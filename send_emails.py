"""
send_emails.py
--------------
Renders the email template for each provider and sends via SMTP.

Usage:
  python send_emails.py --period "Jan–Feb 2026"             # send to all
  python send_emails.py --period "Jan–Feb 2026" --dry-run   # render only, no sending
  python send_emails.py --period "Jan–Feb 2026" --provider SMITH001  # one provider only
"""

import argparse
import csv
import getpass
import os
import smtplib
import sys
from datetime import datetime, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from jinja2 import Environment, FileSystemLoader

load_dotenv()

SUMMARY_FILE  = "data/processed_summary.csv"
PROVIDER_LIST = "data/provider_list.csv"
SEND_LOG      = "data/send_log.csv"
TEMPLATE_DIR  = "templates"
TEMPLATE_FILE = "email_template.html"
TARGET_RATE   = int(os.getenv("TARGET_RATE", 80))


def load_send_log() -> set[tuple[str, str]]:
    """Return a set of (provider_id, period) pairs already sent."""
    path = Path(SEND_LOG)
    if not path.exists():
        return set()
    with path.open(newline="") as f:
        return {(row["provider_id"], row["period"]) for row in csv.DictReader(f)}


def record_send(provider_id: str, period: str) -> None:
    """Append a successful send to the log."""
    path = Path(SEND_LOG)
    write_header = not path.exists()
    with path.open("a", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["provider_id", "period", "sent_at"])
        if write_header:
            writer.writeheader()
        writer.writerow({
            "provider_id": provider_id,
            "period":      period,
            "sent_at":     datetime.now(timezone.utc).isoformat(),
        })


def load_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    summary   = pd.read_csv(SUMMARY_FILE)
    providers = pd.read_csv(PROVIDER_LIST)
    merged    = summary.merge(providers, on="provider_id", how="inner")
    return merged, providers


def compute_group_stats(merged: pd.DataFrame) -> dict:
    """Compute screening rate averages by provider type (attending / fellow)."""
    stats = {}
    for ptype in ("attending", "fellow"):
        group = merged[merged["provider_type"] == ptype]
        stats[ptype] = {
            "avg": round(group["screening_rate"].mean(), 1) if len(group) > 0 else 0.0,
            "n":   len(group),
        }
    return stats


def rate_color(rate: float, target_rate: int = TARGET_RATE) -> str:
    if rate >= target_rate:
        return "#3a7d44"            # green — at or above target
    elif rate >= target_rate * 0.75:
        return "#e8a838"            # amber — within 25% of target
    else:
        return "#c1440e"            # red — below 75% of target


def missing_count_for(row: pd.Series, component_key: str) -> int:
    col = f"missing_{component_key}"
    return int(row[col]) if col in row and not pd.isna(row[col]) else 0


COMPONENT_KEYS = ["lipids", "a1c", "bp", "bmi", "smoking"]
COMPONENT_LABELS = {
    "lipids":  "Lipids (HDL + Total Chol)",
    "a1c":     "HbA1c",
    "bp":      "Blood Pressure",
    "bmi":     "BMI",
    "smoking": "Smoking Status",
}


def build_context(row: pd.Series, group_stats: dict, period_label: str,
                  is_top_performer: bool = False,
                  screening_name: str | None = None,
                  team_label: str | None = None,
                  dashboard_url: str | None = None,
                  target_rate: int | None = None) -> dict:
    ptype      = row["provider_type"]
    other_type = "fellow" if ptype == "attending" else "attending"

    _target_rate = target_rate if target_rate is not None else TARGET_RATE

    return {
        "display_name":        row["display_name"],
        "period_label":        period_label,
        "screening_rate":      row["screening_rate"],
        "eligible_patients":   int(row["eligible_patients"]),
        "screened_patients":   int(row["screened_patients"]),
        "rate_color":          rate_color(row["screening_rate"], _target_rate),
        "provider_type_label": ptype.capitalize() + "s",
        "group_avg":           group_stats[ptype]["avg"],
        "group_n":             group_stats[ptype]["n"],
        "other_type_label":    other_type.capitalize() + "s",
        "other_avg":           group_stats[other_type]["avg"],
        "other_n":             group_stats[other_type]["n"],
        "top_missing_1":       row.get("top_missing_1"),
        "top_missing_2":       row.get("top_missing_2"),
        "missing_count_1":     int(row.get("missing_count_1") or 0),
        "missing_count_2":     int(row.get("missing_count_2") or 0),
        "target_rate":         _target_rate,
        "dashboard_url":       dashboard_url if dashboard_url is not None else os.getenv("DASHBOARD_URL", ""),
        "team_label":          team_label if team_label is not None else os.getenv("TEAM_LABEL", "QI Team · Your Institution"),
        "screening_name":      screening_name if screening_name is not None else os.getenv("SCREENING_NAME", "Screening QI"),
        "is_top_performer":    is_top_performer,
        "patients_to_screen":  str(row.get("patients_to_screen") or ""),
    }


def render_email(context: dict, env: Environment) -> str:
    template = env.get_template(TEMPLATE_FILE)
    return template.render(**context)


def send_email(to_email: str, subject: str, html_body: str, smtp_password: str,
               from_name: str | None = None,
               smtp_host: str | None = None,
               smtp_port: int | None = None):
    smtp_user = os.getenv("SMTP_USER")
    _from_name = from_name or os.getenv("FROM_NAME", "Screening QI Team")
    _smtp_host = smtp_host or os.getenv("SMTP_HOST", "smtp.office365.com")
    _smtp_port = smtp_port or int(os.getenv("SMTP_PORT", 587))

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = f"{_from_name} <{smtp_user}>"
    msg["To"]      = to_email
    msg.attach(MIMEText(html_body, "html"))

    with smtplib.SMTP(_smtp_host, _smtp_port) as server:
        server.ehlo()
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_user, to_email, msg.as_string())


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--period",   required=True, help='Period label, e.g. "Jan–Feb 2026"')
    parser.add_argument("--dry-run",  action="store_true", help="Render emails to output/ without sending")
    parser.add_argument("--provider", default=None, help="Send to a single provider_id only")
    args = parser.parse_args()

    merged, _ = load_data()
    group_stats = compute_group_stats(merged)
    env = Environment(loader=FileSystemLoader(TEMPLATE_DIR))

    if args.provider:
        subset = merged[merged["provider_id"] == args.provider]
        if subset.empty:
            print(f"ERROR: provider '{args.provider}' not found.", file=sys.stderr)
            sys.exit(1)
    else:
        subset = merged

    screening_name = os.getenv("SCREENING_NAME", "Screening QI")
    subject        = f"{screening_name} Update · {args.period}"
    sent, failed, skipped, already_sent = 0, 0, 0, 0

    max_rate = merged["screening_rate"].max()
    send_log = load_send_log()

    to_send = [
        row for _, row in subset.iterrows()
        if args.dry_run or (row["provider_id"], args.period) not in send_log
    ]
    already_sent = len(subset) - len(to_send)

    if not args.dry_run:
        default_user = os.getenv("SMTP_USER", "")
        user_input   = input(f"\nSend from [{default_user}]: ").strip()
        smtp_user    = user_input if user_input else default_user
        os.environ["SMTP_USER"] = smtp_user

        smtp_password = getpass.getpass(f"Password for {smtp_user}: ")

        print(f"\nWill send to {len(to_send)} provider(s):")
        for row in to_send:
            print(f"  {row['display_name']:<20} <{row['email']}>  —  {row['screening_rate']}%")
        if already_sent:
            print(f"  ({already_sent} already sent for {args.period} — will be skipped)")

        confirm = input("\nSend? [Y/n]: ").strip().lower()
        if confirm not in ("", "y", "yes"):
            print("Aborted — no emails sent.")
            sys.exit(0)
        print()
    else:
        smtp_password = None

    for row in to_send:
        pid = row["provider_id"]

        if not args.dry_run and (pid, args.period) in send_log:
            already_sent += 1
            continue

        is_top  = (row["screening_rate"] == max_rate and max_rate > 0)
        context = build_context(row, group_stats, args.period, is_top_performer=is_top)
        html    = render_email(context, env)

        if args.dry_run:
            out_path = Path("output") / f"email_{pid}.html"
            out_path.parent.mkdir(exist_ok=True)
            out_path.write_text(html)
            print(f"  [dry-run] Rendered → {out_path}")
            skipped += 1
        else:
            try:
                send_email(row["email"], subject, html, smtp_password)
                record_send(pid, args.period)
                ts = datetime.now().strftime("%Y-%m-%d %H:%M")
                print(f"  ✓ Sent  {row['display_name']:<20} <{row['email']}>  [{ts}]")
                sent += 1
            except Exception as e:
                print(f"  ✗ Failed for {pid}: {e}", file=sys.stderr)
                failed += 1

    summary_parts = [f"sent: {sent}", f"dry-run: {skipped}", f"failed: {failed}"]
    if already_sent:
        summary_parts.append(f"skipped (already sent): {already_sent}")
    print(f"\nDone — {',  '.join(summary_parts)}")


if __name__ == "__main__":
    main()
