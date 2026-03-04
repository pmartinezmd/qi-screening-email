"""
app.py — Screening QI Email Pipeline
--------------------------------------
Streamlit front end for processing EMR exports and sending provider feedback emails.

Usage:
  streamlit run app.py
"""

import io
import os
import tempfile
from datetime import datetime
from dateutil.relativedelta import relativedelta
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from dotenv import load_dotenv
from jinja2 import Environment, FileSystemLoader

load_dotenv()

def _secret(key: str, default: str = "") -> str:
    """Read from st.secrets first (Streamlit Cloud), then env vars, then default."""
    try:
        return st.secrets.get(key, os.getenv(key, default))
    except Exception:
        return os.getenv(key, default)

from process_data import (
    load_file,
    parse_provider,
    detect_available_components,
    aggregate_by_provider,
    APPROVED_PROVIDERS,
    TARGET_DX_PATTERN,
    PROBLEMS_COL,
    PROVIDER_COL,
)
from send_emails import (
    build_context,
    render_email,
    compute_group_stats,
    send_email,
    load_send_log,
    record_send,
    TEMPLATE_DIR,
    TEMPLATE_FILE,
    PROVIDER_LIST,
    SUMMARY_FILE,
)

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="QI Email Pipeline",
    page_icon="📊",
    layout="wide",
)

# ── Settings sidebar ──────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")
    st.caption("Configure your QI project. Changes apply immediately — no code editing needed.")

    with st.expander("🏷️ Branding", expanded=True):
        cfg_screening_name = st.text_input(
            "QI project name",
            value=_secret("SCREENING_NAME", "Screening QI"),
            help="Appears in the app title, email subject line, and email header.",
        )
        cfg_team_label = st.text_input(
            "Team / institution label",
            value=_secret("TEAM_LABEL", "QI Team · Your Institution"),
            help="Shown in the email header and footer (e.g. 'Rheumatology QI · Children's Hospital').",
        )
        cfg_from_name = st.text_input(
            "Sender display name",
            value=_secret("FROM_NAME", "Screening QI Team"),
            help="The 'From' name recipients see (not the email address).",
        )

    with st.expander("📧 Email content"):
        cfg_target_rate = st.number_input(
            "Screening target (%)",
            min_value=1, max_value=100,
            value=int(_secret("TARGET_RATE", "80")),
            help="The target completion rate shown in the email (green ≥ target, amber ≥ 75% of target, red below).",
        )
        cfg_dashboard_url = st.text_input(
            "Dashboard URL (optional)",
            value=_secret("DASHBOARD_URL", ""),
            help="If provided, a link appears in the email nudge section.",
        )

    with st.expander("🔌 SMTP (advanced)"):
        cfg_smtp_host = st.text_input(
            "SMTP server",
            value=_secret("SMTP_HOST", "smtp.office365.com"),
        )
        cfg_smtp_port = st.number_input(
            "SMTP port",
            min_value=1, max_value=65535,
            value=int(_secret("SMTP_PORT", "587")),
        )

# Aliases used throughout the app
screening_name = cfg_screening_name
team_label     = cfg_team_label

st.title(f"{screening_name} — Email Pipeline")
st.caption(team_label)

tab1, tab2, tab3 = st.tabs(["📊 Process Data", "📧 Preview Email", "🚀 Send Emails"])

# Default reporting period: the two calendar months before today
_today = datetime.today()
_m1 = _today - relativedelta(months=2)
_m2 = _today - relativedelta(months=1)
DEFAULT_PERIOD = (
    f"{_m1.strftime('%B')} – {_m2.strftime('%B')} {_m2.year}"
)


# ── Helpers ───────────────────────────────────────────────────────────────────

def load_summary_and_providers(summary_source=None):
    """Load and merge summary with provider list. summary_source can be a path or UploadedFile."""
    if summary_source is None:
        summary_path = Path(SUMMARY_FILE)
        if not summary_path.exists():
            return None
        summary = pd.read_csv(summary_path)
    else:
        summary = pd.read_csv(summary_source)

    provider_path = Path(PROVIDER_LIST)
    if not provider_path.exists():
        st.error(f"`{PROVIDER_LIST}` not found. Make sure you're running from the project folder.")
        return None

    providers = pd.read_csv(provider_path)
    return summary.merge(providers, on="provider_id", how="inner")


def rate_badge(rate):
    if rate >= 80:
        color = "#3a7d44"
    elif rate >= 60:
        color = "#e8a838"
    else:
        color = "#c1440e"
    return f'<span style="background:{color};color:white;padding:2px 8px;border-radius:4px;font-weight:bold">{rate}%</span>'


# ── Tab 1: Process Data ──────────────────────────────────────────────────────
with tab1:
    st.error("🏥 **Run this step on the institutional workstation only.** Patient data must not leave the institutional machine.")

    st.markdown("Upload the password-encrypted EMR export. The raw file will be deleted automatically after processing.")

    col1, col2 = st.columns([2, 1])
    with col1:
        uploaded_xlsx = st.file_uploader(
            "EMR export (.xlsx)",
            type=["xlsx", "xls"],
            help="The password-encrypted file exported from your EMR system",
        )
    with col2:
        excel_password = st.text_input(
            "Excel password",
            type="password",
            help="The decryption password for the EMR export",
        )

    if st.button("Process Data", type="primary", disabled=not (uploaded_xlsx and excel_password)):

        suffix   = Path(uploaded_xlsx.name).suffix
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
                tmp.write(uploaded_xlsx.getbuffer())
                tmp_path = tmp.name

            with st.spinner("Decrypting and loading…"):
                try:
                    df = load_file(tmp_path, excel_password)
                except SystemExit:
                    st.error("Could not open the file. Check the password and try again.")
                    st.stop()

        finally:
            if tmp_path and Path(tmp_path).exists():
                Path(tmp_path).unlink()

        st.success(f"Loaded {len(df):,} rows. Raw file deleted from this machine.")

        if PROVIDER_COL not in df.columns:
            st.error(f"Column `{PROVIDER_COL}` not found. Available columns: {df.columns.tolist()}")
            st.stop()

        with st.spinner("Filtering patients…"):
            df["_provider"] = df[PROVIDER_COL].apply(parse_provider)

            def has_target_dx(problems):
                if pd.isna(problems):
                    return False
                return bool(TARGET_DX_PATTERN.search(str(problems)))

            before   = len(df)
            df       = df[df[PROBLEMS_COL].apply(has_target_dx)].copy()
            after_dx = len(df)

            unassigned = df[~df["_provider"].isin(APPROVED_PROVIDERS)].copy()
            df         = df[df["_provider"].isin(APPROVED_PROVIDERS)].copy()

        st.info(f"After diagnosis filter: **{after_dx}** of {before} patients  ·  After provider filter: **{len(df)}** in scope")

        if len(unassigned):
            with st.expander(f"⚠️ {len(unassigned)} patients with target diagnosis but unapproved provider — review only, do not export"):
                id_cols = [c for c in ["DOB", "Age", "Sex", "_provider"] if c in unassigned.columns]
                st.dataframe(
                    unassigned[id_cols].rename(columns={"_provider": "Current Provider"}),
                    use_container_width=True,
                )

        with st.spinner("Calculating screening rates…"):
            available = detect_available_components(df)
            summary   = aggregate_by_provider(df, available)

        st.subheader("Screening Rates by Provider")
        display = summary[["provider_id", "eligible_patients", "screened_patients", "screening_rate",
                           "top_missing_1", "top_missing_2"]].copy()
        display = display.sort_values("screening_rate", ascending=False).reset_index(drop=True)
        display.columns = ["Provider", "Eligible", "Screened", "Rate (%)", "Top Gap 1", "Top Gap 2"]

        team_avg = summary["screening_rate"].mean()
        col1, col2, col3 = st.columns(3)
        col1.metric("Providers", len(summary))
        col2.metric("Total Patients", int(summary["eligible_patients"].sum()))
        col3.metric("Team Average", f"{team_avg:.1f}%", delta=f"{team_avg - cfg_target_rate:.1f}% vs {cfg_target_rate}% target")

        st.dataframe(
            display.style.background_gradient(subset=["Rate (%)"], cmap="RdYlGn", vmin=0, vmax=100),
            use_container_width=True,
            hide_index=True,
        )

        csv_bytes = summary.to_csv(index=False).encode()
        st.download_button(
            label="⬇️ Download processed_summary.csv",
            data=csv_bytes,
            file_name="processed_summary.csv",
            mime="text/csv",
            help="Copy this file to your personal laptop — it contains no patient identifiers",
        )

        st.success("Done. Copy `processed_summary.csv` to your personal laptop to send emails.")


# ── Tab 2: Preview Email ─────────────────────────────────────────────────────
with tab2:
    st.markdown("Preview any provider's email before sending. Upload the summary CSV generated in Step 1.")

    col1, col2 = st.columns([2, 1])
    with col1:
        summary_file_preview = st.file_uploader(
            "processed_summary.csv",
            type=["csv"],
            key="preview_summary",
        )
    with col2:
        period_preview = st.text_input(
            "Reporting period",
            value=DEFAULT_PERIOD,
            key="period_preview",
        )

    merged_preview = None
    if summary_file_preview:
        merged_preview = load_summary_and_providers(summary_file_preview)
    elif Path(SUMMARY_FILE).exists():
        merged_preview = load_summary_and_providers()
        st.caption(f"Using existing `{SUMMARY_FILE}`")

    if merged_preview is not None and not merged_preview.empty:
        provider_options  = merged_preview["provider_id"].tolist()
        selected_provider = st.selectbox("Select provider", provider_options)

        if st.button("Preview Email", type="primary"):
            if not period_preview:
                st.warning("Enter a reporting period label first.")
            else:
                env       = Environment(loader=FileSystemLoader(TEMPLATE_DIR))
                row       = merged_preview[merged_preview["provider_id"] == selected_provider].iloc[0]
                group_stats = compute_group_stats(merged_preview)
                max_rate  = merged_preview["screening_rate"].max()
                is_top    = (row["screening_rate"] == max_rate and max_rate > 0)
                context   = build_context(
                    row, group_stats, period_preview,
                    is_top_performer=is_top,
                    screening_name=cfg_screening_name,
                    team_label=cfg_team_label,
                    dashboard_url=cfg_dashboard_url,
                    target_rate=cfg_target_rate,
                )
                html      = render_email(context, env)
                components.html(html, height=900, scrolling=True)
    else:
        st.info("Upload `processed_summary.csv` to preview emails.")


# ── Tab 3: Send Emails ───────────────────────────────────────────────────────
with tab3:
    st.markdown("Send personalized emails to all providers. Run this on your personal laptop.")

    col1, col2 = st.columns([2, 1])
    with col1:
        summary_file_send = st.file_uploader(
            "processed_summary.csv",
            type=["csv"],
            key="send_summary",
        )
    with col2:
        period_send = st.text_input(
            "Reporting period",
            value=DEFAULT_PERIOD,
            key="period_send",
        )

    st.divider()

    col1, col2 = st.columns(2)
    with col1:
        default_user = os.getenv("SMTP_USER", "")
        from_address = st.text_input("Send from", value=default_user)
    with col2:
        smtp_password = st.text_input("Email password", type="password")

    st.divider()

    merged_send = None
    if summary_file_send:
        merged_send = load_summary_and_providers(summary_file_send)
    elif Path(SUMMARY_FILE).exists():
        merged_send = load_summary_and_providers()
        st.caption(f"Using existing `{SUMMARY_FILE}`")

    if merged_send is not None and not merged_send.empty and period_send:
        send_log = load_send_log()

        st.subheader("Recipients")
        rows = []
        for _, row in merged_send.iterrows():
            pid     = row["provider_id"]
            already = (pid, period_send) in send_log
            rows.append({
                "_pid":   pid,
                "_row":   row,
                "Send":   not already,
                "Provider": row["display_name"],
                "Email":  row["email"],
                "Rate":   f"{row['screening_rate']}%",
                "Status": "Already sent" if already else "Pending",
            })

        selected_pids = []
        for r in rows:
            disabled = r["Status"] == "Already sent"
            checked  = st.checkbox(
                f"{r['Provider']}  —  {r['Email']}  —  {r['Rate']}  {'✓ already sent' if disabled else ''}",
                value=r["Send"],
                disabled=disabled,
                key=f"chk_{r['_pid']}",
            )
            if checked and not disabled:
                selected_pids.append(r["_pid"])

        st.divider()

        can_send = bool(selected_pids and from_address and smtp_password and period_send)
        if st.button(f"Send to {len(selected_pids)} provider(s)", type="primary", disabled=not can_send):
            os.environ["SMTP_USER"] = from_address
            subject     = f"{screening_name} Update · {period_send}"
            env         = Environment(loader=FileSystemLoader(TEMPLATE_DIR))
            group_stats = compute_group_stats(merged_send)
            max_rate    = merged_send["screening_rate"].max()

            results      = []
            progress     = st.progress(0)
            status_area  = st.empty()

            to_send = [r for r in rows if r["_pid"] in selected_pids]

            for i, r in enumerate(to_send):
                row = r["_row"]
                pid = r["_pid"]
                status_area.info(f"Sending to {r['Provider']}…")

                is_top  = (row["screening_rate"] == max_rate and max_rate > 0)
                context = build_context(
                    row, group_stats, period_send,
                    is_top_performer=is_top,
                    screening_name=cfg_screening_name,
                    team_label=cfg_team_label,
                    dashboard_url=cfg_dashboard_url,
                    target_rate=cfg_target_rate,
                )
                html    = render_email(context, env)

                try:
                    send_email(
                        row["email"], subject, html, smtp_password,
                        from_name=cfg_from_name,
                        smtp_host=cfg_smtp_host,
                        smtp_port=cfg_smtp_port,
                    )
                    record_send(pid, period_send)
                    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
                    results.append({"Provider": r["Provider"], "Email": r["Email"], "Status": "✓ Sent", "Time": ts})
                except Exception as e:
                    results.append({"Provider": r["Provider"], "Email": r["Email"], "Status": f"✗ Failed: {e}", "Time": ""})

                progress.progress((i + 1) / len(to_send))

            status_area.empty()
            progress.empty()

            sent   = sum(1 for r in results if r["Status"].startswith("✓"))
            failed = len(results) - sent

            if failed == 0:
                st.success(f"All {sent} emails sent successfully.")
            else:
                st.warning(f"Sent: {sent}  ·  Failed: {failed}")

            st.dataframe(pd.DataFrame(results), use_container_width=True, hide_index=True)

    elif not period_send:
        st.info("Enter a reporting period label to load the recipient list.")
    else:
        st.info("Upload `processed_summary.csv` to load recipients.")
