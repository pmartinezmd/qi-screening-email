"""
preview.py
----------
Renders one or all provider emails to output/preview_<id>.html and opens in the browser.
Useful for checking layout and content before sending.

Usage:
  python preview.py --provider SMITH001 --period "Jan–Feb 2026"
  python preview.py --all --period "Jan–Feb 2026"
"""

import argparse
import subprocess
import sys
from pathlib import Path

import pandas as pd
from jinja2 import Environment, FileSystemLoader, select_autoescape

from send_emails import (
    SUMMARY_FILE, PROVIDER_LIST, TEMPLATE_DIR, TEMPLATE_FILE,
    compute_group_stats, build_context, render_email,
)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--provider", default=None, help="Provider ID to preview")
    parser.add_argument("--all",      action="store_true", help="Render all providers")
    parser.add_argument("--period",   default="Jan–Feb 2026", help="Period label")
    args = parser.parse_args()

    summary   = pd.read_csv(SUMMARY_FILE)
    providers = pd.read_csv(PROVIDER_LIST)
    merged    = summary.merge(providers, on="provider_id", how="inner")
    group_stats = compute_group_stats(merged)

    env = Environment(loader=FileSystemLoader(TEMPLATE_DIR), autoescape=select_autoescape(["html"]))

    if args.all:
        subset = merged
    elif args.provider:
        subset = merged[merged["provider_id"] == args.provider]
        if subset.empty:
            print(f"ERROR: provider '{args.provider}' not found.", file=sys.stderr)
            sys.exit(1)
    else:
        print("Specify --provider <id> or --all", file=sys.stderr)
        sys.exit(1)

    Path("output").mkdir(exist_ok=True)
    paths = []

    max_rate = merged["screening_rate"].max()
    for _, row in subset.iterrows():
        is_top  = (row["screening_rate"] == max_rate and max_rate > 0)
        context = build_context(row, group_stats, args.period, is_top_performer=is_top)
        html    = render_email(context, env)
        out_path = Path("output") / f"preview_{row['provider_id']}.html"
        out_path.write_text(html)
        print(f"  Rendered → {out_path}")
        paths.append(out_path)

    if paths:
        subprocess.run(["open", str(paths[0])])


if __name__ == "__main__":
    main()
