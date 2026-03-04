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
    ComponentDef,
    DataConfig,
    aggregate_by_provider_generic,
    PROVIDER_COL,
    PROBLEMS_COL,
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

    with st.expander("🧬 Screening components"):
        st.caption(
            "Define the measures tracked in your QI project. "
            "These drive the Excel template and the processing logic."
        )
        cfg_provider_col = st.text_input(
            "Provider column name",
            value=_secret("PROVIDER_COL", PROVIDER_COL),
            help="Column in the EMR export that identifies the encounter provider.",
        )
        cfg_problems_col = st.text_input(
            "Diagnosis column name",
            value=_secret("PROBLEMS_COL", PROBLEMS_COL),
            help="Column that lists patient diagnoses / problem list.",
        )
        cfg_diagnosis_keywords = st.text_area(
            "Diagnosis keywords (comma-separated)",
            value=_secret(
                "DIAGNOSIS_KEYWORDS",
                "lupus, systemic lupus, SLE, DLE, MCTD, mixed connective tissue, "
                "JIA, JRA, juvenile idiopathic arthritis, juvenile rheumatoid arthritis, "
                "juvenile arthritis, polyarticular, oligoarticular, pauciarticular, "
                "enthesitis",
            ),
            help="Patients are included only if their diagnosis column contains at least one of these keywords (case-insensitive).",
        )

        # Default components — read from secrets [[components]] if present
        try:
            _secret_comps = list(st.secrets.get("components", []) or [])
        except Exception:
            _secret_comps = []

        _default_comps = _secret_comps or [
            {"label": "Lipids",          "has_date": True},
            {"label": "HbA1c",           "has_date": True},
            {"label": "Blood Pressure",  "has_date": False},
            {"label": "BMI",             "has_date": False},
            {"label": "Smoking Status",  "has_date": False},
        ]

        cfg_components_df = st.data_editor(
            pd.DataFrame(_default_comps),
            column_config={
                "label":    st.column_config.TextColumn("Component name", required=True),
                "has_date": st.column_config.CheckboxColumn(
                    "Has date column?",
                    help="If checked, the template includes a '{name} Date' column and checks it's within the lookback window.",
                ),
            },
            num_rows="dynamic",
            use_container_width=True,
            key="cfg_components",
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

def load_providers() -> pd.DataFrame | None:
    """Load provider list from st.secrets, then fall back to local CSV."""
    try:
        secret_providers = st.secrets.get("providers", None)
        if secret_providers:
            return pd.DataFrame(secret_providers)
    except Exception:
        pass

    provider_path = Path(PROVIDER_LIST)
    if not provider_path.exists():
        st.error(
            f"`{PROVIDER_LIST}` not found and no `[[providers]]` entries in Secrets. "
            "Add providers to Streamlit Secrets or upload the CSV."
        )
        return None
    return pd.read_csv(provider_path)


def load_summary_and_providers(summary_source=None):
    """Load and merge summary with provider list. summary_source can be a path or UploadedFile."""
    if summary_source is None:
        summary_path = Path(SUMMARY_FILE)
        if not summary_path.exists():
            return None
        summary = pd.read_csv(summary_path)
    else:
        summary = pd.read_csv(summary_source)

    providers = load_providers()
    if providers is None:
        return None
    return summary.merge(providers, on="provider_id", how="inner")


def _build_data_config() -> DataConfig:
    """Construct a DataConfig from the current sidebar settings."""
    keywords = [k.strip() for k in cfg_diagnosis_keywords.split(",") if k.strip()]
    components = [
        ComponentDef(
            key=f"comp_{i}",
            label=row["label"],
            has_date=bool(row.get("has_date", False)),
        )
        for i, (_, row) in enumerate(cfg_components_df.iterrows())
        if str(row.get("label", "")).strip()
    ]
    return DataConfig(
        provider_col=cfg_provider_col,
        problems_col=cfg_problems_col,
        diagnosis_keywords=keywords,
        components=components,
    )


def _generate_template(config: DataConfig) -> bytes:
    """Return an Excel workbook (bytes) with the expected column headers and one example row."""
    columns = [config.provider_col, config.problems_col]
    for comp in config.components:
        columns.append(comp.label)
        if comp.has_date:
            columns.append(f"{comp.label} Date")

    example = {config.provider_col: "SMITH, JANE A", config.problems_col: "Lupus; SLE"}
    for comp in config.components:
        example[comp.label] = "documented value"
        if comp.has_date:
            example[f"{comp.label} Date"] = "2024-01-15"

    df_tpl = pd.DataFrame([example], columns=columns)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_tpl.to_excel(writer, index=False, sheet_name="Data")
    return buf.getvalue()


def _validate_columns(df: pd.DataFrame, config: DataConfig) -> list[str]:
    """Return a list of column names that are expected but missing from df."""
    missing = []
    for col in [config.provider_col, config.problems_col]:
        if col not in df.columns:
            missing.append(col)
    for comp in config.components:
        if comp.label not in df.columns:
            missing.append(comp.label)
        if comp.has_date and f"{comp.label} Date" not in df.columns:
            missing.append(f"{comp.label} Date")
    return missing


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

    # ── Download template ─────────────────────────────────────────────────────
    with st.expander("📥 Step 0 — Download the input template", expanded=True):
        st.markdown(
            "Download the Excel template, fill it with data from your EMR export, "
            "then upload it below. The template columns are defined by your "
            "**Screening components** settings in the sidebar."
        )
        data_config_for_tpl = _build_data_config()
        st.download_button(
            label="⬇️ Download EMR template (.xlsx)",
            data=_generate_template(data_config_for_tpl),
            file_name="emr_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption(
            f"Required columns: **{data_config_for_tpl.provider_col}**, "
            f"**{data_config_for_tpl.problems_col}**, "
            + ", ".join(
                f"**{c.label}**" + (f", **{c.label} Date**" if c.has_date else "")
                for c in data_config_for_tpl.components
            )
        )

    st.markdown("---")
    st.markdown("Upload the completed (password-encrypted) file. The raw file will be deleted automatically after processing.")

    col1, col2 = st.columns([2, 1])
    with col1:
        uploaded_xlsx = st.file_uploader(
            "Completed EMR template (.xlsx)",
            type=["xlsx", "xls"],
            help="The filled-in template — password-protected is fine",
        )
    with col2:
        excel_password = st.text_input(
            "Excel password (if encrypted)",
            type="password",
            help="Leave blank if the file is not password-protected",
        )

    if st.button("Process Data", type="primary", disabled=not uploaded_xlsx):

        suffix   = Path(uploaded_xlsx.name).suffix
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
                tmp.write(uploaded_xlsx.getbuffer())
                tmp_path = tmp.name

            with st.spinner("Decrypting and loading…"):
                try:
                    df = load_file(tmp_path, excel_password or None)
                except SystemExit:
                    st.error("Could not open the file. Check the password and try again.")
                    st.stop()

        finally:
            if tmp_path and Path(tmp_path).exists():
                Path(tmp_path).unlink()

        st.success(f"Loaded {len(df):,} rows. Raw file deleted from this machine.")

        # ── Validate columns ──────────────────────────────────────────────────
        data_config = _build_data_config()
        missing_cols = _validate_columns(df, data_config)
        if missing_cols:
            st.error(
                "**Missing columns in the uploaded file.** "
                "Make sure you're using the downloaded template and all columns are present.\n\n"
                "Missing: " + ", ".join(f"`{c}`" for c in missing_cols) + "\n\n"
                "Columns found: " + ", ".join(f"`{c}`" for c in df.columns.tolist())
            )
            st.stop()

        # ── Filter by diagnosis ───────────────────────────────────────────────
        keywords = data_config.diagnosis_keywords

        def has_target_dx(problems):
            if pd.isna(problems):
                return False
            text = str(problems).lower()
            return any(kw.lower() in text for kw in keywords)

        df["_provider"] = df[data_config.provider_col].apply(parse_provider)
        before   = len(df)
        df       = df[df[data_config.problems_col].apply(has_target_dx)].copy()
        after_dx = len(df)

        # ── Filter by approved providers (from secrets / provider_list.csv) ───
        providers_df = load_providers()
        approved     = set(providers_df["provider_id"].tolist()) if providers_df is not None else set()

        unassigned = df[~df["_provider"].isin(approved)].copy()
        df         = df[df["_provider"].isin(approved)].copy()

        st.info(f"After diagnosis filter: **{after_dx}** of {before} patients  ·  After provider filter: **{len(df)}** in scope")

        if len(unassigned):
            with st.expander(f"⚠️ {len(unassigned)} patients with target diagnosis but unapproved provider — review only, do not export"):
                id_cols = [c for c in ["DOB", "Age", "Sex", "_provider"] if c in unassigned.columns]
                st.dataframe(
                    unassigned[id_cols].rename(columns={"_provider": "Current Provider"}),
                    use_container_width=True,
                )

        with st.spinner("Calculating screening rates…"):
            summary = aggregate_by_provider_generic(df, data_config)

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
