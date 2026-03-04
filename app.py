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
    load_excel_sheets,
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


def _generate_summary_template(target_rate: int, provider_col: str) -> bytes:
    """Return a 2-sheet Excel workbook for provider-level summary data input.

    Sheet 1 'Summary'  — one row per provider with aggregate stats.
    Sheet 2 'Patient List' — optional; one row per unscreened patient.
    """
    eligible_ex   = 20
    screened_pct_ex = 65.0
    target_no_ex  = round(eligible_ex * target_rate / 100)

    summary_df = pd.DataFrame([{
        provider_col:        "SMITH, JANE A",
        "Eligible Patients": eligible_ex,
        "Screened %":        screened_pct_ex,
        "Target No.":        target_no_ex,
        "Target %":          target_rate,
    }])

    patient_df = pd.DataFrame([
        {provider_col: "SMITH, JANE A", "Patient Name": "JONES, ROBERT"},
        {provider_col: "SMITH, JANE A", "Patient Name": "WILLIAMS, MARY"},
    ])

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
        patient_df.to_excel(writer, index=False, sheet_name="Patient List")
    return buf.getvalue()


def _generate_patient_template(config: DataConfig) -> bytes:
    """Return a single-sheet Excel workbook with patient-level columns and one example row."""
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

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame([example], columns=columns).to_excel(writer, index=False, sheet_name="Data")
    return buf.getvalue()


def _validate_summary_columns(df: pd.DataFrame, provider_col: str) -> list[str]:
    """Return column names expected in the Summary sheet but missing from df."""
    required = [provider_col, "Eligible Patients", "Screened %", "Target No.", "Target %"]
    return [c for c in required if c not in df.columns]


def _validate_patient_columns(df: pd.DataFrame, config: DataConfig) -> list[str]:
    """Return column names expected in the patient-level sheet but missing from df."""
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


def _process_summary_format(
    df_summary: pd.DataFrame,
    df_patients: pd.DataFrame | None,
    providers_df: pd.DataFrame | None,
    provider_col: str,
) -> pd.DataFrame:
    """Process provider-level summary sheet into the standard processed_summary schema."""
    rows = []
    for _, row in df_summary.iterrows():
        raw         = str(row[provider_col]).strip()
        provider_id = parse_provider(raw)
        eligible    = int(row["Eligible Patients"])
        screened_pct = float(row["Screened %"])
        target_no   = int(row.get("Target No.", 0) or 0)
        screened    = round(screened_pct / 100 * eligible)
        gap         = max(0, target_no - screened)

        patient_names: list[str] = []
        if df_patients is not None and not df_patients.empty and provider_col in df_patients.columns:
            mask          = df_patients[provider_col].astype(str).str.strip() == raw
            patient_names = df_patients[mask]["Patient Name"].dropna().astype(str).tolist()

        patient_count = len(patient_names) if patient_names else gap

        rows.append({
            "provider_id":        provider_id,
            "eligible_patients":  eligible,
            "screened_patients":  screened,
            "screening_rate":     round(screened_pct, 1),
            "top_missing_1":      None,
            "top_missing_2":      None,
            "missing_count_1":    patient_count,
            "missing_count_2":    0,
            "patients_to_screen": ", ".join(patient_names),
        })

    summary = pd.DataFrame(rows)
    if providers_df is not None:
        summary = summary.merge(providers_df, on="provider_id", how="inner")
    return summary


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
        st.markdown("Choose the template that matches your EMR report format.")
        tpl_col1, tpl_col2 = st.columns(2)

        with tpl_col1:
            st.markdown("**Provider summary** *(pre-aggregated)*")
            st.caption(
                "Use this when your EMR report already shows totals per provider: "
                "number of eligible patients, % screened, and target. "
                "Optionally list specific unscreened patients in the second sheet."
            )
            st.download_button(
                label="⬇️ Summary template (.xlsx)",
                data=_generate_summary_template(cfg_target_rate, cfg_provider_col),
                file_name="screening_summary_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            st.caption(
                f"Sheet 1 — `{cfg_provider_col}`, `Eligible Patients`, `Screened %`, `Target No.`, `Target %`  \n"
                f"Sheet 2 *(optional)* — `{cfg_provider_col}`, `Patient Name`"
            )

        with tpl_col2:
            st.markdown("**Patient-level detail** *(raw EMR export)*")
            st.caption(
                "Use this when your EMR report has one row per patient. "
                "The app filters by diagnosis and calculates rates automatically. "
                "Screening components are defined in the sidebar."
            )
            _data_config_tpl = _build_data_config()
            st.download_button(
                label="⬇️ Patient-level template (.xlsx)",
                data=_generate_patient_template(_data_config_tpl),
                file_name="emr_patient_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            st.caption(
                f"`{_data_config_tpl.provider_col}`, `{_data_config_tpl.problems_col}`, "
                + ", ".join(
                    f"`{c.label}`" + (f", `{c.label} Date`" if c.has_date else "")
                    for c in _data_config_tpl.components
                )
            )

    st.markdown("---")
    st.markdown("Upload the completed (optionally password-encrypted) file. The raw file is deleted automatically after processing.")

    col1, col2 = st.columns([2, 1])
    with col1:
        uploaded_xlsx = st.file_uploader(
            "Completed template (.xlsx) — summary or patient-level",
            type=["xlsx", "xls"],
            help="Either template format is accepted — the app detects which one you uploaded",
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

            with st.spinner("Loading file…"):
                try:
                    sheets = load_excel_sheets(tmp_path, excel_password or None)
                except SystemExit:
                    st.error("Could not open the file. Check the password and try again.")
                    st.stop()

        finally:
            if tmp_path and Path(tmp_path).exists():
                Path(tmp_path).unlink()

        # ── Detect format ─────────────────────────────────────────────────────
        is_summary_format = "Summary" in sheets

        providers_df = load_providers()

        if is_summary_format:
            # ── Provider summary format ───────────────────────────────────────
            df_summary  = sheets["Summary"].dropna(how="all")
            df_patients = sheets.get("Patient List")
            if df_patients is not None:
                df_patients = df_patients.dropna(how="all")
                if df_patients.empty:
                    df_patients = None

            st.success(f"Detected **summary format** — {len(df_summary)} provider rows loaded. Raw file deleted.")

            missing_cols = _validate_summary_columns(df_summary, cfg_provider_col)
            if missing_cols:
                st.error(
                    "**Missing columns in the Summary sheet.**\n\n"
                    "Missing: " + ", ".join(f"`{c}`" for c in missing_cols) + "\n\n"
                    "Columns found: " + ", ".join(f"`{c}`" for c in df_summary.columns.tolist())
                )
                st.stop()

            with st.spinner("Calculating screening rates…"):
                summary = _process_summary_format(df_summary, df_patients, providers_df, cfg_provider_col)

        else:
            # ── Patient-level format ──────────────────────────────────────────
            first_sheet = next(iter(sheets))
            df = sheets[first_sheet].dropna(how="all")
            st.success(f"Detected **patient-level format** — {len(df):,} rows loaded. Raw file deleted.")

            data_config  = _build_data_config()
            missing_cols = _validate_patient_columns(df, data_config)
            if missing_cols:
                st.error(
                    "**Missing columns in the uploaded file.**\n\n"
                    "Missing: " + ", ".join(f"`{c}`" for c in missing_cols) + "\n\n"
                    "Columns found: " + ", ".join(f"`{c}`" for c in df.columns.tolist())
                )
                st.stop()

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

            approved   = set(providers_df["provider_id"].tolist()) if providers_df is not None else set()
            unassigned = df[~df["_provider"].isin(approved)].copy()
            df         = df[df["_provider"].isin(approved)].copy()

            st.info(f"After diagnosis filter: **{after_dx}** of {before} patients  ·  After provider filter: **{len(df)}** in scope")

            if len(unassigned):
                with st.expander(f"⚠️ {len(unassigned)} patients with target diagnosis but unapproved provider"):
                    id_cols = [c for c in ["DOB", "Age", "Sex", "_provider"] if c in unassigned.columns]
                    st.dataframe(unassigned[id_cols].rename(columns={"_provider": "Current Provider"}), use_container_width=True)

            with st.spinner("Calculating screening rates…"):
                summary = aggregate_by_provider_generic(df, data_config)
                summary["patients_to_screen"] = ""

        # ── Results (shared by both paths) ────────────────────────────────────
        if summary.empty:
            st.warning(
                "No providers matched. Check that provider names match your configured providers exactly."
            )
            st.stop()

        unmatched = (len(df_summary) if is_summary_format else 0) - len(summary) if is_summary_format else 0
        if unmatched > 0:
            st.info(f"Note: {unmatched} provider row(s) did not match any configured provider and were excluded.")

        st.subheader("Screening Rates by Provider")

        gap_col    = "missing_count_1" if "missing_count_1" in summary.columns else None
        disp_cols  = ["provider_id", "eligible_patients", "screened_patients", "screening_rate"]
        disp_names = ["Provider", "Eligible", "Screened", "Rate (%)"]
        if gap_col:
            disp_cols.append(gap_col)
            disp_names.append("Patients to Screen")
        if "top_missing_1" in summary.columns:
            disp_cols += ["top_missing_1", "top_missing_2"]
            disp_names += ["Top Gap 1", "Top Gap 2"]

        display = summary[disp_cols].copy().sort_values("screening_rate", ascending=False).reset_index(drop=True)
        display.columns = disp_names

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

        # ── Patient list (shown here only — not exported) ─────────────────────
        if "patients_to_screen" in summary.columns:
            has_patient_names = summary["patients_to_screen"].fillna("").str.strip().ne("").any()
            if has_patient_names:
                with st.expander("👥 Patients to screen by provider"):
                    for _, row in summary.iterrows():
                        pts = str(row.get("patients_to_screen", "")).strip()
                        if pts:
                            st.markdown(f"**{row['provider_id']}** ({int(row['missing_count_1'])}): {pts}")

        export_summary = summary.drop(columns=["patients_to_screen"], errors="ignore")
        csv_bytes = export_summary.to_csv(index=False).encode()
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
