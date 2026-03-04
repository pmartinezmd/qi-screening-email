"""
process_data.py
---------------
Reads the raw EMR export (password-protected .xlsx) and produces
processed_summary.csv with per-provider screening completion rates
and missing-component counts.

Customize the sections marked ── CONFIGURE ── for your QI project:
  - TARGET_DX_PATTERN: diagnosis filter (regex matching your patient population)
  - APPROVED_PROVIDERS: set of provider names as they appear in the EMR export
  - COMPONENT_LABELS / column constants: screening components you want to track
  - PROVIDER_COL / PROBLEMS_COL: column names in your EMR export

Usage:
  python process_data.py
  python process_data.py --input data/my_export.xlsx --output data/processed_summary.csv
"""

import argparse
import getpass
import io
import re
import sys
import os
import pandas as pd
from dataclasses import dataclass, field
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# ── File paths ─────────────────────────────────────────────────────────────────
INPUT_FILE  = "data/emr_export.xlsx"   # update each cycle or pass via --input
OUTPUT_FILE = "data/processed_summary.csv"

# ── CONFIGURE: EMR column names ────────────────────────────────────────────────
# Update these to match the column headers in your EMR export.
PROVIDER_COL  = "Encounter Provider"   # column identifying the encounter provider
PROBLEMS_COL  = "Problems"             # column containing patient diagnoses

# Screening component column names — adjust to match your export format
HDL_COL       = "Last HDL"
HDL_DATE_COL  = "HDL Date"
CHOL_DATE_COL = "CHOL Date"
BP_COL        = "BP"
SMOKING_COL   = "Smoking Status"
WEIGHT_COL    = "Weight"
HEIGHT_COL    = "Height"
BMI_COL       = "Last BMI"
A1C_COL       = "Last HGBA1C Value "   # trailing space may be in the actual column name
A1C_DATE_COL  = "HBA1C Date"

# ── CONFIGURE: lookback window ─────────────────────────────────────────────────
LOOKBACK_DAYS = 365
REPORT_DATE   = pd.Timestamp.today().normalize()

# ── CONFIGURE: diagnosis filter ────────────────────────────────────────────────
# Patients are included only if their PROBLEMS_COL field matches this pattern.
# The example below matches lupus and JIA subtypes — replace with your diagnoses.
TARGET_DX_PATTERN = re.compile(
    r'\blupus\b'
    r'|systemic lupus'
    r'|\bSLE\b'
    r'|\bDLE\b'
    r'|\bMCTD\b'
    r'|mixed connective tissue'
    r'|\bJIA\b'
    r'|\bJRA\b'
    r'|juvenile idiopathic arthritis'
    r'|juvenile rheumatoid arthritis'
    r'|juvenile arthritis'
    r'|juvenile ankylosing spondylitis'
    r'|juvenile polyarthritis'
    r'|juvenile psoriatic arthritis'
    r'|polyarticular'
    r'|oligoarticular'
    r'|pauciarticular'
    r'|enthesitis.related arthritis',
    re.IGNORECASE
)

# ── CONFIGURE: approved providers ─────────────────────────────────────────────
# Add the names of providers to include, exactly as they appear in PROVIDER_COL
# (first line only, if the field contains multiple lines).
# Example: "SMITH, JANE A"
APPROVED_PROVIDERS = {
    # "LAST, FIRST M",
}

# ── CONFIGURE: screening components ───────────────────────────────────────────
# Labels used in the output CSV and email. Keys must match check_* functions below.
COMPONENT_LABELS = {
    "lipids":  "Lipids (last panel)",
    "bp":      "Blood Pressure",
    "smoking": "Smoking Status",
    "bmi":     "BMI",
    "a1c":     "HbA1c",
}


# ── Generic processing (used by the Streamlit app) ────────────────────────────

@dataclass
class ComponentDef:
    """One screening component as configured in the app sidebar / secrets."""
    key: str       # slug used in CSV column names, e.g. "lipids"
    label: str     # human label = also the Excel template column header
    has_date: bool = False  # if True, expects a "{label} Date" column

@dataclass
class DataConfig:
    """Full configuration for one QI project's processing run."""
    provider_col: str
    problems_col: str
    diagnosis_keywords: list[str]        # case-insensitive substring match
    components: list[ComponentDef]
    lookback_days: int = 365


def _has_value(val) -> bool:
    """True if the cell contains a meaningful documented value."""
    if pd.isna(val):
        return False
    return str(val).strip().lower() not in ("", "n/a", "null", "unknown",
                                            "never assessed", "not documented")


def _within_n_days(date_val, lookback_days: int) -> bool:
    if pd.isna(date_val):
        return False
    try:
        d = pd.Timestamp(date_val)
        report_date = pd.Timestamp.today().normalize()
        return 0 <= (report_date - d).days <= lookback_days
    except Exception:
        return False


def assess_row_generic(row: pd.Series, config: DataConfig) -> dict:
    results = {}
    for comp in config.components:
        val = row.get(comp.label)
        if not _has_value(val):
            results[comp.key] = False
        elif comp.has_date:
            results[comp.key] = _within_n_days(
                row.get(f"{comp.label} Date"), config.lookback_days
            )
        else:
            results[comp.key] = True
    results["complete"] = all(results[c.key] for c in config.components)
    return results


def aggregate_by_provider_generic(df: pd.DataFrame, config: DataConfig) -> pd.DataFrame:
    df = df.copy()
    df["_provider"] = df[config.provider_col].apply(parse_provider)

    records = []
    for provider, group in df.groupby("_provider"):
        assessments = group.apply(
            lambda row: assess_row_generic(row, config), axis=1
        ).tolist()

        n_total    = len(assessments)
        n_complete = sum(a["complete"] for a in assessments)
        rate       = round(n_complete / n_total * 100, 1) if n_total > 0 else 0.0

        missing_counts = {
            c.key: sum(1 for a in assessments if not a[c.key])
            for c in config.components
        }
        assessed_missing = sorted(
            [(c, missing_counts[c.key]) for c in config.components
             if missing_counts[c.key] > 0],
            key=lambda x: -x[1],
        )
        top1 = assessed_missing[0][0] if len(assessed_missing) > 0 else None
        top2 = assessed_missing[1][0] if len(assessed_missing) > 1 else None

        record = {
            "provider_id":         provider,
            "eligible_patients":   n_total,
            "screened_patients":   n_complete,
            "screening_rate":      rate,
            "top_missing_1":       top1.label if top1 else None,
            "top_missing_2":       top2.label if top2 else None,
            "missing_count_1":     missing_counts[top1.key] if top1 else 0,
            "missing_count_2":     missing_counts[top2.key] if top2 else 0,
            "components_assessed": ", ".join(c.label for c in config.components),
        }
        for c in config.components:
            record[f"missing_{c.key}"] = missing_counts[c.key]

        records.append(record)

    return pd.DataFrame(records)


# ── File loading ───────────────────────────────────────────────────────────────

def load_file(path: str, password: str | None = None) -> pd.DataFrame:
    p = Path(path)
    if not p.exists():
        print(f"ERROR: file not found: {path}", file=sys.stderr)
        sys.exit(1)

    if p.suffix.lower() == ".csv":
        return pd.read_csv(p)

    # Try reading directly first (unencrypted)
    try:
        return pd.read_excel(p, engine="openpyxl")
    except Exception:
        pass

    # Encrypted — decrypt with msoffcrypto
    if not password:
        print("ERROR: file appears to be encrypted but no password was provided.", file=sys.stderr)
        sys.exit(1)

    try:
        import msoffcrypto
    except ImportError:
        print("ERROR: install msoffcrypto-tool:  pip install msoffcrypto-tool", file=sys.stderr)
        sys.exit(1)

    with open(p, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password=password)
        decrypted = io.BytesIO()
        office_file.decrypt(decrypted)

    decrypted.seek(0)
    return pd.read_excel(decrypted, engine="openpyxl")


def load_excel_sheets(path: str, password: str | None = None) -> dict[str, pd.DataFrame]:
    """Return all sheets from an Excel workbook as {sheet_name: DataFrame}."""
    p = Path(path)
    if not p.exists():
        print(f"ERROR: file not found: {path}", file=sys.stderr)
        sys.exit(1)

    try:
        return pd.read_excel(p, sheet_name=None, engine="openpyxl")
    except Exception:
        pass

    if not password:
        print("ERROR: file appears to be encrypted but no password was provided.", file=sys.stderr)
        sys.exit(1)

    try:
        import msoffcrypto
    except ImportError:
        print("ERROR: install msoffcrypto-tool:  pip install msoffcrypto-tool", file=sys.stderr)
        sys.exit(1)

    with open(p, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password=password)
        decrypted = io.BytesIO()
        office_file.decrypt(decrypted)

    decrypted.seek(0)
    return pd.read_excel(decrypted, sheet_name=None, engine="openpyxl")


# ── Parsing helpers ────────────────────────────────────────────────────────────

def parse_provider(val) -> str:
    """Take the first line of a multi-line provider field and strip whitespace."""
    return str(val).split("\n")[0].strip()


def within_12_months(date_val) -> bool:
    """True if date_val falls within LOOKBACK_DAYS before REPORT_DATE."""
    if pd.isna(date_val):
        return False
    try:
        d = pd.Timestamp(date_val)
        return 0 <= (REPORT_DATE - d).days <= LOOKBACK_DAYS
    except Exception:
        return False


def parse_weight_kg(val) -> float | None:
    """'44.5 kg' → 44.5"""
    if pd.isna(val):
        return None
    m = re.search(r"([\d.]+)\s*kg", str(val), re.IGNORECASE)
    return float(m.group(1)) if m else None


def parse_height_cm(val) -> float | None:
    """'143.6 cm (4\' 8.54")' → 143.6"""
    if pd.isna(val):
        return None
    m = re.search(r"([\d.]+)\s*cm", str(val), re.IGNORECASE)
    return float(m.group(1)) if m else None


# ── CONFIGURE: screening component checks ─────────────────────────────────────
# Each function returns True if that component is considered complete for a patient.
# Adapt these rules to match your screening measure definition.

def check_lipids(hdl_val, hdl_date, chol_date) -> bool:
    """True if HDL value is present AND a lipid panel date is within 12 months."""
    if pd.isna(hdl_val):
        return False
    return within_12_months(hdl_date) or within_12_months(chol_date)


def check_bp(val) -> bool:
    """True if BP is present (strips EMR alert flags like '(!) ')."""
    if pd.isna(val):
        return False
    cleaned = re.sub(r"^\(!\)\s*", "", str(val)).strip()
    return bool(cleaned)


def check_smoking(val) -> bool:
    """True if Smoking Status has been formally assessed."""
    if pd.isna(val):
        return False
    return str(val).strip().lower() not in ("", "n/a", "null", "unknown", "never assessed")


def check_bmi(last_bmi_val, weight_val, height_val) -> bool:
    """True if BMI is available — from Last BMI field or calculated from Weight + Height."""
    if pd.notna(last_bmi_val) and re.search(r"[\d.]+", str(last_bmi_val)):
        return True
    return parse_weight_kg(weight_val) is not None and parse_height_cm(height_val) is not None


def check_a1c(value, date_val) -> bool:
    """True if A1C has a numeric value AND the date is within 12 months."""
    if pd.isna(value):
        return False
    return within_12_months(date_val)


# ── Detect available components ────────────────────────────────────────────────

def detect_available_components(df: pd.DataFrame) -> list[str]:
    """Report data coverage for each component. All 5 are always assessed."""
    available = ["lipids", "bp", "smoking", "bmi", "a1c"]
    n = len(df)

    n_lipids  = df[HDL_DATE_COL].apply(within_12_months).sum() if HDL_DATE_COL in df.columns else 0
    n_bp      = df[BP_COL].apply(check_bp).sum() if BP_COL in df.columns else 0
    n_smoking = df[SMOKING_COL].apply(check_smoking).sum() if SMOKING_COL in df.columns else 0
    n_bmi     = sum(
        check_bmi(row.get(BMI_COL), row.get(WEIGHT_COL), row.get(HEIGHT_COL))
        for _, row in df.iterrows()
    )
    if A1C_COL in df.columns and A1C_DATE_COL in df.columns:
        n_a1c = sum(
            check_a1c(row.get(A1C_COL), row.get(A1C_DATE_COL))
            for _, row in df.iterrows()
        )
    else:
        n_a1c = 0

    for key, count in [("lipids", n_lipids), ("bp", n_bp), ("smoking", n_smoking),
                       ("bmi", n_bmi), ("a1c", n_a1c)]:
        pct = count / n * 100 if n > 0 else 0
        symbol = "✓" if pct >= 20 else "⚠"
        print(f"  {symbol} {COMPONENT_LABELS[key]}: {count}/{n} ({pct:.0f}%)")

    return available


# ── Per-row assessment ─────────────────────────────────────────────────────────

def assess_row(row, available_components: list[str]) -> dict:
    results = {
        "lipids":  check_lipids(row.get(HDL_COL), row.get(HDL_DATE_COL), row.get(CHOL_DATE_COL)),
        "bp":      check_bp(row.get(BP_COL)),
        "smoking": check_smoking(row.get(SMOKING_COL)),
        "bmi":     check_bmi(row.get(BMI_COL), row.get(WEIGHT_COL), row.get(HEIGHT_COL)),
        "a1c":     check_a1c(row.get(A1C_COL), row.get(A1C_DATE_COL)),
    }
    results["complete"] = all(results[c] for c in available_components)
    return results


# ── Aggregation ────────────────────────────────────────────────────────────────

def aggregate_by_provider(df: pd.DataFrame, available_components: list[str]) -> pd.DataFrame:
    df = df.copy()
    df["_provider"] = df[PROVIDER_COL].apply(parse_provider)

    records = []
    for provider, group in df.groupby("_provider"):
        assessments = group.apply(assess_row, available_components=available_components, axis=1).tolist()

        n_total    = len(assessments)
        n_complete = sum(a["complete"] for a in assessments)
        rate       = round(n_complete / n_total * 100, 1) if n_total > 0 else 0.0

        missing_counts = {c: sum(1 for a in assessments if not a[c]) for c in COMPONENT_LABELS}

        assessed_missing = [(c, missing_counts[c]) for c in available_components if missing_counts[c] > 0]
        assessed_missing.sort(key=lambda x: -x[1])

        top1_key = assessed_missing[0][0] if len(assessed_missing) > 0 else None
        top2_key = assessed_missing[1][0] if len(assessed_missing) > 1 else None

        records.append({
            "provider_id":         provider,
            "eligible_patients":   n_total,
            "screened_patients":   n_complete,
            "screening_rate":      rate,
            "missing_lipids":      missing_counts["lipids"],
            "missing_bp":          missing_counts["bp"],
            "missing_smoking":     missing_counts["smoking"],
            "missing_bmi":         missing_counts["bmi"],
            "missing_a1c":         missing_counts["a1c"],
            "top_missing_1":       COMPONENT_LABELS[top1_key] if top1_key else None,
            "top_missing_2":       COMPONENT_LABELS[top2_key] if top2_key else None,
            "missing_count_1":     missing_counts[top1_key] if top1_key else 0,
            "missing_count_2":     missing_counts[top2_key] if top2_key else 0,
            "components_assessed": ", ".join(COMPONENT_LABELS[c] for c in available_components),
        })

    return pd.DataFrame(records)


# ── Main ───────────────────────────────────────────────────────────────────────

def main(input_path: str, output_path: str, password: str | None):
    print(f"\nReading: {input_path}")
    df = load_file(input_path, password)
    print(f"  Rows loaded: {len(df)}")
    print(f"  Report date: {REPORT_DATE.date()}  (12-month lookback window)")

    if PROVIDER_COL not in df.columns:
        print(f"\nERROR: Column '{PROVIDER_COL}' not found.", file=sys.stderr)
        print(f"Available columns: {df.columns.tolist()}", file=sys.stderr)
        sys.exit(1)

    # ── Filter 1: target diagnosis ─────────────────────────────────────────────
    def has_target_dx(problems):
        if pd.isna(problems):
            return False
        return bool(TARGET_DX_PATTERN.search(str(problems)))

    df["_provider"] = df[PROVIDER_COL].apply(lambda x: str(x).split("\n")[0].strip())
    before = len(df)
    df = df[df[PROBLEMS_COL].apply(has_target_dx)].copy()
    print(f"  After diagnosis filter: {len(df)} of {before} patients")

    # ── Filter 2: approved providers only ──────────────────────────────────────
    id_cols = [c for c in ["DOB", "Age", "Sex", "_provider"] if c in df.columns]
    unassigned = df[~df["_provider"].isin(APPROVED_PROVIDERS)][id_cols].copy()
    unassigned.rename(columns={"_provider": "Current Provider"}, inplace=True)
    if len(unassigned):
        unassigned_path = Path(output_path).parent / "unassigned_patients.csv"
        unassigned.to_csv(unassigned_path, index=False)
        print(f"\n  ⚠ {len(unassigned)} patients with target diagnosis but unapproved provider"
              f" → saved to {unassigned_path}")
        print(unassigned.to_string(index=False))
        print()

    before = len(df)
    df = df[df["_provider"].isin(APPROVED_PROVIDERS)].copy()
    print(f"  After provider filter:  {len(df)} of {before} patients")
    print(f"  Providers in scope: {sorted(df['_provider'].unique())}\n")

    print("Detecting available screening components:")
    available = detect_available_components(df)
    print(f"  Assessing: {', '.join(available)}\n")

    summary = aggregate_by_provider(df, available)
    summary.to_csv(output_path, index=False)

    # ── Print results ──────────────────────────────────────────────────────────
    print(f"{'─'*65}")
    print(f"{'Provider':<35} {'Eligible':>8} {'Screened':>9} {'Rate':>6}")
    print(f"{'─'*65}")
    for _, r in summary.sort_values("screening_rate", ascending=False).iterrows():
        flag = " ✓" if r["screening_rate"] >= 80 else ("  " if r["screening_rate"] >= 60 else " !")
        print(f"{r['provider_id']:<35} {int(r['eligible_patients']):>8} "
              f"{int(r['screened_patients']):>9} {r['screening_rate']:>5.1f}%{flag}")
    print(f"{'─'*65}")

    team_avg = summary["screening_rate"].mean()
    print(f"\n  Team average: {team_avg:.1f}%   Target: ≥80%")
    print(f"  Components assessed: {available}")
    print(f"\nSummary written to: {output_path}")

    # Delete the raw EMR export — it contains patient data and must not persist.
    input_path_obj = Path(input_path)
    if input_path_obj.exists():
        input_path_obj.unlink()
        print(f"Deleted raw export:  {input_path}  (patient data removed)\n")
    else:
        print()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--input",    default=INPUT_FILE)
    parser.add_argument("--output",   default=OUTPUT_FILE)
    parser.add_argument("--password", default=None,
                        help="Excel password (omit to be prompted)")
    args = parser.parse_args()

    password = args.password or getpass.getpass(f"Excel password for {args.input}: ")
    main(args.input, args.output, password)
