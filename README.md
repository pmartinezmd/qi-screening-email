# QI Screening Email Pipeline

A bimonthly provider feedback pipeline for quality improvement screening projects.
Processes encrypted EMR exports and sends personalized HTML emails to each provider showing their screening completion rate, a comparison with peers, and their most commonly missing components.

---

## How it works

```
Institutional workstation (PHI stays here)      Personal laptop (no PHI)
──────────────────────────────────────────      ────────────────────────
Pull EMR export
Run process_data.py (or Tab 1 in the app)
  → Excel password prompted at runtime
  → Diagnoses + provider filter applied
  → Raw .xlsx deleted automatically
  → processed_summary.csv created          ──copy──►  processed_summary.csv
                                                        Run send_emails.py
                                                        (or Tab 3 in the app)
                                                          → prompted for password
                                                          → confirm recipient list
                                                          → emails sent
                                                          → send_log.csv updated
```

---

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure `.env`

```bash
cp .env.example .env
```

Edit `.env` with your settings:

```
SMTP_HOST=smtp.office365.com
SMTP_PORT=587
SMTP_USER=your.email@institution.edu
FROM_NAME=Screening QI Team
TEAM_LABEL=QI Team · Your Department
SCREENING_NAME=CV Risk Screening QI
DASHBOARD_URL=                         # optional
```

> Neither the Excel password nor the SMTP password is stored in `.env`.
> Both are prompted at runtime.

### 3. Create `data/provider_list.csv`

Copy `provider_list.example.csv` to `data/provider_list.csv` and fill in your team:

```
provider_id,display_name,email,provider_type
SMITH, JANE A,Dr. Jane Smith,jsmith@institution.edu,attending
DOE, JOHN B,Dr. John Doe,jdoe@institution.edu,fellow
```

- `provider_id` — must match the exact value in your EMR export (first line of the Encounter Provider field)
- `provider_type` — `attending` or `fellow` (used for group comparison bars)

### 4. Configure the diagnosis filter and providers

Open `process_data.py` and update the **CONFIGURE** sections:

- **`TARGET_DX_PATTERN`** — regex matching your patient population's diagnoses (the `Problems` column in the EMR export)
- **`APPROVED_PROVIDERS`** — set of provider names exactly as they appear in the EMR export
- **Column names** — update `PROVIDER_COL`, `PROBLEMS_COL`, and component column names if your EMR export uses different headers

---

## Running the app (Streamlit)

```bash
streamlit run app.py
```

Opens at `http://localhost:8501` with three tabs:

| Tab | Purpose | Run on |
|---|---|---|
| 📊 Process Data | Upload EMR export, enter password, generate summary | Institutional workstation |
| 📧 Preview Email | Upload summary, preview any provider's email | Personal laptop |
| 🚀 Send Emails | Upload summary, send to selected providers | Personal laptop |

---

## Running via CLI

### Process data (institutional workstation)

```bash
python process_data.py
# or with explicit paths:
python process_data.py --input data/export.xlsx --output data/processed_summary.csv
```

You will be prompted for the Excel password. The raw `.xlsx` is deleted automatically after processing.

### Preview emails (personal laptop)

```bash
python preview.py --provider "SMITH, JANE A" --period "May–Jun 2026"
python preview.py --all --period "May–Jun 2026"
```

### Send emails (personal laptop)

```bash
python send_emails.py --period "May–Jun 2026"
```

You will be prompted for the sender address (defaults to `SMTP_USER` in `.env`), your password, and a recipient confirmation before any email is sent.

Safe to re-run — providers already sent for the current period are automatically skipped. To resend, remove their row from `data/send_log.csv`.

---

## Data files

| File | Contains | Commit? |
|---|---|---|
| EMR export (`.xlsx`) | Patient-level PHI | Never |
| `data/processed_summary.csv` | Provider-level aggregates only, no patient IDs | No (gitignored) |
| `data/provider_list.csv` | Provider names and emails | No (gitignored) |
| `data/send_log.csv` | Send timestamps per provider per period | No (gitignored) |
| `.env` | SMTP config (no passwords) | No (gitignored) |

All data files in `data/` and `output/` are gitignored. No patient data can be accidentally committed.

---

## Email content

Each provider receives:

- **Rate card** — their screening completion rate, color-coded against the 80% target
- **Top performer banner** — shown only to the highest-scoring provider this period
- **Comparison bars** — their rate vs. attending average, fellow average, and 80% target
- **Top 2 missing components** — the most commonly unscreened items across their patients
- **Nudge** — a short call to action with a link to the dashboard (if configured)
- **FAQ** — answers to common questions about eligibility and data

---

## Privacy safeguards

1. **No passwords stored** — Excel and SMTP passwords are prompted at runtime only
2. **Raw EMR export auto-deleted** — `process_data.py` deletes the input file immediately after writing the summary
3. **PHI stays on the institutional workstation** — only the aggregate `processed_summary.csv` is transferred
4. **All data files gitignored** — no patient data or credentials can be committed
5. **Duplicate-send guard** — `send_log.csv` prevents accidental duplicate emails per period
6. **Recipient confirmation** — the CLI requires explicit confirmation before sending
