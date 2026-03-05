# QI Screening Email Pipeline

A bimonthly provider feedback pipeline for quality improvement screening projects.
Processes encrypted EMR exports and sends personalized HTML emails to each provider showing their screening completion rate, a comparison with peers, and their most commonly missing components.

---

## Two ways to use this app

### Option A — Developer setup (original)
For team members with Python experience who work directly in the repository.

| | |
|---|---|
| **Who** | QI team lead, developer |
| **How** | `pip install -r requirements.txt` → `streamlit run app.py` |
| **Benefits** | Full control, easy to modify code and templates, can use CLI tools |
| **Requirements** | Python installed, comfortable with terminal |

### Option B — Portable app (no install, no admin rights)
A self-contained folder distributed via USB or shared drive for colleagues without a programming background.

| | |
|---|---|
| **Who** | Non-programmer colleagues on institutional workstations |
| **How** | Double-click `INSTALL.bat` once → double-click the Desktop shortcut to run |
| **Benefits** | No admin rights needed, no Python installation, one-click updates via `UPDATE.bat` |
| **Requirements** | Windows, internet connection for first-time setup only |

> The portable folder (`portable_app/`) is excluded from this repository because it includes a bundled Python runtime. Distribute it via USB or a shared institutional drive.

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

## Setup (Option A — Developer)

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

> Neither the Excel password nor the SMTP password is stored in `.env`. Both are prompted at runtime.

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

## Setup (Option B — Portable app)

1. Copy the `portable_app/` folder to the target computer (USB or shared drive)
2. Double-click **`INSTALL.bat`** — downloads Python (~25 MB), installs packages, creates a Desktop shortcut
3. To apply future updates: double-click **`UPDATE.bat`** after being notified of a new version

---

## Running the app (Streamlit)

```bash
streamlit run app.py
```

Opens at `http://localhost:8501` with three tabs:

| Tab | Purpose | Run on |
|---|---|---|
| 📊 Process Data | Upload EMR export, enter password, generate summary | Institutional workstation |
| 📧 Preview Email | Upload summary, preview any provider's email | Any machine |
| 🚀 Send Emails | Upload summary, send to selected providers | Any machine |

The app detects whether it is running locally or on Streamlit Cloud and displays a warning banner accordingly. **Never upload patient data to a cloud-hosted instance.**

---

## Running via CLI

### Process data (institutional workstation)

```bash
python process_data.py
# or with explicit paths:
python process_data.py --input data/export.xlsx --output data/processed_summary.csv
```

You will be prompted for the Excel password. The raw `.xlsx` is deleted automatically after processing.

### Preview emails

```bash
python preview.py --provider "SMITH, JANE A" --period "May–Jun 2026"
python preview.py --all --period "May–Jun 2026"
```

### Send emails

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

- **Rate card** — their screening completion rate, color-coded against the target
- **Top performer banner** — shown only to the highest-scoring provider this period
- **Comparison bars** — their rate vs. attending average, fellow average, and target
- **Component breakdown** — completion rate per individual screening component
- **Top 2 missing components** — the most commonly unscreened items across their patients
- **Nudge** — a short call to action with a link to the dashboard (if configured)

---

## Security and privacy

### What data is handled

- **EMR export (.xlsx)** — contains patient-level PHI (names, diagnoses, visit data). This file lives only on the institutional workstation and is deleted automatically by `process_data.py` immediately after processing.
- **processed_summary.csv** — contains provider-level aggregates only (counts, rates, missing component labels). No patient names or IDs. This is the only file that leaves the institutional workstation.
- **provider_list.csv** — contains provider names and email addresses. Not committed to version control.

### What leaves the workstation

Only `processed_summary.csv` (aggregate statistics, no PHI) is transferred to the sending machine. No patient-identifiable information is included.

### Credentials

| Credential | Stored? | How handled |
|---|---|---|
| EMR Excel password | Never | Prompted at runtime, held in memory only |
| SMTP password | Never | Prompted at runtime, held in memory only |
| SMTP username | `.env` only | Not committed (gitignored) |

### Deployment rules

| Environment | Allowed? | Notes |
|---|---|---|
| Local — institutional workstation | Yes | Recommended. Data never leaves the machine. |
| Local — personal computer | With caution | Only if the device is covered by your institution's BAA and data security policy. |
| Streamlit Cloud / any public cloud | No | The app will display a red warning banner if it detects a cloud environment. Do not upload patient data. |

### Code safeguards

1. **Jinja2 HTML autoescape enabled** — prevents injection from provider names or patient data rendered in email templates
2. **Raw EMR export auto-deleted** — `process_data.py` removes the input file immediately after writing the summary
3. **All data files gitignored** — `.gitignore` blocks all `.csv`, `.xlsx`, `.xls`, `output/`, and `winpython/` from being committed
4. **Duplicate-send guard** — `send_log.csv` prevents accidental duplicate emails per period
5. **Recipient confirmation** — the CLI and app require explicit confirmation before any email is sent
6. **Cloud detection** — the app detects Streamlit Cloud (`HOME=/home/appuser`) and blocks use with a visible error banner

### HIPAA considerations

This tool is designed for use within an institutional environment covered by your organization's data security policies. Before use, confirm with your compliance office or IRB that:

- The workstation used for data processing is covered under your institution's security plan
- Any inter-device file transfer (e.g., USB, shared drive) complies with your data transfer policy
- The email system used for sending (SMTP server) is approved for provider-facing communications

Patient data is **never** transmitted to GitHub, Streamlit Cloud, or any external service by this application.
