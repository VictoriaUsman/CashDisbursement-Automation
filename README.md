# Cash Disbursement Consolidator

Extracts cash disbursement entries from one or more Google Sheets and writes them into a consolidated Google Sheet. Includes a General Ledger tab with SUMIF formulas. Runs as a Streamlit web app, deployable locally or on Google Cloud Run.

---

## Requirements

- Python 3.8+
- A Google Cloud Service Account with **Google Sheets API** and **Google Drive API** enabled

---

## Local Setup (One-Time)

### 1. Create a Google Cloud Service Account

1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create or select a project
3. Enable **Google Sheets API**, **Google Drive API**
4. Navigate to **IAM & Admin → Service Accounts → Create Service Account**
5. Download the JSON key and save it as `service_account.json` in this folder

> ⚠️ Never commit `service_account.json` to git — it is in `.gitignore`

### 2. Share Sheets with the Service Account

- Copy the `client_email` from `service_account.json`
- Share each **source sheet** → **Viewer**
- Share each **output sheet** → **Editor**

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Run

```bash
streamlit run app.py
```

Opens at `http://localhost:8501`.

---

## Deploying to Google Cloud Run

### Prerequisites

- Google Cloud project with billing enabled
- `gcloud` CLI installed and authenticated

```bash
gcloud auth login
gcloud config set project consolidation-495808
```

Enable required APIs:

```bash
gcloud services enable run.googleapis.com cloudbuild.googleapis.com \
  secretmanager.googleapis.com containerregistry.googleapis.com \
  --project=consolidation-495808
```

Grant the Cloud Run service account access to Secret Manager:

```bash
gcloud projects add-iam-policy-binding consolidation-495808 \
  --member="serviceAccount:640979934576-compute@developer.gserviceaccount.com" \
  --role="roles/secretmanager.secretAccessor"
```

### Deploy

```bash
./deploy.sh
```

`deploy.sh` will:
1. Upload `service_account.json` to Secret Manager (creates or updates)
2. Build and push the Docker image to Container Registry
3. Deploy to Cloud Run in `asia-southeast1`, injecting the secret at runtime

---

## UI Overview

### Source Sheets Table

Each row represents one source-to-output mapping:

| Field | Description |
|---|---|
| **Label** | Friendly name for this source (e.g. DAVAO, LOMA) |
| **Source Sheet ID** | ID of the Google Sheet to read from |
| **Source Tab** | Tab name inside the source sheet |
| **Output Sheet ID** | ID of the Google Sheet to write to |
| **Output Tab** | Tab name to write consolidated results into |

- **＋ Add Sheet** — add a new source row
- **✕** — remove a row
- **💾 Save Configuration** — persist all entries to `config.json` (reloads on next launch)

### Grouping Behavior

Sources sharing the same **Output Sheet ID + Output Tab** are consolidated into one tab with a shared header row.

### Generated Tabs (per output sheet)

| Tab | Description |
|---|---|
| Per source tab | Consolidated cash disbursement rows, filtered to entries with an Amount Paid |
| **General Ledger** | Unique account titles with SUMIF formulas summing all amount columns across all consolidated tabs |

### Buttons

| Button | Action |
|---|---|
| **🔍 Preview Data** | Fetches all source data and displays a table — no writes |
| **🚀 Run Consolidation** | Reads, filters, writes consolidated tabs + General Ledger |

### Filtering

Rows where **Amount Paid** is blank are automatically excluded.

---

## Finding a Sheet ID

```
https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit
```

---

## File Structure

```
Accounting/
├── app.py                # Streamlit web UI
├── consolidate.py        # Headless CLI script (legacy)
├── config.json           # Saved sheet configuration (auto-generated, safe to commit)
├── Dockerfile            # Container definition for Cloud Run
├── deploy.sh             # One-command Cloud Run deployment
├── requirements.txt      # Python dependencies
├── service_account.json  # Google credentials — DO NOT COMMIT
└── README.md
```

---

## Security Notes

- `service_account.json` is in `.gitignore` — never commit it
- On Cloud Run, credentials are injected via Secret Manager at runtime
- If credentials are ever accidentally pushed to git, rotate the key immediately at [console.cloud.google.com/iam-admin/serviceaccounts](https://console.cloud.google.com/iam-admin/serviceaccounts)
