# Cash Disbursement Consolidator

Extracts cash disbursement entries from one or more Google Sheets and writes them into a consolidated Google Sheet. Includes a General Ledger tab with SUMIF formulas. Runs as a Streamlit web app, deployable locally or on Google Cloud Run.

---

## Requirements

- Python 3.8+
- A Google Cloud Service Account with **Google Sheets API** and **Google Drive API** enabled
- An **Anthropic API key** (for AI summary on the P&L page)

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

### 3. Configure Environment Variables

Create a `.env` file in this folder:

```
ANTHROPIC_API_KEY=sk-ant-api03-your-key-here
```

> ⚠️ Never commit `.env` to git — it is in `.gitignore`

### 4. Install Dependencies

```bash
pip install -r requirements.txt
```

### 5. Run

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

### Store Secrets in Secret Manager (One-Time)

Both secrets must exist in Secret Manager before the first deploy:

```bash
# Google service account credentials
gcloud secrets create service-account-json \
  --data-file=service_account.json \
  --project=consolidation-495808

# Anthropic API key
echo -n "sk-ant-api03-your-key-here" | \
  gcloud secrets create anthropic-api-key \
  --data-file=- \
  --project=consolidation-495808
```

`deploy.sh` will update these secrets automatically on subsequent deploys.

### Deploy

```bash
./deploy.sh
```

`deploy.sh` will:
1. Upload `service_account.json` to Secret Manager (creates or updates)
2. Build and push the Docker image to Container Registry
3. Deploy to Cloud Run in `asia-southeast1`, injecting both secrets at runtime

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

### Navigation Pages

| Page | Description |
|---|---|
| **📋 Consolidation** | Configure sources, preview data, run consolidation |
| **📒 General Ledger** | Per-account breakdown with subtotals; filter by month; download Excel |
| **📊 Disbursement Summary** | Pivot table with months as columns, accounts grouped by COS/OPEX/OTHER; download Excel |
| **📈 P&L Statement** | Income statement with revenues, COS, gross profit, OPEX, net income; AI summary button; download Excel |

### Buttons

| Button | Action |
|---|---|
| **🔍 Preview Data** | Fetches all source data and displays a table — no writes |
| **🚀 Run Consolidation** | Reads, filters, writes consolidated tabs |
| **⬇️ Download Excel** | Downloads the current page's data as a formatted `.xlsx` file |
| **✨ Generate AI Summary** | Uses Claude API to generate a 3-sentence P&L narrative (P&L page only) |

### Filtering

Rows where **Amount Paid** is blank are automatically excluded. In the General Ledger view, rows where **TB Amount** is zero are hidden.

### Account Classification

Accounts are classified into **COS**, **OPEX**, and **SALES** using a hardcoded `ACCOUNT_TAGS` dictionary in `app.py`. Use the **⚙️ Account Classification Override** expander on the P&L page to add accounts not covered by the hardcoded list. Overrides are saved to `config.json`.

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
├── .env                  # Local environment variables — DO NOT COMMIT
├── Dockerfile            # Container definition for Cloud Run
├── deploy.sh             # One-command Cloud Run deployment
├── requirements.txt      # Python dependencies
├── service_account.json  # Google credentials — DO NOT COMMIT
└── README.md
```

---

## Security Notes

- `service_account.json` and `.env` are in `.gitignore` — never commit them
- On Cloud Run, both secrets are injected via Secret Manager at runtime
- If credentials are ever accidentally pushed to git, rotate the key immediately:
  - Google: [console.cloud.google.com/iam-admin/serviceaccounts](https://console.cloud.google.com/iam-admin/serviceaccounts)
  - Anthropic: [console.anthropic.com](https://console.anthropic.com)
