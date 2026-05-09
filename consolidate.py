import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# ── CONFIG ──────────────────────────────────────────────────────────────────
SERVICE_ACCOUNT_FILE = "service_account.json"

SOURCES = [
    {"sheet_id": "1uGiFSN9L3bUl__o5fXQK9SNgv7l4fL7g", "tab": "DAVAO Cash Disbursement"},
    # Add more sources here later:
    # {"sheet_id": "...", "tab": "..."},
]

OUTPUT_SHEET_TITLE = f"Cash Disbursement Consolidated – {datetime.today().strftime('%Y-%m-%d')}"
# ────────────────────────────────────────────────────────────────────────────

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def main():
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)

    all_rows = []
    header = None

    for source in SOURCES:
        print(f"Reading: {source['sheet_id']} → {source['tab']}")
        ws = gc.open_by_key(source["sheet_id"]).worksheet(source["tab"])
        rows = ws.get_all_values()

        if not rows:
            print(f"  ⚠ No data found, skipping.")
            continue

        if header is None:
            header = rows[0]          # take header from first source
            all_rows.append(header)

        # Skip the header row for subsequent reads; append data rows only
        data_rows = rows[1:] if len(rows) > 1 else []
        all_rows.extend(data_rows)
        print(f"  ✓ {len(data_rows)} rows collected.")

    if not all_rows:
        print("No data to write. Exiting.")
        return

    print(f"\nCreating output sheet: '{OUTPUT_SHEET_TITLE}' ...")
    out_sheet = gc.create(OUTPUT_SHEET_TITLE)
    ws_out = out_sheet.sheet1
    ws_out.update("A1", all_rows)

    # Make header row bold
    ws_out.format("1:1", {"textFormat": {"bold": True}})

    # Share with your Google account so you can open it
    out_sheet.share("mcccpa@gmail.com", perm_type="user", role="writer")

    print(f"✓ Done! Sheet URL: https://docs.google.com/spreadsheets/d/{out_sheet.id}")


if __name__ == "__main__":
    main()
