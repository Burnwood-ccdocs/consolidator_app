import time
import re
import datetime
from typing import List, Tuple, Dict
import json
import hashlib
import os

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Import the list of all known appointment sheet URLs
# from add_to_sheet import SHEET_URLS as SOURCE_SHEET_URLS

# ---------------- CONFIGURATION ----------------
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets'
]

# The master spreadsheet whose column T lists the URLs of all appointment sheets
MASTER_SPREADSHEET_ID = '1JycrOH4URGHtLKDQhlP3Sf50-0rJtpqvih7NWvqR4Qo'
MASTER_SHEET_NAME = 'Active Clients'
COMPANY_COLUMN_NAME = 'Company'
URL_COLUMN_NAME = 'Appointment Spreadsheet:'

# Destination spreadsheet – change this to your own blank/target sheet
TARGET_SPREADSHEET_ID = '1Gk7SYNBXebnlgq4aH85PgxPo_BLVq9MkpcSi7ntq320'
TARGET_SHEET_NAME = 'Sheet1'          # Name of the tab that will receive the data
TARGET_START_RANGE = f'{TARGET_SHEET_NAME}!A1'

# How long to pause after finishing one source sheet (in seconds)
# SLEEP_BETWEEN_SHEETS = 2 - This is no longer used per sheet.

# Date range filtering - No longer used, replaced by hashing logic
# DATE_COLUMN_NAME = 'Date Submitted'
# START_DATE = datetime.date(2025, 5, 27)
# END_DATE = datetime.date(2025, 6, 12)
# ------------------------------------------------

PROCESSED_HASHES_FILE = 'processed_rows.json'


def load_processed_hashes() -> Dict[str, List[str]]:
    """Load the dictionary of processed row hashes from the JSON file."""
    if not os.path.exists(PROCESSED_HASHES_FILE):
        return {}
    with open(PROCESSED_HASHES_FILE, 'r') as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            # If file is empty or corrupt, start fresh
            return {}


def save_processed_hashes(hashes: Dict[str, List[str]]):
    """Save the dictionary of processed row hashes to the JSON file."""
    with open(PROCESSED_HASHES_FILE, 'w') as f:
        json.dump(hashes, f, indent=2)


def get_row_hash(row: List[str]) -> str:
    """Compute a SHA256 hash for a row to uniquely identify it."""
    row_str = "".join(map(str, row))
    return hashlib.sha256(row_str.encode('utf-8')).hexdigest()


def get_google_sheets_service():
    """Authenticate the service account and return a Google Sheets service object."""
    creds = service_account.Credentials.from_service_account_file(
        'service-account.json', scopes=SCOPES
    )
    return build('sheets', 'v4', credentials=creds)


def extract_spreadsheet_info_from_url(url: str) -> Tuple[str, int]:
    """Return (spreadsheet_id, gid) extracted from a Sheets URL.

    If gid is not present, it defaults to 0.
    """
    spreadsheet_id, gid = None, 0

    # Spreadsheet id
    patterns = [
        r"/spreadsheets/d/([a-zA-Z0-9-_]+)",  # Normal
        r"id=([a-zA-Z0-9-_]+)"                # Alternate link style
    ]
    for pattern in patterns:
        m = re.search(pattern, url)
        if m:
            spreadsheet_id = m.group(1)
            break

    # gid (sheet/tab id) if present
    m_gid = re.search(r"gid=(\d+)", url)
    if m_gid:
        gid = int(m_gid.group(1))

    return spreadsheet_id, gid


def get_sheet_name_from_gid(service, spreadsheet_id: str, gid: int) -> str:
    """Return the sheet/tab name that corresponds to the given gid.

    If gid isn't found, fall back to the first sheet.
    """
    try:
        meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        for sheet in meta.get('sheets', []):
            if sheet['properties']['sheetId'] == gid:
                return sheet['properties']['title']
        # Fallback – first sheet
        return meta['sheets'][0]['properties']['title']
    except HttpError as err:
        print(f"Failed to get metadata for {spreadsheet_id}: {err}")
    return None


def get_source_sheet_urls(service) -> List[Dict[str, str]]:
    """Read URLs and company names from the master spreadsheet by dynamically finding columns."""
    try:
        # 1. Get header row to find column indices
        header_res = service.spreadsheets().values().get(
            spreadsheetId=MASTER_SPREADSHEET_ID,
            range=f"'{MASTER_SHEET_NAME}'!1:1"
        ).execute()
        headers = header_res.get('values', [[]])[0]

        company_col_idx = -1
        url_col_idx = -1

        for i, header in enumerate(headers):
            if header.strip() == COMPANY_COLUMN_NAME:
                company_col_idx = i
            if header.strip() == URL_COLUMN_NAME:
                url_col_idx = i

        if company_col_idx == -1:
            print(f"Error: Could not find column '{COMPANY_COLUMN_NAME}' in master sheet '{MASTER_SHEET_NAME}'.")
            return []
        if url_col_idx == -1:
            print(f"Error: Could not find column '{URL_COLUMN_NAME}' in master sheet '{MASTER_SHEET_NAME}'.")
            return []

        print(f"Found '{COMPANY_COLUMN_NAME}' in column {chr(ord('A') + company_col_idx)} and '{URL_COLUMN_NAME}' in column {chr(ord('A') + url_col_idx)}")

        # 2. Get all data rows
        data_res = service.spreadsheets().values().get(
            spreadsheetId=MASTER_SPREADSHEET_ID,
            range=f"'{MASTER_SHEET_NAME}'!A2:ZZ"  # Read a wide range
        ).execute()
        values = data_res.get('values', [])

    except HttpError as err:
        print(f"Error reading master spreadsheet: {err}")
        return []

    urls_with_names: List[Tuple[str, str]] = []
    for row in values:
        if not row:
            continue

        if len(row) <= max(company_col_idx, url_col_idx):
            continue

        company_name = row[company_col_idx].strip() if len(row) > company_col_idx else ''
        cell_url_data = row[url_col_idx] if len(row) > url_col_idx else ''

        if not company_name or not cell_url_data:
            continue

        url = ''
        if 'http' in str(cell_url_data) and 'spreadsheets' in str(cell_url_data):
            # The cell itself is the URL
            url = str(cell_url_data).strip()
        else:
            # Search inside the cell text for a URL
            match = re.search(r'(https?://\S+)', str(cell_url_data))
            if match:
                url = match.group(1)

        if url:
            urls_with_names.append((url, company_name))

    # Remove accidental duplicates while preserving order
    seen_urls = set()
    unique_sources = []
    for url, name in urls_with_names:
        if url not in seen_urls:
            unique_sources.append({'url': url, 'name': name})
            seen_urls.add(url)

    print(f"Loaded {len(unique_sources)} appointment sheet sources from master spreadsheet")
    return unique_sources


def prepare_target_sheet(service, headers: List[str]):
    """
    Ensure the destination sheet exists and write headers only if the sheet is empty.
    If the sheet/tab doesn't exist yet, it is created.
    """
    # Check if the sheet exists; create it if necessary
    meta = service.spreadsheets().get(spreadsheetId=TARGET_SPREADSHEET_ID).execute()
    sheet_titles = [s['properties']['title'] for s in meta.get('sheets', [])]
    if TARGET_SHEET_NAME not in sheet_titles:
        req = {
            'addSheet': {
                'properties': {
                    'title': TARGET_SHEET_NAME,
                    'gridProperties': {
                        'rowCount': 1000,
                        'columnCount': 26  # A–Z at least
                    }
                }
            }
        }
        service.spreadsheets().batchUpdate(
            spreadsheetId=TARGET_SPREADSHEET_ID,
            body={'requests': [req]}
        ).execute()
        print(f"Created new sheet/tab '{TARGET_SHEET_NAME}' in target spreadsheet")

    # Check if sheet is empty by looking at A1
    try:
        res = service.spreadsheets().values().get(
            spreadsheetId=TARGET_SPREADSHEET_ID,
            range=f"{TARGET_SHEET_NAME}!A1:A1"
        ).execute()
    except HttpError as err:
        print(f"Error checking target sheet status: {err}")
        return

    if not res.get('values'):
        # Sheet is empty, write headers
        body = {'values': [headers]}
        service.spreadsheets().values().update(
            spreadsheetId=TARGET_SPREADSHEET_ID,
            range=TARGET_START_RANGE,
            valueInputOption='RAW',
            body=body
        ).execute()
        print(f"Target sheet '{TARGET_SHEET_NAME}' was empty, so headers were written.")
    else:
        print(f"Target sheet '{TARGET_SHEET_NAME}' already contains data. Headers not written.")


def append_rows(service, rows: List[List[str]]):
    """Insert rows at the top of the destination sheet (starting at row 2)."""
    if not rows:
        return
    
    try:
        # 1. First, get the current data to determine how many rows we need to insert
        result = service.spreadsheets().values().get(
            spreadsheetId=TARGET_SPREADSHEET_ID,
            range=f"'{TARGET_SHEET_NAME}'!A:A"
        ).execute()
        
        current_values = result.get('values', [])
        if not current_values:
            # If sheet is empty, just update directly
            body = {'values': rows}
            service.spreadsheets().values().update(
                spreadsheetId=TARGET_SPREADSHEET_ID,
                range=f"'{TARGET_SHEET_NAME}'!A2",
                valueInputOption='RAW',
                body=body
            ).execute()
            return
            
        # 2. Insert blank rows at row 2 (right after header)
        requests = [{
            'insertDimension': {
                'range': {
                    'sheetId': get_sheet_id_from_name(service, TARGET_SPREADSHEET_ID, TARGET_SHEET_NAME),
                    'dimension': 'ROWS',
                    'startIndex': 1,  # 0-based, so 1 = row 2
                    'endIndex': 1 + len(rows)
                },
                'inheritFromBefore': False
            }
        }]
        
        service.spreadsheets().batchUpdate(
            spreadsheetId=TARGET_SPREADSHEET_ID,
            body={'requests': requests}
        ).execute()
        
        # 3. Write the data to the newly inserted rows
        body = {'values': rows}
        service.spreadsheets().values().update(
            spreadsheetId=TARGET_SPREADSHEET_ID,
            range=f"'{TARGET_SHEET_NAME}'!A2",
            valueInputOption='RAW',
            body=body
        ).execute()
        
    except Exception as e:
        print(f"Error inserting rows at top of sheet: {e}")


def get_sheet_id_from_name(service, spreadsheet_id: str, sheet_name: str) -> int:
    """Get the sheetId (numeric ID) for a sheet by its name."""
    try:
        meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        for sheet in meta.get('sheets', []):
            if sheet['properties']['title'] == sheet_name:
                return sheet['properties']['sheetId']
        return 0  # Default to first sheet if not found
    except HttpError as err:
        print(f"Error getting sheet ID: {err}")
        return 0


def fetch_entire_sheet(service, spreadsheet_id: str, sheet_name: str) -> List[List[str]]:
    """Return all rows (including header) from the given sheet, columns A through Q only."""
    try:
        res = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f'{sheet_name}!A:Q'  # A-Q only to avoid formatting issues
        ).execute()
        return res.get('values', [])
    except HttpError as err:
        print(f"Error reading {spreadsheet_id} – {sheet_name}: {err}")
        return []


def get_all_sheet_titles(service, spreadsheet_id: str) -> List[str]:
    """Return the list of every tab title within the spreadsheet."""
    try:
        meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        return [s['properties']['title'] for s in meta.get('sheets', [])]
    except HttpError as err:
        print(f"Error fetching metadata for {spreadsheet_id}: {err}")
        return []


def parse_date_value(value: str) -> datetime.date:
    """Parse a date string into a datetime.date; returns None on failure."""
    for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%Y/%m/%d', '%m/%d/%y'):
        try:
            return datetime.datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    try:
        return datetime.date.fromisoformat(value)
    except Exception:
        return None


def main():
    if not TARGET_SPREADSHEET_ID:
        print("💥 ERROR: TARGET_SPREADSHEET_ID is not set in the configuration.")
        print("Please edit the script to provide the ID of the destination spreadsheet.")
        return

    service = get_google_sheets_service()

    while True:
        print("\n" + "="*50)
        print(f"Starting consolidation run at {datetime.datetime.now()}")
        print("="*50)

        processed_hashes = load_processed_hashes()
        source_infos = get_source_sheet_urls(service)

        if not source_infos:
            print("No appointment URLs to process. Waiting for next cycle.")
            time.sleep(4 * 60 * 60)
            continue

        all_new_rows_for_upload = []
        master_header_with_company = None
        newly_processed_this_run = {}
        headers_prepared_this_run = False

        for ss_index, source_info in enumerate(source_infos, start=1):
            url = source_info['url']
            company_name = source_info['name']
            print(f"\n📄 Checking {ss_index}/{len(source_infos)}: {company_name} ({url})")

            ss_id, gid = extract_spreadsheet_info_from_url(url)
            if not ss_id:
                print("  ⚠️  Couldn't parse spreadsheet ID – skipping.")
                continue

            sheet_name = get_sheet_name_from_gid(service, ss_id, gid)
            if not sheet_name:
                print("  ⚠️  Couldn't resolve sheet name for gid; skipping.")
                continue

            print(f"    → Reading tab '{sheet_name}' (gid {gid})")
            rows = fetch_entire_sheet(service, ss_id, sheet_name)
            if not rows or len(rows) <= 1:
                print("      (no data in sheet)")
                continue

            header = rows[0]
            data_rows = rows[1:]

            source_key = f"{ss_id}_{gid}"
            if source_key not in processed_hashes:
                processed_hashes[source_key] = []
            
            existing_hashes_for_source = set(processed_hashes[source_key])
            newly_found_rows = []
            
            if source_key not in newly_processed_this_run:
                newly_processed_this_run[source_key] = []

            for row in data_rows:
                row_hash = get_row_hash(row)
                if row_hash not in existing_hashes_for_source:
                    newly_found_rows.append(row)
                    newly_processed_this_run[source_key].append(row_hash)
            
            if not newly_found_rows:
                print("      (no new rows found)")
                time.sleep(3) # Brief pause to be nice to the API
                continue
            
            print(f"      Found {len(newly_found_rows)} new rows to process.")

            if not master_header_with_company:
                master_header_with_company = header + ['Company Name']

            for row in newly_found_rows:
                padded_row = row + [''] * (len(header) - len(row))
                all_new_rows_for_upload.append(padded_row + [company_name])
            
            time.sleep(3) # Brief pause to be nice to the API

            # --- Batch write logic ---
            # Write to sheet every 25 sheets or on the very last sheet
            if all_new_rows_for_upload and (ss_index % 25 == 0 or ss_index == len(source_infos)):
                print(f"\nProcessed {ss_index}/{len(source_infos)} sheets. Found a batch of {len(all_new_rows_for_upload)} new rows to write.")
                
                # Prepare headers only on the first write of the run
                if not headers_prepared_this_run and master_header_with_company:
                    prepare_target_sheet(service, master_header_with_company)
                    headers_prepared_this_run = True

                print("      ➕  Appending batch to the target sheet...")
                append_rows(service, all_new_rows_for_upload)
                print(f"      ✅  Successfully added {len(all_new_rows_for_upload)} rows.")

                # Update the persistent hash log *after* the successful write
                print("💾 Updating processed entries log for this batch...")
                for source_key, new_hashes in newly_processed_this_run.items():
                    if new_hashes:
                        if source_key not in processed_hashes:
                            processed_hashes[source_key] = []
                        processed_hashes[source_key].extend(new_hashes)
                save_processed_hashes(processed_hashes)
                print("      ...log updated.")

                # Reset for the next batch
                all_new_rows_for_upload = []
                newly_processed_this_run = {}

        print(f"\n✅ Consolidation run complete.")
        print(f"--- Sleeping for 4 hours until the next run (at {datetime.datetime.now() + datetime.timedelta(hours=4)}) ---")
        time.sleep(4 * 60 * 60)


if __name__ == '__main__':
    main() 