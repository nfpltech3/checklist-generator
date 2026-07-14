import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import re

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

class GSheetsClient:
    def __init__(self, credentials_path: str):
        self.credentials_path = credentials_path
        self.client = None
        self._authenticate()

    def _authenticate(self):
        try:
            credentials = Credentials.from_service_account_file(
                self.credentials_path, scopes=SCOPES
            )
            self.client = gspread.authorize(credentials)
        except Exception as e:
            raise Exception(f"Failed to authenticate with Google Sheets: {e}")

    def extract_sheet_id(self, url: str) -> str:
        """Extract the spreadsheet ID from a Google Sheets URL."""
        match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
        if match:
            return match.group(1)
        # If it's already an ID, just return it
        if len(url) > 20 and '/' not in url:
            return url
        raise ValueError("Invalid Google Sheets URL or ID")

    def get_sheet_data(self, url: str, sheet_name: str = "Master Sheet") -> pd.DataFrame:
        """Fetch data from Google Sheets and return as a pandas DataFrame."""
        sheet_id = self.extract_sheet_id(url)
        try:
            spreadsheet = self.client.open_by_key(sheet_id)
        except Exception as e:
            # Check for network connection errors
            err_str = str(e).lower()
            if any(term in err_str for term in ["nameresolutionerror", "failed to resolve", "getaddrinfo", "connectionerror", "max retries exceeded", "timeout", "unreachable"]):
                raise Exception(f"Network Error: Failed to connect to Google Sheets API. Please check your internet connection or proxy settings. (Error: {e})")
            raise Exception(f"Could not open spreadsheet. Make sure the service account email is shared on the sheet. Error: {e}")
        
        try:
            worksheet = spreadsheet.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            # Fallback to the first sheet if "Master Sheet" is not found
            worksheet = spreadsheet.sheet1

        data = worksheet.get_all_values()
        if not data:
            return pd.DataFrame()
        
        # Assume first row is header
        headers = data[0]
        rows = data[1:]
        df = pd.DataFrame(rows, columns=headers)
        return df, worksheet

    def append_rows(self, worksheet, rows: list):
        """Append new rows to the bottom of the worksheet."""
        if not rows:
            return
        worksheet.append_rows(rows, value_input_option='USER_ENTERED')

    def update_cells(self, worksheet, cell_updates: list):
        """
        Update specific cells.
        cell_updates is a list of dicts: [{'row': int, 'col': int, 'value': str}]
        (1-indexed for gspread)
        """
        if not cell_updates:
            return
        
        # Batch update is more efficient
        cells = []
        for update in cell_updates:
            cell = gspread.Cell(row=update['row'], col=update['col'], value=update['value'])
            cells.append(cell)
        
        worksheet.update_cells(cells, value_input_option='USER_ENTERED')

