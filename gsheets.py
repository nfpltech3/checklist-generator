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

    def add_pending_models(self, url: str, pending_rows: list):
        """
        Add missing models to the Pending_Model_Approval sheet if they don't already exist.
        pending_rows is a list of dicts with keys matching the sheet headers.
        Columns: Model, Product Desc, Country of Origin, Generic Description, Importer, Added_By, Added_At, Status
        """
        if not pending_rows:
            return
            
        try:
            df, ws = self.get_sheet_data(url, "Pending_Model_Approval")
        except Exception as e:
            print(f"Failed to fetch Pending_Model_Approval sheet: {e}")
            return
            
        headers = df.columns.tolist() if not df.empty else ["Model", "Product Desc", "Country of Origin", "Generic Description", "Importer", "Added_By", "Added_At", "Status"]
        
        # Check for existing pending items to avoid duplicates
        # We consider a duplicate if Model + Importer match and Status is not 'Updated to Master'
        existing_pending = set()
        if not df.empty:
            pending_df = df[df['Status'].str.strip().str.lower() != 'updated to master']
            for _, row in pending_df.iterrows():
                existing_pending.add((str(row.get('Model', '')).strip().lower(), str(row.get('Importer', '')).strip().lower()))
                
        new_sheet_rows = []
        for pr in pending_rows:
            m = str(pr.get('Model', '')).strip().lower()
            imp = str(pr.get('Importer', '')).strip().lower()
            if (m, imp) not in existing_pending:
                row_list = [str(pr.get(h, "")) for h in headers]
                new_sheet_rows.append(row_list)
                existing_pending.add((m, imp))
                
        if new_sheet_rows:
            self.append_rows(ws, new_sheet_rows)

    def update_pending_model_row(self, url: str, model: str, importer: str, updates_dict: dict):
        """
        Update specific fields of a pending model in the Pending_Model_Approval sheet.
        Finds the row where Model and Importer match and Status is not 'Updated to Master'.
        """
        try:
            df, ws = self.get_sheet_data(url, "Pending_Model_Approval")
        except Exception as e:
            print(f"Failed to fetch for update_pending_model_row: {e}")
            return
            
        if df.empty:
            return
            
        m_lower = str(model).strip().lower()
        imp_lower = str(importer).strip().lower()
        
        # Find row to update
        target_idx = None
        for idx, row in df.iterrows():
            if str(row.get('Status', '')).strip().lower() != 'updated to master':
                if str(row.get('Model', '')).strip().lower() == m_lower and str(row.get('Importer', '')).strip().lower() == imp_lower:
                    target_idx = idx
                    break
                    
        if target_idx is None:
            return
            
        sheet_row = target_idx + 2 # 0-indexed idx -> 2-indexed sheet row (1 for header + 1 for idx offset)
        
        updates = []
        for col_name, value in updates_dict.items():
            for i, col in enumerate(df.columns):
                if str(col).strip().lower() == str(col_name).strip().lower():
                    updates.append({'row': sheet_row, 'col': i + 1, 'value': value})
                    break
                    
        if updates:
            self.update_cells(ws, updates)

    def mark_pending_model_resolved(self, url: str, model: str, importer: str):
        """
        Mark a pending model as Updated to Master in the Pending_Model_Approval sheet.
        """
        try:
            df, ws = self.get_sheet_data(url, "Pending_Model_Approval")
        except Exception:
            return
            
        if df.empty:
            return
            
        m_lower = str(model).strip().lower()
        imp_lower = str(importer).strip().lower()
        
        status_col_idx = None
        for i, col in enumerate(df.columns):
            if str(col).strip().lower() == "status":
                status_col_idx = i + 1
                break
                
        if not status_col_idx:
            return
            
        updates = []
        for idx, row in df.iterrows():
            status = str(row.get('Status', '')).strip().lower()
            if status != 'updated to master' and status != 'resolved':
                if str(row.get('Model', '')).strip().lower() == m_lower and str(row.get('Importer', '')).strip().lower() == imp_lower:
                    # idx is 0-based for df, meaning row 2 in sheet (since row 1 is header)
                    sheet_row = idx + 2
                    updates.append({'row': sheet_row, 'col': status_col_idx, 'value': 'Updated to Master'})
                    
        if updates:
            self.update_cells(ws, updates)

    def update_mapping_last_used(self, url: str, importer: str, format_name: str):
        """
        Update the Last_Used_At timestamp for a specific Format_Name in Invoice_Header_Mappings.
        """
        from datetime import datetime
        try:
            df, ws = self.get_sheet_data(url, "Invoice_Header_Mappings")
        except Exception as e:
            print(f"Failed to fetch Invoice_Header_Mappings sheet for update: {e}")
            return
            
        if df.empty:
            return
            
        imp_lower = str(importer).strip().lower()
        fmt_lower = str(format_name).strip().lower()
        
        last_used_col_idx = None
        for i, col in enumerate(df.columns):
            if str(col).strip().lower() == "last_used_at":
                last_used_col_idx = i + 1
                break
                
        if not last_used_col_idx:
            # If the column doesn't exist, we skip updating
            return
            
        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        updates = []
        
        for idx, row in df.iterrows():
            if str(row.get('Importer', '')).strip().lower() == imp_lower and str(row.get('Format_Name', '')).strip().lower() == fmt_lower:
                # idx is 0-based for df, meaning row 2 in sheet
                sheet_row = idx + 2
                updates.append({'row': sheet_row, 'col': last_used_col_idx, 'value': current_time})
                
        if updates:
            self.update_cells(ws, updates)
            
    def save_new_header_mapping(self, url: str, mappings: list):
        """
        Append new header mapping rows to Invoice_Header_Mappings.
        mappings should be a list of dicts with: Importer, Format_Name, Source_Header, Target_Field, Last_Used_At
        """
        if not mappings:
            return
            
        try:
            df, ws = self.get_sheet_data(url, "Invoice_Header_Mappings")
        except Exception as e:
            print(f"Failed to fetch Invoice_Header_Mappings sheet for save: {e}")
            return
            
        headers = df.columns.tolist() if not df.empty else ["Importer", "Format_Name", "Source_Header", "Target_Field", "Last_Used_At"]
        
        new_sheet_rows = []
        for m in mappings:
            row_list = [str(m.get(h, "")) for h in headers]
            new_sheet_rows.append(row_list)
            
        if new_sheet_rows:
            self.append_rows(ws, new_sheet_rows)
