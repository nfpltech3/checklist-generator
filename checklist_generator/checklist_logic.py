import gspread
import pandas as pd
import openpyxl
from openpyxl.styles.numbers import FORMAT_TEXT
from datetime import datetime
import sys
import os
import csv

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from gsheets import GSheetsClient

# The exact literal headers extracted from the Logisys template
LOGISYS_HEADERS = ['Inv No', 'Date', 'TOI', 'Product Desc', 'Quantity', 'Unit', 'Currency', 'Rate', 'CTH', 'Basic notification', 'Sr no', 'CETH', 'Excise notification', 'Sr no', 'Mrp Sr no', 'Amount', 'SAD notification', 'Sr no', 'Exim Code', 'Scheme Notn', 'Scheme SrNo', 'Generic Description', 'BRAND', 'Model', 'END USE', 'EduCessNotnNo', 'EduCessNotnSrNo ', 'SHEduCessNotnNo', 'SHEduCessNotnSrNo', 'CEXEduCessNotnNo', 'CEXEduCessNotnSrNo ', 'Country of origin', 'RSP Notf', 'Sr. no.', 'Sapta Notif No', 'SAPTA Sr No', 'Manufacturer', 'Address', 'MRP Status', 'Abatement Rate', 'IGST Notn No', 'IGST SrNo', 'GST CCESS Notn No', 'GST CCESS SrNo', 'IGST Exemp Notn Type', 'IGST Exemp Notn No', 'IGST Exemp SrNo', 'GST CCESS Exemp NotnType', 'GST CCESS Exemp Notn No', 'GST CCESS Exemp SrNo', 'Manufacturer Cntry', 'Manufacturer State', 'Manuf Postal Code', 'SWS Notn No', 'SWS Notn SrNo', 'Duty Exemption Type', 'Regn No.', 'Regn Date', 'Regn Port', 'Item SNo. In Lic.', 'Product Material Code', 'Type Of Item', 'Source Country', 'Transit Country', 'SQC_Qty', 'Consigner Name', 'Consigner Address', 'Consigner City', 'Consigner Country', 'SIMS Category', 'SIMS Code', 'Custom Health Cess Notn No', 'Custom Health Cess Notn SrNo', 'COO_No ', 'COO_Date of Issue ', 'COO_Issuing Country ', 'COO_Origin Criteria ', 'COO_Origin Criteria remarks ', 'COO_Accumulation/Cumulation Status ', 'FOC Item', 'AIDC Levy Notification No ', 'AIDC Levy Notification Sr.No. ', 'AIDC Exemption Notification No. ', 'AIDC Exemption Notification Sr.No. ', 'AIDCNotn_Excise', 'ADICNotnSr_Excise', 'Previous BE No.', 'Previous BE Date', 'Previous BE IGM No.', 'Previous BE IGM Date', 'Previous BE Currency', 'Previous BE Unit Price', 'Previous BE Custom House', 'Aggregate Duty Notification No.', 'Aggregate Duty Notification Sr.No.', 'COO_Tariff Shift', 'COO_Retroactive Issuance Status', 'COO_Direct Consignment Status', 'GST_Comp_Cess_SalePrice_INR', 'CVD_Notn', 'CVD_SLNo', 'CVD_Rate', 'CVD_Calc', 'CVD_ItemSLNo', 'CVD_SuppSLNo', 'COO_ItemSrNoCert']

def lookup_model_local(df: pd.DataFrame, model: str):
    """
    Looks up a model in the preloaded dataframe.
    Matches strictly as string, ignoring case and all whitespace (internal, leading, and trailing).
    Returns a list of dictionaries for all matching rows.
    """
    if df.empty or 'Model' not in df.columns:
        return []
        
    def norm(x):
        return "".join(str(x).split()).lower()
        
    target_norm = norm(model)
    matches = df[df['Model'].astype(str).apply(norm) == target_norm]
    return matches.to_dict('records')

def export_checklist(data_rows: list, output_path: str):
    """
    Exports the data rows to a CSV file, exactly matching LOGISYS_HEADERS.
    """
    with open(output_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        
        # 1. Write exact headers
        writer.writerow(LOGISYS_HEADERS)
        
        # 2. Write data rows
        for data in data_rows:
            row = []
            for col_idx in range(1, len(LOGISYS_HEADERS) + 1):
                col_key = f"col_{col_idx}"
                row.append(data.get(col_key, ""))
            writer.writerow(row)

def log_audit_action(url: str, cred_path: str, user: str, importer: str, models: list, action_type: str = "checklist_generated"):
    """
    Logs an action to the Audit_Log tab in Google Sheets.
    """
    try:
        client = GSheetsClient(cred_path)
        sheet_id = client.extract_sheet_id(url)
        spreadsheet = client.client.open_by_key(sheet_id)
        
        try:
            worksheet = spreadsheet.worksheet("Audit_Log")
        except gspread.exceptions.WorksheetNotFound:
            audit_headers = ["Timestamp", "User", "Action", "Importer", "Models", "Field_Changed", "Old_Value", "New_Value"]
            worksheet = spreadsheet.add_worksheet(title="Audit_Log", rows=1000, cols=len(audit_headers))
            worksheet.append_row(audit_headers, value_input_option='USER_ENTERED')
            
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        models_str = ", ".join([str(m) for m in models])
        
        # User might be a dict now (from auth session)
        user_name = user.get('Username', str(user)) if isinstance(user, dict) else str(user)
        
        row = [
            timestamp,
            user_name,
            action_type,
            importer,
            models_str,
            "", # Field_Changed
            "", # Old_Value
            ""  # New_Value
        ]
        
        client.append_rows(worksheet, [row])
        
    except Exception as e:
        print(f"Warning: Failed to log audit action: {str(e)}")
