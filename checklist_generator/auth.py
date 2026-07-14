import gspread
import sys
import os

# Ensure the root directory is in the path to import gsheets
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from gsheets import GSheetsClient

# Global session cache
_CURRENT_SESSION = None

def get_current_session():
    return _CURRENT_SESSION

def set_current_session(username):
    global _CURRENT_SESSION
    _CURRENT_SESSION = username

def clear_session():
    global _CURRENT_SESSION
    _CURRENT_SESSION = None

def authenticate_user(url: str, cred_path: str, username: str, password: str):
    """
    Authenticates a user against the 'Tool_Users' tab in the Google Sheet.
    Validates plain text password and 'Can_Generate_Checklist' column.
    
    Returns: (bool success, str message)
    """
    if not username or not password:
        return False, "Username and password cannot be empty."

    try:
        client = GSheetsClient(cred_path)
        df, _ = client.get_sheet_data(url, sheet_name="Tool_Users")
        
        if df.empty:
            return False, "Tool_Users sheet is empty or not found."
            
        # Standardize column names to remove leading/trailing spaces just in case
        df.columns = df.columns.str.strip()
        
        # Check if required columns exist
        required_cols = ['Username', 'Password', 'Can_Generate_Checklist']
        for col in required_cols:
            if col not in df.columns:
                return False, f"Missing required column in Tool_Users: '{col}'"
                
        # Find user row
        user_row = df[df['Username'] == username]
        
        if user_row.empty:
            return False, "User not found."
            
        user_record = user_row.iloc[0]
        
        # Check password (plain text)
        if str(user_record['Password']) != password:
            return False, "Invalid password."
            
        # Check permissions
        can_generate = str(user_record['Can_Generate_Checklist']).strip().upper()
        if can_generate != 'Y':
            return False, "Access denied. You do not have permission to generate checklists."
            
        can_add_edit = False
        if 'Can_Add_Edit_Model' in user_record.index:
            can_add_edit = str(user_record['Can_Add_Edit_Model']).strip().upper() == 'Y'
            
        # Authentication successful
        set_current_session({'Username': username, 'Can_Add_Edit_Model': can_add_edit})
        return True, "Login successful."
        
    except Exception as e:
        return False, f"Authentication error: {str(e)}"
