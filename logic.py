import pandas as pd
import numpy as np

def normalize_string(val):
    if pd.isna(val) or val is None:
        return ""
    # Convert to string, strip whitespace, and upper case
    return str(val).strip().upper()

def normalize_number(val):
    if pd.isna(val) or val is None or str(val).strip() == "":
        return ""
    try:
        # Convert to float then string to handle formatting differences like "10" vs "10.0"
        return str(float(str(val).strip()))
    except ValueError:
        return str(val).strip().upper()

def compare_dataframes(df_item, df_master):
    """
    Compares the Item Report with the Master Sheet.
    Returns:
        mismatches: List of dicts representing field mismatches.
        new_models: List of dicts for models present in Item Report but missing in Master Sheet.
        master_row_map: Dictionary mapping Model to its row index (1-indexed) in Google Sheets.
    """
    # Columns to ignore during field-by-field comparison
    ignore_cols = {'JOB NO', 'BE DATE', 'MODEL', 'JOB DATE', 'INVOICE NO', 'INVOICE DATE', 'QUANTITY', 'AMOUNT', 'TOTAL PRICE'}
    
    # Standardize column names to upper case for mapping
    item_cols_upper = {c.upper(): c for c in df_item.columns}
    master_cols_upper = {c.upper(): c for c in df_master.columns}
    
    # Find overlapping columns that we need to verify
    verify_cols = []
    for m_col_upper, m_col_orig in master_cols_upper.items():
        if m_col_upper in item_cols_upper and m_col_upper not in ignore_cols:
            verify_cols.append((m_col_orig, item_cols_upper[m_col_upper]))
            
    # Normalize 'Model' column for joining
    df_item_copy = df_item.copy()
    df_master_copy = df_master.copy()
    
    if 'Model' not in df_item.columns:
        # Fallback if Model doesn't exist but maybe something else does, but requirements say Model is primary.
        # Let's assume Model is there. If not, raise error.
        raise ValueError("Item Report must contain a 'Model' column.")
    if 'Model' not in df_master.columns:
        raise ValueError("Master Sheet must contain a 'Model' column.")
        
    df_item_copy['match_key'] = df_item_copy['Model'].apply(normalize_string)
    df_master_copy['match_key'] = df_master_copy['Model'].apply(normalize_string)
    
    # Create a map of match_key to master sheet row index (Google Sheets is 1-indexed, header is row 1)
    # DataFrame index 0 corresponds to Google Sheet row 2.
    df_master_copy['gsheet_row'] = df_master_copy.index + 2
    master_row_map = df_master_copy.set_index('match_key')['gsheet_row'].to_dict()
    master_records = df_master_copy.set_index('match_key').to_dict('index')
    
    mismatches = []
    new_models = []
    seen_new_models = set()
    
    for idx, item_row in df_item_copy.iterrows():
        key = item_row['match_key']
        if not key:
            continue
            
        if key not in master_records:
            if key not in seen_new_models:
                seen_new_models.add(key)
                # It's a new model
                new_model_data = {c: item_row.get(c, "") for c in df_master.columns}
                new_model_data['Model'] = item_row['Model'] # Ensure Model is set correctly
                new_models.append({
                    'Model': item_row['Model'],
                    'Item Data': new_model_data,
                    'Job No': item_row.get('Job No', ''),
                    'Raw Row': item_row
                })
        else:
            # It's an existing model, compare fields
            master_row = master_records[key]
            for m_col, i_col in verify_cols:
                m_val = master_row.get(m_col, "")
                i_val = item_row.get(i_col, "")
                
                # Try numeric normalization first
                m_norm = normalize_number(m_val)
                i_norm = normalize_number(i_val)
                
                # If numeric parsing failed or gave different results, try string normalization
                if m_norm != i_norm:
                    m_str_norm = normalize_string(m_val)
                    i_str_norm = normalize_string(i_val)
                    if m_str_norm != i_str_norm:
                        # Mismatch found
                        mismatches.append({
                            'Model': item_row['Model'],
                            'Field': m_col,
                            'Master Value': m_val,
                            'Item Report Value': i_val,
                            'Job No': item_row.get('Job No', ''),
                            'gsheet_row': master_row['gsheet_row'],
                            'gsheet_col_name': m_col,
                            'match_key': key
                        })

    # Deduplicate mismatches (since Item Report might have multiple rows for same model with same mismatch)
    unique_mismatches = []
    seen_mismatches = set()
    for m in mismatches:
        sig = (m['match_key'], m['Field'], str(m['Master Value']), str(m['Item Report Value']))
        if sig not in seen_mismatches:
            seen_mismatches.add(sig)
            unique_mismatches.append(m)

    return unique_mismatches, new_models, list(df_master.columns)

