import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys
import threading
from pathlib import Path
from datetime import datetime
PENDING_HEADERS = [
    "Model", "Product Desc", "Country of Origin",
    "CTH", "Basic Duty Notn", "Basic Duty Notn SNo", "Basic Duty Rate", 
    "Customs Health Cess Notn", "Customs Health Cess SNo", "Customs Health Cess Rate", 
    "SWS Notification", "SWS Notification SrNo", "SWS Rate", "IGST Notification", 
    "IGST Notification SrNo", "IGST Rate", "AIDC Notn (Customs)", "AIDC Notn Sr.No.(Customs)", 
    "End Use", "Generic Description", "Added_By", "Added_At", "Importer"
]

MASTER_SHEET_COLUMNS = [
    "Job No", "BE Date", "Model", "Product Desc", "CTH", 
    "Basic Duty Notn", "Basic Duty Notn SNo", "Basic Duty Rate", 
    "Customs Health Cess Notn", "Customs Health Cess SNo", "Customs Health Cess Rate", 
    "SWS Notification", "SWS Notification SrNo", "SWS Rate", 
    "IGST Notification", "IGST Notification SrNo", "IGST Rate", 
    "AIDC Notn (Customs)", "AIDC Notn Sr.No.(Customs)", 
    "End Use", "Generic Description", "Country of Origin"
]

# Import tksheet
from tksheet import Sheet

# Adjust path to import auth and logic
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
import auth
import checklist_logic

# Brand Constants
_PRIMARY_BLUE = "#1F3F6E"; _ACCENT_RED = "#D8232A"; _DARK_TEXT = "#1F2937"
_MUTED_GRAY = "#6B7280"; _LIGHT_BG = "#F4F6F8"; _PANEL_WHITE = "#FFFFFF"
_BORDER_GRAY = "#E5E7EB"; _HOVER_BLUE = "#2A528F"; _HEADER_BG = "#D6E4F0"

# 1-based indices of editable manual columns
EDITABLE_INDICES = {1, 2, 3, 5, 6, 7, 8, 16, 23, 32}
# 1-based indices of master-sourced columns
MASTER_INDICES = {4, 9, 10, 11, 22, 24, 25, 41, 42, 54, 55, 81, 82}

# Column Widths classifications
NARROW_COLS = {'Quantity', 'Unit', 'Currency', 'Rate', 'CTH', 'Sr no', 'IGST SrNo', 'SWS SrNo'}
MEDIUM_COLS = {'Inv No', 'Date', 'TOI', 'Model', 'BRAND', 'Basic notification', 'IGST Notn No', 'SWS Notn No'}
WIDE_COLS = {'Product Desc', 'Generic Description', 'END USE', 'Country of origin'}

def wrap_header(text, max_len=15):
    """Returns the cleaned header text on a single line (no wrapping)."""
    return str(text).strip()

# Pre-calculate column widths
GRID_WIDTHS = []
for h in checklist_logic.LOGISYS_HEADERS:
    h_clean = h.strip()
    if h_clean in NARROW_COLS:
        GRID_WIDTHS.append(80)
    elif h_clean in MEDIUM_COLS:
        GRID_WIDTHS.append(140)
    elif h_clean in WIDE_COLS:
        GRID_WIDTHS.append(250)
    else:
        GRID_WIDTHS.append(100)


def get_base_path() -> Path:
    return Path(sys._MEIPASS) if getattr(sys, 'frozen', False) else Path(__file__).parent.parent

def _brand_button(parent, text: str, command, state=tk.NORMAL, bg_color=_PRIMARY_BLUE, secondary=False) -> tk.Button:
    if secondary:
        btn = tk.Button(parent, text=text, command=command, state=state,
                        font=("Segoe UI",10,"bold"), fg=_PRIMARY_BLUE, bg=_PANEL_WHITE,
                        activebackground=_LIGHT_BG, activeforeground=_PRIMARY_BLUE,
                        bd=1, relief=tk.SOLID, highlightthickness=0, highlightbackground=_PRIMARY_BLUE,
                        padx=14, pady=5, cursor="hand2")
    else:
        btn = tk.Button(parent, text=text, command=command, state=state,
                        font=("Segoe UI",10,"bold"), fg="#FFF", bg=bg_color,
                        activebackground=_HOVER_BLUE, activeforeground="#FFF",
                        bd=0, padx=14, pady=6, cursor="hand2", relief=tk.FLAT)
        if bg_color == _PRIMARY_BLUE:
            btn.bind("<Enter>", lambda e: btn.configure(bg=_HOVER_BLUE) if btn['state'] != 'disabled' else None)
            btn.bind("<Leave>", lambda e: btn.configure(bg=_PRIMARY_BLUE) if btn['state'] != 'disabled' else None)
    return btn

def align_sheet_columns(sheet, headers):
    for idx, header in enumerate(headers):
        h_clean = str(header).strip().lower()
        if any(term in h_clean for term in ["qty", "quantity", "rate", "amount", "sno", "srno", "price", "duty", "val"]):
            sheet.align_columns(columns=idx, align="e")
        elif any(term in h_clean for term in ["desc", "model", "generic", "use", "brand", "importer", "country"]):
            sheet.align_columns(columns=idx, align="w")
        else:
            sheet.align_columns(columns=idx, align="center")

def get_flexible(d, key):
    """Retrieves dict values matching key case-insensitively and stripped of whitespace."""
    if key in d:
        return d[key]
    key_clean = str(key).strip().lower()
    for k, v in d.items():
        if str(k).strip().lower() == key_clean:
            return v
    return ""


class ChecklistApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Nagarkot Checklist Template Generator")
        self.root.configure(bg=_LIGHT_BG)
        self.root.state("zoomed")
        self.root.minsize(1024, 600)
        
        # Configure standard TTK styles
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('TEntry', padding=6, font=("Segoe UI", 10), bordercolor=_BORDER_GRAY, lightcolor=_BORDER_GRAY, darkcolor=_BORDER_GRAY, fieldbackground=_PANEL_WHITE)
        self.style.configure('TCombobox', padding=6, font=("Segoe UI", 10), bordercolor=_BORDER_GRAY, lightcolor=_BORDER_GRAY, darkcolor=_BORDER_GRAY, fieldbackground=_PANEL_WHITE)
        self.style.configure('TNotebook', background=_LIGHT_BG)
        self.style.configure('TNotebook.Tab', padding=(16, 6), font=("Segoe UI", 10, "bold"), background=_PANEL_WHITE, foreground=_MUTED_GRAY)
        self.style.map('TNotebook.Tab', 
                       background=[('selected', _PRIMARY_BLUE)], 
                       foreground=[('selected', '#FFFFFF')],
                       padding=[('selected', (16, 6))])
        
        self.gsheets_url = tk.StringVar(value="https://docs.google.com/spreadsheets/d/1nZyxJGCwG7bK8-wg2Tti9l1C3sS5KcULefBvquW3u8E/edit?gid=1228640769#gid=1228640769")
        self.cred_path = tk.StringVar(value=str(get_base_path() / "credentials.json"))
        
        self.importer = None
        self.importer_df = None # Cached importer DataFrame
        self.rows_data = [] # List of dicts representing each added model row
        
        # 0-based index configs for tksheet
        self.editable_cols_0based = {i - 1 for i in EDITABLE_INDICES}
        self.readonly_cols_0based = [i for i in range(len(checklist_logic.LOGISYS_HEADERS)) if i not in self.editable_cols_0based]
        
        # Build main container
        self._build_header()
        
        self.body_container = tk.Frame(self.root, bg=_LIGHT_BG)
        self.body_container.pack(fill=tk.BOTH, expand=True, padx=16, pady=(12,4))
        
        self._build_footer()
        
        self.show_login_screen()

    def _build_header(self) -> None:
        header = tk.Frame(self.root, bg=_PANEL_WHITE, bd=0, highlightthickness=1, highlightbackground=_BORDER_GRAY)
        header.pack(fill=tk.X, side=tk.TOP); header.pack_propagate(False); header.configure(height=64)
        
        # Logo Left
        left_frame = tk.Frame(header, bg=_PANEL_WHITE)
        left_frame.pack(side=tk.LEFT, fill=tk.Y)
        try:
            self._logo_img = tk.PhotoImage(file=str(get_base_path() / "logo.png"))
            factor = max(1, self._logo_img.height() // 20)
            self._logo_img = self._logo_img.subsample(factor)
            tk.Label(left_frame, image=self._logo_img, bg=_PANEL_WHITE).pack(side=tk.LEFT, padx=24, pady=22)
        except Exception:
            tk.Label(left_frame, text="Nagarkot", font=("Segoe UI",12,"bold"), fg=_PRIMARY_BLUE, bg=_PANEL_WHITE).pack(side=tk.LEFT, padx=24, pady=22)
            
        # Absolute Center Title
        tf = tk.Frame(header, bg=_PANEL_WHITE); tf.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        tk.Label(tf, text="Checklist Template Generator", font=("Segoe UI",14,"bold"), fg=_PRIMARY_BLUE, bg=_PANEL_WHITE).pack()
        tk.Label(tf, text="Nagarkot Forwarders Pvt Ltd", font=("Segoe UI",9), fg=_MUTED_GRAY, bg=_PANEL_WHITE).pack()

        # Right Actions
        right_frame = tk.Frame(header, bg=_PANEL_WHITE)
        right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=24)
        
        self.logout_btn = tk.Button(right_frame, text="Logout", command=self.show_login_screen, bg=_PANEL_WHITE, fg=_ACCENT_RED, bd=0, font=("Segoe UI", 10, "bold"), cursor="hand2")
        self.logout_btn.pack(side=tk.RIGHT, pady=18, padx=(10, 0))
        self.logout_btn.bind("<Enter>", lambda e: self.logout_btn.configure(bg="#FEE2E2"))
        self.logout_btn.bind("<Leave>", lambda e: self.logout_btn.configure(bg=_PANEL_WHITE))
        self.logout_btn.pack_forget()
        
        self.user_label = tk.Label(right_frame, text="", font=("Segoe UI", 10, "italic"), fg=_MUTED_GRAY, bg=_PANEL_WHITE)
        self.user_label.pack(side=tk.RIGHT, pady=18)
        self.user_label.pack_forget()

    def _build_footer(self) -> None:
        ft = tk.Frame(self.root, bg=_PANEL_WHITE, height=28, highlightthickness=1, highlightbackground=_BORDER_GRAY)
        ft.pack(fill=tk.X, side=tk.BOTTOM); ft.pack_propagate(False)
        tk.Label(ft, text="Nagarkot Forwarders Pvt. Ltd. ©", font=("Segoe UI",8), fg=_MUTED_GRAY, bg=_PANEL_WHITE).pack(side=tk.LEFT, padx=12)

    def _clear_body(self):
        for widget in self.body_container.winfo_children():
            widget.destroy()

    # --- SCREEN 1: LOGIN ---
    def show_login_screen(self):
        self.logout_btn.pack_forget()
        if hasattr(self, 'user_label'):
            self.user_label.pack_forget()
        self._clear_body()
        auth.clear_session()
        
        login_frame = tk.Frame(self.body_container, bg=_PANEL_WHITE, padx=40, pady=40, highlightthickness=1, highlightbackground=_BORDER_GRAY)
        login_frame.place(relx=0.5, rely=0.4, anchor=tk.CENTER)
        
        tk.Label(login_frame, text="System Login", font=("Segoe UI", 16, "bold"), fg=_PRIMARY_BLUE, bg=_PANEL_WHITE).pack(pady=(0, 20))
        
        # Credentials & URL Setup
        tk.Label(login_frame, text="GSheets Master URL:", bg=_PANEL_WHITE, font=("Segoe UI", 9)).pack(anchor='w')
        ttk.Entry(login_frame, textvariable=self.gsheets_url, width=50).pack(pady=(0, 10))
        
        tk.Label(login_frame, text="Credentials JSON:", bg=_PANEL_WHITE, font=("Segoe UI", 9)).pack(anchor='w')
        cred_entry = ttk.Entry(login_frame, textvariable=self.cred_path, width=50)
        cred_entry.pack(pady=(0, 10))
        
        tk.Label(login_frame, text="Username:", bg=_PANEL_WHITE, font=("Segoe UI", 9)).pack(anchor='w')
        username_var = tk.StringVar()
        ttk.Entry(login_frame, textvariable=username_var, width=50).pack(pady=(0, 10))
        
        tk.Label(login_frame, text="Password:", bg=_PANEL_WHITE, font=("Segoe UI", 9)).pack(anchor='w')
        password_var = tk.StringVar()
        ttk.Entry(login_frame, textvariable=password_var, width=50, show="*").pack(pady=(0, 20))
        
        def do_login():
            url = self.gsheets_url.get().strip()
            if not url:
                messagebox.showerror("Error", "GSheets URL is required.")
                return
            
            self._clear_body()
            loading_label = tk.Label(self.body_container, text="⌛ Logging in and fetching importer list...", 
                                     font=("Segoe UI", 12, "bold"), fg=_PRIMARY_BLUE, bg=_LIGHT_BG)
            loading_label.place(relx=0.5, rely=0.4, anchor=tk.CENTER)
            self.root.update()
            
            def auth_thread():
                try:
                    success, msg = auth.authenticate_user(url, self.cred_path.get(), username_var.get().strip(), password_var.get().strip())
                    if success:
                        from gsheets import GSheetsClient
                        client = GSheetsClient(self.cred_path.get())
                        sheet_id = client.extract_sheet_id(self.gsheets_url.get())
                        spreadsheet = client.client.open_by_key(sheet_id)
                        self.tab_names = [ws.title for ws in spreadsheet.worksheets() if ws.title not in ['Tool_Users', 'Audit_Log', 'Invoice_Header_Mappings', 'Pending_Model_Approval']]
                        
                        try:
                            self.invoice_mappings_df, _ = client.get_sheet_data(self.gsheets_url.get(), "Invoice_Header_Mappings")
                        except Exception as mapping_err:
                            print(f"Warning: Could not load Invoice_Header_Mappings. {mapping_err}")
                            self.invoice_mappings_df = None
                            
                        self.importer = None
                        self.importer_df = None
                        self.rows_data = []
                        self.root.after(0, self.show_builder_screen)
                    else:
                        self.root.after(0, lambda m=msg: handle_login_fail(m))
                except Exception as e:
                    self.root.after(0, lambda err=str(e): handle_login_fail(err))
                    
            def handle_login_fail(err_msg):
                messagebox.showerror("Login Failed", err_msg)
                self.show_login_screen()
                
            threading.Thread(target=auth_thread, daemon=True).start()
                
        _brand_button(login_frame, "Login", do_login).pack(fill=tk.X)



    # --- SCREEN 3: CHECKLIST BUILDER ---
    def show_builder_screen(self):
        self.logout_btn.pack(side=tk.RIGHT, pady=18, padx=(10, 0))
        if hasattr(self, 'user_label'):
            session = auth.get_current_session()
            if session and isinstance(session, dict):
                username = session.get('Username', 'User')
                self.user_label.config(text=f"👤 {username}")
                self.user_label.pack(side=tk.RIGHT, pady=18)
        self._clear_body()
        
        def _do_clear():
            if self.rows_data and messagebox.askyesno("Confirm Clear", "Are you sure you want to clear all rows?"):
                self.rows_data = []
                self.session_queue = []
                self._reload_table_rows()

        # Control Panel Wrapper
        ctrl_wrapper = tk.Frame(self.body_container, bg=_PANEL_WHITE, highlightthickness=1, highlightbackground=_BORDER_GRAY)
        ctrl_wrapper.pack(fill=tk.X, pady=(0, 16))
        
        # Left Side: Importer
        left_ctrl = tk.Frame(ctrl_wrapper, bg=_PANEL_WHITE, padx=16, pady=16)
        left_ctrl.pack(side=tk.LEFT, fill=tk.Y)
        
        tk.Label(left_ctrl, text="Importer:", font=("Segoe UI", 9, "bold"), fg=_MUTED_GRAY, bg=_PANEL_WHITE).pack(anchor='w', pady=(0,4))
        
        importer_combo = ttk.Combobox(left_ctrl, width=25)
        if hasattr(self, 'tab_names') and self.tab_names:
            importer_combo['values'] = self.tab_names
            if self.importer in self.tab_names:
                importer_combo.set(self.importer)
            else:
                importer_combo.set("Select Importer...")
        else:
            importer_combo.set("Select Importer...")
        importer_combo.pack(anchor='w')
        
        # Right Side: Actions
        right_ctrl = tk.Frame(ctrl_wrapper, bg=_PANEL_WHITE, padx=16, pady=16)
        right_ctrl.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        tk.Label(right_ctrl, text="Add Models (comma-separated):", font=("Segoe UI", 9, "bold"), fg=_MUTED_GRAY, bg=_PANEL_WHITE).pack(anchor='w', pady=(0,4))

        input_row = tk.Frame(right_ctrl, bg=_PANEL_WHITE)
        input_row.pack(fill=tk.X)

        models_var = tk.StringVar()
        models_entry = ttk.Entry(input_row, textvariable=models_var)
        models_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # We will pack the action buttons into input_row later below.
        
        status_row = tk.Frame(right_ctrl, bg=_PANEL_WHITE)
        status_row.pack(fill=tk.X, pady=(4,0))
        
        def filter_builder_tabs(event):
            if event.keysym in ("Up", "Down", "Left", "Right", "Return", "Escape"):
                return
            typed = importer_combo.get().strip()
            if not typed:
                importer_combo['values'] = self.tab_names
            else:
                filtered = [t for t in self.tab_names if typed.lower() in t.lower()]
                importer_combo['values'] = filtered
        importer_combo.bind('<KeyRelease>', filter_builder_tabs)
        
        def on_importer_changed(event):
            new_imp = importer_combo.get().strip()
            if not new_imp:
                return
            if hasattr(self, 'tab_names') and self.tab_names:
                if new_imp not in self.tab_names:
                    matches = [t for t in self.tab_names if new_imp.lower() in t.lower()]
                    if len(matches) == 1:
                        new_imp = matches[0]
                    else:
                        messagebox.showerror("Error", f"Invalid importer name: '{new_imp}'")
                        importer_combo.set(self.importer)
                        return
            if new_imp and new_imp != self.importer:
                if self.rows_data:
                    if not messagebox.askyesno("Confirm", "Changing importer will clear your current checklist rows. Proceed?"):
                        importer_combo.set(self.importer)
                        return
                
                self.importer = new_imp
                self._clear_body()
                loading_imp = tk.Label(self.body_container, text=f"⌛ Loading master database for '{new_imp}'...", 
                                       font=("Segoe UI", 12, "bold"), fg=_PRIMARY_BLUE, bg=_LIGHT_BG)
                loading_imp.place(relx=0.5, rely=0.4, anchor=tk.CENTER)
                self.root.update()
                
                def load_df_thread():
                    try:
                        from gsheets import GSheetsClient
                        client = GSheetsClient(self.cred_path.get())
                        df, _ = client.get_sheet_data(self.gsheets_url.get(), new_imp)
                        if not df.empty:
                            for col in df.columns:
                                if str(col).strip().lower() == 'model':
                                    df.rename(columns={col: 'Model'}, inplace=True)
                                    break
                            if 'Model' in df.columns:
                                df['Model'] = df['Model'].astype(str).str.strip()
                        self.importer_df = df
                        self.rows_data = [] # Reset rows
                        self.session_queue = []
                        self.root.after(0, self.show_builder_screen)
                    except Exception as e:
                        def handle_error(err):
                            messagebox.showerror("Error", f"Failed to load data:\n{err}")
                            self.importer = None
                            self.importer_df = None
                            self.show_builder_screen()
                        self.root.after(0, lambda: handle_error(str(e)))
                        
                threading.Thread(target=load_df_thread, daemon=True).start()
                
        importer_combo.bind("<<ComboboxSelected>>", on_importer_changed)
        importer_combo.bind("<Return>", on_importer_changed)
        
        def do_add_models():
            if not self.importer or self.importer_df is None:
                messagebox.showerror("Error", "Please select an importer first.")
                return
            input_str = models_var.get().strip()
            if not input_str:
                messagebox.showerror("Error", "Please enter at least one model.")
                return
                
            models_to_lookup = [m.strip() for m in input_str.split(',') if m.strip()]
            
            if not hasattr(self, 'session_queue'):
                self.session_queue = []
            self.session_queue.extend(models_to_lookup)
            
            self.root.config(cursor="watch")
            self._process_model_queue(models_to_lookup)
            models_var.set("") # Clear entry
            
        self.status_label = tk.Label(status_row, text="", font=("Segoe UI", 9, "italic"), bg=_PANEL_WHITE, fg=_MUTED_GRAY)
        
        def _handle_upload_err(err):
            self.root.config(cursor="")
            self.status_label.config(text="")
            messagebox.showerror("Error", f"Failed to read Excel file: {err}")
            
        def _do_upload_invoice():
            if not self.importer or self.importer_df is None:
                messagebox.showerror("Error", "Please select an importer first.")
                return
            if getattr(self, 'invoice_mappings_df', None) is None:
                messagebox.showerror("Error", "Invoice mapping configuration could not be loaded. Try relogging.")
                return
                
            file_path = filedialog.askopenfilename(
                title="Select Invoice",
                filetypes=[("Excel & PDF Files", "*.xlsx *.xls *.pdf")]
            )
            if not file_path:
                return
                
            self.status_label.config(text="⏳ Reading file...")
            self.root.config(cursor="watch")
            self.root.update_idletasks()
            
            def parse_thread():
                try:
                    import pandas as pd
                    if self.importer.lower() == 'advics':
                        invoice_df = pd.read_excel(file_path, header=None)
                        self.root.after(0, lambda: _parse_advics_invoice(invoice_df))
                    elif self.importer.lower() == 'arjo':
                        invoice_df = pd.read_excel(file_path, header=None)
                        self.root.after(0, lambda: _parse_arjo_invoice(invoice_df))
                    elif self.importer.lower() == 'ansell':
                        if not file_path.lower().endswith('.pdf'):
                            raise ValueError("Ansell importer requires a PDF invoice.")
                        self.root.after(0, lambda: _parse_ansell_invoice(file_path))
                    elif self.importer.lower() == 'aviat':
                        if file_path.lower().endswith('.pdf'):
                            self.root.after(0, lambda: _parse_aviat_invoice(file_path))
                        else:
                            invoice_df = pd.read_excel(file_path)
                            self.root.after(0, lambda: _parse_aviat_excel(invoice_df))
                    else:
                        invoice_df = pd.read_excel(file_path)
                        self.root.after(0, lambda: _on_file_parsed(invoice_df))
                except Exception as e:
                    self.root.after(0, lambda err=e: _handle_upload_err(err))
                    
            threading.Thread(target=parse_thread, daemon=True).start()
            
        def _parse_aviat_excel(df):
            self.root.config(cursor="")
            self.status_label.config(text="")
            
            # Check for required columns based on user instruction
            req_cols = ["Commercial Invoice No", "Document Date (Long)", "Part No", "Quantity", "Unit Price"]
            missing_cols = [c for c in req_cols if c not in df.columns]
            if missing_cols:
                messagebox.showerror("Error", f"Missing required columns in Excel: {', '.join(missing_cols)}")
                return
                
            if "CoO" not in df.columns:
                messagebox.showerror("Missing Column", "Country of origin is blank.\n\nPlease add a 'CoO' column to the excel file and enter the CoO.")
                return
                
            if "Currency" not in df.columns:
                messagebox.showerror("Missing Column", "Currency column is missing.\n\nPlease add a 'Currency' column to the excel file and add details in it.")
                return
                
            target_to_col_idx = {}
            mapping = {
                "Inv No": "col_1", "Date": "col_2",
                "Quantity": "col_5", "Unit": "col_6", "Currency": "col_7",
                "Rate": "col_8", "Country of origin": "col_32"
            }
            # Add missing mappings dynamically
            import pandas as pd
            for t_key in mapping.keys():
                for i, h in enumerate(checklist_logic.LOGISYS_HEADERS):
                    if str(h).strip().lower() == t_key.lower():
                        target_to_col_idx[t_key] = f"col_{i+1}"
                        break
                        
            queue_items = []
            for _, row in df.iterrows():
                model_val = str(row.get("Part No", "")).strip()
                if not model_val or pd.isna(row.get("Part No")) or model_val.lower() == 'nan':
                    continue
                    
                inv_no = str(row.get("Commercial Invoice No", "")).strip()
                if pd.isna(row.get("Commercial Invoice No")) or inv_no.lower() == 'nan':
                    inv_no = ""
                    
                date_val = row.get("Document Date (Long)", "")
                if hasattr(date_val, 'strftime'):
                    date_val = date_val.strftime('%d-%m-%Y')
                else:
                    date_val = str(date_val).strip()
                    if date_val.lower() == 'nan': date_val = ""
                    
                qty = str(row.get("Quantity", "")).strip()
                if qty.lower() == 'nan': qty = ""
                
                rate = str(row.get("Unit Price", "")).strip()
                if rate.lower() == 'nan': rate = ""
                
                coo = str(row.get("CoO", "")).strip()
                if pd.isna(row.get("CoO")) or coo.lower() == 'nan' or not coo:
                    messagebox.showerror("Missing CoO", f"Country of origin is blank for part {model_val}.\n\nPlease add a 'CoO' column to the excel file and enter the CoO.")
                    return
                    
                currency = str(row.get("Currency", "")).strip()
                if pd.isna(row.get("Currency")) or currency.lower() == 'nan' or not currency:
                    messagebox.showerror("Missing Currency", f"Currency is blank for part {model_val}.\n\nPlease add details in the Currency column.")
                    return
                
                prefilled = {
                    target_to_col_idx.get("Inv No"): inv_no,
                    target_to_col_idx.get("Date"): date_val,
                    target_to_col_idx.get("Quantity"): qty,
                    target_to_col_idx.get("Rate"): rate,
                    target_to_col_idx.get("Country of origin"): coo,
                    target_to_col_idx.get("Unit"): "NOS",
                    target_to_col_idx.get("Currency"): currency
                }
                prefilled = {k: v for k, v in prefilled.items() if k is not None}
                queue_items.append({'model': model_val, 'prefilled': prefilled})
                
            if queue_items:
                if not hasattr(self, 'session_queue'):
                    self.session_queue = []
                self.session_queue.extend(queue_items)
                
                self.status_label.config(text=f"⚙️ Matching {len(queue_items)} models...")
                self.root.update_idletasks()
                self._process_model_queue(queue_items)
            else:
                messagebox.showinfo("Info", "No valid rows found in the Aviat Excel.")
            
        def _parse_aviat_invoice(pdf_path):
            try:
                import sys, os
                sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
                from aviat_inv_extractor import parse_commercial_invoice, format_date_helper, trim_hs_code
            except ImportError as e:
                self.root.config(cursor="")
                self.status_label.config(text="")
                messagebox.showerror("Error", f"Failed to load Aviat extractor: {e}")
                return
                
            try:
                inv_no, date_raw, ci_items = parse_commercial_invoice(pdf_path)
            except Exception as e:
                self.root.config(cursor="")
                self.status_label.config(text="")
                messagebox.showerror("Error", f"Failed to parse Aviat PDF: {e}")
                return
                
            if not ci_items:
                self.root.config(cursor="")
                self.status_label.config(text="")
                messagebox.showinfo("Info", "No valid rows found in the Aviat invoice.")
                return
                
            # Currency detection: look below the word 'CURRENCY'
            currency_val = ""
            try:
                import pdfplumber
                with pdfplumber.open(pdf_path) as pdf:
                    first_page_text = pdf.pages[0].extract_text() or ""
                    lines = first_page_text.split('\n')
                    for i, line in enumerate(lines):
                        if "CURRENCY" in line.upper():
                            if i + 1 < len(lines):
                                next_line = lines[i+1].strip()
                                parts = next_line.split()
                                if parts:
                                    currency_val = parts[-1] # The currency code is usually the last word
                            break
            except Exception:
                pass
                
            date_short, date_long = format_date_helper(date_raw)
            
            target_to_col_idx = {}
            mapping = {
                "Inv No": "col_1", "Date": "col_2", "Product Desc": "col_4", 
                "Quantity": "col_5", "Unit": "col_6", "Currency": "col_7", 
                "Rate": "col_8", "Country of origin": "col_32", "CT (HS Code)": "col_42"
            }
            # Add missing mappings dynamically
            for t_key in mapping.keys():
                for i, h in enumerate(checklist_logic.LOGISYS_HEADERS):
                    if str(h).strip().lower() == t_key.lower():
                        target_to_col_idx[t_key] = f"col_{i+1}"
                        break
                        
            queue_items = []
            for item in ci_items:
                model_val = item.get("part_no", "")
                if not model_val or model_val == "N/A":
                    continue
                    
                desc_upper = (item.get("description") or "").upper()
                ct_trimmed = trim_hs_code(item.get("ct"))
                
                prefilled = {
                    target_to_col_idx.get("Inv No"): inv_no,
                    target_to_col_idx.get("Date"): date_short,
                    target_to_col_idx.get("Product Desc"): desc_upper,
                    target_to_col_idx.get("Quantity"): str(item.get("quantity", "")),
                    target_to_col_idx.get("Unit"): "NOS",
                    target_to_col_idx.get("Currency"): currency_val,
                    target_to_col_idx.get("Rate"): str(item.get("unit_price", "")),
                    target_to_col_idx.get("Country of origin"): item.get("coo", ""),
                    target_to_col_idx.get("CT (HS Code)"): ct_trimmed,
                }
                prefilled = {k: v for k, v in prefilled.items() if k is not None}
                queue_items.append({'model': model_val, 'prefilled': prefilled})
                
            if queue_items:
                if not hasattr(self, 'session_queue'):
                    self.session_queue = []
                self.session_queue.extend(queue_items)
                
                self.status_label.config(text=f"⚙️ Matching {len(queue_items)} models...")
                self.root.update_idletasks()
                self._process_model_queue(queue_items)
            else:
                self.root.config(cursor="")
                self.status_label.config(text="")
                messagebox.showinfo("Info", "No valid models found in the Aviat invoice.")
            
        def _parse_ansell_invoice(pdf_path):
            try:
                import pdfplumber
            except ImportError:
                self.root.config(cursor="")
                self.status_label.config(text="")
                messagebox.showerror("Error", "pdfplumber is not installed.")
                return
                
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    all_tables = []
                    for page in pdf.pages:
                        all_tables.extend(page.extract_tables())
            except Exception as e:
                self.root.config(cursor="")
                self.status_label.config(text="")
                messagebox.showerror("Error", f"Failed to read PDF: {e}")
                return

            inv_table = None
            for tbl in all_tables:
                if len(tbl) >= 3 and tbl[0] and tbl[0][0] and "Commercial Invoice" in tbl[0][0]:
                    inv_table = tbl
                    break

            if inv_table is None:
                self.root.config(cursor="")
                self.status_label.config(text="")
                messagebox.showerror("Error", f"Invoice header table not found in {pdf_path}")
                return

            invoice_no = inv_table[2][0].strip() if inv_table[2][0] else ""
            invoice_date = ""
            if len(inv_table[2]) > 1 and inv_table[2][1]:
                invoice_date = inv_table[2][1].strip()

            item_tables = []
            for tbl in all_tables:
                if not tbl or not tbl[0]: continue
                if any(cell and "Product Description" in cell for cell in tbl[0]):
                    item_tables.append(tbl)

            if not item_tables:
                self.root.config(cursor="")
                self.status_label.config(text="")
                messagebox.showerror("Error", f"No item tables found in {pdf_path}")
                return

            def find_col(keyword, rows):
                kw = keyword.lower()
                for row in rows:
                    if not row: continue
                    for idx, cell in enumerate(row):
                        if cell and kw in cell.lower(): return idx
                return None

            records = []
            for tbl in item_tables:
                hdr = tbl[0]
                sub = tbl[1] if len(tbl) > 1 else [None] * len(hdr)

                idx_prod_code = find_col("product", [hdr])
                if idx_prod_code is not None and hdr[idx_prod_code] and "Description" in hdr[idx_prod_code]:
                    for idx, cell in enumerate(hdr):
                        if cell and "Product" in cell and "Code" in cell:
                            idx_prod_code = idx
                            break

                idx_desc = find_col("Product Description", [hdr])
                idx_country = find_col("Country", [hdr])
                
                idx_hs = find_col("HTS/HS", [hdr])
                idx_shipcase = find_col("Case", [hdr])
                idx_value = find_col("Value", [hdr])

                idx_qty = None
                idx_uom = None
                idx_price = None
                for idx, cell in enumerate(sub):
                    if cell == "Qty" and idx_qty is None: idx_qty = idx
                    elif cell == "UOM" and idx_qty is not None and idx_uom is None: idx_uom = idx
                    elif cell == "Price" and idx_price is None: idx_price = idx

                idx_currency = find_col("Cur", [hdr])

                def get(row, idx):
                    if idx is None or idx >= len(row): return ""
                    return (row[idx] or "").strip()

                for row in tbl[2:]:
                    if not row: continue
                    first_cell = (row[0] or "").strip() if len(row) > 0 else ""
                    if first_cell.startswith("Total"): break
                    if first_cell == "": continue

                    record = {
                        "Invoice_No": invoice_no,
                        "Invoice_Date": invoice_date,
                        "Product_Code": get(row, idx_prod_code),
                        "Product_Description": get(row, idx_desc),
                        "Country_of_Origin": get(row, idx_country),
                        "Qty": get(row, idx_qty),
                        "UOM": get(row, idx_uom),
                        "Price": get(row, idx_price),
                        "Currency": get(row, idx_currency),
                        "HS_Code": get(row, idx_hs),
                        "Ship_Case": get(row, idx_shipcase),
                        "Value": get(row, idx_value),
                    }
                    records.append(record)
                    
            target_to_col_idx = {}
            mapping = {
                "Inv No": "col_1", "Date": "col_2", "Product Desc": "col_4", 
                "Quantity": "col_5", "Unit": "col_6", "Currency": "col_7", 
                "Rate": "col_8", "Country of origin": "col_32"
            }
            for t_key in mapping.keys():
                for i, h in enumerate(checklist_logic.LOGISYS_HEADERS):
                    if str(h).strip().lower() == t_key.lower():
                        target_to_col_idx[t_key] = f"col_{i+1}"
                        break
                        
            queue_items = []
            for rec in records:
                model_val = rec.get("Product_Code", "")
                if not model_val:
                    continue
                    
                # Rate calculation: Value / Qty
                ansell_rate = ""
                qty_val = rec.get("Qty", "")
                val_val = rec.get("Value", "")
                if qty_val and val_val:
                    try:
                        q_num = float(qty_val.replace(',', ''))
                        v_num = float(val_val.replace(',', ''))
                        if q_num > 0:
                            ansell_rate = str(round(v_num / q_num, 4))
                    except ValueError:
                        pass
                        
                prefilled = {
                    target_to_col_idx.get("Inv No"): rec.get("Invoice_No", ""),
                    target_to_col_idx.get("Date"): rec.get("Invoice_Date", ""),
                    target_to_col_idx.get("Quantity"): rec.get("Qty", ""),
                    target_to_col_idx.get("Unit"): rec.get("UOM", ""),
                    target_to_col_idx.get("Currency"): rec.get("Currency", ""),
                    target_to_col_idx.get("Rate"): ansell_rate,
                    target_to_col_idx.get("Country of origin"): rec.get("Country_of_Origin", ""),
                    
                    # Internal keys for Ansell description construction
                    "_ansell_raw_desc": rec.get("Product_Description", ""),
                    "_ansell_hs": rec.get("HS_Code", ""),
                    "_ansell_case": rec.get("Ship_Case", ""),
                    "_ansell_unit_price": rec.get("Price", "")
                }
                prefilled = {k: v for k, v in prefilled.items() if k is not None}
                queue_items.append({'model': model_val, 'prefilled': prefilled})

            if queue_items:
                if not hasattr(self, 'session_queue'):
                    self.session_queue = []
                self.session_queue.extend(queue_items)
                
                self.status_label.config(text=f"⚙️ Matching {len(queue_items)} models...")
                self.root.update_idletasks()
                self._process_model_queue(queue_items)
            else:
                self.root.config(cursor="")
                self.status_label.config(text="")
                messagebox.showinfo("Info", "No valid rows found in the Ansell invoice.")
                
        def _parse_advics_invoice(invoice_df):
            import pandas as pd
            inv_no, inv_date, coo_val = "", "", ""
            
            for r in range(min(30, len(invoice_df))):
                for c in range(len(invoice_df.columns)):
                    val = str(invoice_df.iat[r, c]).strip()
                    if val == "INVOICE NO. :":
                        for c_next in range(c + 1, len(invoice_df.columns)):
                            if pd.notna(invoice_df.iat[r, c_next]) and str(invoice_df.iat[r, c_next]).strip():
                                inv_no = str(invoice_df.iat[r, c_next]).strip()
                                break
                    elif val == "DATE :":
                        for c_next in range(c + 1, len(invoice_df.columns)):
                            if pd.notna(invoice_df.iat[r, c_next]) and str(invoice_df.iat[r, c_next]).strip():
                                d_val = invoice_df.iat[r, c_next]
                                if hasattr(d_val, 'strftime'):
                                    inv_date = d_val.strftime('%d-%m-%Y')
                                else:
                                    inv_date = str(d_val).strip()
                                break
                    elif val == "Country of origin :":
                        for c_next in range(c + 1, len(invoice_df.columns)):
                            if pd.notna(invoice_df.iat[r, c_next]) and str(invoice_df.iat[r, c_next]).strip():
                                coo_val = str(invoice_df.iat[r, c_next]).strip()
                                break

            table_start_row = -1
            model_col, desc_col, qty_col, price_col = -1, -1, -1, -1
            currency = "THB"
            
            for r in range(len(invoice_df)):
                row_vals = [str(x).strip() for x in invoice_df.iloc[r] if pd.notna(x)]
                if "PART NO." in row_vals and "DESCRIPTION" in row_vals:
                    table_start_row = r
                    for c in range(len(invoice_df.columns)):
                        val = str(invoice_df.iat[r, c]).strip()
                        if val == "PART NO.": model_col = c
                        elif val == "DESCRIPTION": desc_col = c
                        elif val == "Qty": qty_col = c
                        elif "Unit Price" in val:
                            price_col = c
                            if "(THB)" in val: currency = "THB"
                    break
                    
            if table_start_row == -1 or model_col == -1:
                self.root.config(cursor="")
                self.status_label.config(text="")
                messagebox.showerror("Error", "Could not find 'PART NO.' and 'DESCRIPTION' in Advics invoice.")
                return

            target_to_col_idx = {}
            mapping = {
                "Inv No": "col_1", "Date": "col_2", "Product Desc": "col_4", 
                "Quantity": "col_5", "Unit": "col_6", "Currency": "col_7", 
                "Rate": "col_8", "Country of origin": "col_32"
            }
            for t_key, expected_col in mapping.items():
                for i, h in enumerate(checklist_logic.LOGISYS_HEADERS):
                    if str(h).strip().lower() == t_key.lower():
                        target_to_col_idx[t_key] = f"col_{i+1}"
                        break
            
            queue_items = []
            for r in range(table_start_row + 1, len(invoice_df)):
                if pd.isna(invoice_df.iat[r, model_col]):
                    continue
                    
                model_val = str(invoice_df.iat[r, model_col]).strip()
                if not model_val or model_val.lower() == 'nan':
                    continue
                    
                desc_val = str(invoice_df.iat[r, desc_col]).strip() if desc_col != -1 and pd.notna(invoice_df.iat[r, desc_col]) else ""
                qty_val = str(invoice_df.iat[r, qty_col]).strip() if qty_col != -1 and pd.notna(invoice_df.iat[r, qty_col]) else ""
                price_val = str(invoice_df.iat[r, price_col]).strip() if price_col != -1 and pd.notna(invoice_df.iat[r, price_col]) else ""
                
                unit_val = ""
                if qty_col != -1:
                    for c_next in range(qty_col + 1, len(invoice_df.columns)):
                        c_val = str(invoice_df.iat[r, c_next]).strip()
                        if c_val and c_val.lower() != 'nan':
                            unit_val = c_val
                            break
                            
                if unit_val.lower() in ("pcs.", "pcs"):
                    unit_val = "NOS"
                
                prefilled = {
                    target_to_col_idx.get("Inv No"): inv_no,
                    target_to_col_idx.get("Date"): inv_date,
                    target_to_col_idx.get("Product Desc"): desc_val,
                    target_to_col_idx.get("Quantity"): qty_val,
                    target_to_col_idx.get("Unit"): unit_val,
                    target_to_col_idx.get("Currency"): currency,
                    target_to_col_idx.get("Rate"): price_val,
                    target_to_col_idx.get("Country of origin"): coo_val
                }
                prefilled = {k: v for k, v in prefilled.items() if k is not None}
                queue_items.append({'model': model_val, 'prefilled': prefilled})

            if queue_items:
                if not hasattr(self, 'session_queue'):
                    self.session_queue = []
                self.session_queue.extend(queue_items)
                
                self.status_label.config(text=f"⚙️ Matching {len(queue_items)} models...")
                self.root.update_idletasks()
                self._process_model_queue(queue_items)
            else:
                self.root.config(cursor="")
                self.status_label.config(text="")
                messagebox.showinfo("Info", "No valid rows found in the Advics invoice.")
                
        def _parse_arjo_invoice(invoice_df):
            import pandas as pd
            import re
            inv_no, inv_date = "", ""
            
            for r in range(min(50, len(invoice_df))):
                for c in range(len(invoice_df.columns)):
                    val = str(invoice_df.iat[r, c]).strip()
                    if "Invoice No." in val:
                        match = re.search(r'Invoice No\.\s*(\S+)', val, re.IGNORECASE)
                        if match:
                            inv_no = match.group(1)
                        else:
                            for c_next in range(c + 1, len(invoice_df.columns)):
                                if pd.notna(invoice_df.iat[r, c_next]) and str(invoice_df.iat[r, c_next]).strip():
                                    inv_no = str(invoice_df.iat[r, c_next]).strip()
                                    break
                    elif "Invoice date :" in val:
                        match = re.search(r'Invoice date\s*:\s*(.*)', val, re.IGNORECASE)
                        d_val = ""
                        if match and match.group(1).strip():
                            d_val = match.group(1).strip()
                        else:
                            for c_next in range(c + 1, len(invoice_df.columns)):
                                if pd.notna(invoice_df.iat[r, c_next]) and str(invoice_df.iat[r, c_next]).strip():
                                    d_val = invoice_df.iat[r, c_next]
                                    break
                        if d_val:
                            if hasattr(d_val, 'strftime'):
                                inv_date = d_val.strftime('%d-%m-%Y')
                            else:
                                try:
                                    parsed_date = pd.to_datetime(str(d_val).strip())
                                    inv_date = parsed_date.strftime('%d-%m-%Y')
                                except:
                                    inv_date = str(d_val).strip()

            table_start_row = -1
            model_col, desc_col, qty_col, price_col, coo_col = -1, -1, -1, -1, -1
            currency = ""
            
            for r in range(len(invoice_df)):
                row_vals = [str(x).strip() for x in invoice_df.iloc[r] if pd.notna(x)]
                if "Part No" in row_vals and "Qty" in row_vals:
                    table_start_row = r
                    for c in range(len(invoice_df.columns)):
                        val = str(invoice_df.iat[r, c]).strip()
                        if val == "Part No": model_col = c
                        elif "Product Description" in val or "Product Desc" in val: desc_col = c
                        elif val == "Qty": qty_col = c
                        elif val == "Country Of Origin": coo_col = c
                        elif "Price" in val:
                            price_col = c
                            match = re.search(r'\((.*?)\)', val)
                            if match:
                                currency = match.group(1).strip()
                    break
                    
            if table_start_row == -1 or model_col == -1:
                self.root.config(cursor="")
                self.status_label.config(text="")
                messagebox.showerror("Error", "Could not find 'Part No' and 'Qty' in Arjo invoice.")
                return

            target_to_col_idx = {}
            mapping = {
                "Inv No": "col_1", "Date": "col_2", "Product Desc": "col_4",
                "Quantity": "col_5", "Unit": "col_6", "Currency": "col_7", 
                "Rate": "col_8", "Country of origin": "col_32"
            }
            for t_key, expected_col in mapping.items():
                for i, h in enumerate(checklist_logic.LOGISYS_HEADERS):
                    if str(h).strip().lower() == t_key.lower():
                        target_to_col_idx[t_key] = f"col_{i+1}"
                        break
            
            queue_items = []
            for r in range(table_start_row + 1, len(invoice_df)):
                if pd.isna(invoice_df.iat[r, model_col]):
                    continue
                    
                model_val = str(invoice_df.iat[r, model_col]).strip()
                if not model_val or model_val.lower() == 'nan':
                    continue
                    
                qty_val = str(invoice_df.iat[r, qty_col]).strip() if qty_col != -1 and pd.notna(invoice_df.iat[r, qty_col]) else ""
                desc_val = str(invoice_df.iat[r, desc_col]).strip() if desc_col != -1 and pd.notna(invoice_df.iat[r, desc_col]) else ""
                price_val = str(invoice_df.iat[r, price_col]).strip() if price_col != -1 and pd.notna(invoice_df.iat[r, price_col]) else ""
                coo_val = str(invoice_df.iat[r, coo_col]).strip() if coo_col != -1 and pd.notna(invoice_df.iat[r, coo_col]) else ""
                unit_val = "NOS"
                
                prefilled = {
                    target_to_col_idx.get("Inv No"): inv_no,
                    target_to_col_idx.get("Date"): inv_date,
                    target_to_col_idx.get("Product Desc"): desc_val,
                    target_to_col_idx.get("Quantity"): qty_val,
                    target_to_col_idx.get("Unit"): unit_val,
                    target_to_col_idx.get("Currency"): currency,
                    target_to_col_idx.get("Rate"): price_val,
                    target_to_col_idx.get("Country of origin"): coo_val
                }
                prefilled = {k: v for k, v in prefilled.items() if k is not None}
                queue_items.append({'model': model_val, 'prefilled': prefilled})
                
            if queue_items:
                if not hasattr(self, 'session_queue'):
                    self.session_queue = []
                self.session_queue.extend(queue_items)
                
                self.status_label.config(text=f"⚙️ Matching {len(queue_items)} models...")
                self.root.update_idletasks()
                self._process_model_queue(queue_items)
            else:
                self.root.config(cursor="")
                self.status_label.config(text="")
                messagebox.showinfo("Info", "No valid rows found in the Arjo invoice.")
            
        def _on_file_parsed(invoice_df):
            import pandas as pd
            importer_map = self.invoice_mappings_df[self.invoice_mappings_df['Importer'] == self.importer]
            
            variants = {}
            for _, row in importer_map.iterrows():
                fmt = str(row.get('Format_Name', '')).strip()
                if not fmt: continue
                sh = str(row.get('Source_Header', '')).strip()
                tf = str(row.get('Target_Field', '')).strip()
                if pd.notna(sh):
                    if fmt not in variants:
                        variants[fmt] = {}
                    variants[fmt][sh] = tf
                    
            excel_headers = [str(c).strip() for c in invoice_df.columns]
            excel_header_set = set(excel_headers)
            
            matched_variants = []
            for fmt, mapping in variants.items():
                if set(mapping.keys()) == excel_header_set:
                    matched_variants.append(fmt)
            
            # Superset match: if no exact match, check if file contains all
            # of a variant's headers (plus extras). This handles real invoices
            # that have extra columns we don't need.
            if not matched_variants:
                superset_matches = []
                for fmt, mapping in variants.items():
                    if set(mapping.keys()).issubset(excel_header_set):
                        superset_matches.append(fmt)
                if len(superset_matches) == 1:
                    matched_variants = superset_matches
                elif len(superset_matches) > 1:
                    # Multiple variants are subsets — pick the one with the
                    # most headers (most specific match) to break the tie
                    superset_matches.sort(key=lambda f: len(variants[f]), reverse=True)
                    if len(variants[superset_matches[0]]) > len(variants[superset_matches[1]]):
                        matched_variants = [superset_matches[0]]
                    # else: genuine ambiguity, fall through to guidance
                    
            expected_headers = {}
            self.active_variant_targets = set()
            
            if len(matched_variants) == 1:
                fmt = matched_variants[0]
                self.active_variant_targets = set(variants[fmt].values())
                for sh, tf in variants[fmt].items():
                    if pd.notna(tf) and tf:
                        expected_headers[sh] = tf
                        

            elif len(matched_variants) > 1:
                self.root.config(cursor="")
                self.status_label.config(text="")
                messagebox.showerror("Error", f"Ambiguous format match! Matched variants: {', '.join(matched_variants)}.")
                return
                
            else:
                self.root.config(cursor="")
                self.status_label.config(text="")
                
                # Compute overlap scores to find the closest variant
                best_score = -1
                best_variants = []
                
                for fmt_name, fmt_mapping in variants.items():
                    variant_headers = set(fmt_mapping.keys())
                    overlap = len(variant_headers.intersection(excel_header_set))
                    if overlap > best_score:
                        best_score = overlap
                        best_variants = [(fmt_name, fmt_mapping)]
                    elif overlap == best_score:
                        best_variants.append((fmt_name, fmt_mapping))
                
                # Build the guidance dialog
                top = tk.Toplevel(self.root)
                top.title("Column Mismatch")
                top.grab_set()
                top.configure(bg="white")
                top.resizable(True, True)
                
                # Header bar
                header_frame = tk.Frame(top, bg="#FEF2F2", pady=8, padx=14)
                header_frame.pack(fill=tk.X)
                tk.Label(header_frame, text="⚠  Some column headers need renaming",
                         font=("Segoe UI", 11, "bold"), bg="#FEF2F2", fg="#991B1B").pack(anchor="w")
                tk.Label(header_frame, text="Rename these columns in your Excel file, save, then re-upload.",
                         font=("Segoe UI", 9), bg="#FEF2F2", fg="#7F1D1D").pack(anchor="w")
                
                # Scrollable content
                outer = tk.Frame(top, bg="white")
                outer.pack(fill=tk.BOTH, expand=True)
                canvas = tk.Canvas(outer, bg="white", highlightthickness=0)
                scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
                content = tk.Frame(canvas, bg="white", padx=16, pady=10)
                content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
                canvas.create_window((0, 0), window=content, anchor="nw")
                canvas.configure(yscrollcommand=scrollbar.set)
                canvas.pack(side="left", fill="both", expand=True)
                scrollbar.pack(side="right", fill="y")
                
                def _build_single_variant_guidance(parent, fmt_mapping):
                    """Build guidance showing what's wrong and what to do, returns full header list."""
                    variant_headers = set(fmt_mapping.keys())
                    correct = sorted(variant_headers.intersection(excel_header_set))
                    missing = sorted(variant_headers - excel_header_set)
                    extra = sorted(excel_header_set - variant_headers)
                    
                    if missing:
                        # Section 1: What needs fixing
                        fix_frame = tk.Frame(parent, bg="#FEF2F2", padx=10, pady=8)
                        fix_frame.pack(fill=tk.X, pady=(0, 8))
                        
                        tk.Label(fix_frame, text="❌  Columns that need renaming:",
                                 font=("Segoe UI", 10, "bold"), bg="#FEF2F2", fg="#991B1B",
                                 anchor="w").pack(fill=tk.X, pady=(0, 6))
                        
                        # Table header
                        hdr_row = tk.Frame(fix_frame, bg="#FEF2F2")
                        hdr_row.pack(fill=tk.X, pady=(0, 3))
                        tk.Label(hdr_row, text="Your column", font=("Segoe UI", 8, "bold"),
                                 bg="#FEF2F2", fg="#6B7280", width=24, anchor="w").pack(side=tk.LEFT)
                        tk.Label(hdr_row, text="", bg="#FEF2F2", width=4).pack(side=tk.LEFT)
                        tk.Label(hdr_row, text="Rename to", font=("Segoe UI", 8, "bold"),
                                 bg="#FEF2F2", fg="#6B7280", width=24, anchor="w").pack(side=tk.LEFT)
                        
                        # Pair missing headers with extra (unrecognized) columns
                        for i, required_name in enumerate(missing):
                            pair_row = tk.Frame(fix_frame, bg="#FEF2F2")
                            pair_row.pack(fill=tk.X, pady=2)
                            
                            # Left side: user's likely wrong column (if one exists)
                            if i < len(extra):
                                user_col = extra[i]
                                lbl = tk.Entry(pair_row, font=("Consolas", 9), fg="#DC2626",
                                               bg="#FECACA", readonlybackground="#FECACA",
                                               relief="flat", bd=0, width=26)
                                lbl.insert(0, user_col)
                                lbl.config(state="readonly")
                                lbl.pack(side=tk.LEFT)
                            else:
                                tk.Label(pair_row, text="(column missing)", font=("Segoe UI", 9, "italic"),
                                         bg="#FEF2F2", fg="#9CA3AF", width=24, anchor="w").pack(side=tk.LEFT)
                            
                            # Arrow
                            tk.Label(pair_row, text="  →  ", bg="#FEF2F2", fg="#6B7280",
                                     font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT)
                            
                            # Right side: what it should be renamed to (copyable)
                            e = tk.Entry(pair_row, font=("Consolas", 10, "bold"), fg="#1E40AF",
                                         bg="#DBEAFE", readonlybackground="#DBEAFE",
                                         relief="flat", bd=0, width=26)
                            e.insert(0, required_name)
                            e.config(state="readonly")
                            e.pack(side=tk.LEFT, padx=(0, 4))
                        
                        # Show any remaining extra columns that don't pair up
                        remaining_extra = extra[len(missing):]
                        if remaining_extra:
                            tk.Label(fix_frame, text=f"\nExtra columns in your file (not used): {', '.join(remaining_extra)}",
                                     font=("Segoe UI", 8), bg="#FEF2F2", fg="#9CA3AF",
                                     wraplength=500, justify="left", anchor="w").pack(fill=tk.X, pady=(4, 0))
                    
                    if correct:
                        # Section 2: What's already fine
                        ok_frame = tk.Frame(parent, bg="#F0FDF4", padx=10, pady=8)
                        ok_frame.pack(fill=tk.X, pady=(0, 4))
                        tk.Label(ok_frame, text="✓  Already correct:",
                                 font=("Segoe UI", 9, "bold"), bg="#F0FDF4", fg="#166534",
                                 anchor="w").pack(fill=tk.X)
                        tk.Label(ok_frame, text=", ".join(correct), font=("Segoe UI", 9),
                                 bg="#F0FDF4", fg="#15803D", wraplength=500, justify="left",
                                 anchor="w").pack(fill=tk.X)
                    
                    return list(fmt_mapping.keys())
                
                copy_headers_list = []
                
                if not best_variants:
                    tk.Label(content, text="No known formats found for this importer.",
                             font=("Segoe UI", 10), bg="white", fg="#991B1B").pack(anchor="w")
                elif best_score == 0:
                    tk.Label(content, text="We couldn't recognize any of your column headers.\nThis may not be the right format — contact support if needed.",
                             font=("Segoe UI", 9), bg="white", fg="#6B7280", justify="left",
                             wraplength=500).pack(anchor="w", pady=(0, 8))
                    tk.Label(content, text="If this is a known format, rename your columns to exactly match:",
                             font=("Segoe UI", 10), bg="white", fg="#1E293B").pack(anchor="w", pady=(0, 4))
                    _, fmt_mapping = best_variants[0]
                    copy_headers_list = _build_single_variant_guidance(content, fmt_mapping)
                elif len(best_variants) == 1:
                    _, fmt_mapping = best_variants[0]
                    copy_headers_list = _build_single_variant_guidance(content, fmt_mapping)
                else:
                    # Tie — show multiple options
                    tk.Label(content, text="Your file matches multiple known formats equally.\nPick the one that fits your data:",
                             font=("Segoe UI", 9), bg="white", fg="#6B7280", wraplength=500,
                             justify="left").pack(anchor="w", pady=(0, 6))
                    
                    opts = ['A', 'B', 'C', 'D']
                    copy_headers_per_option = {}
                    for idx, (fmt_name, fmt_mapping) in enumerate(best_variants):
                        opt_label = opts[idx] if idx < len(opts) else str(idx + 1)
                        
                        targets = [str(v).lower().strip() for v in fmt_mapping.values()]
                        has_coo = 'country of origin' in targets
                        desc = "with Country of Origin" if has_coo else "without Country of Origin"
                        
                        opt_frame = tk.LabelFrame(content, text=f"  Option {opt_label} ({desc})  ",
                                                  font=("Segoe UI", 9, "bold"), bg="white",
                                                  fg="#1E293B", padx=8, pady=6)
                        opt_frame.pack(fill=tk.X, pady=4)
                        
                        hdrs = _build_single_variant_guidance(opt_frame, fmt_mapping)
                        copy_headers_per_option[opt_label] = hdrs
                    
                    if copy_headers_per_option:
                        copy_headers_list = list(copy_headers_per_option.values())[0]
                
                # Separator + action buttons
                sep = tk.Frame(top, bg="#E5E7EB", height=1)
                sep.pack(fill=tk.X, padx=16)
                
                btn_frame = tk.Frame(top, bg="white", pady=10, padx=16)
                btn_frame.pack(fill=tk.X)
                
                def _copy_headers():
                    if copy_headers_list:
                        tab_separated = "\t".join(copy_headers_list)
                        top.clipboard_clear()
                        top.clipboard_append(tab_separated)
                        copy_btn.config(text="✓ Copied to clipboard!", fg="white", bg="#059669")
                        top.after(1500, lambda: copy_btn.config(
                            text="📋  Copy full header row (paste into Excel Row 1)",
                            fg="white", bg="#2563EB"))
                
                copy_btn = tk.Button(btn_frame,
                                     text="📋  Copy full header row (paste into Excel Row 1)",
                                     command=_copy_headers, font=("Segoe UI", 9, "bold"),
                                     bg="#2563EB", fg="white", relief="flat", padx=12, pady=4,
                                     cursor="hand2")
                copy_btn.pack(side=tk.LEFT)
                
                tk.Button(btn_frame, text="Close", command=top.destroy, font=("Segoe UI", 9),
                          relief="flat", padx=12, pady=4, bg="#F3F4F6", fg="#374151",
                          cursor="hand2").pack(side=tk.RIGHT)
                
                # Size the window to fit content, capped at a reasonable height
                top.update_idletasks()
                req_h = min(content.winfo_reqheight() + 140, 600)
                top.geometry(f"600x{req_h}")
                
                self.root.wait_window(top)
                return

            mapped_cols = {}
            expected_headers_lower = {k.lower(): v for k, v in expected_headers.items()}
            for col in invoice_df.columns:
                col_str = str(col).strip()
                if col_str in expected_headers:
                    mapped_cols[col] = expected_headers[col_str]
                elif col_str.lower() in expected_headers_lower:
                    mapped_cols[col] = expected_headers_lower[col_str.lower()]
                    
            if not mapped_cols:
                self.root.config(cursor="")
                self.status_label.config(text="")
                messagebox.showerror("Error", "No valid mappings applied.")
                return
                
            model_excel_col = None
            for col, target in mapped_cols.items():
                if target == 'Model' or target.lower() == 'model':
                    model_excel_col = col
                    break
                    
            if not model_excel_col:
                self.root.config(cursor="")
                self.status_label.config(text="")
                messagebox.showerror("Error", "No mapping found for 'Model'. It is required.")
                return
                
            target_to_col_idx = {}
            for target in set(mapped_cols.values()):
                t_lower = target.lower()
                for i, h in enumerate(checklist_logic.LOGISYS_HEADERS):
                    if str(h).strip().lower() == t_lower:
                        target_to_col_idx[target] = f"col_{i+1}"
                        break
                        
            queue_items = []
            for _, row in invoice_df.iterrows():
                model_val = str(row[model_excel_col]).strip() if pd.notna(row[model_excel_col]) else ""
                if not model_val:
                    continue
                    
                prefilled = {}
                for col, target in mapped_cols.items():
                    if target in target_to_col_idx:
                        val = row[col]
                        # Handle datetime objects
                        if hasattr(val, 'strftime'):
                            val = val.strftime('%d-%m-%Y')
                        prefilled[target_to_col_idx[target]] = str(val).strip() if pd.notna(val) else ""
                        
                queue_items.append({'model': model_val, 'prefilled': prefilled})
                
            if queue_items:
                if not hasattr(self, 'session_queue'):
                    self.session_queue = []
                self.session_queue.extend(queue_items)
                
                self.status_label.config(text=f"⚙️ Matching {len(queue_items)} models...")
                self.root.update_idletasks()
                self._process_model_queue(queue_items)
            else:
                self.root.config(cursor="")
                self.status_label.config(text="")
                messagebox.showinfo("Info", "No valid rows found in the uploaded file.")
                
        _brand_button(input_row, "+ Add Models", do_add_models).pack(side=tk.LEFT, padx=(10,0))
        _brand_button(status_row, "✖ Clear", _do_clear, secondary=True).pack(side=tk.LEFT, padx=(10,0))
        _brand_button(input_row, "📁 Upload Invoice Excel", _do_upload_invoice, secondary=True).pack(side=tk.LEFT, padx=(10,0))
        self.btn_export = _brand_button(input_row, "📥 Generate CSV", self._do_export)
        self.btn_export.pack(side=tk.LEFT, padx=(10,0))
        self.lbl_unresolved_count = tk.Label(input_row, text="", font=("Segoe UI", 9, "bold"), fg="#D97706", bg=_PANEL_WHITE)
        self.lbl_unresolved_count.pack(side=tk.LEFT, padx=(10,0))
        self.status_label.pack(side=tk.RIGHT)
        
        # 3. Tabs Area (Notebook)
        notebook_container = tk.Frame(self.body_container)
        notebook_container.pack(fill=tk.BOTH, expand=True)
        self.notebook = ttk.Notebook(notebook_container)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        tab_preview = tk.Frame(self.notebook, bg=_PANEL_WHITE)
        tab_missing = tk.Frame(self.notebook, bg=_PANEL_WHITE)
        self.notebook.add(tab_preview, text="Checklist Preview")
        self.notebook.add(tab_missing, text="Models Not Found (0)")
        
        # --- Checklist Preview Tab ---
        # Resolution panel (shown only when unresolved multi-match models exist)
        self.resolution_panel = tk.Frame(tab_preview, bg="#FFFBEB", bd=1, relief=tk.SOLID)
        # Packed/unpacked dynamically in _rebuild_resolution_panel
        
        table_frame = tk.Frame(tab_preview, bg=_PANEL_WHITE)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Instantiate tksheet Sheet widget
        wrapped_headers = [wrap_header(h) for h in checklist_logic.LOGISYS_HEADERS]
        if self.importer and self.importer.strip().lower() == 'ansell':
            wrapped_headers[3] = wrap_header("Product Desc [Auto]")
        self.sheet = Sheet(
            table_frame,
            headers=wrapped_headers,
            empty_horizontal=0,
            empty_vertical=0,
            show_row_index=True,
            show_column_header=True,
            row_height=28
        )
        self.sheet.pack(fill=tk.BOTH, expand=True)
        
        # Apply themes and layout options
        self.sheet.set_options(
            table_bg="#F9FAFB",
            even_rows_bg="#F9FAFB",
            odd_rows_bg="#FFFFFF",
            header_bg="#F3F4F6",
            header_fg=_PRIMARY_BLUE,
            header_font=("Segoe UI", 10, "bold"),
            header_grid_color=_BORDER_GRAY,
            grid_color=_BORDER_GRAY,
            selected_cells_bg=_HOVER_BLUE,
            selected_cells_fg="#FFFFFF",
            selected_rows_bg=_HOVER_BLUE,
            selected_rows_fg="#FFFFFF",
            row_index_bg="#F3F4F6",
            row_index_fg=_MUTED_GRAY
        )
        self.sheet.set_header_height_lines(1)
        self.sheet.set_column_widths(GRID_WIDTHS)
        
        # Highlight headers for master columns
        for col in range(len(checklist_logic.LOGISYS_HEADERS)):
            if (col + 1) in MASTER_INDICES:
                self.sheet.highlight_cells(row="all", column=col, canvas="header", bg="#D6E4F0", fg=_PRIMARY_BLUE)
            
        # Enable basic bindings
        self.sheet.enable_bindings(
            "edit_cell",
            "single_select",
            "row_select",
            "column_width_resize",
            "row_height_resize",
            "arrowkeys",
            "copy",
            "rc_select"
        )
        self.sheet.extra_bindings("end_edit_cell", self._on_cell_edited)
        self.root.bind("<Delete>", self._on_delete_key)
        self.root.bind("<BackSpace>", self._on_delete_key)
        self.sheet.popup_menu_add_command("✅ Resolve Conflict: Keep this row", self._resolve_conflict)
        
        # --- Missing Models Tab ---
        session = auth.get_current_session()
        can_add = session.get('Can_Add_Edit_Model', False) if isinstance(session, dict) else False
        
        self.missing_table_frame = tk.Frame(tab_missing, bg=_PANEL_WHITE)
        self.missing_table_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Success Frame overlay when empty
        self.success_frame = tk.Frame(self.missing_table_frame, bg=_PANEL_WHITE)
        self.success_msg = tk.Label(self.success_frame, text="", font=("Segoe UI", 16, "bold"), fg="#10B981", bg=_PANEL_WHITE)
        self.success_msg.pack(expand=True, pady=100)
        
        self.missing_sheet = Sheet(
            self.missing_table_frame,
            headers=[wrap_header(h) for h in PENDING_HEADERS],
            empty_horizontal=0,
            empty_vertical=0,
            show_row_index=True,
            show_column_header=True,
            row_height=28
        )
        self.missing_sheet.set_header_height_lines(1)
        
        # Determine column widths for missing sheet
        missing_widths = []
        for h in PENDING_HEADERS:
            h_clean = h.strip()
            # Ensure the width is at least enough to fit the header text clearly
            min_header_width = len(h_clean) * 7 + 30
            
            if h_clean in NARROW_COLS:
                base_w = 80
            elif h_clean in MEDIUM_COLS:
                base_w = 140
            elif h_clean in WIDE_COLS:
                base_w = 250
            else:
                h_lower = h_clean.lower()
                if any(term in h_lower for term in ['rate', 'sno', 'sr.no', 'srno', 'cth']):
                    base_w = 80
                elif any(term in h_lower for term in ['notn', 'notification']):
                    base_w = 140
                else:
                    base_w = 120
                    
            missing_widths.append(max(base_w, min_header_width))
            
        self.missing_sheet.set_column_widths(missing_widths)
        
        # Hide the Importer column from the UI
        if "Importer" in PENDING_HEADERS:
            self.missing_sheet.hide_columns(columns=[PENDING_HEADERS.index("Importer")])
        
        # Enable editing if allowed
        if not can_add:
            self.missing_sheet.enable_bindings(
                "single_select",
                "row_select",
                "column_width_resize",
                "row_height_resize",
                "arrowkeys",
                "copy",
                "rc_select"
            )
        else:
            self.missing_sheet.enable_bindings(
                "edit_cell",
                "single_select",
                "row_select",
                "column_width_resize",
                "row_height_resize",
                "arrowkeys",
                "copy",
                "rc_select",
                "paste",
                "delete",
                "undo"
            )
            
        # Context columns are always read-only: Importer (0), Added_By (1), Added_At (2), Model (3), Product Desc (4), Country of Origin (5)
        self.missing_sheet.readonly_columns(columns=[0, 1, 2, 3, 4, 5], readonly=True)
            
        # Tab Actions Frame (aligned with the Notebook tabs on the top right)
        self.tab_actions_frame = tk.Frame(notebook_container, bg=_LIGHT_BG)
        self.tab_actions_frame.place(relx=1.0, x=-10, y=2, anchor="ne")
        
        self.btn_recheck = _brand_button(self.tab_actions_frame, "🔄 Re-check Models", self._do_recheck_models, secondary=True)
        self.btn_submit_master = _brand_button(self.tab_actions_frame, "📤 Update Master Sheet", self._do_submit_master, bg_color=_PRIMARY_BLUE)
        
        if not can_add:
            self.btn_submit_master.config(state="disabled", bg="#9CA3AF")
            self.view_only_label = tk.Label(self.tab_actions_frame, text="(View Only - No Permission)", fg="#DC2626", bg=_LIGHT_BG, font=("Segoe UI", 9, "bold"))
        else:
            self.view_only_label = None

        def on_tab_changed(event):
            selected_tab = event.widget.select()
            tab_text = event.widget.tab(selected_tab, "text")
            
            # Hide all
            self.btn_recheck.pack_forget()
            self.btn_submit_master.pack_forget()
            if self.view_only_label:
                self.view_only_label.pack_forget()
                
            if "Checklist Preview" in tab_text:
                self.btn_recheck.pack(side=tk.RIGHT)
            else:
                self.btn_submit_master.pack(side=tk.RIGHT)
                if self.view_only_label:
                    self.view_only_label.pack(side=tk.RIGHT, padx=(0, 10))
                    
        self.notebook.bind("<<NotebookTabChanged>>", on_tab_changed)

        # Reload current rows (when switching importer screens)
        self._reload_table_rows()
        self._refresh_pending_tab()


    def _reload_table_rows(self):
        self.sheet.pack(fill=tk.BOTH, expand=True)
        
        # Map self.rows_data dict list to data list of lists
        data_list = []
        for row in self.rows_data:
            row_list = []
            for c_idx in range(1, len(checklist_logic.LOGISYS_HEADERS) + 1):
                row_list.append(row.get(f"col_{c_idx}", ""))
            data_list.append(row_list)
            
        self.sheet.set_sheet_data(data_list, redraw=False)
        
        # Apply Nagarkot standard styling options to sheet
        self.sheet.set_options(
            table_bg="#F9FAFB",
            even_rows_bg="#F9FAFB",
            odd_rows_bg="#FFFFFF",
            header_bg="#F3F4F6",
            header_fg=_PRIMARY_BLUE,
            header_font=("Segoe UI", 10, "bold"),
            header_grid_color=_BORDER_GRAY,
            grid_color=_BORDER_GRAY,
            selected_cells_bg=_HOVER_BLUE,
            selected_cells_fg="#FFFFFF",
            selected_rows_bg=_HOVER_BLUE,
            selected_rows_fg="#FFFFFF",
            row_index_bg="#F3F4F6",
            row_index_fg=_MUTED_GRAY
        )
        self.sheet.set_header_height_lines(2)
        self.sheet.set_column_widths(GRID_WIDTHS)
        
        # Dynamically align columns based on data types
        align_sheet_columns(self.sheet, checklist_logic.LOGISYS_HEADERS)
        
        # Clear existing highlights first
        self.sheet.dehighlight_all()
        
        # Highlight headers for master columns
        for col in range(len(checklist_logic.LOGISYS_HEADERS)):
            if (col + 1) in MASTER_INDICES:
                self.sheet.highlight_cells(row="all", column=col, canvas="header", bg="#D6E4F0")
                
        # Highlight missing and unresolved rows cell-by-cell (overriding alternates)
        for r_idx, row in enumerate(self.rows_data):
            if row.get('_missing_master'):
                for col_idx in range(len(checklist_logic.LOGISYS_HEADERS)):
                    self.sheet.highlight_cells(row=r_idx, column=col_idx, bg="#FCA5A5") # light red
            elif row.get('_unresolved_group') or row.get('_awaiting_resolution'):
                for col_idx in range(len(checklist_logic.LOGISYS_HEADERS)):
                    self.sheet.highlight_cells(row=r_idx, column=col_idx, bg="#FDE047") # light yellow
            
        self.sheet.redraw()
        self._rebuild_resolution_panel()
        self._update_missing_tab_visibility()

    def _rebuild_resolution_panel(self):
        # Clear existing widgets
        for widget in self.resolution_panel.winfo_children():
            widget.destroy()
            
        # Find first unresolved group
        target_group = None
        target_model = None
        for row in self.rows_data:
            if row.get('_unresolved_group'):
                target_group = row['_unresolved_group']
                target_model = row.get('col_24', 'Unknown Model')
                break
                
        if not target_group:
            self.resolution_panel.pack_forget()
            return
            
        # We have an unresolved group. Build the UI.
        self.resolution_panel.pack(fill=tk.X, padx=5, pady=5)
        
        header = tk.Label(self.resolution_panel, 
                          text=f"⚠ Model '{target_model}' has multiple master matches. Choose the correct one below:",
                          font=("Segoe UI", 10, "bold"), fg="#D97706", bg="#FFFBEB")
        header.pack(anchor="w", padx=10, pady=(5, 5))
        
        btn_container = tk.Frame(self.resolution_panel, bg="#FFFBEB")
        btn_container.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        # Build a button for each candidate row in this group
        for r_idx, row in enumerate(self.rows_data):
            if row.get('_unresolved_group') == target_group:
                cth = row.get('col_9', '')
                notn = row.get('col_10', '')
                srno = row.get('col_11', '')
                details = f"Row {r_idx+1}: CTH {cth}" if cth else f"Row {r_idx+1}: No CTH"
                if notn:
                    details += f" | Notn {notn}"
                    if srno:
                        details += f" (SrNo {srno})"
                        
                row_frame = tk.Frame(btn_container, bg="#FFFBEB")
                row_frame.pack(fill=tk.X, pady=2)
                
                # We need to capture r_idx properly in the lambda
                btn = tk.Button(row_frame, text="✓ Use This Row", bg=_PRIMARY_BLUE, fg="white",
                                font=("Segoe UI", 9, "bold"), relief=tk.FLAT, padx=10, pady=2,
                                command=lambda r=r_idx: self._resolve_conflict_for_row(r))
                btn.pack(side=tk.LEFT, padx=(0, 10))
                
                lbl = tk.Label(row_frame, text=details, font=("Segoe UI", 9), bg="#FFFBEB", fg="#1F2937")
                lbl.pack(side=tk.LEFT)


    def _update_missing_tab_visibility(self):
        try:
            missing_rows = self.missing_sheet.get_sheet_data()
            missing_count = len([r for r in missing_rows if any(r)])
        except Exception:
            missing_count = 0
            
        self.notebook.tab(1, text=f"Models Not Found ({missing_count})")
        
        export_blocked = False
        try:
            if hasattr(self, 'session_queue') and missing_count > 0:
                pending_models_in_sheet = [str(r[PENDING_HEADERS.index("Model")]).strip().lower() for r in missing_rows if any(r)]
                for item in self.session_queue:
                    model = item.get('model', '').strip().lower() if isinstance(item, dict) else str(item).strip().lower()
                    if model in pending_models_in_sheet:
                        export_blocked = True
                        break
        except Exception:
            pass
            
        unresolved_models = set(str(row.get('col_24', '')).strip().lower() for row in self.rows_data if row.get('_unresolved_group'))
        unresolved_count = len(unresolved_models)
        
        if hasattr(self, 'lbl_unresolved_count'):
            if unresolved_count > 0:
                self.lbl_unresolved_count.config(text=f"⚠ {unresolved_count} model(s) need conflict resolution")
            else:
                self.lbl_unresolved_count.config(text="")
                
        if unresolved_count > 0:
            export_blocked = True
        
        if missing_count == 0:
            if hasattr(self, 'btn_export'):
                self.btn_export.config(state="normal" if not export_blocked else "disabled", 
                                       bg=_PRIMARY_BLUE if not export_blocked else "#9CA3AF")
            self.missing_sheet.pack_forget()
            if hasattr(self, 'rows_data') and self.rows_data:
                self.success_msg.config(text="🎉 No pending models for this importer!")
            else:
                self.success_msg.config(text="")
            self.success_frame.pack(fill=tk.BOTH, expand=True)
        else:
            if hasattr(self, 'btn_export'):
                if export_blocked:
                    self.btn_export.config(state="disabled", bg="#9CA3AF")
                else:
                    self.btn_export.config(state="normal", bg=_PRIMARY_BLUE)
            self.success_frame.pack_forget()
            self.missing_sheet.pack(fill=tk.BOTH, expand=True)
            
            # Apply standard styling and alignments to missing sheet
            self.missing_sheet.set_options(
                table_bg="#F9FAFB",
                even_rows_bg="#F9FAFB",
                odd_rows_bg="#FFFFFF",
                header_bg="#F3F4F6",
                header_fg=_PRIMARY_BLUE,
                header_font=("Segoe UI", 10, "bold"),
                header_grid_color=_BORDER_GRAY,
                grid_color=_BORDER_GRAY,
                selected_cells_bg=_HOVER_BLUE,
                selected_cells_fg="#FFFFFF",
                row_index_bg="#F3F4F6",
                row_index_fg=_MUTED_GRAY
            )
            align_sheet_columns(self.missing_sheet, PENDING_HEADERS)
            self.missing_sheet.redraw()


    def _on_cell_edited(self, event):
        r = event.row
        c = event.column
        new_val = str(event.value).strip()
        col_key = f"col_{c+1}"
        
        # Validate cell input
        success, err_msg = self._validate_cell(f"row_{r}", col_key, new_val)
        if not success:
            messagebox.showerror("Validation Error", err_msg)
            # Revert cell value in sheet
            old_val = self.rows_data[r].get(col_key, '')
            self.sheet.set_cell_data(r, c, old_val, redraw=True)
            return
            
        self.rows_data[r][col_key] = new_val

    def _resolve_conflict(self, event=None):
        try:
            cells = self.sheet.get_selected_cells()
        except Exception:
            return
        if not cells: return
        r = cells[0][0]
        self._resolve_conflict_for_row(r)
        
    def _resolve_conflict_for_row(self, r):
        if r >= len(self.rows_data): return
        
        row_data = self.rows_data[r]
        group_id = row_data.get('_unresolved_group')
        if not group_id: return
        
        model = row_data.get('col_24')
        model_key = "".join(str(model).split()).lower()
        
        cth = row_data.get('col_9', '')
        notn = row_data.get('col_10', '')
        srno = row_data.get('col_11', '')
        rate = row_data.get('col_15', '') # In case there's a rate, though not mapped by default
        details = f"CTH: {cth}" if cth else "No CTH"
        if notn:
            details += f" | Basic Notn: {notn}"
            if srno:
                details += f" (SrNo: {srno})"
                
        # Count total occurrences (1 group of candidates + N placeholders)
        count_subsequent = sum(1 for r_row in self.rows_data if r_row.get('_awaiting_resolution') == model_key)
        total_occurrences = 1 + count_subsequent
        
        msg = f"You selected:\n\nModel: {model}\n{details}\n\nApply this master data to all {total_occurrences} occurrence(s) of this model in the checklist?"
        if not messagebox.askyesno("Confirm Resolution", msg):
            return
            
        master_keys = ['col_4', 'col_9', 'col_10', 'col_11', 'col_22', 'col_25', 'col_32', 'col_41', 'col_42', 'col_54', 'col_55', 'col_81', 'col_82']
        
        new_rows = []
        for r_idx, row in enumerate(self.rows_data):
            if row.get('_unresolved_group') == group_id:
                if r_idx != r:
                    continue # Skip this un-picked candidate
                else:
                    # Remove the unresolved flag from the picked candidate
                    row.pop('_unresolved_group', None)
                    new_rows.append(row)
            elif row.get('_awaiting_resolution') == model_key:
                # Update placeholder row with master data from picked candidate
                for mk in master_keys:
                    row[mk] = row_data.get(mk, "")
                row.pop('_awaiting_resolution', None)
                new_rows.append(row)
            else:
                new_rows.append(row)
            
        self.rows_data = new_rows
        self._reload_table_rows()

    def _on_delete_key(self, event):
        if isinstance(event.widget, (tk.Entry, tk.Text)):
            return
        if not hasattr(self, 'sheet') or not self.sheet.winfo_viewable():
            return
        try:
            cells = self.sheet.get_selected_cells()
        except Exception:
            return
        if not cells:
            return
        modified = False
        for r, c in cells:
            col_key = f"col_{c+1}"
            if self.rows_data[r].get(col_key, "") != "":
                self.rows_data[r][col_key] = ""
                self.sheet.set_cell_data(r, c, "", redraw=False)
                modified = True
        if modified:
            self.sheet.redraw()

    def _process_model_queue(self, queue):
        import uuid
        
        models_to_push = []
        # Track seen multi-match models: model_str -> group_id
        seen_multi_match = {}
        
        for item in queue:
            if isinstance(item, dict):
                model = item.get('model', '')
                prefilled = item.get('prefilled', {})
            else:
                model = item
                prefilled = {}
                
            model_key = "".join(str(model).split()).lower()
                
            matches = checklist_logic.lookup_model_local(self.importer_df, model)
            
            if len(matches) == 0:
                self._add_row_to_table(model, {}, prefilled, _missing_master=True)
                models_to_push.append({'model': model, 'prefilled': prefilled})
                
            elif len(matches) == 1:
                self._add_row_to_table(model, matches[0], prefilled)
            else:
                if model_key in seen_multi_match:
                    # Already generated candidates for this model. Add a placeholder row.
                    self._add_row_to_table(model, matches[0], prefilled, _awaiting_resolution=model_key)
                else:
                    # Add all candidate rows with a unique group ID
                    group_id = str(uuid.uuid4())
                    seen_multi_match[model_key] = group_id
                    for m in matches:
                        self._add_row_to_table(model, m, prefilled, _unresolved_group=group_id)
                    
        if models_to_push:
            # Send to Pending_Model_Approval and then refresh tab
            def _push_pending():
                try:
                    from gsheets import GSheetsClient
                    client = GSheetsClient(self.cred_path.get())
                    url = self.gsheets_url.get()
                    session = auth.get_current_session()
                    username = session.get('Username', 'User') if isinstance(session, dict) else 'User'
                    
                    pending_rows = []
                    for m_info in models_to_push:
                        model = m_info['model']
                        prefilled = m_info['prefilled']
                        
                        # Extract any mapped fields from invoice excel
                        extracted_vals = {"Product Desc": "", "Country of Origin": "", "Generic Description": ""}
                        for h_idx, header_name in enumerate(checklist_logic.LOGISYS_HEADERS):
                            col_key = f"col_{h_idx+1}"
                            if col_key in prefilled:
                                clean_header = str(header_name).strip().lower()
                                for k in extracted_vals.keys():
                                    if k.lower() == clean_header:
                                        extracted_vals[k] = prefilled[col_key]
                                        
                        if self.importer and self.importer.strip().lower() == 'ansell' and '_ansell_raw_desc' in prefilled:
                            raw_desc = prefilled.get('_ansell_raw_desc', '')
                            ship_case = str(prefilled.get('_ansell_case', '')).strip()
                            if ship_case.endswith('.00'):
                                ship_case = ship_case[:-3]
                            curr = prefilled.get('col_7') or extracted_vals.get('Currency', '')
                            unit_price = prefilled.get('_ansell_unit_price', '')
                            # Empty Generic Description for pending queue
                            extracted_vals["Product Desc"] = f"PRODUCT CODE. {model} {raw_desc} () [QTY {ship_case} CTN @ {curr} {unit_price} PER CTN]"
                                        
                        pr = {
                            "Model": model,
                            "Product Desc": extracted_vals["Product Desc"],
                            "Country of Origin": extracted_vals["Country of Origin"],
                            "Generic Description": extracted_vals["Generic Description"],
                            "Importer": self.importer,
                            "Added_By": username,
                            "Added_At": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                            "Status": "Pending"
                        }
                        pending_rows.append(pr)
                        
                    client.add_pending_models(url, pending_rows)
                except Exception as e:
                    print(f"Failed to push to Pending Model Approval queue: {e}")
                finally:
                    # Refresh the tab to show newly pending items
                    self.root.after(0, self._refresh_pending_tab)
                    
            threading.Thread(target=_push_pending, daemon=True).start()
        else:
            self._update_missing_tab_visibility()

        self.root.config(cursor="")
        if hasattr(self, 'status_label'):
            self.status_label.config(text="")

    def _add_row_to_table(self, model, master_data, prefilled=None, _missing_master=False, _unresolved_group=None, _awaiting_resolution=None):
        prefilled = prefilled or {}
        # Auto-inherit invoice header details from existing rows if typed
        inv_no, inv_date, toi, currency, brand = "", "", "", "", ""
        if self.rows_data and not prefilled:
            first_row = self.rows_data[0]
            inv_no = first_row.get('col_1', '')
            inv_date = first_row.get('col_2', '')
            toi = first_row.get('col_3', '')
            currency = first_row.get('col_7', '')
            brand = first_row.get('col_23', '')
            
        row_data = {}
        for c_idx in range(1, len(checklist_logic.LOGISYS_HEADERS) + 1):
            row_data[f"col_{c_idx}"] = ""
            
        for col_key, val in prefilled.items():
            row_data[col_key] = val
            
        # Set invoice-level columns if not already populated
        if not row_data.get('col_1'): row_data['col_1'] = inv_no
        if not row_data.get('col_2'): row_data['col_2'] = inv_date
        if not row_data.get('col_3'): row_data['col_3'] = toi
        if not row_data.get('col_7'): row_data['col_7'] = currency
        if not row_data.get('col_23'): row_data['col_23'] = brand
        
        # Map master-sourced columns using case-flexible matching (ONLY if not mapped from invoice excel already)
        row_data['col_24'] = model # Model
        
        # get_flexible is defined at module level in this file
        
        # Country of Origin Logic: Exclusively from invoice, UNLESS active variant has no COO mapping
        has_coo_mapping = False
        if hasattr(self, 'active_variant_targets'):
            for target in self.active_variant_targets:
                if str(target).strip().lower() == 'country of origin':
                    has_coo_mapping = True
                    break
                    
        if not has_coo_mapping and 'col_32' not in prefilled:
            row_data['col_32'] = ""
            
        if self.importer and self.importer.strip().upper() == 'BBRAUN':
            row_data['col_12'] = "NOEXCISE"
            row_data['col_39'] = "N"
            
        def safe_set(col_key, md_key):
            if not row_data.get(col_key) and master_data:
                row_data[col_key] = get_flexible(master_data, md_key)
                
        if self.importer and self.importer.strip().lower() == 'ansell':
            # Dynamic construction for Ansell (even if missing master)
            raw_desc = prefilled.get('_ansell_raw_desc', '')
            gen_desc = get_flexible(master_data, 'Generic Description') if master_data else ''
            ship_case = str(prefilled.get('_ansell_case', '')).strip()
            if ship_case.endswith('.00'):
                ship_case = ship_case[:-3]
            curr = prefilled.get('col_7') or row_data.get('col_7', '')
            unit_price = prefilled.get('_ansell_unit_price', '')
            
            # PRODUCT CODE. {Model} {raw} ({Generic}) [QTY {Ship Case} CTN @ {Currency} {Unit Price} PER CTN]
            # Enforce exactly one space after "PRODUCT CODE."
            constructed = f"PRODUCT CODE. {model} {raw_desc} ({gen_desc}) [QTY {ship_case} CTN @ {curr} {unit_price} PER CTN]"
            row_data['col_4'] = constructed
                
            safe_set('col_9', 'CTH')
            if not row_data.get('col_9'):
                row_data['col_9'] = prefilled.get('_ansell_hs', '')
        else:
            if master_data:
                row_data['col_4'] = get_flexible(master_data, 'Product Desc')
            else:
                safe_set('col_4', 'Product Desc')
            safe_set('col_9', 'CTH')
        safe_set('col_10', 'Basic Duty Notn')
        safe_set('col_11', 'Basic Duty Notn SNo')
        safe_set('col_22', 'Generic Description')
        safe_set('col_25', 'End Use')
        safe_set('col_41', 'IGST Notification')
        safe_set('col_42', 'IGST Notification SrNo')
        safe_set('col_54', 'SWS Notification')
        safe_set('col_55', 'SWS Notification SrNo')
        safe_set('col_81', 'AIDC Notn (Customs)')
        safe_set('col_82', 'AIDC Notn Sr.No.(Customs)')
        
        # Set manual default values
        

        if _missing_master: row_data['_missing_master'] = True
        if _unresolved_group: row_data['_unresolved_group'] = _unresolved_group
        if _awaiting_resolution: row_data['_awaiting_resolution'] = _awaiting_resolution
        
        self.rows_data.append(row_data)
        
        # Reload spreadsheet to reflect additions
        self._reload_table_rows()

    def _validate_cell(self, row_id, col_name, new_val):
        if col_name == 'col_5': # Quantity
            if not new_val:
                return False, "'Quantity' cannot be empty."
            try:
                float(new_val)
            except ValueError:
                return False, "'Quantity' must be numeric."
        elif col_name in ['col_8', 'col_16']: # Rate or Amount
            if new_val:
                try:
                    float(new_val)
                except ValueError:
                    name = "Rate" if col_name == 'col_8' else "Amount"
                    return False, f"'{name}' must be numeric."
        return True, ""

    def _do_export(self):
        if not self.rows_data:
            messagebox.showerror("Error", "No rows added to the checklist.")
            return
            
        for row in self.rows_data:
            if row.get('_unresolved_group') or row.get('_awaiting_resolution'):
                messagebox.showerror("Validation Error", "You have unresolved model conflicts (highlighted in yellow).\nRight click the correct candidate row and select '✅ Resolve Conflict: Keep this row'.")
                return
                
        default_name = f"checklist_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        p = filedialog.asksaveasfilename(
            defaultextension=".csv", 
            filetypes=[("CSV Files", "*.csv")], 
            title="Save Checklist",
            initialfile=default_name
        )
        if not p:
            return
            
        try:
            checklist_logic.export_checklist(self.rows_data, p)
            
            # Log audit trail
            models_added = [r['col_24'] for r in self.rows_data]
            checklist_logic.log_audit_action(self.gsheets_url.get(), self.cred_path.get(), auth.get_current_session(), self.importer, models_added)
            
            messagebox.showinfo("Success", f"Checklist successfully generated!\nSaved to: {p}")
                
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export template: {str(e)}")

    def _refresh_pending_tab(self):
        if not self.importer:
            return
            
        self.missing_sheet.set_sheet_data([])
        self.btn_submit_master.config(state="disabled")
        self.btn_recheck.config(state="disabled")
        self.root.update_idletasks()
        
        def load_thread():
            try:
                from gsheets import GSheetsClient
                client = GSheetsClient(self.cred_path.get())
                df, _ = client.get_sheet_data(self.gsheets_url.get(), "Pending_Model_Approval")
                
                rows_to_show = []
                if not df.empty:
                    # Filter by status and importer
                    pending_df = df[(df['Status'].str.strip().str.lower() == 'pending') & (df['Importer'].str.strip() == self.importer)]
                    for _, row in pending_df.iterrows():
                        r_data = []
                        for h in PENDING_HEADERS:
                            r_data.append(str(row.get(h, '')).strip())
                        rows_to_show.append(r_data)
                
                self.root.after(0, lambda d=rows_to_show: self._on_pending_loaded(d))
            except Exception as e:
                self.root.after(0, lambda err=e: self._on_pending_err(err))
                
        threading.Thread(target=load_thread, daemon=True).start()
        
    def _on_pending_loaded(self, data):
        if hasattr(self, 'btn_recheck') and self.btn_recheck.winfo_exists():
            self.btn_recheck.config(state="normal")
            
        session = auth.get_current_session()
        can_add = session.get('Can_Add_Edit_Model', False) if isinstance(session, dict) else False
        if can_add and hasattr(self, 'btn_submit_master') and self.btn_submit_master.winfo_exists():
            self.btn_submit_master.config(state="normal", text="📤 Update Master Sheet (Selected Row)")
            
        # If user cannot edit, render duty columns empty to prevent confusion
        if not can_add:
            for r in data:
                # Keep Context columns (0..5), clear the rest (6..)
                for i in range(6, len(r)):
                    r[i] = ""
                    
        if hasattr(self, 'missing_sheet') and self.missing_sheet.winfo_exists():
            self.missing_sheet.set_sheet_data(data)
            self.missing_sheet.redraw()
            self._update_missing_tab_visibility()
            
    def _on_pending_err(self, err):
        if hasattr(self, 'btn_recheck') and self.btn_recheck.winfo_exists():
            self.btn_recheck.config(state="normal")
        messagebox.showerror("Error", f"Failed to load pending models:\n{err}")

    def _do_submit_master(self):
        selected_rows = self.missing_sheet.get_selected_rows()
        if not selected_rows:
            messagebox.showerror("Error", "Please select a row to submit.")
            return
            
        if len(selected_rows) > 1:
            messagebox.showerror("Error", "Please select only one row at a time to update.")
            return
            
        row_idx = list(selected_rows)[0]
        row_data = self.missing_sheet.get_row_data(row_idx)
        
        # Convert row data to dict
        row_dict = {h: row_data[i] for i, h in enumerate(PENDING_HEADERS)}
        
        # Ansell specific: Auto-fill Generic Description into the Product Desc template before saving to Master Sheet
        if self.importer and self.importer.strip().lower() == 'ansell':
            gen_desc = row_dict.get("Generic Description", "").strip()
            prod_desc = row_dict.get("Product Desc", "")
            if gen_desc and " () " in prod_desc:
                row_dict["Product Desc"] = prod_desc.replace(" () ", f" ({gen_desc}) ")
        
        # Validation
        if not row_dict.get("CTH"):
            messagebox.showerror("Validation Error", "CTH cannot be empty.")
            return
            
        # Validate Duty pairs (Notification+SrNo OR Rate)
        duty_pairs = [
            ("Basic Duty", "Basic Duty Notn", "Basic Duty Notn SNo", "Basic Duty Rate"),
            ("Customs Health Cess", "Customs Health Cess Notn", "Customs Health Cess SNo", "Customs Health Cess Rate"),
            ("SWS", "SWS Notification", "SWS Notification SrNo", "SWS Rate"),
            ("IGST", "IGST Notification", "IGST Notification SrNo", "IGST Rate"),
            ("AIDC", "AIDC Notn (Customs)", "AIDC Notn Sr.No.(Customs)", None)
        ]
        
        for d_name, notn_col, srno_col, rate_col in duty_pairs:
            has_notn = bool(row_dict.get(notn_col))
            has_srno = bool(row_dict.get(srno_col))
            has_rate = bool(row_dict.get(rate_col)) if rate_col else False
            
            if has_notn or has_srno or has_rate:
                if rate_col:
                    if not ((has_notn and has_srno) ^ has_rate):
                        if not messagebox.askyesno("Duty Validation", f"For {d_name}, you must provide EITHER (Notification + SrNo) OR Rate, not both and not incomplete.\n\nDo you want to proceed anyway?"):
                            return
                else:
                    if has_notn != has_srno:
                        if not messagebox.askyesno("Duty Validation", f"For {d_name}, you must provide BOTH Notification and SrNo.\n\nDo you want to proceed anyway?"):
                            return
                            
        self.btn_submit_master.config(state="disabled", text="⏳ Updating...")
        self.root.config(cursor="watch")
        self.root.update_idletasks()
        
        def submit_thread():
            try:
                session = auth.get_current_session()
                if not session or not session.get('Can_Add_Edit_Model', False):
                    raise Exception("You do not have permission to add/edit models.")
                    
                from gsheets import GSheetsClient
                client = GSheetsClient(self.cred_path.get())
                url = self.gsheets_url.get()
                
                sheet_id = client.extract_sheet_id(url)
                spreadsheet = client.client.open_by_key(sheet_id)
                importer_name = row_dict.get("Importer")
                
                try:
                    worksheet = spreadsheet.worksheet(importer_name)
                except Exception:
                    raise Exception(f"Importer sheet '{importer_name}' not found.")
                    
                live_headers = [str(h) for h in worksheet.row_values(1)]
                if len(live_headers) != len(MASTER_SHEET_COLUMNS):
                    raise Exception(f"Master sheet structure count mismatch for {importer_name}.")
                    
                for pos, (live_h, expected_h) in enumerate(zip(live_headers, MASTER_SHEET_COLUMNS)):
                    if live_h != expected_h:
                        raise Exception(f"Master sheet header mismatch at column {pos + 1}.")
                        
                master_row = [""] * len(MASTER_SHEET_COLUMNS)
                for h in MASTER_SHEET_COLUMNS:
                    if h in row_dict:
                        master_row[MASTER_SHEET_COLUMNS.index(h)] = row_dict[h]
                        
                client.append_rows(worksheet, [master_row])
                checklist_logic.log_audit_action(url, self.cred_path.get(), session, importer_name, [row_dict.get("Model")], action_type="model_added")
                client.mark_pending_model_resolved(url, row_dict.get("Model"), importer_name)
                
                self.root.after(0, lambda: self._on_submit_success())
            except Exception as e:
                self.root.after(0, lambda err=e: self._on_submit_error(err))
                
        threading.Thread(target=submit_thread, daemon=True).start()
        
    def _on_submit_success(self):
        if hasattr(self, 'btn_submit_master') and self.btn_submit_master.winfo_exists():
            self.btn_submit_master.config(state="normal", text="📤 Update Master Sheet (Selected Row)")
        self.root.config(cursor="")
        messagebox.showinfo("Success", "Successfully appended row to Master Sheet and marked as resolved.")
        self._refresh_pending_tab()
        
    def _on_submit_error(self, err):
        if hasattr(self, 'btn_submit_master') and self.btn_submit_master.winfo_exists():
            self.btn_submit_master.config(state="normal", text="📤 Update Master Sheet (Selected Row)")
        self.root.config(cursor="")
        messagebox.showerror("Error", f"Failed to submit to master:\n{err}")
        
    def _do_recheck_models(self):
        if not self.importer:
            return
            
        self.root.config(cursor="watch")
        if hasattr(self, 'status_label'):
            self.status_label.config(text="⚙️ Re-checking models...")
        self.root.update_idletasks()
        
        def recheck_thread():
            try:
                from gsheets import GSheetsClient
                client = GSheetsClient(self.cred_path.get())
                df, _ = client.get_sheet_data(self.gsheets_url.get(), self.importer)
                if not df.empty:
                    for col in df.columns:
                        if str(col).strip().lower() == 'model':
                            df.rename(columns={col: 'Model'}, inplace=True)
                            break
                    if 'Model' in df.columns:
                        df['Model'] = df['Model'].astype(str).str.strip()
                
                self.root.after(0, lambda d=df: self._on_recheck_success(d))
            except Exception as e:
                self.root.after(0, lambda err=e: self._on_recheck_error(err))
                
        threading.Thread(target=recheck_thread, daemon=True).start()
        
    def _on_recheck_success(self, new_df):
        self.importer_df = new_df
        
        if hasattr(self, 'session_queue') and self.session_queue:
            # Replay the entire session queue against the fresh master sheet
            self.rows_data = []
            if hasattr(self, 'sheet') and self.sheet.winfo_exists():
                self.sheet.set_sheet_data([])
            try:
                if hasattr(self, 'missing_sheet') and self.missing_sheet.winfo_exists():
                    self.missing_sheet.set_sheet_data([])
            except Exception:
                pass
            self._process_model_queue(self.session_queue)
        else:
            # Fallback if queue is empty for some reason
            self._reload_table_rows()
            
        self.root.config(cursor="")
        if hasattr(self, 'status_label') and self.status_label.winfo_exists():
            self.status_label.config(text="")
            
        messagebox.showinfo("Re-check Complete", "Successfully re-checked models against latest master.")
        
    def _on_recheck_error(self, err):
        self.root.config(cursor="")
        if hasattr(self, 'status_label') and self.status_label.winfo_exists():
            self.status_label.config(text="")
        messagebox.showerror("Error", f"Failed to reload master: {err}")

