import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys
import threading
from pathlib import Path
from datetime import datetime

MISSING_HEADERS = [
    "Model", "Product Desc", "CTH", "Basic Duty Notn", "Basic Duty Notn SNo", "Basic Duty Rate", 
    "Customs Health Cess Notn", "Customs Health Cess SNo", "Customs Health Cess Rate", 
    "SWS Notification", "SWS Notification SrNo", "SWS Rate", "IGST Notification", 
    "IGST Notification SrNo", "IGST Rate", "AIDC Notn (Customs)", "AIDC Notn Sr.No.(Customs)", 
    "End Use", "Generic Description", "Country of Origin"
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
                        self.tab_names = [ws.title for ws in spreadsheet.worksheets() if ws.title not in ['Tool_Users', 'Audit_Log', 'Invoice_Header_Mappings']]
                        
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
                        self.root.after(0, lambda: handle_login_fail(msg))
                except Exception as e:
                    self.root.after(0, lambda: handle_login_fail(str(e)))
                    
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
                    elif self.importer.lower() == 'ansell':
                        if not file_path.lower().endswith('.pdf'):
                            raise ValueError("Ansell importer requires a PDF invoice.")
                        self.root.after(0, lambda: _parse_ansell_invoice(file_path))
                    else:
                        invoice_df = pd.read_excel(file_path)
                        self.root.after(0, lambda: _on_file_parsed(invoice_df))
                except Exception as e:
                    self.root.after(0, lambda err=e: _handle_upload_err(err))
                    
            threading.Thread(target=parse_thread, daemon=True).start()
            
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
                    
                prefilled = {
                    target_to_col_idx.get("Inv No"): rec.get("Invoice_No", ""),
                    target_to_col_idx.get("Date"): rec.get("Invoice_Date", ""),
                    target_to_col_idx.get("Product Desc"): rec.get("Product_Description", ""),
                    target_to_col_idx.get("Quantity"): rec.get("Qty", ""),
                    target_to_col_idx.get("Unit"): rec.get("UOM", ""),
                    target_to_col_idx.get("Currency"): rec.get("Currency", ""),
                    target_to_col_idx.get("Rate"): rec.get("Price", ""),
                    target_to_col_idx.get("Country of origin"): rec.get("Country_of_Origin", "")
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
            
        def _on_file_parsed(invoice_df):
            import pandas as pd
            importer_map = self.invoice_mappings_df[self.invoice_mappings_df['Importer'] == self.importer]
            expected_headers = {}
            for _, row in importer_map.iterrows():
                if pd.notna(row.get('Source_Header')) and pd.notna(row.get('Target_Field')):
                    expected_headers[str(row['Source_Header']).strip().lower()] = str(row['Target_Field']).strip()
                    
            if not expected_headers:
                self.root.config(cursor="")
                self.status_label.config(text="")
                messagebox.showerror("Error", f"No header mappings configured for importer '{self.importer}'.")
                return
                
            mapped_cols = {}
            for col in invoice_df.columns:
                col_str = str(col).strip()
                if col_str.lower() in expected_headers:
                    mapped_cols[col] = expected_headers[col_str.lower()]
                    
            if not mapped_cols:
                self.root.config(cursor="")
                self.status_label.config(text="")
                messagebox.showerror("Error", "None of the headers in the uploaded Excel match the configured mapping.")
                return
                
            model_excel_col = None
            for col, target in mapped_cols.items():
                if target.lower() == 'model':
                    model_excel_col = col
                    break
                    
            if not model_excel_col:
                top = tk.Toplevel(self.root)
                top.title("Select Model Column")
                top.geometry("400x150")
                top.grab_set()
                tk.Label(top, text="No mapping found for 'Model'.\nPlease select the column containing Model data:", pady=10).pack()
                
                sel_var = tk.StringVar()
                combo = ttk.Combobox(top, textvariable=sel_var, values=list(invoice_df.columns), state="readonly")
                combo.pack()
                if len(invoice_df.columns) > 0:
                    combo.current(0)
                    
                result_var = tk.StringVar(value="")
                def on_confirm():
                    result_var.set(sel_var.get())
                    top.destroy()
                
                tk.Button(top, text="Confirm", command=on_confirm).pack(pady=10)
                self.root.wait_window(top)
                
                if not result_var.get():
                    self.root.config(cursor="")
                    self.status_label.config(text="")
                    return 
                    
                model_excel_col = result_var.get()
                mapped_cols[model_excel_col] = 'Model'
                
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
        table_frame = tk.Frame(tab_preview, bg=_PANEL_WHITE)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Instantiate tksheet Sheet widget
        wrapped_headers = [wrap_header(h) for h in checklist_logic.LOGISYS_HEADERS]
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
            headers=[wrap_header(h) for h in MISSING_HEADERS],
            empty_horizontal=0,
            empty_vertical=0,
            show_row_index=True,
            show_column_header=True,
            row_height=28
        )
        self.missing_sheet.set_header_height_lines(1)
        
        # Determine column widths for missing sheet
        missing_widths = []
        for h in MISSING_HEADERS:
            h_clean = h.strip()
            if h_clean in NARROW_COLS:
                missing_widths.append(80)
            elif h_clean in MEDIUM_COLS:
                missing_widths.append(140)
            elif h_clean in WIDE_COLS:
                missing_widths.append(250)
            else:
                h_lower = h_clean.lower()
                if any(term in h_lower for term in ['rate', 'sno', 'sr.no', 'srno', 'cth']):
                    missing_widths.append(80)
                elif any(term in h_lower for term in ['notn', 'notification']):
                    missing_widths.append(140)
                else:
                    missing_widths.append(120)
        self.missing_sheet.set_column_widths(missing_widths)
        
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
        
        # Highlight headers for master columns
        for col in range(len(checklist_logic.LOGISYS_HEADERS)):
            if (col + 1) in MASTER_INDICES:
                self.sheet.highlight_cells(row="all", column=col, canvas="header", bg="#D6E4F0")
                
        # Highlight missing and unresolved rows cell-by-cell (overriding alternates)
        for r_idx, row in enumerate(self.rows_data):
            if row.get('_missing_master'):
                for col_idx in range(len(checklist_logic.LOGISYS_HEADERS)):
                    self.sheet.highlight_cells(row=r_idx, column=col_idx, bg="#FCA5A5") # light red
            elif row.get('_unresolved_group'):
                for col_idx in range(len(checklist_logic.LOGISYS_HEADERS)):
                    self.sheet.highlight_cells(row=r_idx, column=col_idx, bg="#FDE047") # light yellow
            
        self.sheet.redraw()
        self._update_missing_tab_visibility()

    def _update_missing_tab_visibility(self):
        try:
            missing_rows = self.missing_sheet.get_sheet_data()
            missing_count = len([r for r in missing_rows if any(r)])
        except Exception:
            missing_count = 0
            
        self.notebook.tab(1, text=f"Models Not Found ({missing_count})")
        
        if missing_count == 0:
            if hasattr(self, 'btn_export'):
                self.btn_export.config(state="normal", bg=_PRIMARY_BLUE)
            self.missing_sheet.pack_forget()
            if self.rows_data:
                self.success_msg.config(text="All models matched ✓")
            else:
                self.success_msg.config(text="")
            self.success_frame.pack(fill=tk.BOTH, expand=True)
        else:
            if hasattr(self, 'btn_export'):
                self.btn_export.config(state="disabled", bg="#9CA3AF")
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
            align_sheet_columns(self.missing_sheet, MISSING_HEADERS)
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
        if r >= len(self.rows_data): return
        
        row_data = self.rows_data[r]
        group_id = row_data.get('_unresolved_group')
        if not group_id: return
        
        # Remove all other rows with this group_id
        new_rows = []
        for r_idx, row in enumerate(self.rows_data):
            if row.get('_unresolved_group') == group_id and r_idx != r:
                continue # Skip this un-picked candidate
            if r_idx == r:
                # Remove the unresolved flag from the picked candidate
                row.pop('_unresolved_group', None)
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
        
        # Load current missing sheet data
        try:
            missing_data = self.missing_sheet.get_sheet_data()
        except Exception:
            missing_data = []
            
        for item in queue:
            if isinstance(item, dict):
                model = item.get('model', '')
                prefilled = item.get('prefilled', {})
            else:
                model = item
                prefilled = {}
                
            matches = checklist_logic.lookup_model_local(self.importer_df, model)
            
            if len(matches) == 0:
                self._add_row_to_table(model, {}, prefilled, _missing_master=True)
                
                # Add to missing sheet
                new_missing_row = [""] * len(MISSING_HEADERS)
                new_missing_row[MISSING_HEADERS.index("Model")] = model
                
                # Attempt to extract any mapped fields from invoice excel that match the missing sheet headers
                for h_idx, header_name in enumerate(checklist_logic.LOGISYS_HEADERS):
                    col_key = f"col_{h_idx+1}"
                    if col_key in prefilled:
                        clean_header = str(header_name).strip().lower()
                        for m_idx, m_header in enumerate(MISSING_HEADERS):
                            if str(m_header).strip().lower() == clean_header:
                                new_missing_row[m_idx] = prefilled[col_key]
                            
                missing_data.append(new_missing_row)
                
            elif len(matches) == 1:
                self._add_row_to_table(model, matches[0], prefilled)
            else:
                # Add all candidate rows with a unique group ID
                group_id = str(uuid.uuid4())
                for m in matches:
                    self._add_row_to_table(model, m, prefilled, _unresolved_group=group_id)
                    
        if missing_data:
            # Filter out entirely empty rows just in case
            missing_data = [r for r in missing_data if any(r)]
            self.missing_sheet.set_sheet_data(missing_data)
                    
        self.root.config(cursor="")
        if hasattr(self, 'status_label'):
            self.status_label.config(text="")
            
        self._update_missing_tab_visibility()

    def _add_row_to_table(self, model, master_data, prefilled=None, _missing_master=False, _unresolved_group=None):
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
        
        def safe_set(col_key, md_key):
            if not row_data.get(col_key) and master_data:
                row_data[col_key] = get_flexible(master_data, md_key)
                
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
        if not row_data.get('col_6'): row_data['col_6'] = 'PCS'
        
        if _missing_master: row_data['_missing_master'] = True
        if _unresolved_group: row_data['_unresolved_group'] = _unresolved_group
        
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
            if row.get('_unresolved_group'):
                messagebox.showerror("Validation Error", "You have unresolved model conflicts (highlighted in yellow).\nRight click the correct row and select '✅ Resolve Conflict: Keep this row'.")
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

    def _do_submit_master(self):
        try:
            missing_data = self.missing_sheet.get_sheet_data()
        except Exception:
            return
            
        if not missing_data:
            messagebox.showinfo("Info", "No missing models to submit.")
            return
            
        self.btn_submit_master.config(state="disabled", text="⏳ Submitting...")
        self.root.update_idletasks()
        
        def submit_thread():
            try:
                from gsheets import GSheetsClient
                client = GSheetsClient(self.cred_path.get())
                sheet_id = client.extract_sheet_id(self.gsheets_url.get())
                spreadsheet = client.client.open_by_key(sheet_id)
                worksheet = spreadsheet.worksheet(self.importer)
                
                # SAFETY CHECK: Read row 1 of the specific importer tab
                live_headers = [str(h) for h in worksheet.row_values(1)]
                
                # Check column count
                if len(live_headers) != len(MASTER_SHEET_COLUMNS):
                    raise Exception(
                        f"Master sheet structure count mismatch.\n"
                        f"Expected {len(MASTER_SHEET_COLUMNS)} columns, but found {len(live_headers)} columns.\n"
                        f"Refusing to write to prevent data corruption."
                    )
                
                # Position-by-position EXACT comparison (case-sensitive, no trimming)
                for pos, (live_h, expected_h) in enumerate(zip(live_headers, MASTER_SHEET_COLUMNS)):
                    if live_h != expected_h:
                        raise Exception(
                            f"Master sheet header mismatch at column {pos + 1}.\n"
                            f"Expected: '{expected_h}'\n"
                            f"Found: '{live_h}'\n"
                            f"Refusing to write to prevent data corruption."
                        )
                
                new_rows = []
                models_submitted = []
                for row in missing_data:
                    # Filter out empty rows
                    if not any(row): continue
                    
                    master_row = [""] * len(MASTER_SHEET_COLUMNS)
                    for m_idx, m_val in enumerate(row):
                        m_header = MISSING_HEADERS[m_idx]
                        # Position N of form field maps to corresponding column in MASTER_SHEET_COLUMNS
                        target_idx = MASTER_SHEET_COLUMNS.index(m_header)
                        master_row[target_idx] = str(m_val).strip()
                        
                    new_rows.append(master_row)
                    
                    model_idx = MISSING_HEADERS.index("Model")
                    if row[model_idx]:
                        models_submitted.append(row[model_idx])
                        
                if new_rows:
                    client.append_rows(worksheet, new_rows)
                    # Log audit action for each model
                    checklist_logic.log_audit_action(
                        self.gsheets_url.get(), 
                        self.cred_path.get(), 
                        auth.get_current_session(), 
                        self.importer, 
                        models_submitted,
                        action_type="model_added"
                    )
                    
                self.root.after(0, lambda: self._on_submit_success(len(new_rows)))
            except Exception as e:
                self.root.after(0, lambda err=e: self._on_submit_error(err))
                
        threading.Thread(target=submit_thread, daemon=True).start()
        
    def _on_submit_success(self, count):
        self.btn_submit_master.config(state="normal", text="📤 Update Master Sheet")
        self.missing_sheet.set_sheet_data([]) # Clear table
        messagebox.showinfo("Success", f"Successfully appended {count} row(s) to Master Sheet.\nPlease click 'Re-check Models' on the Preview tab.")
        
    def _on_submit_error(self, err):
        self.btn_submit_master.config(state="normal", text="📤 Update Master Sheet")
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
            self.sheet.set_sheet_data([])
            try:
                self.missing_sheet.set_sheet_data([])
            except Exception:
                pass
            self._process_model_queue(self.session_queue)
        else:
            # Fallback if queue is empty for some reason
            self._reload_table_rows()
            
        self.root.config(cursor="")
        if hasattr(self, 'status_label'):
            self.status_label.config(text="")
            
        messagebox.showinfo("Re-check Complete", "Successfully re-checked models against latest master.")
        
    def _on_recheck_error(self, err):
        self.root.config(cursor="")
        if hasattr(self, 'status_label'):
            self.status_label.config(text="")
        messagebox.showerror("Error", f"Failed to reload master: {err}")
