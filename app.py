import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import threading
import os
import sys
from pathlib import Path

from gsheets import GSheetsClient
from logic import compare_dataframes

_PRIMARY_BLUE = "#1F3F6E"; _ACCENT_RED = "#D8232A"; _DARK_TEXT = "#1E1E1E"
_MUTED_GRAY = "#6B7280"; _LIGHT_BG = "#F4F6F8"; _PANEL_WHITE = "#FFFFFF"
_BORDER_GRAY = "#E5E7EB"; _HOVER_BLUE = "#2A528F"; _HEADER_BG = "#D6E4F0"

def get_base_path() -> Path:
    return Path(sys._MEIPASS) if getattr(sys, 'frozen', False) else Path(__file__).parent

def _brand_button(parent, text: str, command, state=tk.NORMAL) -> tk.Button:
    btn = tk.Button(parent, text=text, command=command, state=state,
                    font=("Segoe UI",10,"bold"), fg="#FFF", bg=_PRIMARY_BLUE,
                    activebackground=_HOVER_BLUE, activeforeground="#FFF",
                    bd=0, padx=14, pady=6, cursor="hand2", relief=tk.FLAT)
    btn.bind("<Enter>", lambda e: btn.configure(bg=_HOVER_BLUE) if btn['state'] != 'disabled' else None)
    btn.bind("<Leave>", lambda e: btn.configure(bg=_PRIMARY_BLUE) if btn['state'] != 'disabled' else None)
    return btn

class MismatchCheckerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("General Item Mismatch Checker")
        self.root.configure(bg=_LIGHT_BG)
        self.root.state("zoomed")
        self.root.minsize(1024, 600)

        self.item_path = ""
        self.gsheets_url = tk.StringVar()
        self.cred_path = tk.StringVar(value=str(get_base_path() / "credentials.json"))
        
        self.gsheets_client = None
        self.worksheet = None
        self.master_cols = []
        
        self.mismatches = []
        self.new_models = []

        self._build_header()
        self._build_body()
        self._build_footer()

    def _build_header(self) -> None:
        header = tk.Frame(self.root, bg=_PANEL_WHITE, bd=0, highlightthickness=1, highlightbackground=_BORDER_GRAY)
        header.pack(fill=tk.X, side=tk.TOP); header.pack_propagate(False); header.configure(height=85)
        
        # Logo Left
        try:
            self._logo_img = tk.PhotoImage(file=str(get_base_path() / "logo.png"))
            factor = max(1, self._logo_img.height() // 20)
            self._logo_img = self._logo_img.subsample(factor)
            tk.Label(header, image=self._logo_img, bg=_PANEL_WHITE).pack(side=tk.LEFT, padx=20)
        except Exception:
            tk.Label(header, text="Nagarkot", font=("Segoe UI",12,"bold"), fg=_PRIMARY_BLUE, bg=_PANEL_WHITE).pack(side=tk.LEFT, padx=20)
            
        # Absolute Center Title
        tf = tk.Frame(header, bg=_PANEL_WHITE); tf.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        tk.Label(tf, text="General Item Mismatch Checker", font=("Segoe UI",18,"bold"), fg=_PRIMARY_BLUE, bg=_PANEL_WHITE).pack()
        tk.Label(tf, text="Automated Google Sheets Master Reconciliation", font=("Segoe UI",10), fg=_MUTED_GRAY, bg=_PANEL_WHITE).pack()

    def _build_body(self) -> None:
        body = tk.Frame(self.root, bg=_LIGHT_BG)
        body.pack(fill=tk.BOTH, expand=True, padx=16, pady=(12,4))

        # Inputs Frame
        input_frame = tk.Frame(body, bg=_LIGHT_BG)
        input_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(input_frame, text="Google Sheets URL:", bg=_LIGHT_BG, font=("Segoe UI", 10)).grid(row=0, column=0, sticky='w', pady=5)
        tk.Entry(input_frame, textvariable=self.gsheets_url, width=60, font=("Segoe UI", 10)).grid(row=0, column=1, padx=10, pady=5)

        tk.Label(input_frame, text="Credentials JSON:", bg=_LIGHT_BG, font=("Segoe UI", 10)).grid(row=1, column=0, sticky='w', pady=5)
        cred_entry = tk.Entry(input_frame, textvariable=self.cred_path, width=60, font=("Segoe UI", 10))
        cred_entry.grid(row=1, column=1, padx=10, pady=5)
        tk.Button(input_frame, text="Browse...", command=self._pick_cred, bg="#E5E7EB", bd=0, padx=10).grid(row=1, column=2)

        toolbar = tk.Frame(body, bg=_LIGHT_BG); toolbar.pack(fill=tk.X, pady=(0,6))
        self.btn_item = _brand_button(toolbar, "📂 Select Item Report", self._pick_item)
        self.btn_item.pack(side=tk.LEFT, padx=(0,10))

        self.btn_run = tk.Button(toolbar, text="▶  Run Comparison", font=("Segoe UI",10,"bold"),
                                 fg="#FFF", bg="#A0AAB5", activebackground="#b71c1c", activeforeground="#FFF",
                                 disabledforeground="#FFFFFF", bd=0, padx=18, pady=7, cursor="hand2", state=tk.DISABLED, command=self._run)
        self.btn_run.pack(side=tk.LEFT, padx=10)
        
        self.status_var = tk.StringVar(value="Provide inputs and select Item Report to begin")
        tk.Label(toolbar, textvariable=self.status_var, font=("Segoe UI",10,"bold"), fg=_PRIMARY_BLUE, bg=_LIGHT_BG).pack(side=tk.LEFT, padx=20)

        # Tabs for Mismatches and New Models
        self.notebook = ttk.Notebook(body)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=10)

        self.tab_mismatches = tk.Frame(self.notebook, bg=_PANEL_WHITE)
        self.notebook.add(self.tab_mismatches, text="Mismatches")
        
        self.tab_new = tk.Frame(self.notebook, bg=_PANEL_WHITE)
        self.notebook.add(self.tab_new, text="New Models")

        self._build_tables()

        # Bottom Actions
        action_frame = tk.Frame(body, bg=_LIGHT_BG)
        action_frame.pack(fill=tk.X, pady=5)
        self.btn_update = _brand_button(action_frame, "☁  Apply Selected Updates to Google Sheets", self._apply_updates, state=tk.DISABLED)
        self.btn_update.pack(side=tk.RIGHT)

    def _build_footer(self) -> None:
        ft = tk.Frame(self.root, bg=_PANEL_WHITE, height=28, highlightthickness=1, highlightbackground=_BORDER_GRAY)
        ft.pack(fill=tk.X, side=tk.BOTTOM); ft.pack_propagate(False)
        tk.Label(ft, text="Nagarkot Forwarders Pvt. Ltd. ©", font=("Segoe UI",8), fg=_MUTED_GRAY, bg=_PANEL_WHITE).pack(side=tk.LEFT, padx=12)

    def _build_tables(self):
        style = ttk.Style(); style.theme_use("clam")
        style.configure("Treeview", background=_PANEL_WHITE, foreground=_DARK_TEXT, rowheight=28)
        style.configure("Treeview.Heading", background=_HEADER_BG, foreground=_PRIMARY_BLUE, font=("Segoe UI",9,"bold"))
        style.map("Treeview", background=[("selected",_HOVER_BLUE)], foreground=[("selected","#FFF")])

        # Mismatches Tree
        self.tree_m = ttk.Treeview(self.tab_mismatches, columns=("Apply", "Model", "Field", "Master Value", "Item Value", "Job No"), show="headings")
        self.tree_m.heading("Apply", text="Update? (Y/N)")
        self.tree_m.heading("Model", text="Model")
        self.tree_m.heading("Field", text="Mismatch Field")
        self.tree_m.heading("Master Value", text="Current Master Value")
        self.tree_m.heading("Item Value", text="New Item Value")
        self.tree_m.heading("Job No", text="Job No")
        
        self.tree_m.column("Apply", width=100, anchor='center')
        self.tree_m.column("Model", width=150)
        self.tree_m.column("Field", width=150)
        self.tree_m.column("Master Value", width=150)
        self.tree_m.column("Item Value", width=150)
        self.tree_m.column("Job No", width=100)
        
        scroll_m = ttk.Scrollbar(self.tab_mismatches, orient=tk.VERTICAL, command=self.tree_m.yview)
        self.tree_m.configure(yscroll=scroll_m.set)
        self.tree_m.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_m.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree_m.bind('<Double-1>', self._toggle_mismatch)

        # New Models Tree
        self.tree_n = ttk.Treeview(self.tab_new, columns=("Apply", "Model", "Job No"), show="headings")
        self.tree_n.heading("Apply", text="Add? (Y/N)")
        self.tree_n.heading("Model", text="Model")
        self.tree_n.heading("Job No", text="Job No")
        
        self.tree_n.column("Apply", width=100, anchor='center')
        self.tree_n.column("Model", width=300)
        self.tree_n.column("Job No", width=150)
        
        scroll_n = ttk.Scrollbar(self.tab_new, orient=tk.VERTICAL, command=self.tree_n.yview)
        self.tree_n.configure(yscroll=scroll_n.set)
        self.tree_n.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_n.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree_n.bind('<Double-1>', self._toggle_new)

    def _toggle_mismatch(self, event):
        item = self.tree_m.identify_row(event.y)
        if item:
            vals = list(self.tree_m.item(item, 'values'))
            vals[0] = "[ ]" if vals[0] == "[X]" else "[X]"
            self.tree_m.item(item, values=vals)

    def _toggle_new(self, event):
        item = self.tree_n.identify_row(event.y)
        if item:
            vals = list(self.tree_n.item(item, 'values'))
            vals[0] = "[ ]" if vals[0] == "[X]" else "[X]"
            self.tree_n.item(item, values=vals)

    def _pick_cred(self):
        p = filedialog.askopenfilename(title="Select Credentials JSON", filetypes=[("JSON Files","*.json")])
        if p: self.cred_path.set(p)

    def _pick_item(self):
        p = filedialog.askopenfilename(title="Select Item Report", filetypes=[("Excel Files","*.xlsx *.xls")])
        if p:
            self.item_path = p
            self.btn_item.configure(text=f"✅ {Path(p).name[:30]}")
            self._check_ready()

    def _check_ready(self):
        if self.item_path:
            self.btn_run.configure(state=tk.NORMAL, bg=_ACCENT_RED)
            self.status_var.set("Ready to run comparison.")
        else:
            self.btn_run.configure(state=tk.DISABLED, bg="#A0AAB5")

    def _run(self):
        url = self.gsheets_url.get().strip()
        if not url:
            messagebox.showerror("Input Error", "Please provide a Google Sheets URL.")
            return
        if not os.path.exists(self.cred_path.get()):
            messagebox.showerror("Auth Error", "Credentials JSON not found. Please provide a valid path.")
            return

        self.btn_run.configure(state=tk.DISABLED, text="⌛ Processing...")
        self.status_var.set("Connecting to Google Sheets...")
        self.root.update()

        threading.Thread(target=self._process, args=(url,), daemon=True).start()

    def _process(self, url):
        try:
            # 1. Connect to GSheets
            self.gsheets_client = GSheetsClient(self.cred_path.get())
            df_master, self.worksheet = self.gsheets_client.get_sheet_data(url)
            
            if df_master.empty:
                raise Exception("Master Sheet is empty or could not be read.")

            # 2. Load Item Report
            self.status_var.set("Loading Item Report...")
            self.root.update_idletasks()
            df_item = pd.read_excel(self.item_path, dtype=str)

            # 3. Compare
            self.status_var.set("Comparing data...")
            self.root.update_idletasks()
            self.mismatches, self.new_models, self.master_cols = compare_dataframes(df_item, df_master)

            # 4. Update UI
            self.root.after(0, self._populate_tables)

        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
            self.root.after(0, lambda: self.status_var.set("Comparison failed."))
        finally:
            self.root.after(0, lambda: self.btn_run.configure(state=tk.NORMAL, text="▶  Run Comparison"))

    def _populate_tables(self):
        for item in self.tree_m.get_children(): self.tree_m.delete(item)
        for item in self.tree_n.get_children(): self.tree_n.delete(item)

        for i, m in enumerate(self.mismatches):
            self.tree_m.insert("", tk.END, iid=f"m_{i}", values=(
                "[ ]", m['Model'], m['Field'], m['Master Value'], m['Item Report Value'], m['Job No']
            ))

        for i, n in enumerate(self.new_models):
            self.tree_n.insert("", tk.END, iid=f"n_{i}", values=(
                "[ ]", n['Model'], n['Job No']
            ))

        self.notebook.tab(self.tab_mismatches, text=f"Mismatches ({len(self.mismatches)})")
        self.notebook.tab(self.tab_new, text=f"New Models ({len(self.new_models)})")
        
        self.status_var.set(f"Found {len(self.mismatches)} mismatches and {len(self.new_models)} new models.")
        
        if self.mismatches or self.new_models:
            self.btn_update.configure(state=tk.NORMAL)

    def _apply_updates(self):
        if not messagebox.askyesno("Confirm Updates", "This will modify the Google Sheet. Are you sure you want to proceed?"):
            return

        self.btn_update.configure(state=tk.DISABLED, text="⌛ Updating...")
        self.status_var.set("Applying updates to Google Sheets...")
        self.root.update()

        threading.Thread(target=self._process_updates, daemon=True).start()

    def _process_updates(self):
        try:
            # Gather mismatch updates
            cell_updates = []
            
            # Map column names to 1-indexed column numbers
            col_map = {col_name: idx + 1 for idx, col_name in enumerate(self.master_cols)}
            
            for item in self.tree_m.get_children():
                vals = self.tree_m.item(item, 'values')
                if vals[0] == "[X]":
                    idx = int(item.split('_')[1])
                    mismatch = self.mismatches[idx]
                    row_idx = mismatch['gsheet_row']
                    col_name = mismatch['gsheet_col_name']
                    
                    if col_name in col_map:
                        cell_updates.append({
                            'row': row_idx,
                            'col': col_map[col_name],
                            'value': mismatch['Item Report Value']
                        })

            # Gather new models to add
            rows_to_add = []
            for item in self.tree_n.get_children():
                vals = self.tree_n.item(item, 'values')
                if vals[0] == "[X]":
                    idx = int(item.split('_')[1])
                    new_model = self.new_models[idx]
                    
                    # Create row matching master columns
                    row_data = []
                    for col in self.master_cols:
                        row_data.append(new_model['Item Data'].get(col, ""))
                    rows_to_add.append(row_data)

            # Apply
            if cell_updates:
                self.gsheets_client.update_cells(self.worksheet, cell_updates)
            
            if rows_to_add:
                self.gsheets_client.append_rows(self.worksheet, rows_to_add)

            msg = f"Successfully updated {len(cell_updates)} cells and added {len(rows_to_add)} new models."
            self.root.after(0, lambda: messagebox.showinfo("Success", msg))
            self.root.after(0, lambda: self.status_var.set("Updates complete!"))
            
            # Optional: Refresh data by triggering a new run
            # self.root.after(0, self._run)

        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Update Error", str(e)))
            self.root.after(0, lambda: self.status_var.set("Update failed."))
        finally:
            self.root.after(0, lambda: self.btn_update.configure(state=tk.NORMAL, text="☁  Apply Selected Updates to Google Sheets"))

if __name__ == "__main__":
    root = tk.Tk()
    app = MismatchCheckerApp(root)
    root.mainloop()
