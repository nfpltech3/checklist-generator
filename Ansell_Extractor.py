import os
import sys  # Added for PyInstaller
import csv
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None

# ---------- RESOURCE PATH FUNCTION ----------

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# ---------- CORE EXTRACTION LOGIC ----------

def extract_invoice_data(pdf_path):
    """
    Extracts line-item data from one Ansell commercial invoice PDF.
    Returns a list of dicts (one per line item).
    """
    with pdfplumber.open(pdf_path) as pdf:
        all_tables = []
        for page in pdf.pages:
            all_tables.extend(page.extract_tables())

    # 1) Find the invoice header table (contains "Commercial Invoice")
    inv_table = None
    for tbl in all_tables:
        if len(tbl) >= 3 and tbl[0] and tbl[0][0] and "Commercial Invoice" in tbl[0][0]:
            inv_table = tbl
            break

    if inv_table is None:
        raise ValueError(f"Invoice header table not found in {pdf_path}")

    invoice_no = inv_table[2][0].strip() if inv_table[2][0] else ""
    invoice_date = ""
    if len(inv_table[2]) > 1 and inv_table[2][1]:
        invoice_date = inv_table[2][1].strip()

    # 2) Find all line-item tables (header has "Product Description")
    item_tables = []
    for tbl in all_tables:
        if not tbl or not tbl[0]:
            continue
        header_row = tbl[0]
        if any(cell and "Product Description" in cell for cell in header_row):
            item_tables.append(tbl)

    if not item_tables:
        raise ValueError(f"No item tables found in {pdf_path}")

    def find_col(keyword, rows):
        kw = keyword.lower()
        for row in rows:
            if not row:
                continue
            for idx, cell in enumerate(row):
                if cell and kw in cell.lower():
                    return idx
        return None

    records = []

    for tbl in item_tables:
        hdr = tbl[0]
        sub = tbl[1] if len(tbl) > 1 else [None] * len(hdr)

        idx_prod_code = find_col("product", [hdr])
        if idx_prod_code is not None and hdr[idx_prod_code] and "Description" in hdr[idx_prod_code]:
            # if we accidentally hit "Product Description", look again for "Product Code"
            for idx, cell in enumerate(hdr):
                if cell and "Product" in cell and "Code" in cell:
                    idx_prod_code = idx
                    break

        idx_desc = find_col("Product Description", [hdr])
        idx_country = find_col("Country", [hdr])
        idx_hs = find_col("HTS/HS", [hdr])
        idx_lot = find_col("Lot No", [hdr])
        idx_shipcase = find_col("Case", [hdr])

        idx_qty = None
        idx_uom = None
        idx_price = None
        for idx, cell in enumerate(sub):
            if cell == "Qty" and idx_qty is None:
                idx_qty = idx
            elif cell == "UOM" and idx_qty is not None and idx_uom is None:
                idx_uom = idx
            elif cell == "Price" and idx_price is None:
                idx_price = idx

        idx_currency = find_col("Cur", [hdr])
        idx_value = find_col("Value", [hdr])

        def get(row, idx):
            if idx is None:
                return ""
            if idx >= len(row):
                return ""
            return (row[idx] or "").strip()

        for row in tbl[2:]:
            if not row:
                continue
            first_cell = (row[0] or "").strip() if len(row) > 0 else ""
            if first_cell.startswith("Total"):
                break
            if first_cell == "":
                continue

            lot_raw = get(row, idx_lot)
            lot_no = lot_raw.splitlines()[0].strip() if lot_raw else ""

            record = {
                "SourceFile": os.path.basename(pdf_path),
                "Invoice_No": invoice_no,
                "Invoice_Date": invoice_date,
                "Product_Code": get(row, idx_prod_code),
                "Product_Description": get(row, idx_desc),
                "Country_of_Origin": get(row, idx_country),
                "HS_Code": get(row, idx_hs),
                "Lot_No": lot_no,
                "Ship_Case": get(row, idx_shipcase),
                "Qty": get(row, idx_qty),
                "UOM": get(row, idx_uom),
                "Price": get(row, idx_price),
                "Currency": get(row, idx_currency),
                "Value": get(row, idx_value),
            }
            records.append(record)

    return records


def write_combined_csv(output_path, all_records):
    fieldnames = [
        "SourceFile",
        "Invoice_No",
        "Invoice_Date",
        "Product_Code",
        "Product_Description",
        "Country_of_Origin",
        "HS_Code",
        "Lot_No",
        "Ship_Case",
        "Qty",
        "UOM",
        "Price",
        "Currency",
        "Value",
    ]

    with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(all_records)


# ---------- NAGARKOT GUI IMPLEMENTATION ----------

class InvoiceExtractorGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Ansell Invoice Extractor")
        self.root.geometry("900x650")
        self.root.state('zoomed')  # Open maximized

        # Colors and Styles
        self.bg_color = "#ffffff"
        self.brand_color = "#0056b3"
        self.root.configure(bg=self.bg_color)

        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Configure TFrame
        self.style.configure("TFrame", background=self.bg_color)
        
        # Configure TLabel
        self.style.configure("TLabel", background=self.bg_color, font=("Segoe UI", 10))
        self.style.configure("Header.TLabel", font=("Helvetica", 18, "bold"), foreground=self.brand_color, background=self.bg_color)
        self.style.configure("Subtitle.TLabel", font=("Segoe UI", 11), foreground="gray", background=self.bg_color)
        self.style.configure("Footer.TLabel", font=("Segoe UI", 9), foreground="#555555", background=self.bg_color)
        
        # Configure TButton (Primary & Secondary)
        self.style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), background=self.brand_color, foreground="white", borderwidth=0, focuscolor=self.brand_color)
        self.style.map("Primary.TButton", background=[('active', '#004494')])
        
        self.style.configure("Secondary.TButton", font=("Segoe UI", 10), background="#f0f0f0", foreground="#333333", borderwidth=1)
        self.style.map("Secondary.TButton", background=[('active', '#e0e0e0')])

        # Configure TLabelframe
        self.style.configure("TLabelframe", background=self.bg_color)
        self.style.configure("TLabelframe.Label", background=self.bg_color, foreground=self.brand_color, font=("Segoe UI", 10, "bold"))

        # Configure Treeview
        self.style.configure("Treeview", font=("Segoe UI", 9), rowheight=25)
        self.style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), foreground=self.brand_color)

        self.setup_ui()
        self.selected_files = []

    def setup_ui(self):
        # --- HEADER ---
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill="x", pady=20, padx=20)
        
        # Grid layout for header: 3 columns (Logo, Title, Spacer)
        header_frame.columnconfigure(0, weight=0) # Logo
        header_frame.columnconfigure(1, weight=1) # Title (Center)
        header_frame.columnconfigure(2, weight=0) # Spacer (Right) or extra content

        # 1. Logo (Left)
        try:
            if Image and ImageTk:
                logo_path = resource_path("Nagarkot Logo.png")
                if os.path.exists(logo_path):
                    pil_img = Image.open(logo_path)
                    # Resize to height 20px, maintaining aspect ratio
                    h_percent = (20 / float(pil_img.size[1]))
                    w_size = int((float(pil_img.size[0]) * float(h_percent)))
                    pil_img = pil_img.resize((w_size, 20), Image.Resampling.LANCZOS)
                    self.logo_img = ImageTk.PhotoImage(pil_img)
                    logo_lbl = ttk.Label(header_frame, image=self.logo_img)
                    logo_lbl.grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 20))
                else:
                    print("Warning: Nagarkot Logo.png not found.")
                    ttk.Label(header_frame, text="[LOGO]", foreground="gray").grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 20))
            else:
                 ttk.Label(header_frame, text="[PIL Missing]", foreground="red").grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 20))
        except Exception as e:
            print(f"Error loading logo: {e}")
            ttk.Label(header_frame, text="[LOGO ERROR]", foreground="red").grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 20))

        # 2. Title (Center)
        title_lbl = ttk.Label(header_frame, text="Ansell Invoice Extractor", style="Header.TLabel")
        title_lbl.grid(row=0, column=1, sticky="n")  # Center aligned by column weight
        
        subtitle_lbl = ttk.Label(header_frame, text="Commercial Invoice Data Processing Tool", style="Subtitle.TLabel")
        subtitle_lbl.grid(row=1, column=1, sticky="n")

        # --- MAIN CONTENT ---
        content_frame = ttk.Frame(self.root, padding="20 10 20 10")
        content_frame.pack(fill="both", expand=True)

        # File Selection Area
        file_frame = ttk.LabelFrame(content_frame, text="File Selection", padding="15")
        file_frame.pack(fill="x", pady=(0, 15))

        btn_container = ttk.Frame(file_frame)
        btn_container.pack(fill="x")
        
        self.btn_select = ttk.Button(btn_container, text="Select PDFs", command=self.select_files, style="Secondary.TButton")
        self.btn_select.pack(side="left", padx=(0, 10))
        
        self.btn_clear = ttk.Button(btn_container, text="Clear List", command=self.clear_files, style="Secondary.TButton")
        self.btn_clear.pack(side="left")

        # Data Preview / Log Area
        preview_frame = ttk.LabelFrame(content_frame, text="Data Preview / Processing Queue", padding="15")
        preview_frame.pack(fill="both", expand=True)

        # Treeview
        cols = ("File Name", "Status", "Message")
        self.tree = ttk.Treeview(preview_frame, columns=cols, show="headings", selectmode="extended")
        
        self.tree.heading("File Name", text="File Name")
        self.tree.heading("Status", text="Status")
        self.tree.heading("Message", text="Details")
        
        self.tree.column("File Name", width=300, anchor="w")
        self.tree.column("Status", width=100, anchor="center")
        self.tree.column("Message", width=400, anchor="w")
        
        scrollbar = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # --- FOOTER ---
        footer_frame = ttk.Frame(self.root, padding="10")
        footer_frame.pack(side="bottom", fill="x")
        
        # Copyright (Left)
        copyright_lbl = ttk.Label(footer_frame, text="© Nagarkot Forwarders Pvt Ltd", style="Footer.TLabel")
        copyright_lbl.pack(side="left", anchor="s")

        # Action Button (Right)
        self.btn_run = ttk.Button(footer_frame, text="Extract & Generate CSV", command=self.run_extraction, style="Primary.TButton")
        self.btn_run.pack(side="right")

    def select_files(self):
        filetypes = [("PDF files", "*.pdf")]
        files = filedialog.askopenfilenames(
            title="Select Ansell Commercial Invoice PDFs",
            filetypes=filetypes
        )
        if files:
            for f in files:
                if f not in self.selected_files:
                    self.selected_files.append(f)
                    self.tree.insert("", "end", values=(os.path.basename(f), "Ready", "Waiting to process"))

    def clear_files(self):
        self.selected_files = []
        for item in self.tree.get_children():
            self.tree.delete(item)

    def run_extraction(self):
        if not self.selected_files:
            messagebox.showwarning("No files selected", "Please select PDFs first.")
            return

        # Disable buttons during processing
        self.btn_run.state(["disabled"])
        self.btn_select.state(["disabled"])
        self.btn_clear.state(["disabled"])
        self.root.update_idletasks()

        combined_records = []
        errors = []

        # Reset status in tree
        for item in self.tree.get_children():
            self.tree.set(item, "Status", "Pending")
            self.tree.set(item, "Message", "...")

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        
        for i, pdf_path in enumerate(self.selected_files):
            # Find item in tree
            tree_item = self.tree.get_children()[i]
            self.tree.set(tree_item, "Status", "Processing")
            self.tree.selection_set(tree_item)
            self.tree.see(tree_item)
            self.root.update_idletasks()

            try:
                recs = extract_invoice_data(pdf_path)
                combined_records.extend(recs)
                self.tree.set(tree_item, "Status", "Success")
                self.tree.set(tree_item, "Message", f"Extracted {len(recs)} lines")
            except Exception as e:
                errors.append(f"{os.path.basename(pdf_path)} → {e}")
                self.tree.set(tree_item, "Status", "Error")
                self.tree.set(tree_item, "Message", str(e))

        if not combined_records and not errors:
            messagebox.showinfo("Result", "No data extracted.")
            self._reset_buttons()
            return

        if combined_records:
            folder = os.path.dirname(self.selected_files[0])
            out_path = os.path.join(folder, f"Combined_Extracted_{timestamp}.csv")
            try:
                write_combined_csv(out_path, combined_records)
                messagebox.showinfo("Completed", f"Combined CSV created successfully:\n\n{out_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to write CSV:\n{e}")
        
        if errors:
            msg = "Some files had errors. Check the status list or logs."
            messagebox.showwarning("Partial Errors", msg)

        self._reset_buttons()

    def _reset_buttons(self):
        self.btn_run.state(["!disabled"])
        self.btn_select.state(["!disabled"])
        self.btn_clear.state(["!disabled"])

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = InvoiceExtractorGUI()
    app.run()