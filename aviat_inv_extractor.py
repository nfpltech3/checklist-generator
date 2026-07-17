"""
Aviat Invoice and Packing List Extractor
Extracts item details from Commercial Invoice & Packing List PDFs.
Formats descriptions to uppercase, trims HS code decimals, matches weights, and exports to Excel.

Developed for Nagarkot Forwarders Pvt Ltd.
"""

import os
import re
import sys
import datetime
import logging
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import pdfplumber
from PIL import Image, ImageTk

# ---------------------------------------------------------------------------
# Logging Configuration
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Design System & Colors (Nagarkot Corporate Colors)
# ---------------------------------------------------------------------------
BRAND_DARK = "#1B2A4A"      # Deep Navy
BRAND_MID = "#2C4370"       # Mid-blue
BRAND_ACCENT = "#3B82F6"    # Bright blue
BRAND_LIGHT = "#E8EDF5"     # Panel background
BRAND_WHITE = "#FFFFFF"
BRAND_GREEN = "#10B981"     # Success
BRAND_RED = "#EF4444"       # Mismatch / error
BRAND_TEXT = "#1E293B"      # Body text
BRAND_TEXT_LIGHT = "#64748B"

# Keywords to ignore in PDF item descriptions
IGNORE_KEYWORDS = [
    "COMMERCIAL INVOICE", "PACKING LIST", "Seller Shipped From", "Bill To Ship To",
    "LINE # PART NO.", "UNIT", "PRICE", "EXTENSION", "SERIAL", "PLT GROSS", "NUMBER W.",
    "AVIAT NETWORKS", "AVIAT d.o.o.", "Gasper Jezersek", "GSTIN -", "PAYMENT TERMS",
    "FREIGHT TERMS", "TRACKING NO", "cargo-partner", "Harsh Bhardwaj", "Harsh.Bhardwaj",
    "SHIPPED VIA", "SIGNATURE OF SHIPPER", "These goods are essential part", "No payment, value mentioned",
    "DIMENSIONS PACKAGING", "TOTAL:", "188x78x66cm", "127x82x68cm", "80x55x38cm", "73x55x52cm", "120x80x55cm",
    "VALUE FREIGHT INSURANCE"
]

# ---------------------------------------------------------------------------
# Parsing Helpers
# ---------------------------------------------------------------------------
def should_ignore(line: str) -> bool:
    line_upper = line.upper()
    for kw in IGNORE_KEYWORDS:
        if kw.upper() in line_upper:
            return True
    return False

def format_date_helper(date_str: str):
    """Parse a date string and return both short (d-b-y) and long (d-b-Y) formats."""
    if not date_str:
        return "", ""
    for fmt in ("%d-%b-%y", "%d-%b-%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            dt = datetime.datetime.strptime(date_str, fmt)
            day = dt.day
            month = dt.strftime("%b")
            year_short = dt.strftime("%y")
            year_long = dt.strftime("%Y")
            short_date = f"{day}-{month}-{year_short}"
            long_date = f"{day}-{month}-{year_long}"
            return short_date, long_date
        except ValueError:
            continue
    # Clean up string if parsing fails
    clean_date = re.sub(r"\s+", " ", date_str).strip()
    return clean_date, clean_date

def trim_hs_code(hs_code: str) -> str:
    """Remove decimal dots from HS Code."""
    if not hs_code:
        return ""
    return hs_code.replace(".", "").strip()

def clean_description(desc: str) -> str:
    """Format description to all uppercase and remove artifacts."""
    if not desc:
        return ""
    desc_cleaned = desc.replace("\t", " ").replace("(cid:9)", " ").replace("(cid:10)", " ")
    desc_cleaned = re.sub(r"\s+", " ", desc_cleaned).strip()
    return desc_cleaned.upper()

# ---------------------------------------------------------------------------
# Core Extraction Logic
# ---------------------------------------------------------------------------
def parse_commercial_invoice(pdf_path: str):
    """Parse Commercial Invoice PDF to retrieve invoice level details and line items."""
    items = {}
    last_line_no = None
    invoice_no = ""
    doc_date_raw = ""

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            lines = text.split("\n")
            
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                # Extract Invoice Number and Date from first page
                if page_idx == 0 and not invoice_no:
                    # Look for standard CI invoice pattern (e.g. CI-015-26)
                    inv_match = re.search(r"(CI-\d+-\d+)", line)
                    if inv_match:
                        invoice_no = inv_match.group(1)
                    date_match = re.search(r"(\d+-[A-Za-z]+-\d+)", line)
                    if date_match:
                        doc_date_raw = date_match.group(1)

                # Check if it starts with Line Number (e.g. 15.1)
                item_start_match = re.match(r"^(\d+\.\d+)\s+(.+)$", line)
                if item_start_match and len(item_start_match.group(1).split('.')[0]) <= 3:
                    line_no = item_start_match.group(1)
                    rest = item_start_match.group(2).strip()
                    rest = rest.replace("(cid:9)", "\t")
                    
                    # Look for Qty, UnitPrice, Extension at end of row
                    end_match = re.search(r"\s+(\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)$", rest)
                    if end_match:
                        qty = int(end_match.group(1))
                        unit_price = float(end_match.group(2).replace(",", ""))
                        extension = float(end_match.group(3).replace(",", ""))
                        middle = rest[:end_match.start()].strip()
                    else:
                        qty = None
                        unit_price = None
                        extension = None
                        middle = rest
                    
                    # Split middle into Part No and Description
                    if "\t" in middle:
                        parts = middle.split("\t")
                        part_no = parts[0].strip()
                        desc = " ".join(parts[1:]).strip()
                    else:
                        tokens = middle.split()
                        if tokens:
                            part_no = tokens[0]
                            desc = " ".join(tokens[1:])
                        else:
                            part_no = ""
                            desc = ""
                    
                    # Combine split WGA part numbers e.g. WGA 180230-UBR220(A)-SMA
                    if part_no == "WGA" and desc:
                        desc_tokens = desc.split()
                        if desc_tokens and desc_tokens[0].endswith("-SMA"):
                            part_no = "WGA " + desc_tokens[0]
                            desc = " ".join(desc_tokens[1:])
                            
                    # Fix for item 15.38 where part no contains space: "26350-L350 UG-387/U-AC"
                    if part_no == "26350-L350" and desc.startswith("UG-387/U-AC"):
                        desc_tokens = desc.split()
                        if len(desc_tokens) > 1 and desc_tokens[1] == "UG-387/U-AC":
                            part_no = "26350-L350 " + desc_tokens[0] + " " + desc_tokens[1]
                            desc = " ".join(desc_tokens[2:])
                        else:
                            part_no = "26350-L350 " + desc_tokens[0]
                            desc = " ".join(desc_tokens[1:])
                    
                    desc = clean_description(desc)
                    
                    if line_no not in items:
                        items[line_no] = {
                            "line_no": line_no,
                            "part_no": part_no,
                            "description": desc,
                            "quantity": qty,
                            "unit_price": unit_price,
                            "extension": extension,
                            "coo": "",
                            "ct": ""
                        }
                    else:
                        if part_no and not items[line_no]["part_no"]:
                            items[line_no]["part_no"] = part_no
                        if desc:
                            existing = items[line_no]["description"]
                            if desc not in existing:
                                items[line_no]["description"] = f"{existing} {desc}".strip()
                        if qty is not None:
                            items[line_no]["quantity"] = qty
                        if unit_price is not None:
                            items[line_no]["unit_price"] = unit_price
                        if extension is not None:
                            items[line_no]["extension"] = extension
                            
                    last_line_no = line_no
                    continue
                
                # Check for ECCN / CoO / CT row
                if "ECCN :" in line or "CoO :" in line:
                    if last_line_no in items:
                        coo_match = re.search(r"CoO\s*:\s*([A-Za-z]+)", line)
                        if coo_match:
                            items[last_line_no]["coo"] = coo_match.group(1).upper()
                        
                        ct_match = re.search(r"CT\s*:\s*([\d.]+)", line)
                        if ct_match:
                            items[last_line_no]["ct"] = ct_match.group(1)
                    
                    # Reset last line tracker to avoid matching foreign footer text
                    last_line_no = None
                    continue
                
                # Append subsequent line description if tracking an item
                if last_line_no in items:
                    if not should_ignore(line):
                        line_cleaned = clean_description(line)
                        existing_desc = items[last_line_no]["description"]
                        if line_cleaned not in existing_desc:
                            items[last_line_no]["description"] = f"{existing_desc} {line_cleaned}".strip()

    # Sort items numerically by Line Number
    sorted_items = []
    try:
        sorted_keys = sorted(items.keys(), key=lambda x: [float(i) for i in x.split('.')])
    except Exception:
        sorted_keys = sorted(items.keys())
        
    for k in sorted_keys:
        sorted_items.append(items[k])
        
    return invoice_no, doc_date_raw, sorted_items

def parse_packing_list(pdf_path: str):
    """Parse Packing List PDF to retrieve Net Weights associated with each Line Number."""
    items = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            lines = text.split("\n")
            
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                # Check if it starts with Line Number
                item_start_match = re.match(r"^(\d+\.\d+)\s+(.+)$", line)
                if item_start_match and len(item_start_match.group(1).split('.')[0]) <= 3:
                    line_no = item_start_match.group(1)
                    rest = item_start_match.group(2).strip()
                    rest = rest.replace("(cid:9)", "\t")
                    
                    # Match [SerialNo or N/A] [Qty] [NetW] at the end
                    end_match = re.search(r"\s+(\S+)\s+(\d+)\s+([\d,]+\.?\d*)$", rest)
                    if end_match:
                        net_w_raw = end_match.group(3).replace(",", "")
                        try:
                            net_w = float(net_w_raw)
                            # Display integer if decimal is 0
                            if net_w.is_integer():
                                net_w = int(net_w)
                        except ValueError:
                            net_w = net_w_raw
                            
                        items[line_no] = {
                            "net_w": net_w
                        }
                    continue
    return items

# ---------------------------------------------------------------------------
# GUI Application Class
# ---------------------------------------------------------------------------
class AviatExtractorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("AVIAT Invoice & Packing List Extractor — Nagarkot")
        self.root.geometry("1200x750")
        self.root.configure(bg=BRAND_DARK)

        self.root.state("zoomed")
        self.root.minsize(1000, 650)

        # File path variables
        self.invoice_path = tk.StringVar()
        self.packing_path = tk.StringVar()
        
        # Parsed Data store
        self.extracted_invoice_no = ""
        self.extracted_date_short = ""
        self.extracted_date_long = ""
        self.items_data = []

        # Load logo
        self.logo_image = None
        self._load_logo()

        # Build GUI
        self._build_header()
        self._build_body()
        self._build_footer()
        
        # Auto-detect disabled per user request

    def _get_resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def _load_logo(self):
        """Try to load the Nagarkot logo or assets."""
        possible_paths = [
            "nagarkot_logo.png",
            "logo.png",
            "assets/nagarkot_logo.png"
        ]
        
        # Check bundled first (via sys._MEIPASS helper)
        for p in possible_paths:
            resolved = self._get_resource_path(p)
            if os.path.isfile(resolved):
                try:
                    img = Image.open(resolved)
                    img = img.resize((45, 45), Image.Resampling.LANCZOS)
                    self.logo_image = ImageTk.PhotoImage(img)
                    logger.info("Loaded logo from: %s", resolved)
                    return
                except Exception as e:
                    logger.warning("Could not load logo from %s: %s", resolved, e)

        # Fallback to direct paths if not bundled
        for p in possible_paths + ["../Skoda 1702/Skoda export checklist verifier/nagarkot_logo.png"]:
            if os.path.isfile(p):
                try:
                    img = Image.open(p)
                    img = img.resize((45, 45), Image.Resampling.LANCZOS)
                    self.logo_image = ImageTk.PhotoImage(img)
                    logger.info("Loaded logo from fallback path: %s", p)
                    return
                except Exception as e:
                    logger.warning("Could not load logo from %s: %s", p, e)


    def _build_header(self):
        header = tk.Frame(self.root, bg=BRAND_DARK, height=75)
        header.pack(fill=tk.X, side=tk.TOP)
        header.pack_propagate(False)

        # Left Logo
        logo_frame = tk.Frame(header, bg=BRAND_DARK)
        logo_frame.pack(side=tk.LEFT, padx=20, pady=10)
        if self.logo_image:
            lbl_logo = tk.Label(logo_frame, image=self.logo_image, bg=BRAND_DARK)
            lbl_logo.pack()
        else:
            # Elegant text fallback
            lbl_logo_fallback = tk.Label(
                logo_frame, 
                text="NAGARKOT", 
                font=("Segoe UI Semibold", 13), 
                fg=BRAND_ACCENT, 
                bg=BRAND_DARK
            )
            lbl_logo_fallback.pack()

        # Center Title
        title_label = tk.Label(
            header,
            text="AVIAT INVOICE & PACKING LIST EXTRACTOR",
            font=("Segoe UI Semibold", 18),
            fg=BRAND_WHITE,
            bg=BRAND_DARK
        )
        title_label.pack(side=tk.LEFT, expand=True)

        # Right Status
        self.lbl_status = tk.Label(
            header,
            text="● Ready",
            font=("Segoe UI", 11),
            fg=BRAND_GREEN,
            bg=BRAND_DARK
        )
        self.lbl_status.pack(side=tk.RIGHT, padx=25)

    def _build_body(self):
        body = tk.Frame(self.root, bg=BRAND_LIGHT)
        body.pack(fill=tk.BOTH, expand=True)

        # 1. File Selection Card
        card_file = tk.Frame(body, bg=BRAND_WHITE, relief="flat", bd=0)
        card_file.pack(fill=tk.X, padx=20, pady=(20, 10))

        inner_file = tk.Frame(card_file, bg=BRAND_WHITE, padx=20, pady=15)
        inner_file.pack(fill=tk.X)

        # Commercial Invoice Browse
        lbl_ci = tk.Label(
            inner_file, 
            text="📄  Commercial Invoice PDF", 
            font=("Segoe UI Semibold", 10), 
            fg=BRAND_TEXT, 
            bg=BRAND_WHITE
        )
        lbl_ci.grid(row=0, column=0, sticky="w", pady=(0, 5))
        
        entry_ci = tk.Entry(
            inner_file, 
            textvariable=self.invoice_path, 
            font=("Segoe UI", 9), 
            bg=BRAND_LIGHT, 
            fg=BRAND_TEXT, 
            relief="flat", 
            bd=0
        )
        entry_ci.grid(row=1, column=0, sticky="ew", ipady=6, padx=(0, 10))
        
        btn_ci = tk.Button(
            inner_file, 
            text="Browse", 
            command=self._browse_invoice,
            font=("Segoe UI", 9), 
            bg=BRAND_MID, 
            fg=BRAND_WHITE, 
            activebackground=BRAND_ACCENT,
            activeforeground=BRAND_WHITE, 
            relief="flat", 
            cursor="hand2", 
            padx=15
        )
        btn_ci.grid(row=1, column=1)

        # Packing List Browse
        lbl_pl = tk.Label(
            inner_file, 
            text="📦  Packing List PDF", 
            font=("Segoe UI Semibold", 10), 
            fg=BRAND_TEXT, 
            bg=BRAND_WHITE
        )
        lbl_pl.grid(row=0, column=2, sticky="w", pady=(0, 5), padx=(20, 0))
        
        entry_pl = tk.Entry(
            inner_file, 
            textvariable=self.packing_path, 
            font=("Segoe UI", 9), 
            bg=BRAND_LIGHT, 
            fg=BRAND_TEXT, 
            relief="flat", 
            bd=0
        )
        entry_pl.grid(row=1, column=2, sticky="ew", ipady=6, padx=(20, 10))
        
        btn_pl = tk.Button(
            inner_file, 
            text="Browse", 
            command=self._browse_packing,
            font=("Segoe UI", 9), 
            bg=BRAND_MID, 
            fg=BRAND_WHITE, 
            activebackground=BRAND_ACCENT,
            activeforeground=BRAND_WHITE, 
            relief="flat", 
            cursor="hand2", 
            padx=15
        )
        btn_pl.grid(row=1, column=3)

        inner_file.columnconfigure(0, weight=1)
        inner_file.columnconfigure(2, weight=1)

        # Action Buttons Row
        action_bar = tk.Frame(inner_file, bg=BRAND_WHITE)
        action_bar.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(15, 0))

        btn_extract = tk.Button(
            action_bar,
            text="⚡  EXTRACT DATA",
            command=self._extract_data,
            font=("Segoe UI Semibold", 11),
            bg=BRAND_ACCENT,
            fg=BRAND_WHITE,
            activebackground=BRAND_MID,
            activeforeground=BRAND_WHITE,
            relief="flat",
            cursor="hand2",
            padx=25,
            pady=6
        )
        btn_extract.pack(side=tk.LEFT)

        btn_clear = tk.Button(
            action_bar,
            text="🗑️  CLEAR",
            command=self._clear_all,
            font=("Segoe UI Semibold", 11),
            bg=BRAND_RED,
            fg=BRAND_WHITE,
            activebackground=BRAND_DARK,
            activeforeground=BRAND_WHITE,
            relief="flat",
            cursor="hand2",
            padx=25,
            pady=6
        )
        btn_clear.pack(side=tk.LEFT, padx=(15, 0))

        # 2. Invoice Meta Details Info Box
        self.meta_frame = tk.Frame(body, bg=BRAND_LIGHT)
        self.meta_frame.pack(fill=tk.X, padx=20, pady=(10, 5))

        self.lbl_meta_info = tk.Label(
            self.meta_frame,
            text="",
            font=("Segoe UI Semibold", 11),
            fg=BRAND_DARK,
            bg=BRAND_LIGHT,
            anchor="w"
        )
        self.lbl_meta_info.pack(fill=tk.X)

        # 3. Treeview Visual Table
        card_table = tk.Frame(body, bg=BRAND_WHITE, relief="flat", bd=0)
        card_table.pack(fill=tk.BOTH, expand=True, padx=20, pady=(5, 15))

        table_header = tk.Frame(card_table, bg=BRAND_WHITE, padx=15, pady=10)
        table_header.pack(fill=tk.X)
        
        lbl_table_title = tk.Label(
            table_header, 
            text="Line Items Extracted Table", 
            font=("Segoe UI Semibold", 11), 
            fg=BRAND_TEXT, 
            bg=BRAND_WHITE
        )
        lbl_table_title.pack(side=tk.LEFT)

        # Actions in table header
        self.btn_export = tk.Button(
            table_header,
            text="📥  Export to Excel",
            command=self._export_to_excel,
            font=("Segoe UI", 9),
            bg=BRAND_GREEN,
            fg=BRAND_WHITE,
            activebackground=BRAND_MID,
            activeforeground=BRAND_WHITE,
            relief="flat",
            cursor="hand2",
            padx=12,
            pady=3,
            state=tk.DISABLED
        )
        self.btn_export.pack(side=tk.RIGHT, padx=(10, 0))

        self.btn_copy = tk.Button(
            table_header,
            text="📋  Copy to Clipboard",
            command=self._copy_to_clipboard,
            font=("Segoe UI", 9),
            bg=BRAND_MID,
            fg=BRAND_WHITE,
            activebackground=BRAND_ACCENT,
            activeforeground=BRAND_WHITE,
            relief="flat",
            cursor="hand2",
            padx=12,
            pady=3,
            state=tk.DISABLED
        )
        self.btn_copy.pack(side=tk.RIGHT)

        # Styled Treeview
        style = ttk.Style()
        style.theme_use("clam")
        
        # Table styles
        style.configure(
            "Treeview",
            background=BRAND_WHITE,
            foreground=BRAND_TEXT,
            fieldbackground=BRAND_WHITE,
            rowheight=26,
            font=("Segoe UI", 9)
        )
        style.map(
            "Treeview",
            background=[("selected", BRAND_ACCENT)],
            foreground=[("selected", BRAND_WHITE)]
        )
        style.configure(
            "Treeview.Heading",
            background=BRAND_LIGHT,
            foreground=BRAND_TEXT,
            font=("Segoe UI Semibold", 9),
            relief="flat",
            borderwidth=1
        )

        table_container = tk.Frame(card_table, bg=BRAND_WHITE, padx=15, pady=0)
        table_container.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        columns = (
            "part_no", "description", "coo", "ct", "qty", "unit_price", "extension", "net_w"
        )
        
        self.tree = ttk.Treeview(
            table_container, 
            columns=columns, 
            show="headings", 
            selectmode="extended"
        )

        # Set Column Headers & Widths
        headers = {
            "part_no": "Part No",
            "description": "Description & Compliance Info",
            "coo": "CoO",
            "ct": "CT (HS Code)",
            "qty": "Quantity",
            "unit_price": "Unit Price",
            "extension": "Extension",
            "net_w": "Net Weight (kg)"
        }
        
        widths = {
            "part_no": 130,
            "description": 380,
            "coo": 50,
            "ct": 100,
            "qty": 80,
            "unit_price": 90,
            "extension": 90,
            "net_w": 110
        }

        for col, header_text in headers.items():
            self.tree.heading(col, text=header_text, anchor="center")
            self.tree.column(col, width=widths[col], anchor="center" if col != "description" else "w")

        # Scrollbars
        vsb = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        table_container.columnconfigure(0, weight=1)
        table_container.rowconfigure(0, weight=1)

    def _build_footer(self):
        footer = tk.Frame(self.root, bg=BRAND_DARK, height=35)
        footer.pack(fill=tk.X, side=tk.BOTTOM)
        footer.pack_propagate(False)

        lbl_footer = tk.Label(
            footer,
            text="© 2026 Nagarkot Forwarders Pvt Ltd. Rebranded for Shakti database compatibility. Strictly Confidential.",
            font=("Segoe UI", 8),
            fg=BRAND_TEXT_LIGHT,
            bg=BRAND_DARK
        )
        lbl_footer.pack(pady=8)

    # ---------------------------------------------------------------------------
    # GUI Interactive Commands
    # ---------------------------------------------------------------------------
    def _browse_invoice(self):
        file = filedialog.askopenfilename(
            title="Select Commercial Invoice PDF",
            filetypes=[("PDF files", "*.pdf")]
        )
        if file:
            self.invoice_path.set(os.path.abspath(file))
            self._set_status("Selected Commercial Invoice.", BRAND_MID)

    def _browse_packing(self):
        file = filedialog.askopenfilename(
            title="Select Packing List PDF",
            filetypes=[("PDF files", "*.pdf")]
        )
        if file:
            self.packing_path.set(os.path.abspath(file))
            self._set_status("Selected Packing List.", BRAND_MID)

    def _set_status(self, text: str, color: str = BRAND_WHITE):
        self.lbl_status.config(text=f"● {text}", fg=color)

    def _clear_all(self):
        self.invoice_path.set("")
        self.packing_path.set("")
        self.extracted_invoice_no = ""
        self.extracted_date_short = ""
        self.extracted_date_long = ""
        self.items_data = []
        self.lbl_meta_info.config(text="")
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.btn_export.config(state=tk.DISABLED)
        self.btn_copy.config(state=tk.DISABLED)
        self._set_status("Ready", BRAND_GREEN)

    def _extract_data(self):
        ci = self.invoice_path.get()
        pl = self.packing_path.get()

        if not ci or not pl:
            messagebox.showerror(
                "Missing Files", 
                "Please select both the Commercial Invoice and Packing List PDF files before extracting."
            )
            return

        if not os.path.exists(ci) or not os.path.exists(pl):
            messagebox.showerror("Error", "One or both of the selected files do not exist.")
            return

        # Check for swapped files by looking at first page content
        def check_pdf_type(path):
            try:
                with pdfplumber.open(path) as pdf:
                    text = pdf.pages[0].extract_text() or ""
                    if "PACKING LIST" in text.upper():
                        return "PL"
                    if "COMMERCIAL INVOICE" in text.upper():
                        return "CI"
            except Exception:
                pass
            return "UNKNOWN"

        ci_type = check_pdf_type(ci)
        pl_type = check_pdf_type(pl)

        if ci_type == "PL" and (pl_type == "CI" or "CI" in os.path.basename(pl)):
            ci, pl = pl, ci
            self.invoice_path.set(ci)
            self.packing_path.set(pl)
            logger.info("Auto-swapped CI and PL inputs based on file contents.")

        self._set_status("Parsing PDFs...", BRAND_ACCENT)
        self.root.update_idletasks()

        try:
            # Parse Commercial Invoice
            inv_no, date_raw, ci_items = parse_commercial_invoice(ci)
            
            # Parse Packing List
            pl_weights = parse_packing_list(pl)

            if not ci_items:
                messagebox.showwarning(
                    "Parsing Warning", 
                    "No items found in the Commercial Invoice. Please check if the file matches the expected template."
                )
                self._set_status("Extraction failed", BRAND_RED)
                return

            # Format Dates
            date_short, date_long = format_date_helper(date_raw)
            self.extracted_invoice_no = inv_no or "UNKNOWN"
            self.extracted_date_short = date_short
            self.extracted_date_long = date_long

            # Merge Weights & Format Fields
            self.items_data = []
            for item in ci_items:
                line_no = item["line_no"]
                pl_data = pl_weights.get(line_no, {})
                net_w = pl_data.get("net_w", "N/A")
                
                # Format fields as requested
                # uppercase description
                desc_upper = (item["description"] or "").upper()
                
                # Trim CT (HS Code) decimals
                ct_trimmed = trim_hs_code(item["ct"])

                self.items_data.append({
                    "line_no": line_no,
                    "part_no": item["part_no"] or "N/A",
                    "description": desc_upper,
                    "coo": item["coo"] or "N/A",
                    "ct": ct_trimmed or "N/A",
                    "qty": item["quantity"] if item["quantity"] is not None else 0,
                    "unit_price": item["unit_price"] if item["unit_price"] is not None else 0.0,
                    "extension": item["extension"] if item["extension"] is not None else 0.0,
                    "net_w": net_w
                })

            # Update GUI Elements
            # Display meta info
            self.lbl_meta_info.config(
                text=f"Commercial Invoice No: {self.extracted_invoice_no}   |   "
                     f"Date (Long): {self.extracted_date_long}"
            )

            # Clear Table
            for row in self.tree.get_children():
                self.tree.delete(row)

            # Populate Table
            for item in self.items_data:
                self.tree.insert("", "end", values=(
                    item["part_no"],
                    item["description"],
                    item["coo"],
                    item["ct"],
                    item["qty"],
                    f"{item['unit_price']:.2f}" if isinstance(item['unit_price'], (int, float)) else item['unit_price'],
                    f"{item['extension']:.2f}" if isinstance(item['extension'], (int, float)) else item['extension'],
                    item["net_w"]
                ))

            # Enable Actions
            self.btn_export.config(state=tk.NORMAL)
            self.btn_copy.config(state=tk.NORMAL)

            self._set_status(f"Successfully extracted {len(self.items_data)} items.", BRAND_GREEN)
            logger.info("Parsed %d items for invoice %s", len(self.items_data), self.extracted_invoice_no)

        except Exception as e:
            logger.exception("Error extracting data")
            messagebox.showerror("Error During Extraction", f"An error occurred while parsing the PDFs:\n{str(e)}")
            self._set_status("Extraction error", BRAND_RED)

    def _export_to_excel(self):
        if not self.items_data:
            return

        # Prepare path suggestions
        date_str = self.extracted_date_long if self.extracted_date_long else "date"
        suggested_name = f"{self.extracted_invoice_no}_{date_str}.xlsx"
        file_path = filedialog.asksaveasfilename(
            title="Save Extracted Data to Excel",
            initialfile=suggested_name,
            filetypes=[("Excel files", "*.xlsx")],
            defaultextension=".xlsx"
        )

        if not file_path:
            return

        try:
            # Build DataFrame
            rows = []
            for item in self.items_data:
                rows.append({
                    "Commercial Invoice No": self.extracted_invoice_no,
                    "Document Date (Long)": self.extracted_date_long,
                    "Part No": item["part_no"],
                    "Description and Compliance Info": item["description"],
                    "CoO": item["coo"],
                    "CT (HS Code)": item["ct"],
                    "Quantity": item["qty"],
                    "Unit Price": item["unit_price"],
                    "Extension": item["extension"],
                    "Net Weight (kg)": item["net_w"]
                })

            df = pd.DataFrame(rows)

            # Export using openpyxl for clean styling
            with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Invoice Extracted Data")
                
                # Apply formatting
                workbook = writer.book
                worksheet = writer.sheets["Invoice Extracted Data"]
                
                # Auto-fit columns
                for col in worksheet.columns:
                    max_len = 0
                    col_letter = col[0].column_letter
                    for cell in col:
                        if cell.value:
                            max_len = max(max_len, len(str(cell.value)))
                    # cap max width at 60 for descriptions to keep readable
                    worksheet.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 60)

            messagebox.showinfo("Export Success", f"Extracted data successfully exported to:\n{file_path}")
            self._set_status("Excel exported successfully.", BRAND_GREEN)
            logger.info("Exported data to Excel file: %s", file_path)

        except Exception as e:
            logger.exception("Excel export error")
            messagebox.showerror("Export Failed", f"An error occurred while saving the Excel file:\n{str(e)}")

    def _copy_to_clipboard(self):
        if not self.items_data:
            return

        try:
            # Create a tab-separated string representation of the data (similar to Excel copy-paste)
            header = "Commercial Invoice No\tDocument Date (Long)\tPart No\tDescription\tCoO\tCT\tQuantity\tUnit Price\tExtension\tNet Weight (kg)\n"
            lines = [header]
            for item in self.items_data:
                line = (
                    f"{self.extracted_invoice_no}\t"
                    f"{self.extracted_date_long}\t"
                    f"{item['part_no']}\t"
                    f"{item['description']}\t"
                    f"{item['coo']}\t"
                    f"{item['ct']}\t"
                    f"{item['qty']}\t"
                    f"{item['unit_price']}\t"
                    f"{item['extension']}\t"
                    f"{item['net_w']}\n"
                )
                lines.append(line)

            clipboard_text = "".join(lines)
            
            # Clear clipboard and set text
            self.root.clipboard_clear()
            self.root.clipboard_append(clipboard_text)
            self.root.update()  # keep clipboard contents after app closes

            messagebox.showinfo("Clipboard", f"Copied {len(self.items_data)} items to clipboard successfully!")
            self._set_status("Copied to clipboard.", BRAND_GREEN)

        except Exception as e:
            messagebox.showerror("Clipboard Copy Failed", f"Failed to copy data to clipboard:\n{str(e)}")

# ---------------------------------------------------------------------------
# Application Entry Point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    # Force system scaling awareness on Windows
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = tk.Tk()
    app = AviatExtractorApp(root)
    root.mainloop()