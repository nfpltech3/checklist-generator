import pdfplumber
import re
from datetime import datetime

def should_ignore(line: str) -> bool:
    """Ignore headers, footers and page numbers that might break multiline descriptions."""
    ignore_patterns = [
        "CERAGON NETWORKS LTD",
        "Page ",
        "Commercial Invoice No.",
        "Bill To:", "Sold To:", "Ship To:",
        "Customer VAT:", "Tel:", "Fax:",
        "Sales Order No.", "Customer P.O. No:", "Project:", "Packing List No.",
        "S.O. Type:", "Shipment Terms:", "Ship Via:", "Credit Terms:", "End User:",
        "Delivery Gross weight:", "Delivery Net weight:",
        "Line PO Model No.", "# Line P/N Net", "# Weight Origin",
        "Shipment Declaration", "Sub Total:", "V.A.T:", "Total Amount", "For Customs",
        "Remarks", "Identifier =", "Covered under Internet Protocol",
        "Company Registration:", "Self-Declaration:", "Manufactured by",
        "Wireless Radio Link", "V.A.T. File No:"
    ]
    
    for pattern in ignore_patterns:
        if pattern.lower() in line.lower():
            return True
            
    # Also ignore completely dashed/underline lines
    if re.match(r'^[-_]+$', line):
        return True
        
    return False

def parse_ceragon_invoice(pdf_path: str):
    """
    Parse Bharti Airtel Ceragon Commercial Invoice PDF.
    Returns: (invoice_no, doc_date_formatted, currency, items_list)
    """
    invoice_no = ""
    doc_date_raw = ""
    currency = "USD"
    items = {}
    last_line_no = None

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            words = page.extract_words()
            text = page.extract_text() or ""
            lines = text.split("\n")
            
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                # Extract Invoice Number
                if not invoice_no and "Commercial Invoice No." in line:
                    match = re.search(r"Commercial Invoice No\.\s+(\S+)", line)
                    if match:
                        invoice_no = match.group(1)
                
                # Extract Date (Usually near Place &Date or Company Registration on last page)
                date_match = re.search(r"(\d{1,2}-[A-Za-z]{3}-\d{4})", line)
                if date_match and not doc_date_raw:
                    doc_date_raw = date_match.group(1)
                    
                # Extract Currency from header 'Unit Price [USD]'
                curr_match = re.search(r"Unit Price\s*\[([A-Z]{3})\]", line)
                if curr_match:
                    currency = curr_match.group(1)
                elif "[USD]" in line and "Unit Price" in line:
                    currency = "USD"
                elif "Total Amount [" in line:
                    c_m = re.search(r"Total Amount \[([A-Z]{3})\]", line)
                    if c_m:
                        currency = c_m.group(1)
                
                # Try to extract a line item row
                # Pattern A (Airtel): Line# PO Model Desc... KG Weight CoO Qty UnitPrice Total  (10 groups)
                item_start_match = re.match(r"^(\d+\.\d+)\s+(\S+)\s+(\S+)\s+(.+?)\s+([A-Z]{2,3})\s+([0-9.,]+)\s+([A-Z]{2})\s+([0-9.,]+)\s+([0-9.,]+)\s+([0-9.,]+)$", line)
                if item_start_match:
                    line_no = item_start_match.group(1)
                    part_no = item_start_match.group(2)
                    desc = item_start_match.group(4).strip()
                    coo = item_start_match.group(7).strip()
                    qty = item_start_match.group(8).replace(',', '')
                    unit_price = item_start_match.group(9).replace(',', '')
                else:
                    # Pattern B (Hexacom): Line# PO Model Desc... CoO Qty UnitPrice Total  (8 groups, no KG weight)
                    item_start_match = re.match(r"^(\d+\.\d+)\s+(\S+)\s+(\S+)\s+(.+?)\s+([A-Z]{2})\s+([0-9.,]+)\s+([0-9.,]+)\s+([0-9.,]+)$", line)
                    if item_start_match:
                        line_no = item_start_match.group(1)
                        part_no = item_start_match.group(2)
                        desc = item_start_match.group(4).strip()
                        coo = item_start_match.group(5).strip()
                        qty = item_start_match.group(6).replace(',', '')
                        unit_price = item_start_match.group(7).replace(',', '')
                    else:
                        item_start_match = None
                        
                if item_start_match:
                    # Find coordinates for line_no to search below it for wrapped model suffixes
                    item_top = None
                    for w in words:
                        if w['text'] == line_no and abs(w['x0'] - 41.75) < 5:
                            item_top = w['top']
                            break
                            
                    suffix = ""
                    if item_top is not None:
                        for w in words:
                            # Search for a word on the next line (top is ~11-12pt lower) aligned with model column (x0=107)
                            if (item_top + 5) < w['top'] < (item_top + 16) and abs(w['x0'] - 107.0) < 5:
                                suffix = w['text']
                                break
                    if suffix:
                        part_no += suffix
                    
                    if line_no not in items:
                        items[line_no] = {
                            "line_no": line_no,
                            "part_no": part_no,
                            "description": desc,
                            "quantity": qty,
                            "unit_price": unit_price,
                            "coo": coo
                        }
                    last_line_no = line_no
                    continue
                
                # Append subsequent line description if tracking an item
                if last_line_no in items:
                    if not should_ignore(line):
                        existing_desc = items[last_line_no]["description"]
                        if line not in existing_desc:
                            items[last_line_no]["description"] = f"{existing_desc} {line}".strip()

    # Format date
    doc_date_formatted = ""
    if doc_date_raw:
        try:
            parsed_date = datetime.strptime(doc_date_raw, "%d-%b-%Y")
            doc_date_formatted = parsed_date.strftime("%d-%m-%Y")
        except ValueError:
            doc_date_formatted = doc_date_raw

    # Sort items numerically by Line Number
    sorted_items = []
    try:
        sorted_keys = sorted(items.keys(), key=lambda x: [float(i) for i in x.split('.')])
    except Exception:
        sorted_keys = sorted(items.keys())
        
    for k in sorted_keys:
        sorted_items.append(items[k])
        
    return invoice_no, doc_date_formatted, currency, sorted_items
