import pdfplumber
import re
from datetime import datetime


def parse_eci_invoice(pdf_path: str):
    """
    Parse Bharti Airtel ECI Commercial Invoice PDF.
    Returns: (invoice_no, doc_date_formatted, currency, coo, items_list)
    """
    invoice_no = ""
    doc_date_formatted = ""
    currency = "USD"
    coo = ""
    items = []

    with pdfplumber.open(pdf_path) as pdf:
        full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        
        # 1. Invoice No — against "Invoice / Delivery No:"
        inv_match = re.search(r"Invoice\s*/?\s*Delivery\s+No:\s*(\S+)", full_text, re.IGNORECASE)
        if inv_match:
            invoice_no = inv_match.group(1)
        
        # 2. Date — against "Date:"
        date_match = re.search(r"Date:\s*(\S+)", full_text, re.IGNORECASE)
        if date_match:
            raw_date = date_match.group(1)
            # Try DD-MMM-YY or DD-MMM-YYYY
            for fmt in ("%d-%b-%y", "%d-%b-%Y"):
                try:
                    parsed_date = datetime.strptime(raw_date, fmt)
                    doc_date_formatted = parsed_date.strftime("%d-%m-%Y")
                    break
                except ValueError:
                    continue
            if not doc_date_formatted:
                doc_date_formatted = raw_date
        
        # 3. Currency — against "Currency:"
        curr_match = re.search(r"Currency:\s*(\S+)", full_text, re.IGNORECASE)
        if curr_match:
            currency = curr_match.group(1)
        
        # 4. Country of Origin — against "COO:"
        coo_match = re.search(r"COO:\s*(\S+)", full_text, re.IGNORECASE)
        if coo_match:
            coo = coo_match.group(1).upper()
        
        # 5. Line Items — find table across all pages
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines = text.split("\n")
            
            # Find the header row that contains "Item Code" or "item code"
            header_idx = None
            for idx, line in enumerate(lines):
                if re.search(r"Item\s+Code", line, re.IGNORECASE):
                    header_idx = idx
                    break
            
            if header_idx is None:
                continue
            
            # Parse column positions from the header
            header_line = lines[header_idx]
            
            # Process lines below the header
            for line in lines[header_idx + 1:]:
                line = line.strip()
                if not line:
                    continue
                
                # Stop at subtotal/grand total/footer lines
                if re.search(r"(SUBTOTAL|GRAND TOTAL|NO RETURNS|Signature|Prepaid|ECI Telecom)", line, re.IGNORECASE):
                    break
                
                # Try to match the data row pattern:
                # LineNo ProductName ItemCode UnitPrice OrderedQty ShippedQty UOM Total
                # Example: 1 B0HSTHXP2 - NPT-1100H DC Shelf Assembled or Fiber Optic Product X44994H 2,071.46000 30 30 EA 62,143.80
                row_match = re.match(
                    r"^(\d+)\s+"          # Line No
                    r"(.+?)\s+"           # Product Name (greedy but minimal to catch Item Code)
                    r"([A-Z0-9][\w-]*)\s+"  # Item Code (Model)
                    r"([0-9,]+\.\d+)\s+"  # Unit Price
                    r"(\d+)\s+"           # Ordered Qty
                    r"(\d+)\s+"           # Shipped Qty
                    r"([A-Za-z]{2,3})\s+" # UOM
                    r"([0-9,]+\.\d+)$",   # Total
                    line
                )
                if row_match:
                    product_name = row_match.group(2).strip()
                    item_code = row_match.group(3).strip()
                    unit_price = row_match.group(4).replace(",", "")
                    shipped_qty = row_match.group(6)  # group 6 = Shipped Qty
                    uom = row_match.group(7)
                    
                    items.append({
                        "part_no": item_code,
                        "description": product_name,
                        "quantity": shipped_qty,
                        "unit": uom,
                        "rate": unit_price,
                        "coo": coo,
                    })
    
    return invoice_no, doc_date_formatted, currency, items
