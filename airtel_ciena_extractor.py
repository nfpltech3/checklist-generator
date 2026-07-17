import pdfplumber
import re
from datetime import datetime

def parse_ciena_invoice(pdf_path: str):
    """
    Parse Bharti Airtel Ciena Commercial Invoice PDF.
    Returns: (invoice_no, doc_date_formatted, currency, items_list)
    """
    invoice_no = ""
    doc_date_formatted = ""
    currency = "USD"
    item_rows = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            words = page.extract_words()
            text = page.extract_text() or ""
            lines = text.split("\n")
            
            # 1. Extract Invoice Number
            if not invoice_no:
                inv_match = re.search(r"COMMERCIAL\s+INVOICE\s+NO\.\s*(\S+)", text, re.IGNORECASE)
                if inv_match:
                    invoice_no = inv_match.group(1)
                else:
                    fallback_match = re.search(r"Invoice\s*#:\s*(?:[^\n]*\n)?\s*(\S+)", text, re.IGNORECASE)
                    if fallback_match:
                        invoice_no = fallback_match.group(1)

            # 2. Extract Date (Usually near Place &Date or Company Registration on last page)
            date_match = re.search(r"(\d{1,2}-[A-Za-z]{3}-\d{4})", text)
            if date_match and not doc_date_formatted:
                doc_date_raw = date_match.group(1)
                try:
                    parsed_date = datetime.strptime(doc_date_raw, "%d-%b-%Y")
                    doc_date_formatted = parsed_date.strftime("%d-%m-%Y")
                except ValueError:
                    doc_date_formatted = doc_date_raw
                    
            # 3. Extract Currency
            curr_match = re.search(r"Currency:\s*(\S+)", text, re.IGNORECASE)
            if curr_match:
                currency = curr_match.group(1)

            # 4. Extract Line Items
            page_items = []
            for line in lines:
                # Pattern: Line# Qty UOM Part# HTS COO UnitPrice Total
                match = re.match(r"^(\d+)\s+(\d+)\s+([A-Za-z]{2,3})\s+(\S+)\s+(\d+)\s+([A-Z]{2})\s+([0-9.,]+)\s+([0-9.,]+)$", line)
                if match:
                    line_no = match.group(1)
                    qty = match.group(2)
                    uom = match.group(3)
                    part_no = match.group(4)
                    hts = match.group(5)
                    coo = match.group(6)
                    unit_val = match.group(7).replace(',', '')
                    
                    item_top = None
                    for w in words:
                        if w['text'] == part_no and abs(w['x0'] - 239.15) < 5:
                            item_top = w['top']
                            break
                            
                    # Find HTS suffix (like 090) on the line below
                    hts_suffix = ''
                    if item_top is not None:
                        for w in words:
                            if (item_top + 5) < w['top'] < (item_top + 12) and abs(w['x0'] - 406.2) < 10:
                                hts_suffix = w['text']
                                break
                    if hts_suffix:
                        hts += hts_suffix
                        
                    page_items.append({
                        'line_no': line_no,
                        'qty': qty,
                        'uom': uom,
                        'part_no': part_no,
                        'hts': hts,
                        'coo': coo,
                        'unit_val': unit_val,
                        'top': item_top
                    })
            
            page_items = sorted(page_items, key=lambda x: x['top'])
            
            # Find totals boundary
            totals_top = 800.0
            if page_items:
                last_top = page_items[-1]['top']
                for w in words:
                    if w['top'] > last_top and w['text'].lower() in ['total', 'declaration', 'freight:', 'insurance:', 'invoice']:
                        if w['top'] < totals_top:
                            totals_top = w['top']
            
            # Extract description for each item on the page
            for i, item in enumerate(page_items):
                next_top = page_items[i+1]['top'] if i + 1 < len(page_items) else totals_top
                desc_words = []
                for w in words:
                    if item['top'] < w['top'] < next_top:
                        if 230 <= w['x0'] <= 390:
                            desc_words.append(w)
                desc_words = sorted(desc_words, key=lambda w: (w['top'], w['x0']))
                desc_text = ' '.join([w['text'] for w in desc_words])
                
                full_desc = f"{item['part_no']} {desc_text}".strip()
                
                item_rows.append({
                    'line_no': item['line_no'],
                    'part_no': item['part_no'],
                    'description': full_desc,
                    'quantity': item['qty'],
                    'unit': item['uom'],
                    'currency': currency,
                    'rate': item['unit_val'],
                    'cth': item['hts'],
                    'coo': item['coo']
                })
                
    return invoice_no, doc_date_formatted, currency, item_rows
