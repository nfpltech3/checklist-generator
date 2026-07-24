import sys, re, pdfplumber
sys.path.insert(0, '.')
import aviat_inv_extractor as ax

pdf = pdfplumber.open('243583994_CIPL-014-26 (rev1).pdf')
items = {}
last_line_no = None

for page_idx, page in enumerate(pdf.pages):
    text = page.extract_text() or ''
    if "PACKING LIST" in text[:200] and "COMMERCIAL INVOICE NO" not in text[:200]:
        break
    lines = text.split('\n')
    for line in lines:
        line = line.strip()
        if not line: continue
        
        item_start_match = re.match(r"^(\d+[.,]\d+)\s+(.+)$", line)
        if item_start_match:
            line_no_raw = item_start_match.group(1)
            if len(re.split(r'[.,]', line_no_raw)[0]) <= 3:
                line_no = line_no_raw.replace(',', '.')
                rest = item_start_match.group(2).strip()
                rest = rest.replace("(cid:9)", "\t")
                end_match = re.search(r"\s+(\d+)\s+([\d.,]+)\s+([\d.,]+)$", rest)
                
                if end_match:
                    middle = rest[:end_match.start()].strip()
                    desc = " ".join(middle.split('\t')[1:]).strip() if '\t' in middle else " ".join(middle.split()[1:])
                else:
                    middle = rest
                    desc = rest
                    
                desc = ax.clean_description(desc)
                
                if line_no == '14.13':
                    print("MATCH 14.13:", repr(line))
                    print("END MATCH SUCCESS:", bool(end_match))
                    print("DESC:", repr(desc))
                
                if line_no not in items:
                    items[line_no] = {'description': desc}
                else:
                    items[line_no]['description'] = f"{items[line_no]['description']} {desc}".strip()
                last_line_no = line_no
                continue
                
        if "ECCN :" in line or "CoO :" in line:
            last_line_no = None
            continue
            
        if last_line_no == '14.13':
            if not ax.should_ignore(line):
                line_cleaned = ax.clean_description(line)
                print("APPENDING TO 14.13:", repr(line_cleaned))
                items[last_line_no]['description'] = f"{items[last_line_no]['description']} {line_cleaned}".strip()

print("FINAL 14.13 DESC:", items.get('14.13', {}).get('description'))
