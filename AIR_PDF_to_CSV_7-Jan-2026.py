import csv
import re
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False
import pdfplumber
import pandas as pd
import os
from datetime import datetime
import logging
import sys

# Setup logging
logging.basicConfig(
    filename='integrated_do_invoice_processor.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger()

class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
    def emit(self, record):
        try:
            if self.text_widget.winfo_exists():
                msg = self.format(record)
                self.text_widget.config(state='normal')
                self.text_widget.insert(tk.END, msg + '\n')
                self.text_widget.config(state='disabled')
                self.text_widget.see(tk.END)
        except Exception:
            pass

def load_job_register(job_register_path, log_callback):
    log_callback(f"Loading job register: {os.path.basename(job_register_path)}")
    ref_no_mapping = {}
    try:
        with open(job_register_path, 'r', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            for row in reader:
                # Strip spaces from values to prevent lookup failures
                be_no = str(row.get('BE No') or row.get('BOE No') or "").strip()
                job_no = str(row.get('Job No') or row.get('Job No.') or "").strip()
                if be_no and job_no:
                    ref_no_mapping[be_no] = job_no
        log_callback(f"Loaded {len(ref_no_mapping)} mappings from job register")
        logger.info(f"Loaded {len(ref_no_mapping)} mappings from job register")
        return ref_no_mapping
    except Exception as e:
        log_callback(f"Error loading job register: {e}")
        logger.error(f"Error loading job register: {e}")
        return None

def extract_text_from_pdf(pdf_path, log_callback):
    log_callback(f"Extracting text from {os.path.basename(pdf_path)}...")
    try:
        text = ""
        tables_data = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                if page.extract_text():
                    text += page.extract_text() + "\n"
                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        for row in table:
                            row_text = [str(cell) if cell else "" for cell in row]
                            tables_data.append(row_text)
        combined_text = text + "\n" + "\n".join([" ".join(row) for row in tables_data])
        return combined_text, tables_data
    except Exception as e:
        log_callback(f"Failed to extract text from {os.path.basename(pdf_path)}: {str(e)}")
        logger.error(f"Failed to extract text from {pdf_path}: {e}")
        return None, []

# --- Invoice Extraction Logic ---
def detect_invoice_type(text):
    if re.search(r"MUMBAI\s+CARGO\s+SERVICE\s+CENTER\s+AIRPORT", text, re.IGNORECASE):
        return "MIAL"
    if re.search(r"\bAI Airport Services Limited\b|\bAir India\b", text, re.IGNORECASE):
        return "AirIndia"
    return None

def extract_mial_fields(text, ref_no_mapping, log_callback):
    log_callback("Extracting MIAL fields with regex...")
    invoice_no_match = re.search(r"INVOICE No\s*:\s*([A-Z0-9]+)", text)
    invoice_date_match = re.search(r"Date & Time\s*:\s*([0-9]{2}-[A-Za-z]{3}-[0-9]{4}|[0-9]{2}-[0-9]{2}-[0-9]{4})", text)
    boe_no_match = re.search(r"BoE\. No / Date\s*:\s*([0-9]+)", text)
    demurrage_amount = "NotFound"
    demurrage_basic_amount = None
    demurrage_match = re.search(r"^Import Demurrage Charges.*?([\d,]+\.\d{2})\s*$", text, re.MULTILINE)
    if demurrage_match:
        demurrage_amount = demurrage_match.group(1).replace(",", "")
    # --- Extract basic DMC value from table ---
    lines = text.splitlines()
    for idx, line in enumerate(lines):
        if line.strip().lower().startswith('import demurrage charges'):
            # Split by whitespace, skip the charge name
            parts = line.strip().split()
            if len(parts) > 1:
                # The first number after the charge name is the basic DMC
                for part in parts[1:]:
                    try:
                        demurrage_basic_amount = float(part.replace(',', ''))
                        break
                    except ValueError:
                        continue
            break
    # If not found, fallback to demurrage_amount
    if demurrage_basic_amount is None:
        try:
            demurrage_basic_amount = float(demurrage_amount)
        except Exception:
            demurrage_basic_amount = None
    total_matches = re.findall(r"Total[^\d]*([\d,]+\.\d{2})", text)
    invoice_no = invoice_no_match.group(1) if invoice_no_match else "NotFound"
    invoice_date = invoice_date_match.group(1) if invoice_date_match else "NotFound"
    boe_no = boe_no_match.group(1) if boe_no_match else "NotFound"
    total_amount = total_matches[-1].replace(',', '') if total_matches else "NotFound"
    log_callback(f"[MIAL Extracted] Invoice No: {invoice_no}, Invoice Date: {invoice_date}, BOE No: {boe_no}, Demurrage: {demurrage_amount}, Total: {total_amount}")
    try:
        parsed_date = datetime.strptime(invoice_date, "%d-%b-%Y")
        invoice_date = parsed_date.strftime("%d-%b-%Y")
    except ValueError:
        try:
            parsed_date = datetime.strptime(invoice_date, "%d-%m-%Y")
            invoice_date = parsed_date.strftime("%d-%b-%Y")
        except Exception:
            invoice_date = "27-Jun-2025"
    cleaned_boe_no = "".join(c for c in str(boe_no) if c.isdigit())
    if boe_no != "NotFound" and (len(cleaned_boe_no) != 7 or not cleaned_boe_no.isdigit()):
        boe_no = "NotFound"
    ref_no = ref_no_mapping.get(boe_no, f"Unknown_{boe_no}" if boe_no != "NotFound" else "Unknown")
    try:
        demurrage = float(demurrage_amount) if demurrage_amount != "NotFound" else None
    except ValueError:
        demurrage = None
    try:
        total = float(total_amount) if total_amount != "NotFound" else None
    except ValueError:
        total = None
    # Table-based demurrage-only case: Only if the table under 'Charges Paid (If any) :' contains only Import Demurrage Charges and Round off Amount
    demurrage_only_case = False
    lines = text.splitlines()
    # Known table header patterns to ignore
    header_patterns = [
        'charges no.of', 'days', 'waiver', 'amount', 'tax', 'waivered', 'cgst', '(9%)', 'sgst', 'igst', '(18%)', 'total'
    ]
    def is_header_line(line):
        l = line.strip().lower()
        # If the line is empty, or matches any header pattern, or is just 'total', treat as header
        if not l:
            return True
        for pat in header_patterns:
            if l == pat or l == pat + ' total' or l == pat + ' :' or l == pat + ':' or l == pat + ' (if any) :' or l == pat + ' (if any)':
                return True
        # If the line is just a combination of header words, treat as header
        if all(any(word in l for word in header_patterns) for word in l.split()):
            return True
        return False
    # Find the start of the charges table
    table_start = None
    for idx, line in enumerate(lines):
        if 'Charges Paid (If any)' in line:
            table_start = idx + 1
            break
    table_lines = []
    charge_names = set()
    if table_start is not None:
        for line in lines[table_start:]:
            l = line.strip()
            # Stop at a line that starts with 'Total' (case-insensitive)
            if l.lower().startswith('total'):
                break
            # Only collect lines that match the charge line pattern (charge name followed by a number)
            if re.match(r'^[A-Za-z ()/\-]+\s+\d', l):
                table_lines.append(l)
                # Extract first column (charge name) using regex (everything up to first number)
                match = re.match(r'^([A-Za-z ()/\-]+)', l)
                if match:
                    charge_name = match.group(1).strip().lower()
                    charge_names.add(charge_name)
    log_callback(f"[DEBUG] Table charge lines: {table_lines}")
    log_callback(f"[DEBUG] Table extracted charge names (regex first column): {charge_names}")
    if charge_names == {'import demurrage charges', 'round off amount'}:
        demurrage_only_case = True
    invoice_data_list = []
    base_data = {
        'Type': 'MIAL',
        'Vendor Inv No': invoice_no,
        'Vendor Inv Date': invoice_date,
        'BOE No': boe_no,
        'Ref No': ref_no,
        'Organization': 'MUMBAI CARGO SERVICE CENTRE AIRPORT PRIVATE LIMITED',
        'Branch': 'MUMBAI',
        'Currency': 'INR',
    }
    if total is None:
        return []
    if demurrage_only_case:
        # Assign the value from the 'Total' line to Charge or GL Amount and Amount
        narration_val = demurrage_basic_amount if demurrage_basic_amount is not None else (demurrage_amount if demurrage_amount != "NotFound" else f"{total:.2f}")
        invoice_data_list.append({
            **base_data,
            'Charge or GL Name': 'Demurrage charges',
            'Total Amount': f"{total:.2f}",
            'Narration': f"Being Entry posted for MIAL charges / PD A/c 615 / DMC ({narration_val}) / {ref_no}"
        })
        return invoice_data_list
    if demurrage is not None and demurrage > 0:
        cargo_amount = total - demurrage if demurrage <= total else 0
        narration_val = demurrage_basic_amount if demurrage_basic_amount is not None else demurrage
        invoice_data_list.append({
            **base_data,
            'Charge or GL Name': 'Demurrage charges',
            'Total Amount': f"{demurrage:.2f}",
            'Narration': f"Being Entry posted for MIAL charges / PD A/c 615 / DMC ({narration_val}) / {ref_no}"
        })
    else:
        cargo_amount = total
    if cargo_amount > 0:
        invoice_data_list.append({
            **base_data,
            'Charge or GL Name': 'MUMBAI CARGO SERVICE (1)',
            'Total Amount': f"{cargo_amount:.2f}",
            'Narration': f'Being Entry posted for MIAL charges / PD A/c 615 / {ref_no}'
        })
    return invoice_data_list

def extract_airindia_fields(text, ref_no_mapping, log_callback):
    log_callback("Extracting Air India fields with regex...")
    invoice_no_match = re.search(r"Invoice No:?\s*([A-Z0-9]+)", text)
    invoice_no = invoice_no_match.group(1) if invoice_no_match else "Not Found"
    invoice_date_match = re.search(r"Invoice Date:?\s*([0-9]{2}/[0-9]{2}/[0-9]{4}|[0-9]{2}-[0-9]{2}-[0-9]{4}|[0-9]{2}-[A-Za-z]{3}-[0-9]{4})", text)
    invoice_date = invoice_date_match.group(1) if invoice_date_match else "Not Found"
    boe_no_match = re.search(r"BOE No\.?\s*([0-9]+)", text, re.IGNORECASE)
    boe_no = boe_no_match.group(1) if boe_no_match else "Not Found"
    # Extract TSP and Demurrage Charges
    tsp_match = re.search(r"TSP CHARGES\s+\d+\s+([\d,]+\.\d{2})", text)
    tsp_amount = tsp_match.group(1).replace(",", "") if tsp_match else None
    demurrage_match = re.search(r"DEMURRAGE CHARGES\s+\d+\s+([\d,]+\.\d{2})", text)
    demurrage_amount = demurrage_match.group(1).replace(",", "") if demurrage_match else None
    # Broadened regex for NET PAYABLE
    net_payable_match = re.search(r"NET PAYABLE\s*:?\s*(?:INR)?\s*([\d,]+\.\d{2})", text, re.IGNORECASE)
    if not net_payable_match:
        net_payable_match = re.search(r"NET PAYABLE[^\d]*([\d,]+\.\d{2})", text, re.IGNORECASE)
    net_payable = net_payable_match.group(1).replace(",", "") if net_payable_match else None
    # Fallback: use TSP CHARGES if present and non-zero, else DEMURRAGE CHARGES
    if not net_payable or float(net_payable) == 0.0:
        if tsp_amount and float(tsp_amount) != 0.0:
            net_payable = tsp_amount
        elif demurrage_amount and float(demurrage_amount) != 0.0:
            net_payable = demurrage_amount
        else:
            net_payable = None
    log_callback(f"[Air India Extracted] Invoice No: {invoice_no}, Invoice Date: {invoice_date}, BOE No: {boe_no}, NET PAYABLE: {net_payable}, TSP Charges: {tsp_amount}, Demurrage Charges: {demurrage_amount}")
    ref_no = ref_no_mapping.get(boe_no, f"Unknown_{boe_no}" if boe_no != "Not Found" else "Unknown")
    if invoice_date != "Not Found":
        try:
            parsed_date = datetime.strptime(invoice_date, "%d/%m/%Y")
            invoice_date = parsed_date.strftime("%d-%b-%Y")
        except ValueError:
            try:
                parsed_date = datetime.strptime(invoice_date, "%d-%m-%Y")
                invoice_date = parsed_date.strftime("%d-%b-%Y")
            except Exception:
                invoice_date = "Not Found"
    invoice_data_list = []
    base_data = {
        'Type': 'AirIndia',
        'Vendor Inv No': invoice_no,
        'Vendor Inv Date': invoice_date,
        'BOE No': boe_no,
        'Ref No': ref_no,
        'Organization': 'AI AIRPORT SERVICES LIMITED',
        'Branch': 'MUMBAI',
        'Currency': 'INR',
    }
    try:
        demurrage_val = float(demurrage_amount) if demurrage_amount else 0.0
        net_payable_val = float(net_payable) if net_payable else 0.0
    except Exception:
        demurrage_val = 0.0
        net_payable_val = 0.0
    if demurrage_val == 0:
        # Only AIR INDIA CHARGES (1) = NET PAYABLE, no GST multiplication
        if net_payable_val >= 1:
            invoice_data_list.append({
                **base_data,
                'Charge or GL Name': 'AIR INDIA CHARGES (1)',
                'Total Amount': f"{net_payable_val:.2f}",
                'Narration': f'Being Entry posted for Air India / PDA - 1170 / {ref_no}'
            })
    elif demurrage_val > 0:
        demurrage_total = demurrage_val * 1.18
        airindia_charges = net_payable_val - demurrage_total
        # If AIR INDIA CHARGES (1) < 1, add net_payable_val to demurrage row and do not append AIR INDIA CHARGES (1) row
        if airindia_charges < 1:
            invoice_data_list.append({
                **base_data,
                'Charge or GL Name': 'Demurrage Charges',
                'Total Amount': f"{net_payable_val:.2f}",
                'Narration': f"Being Entry posted for Air India / PDA - 1170 / DMC ({demurrage_val}) / {ref_no}"
            })
        else:
            invoice_data_list.append({
                **base_data,
                'Charge or GL Name': 'Demurrage Charges',
                'Total Amount': f"{demurrage_total:.2f}",
                'Narration': f"Being Entry posted for Air India / PDA - 1170 / DMC ({demurrage_val}) / {ref_no}"
            })
            invoice_data_list.append({
                **base_data,
                'Charge or GL Name': 'AIR INDIA CHARGES (1)',
                'Total Amount': f"{airindia_charges:.2f}",
                'Narration': f'Being Entry posted for Air India / PDA - 1170 / {ref_no}'
            })
    return invoice_data_list

# --- DO Extraction Logic ---
def clean_numeric_string(value):
    if isinstance(value, (str, float, int)):
        value = str(value).replace(",", "").strip()
        try:
            float_val = float(value)
            if float_val.is_integer():
                return str(int(float_val))
            return str(float_val)
        except ValueError:
            return str(value)
    return str(value)

def calculate_wh_tax(wh_tax_taxable):
    try:
        wh_tax_taxable = float(clean_numeric_string(wh_tax_taxable))
        wh_tax_percentage = 2  # Fixed as per template
        wh_tax_amount = wh_tax_taxable * (wh_tax_percentage / 100)
        return round(wh_tax_taxable, 2), round(wh_tax_amount, 2)
    except ValueError:
        return "Not Found", "Not Found"

def extract_invoice_details_with_regex(text, tables_data, log_callback):
    log_callback("Extracting fields using regex...")
    results = []
    try:
        if "Schenker India Pvt Ltd" in text:
            ref_no = re.search(r"HB/L No\.?[:\s]*([A-Z0-9/\-]+)", text)
            ref_val = ref_no.group(1) if ref_no else None
            mawb_no = None
            mawb_match = re.search(r"MAWB No\.?[:\s]*([A-Z0-9/\-]+)", text)
            if mawb_match:
                mawb_no = mawb_match.group(1)
            else:
                mawb_no = "Not Found"
            if not ref_val:
                for row in tables_data:
                    for cell in row:
                        m = re.search(r"HB/L No\.?[:\s]*([A-Z0-9/\-]+)", str(cell))
                        if m:
                            ref_val = m.group(1)
                            break
                    if ref_val:
                        break
            invoice_no = re.search(r"Invoice No\.?[:\s]*([0-9]{10})", text)
            invoice_date = re.search(r"Invoice Date\.?[:\s]*(\d{2}\.\d{2}\.\d{4}|\d{2}-\d{2}-\d{4}|\d{2}/\d{2}/\d{4})", text)
            total = re.search(r"Total Invoice / Credit Amount[:\s]*INR\s*([0-9,.]+)", text)
            taxable = re.search(r"Total net amount taxable[:\s]*INR\s*([0-9,.]+)", text)
            total_val = float(total.group(1).replace(',', '')) if total else 0.0
            taxable_val = float(taxable.group(1).replace(',', '')) if taxable else 0.0
            wh_tax_amount = taxable_val * 0.02
            extracted_date = invoice_date.group(1) if invoice_date else "Not Found"
            log_callback(f"Extracted invoice date: {extracted_date}")
            results.append({
                "Organization": "SCHENKER INDIA PVT LTD",
                "Vendor Inv No": invoice_no.group(1) if invoice_no else "Not Found",
                "Vendor Inv Date": extracted_date,
                "Narration": "DO CHARGES PAYABLE TO SCHENKER INDIA",
                "Amount": str(round(total_val, 2)),
                "Charge or GL Amount": str(round(total_val, 2)),
                "WH Tax Taxable": str(round(taxable_val, 2)),
                "WH Tax Amount": str(round(wh_tax_amount, 2)),
                "Ref No": ref_val if ref_val else "Not Found",
                "MAWB No": mawb_no
            })
        elif "DHL Logistics Pvt. Ltd." in text:
            ref_val = None
            mawb_no = None
            lines = text.splitlines()
            for i, line in enumerate(lines):
                if "HAWB" in line.upper():
                    for j in range(i+1, len(lines)):
                        next_line = lines[j].strip()
                        if next_line:
                            for word in next_line.split():
                                if len(word) >= 7 and re.match(r"^[A-Z0-9/\-]+$", word):
                                    ref_val = word
                                    break
                            break
                    if ref_val:
                        break
            for i, line in enumerate(lines):
                if "MAWB NUMBER" in line.upper():
                    for j in range(i+1, len(lines)):
                        next_line = lines[j].strip()
                        if next_line:
                            mawb_no = next_line.split()[0]
                            break
                    break
            if not mawb_no:
                mawb_no = "Not Found"
            if not ref_val:
                for row in tables_data:
                    for cell in row:
                        m = re.search(r"HAWB\s*([A-Z0-9]+)", str(cell))
                        if m:
                            ref_val = m.group(1)
                            break
                    if ref_val:
                        break
            invoice_no_patterns = [
                r'\b(IM[A-Z0-9]{5,})\b',
                r'Invoice Number\s*[:\-]?\s*([A-Z0-9]+)',
                r'Invoice No\.?\s*[:\-]?\s*([A-Z0-9]+)',
                r'Invoice\s*[A-Za-z]*\s*[:\-]?\s*([A-Z0-9]{6,})',
            ]
            extracted_invoice_no = None
            for pat in invoice_no_patterns:
                m = re.search(pat, text)
                if m:
                    extracted_invoice_no = m.group(1)
                    break
            invoice_date_patterns = [
                r'Invoice Date\s*[:\-]?\s*(\d{1,2}-[A-Za-z]{3}-\d{4})',
                r'Invoice Date\s*[:\-]?\s*(\d{1,2}/\d{1,2}/\d{4})',
                r'Invoice Date\s*[:\-]?\s*([A-Za-z]{3}\s+\d{1,2},\s+\d{4})',
                r'Invoice Date\s*[:\-]?\s*(\d{1,2}-[A-Za-z]{3}-\d{2})',
                r'Invoice\s*Date\s*[:\-]?\s*(\d{1,2}\.\d{1,2}\.\d{4})',
                r'\b(\d{1,2}-[A-Za-z]{3}-\d{4})\b',
                r'\b(\d{1,2}/\d{1,2}/\d{4})\b',
            ]
            extracted_date = None
            for pat in invoice_date_patterns:
                m = re.search(pat, text)
                if m:
                    extracted_date = m.group(1)
                    break
            if not extracted_date:
                extracted_date = "Not Found"
            log_callback(f"Extracted invoice date: {extracted_date}")
            total = re.search(r"DEBIT\s*INR\s*([0-9,.]+)", text)
            taxable = re.search(r"Taxable Amount \(INR\)\s*:\s*([0-9,.]+)", text)
            total_val = float(total.group(1).replace(',', '')) if total else 0.0
            taxable_val = float(taxable.group(1).replace(',', '')) if taxable else 0.0
            wh_tax_amount = taxable_val * 0.02
            results.append({
                "Organization": "DHL LOGISTICS PVT LTD",
                "Vendor Inv No": extracted_invoice_no if extracted_invoice_no else "Not Found",
                "Vendor Inv Date": extracted_date,
                "Narration": "DO CHARGES PAYABLE TO DHL",
                "Amount": str(round(total_val, 2)),
                "Charge or GL Amount": str(round(total_val, 2)),
                "WH Tax Taxable": str(round(taxable_val, 2)),
                "WH Tax Amount": str(round(wh_tax_amount, 2)),
                "Ref No": ref_val if ref_val else "Not Found",
                "MAWB No": mawb_no
            })
        elif "Hellmann Worldwide Logistics India" in text:
            ref_val = None
            mawb_no = "Not Found"
            lines = text.splitlines()
            for i, line in enumerate(lines):
                if "HAWB" in line.upper():
                    for j in range(i+1, len(lines)):
                        next_line = lines[j].strip()
                        if next_line:
                            for word in next_line.split():
                                if len(word) >= 7 and re.match(r"^[A-Z0-9/\-]+$", word):
                                    ref_val = word
                                    break
                            break
                    if ref_val:
                        break
            if not ref_val:
                for row in tables_data:
                    for cell in row:
                        m = re.search(r"HAWB\s*([A-Z0-9]+)", str(cell))
                        if m:
                            ref_val = m.group(1)
                            break
                    if ref_val:
                        break
            invoice_no = re.search(r"Tax Invoice\s+([A-Z0-9]+)", text)
            invoice_date = re.search(r"INVOICE DATE\s*(\d{2}-[A-Za-z]{3}-\d{2,4})", text)
            total = re.search(r"TOTAL INR\s*([0-9,.]+)", text)
            taxable = re.search(r"SUBTOTAL\s*([0-9,.]+)", text)
            total_val = float(total.group(1).replace(',', '')) if total else 0.0
            taxable_val = float(taxable.group(1).replace(',', '')) if taxable else 0.0
            wh_tax_amount = taxable_val * 0.02
            extracted_date = invoice_date.group(1) if invoice_date else "Not Found"
            log_callback(f"Extracted invoice date: {extracted_date}")
            results.append({
                "Organization": "HELLMANN WORLDWIDE LOGISTICS INDIA PVT. LTD.",
                "Vendor Inv No": invoice_no.group(1) if invoice_no else "Not Found",
                "Vendor Inv Date": extracted_date,
                "Narration": "DO CHARGES PAYABLE TO HELLMANN",
                "Amount": str(round(total_val, 2)),
                "Charge or GL Amount": str(round(total_val, 2)),
                "WH Tax Taxable": str(round(taxable_val, 2)),
                "WH Tax Amount": str(round(wh_tax_amount, 2)),
                "Ref No": ref_val if ref_val else "Not Found",
                "MAWB No": mawb_no
            })
        elif "DSV Air & Sea Pvt. Ltd." in text:
            ref_val = None
            dsv_type = None
            mawb_no = "Not Found"
            lines = text.splitlines()
            found_sea = False
            for i, line in enumerate(lines):
                if "HOUSE BILL OF LADING" in line.upper():
                    dsv_type = "SEA"
                    found_sea = True
                    for j in range(i+1, len(lines)):
                        next_line = lines[j].strip()
                        if next_line:
                            words = next_line.split()
                            if words:
                                ref_val = words[-1]
                            break
                    if ref_val:
                        break
            if not found_sea:
                for i, line in enumerate(lines):
                    if line.strip() == "HAWB":
                        dsv_type = "AIR"
                        for j in range(i+1, len(lines)):
                            next_line = lines[j].strip()
                            if next_line:
                                words = next_line.split()
                                if words:
                                    ref_val = words[-1]
                                break
                        if ref_val:
                            break
                if not ref_val or mawb_no == "Not Found":
                    for i, line in enumerate(lines):
                        if "MAWB" in line and "HAWB" in line:
                            for j in range(i+1, len(lines)):
                                next_line = lines[j].strip()
                                if next_line:
                                    words = next_line.split()
                                    if words:
                                        ref_val = words[-1]
                                        if len(words) >= 2:
                                            mawb_no = words[-2]
                                        else:
                                            mawb_no = words[0]
                                    break
                            if ref_val or mawb_no != "Not Found":
                                break
            if not ref_val:
                if dsv_type == "SEA":
                    for row in tables_data:
                        for cell in row:
                            m = re.search(r"HOUSE BILL OF LADING\s*([A-Z0-9]+)", str(cell))
                            if m:
                                ref_val = m.group(1)
                                break
                        if ref_val:
                            break
                elif dsv_type == "AIR":
                    for row in tables_data:
                        for cell in row:
                            m = re.search(r"HAWB\s*([A-Z0-9]+)", str(cell))
                            if m:
                                ref_val = m.group(1)
                                break
                            if "MAWB" in str(cell) and "HAWB" in str(cell):
                                idx = row.index(cell)
                                if idx + 1 < len(row):
                                    next_cell = row[idx + 1]
                                    words = str(next_cell).split()
                                    if words:
                                        ref_val = words[-1]
                                        if len(words) >= 2:
                                            mawb_no = words[-2]
                                        else:
                                            mawb_no = words[0]
                                        break
                        if ref_val:
                            break
            invoice_no = re.search(r"TAX INVOICE\s+(IN1BOM[0-9A-Z]+)", text)
            invoice_date = re.search(r"INVOICE DATE\s*(\d{2}-[A-Za-z]{3}-\d{2,4})", text)
            total = re.search(r"TOTAL INR\s*([0-9,.]+)", text)
            taxable = re.search(r"SUBTOTAL\s*([0-9,.]+)", text)
            total_val = float(total.group(1).replace(',', '')) if total else 0.0
            taxable_val = float(taxable.group(1).replace(',', '')) if taxable else 0.0
            wh_tax_amount = taxable_val * 0.02
            extracted_date = invoice_date.group(1) if invoice_date else "Not Found"
            log_callback(f"Extracted invoice date: {extracted_date}")
            results.append({
                "Organization": "DSV AIR AND SEA PVT LTD",
                "Vendor Inv No": invoice_no.group(1) if invoice_no else "Not Found",
                "Vendor Inv Date": extracted_date,
                "Narration": "DO CHARGES PAYABLE TO DSV AIR & SEA",
                "Amount": str(round(total_val, 2)),
                "Charge or GL Amount": str(round(total_val, 2)),
                "WH Tax Taxable": str(round(taxable_val, 2)),
                "WH Tax Amount": str(round(wh_tax_amount, 2)),
                "Ref No": ref_val if ref_val else "Not Found",
                "MAWB No": mawb_no
            })
        elif "Expeditors International (India) Private Limited" in text:
            hawb = None
            mawb = None
            lines = text.splitlines()
            for line in lines:
                if "HAWB / HBL" in line.upper():
                    m = re.search(r"HAWB / HBL[:\s]*([A-Z0-9\-/]+)", line, re.IGNORECASE)
                    if m:
                        hawb = m.group(1)
                if "AWB / BL" in line.upper():
                    m = re.search(r"AWB / BL[:\s]*([A-Z0-9\-/]+)", line, re.IGNORECASE)
                    if m:
                        mawb = m.group(1)
            if not hawb:
                for row in tables_data:
                    for cell in row:
                        m = re.search(r"HAWB / HBL[:\s]*([A-Z0-9\-/]+)", str(cell), re.IGNORECASE)
                        if m:
                            hawb = m.group(1)
                            break
                    if hawb:
                        break
            if not mawb:
                for row in tables_data:
                    for cell in row:
                        m = re.search(r"AWB / BL[:\s]*([A-Z0-9\-/]+)", str(cell), re.IGNORECASE)
                        if m:
                            mawb = m.group(1)
                            break
                    if mawb:
                        break
            invoice_no = re.search(r"INVOICE NUMBER[:\s]*([A-Z0-9]+)", text, re.IGNORECASE)
            invoice_date = re.search(r"INVOICE DATE[:\s]*(\d{2}/\d{2}/\d{4})", text, re.IGNORECASE)
            rounded_amount = re.search(r"ROUNDED-OFF AMOUNT[:\s]*([\d,]+\.\d{2})", text, re.IGNORECASE)
            subtotal = re.search(r"Sub-total\s*([\d,]+\.\d{2})", text, re.IGNORECASE)
            if not rounded_amount:
                rounded_amount = re.search(r"TOTAL INVOICE AMOUNT[:\s]*([\d,]+\.\d{2})", text, re.IGNORECASE)
            total_val = float(rounded_amount.group(1).replace(',', '')) if rounded_amount else 0.0
            taxable_val = float(subtotal.group(1).replace(',', '')) if subtotal else 0.0
            wh_tax_amount = taxable_val * 0.02
            extracted_date = invoice_date.group(1) if invoice_date else "Not Found"
            log_callback(f"Extracted invoice date: {extracted_date}")
            results.append({
                "Organization": "EXPEDITORS INTERNATIONAL (INDIA) PVT.LTD.",
                "Vendor Inv No": invoice_no.group(1) if invoice_no else "Not Found",
                "Vendor Inv Date": extracted_date,
                "Narration": "DO CHARGES PAYABLE TO EXPEDITORS",
                "Amount": str(round(total_val, 2)),
                "Charge or GL Amount": str(round(total_val, 2)),
                "WH Tax Taxable": str(round(taxable_val, 2)),
                "WH Tax Amount": str(round(wh_tax_amount, 2)),
                "Ref No": hawb if hawb else "Not Found",
                "MAWB No": mawb if mawb else "Not Found"
            })
        elif "Kuehne + Nagel" in text or "KUEHNE+NAGEL" in text.upper():
            # KUEHNE + NAGEL extraction
            hawb = None
            mawb = None
            # Extract AWB NO(S) - format: "AWB NO(S) : 1070505250 / 020-71533324"
            awb_match = re.search(r"AWB NO\(S\)\s*:\s*([A-Z0-9]+)\s*/\s*([A-Z0-9\-]+)", text, re.IGNORECASE)
            if awb_match:
                mawb = awb_match.group(1)  # First part is MAWB
                hawb = awb_match.group(2)  # Second part is HAWB
            # Invoice number from KN TRACKING NUMBER (remove spaces)
            invoice_no_match = re.search(r"KN TRACKING NUMBER\s*([\d\s]+)", text, re.IGNORECASE)
            invoice_no = invoice_no_match.group(1).replace(" ", "").strip() if invoice_no_match else None
            # Invoice date from "INVOICE NO. / DATE [optional invoice no] DD.MM.YYYY"
            invoice_date = re.search(r"INVOICE NO\.?\s*/\s*DATE\s*(?:[A-Z0-9]+\s+)?(\d{2}\.\d{2}\.\d{4})", text, re.IGNORECASE)
            # Total amount from "TOTAL DUE INR X,XXX.XX"
            total = re.search(r"TOTAL DUE\s*INR\s*([\d,]+\.\d{2})", text, re.IGNORECASE)
            # Taxable amount from "SUBTOTAL INR X,XXX.XX"
            subtotal = re.search(r"SUBTOTAL\s*INR\s*([\d,]+\.\d{2})", text, re.IGNORECASE)
            total_val = float(total.group(1).replace(',', '')) if total else 0.0
            taxable_val = float(subtotal.group(1).replace(',', '')) if subtotal else 0.0
            wh_tax_amount = taxable_val * 0.02
            extracted_date = invoice_date.group(1) if invoice_date else "Not Found"
            log_callback(f"[KUEHNE+NAGEL] Invoice: {invoice_no if invoice_no else 'Not Found'}, Date: {extracted_date}, AWB: {mawb}, HAWB: {hawb}, Total: {total_val}")
            results.append({
                "Organization": "KUEHNE + NAGEL PVT LTD",
                "Vendor Inv No": invoice_no if invoice_no else "Not Found",
                "Vendor Inv Date": extracted_date,
                "Narration": "DO CHARGES PAYABLE TO KUEHNE + NAGEL",
                "Amount": str(round(total_val, 2)),
                "Charge or GL Amount": str(round(total_val, 2)),
                "WH Tax Taxable": str(round(taxable_val, 2)),
                "WH Tax Amount": str(round(wh_tax_amount, 2)),
                "Ref No": hawb if hawb else "Not Found",
                "MAWB No": mawb if mawb else "Not Found"
            })
        elif "MAERSK" in text.upper() or "Senator" in text:
            # SENATOR / MAERSK extraction
            hawb = None
            mawb = None
            # Extract TAX INVOICE NO - format: "TAX INVOICE NO. I9100069293" or "TAX INVOICE NO I9100069293"
            invoice_no_match = re.search(r"TAX INVOICE NO\.?\s*([A-Z0-9]+)", text, re.IGNORECASE)
            invoice_no = invoice_no_match.group(1) if invoice_no_match else None
            # Extract INVOICE DATE - format: "INVOICE DATE 16-Dec-25"
            invoice_date = re.search(r"INVOICE DATE\s*(\d{1,2}-[A-Za-z]{3}-\d{2,4})", text, re.IGNORECASE)
            # Extract MAWB and HAWB from table - format: "FLIGHT / DATE MAWB HAWB" header
            # followed by values like "EK0046 / 12-Dec 17612130580 HAJ00021951"
            # MAWB is typically 11 digits, HAWB starts with HAJ
            mawb_match = re.search(r"(\d{10,12})\s+(HAJ\d+)", text, re.IGNORECASE)
            if mawb_match:
                mawb = mawb_match.group(1)
                hawb = mawb_match.group(2)
            else:
                # Alternative: look for any 10+ digit number followed by alphanumeric HAWB
                # after the MAWB HAWB header line
                alt_match = re.search(r"MAWB\s+HAWB\s*\n[^\n]*?(\d{10,})\s+([A-Z0-9]{8,})", text, re.IGNORECASE)
                if alt_match:
                    mawb = alt_match.group(1)
                    hawb = alt_match.group(2)
            # Total amount from "TOTAL INR X,XXX.XX"
            total = re.search(r"TOTAL\s+INR\s*([0-9,]+\.?\d*)", text, re.IGNORECASE)
            # Taxable amount from "SUBTOTAL X,XXX.XX"
            subtotal = re.search(r"SUBTOTAL\s*([0-9,]+\.?\d*)", text, re.IGNORECASE)
            total_val = float(total.group(1).replace(',', '')) if total else 0.0
            taxable_val = float(subtotal.group(1).replace(',', '')) if subtotal else 0.0
            wh_tax_amount = taxable_val * 0.02
            extracted_date = invoice_date.group(1) if invoice_date else "Not Found"
            log_callback(f"[SENATOR/MAERSK] Invoice: {invoice_no if invoice_no else 'Not Found'}, Date: {extracted_date}, MAWB: {mawb}, HAWB: {hawb}, Total: {total_val}")
            results.append({
                "Organization": "SENATOR LOGISTICS INDIA PRIVATE LIMITED",
                "Vendor Inv No": invoice_no if invoice_no else "Not Found",
                "Vendor Inv Date": extracted_date,
                "Narration": "DO CHARGES PAYABLE TO SENATOR",
                "Amount": str(round(total_val, 2)),
                "Charge or GL Amount": str(round(total_val, 2)),
                "WH Tax Taxable": str(round(taxable_val, 2)),
                "WH Tax Amount": str(round(wh_tax_amount, 2)),
                "Ref No": hawb if hawb else "Not Found",
                "MAWB No": mawb if mawb else "Not Found"
            })
        elif "OCEAN NETWORK EXPRESS" in text.upper() or "ONE MISSION" in text or "ONE LINE" in text.upper():
            # OCEAN NETWORK EXPRESS (ONE) extraction - SEA freight
            bl_no = None
            container_no = None
            # Extract Invoice No - format: "Invoice No IN27250041202"
            invoice_no_match = re.search(r"Invoice No\s*([A-Z]{2}\d{11,})", text, re.IGNORECASE)
            invoice_no = invoice_no_match.group(1) if invoice_no_match else None
            # Extract Issue Date - format: "Issue Date 14Jun2025"
            invoice_date = re.search(r"Issue Date\s*(\d{1,2}[A-Za-z]{3}\d{4})", text, re.IGNORECASE)
            # Extract B/L No - format: "B/L No RTMF10438700" (appears after Booking No)
            bl_match = re.search(r"B/L No\s*(\w+)", text, re.IGNORECASE)
            if bl_match:
                bl_no = bl_match.group(1)
            else:
                # Alternative: look for Booking No
                bl_match = re.search(r"Booking\s*No\s*(\w+)", text, re.IGNORECASE)
                if bl_match:
                    bl_no = bl_match.group(1)
            # Extract Container No - format: "CONTAINER NO : FFAU6645735"
            container_match = re.search(r"CONTAINER NO\s*:\s*([A-Z]{4}\d{7})", text, re.IGNORECASE)
            if container_match:
                container_no = container_match.group(1)
            # Total Invoice Value - format: "Total Invoice Value (in figure) 43,217.50"
            total = re.search(r"Total Invoice Value \(in figure\)\s*([0-9,]+\.?\d*)", text, re.IGNORECASE)
            # Taxable amount from table (the value before tax) - look for pattern in tax table
            # CGST INR 36,625.00 3,296.25 -> taxable is 36,625.00
            subtotal = re.search(r"CGST\s+INR\s+([0-9,]+\.?\d*)", text, re.IGNORECASE)
            if not subtotal:
                # Alternative: look at the line before Total Tax
                subtotal = re.search(r"([0-9,]+\.\d{2})\s+[0-9,]+\.\d{2}\s*$.*Total Invoice", text, re.MULTILINE | re.IGNORECASE)
            total_val = float(total.group(1).replace(',', '')) if total else 0.0
            taxable_val = float(subtotal.group(1).replace(',', '')) if subtotal else 0.0
            # WH Tax is 0 for Ocean Network Express
            wh_tax_amount = 0.0
            # Format date: 14Jun2025 -> 14-Jun-2025
            extracted_date = "Not Found"
            if invoice_date:
                raw_date = invoice_date.group(1)
                # Parse 14Jun2025 format
                try:
                    parsed_date = datetime.strptime(raw_date, "%d%b%Y")
                    extracted_date = parsed_date.strftime("%d-%b-%Y")
                except ValueError:
                    extracted_date = raw_date
            log_callback(f"[OCEAN NETWORK EXPRESS] Invoice: {invoice_no if invoice_no else 'Not Found'}, Date: {extracted_date}, B/L No: {bl_no}, Container: {container_no}, Total: {total_val}")
            results.append({
                "Organization": "OCEAN NETWORK EXPRESS PTE LTD",
                "Vendor Inv No": invoice_no if invoice_no else "Not Found",
                "Vendor Inv Date": extracted_date,
                "Narration": "DO CHARGES PAYABLE TO OCEAN NETWORK EXPRESS",
                "Amount": str(round(total_val, 2)),
                "Charge or GL Amount": str(round(total_val, 2)),
                "WH Tax Taxable": "0",
                "WH Tax Amount": "0",
                "Ref No": bl_no if bl_no else "Not Found",
                "MAWB No": container_no if container_no else "Not Found"
            })
        elif "DACHSER India Pvt" in text or "DACHSER INDIA" in text.upper():
            # DACHSER INDIA PVT LTD extraction
            hawb = None
            mawb = None
            
            # Extract Our Reference as Invoice No - format: "Our Reference" followed by value like "70300198673"
            # The value is typically 10+ digits
            our_ref_match = re.search(r"Our Reference\s*\n?\s*(\d{10,})", text, re.IGNORECASE)
            if not our_ref_match:
                # Alternative: look for the pattern with newline between label and value
                our_ref_match = re.search(r"Our Reference[^\d]*(\d{10,})", text, re.IGNORECASE)
            if not our_ref_match:
                # Fallback: look for pattern with "/" separator like "70300198673/3"
                our_ref_match = re.search(r"Our Reference\s*\n?\s*(\d+/?\d*)", text, re.IGNORECASE)
            invoice_no = our_ref_match.group(1) if our_ref_match else None
            
            # Extract Date - the date is in the line after "Document No. Customer No. Date" header
            # Format: "2025-10-27" in YYYY-MM-DD format
            lines = text.split('\n')
            invoice_date = None
            for i, line in enumerate(lines):
                if 'Document No.' in line and 'Date' in line:
                    if i + 1 < len(lines):
                        next_line = lines[i + 1]
                        date_match = re.search(r'(\d{4}-\d{2}-\d{2})', next_line)
                        if date_match:
                            invoice_date = date_match
                            break
            
            # Extract HAWB and MAWB - they are in the line after "HAWB No. MAWB No." header
            # Format: "AAI-0035293 157-49686744 06 Haryana"
            for i, line in enumerate(lines):
                if 'HAWB No' in line and 'MAWB No' in line:
                    if i + 1 < len(lines):
                        next_line = lines[i + 1]
                        # HAWB is the first alphanumeric pattern (e.g., AAI-0035293)
                        hawb_match = re.search(r'^([A-Z]{2,3}-?\d+)', next_line, re.IGNORECASE)
                        if hawb_match:
                            hawb = hawb_match.group(1)
                        # MAWB is the numeric pattern with dashes (e.g., 157-49686744)
                        mawb_match = re.search(r'(\d{3}-\d{8})', next_line)
                        if mawb_match:
                            mawb = mawb_match.group(1)
                    break
            
            # Extract Gross Total from tax table - format: "Gross Total INR" column with value like "14,624.33"
            # The Gross Total INR is typically the last value on a line in the tax summary table
            total_val = None
            total = re.search(r"Gross Total\s*INR?\s*.*?([0-9,]+\.\d{2})\s*$", text, re.MULTILINE | re.IGNORECASE)
            if total:
                total_val = float(total.group(1).replace(',', ''))
            else:
                # Alternative: look for the last decimal number on a line containing "Gross Total"
                gross_lines = re.findall(r".*Gross Total.*", text, re.IGNORECASE)
                for line in gross_lines:
                    amounts = re.findall(r"([0-9,]+\.\d{2})", line)
                    if amounts:
                        # Take the last amount on the line (should be Gross Total INR)
                        total_val = float(amounts[-1].replace(',', ''))
                        break
            
            if total_val is None:
                # Final fallback
                total = re.search(r"Gross Total[^\d]*(\d[\d,]*\.\d{2})", text, re.IGNORECASE)
                total_val = float(total.group(1).replace(',', '')) if total else 0.0
            
            # Extract Net Total (subtotal/taxable amount) - format: "Net Total" followed by amount
            subtotal = re.search(r"Net Total\s*([0-9,]+\.?\d*)", text, re.IGNORECASE)
            if not subtotal:
                # Alternative: look for SUBTOTAL pattern
                subtotal = re.search(r"SUBTOTAL\s*([0-9,]+\.?\d*)", text, re.IGNORECASE)
            
            taxable_val = float(subtotal.group(1).replace(',', '')) if subtotal else 0.0
            wh_tax_amount = taxable_val * 0.02
            
            # Format date: 2025-10-27 -> 27-Oct-2025
            extracted_date = "Not Found"
            if invoice_date:
                raw_date = invoice_date.group(1)
                try:
                    # Try YYYY-MM-DD format
                    parsed_date = datetime.strptime(raw_date, "%Y-%m-%d")
                    extracted_date = parsed_date.strftime("%d-%b-%Y")
                except ValueError:
                    try:
                        # Try DD-MM-YYYY format
                        parsed_date = datetime.strptime(raw_date, "%d-%m-%Y")
                        extracted_date = parsed_date.strftime("%d-%b-%Y")
                    except ValueError:
                        try:
                            # Try DD/MM/YYYY format
                            parsed_date = datetime.strptime(raw_date, "%d/%m/%Y")
                            extracted_date = parsed_date.strftime("%d-%b-%Y")
                        except ValueError:
                            extracted_date = raw_date
            
            log_callback(f"[DACHSER] Invoice: {invoice_no if invoice_no else 'Not Found'}, Date: {extracted_date}, MAWB: {mawb}, HAWB: {hawb}, Total: {total_val}")
            results.append({
                "Organization": "DACHSER INDIA PVT LTD",
                "Vendor Inv No": invoice_no if invoice_no else "Not Found",
                "Vendor Inv Date": extracted_date,
                "Narration": "DO CHARGES PAYABLE TO DACHSER",
                "Amount": str(round(total_val, 2)),
                "Charge or GL Amount": str(round(total_val, 2)),
                "WH Tax Taxable": "0",
                "WH Tax Amount": "0",
                "Ref No": hawb if hawb else "Not Found",
                "MAWB No": mawb if mawb else "Not Found"
            })
        elif "Allcargo Logistics Limited" in text or "ALLCARGO LOGISTICS" in text.upper():
            # ALLCARGO LOGISTICS LIMITED extraction
            docket_no = None
            
            # Extract Invoice No - format: "Invoice No :" followed by value like "MH/PD/26/0045190"
            invoice_no_match = re.search(r"Invoice No\s*:?\s*([A-Z]{2}/[A-Z]+/\d+/\d+)", text, re.IGNORECASE)
            if not invoice_no_match:
                # Alternative pattern
                invoice_no_match = re.search(r"Invoice No[:\s]+([A-Z0-9/]+)", text, re.IGNORECASE)
            invoice_no = invoice_no_match.group(1) if invoice_no_match else None
            
            # Extract Invoice Date - format: "Invoice Date :" followed by value like "19-DEC-25"
            invoice_date = re.search(r"Invoice Date\s*:?\s*(\d{1,2}-[A-Za-z]{3}-\d{2,4})", text, re.IGNORECASE)
            if not invoice_date:
                # Alternative format: DD/MM/YYYY
                invoice_date = re.search(r"Invoice Date\s*:?\s*(\d{1,2}/\d{1,2}/\d{2,4})", text, re.IGNORECASE)
            
            # Extract Docket No as reference - format: "Docket No :" followed by value like "423338731"
            docket_match = re.search(r"Docket No\s*:?\s*(\d+)", text, re.IGNORECASE)
            if docket_match:
                docket_no = docket_match.group(1)
            
            # Extract TOTAL amount - format: "TOTAL" followed by amount like "6460.5"
            total = re.search(r"TOTAL\s*[:\s]*([0-9,]+\.?\d*)", text, re.IGNORECASE)
            if not total:
                # Alternative: look for total in table
                total = re.search(r"TOTAL\s+([0-9,]+\.?\d*)", text, re.IGNORECASE)
            
            # Extract Taxable Value - format: "Taxable Value" followed by amount like "5475"
            taxable = re.search(r"Taxable Value\s*[:\s]*([0-9,]+\.?\d*)", text, re.IGNORECASE)
            if not taxable:
                # Alternative: look for amount before CGST
                taxable = re.search(r"Amount\s*\n?\s*([0-9,]+\.?\d*)\s*\n?\s*Taxable", text, re.IGNORECASE | re.MULTILINE)
            
            total_val = float(total.group(1).replace(',', '')) if total else 0.0
            taxable_val = float(taxable.group(1).replace(',', '')) if taxable else 0.0
            wh_tax_amount = taxable_val * 0.02
            
            # Format date
            extracted_date = "Not Found"
            if invoice_date:
                raw_date = invoice_date.group(1)
                try:
                    # Try DD-Mon-YY format
                    parsed_date = datetime.strptime(raw_date, "%d-%b-%y")
                    extracted_date = parsed_date.strftime("%d-%b-%Y")
                except ValueError:
                    try:
                        # Try DD-Mon-YYYY format
                        parsed_date = datetime.strptime(raw_date, "%d-%b-%Y")
                        extracted_date = parsed_date.strftime("%d-%b-%Y")
                    except ValueError:
                        try:
                            # Try DD/MM/YYYY format
                            parsed_date = datetime.strptime(raw_date, "%d/%m/%Y")
                            extracted_date = parsed_date.strftime("%d-%b-%Y")
                        except ValueError:
                            extracted_date = raw_date
            
            log_callback(f"[ALLCARGO] Invoice: {invoice_no if invoice_no else 'Not Found'}, Date: {extracted_date}, Docket: {docket_no}, Total: {total_val}")
            results.append({
                "Organization": "ALLCARGO LOGISTICS LTD",
                "Vendor Inv No": invoice_no if invoice_no else "Not Found",
                "Vendor Inv Date": extracted_date,
                "Narration": "DO CHARGES PAYABLE TO ALLCARGO LOGISTICS",
                "Amount": str(round(total_val, 2)),
                "Charge or GL Amount": str(round(total_val, 2)),
                "WH Tax Taxable": str(round(taxable_val, 2)),
                "WH Tax Amount": str(round(wh_tax_amount, 2)),
                "Ref No": docket_no if docket_no else "Not Found",
                "MAWB No": "Not Found"
            })
        elif "APEXGLOBAL FORWARDERS" in text.upper() or "Apexlogistics" in text:
            # APEXGLOBAL FORWARDERS INDIA PRIVATE LIMITED extraction
            hawb = None
            mawb = None
            
            # Extract Invoice Number - format: "Invoice Number" followed by value like "MAAAR25121733"
            invoice_no_match = re.search(r"Invoice Number\s*[:\s]*([A-Z0-9]+)", text, re.IGNORECASE)
            invoice_no = invoice_no_match.group(1) if invoice_no_match else None
            
            # Extract Date - format: "Date" followed by value like "12.23.2025" (MM.DD.YYYY)
            # The date appears after "Invoice Number" line
            invoice_date = re.search(r"\bDate\s*[:\s]*(\d{1,2}\.\d{1,2}\.\d{4})", text, re.IGNORECASE)
            if not invoice_date:
                # Alternative format: DD-MM-YYYY or DD/MM/YYYY
                invoice_date = re.search(r"\bDate\s*[:\s]*(\d{1,2}[-/]\d{1,2}[-/]\d{4})", text, re.IGNORECASE)
            
            # Extract House Number as HAWB - format: "House Number" followed by value like "SHAAEB25442"
            house_match = re.search(r"House Number\s*[:\s]*([A-Z0-9]+)", text, re.IGNORECASE)
            if house_match:
                hawb = house_match.group(1)
            
            # Extract Master Number as MAWB - format: "Master Number" followed by value like "176-22541234"
            master_match = re.search(r"Master Number\s*[:\s]*([\d\-]+)", text, re.IGNORECASE)
            if master_match:
                mawb = master_match.group(1)
            
            # Extract Grand Total - format: "Grand Total" followed by amount like "14768"
            total = re.search(r"Grand Total\s*[:\s]*([0-9,]+\.?\d*)", text, re.IGNORECASE)
            if not total:
                # Alternative: look for "Total" at end of document
                total = re.search(r"\bTotal\s+([0-9,]+\.?\d*)\s*\n\s*Grand Total", text, re.IGNORECASE)
            
            # Extract SUBTOTAL as taxable - format: "SUBTOTAL" followed by amount like "12515"
            subtotal = re.search(r"SUBTOTAL\s*[:\s]*([0-9,]+\.?\d*)", text, re.IGNORECASE)
            
            total_val = float(total.group(1).replace(',', '')) if total else 0.0
            taxable_val = float(subtotal.group(1).replace(',', '')) if subtotal else 0.0
            wh_tax_amount = taxable_val * 0.02
            
            # Format date: 12.23.2025 (MM.DD.YYYY) -> 23-Dec-2025
            extracted_date = "Not Found"
            if invoice_date:
                raw_date = invoice_date.group(1)
                try:
                    # Try MM.DD.YYYY format
                    parsed_date = datetime.strptime(raw_date, "%m.%d.%Y")
                    extracted_date = parsed_date.strftime("%d-%b-%Y")
                except ValueError:
                    try:
                        # Try DD.MM.YYYY format
                        parsed_date = datetime.strptime(raw_date, "%d.%m.%Y")
                        extracted_date = parsed_date.strftime("%d-%b-%Y")
                    except ValueError:
                        try:
                            # Try DD-MM-YYYY format
                            parsed_date = datetime.strptime(raw_date, "%d-%m-%Y")
                            extracted_date = parsed_date.strftime("%d-%b-%Y")
                        except ValueError:
                            try:
                                # Try DD/MM/YYYY format
                                parsed_date = datetime.strptime(raw_date, "%d/%m/%Y")
                                extracted_date = parsed_date.strftime("%d-%b-%Y")
                            except ValueError:
                                extracted_date = raw_date
            
            log_callback(f"[APEXGLOBAL] Invoice: {invoice_no if invoice_no else 'Not Found'}, Date: {extracted_date}, MAWB: {mawb}, HAWB: {hawb}, Total: {total_val}")
            results.append({
                "Organization": "APEXGLOBAL FORWARDERS INDIA PRIVATE LIMITED",
                "Vendor Inv No": invoice_no if invoice_no else "Not Found",
                "Vendor Inv Date": extracted_date,
                "Narration": "DO CHARGES PAYABLE TO APEXGLOBAL FORWARDERS",
                "Amount": str(round(total_val, 2)),
                "Charge or GL Amount": str(round(total_val, 2)),
                "WH Tax Taxable": str(round(taxable_val, 2)),
                "WH Tax Amount": str(round(wh_tax_amount, 2)),
                "Ref No": hawb if hawb else "Not Found",
                "MAWB No": mawb if mawb else "Not Found"
            })
        elif "ANIL MANTRA AVIATION" in text.upper():
            # ANIL MANTRA AVIATION PVT LTD extraction
            hawb = None
            mawb = None
            
            # Extract Invoice No - format: "Invoice No" followed by value like "MUM/2526/063"
            invoice_no_match = re.search(r"Invoice No\s*[:\s]*([A-Z]+/\d+/\d+)", text, re.IGNORECASE)
            if not invoice_no_match:
                # Alternative pattern
                invoice_no_match = re.search(r"Invoice No\s*[:\s]*([A-Z0-9/]+)", text, re.IGNORECASE)
            invoice_no = invoice_no_match.group(1) if invoice_no_match else None
            
            # Extract Date - format: "Date" followed by value like "30-Dec-25" or ":30-Dec-25"
            invoice_date = re.search(r"\bDate\s*:?\s*(\d{1,2}-[A-Za-z]{3}-\d{2,4})", text, re.IGNORECASE)
            if not invoice_date:
                # Alternative: look after Invoice No line
                invoice_date = re.search(r"Invoice No[^\n]*Date\s*:?\s*(\d{1,2}-[A-Za-z]{3}-\d{2,4})", text, re.IGNORECASE)
            
            # Extract Mawb/Hawb No - format: "Mawb/Hawb No" followed by value like ":312-9516-0833/SBKK00476263"
            # Some PDFs have spaces within the MAWB (e.g., "217-0854 8363/SBKK00496287")
            # Split into MAWB and HAWB
            mawb_hawb_match = re.search(r"Mawb/Hawb No\s*:?\s*([\d\-\s]+)/([A-Z0-9]+)", text, re.IGNORECASE)
            if mawb_hawb_match:
                mawb = mawb_hawb_match.group(1).replace(" ", "")
                hawb = mawb_hawb_match.group(2)
            else:
                # Try separate patterns
                mawb_match = re.search(r"Mawb[^\d]*([\d\-\s]{7,})", text, re.IGNORECASE)
                hawb_match = re.search(r"Hawb\s*(?:No)?[:\s]+([A-Z][A-Z0-9]+)", text, re.IGNORECASE)
                if mawb_match:
                    mawb = mawb_match.group(1).replace(" ", "")
                if hawb_match:
                    hawb = hawb_match.group(1)
            
            # Extract Grand Total - format: "Grand Total" followed by amount like "₹ 11,474.32" or "11,474.32"
            total = re.search(r"Grand Total\s*[₹Rs.\s]*([0-9,]+\.?\d*)", text, re.IGNORECASE)
            if not total:
                # Alternative: look for Grand Total with rupee symbol
                total = re.search(r"Grand Total.*?([0-9,]+\.\d{2})", text, re.IGNORECASE)
            
            # Extract Total Amount as taxable - format: "Total Amount" followed by amount like "9,724.00"
            taxable = re.search(r"Total Amount\s*[:\s]*([0-9,]+\.?\d*)", text, re.IGNORECASE)
            if not taxable:
                # Alternative: look for Taxable Value
                taxable = re.search(r"Taxable\s*Value\s*[:\s]*([0-9,]+\.?\d*)", text, re.IGNORECASE)
            if not taxable:
                # Alternative: look for "Total" row before Grand Total
                taxable = re.search(r"\bTotal\s+([0-9,]+\.?\d*)\s*\n.*Total Amount", text, re.IGNORECASE)
            
            total_val = float(total.group(1).replace(',', '')) if total else 0.0
            taxable_val = float(taxable.group(1).replace(',', '')) if taxable else 0.0
            wh_tax_amount = taxable_val * 0.02
            
            # Format date: 30-Dec-25 -> 30-Dec-2025
            extracted_date = "Not Found"
            if invoice_date:
                raw_date = invoice_date.group(1)
                try:
                    # Try DD-Mon-YY format
                    parsed_date = datetime.strptime(raw_date, "%d-%b-%y")
                    extracted_date = parsed_date.strftime("%d-%b-%Y")
                except ValueError:
                    try:
                        # Try DD-Mon-YYYY format
                        parsed_date = datetime.strptime(raw_date, "%d-%b-%Y")
                        extracted_date = parsed_date.strftime("%d-%b-%Y")
                    except ValueError:
                        extracted_date = raw_date
            
            log_callback(f"[ANIL MANTRA] Invoice: {invoice_no if invoice_no else 'Not Found'}, Date: {extracted_date}, MAWB: {mawb}, HAWB: {hawb}, Total: {total_val}")
            results.append({
                "Organization": "ANIL MANTRA AVIATION PVT LTD",
                "Vendor Inv No": invoice_no if invoice_no else "Not Found",
                "Vendor Inv Date": extracted_date,
                "Narration": "DO CHARGES PAYABLE TO ANIL MANTRA AVIATION",
                "Amount": str(round(total_val, 2)),
                "Charge or GL Amount": str(round(total_val, 2)),
                "WH Tax Taxable": str(round(taxable_val, 2)),
                "WH Tax Amount": str(round(wh_tax_amount, 2)),
                "Ref No": hawb if hawb else "Not Found",
                "MAWB No": mawb if mawb else "Not Found"
            })
    except Exception as e:
        log_callback(f"Regex extraction error: {e}")
    return results

# --- CSV Row Conversion Functions ---
MIAL_HEADER = [
    'Entry Date', 'Posting Date', 'Organization', 'Organization Branch', 'Vendor Inv No',
    'Vendor Inv Date', 'Currency', 'ExchRate', 'Narration', 'Due Date', 'Charge or GL',
    'Charge or GL Name', 'Charge or GL Amount', 'DR or CR', 'Cost Center', 'Branch',
    'Charge Narration', 'TaxGroup', 'Tax Type', 'SAC or HSN', 'Taxcode1', 'Taxcode1 Amt',
    'Taxcode2', 'Taxcode2 Amt', 'Taxcode3', 'Taxcode3 Amt', 'Taxcode4', 'Taxcode4 Amt',
    'Avail Tax Credit', 'LOB', 'Ref Type', 'Ref No', 'Amount', 'Start Date', 'End Date',
    'WH Tax Code', 'WH Tax Percentage', 'WH Tax Taxable', 'WH Tax Amount', 'Round Off',
    'CC Code'
]

def convert_date_format(date_str, out_fmt="%d-%b-%Y"):
    """Convert date string to DD-Mon-YYYY format.
    Handles both 2-digit and 4-digit year inputs.
    """
    if not date_str or date_str in ["Not Found", "NotFound", ""]:
        return date_str
    
    # Input formats - order matters: try 4-digit year first, then 2-digit
    date_formats = [
        "%d-%b-%Y",   # 17-Dec-2025 (4-digit year)
        "%d-%b-%y",   # 17-Dec-25 (2-digit year) -> converts to 2025
        "%d-%m-%Y",   # 17-12-2025
        "%d-%m-%y",   # 17-12-25
        "%d/%m/%Y",   # 17/12/2025
        "%d/%m/%y",   # 17/12/25
        "%d.%m.%Y",   # 17.12.2025
        "%d.%m.%y",   # 17.12.25
    ]
    
    for fmt in date_formats:
        try:
            parsed_date = datetime.strptime(date_str.strip(), fmt)
            return parsed_date.strftime(out_fmt)
        except ValueError:
            continue
    return date_str  # fallback: return as-is if not parseable

def mial_row_to_csv(row, today_date):
    row_dict = {col: "" for col in MIAL_HEADER}
    raw_date = row.get('Vendor Inv Date', '')
    norm_date = convert_date_format(raw_date, "%d-%b-%Y") if raw_date else ""
    row_dict.update({
        'Entry Date': today_date,
        'Posting Date': today_date,
        'Organization': row.get('Organization', ''),
        'Organization Branch': row.get('Branch', ''),
        'Vendor Inv No': row.get('Vendor Inv No', ''),
        'Vendor Inv Date': norm_date,
        'Currency': row.get('Currency', ''),
        'ExchRate': '1',
        'Narration': row.get('Narration', ''),
        'Charge or GL': 'Charge',
        'Charge or GL Name': row.get('Charge or GL Name', ''),
        'Charge or GL Amount': row.get('Total Amount', ''),
        'DR or CR': 'DR',
        'Branch': 'HO',
        'Charge Narration': '',
        'TaxGroup': 'GSTIN',
        'Tax Type': 'Pure Agent',
        'SAC or HSN': row.get('SAC or HSN', ''),
        'Avail Tax Credit': 'No',
        'LOB': 'CCL IMP',
        'Ref No': row.get('Ref No', ''),
        'Amount': row.get('Total Amount', ''),
        'Round Off': 'Yes'
    })
    return [row_dict[col] for col in MIAL_HEADER]

def airindia_row_to_csv(row, today_date):
    row_dict = {col: "" for col in MIAL_HEADER}
    raw_date = row.get('Vendor Inv Date', '')
    norm_date = convert_date_format(raw_date, "%d-%b-%Y") if raw_date else ""
    row_dict.update({
        'Entry Date': today_date,
        'Posting Date': today_date,
        'Organization': row.get('Organization', ''),
        'Organization Branch': row.get('Branch', ''),
        'Vendor Inv No': row.get('Vendor Inv No', ''),
        'Vendor Inv Date': norm_date,
        'Currency': row.get('Currency', ''),
        'ExchRate': '1',
        'Narration': row.get('Narration', ''),
        'Charge or GL': 'Charge',
        'Charge or GL Name': row.get('Charge or GL Name', ''),
        'Charge or GL Amount': row.get('Total Amount', ''),
        'DR or CR': 'DR',
        'Branch': 'HO',
        'Charge Narration': '',
        'TaxGroup': 'Pure Agent',
        'Tax Type': 'Non-Taxable',
        'SAC or HSN': row.get('SAC or HSN', ''),
        'Avail Tax Credit': 'No',
        'LOB': 'CCL IMP',
        'Ref No': row.get('Ref No', ''),
        'Amount': row.get('Total Amount', ''),
        'Round Off': 'Yes'
    })
    return [row_dict[col] for col in MIAL_HEADER]

def do_row_to_csv(row, today_date):
    row_dict = {col: "" for col in MIAL_HEADER}
    raw_date = row.get('Vendor Inv Date', '')
    norm_date = convert_date_format(raw_date, "%d-%b-%Y") if raw_date else ""
    narration = row.get('Narration', '')
    ref_no = row.get('Ref No', '')
    if ref_no and ref_no != 'No match found':
        narration = f"{narration} / {ref_no}"
    
    # Determine WH Tax Percentage based on organization
    org = row.get('Organization', '').upper()
    if "OCEAN NETWORK EXPRESS" in org or "DACHSER" in org:
        wh_tax_percentage = "0"
        wh_tax_code = ""
    else:
        wh_tax_percentage = "2"
        wh_tax_code = "194C"
    
    # Determine Organization Branch based on organization
    if "SENATOR" in org:
        org_branch = "CHENNAI"
    elif "APEXGLOBAL" in org or "APEX GLOBAL" in org:
        org_branch = "TAMIL NADU"
    else:
        org_branch = "MUMBAI"
    
    row_dict.update({
        'Entry Date': today_date,
        'Posting Date': today_date,
        'Organization': row.get('Organization', ''),
        'Organization Branch': org_branch,
        'Vendor Inv No': row.get('Vendor Inv No', ''),
        'Vendor Inv Date': norm_date,
        'Currency': "INR",
        'ExchRate': "1",
        'Narration': narration,
        'Charge or GL': "CHARGE",
        'Charge or GL Name': "DELIVERY ORDER CHARGES (1)",
        'Charge or GL Amount': row.get('Charge or GL Amount', ''),
        'DR or CR': "DR",
        'Branch': "HO",
        'Avail Tax Credit': "No",
        'LOB': "CCL IMP",
        'Ref No': row.get('Ref No', ''),
        'Amount': row.get('Amount', ''),
        'WH Tax Code': wh_tax_code,
        'WH Tax Percentage': wh_tax_percentage,
        'WH Tax Taxable': row.get('WH Tax Taxable', ''),
        'WH Tax Amount': row.get('WH Tax Amount', ''),
        'Round Off': "yes"
    })
    return [row_dict[col] for col in MIAL_HEADER]

# --- Nagarkot Brand Color Palette ---
BG_COLOR = "#F4F6F8"
CARD_BG = "#FFFFFF"
ACCENT = "#1F3F6E"
ACCENT_HOVER = "#2A528F"
TEXT_PRIMARY = "#1E1E1E"
TEXT_SECONDARY = "#6B7280"
BORDER_COLOR = "#E5E7EB"
ERROR_RED = "#D8232A"
LOG_BG = "#FAFBFC"
LOG_FG = "#1E1E1E"

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and PyInstaller."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- GUI ---
class IntegratedDOInvoiceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Integrated DO + Invoice to CSV Converter")
        try:
            self.root.state("zoomed")
        except Exception:
            self.root.attributes("-fullscreen", True)
        self.root.configure(bg=BG_COLOR)

        self.pdf_paths = []
        self.job_register_path = None
        self.ref_no_mapping = None
        self._logo_image = None

        self._setup_styles()
        self._create_widgets()

        # Logging setup
        text_handler = TextHandler(self.log_text)
        text_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(text_handler)

    def _setup_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Card.TLabelframe", background=CARD_BG, borderwidth=1, relief="solid")
        style.configure("Card.TLabelframe.Label", background=CARD_BG, foreground=TEXT_PRIMARY, font=("Segoe UI", 10, "bold"))
        style.configure("Modern.TButton", font=("Segoe UI", 9), padding=(14, 6), background=CARD_BG, borderwidth=1, relief="solid")
        style.map("Modern.TButton", background=[("active", "#F5F5F5"), ("pressed", "#EEEEEE")])
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(20, 8), foreground="#FFFFFF", background=ACCENT, borderwidth=0)
        style.map("Accent.TButton", background=[("active", ACCENT_HOVER), ("pressed", ACCENT_HOVER), ("disabled", "#90CAF9")], foreground=[("disabled", "#FFFFFF")])

    def _create_widgets(self):
        main_frame = tk.Frame(self.root, bg=BG_COLOR)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- HEADER ---
        header_frame = tk.Frame(main_frame, bg=CARD_BG, pady=16, padx=24)
        header_frame.pack(fill=tk.X)
        tk.Frame(main_frame, bg=BORDER_COLOR, height=1).pack(fill=tk.X)

        logo_path = resource_path("logo.png")
        if HAS_PIL and os.path.isfile(logo_path):
            try:
                img = Image.open(logo_path)
                h = 40
                w = int(img.width * h / img.height)
                img = img.resize((w, h), Image.LANCZOS)
                self._logo_image = ImageTk.PhotoImage(img)
                tk.Label(header_frame, image=self._logo_image, bg=CARD_BG).pack(side=tk.LEFT)
            except Exception:
                tk.Label(header_frame, text="NAGARKOT", font=("Segoe UI", 12, "bold"), fg=ACCENT, bg=CARD_BG).pack(side=tk.LEFT)
        else:
            tk.Label(header_frame, text="NAGARKOT", font=("Segoe UI", 12, "bold"), fg=ACCENT, bg=CARD_BG).pack(side=tk.LEFT)

        tk.Label(
            header_frame, text="Integrated DO + Invoice to CSV Converter",
            font=("Segoe UI", 16, "bold"), bg=CARD_BG, fg=TEXT_PRIMARY,
        ).place(relx=0.5, rely=0.3, anchor="center")
        tk.Label(
            header_frame, text="Extract Invoice & Delivery Order data from PDFs into a single CSV",
            font=("Segoe UI", 9), bg=CARD_BG, fg=TEXT_SECONDARY,
        ).place(relx=0.5, rely=0.75, anchor="center")

        # --- BODY ---
        body = tk.Frame(main_frame, bg=BG_COLOR, padx=40, pady=30)
        body.pack(fill=tk.BOTH, expand=True)

        # File Selection Card
        file_card = ttk.LabelFrame(body, text="  Input Files  ", style="Card.TLabelframe", padding=20)
        file_card.pack(fill=tk.X, pady=(0, 20))
        file_inner = tk.Frame(file_card, bg=CARD_BG)
        file_inner.pack(fill=tk.BOTH, expand=True)

        self.job_status_label = tk.Label(file_inner, text="Job Register: Not Selected", fg=TEXT_SECONDARY, bg=CARD_BG, font=("Segoe UI", 9))
        self.job_status_label.pack(anchor=tk.W, pady=(0, 5))
        self.pdf_status_label = tk.Label(file_inner, text="Invoice/DO PDFs: Not Selected", fg=TEXT_SECONDARY, bg=CARD_BG, font=("Segoe UI", 9))
        self.pdf_status_label.pack(anchor=tk.W, pady=(0, 15))

        btn_frame = tk.Frame(file_inner, bg=CARD_BG)
        btn_frame.pack(fill=tk.X)
        ttk.Button(btn_frame, text="Select Job Register", command=self.select_job, style="Modern.TButton").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="Select PDFs (Invoice/DO)", command=self.select_pdfs, style="Modern.TButton").pack(side=tk.LEFT)

        # Action Area
        action_frame = tk.Frame(body, bg=BG_COLOR)
        action_frame.pack(fill=tk.X, pady=(0, 20))
        self.process_button = ttk.Button(action_frame, text="\u25B6  Process & Generate CSV", command=self.process_files, style="Accent.TButton")
        self.process_button.pack(side=tk.LEFT, padx=(0, 20))
        self.status_label = tk.Label(action_frame, text="Ready", fg=TEXT_SECONDARY, bg=BG_COLOR, font=("Segoe UI", 9), wraplength=600, anchor="w", justify="left")
        self.status_label.pack(side=tk.LEFT, fill=tk.X)

        # Log Card
        log_card = ttk.LabelFrame(body, text="  Processing Log  ", style="Card.TLabelframe", padding=15)
        log_card.pack(fill=tk.BOTH, expand=True)
        log_inner = tk.Frame(log_card, bg=CARD_BG)
        log_inner.pack(fill=tk.BOTH, expand=True)
        self.log_text = scrolledtext.ScrolledText(
            log_inner, height=10, wrap=tk.WORD, state="disabled",
            bg=LOG_BG, fg=LOG_FG, font=("Consolas", 9),
            relief="flat", padx=10, pady=10,
        )
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # --- FOOTER ---
        footer_frame = tk.Frame(main_frame, bg=CARD_BG, padx=24, pady=10)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        tk.Frame(main_frame, bg=BORDER_COLOR, height=1).pack(fill=tk.X, side=tk.BOTTOM)
        tk.Label(footer_frame, text="Nagarkot Forwarders Pvt. Ltd. \u00A9", fg=TEXT_SECONDARY, bg=CARD_BG, font=("Segoe UI", 8)).pack(side=tk.LEFT)
        ttk.Button(footer_frame, text="Exit", command=self.root.destroy, style="Modern.TButton").pack(side=tk.RIGHT)

    def log(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}: {message}\n")
        self.log_text.config(state='disabled')
        self.log_text.see(tk.END)
        self.root.update()

    def select_job(self):
        filetypes = [
            ("CSV and Excel files", "*.csv;*.xls;*.xlsx"),
            ("CSV files", "*.csv"),
            ("Excel files", "*.xls;*.xlsx")
        ]
        file_path = filedialog.askopenfilename(filetypes=filetypes)
        if not file_path:
            self.status_label.config(text="No Job Register file selected.", fg=TEXT_SECONDARY)
            self.log("No Job Register file selected.")
            logger.info("No Job Register file selected.")
            return
        self.job_register_path = file_path
        # Load job register (CSV or Excel)
        ext = os.path.splitext(file_path)[1].lower()
        try:
            if ext == '.csv':
                self.ref_no_mapping = load_job_register(file_path, self.log)
            elif ext in ['.xls', '.xlsx']:
                self.log(f"Loading job register (Excel): {os.path.basename(file_path)}")
                df = pd.read_excel(file_path, dtype=str)
                # Save as temp CSV for compatibility with rest of code
                temp_csv = os.path.splitext(file_path)[0] + "_temp_job_register.csv"
                df.to_csv(temp_csv, index=False)
                self.ref_no_mapping = load_job_register(temp_csv, self.log)
                # Optionally, delete temp_csv after loading if you want
            else:
                self.status_label.config(text="Unsupported file type selected.")
                self.log("Unsupported file type selected.")
                logger.error("Unsupported file type selected.")
                return
            self.status_label.config(text=f"Job Register selected: {os.path.basename(file_path)}", fg=TEXT_PRIMARY)
            self.job_status_label.config(text=f"Job Register: {os.path.basename(file_path)}", fg=TEXT_PRIMARY)
            self.log(f"Job Register selected: {os.path.basename(file_path)}")
            logger.info(f"Job Register selected: {file_path}")
        except Exception as e:
            self.status_label.config(text=f"Error loading job register: {e}", fg=ERROR_RED)
            self.log(f"Error loading job register: {e}")
            logger.error(f"Error loading job register: {e}")

    def select_pdfs(self):
        pdf_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if not pdf_paths:
            self.status_label.config(text="No PDFs selected.", fg=TEXT_SECONDARY)
            self.log("No PDFs selected.")
            logger.info("No PDFs selected.")
            return
        self.pdf_paths = pdf_paths
        self.status_label.config(text=f"Selected {len(pdf_paths)} PDFs", fg=TEXT_PRIMARY)
        self.pdf_status_label.config(text=f"Invoice/DO PDFs: {len(pdf_paths)} selected", fg=TEXT_PRIMARY)
        self.log(f"Selected {len(pdf_paths)} PDFs: {', '.join([os.path.basename(path) for path in pdf_paths])}")
        logger.info(f"Selected {len(pdf_paths)} PDFs: {', '.join([os.path.basename(path) for path in pdf_paths])}")

    def process_files(self):
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        logger.info("Starting integrated file processing")
        if not self.pdf_paths:
            messagebox.showerror("Error", "Please select at least one PDF")
            self.status_label.config(text="No PDFs selected.", fg=ERROR_RED)
            self.log("No PDFs selected.")
            logger.error("No PDFs selected")
            return
        if not self.ref_no_mapping:
            messagebox.showerror("Error", "Please select a Job Register CSV")
            self.status_label.config(text="No Job Register CSV selected.", fg=ERROR_RED)
            self.log("No Job Register CSV selected.")
            logger.error("No Job Register CSV selected")
            return
        # Load job register for DO job number mapping
        job_register = []
        try:
            df = pd.read_csv(self.job_register_path, dtype=str)
            df.columns = [c.strip().lower() for c in df.columns]
            job_col = 'job no'
            hawb_col = 'hawb/hbl no'
            mawb_col = 'awb/bl no.'
            type_col = 'type of b/e'
            for _, row in df.iterrows():
                job_register.append({
                    'job_no': str(row.get(job_col, '')).strip(),
                    'hawb': str(row.get(hawb_col, '')).replace(' ','').strip().replace('-','').replace('/','').upper(),
                    'mawb': str(row.get(mawb_col, '')).replace(' ','').strip().replace('-','').replace('/','').upper(),
                    'type': str(row.get(type_col, '')).strip().upper()
                })
        except Exception as e:
            self.log(f"Failed to load job register for DO mapping: {e}")
            job_register = []
        def match_job_no(hawb, mawb):
            def norm(x):
                return str(x).replace(' ','').replace('-','').replace('/','').upper()
            hawb = norm(hawb)
            mawb = norm(mawb)
            def prefer_sez_z(matches):
                sez_z_matches = [r for r in matches if r['type'] == 'SEZ-Z']
                if sez_z_matches:
                    return sez_z_matches[0]['job_no']
                return 'No match found'
            matches = [r for r in job_register if r['hawb'] == hawb]
            if matches:
                if len(matches) == 1:
                    return matches[0]['job_no']
                return prefer_sez_z(matches)
            matches = [r for r in job_register if r['mawb'] == hawb]
            if matches:
                if len(matches) == 1:
                    return matches[0]['job_no']
                return prefer_sez_z(matches)
            matches = [r for r in job_register if r['mawb'] == mawb]
            if matches:
                if len(matches) == 1:
                    return matches[0]['job_no']
                return prefer_sez_z(matches)
            return 'No match found'
        output_dir = os.path.join(os.getcwd(), 'CSV_Output')
        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%d-%m-%y_%H-%M")
        output_csv = os.path.join(output_dir, f"Integrated_{timestamp}.csv")
        today_date = datetime.now().strftime("%d-%b-%Y")
        all_rows = []
        for pdf_path in self.pdf_paths:
            self.log(f"Processing PDF: {os.path.basename(pdf_path)}")
            logger.info(f"Processing PDF: {pdf_path}")
            text, tables_data = extract_text_from_pdf(pdf_path, self.log)
            if text is None:
                self.log(f"Failed to process {os.path.basename(pdf_path)}")
                logger.warning(f"Failed to process {pdf_path}")
                continue
            invoice_type = detect_invoice_type(text)
            if invoice_type == "MIAL":
                rows = extract_mial_fields(text, self.ref_no_mapping, self.log)
                for row in rows:
                    all_rows.append(mial_row_to_csv(row, today_date))
            elif invoice_type == "AirIndia":
                rows = extract_airindia_fields(text, self.ref_no_mapping, self.log)
                for row in rows:
                    if 'SAC or Service' in row:
                        row['SAC or HSN'] = row['SAC or Service']
                    all_rows.append(airindia_row_to_csv(row, today_date))
            else:
                do_rows = extract_invoice_details_with_regex(text, tables_data, self.log)
                self.log(f"DO extraction result: {do_rows}")
                for row in do_rows:
                    hawb = row.get("Ref No", "")
                    mawb = row.get("MAWB No", "")
                    mapped_job_no = match_job_no(hawb, mawb)
                    self.log(f"Mapped job no for HAWB={hawb}, MAWB={mawb}: {mapped_job_no}")
                    row["Ref No"] = mapped_job_no
                    all_rows.append(do_row_to_csv(row, today_date))
        if not all_rows:
            messagebox.showerror("Error", "No valid data extracted from PDFs.")
            self.status_label.config(text="No valid data extracted from PDFs.", fg=ERROR_RED)
            self.log("No valid data extracted from PDFs.")
            logger.error("No valid data extracted")
            return
        with open(output_csv, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(MIAL_HEADER)
            for row in all_rows:
                writer.writerow(row)
        self.status_label.config(text=f"CSV generated: {os.path.basename(output_csv)}", fg=ACCENT)
        self.log(f"CSV generated: {os.path.basename(output_csv)}")
        messagebox.showinfo("Success", f"CSV generated: {output_csv}")
        logger.info(f"CSV generated: {output_csv}")

def get_base_path():
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle (PyInstaller)
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

if __name__ == "__main__":
    try:
        logger.info("Starting Integrated DO + Invoice to CSV Converter")
        root = tk.Tk()
        app = IntegratedDOInvoiceApp(root)
        root.mainloop()
        logger.info("Application closed")
    except Exception as e:
        logger.error(f"Application error: {e}")
        messagebox.showerror("Error", f"Application error: {e}")
