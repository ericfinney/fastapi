import os
import uuid
import json
import logging
import re
import math
import time
import tempfile
import pandas as pd
from copy import copy
from typing import Dict, Any, List, Optional, Tuple
from xml.etree import ElementTree as ET

from fastapi import FastAPI, File, UploadFile, Body, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel
import httpx

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Protection

# Import PDF extraction functionality
import pdfplumber

logging.basicConfig(level=logging.INFO)

TEMPLATE_PATH = os.environ.get("BOYD_TEMPLATE_PATH", "Blank.xlsx")
SHEET_NAME = os.environ.get("BOYD_SHEET_NAME", "Proposal")
OUTPUT_DIR = os.environ.get("BOYD_OUTPUT_DIR", "/tmp/output")
LOGO_PATH = os.environ.get("BOYD_LOGO_PATH", "assets/logo.png")

os.makedirs(OUTPUT_DIR, exist_ok=True)

app = FastAPI()


# =========================================================
# URL request models
# =========================================================
class PdfUrlRequest(BaseModel):
    file_url: str


async def _download_pdf_bytes(file_url: str) -> bytes:
    """
    Downloads a PDF from a URL and returns raw bytes.
    This is used when the GPT Action platform can't send multipart file uploads.
    """
    if not file_url or not isinstance(file_url, str):
        raise HTTPException(status_code=400, detail="file_url must be a non-empty string")

    # Basic sanity check (optional)
    if not (file_url.startswith("http://") or file_url.startswith("https://")):
        raise HTTPException(status_code=400, detail="file_url must start with http:// or https://")

    try:
        async with httpx.AsyncClient(timeout=120, follow_redirects=True) as client:
            resp = await client.get(file_url)
            resp.raise_for_status()
            content_type = resp.headers.get("content-type", "").lower()

            pdf_bytes = resp.content
            if not pdf_bytes or len(pdf_bytes) < 1000:
                raise HTTPException(status_code=400, detail="Downloaded file appears empty or invalid")

            # If content-type isn't pdf, still allow as long as bytes look like a PDF
            # PDF signature usually starts with %PDF
            if "pdf" not in content_type and not pdf_bytes.startswith(b"%PDF"):
                raise HTTPException(
                    status_code=400,
                    detail=f"Downloaded content does not appear to be a PDF (content-type: {content_type})"
                )
            return pdf_bytes

    except httpx.HTTPStatusError as e:
        raise HTTPException(status_code=400, detail=f"Failed to download file_url (HTTP): {str(e)}")
    except httpx.RequestError as e:
        raise HTTPException(status_code=400, detail=f"Failed to download file_url (network): {str(e)}")


# =========================================================
# PDF EXTRACTION CLASSES (from extract_estimate_data.py)
# =========================================================
class EstimateExtractor:
    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self.header_data = {}
        self.sign_items = []
        self.totals = {}
        self.current_section = 'Unknown'
    
    def clean_price(self, price_str: str) -> float:
        """Convert price string to float"""
        if not price_str or price_str.strip() == '':
            return 0.0
        cleaned = price_str.replace('$', '').replace(',', '').strip()
        try:
            return float(cleaned)
        except ValueError:
            return 0.0
    
    def extract_header_info(self, first_page_text: str) -> Dict:
        """Extract header information from first page"""
        header = {}
        
        # Extract To: Company
        to_company_match = re.search(r'To:\s+([^\n]+?)(?=Ship To:|$)', first_page_text)
        if to_company_match:
            to_text = to_company_match.group(1).strip()
            header['To_Company'] = to_text.split('Ship To:')[0].strip()
        
        # Extract Ship To: Location
        ship_to_match = re.search(r'Ship To:\s+([^\n]+)', first_page_text)
        if ship_to_match:
            header['Ship_To_Address'] = ship_to_match.group(1).strip()
        
        # Extract Project ID
        proj_id_match = re.search(r'Project Id\s*:\s*(\d+)', first_page_text)
        if proj_id_match:
            header['Project_ID'] = proj_id_match.group(1)
        
        # Extract Project Description
        proj_desc_match = re.search(r'Project Desc\.\s*:\s*(.+?)(?=Ship Via)', first_page_text)
        if proj_desc_match:
            header['Project_Description'] = proj_desc_match.group(1).strip()
        
        # Extract Ship Via
        ship_via_match = re.search(r'Ship Via\s*:\s*(.+?)(?=\n|P\.O\.)', first_page_text)
        if ship_via_match:
            header['Ship_Via'] = ship_via_match.group(1).strip()
        
        # Extract Sales Rep
        sales_match = re.search(r'Salesperson\s*:\s*(.+?)(?=\n|$)', first_page_text)
        if sales_match:
            header['Sales_Rep'] = sales_match.group(1).strip()
        
        # Extract Project Manager
        pm_match = re.search(r'Project Manager:\s*(.+?)(?=\n|$)', first_page_text)
        if pm_match:
            header['Project_Manager'] = pm_match.group(1).strip()
        
        # Extract Project Engineer
        pe_match = re.search(r'Project Engineer:\s*(.+?)(?=\n|$)', first_page_text)
        if pe_match:
            header['Project_Engineer'] = pe_match.group(1).strip()
        
        # Extract Date
        date_match = re.search(r'Date\s+(\d{1,2}/\d{1,2}/\d{2,4})', first_page_text)
        if date_match:
            header['Estimate_Date'] = date_match.group(1)
        
        return header
    
    def extract_sign_items_from_text(self, text: str, current_section: str = '') -> List[Dict]:
        """Extract sign items from text using pattern matching"""
        items = []
        
        lines = text.split('\n')
        
        for line in lines:
            line = line.strip()
            
            # Match the price section at the end: QTY $ PRICE $ TOTAL
            price_match = re.search(r'(\d+)\s+\$\s*([\d,]+\.\d+)\s+\$\s*([\d,]+\.\d+)$', line)
            
            if price_match:
                qty = int(price_match.group(1))
                unit_price = self.clean_price(price_match.group(2))
                total = self.clean_price(price_match.group(3))
                
                # Extract the code and description (everything before the price section)
                prefix = line[:price_match.start()].strip()
                
                # Match CODE - DESCRIPTION
                # CODE can be: letters, optionally periods and more letters, then digits, then optionally more letters/digits/periods
                code_match = re.match(r'^([a-zA-Z]+[\.a-zA-Z]*\d+[a-zA-Z0-9\.]*)\s+-\s+(.+)$', prefix)
                
                if code_match:
                    sign_code = code_match.group(1)
                    description = code_match.group(2).strip()
                    
                    # Skip installation, packing, shipping, and permitting items
                    if any(keyword in description.lower() for keyword in ['installation', 'packing', 'crating', 'shipping', 'permitting']):
                        continue
                    
                    items.append({
                        'Section': current_section,
                        'Sign_Type': sign_code,
                        'Description': description,
                        'Quantity': qty,
                        'Unit_Price': unit_price,
                        'Extended_Total': total
                    })
        
        return items
    
    def extract_cost_items_from_text(self, text: str, current_section: str = '') -> List[Dict]:
        """Extract packing, shipping, mobilization, installation, and permitting as line items"""
        items = []
        
        # Packing/Crating
        packing_match = re.search(r'Packing/Crating\s+(\d+)\s+\$\s*([\d,]+\.\d+)\s+\$\s*([\d,]+\.\d+)', text)
        if packing_match:
            items.append({
                'Section': current_section,
                'Sign_Type': 'PACKING',
                'Description': 'Packing/Crating',
                'Quantity': int(packing_match.group(1)),
                'Unit_Price': self.clean_price(packing_match.group(2)),
                'Extended_Total': self.clean_price(packing_match.group(3))
            })
        
        # Shipping
        shipping_match = re.search(r'Shipping\s+(\d+)\s+\$\s*([\d,]+\.\d+)\s+\$\s*([\d,]+\.\d+)', text)
        if shipping_match:
            items.append({
                'Section': current_section,
                'Sign_Type': 'SHIPPING',
                'Description': 'Shipping',
                'Quantity': int(shipping_match.group(1)),
                'Unit_Price': self.clean_price(shipping_match.group(2)),
                'Extended_Total': self.clean_price(shipping_match.group(3))
            })
        
        # Mobilization
        mob_match = re.search(r'Installation - Mobilization & Other Expenses\s+(\d+)\s*\$\s*([\d,]+\.\d+)\s+\$\s*([\d,]+\.\d+)', text)
        if mob_match:
            items.append({
                'Section': current_section,
                'Sign_Type': 'MOBILIZATION',
                'Description': 'Installation - Mobilization & Other Expenses',
                'Quantity': int(mob_match.group(1)),
                'Unit_Price': self.clean_price(mob_match.group(2)),
                'Extended_Total': self.clean_price(mob_match.group(3))
            })
        
        # Installation Labor (Boyd)
        labor_match = re.search(r'Installation - Onsite Labor\s+(\d+)\s+\$\s*([\d,]+\.\d+)\s+\$\s*([\d,]+\.\d+)', text)
        if labor_match:
            items.append({
                'Section': current_section,
                'Sign_Type': 'INSTALLATION',
                'Description': 'Installation - Onsite Labor',
                'Quantity': int(labor_match.group(1)),
                'Unit_Price': self.clean_price(labor_match.group(2)),
                'Extended_Total': self.clean_price(labor_match.group(3))
            })
        
        # More Vang Installation (or other subcontractor)
        more_vang_patterns = [
            r'Full.*?Installation by More Vang\s+(\d+)\s+\$\s*([\d,]+\.\d+)\s+\$\s*([\d,]+\.\d+)',
            r'Installation - Onsite Labor by.*?\s+(\d+)\s+\$\s*([\d,]+\.\d+)\s+\$\s*([\d,]+\.\d+)'
        ]
        
        for pattern in more_vang_patterns:
            more_vang_match = re.search(pattern, text)
            if more_vang_match:
                desc_match = re.search(r'(Full.*?Installation by [^\n]+)', text)
                description = desc_match.group(1) if desc_match else 'Subcontractor Installation'
                
                items.append({
                    'Section': current_section,
                    'Sign_Type': 'INSTALLATION',
                    'Description': description,
                    'Quantity': int(more_vang_match.group(1)),
                    'Unit_Price': self.clean_price(more_vang_match.group(2)),
                    'Extended_Total': self.clean_price(more_vang_match.group(3))
                })
                break
        
        # Permitting
        permitting_match = re.search(r'Permitting\s+(\d+)\s+\$\s*([\d,]+\.\d+)\s+\$\s*([\d,]+\.\d+)', text)
        if permitting_match:
            items.append({
                'Section': current_section,
                'Sign_Type': 'PERMITTING',
                'Description': 'Permitting',
                'Quantity': int(permitting_match.group(1)),
                'Unit_Price': self.clean_price(permitting_match.group(2)),
                'Extended_Total': self.clean_price(permitting_match.group(3))
            })
        
        return items
    
    def extract_totals(self, last_page_text: str) -> Dict:
        """Extract final totals from last page"""
        totals = {}
        
        # Sub-total
        subtotal_match = re.search(r'SUB-TOTAL\s+\$\s*([\d,]+\.\d{2})', last_page_text)
        if subtotal_match:
            totals['Sub_Total'] = self.clean_price(subtotal_match.group(1))
        
        # Mark-up
        markup_match = re.search(r'MARK-UP\s+\$\s*([\d,]+\.\d{2})', last_page_text)
        if markup_match:
            totals['Mark_Up'] = self.clean_price(markup_match.group(1))
        
        # Total (final) - use negative lookbehind to avoid SUB-TOTAL
        total_match = re.search(r'(?<!SUB-)TOTAL\s+\$\s*([\d,]+\.\d{2})', last_page_text)
        if total_match:
            totals['Grand_Total'] = self.clean_price(total_match.group(1))
        
        return totals
    
    def process_pdf(self):
        """Main method to process the PDF"""
        with pdfplumber.open(self.pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                
                # Extract header from first page
                if page_num == 0:
                    self.header_data = self.extract_header_info(text)
                
                # Process line by line for section tracking
                lines = text.split('\n')
                page_section = self.current_section
                
                for line_num, line in enumerate(lines):
                    # Check if this line is a section marker
                    # Check more specific patterns first
                    line_upper = line.upper()
                    new_section = None
                    
                    if 'SIGNAGE SAMPLES PACKAGE' in line_upper:
                        new_section = 'Signage Samples Package'
                    elif 'ECCLES - INTERIOR SIGNAGE PACKAGE' in line_upper:
                        new_section = 'Eccles - Interior Signage Package'
                    elif '1951 - INTERIOR SIGNAGE PACKAGE' in line_upper:
                        new_section = '1951 - Interior Signage Package'
                    elif '1951 - PARKING GARAGE SIGNAGE PACKAGE' in line_upper:
                        new_section = '1951 - Parking Garage Signage Package'
                    elif 'ECCLES - GRAPHICS SIGNAGE PACKAGE' in line_upper:
                        new_section = 'Eccles - Graphics Signage Package'
                    elif '1951 - GRAPHICS SIGNAGE PACKAGE' in line_upper:
                        new_section = '1951 - Graphics Signage Package'
                    elif 'SITE - SIGNAGE PACKAGE' in line_upper:
                        new_section = 'Site - Signage Package'
                    elif 'SIGNAGE PACKAGE' in line_upper and 'SAMPLES' not in line_upper:
                        new_section = 'Signage Package'
                    
                    if new_section:
                        page_section = new_section
                        self.current_section = new_section
                        continue
                    
                    # Try to extract sign item from this line
                    items = self.extract_sign_items_from_text(line, page_section)
                    self.sign_items.extend(items)
                    
                    # Try to extract cost item from this line
                    cost_items = self.extract_cost_items_from_text(line, page_section)
                    self.sign_items.extend(cost_items)
                
                # Extract totals from last page
                if page_num == len(pdf.pages) - 1:
                    self.totals = self.extract_totals(text)
    
    def to_json_format(self) -> Dict:
        """Convert extracted data to the JSON format expected by main.py"""
        # Filter out cost items (PACKING, SHIPPING, MOBILIZATION, INSTALLATION, PERMITTING)
        cost_types = {'PACKING', 'SHIPPING', 'MOBILIZATION', 'INSTALLATION', 'PERMITTING'}
        
        sign_items = [item for item in self.sign_items if item['Sign_Type'] not in cost_types]
        shipping_items = [item for item in self.sign_items if item['Sign_Type'] == 'SHIPPING']
        installation_items = [item for item in self.sign_items if item['Sign_Type'] in {'INSTALLATION', 'MOBILIZATION'}]
        
        # Format for the Excel generator
        result = {
            "estimate_date": self.header_data.get('Estimate_Date', ''),
            "project_id": self.header_data.get('Project_ID', ''),
            "salesperson": self.header_data.get('Sales_Rep', ''),
            "project_manager": self.header_data.get('Project_Manager', ''),
            "project_description": self.header_data.get('Project_Description', ''),
            "sold_to": {
                "name": self.header_data.get('To_Company', ''),
                "address_lines": [],
                "city": "",
                "state": "",
                "zip": "",
                "phone": ""
            },
            "ship_to": {
                "name": self.header_data.get('Ship_To_Address', ''),
                "address_lines": [],
                "city": "",
                "state": "",
                "zip": "",
                "phone": ""
            },
            "sign_types": [
                {
                    "sign_type": f"{item['Sign_Type']} - {item['Description']}",
                    "description": item['Description'],
                    "qty": item['Quantity'],
                    "unit_price": item['Unit_Price']
                }
                for item in sign_items
            ],
            "shipping": [
                {
                    "description": item['Description'],
                    "extended": item['Extended_Total']
                }
                for item in shipping_items
            ],
            "installation": [
                {
                    "description": item['Description'],
                    "extended": item['Extended_Total']
                }
                for item in installation_items
            ],
            "totals": {
                "subtotal": self.totals.get('Sub_Total', 0),
                "markup": self.totals.get('Mark_Up', 0),
                "total": self.totals.get('Grand_Total', 0)
            }
        }
        
        return result


# =========================================================
# Cleanup: delete old generated workbooks (older than N minutes)
# =========================================================
def cleanup_old_generated_workbooks(
    directory: str,
    older_than_minutes: int = 30,
    filename_prefix: str = "Boyd_Proposal_",
    allowed_exts: Tuple[str, ...] = (".xlsx", ".xlsm"),
) -> None:
    if not os.path.isdir(directory):
        return

    now = time.time()
    max_age_sec = older_than_minutes * 60

    for fn in os.listdir(directory):
        fn_lower = fn.lower()
        if not fn.startswith(filename_prefix):
            continue
        if not fn_lower.endswith(tuple(ext.lower() for ext in allowed_exts)):
            continue

        path = os.path.join(directory, fn)
        if not os.path.isfile(path):
            continue

        try:
            age_sec = now - os.path.getmtime(path)
            if age_sec > max_age_sec:
                os.remove(path)
                logging.info("Deleted old output workbook: %s (age %.1f min)", fn, age_sec / 60.0)
        except Exception as e:
            logging.warning("Failed deleting %s: %s", fn, e)


# =========================================================
# Basic helpers
# =========================================================
def safe_str(x) -> str:
    return "" if x is None else str(x)

def safe_num(x):
    try:
        if x is None or x == "":
            return None
        return float(x)
    except Exception:
        return None

def round_up_dollars(value):
    if value is None or value == "":
        return None
    try:
        return math.ceil(float(value))
    except Exception:
        return None

def round_nearest_dollar(value):
    if value is None or value == "":
        return None
    try:
        return round(float(value))
    except Exception:
        return None

def join_address_lines(addr_lines: List[str]) -> str:
    return "\n".join([line for line in addr_lines if line and line.strip()])

def write_cell(ws, cell: str, value):
    ws[cell].value = value

def insert_logo(ws):
    if not os.path.exists(LOGO_PATH):
        logging.warning("Logo not found at %s; skipping insert.", LOGO_PATH)
        return
    img = XLImage(LOGO_PATH)
    ws.add_image(img, "A1")


# =========================================================
# Lock/Unlock helpers
# =========================================================
def lock_cell(ws, cell_ref: str):
    ws[cell_ref].protection = Protection(locked=True)

def unlock_cell(ws, cell_ref: str):
    ws[cell_ref].protection = Protection(locked=False)

def lock_all_cells(ws, max_row: int, max_col: int) -> None:
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(row=r, column=c).protection = Protection(locked=True)


# =========================================================
# Footer row height capture/restore
# =========================================================
def capture_row_heights(ws, start_row: int, end_row: int) -> dict:
    heights = {}
    for r in range(start_row, end_row + 1):
        heights[r] = ws.row_dimensions[r].height
    return heights

def restore_row_heights(ws, heights: dict, row_offset: int):
    for original_row, height in heights.items():
        target_row = original_row + row_offset
        ws.row_dimensions[target_row].height = height


# =========================================================
# Sign type + summary split (ROBUST)
# =========================================================
TYPE_DESC_SPLIT_RE = re.compile(r"\s*[-–—]\s*", flags=re.UNICODE)

def looks_like_sign_code(code: str) -> bool:
    if not code:
        return False
    c = code.strip()
    if not c or len(c) > 60:
        return False
    if not re.match(r"^[A-Za-z0-9]", c):
        return False
    if not re.match(r"^[A-Za-z0-9./&_ ,]+$", c):
        return False
    return True

def split_sign_type_and_summary(raw_sign_type: str) -> Tuple[str, str]:
    if not raw_sign_type:
        return "", ""
    s = str(raw_sign_type).strip()
    if not s:
        return "", ""

    parts = TYPE_DESC_SPLIT_RE.split(s, maxsplit=1)
    if len(parts) == 2:
        code = parts[0].strip()
        summary = parts[1].strip()
        if looks_like_sign_code(code):
            return code, summary

    return s, ""


def sum_extended(items: Optional[List[Dict[str, Any]]]) -> Optional[float]:
    if not items:
        return None
    total = 0.0
    for item in items:
        ext = safe_num(item.get("extended"))
        if ext is not None:
            total += ext
    return total if total > 0 else None


def shift_cell_ref(cell_ref: str, offset: int) -> str:
    match = re.match(r"^([A-Z]+)(\d+)$", cell_ref)
    if not match:
        return cell_ref
    col, row = match.groups()
    new_row = int(row) + offset
    return f"{col}{new_row}"


def adjust_body_rows_preserve_footer(
    ws,
    sign_count: int,
    body_start: int = 28,
    body_end: int = 47,
    extra_blank_rows: int = 3,
):
    default_body_rows = body_end - body_start + 1
    total_body_rows_needed = sign_count + extra_blank_rows
    row_offset = total_body_rows_needed - default_body_rows

    if row_offset > 0:
        ws.insert_rows(body_end + 1, amount=row_offset)
        logging.info(f"Inserted {row_offset} rows after row {body_end}")
    elif row_offset < 0:
        ws.delete_rows(body_start, amount=abs(row_offset))
        logging.info(f"Deleted {abs(row_offset)} rows starting from row {body_start}")
    else:
        logging.info("No row adjustment needed.")

    return row_offset


def approximate_autofit_rows(
    ws,
    row_start: int,
    row_end: int,
    text_cols: List[str],
    min_height: float = 15.0,
    chars_per_line: int = 50
) -> None:
    for r in range(row_start, row_end + 1):
        max_lines = 1
        for col_letter in text_cols:
            cell = ws[f"{col_letter}{r}"]
            text = str(cell.value) if cell.value else ""
            lines_estimate = max(1, len(text) // chars_per_line)
            if lines_estimate > max_lines:
                max_lines = lines_estimate
        new_height = max(min_height, min_height * max_lines)
        ws.row_dimensions[r].height = new_height


def apply_sheet_protection_for_selection(
    ws,
    body_row_start: int,
    body_row_end: int,
    password: Optional[str] = None
) -> None:
    lock_all_cells(ws, ws.max_row, ws.max_column)
    for r in range(body_row_start, body_row_end + 1):
        for c in ["B", "C", "D", "E"]:
            unlock_cell(ws, f"{c}{r}")

    ws.protection.sheet = True
    ws.protection.password = password
    ws.protection.enable()

    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = True
    ws.protection.formatCells = False
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    ws.protection.insertColumns = False
    ws.protection.insertRows = False
    ws.protection.insertHyperlinks = False
    ws.protection.deleteColumns = False
    ws.protection.deleteRows = False
    ws.protection.sort = False
    ws.protection.autoFilter = False
    ws.protection.pivotTables = False
    ws.protection.objects = False
    ws.protection.scenarios = False

    logging.info("Applied sheet protection - users can ONLY select & modify B-E in body rows.")


# =========================================================
# PDF Upload and Conversion Endpoint (JSON only)
# =========================================================
@app.post("/convert-pdf")
async def convert_pdf(file: UploadFile = File(...)):
    if not file.filename.endswith('.pdf'):
        raise HTTPException(status_code=400, detail="File must be a PDF")
    
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
        content = await file.read()
        tmp_file.write(content)
        tmp_path = tmp_file.name
    
    try:
        extractor = EstimateExtractor(tmp_path)
        extractor.process_pdf()
        json_data = extractor.to_json_format()
        logging.info(f"Extracted {len(json_data['sign_types'])} sign items from PDF")
        return JSONResponse({"success": True, "data": json_data})
    except Exception as e:
        logging.exception("PDF extraction failed")
        raise HTTPException(status_code=500, detail=f"PDF extraction failed: {str(e)}")
    finally:
        try:
            os.unlink(tmp_path)
        except:
            pass


# =========================================================
# NEW: URL-based PDF extraction endpoint (JSON only)
# =========================================================
@app.post("/convert-pdf-url")
async def convert_pdf_url(req: PdfUrlRequest):
    pdf_bytes = await _download_pdf_bytes(req.file_url)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
        tmp_file.write(pdf_bytes)
        tmp_path = tmp_file.name

    try:
        extractor = EstimateExtractor(tmp_path)
        extractor.process_pdf()
        json_data = extractor.to_json_format()
        logging.info(f"Extracted {len(json_data['sign_types'])} sign items from PDF (via URL)")
        return JSONResponse({"success": True, "data": json_data})
    except Exception as e:
        logging.exception("URL PDF extraction failed")
        raise HTTPException(status_code=500, detail=f"URL PDF extraction failed: {str(e)}")
    finally:
        try:
            os.unlink(tmp_path)
        except:
            pass


# =========================================================
# Generate proposal from JSON (original endpoint)
# =========================================================
@app.post("/generate-proposal")
def generate_proposal(payload: Dict[str, Any] = Body(...)):
    cleanup_old_generated_workbooks(OUTPUT_DIR, older_than_minutes=30)

    try:
        estimate_data = json.loads(payload["payload"])
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid JSON string in 'payload': {str(e)}")

    if not isinstance(estimate_data, dict) or not estimate_data:
        raise HTTPException(status_code=400, detail="Decoded 'payload' must be a non-empty JSON object.")

    return _generate_excel_from_data(estimate_data)


# =========================================================
# Combined endpoint - PDF upload -> Excel in one call
# =========================================================
@app.post("/pdf-to-excel")
async def pdf_to_excel(file: UploadFile = File(...)):
    if not file.filename.endswith('.pdf'):
        raise HTTPException(status_code=400, detail="File must be a PDF")
    
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
        content = await file.read()
        tmp_file.write(content)
        tmp_path = tmp_file.name
    
    try:
        extractor = EstimateExtractor(tmp_path)
        extractor.process_pdf()
        estimate_data = extractor.to_json_format()
        logging.info(f"Extracted {len(estimate_data['sign_types'])} sign items from PDF")
        return _generate_excel_from_data(estimate_data)
    except Exception as e:
        logging.exception("PDF to Excel conversion failed")
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")
    finally:
        try:
            os.unlink(tmp_path)
        except:
            pass


# =========================================================
# NEW: Combined endpoint - PDF URL -> Excel in one call
# =========================================================
@app.post("/pdf-to-excel-url")
async def pdf_to_excel_url(req: PdfUrlRequest):
    """
    This endpoint exists for GPT Actions platforms that cannot reliably send multipart file uploads.
    The GPT sends a file_url (JSON). Server downloads the PDF, parses it, generates Excel, returns download_url.
    """
    cleanup_old_generated_workbooks(OUTPUT_DIR, older_than_minutes=30)

    pdf_bytes = await _download_pdf_bytes(req.file_url)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
        tmp_file.write(pdf_bytes)
        tmp_path = tmp_file.name

    try:
        extractor = EstimateExtractor(tmp_path)
        extractor.process_pdf()
        estimate_data = extractor.to_json_format()
        logging.info(f"Extracted {len(estimate_data['sign_types'])} sign items from PDF (via URL)")
        return _generate_excel_from_data(estimate_data)
    except Exception as e:
        logging.exception("URL PDF to Excel conversion failed")
        raise HTTPException(status_code=500, detail=f"URL conversion failed: {str(e)}")
    finally:
        try:
            os.unlink(tmp_path)
        except:
            pass


def _generate_excel_from_data(estimate_data: Dict[str, Any]):
    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(status_code=500, detail=f"Template not found at {TEMPLATE_PATH}")

    try:
        wb = load_workbook(TEMPLATE_PATH)
        if SHEET_NAME not in wb.sheetnames:
            raise HTTPException(status_code=500, detail=f"Sheet '{SHEET_NAME}' not found in workbook.")
        ws = wb[SHEET_NAME]
        
        # DON'T replace the protection object - just disable it temporarily
        ws.protection.sheet = False
        
        logging.info("Disabled template sheet protection (preserving settings)")

        insert_logo(ws)

        FOOTER_HEIGHT_START = 48
        FOOTER_HEIGHT_END = 120
        footer_row_heights = capture_row_heights(ws, FOOTER_HEIGHT_START, FOOTER_HEIGHT_END)

        # Header mapping
        write_cell(ws, "E5", safe_str(estimate_data.get("estimate_date")))
        write_cell(ws, "D8", safe_str(estimate_data.get("project_id")))
        write_cell(ws, "C22", safe_str(estimate_data.get("salesperson")))
        write_cell(ws, "C23", safe_str(estimate_data.get("project_manager")))
        write_cell(ws, "C25", safe_str(estimate_data.get("project_description")))

        sold_to = estimate_data.get("sold_to", {}) or {}
        ship_to = estimate_data.get("ship_to", {}) or {}

        write_cell(ws, "D11", safe_str(sold_to.get("name")))
        write_cell(ws, "D13", join_address_lines(sold_to.get("address_lines") or []))
        sold_csz = " ".join([p for p in [
            safe_str(sold_to.get("city")),
            safe_str(sold_to.get("state")),
            safe_str(sold_to.get("zip"))
        ] if p.strip()])
        write_cell(ws, "D16", sold_csz)
        write_cell(ws, "D17", safe_str(sold_to.get("phone")))

        write_cell(ws, "C11", safe_str(ship_to.get("name")))
        write_cell(ws, "C13", join_address_lines(ship_to.get("address_lines") or []))
        ship_csz = " ".join([p for p in [
            safe_str(ship_to.get("city")),
            safe_str(ship_to.get("state")),
            safe_str(ship_to.get("zip"))
        ] if p.strip()])
        write_cell(ws, "C16", ship_csz)
        write_cell(ws, "C17", safe_str(ship_to.get("phone")))

        sign_types = estimate_data.get("sign_types", []) or []
        sign_count = len(sign_types)

        BODY_START = 28
        BODY_END = 47
        EXTRA_BLANK = 3

        footer_row_offset = adjust_body_rows_preserve_footer(
            ws,
            sign_count=sign_count,
            body_start=BODY_START,
            body_end=BODY_END,
            extra_blank_rows=EXTRA_BLANK
        )

        total_body_rows_needed = sign_count + EXTRA_BLANK
        body_last_row = BODY_START + total_body_rows_needed - 1
        body_change_count = footer_row_offset - 5

        if body_change_count > 0:
            change_text = f"Add {body_change_count}"
        elif body_change_count < 0:
            change_text = f"Subtract {body_change_count}"
        else:
            change_text = "No Change"
        ws["H3"].value = change_text
        logging.info(f"Body row change: {change_text}")

        COL_ITEM, COL_SIGN_TYPE, COL_DESC, COL_QTY, COL_UNIT, COL_TOTAL = "A", "B", "C", "D", "E", "F"
        current_row = BODY_START
        item_num = 1

        sign_code_counts: Dict[str, int] = {}

        for sign in sign_types:
            ws[f"{COL_ITEM}{current_row}"].value = item_num

            raw_type = safe_str(sign.get("sign_type"))
            clean_type, summary = split_sign_type_and_summary(raw_type)

            sign_code_counts.setdefault(clean_type, 0)
            sign_code_counts[clean_type] += 1
            occurrence = sign_code_counts[clean_type]

            desc_summary = summary.strip() if summary else safe_str(sign.get("description")).strip()

            if occurrence == 1:
                ws[f"{COL_SIGN_TYPE}{current_row}"].value = clean_type
                qty_val = safe_num(sign.get("qty"))
                unit_val = round_nearest_dollar(sign.get("unit_price"))
                ws[f"{COL_QTY}{current_row}"].value = qty_val
                ws[f"{COL_DESC}{current_row}"].value = desc_summary
                ws[f"{COL_UNIT}{current_row}"].value = unit_val
                if qty_val is not None and unit_val is not None:
                    ws[f"{COL_TOTAL}{current_row}"].value = f"=D{current_row}*E{current_row}"
                else:
                    ws[f"{COL_TOTAL}{current_row}"].value = None
            else:
                ws[f"{COL_SIGN_TYPE}{current_row}"].value = None
                ws[f"{COL_QTY}{current_row}"].value = None
                ws[f"{COL_DESC}{current_row}"].value = f"ALTERNATE {desc_summary}"
                unit_val = round_nearest_dollar(sign.get("unit_price"))
                ws[f"{COL_UNIT}{current_row}"].value = unit_val
                ws[f"{COL_TOTAL}{current_row}"].value = None

            current_row += 1
            item_num += 1

        totals = estimate_data.get("totals", {}) or {}
        grand_total = safe_num(totals.get("total"))
        shipping_total = sum_extended(estimate_data.get("shipping"))
        install_total = sum_extended(estimate_data.get("installation"))

        SUBTOTAL_CELL = "F48"
        SHIPPING_CELL = "F49"
        INSTALL_CELL = "F53"
        TOTAL_CELL = "F54"

        subtotal_cell_ref = shift_cell_ref(SUBTOTAL_CELL, footer_row_offset)
        body_sum_range = f"F{BODY_START}:F{body_last_row}"
        ws[subtotal_cell_ref].value = f"=SUM({body_sum_range})"
        
        shipping_value = shipping_total if shipping_total else 0.00
        write_cell(ws, shift_cell_ref(SHIPPING_CELL, footer_row_offset), shipping_value)
        
        install_value = install_total if install_total else 0.00
        write_cell(ws, shift_cell_ref(INSTALL_CELL, footer_row_offset), install_value)
        
        if grand_total is not None:
            write_cell(ws, shift_cell_ref(TOTAL_CELL, footer_row_offset), grand_total)

        last_used_row = ws.max_row
        approximate_autofit_rows(
            ws,
            row_start=27,
            row_end=last_used_row,
            text_cols=["C"],
            min_height=15.0
        )

        restore_row_heights(ws, footer_row_heights, footer_row_offset)

        apply_sheet_protection_for_selection(ws, body_row_start=BODY_START, body_row_end=body_last_row)
        
        logging.info("About to save workbook - no more cell modifications after this point")

        file_id = uuid.uuid4().hex
        out_name = f"Boyd_Proposal_{file_id}.xlsx"
        out_path = os.path.join(OUTPUT_DIR, out_name)
        wb.save(out_path)

    except Exception as e:
        logging.exception("Proposal generation failed")
        raise HTTPException(status_code=500, detail=str(e))

    base_url = os.environ.get("RAILWAY_PUBLIC_URL", "").rstrip("/")
    if not base_url:
        base_url = "https://fastapi-production-37f6.up.railway.app"

    download_url = f"{base_url}/download/{out_name}"
    return JSONResponse({"download_url": download_url, "filename": out_name})


@app.get("/download/{filename}")
def download_file(filename: str):
    file_path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found")

    return FileResponse(
        file_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename
    )


@app.get("/")
def root():
    return {
        "service": "Boyd Sign Systems - Estimate Processor",
        "endpoints": {
            "/convert-pdf": "POST - Upload PDF, get JSON data",
            "/convert-pdf-url": "POST - Send {file_url}, get JSON data",
            "/pdf-to-excel": "POST - Upload PDF, get Excel file directly",
            "/pdf-to-excel-url": "POST - Send {file_url}, get Excel file directly",
            "/generate-proposal": "POST - Send JSON data, get Excel file"
        }
    }
