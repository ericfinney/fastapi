import os
import uuid
import json
import logging
import re
from typing import Dict, Any, List, Optional

from fastapi import FastAPI, Body, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

logging.basicConfig(level=logging.INFO)

TEMPLATE_PATH = os.environ.get("BOYD_TEMPLATE_PATH", "templates/Blank.xlsx")
SHEET_NAME = os.environ.get("BOYD_SHEET_NAME", "Proposal")
OUTPUT_DIR = os.environ.get("BOYD_OUTPUT_DIR", "/tmp/output")
LOGO_PATH = os.environ.get("BOYD_LOGO_PATH", "assets/logo.png")

os.makedirs(OUTPUT_DIR, exist_ok=True)

app = FastAPI()


# ---------------------------------------------------------
# Helpers
# ---------------------------------------------------------
def safe_str(x) -> str:
    return "" if x is None else str(x)

def safe_num(x):
    try:
        if x is None or x == "":
            return None
        return float(x)
    except Exception:
        return None

def join_address_lines(addr_lines: List[str]) -> str:
    return "\n".join([line for line in addr_lines if line and line.strip()])

def write_cell(ws, cell: str, value):
    ws[cell].value = value

def summarize_components(components: List[Dict[str, Any]], max_lines: int = 4) -> str:
    if not components:
        return ""
    lines = []
    for c in components[:max_lines]:
        desc = safe_str(c.get("description")).strip()
        dims = safe_str(c.get("dimensions")).strip()
        qty = c.get("qty")

        parts = []
        if desc:
            parts.append(desc)
        if dims:
            parts.append(dims)
        if qty is not None:
            parts.append(f"Qty {qty}")

        if parts:
            lines.append("• " + " | ".join(parts))

    if len(components) > max_lines:
        lines.append("• (additional components omitted)")

    return "\n".join(lines)

def split_sign_type_and_summary(raw_sign_type: str):
    """
    Handles:
      'D - Donor Room'
      'D- Donor Room'
      'D-Donor Room'
    => sign_type='D', summary='Donor Room'
    """
    if not raw_sign_type:
        return "", ""
    s = raw_sign_type.strip()
    m = re.match(r"^([A-Za-z0-9]+)\s*-\s*(.+)$", s)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return s, ""

def build_description_one_cell(sign: Dict[str, Any]) -> str:
    """
    ONE CELL description format:
      Line 1: summary (from sign_type split) if available
      Line 2+: base description (if not duplicate)
      Line 3+: bullet components
    """
    raw_sign_type = safe_str(sign.get("sign_type"))
    _, summary = split_sign_type_and_summary(raw_sign_type)

    base = safe_str(sign.get("description")).strip()
    comps = sign.get("components") or []
    comp_summary = summarize_components(comps).strip()

    lines = []
    if summary:
        lines.append(summary)
        if base and base.lower() != summary.lower():
            lines.append(base)
    else:
        if base:
            lines.append(base)

    if comp_summary:
        lines.append(comp_summary)

    return "\n".join([ln for ln in lines if ln])

def sum_extended(items: Optional[List[Dict[str, Any]]]) -> Optional[float]:
    if not items:
        return None
    total = 0.0
    found = False
    for it in items:
        val = safe_num(it.get("extended_total"))
        if val is not None:
            total += val
            found = True
    return total if found else None

def insert_logo(ws):
    """
    Reinserts logo at A1 every time.
    Requires Pillow installed (pip install pillow).
    """
    if not os.path.exists(LOGO_PATH):
        logging.warning(f"Logo not found at {LOGO_PATH}; skipping insert.")
        return

    img = XLImage(LOGO_PATH)
    ws.add_image(img, "A1")


def adjust_body_rows(ws, sign_count: int, body_start: int = 28, body_end: int = 47, extra_blank_rows: int = 3):
    """
    Ensures the body section has enough rows to fit:
      sign_count + extra_blank_rows
    Body is initially body_start..body_end (inclusive).
    Footer begins at body_end + 1.
    
    If needed rows > base rows, inserts rows right before footer.
    If needed rows < base rows, deletes rows from within the body so footer shifts up.
    """
    base_rows = body_end - body_start + 1
    needed_rows = sign_count + extra_blank_rows

    footer_start = body_end + 1
    diff = needed_rows - base_rows

    if diff > 0:
        # Need more body rows: insert diff rows before footer
        logging.info(f"Inserting {diff} rows at {footer_start} to expand body from {base_rows} to {needed_rows}.")
        ws.insert_rows(footer_start, amount=diff)

    elif diff < 0:
        # Too many body rows: delete rows from the bottom part of the body
        delete_count = abs(diff)
        delete_start = body_start + needed_rows  # first row AFTER the needed body content
        logging.info(f"Deleting {delete_count} rows at {delete_start} to shrink body from {base_rows} to {needed_rows}.")
        ws.delete_rows(delete_start, amount=delete_count)

    # If diff == 0, body is already correct size
    return needed_rows


# ---------------------------------------------------------
# Health check
# ---------------------------------------------------------
@app.get("/")
def root():
    return {
        "status": "ok",
        "template_exists": os.path.exists(TEMPLATE_PATH),
        "template_path": TEMPLATE_PATH,
        "sheet_name": SHEET_NAME,
        "logo_exists": os.path.exists(LOGO_PATH),
        "logo_path": LOGO_PATH
    }


# ---------------------------------------------------------
# Generate proposal (returns download URL)
# ---------------------------------------------------------
@app.post("/generate_proposal")
def generate_proposal(payload: Dict[str, Any] = Body(default=None)):
    logging.info("Incoming request: payload keys = %s", list(payload.keys()) if payload else None)

    if not payload or "payload" not in payload:
        raise HTTPException(status_code=400, detail="Missing required field 'payload' (JSON string).")

    try:
        estimate_data = json.loads(payload["payload"])
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid JSON string in 'payload': {str(e)}")

    if not isinstance(estimate_data, dict) or not estimate_data:
        raise HTTPException(status_code=400, detail="Decoded 'payload' must be a non-empty JSON object.")

    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(status_code=500, detail=f"Template not found at {TEMPLATE_PATH}")

    try:
        wb = load_workbook(TEMPLATE_PATH)
        if SHEET_NAME not in wb.sheetnames:
            raise HTTPException(status_code=500, detail=f"Sheet '{SHEET_NAME}' not found in workbook.")
        ws = wb[SHEET_NAME]

        # ✅ Ensure logo is present in output
        insert_logo(ws)

        # ---------------------------------------------------------
        # Header / mapped fields
        # ---------------------------------------------------------
        write_cell(ws, "E5", safe_str(estimate_data.get("estimate_date")))
        write_cell(ws, "D8", safe_str(estimate_data.get("project_id")))
        write_cell(ws, "C22", safe_str(estimate_data.get("salesperson")))
        write_cell(ws, "C23", safe_str(estimate_data.get("project_manager")))
        write_cell(ws, "C25", safe_str(estimate_data.get("project_description")))

        # ---------------------------------------------------------
        # Sold-to / Ship-to blocks
        # ---------------------------------------------------------
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

        # ---------------------------------------------------------
        # BODY RANGE RULES
        # Body is rows 28..47 in the template (20 rows)
        # Must fit sign_count + 3 blank lines
        # ---------------------------------------------------------
        sign_types = estimate_data.get("sign_types", []) or []
        sign_count = len(sign_types)

        BODY_START = 28
        BODY_END = 47
        EXTRA_BLANK = 3

        needed_body_rows = adjust_body_rows(
            ws,
            sign_count=sign_count,
            body_start=BODY_START,
            body_end=BODY_END,
            extra_blank_rows=EXTRA_BLANK
        )

        # ---------------------------------------------------------
        # Line items table (shift left): A-F
        # ---------------------------------------------------------
        COL_ITEM, COL_SIGN_TYPE, COL_DESC, COL_QTY, COL_UNIT, COL_TOTAL = "A", "B", "C", "D", "E", "F"

        current_row = BODY_START
        item_num = 1

        # Write sign types
        for sign in sign_types:
            ws[f"{COL_ITEM}{current_row}"].value = item_num

            raw_type = safe_str(sign.get("sign_type"))
            clean_type, _ = split_sign_type_and_summary(raw_type)
            ws[f"{COL_SIGN_TYPE}{current_row}"].value = clean_type

            ws[f"{COL_DESC}{current_row}"].value = build_description_one_cell(sign)
            ws[f"{COL_QTY}{current_row}"].value = safe_num(sign.get("qty"))

            unit_price = safe_num(sign.get("unit_price"))
            ws[f"{COL_UNIT}{current_row}"].value = round(unit_price) if unit_price is not None else None

            ws[f"{COL_TOTAL}{current_row}"].value = safe_num(sign.get("extended_total"))

            current_row += 1
            item_num += 1

        # Ensure 3 blank lines after the last sign type
        blank_rows_to_write = EXTRA_BLANK
        for _ in range(blank_rows_to_write):
            # Clear cells in the blank rows, just in case template had placeholders
            for col in [COL_ITEM, COL_SIGN_TYPE, COL_DESC, COL_QTY, COL_UNIT, COL_TOTAL]:
                ws[f"{col}{current_row}"].value = None
            current_row += 1

        # (If the template body is larger than needed, adjust_body_rows already deleted extras)

        # ---------------------------------------------------------
        # Totals section
        # ⚠️ NOTE: Totals cells may SHIFT if footer moved.
        # Because we insert/delete rows ABOVE the footer, totals positions shift with the footer automatically.
        # But the cell references (F48 etc.) are now no longer stable.
        #
        # ✅ Best practice: Anchor totals relative to the "SUBTOTAL" label cell.
        # For now, if your totals are below the body, you should locate them by label.
        # ---------------------------------------------------------

        # --- FIND totals by label (recommended, stable) ---
        def find_cell_with_text(ws, text: str):
            text = text.strip().lower()
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.strip().lower() == text:
                        return cell
            return None

        totals = estimate_data.get("totals", {}) or {}
        subtotal = safe_num(totals.get("sub_total"))
        grand_total = safe_num(totals.get("total"))
        shipping_total = sum_extended(estimate_data.get("shipping"))
        install_total = sum_extended(estimate_data.get("installation"))

        # Example: label "Subtotal:" might be in column E and value in F. Adjust if needed.
        subtotal_label = find_cell_with_text(ws, "subtotal:")
        if subtotal_label and subtotal is not None:
            ws.cell(row=subtotal_label.row, column=subtotal_label.column + 1).value = subtotal

        shipping_label = find_cell_with_text(ws, "shipping:")
        if shipping_label and shipping_total is not None:
            ws.cell(row=shipping_label.row, column=shipping_label.column + 1).value = shipping_total

        installation_label = find_cell_with_text(ws, "installation:")
        if installation_label and install_total is not None:
            ws.cell(row=installation_label.row, column=installation_label.column + 1).value = install_total

        total_label = find_cell_with_text(ws, "total:")
        if total_label and grand_total is not None:
            ws.cell(row=total_label.row, column=total_label.column + 1).value = grand_total

        # ---------------------------------------------------------
        # Save output workbook
        # ---------------------------------------------------------
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
        filename=filename,
    )
