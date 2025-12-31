import os
import uuid
import json
import logging
import re
from typing import Dict, Any, List, Optional

from fastapi import FastAPI, Body, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from openpyxl import load_workbook

logging.basicConfig(level=logging.INFO)

# ✅ Switch to XLSX template (no macros)
TEMPLATE_PATH = os.environ.get("BOYD_TEMPLATE_PATH", "templates/Blank.xlsx")
SHEET_NAME = os.environ.get("BOYD_SHEET_NAME", "Proposal")

# Use /tmp for hosted environments
OUTPUT_DIR = os.environ.get("BOYD_OUTPUT_DIR", "/tmp/output")
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
    """
    Condenses component details into short bullet lines.
    """
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

    Also supports:
      'A1 - Something'
      'D2- Something'
    => sign_type='A1', summary='Something'
    """
    if not raw_sign_type:
        return "", ""

    s = raw_sign_type.strip()

    # ✅ Leading code + dash + summary (spaces optional around dash)
    m = re.match(r"^([A-Za-z0-9]+)\s*-\s*(.+)$", s)
    if m:
        code = m.group(1).strip()
        summary = m.group(2).strip()
        return code, summary

    return s, ""

def build_sign_description(sign: Dict[str, Any]) -> str:
    """
    Description cell rules:
    - Line 1: summary extracted from sign_type if formatted "CODE - SUMMARY"
    - Line 2+: base description (if not duplicate)
    - Then: component bullet summary
    """
    raw_sign_type = safe_str(sign.get("sign_type"))
    _, summary_from_type = split_sign_type_and_summary(raw_sign_type)

    base = safe_str(sign.get("description")).strip()
    comps = sign.get("components") or []
    comp_summary = summarize_components(comps).strip()

    lines = []

    # ✅ First line is always summary if available
    if summary_from_type:
        lines.append(summary_from_type)
        if base and base.lower() != summary_from_type.lower():
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


# ---------------------------------------------------------
# Health check
# ---------------------------------------------------------
@app.get("/")
def root():
    return {
        "status": "ok",
        "template_exists": os.path.exists(TEMPLATE_PATH),
        "template_path": TEMPLATE_PATH,
        "sheet_name": SHEET_NAME
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
        # Line items table
        # Start row = 28
        # Shift LEFT by 2 columns: Item=A, SignType=B, Desc=C, Qty=D, Unit=E, Total=F
        # ---------------------------------------------------------
        start_row = 28
        current_row = start_row

        COL_ITEM, COL_SIGN_TYPE, COL_DESC, COL_QTY, COL_UNIT, COL_TOTAL = "A", "B", "C", "D", "E", "F"

        sign_types = estimate_data.get("sign_types", []) or []
        item_num = 1

        for sign in sign_types:
            ws[f"{COL_ITEM}{current_row}"].value = item_num

            raw_type = safe_str(sign.get("sign_type"))
            clean_type, _ = split_sign_type_and_summary(raw_type)
            ws[f"{COL_SIGN_TYPE}{current_row}"].value = clean_type

            ws[f"{COL_DESC}{current_row}"].value = build_sign_description(sign)

            ws[f"{COL_QTY}{current_row}"].value = safe_num(sign.get("qty"))

            # ✅ Round unit price to nearest dollar
            unit_price = safe_num(sign.get("unit_price"))
            ws[f"{COL_UNIT}{current_row}"].value = round(unit_price) if unit_price is not None else None

            ws[f"{COL_TOTAL}{current_row}"].value = safe_num(sign.get("extended_total"))

            current_row += 1
            item_num += 1

        # ---------------------------------------------------------
        # Totals section (shift left by 2: H -> F)
        # ---------------------------------------------------------
        totals = estimate_data.get("totals", {}) or {}
        subtotal = safe_num(totals.get("sub_total"))
        grand_total = safe_num(totals.get("total"))

        shipping_total = sum_extended(estimate_data.get("shipping"))
        install_total = sum_extended(estimate_data.get("installation"))

        if subtotal is not None:
            write_cell(ws, "F48", subtotal)
        if shipping_total is not None:
            write_cell(ws, "F49", shipping_total)
        if install_total is not None:
            write_cell(ws, "F53", install_total)
        if grand_total is not None:
            write_cell(ws, "F54", grand_total)

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
