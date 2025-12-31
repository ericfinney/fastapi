import os
import uuid
import json
import logging
from typing import Dict, Any, List, Optional

from fastapi import FastAPI, Body, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from openpyxl import load_workbook

logging.basicConfig(level=logging.INFO)

TEMPLATE_PATH = os.environ.get("BOYD_TEMPLATE_PATH", "templates/Blank.xlsm")
SHEET_NAME = os.environ.get("BOYD_SHEET_NAME", "Proposal")

# Use /tmp for hosted environments
OUTPUT_DIR = os.environ.get("BOYD_OUTPUT_DIR", "/tmp/output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = FastAPI()


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
        desc = safe_str(c.get("description"))
        dims = safe_str(c.get("dimensions"))
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

def build_sign_description(sign: Dict[str, Any]) -> str:
    base = safe_str(sign.get("description"))
    comps = sign.get("components") or []
    comp_summary = summarize_components(comps)
    if base and comp_summary:
        return base + "\n" + comp_summary
    if base:
        return base
    return comp_summary

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


@app.get("/")
def root():
    return {
        "status": "ok",
        "template_exists": os.path.exists(TEMPLATE_PATH),
        "template_path": TEMPLATE_PATH,
        "sheet_name": SHEET_NAME
    }


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
        wb = load_workbook(TEMPLATE_PATH, keep_vba=True)
        if SHEET_NAME not in wb.sheetnames:
            raise HTTPException(status_code=500, detail=f"Sheet '{SHEET_NAME}' not found in workbook.")
        ws = wb[SHEET_NAME]

        # Header
        write_cell(ws, "E5", safe_str(estimate_data.get("estimate_date")))

        # Sold-to / Ship-to
        sold_to = estimate_data.get("sold_to", {}) or {}
        ship_to = estimate_data.get("ship_to", {}) or {}

        write_cell(ws, "D11", safe_str(sold_to.get("name")))
        write_cell(ws, "D13", join_address_lines(sold_to.get("address_lines") or []))
        sold_csz = " ".join([p for p in [safe_str(sold_to.get("city")), safe_str(sold_to.get("state")), safe_str(sold_to.get("zip"))] if p.strip()])
        write_cell(ws, "D16", sold_csz)
        write_cell(ws, "D17", safe_str(sold_to.get("phone")))

        write_cell(ws, "C11", safe_str(ship_to.get("name")))
        write_cell(ws, "C13", join_address_lines(ship_to.get("address_lines") or []))
        ship_csz = " ".join([p for p in [safe_str(ship_to.get("city")), safe_str(ship_to.get("state")), safe_str(ship_to.get("zip"))] if p.strip()])
        write_cell(ws, "C16", ship_csz)
        write_cell(ws, "C17", safe_str(ship_to.get("phone")))

        # Line items
        start_row = 27
        current_row = start_row
        COL_ITEM, COL_SIGN_TYPE, COL_DESC, COL_QTY, COL_UNIT, COL_TOTAL = "C", "D", "E", "F", "G", "H"

        sign_types = estimate_data.get("sign_types", []) or []
        item_num = 1
        for sign in sign_types:
            ws[f"{COL_ITEM}{current_row}"].value = item_num
            ws[f"{COL_SIGN_TYPE}{current_row}"].value = safe_str(sign.get("sign_type"))
            ws[f"{COL_DESC}{current_row}"].value = build_sign_description(sign)
            ws[f"{COL_QTY}{current_row}"].value = safe_num(sign.get("qty"))
            ws[f"{COL_UNIT}{current_row}"].value = safe_num(sign.get("unit_price"))
            ws[f"{COL_TOTAL}{current_row}"].value = safe_num(sign.get("extended_total"))
            current_row += 1
            item_num += 1

        # Totals
        totals = estimate_data.get("totals", {}) or {}
        subtotal = safe_num(totals.get("sub_total"))
        grand_total = safe_num(totals.get("total"))
        shipping_total = sum_extended(estimate_data.get("shipping"))
        install_total = sum_extended(estimate_data.get("installation"))

        if subtotal is not None:
            write_cell(ws, "H48", subtotal)
        if shipping_total is not None:
            write_cell(ws, "H49", shipping_total)
        if install_total is not None:
            write_cell(ws, "H53", install_total)
        if grand_total is not None:
            write_cell(ws, "H54", grand_total)

        # Save file
        file_id = uuid.uuid4().hex
        out_name = f"Boyd_Proposal_{file_id}.xlsm"
        out_path = os.path.join(OUTPUT_DIR, out_name)
        wb.save(out_path)

    except Exception as e:
        logging.exception("Proposal generation failed")
        raise HTTPException(status_code=500, detail=str(e))

    # Return JSON with download URL
    base_url = os.environ.get("RAILWAY_PUBLIC_URL", "").rstrip("/")
    if not base_url:
        # Fallback (works if your app is accessed via its public domain)
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
        media_type="application/vnd.ms-excel.sheet.macroEnabled.12",
        filename=filename
    )
