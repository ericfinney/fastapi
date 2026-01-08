import os
import uuid
import json
import logging
import re
import math
import time
from copy import copy
from typing import Dict, Any, List, Optional, Tuple

from fastapi import FastAPI, Body, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Protection

logging.basicConfig(level=logging.INFO)

TEMPLATE_PATH = os.environ.get("BOYD_TEMPLATE_PATH", "templates/Blank.xlsx")
SHEET_NAME = os.environ.get("BOYD_SHEET_NAME", "Proposal")
OUTPUT_DIR = os.environ.get("BOYD_OUTPUT_DIR", "/tmp/output")
LOGO_PATH = os.environ.get("BOYD_LOGO_PATH", "assets/logo.png")

os.makedirs(OUTPUT_DIR, exist_ok=True)

app = FastAPI()


# =========================================================
# 🔥 Cleanup old proposal files (30 min)
# =========================================================
def cleanup_old_proposals(
    directory: str,
    max_age_minutes: int = 30
):
    now = time.time()
    max_age_seconds = max_age_minutes * 60

    if not os.path.isdir(directory):
        return

    for filename in os.listdir(directory):
        if not (
            filename.startswith("Boyd_Proposal_")
            and filename.lower().endswith(".xlsx")
        ):
            continue

        path = os.path.join(directory, filename)
        if not os.path.isfile(path):
            continue

        try:
            age_seconds = now - os.path.getmtime(path)
            if age_seconds > max_age_seconds:
                os.remove(path)
                logging.info(f"Deleted old proposal: {filename}")
        except Exception as e:
            logging.warning(f"Failed to delete {filename}: {e}")


# =========================================================
# Helpers
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
    try:
        return math.ceil(float(value))
    except Exception:
        return None

def join_address_lines(addr_lines: List[str]) -> str:
    return "\n".join([line for line in addr_lines if line and line.strip()])

def write_cell(ws, cell: str, value):
    ws[cell].value = value

def insert_logo(ws):
    if os.path.exists(LOGO_PATH):
        ws.add_image(XLImage(LOGO_PATH), "A1")


def lock_cell(ws, cell_ref: str):
    ws[cell_ref].protection = Protection(locked=True)

def unlock_cell(ws, cell_ref: str):
    ws[cell_ref].protection = Protection(locked=False)


# =========================================================
# Sign type parsing
# =========================================================
TYPE_DESC_SPLIT_RE = re.compile(r"\s*[-–—]\s*", flags=re.UNICODE)

def looks_like_sign_code(code: str) -> bool:
    if not code or len(code) > 40:
        return False
    return bool(re.match(r"^[A-Za-z0-9][A-Za-z0-9./&_ ,]+$", code))

def split_sign_type_and_summary(raw: str) -> Tuple[str, str]:
    if not raw:
        return "", ""
    parts = TYPE_DESC_SPLIT_RE.split(raw.strip(), maxsplit=1)
    if len(parts) == 2 and looks_like_sign_code(parts[0]):
        return parts[0].strip(), parts[1].strip()
    return raw.strip(), ""


# =========================================================
# FastAPI
# =========================================================
@app.post("/generate_proposal")
def generate_proposal(payload: Dict[str, Any] = Body(...)):

    # 🔥 CLEAN OLD FILES FIRST
    cleanup_old_proposals(OUTPUT_DIR, max_age_minutes=30)

    try:
        estimate_data = json.loads(payload["payload"])
    except Exception as e:
        raise HTTPException(400, f"Invalid JSON payload: {e}")

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb[SHEET_NAME]

    insert_logo(ws)

    # Header
    write_cell(ws, "E5", estimate_data.get("estimate_date"))
    write_cell(ws, "D8", estimate_data.get("project_id"))
    write_cell(ws, "C22", estimate_data.get("salesperson"))
    write_cell(ws, "C23", estimate_data.get("project_manager"))
    write_cell(ws, "C25", estimate_data.get("project_description"))

    # Body
    BODY_START = 28
    sign_types = estimate_data.get("sign_types", [])

    row = BODY_START
    for idx, sign in enumerate(sign_types, start=1):
        sign_type, desc = split_sign_type_and_summary(sign.get("sign_type"))

        ws[f"A{row}"].value = idx
        ws[f"B{row}"].value = sign_type
        ws[f"C{row}"].value = desc
        ws[f"D{row}"].value = safe_num(sign.get("qty"))
        ws[f"E{row}"].value = round_up_dollars(sign.get("unit_price"))
        ws[f"F{row}"].value = round_up_dollars(sign.get("extended_total"))

        row += 1

    # 🔒 Locking
    for r in range(BODY_START, row):
        unlock_cell(ws, f"B{r}")
        unlock_cell(ws, f"C{r}")
        unlock_cell(ws, f"D{r}")
        unlock_cell(ws, f"E{r}")
        lock_cell(ws, f"F{r}")

    ws.protection.sheet = True
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = True

    # Save
    out_name = f"Boyd_Proposal_{uuid.uuid4().hex}.xlsx"
    out_path = os.path.join(OUTPUT_DIR, out_name)
    wb.save(out_path)

    base_url = os.environ.get("RAILWAY_PUBLIC_URL", "").rstrip("/")
    if not base_url:
        base_url = "https://fastapi-production-37f6.up.railway.app"

    return {
        "download_url": f"{base_url}/download/{out_name}",
        "filename": out_name
    }


@app.get("/download/{filename}")
def download_file(filename: str):
    path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(path):
        raise HTTPException(404, "File not found")
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename
    )
