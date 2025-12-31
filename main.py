from fastapi import FastAPI, Body
from fastapi.responses import FileResponse
import uuid
import os
from typing import Dict, Any, List

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

app = FastAPI()

TEMPLATE_PATH = os.environ.get("BOYD_TEMPLATE_PATH", "templates/Blank.xlsm")
SHEET_NAME = os.environ.get("BOYD_SHEET_NAME", "Proposal")

OUTPUT_DIR = os.environ.get("BOYD_OUTPUT_DIR", "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)


# -------------------------
# Helper functions
# -------------------------

def safe_str(x):
    return "" if x is None else str(x)

def safe_num(x):
    try:
        return float(x)
    except:
        return None

def join_address_lines(addr_lines: List[str]) -> str:
    # keeps multi-line address formatting in a single cell if needed
    return "\n".join([line for line in addr_lines if line and line.strip()])

def write_merged_cell(ws, cell, value):
    # openpyxl writing to the top-left of a merged range works
    ws[cell].value = value

def summarize_components(components: List[Dict[str, Any]], max_lines=4) -> str:
    """
    Condense component detail into short bullet-like lines.
    Example: "• Backer Panel (1/4in clear acrylic) - Qty 2"
    """
    lines = []
    for c in components:
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
        if len(lines) >= max_lines:
            break
    if len(components) > max_lines:
        lines.append("• (additional components omitted)")
    return "\n".join(lines)

def build_sign_description(sign: Dict[str, Any]) -> str:
    base = safe_str(sign.get("description"))
    comps = sign.get("components") or []
    comp_summary = summarize_components(comps)
    if comp_summary:
        if base:
            return base + "\n" + comp_summary
        return comp_summary
    return base


# -------------------------
# Core endpoint
# -------------------------

@app.post("/generate_proposal")
def generate_proposal(payload: Dict[str, Any] = Body(...)):
    estimate_data = payload["body"] 

    wb = load_workbook(TEMPLATE_PATH, keep_vba=True)
    ws = wb[SHEET_NAME]

    # -----------------------------------------
    # HEADER FIELDS (Adjust cell references here)
    # -----------------------------------------
    # These are common placements based on your template structure:
    # Date in E5 (merged)
    write_merged_cell(ws, "E5", safe_str(estimate_data.get("estimate_date")))

    # You likely have job/project id somewhere near the top;
    # Adjust these if your template has specific cells.
    # Example:
    # write_merged_cell(ws, "E7", safe_str(estimate_data.get("project_id")))
    # write_merged_cell(ws, "E8", safe_str(estimate_data.get("project_description")))

    # -----------------------------------------
    # SOLD TO / SHIP TO (Adjust cell references)
    # -----------------------------------------
    sold_to = estimate_data.get("sold_to", {})
    ship_to = estimate_data.get("ship_to", {})

    # SOLD TO block — typical placement right of Ship-to (column D)
    write_merged_cell(ws, "D11", safe_str(sold_to.get("name")))
    write_merged_cell(ws, "D13", join_address_lines(sold_to.get("address_lines") or []))

    city = safe_str(sold_to.get("city"))
    state = safe_str(sold_to.get("state"))
    zipc = safe_str(sold_to.get("zip"))
    csz = " ".join([p for p in [city, state, zipc] if p])
    write_merged_cell(ws, "D16", csz)
    write_merged_cell(ws, "D17", safe_str(sold_to.get("phone")))

    # SHIP TO block — typical placement left (column C)
    write_merged_cell(ws, "C11", safe_str(ship_to.get("name")))
    write_merged_cell(ws, "C13", join_address_lines(ship_to.get("address_lines") or []))

    city = safe_str(ship_to.get("city"))
    state = safe_str(ship_to.get("state"))
    zipc = safe_str(ship_to.get("zip"))
    csz = " ".join([p for p in [city, state, zipc] if p])
    write_merged_cell(ws, "C16", csz)
    write_merged_cell(ws, "C17", safe_str(ship_to.get("phone")))

    # -----------------------------------------
    # LINE ITEMS TABLE
    # -----------------------------------------
    # Table header is row 26, items start at row 27
    start_row = 27
    current_row = start_row

    sign_types = estimate_data.get("sign_types", []) or []

    # These columns match your table:
    # B: Item (or A), C: Sign Type, D: Description, E: Qty, F: Unit Price, G: Total
    # Your template shows: C=Item, D=Sign Type, E=Description, F=Qty, G=Unit Price, H=Total
    #
    # We'll set defaults that you can adjust by changing these column letters:
    COL_ITEM = "C"
    COL_SIGN_TYPE = "D"
    COL_DESC = "E"
    COL_QTY = "F"
    COL_UNIT = "G"
    COL_TOTAL = "H"

    item_num = 1
    for sign in sign_types:
        ws[f"{COL_ITEM}{current_row}"].value = item_num
        ws[f"{COL_SIGN_TYPE}{current_row}"].value = safe_str(sign.get("sign_type"))
        ws[f"{COL_DESC}{current_row}"].value = build_sign_description(sign)

        qty = sign.get("qty")
        unit_price = sign.get("unit_price")
        total = sign.get("extended_total")

        ws[f"{COL_QTY}{current_row}"].value = safe_num(qty)
        ws[f"{COL_UNIT}{current_row}"].value = safe_num(unit_price)

        # If the template has formulas for totals, you can omit writing totals.
        # If not, write it:
        ws[f"{COL_TOTAL}{current_row}"].value = safe_num(total)

        item_num += 1
        current_row += 1

    # -----------------------------------------
    # SHIPPING + INSTALLATION TOTALS
    # -----------------------------------------
    def sum_extended(items):
        s = 0.0
        any_found = False
        for it in items or []:
            val = safe_num(it.get("extended_total"))
            if val is not None:
                s += val
                any_found = True
        return s if any_found else None

    shipping_total = sum_extended(estimate_data.get("shipping"))
    install_total = sum_extended(estimate_data.get("installation"))

    totals = estimate_data.get("totals", {}) or {}
    subtotal = safe_num(totals.get("sub_total"))
    grand_total = safe_num(totals.get("total"))

    # These cell placements should be adjusted to match your sheet:
    # Based on your template’s totals section near bottom:
    # Subtotal row around 48, shipping row around 49, installation around 53, grand total around 54.
    #
    # Adjust these cells if your template differs:
    if subtotal is not None:
        ws["H48"].value = subtotal
    if shipping_total is not None:
        ws["H49"].value = shipping_total
    if install_total is not None:
        ws["H53"].value = install_total
    if grand_total is not None:
        ws["H54"].value = grand_total

    # -----------------------------------------
    # Save output
    # -----------------------------------------
    out_name = f"Boyd_Proposal_{uuid.uuid4().hex}.xlsm"
    out_path = os.path.join(OUTPUT_DIR, out_name)
    wb.save(out_path)

    return FileResponse(
        path=out_path,
        media_type="application/vnd.ms-excel.sheet.macroEnabled.12",
        filename=out_name
    )
