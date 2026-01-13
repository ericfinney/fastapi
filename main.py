import os
import uuid
import json
import logging
import re
import base64
from io import BytesIO
from pathlib import Path

from fastapi import FastAPI, Body, HTTPException
from fastapi.responses import FileResponse, JSONResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from pypdf import PdfReader

# Import the existing main.py functionality
import sys
sys.path.insert(0, '/mnt/user-data/uploads')
from main import generate_proposal_internal, cleanup_old_generated_workbooks, OUTPUT_DIR

logging.basicConfig(level=logging.INFO)

app = FastAPI(title="Boyd Proposal Generator API")

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Models
class GenerateProposalRequest(BaseModel):
    payload: str

class ExtractPdfRequest(BaseModel):
    pdf_base64: str

class GenerateFromPdfRequest(BaseModel):
    pdf_base64: str
    filename: str = "estimate.pdf"


# =========================================================
# Web Interface Route
# =========================================================
@app.get("/", response_class=HTMLResponse)
async def web_interface():
    """Serve the HTML upload interface"""
    html_path = Path(__file__).parent / "templates" / "upload.html"
    
    if not html_path.exists():
        return """
        <html>
        <body>
        <h1>Boyd Proposal Generator</h1>
        <p>Web interface not found. Use the API endpoint directly:</p>
        <pre>POST /generate_proposal</pre>
        </body>
        </html>
        """
    
    with open(html_path) as f:
        return f.read()


# =========================================================
# PDF Text Extraction Endpoint
# =========================================================
@app.post("/extract_pdf_text")
async def extract_pdf_text(request: ExtractPdfRequest):
    """Extract text from a base64-encoded PDF"""
    try:
        # Decode base64 PDF
        pdf_bytes = base64.b64decode(request.pdf_base64)
        
        # Read PDF
        pdf_file = BytesIO(pdf_bytes)
        reader = PdfReader(pdf_file)
        
        # Extract text from all pages
        text = ""
        for page in reader.pages:
            text += page.extract_text() + "\n"
        
        return {"text": text, "num_pages": len(reader.pages)}
        
    except Exception as e:
        logging.error(f"Error extracting PDF text: {e}")
        raise HTTPException(status_code=400, detail=f"Failed to extract PDF text: {str(e)}")


# =========================================================
# Enhanced PDF Parsing Function
# =========================================================
def parse_boyd_estimate_from_text(text: str) -> dict:
    """
    Parse Boyd estimate PDF text into structured JSON data.
    This handles the complex multi-page format with line items.
    
    The key pattern is:
    Line items are indented and follow this format:
    "   <description> <qty> $ <unit_price> $ <extended>"
    Example: "   i01 - Typical Room ID 254 $ 484.374 $ 123,031.00"
    """
    lines = [line for line in text.split('\n')]
    
    data = {
        "estimate_date": "",
        "project_id": "",
        "project_description": "",
        "salesperson": "",
        "project_manager": "",
        "project_engineer": "",
        "sold_to": {
            "name": "",
            "address_lines": [],
            "city": "",
            "state": "",
            "zip": "",
            "phone": ""
        },
        "ship_to": {
            "name": "",
            "address_lines": [],
            "city": "",
            "state": "",
            "zip": ""
        },
        "sign_types": [],
        "shipping": [],
        "installation": [],
        "totals": {
            "sub_total": 0.0,
            "mark_up": 0.0,
            "total": 0.0
        }
    }
    
    # Extract header information
    for line in lines:
        # Date
        if 'Date' in line and not data['estimate_date']:
            match = re.search(r'Date\s+(\d{1,2}/\d{1,2}/\d{2,4})', line)
            if match:
                data['estimate_date'] = match.group(1)
        
        # Project ID
        if 'Project Id' in line:
            match = re.search(r'Project Id\s*:\s*(\S+)', line)
            if match:
                data['project_id'] = match.group(1)
        
        # Project Description
        if 'Project Desc' in line:
            match = re.search(r'Project Desc\.?\s*:\s*(.+?)(?:Ship Via|$)', line)
            if match:
                data['project_description'] = match.group(1).strip()
        
        # Salesperson
        if 'Salesperson' in line and ':' in line:
            match = re.search(r'Salesperson\s*:\s*(.+?)(?:\s|$)', line)
            if match:
                data['salesperson'] = match.group(1).strip()
        
        # Project Manager
        if 'Project Manager:' in line:
            match = re.search(r'Project Manager:\s*(.+?)(?:Project|$)', line)
            if match:
                data['project_manager'] = match.group(1).strip()
        
        # Project Engineer
        if 'Project Engineer:' in line:
            match = re.search(r'Project Engineer:\s*(.+?)(?:REQUESTED|$)', line)
            if match:
                data['project_engineer'] = match.group(1).strip()
    
    # Extract line items - they are INDENTED lines with specific pattern
    # Pattern: "   <description> <qty> $ <unit_price>$ <extended>"
    # or:      "   <description> <qty> $ <unit_price> $ <extended>"
    
    in_signage_package = False
    
    for line in lines:
        # Skip samples package
        if 'SIGNAGE SAMPLES PACKAGE' in line:
            in_signage_package = False
            continue
        
        # Start of actual signage package
        if 'SIGNAGE PACKAGE' in line and 'SAMPLES' not in line:
            in_signage_package = True
            continue
        
        # Stop at totals
        if 'SUB-TOTAL' in line:
            break
        
        if in_signage_package:
            # Line items are indented - look for pattern:
            # "   description qty $ price $ total" or "   description qty $ price$ total"
            match = re.match(r'^\s+(.+?)\s+(\d+)\s*\$\s*([\d,]+\.?\d*)\s*\$?\s*([\d,]+\.?\d*)$', line)
            
            if match:
                description = match.group(1).strip()
                try:
                    qty = int(match.group(2))
                    unit_price = float(match.group(3).replace(',', ''))
                    extended = float(match.group(4).replace(',', ''))
                except ValueError:
                    continue
                
                # Skip installation/shipping/packing lines
                lower_desc = description.lower()
                if ('installation' not in lower_desc and 
                    'packing' not in lower_desc and 
                    'shipping' not in lower_desc and
                    'mobilization' not in lower_desc and
                    'permitting' not in lower_desc):
                    
                    data['sign_types'].append({
                        'description': description,
                        'qty': qty,
                        'unit_price': unit_price,
                        'extended': extended
                    })
    
    # Extract totals
    for i, line in enumerate(lines):
        if 'SUB-TOTAL' in line:
            match = re.search(r'\$\s*([\d,]+\.?\d*)', line)
            if match:
                data['totals']['sub_total'] = float(match.group(1).replace(',', ''))
        
        if 'MARK-UP' in line:
            match = re.search(r'\$\s*([\d,]+\.?\d*)', line)
            if match:
                data['totals']['mark_up'] = float(match.group(1).replace(',', ''))
        
        # TOTAL amount appears on line BEFORE "TOTAL" text
        if line.strip() == 'TOTAL' and i > 0:
            prev_line = lines[i-1]
            match = re.search(r'\$\s*([\d,]+\.?\d*)', prev_line)
            if match:
                data['totals']['total'] = float(match.group(1).replace(',', ''))
    
    return data


# =========================================================
# Combined PDF Upload and Processing Endpoint
# =========================================================
@app.post("/generate_proposal_from_pdf")
async def generate_proposal_from_pdf(request: GenerateFromPdfRequest):
    """
    Combined endpoint: Upload PDF, extract text, parse data, generate Excel.
    This is the main endpoint used by the web interface.
    """
    try:
        logging.info(f"Processing PDF: {request.filename}")
        
        # Step 1: Decode and extract PDF text
        pdf_bytes = base64.b64decode(request.pdf_base64)
        pdf_file = BytesIO(pdf_bytes)
        reader = PdfReader(pdf_file)
        
        text = ""
        for page in reader.pages:
            text += page.extract_text() + "\n"
        
        logging.info(f"Extracted {len(text)} characters from {len(reader.pages)} pages")
        
        # Step 2: Parse estimate data from text
        estimate_data = parse_boyd_estimate_from_text(text)
        
        items_count = len(estimate_data.get('sign_types', []))
        logging.info(f"Parsed {items_count} line items")
        
        if items_count == 0:
            raise HTTPException(
                status_code=400,
                detail="No line items found in PDF. Please check the PDF format."
            )
        
        # Step 3: Generate Excel file
        cleanup_old_generated_workbooks(OUTPUT_DIR)
        
        unique_id = uuid.uuid4().hex
        output_filename = f"Boyd_Proposal_{unique_id}.xlsx"
        output_path = os.path.join(OUTPUT_DIR, output_filename)
        
        generate_proposal_internal(estimate_data, output_path)
        
        # Step 4: Return download URL
        base_url = os.environ.get("RAILWAY_PUBLIC_URL", "").rstrip("/")
        if not base_url:
            base_url = "http://localhost:8000"
        
        download_url = f"{base_url}/download/{output_filename}"
        
        return {
            "status": "success",
            "filename": output_filename,
            "download_url": download_url,
            "items_count": items_count,
            "project_id": estimate_data.get('project_id', 'N/A'),
            "total": estimate_data.get('totals', {}).get('total', 0)
        }
        
    except Exception as e:
        logging.error(f"Error processing PDF: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Error processing PDF: {str(e)}")


# =========================================================
# Generate Proposal Endpoint (existing)
# =========================================================
@app.post("/generate_proposal")
async def generate_proposal(request: GenerateProposalRequest):
    """
    Generate Excel proposal from JSON payload.
    Accepts either parsed data or can parse from PDF text.
    """
    try:
        cleanup_old_generated_workbooks(OUTPUT_DIR)
        
        # Parse the payload
        try:
            estimate_data = json.loads(request.payload)
        except json.JSONDecodeError as e:
            raise HTTPException(status_code=400, detail=f"Invalid JSON in payload: {str(e)}")
        
        # Generate unique filename
        unique_id = uuid.uuid4().hex
        output_filename = f"Boyd_Proposal_{unique_id}.xlsx"
        output_path = os.path.join(OUTPUT_DIR, output_filename)
        
        # Call the existing generate_proposal function from original main.py
        generate_proposal_internal(estimate_data, output_path)
        
        # Get base URL for download
        base_url = os.environ.get("RAILWAY_PUBLIC_URL", "").rstrip("/")
        if not base_url:
            base_url = "http://localhost:8000"
        
        download_url = f"{base_url}/download/{output_filename}"
        
        return {
            "status": "success",
            "filename": output_filename,
            "download_url": download_url
        }
        
    except Exception as e:
        logging.error(f"Error generating proposal: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


# =========================================================
# Download Endpoint
# =========================================================
@app.get("/download/{filename}")
async def download_file(filename: str):
    """Download a generated proposal file"""
    file_path = os.path.join(OUTPUT_DIR, filename)
    
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found")
    
    if not filename.startswith("Boyd_Proposal_"):
        raise HTTPException(status_code=403, detail="Invalid filename")
    
    return FileResponse(
        file_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename
    )


# =========================================================
# Health Check
# =========================================================
@app.get("/health")
async def health_check():
    """Health check endpoint"""
    return {"status": "healthy", "service": "boyd-proposal-generator"}


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
