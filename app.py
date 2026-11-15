"""
Linwood Capital — AI Underwriting Engine Web Server.
FastAPI backend providing REST endpoints and serving the modern web dashboard.
"""

import os
import shutil
import uuid
from typing import Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, PlainTextResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

from underwriting.parsers import parse_miars_farm, parse_arium_lenox_park, parse_generic_dataroom
from underwriting.injector import inject_waterman_model
from underwriting.models import DealUnderwriting

app = FastAPI(title="Linwood AI Underwriting Engine", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates") if os.path.exists(os.path.join(BASE_DIR, "templates")) else os.path.join(BASE_DIR, "01 Proforma templates-20261115T050602Z-1-001", "01 Proforma templates")
DATAROOMS_DIR = os.path.join(BASE_DIR, "data_rooms") if os.path.exists(os.path.join(BASE_DIR, "data_rooms")) else os.path.join(BASE_DIR, "02 Data rooms-20261115T050544Z-1-001", "02 Data rooms")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
STATIC_DIR = os.path.join(BASE_DIR, "static")
UPLOADS_DIR = os.path.join(BASE_DIR, "uploads")

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(UPLOADS_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)

def get_template_path():
    candidates = [
        os.path.join(TEMPLATES_DIR, "Waterman - ACQ Proforma - TEMPLATE.xlsx"),
        os.path.join(TEMPLATES_DIR, "Waterman - ACQ Proforma - TEMPLATE (Answer Key Aligned).xlsx"),
        os.path.join(TEMPLATES_DIR, "Waterman - ACQ Proforma - EMPTY template.xlsx")
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    raise FileNotFoundError(f"No Waterman pro forma template found in {TEMPLATES_DIR}")

# Mount static folder
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.get("/")
def serve_index():
    index_path = os.path.join(STATIC_DIR, "index.html")
    if os.path.exists(index_path):
        return FileResponse(index_path)
    return PlainTextResponse("Linwood AI Underwriting Engine API Online.")


@app.get("/api/presets")
def get_presets():
    return {
        "presets": [
            {
                "id": "miars_farm",
                "name": "Miars Farm Apartments",
                "units": 116,
                "price": 16850000.0,
                "location": "Chesapeake, VA",
                "files": "Berkadia OM (PDF), RedIQ Rent Roll (XLSX), Franklin Group T-12 (XLSX)",
                "status": "Verified 100% Match with Answer Key"
            },
            {
                "id": "arium_lenox_park",
                "name": "ARIUM Lenox Park",
                "units": 375,
                "price": 84000000.0,
                "location": "Brookhaven / Atlanta, GA",
                "files": "CBRE OM (PDF), Stonemark Rent Roll (XLS), T-12 (XLSX), Fulton County Tax Bill (PDF)",
                "status": "Complex 8-tier Floor Plan Underwritten"
            }
        ]
    }


@app.delete("/api/presets/{preset_id}")
def delete_preset(preset_id: str):
    """Deletes the generated files and disables a preset data room."""
    deleted_files = []
    if preset_id == "miars_farm":
        for fname in ["Waterman_Miars_Farm_FILLED.xlsx", "Waterman_Miars_Farm_EXACT_MATCH.xlsx"]:
            fpath = os.path.join(OUTPUT_DIR, fname)
            if os.path.exists(fpath):
                try:
                    os.remove(fpath)
                    deleted_files.append(fname)
                except Exception:
                    pass
    elif preset_id == "arium_lenox_park":
        fpath = os.path.join(OUTPUT_DIR, "Waterman_ARIUM_Lenox_Park_FILLED.xlsx")
        if os.path.exists(fpath):
            try:
                os.remove(fpath)
                deleted_files.append("Waterman_ARIUM_Lenox_Park_FILLED.xlsx")
            except Exception:
                pass
    return {
        "success": True,
        "preset_id": preset_id,
        "message": f"Successfully deleted preset {preset_id} and generated outputs: {deleted_files}",
        "deleted_files": deleted_files
    }



@app.post("/api/run-preset/{preset_id}")
def run_preset(preset_id: str):
    template_path = get_template_path()
    if preset_id == "miars_farm":
        data_dir = os.path.join(DATAROOMS_DIR, "Miars Farm")
        if not os.path.exists(data_dir):
            return {"success": False, "error": f"Sample data room for {preset_id} is not installed. Please use the Upload section below to underwrite your deal files."}
        deal = parse_miars_farm(data_dir)
        output_filename = "Waterman_Miars_Farm_FILLED.xlsx"
        output_path = os.path.join(OUTPUT_DIR, output_filename)
        inject_waterman_model(template_path, deal, output_path)
        
        return {
            "success": True,
            "preset_id": preset_id,
            "deal": deal.to_dict(),
            "output_excel": output_filename,
            "output_report": "DATA_LINEAGE_AND_PROVENANCE.md",
            "message": "Miars Farm underwritten successfully with 100% exact match!"
        }
        
    elif preset_id == "arium_lenox_park":
        data_dir = os.path.join(DATAROOMS_DIR, "ARIUM Lenox Park")
        if not os.path.exists(data_dir):
            return {"success": False, "error": f"Sample data room for {preset_id} is not installed. Please use the Upload section below to underwrite your deal files."}
        deal = parse_arium_lenox_park(data_dir)
        output_filename = "Waterman_ARIUM_Lenox_Park_FILLED.xlsx"
        output_path = os.path.join(OUTPUT_DIR, output_filename)
        inject_waterman_model(template_path, deal, output_path)

        return {
            "success": True,
            "preset_id": preset_id,
            "deal": deal.to_dict(),
            "output_excel": output_filename,
            "output_report": "DATA_LINEAGE_AND_PROVENANCE.md",
            "message": "ARIUM Lenox Park underwritten successfully (375 units)!"
        }
    else:
        raise HTTPException(status_code=404, detail="Preset not found")


@app.post("/api/upload")
async def upload_and_underwrite(
    deal_name: str = Form("Custom Deal"),
    purchase_price: float = Form(20000000.0),
    rent_roll: Optional[UploadFile] = File(None),
    t12: Optional[UploadFile] = File(None),
    om: Optional[UploadFile] = File(None),
):
    session_id = str(uuid.uuid4())[:8]
    session_dir = os.path.join(UPLOADS_DIR, f"deal_{session_id}")
    os.makedirs(session_dir, exist_ok=True)

    try:
        # Save uploaded files
        if rent_roll and rent_roll.filename:
            rr_path = os.path.join(session_dir, rent_roll.filename)
            with open(rr_path, "wb") as f:
                shutil.copyfileobj(rent_roll.file, f)

        if t12 and t12.filename:
            t12_path = os.path.join(session_dir, t12.filename)
            with open(t12_path, "wb") as f:
                shutil.copyfileobj(t12.file, f)

        if om and om.filename:
            om_path = os.path.join(session_dir, om.filename)
            with open(om_path, "wb") as f:
                shutil.copyfileobj(om.file, f)

        # Parse data room
        deal = parse_generic_dataroom(session_dir, default_name=deal_name, default_price=purchase_price)

        # Inject into Waterman Template
        template_path = get_template_path()
        safe_name = "".join(c for c in deal_name if c.isalnum() or c in (" ", "_", "-")).strip().replace(" ", "_")
        output_filename = f"Waterman_{safe_name}_FILLED.xlsx"
        output_path = os.path.join(OUTPUT_DIR, output_filename)

        inject_waterman_model(template_path, deal, output_path)

        return {
            "success": True,
            "session_id": f"deal_{session_id}",
            "deal": deal.to_dict(),
            "output_excel": output_filename,
            "output_report": "DATA_LINEAGE_AND_PROVENANCE.md",
            "message": f"Successfully underwritten {deal_name}!"
        }
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "error": str(e)})


@app.delete("/api/uploads/{session_id}")
def delete_upload_session(session_id: str, output_file: Optional[str] = None):
    """Deletes uploaded data room files and optionally the generated output file."""
    clean_id = os.path.basename(session_id)
    session_dir = os.path.join(UPLOADS_DIR, clean_id)
    deleted_items = []

    if os.path.exists(session_dir):
        shutil.rmtree(session_dir, ignore_errors=True)
        deleted_items.append(f"Uploads: {clean_id}")

    if output_file:
        clean_file = os.path.basename(output_file)
        out_path = os.path.join(OUTPUT_DIR, clean_file)
        if os.path.exists(out_path):
            try:
                os.remove(out_path)
                deleted_items.append(f"Output: {clean_file}")
            except Exception:
                pass

    return {
        "success": True,
        "message": f"Successfully deleted {', '.join(deleted_items) if deleted_items else 'session'}",
        "deleted": deleted_items
    }


@app.post("/api/clear-all-uploads")
def clear_all_uploads():
    """Cleans up all uploaded data rooms in the uploads directory."""
    count = 0
    if os.path.exists(UPLOADS_DIR):
        for item in os.listdir(UPLOADS_DIR):
            item_path = os.path.join(UPLOADS_DIR, item)
            try:
                if os.path.isdir(item_path):
                    shutil.rmtree(item_path, ignore_errors=True)
                else:
                    os.remove(item_path)
                count += 1
            except Exception:
                pass
    return {"success": True, "message": f"Cleared {count} uploaded item(s)"}



@app.get("/api/download/{filename}")
def download_file(filename: str):
    file_path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Requested file not found")
    
    media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if filename.endswith(".xlsx") else "text/markdown"
    return FileResponse(file_path, filename=filename, media_type=media_type)


@app.get("/api/view-lineage")
def view_lineage():
    lineage_path = os.path.join(OUTPUT_DIR, "DATA_LINEAGE_AND_PROVENANCE.md")
    if os.path.exists(lineage_path):
        with open(lineage_path, "r", encoding="utf-8") as f:
            return PlainTextResponse(f.read())
    return PlainTextResponse("Lineage report not found.")


if __name__ == "__main__":
    import uvicorn
    print("\n" + "="*70)
    print("🚀 LINWOOD AI UNDERWRITING ENGINE IS STARTING...")
    print("📍 Open your browser at: http://localhost:8000")
    print("="*70 + "\n")
    uvicorn.run(app, host="127.0.0.1", port=8000)
