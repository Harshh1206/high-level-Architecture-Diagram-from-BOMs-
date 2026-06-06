"""
main.py — FastAPI backend for DeltaV Architecture Generator React UI
Run: uvicorn main:app --reload --port 8000

Endpoints:
  POST /api/upload   — parse + classify + group a BOM file
  POST /api/generate — generate PPTX from structure + overrides
  GET  /api/health   — health check
"""

import os, sys, tempfile, json
sys.path.insert(0, os.path.dirname(__file__))

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel
from typing import Any

from parser     import parse_bom
from classifier import classify_dataframe
from grouper    import group_bom
from generator  import generate_pptx

app = FastAPI(title="DeltaV Arch Gen API")

# Allow React dev server (localhost:5173 or 3000)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://localhost:3000", "*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# ── /api/health ───────────────────────────────────────────────────────────────
@app.get("/api/health")
def health():
    return {"status": "ok"}


# ── /api/upload ───────────────────────────────────────────────────────────────
@app.post("/api/upload")
async def upload(
    file: UploadFile = File(...),
    project_title: str = Form("DeltaV Architecture"),
):
    # Save to temp file
    ext = os.path.splitext(file.filename)[1].lower()
    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
        tmp.write(await file.read())
        tmp_path = tmp.name

    try:
        # Run Python pipeline
        df         = parse_bom(tmp_path)
        df         = classify_dataframe(df)
        structure  = group_bom(df)

        # Build flat_items list for the React editor
        flat_items = []
        item_id    = 0
        for room, cabinets in structure.items():
            for cabinet, items in cabinets.items():
                for it in items:
                    flat_items.append({
                        "id":            f"i{item_id}",
                        "description":   it.get("description", ""),
                        "diagram_class": it.get("diagram_class", "UNKNOWN"),
                        "qty":           it.get("qty", 1),
                        "room":          room,
                        "cabinet":       cabinet,
                        "confidence":    it.get("confidence", "MEDIUM"),
                        "color_hex":     it.get("color_hex", "888888"),
                        "text_color":    it.get("text_color", "FFFFFF"),
                        "part_number":   it.get("part_number", ""),
                        "label":         it.get("label", it.get("diagram_class", "")),
                    })
                    item_id += 1

        return JSONResponse({
            "structure":   structure,
            "flat_items":  flat_items,
            "row_count":   len(df),
            "project_title": project_title,
        })

    except Exception as e:
        raise HTTPException(status_code=422, detail=str(e))
    finally:
        os.unlink(tmp_path)


# ── /api/generate ─────────────────────────────────────────────────────────────
class GenerateRequest(BaseModel):
    structure:     Any
    project_title: str = "DeltaV System Architecture"


@app.post("/api/generate")
def generate(req: GenerateRequest):
    out = tempfile.mktemp(suffix=".pptx")
    try:
        generate_pptx(req.structure, out, project_title=req.project_title)
        fname = req.project_title.replace(" ", "_").replace("/", "_") + "_Architecture.pptx"
        return FileResponse(
            out,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename=fname,
            background=None,   # keep file alive until response sent
        )
    except Exception as e:
        if os.path.exists(out):
            os.unlink(out)
        raise HTTPException(status_code=500, detail=str(e))


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)