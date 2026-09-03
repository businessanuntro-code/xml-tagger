```python
from pathlib import Path
import shutil

from fastapi import FastAPI, Request, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates


# ============================================================
# XML TAGGER
# ============================================================

app = FastAPI(
    title="XML Tagger",
    version="1.0.0"
)


# ============================================================
# DIRECTOARE
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

TEMPLATES_DIR = BASE_DIR / "templates"
UPLOADS_DIR = BASE_DIR / "uploads"

# Creăm folderul uploads dacă nu există
UPLOADS_DIR.mkdir(parents=True, exist_ok=True)


# ============================================================
# TEMPLATE-URI
# ============================================================

templates = Jinja2Templates(
    directory=str(TEMPLATES_DIR)
)


# ============================================================
# PDF-URI STATICE
# ============================================================

app.mount(
    "/pdf",
    StaticFiles(directory=str(UPLOADS_DIR)),
    name="pdf"
)


# ============================================================
# PAGINA PRINCIPALĂ
# ============================================================

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={
            "request": request
        }
    )


# ============================================================
# UPLOAD PDF
# ============================================================

@app.post("/upload")
async def upload_pdf(file: UploadFile = File(...)):

    # --------------------------------------------------------
    # Verificăm extensia
    # --------------------------------------------------------

    if not file.filename:
        raise HTTPException(
            status_code=400,
            detail="Nu a fost selectat niciun fișier."
        )

    original_name = Path(file.filename).name

    if not original_name.lower().endswith(".pdf"):
        raise HTTPException(
            status_code=400,
            detail="Este permis doar formatul PDF."
        )


    # --------------------------------------------------------
    # Salvăm PDF-ul
    # --------------------------------------------------------

    destination = UPLOADS_DIR / original_name

    try:

        with destination.open("wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=f"Eroare la salvarea PDF-ului: {str(e)}"
        )

    finally:

        await file.close()


    # --------------------------------------------------------
    # Răspuns
    # --------------------------------------------------------

    return {
        "status": "success",
        "filename": original_name,
        "url": f"/pdf/{original_name}"
    }


# ============================================================
# HEALTH CHECK
# ============================================================

@app.get("/health")
async def health():
    return {
        "status": "ok",
        "application": "xml_tagger"
    }
```
