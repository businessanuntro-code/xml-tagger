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
HTML_UPLOADS_DIR = UPLOADS_DIR / "html"
PDF_UPLOADS_DIR = UPLOADS_DIR / "pdf"

UPLOADS_DIR.mkdir(
    parents=True,
    exist_ok=True
)

HTML_UPLOADS_DIR.mkdir(
    parents=True,
    exist_ok=True
)

PDF_UPLOADS_DIR.mkdir(
    parents=True,
    exist_ok=True
)


# ============================================================
# TEMPLATE-URI
# ============================================================

templates = Jinja2Templates(
    directory=str(TEMPLATES_DIR)
)


# ============================================================
# FISIERE STATICE
# ============================================================

app.mount(
    "/files/html",
    StaticFiles(
        directory=str(HTML_UPLOADS_DIR)
    ),
    name="html_files"
)

app.mount(
    "/files/pdf",
    StaticFiles(
        directory=str(PDF_UPLOADS_DIR)
    ),
    name="pdf_files"
)


# ============================================================
# PAGINA PRINCIPALA
# ============================================================

@app.get(
    "/",
    response_class=HTMLResponse
)
async def index(request: Request):

    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={
            "request": request
        }
    )


# ============================================================
# UPLOAD HTML5
# ============================================================

@app.post("/upload-html")
async def upload_html(
    file: UploadFile = File(...)
):

    if not file.filename:

        raise HTTPException(
            status_code=400,
            detail="Nu a fost selectat niciun fișier."
        )

    # --------------------------------------------------------
    # Eliminăm eventualele directoare din numele fișierului
    # --------------------------------------------------------

    original_name = Path(
        file.filename
    ).name

    # --------------------------------------------------------
    # Acceptăm HTML / HTM
    # --------------------------------------------------------

    extension = Path(
        original_name
    ).suffix.lower()

    if extension not in (
        ".html",
        ".htm"
    ):

        raise HTTPException(
            status_code=400,
            detail="Este permis doar formatul HTML sau HTM."
        )

    destination = HTML_UPLOADS_DIR / original_name

    try:

        with destination.open("wb") as buffer:

            shutil.copyfileobj(
                file.file,
                buffer
            )

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=f"Eroare la salvarea fișierului HTML: {str(e)}"
        )

    finally:

        await file.close()

    return {
        "status": "success",
        "filename": original_name,
        "url": f"/files/html/{original_name}"
    }


# ============================================================
# UPLOAD PDF
#
# Păstrăm endpoint-ul existent pentru moment.
# Nu îl eliminăm deoarece încă folosim PDF-ul actual.
# ============================================================

@app.post("/upload")
async def upload_pdf(
    file: UploadFile = File(...)
):

    if not file.filename:

        raise HTTPException(
            status_code=400,
            detail="Nu a fost selectat niciun fișier."
        )

    # --------------------------------------------------------
    # Eliminăm eventualele directoare din numele fișierului
    # --------------------------------------------------------

    original_name = Path(
        file.filename
    ).name

    # --------------------------------------------------------
    # Acceptăm numai PDF
    # --------------------------------------------------------

    if not original_name.lower().endswith(".pdf"):

        raise HTTPException(
            status_code=400,
            detail="Este permis doar formatul PDF."
        )

    destination = PDF_UPLOADS_DIR / original_name

    try:

        with destination.open("wb") as buffer:

            shutil.copyfileobj(
                file.file,
                buffer
            )

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=f"Eroare la salvarea PDF-ului: {str(e)}"
        )

    finally:

        await file.close()

    return {
        "status": "success",
        "filename": original_name,
        "url": f"/files/pdf/{original_name}"
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
