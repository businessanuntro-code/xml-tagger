from fastapi import FastAPI, File, UploadFile, Request
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from pathlib import Path
from html import escape
import fitz
import shutil
import uuid
import re


# ============================================================
# CONFIGURARE
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

TEMPLATES_DIR = BASE_DIR / "templates"
UPLOADS_DIR = BASE_DIR / "uploads"

PDF_UPLOADS_DIR = UPLOADS_DIR / "pdf"
PDF_PAGES_DIR = UPLOADS_DIR / "pdf_pages"

PDF_UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
PDF_PAGES_DIR.mkdir(parents=True, exist_ok=True)


# ============================================================
# FASTAPI
# ============================================================

app = FastAPI(
    title="XML Tagger",
    version="4.0.0"
)

templates = Jinja2Templates(
    directory=str(TEMPLATES_DIR)
)


# ============================================================
# STATIC
# ============================================================

app.mount(
    "/files/pdf",
    StaticFiles(directory=str(PDF_UPLOADS_DIR)),
    name="pdf_files"
)

app.mount(
    "/files/pages",
    StaticFiles(directory=str(PDF_PAGES_DIR)),
    name="pdf_pages"
)


# ============================================================
# UTILITARE
# ============================================================

def safe_filename(filename: str) -> str:
    """
    Curata numele fisierului.
    """

    if not filename:
        return "document.pdf"

    filename = Path(filename).name

    filename = re.sub(
        r"[^a-zA-Z0-9_\-.ăâîșțĂÂÎȘȚ ]+",
        "_",
        filename
    )

    filename = filename.replace(" ", "_")

    if not filename.lower().endswith(".pdf"):
        filename += ".pdf"

    return filename


# ============================================================
# RANDARE PDF
# ============================================================

def render_pdf_page(
    page,
    output_path: Path,
    dpi: int = 144
):
    """
    Randeaza pagina PDF ca PNG.

    PDF-ul ramane sursa vizuala.
    Nu reconstruim documentul in HTML.
    """

    zoom = dpi / 72.0

    matrix = fitz.Matrix(
        zoom,
        zoom
    )

    pix = page.get_pixmap(
        matrix=matrix,
        alpha=False
    )

    pix.save(
        str(output_path)
    )

    return {
        "width": pix.width,
        "height": pix.height
    }


# ============================================================
# EXTRAGERE TEXT CU COORDONATE
# ============================================================

def extract_words(page):
    """
    Extrage fiecare cuvant impreuna cu pozitia sa exacta
    pe pagina PDF.

    Format fitz:
    x0, y0, x1, y1, text, block_no, line_no, word_no
    """

    words = page.get_text(
        "words",
        sort=True
    )

    result = []

    for index, item in enumerate(words):

        if len(item) < 8:
            continue

        x0 = float(item[0])
        y0 = float(item[1])
        x1 = float(item[2])
        y1 = float(item[3])

        text = str(item[4])

        if not text.strip():
            continue

        result.append({
            "id": index,
            "text": text,

            "x0": x0,
            "y0": y0,
            "x1": x1,
            "y1": y1,

            "block": int(item[5]),
            "line": int(item[6]),
            "word": int(item[7])
        })

    return result


# ============================================================
# EXTRAGERE TEXT COMPLET
# ============================================================

def extract_page_text(page):
    """
    Textul complet al paginii.
    """

    text = page.get_text(
        "text",
        sort=True
    )

    return text


# ============================================================
# PROCESARE PDF
# ============================================================

def process_pdf(pdf_path: Path, file_id: str):
    """
    Proceseaza PDF-ul:

        PDF original
             |
             +---- PNG fiecare pagina
             |
             +---- text
             |
             +---- coordonate cuvinte
    """

    document = fitz.open(
        str(pdf_path)
    )

    pages = []

    try:

        for page_number in range(document.page_count):

            page = document.load_page(
                page_number
            )

            page_id = (
                f"{file_id}_page_"
                f"{page_number + 1}"
            )

            image_filename = (
                f"{page_id}.png"
            )

            image_path = (
                PDF_PAGES_DIR /
                image_filename
            )

            dimensions = render_pdf_page(
                page,
                image_path,
                dpi=144
            )

            rect = page.rect

            words = extract_words(
                page
            )

            page_text = extract_page_text(
                page
            )

            pages.append({
                "page": page_number + 1,

                "width": dimensions["width"],
                "height": dimensions["height"],

                "pdf_width": float(rect.width),
                "pdf_height": float(rect.height),

                "image": (
                    f"/files/pages/"
                    f"{image_filename}"
                ),

                "text": page_text,

                "words": words
            })

    finally:
        document.close()

    return pages


# ============================================================
# HOME
# ============================================================

@app.get(
    "/",
    response_class=HTMLResponse
)
async def home(request: Request):

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

@app.post("/upload-pdf")
async def upload_pdf(
    file: UploadFile = File(...)
):

    if not file.filename:
        return JSONResponse(
            status_code=400,
            content={
                "status": "error",
                "message": "Nu a fost selectat niciun fisier."
            }
        )

    original_filename = file.filename

    if not original_filename.lower().endswith(".pdf"):
        return JSONResponse(
            status_code=400,
            content={
                "status": "error",
                "message": "Fisierul trebuie sa fie PDF."
            }
        )

    clean_filename = safe_filename(
        original_filename
    )

    file_id = uuid.uuid4().hex

    stored_filename = (
        f"{file_id}_{clean_filename}"
    )

    pdf_path = (
        PDF_UPLOADS_DIR /
        stored_filename
    )

    try:

        with pdf_path.open("wb") as buffer:
            shutil.copyfileobj(
                file.file,
                buffer
            )

        pages = process_pdf(
            pdf_path,
            file_id
        )

        return JSONResponse(
            content={
                "status": "success",

                "filename": original_filename,

                "stored_filename": stored_filename,

                "url": (
                    f"/files/pdf/"
                    f"{stored_filename}"
                ),

                "page_count": len(pages),

                "pages": pages
            }
        )

    except Exception as e:

        try:
            if pdf_path.exists():
                pdf_path.unlink()
        except Exception:
            pass

        return JSONResponse(
            status_code=500,
            content={
                "status": "error",
                "message": (
                    "Eroare la procesarea PDF-ului: "
                    f"{str(e)}"
                )
            }
        )


# ============================================================
# HEALTH
# ============================================================

@app.get("/health")
async def health():

    return {
        "status": "ok",
        "app": "XML Tagger",
        "version": "4.0.0"
    }
