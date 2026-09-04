from pathlib import Path
from io import BytesIO
import base64
import html
import shutil

from fastapi import (
    FastAPI,
    Request,
    UploadFile,
    File,
    HTTPException,
)

from fastapi.responses import HTMLResponse

from fastapi.staticfiles import StaticFiles

from fastapi.templating import Jinja2Templates

from docx import Document


# ============================================================
# APPLICATION
# ============================================================

app = FastAPI(
    title="XML Tagger",
    version="2.0.0"
)


# ============================================================
# DIRECTORIES
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

TEMPLATES_DIR = BASE_DIR / "templates"

UPLOADS_DIR = BASE_DIR / "uploads"

WORD_UPLOADS_DIR = UPLOADS_DIR / "word"


UPLOADS_DIR.mkdir(
    parents=True,
    exist_ok=True
)

WORD_UPLOADS_DIR.mkdir(
    parents=True,
    exist_ok=True
)


# ============================================================
# TEMPLATES
# ============================================================

templates = Jinja2Templates(
    directory=str(TEMPLATES_DIR)
)


# ============================================================
# STATIC WORD FILES
# ============================================================

app.mount(
    "/files/word",
    StaticFiles(
        directory=str(WORD_UPLOADS_DIR)
    ),
    name="word_files"
)


# ============================================================
# HOME
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
# SAFE FILENAME
# ============================================================

def safe_filename(filename: str) -> str:

    return Path(filename).name


# ============================================================
# ESCAPE HTML
# ============================================================

def escape_html(text: str) -> str:

    return html.escape(
        text or "",
        quote=False
    )


# ============================================================
# IMAGE -> DATA URL
# ============================================================

def image_part_to_data_url(part):

    try:

        content_type = part.content_type

        blob = part.blob

        encoded = base64.b64encode(
            blob
        ).decode("ascii")

        return (
            f"data:{content_type};base64,{encoded}"
        )

    except Exception:

        return ""


# ============================================================
# RUN HTML
# ============================================================

def run_to_html(run):

    text = run.text or ""

    if not text:
        return ""


    value = escape_html(text)


    if run.bold:
        value = f"<strong>{value}</strong>"


    if run.italic:
        value = f"<em>{value}</em>"


    if run.underline:
        value = f"<u>{value}</u>"


    if run.font.superscript:
        value = f"<sup>{value}</sup>"


    if run.font.subscript:
        value = f"<sub>{value}</sub>"


    return value


# ============================================================
# PARAGRAPH STYLE -> HTML
#
# IMPORTANT:
# Aceste stiluri sunt folosite DOAR pentru afișarea
# documentului Word.
#
# Nu sunt folosite pentru generarea XML.
# ============================================================

def paragraph_to_html(paragraph):

    text = paragraph.text or ""

    style_name = ""

    try:
        style_name = (
            paragraph.style.name or ""
        )
    except Exception:
        style_name = ""


    content = ""


    # --------------------------------------------------------
    # RUNS
    # --------------------------------------------------------

    for run in paragraph.runs:

        content += run_to_html(run)


    # --------------------------------------------------------
    # Dacă nu există runs utile
    # --------------------------------------------------------

    if not content and text:

        content = escape_html(text)


    # --------------------------------------------------------
    # Paragraf gol
    # --------------------------------------------------------

    if not content:

        return "<p><br></p>"


    # --------------------------------------------------------
    # Word headings
    #
    # Doar pentru afișare.
    # --------------------------------------------------------

    style_lower = style_name.lower()


    if style_lower == "title":

        return (
            f'<h1 data-word-style="Title">'
            f'{content}'
            f'</h1>'
        )


    if style_lower == "subtitle":

        return (
            f'<h2 data-word-style="Subtitle">'
            f'{content}'
            f'</h2>'
        )


    if style_lower == "heading 1":

        return (
            f'<h2 data-word-style="Heading 1">'
            f'{content}'
            f'</h2>'
        )


    if style_lower == "heading 2":

        return (
            f'<h3 data-word-style="Heading 2">'
            f'{content}'
            f'</h3>'
        )


    if style_lower == "heading 3":

        return (
            f'<h3 data-word-style="Heading 3">'
            f'{content}'
            f'</h3>'
        )


    if style_lower == "heading 4":

        return (
            f'<h4 data-word-style="Heading 4">'
            f'{content}'
            f'</h4>'
        )


    if style_lower == "heading 5":

        return (
            f'<h5 data-word-style="Heading 5">'
            f'{content}'
            f'</h5>'
        )


    if style_lower == "heading 6":

        return (
            f'<h6 data-word-style="Heading 6">'
            f'{content}'
            f'</h6>'
        )


    # --------------------------------------------------------
    # LIST PARAGRAPH
    #
    # Pentru moment este afișat ca paragraf.
    # XML tagging-ul este făcut manual.
    # --------------------------------------------------------

    if "list paragraph" in style_lower:

        return (
            f'<p data-word-style="List Paragraph">'
            f'{content}'
            f'</p>'
        )


    # --------------------------------------------------------
    # NORMAL / BODY TEXT / REST
    # --------------------------------------------------------

    return (
        f'<p data-word-style="{escape_html(style_name)}">'
        f'{content}'
        f'</p>'
    )


# ============================================================
# TABLE -> HTML
# ============================================================

def table_to_html(table):

    output = []

    output.append("<table>")


    for row_index, row in enumerate(table.rows):

        output.append("<tr>")


        for cell in row.cells:

            cell_parts = []


            for paragraph in cell.paragraphs:

                paragraph_html = paragraph_to_html(
                    paragraph
                )

                cell_parts.append(
                    paragraph_html
                )


            cell_html = "".join(
                cell_parts
            )


            if row_index == 0:

                output.append(
                    f"<th>{cell_html}</th>"
                )

            else:

                output.append(
                    f"<td>{cell_html}</td>"
                )


        output.append("</tr>")


    output.append("</table>")


    return "".join(output)


# ============================================================
# INLINE IMAGES
#
# Word document can contain images.
#
# They are displayed only.
# They are NOT turned into XML tags.
# ============================================================

def document_images_html(document):

    images = []


    try:

        for rel in document.part.rels.values():

            target = rel.target_part

            if not hasattr(target, "blob"):
                continue


            content_type = getattr(
                target,
                "content_type",
                ""
            )


            if not content_type.startswith("image/"):
                continue


            encoded = base64.b64encode(
                target.blob
            ).decode("ascii")


            images.append(
                f'<img src="data:{content_type};base64,{encoded}" '
                f'alt="Imagine din document">'
            )

    except Exception:

        pass


    return images


# ============================================================
# DOCX -> HTML
# ============================================================

def docx_to_html(document):

    output = []


    # --------------------------------------------------------
    # PARAGRAPHS
    #
    # python-docx păstrează ordinea logică a paragrafelor.
    # Aceasta este ceea ce ne interesează pentru selecție.
    # --------------------------------------------------------

    for paragraph in document.paragraphs:

        output.append(
            paragraph_to_html(
                paragraph
            )
        )


    # --------------------------------------------------------
    # TABLES
    #
    # Le afișăm în document, dar nu există zonă de tag
    # specială pentru tabele.
    # --------------------------------------------------------

    for table in document.tables:

        output.append(
            table_to_html(table)
        )


    if not output:

        output.append(
            '<p>Documentul nu conține text.</p>'
        )


    return "\n".join(output)


# ============================================================
# UPLOAD WORD
# ============================================================

@app.post("/upload-word")
async def upload_word(
    file: UploadFile = File(...)
):

    # --------------------------------------------------------
    # CHECK FILE
    # --------------------------------------------------------

    if not file.filename:

        raise HTTPException(
            status_code=400,
            detail="Nu a fost selectat niciun fișier."
        )


    original_name = safe_filename(
        file.filename
    )


    # --------------------------------------------------------
    # ONLY DOCX
    # --------------------------------------------------------

    if not original_name.lower().endswith(".docx"):

        raise HTTPException(
            status_code=400,
            detail="Este permis doar formatul Word .docx."
        )


    destination = (
        WORD_UPLOADS_DIR /
        original_name
    )


    # --------------------------------------------------------
    # SAVE
    # --------------------------------------------------------

    try:

        with destination.open("wb") as buffer:

            shutil.copyfileobj(
                file.file,
                buffer
            )

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=(
                "Eroare la salvarea documentului Word: "
                f"{str(e)}"
            )
        )

    finally:

        await file.close()


    # --------------------------------------------------------
    # READ DOCX
    # --------------------------------------------------------

    try:

        document = Document(
            str(destination)
        )

    except Exception as e:

        try:
            destination.unlink()
        except Exception:
            pass

        raise HTTPException(
            status_code=400,
            detail=(
                "Fișierul nu poate fi citit ca document "
                f"Word .docx: {str(e)}"
            )
        )


    # --------------------------------------------------------
    # CONVERT TO CONTROLLED HTML
    # --------------------------------------------------------

    try:

        rendered_html = docx_to_html(
            document
        )

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=(
                "Eroare la transformarea documentului "
                f"Word în format de lucru: {str(e)}"
            )
        )


    # --------------------------------------------------------
    # RESPONSE
    # --------------------------------------------------------

    return {
        "status": "success",
        "filename": original_name,
        "url": (
            f"/files/word/{original_name}"
        ),
        "html": rendered_html
    }


# ============================================================
# HEALTH
# ============================================================

@app.get("/health")
async def health():

    return {
        "status": "ok",
        "application": "xml_tagger",
        "version": "2.0.0"
    }
