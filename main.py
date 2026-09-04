from pathlib import Path
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
from docx.document import Document as _Document
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl


# ============================================================
# APPLICATION
# ============================================================

app = FastAPI(
    title="XML Tagger",
    version="2.1.0"
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
# RUN -> HTML
# ============================================================

def run_to_html(run):

    text = run.text or ""

    if not text:
        return ""


    value = escape_html(text)


    # --------------------------------------------------------
    # SUPERSCRIPT
    # --------------------------------------------------------

    if run.font.superscript:

        value = f"<sup>{value}</sup>"


    # --------------------------------------------------------
    # SUBSCRIPT
    # --------------------------------------------------------

    elif run.font.subscript:

        value = f"<sub>{value}</sub>"


    # --------------------------------------------------------
    # BOLD
    # --------------------------------------------------------

    if run.bold:

        value = f"<strong>{value}</strong>"


    # --------------------------------------------------------
    # ITALIC
    # --------------------------------------------------------

    if run.italic:

        value = f"<em>{value}</em>"


    # --------------------------------------------------------
    # UNDERLINE
    # --------------------------------------------------------

    if run.underline:

        value = f"<u>{value}</u>"


    return value


# ============================================================
# PARAGRAPH -> HTML
# ============================================================

def paragraph_to_html(paragraph: Paragraph):

    text = paragraph.text or ""

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
    # FALLBACK
    # --------------------------------------------------------

    if not content and text:

        content = escape_html(text)


    # --------------------------------------------------------
    # EMPTY PARAGRAPH
    # --------------------------------------------------------

    if not content:

        return (
            '<p class="word-empty-paragraph">'
            '<br>'
            '</p>'
        )


    style_lower = style_name.lower()


    # ========================================================
    # TITLE
    # ========================================================

    if style_lower == "title":

        return (
            '<h1 '
            'class="word-title" '
            'data-word-style="Title">'
            f'{content}'
            '</h1>'
        )


    # ========================================================
    # SUBTITLE
    # ========================================================

    if style_lower == "subtitle":

        return (
            '<h2 '
            'class="word-subtitle" '
            'data-word-style="Subtitle">'
            f'{content}'
            '</h2>'
        )


    # ========================================================
    # HEADING 1
    # ========================================================

    if style_lower == "heading 1":

        return (
            '<h2 '
            'class="word-heading word-heading-1" '
            'data-word-style="Heading 1">'
            f'{content}'
            '</h2>'
        )


    # ========================================================
    # HEADING 2
    # ========================================================

    if style_lower == "heading 2":

        return (
            '<h3 '
            'class="word-heading word-heading-2" '
            'data-word-style="Heading 2">'
            f'{content}'
            '</h3>'
        )


    # ========================================================
    # HEADING 3
    # ========================================================

    if style_lower == "heading 3":

        return (
            '<h3 '
            'class="word-heading word-heading-3" '
            'data-word-style="Heading 3">'
            f'{content}'
            '</h3>'
        )


    # ========================================================
    # HEADING 4
    # ========================================================

    if style_lower == "heading 4":

        return (
            '<h4 '
            'class="word-heading word-heading-4" '
            'data-word-style="Heading 4">'
            f'{content}'
            '</h4>'
        )


    # ========================================================
    # HEADING 5
    # ========================================================

    if style_lower == "heading 5":

        return (
            '<h5 '
            'class="word-heading word-heading-5" '
            'data-word-style="Heading 5">'
            f'{content}'
            '</h5>'
        )


    # ========================================================
    # HEADING 6
    # ========================================================

    if style_lower == "heading 6":

        return (
            '<h6 '
            'class="word-heading word-heading-6" '
            'data-word-style="Heading 6">'
            f'{content}'
            '</h6>'
        )


    # ========================================================
    # LIST PARAGRAPH
    # ========================================================

    if "list paragraph" in style_lower:

        return (
            '<p '
            'class="word-list-paragraph" '
            'data-word-style="List Paragraph">'
            f'{content}'
            '</p>'
        )


    # ========================================================
    # NORMAL / BODY TEXT / OTHER
    # ========================================================

    return (
        '<p '
        'class="word-paragraph" '
        f'data-word-style="{escape_html(style_name)}">'
        f'{content}'
        '</p>'
    )


# ============================================================
# TABLE -> HTML
# ============================================================

def table_to_html(table: Table):

    output = []

    output.append(
        '<div class="word-table-wrapper">'
    )

    output.append("<table>")


    for row_index, row in enumerate(table.rows):

        output.append("<tr>")


        for cell in row.cells:

            cell_content = []


            for paragraph in cell.paragraphs:

                cell_content.append(
                    paragraph_to_html(
                        paragraph
                    )
                )


            content = "".join(
                cell_content
            )


            if row_index == 0:

                output.append(
                    f"<th>{content}</th>"
                )

            else:

                output.append(
                    f"<td>{content}</td>"
                )


        output.append("</tr>")


    output.append("</table>")

    output.append(
        "</div>"
    )


    return "".join(output)


# ============================================================
# DOCUMENT BLOCKS
#
# FOARTE IMPORTANT:
#
# document.paragraphs + document.tables NU păstrează ordinea.
#
# Funcția de mai jos parcurge XML-ul intern Word și păstrează
# ordinea reală:
#
# paragraph
# paragraph
# table
# paragraph
# paragraph
# etc.
# ============================================================

def iter_block_items(parent):

    if isinstance(parent, _Document):

        parent_elm = parent.element.body

    elif isinstance(parent, _Cell):

        parent_elm = parent._tc

    else:

        raise ValueError(
            "Parent necunoscut pentru iter_block_items."
        )


    for child in parent_elm.iterchildren():

        if isinstance(child, CT_P):

            yield Paragraph(
                child,
                parent
            )

        elif isinstance(child, CT_Tbl):

            yield Table(
                child,
                parent
            )


# ============================================================
# DOCX -> HTML
# ============================================================

def docx_to_html(document: Document):

    output = []


    # --------------------------------------------------------
    # PARCURGEM DOCUMENTUL ÎN ORDINEA REALĂ
    # --------------------------------------------------------

    for block in iter_block_items(document):

        if isinstance(block, Paragraph):

            output.append(
                paragraph_to_html(
                    block
                )
            )

        elif isinstance(block, Table):

            output.append(
                table_to_html(
                    block
                )
            )


    # --------------------------------------------------------
    # EMPTY DOCUMENT
    # --------------------------------------------------------

    if not output:

        output.append(
            '<p class="word-paragraph">'
            'Documentul nu conține text.'
            '</p>'
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
    # EXTENSION
    # --------------------------------------------------------

    if not original_name.lower().endswith(".docx"):

        raise HTTPException(
            status_code=400,
            detail=(
                "Este permis doar formatul "
                "Microsoft Word .docx."
            )
        )


    destination = (
        WORD_UPLOADS_DIR /
        original_name
    )


    # --------------------------------------------------------
    # SAVE FILE
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
                "Eroare la salvarea fișierului Word: "
                f"{str(e)}"
            )
        )

    finally:

        await file.close()


    # --------------------------------------------------------
    # OPEN DOCX
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
                "Fișierul nu poate fi deschis ca document "
                f"Word .docx: {str(e)}"
            )
        )


    # --------------------------------------------------------
    # CONVERT
    # --------------------------------------------------------

    try:

        rendered_html = docx_to_html(
            document
        )

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=(
                "Documentul Word a fost încărcat, dar "
                "nu a putut fi afișat: "
                f"{str(e)}"
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
        "version": "2.1.0"
    }
