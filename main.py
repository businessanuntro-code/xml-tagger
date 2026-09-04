# ============================================================
# XML TAGGER - POWERPOINT
# ============================================================
# Versiune: 5.0.0
#
# Flux:
# PPTX
#   ↓
# python-pptx
#   ↓
# extragere slide-uri / textbox-uri / texte / imagini
#   ↓
# index.html
#   ↓
# selecție text
#   ↓
# XML
#
# PDF NU mai este folosit ca fișier de upload.
# ============================================================

from fastapi import FastAPI, File, UploadFile, Request
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from pathlib import Path
import uuid
import re
import shutil
import base64

from pptx import Presentation


# ============================================================
# DIRECTOARE
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

TEMPLATES_DIR = BASE_DIR / "templates"

UPLOADS_DIR = BASE_DIR / "uploads"
PPT_UPLOADS_DIR = UPLOADS_DIR / "ppt"
PPT_IMAGES_DIR = UPLOADS_DIR / "ppt_images"

PPT_UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
PPT_IMAGES_DIR.mkdir(parents=True, exist_ok=True)


# ============================================================
# FASTAPI
# ============================================================

app = FastAPI(
    title="XML Tagger - PowerPoint",
    version="5.0.0"
)

templates = Jinja2Templates(
    directory=str(TEMPLATES_DIR)
)


# ============================================================
# STATIC
# ============================================================

app.mount(
    "/files/ppt",
    StaticFiles(directory=str(PPT_UPLOADS_DIR)),
    name="ppt_files"
)

app.mount(
    "/files/ppt_images",
    StaticFiles(directory=str(PPT_IMAGES_DIR)),
    name="ppt_images"
)


# ============================================================
# UTILS
# ============================================================

def safe_filename(filename: str) -> str:
    """
    Curăță numele fișierului.
    """

    filename = filename or "document.pptx"

    filename = Path(filename).name

    filename = re.sub(
        r"[^A-Za-z0-9._-]+",
        "_",
        filename
    )

    if not filename:
        filename = "document.pptx"

    return filename


def rgb_to_hex(color):
    """
    Încearcă să extragă culoarea RGB dintr-un run.
    """

    try:
        if color is None:
            return None

        rgb = color.rgb

        if rgb is None:
            return None

        return "#" + str(rgb)
    except Exception:
        return None


def get_font_size(run):
    """
    Font size în pixeli aproximativi.

    PPTX folosește puncte.
    1 pt ≈ 1.333 px
    """

    try:
        if run.font.size is None:
            return None

        points = run.font.size.pt

        return round(points * 1.3333, 2)

    except Exception:
        return None


def get_run_style(run):
    """
    Extrage stilul unui run.
    """

    return {
        "bold": bool(run.font.bold),
        "italic": bool(run.font.italic),
        "underline": bool(run.font.underline),
        "font_size": get_font_size(run),
        "font_name": run.font.name,
        "color": rgb_to_hex(
            run.font.color
        )
    }


def extract_paragraph(paragraph):
    """
    Extrage un paragraf PowerPoint.

    Păstrăm runs separat pentru a păstra
    cât mai mult din formatarea originală.
    """

    runs = []

    for run in paragraph.runs:

        text = run.text or ""

        if not text:
            continue

        runs.append({
            "text": text,
            "style": get_run_style(run)
        })

    # Dacă nu există runs, încercăm textul
    if not runs:

        text = paragraph.text or ""

        if text:
            runs.append({
                "text": text,
                "style": {
                    "bold": False,
                    "italic": False,
                    "underline": False,
                    "font_size": None,
                    "font_name": None,
                    "color": None
                }
            })

    return {
        "text": paragraph.text or "",
        "runs": runs
    }


def extract_text_frame(shape):
    """
    Extrage structura unui TextBox / shape cu text.
    """

    paragraphs = []

    try:
        for paragraph in shape.text_frame.paragraphs:
            paragraphs.append(
                extract_paragraph(paragraph)
            )
    except Exception:
        pass

    return paragraphs


def save_image(blob: bytes, extension: str, file_id: str, index: int):
    """
    Salvează imaginea extrasă din PPTX.
    """

    extension = extension or "png"

    extension = extension.lower()

    if extension == "jpeg":
        extension = "jpg"

    filename = (
        f"{file_id}_image_{index}.{extension}"
    )

    output = PPT_IMAGES_DIR / filename

    with open(output, "wb") as f:
        f.write(blob)

    return f"/files/ppt_images/{filename}"


# ============================================================
# EXTRAGERE PPTX
# ============================================================

def process_pptx(pptx_path: Path, file_id: str):

    prs = Presentation(
        str(pptx_path)
    )

    # Dimensiunea slide-ului în EMU
    slide_width_emu = prs.slide_width
    slide_height_emu = prs.slide_height

    # Conversie EMU -> px
    # folosim 120 px/inch ca bază de afișare
    EMU_PER_INCH = 914400
    DISPLAY_PX_PER_INCH = 120

    slide_width_px = (
        slide_width_emu / EMU_PER_INCH
    ) * DISPLAY_PX_PER_INCH

    slide_height_px = (
        slide_height_emu / EMU_PER_INCH
    ) * DISPLAY_PX_PER_INCH

    slides = []

    image_index = 0

    for slide_number, slide in enumerate(
        prs.slides,
        start=1
    ):

        shapes = []

        # ----------------------------------------------------
        # SHAPES
        # ----------------------------------------------------

        for shape_index, shape in enumerate(
            slide.shapes
        ):

            try:
                left = (
                    shape.left / EMU_PER_INCH
                ) * DISPLAY_PX_PER_INCH

                top = (
                    shape.top / EMU_PER_INCH
                ) * DISPLAY_PX_PER_INCH

                width = (
                    shape.width / EMU_PER_INCH
                ) * DISPLAY_PX_PER_INCH

                height = (
                    shape.height / EMU_PER_INCH
                ) * DISPLAY_PX_PER_INCH

            except Exception:
                continue

            shape_data = {
                "index": shape_index,
                "type": str(shape.shape_type),
                "left": round(left, 3),
                "top": round(top, 3),
                "width": round(width, 3),
                "height": round(height, 3),
            }

            # ------------------------------------------------
            # TEXT
            # ------------------------------------------------

            if getattr(
                shape,
                "has_text_frame",
                False
            ):

                paragraphs = extract_text_frame(
                    shape
                )

                shape_data["kind"] = "text"
                shape_data["paragraphs"] = paragraphs

                # text complet
                try:
                    shape_data["text"] = (
                        shape.text or ""
                    )
                except Exception:
                    shape_data["text"] = ""

                shapes.append(shape_data)

                continue

            # ------------------------------------------------
            # IMAGE
            # ------------------------------------------------

            if getattr(
                shape,
                "shape_type",
                None
            ) == 13:

                try:

                    image = shape.image

                    blob = image.blob

                    extension = (
                        image.ext
                        or "png"
                    )

                    image_url = save_image(
                        blob,
                        extension,
                        file_id,
                        image_index
                    )

                    image_index += 1

                    shape_data["kind"] = "image"
                    shape_data["url"] = image_url

                    shapes.append(
                        shape_data
                    )

                    continue

                except Exception:
                    pass

            # ------------------------------------------------
            # ALTE SHAPE-URI
            # ------------------------------------------------

            shape_data["kind"] = "shape"

            shapes.append(
                shape_data
            )

        # ----------------------------------------------------
        # SLIDE
        # ----------------------------------------------------

        slides.append({
            "slide": slide_number,
            "width": round(slide_width_px, 3),
            "height": round(slide_height_px, 3),
            "shapes": shapes
        })

    return {
        "width": round(slide_width_px, 3),
        "height": round(slide_height_px, 3),
        "width_emu": slide_width_emu,
        "height_emu": slide_height_emu,
        "slides": slides
    }


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
# UPLOAD PPTX
# ============================================================

@app.post("/upload-ppt/")
async def upload_ppt(
    file: UploadFile = File(...)
):

    original_name = file.filename or ""

    extension = Path(
        original_name
    ).suffix.lower()

    if extension not in (
        ".ppt",
        ".pptx"
    ):

        return JSONResponse(
            status_code=400,
            content={
                "status": "error",
                "message": (
                    "Fișier invalid. "
                    "Încarcă un fișier PPT sau PPTX."
                )
            }
        )

    # --------------------------------------------------------
    # ID
    # --------------------------------------------------------

    file_id = uuid.uuid4().hex

    clean_name = safe_filename(
        original_name
    )

    stored_filename = (
        f"{file_id}_{clean_name}"
    )

    output_path = (
        PPT_UPLOADS_DIR /
        stored_filename
    )

    # --------------------------------------------------------
    # SAVE
    # --------------------------------------------------------

    try:

        with open(
            output_path,
            "wb"
        ) as buffer:

            shutil.copyfileobj(
                file.file,
                buffer
            )

    except Exception as e:

        return JSONResponse(
            status_code=500,
            content={
                "status": "error",
                "message": (
                    "Nu am putut salva "
                    "fișierul.",
                ),
                "detail": str(e)
            }
        )

    # --------------------------------------------------------
    # PPT
    # --------------------------------------------------------

    # python-pptx lucrează nativ cu PPTX.
    # Pentru .ppt vechi, utilizatorul trebuie să-l
    # salveze ca .pptx.
    if extension == ".ppt":

        return JSONResponse(
            status_code=400,
            content={
                "status": "error",
                "message": (
                    "Formatul PPT vechi nu este "
                    "procesat direct. "
                    "Salvează documentul ca PPTX "
                    "și încarcă din nou."
                )
            }
        )

    # --------------------------------------------------------
    # PROCESS
    # --------------------------------------------------------

    try:

        presentation = process_pptx(
            output_path,
            file_id
        )

    except Exception as e:

        try:
            output_path.unlink(
                missing_ok=True
            )
        except Exception:
            pass

        return JSONResponse(
            status_code=500,
            content={
                "status": "error",
                "message": (
                    "Nu am putut procesa "
                    "fișierul PPTX."
                ),
                "detail": str(e)
            }
        )

    # --------------------------------------------------------
    # RESPONSE
    # --------------------------------------------------------

    return JSONResponse(
        content={
            "status": "success",
            "filename": original_name,
            "stored_filename": stored_filename,
            "url": (
                f"/files/ppt/"
                f"{stored_filename}"
            ),
            "slide_count": len(
                presentation["slides"]
            ),
            "presentation": presentation
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
        "version": "5.0.0",
        "format": "PPTX"
    }
