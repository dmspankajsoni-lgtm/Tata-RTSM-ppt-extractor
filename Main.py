import io
import os
import re
import zipfile
import logging
from typing import Optional

from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import JSONResponse
from pptx import Presentation


# ============================================================
# APPLICATION
# ============================================================

app = FastAPI(
    title="Tata RTSM PPT Extractor",
    version="3.0"
)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("rtms-ppt-extractor")


# ============================================================
# CONFIGURATION
# ============================================================

MAX_FILE_SIZE = 50 * 1024 * 1024  # 50 MB

# ZIP/PPTX signatures
PPTX_SIGNATURE = b"PK\x03\x04"

# Old Microsoft PowerPoint .PPT OLE signature
PPT_OLE_SIGNATURE = bytes.fromhex("D0CF11E0A1B11AE1")


# ============================================================
# BASIC ROUTES
# ============================================================

@app.get("/")
async def root():
    return {
        "status": "live",
        "service": "Tata RTSM PPT Extractor",
        "version": "3.0",
        "accepted_files": [".pptx", ".ppt"],
        "endpoint": "/extract"
    }


@app.get("/health")
async def health():
    return {
        "status": "healthy",
        "service": "Tata RTSM PPT Extractor"
    }


# ============================================================
# FILE TYPE DETECTION
# ============================================================

def find_pptx_signature(data: bytes) -> int:
    """
    Find the PK ZIP signature.

    Normal PPTX:
        50 4B 03 04

    Your Power Automate request currently appears to contain:
        0A 50 4B 03 04

    Therefore we search for PK within the beginning of the file
    and remove any unwanted leading bytes.
    """

    if not data:
        return -1

    # Normal PPTX
    if data.startswith(PPTX_SIGNATURE):
        return 0

    # Search only the first 1 MB.
    # This protects against treating arbitrary binary data as PPTX.
    search_area = data[:1024 * 1024]

    position = search_area.find(PPTX_SIGNATURE)

    if position == -1:
        return -1

    return position


def normalize_pptx(data: bytes) -> bytes:
    """
    Remove accidental bytes before the PPTX ZIP package.

    Example received:
        0A 50 4B 03 04 ...

    Converted to:
        50 4B 03 04 ...
    """

    position = find_pptx_signature(data)

    if position == -1:
        raise ValueError(
            "PPTX ZIP signature PK0304 was not found."
        )

    if position > 0:
        logger.warning(
            "PPTX contains %d leading byte(s) before ZIP signature. "
            "Removing them.",
            position
        )

    return data[position:]


def detect_file_type(data: bytes) -> str:
    """
    Detect PowerPoint format from file signature.
    """

    if not data:
        return "unknown"

    # PPTX
    if find_pptx_signature(data) >= 0:
        return "pptx"

    # Old PPT
    if data.startswith(PPT_OLE_SIGNATURE):
        return "ppt"

    return "unknown"


# ============================================================
# PPTX VALIDATION
# ============================================================

def validate_pptx(data: bytes) -> bytes:
    """
    Validate that the supplied bytes are actually a PPTX ZIP package.
    """

    normalized = normalize_pptx(data)

    try:
        with zipfile.ZipFile(io.BytesIO(normalized), "r") as z:

            if not z.testzip() is None:
                raise ValueError(
                    "PPTX ZIP package contains corrupted data."
                )

            names = z.namelist()

            # These are important PowerPoint package files.
            has_content_types = "[Content_Types].xml" in names
            has_presentation = "ppt/presentation.xml" in names

            if not has_content_types or not has_presentation:
                raise ValueError(
                    "ZIP package is not a valid PowerPoint PPTX package."
                )

    except zipfile.BadZipFile as exc:
        raise ValueError(
            "Uploaded file is not a valid PPTX ZIP package."
        ) from exc

    return normalized


# ============================================================
# PPTX TEXT EXTRACTION
# ============================================================

def extract_text_from_pptx(data: bytes) -> dict:
    """
    Extract text from all slides.
    """

    presentation = Presentation(io.BytesIO(data))

    slides = []
    full_text = []

    for slide_number, slide in enumerate(
        presentation.slides,
        start=1
    ):

        slide_text = []

        for shape in slide.shapes:

            # Normal text boxes / placeholders
            if hasattr(shape, "text"):
                text = shape.text

                if text:
                    text = text.strip()

                    if text:
                        slide_text.append(text)

            # Tables
            if getattr(shape, "has_table", False):

                for row in shape.table.rows:

                    row_values = []

                    for cell in row.cells:

                        value = cell.text.strip()

                        if value:
                            row_values.append(value)

                    if row_values:
                        slide_text.append(
                            " | ".join(row_values)
                        )

        slide_result = {
            "slide_number": slide_number,
            "text": "\n".join(slide_text)
        }

        slides.append(slide_result)

        if slide_text:
            full_text.append(
                f"SLIDE {slide_number}\n"
                + "\n".join(slide_text)
            )

    return {
        "slide_count": len(slides),
        "slides": slides,
        "text": "\n\n".join(full_text)
    }


# ============================================================
# REQUEST BODY READER
# ============================================================

async def read_raw_body(request: Request) -> bytes:
    """
    Read the raw binary body sent by Power Automate.

    Power Automate's 'Get file content' is expected to send
    application/octet-stream.
    """

    body = await request.body()

    if not body:
        raise HTTPException(
            status_code=400,
            detail="No file content received."
        )

    if len(body) > MAX_FILE_SIZE:
        raise HTTPException(
            status_code=413,
            detail=(
                f"File too large. Maximum allowed size is "
                f"{MAX_FILE_SIZE // (1024 * 1024)} MB."
            )
        )

    return body


# ============================================================
# MAIN EXTRACTION ENDPOINT
# ============================================================

@app.post("/extract")
async def extract(request: Request):

    logger.info("=" * 70)
    logger.info("RTSM PPT EXTRACT REQUEST STARTED")
    logger.info("=" * 70)

    try:

        # ----------------------------------------------------
        # READ RAW FILE
        # ----------------------------------------------------

        body = await read_raw_body(request)

        logger.info(
            "Content-Type: %s",
            request.headers.get("content-type")
        )

        logger.info(
            "Content-Length: %s",
            request.headers.get("content-length")
        )

        logger.info(
            "Received file size: %d bytes",
            len(body)
        )

        logger.info(
            "First 32 bytes: %s",
            body[:32].hex()
        )

        # ----------------------------------------------------
        # DETECT FILE TYPE
        # ----------------------------------------------------

        file_type = detect_file_type(body)

        logger.info(
            "Detected file type: %s",
            file_type
        )

        # ----------------------------------------------------
        # ONLY POWERPOINT ALLOWED
        # ----------------------------------------------------

        if file_type not in ("pptx", "ppt"):

            logger.error(
                "VALIDATION ERROR: File is not PowerPoint."
            )

            raise HTTPException(
                status_code=400,
                detail=(
                    "Only PowerPoint files (.ppt/.pptx) are "
                    "accepted. PDF, Excel, Word, images and "
                    "other file types are ignored."
                )
            )

        # ----------------------------------------------------
        # OLD .PPT
        # ----------------------------------------------------

        if file_type == "ppt":

            logger.error(
                "Old binary .PPT detected."
            )

            return JSONResponse(
                status_code=400,
                content={
                    "success": False,
                    "error": (
                        "Old .ppt format detected. "
                        "Please convert the file to .pptx "
                        "before sending it to this service."
                    )
                }
            )

        # ----------------------------------------------------
        # NORMALIZE PPTX
        # ----------------------------------------------------

        logger.info(
            "Validating PPTX ZIP package..."
        )

        pptx_data = validate_pptx(body)

        logger.info(
            "PPTX validation successful."
        )

        # ----------------------------------------------------
        # EXTRACT
        # ----------------------------------------------------

        logger.info(
            "Extracting PowerPoint content..."
        )

        result = extract_text_from_pptx(
            pptx_data
        )

        logger.info(
            "Extraction successful."
        )

        logger.info(
            "Slides extracted: %d",
            result["slide_count"]
        )

        logger.info("=" * 70)
        logger.info("RTSM PPT EXTRACT REQUEST COMPLETED")
        logger.info("=" * 70)

        # ----------------------------------------------------
        # RESPONSE
        # ----------------------------------------------------

        return {
            "success": True,
            "file_type": "pptx",
            "file_size": len(body),
            "slide_count": result["slide_count"],
            "slides": result["slides"],
            "text": result["text"]
        }

    except HTTPException:
        raise

    except ValueError as exc:

        logger.error(
            "VALIDATION ERROR: %s",
            str(exc)
        )

        raise HTTPException(
            status_code=400,
            detail=str(exc)
        )

    except Exception as exc:

        logger.exception(
            "UNEXPECTED EXTRACTION ERROR"
        )

        raise HTTPException(
            status_code=500,
            detail=(
                "PowerPoint extraction failed: "
                + str(exc)
            )
        )


# ============================================================
# LOCAL DEVELOPMENT
# ============================================================

if __name__ == "__main__":

    import uvicorn

    port = int(
        os.environ.get(
            "PORT",
            "10000"
        )
    )

    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=port
    )
