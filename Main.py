import io
import os
import logging
import zipfile
from typing import Optional

from fastapi import FastAPI, Request, UploadFile, File
from fastapi.responses import JSONResponse
from pptx import Presentation


# ============================================================
# APPLICATION SETUP
# ============================================================

app = FastAPI(
    title="Tata RTSM PPT Extractor",
    version="2.0"
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

logger = logging.getLogger("rtm-ppt-extractor")


# ============================================================
# HEALTH CHECK
# ============================================================

@app.get("/")
async def health_check():
    return {
        "status": "ok",
        "service": "tata-rtsm-ppt-extractor",
        "version": "2.0",
        "message": "Your service is live 🎉"
    }


# ============================================================
# BASIC PPTX VALIDATION
# ============================================================

def validate_pptx(data: bytes):

    if not data:
        raise ValueError("Request body is empty.")

    logger.info(
        "Received file size: %d bytes",
        len(data)
    )

    # PPTX is a ZIP container and normally starts with PK
    if not data.startswith(b"PK"):
        logger.warning(
            "File does not start with PK ZIP signature. First 16 bytes: %s",
            data[:16].hex()
        )

        raise ValueError(
            "Uploaded file is not a valid PPTX ZIP package."
        )

    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            bad_file = z.testzip()

            if bad_file:
                raise ValueError(
                    f"PPTX ZIP package is corrupted: {bad_file}"
                )

            names = z.namelist()

            if "[Content_Types].xml" not in names:
                raise ValueError(
                    "Invalid Office file: [Content_Types].xml missing."
                )

            if not any(
                name.startswith("ppt/")
                for name in names
            ):
                raise ValueError(
                    "Invalid PPTX: ppt/ directory not found."
                )

    except zipfile.BadZipFile:
        raise ValueError(
            "The uploaded file is not a valid ZIP/PPTX file."
        )


# ============================================================
# TEXT EXTRACTION
# ============================================================

def extract_shape_text(shape):

    result = []

    # Normal text box / title / placeholder
    if hasattr(shape, "text"):

        text = shape.text.strip()

        if text:
            result.append(text)

    # Tables
    if getattr(shape, "has_table", False):

        table_data = []

        for row in shape.table.rows:

            row_data = []

            for cell in row.cells:

                cell_text = cell.text.strip()

                row_data.append(cell_text)

            table_data.append(row_data)

        if table_data:

            result.append(
                "\n".join(
                    " | ".join(row)
                    for row in table_data
                )
            )

    return result


# ============================================================
# PPTX EXTRACTION
# ============================================================

def extract_presentation(data: bytes):

    prs = Presentation(io.BytesIO(data))

    slides_output = []

    total_characters = 0

    for slide_number, slide in enumerate(
        prs.slides,
        start=1
    ):

        slide_text_parts = []

        title = ""

        # ----------------------------------------------------
        # Find title
        # ----------------------------------------------------

        for shape in slide.shapes:

            if not hasattr(shape, "text"):
                continue

            text = shape.text.strip()

            if not text:
                continue

            try:

                if (
                    shape.is_placeholder
                    and shape.placeholder_format.type == 1
                ):
                    title = text
                    break

            except Exception:
                pass

        # ----------------------------------------------------
        # Extract all slide content
        # ----------------------------------------------------

        for shape in slide.shapes:

            try:

                extracted = extract_shape_text(shape)

                for item in extracted:

                    if item:
                        slide_text_parts.append(item)

            except Exception as shape_error:

                logger.warning(
                    "Slide %d shape extraction failed: %s",
                    slide_number,
                    str(shape_error)
                )

        # Remove duplicate text while preserving order
        cleaned_parts = []

        for item in slide_text_parts:

            if item not in cleaned_parts:
                cleaned_parts.append(item)

        slide_text = "\n".join(cleaned_parts).strip()

        total_characters += len(slide_text)

        slides_output.append(
            {
                "slide_number": slide_number,
                "title": title,
                "text": slide_text
            }
        )

    return {
        "success": True,
        "file_type": "PPTX",
        "slide_count": len(slides_output),
        "total_characters": total_characters,
        "slides": slides_output
    }


# ============================================================
# FILE EXTRACTION HELPER
# ============================================================

async def get_request_file(
    request: Request,
    file: Optional[UploadFile]
):

    # --------------------------------------------------------
    # Case 1:
    # Multipart upload
    # --------------------------------------------------------

    if file is not None:

        logger.info(
            "Multipart file received: %s",
            file.filename
        )

        data = await file.read()

        logger.info(
            "Multipart file bytes: %d",
            len(data)
        )

        return data

    # --------------------------------------------------------
    # Case 2:
    # Raw binary upload
    #
    # THIS IS THE FORMAT CURRENT POWER AUTOMATE IS SENDING
    # --------------------------------------------------------

    data = await request.body()

    content_type = request.headers.get(
        "content-type",
        ""
    )

    logger.info(
        "Raw request received | content-type=%s | bytes=%d",
        content_type,
        len(data)
    )

    if data:

        logger.info(
            "First 16 bytes: %s",
            data[:16].hex()
        )

    return data


# ============================================================
# MAIN EXTRACT ENDPOINT
# ============================================================

@app.post("/extract")
async def extract_ppt(
    request: Request,
    file: Optional[UploadFile] = File(None)
):

    logger.info("=" * 70)
    logger.info("RTSM PPT EXTRACT REQUEST STARTED")
    logger.info("=" * 70)

    try:

        content_type = request.headers.get(
            "content-type",
            ""
        )

        content_length = request.headers.get(
            "content-length",
            "unknown"
        )

        logger.info(
            "Content-Type: %s",
            content_type
        )

        logger.info(
            "Content-Length: %s",
            content_length
        )

        # ----------------------------------------------------
        # Get actual PPTX bytes
        # ----------------------------------------------------

        data = await get_request_file(
            request=request,
            file=file
        )

        # ----------------------------------------------------
        # Validate
        # ----------------------------------------------------

        validate_pptx(data)

        # ----------------------------------------------------
        # Extract
        # ----------------------------------------------------

        logger.info(
            "Starting PPTX extraction..."
        )

        result = extract_presentation(data)

        logger.info(
            "Extraction completed successfully."
        )

        logger.info(
            "Slides extracted: %d",
            result["slide_count"]
        )

        logger.info(
            "Characters extracted: %d",
            result["total_characters"]
        )

        logger.info("=" * 70)
        logger.info("RTSM PPT EXTRACT REQUEST COMPLETED")
        logger.info("=" * 70)

        return JSONResponse(
            status_code=200,
            content=result
        )

    except ValueError as e:

        logger.error(
            "VALIDATION ERROR: %s",
            str(e)
        )

        return JSONResponse(
            status_code=400,
            content={
                "success": False,
                "error_type": "ValidationError",
                "message": str(e)
            }
        )

    except zipfile.BadZipFile as e:

        logger.error(
            "BAD PPTX ZIP: %s",
            str(e)
        )

        return JSONResponse(
            status_code=400,
            content={
                "success": False,
                "error_type": "InvalidPPTX",
                "message": "The uploaded file is not a valid PPTX."
            }
        )

    except Exception as e:

        logger.exception(
            "UNEXPECTED EXTRACTION ERROR"
        )

        return JSONResponse(
            status_code=500,
            content={
                "success": False,
                "error_type": "InternalServerError",
                "message": str(e)
            }
        )


# ============================================================
# OPTIONAL DEBUG ENDPOINT
# ============================================================

@app.get("/status")
async def status():

    return {
        "service": "tata-rtsm-ppt-extractor",
        "status": "running",
        "version": "2.0",
        "extract_endpoint": "/extract"
    }
