import os
import json
import base64
import logging
import zipfile
from io import BytesIO

from fastapi import FastAPI, Request, HTTPException
from pptx import Presentation

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("rtsm-ppt-extractor")

app = FastAPI(
    title="RTSM PPTX Text Extractor",
    version="2.0"
)


@app.get("/")
def home():
    return {
        "status": "online",
        "service": "RTSM PPTX Extractor",
        "endpoint": "/extract",
        "version": "2.0"
    }


def is_valid_pptx(data: bytes) -> bool:
    """Check that bytes look like a real PPTX ZIP package."""
    if not data or not data.startswith(b"PK"):
        return False

    try:
        with zipfile.ZipFile(BytesIO(data)) as z:
            names = set(z.namelist())
            return "[Content_Types].xml" in names and any(
                n.startswith("ppt/") for n in names
            )
    except Exception:
        return False


def decode_base64_candidate(value):
    """Decode a base64 candidate only when it produces a valid PPTX."""
    if not isinstance(value, str):
        return None

    value = value.strip()

    # Remove a possible data URL prefix.
    if value.startswith("data:") and "," in value:
        value = value.split(",", 1)[1]

    # Remove whitespace/newlines commonly present in base64.
    value = "".join(value.split())

    if not value:
        return None

    try:
        decoded = base64.b64decode(value, validate=False)
    except Exception:
        return None

    if is_valid_pptx(decoded):
        return decoded

    return None


def find_pptx_in_json(obj):
    """
    Power Automate can wrap file content inside JSON, commonly using
    $content / $contentBytes / contentBytes. Search recursively so the
    endpoint remains tolerant of different HTTP body shapes.
    """
    if isinstance(obj, dict):
        preferred_keys = [
            "$content",
            "$contentBytes",
            "contentBytes",
            "fileContent",
            "file_content",
            "content",
            "body",
        ]

        for key in preferred_keys:
            if key in obj:
                found = find_pptx_in_json(obj[key])
                if found:
                    return found

        for value in obj.values():
            found = find_pptx_in_json(value)
            if found:
                return found

    elif isinstance(obj, list):
        for item in obj:
            found = find_pptx_in_json(item)
            if found:
                return found

    elif isinstance(obj, str):
        return decode_base64_candidate(obj)

    return None


def extract_file_bytes(raw_body: bytes, content_type: str) -> bytes:
    """
    Accept all common Power Automate HTTP body forms:
    1. Raw PPTX binary
    2. Base64 text
    3. JSON containing base64 file content
    4. Multipart/form-data containing a PPTX
    """
    if is_valid_pptx(raw_body):
        return raw_body

    # If multipart/form-data or another wrapper contains the ZIP payload,
    # locate the PPTX ZIP header and validate the candidate.
    zip_start = raw_body.find(b"PK\x03\x04")
    if zip_start > 0:
        candidate = raw_body[zip_start:]
        if is_valid_pptx(candidate):
            return candidate

    # JSON body from Power Automate.
    try:
        text = raw_body.decode("utf-8")
        parsed = json.loads(text)
        found = find_pptx_in_json(parsed)
        if found:
            return found

        # JSON may simply be a quoted base64 string.
        found = decode_base64_candidate(parsed) if isinstance(parsed, str) else None
        if found:
            return found
    except Exception:
        pass

    # Plain base64 body.
    found = decode_base64_candidate(raw_body.decode("utf-8", errors="ignore"))
    if found:
        return found

    raise HTTPException(
        status_code=400,
        detail={
            "error": "Invalid PPTX payload",
            "message": "The /extract endpoint received data, but it could not find a valid PPTX file.",
            "content_type": content_type,
            "received_bytes": len(raw_body),
            "first_bytes": raw_body[:20].hex()
        }
    )


def extract_slide_content(prs):
    slides = []

    for slide_number, slide in enumerate(prs.slides, start=1):
        slide_lines = []

        for shape in slide.shapes:
            # Text boxes, titles and placeholders
            if getattr(shape, "has_text_frame", False):
                text = shape.text.strip()
                if text:
                    slide_lines.append(text)

            # Tables
            if getattr(shape, "has_table", False):
                table_lines = []

                for row in shape.table.rows:
                    cells = []
                    for cell in row.cells:
                        cells.append(cell.text.strip())

                    if any(cells):
                        table_lines.append(" | ".join(cells))

                if table_lines:
                    slide_lines.append("[TABLE]")
                    slide_lines.extend(table_lines)

        slides.append({
            "slide": slide_number,
            "text": "\n".join(slide_lines)
        })

    return slides


@app.post("/extract")
async def extract_pptx(request: Request):
    content_type = request.headers.get("content-type", "")
    raw_body = await request.body()

    logger.info(
        "POST /extract received: content_type=%s bytes=%d first_bytes=%s",
        content_type,
        len(raw_body),
        raw_body[:20].hex()
    )

    if not raw_body:
        raise HTTPException(
            status_code=400,
            detail="No file content received in request body."
        )

    try:
        file_data = extract_file_bytes(raw_body, content_type)

        logger.info(
            "PPTX payload decoded successfully: %d bytes",
            len(file_data)
        )

        prs = Presentation(BytesIO(file_data))
        slides = extract_slide_content(prs)

        logger.info(
            "PPTX extraction successful: slides=%d",
            len(slides)
        )

        return {
            "success": True,
            "filename": "uploaded.pptx",
            "slide_count": len(slides),
            "slides": slides
        }

    except HTTPException:
        raise

    except Exception as e:
        logger.exception("PPTX extraction failed")

        # Return the actual exception to Power Automate instead of hiding it
        # behind a generic 500. This makes the next failure diagnosable.
        raise HTTPException(
            status_code=500,
            detail={
                "error": "PPTX extraction failed",
                "message": str(e),
                "content_type": content_type,
                "received_bytes": len(raw_body)
            }
        )


if __name__ == "__main__":
    import uvicorn

    port = int(os.environ.get("PORT", "10000"))
    uvicorn.run(
        app,
        host="0.0.0.0",
        port=port
    )
