from fastapi import FastAPI, UploadFile, File, HTTPException
from pptx import Presentation
from io import BytesIO

app = FastAPI(
    title="RTSM PPTX Text Extractor",
    version="1.0"
)


@app.get("/")
def home():
    return {
        "status": "online",
        "service": "RTSM PPTX Extractor"
    }


@app.post("/extract")
async def extract_pptx(file: UploadFile = File(...)):

    if not file.filename.lower().endswith(".pptx"):
        raise HTTPException(
            status_code=400,
            detail="Only PPTX files are supported."
        )

    file_data = await file.read()

    try:
        presentation = Presentation(BytesIO(file_data))
    except Exception as e:
        raise HTTPException(
            status_code=400,
            detail=f"Unable to read PPTX: {str(e)}"
        )

    slides = []

    for slide_number, slide in enumerate(
        presentation.slides,
        start=1
    ):

        slide_lines = []

        slide_lines.append(
            f"===== SLIDE {slide_number} ====="
        )

        for shape in slide.shapes:

            # Text boxes / titles / placeholders
            if shape.has_text_frame:

                text = shape.text.strip()

                if text:
                    slide_lines.append(text)

            # Tables
            if shape.has_table:

                slide_lines.append(
                    "[TABLE]"
                )

                for row in shape.table.rows:

                    cells = []

                    for cell in row.cells:

                        cell_text = cell.text.strip()

                        cells.append(cell_text)

                    slide_lines.append(
                        " | ".join(cells)
                    )

        slides.append(
            "\n".join(slide_lines)
        )

    extracted_text = "\n\n".join(slides)

    return {
        "success": True,
        "filename": file.filename,
        "slide_count": len(presentation.slides),
        "text": extracted_text
    }
