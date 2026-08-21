from fastapi import FastAPI, Request, HTTPException
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
async def extract_pptx(request: Request):

    file_data = await request.body()

    if not file_data:
        raise HTTPException(
            status_code=400,
            detail="No PPTX file received."
        )

    try:
        prs = Presentation(BytesIO(file_data))

        slides = []

        for slide_number, slide in enumerate(prs.slides, start=1):

            slide_text = []

            for shape in slide.shapes:

                if hasattr(shape, "text") and shape.text:
                    slide_text.append(shape.text.strip())

            slides.append({
                "slide": slide_number,
                "text": "\n".join(slide_text)
            })

        return {
            "success": True,
            "filename": "uploaded.pptx",
            "slide_count": len(slides),
            "slides": slides
        }

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=f"PPTX extraction failed: {str(e)}"
        )
