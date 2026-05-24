import os
import uuid
from pathlib import Path

from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pptx import Presentation


app = FastAPI(title="PPT Translator")

BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
RESULT_DIR = BASE_DIR / "results"
STATIC_DIR = BASE_DIR / "static"

UPLOAD_DIR.mkdir(exist_ok=True)
RESULT_DIR.mkdir(exist_ok=True)
STATIC_DIR.mkdir(exist_ok=True)


@app.get("/", response_class=HTMLResponse)
def home():
    index_path = STATIC_DIR / "index.html"

    if index_path.exists():
        return index_path.read_text(encoding="utf-8")

    return """
    <html>
        <body>
            <h2>PPT Translator is running.</h2>
            <p>Please create static/index.html.</p>
        </body>
    </html>
    """


@app.post("/api/translate-ppt")
async def translate_ppt(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pptx"):
        return JSONResponse(
            status_code=400,
            content={
                "status": "failed",
                "message": "Only .pptx files are supported."
            }
        )

    original_name = Path(file.filename).stem
    safe_name = original_name.replace(" ", "_")
    
    file_id = str(uuid.uuid4())
    
    input_filename = f"{file_id}.pptx"
    output_filename = f"{safe_name}_en.pptx"
    
    input_path = UPLOAD_DIR / input_filename
    output_path = RESULT_DIR / output_filename
    
    content = await file.read()
    input_path.write_bytes(content)

    try:
        prs = Presentation(str(input_path))

        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    original_text = shape.text.strip()

                    if original_text:
                        translated_text = fake_translate(original_text)
                        shape.text = translated_text
                        simple_reduce_font(shape, original_text, translated_text)

        prs.save(str(output_path))

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "status": "failed",
                "message": f"PPT processing failed: {str(e)}"
            }
        )

    return {
        "status": "success",
        "download_url": f"/api/download/{output_filename}"
    }


@app.get("/api/download/{filename}")
def download_file(filename: str):
    file_path = RESULT_DIR / filename

    if not file_path.exists():
        return JSONResponse(
            status_code=404,
            content={
                "status": "failed",
                "message": "File not found."
            }
        )

    return FileResponse(
        path=str(file_path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )


def fake_translate(text: str) -> str:
    """
    第一版临时翻译函数。
    现在先不接 GPT / 豆包 API，只用它验证：
    1. 上传是否成功；
    2. PPT 是否能被读取；
    3. 文本是否能被替换；
    4. 新 PPT 是否能下载。

    后面再把这个函数替换成真正的大模型翻译。
    """
    return "[EN] " + text


def simple_reduce_font(shape, source_text: str, translated_text: str):
    """
    最简单的字号缩小规则。
    如果英文比中文长很多，就把字号稍微缩小。
    """
    ratio = len(translated_text) / max(len(source_text), 1)

    if ratio > 3:
        reduce_size = 6
    elif ratio > 2:
        reduce_size = 4
    elif ratio > 1.5:
        reduce_size = 2
    else:
        reduce_size = 0

    if reduce_size == 0:
        return

    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            if run.font.size:
                new_size = run.font.size.pt - reduce_size
                if new_size < 10:
                    new_size = 10
                run.font.size = int(new_size * 12700)
