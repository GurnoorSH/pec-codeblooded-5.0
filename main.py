from fastapi import FastAPI, Form, File, UploadFile
from starlette.responses import FileResponse
from utils import html_to_docx
import os
from fastapi.middleware.cors import CORSMiddleware


app = FastAPI()
origins = [
    "http://127.0.0.1",
    "http://127.0.0.1:5500",
    "http://127.0.0.1:50895",
    "http://localhost:5500"

]
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/convert/")
async def convert(file: UploadFile = File(...)):
    html_content = await file.read()
    docx_file_path = html_to_docx(html_content.decode('utf-8'))
    
    return FileResponse(docx_file_path, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename=os.path.basename(docx_file_path))

# /uvicorn main:app --reload