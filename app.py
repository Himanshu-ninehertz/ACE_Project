from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
import os
from test import measure_ducts, generate_pdf_report


app = FastAPI()

# static directory for frontend
app.mount("/static", StaticFiles(directory="static"), name="static")

UPLOAD_DIR = "uploads"
OUTPUT_DIR = "outputs"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


@app.get("/")
def root():
    """Serve the HTML UI."""
    with open("static/index.html") as f:
        return HTMLResponse(f.read())


@app.post("/process")
async def process_pdf(file: UploadFile = File(...)):
    """Upload a PDF, process ducts, export Excel"""
    input_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(input_path, "wb") as f:
        f.write(await file.read())

    # run duct measurement
    results = measure_ducts(input_path)

    # Generate PDF report
    output_path = os.path.join(OUTPUT_DIR, file.filename.replace(".pdf", "_report.pdf"))
    generate_pdf_report(results, output_path, input_path, out_img_dir=OUTPUT_DIR)

    return {"results": results, "pdf_url": f"/download/{os.path.basename(output_path)}"}


@app.get("/download/{fname}")
def download_file(fname: str):
    path = os.path.join(OUTPUT_DIR, fname)
    media_type = "application/pdf" if fname.endswith(".pdf") else "application/octet-stream"
    return FileResponse(path, media_type=media_type)
