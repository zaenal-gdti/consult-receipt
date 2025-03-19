from fastapi import FastAPI, File, UploadFile, BackgroundTasks, HTTPException
import pandas as pd
import time
import uuid
import shutil
import zipfile
import os
from script.mail_merge import run_mail_merge
from datetime import datetime
from fastapi.responses import FileResponse

app = FastAPI()

# Dictionary to store job statuses and results
job_status = {}
job_results = {}
job_start_time = {}  # Store job start times

UPLOAD_DIR = "uploads"
ZIP_DIR = "zipped_files"

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(ZIP_DIR, exist_ok=True)

# Background job: Process and zip an Excel file
def process_and_zip(job_id: str, file_path: str):
    try:
        job_status[job_id] = "🔄 Processing Excel file..."
        job_start_time[job_id] = datetime.now()  # Store start time

        # Simulate processing (read file)
        now = datetime.now().strftime('%Y%m%d_%H%M%S')
        file_path_base = os.path.basename(file_path)
        zip_output = f'{file_path_base}_{now}'
        run_mail_merge(file_path, zip_output)
        
        job_status[job_id] = "✅ Completed!"
        job_results[job_id] = f'output/{zip_output}.zip'  # Store zip file path for download

    except Exception as e:
        job_status[job_id] = f"❌ Failed: {str(e)}"

# API to upload an Excel file and start background job
@app.post("/upload-excel/")
async def upload_excel(file: UploadFile = File(...), background_tasks: BackgroundTasks = None):
    job_id = str(uuid.uuid4())  # Unique job ID
    job_status[job_id] = "⏳ Pending..."
    
    # Save uploaded file
    file_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # Define output zip file path

    # Start background task
    background_tasks.add_task(process_and_zip, job_id, file_path)

    return {"message": "File uploaded, processing started!", "job_id": job_id}

# API to check job status
@app.get("/job-status/{job_id}")
async def check_job_status(job_id: str):
    status = job_status.get(job_id, "Job not found!")
    # Calculate elapsed time
    start_time = job_start_time.get(job_id)
    elapsed_time = None
    if start_time:
        elapsed_time = (datetime.now() - start_time).total_seconds()
    
    return {
        "job_id": job_id,
        "status": status,
        "elapsed_time": f"{elapsed_time:.2f} seconds" if elapsed_time else "N/A"
    }

# API to download processed ZIP file
@app.get("/download/{job_id}")
async def download_zip(job_id: str):
    zip_path = job_results.get(job_id)
    if zip_path and os.path.exists(zip_path):
        return FileResponse(zip_path, media_type="application/zip", filename=os.path.basename(zip_path))
    raise HTTPException(status_code=404, detail="File not found or job failed!")