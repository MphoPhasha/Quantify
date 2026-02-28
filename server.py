from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
import os
import shutil
import uuid
import zipfile
from typing import List

# Import our backend logic
from quantify import run_quantification

app = FastAPI()

# Temporary directory for uploads
UPLOAD_BASE_DIR = "temp_uploads"
if not os.path.exists(UPLOAD_BASE_DIR):
    os.makedirs(UPLOAD_BASE_DIR)

app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/")
def read_root():
    return FileResponse("static/index.html")

@app.post("/api/quantify")
async def quantify_endpoint(
    network_type: str = Form(...),
    mode: str = Form(...),
    mhc_filename: str = Form(...),
    custom_branch: str = Form(None),
    files: List[UploadFile] = File(...)
):
    # Create a unique session ID
    session_id = str(uuid.uuid4())
    session_dir = os.path.join(UPLOAD_BASE_DIR, session_id)
    os.makedirs(session_dir)

    try:
        # Save or extract files
        for file in files:
            file_path = os.path.join(session_dir, file.filename)
            with open(file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            
            # If it's a zip file, extract it
            if file.filename.endswith(".zip"):
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    zip_ref.extractall(session_dir)
                os.remove(file_path) # Clean up the zip itself

        # Run quantification using the session directory as the root_folder
        output_file = run_quantification(
            network_type=network_type,
            mhc_filename=mhc_filename.lower(),
            root_folder=session_dir,
            mode=mode,
            custom_branches=[custom_branch] if custom_branch else None
        )

        # Move output file to a static location or handle it
        final_output_path = f"static/{output_file}"
        if os.path.exists(output_file):
            shutil.move(output_file, final_output_path)

        return {
            "status": "success", 
            "message": f"Successfully generated {output_file}", 
            "file_url": f"/static/{output_file}"
        }

    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        # We should probably keep the files until the user downloads the result,
        # but for now let's just cleanup logic if needed or use a background task.
        pass

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
