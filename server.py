from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel
import os
import sys

# Import our backend logic
from quantify import run_quantification

app = FastAPI()

# Serve static files (HTML, CSS, JS)
app.mount("/static", StaticFiles(directory="static"), name="static")

class QuantifyRequest(BaseModel):
    network_type: str
    mode: str
    mhc_filename: str
    root_folder: str
    custom_branch: str = None

@app.get("/")
def read_root():
    return FileResponse("static/index.html")

@app.post("/api/quantify")
def quantify_endpoint(req: QuantifyRequest):
    try:
        custom_branches = [req.custom_branch] if req.custom_branch else None
        # Call our decoupled backend function
        output_file = run_quantification(
            network_type=req.network_type,
            mhc_filename=req.mhc_filename.lower(),
            root_folder=req.root_folder,
            mode=req.mode,
            custom_branches=custom_branches
        )
        return {"status": "success", "message": f"Successfully generated {output_file}", "file": output_file}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    # Make sure we run on 8000
    uvicorn.run(app, host="127.0.0.1", port=8000)
