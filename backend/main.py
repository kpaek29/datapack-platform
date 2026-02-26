"""
DataPack Platform - Main API
Secure data pack generation for PE teams
"""
from fastapi import FastAPI, UploadFile, File, Depends, HTTPException, status, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, HTMLResponse
from typing import List, Optional
from pathlib import Path
import shutil
import uuid
from datetime import datetime
import json

from .config import UPLOAD_DIR, OUTPUT_DIR, TEMPLATE_DIR, BASE_DIR, OPENAI_API_KEY, DATA_DIR

# Training library directory (uses persistent disk on Render)
TRAINING_DIR = DATA_DIR / "training_library"
try:
    TRAINING_DIR.mkdir(parents=True, exist_ok=True)
except Exception:
    # Fallback if disk not available
    TRAINING_DIR = BASE_DIR / "training_library"
    TRAINING_DIR.mkdir(parents=True, exist_ok=True)
from .auth import (
    User, Token, authenticate_user, create_access_token, 
    get_current_user, create_user, get_users_db
)
from .processor import DataPackProcessor
from .generators import PPTGenerator, ExcelGenerator
from .ai_analyzer import SmartDataTransformer
from .datapack_generator import generate_datapack
from .sectors import SECTORS, SECTOR_CATEGORIES, get_all_sectors, get_sectors_by_category, validate_sector
from .smart_generator import SmartPPTGenerator, IterativeAnalyzer, QualityValidator, SmartTableFormatter

app = FastAPI(
    title="DataPack Platform",
    description="Secure data pack generation for PE teams",
    version="1.0.0"
)

# CORS - restrict in production
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Restrict to your domain in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Mount static files
FRONTEND_DIR = BASE_DIR / "frontend"
if FRONTEND_DIR.exists():
    app.mount("/static", StaticFiles(directory=FRONTEND_DIR), name="static")


# ============ AUTH ENDPOINTS ============

@app.post("/api/auth/login", response_model=Token)
async def login(username: str = Form(...), password: str = Form(...)):
    """Login and get access token"""
    user = authenticate_user(username, password)
    if not user:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Incorrect username or password"
        )
    access_token = create_access_token(data={"sub": user.username})
    return {"access_token": access_token, "token_type": "bearer"}


# Registration disabled - use CLI to create users:
# python -m backend.cli add-user <username> <password> [--email EMAIL] [--name NAME]


@app.get("/api/auth/me", response_model=User)
async def get_me(current_user: User = Depends(get_current_user)):
    """Get current user info"""
    return current_user


# ============ FILE UPLOAD ENDPOINTS ============

@app.post("/api/upload")
async def upload_files(
    files: List[UploadFile] = File(...),
    current_user: User = Depends(get_current_user)
):
    """Upload Excel files for processing"""
    session_id = str(uuid.uuid4())
    session_dir = UPLOAD_DIR / session_id
    session_dir.mkdir(parents=True, exist_ok=True)
    
    uploaded = []
    for file in files:
        if not file.filename.endswith(('.xlsx', '.xls', '.csv')):
            continue
        
        file_path = session_dir / file.filename
        with open(file_path, 'wb') as f:
            shutil.copyfileobj(file.file, f)
        uploaded.append(file.filename)
    
    # Save session metadata
    metadata = {
        "session_id": session_id,
        "user": current_user.username,
        "timestamp": datetime.now().isoformat(),
        "files": uploaded
    }
    with open(session_dir / "metadata.json", 'w') as f:
        json.dump(metadata, f)
    
    return {
        "session_id": session_id,
        "files_uploaded": uploaded,
        "message": f"Uploaded {len(uploaded)} files"
    }


# ============ PROCESSING ENDPOINTS ============

@app.post("/api/process/{session_id}")
async def process_data(
    session_id: str,
    current_user: User = Depends(get_current_user)
):
    """Process uploaded files and generate analysis"""
    session_dir = UPLOAD_DIR / session_id
    
    if not session_dir.exists():
        raise HTTPException(status_code=404, detail="Session not found")
    
    # Get uploaded files
    files = list(session_dir.glob("*.xlsx")) + list(session_dir.glob("*.xls")) + list(session_dir.glob("*.csv"))
    
    if not files:
        raise HTTPException(status_code=400, detail="No files to process")
    
    # Process
    processor = DataPackProcessor(files)
    summary = processor.generate_summary()
    
    # Save analysis results
    with open(session_dir / "analysis.json", 'w') as f:
        json.dump(summary, f, indent=2, default=str)
    
    return {
        "session_id": session_id,
        "summary": summary
    }


@app.post("/api/generate/{session_id}")
async def generate_outputs(
    session_id: str,
    pack_name: str = Form("Data Pack"),
    current_user: User = Depends(get_current_user)
):
    """Generate PPT and Excel outputs"""
    session_dir = UPLOAD_DIR / session_id
    output_session_dir = OUTPUT_DIR / session_id
    output_session_dir.mkdir(parents=True, exist_ok=True)
    
    if not session_dir.exists():
        raise HTTPException(status_code=404, detail="Session not found")
    
    # Load analysis
    analysis_file = session_dir / "analysis.json"
    if not analysis_file.exists():
        raise HTTPException(status_code=400, detail="Run /process first")
    
    with open(analysis_file) as f:
        analysis = json.load(f)
    
    # Load original dataframes for output
    files = list(session_dir.glob("*.xlsx")) + list(session_dir.glob("*.xls"))
    processor = DataPackProcessor(files)
    processor.load_files()
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Generate PPT
    ppt_path = output_session_dir / f"{pack_name.replace(' ', '_')}_{timestamp}.pptx"
    ppt = PPTGenerator(ppt_path)
    ppt.add_title_slide(pack_name, f"Generated {datetime.now().strftime('%B %d, %Y')}")
    
    # Add slides for each file/sheet
    for filename, sheets in processor.dataframes.items():
        ppt.add_section_slide(filename)
        for sheet_name, df in sheets.items():
            if len(df) > 0:
                ppt.add_table_slide(f"{filename} - {sheet_name}", df)
    
    ppt.save()
    
    # Generate Excel backup
    excel_path = output_session_dir / f"{pack_name.replace(' ', '_')}_{timestamp}_backup.xlsx"
    excel = ExcelGenerator(excel_path)
    excel.add_summary_sheet(analysis.get("analyses", {}))
    
    for filename, sheets in processor.dataframes.items():
        for sheet_name, df in sheets.items():
            safe_name = f"{filename[:10]}_{sheet_name[:15]}"
            excel.add_dataframe_sheet(safe_name, df)
    
    excel.save()
    
    return {
        "session_id": session_id,
        "outputs": {
            "ppt": str(ppt_path.name),
            "excel": str(excel_path.name)
        }
    }


@app.post("/api/generate-smart/{session_id}")
async def generate_smart_outputs(
    session_id: str,
    company_name: str = Form("Company"),
    pack_date: str = Form(None),
    current_user: User = Depends(get_current_user)
):
    """
    Smart generation using AI analysis
    Automatically detects data types and generates appropriate outputs
    """
    session_dir = UPLOAD_DIR / session_id
    output_session_dir = OUTPUT_DIR / session_id
    output_session_dir.mkdir(parents=True, exist_ok=True)
    
    if not session_dir.exists():
        raise HTTPException(status_code=404, detail="Session not found")
    
    # Get uploaded files
    files = list(session_dir.glob("*.xlsx")) + list(session_dir.glob("*.xls")) + list(session_dir.glob("*.csv"))
    
    if not files:
        raise HTTPException(status_code=400, detail="No data files found")
    
    # Use smart transformer
    transformer = SmartDataTransformer(OPENAI_API_KEY)
    
    try:
        financial_data, customer_data, meta = transformer.process_files(files, company_name)
        
        # Use detected company name if not provided
        if company_name == "Company" and meta.get('company_name'):
            company_name = meta['company_name']
            
        # Generate data pack
        date_str = pack_date or datetime.now().strftime("%B %Y")
        
        outputs = generate_datapack(
            company_name=company_name,
            financial_data=financial_data,
            customer_data=customer_data,
            output_dir=output_session_dir,
            date_str=date_str
        )
        
        # Save analysis info
        with open(session_dir / "smart_analysis.json", 'w') as f:
            json.dump({
                'company_name': company_name,
                'analysis': meta.get('analysis', {}),
                'generated_at': datetime.now().isoformat()
            }, f, indent=2, default=str)
        
        return {
            "session_id": session_id,
            "company_name": company_name,
            "ai_analyzed": meta.get('analysis', {}).get('_ai_analyzed', False),
            "outputs": {
                "ppt": outputs['ppt'].name,
                "data_backup": outputs['data_backup'].name,
                "customer_backup": outputs['customer_backup'].name
            }
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Generation failed: {str(e)}")


@app.get("/api/download/{session_id}/{filename}")
async def download_file(
    session_id: str,
    filename: str,
    current_user: User = Depends(get_current_user)
):
    """Download generated output file"""
    file_path = OUTPUT_DIR / session_id / filename
    
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="File not found")
    
    return FileResponse(
        file_path,
        filename=filename,
        media_type="application/octet-stream"
    )


# ============ SESSION MANAGEMENT ============

@app.get("/api/sessions")
async def list_sessions(current_user: User = Depends(get_current_user)):
    """List user's processing sessions"""
    sessions = []
    for session_dir in UPLOAD_DIR.iterdir():
        if session_dir.is_dir():
            meta_file = session_dir / "metadata.json"
            if meta_file.exists():
                with open(meta_file) as f:
                    meta = json.load(f)
                if meta.get("user") == current_user.username:
                    sessions.append(meta)
    
    return {"sessions": sorted(sessions, key=lambda x: x["timestamp"], reverse=True)}


@app.delete("/api/sessions/{session_id}")
async def delete_session(
    session_id: str,
    current_user: User = Depends(get_current_user)
):
    """Delete a processing session and its files"""
    session_dir = UPLOAD_DIR / session_id
    output_dir = OUTPUT_DIR / session_id
    
    if session_dir.exists():
        shutil.rmtree(session_dir)
    if output_dir.exists():
        shutil.rmtree(output_dir)
    
    return {"message": "Session deleted"}


# ============ FRONTEND ============

@app.get("/", response_class=HTMLResponse)
async def root():
    """Serve frontend"""
    index_path = FRONTEND_DIR / "index.html"
    if index_path.exists():
        return HTMLResponse(content=index_path.read_text())
    return HTMLResponse(content="<h1>DataPack Platform</h1><p>Frontend not installed.</p>")


# ============ ITERATIVE ANALYSIS ============

@app.post("/api/analyze-request")
async def analyze_request(
    request: str = Form(...),
    session_id: str = Form(None),
    current_user: User = Depends(get_current_user)
):
    """
    Parse a natural language analysis request
    Example: "Add customer retention analysis" or "Show revenue by segment"
    """
    analyzer = IterativeAnalyzer()
    parsed = analyzer.parse_request(request)
    
    return {
        "request": request,
        "matched_analyses": parsed['matched_analyses'],
        "available_analyses": list(IterativeAnalyzer.ANALYSIS_TYPES.keys()),
        "details": parsed['details']
    }


@app.post("/api/generate-analysis/{session_id}")
async def generate_specific_analysis(
    session_id: str,
    analysis_type: str = Form(...),
    parameters: str = Form("{}"),
    current_user: User = Depends(get_current_user)
):
    """
    Generate a specific analysis and add it to the data pack
    """
    session_dir = UPLOAD_DIR / session_id
    output_session_dir = OUTPUT_DIR / session_id
    
    if not session_dir.exists():
        raise HTTPException(status_code=404, detail="Session not found")
    
    # Load data
    files = list(session_dir.glob("*.xlsx")) + list(session_dir.glob("*.xls")) + list(session_dir.glob("*.csv"))
    
    if not files:
        raise HTTPException(status_code=400, detail="No data files found")
    
    # Process files
    transformer = SmartDataTransformer(OPENAI_API_KEY)
    financial_data, customer_data, meta = transformer.process_files(files)
    
    # Combine data for analysis
    all_data = {**financial_data, **customer_data}
    
    # Parse parameters
    try:
        params = json.loads(parameters)
    except:
        params = {}
    
    # Generate analysis
    analyzer = IterativeAnalyzer()
    result = analyzer.generate_analysis(analysis_type, all_data, params)
    
    # Validate quality
    validator = QualityValidator()
    if isinstance(result.get('data'), pd.DataFrame):
        quality = validator.validate_dataframe(result['data'], analysis_type)
        result['quality'] = quality
    
    return {
        "session_id": session_id,
        "analysis_type": analysis_type,
        "title": result.get('title'),
        "subtitle": result.get('subtitle'),
        "insights": result.get('insights'),
        "has_data": not result.get('data', pd.DataFrame()).empty,
        "data_preview": result.get('data', pd.DataFrame()).head(10).to_dict() if isinstance(result.get('data'), pd.DataFrame) else None,
        "quality": result.get('quality')
    }


@app.post("/api/generate-smart-v2/{session_id}")
async def generate_smart_outputs_v2(
    session_id: str,
    company_name: str = Form("Company"),
    pack_date: str = Form(None),
    additional_analyses: str = Form("[]"),
    current_user: User = Depends(get_current_user)
):
    """
    Smart generation v2 with improved formatting and optional additional analyses
    """
    session_dir = UPLOAD_DIR / session_id
    output_session_dir = OUTPUT_DIR / session_id
    output_session_dir.mkdir(parents=True, exist_ok=True)
    
    if not session_dir.exists():
        raise HTTPException(status_code=404, detail="Session not found")
    
    files = list(session_dir.glob("*.xlsx")) + list(session_dir.glob("*.xls")) + list(session_dir.glob("*.csv"))
    
    if not files:
        raise HTTPException(status_code=400, detail="No data files found")
    
    # Parse additional analyses
    try:
        extra_analyses = json.loads(additional_analyses)
    except:
        extra_analyses = []
    
    # Process files
    transformer = SmartDataTransformer(OPENAI_API_KEY)
    financial_data, customer_data, meta = transformer.process_files(files, company_name)
    
    # Use detected company name
    if company_name == "Company" and meta.get('company_name'):
        company_name = meta['company_name']
    
    date_str = pack_date or datetime.now().strftime("%B %Y")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Generate PPT with smart generator
    ppt_path = output_session_dir / f"{company_name.replace(' ', '_')}_Data_Pack_{timestamp}.pptx"
    
    ppt = SmartPPTGenerator(ppt_path, company_name, date_str)
    
    # Title slide
    ppt.add_title_slide()
    
    # Financial section
    if financial_data.get('consolidated_pl') is not None and not financial_data['consolidated_pl'].empty:
        ppt.add_section_slide("Financial Summary")
        ppt.add_pl_slide(f"P&L Summary – {company_name}", financial_data['consolidated_pl'])
    
    # Customer section
    if customer_data.get('top_customers') is not None and not customer_data['top_customers'].empty:
        ppt.add_section_slide("Customer Analysis")
        ppt.add_customer_slide(f"Top Customers – {company_name}", customer_data['top_customers'])
    
    # Additional requested analyses
    if extra_analyses:
        analyzer = IterativeAnalyzer()
        all_data = {**financial_data, **customer_data}
        
        for analysis_type in extra_analyses:
            result = analyzer.generate_analysis(analysis_type, all_data)
            if result.get('data') is not None and not result['data'].empty:
                ppt.add_table_slide(
                    result['title'],
                    result['data'],
                    subtitle=result.get('subtitle'),
                    footnote=result.get('insights')
                )
    
    ppt.save()
    
    # Generate Excel backup
    excel_path = output_session_dir / f"{company_name.replace(' ', '_')}_Data_Pack_Backup_{timestamp}.xlsx"
    
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        for name, df in financial_data.items():
            if isinstance(df, pd.DataFrame) and not df.empty:
                df.to_excel(writer, sheet_name=name[:31], index=False)
        for name, df in customer_data.items():
            if isinstance(df, pd.DataFrame) and not df.empty:
                df.to_excel(writer, sheet_name=name[:31], index=False)
    
    # Validate quality
    validator = QualityValidator()
    slides_data = [
        {'title': 'P&L', 'data': financial_data.get('consolidated_pl', pd.DataFrame())},
        {'title': 'Customers', 'data': customer_data.get('top_customers', pd.DataFrame())}
    ]
    quality_report = validator.validate_presentation(slides_data)
    
    return {
        "session_id": session_id,
        "company_name": company_name,
        "outputs": {
            "ppt": ppt_path.name,
            "excel": excel_path.name
        },
        "quality": quality_report,
        "analyses_included": ['financial_summary', 'customer_analysis'] + extra_analyses
    }


@app.get("/api/available-analyses")
async def list_available_analyses(current_user: User = Depends(get_current_user)):
    """List all available analysis types"""
    return {
        "analyses": IterativeAnalyzer.ANALYSIS_TYPES
    }


# ============ SECTORS ============

@app.get("/api/sectors")
async def list_sectors():
    """Get all available sectors"""
    return {
        "sectors": get_all_sectors(),
        "count": len(SECTORS)
    }


@app.get("/api/sectors/grouped")
async def list_sectors_grouped():
    """Get sectors grouped by category"""
    return {
        "categories": get_sectors_by_category(),
        "total": len(SECTORS)
    }


# ============ TRAINING LIBRARY ============

@app.post("/api/training/upload")
async def upload_training_files(
    files: List[UploadFile] = File(...),
    sector: str = Form("general"),
    description: str = Form(""),
    current_user: User = Depends(get_current_user)
):
    """Upload example data packs to train the AI"""
    # Validate sector if not "general"
    if sector != "general" and not validate_sector(sector):
        raise HTTPException(
            status_code=400, 
            detail=f"Invalid sector. Use /api/sectors to get valid options."
        )
    
    sector_dir = TRAINING_DIR / sector.lower().replace(" ", "_").replace("/", "_").replace("&", "and")
    sector_dir.mkdir(parents=True, exist_ok=True)
    
    uploaded = []
    for file in files:
        if file.filename and file.filename.endswith(('.xlsx', '.xls', '.pptx', '.ppt', '.csv')):
            file_path = sector_dir / file.filename
            with open(file_path, 'wb') as f:
                shutil.copyfileobj(file.file, f)
            uploaded.append(file.filename)
    
    # Save metadata
    meta_file = sector_dir / "_metadata.json"
    meta = {}
    if meta_file.exists():
        with open(meta_file) as f:
            meta = json.load(f)
    
    meta[datetime.now().isoformat()] = {
        "files": uploaded,
        "description": description,
        "uploaded_by": current_user.username
    }
    
    with open(meta_file, 'w') as f:
        json.dump(meta, f, indent=2)
    
    return {
        "sector": sector,
        "files_uploaded": uploaded,
        "message": f"Uploaded {len(uploaded)} training files"
    }


@app.get("/api/training/list")
async def list_training_files(current_user: User = Depends(get_current_user)):
    """List all training files by sector"""
    result = {}
    
    if TRAINING_DIR.exists():
        for sector_dir in TRAINING_DIR.iterdir():
            if sector_dir.is_dir():
                files = [f.name for f in sector_dir.glob("*") if f.is_file() and not f.name.startswith("_")]
                if files:
                    result[sector_dir.name] = {
                        "files": files,
                        "count": len(files)
                    }
    
    return {"sectors": result, "total_files": sum(s["count"] for s in result.values())}


@app.delete("/api/training/{sector}/{filename}")
async def delete_training_file(
    sector: str,
    filename: str,
    current_user: User = Depends(get_current_user)
):
    """Delete a training file"""
    file_path = TRAINING_DIR / sector / filename
    
    if file_path.exists():
        file_path.unlink()
        return {"message": f"Deleted {filename}"}
    
    raise HTTPException(status_code=404, detail="File not found")


# ============ HEALTH CHECK ============

@app.get("/health")
async def health():
    return {"status": "healthy", "timestamp": datetime.now().isoformat()}


# Initialize first user if none exist
@app.on_event("startup")
async def startup():
    users = get_users_db()
    if not users:
        # Create default admin user - CHANGE PASSWORD IN PRODUCTION
        create_user("admin", "changeme123", "admin@example.com", "Admin User")
        print("Created default admin user (username: admin, password: changeme123)")
