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
import pandas as pd

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
from .smart_generator import SmartPPTGenerator, IterativeAnalyzer, QualityValidator, SmartTableFormatter, AnalysisSuggester
from .calculations import DataPackCalculations, detect_columns
from .excel_builder import DataPackExcelBuilder

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


# ============ ANALYSIS SUGGESTIONS ============

@app.post("/api/suggest-analyses/{session_id}")
async def suggest_analyses(
    session_id: str,
    current_user: User = Depends(get_current_user)
):
    """
    AI-powered analysis suggestions based on uploaded data
    Call after upload to get recommended analyses
    """
    session_dir = UPLOAD_DIR / session_id
    
    if not session_dir.exists():
        raise HTTPException(status_code=404, detail="Session not found")
    
    # Load uploaded files
    files = list(session_dir.glob("*.xlsx")) + list(session_dir.glob("*.xls")) + list(session_dir.glob("*.csv"))
    
    if not files:
        raise HTTPException(status_code=400, detail="No data files found")
    
    # Load dataframes
    all_dfs = {}
    for filepath in files:
        try:
            if str(filepath).endswith('.csv'):
                df = pd.read_csv(filepath, nrows=100)
                all_dfs[filepath.stem] = df
            else:
                xlsx = pd.ExcelFile(filepath)
                for sheet in xlsx.sheet_names[:10]:
                    try:
                        df = pd.read_excel(xlsx, sheet_name=sheet, nrows=100)
                        all_dfs[f"{filepath.stem}_{sheet}"] = df
                    except:
                        pass
        except Exception as e:
            continue
    
    if not all_dfs:
        return {
            "session_id": session_id,
            "suggested": [],
            "reasons": {},
            "additional": list(IterativeAnalyzer.ANALYSIS_TYPES.keys()),
            "message": "Could not parse uploaded files"
        }
    
    # Get suggestions
    suggester = AnalysisSuggester()
    suggestions = suggester.analyze_dataframes(all_dfs)
    
    return {
        "session_id": session_id,
        "suggested": suggestions['suggested'],
        "reasons": suggestions['reasons'],
        "confidence": suggestions['confidence'],
        "additional": suggestions['additional'],
        "data_summary": {
            "files": len(files),
            "sheets_analyzed": len(all_dfs)
        }
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


# ============ COLUMN MAPPING & GUIDED GENERATION ============

@app.post("/api/detect-columns/{session_id}")
async def detect_columns_endpoint(
    session_id: str,
    current_user: User = Depends(get_current_user)
):
    """
    Detect columns in uploaded data and return all columns for mapping
    """
    session_dir = UPLOAD_DIR / session_id
    
    if not session_dir.exists():
        raise HTTPException(status_code=404, detail="Session not found")
    
    files = list(session_dir.glob("*.xlsx")) + list(session_dir.glob("*.xls")) + list(session_dir.glob("*.csv"))
    
    if not files:
        raise HTTPException(status_code=400, detail="No data files found")
    
    # Load first file
    filepath = files[0]
    try:
        if str(filepath).endswith('.csv'):
            df = pd.read_csv(filepath)
        else:
            xlsx = pd.ExcelFile(filepath)
            df = pd.read_excel(xlsx, sheet_name=xlsx.sheet_names[0])
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Could not read file: {str(e)}")
    
    # Detect columns with confidence
    detection_result = detect_columns(df)
    detected = detection_result['detected']
    confidence = detection_result['confidence']
    
    # Get all columns with sample data
    columns_info = []
    for col in df.columns:
        sample = df[col].dropna().head(3).tolist()
        sample_str = [str(s)[:30] for s in sample]
        columns_info.append({
            'name': str(col),
            'dtype': str(df[col].dtype),
            'sample': sample_str,
            'non_null': int(df[col].notna().sum()),
            'total': len(df)
        })
    
    return {
        "session_id": session_id,
        "file": filepath.name,
        "rows": len(df),
        "columns": columns_info,
        "detected": detected,
        "confidence": confidence,
        "available_analyses": [
            {"id": "top_customers", "name": "Top Customers", "requires": ["customer", "revenue"]},
            {"id": "concentration", "name": "Customer Concentration", "requires": ["customer", "revenue"]},
            {"id": "retention", "name": "Customer Retention", "requires": ["customer", "date"]},
            {"id": "revenue_by_period", "name": "Revenue by Period", "requires": ["date", "revenue"]},
            {"id": "revenue_by_segment", "name": "Revenue by Segment", "requires": ["segment", "revenue"]},
            {"id": "yoy_comparison", "name": "Year-over-Year", "requires": ["date", "revenue"]},
            {"id": "cohort", "name": "Cohort Analysis", "requires": ["customer", "date"]}
        ]
    }


@app.post("/api/generate-with-mapping/{session_id}")
async def generate_with_mapping(
    session_id: str,
    company_name: str = Form("Company"),
    customer_col: str = Form(None),
    revenue_col: str = Form(None),
    date_col: str = Form(None),
    segment_col: str = Form(None),
    analyses: str = Form("[]"),  # JSON array of analysis IDs
    current_user: User = Depends(get_current_user)
):
    """
    Generate data pack with explicit column mappings
    """
    session_dir = UPLOAD_DIR / session_id
    output_dir = OUTPUT_DIR / session_id
    output_dir.mkdir(parents=True, exist_ok=True)
    
    if not session_dir.exists():
        raise HTTPException(status_code=404, detail="Session not found")
    
    # Parse analyses
    try:
        selected_analyses = json.loads(analyses)
    except:
        selected_analyses = []
    
    # Load data
    files = list(session_dir.glob("*.xlsx")) + list(session_dir.glob("*.xls")) + list(session_dir.glob("*.csv"))
    if not files:
        raise HTTPException(status_code=400, detail="No data files found")
    
    filepath = files[0]
    try:
        if str(filepath).endswith('.csv'):
            df = pd.read_csv(filepath)
        else:
            xlsx = pd.ExcelFile(filepath)
            df = pd.read_excel(xlsx, sheet_name=xlsx.sheet_names[0])
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Could not read file: {str(e)}")
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    calc = DataPackCalculations()
    
    # Run selected analyses
    results = {}
    excel_sheets = {}
    
    if 'top_customers' in selected_analyses and customer_col and revenue_col:
        try:
            results['top_customers'] = calc.top_customers(df, customer_col, revenue_col)
            excel_sheets['Top Customers'] = results['top_customers']
        except Exception as e:
            results['top_customers_error'] = str(e)
    
    if 'concentration' in selected_analyses and customer_col and revenue_col:
        try:
            results['concentration'] = calc.customer_concentration(df, customer_col, revenue_col)
            excel_sheets['Concentration'] = results['concentration']
        except Exception as e:
            results['concentration_error'] = str(e)
    
    if 'retention' in selected_analyses and customer_col and date_col:
        try:
            results['retention'] = calc.customer_retention(df, customer_col, date_col)
            excel_sheets['Retention'] = results['retention']
        except Exception as e:
            results['retention_error'] = str(e)
    
    if 'revenue_by_period' in selected_analyses and date_col and revenue_col:
        try:
            results['revenue_by_period'] = calc.revenue_by_period(df, date_col, revenue_col)
            excel_sheets['Revenue by Period'] = results['revenue_by_period']
        except Exception as e:
            results['revenue_by_period_error'] = str(e)
    
    if 'revenue_by_segment' in selected_analyses and segment_col and revenue_col:
        try:
            results['revenue_by_segment'] = calc.revenue_by_segment(df, segment_col, revenue_col)
            excel_sheets['Revenue by Segment'] = results['revenue_by_segment']
        except Exception as e:
            results['revenue_by_segment_error'] = str(e)
    
    if 'yoy_comparison' in selected_analyses and date_col and revenue_col:
        try:
            results['yoy_comparison'] = calc.yoy_comparison(df, date_col, revenue_col)
            excel_sheets['YoY Comparison'] = results['yoy_comparison']
        except Exception as e:
            results['yoy_comparison_error'] = str(e)
    
    if 'cohort' in selected_analyses and customer_col and date_col:
        try:
            results['cohort'] = calc.cohort_analysis(df, customer_col, date_col)
            excel_sheets['Cohort Analysis'] = results['cohort']
        except Exception as e:
            results['cohort_error'] = str(e)
    
    # Generate PPT
    ppt_path = output_dir / f"{company_name.replace(' ', '_')}_Data_Pack_{timestamp}.pptx"
    ppt = SmartPPTGenerator(ppt_path, company_name, datetime.now().strftime("%B %Y"))
    
    ppt.add_title_slide()
    
    # Add analysis slides
    analysis_titles = {
        'top_customers': f'Top Customers – {company_name}',
        'concentration': f'Customer Concentration – {company_name}',
        'retention': f'Customer Retention – {company_name}',
        'revenue_by_period': f'Revenue by Period – {company_name}',
        'revenue_by_segment': f'Revenue by Segment – {company_name}',
        'yoy_comparison': f'Year-over-Year Comparison – {company_name}',
        'cohort': f'Cohort Analysis – {company_name}'
    }
    
    for analysis_id, title in analysis_titles.items():
        if analysis_id in results and isinstance(results[analysis_id], pd.DataFrame):
            if not results[analysis_id].empty:
                ppt.add_table_slide(title, results[analysis_id])
    
    ppt.save()
    
    # Generate Excel with formulas
    excel_path = output_dir / f"{company_name.replace(' ', '_')}_Data_Pack_{timestamp}.xlsx"
    excel_builder = DataPackExcelBuilder(excel_path)
    
    # Add raw data first (this is the source for formulas)
    excel_builder.add_raw_data(df, "Raw Data")
    
    # Add output tabs with formulas where possible
    analyses_added = []
    
    if 'top_customers' in selected_analyses and customer_col and revenue_col:
        try:
            excel_builder.add_top_customers_with_formulas(df, customer_col, revenue_col)
            analyses_added.append("Top Customers")
        except:
            if 'top_customers' in results:
                excel_builder.add_static_output(results['top_customers'], "Top Customers", "Top Customers")
                analyses_added.append("Top Customers")
    
    if 'concentration' in selected_analyses and customer_col and revenue_col:
        try:
            excel_builder.add_concentration_with_formulas(df, customer_col, revenue_col)
            analyses_added.append("Concentration")
        except:
            if 'concentration' in results:
                excel_builder.add_static_output(results['concentration'], "Concentration", "Customer Concentration")
                analyses_added.append("Concentration")
    
    if 'revenue_by_period' in selected_analyses and date_col and revenue_col:
        try:
            excel_builder.add_revenue_by_period_with_formulas(df, date_col, revenue_col)
            analyses_added.append("Revenue by Period")
        except:
            if 'revenue_by_period' in results:
                excel_builder.add_static_output(results['revenue_by_period'], "Revenue by Period", "Revenue by Period")
                analyses_added.append("Revenue by Period")
    
    # Add remaining analyses as static outputs
    for analysis_id in ['retention', 'revenue_by_segment', 'yoy_comparison', 'cohort']:
        if analysis_id in results and isinstance(results[analysis_id], pd.DataFrame):
            if not results[analysis_id].empty:
                title = analysis_id.replace('_', ' ').title()
                excel_builder.add_static_output(results[analysis_id], title, title)
                analyses_added.append(title)
    
    # Add index sheet
    excel_builder.add_index_sheet(analyses_added)
    
    excel_builder.save()
    
    return {
        "session_id": session_id,
        "company_name": company_name,
        "analyses_run": [k for k in results.keys() if not k.endswith('_error')],
        "errors": {k: v for k, v in results.items() if k.endswith('_error')},
        "outputs": {
            "ppt": ppt_path.name,
            "excel": excel_path.name
        }
    }


@app.post("/api/chat-refine")
async def chat_refine(
    message: str = Form(...),
    session_id: str = Form(...),
    current_config: str = Form("{}"),
    current_user: User = Depends(get_current_user)
):
    """
    Chat-based refinement of data pack
    Interprets natural language requests and applies changes
    """
    try:
        config = json.loads(current_config)
    except:
        config = {}
    
    message_lower = message.lower()
    
    # Parse the request
    response = ""
    action = "none"
    new_params = {}
    
    # Top N changes
    import re
    top_n_match = re.search(r'top\s*(\d+)', message_lower)
    if top_n_match:
        n = int(top_n_match.group(1))
        new_params['top_n'] = n
        response = f"I'll show the top {n} customers instead. Regenerating..."
        action = "regenerate"
    
    # Period changes
    if 'quarterly' in message_lower:
        new_params['period'] = 'Q'
        response = "Switching to quarterly grouping. Regenerating..."
        action = "regenerate"
    elif 'monthly' in message_lower:
        new_params['period'] = 'M'
        response = "Switching to monthly grouping. Regenerating..."
        action = "regenerate"
    elif 'yearly' in message_lower or 'annual' in message_lower:
        new_params['period'] = 'Y'
        response = "Switching to yearly grouping. Regenerating..."
        action = "regenerate"
    
    # Add analysis
    if 'add' in message_lower:
        if 'retention' in message_lower:
            new_params['add_analysis'] = 'retention'
            response = "Adding customer retention analysis. Regenerating..."
            action = "regenerate"
        elif 'cohort' in message_lower:
            new_params['add_analysis'] = 'cohort'
            response = "Adding cohort analysis. Regenerating..."
            action = "regenerate"
        elif 'concentration' in message_lower:
            new_params['add_analysis'] = 'concentration'
            response = "Adding concentration analysis. Regenerating..."
            action = "regenerate"
    
    # If no specific action recognized
    if action == "none":
        response = "I understood your request but I'm not sure how to apply it yet. Try:\n• 'Show top 30 customers'\n• 'Switch to quarterly'\n• 'Add retention analysis'"
        return {"response": response, "action": "none"}
    
    # Regenerate with new params if action required
    if action == "regenerate":
        session_dir = UPLOAD_DIR / session_id
        output_dir = OUTPUT_DIR / session_id
        
        if not session_dir.exists():
            return {"response": "Session not found. Please upload again.", "action": "error"}
        
        # Load data
        files = list(session_dir.glob("*.xlsx")) + list(session_dir.glob("*.xls")) + list(session_dir.glob("*.csv"))
        if not files:
            return {"response": "No data file found.", "action": "error"}
        
        filepath = files[0]
        try:
            if str(filepath).endswith('.csv'):
                df = pd.read_csv(filepath)
            else:
                xlsx = pd.ExcelFile(filepath)
                df = pd.read_excel(xlsx, sheet_name=xlsx.sheet_names[0])
        except:
            return {"response": "Could not read data file.", "action": "error"}
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        calc = DataPackCalculations()
        company_name = config.get('company_name', 'Company')
        
        # Get column mappings
        customer_col = config.get('customer_col')
        revenue_col = config.get('revenue_col')
        date_col = config.get('date_col')
        
        results = {}
        excel_sheets = {}
        
        # Apply top_n if specified
        top_n = new_params.get('top_n', 20)
        
        if customer_col and revenue_col:
            results['top_customers'] = calc.top_customers(df, customer_col, revenue_col, top_n=top_n)
            excel_sheets['Top Customers'] = results['top_customers']
            
            results['concentration'] = calc.customer_concentration(df, customer_col, revenue_col)
            excel_sheets['Concentration'] = results['concentration']
        
        if customer_col and date_col:
            period = new_params.get('period', 'M')
            results['retention'] = calc.customer_retention(df, customer_col, date_col, period=period)
            excel_sheets['Retention'] = results['retention']
        
        if date_col and revenue_col:
            period = new_params.get('period', 'M')
            results['revenue_by_period'] = calc.revenue_by_period(df, date_col, revenue_col, period=period)
            excel_sheets['Revenue by Period'] = results['revenue_by_period']
        
        excel_sheets['Raw Data'] = df
        
        # Generate PPT
        ppt_path = output_dir / f"{company_name.replace(' ', '_')}_Data_Pack_{timestamp}.pptx"
        ppt = SmartPPTGenerator(ppt_path, company_name, datetime.now().strftime("%B %Y"))
        ppt.add_title_slide()
        
        for key, result_df in results.items():
            if isinstance(result_df, pd.DataFrame) and not result_df.empty:
                title = key.replace('_', ' ').title() + f" – {company_name}"
                ppt.add_table_slide(title, result_df)
        
        ppt.save()
        
        # Generate Excel
        excel_path = output_dir / f"{company_name.replace(' ', '_')}_Data_Pack_{timestamp}.xlsx"
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            for sheet_name, sheet_df in excel_sheets.items():
                if isinstance(sheet_df, pd.DataFrame) and not sheet_df.empty:
                    sheet_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
        
        return {
            "response": response,
            "action": "regenerate",
            "new_outputs": {
                "ppt": ppt_path.name,
                "excel": excel_path.name
            }
        }
    
    return {"response": response, "action": action}


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
