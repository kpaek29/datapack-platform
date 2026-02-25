"""
Data Processing Module
Handles Excel file parsing and analysis
"""
import pandas as pd
from pathlib import Path
from typing import Dict, List, Any, Optional
import json

class DataPackProcessor:
    """Process uploaded Excel files and extract insights"""
    
    def __init__(self, files: List[Path]):
        self.files = files
        self.dataframes: Dict[str, pd.DataFrame] = {}
        self.analyses: Dict[str, Any] = {}
    
    def load_files(self) -> Dict[str, List[str]]:
        """Load all Excel files and return sheet names"""
        file_info = {}
        for file_path in self.files:
            try:
                xl = pd.ExcelFile(file_path)
                self.dataframes[file_path.name] = {}
                for sheet in xl.sheet_names:
                    df = pd.read_excel(xl, sheet_name=sheet)
                    self.dataframes[file_path.name][sheet] = df
                file_info[file_path.name] = xl.sheet_names
            except Exception as e:
                file_info[file_path.name] = [f"Error: {str(e)}"]
        return file_info
    
    def detect_data_types(self) -> Dict[str, Any]:
        """Detect what type of data is in each file/sheet"""
        detected = {}
        
        financial_keywords = ['revenue', 'ebitda', 'margin', 'income', 'expense', 'profit', 'loss', 'sales']
        customer_keywords = ['customer', 'client', 'account', 'retention', 'churn', 'cohort']
        
        for filename, sheets in self.dataframes.items():
            detected[filename] = {}
            for sheet_name, df in sheets.items():
                columns_lower = [str(c).lower() for c in df.columns]
                
                data_type = "unknown"
                if any(kw in ' '.join(columns_lower) for kw in financial_keywords):
                    data_type = "financial"
                elif any(kw in ' '.join(columns_lower) for kw in customer_keywords):
                    data_type = "customer"
                
                detected[filename][sheet_name] = {
                    "type": data_type,
                    "columns": list(df.columns),
                    "rows": len(df),
                    "numeric_columns": list(df.select_dtypes(include=['number']).columns),
                    "date_columns": list(df.select_dtypes(include=['datetime']).columns)
                }
        
        return detected
    
    def analyze_financials(self, df: pd.DataFrame, date_col: str = None, value_cols: List[str] = None) -> Dict:
        """Generate financial analysis from a dataframe"""
        analysis = {}
        
        if value_cols is None:
            value_cols = list(df.select_dtypes(include=['number']).columns)
        
        for col in value_cols:
            if col in df.columns:
                analysis[col] = {
                    "total": float(df[col].sum()),
                    "mean": float(df[col].mean()),
                    "min": float(df[col].min()),
                    "max": float(df[col].max()),
                    "growth": None
                }
                
                # Calculate growth if we have a date column
                if date_col and date_col in df.columns:
                    df_sorted = df.sort_values(date_col)
                    if len(df_sorted) > 1:
                        first_val = df_sorted[col].iloc[0]
                        last_val = df_sorted[col].iloc[-1]
                        if first_val != 0:
                            analysis[col]["growth"] = float((last_val - first_val) / first_val * 100)
        
        return analysis
    
    def analyze_customers(self, df: pd.DataFrame) -> Dict:
        """Generate customer analysis from a dataframe"""
        analysis = {
            "total_customers": len(df),
            "segments": {},
            "retention": None
        }
        
        # Try to find segment/category columns
        for col in df.columns:
            if df[col].dtype == 'object' and df[col].nunique() < 20:
                analysis["segments"][col] = df[col].value_counts().to_dict()
        
        return analysis
    
    def generate_summary(self) -> Dict[str, Any]:
        """Generate overall data pack summary"""
        self.load_files()
        data_types = self.detect_data_types()
        
        summary = {
            "files_processed": len(self.files),
            "data_types": data_types,
            "analyses": {}
        }
        
        for filename, sheets in self.dataframes.items():
            summary["analyses"][filename] = {}
            for sheet_name, df in sheets.items():
                sheet_type = data_types.get(filename, {}).get(sheet_name, {}).get("type", "unknown")
                
                if sheet_type == "financial":
                    summary["analyses"][filename][sheet_name] = self.analyze_financials(df)
                elif sheet_type == "customer":
                    summary["analyses"][filename][sheet_name] = self.analyze_customers(df)
                else:
                    # Basic stats for unknown types
                    summary["analyses"][filename][sheet_name] = {
                        "rows": len(df),
                        "columns": list(df.columns)
                    }
        
        return summary
