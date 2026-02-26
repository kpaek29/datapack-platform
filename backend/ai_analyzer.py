"""
AI-Powered Data Analyzer
Uses GPT to intelligently analyze and map raw data files
"""
import pandas as pd
import json
import os
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
from openai import OpenAI
from .config import OPENAI_API_KEY


class AIDataAnalyzer:
    """
    Intelligently analyzes raw Excel/CSV data using GPT
    Identifies data types, columns, and optimal mappings
    """
    
    def __init__(self, api_key: str = None):
        self.api_key = api_key or OPENAI_API_KEY
        if self.api_key:
            self.client = OpenAI(api_key=self.api_key)
        else:
            self.client = None
            
    def analyze_file(self, filepath: Path) -> Dict[str, Any]:
        """
        Analyze an Excel/CSV file and return intelligent mapping
        
        Returns:
            {
                'file_type': 'financial' | 'customer' | 'mixed',
                'company_name': detected company name,
                'sheets': {
                    'sheet_name': {
                        'data_type': 'pl' | 'revenue' | 'customers' | 'fleet' | 'qoe' | 'other',
                        'columns': {...column mappings...},
                        'time_periods': [...detected periods...],
                        'metrics': [...detected metrics...]
                    }
                },
                'summary': 'human readable summary'
            }
        """
        # Load file
        if str(filepath).endswith('.csv'):
            sheets = {'Sheet1': pd.read_csv(filepath, nrows=100)}
        else:
            xlsx = pd.ExcelFile(filepath)
            sheets = {}
            for name in xlsx.sheet_names[:15]:  # Limit to 15 sheets
                try:
                    df = pd.read_excel(xlsx, sheet_name=name, nrows=50)
                    if not df.empty:
                        sheets[name] = df
                except:
                    pass
        
        if not self.client:
            # Fallback to rule-based analysis if no API key
            return self._analyze_without_ai(sheets)
        
        # Build analysis prompt
        analysis = self._ai_analyze_sheets(sheets)
        return analysis
    
    def _get_sheet_preview(self, df: pd.DataFrame) -> str:
        """Get a text preview of a dataframe for GPT"""
        preview = []
        
        # Column names
        cols = [str(c) for c in df.columns[:15]]
        preview.append(f"Columns: {cols}")
        
        # First few rows as text
        for i, row in df.head(8).iterrows():
            row_vals = [str(v)[:30] for v in row.values[:15]]
            preview.append(f"Row {i}: {row_vals}")
            
        # Data types
        dtypes = df.dtypes.head(10).to_dict()
        preview.append(f"Types: {dtypes}")
        
        return '\n'.join(preview)
    
    def _ai_analyze_sheets(self, sheets: Dict[str, pd.DataFrame]) -> Dict[str, Any]:
        """Use GPT to analyze sheets"""
        
        # Build prompt with sheet previews
        sheet_previews = []
        for name, df in list(sheets.items())[:10]:
            preview = self._get_sheet_preview(df)
            sheet_previews.append(f"=== Sheet: {name} ===\n{preview}")
        
        prompt = f"""Analyze these Excel sheets from a business data file. Identify:
1. What type of data each sheet contains (P&L, revenue, customers, fleet, QoE, transactions, etc.)
2. Key column mappings (which columns are revenue, customer names, dates, amounts, etc.)
3. Time periods covered
4. Company name if visible

Sheets:
{chr(10).join(sheet_previews)}

Respond in JSON format:
{{
    "company_name": "detected name or null",
    "file_type": "financial|customer|mixed|unknown",
    "sheets": {{
        "sheet_name": {{
            "data_type": "pl|revenue|customers|fleet|qoe|transactions|summary|other",
            "description": "brief description",
            "key_columns": {{
                "date": "column name or null",
                "amount": "column name or null", 
                "customer": "column name or null",
                "category": "column name or null",
                "period": "column name or null"
            }},
            "time_periods": ["2023", "2024", etc],
            "has_monthly_data": true/false,
            "usefulness": "high|medium|low"
        }}
    }},
    "recommended_outputs": ["ppt_pl", "ppt_revenue_chart", "customer_analysis", etc],
    "summary": "Brief summary of the data"
}}"""

        try:
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "You are a data analyst expert at understanding business Excel files. Always respond with valid JSON."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.1,
                max_tokens=2000
            )
            
            result_text = response.choices[0].message.content
            
            # Parse JSON from response
            # Handle markdown code blocks
            if '```json' in result_text:
                result_text = result_text.split('```json')[1].split('```')[0]
            elif '```' in result_text:
                result_text = result_text.split('```')[1].split('```')[0]
                
            result = json.loads(result_text)
            result['_ai_analyzed'] = True
            return result
            
        except Exception as e:
            print(f"AI analysis failed: {e}")
            return self._analyze_without_ai(sheets)
    
    def _analyze_without_ai(self, sheets: Dict[str, pd.DataFrame]) -> Dict[str, Any]:
        """Rule-based fallback analysis"""
        result = {
            'company_name': None,
            'file_type': 'unknown',
            'sheets': {},
            'recommended_outputs': [],
            'summary': 'Rule-based analysis (no AI)',
            '_ai_analyzed': False
        }
        
        financial_keywords = ['revenue', 'ebitda', 'income', 'expense', 'profit', 'p&l', 'qoe']
        customer_keywords = ['customer', 'client', 'account', 'invoice', 'order']
        fleet_keywords = ['fleet', 'vehicle', 'truck', 'equipment', 'asset']
        
        for name, df in sheets.items():
            sheet_info = {
                'data_type': 'other',
                'description': '',
                'key_columns': {},
                'time_periods': [],
                'has_monthly_data': False,
                'usefulness': 'low'
            }
            
            # Check column names and content
            all_text = ' '.join([str(c).lower() for c in df.columns])
            all_text += ' ' + ' '.join([str(v).lower() for v in df.values.flatten()[:500] if pd.notna(v)])
            
            # Classify sheet
            if any(kw in all_text for kw in financial_keywords):
                sheet_info['data_type'] = 'financial'
                sheet_info['usefulness'] = 'high'
                result['file_type'] = 'financial'
                
            if any(kw in all_text for kw in customer_keywords):
                sheet_info['data_type'] = 'customers'
                sheet_info['usefulness'] = 'high'
                if result['file_type'] == 'unknown':
                    result['file_type'] = 'customer'
                elif result['file_type'] == 'financial':
                    result['file_type'] = 'mixed'
                    
            if any(kw in all_text for kw in fleet_keywords):
                sheet_info['data_type'] = 'fleet'
                sheet_info['usefulness'] = 'medium'
            
            # Look for dates
            for col in df.columns:
                col_str = str(col).lower()
                if any(year in col_str for year in ['2022', '2023', '2024', '2025']):
                    sheet_info['time_periods'].append(col_str)
                    sheet_info['has_monthly_data'] = True
                    
            # Find key columns
            for col in df.columns:
                col_lower = str(col).lower()
                if 'date' in col_lower:
                    sheet_info['key_columns']['date'] = str(col)
                if any(x in col_lower for x in ['amount', 'revenue', 'total', 'sum']):
                    sheet_info['key_columns']['amount'] = str(col)
                if any(x in col_lower for x in ['customer', 'client', 'name', 'account']):
                    sheet_info['key_columns']['customer'] = str(col)
                    
            result['sheets'][name] = sheet_info
            
        return result
    
    def get_extraction_instructions(self, analysis: Dict[str, Any]) -> Dict[str, Any]:
        """
        Based on analysis, return instructions for data extraction
        """
        instructions = {
            'pl_sheets': [],
            'revenue_sheets': [],
            'customer_sheets': [],
            'fleet_sheets': [],
            'date_column': None,
            'amount_column': None,
            'customer_column': None
        }
        
        for sheet_name, info in analysis.get('sheets', {}).items():
            dtype = info.get('data_type', 'other')
            
            if dtype in ['pl', 'financial', 'qoe']:
                instructions['pl_sheets'].append(sheet_name)
            if dtype == 'revenue' or (dtype == 'financial' and info.get('has_monthly_data')):
                instructions['revenue_sheets'].append(sheet_name)
            if dtype == 'customers':
                instructions['customer_sheets'].append(sheet_name)
            if dtype == 'fleet':
                instructions['fleet_sheets'].append(sheet_name)
                
            # Get column mappings
            cols = info.get('key_columns', {})
            if cols.get('date') and not instructions['date_column']:
                instructions['date_column'] = cols['date']
            if cols.get('amount') and not instructions['amount_column']:
                instructions['amount_column'] = cols['amount']
            if cols.get('customer') and not instructions['customer_column']:
                instructions['customer_column'] = cols['customer']
                
        return instructions


class SmartDataTransformer:
    """
    Combines AI analysis with data extraction
    """
    
    def __init__(self, api_key: str = None):
        self.analyzer = AIDataAnalyzer(api_key)
        
    def process_files(
        self, 
        files: List[Path],
        company_name: str = None
    ) -> Tuple[Dict[str, pd.DataFrame], Dict[str, pd.DataFrame], Dict[str, Any]]:
        """
        Process uploaded files intelligently
        
        Returns:
            (financial_data, customer_data, analysis_info)
        """
        all_analyses = []
        all_sheets = {}
        
        # Analyze each file
        for filepath in files:
            analysis = self.analyzer.analyze_file(filepath)
            all_analyses.append(analysis)
            
            # Load sheets
            if str(filepath).endswith('.csv'):
                all_sheets[f"{filepath.stem}_data"] = pd.read_csv(filepath)
            else:
                xlsx = pd.ExcelFile(filepath)
                for name in xlsx.sheet_names:
                    try:
                        df = pd.read_excel(xlsx, sheet_name=name)
                        all_sheets[f"{filepath.stem}_{name}"] = df
                    except:
                        pass
        
        # Merge analyses
        merged = self._merge_analyses(all_analyses)
        
        # Get extraction instructions
        instructions = self.analyzer.get_extraction_instructions(merged)
        
        # Extract data based on instructions
        financial_data = self._extract_financial(all_sheets, instructions, merged)
        customer_data = self._extract_customers(all_sheets, instructions, merged)
        
        # Use detected or provided company name
        if not company_name and merged.get('company_name'):
            company_name = merged['company_name']
            
        return financial_data, customer_data, {
            'analysis': merged,
            'instructions': instructions,
            'company_name': company_name or 'Company'
        }
    
    def _merge_analyses(self, analyses: List[Dict]) -> Dict:
        """Merge multiple file analyses"""
        if len(analyses) == 1:
            return analyses[0]
            
        merged = {
            'company_name': None,
            'file_type': 'mixed',
            'sheets': {},
            'summary': 'Multiple files analyzed'
        }
        
        for a in analyses:
            if a.get('company_name'):
                merged['company_name'] = a['company_name']
            merged['sheets'].update(a.get('sheets', {}))
            
        return merged
    
    def _extract_financial(
        self, 
        sheets: Dict[str, pd.DataFrame],
        instructions: Dict,
        analysis: Dict
    ) -> Dict[str, pd.DataFrame]:
        """Extract financial data based on AI instructions"""
        result = {
            'consolidated_pl': pd.DataFrame(),
            'monthly_revenue': pd.DataFrame(),
            'qoe': pd.DataFrame()
        }
        
        # Find best P&L sheet
        for sheet_name in instructions.get('pl_sheets', []):
            for full_name, df in sheets.items():
                if sheet_name in full_name:
                    # Extract P&L structure
                    pl_data = self._extract_pl_from_sheet(df)
                    if not pl_data.empty:
                        result['consolidated_pl'] = pl_data
                        break
            if not result['consolidated_pl'].empty:
                break
                
        # Find revenue data
        for sheet_name in instructions.get('revenue_sheets', []):
            for full_name, df in sheets.items():
                if sheet_name in full_name:
                    revenue_data = self._extract_revenue_from_sheet(df)
                    if not revenue_data.empty:
                        result['monthly_revenue'] = revenue_data
                        break
            if not result['monthly_revenue'].empty:
                break
                
        return result
    
    def _extract_customers(
        self,
        sheets: Dict[str, pd.DataFrame],
        instructions: Dict,
        analysis: Dict
    ) -> Dict[str, pd.DataFrame]:
        """Extract customer data based on AI instructions"""
        result = {
            'top_customers': pd.DataFrame(),
            'customer_analysis': pd.DataFrame()
        }
        
        for sheet_name in instructions.get('customer_sheets', []):
            for full_name, df in sheets.items():
                if sheet_name in full_name:
                    # Try to find customer and amount columns
                    cust_col = instructions.get('customer_column')
                    amt_col = instructions.get('amount_column')
                    
                    if cust_col and cust_col in df.columns:
                        # Aggregate by customer
                        if amt_col and amt_col in df.columns:
                            top = df.groupby(cust_col)[amt_col].sum().sort_values(ascending=False).head(20)
                            result['top_customers'] = top.reset_index()
                            result['top_customers'].columns = ['Customer', 'Total']
                        else:
                            # Just get customer list
                            result['top_customers'] = df[[cust_col]].drop_duplicates().head(20)
                            result['top_customers'].columns = ['Customer']
                    break
                    
        return result
    
    def _extract_pl_from_sheet(self, df: pd.DataFrame) -> pd.DataFrame:
        """Extract P&L data from a sheet"""
        # Look for revenue row
        result = []
        
        for i, row in df.iterrows():
            row_text = ' '.join([str(x).lower() for x in row.values if pd.notna(x)])
            
            if any(kw in row_text for kw in ['revenue', 'sales', 'income', 'ebitda', 'gross profit', 'operating']):
                # This looks like a P&L line item
                label = None
                values = []
                
                for val in row.values:
                    if pd.notna(val):
                        if isinstance(val, str) and len(val) > 2 and not val.replace('.','').replace('-','').isdigit():
                            if not label:
                                label = val[:50]
                        elif isinstance(val, (int, float)):
                            values.append(val)
                            
                if label and values:
                    result.append({'Line Item': label, **{f'Col{i}': v for i, v in enumerate(values[:5])}})
                    
        return pd.DataFrame(result) if result else pd.DataFrame()
    
    def _extract_revenue_from_sheet(self, df: pd.DataFrame) -> pd.DataFrame:
        """Extract monthly revenue data"""
        result = []
        
        # Look for date columns
        date_cols = []
        for col in df.columns:
            col_str = str(col)
            if any(year in col_str for year in ['2022', '2023', '2024', '2025']):
                date_cols.append(col)
            elif hasattr(col, 'strftime'):  # datetime
                date_cols.append(col)
                
        if not date_cols:
            return pd.DataFrame()
            
        # Find revenue row
        for i, row in df.iterrows():
            row_text = ' '.join([str(x).lower() for x in row.values if pd.notna(x)])
            if 'revenue' in row_text and 'growth' not in row_text:
                for col in date_cols[:24]:
                    val = row.get(col)
                    if pd.notna(val) and isinstance(val, (int, float)):
                        date_str = col.strftime('%b %Y') if hasattr(col, 'strftime') else str(col)[:10]
                        result.append({'date': date_str, 'revenue': val})
                break
                
        return pd.DataFrame(result) if result else pd.DataFrame()
