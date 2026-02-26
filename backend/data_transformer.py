"""
Data Transformer - Extract and transform raw data for DataPack generation
Handles Silver Oak PE data pack format
"""
import pandas as pd
import numpy as np
from pathlib import Path
from typing import Dict, List, Tuple, Optional
from datetime import datetime
import re


class DataPackTransformer:
    """
    Transform raw Excel backups into structured data for DataPack generation
    """
    
    def __init__(self):
        self.financial_data = {}
        self.customer_data = {}
        self.company_name = "Company"
        self.segments = []
        
    def load_financial_backup(self, filepath: Path) -> Dict[str, pd.DataFrame]:
        """Load and parse the financial backup Excel file"""
        xlsx = pd.ExcelFile(filepath)
        
        result = {
            'raw_sheets': {},
            'pl_consolidated': None,
            'pl_by_segment': {},
            'monthly_revenue': None,
            'qoe': None,
            'fleet': None
        }
        
        for sheet in xlsx.sheet_names:
            try:
                df = pd.read_excel(xlsx, sheet_name=sheet, header=None)
                result['raw_sheets'][sheet] = df
            except:
                pass
                
        # Extract structured data
        result['pl_consolidated'] = self._extract_consolidated_pl(result['raw_sheets'])
        result['monthly_revenue'] = self._extract_monthly_revenue(result['raw_sheets'])
        result['qoe'] = self._extract_qoe(result['raw_sheets'])
        result['pl_by_segment'] = self._extract_segment_pl(result['raw_sheets'])
        result['fleet'] = self._extract_fleet(result['raw_sheets'])
        
        return result
        
    def load_customer_backup(self, filepath: Path) -> Dict[str, pd.DataFrame]:
        """Load and parse the customer backup Excel file"""
        xlsx = pd.ExcelFile(filepath)
        
        result = {
            'raw_sheets': {},
            'top_customers': None,
            'customer_analysis': None,
            'service_analysis': None,
            'by_segment': {}
        }
        
        for sheet in xlsx.sheet_names:
            try:
                df = pd.read_excel(xlsx, sheet_name=sheet, header=None)
                result['raw_sheets'][sheet] = df
            except:
                pass
                
        # Extract structured data
        result['top_customers'] = self._extract_top_customers(result['raw_sheets'])
        result['customer_analysis'] = self._extract_customer_analysis(result['raw_sheets'])
        result['service_analysis'] = self._extract_service_analysis(result['raw_sheets'])
        
        return result
    
    def _find_header_row(self, df: pd.DataFrame, keywords: List[str]) -> int:
        """Find the row that contains header keywords"""
        for i, row in df.iterrows():
            row_str = ' '.join([str(x).lower() for x in row.values if pd.notna(x)])
            if any(kw.lower() in row_str for kw in keywords):
                return i
        return 0
        
    def _extract_consolidated_pl(self, sheets: Dict) -> pd.DataFrame:
        """Extract consolidated P&L data"""
        # Look for QofE or Consol sheet
        for name in ['QofE', 'Consol', 'Consolidating_IS_Unadj']:
            if name in sheets:
                df = sheets[name]
                
                # Find header row with years
                header_row = self._find_header_row(df, ['2023', '2024', 'TTM', 'Revenue'])
                
                if header_row > 0:
                    # Extract P&L structure
                    pl_data = []
                    for i in range(header_row + 1, min(header_row + 30, len(df))):
                        row = df.iloc[i]
                        # Look for line item in first few columns
                        line_item = None
                        for j in range(5):
                            if pd.notna(row.iloc[j]) and isinstance(row.iloc[j], str):
                                if len(row.iloc[j].strip()) > 2:
                                    line_item = row.iloc[j].strip()
                                    break
                        
                        if line_item and line_item not in ['NaN', '']:
                            # Extract numeric values
                            values = []
                            for val in row.values[3:8]:
                                if pd.notna(val) and isinstance(val, (int, float)):
                                    values.append(val)
                            
                            if values:
                                pl_data.append({
                                    'Line Item': line_item,
                                    'Values': values[:5]  # Take first 5 numeric values
                                })
                    
                    if pl_data:
                        # Create DataFrame
                        result_df = pd.DataFrame(pl_data)
                        # Expand values to columns
                        if len(result_df) > 0:
                            max_vals = max(len(r['Values']) for r in pl_data)
                            cols = ['2022', '2023', '2024', 'TTM', 'YTD'][:max_vals]
                            for i, col in enumerate(cols):
                                result_df[col] = result_df['Values'].apply(
                                    lambda x: x[i] if len(x) > i else None
                                )
                            result_df = result_df.drop('Values', axis=1)
                            return result_df
        
        return pd.DataFrame()
        
    def _extract_monthly_revenue(self, sheets: Dict) -> pd.DataFrame:
        """Extract monthly revenue data for charts"""
        # Look for QofE or IS sheets
        for name in ['QofE', 'IS_BO', 'IS_UT']:
            if name in sheets:
                df = sheets[name]
                
                # Find row with Revenue
                for i, row in df.iterrows():
                    row_vals = [str(x) for x in row.values if pd.notna(x)]
                    if 'Revenue' in row_vals or 'Revenue - unadjusted' in row_vals:
                        # This row has revenue data
                        # Look for date columns
                        header_row = None
                        for j in range(max(0, i-5), i):
                            check_row = df.iloc[j]
                            for val in check_row.values:
                                if isinstance(val, datetime) or (isinstance(val, str) and '2023' in val):
                                    header_row = j
                                    break
                            if header_row:
                                break
                        
                        if header_row is not None:
                            dates = []
                            revenues = []
                            headers = df.iloc[header_row]
                            values = df.iloc[i]
                            
                            for k, (h, v) in enumerate(zip(headers, values)):
                                if isinstance(h, datetime):
                                    dates.append(h.strftime('%b %Y'))
                                    if pd.notna(v) and isinstance(v, (int, float)):
                                        revenues.append(v)
                                    else:
                                        revenues.append(0)
                            
                            if dates and revenues:
                                return pd.DataFrame({
                                    'date': dates[:24],  # Last 24 months
                                    'revenue': revenues[:24]
                                })
        
        return pd.DataFrame()
    
    def _extract_qoe(self, sheets: Dict) -> pd.DataFrame:
        """Extract Quality of Earnings data"""
        if 'QofE' in sheets:
            df = sheets['QofE']
            
            # Find the main data section
            header_row = self._find_header_row(df, ['Revenue', 'EBITDA', 'Net income'])
            
            if header_row > 0:
                # Extract key metrics
                metrics = []
                for i in range(header_row, min(header_row + 50, len(df))):
                    row = df.iloc[i]
                    label = None
                    for j in range(3):
                        if pd.notna(row.iloc[j]) and isinstance(row.iloc[j], str):
                            label = row.iloc[j].strip()
                            break
                    
                    if label and any(kw in label.lower() for kw in 
                                    ['revenue', 'ebitda', 'net income', 'gross profit']):
                        values = [v for v in row.values if isinstance(v, (int, float)) and pd.notna(v)]
                        if values:
                            metrics.append({
                                'Metric': label,
                                'Value': values[0] if values else 0
                            })
                
                return pd.DataFrame(metrics)
        
        return pd.DataFrame()
    
    def _extract_segment_pl(self, sheets: Dict) -> Dict[str, pd.DataFrame]:
        """Extract P&L by segment"""
        segments = {}
        
        for name in ['IS_BO', 'IS_UT', 'BOTR', 'UT']:
            if name in sheets:
                df = sheets[name]
                
                # Determine segment name
                segment_name = 'Brian Omps' if 'BO' in name else 'Ultimate Towing'
                
                # Find revenue row
                for i, row in df.iterrows():
                    row_str = ' '.join([str(x) for x in row.values if pd.notna(x)])
                    if 'Revenue' in row_str:
                        # Extract data
                        pl_items = []
                        for j in range(max(0, i-2), min(i+20, len(df))):
                            check_row = df.iloc[j]
                            label = None
                            for k in range(5):
                                if pd.notna(check_row.iloc[k]) and isinstance(check_row.iloc[k], str):
                                    if len(check_row.iloc[k].strip()) > 2:
                                        label = check_row.iloc[k].strip()
                                        break
                            
                            if label:
                                values = [v for v in check_row.values if isinstance(v, (int, float)) and pd.notna(v)]
                                if values:
                                    pl_items.append({'Item': label, 'Value': values[0]})
                        
                        if pl_items:
                            segments[segment_name] = pd.DataFrame(pl_items)
                        break
        
        return segments
    
    def _extract_fleet(self, sheets: Dict) -> pd.DataFrame:
        """Extract fleet summary data"""
        for name in ['Consol', 'BOTR', 'Fleet Overview', 'Summary']:
            if name in sheets:
                df = sheets[name]
                
                # Look for fleet-related rows
                if 'Fleet' in str(df.iloc[:5].values):
                    # This might be a fleet sheet
                    header_row = self._find_header_row(df, ['Type', 'Year', 'Make', 'Model', 'Branch'])
                    
                    if header_row >= 0:
                        # Extract fleet data
                        fleet_df = df.iloc[header_row+1:].copy()
                        fleet_df.columns = df.iloc[header_row].values
                        fleet_df = fleet_df.dropna(how='all')
                        
                        # Clean up
                        if len(fleet_df) > 0:
                            return fleet_df.head(50)  # First 50 vehicles
        
        return pd.DataFrame()
    
    def _extract_top_customers(self, sheets: Dict) -> pd.DataFrame:
        """Extract top customers data"""
        for name in ['Brian Omps Top Customers', 'UT Top Customers', 'Top Customers', 'Customer Analysis']:
            if name in sheets:
                df = sheets[name]
                
                # Find header row
                header_row = self._find_header_row(df, ['Customer', 'Revenue', '2023', '2024'])
                
                if header_row >= 0:
                    # Extract customer data
                    result = []
                    for i in range(header_row + 1, min(header_row + 30, len(df))):
                        row = df.iloc[i]
                        
                        # Skip empty rows
                        if row.isna().all():
                            continue
                            
                        customer = None
                        for j in range(3):
                            if pd.notna(row.iloc[j]) and isinstance(row.iloc[j], str):
                                customer = row.iloc[j].strip()
                                break
                        
                        if customer and len(customer) > 2:
                            values = [v for v in row.values if isinstance(v, (int, float)) and pd.notna(v)]
                            if values:
                                result.append({
                                    'Customer': customer[:30],  # Truncate long names
                                    '2023': values[0] if len(values) > 0 else 0,
                                    '2024': values[1] if len(values) > 1 else 0,
                                    'TTM': values[2] if len(values) > 2 else values[0] if values else 0
                                })
                    
                    if result:
                        result_df = pd.DataFrame(result)
                        # Calculate % of total
                        total = result_df['TTM'].sum()
                        if total > 0:
                            result_df['% of Total'] = (result_df['TTM'] / total * 100).round(1).astype(str) + '%'
                        return result_df.head(20)  # Top 20
        
        return pd.DataFrame()
    
    def _extract_customer_analysis(self, sheets: Dict) -> pd.DataFrame:
        """Extract detailed customer analysis"""
        if 'Customer Analysis' in sheets:
            df = sheets['Customer Analysis']
            
            # Find header row
            header_row = self._find_header_row(df, ['Customer', 'Revenue', 'Invoices', 'Jobs'])
            
            if header_row >= 0:
                result_df = df.iloc[header_row+1:].copy()
                result_df.columns = df.iloc[header_row].values
                result_df = result_df.dropna(how='all')
                return result_df.head(1000)  # First 1000 rows
        
        return pd.DataFrame()
    
    def _extract_service_analysis(self, sheets: Dict) -> pd.DataFrame:
        """Extract service analysis data"""
        for name in ['Service Analysis', 'Service Analysis (BOTR)', 'Service Analysis (UT)']:
            if name in sheets:
                df = sheets[name]
                
                header_row = self._find_header_row(df, ['Service', 'Revenue', 'Jobs', 'Type'])
                
                if header_row >= 0:
                    result_df = df.iloc[header_row+1:].copy()
                    result_df.columns = df.iloc[header_row].values
                    result_df = result_df.dropna(how='all')
                    return result_df.head(200)
        
        return pd.DataFrame()


def transform_for_datapack(
    financial_file: Path,
    customer_file: Path = None,
    company_name: str = "Company"
) -> Tuple[Dict[str, pd.DataFrame], Dict[str, pd.DataFrame]]:
    """
    Main transformation function
    
    Args:
        financial_file: Path to financial backup Excel
        customer_file: Path to customer backup Excel (optional)
        company_name: Name of the company
        
    Returns:
        Tuple of (financial_data, customer_data) dicts ready for generator
    """
    transformer = DataPackTransformer()
    
    # Load financial data
    fin_result = transformer.load_financial_backup(financial_file)
    
    financial_data = {
        'consolidated_pl': fin_result.get('pl_consolidated', pd.DataFrame()),
        'monthly_revenue': fin_result.get('monthly_revenue', pd.DataFrame()),
        'qoe': fin_result.get('qoe', pd.DataFrame()),
        'fleet': fin_result.get('fleet', pd.DataFrame()),
    }
    
    # Add segment P&Ls
    for seg_name, seg_df in fin_result.get('pl_by_segment', {}).items():
        financial_data[f'pl_{seg_name.replace(" ", "_").lower()}'] = seg_df
    
    # Load customer data if provided
    customer_data = {}
    if customer_file and customer_file.exists():
        cust_result = transformer.load_customer_backup(customer_file)
        customer_data = {
            'top_customers': cust_result.get('top_customers', pd.DataFrame()),
            'customer_analysis': cust_result.get('customer_analysis', pd.DataFrame()),
            'service_analysis': cust_result.get('service_analysis', pd.DataFrame()),
        }
    
    return financial_data, customer_data
