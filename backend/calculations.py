"""
Data Pack Calculations Module
Actual analysis calculations - deterministic, no AI
"""
import pandas as pd
import numpy as np
from typing import Dict, List, Any, Optional, Tuple
from datetime import datetime
from dateutil.relativedelta import relativedelta


class DataPackCalculations:
    """
    Core calculation engine for PE data pack analyses
    All calculations are deterministic - no AI, just math
    """
    
    @staticmethod
    def top_customers(
        df: pd.DataFrame,
        customer_col: str,
        revenue_col: str,
        top_n: int = 20
    ) -> pd.DataFrame:
        """
        Top N customers by revenue
        
        Returns DataFrame with:
        - Rank
        - Customer
        - Revenue
        - % of Total
        - Cumulative %
        """
        # Aggregate by customer
        summary = df.groupby(customer_col)[revenue_col].sum().reset_index()
        summary.columns = ['Customer', 'Revenue']
        
        # Sort and rank
        summary = summary.sort_values('Revenue', ascending=False).head(top_n)
        summary = summary.reset_index(drop=True)
        summary.insert(0, 'Rank', range(1, len(summary) + 1))
        
        # Calculate percentages
        total = summary['Revenue'].sum()
        summary['% of Total'] = summary['Revenue'] / total * 100
        summary['Cumulative %'] = summary['% of Total'].cumsum()
        
        # Format
        summary['Revenue'] = summary['Revenue'].apply(lambda x: f"${x:,.0f}")
        summary['% of Total'] = summary['% of Total'].apply(lambda x: f"{x:.1f}%")
        summary['Cumulative %'] = summary['Cumulative %'].apply(lambda x: f"{x:.1f}%")
        
        return summary
    
    @staticmethod
    def customer_concentration(
        df: pd.DataFrame,
        customer_col: str,
        revenue_col: str
    ) -> pd.DataFrame:
        """
        Customer concentration analysis
        
        Returns DataFrame showing revenue concentration:
        - Top 1, 5, 10, 20 customers
        - Revenue amount
        - % of total
        """
        # Aggregate by customer
        by_customer = df.groupby(customer_col)[revenue_col].sum().sort_values(ascending=False)
        total = by_customer.sum()
        
        results = []
        for n in [1, 5, 10, 20, 50]:
            if len(by_customer) >= n:
                top_n_rev = by_customer.head(n).sum()
                results.append({
                    'Segment': f'Top {n} Customer{"s" if n > 1 else ""}',
                    'Revenue': f"${top_n_rev:,.0f}",
                    '% of Total': f"{top_n_rev/total*100:.1f}%"
                })
        
        # Add total customers count
        results.append({
            'Segment': 'Total Customers',
            'Revenue': f"${total:,.0f}",
            '% of Total': f"{len(by_customer):,}"
        })
        
        return pd.DataFrame(results)
    
    @staticmethod
    def revenue_by_period(
        df: pd.DataFrame,
        date_col: str,
        revenue_col: str,
        period: str = 'M'  # M=monthly, Q=quarterly, Y=yearly
    ) -> pd.DataFrame:
        """
        Revenue aggregated by time period
        
        Returns DataFrame with:
        - Period
        - Revenue
        - Growth %
        """
        df = df.copy()
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        df = df.dropna(subset=[date_col])
        
        # Group by period
        df['Period'] = df[date_col].dt.to_period(period)
        by_period = df.groupby('Period')[revenue_col].sum().reset_index()
        by_period.columns = ['Period', 'Revenue']
        by_period['Period'] = by_period['Period'].astype(str)
        
        # Calculate growth
        by_period['Prior'] = by_period['Revenue'].shift(1)
        by_period['Growth %'] = ((by_period['Revenue'] - by_period['Prior']) / by_period['Prior'] * 100)
        
        # Format
        by_period['Revenue'] = by_period['Revenue'].apply(lambda x: f"${x:,.0f}")
        by_period['Growth %'] = by_period['Growth %'].apply(
            lambda x: f"{x:+.1f}%" if pd.notna(x) else "—"
        )
        
        return by_period[['Period', 'Revenue', 'Growth %']]
    
    @staticmethod
    def customer_retention(
        df: pd.DataFrame,
        customer_col: str,
        date_col: str,
        period: str = 'M'
    ) -> pd.DataFrame:
        """
        Customer retention analysis by period
        
        Returns DataFrame with:
        - Period
        - Active Customers
        - New Customers
        - Retained Customers
        - Churned Customers
        - Retention Rate
        """
        df = df.copy()
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        df = df.dropna(subset=[date_col])
        df['Period'] = df[date_col].dt.to_period(period)
        
        periods = sorted(df['Period'].unique())
        
        results = []
        prev_customers = set()
        all_seen = set()
        
        for period in periods:
            current_customers = set(df[df['Period'] == period][customer_col].unique())
            
            retained = current_customers & prev_customers
            new = current_customers - all_seen
            churned = prev_customers - current_customers
            
            retention_rate = len(retained) / len(prev_customers) * 100 if prev_customers else 0
            
            results.append({
                'Period': str(period),
                'Active': len(current_customers),
                'New': len(new),
                'Retained': len(retained),
                'Churned': len(churned),
                'Retention %': f"{retention_rate:.1f}%" if prev_customers else "—"
            })
            
            prev_customers = current_customers
            all_seen |= current_customers
        
        return pd.DataFrame(results)
    
    @staticmethod
    def cohort_analysis(
        df: pd.DataFrame,
        customer_col: str,
        date_col: str,
        revenue_col: str = None,
        periods: int = 12
    ) -> pd.DataFrame:
        """
        Cohort retention analysis
        
        Groups customers by first purchase month, tracks retention over time
        
        Returns DataFrame (cohort x period matrix)
        """
        df = df.copy()
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        df = df.dropna(subset=[date_col])
        df['Period'] = df[date_col].dt.to_period('M')
        
        # Find first period for each customer
        first_period = df.groupby(customer_col)['Period'].min().reset_index()
        first_period.columns = [customer_col, 'Cohort']
        
        df = df.merge(first_period, on=customer_col)
        
        # Calculate periods since cohort start
        df['Periods Since'] = (df['Period'].astype('datetime64[ns]').dt.to_period('M') - 
                               df['Cohort']).apply(lambda x: x.n if hasattr(x, 'n') else 0)
        
        # Pivot: count unique customers per cohort per period
        cohort_data = df.groupby(['Cohort', 'Periods Since'])[customer_col].nunique().reset_index()
        cohort_data.columns = ['Cohort', 'Period', 'Customers']
        
        # Get cohort sizes
        cohort_sizes = cohort_data[cohort_data['Period'] == 0][['Cohort', 'Customers']]
        cohort_sizes.columns = ['Cohort', 'Cohort Size']
        
        cohort_data = cohort_data.merge(cohort_sizes, on='Cohort')
        cohort_data['Retention %'] = cohort_data['Customers'] / cohort_data['Cohort Size'] * 100
        
        # Pivot to matrix
        matrix = cohort_data.pivot(index='Cohort', columns='Period', values='Retention %')
        matrix = matrix.iloc[:, :min(periods, len(matrix.columns))]
        
        # Format
        matrix = matrix.applymap(lambda x: f"{x:.0f}%" if pd.notna(x) else "")
        matrix.index = matrix.index.astype(str)
        matrix.columns = [f"M{i}" for i in matrix.columns]
        
        return matrix.reset_index()
    
    @staticmethod
    def revenue_by_segment(
        df: pd.DataFrame,
        segment_col: str,
        revenue_col: str
    ) -> pd.DataFrame:
        """
        Revenue breakdown by segment/category
        
        Returns DataFrame with:
        - Segment
        - Revenue
        - % of Total
        """
        by_segment = df.groupby(segment_col)[revenue_col].sum().reset_index()
        by_segment.columns = ['Segment', 'Revenue']
        by_segment = by_segment.sort_values('Revenue', ascending=False)
        
        total = by_segment['Revenue'].sum()
        by_segment['% of Total'] = by_segment['Revenue'] / total * 100
        
        # Format
        by_segment['Revenue'] = by_segment['Revenue'].apply(lambda x: f"${x:,.0f}")
        by_segment['% of Total'] = by_segment['% of Total'].apply(lambda x: f"{x:.1f}%")
        
        return by_segment
    
    @staticmethod
    def yoy_comparison(
        df: pd.DataFrame,
        date_col: str,
        revenue_col: str,
        metric_name: str = "Revenue"
    ) -> pd.DataFrame:
        """
        Year-over-year comparison
        
        Returns DataFrame comparing current vs prior year by month
        """
        df = df.copy()
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        df = df.dropna(subset=[date_col])
        
        df['Year'] = df[date_col].dt.year
        df['Month'] = df[date_col].dt.month
        
        by_month = df.groupby(['Year', 'Month'])[revenue_col].sum().reset_index()
        
        # Get last 2 years
        years = sorted(by_month['Year'].unique())[-2:]
        if len(years) < 2:
            return pd.DataFrame({'Note': ['Insufficient data for YoY comparison']})
        
        prior_year, current_year = years
        
        prior = by_month[by_month['Year'] == prior_year].set_index('Month')[revenue_col]
        current = by_month[by_month['Year'] == current_year].set_index('Month')[revenue_col]
        
        months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                  'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        
        results = []
        for m in range(1, 13):
            p = prior.get(m, 0)
            c = current.get(m, 0)
            growth = ((c - p) / p * 100) if p > 0 else 0
            
            results.append({
                'Month': months[m-1],
                f'{prior_year}': f"${p:,.0f}" if p > 0 else "—",
                f'{current_year}': f"${c:,.0f}" if c > 0 else "—",
                'Growth %': f"{growth:+.1f}%" if p > 0 and c > 0 else "—"
            })
        
        return pd.DataFrame(results)
    
    @staticmethod
    def summary_stats(
        df: pd.DataFrame,
        customer_col: str = None,
        revenue_col: str = None,
        date_col: str = None
    ) -> Dict[str, Any]:
        """
        Quick summary statistics
        """
        stats = {}
        
        if revenue_col and revenue_col in df.columns:
            stats['Total Revenue'] = f"${df[revenue_col].sum():,.0f}"
            stats['Avg Transaction'] = f"${df[revenue_col].mean():,.0f}"
        
        if customer_col and customer_col in df.columns:
            stats['Unique Customers'] = f"{df[customer_col].nunique():,}"
            stats['Total Transactions'] = f"{len(df):,}"
        
        if date_col and date_col in df.columns:
            dates = pd.to_datetime(df[date_col], errors='coerce').dropna()
            if len(dates) > 0:
                stats['Date Range'] = f"{dates.min().strftime('%b %Y')} - {dates.max().strftime('%b %Y')}"
        
        return stats


def detect_columns(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Smart column detection with confidence scores
    """
    detected = {
        'customer': None,
        'revenue': None,
        'date': None,
        'segment': None
    }
    confidence = {
        'customer': 0,
        'revenue': 0,
        'date': 0,
        'segment': 0
    }
    
    for col in df.columns:
        col_lower = str(col).lower().strip()
        sample_values = df[col].dropna().head(10).tolist()
        
        # ===== CUSTOMER DETECTION =====
        customer_score = 0
        # Strong keywords
        if any(kw in col_lower for kw in ['customer', 'client', 'account name', 'company name']):
            customer_score = 95
        # Medium keywords
        elif any(kw in col_lower for kw in ['name', 'account', 'company', 'vendor', 'buyer']):
            customer_score = 70
        # Check if values look like names (strings, varied)
        elif df[col].dtype == 'object':
            unique_ratio = df[col].nunique() / len(df) if len(df) > 0 else 0
            if unique_ratio > 0.1 and unique_ratio < 0.9:  # Not all same, not all unique
                avg_len = df[col].astype(str).str.len().mean()
                if avg_len > 3 and avg_len < 50:
                    customer_score = 40
        
        if customer_score > confidence['customer']:
            detected['customer'] = col
            confidence['customer'] = customer_score
        
        # ===== REVENUE DETECTION =====
        revenue_score = 0
        # Strong keywords
        if any(kw in col_lower for kw in ['revenue', 'sales', 'amount', 'total revenue', 'net sales']):
            revenue_score = 95
        # Medium keywords  
        elif any(kw in col_lower for kw in ['total', 'value', 'amt', 'price', 'sum', 'gross']):
            revenue_score = 70
        # Check if numeric with currency-like values
        elif df[col].dtype in ['int64', 'float64']:
            if df[col].mean() > 100:  # Likely dollar amounts
                revenue_score = 50
            if df[col].min() >= 0:  # Non-negative
                revenue_score += 10
        
        if revenue_score > confidence['revenue']:
            detected['revenue'] = col
            confidence['revenue'] = revenue_score
        
        # ===== DATE DETECTION =====
        date_score = 0
        # Strong keywords
        if any(kw in col_lower for kw in ['date', 'transaction date', 'order date', 'invoice date']):
            date_score = 95
        # Medium keywords
        elif any(kw in col_lower for kw in ['period', 'month', 'year', 'time', 'day', 'created', 'timestamp']):
            date_score = 70
        
        # Check actual values
        if df[col].dtype == 'datetime64[ns]':
            date_score = max(date_score, 90)
        elif df[col].dtype == 'object':
            try:
                parsed = pd.to_datetime(df[col].head(20), errors='coerce')
                valid_ratio = parsed.notna().sum() / min(20, len(df))
                if valid_ratio > 0.7:
                    date_score = max(date_score, 85)
            except:
                pass
        
        # Check for year-like values (2020, 2021, etc.)
        if df[col].dtype in ['int64', 'float64']:
            sample = df[col].dropna().head(20)
            if len(sample) > 0:
                if sample.min() >= 2000 and sample.max() <= 2030:
                    date_score = max(date_score, 60)
        
        if date_score > confidence['date']:
            detected['date'] = col
            confidence['date'] = date_score
        
        # ===== SEGMENT DETECTION =====
        segment_score = 0
        # Strong keywords
        if any(kw in col_lower for kw in ['segment', 'category', 'product type', 'service type']):
            segment_score = 95
        # Medium keywords
        elif any(kw in col_lower for kw in ['type', 'group', 'region', 'division', 'department', 'class']):
            segment_score = 70
        # Check if categorical (few unique values)
        elif df[col].dtype == 'object':
            unique_count = df[col].nunique()
            if 2 <= unique_count <= 20:  # Typical category range
                segment_score = 50
        
        if segment_score > confidence['segment']:
            detected['segment'] = col
            confidence['segment'] = segment_score
    
    return {
        'detected': detected,
        'confidence': confidence
    }
