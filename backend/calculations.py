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


def detect_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    """
    Simple heuristic column detection - no AI
    """
    detected = {
        'customer': None,
        'revenue': None,
        'date': None,
        'segment': None
    }
    
    for col in df.columns:
        col_lower = str(col).lower()
        
        # Customer detection
        if any(kw in col_lower for kw in ['customer', 'client', 'account', 'name', 'company']):
            if detected['customer'] is None:
                detected['customer'] = col
        
        # Revenue detection
        if any(kw in col_lower for kw in ['revenue', 'amount', 'sales', 'total', 'value', 'amt']):
            if detected['revenue'] is None:
                detected['revenue'] = col
        
        # Date detection
        if any(kw in col_lower for kw in ['date', 'period', 'month', 'time']):
            if detected['date'] is None:
                detected['date'] = col
        
        # Segment detection  
        if any(kw in col_lower for kw in ['segment', 'category', 'type', 'group', 'region']):
            if detected['segment'] is None:
                detected['segment'] = col
    
    # Also check data types
    for col in df.columns:
        if detected['date'] is None:
            if df[col].dtype == 'datetime64[ns]':
                detected['date'] = col
            elif df[col].dtype == 'object':
                # Try parsing as date
                try:
                    parsed = pd.to_datetime(df[col].head(10), errors='coerce')
                    if parsed.notna().sum() >= 5:
                        detected['date'] = col
                except:
                    pass
        
        if detected['revenue'] is None:
            if df[col].dtype in ['int64', 'float64']:
                # Large numbers likely revenue
                if df[col].mean() > 100:
                    detected['revenue'] = col
    
    return detected
