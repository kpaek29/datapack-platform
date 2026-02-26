"""
DataPack Generator - PE Data Pack Creation Engine
Generates PPT and Excel outputs matching Silver Oak style
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')
from io import BytesIO
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional
import tempfile

# ============ STYLING CONSTANTS ============

class DataPackStyle:
    """Silver Oak PE Data Pack styling"""
    
    # Dimensions (4:3)
    SLIDE_WIDTH = Inches(10)
    SLIDE_HEIGHT = Inches(7.5)
    
    # Colors
    DARK = RGBColor(0x51, 0x51, 0x51)      # #515151 - charcoal
    LIGHT = RGBColor(0xE5, 0xE5, 0xE5)     # #E5E5E5 - light gray
    GREEN = RGBColor(0x3E, 0x77, 0x33)      # #3E7733 - accent green
    NAVY = RGBColor(0x08, 0x46, 0x8D)       # #08468D - accent blue
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    
    # Matplotlib colors
    MPL_DARK = '#515151'
    MPL_GREEN = '#3E7733'
    MPL_NAVY = '#08468D'
    MPL_LIGHT = '#E5E5E5'
    
    # Fonts
    HEADING_FONT = 'Times New Roman'
    BODY_FONT = 'Arial'
    
    # Positions
    TITLE_LEFT = Inches(0.4)
    TITLE_TOP = Inches(0.3)
    TITLE_WIDTH = Inches(9.2)
    
    PAGE_NUM_LEFT = Inches(9.3)
    PAGE_NUM_TOP = Inches(7.2)
    
    CONTENT_LEFT = Inches(0.4)
    CONTENT_TOP = Inches(0.9)
    CONTENT_WIDTH = Inches(9.2)


class DataPackPPTGenerator:
    """Generate PE Data Pack PowerPoint presentations"""
    
    def __init__(self, output_path: str, company_name: str, date_str: str = None):
        self.prs = Presentation()
        self.prs.slide_width = DataPackStyle.SLIDE_WIDTH
        self.prs.slide_height = DataPackStyle.SLIDE_HEIGHT
        self.output_path = Path(output_path)
        self.company_name = company_name
        self.date_str = date_str or datetime.now().strftime("%B %Y")
        self.page_num = 0
        
        # Use blank layout
        self.blank_layout = self.prs.slide_layouts[6]  # Blank
        
    def _add_title(self, slide, text: str, subtitle: str = None):
        """Add title text box to slide"""
        title_box = slide.shapes.add_textbox(
            DataPackStyle.TITLE_LEFT, 
            DataPackStyle.TITLE_TOP,
            DataPackStyle.TITLE_WIDTH,
            Inches(0.5)
        )
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = text
        p.font.name = DataPackStyle.HEADING_FONT
        p.font.size = Pt(24)
        p.font.color.rgb = DataPackStyle.DARK
        p.font.bold = True
        
    def _add_page_number(self, slide):
        """Add page number to bottom right"""
        self.page_num += 1
        num_box = slide.shapes.add_textbox(
            DataPackStyle.PAGE_NUM_LEFT,
            DataPackStyle.PAGE_NUM_TOP,
            Inches(0.5),
            Inches(0.3)
        )
        tf = num_box.text_frame
        p = tf.paragraphs[0]
        p.text = str(self.page_num)
        p.font.name = DataPackStyle.BODY_FONT
        p.font.size = Pt(10)
        p.font.color.rgb = DataPackStyle.DARK
        
    def _add_text_box(self, slide, text: str, left: float, top: float, 
                      width: float = 9.0, height: float = 0.4,
                      font_size: int = 12, bold: bool = False):
        """Add a text box with styling"""
        box = slide.shapes.add_textbox(
            Inches(left), Inches(top), Inches(width), Inches(height)
        )
        tf = box.text_frame
        p = tf.paragraphs[0]
        p.text = text
        p.font.name = DataPackStyle.BODY_FONT
        p.font.size = Pt(font_size)
        p.font.color.rgb = DataPackStyle.DARK
        p.font.bold = bold
        return box
        
    def add_title_slide(self):
        """Add cover/title slide"""
        slide = self.prs.slides.add_slide(self.blank_layout)
        
        # Company name and title
        title_box = slide.shapes.add_textbox(
            Inches(0.8), Inches(5.2), Inches(8.4), Inches(0.6)
        )
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = f"{self.company_name} Data Pack"
        p.font.name = DataPackStyle.HEADING_FONT
        p.font.size = Pt(32)
        p.font.color.rgb = DataPackStyle.DARK
        p.font.bold = True
        
        # Date
        date_box = slide.shapes.add_textbox(
            Inches(0.8), Inches(5.8), Inches(4), Inches(0.4)
        )
        tf = date_box.text_frame
        p = tf.paragraphs[0]
        p.text = self.date_str
        p.font.name = DataPackStyle.BODY_FONT
        p.font.size = Pt(18)
        p.font.color.rgb = DataPackStyle.DARK
        
    def add_agenda_slide(self, sections: List[str], current_section: str = None):
        """Add agenda slide with sections"""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_title(slide, "Agenda")
        self._add_page_number(slide)
        
        # Agenda items
        y_pos = 1.2
        for section in sections:
            is_current = section == current_section
            self._add_text_box(
                slide, section, 0.6, y_pos,
                font_size=14,
                bold=is_current
            )
            y_pos += 0.4
            
    def add_pl_summary_slide(self, title: str, data: pd.DataFrame):
        """Add P&L summary slide with table"""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_title(slide, title)
        self._add_page_number(slide)
        
        # Create table
        rows = min(len(data) + 1, 20)  # Header + data, max 20
        cols = min(len(data.columns), 10)
        
        table = slide.shapes.add_table(
            rows, cols,
            Inches(0.4), Inches(1.0),
            Inches(9.2), Inches(5.5)
        ).table
        
        # Style header row
        for j, col in enumerate(data.columns[:cols]):
            cell = table.cell(0, j)
            cell.text = str(col)
            cell.fill.solid()
            cell.fill.fore_color.rgb = DataPackStyle.NAVY
            p = cell.text_frame.paragraphs[0]
            p.font.color.rgb = DataPackStyle.WHITE
            p.font.size = Pt(9)
            p.font.bold = True
            
        # Fill data
        for i, (idx, row) in enumerate(data.iterrows()):
            if i >= rows - 1:
                break
            for j, val in enumerate(row[:cols]):
                cell = table.cell(i + 1, j)
                cell.text = str(val) if pd.notna(val) else ""
                p = cell.text_frame.paragraphs[0]
                p.font.size = Pt(8)
                p.font.name = DataPackStyle.BODY_FONT
                
    def add_chart_slide(self, title: str, chart_image: bytes, 
                        subtitle: str = None, footnote: str = None):
        """Add slide with chart image"""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_title(slide, title)
        self._add_page_number(slide)
        
        if subtitle:
            self._add_text_box(slide, subtitle, 0.4, 0.9, font_size=14, bold=True)
            
        # Add chart image
        img_stream = BytesIO(chart_image)
        slide.shapes.add_picture(
            img_stream,
            Inches(0.4), Inches(1.4),
            width=Inches(9.0)
        )
        
        if footnote:
            self._add_text_box(slide, footnote, 0.3, 6.5, font_size=8)
            
    def add_dual_chart_slide(self, title: str, 
                             chart1_image: bytes, chart1_title: str,
                             chart2_image: bytes, chart2_title: str):
        """Add slide with two charts stacked"""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_title(slide, title)
        self._add_page_number(slide)
        
        # Top chart
        self._add_text_box(slide, chart1_title, 0.4, 0.9, font_size=12, bold=True)
        img1 = BytesIO(chart1_image)
        slide.shapes.add_picture(img1, Inches(0.4), Inches(1.2), width=Inches(9.0), height=Inches(2.5))
        
        # Bottom chart
        self._add_text_box(slide, chart2_title, 0.4, 3.9, font_size=12, bold=True)
        img2 = BytesIO(chart2_image)
        slide.shapes.add_picture(img2, Inches(0.4), Inches(4.2), width=Inches(9.0), height=Inches(2.5))
        
    def add_top_customers_slide(self, title: str, data: pd.DataFrame):
        """Add top customers table slide"""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_title(slide, title)
        self._add_page_number(slide)
        
        # Same table logic as P&L
        rows = min(len(data) + 1, 25)
        cols = min(len(data.columns), 8)
        
        table = slide.shapes.add_table(
            rows, cols,
            Inches(0.3), Inches(0.9),
            Inches(9.4), Inches(6.0)
        ).table
        
        # Header
        for j, col in enumerate(data.columns[:cols]):
            cell = table.cell(0, j)
            cell.text = str(col)
            cell.fill.solid()
            cell.fill.fore_color.rgb = DataPackStyle.NAVY
            p = cell.text_frame.paragraphs[0]
            p.font.color.rgb = DataPackStyle.WHITE
            p.font.size = Pt(8)
            p.font.bold = True
            
        # Data
        for i, (idx, row) in enumerate(data.iterrows()):
            if i >= rows - 1:
                break
            for j, val in enumerate(row[:cols]):
                cell = table.cell(i + 1, j)
                cell.text = str(val) if pd.notna(val) else ""
                p = cell.text_frame.paragraphs[0]
                p.font.size = Pt(7)
                
    def save(self):
        """Save the presentation"""
        self.prs.save(self.output_path)
        return self.output_path


class ChartGenerator:
    """Generate charts matching Silver Oak style"""
    
    @staticmethod
    def setup_style():
        """Configure matplotlib for Silver Oak style"""
        plt.rcParams['font.family'] = 'sans-serif'
        plt.rcParams['font.sans-serif'] = ['Arial']
        plt.rcParams['axes.facecolor'] = 'white'
        plt.rcParams['figure.facecolor'] = 'white'
        plt.rcParams['axes.edgecolor'] = DataPackStyle.MPL_DARK
        plt.rcParams['axes.labelcolor'] = DataPackStyle.MPL_DARK
        plt.rcParams['xtick.color'] = DataPackStyle.MPL_DARK
        plt.rcParams['ytick.color'] = DataPackStyle.MPL_DARK
        
    @staticmethod
    def monthly_revenue_chart(dates: List, values: List, title: str = "") -> bytes:
        """Generate monthly revenue bar chart"""
        ChartGenerator.setup_style()
        
        fig, ax = plt.subplots(figsize=(9, 2.5), dpi=150)
        
        bars = ax.bar(range(len(values)), values, color=DataPackStyle.MPL_NAVY)
        
        ax.set_xticks(range(len(dates)))
        ax.set_xticklabels(dates, rotation=45, ha='right', fontsize=7)
        ax.set_ylabel('Revenue ($000s)', fontsize=9)
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.0f}'))
        
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        
        plt.tight_layout()
        
        buf = BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight', dpi=150)
        plt.close()
        buf.seek(0)
        return buf.read()
        
    @staticmethod
    def ttm_revenue_chart(dates: List, values: List, title: str = "") -> bytes:
        """Generate TTM revenue line chart"""
        ChartGenerator.setup_style()
        
        fig, ax = plt.subplots(figsize=(9, 2.5), dpi=150)
        
        ax.plot(range(len(values)), values, color=DataPackStyle.MPL_GREEN, 
                linewidth=2, marker='o', markersize=4)
        ax.fill_between(range(len(values)), values, alpha=0.2, color=DataPackStyle.MPL_GREEN)
        
        ax.set_xticks(range(len(dates)))
        ax.set_xticklabels(dates, rotation=45, ha='right', fontsize=7)
        ax.set_ylabel('TTM Revenue ($M)', fontsize=9)
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.1f}'))
        
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        
        plt.tight_layout()
        
        buf = BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight', dpi=150)
        plt.close()
        buf.seek(0)
        return buf.read()
        
    @staticmethod  
    def segment_breakdown_chart(segments: Dict[str, float]) -> bytes:
        """Generate segment breakdown pie/bar chart"""
        ChartGenerator.setup_style()
        
        fig, ax = plt.subplots(figsize=(4.5, 3), dpi=150)
        
        colors = [DataPackStyle.MPL_NAVY, DataPackStyle.MPL_GREEN, 
                  DataPackStyle.MPL_DARK, DataPackStyle.MPL_LIGHT]
        
        bars = ax.barh(list(segments.keys()), list(segments.values()), 
                       color=colors[:len(segments)])
        
        ax.set_xlabel('Revenue ($000s)', fontsize=9)
        ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.0f}'))
        
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        
        plt.tight_layout()
        
        buf = BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight', dpi=150)
        plt.close()
        buf.seek(0)
        return buf.read()


class DataPackExcelGenerator:
    """Generate Data Pack backup Excel files"""
    
    def __init__(self, output_path: str):
        self.output_path = Path(output_path)
        self.writer = None
        self.sheets = {}
        
    def __enter__(self):
        self.writer = pd.ExcelWriter(self.output_path, engine='openpyxl')
        return self
        
    def __exit__(self, *args):
        if self.writer:
            self.writer.close()
            
    def add_sheet(self, name: str, df: pd.DataFrame, index: bool = True):
        """Add a sheet with data"""
        df.to_excel(self.writer, sheet_name=name[:31], index=index)
        
    def add_pl_sheet(self, name: str, data: pd.DataFrame):
        """Add formatted P&L sheet"""
        self.add_sheet(name, data)
        
    def add_customer_analysis(self, name: str, data: pd.DataFrame):
        """Add customer analysis sheet"""
        self.add_sheet(name, data, index=False)


def generate_datapack(
    company_name: str,
    financial_data: Dict[str, pd.DataFrame],
    customer_data: Dict[str, pd.DataFrame],
    output_dir: Path,
    date_str: str = None
) -> Dict[str, Path]:
    """
    Main function to generate a complete data pack
    
    Args:
        company_name: Name of the company
        financial_data: Dict of DataFrames with financial sheets
        customer_data: Dict of DataFrames with customer analysis
        output_dir: Directory for output files
        date_str: Date string for the pack (e.g., "October 2025")
        
    Returns:
        Dict with paths to generated files
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Generate PPT
    ppt_path = output_dir / f"{company_name.replace(' ', '_')}_Data_Pack_{timestamp}.pptx"
    ppt = DataPackPPTGenerator(ppt_path, company_name, date_str)
    
    # Title slide
    ppt.add_title_slide()
    
    # Agenda
    sections = [
        "Financial Trends",
        "By Segment", 
        "EBITDA Schedule",
        "Customer Trends"
    ]
    ppt.add_agenda_slide(sections, "Financial Trends")
    
    # P&L Summary slides
    if 'consolidated_pl' in financial_data:
        ppt.add_pl_summary_slide(f"Summary P&L – {company_name}", financial_data['consolidated_pl'])
        
    # Revenue charts
    if 'monthly_revenue' in financial_data:
        rev_data = financial_data['monthly_revenue']
        
        monthly_chart = ChartGenerator.monthly_revenue_chart(
            rev_data['date'].tolist(),
            rev_data['revenue'].tolist()
        )
        
        ttm_chart = ChartGenerator.ttm_revenue_chart(
            rev_data['date'].tolist(),
            rev_data['ttm_revenue'].tolist() if 'ttm_revenue' in rev_data else rev_data['revenue'].tolist()
        )
        
        ppt.add_dual_chart_slide(
            f"Monthly and Rolling TTM Revenue – {company_name}",
            monthly_chart, "Monthly Revenue",
            ttm_chart, "Rolling TTM Revenue"
        )
        
    # Segment analysis
    ppt.add_agenda_slide(sections, "By Segment")
    
    # Customer analysis
    ppt.add_agenda_slide(sections, "Customer Trends")
    
    if 'top_customers' in customer_data:
        ppt.add_top_customers_slide(
            f"Top Customers – {company_name}",
            customer_data['top_customers']
        )
    
    ppt.save()
    
    # Generate Excel backup
    excel_path = output_dir / f"{company_name.replace(' ', '_')}_Data_Pack_Backup_{timestamp}.xlsx"
    
    with DataPackExcelGenerator(excel_path) as excel:
        for sheet_name, df in financial_data.items():
            excel.add_sheet(sheet_name, df)
            
    # Generate Customer backup
    customer_excel_path = output_dir / f"{company_name.replace(' ', '_')}_Customer_Backup_{timestamp}.xlsx"
    
    with DataPackExcelGenerator(customer_excel_path) as excel:
        for sheet_name, df in customer_data.items():
            excel.add_sheet(sheet_name, df, index=False)
    
    return {
        'ppt': ppt_path,
        'data_backup': excel_path,
        'customer_backup': customer_excel_path
    }
