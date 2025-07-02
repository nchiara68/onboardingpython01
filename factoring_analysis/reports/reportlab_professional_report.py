"""
Professional PDF Report Generator using ReportLab
Advanced business document with charts, tables, and professional styling
File: reports/reportlab_professional_report.py
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime
import os
import sys
from pathlib import Path
from io import BytesIO
import base64

# Add project root to path for imports
project_root = Path(__file__).parent.parent
sys.path.append(str(project_root))

# Install ReportLab if not available
try:
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle, Paragraph, 
                                   Spacer, Image, PageBreak, KeepTogether)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch, cm
    from reportlab.lib.colors import Color, HexColor, black, white, red, green, blue
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
    from reportlab.graphics.shapes import Drawing, Rect, Circle, String
    from reportlab.graphics.charts.piecharts import Pie
    from reportlab.graphics.charts.barcharts import VerticalBarChart, HorizontalBarChart
    from reportlab.graphics.charts.linecharts import HorizontalLineChart
    from reportlab.graphics.widgets.markers import makeMarker
    from reportlab.graphics import renderPDF
    from reportlab.lib.utils import ImageReader
    print("✅ ReportLab available")
except ImportError:
    print("📦 Installing ReportLab...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "reportlab"])
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle, Paragraph, 
                                   Spacer, Image, PageBreak, KeepTogether)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch, cm
    from reportlab.lib.colors import Color, HexColor, black, white, red, green, blue
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
    from reportlab.graphics.shapes import Drawing, Rect, Circle, String
    from reportlab.graphics.charts.piecharts import Pie
    from reportlab.graphics.charts.barcharts import VerticalBarChart, HorizontalBarChart
    from reportlab.graphics.charts.linecharts import HorizontalLineChart
    from reportlab.graphics.widgets.markers import makeMarker
    from reportlab.graphics import renderPDF
    from reportlab.lib.utils import ImageReader

import warnings
warnings.filterwarnings('ignore')

class ReportLabProfessionalReport:
    def __init__(self, analyzer, results):
        """
        Initialize the ReportLab professional report generator
        
        Parameters:
        analyzer: CorrectedFactoringAnalyzer instance
        results: Results dictionary from run_full_analysis()
        """
        self.analyzer = analyzer
        self.results = results
        
        # Set up paths
        self.project_root = Path(__file__).parent.parent
        self.output_dir = self.project_root / 'output'
        self.charts_dir = self.output_dir / 'charts'
        self.reports_dir = self.output_dir / 'reports'
        
        # Create directories
        self.charts_dir.mkdir(parents=True, exist_ok=True)
        self.reports_dir.mkdir(parents=True, exist_ok=True)
        
        print(f"📁 Charts directory: {self.charts_dir}")
        print(f"📁 Reports directory: {self.reports_dir}")
        
        # Define corporate color scheme
        self.colors = {
            'primary': HexColor('#2E8B57'),      # Sea Green
            'secondary': HexColor('#4682B4'),    # Steel Blue
            'accent': HexColor('#FF6B35'),       # Orange Red
            'success': HexColor('#28a745'),      # Success Green
            'warning': HexColor('#ffc107'),      # Warning Yellow
            'danger': HexColor('#dc3545'),       # Danger Red
            'light_gray': HexColor('#f8f9fa'),   # Light Gray
            'dark_gray': HexColor('#6c757d'),    # Dark Gray
            'white': white,
            'black': black
        }
        
        # Set up custom styles
        self.styles = self.setup_styles()
    
    def setup_styles(self):
        """Create custom paragraph styles"""
        styles = getSampleStyleSheet()
        
        # Corporate Title Style
        styles.add(ParagraphStyle(
            name='CorporateTitle',
            parent=styles['Title'],
            fontSize=28,
            textColor=self.colors['primary'],
            spaceAfter=30,
            alignment=TA_CENTER,
            fontName='Helvetica-Bold'
        ))
        
        # Subtitle Style
        styles.add(ParagraphStyle(
            name='Subtitle',
            parent=styles['Normal'],
            fontSize=16,
            textColor=self.colors['secondary'],
            spaceAfter=20,
            alignment=TA_CENTER,
            fontName='Helvetica-Bold'
        ))
        
        # Section Header Style
        styles.add(ParagraphStyle(
            name='SectionHeader',
            parent=styles['Heading1'],
            fontSize=18,
            textColor=self.colors['primary'],
            spaceAfter=15,
            spaceBefore=25,
            fontName='Helvetica-Bold'
        ))
        
        # Subsection Header Style
        styles.add(ParagraphStyle(
            name='SubsectionHeader',
            parent=styles['Heading2'],
            fontSize=14,
            textColor=self.colors['secondary'],
            spaceAfter=10,
            spaceBefore=15,
            fontName='Helvetica-Bold'
        ))
        
        # Metric Style
        styles.add(ParagraphStyle(
            name='MetricValue',
            parent=styles['Normal'],
            fontSize=24,
            textColor=self.colors['primary'],
            alignment=TA_CENTER,
            fontName='Helvetica-Bold'
        ))
        
        # Metric Label Style
        styles.add(ParagraphStyle(
            name='MetricLabel',
            parent=styles['Normal'],
            fontSize=10,
            textColor=self.colors['dark_gray'],
            alignment=TA_CENTER,
            fontName='Helvetica'
        ))
        
        # Body text with corporate styling
        styles.add(ParagraphStyle(
            name='CorporateBody',
            parent=styles['Normal'],
            fontSize=11,
            textColor=black,
            spaceAfter=12,
            alignment=TA_JUSTIFY,
            fontName='Helvetica'
        ))
        
        # Risk indicator styles
        styles.add(ParagraphStyle(
            name='RiskHigh',
            parent=styles['Normal'],
            fontSize=16,
            textColor=white,
            alignment=TA_CENTER,
            fontName='Helvetica-Bold',
            backColor=self.colors['danger']
        ))
        
        styles.add(ParagraphStyle(
            name='RiskMedium',
            parent=styles['Normal'],
            fontSize=16,
            textColor=black,
            alignment=TA_CENTER,
            fontName='Helvetica-Bold',
            backColor=self.colors['warning']
        ))
        
        styles.add(ParagraphStyle(
            name='RiskLow',
            parent=styles['Normal'],
            fontSize=16,
            textColor=white,
            alignment=TA_CENTER,
            fontName='Helvetica-Bold',
            backColor=self.colors['success']
        ))
        
        return styles
    
    def create_matplotlib_chart_as_image(self, chart_function, width=400, height=300):
        """Create matplotlib chart and return as ReportLab Image"""
        # Create matplotlib figure
        fig, ax = plt.subplots(figsize=(width/100, height/100), dpi=100)
        fig.patch.set_facecolor('white')
        
        # Call the chart creation function
        chart_function(ax)
        
        # Save to BytesIO
        img_buffer = BytesIO()
        fig.savefig(img_buffer, format='PNG', dpi=150, bbox_inches='tight', 
                   facecolor='white', edgecolor='none')
        img_buffer.seek(0)
        
        # Create ReportLab Image
        img = Image(img_buffer, width=width, height=height)
        plt.close(fig)
        
        return img
    
    def create_collection_rate_pie_chart(self):
        """Create collection rate pie chart using ReportLab graphics"""
        exec_summary = self.results['executive_summary']
        
        drawing = Drawing(300, 200)
        
        # Create pie chart
        pie = Pie()
        pie.x = 50
        pie.y = 50
        pie.width = 150
        pie.height = 150
        
        # Data
        pie.data = [exec_summary['collection_rate'], 100-exec_summary['collection_rate']]
        pie.labels = ['Collected', 'Outstanding']
        pie.slices.strokeWidth = 2
        pie.slices[0].fillColor = self.colors['success']
        pie.slices[1].fillColor = self.colors['danger']
        
        # Add percentage labels
        pie.slices.labelRadius = 1.2
        pie.slices.fontName = 'Helvetica-Bold'
        pie.slices.fontSize = 12
        
        drawing.add(pie)
        
        # Add title
        title = String(150, 180, 'Collection Rate Performance', 
                      textAnchor='middle', fontSize=14, fontName='Helvetica-Bold')
        drawing.add(title)
        
        # Add center percentage
        center_text = String(125, 125, f'{exec_summary["collection_rate"]:.1f}%',
                           textAnchor='middle', fontSize=16, fontName='Helvetica-Bold',
                           fillColor=self.colors['primary'])
        drawing.add(center_text)
        
        return drawing
    
    def create_aging_bar_chart(self):
        """Create aging analysis bar chart"""
        if self.results['outstanding_invoices_aging'] is None:
            return None
            
        outstanding_aging = self.results['outstanding_invoices_aging']
        
        drawing = Drawing(400, 250)
        
        # Create bar chart
        bc = VerticalBarChart()
        bc.x = 50
        bc.y = 50
        bc.height = 150
        bc.width = 300
        
        # Data
        bc.data = [list(outstanding_aging['Invoice_Count'])]
        bc.categoryAxis.categoryNames = list(outstanding_aging.index)
        
        # Styling
        bc.bars[0].fillColor = self.colors['primary']
        bc.bars[0].strokeColor = self.colors['dark_gray']
        bc.bars[0].strokeWidth = 1
        
        # Axes
        bc.valueAxis.valueMin = 0
        bc.valueAxis.valueMax = max(outstanding_aging['Invoice_Count']) * 1.1
        bc.categoryAxis.labels.boxAnchor = 'ne'
        bc.categoryAxis.labels.dx = -5
        bc.categoryAxis.labels.dy = -5
        bc.categoryAxis.labels.angle = 45
        bc.categoryAxis.labels.fontSize = 10
        
        drawing.add(bc)
        
        # Add title
        title = String(200, 220, 'Outstanding Invoices by Age Category',
                      textAnchor='middle', fontSize=14, fontName='Helvetica-Bold')
        drawing.add(title)
        
        return drawing
    
    def create_client_analysis_chart(self):
        """Create top clients horizontal bar chart"""
        top_clients = self.analyzer.df_invoice.groupby('Applied to')['Amount (USD)'].sum().sort_values(ascending=True).tail(8)
        
        drawing = Drawing(500, 300)
        
        # Create horizontal bar chart
        bc = HorizontalBarChart()
        bc.x = 150
        bc.y = 50
        bc.height = 200
        bc.width = 300
        
        # Data
        bc.data = [list(top_clients.values)]
        bc.categoryAxis.categoryNames = [name[:20] + '...' if len(name) > 20 else name 
                                        for name in top_clients.index]
        
        # Styling
        bc.bars[0].fillColor = self.colors['secondary']
        bc.bars[0].strokeColor = self.colors['dark_gray']
        bc.bars[0].strokeWidth = 1
        
        # Axes
        bc.valueAxis.valueMin = 0
        bc.valueAxis.valueMax = max(top_clients.values) * 1.1
        bc.categoryAxis.labels.fontSize = 9
        bc.valueAxis.labels.fontSize = 9
        
        drawing.add(bc)
        
        # Add title
        title = String(300, 270, 'Top 8 Clients by Total Amount',
                      textAnchor='middle', fontSize=14, fontName='Helvetica-Bold')
        drawing.add(title)
        
        return drawing
    
    def create_metric_boxes(self):
        """Create professional metric display boxes"""
        exec_summary = self.results['executive_summary']
        
        metrics = [
            ('Total Invoices', f"{len(self.analyzer.df_invoice):,}"),
            ('Collection Rate', f"{exec_summary['collection_rate']:.1f}%"),
            ('Avg Payment Delay', f"{exec_summary['avg_payment_delay']:.1f} days"),
            ('Total Outstanding', f"${exec_summary['total_outstanding']:,.0f}")
        ]
        
        # Create table for metrics
        metric_data = []
        for i in range(0, len(metrics), 2):
            row = []
            for j in range(2):
                if i + j < len(metrics):
                    label, value = metrics[i + j]
                    row.append(f'<para align="center"><b>{value}</b><br/><font size="10" color="#6c757d">{label}</font></para>')
                else:
                    row.append('')
            metric_data.append(row)
        
        metric_table = Table(metric_data, colWidths=[2.5*inch, 2.5*inch])
        metric_table.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,0), (-1,-1), 18),
            ('TEXTCOLOR', (0,0), (-1,-1), self.colors['primary']),
            ('BOX', (0,0), (-1,-1), 2, self.colors['primary']),
            ('GRID', (0,0), (-1,-1), 1, self.colors['light_gray']),
            ('BACKGROUND', (0,0), (-1,-1), self.colors['light_gray']),
            ('ROWBACKGROUNDS', (0,0), (-1,-1), [white, self.colors['light_gray']]),
            ('LEFTPADDING', (0,0), (-1,-1), 20),
            ('RIGHTPADDING', (0,0), (-1,-1), 20),
            ('TOPPADDING', (0,0), (-1,-1), 15),
            ('BOTTOMPADDING', (0,0), (-1,-1), 15),
        ]))
        
        return metric_table
    
    def create_risk_indicator(self):
        """Create risk level indicator"""
        exec_summary = self.results['executive_summary']
        risk_percentage = exec_summary['risk_percentage']
        
        if risk_percentage < 10:
            risk_level = "LOW RISK"
            style_name = 'RiskLow'
        elif risk_percentage < 20:
            risk_level = "MODERATE RISK"
            style_name = 'RiskMedium'
        else:
            risk_level = "HIGH RISK"
            style_name = 'RiskHigh'
        
        risk_text = f"<para align='center'><b>OVERALL RISK LEVEL: {risk_level}</b><br/>({risk_percentage:.1f}% of outstanding amount is 90+ days overdue)</para>"
        
        return Paragraph(risk_text, self.styles[style_name])
    
    def create_outstanding_aging_table(self):
        """Create detailed outstanding aging table"""
        if self.results['outstanding_invoices_aging'] is None:
            return None
            
        outstanding_aging = self.results['outstanding_invoices_aging']
        
        # Prepare table data
        headers = ['Age Category', 'Invoice Count', 'Total Amount', 'Total Due', 'Avg Days Overdue', 'Percentage']
        data = [headers]
        
        for idx, row in outstanding_aging.iterrows():
            data.append([
                idx,
                f"{row['Invoice_Count']:,}",
                f"${row['Total_Amount']:,.0f}",
                f"${row['Total_Due']:,.0f}",
                f"{row['Avg_Days_Overdue']:.1f}",
                f"{row['Percentage']:.1f}%"
            ])
        
        # Create table
        table = Table(data, colWidths=[1.2*inch, 0.8*inch, 1*inch, 1*inch, 1*inch, 0.8*inch])
        
        # Style table
        table.setStyle(TableStyle([
            # Header styling
            ('BACKGROUND', (0,0), (-1,0), self.colors['primary']),
            ('TEXTCOLOR', (0,0), (-1,0), white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 10),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            
            # Data styling
            ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,1), (-1,-1), 9),
            ('ALIGN', (1,1), (-1,-1), 'RIGHT'),  # Numbers right-aligned
            ('ALIGN', (0,1), (0,-1), 'LEFT'),    # Category left-aligned
            
            # Borders and grid
            ('GRID', (0,0), (-1,-1), 1, self.colors['dark_gray']),
            ('BOX', (0,0), (-1,-1), 2, self.colors['primary']),
            
            # Alternating row colors
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [white, self.colors['light_gray']]),
            
            # Padding
            ('LEFTPADDING', (0,0), (-1,-1), 8),
            ('RIGHTPADDING', (0,0), (-1,-1), 8),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ]))
        
        return table
    
    def create_client_performance_table(self):
        """Create top clients performance table"""
        # Calculate client statistics
        client_stats = self.analyzer.df_invoice.groupby('Applied to').agg({
            'Number': 'count',
            'Amount (USD)': ['sum', 'mean'],
            'Amt. Paid (USD)': 'sum',
            'Amt. Due (USD)': 'sum'
        }).round(2)
        
        client_stats.columns = ['Invoice_Count', 'Total_Amount', 'Avg_Amount', 'Total_Paid', 'Total_Due']
        client_stats['Collection_Rate'] = (client_stats['Total_Paid'] / client_stats['Total_Amount'] * 100).round(1)
        
        # Get top 10 clients by total amount
        top_clients = client_stats.sort_values('Total_Amount', ascending=False).head(10)
        
        # Prepare table data
        headers = ['Client', 'Invoices', 'Total Amount', 'Total Paid', 'Total Due', 'Collection Rate']
        data = [headers]
        
        for client, row in top_clients.iterrows():
            # Truncate long client names
            client_name = client[:25] + '...' if len(client) > 25 else client
            collection_rate = row['Collection_Rate']
            
            data.append([
                client_name,
                f"{int(row['Invoice_Count']):,}",
                f"${row['Total_Amount']:,.0f}",
                f"${row['Total_Paid']:,.0f}",
                f"${row['Total_Due']:,.0f}",
                f"{collection_rate:.1f}%"
            ])
        
        # Create table
        table = Table(data, colWidths=[2*inch, 0.7*inch, 1*inch, 1*inch, 1*inch, 0.8*inch])
        
        # Style table with performance-based colors
        table_style = [
            # Header styling
            ('BACKGROUND', (0,0), (-1,0), self.colors['secondary']),
            ('TEXTCOLOR', (0,0), (-1,0), white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 10),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            
            # Data styling
            ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,1), (-1,-1), 9),
            ('ALIGN', (1,1), (-1,-1), 'RIGHT'),  # Numbers right-aligned
            ('ALIGN', (0,1), (0,-1), 'LEFT'),    # Client names left-aligned
            
            # Borders and grid
            ('GRID', (0,0), (-1,-1), 1, self.colors['dark_gray']),
            ('BOX', (0,0), (-1,-1), 2, self.colors['secondary']),
            
            # Padding
            ('LEFTPADDING', (0,0), (-1,-1), 8),
            ('RIGHTPADDING', (0,0), (-1,-1), 8),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ]
        
        # Color-code collection rates
        for i, (client, row) in enumerate(top_clients.iterrows(), 1):
            collection_rate = row['Collection_Rate']
            if collection_rate > 95:
                bg_color = HexColor('#d4edda')  # Light green
            elif collection_rate > 85:
                bg_color = HexColor('#fff3cd')  # Light yellow
            elif collection_rate > 70:
                bg_color = HexColor('#ffeaa7')  # Light orange
            else:
                bg_color = HexColor('#f8d7da')  # Light red
            
            table_style.append(('BACKGROUND', (5,i), (5,i), bg_color))
        
        table.setStyle(TableStyle(table_style))
        
        return table
    
    def create_recommendations_section(self):
        """Create recommendations based on analysis"""
        exec_summary = self.results['executive_summary']
        risk_percentage = exec_summary['risk_percentage']
        collection_rate = exec_summary['collection_rate']
        avg_delay = exec_summary['avg_payment_delay']
        
        recommendations = []
        
        # Risk-based recommendations
        if risk_percentage > 20:
            recommendations.extend([
                "🚨 <b>HIGH PRIORITY - IMMEDIATE ACTION REQUIRED:</b>",
                "• Deploy emergency collection procedures for 90+ day receivables",
                "• Implement daily monitoring and escalation protocols",
                "• Review and tighten credit policies immediately",
                "• Consider collection agency for severely overdue accounts",
                "• Establish payment plans for large outstanding amounts",
                ""
            ])
        elif risk_percentage > 10:
            recommendations.extend([
                "⚠️ <b>MODERATE PRIORITY - ENHANCED MONITORING:</b>",
                "• Implement weekly aging reports and collection calls",
                "• Establish payment performance scorecards for clients",
                "• Review credit terms and limits for repeat late payers",
                "• Consider early payment discounts to improve cash flow",
                ""
            ])
        else:
            recommendations.extend([
                "✅ <b>GOOD PERFORMANCE - MAINTAIN CURRENT PROCEDURES:</b>",
                "• Continue existing collection processes",
                "• Maintain regular client relationship management",
                "• Monitor for any deterioration trends",
                ""
            ])
        
        # Collection rate recommendations
        if collection_rate < 85:
            recommendations.extend([
                "<b>COLLECTION EFFICIENCY IMPROVEMENTS:</b>",
                "• Implement more aggressive collection procedures",
                "• Train collection staff on best practices",
                "• Automate payment reminders and follow-ups",
                "• Review client creditworthiness more frequently",
                ""
            ])
        
        # Payment delay recommendations
        if avg_delay > 30:
            recommendations.extend([
                "<b>PAYMENT DELAY MANAGEMENT:</b>",
                "• Review payment terms with chronically late clients",
                "• Implement stricter credit approval processes",
                "• Consider requiring deposits or guarantees",
                "• Improve invoice clarity and delivery methods",
                ""
            ])
        
        # General best practices
        recommendations.extend([
            "<b>ONGOING MONITORING & REPORTING:</b>",
            "• Generate weekly aging reports for management review",
            "• Implement monthly client performance assessments",
            "• Conduct quarterly portfolio risk evaluations",
            "• Annual review of credit policies and procedures",
            "",
            "<b>SUCCESS METRICS TO TRACK:</b>",
            f"• Target collection rate: >95% (currently {collection_rate:.1f}%)",
            f"• Target 90+ day receivables: <5% (currently {risk_percentage:.1f}%)",
            f"• Target average payment delay: <15 days (currently {avg_delay:.1f} days)",
            "• Client satisfaction and retention rates"
        ])
        
        return recommendations
    
    def generate_report(self, output_filename='ReportLab_Professional_Report.pdf'):
        """Generate the complete professional ReportLab report"""
        print("📄 GENERATING REPORTLAB PROFESSIONAL REPORT")
        print("="*60)
        
        output_path = self.reports_dir / output_filename
        
        # Create document
        doc = SimpleDocTemplate(
            str(output_path),
            pagesize=A4,
            rightMargin=0.75*inch,
            leftMargin=0.75*inch,
            topMargin=1*inch,
            bottomMargin=1*inch,
            title="Factoring Risk Analysis Report",
            author="Professional Factoring Analysis System",
            subject="Invoice Portfolio Assessment"
        )
        
        # Build story (document content)
        story = []
        
        # PAGE 1: COVER PAGE
        print("   📝 Creating cover page...")
        story.append(Spacer(1, 0.5*inch))
        
        # Main title
        story.append(Paragraph('FACTORING RISK ANALYSIS', self.styles['CorporateTitle']))
        story.append(Paragraph('Professional Invoice Portfolio Assessment Report', self.styles['Subtitle']))
        story.append(Spacer(1, 0.5*inch))
        
        # Key metrics
        story.append(self.create_metric_boxes())
        story.append(Spacer(1, 0.3*inch))
        
        # Risk indicator
        story.append(self.create_risk_indicator())
        story.append(Spacer(1, 0.3*inch))
        
        # Report details
        report_info = f"""
        <para align='center'>
        <b>Analysis Period:</b> January 2024 - June 2025<br/>
        <b>Report Generated:</b> {datetime.now().strftime('%B %d, %Y at %I:%M %p')}<br/>
        <b>Data Snapshot:</b> June 23, 2025<br/>
        <b>System:</b> Professional Factoring Analysis Platform
        </para>
        """
        story.append(Paragraph(report_info, self.styles['CorporateBody']))
        story.append(Spacer(1, 0.5*inch))
        
        # Confidentiality notice
        confidential = "<para align='center'><b><i>CONFIDENTIAL BUSINESS INFORMATION - INTERNAL USE ONLY</i></b></para>"
        story.append(Paragraph(confidential, self.styles['CorporateBody']))
        
        # PAGE 2: EXECUTIVE SUMMARY
        story.append(PageBreak())
        print("   📊 Creating executive summary...")
        
        story.append(Paragraph('EXECUTIVE SUMMARY', self.styles['SectionHeader']))
        
        # Executive summary content
        exec_summary = self.results['executive_summary']
        
        summary_text = f"""
        <para>
        This comprehensive analysis of the invoice portfolio reveals a collection rate of <b>{exec_summary['collection_rate']:.1f}%</b> 
        with an average payment delay of <b>{exec_summary['avg_payment_delay']:.1f} days</b>. The portfolio contains 
        <b>{len(self.analyzer.df_invoice):,} total invoices</b> with <b>${exec_summary['total_outstanding']:,.0f}</b> 
        currently outstanding across <b>{len(self.analyzer.df_outstanding_invoices):,} invoices</b>.
        </para>
        <para>
        Risk assessment indicates that <b>{exec_summary['risk_percentage']:.1f}%</b> of the outstanding amount 
        is 90+ days overdue, representing a {'<font color="red">HIGH RISK</font>' if exec_summary['risk_percentage'] > 20 else '<font color="orange">MODERATE RISK</font>' if exec_summary['risk_percentage'] > 10 else '<font color="green">LOW RISK</font>'} 
        situation requiring {'immediate intervention' if exec_summary['risk_percentage'] > 20 else 'enhanced monitoring' if exec_summary['risk_percentage'] > 10 else 'continued vigilance'}.
        </para>
        """
        story.append(Paragraph(summary_text, self.styles['CorporateBody']))
        story.append(Spacer(1, 0.3*inch))
        
        # Add collection rate chart
        story.append(Paragraph('Collection Rate Performance', self.styles['SubsectionHeader']))
        story.append(self.create_collection_rate_pie_chart())
        story.append(Spacer(1, 0.3*inch))
        
        # PAGE 3: PORTFOLIO ANALYSIS
        story.append(PageBreak())
        print("   📈 Creating portfolio analysis...")
        
        story.append(Paragraph('PORTFOLIO ANALYSIS', self.styles['SectionHeader']))
        
        # Outstanding aging analysis
        if self.results['outstanding_invoices_aging'] is not None:
            story.append(Paragraph('Outstanding Invoices by Age Category', self.styles['SubsectionHeader']))
            story.append(self.create_aging_bar_chart())
            story.append(Spacer(1, 0.3*inch))
            
            # Detailed aging table
            story.append(Paragraph('Detailed Aging Statistics', self.styles['SubsectionHeader']))
            aging_table = self.create_outstanding_aging_table()
            if aging_table:
                story.append(aging_table)
                story.append(Spacer(1, 0.3*inch))
        
        # PAGE 4: CLIENT ANALYSIS
        story.append(PageBreak())
        print("   👥 Creating client analysis...")
        
        story.append(Paragraph('CLIENT ANALYSIS', self.styles['SectionHeader']))
        
        # Top clients chart
        story.append(Paragraph('Top Clients by Total Amount', self.styles['SubsectionHeader']))
        story.append(self.create_client_analysis_chart())
        story.append(Spacer(1, 0.3*inch))
        
        # Client performance table
        story.append(Paragraph('Client Performance Summary', self.styles['SubsectionHeader']))
        client_table = self.create_client_performance_table()
        if client_table:
            story.append(client_table)
            story.append(Spacer(1, 0.3*inch))
        
        # PAGE 5: RECOMMENDATIONS
        story.append(PageBreak())
        print("   📌 Creating recommendations...")
        
        story.append(Paragraph('RECOMMENDATIONS & ACTION PLAN', self.styles['SectionHeader']))
        
        recommendations = self.create_recommendations_section()
        for rec in recommendations:
            if rec.strip():  # Skip empty lines
                story.append(Paragraph(rec, self.styles['CorporateBody']))
            else:
                story.append(Spacer(1, 0.1*inch))
        
        # Build PDF
        print("   📄 Building PDF document...")
        doc.build(story)
        
        print(f"✅ ReportLab professional report generated: {output_path}")
        return str(output_path)

def generate_reportlab_professional_report(analyzer, results, filename='ReportLab_Professional_Report.pdf'):
    """
    Main function to generate ReportLab professional report
    
    Usage:
    report_path = generate_reportlab_professional_report(analyzer, results)
    """
    
    # Generate the report
    report_generator = ReportLabProfessionalReport(analyzer, results)
    return report_generator.generate_report(filename)

if __name__ == "__main__":
    print("✅ ReportLab Professional Report Generator loaded successfully!")
    print("📄 Ready to generate sophisticated business reports!")