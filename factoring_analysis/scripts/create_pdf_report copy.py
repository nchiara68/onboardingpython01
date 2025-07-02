"""
🏢 Comprehensive Factoring Analysis PDF Report Generator
Using ReportLab Plus Business
Save this file as: factoring_analysis/scripts/create_pdf_report.py
"""

import pandas as pd
import numpy as np
import sys
import os
from datetime import datetime

# ReportLab Plus Business imports
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.platypus import Image, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfgen import canvas

# ReportLab Plus Business advanced features
try:
    from reportlab.graphics.charts.barcharts import VerticalBarChart
    from reportlab.graphics.charts.piecharts import Pie
    from reportlab.graphics.shapes import Drawing
    from reportlab.graphics import renderPDF
    CHARTS_AVAILABLE = True
except ImportError:
    CHARTS_AVAILABLE = False
    print("⚠️ Chart features not available - using tables only")

# Add the scripts directory to Python path
sys.path.append(os.path.dirname(__file__))

try:
    from factoring_analyzer import FactoringAnalyzer
except ImportError:
    print("❌ Could not import FactoringAnalyzer")
    sys.exit(1)

class FactoringPDFReport:
    def __init__(self, csv_path, output_path, page_style='simple'):
        self.csv_path = csv_path
        self.output_path = output_path
        self.page_style = page_style  # 'simple', 'center', or 'header'
        self.analyzer = None
        self.df_paid_invoices = None
        self.df_outstanding_invoices = None
        self.story = []
        
        # Define styles
        self.styles = getSampleStyleSheet()
        self.setup_custom_styles()
        
    def add_page_number_only(self, canvas, doc):
        """Add only page number to each page (simpler version)"""
        page_num = canvas.getPageNumber()
        
        # Save the current state
        canvas.saveState()
        
        # Footer with page number only
        canvas.setFont('Helvetica', 10)
        canvas.drawRightString(letter[0] - 72, 30, f"Page {page_num}")
        
        # Restore the state
        canvas.restoreState()
    
    def add_page_number_center(self, canvas, doc):
        """Add page number centered at bottom"""
        page_num = canvas.getPageNumber()
        
        # Save the current state
        canvas.saveState()
        
        # Footer with centered page number
        canvas.setFont('Helvetica', 10)
        canvas.drawCentredString(letter[0] / 2, 30, f"Page {page_num}")
        
        # Restore the state
        canvas.restoreState()
    
    def add_header_footer(self, canvas, doc):
        """Add header and footer to each page"""
        page_num = canvas.getPageNumber()
        
        # Save the current state
        canvas.saveState()
        
        # Header (optional - you can customize this)
        canvas.setFont('Helvetica-Bold', 12)
        canvas.drawString(72, letter[1] - 50, "Invoice Analysis Report")
        
        # Footer with page number
        canvas.setFont('Helvetica', 10)
        canvas.drawRightString(letter[0] - 72, 30, f"Page {page_num}")
        
        # Add a line under header (optional)
        canvas.setStrokeColor(colors.grey)
        canvas.setLineWidth(0.5)
        canvas.line(72, letter[1] - 60, letter[0] - 72, letter[1] - 60)
        
        # Restore the state
        canvas.restoreState()
        
    def setup_custom_styles(self):
        """Setup custom paragraph styles"""
        # Title style
        self.title_style = ParagraphStyle(
            'CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=24,
            spaceAfter=30,
            alignment=TA_CENTER,
            textColor=colors.darkblue,
            fontName='Helvetica-Bold'
        )
        
        # Section header style
        self.section_style = ParagraphStyle(
            'SectionHeader',
            parent=self.styles['Heading2'],
            fontSize=16,
            spaceAfter=12,
            spaceBefore=20,
            textColor=colors.darkgreen,
            fontName='Helvetica-Bold'
        )
        
        # Table header style
        self.table_header_style = ParagraphStyle(
            'TableHeader',
            parent=self.styles['Heading3'],
            fontSize=12,
            spaceAfter=6,
            spaceBefore=12,
            textColor=colors.black,
            fontName='Helvetica-Bold'
        )
        
        # Normal text style
        self.normal_style = ParagraphStyle(
            'CustomNormal',
            parent=self.styles['Normal'],
            fontSize=10,
            spaceAfter=6,
            fontName='Helvetica'
        )

    def load_and_prepare_data(self):
        """Load data using FactoringAnalyzer and prepare datasets"""
        print("📊 DEBUG: Starting data loading...")
        
        try:
            print(f"🔍 DEBUG: CSV path: {self.csv_path}")
            print(f"🔍 DEBUG: CSV file exists: {os.path.exists(self.csv_path)}")
            
            self.analyzer = FactoringAnalyzer(self.csv_path)
            print("✅ DEBUG: FactoringAnalyzer created successfully")
            
            # Get base invoice data
            base_df = self.analyzer.df_invoice.copy()
            print(f"🔍 DEBUG: Base dataframe shape: {base_df.shape}")
            print(f"🔍 DEBUG: Base dataframe columns: {list(base_df.columns)}")
            
            # Create data separations
            df_paid_raw = base_df[base_df['Amt. Due (USD)'] == 0].copy()
            df_outstanding_raw = base_df[base_df['Amt. Due (USD)'] > 0].copy()
            
            print(f"🔍 DEBUG: Raw paid invoices: {len(df_paid_raw)}")
            print(f"🔍 DEBUG: Raw outstanding invoices: {len(df_outstanding_raw)}")
            print(f"🔍 DEBUG: Credit memos: {len(self.analyzer.df_credit_memo)}")
            
            # Calculate aging correctly
            print("🔧 DEBUG: Calculating aging for paid invoices...")
            self.df_paid_invoices = self.calculate_aging_correctly(df_paid_raw, 'paid')
            print(f"✅ DEBUG: Paid invoices processed: {len(self.df_paid_invoices)}")
            
            print("🔧 DEBUG: Calculating aging for outstanding invoices...")
            self.df_outstanding_invoices = self.calculate_aging_correctly(df_outstanding_raw, 'outstanding')
            print(f"✅ DEBUG: Outstanding invoices processed: {len(self.df_outstanding_invoices)}")
            
            # Debug aging bucket distributions
            if len(self.df_paid_invoices) > 0:
                print(f"🔍 DEBUG: Paid aging distribution: {self.df_paid_invoices['Aging_Bucket'].value_counts().to_dict()}")
            
            if len(self.df_outstanding_invoices) > 0:
                print(f"🔍 DEBUG: Outstanding aging distribution: {self.df_outstanding_invoices['Aging_Bucket'].value_counts().to_dict()}")
            
            print("✅ DEBUG: Data loading and preparation completed")
            
        except Exception as e:
            print(f"❌ DEBUG: Error in data loading: {e}")
            print(f"🔍 DEBUG: Error type: {type(e).__name__}")
            import traceback
            traceback.print_exc()
            raise

    def calculate_aging_correctly(self, df, invoice_type):
        """Calculate aging delay days correctly based on invoice type"""
        if len(df) == 0:
            return df
        
        # Ensure date columns are datetime
        df['Due Date'] = pd.to_datetime(df['Due Date'], errors='coerce')
        if 'Last Payment Date' in df.columns:
            df['Last Payment Date'] = pd.to_datetime(df['Last Payment Date'], errors='coerce')
        
        # Set cutoff date - June 23, 2025
        cutoff_date = pd.Timestamp('2025-06-23')
        
        if invoice_type == 'paid':
            # For paid invoices: Aging Delay Days = Last Payment Date - Due Date
            df['Aging Delay Days'] = (df['Last Payment Date'] - df['Due Date']).dt.days
            df['Aging Delay Days'] = df['Aging Delay Days'].fillna(0)
        elif invoice_type == 'outstanding':
            # For outstanding invoices: Aging Delay Days = June 23, 2025 - Due Date
            df['Aging Delay Days'] = (cutoff_date - df['Due Date']).dt.days
            df['Aging Delay Days'] = df['Aging Delay Days'].fillna(0)
        
        # Create aging buckets
        def create_aging_bucket(days, inv_type):
            if days <= 0:
                return 'On-time' if inv_type == 'paid' else 'Current'
            elif 1 <= days <= 30:
                return '1-30 Days'
            elif 31 <= days <= 60:
                return '31-60 Days'
            elif 61 <= days <= 90:
                return '61-90 Days'
            else:
                return '90+ Days'
        
        df['Aging_Bucket'] = df['Aging Delay Days'].apply(lambda x: create_aging_bucket(x, invoice_type))
        
        return df

    def sort_aging_buckets(self, df, aging_col='Aging_Bucket'):
        """Sort aging buckets in logical order"""
        aging_order = ['On-time', 'Current', '1-30 Days', '31-60 Days', '61-90 Days', '90+ Days']
        
        if aging_col in df.columns:
            df[aging_col] = pd.Categorical(df[aging_col], categories=aging_order, ordered=True)
            df = df.sort_values(aging_col)
            df[aging_col] = df[aging_col].astype(str)
        
        return df

    def format_currency(self, amount):
        """Format amount as currency"""
        return f"${amount:,.0f}"

    def format_percentage(self, value):
        """Format value as percentage"""
        return f"{value:.1f}%"

    def create_pdf_table(self, data, title, col_widths=None, currency_cols=None, percentage_cols=None, pre_formatted=False):
        """Create a styled PDF table"""
        if data.empty:
            return []
        
        # Create table title
        elements = []
        elements.append(Paragraph(title, self.table_header_style))
        
        # Prepare data for table
        table_data = []
        
        # Add headers
        headers = list(data.columns)
        table_data.append(headers)
        
        # Add data rows with formatting
        for _, row in data.iterrows():
            formatted_row = []
            for i, (col, value) in enumerate(row.items()):
                # If pre_formatted is True, don't apply additional formatting
                if pre_formatted:
                    formatted_row.append(str(value) if pd.notna(value) else "")
                elif currency_cols and col in currency_cols:
                    if pd.notna(value) and isinstance(value, (int, float)):
                        formatted_row.append(self.format_currency(value))
                    else:
                        formatted_row.append("$0")
                elif percentage_cols and col in percentage_cols:
                    if pd.notna(value) and isinstance(value, (int, float)):
                        formatted_row.append(self.format_percentage(value))
                    else:
                        formatted_row.append("0.0%")
                else:
                    formatted_row.append(str(value) if pd.notna(value) else "")
            table_data.append(formatted_row)
        
        # Calculate column widths
        if col_widths is None:
            num_cols = len(headers)
            available_width = 7.5 * inch  # Letter size minus margins
            col_widths = [available_width / num_cols] * num_cols
        
        # Create table
        table = Table(table_data, colWidths=col_widths, repeatRows=1)
        
        # Apply table style with special handling for alternating customer rows
        style_commands = [
            # Header row styling
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkgreen),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            
            # Data rows styling
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ]
        
        # Special styling for customer aging tables (detect by title)
        if "Aging Distribution" in title and "Top 15 Customers" in title:
            # Alternate background colors for customer groups (every 2 rows after header)
            for i in range(1, len(table_data), 2):
                if i < len(table_data):
                    style_commands.append(('BACKGROUND', (0, i), (-1, i), colors.lightblue))
                if i + 1 < len(table_data):
                    style_commands.append(('BACKGROUND', (0, i + 1), (-1, i + 1), colors.lightgrey))
        else:
            # Standard alternating row colors
            style_commands.append(('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]))
        
        table.setStyle(TableStyle(style_commands))
        
        elements.append(table)
        elements.append(Spacer(1, 12))
        
        return elements

    def create_summary_section(self):
        """Create executive summary section"""
        self.story.append(Paragraph("Executive Summary", self.section_style))
        
        # Create summary data
        summary_data = {
            'Metric': ['Total Invoices', 'Paid Invoices', 'Outstanding Invoices', 'Credit Memos'],
            'Count': [
                len(self.analyzer.df_invoice),
                len(self.df_paid_invoices),
                len(self.df_outstanding_invoices),
                len(self.analyzer.df_credit_memo)
            ],
            'Total Amount': [
                self.analyzer.df_invoice['Amount (USD)'].sum(),
                self.df_paid_invoices['Amount (USD)'].sum() if len(self.df_paid_invoices) > 0 else 0,
                self.df_outstanding_invoices['Amt. Due (USD)'].sum() if len(self.df_outstanding_invoices) > 0 else 0,
                self.analyzer.df_credit_memo['Amount (USD)'].sum() if len(self.analyzer.df_credit_memo) > 0 else 0
            ]
        }
        
        summary_df = pd.DataFrame(summary_data)
        
        # Add percentages
        total_amount = summary_df.loc[0, 'Total Amount']
        summary_df['Percentage'] = (summary_df['Total Amount'] / total_amount * 100).round(1)
        
        # Create table
        table_elements = self.create_pdf_table(
            summary_df, 
            "Portfolio Overview",
            col_widths=[2*inch, 1*inch, 1.5*inch, 1*inch],
            currency_cols=['Total Amount'],
            percentage_cols=['Percentage']
        )
        
        self.story.extend(table_elements)

    def create_paid_invoices_section(self):
        """Create paid invoices analysis section"""
        if len(self.df_paid_invoices) == 0:
            self.story.append(Paragraph("Paid Invoices Analysis", self.section_style))
            self.story.append(Paragraph("No paid invoices found in the dataset.", self.normal_style))
            return
        
        self.story.append(Paragraph("Paid (Historical) Invoices Analysis", self.section_style))
        self.story.append(Paragraph(f"Total Paid Invoices: {len(self.df_paid_invoices):,}", self.normal_style))
        
        # Table 1.1: Count by Aging Bucket
        paid_count = self.df_paid_invoices.groupby('Aging_Bucket')['Number'].count().reset_index()
        paid_count.columns = ['Aging Bucket', 'Number of Invoices']
        paid_count = self.sort_aging_buckets(paid_count, 'Aging Bucket')
        
        table_elements = self.create_pdf_table(
            paid_count,
            "1.1 Paid Invoices: Count by Aging Bucket",
            col_widths=[3*inch, 2*inch]
        )
        self.story.extend(table_elements)
        
        # Table 1.2: Amount by Aging Bucket
        paid_amount = self.df_paid_invoices.groupby('Aging_Bucket')['Amount (USD)'].sum().reset_index()
        paid_amount.columns = ['Aging Bucket', 'Total Amount']
        paid_amount = self.sort_aging_buckets(paid_amount, 'Aging Bucket')
        
        table_elements = self.create_pdf_table(
            paid_amount,
            "1.2 Paid Invoices: Total Amount by Aging Bucket",
            col_widths=[3*inch, 2*inch],
            currency_cols=['Total Amount']
        )
        self.story.extend(table_elements)
        
        # Table 1.3: Percentages
        total_paid_invoices = len(self.df_paid_invoices)
        total_paid_amount = self.df_paid_invoices['Amount (USD)'].sum()
        
        paid_pct = self.df_paid_invoices.groupby('Aging_Bucket').agg({
            'Number': 'count',
            'Amount (USD)': 'sum'
        }).reset_index()
        
        paid_pct['Count Percentage'] = (paid_pct['Number'] / total_paid_invoices * 100)
        paid_pct['Amount Percentage'] = (paid_pct['Amount (USD)'] / total_paid_amount * 100)
        paid_pct_final = paid_pct[['Aging_Bucket', 'Count Percentage', 'Amount Percentage']].copy()
        paid_pct_final.columns = ['Aging Bucket', 'Count %', 'Amount %']
        paid_pct_final = self.sort_aging_buckets(paid_pct_final, 'Aging Bucket')
        
        table_elements = self.create_pdf_table(
            paid_pct_final,
            "1.3 Paid Invoices: Percentage Distribution by Aging Bucket",
            col_widths=[2.5*inch, 1.5*inch, 1.5*inch],
            percentage_cols=['Count %', 'Amount %']
        )
        self.story.extend(table_elements)
        
        # Table 1.4: Customer Analysis with alternating % and $ rows
        try:
            print("🔧 DEBUG: Creating Table 1.4 - Customer aging analysis...")
            
            # Create pivot table for customers vs aging buckets (TOTAL VALUES)
            customer_aging_paid_values = pd.crosstab(self.df_paid_invoices['Applied to'], 
                                                    self.df_paid_invoices['Aging_Bucket'], 
                                                    values=self.df_paid_invoices['Amount (USD)'], 
                                                    aggfunc='sum')
            customer_aging_paid_values = customer_aging_paid_values.fillna(0)
            
            # Sort columns in logical aging order
            aging_order = ['On-time', 'Current', '1-30 Days', '31-60 Days', '61-90 Days', '90+ Days']
            available_cols = [col for col in aging_order if col in customer_aging_paid_values.columns]
            customer_aging_paid_values = customer_aging_paid_values.reindex(columns=available_cols, fill_value=0)
            
            # Calculate percentages (row-wise) - handle division by zero
            customer_aging_paid_pct = customer_aging_paid_values.div(customer_aging_paid_values.sum(axis=1), axis=0) * 100
            customer_aging_paid_pct = customer_aging_paid_pct.fillna(0).round(1)
            
            # Get top 15 customers by total paid amount for readability
            top_customers_paid = self.df_paid_invoices.groupby('Applied to')['Amount (USD)'].sum().nlargest(15).index
            
            if len(top_customers_paid) > 0:
                # Filter for top customers
                customer_aging_paid_values_top = customer_aging_paid_values.loc[top_customers_paid]
                customer_aging_paid_pct_top = customer_aging_paid_pct.loc[top_customers_paid]
                
                # Create combined table with alternating percentage and value rows
                combined_data = []
                
                for customer in top_customers_paid:
                    # Add percentage row
                    pct_row = [customer + " (%)"] + [f"{customer_aging_paid_pct_top.loc[customer, col]:.1f}%" 
                                                    for col in available_cols]
                    combined_data.append(pct_row)
                    
                    # Add value row
                    val_row = [customer + " ($)"] + [f"${customer_aging_paid_values_top.loc[customer, col]:,.0f}" 
                                                    for col in available_cols]
                    combined_data.append(val_row)
                
                # Create DataFrame
                column_names = ['Customer'] + available_cols
                customer_combined_df = pd.DataFrame(combined_data, columns=column_names)
                
                # Calculate column widths - make customer column wider
                customer_col_width = 3*inch
                aging_col_width = (7.5*inch - customer_col_width) / len(available_cols)
                col_widths = [customer_col_width] + [aging_col_width] * len(available_cols)
                
                # Create table without currency/percentage formatting (already formatted)
                table_elements = self.create_pdf_table(
                    customer_combined_df,
                    "1.4 Top 15 Customers: Aging Distribution (% and $ Values)",
                    col_widths=col_widths,
                    pre_formatted=True
                )
                self.story.extend(table_elements)
                
                print("✅ DEBUG: Table 1.4 created successfully")
            else:
                print("⚠️ DEBUG: No customers found for Table 1.4")
                
        except Exception as e:
            print(f"❌ DEBUG: Error creating Table 1.4: {e}")
            # Fallback to simple customer table
            customer_paid = self.df_paid_invoices.groupby('Applied to').agg({
                'Number': 'count',
                'Amount (USD)': 'sum'
            }).reset_index()
            customer_paid.columns = ['Customer', 'Invoice Count', 'Total Amount']
            customer_paid = customer_paid.sort_values('Total Amount', ascending=False).head(15)
            
            table_elements = self.create_pdf_table(
                customer_paid,
                "1.4 Top 15 Customers: Paid Invoices (Fallback)",
                col_widths=[3*inch, 1.5*inch, 2*inch],
                currency_cols=['Total Amount']
            )
            self.story.extend(table_elements)

    def create_outstanding_invoices_section(self):
        """Create outstanding invoices analysis section"""
        if len(self.df_outstanding_invoices) == 0:
            self.story.append(Paragraph("Outstanding Invoices Analysis", self.section_style))
            self.story.append(Paragraph("No outstanding invoices found in the dataset.", self.normal_style))
            return
        
        self.story.append(PageBreak())
        self.story.append(Paragraph("Outstanding Invoices Analysis", self.section_style))
        self.story.append(Paragraph(f"Total Outstanding Invoices: {len(self.df_outstanding_invoices):,}", self.normal_style))
        
        # Table 2.1: Count by Aging Bucket
        out_count = self.df_outstanding_invoices.groupby('Aging_Bucket')['Number'].count().reset_index()
        out_count.columns = ['Aging Bucket', 'Number of Invoices']
        out_count = self.sort_aging_buckets(out_count, 'Aging Bucket')
        
        table_elements = self.create_pdf_table(
            out_count,
            "2.1 Outstanding Invoices: Count by Aging Bucket",
            col_widths=[3*inch, 2*inch]
        )
        self.story.extend(table_elements)
        
        # Table 2.2: Amount by Aging Bucket
        out_amount = self.df_outstanding_invoices.groupby('Aging_Bucket')['Amt. Due (USD)'].sum().reset_index()
        out_amount.columns = ['Aging Bucket', 'Total Amount Due']
        out_amount = self.sort_aging_buckets(out_amount, 'Aging Bucket')
        
        table_elements = self.create_pdf_table(
            out_amount,
            "2.2 Outstanding Invoices: Total Amount Due by Aging Bucket",
            col_widths=[3*inch, 2*inch],
            currency_cols=['Total Amount Due']
        )
        self.story.extend(table_elements)
        
        # Table 2.3: Percentages
        total_out_invoices = len(self.df_outstanding_invoices)
        total_out_amount = self.df_outstanding_invoices['Amt. Due (USD)'].sum()
        
        out_pct = self.df_outstanding_invoices.groupby('Aging_Bucket').agg({
            'Number': 'count',
            'Amt. Due (USD)': 'sum'
        }).reset_index()
        
        out_pct['Count Percentage'] = (out_pct['Number'] / total_out_invoices * 100)
        out_pct['Amount Percentage'] = (out_pct['Amt. Due (USD)'] / total_out_amount * 100)
        out_pct_final = out_pct[['Aging_Bucket', 'Count Percentage', 'Amount Percentage']].copy()
        out_pct_final.columns = ['Aging Bucket', 'Count %', 'Amount %']
        out_pct_final = self.sort_aging_buckets(out_pct_final, 'Aging Bucket')
        
        table_elements = self.create_pdf_table(
            out_pct_final,
            "2.3 Outstanding Invoices: Percentage Distribution by Aging Bucket",
            col_widths=[2.5*inch, 1.5*inch, 1.5*inch],
            percentage_cols=['Count %', 'Amount %']
        )
        self.story.extend(table_elements)
        
        # Table 2.4: Customer Analysis with alternating % and $ rows
        try:
            print("🔧 DEBUG: Creating Table 2.4 - Outstanding customer aging analysis...")
            
            # Create pivot table for customers vs aging buckets (TOTAL VALUES)
            customer_aging_outstanding_values = pd.crosstab(self.df_outstanding_invoices['Applied to'], 
                                                          self.df_outstanding_invoices['Aging_Bucket'], 
                                                          values=self.df_outstanding_invoices['Amt. Due (USD)'], 
                                                          aggfunc='sum')
            customer_aging_outstanding_values = customer_aging_outstanding_values.fillna(0)
            
            # Sort columns in logical aging order
            aging_order = ['On-time', 'Current', '1-30 Days', '31-60 Days', '61-90 Days', '90+ Days']
            available_cols = [col for col in aging_order if col in customer_aging_outstanding_values.columns]
            customer_aging_outstanding_values = customer_aging_outstanding_values.reindex(columns=available_cols, fill_value=0)
            
            # Calculate percentages (row-wise) - handle division by zero
            customer_aging_outstanding_pct = customer_aging_outstanding_values.div(customer_aging_outstanding_values.sum(axis=1), axis=0) * 100
            customer_aging_outstanding_pct = customer_aging_outstanding_pct.fillna(0).round(1)
            
            # Get top 15 customers by total outstanding amount for readability
            top_customers_outstanding = self.df_outstanding_invoices.groupby('Applied to')['Amt. Due (USD)'].sum().nlargest(15).index
            
            if len(top_customers_outstanding) > 0:
                # Filter for top customers
                customer_aging_outstanding_values_top = customer_aging_outstanding_values.loc[top_customers_outstanding]
                customer_aging_outstanding_pct_top = customer_aging_outstanding_pct.loc[top_customers_outstanding]
                
                # Create combined table with alternating percentage and value rows
                combined_data = []
                
                for customer in top_customers_outstanding:
                    # Add percentage row
                    pct_row = [customer + " (%)"] + [f"{customer_aging_outstanding_pct_top.loc[customer, col]:.1f}%" 
                                                    for col in available_cols]
                    combined_data.append(pct_row)
                    
                    # Add value row
                    val_row = [customer + " ($)"] + [f"${customer_aging_outstanding_values_top.loc[customer, col]:,.0f}" 
                                                    for col in available_cols]
                    combined_data.append(val_row)
                
                # Create DataFrame
                column_names = ['Customer'] + available_cols
                customer_combined_df = pd.DataFrame(combined_data, columns=column_names)
                
                # Calculate column widths - make customer column wider
                customer_col_width = 3*inch
                aging_col_width = (7.5*inch - customer_col_width) / len(available_cols)
                col_widths = [customer_col_width] + [aging_col_width] * len(available_cols)
                
                # Create table without currency/percentage formatting (already formatted)
                table_elements = self.create_pdf_table(
                    customer_combined_df,
                    "2.4 Top 15 Customers: Outstanding Aging Distribution (% and $ Values)",
                    col_widths=col_widths,
                    pre_formatted=True
                )
                self.story.extend(table_elements)
                
                print("✅ DEBUG: Table 2.4 created successfully")
            else:
                print("⚠️ DEBUG: No customers found for Table 2.4")
                
        except Exception as e:
            print(f"❌ DEBUG: Error creating Table 2.4: {e}")
            # Fallback to simple customer table
            customer_out = self.df_outstanding_invoices.groupby('Applied to').agg({
                'Number': 'count',
                'Amt. Due (USD)': 'sum'
            }).reset_index()
            customer_out.columns = ['Customer', 'Invoice Count', 'Total Amount Due']
            customer_out = customer_out.sort_values('Total Amount Due', ascending=False).head(15)
            
            table_elements = self.create_pdf_table(
                customer_out,
                "2.4 Top 15 Customers: Outstanding Invoices (Fallback)",
                col_widths=[3*inch, 1.5*inch, 2*inch],
                currency_cols=['Total Amount Due']
            )
            self.story.extend(table_elements)

    def create_credit_memo_section(self):
        """Create credit memo analysis section"""
        if len(self.analyzer.df_credit_memo) == 0:
            self.story.append(Paragraph("Credit Memo Analysis", self.section_style))
            self.story.append(Paragraph("No credit memos found in the dataset.", self.normal_style))
            return
        
        self.story.append(PageBreak())
        self.story.append(Paragraph("Credit Memo Analysis", self.section_style))
        self.story.append(Paragraph(f"Total Credit Memos: {len(self.analyzer.df_credit_memo):,}", self.normal_style))
        
        # Table 3.1: Credit by Customer
        credit_by_customer = self.analyzer.df_credit_memo.groupby('Applied to')['Amount (USD)'].sum().reset_index()
        credit_by_customer.columns = ['Customer', 'Total Credit Amount']
        credit_by_customer = credit_by_customer.sort_values('Total Credit Amount', ascending=False).head(20)
        
        # Add percentage
        total_credit_amount = self.analyzer.df_credit_memo['Amount (USD)'].sum()
        credit_by_customer['Percentage'] = (credit_by_customer['Total Credit Amount'] / total_credit_amount * 100)
        
        table_elements = self.create_pdf_table(
            credit_by_customer,
            "3.1 Credit Memos: Total Amount by Customer",
            col_widths=[3*inch, 2*inch, 1.5*inch],
            currency_cols=['Total Credit Amount'],
            percentage_cols=['Percentage']
        )
        self.story.extend(table_elements)
        
        # Table 3.2: Credit Summary Statistics
        credit_stats = pd.DataFrame({
            'Metric': ['Total Credit Memos', 'Total Credit Amount', 'Average Credit Amount', 
                      'Largest Credit Memo', 'Customers with Credits'],
            'Value': [
                f"{len(self.analyzer.df_credit_memo):,}",
                self.format_currency(self.analyzer.df_credit_memo['Amount (USD)'].sum()),
                self.format_currency(self.analyzer.df_credit_memo['Amount (USD)'].mean()),
                self.format_currency(self.analyzer.df_credit_memo['Amount (USD)'].max()),
                f"{self.analyzer.df_credit_memo['Applied to'].nunique():,}"
            ]
        })
        
        table_elements = self.create_pdf_table(
            credit_stats,
            "3.2 Credit Memo Summary Statistics",
            col_widths=[3*inch, 2*inch]
        )
        self.story.extend(table_elements)

    def generate_pdf_report(self):
        """Generate the complete PDF report"""
        print("📄 DEBUG: Starting PDF report generation...")
        
        try:
            # Load and prepare data
            print("🔧 DEBUG: Loading and preparing data...")
            self.load_and_prepare_data()
            print("✅ DEBUG: Data loading completed")
            
            # Create PDF document with appropriate margins based on page style
            print(f"🔧 DEBUG: Creating PDF document at: {self.output_path}")
            print(f"🔍 DEBUG: Output directory exists: {os.path.exists(os.path.dirname(self.output_path))}")
            print(f"🔍 DEBUG: Page style: {self.page_style}")
            
            # Adjust margins based on page style
            if self.page_style == 'header':
                top_margin = 90  # More space for header
                bottom_margin = 50
            else:
                top_margin = 72  # Standard margin
                bottom_margin = 50  # Space for page number
            
            doc = SimpleDocTemplate(
                self.output_path,
                pagesize=letter,
                rightMargin=72,
                leftMargin=72,
                topMargin=top_margin,
                bottomMargin=bottom_margin
            )
            print("✅ DEBUG: SimpleDocTemplate created")
            
            # Add title
            title = f"Invoice Analysis Report"
            subtitle = f"Generated on {datetime.now().strftime('%B %d, %Y')}"
            
            print("🔧 DEBUG: Adding title and subtitle...")
            self.story.append(Paragraph(title, self.title_style))
            self.story.append(Paragraph(subtitle, self.normal_style))
            self.story.append(Spacer(1, 30))
            print("✅ DEBUG: Title added to story")
            
            # Add sections
            print("🔧 DEBUG: Creating summary section...")
            self.create_summary_section()
            print("✅ DEBUG: Summary section completed")
            
            print("🔧 DEBUG: Creating paid invoices section...")
            self.create_paid_invoices_section()
            print("✅ DEBUG: Paid invoices section completed")
            
            print("🔧 DEBUG: Creating outstanding invoices section...")
            self.create_outstanding_invoices_section()
            print("✅ DEBUG: Outstanding invoices section completed")
            
            print("🔧 DEBUG: Creating credit memo section...")
            self.create_credit_memo_section()
            print("✅ DEBUG: Credit memo section completed")
            
            # Check story content before building
            print(f"🔍 DEBUG: Story elements count: {len(self.story)}")
            print(f"🔍 DEBUG: Story element types: {[type(elem).__name__ for elem in self.story[:5]]}")
            
            # Build PDF with appropriate page numbering style
            print("🔧 DEBUG: Building PDF document with page numbers...")
            
            # Choose page numbering function based on style
            if self.page_style == 'header':
                page_func = self.add_header_footer
            elif self.page_style == 'center':
                page_func = self.add_page_number_center
            else:  # 'simple' or default
                page_func = self.add_page_number_only
            
            print(f"🔍 DEBUG: Using page numbering style: {self.page_style}")
            
            doc.build(self.story, 
                     onFirstPage=page_func,
                     onLaterPages=page_func)
            print("✅ DEBUG: PDF document built successfully")
            
            # Verify file creation
            if os.path.exists(self.output_path):
                file_size = os.path.getsize(self.output_path)
                print(f"✅ DEBUG: PDF file created successfully - Size: {file_size} bytes")
            else:
                print(f"❌ DEBUG: PDF file not found after build: {self.output_path}")
                return None
            
            print(f"✅ PDF report created: {self.output_path}")
            return self.output_path
            
        except Exception as e:
            print(f"❌ DEBUG: Error in PDF generation: {e}")
            print(f"🔍 DEBUG: Error type: {type(e).__name__}")
            import traceback
            traceback.print_exc()
            
            # Check if partial file exists
            if os.path.exists(self.output_path):
                print(f"🔍 DEBUG: Partial file found: {self.output_path}")
                try:
                    size = os.path.getsize(self.output_path)
                    print(f"🔍 DEBUG: Partial file size: {size} bytes")
                except:
                    print(f"🔍 DEBUG: Could not get partial file size")
            
            raise

def create_factoring_pdf_report(page_style='simple'):
    """
    Main function to create PDF report
    page_style options:
    - 'simple': Page numbers only (bottom right)
    - 'center': Page numbers centered at bottom
    - 'header': Full header with report title and page numbers
    """
    print("🏢 FACTORING ANALYSIS PDF REPORT GENERATOR")
    print("="*60)
    
    # DEBUG: Show current working directory
    current_working_dir = os.getcwd()
    print(f"🔍 DEBUG: Current working directory: {current_working_dir}")
    
    # File paths
    current_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"🔍 DEBUG: Script directory: {current_dir}")
    
    project_root = os.path.dirname(current_dir)
    print(f"🔍 DEBUG: Project root: {project_root}")
    
    csv_path = os.path.join(project_root, 'data', 'TecnoCargoInvoiceDataset01.csv')
    print(f"🔍 DEBUG: CSV path: {csv_path}")
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"factoring_analysis_report_{timestamp}.pdf"
    output_path = os.path.join(project_root, 'output', output_filename)
    print(f"🔍 DEBUG: Output filename: {output_filename}")
    print(f"🔍 DEBUG: Full output path: {output_path}")
    
    # DEBUG: Check if directories exist
    print(f"🔍 DEBUG: Project root exists: {os.path.exists(project_root)}")
    print(f"🔍 DEBUG: Data directory exists: {os.path.exists(os.path.dirname(csv_path))}")
    
    # Create output directory with debugging
    output_dir = os.path.dirname(output_path)
    print(f"🔍 DEBUG: Output directory: {output_dir}")
    print(f"🔍 DEBUG: Output directory exists before creation: {os.path.exists(output_dir)}")
    
    try:
        os.makedirs(output_dir, exist_ok=True)
        print(f"✅ DEBUG: Output directory created/verified: {output_dir}")
        print(f"🔍 DEBUG: Output directory exists after creation: {os.path.exists(output_dir)}")
        
        # Check directory permissions
        if os.path.exists(output_dir):
            print(f"🔍 DEBUG: Output directory is writable: {os.access(output_dir, os.W_OK)}")
        
    except Exception as e:
        print(f"❌ DEBUG: Failed to create output directory: {e}")
        return None
    
    print(f"📁 Input: {csv_path}")
    print(f"📁 Output: {output_path}")
    
    # Check if input file exists
    if not os.path.exists(csv_path):
        print(f"❌ Input file not found: {csv_path}")
        # DEBUG: List what's actually in the data directory
        data_dir = os.path.dirname(csv_path)
        if os.path.exists(data_dir):
            print(f"🔍 DEBUG: Files in data directory: {os.listdir(data_dir)}")
        else:
            print(f"🔍 DEBUG: Data directory doesn't exist: {data_dir}")
        return None
    
    print(f"✅ DEBUG: Input file found and accessible")
    
    # DEBUG: Test ReportLab import
    try:
        from reportlab.lib import colors
        from reportlab.platypus import SimpleDocTemplate
        print(f"✅ DEBUG: ReportLab imports successful")
    except ImportError as e:
        print(f"❌ DEBUG: ReportLab import failed: {e}")
        return None
    
    try:
        # Create report generator with page numbering
        print(f"🔧 DEBUG: Creating FactoringPDFReport instance with page style: {page_style}")
        report_generator = FactoringPDFReport(csv_path, output_path, page_style=page_style)
        print(f"✅ DEBUG: FactoringPDFReport instance created")
        
        # Generate PDF
        print(f"🔧 DEBUG: Starting PDF generation...")
        output_file = report_generator.generate_pdf_report()
        print(f"✅ DEBUG: PDF generation completed")
        
        # DEBUG: Check if file was actually created
        if output_file and os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print(f"✅ DEBUG: PDF file exists with size: {file_size} bytes")
            print(f"🔍 DEBUG: File location confirmed: {output_file}")
            
            # List files in output directory
            output_dir = os.path.dirname(output_file)
            print(f"🔍 DEBUG: Files in output directory: {os.listdir(output_dir)}")
            
        else:
            print(f"❌ DEBUG: PDF file was not created or not found at: {output_path}")
            # Check what's in the output directory
            if os.path.exists(output_dir):
                print(f"🔍 DEBUG: Files in output directory: {os.listdir(output_dir)}")
            
        print("\n🎉 PDF Report Generation Completed!")
        print(f"📄 Your professional PDF report is ready: {output_file}")
        
        return output_file
        
    except Exception as e:
        print(f"\n❌ ERROR: PDF generation failed: {e}")
        print(f"🔍 DEBUG: Error type: {type(e).__name__}")
        import traceback
        traceback.print_exc()
        
        # Additional debugging - check if partial file was created
        if os.path.exists(output_path):
            print(f"🔍 DEBUG: Partial file exists: {output_path}")
            print(f"🔍 DEBUG: Partial file size: {os.path.getsize(output_path)} bytes")
        
        return None

def test_setup():
    """Test basic setup before running full PDF generation"""
    print("🧪 TESTING BASIC SETUP")
    print("="*40)
    
    # Test 1: Current directory and paths
    current_working_dir = os.getcwd()
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(current_dir)
    
    print(f"🔍 Current working directory: {current_working_dir}")
    print(f"🔍 Script directory: {current_dir}")
    print(f"🔍 Project root: {project_root}")
    
    # Test 2: File paths
    csv_path = os.path.join(project_root, 'data', 'TecnoCargoInvoiceDataset01.csv')
    output_dir = os.path.join(project_root, 'output')
    
    print(f"🔍 CSV path: {csv_path}")
    print(f"🔍 Output directory: {output_dir}")
    
    # Test 3: File existence
    print(f"📁 Project root exists: {os.path.exists(project_root)}")
    print(f"📁 CSV file exists: {os.path.exists(csv_path)}")
    print(f"📁 Output directory exists: {os.path.exists(output_dir)}")
    
    # Test 4: List directories
    if os.path.exists(project_root):
        print(f"📋 Project root contents: {os.listdir(project_root)}")
    
    data_dir = os.path.dirname(csv_path)
    if os.path.exists(data_dir):
        print(f"📋 Data directory contents: {os.listdir(data_dir)}")
    
    # Test 5: ReportLab imports
    try:
        from reportlab.lib import colors
        from reportlab.platypus import SimpleDocTemplate, Paragraph
        from reportlab.lib.styles import getSampleStyleSheet
        print("✅ ReportLab imports successful")
    except ImportError as e:
        print(f"❌ ReportLab import failed: {e}")
        return False
    
    # Test 6: FactoringAnalyzer import
    try:
        from factoring_analyzer import FactoringAnalyzer
        print("✅ FactoringAnalyzer import successful")
    except ImportError as e:
        print(f"❌ FactoringAnalyzer import failed: {e}")
        return False
    
    # Test 7: Create output directory
    try:
        os.makedirs(output_dir, exist_ok=True)
        print(f"✅ Output directory created/verified: {output_dir}")
        print(f"📁 Directory writable: {os.access(output_dir, os.W_OK)}")
    except Exception as e:
        print(f"❌ Cannot create output directory: {e}")
        return False
    
    # Test 8: Try basic data loading
    if os.path.exists(csv_path):
        try:
            analyzer = FactoringAnalyzer(csv_path)
            print(f"✅ FactoringAnalyzer loaded data successfully")
            print(f"📊 Invoice count: {len(analyzer.df_invoice)}")
        except Exception as e:
            print(f"❌ Data loading failed: {e}")
            return False
    
    print("🎉 Basic setup test completed successfully!")
    return True

if __name__ == "__main__":
    # Add option to run test first or specify page style
    import sys
    
    if len(sys.argv) > 1:
        if sys.argv[1] == "--test":
            test_setup()
        elif sys.argv[1] == "--help":
            print("📄 PDF Report Generator Options:")
            print("  python create_pdf_report.py --test         # Test setup")
            print("  python create_pdf_report.py --simple       # Simple page numbers (default)")
            print("  python create_pdf_report.py --center       # Centered page numbers")
            print("  python create_pdf_report.py --header       # Full header with page numbers")
            print("  python create_pdf_report.py --help         # Show this help")
        elif sys.argv[1] == "--simple":
            create_factoring_pdf_report(page_style='simple')
        elif sys.argv[1] == "--center":
            create_factoring_pdf_report(page_style='center')
        elif sys.argv[1] == "--header":
            create_factoring_pdf_report(page_style='header')
        else:
            print(f"Unknown option: {sys.argv[1]}")
            print("Use --help to see available options")
    else:
        # Default: simple page numbering
        create_factoring_pdf_report(page_style='simple')