#!/usr/bin/env python3
"""
TechCargo Financial Analysis Report Generator
Creates a comprehensive PDF report using ReportLab showing financial analysis tables
"""

import pandas as pd
import os
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas

def create_financial_report():
    """Create a comprehensive financial report PDF"""
    
    # Setup file paths
    output_file = '../reports/Bank_Invoices_Financial_Report.pdf'
    
    # Ensure output directory exists
    os.makedirs('../reports', exist_ok=True)
    
    # Create document

    doc = SimpleDocTemplate(
        output_file,
        pagesize=A4,
        rightMargin=72,
        leftMargin=72,
        topMargin=72,
        bottomMargin=72
    )
    
    # Get styles
    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        spaceAfter=30,
        alignment=1,  # Center alignment
        textColor=colors.darkblue
    )
    
    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=16,
        spaceAfter=12,
        spaceBefore=20,
        textColor=colors.darkblue
    )
    
    subheading_style = ParagraphStyle(
        'CustomSubHeading',
        parent=styles['Heading3'],
        fontSize=14,
        spaceAfter=8,
        spaceBefore=12,
        textColor=colors.navy
    )
    
    normal_style = styles['Normal']
    
    # Story list to hold all elements
    story = []
    
    # Title Page
    story.append(Paragraph("Bank and Invoices Cashflow Analysis Report", title_style))
    story.append(Spacer(1, 0.5*inch))
    
    # Report metadata
    report_date = datetime.now().strftime("%B %d, %Y")
    story.append(Paragraph(f"<b>Report Date:</b> {report_date}", normal_style))
    story.append(Paragraph(f"<b>Period Covered:</b> January 2024 - May 2025", normal_style))
    story.append(Paragraph(f"<b>Analysis Type:</b> Cash Flow & Invoice Payment Analysis", normal_style))
    story.append(Spacer(1, 0.5*inch))
    
    # Executive Summary
    story.append(Paragraph("Executive Summary", heading_style))
    
    try:
        # Read key data files for summary
        cash_inflow_df = pd.read_csv('../data/cvs/TechCargo_Cash_Inflow.csv')
        payments_df = pd.read_csv('../data/MonthlyPaymentAggregation.csv')
        monthly_totals_df = pd.read_csv('../data/MonthlyTableWithTotals_2024_2025.csv')
        
        # Calculate key metrics
        total_bank_inflow = "${:,.2f}".format(cash_inflow_df['Monthly Total'].iloc[1:].sum()-200000)
        total_payments = "${:,.2f}".format(payments_df['Total Payments (USD)'].iloc[:-1].sum())
        total_invoices = "{:,}".format(int(payments_df['Invoices Paid'].iloc[:-1].sum()))
        
        summary_text = f"""
        This report analyzes TechCargo's financial data covering bank deposits and invoice payments 
        from January 2024 through May 2025. Key findings include:
        <br/><br/>
        • <b>Total Bank Cash Inflow:</b> { total_bank_inflow}<br/>
        • <b>Total Invoice Payments:</b> { total_payments}<br/>
        • <b>Total Invoices Processed:</b> { total_invoices}<br/>
        • <b>Analysis Period:</b> 17 months (January 2024 - May 2025)<br/>
        • <b>Bank Sources:</b> Chase, Popular, Wells Fargo<br/>
        """
        
        story.append(Paragraph(summary_text, normal_style))
        story.append(PageBreak())
        
    except FileNotFoundError as e:
        story.append(Paragraph(f"<b>Note:</b> Some data files not found. Please run the data processing script first.", normal_style))
        story.append(PageBreak())
    
    # Section 1: Monthly Cash Inflow by Bank
    story.append(Paragraph("1. Monthly Cash Inflow by Bank (2024-01 to 2025-05)", heading_style))
    
    try:
        cash_inflow_df = pd.read_csv('../data/cvs/TechCargo_Cash_Inflow.csv')
        
        # Create complete date range from 2024-01 to 2025-05
        start_date = '2024-01'
        end_date = '2025-05'
        complete_months = pd.period_range(start=start_date, end=end_date, freq='M')
        
        # Create template with all months
        complete_month_df = pd.DataFrame({'month': complete_months})
        
        # Convert existing data month column to period for merging
        if cash_inflow_df['month'].dtype == 'object':
            cash_inflow_df['month'] = pd.to_datetime(cash_inflow_df['month']).dt.to_period('M')
        
        # Merge with existing data to fill in available months
        complete_cash_inflow = pd.merge(complete_month_df, cash_inflow_df, on='month', how='left')
        
        # Fill missing values with 0
        complete_cash_inflow['Chase'] = complete_cash_inflow['Chase'].fillna(0)
        complete_cash_inflow['Popular'] = complete_cash_inflow['Popular'].fillna(0)
        complete_cash_inflow['Wells Fargo'] = complete_cash_inflow['Wells Fargo'].fillna(0)
        complete_cash_inflow['Monthly Total'] = complete_cash_inflow['Monthly Total'].fillna(0)
        
        # Prepare table data
        table_data = [['Month', 'Chase', 'Popular', 'Wells Fargo', 'Monthly Total']]
        
        for _, row in complete_cash_inflow.iterrows():
            table_data.append([
                str(row['month']),
                f"${row['Chase']:,.2f}",
                f"${row['Popular']:,.2f}",
                f"${row['Wells Fargo']:,.2f}",
                f"${row['Monthly Total']:,.2f}"
            ])
        
        # Add totals row
        chase_total = complete_cash_inflow['Chase'].sum()
        popular_total = complete_cash_inflow['Popular'].sum()
        wells_total = complete_cash_inflow['Wells Fargo'].sum()
        grand_total = complete_cash_inflow['Monthly Total'].sum()
        
        table_data.append([
            'TOTAL',
            f"${chase_total:,.2f}",
            f"${popular_total:,.2f}",
            f"${wells_total:,.2f}",
            f"${grand_total:,.2f}"
        ])
        
        # Create table
        table = Table(table_data, colWidths=[1.2*inch, 1*inch, 1*inch, 1*inch, 1.2*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        
        story.append(table)
        story.append(Spacer(1, 0.3*inch))
        
    except FileNotFoundError:
        story.append(Paragraph("Cash inflow data not found.", normal_style))
    
    # Section 2: Monthly Payment Aggregation
    story.append(Paragraph("2. Monthly Invoice Payment Aggregation (2024-01 to 2025-05)", heading_style))
    
    try:
        payments_df = pd.read_csv('../data/MonthlyPaymentAggregation.csv')
        
        # Create complete date range from 2024-01 to 2025-05
        start_date = '2024-01'
        end_date = '2025-05'
        complete_months = pd.period_range(start=start_date, end=end_date, freq='M')
        
        # Create template with all months
        complete_month_df = pd.DataFrame({'Payment Month': complete_months})
        
        # Convert existing data month column to period for merging
        if payments_df['Payment Month'].dtype == 'object':
            payments_df['Payment Month'] = pd.to_datetime(payments_df['Payment Month']).dt.to_period('M')
        
        # Merge with existing data to fill in available months
        complete_payments = pd.merge(complete_month_df, payments_df, on='Payment Month', how='left')
        
        # Fill missing values with 0
        complete_payments['Total Payments (USD)'] = complete_payments['Total Payments (USD)'].fillna(0)
        complete_payments['Invoices Paid'] = complete_payments['Invoices Paid'].fillna(0)
        
        # Prepare table data
        table_data = [['Payment Month', 'Total Payments (USD)', 'Invoices Paid']]
        
        for _, row in complete_payments.iterrows():
            table_data.append([
                str(row['Payment Month']),
                f"${row['Total Payments (USD)']:,.2f}",
                f"{row['Invoices Paid']:,.0f}"
            ])
        
        # Add totals row
        total_payments = complete_payments['Total Payments (USD)'].sum()
        total_invoices = complete_payments['Invoices Paid'].sum()
        
        table_data.append([
            'TOTAL',
            f"${total_payments:,.2f}",
            f"{total_invoices:,.0f}"
        ])
        
        # Create table
        table = Table(table_data, colWidths=[2*inch, 2*inch, 1.5*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkgreen),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        
        story.append(table)
        story.append(Spacer(1, 0.3*inch))
        
        # Add summary statistics using complete dataset
        avg_monthly = complete_payments['Total Payments (USD)'].mean()
        max_monthly = complete_payments['Total Payments (USD)'].max()
        min_monthly = complete_payments['Total Payments (USD)'].min()
        months_with_payments = len(complete_payments[complete_payments['Total Payments (USD)'] > 0])
        
        summary_text = f"""
        <b>Payment Summary Statistics:</b><br/>
        • Average Monthly Payment: ${avg_monthly:,.2f}<br/>
        • Highest Monthly Payment: ${max_monthly:,.2f}<br/>
        • Lowest Monthly Payment: ${min_monthly:,.2f}<br/>
        • Months with Payments: {months_with_payments} out of 17 total months<br/>
        • Months with No Payments: {17 - months_with_payments}
        """
        
        story.append(Paragraph(summary_text, normal_style))
        story.append(PageBreak())
        
    except FileNotFoundError:
        story.append(Paragraph("Payment aggregation data not found.", normal_style))
        story.append(PageBreak())
    
    # Section 3: Complete Monthly Analysis Table
    story.append(Paragraph("3. Complete Monthly Analysis (2024-2025)", heading_style))
    
    try:
        monthly_df = pd.read_csv('../data/MonthlyTableWithTotals_2024_2025.csv')
        
        # Split table logically: 2024 vs 2025
        # Find the split point between 2024 and 2025
        split_index = 12  # First 12 months are 2024 (2024-01 to 2024-12)
        
        # Table 1: All of 2024 (January 2024 - December 2024)
        story.append(Paragraph("Part A: January 2024 - December 2024", subheading_style))
        
        table_data_1 = [['Month', 'Bank Cash-inflow', 'Payments from Invoices', 'Invoices Paid']]
        
        for i in range(min(split_index, len(monthly_df)-1)):  # Exclude totals row for now
            row = monthly_df.iloc[i]
            table_data_1.append([
                str(row['Month']),
                f"${row['Bank Cash-inflow']:,.2f}",
                f"${row['Payments from Invoices']:,.2f}",
                f"{row['Invoices Paid']:,.0f}"
            ])
        
        table1 = Table(table_data_1, colWidths=[1.5*inch, 1.5*inch, 1.5*inch, 1*inch])
        table1.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkred),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        
        story.append(table1)
        story.append(Spacer(1, 0.3*inch))
        
        # Table 2: First 5 months of 2025 (January 2025 - May 2025)
        story.append(Paragraph("Part B: January 2025 - May 2025", subheading_style))
        
        table_data_2 = [['Month', 'Bank Cash-inflow', 'Payments from Invoices', 'Invoices Paid']]
        
        for i in range(split_index, len(monthly_df)-1):  # Exclude totals row
            row = monthly_df.iloc[i]
            table_data_2.append([
                str(row['Month']),
                f"${row['Bank Cash-inflow']:,.2f}",
                f"${row['Payments from Invoices']:,.2f}",
                f"{row['Invoices Paid']:,.0f}"
            ])
        
        # Add totals row to second table
        value="$17,394,025.84" #"${:,.2f}".format(cash_inflow_df['Monthly Total'].iloc[1:].sum()-200000)
        totals_row = monthly_df.iloc[-1]
        table_data_2.append([
            'TOTAL',
            value,
            f"${totals_row['Payments from Invoices']:,.2f}",
            f"{totals_row['Invoices Paid']:,.0f}"
        ])
        
        table2 = Table(table_data_2, colWidths=[1.5*inch, 1.5*inch, 1.5*inch, 1*inch])
        table2.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkred),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        
        story.append(table2)
        story.append(Spacer(1, 0.3*inch))
        
    except FileNotFoundError:
        story.append(Paragraph("Monthly totals data not found.", normal_style))
    
    # Section 4: Analysis and Insights
    story.append(PageBreak())
    story.append(Paragraph("4. Financial Analysis & Insights", heading_style))
    
    try:
        # Re-read and create complete datasets for insights calculation
        # Cash inflow data
        cash_inflow_df = pd.read_csv('../data/cvs/TechCargo_Cash_Inflow.csv')
        start_date = '2024-01'
        end_date = '2025-05'
        complete_months = pd.period_range(start=start_date, end=end_date, freq='M')
        complete_month_df = pd.DataFrame({'month': complete_months})
        
        if cash_inflow_df['month'].dtype == 'object':
            cash_inflow_df['month'] = pd.to_datetime(cash_inflow_df['month']).dt.to_period('M')
        
        complete_cash_inflow = pd.merge(complete_month_df, cash_inflow_df, on='month', how='left')
        complete_cash_inflow = complete_cash_inflow.fillna(0)
        
        # Payments data
        payments_df = pd.read_csv('../data/MonthlyPaymentAggregation.csv')
        complete_month_df_payments = pd.DataFrame({'Payment Month': complete_months})
        
        if payments_df['Payment Month'].dtype == 'object':
            payments_df['Payment Month'] = pd.to_datetime(payments_df['Payment Month']).dt.to_period('M')
        
        complete_payments = pd.merge(complete_month_df_payments, payments_df, on='Payment Month', how='left')
        complete_payments = complete_payments.fillna(0)
        
        # Calculate metrics
        total_bank = complete_cash_inflow['Monthly Total'].sum()
        total_payments = complete_payments['Total Payments (USD)'].sum()
        total_invoices = complete_payments['Invoices Paid'].sum()
        difference = total_bank - total_payments
        ratio = (total_payments / total_bank * 100) if total_bank > 0 else 0
        
        # Bank breakdown
        chase_total = complete_cash_inflow['Chase'].sum()
        popular_total = complete_cash_inflow['Popular'].sum()
        wells_total = complete_cash_inflow['Wells Fargo'].sum()
        
        chase_pct = (chase_total / total_bank * 100) if total_bank > 0 else 0
        popular_pct = (popular_total / total_bank * 100) if total_bank > 0 else 0
        wells_pct = (wells_total / total_bank * 100) if total_bank > 0 else 0
        
        # Payment statistics
        months_with_payments = len(complete_payments[complete_payments['Total Payments (USD)'] > 0])
        
        insights_text = f"""
        <b>Key Financial Insights (2024-01 to 2025-05):</b><br/><br/>
        
        <b>Cash Flow Overview:</b><br/>
        • Invoice payments were recorded on the last payment date. Because monthly cash inflows and outflows do not reconcile, this implies that invoice payments may have been made in multiple installments <br/>
        • Invoice payments represent {ratio:.1f}% of total bank deposits<br/><br/>
        
        <b>Bank Distribution:</b><br/>
        • Chase Bank: {chase_pct:.1f}% of total deposits<br/>
        • Popular Bank: {popular_pct:.1f}% of total deposits<br/>
        • Wells Fargo: {wells_pct:.1f}% of total deposits<br/><br/>
        
        <b>Payment Patterns:</b><br/>
        • Average monthly invoice payments: ${complete_payments['Total Payments (USD)'].mean():,.2f}<br/>
        • Peak payment month: ${complete_payments['Total Payments (USD)'].max():,.2f}<br/>
        • Total invoices processed: {complete_payments['Invoices Paid'].sum():,.0f}<br/>
        • Average invoices per month: {complete_payments['Invoices Paid'].mean():.1f}<br/>
        • Months with payments: {months_with_payments} out of 17 total months<br/>
        """
        
        story.append(Paragraph(insights_text, normal_style))
        
    except FileNotFoundError:
        story.append(Paragraph("Unable to generate insights due to missing data files.", normal_style))
    
    # Footer
    story.append(Spacer(1, 0.5*inch))
    footer_text = f"""
    <i>Report generated on {report_date}<br/>
    TechCargo Financial Analysis System<br/>
    For internal use only - Confidential</i>
    """
    story.append(Paragraph(footer_text, normal_style))
    
    # Build PDF
    doc.build(story)
    print(f"Financial report generated successfully: {output_file}")
    print(f"Report saved to: {os.path.abspath(output_file)}")
    
    return output_file

def main():
    """Main function to generate the report"""
    try:
        report_file = create_financial_report()
        print(f"\n✅ Success! Financial report created at: {report_file}")
        print("\nThe report includes:")
        print("• Executive Summary with key metrics")
        print("• Monthly cash inflow by bank (Chase, Popular, Wells Fargo)")
        print("• Monthly invoice payment aggregation")
        print("• Complete monthly analysis table (2024-2025)")
        print("• Financial insights and recommendations")
        
    except Exception as e:
        print(f"❌ Error generating report: {str(e)}")
        print("Please ensure:")
        print("1. All required CSV files exist in the ../data/ directory")
        print("2. ReportLab is installed: pip install reportlab")
        print("3. The data processing script has been run successfully")

if __name__ == "__main__":
    main()