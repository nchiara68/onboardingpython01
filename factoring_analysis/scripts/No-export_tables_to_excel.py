# ============================================================================
# NOTEBOOK CELL: Export Factoring Analysis Tables to Excel
# Use this in a Jupyter notebook in: factoring_analysis/notebooks/
# ============================================================================

import pandas as pd
import sys
import os
from datetime import datetime

# Add the scripts directory to Python path
sys.path.append('../scripts')
from factoring_analyzer import FactoringAnalyzer

def quick_export_to_excel():
    """Quick function to export all tables to Excel from notebook"""
    
    # File paths
    csv_path = '../data/TecnoCargoInvoiceDataset01.csv'
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = f'../output/factoring_tables_{timestamp}.xlsx'
    
    # Create output directory
    os.makedirs('../output', exist_ok=True)
    
    print("🏢 Loading data and creating analysis...")
    analyzer = FactoringAnalyzer(csv_path)
    
    # Create Excel file
    print("📊 Exporting to Excel...")
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        
        # PAID INVOICES TABLES
        if len(analyzer.df_paid_invoices) > 0:
            # 1.1: Count by aging
            paid_count = analyzer.df_paid_invoices.groupby('Aging_Bucket')['Number'].count().reset_index()
            paid_count.columns = ['Aging Bucket', 'Count']
            paid_count.to_excel(writer, sheet_name='1.1_Paid_Count', index=False)
            
            # 1.2: Amount by aging  
            paid_amount = analyzer.df_paid_invoices.groupby('Aging_Bucket')['Amount (USD)'].sum().reset_index()
            paid_amount.columns = ['Aging Bucket', 'Total Amount']
            paid_amount.to_excel(writer, sheet_name='1.2_Paid_Amount', index=False)
            
            # 1.3: Percentages
            total_paid = len(analyzer.df_paid_invoices)
            total_amount = analyzer.df_paid_invoices['Amount (USD)'].sum()
            paid_pct = analyzer.df_paid_invoices.groupby('Aging_Bucket').agg({
                'Number': 'count', 'Amount (USD)': 'sum'}).reset_index()
            paid_pct['Count %'] = (paid_pct['Number'] / total_paid * 100).round(2)
            paid_pct['Amount %'] = (paid_pct['Amount (USD)'] / total_amount * 100).round(2)
            paid_pct[['Aging_Bucket', 'Count %', 'Amount %']].to_excel(writer, sheet_name='1.3_Paid_Percentages', index=False)
            
            # 1.4: Customer analysis
            try:
                customer_pivot = pd.crosstab(analyzer.df_paid_invoices['Applied to'], 
                                           analyzer.df_paid_invoices['Aging_Bucket'], 
                                           values=analyzer.df_paid_invoices['Amount (USD)'], 
                                           aggfunc='sum').fillna(0)
                customer_pivot.to_excel(writer, sheet_name='1.4_Paid_Customer_Matrix')
            except:
                pass
        
        # OUTSTANDING INVOICES TABLES
        if len(analyzer.df_outstanding_invoices) > 0:
            # 2.1: Count by aging
            out_count = analyzer.df_outstanding_invoices.groupby('Aging_Bucket')['Number'].count().reset_index()
            out_count.columns = ['Aging Bucket', 'Count']
            out_count.to_excel(writer, sheet_name='2.1_Outstanding_Count', index=False)
            
            # 2.2: Amount by aging
            out_amount = analyzer.df_outstanding_invoices.groupby('Aging_Bucket')['Amt. Due (USD)'].sum().reset_index()
            out_amount.columns = ['Aging Bucket', 'Amount Due']
            out_amount.to_excel(writer, sheet_name='2.2_Outstanding_Amount', index=False)
            
            # 2.3: Percentages
            total_out = len(analyzer.df_outstanding_invoices)
            total_due = analyzer.df_outstanding_invoices['Amt. Due (USD)'].sum()
            out_pct = analyzer.df_outstanding_invoices.groupby('Aging_Bucket').agg({
                'Number': 'count', 'Amt. Due (USD)': 'sum'}).reset_index()
            out_pct['Count %'] = (out_pct['Number'] / total_out * 100).round(2)
            out_pct['Amount %'] = (out_pct['Amt. Due (USD)'] / total_due * 100).round(2)
            out_pct[['Aging_Bucket', 'Count %', 'Amount %']].to_excel(writer, sheet_name='2.3_Outstanding_Percentages', index=False)
            
            # 2.4: Customer analysis
            try:
                customer_out_pivot = pd.crosstab(analyzer.df_outstanding_invoices['Applied to'], 
                                               analyzer.df_outstanding_invoices['Aging_Bucket'], 
                                               values=analyzer.df_outstanding_invoices['Amt. Due (USD)'], 
                                               aggfunc='sum').fillna(0)
                customer_out_pivot.to_excel(writer, sheet_name='2.4_Outstanding_Customer_Matrix')
            except:
                pass
        
        # CREDIT MEMO TABLES
        if len(analyzer.df_credit_memo) > 0:
            # 3.1: Credit by customer
            credit_summary = analyzer.df_credit_memo.groupby('Applied to')['Amount (USD)'].sum().reset_index()
            credit_summary.columns = ['Customer', 'Credit Amount']
            credit_summary = credit_summary.sort_values('Credit Amount', ascending=False)
            credit_summary.to_excel(writer, sheet_name='3.1_Credit_by_Customer', index=False)
            
            # 3.2: Credit summary stats
            credit_stats = pd.DataFrame({
                'Metric': ['Total Credits', 'Total Amount', 'Average', 'Max', 'Customers'],
                'Value': [len(analyzer.df_credit_memo), 
                         analyzer.df_credit_memo['Amount (USD)'].sum(),
                         analyzer.df_credit_memo['Amount (USD)'].mean(),
                         analyzer.df_credit_memo['Amount (USD)'].max(),
                         analyzer.df_credit_memo['Applied to'].nunique()]
            })
            credit_stats.to_excel(writer, sheet_name='3.2_Credit_Summary', index=False)
        
        # RAW DATA
        analyzer.df_paid_invoices.to_excel(writer, sheet_name='RAW_Paid', index=False)
        analyzer.df_outstanding_invoices.to_excel(writer, sheet_name='RAW_Outstanding', index=False)
        if len(analyzer.df_credit_memo) > 0:
            analyzer.df_credit_memo.to_excel(writer, sheet_name='RAW_Credits', index=False)
    
    print(f"✅ Excel file created: {output_path}")
    print(f"📊 Worksheets: {11 + (3 if len(analyzer.df_credit_memo) > 0 else 1)} total")
    
    return output_path

# Run the export
output_file = quick_export_to_excel()