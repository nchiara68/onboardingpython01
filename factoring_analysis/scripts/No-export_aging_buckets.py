"""
📊 Aging Buckets Export Routine
Exports invoice aging analysis to separate Excel worksheets
Save this as: scripts/export_aging_buckets.py
"""

import pandas as pd
import numpy as np
from datetime import datetime
import os


def export_aging_buckets(csv_path='../data/TecnoCargoInvoiceDataset01.csv', 
                        output_path='../data/Invoice_Aging_Analysis.xlsx',
                        encoding='latin1',
                        include_summary=True,
                        include_credit_memos=False):
    """
    Export invoice aging buckets to separate Excel worksheets
    
    Parameters:
    csv_path (str): Path to the input CSV file
    output_path (str): Path for the output Excel file
    encoding (str): CSV file encoding
    include_summary (bool): Include summary worksheet
    include_credit_memos (bool): Include credit memos in separate sheet
    
    Returns:
    dict: Summary of exported data
    """
    
    print("📊 EXPORTING AGING BUCKETS TO EXCEL")
    print("="*50)
    
    # Load and prepare data
    print("1. Loading data...")
    try:
        df = pd.read_csv(csv_path, encoding=encoding)
        print(f"   ✅ Loaded {len(df)} records")
    except Exception as e:
        print(f"   ❌ Error loading data: {e}")
        return None
    
    # Clean monetary columns
    print("2. Cleaning monetary columns...")
    monetary_columns = ['Amount (USD)', 'Amt. Paid (USD)', 'Amt. Due (USD)']
    for col in monetary_columns:
        if col in df.columns:
            df[col] = (df[col]
                      .astype(str)
                      .str.replace('$', '', regex=False)
                      .str.replace(',', '', regex=False)
                      .str.replace(' ', '', regex=False)
                      .str.replace('(', '-', regex=False)
                      .str.replace(')', '', regex=False)
                      .replace('nan', '0')
                      .replace('', '0'))
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    # Convert date columns
    print("3. Converting date columns...")
    date_columns = ['Transaction Date', 'Due Date', 'Last Payment Date']
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    # Split by type
    df_invoices = df[df['Type'] == 'Invoice'].copy()
    df_credit_memos = df[df['Type'] == 'Credit Memo'].copy()
    
    print(f"   📄 Invoice records: {len(df_invoices)}")
    print(f"   📝 Credit memo records: {len(df_credit_memos)}")
    
    # Calculate aging for invoices
    print("4. Calculating aging buckets...")
    today = datetime.now()
    df_invoices['Days_Past_Due'] = (today - df_invoices['Due Date']).dt.days
    df_invoices['Days_Since_Transaction'] = (today - df_invoices['Transaction Date']).dt.days
    
    # Create aging buckets
    df_invoices['Aging_Bucket'] = pd.cut(
        df_invoices['Days_Past_Due'],
        bins=[-float('inf'), 0, 30, 60, 90, float('inf')],
        labels=['Current', '1-30 Days', '31-60 Days', '61-90 Days', '90+ Days']
    )
    
    # Split into aging buckets
    aging_buckets = {}
    bucket_names = ['Current', '1-30 Days', '31-60 Days', '61-90 Days', '90+ Days']
    
    for bucket in bucket_names:
        bucket_data = df_invoices[df_invoices['Aging_Bucket'] == bucket].copy()
        aging_buckets[bucket] = bucket_data
        print(f"   {bucket}: {len(bucket_data)} invoices")
    
    # Create Excel file with multiple worksheets
    print("5. Creating Excel file...")
    
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            
            # Export each aging bucket to separate worksheet
            for bucket_name, bucket_data in aging_buckets.items():
                if len(bucket_data) > 0:
                    # Clean worksheet name (Excel limitations)
                    sheet_name = bucket_name.replace('+', '_Plus').replace('-', '_')
                    if len(sheet_name) > 31:  # Excel sheet name limit
                        sheet_name = sheet_name[:31]
                    
                    # Select columns to export
                    export_columns = [
                        'Number', 'Transaction Date', 'Applied to', 'Amount (USD)', 
                        'Due Date', 'Status', 'Last Payment Date', 'Amt. Paid (USD)', 
                        'Amt. Due (USD)', 'Days_Past_Due', 'Days_Since_Transaction'
                    ]
                    
                    # Filter to available columns
                    available_columns = [col for col in export_columns if col in bucket_data.columns]
                    export_data = bucket_data[available_columns].copy()
                    
                    # Sort by Days_Past_Due (worst first)
                    if 'Days_Past_Due' in export_data.columns:
                        export_data = export_data.sort_values('Days_Past_Due', ascending=False)
                    
                    # Export to worksheet
                    export_data.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"   ✅ Exported {len(export_data)} records to '{sheet_name}' worksheet")
                else:
                    print(f"   ⚠️ Skipped '{bucket_name}' - no records")
            
            # Create summary worksheet
            if include_summary:
                print("6. Creating summary worksheet...")
                summary_data = []
                
                for bucket_name in bucket_names:
                    bucket_data = aging_buckets[bucket_name]
                    if len(bucket_data) > 0:
                        summary_data.append({
                            'Aging_Bucket': bucket_name,
                            'Invoice_Count': len(bucket_data),
                            'Total_Amount_USD': bucket_data['Amount (USD)'].sum(),
                            'Total_Due_USD': bucket_data['Amt. Due (USD)'].sum(),
                            'Total_Paid_USD': bucket_data['Amt. Paid (USD)'].sum(),
                            'Avg_Days_Past_Due': bucket_data['Days_Past_Due'].mean(),
                            'Max_Days_Past_Due': bucket_data['Days_Past_Due'].max(),
                            'Percentage_of_Total': (len(bucket_data) / len(df_invoices)) * 100
                        })
                
                summary_df = pd.DataFrame(summary_data)
                
                # Add totals row
                totals_row = {
                    'Aging_Bucket': 'TOTAL',
                    'Invoice_Count': summary_df['Invoice_Count'].sum(),
                    'Total_Amount_USD': summary_df['Total_Amount_USD'].sum(),
                    'Total_Due_USD': summary_df['Total_Due_USD'].sum(),
                    'Total_Paid_USD': summary_df['Total_Paid_USD'].sum(),
                    'Avg_Days_Past_Due': df_invoices['Days_Past_Due'].mean(),
                    'Max_Days_Past_Due': df_invoices['Days_Past_Due'].max(),
                    'Percentage_of_Total': 100.0
                }
                
                summary_df = pd.concat([summary_df, pd.DataFrame([totals_row])], ignore_index=True)
                
                # Round numeric columns
                numeric_columns = ['Total_Amount_USD', 'Total_Due_USD', 'Total_Paid_USD', 
                                 'Avg_Days_Past_Due', 'Percentage_of_Total']
                for col in numeric_columns:
                    summary_df[col] = summary_df[col].round(2)
                
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
                print(f"   ✅ Created summary worksheet")
            
            # Add credit memos if requested
            if include_credit_memos and len(df_credit_memos) > 0:
                print("7. Adding credit memos worksheet...")
                credit_export_columns = [
                    'Number', 'Transaction Date', 'Applied to', 'Amount (USD)', 
                    'Status', 'Last Payment Date', 'Amt. Paid (USD)', 'Amt. Due (USD)'
                ]
                available_credit_columns = [col for col in credit_export_columns if col in df_credit_memos.columns]
                credit_export = df_credit_memos[available_credit_columns].copy()
                
                credit_export.to_excel(writer, sheet_name='Credit_Memos', index=False)
                print(f"   ✅ Exported {len(credit_export)} credit memos")
        
        print(f"\n🎉 SUCCESS! Excel file created: {output_path}")
        
        # Return summary
        result_summary = {
            'output_file': output_path,
            'total_invoices': len(df_invoices),
            'total_credit_memos': len(df_credit_memos),
            'aging_breakdown': {bucket: len(data) for bucket, data in aging_buckets.items()},
            'export_date': today.strftime('%Y-%m-%d %H:%M:%S')
        }
        
        print(f"\n📋 EXPORT SUMMARY:")
        print(f"   Output file: {os.path.basename(output_path)}")
        print(f"   Total invoices processed: {len(df_invoices)}")
        print(f"   Worksheets created: {len([b for b in aging_buckets.values() if len(b) > 0]) + (1 if include_summary else 0) + (1 if include_credit_memos and len(df_credit_memos) > 0 else 0)}")
        print(f"   Export completed: {today.strftime('%Y-%m-%d %H:%M:%S')}")
        
        return result_summary
        
    except Exception as e:
        print(f"❌ Error creating Excel file: {e}")
        return None


def export_aging_quick(csv_path='../data/TecnoCargoInvoiceDataset01.csv'):
    """
    Quick export with default settings
    """
    return export_aging_buckets(csv_path)


def export_aging_detailed(csv_path='../data/TecnoCargoInvoiceDataset01.csv',
                         output_filename='Detailed_Aging_Analysis.xlsx'):
    """
    Detailed export with all options enabled
    """
    output_path = f'../data/{output_filename}'
    return export_aging_buckets(
        csv_path=csv_path,
        output_path=output_path,
        include_summary=True,
        include_credit_memos=True
    )


def export_custom_aging_buckets(csv_path, aging_ranges, output_path):
    """
    Export with custom aging bucket ranges
    
    Parameters:
    csv_path (str): Input CSV path
    aging_ranges (list): List of tuples defining custom ranges [(0, 30), (31, 60), etc.]
    output_path (str): Output Excel path
    """
    # Implementation for custom aging ranges
    # This would allow users to define their own aging categories
    pass


if __name__ == "__main__":
    # Run export if script is executed directly
    print("🚀 Running aging buckets export...")
    result = export_aging_quick()
    
    if result:
        print("✅ Export completed successfully!")
        print(f"📁 File saved: {result['output_file']}")
    else:
        print("❌ Export failed!")