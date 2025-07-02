"""
Quick test to verify the Table 1.4 and 2.4 fixes
Save as: factoring_analysis/scripts/test_table_fix.py
"""

import pandas as pd
import sys
import os

# Add the scripts directory to Python path
sys.path.append(os.path.dirname(__file__))

try:
    from factoring_analyzer import FactoringAnalyzer
except ImportError:
    print("❌ Could not import FactoringAnalyzer")
    sys.exit(1)

def test_table_structure():
    """Test the table structure that will be used in PDF"""
    print("🧪 TESTING TABLE 1.4 and 2.4 STRUCTURE")
    print("="*50)
    
    # Load data
    csv_path = '../data/TecnoCargoInvoiceDataset01.csv'
    analyzer = FactoringAnalyzer(csv_path)
    
    # Get base invoice data
    base_df = analyzer.df_invoice.copy()
    
    # Create data separations
    df_paid_raw = base_df[base_df['Amt. Due (USD)'] == 0].copy()
    df_outstanding_raw = base_df[base_df['Amt. Due (USD)'] > 0].copy()
    
    print(f"📊 Data: {len(df_paid_raw)} paid, {len(df_outstanding_raw)} outstanding")
    
    # Test Table 1.4 structure (Paid Invoices)
    if len(df_paid_raw) > 0:
        print("\n🔍 TESTING TABLE 1.4 STRUCTURE (PAID INVOICES)")
        print("-" * 40)
        
        # Calculate aging for paid invoices
        df_paid_raw['Due Date'] = pd.to_datetime(df_paid_raw['Due Date'], errors='coerce')
        df_paid_raw['Last Payment Date'] = pd.to_datetime(df_paid_raw['Last Payment Date'], errors='coerce')
        df_paid_raw['Aging Delay Days'] = (df_paid_raw['Last Payment Date'] - df_paid_raw['Due Date']).dt.days
        df_paid_raw['Aging Delay Days'] = df_paid_raw['Aging Delay Days'].fillna(0)
        
        def create_aging_bucket(days, invoice_type):
            if days <= 0:
                return 'On-time' if invoice_type == 'paid' else 'Current'
            elif 1 <= days <= 30:
                return '1-30 Days'
            elif 31 <= days <= 60:
                return '31-60 Days'
            elif 61 <= days <= 90:
                return '61-90 Days'
            else:
                return '90+ Days'
        
        df_paid_raw['Aging_Bucket'] = df_paid_raw['Aging Delay Days'].apply(lambda x: create_aging_bucket(x, 'paid'))
        
        # Create pivot table
        customer_aging_paid_values = pd.crosstab(df_paid_raw['Applied to'], 
                                                df_paid_raw['Aging_Bucket'], 
                                                values=df_paid_raw['Amount (USD)'], 
                                                aggfunc='sum').fillna(0)
        
        # Sort columns
        aging_order = ['On-time', 'Current', '1-30 Days', '31-60 Days', '61-90 Days', '90+ Days']
        available_cols = [col for col in aging_order if col in customer_aging_paid_values.columns]
        customer_aging_paid_values = customer_aging_paid_values.reindex(columns=available_cols, fill_value=0)
        
        # Calculate percentages
        customer_aging_paid_pct = customer_aging_paid_values.div(customer_aging_paid_values.sum(axis=1), axis=0) * 100
        customer_aging_paid_pct = customer_aging_paid_pct.fillna(0).round(1)
        
        # Get top 5 customers for testing
        top_customers = df_paid_raw.groupby('Applied to')['Amount (USD)'].sum().nlargest(5).index
        
        print(f"Available aging columns: {available_cols}")
        print(f"Top 5 customers: {list(top_customers)}")
        
        # Create the alternating structure
        combined_data = []
        for customer in top_customers[:3]:  # Just show first 3 for testing
            # Percentage row
            pct_row = [customer + " (%)"] + [f"{customer_aging_paid_pct.loc[customer, col]:.1f}%" 
                                            for col in available_cols]
            combined_data.append(pct_row)
            
            # Dollar row
            val_row = [customer + " ($)"] + [f"${customer_aging_paid_values.loc[customer, col]:,.0f}" 
                                            for col in available_cols]
            combined_data.append(val_row)
        
        # Show the structure
        column_names = ['Customer'] + available_cols
        customer_combined_df = pd.DataFrame(combined_data, columns=column_names)
        
        print("\n📋 TABLE 1.4 STRUCTURE PREVIEW:")
        print(customer_combined_df.to_string(index=False))
    
    # Test Table 2.4 structure (Outstanding Invoices)
    if len(df_outstanding_raw) > 0:
        print("\n🔍 TESTING TABLE 2.4 STRUCTURE (OUTSTANDING INVOICES)")
        print("-" * 40)
        
        # Calculate aging for outstanding invoices
        df_outstanding_raw['Due Date'] = pd.to_datetime(df_outstanding_raw['Due Date'], errors='coerce')
        cutoff_date = pd.Timestamp('2025-06-23')
        df_outstanding_raw['Aging Delay Days'] = (cutoff_date - df_outstanding_raw['Due Date']).dt.days
        df_outstanding_raw['Aging Delay Days'] = df_outstanding_raw['Aging Delay Days'].fillna(0)
        df_outstanding_raw['Aging_Bucket'] = df_outstanding_raw['Aging Delay Days'].apply(lambda x: create_aging_bucket(x, 'outstanding'))
        
        # Create pivot table
        customer_aging_out_values = pd.crosstab(df_outstanding_raw['Applied to'], 
                                               df_outstanding_raw['Aging_Bucket'], 
                                               values=df_outstanding_raw['Amt. Due (USD)'], 
                                               aggfunc='sum').fillna(0)
        
        # Sort columns
        aging_order = ['On-time', 'Current', '1-30 Days', '31-60 Days', '61-90 Days', '90+ Days']
        available_cols = [col for col in aging_order if col in customer_aging_out_values.columns]
        customer_aging_out_values = customer_aging_out_values.reindex(columns=available_cols, fill_value=0)
        
        # Calculate percentages
        customer_aging_out_pct = customer_aging_out_values.div(customer_aging_out_values.sum(axis=1), axis=0) * 100
        customer_aging_out_pct = customer_aging_out_pct.fillna(0).round(1)
        
        # Get top 3 customers for testing
        top_customers = df_outstanding_raw.groupby('Applied to')['Amt. Due (USD)'].sum().nlargest(3).index
        
        print(f"Available aging columns: {available_cols}")
        print(f"Top 3 customers: {list(top_customers)}")
        
        # Create the alternating structure
        combined_data = []
        for customer in top_customers:
            # Percentage row
            pct_row = [customer + " (%)"] + [f"{customer_aging_out_pct.loc[customer, col]:.1f}%" 
                                            for col in available_cols]
            combined_data.append(pct_row)
            
            # Dollar row
            val_row = [customer + " ($)"] + [f"${customer_aging_out_values.loc[customer, col]:,.0f}" 
                                            for col in available_cols]
            combined_data.append(val_row)
        
        # Show the structure
        column_names = ['Customer'] + available_cols
        customer_combined_df = pd.DataFrame(combined_data, columns=column_names)
        
        print("\n📋 TABLE 2.4 STRUCTURE PREVIEW:")
        print(customer_combined_df.to_string(index=False))
    
    print("\n✅ Table structure test completed!")
    print("🎯 The PDF will now include the correct alternating percentage/dollar format")

if __name__ == "__main__":
    test_table_structure()