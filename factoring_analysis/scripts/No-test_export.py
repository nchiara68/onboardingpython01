"""
Quick test script to verify the export works
Save as: factoring_analysis/scripts/test_export.py
"""

import pandas as pd
import sys
import os
from datetime import datetime

# Add the scripts directory to Python path
sys.path.append(os.path.dirname(__file__))

try:
    from factoring_analyzer import FactoringAnalyzer
except ImportError:
    print("❌ Could not import FactoringAnalyzer")
    sys.exit(1)

def test_export():
    """Test the export functionality"""
    print("🧪 TESTING EXPORT FUNCTIONALITY")
    print("="*50)
    
    # File paths
    csv_path = '../data/TecnoCargoInvoiceDataset01.csv'
    output_path = f'../output/test_export_{datetime.now().strftime("%H%M%S")}.xlsx'
    
    # Create output directory
    os.makedirs('../output', exist_ok=True)
    
    print("📁 Loading analyzer...")
    analyzer = FactoringAnalyzer(csv_path)
    
    # Create separations
    print("🔄 Creating data separations...")
    analyzer.df_paid_invoices = analyzer.df_invoice[analyzer.df_invoice['Amt. Due (USD)'] == 0].copy()
    analyzer.df_outstanding_invoices = analyzer.df_invoice[analyzer.df_invoice['Amt. Due (USD)'] > 0].copy()
    
    print(f"   📊 Invoice breakdown:")
    print(f"   📄 Total invoices: {len(analyzer.df_invoice)}")
    print(f"   📈 Paid invoices: {len(analyzer.df_paid_invoices)}")
    print(f"   ⏳ Outstanding invoices: {len(analyzer.df_outstanding_invoices)}")
    print(f"   💳 Credit memos: {len(analyzer.df_credit_memo)}")
    
    # Test Excel creation
    print("\n📊 Testing Excel creation...")
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Always create summary
            summary_data = {
                'Category': ['Total Records', 'Invoices', 'Paid', 'Outstanding', 'Credits'],
                'Count': [len(analyzer.df_full), len(analyzer.df_invoice), 
                         len(analyzer.df_paid_invoices), len(analyzer.df_outstanding_invoices), 
                         len(analyzer.df_credit_memo)]
            }
            pd.DataFrame(summary_data).to_excel(writer, sheet_name='Test_Summary', index=False)
            
            # Add raw data
            analyzer.df_invoice.to_excel(writer, sheet_name='Test_Raw_Data', index=False)
            
        print(f"✅ Test Excel file created: {output_path}")
        
        # Verify file exists and has content
        test_df = pd.read_excel(output_path, sheet_name='Test_Summary')
        print(f"✅ File verification successful - {len(test_df)} rows in summary")
        
        return True
        
    except Exception as e:
        print(f"❌ Excel creation failed: {e}")
        return False

if __name__ == "__main__":
    success = test_export()
    if success:
        print("\n🎉 Test passed! You can now run the full export script.")
    else:
        print("\n❌ Test failed. Please check the error messages above.")
        