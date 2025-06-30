"""
📊 Data Validation Module for Factoring Analysis
Save this file as: scripts/data_validation.py
"""

import pandas as pd
import numpy as np

def quick_test(csv_path='../data/TecnoCargoInvoiceDataset01.csv', encoding='latin1'):
    """
    Run a quick test of invoice data loading and validation
    
    Parameters:
    csv_path (str): Path to the CSV file
    encoding (str): File encoding (default: 'latin1')
    
    Returns:
    dict: Dictionary containing validation results and cleaned data
    """
    
    print("🔍 TESTING DATA LOADING & VALIDATION")
    print("="*50)
    
    # Load raw data with encoding handling
    print("1. Loading raw data...")
    try:
        df_raw = pd.read_csv(csv_path, encoding=encoding)
        print(f"   ✅ Loaded {len(df_raw)} records with {len(df_raw.columns)} columns (encoding: {encoding})")
    except UnicodeDecodeError:
        print(f"   ⚠️ Failed with {encoding}, trying utf-8...")
        try:
            df_raw = pd.read_csv(csv_path, encoding='utf-8')
            print(f"   ✅ Loaded {len(df_raw)} records with {len(df_raw.columns)} columns (encoding: utf-8)")
        except:
            print(f"   ⚠️ Failed with utf-8, using default...")
            df_raw = pd.read_csv(csv_path)
            print(f"   ✅ Loaded {len(df_raw)} records with {len(df_raw.columns)} columns (default encoding)")
    except Exception as e:
        print(f"   ❌ ERROR: {e}")
        return None
    
    # Show column names and types
    print("\n2. Column information:")
    for i, col in enumerate(df_raw.columns):
        print(f"   {i+1:2d}. {col} ({df_raw[col].dtype})")
    
    # Check Type column
    print(f"\n3. Type distribution:")
    type_counts = df_raw['Type'].value_counts()
    for type_name, count in type_counts.items():
        print(f"   {type_name}: {count} records")
    
    # Show sample monetary values BEFORE cleaning
    print(f"\n4. Sample monetary values (BEFORE cleaning):")
    monetary_cols = ['Amount (USD)', 'Amt. Paid (USD)', 'Amt. Due (USD)']
    for col in monetary_cols:
        if col in df_raw.columns:
            sample_vals = df_raw[col].head(3).tolist()
            print(f"   {col}: {sample_vals}")
            print(f"   Data type: {df_raw[col].dtype}")
    
    print(f"\n5. Testing monetary column conversion...")
    df_test = df_raw.copy()
    
    # Clean monetary columns
    for col in monetary_cols:
        if col in df_test.columns:
            print(f"   Converting {col}...")
            
            # Show original values
            original_sample = df_test[col].head(3).tolist()
            print(f"     Before: {original_sample}")
            
            # Apply conversion
            df_test[col] = (df_test[col]
                           .astype(str)
                           .str.replace('$', '', regex=False)
                           .str.replace(',', '', regex=False)
                           .str.replace(' ', '', regex=False)
                           .str.replace('(', '-', regex=False)
                           .str.replace(')', '', regex=False)
                           .replace('nan', '0')
                           .replace('', '0'))
            
            df_test[col] = pd.to_numeric(df_test[col], errors='coerce').fillna(0)
            
            # Show converted values
            converted_sample = df_test[col].head(3).tolist()
            print(f"     After:  {[f'${x:,.2f}' for x in converted_sample]}")
    
    # Test dataset splitting
    print(f"\n6. Testing dataset splitting...")
    df_invoices = df_test[df_test['Type'] == 'Invoice']
    df_credits = df_test[df_test['Type'] == 'Credit Memo']
    
    print(f"   📄 Invoice records: {len(df_invoices)}")
    print(f"   📝 Credit Memo records: {len(df_credits)}")
    
    # Summary statistics
    print(f"\n7. Summary statistics (after conversion):")
    for col in monetary_cols:
        if col in df_test.columns:
            total = df_test[col].sum()
            mean_val = df_test[col].mean()
            print(f"   {col}:")
            print(f"     Total: ${total:,.2f}")
            print(f"     Average: ${mean_val:,.2f}")
            print(f"     Non-zero records: {(df_test[col] > 0).sum()}")
    
    print(f"\n✅ Data validation test completed!")
    print(f"📊 Ready to proceed with full analysis")
    
    return {
        'raw_data': df_raw,
        'cleaned_data': df_test,
        'invoice_data': df_invoices,
        'credit_data': df_credits,
        'summary': {
            'total_records': len(df_raw),
            'invoice_count': len(df_invoices),
            'credit_count': len(df_credits),
            'monetary_columns': monetary_cols
        }
    }


def validate_file_structure(base_path='..'):
    """
    Validate that the project file structure is correct
    
    Parameters:
    base_path (str): Base path to the project folder
    
    Returns:
    bool: True if structure is valid, False otherwise
    """
    
    import os
    
    print("📁 VALIDATING PROJECT STRUCTURE")
    print("="*40)
    
    required_folders = ['data', 'notebooks', 'scripts', 'reports']
    required_files = ['data/TecnoCargoInvoiceDataset01.csv']
    
    # Check folders
    for folder in required_folders:
        path = os.path.join(base_path, folder)
        if os.path.exists(path):
            print(f"   ✅ {folder}/ folder exists")
        else:
            print(f"   ❌ {folder}/ folder missing")
            return False
    
    # Check files
    for file in required_files:
        path = os.path.join(base_path, file)
        if os.path.exists(path):
            print(f"   ✅ {file} exists")
        else:
            print(f"   ❌ {file} missing")
            return False
    
    print(f"\n✅ Project structure is valid!")
    return True


def check_dependencies():
    """
    Check if all required Python packages are installed
    
    Returns:
    bool: True if all dependencies are available, False otherwise
    """
    
    print("📦 CHECKING DEPENDENCIES")
    print("="*30)
    
    required_packages = {
        'pandas': 'Data manipulation',
        'numpy': 'Numerical computing',
        'matplotlib': 'Plotting',
        'seaborn': 'Statistical visualization',
        'datetime': 'Date/time handling'
    }
    
    missing_packages = []
    
    for package, description in required_packages.items():
        try:
            __import__(package)
            print(f"   ✅ {package} - {description}")
        except ImportError:
            print(f"   ❌ {package} - {description} (MISSING)")
            missing_packages.append(package)
    
    if missing_packages:
        print(f"\n❌ Missing packages: {', '.join(missing_packages)}")
        print(f"💡 Install with: pip install {' '.join(missing_packages)}")
        return False
    else:
        print(f"\n✅ All dependencies are installed!")
        return True


def full_environment_check(csv_path='../data/TecnoCargoInvoiceDataset01.csv'):
    """
    Complete environment and data validation check
    
    Parameters:
    csv_path (str): Path to the CSV file
    
    Returns:
    dict: Complete validation results
    """
    
    print("🚀 FULL ENVIRONMENT CHECK")
    print("="*60)
    
    # Check dependencies
    deps_ok = check_dependencies()
    
    print()
    
    # Check file structure
    structure_ok = validate_file_structure()
    
    print()
    
    # Test data loading
    if deps_ok and structure_ok:
        data_results = quick_test(csv_path)
        
        if data_results:
            print(f"\n🎉 ENVIRONMENT CHECK PASSED!")
            print(f"🚀 Ready to run full factoring analysis!")
            return {
                'dependencies': True,
                'structure': True,
                'data': data_results,
                'status': 'ready'
            }
        else:
            print(f"\n❌ Data validation failed")
            return {'status': 'data_error'}
    else:
        print(f"\n❌ Environment setup incomplete")
        return {'status': 'setup_error'}


# Convenience function for quick access
def test():
    """Quick alias for the most common test"""
    return quick_test()


if __name__ == "__main__":
    # Run full check if script is executed directly
    results = full_environment_check()
    print(f"\nFinal status: {results.get('status', 'unknown')}")