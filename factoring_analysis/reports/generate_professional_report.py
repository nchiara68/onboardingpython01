"""
Integration Script for Professional WeasyPrint Reports
File: reports/generate_professional_report.py

This script integrates the new WeasyPrint reporting system with your existing
factoring analysis code. It generates both old and new format reports for comparison.
"""

import sys
import os
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

# Add project root to Python path
project_root = Path(__file__).parent.parent
sys.path.append(str(project_root))

def setup_environment():
    """Setup and validate the environment"""
    print("🔧 SETTING UP ENVIRONMENT")
    print("="*50)
    
    # Check if required directories exist
    data_dir = project_root / 'data'
    output_dir = project_root / 'output'
    reports_dir = project_root / 'reports'
    
    print(f"📁 Project root: {project_root}")
    print(f"📁 Data directory: {data_dir}")
    print(f"📁 Reports directory: {reports_dir}")
    
    # Create output directories if they don't exist
    (output_dir / 'charts').mkdir(parents=True, exist_ok=True)
    (output_dir / 'reports').mkdir(parents=True, exist_ok=True)
    
    print(f"✅ Output directories created: {output_dir}")
    
    # Check if data file exists
    data_file = data_dir / 'TecnoCargoInvoiceDataset01.csv'
    if not data_file.exists():
        print(f"❌ Data file not found: {data_file}")
        print("Please ensure the CSV file is in the data directory.")
        return False
    
    print(f"✅ Data file found: {data_file}")
    
    # Check required packages
    missing_packages = []
    
    try:
        import pandas as pd
        print("✅ pandas available")
    except ImportError:
        missing_packages.append('pandas')
    
    try:
        import matplotlib.pyplot as plt
        print("✅ matplotlib available")
    except ImportError:
        missing_packages.append('matplotlib')
    
    try:
        import weasyprint
        from jinja2 import Template
        print("✅ weasyprint and jinja2 available")
    except ImportError:
        missing_packages.extend(['weasyprint', 'jinja2'])
    
    if missing_packages:
        print(f"❌ Missing packages: {missing_packages}")
        print("Please install with: pip install " + " ".join(missing_packages))
        return False
    
    print("✅ All required packages available")
    return True

def run_factoring_analysis():
    """Run the factoring analysis using existing code"""
    print("\n🏢 RUNNING FACTORING ANALYSIS")
    print("="*50)
    
    try:
        # Import your existing analyzer
        from reports.routine_supportfor_report_v01 import CorrectedFactoringAnalyzer
        
        # Initialize with correct path
        data_file = project_root / 'data' / 'TecnoCargoInvoiceDataset01.csv'
        print(f"📂 Loading data from: {data_file}")
        
        analyzer = CorrectedFactoringAnalyzer(str(data_file), encoding='latin1')
        print("✅ Analyzer initialized successfully")
        
        # Run the complete analysis
        print("📊 Running complete analysis...")
        results = analyzer.run_full_analysis()
        print("✅ Analysis completed successfully")
        
        return analyzer, results
        
    except Exception as e:
        print(f"❌ Error in analysis: {e}")
        import traceback
        traceback.print_exc()
        return None, None

def generate_original_report(analyzer, results):
    """Generate the original matplotlib-based report"""
    print("\n📊 GENERATING ORIGINAL MATPLOTLIB REPORT")
    print("="*50)
    
    try:
        # Try to import and use your original report generator
        from reports.report_v01 import create_complete_factoring_report
        
        output_file = 'TecnoCargo_Original_Report.pdf'
        report_path = create_complete_factoring_report(
            analyzer=analyzer,
            results=results,
            filename=output_file
        )
        
        print(f"✅ Original report generated: {report_path}")
        return report_path
        
    except ImportError:
        print("⚠️ Original report generator not found (report_v01.py)")
        print("   This is normal if you haven't created it yet.")
        return None
    except Exception as e:
        print(f"❌ Error generating original report: {e}")
        return None

def generate_professional_report(analyzer, results):
    """Generate the new professional WeasyPrint report"""
    print("\n🎨 GENERATING PROFESSIONAL WEASYPRINT REPORT")
    print("="*50)
    
    try:
        from reports.weasyprint_report import generate_professional_factoring_report
        
        output_file = 'TecnoCargo_Professional_Report.pdf'
        report_path = generate_professional_factoring_report(
            analyzer=analyzer,
            results=results,
            filename=output_file
        )
        
        print(f"✅ Professional report generated: {report_path}")
        return report_path
        
    except Exception as e:
        print(f"❌ Error generating professional report: {e}")
        import traceback
        traceback.print_exc()
        return None

def compare_reports(original_path, professional_path):
    """Compare the two report outputs"""
    print("\n📋 REPORT COMPARISON")
    print("="*50)
    
    if original_path and professional_path:
        print("✅ Both reports generated successfully!")
        print(f"📊 Original (matplotlib): {original_path}")
        print(f"🎨 Professional (WeasyPrint): {professional_path}")
        print("\nKey differences:")
        print("• Professional report has modern CSS styling")
        print("• Better chart integration and layout")
        print("• Enhanced table formatting")
        print("• Improved typography and spacing")
        print("• Color-coded risk indicators")
        print("• Professional page headers/footers")
        
    elif professional_path:
        print("✅ Professional report generated successfully!")
        print(f"🎨 Professional (WeasyPrint): {professional_path}")
        print("⚠️ Original report not available for comparison")
        
    elif original_path:
        print("✅ Original report generated successfully!")
        print(f"📊 Original (matplotlib): {original_path}")
        print("❌ Professional report failed to generate")
        
    else:
        print("❌ Both reports failed to generate")

def main():
    """Main execution function"""
    print("🚀 PROFESSIONAL FACTORING REPORT GENERATOR")
    print("="*60)
    print("Integration script for WeasyPrint reporting system")
    print("="*60)
    
    # Step 1: Setup environment
    if not setup_environment():
        print("❌ Environment setup failed. Please fix the issues above.")
        return
    
    # Step 2: Run analysis
    analyzer, results = run_factoring_analysis()
    if analyzer is None or results is None:
        print("❌ Analysis failed. Cannot generate reports.")
        return
    
    # Step 3: Generate reports
    original_path = generate_original_report(analyzer, results)
    professional_path = generate_professional_report(analyzer, results)
    
    # Step 4: Compare results
    compare_reports(original_path, professional_path)
    
    # Step 5: Summary
    print("\n🎉 EXECUTION COMPLETED")
    print("="*50)
    print("Check the output/reports/ directory for generated PDFs")
    print("\nTo customize the professional report:")
    print("1. Edit reports/weasyprint_report.py")
    print("2. Modify the CSS styles in get_css_styles()")
    print("3. Update the HTML template in get_html_template()")
    print("4. Add custom charts in the chart generation methods")

def quick_professional_report():
    """Quick function to generate just the professional report"""
    print("🎨 QUICK PROFESSIONAL REPORT GENERATION")
    print("="*50)
    
    if not setup_environment():
        return None
    
    analyzer, results = run_factoring_analysis()
    if analyzer is None or results is None:
        return None
    
    return generate_professional_report(analyzer, results)

if __name__ == "__main__":
    # Check if running with arguments
    if len(sys.argv) > 1 and sys.argv[1] == "--quick":
        # Quick mode - just generate professional report
        report_path = quick_professional_report()
        if report_path:
            print(f"\n✅ Quick report generated: {report_path}")
    else:
        # Full mode - generate both reports and compare
        main()