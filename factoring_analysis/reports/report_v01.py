"""
📄 Professional PDF Report Generator for Factoring Risk Analysis
Creates a comprehensive PDF report with all analysis results, tables, and figures
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.backends.backend_pdf import PdfPages
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

def generate_factoring_pdf_report(analyzer, results, output_filename='Factoring_Risk_Analysis_Report.pdf'):
    """
    Generate a comprehensive PDF report for factoring risk analysis
    
    Parameters:
    analyzer: CorrectedFactoringAnalyzer instance
    results: Results dictionary from run_full_analysis()
    output_filename: Name of the output PDF file
    """
    
    print("📄 GENERATING COMPREHENSIVE PDF REPORT")
    print("="*60)
    
    # Set up the PDF file
    output_path = f'../data/{output_filename}'
    
    with PdfPages(output_path) as pdf:
        
        # Configure matplotlib for better PDF output
        plt.style.use('default')
        sns.set_palette("husl")
        
        # PAGE 1: COVER PAGE
        print("   📝 Creating cover page...")
        create_cover_page(pdf, analyzer, results)
        
        # PAGE 2: EXECUTIVE SUMMARY
        print("   📊 Creating executive summary...")
        create_executive_summary_page(pdf, analyzer, results)
        
        # PAGE 3: PORTFOLIO OVERVIEW
        print("   📈 Creating portfolio overview...")
        create_portfolio_overview_page(pdf, analyzer, results)
        
        # PAGE 4: PAID INVOICES ANALYSIS
        print("   💰 Creating paid invoices analysis...")
        create_paid_invoices_analysis_page(pdf, analyzer, results)
        
        # PAGE 5: OUTSTANDING INVOICES ANALYSIS
        print("   ⏳ Creating outstanding invoices analysis...")
        create_outstanding_invoices_analysis_page(pdf, analyzer, results)
        
        # PAGE 6: RISK ASSESSMENT
        print("   🚨 Creating risk assessment...")
        create_risk_assessment_page(pdf, analyzer, results)
        
        # PAGE 7: CLIENT ANALYSIS
        print("   👥 Creating client analysis...")
        create_client_analysis_page(pdf, analyzer, results)
        
        # PAGE 8: RECOMMENDATIONS
        print("   📌 Creating recommendations...")
        create_recommendations_page(pdf, analyzer, results)
        
        # PAGE 9: DETAILED TABLES
        print("   📋 Creating detailed tables...")
        create_detailed_tables_page(pdf, analyzer, results)
        
        # Add metadata
        pdf_metadata = pdf.infodict()
        pdf_metadata['Title'] = 'Factoring Risk Analysis Report'
        pdf_metadata['Author'] = 'Factoring Analysis System'
        pdf_metadata['Subject'] = 'Invoice Portfolio Risk Assessment'
        pdf_metadata['Creator'] = 'Python Factoring Analyzer'
        pdf_metadata['CreationDate'] = datetime.now()
    
    print(f"✅ PDF report generated: {output_path}")
    return output_path

def create_cover_page(pdf, analyzer, results):
    """Create the cover page"""
    fig, ax = plt.subplots(figsize=(8.5, 11))
    ax.axis('off')
    
    # Title
    ax.text(0.5, 0.85, 'FACTORING RISK ANALYSIS', 
            fontsize=28, weight='bold', ha='center', va='center')
    ax.text(0.5, 0.8, 'Invoice Portfolio Assessment Report', 
            fontsize=18, ha='center', va='center', style='italic')
    
    # Key metrics box
    exec_summary = results['executive_summary']
    
    # Create a summary box
    summary_text = f"""
PORTFOLIO SUMMARY
{'='*50}

Total Invoices Analyzed: {len(analyzer.df_invoice):,}
Collection Rate: {exec_summary['collection_rate']:.1f}%
Average Payment Delay: {exec_summary['avg_payment_delay']:.1f} days
Total Outstanding: ${exec_summary['total_outstanding']:,.2f}
Risk Level (90+ days): {exec_summary['risk_percentage']:.1f}%

Analysis Period: January 2024 - June 2025
Report Date: {datetime.now().strftime('%B %d, %Y')}
Data Snapshot: June 23, 2025
"""
    
    ax.text(0.5, 0.55, summary_text, fontsize=12, ha='center', va='center',
            bbox=dict(boxstyle="round,pad=1", facecolor="lightblue", alpha=0.8))
    
    # Risk assessment indicator
    risk_level = "LOW" if exec_summary['risk_percentage'] < 10 else "MODERATE" if exec_summary['risk_percentage'] < 20 else "HIGH"
    risk_color = "green" if risk_level == "LOW" else "orange" if risk_level == "MODERATE" else "red"
    
    ax.text(0.5, 0.25, f'OVERALL RISK LEVEL: {risk_level}', 
            fontsize=20, weight='bold', ha='center', va='center',
            bbox=dict(boxstyle="round,pad=0.5", facecolor=risk_color, alpha=0.7))
    
    # Footer
    ax.text(0.5, 0.1, 'CONFIDENTIAL - For Internal Use Only', 
            fontsize=10, ha='center', va='center', style='italic', alpha=0.7)
    
    plt.tight_layout()
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()

def create_executive_summary_page(pdf, analyzer, results):
    """Create executive summary page"""
    fig = plt.figure(figsize=(8.5, 11))
    
    # Title
    fig.suptitle('EXECUTIVE SUMMARY', fontsize=20, weight='bold', y=0.95)
    
    exec_summary = results['executive_summary']
    
    # Create 4 subplots for key metrics
    gs = fig.add_gridspec(3, 2, height_ratios=[1, 1, 1.5], hspace=0.4, wspace=0.3)
    
    # Key Performance Indicators
    ax1 = fig.add_subplot(gs[0, 0])
    ax1.pie([exec_summary['collection_rate'], 100-exec_summary['collection_rate']], 
           labels=['Collected', 'Outstanding'], autopct='%1.1f%%', 
           colors=['lightgreen', 'lightcoral'])
    ax1.set_title('Collection Rate', weight='bold')
    
    ax2 = fig.add_subplot(gs[0, 1])
    payment_delay = exec_summary['avg_payment_delay']
    colors = ['green' if payment_delay < 15 else 'orange' if payment_delay < 30 else 'red']
    ax2.bar(['Avg Payment Delay'], [payment_delay], color=colors)
    ax2.set_ylabel('Days')
    ax2.set_title('Payment Delay Performance', weight='bold')
    
    # Risk breakdown
    ax3 = fig.add_subplot(gs[1, :])
    if results['outstanding_invoices_aging'] is not None:
        outstanding_aging = results['outstanding_invoices_aging']
        ax3.bar(outstanding_aging.index, outstanding_aging['Invoice_Count'], 
               color=['green', 'yellow', 'orange', 'red', 'darkred'])
        ax3.set_title('Outstanding Invoices by Risk Category', weight='bold')
        ax3.set_ylabel('Number of Invoices')
        plt.setp(ax3.xaxis.get_majorticklabels(), rotation=45)
    
    # Key findings text
    ax4 = fig.add_subplot(gs[2, :])
    ax4.axis('off')
    
    findings_text = f"""
KEY FINDINGS & INSIGHTS

Portfolio Health: {'EXCELLENT' if exec_summary['collection_rate'] > 90 else 'GOOD' if exec_summary['collection_rate'] > 80 else 'NEEDS IMPROVEMENT'}
• Collection rate of {exec_summary['collection_rate']:.1f}% {'exceeds' if exec_summary['collection_rate'] > 90 else 'meets' if exec_summary['collection_rate'] > 80 else 'falls below'} industry standards
• Average payment delay of {exec_summary['avg_payment_delay']:.1f} days is {'excellent' if exec_summary['avg_payment_delay'] < 15 else 'acceptable' if exec_summary['avg_payment_delay'] < 30 else 'concerning'}

Risk Assessment: {exec_summary['risk_percentage']:.1f}% of outstanding amount is 90+ days overdue
• Total outstanding amount: ${exec_summary['total_outstanding']:,.2f}
• Portfolio size: {len(analyzer.df_invoice):,} total invoices
• Paid invoices: {len(analyzer.df_paid_invoices):,} ({len(analyzer.df_paid_invoices)/len(analyzer.df_invoice)*100:.1f}%)
• Outstanding invoices: {len(analyzer.df_outstanding_invoices):,} ({len(analyzer.df_outstanding_invoices)/len(analyzer.df_invoice)*100:.1f}%)

Liquidity Position: {'Strong' if len(analyzer.df_outstanding_invoices)/len(analyzer.df_invoice) < 0.15 else 'Moderate' if len(analyzer.df_outstanding_invoices)/len(analyzer.df_invoice) < 0.25 else 'Weak'}
• Low percentage of outstanding invoices indicates strong cash conversion
• Payment behavior shows {'consistent' if exec_summary['avg_payment_delay'] < 20 else 'variable'} customer payment patterns
"""
    
    ax4.text(0.05, 0.95, findings_text, fontsize=11, ha='left', va='top', 
            transform=ax4.transAxes, 
            bbox=dict(boxstyle="round,pad=0.5", facecolor="lightyellow", alpha=0.8))
    
    plt.tight_layout()
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()

def create_portfolio_overview_page(pdf, analyzer, results):
    """Create portfolio overview page"""
    fig = plt.figure(figsize=(8.5, 11))
    fig.suptitle('PORTFOLIO OVERVIEW', fontsize=20, weight='bold', y=0.95)
    
    gs = fig.add_gridspec(3, 2, height_ratios=[1, 1, 1], hspace=0.4, wspace=0.3)
    
    # Monthly invoice volume (if possible to extract)
    ax1 = fig.add_subplot(gs[0, :])
    if 'Transaction Date' in analyzer.df_invoice.columns:
        monthly_volume = analyzer.df_invoice.groupby(analyzer.df_invoice['Transaction Date'].dt.to_period('M')).size()
        monthly_volume.plot(kind='line', ax=ax1, marker='o')
        ax1.set_title('Monthly Invoice Volume Trend', weight='bold')
        ax1.set_ylabel('Number of Invoices')
        ax1.tick_params(axis='x', rotation=45)
    
    # Amount distribution
    ax2 = fig.add_subplot(gs[1, 0])
    analyzer.df_invoice['Amount (USD)'].hist(bins=20, ax=ax2, alpha=0.7, color='skyblue')
    ax2.set_title('Invoice Amount Distribution', weight='bold')
    ax2.set_xlabel('Amount (USD)')
    ax2.set_ylabel('Frequency')
    
    # Payment status overview
    ax3 = fig.add_subplot(gs[1, 1])
    status_counts = [len(analyzer.df_paid_invoices), len(analyzer.df_outstanding_invoices)]
    labels = ['Paid', 'Outstanding']
    colors = ['lightgreen', 'lightcoral']
    ax3.pie(status_counts, labels=labels, autopct='%1.1f%%', colors=colors)
    ax3.set_title('Payment Status Distribution', weight='bold')
    
    # Summary statistics table
    ax4 = fig.add_subplot(gs[2, :])
    ax4.axis('off')
    
    summary_stats = pd.DataFrame({
        'Metric': [
            'Total Invoices',
            'Total Billed Amount',
            'Total Collected',
            'Total Outstanding',
            'Average Invoice Amount',
            'Largest Invoice',
            'Smallest Invoice',
            'Collection Rate'
        ],
        'Value': [
            f"{len(analyzer.df_invoice):,}",
            f"${analyzer.df_invoice['Amount (USD)'].sum():,.2f}",
            f"${analyzer.df_invoice['Amt. Paid (USD)'].sum():,.2f}",
            f"${analyzer.df_outstanding_invoices['Amt. Due (USD)'].sum():,.2f}" if len(analyzer.df_outstanding_invoices) > 0 else "$0.00",
            f"${analyzer.df_invoice['Amount (USD)'].mean():,.2f}",
            f"${analyzer.df_invoice['Amount (USD)'].max():,.2f}",
            f"${analyzer.df_invoice['Amount (USD)'].min():,.2f}",
            f"{results['executive_summary']['collection_rate']:.1f}%"
        ]
    })
    
    table = ax4.table(cellText=summary_stats.values, colLabels=summary_stats.columns,
                     cellLoc='left', loc='center', colWidths=[0.4, 0.3])
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1.2, 1.5)
    ax4.set_title('Portfolio Summary Statistics', weight='bold', pad=20)
    
    plt.tight_layout()
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()

def create_paid_invoices_analysis_page(pdf, analyzer, results):
    """Create paid invoices analysis page"""
    if results['paid_invoices_aging'] is None or len(analyzer.df_paid_invoices) == 0:
        return
        
    fig = plt.figure(figsize=(8.5, 11))
    fig.suptitle('PAID INVOICES ANALYSIS', fontsize=20, weight='bold', y=0.95)
    
    gs = fig.add_gridspec(3, 2, height_ratios=[1, 1, 1], hspace=0.4, wspace=0.3)
    
    paid_aging = results['paid_invoices_aging']
    
    # Payment delay distribution
    ax1 = fig.add_subplot(gs[0, :])
    paid_aging['Invoice_Count'].plot(kind='bar', ax=ax1, color='lightblue')
    ax1.set_title('Payment Delay Distribution', weight='bold')
    ax1.set_ylabel('Number of Invoices')
    ax1.tick_params(axis='x', rotation=45)
    
    # Average delay by category
    ax2 = fig.add_subplot(gs[1, 0])
    paid_aging['Avg_Days_Late'].plot(kind='bar', ax=ax2, color='orange')
    ax2.set_title('Average Days Late by Category', weight='bold')
    ax2.set_ylabel('Days')
    ax2.tick_params(axis='x', rotation=45)
    
    # Amount by aging category
    ax3 = fig.add_subplot(gs[1, 1])
    ax3.pie(paid_aging['Total_Amount'], labels=paid_aging.index, autopct='%1.1f%%')
    ax3.set_title('Amount Distribution by Aging', weight='bold')
    
    # Detailed statistics table
    ax4 = fig.add_subplot(gs[2, :])
    ax4.axis('off')
    
    # Prepare table data
    table_data = paid_aging.reset_index()
    table = ax4.table(cellText=table_data.values, colLabels=table_data.columns,
                     cellLoc='center', loc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1, 1.5)
    ax4.set_title('Detailed Paid Invoices Aging Statistics', weight='bold', pad=20)
    
    plt.tight_layout()
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()

def create_outstanding_invoices_analysis_page(pdf, analyzer, results):
    """Create outstanding invoices analysis page"""
    if results['outstanding_invoices_aging'] is None or len(analyzer.df_outstanding_invoices) == 0:
        return
        
    fig = plt.figure(figsize=(8.5, 11))
    fig.suptitle('OUTSTANDING INVOICES ANALYSIS', fontsize=20, weight='bold', y=0.95)
    
    gs = fig.add_gridspec(3, 2, height_ratios=[1, 1, 1], hspace=0.4, wspace=0.3)
    
    outstanding_aging = results['outstanding_invoices_aging']
    
    # Outstanding by aging category
    ax1 = fig.add_subplot(gs[0, 0])
    outstanding_aging['Invoice_Count'].plot(kind='bar', ax=ax1, 
                                          color=['green', 'yellow', 'orange', 'red', 'darkred'])
    ax1.set_title('Outstanding Invoices by Category', weight='bold')
    ax1.set_ylabel('Number of Invoices')
    ax1.tick_params(axis='x', rotation=45)
    
    # Amount due by category
    ax2 = fig.add_subplot(gs[0, 1])
    outstanding_aging['Total_Due'].plot(kind='bar', ax=ax2,
                                       color=['green', 'yellow', 'orange', 'red', 'darkred'])
    ax2.set_title('Amount Due by Category', weight='bold')
    ax2.set_ylabel('Amount Due (USD)')
    ax2.tick_params(axis='x', rotation=45)
    
    # Risk distribution pie chart
    ax3 = fig.add_subplot(gs[1, :])
    ax3.pie(outstanding_aging['Total_Due'], labels=outstanding_aging.index, 
           autopct='%1.1f%%', colors=['green', 'yellow', 'orange', 'red', 'darkred'])
    ax3.set_title('Outstanding Amount Risk Distribution', weight='bold')
    
    # Detailed statistics table
    ax4 = fig.add_subplot(gs[2, :])
    ax4.axis('off')
    
    table_data = outstanding_aging.reset_index()
    table = ax4.table(cellText=table_data.values, colLabels=table_data.columns,
                     cellLoc='center', loc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1, 1.5)
    ax4.set_title('Detailed Outstanding Invoices Statistics', weight='bold', pad=20)
    
    plt.tight_layout()
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()

def create_risk_assessment_page(pdf, analyzer, results):
    """Create risk assessment page"""
    fig = plt.figure(figsize=(8.5, 11))
    fig.suptitle('RISK ASSESSMENT', fontsize=20, weight='bold', y=0.95)
    
    exec_summary = results['executive_summary']
    
    # Risk level indicator
    ax1 = fig.add_subplot(3, 1, 1)
    ax1.axis('off')
    
    risk_level = "LOW" if exec_summary['risk_percentage'] < 10 else "MODERATE" if exec_summary['risk_percentage'] < 20 else "HIGH"
    risk_color = "green" if risk_level == "LOW" else "orange" if risk_level == "MODERATE" else "red"
    
    ax1.text(0.5, 0.5, f'OVERALL RISK LEVEL: {risk_level}', 
            fontsize=24, weight='bold', ha='center', va='center',
            bbox=dict(boxstyle="round,pad=0.5", facecolor=risk_color, alpha=0.7),
            transform=ax1.transAxes)
    
    # Risk factors analysis
    ax2 = fig.add_subplot(3, 2, 3)
    risk_factors = ['Collection Rate', 'Payment Delay', '90+ Days Risk', 'Portfolio Size']
    risk_scores = [
        100 - exec_summary['collection_rate'],  # Higher collection rate = lower risk
        min(exec_summary['avg_payment_delay'] * 2, 100),  # Higher delay = higher risk
        exec_summary['risk_percentage'],  # Direct risk percentage
        min(len(analyzer.df_outstanding_invoices) / 100, 100)  # More outstanding = higher risk
    ]
    
    colors = ['green' if score < 20 else 'yellow' if score < 40 else 'orange' if score < 60 else 'red' for score in risk_scores]
    ax2.barh(risk_factors, risk_scores, color=colors)
    ax2.set_title('Risk Factor Analysis', weight='bold')
    ax2.set_xlabel('Risk Score')
    
    # Concentration risk
    ax3 = fig.add_subplot(3, 2, 4)
    if len(analyzer.df_outstanding_invoices) > 0:
        client_concentration = analyzer.df_outstanding_invoices.groupby('Applied to')['Amt. Due (USD)'].sum().sort_values(ascending=False).head(10)
        client_concentration.plot(kind='bar', ax=ax3, color='lightcoral')
        ax3.set_title('Top 10 Client Concentration Risk', weight='bold')
        ax3.set_ylabel('Amount Due (USD)')
        ax3.tick_params(axis='x', rotation=45)
    
    # Risk summary text
    ax4 = fig.add_subplot(3, 1, 3)
    ax4.axis('off')
    
    risk_text = f"""
RISK ASSESSMENT SUMMARY

Primary Risk Factors:
• 90+ Day Receivables: {exec_summary['risk_percentage']:.1f}% of outstanding amount
• Collection Efficiency: {exec_summary['collection_rate']:.1f}% overall collection rate
• Payment Delays: {exec_summary['avg_payment_delay']:.1f} days average delay
• Portfolio Liquidity: {len(analyzer.df_outstanding_invoices)/len(analyzer.df_invoice)*100:.1f}% outstanding ratio

Risk Mitigation Status:
{'✅ LOW RISK: Portfolio shows strong collection performance' if risk_level == 'LOW' else 
 '⚠️ MODERATE RISK: Monitor closely and implement preventive measures' if risk_level == 'MODERATE' else
 '🚨 HIGH RISK: Immediate action required to address collection issues'}

Recommended Actions:
• {'Continue current collection procedures' if risk_level == 'LOW' else 'Implement enhanced collection procedures' if risk_level == 'MODERATE' else 'Urgent collection intervention required'}
• {'Monitor for any deterioration' if risk_level == 'LOW' else 'Weekly risk monitoring' if risk_level == 'MODERATE' else 'Daily risk monitoring'}
• {'Maintain credit policies' if risk_level == 'LOW' else 'Review credit policies' if risk_level == 'MODERATE' else 'Tighten credit policies immediately'}
"""
    
    ax4.text(0.05, 0.95, risk_text, fontsize=11, ha='left', va='top',
            transform=ax4.transAxes,
            bbox=dict(boxstyle="round,pad=0.5", facecolor="lightyellow", alpha=0.8))
    
    plt.tight_layout()
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()

def create_client_analysis_page(pdf, analyzer, results):
    """Create client analysis page"""
    fig = plt.figure(figsize=(8.5, 11))
    fig.suptitle('CLIENT ANALYSIS', fontsize=20, weight='bold', y=0.95)
    
    gs = fig.add_gridspec(3, 2, height_ratios=[1, 1, 1], hspace=0.4, wspace=0.3)
    
    # Top clients by amount
    ax1 = fig.add_subplot(gs[0, 0])
    top_clients = analyzer.df_invoice.groupby('Applied to')['Amount (USD)'].sum().sort_values(ascending=False).head(10)
    top_clients.plot(kind='barh', ax=ax1, color='lightblue')
    ax1.set_title('Top 10 Clients by Total Amount', weight='bold')
    ax1.set_xlabel('Total Amount (USD)')
    
    # Client payment performance
    ax2 = fig.add_subplot(gs[0, 1])
    if len(analyzer.df_paid_invoices) > 0:
        client_performance = analyzer.df_paid_invoices.groupby('Applied to')['Aging_Delay'].mean().sort_values().head(10)
        colors = ['green' if delay <= 0 else 'yellow' if delay <= 15 else 'orange' if delay <= 30 else 'red' for delay in client_performance.values]
        client_performance.plot(kind='barh', ax=ax2, color=colors)
        ax2.set_title('Top 10 Clients by Payment Speed', weight='bold')
        ax2.set_xlabel('Average Payment Delay (Days)')
    
    # Risk clients
    ax3 = fig.add_subplot(gs[1, :])
    if len(analyzer.df_outstanding_invoices) > 0:
        risk_clients = analyzer.df_outstanding_invoices.groupby('Applied to')['Amt. Due (USD)'].sum().sort_values(ascending=False).head(10)
        risk_clients.plot(kind='bar', ax=ax3, color='lightcoral')
        ax3.set_title('Top 10 Clients by Outstanding Amount (Risk)', weight='bold')
        ax3.set_ylabel('Outstanding Amount (USD)')
        ax3.tick_params(axis='x', rotation=45)
    
    # Client summary statistics
    ax4 = fig.add_subplot(gs[2, :])
    ax4.axis('off')
    
    client_stats = analyzer.df_invoice.groupby('Applied to').agg({
        'Number': 'count',
        'Amount (USD)': ['sum', 'mean'],
        'Amt. Paid (USD)': 'sum',
        'Amt. Due (USD)': 'sum'
    }).round(2)
    
    client_stats.columns = ['Invoice_Count', 'Total_Amount', 'Avg_Amount', 'Total_Paid', 'Total_Due']
    client_stats['Collection_Rate'] = (client_stats['Total_Paid'] / client_stats['Total_Amount'] * 100).round(1)
    
    # Show top 10 clients
    top_client_stats = client_stats.sort_values('Total_Amount', ascending=False).head(10)
    
    table = ax4.table(cellText=top_client_stats.values, 
                     rowLabels=top_client_stats.index,
                     colLabels=top_client_stats.columns,
                     cellLoc='center', loc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(8)
    table.scale(1, 1.5)
    ax4.set_title('Top 10 Clients - Detailed Statistics', weight='bold', pad=20)
    
    plt.tight_layout()
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()

def create_recommendations_page(pdf, analyzer, results):
    """Create recommendations page"""
    fig, ax = plt.subplots(figsize=(8.5, 11))
    ax.axis('off')
    
    exec_summary = results['executive_summary']
    
    # Title
    ax.text(0.5, 0.95, 'RECOMMENDATIONS & ACTION PLAN', 
            fontsize=20, weight='bold', ha='center', va='top')
    
    # Generate recommendations based on analysis
    recommendations = []
    
    # Collection rate recommendations
    if exec_summary['collection_rate'] > 90:
        recommendations.append("✅ COLLECTION EFFICIENCY: Excellent performance - maintain current procedures")
    elif exec_summary['collection_rate'] > 80:
        recommendations.append("⚠️ COLLECTION EFFICIENCY: Good but can improve - review slow-paying accounts")
    else:
        recommendations.append("🚨 COLLECTION EFFICIENCY: Poor performance - immediate intervention required")
    
    # Payment delay recommendations
    if exec_summary['avg_payment_delay'] < 15:
        recommendations.append("✅ PAYMENT TIMING: Excellent payment behavior - maintain credit terms")
    elif exec_summary['avg_payment_delay'] < 30:
        recommendations.append("⚠️ PAYMENT TIMING: Acceptable delays - monitor for deterioration")
    else:
        recommendations.append("🚨 PAYMENT TIMING: Excessive delays - review credit policies and terms")
    
    # Risk level recommendations
    if exec_summary['risk_percentage'] < 10:
        recommendations.append("✅ RISK LEVEL: Low risk portfolio - maintain monitoring")
    elif exec_summary['risk_percentage'] < 20:
        recommendations.append("⚠️ RISK LEVEL: Moderate risk - implement enhanced collection procedures")
    else:
        recommendations.append("🚨 RISK LEVEL: High risk - urgent collection action required")
    
    # Specific action items
    action_items = [
        "IMMEDIATE ACTIONS (0-30 days):",
        "• Focus collection efforts on 90+ day receivables",
        "• Contact top 10 outstanding clients for payment plans",
        "• Review and update credit policies if risk is elevated",
        "",
        "SHORT-TERM ACTIONS (30-90 days):",
        "• Implement weekly aging reports",
        "• Establish client payment performance scorecards", 
        "• Consider offering early payment discounts",
        "",
        "LONG-TERM STRATEGY (90+ days):",
        "• Regular portfolio risk assessment (monthly)",
        "• Client credit limit reviews based on payment history",
        "• Implement automated collection procedures",
        "",
        "MONITORING & REPORTING:",
        "• Weekly aging reports for collections team",
        "• Monthly executive dashboard with key metrics",
        "• Quarterly portfolio risk assessment",
        "• Annual credit policy review"
    ]
    
    # Combine recommendations and action items
    full_text = "\n".join(recommendations) + "\n\n" + "\n".join(action_items)
    
    ax.text(0.05, 0.85, full_text, fontsize=11, ha='left', va='top',
            bbox=dict(boxstyle="round,pad=1", facecolor="lightgreen", alpha=0.3))
    
    # Add a priority matrix
    ax.text(0.05, 0.25, 
            "PRIORITY MATRIX:\n\n" +
            "HIGH PRIORITY: 90+ day receivables, large outstanding amounts\n" +
            "MEDIUM PRIORITY: 60-90 day receivables, repeat late payers\n" +
            "LOW PRIORITY: Current and 1-30 day receivables\n\n" +
            "Success will be measured by:\n" +
            "• Reduction in 90+ day receivables\n" +
            "• Improvement in overall collection rate\n" +
            "• Decrease in average payment delays\n" +
            "• Maintenance of strong client relationships",
            fontsize=11, ha='left', va='top',
            bbox=dict(boxstyle="round,pad=0.5", facecolor="lightyellow", alpha=0.8))
    
    plt.tight_layout()
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()

def create_detailed_tables_page(pdf, analyzer, results):
    """Create detailed tables page"""
    fig = plt.figure(figsize=(8.5, 11))
    fig.suptitle('DETAILED DATA TABLES', fontsize=20, weight='bold', y=0.95)
    
    gs = fig.add_gridspec(2, 1, height_ratios=[1, 1], hspace=0.3)
    
    # Paid invoices aging table
    ax1 = fig.add_subplot(gs[0])
    ax1.axis('off')
    if results['paid_invoices_aging'] is not None:
        paid_data = results['paid_invoices_aging'].reset_index()
        table1 = ax1.table(cellText=paid_data.values, colLabels=paid_data.columns,
                          cellLoc='center', loc='center')
        table1.auto_set_font_size(False)
        table1.set_fontsize(10)
        table1.scale(1, 2)
        ax1.set_title('Paid Invoices Aging - Detailed Statistics', weight='bold', pad=20)
    
    # Outstanding invoices aging table  
    ax2 = fig.add_subplot(gs[1])
    ax2.axis('off')
    if results['outstanding_invoices_aging'] is not None:
        outstanding_data = results['outstanding_invoices_aging'].reset_index()
        table2 = ax2.table(cellText=outstanding_data.values, colLabels=outstanding_data.columns,
                          cellLoc='center', loc='center')
        table2.auto_set_font_size(False)
        table2.set_fontsize(10)
        table2.scale(1, 2)
        ax2.set_title('Outstanding Invoices Aging - Detailed Statistics', weight='bold', pad=20)
    
    plt.tight_layout()
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()

# Main function to generate report
def create_complete_factoring_report(analyzer, results, filename='Factoring_Risk_Analysis_Report.pdf'):
    """
    Main function to create the complete PDF report
    
    Usage:
    # After running your analysis:
    report_path = create_complete_factoring_report(analyzer, results)
    """
    
    return generate_factoring_pdf_report(analyzer, results, filename)

print("✅ PDF Report Generator loaded successfully!")
print("📄 Ready to generate comprehensive factoring risk analysis report!")

# 📄 STEP 2: INSTALL REQUIRED PACKAGE (if needed)
import subprocess
import sys

try:
    import matplotlib
    print("✅ Matplotlib already installed")
except ImportError:
    print("Installing matplotlib...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "matplotlib"])

# 📄 STEP 3: GENERATE THE COMPLETE PDF REPORT
print("📄 GENERATING COMPREHENSIVE PDF REPORT...")
print("="*60)

# Generate the PDF report (make sure you have run the analysis first)
report_path = create_complete_factoring_report(
    analyzer=analyzer,  # Your analyzer instance
    results=results,    # Your analysis results
    filename='TecnoCargo_Factoring_Risk_Analysis_Report.pdf'
)

print(f"🎉 SUCCESS!")
print(f"📁 PDF Report saved to: {report_path}")
print(f"📊 Report contains 9 comprehensive pages with all analysis results!")