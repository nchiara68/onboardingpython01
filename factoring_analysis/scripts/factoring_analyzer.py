"""
🏢 Factoring Analyzer Module
Save this file as: scripts/factoring_analyzer.py
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

class FactoringAnalyzer:
    def __init__(self, csv_path, encoding='latin1'):
        """Initialize with invoice data"""
        try:
            self.df = pd.read_csv(csv_path, encoding=encoding)
            print(f"✅ Data loaded successfully with {encoding} encoding")
        except UnicodeDecodeError:
            print(f"❌ Failed with {encoding}, trying utf-8...")
            try:
                self.df = pd.read_csv(csv_path, encoding='utf-8')
                print("✅ Data loaded successfully with utf-8 encoding")
            except:
                print("❌ Failed with utf-8, trying default encoding...")
                self.df = pd.read_csv(csv_path)
                print("✅ Data loaded with default encoding")
        
        self.prepare_data()
    
    def prepare_data(self):
        """Clean and prepare data for analysis"""
        print("🔧 Starting data preparation...")
        
        # 1. Clean monetary columns - convert to numeric
        monetary_columns = ['Amount (USD)', 'Amt. Paid (USD)', 'Amt. Due (USD)']
        for col in monetary_columns:
            if col in self.df.columns:
                print(f"   Converting {col} to numeric...")
                # Remove currency symbols, commas, spaces, and convert to float
                self.df[col] = (self.df[col]
                               .astype(str)
                               .str.replace('$', '', regex=False)
                               .str.replace(',', '', regex=False)
                               .str.replace(' ', '', regex=False)
                               .str.replace('(', '-', regex=False)  # Handle negative values in parentheses
                               .str.replace(')', '', regex=False)
                               .replace('nan', '0')
                               .replace('', '0'))
                
                # Convert to numeric, handling any remaining issues
                self.df[col] = pd.to_numeric(self.df[col], errors='coerce').fillna(0)
        
        print("   ✅ Monetary columns converted to numeric")
        
        # 2. Split dataset by Type
        print("📊 Splitting dataset by Type...")
        
        # Check what types exist
        type_counts = self.df['Type'].value_counts()
        print(f"   Types found: {dict(type_counts)}")
        
        # Create separate datasets
        self.df_invoice = self.df[self.df['Type'] == 'Invoice'].copy()
        self.df_credit_memo = self.df[self.df['Type'] == 'Credit Memo'].copy()
        
        print(f"   📄 Invoice records: {len(self.df_invoice)}")
        print(f"   📝 Credit Memo records: {len(self.df_credit_memo)}")
        
        # For main analysis, we'll primarily focus on invoices
        # but keep the full dataset for reference
        self.df_full = self.df.copy()  # Keep original
        self.df = self.df_invoice.copy()  # Main analysis on invoices
        
        print("   ✅ Dataset split completed")
        
        # 3. Convert date columns
        print("📅 Converting date columns...")
        date_columns = ['Transaction Date', 'Due Date', 'Last Payment Date']
        for col in date_columns:
            if col in self.df.columns:
                self.df[col] = pd.to_datetime(self.df[col], errors='coerce')
        
        print("   ✅ Date columns converted")
        
        # 4. Calculate key metrics (for invoices)
        print("🧮 Calculating key metrics...")
        self.df['Days_Since_Transaction'] = (datetime.now() - self.df['Transaction Date']).dt.days
        self.df['Days_Past_Due'] = (datetime.now() - self.df['Due Date']).dt.days
        
        # Calculate payment metrics where possible
        if 'Last Payment Date' in self.df.columns:
            self.df['Days_to_Payment'] = (self.df['Last Payment Date'] - self.df['Transaction Date']).dt.days
        
        # Create aging categories
        self.df['Aging_Bucket'] = pd.cut(
            self.df['Days_Past_Due'],
            bins=[-float('inf'), 0, 30, 60, 90, float('inf')],
            labels=['Current', '1-30 Days', '31-60 Days', '61-90 Days', '90+ Days']
        )
        
        print("   ✅ Key metrics calculated")
        
        # 5. Data quality summary
        print("\n📋 DATA PREPARATION SUMMARY")
        print("=" * 50)
        print(f"Total records loaded: {len(self.df_full)}")
        print(f"Invoice records: {len(self.df_invoice)}")
        print(f"Credit Memo records: {len(self.df_credit_memo)}")
        print(f"Main analysis dataset (invoices): {self.df.shape}")
        print(f"Date range: {self.df['Transaction Date'].min()} to {self.df['Transaction Date'].max()}")
        
        # Show sample of converted monetary values
        print(f"\nSample monetary values after conversion:")
        for col in monetary_columns:
            if col in self.df.columns and len(self.df) > 0:
                sample_val = self.df[col].iloc[0] if not pd.isna(self.df[col].iloc[0]) else 0
                print(f"   {col}: ${sample_val:,.2f}")
        
        print("✅ Data preparation completed successfully!")

    def accounts_receivable_aging(self):
        """1. AR Aging Analysis"""
        print("\n" + "="*50)
        print("📊 ACCOUNTS RECEIVABLE AGING ANALYSIS")
        print("="*50)
        
        # Aging summary
        aging_summary = self.df.groupby('Aging_Bucket').agg({
            'Number': 'count',
            'Amount (USD)': 'sum',
            'Amt. Due (USD)': 'sum'
        }).round(2)
        
        aging_summary['Pct_of_Total_Due'] = (aging_summary['Amt. Due (USD)'] / 
                                           aging_summary['Amt. Due (USD)'].sum() * 100).round(2)
        
        print("\nAging Buckets Summary:")
        print(aging_summary)
        
        # Visualization
        plt.figure(figsize=(12, 8))
        
        plt.subplot(2, 2, 1)
        aging_summary['Amt. Due (USD)'].plot(kind='bar', color='skyblue')
        plt.title('Amount Due by Aging Bucket')
        plt.ylabel('Amount Due (USD)')
        plt.xticks(rotation=45)
        
        plt.subplot(2, 2, 2)
        plt.pie(aging_summary['Amt. Due (USD)'], labels=aging_summary.index, autopct='%1.1f%%')
        plt.title('AR Distribution by Aging')
        
        plt.subplot(2, 2, 3)
        aging_summary['Number'].plot(kind='bar', color='lightcoral')
        plt.title('Invoice Count by Aging Bucket')
        plt.ylabel('Number of Invoices')
        plt.xticks(rotation=45)
        
        plt.tight_layout()
        plt.show()
        
        return aging_summary

    def client_risk_analysis(self):
        """2. Client Risk & Exposure Analysis"""
        print("\n" + "="*50)
        print("⚠️  CLIENT RISK & EXPOSURE ANALYSIS")
        print("="*50)
        
        # Client exposure summary
        client_summary = self.df.groupby('Applied to').agg({
            'Number': 'count',
            'Amount (USD)': 'sum',
            'Amt. Due (USD)': 'sum',
            'Amt. Paid (USD)': 'sum',
            'Days_Past_Due': 'mean'
        }).round(2)
        
        client_summary['Payment_Rate'] = (client_summary['Amt. Paid (USD)'] / 
                                        client_summary['Amount (USD)'] * 100).round(2)
        client_summary = client_summary.sort_values('Amt. Due (USD)', ascending=False)
        
        print("\nTop 10 Clients by Outstanding Amount:")
        print(client_summary.head(10))
        
        # Risk flags
        high_risk_clients = client_summary[
            (client_summary['Amt. Due (USD)'] > client_summary['Amt. Due (USD)'].quantile(0.8)) |
            (client_summary['Days_Past_Due'] > 60) |
            (client_summary['Payment_Rate'] < 50)
        ]
        
        print(f"\n🚨 HIGH RISK CLIENTS ({len(high_risk_clients)}):")
        print(high_risk_clients)
        
        # Visualization
        plt.figure(figsize=(12, 6))
        
        plt.subplot(1, 2, 1)
        top_10_clients = client_summary.head(10)
        top_10_clients['Amt. Due (USD)'].plot(kind='barh')
        plt.title('Top 10 Clients by Amount Due')
        plt.xlabel('Amount Due (USD)')
        
        plt.subplot(1, 2, 2)
        plt.scatter(client_summary['Amt. Due (USD)'], client_summary['Days_Past_Due'], alpha=0.6)
        plt.xlabel('Amount Due (USD)')
        plt.ylabel('Average Days Past Due')
        plt.title('Client Risk Matrix')
        
        plt.tight_layout()
        plt.show()
        
        return client_summary, high_risk_clients

    def executive_summary(self):
        """Generate Executive Summary"""
        print("\n" + "="*60)
        print("📋 EXECUTIVE SUMMARY")
        print("="*60)
        
        total_outstanding = self.df['Amt. Due (USD)'].sum()
        total_invoices = len(self.df)
        avg_days_past_due = self.df['Days_Past_Due'].mean()
        
        # Key metrics
        print(f"Total Outstanding Amount: ${total_outstanding:,.2f}")
        print(f"Total Invoices: {total_invoices}")
        print(f"Average Days Past Due: {avg_days_past_due:.1f}")
        
        # Risk assessment
        high_risk_amount = self.df[self.df['Days_Past_Due'] > 90]['Amt. Due (USD)'].sum()
        risk_percentage = (high_risk_amount / total_outstanding) * 100
        
        print(f"High Risk Amount (90+ days): ${high_risk_amount:,.2f} ({risk_percentage:.1f}%)")
        
        # Collection efficiency
        total_billed = self.df['Amount (USD)'].sum()
        total_collected = self.df['Amt. Paid (USD)'].sum()
        collection_rate = (total_collected / total_billed) * 100
        
        print(f"Collection Rate: {collection_rate:.1f}%")
        
        # Recommendations
        print("\n📌 KEY RECOMMENDATIONS:")
        if risk_percentage > 20:
            print("• URGENT: High concentration of 90+ day receivables requires immediate attention")
        if collection_rate < 80:
            print("• Implement more aggressive collection procedures")
        if avg_days_past_due > 45:
            print("• Review credit policies and payment terms")
        
        return {
            'total_outstanding': total_outstanding,
            'total_invoices': total_invoices,
            'avg_days_past_due': avg_days_past_due,
            'collection_rate': collection_rate,
            'risk_percentage': risk_percentage
        }

    def run_full_analysis(self):
        """Run complete factoring analysis"""
        print("🏢 FACTORING FIRM INVOICE ANALYSIS")
        print("="*60)
        
        # Run main analyses
        aging_summary = self.accounts_receivable_aging()
        client_summary, high_risk_clients = self.client_risk_analysis()
        executive_summary = self.executive_summary()
        
        return {
            'aging_summary': aging_summary,
            'client_summary': client_summary,
            'high_risk_clients': high_risk_clients,
            'executive_summary': executive_summary
        }


# Convenience function for quick analysis
def analyze_invoices(csv_path='../data/TecnoCargoInvoiceDataset01.csv', encoding='latin1'):
    """
    Quick function to run complete invoice analysis
    
    Parameters:
    csv_path (str): Path to CSV file
    encoding (str): File encoding
    
    Returns:
    dict: Analysis results
    """
    analyzer = FactoringAnalyzer(csv_path, encoding)
    return analyzer.run_full_analysis()


if __name__ == "__main__":
    # Run analysis if script is executed directly
    print("🚀 Running factoring analysis...")
    results = analyze_invoices()
    print("✅ Analysis completed!")