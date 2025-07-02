# 🏢 COMPLETE CORRECTED FACTORING ANALYZER CLASS
# Copy and run this entire cell to define the complete class

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

class CorrectedFactoringAnalyzer:
    def __init__(self, csv_path, encoding='latin1'):
        """Initialize with invoice data"""
        try:
            self.df = pd.read_csv(csv_path, encoding=encoding)
            print(f"✅ Data loaded successfully with {encoding} encoding")
        except Exception as e:
            print(f"❌ Error loading data: {e}")
            return
        
        self.prepare_data()
    
    def prepare_data(self):
        """Clean and prepare data for analysis with MM/DD/YYYY format"""
        print("🔧 Starting data preparation...")
        
        # 1. Clean monetary columns
        monetary_columns = ['Amount (USD)', 'Amt. Paid (USD)', 'Amt. Due (USD)']
        for col in monetary_columns:
            if col in self.df.columns:
                print(f"   Converting {col} to numeric...")
                self.df[col] = pd.to_numeric(
                    self.df[col].astype(str)
                    .str.replace('$', '', regex=False)
                    .str.replace(',', '', regex=False)
                    .str.replace('(', '-', regex=False)
                    .str.replace(')', '', regex=False),
                    errors='coerce'
                ).fillna(0)
        
        print("   ✅ Monetary columns converted")
        
        # 2. Split by type
        self.df_invoice = self.df[self.df['Type'] == 'Invoice'].copy()
        self.df_credit_memo = self.df[self.df['Type'] == 'Credit Memo'].copy()
        self.df_full = self.df.copy()
        
        print(f"   📄 Invoice records: {len(self.df_invoice)}")
        print(f"   📝 Credit Memo records: {len(self.df_credit_memo)}")
        
        # 3. Convert dates with MM/DD/YYYY format
        print("📅 Converting date columns (MM/DD/YYYY format)...")
        date_columns = ['Transaction Date', 'Due Date', 'Last Payment Date']
        
        for col in date_columns:
            if col in self.df_invoice.columns:
                self.df_invoice[col] = pd.to_datetime(
                    self.df_invoice[col], 
                    format='%m/%d/%Y', 
                    errors='coerce'
                )
        
        print("   ✅ Date columns converted")
        
        # 4. Split paid vs outstanding
        self.df_paid_invoices = self.df_invoice[self.df_invoice['Last Payment Date'].notna()].copy()
        self.df_outstanding_invoices = self.df_invoice[self.df_invoice['Last Payment Date'].isna()].copy()
        
        print(f"   💰 Paid invoices: {len(self.df_paid_invoices)}")
        print(f"   ⏳ Outstanding invoices: {len(self.df_outstanding_invoices)}")
        
        # 5. Calculate aging delays
        if len(self.df_paid_invoices) > 0:
            self.df_paid_invoices['Aging_Delay'] = (
                self.df_paid_invoices['Last Payment Date'] - 
                self.df_paid_invoices['Due Date']
            ).dt.days
            
            self.df_paid_invoices['Aging_Bucket'] = pd.cut(
                self.df_paid_invoices['Aging_Delay'],
                bins=[-float('inf'), 0, 30, 60, 90, float('inf')],
                labels=['On-time', '1-30 Days', '31-60 Days', '61-90 Days', '90+ Days']
            )
        
        if len(self.df_outstanding_invoices) > 0:
            dataset_date = datetime(2025, 6, 23)
            self.df_outstanding_invoices['Aging_Delay'] = (
                dataset_date - self.df_outstanding_invoices['Due Date']
            ).dt.days
            
            self.df_outstanding_invoices['Aging_Bucket'] = pd.cut(
                self.df_outstanding_invoices['Aging_Delay'],
                bins=[-float('inf'), 0, 30, 60, 90, float('inf')],
                labels=['Current', '1-30 Days', '31-60 Days', '61-90 Days', '90+ Days']
            )
        
        print("✅ Data preparation completed successfully!")

    def paid_invoices_aging_analysis(self):
        """Analyze aging for PAID invoices"""
        print("\n" + "="*60)
        print("💰 PAID INVOICES AGING ANALYSIS")
        print("(How late were payments when received)")
        print("="*60)
        
        if len(self.df_paid_invoices) == 0:
            print("No paid invoices found.")
            return None
        
        # Aging summary
        paid_aging_summary = self.df_paid_invoices.groupby('Aging_Bucket').agg({
            'Number': 'count',
            'Amount (USD)': 'sum',
            'Aging_Delay': ['mean', 'max'],
        }).round(2)
        
        paid_aging_summary.columns = ['Invoice_Count', 'Total_Amount', 'Avg_Days_Late', 'Max_Days_Late']
        paid_aging_summary['Percentage'] = (
            paid_aging_summary['Invoice_Count'] / len(self.df_paid_invoices) * 100
        ).round(2)
        
        print("\nPaid Invoices Aging Summary:")
        print(paid_aging_summary)
        
        # Payment behavior insights
        early_payments = len(self.df_paid_invoices[self.df_paid_invoices['Aging_Delay'] <= 0])
        late_payments = len(self.df_paid_invoices[self.df_paid_invoices['Aging_Delay'] > 0])
        avg_delay = self.df_paid_invoices['Aging_Delay'].mean()
        
        print(f"\n💡 PAYMENT BEHAVIOR:")
        print(f"   Early/On-time: {early_payments} ({early_payments/len(self.df_paid_invoices)*100:.1f}%)")
        print(f"   Late payments: {late_payments} ({late_payments/len(self.df_paid_invoices)*100:.1f}%)")
        print(f"   Average delay: {avg_delay:.1f} days")
        
        # Visualization
        plt.figure(figsize=(15, 6))
        
        plt.subplot(1, 3, 1)
        paid_aging_summary['Invoice_Count'].plot(kind='bar', color='lightblue')
        plt.title('Paid Invoices by Aging')
        plt.ylabel('Number of Invoices')
        plt.xticks(rotation=45)
        
        plt.subplot(1, 3, 2)
        paid_aging_summary['Total_Amount'].plot(kind='bar', color='lightgreen')
        plt.title('Payment Amount by Aging')
        plt.ylabel('Total Amount (USD)')
        plt.xticks(rotation=45)
        
        plt.subplot(1, 3, 3)
        self.df_paid_invoices['Aging_Delay'].hist(bins=20, alpha=0.7, color='orange')
        plt.title('Distribution of Payment Delays')
        plt.xlabel('Days (negative = early, positive = late)')
        plt.ylabel('Frequency')
        plt.axvline(x=0, color='red', linestyle='--', label='Due Date')
        plt.legend()
        
        plt.tight_layout()
        plt.show()
        
        return paid_aging_summary

    def outstanding_invoices_aging_analysis(self):
        """Analyze aging for OUTSTANDING invoices"""
        print("\n" + "="*60)
        print("⏳ OUTSTANDING INVOICES AGING ANALYSIS")
        print("(Current collection priorities as of June 23, 2025)")
        print("="*60)
        
        if len(self.df_outstanding_invoices) == 0:
            print("No outstanding invoices found.")
            return None
        
        # Aging summary
        outstanding_aging_summary = self.df_outstanding_invoices.groupby('Aging_Bucket').agg({
            'Number': 'count',
            'Amount (USD)': 'sum',
            'Amt. Due (USD)': 'sum',
            'Aging_Delay': ['mean', 'max'],
        }).round(2)
        
        outstanding_aging_summary.columns = ['Invoice_Count', 'Total_Amount', 'Total_Due', 'Avg_Days_Overdue', 'Max_Days_Overdue']
        outstanding_aging_summary['Percentage'] = (
            outstanding_aging_summary['Invoice_Count'] / len(self.df_outstanding_invoices) * 100
        ).round(2)
        
        print("\nOutstanding Invoices Aging Summary:")
        print(outstanding_aging_summary)
        
        # Collection priority insights
        total_outstanding = outstanding_aging_summary['Total_Due'].sum()
        current_amount = outstanding_aging_summary.loc['Current', 'Total_Due'] if 'Current' in outstanding_aging_summary.index else 0
        overdue_amount = total_outstanding - current_amount
        
        print(f"\n🚨 COLLECTION PRIORITIES:")
        print(f"   Total Outstanding: ${total_outstanding:,.2f}")
        print(f"   Current (not yet due): ${current_amount:,.2f}")
        print(f"   Overdue amount: ${overdue_amount:,.2f} ({overdue_amount/total_outstanding*100:.1f}%)")
        
        # Visualization
        plt.figure(figsize=(15, 6))
        
        plt.subplot(1, 3, 1)
        outstanding_aging_summary['Invoice_Count'].plot(kind='bar', color='lightcoral')
        plt.title('Outstanding Invoices by Aging')
        plt.ylabel('Number of Invoices')
        plt.xticks(rotation=45)
        
        plt.subplot(1, 3, 2)
        outstanding_aging_summary['Total_Due'].plot(kind='bar', color='gold')
        plt.title('Amount Due by Aging')
        plt.ylabel('Amount Due (USD)')
        plt.xticks(rotation=45)
        
        plt.subplot(1, 3, 3)
        plt.pie(outstanding_aging_summary['Total_Due'], 
                labels=outstanding_aging_summary.index, 
                autopct='%1.1f%%')
        plt.title('Outstanding Amount Distribution')
        
        plt.tight_layout()
        plt.show()
        
        return outstanding_aging_summary

    def executive_summary(self):
        """Generate Executive Summary"""
        print("\n" + "="*60)
        print("📋 EXECUTIVE SUMMARY")
        print("="*60)
        
        # Portfolio metrics
        total_invoices = len(self.df_invoice)
        total_billed = self.df_invoice['Amount (USD)'].sum()
        total_paid = self.df_invoice['Amt. Paid (USD)'].sum()
        total_outstanding = self.df_outstanding_invoices['Amt. Due (USD)'].sum() if len(self.df_outstanding_invoices) > 0 else 0
        
        collection_rate = (total_paid / total_billed * 100) if total_billed > 0 else 0
        
        print(f"Portfolio Overview:")
        print(f"  Total Invoices: {total_invoices}")
        print(f"  Total Billed: ${total_billed:,.2f}")
        print(f"  Total Collected: ${total_paid:,.2f}")
        print(f"  Collection Rate: {collection_rate:.1f}%")
        print(f"  Outstanding Amount: ${total_outstanding:,.2f}")
        
        # Payment behavior
        avg_payment_delay = 0
        if len(self.df_paid_invoices) > 0:
            avg_payment_delay = self.df_paid_invoices['Aging_Delay'].mean()
            on_time_rate = len(self.df_paid_invoices[self.df_paid_invoices['Aging_Delay'] <= 0]) / len(self.df_paid_invoices) * 100
            
            print(f"\nPayment Behavior:")
            print(f"  Average Payment Delay: {avg_payment_delay:.1f} days")
            print(f"  On-time Payment Rate: {on_time_rate:.1f}%")
        
        # Risk assessment
        risk_percentage = 0
        if len(self.df_outstanding_invoices) > 0:
            overdue_90_plus = self.df_outstanding_invoices[self.df_outstanding_invoices['Aging_Delay'] > 90]
            high_risk_amount = overdue_90_plus['Amt. Due (USD)'].sum()
            risk_percentage = (high_risk_amount / total_outstanding * 100) if total_outstanding > 0 else 0
            
            print(f"\nCurrent Collection Risks:")
            print(f"  90+ Days Overdue: {len(overdue_90_plus)} invoices")
            print(f"  High Risk Amount: ${high_risk_amount:,.2f} ({risk_percentage:.1f}%)")
        
        # Recommendations
        print(f"\n📌 KEY RECOMMENDATIONS:")
        if avg_payment_delay > 15:
            print("• Review credit terms - payment delays are elevated")
        if risk_percentage > 15:
            print("• URGENT: High concentration of 90+ day receivables")
        if collection_rate < 85:
            print("• Implement more aggressive collection procedures")
        if collection_rate > 90:
            print("• Excellent collection performance - maintain current procedures")
        
        return {
            'total_outstanding': total_outstanding,
            'collection_rate': collection_rate,
            'avg_payment_delay': avg_payment_delay,
            'risk_percentage': risk_percentage
        }

    def run_full_analysis(self):
        """Run complete factoring analysis"""
        print("🏢 CORRECTED FACTORING FIRM ANALYSIS")
        print("="*60)
        
        # Run all analyses
        paid_aging = self.paid_invoices_aging_analysis()
        outstanding_aging = self.outstanding_invoices_aging_analysis()
        executive_summary = self.executive_summary()
        
        return {
            'paid_invoices_aging': paid_aging,
            'outstanding_invoices_aging': outstanding_aging,
            'executive_summary': executive_summary
        }

print("✅ Complete CorrectedFactoringAnalyzer class loaded successfully!")
print("🚀 Ready to run full analysis!")
# 🚀 INITIALIZE AND RUN ANALYSIS
print("🏢 INITIALIZING FACTORING ANALYZER...")

# Initialize the analyzer
analyzer = CorrectedFactoringAnalyzer('../data/TecnoCargoInvoiceDataset01.csv', encoding='latin1')

print("\n🚀 RUNNING COMPLETE ANALYSIS...")
# Run the complete analysis
results = analyzer.run_full_analysis()

print("\n🎉 ANALYSIS COMPLETED!")
