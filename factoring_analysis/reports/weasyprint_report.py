"""
Professional PDF Report Generator using WeasyPrint
Adapted for factoring_analysis project structure
File: reports/weasyprint_report.py
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.graph_objects as go
import plotly.express as px
from plotly.offline import plot
import base64
from io import BytesIO
from datetime import datetime
import os
import sys
from pathlib import Path

# Add project root to path for imports
project_root = Path(__file__).parent.parent
sys.path.append(str(project_root))

try:
    from weasyprint import HTML, CSS
    from jinja2 import Template
except ImportError as e:
    print(f"❌ Missing required package: {e}")
    print("Please install with: pip install weasyprint jinja2")
    sys.exit(1)

import warnings
warnings.filterwarnings('ignore')

class ProfessionalFactoringReport:
    def __init__(self, analyzer, results):
        """
        Initialize the professional report generator
        
        Parameters:
        analyzer: CorrectedFactoringAnalyzer instance
        results: Results dictionary from run_full_analysis()
        """
        self.analyzer = analyzer
        self.results = results
        
        # Set up paths relative to project root
        self.project_root = Path(__file__).parent.parent
        self.output_dir = self.project_root / 'output'
        self.charts_dir = self.output_dir / 'charts'
        self.reports_dir = self.output_dir / 'reports'
        
        # Create directories if they don't exist
        self.charts_dir.mkdir(parents=True, exist_ok=True)
        self.reports_dir.mkdir(parents=True, exist_ok=True)
        
        print(f"📁 Charts directory: {self.charts_dir}")
        print(f"📁 Reports directory: {self.reports_dir}")
        
    def generate_chart_base64(self, fig):
        """Convert matplotlib figure to base64 string for embedding"""
        buffer = BytesIO()
        fig.savefig(buffer, format='png', dpi=300, bbox_inches='tight', 
                   facecolor='white', edgecolor='none')
        buffer.seek(0)
        image_base64 = base64.b64encode(buffer.getvalue()).decode()
        plt.close(fig)
        return image_base64
    
    def create_collection_rate_chart(self):
        """Create collection rate pie chart"""
        exec_summary = self.results['executive_summary']
        
        fig, ax = plt.subplots(figsize=(8, 6))
        fig.patch.set_facecolor('white')
        
        colors = ['#2E8B57', '#DC143C']  # Professional green and red
        sizes = [exec_summary['collection_rate'], 100 - exec_summary['collection_rate']]
        labels = ['Collected', 'Outstanding']
        
        wedges, texts, autotexts = ax.pie(sizes, labels=labels, autopct='%1.1f%%', 
                                         colors=colors, startangle=90,
                                         explode=(0.05, 0))  # Slight separation
        
        # Style the text
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
            autotext.set_fontsize(14)
        
        for text in texts:
            text.set_fontsize(12)
            text.set_fontweight('bold')
        
        ax.set_title('Collection Rate Performance', fontsize=16, fontweight='bold', pad=20)
        return self.generate_chart_base64(fig)
    
    def create_aging_analysis_chart(self):
        """Create outstanding invoices aging chart"""
        if self.results['outstanding_invoices_aging'] is None:
            return None
            
        outstanding_aging = self.results['outstanding_invoices_aging']
        
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))
        fig.patch.set_facecolor('white')
        
        # Chart 1: Invoice count by aging
        colors = ['#228B22', '#FFD700', '#FF8C00', '#DC143C', '#8B0000']
        bars1 = ax1.bar(outstanding_aging.index, outstanding_aging['Invoice_Count'], color=colors)
        ax1.set_title('Outstanding Invoices by Age', fontweight='bold', fontsize=14)
        ax1.set_ylabel('Number of Invoices', fontweight='bold')
        ax1.tick_params(axis='x', rotation=45)
        ax1.grid(True, alpha=0.3, axis='y')
        
        # Add value labels on bars
        for bar in bars1:
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height + height*0.01,
                    f'{int(height)}', ha='center', va='bottom', fontweight='bold')
        
        # Chart 2: Amount due by aging
        bars2 = ax2.bar(outstanding_aging.index, outstanding_aging['Total_Due'], color=colors)
        ax2.set_title('Amount Due by Age', fontweight='bold', fontsize=14)
        ax2.set_ylabel('Amount Due (USD)', fontweight='bold')
        ax2.tick_params(axis='x', rotation=45)
        ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.0f}'))
        ax2.grid(True, alpha=0.3, axis='y')
        
        # Add value labels on bars
        for bar in bars2:
            height = bar.get_height()
            ax2.text(bar.get_x() + bar.get_width()/2., height + height*0.01,
                    f'${height:,.0f}', ha='center', va='bottom', fontweight='bold', fontsize=10)
        
        plt.tight_layout()
        return self.generate_chart_base64(fig)
    
    def create_payment_delay_chart(self):
        """Create payment delay distribution chart"""
        if len(self.analyzer.df_paid_invoices) == 0:
            return None
            
        fig, ax = plt.subplots(figsize=(12, 6))
        fig.patch.set_facecolor('white')
        
        # Create histogram of payment delays
        delays = self.analyzer.df_paid_invoices['Aging_Delay']
        n, bins, patches = ax.hist(delays, bins=30, alpha=0.7, color='#4682B4', 
                                  edgecolor='black', linewidth=0.5)
        
        # Color bars based on delay (green for early, red for late)
        for i, patch in enumerate(patches):
            bin_center = (bins[i] + bins[i+1]) / 2
            if bin_center <= 0:
                patch.set_facecolor('#28a745')  # Green for early/on-time
            elif bin_center <= 30:
                patch.set_facecolor('#ffc107')  # Yellow for slightly late
            else:
                patch.set_facecolor('#dc3545')  # Red for very late
        
        ax.axvline(x=0, color='red', linestyle='--', linewidth=3, label='Due Date')
        ax.axvline(x=delays.mean(), color='orange', linestyle='-', linewidth=3, 
                  label=f'Average ({delays.mean():.1f} days)')
        
        ax.set_title('Payment Delay Distribution', fontweight='bold', fontsize=16)
        ax.set_xlabel('Days (negative = early, positive = late)', fontweight='bold')
        ax.set_ylabel('Number of Invoices', fontweight='bold')
        ax.legend(fontsize=12)
        ax.grid(True, alpha=0.3)
        
        return self.generate_chart_base64(fig)
    
    def create_client_analysis_chart(self):
        """Create top clients analysis chart"""
        top_clients = self.analyzer.df_invoice.groupby('Applied to')['Amount (USD)'].sum().sort_values(ascending=False).head(10)
        
        fig, ax = plt.subplots(figsize=(12, 8))
        fig.patch.set_facecolor('white')
        
        # Create color gradient
        colors = plt.cm.viridis(np.linspace(0.3, 0.9, len(top_clients)))
        
        bars = ax.barh(range(len(top_clients)), top_clients.values, color=colors)
        ax.set_yticks(range(len(top_clients)))
        ax.set_yticklabels([client[:30] + '...' if len(client) > 30 else client 
                           for client in top_clients.index], fontsize=10)
        ax.set_xlabel('Total Amount (USD)', fontsize=12, fontweight='bold')
        ax.set_title('Top 10 Clients by Total Amount', fontweight='bold', fontsize=16)
        ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.0f}'))
        
        # Add value labels on bars
        for i, (bar, value) in enumerate(zip(bars, top_clients.values)):
            ax.text(value + max(top_clients.values) * 0.01, i, f'${value:,.0f}', 
                   va='center', fontweight='bold', fontsize=10)
        
        ax.grid(True, alpha=0.3, axis='x')
        plt.tight_layout()
        return self.generate_chart_base64(fig)
    
    def get_css_styles(self):
        """Return professional CSS styles"""
        return """
        @page {
            size: A4;
            margin: 0.8in;
            @top-center {
                content: "Factoring Risk Analysis Report - " string(report-title);
                font-family: Arial, sans-serif;
                font-size: 10pt;
                color: #666;
                border-bottom: 1px solid #ddd;
                padding-bottom: 5px;
            }
            @bottom-center {
                content: "Page " counter(page) " of " counter(pages) " | Generated: " string(generation-date);
                font-family: Arial, sans-serif;
                font-size: 9pt;
                color: #666;
                border-top: 1px solid #ddd;
                padding-top: 5px;
            }
        }
        
        body {
            font-family: 'Arial', 'Helvetica', sans-serif;
            line-height: 1.6;
            color: #333;
            background-color: #fff;
            margin: 0;
            padding: 0;
        }
        
        .header {
            text-align: center;
            border-bottom: 4px solid #2E8B57;
            padding-bottom: 25px;
            margin-bottom: 35px;
            background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        
        .header h1 {
            font-size: 32pt;
            color: #2E8B57;
            margin: 0;
            font-weight: bold;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
            string-set: report-title "Factoring Risk Analysis";
        }
        
        .header h2 {
            font-size: 18pt;
            color: #666;
            margin: 15px 0;
            font-style: italic;
            font-weight: 300;
        }
        
        .summary-box {
            background: linear-gradient(135deg, #f8f9fa 0%, #e8f5e8 100%);
            border: 2px solid #2E8B57;
            border-radius: 12px;
            padding: 25px;
            margin: 25px 0;
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        }
        
        .risk-indicator {
            text-align: center;
            padding: 20px;
            border-radius: 12px;
            margin: 25px 0;
            font-weight: bold;
            font-size: 20pt;
            box-shadow: 0 4px 8px rgba(0,0,0,0.2);
            text-transform: uppercase;
            letter-spacing: 2px;
        }
        
        .risk-low { 
            background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%); 
            color: #155724; 
            border: 3px solid #28a745; 
        }
        .risk-moderate { 
            background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%); 
            color: #856404; 
            border: 3px solid #ffc107; 
        }
        .risk-high { 
            background: linear-gradient(135deg, #f8d7da 0%, #f5c6cb 100%); 
            color: #721c24; 
            border: 3px solid #dc3545; 
        }
        
        .section {
            margin: 35px 0;
            page-break-inside: avoid;
        }
        
        .section h2 {
            color: #2E8B57;
            border-bottom: 3px solid #2E8B57;
            padding-bottom: 8px;
            font-size: 22pt;
            margin-bottom: 25px;
            string-set: generation-date attr(data-date);
        }
        
        .section h3 {
            color: #4682B4;
            font-size: 16pt;
            margin: 25px 0 15px 0;
            border-left: 4px solid #4682B4;
            padding-left: 15px;
        }
        
        .metrics-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin: 25px 0;
        }
        
        .metric-box {
            background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
            border: 2px solid #dee2e6;
            border-radius: 8px;
            padding: 20px;
            text-align: center;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            transition: transform 0.3s ease;
        }
        
        .metric-box:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.15);
        }
        
        .metric-value {
            font-size: 28pt;
            font-weight: bold;
            color: #2E8B57;
            display: block;
            margin-bottom: 5px;
        }
        
        .metric-label {
            font-size: 12pt;
            color: #666;
            font-weight: 600;
        }
        
        .chart-container {
            text-align: center;
            margin: 25px 0;
            page-break-inside: avoid;
            background: #fff;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .chart-container img {
            max-width: 100%;
            height: auto;
            border: 1px solid #ddd;
            border-radius: 6px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 25px 0;
            background: white;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        
        th {
            background: linear-gradient(135deg, #2E8B57 0%, #228B22 100%);
            color: white;
            padding: 15px 12px;
            text-align: left;
            font-weight: bold;
            font-size: 11pt;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        td {
            padding: 12px;
            border-bottom: 1px solid #dee2e6;
            font-size: 10pt;
        }
        
        tr:nth-child(even) {
            background: #f8f9fa;
        }
        
        tr:hover {
            background: #e8f5e8;
            transition: background-color 0.3s ease;
        }
        
        .recommendations {
            background: linear-gradient(135deg, #e8f5e8 0%, #d4edda 100%);
            border-left: 6px solid #2E8B57;
            padding: 25px;
            margin: 25px 0;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .recommendations h3 {
            color: #2E8B57;
            margin-top: 0;
            border: none;
            padding: 0;
        }
        
        .recommendations h4 {
            margin: 20px 0 10px 0;
            font-size: 14pt;
        }
        
        .recommendations ul {
            margin: 15px 0;
            padding-left: 25px;
        }
        
        .recommendations li {
            margin: 10px 0;
            line-height: 1.6;
        }
        
        .priority-high { 
            color: #dc3545; 
            font-weight: bold; 
            background: #f8d7da;
            padding: 2px 6px;
            border-radius: 4px;
        }
        .priority-medium { 
            color: #fd7e14; 
            font-weight: bold; 
            background: #fff3cd;
            padding: 2px 6px;
            border-radius: 4px;
        }
        .priority-low { 
            color: #28a745; 
            font-weight: bold; 
            background: #d4edda;
            padding: 2px 6px;
            border-radius: 4px;
        }
        
        .page-break { 
            page-break-before: always; 
        }
        
        .footer {
            text-align: center;
            font-size: 10pt;
            color: #666;
            margin-top: 50px;
            padding-top: 25px;
            border-top: 2px solid #ddd;
            background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
            padding: 25px;
            border-radius: 8px;
        }
        
        .text-right { text-align: right; }
        .text-center { text-align: center; }
        .font-bold { font-weight: bold; }
        .text-success { color: #28a745; font-weight: bold; }
        .text-warning { color: #fd7e14; font-weight: bold; }
        .text-danger { color: #dc3545; font-weight: bold; }
        
        /* Ensure proper spacing and avoid orphans */
        p, li { orphans: 2; widows: 2; }
        h1, h2, h3, h4 { page-break-after: avoid; }
        """
    
    def get_html_template(self):
        """Return the HTML template with improved structure"""
        return """
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Factoring Risk Analysis Report</title>
        </head>
        <body data-date="{{ generation_date }}">
            <!-- Cover Page -->
            <div class="header">
                <h1>FACTORING RISK ANALYSIS</h1>
                <h2>Invoice Portfolio Assessment Report</h2>
                <div class="summary-box">
                    <div class="metrics-grid">
                        <div class="metric-box">
                            <span class="metric-value">{{ total_invoices }}</span>
                            <div class="metric-label">Total Invoices</div>
                        </div>
                        <div class="metric-box">
                            <span class="metric-value">{{ collection_rate }}%</span>
                            <div class="metric-label">Collection Rate</div>
                        </div>
                        <div class="metric-box">
                            <span class="metric-value">{{ avg_payment_delay }} days</span>
                            <div class="metric-label">Avg Payment Delay</div>
                        </div>
                        <div class="metric-box">
                            <span class="metric-value">${{ total_outstanding }}</span>
                            <div class="metric-label">Total Outstanding</div>
                        </div>
                    </div>
                </div>
                
                <div class="risk-indicator {{ risk_class }}">
                    Overall Risk Level: {{ risk_level }}
                </div>
                
                <p style="font-size: 12pt; color: #666; margin-top: 20px;">
                    <strong>Analysis Period:</strong> January 2024 - June 2025<br>
                    <strong>Report Date:</strong> {{ report_date }}<br>
                    <strong>Data Snapshot:</strong> June 23, 2025
                </p>
            </div>
            
            <!-- Executive Summary -->
            <div class="section">
                <h2>Executive Summary</h2>
                
                {% if collection_chart %}
                <div class="chart-container">
                    <h3>Collection Rate Performance</h3>
                    <img src="data:image/png;base64,{{ collection_chart }}" alt="Collection Rate Chart">
                </div>
                {% endif %}
                
                <div class="summary-box">
                    <h3>Key Findings & Insights</h3>
                    <p><strong>Portfolio Health:</strong> {{ portfolio_health }}</p>
                    <ul>
                        <li>Collection rate of {{ collection_rate }}% {{ collection_performance }}</li>
                        <li>Average payment delay of {{ avg_payment_delay }} days is {{ delay_assessment }}</li>
                        <li>{{ risk_percentage }}% of outstanding amount is 90+ days overdue</li>
                    </ul>
                    
                    <p><strong>Risk Assessment:</strong> {{ risk_assessment }}</p>
                    <ul>
                        <li>Total outstanding amount: ${{ total_outstanding }}</li>
                        <li>Portfolio size: {{ total_invoices }} total invoices</li>
                        <li>Paid invoices: {{ paid_count }} ({{ paid_percentage }}%)</li>
                        <li>Outstanding invoices: {{ outstanding_count }} ({{ outstanding_percentage }}%)</li>
                    </ul>
                </div>
            </div>
            
            <div class="page-break"></div>
            
            <!-- Portfolio Overview -->
            <div class="section">
                <h2>Portfolio Overview</h2>
                
                {% if payment_delay_chart %}
                <div class="chart-container">
                    <h3>Payment Delay Distribution Analysis</h3>
                    <img src="data:image/png;base64,{{ payment_delay_chart }}" alt="Payment Delay Distribution">
                </div>
                {% endif %}
                
                <h3>Portfolio Summary Statistics</h3>
                <table>
                    <thead>
                        <tr>
                            <th>Metric</th>
                            <th>Value</th>
                            <th>Performance</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td>Total Invoices</td>
                            <td>{{ total_invoices }}</td>
                            <td class="text-center">📊</td>
                        </tr>
                        <tr>
                            <td>Total Billed Amount</td>
                            <td>${{ total_billed }}</td>
                            <td class="text-center">💰</td>
                        </tr>
                        <tr>
                            <td>Total Collected</td>
                            <td>${{ total_collected }}</td>
                            <td class="text-success">✓ Collected</td>
                        </tr>
                        <tr>
                            <td>Total Outstanding</td>
                            <td>${{ total_outstanding }}</td>
                            <td class="text-warning">⏳ Pending</td>
                        </tr>
                        <tr>
                            <td>Average Invoice Amount</td>
                            <td>${{ avg_invoice_amount }}</td>
                            <td class="text-center">📈</td>
                        </tr>
                        <tr>
                            <td>Collection Rate</td>
                            <td>{{ collection_rate }}%</td>
                            <td class="{{ 'text-success' if collection_rate|float > 90 else 'text-warning' if collection_rate|float > 75 else 'text-danger' }}">
                                {{ '✓ Excellent' if collection_rate|float > 90 else '△ Good' if collection_rate|float > 75 else '⚠ Needs Improvement' }}
                            </td>
                        </tr>
                    </tbody>
                </table>
            </div>
            
            <div class="page-break"></div>
            
            <!-- Outstanding Invoices Analysis -->
            <div class="section">
                <h2>Outstanding Invoices Analysis</h2>
                
                {% if aging_chart %}
                <div class="chart-container">
                    <h3>Outstanding Invoices by Age Category</h3>
                    <img src="data:image/png;base64,{{ aging_chart }}" alt="Aging Analysis Chart">
                </div>
                {% endif %}
                
                {% if outstanding_aging_table %}
                <h3>Detailed Outstanding Invoices Statistics</h3>
                <table>
                    <thead>
                        <tr>
                            <th>Age Category</th>
                            <th>Invoice Count</th>
                            <th>Total Amount</th>
                            <th>Total Due</th>
                            <th>Avg Days Overdue</th>
                            <th>Percentage</th>
                            <th>Risk Level</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for row in outstanding_aging_table %}
                        <tr>
                            <td class="font-bold">{{ row.index }}</td>
                            <td>{{ row.Invoice_Count }}</td>
                            <td>${{ row.Total_Amount }}</td>
                            <td>${{ row.Total_Due }}</td>
                            <td>{{ row.Avg_Days_Overdue }}</td>
                            <td>{{ row.Percentage }}%</td>
                            <td class="{{ 'text-success' if row.index == 'Current' else 'text-warning' if '1-30' in row.index else 'text-danger' }}">
                                {{ 'Low' if row.index == 'Current' else 'Medium' if '1-30' in row.index or '31-60' in row.index else 'High' }}
                            </td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
                {% endif %}
            </div>
            
            <div class="page-break"></div>
            
            <!-- Client Analysis -->
            <div class="section">
                <h2>Client Analysis</h2>
                
                {% if client_chart %}
                <div class="chart-container">
                    <h3>Top 10 Clients by Total Amount</h3>
                    <img src="data:image/png;base64,{{ client_chart }}" alt="Top Clients Chart">
                </div>
                {% endif %}
                
                <h3>Client Performance Summary</h3>
                <table>
                    <thead>
                        <tr>
                            <th>Client</th>
                            <th>Invoice Count</th>
                            <th>Total Amount</th>
                            <th>Total Paid</th>
                            <th>Total Due</th>
                            <th>Collection Rate</th>
                            <th>Performance</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for row in top_clients_table %}
                        <tr>
                            <td class="font-bold">{{ row.client[:40] }}{{ '...' if row.client|length > 40 else '' }}</td>
                            <td>{{ row.invoice_count }}</td>
                            <td>${{ row.total_amount }}</td>
                            <td>${{ row.total_paid }}</td>
                            <td>${{ row.total_due }}</td>
                            <td class="{{ row.collection_class }}">{{ row.collection_rate }}%</td>
                            <td class="{{ row.collection_class }}">
                                {{ '⭐ Excellent' if row.collection_rate|float > 95 else '✓ Good' if row.collection_rate|float > 85 else '△ Fair' if row.collection_rate|float > 70 else '⚠ Poor' }}
                            </td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
            
            <div class="page-break"></div>
            
            <!-- Recommendations -->
            <div class="section">
                <h2>Recommendations & Action Plan</h2>
                
                <div class="recommendations">
                    <h3>🎯 Priority Actions</h3>
                    
                    <h4 class="priority-high">🚨 HIGH PRIORITY (0-30 days):</h4>
                    <ul>
                        <li>Focus collection efforts on 90+ day receivables immediately</li>
                        <li>Contact top 10 outstanding clients for payment plans</li>
                        <li>Implement daily monitoring for high-risk accounts</li>
                        {% if risk_level == 'HIGH' %}
                        <li class="priority-high">URGENT: Deploy emergency collection procedures</li>
                        <li class="priority-high">Review credit limits for all new transactions</li>
                        {% endif %}
                    </ul>
                    
                    <h4 class="priority-medium">⚠️ MEDIUM PRIORITY (30-90 days):</h4>
                    <ul>
                        <li>Implement weekly aging reports and automated alerts</li>
                        <li>Establish client payment performance scorecards</li>
                        <li>Consider offering early payment discounts (2/10 net 30)</li>
                        <li>Develop client risk profiles based on payment history</li>
                        {% if avg_payment_delay|float > 30 %}
                        <li class="priority-medium">Review and tighten credit policies</li>
                        <li class="priority-medium">Implement stricter approval processes</li>
                        {% endif %}
                    </ul>
                    
                    <h4 class="priority-low">✅ LONG-TERM STRATEGY (90+ days):</h4>
                    <ul>
                        <li>Implement monthly portfolio risk assessments</li>
                        <li>Develop predictive analytics for collection timing</li>
                        <li>Client credit limit reviews based on payment history</li>
                        <li>Implement automated collection procedures and workflows</li>
                        <li>Annual comprehensive credit policy review</li>
                        <li>Staff training on collection best practices</li>
                    </ul>
                </div>
                
                <div class="summary-box">
                    <h3>📊 Success Metrics & KPIs</h3>
                    <p><strong>Progress will be measured by:</strong></p>
                    <ul>
                        <li><strong>Collection Rate:</strong> Target > 95% (currently {{ collection_rate }}%)</li>
                        <li><strong>90+ Day Receivables:</strong> Target < 5% (currently {{ risk_percentage }}%)</li>
                        <li><strong>Average Payment Delay:</strong> Target < 15 days (currently {{ avg_payment_delay }} days)</li>
                        <li><strong>Outstanding Ratio:</strong> Target < 15% (currently {{ outstanding_percentage }}%)</li>
                    </ul>
                    
                    <p><strong>Review Schedule:</strong></p>
                    <ul>
                        <li>Daily: High-risk account monitoring</li>
                        <li>Weekly: Aging reports and collection activities</li>
                        <li>Monthly: Portfolio risk assessment</li>
                        <li>Quarterly: Strategy review and policy updates</li>
                    </ul>
                </div>
            </div>
            
            <div class="footer">
                <p><strong>CONFIDENTIAL - For Internal Use Only</strong><br>
                Generated by Professional Factoring Analysis System<br>
                Report Version: 2.0 | Generated: {{ generation_date }}</p>
            </div>
        </body>
        </html>
        """
    
    def prepare_template_data(self):
        """Prepare all data for the HTML template"""
        exec_summary = self.results['executive_summary']
        
        # Generate charts
        print("   📊 Generating collection rate chart...")
        collection_chart = self.create_collection_rate_chart()
        
        print("   📊 Generating aging analysis chart...")
        aging_chart = self.create_aging_analysis_chart()
        
        print("   📊 Generating payment delay chart...")
        payment_delay_chart = self.create_payment_delay_chart()
        
        print("   📊 Generating client analysis chart...")
        client_chart = self.create_client_analysis_chart()
        
        # Risk assessment
        risk_level = "LOW" if exec_summary['risk_percentage'] < 10 else "MODERATE" if exec_summary['risk_percentage'] < 20 else "HIGH"
        risk_class = f"risk-{risk_level.lower()}"
        
        # Portfolio health assessment
        portfolio_health = 'EXCELLENT' if exec_summary['collection_rate'] > 90 else 'GOOD' if exec_summary['collection_rate'] > 80 else 'NEEDS IMPROVEMENT'
        
        # Prepare outstanding aging table
        outstanding_aging_table = []
        if self.results['outstanding_invoices_aging'] is not None:
            outstanding_aging = self.results['outstanding_invoices_aging'].reset_index()
            for _, row in outstanding_aging.iterrows():
                outstanding_aging_table.append({
                    'index': row.iloc[0],  # The index column (age category)
                    'Invoice_Count': f"{row['Invoice_Count']:,}",
                    'Total_Amount': f"{row['Total_Amount']:,.2f}",
                    'Total_Due': f"{row['Total_Due']:,.2f}",
                    'Avg_Days_Overdue': f"{row['Avg_Days_Overdue']:.1f}",
                    'Percentage': f"{row['Percentage']:.1f}"
                })
        
        # Prepare top clients table
        client_stats = self.analyzer.df_invoice.groupby('Applied to').agg({
            'Number': 'count',
            'Amount (USD)': ['sum', 'mean'],
            'Amt. Paid (USD)': 'sum',
            'Amt. Due (USD)': 'sum'
        }).round(2)
        
        client_stats.columns = ['Invoice_Count', 'Total_Amount', 'Avg_Amount', 'Total_Paid', 'Total_Due']
        client_stats['Collection_Rate'] = (client_stats['Total_Paid'] / client_stats['Total_Amount'] * 100).round(1)
        
        top_clients_table = []
        top_clients = client_stats.sort_values('Total_Amount', ascending=False).head(10)
        
        for client, row in top_clients.iterrows():
            collection_rate = row['Collection_Rate']
            collection_class = 'text-success' if collection_rate > 90 else 'text-warning' if collection_rate > 75 else 'text-danger'
            
            top_clients_table.append({
                'client': client,
                'invoice_count': f"{int(row['Invoice_Count']):,}",
                'total_amount': f"{row['Total_Amount']:,.2f}",
                'total_paid': f"{row['Total_Paid']:,.2f}",
                'total_due': f"{row['Total_Due']:,.2f}",
                'collection_rate': f"{collection_rate:.1f}",
                'collection_class': collection_class
            })
        
        return {
            # Basic metrics
            'total_invoices': f"{len(self.analyzer.df_invoice):,}",
            'collection_rate': f"{exec_summary['collection_rate']:.1f}",
            'avg_payment_delay': f"{exec_summary['avg_payment_delay']:.1f}",
            'total_outstanding': f"{exec_summary['total_outstanding']:,.2f}",
            'risk_percentage': f"{exec_summary['risk_percentage']:.1f}",
            'risk_level': risk_level,
            'risk_class': risk_class,
            'portfolio_health': portfolio_health,
            'report_date': datetime.now().strftime('%B %d, %Y'),
            'generation_date': datetime.now().strftime('%Y-%m-%d %H:%M'),
            
            # Detailed metrics
            'total_billed': f"{self.analyzer.df_invoice['Amount (USD)'].sum():,.2f}",
            'total_collected': f"{self.analyzer.df_invoice['Amt. Paid (USD)'].sum():,.2f}",
            'avg_invoice_amount': f"{self.analyzer.df_invoice['Amount (USD)'].mean():,.2f}",
            'paid_count': f"{len(self.analyzer.df_paid_invoices):,}",
            'outstanding_count': f"{len(self.analyzer.df_outstanding_invoices):,}",
            'paid_percentage': f"{len(self.analyzer.df_paid_invoices)/len(self.analyzer.df_invoice)*100:.1f}",
            'outstanding_percentage': f"{len(self.analyzer.df_outstanding_invoices)/len(self.analyzer.df_invoice)*100:.1f}",
            
            # Assessments
            'collection_performance': 'exceeds industry standards' if exec_summary['collection_rate'] > 90 else 'meets expectations' if exec_summary['collection_rate'] > 80 else 'falls below targets',
            'delay_assessment': 'excellent' if exec_summary['avg_payment_delay'] < 15 else 'acceptable' if exec_summary['avg_payment_delay'] < 30 else 'concerning',
            'risk_assessment': f"{exec_summary['risk_percentage']:.1f}% of outstanding amount requires immediate attention",
            
            # Charts
            'collection_chart': collection_chart,
            'aging_chart': aging_chart,
            'payment_delay_chart': payment_delay_chart,
            'client_chart': client_chart,
            
            # Tables
            'outstanding_aging_table': outstanding_aging_table,
            'top_clients_table': top_clients_table
        }
    
    def generate_report(self, output_filename='Professional_Factoring_Report.pdf'):
        """Generate the professional PDF report"""
        print("📄 GENERATING PROFESSIONAL WEASYPRINT REPORT")
        print("="*60)
        
        # Prepare template data
        print("   📊 Preparing charts and data...")
        template_data = self.prepare_template_data()
        
        # Render HTML template
        print("   🎨 Rendering HTML template...")
        template = Template(self.get_html_template())
        html_content = template.render(**template_data)
        
        # Generate PDF
        print("   📄 Generating PDF...")
        output_path = self.reports_dir / output_filename
        
        try:
            # Create WeasyPrint HTML and CSS objects
            html_doc = HTML(string=html_content)
            css_doc = CSS(string=self.get_css_styles())
            
            # Generate the PDF
            html_doc.write_pdf(str(output_path), stylesheets=[css_doc])
            
            print(f"✅ Professional PDF report generated: {output_path}")
            return str(output_path)
            
        except Exception as e:
            print(f"❌ Error generating PDF: {e}")
            
            # Save HTML for debugging
            html_debug_path = self.reports_dir / f"debug_{output_filename.replace('.pdf', '.html')}"
            with open(html_debug_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            print(f"💾 Debug HTML saved to: {html_debug_path}")
            
            raise e

def generate_professional_factoring_report(analyzer, results, filename='Professional_Factoring_Report.pdf'):
    """
    Main function to generate professional WeasyPrint report
    
    Usage:
    report_path = generate_professional_factoring_report(analyzer, results)
    """
    
    # Check for required packages
    try:
        import weasyprint
        from jinja2 import Template
    except ImportError as e:
        print(f"❌ Missing required package: {e}")
        print("Please install with: pip install weasyprint jinja2")
        return None
    
    # Generate the report
    report_generator = ProfessionalFactoringReport(analyzer, results)
    return report_generator.generate_report(filename)

if __name__ == "__main__":
    print("✅ Professional WeasyPrint Report Generator loaded successfully!")
    print("📄 Ready to generate professional reports with beautiful styling!")
    print(f"📁 Project root: {Path(__file__).parent.parent}")