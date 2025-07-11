# %%
import pandas as pd
import os

# Read the original CSV file (using relative path from notebooks folder)
df = pd.read_csv('../data/cvs/PopularAdditionsSubtractions.csv')

# Create deposits dataframe (rows with additions)
# Filter rows where Additions column is not null and not empty
deposits_mask = df['Additions'].notna() & (df['Additions'] != '') & (df['Additions'] != 0)
deposits_df = df[deposits_mask].copy()

# Create the deposits dataframe with required columns
deposits_output = pd.DataFrame({
    'date': deposits_df['Date'],
    'description': deposits_df['Description'], 
    'amount': deposits_df['Additions']
})

# Convert amount to float and round to 2 decimal places
deposits_output['amount'] = pd.to_numeric(deposits_output['amount'], errors='coerce').round(2)

# Create withdrawals dataframe (rows with subtractions)
# Filter rows where Subtractions column is not null and not empty
withdrawals_mask = df['Subtractions'].notna() & (df['Subtractions'] != '') & (df['Subtractions'] != 0)
withdrawals_df = df[withdrawals_mask].copy()

# Create the withdrawals dataframe with required columns
withdrawals_output = pd.DataFrame({
    'date': withdrawals_df['Date'],
    'description': withdrawals_df['Description'],
    'amount': withdrawals_df['Subtractions']
})

# Convert amount to float and round to 2 decimal places
withdrawals_output['amount'] = pd.to_numeric(withdrawals_output['amount'], errors='coerce').round(2)

# Remove any rows where amount conversion failed (resulted in NaN)
deposits_output = deposits_output.dropna(subset=['amount'])
withdrawals_output = withdrawals_output.dropna(subset=['amount'])

# Ensure the output directory exists (using relative path)
os.makedirs('../data/cvs', exist_ok=True)

# Save the deposits CSV with proper float formatting
deposits_output.to_csv('../data/cvs/PopularDeposits.csv', index=False, float_format='%.2f')

# Save the withdrawals CSV with proper float formatting
withdrawals_output.to_csv('../data/cvs/PopularWithdrawal.csv', index=False, float_format='%.2f')

print(f"Created PopularDeposits.csv with {len(deposits_output)} deposit records")
print(f"Created PopularWithdrawal.csv with {len(withdrawals_output)} withdrawal records")

# Display sample of each file for verification
print("\nSample from PopularDeposits.csv:")
print(deposits_output.head())
print(f"Amount column type: {deposits_output['amount'].dtype}")

print("\nSample from PopularWithdrawal.csv:")
print(withdrawals_output.head())
print(f"Amount column type: {withdrawals_output['amount'].dtype}")

# %%
import pandas as pd
import os

# Read the Wells Fargo CSV file (using relative path from notebooks folder)
df = pd.read_csv('../data/cvs/WellFargoAdditionsSubstractions.csv')

# Create deposits dataframe (rows with additions)
# Filter rows where Additions column is not null and not empty
deposits_mask = df['Additions'].notna() & (df['Additions'] != '') & (df['Additions'] != 0)
deposits_df = df[deposits_mask].copy()

# Create the deposits dataframe with required columns
deposits_output = pd.DataFrame({
    'date': deposits_df['Date'],
    'description': deposits_df['Description'], 
    'amount': deposits_df['Additions']
})

# Create withdrawals dataframe (rows with subtractions)
# Filter rows where Subtractions column is not null and not empty
withdrawals_mask = df['Subtractions'].notna() & (df['Subtractions'] != '') & (df['Subtractions'] != 0)
withdrawals_df = df[withdrawals_mask].copy()

# Create the withdrawals dataframe with required columns
withdrawals_output = pd.DataFrame({
    'date': withdrawals_df['Date'],
    'description': withdrawals_df['Description'],
    'amount': withdrawals_df['Subtractions']
})

# Ensure the output directory exists (using relative path)
os.makedirs('../data/cvs', exist_ok=True)

# Save the Wells Fargo deposits CSV
deposits_output.to_csv('../data/cvs/WellFargoDeposits.csv', index=False)

# Save the Wells Fargo withdrawals CSV  
withdrawals_output.to_csv('../data/cvs/WellFargoWithdrawal.csv', index=False)

print(f"Created WellFargoDeposits.csv with {len(deposits_output)} deposit records")
print(f"Created WellFargoWithdrawal.csv with {len(withdrawals_output)} withdrawal records")

# Display sample of each file for verification
print("\nSample from WellFargoDeposits.csv:")
print(deposits_output.head())

print("\nSample from WellFargoWithdrawal.csv:")
print(withdrawals_output.head())

# %%
import pandas as pd
import numpy as np
from datetime import datetime

# Function to aggregate deposits by month for a given bank
def aggregate_monthly_deposits(file_path, bank_name):
    """
    Read deposits CSV and aggregate by month
    """
    # Read the CSV file
    df = pd.read_csv(file_path)
    
    # Display diagnostic info
    print(f"Processing {bank_name}:")
    print(f"Shape: {df.shape}")
    print(f"Column types: {df.dtypes}")
    print("First few rows:")
    print(df.head())
    print()
    
    # Convert date column to datetime
    df['date'] = pd.to_datetime(df['date'], format='%m/%d/%Y', errors='coerce')
    
    # Remove any rows with invalid dates
    initial_count = len(df)
    df = df.dropna(subset=['date'])
    final_count = len(df)
    
    if initial_count != final_count:
        print(f"Removed {initial_count - final_count} rows with invalid dates from {bank_name}")
    
    print(f"Final data shape for {bank_name}: {df.shape}")
    print()
    
    # Create year-month column for grouping
    df['year_month'] = df['date'].dt.to_period('M')
    
    # Aggregate by month
    monthly_agg = df.groupby('year_month')['amount'].sum().reset_index()
    
    # Add bank name column
    monthly_agg['bank'] = bank_name
    
    # Rename columns for clarity
    monthly_agg = monthly_agg.rename(columns={'year_month': 'month', 'amount': 'total_deposits'})
    
    return monthly_agg

# 1. Aggregate Chase Deposits by month
print("=== CHASE DEPOSITS BY MONTH ===")
chase_monthly = aggregate_monthly_deposits('../data/cvs/ChaseDeposits.csv', 'Chase')
chase_monthly_display = chase_monthly[['month', 'total_deposits']].copy()
chase_monthly_display['total_deposits'] = chase_monthly_display['total_deposits'].apply(lambda x: f"${x:,.2f}")
print(chase_monthly_display.to_string(index=False))
print(f"\nChase Total: ${chase_monthly['total_deposits'].sum():,.2f}")

print("\n" + "="*50 + "\n")

# 2. Aggregate Popular Deposits by month  
print("=== POPULAR DEPOSITS BY MONTH ===")
popular_monthly = aggregate_monthly_deposits('../data/cvs/PopularDeposits.csv', 'Popular')
popular_monthly_display = popular_monthly[['month', 'total_deposits']].copy()
popular_monthly_display['total_deposits'] = popular_monthly_display['total_deposits'].apply(lambda x: f"${x:,.2f}")
print(popular_monthly_display.to_string(index=False))
print(f"\nPopular Total: ${popular_monthly['total_deposits'].sum():,.2f}")

print("\n" + "="*50 + "\n")

# 3. Aggregate Wells Fargo Deposits by month
print("=== WELLS FARGO DEPOSITS BY MONTH ===")
wellsfargo_monthly = aggregate_monthly_deposits('../data/cvs/WellFargoDeposits.csv', 'Wells Fargo')
wellsfargo_monthly_display = wellsfargo_monthly[['month', 'total_deposits']].copy()
wellsfargo_monthly_display['total_deposits'] = wellsfargo_monthly_display['total_deposits'].apply(lambda x: f"${x:,.2f}")
print(wellsfargo_monthly_display.to_string(index=False))
print(f"\nWells Fargo Total: ${wellsfargo_monthly['total_deposits'].sum():,.2f}")

print("\n" + "="*70 + "\n")

# 4. Combine all three tables into "TechCargo Cash-inflow"
print("=== TECHCARGO CASH-INFLOW (COMBINED) ===")

# Combine all monthly data
all_monthly = pd.concat([chase_monthly, popular_monthly, wellsfargo_monthly], ignore_index=True)

# Create pivot table with banks as columns and months as rows
cash_inflow = all_monthly.pivot_table(
    values='total_deposits', 
    index='month', 
    columns='bank', 
    fill_value=0,
    aggfunc='sum'
).reset_index()

# Calculate total cash inflow per month
cash_inflow['Monthly Total'] = cash_inflow[['Chase', 'Popular', 'Wells Fargo']].sum(axis=1)

# Format the display version
cash_inflow_display = cash_inflow.copy()
for col in ['Chase', 'Popular', 'Wells Fargo', 'Monthly Total']:
    cash_inflow_display[col] = cash_inflow_display[col].apply(lambda x: f"${x:,.2f}")

print(cash_inflow_display.to_string(index=False))

# Calculate and display totals by bank
print(f"\n=== SUMMARY TOTALS ===")
print(f"Chase Total:        ${cash_inflow['Chase'].sum():,.2f}")
print(f"Popular Total:      ${cash_inflow['Popular'].sum():,.2f}")
print(f"Wells Fargo Total:  ${cash_inflow['Wells Fargo'].sum():,.2f}")
print(f"GRAND TOTAL:        ${cash_inflow['Monthly Total'].sum():,.2f}")

# Save the combined cash inflow table
cash_inflow.to_csv('../data/cvs/TechCargo_Cash_Inflow.csv', index=False)
print(f"\nSaved combined data to: ../data/cvs/TechCargo_Cash_Inflow.csv")

# %%
import pandas as pd
import numpy as np
from datetime import datetime

# Read the source CSV file
df = pd.read_csv('../data/TecnoCargoInvoiceDataset01.csv')

print("Original data info:")
print(f"Shape: {df.shape}")
print(f"Columns: {list(df.columns)}")
print("\nFirst few rows:")
print(df.head())
print()

# Filter for Type = "Invoice" only
invoice_df = df[df['Type'] == 'Invoice'].copy()
print(f"Filtered to {len(invoice_df)} invoice records")

# Select and rename the required columns
result_df = pd.DataFrame({
    'Type': 'Invoice',  # Set all to "Invoice"
    'Number': invoice_df['Number'],
    'Transaction Date': invoice_df['Transaction Date'],
    'Last Payment Date': invoice_df['Last Payment Date'],
    'Amt. Paid (USD)': invoice_df['Amt. Paid (USD)']  # Correct column name without **
})

print("Selected columns, now cleaning data...")

# Clean and convert Transaction Date to proper MM/DD/YYYY format
result_df['Transaction Date'] = pd.to_datetime(result_df['Transaction Date'], errors='coerce')
result_df['Transaction Date'] = result_df['Transaction Date'].dt.strftime('%m/%d/%Y')

# Clean and convert Last Payment Date to proper MM/DD/YYYY format
# Handle empty values first
result_df['Last Payment Date'] = result_df['Last Payment Date'].replace('', np.nan)
result_df['Last Payment Date'] = pd.to_datetime(result_df['Last Payment Date'], errors='coerce')
result_df['Last Payment Date'] = result_df['Last Payment Date'].dt.strftime('%m/%d/%Y')

# Debug: Show original amount values before cleaning
print("Sample of original Amt. Paid values:")
print(result_df['Amt. Paid (USD)'].head(10).tolist())
print(f"Original amount column type: {result_df['Amt. Paid (USD)'].dtype}")

# Clean and convert Amt. Paid (USD) to float with 2 decimal places
result_df['Amt. Paid (USD)'] = result_df['Amt. Paid (USD)'].astype(str)

# Debug: Show values after string conversion
print("\nAfter converting to string:")
print(result_df['Amt. Paid (USD)'].head(10).tolist())

# Remove quotes (if present in data)
result_df['Amt. Paid (USD)'] = result_df['Amt. Paid (USD)'].str.replace('"', '', regex=False)

# Debug: Show values after removing quotes
print("\nAfter removing quotes:")
print(result_df['Amt. Paid (USD)'].head(10).tolist())

# Remove commas (if present in data)
result_df['Amt. Paid (USD)'] = result_df['Amt. Paid (USD)'].str.replace(',', '', regex=False)

# Debug: Show values after removing commas
print("\nAfter removing commas:")
print(result_df['Amt. Paid (USD)'].head(10).tolist())

# Handle 'nan' strings and convert to numeric
result_df['Amt. Paid (USD)'] = result_df['Amt. Paid (USD)'].replace('nan', '0')
result_df['Amt. Paid (USD)'] = pd.to_numeric(result_df['Amt. Paid (USD)'], errors='coerce')

# Debug: Show values after numeric conversion
print(f"\nAfter numeric conversion:")
print(result_df['Amt. Paid (USD)'].head(10).tolist())
print(f"Count of NaN values: {result_df['Amt. Paid (USD)'].isna().sum()}")

# Fill any remaining NaN with 0 and round to 2 decimal places
result_df['Amt. Paid (USD)'] = result_df['Amt. Paid (USD)'].fillna(0).round(2)

# Remove any rows where data conversion failed
initial_count = len(result_df)
result_df = result_df.dropna(subset=['Transaction Date', 'Amt. Paid (USD)'])
final_count = len(result_df)

if initial_count != final_count:
    print(f"Removed {initial_count - final_count} rows with invalid data")

print(f"Final data shape: {result_df.shape}")
print(f"Column types:")
print(result_df.dtypes)

# Save the result to CSV with proper formatting
result_df.to_csv('../data/InvoiceTableForAggregation.csv', index=False, float_format='%.2f')

print(f"\nCreated InvoiceTableForAggregation.csv with {len(result_df)} records")

# Display sample of the output for verification
print("\nSample from InvoiceTableForAggregation.csv:")
print(result_df.head(10))

# Display summary statistics with detailed debugging
print(f"\nDETAILED SUMMARY STATISTICS:")
print(f"Total records processed: {len(result_df)}")
print(f"Records with non-zero amounts: {len(result_df[result_df['Amt. Paid (USD)'] > 0])}")
print(f"Records with zero amounts: {len(result_df[result_df['Amt. Paid (USD)'] == 0])}")
print(f"Records with NaN amounts: {result_df['Amt. Paid (USD)'].isna().sum()}")
print(f"Sum of all amounts: ${result_df['Amt. Paid (USD)'].sum():,.2f}")
print(f"Sum of non-zero amounts: ${result_df[result_df['Amt. Paid (USD)'] > 0]['Amt. Paid (USD)'].sum():,.2f}")
print(f"Max amount: ${result_df['Amt. Paid (USD)'].max():,.2f}")
print(f"Min amount: ${result_df['Amt. Paid (USD)'].min():,.2f}")
print(f"Average amount: ${result_df['Amt. Paid (USD)'].mean():,.2f}")

# Show the largest payments to verify data quality
print(f"\nTop 10 largest payments:")
top_payments = result_df.nlargest(10, 'Amt. Paid (USD)')
print(top_payments[['Number', 'Transaction Date', 'Last Payment Date', 'Amt. Paid (USD)']])

# %%


# %%
import pandas as pd
import numpy as np
from datetime import datetime

# Read the invoice aggregation table
df = pd.read_csv('../data/InvoiceTableForAggregation.csv')

print("=== INVOICE PAYMENT DATA ===")
print(f"Total records: {len(df)}")
print(f"Column types: {df.dtypes}")
print("\nFirst few rows:")
print(df.head())
print()

# Filter out records without Last Payment Date (unpaid invoices)
# Only aggregate invoices that have been paid
paid_invoices = df[df['Last Payment Date'].notna() & (df['Last Payment Date'] != '')].copy()

print(f"Records with Last Payment Date: {len(paid_invoices)}")
print(f"Records without Last Payment Date (unpaid): {len(df) - len(paid_invoices)}")
print()

if len(paid_invoices) == 0:
    print("No paid invoices found - cannot create monthly aggregation")
else:
    # Convert Last Payment Date to datetime
    paid_invoices['Last Payment Date'] = pd.to_datetime(paid_invoices['Last Payment Date'], format='%m/%d/%Y', errors='coerce')
    
    # Remove any rows where date conversion failed
    initial_count = len(paid_invoices)
    paid_invoices = paid_invoices.dropna(subset=['Last Payment Date'])
    final_count = len(paid_invoices)
    
    if initial_count != final_count:
        print(f"Removed {initial_count - final_count} records with invalid payment dates")
    
    # Create year-month column for grouping
    paid_invoices['year_month'] = paid_invoices['Last Payment Date'].dt.to_period('M')
    
    # Aggregate by month
    monthly_payments = paid_invoices.groupby('year_month').agg({
        'Amt. Paid (USD)': 'sum',
        'Number': 'count'  # Count of invoices paid in each month
    }).reset_index()
    
    # Rename columns for clarity
    monthly_payments = monthly_payments.rename(columns={
        'year_month': 'Payment Month',
        'Amt. Paid (USD)': 'Total Payments (USD)',
        'Number': 'Invoices Paid'
    })
    
    # Sort by month
    monthly_payments = monthly_payments.sort_values('Payment Month')
    
    print("=== MONTHLY PAYMENT AGGREGATION ===")
    
    # Create display version with formatted amounts
    monthly_display = monthly_payments.copy()
    monthly_display['Total Payments (USD)'] = monthly_display['Total Payments (USD)'].apply(lambda x: f"${x:,.2f}")
    
    print(monthly_display.to_string(index=False))
    
    # Calculate and display summary statistics
    print(f"\n=== SUMMARY STATISTICS ===")
    print(f"Total Months with Payments: {len(monthly_payments)}")
    print(f"Total Amount Paid: ${monthly_payments['Total Payments (USD)'].sum():,.2f}")
    print(f"Average Monthly Payment: ${monthly_payments['Total Payments (USD)'].mean():,.2f}")
    print(f"Highest Monthly Payment: ${monthly_payments['Total Payments (USD)'].max():,.2f}")
    print(f"Lowest Monthly Payment: ${monthly_payments['Total Payments (USD)'].min():,.2f}")
    print(f"Total Invoices Paid: {monthly_payments['Invoices Paid'].sum():,}")
    print(f"Average Invoices Paid per Month: {monthly_payments['Invoices Paid'].mean():.1f}")
    
    # Save the monthly aggregation
    monthly_payments.to_csv('../data/MonthlyPaymentAggregation.csv', index=False, float_format='%.2f')
    print(f"\nSaved monthly aggregation to: ../data/MonthlyPaymentAggregation.csv")
    
    # Show payment trends
    if len(monthly_payments) > 1:
        print(f"\n=== PAYMENT TRENDS ===")
        first_month = monthly_payments.iloc[0]
        last_month = monthly_payments.iloc[-1]
        print(f"First payment month: {first_month['Payment Month']} - ${first_month['Total Payments (USD)']:,.2f}")
        print(f"Last payment month: {last_month['Payment Month']} - ${last_month['Total Payments (USD)']:,.2f}")
        
        # Calculate month-over-month change if we have recent data
        if len(monthly_payments) >= 2:
            recent_change = last_month['Total Payments (USD)'] - monthly_payments.iloc[-2]['Total Payments (USD)']
            print(f"Month-over-month change: ${recent_change:,.2f}")

print(f"\n=== UNPAID INVOICES SUMMARY ===")
unpaid_invoices = df[df['Last Payment Date'].isna() | (df['Last Payment Date'] == '')].copy()
if len(unpaid_invoices) > 0:
    # Note: Unpaid invoices will have Amt. Paid = 0, but let's show the outstanding amounts would be in Amt. Due
    print(f"Number of unpaid invoices: {len(unpaid_invoices)}")
    print(f"These invoices are not included in the monthly payment aggregation")
else:
    print("All invoices have payment dates")

# %%
import pandas as pd
import numpy as np

# Read the bank cash inflow data
print("=== READING BANK CASH INFLOW DATA ===")
cash_inflow_df = pd.read_csv('../data/cvs/TechCargo_Cash_Inflow.csv')
print(f"Cash Inflow Shape: {cash_inflow_df.shape}")
print(f"Cash Inflow Columns: {list(cash_inflow_df.columns)}")
print("Sample Cash Inflow Data:")
print(cash_inflow_df.head())
print()

# Read the invoice payments data
print("=== READING INVOICE PAYMENTS DATA ===")
payments_df = pd.read_csv('../data/MonthlyPaymentAggregation.csv')
print(f"Payments Shape: {payments_df.shape}")
print(f"Payments Columns: {list(payments_df.columns)}")
print("Sample Payments Data:")
print(payments_df.head())
print()

# Prepare cash inflow data
# Select month and Monthly Total columns
cash_inflow_clean = cash_inflow_df[['month', 'Monthly Total']].copy()
cash_inflow_clean = cash_inflow_clean.rename(columns={
    'month': 'Month',
    'Monthly Total': 'Bank Cash-inflow'
})

# Convert month to period if it's not already
if cash_inflow_clean['Month'].dtype == 'object':
    cash_inflow_clean['Month'] = pd.to_datetime(cash_inflow_clean['Month']).dt.to_period('M')

print("Cleaned Cash Inflow Data:")
print(cash_inflow_clean)
print()

# Prepare payments data
# Select required columns and rename
payments_clean = payments_df[['Payment Month', 'Total Payments (USD)', 'Invoices Paid']].copy()
payments_clean = payments_clean.rename(columns={
    'Payment Month': 'Month',
    'Total Payments (USD)': 'Payments from Invoices'
})

# Convert month to period if it's not already
if payments_clean['Month'].dtype == 'object':
    payments_clean['Month'] = pd.to_datetime(payments_clean['Month']).dt.to_period('M')

print("Cleaned Payments Data:")
print(payments_clean)
print()

# Merge the two datasets on Month
# Use outer join to include all months from both datasets
combined_df = pd.merge(
    cash_inflow_clean, 
    payments_clean, 
    on='Month', 
    how='outer'
)

# Fill NaN values with 0 for missing data
combined_df['Bank Cash-inflow'] = combined_df['Bank Cash-inflow'].fillna(0)
combined_df['Payments from Invoices'] = combined_df['Payments from Invoices'].fillna(0)
combined_df['Invoices Paid'] = combined_df['Invoices Paid'].fillna(0)

# Sort by month
combined_df = combined_df.sort_values('Month')

# Reset index
combined_df = combined_df.reset_index(drop=True)

print("=== COMBINED MONTHLY DATA ===")
print(f"Combined data shape: {combined_df.shape}")
print()

# Create display version with formatted amounts
combined_display = combined_df.copy()
combined_display['Bank Cash-inflow'] = combined_display['Bank Cash-inflow'].apply(lambda x: f"${x:,.2f}")
combined_display['Payments from Invoices'] = combined_display['Payments from Invoices'].apply(lambda x: f"${x:,.2f}")
combined_display['Invoices Paid'] = combined_display['Invoices Paid'].apply(lambda x: f"{x:,.0f}")

print("MONTHLY CASH FLOW ANALYSIS:")
print(combined_display.to_string(index=False))

# Calculate summary statistics
print(f"\n=== SUMMARY STATISTICS ===")
print(f"Total Bank Cash Inflow: ${combined_df['Bank Cash-inflow'].sum():,.2f}")
print(f"Total Invoice Payments: ${combined_df['Payments from Invoices'].sum():,.2f}")
print(f"Total Invoices Paid: {combined_df['Invoices Paid'].sum():,.0f}")
print(f"Number of months with bank data: {len(combined_df[combined_df['Bank Cash-inflow'] > 0])}")
print(f"Number of months with payment data: {len(combined_df[combined_df['Payments from Invoices'] > 0])}")

# Calculate differences and ratios
combined_df['Difference (Bank - Payments)'] = combined_df['Bank Cash-inflow'] - combined_df['Payments from Invoices']
combined_df['Payment Ratio (%)'] = (combined_df['Payments from Invoices'] / combined_df['Bank Cash-inflow'] * 100).replace([np.inf, -np.inf], 0).fillna(0)

print(f"\n=== CASH FLOW ANALYSIS ===")
print(f"Average monthly bank inflow: ${combined_df['Bank Cash-inflow'].mean():,.2f}")
print(f"Average monthly invoice payments: ${combined_df['Payments from Invoices'].mean():,.2f}")

# Show analysis with differences
analysis_display = combined_df.copy()
analysis_display['Bank Cash-inflow'] = analysis_display['Bank Cash-inflow'].apply(lambda x: f"${x:,.2f}")
analysis_display['Payments from Invoices'] = analysis_display['Payments from Invoices'].apply(lambda x: f"${x:,.2f}")
analysis_display['Difference (Bank - Payments)'] = analysis_display['Difference (Bank - Payments)'].apply(lambda x: f"${x:,.2f}")
analysis_display['Payment Ratio (%)'] = analysis_display['Payment Ratio (%)'].apply(lambda x: f"{x:.1f}%")
analysis_display['Invoices Paid'] = analysis_display['Invoices Paid'].apply(lambda x: f"{x:,.0f}")

print(f"\nDETAILED MONTHLY ANALYSIS:")
print(analysis_display.to_string(index=False))

# Save the combined data
combined_df_save = combined_df[['Month', 'Bank Cash-inflow', 'Payments from Invoices', 'Invoices Paid']].copy()
combined_df_save.to_csv('../data/MonthlyCashFlowAnalysis.csv', index=False, float_format='%.2f')
print(f"\nSaved combined analysis to: ../data/MonthlyCashFlowAnalysis.csv")

# Key insights
print(f"\n=== KEY INSIGHTS ===")
total_diff = combined_df['Difference (Bank - Payments)'].sum()
if total_diff > 0:
    print(f"Bank deposits exceed invoice payments by ${total_diff:,.2f}")
    print("This suggests additional revenue sources beyond invoiced services")
elif total_diff < 0:
    print(f"Invoice payments exceed bank deposits by ${abs(total_diff):,.2f}")
    print("This may indicate timing differences or collection issues")
else:
    print("Bank deposits and invoice payments are perfectly matched")

avg_ratio = combined_df[combined_df['Bank Cash-inflow'] > 0]['Payment Ratio (%)'].mean()
print(f"Average invoice payments as % of bank deposits: {avg_ratio:.1f}%")

# %%
import pandas as pd
import numpy as np

# Read the existing monthly cash flow analysis file
print("=== READING MONTHLY CASH FLOW DATA ===")
df = pd.read_csv('../data/MonthlyCashFlowAnalysis.csv')

print(f"Original data shape: {df.shape}")
print(f"Columns: {list(df.columns)}")
print("Sample data:")
print(df.head())
print()

# Convert Month column to period for easier handling
if df['Month'].dtype == 'object':
    df['Month'] = pd.to_datetime(df['Month']).dt.to_period('M')

print("Data after month conversion:")
print(df.head())
print()

# Create complete date range from 2024-01 to 2025-05
start_date = '2024-01'
end_date = '2025-05'
complete_months = pd.period_range(start=start_date, end=end_date, freq='M')

print(f"Creating complete month range from {start_date} to {end_date}:")
print(f"Total months: {len(complete_months)}")
print()

# Create template with all months
complete_month_df = pd.DataFrame({'Month': complete_months})

# Merge with existing data to fill in available months
result_df = pd.merge(complete_month_df, df, on='Month', how='left')

# Fill missing values with 0
result_df['Bank Cash-inflow'] = result_df['Bank Cash-inflow'].fillna(0)
result_df['Payments from Invoices'] = result_df['Payments from Invoices'].fillna(0)
result_df['Invoices Paid'] = result_df['Invoices Paid'].fillna(0)

# Calculate totals for summary row
total_bank_inflow = result_df['Bank Cash-inflow'].sum()
total_payments = result_df['Payments from Invoices'].sum()
total_invoices = result_df['Invoices Paid'].sum()

print(f"=== CALCULATED TOTALS ===")
print(f"Total Bank Cash-inflow: ${total_bank_inflow:,.2f}")
print(f"Total Payments from Invoices: ${total_payments:,.2f}")
print(f"Total Invoices Paid: {total_invoices:,.0f}")
print()

# Create totals row
totals_row = pd.DataFrame({
    'Month': ['TOTAL'],
    'Bank Cash-inflow': [total_bank_inflow],
    'Payments from Invoices': [total_payments],
    'Invoices Paid': [total_invoices]
})

# Add totals row to the main data
final_table = pd.concat([result_df, totals_row], ignore_index=True)

print("=== MONTHLY CASH FLOW TABLE (2024-01 to 2025-05) ===")

# Create display version with formatted amounts
display_table = final_table.copy()

# Format the amounts for display
display_table['Bank Cash-inflow'] = display_table['Bank Cash-inflow'].apply(lambda x: f"${x:,.2f}")
display_table['Payments from Invoices'] = display_table['Payments from Invoices'].apply(lambda x: f"${x:,.2f}")
display_table['Invoices Paid'] = display_table['Invoices Paid'].apply(lambda x: f"{x:,.0f}")

# Display the complete table
print(display_table.to_string(index=False))

print(f"\n=== TABLE SUMMARY ===")
print(f"Period covered: {start_date} to {end_date}")
print(f"Total months: {len(complete_months)}")
print(f"Months with bank data: {len(result_df[result_df['Bank Cash-inflow'] > 0])}")
print(f"Months with payment data: {len(result_df[result_df['Payments from Invoices'] > 0])}")

# Save the final table with totals
final_table_save = final_table.copy()
final_table_save.to_csv('../data/MonthlyTableWithTotals_2024_2025.csv', index=False, float_format='%.2f')
print(f"\nSaved complete table to: ../data/MonthlyTableWithTotals_2024_2025.csv")

# Additional analysis
print(f"\n=== MONTHLY AVERAGES ===")
# Calculate averages excluding the totals row
monthly_data = result_df.copy()
avg_bank_inflow = monthly_data['Bank Cash-inflow'].mean()
avg_payments = monthly_data['Payments from Invoices'].mean()
avg_invoices = monthly_data['Invoices Paid'].mean()

print(f"Average monthly bank inflow: ${avg_bank_inflow:,.2f}")
print(f"Average monthly payments: ${avg_payments:,.2f}")
print(f"Average invoices paid per month: {avg_invoices:.1f}")

# Show months with highest activity
print(f"\n=== TOP MONTHS ===")
top_bank_month = monthly_data.loc[monthly_data['Bank Cash-inflow'].idxmax()]
top_payment_month = monthly_data.loc[monthly_data['Payments from Invoices'].idxmax()]

print(f"Highest bank inflow: {top_bank_month['Month']} - ${top_bank_month['Bank Cash-inflow']:,.2f}")
print(f"Highest payments: {top_payment_month['Month']} - ${top_payment_month['Payments from Invoices']:,.2f}")


