import pandas as pd
import re
from datetime import datetime

def extract_account_number_and_name(sheet_name):
    """Extracts the account number and name from the sheet name."""
    match = re.match(r'(\d+)\s+(.+)', sheet_name)
    if match:
        return match.group(1), match.group(2)
    return None, None

def classify_account(account_number):
    """Classifies an account based on its starting digit."""
    if not account_number:
        return None
    first_digit = account_number[0]
    if first_digit == '1':
        return 'Asset'
    elif first_digit == '2':
        return 'Liability'
    elif first_digit == '3':
        return 'Revenue'
    elif first_digit in '45678':
        return 'Expense'
    return None

def process_account_data(df, account_number):
    """Processes account data to compute net amounts per year."""
    # Convert Date to datetime
    df['Date'] = pd.to_datetime(df['Date'], format='%d.%m.%Y', errors='coerce')
    # Extract year
    df['Year'] = df['Date'].dt.year
    # Convert Débit and Crédit to numeric, handling empty values
    df['Débit'] = pd.to_numeric(df['Débit'], errors='coerce').fillna(0)
    df['Crédit'] = pd.to_numeric(df['Crédit'], errors='coerce').fillna(0)
    
    # Compute net amount (debits - credits) per year
    yearly_net = df.groupby('Year').agg({
        'Débit': 'sum',
        'Crédit': 'sum'
    }).reset_index()
    yearly_net['Net'] = yearly_net['Débit'] - yearly_net['Crédit']
    
    return yearly_net[['Year', 'Net']]

def generate_financial_statements(input_file, output_file):
    """Generates Balance Sheet and Income Statement from Comptes_Cleans.xlsx."""
    # Read the input Excel file
    xl = pd.ExcelFile(input_file)
    
    # Initialize data structures
    balance_sheet_data = {'Asset': {}, 'Liability': {}}
    income_statement_data = {'Revenue': {}, 'Expense': {}}
    all_years = set()
    
    # Process each sheet (account)
    for sheet_name in xl.sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        account_number, account_name = extract_account_number_and_name(sheet_name)
        if not account_number:
            continue
        
        account_type = classify_account(account_number)
        if not account_type:
            continue
        
        # Get net amounts per year
        yearly_net = process_account_data(df, account_number)
        
        # Store data based on account type
        if account_type in ['Asset', 'Liability']:
            balance_sheet_data[account_type][(account_number, account_name)] = yearly_net
        elif account_type in ['Revenue', 'Expense']:
            income_statement_data[account_type][(account_number, account_name)] = yearly_net
        
        # Collect all years
        all_years.update(yearly_net['Year'].dropna().astype(int))
    
    # Sort years
    all_years = sorted(list(all_years))
    
    # Prepare Balance Sheet
    balance_sheet_rows = []
    for account_type in ['Asset', 'Liability']:
        for (account_number, account_name), yearly_net in balance_sheet_data[account_type].items():
            row = {'Account Number': account_number, 'Account Name': account_name}
            # Initialize cumulative balance
            cumulative = 0
            for year in all_years:
                net = yearly_net[yearly_net['Year'] == year]['Net']
                net_value = net.iloc[0] if not net.empty else 0
                cumulative += net_value
                row[str(year)] = cumulative
            balance_sheet_rows.append(row)
    
    # Prepare Income Statement
    income_statement_rows = []
    for account_type in ['Revenue', 'Expense']:
        for (account_number, account_name), yearly_net in income_statement_data[account_type].items():
            row = {'Account Number': account_number, 'Account Name': account_name}
            for year in all_years:
                net = yearly_net[yearly_net['Year'] == year]['Net']
                row[str(year)] = net.iloc[0] if not net.empty else 0
            income_statement_rows.append(row)
    
    # Create DataFrames
    balance_sheet_df = pd.DataFrame(balance_sheet_rows)
    income_statement_df = pd.DataFrame(income_statement_rows)
    
    # Adjustment: Check Balance Sheet column sums and add to Balance Sheet
    if not balance_sheet_df.empty and not income_statement_df.empty:
        for year in all_years:
            year_str = str(year)
            if year_str in balance_sheet_df.columns:
                # Compute sum of Balance Sheet column
                balance_sum = pd.to_numeric(balance_sheet_df[year_str], errors='coerce').sum()
                # Check if sum is non-zero (allow small floating-point differences)
                if abs(balance_sum) > 0.01:
                    # Compute sum of Income Statement column (net profit/loss)
                    income_sum = pd.to_numeric(income_statement_df[year_str], errors='coerce').sum()
                    # Add or update account 2979 in Balance Sheet
                    new_row = {'Account Number': '2979', 'Account Name': 'Résultat de l’exercice'}
                    for y in all_years:
                        new_row[str(y)] = income_sum if y == year else 0
                    # Check if 2979 already exists
                    if '2979' in balance_sheet_df['Account Number'].values:
                        balance_sheet_df.loc[balance_sheet_df['Account Number'] == '2979', year_str] = income_sum
                    else:
                        balance_sheet_df = pd.concat([balance_sheet_df, pd.DataFrame([new_row])], ignore_index=True)
    
    # Write to Excel
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # Write Balance Sheet
        if not balance_sheet_df.empty:
            balance_sheet_df.to_excel(writer, sheet_name='Balance Sheet', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Balance Sheet']
            # Format numbers
            number_format = workbook.add_format({'num_format': '#,##0.00'})
            for col in range(2, len(all_years) + 2):  # Start from column 2 (after Account Number, Account Name)
                worksheet.set_column(col, col, None, number_format)
        
        # Write Income Statement
        if not income_statement_df.empty:
            income_statement_df.to_excel(writer, sheet_name='Income Statement', index=False)
            worksheet = writer.sheets['Income Statement']
            for col in range(2, len(all_years) + 2):
                worksheet.set_column(col, col, None, number_format)

if __name__ == "__main__":
    input_file = "Comptes_Cleans.xlsx"
    output_file = "Financial_Statements.xlsx"
    generate_financial_statements(input_file, output_file)