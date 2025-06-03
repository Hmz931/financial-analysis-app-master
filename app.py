
from flask import Flask, render_template, request, send_file, flash, redirect, url_for, jsonify
import pandas as pd
import json
import os
from werkzeug.utils import secure_filename
from datetime import datetime
import traceback

# Import our processing modules
try:
    from GL_Cleaner import main as clean_gl_data
    from generate_financial_statements import generate_financial_statements
except ImportError as e:
    print(f"Import error: {e}")
    print("Make sure GL_Cleaner.py and generate_financial_statements.py are in the same directory")

app = Flask(__name__)
app.secret_key = os.urandom(24)

# Configuration
UPLOAD_FOLDER = 'uploads'
RESULTS_FOLDER = 'results'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Create directories
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULTS_FOLDER, exist_ok=True)

def allowed_file(filename):
    """Check if file extension is allowed."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def calculate_financial_ratios(balance_sheet_df, income_statement_df):
    """Calculate comprehensive financial ratios from the financial statements."""
    ratios_data = {}
    
    # Get all year columns (exclude account info columns)
    year_columns = [col for col in balance_sheet_df.columns 
                   if col not in ['Account Number', 'Account Name']]
    
    for year in year_columns:
        try:
            # Convert to numeric, handling any string values
            bs_year = pd.to_numeric(balance_sheet_df[year], errors='coerce').fillna(0)
            is_year = pd.to_numeric(income_statement_df[year], errors='coerce').fillna(0)
            
            # Get account numbers for easier filtering
            bs_accounts = balance_sheet_df['Account Number'].astype(str)
            is_accounts = income_statement_df['Account Number'].astype(str)
            
            # Assets (accounts starting with 1)
            current_assets = bs_year[bs_accounts.str.startswith(('10', '11', '12', '13'))].sum()
            cash_equivalents = bs_year[bs_accounts.str.startswith('10')].sum()
            inventory = bs_year[bs_accounts.str.startswith('12')].sum()
            total_assets = bs_year[bs_accounts.str.startswith('1')].sum()
            fixed_assets = bs_year[bs_accounts.str.startswith(('15', '16', '17', '18'))].sum()
            
            # Liabilities (accounts starting with 2, excluding equity)
            current_liabilities = bs_year[bs_accounts.str.startswith(('20', '21', '22', '23'))].sum()
            long_term_debt = bs_year[bs_accounts.str.startswith(('24', '25', '27'))].sum()
            total_debt = current_liabilities + long_term_debt
            
            # Equity (accounts 28, 29)
            equity = bs_year[bs_accounts.str.startswith(('28', '29'))].sum()
            
            # Working capital
            working_capital = current_assets - current_liabilities
            
            # Income Statement items - handle Swiss accounting where revenues can be negative
            revenues = abs(is_year[is_accounts.str.startswith('3')].sum())
            cost_of_goods = abs(is_year[is_accounts.str.startswith('4')].sum())
            personnel_costs = abs(is_year[is_accounts.str.startswith('5')].sum())
            other_expenses = abs(is_year[is_accounts.str.startswith(('6', '7'))].sum())
            financial_expenses = abs(is_year[is_accounts.str.startswith('8')].sum())
            
            total_expenses = cost_of_goods + personnel_costs + other_expenses + financial_expenses
            net_income = revenues - total_expenses
            
            # EBITDA approximation (before depreciation and financial costs)
            ebitda = revenues - cost_of_goods - personnel_costs - other_expenses
            
            # Calculate ratios
            ratios = {}
            
            # Liquidity Ratios (Swiss standard)
            ratios['current_ratio'] = current_assets / current_liabilities if current_liabilities != 0 else 0
            ratios['quick_ratio'] = (current_assets - inventory) / current_liabilities if current_liabilities != 0 else 0
            ratios['cash_ratio'] = cash_equivalents / current_liabilities if current_liabilities != 0 else 0
            ratios['working_capital'] = working_capital
            
            # Profitability Ratios
            ratios['net_margin'] = (net_income / revenues * 100) if revenues != 0 else 0
            ratios['roa'] = (net_income / total_assets * 100) if total_assets != 0 else 0
            ratios['roe'] = (net_income / equity * 100) if equity != 0 else 0
            ratios['ebitda_margin'] = (ebitda / revenues * 100) if revenues != 0 else 0
            
            # Solvency Ratios (Swiss standard)
            ratios['equity_ratio'] = equity / total_assets if total_assets != 0 else 0
            ratios['debt_to_equity'] = total_debt / equity if equity != 0 else 0
            ratios['debt_to_assets'] = total_debt / total_assets if total_assets != 0 else 0
            ratios['interest_coverage'] = ebitda / financial_expenses if financial_expenses != 0 else 0
            
            # Efficiency Ratios
            ratios['asset_turnover'] = revenues / total_assets if total_assets != 0 else 0
            ratios['fixed_asset_turnover'] = revenues / fixed_assets if fixed_assets != 0 else 0
            
            ratios_data[year] = ratios
            
        except Exception as e:
            print(f"Error calculating ratios for year {year}: {e}")
            continue
    
    return ratios_data

def prepare_chart_data(balance_sheet_df, income_statement_df):
    """Prepare data for charts."""
    chart_data = {
        'assets_breakdown': {},
        'liabilities_breakdown': {},
        'revenue_breakdown': {},
        'expense_breakdown': {}
    }
    
    year_columns = [col for col in balance_sheet_df.columns 
                   if col not in ['Account Number', 'Account Name']]
    
    for year in year_columns:
        try:
            # Assets breakdown
            bs_year = pd.to_numeric(balance_sheet_df[year], errors='coerce').fillna(0)
            bs_accounts = balance_sheet_df['Account Number'].astype(str)
            
            assets_data = {
                'Liquidités & équivalents': bs_year[bs_accounts.str.startswith('10')].sum(),
                'Créances': bs_year[bs_accounts.str.startswith('11')].sum(),
                'Stocks': bs_year[bs_accounts.str.startswith('12')].sum(),
                'Immobilisations': bs_year[bs_accounts.str.startswith(('15', '16', '17', '18'))].sum(),
                'Autres actifs': bs_year[bs_accounts.str.startswith(('13', '14'))].sum()
            }
            # Filter out zero values
            assets_data = {k: v for k, v in assets_data.items() if v > 0}
            chart_data['assets_breakdown'][year] = assets_data
            
            # Liabilities breakdown
            liabilities_data = {
                'Dettes à court terme': bs_year[bs_accounts.str.startswith(('20', '21', '22', '23'))].sum(),
                'Dettes à long terme': bs_year[bs_accounts.str.startswith(('24', '25', '27'))].sum(),
                'Capitaux propres': bs_year[bs_accounts.str.startswith(('28', '29'))].sum()
            }
            liabilities_data = {k: v for k, v in liabilities_data.items() if v != 0}
            chart_data['liabilities_breakdown'][year] = liabilities_data
            
            # Income statement breakdown
            is_year = pd.to_numeric(income_statement_df[year], errors='coerce').fillna(0)
            is_accounts = income_statement_df['Account Number'].astype(str)
            
            # Revenue breakdown - handle both positive and negative values
            revenue_data = {}
            for _, row in income_statement_df.iterrows():
                if str(row['Account Number']).startswith('3'):
                    value = pd.to_numeric(row[year], errors='coerce')
                    if pd.notna(value) and value != 0:
                        # Convert negative revenues to positive for display
                        revenue_data[row['Account Name'][:30]] = abs(value)
            chart_data['revenue_breakdown'][year] = revenue_data
            
            # Expense breakdown
            expense_data = {
                'Coûts directs': abs(is_year[is_accounts.str.startswith('4')].sum()),
                'Charges de personnel': abs(is_year[is_accounts.str.startswith('5')].sum()),
                'Charges d\'exploitation': abs(is_year[is_accounts.str.startswith('6')].sum()),
                'Autres charges': abs(is_year[is_accounts.str.startswith(('7', '8'))].sum())
            }
            expense_data = {k: v for k, v in expense_data.items() if v > 0}
            chart_data['expense_breakdown'][year] = expense_data
            
        except Exception as e:
            print(f"Error preparing chart data for year {year}: {e}")
            continue
    
    return chart_data

@app.route('/', methods=['GET', 'POST'])
def index():
    """Main route for file upload and analysis display."""
    if request.method == 'POST':
        # Check if file was uploaded
        if 'file' not in request.files:
            flash('No file selected', 'error')
            return redirect(request.url)
        
        file = request.files['file']
        if file.filename == '':
            flash('No file selected', 'error')
            return redirect(request.url)
        
        if file and allowed_file(file.filename):
            try:
                # Save uploaded file
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                
                # Process the file
                flash('Processing file...', 'info')
                
                # Step 1: Clean the GL data
                clean_gl_data(filepath)
                
                # Step 2: Generate financial statements
                generate_financial_statements('Comptes_Cleans.xlsx', 'Financial_Statements.xlsx')
                
                # Step 3: Load and analyze the financial statements
                balance_sheet_df = pd.read_excel('Financial_Statements.xlsx', sheet_name='Balance Sheet')
                income_statement_df = pd.read_excel('Financial_Statements.xlsx', sheet_name='Income Statement')
                
                # Step 4: Calculate ratios and prepare chart data
                ratios_data = calculate_financial_ratios(balance_sheet_df, income_statement_df)
                chart_data = prepare_chart_data(balance_sheet_df, income_statement_df)
                
                # Get available years for the interface
                years = [col for col in balance_sheet_df.columns 
                        if col not in ['Account Number', 'Account Name']]
                
                flash('Analysis completed successfully!', 'success')
                
                return render_template('index.html', 
                                     ratios=ratios_data,
                                     chart_data=chart_data,
                                     years=years,
                                     has_data=True)
                
            except Exception as e:
                error_msg = f"Error processing file: {str(e)}"
                print(f"Full error: {traceback.format_exc()}")
                flash(error_msg, 'error')
                return redirect(request.url)
        else:
            flash('Invalid file type. Please upload an Excel file (.xlsx or .xls)', 'error')
            return redirect(request.url)
    
    return render_template('index.html', has_data=False)

@app.route('/download/<filename>')
def download_file(filename):
    """Download generated files."""
    try:
        return send_file(filename, as_attachment=True)
    except FileNotFoundError:
        flash(f'File {filename} not found', 'error')
        return redirect(url_for('index'))

@app.route('/api/financial-data')
def api_financial_data():
    """API endpoint to get financial data as JSON."""
    try:
        # Check if files exist
        if not os.path.exists('Financial_Statements.xlsx'):
            return jsonify({
                'error': 'Financial statements not found. Please upload and process a file first.',
                'status': 'error'
            }), 404
        
        balance_sheet_df = pd.read_excel('Financial_Statements.xlsx', sheet_name='Balance Sheet')
        income_statement_df = pd.read_excel('Financial_Statements.xlsx', sheet_name='Income Statement')
        
        # Debug: Print data shapes
        print(f"Balance sheet shape: {balance_sheet_df.shape}")
        print(f"Income statement shape: {income_statement_df.shape}")
        print(f"Balance sheet columns: {balance_sheet_df.columns.tolist()}")
        
        ratios_data = calculate_financial_ratios(balance_sheet_df, income_statement_df)
        chart_data = prepare_chart_data(balance_sheet_df, income_statement_df)
        
        print(f"Ratios data keys: {list(ratios_data.keys())}")
        print(f"Chart data keys: {list(chart_data.keys())}")
        
        return jsonify({
            'ratios': ratios_data,
            'charts': chart_data,
            'status': 'success'
        })
    except Exception as e:
        print(f"API Error: {str(e)}")
        print(f"Full traceback: {traceback.format_exc()}")
        return jsonify({
            'error': str(e),
            'status': 'error'
        }), 500

@app.route('/test-data')
def test_data():
    """Test endpoint to check available data files."""
    result = {
        'files': {},
        'status': 'success'
    }
    
    files_to_check = [
        'Comptes_Cleans.xlsx',
        'Financial_Statements.xlsx', 
        'Plan_Comptable.xlsx',
        'Summary.xlsx'
    ]
    
    for filename in files_to_check:
        if os.path.exists(filename):
            try:
                xl = pd.ExcelFile(filename)
                result['files'][filename] = {
                    'exists': True,
                    'sheets': xl.sheet_names
                }
                if filename == 'Financial_Statements.xlsx':
                    # Get more details for this file
                    if 'Balance Sheet' in xl.sheet_names:
                        bs_df = pd.read_excel(filename, sheet_name='Balance Sheet')
                        result['files'][filename]['balance_sheet_shape'] = bs_df.shape
                        result['files'][filename]['balance_sheet_columns'] = bs_df.columns.tolist()
                        result['files'][filename]['balance_sheet_sample'] = bs_df.head(5).to_dict('records')
                    if 'Income Statement' in xl.sheet_names:
                        is_df = pd.read_excel(filename, sheet_name='Income Statement')
                        result['files'][filename]['income_statement_shape'] = is_df.shape
                        result['files'][filename]['income_statement_columns'] = is_df.columns.tolist()
                        result['files'][filename]['income_statement_sample'] = is_df.head(5).to_dict('records')
                        
                elif filename == 'Comptes_Cleans.xlsx':
                    # Show structure of cleaned accounts
                    result['files'][filename]['sheet_count'] = len(xl.sheet_names)
                    result['files'][filename]['sample_sheets'] = xl.sheet_names[:5]
                    
            except Exception as e:
                result['files'][filename] = {
                    'exists': True,
                    'error': str(e)
                }
        else:
            result['files'][filename] = {'exists': False}
    
    return jsonify(result)

@app.errorhandler(413)
def too_large(e):
    flash('File too large. Maximum size is 16MB.', 'error')
    return redirect(url_for('index'))

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_ENV') == 'development'
    app.run(host='0.0.0.0', port=port, debug=debug)
