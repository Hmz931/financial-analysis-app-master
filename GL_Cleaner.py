import pandas as pd
import re
from datetime import datetime
import uuid

# Mapping des codes d'origine
ORIGIN_MAPPING = {
    'F': 'Comptabilité financière',
    'K': 'Saisie facture d’achat',
    'k': 'Paiement facture d’achat',
    'D': 'Saisie facture de vente',
    'd': 'Paiement facture de vente',
    'Y': 'EBICS (Electronic Banking)',
    'L': 'Salaire (Lohn)',
    '': 'Écriture manuelle ou inconnue'
}

# Mapping de la nature des comptes
NATURE_MAPPING = {
    '1': 'Actif',
    '2': 'Passif',
    '3': 'Produit',
    '4': 'Charge directe',
    '5': 'Charges de personnel',
    '6': 'Autres charges d’exploitation',
    '7': 'Charges/produits annexes',
    '8': 'Charges/produits extraordinaires',
    '9': 'Comptes auxiliaires/clôtures'
}

def parse_sheet_name(sheet_name):
    """Extrait le numéro et le nom du compte à partir du nom de la feuille."""
    match = re.match(r'_(\d+)_(.+)', sheet_name)
    if match:
        account_number = match.group(1)
        account_name = match.group(2).replace('___', ' ').replace('_', ' ')
        return account_number, account_name
    return None, None

def get_period_and_initial_balance(df):
    """Extrait la période et le solde initial."""
    period_row = df[df['A'].str.contains(r'Solde \d{2}\.\d{2}\.\d{4} - \d{2}\.\d{2}\.\d{4}', na=False)]
    start_date = '01.01.2023'
    end_date = '31.12.2023'
    if not period_row.empty:
        period_text = period_row.iloc[0]['A']
        match = re.search(r'(\d{2}\.\d{2}\.\d{4}) - (\d{2}\.\d{2}\.\d{4})', period_text)
        if match:
            start_date = match.group(1)
            end_date = match.group(2)

    initial_balance_row = df[df['A'] == 'Report de solde']
    initial_balance = 0.0
    if not initial_balance_row.empty:
        initial_balance = float(initial_balance_row.iloc[0]['I']) if pd.notnull(initial_balance_row.iloc[0]['I']) else 0.0
    
    return start_date, end_date, initial_balance

def is_tva_row(row, last_date):
    """Vérifie si une ligne est une ligne de TVA sans date."""
    is_no_date = pd.isna(row['A']) or not re.match(r'\d{2}\.\d{2}\.\d{4}', str(row['A']))
    has_amount = pd.notnull(row['G']) or pd.notnull(row['H']) or pd.notnull(row['I'])
    is_tva = str(row['D']).startswith(('117', '2200')) or re.search(r'TVA|VAT', str(row['B']), re.IGNORECASE)
    return is_no_date and has_amount and is_tva and last_date is not None

def is_change_row(row):
    """Vérifie si une ligne est une compensation de change."""
    return str(row['B']).lower().startswith('compensation de change')

def process_sheet(df, account_number, start_date, initial_balance):
    """Traite une feuille pour nettoyer les données, incluant les lignes TVA sans date."""
    columns = ['Date', 'Texte', 'Compte', 'Contre écr', 'Code', 'Origine', 'Document', 'Débit', 'Crédit', 'Solde']
    cleaned_data = []
    
    # Ajouter la ligne de solde initial
    debit = initial_balance if initial_balance >= 0 else 0.0
    credit = abs(initial_balance) if initial_balance < 0 else 0.0
    cleaned_data.append({
        'Date': start_date,
        'Texte': 'Report de solde',
        'Compte': account_number,
        'Contre écr': '',
        'Code': '',
        'Origine': '',
        'Document': '',
        'Débit': debit,
        'Crédit': credit,
        'Solde': initial_balance
    })
    
    last_date = start_date
    i = 0
    while i < len(df):
        row = df.iloc[i]
        
        # Ignorer les lignes avec URL
        if pd.isna(row['A']) and str(row['B']).startswith('http'):
            i += 1
            continue
        
        # Traiter les lignes TVA sans date
        if is_tva_row(row, last_date):
            debit = float(row['G']) if pd.notnull(row['G']) else 0.0
            credit = float(row['H']) if pd.notnull(row['H']) else 0.0
            solde = float(row['I']) if pd.notnull(row['I']) else 0.0
            code = str(row['E']) if pd.notnull(row['E']) else ''
            origin = ORIGIN_MAPPING.get(code, 'Écriture manuelle ou inconnue')
            
            cleaned_data.append({
                'Date': last_date,
                'Texte': row['B'],
                'Compte': account_number,
                'Contre écr': str(row['D']) if pd.notnull(row['D']) else '',
                'Code': code,
                'Origine': origin,
                'Document': str(row['F']) if pd.notnull(row['F']) else '',
                'Débit': debit if debit != 0 else '',
                'Crédit': credit if credit != 0 else '',
                'Solde': solde
            })
            i += 1
            continue
        
        # Ignorer les lignes de compensation de change
        if is_change_row(row):
            i += 1
            continue
        
        # Traiter les lignes principales avec date
        if pd.notnull(row['A']) and re.match(r'\d{2}\.\d{2}\.\d{4}', str(row['A'])):
            last_date = row['A']
            debit = float(row['G']) if pd.notnull(row['G']) else 0.0
            credit = float(row['H']) if pd.notnull(row['H']) else 0.0
            solde = float(row['I']) if pd.notnull(row['I']) else 0.0
            code = str(row['E']) if pd.notnull(row['E']) else ''
            origin = ORIGIN_MAPPING.get(code, 'Écriture manuelle ou inconnue')
            
            # Vérifier les lignes suivantes pour TVA ou change
            j = i + 1
            while j < len(df) and (is_tva_row(df.iloc[j], last_date) or is_change_row(df.iloc[j])):
                if is_tva_row(df.iloc[j], last_date):
                    tva_row = df.iloc[j]
                    debit += float(tva_row['G']) if pd.notnull(tva_row['G']) else 0.0
                    credit += float(tva_row['H']) if pd.notnull(tva_row['H']) else 0.0
                    solde = float(tva_row['I']) if pd.notnull(tva_row['I']) else solde
                j += 1
            
            cleaned_data.append({
                'Date': row['A'],
                'Texte': row['B'],
                'Compte': account_number,
                'Contre écr': str(row['D']) if pd.notnull(row['D']) else '',
                'Code': code,
                'Origine': origin,
                'Document': str(row['F']) if pd.notnull(row['F']) else '',
                'Débit': debit if debit != 0 else '',
                'Crédit': credit if credit != 0 else '',
                'Solde': solde
            })
            i = j
        else:
            i += 1
    
    return pd.DataFrame(cleaned_data, columns=columns)

def compute_aggregations(cleaned_sheets):
    """Calcule les agrégations par compte, incluant TVA et totaux mensuels/trimestriels."""
    summary_data = []
    
    for sheet_name, df in cleaned_sheets.items():
        account_number, account_name = sheet_name.split(' ', 1)
        nature = NATURE_MAPPING.get(account_number[0], 'Inconnue')
        
        # Convertir les dates en datetime
        df['Date'] = pd.to_datetime(df['Date'], format='%d.%m.%Y', errors='coerce')
        
        # Totaux généraux
        total_debit = pd.to_numeric(df['Débit'], errors='coerce').fillna(0).sum()
        total_credit = pd.to_numeric(df['Crédit'], errors='coerce').fillna(0).sum()
        
        # Détection des transactions TVA
        vat_mask = (df['Contre écr'].str.startswith(('117', '2200'), na=False)) | \
                   (df['Texte'].str.contains(r'TVA|VAT', case=False, na=False))
        vat_debit = pd.to_numeric(df[vat_mask]['Débit'], errors='coerce').fillna(0).sum()
        vat_credit = pd.to_numeric(df[vat_mask]['Crédit'], errors='coerce').fillna(0).sum()
        net_vat = vat_credit - vat_debit
        
        # Totaux mensuels
        monthly = df.groupby(df['Date'].dt.to_period('M')).agg({
            'Débit': lambda x: pd.to_numeric(x, errors='coerce').fillna(0).sum(),
            'Crédit': lambda x: pd.to_numeric(x, errors='coerce').fillna(0).sum()
        }).reset_index()
        monthly['Date'] = monthly['Date'].astype(str)
        
        # Totaux trimestriels
        quarterly = df.groupby(df['Date'].dt.to_period('Q')).agg({
            'Débit': lambda x: pd.to_numeric(x, errors='coerce').fillna(0).sum(),
            'Crédit': lambda x: pd.to_numeric(x, errors='coerce').fillna(0).sum()
        }).reset_index()
        quarterly['Date'] = quarterly['Date'].astype(str)
        
        summary_data.append({
            'Account Number': account_number,
            'Account Name': account_name,
            'Nature': nature,
            'Total Debit': total_debit,
            'Total Credit': total_credit,
            'Net VAT': net_vat,
            'Monthly Summary': monthly.to_dict('records'),
            'Quarterly Summary': quarterly.to_dict('records')
        })
    
    return pd.DataFrame(summary_data)

def main(input_file):
    xl = pd.ExcelFile(input_file)
    plan_comptable_data = []
    cleaned_sheets = {}
    
    for sheet_name in xl.sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet_name, dtype=str)
        df.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
        
        account_number, account_name = parse_sheet_name(sheet_name)
        if not account_number:
            continue
        
        nature = NATURE_MAPPING.get(account_number[0], 'Inconnue')
        plan_comptable_data.append({
            'Numéro de compte': account_number,
            'Nom de compte': account_name,
            'Nature du compte': nature
        })
        
        start_date, end_date, initial_balance = get_period_and_initial_balance(df)
        cleaned_df = process_sheet(df, account_number, start_date, initial_balance)
        cleaned_sheet_name = f"{account_number} {account_name}"
        cleaned_sheets[cleaned_sheet_name] = cleaned_df
    
    # Créer Plan_Comptable.xlsx
    plan_comptable_df = pd.DataFrame(plan_comptable_data)
    plan_comptable_df.to_excel('Plan_Comptable.xlsx', sheet_name='Plan_Comptable', index=False)
    
    # Créer Comptes_Cleans.xlsx
    with pd.ExcelWriter('Comptes_Cleans.xlsx', engine='xlsxwriter') as writer:
        for sheet_name, df in cleaned_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)  # Truncate sheet name limit
    
    # Créer Summary.xlsx
    with pd.ExcelWriter('Summary.xlsx', engine='xlsxwriter') as writer:
        summary_df = compute_aggregations(cleaned_sheets)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)

if __name__ == "__main__":
    input_file = "GL.xlsx"
    main(input_file)