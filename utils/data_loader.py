import pandas as pd
import os
import streamlit as st

def load_data():
    """Load M&A and Investment data from Excel file"""
    try:
        # Try to load from data directory
        file_path = 'data/MedTech_YTD_Standardized.xlsx'
        
        if not os.path.exists(file_path):
            st.error(f"Data file not found at {file_path}")
            # Return empty dataframes with correct structure
            ma_columns = ['Company', 'Acquirer', 'Deal Type (Merger / Acquisition)', 
                         'Technology/Description', 'Deal Value', 'Quarter', 'Month']
            inv_columns = ['Company', 'Funding type (VC / PE)', 'Technology/Description', 
                          'Amount Raised', 'Lead Investors', 'Quarter', 'Month']
            return pd.DataFrame(columns=ma_columns), pd.DataFrame(columns=inv_columns)
        
        ma_data = pd.read_excel(file_path, sheet_name='YTD M&A Activity')
        inv_data = pd.read_excel(file_path, sheet_name='YTD Investment Activity')
        
        return ma_data, inv_data
    
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        ma_columns = ['Company', 'Acquirer', 'Deal Type (Merger / Acquisition)', 
                     'Technology/Description', 'Deal Value', 'Quarter', 'Month']
        inv_columns = ['Company', 'Funding type (VC / PE)', 'Technology/Description', 
                      'Amount Raised', 'Lead Investors', 'Quarter', 'Month']
        return pd.DataFrame(columns=ma_columns), pd.DataFrame(columns=inv_columns)

def save_data(ma_data, inv_data):
    """Save M&A and Investment data to Excel file"""
    try:
        file_path = 'data/MedTech_YTD_Standardized.xlsx'
        
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            ma_data.to_excel(writer, sheet_name='YTD M&A Activity', index=False)
            inv_data.to_excel(writer, sheet_name='YTD Investment Activity', index=False)
        
        return True
    
    except Exception as e:
        st.error(f"Error saving data: {str(e)}")
        return False

def check_duplicates(new_deal, existing_data, deal_type):
    """Check if a deal already exists in the dataset"""
    if deal_type == "M&A":
        # Check by Company and Acquirer
        company_match = existing_data['Company'].str.lower() == new_deal.get('Company', '').lower()
        acquirer_match = existing_data['Acquirer'].str.lower() == new_deal.get('Acquirer', '').lower()
        return (company_match & acquirer_match).any()
    
    else:  # Investment
        # Check by Company and Amount Raised
        company_match = existing_data['Company'].str.lower() == new_deal.get('Company', '').lower()
        amount_match = existing_data['Amount Raised'].astype(str) == str(new_deal.get('Amount Raised', ''))
        return (company_match & amount_match).any()

def format_currency(value):
    """Format currency values to B (billions) or M (millions)"""
    if pd.isna(value) or value == 'Undisclosed':
        return 'Undisclosed'
    
    # If it's already a string, return it
    if isinstance(value, str):
        # Check if already formatted
        if 'B' in value or 'M' in value:
            return value
        
        # Try to extract numeric value
        try:
            # Remove $ and commas
            clean_value = value.replace('$', '').replace(',', '')
            numeric_value = float(clean_value)
        except:
            return value
    else:
        numeric_value = float(value)
    
    # Format as B or M
    if numeric_value >= 1_000_000_000:
        return f"${numeric_value / 1_000_000_000:.1f}B"
    elif numeric_value >= 1_000_000:
        return f"${numeric_value / 1_000_000:.1f}M"
    else:
        return f"${numeric_value:,.0f}"

def extract_numeric_value(value_str):
    """Extract numeric value from currency string for calculations"""
    if pd.isna(value_str) or value_str == 'Undisclosed':
        return 0
    
    if isinstance(value_str, (int, float)):
        return float(value_str)
    
    # Remove $ and commas
    clean_value = str(value_str).replace('$', '').replace(',', '')
    
    # Handle B (billions) and M (millions)
    if 'B' in clean_value:
        return float(clean_value.replace('B', '')) * 1_000_000_000
    elif 'M' in clean_value:
        return float(clean_value.replace('M', '')) * 1_000_000
    
    try:
        return float(clean_value)
    except:
        return 0
