import pandas as pd
import numpy as np


def create_FDL_quarter_df(monthlyDF, fdlreport_file_path):
    """
    Creates FDL quarterly DataFrame with last 3 months of data.
    
    Parameters:
    monthlyDF (DataFrame): Main monthly data containing FDL records
    fdlreport_file_path (str): Path to the FDL report Excel file
    
    Returns:
    DataFrame: Processed FDL loans data for the last 3 months
    """
    
    # Read the status file
    monthlyDF_astatus = pd.read_excel(fdlreport_file_path, skiprows=3)
    
    # Filter for FDL products - FIXED: use .copy() to avoid SettingWithCopyWarning
    monthlyDF = monthlyDF[monthlyDF['Product'] == 'FDL'].copy()
    
    # Map loan dates from status file
    contract_to_date = dict(zip(monthlyDF_astatus['Loan No'], monthlyDF_astatus["Loan Date"]))
    monthlyDF['Loan Date'] = monthlyDF['Contract No'].map(contract_to_date)
    
    # FIXED: Added error handling for date conversion
    monthlyDF['Loan Date'] = pd.to_datetime(monthlyDF['Loan Date'], format='%d/%m/%Y', errors='coerce')
    
    # Remove rows with invalid dates
    monthlyDF = monthlyDF.dropna(subset=['Loan Date'])
    
    # Extract month in 'May-25' format
    monthlyDF['Month_Formatted'] = monthlyDF['Loan Date'].dt.strftime('%b-%y')
    
    # Sort by date (descending)
    monthlyDF = monthlyDF.sort_values('Loan Date', ascending=False)

    # Get the 3 most recent months
    recent_months = monthlyDF['Loan Date'].dt.to_period('M').unique()
    recent_months_sorted = sorted(recent_months, reverse=True)[:3]

    # Filter for last 3 months
    monthlyDF = monthlyDF[monthlyDF['Loan Date'].dt.to_period('M').isin(recent_months_sorted)].copy()
    
    # FIXED: Reset index to avoid index misalignment issues
    monthlyDF = monthlyDF.reset_index(drop=True)
    
    # Create the output DataFrame
    fdl_loans_df = pd.DataFrame()
    
    # Map columns - FIXED: Added .values to ensure proper assignment
    fdl_loans_df['Month'] = monthlyDF['Month_Formatted'].values
    fdl_loans_df['Contract No'] = monthlyDF['Contract No'].values
    fdl_loans_df['Product'] = monthlyDF['Product'].values
    fdl_loans_df['Client Code'] = np.nan
    fdl_loans_df['Equipment'] = np.nan
    fdl_loans_df['Purpose'] = np.nan
    fdl_loans_df['Corporate Clients'] = monthlyDF['Corporate Clients'].values
    fdl_loans_df['Type of Loans (1.3.1.0.0.0)'] = 'Loans against Cash/Deposits'
    fdl_loans_df['Normal/Concessionary'] = 'Normal'
    fdl_loans_df['Frequency'] = monthlyDF['Frequency'].values
    fdl_loans_df['Contract Period'] = monthlyDF['Contract Period'].values
    fdl_loans_df['Tenure (Months)'] = monthlyDF['Tenure (Months)'].values
    
    # FIXED: Corrected typo - 'Contractual Interest rate' (lowercase 'r')
    fdl_loans_df['Minimum Rate (Final)'] = monthlyDF['Contractual Interest rate'].values
    fdl_loans_df['Contractual Interest Rate'] = monthlyDF['Contractual Interest rate'].values
    fdl_loans_df['Base Rate'] = monthlyDF['Contractual Interest rate'].values
    
    fdl_loans_df['Disbursed Amount (Net of DC)'] = monthlyDF['Contract Amount'].values
    fdl_loans_df['Contract Amount'] = monthlyDF['Contract Amount'].values
    fdl_loans_df['Annual Interest Cost'] = monthlyDF['Annual Contract Interest'].values
    fdl_loans_df['EIR (%)'] = monthlyDF['EIR (%)'].values
    fdl_loans_df['Gross Exposure'] = monthlyDF['Gross Outstanding'].values
    fdl_loans_df['PD Category'] = monthlyDF['PD Category'].values
    fdl_loans_df['Collateral/Security Type'] = monthlyDF['Collateral/Security Type'].values
    fdl_loans_df['Initial Valuation'] = monthlyDF['Initial Valuation'].values
    fdl_loans_df['LTV %'] = monthlyDF['LTV %'].values
    fdl_loans_df['WALTV %'] = monthlyDF['WALTV %'].values
    
    print(f"Months included: {sorted(fdl_loans_df['Month'].unique())}")
    print(f"Total records: {len(fdl_loans_df)}")

    return fdl_loans_df