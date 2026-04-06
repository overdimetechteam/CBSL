import pandas as pd
import numpy as np

def create_fixed_loans_df(lrfile_path, monthlyDF, maindflen = 81689):

    lr_df = pd.read_excel(lrfile_path,skiprows=3)
    contract_to_garnt= dict(zip(monthlyDF["Contract No"], monthlyDF["Contract Amount"]))
    contract_to_initial= dict(zip(monthlyDF["Contract No"], monthlyDF["Initial Valuation"]))

    fixed_loans_df = pd.DataFrame()
    
    # Assign columns with NaN values
    fixed_loans_df['Contract No'] = lr_df['Loan No']
    fixed_loans_df['Client No'] = np.nan#TODO: check if its client no or client code
    fixed_loans_df['Branch'] = lr_df['Unnamed: 16']
    fixed_loans_df['Outstanding'] = lr_df['Loan Balance']
    fixed_loans_df['Gross Outstanding'] = lr_df['Loan Balance']
    fixed_loans_df['Original DPD'] = np.nan
    fixed_loans_df['P/NP'] = "P"
    fixed_loans_df['Model Provision'] = np.nan
    fixed_loans_df['100% Provision Contracts'] = np.nan
    fixed_loans_df['Final Provision'] = np.nan
    fixed_loans_df['Yard Contracts'] = "NON-YARD"
    fixed_loans_df['Lease/Loan'] = np.nan
    fixed_loans_df['Age Bucket alligned to Staging'] = np.nan
    fixed_loans_df['Arrears'] = np.nan
    fixed_loans_df['EQT_DESC'] = np.nan
    fixed_loans_df['Auto/Non Auto'] = np.nan
    fixed_loans_df['Reigion'] = np.nan
    fixed_loans_df['PD Category'] = np.nan
    fixed_loans_df['NPA Product Category'] = np.nan
    fixed_loans_df['IIS'] = np.nan
    fixed_loans_df['Total Net Exposure-90 Model'] = lr_df['Loan Balance']
    fixed_loans_df['Loss %'] = np.nan
    fixed_loans_df['DP'] = np.nan
    fixed_loans_df['Provision for DP'] = np.nan
    fixed_loans_df['EXP+DP'] = lr_df['Loan Balance']
    fixed_loans_df['IMP+Pro.DP'] = np.nan
    fixed_loans_df['Product'] = np.nan
    fixed_loans_df['FC'] = np.nan
    fixed_loans_df['TA (6)'] = np.nan
    fixed_loans_df['Stage'] = 1
    fixed_loans_df['Frequency'] = "M"
    fixed_loans_df['Repayment Cycle'] = "Monthly Basis"
    fixed_loans_df['CBSL DPD'] = 0
    fixed_loans_df['Provision Category'] = "Performing"
    fixed_loans_df['Provision - CBSL Guideline'] = 0
    fixed_loans_df['CBSL P/NP'] = "P"
    fixed_loans_df['Forbone Loans (Yes/No)'] = "No"
    fixed_loans_df['EXP+DP-IIS'] = [f'=Y{i + 3 + maindflen}-T{i + 3 + maindflen}' for i in range(len(fixed_loans_df))]
    fixed_loans_df['Gross-Imp-IIS'] = [f'=E{i + 3 + maindflen}-J{i + 3 + maindflen}-T{i + 3 + maindflen}' for i in range(len(fixed_loans_df))]
    fixed_loans_df['Exp+DP-IMP-IIS'] = [f'=Y{i + 3 + maindflen}-Z{i + 3 + maindflen}-T{i + 3 + maindflen}' for i in range(len(fixed_loans_df))]
    fixed_loans_df['Collateral/Security Type'] = "Deposits (Cash-Backed)"
    fixed_loans_df['Sector'] = np.nan #TODO selenium
    fixed_loans_df['Sub-Sector'] = np.nan#TODO selenium
    fixed_loans_df['CBSL Sector'] = np.nan
    fixed_loans_df['CustomerDistrict(ClientMain)'] = np.nan
    fixed_loans_df['Customer Name'] = np.nan
    fixed_loans_df['Micro/Small/Medium'] = np.nan
    fixed_loans_df['Gender'] = np.nan
    fixed_loans_df['Old Contract No (Before Reschedule)'] = np.nan
    fixed_loans_df['Grant Amount'] = fixed_loans_df['Contract No'].map(contract_to_garnt)
    fixed_loans_df['Initial Valuation'] = fixed_loans_df['Contract No'].map(contract_to_initial)
    fixed_loans_df['LTV %'] = np.where(fixed_loans_df['Initial Valuation'].isna() | fixed_loans_df['Grant Amount'].isna(), np.nan, fixed_loans_df['Grant Amount'] / fixed_loans_df['Initial Valuation'])
    fixed_loans_df['WALTV %'] = [f'=(Y{i + 3 + maindflen}/SUM($Y${3 + maindflen}:$Y${3 + maindflen + len(fixed_loans_df)})))*AZ{i + 3 + maindflen}' for i in range(len(fixed_loans_df))]
    print(fixed_loans_df.shape)

    return fixed_loans_df
