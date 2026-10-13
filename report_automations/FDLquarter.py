import pandas as pd

# IA Working sheet layout: header row is row 2, so skiprows=1 lines up with how
# this sheet is read everywhere else in the pipeline (e.g. process_summary.py).
IA_WORKING_SKIPROWS = 1

# Output column order matches the FDL Quarter sheet's layout. 'Product' is
# handled separately (see below) since its source value needs normalizing.
FDL_OUTPUT_COLUMNS_BEFORE_PRODUCT = ['Contract No']
FDL_OUTPUT_COLUMNS_AFTER_PRODUCT = [
    'Client Code', 'Equipment', 'Purpose', 'Corporate Clients',
    'Type of Loans (1.3.1.0.0.0)', 'Normal/Concessionary', 'Frequency',
    'Contract Period', 'Tenure (Months)', 'Minimum Rate (Final)',
    'Contractual Interest Rate', 'Base Rate', 'Disbursed Amount (Net of DC)',
    'Contract Amount', 'Annual Interest Cost', 'EIR (%)', 'Gross Exposure',
    'PD Category', 'Collateral/Security Type', 'Initial Valuation', 'LTV %', 'WALTV %',
]


def create_FDL_quarter_df(month_file_pairs):
    """
    Builds the FDL Quarter sheet's data from 3 separate monthly
    "Prod. wise Class. of Loans - <Month> <Year>" files, one row set per month.

    Parameters:
    month_file_pairs (list[tuple[str, str]]): ordered (month_label, file_path)
        pairs - one per month to include, in the order rows should appear in
        the output (e.g. [('Sep-25', ...), ('Aug-25', ...), ('Jul-25', ...)]).
        month_label is written into the output 'Month' column as-is.

    Returns:
    DataFrame: FDL rows for all given months, newest-first (per month_file_pairs order).
    """

    frames = []
    for month_label, file_path in month_file_pairs:
        df = pd.read_excel(file_path, sheet_name='IA Working', skiprows=IA_WORKING_SKIPROWS, engine='pyxlsb')
        # The sheet's actual header has a stray space ("Normal/ Concessionary");
        # normalize it to match the output column name used below.
        df = df.rename(columns={'Normal/ Concessionary': 'Normal/Concessionary'})

        # FDL rows are identified by IA Working's 'Product' column being 'FDL'
        # OR by the resolved loan type, so a row is still caught if one of the
        # two is inconsistent for a given contract.
        product_code = df['Product'].astype(str).str.strip().str.upper()
        type_of_loans = df['Type of Loans (1.3.1.0.0.0)'].astype(str).str.strip()
        is_fdl = (product_code == 'FDL') | (type_of_loans == 'Loans against Cash/Deposits')

        month_df = df[is_fdl].copy()
        month_df['Month'] = month_label
        frames.append(month_df)

    combined = pd.concat(frames, ignore_index=True)

    fdl_loans_df = pd.DataFrame()
    fdl_loans_df['Month'] = combined['Month'].values
    for col in FDL_OUTPUT_COLUMNS_BEFORE_PRODUCT:
        fdl_loans_df[col] = combined[col].values
    # Historical FDL Quarter output always shows the spelled-out product name,
    # regardless of the raw 'FD' code used to find these rows.
    fdl_loans_df['Product'] = 'FDL'
    for col in FDL_OUTPUT_COLUMNS_AFTER_PRODUCT:
        fdl_loans_df[col] = combined[col].values

    # LTV % and WALTV % are calculated in the output sheet itself rather than
    # pulled from IA Working - process_summary.py writes this dataframe
    # starting at sheet row 3 (write_dataframe_to_sheet(ws_fdl, ..., start_row=3)),
    # so row i (0-indexed here) lands on sheet row 3 + i. Column letters below
    # match FDL_OUTPUT_COLUMNS_AFTER_PRODUCT's fixed layout: Q=Contract Amount,
    # W=Initial Valuation, T=Gross Exposure, X=LTV %.
    #   LTV %   = Contract Amount / Initial Valuation
    #   WALTV % = (Gross Exposure / SUM of all rows' Gross Exposure) * LTV %
    n = len(fdl_loans_df)
    last_row = 2 + n
    fdl_loans_df['LTV %'] = [f'=IF(N(W{3 + i})=0,0,Q{3 + i}/W{3 + i})' for i in range(n)]
    fdl_loans_df['WALTV %'] = [f'=T{3 + i}/SUM($T$3:$T${last_row})*X{3 + i}' for i in range(n)]

    print(f"Months included: {[label for label, _ in month_file_pairs]}")
    print(f"Total records: {len(fdl_loans_df)}")

    return fdl_loans_df
