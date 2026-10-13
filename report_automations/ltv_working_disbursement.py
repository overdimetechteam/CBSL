import pandas as pd

# LTV Working - Disbursement sheet layout: title in row 1, quarter-period
# label in row 2, SUM totals in row 3, headers in row 4, data from row 5.
LTV_WORKING_DISBURSEMENT_START_ROW = 5

LTV_WORKING_DISBURSEMENT_COLUMNS = [
    'Contract No', 'Customer Name', 'Product', 'Equipment', 'Gross Exposure',
    'Collateral Type', 'Grant Amount', 'Initial Valuation', 'LTV %', 'WALTV %',
]


def create_ltv_working_disbursement_df(month_file_pairs, mainDF):
    """
    Builds the 'LTV Working - Disbursement' sheet's data: Vehicles and
    Machinery-collateral disbursements across the quarter's 3 monthly
    "Prod. wise Class. of Loans - <Month> <Year>" files, one flat table for
    the whole quarter - unlike Disbursement/FDL Quarter, this sheet has no
    per-row Month column; the quarter period is named once in row 2.

    Parameters:
    month_file_pairs (list[tuple[str, str]]): ordered (month_label, file_path)
        pairs for the quarter's 3 months, as built for FDL Quarter - only
        file_path is used here.
    mainDF: Main/Summary dataframe, used to look up Customer Name by Contract
        No (same lookup the Disbursement sheet uses).

    Returns:
    DataFrame: one row per Vehicles and Machinery-collateral disbursement,
    columns in the sheet's own order. Product/Gross Exposure/Initial
    Valuation/LTV %/WALTV % are written as row-relative Excel formulas,
    matching the sheet's existing formula design, rather than computed here.
    Initial Valuation additionally falls back (via IFNA) to the row's own
    IA Working "Initial Valuation" value when the Summary lookup misses.
    """

    frames = []
    for _, file_path in month_file_pairs:
        df = pd.read_excel(file_path, sheet_name='IA Working', skiprows=1, engine='pyxlsb')
        if 'Initial Valuation' not in df.columns:
            df['Initial Valuation'] = pd.NA
        collateral_type = df['Collateral/Security Type'].astype(str).str.strip()
        frames.append(df[collateral_type == 'Vehicles and Machinery'].copy())

    combined = pd.concat(frames, ignore_index=True)
    combined = combined.dropna(subset=['Contract No'])

    contract_to_name = dict(zip(mainDF['CONTRACT NO'], mainDF['Customer Name']))

    contract_no = combined['Contract No'].values
    rows = range(LTV_WORKING_DISBURSEMENT_START_ROW, LTV_WORKING_DISBURSEMENT_START_ROW + len(combined))

    # Initial Valuation: primarily VLOOKUP'd from Summary (col AY, populated from
    # the current month's C1 & C2 Working file - see process_summary.py). That
    # lookup misses contracts sourced from the prior 2 months' files, or any
    # contract C1 & C2 Working doesn't have - IFNA falls back to this row's own
    # IA Working "Initial Valuation" value (populated by the valuation bots) in
    # those cases.
    ia_working_initial_valuation = pd.to_numeric(combined['Initial Valuation'], errors='coerce').values

    def initial_valuation_formula(r, fallback):
        base = f'VLOOKUP(A{r},Summary!A:AY,51,0)'
        if pd.notna(fallback):
            return f'=_xlfn.IFNA({base},{fallback:.2f})'
        return f'={base}'

    ltv_df = pd.DataFrame({
        'Contract No': contract_no,
        'Customer Name': [contract_to_name.get(c) for c in contract_no],
        'Product': [f'=MID(A{r},3,2)' for r in rows],
        'Equipment': combined['Equipment'].values,
        'Gross Exposure': [f'=_xlfn.IFNA(VLOOKUP(A{r},Summary!A:E,5,0),G{r})' for r in rows],
        'Collateral Type': 'Vehicles and Machinery',
        'Grant Amount': combined['Contract Amount'].values,
        'Initial Valuation': [
            initial_valuation_formula(r, fallback)
            for r, fallback in zip(rows, ia_working_initial_valuation)
        ],
        'LTV %': [f'=G{r}/H{r}' for r in rows],
        'WALTV %': [f'=(E{r}/$E$3)*I{r}' for r in rows],
    }, columns=LTV_WORKING_DISBURSEMENT_COLUMNS)

    print(f"LTV Working - Disbursement records: {len(ltv_df)}")

    return ltv_df
