import pandas as pd


DISBURSEMENT_SHEET_START_ROW = 5


def get_disbursement_df(month_file_pairs, mainDF, SectorReportDF):
    """
    Builds the Disbursement sheet's data from the quarter's 3 monthly
    "Prod. wise Class. of Loans - <Month> <Year>" files (current month plus
    the previous two) - the same file set used for FDL Quarter and LTV
    Working - Disbursement - reading each month's own IA Working sheet
    directly rather than filtering a single month's snapshot by activation
    date.

    Parameters:
    month_file_pairs (list[tuple[str, str]]): ordered (month_label, file_path)
        pairs - one per month to include, in the order rows should appear in
        the output (e.g. [('September', file for Sep), ('August', file for
        Aug), ('July', file for Jul)]). month_label is written into the
        output 'Month' column as-is.
    mainDF: Main/Summary dataframe, used only to enrich each IA Working
        contract with Customer Name/Sector/Sub-Sector (fields not present in
        IA Working), keyed by Contract No.
    SectorReportDF: the "1.Sector Report - <Month> <Year>" file's 'Portfolio'
        sheet, used as the source of truth for CBSL Sector (its 'CBSL Sector
        1 Final' column, keyed by Contract No) - not mainDF's own CBSL Sector,
        which can be blank for contracts outside the Summary snapshot.

    Returns:
    DataFrame: Disbursement sheet rows for all given months, newest-first
    (per month_file_pairs order).
    """

    frames = []
    for month_label, file_path in month_file_pairs:
        df = pd.read_excel(file_path, sheet_name='IA Working', skiprows=1, engine='pyxlsb')
        # Exclude Margin Trading rows (IA Working carries a "Margin trading"
        # subtotal row with Product="Margin Trading" - not a real contract,
        # and Margin Trading is out of scope for this sheet either way).
        product = df['Product'].astype(str).str.strip()
        df = df[product.str.upper() != 'MARGIN TRADING']
        month_df = df[[
            'Contract No', 'Client Code', 'Disbursed Amount (Net of DC)', 'MSME Classification'
        ]].dropna(subset=['Contract No']).copy()
        month_df['Month'] = month_label
        frames.append(month_df)

    combined = pd.concat(frames, ignore_index=True)

    disbursementDF = combined.rename(columns={
        'Contract No': 'CONTRACT NO',
        'Client Code': 'CLIENT NO',
        'Disbursed Amount (Net of DC)': 'Disbursement',
        'MSME Classification': 'Micro/Small/Medium',
    })

    # Customer Name/Sector/Sub-Sector are not in IA Working, so they're
    # enriched from mainDF, keyed by Contract No.
    contract_to_name = dict(zip(mainDF['CONTRACT NO'], mainDF['Customer Name']))
    contract_to_sector = dict(zip(mainDF['CONTRACT NO'], mainDF['Sector']))
    contract_to_subsector = dict(zip(mainDF['CONTRACT NO'], mainDF['Sub-Sector']))

    disbursementDF['Customer Name'] = disbursementDF['CONTRACT NO'].map(contract_to_name)
    disbursementDF['Sector'] = disbursementDF['CONTRACT NO'].map(contract_to_sector)
    disbursementDF['Sub-Sector'] = disbursementDF['CONTRACT NO'].map(contract_to_subsector)

    # CBSL Sector comes straight from the Sector Report file's 'CBSL Sector 1
    # Final' column, keyed by Contract No - not from mainDF's own CBSL
    # Sector, which can be blank for contracts outside the Summary snapshot.
    contract_to_cbslsector = dict(zip(SectorReportDF['Contract No'], SectorReportDF['CBSL Sector 1 Final']))
    disbursementDF['CBSL Sector'] = disbursementDF['CONTRACT NO'].map(contract_to_cbslsector)

    # Gender formula mirrors the Summary sheet's logic, re-pointed at the
    # Disbursement sheet's own columns: Client Code is column A here (vs B
    # on Summary) and Customer Name is column C here (vs AT on Summary).
    # Company codes (Client Code starts with '2') get an empty Gender rather
    # than the Summary sheet's BusinessGender lookup/"No Data" fallback.
    disbursementDF['Gender'] = [
        f'=IF(LEFT(A{r},1)="2","",'
        f'IF(LEFT(C{r},3)="Mr.","Male",'
        f'IF(LEFT(C{r},3)="Rev","Male",'
        f'IF(LEFT(C{r},4)="Miss","Female",'
        f'IF(LEFT(C{r},3)="Ms.","Female",'
        f'IF(LEFT(C{r},4)="Mrs.","Female"))))))'
        for r in range(DISBURSEMENT_SHEET_START_ROW, DISBURSEMENT_SHEET_START_ROW + len(disbursementDF))
    ]

    # Extract required columns
    columns_to_extract = [
        'CLIENT NO',
        'CONTRACT NO',
        'Customer Name',
        'Disbursement',
        'Micro/Small/Medium',
        'Gender',
        'Sector',
        'Sub-Sector',
        'CBSL Sector',
        'Month'
    ]

    disbursementDF = disbursementDF[columns_to_extract]

    print(f"Months included: {[label for label, _ in month_file_pairs]}")
    print(f"Total records: {len(disbursementDF)}")

    return disbursementDF
