import pandas as pd
import numpy as np
from datetime import datetime, timedelta


def get_disbursement_df(mainDF, NetportfolioDF):
    """
    Process mainDF and NetportfolioDF to return disbursement data for last 3 months.

    Parameters:
    - mainDF: Main dataframe containing disbursement information
    - NetportfolioDF: Net portfolio dataframe containing activation dates

    Returns:
    - disbursementDF: Filtered dataframe with disbursement data from last 3 months
    """

    # Create a dictionary mapping CONTRACT_NO to ACTIVATION_DATE from NetportfolioDF
    activation_date_mapping = dict(zip(
        NetportfolioDF['CONTRACT_NO'],
        NetportfolioDF['ACTIVATION_DATE']
    ))

    # Map ACTIVATION_DATE to mainDF based on Contract No
    mainDF['ACTIVATION_DATE'] = mainDF['CONTRACT NO'].map(activation_date_mapping)

    # Drop rows with NaN ACTIVATION_DATE values
    mainDF = mainDF.dropna(subset=['ACTIVATION_DATE'])

    # Convert ACTIVATION_DATE to datetime, coercing errors to NaT
    mainDF['ACTIVATION_DATE'] = pd.to_datetime(mainDF['ACTIVATION_DATE'], errors='coerce')

    # Drop any rows where conversion failed (resulted in NaT)
    mainDF = mainDF.dropna(subset=['ACTIVATION_DATE'])

    # Calculate date range for last 3 months excluding current month
    current_date = datetime.now()
    # First day of current month
    first_day_current_month = current_date.replace(day=1)
    # Last day of 3 months ago (end of the 3-month period)
    end_date = first_day_current_month - timedelta(days=1)
    # First day of 3 months ago (start 3 months back from first day of current month)
    start_date = (first_day_current_month - timedelta(days=90)).replace(day=1)

    # Normalize dates to remove time component for accurate comparison
    start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
    end_date = end_date.replace(hour=23, minute=59, second=59, microsecond=999999)

    # Debug: Print date range and sample dates
    print(f"Filtering date range: {start_date.date()} to {end_date.date()}")
    print(f"Total records before filtering: {len(mainDF)}")
    if len(mainDF) > 0:
        print(f"Sample ACTIVATION_DATE values: {mainDF['ACTIVATION_DATE'].head(10).tolist()}")
        print(f"Min date: {mainDF['ACTIVATION_DATE'].min()}, Max date: {mainDF['ACTIVATION_DATE'].max()}")

    # Filter mainDF for last 3 months excluding current month
    disbursementDF = mainDF[
        (mainDF['ACTIVATION_DATE'] >= start_date) &
        (mainDF['ACTIVATION_DATE'] <= end_date)
    ].copy()

    print(f"Total records after filtering: {len(disbursementDF)}")

    # Create Month column with full month name
    disbursementDF['Month'] = disbursementDF['ACTIVATION_DATE'].dt.strftime('%B')

    # Extract required columns
    columns_to_extract = [
        'CLIENT NO',
        'CONTRACT NO',
        'Customer Name',
        'Grant Amount',
        'Micro/Small/Medium',
        'Gender',
        'Sector',
        'Sub-Sector',
        'CBSL Sector',
        'Month'
    ]

    disbursementDF = disbursementDF[columns_to_extract]

    return disbursementDF