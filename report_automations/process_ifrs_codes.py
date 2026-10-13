import pandas as pd
import xlwings as xw
import os

# Define the IFRS codes to filter
IFRS_CODES = [
    'BS-1.2.6.0.0.0',
    'BS-1.2.7.0.0.0',
    'BS-1.2.8.0.0.0',
    'BS-1.2.9.0.0.0',
    'BS-1.2.10.0.0.0',
    'BS-1.2.11.0.0.0',
    'BS-1.1.17.0.0.0',
    'BS-1.1.18.0.0.0'
]

def process_ifrs_codes(ifrs_file="ifrs_code.xlsx", nbd_file="NBD-MF-10-GA & NBD-MF-11-IS July 2025.xlsx"):
    """
    Process IFRS codes and update NBD file with missing Account # records.

    Args:
        ifrs_file: Path to the IFRS code Excel file
        nbd_file: Path to the NBD-MF-10-GA & NBD-MF-11-IS Excel file

    Returns:
        bool: True if successful, False otherwise
    """
    try:
        print("\n" + "="*70)
        print("Starting IFRS Code Processing")
        print("="*70)

        print("\nStep 1: Reading Ifrs_code.xlsx (columns A:C)...")
        # Read columns A:C from Ifrs_code.xlsx
        ifrs_df = pd.read_excel(ifrs_file, usecols="A:C")
        print(f"Loaded {len(ifrs_df)} rows from Ifrs_code.xlsx")
        print(f"Columns: {ifrs_df.columns.tolist()}")
        print(f"First few rows:\n{ifrs_df.head()}")

        # Identify the IFRS Code column name
        ifrs_code_col = None
        for col in ifrs_df.columns:
            if 'IFRS' in str(col).upper() and 'CODE' in str(col).upper():
                ifrs_code_col = col
                break

        if ifrs_code_col is None:
            # Try to identify by position or other names
            print("Column names:", ifrs_df.columns.tolist())
            # Assume column B (index 1) is IFRS Code
            ifrs_code_col = ifrs_df.columns[1]
            print(f"WARNING: Could not find 'IFRS Code' column, using column: {ifrs_code_col}")

        print(f"\nStep 2: Filtering for IFRS codes: {IFRS_CODES}")
        # Filter for specific IFRS codes
        filtered_df = ifrs_df[ifrs_df[ifrs_code_col].isin(IFRS_CODES)].copy()
        print(f"Filtered to {len(filtered_df)} rows")
        print(f"Filtered data:\n{filtered_df}")

        # Identify column names for GL Code and Account Name
        gl_code_col = None
        account_name_col = None

        for col in filtered_df.columns:
            col_upper = str(col).upper()
            if 'GL' in col_upper or 'ACCOUNT' in col_upper and '#' in col_upper:
                gl_code_col = col
            elif 'ACCOUNT' in col_upper and 'NAME' in col_upper:
                account_name_col = col
            elif 'NAME' in col_upper and account_name_col is None:
                account_name_col = col

        # If not found, use column positions
        if gl_code_col is None:
            gl_code_col = filtered_df.columns[0]  # Column A
            print(f"Using column A as GL Code: {gl_code_col}")

        if account_name_col is None:
            account_name_col = filtered_df.columns[2]  # Column C
            print(f"Using column C as Account Name: {account_name_col}")

        print(f"\nIdentified columns:")
        print(f"  GL Code column: {gl_code_col}")
        print(f"  IFRS Code column: {ifrs_code_col}")
        print(f"  Account Name column: {account_name_col}")

        print(f"\nStep 3: Reading NBD-MF-10-GA & NBD-MF-11-IS file...")
        # Read column B from Other Assets sheet
        other_assets_df = pd.read_excel(nbd_file, sheet_name="Other Assets", usecols="B")
        print(f"Loaded {len(other_assets_df)} rows from Other Assets sheet")
        print(f"Column: {other_assets_df.columns.tolist()}")

        # Read column B from Other Liabilities sheet
        other_liabilities_df = pd.read_excel(nbd_file, sheet_name="Other Liabilities", usecols="B")
        print(f"Loaded {len(other_liabilities_df)} rows from Other Liabilities sheet")
        print(f"Column: {other_liabilities_df.columns.tolist()}")

        print(f"\nStep 4: Merging column B from both sheets and removing duplicates...")
        # Merge the two column B dataframes
        merged_col_b = pd.concat([other_assets_df, other_liabilities_df], ignore_index=True)
        print(f"Merged to {len(merged_col_b)} rows")

        # Remove duplicates and NaN values
        merged_col_b = merged_col_b.dropna().drop_duplicates()
        print(f"After removing duplicates and NaN: {len(merged_col_b)} rows")

        # Get the column name for column B
        col_b_name = merged_col_b.columns[0]
        print(f"Column B name: {col_b_name}")

        print(f"\nStep 5: Finding missing GL Codes...")
        # Compare merged DF with filtered DF
        # Convert to integers for proper comparison (handles both int and float types)
        merged_values = set(merged_col_b[col_b_name].astype(int))
        filtered_gl_codes_int = filtered_df[gl_code_col].astype(int)

        # Find missing values
        missing_mask = ~filtered_gl_codes_int.isin(merged_values)
        missing_df = filtered_df[missing_mask].copy()

        print(f"Found {len(missing_df)} missing GL Codes before filtering")

        # Filter out Account # codes before 2876 (keep only >= 2876)
        print(f"\nStep 5.5: Filtering Account # >= 2876...")
        missing_df = missing_df[missing_df[gl_code_col] >= 2876].copy()

        print(f"Found {len(missing_df)} missing GL Codes after filtering:")
        print(missing_df)

        if len(missing_df) == 0:
            print("\nNo missing values found. Nothing to add to the sheets.")
            print("\n" + "="*70)
            print("IFRS Code Processing completed (no changes needed)")
            print("="*70)
            return True
        else:
            print(f"\nStep 6: Preparing to insert missing values into NBD file using xlwings...")

            # Use xlwings to open Excel file (this interfaces directly with Excel)
            # xlwings operations work EXACTLY like manual Excel operations
            app = None
            wb = None

            try:
                # Open Excel application (visible=False for background processing)
                app = xw.App(visible=False, add_book=False)
                app.display_alerts = False
                app.screen_updating = False

                # Open the workbook
                wb = app.books.open(os.path.abspath(nbd_file))
                print(f"Opened workbook: {nbd_file}")

                # Process both sheets
                for sheet_name in ["Other Assets", "Other Liabilities"]:
                    print(f"\nProcessing sheet: {sheet_name}")
                    ws = wb.sheets[sheet_name]

                    # Find the first blank cell in Column A from top to bottom
                    insert_row = None
                    max_check_row = 1000  # Check up to row 1000

                    for row in range(1, max_check_row):
                        cell_value = ws.range(f'A{row}').value
                        if cell_value is None or (isinstance(cell_value, str) and cell_value.strip() == ''):
                            insert_row = row
                            print(f"  First blank cell in Column A found at row {insert_row}")
                            break

                    if insert_row is None:
                        insert_row = ws.used_range.last_cell.row + 1
                        print(f"  No blank cell found, using row {insert_row}")

                    # Insert rows for each missing value
                    for _, row_data in missing_df.iterrows():
                        print(f"\n  Inserting row at {insert_row} for Account #: {row_data[gl_code_col]}")

                        # Get the formula from row above (for reference)
                        if insert_row > 1:
                            formula_d_above = ws.range(f'D{insert_row - 1}').formula
                            print(f"  Formula in D{insert_row - 1}: {formula_d_above}")

                        # INSERT ROW using xlwings - this works EXACTLY like Excel's Insert > Row
                        # All formulas in the workbook will be automatically adjusted!
                        ws.range(f'{insert_row}:{insert_row}').api.Insert(Shift=-4121)  # xlShiftDown
                        print(f"  Row inserted at {insert_row} - ALL formulas automatically adjusted!")

                        # Add values to the new row
                        # Column A: IFRS Code
                        ws.range(f'A{insert_row}').value = row_data[ifrs_code_col]
                        print(f"    A{insert_row} = {row_data[ifrs_code_col]}")

                        # Column B: Account # (as text)
                        ws.range(f'B{insert_row}').number_format = '@'  # Set as text format first
                        ws.range(f'B{insert_row}').value = str(row_data[gl_code_col])
                        print(f"    B{insert_row} = {row_data[gl_code_col]} (formatted as text)")

                        # Column C: Account Name
                        ws.range(f'C{insert_row}').value = row_data[account_name_col]
                        print(f"    C{insert_row} = {row_data[account_name_col]}")

                        # Copy formula from row above to column D using Excel's copy
                        if insert_row > 1:
                            # Copy the formula from D(insert_row-1) to D(insert_row)
                            # This preserves relative references and adjusts them automatically
                            source_cell = ws.range(f'D{insert_row - 1}')
                            target_cell = ws.range(f'D{insert_row}')

                            if source_cell.formula and source_cell.formula.startswith('='):
                                source_cell.copy(target_cell)
                                print(f"    D{insert_row} = Copied formula from D{insert_row - 1}")

                        # Add formula to column G: =D{insert_row}
                        ws.range(f'G{insert_row}').formula = f'=D{insert_row}'
                        print(f"    G{insert_row} = =D{insert_row}")

                        # Move to next row for next insertion
                        insert_row += 1

                print(f"\nStep 7: Saving the file...")
                # Save and close
                wb.save()
                wb.close()
                print(f"File saved successfully: {nbd_file}")

            finally:
                # Clean up: close workbook and quit Excel application
                if wb:
                    try:
                        wb.close()
                    except:
                        pass
                if app:
                    try:
                        app.quit()
                    except:
                        pass

            print("\n" + "="*70)
            print("IFRS Code Processing completed successfully!")
            print("="*70)
            return True

    except Exception as e:
        print(f"\nERROR in IFRS Code Processing: {e}")
        import traceback
        traceback.print_exc()
        print("\n" + "="*70)
        print("IFRS Code Processing FAILED")
        print("="*70)
        return False


if __name__ == "__main__":
    # Run as standalone script with default file paths
    process_ifrs_codes()
