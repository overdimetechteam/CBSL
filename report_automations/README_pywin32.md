# PyWin32 Excel Automation Guide

This guide explains how to use `pywin32` to directly manipulate `.xlsb` files with Excel COM automation, including writing data and using Ctrl+D (Fill Down) formulas.

## Overview

The `NBD_MF_23_IA.py` script has been updated to use `pywin32` instead of `pyxlsb` and `pandas` for Excel manipulation. This allows you to:

- ✅ **Directly edit `.xlsb` files** without conversion
- ✅ **Write formulas** including VLOOKUP, calculations, etc.
- ✅ **Use Fill Down (Ctrl+D)** functionality programmatically
- ✅ **Manipulate Excel objects** through COM automation
- ✅ **Preserve original file format** and structure

## Installation

```bash
pip install pywin32>=306
```

## Key Features

### 1. Excel COM Initialization

```python
# Initialize Excel COM application
automation.initialize_excel()

# Always close Excel when done
automation.close_excel()
```

### 2. Basic Cell Operations

```python
# Write values to cells
automation.write_cell_value("Sheet Name", row, col, value)

# Read values from cells
value = automation.read_cell_value("Sheet Name", row, col)

# Write formulas
automation.write_cell_formula("Sheet Name", row, col, "=A1+B1")
```

### 3. Fill Down (Ctrl+D) Functionality

```python
# Fill down a formula from start_row to end_row
automation.fill_down_formula("Sheet Name", start_row, end_row, col, formula)

# Example: Fill down VLOOKUP formula
automation.fill_down_formula("IA Working", 3, 100, 5, "=VLOOKUP(A3,NetPortfolio!A:C,2,FALSE)")
```

### 4. VLOOKUP Formula Application

```python
# Apply VLOOKUP formula to a range with automatic Fill Down
automation.apply_vlookup_formula(
    sheet_name="IA Working",
    target_col=3,           # Column C
    lookup_value_col=1,     # Column A (contract number)
    table_array="NetPortfolio!A:C",  # Lookup table
    col_index=2,            # Return column 2 from lookup table
    start_row=3,            # Start from row 3
    end_row=100             # End at row 100
)
```

### 5. Calculation Formula Application

```python
# Apply calculation formula to a range with automatic Fill Down
automation.apply_calculation_formula(
    sheet_name="IA Working",
    target_col=22,          # Column V
    formula="=P3/V3",       # Formula to apply
    start_row=3,            # Start from row 3
    end_row=100             # End at row 100
)
```

### 6. Range Operations

```python
# Clear a range of cells
automation.clear_range("Sheet Name", start_row, end_row, start_col, end_col)

# Read a range of values
values = automation.read_range_values("Sheet Name", start_row, end_row, start_col, end_col)

# Find the last row with data
last_row = automation.find_last_row("Sheet Name", col)
```

## Practical Examples

### Example 1: Basic Data Entry

```python
# Write contract data starting from row 3
for i, contract in enumerate(contracts):
    row = i + 3
    automation.write_cell_value("IA Working", row, 1, contract)      # Column A
    automation.write_cell_value("IA Working", row, 15, amounts[i])   # Column O
    automation.write_cell_value("IA Working", row, 14, rates[i])     # Column N
```

### Example 2: VLOOKUP with Fill Down

```python
# Apply VLOOKUP to look up client codes
automation.apply_vlookup_formula(
    "IA Working", 3, 1, "NetPortfolio!A:C", 2, 3, 100
)
```

### Example 3: Financial Calculations

```python
# Calculate monthly payments and fill down
automation.apply_calculation_formula(
    "IA Working", 22, "=PMT(M3/12, J3, P3)", 3, 100
)
```

### Example 4: Conditional Logic

```python
# Apply different formulas based on conditions
for row in range(3, 101):
    status = automation.read_cell_value("IA Working", row, 2)  # Column B
    
    if status == "FDL":
        automation.write_cell_value("IA Working", row, 24, "Small")
    elif status == "Margin Trading":
        automation.write_cell_value("IA Working", row, 24, "Medium")
```

## File Structure

```
report_automations/
├── NBD_MF_23_IA.py          # Main automation script with pywin32
├── test_pywin32.py          # Basic test script
├── pywin32_examples.py      # Comprehensive examples
└── README_pywin32.md        # This guide
```

## Running the Scripts

### Test Basic Functionality

```bash
python test_pywin32.py
```

### Run Comprehensive Examples

```bash
python pywin32_examples.py
```

### Run Main Automation

```bash
python NBD_MF_23_IA.py
```

## Advantages of PyWin32

1. **Direct File Manipulation**: No need to convert `.xlsb` to `.xlsx`
2. **Formula Support**: Full Excel formula capabilities including VLOOKUP, PMT, etc.
3. **Fill Down**: Programmatic Ctrl+D functionality
4. **Native Excel Features**: Access to all Excel COM automation features
5. **Performance**: Faster than reading/writing entire files to memory
6. **Compatibility**: Works with all Excel file formats

## Important Notes

1. **Excel Installation Required**: PyWin32 requires Microsoft Excel to be installed
2. **Windows Only**: This solution works only on Windows systems
3. **COM Automation**: Uses Excel's COM interface for automation
4. **File Locks**: Excel will lock the file while it's open
5. **Error Handling**: Always use try/finally to ensure Excel is closed

## Error Handling Best Practices

```python
try:
    automation.initialize_excel()
    # Your automation code here
    automation.run_automation()
finally:
    # Always close Excel
    automation.close_excel()
```

## Troubleshooting

### Common Issues

1. **"Excel application not found"**: Ensure Microsoft Excel is installed
2. **"Permission denied"**: Close the file in Excel before running the script
3. **"COM object error"**: Restart the script and ensure Excel is properly closed
4. **"File locked"**: Close the workbook in Excel and try again

### Performance Tips

1. **Batch Operations**: Group multiple cell operations together
2. **Minimize Excel Visibility**: Keep `Visible = False` for better performance
3. **Efficient Ranges**: Use range operations instead of individual cell operations
4. **Proper Cleanup**: Always close Excel to free up system resources

## Next Steps

1. **Customize Formulas**: Modify the formula templates for your specific needs
2. **Add Error Handling**: Implement robust error handling for production use
3. **Optimize Performance**: Batch operations and minimize Excel interactions
4. **Extend Functionality**: Add more Excel automation features as needed

## Support

For issues or questions about the pywin32 implementation:

1. Check the error logs for specific error messages
2. Ensure Excel is properly installed and accessible
3. Verify file paths and permissions
4. Test with the example scripts first

---

**Note**: This implementation replaces the previous pandas/pyxlsb approach with direct Excel COM automation, providing full access to Excel's native capabilities while maintaining the original file format.
