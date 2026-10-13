' load_matcap_pivot.vbs
' Loads MATCAP data and exports to CSV for pivot table creation
' Usage: cscript load_matcap_pivot.vbs <input_file> <output_csv>

Option Explicit

Dim objArgs, inputFile, outputFile
Dim objExcel, objWorkbook, objWorksheet
Dim lastRow, lastCol, row, col
Dim outputStream, line, cellValue
Dim sheetName, conNo, npP, typ, headerRow

' Get command line arguments
Set objArgs = WScript.Arguments
If objArgs.Count < 2 Then
    WScript.Echo "ERROR: Missing arguments"
    WScript.Echo "Usage: cscript load_matcap_pivot.vbs <input_file> <output_csv>"
    WScript.Quit 1
End If

inputFile = objArgs(0)
outputFile = objArgs(1)

' Verify input file exists
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(inputFile) Then
    WScript.Echo "ERROR: Input file not found: " & inputFile
    WScript.Quit 1
End If

' Create Excel application
On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Could not create Excel application: " & Err.Description
    WScript.Quit 1
End If
On Error Goto 0

objExcel.Visible = False
objExcel.DisplayAlerts = False

WScript.Echo "Opening workbook: " & inputFile

' Open workbook
On Error Resume Next
Set objWorkbook = objExcel.Workbooks.Open(inputFile, 0, True) ' ReadOnly = True
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Could not open workbook: " & Err.Description
    objExcel.Quit
    Set objExcel = Nothing
    WScript.Quit 1
End If
On Error Goto 0

' Detect sheet name
sheetName = DetectSheetName(objWorkbook, inputFile)
If sheetName = "" Then
    WScript.Echo "ERROR: Could not determine sheet name"
    WScript.Echo "Available sheets:"
    Dim sheet
    For Each sheet In objWorkbook.Sheets
        WScript.Echo "  - " & sheet.Name
    Next
    objWorkbook.Close False
    objExcel.Quit
    Set objWorkbook = Nothing
    Set objExcel = Nothing
    WScript.Quit 1
End If

WScript.Echo "Using sheet: " & sheetName

' Get worksheet
On Error Resume Next
Set objWorksheet = objWorkbook.Sheets(sheetName)
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Could not access sheet: " & Err.Description
    objWorkbook.Close False
    objExcel.Quit
    Set objWorkbook = Nothing
    Set objExcel = Nothing
    WScript.Quit 1
End If
On Error Goto 0

' Get data dimensions using safer method
On Error Resume Next
lastRow = objWorksheet.UsedRange.Rows.Count
lastCol = objWorksheet.UsedRange.Columns.Count
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Could not determine sheet dimensions: " & Err.Description
    objWorkbook.Close False
    objExcel.Quit
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    Set objExcel = Nothing
    WScript.Quit 1
End If
On Error Goto 0

WScript.Echo "Sheet dimensions: " & lastRow & " rows x " & lastCol & " columns"

' Ensure we have minimum required columns (at least up to column 104 = CZ)
If lastCol < 104 Then
    WScript.Echo "WARNING: Sheet only has " & lastCol & " columns, expected at least 104"
End If

' Create output CSV file
On Error Resume Next
Set outputStream = CreateObject("ADODB.Stream")
outputStream.Type = 2 ' adTypeText
outputStream.Charset = "utf-8"
outputStream.Open
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Could not create output stream: " & Err.Description
    objWorkbook.Close False
    objExcel.Quit
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    Set objExcel = Nothing
    WScript.Quit 1
End If
On Error Goto 0

' Write header row
headerRow = "CON_NO,NP/P,Type,7-30 Days (c),1-3 Months (d),3-6 Months (e),6-12 Months (f),1-3 Years (g),3-5 Years (h),Over 5 Years (i),Total"
outputStream.WriteText headerRow & vbCrLf

Dim rowCount, progressInterval
rowCount = 0
progressInterval = 1000  ' Report progress every 1000 rows

WScript.Echo "Processing " & (lastRow - 1) & " data rows..."

' Process data rows (starting from row 2)
For row = 2 To lastRow
    On Error Resume Next

    ' Report progress every 1000 rows
    If (row - 2) Mod progressInterval = 0 And row > 2 Then
        WScript.Echo "  Processed " & (row - 2) & " / " & (lastRow - 1) & " rows..."
    End If
    
    ' Get key values
    conNo = objWorksheet.Cells(row, 1).Value ' Column A
    npP = objWorksheet.Cells(row, 2).Value   ' Column B
    typ = objWorksheet.Cells(row, 4).Value   ' Column D
    
    If Err.Number <> 0 Then
        ' Skip problematic rows
        Err.Clear
    Else
        ' Skip if invalid
        If IsNull(conNo) Or IsEmpty(conNo) Then
            ' Skip empty rows
        ElseIf LCase(Trim(CStr(conNo))) = "total" Then
            ' Skip total row
        ElseIf (npP <> "NP" And npP <> "P") Then
            ' Skip if NP/P is not valid
        ElseIf IsNull(typ) Or IsEmpty(typ) Or Trim(CStr(typ)) = "" Then
            ' Skip if Type is empty
        Else
            ' Calculate bucket values
            Dim v7_30, v1_3m, v3_6m, v6_12m, v1_3y, v3_5y, vOver5y, total
            
            ' 7-30 Days: Column N (14)
            v7_30 = GetNumericValue(objWorksheet.Cells(row, 14).Value)
            
            ' 1-3 Months: Sum O:P (15:16)
            v1_3m = SumRange(objWorksheet, row, 15, 16, lastCol)
            
            ' 3-6 Months: Sum Q:S (17:19)
            v3_6m = SumRange(objWorksheet, row, 17, 19, lastCol)
            
            ' 6-12 Months: Sum T:Y (20:25)
            v6_12m = SumRange(objWorksheet, row, 20, 25, lastCol)
            
            ' 1-3 Years: Sum Z:AW (26:49)
            v1_3y = SumRange(objWorksheet, row, 26, 49, lastCol)
            
            ' 3-5 Years: Sum AX:BU (50:73)
            v3_5y = SumRange(objWorksheet, row, 50, 73, lastCol)
            
            ' Over 5 Years: Sum BV:CZ (74:104)
            vOver5y = SumRange(objWorksheet, row, 74, 104, lastCol)
            
            ' Total
            total = v7_30 + v1_3m + v3_6m + v6_12m + v1_3y + v3_5y + vOver5y
            
            ' Write CSV line
            line = EscapeCSV(conNo) & "," & _
                   EscapeCSV(npP) & "," & _
                   EscapeCSV(typ) & "," & _
                   FormatNumber(v7_30, 2, -1, 0, 0) & "," & _
                   FormatNumber(v1_3m, 2, -1, 0, 0) & "," & _
                   FormatNumber(v3_6m, 2, -1, 0, 0) & "," & _
                   FormatNumber(v6_12m, 2, -1, 0, 0) & "," & _
                   FormatNumber(v1_3y, 2, -1, 0, 0) & "," & _
                   FormatNumber(v3_5y, 2, -1, 0, 0) & "," & _
                   FormatNumber(vOver5y, 2, -1, 0, 0) & "," & _
                   FormatNumber(total, 2, -1, 0, 0)
            
            outputStream.WriteText line & vbCrLf
            rowCount = rowCount + 1
        End If
    End If
    
    On Error Goto 0
Next

WScript.Echo "Completed processing " & (lastRow - 1) & " rows"
WScript.Echo "Valid data rows extracted: " & rowCount

' Save and close
WScript.Echo "Saving output file..."
outputStream.SaveToFile outputFile, 2 ' adSaveCreateOverWrite
outputStream.Close

WScript.Echo "Processed " & rowCount & " valid records"
WScript.Echo "Output saved to: " & outputFile

' Cleanup
objWorkbook.Close False
objExcel.Quit
Set objWorksheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing
Set outputStream = Nothing
Set fso = Nothing

WScript.Echo "SUCCESS: Data exported"
WScript.Quit 0

' ===== Helper Functions =====

Function DetectSheetName(workbook, filename)
    Dim sheet, sheetName
    sheetName = ""
    
    ' Check for Portfolio sheet (CBSL files)
    If InStr(1, filename, "CBSL", vbTextCompare) > 0 Or _
       InStr(1, filename, ".xlsb", vbTextCompare) > 0 Then
        For Each sheet In workbook.Sheets
            If sheet.Name = "Portfolio" Then
                sheetName = "Portfolio"
                Exit For
            End If
        Next
    End If
    
    ' Look for sheet starting with MATCAP
    If sheetName = "" Then
        For Each sheet In workbook.Sheets
            If InStr(1, sheet.Name, "MATCAP", vbTextCompare) > 0 Then
                sheetName = sheet.Name
                Exit For
            End If
        Next
    End If
    
    ' Look for Sheet1
    If sheetName = "" Then
        For Each sheet In workbook.Sheets
            If sheet.Name = "Sheet1" Then
                sheetName = "Sheet1"
                Exit For
            End If
        Next
    End If
    
    ' Use first sheet as fallback
    If sheetName = "" And workbook.Sheets.Count > 0 Then
        sheetName = workbook.Sheets(1).Name
    End If
    
    DetectSheetName = sheetName
End Function

Function GetNumericValue(value)
    On Error Resume Next
    If IsNull(value) Or IsEmpty(value) Then
        GetNumericValue = 0
    ElseIf IsNumeric(value) Then
        GetNumericValue = CDbl(value)
    Else
        GetNumericValue = 0
    End If
    On Error Goto 0
End Function

Function SumRange(worksheet, row, startCol, endCol, maxCol)
    Dim col, value, sum
    sum = 0
    
    ' Don't exceed available columns
    If endCol > maxCol Then
        endCol = maxCol
    End If
    
    On Error Resume Next
    For col = startCol To endCol
        value = worksheet.Cells(row, col).Value
        sum = sum + GetNumericValue(value)
    Next
    On Error Goto 0
    
    SumRange = sum
End Function

Function EscapeCSV(value)
    Dim strValue
    On Error Resume Next
    If IsNull(value) Or IsEmpty(value) Then
        EscapeCSV = ""
    Else
        strValue = CStr(value)
        ' Escape quotes and wrap in quotes if contains comma or quote
        If InStr(strValue, ",") > 0 Or InStr(strValue, """") > 0 Or InStr(strValue, vbCrLf) > 0 Or InStr(strValue, vbLf) > 0 Then
            strValue = Replace(strValue, """", """""")
            EscapeCSV = """" & strValue & """"
        Else
            EscapeCSV = strValue
        End If
    End If
    On Error Goto 0
End Function