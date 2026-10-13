Option Explicit

Dim objArgs, inputFile, outputFile
Dim objExcel, objWorkbook, objWorksheet
Dim fso, outputStream
Dim lastRow, col, row

' Get command line arguments
Set objArgs = WScript.Arguments
If objArgs.Count < 2 Then
    WScript.Echo "ERROR: Missing arguments"
    WScript.Echo "Usage: cscript debug_cbsl_structure.vbs <cbsl_file> <output_file>"
    WScript.Quit 1
End If

inputFile = objArgs(0)
outputFile = objArgs(1)

' Verify input file exists
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(inputFile) Then
    WScript.Echo "ERROR: Input file not found: " & inputFile
    WScript.Quit 1
End If

' Create output stream for results
Set outputStream = CreateObject("ADODB.Stream")
outputStream.Type = 2
outputStream.Charset = "utf-8"
outputStream.Open

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
Set objWorkbook = objExcel.Workbooks.Open(inputFile, 0, True)
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Could not open workbook: " & Err.Description
    objExcel.Quit
    WScript.Quit 1
End If
On Error Goto 0

' Get Portfolio sheet
On Error Resume Next
Set objWorksheet = objWorkbook.Sheets("Portfolio")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Could not access Portfolio sheet: " & Err.Description
    objWorkbook.Close False
    objExcel.Quit
    WScript.Quit 1
End If
On Error Goto 0

' Get dimensions
lastRow = objWorksheet.UsedRange.Rows.Count
WScript.Echo "Processing " & lastRow & " rows from .xlsb file"

' === DEBUGGING CBSL FILE STRUCTURE ===
outputStream.WriteText "=== DEBUGGING CBSL FILE STRUCTURE ===" & vbCrLf
outputStream.WriteText "Checking first row headers around column AK:" & vbCrLf

Dim colLetters, colLetter, headerValue, colIdx
colLetters = Array("AI", "AJ", "AK", "AL", "AM")
colIdx = 35

For Each colLetter In colLetters
    headerValue = objWorksheet.Range(colLetter & "1").Value
    outputStream.WriteText "  Column " & colLetter & " (index " & colIdx & "): " & SafeString(headerValue) & vbCrLf
    colIdx = colIdx + 1
Next

outputStream.WriteText vbCrLf
outputStream.WriteText "Sample data from row 3:" & vbCrLf

' Get Contract Number from A3
Dim contractNoSample
contractNoSample = objWorksheet.Range("A3").Value
outputStream.WriteText "  A3 (Contract): " & SafeString(contractNoSample) & vbCrLf

' Get sample data from columns around AK
For Each colLetter In colLetters
    Dim dataValue
    dataValue = objWorksheet.Range(colLetter & "3").Value
    outputStream.WriteText "  " & colLetter & "3: " & SafeString(dataValue) & vbCrLf
Next

outputStream.WriteText String(40, "=") & vbCrLf & vbCrLf

' Process first 100 rows for debugging
Dim maxRows, rowIdx, contractNo, nppValue, contractStr
maxRows = 100
If lastRow < maxRows Then
    maxRows = lastRow
End If

outputStream.WriteText "Processing first 100 rows:" & vbCrLf

Dim rowCount, npCount, pCount, naCount
rowCount = 0
npCount = 0
pCount = 0
naCount = 0

For rowIdx = 3 To maxRows
    contractNo = objWorksheet.Range("A" & rowIdx).Value
    nppValue = objWorksheet.Range("AK" & rowIdx).Value
    
    If Not IsNull(contractNo) And Not IsEmpty(contractNo) Then
        contractStr = Trim(UCase(CStr(contractNo)))
        rowCount = rowCount + 1
        
        ' Log first 5 mappings
        If rowCount <= 5 Then
            outputStream.WriteText "Row " & rowIdx & ": Contract=" & contractStr & ", NP/P=" & SafeString(nppValue) & vbCrLf
        End If
        
        ' Count NP/P/NA values
        If IsNull(nppValue) Or IsEmpty(nppValue) Then
            naCount = naCount + 1
        ElseIf Trim(UCase(CStr(nppValue))) = "NP" Then
            npCount = npCount + 1
        ElseIf Trim(UCase(CStr(nppValue))) = "P" Then
            pCount = pCount + 1
        ElseIf Trim(CStr(nppValue)) = "N/A" Then
            naCount = naCount + 1
        End If
    End If
Next

outputStream.WriteText vbCrLf

' Now process ALL rows for complete count
WScript.Echo "Analyzing complete file for NP/P distribution..."

For rowIdx = 3 To lastRow
    nppValue = objWorksheet.Range("AK" & rowIdx).Value
    
    If IsNull(nppValue) Or IsEmpty(nppValue) Then
        ' Already counted in first 100
        If rowIdx > maxRows Then
            naCount = naCount + 1
        End If
    ElseIf Trim(UCase(CStr(nppValue))) = "NP" Then
        If rowIdx > maxRows Then
            npCount = npCount + 1
        End If
    ElseIf Trim(UCase(CStr(nppValue))) = "P" Then
        If rowIdx > maxRows Then
            pCount = pCount + 1
        End If
    ElseIf Trim(CStr(nppValue)) = "N/A" Then
        If rowIdx > maxRows Then
            naCount = naCount + 1
        End If
    End If
Next

' Output final statistics
outputStream.WriteText "Loaded CBSL mapping with " & (lastRow - 2) & " contracts" & vbCrLf
outputStream.WriteText vbCrLf
outputStream.WriteText "NP/P distribution: NP=" & npCount & ", P=" & pCount & ", N/A=" & naCount & vbCrLf

WScript.Echo "NP/P distribution: NP=" & npCount & ", P=" & pCount & ", N/A=" & naCount

' Check other columns for NP/P pattern if column AK has no data
If npCount = 0 And pCount = 0 Then
    outputStream.WriteText vbCrLf
    outputStream.WriteText "WARNING: No NP/P values found in column AK!" & vbCrLf
    outputStream.WriteText "Searching other columns for NP/P pattern..." & vbCrLf
    
    Dim searchCols, searchCol, searchNP, searchP
    searchCols = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", _
                       "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", _
                       "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", _
                       "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO")
    
    For Each searchCol In searchCols
        searchNP = 0
        searchP = 0
        
        ' Check first 100 rows
        For row = 3 To 102
            Dim testValue
            testValue = objWorksheet.Range(searchCol & row).Value
            If Not IsNull(testValue) And Not IsEmpty(testValue) Then
                If Trim(UCase(CStr(testValue))) = "NP" Then
                    searchNP = searchNP + 1
                ElseIf Trim(UCase(CStr(testValue))) = "P" Then
                    searchP = searchP + 1
                End If
            End If
        Next
        
        ' If column has significant NP/P values, report it
        If searchNP > 5 Or searchP > 5 Then
            outputStream.WriteText "  Column " & searchCol & " has NP=" & searchNP & ", P=" & searchP & vbCrLf
            WScript.Echo "  Found NP/P in column " & searchCol & ": NP=" & searchNP & ", P=" & searchP
        End If
    Next
End If

' Save debug output to a separate file
Dim debugOutputFile
debugOutputFile = Replace(outputFile, ".csv", "_debug.txt")
outputStream.SaveToFile debugOutputFile, 2
outputStream.Close

' Now create the actual CSV output file with proper format
Set outputStream = CreateObject("ADODB.Stream")
outputStream.Type = 2
outputStream.Charset = "utf-8"
outputStream.Open

' Write CSV header
outputStream.WriteText "contract_no,npp_value,product_code" & vbCrLf

' Write data for all rows
WScript.Echo "Creating CSV output with contract mappings..."
WScript.Echo "Reading from columns: A (Contract), AK (NP/P), B (Product)"

Dim csvRowCount, npCountCSV, pCountCSV, naCountCSV
csvRowCount = 0
npCountCSV = 0
pCountCSV = 0
naCountCSV = 0
For rowIdx = 3 To lastRow
    On Error Resume Next

    ' Read values using .Value2 to get actual values, not formulas
    contractNo = objWorksheet.Range("A" & rowIdx).Value2
    If Err.Number <> 0 Then
        contractNo = objWorksheet.Range("A" & rowIdx).Value
        Err.Clear
    End If

    nppValue = objWorksheet.Range("AK" & rowIdx).Value2
    If Err.Number <> 0 Then
        nppValue = objWorksheet.Range("AK" & rowIdx).Value
        Err.Clear
    End If

    ' Get product code from column B (adjust if needed)
    Dim productCode
    productCode = objWorksheet.Range("B" & rowIdx).Value2
    If Err.Number <> 0 Then
        productCode = objWorksheet.Range("B" & rowIdx).Value
        Err.Clear
    End If

    On Error Goto 0

    If Not IsNull(contractNo) And Not IsEmpty(contractNo) Then
        contractStr = Trim(CStr(contractNo))

        ' Handle nppValue - check for errors and convert properly
        Dim nppStr
        If IsNull(nppValue) Or IsEmpty(nppValue) Then
            nppStr = "N/A"
        Else
            ' Try to convert to string
            On Error Resume Next
            nppStr = Trim(CStr(nppValue))
            If Err.Number <> 0 Then
                ' If conversion failed (likely an error value), use N/A
                nppStr = "N/A"
                Err.Clear
            ElseIf nppStr = "" Then
                ' If the value is empty after trimming
                nppStr = "N/A"
            End If
            On Error Goto 0
        End If

        ' Handle productCode
        Dim productStr
        If IsNull(productCode) Or IsEmpty(productCode) Then
            productStr = ""
        Else
            ' Try to convert to string
            On Error Resume Next
            productStr = Trim(CStr(productCode))
            If Err.Number <> 0 Then
                ' If conversion failed, use empty string
                productStr = ""
                Err.Clear
            End If
            On Error Goto 0
        End If

        ' Count NP/P distribution in CSV output
        If UCase(nppStr) = "NP" Then
            npCountCSV = npCountCSV + 1
        ElseIf UCase(nppStr) = "P" Then
            pCountCSV = pCountCSV + 1
        ElseIf UCase(nppStr) = "N/A" Then
            naCountCSV = naCountCSV + 1
        End If

        ' Write CSV row
        outputStream.WriteText contractStr & "," & nppStr & "," & productStr & vbCrLf
        csvRowCount = csvRowCount + 1

        ' Log first 5 entries for verification
        If csvRowCount <= 5 Then
            WScript.Echo "  Sample " & csvRowCount & ": Contract=" & contractStr & ", NP/P=" & nppStr & ", Product=" & productStr
        End If
    End If
Next

WScript.Echo ""
WScript.Echo "Created CSV with " & csvRowCount & " contract mappings"
WScript.Echo "CSV NP/P distribution: NP=" & npCountCSV & ", P=" & pCountCSV & ", N/A=" & naCountCSV

' Save CSV output
outputStream.SaveToFile outputFile, 2
outputStream.Close

' Cleanup
objWorkbook.Close False
objExcel.Quit
Set objWorksheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing
Set outputStream = Nothing
Set fso = Nothing

WScript.Echo ""
WScript.Echo "CSV output saved to: " & outputFile
WScript.Echo "Debug output saved to: " & debugOutputFile
WScript.Echo "SUCCESS"
WScript.Quit 0

' ===== Helper Functions =====

Function SafeString(value)
    If IsNull(value) Then
        SafeString = "None"
    ElseIf IsEmpty(value) Then
        SafeString = "None"
    Else
        SafeString = CStr(value)
    End If
End Function