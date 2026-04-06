' VBScript to copy D14 value from DM file to LA file C7 cell
' Usage: cscript copy_deposit_liability_value.vbs "source_file_path" "target_file_path"
'
' Source: NBD-WF-18-DM sheet, D14 cell
' Target: NBD-WF-15-LA sheet, C7 cell

Option Explicit

Dim objExcel, sourceWb, targetWb, sourceSheet, targetSheet
Dim sourceFile, targetFile, d14Value
Dim objFSO, objArgs, objShell

' Get command line arguments
Set objArgs = WScript.Arguments
If objArgs.Count < 2 Then
    WScript.Echo "ERROR: Missing arguments"
    WScript.Echo "Usage: cscript copy_deposit_liability_value.vbs ""source_file"" ""target_file"""
    WScript.Quit 1
End If

sourceFile = objArgs(0)
targetFile = objArgs(1)

' Verify files exist
Set objFSO = CreateObject("Scripting.FileSystemObject")
If Not objFSO.FileExists(sourceFile) Then
    WScript.Echo "ERROR: Source file not found: " & sourceFile
    WScript.Quit 1
End If

If Not objFSO.FileExists(targetFile) Then
    WScript.Echo "ERROR: Target file not found: " & targetFile
    WScript.Quit 1
End If

' Close any existing Excel instances that might be locking files
WScript.Echo "Checking for existing Excel instances..."
On Error Resume Next
Set objShell = CreateObject("WScript.Shell")
objShell.Run "taskkill /F /IM EXCEL.EXE", 0, True
WScript.Sleep 2000  ' Wait longer for Excel to fully close

' Double-check by trying to release any COM references
Dim existingExcel
Set existingExcel = GetObject(, "Excel.Application")
If Not existingExcel Is Nothing Then
    WScript.Echo "Found existing Excel instance, closing it..."
    existingExcel.Quit
    Set existingExcel = Nothing
    WScript.Sleep 2000
End If
Err.Clear
On Error GoTo 0

' Create new Excel application instance
WScript.Echo "Creating Excel application..."
On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Failed to create Excel application: " & Err.Description
    WScript.Quit 1
End If
Err.Clear
On Error GoTo 0

objExcel.Visible = False
objExcel.DisplayAlerts = False
objExcel.ScreenUpdating = False
objExcel.EnableEvents = False
objExcel.AskToUpdateLinks = False

' Add small delay after Excel creation
WScript.Sleep 500

' Open source file (DM file) and get D14 value
WScript.Echo "Opening source file: " & sourceFile
On Error Resume Next

' Attempt 1: Try opening with CorruptLoad parameter (xlRepairFile = 1)
WScript.Echo "Attempt 1: Opening with repair mode..."
Set sourceWb = objExcel.Workbooks.Open(sourceFile, 0, True, , , , True, , , , , , , , 1)
If Err.Number <> 0 Then
    WScript.Echo "Attempt 1 failed: " & Err.Description
    Err.Clear
    WScript.Sleep 1000

    ' Attempt 2: Try opening with CorruptLoad = xlExtractData (2)
    WScript.Echo "Attempt 2: Opening with data extraction mode..."
    Set sourceWb = objExcel.Workbooks.Open(sourceFile, 0, True, , , , True, , , , , , , , 2)
    If Err.Number <> 0 Then
        WScript.Echo "Attempt 2 failed: " & Err.Description
        Err.Clear
        WScript.Sleep 1000

        ' Attempt 3: Standard open with minimal parameters
        WScript.Echo "Attempt 3: Standard open..."
        Set sourceWb = objExcel.Workbooks.Open(sourceFile, False, True)
        If Err.Number <> 0 Then
            WScript.Echo "Attempt 3 failed: " & Err.Description
            Err.Clear
            WScript.Sleep 1000

            ' Attempt 4: Basic open
            WScript.Echo "Attempt 4: Basic open..."
            Set sourceWb = objExcel.Workbooks.Open(sourceFile)
            If Err.Number <> 0 Then
                WScript.Echo "ERROR: All open attempts failed: " & Err.Description
                objExcel.Quit
                Set objExcel = Nothing
                WScript.Quit 1
            End If
        End If
    End If
End If
On Error GoTo 0

WScript.Echo "Source file opened successfully"

' Find NBD-WF-18-DM sheet
Set sourceSheet = Nothing
On Error Resume Next
Set sourceSheet = sourceWb.Sheets("NBD-WF-18-DM")
On Error GoTo 0

If sourceSheet Is Nothing Then
    WScript.Echo "ERROR: Sheet 'NBD-WF-18-DM' not found in source file"
    WScript.Echo "Available sheets:"
    Dim ws
    For Each ws In sourceWb.Sheets
        WScript.Echo "  - " & ws.Name
    Next
    sourceWb.Close False
    objExcel.Quit
    Set objExcel = Nothing
    WScript.Quit 1
End If

' Force recalculation of the sheet first
WScript.Echo "Recalculating sheet..."
On Error Resume Next
sourceSheet.Calculate
objExcel.CalculateFull
WScript.Sleep 500
On Error GoTo 0

' Get D14 value (this will get the calculated value, not the formula)
d14Value = sourceSheet.Range("D14").Value
WScript.Echo "D14 value extracted: " & d14Value

If IsEmpty(d14Value) Or IsNull(d14Value) Or d14Value = 0 Then
    WScript.Echo "WARNING: D14 value is empty, null, or zero - trying full recalculation"

    ' Try full workbook recalculation
    On Error Resume Next
    sourceWb.RefreshAll
    objExcel.CalculateFullRebuild
    WScript.Sleep 1000
    On Error GoTo 0

    d14Value = sourceSheet.Range("D14").Value
    WScript.Echo "D14 value after full recalculation: " & d14Value
End If

' Close source file
sourceWb.Close False
Set sourceWb = Nothing

If IsEmpty(d14Value) Or IsNull(d14Value) Then
    WScript.Echo "ERROR: Could not extract D14 value"
    objExcel.Quit
    Set objExcel = Nothing
    WScript.Quit 1
End If

' Open target file (LA file) and update C7
WScript.Echo "Opening target file: " & targetFile
On Error Resume Next
Set targetWb = objExcel.Workbooks.Open(targetFile, 0, False)
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Failed to open target file: " & Err.Description
    objExcel.Quit
    Set objExcel = Nothing
    WScript.Quit 1
End If
On Error GoTo 0

WScript.Echo "Target file opened successfully"

' Find NBD-WF-15-LA sheet
Set targetSheet = Nothing
On Error Resume Next
Set targetSheet = targetWb.Sheets("NBD-WF-15-LA")
On Error GoTo 0

If targetSheet Is Nothing Then
    WScript.Echo "ERROR: Sheet 'NBD-WF-15-LA' not found in target file"
    WScript.Echo "Available sheets:"
    For Each ws In targetWb.Sheets
        WScript.Echo "  - " & ws.Name
    Next
    targetWb.Close False
    objExcel.Quit
    Set objExcel = Nothing
    WScript.Quit 1
End If

' Update C7 with D14 value
WScript.Echo "Updating C7 with value: " & d14Value
targetSheet.Range("C7").Value = d14Value

' Save and close target file
WScript.Echo "Saving target file..."
targetWb.Save
targetWb.Close False
Set targetWb = Nothing

' Clean up
objExcel.Quit
Set objExcel = Nothing

WScript.Echo "SUCCESS: D14 value (" & d14Value & ") copied to C7"
WScript.Quit 0
