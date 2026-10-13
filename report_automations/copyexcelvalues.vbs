Option Explicit

Dim objFSO, objFolder, objSubFolder, objFile
Dim strBasePath, strDateFolder, strSourceFile, strDestFile
Dim objExcel, wbSource, wbDest, wsSource, wsDest
Dim rngSource, rngDest, cell
Dim folderCount, lastRow, i

' Create FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Base path
strBasePath = "C:\CBSL\Script\working\NBD_MF_10_GA_NBD_MF_11_IS"

' Check if base path exists
If Not objFSO.FolderExists(strBasePath) Then
    WScript.Echo "Error: Base path not found - " & strBasePath
    WScript.Quit
End If

' Get the only folder in the base path
Set objFolder = objFSO.GetFolder(strBasePath)

' Find folders that look like date folders (format: DD-MM-YYYY)
folderCount = 0
strDateFolder = ""

For Each objSubFolder In objFolder.SubFolders
    ' Check if folder name matches date pattern (DD-MM-YYYY)
    If IsDateFolder(objSubFolder.Name) Then
        strDateFolder = objSubFolder.Name
        folderCount = folderCount + 1
        WScript.Echo "Found date folder: " & strDateFolder
    End If
Next

' Check if exactly one date folder was found
If folderCount = 0 Then
    WScript.Echo "Error: No date folder found in " & strBasePath
    WScript.Echo "Looking for folder with format: DD-MM-YYYY"
    WScript.Quit
ElseIf folderCount > 1 Then
    WScript.Echo "Error: Multiple date folders found in " & strBasePath
    WScript.Echo "Please ensure only one date folder exists"
    WScript.Quit
End If

' Find files in the date folder
Set objFolder = objFSO.GetFolder(strBasePath & "\" & strDateFolder)

' Find source file (starts with "NBD-MF-01-SOFP")
For Each objFile In objFolder.Files
    If Left(objFile.Name, 15) = "NBD-MF-01-SOFP " And LCase(Right(objFile.Name, 5)) = ".xlsx" Then
        strSourceFile = objFile.Path
        WScript.Echo "Found source file: " & objFile.Name
        Exit For
    End If
Next

' Find destination file (starts with "NBD-MF-10-GA")
For Each objFile In objFolder.Files
    If Left(objFile.Name, 12) = "NBD-MF-10-GA" And LCase(Right(objFile.Name, 5)) = ".xlsx" Then
        strDestFile = objFile.Path
        WScript.Echo "Found destination file: " & objFile.Name
        Exit For
    End If
Next

' Check if files were found
If strSourceFile = "" Then
    WScript.Echo "Error: Source file not found (looking for file starting with 'NBD-MF-01-SOFP ')"
    WScript.Quit
End If

If strDestFile = "" Then
    WScript.Echo "Error: Destination file not found (looking for file starting with 'NBD-MF-10-GA')"
    WScript.Quit
End If

WScript.Echo "Source: " & strSourceFile
WScript.Echo "Destination: " & strDestFile

' Create Excel Application
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

On Error Resume Next

' Open source workbook
Set wbSource = objExcel.Workbooks.Open(strSourceFile)
If Err.Number <> 0 Then
    WScript.Echo "Error opening source file: " & Err.Description
    objExcel.Quit
    WScript.Quit
End If

' Open destination workbook
Set wbDest = objExcel.Workbooks.Open(strDestFile)
If Err.Number <> 0 Then
    WScript.Echo "Error opening destination file: " & Err.Description
    wbSource.Close False
    objExcel.Quit
    WScript.Quit
End If

' Get worksheets
Set wsSource = wbSource.Worksheets("NBD-MF-01-SOFP")
If Err.Number <> 0 Then
    WScript.Echo "Error: Sheet 'NBD-MF-01-SOFP' not found in source file"
    wbSource.Close False
    wbDest.Close False
    objExcel.Quit
    WScript.Quit
End If

Set wsDest = wbDest.Worksheets("NBD-MF-01-SOFP")
If Err.Number <> 0 Then
    WScript.Echo "Error: Sheet 'NBD-MF-01-SOFP' not found in destination file"
    wbSource.Close False
    wbDest.Close False
    objExcel.Quit
    WScript.Quit
End If

' Find the last row with data in source sheet
lastRow = wsSource.Cells(wsSource.Rows.Count, "D").End(-4162).Row ' xlUp = -4162
If lastRow < 1 Then lastRow = 1

WScript.Echo "Processing " & lastRow & " rows..."

' Clear destination columns D:E first
wsDest.Range("D:E").ClearContents

' Copy values cell by cell to handle merged cells
For i = 1 To lastRow
    On Error Resume Next
    
    ' Copy column D
    If Not IsEmpty(wsSource.Cells(i, 4).Value) Then
        wsDest.Cells(i, 4).Value = wsSource.Cells(i, 4).Value
    End If
    
    ' Copy column E
    If Not IsEmpty(wsSource.Cells(i, 5).Value) Then
        wsDest.Cells(i, 5).Value = wsSource.Cells(i, 5).Value
    End If
    
    On Error GoTo 0
Next

' Save and close destination workbook
wbDest.Save
wbDest.Close

' Close source workbook without saving
wbSource.Close False

' Quit Excel
objExcel.Quit

' Clean up objects
Set wsDest = Nothing
Set wsSource = Nothing
Set wbDest = Nothing
Set wbSource = Nothing
Set objExcel = Nothing
Set objFile = Nothing
Set objSubFolder = Nothing
Set objFolder = Nothing
Set objFSO = Nothing

WScript.Echo "Success: Data copied from columns D:E successfully"
WScript.Echo "Date folder used: " & strDateFolder

' Function to check if folder name matches date pattern
Function IsDateFolder(folderName)
    Dim regex
    Set regex = New RegExp
    regex.Pattern = "^\d{2}-\d{2}-\d{4}$"
    IsDateFolder = regex.Test(folderName)
    Set regex = Nothing
End Function