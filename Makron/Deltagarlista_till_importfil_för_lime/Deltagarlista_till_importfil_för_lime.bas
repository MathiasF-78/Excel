Attribute VB_Name = "importDeltagareToLime"
Sub FormatDeltagareForImport()

'=================================================================
' SPLIT WORKSHEET INTO NEW SHEET EVERY 1 000 ROW ---WORKS
'=================================================================

    Dim lngLastRow As Long
    Dim lngNumberOfRows As Long
    Dim lngI As Long
    Dim strMainSheetName As String
    Dim WS As Worksheet
    Dim prevSheet As Worksheet
    
    'Number of rows to split among worksheets
    lngNumberOfRows = 1000
    'Current worksheet in workbook
    Set prevSheet = ThisWorkbook.ActiveSheet
    'First worksheet name
    strMainSheetName = prevSheet.Name
    'Number of rows in worksheet
    lngLastRow = prevSheet.Cells(Rows.Count, 1).End(xlUp).Row
    'Worksheet counter for added worksheets
    lngI = 1
    While lngLastRow > lngNumberOfRows
        Set WS = ThisWorkbook.Worksheets.Add
        With WS
           .Move after:=Worksheets(Worksheets.Count)
           .Name = strMainSheetName + "(" + CStr(lngI) + ")"
        End With

        With prevSheet.Rows(lngNumberOfRows + 1 & ":" & lngLastRow).EntireRow
            .Cut WS.Range("A1")
        End With

        lngLastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        Set prevSheet = WS
        lngI = lngI + 1
    Wend

'=================================================================
' COPY FIRST ROW AND INSERT TO EVERY NEW SHEET---WORKS
'=================================================================

For Each WS In ThisWorkbook.Worksheets
    If Not WS.Name = "deltagare" Then
        If WorksheetFunction.CountA(WS.Rows(1)) > 0 Then _
            WS.Rows(1).Insert
    End If
Next WS

Worksheets.FillAcrossSheets Worksheets("deltagare").Rows(1), xlFillWithAll

'=================================================================
' EXPORT EVERY SHEET AS TAB DELIMITED TEXTFILE ---WORKS
'=================================================================

Dim MyStr1 As String
Dim MyStr2 As String
Dim MyPath As String
Dim SavePath As String
Dim MyDate
Dim MyTime

MyDate = Date    ' MyDate Returns the current system date.
MyTime = Time    ' Returns current system time.

Application.DisplayAlerts = False
Application.ScreenUpdating = False
On Error Resume Next

MyStr1 = Format(MyDate, "DD-MM-YYYY")
'Use MyStr2 If You Require Time Stamp In File Name
'MyStr2 = Format(MyTime, "HH.MM.SS")
' MyPath = "C:\Documents and Settings\Administrator\My Documents\"
MyPath = ThisWorkbook.path
MkDir MyPath & MyStr1 & "_" & ThisWorkbook.Name
SavePath = MyPath & MyStr1 & "_" & ThisWorkbook.Name & "\"

For Each WS In ThisWorkbook.Sheets
WS.Activate
ActiveSheet.Copy
'Exporting Sheet as Tab Delimited Text File To Target Path

ActiveSheet.SaveAs Filename:=SavePath & WS.Name, FileFormat:=xlTextWindows
'Saving a Backup Copy of a Sheet in Target Path
'ActiveSheet.SaveAs Filename:=SavePath & WS.Name, FileFormat:=xlOpenXMLWorkbook
ActiveWorkbook.Close Savechanges:=True
Next WS
Application.DisplayAlerts = True
Application.ScreenUpdating = True
ThisWorkbook.Activate
'ThisWorkbook.Close Savechanges:=True
'Application.Quit

End Sub
