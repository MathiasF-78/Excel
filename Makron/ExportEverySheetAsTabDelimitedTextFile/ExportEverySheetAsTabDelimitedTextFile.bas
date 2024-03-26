Attribute VB_Name = "Modul2"
Sub Export_Each_Sheet_As_TabDelimited_TextFile()
Dim WS As Worksheet
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
