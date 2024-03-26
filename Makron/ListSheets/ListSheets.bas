Attribute VB_Name = "Module1"
Sub ListSheets()
 
Dim ws As Worksheet
Dim x As Integer
 
x = 1
 
'Sheets("Sheet1").Range("A:A").Clear
ActiveSheet.Range("A:A").Clear
 
For Each ws In Worksheets
     'Sheets("Sheet1").Cells(x, 1) = ws.Name
     ActiveSheet.Cells(x, 1) = ws.Name
     x = x + 1
Next ws
 
End Sub
