Attribute VB_Name = "AddIndex"
'Let's say you have more than 100 worksheets in your workbook and it's hard to navigate now.
'Don 't worry this macro code will rescue everything.

Sub TableofContent()
Dim i As Long
On Error Resume Next
Application.DisplayAlerts = False
Worksheets("Table of Content").Delete
Application.DisplayAlerts = True
On Error GoTo 0
Sheets.Add(Before:=Sheets(1)).name = "TableOfContent"
'ActiveSheet.name = "Table of Content"
For i = 1 To Sheets.count
With ActiveSheet
.Hyperlinks.Add _
Anchor:=ActiveSheet.Cells(i, 1), _
Address:="", _
SubAddress:="'" & Sheets(i).name & "'!A1", _
ScreenTip:=Sheets(i).name, _
TextToDisplay:=Sheets(i).name
End With
Next i
End Sub

