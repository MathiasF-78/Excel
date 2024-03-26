Attribute VB_Name = "AddMultipleSheets"
'Adds requested number of sheets to the workbook

Sub AddMultipleSheets()

 Dim ShtCount As Integer, i As Integer

 ShtCount = Application.InputBox("How Many Sheets you would like to insert?",
 "Add Sheets", , , , , , 1)

 If ShtCount = False Then
  Exit Sub
 Else
   For i = 1 To ShtCount
   Worksheets.Add
   Next i
 End If
End Sub