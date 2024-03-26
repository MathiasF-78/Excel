Attribute VB_Name = "InsertRowsAdvanced"
Sub Insert_Row()
Dim my As Integer, ur As Integer
On Error GoTo Getout
ur = InputBox("Enter How Many Row Want to Insert")
my = InputBox("Enter How Many Rows Want to Skip")
Application.ScreenUpdating = False
If my = 0 Or ur = 0 Then Exit Sub
On Error GoTo Getout
Range("A" & 2 + my).Select
Do While ActiveCell.Value <> ""
5 Range(ActiveCell, ActiveCell.Offset(ur - 1, 0)).EntireRow.Insert
ActiveCell.Offset(1 + my + ur - 1, 0).Select
Loop
Getout:
Application.ScreenUpdating = True
End Sub
