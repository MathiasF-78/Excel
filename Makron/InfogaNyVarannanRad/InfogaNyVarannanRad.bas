Attribute VB_Name = "InfogaNyVarannanRad"
Sub InsertAlternateRows()
'This code will insert a row after every row in the selection
'This code has been created by Sumit Bansal from trumpexcel.com
Dim Rng As Range
Dim CountRow As Integer
Dim i As Integer
Set Rng = Selection
CountRow = Rng.EntireRow.count

For i = 1 To CountRow
    ActiveCell.Offset(1, 0).EntireRow.Insert
    ActiveCell.Offset(2, 0).Select
Next i

End Sub
