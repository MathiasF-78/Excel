Attribute VB_Name = "RemoveWrapText"
Sub RemoveWrapText()
Cells.Select
Selection.WrapText = False
Cells.EntireRow.AutoFit
Cells.EntireColumn.AutoFit
End Sub
