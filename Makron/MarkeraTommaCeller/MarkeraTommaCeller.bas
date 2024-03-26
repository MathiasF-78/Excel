Attribute VB_Name = "MarkeraTommaCeller"
Sub MarkeraTommaCeller()
'Markerar alla tomma celler inom markeringen
Dim Dataset As Range
Set Dataset = Selection
Dataset.SpecialCells(xlCellTypeBlanks).Interior.ColorIndex = 36
End Sub
