Attribute VB_Name = "MarkeraKommentarer"
Sub MarkeraKommentarer()
'Markerar alla celler med kommentarer med RGB färg (Gult som standard)
ActiveSheet.Cells.SpecialCells(xlCellTypeComments).Interior.ColorIndex = 44
ActiveSheet.Cells.SpecialCells(xlCellTypeComments).Font.ColorIndex = 1
End Sub
