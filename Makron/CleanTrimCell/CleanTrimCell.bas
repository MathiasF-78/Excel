Attribute VB_Name = "CleanTrimCells"
Sub cleanTrim()
'rensar cell på formler och formatera cell inom markering

Dim rng As Range
Dim Area As Range

'ta bort formler i selection
  If Selection.Cells.Count = 1 Then
    Set rng = Selection
  Else
    Set rng = Selection.SpecialCells(xlCellTypeConstants)
  End If

'clean & trim
  For Each Area In rng.Areas
    Area.Value = Evaluate("IF(ROW(" & Area.Address & "),CLEAN(TRIM(" & Area.Address & ")))")
  Next Area
  
End Sub
