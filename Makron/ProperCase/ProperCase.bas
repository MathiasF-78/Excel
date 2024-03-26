Attribute VB_Name = "ProperCase"
Sub properCase()
'Konverterar Text Till Proper Case Inom Markering
Dim rng As Range
For Each rng In Selection.Cells
If rng.HasFormula = False Then
rng.Value = WorksheetFunction.Proper(rng.Value)
End If
Next rng
End Sub