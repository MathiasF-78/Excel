Attribute VB_Name = "gemener"
Sub gemener()
'konverterar text till gemener inom markering
Dim Rng As Range
For Each Rng In Selection.Cells
If Rng.HasFormula = False Then
Rng.Value = LCase(Rng.Value)
End If
Next Rng
End Sub

