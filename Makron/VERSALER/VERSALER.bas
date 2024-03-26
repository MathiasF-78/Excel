Attribute VB_Name = "VERSALER"
Sub VERSALER()
'konverterar text till versaler inom markering
Dim Rng As Range
For Each Rng In Selection.Cells
If Rng.HasFormula = False Then
Rng.Value = UCase(Rng.Value)
End If
Next Rng
End Sub
