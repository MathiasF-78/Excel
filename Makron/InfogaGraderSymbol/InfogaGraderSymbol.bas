Attribute VB_Name = "InfogaGraderSymbol"
Sub LäggTillGraderSymbol()
Dim rng As Range
For Each rng In Selection
rng.Select
If ActiveCell <> "" Then
If IsNumeric(ActiveCell.Value) Then
ActiveCell.Value = ActiveCell.Value & "°"
End If
End If
Next
End Sub
