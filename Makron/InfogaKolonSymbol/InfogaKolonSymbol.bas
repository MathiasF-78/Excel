Attribute VB_Name = "InfogaKolonSymbol"
Sub KolonSymbol()
'adderar kolon till markerade celler MED innehåll, om cellen är tom förblir den tom
Dim Rng As Range
For Each Rng In Selection
Rng.Select
If ActiveCell <> "" Then
ActiveCell.Value = ActiveCell.Value & ":"
End If
Next
End Sub
