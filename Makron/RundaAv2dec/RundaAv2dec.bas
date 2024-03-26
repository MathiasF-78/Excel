Attribute VB_Name = "RundaAv2dec"
Sub RundaAv2dec()
Dim Rng As Range
    For Each Rng In Selection
        Rng.Value = "=Round(" & Mid(Rng.Formula, 2, 9999) & ",2)"
    Next Rng
    
End Sub
