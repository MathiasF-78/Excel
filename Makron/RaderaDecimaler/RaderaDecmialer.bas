Attribute VB_Name = "RaderaDecmialer"
Sub RaderaDecimaler()
'kopierar och rundar upp varde, klistrar in varde utan decimaler
Dim Rng As Range
    For Each Rng In Selection
        Rng.Value = "=Round(" & Mid(Rng.Formula, 2, 9999) & ",0)"
    Next Rng
    
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues

End Sub
