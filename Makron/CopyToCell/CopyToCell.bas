
Sub CopyToCell()
'
' Macro21 Macro
'
'  CopyCell Value from O2 to O1
'
    Range("O2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[1]C"
    Range("A3").Select
    

End Sub
Sub Paste()
'
' Macro21 Macro
'
' Paste formula value as text A2 to A1
'
    Range("O1").Copy
   Range("O1").PasteSpecial xlPasteValues
   
        
End Sub
