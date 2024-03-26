Sub ConvertFormulasToValuesInSelection()
   ActiveSheet.Range("D5").Select
    Dim rng As Range
    Range("O1").Formula = "=O2"
    For Each rng In Selection
        If rng.HasFormula Then
            rng.Formula = rng.Value
        End If
    Next rng
End Sub