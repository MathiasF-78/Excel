Attribute VB_Name = "HittaBlank"
Sub HittaBlank()
'markerar celler med endast ett mellanslag som tecken
Dim Rng As Range
Dim myCell As Range
For Each Rng In ActiveSheet.UsedRange
If Rng.Value = " " Then
Rng.Interior.ColorIndex = 36

End If
Next Rng
End Sub
