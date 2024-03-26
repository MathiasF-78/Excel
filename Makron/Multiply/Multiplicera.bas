Attribute VB_Name = "Multiplicera"
Sub Multiply()
Dim Rng As Range
Dim c As Integer
c = InputBox("Enter number to multiple", "Input Required")
For Each Rng In Selection
If WorksheetFunction.IsNumber(Rng) Then
Rng.Value = Rng * c
Else
End If
Next Rng
End Sub
