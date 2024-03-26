Attribute VB_Name = "HighlightMaxValue"
Sub HighlightMaxValue()
Dim rng As Range
For Each rng In Selection
If rng = WorksheetFunction.Max(Selection) Then
rng.Style = "Good"
End If
Next rng
End Sub
