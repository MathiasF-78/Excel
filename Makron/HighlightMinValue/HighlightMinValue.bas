Attribute VB_Name = "HighlightMinValue"
Sub HighlightMinValue()
Dim rng As Range
For Each rng In Selection
If rng = WorksheetFunction.Min(Selection) Then
rng.Style = "Good"
End If
Next rng
End Sub
