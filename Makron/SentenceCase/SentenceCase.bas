Attribute VB_Name = "SentenceCase"
Sub SentenceCase()
Dim Rng As Range
Dim WorkRng As Range
On Error Resume Next
Set WorkRng = Application.Selection
For Each Rng In WorkRng
    xValue = Rng.Value
    xStart = True
    For i = 1 To VBA.Len(xValue)
        ch = Mid(xValue, i, 1)
        Select Case ch
            Case "."
            xStart = True
            Case "?"
            xStart = True
            Case "a" To "z"
            If xStart Then
                ch = UCase(ch)
                xStart = False
            End If
            Case "A" To "Z"
            If xStart Then
                xStart = False
            Else
                ch = LCase(ch)
            End If
        End Select
        Mid(xValue, i, 1) = ch
    Next
    Rng.Value = xValue
Next
End Sub
