Attribute VB_Name = "AddNumbersFromTop"
'Add numbers i cells from top
'Sub AddNumbersFromTop()

  Dim i As Integer

  For i = 1 To 10
    Cells(i, 1).Value = i
  Next i

End Sub