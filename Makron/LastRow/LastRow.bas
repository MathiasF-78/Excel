Sub LastRow()

  Dim LR As Long

  LR = Cells(Rows.Count, 1).End(xlUp).Row
  MsgBox LR

End Sub