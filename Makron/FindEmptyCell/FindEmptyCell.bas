Attribute VB_Name = "FindEmptyCell"
Sub FindEmptyCell()
    ActiveCell.Offset(1, 0).Select
       Do While Not IsEmpty(ActiveCell)
          ActiveCell.Offset(1, 0).Select
       Loop
End Sub
