Attribute VB_Name = "Modul2"
Sub convertToNumbers()
With Selection
    .NumberFormat = "General"
    .Value = .Value
End With

End Sub
