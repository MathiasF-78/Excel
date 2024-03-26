Attribute VB_Name = "InsertSequentialNumbers"
Sub fillDown()
With Selection
    .Value = 1
    .Resize(Range("A1").End(4).Row).DataSeries
End With
End Sub
