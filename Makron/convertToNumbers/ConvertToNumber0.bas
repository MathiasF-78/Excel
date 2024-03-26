Attribute VB_Name = "ConvertToNumbers"
Sub ConvertTextToNumber()
'converts to number format 0
     With Selection
     Selection.NumberFormat = "0"
        Selection.Value = .Value
    End With

End Sub
