Attribute VB_Name = "AddCheckboxes"
'skapar kryssrutor i markerade celler
Sub AddCheckboxes()
    On Error Resume Next
    Dim c As Range, myRange As Range
    Set myRange = Selection
    For Each c In myRange.Cells
        ActiveSheet.CheckBoxes.Add(c.Left, c.Top, c.Width, c.Height).Select
            With Selection
                .LinkedCell = c.Address
                .Characters.Text = ""
                .name = c.Address
            End With
            c.Select
            With Selection
                .FormatConditions.Delete
                .FormatConditions.Add Type:=xlExpression, _
                    Formula1:="=" & c.Address & "=x"
                    
                .FormatConditions(1).Font.ColorIndex = 6 'change for other color when ticked
                .FormatConditions(1).Interior.ColorIndex = 6 'change for other color when ticked
                .Font.ColorIndex = 2 'cell background color = White
            End With
        Next
        myRange.Select
End Sub

