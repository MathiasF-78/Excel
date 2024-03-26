Attribute VB_Name = "MarkeraDubletter"
Sub markeraDubletter()
'markerar dubletter inom markeringen med grön fyllning
    'Declare your variables
        Dim myRange As Range
        Dim myCell As Range
    'Define the target Range.
        Set myRange = Selection
    'Start looping through the range.
        For Each myCell In myRange
    'Ensure the cell has Text formatting.
            If WorksheetFunction.CountIf(myRange, myCell.Value) > 1 Then
                myCell.Interior.ColorIndex = 36
            End If
    'Get the next cell in the range
        Next myCell
End Sub
