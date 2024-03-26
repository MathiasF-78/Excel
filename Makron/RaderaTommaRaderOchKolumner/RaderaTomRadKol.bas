Attribute VB_Name = "RaderaTomRadKol"
Sub RaderaTomRadKol()
'raderar samtliga tomma rader och kolumner i arbetsbladet
    'Declare your variables.
        Dim myRange As Range
        Dim iCounter As Long
    'Define the target Range.
        Set myRange = ActiveSheet.UsedRange
        'Start reverse looping through the range of Rows.
        For iCounter = myRange.Rows.count To 1 Step -1
    'If entire row is empty then delete it.
           If Application.CountA(Rows(iCounter).EntireRow) = 0 Then
               Rows(iCounter).Delete
               'Remove comment to See which are the empty rows
               'MsgBox "row " & iCounter & " is empty"
           End If
    'Increment the counter down
        Next iCounter
    'Step 6:  Start reverse looping through the range of Columns.
        For iCounter = myRange.Columns.count To 1 Step -1
    'Step 7: If entire column is empty then delete it.
               If Application.CountA(Columns(iCounter).EntireColumn) = 0 Then
                Columns(iCounter).Delete
               End If
    'Step 8: Increment the counter down
        Next iCounter
End Sub
