Attribute VB_Name = "RaderaTommaBlad"
Sub RaderaTommaBlad()
'Tar bort alla tomma blad i arbetsboken
Dim ws As Worksheet
On Error Resume Next
Application.ScreenUpdating = False
Application.DisplayAlerts = False
For Each ws In Application.Worksheets
If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then
ws.Delete
End If
Next
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
