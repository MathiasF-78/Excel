Attribute VB_Name = "Modul5"
Option Explicit
Sub CopyToAllSheets()
  Dim WS As Worksheet
For Each WS In ThisWorkbook.Worksheets
    If Not WS.Name = "deltagare" Then
        If WorksheetFunction.CountA(WS.Rows(1)) > 0 Then _
            WS.Rows(1).Insert
    End If
Next WS

Worksheets.FillAcrossSheets Worksheets("deltagare").Rows(1), xlFillWithAll
End Sub

