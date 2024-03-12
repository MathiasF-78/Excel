Sub runMakroOnAllSheets_5()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In Worksheets
        ws.Select
        Call MODULNAMN.MAKRONAMN
    Next
    Application.ScreenUpdating = True
End Sub

Sub MAKRONAMN
	COD

End Sub
