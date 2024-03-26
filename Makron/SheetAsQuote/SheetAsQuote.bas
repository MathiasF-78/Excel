Attribute VB_Name = "SheetAsQuote"
Private Sub CommandButton1_Click()
 PDFname = ActiveSheet.Name
    ChDir ActiveWorkbook.Path & "\"
    Dim fN1 As String
    'Dim fN2 As String
    fN1 = Range("F8").Text
    'fN2 = Range("B5").Text
      fileSaveName = ActiveSheet.Name & "_" & fN1 & "-" & Format(Now, "yyyymmdd")
      'fileSaveName = fN1
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        fileSaveName _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, ignoreprintareas _
        :=False, openafterpublish:=True
        'MsgBox "File Saved " & " " & fileSaveName
End Sub


