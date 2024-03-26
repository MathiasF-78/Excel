Attribute VB_Name = "ExporteraEffektsignaturerv2"
Function SaveSelectToPDFv6()

 
Dim str As String, myfolder As String, myfile As String
Dim FolderPath As String
Dim myValue As Variant
myValue = ThisWorkbook.Worksheets("ber_anl").Range("A5").Value
myValue = Replace(myValue, ":", "_")

'FolderPath = "H:\Effektsignaturer"
ChDir ActiveWorkbook.Path & "\"
'MkDir FolderPath
    Sheets(Array("sign.alla", "sign.behandlad")).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=myValue, _
        openafterpublish:=False, ignoreprintareas:=False
    
MsgBox "All PDF's have been successfully exported."
End Function

