Attribute VB_Name = "neweffektexport"
 Sub neweffektexport()
   Range("O2").Copy
   Range("O1").PasteSpecial xlPasteValues
End Sub
      
Function SaveSelectToPDFv6()
Dim str As String, myfolder As String, myfile As String
Dim FolderPath As String
Dim myValue As Variant
'myValue = ThisWorkbook.Worksheets("anl.analys").Range("O1").Value
myValue = ActiveSheet.Range("O1").Value
'myValue = Sheet12.Range("O1").Value
myValue = Replace(myValue, ":", "_")

'FolderPath = "H:\Effektsignaturer"
ChDir ActiveWorkbook.Path & "\"
'MkDir FolderPath
    Sheets(Array("sign.alla", "sign.behandlad")).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=myValue, _
        openafterpublish:=False, ignoreprintareas:=False
    
MsgBox "All PDF's have been successfully exported."
End Function

