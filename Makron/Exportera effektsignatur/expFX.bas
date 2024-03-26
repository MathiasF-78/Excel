Attribute VB_Name = "expFX"
Function expefX()
  Dim str As String, myfolder As String, myfile As String
  Dim FolderPath As String
  Dim myValue As Variant
  Dim Filename As Variant
   
   Range("O2").Copy
  Range("O1").PasteSpecial xlPasteValues

   Filename = ActiveSheet.Range("O1").Value

'myValue = ThisWorkbook.Worksheets("anl.analys").Range("O1").Value
'myValue = ActiveSheet.Range("O1").Text
'myValue = Sheet12.Range("O1").Value
myValue = Replace(Filename, ":", "_")

'FolderPath = ActiveSheet.Range("O1").Text
ChDir ActiveWorkbook.Path & "\"
'MkDir FolderPath
    Sheets(Array("sign.alla", "sign.behandlad")).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=myValue, _
        openafterpublish:=False, ignoreprintareas:=False
    
MsgBox "All PDF's have been successfully exported."
End Function
'fungerar. sparar som cellvärde i O1

