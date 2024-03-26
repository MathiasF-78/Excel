Attribute VB_Name = "InsertList"
Sub InsertDoclinks()
Dim objFSO As Object
Dim objFolder As Object
Dim objFile As Object
Dim i As Integer

'Create an instance of the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Get the folder object
Set objFolder = objFSO.GetFolder("C:\Users\mfn\OneDrive - Norrtälje Energi\Fjärrvärme\Effektsignaturer 2021\brev\attachments")
i = 1
'loops through each file in the directory
For Each objFile In objFolder.Files
    'select cell
    Range(Cells(i + 1, 1), Cells(i + 1, 1)).Select
    'create hyperlink in selected cell
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:= _
        objFile.Path, _
        TextToDisplay:=objFile.Name
    i = i + 1
Next objFile
End Sub

