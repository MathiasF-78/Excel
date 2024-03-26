Attribute VB_Name = "FileNameFromCellContent"
Sub FileNameAsCellContent()

Dim FileName As String
Dim Path As String

Application.DisplayAlerts = False

Path = "C:\" 'Change the directory path here where you want to save the file
FileName = Range("A1").Value & ".xlsx" 'Change extension here

ActiveWorkbook.SaveAs Path & FileName, xlOpenXMLWorkbook 'Change the format here which matches with the extention above. Choose from the following link http://msdn.microsoft.com/en-us/library/office/ff198017.aspx

Application.DisplayAlerts = True

ActiveWorkbook.Close

End Sub
