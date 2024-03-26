Attribute VB_Name = "FileBackUp"
Sub FileBackUp()
ThisWorkbook.SaveCopyAs FileName:=ThisWorkbook.Path & _
"" & Format(Date, "yyyy-mm-dd") & " " & _
ThisWorkbook.name
End Sub

