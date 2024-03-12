Attribute VB_Name = "AddSuffix"
Sub AddSuffix()
'Updateby20131128
Dim Rng As Range
Dim WorkRng As Range
Dim addStr As String
On Error Resume Next
'xTitleId = "KutoolsforExcel"
Set WorkRng = Application.Selection
'Set WorkRng = Application.inputbox("Range", xTitleId, WorkRng.Address, Type:=8)
addStr = Application.InputBox("Add text", xTitleId, "", Type:=2)
For Each Rng In WorkRng
    Rng.Value = Rng.Value & addStr
Next
End Sub
