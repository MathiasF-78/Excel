Attribute VB_Name = "AddPrefix"
Sub AddPrefix()
'Updateby20131128
Dim Rng As Range
Dim WorkRng As Range
Dim addStr As String
On Error Resume Next
'xTitleId = "KutoolsforExcel"
Set WorkRng = Application.Selection
'Set WorkRng = Application.inputbox("Range", xTitleId, WorkRng.Address, Type:=8)
addStr = Application.inputbox("Add text", xTitleId, "", Type:=2)
For Each Rng In WorkRng
    'Rng.Value = addStr & " " & Rng.Value ' denna lägger till ett mellanslag efter vald text
    Rng.Value = addStr & Rng.Value
Next
End Sub
