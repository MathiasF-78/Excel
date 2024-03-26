Attribute VB_Name = "KommentarTillCell"
Sub KommentarTillText()
'kopierar kommentarer till celler inom markering
Dim Rng As Range
Dim WorkRng As Range
On Error Resume Next
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Range", WorkRng.Address, Type:=8)
For Each Rng In WorkRng
    Rng.Value = Rng.NoteText
Next
End Sub
