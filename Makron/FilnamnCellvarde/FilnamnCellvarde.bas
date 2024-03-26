Attribute VB_Name = "FilnamnCellvarde"
Sub FilnamnCellvarde()
'Sparar filen med filnamn efter värdet i cell A1
    Dim FName           As String
    Dim FPath           As String
     FPath = "C:"
    FName = ActiveWorkbook &("A1").Text
    ActiveWorkbook.SaveAs Filename:=FPath & "\" & FName
End Sub
