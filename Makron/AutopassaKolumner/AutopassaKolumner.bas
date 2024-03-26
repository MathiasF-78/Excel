Attribute VB_Name = "AutopassaKolumner"
Sub AutopassaKolumner()
'justerar kolumnbredd efter innehåll i hela arbetsbladet
Cells.Select
Cells.EntireColumn.AutoFit
End Sub
