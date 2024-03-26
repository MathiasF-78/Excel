Attribute VB_Name = "TaBortRadbrytningar"
Sub TaBortRadbrytningar()
'Tar bort radbyte i de celller som har detta formaterat
Cells.Select
Selection.WrapText = False
Cells.EntireRow.AutoFit
Cells.EntireColumn.AutoFit
End Sub
