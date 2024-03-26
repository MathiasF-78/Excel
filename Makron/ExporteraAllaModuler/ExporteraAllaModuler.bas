Attribute VB_Name = "ExporteraAllaModuler"
Sub ExporteraAllaModuler()
For Each cp In ThisWorkbook.VBProject.VBComponents
cp.export "C:\" & cp.name & Switch(cp.Type = 1, ".bas", cp.Type = 3, ".frm", cp.Type = 2, ".cls", cp.Type = 100, ".cls")
Next
End Sub
