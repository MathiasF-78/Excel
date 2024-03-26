Attribute VB_Name = "SumByYear"
Sub sumByYear()
'declare a variable
Dim ws As Worksheet
Set ws = ActiveSheet
'apply the Excel SUMIF function
ws.Range("H2") = "2015"
ws.Range("I2") = Application.WorksheetFunction.SumIf(ws.Range("F2:F200"), "2015*", ws.Range("G2:G200"))
ws.Range("H3") = "2016"
ws.Range("I3") = Application.WorksheetFunction.SumIf(ws.Range("F2:F200"), "2016*", ws.Range("G2:G200"))
ws.Range("H4") = "2017"
ws.Range("I4") = Application.WorksheetFunction.SumIf(ws.Range("F2:F200"), "2017*", ws.Range("G2:G200"))
ws.Range("H5") = "2018"
ws.Range("I5") = Application.WorksheetFunction.SumIf(ws.Range("F2:F200"), "2018*", ws.Range("G2:G200"))
ws.Range("H6") = "2019"
ws.Range("I6") = Application.WorksheetFunction.SumIf(ws.Range("F2:F200"), "2019*", ws.Range("G2:G200"))

End Sub
