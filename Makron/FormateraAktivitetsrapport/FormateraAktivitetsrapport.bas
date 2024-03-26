Attribute VB_Name = "Aktivitetsrapport"
Sub Aktivitetsrapport()

 ' döp om blad
   ActiveSheet.Name = ActiveSheet.Range("J1")

 ' hide Mathias and Monica, rad 8 och 9
Rows("8:9").EntireRow.Hidden = True

' formatera rubriker
    Range("B1").Select
    Selection.Cut
    Range("A2").Select
    ActiveSheet.Paste
    
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Range("J1").Select
    ActiveCell.Formula = "Totalt"
    Range("K1").Select
    ActiveCell.Formula = "%"

 Range("A2:J7").Select
  ActiveSheet.sort.SortFields.Clear
  ActiveSheet.sort.SortFields.Add2 Key:=Range("J4:J11") _
  , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    
    With ActiveSheet.sort
        .SetRange Range("A2:J9")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

' ruta in område
   
    Range("A1:J10").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
  End With
  
' Fet stil och design
    
    Range("1:1,A:A").Select
    Selection.Font.Bold = True
    Range("A9:J9").Select
    Range("J9").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A10:J10").Select
    Selection.Font.Bold = False
    Selection.Font.Bold = True
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A10:J10").Select
    Selection.Font.Bold = False
    Selection.Font.Bold = True

    Cells.EntireColumn.AutoFit

' Fat Line

    Range("A9:J9").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

' summera rader

    Range("B10").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-8]C,R[-7]C,R[-6]C,R[-5]C,R[-2]C,R[-1]C)"
    Range("B10").Select
    Selection.Copy
    Range("C10").Select
    ActiveSheet.Paste
    Range("D10").Select
    ActiveSheet.Paste
    Range("E10").Select
    ActiveSheet.Paste
    Range("F10").Select
    ActiveSheet.Paste
    Range("G10").Select
    ActiveSheet.Paste
    Range("H10").Select
    ActiveSheet.Paste
    Range("I10").Select
    ActiveSheet.Paste
    Range("J10").Select
    ActiveSheet.Paste
    Range("K10").Select
    Application.CutCopyMode = False
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Fet"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveCell.FormulaR1C1 = "=SUM(100/RC[-1])"
    Range("K2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]*R[8]C"
    Range("K3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]*R[7]C"
    Range("K4").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]*R[6]C"
    Range("K5").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]*R[5]C"
    Range("K8").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]*R[2]C"
    Range("K9").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]*R[1]C"
    Range("K2:K10").Select
    Selection.Style = "Percent"
    Range("N10").Select


' skapa procent kolumn

    Range("K2").Select
    ActiveCell.FormulaR1C1 = "=(RC[-1]/R[8]C[-1])"
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "=(RC[-1]/R[7]C[-1])"
    Range("K4").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/R[6]C[-1]"
    Range("K5").Select
    ActiveCell.FormulaR1C1 = "=(RC[-1]/R[5]C[-1])"
    Range("K8").Select
    ActiveCell.FormulaR1C1 = "=(RC[-1]/R[2]C[-1])"
    Range("K9").Select
    ActiveCell.FormulaR1C1 = "=(RC[-1]/R[1]C[-1])"
    Range("K2:K9").Select
    Selection.NumberFormat = "0.0%"
    Range("K1:K10").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("K9").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("K10").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("K10").Select
    ActiveCell.FormulaR1C1 = "100%"
    Range("K11").Select

End Sub
