Attribute VB_Name = "percentageChange"
Option Explicit

Sub Percent_Change_Formula()
'Description: Creates a percentage change formula
'Source: https://www.excelcampus.com/vba/percentage-change-formulas-macro/

Dim rOld As Range
Dim rNew As Range
Dim sOld As String
Dim sNew As String
Dim sFormula As String
    
    'End the macro on any input errors
    'or if the user hits Cancel in the InputBox
    On Error GoTo ErrExit
    
    'Prompt the user to select the cells
    Set rNew = Application.InputBox( _
            "Markera det NYA värdet", _
            "Markera NY cell", Type:=8)
    Set rOld = Application.InputBox( _
            "Markera deet GAMLA värdet", _
            "Markera GAMMAL cell", Type:=8)
    
    'Get the cell addresses for the formula
    sNew = rNew.Address(False, False)
    sOld = rOld.Address(False, False)
        
    'Create the formula
    sFormula = "=IFERROR((" & sNew & " - " & sOld & ")/" & sOld & ",0)"

    'Simplified formula - uncomment line below to use this version
    'sFormula = "=IFERROR((" & sNew & "/" & sOld & ")-1,0)"


    'Create the formula in the activecell
    ActiveCell.Formula = sFormula
    ActiveCell.NumberFormat = "0.00%"
    
    
ErrExit:
    
End Sub
