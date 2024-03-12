Attribute VB_Name = "Modul9"
Sub everything()
'------------------------SETUP------------------------
  
'------------------------Rename ActiveSheet------------------------
   
 '  ActiveSheet.Name = ActiveSheet.Range("B1")

'------------------------Hide specific colleagues------------------------
    'Rows("8:9").EntireRow.Hidden = True

'------------------------Formatera rubriker------------------------
' formatera rubriker
    
    
    Range("B1").Select
    Selection.Cut
    Range("A2").Select
    ActiveSheet.Paste
    
    'lägg till rubriker
    Range("G2").Select
    ActiveCell.Formula = "Totalt"
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Range("H1").Select
    ActiveCell.Formula = "%"

'------------------------Kontrollera och lägg till eventuella rubriker------------------------

    CheckColumnHeadingIsPresent "BFUS notifikation", 2
    CheckColumnHeadingIsPresent "Checklistkommentar", 3
    CheckColumnHeadingIsPresent "Kommentar", 4
    CheckColumnHeadingIsPresent "Sökt, ej på plats", 5
    CheckColumnHeadingIsPresent "Fått e-post", 6
    CheckColumnHeadingIsPresent "Säljsamtal", 7
    CheckColumnHeadingIsPresent "Skickat e-post", 8
    CheckColumnHeadingIsPresent "Pratat med", 9
    CheckColumnHeadingIsPresent "Totalt", 10
    CheckColumnHeadingIsPresent "%", 11

End Sub


Sub CheckColumnHeadingIsPresent(strColumnHeading As String, intPosition As Integer)

Dim wsData As Worksheet
If ActiveSheet.Cells(1, intPosition) <> strColumnHeading Then
 
 ' The column heading is missing so insert it at intPosition.
 ActiveSheet.Columns(intPosition).Select
 Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
 ActiveSheet.Cells(1, intPosition) = strColumnHeading
End If

'------------------------Summering och procent - setup------------------------

Dim cl As Range
Set cl = Range("A:A").Find("Totalsumma")    ' hitta raden totalsumma
If Not cl Is Nothing Then cl.Select

With Selection
    .Offset(, 1).Formula = "=SUM(B2:B8)"
    .Offset(, 2).Formula = "=SUM(C2:C8)"
    .Offset(, 3).Formula = "=SUM(D2:D8)"
    .Offset(, 4).Formula = "=SUM(E2:E8)"
    .Offset(, 5).Formula = "=SUM(F2:F8)"
    .Offset(, 6).Formula = "=SUM(G2:G8)"
    .Offset(, 7).Formula = "=SUM(H2:H8)"
    End With
    
Set cl = Range("A:A").Find("Totalsumma")    ' hitta raden totalsumma
If Not cl Is Nothing Then cl.Select

n = Range("J9")
    
' benämner alla variabler
      number_1 = Range("J2").Value
      number_2 = Range("J3").Value
      number_3 = Range("J4").Value
      number_4 = Range("J5").Value
      number_5 = Range("J6").Value
      number_6 = Range("J7").Value
      number_7 = Range("J8").Value
      number_8 = Range("J9").Value

 Range("K2").Formula = "=SUM(J2/J9)"
 Range("K3").Formula = "=SUM(J3/J9)"
  Range("K4").Formula = "=SUM(J4/J9)"
   Range("K5").Formula = "=SUM(J5/J9)"
    Range("K6").Formula = "=SUM(J6/J9)"
     Range("K7").Formula = "=SUM(J7/J9)"
      Range("K8").Formula = "=SUM(J8/J9)"
       Range("K9").Formula = "=SUM(J9/J9)"
       
 Range("K2:K9").NumberFormat = "0.00%"
 
End Sub
