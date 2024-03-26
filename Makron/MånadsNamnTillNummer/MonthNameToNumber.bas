Attribute VB_Name = "MonthNameToNumber"
Sub columnMonthNumber()
'kopierar kolumn C och konverterar ny kolumn till månadens nummer
  Range("C:C").Copy 'kopierar kolumn C
  Range("C:C").Insert 'klistrar in kolumnen till vänster om kolumn C
     Columns("C:C").Select
       Range("C:C").Replace What:="jan", Replacement:="1"
       Range("C:C").Replace What:="feb", Replacement:="2"
       Range("C:C").Replace What:="mar", Replacement:="3"
       Range("C:C").Replace What:="apr", Replacement:="4"
       Range("C:C").Replace What:="maj", Replacement:="5"
       Range("C:C").Replace What:="jun", Replacement:="6"
       Range("C:C").Replace What:="jul", Replacement:="7"
       Range("C:C").Replace What:="aug", Replacement:="8"
       Range("C:C").Replace What:="sep", Replacement:="9"
       Range("C:C").Replace What:="okt", Replacement:="10"
       Range("C:C").Replace What:="nov", Replacement:="11"
       Range("C:C").Replace What:="dec", Replacement:="12"
End Sub

