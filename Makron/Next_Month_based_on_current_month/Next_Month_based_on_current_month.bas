Attribute VB_Name = "Modul2"
Sub Next_Month_based_on_Current_Month()
'declare a variable
Dim ws As Worksheet
Set ws = ActiveSheet
'return the next month based on the current month
ws.Range("C1") = Format(DateAdd("m", 1, Date), "mmmm")
End Sub
