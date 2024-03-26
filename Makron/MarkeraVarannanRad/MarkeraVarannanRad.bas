Attribute VB_Name = "MarkeraVarannanRad"
Sub MarkeraVarannanRad()
'Markerar varannan rad i markeringen med gul
Dim myRange As Range
Dim Myrow As Range
Set myRange = Selection
For Each Myrow In myRange.Rows
   If Myrow.Row Mod 2 = 1 Then
      Myrow.Interior.ColorIndex = 36
      '= RGB(230, 230, 230)
   End If
Next Myrow
End Sub
