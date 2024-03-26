Attribute VB_Name = "copyPasteValue"
 Sub copyPasteValue()
   Range("O2").Copy
   Range("O1").PasteSpecial xlPasteValues
End Sub
   
