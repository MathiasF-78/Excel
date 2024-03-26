Attribute VB_Name = "AddHeader"
Sub InfogaSidhuvud()
'infogar sidhuvud i dokumentet med valfritt innehåll
Dim myText As String
myText = InputBox("Enter your text here", "Enter Text")
With ActiveSheet.PageSetup
.LeftHeader = ""
.CenterHeader = myText
.RightHeader = ""
.LeftFooter = ""
.CenterFooter = ""
.RightFooter = ""
End With
End Sub
