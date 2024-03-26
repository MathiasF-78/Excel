Attribute VB_Name = "InfogaHuvud"
Sub InfogaSidhuvud()
'Infogar text i dokumentethuvud
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
