Attribute VB_Name = "CustomHeader"
Sub CustomHeader()
Dim myText As Stringmy
Text = InputBox("Enter your text here", "Enter Text")
With ActiveSheet.PageSetup
.LeftHeader = ""
.CenterHeader = myText
.RightHeader = ""
.LeftFooter = ""
.CenterFooter = ""
.RightFooter = "" End Sub
