Attribute VB_Name = "DateInHeader"
Sub DateInHeader()
With ActiveSheet.PageSetup
.LeftHeader = ""
.CenterHeader = "&D"
.RightHeader = ""
.LeftFooter = ""
.CenterFooter = ""
.RightFooter = ""
End With
ActiveWindow.View = xlNormalView
End Sub
