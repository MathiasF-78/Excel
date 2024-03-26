Attribute VB_Name = "DatumSidhuvud"
Sub DatumSidhuvud()
'Infogar dagens datum i sidhuvudet på dokumentet (syns under Visa/Sidlayout)
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
