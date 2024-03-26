Attribute VB_Name = "InfogaDatumTid"
Sub InfogaDatumTid()
'Infogar dagens datum och tid i markerad cell
ActiveCell = Format(Now, "yyyy-mm-dd HH:mm:ss")
End Sub
