Attribute VB_Name = "savetosharepoint"
Sub savetosharepoint()
Attribute savetosharepoint.VB_ProcData.VB_Invoke_Func = " \n14"

    ActiveWorkbook.SaveAs Filename:="https://neabonline.sharepoint.com/sites/Lime-testmapp/Shared Documents/elpriser(demo)/mfn-SVK-tim-mall.xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
End Sub
Sub Makro2()
Attribute Makro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makro2 Makro
'

'
    ActiveWorkbook.Save
    ActiveWorkbook.SaveAs Filename:= _
        "https://neabonline.sharepoint.com/sites/Lime-testmapp/Delade%20dokument/elpriser(demo)/mfn-SVK-tim-mall.xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
End Sub
Sub Makro3()
Attribute Makro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makro3 Makro
'

'
    ActiveWorkbook.Save
    ActiveWorkbook.Save
    Range("H22").Select
End Sub
Sub Makro4()
Attribute Makro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makro4 Makro
'

'
    ActiveWorkbook.Save
    ActiveWorkbook.SaveAs Filename:= _
        "https://neabonline.sharepoint.com/sites/FrsljningKundcenter/Delade%20dokument/elpriser/aktuella-elpriser-.xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
End Sub
