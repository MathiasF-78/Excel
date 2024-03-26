Attribute VB_Name = "SammanfogaCeller"
Sub SammanfogaCeller()
Attribute SammanfogaCeller.VB_Description = "Sammanfoga markerade celler"
Attribute SammanfogaCeller.VB_ProcData.VB_Invoke_Func = "J\n14"
'sammanfogar markerade celler
    Dim cell As Range
    Dim CombinedEntry As String
 
'   Loop through selection and build string and erase values in cells being merged
    For Each cell In Selection
        CombinedEntry = CombinedEntry & cell.Value & " "
        cell.ClearContents
    Next cell
 
'   Remove last space from string
    CombinedEntry = Left(CombinedEntry, Len(CombinedEntry) - 1)
 
'   Merge cells and populate entry
    Selection.Merge
    ActiveCell.Value = CombinedEntry
 
End Sub
