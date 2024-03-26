Attribute VB_Name = "RaderaDubletter"
Sub RemoveDupsEx1()
    Dim Selection As Range
    
    Range("A2:P6000").RemoveDuplicates Columns:=1, Header:=xlYes 'defined range
    'ActiveSheet.UsedRange.RemoveDuplicates Columns:=1, Header:=xlYes 'used range
    'ActiveSheet.Selection.RemoveDuplicates Columns:=1, Header:=xlYes 'selection

End Sub
