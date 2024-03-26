Attribute VB_Name = "SortByYear"
Sub SortMultipleColumns()
With ActiveSheet
        With .Cells(1, "A").CurrentRegion
        'year i kolumn F Month i kolumn C
            .Cells.Sort Key1:=.Range("F2"), Order1:=xlAscending, _
             Key2:=.Range("C2"), Order2:=xlAscending, _
                        Orientation:=xlTopToBottom, Header:=xlYes
        End With
    End With

End Sub



