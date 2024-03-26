Attribute VB_Name = "ToppTio"
Sub ToppTio()
'markerar de 10 högsta värdena inom markeringen

    Selection.FormatConditions.AddTop10
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
        With Selection.FormatConditions(1)
            .TopBottom = xlTop10Top
            'Change the rank here to highlight a different number of values
            .Rank = 10
            .Percent = False
        End With
        With Selection.FormatConditions(1).Font
            .Color = -16752384
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .ColorIndex = 36
'.Color = 13561798
            .TintAndShade = 0
        End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub
