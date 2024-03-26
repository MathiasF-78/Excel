Attribute VB_Name = "MarkeraUnik"
Sub MarkeraUnik()
'Markerar unika värden inom markering (som villkorsstyrd formatering)
Dim Rng As Range
Set Rng = Selection
Rng.FormatConditions.Delete
Dim uv As UniqueValues
Set uv = Rng.FormatConditions.AddUniqueValues
uv.DupeUnique = xlUnique
uv.Interior.ColorIndex = 36
'= vbGreen
End Sub
