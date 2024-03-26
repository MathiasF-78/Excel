Attribute VB_Name = "OrgNrFormattingBFUS"
Sub BFUS_orgnr()
'formaterar organisationsnumme, raderar 16 och 00. 16xxxxxxxxxx00 -> xxxxxxxxxx

Dim c As Range
For Each c In Range("A1:A" & Range("A" & Rows.Count).End(xlUp).Row)
    If Left(c, 2) Like "16" Then c = Mid(c, 3)
Next

For Each c In Range("A1:A" & Range("A" & Rows.Count).End(xlUp).Row)
    If Right(c, 2) Like "00" Then c = Mid(c, 3)
Next
End Sub

