Attribute VB_Name = "Modul3"
Sub Delete_Text_From_Left(rng As Range, delChars As Byte)
  'For example: Deleting left 2 characters will turn CLASS into ASS.
  Dim myCell As Range
     On Error Resume Next
     For Each myCell In rng.Cells
        'Ignore if cell length < number of characters to be deleted.
        If Len(myCell.Value) > delChars Then
          'Take the rightmost characters and delete from left.
          myCell.Value = Right(myCell.Value, Len(myCell.Value) - delChars)
        End If
     Next
     Set myCell = Nothing
     On Error GoTo 0
End Sub
