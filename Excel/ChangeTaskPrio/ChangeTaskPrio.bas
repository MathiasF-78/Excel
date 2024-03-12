Attribute VB_Name = "ChangeTaskPrio"
Public Sub ChangeTaskPrio()
  Dim Selection As Outlook.Selection
  Dim obj As Object
  Dim Task As Outlook.TaskItem ' Task object used in For Each loop
  Dim i&
 Set Selection = Application.ActiveExplorer.Selection
  'If Selection.Count Then
    For Each obj In Selection
      'If TypeOf obj Is Outlook.TaskItem Then
      'Set Task = obj
'Set Task.Importance = olImportanceHigh
.Importance = olImportanceNormal
        'Task.Importance = olImportanceHigh
      Task.Save
       Next
  'End If
End Sub
