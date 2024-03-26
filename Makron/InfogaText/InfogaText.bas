Attribute VB_Name = "InfogaText"
Sub InfogaText()
'Lägger till valfri text till höger i markerade celler
' Declare variables
Dim Msg, Title, Default, userText, c

' Define text Input dialog parameters...
Msg = "Enter text to append in selected cells: "
Title = "Append Text"
Default = ""

' Prompt user to enter text to append to cells...
userText = InputBox(Msg, Title, Default)

' Append user text to selected cells...
For Each c In Selection.Cells
c.Value = c.Value & userText
Next c

End Sub
