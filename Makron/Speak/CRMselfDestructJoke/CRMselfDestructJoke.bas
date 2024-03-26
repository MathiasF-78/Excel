Attribute VB_Name = "CRMselfDestructJoke"
Sub Speak()
    Dim Message As String
    Dim Speak As Object
    
    ' Prompt user for input
    'Message = InputBox("Enter text", "Speak")
    Message = "Just kidding!"
    
    ' Check if the user entered text
    If Len(Trim(Message)) > 0 Then
        ' Flag to indicate if it's the first time speaking
        Dim isFirstSpeak As Boolean
        isFirstSpeak = True

        ' Display countdown from 10 seconds to 1
        'For i = 10 To 1 Step -1
        ' Display countdown from 10 seconds to 8
        For i = 10 To 8 Step -1
            Application.Wait Now + TimeValue("00:00:01")
            If isFirstSpeak Then
                SpeakText "This CRM system will self destruct in " & i & " seconds"
                isFirstSpeak = False
            Else
                SpeakText CStr(i) ' Convert i to string
            End If
        Next i

        ' Speak the entered text
        Set Speak = CreateObject("sapi.spvoice")
        Speak.Speak Message
    Else
        ' Inform the user if no text was entered
        MsgBox "No text entered. Exiting."
    End If
End Sub

Sub SpeakText(TextToSpeak As Variant)
    Dim Speak As Object
    Set Speak = CreateObject("sapi.spvoice")
    Speak.Speak CStr(TextToSpeak) ' Convert to string if not already
End Sub

