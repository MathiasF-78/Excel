Attribute VB_Name = "SendAsMail"
Sub SendAsMail()
' Kopiera hela modulen.
' Konstaterat fungerar i Office 2016
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Dim SigString As String
    Dim Signature As String

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

'ändra "Hej," och "Se bifogat dokument." till valfritt.
'du kan lägag till mer text mellan " och <br> på rad 2 och 3 samt lägga till felra genom att koiera och klistra in dessa
    strbody = "<BODY style=font-size:10pt;font-family:Calibri>Hej," & _
              "<br>" & _
              "<br>" & _
              "Se bifogat dokument.</A>" & _
              "<br><br></BODY>"
    'Ändra  Mathias.htm till namnet på din signatur i outlook, oftast "Standard.htm"
    SigString = Environ("appdata") & _
                "\Microsoft\Signaturer\Mathias.htm"

    If Dir(SigString) <> "" Then
        Signature = GetBoiler(SigString)
    Else
        Signature = ""
    End If

    On Error Resume Next

    With OutMail
        .To = ""
        ' du kan skriva in ett kontaktnamn kontakt eller grupp från ditt outlook kontaktregister om du alltid mailar samma mottagare varje gång
                .CC = ""
        .BCC = "" ' samma här
        .Subject = "bifogat dokument"
        ' ändra detta till valfri ämnesrad på mailet
        .HTMLBody = strbody & "<br>" & Signature
        .Attachments.Add ActiveWorkbook.FullName
        .display 'eller använd .Send om du vill skicka mailet direkt utan förhandsvisning
       ' .Send    'eller använd .display om du vill förhandsvisa mailet innan du skickar det
    End With

    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub


Function GetBoiler(ByVal sFile As String) As String
'Dick Kusleika
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)
    GetBoiler = ts.readall
    ts.Close
End Function



