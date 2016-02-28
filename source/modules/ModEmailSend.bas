Option Compare Database

' INSERIRE LIBRERIA MICROSOFTcdo
' Ex. ?SendAMessage("mmorotti@scmgroup.com", "marco.morotti@gmail.com", "", "Oggetto Prova", "Corpo linea 1")

Public Function SendAMessage(strFrom As String, strTo As String, _
    strCC As String, strSubject As String, strTextBody As String, _
    Optional strBcc As String, Optional strAttachDoc As String)
Dim objMessage As CDO.Message

On Error GoTo MyErrorHadler

Set objMessage = New CDO.Message

With objMessage
    .From = strFrom
    .To = strTo
    If Len(Trim$(strCC)) > 0 Then
        .CC = strCC
    End If
    If Len(strBcc) > 0 Then
        .BCC = strBcc
    End If
    ''' On behalf of
    '.Sender = "Cheryl.Smith@abc.com"
    .Subject = strSubject
    .TextBody = strTextBody
    
    If Len(strAttachDoc) > 0 Then
        .AddAttachment strAttachDoc
    End If
    
    With .Configuration.Fields
        .item(CDO.cdoSMTPServer) = "ocsbh.scmgroup.com"
        .item(CDO.cdoSMTPServerPort) = 25
        .item(CDO.cdoSendUsingMethod) = CDO.cdoSendUsingPort
        .item(cdoSMTPConnectionTimeout) = 10
        .Update
    End With
    .Send
End With

Set objMessage = Nothing

Exit Function
MyErrorHadler:

End Function