Public Class Redirect

    Public Shared Sub AutoReply(ByVal redirectAddress As String, ByRef Original As Outlook.MailItem, Optional ByVal dbg As Boolean = False)

        If dbg = True Then

            MsgBox("Debugging mode active!" & vbCr & "Message will be generated but not sent.", vbOKOnly)

        End If

        ' Set reply to the reply to the original item
        ' NB: I was going to use ReplyAll here instead, but I see no point in unnecessarily address shaming the email sender in front of the CCs.
        Dim Reply As Outlook.MailItem
        'Dim Original As Outlook.MailItem

        Reply = Original.Reply
        ' Throw the original email on as an attachment
        Reply.Attachments.Add(Original)

        ' Let's add that to the email reply
        ' as a CC because why not
        Dim infoRecip = Reply.Recipients.Add(redirectAddress)
        infoRecip.Type = Outlook.OlMailRecipientType.olCC
        ' and resolve it, so it looks nice
        ' Probably don't really need this, but I'm keeping it
        Reply.Recipients.ResolveAll()

        ' Add subject line to reply email
        Reply.Subject = "RE: " & Original.Subject
        ' Let's keep a copy, shall we?
        Reply.DeleteAfterSubmit = False

        ' Put some body in the reply!
        Reply.BodyFormat = Original.BodyFormat
        Reply.Body = ""
        Reply.Body = Reply.Body & "THIS IS AN AUTOMATED RESPONSE:" & vbCrLf
        Reply.Body = Reply.Body & "==============================" & vbCrLf & vbCrLf
        Reply.Body = Reply.Body & "In future, please send all emails of this nature to " & redirectAddress & "." & vbCrLf
        Reply.Body = Reply.Body & "This email has been copied to that address to be actioned." & vbCrLf & vbCrLf
        Reply.Body = Reply.Body & "==============================" & vbCrLf & vbCrLf & vbCrLf
        Reply.Body = Reply.Body & Original.Body

        ' Show it while Debugging
        If dbg = True Then
            Reply.Display()
        Else
            ' Send it when not debugging
            Reply.Send()
            ' Close the original, discarding any changes
            Original.Close(Outlook.OlInspectorClose.olDiscard)
        End If

    End Sub

End Class
