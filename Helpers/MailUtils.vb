Imports System.IO

Public Module MailUtils
    ' Liest die Metadaten der ausgewählten Mail und befüllt die Properties der übergebenen Session
    Public Sub ReadMailMeta(session As Session)
        Dim app As Outlook.Application = Globals.ThisAddIn.Application
        Dim explorer = app.ActiveExplorer()
        If explorer Is Nothing OrElse explorer.Selection.Count <> 1 Then
            Throw New InvalidOperationException("Es muss genau eine E-Mail ausgewählt sein.")
        End If
        Dim mail = TryCast(explorer.Selection.Item(1), Outlook.MailItem)
        If mail Is Nothing Then
            Throw New InvalidOperationException("Das ausgewählte Element ist keine E-Mail.")
        End If
        session.Absender = mail.SenderName
        If mail.SenderEmailType = "SMTP" AndAlso mail.SenderEmailAddress.Contains("@") Then
            session.AbsenderDomain = mail.SenderEmailAddress.Split("@"c).Last()
        End If
        session.Empfaenger = mail.To
        session.Betreff = mail.Subject
        session.Datum = mail.ReceivedTime
        session.DatumFormatiert = mail.ReceivedTime.ToString("yyyyMMdd")
    End Sub
End Module
