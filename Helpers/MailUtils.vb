Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.IO

Public Module MailUtils
    ' Liest die Metadaten der ausgew�hlten Mail und bef�llt die Properties der �bergebenen Session
    Public Sub ReadMailMeta(session As Session)
        Dim app As Outlook.Application = Globals.ThisAddIn.Application
        Dim explorer = app.ActiveExplorer()
        If explorer Is Nothing OrElse explorer.Selection.Count <> 1 Then
            Throw New InvalidOperationException("Es muss genau eine E-Mail ausgew�hlt sein.")
        End If
        Dim mail = TryCast(explorer.Selection.Item(1), Outlook.MailItem)
        If mail Is Nothing Then
            Throw New InvalidOperationException("Das ausgew�hlte Element ist keine E-Mail.")
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

    ' Speichert die markierte Mail als .msg im Ablageordner
    Public Function SaveSelectedMailAsMsg(msgZielPfad As String) As String
        Try
            Dim app As Outlook.Application = Globals.ThisAddIn.Application
            Dim explorer = app.ActiveExplorer()
            Dim mail = TryCast(explorer.Selection.Item(1), Outlook.MailItem)
            If mail Is Nothing Then
                Return "Bitte w�hlen Sie eine einzelne E-Mail aus."
            End If
            Dim vollPfad = msgZielPfad
            If Not vollPfad.ToLower().EndsWith(".msg") Then
                vollPfad &= ".msg"
            End If
            mail.SaveAs(vollPfad, Outlook.OlSaveAsType.olMSG)
            Return String.Empty
        Catch ex As Exception
            Return $"Fehler beim Speichern der E-Mail: {ex.Message}"
        End Try
    End Function

    ' Speichert alle Anh�nge der Mail im Ablageordner
    Public Function SaveMailAttachments(anhangZielpfade As List(Of String)) As String
        Try
            Dim app As Outlook.Application = Globals.ThisAddIn.Application
            Dim explorer = app.ActiveExplorer()
            Dim mail = TryCast(explorer.Selection.Item(1), Outlook.MailItem)
            If mail Is Nothing Then
                Return "Bitte w�hlen Sie eine einzelne E-Mail aus."
            End If
            For i As Integer = 1 To mail.Attachments.Count
                If i <= anhangZielpfade.Count Then
                    Dim att = mail.Attachments(i)
                    Dim anhangPfad = anhangZielpfade(i - 1)
                    att.SaveAsFile(anhangPfad)
                End If
            Next
            Return String.Empty
        Catch ex As Exception
            Return $"Fehler beim Speichern der Anh�nge: {ex.Message}"
        End Try
    End Function

End Module
