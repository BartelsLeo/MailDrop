Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.IO

Public Module MailUtils
    ' Liest die Metadaten der ausgewählten Mail und befüllt die Properties der übergebenen Session
    Public Sub ReadMailMeta(session As Session)
        Try
            Dim app As Outlook.Application = Globals.ThisAddIn.Application
            Dim explorer = app.ActiveExplorer()
            If explorer Is Nothing OrElse explorer.Selection.Count <> 1 Then
                System.Windows.MessageBox.Show("Eine Mail auswählen", "Warnung", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning)
                Return
            End If
            Dim mail = TryCast(explorer.Selection.Item(1), Outlook.MailItem)
            If mail Is Nothing Then
                System.Windows.MessageBox.Show("Bitte eine einzelne E-Mail auswählen.", "Warnung", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning)
                Return
            End If
            session.Absender = mail.SenderName
            If mail.SenderEmailType = "SMTP" AndAlso mail.SenderEmailAddress.Contains("@") Then
                session.AbsenderDomain = mail.SenderEmailAddress.Split("@"c).Last()
            End If
            session.Empfaenger = mail.To
            session.Betreff = mail.Subject
            session.Datum = mail.ReceivedTime
            session.DatumFormatiert = mail.ReceivedTime.ToString("yyyyMMdd")
        Catch ex As Exception
            System.Windows.MessageBox.Show($"Fehler beim Auslesen der E-Mail: {ex.Message}", "Fehler", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error)
        End Try
    End Sub

    ' Speichert die markierte Mail als .msg im Ablageordner
    Public Function SaveSelectedMailAsMsg(msgZielPfad As String) As String
        Try
            Dim app As Outlook.Application = Globals.ThisAddIn.Application
            Dim explorer = app.ActiveExplorer()
            Dim mail = TryCast(explorer.Selection.Item(1), Outlook.MailItem)
            If mail Is Nothing Then
                Return "Bitte wählen Sie eine einzelne E-Mail aus."
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

    ' Speichert alle Anhänge der Mail im Ablageordner
    Public Function SaveMailAttachments(anhangZielpfade As List(Of String)) As String
        Try
            Dim app As Outlook.Application = Globals.ThisAddIn.Application
            Dim explorer = app.ActiveExplorer()
            Dim mail = TryCast(explorer.Selection.Item(1), Outlook.MailItem)
            If mail Is Nothing Then
                Return "Bitte wählen Sie eine einzelne E-Mail aus."
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
            Return $"Fehler beim Speichern der Anhänge: {ex.Message}"
        End Try
    End Function

End Module
