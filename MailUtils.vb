Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.IO

Public Module MailUtils
    ' Liest die Metadaten der ausgewählten Mail und gibt ein MailMetaInfo-Objekt zurück
    Public Function ReadMailMeta() As MailMetaInfo
        Try
            Dim app As Outlook.Application = Globals.ThisAddIn.Application
            Dim explorer = app.ActiveExplorer()
            If explorer Is Nothing OrElse explorer.Selection.Count <> 1 Then
                System.Windows.MessageBox.Show("Eine Mail auswählen", "Warnung", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning)
                Return Nothing
            End If
            Dim mail = TryCast(explorer.Selection.Item(1), Outlook.MailItem)
            If mail Is Nothing Then
                System.Windows.MessageBox.Show("Bitte eine einzelne E-Mail auswählen.", "Warnung", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning)
                Return Nothing
            End If
            Dim info As New MailMetaInfo()
            info.Sender = mail.SenderName
            If mail.SenderEmailType = "SMTP" AndAlso mail.SenderEmailAddress.Contains("@") Then
                info.SenderDomain = mail.SenderEmailAddress.Split("@"c).Last()
            End If
            info.Empfaenger = mail.To
            If Not String.IsNullOrEmpty(mail.To) Then
                Dim firstTo = mail.To.Split({";"}, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()
                If Not String.IsNullOrEmpty(firstTo) Then
                    info.EmpfaengerKurz = firstTo.Split("@"c)(0).Trim()
                End If
            End If
            info.Betreff = mail.Subject
            info.Datum = mail.ReceivedTime
            info.DatumFormatiert = mail.ReceivedTime.ToString("yyyyMMdd")
            Return info
        Catch ex As Exception
            System.Windows.MessageBox.Show($"Fehler beim Auslesen der E-Mail: {ex.Message}", "Fehler", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error)
            Return Nothing
        End Try
    End Function

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

    Public Class MailMetaInfo
        <DisplayName("Absender")>
        Public Property Sender As String
        <DisplayName("Absender-Domain")>
        Public Property SenderDomain As String
        <DisplayName("Empfänger")>
        Public Property Empfaenger As String
        <DisplayName("Empfänger (kurz)")>
        Public Property EmpfaengerKurz As String
        <DisplayName("Betreff")>
        Public Property Betreff As String
        <DisplayName("Datum")>
        Public Property Datum As DateTime
        <DisplayName("Datum (formatiert)")>
        Public Property DatumFormatiert As String
    End Class
End Module
