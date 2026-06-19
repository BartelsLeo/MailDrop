Imports System.IO
Imports System.Runtime.InteropServices

Public Module MailUtils
    Private Sub ReleaseComObjectSafe(comObject As Object)
        If comObject Is Nothing Then Return
        If Marshal.IsComObject(comObject) Then
            Marshal.FinalReleaseComObject(comObject)
        End If
    End Sub

    ' Liest die Metadaten der ausgew�hlten Mail und bef�llt die Properties der �bergebenen Session
    Public Sub ReadMailMeta(session As Session)
        Dim explorer As Object = Nothing
        Dim mail As Object = Nothing
        Try
            Dim app As Outlook.Application = Globals.ThisAddIn.Application
            explorer = app.ActiveExplorer()
            If explorer Is Nothing OrElse explorer.Selection.Count <> 1 Then
                System.Windows.MessageBox.Show("Eine Mail ausw�hlen", "Warnung", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning)
                Return
            End If
            mail = TryCast(explorer.Selection.Item(1), Outlook.MailItem)
            If mail Is Nothing Then
                System.Windows.MessageBox.Show("Bitte eine einzelne E-Mail ausw�hlen.", "Warnung", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning)
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
        Finally
            ReleaseComObjectSafe(mail)
            ReleaseComObjectSafe(explorer)
        End Try
    End Sub

    ' Speichert die markierte Mail als .msg im Ablageordner
    Public Function SaveSelectedMailAsMsg(msgZielPfad As String) As String
        Dim explorer As Object = Nothing
        Dim mail As Object = Nothing
        Try
            Dim app As Outlook.Application = Globals.ThisAddIn.Application
            explorer = app.ActiveExplorer()
            If explorer Is Nothing OrElse explorer.Selection.Count < 1 Then
                Return "Bitte wählen Sie eine einzelne E-Mail aus."
            End If
            mail = TryCast(explorer.Selection.Item(1), Outlook.MailItem)
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
        Finally
            ReleaseComObjectSafe(mail)
            ReleaseComObjectSafe(explorer)
        End Try
    End Function

    ' Speichert alle Anh�nge der Mail im Ablageordner
    Public Function SaveMailAttachments(anhangZielpfade As List(Of String)) As String
        Dim explorer As Object = Nothing
        Dim mail As Object = Nothing
        Try
            Dim app As Outlook.Application = Globals.ThisAddIn.Application
            explorer = app.ActiveExplorer()
            If explorer Is Nothing OrElse explorer.Selection.Count < 1 Then
                Return "Bitte w�hlen Sie eine einzelne E-Mail aus."
            End If
            mail = TryCast(explorer.Selection.Item(1), Outlook.MailItem)
            If mail Is Nothing Then
                Return "Bitte w�hlen Sie eine einzelne E-Mail aus."
            End If
            For i As Integer = 1 To mail.Attachments.Count
                If i <= anhangZielpfade.Count Then
                    Dim att As Object = Nothing
                    Try
                        att = mail.Attachments(i)
                        Dim anhangPfad = anhangZielpfade(i - 1)
                        att.SaveAsFile(anhangPfad)
                    Finally
                        ReleaseComObjectSafe(att)
                    End Try
                End If
            Next
            Return String.Empty
        Catch ex As Exception
            Return $"Fehler beim Speichern der Anh�nge: {ex.Message}"
        Finally
            ReleaseComObjectSafe(mail)
            ReleaseComObjectSafe(explorer)
        End Try
    End Function

End Module
