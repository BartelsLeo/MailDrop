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
            Dim vollPfad = If(msgZielPfad.ToLower().EndsWith(".msg"), msgZielPfad, msgZielPfad & ".msg")

            ' Datei existiert bereits: Überschreiben / Umbenennen / Abbrechen
            If File.Exists(vollPfad) Then
                Dim choice = AskOverwriteRenameCancel(Path.GetFileName(vollPfad))
                Select Case choice
                    Case "cancel"
                        Return "Vorgang abgebrochen."
                    Case "rename"
                        Dim newName = AskNewName(Path.GetFileName(vollPfad), Path.GetDirectoryName(vollPfad),
                                                 "Bitte geben Sie einen neuen Dateinamen ein:")
                        If String.IsNullOrEmpty(newName) Then Return "Vorgang abgebrochen."
                        vollPfad = Path.Combine(Path.GetDirectoryName(vollPfad), newName)
                    ' "overwrite": direkt fortfahren
                End Select
            End If

            ' Speichern; bei gesperrter Datei: Umbenennen / Abbrechen
            Try
                mail.SaveAs(vollPfad, Outlook.OlSaveAsType.olMSG)
            Catch ex As Exception
                Dim newName = AskRenameOrCancel(Path.GetFileName(vollPfad), ex.Message,
                                                Path.GetDirectoryName(vollPfad))
                If String.IsNullOrEmpty(newName) Then Return "Vorgang abgebrochen."
                vollPfad = Path.Combine(Path.GetDirectoryName(vollPfad), newName)
                mail.SaveAs(vollPfad, Outlook.OlSaveAsType.olMSG)
            End Try

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

                    ' Datei existiert bereits: Überschreiben / Umbenennen / Abbrechen
                    If File.Exists(anhangPfad) Then
                        Dim choice = AskOverwriteRenameCancel(Path.GetFileName(anhangPfad))
                        Select Case choice
                            Case "cancel"
                                Return "Vorgang abgebrochen."
                            Case "rename"
                                Dim newName = AskNewName(Path.GetFileName(anhangPfad), Path.GetDirectoryName(anhangPfad),
                                                         "Bitte geben Sie einen neuen Dateinamen ein:")
                                If String.IsNullOrEmpty(newName) Then Return "Vorgang abgebrochen."
                                anhangPfad = Path.Combine(Path.GetDirectoryName(anhangPfad), newName)
                            ' "overwrite": direkt fortfahren
                        End Select
                    End If

                    ' Speichern; bei gesperrter Datei: Umbenennen / Abbrechen
                    Try
                        att.SaveAsFile(anhangPfad)
                    Catch ex As Exception
                        Dim newName = AskRenameOrCancel(Path.GetFileName(anhangPfad), ex.Message,
                                                        Path.GetDirectoryName(anhangPfad))
                        If String.IsNullOrEmpty(newName) Then Return "Vorgang abgebrochen."
                        anhangPfad = Path.Combine(Path.GetDirectoryName(anhangPfad), newName)
                        att.SaveAsFile(anhangPfad)
                    End Try
                End If
            Next
            Return String.Empty
        Catch ex As Exception
            Return $"Fehler beim Speichern der Anh�nge: {ex.Message}"
        End Try
    End Function

    ' Zeigt "Überschreiben / Umbenennen / Abbrechen" — gibt "overwrite", "rename" oder "cancel" zurück
    Private Function AskOverwriteRenameCancel(fileName As String) As String
        Dim result = System.Windows.MessageBox.Show(
            $"Die Datei ""{fileName}"" existiert bereits.{vbCrLf}" &
            "Ja = �berschreiben     Nein = Umbenennen     Abbrechen = Vorgang abbrechen",
            "Datei existiert bereits",
            System.Windows.MessageBoxButton.YesNoCancel,
            System.Windows.MessageBoxImage.Question)
        Select Case result
            Case System.Windows.MessageBoxResult.Yes : Return "overwrite"
            Case System.Windows.MessageBoxResult.No : Return "rename"
            Case Else : Return "cancel"
        End Select
    End Function

    ' Zeigt bei gesperrter/nicht speicherbarer Datei "Umbenennen / Abbrechen", öffnet bei Ja den Rename-Dialog
    ' Gibt den neuen Namen zurück oder String.Empty wenn abgebrochen
    Private Function AskRenameOrCancel(fileName As String, errorMsg As String, basePath As String) As String
        Dim result = System.Windows.MessageBox.Show(
            $"Die Datei ""{fileName}"" konnte nicht gespeichert werden:{vbCrLf}{errorMsg}{vbCrLf}{vbCrLf}" &
            "Soll ein anderer Dateiname vergeben werden?",
            "Datei nicht speicherbar",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Warning)
        If result <> System.Windows.MessageBoxResult.Yes Then Return String.Empty
        Return AskNewName(fileName, basePath, "Bitte geben Sie einen neuen Dateinamen ein:")
    End Function

    ' Öffnet den Rename-Dialog und gibt den neuen Namen oder String.Empty zurück
    Private Function AskNewName(currentName As String, basePath As String, hintText As String) As String
        Return InputChecker.ShowAttachmentRenameDialog(currentName, basePath, hintText, "Datei umbenennen")
    End Function

End Module
