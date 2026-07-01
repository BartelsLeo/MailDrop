Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Diagnostics

Public Module MailUtils
    Private Sub ReleaseComObjectSafe(comObject As Object)
        If comObject Is Nothing Then Return
        If Marshal.IsComObject(comObject) Then
            Marshal.FinalReleaseComObject(comObject)
        End If
    End Sub

    ' Liest die Metadaten der ausgew�hlten Mail und bef�llt die Properties der �bergebenen Session
    Public Sub ReadMailMeta(session As Session)
        Dim mail As Object = Nothing
        Debug.WriteLine("[MailUtils] ReadMailMeta called.")
        Try
            Dim app As Outlook.Application = Globals.ThisAddIn.Application
            Dim explorer As Object = app.ActiveExplorer()
            If explorer Is Nothing OrElse explorer.Selection.Count <> 1 Then
                Debug.WriteLine($"[MailUtils] ReadMailMeta: selection count={If(explorer Is Nothing, "no explorer", explorer.Selection.Count.ToString())} - skipping.")
                Return
            End If
            mail = TryCast(explorer.Selection.Item(1), Outlook.MailItem)
            If mail Is Nothing Then
                Debug.WriteLine("[MailUtils] ReadMailMeta: selected item is not a MailItem - skipping.")
                Return
            End If
            session.Absender = mail.SenderName
            If mail.SenderEmailType = "SMTP" AndAlso mail.SenderEmailAddress.Contains("@") Then
                Dim emailParts = mail.SenderEmailAddress.Split("@"c)
                session.AbsenderDomain = emailParts(emailParts.Length - 1)
            End If
            session.Empfaenger = mail.To
            session.Betreff = mail.Subject
            Dim receivedTime As DateTime = CType(mail.ReceivedTime, DateTime)
            session.Datum = receivedTime
            session.DatumFormatiert = receivedTime.ToString("yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture)
            Debug.WriteLine($"[MailUtils] ReadMailMeta written: Betreff='{session.Betreff}', Absender='{session.Absender}', AbsenderDomain='{session.AbsenderDomain}', Empfaenger='{session.Empfaenger}', Datum='{session.Datum:yyyy-MM-dd HH:mm:ss}', DatumFormatiert='{session.DatumFormatiert}', AusfueBenutzer='{session.AusfueBenutzer}', AusfueDatum='{session.AusfueDatum:yyyy-MM-dd HH:mm:ss}'")
        Catch ex As Exception
            Debug.WriteLine($"[MailUtils] ReadMailMeta exception: {ex.Message}")
        Finally
            ReleaseComObjectSafe(mail)
            ' Explorer is intentionally NOT released: app.ActiveExplorer() returns the same RCW
            ' as _currentExplorer in ThisAddIn. FinalReleaseComObject on it would destroy the
            ' SelectionChange event connection permanently.
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

    ' Liest Anhang-Namen und -Indizes aus der selektierten Mail und befüllt session.Anhaenge
    Public Sub ReadAttachmentNames(session As Session)
        Dim mail As Object = Nothing
        Try
            Dim app As Outlook.Application = Globals.ThisAddIn.Application
            Dim explorer As Object = app.ActiveExplorer()
            If explorer Is Nothing OrElse explorer.Selection.Count <> 1 Then Return
            mail = TryCast(explorer.Selection.Item(1), Outlook.MailItem)
            If mail Is Nothing Then Return
            session.Anhaenge.Clear()
            For i As Integer = 1 To mail.Attachments.Count
                Dim att As Object = Nothing
                Try
                    att = mail.Attachments(i)
                    session.Anhaenge.Add(New AttachmentItem() With {
                        .Name = att.FileName,
                        .OutlookIndex = i,
                        .IsSelected = True
                    })
                Finally
                    ReleaseComObjectSafe(att)
                End Try
            Next
        Catch ex As Exception
            Debug.WriteLine($"[MailUtils] ReadAttachmentNames exception: {ex.Message}")
        Finally
            ReleaseComObjectSafe(mail)
        End Try
    End Sub

    ' Speichert die selektierten Anhänge der Mail; Zuordnung per Dateiname (nicht per Index)
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
            For Each anhangPfad In anhangZielpfade
                Dim targetName = Path.GetFileName(anhangPfad)
                For i As Integer = 1 To mail.Attachments.Count
                    Dim att As Object = Nothing
                    Try
                        att = mail.Attachments(i)
                        If String.Equals(att.FileName, targetName, StringComparison.OrdinalIgnoreCase) Then
                            att.SaveAsFile(anhangPfad)
                            Exit For
                        End If
                    Finally
                        ReleaseComObjectSafe(att)
                    End Try
                Next
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
