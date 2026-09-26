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

    ' Liefert die abzulegende Mail: aus dem geoeffneten Mail-Fenster (session.SourceInspector),
    ' sonst die einzelne Auswahl im Explorer. Nothing, wenn keine einzelne MailItem verfuegbar ist.
    ' Der Explorer/Inspector selbst wird NIE freigegeben (siehe Kommentar in ReadMailMeta) - nur
    ' das zurueckgegebene MailItem gibt der Aufrufer im Finally frei.
    Private Function GetSourceMail(session As Session) As Outlook.MailItem
        Dim inspector As Outlook.Inspector = session?.SourceInspector
        If inspector IsNot Nothing Then
            Return TryCast(inspector.CurrentItem, Outlook.MailItem)
        End If
        Dim explorer As Outlook.Explorer = Globals.ThisAddIn.Application.ActiveExplorer()
        If explorer Is Nothing OrElse explorer.Selection.Count <> 1 Then Return Nothing
        Return TryCast(explorer.Selection.Item(1), Outlook.MailItem)
    End Function

    ' Prueft, ob noch dieselbe Mail aktiv ist, fuer die PrepareSession die Felder befuellt hat.
    Private Function IsSameMailAsPrepared(session As Session, mail As Outlook.MailItem) As Boolean
        If String.IsNullOrEmpty(session.SourceMailEntryId) Then Return True
        Return String.Equals(session.SourceMailEntryId, mail.EntryID, StringComparison.Ordinal)
    End Function

    ' Liest die Metadaten der ausgew�hlten Mail und bef�llt die Properties der �bergebenen Session
    Public Sub ReadMailMeta(session As Session)
        Dim mail As Object = Nothing
        Debug.WriteLine("[MailUtils] ReadMailMeta called.")
        Try
            mail = GetSourceMail(session)
            If mail Is Nothing Then
                Debug.WriteLine("[MailUtils] ReadMailMeta: no single MailItem available - skipping.")
                Return
            End If
            session.SourceMailEntryId = mail.EntryID
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
            ' SelectionChange event connection permanently. Same for session.SourceInspector: its
            ' RCW is held by ThisAddIn for the Inspector.Close handler of that window's task pane.
        End Try
    End Sub

    ' Speichert die markierte Mail als .msg im Ablageordner
    Public Function SaveSelectedMailAsMsg(session As Session, msgZielPfad As String) As String
        Dim mail As Object = Nothing
        Try
            mail = GetSourceMail(session)
            If mail Is Nothing Then
                Return "Bitte wählen Sie eine einzelne E-Mail aus."
            End If
            If Not IsSameMailAsPrepared(session, mail) Then
                Return "Die angezeigte Mail hat sich geändert. Bitte MailDrop erneut öffnen."
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
            ' Explorer is intentionally NOT released: app.ActiveExplorer() returns the same RCW
            ' as _currentExplorer in ThisAddIn. FinalReleaseComObject on it would destroy the
            ' SelectionChange event connection permanently (see ReadMailMeta).
        End Try
    End Function

    ' Liest Anhang-Namen und -Indizes aus der selektierten Mail und befüllt session.Anhaenge
    Public Sub ReadAttachmentNames(session As Session)
        Dim mail As Object = Nothing
        Try
            mail = GetSourceMail(session)
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
    Public Function SaveMailAttachments(session As Session, anhangZielpfade As List(Of String)) As String
        Dim mail As Object = Nothing
        Try
            mail = GetSourceMail(session)
            If mail Is Nothing Then
                Return "Bitte w�hlen Sie eine einzelne E-Mail aus."
            End If
            If Not IsSameMailAsPrepared(session, mail) Then
                Return "Die angezeigte Mail hat sich geändert. Bitte MailDrop erneut öffnen."
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
            ' Explorer is intentionally NOT released: app.ActiveExplorer() returns the same RCW
            ' as _currentExplorer in ThisAddIn. FinalReleaseComObject on it would destroy the
            ' SelectionChange event connection permanently (see ReadMailMeta).
        End Try
    End Function

End Module
