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

    ' SMTP-Adresse des Absenders. Bei Absendern aus derselben Exchange-Organisation ist
    ' SenderEmailType "EX" und SenderEmailAddress ein X.500-Pfad ("/O=EXCHANGELABS/OU=...") ohne
    ' "@" - AbsenderDomain blieb dann leer und das Feature war fuer alle internen Mails wertlos.
    ' Fallback ueber AddressEntry.GetExchangeUser().PrimarySmtpAddress; schlaegt das fehl (z.B.
    ' offline ohne Adressbuch), bleibt es beim leeren Ergebnis wie bisher.
    Private Function GetSenderSmtpAddress(mail As Object) As String
        Dim adresse As String = Nothing
        Try
            adresse = mail.SenderEmailAddress
            If Not String.Equals(CStr(mail.SenderEmailType), "EX", StringComparison.OrdinalIgnoreCase) Then Return adresse
        Catch ex As Exception
            Debug.WriteLine($"[MailUtils] GetSenderSmtpAddress: {ex.Message}")
            Return adresse
        End Try
        Dim sender As Object = Nothing
        Dim exchangeUser As Object = Nothing
        Try
            sender = mail.Sender
            If sender IsNot Nothing Then exchangeUser = sender.GetExchangeUser()
            If exchangeUser IsNot Nothing Then Return exchangeUser.PrimarySmtpAddress
        Catch ex As Exception
            Debug.WriteLine($"[MailUtils] GetSenderSmtpAddress (Exchange): {ex.Message}")
        Finally
            ReleaseComObjectSafe(exchangeUser)
            ReleaseComObjectSafe(sender)
        End Try
        Return adresse
    End Function

    ' Liest die Metadaten der ausgewählten Mail und befüllt die Properties der übergebenen Session
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
            Dim smtpAdresse = GetSenderSmtpAddress(mail)
            If Not String.IsNullOrEmpty(smtpAdresse) AndAlso smtpAdresse.Contains("@") Then
                Dim emailParts = smtpAdresse.Split("@"c)
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

    ' Speichert die selektierten Anhänge der Mail. Zuordnung per Index (AnhangZiel.OutlookIndex aus
    ' ReadAttachmentNames), nicht per Dateiname: der Zielname kann vom Originalnamen abweichen
    ' (Umbenennen-Dialog, Nummerierung gleichnamiger Anhänge). Zuvor wurde per Name gesucht -
    ' umbenannte Anhänge wurden dadurch stillschweigend nicht gespeichert, und von zwei
    ' gleichnamigen Anhängen landete zweimal der erste. Die Indizes sind stabil, weil
    ' IsSameMailAsPrepared sicherstellt, dass es noch dieselbe Mail ist.
    Public Function SaveMailAttachments(session As Session, anhaenge As List(Of InputChecker.AnhangZiel)) As String
        Dim mail As Object = Nothing
        Try
            mail = GetSourceMail(session)
            If mail Is Nothing Then
                Return "Bitte wählen Sie eine einzelne E-Mail aus."
            End If
            If Not IsSameMailAsPrepared(session, mail) Then
                Return "Die angezeigte Mail hat sich geändert. Bitte MailDrop erneut öffnen."
            End If
            Dim anzahl As Integer = mail.Attachments.Count
            For Each anhang In anhaenge
                If anhang.OutlookIndex < 1 OrElse anhang.OutlookIndex > anzahl Then
                    Return $"Anhang Nr. {anhang.OutlookIndex} wurde in der Mail nicht gefunden."
                End If
                Dim att As Object = Nothing
                Try
                    att = mail.Attachments(anhang.OutlookIndex)
                    att.SaveAsFile(anhang.Zielpfad)
                Finally
                    ReleaseComObjectSafe(att)
                End Try
            Next
            Return String.Empty
        Catch ex As Exception
            Return $"Fehler beim Speichern der Anhänge: {ex.Message}"
        Finally
            ReleaseComObjectSafe(mail)
            ' Explorer is intentionally NOT released: app.ActiveExplorer() returns the same RCW
            ' as _currentExplorer in ThisAddIn. FinalReleaseComObject on it would destroy the
            ' SelectionChange event connection permanently (see ReadMailMeta).
        End Try
    End Function

End Module
