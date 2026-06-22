Imports System.IO
Imports System.Collections.ObjectModel

Public Module InputChecker
    ' Hilfsklasse für die geprüften Pfade
    Public Class CheckedInputResult
        Public Property CheckedAblageOrdner As String
        Public Property CheckedMsgZielpfad As String
        Public Property CheckedAnhZielpfade As List(Of String)
        Public Property ErrorMessage As String
    End Class

    ' Prüft Ordnernamen und Pfadlänge
    Public Function CheckFolderNameAndPath(folderPath As String) As String
        Dim invalidChars = Path.GetInvalidPathChars().Concat(Path.GetInvalidFileNameChars()).ToArray()
        Dim folderName = Path.GetFileName(folderPath)
        If folderName.IndexOfAny(invalidChars) >= 0 Then
            Return $"Der Ablageordner enthält ungültige Zeichen: {folderName}"
        End If
        If folderPath.Length > 255 Then
            Dim diff = folderPath.Length - 255
            Return $"Der Ablageordner-Pfad ist zu lang (max. 255 Zeichen).{vbCrLf}Pfad: {folderPath}{vbCrLf}Überlänge: {diff} Zeichen."
        End If
        Return String.Empty
    End Function

    ' Prüft Dateinamen und Pfadlänge
    Public Function CheckFileNameAndPath(filePath As String) As String
        Dim invalidChars = Path.GetInvalidPathChars().Concat(Path.GetInvalidFileNameChars()).ToArray()
        Dim fileName = Path.GetFileName(filePath)
        If fileName.IndexOfAny(invalidChars) >= 0 Then
            Return $"Der Dateiname enthält ungültige Zeichen: {fileName}"
        End If
        If filePath.Length > 255 Then
            Dim diff = filePath.Length - 255
            Return $"Der vollständige Dateipfad ist zu lang (max. 255 Zeichen).{vbCrLf}Pfad: {filePath}{vbCrLf}Überlänge: {diff} Zeichen."
        End If
        Return String.Empty
    End Function

    ' Zeigt einen Dialog zur Umbenennung, gibt neuen Namen oder String.Empty zurück
    Public Function ShowAttachmentRenameDialog(currentName As String, basePath As String,
                                               Optional hintText As String = Nothing,
                                               Optional windowTitle As String = Nothing) As String
        Dim dlg As New AttachmentRenameDialog(currentName, basePath, hintText, windowTitle)
        If System.Windows.Application.Current IsNot Nothing AndAlso System.Windows.Application.Current.MainWindow IsNot Nothing Then
            dlg.Owner = System.Windows.Application.Current.MainWindow
        End If
        If dlg.ShowDialog() = True Then
            Return dlg.FileName
        Else
            Return String.Empty
        End If
    End Function

    ' Prüft alle Eingaben (Projektpfad, Projektstruktur, Ablageordner, msg-Dateiname)
    Public Function CheckInput(session As Session) As CheckedInputResult
        Dim result As New CheckedInputResult()
        result.CheckedAnhZielpfade = New List(Of String)()
        Dim projektPfad As String = session.ProjektPfad
        If String.IsNullOrWhiteSpace(projektPfad) OrElse Not Directory.Exists(projektPfad) Then
            result.ErrorMessage = "Bitte wählen Sie einen gültigen Projektpfad aus."
            Return result
        End If
        If String.IsNullOrWhiteSpace(session.ProjektstrukturPfad) Then
            result.ErrorMessage = "Bitte wählen Sie eine gültige Projektstruktur aus."
            Return result
        End If
        Dim projektstrukturPfad As String = Path.Combine(projektPfad, session.ProjektstrukturPfad)
        If Not Directory.Exists(projektstrukturPfad) Then
            result.ErrorMessage = "Bitte wählen Sie eine gültige Projektstruktur aus."
            Return result
        End If
        Dim ablageOrdnerPfad As String = Path.Combine(projektstrukturPfad, session.AblageordnerAufgeloest)
        Dim ablageOrdnerCheck = CheckFolderNameAndPath(ablageOrdnerPfad)
        If ablageOrdnerCheck <> String.Empty Then
            result.ErrorMessage = ablageOrdnerCheck
            Return result
        End If
        result.CheckedAblageOrdner = ablageOrdnerPfad
        If String.IsNullOrWhiteSpace(session.MsgDateinameAufgeloest) Then
            result.ErrorMessage = "Bitte geben Sie einen gültigen Dateinamen für die E-Mail an."
            Return result
        End If
        Dim msgZielPfad As String = Path.Combine(ablageOrdnerPfad, session.MsgDateinameAufgeloest)
        Dim msgDateinameCheck = CheckFileNameAndPath(msgZielPfad)
        If msgDateinameCheck <> String.Empty Then
            result.ErrorMessage = msgDateinameCheck
            Return result
        End If
        result.CheckedMsgZielpfad = msgZielPfad
        If session.AnhaengeAblegen Then
            Dim app As Outlook.Application = Globals.ThisAddIn.Application
            Dim explorer = app.ActiveExplorer()
            If explorer IsNot Nothing AndAlso explorer.Selection.Count = 1 Then
                Dim mail = TryCast(explorer.Selection.Item(1), Outlook.MailItem)
                If mail IsNot Nothing Then
                    For i As Integer = 1 To mail.Attachments.Count
                        Dim att = mail.Attachments(i)
                        Dim anhangName = att.FileName
                        Dim anhangPfad = Path.Combine(ablageOrdnerPfad, anhangName)
                        If anhangPfad.Length > 255 Then
                            Dim newName = ShowAttachmentRenameDialog(anhangName, ablageOrdnerPfad)
                            If String.IsNullOrEmpty(newName) Then
                                Continue For
                            End If
                            anhangName = newName
                            anhangPfad = Path.Combine(ablageOrdnerPfad, anhangName)
                        End If
                        Dim anhangNameCheck = CheckFileNameAndPath(anhangPfad)
                        If anhangNameCheck <> String.Empty Then
                            result.ErrorMessage = anhangNameCheck
                            Return result
                        End If
                        result.CheckedAnhZielpfade.Add(anhangPfad)
                    Next
                End If
            End If
        End If
        Return result
    End Function
End Module
