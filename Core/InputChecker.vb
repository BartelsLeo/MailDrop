Imports System.IO
Imports System.Runtime.InteropServices

Public Module InputChecker
    Private Sub ReleaseComObjectSafe(comObject As Object)
        If comObject Is Nothing Then Return
        If Marshal.IsComObject(comObject) Then
            Marshal.FinalReleaseComObject(comObject)
        End If
    End Sub

    ' Hilfsklasse für die geprüften Pfade
    Public Class CheckedInputResult
        Public Property CheckedAblageOrdner As String
        Public Property CheckedMsgZielpfad As String
        Public Property CheckedAnhZielpfade As List(Of String)
        Public Property ErrorMessage As String
        Public Property DuplicateWarning As String
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

    ' Stellt sicher, dass ein abgeleiteter Zielpfad seinen Basisordner nicht verlässt.
    ' Nötig, weil Ablageordner- und Dateinamen aus frei editierbaren Feldern (Schema, Titel)
    ' entstehen: Path.Combine verwirft sein erstes Argument vollständig, sobald das zweite ein
    ' Wurzelpfad ist ("C:\..." oder "\..."), und ".." führt aus dem Projektordner heraus. Die Mail
    ' würde dann ohne jeden Hinweis außerhalb des Projekts abgelegt.
    Public Function IsInsideBaseFolder(basePath As String, candidatePath As String) As Boolean
        Try
            Dim normalizedBase = Path.GetFullPath(basePath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            Dim normalizedCandidate = Path.GetFullPath(candidatePath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            If String.Equals(normalizedBase, normalizedCandidate, StringComparison.OrdinalIgnoreCase) Then Return True
            Return normalizedCandidate.StartsWith(normalizedBase & Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
        Catch
            ' GetFullPath wirft bei syntaktisch unbrauchbaren Pfaden - die sind ohnehin nicht verwendbar.
            Return False
        End Try
    End Function

    Public Function ShowAttachmentRenameDialog(currentName As String, basePath As String) As String
        Dim dlg As New AttachmentRenameDialog(currentName, basePath)
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
        ' ProjektstrukturPfad = Nothing bedeutet "noch nichts ausgewaehlt" (ungueltig).
        ' ProjektstrukturPfad = String.Empty bedeutet "der Projektpfad-Root-Knoten wurde explizit
        ' ausgewaehlt" (gueltig - Mail wird direkt in ProjektPfad abgelegt, siehe TreeView-Root-Knoten
        ' in DirectoryTreeHelper.BuildDirectoryTree). Nicht IsNullOrWhiteSpace verwenden, da das auch
        ' den gueltigen leeren String ablehnen wuerde.
        If session.ProjektstrukturPfad Is Nothing Then
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
        If Not IsInsideBaseFolder(projektstrukturPfad, ablageOrdnerPfad) Then
            result.ErrorMessage = "Der aufgelöste Ablageordner liegt außerhalb der gewählten Projektstruktur. Bitte prüfen Sie das Ablageordner-Schema und den Titel."
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
        If Not IsInsideBaseFolder(ablageOrdnerPfad, msgZielPfad) Then
            result.ErrorMessage = "Der aufgelöste msg-Dateiname liegt außerhalb des Ablageordners. Bitte prüfen Sie das msg-Dateinamen-Schema."
            Return result
        End If
        result.CheckedMsgZielpfad = msgZielPfad
        If session.AnhaengeAblegen Then
            For Each item In session.Anhaenge
                If Not item.IsSelected Then Continue For
                Dim anhangName = item.Name
                Dim anhangPfad = Path.Combine(ablageOrdnerPfad, anhangName)
                If anhangPfad.Length > 255 Then
                    Dim newName = ShowAttachmentRenameDialog(anhangName, ablageOrdnerPfad)
                    If String.IsNullOrEmpty(newName) Then Continue For
                    anhangName = newName
                    anhangPfad = Path.Combine(ablageOrdnerPfad, anhangName)
                End If
                Dim anhangNameCheck = CheckFileNameAndPath(anhangPfad)
                If anhangNameCheck <> String.Empty Then
                    result.ErrorMessage = anhangNameCheck
                    Return result
                End If
                If Not IsInsideBaseFolder(ablageOrdnerPfad, anhangPfad) Then
                    result.ErrorMessage = $"Der Anhang '{anhangName}' würde außerhalb des Ablageordners gespeichert werden."
                    Return result
                End If
                result.CheckedAnhZielpfade.Add(anhangPfad)
            Next
        End If
        Dim existingFiles As New List(Of String)()
        If File.Exists(result.CheckedMsgZielpfad) Then
            existingFiles.Add(Path.GetFileName(result.CheckedMsgZielpfad))
        End If
        For Each anhPfad In result.CheckedAnhZielpfade
            If File.Exists(anhPfad) Then existingFiles.Add(Path.GetFileName(anhPfad))
        Next
        If existingFiles.Count > 0 Then
            result.DuplicateWarning = "Bereits vorhanden: " & String.Join(", ", existingFiles)
        End If
        Return result
    End Function
End Module
