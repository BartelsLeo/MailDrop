Imports System.IO

' Datei-basiertes Fehlerlogging fuer Release-Builds, in denen Debug.WriteLine
' vollstaendig herauskompiliert wird (DefineDebug=false, siehe MailDrop.vbproj).
Public Module Logger

    Private ReadOnly logLock As New Object()

    Public Sub LogError(context As String, ex As Exception)
        Try
            SyncLock logLock
                Directory.CreateDirectory(ThisAddIn.DbDirectory)
                Dim logPath As String = Path.Combine(ThisAddIn.DbDirectory, "error.log")
                Dim entry As String = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{context}] {ex}{Environment.NewLine}"
                File.AppendAllText(logPath, entry)
            End SyncLock
        Catch
            ' Logging darf niemals selbst eine Exception werfen.
        End Try
    End Sub

    ' Wie LogError, aber ohne Exception - fuer Release-sichere Timing-/Diagnose-Traces, die
    ' Debug.WriteLine (in Release herauskompiliert) ersetzen, wenn deren Wert das ueber die
    ' Session-Laufzeit hinaus fortbestehende Auffindenkoennen rechtfertigt (aktuell nur
    ' ThisAddIn_Startup-Timing, siehe CLAUDE.md "Startup timing instrumentation").
    Public Sub LogInfo(context As String, message As String)
        Try
            SyncLock logLock
                Directory.CreateDirectory(ThisAddIn.DbDirectory)
                Dim logPath As String = Path.Combine(ThisAddIn.DbDirectory, "error.log")
                Dim entry As String = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{context}] {message}{Environment.NewLine}"
                File.AppendAllText(logPath, entry)
            End SyncLock
        Catch
            ' Logging darf niemals selbst eine Exception werfen.
        End Try
    End Sub

End Module
