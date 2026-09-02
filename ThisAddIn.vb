Imports Microsoft.Office.Tools.Ribbon
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core
Imports System.IO
Imports System.Diagnostics
Imports System.Windows.Forms

Public Class ThisAddIn
    Private ribbonObj As MailDropRibbon
    Private taskPane As Microsoft.Office.Tools.CustomTaskPane
    Private Shared _currentDatabaseManager As SessionDatabaseManager
    Private Shared ReadOnly _databaseManagerLock As New Object()
    Private WithEvents _explorers As Outlook.Explorers
    Private WithEvents _currentExplorer As Outlook.Explorer

    Public Shared ReadOnly Property DbPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MailDrop", "sessions.db")
    Public Shared ReadOnly Property DbDirectory As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MailDrop")
    Public Shared ReadOnly Property CurrentDatabaseManager As SessionDatabaseManager
        Get
            SyncLock _databaseManagerLock
                If _currentDatabaseManager Is Nothing Then
                    _currentDatabaseManager = New SessionDatabaseManager()
                End If
                Return _currentDatabaseManager
            End SyncLock
        End Get
    End Property

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Dim sw As Stopwatch = Stopwatch.StartNew()
        Debug.WriteLine("[ThisAddIn] Startup BEGIN")

        _explorers = Application.Explorers
        If _explorers.Count > 0 Then
            _currentExplorer = TryCast(_explorers.Item(1), Outlook.Explorer)
            Debug.WriteLine($"[ThisAddIn] Startup: _currentExplorer set. ({sw.ElapsedMilliseconds} ms)")
        Else
            Debug.WriteLine($"[ThisAddIn] Startup: no explorer available yet, waiting for NewExplorer. ({sw.ElapsedMilliseconds} ms)")
        End If

        ' Engine im Hintergrund vorladen (inkl. EmbeddingService-Warmup).
        SuggestionEngine.PreloadSharedInstanceInBackground(1500)

        ' Task pane im Hintergrund vorerstellen: Timer feuert auf dem UI-Thread (WPF/COM-sicher).
        ' Verzögerung > Engine-Preload damit SQLite-Init und ONNX-Load nicht gleichzeitig laufen.
        PreloadTaskPaneInBackground(delayMs:=4000)

        Debug.WriteLine($"[ThisAddIn] Startup END – preloads queued. Total={sw.ElapsedMilliseconds} ms")
    End Sub

    Private Sub PreloadTaskPaneInBackground(delayMs As Integer)
        Dim timer As New System.Windows.Forms.Timer()
        timer.Interval = delayMs
        AddHandler timer.Tick, Sub(s, e)
            timer.Stop()
            timer.Dispose()
            If taskPane IsNot Nothing Then
                Debug.WriteLine("[ThisAddIn] PreloadTaskPane: already created, skipping.")
                Return
            End If
            Try
                Dim sw As Stopwatch = Stopwatch.StartNew()
                Debug.WriteLine("[ThisAddIn] PreloadTaskPane: creating hidden task pane…")
                CreateAndRegisterTaskPane()
                taskPane.Visible = False
                Debug.WriteLine($"[ThisAddIn] PreloadTaskPane: done in {sw.ElapsedMilliseconds} ms — task pane ready.")
            Catch ex As Exception
                Debug.WriteLine($"[ThisAddIn] PreloadTaskPane failed: {ex.Message}")
                Logger.LogError("PreloadTaskPaneInBackground", ex)
            End Try
        End Sub
        timer.Start()
    End Sub

    Private Sub CreateAndRegisterTaskPane()
        Dim sw As Stopwatch = Stopwatch.StartNew()
        Dim paneControl As New MailDropWpfHostControl()
        Debug.WriteLine($"[ThisAddIn]   MailDropWpfHostControl created: {sw.ElapsedMilliseconds} ms")
        taskPane = Me.CustomTaskPanes.Add(paneControl, "Mail ablegen")
        taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        taskPane.Width = 500
        Debug.WriteLine($"[ThisAddIn]   Task pane registered and docked: {sw.ElapsedMilliseconds} ms")
    End Sub

    Private Sub _explorers_NewExplorer(NewExplorer As Outlook.Explorer) Handles _explorers.NewExplorer
        Debug.WriteLine("[ThisAddIn] NewExplorer fired - updating _currentExplorer.")
        _currentExplorer = NewExplorer
    End Sub

    Private Sub _currentExplorer_SelectionChange() Handles _currentExplorer.SelectionChange
        Debug.WriteLine("[ThisAddIn] Explorer_SelectionChange fired.")
        Try
            MailSelected()
        Catch ex As Exception
            Debug.WriteLine($"[ThisAddIn] SelectionChange: MailSelected failed: {ex.Message}")
            Logger.LogError("SelectionChange: MailSelected", ex)
        End Try
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        Try
            SuggestionEngine.DisposeSharedInstance()
        Catch
            ' Shutdown cleanup should never block Outlook exit.
        End Try
    End Sub

    ' Kapselt die Logik für die TaskPane-Initialisierung und Editierbarkeit
    Private Sub MailSelected()
        Dim wpfTaskPane As MailDropWpfTaskPane = GetWpfTaskPane()
        Debug.WriteLine($"[ThisAddIn] MailSelected: taskPane={If(taskPane IsNot Nothing, "open", "null")}, wpfTaskPane={If(wpfTaskPane IsNot Nothing, "ok", "null")}")
        If wpfTaskPane IsNot Nothing Then
            Dim singleMail = wpfTaskPane.SingleMailSelected()
            Debug.WriteLine($"[ThisAddIn] MailSelected: SingleMailSelected={singleMail}")
            If singleMail Then
                wpfTaskPane.Session.PrepareSession()
                wpfTaskPane.SetEditMode(True)
            Else
                wpfTaskPane.Session.Reset()
                wpfTaskPane.SetEditMode(False)
            End If
        End If
    End Sub

    ' Hilfsmethode, um die WPF TaskPane Instanz zu bekommen
    Private Function GetWpfTaskPane() As MailDropWpfTaskPane
        If taskPane Is Nothing Then Return Nothing
        Dim wpfPane = TryCast(taskPane.Control.Controls(0), System.Windows.Forms.Integration.ElementHost)
        If wpfPane IsNot Nothing Then
            Return TryCast(wpfPane.Child, MailDropWpfTaskPane)
        End If
        Return Nothing
    End Function

    Protected Overrides Function CreateRibbonExtensibilityObject() As IRibbonExtensibility
        ribbonObj = New MailDropRibbon()
        Return ribbonObj
    End Function

    ' Callback for Ribbon button
    Public Sub MailAblegen_Click(control As Object)
        Try
            If taskPane Is Nothing Then
                ' Preload hasn't fired yet (clicked within first 4 s) — create synchronously.
                Debug.WriteLine("[ThisAddIn] MailAblegen_Click: preload not done, creating task pane now…")
                CreateAndRegisterTaskPane()
            Else
                Debug.WriteLine("[ThisAddIn] MailAblegen_Click: task pane was preloaded — showing immediately.")
            End If
        Catch ex As Exception
            ' Task pane could not be created at all — nothing to show, so surface the error directly.
            Debug.WriteLine($"[ThisAddIn] MailAblegen_Click: CreateAndRegisterTaskPane failed: {ex.Message}")
            Logger.LogError("MailAblegen_Click: CreateAndRegisterTaskPane", ex)
            MessageBox.Show($"MailDrop konnte nicht geoeffnet werden:{Environment.NewLine}{ex.Message}{Environment.NewLine}{Environment.NewLine}Details in: {Path.Combine(DbDirectory, "error.log")}",
                             "MailDrop", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try

        Try
            MailSelected()
        Catch ex As Exception
            ' Task pane exists but couldn't be populated (e.g. DB/SQLite init failed) —
            ' still show it instead of silently doing nothing, and log the cause.
            Debug.WriteLine($"[ThisAddIn] MailAblegen_Click: MailSelected failed: {ex.Message}")
            Logger.LogError("MailAblegen_Click: MailSelected", ex)
        End Try

        taskPane.Visible = True
    End Sub

    Public Sub HideTaskPane()
        If taskPane IsNot Nothing Then
            taskPane.Visible = False
        End If
    End Sub

End Class
