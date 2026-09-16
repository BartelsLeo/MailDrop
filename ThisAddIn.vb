Imports Microsoft.Office.Tools.Ribbon
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core
Imports System.IO
Imports System.Diagnostics
Imports System.Windows.Forms
Imports System.Threading.Tasks

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
        ' Registered first, before anything else in this Sub can throw: catches exceptions on
        ' background threads (e.g. an unobserved Task fault in SuggestionEngine's preload) that
        ' would otherwise crash or destabilize the process silently, with zero trace in error.log.
        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf OnUnhandledException
        AddHandler TaskScheduler.UnobservedTaskException, AddressOf OnUnobservedTaskException

        Dim sw As Stopwatch = Stopwatch.StartNew()
        Debug.WriteLine("[ThisAddIn] Startup BEGIN")
        ' Written unconditionally, before any code below that could throw, so that even a total
        ' Startup failure leaves a record of which build was running and how far it got - this was
        ' previously missing: an exception anywhere in this Sub's body (it was not Try/Catch-wrapped)
        ' meant NO log entry at all, since the only Logger.LogInfo call was the very last line.
        Logger.LogInfo("Startup", "ThisAddIn_Startup entered.")
        LogStartupEnvironmentInfo()
        LogDependencyFilePresence()

        Try
            ' Application.Explorers is a synchronous COM call that can block on Outlook's own
            ' mail-store/session init (Exchange cached-mode sync, PST/OST mounting) - observed
            ' taking 1.7-2s and tripping Outlook's slow-add-in watchdog (threshold 1000 ms),
            ' which then auto-disables MailDrop. Deferred to a short UI-thread timer tick (COM-safe,
            ' Explorer objects need STA affinity so Task.Run is not an option here) so Startup itself
            ' returns near-instantly, same pattern as PreloadTaskPaneInBackground below.
            Dim explorerTimer As New System.Windows.Forms.Timer()
            explorerTimer.Interval = 100
            AddHandler explorerTimer.Tick, Sub(s, e)
                explorerTimer.Stop()
                explorerTimer.Dispose()
                Try
                    Dim explorerSw As Stopwatch = Stopwatch.StartNew()
                    _explorers = Application.Explorers
                    If _explorers.Count > 0 Then
                        _currentExplorer = TryCast(_explorers.Item(1), Outlook.Explorer)
                        Debug.WriteLine($"[ThisAddIn] Deferred explorer wiring: _currentExplorer set. ({sw.ElapsedMilliseconds} ms)")
                    Else
                        Debug.WriteLine($"[ThisAddIn] Deferred explorer wiring: no explorer available yet, waiting for NewExplorer. ({sw.ElapsedMilliseconds} ms)")
                    End If
                    ' Release-safe (survives DefineDebug=false): how long the deferred Application.Explorers
                    ' COM call itself took, to confirm/refute it as a bottleneck in the field - see
                    ' CLAUDE.md "Startup timing instrumentation".
                    Logger.LogInfo("Startup timing", $"Deferred Application.Explorers call took {explorerSw.ElapsedMilliseconds} ms (fired {sw.ElapsedMilliseconds} ms after Startup returned).")
                Catch ex As Exception
                    Debug.WriteLine($"[ThisAddIn] Deferred explorer wiring failed: {ex.Message}")
                    Logger.LogError("ThisAddIn_Startup: deferred explorer wiring", ex)
                End Try
            End Sub
            explorerTimer.Start()

            ' Engine im Hintergrund vorladen (inkl. EmbeddingService-Warmup).
            SuggestionEngine.PreloadSharedInstanceInBackground(1500)

            ' Task pane im Hintergrund vorerstellen: Timer feuert auf dem UI-Thread (WPF/COM-sicher).
            ' Verzögerung > Engine-Preload damit SQLite-Init und ONNX-Load nicht gleichzeitig laufen.
            PreloadTaskPaneInBackground(delayMs:=4000)

            Debug.WriteLine($"[ThisAddIn] Startup END – preloads queued. Total={sw.ElapsedMilliseconds} ms")
            ' Release-safe: Outlook's own slow-add-in watchdog (threshold 1000 ms, see CLAUDE.md) has
            ' been observed disabling MailDrop with "Benötigte Zeit" around 1.7-2s even after deferring
            ' Application.Explorers out of this Sub's body, meaning that fix alone was not sufficient -
            ' this trace tells us whether Startup's OWN synchronous body is actually fast now (pointing
            ' the remaining cost at CLR/ClickOnce/assembly-load time before Startup even runs, which is
            ' outside this Sub's control) or still slow (meaning something else in here still blocks).
            Logger.LogInfo("Startup timing", $"ThisAddIn_Startup Sub body returned after {sw.ElapsedMilliseconds} ms.")
        Catch ex As Exception
            ' Previously unguarded: an exception here (e.g. a static-initializer failure in
            ' SuggestionEngine, or the Timer construction itself) propagated out of Startup
            ' unhandled. Whether VSTO/Outlook's own resiliency tracking treats a thrown Startup
            ' exception the same as a slow load is unconfirmed, but either way this ensures the
            ' failure is at least logged instead of leaving error.log empty.
            Debug.WriteLine($"[ThisAddIn] Startup FAILED: {ex.Message}")
            Logger.LogError("ThisAddIn_Startup", ex)
        End Try
    End Sub

    Private Sub OnUnhandledException(sender As Object, e As UnhandledExceptionEventArgs)
        Try
            Dim ex As Exception = TryCast(e.ExceptionObject, Exception)
            If ex IsNot Nothing Then
                Logger.LogError("AppDomain.UnhandledException", ex)
            Else
                Logger.LogInfo("AppDomain.UnhandledException", $"Non-Exception object thrown: {e.ExceptionObject}. IsTerminating={e.IsTerminating}")
            End If
        Catch
            ' Never let the crash handler itself crash.
        End Try
    End Sub

    Private Sub OnUnobservedTaskException(sender As Object, e As UnobservedTaskExceptionEventArgs)
        Try
            Logger.LogError("TaskScheduler.UnobservedTaskException", e.Exception)
            e.SetObserved()
        Catch
        End Try
    End Sub

    ' Gibt Version/Build-/Umgebungsinfo aus, damit ein spaeteres error.log (oder dessen
    ' Fehlen) einer konkreten Version und einem konkreten Installationspfad (ClickOnce vs.
    ' lokal/vstolocal) zugeordnet werden kann - siehe CLAUDE.md "Distribution" fuer den
    ' Hintergrund der wiederholten Deaktivierungs-/Deployment-Probleme.
    Private Sub LogStartupEnvironmentInfo()
        Try
            Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
            Dim asmVersion As String = asm.GetName().Version.ToString()

            Dim fileVersion As String = "?"
            Try
                fileVersion = FileVersionInfo.GetVersionInfo(asm.Location).FileVersion
            Catch
            End Try

            Dim buildDate As String = "?"
            Try
                ' Kein eingebetteter Linker-Timestamp wird ausgelesen - LastWriteTime der DLL ist
                ' ein Näherungswert fuer "wann deployed/gebaut", aber verlaesslich genug fuer Diagnose.
                buildDate = File.GetLastWriteTime(asm.Location).ToString("yyyy-MM-dd HH:mm:ss")
            Catch
            End Try

            Dim isClickOnce As Boolean = False
            Dim deploymentVersion As String = "n/a"
            Dim deploymentUpdateLocation As String = "n/a"
            Try
                isClickOnce = System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed
                If isClickOnce Then
                    Dim dep = System.Deployment.Application.ApplicationDeployment.CurrentDeployment
                    deploymentVersion = dep.CurrentVersion.ToString()
                    deploymentUpdateLocation = dep.UpdateLocation?.ToString()
                End If
            Catch
            End Try

            Dim outlookVersion As String = "?"
            Try
                outlookVersion = Application.Version
            Catch
            End Try

            Dim proc = Process.GetCurrentProcess()
            Dim elapsedSinceProcessStartMs As Double = 0
            Try
                elapsedSinceProcessStartMs = (DateTime.Now - proc.StartTime).TotalMilliseconds
            Catch
            End Try

            Logger.LogInfo("Startup version info",
                $"AssemblyVersion={asmVersion} FileVersion={fileVersion} AssemblyLastWriteTime(BuildDate-Näherung)={buildDate} " &
                $"ClickOnceDeployed={isClickOnce} ClickOnceVersion={deploymentVersion} ClickOnceUpdateLocation={deploymentUpdateLocation} " &
                $"OutlookVersion={outlookVersion} OS={Environment.OSVersion.VersionString} Is64BitProcess={Environment.Is64BitProcess} " &
                $"CLR={Environment.Version} ProcessStartTime={proc.StartTime:yyyy-MM-dd HH:mm:ss} ElapsedSinceProcessStart={elapsedSinceProcessStartMs:F0}ms " &
                $"AssemblyPath={asm.Location}")
        Catch ex As Exception
            Logger.LogError("ThisAddIn_Startup: LogStartupEnvironmentInfo", ex)
        End Try
    End Sub

    ' Prueft die Anwesenheit der bekannten deployment-anfaelligen Abhaengigkeiten (siehe
    ' CLAUDE.md-Abschnitt zu SQLite.Interop.dll-Platzierungsproblemen) direkt beim Start, statt
    ' erst auf eine spaetere DllNotFoundException zu warten, die den eigentlichen Pfadfehler
    ' nicht immer eindeutig erkennen laesst.
    Private Sub LogDependencyFilePresence()
        Try
            Dim baseDir As String = AppDomain.CurrentDomain.BaseDirectory
            Dim relativePaths() As String = {"x86\SQLite.Interop.dll", "x64\SQLite.Interop.dll", "Models\model.onnx", "Models\vocab.txt"}
            Dim sb As New Text.StringBuilder()
            sb.Append($"BaseDirectory={baseDir}")
            For Each rel As String In relativePaths
                Dim full As String = Path.Combine(baseDir, rel)
                sb.Append($" | {rel}={File.Exists(full)}")
            Next
            Logger.LogInfo("Startup dependency check", sb.ToString())
        Catch ex As Exception
            Logger.LogError("ThisAddIn_Startup: LogDependencyFilePresence", ex)
        End Try
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
        ' If a future error.log never shows this line before the next "Startup entered", Outlook
        ' (or the whole process) was killed/crashed rather than exited normally - useful to tell
        ' apart from an auto-disable, which instead shows no Startup entry at all for a run.
        Logger.LogInfo("Shutdown", "ThisAddIn_Shutdown entered.")
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
