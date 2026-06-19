Imports Microsoft.Office.Tools.Ribbon
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core
Imports System.IO

Public Class ThisAddIn
    Private ribbonObj As MailDropRibbon
    Private taskPane As Microsoft.Office.Tools.CustomTaskPane
    Private Shared _currentDatabaseManager As SessionDatabaseManager
    Private Shared ReadOnly _databaseManagerLock As New Object()

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
        explorer = Application.ActiveExplorer()
        If explorer IsNot Nothing Then
            AddHandler explorer.SelectionChange, AddressOf Explorer_SelectionChange
        End If

        ' Engine im Hintergrund vorladen, damit "Mail ablegen" beim ersten Klick schneller oeffnet.
        SuggestionEngine.PreloadSharedInstanceInBackground(1500)
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        Try
            SuggestionEngine.DisposeSharedInstance()
        Catch
            ' Shutdown cleanup should never block Outlook exit.
        End Try
    End Sub

    Private WithEvents explorer As Outlook.Explorer

    ' Kapselt die Logik für die TaskPane-Initialisierung und Editierbarkeit
    Private Sub MailSelected()
        Dim wpfTaskPane As MailDropWpfTaskPane = GetWpfTaskPane()
        If wpfTaskPane IsNot Nothing Then
            If wpfTaskPane.SingleMailSelected() Then
                wpfTaskPane.Session.PrepareSession()
                wpfTaskPane.SetEditMode(True)
            Else
                wpfTaskPane.Session.Reset()
                wpfTaskPane.SetEditMode(False)
            End If
        End If
    End Sub

    Private Sub Explorer_SelectionChange()
        MailSelected()
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
        Dim wpfTaskPane As MailDropWpfTaskPane = Nothing

        If taskPane Is Nothing Then
            Dim paneControl As New MailDropWpfHostControl()
            Dim wpfPane = TryCast(paneControl.Controls(0), System.Windows.Forms.Integration.ElementHost)
            If wpfPane IsNot Nothing Then
                wpfTaskPane = TryCast(wpfPane.Child, MailDropWpfTaskPane)
            End If
            taskPane = Me.CustomTaskPanes.Add(paneControl, "Mail ablegen")
            taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
            taskPane.Width = 500
        Else
            Dim wpfPane = TryCast(taskPane.Control.Controls(0), System.Windows.Forms.Integration.ElementHost)
            If wpfPane IsNot Nothing Then
                wpfTaskPane = TryCast(wpfPane.Child, MailDropWpfTaskPane)
            End If
        End If

        MailSelected()
        taskPane.Visible = True
    End Sub

    Public Sub HideTaskPane()
        If taskPane IsNot Nothing Then
            taskPane.Visible = False
        End If
    End Sub

End Class
