Imports Microsoft.Office.Tools.Ribbon
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core
Imports System.IO
Imports System.Threading.Tasks

Public Class ThisAddIn
    Private ribbonObj As MailDropRibbon
    Private taskPane As Microsoft.Office.Tools.CustomTaskPane
    Private _wpfTaskPane As MailDropWpfTaskPane
    Private _activeExplorer As Outlook.Explorer

    Public Shared ReadOnly Property DbPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MailDrop", "sessions.db")
    Public Shared ReadOnly Property DbDirectory As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MailDrop")
    Public Shared Property CurrentDatabaseManager As SessionDatabaseManager
    Public Shared CachedSuggestionEngine As SuggestionEngine

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        CurrentDatabaseManager = New SessionDatabaseManager()
        Task.Run(Sub() CachedSuggestionEngine = New SuggestionEngine())
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        DisconnectExplorer()
    End Sub

    Private Sub DisconnectExplorer()
        If _activeExplorer IsNot Nothing Then
            RemoveHandler _activeExplorer.SelectionChange, AddressOf Explorer_SelectionChange
            Marshal.ReleaseComObject(_activeExplorer)
            _activeExplorer = Nothing
        End If
    End Sub

    Private Sub Explorer_SelectionChange()
        Try
            MailSelected()
        Catch ex As Exception
            ErrorHandler.ShowError(ex)
        End Try
    End Sub

    Private Sub MailSelected()
        If _wpfTaskPane IsNot Nothing Then
            If _wpfTaskPane.SingleMailSelected() Then
                _wpfTaskPane.Session.PrepareSession()
                _wpfTaskPane.SetEditMode(True)
            Else
                _wpfTaskPane.Session.Reset()
                _wpfTaskPane.SetEditMode(False)
            End If
        End If
    End Sub

    Protected Overrides Function CreateRibbonExtensibilityObject() As IRibbonExtensibility
        ribbonObj = New MailDropRibbon()
        Return ribbonObj
    End Function

    Public Sub MailAblegen_Click(control As Object)
        Try
            If taskPane Is Nothing Then
                Dim paneControl As New MailDropWpfHostControl()
                _wpfTaskPane = paneControl.WpfTaskPane
                taskPane = Me.CustomTaskPanes.Add(paneControl, "Mail ablegen")
                taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
                taskPane.Width = 1000
            End If
            If _activeExplorer Is Nothing Then
                _activeExplorer = Application.ActiveExplorer()
                AddHandler _activeExplorer.SelectionChange, AddressOf Explorer_SelectionChange
            End If
            MailSelected()
            taskPane.Visible = True
        Catch ex As Exception
            ErrorHandler.ShowError(ex)
        End Try
    End Sub

    Public Sub HideTaskPane()
        If taskPane IsNot Nothing Then
            taskPane.Visible = False
        End If
        DisconnectExplorer()
    End Sub

End Class
