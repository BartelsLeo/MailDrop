Imports Microsoft.Office.Tools.Ribbon
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core

Public Class ThisAddIn
    Private ribbonObj As MailDropRibbon
    Private taskPane As Microsoft.Office.Tools.CustomTaskPane

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

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
            taskPane.Width = 300
        Else
            Dim wpfPane = TryCast(taskPane.Control.Controls(0), System.Windows.Forms.Integration.ElementHost)
            If wpfPane IsNot Nothing Then
                wpfTaskPane = TryCast(wpfPane.Child, MailDropWpfTaskPane)
            End If
        End If

        ' Session Vorbereiten
        If wpfTaskPane IsNot Nothing Then
            wpfTaskPane.Session.PrepareSession()
        End If

        taskPane.Visible = True
    End Sub

End Class
