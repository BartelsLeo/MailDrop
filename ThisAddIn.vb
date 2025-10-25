Imports Microsoft.Office.Tools.Ribbon
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core

Public Class ThisAddIn
    Private ribbonObj As MailDropRibbon

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
        ' Hier die gewünschte Logik einfügen
        System.Windows.Forms.MessageBox.Show("Mail ablegen geklickt!")
    End Sub

End Class
