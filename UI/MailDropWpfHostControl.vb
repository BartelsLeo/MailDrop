Imports System.Windows.Forms
Imports System.Windows.Forms.Integration

Public Class MailDropWpfHostControl
    Inherits UserControl

    Private _wpfTaskPane As MailDropWpfTaskPane

    Public ReadOnly Property WpfTaskPane As MailDropWpfTaskPane
        Get
            Return _wpfTaskPane
        End Get
    End Property

    Public Sub New()
        _wpfTaskPane = New MailDropWpfTaskPane()
        Dim host As New ElementHost()
        host.Dock = DockStyle.Fill
        host.Child = _wpfTaskPane
        Me.Controls.Add(host)
    End Sub

End Class
