Imports System.Windows.Forms
Imports System.Windows.Forms.Integration

Public Class MailDropWpfHostControl
    Inherits UserControl

    Public Sub New()
        Dim host As New ElementHost()
        host.Dock = DockStyle.Fill
        host.Child = New MailDropWpfTaskPane()
        Me.Controls.Add(host)
    End Sub
End Class
