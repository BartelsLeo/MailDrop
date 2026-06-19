Imports System.Windows.Forms
Imports System.Windows.Forms.Integration

Public Class MailDropWpfHostControl
    Inherits UserControl

        Private _host As ElementHost

    Public Sub New()
            _host = New ElementHost()
            _host.Dock = DockStyle.Fill
            _host.Child = New MailDropWpfTaskPane()
            Me.Controls.Add(_host)
    End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then
                _host?.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

End Class
