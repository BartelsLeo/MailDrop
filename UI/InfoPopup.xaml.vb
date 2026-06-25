Imports System.Windows

Public Class InfoPopup
    Inherits Window
    Public Sub New()
        InitializeComponent()
        TxtDbPath.Text = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "MailDrop", "sessions.db")
    End Sub
End Class
