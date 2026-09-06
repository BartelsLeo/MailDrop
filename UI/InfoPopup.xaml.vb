Imports System.Deployment.Application
Imports System.Diagnostics
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Navigation

Public Class InfoPopup
    Inherits Window
    Public Sub New()
        InitializeComponent()
        TxtDbPath.Text = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "MailDrop", "sessions.db")
        TxtVersion.Text = $"Version {GetDisplayVersion()}"
    End Sub

    ' ClickOnce-ApplicationVersion bei Netzwerkbereitstellung, sonst statische AssemblyVersion (lokaler Debug-Start).
    Private Function GetDisplayVersion() As String
        If ApplicationDeployment.IsNetworkDeployed Then
            Return ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString()
        End If
        Return Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()
    End Function

    Private Sub Hyperlink_RequestNavigate(sender As Object, e As RequestNavigateEventArgs)
        Try
            Process.Start(New ProcessStartInfo(e.Uri.AbsoluteUri) With {.UseShellExecute = True})
        Catch ex As Exception
            Logger.LogError("InfoPopup.Hyperlink_RequestNavigate", ex)
        End Try
        e.Handled = True
    End Sub

    Private Sub ButtonGewichteNeuBerechnen_Click(sender As Object, e As RoutedEventArgs)
        ButtonGewichteNeuBerechnen.IsEnabled = False
        ButtonGewichteNeuBerechnen.Content = "Berechnung läuft…"
        Task.Run(Sub()
                     Try
                         SuggestionEngine.GetSharedInstance().RecalculateWeightsFromHistory()
                     Catch ex As Exception
                         System.Diagnostics.Debug.WriteLine("[InfoPopup] Gewichte-Neuberechnung fehlgeschlagen: " & ex.Message)
                     End Try
                     Dispatcher.Invoke(Sub()
                                           ButtonGewichteNeuBerechnen.Content = "Gewichte neu berechnen"
                                           ButtonGewichteNeuBerechnen.IsEnabled = True
                                       End Sub)
                 End Sub)
    End Sub
End Class
