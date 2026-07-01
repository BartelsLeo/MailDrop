Imports System.Threading.Tasks
Imports System.Windows

Public Class InfoPopup
    Inherits Window
    Public Sub New()
        InitializeComponent()
        TxtDbPath.Text = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "MailDrop", "sessions.db")
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
