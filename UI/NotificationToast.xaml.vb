Imports System.Diagnostics
Imports System.IO
Imports System.Windows
Imports System.Windows.Interop
Imports System.Windows.Media
Imports System.Windows.Media.Animation

' Freistehendes Toast-Fenster fuer die Rueckmeldung nach einer erfolgreichen Ablage.
' Frueher lagen diese Meldungen als Border INNERHALB der Task Pane - dadurch musste die Pane
' so lange stehen bleiben, wie die Meldung sichtbar sein sollte (rund 4,4 s Totzeit einer
' 500 px breiten Sidebar, obwohl die Arbeit erledigt war). Als eigenes Fenster ueberlebt die
' Meldung das sofortige Ausblenden der Pane, sodass beides gleichzeitig moeglich ist:
' Platz sofort zurueckgeben und die Meldung trotzdem lange genug zeigen.
Public Class NotificationToast
    Inherits Window

    Public Enum ToastKind
        Success
        OverwriteWarning
        DuplicateWarning
    End Enum

    ' Muss zum Grid-Margin in NotificationToast.xaml passen: das Fenster ist per SizeToContent
    ' exakt so gross wie Border + Margin, der sichtbare Rand sitzt also um diesen Wert eingerueckt.
    Private Const ShadowMargin As Double = 14

    Private ReadOnly _ablageordnerPfad As String
    Private _hideTimer As System.Windows.Threading.DispatcherTimer

    Private Sub New(kind As ToastKind, message As String, ablageordnerPfad As String)
        InitializeComponent()
        _ablageordnerPfad = ablageordnerPfad
        ApplyKind(kind, message)
    End Sub

    ' Zeigt den Toast an der uebergebenen Bildschirmposition (in DIPs, nicht Geraetepixeln).
    ' screenTopLeft ist die linke obere Ecke der Task Pane; der Aufrufer muss sie ermitteln,
    ' BEVOR er die Pane ausblendet.
    Public Shared Function ShowToast(kind As ToastKind, message As String, ablageordnerPfad As String,
                                     screenTopLeft As Point, visibleDuration As TimeSpan) As NotificationToast
        Dim toast As New NotificationToast(kind, message, ablageordnerPfad)

        ' Outlooks Hauptfenster als Owner: der Toast bleibt dadurch ueber Outlook, minimiert mit
        ' Outlook und draengt sich nicht vor andere Anwendungen. Das erspart es, Z-Order und
        ' Minimieren selbst nachzubauen. Schlaegt es fehl, wird der Toast trotzdem angezeigt -
        ' nur ohne diese Kopplung.
        Try
            Dim outlookHandle = Process.GetCurrentProcess().MainWindowHandle
            If outlookHandle <> IntPtr.Zero Then
                Dim helper As New WindowInteropHelper(toast)
                helper.Owner = outlookHandle
            End If
        Catch ex As Exception
            Logger.LogError("NotificationToast: Owner-Fenster setzen", ex)
        End Try

        ' ShadowMargin wird auf BEIDEN Achsen herausgerechnet (die 14px sind ein Grid.Margin auf
        ' allen vier Seiten, siehe NotificationToast.xaml), damit die SICHTBARE linke und obere
        ' Kante des Toasts exakt auf der linken/oberen Kante der Sidebar liegt - unabhaengig von
        ' screenTopLeft, robust gegen Fensterposition/DPI. Fruehere Version rechnete den Rand nur
        ' vertikal heraus und liess den Toast horizontal 14px in die Sidebar hinein eingerueckt,
        ' um das Margin="14,..." der laengst entfernten Inline-Notification nachzubilden - das sah
        ' neben der buendigen Oberkante schief aus, da links und oben unterschiedlich behandelt
        ' wurden, ohne dass das noch einem noch sichtbaren Bezugspunkt entsprach.
        toast.Left = screenTopLeft.X - ShadowMargin
        toast.Top = screenTopLeft.Y - ShadowMargin

        toast.Show()
        toast.BeginAnimation(UIElement.OpacityProperty,
            New DoubleAnimation(0, 1, New Duration(TimeSpan.FromMilliseconds(250))))

        toast._hideTimer = New System.Windows.Threading.DispatcherTimer()
        toast._hideTimer.Interval = visibleDuration
        AddHandler toast._hideTimer.Tick, Sub(s, e)
                                              toast._hideTimer.Stop()
                                              toast.FadeOutAndClose()
                                          End Sub
        toast._hideTimer.Start()

        Return toast
    End Function

    Public Sub FadeOutAndClose()
        Dim fadeOut As New DoubleAnimation(Opacity, 0, New Duration(TimeSpan.FromMilliseconds(600)))
        AddHandler fadeOut.Completed, Sub(s, e) CloseSafely()
        BeginAnimation(UIElement.OpacityProperty, fadeOut)
    End Sub

    ' Mehrfach aufrufbar: der Aufrufer schliesst damit auch einen noch offenen aelteren Toast,
    ' ohne wissen zu muessen, ob dieser bereits von selbst zugegangen ist.
    Public Sub CloseSafely()
        Try
            _hideTimer?.Stop()
            Close()
        Catch ex As Exception
            Logger.LogError("NotificationToast.CloseSafely", ex)
        End Try
    End Sub

    Private Sub ApplyKind(kind As ToastKind, message As String)
        ' Glyphen bewusst per ChrW statt als Literal - diese Datei soll unabhaengig von
        ' Encoding-Eigenheiten bleiben (siehe CLAUDE.md zu gemischten Umlaut-Artefakten).
        Dim haken As String = ChrW(&H2713)
        Dim warnung As String = ChrW(&H26A0)
        Select Case kind
            Case ToastKind.Success
                StyleToast("#ECFDF5", "#6EE7B7", haken, "#059669",
                           "Erfolgreich abgelegt.", "#065F46", True, "#059669")
            Case ToastKind.OverwriteWarning
                StyleToast("#FFFBEB", "#FCD34D", warnung, "#B45309",
                           "Erfolgreich abgelegt. Existierende Dateien " & ChrW(&HFC) & "berschrieben.", "#92400E", True, "#92400E")
            Case Else
                StyleToast("#FFFBEB", "#FCD34D", warnung, "#B45309",
                           message, "#92400E", False, "#92400E")
        End Select
    End Sub

    Private Sub StyleToast(background As String, borderColor As String, icon As String, iconColor As String,
                           text As String, textColor As String, showOpenButton As Boolean, openButtonColor As String)
        ToastBorder.Background = HexBrush(background)
        ToastBorder.BorderBrush = HexBrush(borderColor)
        ToastIcon.Text = icon
        ToastIcon.Foreground = HexBrush(iconColor)
        ToastText.Text = text
        ToastText.Foreground = HexBrush(textColor)
        If showOpenButton AndAlso Not String.IsNullOrWhiteSpace(_ablageordnerPfad) Then
            ButtonOpenAblageordner.Foreground = HexBrush(openButtonColor)
            ButtonOpenAblageordner.Visibility = Visibility.Visible
        End If
    End Sub

    Private Shared Function HexBrush(hex As String) As Brush
        Return CType(New BrushConverter().ConvertFromString(hex), Brush)
    End Function

    Private Sub ButtonOpenAblageordner_Click(sender As Object, e As RoutedEventArgs)
        If String.IsNullOrWhiteSpace(_ablageordnerPfad) OrElse Not Directory.Exists(_ablageordnerPfad) Then Return
        Try
            Process.Start(_ablageordnerPfad)
            ' Der Toast hat seinen Zweck erfuellt, sobald der Ordner offen ist.
            CloseSafely()
        Catch ex As Exception
            Logger.LogError("NotificationToast.ButtonOpenAblageordner_Click", ex)
            MessageBox.Show("Ordner konnte nicht geoeffnet werden: " & ex.Message, "Oeffnen", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

End Class
