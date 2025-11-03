Imports System.Windows.Controls
Imports System.IO
Imports System.Windows

Public Class MailDropWpfTaskPane
    Inherits UserControl

    Public Property Session As New Session()
    Private infoPopup As InfoPopup = Nothing

    Public Sub New()
        InitializeComponent()
        Me.DataContext = Session
        AddHandler ListBox1.SelectionChanged, AddressOf ListBox1_SelectionChanged
        AddHandler Session.PropertyChanged, AddressOf Session_PropertyChanged
    End Sub

    Private Sub ListBox1_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If ListBox1.SelectedItem IsNot Nothing Then
            Session.HandleProjektSelection(ListBox1.SelectedItem.ToString(), Window.GetWindow(Me))
        End If
    End Sub

    Private Sub Session_PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs)
        If e.PropertyName = NameOf(Session.SelectedProjekt) Then
            Session.TreeviewEngine()
        End If
    End Sub

    Private Sub TextBoxAblageordner_GotFocus(sender As Object, e As RoutedEventArgs)
        Session.BeginAblageordnerEdit()
    End Sub

    Private Sub TextBoxAblageordner_LostFocus(sender As Object, e As RoutedEventArgs)
        Session.EndAblageordnerEdit()
    End Sub

    Private Sub TextBoxMsgDateiname_GotFocus(sender As Object, e As RoutedEventArgs)
        Session.BeginMsgDateinameEdit()
    End Sub

    Private Sub TextBoxMsgDateiname_LostFocus(sender As Object, e As RoutedEventArgs)
        Session.EndMsgDateinameEdit()
    End Sub

    Private Sub ButtonInfo_Click(sender As Object, e As RoutedEventArgs)
        If infoPopup Is Nothing OrElse Not infoPopup.IsLoaded Then
            infoPopup = New InfoPopup()
            infoPopup.Owner = Window.GetWindow(Me)
            infoPopup.WindowStartupLocation = WindowStartupLocation.CenterOwner
            infoPopup.Show()
        Else
            infoPopup.Activate()
        End If
    End Sub

    Private Sub ButtonOk_Click(sender As Object, e As RoutedEventArgs)
        ' Hier kann ggf. eine Speichern- oder ‹bernehmen-Logik erg‰nzt werden
        ' Aktuell keine spezielle Aktion
    End Sub

    Private Sub ButtonAbbrechen_Click(sender As Object, e As RoutedEventArgs)
        ' Schlieﬂe InfoPopup, falls offen
        If infoPopup IsNot Nothing AndAlso infoPopup.IsLoaded Then
            infoPopup.Close()
            infoPopup = Nothing
        End If
        Try
            Globals.ThisAddIn.HideTaskPane()
        Catch
            ' Fallback: Fenster schlieﬂen
            Dim wnd = Window.GetWindow(Me)
            If wnd IsNot Nothing Then wnd.Close()
        End Try
    End Sub
End Class
