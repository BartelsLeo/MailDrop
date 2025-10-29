Imports System.Windows.Controls
Imports System.IO
Imports System.Windows

Public Class MailDropWpfTaskPane
    Inherits UserControl

    Public Property Session As New Session()

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
End Class
