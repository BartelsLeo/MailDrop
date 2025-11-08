Imports System.Windows.Controls
Imports System.IO
Imports System.Windows
Imports System.Windows.Media
Imports System.Diagnostics

Public Class MailDropWpfTaskPane
    Inherits UserControl

    Public Property Session As New Session()
    Private infoPopup As InfoPopup = Nothing

    Public Sub New()
        InitializeComponent()
        Me.DataContext = Session
        AddHandler ListBox1.SelectionChanged, AddressOf ListBox1_SelectionChanged
        AddHandler Session.PropertyChanged, AddressOf Session_PropertyChanged
        AddHandler TreeView1.SelectedItemChanged, AddressOf TreeView1_SelectedItemChanged
    End Sub

    Private Sub ListBox1_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If ListBox1.SelectedItem IsNot Nothing Then
            Session.HandleProjektSelection(ListBox1.SelectedItem.ToString(), Window.GetWindow(Me))
        End If
    End Sub

    Private Sub Session_PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs)
        If e.PropertyName = NameOf(Session.ProjektPfad) Then
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
        Dim result As String = Session.ProcessSession()
        If Not String.IsNullOrEmpty(result) Then
            MessageBox.Show(result, "Fehler", MessageBoxButton.OK, MessageBoxImage.Error)
        Else
            MessageBox.Show("Vorgang erfolgreich abgeschlossen.", "Info", MessageBoxButton.OK, MessageBoxImage.Information)
        End If
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

    ' Setzt die Editierbarkeit der TaskPane
    Public Sub SetEditMode(isEditable As Boolean)
        ' Beispiel: Alle Controls auﬂer Info-Button deaktivieren/aktivieren
        For Each ctrl In Me.FindVisualChildren(Of Control)(Me)
            If ctrl.Name <> "ButtonInfo" Then
                ctrl.IsEnabled = isEditable
            End If
        Next
    End Sub

    ' Gibt True zur¸ck, wenn genau eine Mail selektiert ist, sonst False
    Public Function SingleMailSelected() As Boolean
        Dim explorer = Globals.ThisAddIn.Application.ActiveExplorer()
        Dim selection = explorer.Selection
        Return selection.Count = 1 AndAlso TypeOf selection.Item(1) Is Outlook.MailItem
    End Function

    ' Hilfsmethode: Findet alle Controls eines Typs rekursiv
    Private Iterator Function FindVisualChildren(Of T As DependencyObject)(depObj As DependencyObject) As IEnumerable(Of T)
        If depObj IsNot Nothing Then
            For i As Integer = 0 To VisualTreeHelper.GetChildrenCount(depObj) - 1
                Dim child = VisualTreeHelper.GetChild(depObj, i)
                If TypeOf child Is T Then
                    Yield CType(child, T)
                End If
                For Each childOfChild In FindVisualChildren(Of T)(child)
                    Yield childOfChild
                Next
            Next
        End If
    End Function

    Private Sub TreeView1_SelectedItemChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object))
        Dim node = TryCast(TreeView1.SelectedItem, DirectoryNode)
        If node IsNot Nothing Then
            Session.ProjektstrukturPfad = node.RelativePath
            Debug.WriteLine("ProjektstrukturPfad (relativ) gesetzt: " & node.RelativePath)
        End If
    End Sub
End Class
