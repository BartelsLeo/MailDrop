Imports System.Windows
Imports System.Windows.Controls

Public Class AttachmentRenameDialog
    Public Property FileName As String
    Private Const MaxPathLength As Integer = 255

    Public Sub New(currentName As String)
        InitializeComponent()
        TextBoxFileName.Text = currentName
        TextBoxFileName.SelectAll()
        TextBoxFileName.Focus()
        AddHandler TextBoxFileName.TextChanged, AddressOf TextBoxFileName_TextChanged
        UpdateOverlength()
    End Sub

    Private Sub ButtonOk_Click(sender As Object, e As RoutedEventArgs)
        FileName = TextBoxFileName.Text
        DialogResult = True
        Close()
    End Sub

    Private Sub TextBoxFileName_TextChanged(sender As Object, e As TextChangedEventArgs)
        UpdateOverlength()
    End Sub

    Private Sub UpdateOverlength()
        Dim length As Integer = TextBoxFileName.Text.Length
        Dim overlength As Integer = length - MaxPathLength
        If overlength > 0 Then
            TextBlockOverlength.Text = $"Überlänge von {overlength} Zeichen"
            TextBlockOverlength.Visibility = Visibility.Visible
        Else
            TextBlockOverlength.Text = String.Empty
            TextBlockOverlength.Visibility = Visibility.Collapsed
        End If
    End Sub
End Class
