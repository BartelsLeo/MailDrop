Imports System.Windows
Imports System.Windows.Controls
Imports System.IO

Public Class AttachmentRenameDialog
    Public Property FileName As String
    Private Const MaxPathLength As Integer = 255
    Private ReadOnly _basePath As String

    Public Sub New(currentName As String, basePath As String)
        InitializeComponent()
        _basePath = basePath
        TextBoxFileName.Text = currentName
        TextBoxFileName.SelectAll()
        TextBoxFileName.Focus()
        AddHandler TextBoxFileName.TextChanged, AddressOf TextBoxFileName_TextChanged
        UpdateOverlength()
    End Sub

    Private Sub ButtonOk_Click(sender As Object, e As RoutedEventArgs)
        Dim name = TextBoxFileName.Text
        If String.IsNullOrWhiteSpace(name) OrElse name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 Then
            MessageBox.Show("Bitte einen g�ltigen Dateinamen eingeben.", "Ung�ltiger Name", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If
        FileName = name
        DialogResult = True
        Close()
    End Sub

    Private Sub TextBoxFileName_TextChanged(sender As Object, e As TextChangedEventArgs)
        UpdateOverlength()
    End Sub

    Private Sub UpdateOverlength()
        Dim fullPath = Path.Combine(_basePath, TextBoxFileName.Text)
        Dim overlength As Integer = fullPath.Length - MaxPathLength
        If overlength > 0 Then
            TextBlockOverlength.Text = $"�berl�nge von {overlength} Zeichen"
            TextBlockOverlength.Visibility = Visibility.Visible
        Else
            TextBlockOverlength.Text = String.Empty
            TextBlockOverlength.Visibility = Visibility.Collapsed
        End If
    End Sub
End Class
