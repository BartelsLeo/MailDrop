Imports System.ComponentModel
Imports System.IO
Imports System.Windows
Imports System.Windows.Controls

Public Enum SaveItemStatus
    Ok
    Conflict
    Skipped
    Saved
End Enum

Public Class SaveItem
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _status As SaveItemStatus
    Private _fullPath As String
    Private _displayName As String

    Public Property IsFolder As Boolean
    Public Property SaveAction As Action(Of SaveItem)

    Public Property DisplayName As String
        Get
            Return _displayName
        End Get
        Set(value As String)
            _displayName = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(DisplayName)))
        End Set
    End Property

    Public Property FullPath As String
        Get
            Return _fullPath
        End Get
        Set(value As String)
            _fullPath = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(FullPath)))
        End Set
    End Property

    Public Property Status As SaveItemStatus
        Get
            Return _status
        End Get
        Set(value As SaveItemStatus)
            _status = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(Status)))
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(IsOk)))
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(IsConflict)))
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(IsSkipped)))
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(IsSaved)))
        End Set
    End Property

    Public ReadOnly Property IsOk As Boolean
        Get
            Return _status = SaveItemStatus.Ok
        End Get
    End Property

    Public ReadOnly Property IsConflict As Boolean
        Get
            Return _status = SaveItemStatus.Conflict
        End Get
    End Property

    Public ReadOnly Property IsSkipped As Boolean
        Get
            Return _status = SaveItemStatus.Skipped
        End Get
    End Property

    Public ReadOnly Property IsSaved As Boolean
        Get
            Return _status = SaveItemStatus.Saved
        End Get
    End Property
End Class

Public Class SavePreviewDialog
    Private ReadOnly _items As List(Of SaveItem)

    Public Sub New(items As List(Of SaveItem))
        InitializeComponent()
        _items = items
        ItemsListView.ItemsSource = _items
        If System.Windows.Application.Current IsNot Nothing AndAlso
           System.Windows.Application.Current.MainWindow IsNot Nothing Then
            Owner = System.Windows.Application.Current.MainWindow
        End If
    End Sub

    Private Sub EditButton_Click(sender As Object, e As RoutedEventArgs)
        Dim item = DirectCast(DirectCast(sender, Button).Tag, SaveItem)
        ResolveFileConflict(item)
        CheckAndAutoSave()
    End Sub

    Private Sub ResolveFileConflict(item As SaveItem)
        Dim result = MessageBox.Show(
            $"Die Datei ""{item.DisplayName}"" existiert bereits.{vbCrLf}" &
            "Ja = Überschreiben     Nein = Umbenennen     Abbrechen = Überspringen",
            "Datei existiert bereits",
            MessageBoxButton.YesNoCancel,
            MessageBoxImage.Question)

        Select Case result
            Case MessageBoxResult.Yes
                item.Status = SaveItemStatus.Ok

            Case MessageBoxResult.No
                Dim newName = InputChecker.ShowAttachmentRenameDialog(
                    item.DisplayName,
                    Path.GetDirectoryName(item.FullPath),
                    "Bitte geben Sie einen neuen Dateinamen ein:")
                If Not String.IsNullOrEmpty(newName) Then
                    item.FullPath = Path.Combine(Path.GetDirectoryName(item.FullPath), newName)
                    item.DisplayName = newName
                    item.Status = SaveItemStatus.Ok
                Else
                    item.Status = SaveItemStatus.Skipped
                End If

            Case Else
                item.Status = SaveItemStatus.Skipped
        End Select
    End Sub

    Private Sub CheckAndAutoSave()
        If _items.All(Function(i) i.Status = SaveItemStatus.Ok OrElse i.Status = SaveItemStatus.Skipped) Then
            ExecuteSave()
        End If
    End Sub

    Friend Sub ExecuteSave()
        For Each item In _items
            If item.Status = SaveItemStatus.Skipped Then Continue For

            If item.IsFolder Then
                If Not TrySaveFolder(item) Then Return
            Else
                TrySaveFile(item)
            End If
        Next
        DialogResult = True
        Close()
    End Sub

    Private Function TrySaveFolder(item As SaveItem) As Boolean
        Try
            Directory.CreateDirectory(item.FullPath)
            item.Status = SaveItemStatus.Saved
            Return True
        Catch ex As Exception
            Dim dlgResult = MessageBox.Show(
                $"Der Ordner ""{item.DisplayName}"" konnte nicht erstellt werden:{vbCrLf}{ex.Message}{vbCrLf}{vbCrLf}Soll ein anderer Ordnername vergeben werden?",
                "Ordner nicht erstellbar",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning)
            If dlgResult <> MessageBoxResult.Yes Then
                DialogResult = False
                Close()
                Return False
            End If

            Dim parentPath = Path.GetDirectoryName(item.FullPath)
            Dim newName = InputChecker.ShowAttachmentRenameDialog(
                item.DisplayName, parentPath,
                "Bitte geben Sie einen neuen Ordnernamen ein:", "Ordner umbenennen")
            If String.IsNullOrEmpty(newName) Then
                DialogResult = False
                Close()
                Return False
            End If

            Dim newFolderPath = Path.Combine(parentPath, newName)
            Try
                Directory.CreateDirectory(newFolderPath)
            Catch ex2 As Exception
                MessageBox.Show($"Fehler beim Erstellen des Ordners: {ex2.Message}", "Fehler",
                                MessageBoxButton.OK, MessageBoxImage.Error)
                DialogResult = False
                Close()
                Return False
            End Try

            ' Update folder item and cascade paths for all file items
            item.FullPath = newFolderPath
            item.DisplayName = newName
            item.Status = SaveItemStatus.Saved
            For Each fileItem In _items.Where(Function(i) Not i.IsFolder)
                fileItem.FullPath = Path.Combine(newFolderPath, Path.GetFileName(fileItem.FullPath))
            Next
            Return True
        End Try
    End Function

    Private Sub TrySaveFile(item As SaveItem)
        Try
            item.SaveAction(item)
            item.Status = SaveItemStatus.Saved
        Catch ex As Exception
            Dim newName = InputChecker.ShowAttachmentRenameDialog(
                item.DisplayName,
                Path.GetDirectoryName(item.FullPath),
                $"'{item.DisplayName}' konnte nicht gespeichert werden:{vbCrLf}{ex.Message}{vbCrLf}Neuen Namen eingeben:",
                "Datei nicht speicherbar")
            If String.IsNullOrEmpty(newName) Then
                item.Status = SaveItemStatus.Skipped
                Return
            End If
            item.FullPath = Path.Combine(Path.GetDirectoryName(item.FullPath), newName)
            item.DisplayName = newName
            Try
                item.SaveAction(item)
                item.Status = SaveItemStatus.Saved
            Catch ex2 As Exception
                item.Status = SaveItemStatus.Skipped
            End Try
        End Try
    End Sub

    Private Sub ButtonOk_Click(sender As Object, e As RoutedEventArgs)
        If _items.Any(Function(i) i.Status = SaveItemStatus.Conflict) Then
            MessageBox.Show("Bitte lösen Sie zuerst alle Konflikte auf.",
                            "Konflikte ausstehend", MessageBoxButton.OK, MessageBoxImage.Information)
            Return
        End If
        ExecuteSave()
    End Sub

    Private Sub ButtonAbbrechen_Click(sender As Object, e As RoutedEventArgs)
        DialogResult = False
        Close()
    End Sub
End Class
