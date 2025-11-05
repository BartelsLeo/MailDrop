Imports System.Collections.ObjectModel
Imports System.IO

Public Module DirectoryTreeHelper
    ' Erstellt die Directory-Struktur für das TreeView
    Public Function BuildDirectoryTree(selectedProjekt As String) As ObservableCollection(Of DirectoryNode)
        If String.IsNullOrEmpty(selectedProjekt) OrElse Not Directory.Exists(selectedProjekt) Then
            Return New ObservableCollection(Of DirectoryNode)()
        End If
        Dim rootChildren As New ObservableCollection(Of DirectoryNode)()
        For Each dir As String In Directory.GetDirectories(selectedProjekt)
            Dim childNode As DirectoryNode = CreateDirectoryNodeWithExpand(dir, 1)
            rootChildren.Add(childNode)
        Next
        Return rootChildren
    End Function

    ' level: 1 = erste Ebene unter SelectedProjekt
    Public Function CreateDirectoryNodeWithExpand(dirPath As String, level As Integer) As DirectoryNode
        Dim node As New DirectoryNode With {
            .Name = Path.GetFileName(dirPath),
            .FullPath = dirPath,
            .Children = New ObservableCollection(Of DirectoryNode)(),
            .IsExpanded = (level <= 2)
        }
        Try
            For Each dir As String In Directory.GetDirectories(dirPath)
                node.Children.Add(CreateDirectoryNodeWithExpand(dir, level + 1))
            Next
        Catch ex As Exception
            ' Fehlerausgabe entfernt, da Debug nicht verfügbar ist
        End Try
        Return node
    End Function
End Module

Public Class DirectoryNode
    Public Property Name As String
    Public Property FullPath As String
    Public Property Children As ObservableCollection(Of DirectoryNode)
    Public Property IsExpanded As Boolean ' Für automatische Expansion im TreeView
End Class
