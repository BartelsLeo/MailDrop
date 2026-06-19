Imports System.Collections.ObjectModel
Imports System.IO

Public Module DirectoryTreeHelper
    ' Erstellt die Directory-Struktur f�r das TreeView
    Public Function BuildDirectoryTree(projektPfad As String) As ObservableCollection(Of DirectoryNode)
        If String.IsNullOrEmpty(projektPfad) OrElse Not Directory.Exists(projektPfad) Then
            Return New ObservableCollection(Of DirectoryNode)()
        End If
        Dim rootChildren As New ObservableCollection(Of DirectoryNode)()
        For Each dir As String In Directory.GetDirectories(projektPfad)
            Dim childNode As DirectoryNode = CreateDirectoryNodeWithExpand(dir, 1, projektPfad)
            rootChildren.Add(childNode)
        Next
        Return rootChildren
    End Function

    ' level: 1 = erste Ebene unter ProjektPfad
    Public Function CreateDirectoryNodeWithExpand(dirPath As String, level As Integer, basePath As String, Optional maxDepth As Integer = 20) As DirectoryNode
        Dim relPath = If(dirPath.StartsWith(basePath), dirPath.Substring(basePath.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar), dirPath)
        Dim node As New DirectoryNode With {
            .Name = Path.GetFileName(dirPath),
            .FullPath = dirPath,
            .RelativePath = relPath,
            .Children = New ObservableCollection(Of DirectoryNode)(),
            .IsExpanded = (level <= 2)
        }
        If level > maxDepth Then Return node
        Try
            For Each dir As String In Directory.GetDirectories(dirPath)
                node.Children.Add(CreateDirectoryNodeWithExpand(dir, level + 1, basePath, maxDepth))
            Next
        Catch ex As Exception
            ' Fehlerausgabe entfernt, da Debug nicht verf�gbar ist
        End Try
        Return node
    End Function
End Module

Public Class DirectoryNode
    Public Property Name As String
    Public Property FullPath As String
    Public Property RelativePath As String
    Public Property Children As ObservableCollection(Of DirectoryNode)
    Public Property IsExpanded As Boolean ' F�r automatische Expansion im TreeView
End Class
