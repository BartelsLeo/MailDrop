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

    ' Fuegt einen neu angelegten, leeren Ordner als einzelnen Knoten in einen bestehenden
    ' TreeView-Baum ein, statt den kompletten Baum per BuildDirectoryTree() neu vom
    ' Dateisystem einzulesen. Das vermeidet sowohl die Laufzeit eines vollstaendigen
    ' Neuaufbaus als auch den Verlust des vom User manuell auf-/zugeklappten Zustands
    ' (BuildDirectoryTree setzt IsExpanded pauschal fuer die ersten zwei Ebenen neu).
    ' parentFullPath ist entweder projektPfad selbst (Root-Ebene) oder der FullPath eines
    ' bereits vorhandenen Knotens.
    Public Function InsertDirectoryNode(rootChildren As ObservableCollection(Of DirectoryNode), projektPfad As String, parentFullPath As String, newFolderPath As String) As DirectoryNode
        Dim newNode As New DirectoryNode With {
            .Name = Path.GetFileName(newFolderPath),
            .FullPath = newFolderPath,
            .RelativePath = If(newFolderPath.StartsWith(projektPfad), newFolderPath.Substring(projektPfad.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar), newFolderPath),
            .Children = New ObservableCollection(Of DirectoryNode)(),
            .IsExpanded = False
        }

        Dim normalizedParent = parentFullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        Dim normalizedProjekt = projektPfad.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)

        Dim targetChildren = rootChildren
        If Not String.Equals(normalizedParent, normalizedProjekt, StringComparison.OrdinalIgnoreCase) Then
            Dim parentNode = FindNodeByFullPath(rootChildren, normalizedParent)
            If parentNode Is Nothing Then
                ' Elternknoten nicht gefunden (sollte nicht vorkommen) - Baum unveraendert lassen.
                Return Nothing
            End If
            parentNode.IsExpanded = True
            targetChildren = parentNode.Children
        End If

        Dim insertIndex = 0
        While insertIndex < targetChildren.Count AndAlso String.Compare(targetChildren(insertIndex).Name, newNode.Name, StringComparison.OrdinalIgnoreCase) < 0
            insertIndex += 1
        End While
        targetChildren.Insert(insertIndex, newNode)

        Return newNode
    End Function

    Private Function FindNodeByFullPath(children As ObservableCollection(Of DirectoryNode), fullPath As String) As DirectoryNode
        For Each child In children
            If String.Equals(child.FullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar), fullPath, StringComparison.OrdinalIgnoreCase) Then
                Return child
            End If
            Dim found = FindNodeByFullPath(child.Children, fullPath)
            If found IsNot Nothing Then Return found
        Next
        Return Nothing
    End Function
End Module

Public Class DirectoryNode
    Public Property Name As String
    Public Property FullPath As String
    Public Property RelativePath As String
    Public Property Children As ObservableCollection(Of DirectoryNode)
    Public Property IsExpanded As Boolean ' F�r automatische Expansion im TreeView
End Class
