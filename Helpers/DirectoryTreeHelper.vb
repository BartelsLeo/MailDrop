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

    ' Laedt nur die direkten Kinder von parentFullPath neu vom Dateisystem und ersetzt sie im
    ' uebergebenen Baum - statt (wie BuildDirectoryTree) den kompletten Baum neu einzulesen.
    ' Alles ausserhalb dieses Teilbaums (Geschwister von parentFullPath, dessen Vorfahren,
    ' andere Zweige) wird nicht angefasst und behaelt daher seinen Auf-/Zugeklappt-Zustand.
    ' Dieser eine Mechanismus deckt Erstellen, Loeschen und Umbenennen ab: alle drei Aktionen
    ' rufen nach der Dateisystem-Aenderung einfach RefreshChildren mit dem betroffenen
    ' Elternordner auf, statt selbst zu wissen, wie sich der Baum konkret veraendert hat -
    ' das ist robuster als eine Aktion-spezifische manuelle Node-Manipulation (z.B. bei
    ' Rename muesste man sonst Name/FullPath/RelativePath des Knotens und aller Nachfahren
    ' konsistent aktualisieren) und macht das TreeView nach der Aenderung wieder exakt
    ' konsistent mit dem tatsaechlichen Dateisystemzustand.
    ' Einziger Nachteil: Knoten UNTERHALB von parentFullPath (dessen (Enkel-)Kinder) werden
    ' komplett neu aufgebaut und verlieren dabei ihren eigenen Auf-/Zugeklappt-Zustand
    ' (Level<=2 relativ zu parentFullPath wird per CreateDirectoryNodeWithExpand neu gesetzt) -
    ' das betrifft aber nur den direkt bearbeiteten Ordner, nicht den Rest des Baums.
    ' parentFullPath ist entweder projektPfad selbst (Root-Ebene) oder der FullPath eines
    ' bereits vorhandenen Knotens. Gibt False zurueck, wenn der Elternknoten nicht gefunden
    ' wurde (Baum bleibt dann unveraendert - Aufrufer sollte auf BuildDirectoryTree() zurueckfallen).
    Public Function RefreshChildren(rootChildren As ObservableCollection(Of DirectoryNode), projektPfad As String, parentFullPath As String) As Boolean
        Dim normalizedParent = parentFullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        Dim normalizedProjekt = projektPfad.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)

        Dim targetChildren As ObservableCollection(Of DirectoryNode)
        Dim parentLevel As Integer

        If String.Equals(normalizedParent, normalizedProjekt, StringComparison.OrdinalIgnoreCase) Then
            targetChildren = rootChildren
            parentLevel = 0
        Else
            Dim parentNode = FindNodeByFullPath(rootChildren, normalizedParent)
            If parentNode Is Nothing Then
                ' Elternknoten nicht gefunden (sollte nicht vorkommen) - Baum unveraendert lassen.
                Return False
            End If
            parentNode.IsExpanded = True
            targetChildren = parentNode.Children
            parentLevel = parentNode.RelativePath.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar).Length
        End If

        targetChildren.Clear()
        If Directory.Exists(normalizedParent) Then
            For Each dir As String In Directory.GetDirectories(normalizedParent)
                targetChildren.Add(CreateDirectoryNodeWithExpand(dir, parentLevel + 1, projektPfad))
            Next
        End If

        Return True
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
