Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.IO

Public Class Session
    Implements INotifyPropertyChanged

    Private _selectedProjekt As String
    Private _titel As String
    Private _ablageordner As String
    Private _msgDateiname As String
    Private _anhaengeAblegen As Boolean
    Private _selectedMetadaten As New List(Of String)()
    Private _selectedOrdner As String
    Private _treeViewData As ObservableCollection(Of DirectoryNode)

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Protected Sub OnPropertyChanged(propertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub

    Public Property SelectedProjekt As String
        Get
            Return _selectedProjekt
        End Get
        Set(value As String)
            If _selectedProjekt <> value Then
                _selectedProjekt = value
                OnPropertyChanged(NameOf(SelectedProjekt))
            End If
        End Set
    End Property

    Public Property Titel As String
        Get
            Return _titel
        End Get
        Set(value As String)
            If _titel <> value Then
                _titel = value
                OnPropertyChanged(NameOf(Titel))
            End If
        End Set
    End Property

    Public Property Ablageordner As String
        Get
            Return _ablageordner
        End Get
        Set(value As String)
            If _ablageordner <> value Then
                _ablageordner = value
                OnPropertyChanged(NameOf(Ablageordner))
            End If
        End Set
    End Property

    Public Property MsgDateiname As String
        Get
            Return _msgDateiname
        End Get
        Set(value As String)
            If _msgDateiname <> value Then
                _msgDateiname = value
                OnPropertyChanged(NameOf(MsgDateiname))
            End If
        End Set
    End Property

    Public Property AnhaengeAblegen As Boolean
        Get
            Return _anhaengeAblegen
        End Get
        Set(value As Boolean)
            If _anhaengeAblegen <> value Then
                _anhaengeAblegen = value
                OnPropertyChanged(NameOf(AnhaengeAblegen))
            End If
        End Set
    End Property

    Public Property SelectedMetadaten As List(Of String)
        Get
            Return _selectedMetadaten
        End Get
        Set(value As List(Of String))
            If Not _selectedMetadaten.SequenceEqual(value) Then
                _selectedMetadaten = value
                OnPropertyChanged(NameOf(SelectedMetadaten))
            End If
        End Set
    End Property

    Public Property SelectedOrdner As String
        Get
            Return _selectedOrdner
        End Get
        Set(value As String)
            If _selectedOrdner <> value Then
                _selectedOrdner = value
                OnPropertyChanged(NameOf(SelectedOrdner))
            End If
        End Set
    End Property

    Public Property TreeViewData As ObservableCollection(Of DirectoryNode)
        Get
            Return _treeViewData
        End Get
        Set(value As ObservableCollection(Of DirectoryNode))
            If _treeViewData IsNot value Then
                _treeViewData = value
                OnPropertyChanged(NameOf(TreeViewData))
            End If
        End Set
    End Property

    Public Property ProjektVerzeichnisse As ObservableCollection(Of String) = New ObservableCollection(Of String) From {
        "P:\Leo\Projekte\MailDrop\TestDirectory\P-23001",
        "C:/Projekte/P-230111",
        "C:/Projekte/P-23251",
        "C:/Projekte/P-23052",
        "anderes..."
    }

    ' Setzt alle Properties auf Standardwerte zurück
    Public Sub Reset()
        SelectedProjekt = Nothing
        Titel = String.Empty
        Ablageordner = String.Empty
        MsgDateiname = String.Empty
        AnhaengeAblegen = False
        SelectedMetadaten.Clear()
        OnPropertyChanged(NameOf(SelectedMetadaten))
        SelectedOrdner = Nothing
        Debug.WriteLine("[Session] Reset ausgeführt")
    End Sub

    ' Vorbereitung der Session. Mail Daten auslesen und Felder aus- und vorausfüllen
    Public Sub PrepareSession()
        ReadMailMeta()
        TreeviewEngine()
        SchemaEngine()
        Debug.WriteLine("[Session] PrepareSession ausgeführt")
    End Sub

    ' Methoden für die Verarbeitungslogik
    Public Sub ReadMailMeta()
        Debug.WriteLine("[Session] ReadMailMeta ausgeführt")
        ' TODO: Implementiere das Auslesen der E-Mail-Metadaten
    End Sub

    ' Methode zum Bauen der Directory-Struktur
    Public Sub BuildDirectoryTree()
        If String.IsNullOrEmpty(SelectedProjekt) OrElse Not Directory.Exists(SelectedProjekt) Then
            TreeViewData = New ObservableCollection(Of DirectoryNode)()
            Return
        End If
        ' Nur die Unterordner von SelectedProjekt als Wurzelknoten verwenden
        Dim rootChildren As New ObservableCollection(Of DirectoryNode)()
        For Each dir As String In Directory.GetDirectories(SelectedProjekt)
            Dim childNode As DirectoryNode = CreateDirectoryNodeWithExpand(dir, 1)
            rootChildren.Add(childNode)
        Next
        TreeViewData = rootChildren
    End Sub

    ' level: 1 = erste Ebene unter SelectedProjekt
    Private Function CreateDirectoryNodeWithExpand(dirPath As String, level As Integer) As DirectoryNode
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
            Debug.WriteLine($"[Session] Zugriff verweigert auf {dirPath}")
        End Try
        Return node
    End Function

    ' TreeviewEngine ruft BuildDirectoryTree auf
    Public Sub TreeviewEngine()
        BuildDirectoryTree()
    End Sub

    Public Sub SchemaEngine()
        Debug.WriteLine("[Session] SchemaEngine ausgeführt")
        ' TODO: Implementiere die Logik für das Schema-Handling
    End Sub

    Public Sub SubmitSession()
        Debug.WriteLine("[Session] SubmitSession ausgeführt")
        ' TODO: Implementiere die Logik für das Abschicken/Speichern der Session
    End Sub

    Public Sub CancelSession()
        Debug.WriteLine("[Session] CancelSession ausgeführt")
        ' TODO: Implementiere die Logik für das Abbrechen der Session
    End Sub

    Public Sub HandleProjektSelection(selectedValue As String, uiContext As System.Windows.Window)
        If selectedValue = "anderes..." Then
            Dim dialog As New System.Windows.Forms.FolderBrowserDialog()
            dialog.Description = "Bitte Projektordner auswählen"
            If dialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                If Not ProjektVerzeichnisse.Contains(dialog.SelectedPath) Then
                    ProjektVerzeichnisse.Insert(0, dialog.SelectedPath)
                End If
                selectedValue = dialog.SelectedPath
            Else
                SelectedProjekt = Nothing
                BuildDirectoryTree()
                Return
            End If
        End If

        ' Prüfe, ob der Pfad existiert
        If Not String.IsNullOrEmpty(selectedValue) AndAlso Not Directory.Exists(selectedValue) Then
            System.Windows.MessageBox.Show($"Das Verzeichnis '{selectedValue}' existiert nicht!", "Pfad nicht gefunden", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning)
            SelectedProjekt = Nothing
        Else
            SelectedProjekt = selectedValue
        End If

        ' Treeview aktualisieren
        BuildDirectoryTree()
    End Sub
End Class

' Datenmodell für TreeView
Public Class DirectoryNode
    Public Property Name As String
    Public Property FullPath As String
    Public Property Children As ObservableCollection(Of DirectoryNode)
    Public Property IsExpanded As Boolean ' Für automatische Expansion im TreeView
End Class
