Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.IO

Public Class Session
    Implements INotifyPropertyChanged

    Private _selectedProjekt As String
    Private _titel As String
    Private _anhaengeAblegen As Boolean
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
                ' Nach Änderung des Titels Resolved- und Feld-Properties aktualisieren
                UpdateResolvedAfterTitelChange()
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

    ' Ablageordner mit Platzhaltern
    Private _ablageordnerTemplate As String
    Public Property AblageordnerTemplate As String
        Get
            Return _ablageordnerTemplate
        End Get
        Set(value As String)
            If _ablageordnerTemplate <> value Then
                _ablageordnerTemplate = value
                OnPropertyChanged(NameOf(AblageordnerTemplate))
                UpdateAblageordnerResolved()
            End If
        End Set
    End Property

    ' Ablageordner mit ausgewerteten Platzhaltern (readonly)
    Private _ablageordnerResolved As String
    Public ReadOnly Property AblageordnerResolved As String
        Get
            Return _ablageordnerResolved
        End Get
    End Property

    ' MsgDateiname mit Platzhaltern
    Private _msgDateinameTemplate As String
    Public Property MsgDateinameTemplate As String
        Get
            Return _msgDateinameTemplate
        End Get
        Set(value As String)
            If _msgDateinameTemplate <> value Then
                _msgDateinameTemplate = value
                OnPropertyChanged(NameOf(MsgDateinameTemplate))
                UpdateMsgDateinameResolved()
            End If
        End Set
    End Property

    ' MsgDateiname mit ausgewerteten Platzhaltern (readonly)
    Private _msgDateinameResolved As String
    Public ReadOnly Property MsgDateinameResolved As String
        Get
            Return _msgDateinameResolved
        End Get
    End Property

    ' Methode zum Ersetzen der Platzhalter im Ablageordner
    Private Sub UpdateAblageordnerResolved()
        _ablageordnerResolved = ReplacePlaceholders(_ablageordnerTemplate)
        OnPropertyChanged(NameOf(AblageordnerResolved))
    End Sub

    ' Methode zum Ersetzen der Platzhalter im MsgDateiname
    Private Sub UpdateMsgDateinameResolved()
        _msgDateinameResolved = ReplacePlaceholders(_msgDateinameTemplate)
        OnPropertyChanged(NameOf(MsgDateinameResolved))
    End Sub

    ' Hilfsmethode zum Ersetzen von Platzhaltern mit Friendly Names
    Private Function ReplacePlaceholders(template As String) As String
        If String.IsNullOrEmpty(template) Then Return String.Empty
        Dim result = template
        ' Titel als Platzhalter ersetzen
        If Not String.IsNullOrEmpty(Titel) Then
            result = result.Replace("[Titel]", Titel)
        End If
        ' AbsenderKurz als Platzhalter ersetzen
        If Not String.IsNullOrEmpty(AbsenderKurz) Then
            result = result.Replace("[Absender (kurz)]", AbsenderKurz)
        End If
        Return result
    End Function

    ' Methode zum Aktualisieren der Resolved- und Feld-Properties nach Titeländerung
    Public Sub UpdateResolvedAfterTitelChange()
        UpdateAblageordnerResolved()
        AblageordnerFeld = AblageordnerResolved
        UpdateMsgDateinameResolved()
        MsgDateinameFeld = MsgDateinameResolved
    End Sub

    ' AbsenderKurz analog zu Titel
    Private _absenderKurz As String
    Public Property AbsenderKurz As String
        Get
            Return _absenderKurz
        End Get
        Set(value As String)
            If _absenderKurz <> value Then
                _absenderKurz = value
                OnPropertyChanged(NameOf(AbsenderKurz))
                UpdateResolvedAfterAbsenderKurzChange()
            End If
        End Set
    End Property

    ' Methode zum Aktualisieren der Resolved- und Feld-Properties nach AbsenderKurz-Änderung
    Public Sub UpdateResolvedAfterAbsenderKurzChange()
        UpdateAblageordnerResolved()
        AblageordnerFeld = AblageordnerResolved
        UpdateMsgDateinameResolved()
        MsgDateinameFeld = MsgDateinameResolved
    End Sub

    ' Setzt alle Properties auf Standardwerte zurück
    Public Sub Reset()
        SelectedProjekt = Nothing
        Titel = String.Empty
        AblageordnerTemplate = String.Empty
        _ablageordnerResolved = String.Empty
        AblageordnerFeld = String.Empty
        MsgDateinameTemplate = String.Empty
        _msgDateinameResolved = String.Empty
        MsgDateinameFeld = String.Empty
        AnhaengeAblegen = False
        SelectedOrdner = Nothing
        Debug.WriteLine("[Session] Reset ausgeführt")
    End Sub

    ' Vorbereitung der Session. Mail Daten auslesen und Felder aus- und vorausfüllen
    Public Sub PrepareSession()
        Reset()
        MailUtils.ReadMailMeta()
        TreeviewEngine()
        Debug.WriteLine("[Session] PrepareSession ausgeführt")
    End Sub

    ' Methode zum Bauen der Directory-Struktur
    Public Sub BuildDirectoryTree()
        TreeViewData = DirectoryTreeHelper.BuildDirectoryTree(SelectedProjekt)
    End Sub

    ' TreeviewEngine ruft BuildDirectoryTree auf
    Public Sub TreeviewEngine()
        BuildDirectoryTree()
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

    ' Ablageordner Feld für das UI
    Private _ablageordnerFeld As String
    Public Property AblageordnerFeld As String
        Get
            Return _ablageordnerFeld
        End Get
        Set(value As String)
            If _ablageordnerFeld <> value Then
                _ablageordnerFeld = value
                OnPropertyChanged(NameOf(AblageordnerFeld))
            End If
        End Set
    End Property

    ' MsgDateiname Feld für das UI
    Private _msgDateinameFeld As String
    Public Property MsgDateinameFeld As String
        Get
            Return _msgDateinameFeld
        End Get
        Set(value As String)
            If _msgDateinameFeld <> value Then
                _msgDateinameFeld = value
                OnPropertyChanged(NameOf(MsgDateinameFeld))
            End If
        End Set
    End Property

    ' Methoden für Focus Handling
    Public Sub BeginAblageordnerEdit()
        AblageordnerFeld = AblageordnerTemplate
    End Sub

    Public Sub EndAblageordnerEdit()
        AblageordnerTemplate = AblageordnerFeld
        UpdateAblageordnerResolved()
        AblageordnerFeld = AblageordnerResolved
    End Sub

    Public Sub BeginMsgDateinameEdit()
        MsgDateinameFeld = MsgDateinameTemplate
    End Sub

    Public Sub EndMsgDateinameEdit()
        MsgDateinameTemplate = MsgDateinameFeld
        UpdateMsgDateinameResolved()
        MsgDateinameFeld = MsgDateinameResolved
    End Sub

    ' Prüft die Session-Eingaben und führt die Ablage durch
    Public Function ProcessSession() As String
        Dim checkedInput = InputChecker.CheckInput(Me)
        If checkedInput.ErrorMessage <> String.Empty Then
            Return checkedInput.ErrorMessage
        End If
        ' 1. Ablageordner erstellen
        Dim ablageResult = Me.CreateAblageordner(checkedInput.CheckedAblageOrdner)
        If ablageResult <> String.Empty Then
            Return ablageResult
        End If
        ' 2. Mail als msg speichern
        Dim mailResult = MailUtils.SaveSelectedMailAsMsg(checkedInput.CheckedMsgZielpfad)
        If mailResult <> String.Empty Then
            Return mailResult
        End If
        ' 3. Anhänge speichern (nur wenn aktiviert)
        If AnhaengeAblegen Then
            Dim anhangResult = MailUtils.SaveMailAttachments(checkedInput.CheckedAnhZielpfade)
            If anhangResult <> String.Empty Then
                Return anhangResult
            End If
        End If
        ' Nach erfolgreichem Abschluss alles zurücksetzen
        Me.Reset()
        Return String.Empty
    End Function

    ' Erstellt den Ablageordner, falls nicht vorhanden
    Private Function CreateAblageordner(ablageOrdnerPfad As String) As String
        Try
            If Not Directory.Exists(ablageOrdnerPfad) Then
                Directory.CreateDirectory(ablageOrdnerPfad)
            End If
            Return String.Empty
        Catch ex As Exception
            Return $"Fehler beim Erstellen des Ablageordners: {ex.Message}"
        End Try
    End Function

End Class
