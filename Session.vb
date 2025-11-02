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
    Private _mailMetaInfo As New MailMetaInfo()

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

    Public Property MailMetaInfo As MailMetaInfo
        Get
            Return _mailMetaInfo
        End Get
        Set(value As MailMetaInfo)
            _mailMetaInfo = value
            OnPropertyChanged(NameOf(MailMetaInfo))
            OnPropertyChanged(NameOf(MailMetaInfoList))
        End Set
    End Property

    Public ReadOnly Property MailMetaInfoList As List(Of KeyValuePair(Of String, String))
        Get
            Dim result As New List(Of KeyValuePair(Of String, String))()
            If MailMetaInfo Is Nothing Then Return result
            Dim props = GetType(MailMetaInfo).GetProperties()
            For Each prop In props
                Dim displayNameAttr = CType(Attribute.GetCustomAttribute(prop, GetType(DisplayNameAttribute)), DisplayNameAttribute)
                Dim key As String = If(displayNameAttr IsNot Nothing, displayNameAttr.DisplayName, prop.Name)
                Dim valueObj = prop.GetValue(MailMetaInfo)
                Dim valueStr = If(valueObj IsNot Nothing, valueObj.ToString(), String.Empty)
                result.Add(New KeyValuePair(Of String, String)(key, valueStr))
            Next
            Return result
        End Get
    End Property

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
        If MailMetaInfo IsNot Nothing Then
            Dim props = GetType(MailMetaInfo).GetProperties()
            For Each prop In props
                Dim displayNameAttr = CType(Attribute.GetCustomAttribute(prop, GetType(DisplayNameAttribute)), DisplayNameAttribute)
                If displayNameAttr IsNot Nothing Then
                    Dim friendlyName = displayNameAttr.DisplayName
                    Dim valueObj = prop.GetValue(MailMetaInfo)
                    Dim valueStr = If(valueObj IsNot Nothing, valueObj.ToString(), String.Empty)
                    result = result.Replace($"[{friendlyName}]", valueStr)
                End If
            Next
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
        MailMetaInfo = New MailMetaInfo()
        SelectedOrdner = Nothing
        Debug.WriteLine("[Session] Reset ausgeführt")
    End Sub

    ' Vorbereitung der Session. Mail Daten auslesen und Felder aus- und vorausfüllen
    Public Sub PrepareSession()
        Reset()
        ReadMailMeta()
        TreeviewEngine()
        SchemaEngine()
        Debug.WriteLine("[Session] PrepareSession ausgeführt")
    End Sub

    ' Methoden für die Verarbeitungslogik
    Public Sub ReadMailMeta()
        Debug.WriteLine("[Session] ReadMailMeta ausgeführt")
        Try
            ' Verwende die VSTO-Instanz für Outlook
            Dim app As Outlook.Application = Globals.ThisAddIn.Application
            Dim explorer = app.ActiveExplorer()
            If explorer Is Nothing OrElse explorer.Selection.Count <> 1 Then
                System.Windows.MessageBox.Show("Eine Mail auswählen", "Warnung", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning)
                Return
            End If
            Dim mail = TryCast(explorer.Selection.Item(1), Outlook.MailItem)
            If mail Is Nothing Then
                System.Windows.MessageBox.Show("Bitte eine einzelne E-Mail auswählen.", "Warnung", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning)
                Return
            End If
            Dim info As New MailMetaInfo()
            info.Sender = mail.SenderName
            If mail.SenderEmailType = "SMTP" AndAlso mail.SenderEmailAddress.Contains("@") Then
                info.SenderDomain = mail.SenderEmailAddress.Split("@"c).Last()
            End If
            info.Empfaenger = mail.To
            If Not String.IsNullOrEmpty(mail.To) Then
                Dim firstTo = mail.To.Split({";"}, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()
                If Not String.IsNullOrEmpty(firstTo) Then
                    info.EmpfaengerKurz = firstTo.Split("@"c)(0).Trim()
                End If
            End If
            info.Betreff = mail.Subject
            info.Datum = mail.ReceivedTime
            info.DatumFormatiert = mail.ReceivedTime.ToString("yyyyMMdd")
            MailMetaInfo = info
        Catch ex As Exception
            System.Windows.MessageBox.Show($"Fehler beim Auslesen der E-Mail: {ex.Message}", "Fehler", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error)
        End Try
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
End Class

' Datenmodell für TreeView
Public Class DirectoryNode
    Public Property Name As String
    Public Property FullPath As String
    Public Property Children As ObservableCollection(Of DirectoryNode)
    Public Property IsExpanded As Boolean ' Für automatische Expansion im TreeView
End Class

Public Class MailMetaInfo
    <DisplayName("Absender")>
    Public Property Sender As String
    <DisplayName("Absender-Domain")>
    Public Property SenderDomain As String
    <DisplayName("Empfänger")>
    Public Property Empfaenger As String
    <DisplayName("Empfänger (kurz)")>
    Public Property EmpfaengerKurz As String
    <DisplayName("Betreff")>
    Public Property Betreff As String
    <DisplayName("Datum")>
    Public Property Datum As DateTime
    <DisplayName("Datum (formatiert)")>
    Public Property DatumFormatiert As String
End Class
