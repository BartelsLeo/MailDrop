Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.IO

Public Class Session
    Implements INotifyPropertyChanged

    Private _projektPfad As String
    Private _titel As String
    Private _anhaengeAblegen As Boolean
    Private _projektstrukturPfad As String
    Private _treeViewData As ObservableCollection(Of DirectoryNode)
    Private _ablageordnerSchema As String
    Private _ablageordnerAufgeloest As String
    Private _msgDateinameSchema As String
    Private _msgDateinameAufgeloest As String
    Private _ablageordnerFeld As String
    Private _msgDateinameFeld As String
    Private _ausfueDatum As DateTime = DateTime.Now
    Private _ausfueBenutzer As String = Environment.UserName

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Protected Sub OnPropertyChanged(propertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub

    Public Property ProjektPfad As String
        Get
            Return _projektPfad
        End Get
        Set(value As String)
            If _projektPfad <> value Then
                _projektPfad = value
                OnPropertyChanged(NameOf(ProjektPfad))
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

    Public Property ProjektstrukturPfad As String
        Get
            Return _projektstrukturPfad
        End Get
        Set(value As String)
            If _projektstrukturPfad <> value Then
                _projektstrukturPfad = value
                OnPropertyChanged(NameOf(ProjektstrukturPfad))
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

    Public Property AblageordnerSchema As String
        Get
            Return _ablageordnerSchema
        End Get
        Set(value As String)
            If _ablageordnerSchema <> value Then
                _ablageordnerSchema = value
                OnPropertyChanged(NameOf(AblageordnerSchema))
                UpdateAblageordnerAufgeloest()
            End If
        End Set
    End Property

    Public ReadOnly Property AblageordnerAufgeloest As String
        Get
            Return _ablageordnerAufgeloest
        End Get
    End Property

    Public Property MsgDateinameSchema As String
        Get
            Return _msgDateinameSchema
        End Get
        Set(value As String)
            If _msgDateinameSchema <> value Then
                _msgDateinameSchema = value
                OnPropertyChanged(NameOf(MsgDateinameSchema))
                UpdateMsgDateinameAufgeloest()
            End If
        End Set
    End Property

    Public ReadOnly Property MsgDateinameAufgeloest As String
        Get
            Return _msgDateinameAufgeloest
        End Get
    End Property

    Private Sub UpdateAblageordnerAufgeloest()
        _ablageordnerAufgeloest = ReplacePlaceholders(_ablageordnerSchema)
        OnPropertyChanged(NameOf(AblageordnerAufgeloest))
    End Sub

    Private Sub UpdateMsgDateinameAufgeloest()
        _msgDateinameAufgeloest = ReplacePlaceholders(_msgDateinameSchema)
        OnPropertyChanged(NameOf(MsgDateinameAufgeloest))
    End Sub

    Private Function ReplacePlaceholders(template As String) As String
        If String.IsNullOrEmpty(template) Then Return String.Empty
        Dim result = template
        If Not String.IsNullOrEmpty(Titel) Then
            result = result.Replace("[Titel]", Titel)
        End If
        Return result
    End Function

    Public Sub UpdateResolvedAfterTitelChange()
        UpdateAblageordnerAufgeloest()
        AblageordnerFeld = AblageordnerAufgeloest
        UpdateMsgDateinameAufgeloest()
        MsgDateinameFeld = MsgDateinameAufgeloest
    End Sub

    Public Sub Reset()
        ProjektPfad = Nothing
        Titel = String.Empty
        AblageordnerSchema = String.Empty
        _ablageordnerAufgeloest = String.Empty
        AblageordnerFeld = String.Empty
        MsgDateinameSchema = String.Empty
        _msgDateinameAufgeloest = String.Empty
        MsgDateinameFeld = String.Empty
        AnhaengeAblegen = False
        ProjektstrukturPfad = Nothing
        Debug.WriteLine("[Session] Reset ausgeführt")
    End Sub

    Public Sub PrepareSession()
        Reset()
        MailUtils.ReadMailMeta(Me)
        TreeviewEngine()
        Debug.WriteLine("[Session] PrepareSession ausgeführt")
    End Sub

    Public Sub BuildDirectoryTree()
        TreeViewData = DirectoryTreeHelper.BuildDirectoryTree(ProjektPfad)
    End Sub

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
                ProjektPfad = Nothing
                BuildDirectoryTree()
                Return
            End If
        End If
        If Not String.IsNullOrEmpty(selectedValue) AndAlso Not Directory.Exists(selectedValue) Then
            System.Windows.MessageBox.Show($"Das Verzeichnis '{selectedValue}' existiert nicht!", "Pfad nicht gefunden", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning)
            ProjektPfad = Nothing
        Else
            ProjektPfad = selectedValue
        End If
        BuildDirectoryTree()
    End Sub

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

    Public Sub BeginAblageordnerEdit()
        AblageordnerFeld = AblageordnerSchema
    End Sub

    Public Sub EndAblageordnerEdit()
        AblageordnerSchema = AblageordnerFeld
        UpdateAblageordnerAufgeloest()
        AblageordnerFeld = AblageordnerAufgeloest
    End Sub

    Public Sub BeginMsgDateinameEdit()
        MsgDateinameFeld = MsgDateinameSchema
    End Sub

    Public Sub EndMsgDateinameEdit()
        MsgDateinameSchema = MsgDateinameFeld
        UpdateMsgDateinameAufgeloest()
        MsgDateinameFeld = MsgDateinameAufgeloest
    End Sub

    Public Function ProcessSession() As String
        Dim checkedInput = InputChecker.CheckInput(Me)
        If checkedInput.ErrorMessage <> String.Empty Then
            Return checkedInput.ErrorMessage
        End If
        Dim ablageResult = Me.CreateAblageordner(checkedInput.CheckedAblageOrdner)
        If ablageResult <> String.Empty Then
            Return ablageResult
        End If
        Dim mailResult = MailUtils.SaveSelectedMailAsMsg(checkedInput.CheckedMsgZielpfad)
        If mailResult <> String.Empty Then
            Return mailResult
        End If
        If AnhaengeAblegen Then
            Dim anhangResult = MailUtils.SaveMailAttachments(checkedInput.CheckedAnhZielpfade)
            If anhangResult <> String.Empty Then
                Return anhangResult
            End If
        End If
        ThisAddIn.CurrentDatabaseManager.SaveSessionRecord(Me.ToSessionRecord())
        PredictionEngine.EncodeSessions()
        Me.Reset()
        Return String.Empty
    End Function

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

    Public Property AusfueDatum As DateTime
        Get
            Return _ausfueDatum
        End Get
        Set(value As DateTime)
            If _ausfueDatum <> value Then
                _ausfueDatum = value
                OnPropertyChanged(NameOf(AusfueDatum))
            End If
        End Set
    End Property

    Public Property AusfueBenutzer As String
        Get
            Return _ausfueBenutzer
        End Get
        Set(value As String)
            If _ausfueBenutzer <> value Then
                _ausfueBenutzer = value
                OnPropertyChanged(NameOf(AusfueBenutzer))
            End If
        End Set
    End Property

    <DisplayName("Absender")>
    Public Property Absender As String
    <DisplayName("Absender-Domain")>
    Public Property AbsenderDomain As String
    <DisplayName("AbsenderKurz")>
    Public Property AbsenderKurz As String
    <DisplayName("Empfänger")>
    Public Property Empfaenger As String
    <DisplayName("Betreff")>
    Public Property Betreff As String
    <DisplayName("Datum")>
    Public Property Datum As DateTime
    <DisplayName("Datum (formatiert)")>
    Public Property DatumFormatiert As String

    Public Property ID As Integer

    Public Function ToSessionRecord() As SessionRecord
        Dim record As New SessionRecord()
        For Each prop In GetType(SessionRecord).GetProperties()
            Dim sessionProp = Me.GetType().GetProperty(prop.Name)
            If sessionProp IsNot Nothing Then
                Dim value = sessionProp.GetValue(Me)
                prop.SetValue(record, value)
            End If
        Next
        Return record
    End Function
End Class
