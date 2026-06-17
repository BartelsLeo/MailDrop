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
    Private _absender As String
    Private _absenderDomain As String
    Private _absenderKurz As String
    Private _empfaenger As String
    Private _betreff As String
    Private _betreffEmbedded As Single()
    Private _datum As DateTime
    Private _datumFormatiert As String
    Private _isSuggestedProjektPfad As Boolean
    Private _isSuggestedProjektstrukturPfad As Boolean
    Private _isSuggestedTitel As Boolean
    Private _isSuggestedAbsenderKurz As Boolean
    Private _isSuggestedAblageordnerSchema As Boolean
    Private _isSuggestedMsgDateinameSchema As Boolean
    Private _isSuggestedAnhaengeAblegen As Boolean

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
                BuildDirectoryTree()
                SuggestionEngineInstance?.RecalculateProjektPfadDistances(Me)
                Dim suggestedProjstr = SuggestionEngineInstance?.SuggestProjektstrukturPfad(Me)
                If Not String.IsNullOrWhiteSpace(suggestedProjstr) AndAlso
                   Not String.IsNullOrWhiteSpace(_projektPfad) AndAlso
                   IO.Directory.Exists(IO.Path.Combine(_projektPfad, suggestedProjstr)) Then
                    SuggestProjektstrukturPfad(suggestedProjstr)
                End If
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
                SuggestionEngineInstance?.RecalculateTitelDistances(Me)
                Dim suggestedAbsenderKurz = SuggestionEngineInstance?.SuggestAbsenderKurz(Me)
                If Not String.IsNullOrWhiteSpace(suggestedAbsenderKurz) Then
                    SuggestAbsenderKurz(suggestedAbsenderKurz)
                Else
                    Dim suggestedAbl = SuggestionEngineInstance?.SuggestAblageordnerSchema(Me)
                    If Not String.IsNullOrWhiteSpace(suggestedAbl) Then
                        SuggestAblageordnerSchema(suggestedAbl)
                    End If
                End If
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
                SuggestionEngineInstance?.RecalculateProjektstrukturPfadDistances(Me)
                Dim suggestedTitel = SuggestionEngineInstance?.SuggestTitel(Me)
                If Not String.IsNullOrWhiteSpace(suggestedTitel) Then
                    SuggestTitel(suggestedTitel)
                End If
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

    Private _projektVerzeichnisse As ObservableCollection(Of String)
    Public Property ProjektVerzeichnisse As ObservableCollection(Of String)
        Get
            Return _projektVerzeichnisse
        End Get
        Set(value As ObservableCollection(Of String))
            If _projektVerzeichnisse IsNot value Then
                _projektVerzeichnisse = value
                OnPropertyChanged(NameOf(ProjektVerzeichnisse))
            End If
        End Set
    End Property

    Public Property AblageordnerSchema As String
        Get
            Return _ablageordnerSchema
        End Get
        Set(value As String)
            If _ablageordnerSchema <> value Then
                _ablageordnerSchema = value
                OnPropertyChanged(NameOf(AblageordnerSchema))
                UpdateAblageordnerAufgeloest()
                AblageordnerFeld = AblageordnerAufgeloest
                Dim suggestedMsgDateinameSchema = SuggestionEngineInstance?.SuggestMsgDateinameSchema(Me)
                If Not String.IsNullOrWhiteSpace(suggestedMsgDateinameSchema) Then
                    SuggestMsgDateinameSchema(suggestedMsgDateinameSchema)
                End If
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
                MsgDateinameFeld = MsgDateinameAufgeloest
                Dim suggestedAnhaengeAblegen = SuggestionEngineInstance?.SuggestAnhaengeAblegen(Me)
                If suggestedAnhaengeAblegen.HasValue Then
                    SuggestAnhaengeAblegen(suggestedAnhaengeAblegen.Value)
                End If
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
        SuggestionEngineInstance?.RecalculateAblageordnerDistances(Me)
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
        If Not String.IsNullOrEmpty(Absender) Then
            result = result.Replace("[Absender]", Absender)
        End If
        If Not String.IsNullOrEmpty(AbsenderDomain) Then
            result = result.Replace("[Absender-Domain]", AbsenderDomain)
        End If
        If Not String.IsNullOrEmpty(Empfaenger) Then
            result = result.Replace("[Empf�nger]", Empfaenger)
        End If
        If Not String.IsNullOrEmpty(Empfaenger) Then
            result = result.Replace("[Empf�nger (kurz)]", Empfaenger)
        End If
        If Not String.IsNullOrEmpty(Betreff) Then
            result = result.Replace("[Betreff]", Betreff)
        End If
        If Datum <> Date.MinValue Then
            result = result.Replace("[Datum]", Datum.ToString("yyyy-MM-dd"))
        End If
        If Not String.IsNullOrEmpty(DatumFormatiert) Then
            result = result.Replace("[Datum (formatiert)]", DatumFormatiert)
        End If
        If Not String.IsNullOrEmpty(AbsenderKurz) Then
            result = result.Replace("[Absender (kurz)]", AbsenderKurz)
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
        IsSuggestedProjektPfad = False
        IsSuggestedProjektstrukturPfad = False
        IsSuggestedTitel = False
        IsSuggestedAbsenderKurz = False
        IsSuggestedAblageordnerSchema = False
        IsSuggestedMsgDateinameSchema = False
        IsSuggestedAnhaengeAblegen = False
        Debug.WriteLine("[Session] Reset ausgef�hrt")
    End Sub

    Public Property SuggestionEngineInstance As SuggestionEngine

    Public Property IsSuggestedProjektPfad As Boolean
        Get
            Return _isSuggestedProjektPfad
        End Get
        Set(value As Boolean)
            If _isSuggestedProjektPfad <> value Then
                _isSuggestedProjektPfad = value
                OnPropertyChanged(NameOf(IsSuggestedProjektPfad))
            End If
        End Set
    End Property

    Public Property IsSuggestedProjektstrukturPfad As Boolean
        Get
            Return _isSuggestedProjektstrukturPfad
        End Get
        Set(value As Boolean)
            If _isSuggestedProjektstrukturPfad <> value Then
                _isSuggestedProjektstrukturPfad = value
                OnPropertyChanged(NameOf(IsSuggestedProjektstrukturPfad))
            End If
        End Set
    End Property

    Public Property IsSuggestedTitel As Boolean
        Get
            Return _isSuggestedTitel
        End Get
        Set(value As Boolean)
            If _isSuggestedTitel <> value Then
                _isSuggestedTitel = value
                OnPropertyChanged(NameOf(IsSuggestedTitel))
            End If
        End Set
    End Property

    Public Property IsSuggestedAbsenderKurz As Boolean
        Get
            Return _isSuggestedAbsenderKurz
        End Get
        Set(value As Boolean)
            If _isSuggestedAbsenderKurz <> value Then
                _isSuggestedAbsenderKurz = value
                OnPropertyChanged(NameOf(IsSuggestedAbsenderKurz))
            End If
        End Set
    End Property

    Public Property IsSuggestedAblageordnerSchema As Boolean
        Get
            Return _isSuggestedAblageordnerSchema
        End Get
        Set(value As Boolean)
            If _isSuggestedAblageordnerSchema <> value Then
                _isSuggestedAblageordnerSchema = value
                OnPropertyChanged(NameOf(IsSuggestedAblageordnerSchema))
            End If
        End Set
    End Property

    Public Property IsSuggestedMsgDateinameSchema As Boolean
        Get
            Return _isSuggestedMsgDateinameSchema
        End Get
        Set(value As Boolean)
            If _isSuggestedMsgDateinameSchema <> value Then
                _isSuggestedMsgDateinameSchema = value
                OnPropertyChanged(NameOf(IsSuggestedMsgDateinameSchema))
            End If
        End Set
    End Property

    Public Property IsSuggestedAnhaengeAblegen As Boolean
        Get
            Return _isSuggestedAnhaengeAblegen
        End Get
        Set(value As Boolean)
            If _isSuggestedAnhaengeAblegen <> value Then
                _isSuggestedAnhaengeAblegen = value
                OnPropertyChanged(NameOf(IsSuggestedAnhaengeAblegen))
            End If
        End Set
    End Property

    Public Sub SuggestProjektPfad(value As String)
        ProjektPfad = value
        IsSuggestedProjektPfad = True
    End Sub

    Public Sub SuggestProjektstrukturPfad(value As String)
        ProjektstrukturPfad = value
        IsSuggestedProjektstrukturPfad = True
    End Sub

    Public Sub SuggestTitel(value As String)
        Titel = value
        IsSuggestedTitel = True
    End Sub

    Public Sub SuggestAbsenderKurz(value As String)
        AbsenderKurz = value
        IsSuggestedAbsenderKurz = True
    End Sub

    Public Sub SuggestAblageordnerSchema(value As String)
        AblageordnerSchema = value
        IsSuggestedAblageordnerSchema = True
    End Sub

    Public Sub SuggestMsgDateinameSchema(value As String)
        MsgDateinameSchema = value
        IsSuggestedMsgDateinameSchema = True
    End Sub

    Public Sub SuggestAnhaengeAblegen(value As Boolean)
        AnhaengeAblegen = value
        IsSuggestedAnhaengeAblegen = True
    End Sub

    Public Sub PrepareSession()
        Reset()
        MailUtils.ReadMailMeta(Me)
        ' Nach dem Einlesen der Mail: Projektverzeichnisse aktualisieren
        GetProjektVerzeichnisse()

        ' Shared SuggestionEngine nutzen (Historie wird pro Outlook-Start lazy geladen und gecacht).
        SuggestionEngineInstance = SuggestionEngine.GetSharedInstance()

        ' Alle Feature-Distanzlisten initialisieren: fixe Features berechnen, mutable Features auf 0 setzen.
        SuggestionEngineInstance.CalculateInitialFeatureDistances(Me)

        ' Vorhersage für Projektpfad aus den vorberechneten Distanzlisten berechnen.
        Dim suggestedProjektPfad = SuggestionEngineInstance.SuggestProjektPfad(Me)

        ' Gültige Vorschläge übernehmen und Cascade-Vorschläge auslösen.
        If Not String.IsNullOrWhiteSpace(suggestedProjektPfad) AndAlso Directory.Exists(suggestedProjektPfad) Then
            If ProjektVerzeichnisse IsNot Nothing AndAlso Not ProjektVerzeichnisse.Contains(suggestedProjektPfad) Then
                ProjektVerzeichnisse.Insert(0, suggestedProjektPfad)
            End If
            SuggestProjektPfad(suggestedProjektPfad)
        End If

        Debug.WriteLine("[Session] PrepareSession ausgef�hrt")
    End Sub

    ' Holt die letzten vier eindeutigen Projektverzeichnisse des aktuellen Benutzers aus der Datenbank
    Public Sub GetProjektVerzeichnisse()
        Dim verzeichnisse = ThisAddIn.CurrentDatabaseManager.GetLastProjektVerzeichnisseForUser(Me.AusfueBenutzer)
        ' Maximal 3 aus der DB, letzter Eintrag immer "anderes..."
        Dim list As New List(Of String)(verzeichnisse.Take(3))
        list.Add("anderes...")
        ProjektVerzeichnisse = New ObservableCollection(Of String)(list)
        Debug.WriteLine($"[Session] ProjektVerzeichnisse f�r Benutzer '{Me.AusfueBenutzer}': {String.Join(", ", list)}")
    End Sub

    Public Sub BuildDirectoryTree()
        TreeViewData = DirectoryTreeHelper.BuildDirectoryTree(ProjektPfad)
    End Sub

    Public Sub CancelSession()
        Debug.WriteLine("[Session] CancelSession ausgef�hrt")
        ' TODO: Implementiere die Logik f�r das Abbrechen der Session
    End Sub

    Public Sub HandleProjektSelection(selectedValue As String, uiContext As System.Windows.Window)
        If selectedValue = "anderes..." Then
            Dim dialog As New System.Windows.Forms.FolderBrowserDialog()
            dialog.Description = "Bitte Projektordner ausw�hlen"
            If dialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                If Not ProjektVerzeichnisse.Contains(dialog.SelectedPath) Then
                    ProjektVerzeichnisse.Insert(0, dialog.SelectedPath)
                End If
                selectedValue = dialog.SelectedPath
            Else
                ProjektPfad = Nothing
                Return
            End If
        End If
        If Not String.IsNullOrEmpty(selectedValue) AndAlso Not Directory.Exists(selectedValue) Then
            System.Windows.MessageBox.Show($"Das Verzeichnis '{selectedValue}' existiert nicht!", "Pfad nicht gefunden", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning)
            ProjektPfad = Nothing
        Else
            ProjektPfad = selectedValue
        End If
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
        ' PredictionEngine.TrainDecisionTreeModels()
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
        Get
            Return _absender
        End Get
        Set(value As String)
            If _absender <> value Then
                _absender = value
                OnPropertyChanged(NameOf(Absender))
            End If
        End Set
    End Property

    <DisplayName("Absender-Domain")>
    Public Property AbsenderDomain As String
        Get
            Return _absenderDomain
        End Get
        Set(value As String)
            If _absenderDomain <> value Then
                _absenderDomain = value
                OnPropertyChanged(NameOf(AbsenderDomain))
            End If
        End Set
    End Property

    <DisplayName("AbsenderKurz")>
    Public Property AbsenderKurz As String
        Get
            Return _absenderKurz
        End Get
        Set(value As String)
            If _absenderKurz <> value Then
                _absenderKurz = value
                OnPropertyChanged(NameOf(AbsenderKurz))
                UpdateResolvedAfterTitelChange()
                Dim suggestedAbl = SuggestionEngineInstance?.SuggestAblageordnerSchema(Me)
                If Not String.IsNullOrWhiteSpace(suggestedAbl) Then
                    SuggestAblageordnerSchema(suggestedAbl)
                End If
            End If
        End Set
    End Property

    <DisplayName("Empf�nger")>
    Public Property Empfaenger As String
        Get
            Return _empfaenger
        End Get
        Set(value As String)
            If _empfaenger <> value Then
                _empfaenger = value
                OnPropertyChanged(NameOf(Empfaenger))
            End If
        End Set
    End Property

    <DisplayName("Betreff")>
    Public Property Betreff As String
        Get
            Return _betreff
        End Get
        Set(value As String)
            If _betreff <> value Then
                _betreff = value
                OnPropertyChanged(NameOf(Betreff))
            End If
        End Set
    End Property

    <DisplayName("Datum")>
    Public Property BetreffEmbedded As Single()
        Get
            Return _betreffEmbedded
        End Get
        Set(value As Single())
            _betreffEmbedded = value
            OnPropertyChanged(NameOf(BetreffEmbedded))
        End Set
    End Property

    <DisplayName("Betreff Embedded")>
    Public Property Datum As DateTime
        Get
            Return _datum
        End Get
        Set(value As DateTime)
            If _datum <> value Then
                _datum = value
                OnPropertyChanged(NameOf(Datum))
            End If
        End Set
    End Property

    <DisplayName("Datum (formatiert)")>
    Public Property DatumFormatiert As String
        Get
            Return _datumFormatiert
        End Get
        Set(value As String)
            If _datumFormatiert <> value Then
                _datumFormatiert = value
                OnPropertyChanged(NameOf(DatumFormatiert))
            End If
        End Set
    End Property

    Public Property ID As Integer

    ' Wandelt das aktuelle Session-Objekt in ein SessionRecord-Objekt um,
    ' sodass es f�r die Speicherung in der Datenbank oder f�r maschinelles Lernen verwendet werden kann.
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
