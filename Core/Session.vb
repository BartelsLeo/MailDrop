Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.IO
Imports System.Threading.Tasks

Public Class AttachmentItem
    Implements INotifyPropertyChanged

    Public Property Name As String
    Public Property OutlookIndex As Integer

    Private _isSelected As Boolean = True
    Public Property IsSelected As Boolean
        Get
            Return _isSelected
        End Get
        Set(value As Boolean)
            If _isSelected <> value Then
                _isSelected = value
                RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(IsSelected)))
            End If
        End Set
    End Property

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
End Class

Public Class Session
    Implements INotifyPropertyChanged

    Public Property LastDuplicateWarning As String
    Public Property LastOverwriteWarning As String
    Public Property Anhaenge As New ObservableCollection(Of AttachmentItem)()

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

    Private Const DefaultAblageSchema As String = "[Datum (formatiert)]_[Absender (kurz)]_[Titel]"

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
                SuggestionEngineInstance?.RunSuggestionCascade(Me, SuggestionEngine.CascadeStep.ProjektstrukturPfad)
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
                SuggestionEngineInstance?.RunSuggestionCascade(Me, SuggestionEngine.CascadeStep.AbsenderKurz)
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

    Public ReadOnly Property HasAnhaenge As Boolean
        Get
            Return Anhaenge.Count > 0
        End Get
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
                SuggestionEngineInstance?.RunSuggestionCascade(Me, SuggestionEngine.CascadeStep.Titel)
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
                SuggestionEngineInstance?.RunSuggestionCascade(Me, SuggestionEngine.CascadeStep.MsgDateinameSchema)
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
                SuggestionEngineInstance?.RunSuggestionCascade(Me, SuggestionEngine.CascadeStep.AnhaengeAblegen)
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

        Dim placeholderValues As New Dictionary(Of String, String) From {
            {"[Titel]", If(Titel, String.Empty)},
            {"[Absender]", If(Absender, String.Empty)},
            {"[Absender-Domain]", If(AbsenderDomain, String.Empty)},
            {"[Empf�nger]", If(Empfaenger, String.Empty)},
            {"[Empf�nger (kurz)]", If(Empfaenger, String.Empty)},
            {"[Betreff]", If(Betreff, String.Empty)},
            {"[Datum]", If(Datum <> Date.MinValue, Datum.ToString("yyyy-MM-dd"), String.Empty)},
            {"[Datum (formatiert)]", If(DatumFormatiert, String.Empty)},
            {"[Absender (kurz)]", If(AbsenderKurz, String.Empty)}
        }

        Dim pattern = String.Join("|", placeholderValues.Keys.Select(Function(k) Text.RegularExpressions.Regex.Escape(k)))
        Dim tokenRegex As New Text.RegularExpressions.Regex(pattern)

        ' Vorlage in wechselnde literale/Platzhalter-Segmente zerlegen: literals(0), Wert(0),
        ' literals(1), Wert(1), ..., literals(n). isEmptyPlaceholder(i) = Nothing markiert ein
        ' literales Segment (z.B. ein vom User frei gewaehltes Trennzeichen wie "_"), sonst
        ' True/False je nachdem ob der Platzhalter an dieser Stelle leer aufgeloest wurde.
        Dim literals = tokenRegex.Split(template)
        Dim matches = tokenRegex.Matches(template)

        Dim parts As New List(Of String)
        Dim isEmptyPlaceholder As New List(Of Boolean?)

        parts.Add(literals(0))
        isEmptyPlaceholder.Add(Nothing)
        For i = 0 To matches.Count - 1
            Dim value = placeholderValues(matches(i).Value)
            parts.Add(value)
            isEmptyPlaceholder.Add(String.IsNullOrEmpty(value))
            parts.Add(literals(i + 1))
            isEmptyPlaceholder.Add(Nothing)
        Next

        ' "Skip-empty-join": ein literales Segment, das direkt zwischen zwei Platzhaltern liegt
        ' (ein "Connector", z.B. das vom User getippte "_"), wird nicht sofort ausgegeben,
        ' sondern zwischengespeichert (pendingConnector) und erst unmittelbar vor dem naechsten
        ' NICHT leeren Platzhalter ausgegeben - und auch nur, wenn diesem bereits ein anderer,
        ' nicht leerer Platzhalter vorausging. Ein leerer Platzhalter selbst wird uebersprungen,
        ' ohne den zwischengespeicherten Connector zu verwerfen (der naechste Connector
        ' ueberschreibt ihn einfach). Dadurch bleibt genau EIN Trennzeichen zwischen zwei
        ' tatsaechlich befuellten Platzhaltern erhalten, egal wie viele leere Platzhalter dazwischen
        ' liegen, und es entsteht nie ein fuehrendes/abschliessendes Trennzeichen. Literaler Text,
        ' der nur an einen Platzhalter grenzt (z.B. ein fixes Prefix vor dem ersten Platzhalter),
        ' ist kein Connector und wird immer unveraendert uebernommen, auch wenn dieser Platzhalter leer ist.
        Dim result As New Text.StringBuilder()
        Dim pendingConnector As String = Nothing
        Dim emittedAny = False

        For i = 0 To parts.Count - 1
            If isEmptyPlaceholder(i) Is Nothing Then
                Dim isConnector = i > 0 AndAlso i < parts.Count - 1 AndAlso isEmptyPlaceholder(i - 1).HasValue AndAlso isEmptyPlaceholder(i + 1).HasValue
                If isConnector Then
                    pendingConnector = parts(i)
                Else
                    result.Append(parts(i))
                End If
            ElseIf Not String.IsNullOrEmpty(parts(i)) Then
                If emittedAny AndAlso pendingConnector IsNot Nothing Then
                    result.Append(pendingConnector)
                End If
                result.Append(parts(i))
                emittedAny = True
                pendingConnector = Nothing
            End If
        Next

        Return result.ToString()
    End Function

    Public Sub UpdateResolvedAfterTitelChange()
        UpdateAblageordnerAufgeloest()
        AblageordnerFeld = AblageordnerAufgeloest
        UpdateMsgDateinameAufgeloest()
        MsgDateinameFeld = MsgDateinameAufgeloest
    End Sub

    Public Sub Reset()
        LastDuplicateWarning = String.Empty
        Anhaenge.Clear()
        OnPropertyChanged(NameOf(HasAnhaenge))
        ProjektPfad = Nothing
        Titel = String.Empty
        AblageordnerSchema = String.Empty
        _ablageordnerAufgeloest = String.Empty
        AblageordnerFeld = String.Empty
        MsgDateinameSchema = String.Empty
        _msgDateinameAufgeloest = String.Empty
        MsgDateinameFeld = String.Empty
        AnhaengeAblegen = True
        ProjektstrukturPfad = Nothing
        IsSuggestedProjektPfad = False
        IsSuggestedProjektstrukturPfad = False
        IsSuggestedTitel = False
        IsSuggestedAbsenderKurz = False
        IsSuggestedAblageordnerSchema = False
        IsSuggestedMsgDateinameSchema = False
        IsSuggestedAnhaengeAblegen = False
        ' AbsenderKurz is not set by ReadMailMeta — clear backing field directly to avoid
        ' the setter's cascade guard (If _absenderKurz <> value) suppressing the next suggestion.
        _absenderKurz = String.Empty
        OnPropertyChanged(NameOf(AbsenderKurz))
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
        Dim sw As Stopwatch = Stopwatch.StartNew()
        Dim t As Long
        Debug.WriteLine("[Session] PrepareSession BEGIN")

        Reset()
        Debug.WriteLine($"[Session]   Reset:                 {sw.ElapsedMilliseconds} ms")

        MailUtils.ReadMailMeta(Me)
        t = sw.ElapsedMilliseconds : Debug.WriteLine($"[Session]   ReadMailMeta:          {t} ms")

        MailUtils.ReadAttachmentNames(Me)
        Debug.WriteLine($"[Session]   ReadAttachmentNames:   {sw.ElapsedMilliseconds - t} ms ({Anhaenge.Count} attachments)") : t = sw.ElapsedMilliseconds
        OnPropertyChanged(NameOf(HasAnhaenge))

        ' Nach dem Einlesen der Mail: Projektverzeichnisse aktualisieren
        GetProjektVerzeichnisse()
        Debug.WriteLine($"[Session]   GetProjektVerzeichnisse:{sw.ElapsedMilliseconds - t} ms") : t = sw.ElapsedMilliseconds

        ' Shared SuggestionEngine nutzen (Historie wird pro Outlook-Start lazy geladen und gecacht).
        SuggestionEngineInstance = SuggestionEngine.GetSharedInstance()
        Debug.WriteLine($"[Session]   GetSharedInstance:     {sw.ElapsedMilliseconds - t} ms") : t = sw.ElapsedMilliseconds

        ' Alle Feature-Distanzlisten initialisieren: fixe Features berechnen, mutable Features auf 0 setzen.
        SuggestionEngineInstance.CalculateInitialFeatureDistances(Me)
        Debug.WriteLine($"[Session]   CalculateInitialFeatureDistances: {sw.ElapsedMilliseconds - t} ms") : t = sw.ElapsedMilliseconds

        ' Schema-Standardwerte direkt in Backing-Fields schreiben, um keinen Cascade auszulösen.
        ' DatumFormatiert und Absender sind zu diesem Zeitpunkt bereits gesetzt, sodass
        ' [Datum (formatiert)] sofort aufgelöst wird; Titel und AbsenderKurz werden durch
        ' den nachfolgenden Cascade-Lauf nachgefüllt.
        _ablageordnerSchema = DefaultAblageSchema
        OnPropertyChanged(NameOf(AblageordnerSchema))
        UpdateAblageordnerAufgeloest()
        AblageordnerFeld = AblageordnerAufgeloest
        _msgDateinameSchema = DefaultAblageSchema
        OnPropertyChanged(NameOf(MsgDateinameSchema))
        UpdateMsgDateinameAufgeloest()
        MsgDateinameFeld = MsgDateinameAufgeloest
        Debug.WriteLine($"[Session]   Schema-Defaults:        {sw.ElapsedMilliseconds - t} ms") : t = sw.ElapsedMilliseconds

        ' Vorhersage für Projektpfad aus den vorberechneten Distanzlisten berechnen.
        Dim suggestedProjektPfad = SuggestionEngineInstance.SuggestProjektPfad(Me)
        Debug.WriteLine($"[Session]   SuggestProjektPfad:    {sw.ElapsedMilliseconds - t} ms → '{suggestedProjektPfad}'") : t = sw.ElapsedMilliseconds

        ' Gültige Vorschläge übernehmen und Cascade-Vorschläge auslösen.
        ' SuggestProjektPfad garantiert bereits, dass der Pfad existiert.
        If String.IsNullOrWhiteSpace(suggestedProjektPfad) Then
            Debug.WriteLine("[Session]   ProjektPfad: kein Vorschlag vom Engine")
        Else
            If ProjektVerzeichnisse IsNot Nothing AndAlso Not ProjektVerzeichnisse.Contains(suggestedProjektPfad) Then
                ProjektVerzeichnisse.Insert(0, suggestedProjektPfad)
            End If
            SuggestProjektPfad(suggestedProjektPfad)
        End If

        ' Sicherheitsnetz: Mails ohne Anhänge immer mit deaktivierter Checkbox abschließen,
        ' unabhängig davon ob der Cascade einen Vorschlag gesetzt hat.
        If Not HasAnhaenge Then
            AnhaengeAblegen = False
            IsSuggestedAnhaengeAblegen = False
        End If
        Debug.WriteLine($"[Session]   Cascade+Suggest:       {sw.ElapsedMilliseconds - t} ms")

        Debug.WriteLine($"[Session] PrepareSession END – total: {sw.ElapsedMilliseconds} ms")
    End Sub

    ' Holt die letzten vier eindeutigen Projektverzeichnisse des aktuellen Benutzers aus der Datenbank
    Public Sub GetProjektVerzeichnisse()
        Dim verzeichnisse = ThisAddIn.CurrentDatabaseManager.GetLastProjektVerzeichnisseForUser(Me.AusfueBenutzer)
        ' Maximal 10 aus der DB, letzter Eintrag immer "anderes..."
        Dim list As New List(Of String)(verzeichnisse.Take(10))
        list.Add("anderes...")
        ProjektVerzeichnisse = New ObservableCollection(Of String)(list)
        Debug.WriteLine($"[Session] ProjektVerzeichnisse f�r Benutzer '{Me.AusfueBenutzer}': {String.Join(", ", list)}")
    End Sub

    Public Sub BuildDirectoryTree()
        TreeViewData = DirectoryTreeHelper.BuildDirectoryTree(ProjektPfad)
    End Sub

    ' Laedt nur die Kinder von parentFullPath (Root-Ebene = ProjektPfad, oder ein bestehender
    ' Knoten) neu vom Dateisystem, statt den kompletten TreeViewData-Baum neu aufzubauen
    ' (siehe DirectoryTreeHelper.RefreshChildren). Gibt False zurueck, wenn parentFullPath im
    ' aktuellen Baum nicht gefunden wurde - der Aufrufer sollte dann auf BuildDirectoryTree() zurueckfallen.
    Public Function RefreshTreeViewChildren(parentFullPath As String) As Boolean
        Return DirectoryTreeHelper.RefreshChildren(TreeViewData, ProjektPfad, parentFullPath)
    End Function

    Public Sub CancelSession()
        Reset()
        Debug.WriteLine("[Session] CancelSession ausgef�hrt")
    End Sub

    Public Sub HandleProjektSelection(selectedValue As String)
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
        LastDuplicateWarning = String.Empty
        LastOverwriteWarning = String.Empty
        Dim checkedInput = InputChecker.CheckInput(Me)
        If checkedInput.ErrorMessage <> String.Empty Then
            Return checkedInput.ErrorMessage
        End If
        LastDuplicateWarning = checkedInput.DuplicateWarning
        Dim ablageResult = Me.CreateAblageordner(checkedInput.CheckedAblageOrdner)
        If ablageResult <> String.Empty Then
            Return ablageResult
        End If
        ' Prüfen ob Dateien bereits existieren; Überschreiben wird durchgeführt.
        Dim msgPfad = checkedInput.CheckedMsgZielpfad
        If Not msgPfad.ToLower().EndsWith(".msg") Then msgPfad &= ".msg"
        Dim overwriteFound = File.Exists(msgPfad) OrElse
                             checkedInput.CheckedAnhZielpfade.Any(Function(p) File.Exists(p))
        If overwriteFound Then
            LastOverwriteWarning = "Erfolgreich abgelegt. Existierende Dateien überschrieben."
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
        Dim newRecord = Me.ToSessionRecord()
        ThisAddIn.CurrentDatabaseManager.SaveSessionRecord(newRecord)
        SuggestionEngine.GetSharedInstance().AppendHistoricalRecord(newRecord)
        Dim recordCount As Integer = ThisAddIn.CurrentDatabaseManager.GetSessionRecordCount()
        If recordCount Mod 50 = 0 Then
            Task.Run(Sub() SuggestionEngine.GetSharedInstance().RecalculateWeightsFromHistory())
        End If
        ' Warnungswerte vor Reset() sichern, da Reset() sie löscht.
        Dim savedDuplicate = LastDuplicateWarning
        Dim savedOverwrite = LastOverwriteWarning
        Me.Reset()
        LastDuplicateWarning = savedDuplicate
        LastOverwriteWarning = savedOverwrite
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
                SuggestionEngineInstance?.RunSuggestionCascade(Me, SuggestionEngine.CascadeStep.AblageordnerSchema)
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
