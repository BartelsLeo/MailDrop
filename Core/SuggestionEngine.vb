Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks

Public Class SuggestionEngine
    Implements IDisposable

    Private Shared ReadOnly SharedEngineLazy As New Lazy(Of SuggestionEngine)(
        Function()
            Debug.WriteLine("[SuggestionEngine] Shared instance is being created.")
            Return New SuggestionEngine()
        End Function,
        LazyThreadSafetyMode.ExecutionAndPublication)

    Public Property EnginesHistoricalSessionRecords As List(Of SessionRecord)
    Private EnginesEmbeddingService As EmbeddingService

    ' Diese Listen werden pro aktueller Session einmal berechnet und anschließend beim Suggest genutzt.
    Private Property BetreffDistances As List(Of Double)
    Private Property DatumsDistances As List(Of Double)
    Private Property AbsenderDomainDistances As List(Of Double)
    Private Property AbsenderDistances As List(Of Double)
    Private Property AusfueBenutzerDistances As List(Of Double)
    Private Property AusfueDatumsDistances As List(Of Double)
    Private Property TitelDistances As List(Of Double)
    Private Property AblageordnerDistances As List(Of Double)
    Private Property ProjektPfadDistances As List(Of Double)
    Private Property ProjektstrukturPfadDistances As List(Of Double)
    Private _disposed As Boolean = False
    Private ReadOnly _embeddingServiceLock As New Object()
    Private _cascadeInProgress As Boolean = False
    Private _computedWeights As Dictionary(Of String, Dictionary(Of String, Double))

    ' Maps mutable engine feature names to the CascadeStep at which they first become available as INPUT.
    ' Uses the existing CascadeStep enum — no separate cascade order definition.
    Private Shared ReadOnly MutableFeatureAvailableFromStep As New Dictionary(Of String, CascadeStep)(StringComparer.OrdinalIgnoreCase) From {
        {"ProjektPfad", CascadeStep.ProjektstrukturPfad},
        {"ProjektstrukturPfad", CascadeStep.Titel},
        {"Titel", CascadeStep.AbsenderKurz},
        {"Ablageordner", CascadeStep.MsgDateinameSchema}
    }

    Public Enum CascadeStep
        ProjektstrukturPfad = 0
        Titel = 1
        AbsenderKurz = 2
        AblageordnerSchema = 3
        MsgDateinameSchema = 4
        AnhaengeAblegen = 5
    End Enum

    Public Sub New()
        Dim sw As Stopwatch = Stopwatch.StartNew()
        Debug.WriteLine("[SuggestionEngine] New() BEGIN")
        Try
            EnginesHistoricalSessionRecords = ThisAddIn.CurrentDatabaseManager.GetAllSessionRecords()
            Debug.WriteLine($"[SuggestionEngine]   GetAllSessionRecords: {sw.ElapsedMilliseconds} ms – {EnginesHistoricalSessionRecords.Count} records")
        Catch ex As Exception
            Debug.WriteLine("[SuggestionEngine] Fehler beim Laden der Session-Historie: " & ex.Message)
            EnginesHistoricalSessionRecords = New List(Of SessionRecord)()
        End Try
        LoadAllComputedWeights()
        Debug.WriteLine($"[SuggestionEngine] New() END – total: {sw.ElapsedMilliseconds} ms")
    End Sub

    Private Sub LoadAllComputedWeights()
        Dim sw As Stopwatch = Stopwatch.StartNew()
        Debug.WriteLine("[SuggestionEngine] LoadAllComputedWeights BEGIN")
        Dim loaded As Dictionary(Of String, Dictionary(Of String, Double))
        Try
            ' Eine Query fuer alle TargetFields statt sieben einzelne Verbindungen (Startpfad!).
            loaded = ThisAddIn.CurrentDatabaseManager.LoadAllComputedWeights()
        Catch ex As Exception
            Debug.WriteLine("[SuggestionEngine] Fehler beim Laden der ComputedWeights: " & ex.Message)
            loaded = New Dictionary(Of String, Dictionary(Of String, Double))(StringComparer.OrdinalIgnoreCase)
        End Try
        _computedWeights = loaded
        Debug.WriteLine($"[SuggestionEngine] LoadAllComputedWeights END – {loaded.Count} fields loaded: {sw.ElapsedMilliseconds} ms")
    End Sub

    Public Shared Function GetSharedInstance() As SuggestionEngine
        Return SharedEngineLazy.Value
    End Function

    Public Shared Sub PreloadSharedInstanceInBackground(Optional delayMs As Integer = 1500)
        Task.Run(Sub()
                     Dim sw As Stopwatch = Stopwatch.StartNew()
                     Try
                         If delayMs > 0 Then
                             Debug.WriteLine($"[SuggestionEngine] Preload: sleeping {delayMs} ms before init…")
                             Thread.Sleep(delayMs)
                         End If
                         Debug.WriteLine($"[SuggestionEngine] Preload: starting SharedEngineLazy.Value… ({sw.ElapsedMilliseconds} ms since queued)")
                         Dim engine = SharedEngineLazy.Value
                         Debug.WriteLine($"[SuggestionEngine] Preload: engine ready at {sw.ElapsedMilliseconds} ms — warming up EmbeddingService…")
                         engine.PreloadEmbeddingService()
                         Debug.WriteLine($"[SuggestionEngine] Preload: fully done. Total={sw.ElapsedMilliseconds} ms")
                     Catch ex As Exception
                         Debug.WriteLine($"[SuggestionEngine] Background preload failed after {sw.ElapsedMilliseconds} ms: {ex.Message}")
                     End Try
                 End Sub)
    End Sub

    ' Erstellt EmbeddingService und führt eine Dummy-Inferenz durch, damit ONNX beim ersten
    ' echten Aufruf nicht kalt startet. Sicher auf beliebigem Thread aufrufbar.
    Public Sub PreloadEmbeddingService()
        Dim svc = GetEmbeddingService()
        If svc Is Nothing Then Return
        Try
            Dim sw As Stopwatch = Stopwatch.StartNew()
            svc.GenerateEmbedding("warmup")
            Debug.WriteLine($"[SuggestionEngine] PreloadEmbeddingService: warmup inference done in {sw.ElapsedMilliseconds} ms")
        Catch ex As Exception
            Debug.WriteLine($"[SuggestionEngine] PreloadEmbeddingService: warmup failed: {ex.Message}")
        End Try
        EnsureHistoryEmbeddingsMatchModel(svc)
    End Sub

    Private Const EmbeddingModelIdKey As String = "EmbeddingModelId"
    ' Wird an die Modell-ID angehaengt, um die Neuberechnung auch OHNE Modellwechsel einmalig zu
    ' erzwingen. "+r2" (2026-09-30): bis dahin wurde das Betreff-Embedding der ersten Mail einer
    ' Outlook-Sitzung fuer alle weiteren wiederverwendet (Session.Reset leerte den Cache nicht) und
    ' so auch falsch gespeichert. Bei kuenftigen Fehlern dieser Art hochzaehlen.
    Private Const HistoryEmbeddingRevision As String = "+r2"

    ' Gespeicherte Betreff-Embeddings stammen von dem Modell, das beim Ablegen aktiv war. Nach einem
    ' Modellwechsel (z. B. englisches -> deutschfähiges Modell oder ein neu gebautes Modell) wären sie
    ' zwar gleich lang (384), aber nicht mehr mit neuen Embeddings vergleichbar - die Betreff-
    ' Ähnlichkeit würde stillschweigend zu Rauschen. Deshalb hier (Hintergrund-Thread des Preloads)
    ' alle Betreffs mit dem aktuellen Modell neu berechnen, wenn die in der Datenbank vermerkte
    ' Modell-ID abweicht (oder noch fehlt). Kosten: ein Embedding je Verlaufseintrag, einmalig.
    Private Sub EnsureHistoryEmbeddingsMatchModel(svc As EmbeddingService)
        Try
            Dim db = ThisAddIn.CurrentDatabaseManager
            Dim expectedId = svc.ModelId & HistoryEmbeddingRevision
            If db.GetMetaValue(EmbeddingModelIdKey) = expectedId Then Return
            Dim sw As Stopwatch = Stopwatch.StartNew()
            Dim updates As New Dictionary(Of Integer, Single())
            ' Per Index bis zur aktuellen Länge: der UI-Thread kann parallel neue Einträge anhängen
            ' (AppendHistoricalRecord) - diese tragen bereits Embeddings des aktuellen Modells.
            Dim count As Integer = EnginesHistoricalSessionRecords.Count
            For i As Integer = 0 To count - 1
                Dim record = EnginesHistoricalSessionRecords(i)
                If String.IsNullOrWhiteSpace(record.Betreff) Then Continue For
                Dim embedding = svc.GenerateEmbedding(record.Betreff)
                record.BetreffEmbedded = embedding
                updates(record.ID) = embedding
            Next
            db.UpdateBetreffEmbeddings(updates)
            db.SetMetaValue(EmbeddingModelIdKey, expectedId)
            Logger.LogInfo("Embedding-Modell", $"Betreff-Embeddings von {updates.Count} Verlaufseinträgen für Modell {svc.ModelId.Substring(0, 12)} neu berechnet in {sw.ElapsedMilliseconds} ms.")
        Catch ex As Exception
            Logger.LogError("EnsureHistoryEmbeddingsMatchModel", ex)
        End Try
    End Sub

    Public Shared Sub DisposeSharedInstance()
        If SharedEngineLazy.IsValueCreated Then
            SharedEngineLazy.Value.Dispose()
        End If
    End Sub

    ' Doppelt gepruefte Sperre: PreloadEmbeddingService laeuft per Task.Run auf einem
    ' Hintergrund-Thread, waehrend PrepareSession auf dem UI-Thread ueber
    ' GetOrCreateCurrentBetreffEmbedding denselben Weg nimmt. Oeffnet der Benutzer die Task Pane
    ' innerhalb des Preload-Fensters, konnten bisher beide Threads gleichzeitig Nothing sehen und
    ' JE eine EmbeddingService-Instanz erzeugen - das ONNX-Modell wurde also doppelt geladen und
    ' eine der beiden InferenceSessions nie disposed, weil das Feld nur eine davon behaelt.
    Private Function GetEmbeddingService() As EmbeddingService
        If EnginesEmbeddingService IsNot Nothing Then Return EnginesEmbeddingService
        SyncLock _embeddingServiceLock
            If EnginesEmbeddingService IsNot Nothing Then Return EnginesEmbeddingService
            Try
                Debug.WriteLine("[SuggestionEngine] GetEmbeddingService: constructing EmbeddingService (first call)…")
                Dim sw As Stopwatch = Stopwatch.StartNew()
                EnginesEmbeddingService = New EmbeddingService()
                Debug.WriteLine($"[SuggestionEngine] GetEmbeddingService: done in {sw.ElapsedMilliseconds} ms")
            Catch ex As Exception
                Debug.WriteLine("[SuggestionEngine] EmbeddingService konnte nicht erstellt werden: " & ex.Message)
                Return Nothing
            End Try
        End SyncLock
        Return EnginesEmbeddingService
    End Function

    Public Sub Dispose() Implements IDisposable.Dispose
        If _disposed Then Return
        ' Gleiche Sperre wie GetEmbeddingService: Shutdown kann waehrend eines noch laufenden
        ' Hintergrund-Preloads eintreffen.
        SyncLock _embeddingServiceLock
            EnginesEmbeddingService?.Dispose()
            EnginesEmbeddingService = Nothing
            _disposed = True
        End SyncLock
    End Sub

    ' Berechnet fixe Feature-Distanzen einmalig und initialisiert mutable Features mit 0.
    ' Wird direkt nach New() in PrepareSession aufgerufen.
    Public Sub CalculateInitialFeatureDistances(session As Session)
        If session Is Nothing Then Throw New ArgumentNullException(NameOf(session))
        Dim sw As Stopwatch = Stopwatch.StartNew()
        Dim t As Long
        Debug.WriteLine($"[SuggestionEngine] CalculateInitialFeatureDistances BEGIN ({EnginesHistoricalSessionRecords.Count} records)")

        RecalculateBetreffDistances(session)
        t = sw.ElapsedMilliseconds : Debug.WriteLine($"[SuggestionEngine]   BetreffDistances:       {t} ms")

        RecalculateDatumsDistances(session)
        Debug.WriteLine($"[SuggestionEngine]   DatumsDistances:        {sw.ElapsedMilliseconds - t} ms") : t = sw.ElapsedMilliseconds

        RecalculateAbsenderDomainDistances(session)
        Debug.WriteLine($"[SuggestionEngine]   AbsenderDomainDistances:{sw.ElapsedMilliseconds - t} ms") : t = sw.ElapsedMilliseconds

        RecalculateAbsenderDistances(session)
        Debug.WriteLine($"[SuggestionEngine]   AbsenderDistances:      {sw.ElapsedMilliseconds - t} ms") : t = sw.ElapsedMilliseconds

        RecalculateAusfueBenutzerDistances(session)
        Debug.WriteLine($"[SuggestionEngine]   AusfueBenutzerDistances:{sw.ElapsedMilliseconds - t} ms") : t = sw.ElapsedMilliseconds

        RecalculateAusfueDatumsDistances(session)
        Debug.WriteLine($"[SuggestionEngine]   AusfueDatumsDistances:  {sw.ElapsedMilliseconds - t} ms") : t = sw.ElapsedMilliseconds

        ' Mutable Features sind zur Initialisierungszeit leer – 0 als Startwert.
        Dim recordCount = EnginesHistoricalSessionRecords.Count
        TitelDistances = Enumerable.Repeat(0.0, recordCount).ToList()
        AblageordnerDistances = Enumerable.Repeat(0.0, recordCount).ToList()
        ProjektPfadDistances = Enumerable.Repeat(0.0, recordCount).ToList()
        ProjektstrukturPfadDistances = Enumerable.Repeat(0.0, recordCount).ToList()
        Debug.WriteLine($"[SuggestionEngine] CalculateInitialFeatureDistances END – total: {sw.ElapsedMilliseconds} ms")
    End Sub

    ' === Fixe Features – einmalig pro Mail-Selektion, aufgerufen aus CalculateInitialFeatureDistances ===

    Public Sub RecalculateBetreffDistances(session As Session)
        If session Is Nothing Then Return
        Dim newDistances As New List(Of Double)(EnginesHistoricalSessionRecords.Count)
        Dim currentEmbedding = GetOrCreateCurrentBetreffEmbedding(session)
        For Each record In EnginesHistoricalSessionRecords
            newDistances.Add(CalculateCosineSimilarity(currentEmbedding, record.BetreffEmbedded))
        Next
        BetreffDistances = newDistances
        Debug.WriteLine("[SuggestionEngine] BetreffDistances: " &
            String.Join(", ", EnginesHistoricalSessionRecords.Select(
                Function(r, i) $"Id={r.ID}:{newDistances(i):F3}")))
    End Sub

    Public Sub RecalculateDatumsDistances(session As Session)
        If session Is Nothing Then Return
        Dim rawDays = EnginesHistoricalSessionRecords.Select(Function(r) DateDistanceInDays(session.Datum, r.Datum)).ToList()
        DatumsDistances = NormalizeDateDistances(rawDays)
    End Sub

    Public Sub RecalculateAbsenderDomainDistances(session As Session)
        If session Is Nothing Then Return
        Dim newDistances As New List(Of Double)(EnginesHistoricalSessionRecords.Count)
        For Each record In EnginesHistoricalSessionRecords
            newDistances.Add(CalculateCategoricalSimilarity(session.AbsenderDomain, record.AbsenderDomain))
        Next
        AbsenderDomainDistances = newDistances
    End Sub

    Public Sub RecalculateAbsenderDistances(session As Session)
        If session Is Nothing Then Return
        Dim newDistances As New List(Of Double)(EnginesHistoricalSessionRecords.Count)
        For Each record In EnginesHistoricalSessionRecords
            newDistances.Add(CalculateCategoricalSimilarity(session.Absender, record.Absender))
        Next
        AbsenderDistances = newDistances
    End Sub

    Public Sub RecalculateAusfueBenutzerDistances(session As Session)
        If session Is Nothing Then Return
        Dim newDistances As New List(Of Double)(EnginesHistoricalSessionRecords.Count)
        For Each record In EnginesHistoricalSessionRecords
            newDistances.Add(CalculateCategoricalSimilarity(session.AusfueBenutzer, record.AusfueBenutzer))
        Next
        AusfueBenutzerDistances = newDistances
    End Sub

    Public Sub RecalculateAusfueDatumsDistances(session As Session)
        If session Is Nothing Then Return
        Dim rawDays = EnginesHistoricalSessionRecords.Select(Function(r) DateDistanceInDays(session.AusfueDatum, r.AusfueDatum)).ToList()
        AusfueDatumsDistances = NormalizeDateDistances(rawDays)
    End Sub

    ' === Mutable Features – ausgelöst bei Feldänderung über Session-Property-Setter ===

    Public Sub RecalculateTitelDistances(session As Session)
        If session Is Nothing Then Return
        Dim newDistances As New List(Of Double)(EnginesHistoricalSessionRecords.Count)
        Dim currentTokens = TokenizeToSet(session.Titel)
        For Each record In EnginesHistoricalSessionRecords
            newDistances.Add(CalculateTextSimilarityWithTokens(currentTokens, record.Titel))
        Next
        TitelDistances = newDistances
    End Sub

    Public Sub RecalculateAblageordnerDistances(session As Session)
        If session Is Nothing Then Return
        Dim newDistances As New List(Of Double)(EnginesHistoricalSessionRecords.Count)
        Dim currentTokens = TokenizeToSet(session.AblageordnerAufgeloest)
        For Each record In EnginesHistoricalSessionRecords
            newDistances.Add(CalculateTextSimilarityWithTokens(currentTokens, record.AblageordnerAufgeloest))
        Next
        AblageordnerDistances = newDistances
    End Sub

    Public Sub RecalculateProjektPfadDistances(session As Session)
        If session Is Nothing Then Return
        Dim newDistances As New List(Of Double)(EnginesHistoricalSessionRecords.Count)
        For Each record In EnginesHistoricalSessionRecords
            newDistances.Add(CalculateCategoricalSimilarity(session.ProjektPfad, record.ProjektPfad))
        Next
        ProjektPfadDistances = newDistances
    End Sub

    Public Sub RecalculateProjektstrukturPfadDistances(session As Session)
        If session Is Nothing Then Return
        Dim newDistances As New List(Of Double)(EnginesHistoricalSessionRecords.Count)
        For Each record In EnginesHistoricalSessionRecords
            newDistances.Add(CalculateCategoricalSimilarity(session.ProjektstrukturPfad, record.ProjektstrukturPfad))
        Next
        ProjektstrukturPfadDistances = newDistances
    End Sub

    ' Adds a newly saved record to the in-memory list AND extends every distance array by one zero,
    ' keeping all arrays in sync with the list length so FindBestRecordByField never goes out of range.
    Public Sub AppendHistoricalRecord(record As SessionRecord)
        If record Is Nothing Then Return
        EnginesHistoricalSessionRecords.Add(record)
        BetreffDistances?.Add(0.0)
        DatumsDistances?.Add(0.0)
        AbsenderDomainDistances?.Add(0.0)
        AbsenderDistances?.Add(0.0)
        AusfueBenutzerDistances?.Add(0.0)
        AusfueDatumsDistances?.Add(0.0)
        TitelDistances?.Add(0.0)
        AblageordnerDistances?.Add(0.0)
        ProjektPfadDistances?.Add(0.0)
        ProjektstrukturPfadDistances?.Add(0.0)
    End Sub

    Private Const DefaultSchemaTemplate As String = "[Datum (formatiert)]_[Absender (kurz)]_[Titel]"

    ' Viele Verlaufseintraege teilen sich denselben Pfad; ohne Cache wurde derselbe (Netz-)Pfad pro
    ' Kandidat erneut geprueft - bei nicht erreichbaren Freigaben jeweils mit Timeout. Der Cache
    ' gilt nur fuer einen einzelnen Vorschlagsaufruf, damit neu angelegte Ordner sofort zaehlen.
    Private Shared Function DirectoryExistsCached(path As String, cache As Dictionary(Of String, Boolean)) As Boolean
        Dim exists As Boolean
        If cache.TryGetValue(path, exists) Then Return exists
        exists = IO.Directory.Exists(path)
        cache(path) = exists
        Return exists
    End Function

    Public Function SuggestProjektPfad(session As Session) As String
        If session Is Nothing OrElse EnginesHistoricalSessionRecords.Count = 0 Then Return String.Empty
        Dim existsCache As New Dictionary(Of String, Boolean)(StringComparer.OrdinalIgnoreCase)
        For Each record In FindRecordsSortedByScore(Function(r) r.ProjektPfad, GetFeatureWeightsForProjektPfadSuggestion(), minScore:=SuggestionScoreThreshold)
            If Not String.IsNullOrWhiteSpace(record.ProjektPfad) AndAlso DirectoryExistsCached(record.ProjektPfad, existsCache) Then
                Debug.WriteLine($"[SuggestionEngine] SuggestProjektPfad: accepted '{record.ProjektPfad}'")
                Return record.ProjektPfad
            End If
            Debug.WriteLine($"[SuggestionEngine] SuggestProjektPfad: skipping non-existent '{record.ProjektPfad}'")
        Next
        Return String.Empty
    End Function

    ' Rueckgabe-Konvention: Nothing bedeutet "kein Vorschlag gefunden" (Cascade ueberspringt den
    ' Schritt). String.Empty ist ein GUELTIGER Vorschlag - z.B. der TreeView-Root-Knoten
    ' "Projektpfad" (direkte Ablage in ProjektPfad, RelativePath=String.Empty, siehe
    ' DirectoryTreeHelper). Vorher kollabierte String.Empty beide Faelle ("nichts gefunden" UND
    ' "leerer Wert vorgeschlagen") auf denselben Rueckgabewert, wodurch ein leerer
    ' ProjektstrukturPfad in der Historie nie als Vorschlag ankommen konnte, selbst wenn er der
    ' beste Treffer war. requireNonEmptyField:=False laesst solche Records ueberhaupt erst als
    ' Kandidaten zu (FindRecordsSortedByScore filtert sie sonst per Default heraus).
    Public Function SuggestProjektstrukturPfad(session As Session) As String
        If session Is Nothing OrElse EnginesHistoricalSessionRecords.Count = 0 Then Return Nothing
        If String.IsNullOrWhiteSpace(session.ProjektPfad) Then Return Nothing
        Dim existsCache As New Dictionary(Of String, Boolean)(StringComparer.OrdinalIgnoreCase)
        For Each record In FindRecordsSortedByScore(Function(r) r.ProjektstrukturPfad, GetFeatureWeightsForProjektstrukturPfadSuggestion(), requireNonEmptyField:=False, minScore:=SuggestionScoreThreshold)
            ' Path.Combine(ProjektPfad, "") = ProjektPfad selbst - existiert bereits validiert.
            Dim fullPath = IO.Path.Combine(session.ProjektPfad, If(record.ProjektstrukturPfad, String.Empty))
            If DirectoryExistsCached(fullPath, existsCache) Then
                Debug.WriteLine($"[SuggestionEngine] SuggestProjektstrukturPfad: accepted '{record.ProjektstrukturPfad}'")
                Return record.ProjektstrukturPfad
            End If
            Debug.WriteLine($"[SuggestionEngine] SuggestProjektstrukturPfad: skipping non-existent '{fullPath}'")
        Next
        Return Nothing
    End Function

    ' Selbe Nothing/String.Empty-Konvention wie SuggestProjektstrukturPfad: ein leerer Titel ist
    ' ein gueltiger, historisch beobachteter Wert (nicht jede Mail bekommt einen Titel), kein
    ' "nichts gefunden"-Signal.
    Public Function SuggestTitel(session As Session) As String
        If session Is Nothing OrElse EnginesHistoricalSessionRecords.Count = 0 Then Return Nothing
        Dim best = FindBestRecordByField(Function(r) r.Titel, GetFeatureWeightsForTitelSuggestion(), "Titel", requireNonEmptyField:=False, minScore:=SuggestionScoreThreshold)
        Return If(best Is Nothing, Nothing, best.Titel)
    End Function

    ' Selbe Nothing/String.Empty-Konvention wie SuggestProjektstrukturPfad/SuggestTitel: ein leerer
    ' Absender (kurz) ist ein gueltiger historischer Wert, kein "nichts gefunden"-Signal.
    Public Function SuggestAbsenderKurz(session As Session) As String
        If session Is Nothing OrElse EnginesHistoricalSessionRecords.Count = 0 Then Return Nothing
        Dim best = FindBestRecordByField(Function(r) r.AbsenderKurz, GetFeatureWeightsForAbsenderKurzSuggestion(), "AbsenderKurz", requireNonEmptyField:=False, minScore:=SuggestionScoreThreshold)
        Return If(best Is Nothing, Nothing, best.AbsenderKurz)
    End Function

    ' Hier kollidiert der "nichts gefunden"-Fall nicht mit einem gueltigen leeren Ablageordner-
    ' Schema, weil dessen Sentinel DefaultSchemaTemplate ist (nie leer) statt String.Empty - anders
    ' als bei Titel/AbsenderKurz/ProjektstrukturPfad reicht hier requireNonEmptyField:=False allein;
    ' best?.AblageordnerSchema liefert bereits korrekt "" durch, wenn best gefunden wurde.
    Public Function SuggestAblageordnerSchema(session As Session) As String
        If session Is Nothing OrElse EnginesHistoricalSessionRecords.Count = 0 Then Return DefaultSchemaTemplate
        Dim best = FindBestRecordByField(Function(r) r.AblageordnerSchema, GetFeatureWeightsForAblageordnerSuggestion(), "AblageordnerSchema", requireNonEmptyField:=False, minScore:=SuggestionScoreThreshold)
        Return If(best?.AblageordnerSchema, DefaultSchemaTemplate)
    End Function

    Public Function SuggestMsgDateinameSchema(session As Session) As String
        If session Is Nothing OrElse EnginesHistoricalSessionRecords.Count = 0 Then Return DefaultSchemaTemplate
        Dim best = FindBestRecordByField(Function(r) r.MsgDateinameSchema, GetFeatureWeightsForMsgDateinameSuggestion(), "MsgDateinameSchema", minScore:=SuggestionScoreThreshold)
        Return If(best?.MsgDateinameSchema, DefaultSchemaTemplate)
    End Function

    Public Function SuggestAnhaengeAblegen(session As Session) As Boolean?
        If session Is Nothing OrElse EnginesHistoricalSessionRecords.Count = 0 Then Return Nothing
        Dim bestRecord = FindBestRecordByField(Function(r) String.Empty, GetFeatureWeightsForAnhaengeAblegenSuggestion(), "AnhaengeAblegen", requireNonEmptyField:=False, minScore:=SuggestionScoreThreshold)
        If bestRecord Is Nothing Then Return Nothing
        Return bestRecord.AnhaengeAblegen
    End Function

    ' Orchestriert alle Vorschlagsschritte ab startFrom abwärts.
    ' Wird von Session-Property-Settern und PrepareSession aufgerufen.
    ' _cascadeInProgress verhindert Rekursion wenn ein Suggest-Aufruf den Setter triggert.
    Public Sub RunSuggestionCascade(session As Session, startFrom As CascadeStep)
        If session Is Nothing OrElse _cascadeInProgress Then Return
        _cascadeInProgress = True
        Try
            If startFrom <= CascadeStep.ProjektstrukturPfad Then
                Dim suggested = SuggestProjektstrukturPfad(session)
                If suggested IsNot Nothing Then
                    session.SuggestProjektstrukturPfad(suggested)
                Else
                    Debug.WriteLine($"[SuggestionEngine] Cascade: ProjektstrukturPfad übersprungen")
                End If
            End If

            If startFrom <= CascadeStep.Titel Then
                Dim suggested = SuggestTitel(session)
                If suggested IsNot Nothing Then
                    session.SuggestTitel(suggested)
                Else
                    Debug.WriteLine($"[SuggestionEngine] Cascade: Titel übersprungen")
                End If
            End If

            If startFrom <= CascadeStep.AbsenderKurz Then
                Dim suggested = SuggestAbsenderKurz(session)
                If suggested IsNot Nothing Then
                    session.SuggestAbsenderKurz(suggested)
                Else
                    Debug.WriteLine($"[SuggestionEngine] Cascade: AbsenderKurz übersprungen")
                End If
            End If

            ' SuggestAblageordnerSchema gibt nie Nothing zurueck (Sentinel fuer "nichts gefunden"
            ' ist DefaultSchemaTemplate, siehe dort) - dieser Zweig wendet daher immer an, auch den
            ' Default. Als IsNot-Nothing-Check geschrieben, um dieselbe Struktur wie die anderen
            ' Cascade-Schritte zu behalten.
            If startFrom <= CascadeStep.AblageordnerSchema Then
                Dim suggested = SuggestAblageordnerSchema(session)
                If suggested IsNot Nothing Then
                    session.SuggestAblageordnerSchema(suggested)
                Else
                    Debug.WriteLine($"[SuggestionEngine] Cascade: AblageordnerSchema übersprungen")
                End If
            End If

            If startFrom <= CascadeStep.MsgDateinameSchema Then
                Dim suggested = SuggestMsgDateinameSchema(session)
                If Not String.IsNullOrWhiteSpace(suggested) Then
                    session.SuggestMsgDateinameSchema(suggested)
                Else
                    Debug.WriteLine($"[SuggestionEngine] Cascade: MsgDateinameSchema übersprungen")
                End If
            End If

            If startFrom <= CascadeStep.AnhaengeAblegen Then
                If session.HasAnhaenge Then
                    Dim suggested = SuggestAnhaengeAblegen(session)
                    If suggested.HasValue Then
                        session.SuggestAnhaengeAblegen(suggested.Value)
                    Else
                        Debug.WriteLine($"[SuggestionEngine] Cascade: AnhaengeAblegen übersprungen")
                    End If
                Else
                    Debug.WriteLine($"[SuggestionEngine] Cascade: AnhaengeAblegen übersprungen (keine Anhänge)")
                End If
            End If
        Finally
            _cascadeInProgress = False
        End Try
    End Sub

    Private Const SuggestionScoreThreshold As Double = 0.5

    ' Gibt alle historischen Datensätze zurück, die den minScore erreichen, sortiert nach Score absteigend.
    ' Wird für ProjektPfad/ProjektstrukturPfad genutzt, um beim besten Treffer zu starten und
    ' zum nächsten auszuweichen, wenn der Pfad nicht existiert.
    Private Function FindRecordsSortedByScore(
        fieldSelector As Func(Of SessionRecord, String),
        featureWeights As IDictionary(Of String, Double),
        Optional requireNonEmptyField As Boolean = True,
        Optional minScore As Double = 0.0) As IEnumerable(Of SessionRecord)

        ' Gewichte einmal pro Suchlauf aufloesen statt zehn Dictionary-Lookups je Record.
        Dim wBetreff = FeatureWeight(featureWeights, "Betreff")
        Dim wDatum = FeatureWeight(featureWeights, "Datum")
        Dim wAbsenderDomain = FeatureWeight(featureWeights, "AbsenderDomain")
        Dim wAbsender = FeatureWeight(featureWeights, "Absender")
        Dim wAusfueBenutzer = FeatureWeight(featureWeights, "AusfueBenutzer")
        Dim wAusfueDatum = FeatureWeight(featureWeights, "AusfueDatum")
        Dim wTitel = FeatureWeight(featureWeights, "Titel")
        Dim wAblageordner = FeatureWeight(featureWeights, "Ablageordner")
        Dim wProjektPfad = FeatureWeight(featureWeights, "ProjektPfad")
        Dim wProjektstrukturPfad = FeatureWeight(featureWeights, "ProjektstrukturPfad")

        Dim scored As New List(Of KeyValuePair(Of Double, SessionRecord))()
        For i As Integer = 0 To EnginesHistoricalSessionRecords.Count - 1
            Dim record = EnginesHistoricalSessionRecords(i)
            If requireNonEmptyField AndAlso String.IsNullOrWhiteSpace(fieldSelector(record)) Then Continue For
            Dim score =
                wBetreff * BetreffDistances(i) +
                wDatum * DatumsDistances(i) +
                wAbsenderDomain * AbsenderDomainDistances(i) +
                wAbsender * AbsenderDistances(i) +
                wAusfueBenutzer * AusfueBenutzerDistances(i) +
                wAusfueDatum * AusfueDatumsDistances(i) +
                wTitel * TitelDistances(i) +
                wAblageordner * AblageordnerDistances(i) +
                wProjektPfad * ProjektPfadDistances(i) +
                wProjektstrukturPfad * ProjektstrukturPfadDistances(i)
            If score >= minScore Then
                scored.Add(New KeyValuePair(Of Double, SessionRecord)(score, record))
            End If
        Next
        Return scored.OrderByDescending(Function(kvp) kvp.Key).Select(Function(kvp) kvp.Value)
    End Function

    ' Findet den historischen Datensatz mit dem höchsten Gesamtscore, der für fieldSelector einen nicht-leeren Wert hat.
    ' Gibt Nothing zurück wenn der beste Score unter minScore liegt.
    Private Function FindBestRecordByField(fieldSelector As Func(Of SessionRecord, String), featureWeights As IDictionary(Of String, Double), suggestionName As String, Optional requireNonEmptyField As Boolean = True, Optional minScore As Double = 0.0) As SessionRecord

        Dim bestScore As Double = Double.MinValue
        Dim bestRecord As SessionRecord = Nothing
        Dim bestIndex As Integer = -1
        Dim bestUnconstrainedScore As Double = Double.MinValue
        Dim bestUnconstrainedIndex As Integer = -1

        ' Gewichte einmal pro Suchlauf aufloesen statt zehn Dictionary-Lookups je Record.
        Dim wBetreff = FeatureWeight(featureWeights, "Betreff")
        Dim wDatum = FeatureWeight(featureWeights, "Datum")
        Dim wAbsenderDomain = FeatureWeight(featureWeights, "AbsenderDomain")
        Dim wAbsender = FeatureWeight(featureWeights, "Absender")
        Dim wAusfueBenutzer = FeatureWeight(featureWeights, "AusfueBenutzer")
        Dim wAusfueDatum = FeatureWeight(featureWeights, "AusfueDatum")
        Dim wTitel = FeatureWeight(featureWeights, "Titel")
        Dim wAblageordner = FeatureWeight(featureWeights, "Ablageordner")
        Dim wProjektPfad = FeatureWeight(featureWeights, "ProjektPfad")
        Dim wProjektstrukturPfad = FeatureWeight(featureWeights, "ProjektstrukturPfad")

        For i As Integer = 0 To EnginesHistoricalSessionRecords.Count - 1
            Dim record = EnginesHistoricalSessionRecords(i)

            Dim score =
                wBetreff * BetreffDistances(i) +
                wDatum * DatumsDistances(i) +
                wAbsenderDomain * AbsenderDomainDistances(i) +
                wAbsender * AbsenderDistances(i) +
                wAusfueBenutzer * AusfueBenutzerDistances(i) +
                wAusfueDatum * AusfueDatumsDistances(i) +
                wTitel * TitelDistances(i) +
                wAblageordner * AblageordnerDistances(i) +
                wProjektPfad * ProjektPfadDistances(i) +
                wProjektstrukturPfad * ProjektstrukturPfadDistances(i)

            If score > bestUnconstrainedScore Then
                bestUnconstrainedScore = score
                bestUnconstrainedIndex = i
            End If

            If requireNonEmptyField AndAlso String.IsNullOrWhiteSpace(fieldSelector(record)) Then Continue For

            If score > bestScore Then
                bestScore = score
                bestRecord = record
                bestIndex = i
            End If
        Next

        If bestRecord Is Nothing Then
            Dim nonEmptyCount = EnginesHistoricalSessionRecords.Where(Function(r) Not String.IsNullOrWhiteSpace(fieldSelector(r))).Count()
            If bestUnconstrainedIndex >= 0 Then
                Dim ub = EnginesHistoricalSessionRecords(bestUnconstrainedIndex)
                Debug.WriteLine($"[SuggestionEngine] FindBestRecord for {suggestionName}: no suggestion — {nonEmptyCount} of {EnginesHistoricalSessionRecords.Count} records have a non-empty field; best overall: score={bestUnconstrainedScore:F3}, RecordId={ub.ID} | " &
                    FormatFeatureBreakdown(bestUnconstrainedIndex, featureWeights))
            Else
                Debug.WriteLine($"[SuggestionEngine] FindBestRecord for {suggestionName}: no suggestion — no records available")
            End If
            Return Nothing
        End If

        If bestScore < minScore Then
            Debug.WriteLine($"[SuggestionEngine] FindBestRecord for {suggestionName}: no suggestion — best score {bestScore:F3} below threshold {minScore}, RecordId={bestRecord.ID} | " &
                FormatFeatureBreakdown(bestIndex, featureWeights))
            Return Nothing
        End If

        Debug.WriteLine($"[SuggestionEngine] FindBestRecord for {suggestionName}: Bestscore={bestScore:F3}, BestRecordId={bestRecord.ID} | " &
            FormatFeatureBreakdown(bestIndex, featureWeights))
        Return bestRecord
    End Function

    Private Function FormatFeatureBreakdown(index As Integer, featureWeights As IDictionary(Of String, Double)) As String
        Return $"Betreff={BetreffDistances(index):F3}*{If(featureWeights.ContainsKey("Betreff"), featureWeights("Betreff"), 0)} " &
               $"Datum={DatumsDistances(index):F3}*{If(featureWeights.ContainsKey("Datum"), featureWeights("Datum"), 0)} " &
               $"Domain={AbsenderDomainDistances(index):F3}*{If(featureWeights.ContainsKey("AbsenderDomain"), featureWeights("AbsenderDomain"), 0)} " &
               $"Absender={AbsenderDistances(index):F3}*{If(featureWeights.ContainsKey("Absender"), featureWeights("Absender"), 0)} " &
               $"Benutzer={AusfueBenutzerDistances(index):F3}*{If(featureWeights.ContainsKey("AusfueBenutzer"), featureWeights("AusfueBenutzer"), 0)} " &
               $"AusfueDatum={AusfueDatumsDistances(index):F3}*{If(featureWeights.ContainsKey("AusfueDatum"), featureWeights("AusfueDatum"), 0)} " &
               $"Titel={TitelDistances(index):F3}*{If(featureWeights.ContainsKey("Titel"), featureWeights("Titel"), 0)} " &
               $"Ablageordner={AblageordnerDistances(index):F3}*{If(featureWeights.ContainsKey("Ablageordner"), featureWeights("Ablageordner"), 0)} " &
               $"ProjektPfad={ProjektPfadDistances(index):F3}*{If(featureWeights.ContainsKey("ProjektPfad"), featureWeights("ProjektPfad"), 0)} " &
               $"ProjektstrukturPfad={ProjektstrukturPfadDistances(index):F3}*{If(featureWeights.ContainsKey("ProjektstrukturPfad"), featureWeights("ProjektstrukturPfad"), 0)}"
    End Function

    Private Function GetFeatureWeightsForProjektPfadSuggestion() As IDictionary(Of String, Double)
        Dim computed As Dictionary(Of String, Double) = Nothing
        If _computedWeights?.TryGetValue("ProjektPfad", computed) Then Return computed
        Return New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
            {"Betreff", 0.46},
            {"AbsenderDomain", 0.18},
            {"Absender", 0.18},
            {"AusfueDatum", 0.18}
        }
    End Function

    Private Function GetFeatureWeightsForProjektstrukturPfadSuggestion() As IDictionary(Of String, Double)
        Dim computed As Dictionary(Of String, Double) = Nothing
        If _computedWeights?.TryGetValue("ProjektstrukturPfad", computed) Then Return computed
        Return New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
            {"Betreff", 0.46},
            {"AbsenderDomain", 0.09},
            {"Absender", 0.09},
            {"AusfueDatum", 0.1},
            {"ProjektPfad", 0.18}
        }
    End Function

    Private Function GetFeatureWeightsForTitelSuggestion() As IDictionary(Of String, Double)
        Dim computed As Dictionary(Of String, Double) = Nothing
        If _computedWeights?.TryGetValue("Titel", computed) Then Return computed
        Return New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
            {"Betreff", 0.55},
            {"Absender", 0.1},
            {"AusfueDatum", 0.15},
            {"ProjektPfad", 0.1},
            {"ProjektstrukturPfad", 0.1}
        }
    End Function

    Private Function GetFeatureWeightsForAblageordnerSuggestion() As IDictionary(Of String, Double)
        Dim computed As Dictionary(Of String, Double) = Nothing
        If _computedWeights?.TryGetValue("AblageordnerSchema", computed) Then Return computed
        Return New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
            {"AusfueDatum", 0.18},
            {"ProjektPfad", 0.36},
            {"ProjektstrukturPfad", 0.41}
        }
    End Function

    Private Function GetFeatureWeightsForAbsenderKurzSuggestion() As IDictionary(Of String, Double)
        Dim computed As Dictionary(Of String, Double) = Nothing
        If _computedWeights?.TryGetValue("AbsenderKurz", computed) Then Return computed
        Return New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
            {"AbsenderDomain", 0.26},
            {"Absender", 0.17},
            {"AusfueDatum", 0.17},
            {"ProjektPfad", 0.36}
        }
    End Function

    Private Function GetFeatureWeightsForMsgDateinameSuggestion() As IDictionary(Of String, Double)
        Dim computed As Dictionary(Of String, Double) = Nothing
        If _computedWeights?.TryGetValue("MsgDateinameSchema", computed) Then Return computed
        Return New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
            {"AusfueDatum", 0.18},
            {"ProjektPfad", 0.36},
            {"ProjektstrukturPfad", 0.41}
        }
    End Function

    Private Function GetFeatureWeightsForAnhaengeAblegenSuggestion() As IDictionary(Of String, Double)
        Dim computed As Dictionary(Of String, Double) = Nothing
        If _computedWeights?.TryGetValue("AnhaengeAblegen", computed) Then Return computed
        Return New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
            {"Absender", 0.2},
            {"AusfueDatum", 0.08},
            {"ProjektPfad", 0.16},
            {"ProjektstrukturPfad", 0.41}
        }
    End Function

    ' Erzeugt das Embedding für den aktuellen Betreff und speichert es in der Session.
    Private Function GetOrCreateCurrentBetreffEmbedding(currentSession As Session) As Single()
        If currentSession.BetreffEmbedded IsNot Nothing AndAlso currentSession.BetreffEmbedded.Length > 0 Then
            Return currentSession.BetreffEmbedded
        End If
        If String.IsNullOrWhiteSpace(currentSession.Betreff) Then
            Return Nothing
        End If

        Try
            currentSession.BetreffEmbedded = GetEmbeddingService()?.GenerateEmbedding(currentSession.Betreff)
        Catch ex As Exception
            Debug.WriteLine("[SuggestionEngine] BetreffEmbedding generation failed: " & ex.Message)
            currentSession.BetreffEmbedded = Nothing
        End Try
        Return currentSession.BetreffEmbedded
    End Function

    ' Berechnet Cosine Similarity im Bereich [-1, 1]. Für fehlende Werte wird 0 verwendet.
    Private Function CalculateCosineSimilarity(vectorA As Single(), vectorB As Single()) As Double
        If vectorA Is Nothing OrElse vectorB Is Nothing Then
            Return 0.0
        End If
        If vectorA.Length = 0 OrElse vectorB.Length = 0 OrElse vectorA.Length <> vectorB.Length Then
            Return 0.0
        End If

        Dim dotProduct As Double = 0
        Dim normA As Double = 0
        Dim normB As Double = 0

        For i As Integer = 0 To vectorA.Length - 1
            dotProduct += vectorA(i) * vectorB(i)
            normA += vectorA(i) * vectorA(i)
            normB += vectorB(i) * vectorB(i)
        Next

        If normA = 0 OrElse normB = 0 Then
            Return 0.0
        End If
        Return dotProduct / (Math.Sqrt(normA) * Math.Sqrt(normB))
    End Function

    ' Vergleicht zwei kategoriale Werte: exakter Match = 1, sonst 0.
    Private Function CalculateCategoricalSimilarity(currentValue As String, historicalValue As String) As Double
        If String.IsNullOrWhiteSpace(currentValue) OrElse String.IsNullOrWhiteSpace(historicalValue) Then
            Return 0.0
        End If
        Return If(String.Equals(currentValue.Trim(), historicalValue.Trim(), StringComparison.OrdinalIgnoreCase), 1.0, 0.0)
    End Function

    ' Normalisiert Datumsunterschiede auf [0, 1], wobei 1 = sehr ähnlich.
    ' Returns the raw absolute distance in days; -1 when either date is missing (sentinel for NormalizeDateDistances).
    Private Function DateDistanceInDays(a As DateTime, b As DateTime) As Double
        If a = Date.MinValue OrElse b = Date.MinValue Then Return -1.0
        Return Math.Abs((a - b).TotalDays)
    End Function

    ' Converts a list of raw day-distances (from DateDistanceInDays) to [0,1] similarities.
    ' The most-distant valid entry becomes 0.0, the closest becomes 1.0.
    ' Entries with the sentinel value -1 (missing date) map to 0.0.
    Private Function NormalizeDateDistances(rawDays As List(Of Double)) As List(Of Double)
        Dim maxCalc = rawDays.Where(Function(d) d >= 0).DefaultIfEmpty(0).Max()
        Dim maxUse = Math.Min(180.0, maxCalc)
        If maxUse <= 0 Then maxUse = 1.0
        Return rawDays.Select(Function(d)
                                  If d < 0 Then Return 0.0
                                  Return Math.Max(0.0, 1.0 - d / maxUse)
                              End Function).ToList()
    End Function

    ' Textähnlichkeit auf Basis normalisierter Token-Überlappung (Jaccard), Bereich [0,1].
    Private Function CalculateTextSimilarity(currentValue As String, historicalValue As String) As Double
        Return JaccardSimilarity(TokenizeToSet(currentValue), TokenizeToSet(historicalValue))
    End Function

    ' Variante mit bereits berechneter Tokenmenge fuer die "aktuelle" Seite: in den
    ' Recalculate*Distances-Schleifen ist diese Seite ueber alle Records konstant, wurde bisher
    ' aber fuer jeden einzelnen Record erneut tokenisiert und in ein neues HashSet kopiert.
    Private Function CalculateTextSimilarityWithTokens(currentTokens As HashSet(Of String), historicalValue As String) As Double
        Return JaccardSimilarity(currentTokens, TokenizeToSet(historicalValue))
    End Function

    Private Function TokenizeToSet(value As String) As HashSet(Of String)
        If String.IsNullOrWhiteSpace(value) Then Return New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Return New HashSet(Of String)(TokenizeForSimilarity(value), StringComparer.OrdinalIgnoreCase)
    End Function

    ' Jaccard-Index zweier Tokenmengen. Die Schnittmenge wird direkt gezaehlt und die
    ' Vereinigungsgroesse ueber |A| + |B| - |A geschnitten B| abgeleitet - das spart die beiden
    ' LINQ-Zwischensequenzen, die Intersect()/Union() zuvor pro Vergleich aufgebaut haben.
    Private Function JaccardSimilarity(currentTokens As HashSet(Of String), historicalTokens As HashSet(Of String)) As Double
        If currentTokens Is Nothing OrElse historicalTokens Is Nothing Then Return 0.0
        If currentTokens.Count = 0 OrElse historicalTokens.Count = 0 Then Return 0.0

        Dim intersectionCount As Integer = 0
        For Each token In historicalTokens
            If currentTokens.Contains(token) Then intersectionCount += 1
        Next

        Dim unionCount = currentTokens.Count + historicalTokens.Count - intersectionCount
        If unionCount = 0 Then Return 0.0
        Return intersectionCount / unionCount
    End Function

    ' Liefert das Gewicht eines einzelnen Features (0.0, wenn im Profil nicht gesetzt).
    Private Function FeatureWeight(featureWeights As IDictionary(Of String, Double), featureName As String) As Double
        Dim weight As Double = 0.0
        If featureWeights Is Nothing OrElse Not featureWeights.TryGetValue(featureName, weight) Then
            Return 0.0
        End If
        Return weight
    End Function

    ' Zerlegt Strings in einfache Vergleichstoken und entfernt leere Segmente.
    Private Function TokenizeForSimilarity(value As String) As IEnumerable(Of String)
        Return value _
            .ToLowerInvariant() _
            .Split(New Char() {" "c, "_"c, "-"c, "."c, "/"c, "\"c, ","c, ";"c, ":"c, "("c, ")"c, "["c, "]"c}, StringSplitOptions.RemoveEmptyEntries) _
            .Where(Function(token) token.Trim().Length > 0)
    End Function

    ' === Korrelationsbasierte Gewichtsberechnung ===

    ' Geometrischer statt fixer Recalc-Trigger: neu berechnen, wenn recordCount eines der
    ' geometrisch wachsenden "Meilenstein"-Folgenglieder 1, 2, 3, 4, 5, 7, 9, 12, 15, 19, 24, 30,
    ' 38, 48, 60, 75, 94, ... (milestone_(i+1) = max(milestone_i + 1, ceil(milestone_i *
    ' growthFactor)), milestone_0 = 0) ist, statt bei jedem n-ten Datensatz. Grund:
    ' RecalculateWeightsFromHistory ist O(n^2) (Paarschleife ueber alle Records x 7 Zielfelder),
    ' waehrend der statistische Nutzen weiterer Datenpunkte mit O(1/sqrt(n)) abnimmt
    ' (Pearson-Schaetzer). Ein fixes Intervall (z.B. alle 50) fuehrt zu kubisch wachsenden
    ' Lebenszeit-Gesamtkosten (Summe von (50i)^2 ueber alle Meilensteine ~ n^3), da spaete, teure
    ' Neuberechnungen genauso oft anfallen wie fruehe, guenstige. Die geometrische Folge braucht
    ' dagegen nur O(log n) Meilensteine ueber die Lebenszeit; da jede Stufe um growthFactor
    ' groesser ist als die vorherige, ist die kumulierte Kost von der jeweils letzten (groessten)
    ' Neuberechnung dominiert - die Gesamtkosten bleiben ein konstantes Vielfaches (~ growthFactor^2 /
    ' (growthFactor^2 - 1), fuer 1.25 also Faktor ~2.3) der Kosten einer einzigen Neuberechnung bei
    ' aktueller Groesse, statt unbegrenzt zu wachsen. Frueh loest das nahezu bei jeder Ablage aus,
    ' wo die Neuberechnung noch billig ist und jeder zusaetzliche Datenpunkt die Gewichte noch
    ' spuerbar veraendern kann; spaeter werden die Abstaende automatisch groesser.
    '
    ' Die Meilenstein-Folge selbst haengt nur von growthFactor ab, nicht davon, ob/wann eine
    ' vorherige Neuberechnung tatsaechlich stattfand oder in ComputedWeights geschrieben wurde -
    ' recordCount allein reicht daher aus, um per Simulation zu entscheiden, ob es ein
    ' Folgenglied ist. Das ersetzt einen fruehreren Entwurf, der zusaetzlich den RecordCount der
    ' letzten erfolgreichen Neuberechnung aus der DB nachschlug (SessionDatabaseManager.
    ' GetLastWeightRecalcRecordCount(), inzwischen wieder entfernt): unnoetig, da die Folge
    ' ohnehin deterministisch ist, und eine DB-Abfrage bei jeder Ablage spart. Die Schleife
    ' braucht dank geometrischen Wachstums nur O(log_growthFactor(recordCount)) Schritte
    ' (~40 bei 10.000 Datensaetzen und growthFactor=1.25) - vernachlaessigbar gegenueber der
    ' bisherigen SQLite-Abfrage, geschweige denn gegenueber RecalculateWeightsFromHistory selbst.
    Friend Shared Function ShouldRecalculateWeights(recordCount As Integer, Optional growthFactor As Double = 1.25) As Boolean
        If recordCount <= 0 Then Return False
        Dim milestone As Integer = 0
        Do
            milestone = Math.Max(milestone + 1, CInt(Math.Ceiling(milestone * growthFactor)))
            If milestone = recordCount Then Return True
        Loop While milestone < recordCount
        Return False
    End Function

    ' Serialisiert parallele Neuberechnungen (automatischer Trigger nach dem Speichern und der
    ' Button im Info-Popup koennen gleichzeitig laufen).
    Private ReadOnly _recalcLock As New Object()

    Public Sub RecalculateWeightsFromHistory()
        SyncLock _recalcLock
            RecalculateWeightsFromHistoryCore()
        End SyncLock
    End Sub

    Private Sub RecalculateWeightsFromHistoryCore()
        Dim records As List(Of SessionRecord)
        Try
            records = ThisAddIn.CurrentDatabaseManager.GetAllSessionRecords()
        Catch ex As Exception
            Debug.WriteLine("[SuggestionEngine] RecalculateWeightsFromHistory: DB-Ladefehler: " & ex.Message)
            Return
        End Try
        If records.Count < 2 Then
            Debug.WriteLine($"[SuggestionEngine] RecalculateWeightsFromHistory: Zu wenige Records ({records.Count}), Abbruch.")
            Return
        End If

        Dim oldWeights = _computedWeights
        Dim globalMaxDatumDist As Double = ComputeGlobalMaxDateDistance(records, Function(r) r.Datum)
        Dim globalMaxAusfueDatumDist As Double = ComputeGlobalMaxDateDistance(records, Function(r) r.AusfueDatum)

        Dim allFeatureNames() As String = WeightFeatureNames
        Dim targetFields() As String = {
            "ProjektPfad", "ProjektstrukturPfad", "Titel", "AbsenderKurz",
            "AblageordnerSchema", "MsgDateinameSchema", "AnhaengeAblegen"
        }

        ' Zielfelder, die die Schwelle erreichen, vorab bestimmen: die Korrelationen aller dieser
        ' Felder entstehen dann in EINEM Durchlauf ueber die Record-Paare (ComputeRawCorrelations).
        Dim passingTargets As New List(Of String)
        For Each targetField In targetFields
            Dim K0, M0, t0, f0 As Integer
            GetThresholdInfo(records, targetField, K0, M0, t0, f0)
            If String.Equals(targetField, "AnhaengeAblegen", StringComparison.OrdinalIgnoreCase) Then
                If t0 >= 2 AndAlso f0 >= 2 Then passingTargets.Add(targetField)
            ElseIf K0 >= 2 Then
                passingTargets.Add(targetField)
            End If
        Next
        Dim allRawCorrs = ComputeRawCorrelations(records, passingTargets, globalMaxDatumDist, globalMaxAusfueDatumDist)

        ' Mit den bisherigen Gewichten starten und nur neu berechnete Zielfelder ersetzen. Zuvor
        ' wurde das Dictionary komplett ersetzt: Zielfelder, die diesmal unter der Schwelle lagen
        ' oder nur Null-Korrelationen hatten, fielen im Speicher auf die festen Gewichte zurueck,
        ' obwohl ihre gelernten Gewichte in der Datenbank blieben (und nach dem naechsten
        ' Outlook-Start wieder galten).
        Dim newWeights As New Dictionary(Of String, Dictionary(Of String, Double))(StringComparer.OrdinalIgnoreCase)
        If oldWeights IsNot Nothing Then
            For Each kvp In oldWeights
                newWeights(kvp.Key) = kvp.Value
            Next
        End If
        For Each targetField In targetFields
            Dim K, M, trueCount, falseCount As Integer
            GetThresholdInfo(records, targetField, K, M, trueCount, falseCount)
            Dim isAnhaenge = String.Equals(targetField, "AnhaengeAblegen", StringComparison.OrdinalIgnoreCase)

            Debug.WriteLine($"[SuggestionEngine] RecalculateWeights | TargetField={targetField} | Records={records.Count}")
            Dim passes As Boolean
            If isAnhaenge Then
                passes = trueCount >= 2 AndAlso falseCount >= 2
                Debug.WriteLine($"  Threshold: True={trueCount}, False={falseCount}" &
                                If(passes, " → PASSED", $" → FAILED ({If(trueCount < 2, "True", "False")}<2)"))
            Else
                passes = K >= 2
                Debug.WriteLine($"  Threshold: K={K} distinct values" &
                                If(passes, " → PASSED", " → FAILED (K<2)"))
            End If
            If Not passes Then Continue For

            Dim featureNames = GetCascadeAwareFeaturesForTarget(targetField)
            Dim rawCorrs As Dictionary(Of String, Double) = Nothing
            allRawCorrs.TryGetValue(targetField, rawCorrs)
            Dim weights = NormalizeCorrelations(rawCorrs, featureNames)

            Debug.WriteLine("  Pearson-Korrelationen (Cascade-Features):")
            For Each feat In featureNames
                Dim rawVal As Double = Double.NaN
                If rawCorrs IsNot Nothing Then rawCorrs.TryGetValue(feat, rawVal)
                If Double.IsNaN(rawVal) Then
                    Debug.WriteLine($"    {feat,-22} r=NaN    → Feature-Vektor konstant (Pearson nicht definiert)")
                ElseIf rawVal < 0 Then
                    Debug.WriteLine($"    {feat,-22} r={rawVal,-8:F3} → negativ, auf 0 geclippt")
                Else
                    Debug.WriteLine($"    {feat,-22} r=+{rawVal,-7:F3} → gewichtet")
                End If
            Next
            Dim skipped = allFeatureNames.Where(Function(f) Not featureNames.Contains(f, StringComparer.OrdinalIgnoreCase)).ToArray()
            If skipped.Length > 0 Then
                Debug.WriteLine("  Nicht in Cascade-Stufe (übersprungen): " & String.Join(", ", skipped))
            End If

            If weights IsNot Nothing AndAlso weights.Count > 0 Then
                Dim beforeW As IDictionary(Of String, Double)
                Dim oldW As Dictionary(Of String, Double) = Nothing
                If oldWeights IsNot Nothing AndAlso oldWeights.TryGetValue(targetField, oldW) Then
                    beforeW = oldW
                Else
                    beforeW = GetHardcodedWeightsForTarget(targetField)
                End If
                Dim beforeSum = If(beforeW IsNot Nothing AndAlso beforeW.Count > 0, beforeW.Values.Sum(), 1.0)
                If beforeSum <= 0 Then beforeSum = 1.0

                Debug.WriteLine("  Gewichtsvergleich (Feature | Vorher | Nachher):")
                For Each feat In featureNames
                    Dim before As Double = 0
                    beforeW?.TryGetValue(feat, before)
                    Dim after As Double = 0
                    weights.TryGetValue(feat, after)
                    Debug.WriteLine($"    {feat,-22} {before / beforeSum:F3} → {after:F3}")
                Next

                Try
                    ThisAddIn.CurrentDatabaseManager.SaveComputedWeights(targetField, weights, records.Count)
                Catch ex As Exception
                    Debug.WriteLine($"[SuggestionEngine] Fehler beim Speichern der Gewichte für '{targetField}': " & ex.Message)
                End Try
                newWeights(targetField) = weights
            Else
                Debug.WriteLine("  Alle Korrelationen = 0, kein Update (degenerierter Fall)")
            End If
        Next
        _computedWeights = newWeights
    End Sub

    Private Function ComputeGlobalMaxDateDistance(records As List(Of SessionRecord),
                                                  dateSelector As Func(Of SessionRecord, DateTime)) As Double
        ' Das Maximum von |di - dj| ueber alle Paare ist definitionsgemaess
        ' (groesstes Datum - kleinstes Datum). Ein O(n)-Durchlauf liefert also exakt dasselbe
        ' Ergebnis wie die fruehere O(n^2)-Paarschleife (inkl. der Untergrenze 1.0).
        ' Datensaetze ohne Datum (Date.MinValue) bleiben wie bisher unberuecksichtigt.
        Dim minDate As DateTime = Date.MaxValue
        Dim maxDate As DateTime = Date.MinValue
        Dim validCount As Integer = 0
        For Each record In records
            Dim d = dateSelector(record)
            If d = Date.MinValue Then Continue For
            validCount += 1
            If d < minDate Then minDate = d
            If d > maxDate Then maxDate = d
        Next
        If validCount < 2 Then Return 1.0
        Return Math.Max(1.0, (maxDate - minDate).TotalDays)
    End Function

    Private Shared ReadOnly WeightFeatureNames() As String = {
        "Betreff", "Datum", "AbsenderDomain", "Absender", "AusfueBenutzer", "AusfueDatum",
        "Titel", "Ablageordner", "ProjektPfad", "ProjektstrukturPfad"
    }

    ' Pearson-Korrelation jedes Features (Aehnlichkeit eines Record-Paars) mit dem Label "beide
    ' Records haben denselben Zielwert" - fuer ALLE uebergebenen Zielfelder in EINEM Durchlauf ueber
    ' die O(n^2) Paare, mit laufenden Summen statt Vektoren. Zuvor wurde pro Zielfeld (bis zu 7x)
    ' erneut ueber alle Paare iteriert, jedes Mal u.a. die 384-dimensionale Cosine-Similarity
    ' neu berechnet, und je Feature eine Liste mit einem Double pro Paar aufgebaut (bei 1.000
    ' Records ~500.000 Paare * 10 Features * 8 Byte = ~40 MB je Zielfeld plus Listen-Wachstum).
    ' Ergebnis je Zielfeld: Feature -> r (Double.NaN = Feature-Vektor konstant, 0 = Label konstant),
    ' identisch zur frueheren Berechnung bis auf Rundung. Die Werte werden relativ zum ersten Paar
    ' verschoben aufsummiert, damit ein konstanter Feature-Vektor exakt Varianz 0 ergibt.
    Private Function ComputeRawCorrelations(records As List(Of SessionRecord),
                                            targetFields As List(Of String),
                                            globalMaxDatumDist As Double,
                                            globalMaxAusfueDatumDist As Double) As Dictionary(Of String, Dictionary(Of String, Double))
        Dim result As New Dictionary(Of String, Dictionary(Of String, Double))(StringComparer.OrdinalIgnoreCase)
        Dim n = records.Count
        Dim tCount = targetFields.Count
        If tCount = 0 OrElse n < 2 Then Return result
        Const FeatureCount As Integer = 10

        Dim maxDatumUse = Math.Min(180.0, globalMaxDatumDist)
        Dim maxAusfueUse = Math.Min(180.0, globalMaxAusfueDatumDist)
        If maxDatumUse <= 0 Then maxDatumUse = 1.0
        If maxAusfueUse <= 0 Then maxAusfueUse = 1.0

        ' Je Record EINMAL vorberechnen: Embedding-Norm, Kategorie-IDs, Tokenmengen, Zielwert-IDs.
        Dim embeddings As Single()() = New Single(n - 1)() {}
        Dim norms(n - 1) As Double
        For i As Integer = 0 To n - 1
            Dim e = records(i).BetreffEmbedded
            embeddings(i) = e
            If e IsNot Nothing AndAlso e.Length > 0 Then
                Dim sq As Double = 0
                For k As Integer = 0 To e.Length - 1
                    sq += CDbl(e(k)) * e(k)
                Next
                norms(i) = Math.Sqrt(sq)
            End If
        Next
        Dim domainIds = GetCategoryIds(records, Function(r) r.AbsenderDomain)
        Dim absenderIds = GetCategoryIds(records, Function(r) r.Absender)
        Dim benutzerIds = GetCategoryIds(records, Function(r) r.AusfueBenutzer)
        Dim projektIds = GetCategoryIds(records, Function(r) r.ProjektPfad)
        Dim strukturIds = GetCategoryIds(records, Function(r) r.ProjektstrukturPfad)
        Dim titelTokens = records.Select(Function(r) TokenizeToSet(r.Titel)).ToArray()
        Dim ablageordnerTokens = records.Select(Function(r) TokenizeToSet(r.AblageordnerAufgeloest)).ToArray()
        Dim targetIds As Integer()() = New Integer(tCount - 1)() {}
        For t As Integer = 0 To tCount - 1
            targetIds(t) = ToValueIds(GetTargetValues(records, targetFields(t)))
        Next

        Dim sx(FeatureCount - 1) As Double
        Dim sxx(FeatureCount - 1) As Double
        Dim shift(FeatureCount - 1) As Double
        Dim sy(tCount - 1) As Double
        Dim sxy(tCount - 1, FeatureCount - 1) As Double
        Dim f(FeatureCount - 1) As Double
        Dim pairCount As Long = 0

        For i As Integer = 0 To n - 2
            For j As Integer = i + 1 To n - 1
                f(0) = CosineWithNorms(embeddings(i), norms(i), embeddings(j), norms(j))
                Dim rawDatum = DateDistanceInDays(records(i).Datum, records(j).Datum)
                f(1) = If(rawDatum < 0, 0.0, Math.Max(0.0, 1.0 - rawDatum / maxDatumUse))
                f(2) = If(domainIds(i) >= 0 AndAlso domainIds(i) = domainIds(j), 1.0, 0.0)
                f(3) = If(absenderIds(i) >= 0 AndAlso absenderIds(i) = absenderIds(j), 1.0, 0.0)
                f(4) = If(benutzerIds(i) >= 0 AndAlso benutzerIds(i) = benutzerIds(j), 1.0, 0.0)
                Dim rawAusfue = DateDistanceInDays(records(i).AusfueDatum, records(j).AusfueDatum)
                f(5) = If(rawAusfue < 0, 0.0, Math.Max(0.0, 1.0 - rawAusfue / maxAusfueUse))
                f(6) = JaccardSimilarity(titelTokens(i), titelTokens(j))
                f(7) = JaccardSimilarity(ablageordnerTokens(i), ablageordnerTokens(j))
                f(8) = If(projektIds(i) >= 0 AndAlso projektIds(i) = projektIds(j), 1.0, 0.0)
                f(9) = If(strukturIds(i) >= 0 AndAlso strukturIds(i) = strukturIds(j), 1.0, 0.0)

                If pairCount = 0 Then Array.Copy(f, shift, FeatureCount)
                pairCount += 1
                For k As Integer = 0 To FeatureCount - 1
                    Dim x = f(k) - shift(k)
                    sx(k) += x
                    sxx(k) += x * x
                Next
                For t As Integer = 0 To tCount - 1
                    Dim ids = targetIds(t)
                    If ids(i) = ids(j) Then
                        sy(t) += 1.0
                        For k As Integer = 0 To FeatureCount - 1
                            sxy(t, k) += f(k) - shift(k)
                        Next
                    End If
                Next
            Next
        Next

        Dim np As Double = pairCount
        For t As Integer = 0 To tCount - 1
            Dim corrs As New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase)
            ' Label ist binaer (y*y = y), daher Summe der Quadrate = sy.
            Dim varY = sy(t) - sy(t) * sy(t) / np
            For k As Integer = 0 To FeatureCount - 1
                Dim varX = sxx(k) - sx(k) * sx(k) / np
                Dim r As Double
                If varX <= 0.0 Then
                    r = Double.NaN            ' feature vector is constant → Pearson undefined
                ElseIf varY <= 0.0 Then
                    r = 0.0                   ' label vector is constant → degenerate case
                Else
                    Dim cov = sxy(t, k) - sx(k) * sy(t) / np
                    r = Math.Max(-1.0, Math.Min(1.0, cov / Math.Sqrt(varX * varY)))
                End If
                corrs(WeightFeatureNames(k)) = r
            Next
            result(targetFields(t)) = corrs
        Next
        Return result
    End Function

    ' Clippt die Roh-Korrelationen der kaskadenkonformen Features auf [0, ∞) und normiert sie auf
    ' Summe 1. Nothing, wenn keine positive Korrelation uebrig bleibt (degenerierter Fall).
    Private Function NormalizeCorrelations(rawCorrs As Dictionary(Of String, Double),
                                           featureNames As String()) As Dictionary(Of String, Double)
        If rawCorrs Is Nothing Then Return Nothing
        Dim clipped As New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase)
        For Each feat In featureNames
            Dim r As Double = Double.NaN
            rawCorrs.TryGetValue(feat, r)
            clipped(feat) = If(Double.IsNaN(r), 0.0, Math.Max(0.0, r))
        Next
        Dim total = clipped.Values.Sum()
        If total <= 0.0 Then Return Nothing
        Dim normalized As New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase)
        For Each kvp In clipped
            normalized(kvp.Key) = kvp.Value / total
        Next
        Return normalized
    End Function

    ' Cosine-Similarity mit vorberechneten Normen; gleiche Randfaelle wie CalculateCosineSimilarity
    ' (fehlender/leerer Vektor, ungleiche Laenge oder Norm 0 -> 0).
    Private Shared Function CosineWithNorms(a As Single(), normA As Double, b As Single(), normB As Double) As Double
        If a Is Nothing OrElse b Is Nothing OrElse a.Length = 0 OrElse a.Length <> b.Length Then Return 0.0
        If normA = 0 OrElse normB = 0 Then Return 0.0
        Dim dot As Double = 0
        For k As Integer = 0 To a.Length - 1
            dot += CDbl(a(k)) * b(k)
        Next
        Return dot / (normA * normB)
    End Function

    ' Kategorie-ID je Record (gleiche Semantik wie CalculateCategoricalSimilarity: getrimmt,
    ' Gross-/Kleinschreibung egal); -1 fuer leere Werte, die nie uebereinstimmen.
    Private Shared Function GetCategoryIds(records As List(Of SessionRecord), selector As Func(Of SessionRecord, String)) As Integer()
        Dim ids(records.Count - 1) As Integer
        Dim map As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        For i As Integer = 0 To records.Count - 1
            Dim v = selector(records(i))
            If String.IsNullOrWhiteSpace(v) Then
                ids(i) = -1
                Continue For
            End If
            Dim key = v.Trim()
            Dim id As Integer
            If Not map.TryGetValue(key, id) Then
                id = map.Count
                map(key) = id
            End If
            ids(i) = id
        Next
        Return ids
    End Function

    ' Zielwert-IDs (TargetValuesMatch-Semantik: Gross-/Kleinschreibung egal, leer ist eine eigene Kategorie).
    Private Shared Function ToValueIds(values As String()) As Integer()
        Dim ids(values.Length - 1) As Integer
        Dim map As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        For i As Integer = 0 To values.Length - 1
            Dim key = If(values(i), String.Empty)
            Dim id As Integer
            If Not map.TryGetValue(key, id) Then
                id = map.Count
                map(key) = id
            End If
            ids(i) = id
        Next
        Return ids
    End Function

    ' Returns which engine features (by name) are causally available as inputs when targetField is being suggested.
    ' Uses MutableFeatureAvailableFromStep + the existing CascadeStep enum — no separate cascade order definition.
    Private Function GetCascadeAwareFeaturesForTarget(targetField As String) As String()
        Dim fixedFeatures() As String = {"Betreff", "Datum", "AbsenderDomain", "Absender", "AusfueBenutzer", "AusfueDatum"}
        If String.Equals(targetField, "ProjektPfad", StringComparison.OrdinalIgnoreCase) Then Return fixedFeatures
        Dim stepEnum As CascadeStep
        If Not [Enum].TryParse(Of CascadeStep)(targetField, True, stepEnum) Then Return fixedFeatures
        Dim result As New List(Of String)(fixedFeatures)
        For Each kvp In MutableFeatureAvailableFromStep
            If CInt(kvp.Value) <= CInt(stepEnum) Then result.Add(kvp.Key)
        Next
        Return result.ToArray()
    End Function

    ' Populates threshold diagnostic values for debug output. For AnhaengeAblegen: K/M unused, trueCount/falseCount relevant.
    ' For string fields: trueCount/falseCount unused, K = distinct value count, M = min count per value (informational only; not used in threshold).
    Private Sub GetThresholdInfo(records As List(Of SessionRecord), targetField As String,
                                  ByRef K As Integer, ByRef M As Integer,
                                  ByRef trueCount As Integer, ByRef falseCount As Integer)
        K = 0 : M = 0 : trueCount = 0 : falseCount = 0
        If String.Equals(targetField, "AnhaengeAblegen", StringComparison.OrdinalIgnoreCase) Then
            trueCount = records.Where(Function(r) r.AnhaengeAblegen = True).Count()
            falseCount = records.Count - trueCount
            Return
        End If
        Dim propInfo = GetType(SessionRecord).GetProperty(targetField)
        If propInfo Is Nothing Then Return
        ' Ein leerer Zielwert ist ein gueltiger, eigenstaendiger Wert (z.B. Titel oder
        ' ProjektstrukturPfad leer gelassen) und zaehlt daher als eigene Kategorie mit, statt
        ' herausgefiltert zu werden - sonst wuerde die Gewichtung nie lernen, wie gut Features
        ' einen "leer bleibt es"-Fall vorhersagen.
        Dim valueCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        For Each record In records
            Dim v = TryCast(propInfo.GetValue(record), String)
            Dim key = If(v, String.Empty).Trim()
            If Not valueCounts.ContainsKey(key) Then valueCounts(key) = 0
            valueCounts(key) += 1
        Next
        K = valueCounts.Count
        M = If(valueCounts.Count > 0, valueCounts.Values.Min(), 0)
    End Sub

    ' Returns the hardcoded fallback weights for targetField (without consulting _computedWeights).
    ' Used as the "Vorher" baseline in debug weight comparison output.
    Private Function GetHardcodedWeightsForTarget(targetField As String) As IDictionary(Of String, Double)
        Select Case targetField
            Case "ProjektPfad"
                Return New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
                    {"Betreff", 0.46}, {"AbsenderDomain", 0.18}, {"Absender", 0.18}, {"AusfueDatum", 0.18}}
            Case "ProjektstrukturPfad"
                Return New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
                    {"Betreff", 0.46}, {"AbsenderDomain", 0.09}, {"Absender", 0.09}, {"AusfueDatum", 0.1}, {"ProjektPfad", 0.18}}
            Case "Titel"
                Return New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
                    {"Betreff", 0.55}, {"Absender", 0.1}, {"AusfueDatum", 0.15}, {"ProjektPfad", 0.1}, {"ProjektstrukturPfad", 0.1}}
            Case "AbsenderKurz"
                Return New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
                    {"AbsenderDomain", 0.26}, {"Absender", 0.17}, {"AusfueDatum", 0.17}, {"ProjektPfad", 0.36}}
            Case "AblageordnerSchema"
                Return New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
                    {"AusfueDatum", 0.18}, {"ProjektPfad", 0.36}, {"ProjektstrukturPfad", 0.41}}
            Case "MsgDateinameSchema"
                Return New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
                    {"AusfueDatum", 0.18}, {"ProjektPfad", 0.36}, {"ProjektstrukturPfad", 0.41}}
            Case "AnhaengeAblegen"
                Return New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
                    {"Absender", 0.2}, {"AusfueDatum", 0.08}, {"ProjektPfad", 0.16}, {"ProjektstrukturPfad", 0.41}}
            Case Else
                Return New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase)
        End Select
    End Function

    ' Liest den Zielwert jedes Records EINMAL aus, statt ihn (wie zuvor) per Reflection fuer jedes
    ' der O(n^2) Record-Paare erneut zu ermitteln: bei 500 Records sind das 124.750 Paare * 2
    ' GetProperty/GetValue-Aufrufe pro TargetField, also ueber eine Million Reflection-Aufrufe je
    ' Gewichtsneuberechnung. Ein leerer Zielwert wird zu String.Empty normalisiert statt (wie
    ' zuvor) zu Nothing - Nothing bedeutete "matcht nie", wodurch zwei Records, die beide denselben
    ' Zielwert leer gelassen haben, nie als "gleich" fuer das Pearson-Label gezaehlt wurden, selbst
    ' wenn "leer" der historisch haeufigste/konsistenteste Wert war.
    Private Function GetTargetValues(records As List(Of SessionRecord), targetField As String) As String()
        Dim values(records.Count - 1) As String
        If String.Equals(targetField, "AnhaengeAblegen", StringComparison.OrdinalIgnoreCase) Then
            For i As Integer = 0 To records.Count - 1
                values(i) = If(records(i).AnhaengeAblegen, "1", "0")
            Next
            Return values
        End If
        Dim propInfo = GetType(SessionRecord).GetProperty(targetField)
        If propInfo Is Nothing Then Return values
        For i As Integer = 0 To records.Count - 1
            Dim v = TryCast(propInfo.GetValue(records(i)), String)
            values(i) = If(v, String.Empty).Trim()
        Next
        Return values
    End Function

End Class
