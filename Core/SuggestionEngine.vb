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
    Private _cascadeInProgress As Boolean = False
    Private _computedWeights As Dictionary(Of String, Dictionary(Of String, Double))

    ' Maps mutable engine feature names to the CascadeStep at which they first become available as INPUT.
    ' Uses the existing CascadeStep enum — no separate cascade order definition.
    Private Shared ReadOnly MutableFeatureAvailableFromStep As New Dictionary(Of String, CascadeStep)(StringComparer.OrdinalIgnoreCase) From {
        {"ProjektPfad",         CascadeStep.ProjektstrukturPfad},
        {"ProjektstrukturPfad", CascadeStep.Titel},
        {"Titel",               CascadeStep.AbsenderKurz},
        {"Ablageordner",        CascadeStep.MsgDateinameSchema}
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
        Try
            EnginesHistoricalSessionRecords = ThisAddIn.CurrentDatabaseManager.GetAllSessionRecords()
        Catch ex As Exception
            Debug.WriteLine("[SuggestionEngine] Fehler beim Laden der Session-Historie: " & ex.Message)
            EnginesHistoricalSessionRecords = New List(Of SessionRecord)()
        End Try
        LoadAllComputedWeights()
    End Sub

    Private Sub LoadAllComputedWeights()
        Dim loaded As New Dictionary(Of String, Dictionary(Of String, Double))(StringComparer.OrdinalIgnoreCase)
        Dim targetFields() As String = {
            "ProjektPfad", "ProjektstrukturPfad", "Titel", "AbsenderKurz",
            "AblageordnerSchema", "MsgDateinameSchema", "AnhaengeAblegen"
        }
        Try
            For Each field In targetFields
                Dim weights = ThisAddIn.CurrentDatabaseManager.LoadComputedWeights(field)
                If weights.Count > 0 Then
                    loaded(field) = weights
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine("[SuggestionEngine] Fehler beim Laden der ComputedWeights: " & ex.Message)
        End Try
        _computedWeights = loaded
    End Sub

    Public Shared Function GetSharedInstance() As SuggestionEngine
        Return SharedEngineLazy.Value
    End Function

    Public Shared Sub PreloadSharedInstanceInBackground(Optional delayMs As Integer = 1500)
        Task.Run(Sub()
                     Try
                         If delayMs > 0 Then
                             Thread.Sleep(delayMs)
                         End If
                         Dim ignored = SharedEngineLazy.Value
                         Debug.WriteLine("[SuggestionEngine] Shared instance preloaded in background.")
                     Catch ex As Exception
                         Debug.WriteLine("[SuggestionEngine] Background preload failed: " & ex.Message)
                     End Try
                 End Sub)
    End Sub

    Public Shared Sub DisposeSharedInstance()
        If SharedEngineLazy.IsValueCreated Then
            SharedEngineLazy.Value.Dispose()
        End If
    End Sub

    Private Function GetEmbeddingService() As EmbeddingService
        If EnginesEmbeddingService Is Nothing Then
            Try
                EnginesEmbeddingService = New EmbeddingService()
            Catch ex As Exception
                Debug.WriteLine("[SuggestionEngine] EmbeddingService konnte nicht erstellt werden: " & ex.Message)
                Return Nothing
            End Try
        End If
        Return EnginesEmbeddingService
    End Function

    Public Sub Dispose() Implements IDisposable.Dispose
        If _disposed Then Return
        EnginesEmbeddingService?.Dispose()
        EnginesEmbeddingService = Nothing
        _disposed = True
    End Sub

    ' Berechnet fixe Feature-Distanzen einmalig und initialisiert mutable Features mit 0.
    ' Wird direkt nach New() in PrepareSession aufgerufen.
    Public Sub CalculateInitialFeatureDistances(session As Session)
        If session Is Nothing Then Throw New ArgumentNullException(NameOf(session))

        RecalculateBetreffDistances(session)
        RecalculateDatumsDistances(session)
        RecalculateAbsenderDomainDistances(session)
        RecalculateAbsenderDistances(session)
        RecalculateAusfueBenutzerDistances(session)
        RecalculateAusfueDatumsDistances(session)

        ' Mutable Features sind zur Initialisierungszeit leer – 0 als Startwert.
        Dim recordCount = EnginesHistoricalSessionRecords.Count
        TitelDistances = Enumerable.Repeat(0.0, recordCount).ToList()
        AblageordnerDistances = Enumerable.Repeat(0.0, recordCount).ToList()
        ProjektPfadDistances = Enumerable.Repeat(0.0, recordCount).ToList()
        ProjektstrukturPfadDistances = Enumerable.Repeat(0.0, recordCount).ToList()
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
        For Each record In EnginesHistoricalSessionRecords
            newDistances.Add(CalculateTextSimilarity(session.Titel, record.Titel))
        Next
        TitelDistances = newDistances
    End Sub

    Public Sub RecalculateAblageordnerDistances(session As Session)
        If session Is Nothing Then Return
        Dim newDistances As New List(Of Double)(EnginesHistoricalSessionRecords.Count)
        For Each record In EnginesHistoricalSessionRecords
            newDistances.Add(CalculateTextSimilarity(session.AblageordnerAufgeloest, record.AblageordnerAufgeloest))
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

    Public Function SuggestProjektPfad(session As Session) As String
        If session Is Nothing OrElse EnginesHistoricalSessionRecords.Count = 0 Then Return String.Empty
        Return If(FindBestRecordByField(Function(r) r.ProjektPfad, GetFeatureWeightsForProjektPfadSuggestion(), "ProjektPfad", minScore:=SuggestionScoreThreshold)?.ProjektPfad, String.Empty)
    End Function

    Public Function SuggestProjektstrukturPfad(session As Session) As String
        If session Is Nothing OrElse EnginesHistoricalSessionRecords.Count = 0 Then Return String.Empty
        Return If(FindBestRecordByField(Function(r) r.ProjektstrukturPfad, GetFeatureWeightsForProjektstrukturPfadSuggestion(), "ProjektstrukturPfad", minScore:=SuggestionScoreThreshold)?.ProjektstrukturPfad, String.Empty)
    End Function

    Public Function SuggestTitel(session As Session) As String
        If session Is Nothing OrElse EnginesHistoricalSessionRecords.Count = 0 Then Return String.Empty
        Return If(FindBestRecordByField(Function(r) r.Titel, GetFeatureWeightsForTitelSuggestion(), "Titel", minScore:=SuggestionScoreThreshold)?.Titel, String.Empty)
    End Function

    Public Function SuggestAbsenderKurz(session As Session) As String
        If session Is Nothing OrElse EnginesHistoricalSessionRecords.Count = 0 Then Return String.Empty
        Return If(FindBestRecordByField(Function(r) r.AbsenderKurz, GetFeatureWeightsForAbsenderKurzSuggestion(), "AbsenderKurz", minScore:=SuggestionScoreThreshold)?.AbsenderKurz, String.Empty)
    End Function

    Public Function SuggestAblageordnerSchema(session As Session) As String
        If session Is Nothing OrElse EnginesHistoricalSessionRecords.Count = 0 Then Return DefaultSchemaTemplate
        Dim best = FindBestRecordByField(Function(r) r.AblageordnerSchema, GetFeatureWeightsForAblageordnerSuggestion(), "AblageordnerSchema", minScore:=SuggestionScoreThreshold)
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
                If Not String.IsNullOrWhiteSpace(suggested) AndAlso
                   Not String.IsNullOrWhiteSpace(session.ProjektPfad) AndAlso
                   IO.Directory.Exists(IO.Path.Combine(session.ProjektPfad, suggested)) Then
                    session.SuggestProjektstrukturPfad(suggested)
                Else
                    Debug.WriteLine($"[SuggestionEngine] Cascade: ProjektstrukturPfad übersprungen")
                End If
            End If

            If startFrom <= CascadeStep.Titel Then
                Dim suggested = SuggestTitel(session)
                If Not String.IsNullOrWhiteSpace(suggested) Then
                    session.SuggestTitel(suggested)
                Else
                    Debug.WriteLine($"[SuggestionEngine] Cascade: Titel übersprungen")
                End If
            End If

            If startFrom <= CascadeStep.AbsenderKurz Then
                Dim suggested = SuggestAbsenderKurz(session)
                If Not String.IsNullOrWhiteSpace(suggested) Then
                    session.SuggestAbsenderKurz(suggested)
                Else
                    Debug.WriteLine($"[SuggestionEngine] Cascade: AbsenderKurz übersprungen")
                End If
            End If

            If startFrom <= CascadeStep.AblageordnerSchema Then
                Dim suggested = SuggestAblageordnerSchema(session)
                If Not String.IsNullOrWhiteSpace(suggested) Then
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
                Dim suggested = SuggestAnhaengeAblegen(session)
                If suggested.HasValue Then
                    session.SuggestAnhaengeAblegen(suggested.Value)
                Else
                    Debug.WriteLine($"[SuggestionEngine] Cascade: AnhaengeAblegen übersprungen")
                End If
            End If
        Finally
            _cascadeInProgress = False
        End Try
    End Sub

    Private Const SuggestionScoreThreshold As Double = 0.5

    ' Findet den historischen Datensatz mit dem höchsten Gesamtscore, der für fieldSelector einen nicht-leeren Wert hat.
    ' Gibt Nothing zurück wenn der beste Score unter minScore liegt.
    Private Function FindBestRecordByField(fieldSelector As Func(Of SessionRecord, String), featureWeights As IDictionary(Of String, Double), suggestionName As String, Optional requireNonEmptyField As Boolean = True, Optional minScore As Double = 0.0) As SessionRecord

        Dim bestScore As Double = Double.MinValue
        Dim bestRecord As SessionRecord = Nothing
        Dim bestIndex As Integer = -1
        Dim bestUnconstrainedScore As Double = Double.MinValue
        Dim bestUnconstrainedIndex As Integer = -1

        For i As Integer = 0 To EnginesHistoricalSessionRecords.Count - 1
            Dim record = EnginesHistoricalSessionRecords(i)

            Dim score =
                WeightedFeatureScore(featureWeights, "Betreff", BetreffDistances(i)) +
                WeightedFeatureScore(featureWeights, "Datum", DatumsDistances(i)) +
                WeightedFeatureScore(featureWeights, "AbsenderDomain", AbsenderDomainDistances(i)) +
                WeightedFeatureScore(featureWeights, "Absender", AbsenderDistances(i)) +
                WeightedFeatureScore(featureWeights, "AusfueBenutzer", AusfueBenutzerDistances(i)) +
                WeightedFeatureScore(featureWeights, "AusfueDatum", AusfueDatumsDistances(i)) +
                WeightedFeatureScore(featureWeights, "Titel", TitelDistances(i)) +
                WeightedFeatureScore(featureWeights, "Ablageordner", AblageordnerDistances(i)) +
                WeightedFeatureScore(featureWeights, "ProjektPfad", ProjektPfadDistances(i)) +
                WeightedFeatureScore(featureWeights, "ProjektstrukturPfad", ProjektstrukturPfadDistances(i))

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
            currentSession.BetreffEmbedded = GetEmbeddingService()?.GenerateEmbedding(currentSession.Betreff.ToLower())
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
        If String.IsNullOrWhiteSpace(currentValue) OrElse String.IsNullOrWhiteSpace(historicalValue) Then
            Return 0.0
        End If

        Dim currentTokens = New HashSet(Of String)(TokenizeForSimilarity(currentValue), StringComparer.OrdinalIgnoreCase)
        Dim historicalTokens = New HashSet(Of String)(TokenizeForSimilarity(historicalValue), StringComparer.OrdinalIgnoreCase)

        If currentTokens.Count = 0 OrElse historicalTokens.Count = 0 Then
            Return 0.0
        End If

        Dim intersectionCount = currentTokens.Intersect(historicalTokens, StringComparer.OrdinalIgnoreCase).Count()
        Dim unionCount = currentTokens.Union(historicalTokens, StringComparer.OrdinalIgnoreCase).Count()
        If unionCount = 0 Then
            Return 0.0
        End If

        Return intersectionCount / unionCount
    End Function

    ' Liefert den gewichteten Beitrag eines einzelnen Features.
    Private Function WeightedFeatureScore(featureWeights As IDictionary(Of String, Double), featureName As String, featureDistance As Double) As Double
        Dim weight As Double = 0.0
        If featureWeights Is Nothing OrElse Not featureWeights.TryGetValue(featureName, weight) Then
            Return 0.0
        End If
        Return weight * featureDistance
    End Function

    ' Zerlegt Strings in einfache Vergleichstoken und entfernt leere Segmente.
    Private Function TokenizeForSimilarity(value As String) As IEnumerable(Of String)
        Return value _
            .ToLowerInvariant() _
            .Split(New Char() {" "c, "_"c, "-"c, "."c, "/"c, "\"c, ","c, ";"c, ":"c, "("c, ")"c, "["c, "]"c}, StringSplitOptions.RemoveEmptyEntries) _
            .Where(Function(token) token.Trim().Length > 0)
    End Function

    ' === Korrelationsbasierte Gewichtsberechnung ===

    Public Sub RecalculateWeightsFromHistory()
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

        Dim allFeatureNames() As String = {
            "Betreff", "Datum", "AbsenderDomain", "Absender", "AusfueBenutzer", "AusfueDatum",
            "Titel", "Ablageordner", "ProjektPfad", "ProjektstrukturPfad"
        }
        Dim targetFields() As String = {
            "ProjektPfad", "ProjektstrukturPfad", "Titel", "AbsenderKurz",
            "AblageordnerSchema", "MsgDateinameSchema", "AnhaengeAblegen"
        }

        Dim newWeights As New Dictionary(Of String, Dictionary(Of String, Double))(StringComparer.OrdinalIgnoreCase)
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
                passes = K >= 3
                Debug.WriteLine($"  Threshold: K={K} distinct values" &
                                If(passes, " → PASSED", " → FAILED (K<3)"))
            End If
            If Not passes Then Continue For

            Dim featureNames = GetCascadeAwareFeaturesForTarget(targetField)
            Dim rawCorrs As Dictionary(Of String, Double) = Nothing
            Dim weights = ComputeWeightsForTarget(records, targetField, featureNames, globalMaxDatumDist, globalMaxAusfueDatumDist, rawCorrs)

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
        Dim maxDist As Double = 1.0
        For i As Integer = 0 To records.Count - 2
            Dim di = dateSelector(records(i))
            If di = Date.MinValue Then Continue For
            For j As Integer = i + 1 To records.Count - 1
                Dim dj = dateSelector(records(j))
                If dj = Date.MinValue Then Continue For
                Dim dist = Math.Abs((di - dj).TotalDays)
                If dist > maxDist Then maxDist = dist
            Next
        Next
        Return maxDist
    End Function

    ' featureNames: cascade-aware subset of features to evaluate (from GetCascadeAwareFeaturesForTarget).
    ' rawCorrs: populated with raw Pearson values (before clipping; NaN = constant feature vector).
    Private Function ComputeWeightsForTarget(records As List(Of SessionRecord),
                                             targetField As String,
                                             featureNames As String(),
                                             globalMaxDatumDist As Double,
                                             globalMaxAusfueDatumDist As Double,
                                             ByRef rawCorrs As Dictionary(Of String, Double)) As Dictionary(Of String, Double)
        Dim featureVectors As New Dictionary(Of String, List(Of Double))(StringComparer.OrdinalIgnoreCase)
        For Each feat In featureNames
            featureVectors(feat) = New List(Of Double)()
        Next

        Dim labelVector As New List(Of Double)()
        Dim maxDatumUse = Math.Min(180.0, globalMaxDatumDist)
        Dim maxAusfueUse = Math.Min(180.0, globalMaxAusfueDatumDist)
        If maxDatumUse <= 0 Then maxDatumUse = 1.0
        If maxAusfueUse <= 0 Then maxAusfueUse = 1.0

        For i As Integer = 0 To records.Count - 2
            Dim ri = records(i)
            For j As Integer = i + 1 To records.Count - 1
                Dim rj = records(j)
                labelVector.Add(If(TargetFieldMatch(ri, rj, targetField), 1.0, 0.0))
                If featureVectors.ContainsKey("Betreff") Then featureVectors("Betreff").Add(CalculateCosineSimilarity(ri.BetreffEmbedded, rj.BetreffEmbedded))
                If featureVectors.ContainsKey("Datum") Then
                    Dim rawDatum = DateDistanceInDays(ri.Datum, rj.Datum)
                    featureVectors("Datum").Add(If(rawDatum < 0, 0.0, Math.Max(0.0, 1.0 - rawDatum / maxDatumUse)))
                End If
                If featureVectors.ContainsKey("AbsenderDomain") Then featureVectors("AbsenderDomain").Add(CalculateCategoricalSimilarity(ri.AbsenderDomain, rj.AbsenderDomain))
                If featureVectors.ContainsKey("Absender") Then featureVectors("Absender").Add(CalculateCategoricalSimilarity(ri.Absender, rj.Absender))
                If featureVectors.ContainsKey("AusfueBenutzer") Then featureVectors("AusfueBenutzer").Add(CalculateCategoricalSimilarity(ri.AusfueBenutzer, rj.AusfueBenutzer))
                If featureVectors.ContainsKey("AusfueDatum") Then
                    Dim rawAusfue = DateDistanceInDays(ri.AusfueDatum, rj.AusfueDatum)
                    featureVectors("AusfueDatum").Add(If(rawAusfue < 0, 0.0, Math.Max(0.0, 1.0 - rawAusfue / maxAusfueUse)))
                End If
                If featureVectors.ContainsKey("Titel") Then featureVectors("Titel").Add(CalculateTextSimilarity(ri.Titel, rj.Titel))
                If featureVectors.ContainsKey("Ablageordner") Then featureVectors("Ablageordner").Add(CalculateTextSimilarity(ri.AblageordnerAufgeloest, rj.AblageordnerAufgeloest))
                If featureVectors.ContainsKey("ProjektPfad") Then featureVectors("ProjektPfad").Add(CalculateCategoricalSimilarity(ri.ProjektPfad, rj.ProjektPfad))
                If featureVectors.ContainsKey("ProjektstrukturPfad") Then featureVectors("ProjektstrukturPfad").Add(CalculateCategoricalSimilarity(ri.ProjektstrukturPfad, rj.ProjektstrukturPfad))
            Next
        Next

        If labelVector.Count = 0 Then Return Nothing

        Dim rawCorrelations As New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase)
        Dim clippedCorrelations As New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase)
        Dim allZero As Boolean = True
        For Each kvp In featureVectors
            Dim r = PearsonCorrelation(kvp.Value, labelVector)
            rawCorrelations(kvp.Key) = r
            Dim clipped = If(Double.IsNaN(r), 0.0, Math.Max(0.0, r))
            clippedCorrelations(kvp.Key) = clipped
            If clipped > 0.0 Then allZero = False
        Next
        rawCorrs = rawCorrelations

        If allZero Then Return Nothing

        Dim total = clippedCorrelations.Values.Sum()
        Dim normalized As New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase)
        For Each kvp In clippedCorrelations
            normalized(kvp.Key) = If(total > 0.0, kvp.Value / total, 0.0)
        Next
        Return normalized
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
        Dim valueCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        For Each record In records
            Dim v = TryCast(propInfo.GetValue(record), String)
            If String.IsNullOrWhiteSpace(v) Then Continue For
            Dim key = v.Trim()
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

    Private Function TargetFieldMatch(ri As SessionRecord, rj As SessionRecord, targetField As String) As Boolean
        If String.Equals(targetField, "AnhaengeAblegen", StringComparison.OrdinalIgnoreCase) Then
            Return ri.AnhaengeAblegen = rj.AnhaengeAblegen
        End If
        Dim propInfo = GetType(SessionRecord).GetProperty(targetField)
        If propInfo Is Nothing Then Return False
        Dim vi = TryCast(propInfo.GetValue(ri), String)
        Dim vj = TryCast(propInfo.GetValue(rj), String)
        If String.IsNullOrWhiteSpace(vi) OrElse String.IsNullOrWhiteSpace(vj) Then Return False
        Return String.Equals(vi.Trim(), vj.Trim(), StringComparison.OrdinalIgnoreCase)
    End Function

    Private Function PearsonCorrelation(x As List(Of Double), y As List(Of Double)) As Double
        If x Is Nothing OrElse y Is Nothing Then Return 0.0
        Dim n = x.Count
        If n <> y.Count OrElse n < 2 Then Return 0.0

        Dim sumX As Double = 0, sumY As Double = 0
        For i As Integer = 0 To n - 1
            sumX += x(i)
            sumY += y(i)
        Next
        Dim meanX = sumX / n
        Dim meanY = sumY / n

        Dim cov As Double = 0, varX As Double = 0, varY As Double = 0
        For i As Integer = 0 To n - 1
            Dim dx = x(i) - meanX
            Dim dy = y(i) - meanY
            cov += dx * dy
            varX += dx * dx
            varY += dy * dy
        Next

        If varX = 0.0 Then Return Double.NaN  ' feature vector is constant → Pearson undefined
        If varY = 0.0 Then Return 0.0          ' label vector is constant → degenerate case
        Return cov / Math.Sqrt(varX * varY)
    End Function

End Class
