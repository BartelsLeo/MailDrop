Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
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

    Public Sub New()
        Try
            EnginesHistoricalSessionRecords = ThisAddIn.CurrentDatabaseManager.GetAllSessionRecords()
        Catch ex As Exception
            Debug.WriteLine("[SuggestionEngine] Fehler beim Laden der Session-Historie: " & ex.Message)
            EnginesHistoricalSessionRecords = New List(Of SessionRecord)()
        End Try
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
        Dim newDistances As New List(Of Double)(EnginesHistoricalSessionRecords.Count)
        Dim maxDistance = GetMaxDateDistanceInDays(session.Datum)
        For Each record In EnginesHistoricalSessionRecords
            newDistances.Add(CalculateDateSimilarity(session.Datum, record.Datum, maxDistance))
        Next
        DatumsDistances = newDistances
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
        Dim newDistances As New List(Of Double)(EnginesHistoricalSessionRecords.Count)
        Dim maxDistance = GetMaxDateDistanceInDays(session.AusfueDatum)
        For Each record In EnginesHistoricalSessionRecords
            newDistances.Add(CalculateDateSimilarity(session.AusfueDatum, record.AusfueDatum, maxDistance))
        Next
        AusfueDatumsDistances = newDistances
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

    Private Const SuggestionScoreThreshold As Double = 0.3

    ' Findet den historischen Datensatz mit dem höchsten Gesamtscore, der für fieldSelector einen nicht-leeren Wert hat.
    ' Gibt Nothing zurück wenn der beste Score unter minScore liegt.
    Private Function FindBestRecordByField(fieldSelector As Func(Of SessionRecord, String), featureWeights As IDictionary(Of String, Double), suggestionName As String, Optional requireNonEmptyField As Boolean = True, Optional minScore As Double = 0.0) As SessionRecord

        Dim bestScore As Double = Double.MinValue
        Dim bestRecord As SessionRecord = Nothing
        Dim bestIndex As Integer = -1
        For i As Integer = 0 To EnginesHistoricalSessionRecords.Count - 1
            Dim record = EnginesHistoricalSessionRecords(i)
            If requireNonEmptyField AndAlso String.IsNullOrWhiteSpace(fieldSelector(record)) Then Continue For

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

            If score > bestScore Then
                bestScore = score
                bestRecord = record
                bestIndex = i
            End If
        Next

        If bestRecord Is Nothing Then
            Debug.WriteLine($"[SuggestionEngine] FindBestRecord for {suggestionName}: no matching record found (records with non-empty field: " &
                $"{EnginesHistoricalSessionRecords.Count(Function(r) Not String.IsNullOrWhiteSpace(fieldSelector(r)))} of {EnginesHistoricalSessionRecords.Count})")
            Return Nothing
        End If
        If bestScore < minScore Then
            Debug.WriteLine($"[SuggestionEngine] FindBestRecord for {suggestionName}: score {bestScore:F3} below threshold {minScore} — no suggestion")
            Return Nothing
        End If
        Debug.WriteLine($"[SuggestionEngine] FindBestRecord for {suggestionName}: Bestscore={bestScore:F3}, BestRecordId={bestRecord.ID} | " &
                $"Betreff={BetreffDistances(bestIndex):F3}*{If(featureWeights.ContainsKey("Betreff"), featureWeights("Betreff"), 0)} " &
                $"Datum={DatumsDistances(bestIndex):F3}*{If(featureWeights.ContainsKey("Datum"), featureWeights("Datum"), 0)} " &
                $"Domain={AbsenderDomainDistances(bestIndex):F3}*{If(featureWeights.ContainsKey("AbsenderDomain"), featureWeights("AbsenderDomain"), 0)} " &
                $"Absender={AbsenderDistances(bestIndex):F3}*{If(featureWeights.ContainsKey("Absender"), featureWeights("Absender"), 0)} " &
                $"Benutzer={AusfueBenutzerDistances(bestIndex):F3}*{If(featureWeights.ContainsKey("AusfueBenutzer"), featureWeights("AusfueBenutzer"), 0)} " &
                $"AusfueDatum={AusfueDatumsDistances(bestIndex):F3}*{If(featureWeights.ContainsKey("AusfueDatum"), featureWeights("AusfueDatum"), 0)} " &
                $"Titel={TitelDistances(bestIndex):F3}*{If(featureWeights.ContainsKey("Titel"), featureWeights("Titel"), 0)} " &
                $"Ablageordner={AblageordnerDistances(bestIndex):F3}*{If(featureWeights.ContainsKey("Ablageordner"), featureWeights("Ablageordner"), 0)} " &
                $"ProjektPfad={ProjektPfadDistances(bestIndex):F3}*{If(featureWeights.ContainsKey("ProjektPfad"), featureWeights("ProjektPfad"), 0)} " &
                $"ProjektstrukturPfad={ProjektstrukturPfadDistances(bestIndex):F3}*{If(featureWeights.ContainsKey("ProjektstrukturPfad"), featureWeights("ProjektstrukturPfad"), 0)}")
        Return bestRecord
    End Function

    Private Function GetFeatureWeightsForProjektPfadSuggestion() As IDictionary(Of String, Double)
        Return New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
            {"Betreff", 0.4},
            {"Datum", 0.1},
            {"AbsenderDomain", 0.2},
            {"Absender", 0.2},
            {"AusfueBenutzer", 0.05},
            {"AusfueDatum", 0.2},
            {"Titel", 0},
            {"Ablageordner", 0},
            {"ProjektPfad", 0},
            {"ProjektstrukturPfad", 0}
        }
    End Function

    Private Function GetFeatureWeightsForProjektstrukturPfadSuggestion() As IDictionary(Of String, Double)
        Return New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
            {"Betreff", 0.4},
            {"Datum", 0.1},
            {"AbsenderDomain", 0.1},
            {"Absender", 0.1},
            {"AusfueBenutzer", 0.05},
            {"AusfueDatum", 0.2},
            {"Titel", 0},
            {"Ablageordner", 0},
            {"ProjektPfad", 0.2},
            {"ProjektstrukturPfad", 0}
        }
    End Function

    Private Function GetFeatureWeightsForTitelSuggestion() As IDictionary(Of String, Double)
        ' Titel and Ablageordner are always 0 at suggestion time (mutable, not yet set in cascade).
        ' ProjektstrukturPfad is 0 intentionally: it was itself suggested from the cascade (often from
        ' the same record as ProjektPfad), creating a circular self-reinforcing bias — excluded here.
        Return New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
            {"Betreff", 0.55},
            {"Datum", 0.0},
            {"AbsenderDomain", 0.1},
            {"Absender", 0.1},
            {"AusfueBenutzer", 0.0},
            {"AusfueDatum", 0.15},
            {"Titel", 0},
            {"Ablageordner", 0},
            {"ProjektPfad", 0.1},
            {"ProjektstrukturPfad", 0.0}
        }
    End Function

    Private Function GetFeatureWeightsForAblageordnerSuggestion() As IDictionary(Of String, Double)
        Return New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
            {"Betreff", 0},
            {"Datum", 0.05},
            {"AbsenderDomain", 0},
            {"Absender", 0},
            {"AusfueBenutzer", 0.05},
            {"AusfueDatum", 0.2},
            {"Titel", 0},
            {"Ablageordner", 0},
            {"ProjektPfad", 0.4},
            {"ProjektstrukturPfad", 0.45}
        }
    End Function

    Private Function GetFeatureWeightsForAbsenderKurzSuggestion() As IDictionary(Of String, Double)
        Return New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
            {"Betreff", 0},
            {"Datum", 0.05},
            {"AbsenderDomain", 0.3},
            {"Absender", 0.2},
            {"AusfueBenutzer", 0.025},
            {"AusfueDatum", 0.2},
            {"Titel", 0},
            {"Ablageordner", 0},
            {"ProjektPfad", 0.4},
            {"ProjektstrukturPfad", 0}
        }
    End Function

    Private Function GetFeatureWeightsForMsgDateinameSuggestion() As IDictionary(Of String, Double)
        Return New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
            {"Betreff", 0},
            {"Datum", 0.05},
            {"AbsenderDomain", 0},
            {"Absender", 0},
            {"AusfueBenutzer", 0.05},
            {"AusfueDatum", 0.2},
            {"Titel", 0},
            {"Ablageordner", 0},
            {"ProjektPfad", 0.4},
            {"ProjektstrukturPfad", 0.45}
        }
    End Function

    Private Function GetFeatureWeightsForAnhaengeAblegenSuggestion() As IDictionary(Of String, Double)
        Return New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
            {"Betreff", 0},
            {"Datum", 0.05},
            {"AbsenderDomain", 0},
            {"Absender", 0},
            {"AusfueBenutzer", 0.05},
            {"AusfueDatum", 0.2},
            {"Titel", 0},
            {"Ablageordner", 0},
            {"ProjektPfad", 0.4},
            {"ProjektstrukturPfad", 0.45}
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
    Private Function CalculateDateSimilarity(currentDate As DateTime, historicalDate As DateTime, maxDistanceInDays As Double) As Double
        If currentDate = Date.MinValue OrElse historicalDate = Date.MinValue Then
            Return 0.0
        End If

        Dim differenceInDays = Math.Abs((currentDate - historicalDate).TotalDays)
        Dim normalizedDistance = differenceInDays / maxDistanceInDays
        Dim similarity = 1.0 - normalizedDistance
        Return Math.Max(0.0, Math.Min(1.0, similarity))
    End Function

    ' Liefert den größten Datumsabstand zwischen aktueller Session und Historie als Normierungsbasis.
    Private Function GetMaxDateDistanceInDays(currentDate As DateTime) As Double
        If currentDate = Date.MinValue OrElse EnginesHistoricalSessionRecords.Count = 0 Then
            Return 1.0
        End If

        Dim maxDistance = EnginesHistoricalSessionRecords _
            .Where(Function(r) r.Datum <> Date.MinValue) _
            .Select(Function(r) Math.Abs((currentDate - r.Datum).TotalDays)) _
            .DefaultIfEmpty(0) _
            .Max()

        If maxDistance <= 0 Then
            Return 1.0
        End If

        Return maxDistance
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

End Class
