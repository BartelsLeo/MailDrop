Imports System.Diagnostics
Imports System.Linq

Public Class SuggestionEngine

    Public Property EnginesHistoricalSessionRecords As List(Of SessionRecord)
    Private ReadOnly EnginesEmbeddingService As EmbeddingService

    ' Diese Listen werden pro aktueller Session einmal berechnet und anschließend beim Suggest genutzt.
    Private Property BetreffDistances As List(Of Double)
    Private Property DatumsDistances As List(Of Double)
    Private Property AbsenderDomainDistances As List(Of Double)
    Private Property AbsenderDistances As List(Of Double)
    Private Property AusfueBenutzerDistances As List(Of Double)
    Private Property TitelDistances As List(Of Double)
    Private Property AblageordnerDistances As List(Of Double)
    Private Property ProjektPfadDistances As List(Of Double)
    Private Property ProjektstrukturPfadDistances As List(Of Double)

    Public Sub New()
        ' Lädt historische Sitzungen einmalig beim Erstellen der Engine.
        EnginesHistoricalSessionRecords = ThisAddIn.CurrentDatabaseManager.GetAllSessionRecords()

        ' Startet den Embedding-Service einmalig.
        EnginesEmbeddingService = New EmbeddingService()

        ' Initialisiert leere Distanzlisten.
        BetreffDistances = New List(Of Double)()
        DatumsDistances = New List(Of Double)()
        AbsenderDomainDistances = New List(Of Double)()
        AbsenderDistances = New List(Of Double)()
        AusfueBenutzerDistances = New List(Of Double)()
        TitelDistances = New List(Of Double)()
        AblageordnerDistances = New List(Of Double)()
        ProjektPfadDistances = New List(Of Double)()
        ProjektstrukturPfadDistances = New List(Of Double)()

        ' Stellt sicher, dass historische Datensätze ein Betreff-Embedding besitzen.
        EnsureHistoricalBetreffEmbeddings()
    End Sub

    ' Berechnet alle Distanzlisten zwischen aktueller Session und historischen Datensätzen.
    ' Diese Methode soll vor der eigentlichen Projektempfehlung aufgerufen werden.
    Public Sub RecalculateFeatureDistances(currentSession As Session)
        If currentSession Is Nothing Then
            Throw New ArgumentNullException(NameOf(currentSession))
        End If

        Dim recordCount = EnginesHistoricalSessionRecords.Count
        BetreffDistances = New List(Of Double)(recordCount)
        DatumsDistances = New List(Of Double)(recordCount)
        AbsenderDomainDistances = New List(Of Double)(recordCount)
        AbsenderDistances = New List(Of Double)(recordCount)
        AusfueBenutzerDistances = New List(Of Double)(recordCount)
        TitelDistances = New List(Of Double)(recordCount)
        AblageordnerDistances = New List(Of Double)(recordCount)
        ProjektPfadDistances = New List(Of Double)(recordCount)
        ProjektstrukturPfadDistances = New List(Of Double)(recordCount)

        If recordCount = 0 Then
            Return
        End If

        Dim currentBetreffEmbedding = GetOrCreateCurrentBetreffEmbedding(currentSession)
        Dim maxDateDistanceInDays = GetMaxDateDistanceInDays(currentSession.Datum)

        For Each record In EnginesHistoricalSessionRecords
            BetreffDistances.Add(CalculateCosineSimilarity(currentBetreffEmbedding, record.BetreffEmbedded))
            DatumsDistances.Add(CalculateDateSimilarity(currentSession.Datum, record.Datum, maxDateDistanceInDays))
            AbsenderDomainDistances.Add(CalculateCategoricalSimilarity(currentSession.AbsenderDomain, record.AbsenderDomain))
            AbsenderDistances.Add(CalculateCategoricalSimilarity(currentSession.Absender, record.Absender))
            AusfueBenutzerDistances.Add(CalculateCategoricalSimilarity(currentSession.AusfueBenutzer, record.AusfueBenutzer))
            TitelDistances.Add(CalculateTextSimilarity(currentSession.Titel, record.Titel))
            AblageordnerDistances.Add(CalculateTextSimilarity(currentSession.AblageordnerAufgeloest, record.AblageordnerAufgeloest))
            ProjektPfadDistances.Add(CalculateCategoricalSimilarity(currentSession.ProjektPfad, record.ProjektPfad))
            ProjektstrukturPfadDistances.Add(CalculateCategoricalSimilarity(currentSession.ProjektstrukturPfad, record.ProjektstrukturPfad))
        Next
    End Sub

    Public Function SuggestProjektPfad(session As Session) As String
        ' Nutzt die bereits vorberechneten Distanzlisten.
        ' Falls noch nicht vorbereitet oder Session geändert, wird hier defensiv neu berechnet.
        If session Is Nothing Then
            Return String.Empty
        End If

        ' Die Feature-Gewichte werden direkt im Suggest-Aufruf definiert,
        ' damit Scoring-Logik und Tuning-Werte gemeinsam gepflegt werden.
        Dim featureWeights As New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
            {"Betreff", 0.3},
            {"Datum", 0.15},
            {"AbsenderDomain", 0.08},
            {"Absender", 0.08},
            {"AusfueBenutzer", 0.15},
            {"Titel", 0.1},
            {"Ablageordner", 0.07},
            {"ProjektPfad", 0.04},
            {"ProjektstrukturPfad", 0.03}
        }

        If Not AreDistanceCachesComplete() Then
            RecalculateFeatureDistances(session)
        End If

        If EnginesHistoricalSessionRecords.Count = 0 Then
            Return String.Empty
        End If

        Dim bestScore As Double = Double.MinValue
        Dim bestScoreIndex As Integer = -1
        For i As Integer = 0 To EnginesHistoricalSessionRecords.Count - 1
            Dim record = EnginesHistoricalSessionRecords(i)
            If String.IsNullOrWhiteSpace(record.ProjektPfad) Then
                Continue For
            End If

            ' Bildet einen Gesamtscore aus den relevanten Features.
            Dim score =
                WeightedFeatureScore(featureWeights, "Betreff", BetreffDistances(i)) +
                WeightedFeatureScore(featureWeights, "Datum", DatumsDistances(i)) +
                WeightedFeatureScore(featureWeights, "AbsenderDomain", AbsenderDomainDistances(i)) +
                WeightedFeatureScore(featureWeights, "Absender", AbsenderDistances(i)) +
                WeightedFeatureScore(featureWeights, "AusfueBenutzer", AusfueBenutzerDistances(i)) +
                WeightedFeatureScore(featureWeights, "Titel", TitelDistances(i)) +
                WeightedFeatureScore(featureWeights, "Ablageordner", AblageordnerDistances(i)) +
                WeightedFeatureScore(featureWeights, "ProjektPfad", ProjektPfadDistances(i)) +
                WeightedFeatureScore(featureWeights, "ProjektstrukturPfad", ProjektstrukturPfadDistances(i))

            If score > bestScore Then
                bestScore = score
                bestScoreIndex = i
            End If
        Next

        ' Nimmt den Projektpfad des besten historischen Treffers.
        If bestScoreIndex >= 0 Then
            Dim bestProjektPfad = EnginesHistoricalSessionRecords(bestScoreIndex).ProjektPfad
            Debug.WriteLine($"[SuggestionEngine] Best score: {bestScore}, index: {bestScoreIndex}, ProjektPfad: {bestProjektPfad}")
            Return bestProjektPfad
        Else
            Return String.Empty
        End If
    End Function

    ' Erzeugt fehlende historische Betreff-Embeddings einmalig beim Engine-Start.
    Private Sub EnsureHistoricalBetreffEmbeddings()
        For Each record In EnginesHistoricalSessionRecords
            If record.BetreffEmbedded Is Nothing OrElse record.BetreffEmbedded.Length = 0 Then
                If String.IsNullOrWhiteSpace(record.Betreff) Then
                    Continue For
                End If
                Try
                    record.BetreffEmbedded = EnginesEmbeddingService.GenerateEmbedding(record.Betreff.ToLower())
                Catch ex As Exception
                    Debug.WriteLine($"[SuggestionEngine] Konnte historisches Betreff-Embedding nicht erstellen: {ex.Message}")
                End Try
            End If
        Next
    End Sub

    ' Erzeugt (falls nötig) das Embedding der aktuellen Session und gibt es zurück.
    Private Function GetOrCreateCurrentBetreffEmbedding(currentSession As Session) As Single()
        If currentSession.BetreffEmbedded IsNot Nothing AndAlso currentSession.BetreffEmbedded.Length > 0 Then
            Return currentSession.BetreffEmbedded
        End If
        If String.IsNullOrWhiteSpace(currentSession.Betreff) Then
            Return Nothing
        End If

        currentSession.BetreffEmbedded = EnginesEmbeddingService.GenerateEmbedding(currentSession.Betreff.ToLower())
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

    ' Prüft, ob alle Cache-Listen zur Anzahl historischer Datensätze passen.
    Private Function AreDistanceCachesComplete() As Boolean
        Dim expectedCount = EnginesHistoricalSessionRecords.Count
        Return
            BetreffDistances.Count = expectedCount AndAlso
            DatumsDistances.Count = expectedCount AndAlso
            AbsenderDomainDistances.Count = expectedCount AndAlso
            AbsenderDistances.Count = expectedCount AndAlso
            AusfueBenutzerDistances.Count = expectedCount AndAlso
            TitelDistances.Count = expectedCount AndAlso
            AblageordnerDistances.Count = expectedCount AndAlso
            ProjektPfadDistances.Count = expectedCount AndAlso
            ProjektstrukturPfadDistances.Count = expectedCount
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
            .Split(New Char() {" "c, "_"c, "-"c, "."c, "/"c, "\\"c, ","c, ";"c, ":"c, "("c, ")"c, "["c, "]"c}, StringSplitOptions.RemoveEmptyEntries) _
            .Where(Function(token) token.Trim().Length > 0)
    End Function

End Class
