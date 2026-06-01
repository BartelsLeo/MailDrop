Imports System.Diagnostics
Imports System.Windows

Public Class SuggestionEngine

    Public Property EnginesHistoricalSessionRecords As List(Of SessionRecord)
    Private Property EnginesEmbeddingService As EmbeddingService

    Public Sub New()
        'Beim Erstellen der SuggestionEngine wird...

        'Lade alle historischen Sitzungsdaten aus der Datenbank und speichere sie in der Eigenschaft HistoricalSessionRecords
        EnginesHistoricalSessionRecords = ThisAddIn.CurrentDatabaseManager.GetAllSessionRecords()

        'Starte EmbeddingService
        EnginesEmbeddingService = New EmbeddingService()
    End Sub


    Private Class SemanticFeature

        Public Property FeaturesHistoricalSessionRecords As List(Of SessionRecord)
        Public Property FeaturesInputString As String
        Public Property FeaturesInputStringEmbedded As Single()

        Private Property FeaturesEmbeddingService As EmbeddingService

        Public Sub New(EnginesEmbeddingService As EmbeddingService, FeaturesHistoricalSessionRecords As List(Of SessionRecord))
            Me.FeaturesEmbeddingService = EnginesEmbeddingService
            Me.FeaturesHistoricalSessionRecords = FeaturesHistoricalSessionRecords
        End Sub

        Public Function Embed(InputString As String) As Single()
            FeaturesInputString = InputString.ToLower() ' Normalisieren
            FeaturesInputStringEmbedded = Me.FeaturesEmbeddingService.GenerateEmbedding(FeaturesInputString)
            Return FeaturesInputStringEmbedded
        End Function

        Public Function CalculateDistances() As List(Of Double)
            ' TODO: CosineDistance between FeaturesInputString Embedded and the embedded Betreff of each SessionRecord in FeaturesHistoricalSessionRecords
            Dim distances As New List(Of Double)(FeaturesHistoricalSessionRecords.Count)
            Dim Index As Integer = 0

            For Each record In FeaturesHistoricalSessionRecords

                'Get embedded Betreff from SessionRecord
                Dim betreffEmbedded = record.BetreffEmbedded

                'Calculate Cosine Distance between FeaturesInputStringEmbedded and betreffEmbedded
                Dim distance As Double = 0
                Dim dotProduct As Double = 0
                Dim normA As Double = 0
                Dim normB As Double = 0
                For i As Integer = 0 To betreffEmbedded.Length - 1
                    dotProduct += FeaturesInputStringEmbedded(i) * betreffEmbedded(i)
                    normA += FeaturesInputStringEmbedded(i) * FeaturesInputStringEmbedded(i)
                    normB += betreffEmbedded(i) * betreffEmbedded(i)
                Next
                normA = Math.Sqrt(normA)
                normB = Math.Sqrt(normB)
                If normA = 0 OrElse normB = 0 Then
                    distance = 0 ' Vermeidung von Division durch Null
                Else
                    distance = dotProduct / (normA * normB)
                End If

                'Assign distance to distances list
                distances(Index) = distance

                'Next Index
                Index += 1
            Next

            Return distances
        End Function

    End Class

    Private Class CategorialFeature

        Public Function CaclulateDistances() As Double
            ' TODO: Implementieren
            Return 0.0
        End Function

    End Class

    Private Class NumericalFeature

        Public Function CalculateDistances() As Double
            ' TODO: Implementieren

            ' and normalize to range [0,1] if necessary (e.g. for date difference, divide by max date difference in dataset)
            Return 0.0
        End Function

    End Class


    Public Sub EmbedBetreff(Session As Session)
        ' Function creates a Embedded vektor for Betreff
        ' Use: When Session is beeing opened, the embedded vector of Betreff needs to be created 
        ' and stored in the Session object for later use

        'Get Betreff from Session
        Dim betreff = Session.Betreff

        'Create SemanticFeature
        Dim feature As New SemanticFeature(Me.EnginesEmbeddingService, Me.EnginesHistoricalSessionRecords)

        'Embed Betreff
        Dim betreffEmbedded As Single() = feature.Embed(betreff)

        'Assign embedded Betreff to session
        Session.BetreffEmbedded = betreffEmbedded

        Debug.WriteLine($"Betreff = {betreff}")
        Debug.WriteLine($"BetreffEmbedded Länge = {betreffEmbedded.Length}")
        Debug.WriteLine($"BetreffEmbedded Erster Wert: {betreffEmbedded(0)}")

    End Sub

    Public Function SuggestProjektPfad(session As Session) As String
        ' Sub generates a project path suggestion for the given session based on the embedded Betreff and other features of the session
        ' Features to consider: Betreff, 

        Dim betreffDistances As List(Of Double) = New SemanticFeature(Me.EnginesEmbeddingService, Me.EnginesHistoricalSessionRecords).CalculateDistances()

        Dim DatumsDistances As List(Of Double) = New NumericalFeature().CalculateDistances()
        'HIER GEHTS WEITER: Implementiere die Berechnung der Distanzen für die numerischen Features (z.B. Datum) und die kategorialen Features (z.B. Abteilung)

        Dim AbsenderDomainDistances As List(Of Double) = New CategorialFeature().CalculateDistances()

        Dim AbsenderDistances As List(Of Double) = New CategorialFeature().CalculateDistances()

        Dim AusfueBenutzerDistances As List(Of Double) = New CategorialFeature().CalculateDistances()

        ' Calculate score for each historical session record based on the calculated distances and some predefined weights for each feature
        Dim Scores As New List(Of Double)(Me.EnginesHistoricalSessionRecords.Count)
        For i As Integer = 0 To Me.EnginesHistoricalSessionRecords.Count - 1
            Scores(i) = 0.4 * betreffDistances(i) + 0.2 * DatumsDistances(i) + 0.1 * AbsenderDomainDistances(i) + 0.1 * AbsenderDistances(i) + 0.2 * AusfueBenutzerDistances(i)
        Next

        Dim bestScore As Double = Double.MinValue
        Dim bestScoreIndex As Integer = -1
        Dim bestProjektPfad As String = String.Empty

        For i As Integer = 0 To Scores.Count - 1
            If Scores(i) > bestScore Then
                bestScore = Scores(i)
                bestScoreIndex = i
            End If
        Next

        ' Suggest project path based on the project path of the best matching session record
        If bestScoreIndex >= 0 Then
            bestProjektPfad = Me.EnginesHistoricalSessionRecords(bestScoreIndex).ProjektPfad
        Else
            Return String.Empty
        End If

        Debug.WriteLine($"Best Score: {bestScore} at index {bestScoreIndex}, suggesting project path: {bestProjektPfad}")

        Return bestProjektPfad
    End Function

End Class
