Imports System.Diagnostics

Public Class SuggestionEngine

    Public Property HistoricalSessionRecords As List(Of SessionRecord)

    Public Sub New()
        'Beim Erstellen der SuggestionEngine wird...

        'Lade alle historischen Sitzungsdaten aus der Datenbank und speichere sie in der Eigenschaft HistoricalSessionRecords
        HistoricalSessionRecords = ThisAddIn.CurrentDatabaseManager.GetAllSessionRecords()
    End Sub


    Private Class SemanticFeature

        Public Property InputString As String
        Public Property InputStringEmbedded As Single()

        Public Function Embed(inputString As String) As Single()
            ' TODO: Implementieren mit all-MiniLM-L6-v2
            Return Array.Empty(Of Single)()
        End Function

        Public Function CosineDistance() As Double
            ' TODO: Implementieren
            Return 0.0
        End Function

    End Class

    Private Class CategorialFeature

        Public Function MatchDistance() As Double
            ' TODO: Implementieren
            Return 0.0
        End Function

    End Class

    Private Class NumericalFeature

        Public Function Difference() As Double
            ' TODO: Implementieren
            Return 0.0
        End Function

        Public Function Normalize() As Double
            ' TODO: Implementieren
            Return 0.0
        End Function

    End Class


    Public Function EmbedBetreff(Session As Session)
        ' Function creates a Embedded vektor for Betreff
        ' Use: When Session is beeing opened, the embedded vector of Betreff needs to be created 
        Return 0.0
    End Function

    Public Sub SuggestProjektPfad(session As Session)
        Dim ausfueBenutzer = session.AusfueBenutzer
        Dim absender = session.Absender
        Dim absenderDomain = session.AbsenderDomain
        Dim datum = session.Datum
        Dim betreff = session.Betreff

        Dim service As New EmbeddingService()
        Dim betreffEmbedded As Single() = service.GenerateEmbedding(betreff)

        Debug.WriteLine($"Embedding Länge: {betreffEmbedded.Length}")
        Debug.WriteLine($"Erster Wert: {betreffEmbedded(0)}")

    End Sub

End Class
