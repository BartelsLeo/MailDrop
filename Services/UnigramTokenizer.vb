Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Text
Imports Microsoft.VisualBasic

' Unigram-Tokenizer (SentencePiece-Verfahren) für das deutschfähige Modell
' (gekürztes paraphrase-multilingual-MiniLM-L12-v2, siehe tools/model/).
'
' Vokabular-Format (Models/vocab.txt, erzeugt von tools/model/build_german_model.py):
' eine Zeile pro Token-ID, "<piece><TAB><score>", Zeilennummer = ID. Die ersten Zeilen sind die
' Sondertokens <s>, <pad>, </s>, <unk>.
'
' Verfahren wie der Hugging-Face-Tokenizer von XLM-R: NFKC-Normalisierung, Leerraum zu einem
' Leerzeichen zusammenfassen, jedes Wort mit "▁" (U+2581) beginnen lassen, pro Wort per Viterbi
' die Zerlegung mit maximaler Summe der Log-Wahrscheinlichkeiten wählen; nicht abgedeckte Zeichen
' werden zu <unk> (aufeinanderfolgende zusammengefasst). Gleiche Logik wie
' web/src/embedding/unigram.ts - Änderungen immer in beiden Dateien nachziehen. Die vom Build-Skript
' erzeugten Testvektoren (Models/testvectors.json) sichern beide Umsetzungen gegen den
' Original-Tokenizer ab.
Public Class UnigramTokenizer

    Public Const WordPrefix As String = ChrW(&H2581)

    Private Structure PieceEntry
        Public Id As Integer
        Public Score As Double
    End Structure

    Private ReadOnly _pieces As New Dictionary(Of String, PieceEntry)(StringComparer.Ordinal)
    Private ReadOnly _maxPieceLength As Integer
    Private ReadOnly _unkScore As Double

    Public ReadOnly Property MaxLength As Integer
    Public ReadOnly Property BosId As Integer
    Public ReadOnly Property EosId As Integer
    Public ReadOnly Property UnkId As Integer

    Public Sub New(vocabLines As IEnumerable(Of String), Optional maxLength As Integer = 128)
        Me.MaxLength = maxLength
        Dim minScore As Double = 0
        Dim maxLen As Integer = 1
        Dim id As Integer = -1
        For Each line In vocabLines
            id += 1
            If String.IsNullOrEmpty(line) Then Continue For
            Dim tab As Integer = line.LastIndexOf(ControlChars.Tab)
            If tab < 0 Then Throw New FormatException($"Vokabular-Zeile {id + 1} hat kein Tab - kein Unigram-Vokabular.")
            Dim piece As String = line.Substring(0, tab)
            Dim score As Double = Double.Parse(line.Substring(tab + 1), CultureInfo.InvariantCulture)
            If Not _pieces.ContainsKey(piece) Then _pieces(piece) = New PieceEntry With {.Id = id, .Score = score}
            If Not IsSpecial(piece) Then
                minScore = Math.Min(minScore, score)
                maxLen = Math.Max(maxLen, SplitCodePoints(piece).Count)
            End If
        Next
        _maxPieceLength = maxLen
        _unkScore = minScore - 10 ' wie Hugging Face tokenizers (Unigram)
        BosId = RequirePiece("<s>")
        EosId = RequirePiece("</s>")
        UnkId = RequirePiece("<unk>")
    End Sub

    Public Shared Function FromFile(vocabPath As String, Optional maxLength As Integer = 128) As UnigramTokenizer
        Return New UnigramTokenizer(File.ReadAllLines(vocabPath, Encoding.UTF8), maxLength)
    End Function

    ' Unigram-Vokabular = "piece<TAB>score" je Zeile; das englische BERT-Vokabular hat nur ein Token je Zeile.
    Public Shared Function LooksLikeUnigramVocab(vocabPath As String) As Boolean
        Dim firstLine As String = File.ReadLines(vocabPath, Encoding.UTF8).FirstOrDefault()
        Return firstLine IsNot Nothing AndAlso firstLine.Contains(ControlChars.Tab)
    End Function

    Public Function Encode(text As String) As List(Of Integer)
        Dim ids As New List(Of Integer) From {BosId}
        Dim budget As Integer = MaxLength - 2
        For Each word In PreTokenize(text)
            For Each pieceId In Viterbi(word)
                If ids.Count - 1 >= budget Then GoTo Done
                ids.Add(pieceId)
            Next
        Next
Done:
        ids.Add(EosId)
        Return ids
    End Function

    ' Normalisierung + Wortzerlegung: jedes Wort beginnt mit "▁", Zeichen = Unicode-Codepoints.
    Public Function PreTokenize(text As String) As List(Of List(Of String))
        Dim result As New List(Of List(Of String))
        Dim normalized As String = Normalize(text)
        If normalized.Length = 0 Then Return result
        For Each word In normalized.Split(" "c)
            result.Add(SplitCodePoints(WordPrefix & word))
        Next
        Return result
    End Function

    Private Function Viterbi(chars As List(Of String)) As List(Of Integer)
        Dim n As Integer = chars.Count
        Dim bestScore(n) As Double
        Dim bestPrev(n) As Integer
        Dim bestId(n) As Integer
        For i As Integer = 0 To n
            bestScore(i) = Double.NegativeInfinity
            bestPrev(i) = -1
            bestId(i) = -1
        Next
        bestScore(0) = 0
        Dim piece As New StringBuilder()
        For start As Integer = 0 To n - 1
            If Double.IsNegativeInfinity(bestScore(start)) Then Continue For
            Dim hasSingleChar As Boolean = False
            piece.Clear()
            Dim maxEnd As Integer = Math.Min(n, start + _maxPieceLength)
            For [end] As Integer = start + 1 To maxEnd
                piece.Append(chars([end] - 1))
                Dim key As String = piece.ToString()
                Dim entry As PieceEntry
                If Not _pieces.TryGetValue(key, entry) OrElse IsSpecial(key) Then Continue For
                If [end] = start + 1 Then hasSingleChar = True
                Dim score As Double = bestScore(start) + entry.Score
                If score > bestScore([end]) Then
                    bestScore([end]) = score
                    bestPrev([end]) = start
                    bestId([end]) = entry.Id
                End If
            Next
            If Not hasSingleChar Then
                Dim score As Double = bestScore(start) + _unkScore
                If score > bestScore(start + 1) Then
                    bestScore(start + 1) = score
                    bestPrev(start + 1) = start
                    bestId(start + 1) = UnkId
                End If
            End If
        Next
        Dim ids As New List(Of Integer)
        Dim pos As Integer = n
        While pos > 0
            ids.Add(bestId(pos))
            pos = bestPrev(pos)
        End While
        ids.Reverse()
        ' Aufeinanderfolgende <unk> zusammenfassen (fuse_unk)
        Dim fused As New List(Of Integer)(ids.Count)
        For i As Integer = 0 To ids.Count - 1
            If ids(i) = UnkId AndAlso i > 0 AndAlso ids(i - 1) = UnkId Then Continue For
            fused.Add(ids(i))
        Next
        Return fused
    End Function

    Private Function RequirePiece(piece As String) As Integer
        Dim entry As PieceEntry
        If Not _pieces.TryGetValue(piece, entry) Then Throw New FormatException($"Vokabular enthält {piece} nicht.")
        Return entry.Id
    End Function

    Private Shared Function IsSpecial(piece As String) As Boolean
        Return piece = "<s>" OrElse piece = "</s>" OrElse piece = "<pad>" OrElse piece = "<unk>" OrElse piece = "<mask>"
    End Function

    Private Shared Function SplitCodePoints(text As String) As List(Of String)
        Dim result As New List(Of String)(text.Length)
        Dim i As Integer = 0
        While i < text.Length
            If Char.IsHighSurrogate(text(i)) AndAlso i + 1 < text.Length AndAlso Char.IsLowSurrogate(text(i + 1)) Then
                result.Add(text.Substring(i, 2))
                i += 2
            Else
                result.Add(text(i).ToString())
                i += 1
            End If
        End While
        Return result
    End Function

    ' Muss web/src/embedding/unigram.ts normalize() entsprechen: NFKC, unsichtbare Zeichen verwerfen,
    ' Leerraum vereinheitlichen, Steuerzeichen verwerfen, Leerzeichenfolgen zusammenfassen, trimmen.
    Public Shared Function Normalize(text As String) As String
        If String.IsNullOrEmpty(text) Then Return String.Empty
        Dim sb As New StringBuilder(text.Length)
        For Each ch In SplitCodePoints(text.Normalize(NormalizationForm.FormKC))
            Dim cp As Integer = Char.ConvertToUtf32(ch, 0)
            If cp = &HFEFF OrElse cp = &H200B Then Continue For
            ' U+0085 ist in .NET Leerraum, in JavaScript (\s) nicht - dort fällt es unter die Steuerzeichen.
            If cp <> &H85 AndAlso ch.Length = 1 AndAlso Char.IsWhiteSpace(ch(0)) Then
                sb.Append(" "c)
            ElseIf cp < &H20 OrElse (cp >= &H7F AndAlso cp < &HA0) Then
                Continue For
            Else
                sb.Append(ch)
            End If
        Next
        Dim collapsed As String = sb.ToString()
        While collapsed.Contains("  ")
            collapsed = collapsed.Replace("  ", " ")
        End While
        Return collapsed.Trim(" "c)
    End Function

End Class
