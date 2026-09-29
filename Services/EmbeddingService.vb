Imports System.IO
Imports System.Diagnostics
Imports System.Security.Cryptography
Imports Microsoft.ML.OnnxRuntime
Imports Microsoft.ML.OnnxRuntime.Tensors
Imports Microsoft.ML.Tokenizers

' Modell und Tokenizer richten sich nach Models/vocab.txt:
' - Zeilen "piece<TAB>score" = deutschfähiges Modell (gekürztes paraphrase-multilingual-MiniLM-L12-v2,
'   gebaut von tools/model/build_german_model.py) -> UnigramTokenizer
' - sonst das bisherige englische all-MiniLM-Vokabular -> BertTokenizer (Microsoft.ML.Tokenizers)
' Das englische Modell bleibt nur als Rückfall, bis das deutschfähige Modell gebaut und eingecheckt ist.
Public Class EmbeddingService
    Implements IDisposable

    Private ReadOnly _session As InferenceSession
    Private ReadOnly _unigramTokenizer As UnigramTokenizer
    Private ReadOnly _bertTokenizer As BertTokenizer
    Private ReadOnly _usesTokenTypeIds As Boolean

    ' SHA-256 von model.onnx. Gespeicherte Betreff-Embeddings (Sessions.BetreffEmbedded) sind nur mit
    ' Embeddings desselben Modells vergleichbar; SuggestionEngine.EnsureHistoryEmbeddingsMatchModel
    ' berechnet sie neu, wenn sich diese ID gegenüber der Datenbank geändert hat.
    Public ReadOnly Property ModelId As String
    Private _disposed As Boolean = False

    Public Sub New()
        Dim sw As Stopwatch = Stopwatch.StartNew()
        Debug.WriteLine("[EmbeddingService] New() BEGIN")
        Dim codeBase As String = System.Reflection.Assembly.GetExecutingAssembly().CodeBase
        Dim uri As New Uri(codeBase)
        Dim baseDir As String = System.IO.Path.GetDirectoryName(uri.LocalPath)
        Dim modelPath As String = System.IO.Path.Combine(baseDir, "Models", "model.onnx")
        Dim vocabPath As String = System.IO.Path.Combine(baseDir, "Models", "vocab.txt")
        Debug.WriteLine($"[EmbeddingService]   Model path resolved: {sw.ElapsedMilliseconds} ms – {modelPath}")
        _session = New InferenceSession(modelPath)
        Using stream = File.OpenRead(modelPath), sha = SHA256.Create()
            ModelId = BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant()
        End Using
        Debug.WriteLine($"[EmbeddingService]   InferenceSession loaded: {sw.ElapsedMilliseconds} ms")
        ' Das multilinguale Modell (XLM-R-Tokenizer) hat keinen token_type_ids-Eingang.
        _usesTokenTypeIds = _session.InputMetadata.ContainsKey("token_type_ids")
        If UnigramTokenizer.LooksLikeUnigramVocab(vocabPath) Then
            _unigramTokenizer = UnigramTokenizer.FromFile(vocabPath)
            Debug.WriteLine($"[EmbeddingService]   UnigramTokenizer created (deutschfähiges Modell): {sw.ElapsedMilliseconds} ms")
        Else
            _bertTokenizer = BertTokenizer.Create(vocabPath)
            Debug.WriteLine($"[EmbeddingService]   BertTokenizer created (englisches Modell): {sw.ElapsedMilliseconds} ms")
        End If
        Debug.WriteLine($"[EmbeddingService] New() END – total: {sw.ElapsedMilliseconds} ms")
    End Sub

    Public Function GenerateEmbedding(text As String) As Single()
        Dim sw As Stopwatch = Stopwatch.StartNew()

        ' 1. Tokenisieren. Das englische Modell ist "uncased" (Kleinschreibung wie bisher), das
        ' deutschfähige unterscheidet Groß-/Kleinschreibung - dort den Text unverändert lassen.
        Dim encoding As IEnumerable(Of Integer) = If(_unigramTokenizer IsNot Nothing,
                                                     _unigramTokenizer.Encode(text),
                                                     _bertTokenizer.EncodeToIds(If(text, String.Empty).ToLower()))
        Dim tokenIds = encoding.Select(Function(id) CLng(id)).ToArray()
        Dim attentionMask = Enumerable.Repeat(1L, tokenIds.Length).ToArray()
        Dim tokenTypeIds = Enumerable.Repeat(0L, tokenIds.Length).ToArray()

        Dim seqLen As Integer = tokenIds.Length

        ' 2. Tensoren erstellen
        Dim inputIdsTensor = New DenseTensor(Of Long)(tokenIds, New Integer() {1, seqLen})
        Dim attentionTensor = New DenseTensor(Of Long)(attentionMask, New Integer() {1, seqLen})
        Dim tokenTypeTensor = New DenseTensor(Of Long)(tokenTypeIds, New Integer() {1, seqLen})

        ' 3. ONNX Input
        Dim inputs As New List(Of NamedOnnxValue) From {
            NamedOnnxValue.CreateFromTensor("input_ids", inputIdsTensor),
            NamedOnnxValue.CreateFromTensor("attention_mask", attentionTensor)
        }
        If _usesTokenTypeIds Then inputs.Add(NamedOnnxValue.CreateFromTensor("token_type_ids", tokenTypeTensor))

        ' 4. Inferenz
        Dim swInfer As Stopwatch = Stopwatch.StartNew()
        Using results = _session.Run(inputs)
            Debug.WriteLine($"[EmbeddingService] GenerateEmbedding: tokenize+run={sw.ElapsedMilliseconds} ms (inference={swInfer.ElapsedMilliseconds} ms), tokens={seqLen}, text='{If(text?.Length > 40, text.Substring(0, 40) & "…", text)}'")
            Dim output = results.First().AsTensor(Of Single)()

            ' 5. Mean Pooling über alle Token (beide Modelle: 384 Dimensionen)
            Dim hidden As Integer = CInt(output.Length \ seqLen) ' Ausgabe [1, seqLen, hidden]
            Dim embedding(hidden - 1) As Single
            For tokenIdx As Integer = 0 To seqLen - 1
                For d As Integer = 0 To hidden - 1
                    embedding(d) += output(0, tokenIdx, d)
                Next
            Next
            For d As Integer = 0 To hidden - 1
                embedding(d) /= seqLen
            Next

            Return Normalize(embedding)
        End Using

    End Function

    ' L2-Normalisierung
    Private Function Normalize(vector As Single()) As Single()
        Dim norm As Double = Math.Sqrt(vector.Sum(Function(x) CDbl(x) * x))
        Return vector.Select(Function(x) CSng(x / norm)).ToArray()
    End Function

    Public Sub Dispose() Implements IDisposable.Dispose
        If Not _disposed Then
            _session?.Dispose()
            _disposed = True
        End If
    End Sub

End Class