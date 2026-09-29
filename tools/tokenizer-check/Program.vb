Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text.Json

' Gleiche Fälle wie web/test/unigram.test.ts, damit VB- und TypeScript-Tokenizer nachweislich
' gleich arbeiten. Mit Models/testvectors.json (vom Modell-Build) zusätzlich Abgleich der Token-IDs
' mit dem Original-Tokenizer von Hugging Face. Rückgabecode 0 = alles bestanden.
Module Program

    Private _failures As Integer = 0

    Function Main(args As String()) As Integer
        Dim vocab As String() = {
            "<s>" & vbTab & "0", "<pad>" & vbTab & "0", "</s>" & vbTab & "0", "<unk>" & vbTab & "0",
            "▁" & vbTab & "-2", "▁an" & vbTab & "-3", "▁ange" & vbTab & "-4", "bot" & vbTab & "-3", "▁angebot" & vbTab & "-5",
            "a" & vbTab & "-6", "n" & vbTab & "-6", "g" & vbTab & "-6", "e" & vbTab & "-6", "b" & vbTab & "-6", "o" & vbTab & "-6", "t" & vbTab & "-6",
            "▁pr" & vbTab & "-4", "ü" & vbTab & "-5", "fung" & vbTab & "-4", "▁prüfung" & vbTab & "-6.5"}
        Dim tok As New UnigramTokenizer(vocab)
        Dim pieces = Function(ids As List(Of Integer)) String.Join(" ", ids.Select(Function(id) vocab(id).Split(ControlChars.Tab)(0)))

        Check("höchste Gesamtwahrscheinlichkeit", pieces(tok.Encode("angebot")), "<s> ▁angebot </s>")
        Check("cased", pieces(tok.Encode("Angebot")).Contains("▁angebot").ToString(), "False")
        Check("ganzes Wort günstiger", pieces(tok.Encode("prüfung")), "<s> ▁prüfung </s>")
        Check("<unk> zusammengefasst", pieces(tok.Encode("an xyz")), "<s> ▁an ▁ <unk> </s>")
        Check("maxLength", New UnigramTokenizer(vocab, 4).Encode("angebot angebot angebot").Count.ToString(), "4")
        Check("Leerraum", UnigramTokenizer.Normalize("  a" & vbTab & vbTab & "b c" & ChrW(&HFEFF) & " "), "a b c")
        Check("NFKC", UnigramTokenizer.Normalize(ChrW(&HFB01)), "fi")

        Dim modelsDir As String = If(args.Length > 0, args(0), Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "Models"))
        Dim vectorsPath As String = Path.Combine(modelsDir, "testvectors.json")
        If File.Exists(vectorsPath) Then
            Dim real = UnigramTokenizer.FromFile(Path.Combine(modelsDir, "vocab.txt"))
            Using doc = JsonDocument.Parse(File.ReadAllText(vectorsPath))
                Dim count As Integer = 0
                For Each v In doc.RootElement.EnumerateArray()
                    Dim text As String = v.GetProperty("text").GetString()
                    Dim expected = String.Join(",", v.GetProperty("ids").EnumerateArray().Select(Function(e) e.GetInt32()))
                    Check($"Prüfvektor '{text}'", String.Join(",", real.Encode(text)), expected)
                    count += 1
                Next
                Console.WriteLine($"{count} Prüfvektoren geprüft.")
            End Using
        Else
            Console.WriteLine($"Keine Prüfvektoren unter {vectorsPath} - nur künstliches Vokabular geprüft.")
        End If

        Console.WriteLine(If(_failures = 0, "Alle Prüfungen bestanden.", $"{_failures} Prüfung(en) fehlgeschlagen."))
        Return If(_failures = 0, 0, 1)
    End Function

    Private Sub Check(name As String, actual As String, expected As String)
        If actual = expected Then Return
        _failures += 1
        Console.WriteLine($"FEHLER {name}: erwartet [{expected}], erhalten [{actual}]")
    End Sub

End Module
