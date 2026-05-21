Imports System.Data.SQLite
Imports System.IO
Imports Newtonsoft.Json
Imports Accord.MachineLearning.DecisionTrees
Imports Accord.MachineLearning
Imports Accord.Statistics.Filters
Imports Accord.Math
Imports System.Diagnostics

Module PredictionEngine
    ' Verwende den zentralen Ordner aus ThisAddIn
    Private ReadOnly dbPath As String = Path.Combine(ThisAddIn.DbDirectory, "sessions.db")
    Private ReadOnly connectionString As String = $"Data Source={dbPath};Version=3;"

    ' Dictionaries für Label Encoding
    Private labelEncoders As New Dictionary(Of String, Dictionary(Of String, Integer))

    Private Sub EncodeSessions()
        Debug.WriteLine("[PredictionEngine] Starte EncodeSessions...")
        EnsureLabelEncodersLoaded()
        Dim notEncodedSessions = ThisAddIn.CurrentDatabaseManager.GetNotEncodedSessions()
        Debug.WriteLine($"[PredictionEngine] Nicht encodierte Sessions gefunden: {notEncodedSessions.Count}")
        For Each sessionRecord In notEncodedSessions
            Dim encodedSession = EncodeSession(sessionRecord)
            ThisAddIn.CurrentDatabaseManager.SaveEncodedSession(encodedSession)
        Next
        SaveLabelEncoders()
        Debug.WriteLine("[PredictionEngine] EncodeSessions abgeschlossen.")
    End Sub

    ' Encodiert eine einzelne SessionRecord und gibt eine EncodedSession zurück
    Public Function EncodeSession(sessionRecord As SessionRecord) As EncodedSessionRecord
        Debug.WriteLine($"[PredictionEngine] Encodiere SessionRecord mit ID: {sessionRecord.ID}")
        Dim encodedSession As New EncodedSessionRecord()
        encodedSession.SessionID = sessionRecord.ID

        ' Alle Properties dynamisch per Loop encodieren (außer ID)
        For Each prop In GetType(SessionRecord).GetProperties()
            If prop.Name = "ID" Then Continue For
            Dim value = prop.GetValue(sessionRecord)
            Dim encodedValue As Integer
            If prop.PropertyType Is GetType(Boolean) Then
                encodedValue = If(CBool(value), 1, 0)
            ElseIf prop.PropertyType Is GetType(Integer) Then
                encodedValue = CInt(value)
            ElseIf prop.PropertyType Is GetType(DateTime) AndAlso prop.Name = "Datum" Then
                ' Datum als Unix-Timestamp speichern
                encodedValue = CInt(DateTimeOffset.Parse(value.ToString()).ToUnixTimeSeconds())
            Else
                encodedValue = EncodeValue(prop.Name, If(value IsNot Nothing, value.ToString(), ""))
            End If
            ' Setze Property im EncodedSession-Objekt (Name muss übereinstimmen)
            Dim encodedProp = GetType(EncodedSessionRecord).GetProperty(prop.Name)
            If encodedProp IsNot Nothing Then
                encodedProp.SetValue(encodedSession, encodedValue)
            End If
        Next
        Debug.WriteLine($"[PredictionEngine] SessionRecord ID {sessionRecord.ID} encodiert.")
        Return encodedSession
    End Function

    Private Function EncodeValue(column As String, value As String) As Integer
        If Not labelEncoders.ContainsKey(column) Then
            labelEncoders(column) = New Dictionary(Of String, Integer)()
        End If
        Dim encoder = labelEncoders(column)
        If Not encoder.ContainsKey(value) Then
            encoder(value) = encoder.Count
            Debug.WriteLine($"[PredictionEngine] Neuer LabelEncoder für '{column}': '{value}' -> {encoder(value)}")
        End If
        Return encoder(value)
    End Function

    Private Sub SaveLabelEncoders()
        Dim encoderPath = Path.Combine(ThisAddIn.DbDirectory, "label_encoders.json")
        Dim json = JsonConvert.SerializeObject(labelEncoders)
        File.WriteAllText(encoderPath, json)
        Debug.WriteLine($"[PredictionEngine] LabelEncoder gespeichert: {encoderPath}")
    End Sub

    ' Lädt die gespeicherten Label-Encoder aus einer JSON-Datei,
    ' falls diese existiert. Dadurch bleiben die Zuordnungen von Text zu Zahl
    ' zwischen verschiedenen Programmstarts erhalten und werden wiederverwendet.
    Private Sub EnsureLabelEncodersLoaded()
        Dim encoderPath = Path.Combine(ThisAddIn.DbDirectory, "label_encoders.json")
        If File.Exists(encoderPath) Then
            Dim json = File.ReadAllText(encoderPath)
            labelEncoders = JsonConvert.DeserializeObject(Of Dictionary(Of String, Dictionary(Of String, Integer)))(json)
            Debug.WriteLine($"[PredictionEngine] LabelEncoder geladen: {encoderPath}")
        Else
            Debug.WriteLine($"[PredictionEngine] Kein LabelEncoder gefunden, neuer wird erstellt.")
        End If
    End Sub

    ' Trainiert für jeden Output-Parameter ein Decision Tree Modell (angepasste Property-Namen)
    Public Sub TrainDecisionTreeModels()
        Debug.WriteLine("[PredictionEngine] Starte Training der Decision Tree Modelle...")
        ' 1. Encodiere alle Sessions, damit sie als Integer-Werte vorliegen
        EncodeSessions()
        ' 2. Hole alle encodierten Sessions aus der Datenbank
        Dim encodedSessions = ThisAddIn.CurrentDatabaseManager.GetAllEncodedSessions()
        Debug.WriteLine($"[PredictionEngine] Anzahl encodierter Sessions: {encodedSessions.Count}")
        ' 3. Definiere die Input- und Output-Properties für das Modell
        Dim inputProperties As String() = {"AusfueBenutzer", "Betreff", "Absender", "AbsenderDomain", "Empfaenger", "Datum"}
        Dim outputProperties As String() = {"ProjektPfad"} ', "ProjektstrukturPfad", "Titel", "AblageordnerSchema", "MsgDateinameSchema", "AnhaengeAblegen"}
        ' 4. Erzeuge die Input-Matrix für das Training (nur Integer-Werte)
        Dim inputs = encodedSessions.Select(Function(s) inputProperties.Select(Function(p) CInt(s.GetType().GetProperty(p).GetValue(s))).ToArray()).ToArray()
        ' 5. Entferne konstante Spalten aus den Inputs, da sie für das Training nicht relevant sind
        Dim columnsToKeep As New List(Of Integer)
        For i = 0 To inputProperties.Length - 1
            Dim colIndex = i ' Lokale Kopie!
            Dim firstValue = inputs(0)(colIndex)
            If Not inputs.All(Function(row) row(colIndex) = firstValue) Then
                columnsToKeep.Add(colIndex)
            Else
                Debug.WriteLine($"[PredictionEngine] Spalte '{inputProperties(colIndex)}' ist konstant und wird für das Training entfernt.")
            End If
        Next
        Dim filteredInputs = inputs.Select(Function(row) columnsToKeep.Select(Function(i) row(i)).ToArray()).ToArray()
        ' 5.1 Prüfe die Input-Matrix nur einmal, da sie sich im Loop nicht ändert
        If filteredInputs.Length < 2 Then
            Debug.WriteLine("[PredictionEngine] Training abgebrochen: Zu wenig Daten (weniger als 2 Zeilen) für das Training.")
            Return ' Schleife wird nicht ausgeführt
        End If
        If filteredInputs(0).Length < 1 Then
            Debug.WriteLine("[PredictionEngine] Training abgebrochen: Keine Input-Spalten für das Training.")
            Return ' Schleife wird nicht ausgeführt
        End If
        ' Debug-Ausgabe für die Input-Matrix (gilt für alle Outputs)
        Debug.WriteLine($"[PredictionEngine] Inputs: {filteredInputs.Length} Zeilen, {If(filteredInputs.Length > 0, filteredInputs(0).Length, 0)} Spalten")
        ' Debug: Zeige alle Werte der InputProperties im SessionRecord vor dem Encoding
        For Each sessionRecord In ThisAddIn.CurrentDatabaseManager.GetAllSessionRecords()
            Dim values = inputProperties.Select(Function(p) sessionRecord.GetType().GetProperty(p)?.GetValue(sessionRecord)).ToArray()
            Debug.WriteLine($"[Debug] SessionRecord ID: {sessionRecord.ID}, Inputs: {String.Join(", ", values)}")
        Next
        ' Debug: Zeige alle Werte der encodierten Sessions für die InputProperties
        For Each s In encodedSessions
            Dim values = inputProperties.Select(Function(p) s.GetType().GetProperty(p)?.GetValue(s)).ToArray()
            Debug.WriteLine($"[Debug] EncodedSession ID: {s.SessionID}, Inputs: {String.Join(", ", values)}")
        Next
        ' 6. Trainiere für jeden Output-Parameter ein Decision Tree Modell
        For Each outputProp In outputProperties
            ' 6.1 Erzeuge die Output-Werte für das aktuelle Modell
            Dim outputs = encodedSessions.Select(Function(s) CInt(s.GetType().GetProperty(outputProp).GetValue(s))).ToArray()
            ' Debug-Ausgabe für die Output-Werte (pro Output-Parameter)
            Debug.WriteLine($"[PredictionEngine] Outputs: {outputs.Length} Werte, Distinct: {outputs.Distinct().Count()}")
            Dim modelPath = Path.Combine(ThisAddIn.DbDirectory, $"DecisionTree_{outputProp}.bin")
            Dim constPath = Path.Combine(ThisAddIn.DbDirectory, $"DecisionTree_{outputProp}.txt")
            ' 6.2 Prüfe, ob die Output-Länge zu den Inputs passt
            If outputs.Length <> filteredInputs.Length Then
                Debug.WriteLine($"[PredictionEngine] Anzahl Inputs und Outputs stimmt nicht überein für Output '{outputProp}', Training wird übersprungen.")
                If File.Exists(modelPath) Then File.Delete(modelPath)
                If File.Exists(constPath) Then File.Delete(constPath)
                Continue For
            End If
            ' 6.3 Falls der Output konstant ist, speichere ihn als Text und lösche ggf. Modell
            If outputs.Distinct().Count() < 2 Then
                Dim konst = outputs(0).ToString()
                File.WriteAllText(constPath, konst)
                If File.Exists(modelPath) Then File.Delete(modelPath)
                Debug.WriteLine($"[PredictionEngine] Output '{outputProp}' ist konstant, schreibe {constPath} mit Wert {konst} und lösche ggf. Modell.")
                Continue For
            End If
            Try
                ' 6.4 Starte das Training des Decision Tree Modells
                Debug.WriteLine($"[PredictionEngine] Training Decision Tree für Output: {outputProp}")
                Debug.WriteLine($"[PredictionEngine] OutputProp: {outputProp}")
                Debug.WriteLine($"[PredictionEngine] InputProperties: {String.Join(", ", columnsToKeep.Select(Function(i) inputProperties(i)))}")
                Debug.WriteLine($"[PredictionEngine] Erster Input: {String.Join(", ", filteredInputs(0))}")
                Debug.WriteLine($"[PredictionEngine] Erster Output: {outputs(0)}")
                ' Decision Tree Training
                Dim attributes = columnsToKeep.Select(Function(i) New DecisionVariable(inputProperties(i), DecisionVariableKind.Discrete)).ToArray()
                Dim tree = New DecisionTree(attributes, outputs.Distinct().Count())
                Dim teacher = New ID3Learning(tree)
                teacher.Run(filteredInputs, outputs)
                ' 6.6 Speichere das trainierte Modell
                Accord.IO.Serializer.Save(tree, modelPath)
                If File.Exists(constPath) Then File.Delete(constPath)
                Debug.WriteLine($"[PredictionEngine] Modell gespeichert: {modelPath}")
            Catch ex As Exception
                Debug.WriteLine($"[PredictionEngine] Fehler beim Training für Output '{outputProp}': {ex.ToString()}")
            End Try
        Next
        Debug.WriteLine("[PredictionEngine] Training abgeschlossen.")
    End Sub

    ' Gibt für eine neue Session Vorhersagen für alle Output-Parameter zurück und schreibt sie direkt in das Session-Objekt
    Public Sub PredictOutput(session As Session)
        Debug.WriteLine("[PredictionEngine] Starte PredictOutput...")
        Dim inputProperties As String() = {"AusfueBenutzer", "Betreff", "Absender", "AbsenderDomain", "Empfaenger", "Datum"}
        EnsureLabelEncodersLoaded()
        Dim encodedSession = EncodeSession(session.ToSessionRecord())
        Dim inputValues = inputProperties.Select(Function(p) CInt(encodedSession.GetType().GetProperty(p).GetValue(encodedSession))).ToArray()
        Debug.WriteLine($"[PredictionEngine] Eingabewerte: {String.Join(", ", inputValues)}")
        Dim outputProps = New List(Of String) From {
            "ProjektPfad",
            "ProjektstrukturPfad",
            "AbsenderKurz",
            "AblageordnerSchema",
            "MsgDateinameSchema"
        }
        For Each outputProp In outputProps
            Dim modelPath = Path.Combine(ThisAddIn.DbDirectory, $"DecisionTree_{outputProp}.bin")
            Dim constPath = Path.Combine(ThisAddIn.DbDirectory, $"DecisionTree_{outputProp}.txt")
            Dim prop = session.GetType().GetProperty(outputProp)
            If prop Is Nothing OrElse Not prop.CanWrite Then
                Debug.WriteLine($"[PredictionEngine] Property {outputProp} nicht beschreibbar oder nicht vorhanden.")
                Continue For
            End If
            If File.Exists(constPath) Then
                Dim konst = File.ReadAllText(constPath)
                If Not String.IsNullOrWhiteSpace(konst) AndAlso konst <> "0" Then
                    prop.SetValue(session, konst)
                    Debug.WriteLine($"[PredictionEngine] {outputProp} aus Konstante ({konst}) gesetzt.")
                Else
                    Debug.WriteLine($"[PredictionEngine] {outputProp} Konstante existiert, aber leer oder 0.")
                End If
            ElseIf File.Exists(modelPath) Then
                Dim tree = Accord.IO.Serializer.Load(Of DecisionTree)(modelPath)
                Dim prediction = tree.Decide(inputValues)
                Debug.WriteLine($"[PredictionEngine] {outputProp} Vorhersagewert: {prediction}")
                If prediction <> 0 Then
                    prop.SetValue(session, prediction.ToString())
                    Debug.WriteLine($"[PredictionEngine] {outputProp} vorhergesagt und gesetzt: {prediction}")
                Else
                    Debug.WriteLine($"[PredictionEngine] {outputProp} Vorhersagewert ist 0, nicht gesetzt.")
                End If
            Else
                Debug.WriteLine($"[PredictionEngine] Kein Modell/Konstante für {outputProp} vorhanden.")
            End If
        Next
        ' AnhaengeAblegen (Boolean)
        Dim anhaengeModel = Path.Combine(ThisAddIn.DbDirectory, "RandomForest_AnhaengeAblegen.bin")
        Dim anhaengeConst = Path.Combine(ThisAddIn.DbDirectory, "RandomForest_AnhaengeAblegen.txt")
        If File.Exists(anhaengeConst) Then
            Dim konst = File.ReadAllText(anhaengeConst)
            session.AnhaengeAblegen = (konst = "1")
            Debug.WriteLine($"[PredictionEngine] AnhaengeAblegen aus Konstante gesetzt: {konst}")
        ElseIf File.Exists(anhaengeModel) Then
            Dim forest = Accord.IO.Serializer.Load(Of RandomForest)(anhaengeModel)
            Dim prediction = forest.Decide(inputValues)
            session.AnhaengeAblegen = (prediction = 1)
            Debug.WriteLine($"[PredictionEngine] AnhaengeAblegen vorhergesagt: {prediction}")
        Else
            Debug.WriteLine("[PredictionEngine] Kein Modell/Konstante für AnhaengeAblegen vorhanden.")
        End If
        Debug.WriteLine("[PredictionEngine] PredictOutput abgeschlossen.")
    End Sub
End Module
