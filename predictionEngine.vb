Imports System.Data.SQLite
Imports System.IO
Imports Newtonsoft.Json

Module PredictionEngine
    ' Verwende den zentralen Ordner aus ThisAddIn
    Private ReadOnly dbPath As String = Path.Combine(ThisAddIn.DbDirectory, "sessions.db")
    Private ReadOnly connectionString As String = $"Data Source={dbPath};Version=3;"

    ' Dictionaries für Label Encoding
    Private labelEncoders As New Dictionary(Of String, Dictionary(Of String, Integer))

    Public Sub EncodeSessions()
        EnsureLabelEncodersLoaded()
        Dim notEncodedSessions = ThisAddIn.CurrentDatabaseManager.GetNotEncodedSessions()
        For Each sessionRecord In notEncodedSessions
            Dim encodedSession = EncodeSession(sessionRecord)
            ThisAddIn.CurrentDatabaseManager.SaveEncodedSession(encodedSession)
        Next
        SaveLabelEncoders()
    End Sub

    ' Encodiert eine einzelne SessionRecord und gibt eine EncodedSession zurück
    Public Function EncodeSession(sessionRecord As SessionRecord) As EncodedSessionRecord
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
            Else
                encodedValue = EncodeValue(prop.Name, If(value IsNot Nothing, value.ToString(), ""))
            End If
            ' Setze Property im EncodedSession-Objekt (Name muss übereinstimmen)
            Dim encodedProp = GetType(EncodedSessionRecord).GetProperty(prop.Name)
            If encodedProp IsNot Nothing Then
                encodedProp.SetValue(encodedSession, encodedValue)
            End If
        Next
        Return encodedSession
    End Function

    Private Function EncodeValue(column As String, value As String) As Integer
        If Not labelEncoders.ContainsKey(column) Then
            labelEncoders(column) = New Dictionary(Of String, Integer)()
        End If
        Dim encoder = labelEncoders(column)
        If Not encoder.ContainsKey(value) Then
            encoder(value) = encoder.Count
        End If
        Return encoder(value)
    End Function

    Private Sub SaveLabelEncoders()
        Dim encoderPath = Path.Combine(ThisAddIn.DbDirectory, "label_encoders.json")
        Dim json = JsonConvert.SerializeObject(labelEncoders)
        File.WriteAllText(encoderPath, json)
    End Sub

    ' Lädt die gespeicherten Label-Encoder aus einer JSON-Datei,
    ' falls diese existiert. Dadurch bleiben die Zuordnungen von Text zu Zahl
    ' zwischen verschiedenen Programmstarts erhalten und werden wiederverwendet.
    Private Sub EnsureLabelEncodersLoaded()
        Dim encoderPath = Path.Combine(ThisAddIn.DbDirectory, "label_encoders.json")
        If File.Exists(encoderPath) Then
            Dim json = File.ReadAllText(encoderPath)
            labelEncoders = JsonConvert.DeserializeObject(Of Dictionary(Of String, Dictionary(Of String, Integer)))(json)
        End If
    End Sub
End Module
