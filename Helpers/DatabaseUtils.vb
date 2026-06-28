Imports System.Data.SQLite
Imports System.Diagnostics
Imports System.IO

Public Class SessionDatabaseManager
    Private Const CurrentSchemaVersion As Integer = 1
    Private ReadOnly dbPath As String = Path.Combine(ThisAddIn.DbDirectory, "sessions.db")
    Private ReadOnly connectionString As String

    Public Sub New()
        Dim sw As Stopwatch = Stopwatch.StartNew()
        Debug.WriteLine("[DatabaseManager] New() BEGIN")
        dbPath = Path.Combine(ThisAddIn.DbDirectory, "sessions.db")
        connectionString = $"Data Source={dbPath};Version=3;"
        EnsureDatabaseExists()
        Debug.WriteLine($"[DatabaseManager] New() END – total: {sw.ElapsedMilliseconds} ms")
    End Sub

    Private Sub EnsureDatabaseExists()
        Dim sw As Stopwatch = Stopwatch.StartNew()
        Debug.WriteLine($"[DatabaseManager] EnsureDatabaseExists BEGIN – {dbPath}")
        Dim dbDir As String = Path.GetDirectoryName(dbPath)
        If Not Directory.Exists(dbDir) Then
            Directory.CreateDirectory(dbDir)
        End If
        If Not File.Exists(dbPath) Then
            SQLiteConnection.CreateFile(dbPath)
        End If
        Using conn As New SQLiteConnection(connectionString)
            conn.Open()

            ' Tabellen immer anlegen, falls sie fehlen.
            CreateSessionTable(conn)
            CreateComputedWeightsTable(conn)

            Dim dbVersion = GetDatabaseVersion(conn)

            ' Bestehende Datenbank ohne Version auf initiales Schema heben.
            If dbVersion = 0 Then
                SetDatabaseVersion(conn, CurrentSchemaVersion)
                dbVersion = CurrentSchemaVersion
            End If

            If dbVersion < CurrentSchemaVersion Then
                ApplyMigrations(conn, dbVersion, CurrentSchemaVersion)
            ElseIf dbVersion > CurrentSchemaVersion Then
                Throw New InvalidOperationException($"Datenbankversion {dbVersion} ist neuer als die unterstuetzte Version {CurrentSchemaVersion}.")
            End If
        End Using
        'CreateEncodedSessionsTable()
        Debug.WriteLine($"[DatabaseManager] EnsureDatabaseExists END – {sw.ElapsedMilliseconds} ms")
    End Sub

    Private Sub CreateComputedWeightsTable(conn As SQLiteConnection)
        Dim sql As String =
            "CREATE TABLE IF NOT EXISTS ComputedWeights (" &
            "TargetField TEXT NOT NULL, " &
            "FeatureName TEXT NOT NULL, " &
            "Weight REAL NOT NULL, " &
            "RecordCount INTEGER NOT NULL, " &
            "ComputedAt TEXT NOT NULL, " &
            "PRIMARY KEY (TargetField, FeatureName)" &
            ")"
        Using cmd As New SQLiteCommand(sql, conn)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Public Function GetSessionRecordCount() As Integer
        Using conn As New SQLiteConnection(connectionString)
            conn.Open()
            Using cmd As New SQLiteCommand("SELECT COUNT(*) FROM Sessions", conn)
                Return Convert.ToInt32(cmd.ExecuteScalar())
            End Using
        End Using
    End Function

    Public Function LoadComputedWeights(targetField As String) As Dictionary(Of String, Double)
        Dim result As New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase)
        Using conn As New SQLiteConnection(connectionString)
            conn.Open()
            Dim sql As String = "SELECT FeatureName, Weight FROM ComputedWeights WHERE TargetField = @targetField"
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@targetField", targetField)
                Using reader As SQLiteDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        result(reader("FeatureName").ToString()) = Convert.ToDouble(reader("Weight"))
                    End While
                End Using
            End Using
        End Using
        Return result
    End Function

    Public Sub SaveComputedWeights(targetField As String, weights As Dictionary(Of String, Double), recordCount As Integer)
        If weights Is Nothing OrElse weights.Count = 0 Then Return
        Dim computedAt As String = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
        Using conn As New SQLiteConnection(connectionString)
            conn.Open()
            Using txn As SQLiteTransaction = conn.BeginTransaction()
                Dim sql As String =
                    "INSERT OR REPLACE INTO ComputedWeights " &
                    "(TargetField, FeatureName, Weight, RecordCount, ComputedAt) " &
                    "VALUES (@targetField, @featureName, @weight, @recordCount, @computedAt)"
                For Each kvp In weights
                    Using cmd As New SQLiteCommand(sql, conn, txn)
                        cmd.Parameters.AddWithValue("@targetField", targetField)
                        cmd.Parameters.AddWithValue("@featureName", kvp.Key)
                        cmd.Parameters.AddWithValue("@weight", kvp.Value)
                        cmd.Parameters.AddWithValue("@recordCount", recordCount)
                        cmd.Parameters.AddWithValue("@computedAt", computedAt)
                        cmd.ExecuteNonQuery()
                    End Using
                Next
                txn.Commit()
            End Using
        End Using
    End Sub

    Private Sub CreateSessionTable(conn As SQLiteConnection)
        Dim sql As String =
            "CREATE TABLE IF NOT EXISTS Sessions (" &
            "ID INTEGER PRIMARY KEY AUTOINCREMENT, " &
            "AusfueDatum TEXT, " &
            "AusfueBenutzer TEXT, " &
            "Betreff TEXT, " &
            "BetreffEmbedded BLOB, " & ' Enth�lt den embeddet Betreff als Vektor
            "Absender TEXT, " &
            "AbsenderDomain TEXT, " &
            "AbsenderKurz TEXT, " &
            "Empfaenger TEXT, " &
            "Datum TEXT, " &
            "DatumFormatiert TEXT, " &
            "ProjektPfad TEXT, " &
            "ProjektstrukturPfad TEXT, " &
            "Titel TEXT, " &
            "AblageordnerSchema TEXT, " &
            "AblageordnerAufgeloest TEXT, " &
            "MsgDateinameSchema TEXT, " &
            "MsgDateinameAufgeloest TEXT, " &
            "AnhaengeAblegen BOOLEAN" &
            ")"
        Using cmd As New SQLiteCommand(sql, conn)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Function GetDatabaseVersion(conn As SQLiteConnection) As Integer
        Using cmd As New SQLiteCommand("PRAGMA user_version;", conn)
            Return Convert.ToInt32(cmd.ExecuteScalar())
        End Using
    End Function

    Private Sub SetDatabaseVersion(conn As SQLiteConnection, version As Integer)
        Using cmd As New SQLiteCommand($"PRAGMA user_version = {version};", conn)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub ApplyMigrations(conn As SQLiteConnection, fromVersion As Integer, toVersion As Integer)
        For targetVersion As Integer = fromVersion + 1 To toVersion
            Select Case targetVersion
                Case 1
                    ' Initiales Baseline-Schema. Fuer bestehende DBs ohne Version ist nichts weiter noetig.
                Case Else
                    Throw New NotSupportedException($"Keine Migration fuer Datenbankversion {targetVersion} definiert.")
            End Select

            SetDatabaseVersion(conn, targetVersion)
        Next
    End Sub

    ' Neue Methode: EncodedSessions Table anlegen
    'Private Sub CreateEncodedSessionsTable()
    '    Using conn As New SQLiteConnection(connectionString)
    '        conn.Open()
    '        Dim sql As String =
    '            "CREATE TABLE IF NOT EXISTS EncodedSessions (" &
    '            "SessionID INTEGER PRIMARY KEY, " &
    '            "AusfueBenutzer INTEGER, Betreff INTEGER, Absender INTEGER, AbsenderDomain INTEGER, AbsenderKurz INTEGER, " &
    '            "Empfaenger INTEGER, Datum INTEGER, ProjektPfad INTEGER, ProjektstrukturPfad INTEGER, Titel INTEGER, " &
    '            "AblageordnerSchema INTEGER, AblageordnerAufgeloest INTEGER, MsgDateinameSchema INTEGER, MsgDateinameAufgeloest INTEGER, " &
    '            "AnhaengeAblegen INTEGER)"
    '        Using cmd As New SQLiteCommand(sql, conn)
    '            cmd.ExecuteNonQuery()
    '        End Using
    '    End Using
    'End Sub

    ' Gibt alle Sessions zur�ck, die noch nicht encodiert wurden
    'Public Function GetNotEncodedSessions() As List(Of SessionRecord)
    '    Dim result As New List(Of SessionRecord)()
    '    Using conn As New SQLiteConnection(connectionString)
    '        conn.Open()
    '        Dim sql As String = "SELECT * FROM Sessions WHERE ID NOT IN (SELECT SessionID FROM EncodedSessions)"
    '        Using cmd As New SQLiteCommand(sql, conn)
    '            Using reader As SQLiteDataReader = cmd.ExecuteReader()
    '                While reader.Read()
    '                    Dim s As New SessionRecord With {
    '                        .ID = Convert.ToInt32(reader("ID")),
    '                        .AusfueDatum = DateTime.Parse(reader("AusfueDatum").ToString()),
    '                        .AusfueBenutzer = reader("AusfueBenutzer").ToString(),
    '                        .Betreff = reader("Betreff").ToString(),
    '                        .Absender = reader("Absender").ToString(),
    '                        .AbsenderDomain = reader("AbsenderDomain").ToString(),
    '                        .AbsenderKurz = reader("AbsenderKurz").ToString(),
    '                        .Empfaenger = reader("Empfaenger").ToString(),
    '                        .Datum = DateTime.Parse(reader("Datum").ToString()),
    '                        .DatumFormatiert = reader("DatumFormatiert").ToString(),
    '                        .ProjektPfad = reader("ProjektPfad").ToString(),
    '                        .ProjektstrukturPfad = reader("ProjektstrukturPfad").ToString(),
    '                        .Titel = reader("Titel").ToString(),
    '                        .AblageordnerSchema = reader("AblageordnerSchema").ToString(),
    '                        .AblageordnerAufgeloest = reader("AblageordnerAufgeloest").ToString(),
    '                        .MsgDateinameSchema = reader("MsgDateinameSchema").ToString(),
    '                        .MsgDateinameAufgeloest = reader("MsgDateinameAufgeloest").ToString(),
    '                        .AnhaengeAblegen = If(reader("AnhaengeAblegen") IsNot DBNull.Value AndAlso Convert.ToBoolean(reader("AnhaengeAblegen")), True, False)
    '                    }
    '                    result.Add(s)
    '                End While
    '            End Using
    '        End Using
    '    End Using
    '    Return result
    'End Function

    ' Gibt alle EncodedSessions als Liste von EncodedSessionRecord zur�ck
    'Public Function GetAllEncodedSessions() As List(Of EncodedSessionRecord)
    '    Dim result As New List(Of EncodedSessionRecord)()
    '    Using conn As New SQLiteConnection(connectionString)
    '        conn.Open()
    '        Dim sql As String = "SELECT * FROM EncodedSessions"
    '        Using cmd As New SQLiteCommand(sql, conn)
    '            Using reader As SQLiteDataReader = cmd.ExecuteReader()
    '                While reader.Read()
    '                    Dim s As New EncodedSessionRecord With {
    '                        .SessionID = If(reader("SessionID") IsNot DBNull.Value, Convert.ToInt32(reader("SessionID")), 0),
    '                        .AusfueBenutzer = If(reader("AusfueBenutzer") IsNot DBNull.Value, Convert.ToInt32(reader("AusfueBenutzer")), 0),
    '                        .Betreff = If(reader("Betreff") IsNot DBNull.Value, Convert.ToInt32(reader("Betreff")), 0),
    '                        .Absender = If(reader("Absender") IsNot DBNull.Value, Convert.ToInt32(reader("Absender")), 0),
    '                        .AbsenderDomain = If(reader("AbsenderDomain") IsNot DBNull.Value, Convert.ToInt32(reader("AbsenderDomain")), 0),
    '                        .AbsenderKurz = If(reader("AbsenderKurz") IsNot DBNull.Value, Convert.ToInt32(reader("AbsenderKurz")), 0),
    '                        .Empfaenger = If(reader("Empfaenger") IsNot DBNull.Value, Convert.ToInt32(reader("Empfaenger")), 0),
    '                        .Datum = If(reader("Datum") IsNot DBNull.Value, Convert.ToInt32(reader("Datum")), 0),
    '                        .ProjektPfad = If(reader("ProjektPfad") IsNot DBNull.Value, Convert.ToInt32(reader("ProjektPfad")), 0),
    '                        .ProjektstrukturPfad = If(reader("ProjektstrukturPfad") IsNot DBNull.Value, Convert.ToInt32(reader("ProjektstrukturPfad")), 0),
    '                        .Titel = If(reader("Titel") IsNot DBNull.Value, Convert.ToInt32(reader("Titel")), 0),
    '                        .AblageordnerSchema = If(reader("AblageordnerSchema") IsNot DBNull.Value, Convert.ToInt32(reader("AblageordnerSchema")), 0),
    '                        .AblageordnerAufgeloest = If(reader("AblageordnerAufgeloest") IsNot DBNull.Value, Convert.ToInt32(reader("AblageordnerAufgeloest")), 0),
    '                        .MsgDateinameSchema = If(reader("MsgDateinameSchema") IsNot DBNull.Value, Convert.ToInt32(reader("MsgDateinameSchema")), 0),
    '                        .MsgDateinameAufgeloest = If(reader("MsgDateinameAufgeloest") IsNot DBNull.Value, Convert.ToInt32(reader("MsgDateinameAufgeloest")), 0),
    '                        .AnhaengeAblegen = If(reader("AnhaengeAblegen") IsNot DBNull.Value, Convert.ToInt32(reader("AnhaengeAblegen")), 0)
    '                    }
    '                    result.Add(s)
    '                End While
    '            End Using
    '        End Using
    '    End Using
    '    Return result
    'End Function

    '' Speichert eine EncodedSession in der Datenbank (dynamisch per Reflection)
    'Public Sub SaveEncodedSession(encodedSession As EncodedSessionRecord)
    '    Using conn As New SQLiteConnection(connectionString)
    '        conn.Open()
    '        Dim props = GetType(EncodedSessionRecord).GetProperties()
    '        Dim colNames = String.Join(", ", props.Select(Function(p) p.Name))
    '        Dim paramNames = String.Join(", ", props.Select(Function(p) "@" & p.Name))
    '        Dim sql As String = $"INSERT INTO EncodedSessions ({colNames}) VALUES ({paramNames})"
    '        Using cmd As New SQLiteCommand(sql, conn)
    '            For Each prop In props
    '                cmd.Parameters.AddWithValue("@" & prop.Name, prop.GetValue(encodedSession))
    '            Next
    '            cmd.ExecuteNonQuery()
    '        End Using
    '    End Using
    'End Sub

    ' Speichert einen SessionRecord in der Datenbank (dynamisch per Reflection)
    Public Sub SaveSessionRecord(sessionRecord As SessionRecord)
        Using conn As New SQLiteConnection(connectionString)
            conn.Open()
            Dim props = GetType(SessionRecord).GetProperties()
            Dim colNames = String.Join(", ", props.Where(Function(p) p.Name <> "ID").Select(Function(p) p.Name))
            Dim paramNames = String.Join(", ", props.Where(Function(p) p.Name <> "ID").Select(Function(p) "@" & p.Name))
            Dim sql As String = $"INSERT INTO Sessions ({colNames}) VALUES ({paramNames})"
            Using cmd As New SQLiteCommand(sql, conn)
                For Each prop In props
                    If prop.Name = "ID" Then Continue For ' ID ist AUTOINCREMENT
                    Dim value = prop.GetValue(sessionRecord)
                    If prop.Name = "BetreffEmbedded" Then
                        value = If(value IsNot Nothing, DatabaseUtils.FloatsToBytes(CType(value, Single())), DBNull.Value)
                    End If
                    cmd.Parameters.AddWithValue("@" & prop.Name, If(value Is Nothing, DBNull.Value, value))
                Next
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ' Gibt die letzten vier eindeutigen ProjektPfad-Eintr�ge f�r einen Benutzer zur�ck (absteigend nach AusfueDatum)
    Public Function GetLastProjektVerzeichnisseForUser(benutzer As String) As List(Of String)
        Dim result As New List(Of String)()
        Using conn As New SQLiteConnection(connectionString)
            conn.Open()
            Dim sql As String = "SELECT DISTINCT ProjektPfad FROM Sessions WHERE AusfueBenutzer = @benutzer AND ProjektPfad IS NOT NULL AND ProjektPfad <> '' ORDER BY AusfueDatum DESC LIMIT 10"
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@benutzer", benutzer)
                Using reader As SQLiteDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim pfad = reader("ProjektPfad").ToString()
                        If Not String.IsNullOrWhiteSpace(pfad) AndAlso Not result.Contains(pfad) Then
                            result.Add(pfad)
                            If result.Count = 10 Then Exit While
                        End If
                    End While
                End Using
            End Using
        End Using
        Return result
    End Function

    ' Gibt alle Sessions als Liste von SessionRecord zur�ck
    Public Function GetAllSessionRecords() As List(Of SessionRecord)
        Dim sw As Stopwatch = Stopwatch.StartNew()
        Debug.WriteLine("[DatabaseManager] GetAllSessionRecords BEGIN")
        Dim result As New List(Of SessionRecord)()
        Using conn As New SQLiteConnection(connectionString)
            conn.Open()
            Debug.WriteLine($"[DatabaseManager]   DB connection opened: {sw.ElapsedMilliseconds} ms")
            Dim sql As String = "SELECT * FROM Sessions ORDER BY ID ASC"
            Using cmd As New SQLiteCommand(sql, conn)
                Using reader As SQLiteDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim s As New SessionRecord With {
                            .ID = Convert.ToInt32(reader("ID")),
                            .AusfueDatum = DateTime.Parse(reader("AusfueDatum").ToString()),
                            .AusfueBenutzer = reader("AusfueBenutzer").ToString(),
                            .Betreff = reader("Betreff").ToString(),
                            .BetreffEmbedded = If(reader("BetreffEmbedded") IsNot DBNull.Value, DatabaseUtils.BytesToFloats(CType(reader("BetreffEmbedded"), Byte())), Nothing),
                            .Absender = reader("Absender").ToString(),
                            .AbsenderDomain = reader("AbsenderDomain").ToString(),
                            .AbsenderKurz = reader("AbsenderKurz").ToString(),
                            .Empfaenger = reader("Empfaenger").ToString(),
                            .Datum = DateTime.Parse(reader("Datum").ToString()),
                            .DatumFormatiert = reader("DatumFormatiert").ToString(),
                            .ProjektPfad = reader("ProjektPfad").ToString(),
                            .ProjektstrukturPfad = reader("ProjektstrukturPfad").ToString(),
                            .Titel = reader("Titel").ToString(),
                            .AblageordnerSchema = reader("AblageordnerSchema").ToString(),
                            .AblageordnerAufgeloest = reader("AblageordnerAufgeloest").ToString(),
                            .MsgDateinameSchema = reader("MsgDateinameSchema").ToString(),
                            .MsgDateinameAufgeloest = reader("MsgDateinameAufgeloest").ToString(),
                            .AnhaengeAblegen = If(reader("AnhaengeAblegen") IsNot DBNull.Value AndAlso Convert.ToBoolean(reader("AnhaengeAblegen")), True, False)
                        }
                        result.Add(s)
                    End While
                End Using
            End Using
        End Using
        Debug.WriteLine($"[DatabaseManager] GetAllSessionRecords END – {result.Count} records in {sw.ElapsedMilliseconds} ms")
        Return result
    End Function

End Class

' Hilfsklasse f�r EncodedSession
Public Class EncodedSessionRecord
    Public Property SessionID As Integer
    Public Property AusfueBenutzer As Integer
    Public Property Betreff As Integer
    Public Property BetreffEmbedded As Single()
    Public Property Absender As Integer
    Public Property AbsenderDomain As Integer
    Public Property AbsenderKurz As Integer
    Public Property Empfaenger As Integer
    Public Property Datum As Integer ' Erg�nzt: Datum als Integer (z.B. Unix-Timestamp oder Label-encoded)
    Public Property ProjektPfad As Integer
    Public Property ProjektstrukturPfad As Integer
    Public Property Titel As Integer
    Public Property AblageordnerSchema As Integer
    Public Property AblageordnerAufgeloest As Integer
    Public Property MsgDateinameSchema As Integer
    Public Property MsgDateinameAufgeloest As Integer
    Public Property AnhaengeAblegen As Integer
End Class

' Repr�sentiert einen reinen Datenbank-Datensatz der Tabelle Sessions (ohne UI-Logik)
Public Class SessionRecord
    Public Property ID As Integer
    Public Property AusfueDatum As DateTime
    Public Property AusfueBenutzer As String
    Public Property Betreff As String
    Public Property BetreffEmbedded As Single()
    Public Property Absender As String
    Public Property AbsenderDomain As String
    Public Property AbsenderKurz As String
    Public Property Empfaenger As String
    Public Property Datum As DateTime
    Public Property DatumFormatiert As String
    Public Property ProjektPfad As String
    Public Property ProjektstrukturPfad As String
    Public Property Titel As String
    Public Property AblageordnerSchema As String
    Public Property AblageordnerAufgeloest As String
    Public Property MsgDateinameSchema As String
    Public Property MsgDateinameAufgeloest As String
    Public Property AnhaengeAblegen As Boolean
End Class

Public Class DatabaseUtils
    Public Shared Function FloatsToBytes(floats As Single()) As Byte()
        Dim bytes(floats.Length * 4 - 1) As Byte
        System.Buffer.BlockCopy(floats, 0, bytes, 0, bytes.Length)
        Return bytes
    End Function

    Public Shared Function BytesToFloats(bytes As Byte()) As Single()
        Dim floats(bytes.Length \ 4 - 1) As Single
        System.Buffer.BlockCopy(bytes, 0, floats, 0, bytes.Length)
        Return floats
    End Function
End Class
