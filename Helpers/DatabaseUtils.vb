Imports System.Data.SQLite
Imports System.IO

Public Class SessionDatabaseManager
    Private ReadOnly dbPath As String
    Private ReadOnly connectionString As String

    Public Sub New()
        dbPath = Path.Combine(ThisAddIn.DbDirectory, "sessions.db")
        connectionString = $"Data Source={dbPath};Version=3;"
        EnsureDatabaseExists()
    End Sub

    Private Sub EnsureDatabaseExists()
        Dim dbDir As String = Path.GetDirectoryName(dbPath)
        If Not Directory.Exists(dbDir) Then
            Directory.CreateDirectory(dbDir)
        End If
        If Not File.Exists(dbPath) Then
            SQLiteConnection.CreateFile(dbPath)
        End If
        CreateSessionTable()
    End Sub

    Private Sub CreateSessionTable()
        Using conn As New SQLiteConnection(connectionString)
            conn.Open()
            Dim sql As String =
                "CREATE TABLE IF NOT EXISTS Sessions (" &
                "ID INTEGER PRIMARY KEY AUTOINCREMENT, " &
                "AusfueDatum TEXT, " &
                "AusfueBenutzer TEXT, " &
                "Betreff TEXT, " &
                "BetreffEmbedded BLOB, " &
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
        End Using
    End Sub

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
                    If prop.Name = "ID" Then Continue For
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

    ' Gibt die letzten vier eindeutigen ProjektPfad-Einträge für einen Benutzer zurück (absteigend nach AusfueDatum)
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
                            If result.Count = 4 Then Exit While
                        End If
                    End While
                End Using
            End Using
        End Using
        Return result
    End Function

    ' Gibt alle Sessions als Liste von SessionRecord zurück
    Public Function GetAllSessionRecords() As List(Of SessionRecord)
        Dim result As New List(Of SessionRecord)()
        Using conn As New SQLiteConnection(connectionString)
            conn.Open()
            Dim sql As String = "SELECT * FROM Sessions"
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
        Return result
    End Function

End Class

' Repräsentiert einen reinen Datenbank-Datensatz der Tabelle Sessions (ohne UI-Logik)
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
