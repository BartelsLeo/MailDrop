Imports System.Data.SQLite
Imports System.IO

Public Class SessionDatabaseManager
    Private ReadOnly dbPath As String
    Private ReadOnly connectionString As String

    Public Sub New()
        Dim appDataPath As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        dbPath = Path.Combine(appDataPath, "MailDrop", "sessions.db")
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
        ' Tabelle immer anlegen, falls sie fehlt!
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

    ' Speichert eine Session in der Datenbank
    Public Sub SaveSession(session As Session)
        Using conn As New SQLiteConnection(connectionString)
            conn.Open()
            Dim sql As String =
                "INSERT INTO Sessions (" &
                "AusfueDatum, AusfueBenutzer, Betreff, Absender, AbsenderDomain, AbsenderKurz, Empfaenger, Datum, DatumFormatiert, ProjektPfad, ProjektstrukturPfad, Titel, AblageordnerSchema, AblageordnerAufgeloest, MsgDateinameSchema, MsgDateinameAufgeloest, AnhaengeAblegen" &
                ") VALUES (" &
                "@AusfueDatum, @AusfueBenutzer, @Betreff, @Absender, @AbsenderDomain, @AbsenderKurz, @Empfaenger, @Datum, @DatumFormatiert, @ProjektPfad, @ProjektstrukturPfad, @Titel, @AblageordnerSchema, @AblageordnerAufgeloest, @MsgDateinameSchema, @MsgDateinameAufgeloest, @AnhaengeAblegen" &
                ")"
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@AusfueDatum", session.AusfueDatum)
                cmd.Parameters.AddWithValue("@AusfueBenutzer", session.AusfueBenutzer)
                cmd.Parameters.AddWithValue("@Betreff", session.Betreff)
                cmd.Parameters.AddWithValue("@Absender", session.Absender)
                cmd.Parameters.AddWithValue("@AbsenderDomain", session.AbsenderDomain)
                cmd.Parameters.AddWithValue("@AbsenderKurz", session.AbsenderKurz)
                cmd.Parameters.AddWithValue("@Empfaenger", session.Empfaenger)
                cmd.Parameters.AddWithValue("@Datum", session.Datum)
                cmd.Parameters.AddWithValue("@DatumFormatiert", session.DatumFormatiert)
                cmd.Parameters.AddWithValue("@ProjektPfad", session.ProjektPfad)
                cmd.Parameters.AddWithValue("@ProjektstrukturPfad", session.ProjektstrukturPfad)
                cmd.Parameters.AddWithValue("@Titel", session.Titel)
                cmd.Parameters.AddWithValue("@AblageordnerSchema", session.AblageordnerSchema)
                cmd.Parameters.AddWithValue("@AblageordnerAufgeloest", session.AblageordnerAufgeloest)
                cmd.Parameters.AddWithValue("@MsgDateinameSchema", session.MsgDateinameSchema)
                cmd.Parameters.AddWithValue("@MsgDateinameAufgeloest", session.MsgDateinameAufgeloest)
                cmd.Parameters.AddWithValue("@AnhaengeAblegen", If(session.AnhaengeAblegen, 1, 0))
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

End Class
