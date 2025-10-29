Imports System.Diagnostics

Public Class Session
    Public Property SelectedProjekt As String
    Public Property Titel As String
    Public Property Ablageordner As String
    Public Property MsgDateiname As String
    Public Property AnhaengeAblegen As Boolean
    Public Property SelectedMetadaten As New List(Of String)()
    Public Property SelectedOrdner As String ' für TreeView-Auswahl

    ' Setzt alle Properties auf Standardwerte zurück
    Public Sub Reset()
        SelectedProjekt = Nothing
        Titel = String.Empty
        Ablageordner = String.Empty
        MsgDateiname = String.Empty
        AnhaengeAblegen = False
        SelectedMetadaten.Clear()
        SelectedOrdner = Nothing
        Debug.WriteLine("[Session] Reset ausgeführt")
    End Sub

    ' Vorbereitung der Session. Mail Daten auslesen und Felder aus- und vorausfüllen
    Public Sub PrepareSession()
        ReadMailMeta()
        TreeviewEngine()
        SchemaEngine()
        Debug.WriteLine("[Session] PrepareSession ausgeführt")
    End Sub

    ' Methoden für die Verarbeitungslogik
    Public Sub ReadMailMeta()
        Debug.WriteLine("[Session] ReadMailMeta ausgeführt")
        ' TODO: Implementiere das Auslesen der E-Mail-Metadaten
    End Sub

    Public Sub TreeviewEngine()
        Debug.WriteLine("[Session] TreeviewEngine ausgeführt")
        ' TODO: Implementiere die Logik für die Verarbeitung des TreeViews
    End Sub

    Public Sub SchemaEngine()
        Debug.WriteLine("[Session] SchemaEngine ausgeführt")
        ' TODO: Implementiere die Logik für das Schema-Handling
    End Sub

    Public Sub SubmitSession()
        Debug.WriteLine("[Session] SubmitSession ausgeführt")
        ' TODO: Implementiere die Logik für das Abschicken/Speichern der Session
    End Sub

    Public Sub CancelSession()
        Debug.WriteLine("[Session] CancelSession ausgeführt")
        ' TODO: Implementiere die Logik für das Abbrechen der Session
    End Sub
End Class
