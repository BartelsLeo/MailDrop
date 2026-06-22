Public Module ErrorHandler

    Public Sub ShowError(ex As Exception)
        System.Windows.MessageBox.Show(
            BuildErrorString(ex),
            "Fehler",
            System.Windows.MessageBoxButton.OK,
            System.Windows.MessageBoxImage.Error)
    End Sub

    Public Function BuildErrorString(ex As Exception) As String
        Dim sb As New System.Text.StringBuilder()
        Dim current As Exception = ex
        Dim level As Integer = 0
        While current IsNot Nothing
            If level = 0 Then
                sb.AppendLine(current.Message)
            Else
                sb.AppendLine(New String(" "c, level * 2) & "→ " & current.Message)
            End If
            current = current.InnerException
            level += 1
        End While
        Return sb.ToString().TrimEnd()
    End Function

End Module
