Imports System
Imports System.Globalization
Imports System.Windows.Data

Public Class PathShortenerConverter
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        Dim path As String = TryCast(value, String)
        If String.IsNullOrEmpty(path) Then Return String.Empty
        Dim parts = path.Split({"\"c, "/"c})
        If parts.Length < 3 Then Return path
        Dim beforeSecond = String.Join("\", parts.Take(2))
        Dim lastPart = parts(parts.Length - 1)
        Return $"{beforeSecond}\...\{lastPart}"
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function
End Class
