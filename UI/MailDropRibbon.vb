Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core

<ComVisible(True)>
Public Class MailDropRibbon
    Implements IRibbonExtensibility

    Private ribbon As IRibbonUI

    ' Outlook fragt GetCustomUI einmal pro Ribbon-Typ ab. Der Button erscheint in der
    ' Explorer-Ansicht (Reiter "Start") und im geoeffneten Mail-Fenster (Reiter "Nachricht");
    ' Verfassen-/Termin-/Kontakt-Fenster usw. bekommen kein MailDrop-Ribbon.
    Public Function GetCustomUI(ribbonID As String) As String Implements IRibbonExtensibility.GetCustomUI
        Dim resourceName As String
        Select Case ribbonID
            Case "Microsoft.Outlook.Explorer"
                resourceName = "MailDrop.MailDropRibbon.xml"
            Case "Microsoft.Outlook.Mail.Read"
                resourceName = "MailDrop.MailDropRibbonInspector.xml"
            Case Else
                Return Nothing
        End Select
        Dim asm = System.Reflection.Assembly.GetExecutingAssembly()
        Using stream = asm.GetManifestResourceStream(resourceName)
            If stream Is Nothing Then
                Throw New InvalidOperationException($"Embedded ribbon resource '{resourceName}' not found in assembly.")
            End If
            Using reader As New System.IO.StreamReader(stream)
                Return reader.ReadToEnd()
            End Using
        End Using
    End Function

    Public Sub MailAblegen_Click(control As IRibbonControl)
        ' Delegiert an ThisAddIn
        Globals.ThisAddIn.MailAblegen_Click(control)
    End Sub

    Public Sub OnLoad(ribbonUI As IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub
End Class
