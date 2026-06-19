Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core

<ComVisible(True)>
Public Class MailDropRibbon
    Implements IRibbonExtensibility

    Private ribbon As IRibbonUI

    Public Function GetCustomUI(ribbonID As String) As String Implements IRibbonExtensibility.GetCustomUI
        Dim asm = System.Reflection.Assembly.GetExecutingAssembly()
        Using stream = asm.GetManifestResourceStream("MailDrop.MailDropRibbon.xml")
            If stream Is Nothing Then
                Throw New InvalidOperationException("Embedded ribbon resource 'MailDrop.MailDropRibbon.xml' not found in assembly.")
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
