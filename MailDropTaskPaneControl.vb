Imports System.Windows.Forms

Public Class MailDropTaskPaneControl
    Inherits UserControl

    Public Sub New()
        InitializeComponent()
        ' Höhe so setzen, dass genau 5 Zeilen sichtbar sind (inkl. Rahmen)
        ListBox1.Height = ListBox1.ItemHeight * 5 + 4
        ' Dummy-Nodes für TreeView
        TreeView1.Nodes.Clear()
        Dim root As New TreeNode("Root")
        root.Nodes.Add("Ordner 1")
        root.Nodes.Add("Ordner 2")
        root.Nodes.Add("Ordner 3")
        TreeView1.Nodes.Add(root)
        TreeView1.ExpandAll()
    End Sub

    ' Beispiel: Event-Handler für Auswahl
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        If ListBox1.SelectedIndex <> -1 Then
            MessageBox.Show("Ausgewählt: " & ListBox1.SelectedItem.ToString())
        End If
    End Sub

End Class
