Imports System.Windows.Forms
Imports System.Diagnostics

Public Class MailDropTaskPaneControl
    Inherits UserControl

    Public Property Session As Session

    Public Sub New()
        Debug.WriteLine("[TaskPane] Konstruktor aufgerufen")
        InitializeComponent()

        ' Session initialisieren
        Session = New Session()
        Session.Reset()
        Session.PrepareSession()

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
        Debug.WriteLine("[TaskPane] TreeView initialisiert")

        ' DataBindings für ViewModel
        ListBox1.DataBindings.Add("SelectedItem", Session, "SelectedProjekt", False, DataSourceUpdateMode.OnPropertyChanged)
        TextBoxTitel.DataBindings.Add("Text", Session, "Titel", False, DataSourceUpdateMode.OnPropertyChanged)
        TextBoxAblageordner.DataBindings.Add("Text", Session, "Ablageordner", False, DataSourceUpdateMode.OnPropertyChanged)
        TextBoxMsgDateiname.DataBindings.Add("Text", Session, "MsgDateiname", False, DataSourceUpdateMode.OnPropertyChanged)
        CheckBoxAnhaenge.DataBindings.Add("Checked", Session, "AnhaengeAblegen", False, DataSourceUpdateMode.OnPropertyChanged)
        Debug.WriteLine("[TaskPane] DataBindings gesetzt")
        ' ListBoxMetadaten: SelectedItems Binding ist nicht direkt möglich, daher Event nutzen
    End Sub

    ' Beispiel: Event-Handler für Auswahl
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Debug.WriteLine("[TaskPane] ListBox1_SelectedIndexChanged aufgerufen")
        If ListBox1.SelectedIndex <> -1 Then
            Debug.WriteLine("[TaskPane] ListBox1 Auswahl: " & ListBox1.SelectedItem.ToString())
            MessageBox.Show("Ausgewählt: " & ListBox1.SelectedItem.ToString())
        End If
    End Sub

    Private Sub ListBoxMetadaten_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxMetadaten.SelectedIndexChanged
        Debug.WriteLine("[TaskPane] ListBoxMetadaten_SelectedIndexChanged aufgerufen")
        Session.SelectedMetadaten.Clear()
        For Each item In ListBoxMetadaten.SelectedItems
            Session.SelectedMetadaten.Add(item.ToString())
        Next
        Debug.WriteLine("[TaskPane] ListBoxMetadaten Auswahl: " & String.Join(", ", Session.SelectedMetadaten))
    End Sub

    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect
        Debug.WriteLine("[TaskPane] TreeView1_AfterSelect aufgerufen: " & e.Node.Text)
        Session.SelectedOrdner = e.Node.Text
    End Sub

End Class
