Imports System.Windows.Controls
Imports System.IO
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.Animation
Imports System.Windows.Input
Imports System.Diagnostics
Imports Microsoft.VisualBasic

Public Class MailDropWpfTaskPane
    Inherits UserControl

    Public Property Session As New Session()
    Private infoPopup As InfoPopup = Nothing
    Private _applyingTreeViewSuggestion As Boolean = False

    Public Sub New()
        InitializeComponent()
        Me.DataContext = Session
        AddHandler ListBox1.SelectionChanged, AddressOf ListBox1_SelectionChanged
        AddHandler Session.PropertyChanged, AddressOf Session_PropertyChanged
        AddHandler TreeView1.SelectedItemChanged, AddressOf TreeView1_SelectedItemChanged
    End Sub

    Private Function FindTreeViewItem(container As ItemsControl, relativePath As String) As TreeViewItem
        For Each item In container.Items
            Dim node = TryCast(item, DirectoryNode)
            Dim tvi = TryCast(container.ItemContainerGenerator.ContainerFromItem(item), TreeViewItem)
            If node IsNot Nothing AndAlso tvi IsNot Nothing Then
                If node.RelativePath = relativePath Then Return tvi
                tvi.IsExpanded = True
                tvi.UpdateLayout()
                Dim found = FindTreeViewItem(tvi, relativePath)
                If found IsNot Nothing Then Return found
            End If
        Next
        Return Nothing
    End Function

    Private Function GetRelativePath(basePath As String, fullPath As String) As String
        If String.IsNullOrWhiteSpace(basePath) OrElse String.IsNullOrWhiteSpace(fullPath) Then
            Return String.Empty
        End If

        Dim normalizedBase = basePath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        If fullPath.StartsWith(normalizedBase, StringComparison.OrdinalIgnoreCase) Then
            Return fullPath.Substring(normalizedBase.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        End If

        Return fullPath
    End Function

    Private Sub SelectTreeViewPath(relativePath As String)
        If String.IsNullOrWhiteSpace(relativePath) Then
            Return
        End If

        If Not Dispatcher.HasShutdownStarted Then
            Dispatcher.BeginInvoke(New Action(Sub()
                TreeView1.UpdateLayout()
                Dim tvi = FindTreeViewItem(TreeView1, relativePath)
                If tvi IsNot Nothing Then
                    tvi.IsSelected = True
                    tvi.BringIntoView()
                End If
            End Sub))
        End If
    End Sub

    Private Sub ShowSparkle(sparkle As TextBlock)
        sparkle.BeginAnimation(UIElement.OpacityProperty,
            New DoubleAnimation(0, 1, New Duration(TimeSpan.FromMilliseconds(300))))
    End Sub

    Private Sub HideSparkle(sparkle As TextBlock)
        sparkle.BeginAnimation(UIElement.OpacityProperty,
            New DoubleAnimation(sparkle.Opacity, 0, New Duration(TimeSpan.FromMilliseconds(200))))
    End Sub

    Private Sub ListBox1_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        Session.IsSuggestedProjektPfad = False
        If ListBox1.SelectedItem IsNot Nothing Then
            Session.HandleProjektSelection(ListBox1.SelectedItem.ToString())
        End If
    End Sub

    Private Sub Session_PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs)
        Select Case e.PropertyName
            Case NameOf(Session.IsSuggestedProjektPfad)
                If Session.IsSuggestedProjektPfad Then ShowSparkle(SparkleProjektPfad) Else HideSparkle(SparkleProjektPfad)
            Case NameOf(Session.IsSuggestedProjektstrukturPfad)
                If Session.IsSuggestedProjektstrukturPfad Then
                    ShowSparkle(SparkleProjektstrukturPfad)
                    Dim targetPath = Session.ProjektstrukturPfad
                    If Not Dispatcher.HasShutdownStarted Then
                        Dispatcher.BeginInvoke(New Action(Sub()
                            _applyingTreeViewSuggestion = True
                            Dim tvi = FindTreeViewItem(TreeView1, targetPath)
                            If tvi IsNot Nothing Then
                                tvi.IsSelected = True
                                tvi.IsExpanded = True
                                tvi.BringIntoView()
                            End If
                            _applyingTreeViewSuggestion = False
                        End Sub))
                    End If
                Else
                    HideSparkle(SparkleProjektstrukturPfad)
                End If
            Case NameOf(Session.IsSuggestedTitel)
                If Session.IsSuggestedTitel Then ShowSparkle(SparkleTitel) Else HideSparkle(SparkleTitel)
            Case NameOf(Session.IsSuggestedAbsenderKurz)
                If Session.IsSuggestedAbsenderKurz Then ShowSparkle(SparkleAbsenderKurz) Else HideSparkle(SparkleAbsenderKurz)
            Case NameOf(Session.IsSuggestedAblageordnerSchema)
                If Session.IsSuggestedAblageordnerSchema Then ShowSparkle(SparkleAblageordner) Else HideSparkle(SparkleAblageordner)
            Case NameOf(Session.IsSuggestedMsgDateinameSchema)
                If Session.IsSuggestedMsgDateinameSchema Then ShowSparkle(SparkleMsgDateiname) Else HideSparkle(SparkleMsgDateiname)
            Case NameOf(Session.IsSuggestedAnhaengeAblegen)
                If Session.IsSuggestedAnhaengeAblegen Then ShowSparkle(SparkleAnhaengeAblegen) Else HideSparkle(SparkleAnhaengeAblegen)
        End Select
    End Sub

    Private Sub TextBoxTitel_GotFocus(sender As Object, e As RoutedEventArgs)
        Session.IsSuggestedTitel = False
    End Sub

    Private Sub TextBoxAbsenderKurz_GotFocus(sender As Object, e As RoutedEventArgs)
        Session.IsSuggestedAbsenderKurz = False
    End Sub

    Private Sub TextBoxAblageordner_GotFocus(sender As Object, e As RoutedEventArgs)
        Session.IsSuggestedAblageordnerSchema = False
        Session.BeginAblageordnerEdit()
    End Sub

    Private Sub TextBoxAblageordner_LostFocus(sender As Object, e As RoutedEventArgs)
        Session.EndAblageordnerEdit()
    End Sub

    Private Sub TextBoxMsgDateiname_GotFocus(sender As Object, e As RoutedEventArgs)
        Session.IsSuggestedMsgDateinameSchema = False
        Session.BeginMsgDateinameEdit()
    End Sub

    Private Sub TextBoxMsgDateiname_LostFocus(sender As Object, e As RoutedEventArgs)
        Session.EndMsgDateinameEdit()
    End Sub

    Private Sub CheckBoxAnhaenge_Click(sender As Object, e As RoutedEventArgs)
        Session.IsSuggestedAnhaengeAblegen = False
    End Sub

    Private Sub ButtonInfo_Click(sender As Object, e As RoutedEventArgs)
        If infoPopup Is Nothing OrElse Not infoPopup.IsLoaded Then
            infoPopup = New InfoPopup()
            infoPopup.Owner = Window.GetWindow(Me)
            infoPopup.WindowStartupLocation = WindowStartupLocation.CenterOwner
            infoPopup.Show()
        Else
            infoPopup.Activate()
        End If
    End Sub

    Private Sub ButtonOk_Click(sender As Object, e As RoutedEventArgs)
        Dim result As String = Session.ProcessSession()
        If Not String.IsNullOrEmpty(result) Then
            MessageBox.Show(result, "Fehler", MessageBoxButton.OK, MessageBoxImage.Error)
        Else
            ShowSuccessNotification()
        End If
    End Sub

    Private Sub ShowSuccessNotification()
        SuccessNotification.BeginAnimation(UIElement.OpacityProperty,
            New DoubleAnimation(0, 1, New Duration(TimeSpan.FromMilliseconds(250))))
        Dim timer As New System.Windows.Threading.DispatcherTimer()
        timer.Interval = TimeSpan.FromSeconds(2.5)
        AddHandler timer.Tick, Sub(s, ev)
            timer.Stop()
            SuccessNotification.BeginAnimation(UIElement.OpacityProperty,
                New DoubleAnimation(1, 0, New Duration(TimeSpan.FromMilliseconds(600))))
        End Sub
        timer.Start()
    End Sub

    Private Sub ButtonAbbrechen_Click(sender As Object, e As RoutedEventArgs)
        ' Schlie�e InfoPopup, falls offen
        If infoPopup IsNot Nothing AndAlso infoPopup.IsLoaded Then
            infoPopup.Close()
            infoPopup = Nothing
        End If
        Try
            Globals.ThisAddIn.HideTaskPane()
        Catch
            ' Fallback: Fenster schlie�en
            Dim wnd = Window.GetWindow(Me)
            If wnd IsNot Nothing Then wnd.Close()
        End Try
    End Sub

    ' Setzt die Editierbarkeit der TaskPane
    Public Sub SetEditMode(isEditable As Boolean)
        ' Beispiel: Alle Controls au�er Info-Button deaktivieren/aktivieren
        For Each ctrl In Me.FindVisualChildren(Of Control)(Me)
            If ctrl.Name <> "ButtonInfo" AndAlso ctrl.Name <> "ButtonGewichteNeuBerechnen" Then
                ctrl.IsEnabled = isEditable
            End If
        Next
    End Sub

    ' Gibt True zur�ck, wenn genau eine Mail selektiert ist, sonst False
    Public Function SingleMailSelected() As Boolean
        Dim explorer = Globals.ThisAddIn.Application.ActiveExplorer()
        If explorer Is Nothing Then Return False
        Dim selection = explorer.Selection
        Return selection.Count = 1 AndAlso TypeOf selection.Item(1) Is Outlook.MailItem
    End Function

    ' Hilfsmethode: Findet alle Controls eines Typs rekursiv
    Private Iterator Function FindVisualChildren(Of T As DependencyObject)(depObj As DependencyObject) As IEnumerable(Of T)
        If depObj IsNot Nothing Then
            For i As Integer = 0 To VisualTreeHelper.GetChildrenCount(depObj) - 1
                Dim child = VisualTreeHelper.GetChild(depObj, i)
                If TypeOf child Is T Then
                    Yield CType(child, T)
                End If
                For Each childOfChild In FindVisualChildren(Of T)(child)
                    Yield childOfChild
                Next
            Next
        End If
    End Function

    Private Sub TreeView1_SelectedItemChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object))
        If Not _applyingTreeViewSuggestion Then
            Session.IsSuggestedProjektstrukturPfad = False
        End If
        Dim node = TryCast(TreeView1.SelectedItem, DirectoryNode)
        If node IsNot Nothing Then
            Session.ProjektstrukturPfad = node.RelativePath
            Debug.WriteLine("ProjektstrukturPfad (relativ) gesetzt: " & node.RelativePath)
            Dim tvi = FindSelectedTreeViewItem(TreeView1)
            If tvi IsNot Nothing Then tvi.IsExpanded = True
        End If
    End Sub

    ' Findet das aktuell selektierte TreeViewItem, indem nur bereits aufgeklappte Knoten
    ' durchsucht werden (der Benutzer muss die Eltern geöffnet haben, um den Knoten zu wählen).
    Private Function FindSelectedTreeViewItem(container As ItemsControl) As TreeViewItem
        For Each item In container.Items
            Dim tvi = TryCast(container.ItemContainerGenerator.ContainerFromItem(item), TreeViewItem)
            If tvi Is Nothing Then Continue For
            If tvi.IsSelected Then Return tvi
            If tvi.IsExpanded Then
                Dim found = FindSelectedTreeViewItem(tvi)
                If found IsNot Nothing Then Return found
            End If
        Next
        Return Nothing
    End Function

    Private Function EnsureValidProjektPfad() As Boolean
        If String.IsNullOrWhiteSpace(Session.ProjektPfad) OrElse Not Directory.Exists(Session.ProjektPfad) Then
            MessageBox.Show("Bitte zuerst ein gueltiges Projektverzeichnis auswaehlen.", "Projektstruktur", MessageBoxButton.OK, MessageBoxImage.Information)
            Return False
        End If
        Return True
    End Function

    Private Function IsValidFolderName(folderName As String) As Boolean
        If String.IsNullOrWhiteSpace(folderName) Then
            Return False
        End If
        Return folderName.IndexOfAny(Path.GetInvalidFileNameChars()) < 0
    End Function

    Private Sub CreateFolderUnderSelection()
        If Not EnsureValidProjektPfad() Then
            Return
        End If

        Dim selectedNode = TryCast(TreeView1.SelectedItem, DirectoryNode)
        Dim parentPath = If(selectedNode IsNot Nothing, selectedNode.FullPath, Session.ProjektPfad)
        Dim folderName = Interaction.InputBox("Name fuer den neuen Ordner:", "Neuen Ordner erstellen", "Neuer Ordner").Trim()

        If String.IsNullOrWhiteSpace(folderName) Then
            Return
        End If

        If Not IsValidFolderName(folderName) Then
            MessageBox.Show("Der Ordnername enthaelt ungueltige Zeichen.", "Neuen Ordner erstellen", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim newFolderPath = Path.Combine(parentPath, folderName)

        Try
            If Directory.Exists(newFolderPath) Then
                MessageBox.Show("Der Ordner existiert bereits.", "Neuen Ordner erstellen", MessageBoxButton.OK, MessageBoxImage.Information)
            Else
                Directory.CreateDirectory(newFolderPath)
            End If
        Catch ex As Exception
            MessageBox.Show("Ordner konnte nicht erstellt werden: " & ex.Message, "Neuen Ordner erstellen", MessageBoxButton.OK, MessageBoxImage.Error)
            Return
        End Try

        Session.BuildDirectoryTree()
        Dim relativePath = GetRelativePath(Session.ProjektPfad, newFolderPath)
        Session.ProjektstrukturPfad = relativePath
        Session.IsSuggestedProjektstrukturPfad = False
        HideSparkle(SparkleProjektstrukturPfad)
        SelectTreeViewPath(relativePath)
    End Sub

    Private Sub DeleteSelectedFolder()
        If Not EnsureValidProjektPfad() Then
            Return
        End If

        Dim selectedNode = TryCast(TreeView1.SelectedItem, DirectoryNode)
        If selectedNode Is Nothing Then
            MessageBox.Show("Bitte zuerst einen Ordner in der Projektstruktur auswaehlen.", "Ordner loeschen", MessageBoxButton.OK, MessageBoxImage.Information)
            Return
        End If

        If Directory.GetFileSystemEntries(selectedNode.FullPath).Length > 0 Then
            MessageBox.Show("Der Ordner ist nicht leer und kann daher nicht gelöscht werden.", "Ordner löschen", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim question = "Soll der Ordner wirklich gelöscht werden?" & Environment.NewLine & selectedNode.FullPath
        Dim result = MessageBox.Show(question, "Ordner löschen", MessageBoxButton.YesNo, MessageBoxImage.Warning)
        If result <> MessageBoxResult.Yes Then
            Return
        End If

        Dim parentPath = Path.GetDirectoryName(selectedNode.FullPath)

        Try
            Directory.Delete(selectedNode.FullPath, recursive:=False)
        Catch ex As Exception
            MessageBox.Show("Ordner konnte nicht gelöscht werden: " & ex.Message, "Ordner löschen", MessageBoxButton.OK, MessageBoxImage.Error)
            Return
        End Try

        Session.BuildDirectoryTree()
        Dim parentRelativePath = GetRelativePath(Session.ProjektPfad, parentPath)
        Session.ProjektstrukturPfad = parentRelativePath
        Session.IsSuggestedProjektstrukturPfad = False
        HideSparkle(SparkleProjektstrukturPfad)
        SelectTreeViewPath(parentRelativePath)
    End Sub

    Private Sub RenameSelectedFolder()
        If Not EnsureValidProjektPfad() Then
            Return
        End If

        Dim selectedNode = TryCast(TreeView1.SelectedItem, DirectoryNode)
        If selectedNode Is Nothing Then
            MessageBox.Show("Bitte zuerst einen Ordner in der Projektstruktur auswaehlen.", "Ordner umbenennen", MessageBoxButton.OK, MessageBoxImage.Information)
            Return
        End If

        Dim newName = Interaction.InputBox("Neuer Name fuer den Ordner:", "Ordner umbenennen", selectedNode.Name).Trim()
        If String.IsNullOrWhiteSpace(newName) Then
            Return
        End If

        If Not IsValidFolderName(newName) Then
            MessageBox.Show("Der Ordnername enthaelt ungueltige Zeichen.", "Ordner umbenennen", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim parentPath = Path.GetDirectoryName(selectedNode.FullPath)
        If String.IsNullOrWhiteSpace(parentPath) Then
            MessageBox.Show("Der ausgewaehlte Ordner kann nicht umbenannt werden.", "Ordner umbenennen", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim renamedPath = Path.Combine(parentPath, newName)
        If String.Equals(selectedNode.FullPath, renamedPath, StringComparison.OrdinalIgnoreCase) Then
            Return
        End If

        If Directory.Exists(renamedPath) Then
            MessageBox.Show("Ein Ordner mit diesem Namen existiert bereits.", "Ordner umbenennen", MessageBoxButton.OK, MessageBoxImage.Information)
            Return
        End If

        Try
            Directory.Move(selectedNode.FullPath, renamedPath)
        Catch ex As Exception
            MessageBox.Show("Ordner konnte nicht umbenannt werden: " & ex.Message, "Ordner umbenennen", MessageBoxButton.OK, MessageBoxImage.Error)
            Return
        End Try

        Session.BuildDirectoryTree()
        Dim relativePath = GetRelativePath(Session.ProjektPfad, renamedPath)
        Session.ProjektstrukturPfad = relativePath
        Session.IsSuggestedProjektstrukturPfad = False
        HideSparkle(SparkleProjektstrukturPfad)
        SelectTreeViewPath(relativePath)
    End Sub

    Private Sub ButtonNeuenOrdner_Click(sender As Object, e As RoutedEventArgs)
        CreateFolderUnderSelection()
    End Sub

    Private Sub MenuItemNeuerOrdner_Click(sender As Object, e As RoutedEventArgs)
        CreateFolderUnderSelection()
    End Sub

    Private Sub MenuItemLoeschen_Click(sender As Object, e As RoutedEventArgs)
        DeleteSelectedFolder()
    End Sub

    Private Sub MenuItemUmbenennen_Click(sender As Object, e As RoutedEventArgs)
        RenameSelectedFolder()
    End Sub

    Private Sub ButtonGewichteNeuBerechnen_Click(sender As Object, e As RoutedEventArgs)
        Task.Run(Sub()
                     Try
                         SuggestionEngine.GetSharedInstance().RecalculateWeightsFromHistory()
                     Catch ex As Exception
                         Debug.WriteLine("[TaskPane] Gewichte-Neuberechnung fehlgeschlagen: " & ex.Message)
                     End Try
                 End Sub)
    End Sub

    Private Sub TreeView1_PreviewMouseRightButtonDown(sender As Object, e As MouseButtonEventArgs)
        Dim depObj = TryCast(e.OriginalSource, DependencyObject)
        While depObj IsNot Nothing AndAlso Not TypeOf depObj Is TreeViewItem
            depObj = VisualTreeHelper.GetParent(depObj)
        End While

        Dim treeViewItem = TryCast(depObj, TreeViewItem)
        If treeViewItem IsNot Nothing Then
            treeViewItem.IsSelected = True
            treeViewItem.Focus()
        End If
    End Sub
End Class
