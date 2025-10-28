<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MailDropTaskPaneControl
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.LabelProjektordner = New System.Windows.Forms.Label()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.LabelAblageordner = New System.Windows.Forms.Label()
        Me.TreeView1 = New System.Windows.Forms.TreeView()
        Me.LabelTitel = New System.Windows.Forms.Label()
        Me.TextBoxTitel = New System.Windows.Forms.TextBox()
        Me.LabelAblageordnerField = New System.Windows.Forms.Label()
        Me.TextBoxAblageordner = New System.Windows.Forms.TextBox()
        Me.LabelMsgDateiname = New System.Windows.Forms.Label()
        Me.TextBoxMsgDateiname = New System.Windows.Forms.TextBox()
        Me.CheckBoxAnhaenge = New System.Windows.Forms.CheckBox()
        Me.LabelMetadaten = New System.Windows.Forms.Label()
        Me.ListBoxMetadaten = New System.Windows.Forms.ListBox()
        Me.LabelAnhaenge = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        ' 
        ' LabelProjektordner
        ' 
        Me.LabelProjektordner.AutoSize = False
        Me.LabelProjektordner.Location = New System.Drawing.Point(5, 5)
        Me.LabelProjektordner.Name = "LabelProjektordner"
        Me.LabelProjektordner.Size = New System.Drawing.Size(240, 16)
        Me.LabelProjektordner.TabIndex = 0
        Me.LabelProjektordner.Text = "Basisverzeichnis / Projektverzeichnis"
        Me.LabelProjektordner.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        ' 
        ' ListBox1
        ' 
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Items.AddRange(New Object() {"C:/Projekte/P-230511", "C:/Projekte/P-230111", "C:/Projekte/P-23251", "C:/Projekte/P-23052", "anderes..."})
        Me.ListBox1.Location = New System.Drawing.Point(5, 25)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(240, 69)
        Me.ListBox1.TabIndex = 1
        Me.ListBox1.SelectionMode = System.Windows.Forms.SelectionMode.One
        ' 
        ' LabelAblageordner
        ' 
        Me.LabelAblageordner.AutoSize = False
        Me.LabelAblageordner.Location = New System.Drawing.Point(5, 100)
        Me.LabelAblageordner.Name = "LabelAblageordner"
        Me.LabelAblageordner.Size = New System.Drawing.Size(240, 16)
        Me.LabelAblageordner.TabIndex = 2
        Me.LabelAblageordner.Text = "Projektstruktur"
        Me.LabelAblageordner.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        ' 
        ' TreeView1
        ' 
        Me.TreeView1.Location = New System.Drawing.Point(5, 120)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.Size = New System.Drawing.Size(240, 120)
        Me.TreeView1.TabIndex = 3
        ' 
        ' LabelTitel
        ' 
        Me.LabelTitel.AutoSize = True
        Me.LabelTitel.Location = New System.Drawing.Point(5, 250)
        Me.LabelTitel.Name = "LabelTitel"
        Me.LabelTitel.Size = New System.Drawing.Size(27, 13)
        Me.LabelTitel.TabIndex = 4
        Me.LabelTitel.Text = "Titel"
        ' 
        ' TextBoxTitel
        ' 
        Me.TextBoxTitel.Location = New System.Drawing.Point(90, 247)
        Me.TextBoxTitel.Name = "TextBoxTitel"
        Me.TextBoxTitel.Size = New System.Drawing.Size(155, 20)
        Me.TextBoxTitel.TabIndex = 5
        ' 
        ' LabelAblageordnerField
        ' 
        Me.LabelAblageordnerField.AutoSize = False
        Me.LabelAblageordnerField.Location = New System.Drawing.Point(5, 275)
        Me.LabelAblageordnerField.Name = "LabelAblageordnerField"
        Me.LabelAblageordnerField.Size = New System.Drawing.Size(240, 16)
        Me.LabelAblageordnerField.TabIndex = 6
        Me.LabelAblageordnerField.Text = "Ablageordner"
        Me.LabelAblageordnerField.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        ' 
        ' TextBoxAblageordner
        ' 
        Me.TextBoxAblageordner.Location = New System.Drawing.Point(90, 272)
        Me.TextBoxAblageordner.Name = "TextBoxAblageordner"
        Me.TextBoxAblageordner.Size = New System.Drawing.Size(155, 20)
        Me.TextBoxAblageordner.TabIndex = 7
        ' 
        ' LabelMsgDateiname
        ' 
        Me.LabelMsgDateiname.AutoSize = True
        Me.LabelMsgDateiname.Location = New System.Drawing.Point(5, 300)
        Me.LabelMsgDateiname.Name = "LabelMsgDateiname"
        Me.LabelMsgDateiname.Size = New System.Drawing.Size(79, 13)
        Me.LabelMsgDateiname.TabIndex = 8
        Me.LabelMsgDateiname.Text = "msg Dateiname"
        ' 
        ' TextBoxMsgDateiname
        ' 
        Me.TextBoxMsgDateiname.Location = New System.Drawing.Point(90, 297)
        Me.TextBoxMsgDateiname.Name = "TextBoxMsgDateiname"
        Me.TextBoxMsgDateiname.Size = New System.Drawing.Size(155, 20)
        Me.TextBoxMsgDateiname.TabIndex = 9
        ' 
        ' LabelAnhaenge
        ' 
        Me.LabelAnhaenge.AutoSize = False
        Me.LabelAnhaenge.Location = New System.Drawing.Point(5, 325)
        Me.LabelAnhaenge.Name = "LabelAnhaenge"
        Me.LabelAnhaenge.Size = New System.Drawing.Size(110, 16)
        Me.LabelAnhaenge.TabIndex = 13
        Me.LabelAnhaenge.Text = "Anhänge ablegen"
        Me.LabelAnhaenge.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        ' 
        ' CheckBoxAnhaenge
        ' 
        Me.CheckBoxAnhaenge.AutoSize = True
        Me.CheckBoxAnhaenge.Location = New System.Drawing.Point(100, 323)
        Me.CheckBoxAnhaenge.Name = "CheckBoxAnhaenge"
        Me.CheckBoxAnhaenge.Size = New System.Drawing.Size(15, 14)
        Me.CheckBoxAnhaenge.TabIndex = 10
        Me.CheckBoxAnhaenge.Text = ""
        Me.CheckBoxAnhaenge.UseVisualStyleBackColor = True
        ' 
        ' LabelMetadaten
        ' 
        Me.LabelMetadaten.AutoSize = True
        Me.LabelMetadaten.Location = New System.Drawing.Point(5, 350)
        Me.LabelMetadaten.Name = "LabelMetadaten"
        Me.LabelMetadaten.Size = New System.Drawing.Size(87, 13)
        Me.LabelMetadaten.TabIndex = 11
        Me.LabelMetadaten.Text = "Email Metadaten"
        ' 
        ' ListBoxMetadaten
        ' 
        Me.ListBoxMetadaten.FormattingEnabled = True
        Me.ListBoxMetadaten.Items.AddRange(New Object() {"Von", "An", "Betreff", "Datum"})
        Me.ListBoxMetadaten.Location = New System.Drawing.Point(5, 370)
        Me.ListBoxMetadaten.Name = "ListBoxMetadaten"
        Me.ListBoxMetadaten.Size = New System.Drawing.Size(240, 56)
        Me.ListBoxMetadaten.TabIndex = 12
        ' 
        ' MailDropTaskPaneControl
        ' 
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.ListBoxMetadaten)
        Me.Controls.Add(Me.LabelMetadaten)
        Me.Controls.Add(Me.CheckBoxAnhaenge)
        Me.Controls.Add(Me.TextBoxMsgDateiname)
        Me.Controls.Add(Me.LabelMsgDateiname)
        Me.Controls.Add(Me.TextBoxAblageordner)
        Me.Controls.Add(Me.LabelAblageordnerField)
        Me.Controls.Add(Me.TextBoxTitel)
        Me.Controls.Add(Me.LabelTitel)
        Me.Controls.Add(Me.TreeView1)
        Me.Controls.Add(Me.LabelAblageordner)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.LabelProjektordner)
        Me.Controls.Add(Me.LabelAnhaenge)
        Me.Name = "MailDropTaskPaneControl"
        Me.Size = New System.Drawing.Size(250, 400)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents LabelProjektordner As System.Windows.Forms.Label
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents LabelAblageordner As System.Windows.Forms.Label
    Friend WithEvents TreeView1 As System.Windows.Forms.TreeView
    Friend WithEvents LabelTitel As System.Windows.Forms.Label
    Friend WithEvents TextBoxTitel As System.Windows.Forms.TextBox
    Friend WithEvents LabelAblageordnerField As System.Windows.Forms.Label
    Friend WithEvents TextBoxAblageordner As System.Windows.Forms.TextBox
    Friend WithEvents LabelMsgDateiname As System.Windows.Forms.Label
    Friend WithEvents TextBoxMsgDateiname As System.Windows.Forms.TextBox
    Friend WithEvents CheckBoxAnhaenge As System.Windows.Forms.CheckBox
    Friend WithEvents LabelMetadaten As System.Windows.Forms.Label
    Friend WithEvents ListBoxMetadaten As System.Windows.Forms.ListBox
    Friend WithEvents LabelAnhaenge As System.Windows.Forms.Label

End Class
