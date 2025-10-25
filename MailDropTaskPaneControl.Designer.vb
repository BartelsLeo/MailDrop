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
        Me.SuspendLayout()
        ' 
        ' LabelProjektordner
        ' 
        Me.LabelProjektordner.AutoSize = True
        Me.LabelProjektordner.Location = New System.Drawing.Point(5, 5)
        Me.LabelProjektordner.Name = "LabelProjektordner"
        Me.LabelProjektordner.Size = New System.Drawing.Size(143, 13)
        Me.LabelProjektordner.TabIndex = 0
        Me.LabelProjektordner.Text = "Basisverzeichnis / Projektverzeichnis"
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
        Me.LabelAblageordner.AutoSize = True
        Me.LabelAblageordner.Location = New System.Drawing.Point(5, 100)
        Me.LabelAblageordner.Name = "LabelAblageordner"
        Me.LabelAblageordner.Size = New System.Drawing.Size(74, 13)
        Me.LabelAblageordner.TabIndex = 2
        Me.LabelAblageordner.Text = "Ablageordner"
        ' 
        ' TreeView1
        ' 
        Me.TreeView1.Location = New System.Drawing.Point(5, 120)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.Size = New System.Drawing.Size(240, 120)
        Me.TreeView1.TabIndex = 3
        ' 
        ' MailDropTaskPaneControl
        ' 
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.TreeView1)
        Me.Controls.Add(Me.LabelAblageordner)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.LabelProjektordner)
        Me.Name = "MailDropTaskPaneControl"
        Me.Size = New System.Drawing.Size(250, 400)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents LabelProjektordner As System.Windows.Forms.Label
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents LabelAblageordner As System.Windows.Forms.Label
    Friend WithEvents TreeView1 As System.Windows.Forms.TreeView

End Class
