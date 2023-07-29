Imports System.Windows.Forms

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class InteractiveReplNode
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(InteractiveReplNode))
        Me.TextBoxNodeCodeInput = New System.Windows.Forms.TextBox()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.SplitContainer5 = New System.Windows.Forms.SplitContainer()
        Me.ButtonExecute = New System.Windows.Forms.Button()
        Me.SplitContainer6 = New System.Windows.Forms.SplitContainer()
        Me.ButtonAddCell = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer()
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer()
        Me.SplitContainer7 = New System.Windows.Forms.SplitContainer()
        Me.RadioButtonIsMainCell = New System.Windows.Forms.RadioButton()
        Me.Code_Lable = New System.Windows.Forms.Label()
        Me.ToolStripStandard = New System.Windows.Forms.ToolStrip()
        Me.NewToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.OpenToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.SaveToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.PrintToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.toolStripSeparator = New System.Windows.Forms.ToolStripSeparator()
        Me.CutToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.CopyToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.PasteToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.toolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.HelpToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.RichTextBoxNodeOutput = New System.Windows.Forms.RichTextBox()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.SplitContainer5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer5.Panel1.SuspendLayout()
        Me.SplitContainer5.Panel2.SuspendLayout()
        Me.SplitContainer5.SuspendLayout()
        CType(Me.SplitContainer6, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer6.Panel1.SuspendLayout()
        Me.SplitContainer6.Panel2.SuspendLayout()
        Me.SplitContainer6.SuspendLayout()
        CType(Me.SplitContainer3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        CType(Me.SplitContainer4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        CType(Me.SplitContainer7, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer7.Panel1.SuspendLayout()
        Me.SplitContainer7.SuspendLayout()
        Me.ToolStripStandard.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBoxNodeCodeInput
        '
        Me.TextBoxNodeCodeInput.AcceptsReturn = True
        Me.TextBoxNodeCodeInput.AcceptsTab = True
        Me.TextBoxNodeCodeInput.AllowDrop = True
        Me.TextBoxNodeCodeInput.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TextBoxNodeCodeInput.Font = New System.Drawing.Font("Consolas", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxNodeCodeInput.Location = New System.Drawing.Point(0, 0)
        Me.TextBoxNodeCodeInput.Margin = New System.Windows.Forms.Padding(4)
        Me.TextBoxNodeCodeInput.Multiline = True
        Me.TextBoxNodeCodeInput.Name = "TextBoxNodeCodeInput"
        Me.TextBoxNodeCodeInput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBoxNodeCodeInput.Size = New System.Drawing.Size(1376, 225)
        Me.TextBoxNodeCodeInput.TabIndex = 1
        Me.TextBoxNodeCodeInput.Text = "" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(4)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.SplitContainer2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer3)
        Me.SplitContainer1.Size = New System.Drawing.Size(1611, 382)
        Me.SplitContainer1.SplitterDistance = 225
        Me.SplitContainer1.SplitterWidth = 6
        Me.SplitContainer1.TabIndex = 1
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(4)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.TextBoxNodeCodeInput)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.Panel1)
        Me.SplitContainer2.Size = New System.Drawing.Size(1611, 225)
        Me.SplitContainer2.SplitterDistance = 1376
        Me.SplitContainer2.SplitterWidth = 5
        Me.SplitContainer2.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.SplitContainer5)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(230, 57)
        Me.Panel1.TabIndex = 0
        '
        'SplitContainer5
        '
        Me.SplitContainer5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer5.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer5.Name = "SplitContainer5"
        '
        'SplitContainer5.Panel1
        '
        Me.SplitContainer5.Panel1.Controls.Add(Me.ButtonExecute)
        '
        'SplitContainer5.Panel2
        '
        Me.SplitContainer5.Panel2.Controls.Add(Me.SplitContainer6)
        Me.SplitContainer5.Size = New System.Drawing.Size(230, 57)
        Me.SplitContainer5.SplitterDistance = 76
        Me.SplitContainer5.TabIndex = 0
        '
        'ButtonExecute
        '
        Me.ButtonExecute.BackColor = System.Drawing.Color.Lime
        Me.ButtonExecute.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ButtonExecute.Font = New System.Drawing.Font("Continuum Bold", 8.999999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonExecute.ForeColor = System.Drawing.Color.Maroon
        Me.ButtonExecute.Location = New System.Drawing.Point(0, 0)
        Me.ButtonExecute.Name = "ButtonExecute"
        Me.ButtonExecute.Size = New System.Drawing.Size(76, 57)
        Me.ButtonExecute.TabIndex = 4
        Me.ButtonExecute.Text = "Execute"
        Me.ButtonExecute.UseVisualStyleBackColor = False
        '
        'SplitContainer6
        '
        Me.SplitContainer6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer6.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer6.Name = "SplitContainer6"
        '
        'SplitContainer6.Panel1
        '
        Me.SplitContainer6.Panel1.Controls.Add(Me.ButtonAddCell)
        '
        'SplitContainer6.Panel2
        '
        Me.SplitContainer6.Panel2.Controls.Add(Me.Button1)
        Me.SplitContainer6.Size = New System.Drawing.Size(150, 57)
        Me.SplitContainer6.SplitterDistance = 76
        Me.SplitContainer6.TabIndex = 0
        '
        'ButtonAddCell
        '
        Me.ButtonAddCell.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ButtonAddCell.BackColor = System.Drawing.Color.Lime
        Me.ButtonAddCell.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ButtonAddCell.Font = New System.Drawing.Font("Continuum Bold", 8.999999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonAddCell.ForeColor = System.Drawing.Color.Maroon
        Me.ButtonAddCell.Location = New System.Drawing.Point(0, 0)
        Me.ButtonAddCell.Name = "ButtonAddCell"
        Me.ButtonAddCell.Padding = New System.Windows.Forms.Padding(6)
        Me.ButtonAddCell.Size = New System.Drawing.Size(76, 57)
        Me.ButtonAddCell.TabIndex = 5
        Me.ButtonAddCell.Text = "Add"
        Me.ButtonAddCell.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.Button1.BackColor = System.Drawing.Color.Lime
        Me.Button1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button1.Font = New System.Drawing.Font("Continuum Bold", 8.999999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Maroon
        Me.Button1.Location = New System.Drawing.Point(0, 0)
        Me.Button1.Name = "Button1"
        Me.Button1.Padding = New System.Windows.Forms.Padding(6)
        Me.Button1.Size = New System.Drawing.Size(70, 57)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "Remove"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer3.Margin = New System.Windows.Forms.Padding(4)
        Me.SplitContainer3.Name = "SplitContainer3"
        Me.SplitContainer3.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.SplitContainer4)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.RichTextBoxNodeOutput)
        Me.SplitContainer3.Size = New System.Drawing.Size(1611, 151)
        Me.SplitContainer3.SplitterDistance = 79
        Me.SplitContainer3.SplitterWidth = 6
        Me.SplitContainer3.TabIndex = 0
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer4.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer4.Margin = New System.Windows.Forms.Padding(4)
        Me.SplitContainer4.Name = "SplitContainer4"
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.SplitContainer7)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Code_Lable)
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.BackColor = System.Drawing.Color.FloralWhite
        Me.SplitContainer4.Panel2.Controls.Add(Me.ToolStripStandard)
        Me.SplitContainer4.Size = New System.Drawing.Size(1611, 79)
        Me.SplitContainer4.SplitterDistance = 532
        Me.SplitContainer4.SplitterWidth = 5
        Me.SplitContainer4.TabIndex = 0
        '
        'SplitContainer7
        '
        Me.SplitContainer7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer7.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer7.Name = "SplitContainer7"
        '
        'SplitContainer7.Panel1
        '
        Me.SplitContainer7.Panel1.Controls.Add(Me.RadioButtonIsMainCell)
        Me.SplitContainer7.Size = New System.Drawing.Size(532, 61)
        Me.SplitContainer7.SplitterDistance = 177
        Me.SplitContainer7.TabIndex = 1
        '
        'RadioButtonIsMainCell
        '
        Me.RadioButtonIsMainCell.AutoSize = True
        Me.RadioButtonIsMainCell.Location = New System.Drawing.Point(20, 19)
        Me.RadioButtonIsMainCell.Name = "RadioButtonIsMainCell"
        Me.RadioButtonIsMainCell.Size = New System.Drawing.Size(106, 22)
        Me.RadioButtonIsMainCell.TabIndex = 0
        Me.RadioButtonIsMainCell.TabStop = True
        Me.RadioButtonIsMainCell.Text = "IsMainCell"
        Me.RadioButtonIsMainCell.UseVisualStyleBackColor = True
        '
        'Code_Lable
        '
        Me.Code_Lable.AutoSize = True
        Me.Code_Lable.BackColor = System.Drawing.Color.White
        Me.Code_Lable.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Code_Lable.Location = New System.Drawing.Point(0, 61)
        Me.Code_Lable.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Code_Lable.Name = "Code_Lable"
        Me.Code_Lable.Size = New System.Drawing.Size(56, 18)
        Me.Code_Lable.TabIndex = 0
        Me.Code_Lable.Text = "Output"
        '
        'ToolStripStandard
        '
        Me.ToolStripStandard.BackColor = System.Drawing.Color.FloralWhite
        Me.ToolStripStandard.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ToolStripStandard.Font = New System.Drawing.Font("Segoe UI Semilight", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStandard.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ToolStripStandard.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewToolStripButton, Me.OpenToolStripButton, Me.SaveToolStripButton, Me.PrintToolStripButton, Me.toolStripSeparator, Me.CutToolStripButton, Me.CopyToolStripButton, Me.PasteToolStripButton, Me.toolStripSeparator1, Me.HelpToolStripButton})
        Me.ToolStripStandard.Location = New System.Drawing.Point(0, 0)
        Me.ToolStripStandard.MinimumSize = New System.Drawing.Size(0, 62)
        Me.ToolStripStandard.Name = "ToolStripStandard"
        Me.ToolStripStandard.Size = New System.Drawing.Size(1074, 79)
        Me.ToolStripStandard.Stretch = True
        Me.ToolStripStandard.TabIndex = 2
        '
        'NewToolStripButton
        '
        Me.NewToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.NewToolStripButton.Image = CType(resources.GetObject("NewToolStripButton.Image"), System.Drawing.Image)
        Me.NewToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.NewToolStripButton.Name = "NewToolStripButton"
        Me.NewToolStripButton.Size = New System.Drawing.Size(28, 76)
        Me.NewToolStripButton.Text = "&New"
        '
        'OpenToolStripButton
        '
        Me.OpenToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.OpenToolStripButton.Image = CType(resources.GetObject("OpenToolStripButton.Image"), System.Drawing.Image)
        Me.OpenToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.OpenToolStripButton.Name = "OpenToolStripButton"
        Me.OpenToolStripButton.Size = New System.Drawing.Size(28, 76)
        Me.OpenToolStripButton.Text = "&Open"
        '
        'SaveToolStripButton
        '
        Me.SaveToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.SaveToolStripButton.Image = CType(resources.GetObject("SaveToolStripButton.Image"), System.Drawing.Image)
        Me.SaveToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.SaveToolStripButton.Name = "SaveToolStripButton"
        Me.SaveToolStripButton.Size = New System.Drawing.Size(28, 76)
        Me.SaveToolStripButton.Text = "&Save"
        '
        'PrintToolStripButton
        '
        Me.PrintToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.PrintToolStripButton.Image = CType(resources.GetObject("PrintToolStripButton.Image"), System.Drawing.Image)
        Me.PrintToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.PrintToolStripButton.Name = "PrintToolStripButton"
        Me.PrintToolStripButton.Size = New System.Drawing.Size(28, 76)
        Me.PrintToolStripButton.Text = "&Print"
        '
        'toolStripSeparator
        '
        Me.toolStripSeparator.Name = "toolStripSeparator"
        Me.toolStripSeparator.Size = New System.Drawing.Size(6, 79)
        '
        'CutToolStripButton
        '
        Me.CutToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.CutToolStripButton.Image = CType(resources.GetObject("CutToolStripButton.Image"), System.Drawing.Image)
        Me.CutToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.CutToolStripButton.Name = "CutToolStripButton"
        Me.CutToolStripButton.Size = New System.Drawing.Size(28, 76)
        Me.CutToolStripButton.Text = "C&ut"
        '
        'CopyToolStripButton
        '
        Me.CopyToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.CopyToolStripButton.Image = CType(resources.GetObject("CopyToolStripButton.Image"), System.Drawing.Image)
        Me.CopyToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.CopyToolStripButton.Name = "CopyToolStripButton"
        Me.CopyToolStripButton.Size = New System.Drawing.Size(28, 76)
        Me.CopyToolStripButton.Text = "&Copy"
        '
        'PasteToolStripButton
        '
        Me.PasteToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.PasteToolStripButton.Image = CType(resources.GetObject("PasteToolStripButton.Image"), System.Drawing.Image)
        Me.PasteToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.PasteToolStripButton.Name = "PasteToolStripButton"
        Me.PasteToolStripButton.Size = New System.Drawing.Size(28, 76)
        Me.PasteToolStripButton.Text = "&Paste"
        '
        'toolStripSeparator1
        '
        Me.toolStripSeparator1.Name = "toolStripSeparator1"
        Me.toolStripSeparator1.Size = New System.Drawing.Size(6, 79)
        '
        'HelpToolStripButton
        '
        Me.HelpToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.HelpToolStripButton.Image = CType(resources.GetObject("HelpToolStripButton.Image"), System.Drawing.Image)
        Me.HelpToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.HelpToolStripButton.Name = "HelpToolStripButton"
        Me.HelpToolStripButton.Size = New System.Drawing.Size(28, 76)
        Me.HelpToolStripButton.Text = "He&lp"
        '
        'RichTextBoxNodeOutput
        '
        Me.RichTextBoxNodeOutput.BackColor = System.Drawing.SystemColors.Info
        Me.RichTextBoxNodeOutput.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RichTextBoxNodeOutput.Font = New System.Drawing.Font("Consolas", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RichTextBoxNodeOutput.Location = New System.Drawing.Point(0, 0)
        Me.RichTextBoxNodeOutput.Margin = New System.Windows.Forms.Padding(4)
        Me.RichTextBoxNodeOutput.Name = "RichTextBoxNodeOutput"
        Me.RichTextBoxNodeOutput.ReadOnly = True
        Me.RichTextBoxNodeOutput.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.ForcedVertical
        Me.RichTextBoxNodeOutput.Size = New System.Drawing.Size(1611, 66)
        Me.RichTextBoxNodeOutput.TabIndex = 1
        Me.RichTextBoxNodeOutput.Text = ""
        Me.RichTextBoxNodeOutput.Visible = False
        '
        'InteractiveReplNode
        '
        Me.AllowDrop = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoSize = True
        Me.BackColor = System.Drawing.Color.Transparent
        Me.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Controls.Add(Me.SplitContainer1)
        Me.Font = New System.Drawing.Font("Consolas", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "InteractiveReplNode"
        Me.Size = New System.Drawing.Size(1611, 382)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.SplitContainer5.Panel1.ResumeLayout(False)
        Me.SplitContainer5.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer5.ResumeLayout(False)
        Me.SplitContainer6.Panel1.ResumeLayout(False)
        Me.SplitContainer6.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer6, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer6.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer3.ResumeLayout(False)
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel1.PerformLayout()
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        Me.SplitContainer4.Panel2.PerformLayout()
        CType(Me.SplitContainer4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer4.ResumeLayout(False)
        Me.SplitContainer7.Panel1.ResumeLayout(False)
        Me.SplitContainer7.Panel1.PerformLayout()
        CType(Me.SplitContainer7, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer7.ResumeLayout(False)
        Me.ToolStripStandard.ResumeLayout(False)
        Me.ToolStripStandard.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TextBoxNodeCodeInput As TextBox
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents SplitContainer2 As SplitContainer
    Friend WithEvents SplitContainer3 As SplitContainer
    Friend WithEvents SplitContainer4 As SplitContainer
    Friend WithEvents Code_Lable As Label
    Friend WithEvents RichTextBoxNodeOutput As RichTextBox
    Friend WithEvents ToolStripStandard As ToolStrip
    Friend WithEvents Panel1 As Panel
    Friend WithEvents NewToolStripButton As ToolStripButton
    Friend WithEvents OpenToolStripButton As ToolStripButton
    Friend WithEvents SaveToolStripButton As ToolStripButton
    Friend WithEvents PrintToolStripButton As ToolStripButton
    Friend WithEvents toolStripSeparator As ToolStripSeparator
    Friend WithEvents CutToolStripButton As ToolStripButton
    Friend WithEvents CopyToolStripButton As ToolStripButton
    Friend WithEvents PasteToolStripButton As ToolStripButton
    Friend WithEvents toolStripSeparator1 As ToolStripSeparator
    Friend WithEvents HelpToolStripButton As ToolStripButton
    Friend WithEvents SplitContainer5 As SplitContainer
    Friend WithEvents ButtonExecute As Button
    Friend WithEvents SplitContainer6 As SplitContainer
    Friend WithEvents ButtonAddCell As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents SplitContainer7 As SplitContainer
    Friend WithEvents RadioButtonIsMainCell As RadioButton
End Class
