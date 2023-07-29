Imports System.Windows.Forms

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ReplGrid
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ReplGrid))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.ToolStrip_MAIN = New System.Windows.Forms.ToolStrip()
        Me.ToolStripDropDownButton5 = New System.Windows.Forms.ToolStripDropDownButton()
        Me.MainClassNameToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripTextBoxMainClassName = New System.Windows.Forms.ToolStripTextBox()
        Me.ExecuteFunctionNAmeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripTextBoxFunctionName = New System.Windows.Forms.ToolStripTextBox()
        Me.SetProgrammingLanguageToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripComboBoxProgLang = New System.Windows.Forms.ToolStripComboBox()
        Me.ButtonBuild = New System.Windows.Forms.ToolStripDropDownButton()
        Me.ExecuteToolStripREPL = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DLLToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.GroupBoxCells = New System.Windows.Forms.GroupBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPageObjectBrowser = New System.Windows.Forms.TabPage()
        Me.GroupBox38 = New System.Windows.Forms.GroupBox()
        Me.RichTextBoxCompilerInfo = New System.Windows.Forms.RichTextBox()
        Me.GroupBox39 = New System.Windows.Forms.GroupBox()
        Me.TreeViewObjectSyntax = New System.Windows.Forms.TreeView()
        Me.TabPageEmbeddedFiles = New System.Windows.Forms.TabPage()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.CheckedListBoxEmbeddedFiles = New System.Windows.Forms.CheckedListBox()
        Me.ToolStripEmbeddedFiles = New System.Windows.Forms.ToolStrip()
        Me.AddEmbededFileToolStripMenuItem = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator6 = New System.Windows.Forms.ToolStripSeparator()
        Me.RemoveEmbeddedFileToolStripMenuItem = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator7 = New System.Windows.Forms.ToolStripSeparator()
        Me.ClearFilesToolStripMenuItem = New System.Windows.Forms.ToolStripButton()
        Me.TabPageAssemblies = New System.Windows.Forms.TabPage()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.CheckedListBoxAssemblys = New System.Windows.Forms.CheckedListBox()
        Me.ToolStripAssemblies = New System.Windows.Forms.ToolStrip()
        Me.AddAssemblyToolStripMenuItem = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.RemoveAssemblyToolStripMenuItem = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.ClearAssemblyToolStripMenuItem = New System.Windows.Forms.ToolStripButton()
        Me.TabPageInteractiveReplScript = New System.Windows.Forms.TabPage()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.InteractiveReplScript = New Controls.SpydazWebTextBox()
        Me.SpydazWebLineNumbering2 = New Controls.SpydazWebLineNumbering()
        Me.ToolStripInteractiveRepl = New System.Windows.Forms.ToolStrip()
        Me.ToolStripSeparator12 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripButtonExecuteInteractive = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator13 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripButtonClearInteractive = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripTextBoxClassNameInteractive = New System.Windows.Forms.ToolStripTextBox()
        Me.ToolStripTextBoxFunctionNameInteractive = New System.Windows.Forms.ToolStripTextBox()
        Me.ToolStripTextBox1 = New System.Windows.Forms.ToolStripTextBox()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.ToolStrip_MAIN.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPageObjectBrowser.SuspendLayout()
        Me.GroupBox38.SuspendLayout()
        Me.GroupBox39.SuspendLayout()
        Me.TabPageEmbeddedFiles.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.ToolStripEmbeddedFiles.SuspendLayout()
        Me.TabPageAssemblies.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.ToolStripAssemblies.SuspendLayout()
        Me.TabPageInteractiveReplScript.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.ToolStripInteractiveRepl.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.SplitContainer2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TabControl1)
        Me.SplitContainer1.Size = New System.Drawing.Size(1294, 420)
        Me.SplitContainer1.SplitterDistance = 655
        Me.SplitContainer1.TabIndex = 0
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.ToolStrip_MAIN)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.GroupBoxCells)
        Me.SplitContainer2.Size = New System.Drawing.Size(655, 420)
        Me.SplitContainer2.SplitterDistance = 39
        Me.SplitContainer2.TabIndex = 0
        '
        'ToolStrip_MAIN
        '
        Me.ToolStrip_MAIN.BackColor = System.Drawing.Color.Black
        Me.ToolStrip_MAIN.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ToolStrip_MAIN.Font = New System.Drawing.Font("Continuum Bold", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStrip_MAIN.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripDropDownButton5, Me.ButtonBuild, Me.ToolStripTextBox1, Me.ToolStripButton1})
        Me.ToolStrip_MAIN.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip_MAIN.Name = "ToolStrip_MAIN"
        Me.ToolStrip_MAIN.Size = New System.Drawing.Size(655, 39)
        Me.ToolStrip_MAIN.Stretch = True
        Me.ToolStrip_MAIN.TabIndex = 0
        '
        'ToolStripDropDownButton5
        '
        Me.ToolStripDropDownButton5.BackColor = System.Drawing.Color.Black
        Me.ToolStripDropDownButton5.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ToolStripDropDownButton5.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MainClassNameToolStripMenuItem, Me.ExecuteFunctionNAmeToolStripMenuItem, Me.SetProgrammingLanguageToolStripMenuItem})
        Me.ToolStripDropDownButton5.ForeColor = System.Drawing.Color.DimGray
        Me.ToolStripDropDownButton5.Image = My.Resources.Resources.APP_Tools
        Me.ToolStripDropDownButton5.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripDropDownButton5.Name = "ToolStripDropDownButton5"
        Me.ToolStripDropDownButton5.Size = New System.Drawing.Size(141, 36)
        Me.ToolStripDropDownButton5.Text = "Project Settings"
        '
        'MainClassNameToolStripMenuItem
        '
        Me.MainClassNameToolStripMenuItem.BackgroundImage = CType(resources.GetObject("MainClassNameToolStripMenuItem.BackgroundImage"), System.Drawing.Image)
        Me.MainClassNameToolStripMenuItem.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.MainClassNameToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripTextBoxMainClassName})
        Me.MainClassNameToolStripMenuItem.ForeColor = System.Drawing.Color.White
        Me.MainClassNameToolStripMenuItem.Image = Global.BasicUserControls.My.Resources.Resources.fileIconTextEdit
        Me.MainClassNameToolStripMenuItem.Name = "MainClassNameToolStripMenuItem"
        Me.MainClassNameToolStripMenuItem.Size = New System.Drawing.Size(262, 24)
        Me.MainClassNameToolStripMenuItem.Text = "Main ClassName"
        '
        'ToolStripTextBoxMainClassName
        '
        Me.ToolStripTextBoxMainClassName.BackColor = System.Drawing.Color.LavenderBlush
        Me.ToolStripTextBoxMainClassName.Font = New System.Drawing.Font("Comic Sans MS", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripTextBoxMainClassName.Name = "ToolStripTextBoxMainClassName"
        Me.ToolStripTextBoxMainClassName.Size = New System.Drawing.Size(200, 28)
        Me.ToolStripTextBoxMainClassName.Text = "Test"
        '
        'ExecuteFunctionNAmeToolStripMenuItem
        '
        Me.ExecuteFunctionNAmeToolStripMenuItem.BackgroundImage = CType(resources.GetObject("ExecuteFunctionNAmeToolStripMenuItem.BackgroundImage"), System.Drawing.Image)
        Me.ExecuteFunctionNAmeToolStripMenuItem.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ExecuteFunctionNAmeToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripTextBoxFunctionName})
        Me.ExecuteFunctionNAmeToolStripMenuItem.ForeColor = System.Drawing.Color.White
        Me.ExecuteFunctionNAmeToolStripMenuItem.Image = Global.BasicUserControls.My.Resources.Resources.fileIconTextEdit
        Me.ExecuteFunctionNAmeToolStripMenuItem.Name = "ExecuteFunctionNAmeToolStripMenuItem"
        Me.ExecuteFunctionNAmeToolStripMenuItem.Size = New System.Drawing.Size(262, 24)
        Me.ExecuteFunctionNAmeToolStripMenuItem.Text = "Execute FunctionName"
        '
        'ToolStripTextBoxFunctionName
        '
        Me.ToolStripTextBoxFunctionName.BackColor = System.Drawing.Color.LavenderBlush
        Me.ToolStripTextBoxFunctionName.Font = New System.Drawing.Font("Comic Sans MS", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripTextBoxFunctionName.Name = "ToolStripTextBoxFunctionName"
        Me.ToolStripTextBoxFunctionName.Size = New System.Drawing.Size(200, 28)
        Me.ToolStripTextBoxFunctionName.Text = "Main"
        '
        'SetProgrammingLanguageToolStripMenuItem
        '
        Me.SetProgrammingLanguageToolStripMenuItem.BackgroundImage = CType(resources.GetObject("SetProgrammingLanguageToolStripMenuItem.BackgroundImage"), System.Drawing.Image)
        Me.SetProgrammingLanguageToolStripMenuItem.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.SetProgrammingLanguageToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripComboBoxProgLang})
        Me.SetProgrammingLanguageToolStripMenuItem.ForeColor = System.Drawing.Color.White
        Me.SetProgrammingLanguageToolStripMenuItem.Image = Global.BasicUserControls.My.Resources.Resources.fileIconTextEdit
        Me.SetProgrammingLanguageToolStripMenuItem.Name = "SetProgrammingLanguageToolStripMenuItem"
        Me.SetProgrammingLanguageToolStripMenuItem.Size = New System.Drawing.Size(262, 24)
        Me.SetProgrammingLanguageToolStripMenuItem.Text = "Set Programming Language"
        '
        'ToolStripComboBoxProgLang
        '
        Me.ToolStripComboBoxProgLang.AutoCompleteCustomSource.AddRange(New String() {"VisualBasic"})
        Me.ToolStripComboBoxProgLang.BackColor = System.Drawing.Color.LavenderBlush
        Me.ToolStripComboBoxProgLang.Items.AddRange(New Object() {"Visual Basic", "C#"})
        Me.ToolStripComboBoxProgLang.Name = "ToolStripComboBoxProgLang"
        Me.ToolStripComboBoxProgLang.Size = New System.Drawing.Size(200, 25)
        '
        'ButtonBuild
        '
        Me.ButtonBuild.BackColor = System.Drawing.Color.Transparent
        Me.ButtonBuild.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ButtonBuild.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExecuteToolStripREPL, Me.ToolStripMenuItem2})
        Me.ButtonBuild.ForeColor = System.Drawing.Color.DimGray
        Me.ButtonBuild.Image = Global.BasicUserControls.My.Resources.Resources.App_Compile
        Me.ButtonBuild.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ButtonBuild.Name = "ButtonBuild"
        Me.ButtonBuild.Size = New System.Drawing.Size(73, 36)
        Me.ButtonBuild.Text = "Build"
        '
        'ExecuteToolStripREPL
        '
        Me.ExecuteToolStripREPL.BackgroundImage = CType(resources.GetObject("ExecuteToolStripREPL.BackgroundImage"), System.Drawing.Image)
        Me.ExecuteToolStripREPL.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ExecuteToolStripREPL.ForeColor = System.Drawing.Color.White
        Me.ExecuteToolStripREPL.Image = Global.BasicUserControls.My.Resources.Resources.APP_icon_Run
        Me.ExecuteToolStripREPL.Name = "ExecuteToolStripREPL"
        Me.ExecuteToolStripREPL.Size = New System.Drawing.Size(182, 24)
        Me.ExecuteToolStripREPL.Text = "Run Project"
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.BackgroundImage = CType(resources.GetObject("ToolStripMenuItem2.BackgroundImage"), System.Drawing.Image)
        Me.ToolStripMenuItem2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ToolStripMenuItem2.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExeToolStripMenuItem, Me.DLLToolStripMenuItem})
        Me.ToolStripMenuItem2.ForeColor = System.Drawing.Color.White
        Me.ToolStripMenuItem2.Image = Global.BasicUserControls.My.Resources.Resources.APP_icon_Complie
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(182, 24)
        Me.ToolStripMenuItem2.Text = "Compile Project"
        '
        'ExeToolStripMenuItem
        '
        Me.ExeToolStripMenuItem.BackColor = System.Drawing.Color.Black
        Me.ExeToolStripMenuItem.ForeColor = System.Drawing.Color.Silver
        Me.ExeToolStripMenuItem.Image = Global.BasicUserControls.My.Resources.Resources.APP_icon_SucessBuild
        Me.ExeToolStripMenuItem.Name = "ExeToolStripMenuItem"
        Me.ExeToolStripMenuItem.Size = New System.Drawing.Size(180, 24)
        Me.ExeToolStripMenuItem.Text = "EXE"
        '
        'DLLToolStripMenuItem
        '
        Me.DLLToolStripMenuItem.BackColor = System.Drawing.Color.Black
        Me.DLLToolStripMenuItem.ForeColor = System.Drawing.Color.Silver
        Me.DLLToolStripMenuItem.Image = Global.BasicUserControls.My.Resources.Resources.APP_icon_UnsucessBuild
        Me.DLLToolStripMenuItem.Name = "DLLToolStripMenuItem"
        Me.DLLToolStripMenuItem.Size = New System.Drawing.Size(180, 24)
        Me.DLLToolStripMenuItem.Text = "DLL"
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.BackColor = System.Drawing.Color.Lime
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton1.ForeColor = System.Drawing.Color.Gray
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(71, 36)
        Me.ToolStripButton1.Text = "Add Cell"
        '
        'GroupBoxCells
        '
        Me.GroupBoxCells.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBoxCells.Font = New System.Drawing.Font("Continuum Medium", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBoxCells.Location = New System.Drawing.Point(0, 0)
        Me.GroupBoxCells.Name = "GroupBoxCells"
        Me.GroupBoxCells.Size = New System.Drawing.Size(655, 377)
        Me.GroupBoxCells.TabIndex = 0
        Me.GroupBoxCells.TabStop = False
        Me.GroupBoxCells.Text = "Project"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPageObjectBrowser)
        Me.TabControl1.Controls.Add(Me.TabPageEmbeddedFiles)
        Me.TabControl1.Controls.Add(Me.TabPageAssemblies)
        Me.TabControl1.Controls.Add(Me.TabPageInteractiveReplScript)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(635, 420)
        Me.TabControl1.TabIndex = 0
        '
        'TabPageObjectBrowser
        '
        Me.TabPageObjectBrowser.Controls.Add(Me.GroupBox38)
        Me.TabPageObjectBrowser.Controls.Add(Me.GroupBox39)
        Me.TabPageObjectBrowser.Location = New System.Drawing.Point(4, 22)
        Me.TabPageObjectBrowser.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TabPageObjectBrowser.Name = "TabPageObjectBrowser"
        Me.TabPageObjectBrowser.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TabPageObjectBrowser.Size = New System.Drawing.Size(627, 394)
        Me.TabPageObjectBrowser.TabIndex = 3
        Me.TabPageObjectBrowser.Text = "ObjectBrowser"
        Me.TabPageObjectBrowser.UseVisualStyleBackColor = True
        '
        'GroupBox38
        '
        Me.GroupBox38.BackColor = System.Drawing.Color.Black
        Me.GroupBox38.BackgroundImage = Global.BasicUserControls.My.Resources.Resources.App_Texturex15
        Me.GroupBox38.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.GroupBox38.Controls.Add(Me.RichTextBoxCompilerInfo)
        Me.GroupBox38.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox38.Font = New System.Drawing.Font("Comic Sans MS", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox38.ForeColor = System.Drawing.Color.White
        Me.GroupBox38.Location = New System.Drawing.Point(4, 3)
        Me.GroupBox38.Margin = New System.Windows.Forms.Padding(8, 5, 8, 5)
        Me.GroupBox38.Name = "GroupBox38"
        Me.GroupBox38.Padding = New System.Windows.Forms.Padding(8, 5, 8, 5)
        Me.GroupBox38.Size = New System.Drawing.Size(619, 168)
        Me.GroupBox38.TabIndex = 3
        Me.GroupBox38.TabStop = False
        Me.GroupBox38.Text = "Info"
        '
        'RichTextBoxCompilerInfo
        '
        Me.RichTextBoxCompilerInfo.BackColor = System.Drawing.SystemColors.Info
        Me.RichTextBoxCompilerInfo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.RichTextBoxCompilerInfo.Cursor = System.Windows.Forms.Cursors.Default
        Me.RichTextBoxCompilerInfo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RichTextBoxCompilerInfo.EnableAutoDragDrop = True
        Me.RichTextBoxCompilerInfo.Font = New System.Drawing.Font("Microsoft Tai Le", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RichTextBoxCompilerInfo.Location = New System.Drawing.Point(8, 26)
        Me.RichTextBoxCompilerInfo.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.RichTextBoxCompilerInfo.Name = "RichTextBoxCompilerInfo"
        Me.RichTextBoxCompilerInfo.Size = New System.Drawing.Size(603, 137)
        Me.RichTextBoxCompilerInfo.TabIndex = 0
        Me.RichTextBoxCompilerInfo.Text = ""
        Me.RichTextBoxCompilerInfo.ZoomFactor = 2.0!
        '
        'GroupBox39
        '
        Me.GroupBox39.BackColor = System.Drawing.Color.Black
        Me.GroupBox39.BackgroundImage = Global.BasicUserControls.My.Resources.Resources.App_Texturex15
        Me.GroupBox39.Controls.Add(Me.TreeViewObjectSyntax)
        Me.GroupBox39.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.GroupBox39.Font = New System.Drawing.Font("Comic Sans MS", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox39.ForeColor = System.Drawing.Color.White
        Me.GroupBox39.Location = New System.Drawing.Point(4, 171)
        Me.GroupBox39.Margin = New System.Windows.Forms.Padding(8, 5, 8, 5)
        Me.GroupBox39.Name = "GroupBox39"
        Me.GroupBox39.Padding = New System.Windows.Forms.Padding(8, 5, 8, 5)
        Me.GroupBox39.Size = New System.Drawing.Size(619, 220)
        Me.GroupBox39.TabIndex = 2
        Me.GroupBox39.TabStop = False
        Me.GroupBox39.Text = "Loaded Objects"
        '
        'TreeViewObjectSyntax
        '
        Me.TreeViewObjectSyntax.BackColor = System.Drawing.Color.White
        Me.TreeViewObjectSyntax.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TreeViewObjectSyntax.Font = New System.Drawing.Font("Microsoft Tai Le", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TreeViewObjectSyntax.Location = New System.Drawing.Point(8, 26)
        Me.TreeViewObjectSyntax.Margin = New System.Windows.Forms.Padding(8, 5, 8, 5)
        Me.TreeViewObjectSyntax.Name = "TreeViewObjectSyntax"
        Me.TreeViewObjectSyntax.Size = New System.Drawing.Size(603, 189)
        Me.TreeViewObjectSyntax.TabIndex = 0
        '
        'TabPageEmbeddedFiles
        '
        Me.TabPageEmbeddedFiles.Controls.Add(Me.GroupBox2)
        Me.TabPageEmbeddedFiles.Location = New System.Drawing.Point(4, 22)
        Me.TabPageEmbeddedFiles.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TabPageEmbeddedFiles.Name = "TabPageEmbeddedFiles"
        Me.TabPageEmbeddedFiles.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TabPageEmbeddedFiles.Size = New System.Drawing.Size(627, 394)
        Me.TabPageEmbeddedFiles.TabIndex = 4
        Me.TabPageEmbeddedFiles.Text = "EmbeddedFiles"
        Me.TabPageEmbeddedFiles.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.Black
        Me.GroupBox2.BackgroundImage = Global.BasicUserControls.My.Resources.Resources.App_Texturex15
        Me.GroupBox2.Controls.Add(Me.CheckedListBoxEmbeddedFiles)
        Me.GroupBox2.Controls.Add(Me.ToolStripEmbeddedFiles)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Font = New System.Drawing.Font("Comic Sans MS", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.Color.White
        Me.GroupBox2.Location = New System.Drawing.Point(4, 3)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.GroupBox2.Size = New System.Drawing.Size(619, 388)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Embeded Resources"
        '
        'CheckedListBoxEmbeddedFiles
        '
        Me.CheckedListBoxEmbeddedFiles.BackColor = System.Drawing.SystemColors.Info
        Me.CheckedListBoxEmbeddedFiles.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CheckedListBoxEmbeddedFiles.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CheckedListBoxEmbeddedFiles.Font = New System.Drawing.Font("Comic Sans MS", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckedListBoxEmbeddedFiles.FormattingEnabled = True
        Me.CheckedListBoxEmbeddedFiles.HorizontalScrollbar = True
        Me.CheckedListBoxEmbeddedFiles.Location = New System.Drawing.Point(4, 63)
        Me.CheckedListBoxEmbeddedFiles.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.CheckedListBoxEmbeddedFiles.Name = "CheckedListBoxEmbeddedFiles"
        Me.CheckedListBoxEmbeddedFiles.ScrollAlwaysVisible = True
        Me.CheckedListBoxEmbeddedFiles.Size = New System.Drawing.Size(611, 322)
        Me.CheckedListBoxEmbeddedFiles.TabIndex = 10
        '
        'ToolStripEmbeddedFiles
        '
        Me.ToolStripEmbeddedFiles.BackColor = System.Drawing.Color.Black
        Me.ToolStripEmbeddedFiles.BackgroundImage = Global.BasicUserControls.My.Resources.Resources.App_Texturex15
        Me.ToolStripEmbeddedFiles.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ToolStripEmbeddedFiles.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.ToolStripEmbeddedFiles.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AddEmbededFileToolStripMenuItem, Me.ToolStripSeparator6, Me.RemoveEmbeddedFileToolStripMenuItem, Me.ToolStripSeparator7, Me.ClearFilesToolStripMenuItem})
        Me.ToolStripEmbeddedFiles.Location = New System.Drawing.Point(4, 24)
        Me.ToolStripEmbeddedFiles.Name = "ToolStripEmbeddedFiles"
        Me.ToolStripEmbeddedFiles.Size = New System.Drawing.Size(611, 39)
        Me.ToolStripEmbeddedFiles.TabIndex = 9
        Me.ToolStripEmbeddedFiles.Text = "Application"
        '
        'AddEmbededFileToolStripMenuItem
        '
        Me.AddEmbededFileToolStripMenuItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.AddEmbededFileToolStripMenuItem.Image = Global.BasicUserControls.My.Resources.Resources.APP_icon_fileopen
        Me.AddEmbededFileToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.AddEmbededFileToolStripMenuItem.Name = "AddEmbededFileToolStripMenuItem"
        Me.AddEmbededFileToolStripMenuItem.Size = New System.Drawing.Size(36, 36)
        Me.AddEmbededFileToolStripMenuItem.Text = "&Open"
        '
        'ToolStripSeparator6
        '
        Me.ToolStripSeparator6.Name = "ToolStripSeparator6"
        Me.ToolStripSeparator6.Size = New System.Drawing.Size(6, 39)
        '
        'RemoveEmbeddedFileToolStripMenuItem
        '
        Me.RemoveEmbeddedFileToolStripMenuItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.RemoveEmbeddedFileToolStripMenuItem.Image = Global.BasicUserControls.My.Resources.Resources.APP_icon_error
        Me.RemoveEmbeddedFileToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.RemoveEmbeddedFileToolStripMenuItem.Name = "RemoveEmbeddedFileToolStripMenuItem"
        Me.RemoveEmbeddedFileToolStripMenuItem.Size = New System.Drawing.Size(36, 36)
        Me.RemoveEmbeddedFileToolStripMenuItem.Text = "&Removed Selected"
        '
        'ToolStripSeparator7
        '
        Me.ToolStripSeparator7.Name = "ToolStripSeparator7"
        Me.ToolStripSeparator7.Size = New System.Drawing.Size(6, 39)
        '
        'ClearFilesToolStripMenuItem
        '
        Me.ClearFilesToolStripMenuItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ClearFilesToolStripMenuItem.Image = Global.BasicUserControls.My.Resources.Resources.App_Refresh
        Me.ClearFilesToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ClearFilesToolStripMenuItem.Name = "ClearFilesToolStripMenuItem"
        Me.ClearFilesToolStripMenuItem.Size = New System.Drawing.Size(36, 36)
        Me.ClearFilesToolStripMenuItem.Text = "Clear"
        '
        'TabPageAssemblies
        '
        Me.TabPageAssemblies.Controls.Add(Me.GroupBox1)
        Me.TabPageAssemblies.Location = New System.Drawing.Point(4, 22)
        Me.TabPageAssemblies.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TabPageAssemblies.Name = "TabPageAssemblies"
        Me.TabPageAssemblies.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TabPageAssemblies.Size = New System.Drawing.Size(627, 394)
        Me.TabPageAssemblies.TabIndex = 5
        Me.TabPageAssemblies.Text = "Assemblies Files"
        Me.TabPageAssemblies.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Black
        Me.GroupBox1.BackgroundImage = Global.BasicUserControls.My.Resources.Resources.App_Texturex15
        Me.GroupBox1.Controls.Add(Me.CheckedListBoxAssemblys)
        Me.GroupBox1.Controls.Add(Me.ToolStripAssemblies)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Font = New System.Drawing.Font("Comic Sans MS", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.White
        Me.GroupBox1.Location = New System.Drawing.Point(4, 3)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.GroupBox1.Size = New System.Drawing.Size(619, 388)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Assemblies"
        '
        'CheckedListBoxAssemblys
        '
        Me.CheckedListBoxAssemblys.BackColor = System.Drawing.SystemColors.Info
        Me.CheckedListBoxAssemblys.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CheckedListBoxAssemblys.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CheckedListBoxAssemblys.Font = New System.Drawing.Font("Comic Sans MS", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckedListBoxAssemblys.FormattingEnabled = True
        Me.CheckedListBoxAssemblys.HorizontalScrollbar = True
        Me.CheckedListBoxAssemblys.Location = New System.Drawing.Point(4, 63)
        Me.CheckedListBoxAssemblys.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.CheckedListBoxAssemblys.Name = "CheckedListBoxAssemblys"
        Me.CheckedListBoxAssemblys.ScrollAlwaysVisible = True
        Me.CheckedListBoxAssemblys.Size = New System.Drawing.Size(611, 322)
        Me.CheckedListBoxAssemblys.TabIndex = 10
        '
        'ToolStripAssemblies
        '
        Me.ToolStripAssemblies.BackColor = System.Drawing.Color.Black
        Me.ToolStripAssemblies.BackgroundImage = Global.BasicUserControls.My.Resources.Resources.App_Texturex15
        Me.ToolStripAssemblies.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ToolStripAssemblies.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.ToolStripAssemblies.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AddAssemblyToolStripMenuItem, Me.ToolStripSeparator4, Me.RemoveAssemblyToolStripMenuItem, Me.ToolStripSeparator5, Me.ClearAssemblyToolStripMenuItem})
        Me.ToolStripAssemblies.Location = New System.Drawing.Point(4, 24)
        Me.ToolStripAssemblies.Name = "ToolStripAssemblies"
        Me.ToolStripAssemblies.Size = New System.Drawing.Size(611, 39)
        Me.ToolStripAssemblies.TabIndex = 9
        Me.ToolStripAssemblies.Text = "Application"
        '
        'AddAssemblyToolStripMenuItem
        '
        Me.AddAssemblyToolStripMenuItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.AddAssemblyToolStripMenuItem.Image = Global.BasicUserControls.My.Resources.Resources.APP_icon_fileopen
        Me.AddAssemblyToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.AddAssemblyToolStripMenuItem.Name = "AddAssemblyToolStripMenuItem"
        Me.AddAssemblyToolStripMenuItem.Size = New System.Drawing.Size(36, 36)
        Me.AddAssemblyToolStripMenuItem.Text = "&Open"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(6, 39)
        '
        'RemoveAssemblyToolStripMenuItem
        '
        Me.RemoveAssemblyToolStripMenuItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.RemoveAssemblyToolStripMenuItem.Image = Global.BasicUserControls.My.Resources.Resources.APP_icon_error
        Me.RemoveAssemblyToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.RemoveAssemblyToolStripMenuItem.Name = "RemoveAssemblyToolStripMenuItem"
        Me.RemoveAssemblyToolStripMenuItem.Size = New System.Drawing.Size(36, 36)
        Me.RemoveAssemblyToolStripMenuItem.Text = "&Remove Selected"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(6, 39)
        '
        'ClearAssemblyToolStripMenuItem
        '
        Me.ClearAssemblyToolStripMenuItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ClearAssemblyToolStripMenuItem.Image = Global.BasicUserControls.My.Resources.Resources.App_Refresh
        Me.ClearAssemblyToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ClearAssemblyToolStripMenuItem.Name = "ClearAssemblyToolStripMenuItem"
        Me.ClearAssemblyToolStripMenuItem.Size = New System.Drawing.Size(36, 36)
        Me.ClearAssemblyToolStripMenuItem.Text = "Clear"
        '
        'TabPageInteractiveReplScript
        '
        Me.TabPageInteractiveReplScript.Controls.Add(Me.GroupBox4)
        Me.TabPageInteractiveReplScript.Location = New System.Drawing.Point(4, 22)
        Me.TabPageInteractiveReplScript.Name = "TabPageInteractiveReplScript"
        Me.TabPageInteractiveReplScript.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPageInteractiveReplScript.Size = New System.Drawing.Size(627, 394)
        Me.TabPageInteractiveReplScript.TabIndex = 6
        Me.TabPageInteractiveReplScript.Text = "Interactive REPL"
        Me.TabPageInteractiveReplScript.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.BackColor = System.Drawing.Color.Black
        Me.GroupBox4.BackgroundImage = Global.BasicUserControls.My.Resources.Resources.App_Texturex15
        Me.GroupBox4.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.GroupBox4.Controls.Add(Me.InteractiveReplScript)
        Me.GroupBox4.Controls.Add(Me.SpydazWebLineNumbering2)
        Me.GroupBox4.Controls.Add(Me.ToolStripInteractiveRepl)
        Me.GroupBox4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox4.Font = New System.Drawing.Font("Comic Sans MS", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.ForeColor = System.Drawing.Color.White
        Me.GroupBox4.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox4.Margin = New System.Windows.Forms.Padding(8, 5, 8, 5)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Padding = New System.Windows.Forms.Padding(8, 5, 8, 5)
        Me.GroupBox4.Size = New System.Drawing.Size(621, 388)
        Me.GroupBox4.TabIndex = 2
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Interactive VB Repl"
        '
        'InteractiveReplScript
        '
        Me.InteractiveReplScript.BackColor = System.Drawing.Color.Ivory
        Me.InteractiveReplScript.Dock = System.Windows.Forms.DockStyle.Fill
        Me.InteractiveReplScript.Font = New System.Drawing.Font("Courier-Oblique", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.InteractiveReplScript.Location = New System.Drawing.Point(58, 65)
        Me.InteractiveReplScript.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.InteractiveReplScript.MySyntax = CType(resources.GetObject("InteractiveReplScript.MySyntax"), System.Collections.Generic.List(Of String))
        Me.InteractiveReplScript.Name = "InteractiveReplScript"
        Me.InteractiveReplScript.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.ForcedBoth
        Me.InteractiveReplScript.Size = New System.Drawing.Size(555, 318)
        Me.InteractiveReplScript.TabIndex = 13
        Me.InteractiveReplScript.Text = "Imports System" & Global.Microsoft.VisualBasic.ChrW(10) & "Imports system.windows.forms" & Global.Microsoft.VisualBasic.ChrW(10) & "Public Class Test" & Global.Microsoft.VisualBasic.ChrW(10) & "   Public Function " &
    "ExecuteCode()" & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(10) & "Console.WriteLine(""Hello"")" & Global.Microsoft.VisualBasic.ChrW(10) & "MessageBox.Show(""Welcome"")" & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(10) & "    End Fu" &
    "nction" & Global.Microsoft.VisualBasic.ChrW(10) & "End Class" & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'SpydazWebLineNumbering2
        '
        Me.SpydazWebLineNumbering2._SeeThroughMode_ = False
        Me.SpydazWebLineNumbering2.AutoSizing = True
        Me.SpydazWebLineNumbering2.BackgroundGradient_AlphaColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.SpydazWebLineNumbering2.BackgroundGradient_BetaColor = System.Drawing.Color.LightSteelBlue
        Me.SpydazWebLineNumbering2.BackgroundGradient_Direction = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.SpydazWebLineNumbering2.BorderLines_Color = System.Drawing.Color.SlateGray
        Me.SpydazWebLineNumbering2.BorderLines_Style = System.Drawing.Drawing2D.DashStyle.Dot
        Me.SpydazWebLineNumbering2.BorderLines_Thickness = 1.0!
        Me.SpydazWebLineNumbering2.Dock = System.Windows.Forms.DockStyle.Left
        Me.SpydazWebLineNumbering2.DockSide = BasicUserControls.Controls.SpydazWebLineNumbering.LineNumberDockSide.Left
        Me.SpydazWebLineNumbering2.GridLines_Color = System.Drawing.Color.SlateGray
        Me.SpydazWebLineNumbering2.GridLines_Style = System.Drawing.Drawing2D.DashStyle.Dot
        Me.SpydazWebLineNumbering2.GridLines_Thickness = 1.0!
        Me.SpydazWebLineNumbering2.LineNrs_Alignment = System.Drawing.ContentAlignment.TopRight
        Me.SpydazWebLineNumbering2.LineNrs_AntiAlias = True
        Me.SpydazWebLineNumbering2.LineNrs_AsHexadecimal = False
        Me.SpydazWebLineNumbering2.LineNrs_ClippedByItemRectangle = True
        Me.SpydazWebLineNumbering2.LineNrs_LeadingZeroes = True
        Me.SpydazWebLineNumbering2.LineNrs_Offset = New System.Drawing.Size(0, 0)
        Me.SpydazWebLineNumbering2.Location = New System.Drawing.Point(8, 65)
        Me.SpydazWebLineNumbering2.Margin = New System.Windows.Forms.Padding(0)
        Me.SpydazWebLineNumbering2.MarginLines_Color = System.Drawing.Color.SlateGray
        Me.SpydazWebLineNumbering2.MarginLines_Side = BasicUserControls.Controls.SpydazWebLineNumbering.LineNumberDockSide.Right
        Me.SpydazWebLineNumbering2.MarginLines_Style = System.Drawing.Drawing2D.DashStyle.Solid
        Me.SpydazWebLineNumbering2.MarginLines_Thickness = 1.0!
        Me.SpydazWebLineNumbering2.Name = "SpydazWebLineNumbering2"
        Me.SpydazWebLineNumbering2.Padding = New System.Windows.Forms.Padding(0, 0, 2, 0)
        Me.SpydazWebLineNumbering2.ParentRichTextBox = Me.InteractiveReplScript
        Me.SpydazWebLineNumbering2.Show_BackgroundGradient = True
        Me.SpydazWebLineNumbering2.Show_BorderLines = True
        Me.SpydazWebLineNumbering2.Show_GridLines = True
        Me.SpydazWebLineNumbering2.Show_LineNrs = True
        Me.SpydazWebLineNumbering2.Show_MarginLines = True
        Me.SpydazWebLineNumbering2.Size = New System.Drawing.Size(50, 318)
        Me.SpydazWebLineNumbering2.TabIndex = 12
        '
        'ToolStripInteractiveRepl
        '
        Me.ToolStripInteractiveRepl.BackColor = System.Drawing.Color.Black
        Me.ToolStripInteractiveRepl.BackgroundImage = Global.BasicUserControls.My.Resources.Resources.App_Texturex15
        Me.ToolStripInteractiveRepl.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ToolStripInteractiveRepl.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.ToolStripInteractiveRepl.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSeparator12, Me.ToolStripButtonExecuteInteractive, Me.ToolStripSeparator13, Me.ToolStripButtonClearInteractive, Me.ToolStripTextBoxClassNameInteractive, Me.ToolStripTextBoxFunctionNameInteractive})
        Me.ToolStripInteractiveRepl.Location = New System.Drawing.Point(8, 26)
        Me.ToolStripInteractiveRepl.Name = "ToolStripInteractiveRepl"
        Me.ToolStripInteractiveRepl.Size = New System.Drawing.Size(605, 39)
        Me.ToolStripInteractiveRepl.TabIndex = 10
        Me.ToolStripInteractiveRepl.Text = "Application"
        '
        'ToolStripSeparator12
        '
        Me.ToolStripSeparator12.Name = "ToolStripSeparator12"
        Me.ToolStripSeparator12.Size = New System.Drawing.Size(6, 39)
        '
        'ToolStripButtonExecuteInteractive
        '
        Me.ToolStripButtonExecuteInteractive.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonExecuteInteractive.Image = Global.BasicUserControls.My.Resources.Resources.APP_icon_Complie
        Me.ToolStripButtonExecuteInteractive.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButtonExecuteInteractive.Name = "ToolStripButtonExecuteInteractive"
        Me.ToolStripButtonExecuteInteractive.Size = New System.Drawing.Size(36, 36)
        Me.ToolStripButtonExecuteInteractive.Text = "&Execute"
        '
        'ToolStripSeparator13
        '
        Me.ToolStripSeparator13.Name = "ToolStripSeparator13"
        Me.ToolStripSeparator13.Size = New System.Drawing.Size(6, 39)
        '
        'ToolStripButtonClearInteractive
        '
        Me.ToolStripButtonClearInteractive.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonClearInteractive.Image = Global.BasicUserControls.My.Resources.Resources.App_Refresh
        Me.ToolStripButtonClearInteractive.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButtonClearInteractive.Name = "ToolStripButtonClearInteractive"
        Me.ToolStripButtonClearInteractive.Size = New System.Drawing.Size(36, 36)
        Me.ToolStripButtonClearInteractive.Text = "Clear"
        '
        'ToolStripTextBoxClassNameInteractive
        '
        Me.ToolStripTextBoxClassNameInteractive.BackColor = System.Drawing.Color.LavenderBlush
        Me.ToolStripTextBoxClassNameInteractive.Font = New System.Drawing.Font("Comic Sans MS", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripTextBoxClassNameInteractive.Name = "ToolStripTextBoxClassNameInteractive"
        Me.ToolStripTextBoxClassNameInteractive.Size = New System.Drawing.Size(200, 39)
        Me.ToolStripTextBoxClassNameInteractive.Tag = "ClassName"
        Me.ToolStripTextBoxClassNameInteractive.Text = "Test"
        '
        'ToolStripTextBoxFunctionNameInteractive
        '
        Me.ToolStripTextBoxFunctionNameInteractive.BackColor = System.Drawing.Color.LavenderBlush
        Me.ToolStripTextBoxFunctionNameInteractive.Font = New System.Drawing.Font("Comic Sans MS", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripTextBoxFunctionNameInteractive.Name = "ToolStripTextBoxFunctionNameInteractive"
        Me.ToolStripTextBoxFunctionNameInteractive.Size = New System.Drawing.Size(200, 39)
        Me.ToolStripTextBoxFunctionNameInteractive.Text = "ExecuteCode"
        '
        'ToolStripTextBox1
        '
        Me.ToolStripTextBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ToolStripTextBox1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.ToolStripTextBox1.Name = "ToolStripTextBox1"
        Me.ToolStripTextBox1.Size = New System.Drawing.Size(100, 39)
        Me.ToolStripTextBox1.Text = "ProjectName"
        '
        'ReplGrid
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Controls.Add(Me.SplitContainer1)
        Me.ForeColor = System.Drawing.Color.White
        Me.Name = "ReplGrid"
        Me.Size = New System.Drawing.Size(1294, 420)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.ToolStrip_MAIN.ResumeLayout(False)
        Me.ToolStrip_MAIN.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPageObjectBrowser.ResumeLayout(False)
        Me.GroupBox38.ResumeLayout(False)
        Me.GroupBox39.ResumeLayout(False)
        Me.TabPageEmbeddedFiles.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ToolStripEmbeddedFiles.ResumeLayout(False)
        Me.ToolStripEmbeddedFiles.PerformLayout()
        Me.TabPageAssemblies.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ToolStripAssemblies.ResumeLayout(False)
        Me.ToolStripAssemblies.PerformLayout()
        Me.TabPageInteractiveReplScript.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.ToolStripInteractiveRepl.ResumeLayout(False)
        Me.ToolStripInteractiveRepl.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents SplitContainer2 As SplitContainer
    Friend WithEvents ToolStrip_MAIN As ToolStrip
    Friend WithEvents ToolStripDropDownButton5 As ToolStripDropDownButton
    Friend WithEvents MainClassNameToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripTextBoxMainClassName As ToolStripTextBox
    Friend WithEvents ExecuteFunctionNAmeToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripTextBoxFunctionName As ToolStripTextBox
    Friend WithEvents SetProgrammingLanguageToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripComboBoxProgLang As ToolStripComboBox
    Friend WithEvents ButtonBuild As ToolStripDropDownButton
    Friend WithEvents ExecuteToolStripREPL As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As ToolStripMenuItem
    Friend WithEvents ExeToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents DLLToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripButton1 As ToolStripButton
    Friend WithEvents GroupBoxCells As GroupBox
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPageObjectBrowser As TabPage
    Friend WithEvents GroupBox38 As GroupBox
    Friend WithEvents RichTextBoxCompilerInfo As RichTextBox
    Friend WithEvents GroupBox39 As GroupBox
    Friend WithEvents TreeViewObjectSyntax As TreeView
    Friend WithEvents TabPageEmbeddedFiles As TabPage
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents CheckedListBoxEmbeddedFiles As CheckedListBox
    Friend WithEvents ToolStripEmbeddedFiles As ToolStrip
    Friend WithEvents AddEmbededFileToolStripMenuItem As ToolStripButton
    Friend WithEvents ToolStripSeparator6 As ToolStripSeparator
    Friend WithEvents RemoveEmbeddedFileToolStripMenuItem As ToolStripButton
    Friend WithEvents ToolStripSeparator7 As ToolStripSeparator
    Friend WithEvents ClearFilesToolStripMenuItem As ToolStripButton
    Friend WithEvents TabPageAssemblies As TabPage
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents CheckedListBoxAssemblys As CheckedListBox
    Friend WithEvents ToolStripAssemblies As ToolStrip
    Friend WithEvents AddAssemblyToolStripMenuItem As ToolStripButton
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
    Friend WithEvents RemoveAssemblyToolStripMenuItem As ToolStripButton
    Friend WithEvents ToolStripSeparator5 As ToolStripSeparator
    Friend WithEvents ClearAssemblyToolStripMenuItem As ToolStripButton
    Friend WithEvents TabPageInteractiveReplScript As TabPage
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents InteractiveReplScript As BasicUserControls.Controls.SpydazWebTextBox
    Friend WithEvents SpydazWebLineNumbering2 As BasicUserControls.Controls.SpydazWebLineNumbering
    Friend WithEvents ToolStripInteractiveRepl As ToolStrip
    Friend WithEvents ToolStripSeparator12 As ToolStripSeparator
    Friend WithEvents ToolStripButtonExecuteInteractive As ToolStripButton
    Friend WithEvents ToolStripSeparator13 As ToolStripSeparator
    Friend WithEvents ToolStripButtonClearInteractive As ToolStripButton
    Friend WithEvents ToolStripTextBoxClassNameInteractive As ToolStripTextBox
    Friend WithEvents ToolStripTextBoxFunctionNameInteractive As ToolStripTextBox
    Friend WithEvents ToolStripTextBox1 As ToolStripTextBox
End Class
