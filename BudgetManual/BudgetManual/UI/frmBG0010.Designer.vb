<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBG0010
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBG0010))
        Dim TreeNode1 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Menu")
        Me.MenuStrip = New System.Windows.Forms.MenuStrip
        Me.FileMenu = New System.Windows.Forms.ToolStripMenuItem
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolsMenu = New System.Windows.Forms.ToolStripMenuItem
        Me.OptionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.WindowsMenu = New System.Windows.Forms.ToolStripMenuItem
        Me.CascadeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.TileVerticalToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.TileHorizontalToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CloseAllToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.HelpMenu = New System.Windows.Forms.ToolStripMenuItem
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.StatusStrip = New System.Windows.Forms.StatusStrip
        Me.lblToolbarStatus = New System.Windows.Forms.ToolStripStatusLabel
        Me.lblCurrMode = New System.Windows.Forms.ToolStripStatusLabel
        Me.lblUserName = New System.Windows.Forms.ToolStripStatusLabel
        Me.lblPIC = New System.Windows.Forms.ToolStripStatusLabel
        Me.lblProgramVersion = New System.Windows.Forms.ToolStripStatusLabel
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.trvMenu = New System.Windows.Forms.TreeView
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.MenuStrip.SuspendLayout()
        Me.StatusStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip
        '
        Me.MenuStrip.AllowMerge = False
        Me.MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileMenu, Me.ToolsMenu, Me.WindowsMenu, Me.HelpMenu})
        Me.MenuStrip.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow
        Me.MenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip.MdiWindowListItem = Me.WindowsMenu
        Me.MenuStrip.Name = "MenuStrip"
        Me.MenuStrip.Size = New System.Drawing.Size(1016, 24)
        Me.MenuStrip.TabIndex = 0
        Me.MenuStrip.TabStop = True
        Me.MenuStrip.Text = "MenuStrip"
        '
        'FileMenu
        '
        Me.FileMenu.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExitToolStripMenuItem})
        Me.FileMenu.ImageTransparentColor = System.Drawing.SystemColors.ActiveBorder
        Me.FileMenu.Name = "FileMenu"
        Me.FileMenu.Size = New System.Drawing.Size(35, 20)
        Me.FileMenu.Text = "&File"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(92, 22)
        Me.ExitToolStripMenuItem.Text = "E&xit"
        '
        'ToolsMenu
        '
        Me.ToolsMenu.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OptionsToolStripMenuItem})
        Me.ToolsMenu.Enabled = False
        Me.ToolsMenu.Name = "ToolsMenu"
        Me.ToolsMenu.Size = New System.Drawing.Size(44, 20)
        Me.ToolsMenu.Text = "&Tools"
        '
        'OptionsToolStripMenuItem
        '
        Me.OptionsToolStripMenuItem.Name = "OptionsToolStripMenuItem"
        Me.OptionsToolStripMenuItem.Size = New System.Drawing.Size(111, 22)
        Me.OptionsToolStripMenuItem.Text = "&Options"
        '
        'WindowsMenu
        '
        Me.WindowsMenu.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CascadeToolStripMenuItem, Me.TileVerticalToolStripMenuItem, Me.TileHorizontalToolStripMenuItem, Me.CloseAllToolStripMenuItem})
        Me.WindowsMenu.Name = "WindowsMenu"
        Me.WindowsMenu.Size = New System.Drawing.Size(62, 20)
        Me.WindowsMenu.Text = "&Windows"
        '
        'CascadeToolStripMenuItem
        '
        Me.CascadeToolStripMenuItem.Name = "CascadeToolStripMenuItem"
        Me.CascadeToolStripMenuItem.Size = New System.Drawing.Size(141, 22)
        Me.CascadeToolStripMenuItem.Text = "&Cascade"
        '
        'TileVerticalToolStripMenuItem
        '
        Me.TileVerticalToolStripMenuItem.Name = "TileVerticalToolStripMenuItem"
        Me.TileVerticalToolStripMenuItem.Size = New System.Drawing.Size(141, 22)
        Me.TileVerticalToolStripMenuItem.Text = "Tile &Vertical"
        '
        'TileHorizontalToolStripMenuItem
        '
        Me.TileHorizontalToolStripMenuItem.Name = "TileHorizontalToolStripMenuItem"
        Me.TileHorizontalToolStripMenuItem.Size = New System.Drawing.Size(141, 22)
        Me.TileHorizontalToolStripMenuItem.Text = "Tile &Horizontal"
        '
        'CloseAllToolStripMenuItem
        '
        Me.CloseAllToolStripMenuItem.Name = "CloseAllToolStripMenuItem"
        Me.CloseAllToolStripMenuItem.Size = New System.Drawing.Size(141, 22)
        Me.CloseAllToolStripMenuItem.Text = "C&lose All"
        '
        'HelpMenu
        '
        Me.HelpMenu.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutToolStripMenuItem})
        Me.HelpMenu.Name = "HelpMenu"
        Me.HelpMenu.Size = New System.Drawing.Size(40, 20)
        Me.HelpMenu.Text = "&Help"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(118, 22)
        Me.AboutToolStripMenuItem.Text = "&About ..."
        '
        'StatusStrip
        '
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.lblToolbarStatus, Me.lblCurrMode, Me.lblUserName, Me.lblPIC, Me.lblProgramVersion})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 673)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.ShowItemToolTips = True
        Me.StatusStrip.Size = New System.Drawing.Size(1016, 22)
        Me.StatusStrip.TabIndex = 7
        Me.StatusStrip.Text = "StatusStrip"
        '
        'lblToolbarStatus
        '
        Me.lblToolbarStatus.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenInner
        Me.lblToolbarStatus.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.lblToolbarStatus.Name = "lblToolbarStatus"
        Me.lblToolbarStatus.Size = New System.Drawing.Size(728, 17)
        Me.lblToolbarStatus.Spring = True
        Me.lblToolbarStatus.Text = "Ready"
        Me.lblToolbarStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblToolbarStatus.ToolTipText = "Current Status"
        '
        'lblCurrMode
        '
        Me.lblCurrMode.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenInner
        Me.lblCurrMode.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.lblCurrMode.Name = "lblCurrMode"
        Me.lblCurrMode.Size = New System.Drawing.Size(81, 17)
        Me.lblCurrMode.Text = "[Current Mode]"
        Me.lblCurrMode.ToolTipText = "Current Mode"
        '
        'lblUserName
        '
        Me.lblUserName.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenInner
        Me.lblUserName.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.lblUserName.Name = "lblUserName"
        Me.lblUserName.Size = New System.Drawing.Size(59, 17)
        Me.lblUserName.Text = "User Name"
        Me.lblUserName.ToolTipText = "User Name"
        '
        'lblPIC
        '
        Me.lblPIC.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenInner
        Me.lblPIC.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.lblPIC.Name = "lblPIC"
        Me.lblPIC.Size = New System.Drawing.Size(91, 17)
        Me.lblPIC.Text = "Person In Charge"
        Me.lblPIC.ToolTipText = "Person In Charge"
        '
        'lblProgramVersion
        '
        Me.lblProgramVersion.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenInner
        Me.lblProgramVersion.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.lblProgramVersion.Name = "lblProgramVersion"
        Me.lblProgramVersion.Size = New System.Drawing.Size(42, 17)
        Me.lblProgramVersion.Text = "Version"
        Me.lblProgramVersion.ToolTipText = "Program Version"
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "document.ico")
        Me.ImageList1.Images.SetKeyName(1, "document_blank.ico")
        Me.ImageList1.Images.SetKeyName(2, "favorite.ico")
        Me.ImageList1.Images.SetKeyName(3, "folder.ico")
        Me.ImageList1.Images.SetKeyName(4, "home.ico")
        Me.ImageList1.Images.SetKeyName(5, "info.ico")
        Me.ImageList1.Images.SetKeyName(6, "journal.ico")
        Me.ImageList1.Images.SetKeyName(7, "opened_folder.ico")
        Me.ImageList1.Images.SetKeyName(8, "image.ico")
        Me.ImageList1.Images.SetKeyName(9, "PowerPoint.ico")
        Me.ImageList1.Images.SetKeyName(10, "Word.ico")
        Me.ImageList1.Images.SetKeyName(11, "Excel.ico")
        Me.ImageList1.Images.SetKeyName(12, "window.ico")
        Me.ImageList1.Images.SetKeyName(13, "pdf.ico")
        Me.ImageList1.Images.SetKeyName(14, "none.ico")
        '
        'trvMenu
        '
        Me.trvMenu.BackColor = System.Drawing.SystemColors.Window
        Me.trvMenu.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.trvMenu.Dock = System.Windows.Forms.DockStyle.Left
        Me.trvMenu.HideSelection = False
        Me.trvMenu.ImageIndex = 0
        Me.trvMenu.ImageList = Me.ImageList1
        Me.trvMenu.Indent = 20
        Me.trvMenu.ItemHeight = 18
        Me.trvMenu.Location = New System.Drawing.Point(0, 24)
        Me.trvMenu.Margin = New System.Windows.Forms.Padding(2)
        Me.trvMenu.Name = "trvMenu"
        TreeNode1.ImageKey = "folder.ico"
        TreeNode1.Name = "nodMenu"
        TreeNode1.NodeFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        TreeNode1.SelectedImageKey = "opened_folder.ico"
        TreeNode1.Text = "Menu"
        Me.trvMenu.Nodes.AddRange(New System.Windows.Forms.TreeNode() {TreeNode1})
        Me.trvMenu.SelectedImageIndex = 0
        Me.trvMenu.ShowRootLines = False
        Me.trvMenu.Size = New System.Drawing.Size(210, 649)
        Me.trvMenu.TabIndex = 1
        '
        'Splitter1
        '
        Me.Splitter1.Location = New System.Drawing.Point(210, 24)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(3, 649)
        Me.Splitter1.TabIndex = 14
        Me.Splitter1.TabStop = False
        '
        'frmBG0010
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(1016, 695)
        Me.Controls.Add(Me.Splitter1)
        Me.Controls.Add(Me.trvMenu)
        Me.Controls.Add(Me.MenuStrip)
        Me.Controls.Add(Me.StatusStrip)
        Me.DoubleBuffered = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.MenuStrip
        Me.Name = "frmBG0010"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmBG0010"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip.ResumeLayout(False)
        Me.MenuStrip.PerformLayout()
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents HelpMenu As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CloseAllToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents WindowsMenu As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CascadeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TileVerticalToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TileHorizontalToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OptionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents lblToolbarStatus As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FileMenu As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MenuStrip As System.Windows.Forms.MenuStrip
    Friend WithEvents ToolsMenu As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents trvMenu As System.Windows.Forms.TreeView
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents lblUserName As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblPIC As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblProgramVersion As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblCurrMode As System.Windows.Forms.ToolStripStatusLabel

End Class
