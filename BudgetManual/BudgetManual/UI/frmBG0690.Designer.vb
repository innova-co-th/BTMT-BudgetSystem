<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBG0690
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBG0690))
        Me.lblFormTitle = New System.Windows.Forms.Label
        Me.cmdExport = New System.Windows.Forms.Button
        Me.cmdImport = New System.Windows.Forms.Button
        Me.cmdDelete = New System.Windows.Forms.Button
        Me.fraInfo = New System.Windows.Forms.GroupBox
        Me.cboParentPIC = New System.Windows.Forms.ComboBox
        Me.cboChildPIC = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.cmdAdd = New System.Windows.Forms.Button
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.grvMaster = New System.Windows.Forms.DataGridView
        Me.fraFilter = New System.Windows.Forms.GroupBox
        Me.cmdClearFilter = New System.Windows.Forms.Button
        Me.cmdFilter = New System.Windows.Forms.Button
        Me.cboParentPICFilter = New System.Windows.Forms.ComboBox
        Me.cboChildPICFilter = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.fraInfo.SuspendLayout()
        CType(Me.grvMaster, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraFilter.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblFormTitle
        '
        Me.lblFormTitle.AutoSize = True
        Me.lblFormTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblFormTitle.Location = New System.Drawing.Point(12, 9)
        Me.lblFormTitle.Name = "lblFormTitle"
        Me.lblFormTitle.Size = New System.Drawing.Size(295, 24)
        Me.lblFormTitle.TabIndex = 0
        Me.lblFormTitle.Text = "Child Person In Charge Master"
        '
        'cmdExport
        '
        Me.cmdExport.Location = New System.Drawing.Point(345, 606)
        Me.cmdExport.Name = "cmdExport"
        Me.cmdExport.Size = New System.Drawing.Size(75, 23)
        Me.cmdExport.TabIndex = 8
        Me.cmdExport.Text = "&Export"
        Me.cmdExport.UseVisualStyleBackColor = True
        '
        'cmdImport
        '
        Me.cmdImport.Location = New System.Drawing.Point(264, 606)
        Me.cmdImport.Name = "cmdImport"
        Me.cmdImport.Size = New System.Drawing.Size(75, 23)
        Me.cmdImport.TabIndex = 7
        Me.cmdImport.Text = "&Import"
        Me.cmdImport.UseVisualStyleBackColor = True
        '
        'cmdDelete
        '
        Me.cmdDelete.Location = New System.Drawing.Point(174, 606)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(75, 23)
        Me.cmdDelete.TabIndex = 6
        Me.cmdDelete.Text = "&Delete"
        Me.cmdDelete.UseVisualStyleBackColor = True
        '
        'fraInfo
        '
        Me.fraInfo.Controls.Add(Me.cboParentPIC)
        Me.fraInfo.Controls.Add(Me.cboChildPIC)
        Me.fraInfo.Controls.Add(Me.Label1)
        Me.fraInfo.Controls.Add(Me.Label2)
        Me.fraInfo.Location = New System.Drawing.Point(12, 512)
        Me.fraInfo.Name = "fraInfo"
        Me.fraInfo.Size = New System.Drawing.Size(768, 88)
        Me.fraInfo.TabIndex = 3
        Me.fraInfo.TabStop = False
        Me.fraInfo.Text = "Person In Charge Information"
        '
        'cboParentPIC
        '
        Me.cboParentPIC.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboParentPIC.FormattingEnabled = True
        Me.cboParentPIC.Location = New System.Drawing.Point(160, 22)
        Me.cboParentPIC.Name = "cboParentPIC"
        Me.cboParentPIC.Size = New System.Drawing.Size(240, 21)
        Me.cboParentPIC.TabIndex = 1
        '
        'cboChildPIC
        '
        Me.cboChildPIC.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboChildPIC.FormattingEnabled = True
        Me.cboChildPIC.Location = New System.Drawing.Point(160, 49)
        Me.cboChildPIC.Name = "cboChildPIC"
        Me.cboChildPIC.Size = New System.Drawing.Size(240, 21)
        Me.cboChildPIC.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(28, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(126, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Parent Person In Charge:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(28, 52)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(118, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Child Person In Charge:"
        '
        'cmdAdd
        '
        Me.cmdAdd.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdAdd.Location = New System.Drawing.Point(12, 606)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(75, 23)
        Me.cmdAdd.TabIndex = 4
        Me.cmdAdd.Text = "&Add New"
        Me.cmdAdd.UseVisualStyleBackColor = True
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(93, 606)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(75, 23)
        Me.cmdSave.TabIndex = 5
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'cmdClose
        '
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdClose.Location = New System.Drawing.Point(705, 606)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 23)
        Me.cmdClose.TabIndex = 9
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'grvMaster
        '
        Me.grvMaster.AllowUserToAddRows = False
        Me.grvMaster.AllowUserToDeleteRows = False
        Me.grvMaster.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.PowderBlue
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grvMaster.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.grvMaster.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.grvMaster.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.grvMaster.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.SkyBlue
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grvMaster.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.grvMaster.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.Azure
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle3.NullValue = "-"
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grvMaster.DefaultCellStyle = DataGridViewCellStyle3
        Me.grvMaster.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.grvMaster.EnableHeadersVisualStyles = False
        Me.grvMaster.Location = New System.Drawing.Point(12, 129)
        Me.grvMaster.MultiSelect = False
        Me.grvMaster.Name = "grvMaster"
        Me.grvMaster.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grvMaster.RowHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.grvMaster.RowHeadersVisible = False
        Me.grvMaster.RowHeadersWidth = 30
        Me.grvMaster.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grvMaster.RowsDefaultCellStyle = DataGridViewCellStyle5
        Me.grvMaster.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.grvMaster.Size = New System.Drawing.Size(768, 377)
        Me.grvMaster.StandardTab = True
        Me.grvMaster.TabIndex = 2
        '
        'fraFilter
        '
        Me.fraFilter.Controls.Add(Me.cboParentPICFilter)
        Me.fraFilter.Controls.Add(Me.cboChildPICFilter)
        Me.fraFilter.Controls.Add(Me.Label3)
        Me.fraFilter.Controls.Add(Me.Label4)
        Me.fraFilter.Controls.Add(Me.cmdClearFilter)
        Me.fraFilter.Controls.Add(Me.cmdFilter)
        Me.fraFilter.Location = New System.Drawing.Point(12, 36)
        Me.fraFilter.Name = "fraFilter"
        Me.fraFilter.Size = New System.Drawing.Size(768, 87)
        Me.fraFilter.TabIndex = 1
        Me.fraFilter.TabStop = False
        Me.fraFilter.Text = "Filter Section"
        '
        'cmdClearFilter
        '
        Me.cmdClearFilter.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdClearFilter.Location = New System.Drawing.Point(687, 49)
        Me.cmdClearFilter.Name = "cmdClearFilter"
        Me.cmdClearFilter.Size = New System.Drawing.Size(75, 23)
        Me.cmdClearFilter.TabIndex = 5
        Me.cmdClearFilter.Text = "Clear"
        Me.cmdClearFilter.UseVisualStyleBackColor = True
        '
        'cmdFilter
        '
        Me.cmdFilter.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdFilter.Location = New System.Drawing.Point(606, 49)
        Me.cmdFilter.Name = "cmdFilter"
        Me.cmdFilter.Size = New System.Drawing.Size(75, 23)
        Me.cmdFilter.TabIndex = 4
        Me.cmdFilter.Text = "&Filter"
        Me.cmdFilter.UseVisualStyleBackColor = True
        '
        'cboParentPICFilter
        '
        Me.cboParentPICFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboParentPICFilter.FormattingEnabled = True
        Me.cboParentPICFilter.Location = New System.Drawing.Point(160, 24)
        Me.cboParentPICFilter.Name = "cboParentPICFilter"
        Me.cboParentPICFilter.Size = New System.Drawing.Size(240, 21)
        Me.cboParentPICFilter.TabIndex = 1
        '
        'cboChildPICFilter
        '
        Me.cboChildPICFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboChildPICFilter.FormattingEnabled = True
        Me.cboChildPICFilter.Location = New System.Drawing.Point(160, 51)
        Me.cboChildPICFilter.Name = "cboChildPICFilter"
        Me.cboChildPICFilter.Size = New System.Drawing.Size(240, 21)
        Me.cboChildPICFilter.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(28, 27)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(126, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Parent Person In Charge:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(28, 54)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(118, 13)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Child Person In Charge:"
        '
        'frmBG0690
        '
        Me.AcceptButton = Me.cmdFilter
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(792, 578)
        Me.Controls.Add(Me.fraFilter)
        Me.Controls.Add(Me.cmdExport)
        Me.Controls.Add(Me.cmdImport)
        Me.Controls.Add(Me.cmdDelete)
        Me.Controls.Add(Me.fraInfo)
        Me.Controls.Add(Me.cmdAdd)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.grvMaster)
        Me.Controls.Add(Me.lblFormTitle)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "frmBG0690"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "frmBG0690"
        Me.fraInfo.ResumeLayout(False)
        Me.fraInfo.PerformLayout()
        CType(Me.grvMaster, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraFilter.ResumeLayout(False)
        Me.fraFilter.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblFormTitle As System.Windows.Forms.Label
    Friend WithEvents cmdExport As System.Windows.Forms.Button
    Friend WithEvents cmdImport As System.Windows.Forms.Button
    Friend WithEvents cmdDelete As System.Windows.Forms.Button
    Friend WithEvents fraInfo As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents grvMaster As System.Windows.Forms.DataGridView
    Friend WithEvents cboChildPIC As System.Windows.Forms.ComboBox
    Friend WithEvents cboParentPIC As System.Windows.Forms.ComboBox
    Friend WithEvents fraFilter As System.Windows.Forms.GroupBox
    Friend WithEvents cboParentPICFilter As System.Windows.Forms.ComboBox
    Friend WithEvents cboChildPICFilter As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmdClearFilter As System.Windows.Forms.Button
    Friend WithEvents cmdFilter As System.Windows.Forms.Button
End Class
