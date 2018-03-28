<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBG0380
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBG0380))
        Me.lblFormTitle = New System.Windows.Forms.Label
        Me.cmdDelete = New System.Windows.Forms.Button
        Me.fraInfo = New System.Windows.Forms.GroupBox
        Me.cmdBrowse = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtFileTitle = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtFilePath = New System.Windows.Forms.TextBox
        Me.cmdAdd = New System.Windows.Forms.Button
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.grvMaster = New System.Windows.Forms.DataGridView
        Me.FILENODataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.FILETITLEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.FILEPATHDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CREATEUSERIDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CREATEDATEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.UPDATEUSERIDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.UPDATEDATEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataSet1 = New System.Data.DataSet
        Me.dtInformation = New System.Data.DataTable
        Me.DataColumn1 = New System.Data.DataColumn
        Me.DataColumn2 = New System.Data.DataColumn
        Me.DataColumn3 = New System.Data.DataColumn
        Me.DataColumn4 = New System.Data.DataColumn
        Me.DataColumn5 = New System.Data.DataColumn
        Me.DataColumn6 = New System.Data.DataColumn
        Me.DataColumn7 = New System.Data.DataColumn
        Me.fraInfo.SuspendLayout()
        CType(Me.grvMaster, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataSet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtInformation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblFormTitle
        '
        Me.lblFormTitle.AutoSize = True
        Me.lblFormTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblFormTitle.Location = New System.Drawing.Point(12, 9)
        Me.lblFormTitle.Name = "lblFormTitle"
        Me.lblFormTitle.Size = New System.Drawing.Size(240, 24)
        Me.lblFormTitle.TabIndex = 11
        Me.lblFormTitle.Text = "Add/Remove Information"
        '
        'cmdDelete
        '
        Me.cmdDelete.Location = New System.Drawing.Point(174, 532)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(75, 23)
        Me.cmdDelete.TabIndex = 5
        Me.cmdDelete.Text = "&Delete"
        Me.cmdDelete.UseVisualStyleBackColor = True
        '
        'fraInfo
        '
        Me.fraInfo.Controls.Add(Me.cmdBrowse)
        Me.fraInfo.Controls.Add(Me.Label1)
        Me.fraInfo.Controls.Add(Me.txtFileTitle)
        Me.fraInfo.Controls.Add(Me.Label2)
        Me.fraInfo.Controls.Add(Me.txtFilePath)
        Me.fraInfo.Location = New System.Drawing.Point(12, 439)
        Me.fraInfo.Name = "fraInfo"
        Me.fraInfo.Size = New System.Drawing.Size(768, 87)
        Me.fraInfo.TabIndex = 62
        Me.fraInfo.TabStop = False
        Me.fraInfo.Text = "File Information"
        '
        'cmdBrowse
        '
        Me.cmdBrowse.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdBrowse.Location = New System.Drawing.Point(535, 48)
        Me.cmdBrowse.Name = "cmdBrowse"
        Me.cmdBrowse.Size = New System.Drawing.Size(75, 23)
        Me.cmdBrowse.TabIndex = 2
        Me.cmdBrowse.Text = "&Browse..."
        Me.cmdBrowse.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(28, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 13)
        Me.Label1.TabIndex = 41
        Me.Label1.Text = "File Title:"
        '
        'txtFileTitle
        '
        Me.txtFileTitle.Location = New System.Drawing.Point(85, 23)
        Me.txtFileTitle.MaxLength = 100
        Me.txtFileTitle.Name = "txtFileTitle"
        Me.txtFileTitle.Size = New System.Drawing.Size(304, 20)
        Me.txtFileTitle.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(28, 52)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(51, 13)
        Me.Label2.TabIndex = 42
        Me.Label2.Text = "File Path:"
        '
        'txtFilePath
        '
        Me.txtFilePath.Location = New System.Drawing.Point(85, 49)
        Me.txtFilePath.MaxLength = 100
        Me.txtFilePath.Name = "txtFilePath"
        Me.txtFilePath.Size = New System.Drawing.Size(450, 20)
        Me.txtFilePath.TabIndex = 1
        '
        'cmdAdd
        '
        Me.cmdAdd.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdAdd.Location = New System.Drawing.Point(12, 532)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(75, 23)
        Me.cmdAdd.TabIndex = 3
        Me.cmdAdd.Text = "&Add New"
        Me.cmdAdd.UseVisualStyleBackColor = True
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(93, 532)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(75, 23)
        Me.cmdSave.TabIndex = 4
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'cmdClose
        '
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdClose.Location = New System.Drawing.Point(705, 532)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 23)
        Me.cmdClose.TabIndex = 6
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
        Me.grvMaster.AutoGenerateColumns = False
        Me.grvMaster.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.grvMaster.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.grvMaster.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        Me.grvMaster.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grvMaster.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.grvMaster.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.grvMaster.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.FILENODataGridViewTextBoxColumn, Me.FILETITLEDataGridViewTextBoxColumn, Me.FILEPATHDataGridViewTextBoxColumn, Me.CREATEUSERIDDataGridViewTextBoxColumn, Me.CREATEDATEDataGridViewTextBoxColumn, Me.UPDATEUSERIDDataGridViewTextBoxColumn, Me.UPDATEDATEDataGridViewTextBoxColumn})
        Me.grvMaster.DataMember = "dtInformation"
        Me.grvMaster.DataSource = Me.DataSet1
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle3.NullValue = "-"
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grvMaster.DefaultCellStyle = DataGridViewCellStyle3
        Me.grvMaster.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.grvMaster.Location = New System.Drawing.Point(12, 44)
        Me.grvMaster.MultiSelect = False
        Me.grvMaster.Name = "grvMaster"
        Me.grvMaster.ReadOnly = True
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
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grvMaster.RowsDefaultCellStyle = DataGridViewCellStyle5
        Me.grvMaster.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.grvMaster.Size = New System.Drawing.Size(768, 389)
        Me.grvMaster.StandardTab = True
        Me.grvMaster.TabIndex = 7
        '
        'FILENODataGridViewTextBoxColumn
        '
        Me.FILENODataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.FILENODataGridViewTextBoxColumn.DataPropertyName = "FILE_NO"
        Me.FILENODataGridViewTextBoxColumn.HeaderText = "FILE_NO"
        Me.FILENODataGridViewTextBoxColumn.MinimumWidth = 30
        Me.FILENODataGridViewTextBoxColumn.Name = "FILENODataGridViewTextBoxColumn"
        Me.FILENODataGridViewTextBoxColumn.ReadOnly = True
        Me.FILENODataGridViewTextBoxColumn.Visible = False
        Me.FILENODataGridViewTextBoxColumn.Width = 50
        '
        'FILETITLEDataGridViewTextBoxColumn
        '
        Me.FILETITLEDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.FILETITLEDataGridViewTextBoxColumn.DataPropertyName = "FILE_TITLE"
        Me.FILETITLEDataGridViewTextBoxColumn.HeaderText = "File Title"
        Me.FILETITLEDataGridViewTextBoxColumn.MinimumWidth = 100
        Me.FILETITLEDataGridViewTextBoxColumn.Name = "FILETITLEDataGridViewTextBoxColumn"
        Me.FILETITLEDataGridViewTextBoxColumn.ReadOnly = True
        Me.FILETITLEDataGridViewTextBoxColumn.Width = 200
        '
        'FILEPATHDataGridViewTextBoxColumn
        '
        Me.FILEPATHDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.FILEPATHDataGridViewTextBoxColumn.DataPropertyName = "FILE_PATH"
        Me.FILEPATHDataGridViewTextBoxColumn.HeaderText = "File Path"
        Me.FILEPATHDataGridViewTextBoxColumn.MinimumWidth = 100
        Me.FILEPATHDataGridViewTextBoxColumn.Name = "FILEPATHDataGridViewTextBoxColumn"
        Me.FILEPATHDataGridViewTextBoxColumn.ReadOnly = True
        Me.FILEPATHDataGridViewTextBoxColumn.Width = 564
        '
        'CREATEUSERIDDataGridViewTextBoxColumn
        '
        Me.CREATEUSERIDDataGridViewTextBoxColumn.DataPropertyName = "CREATE_USER_ID"
        Me.CREATEUSERIDDataGridViewTextBoxColumn.HeaderText = "CREATE_USER_ID"
        Me.CREATEUSERIDDataGridViewTextBoxColumn.Name = "CREATEUSERIDDataGridViewTextBoxColumn"
        Me.CREATEUSERIDDataGridViewTextBoxColumn.ReadOnly = True
        Me.CREATEUSERIDDataGridViewTextBoxColumn.Visible = False
        Me.CREATEUSERIDDataGridViewTextBoxColumn.Width = 128
        '
        'CREATEDATEDataGridViewTextBoxColumn
        '
        Me.CREATEDATEDataGridViewTextBoxColumn.DataPropertyName = "CREATE_DATE"
        Me.CREATEDATEDataGridViewTextBoxColumn.HeaderText = "CREATE_DATE"
        Me.CREATEDATEDataGridViewTextBoxColumn.Name = "CREATEDATEDataGridViewTextBoxColumn"
        Me.CREATEDATEDataGridViewTextBoxColumn.ReadOnly = True
        Me.CREATEDATEDataGridViewTextBoxColumn.Visible = False
        Me.CREATEDATEDataGridViewTextBoxColumn.Width = 110
        '
        'UPDATEUSERIDDataGridViewTextBoxColumn
        '
        Me.UPDATEUSERIDDataGridViewTextBoxColumn.DataPropertyName = "UPDATE_USER_ID"
        Me.UPDATEUSERIDDataGridViewTextBoxColumn.HeaderText = "UPDATE_USER_ID"
        Me.UPDATEUSERIDDataGridViewTextBoxColumn.Name = "UPDATEUSERIDDataGridViewTextBoxColumn"
        Me.UPDATEUSERIDDataGridViewTextBoxColumn.ReadOnly = True
        Me.UPDATEUSERIDDataGridViewTextBoxColumn.Visible = False
        Me.UPDATEUSERIDDataGridViewTextBoxColumn.Width = 129
        '
        'UPDATEDATEDataGridViewTextBoxColumn
        '
        Me.UPDATEDATEDataGridViewTextBoxColumn.DataPropertyName = "UPDATE_DATE"
        Me.UPDATEDATEDataGridViewTextBoxColumn.HeaderText = "UPDATE_DATE"
        Me.UPDATEDATEDataGridViewTextBoxColumn.Name = "UPDATEDATEDataGridViewTextBoxColumn"
        Me.UPDATEDATEDataGridViewTextBoxColumn.ReadOnly = True
        Me.UPDATEDATEDataGridViewTextBoxColumn.Visible = False
        Me.UPDATEDATEDataGridViewTextBoxColumn.Width = 111
        '
        'DataSet1
        '
        Me.DataSet1.DataSetName = "NewDataSet"
        Me.DataSet1.Tables.AddRange(New System.Data.DataTable() {Me.dtInformation})
        '
        'dtInformation
        '
        Me.dtInformation.Columns.AddRange(New System.Data.DataColumn() {Me.DataColumn1, Me.DataColumn2, Me.DataColumn3, Me.DataColumn4, Me.DataColumn5, Me.DataColumn6, Me.DataColumn7})
        Me.dtInformation.TableName = "dtInformation"
        '
        'DataColumn1
        '
        Me.DataColumn1.Caption = "File Number"
        Me.DataColumn1.ColumnName = "FILE_NO"
        '
        'DataColumn2
        '
        Me.DataColumn2.Caption = "File Title"
        Me.DataColumn2.ColumnName = "FILE_TITLE"
        '
        'DataColumn3
        '
        Me.DataColumn3.Caption = "Path"
        Me.DataColumn3.ColumnName = "FILE_PATH"
        '
        'DataColumn4
        '
        Me.DataColumn4.Caption = "Created by User"
        Me.DataColumn4.ColumnName = "CREATE_USER_ID"
        '
        'DataColumn5
        '
        Me.DataColumn5.Caption = "Created Date"
        Me.DataColumn5.ColumnName = "CREATE_DATE"
        '
        'DataColumn6
        '
        Me.DataColumn6.Caption = "Latest update by"
        Me.DataColumn6.ColumnName = "UPDATE_USER_ID"
        '
        'DataColumn7
        '
        Me.DataColumn7.Caption = "Latest update on"
        Me.DataColumn7.ColumnName = "UPDATE_DATE"
        '
        'frmBG0380
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(792, 567)
        Me.Controls.Add(Me.cmdDelete)
        Me.Controls.Add(Me.fraInfo)
        Me.Controls.Add(Me.cmdAdd)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.grvMaster)
        Me.Controls.Add(Me.lblFormTitle)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "frmBG0380"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "frmBG0380"
        Me.fraInfo.ResumeLayout(False)
        Me.fraInfo.PerformLayout()
        CType(Me.grvMaster, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataSet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtInformation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblFormTitle As System.Windows.Forms.Label
    Friend WithEvents cmdDelete As System.Windows.Forms.Button
    Friend WithEvents fraInfo As System.Windows.Forms.GroupBox
    Friend WithEvents cmdBrowse As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtFileTitle As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtFilePath As System.Windows.Forms.TextBox
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents grvMaster As System.Windows.Forms.DataGridView
    Friend WithEvents DataSet1 As System.Data.DataSet
    Friend WithEvents dtInformation As System.Data.DataTable
    Friend WithEvents DataColumn1 As System.Data.DataColumn
    Friend WithEvents DataColumn2 As System.Data.DataColumn
    Friend WithEvents DataColumn3 As System.Data.DataColumn
    Friend WithEvents DataColumn4 As System.Data.DataColumn
    Friend WithEvents DataColumn5 As System.Data.DataColumn
    Friend WithEvents DataColumn6 As System.Data.DataColumn
    Friend WithEvents DataColumn7 As System.Data.DataColumn
    Friend WithEvents FILENODataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FILETITLEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FILEPATHDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CREATEUSERIDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CREATEDATEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents UPDATEUSERIDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents UPDATEDATEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
