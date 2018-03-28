<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBG0660
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBG0660))
        Me.lblFormTitle = New System.Windows.Forms.Label
        Me.numYear = New System.Windows.Forms.NumericUpDown
        Me.cboPeriodType = New System.Windows.Forms.ComboBox
        Me.lblConfirmPwd = New System.Windows.Forms.Label
        Me.lblYear = New System.Windows.Forms.Label
        Me.cmdExport = New System.Windows.Forms.Button
        Me.cmdImport = New System.Windows.Forms.Button
        Me.cmdDelete = New System.Windows.Forms.Button
        Me.fraInfo = New System.Windows.Forms.GroupBox
        Me.optAddByAccountNo = New System.Windows.Forms.RadioButton
        Me.optAddByOrderNo = New System.Windows.Forms.RadioButton
        Me.cboAccountNo = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.cboToOrderNo = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.cboFromOrderNo = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.cboTransferType = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.cboBGOrderNo = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtTransferRate = New System.Windows.Forms.TextBox
        Me.cmdAdd = New System.Windows.Forms.Button
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.grvMaster = New System.Windows.Forms.DataGridView
        Me.BUDGETYEARDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PERIODTYPEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TRANSFERTYPETEXTDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BUDGETORDERNODataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BUDGETORDERNAMEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ACCOUNTNODataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.COSTTEXTDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.COSTDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DEPTNODataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TRANSFERTYPEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TRANSFERRATEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.FROMORDERNODataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.FROMCOSTTEXTDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.FROMCOSTDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TOORDERNODataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TOCOSTTEXTDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TOCOSTDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CREATE_USER_ID = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CREATE_DATE = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.UPDATE_USER_ID = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.UPDATE_DATE = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataSet1 = New System.Data.DataSet
        Me.dtTCS = New System.Data.DataTable
        Me.DataColumn1 = New System.Data.DataColumn
        Me.DataColumn2 = New System.Data.DataColumn
        Me.DataColumn3 = New System.Data.DataColumn
        Me.DataColumn4 = New System.Data.DataColumn
        Me.DataColumn5 = New System.Data.DataColumn
        Me.DataColumn6 = New System.Data.DataColumn
        Me.DataColumn7 = New System.Data.DataColumn
        Me.DataColumn8 = New System.Data.DataColumn
        Me.DataColumn9 = New System.Data.DataColumn
        Me.DataColumn10 = New System.Data.DataColumn
        Me.DataColumn11 = New System.Data.DataColumn
        Me.DataColumn12 = New System.Data.DataColumn
        Me.DataColumn13 = New System.Data.DataColumn
        Me.DataColumn14 = New System.Data.DataColumn
        Me.DataColumn15 = New System.Data.DataColumn
        Me.DataColumn16 = New System.Data.DataColumn
        Me.DataColumn17 = New System.Data.DataColumn
        Me.DataColumn18 = New System.Data.DataColumn
        Me.DataColumn19 = New System.Data.DataColumn
        Me.DataColumn20 = New System.Data.DataColumn
        Me.DataColumn21 = New System.Data.DataColumn
        Me.numProjectNo = New System.Windows.Forms.NumericUpDown
        Me.lblProjectNo = New System.Windows.Forms.Label
        Me.fraFilter = New System.Windows.Forms.GroupBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtAccountNameFilter = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.txtOrderNameFilter = New System.Windows.Forms.TextBox
        Me.txtToOrderNoFilter = New System.Windows.Forms.TextBox
        Me.txtFromOrderNoFilter = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.cboTransferTypeFilter = New System.Windows.Forms.ComboBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.txtOrderNoFilter = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtAccountNoFilter = New System.Windows.Forms.TextBox
        Me.cmdClearFilter = New System.Windows.Forms.Button
        Me.cmdFilter = New System.Windows.Forms.Button
        CType(Me.numYear, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraInfo.SuspendLayout()
        CType(Me.grvMaster, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataSet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtTCS, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numProjectNo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraFilter.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblFormTitle
        '
        Me.lblFormTitle.AutoSize = True
        Me.lblFormTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblFormTitle.Location = New System.Drawing.Point(12, 9)
        Me.lblFormTitle.Name = "lblFormTitle"
        Me.lblFormTitle.Size = New System.Drawing.Size(202, 24)
        Me.lblFormTitle.TabIndex = 0
        Me.lblFormTitle.Text = "Transfer Cost Master"
        '
        'numYear
        '
        Me.numYear.Location = New System.Drawing.Point(162, 14)
        Me.numYear.Maximum = New Decimal(New Integer() {3000, 0, 0, 0})
        Me.numYear.Minimum = New Decimal(New Integer() {2000, 0, 0, 0})
        Me.numYear.Name = "numYear"
        Me.numYear.Size = New System.Drawing.Size(49, 20)
        Me.numYear.TabIndex = 1
        Me.numYear.Value = New Decimal(New Integer() {2010, 0, 0, 0})
        '
        'cboPeriodType
        '
        Me.cboPeriodType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPeriodType.FormattingEnabled = True
        Me.cboPeriodType.Location = New System.Drawing.Point(162, 40)
        Me.cboPeriodType.Name = "cboPeriodType"
        Me.cboPeriodType.Size = New System.Drawing.Size(167, 21)
        Me.cboPeriodType.TabIndex = 3
        '
        'lblConfirmPwd
        '
        Me.lblConfirmPwd.AutoSize = True
        Me.lblConfirmPwd.Location = New System.Drawing.Point(66, 43)
        Me.lblConfirmPwd.Name = "lblConfirmPwd"
        Me.lblConfirmPwd.Size = New System.Drawing.Size(67, 13)
        Me.lblConfirmPwd.TabIndex = 2
        Me.lblConfirmPwd.Text = "Period &Type:"
        '
        'lblYear
        '
        Me.lblYear.AutoSize = True
        Me.lblYear.Location = New System.Drawing.Point(66, 16)
        Me.lblYear.Name = "lblYear"
        Me.lblYear.Size = New System.Drawing.Size(69, 13)
        Me.lblYear.TabIndex = 0
        Me.lblYear.Text = "Budget &Year:"
        '
        'cmdExport
        '
        Me.cmdExport.Location = New System.Drawing.Point(345, 663)
        Me.cmdExport.Name = "cmdExport"
        Me.cmdExport.Size = New System.Drawing.Size(75, 23)
        Me.cmdExport.TabIndex = 8
        Me.cmdExport.Text = "&Export"
        Me.cmdExport.UseVisualStyleBackColor = True
        '
        'cmdImport
        '
        Me.cmdImport.Location = New System.Drawing.Point(264, 663)
        Me.cmdImport.Name = "cmdImport"
        Me.cmdImport.Size = New System.Drawing.Size(75, 23)
        Me.cmdImport.TabIndex = 7
        Me.cmdImport.Text = "&Import"
        Me.cmdImport.UseVisualStyleBackColor = True
        '
        'cmdDelete
        '
        Me.cmdDelete.Location = New System.Drawing.Point(174, 663)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(75, 23)
        Me.cmdDelete.TabIndex = 6
        Me.cmdDelete.Text = "&Delete"
        Me.cmdDelete.UseVisualStyleBackColor = True
        '
        'fraInfo
        '
        Me.fraInfo.Controls.Add(Me.optAddByAccountNo)
        Me.fraInfo.Controls.Add(Me.optAddByOrderNo)
        Me.fraInfo.Controls.Add(Me.cboAccountNo)
        Me.fraInfo.Controls.Add(Me.Label6)
        Me.fraInfo.Controls.Add(Me.cboToOrderNo)
        Me.fraInfo.Controls.Add(Me.Label5)
        Me.fraInfo.Controls.Add(Me.cboFromOrderNo)
        Me.fraInfo.Controls.Add(Me.Label4)
        Me.fraInfo.Controls.Add(Me.cboTransferType)
        Me.fraInfo.Controls.Add(Me.Label3)
        Me.fraInfo.Controls.Add(Me.cboBGOrderNo)
        Me.fraInfo.Controls.Add(Me.Label1)
        Me.fraInfo.Controls.Add(Me.txtTransferRate)
        Me.fraInfo.Location = New System.Drawing.Point(12, 495)
        Me.fraInfo.Name = "fraInfo"
        Me.fraInfo.Size = New System.Drawing.Size(768, 162)
        Me.fraInfo.TabIndex = 3
        Me.fraInfo.TabStop = False
        Me.fraInfo.Text = "Transfer Cost Information"
        '
        'optAddByAccountNo
        '
        Me.optAddByAccountNo.AutoSize = True
        Me.optAddByAccountNo.Enabled = False
        Me.optAddByAccountNo.Location = New System.Drawing.Point(51, 47)
        Me.optAddByAccountNo.Name = "optAddByAccountNo"
        Me.optAddByAccountNo.Size = New System.Drawing.Size(88, 17)
        Me.optAddByAccountNo.TabIndex = 2
        Me.optAddByAccountNo.Text = "Account No.:"
        Me.optAddByAccountNo.UseVisualStyleBackColor = True
        '
        'optAddByOrderNo
        '
        Me.optAddByOrderNo.AutoSize = True
        Me.optAddByOrderNo.Checked = True
        Me.optAddByOrderNo.Location = New System.Drawing.Point(51, 20)
        Me.optAddByOrderNo.Name = "optAddByOrderNo"
        Me.optAddByOrderNo.Size = New System.Drawing.Size(114, 17)
        Me.optAddByOrderNo.TabIndex = 0
        Me.optAddByOrderNo.TabStop = True
        Me.optAddByOrderNo.Text = "Budget Order No. :"
        Me.optAddByOrderNo.UseVisualStyleBackColor = True
        '
        'cboAccountNo
        '
        Me.cboAccountNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAccountNo.Enabled = False
        Me.cboAccountNo.FormattingEnabled = True
        Me.cboAccountNo.Location = New System.Drawing.Point(166, 46)
        Me.cboAccountNo.Name = "cboAccountNo"
        Me.cboAccountNo.Size = New System.Drawing.Size(310, 21)
        Me.cboAccountNo.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(462, 76)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(15, 13)
        Me.Label6.TabIndex = 8
        Me.Label6.Text = "%"
        '
        'cboToOrderNo
        '
        Me.cboToOrderNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboToOrderNo.FormattingEnabled = True
        Me.cboToOrderNo.Location = New System.Drawing.Point(166, 127)
        Me.cboToOrderNo.Name = "cboToOrderNo"
        Me.cboToOrderNo.Size = New System.Drawing.Size(310, 21)
        Me.cboToOrderNo.TabIndex = 12
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(66, 130)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(75, 13)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "To Order No. :"
        '
        'cboFromOrderNo
        '
        Me.cboFromOrderNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFromOrderNo.FormattingEnabled = True
        Me.cboFromOrderNo.Location = New System.Drawing.Point(166, 100)
        Me.cboFromOrderNo.Name = "cboFromOrderNo"
        Me.cboFromOrderNo.Size = New System.Drawing.Size(310, 21)
        Me.cboFromOrderNo.TabIndex = 10
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(66, 103)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(85, 13)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "From Order No. :"
        '
        'cboTransferType
        '
        Me.cboTransferType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTransferType.FormattingEnabled = True
        Me.cboTransferType.Location = New System.Drawing.Point(166, 73)
        Me.cboTransferType.Name = "cboTransferType"
        Me.cboTransferType.Size = New System.Drawing.Size(128, 21)
        Me.cboTransferType.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(66, 76)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(79, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Transfer Type :"
        '
        'cboBGOrderNo
        '
        Me.cboBGOrderNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBGOrderNo.FormattingEnabled = True
        Me.cboBGOrderNo.Location = New System.Drawing.Point(166, 19)
        Me.cboBGOrderNo.Name = "cboBGOrderNo"
        Me.cboBGOrderNo.Size = New System.Drawing.Size(310, 21)
        Me.cboBGOrderNo.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(326, 76)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(78, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Transfer Rate :"
        '
        'txtTransferRate
        '
        Me.txtTransferRate.Location = New System.Drawing.Point(410, 73)
        Me.txtTransferRate.MaxLength = 100
        Me.txtTransferRate.Name = "txtTransferRate"
        Me.txtTransferRate.Size = New System.Drawing.Size(50, 20)
        Me.txtTransferRate.TabIndex = 7
        Me.txtTransferRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cmdAdd
        '
        Me.cmdAdd.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdAdd.Location = New System.Drawing.Point(12, 663)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(75, 23)
        Me.cmdAdd.TabIndex = 4
        Me.cmdAdd.Text = "&Add New"
        Me.cmdAdd.UseVisualStyleBackColor = True
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(93, 663)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(75, 23)
        Me.cmdSave.TabIndex = 5
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'cmdClose
        '
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdClose.Location = New System.Drawing.Point(705, 663)
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
        Me.grvMaster.AutoGenerateColumns = False
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
        Me.grvMaster.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.BUDGETYEARDataGridViewTextBoxColumn, Me.PERIODTYPEDataGridViewTextBoxColumn, Me.TRANSFERTYPETEXTDataGridViewTextBoxColumn, Me.BUDGETORDERNODataGridViewTextBoxColumn, Me.BUDGETORDERNAMEDataGridViewTextBoxColumn, Me.ACCOUNTNODataGridViewTextBoxColumn, Me.COSTTEXTDataGridViewTextBoxColumn, Me.COSTDataGridViewTextBoxColumn, Me.DEPTNODataGridViewTextBoxColumn, Me.TRANSFERTYPEDataGridViewTextBoxColumn, Me.TRANSFERRATEDataGridViewTextBoxColumn, Me.FROMORDERNODataGridViewTextBoxColumn, Me.FROMCOSTTEXTDataGridViewTextBoxColumn, Me.FROMCOSTDataGridViewTextBoxColumn, Me.TOORDERNODataGridViewTextBoxColumn, Me.TOCOSTTEXTDataGridViewTextBoxColumn, Me.TOCOSTDataGridViewTextBoxColumn, Me.CREATE_USER_ID, Me.CREATE_DATE, Me.UPDATE_USER_ID, Me.UPDATE_DATE})
        Me.grvMaster.DataMember = "dtTCS"
        Me.grvMaster.DataSource = Me.DataSet1
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.Azure
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle4.NullValue = "-"
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grvMaster.DefaultCellStyle = DataGridViewCellStyle4
        Me.grvMaster.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.grvMaster.EnableHeadersVisualStyles = False
        Me.grvMaster.Location = New System.Drawing.Point(12, 229)
        Me.grvMaster.MultiSelect = False
        Me.grvMaster.Name = "grvMaster"
        Me.grvMaster.ReadOnly = True
        Me.grvMaster.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grvMaster.RowHeadersDefaultCellStyle = DataGridViewCellStyle5
        Me.grvMaster.RowHeadersVisible = False
        Me.grvMaster.RowHeadersWidth = 30
        Me.grvMaster.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grvMaster.RowsDefaultCellStyle = DataGridViewCellStyle6
        Me.grvMaster.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.grvMaster.Size = New System.Drawing.Size(768, 260)
        Me.grvMaster.StandardTab = True
        Me.grvMaster.TabIndex = 2
        '
        'BUDGETYEARDataGridViewTextBoxColumn
        '
        Me.BUDGETYEARDataGridViewTextBoxColumn.DataPropertyName = "BUDGET_YEAR"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.BUDGETYEARDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle3
        Me.BUDGETYEARDataGridViewTextBoxColumn.HeaderText = "Year"
        Me.BUDGETYEARDataGridViewTextBoxColumn.Name = "BUDGETYEARDataGridViewTextBoxColumn"
        Me.BUDGETYEARDataGridViewTextBoxColumn.ReadOnly = True
        Me.BUDGETYEARDataGridViewTextBoxColumn.Visible = False
        Me.BUDGETYEARDataGridViewTextBoxColumn.Width = 54
        '
        'PERIODTYPEDataGridViewTextBoxColumn
        '
        Me.PERIODTYPEDataGridViewTextBoxColumn.DataPropertyName = "PERIOD_TYPE"
        Me.PERIODTYPEDataGridViewTextBoxColumn.HeaderText = "Period Type"
        Me.PERIODTYPEDataGridViewTextBoxColumn.Name = "PERIODTYPEDataGridViewTextBoxColumn"
        Me.PERIODTYPEDataGridViewTextBoxColumn.ReadOnly = True
        Me.PERIODTYPEDataGridViewTextBoxColumn.Visible = False
        Me.PERIODTYPEDataGridViewTextBoxColumn.Width = 89
        '
        'TRANSFERTYPETEXTDataGridViewTextBoxColumn
        '
        Me.TRANSFERTYPETEXTDataGridViewTextBoxColumn.DataPropertyName = "TRANSFER_TYPE_TEXT"
        Me.TRANSFERTYPETEXTDataGridViewTextBoxColumn.HeaderText = "Transfer Type"
        Me.TRANSFERTYPETEXTDataGridViewTextBoxColumn.Name = "TRANSFERTYPETEXTDataGridViewTextBoxColumn"
        Me.TRANSFERTYPETEXTDataGridViewTextBoxColumn.ReadOnly = True
        Me.TRANSFERTYPETEXTDataGridViewTextBoxColumn.Width = 80
        '
        'BUDGETORDERNODataGridViewTextBoxColumn
        '
        Me.BUDGETORDERNODataGridViewTextBoxColumn.DataPropertyName = "BUDGET_ORDER_NO"
        Me.BUDGETORDERNODataGridViewTextBoxColumn.HeaderText = "Order No."
        Me.BUDGETORDERNODataGridViewTextBoxColumn.Name = "BUDGETORDERNODataGridViewTextBoxColumn"
        Me.BUDGETORDERNODataGridViewTextBoxColumn.ReadOnly = True
        Me.BUDGETORDERNODataGridViewTextBoxColumn.Width = 80
        '
        'BUDGETORDERNAMEDataGridViewTextBoxColumn
        '
        Me.BUDGETORDERNAMEDataGridViewTextBoxColumn.DataPropertyName = "BUDGET_ORDER_NAME"
        Me.BUDGETORDERNAMEDataGridViewTextBoxColumn.HeaderText = "Order Name"
        Me.BUDGETORDERNAMEDataGridViewTextBoxColumn.Name = "BUDGETORDERNAMEDataGridViewTextBoxColumn"
        Me.BUDGETORDERNAMEDataGridViewTextBoxColumn.ReadOnly = True
        Me.BUDGETORDERNAMEDataGridViewTextBoxColumn.Width = 120
        '
        'ACCOUNTNODataGridViewTextBoxColumn
        '
        Me.ACCOUNTNODataGridViewTextBoxColumn.DataPropertyName = "ACCOUNT_NO"
        Me.ACCOUNTNODataGridViewTextBoxColumn.HeaderText = "Account No."
        Me.ACCOUNTNODataGridViewTextBoxColumn.Name = "ACCOUNTNODataGridViewTextBoxColumn"
        Me.ACCOUNTNODataGridViewTextBoxColumn.ReadOnly = True
        Me.ACCOUNTNODataGridViewTextBoxColumn.Width = 90
        '
        'COSTTEXTDataGridViewTextBoxColumn
        '
        Me.COSTTEXTDataGridViewTextBoxColumn.DataPropertyName = "COST_TEXT"
        Me.COSTTEXTDataGridViewTextBoxColumn.HeaderText = "Cost"
        Me.COSTTEXTDataGridViewTextBoxColumn.Name = "COSTTEXTDataGridViewTextBoxColumn"
        Me.COSTTEXTDataGridViewTextBoxColumn.ReadOnly = True
        Me.COSTTEXTDataGridViewTextBoxColumn.Width = 40
        '
        'COSTDataGridViewTextBoxColumn
        '
        Me.COSTDataGridViewTextBoxColumn.DataPropertyName = "COST"
        Me.COSTDataGridViewTextBoxColumn.HeaderText = "COST"
        Me.COSTDataGridViewTextBoxColumn.Name = "COSTDataGridViewTextBoxColumn"
        Me.COSTDataGridViewTextBoxColumn.ReadOnly = True
        Me.COSTDataGridViewTextBoxColumn.Visible = False
        Me.COSTDataGridViewTextBoxColumn.Width = 61
        '
        'DEPTNODataGridViewTextBoxColumn
        '
        Me.DEPTNODataGridViewTextBoxColumn.DataPropertyName = "DEPT_NO"
        Me.DEPTNODataGridViewTextBoxColumn.HeaderText = "Dept."
        Me.DEPTNODataGridViewTextBoxColumn.Name = "DEPTNODataGridViewTextBoxColumn"
        Me.DEPTNODataGridViewTextBoxColumn.ReadOnly = True
        Me.DEPTNODataGridViewTextBoxColumn.Width = 60
        '
        'TRANSFERTYPEDataGridViewTextBoxColumn
        '
        Me.TRANSFERTYPEDataGridViewTextBoxColumn.DataPropertyName = "TRANSFER_TYPE"
        Me.TRANSFERTYPEDataGridViewTextBoxColumn.HeaderText = "TRANSFER_TYPE"
        Me.TRANSFERTYPEDataGridViewTextBoxColumn.Name = "TRANSFERTYPEDataGridViewTextBoxColumn"
        Me.TRANSFERTYPEDataGridViewTextBoxColumn.ReadOnly = True
        Me.TRANSFERTYPEDataGridViewTextBoxColumn.Visible = False
        Me.TRANSFERTYPEDataGridViewTextBoxColumn.Width = 124
        '
        'TRANSFERRATEDataGridViewTextBoxColumn
        '
        Me.TRANSFERRATEDataGridViewTextBoxColumn.DataPropertyName = "TRANSFER_RATE"
        Me.TRANSFERRATEDataGridViewTextBoxColumn.HeaderText = "Transfer Rate (%)"
        Me.TRANSFERRATEDataGridViewTextBoxColumn.Name = "TRANSFERRATEDataGridViewTextBoxColumn"
        Me.TRANSFERRATEDataGridViewTextBoxColumn.ReadOnly = True
        '
        'FROMORDERNODataGridViewTextBoxColumn
        '
        Me.FROMORDERNODataGridViewTextBoxColumn.DataPropertyName = "FROM_ORDER_NO"
        Me.FROMORDERNODataGridViewTextBoxColumn.HeaderText = "From Order No."
        Me.FROMORDERNODataGridViewTextBoxColumn.Name = "FROMORDERNODataGridViewTextBoxColumn"
        Me.FROMORDERNODataGridViewTextBoxColumn.ReadOnly = True
        Me.FROMORDERNODataGridViewTextBoxColumn.Width = 90
        '
        'FROMCOSTTEXTDataGridViewTextBoxColumn
        '
        Me.FROMCOSTTEXTDataGridViewTextBoxColumn.DataPropertyName = "FROM_COST_TEXT"
        Me.FROMCOSTTEXTDataGridViewTextBoxColumn.HeaderText = "Cost"
        Me.FROMCOSTTEXTDataGridViewTextBoxColumn.Name = "FROMCOSTTEXTDataGridViewTextBoxColumn"
        Me.FROMCOSTTEXTDataGridViewTextBoxColumn.ReadOnly = True
        Me.FROMCOSTTEXTDataGridViewTextBoxColumn.Visible = False
        Me.FROMCOSTTEXTDataGridViewTextBoxColumn.Width = 60
        '
        'FROMCOSTDataGridViewTextBoxColumn
        '
        Me.FROMCOSTDataGridViewTextBoxColumn.DataPropertyName = "FROM_COST"
        Me.FROMCOSTDataGridViewTextBoxColumn.HeaderText = "FROM_COST"
        Me.FROMCOSTDataGridViewTextBoxColumn.Name = "FROMCOSTDataGridViewTextBoxColumn"
        Me.FROMCOSTDataGridViewTextBoxColumn.ReadOnly = True
        Me.FROMCOSTDataGridViewTextBoxColumn.Visible = False
        Me.FROMCOSTDataGridViewTextBoxColumn.Width = 98
        '
        'TOORDERNODataGridViewTextBoxColumn
        '
        Me.TOORDERNODataGridViewTextBoxColumn.DataPropertyName = "TO_ORDER_NO"
        Me.TOORDERNODataGridViewTextBoxColumn.HeaderText = "To Order No."
        Me.TOORDERNODataGridViewTextBoxColumn.Name = "TOORDERNODataGridViewTextBoxColumn"
        Me.TOORDERNODataGridViewTextBoxColumn.ReadOnly = True
        Me.TOORDERNODataGridViewTextBoxColumn.Width = 90
        '
        'TOCOSTTEXTDataGridViewTextBoxColumn
        '
        Me.TOCOSTTEXTDataGridViewTextBoxColumn.DataPropertyName = "TO_COST_TEXT"
        Me.TOCOSTTEXTDataGridViewTextBoxColumn.HeaderText = "Cost"
        Me.TOCOSTTEXTDataGridViewTextBoxColumn.Name = "TOCOSTTEXTDataGridViewTextBoxColumn"
        Me.TOCOSTTEXTDataGridViewTextBoxColumn.ReadOnly = True
        Me.TOCOSTTEXTDataGridViewTextBoxColumn.Visible = False
        Me.TOCOSTTEXTDataGridViewTextBoxColumn.Width = 60
        '
        'TOCOSTDataGridViewTextBoxColumn
        '
        Me.TOCOSTDataGridViewTextBoxColumn.DataPropertyName = "TO_COST"
        Me.TOCOSTDataGridViewTextBoxColumn.HeaderText = "TO_COST"
        Me.TOCOSTDataGridViewTextBoxColumn.Name = "TOCOSTDataGridViewTextBoxColumn"
        Me.TOCOSTDataGridViewTextBoxColumn.ReadOnly = True
        Me.TOCOSTDataGridViewTextBoxColumn.Visible = False
        Me.TOCOSTDataGridViewTextBoxColumn.Width = 82
        '
        'CREATE_USER_ID
        '
        Me.CREATE_USER_ID.DataPropertyName = "CREATE_USER_ID"
        Me.CREATE_USER_ID.HeaderText = "CREATE_USER_ID"
        Me.CREATE_USER_ID.Name = "CREATE_USER_ID"
        Me.CREATE_USER_ID.ReadOnly = True
        Me.CREATE_USER_ID.Visible = False
        '
        'CREATE_DATE
        '
        Me.CREATE_DATE.DataPropertyName = "CREATE_DATE"
        Me.CREATE_DATE.HeaderText = "CREATE_DATE"
        Me.CREATE_DATE.Name = "CREATE_DATE"
        Me.CREATE_DATE.ReadOnly = True
        Me.CREATE_DATE.Visible = False
        '
        'UPDATE_USER_ID
        '
        Me.UPDATE_USER_ID.DataPropertyName = "UPDATE_USER_ID"
        Me.UPDATE_USER_ID.HeaderText = "UPDATE_USER_ID"
        Me.UPDATE_USER_ID.Name = "UPDATE_USER_ID"
        Me.UPDATE_USER_ID.ReadOnly = True
        Me.UPDATE_USER_ID.Visible = False
        '
        'UPDATE_DATE
        '
        Me.UPDATE_DATE.DataPropertyName = "UPDATE_DATE"
        Me.UPDATE_DATE.HeaderText = "UPDATE_DATE"
        Me.UPDATE_DATE.Name = "UPDATE_DATE"
        Me.UPDATE_DATE.ReadOnly = True
        Me.UPDATE_DATE.Visible = False
        '
        'DataSet1
        '
        Me.DataSet1.DataSetName = "NewDataSet"
        Me.DataSet1.Tables.AddRange(New System.Data.DataTable() {Me.dtTCS})
        '
        'dtTCS
        '
        Me.dtTCS.Columns.AddRange(New System.Data.DataColumn() {Me.DataColumn1, Me.DataColumn2, Me.DataColumn3, Me.DataColumn4, Me.DataColumn5, Me.DataColumn6, Me.DataColumn7, Me.DataColumn8, Me.DataColumn9, Me.DataColumn10, Me.DataColumn11, Me.DataColumn12, Me.DataColumn13, Me.DataColumn14, Me.DataColumn15, Me.DataColumn16, Me.DataColumn17, Me.DataColumn18, Me.DataColumn19, Me.DataColumn20, Me.DataColumn21})
        Me.dtTCS.TableName = "dtTCS"
        '
        'DataColumn1
        '
        Me.DataColumn1.ColumnName = "BUDGET_YEAR"
        '
        'DataColumn2
        '
        Me.DataColumn2.ColumnName = "PERIOD_TYPE"
        '
        'DataColumn3
        '
        Me.DataColumn3.ColumnName = "BUDGET_ORDER_NO"
        '
        'DataColumn4
        '
        Me.DataColumn4.ColumnName = "BUDGET_ORDER_NAME"
        '
        'DataColumn5
        '
        Me.DataColumn5.ColumnName = "ACCOUNT_NO"
        '
        'DataColumn6
        '
        Me.DataColumn6.ColumnName = "COST_TEXT"
        '
        'DataColumn7
        '
        Me.DataColumn7.ColumnName = "COST"
        '
        'DataColumn8
        '
        Me.DataColumn8.ColumnName = "DEPT_NO"
        '
        'DataColumn9
        '
        Me.DataColumn9.ColumnName = "TRANSFER_TYPE_TEXT"
        '
        'DataColumn10
        '
        Me.DataColumn10.ColumnName = "TRANSFER_TYPE"
        '
        'DataColumn11
        '
        Me.DataColumn11.ColumnName = "TRANSFER_RATE"
        '
        'DataColumn12
        '
        Me.DataColumn12.ColumnName = "FROM_ORDER_NO"
        '
        'DataColumn13
        '
        Me.DataColumn13.ColumnName = "FROM_COST_TEXT"
        '
        'DataColumn14
        '
        Me.DataColumn14.ColumnName = "FROM_COST"
        '
        'DataColumn15
        '
        Me.DataColumn15.ColumnName = "TO_ORDER_NO"
        '
        'DataColumn16
        '
        Me.DataColumn16.ColumnName = "TO_COST_TEXT"
        '
        'DataColumn17
        '
        Me.DataColumn17.ColumnName = "TO_COST"
        '
        'DataColumn18
        '
        Me.DataColumn18.ColumnName = "CREATE_USER_ID"
        '
        'DataColumn19
        '
        Me.DataColumn19.ColumnName = "CREATE_DATE"
        '
        'DataColumn20
        '
        Me.DataColumn20.ColumnName = "UPDATE_USER_ID"
        '
        'DataColumn21
        '
        Me.DataColumn21.ColumnName = "UPDATE_DATE"
        '
        'numProjectNo
        '
        Me.numProjectNo.Location = New System.Drawing.Point(452, 41)
        Me.numProjectNo.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numProjectNo.Name = "numProjectNo"
        Me.numProjectNo.Size = New System.Drawing.Size(38, 20)
        Me.numProjectNo.TabIndex = 5
        Me.numProjectNo.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'lblProjectNo
        '
        Me.lblProjectNo.AutoSize = True
        Me.lblProjectNo.Location = New System.Drawing.Point(344, 43)
        Me.lblProjectNo.Name = "lblProjectNo"
        Me.lblProjectNo.Size = New System.Drawing.Size(60, 13)
        Me.lblProjectNo.TabIndex = 4
        Me.lblProjectNo.Text = "Project No:"
        '
        'fraFilter
        '
        Me.fraFilter.Controls.Add(Me.Label10)
        Me.fraFilter.Controls.Add(Me.txtAccountNameFilter)
        Me.fraFilter.Controls.Add(Me.Label22)
        Me.fraFilter.Controls.Add(Me.txtOrderNameFilter)
        Me.fraFilter.Controls.Add(Me.txtToOrderNoFilter)
        Me.fraFilter.Controls.Add(Me.txtFromOrderNoFilter)
        Me.fraFilter.Controls.Add(Me.Label8)
        Me.fraFilter.Controls.Add(Me.Label9)
        Me.fraFilter.Controls.Add(Me.cboTransferTypeFilter)
        Me.fraFilter.Controls.Add(Me.Label7)
        Me.fraFilter.Controls.Add(Me.Label21)
        Me.fraFilter.Controls.Add(Me.txtOrderNoFilter)
        Me.fraFilter.Controls.Add(Me.Label2)
        Me.fraFilter.Controls.Add(Me.numProjectNo)
        Me.fraFilter.Controls.Add(Me.txtAccountNoFilter)
        Me.fraFilter.Controls.Add(Me.lblProjectNo)
        Me.fraFilter.Controls.Add(Me.cmdClearFilter)
        Me.fraFilter.Controls.Add(Me.cmdFilter)
        Me.fraFilter.Controls.Add(Me.lblYear)
        Me.fraFilter.Controls.Add(Me.lblConfirmPwd)
        Me.fraFilter.Controls.Add(Me.cboPeriodType)
        Me.fraFilter.Controls.Add(Me.numYear)
        Me.fraFilter.Location = New System.Drawing.Point(12, 36)
        Me.fraFilter.Name = "fraFilter"
        Me.fraFilter.Size = New System.Drawing.Size(768, 187)
        Me.fraFilter.TabIndex = 1
        Me.fraFilter.TabStop = False
        Me.fraFilter.Text = "Filter Section"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(344, 110)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(81, 13)
        Me.Label10.TabIndex = 12
        Me.Label10.Text = "Account Name:"
        '
        'txtAccountNameFilter
        '
        Me.txtAccountNameFilter.Location = New System.Drawing.Point(452, 107)
        Me.txtAccountNameFilter.MaxLength = 100
        Me.txtAccountNameFilter.Name = "txtAccountNameFilter"
        Me.txtAccountNameFilter.Size = New System.Drawing.Size(240, 20)
        Me.txtAccountNameFilter.TabIndex = 13
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(342, 84)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(104, 13)
        Me.Label22.TabIndex = 8
        Me.Label22.Text = "Budget Order Name:"
        '
        'txtOrderNameFilter
        '
        Me.txtOrderNameFilter.Location = New System.Drawing.Point(452, 81)
        Me.txtOrderNameFilter.MaxLength = 100
        Me.txtOrderNameFilter.Name = "txtOrderNameFilter"
        Me.txtOrderNameFilter.Size = New System.Drawing.Size(240, 20)
        Me.txtOrderNameFilter.TabIndex = 9
        '
        'txtToOrderNoFilter
        '
        Me.txtToOrderNoFilter.Location = New System.Drawing.Point(452, 133)
        Me.txtToOrderNoFilter.MaxLength = 13
        Me.txtToOrderNoFilter.Name = "txtToOrderNoFilter"
        Me.txtToOrderNoFilter.Size = New System.Drawing.Size(80, 20)
        Me.txtToOrderNoFilter.TabIndex = 17
        '
        'txtFromOrderNoFilter
        '
        Me.txtFromOrderNoFilter.Location = New System.Drawing.Point(162, 133)
        Me.txtFromOrderNoFilter.MaxLength = 13
        Me.txtFromOrderNoFilter.Name = "txtFromOrderNoFilter"
        Me.txtFromOrderNoFilter.Size = New System.Drawing.Size(80, 20)
        Me.txtFromOrderNoFilter.TabIndex = 15
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(344, 136)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(75, 13)
        Me.Label8.TabIndex = 16
        Me.Label8.Text = "To Order No. :"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(66, 136)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(85, 13)
        Me.Label9.TabIndex = 14
        Me.Label9.Text = "From Order No. :"
        '
        'cboTransferTypeFilter
        '
        Me.cboTransferTypeFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTransferTypeFilter.FormattingEnabled = True
        Me.cboTransferTypeFilter.Location = New System.Drawing.Point(162, 159)
        Me.cboTransferTypeFilter.Name = "cboTransferTypeFilter"
        Me.cboTransferTypeFilter.Size = New System.Drawing.Size(128, 21)
        Me.cboTransferTypeFilter.TabIndex = 19
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(66, 162)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(79, 13)
        Me.Label7.TabIndex = 18
        Me.Label7.Text = "Transfer Type :"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(66, 84)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(90, 13)
        Me.Label21.TabIndex = 6
        Me.Label21.Text = "Budget Order No:"
        '
        'txtOrderNoFilter
        '
        Me.txtOrderNoFilter.Location = New System.Drawing.Point(162, 81)
        Me.txtOrderNoFilter.MaxLength = 13
        Me.txtOrderNoFilter.Name = "txtOrderNoFilter"
        Me.txtOrderNoFilter.Size = New System.Drawing.Size(80, 20)
        Me.txtOrderNoFilter.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(66, 110)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(67, 13)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Account No:"
        '
        'txtAccountNoFilter
        '
        Me.txtAccountNoFilter.Location = New System.Drawing.Point(162, 107)
        Me.txtAccountNoFilter.MaxLength = 6
        Me.txtAccountNoFilter.Name = "txtAccountNoFilter"
        Me.txtAccountNoFilter.Size = New System.Drawing.Size(80, 20)
        Me.txtAccountNoFilter.TabIndex = 11
        '
        'cmdClearFilter
        '
        Me.cmdClearFilter.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdClearFilter.Location = New System.Drawing.Point(687, 157)
        Me.cmdClearFilter.Name = "cmdClearFilter"
        Me.cmdClearFilter.Size = New System.Drawing.Size(75, 23)
        Me.cmdClearFilter.TabIndex = 21
        Me.cmdClearFilter.Text = "Clear"
        Me.cmdClearFilter.UseVisualStyleBackColor = True
        '
        'cmdFilter
        '
        Me.cmdFilter.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdFilter.Location = New System.Drawing.Point(606, 157)
        Me.cmdFilter.Name = "cmdFilter"
        Me.cmdFilter.Size = New System.Drawing.Size(75, 23)
        Me.cmdFilter.TabIndex = 20
        Me.cmdFilter.Text = "&Filter"
        Me.cmdFilter.UseVisualStyleBackColor = True
        '
        'frmBG0660
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
        Me.Name = "frmBG0660"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "frmBG0660"
        CType(Me.numYear, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraInfo.ResumeLayout(False)
        Me.fraInfo.PerformLayout()
        CType(Me.grvMaster, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataSet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtTCS, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numProjectNo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraFilter.ResumeLayout(False)
        Me.fraFilter.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblFormTitle As System.Windows.Forms.Label
    Friend WithEvents numYear As System.Windows.Forms.NumericUpDown
    Friend WithEvents cboPeriodType As System.Windows.Forms.ComboBox
    Friend WithEvents lblConfirmPwd As System.Windows.Forms.Label
    Friend WithEvents lblYear As System.Windows.Forms.Label
    Friend WithEvents cmdExport As System.Windows.Forms.Button
    Friend WithEvents cmdImport As System.Windows.Forms.Button
    Friend WithEvents cmdDelete As System.Windows.Forms.Button
    Friend WithEvents fraInfo As System.Windows.Forms.GroupBox
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents grvMaster As System.Windows.Forms.DataGridView
    Friend WithEvents DataSet1 As System.Data.DataSet
    Friend WithEvents dtTCS As System.Data.DataTable
    Friend WithEvents DataColumn1 As System.Data.DataColumn
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtTransferRate As System.Windows.Forms.TextBox
    Friend WithEvents cboToOrderNo As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboFromOrderNo As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboTransferType As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboBGOrderNo As System.Windows.Forms.ComboBox
    Friend WithEvents DataColumn2 As System.Data.DataColumn
    Friend WithEvents DataColumn3 As System.Data.DataColumn
    Friend WithEvents DataColumn4 As System.Data.DataColumn
    Friend WithEvents DataColumn5 As System.Data.DataColumn
    Friend WithEvents DataColumn6 As System.Data.DataColumn
    Friend WithEvents DataColumn7 As System.Data.DataColumn
    Friend WithEvents DataColumn8 As System.Data.DataColumn
    Friend WithEvents DataColumn9 As System.Data.DataColumn
    Friend WithEvents DataColumn10 As System.Data.DataColumn
    Friend WithEvents DataColumn11 As System.Data.DataColumn
    Friend WithEvents DataColumn12 As System.Data.DataColumn
    Friend WithEvents DataColumn13 As System.Data.DataColumn
    Friend WithEvents DataColumn14 As System.Data.DataColumn
    Friend WithEvents DataColumn15 As System.Data.DataColumn
    Friend WithEvents DataColumn16 As System.Data.DataColumn
    Friend WithEvents DataColumn17 As System.Data.DataColumn
    Friend WithEvents DataColumn18 As System.Data.DataColumn
    Friend WithEvents DataColumn19 As System.Data.DataColumn
    Friend WithEvents DataColumn20 As System.Data.DataColumn
    Friend WithEvents DataColumn21 As System.Data.DataColumn
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents BUDGETYEARDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PERIODTYPEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TRANSFERTYPETEXTDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BUDGETORDERNODataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BUDGETORDERNAMEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ACCOUNTNODataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents COSTTEXTDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents COSTDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DEPTNODataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TRANSFERTYPEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TRANSFERRATEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FROMORDERNODataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FROMCOSTTEXTDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FROMCOSTDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TOORDERNODataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TOCOSTTEXTDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TOCOSTDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CREATE_USER_ID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CREATE_DATE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents UPDATE_USER_ID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents UPDATE_DATE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents optAddByAccountNo As System.Windows.Forms.RadioButton
    Friend WithEvents optAddByOrderNo As System.Windows.Forms.RadioButton
    Friend WithEvents cboAccountNo As System.Windows.Forms.ComboBox
    Friend WithEvents numProjectNo As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblProjectNo As System.Windows.Forms.Label
    Friend WithEvents fraFilter As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtAccountNoFilter As System.Windows.Forms.TextBox
    Friend WithEvents cmdClearFilter As System.Windows.Forms.Button
    Friend WithEvents cmdFilter As System.Windows.Forms.Button
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents txtOrderNoFilter As System.Windows.Forms.TextBox
    Friend WithEvents cboTransferTypeFilter As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtToOrderNoFilter As System.Windows.Forms.TextBox
    Friend WithEvents txtFromOrderNoFilter As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents txtOrderNameFilter As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtAccountNameFilter As System.Windows.Forms.TextBox
End Class
