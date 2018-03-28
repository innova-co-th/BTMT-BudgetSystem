<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBG0390
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
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle14 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBG0390))
        Me.lblFormTitle = New System.Windows.Forms.Label
        Me.grvLog = New System.Windows.Forms.DataGridView
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column13 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column12 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column5 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column9 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PROJECT_NO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column6 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column10 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column7 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column11 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column8 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column14 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.cmdClose = New System.Windows.Forms.Button
        Me.cmdSearch = New System.Windows.Forms.Button
        Me.numYear = New System.Windows.Forms.NumericUpDown
        Me.cboPeriodType = New System.Windows.Forms.ComboBox
        Me.lblConfirmPwd = New System.Windows.Forms.Label
        Me.lblYear = New System.Windows.Forms.Label
        Me.optBudget = New System.Windows.Forms.RadioButton
        Me.optAccount = New System.Windows.Forms.RadioButton
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.numProjectNo = New System.Windows.Forms.NumericUpDown
        Me.lblProjectNo = New System.Windows.Forms.Label
        CType(Me.grvLog, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numYear, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numProjectNo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblFormTitle
        '
        Me.lblFormTitle.AutoSize = True
        Me.lblFormTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblFormTitle.Location = New System.Drawing.Point(12, 9)
        Me.lblFormTitle.Name = "lblFormTitle"
        Me.lblFormTitle.Size = New System.Drawing.Size(212, 24)
        Me.lblFormTitle.TabIndex = 11
        Me.lblFormTitle.Text = "View Transaction Log"
        '
        'grvLog
        '
        Me.grvLog.AllowUserToAddRows = False
        Me.grvLog.AllowUserToDeleteRows = False
        Me.grvLog.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.PowderBlue
        DataGridViewCellStyle1.NullValue = "-"
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grvLog.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.grvLog.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grvLog.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.grvLog.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.grvLog.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.SkyBlue
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grvLog.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.grvLog.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.grvLog.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column13, Me.Column2, Me.Column3, Me.Column12, Me.Column4, Me.Column5, Me.Column9, Me.PROJECT_NO, Me.Column6, Me.Column10, Me.Column7, Me.Column11, Me.Column8, Me.Column14})
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle12.BackColor = System.Drawing.Color.Azure
        DataGridViewCellStyle12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle12.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle12.NullValue = "-"
        DataGridViewCellStyle12.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle12.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle12.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grvLog.DefaultCellStyle = DataGridViewCellStyle12
        Me.grvLog.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.grvLog.EnableHeadersVisualStyles = False
        Me.grvLog.Location = New System.Drawing.Point(12, 144)
        Me.grvLog.MultiSelect = False
        Me.grvLog.Name = "grvLog"
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle13.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle13.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle13.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle13.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle13.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grvLog.RowHeadersDefaultCellStyle = DataGridViewCellStyle13
        Me.grvLog.RowHeadersVisible = False
        Me.grvLog.RowHeadersWidth = 30
        Me.grvLog.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridViewCellStyle14.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grvLog.RowsDefaultCellStyle = DataGridViewCellStyle14
        Me.grvLog.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.grvLog.Size = New System.Drawing.Size(815, 382)
        Me.grvLog.StandardTab = True
        Me.grvLog.TabIndex = 5
        '
        'Column1
        '
        Me.Column1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Column1.DataPropertyName = "LOG_DATE"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        Me.Column1.DefaultCellStyle = DataGridViewCellStyle3
        Me.Column1.HeaderText = "LOG_DATE"
        Me.Column1.Name = "Column1"
        Me.Column1.Visible = False
        '
        'Column13
        '
        Me.Column13.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Column13.DataPropertyName = "LOG_DATE2"
        Me.Column13.HeaderText = "Log Date"
        Me.Column13.Name = "Column13"
        Me.Column13.Width = 76
        '
        'Column2
        '
        Me.Column2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Column2.DataPropertyName = "USER_ID"
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        Me.Column2.DefaultCellStyle = DataGridViewCellStyle4
        Me.Column2.HeaderText = "User ID"
        Me.Column2.Name = "Column2"
        Me.Column2.Width = 68
        '
        'Column3
        '
        Me.Column3.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Column3.DataPropertyName = "OPERATION_CD"
        Me.Column3.HeaderText = "OPERATION_CD"
        Me.Column3.Name = "Column3"
        Me.Column3.Visible = False
        '
        'Column12
        '
        Me.Column12.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Column12.DataPropertyName = "OPERATION_NAME"
        Me.Column12.HeaderText = "Operation Name"
        Me.Column12.Name = "Column12"
        Me.Column12.Width = 109
        '
        'Column4
        '
        Me.Column4.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Column4.DataPropertyName = "BUDGET_YEAR"
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Column4.DefaultCellStyle = DataGridViewCellStyle5
        Me.Column4.HeaderText = "Budget Year"
        Me.Column4.Name = "Column4"
        Me.Column4.Width = 91
        '
        'Column5
        '
        Me.Column5.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Column5.DataPropertyName = "PERIOD_TYPE"
        Me.Column5.HeaderText = "PERIOD_TYPE"
        Me.Column5.Name = "Column5"
        Me.Column5.Visible = False
        '
        'Column9
        '
        Me.Column9.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Column9.DataPropertyName = "PERIOD_TYPE_NAME"
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Column9.DefaultCellStyle = DataGridViewCellStyle6
        Me.Column9.HeaderText = "Period Type"
        Me.Column9.Name = "Column9"
        Me.Column9.Width = 89
        '
        'PROJECT_NO
        '
        Me.PROJECT_NO.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.PROJECT_NO.DataPropertyName = "PROJECT_NO"
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.PROJECT_NO.DefaultCellStyle = DataGridViewCellStyle7
        Me.PROJECT_NO.HeaderText = "Project No."
        Me.PROJECT_NO.Name = "PROJECT_NO"
        Me.PROJECT_NO.Width = 85
        '
        'Column6
        '
        Me.Column6.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Column6.DataPropertyName = "PERSON_IN_CHARGE_NO"
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Column6.DefaultCellStyle = DataGridViewCellStyle8
        Me.Column6.HeaderText = "PERSON_IN_CHARGE_NO"
        Me.Column6.Name = "Column6"
        Me.Column6.Visible = False
        '
        'Column10
        '
        Me.Column10.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Column10.DataPropertyName = "PIC_NAME"
        Me.Column10.HeaderText = "Person In Charge"
        Me.Column10.Name = "Column10"
        Me.Column10.Width = 114
        '
        'Column7
        '
        Me.Column7.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Column7.DataPropertyName = "BUDGET_TYPE"
        Me.Column7.HeaderText = "BUDGET_TYPE"
        Me.Column7.Name = "Column7"
        Me.Column7.Visible = False
        '
        'Column11
        '
        Me.Column11.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Column11.DataPropertyName = "BUDGET_TYPE_NAME"
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Column11.DefaultCellStyle = DataGridViewCellStyle9
        Me.Column11.HeaderText = "Budget Type"
        Me.Column11.Name = "Column11"
        Me.Column11.Width = 93
        '
        'Column8
        '
        Me.Column8.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Column8.DataPropertyName = "REV_NO"
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Column8.DefaultCellStyle = DataGridViewCellStyle10
        Me.Column8.HeaderText = "Rev No"
        Me.Column8.Name = "Column8"
        Me.Column8.Visible = False
        '
        'Column14
        '
        Me.Column14.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Column14.DataPropertyName = "REV_NO2"
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Column14.DefaultCellStyle = DataGridViewCellStyle11
        Me.Column14.HeaderText = "Rev No."
        Me.Column14.Name = "Column14"
        Me.Column14.Width = 72
        '
        'cmdClose
        '
        Me.cmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdClose.Location = New System.Drawing.Point(752, 532)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 23)
        Me.cmdClose.TabIndex = 6
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'cmdSearch
        '
        Me.cmdSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSearch.Location = New System.Drawing.Point(752, 41)
        Me.cmdSearch.Name = "cmdSearch"
        Me.cmdSearch.Size = New System.Drawing.Size(75, 23)
        Me.cmdSearch.TabIndex = 4
        Me.cmdSearch.Text = "&Search"
        Me.cmdSearch.UseVisualStyleBackColor = True
        '
        'numYear
        '
        Me.numYear.Location = New System.Drawing.Point(126, 64)
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
        Me.cboPeriodType.Location = New System.Drawing.Point(262, 63)
        Me.cboPeriodType.Name = "cboPeriodType"
        Me.cboPeriodType.Size = New System.Drawing.Size(143, 21)
        Me.cboPeriodType.TabIndex = 3
        '
        'lblConfirmPwd
        '
        Me.lblConfirmPwd.AutoSize = True
        Me.lblConfirmPwd.Location = New System.Drawing.Point(189, 66)
        Me.lblConfirmPwd.Name = "lblConfirmPwd"
        Me.lblConfirmPwd.Size = New System.Drawing.Size(67, 13)
        Me.lblConfirmPwd.TabIndex = 2
        Me.lblConfirmPwd.Text = "Period &Type:"
        '
        'lblYear
        '
        Me.lblYear.AutoSize = True
        Me.lblYear.Location = New System.Drawing.Point(51, 66)
        Me.lblYear.Name = "lblYear"
        Me.lblYear.Size = New System.Drawing.Size(69, 13)
        Me.lblYear.TabIndex = 0
        Me.lblYear.Text = "Budget &Year:"
        '
        'optBudget
        '
        Me.optBudget.Checked = True
        Me.optBudget.Location = New System.Drawing.Point(16, 41)
        Me.optBudget.Name = "optBudget"
        Me.optBudget.Size = New System.Drawing.Size(355, 17)
        Me.optBudget.TabIndex = 12
        Me.optBudget.TabStop = True
        Me.optBudget.Text = "Budget Journal"
        Me.optBudget.UseVisualStyleBackColor = True
        '
        'optAccount
        '
        Me.optAccount.Location = New System.Drawing.Point(16, 90)
        Me.optAccount.Name = "optAccount"
        Me.optAccount.Size = New System.Drawing.Size(355, 17)
        Me.optAccount.TabIndex = 13
        Me.optAccount.Text = "Operated by Admin"
        Me.optAccount.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(51, 115)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(59, 13)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "&From Date:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(221, 115)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(49, 13)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "&To Date:"
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "yyy/MM/dd"
        Me.dtpFrom.Enabled = False
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(116, 111)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(89, 20)
        Me.dtpFrom.TabIndex = 16
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "yyy/MM/dd"
        Me.dtpTo.Enabled = False
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(276, 111)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(89, 20)
        Me.dtpTo.TabIndex = 17
        '
        'numProjectNo
        '
        Me.numProjectNo.Location = New System.Drawing.Point(490, 64)
        Me.numProjectNo.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numProjectNo.Name = "numProjectNo"
        Me.numProjectNo.Size = New System.Drawing.Size(38, 20)
        Me.numProjectNo.TabIndex = 18
        Me.numProjectNo.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'lblProjectNo
        '
        Me.lblProjectNo.AutoSize = True
        Me.lblProjectNo.Location = New System.Drawing.Point(424, 66)
        Me.lblProjectNo.Name = "lblProjectNo"
        Me.lblProjectNo.Size = New System.Drawing.Size(60, 13)
        Me.lblProjectNo.TabIndex = 19
        Me.lblProjectNo.Text = "Project No:"
        '
        'frmBG0390
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdClose
        Me.ClientSize = New System.Drawing.Size(839, 567)
        Me.Controls.Add(Me.numProjectNo)
        Me.Controls.Add(Me.lblProjectNo)
        Me.Controls.Add(Me.dtpTo)
        Me.Controls.Add(Me.dtpFrom)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.optAccount)
        Me.Controls.Add(Me.optBudget)
        Me.Controls.Add(Me.cmdSearch)
        Me.Controls.Add(Me.numYear)
        Me.Controls.Add(Me.cboPeriodType)
        Me.Controls.Add(Me.lblConfirmPwd)
        Me.Controls.Add(Me.lblYear)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.grvLog)
        Me.Controls.Add(Me.lblFormTitle)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "frmBG0390"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "frmBG0390"
        CType(Me.grvLog, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numYear, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numProjectNo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblFormTitle As System.Windows.Forms.Label
    Friend WithEvents grvLog As System.Windows.Forms.DataGridView
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents cmdSearch As System.Windows.Forms.Button
    Friend WithEvents numYear As System.Windows.Forms.NumericUpDown
    Friend WithEvents cboPeriodType As System.Windows.Forms.ComboBox
    Friend WithEvents lblConfirmPwd As System.Windows.Forms.Label
    Friend WithEvents lblYear As System.Windows.Forms.Label
    Friend WithEvents optBudget As System.Windows.Forms.RadioButton
    Friend WithEvents optAccount As System.Windows.Forms.RadioButton
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column13 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column12 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column9 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PROJECT_NO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column10 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column11 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column8 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column14 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents numProjectNo As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblProjectNo As System.Windows.Forms.Label
End Class
