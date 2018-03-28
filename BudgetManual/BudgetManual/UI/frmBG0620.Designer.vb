<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBG0620
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
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBG0620))
        Me.lblFormTitle = New System.Windows.Forms.Label
        Me.cmdDelete = New System.Windows.Forms.Button
        Me.fraInfo = New System.Windows.Forms.GroupBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtRemarks = New System.Windows.Forms.TextBox
        Me.lblExpenseType = New System.Windows.Forms.Label
        Me.cboExpenseType = New System.Windows.Forms.ComboBox
        Me.lblAssetGroup = New System.Windows.Forms.Label
        Me.cboAssetGroup = New System.Windows.Forms.ComboBox
        Me.chkActive = New System.Windows.Forms.CheckBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.cboPersonInCharge = New System.Windows.Forms.ComboBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.cboDepartment = New System.Windows.Forms.ComboBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.cboCost = New System.Windows.Forms.ComboBox
        Me.lblCost = New System.Windows.Forms.Label
        Me.cboCostType = New System.Windows.Forms.ComboBox
        Me.lblCostType = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtCostCenter = New System.Windows.Forms.TextBox
        Me.cboAccount = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.cboBudgetType = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtOrderNo = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtOrderName = New System.Windows.Forms.TextBox
        Me.cmdAdd = New System.Windows.Forms.Button
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.grvMaster = New System.Windows.Forms.DataGridView
        Me.BUDGETORDERNODataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BUDGETORDERNAMEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BUDGETTYPEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ACCOUNTNODataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.COSTCENTERDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.COSTTYPEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.COSTDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ASSETGROUPNODataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DEPTNODataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PERSON_IN_CHARGE_NO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ACTIVEFLAGDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.EXPENSE_TYPE = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PIC_SHOW_FLAG = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CREATEUSERIDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CREATEDATEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.UPDATEUSERIDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.UPDATEDATEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataSet1 = New System.Data.DataSet
        Me.dtMasterBudgetOrder = New System.Data.DataTable
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
        Me.cmdImport = New System.Windows.Forms.Button
        Me.cmdExport = New System.Windows.Forms.Button
        Me.fraFilter = New System.Windows.Forms.GroupBox
        Me.cmdFilter = New System.Windows.Forms.Button
        Me.Label11 = New System.Windows.Forms.Label
        Me.cboExpenseTypeFilter = New System.Windows.Forms.ComboBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.cboAssetGroupFilter = New System.Windows.Forms.ComboBox
        Me.chkActiveFilter = New System.Windows.Forms.CheckBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.cboPersonInChargeFilter = New System.Windows.Forms.ComboBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.cboDepartmentFilter = New System.Windows.Forms.ComboBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.cboCostFilter = New System.Windows.Forms.ComboBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.cboCostTypeFilter = New System.Windows.Forms.ComboBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.txtCostCenterFilter = New System.Windows.Forms.TextBox
        Me.cboAccountFilter = New System.Windows.Forms.ComboBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.cboBudgetTypeFilter = New System.Windows.Forms.ComboBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.txtOrderNoFilter = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.txtOrderNameFilter = New System.Windows.Forms.TextBox
        Me.cmdClearFilter = New System.Windows.Forms.Button
        Me.fraInfo.SuspendLayout()
        CType(Me.grvMaster, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataSet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtMasterBudgetOrder, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraFilter.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblFormTitle
        '
        Me.lblFormTitle.AutoSize = True
        Me.lblFormTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblFormTitle.Location = New System.Drawing.Point(12, 9)
        Me.lblFormTitle.Name = "lblFormTitle"
        Me.lblFormTitle.Size = New System.Drawing.Size(204, 24)
        Me.lblFormTitle.TabIndex = 0
        Me.lblFormTitle.Text = "Budget Order Master"
        '
        'cmdDelete
        '
        Me.cmdDelete.Location = New System.Drawing.Point(171, 760)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(75, 23)
        Me.cmdDelete.TabIndex = 6
        Me.cmdDelete.Text = "&Delete"
        Me.cmdDelete.UseVisualStyleBackColor = True
        '
        'fraInfo
        '
        Me.fraInfo.Controls.Add(Me.Label6)
        Me.fraInfo.Controls.Add(Me.txtRemarks)
        Me.fraInfo.Controls.Add(Me.lblExpenseType)
        Me.fraInfo.Controls.Add(Me.cboExpenseType)
        Me.fraInfo.Controls.Add(Me.lblAssetGroup)
        Me.fraInfo.Controls.Add(Me.cboAssetGroup)
        Me.fraInfo.Controls.Add(Me.chkActive)
        Me.fraInfo.Controls.Add(Me.Label10)
        Me.fraInfo.Controls.Add(Me.cboPersonInCharge)
        Me.fraInfo.Controls.Add(Me.Label9)
        Me.fraInfo.Controls.Add(Me.cboDepartment)
        Me.fraInfo.Controls.Add(Me.Label8)
        Me.fraInfo.Controls.Add(Me.cboCost)
        Me.fraInfo.Controls.Add(Me.lblCost)
        Me.fraInfo.Controls.Add(Me.cboCostType)
        Me.fraInfo.Controls.Add(Me.lblCostType)
        Me.fraInfo.Controls.Add(Me.Label5)
        Me.fraInfo.Controls.Add(Me.txtCostCenter)
        Me.fraInfo.Controls.Add(Me.cboAccount)
        Me.fraInfo.Controls.Add(Me.Label4)
        Me.fraInfo.Controls.Add(Me.cboBudgetType)
        Me.fraInfo.Controls.Add(Me.Label3)
        Me.fraInfo.Controls.Add(Me.Label1)
        Me.fraInfo.Controls.Add(Me.txtOrderNo)
        Me.fraInfo.Controls.Add(Me.Label2)
        Me.fraInfo.Controls.Add(Me.txtOrderName)
        Me.fraInfo.Location = New System.Drawing.Point(12, 534)
        Me.fraInfo.Name = "fraInfo"
        Me.fraInfo.Size = New System.Drawing.Size(760, 219)
        Me.fraInfo.TabIndex = 3
        Me.fraInfo.TabStop = False
        Me.fraInfo.Text = "Budget Order Information"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(25, 185)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(52, 13)
        Me.Label6.TabIndex = 24
        Me.Label6.Text = "Remarks:"
        '
        'txtRemarks
        '
        Me.txtRemarks.Location = New System.Drawing.Point(135, 182)
        Me.txtRemarks.MaxLength = 500
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.Size = New System.Drawing.Size(609, 20)
        Me.txtRemarks.TabIndex = 25
        '
        'lblExpenseType
        '
        Me.lblExpenseType.AutoSize = True
        Me.lblExpenseType.Location = New System.Drawing.Point(25, 160)
        Me.lblExpenseType.Name = "lblExpenseType"
        Me.lblExpenseType.Size = New System.Drawing.Size(78, 13)
        Me.lblExpenseType.TabIndex = 20
        Me.lblExpenseType.Text = "Expense Type:"
        '
        'cboExpenseType
        '
        Me.cboExpenseType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboExpenseType.FormattingEnabled = True
        Me.cboExpenseType.Location = New System.Drawing.Point(135, 155)
        Me.cboExpenseType.Name = "cboExpenseType"
        Me.cboExpenseType.Size = New System.Drawing.Size(95, 21)
        Me.cboExpenseType.TabIndex = 21
        '
        'lblAssetGroup
        '
        Me.lblAssetGroup.AutoSize = True
        Me.lblAssetGroup.Location = New System.Drawing.Point(406, 78)
        Me.lblAssetGroup.Name = "lblAssetGroup"
        Me.lblAssetGroup.Size = New System.Drawing.Size(68, 13)
        Me.lblAssetGroup.TabIndex = 10
        Me.lblAssetGroup.Text = "Asset Group:"
        '
        'cboAssetGroup
        '
        Me.cboAssetGroup.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAssetGroup.FormattingEnabled = True
        Me.cboAssetGroup.Location = New System.Drawing.Point(504, 75)
        Me.cboAssetGroup.Name = "cboAssetGroup"
        Me.cboAssetGroup.Size = New System.Drawing.Size(240, 21)
        Me.cboAssetGroup.TabIndex = 11
        '
        'chkActive
        '
        Me.chkActive.AutoSize = True
        Me.chkActive.Location = New System.Drawing.Point(504, 159)
        Me.chkActive.Name = "chkActive"
        Me.chkActive.Size = New System.Drawing.Size(15, 14)
        Me.chkActive.TabIndex = 23
        Me.chkActive.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(406, 159)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(40, 13)
        Me.Label10.TabIndex = 22
        Me.Label10.Text = "Active:"
        '
        'cboPersonInCharge
        '
        Me.cboPersonInCharge.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPersonInCharge.FormattingEnabled = True
        Me.cboPersonInCharge.Location = New System.Drawing.Point(504, 129)
        Me.cboPersonInCharge.Name = "cboPersonInCharge"
        Me.cboPersonInCharge.Size = New System.Drawing.Size(240, 21)
        Me.cboPersonInCharge.TabIndex = 19
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(406, 132)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(92, 13)
        Me.Label9.TabIndex = 18
        Me.Label9.Text = "Person In Charge:"
        '
        'cboDepartment
        '
        Me.cboDepartment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDepartment.FormattingEnabled = True
        Me.cboDepartment.Location = New System.Drawing.Point(504, 102)
        Me.cboDepartment.Name = "cboDepartment"
        Me.cboDepartment.Size = New System.Drawing.Size(240, 21)
        Me.cboDepartment.TabIndex = 15
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(406, 105)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(65, 13)
        Me.Label8.TabIndex = 14
        Me.Label8.Text = "Department:"
        '
        'cboCost
        '
        Me.cboCost.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCost.FormattingEnabled = True
        Me.cboCost.Location = New System.Drawing.Point(504, 48)
        Me.cboCost.Name = "cboCost"
        Me.cboCost.Size = New System.Drawing.Size(95, 21)
        Me.cboCost.TabIndex = 7
        '
        'lblCost
        '
        Me.lblCost.AutoSize = True
        Me.lblCost.Location = New System.Drawing.Point(406, 51)
        Me.lblCost.Name = "lblCost"
        Me.lblCost.Size = New System.Drawing.Size(31, 13)
        Me.lblCost.TabIndex = 6
        Me.lblCost.Text = "Cost:"
        '
        'cboCostType
        '
        Me.cboCostType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCostType.FormattingEnabled = True
        Me.cboCostType.Location = New System.Drawing.Point(504, 21)
        Me.cboCostType.Name = "cboCostType"
        Me.cboCostType.Size = New System.Drawing.Size(95, 21)
        Me.cboCostType.TabIndex = 3
        '
        'lblCostType
        '
        Me.lblCostType.AutoSize = True
        Me.lblCostType.Location = New System.Drawing.Point(406, 24)
        Me.lblCostType.Name = "lblCostType"
        Me.lblCostType.Size = New System.Drawing.Size(58, 13)
        Me.lblCostType.TabIndex = 2
        Me.lblCostType.Text = "Cost Type:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(25, 132)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(65, 13)
        Me.Label5.TabIndex = 16
        Me.Label5.Text = "Cost Center:"
        '
        'txtCostCenter
        '
        Me.txtCostCenter.Location = New System.Drawing.Point(135, 129)
        Me.txtCostCenter.MaxLength = 6
        Me.txtCostCenter.Name = "txtCostCenter"
        Me.txtCostCenter.Size = New System.Drawing.Size(80, 20)
        Me.txtCostCenter.TabIndex = 17
        '
        'cboAccount
        '
        Me.cboAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAccount.FormattingEnabled = True
        Me.cboAccount.Location = New System.Drawing.Point(135, 102)
        Me.cboAccount.Name = "cboAccount"
        Me.cboAccount.Size = New System.Drawing.Size(240, 21)
        Me.cboAccount.TabIndex = 13
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(25, 105)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(50, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Account:"
        '
        'cboBudgetType
        '
        Me.cboBudgetType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBudgetType.FormattingEnabled = True
        Me.cboBudgetType.Location = New System.Drawing.Point(135, 75)
        Me.cboBudgetType.Name = "cboBudgetType"
        Me.cboBudgetType.Size = New System.Drawing.Size(95, 21)
        Me.cboBudgetType.TabIndex = 9
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(25, 78)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(71, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Budget Type:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(25, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(90, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Budget Order No:"
        '
        'txtOrderNo
        '
        Me.txtOrderNo.Location = New System.Drawing.Point(135, 23)
        Me.txtOrderNo.MaxLength = 13
        Me.txtOrderNo.Name = "txtOrderNo"
        Me.txtOrderNo.Size = New System.Drawing.Size(80, 20)
        Me.txtOrderNo.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(25, 52)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(104, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Budget Order Name:"
        '
        'txtOrderName
        '
        Me.txtOrderName.Location = New System.Drawing.Point(135, 49)
        Me.txtOrderName.MaxLength = 100
        Me.txtOrderName.Name = "txtOrderName"
        Me.txtOrderName.Size = New System.Drawing.Size(240, 20)
        Me.txtOrderName.TabIndex = 5
        '
        'cmdAdd
        '
        Me.cmdAdd.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdAdd.Location = New System.Drawing.Point(9, 760)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(75, 23)
        Me.cmdAdd.TabIndex = 4
        Me.cmdAdd.Text = "&Add New"
        Me.cmdAdd.UseVisualStyleBackColor = True
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(90, 760)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(75, 23)
        Me.cmdSave.TabIndex = 5
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'cmdClose
        '
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdClose.Location = New System.Drawing.Point(697, 759)
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
        Me.grvMaster.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.BUDGETORDERNODataGridViewTextBoxColumn, Me.BUDGETORDERNAMEDataGridViewTextBoxColumn, Me.BUDGETTYPEDataGridViewTextBoxColumn, Me.ACCOUNTNODataGridViewTextBoxColumn, Me.COSTCENTERDataGridViewTextBoxColumn, Me.COSTTYPEDataGridViewTextBoxColumn, Me.COSTDataGridViewTextBoxColumn, Me.ASSETGROUPNODataGridViewTextBoxColumn, Me.DEPTNODataGridViewTextBoxColumn, Me.PERSON_IN_CHARGE_NO, Me.ACTIVEFLAGDataGridViewTextBoxColumn, Me.EXPENSE_TYPE, Me.PIC_SHOW_FLAG, Me.colRemarks, Me.CREATEUSERIDDataGridViewTextBoxColumn, Me.CREATEDATEDataGridViewTextBoxColumn, Me.UPDATEUSERIDDataGridViewTextBoxColumn, Me.UPDATEDATEDataGridViewTextBoxColumn})
        Me.grvMaster.DataMember = "dtMasterBudgetOrder"
        Me.grvMaster.DataSource = Me.DataSet1
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.Color.Azure
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.NullValue = "-"
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grvMaster.DefaultCellStyle = DataGridViewCellStyle5
        Me.grvMaster.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.grvMaster.EnableHeadersVisualStyles = False
        Me.grvMaster.Location = New System.Drawing.Point(12, 248)
        Me.grvMaster.MultiSelect = False
        Me.grvMaster.Name = "grvMaster"
        Me.grvMaster.ReadOnly = True
        Me.grvMaster.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grvMaster.RowHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.grvMaster.RowHeadersVisible = False
        Me.grvMaster.RowHeadersWidth = 30
        Me.grvMaster.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grvMaster.RowsDefaultCellStyle = DataGridViewCellStyle7
        Me.grvMaster.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.grvMaster.Size = New System.Drawing.Size(760, 280)
        Me.grvMaster.StandardTab = True
        Me.grvMaster.TabIndex = 2
        '
        'BUDGETORDERNODataGridViewTextBoxColumn
        '
        Me.BUDGETORDERNODataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.BUDGETORDERNODataGridViewTextBoxColumn.DataPropertyName = "BUDGET_ORDER_NO"
        Me.BUDGETORDERNODataGridViewTextBoxColumn.HeaderText = "Budget Order No."
        Me.BUDGETORDERNODataGridViewTextBoxColumn.Name = "BUDGETORDERNODataGridViewTextBoxColumn"
        Me.BUDGETORDERNODataGridViewTextBoxColumn.ReadOnly = True
        Me.BUDGETORDERNODataGridViewTextBoxColumn.Width = 132
        '
        'BUDGETORDERNAMEDataGridViewTextBoxColumn
        '
        Me.BUDGETORDERNAMEDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.BUDGETORDERNAMEDataGridViewTextBoxColumn.DataPropertyName = "BUDGET_ORDER_NAME"
        Me.BUDGETORDERNAMEDataGridViewTextBoxColumn.HeaderText = "Budget Order Name"
        Me.BUDGETORDERNAMEDataGridViewTextBoxColumn.Name = "BUDGETORDERNAMEDataGridViewTextBoxColumn"
        Me.BUDGETORDERNAMEDataGridViewTextBoxColumn.ReadOnly = True
        Me.BUDGETORDERNAMEDataGridViewTextBoxColumn.Width = 146
        '
        'BUDGETTYPEDataGridViewTextBoxColumn
        '
        Me.BUDGETTYPEDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.BUDGETTYPEDataGridViewTextBoxColumn.DataPropertyName = "BUDGET_TYPE"
        Me.BUDGETTYPEDataGridViewTextBoxColumn.HeaderText = "Budget Type"
        Me.BUDGETTYPEDataGridViewTextBoxColumn.Name = "BUDGETTYPEDataGridViewTextBoxColumn"
        Me.BUDGETTYPEDataGridViewTextBoxColumn.ReadOnly = True
        Me.BUDGETTYPEDataGridViewTextBoxColumn.Width = 104
        '
        'ACCOUNTNODataGridViewTextBoxColumn
        '
        Me.ACCOUNTNODataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.ACCOUNTNODataGridViewTextBoxColumn.DataPropertyName = "ACCOUNT_NO"
        Me.ACCOUNTNODataGridViewTextBoxColumn.HeaderText = "Account No."
        Me.ACCOUNTNODataGridViewTextBoxColumn.Name = "ACCOUNTNODataGridViewTextBoxColumn"
        Me.ACCOUNTNODataGridViewTextBoxColumn.ReadOnly = True
        Me.ACCOUNTNODataGridViewTextBoxColumn.Width = 101
        '
        'COSTCENTERDataGridViewTextBoxColumn
        '
        Me.COSTCENTERDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.COSTCENTERDataGridViewTextBoxColumn.DataPropertyName = "COST_CENTER"
        Me.COSTCENTERDataGridViewTextBoxColumn.HeaderText = "Cost Center"
        Me.COSTCENTERDataGridViewTextBoxColumn.Name = "COSTCENTERDataGridViewTextBoxColumn"
        Me.COSTCENTERDataGridViewTextBoxColumn.ReadOnly = True
        '
        'COSTTYPEDataGridViewTextBoxColumn
        '
        Me.COSTTYPEDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.COSTTYPEDataGridViewTextBoxColumn.DataPropertyName = "COST_TYPE"
        Me.COSTTYPEDataGridViewTextBoxColumn.HeaderText = "Cost Type"
        Me.COSTTYPEDataGridViewTextBoxColumn.Name = "COSTTYPEDataGridViewTextBoxColumn"
        Me.COSTTYPEDataGridViewTextBoxColumn.ReadOnly = True
        Me.COSTTYPEDataGridViewTextBoxColumn.Width = 90
        '
        'COSTDataGridViewTextBoxColumn
        '
        Me.COSTDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.COSTDataGridViewTextBoxColumn.DataPropertyName = "COST"
        Me.COSTDataGridViewTextBoxColumn.HeaderText = "Cost"
        Me.COSTDataGridViewTextBoxColumn.Name = "COSTDataGridViewTextBoxColumn"
        Me.COSTDataGridViewTextBoxColumn.ReadOnly = True
        Me.COSTDataGridViewTextBoxColumn.Width = 58
        '
        'ASSETGROUPNODataGridViewTextBoxColumn
        '
        Me.ASSETGROUPNODataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.ASSETGROUPNODataGridViewTextBoxColumn.DataPropertyName = "ASSET_GROUP_NO"
        Me.ASSETGROUPNODataGridViewTextBoxColumn.HeaderText = "Asset Group No."
        Me.ASSETGROUPNODataGridViewTextBoxColumn.Name = "ASSETGROUPNODataGridViewTextBoxColumn"
        Me.ASSETGROUPNODataGridViewTextBoxColumn.ReadOnly = True
        Me.ASSETGROUPNODataGridViewTextBoxColumn.Width = 125
        '
        'DEPTNODataGridViewTextBoxColumn
        '
        Me.DEPTNODataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.DEPTNODataGridViewTextBoxColumn.DataPropertyName = "DEPT_NO"
        Me.DEPTNODataGridViewTextBoxColumn.HeaderText = "Department No."
        Me.DEPTNODataGridViewTextBoxColumn.Name = "DEPTNODataGridViewTextBoxColumn"
        Me.DEPTNODataGridViewTextBoxColumn.ReadOnly = True
        Me.DEPTNODataGridViewTextBoxColumn.Width = 123
        '
        'PERSON_IN_CHARGE_NO
        '
        Me.PERSON_IN_CHARGE_NO.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.PERSON_IN_CHARGE_NO.DataPropertyName = "PERSON_IN_CHARGE_NO"
        Me.PERSON_IN_CHARGE_NO.HeaderText = "Person In Charge"
        Me.PERSON_IN_CHARGE_NO.Name = "PERSON_IN_CHARGE_NO"
        Me.PERSON_IN_CHARGE_NO.ReadOnly = True
        Me.PERSON_IN_CHARGE_NO.Width = 132
        '
        'ACTIVEFLAGDataGridViewTextBoxColumn
        '
        Me.ACTIVEFLAGDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.ACTIVEFLAGDataGridViewTextBoxColumn.DataPropertyName = "ACTIVE_FLAG"
        Me.ACTIVEFLAGDataGridViewTextBoxColumn.HeaderText = "Active"
        Me.ACTIVEFLAGDataGridViewTextBoxColumn.Name = "ACTIVEFLAGDataGridViewTextBoxColumn"
        Me.ACTIVEFLAGDataGridViewTextBoxColumn.ReadOnly = True
        Me.ACTIVEFLAGDataGridViewTextBoxColumn.Width = 67
        '
        'EXPENSE_TYPE
        '
        Me.EXPENSE_TYPE.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.EXPENSE_TYPE.DataPropertyName = "EXPENSE_TYPE"
        Me.EXPENSE_TYPE.HeaderText = "Expense Type"
        Me.EXPENSE_TYPE.Name = "EXPENSE_TYPE"
        Me.EXPENSE_TYPE.ReadOnly = True
        Me.EXPENSE_TYPE.Width = 112
        '
        'PIC_SHOW_FLAG
        '
        Me.PIC_SHOW_FLAG.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.PIC_SHOW_FLAG.DataPropertyName = "PIC_SHOW_FLAG"
        Me.PIC_SHOW_FLAG.HeaderText = "PIC show flag"
        Me.PIC_SHOW_FLAG.Name = "PIC_SHOW_FLAG"
        Me.PIC_SHOW_FLAG.ReadOnly = True
        Me.PIC_SHOW_FLAG.Width = 111
        '
        'colRemarks
        '
        Me.colRemarks.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.colRemarks.DataPropertyName = "REMARKS"
        Me.colRemarks.HeaderText = "Remarks"
        Me.colRemarks.MinimumWidth = 50
        Me.colRemarks.Name = "colRemarks"
        Me.colRemarks.ReadOnly = True
        Me.colRemarks.Width = 74
        '
        'CREATEUSERIDDataGridViewTextBoxColumn
        '
        Me.CREATEUSERIDDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.CREATEUSERIDDataGridViewTextBoxColumn.DataPropertyName = "CREATE_USER_ID"
        Me.CREATEUSERIDDataGridViewTextBoxColumn.HeaderText = "Create User"
        Me.CREATEUSERIDDataGridViewTextBoxColumn.Name = "CREATEUSERIDDataGridViewTextBoxColumn"
        Me.CREATEUSERIDDataGridViewTextBoxColumn.ReadOnly = True
        Me.CREATEUSERIDDataGridViewTextBoxColumn.Width = 101
        '
        'CREATEDATEDataGridViewTextBoxColumn
        '
        Me.CREATEDATEDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.CREATEDATEDataGridViewTextBoxColumn.DataPropertyName = "CREATE_DATE"
        DataGridViewCellStyle3.Format = "G"
        DataGridViewCellStyle3.NullValue = Nothing
        Me.CREATEDATEDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle3
        Me.CREATEDATEDataGridViewTextBoxColumn.HeaderText = "Create Date"
        Me.CREATEDATEDataGridViewTextBoxColumn.Name = "CREATEDATEDataGridViewTextBoxColumn"
        Me.CREATEDATEDataGridViewTextBoxColumn.ReadOnly = True
        Me.CREATEDATEDataGridViewTextBoxColumn.Width = 101
        '
        'UPDATEUSERIDDataGridViewTextBoxColumn
        '
        Me.UPDATEUSERIDDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.UPDATEUSERIDDataGridViewTextBoxColumn.DataPropertyName = "UPDATE_USER_ID"
        Me.UPDATEUSERIDDataGridViewTextBoxColumn.HeaderText = "Update User"
        Me.UPDATEUSERIDDataGridViewTextBoxColumn.Name = "UPDATEUSERIDDataGridViewTextBoxColumn"
        Me.UPDATEUSERIDDataGridViewTextBoxColumn.ReadOnly = True
        Me.UPDATEUSERIDDataGridViewTextBoxColumn.Width = 103
        '
        'UPDATEDATEDataGridViewTextBoxColumn
        '
        Me.UPDATEDATEDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.UPDATEDATEDataGridViewTextBoxColumn.DataPropertyName = "UPDATE_DATE"
        DataGridViewCellStyle4.Format = "G"
        DataGridViewCellStyle4.NullValue = Nothing
        Me.UPDATEDATEDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle4
        Me.UPDATEDATEDataGridViewTextBoxColumn.HeaderText = "Update Date"
        Me.UPDATEDATEDataGridViewTextBoxColumn.Name = "UPDATEDATEDataGridViewTextBoxColumn"
        Me.UPDATEDATEDataGridViewTextBoxColumn.ReadOnly = True
        Me.UPDATEDATEDataGridViewTextBoxColumn.Width = 103
        '
        'DataSet1
        '
        Me.DataSet1.DataSetName = "NewDataSet"
        Me.DataSet1.Tables.AddRange(New System.Data.DataTable() {Me.dtMasterBudgetOrder})
        '
        'dtMasterBudgetOrder
        '
        Me.dtMasterBudgetOrder.Columns.AddRange(New System.Data.DataColumn() {Me.DataColumn1, Me.DataColumn2, Me.DataColumn3, Me.DataColumn4, Me.DataColumn5, Me.DataColumn6, Me.DataColumn7, Me.DataColumn8, Me.DataColumn9, Me.DataColumn10, Me.DataColumn11, Me.DataColumn12, Me.DataColumn13, Me.DataColumn14, Me.DataColumn15, Me.DataColumn16, Me.DataColumn17})
        Me.dtMasterBudgetOrder.Constraints.AddRange(New System.Data.Constraint() {New System.Data.UniqueConstraint("Constraint1", New String() {"BUDGET_ORDER_NO"}, True)})
        Me.dtMasterBudgetOrder.PrimaryKey = New System.Data.DataColumn() {Me.DataColumn1}
        Me.dtMasterBudgetOrder.TableName = "dtMasterBudgetOrder"
        '
        'DataColumn1
        '
        Me.DataColumn1.AllowDBNull = False
        Me.DataColumn1.ColumnName = "BUDGET_ORDER_NO"
        '
        'DataColumn2
        '
        Me.DataColumn2.ColumnName = "BUDGET_ORDER_NAME"
        '
        'DataColumn3
        '
        Me.DataColumn3.ColumnName = "BUDGET_TYPE"
        '
        'DataColumn4
        '
        Me.DataColumn4.ColumnName = "ACCOUNT_NO"
        '
        'DataColumn5
        '
        Me.DataColumn5.ColumnName = "COST_CENTER"
        '
        'DataColumn6
        '
        Me.DataColumn6.ColumnName = "COST_TYPE"
        '
        'DataColumn7
        '
        Me.DataColumn7.ColumnName = "COST"
        '
        'DataColumn8
        '
        Me.DataColumn8.ColumnName = "ASSET_GROUP_NO"
        '
        'DataColumn9
        '
        Me.DataColumn9.ColumnName = "DEPT_NO"
        '
        'DataColumn10
        '
        Me.DataColumn10.ColumnName = "PERSON_IN_CHARGE_NO"
        '
        'DataColumn11
        '
        Me.DataColumn11.ColumnName = "ACTIVE_FLAG"
        '
        'DataColumn12
        '
        Me.DataColumn12.ColumnName = "CREATE_USER_ID"
        '
        'DataColumn13
        '
        Me.DataColumn13.ColumnName = "CREATE_DATE"
        '
        'DataColumn14
        '
        Me.DataColumn14.ColumnName = "UPDATE_USER_ID"
        '
        'DataColumn15
        '
        Me.DataColumn15.ColumnName = "UPDATE_DATE"
        '
        'DataColumn16
        '
        Me.DataColumn16.ColumnName = "EXPENSE_TYPE"
        '
        'DataColumn17
        '
        Me.DataColumn17.ColumnName = "PIC_SHOW_FLAG"
        '
        'cmdImport
        '
        Me.cmdImport.Location = New System.Drawing.Point(261, 760)
        Me.cmdImport.Name = "cmdImport"
        Me.cmdImport.Size = New System.Drawing.Size(75, 23)
        Me.cmdImport.TabIndex = 7
        Me.cmdImport.Text = "&Import"
        Me.cmdImport.UseVisualStyleBackColor = True
        '
        'cmdExport
        '
        Me.cmdExport.Location = New System.Drawing.Point(342, 760)
        Me.cmdExport.Name = "cmdExport"
        Me.cmdExport.Size = New System.Drawing.Size(75, 23)
        Me.cmdExport.TabIndex = 8
        Me.cmdExport.Text = "&Export"
        Me.cmdExport.UseVisualStyleBackColor = True
        '
        'fraFilter
        '
        Me.fraFilter.Controls.Add(Me.cmdClearFilter)
        Me.fraFilter.Controls.Add(Me.cmdFilter)
        Me.fraFilter.Controls.Add(Me.Label11)
        Me.fraFilter.Controls.Add(Me.cboExpenseTypeFilter)
        Me.fraFilter.Controls.Add(Me.Label12)
        Me.fraFilter.Controls.Add(Me.cboAssetGroupFilter)
        Me.fraFilter.Controls.Add(Me.chkActiveFilter)
        Me.fraFilter.Controls.Add(Me.Label13)
        Me.fraFilter.Controls.Add(Me.cboPersonInChargeFilter)
        Me.fraFilter.Controls.Add(Me.Label14)
        Me.fraFilter.Controls.Add(Me.cboDepartmentFilter)
        Me.fraFilter.Controls.Add(Me.Label15)
        Me.fraFilter.Controls.Add(Me.cboCostFilter)
        Me.fraFilter.Controls.Add(Me.Label16)
        Me.fraFilter.Controls.Add(Me.cboCostTypeFilter)
        Me.fraFilter.Controls.Add(Me.Label17)
        Me.fraFilter.Controls.Add(Me.Label18)
        Me.fraFilter.Controls.Add(Me.txtCostCenterFilter)
        Me.fraFilter.Controls.Add(Me.cboAccountFilter)
        Me.fraFilter.Controls.Add(Me.Label19)
        Me.fraFilter.Controls.Add(Me.cboBudgetTypeFilter)
        Me.fraFilter.Controls.Add(Me.Label20)
        Me.fraFilter.Controls.Add(Me.Label21)
        Me.fraFilter.Controls.Add(Me.txtOrderNoFilter)
        Me.fraFilter.Controls.Add(Me.Label22)
        Me.fraFilter.Controls.Add(Me.txtOrderNameFilter)
        Me.fraFilter.Location = New System.Drawing.Point(12, 36)
        Me.fraFilter.Name = "fraFilter"
        Me.fraFilter.Size = New System.Drawing.Size(760, 206)
        Me.fraFilter.TabIndex = 1
        Me.fraFilter.TabStop = False
        Me.fraFilter.Text = "Budget Order Filter"
        '
        'cmdFilter
        '
        Me.cmdFilter.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdFilter.Location = New System.Drawing.Point(588, 177)
        Me.cmdFilter.Name = "cmdFilter"
        Me.cmdFilter.Size = New System.Drawing.Size(75, 23)
        Me.cmdFilter.TabIndex = 24
        Me.cmdFilter.Text = "&Filter"
        Me.cmdFilter.UseVisualStyleBackColor = True
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(25, 160)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(78, 13)
        Me.Label11.TabIndex = 20
        Me.Label11.Text = "Expense Type:"
        '
        'cboExpenseTypeFilter
        '
        Me.cboExpenseTypeFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboExpenseTypeFilter.FormattingEnabled = True
        Me.cboExpenseTypeFilter.Location = New System.Drawing.Point(135, 155)
        Me.cboExpenseTypeFilter.Name = "cboExpenseTypeFilter"
        Me.cboExpenseTypeFilter.Size = New System.Drawing.Size(95, 21)
        Me.cboExpenseTypeFilter.TabIndex = 21
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(406, 78)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(68, 13)
        Me.Label12.TabIndex = 10
        Me.Label12.Text = "Asset Group:"
        '
        'cboAssetGroupFilter
        '
        Me.cboAssetGroupFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAssetGroupFilter.FormattingEnabled = True
        Me.cboAssetGroupFilter.Location = New System.Drawing.Point(504, 75)
        Me.cboAssetGroupFilter.Name = "cboAssetGroupFilter"
        Me.cboAssetGroupFilter.Size = New System.Drawing.Size(240, 21)
        Me.cboAssetGroupFilter.TabIndex = 11
        '
        'chkActiveFilter
        '
        Me.chkActiveFilter.AutoSize = True
        Me.chkActiveFilter.Location = New System.Drawing.Point(504, 159)
        Me.chkActiveFilter.Name = "chkActiveFilter"
        Me.chkActiveFilter.Size = New System.Drawing.Size(15, 14)
        Me.chkActiveFilter.TabIndex = 23
        Me.chkActiveFilter.UseVisualStyleBackColor = True
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(406, 159)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(40, 13)
        Me.Label13.TabIndex = 22
        Me.Label13.Text = "Active:"
        '
        'cboPersonInChargeFilter
        '
        Me.cboPersonInChargeFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPersonInChargeFilter.FormattingEnabled = True
        Me.cboPersonInChargeFilter.Location = New System.Drawing.Point(504, 129)
        Me.cboPersonInChargeFilter.Name = "cboPersonInChargeFilter"
        Me.cboPersonInChargeFilter.Size = New System.Drawing.Size(240, 21)
        Me.cboPersonInChargeFilter.TabIndex = 19
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(406, 132)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(92, 13)
        Me.Label14.TabIndex = 18
        Me.Label14.Text = "Person In Charge:"
        '
        'cboDepartmentFilter
        '
        Me.cboDepartmentFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDepartmentFilter.FormattingEnabled = True
        Me.cboDepartmentFilter.Location = New System.Drawing.Point(504, 102)
        Me.cboDepartmentFilter.Name = "cboDepartmentFilter"
        Me.cboDepartmentFilter.Size = New System.Drawing.Size(240, 21)
        Me.cboDepartmentFilter.TabIndex = 15
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(406, 105)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(65, 13)
        Me.Label15.TabIndex = 14
        Me.Label15.Text = "Department:"
        '
        'cboCostFilter
        '
        Me.cboCostFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCostFilter.FormattingEnabled = True
        Me.cboCostFilter.Location = New System.Drawing.Point(504, 48)
        Me.cboCostFilter.Name = "cboCostFilter"
        Me.cboCostFilter.Size = New System.Drawing.Size(95, 21)
        Me.cboCostFilter.TabIndex = 7
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(406, 51)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(31, 13)
        Me.Label16.TabIndex = 6
        Me.Label16.Text = "Cost:"
        '
        'cboCostTypeFilter
        '
        Me.cboCostTypeFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCostTypeFilter.FormattingEnabled = True
        Me.cboCostTypeFilter.Location = New System.Drawing.Point(504, 21)
        Me.cboCostTypeFilter.Name = "cboCostTypeFilter"
        Me.cboCostTypeFilter.Size = New System.Drawing.Size(95, 21)
        Me.cboCostTypeFilter.TabIndex = 3
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(406, 24)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(58, 13)
        Me.Label17.TabIndex = 2
        Me.Label17.Text = "Cost Type:"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(25, 132)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(65, 13)
        Me.Label18.TabIndex = 16
        Me.Label18.Text = "Cost Center:"
        '
        'txtCostCenterFilter
        '
        Me.txtCostCenterFilter.Location = New System.Drawing.Point(135, 129)
        Me.txtCostCenterFilter.MaxLength = 6
        Me.txtCostCenterFilter.Name = "txtCostCenterFilter"
        Me.txtCostCenterFilter.Size = New System.Drawing.Size(80, 20)
        Me.txtCostCenterFilter.TabIndex = 17
        '
        'cboAccountFilter
        '
        Me.cboAccountFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAccountFilter.FormattingEnabled = True
        Me.cboAccountFilter.Location = New System.Drawing.Point(135, 102)
        Me.cboAccountFilter.Name = "cboAccountFilter"
        Me.cboAccountFilter.Size = New System.Drawing.Size(240, 21)
        Me.cboAccountFilter.TabIndex = 13
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(25, 105)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(50, 13)
        Me.Label19.TabIndex = 12
        Me.Label19.Text = "Account:"
        '
        'cboBudgetTypeFilter
        '
        Me.cboBudgetTypeFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBudgetTypeFilter.FormattingEnabled = True
        Me.cboBudgetTypeFilter.Location = New System.Drawing.Point(135, 75)
        Me.cboBudgetTypeFilter.Name = "cboBudgetTypeFilter"
        Me.cboBudgetTypeFilter.Size = New System.Drawing.Size(95, 21)
        Me.cboBudgetTypeFilter.TabIndex = 9
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(25, 78)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(71, 13)
        Me.Label20.TabIndex = 8
        Me.Label20.Text = "Budget Type:"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(25, 26)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(90, 13)
        Me.Label21.TabIndex = 0
        Me.Label21.Text = "Budget Order No:"
        '
        'txtOrderNoFilter
        '
        Me.txtOrderNoFilter.Location = New System.Drawing.Point(135, 23)
        Me.txtOrderNoFilter.MaxLength = 13
        Me.txtOrderNoFilter.Name = "txtOrderNoFilter"
        Me.txtOrderNoFilter.Size = New System.Drawing.Size(80, 20)
        Me.txtOrderNoFilter.TabIndex = 1
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(25, 52)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(104, 13)
        Me.Label22.TabIndex = 4
        Me.Label22.Text = "Budget Order Name:"
        '
        'txtOrderNameFilter
        '
        Me.txtOrderNameFilter.Location = New System.Drawing.Point(135, 49)
        Me.txtOrderNameFilter.MaxLength = 100
        Me.txtOrderNameFilter.Name = "txtOrderNameFilter"
        Me.txtOrderNameFilter.Size = New System.Drawing.Size(240, 20)
        Me.txtOrderNameFilter.TabIndex = 5
        '
        'cmdClearFilter
        '
        Me.cmdClearFilter.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdClearFilter.Location = New System.Drawing.Point(669, 177)
        Me.cmdClearFilter.Name = "cmdClearFilter"
        Me.cmdClearFilter.Size = New System.Drawing.Size(75, 23)
        Me.cmdClearFilter.TabIndex = 25
        Me.cmdClearFilter.Text = "Clear"
        Me.cmdClearFilter.UseVisualStyleBackColor = True
        '
        'frmBG0620
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
        Me.Name = "frmBG0620"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "frmBG0620"
        Me.fraInfo.ResumeLayout(False)
        Me.fraInfo.PerformLayout()
        CType(Me.grvMaster, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataSet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtMasterBudgetOrder, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraFilter.ResumeLayout(False)
        Me.fraFilter.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblFormTitle As System.Windows.Forms.Label
    Friend WithEvents cmdDelete As System.Windows.Forms.Button
    Friend WithEvents fraInfo As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtOrderNo As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtOrderName As System.Windows.Forms.TextBox
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents grvMaster As System.Windows.Forms.DataGridView
    Friend WithEvents cboAccount As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboBudgetType As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents cboPersonInCharge As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cboDepartment As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents cboCost As System.Windows.Forms.ComboBox
    Friend WithEvents lblCost As System.Windows.Forms.Label
    Friend WithEvents cboCostType As System.Windows.Forms.ComboBox
    Friend WithEvents lblCostType As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtCostCenter As System.Windows.Forms.TextBox
    Friend WithEvents chkActive As System.Windows.Forms.CheckBox
    Friend WithEvents cmdImport As System.Windows.Forms.Button
    Friend WithEvents cmdExport As System.Windows.Forms.Button
    Friend WithEvents lblAssetGroup As System.Windows.Forms.Label
    Friend WithEvents cboAssetGroup As System.Windows.Forms.ComboBox
    Friend WithEvents DataSet1 As System.Data.DataSet
    Friend WithEvents dtMasterBudgetOrder As System.Data.DataTable
    Friend WithEvents DataColumn1 As System.Data.DataColumn
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
    Friend WithEvents lblExpenseType As System.Windows.Forms.Label
    Friend WithEvents cboExpenseType As System.Windows.Forms.ComboBox
    Friend WithEvents DataColumn16 As System.Data.DataColumn
    Friend WithEvents DataColumn17 As System.Data.DataColumn
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents BUDGETORDERNODataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BUDGETORDERNAMEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BUDGETTYPEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ACCOUNTNODataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents COSTCENTERDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents COSTTYPEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents COSTDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ASSETGROUPNODataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DEPTNODataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PERSON_IN_CHARGE_NO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ACTIVEFLAGDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EXPENSE_TYPE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PIC_SHOW_FLAG As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CREATEUSERIDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CREATEDATEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents UPDATEUSERIDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents UPDATEDATEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents fraFilter As System.Windows.Forms.GroupBox
    Friend WithEvents cmdFilter As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents cboExpenseTypeFilter As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents cboAssetGroupFilter As System.Windows.Forms.ComboBox
    Friend WithEvents chkActiveFilter As System.Windows.Forms.CheckBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents cboPersonInChargeFilter As System.Windows.Forms.ComboBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents cboDepartmentFilter As System.Windows.Forms.ComboBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents cboCostFilter As System.Windows.Forms.ComboBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents cboCostTypeFilter As System.Windows.Forms.ComboBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtCostCenterFilter As System.Windows.Forms.TextBox
    Friend WithEvents cboAccountFilter As System.Windows.Forms.ComboBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents cboBudgetTypeFilter As System.Windows.Forms.ComboBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents txtOrderNoFilter As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents txtOrderNameFilter As System.Windows.Forms.TextBox
    Friend WithEvents cmdClearFilter As System.Windows.Forms.Button
End Class
