<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBG0610
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBG0610))
        Me.lblFormTitle = New System.Windows.Forms.Label
        Me.grvMaster = New System.Windows.Forms.DataGridView
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column5 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column6 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column7 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.cmdClose = New System.Windows.Forms.Button
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdAdd = New System.Windows.Forms.Button
        Me.fraUserInfo = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtUserID = New System.Windows.Forms.TextBox
        Me.txtPwd2 = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtUserName = New System.Windows.Forms.TextBox
        Me.txtPwd1 = New System.Windows.Forms.TextBox
        Me.cboUserLevel = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.chkExpired = New System.Windows.Forms.CheckBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtEmail = New System.Windows.Forms.TextBox
        Me.cboUserPIC = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.cmdExport = New System.Windows.Forms.Button
        Me.cmdImport = New System.Windows.Forms.Button
        Me.chkHideExpiredUser = New System.Windows.Forms.CheckBox
        Me.fraFilter = New System.Windows.Forms.GroupBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtUserIDFiter = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtUserNameFilter = New System.Windows.Forms.TextBox
        Me.cboUserLevelFilter = New System.Windows.Forms.ComboBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.cboUserPICFilter = New System.Windows.Forms.ComboBox
        Me.cmdClearFilter = New System.Windows.Forms.Button
        Me.cmdFilter = New System.Windows.Forms.Button
        CType(Me.grvMaster, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraUserInfo.SuspendLayout()
        Me.fraFilter.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblFormTitle
        '
        Me.lblFormTitle.AutoSize = True
        Me.lblFormTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblFormTitle.Location = New System.Drawing.Point(12, 9)
        Me.lblFormTitle.Name = "lblFormTitle"
        Me.lblFormTitle.Size = New System.Drawing.Size(121, 24)
        Me.lblFormTitle.TabIndex = 0
        Me.lblFormTitle.Text = "User Master"
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
        Me.grvMaster.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2, Me.Column3, Me.Column4, Me.Column5, Me.Column6, Me.Column7})
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
        Me.grvMaster.Location = New System.Drawing.Point(12, 149)
        Me.grvMaster.MultiSelect = False
        Me.grvMaster.Name = "grvMaster"
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
        Me.grvMaster.Size = New System.Drawing.Size(760, 336)
        Me.grvMaster.StandardTab = True
        Me.grvMaster.TabIndex = 2
        '
        'Column1
        '
        Me.Column1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Column1.DataPropertyName = "USER_ID"
        Me.Column1.HeaderText = "User ID"
        Me.Column1.MinimumWidth = 73
        Me.Column1.Name = "Column1"
        Me.Column1.Width = 73
        '
        'Column2
        '
        Me.Column2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Column2.DataPropertyName = "USER_NAME"
        Me.Column2.HeaderText = "User Name"
        Me.Column2.MinimumWidth = 169
        Me.Column2.Name = "Column2"
        Me.Column2.Width = 169
        '
        'Column3
        '
        Me.Column3.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Column3.DataPropertyName = "USER_LEVEL_NAME"
        Me.Column3.HeaderText = "User Level"
        Me.Column3.MinimumWidth = 150
        Me.Column3.Name = "Column3"
        Me.Column3.Width = 150
        '
        'Column4
        '
        Me.Column4.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Column4.DataPropertyName = "PERSON_IN_CHARGE_NO"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Column4.DefaultCellStyle = DataGridViewCellStyle3
        Me.Column4.HeaderText = "Person In Charge"
        Me.Column4.MinimumWidth = 120
        Me.Column4.Name = "Column4"
        Me.Column4.Width = 120
        '
        'Column5
        '
        Me.Column5.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Column5.DataPropertyName = "EMAIL"
        Me.Column5.HeaderText = "Email"
        Me.Column5.MinimumWidth = 166
        Me.Column5.Name = "Column5"
        Me.Column5.Width = 166
        '
        'Column6
        '
        Me.Column6.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Column6.DataPropertyName = "EXPIRED"
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Column6.DefaultCellStyle = DataGridViewCellStyle4
        Me.Column6.HeaderText = "Expired"
        Me.Column6.MinimumWidth = 70
        Me.Column6.Name = "Column6"
        Me.Column6.Width = 70
        '
        'Column7
        '
        Me.Column7.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Column7.DataPropertyName = "PASSWORD"
        Me.Column7.HeaderText = "PWD"
        Me.Column7.MinimumWidth = 70
        Me.Column7.Name = "Column7"
        Me.Column7.Visible = False
        Me.Column7.Width = 70
        '
        'cmdClose
        '
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdClose.Location = New System.Drawing.Point(699, 637)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 23)
        Me.cmdClose.TabIndex = 8
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(93, 637)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(75, 23)
        Me.cmdSave.TabIndex = 5
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'cmdAdd
        '
        Me.cmdAdd.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdAdd.Location = New System.Drawing.Point(12, 637)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(75, 23)
        Me.cmdAdd.TabIndex = 4
        Me.cmdAdd.Text = "&Add New"
        Me.cmdAdd.UseVisualStyleBackColor = True
        '
        'fraUserInfo
        '
        Me.fraUserInfo.Controls.Add(Me.Label1)
        Me.fraUserInfo.Controls.Add(Me.Label8)
        Me.fraUserInfo.Controls.Add(Me.txtUserID)
        Me.fraUserInfo.Controls.Add(Me.txtPwd2)
        Me.fraUserInfo.Controls.Add(Me.Label2)
        Me.fraUserInfo.Controls.Add(Me.Label7)
        Me.fraUserInfo.Controls.Add(Me.txtUserName)
        Me.fraUserInfo.Controls.Add(Me.txtPwd1)
        Me.fraUserInfo.Controls.Add(Me.cboUserLevel)
        Me.fraUserInfo.Controls.Add(Me.Label6)
        Me.fraUserInfo.Controls.Add(Me.Label3)
        Me.fraUserInfo.Controls.Add(Me.chkExpired)
        Me.fraUserInfo.Controls.Add(Me.Label4)
        Me.fraUserInfo.Controls.Add(Me.txtEmail)
        Me.fraUserInfo.Controls.Add(Me.cboUserPIC)
        Me.fraUserInfo.Controls.Add(Me.Label5)
        Me.fraUserInfo.Location = New System.Drawing.Point(12, 491)
        Me.fraUserInfo.Name = "fraUserInfo"
        Me.fraUserInfo.Size = New System.Drawing.Size(760, 140)
        Me.fraUserInfo.TabIndex = 3
        Me.fraUserInfo.TabStop = False
        Me.fraUserInfo.Text = "User Information"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(24, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(46, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "User ID:"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(398, 104)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(45, 13)
        Me.Label8.TabIndex = 14
        Me.Label8.Text = "Expired:"
        '
        'txtUserID
        '
        Me.txtUserID.Location = New System.Drawing.Point(122, 23)
        Me.txtUserID.MaxLength = 10
        Me.txtUserID.Name = "txtUserID"
        Me.txtUserID.Size = New System.Drawing.Size(80, 20)
        Me.txtUserID.TabIndex = 1
        '
        'txtPwd2
        '
        Me.txtPwd2.Location = New System.Drawing.Point(122, 101)
        Me.txtPwd2.MaxLength = 50
        Me.txtPwd2.Name = "txtPwd2"
        Me.txtPwd2.PasswordChar = Global.Microsoft.VisualBasic.ChrW(9679)
        Me.txtPwd2.Size = New System.Drawing.Size(175, 20)
        Me.txtPwd2.TabIndex = 13
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(24, 52)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "User Name:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(24, 104)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(94, 13)
        Me.Label7.TabIndex = 12
        Me.Label7.Text = "Confirm Password:"
        '
        'txtUserName
        '
        Me.txtUserName.Location = New System.Drawing.Point(122, 49)
        Me.txtUserName.MaxLength = 100
        Me.txtUserName.Name = "txtUserName"
        Me.txtUserName.Size = New System.Drawing.Size(240, 20)
        Me.txtUserName.TabIndex = 5
        '
        'txtPwd1
        '
        Me.txtPwd1.Location = New System.Drawing.Point(122, 75)
        Me.txtPwd1.MaxLength = 50
        Me.txtPwd1.Name = "txtPwd1"
        Me.txtPwd1.PasswordChar = Global.Microsoft.VisualBasic.ChrW(9679)
        Me.txtPwd1.Size = New System.Drawing.Size(175, 20)
        Me.txtPwd1.TabIndex = 9
        '
        'cboUserLevel
        '
        Me.cboUserLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboUserLevel.FormattingEnabled = True
        Me.cboUserLevel.Location = New System.Drawing.Point(496, 49)
        Me.cboUserLevel.Name = "cboUserLevel"
        Me.cboUserLevel.Size = New System.Drawing.Size(240, 21)
        Me.cboUserLevel.TabIndex = 7
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(24, 78)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(56, 13)
        Me.Label6.TabIndex = 8
        Me.Label6.Text = "Password:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(398, 52)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(61, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "User Level:"
        '
        'chkExpired
        '
        Me.chkExpired.AutoSize = True
        Me.chkExpired.Location = New System.Drawing.Point(496, 104)
        Me.chkExpired.Name = "chkExpired"
        Me.chkExpired.Size = New System.Drawing.Size(15, 14)
        Me.chkExpired.TabIndex = 15
        Me.chkExpired.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(398, 78)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(92, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Person In Charge:"
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(496, 23)
        Me.txtEmail.MaxLength = 100
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(240, 20)
        Me.txtEmail.TabIndex = 3
        '
        'cboUserPIC
        '
        Me.cboUserPIC.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboUserPIC.FormattingEnabled = True
        Me.cboUserPIC.Location = New System.Drawing.Point(496, 75)
        Me.cboUserPIC.Name = "cboUserPIC"
        Me.cboUserPIC.Size = New System.Drawing.Size(240, 21)
        Me.cboUserPIC.TabIndex = 11
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(398, 26)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(76, 13)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "Email Address:"
        '
        'cmdExport
        '
        Me.cmdExport.Location = New System.Drawing.Point(265, 637)
        Me.cmdExport.Name = "cmdExport"
        Me.cmdExport.Size = New System.Drawing.Size(75, 23)
        Me.cmdExport.TabIndex = 7
        Me.cmdExport.Text = "&Export"
        Me.cmdExport.UseVisualStyleBackColor = True
        '
        'cmdImport
        '
        Me.cmdImport.Location = New System.Drawing.Point(184, 637)
        Me.cmdImport.Name = "cmdImport"
        Me.cmdImport.Size = New System.Drawing.Size(75, 23)
        Me.cmdImport.TabIndex = 6
        Me.cmdImport.Text = "&Import"
        Me.cmdImport.UseVisualStyleBackColor = True
        '
        'chkHideExpiredUser
        '
        Me.chkHideExpiredUser.AutoSize = True
        Me.chkHideExpiredUser.Checked = True
        Me.chkHideExpiredUser.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkHideExpiredUser.Location = New System.Drawing.Point(467, 47)
        Me.chkHideExpiredUser.Name = "chkHideExpiredUser"
        Me.chkHideExpiredUser.Size = New System.Drawing.Size(111, 17)
        Me.chkHideExpiredUser.TabIndex = 6
        Me.chkHideExpiredUser.Text = "Hide Expired User"
        Me.chkHideExpiredUser.UseVisualStyleBackColor = True
        '
        'fraFilter
        '
        Me.fraFilter.Controls.Add(Me.Label9)
        Me.fraFilter.Controls.Add(Me.chkHideExpiredUser)
        Me.fraFilter.Controls.Add(Me.txtUserIDFiter)
        Me.fraFilter.Controls.Add(Me.Label10)
        Me.fraFilter.Controls.Add(Me.txtUserNameFilter)
        Me.fraFilter.Controls.Add(Me.cboUserLevelFilter)
        Me.fraFilter.Controls.Add(Me.Label11)
        Me.fraFilter.Controls.Add(Me.Label12)
        Me.fraFilter.Controls.Add(Me.cboUserPICFilter)
        Me.fraFilter.Controls.Add(Me.cmdClearFilter)
        Me.fraFilter.Controls.Add(Me.cmdFilter)
        Me.fraFilter.Location = New System.Drawing.Point(12, 36)
        Me.fraFilter.Name = "fraFilter"
        Me.fraFilter.Size = New System.Drawing.Size(760, 107)
        Me.fraFilter.TabIndex = 1
        Me.fraFilter.TabStop = False
        Me.fraFilter.Text = "Filter Section"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(24, 22)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(46, 13)
        Me.Label9.TabIndex = 0
        Me.Label9.Text = "User ID:"
        '
        'txtUserIDFiter
        '
        Me.txtUserIDFiter.Location = New System.Drawing.Point(122, 19)
        Me.txtUserIDFiter.MaxLength = 10
        Me.txtUserIDFiter.Name = "txtUserIDFiter"
        Me.txtUserIDFiter.Size = New System.Drawing.Size(80, 20)
        Me.txtUserIDFiter.TabIndex = 1
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(398, 22)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(63, 13)
        Me.Label10.TabIndex = 2
        Me.Label10.Text = "User Name:"
        '
        'txtUserNameFilter
        '
        Me.txtUserNameFilter.Location = New System.Drawing.Point(467, 19)
        Me.txtUserNameFilter.MaxLength = 100
        Me.txtUserNameFilter.Name = "txtUserNameFilter"
        Me.txtUserNameFilter.Size = New System.Drawing.Size(240, 20)
        Me.txtUserNameFilter.TabIndex = 3
        '
        'cboUserLevelFilter
        '
        Me.cboUserLevelFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboUserLevelFilter.FormattingEnabled = True
        Me.cboUserLevelFilter.Location = New System.Drawing.Point(122, 45)
        Me.cboUserLevelFilter.Name = "cboUserLevelFilter"
        Me.cboUserLevelFilter.Size = New System.Drawing.Size(240, 21)
        Me.cboUserLevelFilter.TabIndex = 5
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(24, 48)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(61, 13)
        Me.Label11.TabIndex = 4
        Me.Label11.Text = "User Level:"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(24, 74)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(92, 13)
        Me.Label12.TabIndex = 7
        Me.Label12.Text = "Person In Charge:"
        '
        'cboUserPICFilter
        '
        Me.cboUserPICFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboUserPICFilter.FormattingEnabled = True
        Me.cboUserPICFilter.Location = New System.Drawing.Point(122, 71)
        Me.cboUserPICFilter.Name = "cboUserPICFilter"
        Me.cboUserPICFilter.Size = New System.Drawing.Size(240, 21)
        Me.cboUserPICFilter.TabIndex = 8
        '
        'cmdClearFilter
        '
        Me.cmdClearFilter.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdClearFilter.Location = New System.Drawing.Point(679, 69)
        Me.cmdClearFilter.Name = "cmdClearFilter"
        Me.cmdClearFilter.Size = New System.Drawing.Size(75, 23)
        Me.cmdClearFilter.TabIndex = 10
        Me.cmdClearFilter.Text = "Clear"
        Me.cmdClearFilter.UseVisualStyleBackColor = True
        '
        'cmdFilter
        '
        Me.cmdFilter.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdFilter.Location = New System.Drawing.Point(598, 69)
        Me.cmdFilter.Name = "cmdFilter"
        Me.cmdFilter.Size = New System.Drawing.Size(75, 23)
        Me.cmdFilter.TabIndex = 9
        Me.cmdFilter.Text = "&Filter"
        Me.cmdFilter.UseVisualStyleBackColor = True
        '
        'frmBG0610
        '
        Me.AcceptButton = Me.cmdFilter
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.CancelButton = Me.cmdClose
        Me.ClientSize = New System.Drawing.Size(809, 578)
        Me.Controls.Add(Me.fraFilter)
        Me.Controls.Add(Me.cmdExport)
        Me.Controls.Add(Me.cmdImport)
        Me.Controls.Add(Me.fraUserInfo)
        Me.Controls.Add(Me.cmdAdd)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.grvMaster)
        Me.Controls.Add(Me.lblFormTitle)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "frmBG0610"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "frmBG0610"
        CType(Me.grvMaster, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraUserInfo.ResumeLayout(False)
        Me.fraUserInfo.PerformLayout()
        Me.fraFilter.ResumeLayout(False)
        Me.fraFilter.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblFormTitle As System.Windows.Forms.Label
    Friend WithEvents grvMaster As System.Windows.Forms.DataGridView
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents fraUserInfo As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtUserID As System.Windows.Forms.TextBox
    Friend WithEvents txtPwd2 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtUserName As System.Windows.Forms.TextBox
    Friend WithEvents txtPwd1 As System.Windows.Forms.TextBox
    Friend WithEvents cboUserLevel As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents chkExpired As System.Windows.Forms.CheckBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents cboUserPIC As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cmdExport As System.Windows.Forms.Button
    Friend WithEvents cmdImport As System.Windows.Forms.Button
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents chkHideExpiredUser As System.Windows.Forms.CheckBox
    Friend WithEvents fraFilter As System.Windows.Forms.GroupBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtUserIDFiter As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtUserNameFilter As System.Windows.Forms.TextBox
    Friend WithEvents cboUserLevelFilter As System.Windows.Forms.ComboBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents cboUserPICFilter As System.Windows.Forms.ComboBox
    Friend WithEvents cmdClearFilter As System.Windows.Forms.Button
    Friend WithEvents cmdFilter As System.Windows.Forms.Button
End Class
