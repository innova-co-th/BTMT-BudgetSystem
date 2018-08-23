<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBG0670
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBG0670))
        Me.lblFormTitle = New System.Windows.Forms.Label
        Me.cboPeriodType = New System.Windows.Forms.ComboBox
        Me.lblConfirmPwd = New System.Windows.Forms.Label
        Me.lblYear = New System.Windows.Forms.Label
        Me.fraInfo = New System.Windows.Forms.GroupBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtSecondHalfWBudget = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtFirstHalfWBudget = New System.Windows.Forms.TextBox
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.fraMTP = New System.Windows.Forms.GroupBox
        Me.lblRRT5 = New System.Windows.Forms.Label
        Me.lblRRT4 = New System.Windows.Forms.Label
        Me.lblRRT3 = New System.Windows.Forms.Label
        Me.lblRRT2 = New System.Windows.Forms.Label
        Me.lblRRT1 = New System.Windows.Forms.Label
        Me.txtRRT0 = New System.Windows.Forms.TextBox
        Me.lblRRT0 = New System.Windows.Forms.Label
        Me.txtRRT1 = New System.Windows.Forms.TextBox
        Me.txtRRT2 = New System.Windows.Forms.TextBox
        Me.txtRRT3 = New System.Windows.Forms.TextBox
        Me.txtRRT4 = New System.Windows.Forms.TextBox
        Me.txtRRT5 = New System.Windows.Forms.TextBox
        Me.lblRRT1p = New System.Windows.Forms.Label
        Me.lblRRT2p = New System.Windows.Forms.Label
        Me.lblRRT5p = New System.Windows.Forms.Label
        Me.lblRRT3p = New System.Windows.Forms.Label
        Me.lblRRT4p = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.cboRevNo = New System.Windows.Forms.ComboBox
        Me.cboBudgetYear = New System.Windows.Forms.ComboBox
        Me.lblProjectNo = New System.Windows.Forms.Label
        Me.cboProjectNo = New System.Windows.Forms.ComboBox
        Me.grbReference = New System.Windows.Forms.GroupBox
        Me.cboRefBudgetYear = New System.Windows.Forms.ComboBox
        Me.cboRefPeriodType = New System.Windows.Forms.ComboBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.cboRefProjectNo = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.cboRefRevNo = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.grbReference2 = New System.Windows.Forms.GroupBox
        Me.cboRefBudgetYear2 = New System.Windows.Forms.ComboBox
        Me.cboRefPeriodType2 = New System.Windows.Forms.ComboBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.cboRefProjectNo2 = New System.Windows.Forms.ComboBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.cboRefRevNo2 = New System.Windows.Forms.ComboBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.fraInfo.SuspendLayout()
        Me.fraMTP.SuspendLayout()
        Me.grbReference.SuspendLayout()
        Me.grbReference2.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblFormTitle
        '
        Me.lblFormTitle.AutoSize = True
        Me.lblFormTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblFormTitle.Location = New System.Drawing.Point(12, 9)
        Me.lblFormTitle.Name = "lblFormTitle"
        Me.lblFormTitle.Size = New System.Drawing.Size(208, 24)
        Me.lblFormTitle.TabIndex = 11
        Me.lblFormTitle.Text = "Budget Adjust Master"
        '
        'cboPeriodType
        '
        Me.cboPeriodType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPeriodType.FormattingEnabled = True
        Me.cboPeriodType.Location = New System.Drawing.Point(86, 71)
        Me.cboPeriodType.Name = "cboPeriodType"
        Me.cboPeriodType.Size = New System.Drawing.Size(160, 21)
        Me.cboPeriodType.TabIndex = 1
        '
        'lblConfirmPwd
        '
        Me.lblConfirmPwd.AutoSize = True
        Me.lblConfirmPwd.Location = New System.Drawing.Point(13, 74)
        Me.lblConfirmPwd.Name = "lblConfirmPwd"
        Me.lblConfirmPwd.Size = New System.Drawing.Size(67, 13)
        Me.lblConfirmPwd.TabIndex = 14
        Me.lblConfirmPwd.Text = "Period &Type:"
        '
        'lblYear
        '
        Me.lblYear.AutoSize = True
        Me.lblYear.Location = New System.Drawing.Point(12, 47)
        Me.lblYear.Name = "lblYear"
        Me.lblYear.Size = New System.Drawing.Size(69, 13)
        Me.lblYear.TabIndex = 12
        Me.lblYear.Text = "Budget &Year:"
        '
        'fraInfo
        '
        Me.fraInfo.Controls.Add(Me.Label7)
        Me.fraInfo.Controls.Add(Me.txtSecondHalfWBudget)
        Me.fraInfo.Controls.Add(Me.Label6)
        Me.fraInfo.Controls.Add(Me.Label3)
        Me.fraInfo.Controls.Add(Me.Label2)
        Me.fraInfo.Controls.Add(Me.txtFirstHalfWBudget)
        Me.fraInfo.Location = New System.Drawing.Point(12, 157)
        Me.fraInfo.Name = "fraInfo"
        Me.fraInfo.Size = New System.Drawing.Size(363, 86)
        Me.fraInfo.TabIndex = 62
        Me.fraInfo.TabStop = False
        Me.fraInfo.Text = "Working Budget"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(130, 52)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(15, 13)
        Me.Label7.TabIndex = 54
        Me.Label7.Text = "%"
        '
        'txtSecondHalfWBudget
        '
        Me.txtSecondHalfWBudget.Location = New System.Drawing.Point(74, 49)
        Me.txtSecondHalfWBudget.MaxLength = 9
        Me.txtSecondHalfWBudget.Name = "txtSecondHalfWBudget"
        Me.txtSecondHalfWBudget.Size = New System.Drawing.Size(53, 20)
        Me.txtSecondHalfWBudget.TabIndex = 4
        Me.txtSecondHalfWBudget.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(130, 26)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(15, 13)
        Me.Label6.TabIndex = 52
        Me.Label6.Text = "%"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(18, 52)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(50, 13)
        Me.Label3.TabIndex = 46
        Me.Label3.Text = "2nd Half:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(22, 26)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(46, 13)
        Me.Label2.TabIndex = 44
        Me.Label2.Text = "1st Half:"
        '
        'txtFirstHalfWBudget
        '
        Me.txtFirstHalfWBudget.Location = New System.Drawing.Point(74, 23)
        Me.txtFirstHalfWBudget.MaxLength = 9
        Me.txtFirstHalfWBudget.Name = "txtFirstHalfWBudget"
        Me.txtFirstHalfWBudget.Size = New System.Drawing.Size(53, 20)
        Me.txtFirstHalfWBudget.TabIndex = 3
        Me.txtFirstHalfWBudget.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(358, 490)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(75, 23)
        Me.cmdSave.TabIndex = 11
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'cmdClose
        '
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdClose.Location = New System.Drawing.Point(439, 490)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 23)
        Me.cmdClose.TabIndex = 12
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'fraMTP
        '
        Me.fraMTP.Controls.Add(Me.lblRRT5)
        Me.fraMTP.Controls.Add(Me.lblRRT4)
        Me.fraMTP.Controls.Add(Me.lblRRT3)
        Me.fraMTP.Controls.Add(Me.lblRRT2)
        Me.fraMTP.Controls.Add(Me.lblRRT1)
        Me.fraMTP.Controls.Add(Me.txtRRT0)
        Me.fraMTP.Controls.Add(Me.lblRRT0)
        Me.fraMTP.Controls.Add(Me.txtRRT1)
        Me.fraMTP.Controls.Add(Me.txtRRT2)
        Me.fraMTP.Controls.Add(Me.txtRRT3)
        Me.fraMTP.Controls.Add(Me.txtRRT4)
        Me.fraMTP.Controls.Add(Me.txtRRT5)
        Me.fraMTP.Controls.Add(Me.lblRRT1p)
        Me.fraMTP.Controls.Add(Me.lblRRT2p)
        Me.fraMTP.Controls.Add(Me.lblRRT5p)
        Me.fraMTP.Controls.Add(Me.lblRRT3p)
        Me.fraMTP.Controls.Add(Me.lblRRT4p)
        Me.fraMTP.Location = New System.Drawing.Point(12, 249)
        Me.fraMTP.Name = "fraMTP"
        Me.fraMTP.Size = New System.Drawing.Size(363, 93)
        Me.fraMTP.TabIndex = 63
        Me.fraMTP.TabStop = False
        Me.fraMTP.Text = "MTP Budget"
        '
        'lblRRT5
        '
        Me.lblRRT5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblRRT5.Location = New System.Drawing.Point(296, 20)
        Me.lblRRT5.Name = "lblRRT5"
        Me.lblRRT5.Size = New System.Drawing.Size(55, 13)
        Me.lblRRT5.TabIndex = 56
        Me.lblRRT5.Text = "Year N+5"
        Me.lblRRT5.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'lblRRT4
        '
        Me.lblRRT4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblRRT4.Location = New System.Drawing.Point(240, 20)
        Me.lblRRT4.Name = "lblRRT4"
        Me.lblRRT4.Size = New System.Drawing.Size(55, 13)
        Me.lblRRT4.TabIndex = 55
        Me.lblRRT4.Text = "Year N+4"
        Me.lblRRT4.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'lblRRT3
        '
        Me.lblRRT3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblRRT3.Location = New System.Drawing.Point(184, 20)
        Me.lblRRT3.Name = "lblRRT3"
        Me.lblRRT3.Size = New System.Drawing.Size(55, 13)
        Me.lblRRT3.TabIndex = 54
        Me.lblRRT3.Text = "Year N+3"
        Me.lblRRT3.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'lblRRT2
        '
        Me.lblRRT2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblRRT2.Location = New System.Drawing.Point(128, 20)
        Me.lblRRT2.Name = "lblRRT2"
        Me.lblRRT2.Size = New System.Drawing.Size(55, 13)
        Me.lblRRT2.TabIndex = 53
        Me.lblRRT2.Text = "Year N+2"
        Me.lblRRT2.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'lblRRT1
        '
        Me.lblRRT1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblRRT1.Location = New System.Drawing.Point(72, 20)
        Me.lblRRT1.Name = "lblRRT1"
        Me.lblRRT1.Size = New System.Drawing.Size(55, 13)
        Me.lblRRT1.TabIndex = 52
        Me.lblRRT1.Text = "Year N+1"
        Me.lblRRT1.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'txtRRT0
        '
        Me.txtRRT0.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtRRT0.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.txtRRT0.Location = New System.Drawing.Point(16, 34)
        Me.txtRRT0.MaxLength = 10
        Me.txtRRT0.Name = "txtRRT0"
        Me.txtRRT0.Size = New System.Drawing.Size(50, 20)
        Me.txtRRT0.TabIndex = 5
        Me.txtRRT0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblRRT0
        '
        Me.lblRRT0.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblRRT0.Location = New System.Drawing.Point(16, 20)
        Me.lblRRT0.Name = "lblRRT0"
        Me.lblRRT0.Size = New System.Drawing.Size(55, 13)
        Me.lblRRT0.TabIndex = 51
        Me.lblRRT0.Text = "Year N"
        Me.lblRRT0.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'txtRRT1
        '
        Me.txtRRT1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtRRT1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.txtRRT1.Location = New System.Drawing.Point(72, 34)
        Me.txtRRT1.MaxLength = 10
        Me.txtRRT1.Name = "txtRRT1"
        Me.txtRRT1.Size = New System.Drawing.Size(50, 20)
        Me.txtRRT1.TabIndex = 6
        Me.txtRRT1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtRRT2
        '
        Me.txtRRT2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtRRT2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.txtRRT2.Location = New System.Drawing.Point(128, 34)
        Me.txtRRT2.MaxLength = 10
        Me.txtRRT2.Name = "txtRRT2"
        Me.txtRRT2.Size = New System.Drawing.Size(50, 20)
        Me.txtRRT2.TabIndex = 7
        Me.txtRRT2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtRRT3
        '
        Me.txtRRT3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtRRT3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.txtRRT3.Location = New System.Drawing.Point(184, 34)
        Me.txtRRT3.MaxLength = 10
        Me.txtRRT3.Name = "txtRRT3"
        Me.txtRRT3.Size = New System.Drawing.Size(50, 20)
        Me.txtRRT3.TabIndex = 8
        Me.txtRRT3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtRRT4
        '
        Me.txtRRT4.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtRRT4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.txtRRT4.Location = New System.Drawing.Point(240, 34)
        Me.txtRRT4.MaxLength = 10
        Me.txtRRT4.Name = "txtRRT4"
        Me.txtRRT4.Size = New System.Drawing.Size(50, 20)
        Me.txtRRT4.TabIndex = 9
        Me.txtRRT4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtRRT5
        '
        Me.txtRRT5.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtRRT5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.txtRRT5.Location = New System.Drawing.Point(296, 34)
        Me.txtRRT5.MaxLength = 10
        Me.txtRRT5.Name = "txtRRT5"
        Me.txtRRT5.Size = New System.Drawing.Size(50, 20)
        Me.txtRRT5.TabIndex = 10
        Me.txtRRT5.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblRRT1p
        '
        Me.lblRRT1p.BackColor = System.Drawing.SystemColors.Control
        Me.lblRRT1p.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblRRT1p.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblRRT1p.Location = New System.Drawing.Point(72, 57)
        Me.lblRRT1p.Name = "lblRRT1p"
        Me.lblRRT1p.Size = New System.Drawing.Size(50, 20)
        Me.lblRRT1p.TabIndex = 45
        Me.lblRRT1p.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblRRT1p.Visible = False
        '
        'lblRRT2p
        '
        Me.lblRRT2p.BackColor = System.Drawing.SystemColors.Control
        Me.lblRRT2p.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblRRT2p.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblRRT2p.Location = New System.Drawing.Point(128, 57)
        Me.lblRRT2p.Name = "lblRRT2p"
        Me.lblRRT2p.Size = New System.Drawing.Size(50, 20)
        Me.lblRRT2p.TabIndex = 46
        Me.lblRRT2p.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblRRT5p
        '
        Me.lblRRT5p.BackColor = System.Drawing.SystemColors.Control
        Me.lblRRT5p.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblRRT5p.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblRRT5p.Location = New System.Drawing.Point(296, 57)
        Me.lblRRT5p.Name = "lblRRT5p"
        Me.lblRRT5p.Size = New System.Drawing.Size(50, 20)
        Me.lblRRT5p.TabIndex = 49
        Me.lblRRT5p.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblRRT3p
        '
        Me.lblRRT3p.BackColor = System.Drawing.SystemColors.Control
        Me.lblRRT3p.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblRRT3p.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblRRT3p.Location = New System.Drawing.Point(184, 57)
        Me.lblRRT3p.Name = "lblRRT3p"
        Me.lblRRT3p.Size = New System.Drawing.Size(50, 20)
        Me.lblRRT3p.TabIndex = 47
        Me.lblRRT3p.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblRRT4p
        '
        Me.lblRRT4p.BackColor = System.Drawing.SystemColors.Control
        Me.lblRRT4p.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblRRT4p.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblRRT4p.Location = New System.Drawing.Point(240, 57)
        Me.lblRRT4p.Name = "lblRRT4p"
        Me.lblRRT4p.Size = New System.Drawing.Size(50, 20)
        Me.lblRRT4p.TabIndex = 48
        Me.lblRRT4p.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 128)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(53, 13)
        Me.Label1.TabIndex = 64
        Me.Label1.Text = "Rev. No.:"
        '
        'cboRevNo
        '
        Me.cboRevNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRevNo.FormattingEnabled = True
        Me.cboRevNo.Location = New System.Drawing.Point(86, 125)
        Me.cboRevNo.Name = "cboRevNo"
        Me.cboRevNo.Size = New System.Drawing.Size(53, 21)
        Me.cboRevNo.TabIndex = 2
        '
        'cboBudgetYear
        '
        Me.cboBudgetYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBudgetYear.FormattingEnabled = True
        Me.cboBudgetYear.Location = New System.Drawing.Point(86, 44)
        Me.cboBudgetYear.Name = "cboBudgetYear"
        Me.cboBudgetYear.Size = New System.Drawing.Size(53, 21)
        Me.cboBudgetYear.TabIndex = 0
        '
        'lblProjectNo
        '
        Me.lblProjectNo.AutoSize = True
        Me.lblProjectNo.Location = New System.Drawing.Point(13, 101)
        Me.lblProjectNo.Name = "lblProjectNo"
        Me.lblProjectNo.Size = New System.Drawing.Size(60, 13)
        Me.lblProjectNo.TabIndex = 66
        Me.lblProjectNo.Text = "Project No:"
        '
        'cboProjectNo
        '
        Me.cboProjectNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboProjectNo.FormattingEnabled = True
        Me.cboProjectNo.Location = New System.Drawing.Point(86, 98)
        Me.cboProjectNo.Name = "cboProjectNo"
        Me.cboProjectNo.Size = New System.Drawing.Size(53, 21)
        Me.cboProjectNo.TabIndex = 67
        '
        'grbReference
        '
        Me.grbReference.Controls.Add(Me.cboRefBudgetYear)
        Me.grbReference.Controls.Add(Me.cboRefPeriodType)
        Me.grbReference.Controls.Add(Me.Label8)
        Me.grbReference.Controls.Add(Me.Label9)
        Me.grbReference.Controls.Add(Me.cboRefProjectNo)
        Me.grbReference.Controls.Add(Me.Label4)
        Me.grbReference.Controls.Add(Me.cboRefRevNo)
        Me.grbReference.Controls.Add(Me.Label5)
        Me.grbReference.Location = New System.Drawing.Point(12, 348)
        Me.grbReference.Name = "grbReference"
        Me.grbReference.Size = New System.Drawing.Size(248, 136)
        Me.grbReference.TabIndex = 68
        Me.grbReference.TabStop = False
        Me.grbReference.Text = "Reference Budget"
        '
        'cboRefBudgetYear
        '
        Me.cboRefBudgetYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRefBudgetYear.Enabled = False
        Me.cboRefBudgetYear.FormattingEnabled = True
        Me.cboRefBudgetYear.Location = New System.Drawing.Point(74, 23)
        Me.cboRefBudgetYear.Name = "cboRefBudgetYear"
        Me.cboRefBudgetYear.Size = New System.Drawing.Size(53, 21)
        Me.cboRefBudgetYear.TabIndex = 72
        '
        'cboRefPeriodType
        '
        Me.cboRefPeriodType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRefPeriodType.Enabled = False
        Me.cboRefPeriodType.FormattingEnabled = True
        Me.cboRefPeriodType.Location = New System.Drawing.Point(74, 50)
        Me.cboRefPeriodType.Name = "cboRefPeriodType"
        Me.cboRefPeriodType.Size = New System.Drawing.Size(160, 21)
        Me.cboRefPeriodType.TabIndex = 73
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(5, 53)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(67, 13)
        Me.Label8.TabIndex = 75
        Me.Label8.Text = "Period &Type:"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(5, 26)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(69, 13)
        Me.Label9.TabIndex = 74
        Me.Label9.Text = "Budget &Year:"
        '
        'cboRefProjectNo
        '
        Me.cboRefProjectNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRefProjectNo.FormattingEnabled = True
        Me.cboRefProjectNo.Location = New System.Drawing.Point(74, 77)
        Me.cboRefProjectNo.Name = "cboRefProjectNo"
        Me.cboRefProjectNo.Size = New System.Drawing.Size(53, 21)
        Me.cboRefProjectNo.TabIndex = 71
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label4.Location = New System.Drawing.Point(5, 80)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(60, 13)
        Me.Label4.TabIndex = 70
        Me.Label4.Text = "Project No:"
        '
        'cboRefRevNo
        '
        Me.cboRefRevNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRefRevNo.FormattingEnabled = True
        Me.cboRefRevNo.Location = New System.Drawing.Point(74, 104)
        Me.cboRefRevNo.Name = "cboRefRevNo"
        Me.cboRefRevNo.Size = New System.Drawing.Size(53, 21)
        Me.cboRefRevNo.TabIndex = 68
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label5.Location = New System.Drawing.Point(5, 107)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(53, 13)
        Me.Label5.TabIndex = 69
        Me.Label5.Text = "Rev. No.:"
        '
        'grbReference2
        '
        Me.grbReference2.Controls.Add(Me.cboRefBudgetYear2)
        Me.grbReference2.Controls.Add(Me.cboRefPeriodType2)
        Me.grbReference2.Controls.Add(Me.Label10)
        Me.grbReference2.Controls.Add(Me.Label11)
        Me.grbReference2.Controls.Add(Me.cboRefProjectNo2)
        Me.grbReference2.Controls.Add(Me.Label12)
        Me.grbReference2.Controls.Add(Me.cboRefRevNo2)
        Me.grbReference2.Controls.Add(Me.Label13)
        Me.grbReference2.Location = New System.Drawing.Point(266, 348)
        Me.grbReference2.Name = "grbReference2"
        Me.grbReference2.Size = New System.Drawing.Size(248, 136)
        Me.grbReference2.TabIndex = 69
        Me.grbReference2.TabStop = False
        Me.grbReference2.Text = "Reference Budget 2"
        '
        'cboRefBudgetYear2
        '
        Me.cboRefBudgetYear2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRefBudgetYear2.Enabled = False
        Me.cboRefBudgetYear2.FormattingEnabled = True
        Me.cboRefBudgetYear2.Location = New System.Drawing.Point(74, 23)
        Me.cboRefBudgetYear2.Name = "cboRefBudgetYear2"
        Me.cboRefBudgetYear2.Size = New System.Drawing.Size(53, 21)
        Me.cboRefBudgetYear2.TabIndex = 72
        '
        'cboRefPeriodType2
        '
        Me.cboRefPeriodType2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRefPeriodType2.Enabled = False
        Me.cboRefPeriodType2.FormattingEnabled = True
        Me.cboRefPeriodType2.Location = New System.Drawing.Point(74, 50)
        Me.cboRefPeriodType2.Name = "cboRefPeriodType2"
        Me.cboRefPeriodType2.Size = New System.Drawing.Size(160, 21)
        Me.cboRefPeriodType2.TabIndex = 73
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(5, 53)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(67, 13)
        Me.Label10.TabIndex = 75
        Me.Label10.Text = "Period &Type:"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(5, 26)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(69, 13)
        Me.Label11.TabIndex = 74
        Me.Label11.Text = "Budget &Year:"
        '
        'cboRefProjectNo2
        '
        Me.cboRefProjectNo2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRefProjectNo2.FormattingEnabled = True
        Me.cboRefProjectNo2.Location = New System.Drawing.Point(74, 77)
        Me.cboRefProjectNo2.Name = "cboRefProjectNo2"
        Me.cboRefProjectNo2.Size = New System.Drawing.Size(53, 21)
        Me.cboRefProjectNo2.TabIndex = 71
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label12.Location = New System.Drawing.Point(5, 80)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(60, 13)
        Me.Label12.TabIndex = 70
        Me.Label12.Text = "Project No:"
        '
        'cboRefRevNo2
        '
        Me.cboRefRevNo2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRefRevNo2.FormattingEnabled = True
        Me.cboRefRevNo2.Location = New System.Drawing.Point(74, 104)
        Me.cboRefRevNo2.Name = "cboRefRevNo2"
        Me.cboRefRevNo2.Size = New System.Drawing.Size(53, 21)
        Me.cboRefRevNo2.TabIndex = 68
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label13.Location = New System.Drawing.Point(5, 107)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(53, 13)
        Me.Label13.TabIndex = 69
        Me.Label13.Text = "Rev. No.:"
        '
        'frmBG0670
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(526, 521)
        Me.Controls.Add(Me.grbReference2)
        Me.Controls.Add(Me.grbReference)
        Me.Controls.Add(Me.cboProjectNo)
        Me.Controls.Add(Me.lblProjectNo)
        Me.Controls.Add(Me.cboBudgetYear)
        Me.Controls.Add(Me.cboRevNo)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.fraMTP)
        Me.Controls.Add(Me.fraInfo)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.cboPeriodType)
        Me.Controls.Add(Me.lblConfirmPwd)
        Me.Controls.Add(Me.lblYear)
        Me.Controls.Add(Me.lblFormTitle)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "frmBG0670"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "frmBG0670"
        Me.fraInfo.ResumeLayout(False)
        Me.fraInfo.PerformLayout()
        Me.fraMTP.ResumeLayout(False)
        Me.fraMTP.PerformLayout()
        Me.grbReference.ResumeLayout(False)
        Me.grbReference.PerformLayout()
        Me.grbReference2.ResumeLayout(False)
        Me.grbReference2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblFormTitle As System.Windows.Forms.Label
    Friend WithEvents cboPeriodType As System.Windows.Forms.ComboBox
    Friend WithEvents lblConfirmPwd As System.Windows.Forms.Label
    Friend WithEvents lblYear As System.Windows.Forms.Label
    Friend WithEvents fraInfo As System.Windows.Forms.GroupBox
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents txtFirstHalfWBudget As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtSecondHalfWBudget As System.Windows.Forms.TextBox
    Friend WithEvents fraMTP As System.Windows.Forms.GroupBox
    Friend WithEvents lblRRT5 As System.Windows.Forms.Label
    Friend WithEvents lblRRT4 As System.Windows.Forms.Label
    Friend WithEvents lblRRT3 As System.Windows.Forms.Label
    Friend WithEvents lblRRT2 As System.Windows.Forms.Label
    Friend WithEvents lblRRT1 As System.Windows.Forms.Label
    Friend WithEvents txtRRT0 As System.Windows.Forms.TextBox
    Friend WithEvents lblRRT0 As System.Windows.Forms.Label
    Friend WithEvents txtRRT1 As System.Windows.Forms.TextBox
    Friend WithEvents txtRRT2 As System.Windows.Forms.TextBox
    Friend WithEvents txtRRT3 As System.Windows.Forms.TextBox
    Friend WithEvents txtRRT4 As System.Windows.Forms.TextBox
    Friend WithEvents txtRRT5 As System.Windows.Forms.TextBox
    Friend WithEvents lblRRT1p As System.Windows.Forms.Label
    Friend WithEvents lblRRT2p As System.Windows.Forms.Label
    Friend WithEvents lblRRT5p As System.Windows.Forms.Label
    Friend WithEvents lblRRT3p As System.Windows.Forms.Label
    Friend WithEvents lblRRT4p As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboRevNo As System.Windows.Forms.ComboBox
    Friend WithEvents cboBudgetYear As System.Windows.Forms.ComboBox
    Friend WithEvents lblProjectNo As System.Windows.Forms.Label
    Friend WithEvents cboProjectNo As System.Windows.Forms.ComboBox
    Friend WithEvents grbReference As System.Windows.Forms.GroupBox
    Friend WithEvents cboRefProjectNo As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboRefRevNo As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboRefBudgetYear As System.Windows.Forms.ComboBox
    Friend WithEvents cboRefPeriodType As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents grbReference2 As System.Windows.Forms.GroupBox
    Friend WithEvents cboRefBudgetYear2 As System.Windows.Forms.ComboBox
    Friend WithEvents cboRefPeriodType2 As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents cboRefProjectNo2 As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents cboRefRevNo2 As System.Windows.Forms.ComboBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
End Class
