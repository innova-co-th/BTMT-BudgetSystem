<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBG0460
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBG0460))
        Me.lblFormTitle = New System.Windows.Forms.Label
        Me.cmdPrint = New System.Windows.Forms.Button
        Me.numYear = New System.Windows.Forms.NumericUpDown
        Me.cboPeriodType = New System.Windows.Forms.ComboBox
        Me.cmdPreview = New System.Windows.Forms.Button
        Me.lblConfirmPwd = New System.Windows.Forms.Label
        Me.lblYear = New System.Windows.Forms.Label
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument
        Me.cmdClose = New System.Windows.Forms.Button
        Me.chkShowMTP = New System.Windows.Forms.CheckBox
        Me.cmdExcel = New System.Windows.Forms.Button
        Me.numProjectNo = New System.Windows.Forms.NumericUpDown
        Me.lblProjectNo = New System.Windows.Forms.Label
        Me.cboRevNo = New System.Windows.Forms.ComboBox
        Me.lblRevNo = New System.Windows.Forms.Label
        Me.cboUserPIC = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
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
        Me.lblFormTitle.Size = New System.Drawing.Size(300, 24)
        Me.lblFormTitle.TabIndex = 0
        Me.lblFormTitle.Text = "Summary By Investment Report"
        '
        'cmdPrint
        '
        Me.cmdPrint.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdPrint.Location = New System.Drawing.Point(16, 154)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(75, 23)
        Me.cmdPrint.TabIndex = 12
        Me.cmdPrint.Text = "&Print"
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'numYear
        '
        Me.numYear.Location = New System.Drawing.Point(105, 45)
        Me.numYear.Maximum = New Decimal(New Integer() {3000, 0, 0, 0})
        Me.numYear.Minimum = New Decimal(New Integer() {2000, 0, 0, 0})
        Me.numYear.Name = "numYear"
        Me.numYear.Size = New System.Drawing.Size(49, 20)
        Me.numYear.TabIndex = 2
        Me.numYear.Value = New Decimal(New Integer() {2010, 0, 0, 0})
        '
        'cboPeriodType
        '
        Me.cboPeriodType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPeriodType.FormattingEnabled = True
        Me.cboPeriodType.Location = New System.Drawing.Point(105, 71)
        Me.cboPeriodType.Name = "cboPeriodType"
        Me.cboPeriodType.Size = New System.Drawing.Size(178, 21)
        Me.cboPeriodType.TabIndex = 4
        '
        'cmdPreview
        '
        Me.cmdPreview.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdPreview.Location = New System.Drawing.Point(97, 154)
        Me.cmdPreview.Name = "cmdPreview"
        Me.cmdPreview.Size = New System.Drawing.Size(75, 23)
        Me.cmdPreview.TabIndex = 13
        Me.cmdPreview.Text = "Pre&view"
        Me.cmdPreview.UseVisualStyleBackColor = True
        '
        'lblConfirmPwd
        '
        Me.lblConfirmPwd.AutoSize = True
        Me.lblConfirmPwd.Location = New System.Drawing.Point(13, 74)
        Me.lblConfirmPwd.Name = "lblConfirmPwd"
        Me.lblConfirmPwd.Size = New System.Drawing.Size(67, 13)
        Me.lblConfirmPwd.TabIndex = 3
        Me.lblConfirmPwd.Text = "Period &Type:"
        '
        'lblYear
        '
        Me.lblYear.AutoSize = True
        Me.lblYear.Location = New System.Drawing.Point(13, 47)
        Me.lblYear.Name = "lblYear"
        Me.lblYear.Size = New System.Drawing.Size(69, 13)
        Me.lblYear.TabIndex = 1
        Me.lblYear.Text = "Budget &Year:"
        '
        'PrintDialog1
        '
        Me.PrintDialog1.AllowSelection = True
        Me.PrintDialog1.UseEXDialog = True
        '
        'cmdClose
        '
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdClose.Location = New System.Drawing.Point(263, 154)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 23)
        Me.cmdClose.TabIndex = 15
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'chkShowMTP
        '
        Me.chkShowMTP.AutoSize = True
        Me.chkShowMTP.Enabled = False
        Me.chkShowMTP.Location = New System.Drawing.Point(289, 73)
        Me.chkShowMTP.Name = "chkShowMTP"
        Me.chkShowMTP.Size = New System.Drawing.Size(90, 17)
        Me.chkShowMTP.TabIndex = 5
        Me.chkShowMTP.Text = "2nd Half Only"
        Me.chkShowMTP.UseVisualStyleBackColor = True
        '
        'cmdExcel
        '
        Me.cmdExcel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdExcel.Location = New System.Drawing.Point(178, 154)
        Me.cmdExcel.Name = "cmdExcel"
        Me.cmdExcel.Size = New System.Drawing.Size(75, 23)
        Me.cmdExcel.TabIndex = 14
        Me.cmdExcel.Text = "&Excel"
        Me.cmdExcel.UseVisualStyleBackColor = True
        '
        'numProjectNo
        '
        Me.numProjectNo.Location = New System.Drawing.Point(105, 98)
        Me.numProjectNo.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numProjectNo.Name = "numProjectNo"
        Me.numProjectNo.Size = New System.Drawing.Size(38, 20)
        Me.numProjectNo.TabIndex = 7
        Me.numProjectNo.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'lblProjectNo
        '
        Me.lblProjectNo.AutoSize = True
        Me.lblProjectNo.Location = New System.Drawing.Point(13, 100)
        Me.lblProjectNo.Name = "lblProjectNo"
        Me.lblProjectNo.Size = New System.Drawing.Size(60, 13)
        Me.lblProjectNo.TabIndex = 6
        Me.lblProjectNo.Text = "Project No:"
        '
        'cboRevNo
        '
        Me.cboRevNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRevNo.FormattingEnabled = True
        Me.cboRevNo.Location = New System.Drawing.Point(230, 97)
        Me.cboRevNo.Name = "cboRevNo"
        Me.cboRevNo.Size = New System.Drawing.Size(53, 21)
        Me.cboRevNo.TabIndex = 9
        Me.cboRevNo.Visible = False
        '
        'lblRevNo
        '
        Me.lblRevNo.AutoSize = True
        Me.lblRevNo.Location = New System.Drawing.Point(171, 100)
        Me.lblRevNo.Name = "lblRevNo"
        Me.lblRevNo.Size = New System.Drawing.Size(53, 13)
        Me.lblRevNo.TabIndex = 8
        Me.lblRevNo.Text = "Rev. No.:"
        Me.lblRevNo.Visible = False
        '
        'cboUserPIC
        '
        Me.cboUserPIC.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboUserPIC.FormattingEnabled = True
        Me.cboUserPIC.Location = New System.Drawing.Point(105, 123)
        Me.cboUserPIC.Name = "cboUserPIC"
        Me.cboUserPIC.Size = New System.Drawing.Size(250, 21)
        Me.cboUserPIC.TabIndex = 11
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 126)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(92, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Person In &Charge:"
        '
        'frmBG0460
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(420, 188)
        Me.Controls.Add(Me.cboUserPIC)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboRevNo)
        Me.Controls.Add(Me.lblRevNo)
        Me.Controls.Add(Me.numProjectNo)
        Me.Controls.Add(Me.lblProjectNo)
        Me.Controls.Add(Me.cmdExcel)
        Me.Controls.Add(Me.chkShowMTP)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.cmdPrint)
        Me.Controls.Add(Me.numYear)
        Me.Controls.Add(Me.cboPeriodType)
        Me.Controls.Add(Me.cmdPreview)
        Me.Controls.Add(Me.lblConfirmPwd)
        Me.Controls.Add(Me.lblYear)
        Me.Controls.Add(Me.lblFormTitle)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MinimumSize = New System.Drawing.Size(428, 200)
        Me.Name = "frmBG0460"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "frmBG0460"
        CType(Me.numYear, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numProjectNo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblFormTitle As System.Windows.Forms.Label
    Friend WithEvents cmdPrint As System.Windows.Forms.Button
    Friend WithEvents numYear As System.Windows.Forms.NumericUpDown
    Friend WithEvents cboPeriodType As System.Windows.Forms.ComboBox
    Friend WithEvents cmdPreview As System.Windows.Forms.Button
    Friend WithEvents lblConfirmPwd As System.Windows.Forms.Label
    Friend WithEvents lblYear As System.Windows.Forms.Label
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    Friend WithEvents PrintDocument1 As System.Drawing.Printing.PrintDocument
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents chkShowMTP As System.Windows.Forms.CheckBox
    Friend WithEvents cmdExcel As System.Windows.Forms.Button
    Friend WithEvents numProjectNo As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblProjectNo As System.Windows.Forms.Label
    Friend WithEvents cboRevNo As System.Windows.Forms.ComboBox
    Friend WithEvents lblRevNo As System.Windows.Forms.Label
    Friend WithEvents cboUserPIC As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
