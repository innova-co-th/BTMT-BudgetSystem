<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBG0360
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBG0360))
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.lblFormTitle = New System.Windows.Forms.Label
        Me.cmdExport = New System.Windows.Forms.Button
        Me.cboAccountNo = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.numYear = New System.Windows.Forms.NumericUpDown
        Me.cboPeriodType = New System.Windows.Forms.ComboBox
        Me.lblConfirmPwd = New System.Windows.Forms.Label
        Me.lblYear = New System.Windows.Forms.Label
        Me.numProjectNo = New System.Windows.Forms.NumericUpDown
        Me.lblProjectNo = New System.Windows.Forms.Label
        CType(Me.numYear, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numProjectNo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(264, 150)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 6
        Me.cmdCancel.Text = "&Close"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'lblFormTitle
        '
        Me.lblFormTitle.AutoSize = True
        Me.lblFormTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblFormTitle.Location = New System.Drawing.Point(12, 9)
        Me.lblFormTitle.Name = "lblFormTitle"
        Me.lblFormTitle.Size = New System.Drawing.Size(189, 24)
        Me.lblFormTitle.TabIndex = 11
        Me.lblFormTitle.Text = "Export File To SAP"
        '
        'cmdExport
        '
        Me.cmdExport.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdExport.Location = New System.Drawing.Point(183, 150)
        Me.cmdExport.Name = "cmdExport"
        Me.cmdExport.Size = New System.Drawing.Size(75, 23)
        Me.cmdExport.TabIndex = 7
        Me.cmdExport.Text = "&Export"
        Me.cmdExport.UseVisualStyleBackColor = True
        '
        'cboAccountNo
        '
        Me.cboAccountNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAccountNo.FormattingEnabled = True
        Me.cboAccountNo.Location = New System.Drawing.Point(87, 123)
        Me.cboAccountNo.Name = "cboAccountNo"
        Me.cboAccountNo.Size = New System.Drawing.Size(250, 21)
        Me.cboAccountNo.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(14, 126)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(67, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "&Account No:"
        '
        'numYear
        '
        Me.numYear.Location = New System.Drawing.Point(88, 45)
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
        Me.cboPeriodType.Location = New System.Drawing.Point(88, 71)
        Me.cboPeriodType.Name = "cboPeriodType"
        Me.cboPeriodType.Size = New System.Drawing.Size(250, 21)
        Me.cboPeriodType.TabIndex = 3
        '
        'lblConfirmPwd
        '
        Me.lblConfirmPwd.AutoSize = True
        Me.lblConfirmPwd.Location = New System.Drawing.Point(14, 74)
        Me.lblConfirmPwd.Name = "lblConfirmPwd"
        Me.lblConfirmPwd.Size = New System.Drawing.Size(67, 13)
        Me.lblConfirmPwd.TabIndex = 2
        Me.lblConfirmPwd.Text = "Period &Type:"
        '
        'lblYear
        '
        Me.lblYear.AutoSize = True
        Me.lblYear.Location = New System.Drawing.Point(13, 47)
        Me.lblYear.Name = "lblYear"
        Me.lblYear.Size = New System.Drawing.Size(69, 13)
        Me.lblYear.TabIndex = 0
        Me.lblYear.Text = "Budget &Year:"
        '
        'numProjectNo
        '
        Me.numProjectNo.Location = New System.Drawing.Point(88, 97)
        Me.numProjectNo.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numProjectNo.Name = "numProjectNo"
        Me.numProjectNo.Size = New System.Drawing.Size(38, 20)
        Me.numProjectNo.TabIndex = 15
        Me.numProjectNo.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'lblProjectNo
        '
        Me.lblProjectNo.AutoSize = True
        Me.lblProjectNo.Location = New System.Drawing.Point(14, 99)
        Me.lblProjectNo.Name = "lblProjectNo"
        Me.lblProjectNo.Size = New System.Drawing.Size(60, 13)
        Me.lblProjectNo.TabIndex = 16
        Me.lblProjectNo.Text = "Project No:"
        '
        'frmBG0360
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(420, 192)
        Me.Controls.Add(Me.numProjectNo)
        Me.Controls.Add(Me.lblProjectNo)
        Me.Controls.Add(Me.cmdExport)
        Me.Controls.Add(Me.cboAccountNo)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.numYear)
        Me.Controls.Add(Me.cboPeriodType)
        Me.Controls.Add(Me.lblConfirmPwd)
        Me.Controls.Add(Me.lblYear)
        Me.Controls.Add(Me.lblFormTitle)
        Me.Controls.Add(Me.cmdCancel)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MinimumSize = New System.Drawing.Size(428, 200)
        Me.Name = "frmBG0360"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "frmBG0360"
        CType(Me.numYear, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numProjectNo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents lblFormTitle As System.Windows.Forms.Label
    Friend WithEvents cmdExport As System.Windows.Forms.Button
    Friend WithEvents cboAccountNo As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents numYear As System.Windows.Forms.NumericUpDown
    Friend WithEvents cboPeriodType As System.Windows.Forms.ComboBox
    Friend WithEvents lblConfirmPwd As System.Windows.Forms.Label
    Friend WithEvents lblYear As System.Windows.Forms.Label
    Friend WithEvents numProjectNo As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblProjectNo As System.Windows.Forms.Label
End Class
