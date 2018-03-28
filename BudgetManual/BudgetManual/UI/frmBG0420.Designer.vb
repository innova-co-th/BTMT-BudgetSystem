<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBG0420
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBG0420))
        Me.lblFormTitle = New System.Windows.Forms.Label
        Me.cmdPrint = New System.Windows.Forms.Button
        Me.numYear = New System.Windows.Forms.NumericUpDown
        Me.cmdPreview = New System.Windows.Forms.Button
        Me.lblYear = New System.Windows.Forms.Label
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument
        Me.cmdClose = New System.Windows.Forms.Button
        Me.cboPeriodType = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.chkHideEstimate = New System.Windows.Forms.CheckBox
        Me.cmdExcel = New System.Windows.Forms.Button
        Me.numProjectNo = New System.Windows.Forms.NumericUpDown
        Me.lblProjectNo = New System.Windows.Forms.Label
        Me.cboRevNo = New System.Windows.Forms.ComboBox
        Me.lblRevNo = New System.Windows.Forms.Label
        Me.gbPrevYear = New System.Windows.Forms.GroupBox
        Me.cboPrevRevno = New System.Windows.Forms.ComboBox
        Me.lblPrevRevNo = New System.Windows.Forms.Label
        Me.numPrevProjectNo = New System.Windows.Forms.NumericUpDown
        Me.Label3 = New System.Windows.Forms.Label
        CType(Me.numYear, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numProjectNo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbPrevYear.SuspendLayout()
        CType(Me.numPrevProjectNo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblFormTitle
        '
        Me.lblFormTitle.AutoSize = True
        Me.lblFormTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblFormTitle.Location = New System.Drawing.Point(12, 9)
        Me.lblFormTitle.Name = "lblFormTitle"
        Me.lblFormTitle.Size = New System.Drawing.Size(363, 24)
        Me.lblFormTitle.TabIndex = 0
        Me.lblFormTitle.Text = "Summary By Person In Charge Report"
        '
        'cmdPrint
        '
        Me.cmdPrint.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdPrint.Location = New System.Drawing.Point(16, 129)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(75, 23)
        Me.cmdPrint.TabIndex = 11
        Me.cmdPrint.Text = "&Print"
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'numYear
        '
        Me.numYear.Location = New System.Drawing.Point(88, 45)
        Me.numYear.Maximum = New Decimal(New Integer() {3000, 0, 0, 0})
        Me.numYear.Minimum = New Decimal(New Integer() {2000, 0, 0, 0})
        Me.numYear.Name = "numYear"
        Me.numYear.Size = New System.Drawing.Size(49, 20)
        Me.numYear.TabIndex = 2
        Me.numYear.Value = New Decimal(New Integer() {2010, 0, 0, 0})
        '
        'cmdPreview
        '
        Me.cmdPreview.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdPreview.Location = New System.Drawing.Point(97, 129)
        Me.cmdPreview.Name = "cmdPreview"
        Me.cmdPreview.Size = New System.Drawing.Size(75, 23)
        Me.cmdPreview.TabIndex = 12
        Me.cmdPreview.Text = "Pre&view"
        Me.cmdPreview.UseVisualStyleBackColor = True
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
        Me.PrintDialog1.UseEXDialog = True
        '
        'cmdClose
        '
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdClose.Location = New System.Drawing.Point(263, 129)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 23)
        Me.cmdClose.TabIndex = 14
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'cboPeriodType
        '
        Me.cboPeriodType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPeriodType.FormattingEnabled = True
        Me.cboPeriodType.Location = New System.Drawing.Point(88, 71)
        Me.cboPeriodType.Name = "cboPeriodType"
        Me.cboPeriodType.Size = New System.Drawing.Size(183, 21)
        Me.cboPeriodType.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 74)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(67, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Period &Type:"
        '
        'chkHideEstimate
        '
        Me.chkHideEstimate.AutoSize = True
        Me.chkHideEstimate.Enabled = False
        Me.chkHideEstimate.Location = New System.Drawing.Point(277, 73)
        Me.chkHideEstimate.Name = "chkHideEstimate"
        Me.chkHideEstimate.Size = New System.Drawing.Size(90, 17)
        Me.chkHideEstimate.TabIndex = 5
        Me.chkHideEstimate.Text = "2nd Half Only"
        Me.chkHideEstimate.UseVisualStyleBackColor = True
        '
        'cmdExcel
        '
        Me.cmdExcel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdExcel.Location = New System.Drawing.Point(178, 129)
        Me.cmdExcel.Name = "cmdExcel"
        Me.cmdExcel.Size = New System.Drawing.Size(75, 23)
        Me.cmdExcel.TabIndex = 13
        Me.cmdExcel.Text = "&Excel"
        Me.cmdExcel.UseVisualStyleBackColor = True
        '
        'numProjectNo
        '
        Me.numProjectNo.Location = New System.Drawing.Point(88, 98)
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
        Me.cboRevNo.Location = New System.Drawing.Point(218, 97)
        Me.cboRevNo.Name = "cboRevNo"
        Me.cboRevNo.Size = New System.Drawing.Size(53, 21)
        Me.cboRevNo.TabIndex = 9
        Me.cboRevNo.Visible = False
        '
        'lblRevNo
        '
        Me.lblRevNo.AutoSize = True
        Me.lblRevNo.Location = New System.Drawing.Point(159, 100)
        Me.lblRevNo.Name = "lblRevNo"
        Me.lblRevNo.Size = New System.Drawing.Size(53, 13)
        Me.lblRevNo.TabIndex = 8
        Me.lblRevNo.Text = "Rev. No.:"
        Me.lblRevNo.Visible = False
        '
        'gbPrevYear
        '
        Me.gbPrevYear.Controls.Add(Me.cboPrevRevno)
        Me.gbPrevYear.Controls.Add(Me.lblPrevRevNo)
        Me.gbPrevYear.Controls.Add(Me.numPrevProjectNo)
        Me.gbPrevYear.Controls.Add(Me.Label3)
        Me.gbPrevYear.Enabled = False
        Me.gbPrevYear.Location = New System.Drawing.Point(16, 124)
        Me.gbPrevYear.Name = "gbPrevYear"
        Me.gbPrevYear.Size = New System.Drawing.Size(323, 51)
        Me.gbPrevYear.TabIndex = 10
        Me.gbPrevYear.TabStop = False
        Me.gbPrevYear.Text = "Previous Year"
        Me.gbPrevYear.Visible = False
        '
        'cboPrevRevno
        '
        Me.cboPrevRevno.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPrevRevno.FormattingEnabled = True
        Me.cboPrevRevno.Location = New System.Drawing.Point(197, 19)
        Me.cboPrevRevno.Name = "cboPrevRevno"
        Me.cboPrevRevno.Size = New System.Drawing.Size(53, 21)
        Me.cboPrevRevno.TabIndex = 3
        Me.cboPrevRevno.Visible = False
        '
        'lblPrevRevNo
        '
        Me.lblPrevRevNo.AutoSize = True
        Me.lblPrevRevNo.Location = New System.Drawing.Point(138, 22)
        Me.lblPrevRevNo.Name = "lblPrevRevNo"
        Me.lblPrevRevNo.Size = New System.Drawing.Size(53, 13)
        Me.lblPrevRevNo.TabIndex = 2
        Me.lblPrevRevNo.Text = "Rev. No.:"
        Me.lblPrevRevNo.Visible = False
        '
        'numPrevProjectNo
        '
        Me.numPrevProjectNo.Location = New System.Drawing.Point(75, 20)
        Me.numPrevProjectNo.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numPrevProjectNo.Name = "numPrevProjectNo"
        Me.numPrevProjectNo.Size = New System.Drawing.Size(38, 20)
        Me.numPrevProjectNo.TabIndex = 1
        Me.numPrevProjectNo.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Project No:"
        '
        'frmBG0420
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(420, 173)
        Me.Controls.Add(Me.cmdExcel)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.cmdPrint)
        Me.Controls.Add(Me.cmdPreview)
        Me.Controls.Add(Me.gbPrevYear)
        Me.Controls.Add(Me.cboRevNo)
        Me.Controls.Add(Me.lblRevNo)
        Me.Controls.Add(Me.numProjectNo)
        Me.Controls.Add(Me.lblProjectNo)
        Me.Controls.Add(Me.chkHideEstimate)
        Me.Controls.Add(Me.cboPeriodType)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.numYear)
        Me.Controls.Add(Me.lblYear)
        Me.Controls.Add(Me.lblFormTitle)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MinimumSize = New System.Drawing.Size(428, 200)
        Me.Name = "frmBG0420"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "frmBG0420"
        CType(Me.numYear, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numProjectNo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbPrevYear.ResumeLayout(False)
        Me.gbPrevYear.PerformLayout()
        CType(Me.numPrevProjectNo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblFormTitle As System.Windows.Forms.Label
    Friend WithEvents cmdPrint As System.Windows.Forms.Button
    Friend WithEvents numYear As System.Windows.Forms.NumericUpDown
    Friend WithEvents cmdPreview As System.Windows.Forms.Button
    Friend WithEvents lblYear As System.Windows.Forms.Label
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    Friend WithEvents PrintDocument1 As System.Drawing.Printing.PrintDocument
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents cboPeriodType As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents chkHideEstimate As System.Windows.Forms.CheckBox
    Friend WithEvents cmdExcel As System.Windows.Forms.Button
    Friend WithEvents numProjectNo As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblProjectNo As System.Windows.Forms.Label
    Friend WithEvents cboRevNo As System.Windows.Forms.ComboBox
    Friend WithEvents lblRevNo As System.Windows.Forms.Label
    Friend WithEvents gbPrevYear As System.Windows.Forms.GroupBox
    Friend WithEvents cboPrevRevno As System.Windows.Forms.ComboBox
    Friend WithEvents lblPrevRevNo As System.Windows.Forms.Label
    Friend WithEvents numPrevProjectNo As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label3 As System.Windows.Forms.Label
End Class
