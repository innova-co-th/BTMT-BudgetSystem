<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBG0350
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBG0350))
        Me.cmdClose = New System.Windows.Forms.Button
        Me.lblFormTitle = New System.Windows.Forms.Label
        Me.cmdImport = New System.Windows.Forms.Button
        Me.cmdBrowse = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtFilePath = New System.Windows.Forms.TextBox
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.cboPeriodType = New System.Windows.Forms.ComboBox
        Me.lblConfirmPwd = New System.Windows.Forms.Label
        Me.numProjectNo = New System.Windows.Forms.NumericUpDown
        Me.lblProjectNo = New System.Windows.Forms.Label
        Me.lblRevNo = New System.Windows.Forms.Label
        Me.numRev = New System.Windows.Forms.NumericUpDown
        CType(Me.numProjectNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numRev, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdClose
        '
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdClose.Location = New System.Drawing.Point(499, 124)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 23)
        Me.cmdClose.TabIndex = 6
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'lblFormTitle
        '
        Me.lblFormTitle.AutoSize = True
        Me.lblFormTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblFormTitle.Location = New System.Drawing.Point(12, 9)
        Me.lblFormTitle.Name = "lblFormTitle"
        Me.lblFormTitle.Size = New System.Drawing.Size(216, 24)
        Me.lblFormTitle.TabIndex = 11
        Me.lblFormTitle.Text = "Import Data From SAP"
        '
        'cmdImport
        '
        Me.cmdImport.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdImport.Location = New System.Drawing.Point(418, 124)
        Me.cmdImport.Name = "cmdImport"
        Me.cmdImport.Size = New System.Drawing.Size(75, 23)
        Me.cmdImport.TabIndex = 5
        Me.cmdImport.Text = "&Import"
        Me.cmdImport.UseVisualStyleBackColor = True
        '
        'cmdBrowse
        '
        Me.cmdBrowse.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdBrowse.Location = New System.Drawing.Point(499, 95)
        Me.cmdBrowse.Name = "cmdBrowse"
        Me.cmdBrowse.Size = New System.Drawing.Size(75, 23)
        Me.cmdBrowse.TabIndex = 4
        Me.cmdBrowse.Text = "&Browse..."
        Me.cmdBrowse.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 100)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(77, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Data File &Path:"
        '
        'txtFilePath
        '
        Me.txtFilePath.Location = New System.Drawing.Point(96, 97)
        Me.txtFilePath.MaxLength = 100
        Me.txtFilePath.Name = "txtFilePath"
        Me.txtFilePath.Size = New System.Drawing.Size(403, 20)
        Me.txtFilePath.TabIndex = 3
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.DefaultExt = "txt"
        Me.OpenFileDialog1.Filter = "Text files|*.txt"
        Me.OpenFileDialog1.Title = "Please select data file"
        '
        'cboPeriodType
        '
        Me.cboPeriodType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPeriodType.FormattingEnabled = True
        Me.cboPeriodType.Items.AddRange(New Object() {"Original Budget", "Estimate Budget", "Revise Budget", "Estimate Budget (Actual Oct)", "Revise Budget (Actual Apr)", "Original Budget (Input Data)", "Estimate Budget (Input Data)", "Revise Budget (Input Data)", "Revise Budget (Input MTP)"})
        Me.cboPeriodType.Location = New System.Drawing.Point(96, 44)
        Me.cboPeriodType.Name = "cboPeriodType"
        Me.cboPeriodType.Size = New System.Drawing.Size(245, 21)
        Me.cboPeriodType.TabIndex = 1
        '
        'lblConfirmPwd
        '
        Me.lblConfirmPwd.AutoSize = True
        Me.lblConfirmPwd.Location = New System.Drawing.Point(13, 47)
        Me.lblConfirmPwd.Name = "lblConfirmPwd"
        Me.lblConfirmPwd.Size = New System.Drawing.Size(67, 13)
        Me.lblConfirmPwd.TabIndex = 0
        Me.lblConfirmPwd.Text = "Period &Type:"
        '
        'numProjectNo
        '
        Me.numProjectNo.Location = New System.Drawing.Point(96, 71)
        Me.numProjectNo.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numProjectNo.Name = "numProjectNo"
        Me.numProjectNo.Size = New System.Drawing.Size(38, 20)
        Me.numProjectNo.TabIndex = 13
        Me.numProjectNo.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'lblProjectNo
        '
        Me.lblProjectNo.AutoSize = True
        Me.lblProjectNo.Location = New System.Drawing.Point(13, 73)
        Me.lblProjectNo.Name = "lblProjectNo"
        Me.lblProjectNo.Size = New System.Drawing.Size(60, 13)
        Me.lblProjectNo.TabIndex = 14
        Me.lblProjectNo.Text = "Project No:"
        '
        'lblRevNo
        '
        Me.lblRevNo.AutoSize = True
        Me.lblRevNo.Location = New System.Drawing.Point(147, 74)
        Me.lblRevNo.Name = "lblRevNo"
        Me.lblRevNo.Size = New System.Drawing.Size(53, 13)
        Me.lblRevNo.TabIndex = 15
        Me.lblRevNo.Text = "Rev. No.:"
        '
        'numRev
        '
        Me.numRev.Location = New System.Drawing.Point(221, 71)
        Me.numRev.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numRev.Name = "numRev"
        Me.numRev.Size = New System.Drawing.Size(38, 20)
        Me.numRev.TabIndex = 16
        Me.numRev.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'frmBG0350
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(587, 166)
        Me.Controls.Add(Me.numRev)
        Me.Controls.Add(Me.lblRevNo)
        Me.Controls.Add(Me.numProjectNo)
        Me.Controls.Add(Me.lblProjectNo)
        Me.Controls.Add(Me.cboPeriodType)
        Me.Controls.Add(Me.lblConfirmPwd)
        Me.Controls.Add(Me.cmdBrowse)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtFilePath)
        Me.Controls.Add(Me.cmdImport)
        Me.Controls.Add(Me.lblFormTitle)
        Me.Controls.Add(Me.cmdClose)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MinimumSize = New System.Drawing.Size(595, 166)
        Me.Name = "frmBG0350"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "frmBG0350"
        CType(Me.numProjectNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numRev, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents lblFormTitle As System.Windows.Forms.Label
    Friend WithEvents cmdImport As System.Windows.Forms.Button
    Friend WithEvents cmdBrowse As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtFilePath As System.Windows.Forms.TextBox
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents cboPeriodType As System.Windows.Forms.ComboBox
    Friend WithEvents lblConfirmPwd As System.Windows.Forms.Label
    Friend WithEvents numProjectNo As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblProjectNo As System.Windows.Forms.Label
    Friend WithEvents lblRevNo As System.Windows.Forms.Label
    Friend WithEvents numRev As System.Windows.Forms.NumericUpDown
End Class
