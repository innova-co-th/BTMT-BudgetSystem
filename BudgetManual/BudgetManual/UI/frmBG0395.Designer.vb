<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBG0395
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBG0395))
        Me.cboPeriod = New System.Windows.Forms.ComboBox
        Me.lblFormTitle = New System.Windows.Forms.Label
        Me.cmdClose = New System.Windows.Forms.Button
        Me.cmdSave = New System.Windows.Forms.Button
        Me.lblYear = New System.Windows.Forms.Label
        Me.rdoHide = New System.Windows.Forms.RadioButton
        Me.rdoShow = New System.Windows.Forms.RadioButton
        Me.SuspendLayout()
        '
        'cboPeriod
        '
        Me.cboPeriod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPeriod.FormattingEnabled = True
        Me.cboPeriod.Location = New System.Drawing.Point(96, 44)
        Me.cboPeriod.Name = "cboPeriod"
        Me.cboPeriod.Size = New System.Drawing.Size(212, 21)
        Me.cboPeriod.TabIndex = 13
        '
        'lblFormTitle
        '
        Me.lblFormTitle.AutoSize = True
        Me.lblFormTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblFormTitle.Location = New System.Drawing.Point(12, 9)
        Me.lblFormTitle.Name = "lblFormTitle"
        Me.lblFormTitle.Size = New System.Drawing.Size(195, 24)
        Me.lblFormTitle.TabIndex = 16
        Me.lblFormTitle.Text = "View Budget Period"
        '
        'cmdClose
        '
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdClose.Location = New System.Drawing.Point(233, 102)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 23)
        Me.cmdClose.TabIndex = 15
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(152, 102)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(75, 23)
        Me.cmdSave.TabIndex = 14
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'lblYear
        '
        Me.lblYear.AutoSize = True
        Me.lblYear.Location = New System.Drawing.Point(13, 47)
        Me.lblYear.Name = "lblYear"
        Me.lblYear.Size = New System.Drawing.Size(77, 13)
        Me.lblYear.TabIndex = 12
        Me.lblYear.Text = "Budget &Period:"
        '
        'rdoHide
        '
        Me.rdoHide.AutoSize = True
        Me.rdoHide.Location = New System.Drawing.Point(96, 72)
        Me.rdoHide.Name = "rdoHide"
        Me.rdoHide.Size = New System.Drawing.Size(47, 17)
        Me.rdoHide.TabIndex = 17
        Me.rdoHide.TabStop = True
        Me.rdoHide.Text = "Hide"
        Me.rdoHide.UseVisualStyleBackColor = True
        '
        'rdoShow
        '
        Me.rdoShow.AutoSize = True
        Me.rdoShow.Location = New System.Drawing.Point(149, 72)
        Me.rdoShow.Name = "rdoShow"
        Me.rdoShow.Size = New System.Drawing.Size(52, 17)
        Me.rdoShow.TabIndex = 17
        Me.rdoShow.TabStop = True
        Me.rdoShow.Text = "Show"
        Me.rdoShow.UseVisualStyleBackColor = True
        '
        'frmBG0395
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(420, 167)
        Me.Controls.Add(Me.rdoShow)
        Me.Controls.Add(Me.rdoHide)
        Me.Controls.Add(Me.cboPeriod)
        Me.Controls.Add(Me.lblFormTitle)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.lblYear)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MinimumSize = New System.Drawing.Size(428, 200)
        Me.Name = "frmBG0395"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "frmBG0395"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cboPeriod As System.Windows.Forms.ComboBox
    Friend WithEvents lblFormTitle As System.Windows.Forms.Label
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents lblYear As System.Windows.Forms.Label
    Friend WithEvents rdoHide As System.Windows.Forms.RadioButton
    Friend WithEvents rdoShow As System.Windows.Forms.RadioButton
End Class
