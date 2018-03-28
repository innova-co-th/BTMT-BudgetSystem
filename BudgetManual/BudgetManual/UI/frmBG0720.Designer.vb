<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBG0720
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBG0720))
        Me.lblLockPic = New System.Windows.Forms.Label
        Me.cmdUnlock = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.lblFormTitle = New System.Windows.Forms.Label
        Me.cboPIC = New System.Windows.Forms.ComboBox
        Me.SuspendLayout()
        '
        'lblLockPic
        '
        Me.lblLockPic.AutoSize = True
        Me.lblLockPic.Location = New System.Drawing.Point(13, 47)
        Me.lblLockPic.Name = "lblLockPic"
        Me.lblLockPic.Size = New System.Drawing.Size(92, 13)
        Me.lblLockPic.TabIndex = 0
        Me.lblLockPic.Text = "Person In Charge:"
        '
        'cmdUnlock
        '
        Me.cmdUnlock.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdUnlock.Location = New System.Drawing.Point(183, 71)
        Me.cmdUnlock.Name = "cmdUnlock"
        Me.cmdUnlock.Size = New System.Drawing.Size(75, 23)
        Me.cmdUnlock.TabIndex = 2
        Me.cmdUnlock.Text = "&Unlock"
        Me.cmdUnlock.UseVisualStyleBackColor = True
        '
        'cmdClose
        '
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdClose.Location = New System.Drawing.Point(264, 71)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 23)
        Me.cmdClose.TabIndex = 3
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'lblFormTitle
        '
        Me.lblFormTitle.AutoSize = True
        Me.lblFormTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblFormTitle.Location = New System.Drawing.Point(12, 9)
        Me.lblFormTitle.Name = "lblFormTitle"
        Me.lblFormTitle.Size = New System.Drawing.Size(243, 24)
        Me.lblFormTitle.TabIndex = 11
        Me.lblFormTitle.Text = "Unlock Person In Charge"
        '
        'cboPIC
        '
        Me.cboPIC.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPIC.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.cboPIC.FormattingEnabled = True
        Me.cboPIC.Location = New System.Drawing.Point(111, 44)
        Me.cboPIC.Name = "cboPIC"
        Me.cboPIC.Size = New System.Drawing.Size(228, 21)
        Me.cboPIC.TabIndex = 1
        '
        'frmBG0720
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(420, 167)
        Me.Controls.Add(Me.cboPIC)
        Me.Controls.Add(Me.lblFormTitle)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.cmdUnlock)
        Me.Controls.Add(Me.lblLockPic)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MinimumSize = New System.Drawing.Size(428, 200)
        Me.Name = "frmBG0720"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "frmBG0720"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblLockPic As System.Windows.Forms.Label
    Friend WithEvents cmdUnlock As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents lblFormTitle As System.Windows.Forms.Label
    Friend WithEvents cboPIC As System.Windows.Forms.ComboBox
End Class
