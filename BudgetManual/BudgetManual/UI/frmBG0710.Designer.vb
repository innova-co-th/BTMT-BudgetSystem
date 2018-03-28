<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBG0710
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBG0710))
        Me.lblNewPwd = New System.Windows.Forms.Label
        Me.lblConfirmPwd = New System.Windows.Forms.Label
        Me.txtNewPwd = New System.Windows.Forms.TextBox
        Me.txtConfirmPwd = New System.Windows.Forms.TextBox
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.lblFormTitle = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'lblNewPwd
        '
        Me.lblNewPwd.AutoSize = True
        Me.lblNewPwd.Location = New System.Drawing.Point(13, 47)
        Me.lblNewPwd.Name = "lblNewPwd"
        Me.lblNewPwd.Size = New System.Drawing.Size(108, 13)
        Me.lblNewPwd.TabIndex = 0
        Me.lblNewPwd.Text = "Type New &Password:"
        '
        'lblConfirmPwd
        '
        Me.lblConfirmPwd.AutoSize = True
        Me.lblConfirmPwd.Location = New System.Drawing.Point(14, 73)
        Me.lblConfirmPwd.Name = "lblConfirmPwd"
        Me.lblConfirmPwd.Size = New System.Drawing.Size(119, 13)
        Me.lblConfirmPwd.TabIndex = 2
        Me.lblConfirmPwd.Text = "&Confirm New Password:"
        '
        'txtNewPwd
        '
        Me.txtNewPwd.Location = New System.Drawing.Point(139, 44)
        Me.txtNewPwd.MaxLength = 50
        Me.txtNewPwd.Name = "txtNewPwd"
        Me.txtNewPwd.PasswordChar = Global.Microsoft.VisualBasic.ChrW(9679)
        Me.txtNewPwd.Size = New System.Drawing.Size(200, 20)
        Me.txtNewPwd.TabIndex = 1
        '
        'txtConfirmPwd
        '
        Me.txtConfirmPwd.Location = New System.Drawing.Point(139, 70)
        Me.txtConfirmPwd.MaxLength = 50
        Me.txtConfirmPwd.Name = "txtConfirmPwd"
        Me.txtConfirmPwd.PasswordChar = Global.Microsoft.VisualBasic.ChrW(9679)
        Me.txtConfirmPwd.Size = New System.Drawing.Size(200, 20)
        Me.txtConfirmPwd.TabIndex = 3
        '
        'cmdSave
        '
        Me.cmdSave.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdSave.Location = New System.Drawing.Point(183, 96)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(75, 23)
        Me.cmdSave.TabIndex = 4
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'cmdClose
        '
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdClose.Location = New System.Drawing.Point(264, 96)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 23)
        Me.cmdClose.TabIndex = 5
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'lblFormTitle
        '
        Me.lblFormTitle.AutoSize = True
        Me.lblFormTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblFormTitle.Location = New System.Drawing.Point(12, 9)
        Me.lblFormTitle.Name = "lblFormTitle"
        Me.lblFormTitle.Size = New System.Drawing.Size(179, 24)
        Me.lblFormTitle.TabIndex = 11
        Me.lblFormTitle.Text = "Change Password"
        '
        'frmBG0710
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(420, 167)
        Me.Controls.Add(Me.lblFormTitle)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.txtConfirmPwd)
        Me.Controls.Add(Me.txtNewPwd)
        Me.Controls.Add(Me.lblConfirmPwd)
        Me.Controls.Add(Me.lblNewPwd)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MinimumSize = New System.Drawing.Size(428, 200)
        Me.Name = "frmBG0710"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "frmBG0710"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblNewPwd As System.Windows.Forms.Label
    Friend WithEvents lblConfirmPwd As System.Windows.Forms.Label
    Friend WithEvents txtNewPwd As System.Windows.Forms.TextBox
    Friend WithEvents txtConfirmPwd As System.Windows.Forms.TextBox
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents lblFormTitle As System.Windows.Forms.Label
End Class
