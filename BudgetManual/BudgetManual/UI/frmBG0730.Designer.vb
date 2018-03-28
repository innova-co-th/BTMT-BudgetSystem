<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBG0730
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
        Me.lblFormTitle = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.txtBackupFile = New System.Windows.Forms.TextBox
        Me.cmdOpenFile = New System.Windows.Forms.Button
        Me.cmdRestore = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.chkUpload = New System.Windows.Forms.CheckBox
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.chkBudget = New System.Windows.Forms.CheckBox
        Me.lblObjectName = New System.Windows.Forms.Label
        Me.chkMaster = New System.Windows.Forms.CheckBox
        Me.lblStatus = New System.Windows.Forms.Label
        Me.cmdBackUp = New System.Windows.Forms.Button
        Me.Panel1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblFormTitle
        '
        Me.lblFormTitle.AutoSize = True
        Me.lblFormTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblFormTitle.Location = New System.Drawing.Point(12, 9)
        Me.lblFormTitle.Name = "lblFormTitle"
        Me.lblFormTitle.Size = New System.Drawing.Size(233, 24)
        Me.lblFormTitle.TabIndex = 0
        Me.lblFormTitle.Text = "Database / Restore Data"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.GroupBox2)
        Me.Panel1.Controls.Add(Me.ProgressBar1)
        Me.Panel1.Controls.Add(Me.cmdClose)
        Me.Panel1.Controls.Add(Me.lblObjectName)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.lblStatus)
        Me.Panel1.Location = New System.Drawing.Point(16, 36)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(393, 270)
        Me.Panel1.TabIndex = 1
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtBackupFile)
        Me.GroupBox2.Controls.Add(Me.cmdOpenFile)
        Me.GroupBox2.Controls.Add(Me.cmdRestore)
        Me.GroupBox2.Location = New System.Drawing.Point(3, 84)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(387, 78)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Restore"
        '
        'txtBackupFile
        '
        Me.txtBackupFile.Location = New System.Drawing.Point(6, 19)
        Me.txtBackupFile.Name = "txtBackupFile"
        Me.txtBackupFile.ReadOnly = True
        Me.txtBackupFile.Size = New System.Drawing.Size(294, 20)
        Me.txtBackupFile.TabIndex = 0
        Me.txtBackupFile.TabStop = False
        '
        'cmdOpenFile
        '
        Me.cmdOpenFile.Location = New System.Drawing.Point(306, 17)
        Me.cmdOpenFile.Name = "cmdOpenFile"
        Me.cmdOpenFile.Size = New System.Drawing.Size(75, 23)
        Me.cmdOpenFile.TabIndex = 1
        Me.cmdOpenFile.Text = "Browse..."
        Me.cmdOpenFile.UseVisualStyleBackColor = True
        '
        'cmdRestore
        '
        Me.cmdRestore.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdRestore.Location = New System.Drawing.Point(306, 46)
        Me.cmdRestore.Name = "cmdRestore"
        Me.cmdRestore.Size = New System.Drawing.Size(75, 23)
        Me.cmdRestore.TabIndex = 2
        Me.cmdRestore.Text = "&Restore"
        Me.cmdRestore.UseVisualStyleBackColor = True
        '
        'cmdClose
        '
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdClose.Location = New System.Drawing.Point(309, 240)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 23)
        Me.cmdClose.TabIndex = 5
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chkUpload)
        Me.GroupBox1.Controls.Add(Me.chkBudget)
        Me.GroupBox1.Controls.Add(Me.chkMaster)
        Me.GroupBox1.Controls.Add(Me.cmdBackUp)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(387, 75)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Backup"
        '
        'chkUpload
        '
        Me.chkUpload.AutoSize = True
        Me.chkUpload.Location = New System.Drawing.Point(101, 19)
        Me.chkUpload.Name = "chkUpload"
        Me.chkUpload.Size = New System.Drawing.Size(86, 17)
        Me.chkUpload.TabIndex = 1
        Me.chkUpload.Text = "Upload Data"
        Me.chkUpload.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(9, 216)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(375, 18)
        Me.ProgressBar1.TabIndex = 4
        '
        'chkBudget
        '
        Me.chkBudget.AutoSize = True
        Me.chkBudget.Location = New System.Drawing.Point(9, 19)
        Me.chkBudget.Name = "chkBudget"
        Me.chkBudget.Size = New System.Drawing.Size(86, 17)
        Me.chkBudget.TabIndex = 0
        Me.chkBudget.Text = "Budget Data"
        Me.chkBudget.UseVisualStyleBackColor = True
        '
        'lblObjectName
        '
        Me.lblObjectName.AutoSize = True
        Me.lblObjectName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblObjectName.Location = New System.Drawing.Point(9, 170)
        Me.lblObjectName.Name = "lblObjectName"
        Me.lblObjectName.Size = New System.Drawing.Size(15, 13)
        Me.lblObjectName.TabIndex = 2
        Me.lblObjectName.Text = "tt"
        '
        'chkMaster
        '
        Me.chkMaster.AutoSize = True
        Me.chkMaster.Location = New System.Drawing.Point(193, 19)
        Me.chkMaster.Name = "chkMaster"
        Me.chkMaster.Size = New System.Drawing.Size(84, 17)
        Me.chkMaster.TabIndex = 2
        Me.chkMaster.Text = "Master Data"
        Me.chkMaster.UseVisualStyleBackColor = True
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(9, 191)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(13, 13)
        Me.lblStatus.TabIndex = 3
        Me.lblStatus.Text = "tt"
        '
        'cmdBackUp
        '
        Me.cmdBackUp.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdBackUp.Location = New System.Drawing.Point(306, 37)
        Me.cmdBackUp.Name = "cmdBackUp"
        Me.cmdBackUp.Size = New System.Drawing.Size(75, 23)
        Me.cmdBackUp.TabIndex = 3
        Me.cmdBackUp.Text = "&Backup"
        Me.cmdBackUp.UseVisualStyleBackColor = True
        '
        'frmBG0730
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(421, 309)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.lblFormTitle)
        Me.Name = "frmBG0730"
        Me.Text = "frmBG0730"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblFormTitle As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cmdBackUp As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents lblObjectName As System.Windows.Forms.Label
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents chkUpload As System.Windows.Forms.CheckBox
    Friend WithEvents chkBudget As System.Windows.Forms.CheckBox
    Friend WithEvents chkMaster As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdRestore As System.Windows.Forms.Button
    Friend WithEvents txtBackupFile As System.Windows.Forms.TextBox
    Friend WithEvents cmdOpenFile As System.Windows.Forms.Button
End Class
