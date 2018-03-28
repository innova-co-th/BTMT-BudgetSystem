<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBG0012
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBG0012))
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.fraAutoMail = New System.Windows.Forms.GroupBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.chkEnableAutoMail = New System.Windows.Forms.CheckBox
        Me.cmdTestSendMail = New System.Windows.Forms.Button
        Me.txtFromAddr = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtSmtpPassword = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtSmtpUser = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtSmtpPort = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.chkUseAuthentication = New System.Windows.Forms.CheckBox
        Me.txtSmtpServer = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.fraGeneral = New System.Windows.Forms.GroupBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtHighLightMTP = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtSharedFolder = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtHomeURL = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cmdClear2 = New System.Windows.Forms.Button
        Me.cmdClear1 = New System.Windows.Forms.Button
        Me.cmdUndo2 = New System.Windows.Forms.Button
        Me.cmdUndo1 = New System.Windows.Forms.Button
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.cmdChangePic2 = New System.Windows.Forms.Button
        Me.cmdChangePic1 = New System.Windows.Forms.Button
        Me.picAuth2 = New System.Windows.Forms.PictureBox
        Me.picAuth1 = New System.Windows.Forms.PictureBox
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.fraAutoMail.SuspendLayout()
        Me.fraGeneral.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.picAuth2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picAuth1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdSave
        '
        Me.cmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSave.Location = New System.Drawing.Point(376, 498)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(75, 23)
        Me.cmdSave.TabIndex = 18
        Me.cmdSave.Text = "Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(457, 498)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 0
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'fraAutoMail
        '
        Me.fraAutoMail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.fraAutoMail.Controls.Add(Me.Label11)
        Me.fraAutoMail.Controls.Add(Me.chkEnableAutoMail)
        Me.fraAutoMail.Controls.Add(Me.cmdTestSendMail)
        Me.fraAutoMail.Controls.Add(Me.txtFromAddr)
        Me.fraAutoMail.Controls.Add(Me.Label10)
        Me.fraAutoMail.Controls.Add(Me.Label6)
        Me.fraAutoMail.Controls.Add(Me.txtSmtpPassword)
        Me.fraAutoMail.Controls.Add(Me.Label8)
        Me.fraAutoMail.Controls.Add(Me.txtSmtpUser)
        Me.fraAutoMail.Controls.Add(Me.Label9)
        Me.fraAutoMail.Controls.Add(Me.txtSmtpPort)
        Me.fraAutoMail.Controls.Add(Me.Label7)
        Me.fraAutoMail.Controls.Add(Me.chkUseAuthentication)
        Me.fraAutoMail.Controls.Add(Me.txtSmtpServer)
        Me.fraAutoMail.Controls.Add(Me.Label4)
        Me.fraAutoMail.Location = New System.Drawing.Point(12, 129)
        Me.fraAutoMail.Name = "fraAutoMail"
        Me.fraAutoMail.Size = New System.Drawing.Size(520, 201)
        Me.fraAutoMail.TabIndex = 22
        Me.fraAutoMail.TabStop = False
        Me.fraAutoMail.Text = "Auto Mail Options"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(68, 171)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(89, 13)
        Me.Label11.TabIndex = 35
        Me.Label11.Text = "Enable Auto mail:"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'chkEnableAutoMail
        '
        Me.chkEnableAutoMail.AutoSize = True
        Me.chkEnableAutoMail.Location = New System.Drawing.Point(162, 171)
        Me.chkEnableAutoMail.Name = "chkEnableAutoMail"
        Me.chkEnableAutoMail.Size = New System.Drawing.Size(15, 14)
        Me.chkEnableAutoMail.TabIndex = 10
        Me.chkEnableAutoMail.UseVisualStyleBackColor = True
        '
        'cmdTestSendMail
        '
        Me.cmdTestSendMail.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdTestSendMail.Location = New System.Drawing.Point(417, 19)
        Me.cmdTestSendMail.Name = "cmdTestSendMail"
        Me.cmdTestSendMail.Size = New System.Drawing.Size(75, 23)
        Me.cmdTestSendMail.TabIndex = 11
        Me.cmdTestSendMail.Text = "Send Test"
        Me.cmdTestSendMail.UseVisualStyleBackColor = True
        '
        'txtFromAddr
        '
        Me.txtFromAddr.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFromAddr.Location = New System.Drawing.Point(162, 21)
        Me.txtFromAddr.MaxLength = 500
        Me.txtFromAddr.Name = "txtFromAddr"
        Me.txtFromAddr.Size = New System.Drawing.Size(243, 20)
        Me.txtFromAddr.TabIndex = 4
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(29, 24)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(127, 13)
        Me.Label10.TabIndex = 32
        Me.Label10.Text = "Auto mail's From Address:"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(56, 99)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(100, 13)
        Me.Label6.TabIndex = 31
        Me.Label6.Text = "Use Authentication:"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtSmtpPassword
        '
        Me.txtSmtpPassword.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSmtpPassword.Enabled = False
        Me.txtSmtpPassword.Location = New System.Drawing.Point(162, 145)
        Me.txtSmtpPassword.MaxLength = 500
        Me.txtSmtpPassword.Name = "txtSmtpPassword"
        Me.txtSmtpPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(9679)
        Me.txtSmtpPassword.Size = New System.Drawing.Size(243, 20)
        Me.txtSmtpPassword.TabIndex = 9
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Enabled = False
        Me.Label8.Location = New System.Drawing.Point(100, 148)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(56, 13)
        Me.Label8.TabIndex = 29
        Me.Label8.Text = "Password:"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtSmtpUser
        '
        Me.txtSmtpUser.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSmtpUser.Enabled = False
        Me.txtSmtpUser.Location = New System.Drawing.Point(162, 119)
        Me.txtSmtpUser.MaxLength = 500
        Me.txtSmtpUser.Name = "txtSmtpUser"
        Me.txtSmtpUser.Size = New System.Drawing.Size(243, 20)
        Me.txtSmtpUser.TabIndex = 8
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Enabled = False
        Me.Label9.Location = New System.Drawing.Point(95, 122)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(61, 13)
        Me.Label9.TabIndex = 27
        Me.Label9.Text = "User name:"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtSmtpPort
        '
        Me.txtSmtpPort.Location = New System.Drawing.Point(162, 73)
        Me.txtSmtpPort.MaxLength = 500
        Me.txtSmtpPort.Name = "txtSmtpPort"
        Me.txtSmtpPort.Size = New System.Drawing.Size(44, 20)
        Me.txtSmtpPort.TabIndex = 6
        Me.txtSmtpPort.Text = "25"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(127, 76)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(29, 13)
        Me.Label7.TabIndex = 25
        Me.Label7.Text = "Port:"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'chkUseAuthentication
        '
        Me.chkUseAuthentication.AutoSize = True
        Me.chkUseAuthentication.Location = New System.Drawing.Point(162, 99)
        Me.chkUseAuthentication.Name = "chkUseAuthentication"
        Me.chkUseAuthentication.Size = New System.Drawing.Size(15, 14)
        Me.chkUseAuthentication.TabIndex = 7
        Me.chkUseAuthentication.UseVisualStyleBackColor = True
        '
        'txtSmtpServer
        '
        Me.txtSmtpServer.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSmtpServer.Location = New System.Drawing.Point(162, 47)
        Me.txtSmtpServer.MaxLength = 500
        Me.txtSmtpServer.Name = "txtSmtpServer"
        Me.txtSmtpServer.Size = New System.Drawing.Size(243, 20)
        Me.txtSmtpServer.TabIndex = 5
        Me.txtSmtpServer.Text = "mail.truemail.co.th"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(82, 50)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(74, 13)
        Me.Label4.TabIndex = 22
        Me.Label4.Text = "SMTP Server:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'fraGeneral
        '
        Me.fraGeneral.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.fraGeneral.Controls.Add(Me.Label5)
        Me.fraGeneral.Controls.Add(Me.txtHighLightMTP)
        Me.fraGeneral.Controls.Add(Me.Label3)
        Me.fraGeneral.Controls.Add(Me.txtSharedFolder)
        Me.fraGeneral.Controls.Add(Me.Label2)
        Me.fraGeneral.Controls.Add(Me.txtHomeURL)
        Me.fraGeneral.Controls.Add(Me.Label1)
        Me.fraGeneral.Location = New System.Drawing.Point(12, 12)
        Me.fraGeneral.Name = "fraGeneral"
        Me.fraGeneral.Size = New System.Drawing.Size(520, 111)
        Me.fraGeneral.TabIndex = 23
        Me.fraGeneral.TabStop = False
        Me.fraGeneral.Text = "General Options"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(206, 77)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(39, 13)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "K Baht"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtHighLightMTP
        '
        Me.txtHighLightMTP.Location = New System.Drawing.Point(162, 74)
        Me.txtHighLightMTP.MaxLength = 500
        Me.txtHighLightMTP.Name = "txtHighLightMTP"
        Me.txtHighLightMTP.Size = New System.Drawing.Size(44, 20)
        Me.txtHighLightMTP.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(53, 77)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(103, 13)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Highlight MTP Over:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtSharedFolder
        '
        Me.txtSharedFolder.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSharedFolder.Location = New System.Drawing.Point(162, 48)
        Me.txtSharedFolder.MaxLength = 500
        Me.txtSharedFolder.Name = "txtSharedFolder"
        Me.txtSharedFolder.Size = New System.Drawing.Size(330, 20)
        Me.txtSharedFolder.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(80, 51)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(76, 13)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "Shared Folder:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtHomeURL
        '
        Me.txtHomeURL.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtHomeURL.Location = New System.Drawing.Point(162, 22)
        Me.txtHomeURL.MaxLength = 500
        Me.txtHomeURL.Name = "txtHomeURL"
        Me.txtHomeURL.Size = New System.Drawing.Size(330, 20)
        Me.txtHomeURL.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(93, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(63, 13)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Home URL:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cmdClear2)
        Me.GroupBox1.Controls.Add(Me.cmdClear1)
        Me.GroupBox1.Controls.Add(Me.cmdUndo2)
        Me.GroupBox1.Controls.Add(Me.cmdUndo1)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.cmdChangePic2)
        Me.GroupBox1.Controls.Add(Me.cmdChangePic1)
        Me.GroupBox1.Controls.Add(Me.picAuth2)
        Me.GroupBox1.Controls.Add(Me.picAuth1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 336)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(520, 155)
        Me.GroupBox1.TabIndex = 24
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Authorize Image"
        '
        'cmdClear2
        '
        Me.cmdClear2.Location = New System.Drawing.Point(386, 93)
        Me.cmdClear2.Name = "cmdClear2"
        Me.cmdClear2.Size = New System.Drawing.Size(64, 23)
        Me.cmdClear2.TabIndex = 17
        Me.cmdClear2.Text = "Clear"
        Me.cmdClear2.UseVisualStyleBackColor = True
        '
        'cmdClear1
        '
        Me.cmdClear1.Location = New System.Drawing.Point(176, 93)
        Me.cmdClear1.Name = "cmdClear1"
        Me.cmdClear1.Size = New System.Drawing.Size(63, 23)
        Me.cmdClear1.TabIndex = 14
        Me.cmdClear1.Text = "Clear"
        Me.cmdClear1.UseVisualStyleBackColor = True
        '
        'cmdUndo2
        '
        Me.cmdUndo2.Location = New System.Drawing.Point(386, 64)
        Me.cmdUndo2.Name = "cmdUndo2"
        Me.cmdUndo2.Size = New System.Drawing.Size(64, 23)
        Me.cmdUndo2.TabIndex = 16
        Me.cmdUndo2.Text = "Undo"
        Me.cmdUndo2.UseVisualStyleBackColor = True
        '
        'cmdUndo1
        '
        Me.cmdUndo1.Location = New System.Drawing.Point(176, 64)
        Me.cmdUndo1.Name = "cmdUndo1"
        Me.cmdUndo1.Size = New System.Drawing.Size(63, 23)
        Me.cmdUndo1.TabIndex = 13
        Me.cmdUndo1.Text = "Undo"
        Me.cmdUndo1.UseVisualStyleBackColor = True
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(300, 118)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(60, 13)
        Me.Label13.TabIndex = 5
        Me.Label13.Text = "Authorize 2"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(90, 118)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(60, 13)
        Me.Label12.TabIndex = 4
        Me.Label12.Text = "Authorize 1"
        '
        'cmdChangePic2
        '
        Me.cmdChangePic2.Location = New System.Drawing.Point(386, 35)
        Me.cmdChangePic2.Name = "cmdChangePic2"
        Me.cmdChangePic2.Size = New System.Drawing.Size(64, 23)
        Me.cmdChangePic2.TabIndex = 15
        Me.cmdChangePic2.Text = "Change"
        Me.cmdChangePic2.UseVisualStyleBackColor = True
        '
        'cmdChangePic1
        '
        Me.cmdChangePic1.Location = New System.Drawing.Point(176, 35)
        Me.cmdChangePic1.Name = "cmdChangePic1"
        Me.cmdChangePic1.Size = New System.Drawing.Size(63, 23)
        Me.cmdChangePic1.TabIndex = 12
        Me.cmdChangePic1.Text = "Change"
        Me.cmdChangePic1.UseVisualStyleBackColor = True
        '
        'picAuth2
        '
        Me.picAuth2.BackColor = System.Drawing.Color.White
        Me.picAuth2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picAuth2.Location = New System.Drawing.Point(280, 35)
        Me.picAuth2.Name = "picAuth2"
        Me.picAuth2.Size = New System.Drawing.Size(100, 80)
        Me.picAuth2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picAuth2.TabIndex = 1
        Me.picAuth2.TabStop = False
        '
        'picAuth1
        '
        Me.picAuth1.BackColor = System.Drawing.Color.White
        Me.picAuth1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picAuth1.Location = New System.Drawing.Point(70, 35)
        Me.picAuth1.Name = "picAuth1"
        Me.picAuth1.Size = New System.Drawing.Size(100, 80)
        Me.picAuth1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picAuth1.TabIndex = 0
        Me.picAuth1.TabStop = False
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.Filter = "JPEG files|*.jpg"
        '
        'frmBG0012
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(544, 533)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.fraGeneral)
        Me.Controls.Add(Me.fraAutoMail)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdSave)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmBG0012"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "frmBG0012"
        Me.fraAutoMail.ResumeLayout(False)
        Me.fraAutoMail.PerformLayout()
        Me.fraGeneral.ResumeLayout(False)
        Me.fraGeneral.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.picAuth2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picAuth1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents fraAutoMail As System.Windows.Forms.GroupBox
    Friend WithEvents txtFromAddr As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtSmtpPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtSmtpUser As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtSmtpPort As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents chkUseAuthentication As System.Windows.Forms.CheckBox
    Friend WithEvents txtSmtpServer As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents fraGeneral As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtHighLightMTP As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtSharedFolder As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtHomeURL As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdTestSendMail As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents chkEnableAutoMail As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdChangePic2 As System.Windows.Forms.Button
    Friend WithEvents cmdChangePic1 As System.Windows.Forms.Button
    Friend WithEvents picAuth2 As System.Windows.Forms.PictureBox
    Friend WithEvents picAuth1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents cmdUndo2 As System.Windows.Forms.Button
    Friend WithEvents cmdUndo1 As System.Windows.Forms.Button
    Friend WithEvents cmdClear2 As System.Windows.Forms.Button
    Friend WithEvents cmdClear1 As System.Windows.Forms.Button

End Class
