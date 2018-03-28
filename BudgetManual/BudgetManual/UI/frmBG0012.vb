Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class frmBG0012

#Region "Variable"
    Private myClsBG0012BL As New clsBG0012BL
#End Region

#Region "Overrides Function"
    Public Sub New(ByVal strFormName As String)
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.Text = strFormName
    End Sub
#End Region

#Region "Function"

#End Region

#Region "Control Event"
    Private Sub frmBG0012_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '// Intialize Controls
        myClsBG0012BL.GetOptions()
        txtHomeURL.Text = myClsBG0012BL.HomePageURL
        txtSharedFolder.Text = myClsBG0012BL.SharedFolder
        txtHighLightMTP.Text = myClsBG0012BL.HighLightMTP
        txtFromAddr.Text = myClsBG0012BL.FromAddr
        txtSmtpServer.Text = myClsBG0012BL.SmtpServer '"mail.truemail.co.th"
        txtSmtpPort.Text = myClsBG0012BL.SmtpPort '"25"
        chkUseAuthentication.Checked = myClsBG0012BL.UseAuthentication
        txtSmtpUser.Text = myClsBG0012BL.SmtpUser
        txtSmtpPassword.Text = myClsBG0012BL.SmtpPassword
        chkEnableAutoMail.Checked = myClsBG0012BL.EnableAutoMail

        Dim ms1 As IO.MemoryStream
        Dim ms2 As IO.MemoryStream

        If myClsBG0012BL.AuthPic1 IsNot Nothing Then
            If myClsBG0012BL.AuthPic1.GetUpperBound(0) > 0 Then
                ms1 = New IO.MemoryStream(myClsBG0012BL.AuthPic1)
                picAuth1.Image = Image.FromStream(ms1)
            End If
        End If

        If myClsBG0012BL.AuthPic2 IsNot Nothing Then
            If myClsBG0012BL.AuthPic2.GetUpperBound(0) > 0 Then
                ms2 = New IO.MemoryStream(myClsBG0012BL.AuthPic2)
                picAuth2.Image = Image.FromStream(ms2)
            End If
        End If

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        If MessageBox.Show("Are you sure to save the options?", Me.Text, MessageBoxButtons.YesNo, _
                      MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        '/ Check validation
        If chkUseAuthentication.Checked And (txtSmtpUser.Text.Trim = "" Or txtSmtpPassword.Text.Trim = "") Then
            MessageBox.Show("Please input SMTP server's user and password.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        '// Save Settings Change
        myClsBG0012BL.HomePageURL = txtHomeURL.Text.Trim
        myClsBG0012BL.SharedFolder = txtSharedFolder.Text.Trim
        myClsBG0012BL.HighLightMTP = txtHighLightMTP.Text
        myClsBG0012BL.FromAddr = txtFromAddr.Text
        myClsBG0012BL.SmtpServer = txtSmtpServer.Text
        myClsBG0012BL.SmtpPort = txtSmtpPort.Text
        myClsBG0012BL.UseAuthentication = chkUseAuthentication.Checked
        myClsBG0012BL.SmtpUser = txtSmtpUser.Text
        myClsBG0012BL.SmtpPassword = txtSmtpPassword.Text
        myClsBG0012BL.EnableAutoMail = chkEnableAutoMail.Checked

        Dim ms1 As New IO.MemoryStream
        Dim ms2 As New IO.MemoryStream

        If picAuth1.Image IsNot Nothing Then
            Try
                picAuth1.Image.Save(ms1, Imaging.ImageFormat.Jpeg)
                myClsBG0012BL.AuthPic1 = ms1.GetBuffer()
            Catch ex As Exception
                MessageBox.Show("Can not save image of Authorize 1", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            myClsBG0012BL.AuthPic1 = Nothing
        End If

        If picAuth2.Image IsNot Nothing Then
            Try
                picAuth2.Image.Save(ms2, Imaging.ImageFormat.Jpeg)
                myClsBG0012BL.AuthPic2 = ms2.GetBuffer()
            Catch ex As Exception
                MessageBox.Show("Can not save image of Authorize 2", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            myClsBG0012BL.AuthPic2 = Nothing
        End If

        If myClsBG0012BL.UpdateOptions() = True Then
            MessageBox.Show("The options saved", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            '// Set New Settings
            If p_frmBG0110 IsNot Nothing Then
                p_frmBG0110.ReloadHome()
            End If
        End If

        '// Set Public Settings
        p_blnSendAutoMail = chkEnableAutoMail.Checked
        p_strAutoMailFromAddr = txtFromAddr.Text

        BGSendMail.SmtpServer = txtSmtpServer.Text
        If IsNumeric(txtSmtpPort.Text) Then
            BGSendMail.SmtpPort = CInt(txtSmtpPort.Text)
        Else
            BGSendMail.SmtpPort = 25
        End If
        If chkUseAuthentication.Checked Then
            BGSendMail.UseAuthentication = True
            BGSendMail.SmtpUser = txtSmtpUser.Text
            BGSendMail.SmtpPassword = txtSmtpPassword.Text
        Else
            BGSendMail.UseAuthentication = False
        End If
        BGSendMail.FromAddress = txtFromAddr.Text

        Me.Close()
    End Sub

    Private Sub cmdTestSendMail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTestSendMail.Click
        '// Test Send auto mail
        BGSendMail.SmtpServer = txtSmtpServer.Text
        If IsNumeric(txtSmtpPort.Text) Then
            BGSendMail.SmtpPort = CInt(txtSmtpPort.Text)
        Else
            BGSendMail.SmtpPort = 25
        End If
        If chkUseAuthentication.Checked Then
            BGSendMail.UseAuthentication = True
            BGSendMail.SmtpUser = txtSmtpUser.Text
            BGSendMail.SmtpPassword = txtSmtpPassword.Text
        Else
            BGSendMail.UseAuthentication = False
        End If
        BGSendMail.FromAddress = txtFromAddr.Text
        BGSendMail.SendMessage(txtFromAddr.Text, "Test auto mail!", _
                               "This is auto mail from " & My.Settings.ProgramTitle & "." & vbNewLine & vbNewLine & _
                               "Sent time: " & Now.ToString("yyyy/MM/dd HH:mm:ss"))
    End Sub

    Private Sub txtHighLightMTP_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtHighLightMTP.KeyPress
        If IsNumeric(e.KeyChar) Or Asc(e.KeyChar) = Keys.Back Or Asc(e.KeyChar) = Keys.Delete Or e.KeyChar = CChar(".") Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub txtSmtpPort_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSmtpPort.KeyPress
        If IsNumeric(e.KeyChar) Or Asc(e.KeyChar) = Keys.Back Or Asc(e.KeyChar) = Keys.Delete Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub chkUseAuthentication_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseAuthentication.CheckedChanged
        If chkUseAuthentication.Checked Then
            txtSmtpUser.Enabled = True
            txtSmtpPassword.Enabled = True
        Else
            txtSmtpUser.Enabled = False
            txtSmtpPassword.Enabled = False
        End If
    End Sub

    Private Sub cmdChangePic1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdChangePic1.Click
        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            picAuth1.ImageLocation = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub cmdChangePic2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdChangePic2.Click
        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            picAuth2.ImageLocation = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub cmdUndo1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUndo1.Click
        Dim ms As IO.MemoryStream

        If myClsBG0012BL.AuthPic1 IsNot Nothing Then
            If myClsBG0012BL.AuthPic1.GetUpperBound(0) > 0 Then
                ms = New IO.MemoryStream(myClsBG0012BL.AuthPic1)
                picAuth1.Image = Image.FromStream(ms)
            End If
        End If
    End Sub

    Private Sub cmdUndo2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUndo2.Click
        Dim ms As IO.MemoryStream

        If myClsBG0012BL.AuthPic2 IsNot Nothing Then
            If myClsBG0012BL.AuthPic2.GetUpperBound(0) > 0 Then
                ms = New IO.MemoryStream(myClsBG0012BL.AuthPic2)
                picAuth2.Image = Image.FromStream(ms)
            End If
        End If
    End Sub

    Private Sub cmdClear1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear1.Click
        picAuth1.Image = Nothing
    End Sub

    Private Sub cmdClear2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear2.Click
        picAuth2.Image = Nothing
    End Sub

#End Region


End Class
