Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Globalization

Public Class frmBG0000

#Region "Variable"
    Private myClsBG0000BL As New clsBG0000BL()
#End Region

#Region "Function"
    ''' <summary>
    ''' InitForm
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub InitForm()
        Me.Text = My.Settings.ProgramTitle
        Me.lblVersion.Text &= Application.ProductVersion
    End Sub
#End Region

#Region "Control Event"
    Private Sub frmBG00000_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '// Set Current Culture to English-US
        System.Threading.Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US")

        '// Get Data Directory
        If My.Application.IsNetworkDeployed Then
            p_strAppPath = My.Application.Info.DirectoryPath
            p_strDataPath = My.Application.Deployment.DataDirectory
        Else
            p_strAppPath = Application.StartupPath
            p_strDataPath = Application.StartupPath
        End If

        InitForm()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        '// Check input data
        If txtUserId.Text.Trim = "" Then
            MessageBox.Show("Please enter User ID.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Dim drChildPIC As DataRow

        '// Set Function's Parameters
        myClsBG0000BL.UserId = txtUserId.Text.Trim
        myClsBG0000BL.Password = txtPassword.Text.Trim

        '// Call CheckLogIn Function
        If myClsBG0000BL.CheckLogin() = True Then

            '// Save current user's info
            p_strUserId = myClsBG0000BL.UserId
            p_strUserName = myClsBG0000BL.UserName
            p_intUserLevelId = myClsBG0000BL.UserLevelId
            p_intUserLevelName = myClsBG0000BL.UserLevelName
            p_strUserPIC = myClsBG0000BL.UserPIC

            '//-- Begin Add 2011/09/20 S.Watcharapong
            If myClsBG0000BL.CheckUserLoggedIn() = True Then
                MessageBox.Show("This User ID was logged in. Please try again later.", _
                                Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                txtPassword.Clear()
                txtUserId.Focus()
                txtUserId.SelectionStart = 0
                txtUserId.SelectionLength = txtUserId.TextLength

                Exit Sub
            End If
            '//-- End Add 2011/09/20

            If myClsBG0000BL.CheckLockPIC() = True Then
                '//-- Begin Edit 2011/08/25 S.Watcharapong
                ''MessageBox.Show("This Person In Charge was logged in by [" & myClsBG0000BL.UserName & "]." & vbNewLine & _
                ''                "Please try again later.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                MessageBox.Show("This Person In Charge was logged in by [" & myClsBG0000BL.UserName & "]." & vbNewLine & _
                "Program will switch to [Read-Only] Mode!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                p_blnReadOnlyMode = True

                p_frmBG0010 = New frmBG0010
                p_frmBG0010.Show()

                Me.Close()
                '//-- End Edit 2011/08/25
            Else

                '// Child PIC Lock

                If myClsBG0000BL.GetChildPicList() = True Then

                    For Each drChildPIC In myClsBG0000BL.ChildPicList.Rows

                        myClsBG0000BL.ChildPIC = drChildPIC("PIC_CHILD_NO").ToString
                        If myClsBG0000BL.CheckLockChildPIC() = True Then
                            MessageBox.Show("Child Person In Charge [" & myClsBG0000BL.ChildPIC & "] was logged In." & vbNewLine & _
                                    "Please be careful if you want to edit data. ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)

                            Exit For

                        End If

                    Next

                End If

                '// Add Lock PIC
                If p_strUserPIC <> "0000" Then
                    myClsBG0000BL.AddLockPIC()
                End If

                p_frmBG0010 = New frmBG0010
                p_frmBG0010.Show()

                Me.Close()
            End If

        Else
            MessageBox.Show("User ID or Password is incorrect." & vbNewLine & "Please try again.", _
                            Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

#End Region

End Class
