Public Class frmBG0710

#Region "Variable"
    Private myClsBG0710BL As New clsBG0710BL
#End Region

#Region "Overrides Function"
    Public Sub New(ByRef frmParent As Form, ByVal strFormName As String, ByVal blnMaximize As Boolean)
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.MdiParent = frmParent
        If blnMaximize Then
            Me.WindowState = FormWindowState.Maximized
        Else
            Me.WindowState = FormWindowState.Normal
        End If
        Me.Text = strFormName
    End Sub
#End Region

#Region "Function"

#End Region

#Region "Control Event"
    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        If txtNewPwd.Text <> txtConfirmPwd.Text Then
            MessageBox.Show("The passwords you typed do not match. Please try again.", Me.Text, _
                            MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            myClsBG0710BL.Password = txtConfirmPwd.Text
            If myClsBG0710BL.ChangePassword() = True Then
                MessageBox.Show("The password changed.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Error, Can not change the password!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

            Me.Close()
        End If
    End Sub

#End Region

End Class