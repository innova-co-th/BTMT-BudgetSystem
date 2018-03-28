Public Class frmBG0720

#Region "Variable"
    Private myClsBG0720BL As New clsBG0720BL
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
    Private Sub LoadPicList()
        myClsBG0720BL.GetPicList()

        Me.cboPIC.DisplayMember = "PIC_NAME"
        Me.cboPIC.ValueMember = "PERSON_IN_CHARGE_NO"
        Me.cboPIC.DataSource = myClsBG0720BL.PicList
    End Sub
#End Region

#Region "Control Event"
    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub cmdUnlock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUnlock.Click
        If cboPIC.SelectedIndex >= 0 Then

            myClsBG0720BL.PicNo = cboPIC.SelectedValue.ToString

            If myClsBG0720BL.UnlockPic() = True Then
                MessageBox.Show("Person In Charge unlock successfully.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Error, Can not unlock the Person In Charge!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

            LoadPicList()
        End If
    End Sub

    Private Sub frmBG0720_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoadPicList()
    End Sub

#End Region

End Class