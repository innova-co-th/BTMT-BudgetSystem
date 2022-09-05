Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class frmBG0310

#Region "Variable"
    Private myClsBG0310BL As New clsBG0310BL
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
    Public Sub LoadPeriodList()
        '// Initialize controls
        numYear.Value = Now.Year

        myClsBG0310BL.OpenPeriodFlg = "1"
        myClsBG0310BL.GetOpenPeriodList()

        If myClsBG0310BL.PeriodList IsNot Nothing AndAlso myClsBG0310BL.PeriodList.Rows.Count > 0 Then
            cboPeriodType.DisplayMember = "PERIOD_TYPE_NAME"
            cboPeriodType.ValueMember = "PERIOD_TYPE_ID"
            cboPeriodType.DataSource = myClsBG0310BL.PeriodList
    
            cboPeriodType.SelectedIndex = 0
        End If

        'cboPeriodType.Items.Clear()
        'cboPeriodType.Items.Add("Original Budget")
        'cboPeriodType.Items.Add("Estimate Budget")
        'cboPeriodType.Items.Add("Forecast Budget")
        'cboPeriodType.Items.Add("MTP Budget")


    End Sub
#End Region

#Region "Control Event"
    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub frmBG0310_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadPeriodList()
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        '// Check if new MTP, must have Forecast budget.
        If CStr(cboPeriodType.SelectedValue) = CStr(enumPeriodType.MBPBudget) Then
            myClsBG0310BL.BudgetYear = numYear.Value.ToString("0000")
            myClsBG0310BL.ProjectNo = numProjectNo.Value.ToString

            If myClsBG0310BL.CheckForecastExist = False Then

                MessageBox.Show("Can not create Budget period!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

                Exit Sub

            End If

        End If

        If MessageBox.Show("Are you sure to create new budget period?", Me.Text, MessageBoxButtons.YesNo, _
                          MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        '// Set Parameters
        myClsBG0310BL.BudgetYear = numYear.Value.ToString("0000")
        myClsBG0310BL.PeriodType = CStr(cboPeriodType.SelectedValue)
        myClsBG0310BL.UserId = p_strUserId
        myClsBG0310BL.ProjectNo = numProjectNo.Value.ToString

        '// Call Function: Insert Budget period
        If myClsBG0310BL.CreateNewPeriod() = True Then
            MessageBox.Show("Budget period was created", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            '// Write Transaction Log
            WriteTransactionLog(CStr(enumOperationCd.OpenNewPeriod), myClsBG0310BL.BudgetYear, myClsBG0310BL.PeriodType, "", "", "", myClsBG0310BL.ProjectNo)

            p_frmBG0010.ShowBudgetMenu()
            Me.Close()
        Else
            MessageBox.Show("Can not create Budget period!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

#End Region

    Private Sub cboPeriodType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPeriodType.SelectedIndexChanged

        If cboPeriodType.SelectedIndex >= 0 Then

            If CStr(cboPeriodType.SelectedValue) = CStr(enumPeriodType.MBPBudget) Then
                Me.numProjectNo.Enabled = True
            Else
                Me.numProjectNo.Value = 1
                Me.numProjectNo.Enabled = False
            End If

        End If

    End Sub
End Class