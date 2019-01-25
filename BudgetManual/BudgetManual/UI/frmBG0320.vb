Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class frmBG0320

#Region "Variable"
    Private myClsBG0320BL As New clsBG0320BL
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
        Dim strTemp As String = String.Empty

        '// Initialize controls
        cboPeriod.Items.Clear()

        If myClsBG0320BL.SearchBudgetPeriod() = True Then

            For Each dr As DataRow In myClsBG0320BL.PeriodList.Rows
                If CInt(dr("PERIOD_TYPE")) = enumPeriodType.OriginalBudget Then
                    strTemp = CStr(dr("BUDGET_YEAR")) & " Original Budget " & CStr(dr("PROJECT_NO"))

                ElseIf CInt(dr("PERIOD_TYPE")) = enumPeriodType.EstimateBudget Then
                    strTemp = CStr(dr("BUDGET_YEAR")) & " Estimate Budget " & CStr(dr("PROJECT_NO"))

                ElseIf CInt(dr("PERIOD_TYPE")) = enumPeriodType.ForecastBudget Then
                    strTemp = CStr(dr("BUDGET_YEAR")) & " Forecast Budget " & CStr(dr("PROJECT_NO"))

                ElseIf CInt(dr("PERIOD_TYPE")) = enumPeriodType.MTPBudget Then
                    strTemp = CStr(dr("BUDGET_YEAR")) & " MTP Budget " & CStr(dr("PROJECT_NO"))

                End If

                cboPeriod.Items.Add(strTemp)
            Next

        End If

        If cboPeriod.Items.Count > 0 Then
            cboPeriod.SelectedIndex = 0
        End If
    End Sub
#End Region

#Region "Control Event"
    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub frmBG0320_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadPeriodList()
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        If cboPeriod.SelectedIndex < 0 Then
            Exit Sub
        End If

        If MessageBox.Show("Are you sure to close this budget period?", Me.Text, MessageBoxButtons.YesNo, _
                          MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        '// Set Parameters
        myClsBG0320BL.BudgetYear = Mid(cboPeriod.Text, 1, 4)
        If cboPeriod.Text.Contains("Original") Then
            myClsBG0320BL.PeriodType = CStr(enumPeriodType.OriginalBudget)
        ElseIf cboPeriod.Text.Contains("Estimate") Then
            myClsBG0320BL.PeriodType = CStr(enumPeriodType.EstimateBudget)
        ElseIf cboPeriod.Text.Contains("Forecast") Then
            myClsBG0320BL.PeriodType = CStr(enumPeriodType.ForecastBudget)
        ElseIf cboPeriod.Text.Contains("MTP") Then
            myClsBG0320BL.PeriodType = CStr(enumPeriodType.MTPBudget)
        End If
        myClsBG0320BL.UserId = p_strUserId
        myClsBG0320BL.ProjectNo = Mid(cboPeriod.Text, cboPeriod.Text.LastIndexOf(" ") + 2, cboPeriod.Text.Length - cboPeriod.Text.LastIndexOf(" "))

        '// Call Function
        If myClsBG0320BL.ClosePeriod() = True Then
            MessageBox.Show("Budget period was closed", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            '// Write Transaction Log
            WriteTransactionLog(CStr(enumOperationCd.ClosePeriod), myClsBG0320BL.BudgetYear, myClsBG0320BL.PeriodType, "", "", "", myClsBG0320BL.ProjectNo)

            p_frmBG0010.ShowBudgetMenu()
            Me.Close()
        Else
            MessageBox.Show("Can not close Budget period!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

#End Region

End Class