Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class frmBG0395

#Region "Variable"
    Private myClsBG0395BL As New clsBG0395BL
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

        If myClsBG0395BL.SearchBudgetPeriod() = True Then

            For Each dr As DataRow In myClsBG0395BL.PeriodList.Rows
                If CInt(dr("PERIOD_TYPE")) = enumPeriodType.OriginalBudget Then
                    strTemp = CStr(dr("BUDGET_YEAR")) & " Original Budget " & CStr(dr("PROJECT_NO"))

                ElseIf CInt(dr("PERIOD_TYPE")) = enumPeriodType.EstimateBudget Then
                    strTemp = CStr(dr("BUDGET_YEAR")) & " Estimate Budget " & CStr(dr("PROJECT_NO"))

                ElseIf CInt(dr("PERIOD_TYPE")) = enumPeriodType.ReviseBudget Then
                    strTemp = CStr(dr("BUDGET_YEAR")) & " Revise Budget " & CStr(dr("PROJECT_NO"))

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
    Private Sub SetHideRadio()
        If cboPeriod.SelectedIndex < 0 Then
            Exit Sub
        End If

        '// Set Parameters
        myClsBG0395BL.BudgetYear = Mid(cboPeriod.Text, 1, 4)
        If cboPeriod.Text.Contains("Original") Then
            myClsBG0395BL.PeriodType = CStr(enumPeriodType.OriginalBudget)
        ElseIf cboPeriod.Text.Contains("Estimate") Then
            myClsBG0395BL.PeriodType = CStr(enumPeriodType.EstimateBudget)
        ElseIf cboPeriod.Text.Contains("Revise") Then
            myClsBG0395BL.PeriodType = CStr(enumPeriodType.ReviseBudget)
        ElseIf cboPeriod.Text.Contains("MTP") Then
            myClsBG0395BL.PeriodType = CStr(enumPeriodType.MTPBudget)
        End If
        myClsBG0395BL.ProjectNo = Mid(cboPeriod.Text, cboPeriod.Text.LastIndexOf(" ") + 2, cboPeriod.Text.Length - cboPeriod.Text.LastIndexOf(" "))

        '// Call Function
        If myClsBG0395BL.GetHideFlag() = True AndAlso myClsBG0395BL.PeriodList.Rows.Count > 0 Then

            If CInt(myClsBG0395BL.PeriodList.Rows(0)("HIDE_FLAG")) = 1 Then
                Me.rdoHide.Checked = True
            Else
                Me.rdoShow.Checked = True
            End If

        End If
    End Sub

#End Region

    Private Sub frmBG0395_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoadPeriodList()
    End Sub

    Private Sub cboPeriod_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPeriod.SelectedIndexChanged
       
        SetHideRadio()
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        If cboPeriod.SelectedIndex < 0 Then
            Exit Sub
        End If

        If MessageBox.Show("Are you sure to save view budget period?", Me.Text, MessageBoxButtons.YesNo, _
                          MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        '// Set Parameters
        myClsBG0395BL.BudgetYear = Mid(cboPeriod.Text, 1, 4)
        If cboPeriod.Text.Contains("Original") Then
            myClsBG0395BL.PeriodType = CStr(enumPeriodType.OriginalBudget)
        ElseIf cboPeriod.Text.Contains("Estimate") Then
            myClsBG0395BL.PeriodType = CStr(enumPeriodType.EstimateBudget)
        ElseIf cboPeriod.Text.Contains("Revise") Then
            myClsBG0395BL.PeriodType = CStr(enumPeriodType.ReviseBudget)
        ElseIf cboPeriod.Text.Contains("MTP") Then
            myClsBG0395BL.PeriodType = CStr(enumPeriodType.MTPBudget)
        End If
        myClsBG0395BL.UserId = p_strUserId
        myClsBG0395BL.ProjectNo = Mid(cboPeriod.Text, cboPeriod.Text.LastIndexOf(" ") + 2, cboPeriod.Text.Length - cboPeriod.Text.LastIndexOf(" "))
        If Me.rdoHide.Checked = True Then
            myClsBG0395BL.HideFlag = "1"
        Else
            myClsBG0395BL.HideFlag = "0"
        End If

        '// Call Function
        If myClsBG0395BL.SaveViewBudgetPeriod() = True Then
            MessageBox.Show("Budget period was saved", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            '// Write Transaction Log
            WriteTransactionLog(CStr(enumOperationCd.EditViewBudgetPeriod), myClsBG0395BL.BudgetYear, myClsBG0395BL.PeriodType, "", "", "", myClsBG0395BL.ProjectNo)

            SetHideRadio()
        Else
            MessageBox.Show("Can not save Budget period!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub
End Class