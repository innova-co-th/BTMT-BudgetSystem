Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class frmBG0330

#Region "Variable"
    Private myClsBG0330BL As New clsBG0330BL
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
        cboPeriod2.Items.Clear()

        If myClsBG0330BL.SearchBudgetPeriod() = True Then

            For Each dr As DataRow In myClsBG0330BL.PeriodList.Rows
                If CInt(dr("PERIOD_TYPE")) = enumPeriodType.OriginalBudget Then
                    strTemp = CStr(dr("BUDGET_YEAR")) & " Original Budget " & CStr(dr("PROJECT_NO"))

                ElseIf CInt(dr("PERIOD_TYPE")) = enumPeriodType.EstimateBudget Then
                    strTemp = CStr(dr("BUDGET_YEAR")) & " Estimate Budget " & CStr(dr("PROJECT_NO"))

                ElseIf CInt(dr("PERIOD_TYPE")) = enumPeriodType.ForecastBudget Then
                    strTemp = CStr(dr("BUDGET_YEAR")) & " Forecast Budget " & CStr(dr("PROJECT_NO"))

                ElseIf CInt(dr("PERIOD_TYPE")) = enumPeriodType.MBPBudget Then
                    strTemp = CStr(dr("BUDGET_YEAR")) & " MBP Budget " & CStr(dr("PROJECT_NO"))

                End If

                cboPeriod.Items.Add(strTemp)
                cboPeriod2.Items.Add(strTemp)
            Next
        End If

        If cboPeriod.Items.Count > 0 Then
            cboPeriod.SelectedIndex = 0
        End If
        If cboPeriod2.Items.Count > 0 Then
            cboPeriod2.SelectedIndex = 0
        End If
    End Sub

    Private Sub SearchDatagrid()
        '// Set Parameters
        myClsBG0330BL.BudgetYear = Mid(cboPeriod2.Text, 1, 4)
        If cboPeriod2.Text.Contains("Original") Then
            myClsBG0330BL.PeriodType = CStr(enumPeriodType.OriginalBudget)
        ElseIf cboPeriod2.Text.Contains("Estimate") Then
            myClsBG0330BL.PeriodType = CStr(enumPeriodType.EstimateBudget)
        ElseIf cboPeriod2.Text.Contains("Forecast") Then
            myClsBG0330BL.PeriodType = CStr(enumPeriodType.ForecastBudget)
        ElseIf cboPeriod2.Text.Contains("MBP") Then
            myClsBG0330BL.PeriodType = CStr(enumPeriodType.MBPBudget)
        End If
        myClsBG0330BL.ProjectNo = Mid(cboPeriod2.Text, cboPeriod2.Text.LastIndexOf(" ") + 2, cboPeriod2.Text.Length - cboPeriod2.Text.LastIndexOf(" "))
        If myClsBG0330BL.SearchReopenAccount() Then
            grvMaster.DataSource = myClsBG0330BL.DtResult
        Else
            grvMaster.DataSource = Nothing
        End If
    End Sub

    Private Function CheckAccountExist(ByVal strAccountNo As String, ByVal strPicNo As String) As Boolean
        If strAccountNo = "All" Or strPicNo = "All" Then

            Return False

        ElseIf grvMaster.DataSource IsNot Nothing Then

            Dim dt As DataTable = CType(grvMaster.DataSource, DataTable)
            If dt.Select("ACCOUNT_NO = '" & strAccountNo & "' AND PERSON_IN_CHARGE_NO = '" & strPicNo & "'").Length > 0 Then

                Return True
            Else

                Return False
            End If
        Else

            Return False
        End If
    End Function

    Private Sub LoadAccountList()
        If myClsBG0330BL.GetAccountList() = True Then

            Dim dt As DataTable = myClsBG0330BL.DtResult
            Dim dr As DataRow = dt.NewRow
            dr("ACCOUNT_NO") = "All"
            dr("ACCOUNT_NAME_2") = "All"
            dt.Rows.InsertAt(dr, 0)

            cboAccount.DisplayMember = "ACCOUNT_NAME_2"
            cboAccount.ValueMember = "ACCOUNT_NO"
            cboAccount.DataSource = dt

            If cboAccount.Items.Count > 0 Then
                cboAccount.SelectedIndex = 0
            End If

        End If
    End Sub

    Private Sub loadPicList()
        If cboAccount.SelectedIndex >= 0 Then

            If cboAccount.SelectedValue.ToString = "All" Then     '// All Account
                If myClsBG0330BL.GetAllPicList() = True Then

                    cboPic.DisplayMember = "PERSON_IN_CHARGE_NO"
                    cboPic.ValueMember = "PERSON_IN_CHARGE_NO"
                    cboPic.DataSource = myClsBG0330BL.DtResult

                    If cboPic.Items.Count > 0 Then
                        cboPic.SelectedIndex = 0
                    End If

                End If
            Else
                myClsBG0330BL.AccountNo = cboAccount.SelectedValue.ToString
                If myClsBG0330BL.GetPicList() = True Then

                    '// Add "All" for all Pic
                    Dim dt As DataTable = myClsBG0330BL.DtResult
                    Dim dr As DataRow = dt.NewRow
                    dr("PERSON_IN_CHARGE_NO") = "All"
                    dt.Rows.InsertAt(dr, 0)

                    cboPic.DisplayMember = "PERSON_IN_CHARGE_NO"
                    cboPic.ValueMember = "PERSON_IN_CHARGE_NO"
                    cboPic.DataSource = dt

                    If cboPic.Items.Count > 0 Then
                        cboPic.SelectedIndex = 0
                    End If

                End If
            End If
        End If
    End Sub

#End Region

#Region "Control Event"
    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub frmBG0330_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadPeriodList()
        LoadAccountList()
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        If cboPeriod.SelectedIndex < 0 Then
            Exit Sub
        End If

        If MessageBox.Show("Are you sure to reopen this budget period?", Me.Text, MessageBoxButtons.YesNo, _
                          MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        '// Set Parameters
        myClsBG0330BL.BudgetYear = Mid(cboPeriod.Text, 1, 4)
        If cboPeriod.Text.Contains("Original") Then
            myClsBG0330BL.PeriodType = CStr(enumPeriodType.OriginalBudget)
        ElseIf cboPeriod.Text.Contains("Estimate") Then
            myClsBG0330BL.PeriodType = CStr(enumPeriodType.EstimateBudget)
        ElseIf cboPeriod.Text.Contains("Forecast") Then
            myClsBG0330BL.PeriodType = CStr(enumPeriodType.ForecastBudget)
        ElseIf cboPeriod.Text.Contains("MBP") Then
            myClsBG0330BL.PeriodType = CStr(enumPeriodType.MBPBudget)
        End If
        myClsBG0330BL.UserId = p_strUserId
        myClsBG0330BL.ProjectNo = Mid(cboPeriod.Text, cboPeriod.Text.LastIndexOf(" ") + 2, cboPeriod.Text.Length - cboPeriod.Text.LastIndexOf(" "))

        '// Call Function
        If myClsBG0330BL.ReopenPeriod() = True Then
            MessageBox.Show("Budget period was reopened", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            '// Write Transaction Log
            WriteTransactionLog(CStr(enumOperationCd.ReopenPeriod), myClsBG0330BL.BudgetYear, myClsBG0330BL.PeriodType, "", "", "", myClsBG0330BL.ProjectNo)

            p_frmBG0010.ShowBudgetMenu()
            Me.Close()
        Else
            MessageBox.Show("Can not reopen Budget period", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Private Sub cboPeriod2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPeriod2.SelectedIndexChanged
        SearchDataGrid()
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        If Me.grvMaster.SelectedRows Is Nothing OrElse Me.grvMaster.SelectedRows.Count = 0 Then
            Exit Sub
        End If

        If MessageBox.Show("Are you sure to delete Account?", Me.Text, MessageBoxButtons.YesNo, _
               MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        '// Set Parameters
        myClsBG0330BL.BudgetYear = Mid(cboPeriod2.Text, 1, 4)
        If cboPeriod2.Text.Contains("Original") Then
            myClsBG0330BL.PeriodType = CStr(enumPeriodType.OriginalBudget)
        ElseIf cboPeriod2.Text.Contains("Estimate") Then
            myClsBG0330BL.PeriodType = CStr(enumPeriodType.EstimateBudget)
        ElseIf cboPeriod2.Text.Contains("Forecast") Then
            myClsBG0330BL.PeriodType = CStr(enumPeriodType.ForecastBudget)
        ElseIf cboPeriod2.Text.Contains("MBP") Then
            myClsBG0330BL.PeriodType = CStr(enumPeriodType.MBPBudget)
        End If
        myClsBG0330BL.AccountNo = Me.grvMaster.SelectedRows(0).Cells(2).Value.ToString
        myClsBG0330BL.PicNo = Me.grvMaster.SelectedRows(0).Cells(4).Value.ToString
        myClsBG0330BL.ProjectNo = Mid(cboPeriod2.Text, cboPeriod2.Text.LastIndexOf(" ") + 2, cboPeriod2.Text.Length - cboPeriod2.Text.LastIndexOf(" "))

        If myClsBG0330BL.DeleteReopenAccount() = True Then
            MessageBox.Show("Account was deleted", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            '// Write Transaction Log
            WriteTransactionLog(CStr(enumOperationCd.EditReopenAccountMaster), myClsBG0330BL.BudgetYear, myClsBG0330BL.PeriodType, myClsBG0330BL.PicNo, "", "", myClsBG0330BL.ProjectNo)

            '// Refresh side menu
            p_frmBG0010.ShowBudgetMenu()

            SearchDatagrid()
        Else
            MessageBox.Show("There are error between delete Account", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub cmdDeleteAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeleteAll.Click
        If Me.grvMaster.SelectedRows.Count = 0 Then
            Exit Sub
        End If

        If MessageBox.Show("Are you sure to delete all Account?", Me.Text, MessageBoxButtons.YesNo, _
               MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        '// Set Parameters
        myClsBG0330BL.BudgetYear = Mid(cboPeriod2.Text, 1, 4)
        If cboPeriod2.Text.Contains("Original") Then
            myClsBG0330BL.PeriodType = CStr(enumPeriodType.OriginalBudget)
        ElseIf cboPeriod2.Text.Contains("Estimate") Then
            myClsBG0330BL.PeriodType = CStr(enumPeriodType.EstimateBudget)
        ElseIf cboPeriod2.Text.Contains("Forecast") Then
            myClsBG0330BL.PeriodType = CStr(enumPeriodType.ForecastBudget)
        ElseIf cboPeriod2.Text.Contains("MBP") Then
            myClsBG0330BL.PeriodType = CStr(enumPeriodType.MBPBudget)
        End If
        myClsBG0330BL.ProjectNo = Mid(cboPeriod2.Text, cboPeriod2.Text.LastIndexOf(" ") + 2, cboPeriod2.Text.Length - cboPeriod2.Text.LastIndexOf(" "))

        If myClsBG0330BL.DeleteAllReopenAccount() = True Then
            MessageBox.Show("All Account was deleted", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            '// Write Transaction Log
            WriteTransactionLog(CStr(enumOperationCd.EditReopenAccountMaster), myClsBG0330BL.BudgetYear, myClsBG0330BL.PeriodType, "", "", "", myClsBG0330BL.ProjectNo)

            '// Refresh side menu
            p_frmBG0010.ShowBudgetMenu()

            SearchDatagrid()
        Else
            MessageBox.Show("There are error between delete Account", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Dim arrList As ArrayList

        '// Set Parameters
        myClsBG0330BL.BudgetYear = Mid(cboPeriod2.Text, 1, 4)
        If cboPeriod2.Text.Contains("Original") Then
            myClsBG0330BL.PeriodType = CStr(enumPeriodType.OriginalBudget)
        ElseIf cboPeriod2.Text.Contains("Estimate") Then
            myClsBG0330BL.PeriodType = CStr(enumPeriodType.EstimateBudget)
        ElseIf cboPeriod2.Text.Contains("Forecast") Then
            myClsBG0330BL.PeriodType = CStr(enumPeriodType.ForecastBudget)
        ElseIf cboPeriod2.Text.Contains("MBP") Then
            myClsBG0330BL.PeriodType = CStr(enumPeriodType.MBPBudget)
        End If
        myClsBG0330BL.AccountNo = Me.cboAccount.SelectedValue.ToString
        myClsBG0330BL.PicNo = Me.cboPic.SelectedValue.ToString
        myClsBG0330BL.UserId = p_strUserId
        myClsBG0330BL.ProjectNo = Mid(cboPeriod2.Text, cboPeriod2.Text.LastIndexOf(" ") + 2, cboPeriod2.Text.Length - cboPeriod2.Text.LastIndexOf(" "))

        If CheckAccountExist(myClsBG0330BL.AccountNo, myClsBG0330BL.PicNo) = True Then
            MessageBox.Show("Selected Account already exist in the table", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        If myClsBG0330BL.AccountNo = "All" Then
            arrList = New ArrayList

            myClsBG0330BL.GetAccountList2()
            If myClsBG0330BL.DtResult.Rows.Count > 0 Then

                For i = 0 To myClsBG0330BL.DtResult.Rows.Count - 1
                    arrList.Add(CStr(myClsBG0330BL.DtResult.Rows(i).Item("ACCOUNT_NO")))
                Next
            End If

            myClsBG0330BL.AddAccountList = arrList

        ElseIf myClsBG0330BL.PicNo = "All" Then
            arrList = New ArrayList

            myClsBG0330BL.GetPicList()
            If myClsBG0330BL.DtResult.Rows.Count > 0 Then

                For i = 0 To myClsBG0330BL.DtResult.Rows.Count - 1
                    arrList.Add(CStr(myClsBG0330BL.DtResult.Rows(i).Item("PERSON_IN_CHARGE_NO")))
                Next
            End If

            myClsBG0330BL.AddPicList = arrList
        End If

        If myClsBG0330BL.AddReopenAccount() = True Then
            MessageBox.Show("Account was saved", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            '// Write Transaction Log
            WriteTransactionLog(CStr(enumOperationCd.EditReopenAccountMaster), _
                                myClsBG0330BL.BudgetYear, _
                                myClsBG0330BL.PeriodType, _
                                myClsBG0330BL.PicNo, _
                                "", _
                                "", _
                                myClsBG0330BL.ProjectNo)

            '// Refresh side menu
            p_frmBG0010.ShowBudgetMenu()

            SearchDatagrid()
        Else
            MessageBox.Show("There are error between add Account", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub cboAccount_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAccount.SelectedIndexChanged
        loadPicList()
    End Sub

#End Region

End Class