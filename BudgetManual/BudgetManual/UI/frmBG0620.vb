Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.IO
Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class frmBG0620

#Region "Variable"
    Private myClsBG0620BL As New clsBG0620BL
    Private dtAccount As New DataTable
    Private dtDepartment As New DataTable
    Private dtBudgetType As New DataTable
    Private dtCostType As New DataTable
    Private dtCost As New DataTable
    Private dtAssetGroup As New DataTable
    Private dtPersonInCharge As New DataTable
    Private dtExpenseType As New DataTable
    Private strSelectedOrderNo As String = String.Empty
    Private currentDS As String = String.Empty
    Private myPicShowFlag As String

    Dim dtBudgetTypeFilter As New DataTable
    Dim dtAccountFilter As New DataTable
    Dim dtDepartmentFilter As New DataTable
    Dim dtAssetGroupFilter As New DataTable
    Dim dtPersonInChargeFilter As New DataTable
    Dim dtCostTypeFilter As New DataTable
    Dim dtCostFilter As New DataTable
    Dim dtExpenseTypeFilter As New DataTable

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

    Private Sub clearData()
        Me.txtOrderNo.Enabled = True
        Me.txtOrderNo.Text = ""
        Me.txtOrderName.Text = ""
        Me.txtCostCenter.Text = ""
        Me.cboBudgetType.SelectedIndex = 0
        Me.cboAccount.SelectedIndex = 0
        Me.cboCostType.SelectedIndex = -1 '0
        Me.cboCost.SelectedIndex = -1 '0
        Me.cboAssetGroup.SelectedIndex = 0
        Me.cboDepartment.SelectedIndex = 0
        Me.cboPersonInCharge.SelectedIndex = 0
        Me.cboExpenseType.SelectedIndex = -1 '0
        Me.chkActive.Checked = True 'False
        Me.txtRemarks.Text = ""
    End Sub

    Private Sub clearFilter()
        Me.txtOrderNoFilter.Text = ""
        Me.txtOrderNameFilter.Text = ""
        Me.txtCostCenterFilter.Text = ""
        Me.cboBudgetTypeFilter.SelectedIndex = 0
        Me.cboAccountFilter.SelectedIndex = 0
        Me.cboCostTypeFilter.SelectedIndex = -1 '0
        Me.cboCostFilter.SelectedIndex = -1 '0
        Me.cboAssetGroupFilter.SelectedIndex = -1
        Me.cboDepartmentFilter.SelectedIndex = 0
        Me.cboPersonInChargeFilter.SelectedIndex = 0
        Me.cboExpenseTypeFilter.SelectedIndex = -1 '0
        Me.chkActiveFilter.Checked = False
    End Sub

    Private Sub init()
        Dim costType As BGConstant.enumCostType
        Dim cost As BGConstant.enumCost
        Dim expenseType As BGConstant.enumExpenseType
        Dim drTemp As DataRow

        '// Add items for budget type
        addDataColumn(dtBudgetType, New DataColumn("KEY"))
        addDataColumn(dtBudgetType, New DataColumn("VALUE"))

        drTemp = dtBudgetType.NewRow
        drTemp("KEY") = BGConstant.P_BUDGET_TYPE_ASSET
        drTemp("VALUE") = "Asset"
        dtBudgetType.Rows.Add(drTemp)

        drTemp = dtBudgetType.NewRow
        drTemp("KEY") = BGConstant.P_BUDGET_TYPE_EXPENSE
        drTemp("VALUE") = "Expense"
        dtBudgetType.Rows.Add(drTemp)

        fillComboBox(Me.cboBudgetType, dtBudgetType, "KEY", "VALUE", False)

        '// Add item for account list
        If myClsBG0620BL.getAccountList Then
            dtAccount = myClsBG0620BL.DtResult
        End If
        fillComboBox(Me.cboAccount, dtAccount, "ACCOUNT_NO", "ACCOUNT_NAME_2", False)

        '// Add item for department list
        If myClsBG0620BL.getDepartmentList Then
            dtDepartment = myClsBG0620BL.DtResult
        End If
        fillComboBox(Me.cboDepartment, dtDepartment, "DEPT_NO", "DEPT_NAME_2", False)

        '// Add item for asset group
        If myClsBG0620BL.getAssetGroupList Then
            dtAssetGroup = myClsBG0620BL.DtResult
        End If
        fillComboBox(Me.cboAssetGroup, dtAssetGroup, "ASSET_GROUP_NO", "ASSET_GROUP_NAME_2", False)

        '// Add item for person in charge
        If myClsBG0620BL.getPersonInChargeList Then
            dtPersonInCharge = myClsBG0620BL.DtResult
        End If
        fillComboBox(Me.cboPersonInCharge, dtPersonInCharge, "PERSON_IN_CHARGE_NO", "PIC_NAME", False)

        '// Add item for cost type
        addDataColumn(dtCostType, New DataColumn("KEY"))
        addDataColumn(dtCostType, New DataColumn("VALUE"))
        For Each costType In [Enum].GetValues(GetType(BGConstant.enumCostType))
            drTemp = dtCostType.NewRow
            drTemp("KEY") = CStr(costType)
            Select Case costType
                Case enumCostType.FixedCost
                    drTemp("VALUE") = "Fixed Cost"
                Case enumCostType.VariableCost
                    drTemp("VALUE") = "Variable Cost"
            End Select
            dtCostType.Rows.Add(drTemp)
        Next
        fillComboBox(Me.cboCostType, dtCostType, "KEY", "VALUE", False)

        '// Add item for cost
        addDataColumn(dtCost, New DataColumn("KEY"))
        addDataColumn(dtCost, New DataColumn("VALUE"))
        For Each cost In [Enum].GetValues(GetType(BGConstant.enumCost))
            drTemp = dtCost.NewRow
            drTemp("KEY") = CStr(cost)
            Select Case cost
                Case enumCost.FC
                    drTemp("VALUE") = "FC"
                Case enumCost.ADMIN
                    drTemp("VALUE") = "Admin"
            End Select

            dtCost.Rows.Add(drTemp)
        Next
        fillComboBox(Me.cboCost, dtCost, "KEY", "VALUE", False)

        '// Add item for expense type
        addDataColumn(dtExpenseType, New DataColumn("KEY"))
        addDataColumn(dtExpenseType, New DataColumn("VALUE"))
        For Each expenseType In [Enum].GetValues(GetType(BGConstant.enumExpenseType))
            drTemp = dtExpenseType.NewRow
            drTemp("KEY") = CStr(expenseType)
            Select Case expenseType
                Case enumExpenseType.FixedExpense
                    drTemp("VALUE") = "Fixed Expense"
                Case enumExpenseType.LaborExpense
                    drTemp("VALUE") = "Labor Expense"
                Case enumExpenseType.VariableExpense
                    drTemp("VALUE") = "Variable Expense"
            End Select

            dtExpenseType.Rows.Add(drTemp)
        Next
        fillComboBox(Me.cboExpenseType, dtExpenseType, "KEY", "VALUE", False)
    End Sub

    Private Sub initFilter()
        Dim costType As BGConstant.enumCostType
        Dim cost As BGConstant.enumCost
        Dim expenseType As BGConstant.enumExpenseType
        Dim drTemp As DataRow

        '// Add items for budget type
        addDataColumn(dtBudgetTypeFilter, New DataColumn("KEY"))
        addDataColumn(dtBudgetTypeFilter, New DataColumn("VALUE"))

        drTemp = dtBudgetTypeFilter.NewRow
        drTemp("KEY") = BGConstant.P_BUDGET_TYPE_ASSET
        drTemp("VALUE") = "Asset"
        dtBudgetTypeFilter.Rows.Add(drTemp)

        drTemp = dtBudgetTypeFilter.NewRow
        drTemp("KEY") = BGConstant.P_BUDGET_TYPE_EXPENSE
        drTemp("VALUE") = "Expense"
        dtBudgetTypeFilter.Rows.Add(drTemp)

        fillComboBox(Me.cboBudgetTypeFilter, dtBudgetTypeFilter, "KEY", "VALUE", True)

        '// Add item for account list
        If myClsBG0620BL.getAccountList Then
            dtAccountFilter = myClsBG0620BL.DtResult
        End If
        fillComboBox(Me.cboAccountFilter, dtAccountFilter, "ACCOUNT_NO", "ACCOUNT_NAME_2", True)

        '// Add item for department list
        If myClsBG0620BL.getDepartmentList Then
            dtDepartmentFilter = myClsBG0620BL.DtResult
        End If
        fillComboBox(Me.cboDepartmentFilter, dtDepartmentFilter, "DEPT_NO", "DEPT_NAME_2", True)

        '// Add item for asset group
        If myClsBG0620BL.getAssetGroupList Then
            dtAssetGroupFilter = myClsBG0620BL.DtResult
        End If
        fillComboBox(Me.cboAssetGroupFilter, dtAssetGroupFilter, "ASSET_GROUP_NO", "ASSET_GROUP_NAME_2", True)

        '// Add item for person in charge
        If myClsBG0620BL.getPersonInChargeList Then
            dtPersonInChargeFilter = myClsBG0620BL.DtResult
        End If
        fillComboBox(Me.cboPersonInChargeFilter, dtPersonInChargeFilter, "PERSON_IN_CHARGE_NO", "PIC_NAME", True)

        '// Add item for cost type
        addDataColumn(dtCostTypeFilter, New DataColumn("KEY"))
        addDataColumn(dtCostTypeFilter, New DataColumn("VALUE"))
        For Each costType In [Enum].GetValues(GetType(BGConstant.enumCostType))
            drTemp = dtCostTypeFilter.NewRow
            drTemp("KEY") = CStr(costType)
            Select Case costType
                Case enumCostType.FixedCost
                    drTemp("VALUE") = "Fixed Cost"
                Case enumCostType.VariableCost
                    drTemp("VALUE") = "Variable Cost"
            End Select
            dtCostTypeFilter.Rows.Add(drTemp)
        Next
        fillComboBox(Me.cboCostTypeFilter, dtCostTypeFilter, "KEY", "VALUE", True)

        '// Add item for cost
        addDataColumn(dtCostFilter, New DataColumn("KEY"))
        addDataColumn(dtCostFilter, New DataColumn("VALUE"))
        For Each cost In [Enum].GetValues(GetType(BGConstant.enumCost))
            drTemp = dtCostFilter.NewRow
            drTemp("KEY") = CStr(cost)
            Select Case cost
                Case enumCost.FC
                    drTemp("VALUE") = "FC"
                Case enumCost.ADMIN
                    drTemp("VALUE") = "Admin"
            End Select

            dtCostFilter.Rows.Add(drTemp)
        Next
        fillComboBox(Me.cboCostFilter, dtCostFilter, "KEY", "VALUE", True)

        '// Add item for expense type
        addDataColumn(dtExpenseTypeFilter, New DataColumn("KEY"))
        addDataColumn(dtExpenseTypeFilter, New DataColumn("VALUE"))
        For Each expenseType In [Enum].GetValues(GetType(BGConstant.enumExpenseType))
            drTemp = dtExpenseTypeFilter.NewRow
            drTemp("KEY") = CStr(expenseType)
            Select Case expenseType
                Case enumExpenseType.FixedExpense
                    drTemp("VALUE") = "Fixed Expense"
                Case enumExpenseType.LaborExpense
                    drTemp("VALUE") = "Labor Expense"
                Case enumExpenseType.VariableExpense
                    drTemp("VALUE") = "Variable Expense"
            End Select

            dtExpenseTypeFilter.Rows.Add(drTemp)
        Next
        fillComboBox(Me.cboExpenseTypeFilter, dtExpenseTypeFilter, "KEY", "VALUE", True)
    End Sub

    Private Sub fillComboBox(ByVal cbo As ComboBox, _
                             ByVal dt As DataTable, _
                             ByVal keyColName As String, _
                             ByVal textColName As String, _
                             ByVal allowFirstEmptyItem As Boolean)
        Dim str As String = String.Empty
        cbo.Items.Clear()
        If dt.Rows.Count > 0 Then
            If allowFirstEmptyItem Then
                Dim dr As DataRow = dt.NewRow
                dr(keyColName) = ""
                dr(textColName) = "All"
                dt.Rows.InsertAt(dr, 0)
            End If
            cbo.DataSource = dt
            cbo.ValueMember = keyColName
            cbo.DisplayMember = textColName
        End If
        cbo.SelectedIndex = 0

    End Sub

    Private Sub findItemValue(ByVal cbo As ComboBox, _
                              ByVal str As String)
        Dim intFoundRow As Integer = -1
        intFoundRow = cbo.FindString(str)

        If intFoundRow < 0 Then
            cbo.SelectedIndex = -1
        Else
            cbo.SelectedIndex = intFoundRow
        End If
    End Sub

    Private Sub displaySelectedItem()
        If Me.grvMaster.RowCount > 0 And Me.grvMaster.SelectedRows.Count = 1 Then
            'Dim costType As BGConstant.enumCostType
            'Dim cost As BGConstant.enumCost
            'Dim expenseType As BGConstant.enumExpenseType
            Dim costTypeValue As String = String.Empty
            Dim costValue As String = String.Empty
            Dim expenseTypeValue As String = String.Empty

            strSelectedOrderNo = Me.grvMaster.SelectedRows(0).Cells(0).Value.ToString

            Me.txtOrderNo.Enabled = False
            Me.txtOrderNo.Text = Me.grvMaster.SelectedRows(0).Cells(0).Value.ToString
            Me.txtOrderName.Text = Me.grvMaster.SelectedRows(0).Cells(1).Value.ToString
            'findItemValue(Me.cboBudgetType, Me.grvMaster.SelectedRows(0).Cells(2).Value.ToString)
            findItemValue(Me.cboAccount, Me.grvMaster.SelectedRows(0).Cells(3).Value.ToString)
            Me.txtCostCenter.Text = Me.grvMaster.SelectedRows(0).Cells(4).Value.ToString

            If Me.grvMaster.SelectedRows(0).Cells(2).Value.ToString = "Expense" Then
                Me.cboBudgetType.SelectedIndex = 1
            ElseIf Me.grvMaster.SelectedRows(0).Cells(2).Value.ToString = "Asset" Then
                Me.cboBudgetType.SelectedIndex = 0
            Else
                Me.cboBudgetType.SelectedIndex = -1
            End If

            If Me.grvMaster.SelectedRows(0).Cells(6).Value.ToString = "FC" Then
                Me.cboCost.SelectedIndex = 0
            ElseIf Me.grvMaster.SelectedRows(0).Cells(6).Value.ToString = "Admin" Then
                Me.cboCost.SelectedIndex = 1
            Else
                Me.cboCost.SelectedIndex = -1
            End If
            'findItemValue(Me.cboCost, costValue)

            If Me.grvMaster.SelectedRows(0).Cells(5).Value.ToString = "Fixed Cost" Then
                Me.cboCostType.SelectedIndex = 0
            ElseIf Me.grvMaster.SelectedRows(0).Cells(5).Value.ToString = "Variable Cost" Then
                Me.cboCostType.SelectedIndex = 1
            Else
                Me.cboCostType.SelectedIndex = -1
            End If
            'findItemValue(Me.cboCostType, costTypeValue)

            findItemValue(Me.cboAssetGroup, Me.grvMaster.SelectedRows(0).Cells(7).Value.ToString)
            findItemValue(Me.cboDepartment, Me.grvMaster.SelectedRows(0).Cells(8).Value.ToString)
            findItemValue(Me.cboPersonInCharge, Me.grvMaster.SelectedRows(0).Cells(9).Value.ToString)

            If Me.grvMaster.SelectedRows(0).Cells(10).Value.ToString = "1" Then
                Me.chkActive.Checked = True
            Else
                Me.chkActive.Checked = False
            End If

            If Me.cboBudgetType.SelectedIndex = 1 Then
                If Me.grvMaster.SelectedRows(0).Cells(11).Value.ToString = "Labor Expense" Then
                    Me.cboExpenseType.SelectedIndex = 0
                ElseIf Me.grvMaster.SelectedRows(0).Cells(11).Value.ToString = "Variable Expense" Then
                    Me.cboExpenseType.SelectedIndex = 1
                ElseIf Me.grvMaster.SelectedRows(0).Cells(11).Value.ToString = "Fixed Expense" Then
                    Me.cboExpenseType.SelectedIndex = 2
                Else
                    Me.cboExpenseType.SelectedIndex = -1
                End If
                'findItemValue(Me.cboExpenseType, expenseTypeValue)
            Else
                Me.cboExpenseType.SelectedIndex = -1
            End If

            Me.txtRemarks.Text = Me.grvMaster.SelectedRows(0).Cells("colRemarks").Value.ToString

            SetBudgetTypeRelatedCtrl()
        End If
    End Sub

    Private Function requiredValueCheck() As Boolean
        If Me.txtOrderNo.Text.Trim = "" Then
            showErrorMessage("Please fill in budget order number")
            Me.txtOrderNo.Focus()

            Return False

        ElseIf Me.txtOrderName.Text.Trim = "" Then
            showErrorMessage("Please fill in budget order name")
            Me.txtOrderName.Focus()

            Return False

        ElseIf Me.cboBudgetType.SelectedIndex < 0 Then
            showErrorMessage("Please select budget type")
            Me.cboBudgetType.Focus()

            Return False

        ElseIf Me.txtCostCenter.Text.Trim = "" Then
            showErrorMessage("Please fill in cost center")
            Me.txtCostCenter.Focus()

            Return False

        ElseIf Me.cboBudgetType.Text.Equals(BGConstant.P_BUDGET_TYPE_EXPENSE) Then
            If Me.cboCostType.SelectedIndex < 0 Then
                showErrorMessage("Please select cost type")
                Me.cboCostType.Focus()

                Return False

            ElseIf Me.cboCost.SelectedIndex < 0 Then
                showErrorMessage("Please select cost")
                Me.cboCost.Focus()

                Return False

            ElseIf Me.cboExpenseType.SelectedIndex < 0 Then
                showErrorMessage("Please select expense type")
                Me.cboExpenseType.Focus()

                Return False
            End If

        ElseIf Me.cboBudgetType.Text.Equals(BGConstant.P_BUDGET_TYPE_ASSET) Then
            If Me.cboAssetGroup.SelectedIndex < 0 Then
                showErrorMessage("Please select asset group")
                Me.cboAssetGroup.Focus()

                Return False
            End If

        ElseIf Me.cboDepartment.SelectedIndex < 0 Then
            showErrorMessage("Please select department")
            Me.cboDepartment.Focus()

            Return False

        ElseIf Me.cboPersonInCharge.SelectedIndex < 0 Then
            showErrorMessage("Please select person in charge")
            Me.cboPersonInCharge.Focus()

            Return False
        End If

        Return True
    End Function

    Private Sub addDataColumn(ByVal dt As DataTable, ByVal dc As DataColumn)
        Dim colIndex As Integer = dt.Columns.IndexOf(dc)
        If colIndex < 0 Then
            dt.Columns.Add(dc)
        End If
    End Sub

    Private Sub saveData()
        If currentDS = "DB" Then

            Dim strBGOrderNo As String = Me.txtOrderNo.Text.Trim
            Dim rowIndex As Integer = -1

            myClsBG0620BL.BudgetOrderNo = Me.txtOrderNo.Text.Trim
            myClsBG0620BL.BudgetOrderName = Me.txtOrderName.Text.Trim
            myClsBG0620BL.BudgetType = Me.cboBudgetType.SelectedValue.ToString
            myClsBG0620BL.Account = Me.cboAccount.SelectedValue.ToString
            myClsBG0620BL.CostCenter = Me.txtCostCenter.Text.Trim

            If myClsBG0620BL.BudgetType = P_BUDGET_TYPE_EXPENSE Then
                If Me.cboCostType.SelectedValue IsNot Nothing Then
                    myClsBG0620BL.CostType = Me.cboCostType.SelectedValue.ToString
                Else
                    myClsBG0620BL.CostType = ""
                End If
                If Me.cboCost.SelectedValue IsNot Nothing Then
                    myClsBG0620BL.Cost = Me.cboCost.SelectedValue.ToString
                Else
                    myClsBG0620BL.Cost = ""
                End If
                If Me.cboExpenseType.SelectedValue IsNot Nothing Then
                    myClsBG0620BL.ExpenseType = Me.cboExpenseType.SelectedValue.ToString
                Else
                    myClsBG0620BL.ExpenseType = ""
                End If
                myClsBG0620BL.AssetGroup = ""

            Else
                myClsBG0620BL.CostType = ""
                myClsBG0620BL.Cost = ""
                myClsBG0620BL.ExpenseType = ""
                If Me.cboAssetGroup.SelectedValue IsNot Nothing Then
                    myClsBG0620BL.AssetGroup = Me.cboAssetGroup.SelectedValue.ToString
                Else
                    myClsBG0620BL.AssetGroup = ""
                End If
            End If

            myClsBG0620BL.Department = Me.cboDepartment.SelectedValue.ToString
            myClsBG0620BL.PersonInCharge = Me.cboPersonInCharge.SelectedValue.ToString

            If Me.chkActive.Checked Then
                myClsBG0620BL.ActiveFlag = "1"
            Else
                myClsBG0620BL.ActiveFlag = "0"
            End If
            myClsBG0620BL.CreateUserId = p_strUserId
            myClsBG0620BL.UpdateUserId = p_strUserId

            myClsBG0620BL.Remarks = Me.txtRemarks.Text.Trim

            '// Get the selected person in charge's PIC Show Flag
            myClsBG0620BL.getPICShowFlag()

            If myClsBG0620BL.saveBudgetOrder Then
                showSystemMessage("Budget order save complete.")

                '// Write Transaction Log
                WriteTransactionLog(CStr(enumOperationCd.EditBudgetOrderMaster), "", "", "", "", "", "")

            End If

            '// Remember current state
            Dim intFirstRow As Integer = grvMaster.FirstDisplayedCell.RowIndex
            Dim intFirstCol As Integer = grvMaster.FirstDisplayedCell.ColumnIndex
            Dim intSelRow As Integer = grvMaster.SelectedCells(0).RowIndex
            Dim intSelCol As Integer = grvMaster.SelectedCells(0).ColumnIndex

            clearData()

            Me.grvMaster.DataSource = Nothing

            ReadFilter()

            If myClsBG0620BL.getBudgetOrderList Then
                Me.Cursor = Cursors.WaitCursor
                Me.Enabled = False
                Me.dtMasterBudgetOrder = myClsBG0620BL.DtResult
                Me.grvMaster.DataSource = Me.dtMasterBudgetOrder

                Dim drRows() As DataRow
                drRows = Me.dtMasterBudgetOrder.Select("BUDGET_ORDER_NO = '" & strBGOrderNo & "'", "")
                If drRows.Length = 1 Then
                    rowIndex = Me.dtMasterBudgetOrder.Rows.IndexOf(drRows(0))
                    Me.grvMaster.Rows(rowIndex).Selected = True
                    displaySelectedItem()
                End If

                Me.Enabled = True
                Me.Cursor = Cursors.Default
            End If

            '// Select edited row
            If intFirstRow < grvMaster.Rows.Count Then
                If grvMaster.Item(intFirstCol, intFirstRow) IsNot Nothing Then
                    grvMaster.FirstDisplayedCell = grvMaster.Item(intFirstCol, intFirstRow)
                End If
            End If
            If intSelRow < grvMaster.Rows.Count Then
                If grvMaster.Item(intSelCol, intSelRow) IsNot Nothing Then
                    grvMaster.Item(intSelCol, intSelRow).Selected = True
                End If
            End If

        ElseIf currentDS = "XLS" Then

            Dim intSaved As Integer = 0
            Dim intTotal As Integer = Me.dtMasterBudgetOrder.Rows.Count

            If showConfirmMessage("This operation will save all budget order data " & vbCrLf & _
                                  "from the selected spreadsheet to the database" & vbCrLf & _
                                  "which has total " & intTotal.ToString & " item(s)." & vbCrLf & vbCrLf & _
                                  "Would you like to continue?") = Windows.Forms.DialogResult.Yes Then
                Me.Enabled = False
                Dim success As Boolean = False

                Dim conn As SqlConnection
                conn = New SqlConnection(My.Settings.ConnStr)
                Dim trans As SqlTransaction

                conn.Open()
                trans = conn.BeginTransaction()
                Try
                    For Each row As DataRow In Me.dtMasterBudgetOrder.Rows
                        myClsBG0620BL.BudgetOrderNo = row.Item("BUDGET_ORDER_NO").ToString
                        myClsBG0620BL.BudgetOrderName = row.Item("BUDGET_ORDER_NAME").ToString
                        myClsBG0620BL.BudgetType = row.Item("BUDGET_TYPE").ToString
                        myClsBG0620BL.Account = row.Item("ACCOUNT_NO").ToString
                        myClsBG0620BL.CostCenter = row.Item("COST_CENTER").ToString
                        If myClsBG0620BL.BudgetType = P_BUDGET_TYPE_EXPENSE Then
                            myClsBG0620BL.CostType = row.Item("COST_TYPE").ToString
                            myClsBG0620BL.Cost = row.Item("COST").ToString
                            myClsBG0620BL.ExpenseType = row.Item("EXPENSE_TYPE").ToString
                            myClsBG0620BL.AssetGroup = ""
                        Else
                            myClsBG0620BL.CostType = ""
                            myClsBG0620BL.Cost = ""
                            myClsBG0620BL.ExpenseType = ""
                            myClsBG0620BL.AssetGroup = row.Item("ASSET_GROUP_NO").ToString
                        End If
                        myClsBG0620BL.Department = row.Item("DEPT_NO").ToString
                        myClsBG0620BL.PersonInCharge = row.Item("PERSON_IN_CHARGE_NO").ToString
                        myClsBG0620BL.PICShowFlag = row.Item("PIC_SHOW_FLAG").ToString
                        If row.Item("CREATE_USER_ID").ToString = "" Then
                            myClsBG0620BL.CreateUserId = p_strUserId
                        Else
                            myClsBG0620BL.CreateUserId = row.Item("CREATE_USER_ID").ToString
                        End If
                        myClsBG0620BL.CreateDate = dateConvert(row.Item("CREATE_DATE").ToString)
                        myClsBG0620BL.ActiveFlag = row.Item("ACTIVE_FLAG").ToString
                        myClsBG0620BL.CreateUserId = p_strUserId
                        myClsBG0620BL.UpdateUserId = p_strUserId

                        myClsBG0620BL.Remarks = row.Item("REMARKS").ToString

                        If myClsBG0620BL.saveBudgetOrderAll(conn, trans) Then
                            intSaved += 1
                            success = True
                        Else
                            success = False
                            Exit For
                        End If
                    Next

                    If success = True Then
                        trans.Commit()
                    Else
                        trans.Rollback()
                    End If
                Catch ex As Exception

                Finally
                    conn.Close()
                End Try

                If intSaved > 0 Then
                    showSystemMessage("Spreadsheet data were imported and successfully saved to database.")

                    '// Write Transaction Log
                    WriteTransactionLog(CStr(enumOperationCd.EditBudgetOrderMaster), "", "", "", "", "", "")

                End If

                Me.grvMaster.DataSource = Nothing

                ReadFilter()

                If myClsBG0620BL.getBudgetOrderList Then
                    Me.Cursor = Cursors.WaitCursor
                    Me.Enabled = False
                    Me.dtMasterBudgetOrder = myClsBG0620BL.DtResult
                    Me.grvMaster.DataSource = Me.dtMasterBudgetOrder
                    Me.Cursor = Cursors.Default
                End If

                currentDS = "DB"
                Me.Enabled = True
            End If

        End If
    End Sub

    Private Sub SetBudgetTypeRelatedCtrl()

        If Me.cboBudgetType.Text = "Asset" Then

            Me.cboCostType.Enabled = False
            Me.lblCostType.Enabled = False
            Me.cboCostType.SelectedIndex = -1

            Me.cboCost.Enabled = False
            Me.lblCost.Enabled = False
            Me.cboCost.SelectedIndex = -1

            Me.cboExpenseType.Enabled = False
            Me.lblExpenseType.Enabled = False
            Me.cboExpenseType.SelectedIndex = -1

            Me.cboAssetGroup.Enabled = True
            Me.lblAssetGroup.Enabled = True
            If Me.cboAssetGroup.Items.Count > 0 And Me.cboAssetGroup.SelectedIndex = -1 Then
                Me.cboAssetGroup.SelectedIndex = 0
            End If

        ElseIf Me.cboBudgetType.Text = "Expense" Then

            Me.cboCostType.Enabled = True
            Me.lblCostType.Enabled = True
            If Me.cboCostType.Items.Count > 0 And Me.cboCostType.SelectedIndex = -1 Then
                Me.cboCostType.SelectedIndex = 0
            End If

            Me.cboCost.Enabled = True
            Me.lblCost.Enabled = True
            If Me.cboCost.Items.Count > 0 And Me.cboCost.SelectedIndex = -1 Then
                Me.cboCost.SelectedIndex = 0
            End If

            Me.cboExpenseType.Enabled = True
            Me.lblExpenseType.Enabled = True
            If Me.cboExpenseType.Items.Count > 0 And Me.cboExpenseType.SelectedIndex = -1 Then
                Me.cboExpenseType.SelectedIndex = 0
            End If

            Me.cboAssetGroup.Enabled = False
            Me.lblAssetGroup.Enabled = False
            Me.cboAssetGroup.SelectedIndex = -1

        End If
    End Sub

    Private Sub SetBudgetTypeRelatedCtrlFilter()

        If Me.cboBudgetTypeFilter.Text = "Asset" Then

            Me.cboCostTypeFilter.Enabled = False
            Me.cboCostTypeFilter.SelectedIndex = -1

            Me.cboCostFilter.Enabled = False
            Me.cboCostFilter.SelectedIndex = -1

            Me.cboExpenseTypeFilter.Enabled = False
            Me.cboExpenseTypeFilter.SelectedIndex = -1

            Me.cboAssetGroupFilter.Enabled = True
            If Me.cboAssetGroupFilter.Items.Count > 0 And Me.cboAssetGroupFilter.SelectedIndex = -1 Then
                Me.cboAssetGroupFilter.SelectedIndex = 0
            End If

        ElseIf Me.cboBudgetTypeFilter.Text = "Expense" Then

            Me.cboCostTypeFilter.Enabled = True
            If Me.cboCostTypeFilter.Items.Count > 0 And Me.cboCostTypeFilter.SelectedIndex = -1 Then
                Me.cboCostTypeFilter.SelectedIndex = 0
            End If

            Me.cboCostFilter.Enabled = True
            If Me.cboCostFilter.Items.Count > 0 And Me.cboCostFilter.SelectedIndex = -1 Then
                Me.cboCostFilter.SelectedIndex = 0
            End If

            Me.cboExpenseTypeFilter.Enabled = True
            If Me.cboExpenseTypeFilter.Items.Count > 0 And Me.cboExpenseTypeFilter.SelectedIndex = -1 Then
                Me.cboExpenseTypeFilter.SelectedIndex = 0
            End If

            Me.cboAssetGroupFilter.Enabled = False
            Me.cboAssetGroupFilter.SelectedIndex = -1

        Else
            Me.cboCostTypeFilter.Enabled = False
            Me.cboCostTypeFilter.SelectedIndex = -1

            Me.cboCostFilter.Enabled = False
            Me.cboCostFilter.SelectedIndex = -1

            Me.cboExpenseTypeFilter.Enabled = False
            Me.cboExpenseTypeFilter.SelectedIndex = -1

            Me.cboAssetGroupFilter.Enabled = False
            Me.cboAssetGroupFilter.SelectedIndex = -1
        End If
    End Sub

    Private Sub FilterData()

        Me.Cursor = Cursors.WaitCursor

        ReadFilter()

        If myClsBG0620BL.getBudgetOrderList Then
            currentDS = "DB"
            Me.dtMasterBudgetOrder = myClsBG0620BL.DtResult
            Me.grvMaster.DataSource = Me.dtMasterBudgetOrder
        End If
        displaySelectedItem()

        Me.Cursor = Cursors.Default

    End Sub

    Private Sub ReadFilter()

        myClsBG0620BL.BudgetOrderNoFilter = Me.txtOrderNoFilter.Text.Trim
        myClsBG0620BL.BudgetOrderNameFilter = Me.txtOrderNameFilter.Text.Trim

        If cboBudgetTypeFilter.SelectedValue IsNot Nothing Then
            myClsBG0620BL.BudgetTypeFilter = Me.cboBudgetTypeFilter.SelectedValue.ToString
        End If

        If cboAccountFilter.SelectedValue IsNot Nothing Then
            myClsBG0620BL.AccountFilter = Me.cboAccountFilter.SelectedValue.ToString
        End If

        myClsBG0620BL.CostCenterFilter = Me.txtCostCenterFilter.Text.Trim

        If myClsBG0620BL.BudgetTypeFilter = P_BUDGET_TYPE_EXPENSE Then
            If Me.cboCostTypeFilter.SelectedValue IsNot Nothing Then
                myClsBG0620BL.CostTypeFilter = Me.cboCostTypeFilter.SelectedValue.ToString
            Else
                myClsBG0620BL.CostTypeFilter = ""
            End If
            If Me.cboCostFilter.SelectedValue IsNot Nothing Then
                myClsBG0620BL.CostFilter = Me.cboCostFilter.SelectedValue.ToString
            Else
                myClsBG0620BL.CostFilter = ""
            End If
            If Me.cboExpenseTypeFilter.SelectedValue IsNot Nothing Then
                myClsBG0620BL.ExpenseTypeFilter = Me.cboExpenseTypeFilter.SelectedValue.ToString
            Else
                myClsBG0620BL.ExpenseTypeFilter = ""
            End If
            myClsBG0620BL.AssetGroupFilter = ""

        Else
            myClsBG0620BL.CostTypeFilter = ""
            myClsBG0620BL.CostFilter = ""
            myClsBG0620BL.ExpenseTypeFilter = ""
            If Me.cboAssetGroupFilter.SelectedValue IsNot Nothing Then
                myClsBG0620BL.AssetGroupFilter = Me.cboAssetGroupFilter.SelectedValue.ToString
            Else
                myClsBG0620BL.AssetGroupFilter = ""
            End If
        End If

        If Me.cboDepartmentFilter.SelectedValue IsNot Nothing Then
            myClsBG0620BL.DepartmentFilter = Me.cboDepartmentFilter.SelectedValue.ToString
        End If

        If Me.cboPersonInChargeFilter.SelectedValue IsNot Nothing Then
            myClsBG0620BL.PersonInChargeFilter = Me.cboPersonInChargeFilter.SelectedValue.ToString
        End If

        If Me.chkActiveFilter.Checked Then
            myClsBG0620BL.ActiveFlagFilter = "1"
        Else
            myClsBG0620BL.ActiveFlagFilter = "0"
        End If
    End Sub

#End Region

#Region "Control Event"

    Private Function dateConvert(ByVal strDate As String) As String
        Dim myDate As Date
        If strDate <> "" Then
            myDate = CDate(strDate)
        Else
            myDate = Date.Now
        End If

        Return myDate.ToString("yyyy-MM-dd HH:mm:ss")
    End Function

    Private Sub frmBG0620_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Cursor = Cursors.WaitCursor

        initFilter()
        init()
        clearData()

        ReadFilter()

        If myClsBG0620BL.getBudgetOrderList Then
            currentDS = "DB"
            Me.dtMasterBudgetOrder = myClsBG0620BL.DtResult
            Me.grvMaster.DataSource = Me.dtMasterBudgetOrder
        End If
        displaySelectedItem()

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub cboBudgetType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboBudgetType.SelectedIndexChanged
        If Me.cboBudgetType.SelectedIndex > -1 Then
            SetBudgetTypeRelatedCtrl()
        End If
    End Sub

    Private Sub grvMaster_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grvMaster.CellClick
        displaySelectedItem()
    End Sub

    Private Sub txtOrderNo_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOrderNo.Leave
        If Me.txtOrderNo.Text.Trim <> "" Or Me.txtOrderNo.Text.Trim <> String.Empty Then
            Dim drRows() As DataRow
            drRows = Me.dtMasterBudgetOrder.Select("BUDGET_ORDER_NO = '" & Me.txtOrderNo.Text.Trim & "'", "")
            If drRows.Length > 0 Then
                showErrorMessage("This budget order number was already in use. Please input another number.")
                Me.txtOrderNo.Focus()
            End If
        End If
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        clearData()
        Me.txtOrderNo.Focus()
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        If Not requiredValueCheck() Then Return

        saveData()
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        If Not requiredValueCheck() Then Return

        '// Calc total value of budget data 
        Dim dblTotal As Double = 0
        Dim clsBG_T_BUDGET_DATA As New BG_T_BUDGET_DATA

        clsBG_T_BUDGET_DATA.BudgetOrderNo = Me.txtOrderNo.Text.Trim

        If clsBG_T_BUDGET_DATA.Select022() = True AndAlso clsBG_T_BUDGET_DATA.dtResult.Rows.Count > 0 Then
            dblTotal = CDbl(Nz(clsBG_T_BUDGET_DATA.dtResult.Rows(0)![Total], 0))
        End If

        If showConfirmMessage("Total value of budget data which reference with this budget order is " & _
                              dblTotal.ToString("#,##0.00") & "K baht." & vbNewLine & _
                              "Are you sure to delete the budget order?") = Windows.Forms.DialogResult.Yes Then

            '// Remember current state
            Dim intFirstRow As Integer = grvMaster.FirstDisplayedCell.RowIndex
            Dim intFirstCol As Integer = grvMaster.FirstDisplayedCell.ColumnIndex
            Dim intSelRow As Integer = grvMaster.SelectedCells(0).RowIndex
            Dim intSelCol As Integer = grvMaster.SelectedCells(0).ColumnIndex

            myClsBG0620BL.BudgetOrderNo = Me.txtOrderNo.Text.Trim
            If myClsBG0620BL.deleteBudgetOrder Then
                showSystemMessage("Delete completed.")

                '// Write Transaction Log
                WriteTransactionLog(CStr(enumOperationCd.EditBudgetOrderMaster), "", "", "", "", "", "")

            End If

            clearData()

            Me.grvMaster.DataSource = Nothing

            ReadFilter()

            If myClsBG0620BL.getBudgetOrderList Then
                Me.Enabled = False
                Me.dtMasterBudgetOrder = myClsBG0620BL.DtResult
                Me.grvMaster.DataSource = Me.dtMasterBudgetOrder
                Me.Enabled = True
                displaySelectedItem()
            End If

            '// Select edited row
            If intFirstRow < grvMaster.Rows.Count Then
                If grvMaster.Item(intFirstCol, intFirstRow) IsNot Nothing Then
                    grvMaster.FirstDisplayedCell = grvMaster.Item(intFirstCol, intFirstRow)
                End If
            End If
            If intSelRow < grvMaster.Rows.Count Then
                If grvMaster.Item(intSelCol, intSelRow) IsNot Nothing Then
                    grvMaster.Item(intSelCol, intSelRow).Selected = True
                End If
            End If

        End If
    End Sub

    Private Sub cmdExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExport.Click
        If Me.grvMaster.Columns.Count = 0 Or Me.grvMaster.Rows.Count = 0 Then
            Exit Sub
        End If

        '// Show dialog box
        Dim sdlgSave As SaveFileDialog = New SaveFileDialog
        sdlgSave.FileName = "BudgetOrderMaster_" & Format(Date.Now, "yyyyMMdd")
        sdlgSave.Filter = "Microsoft Excel Workbook (*.xls)|*.xls"

        Dim dlrConfirm As DialogResult = sdlgSave.ShowDialog()
        If dlrConfirm.Equals(DialogResult.Cancel) Then
            Exit Sub
        End If

        Me.Enabled = False

        Dim excel As New Microsoft.Office.Interop.Excel.Application
        Dim wBook As Microsoft.Office.Interop.Excel.Workbook
        Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet

        wBook = excel.Workbooks.Add()
        wSheet = CType(wBook.ActiveSheet(), Microsoft.Office.Interop.Excel.Worksheet)

        Dim dt As System.Data.DataTable = Me.dtMasterBudgetOrder
        Dim dc As System.Data.DataColumn
        Dim dr As System.Data.DataRow
        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = 0

        For Each dc In dt.Columns
            colIndex = colIndex + 1
            excel.Cells(1, colIndex) = dc.ColumnName
        Next

        For Each dr In dt.Rows
            rowIndex = rowIndex + 1
            colIndex = 0
            For Each dc In dt.Columns
                colIndex = colIndex + 1
                If dc.ColumnName = "BUDGET_TYPE" Then
                    If dr(dc.ColumnName).ToString = "Expense" Then
                        excel.Cells(rowIndex + 1, colIndex) = "E"
                    ElseIf dr(dc.ColumnName).ToString = "Asset" Then
                        excel.Cells(rowIndex + 1, colIndex) = "A"
                    Else
                        excel.Cells(rowIndex + 1, colIndex) = ""
                    End If
                ElseIf dc.ColumnName = "EXPENSE_TYPE" Then
                    If dr(dc.ColumnName).ToString = "Labor Expense" Then
                        excel.Cells(rowIndex + 1, colIndex) = "1"
                    ElseIf dr(dc.ColumnName).ToString = "Variable Expense" Then
                        excel.Cells(rowIndex + 1, colIndex) = "2"
                    ElseIf dr(dc.ColumnName).ToString = "Fixed Expense" Then
                        excel.Cells(rowIndex + 1, colIndex) = "3"
                    Else
                        excel.Cells(rowIndex + 1, colIndex) = ""
                    End If
                ElseIf dc.ColumnName = "COST_TYPE" Then
                    If dr(dc.ColumnName).ToString = "Fixed Cost" Then
                        excel.Cells(rowIndex + 1, colIndex) = "1"
                    ElseIf dr(dc.ColumnName).ToString = "Variable Cost" Then
                        excel.Cells(rowIndex + 1, colIndex) = "2"
                    Else
                        excel.Cells(rowIndex + 1, colIndex) = ""
                    End If
                ElseIf dc.ColumnName = "COST" Then
                    If dr(dc.ColumnName).ToString = "Admin" Then
                        excel.Cells(rowIndex + 1, colIndex) = CStr(enumCost.ADMIN)
                    ElseIf dr(dc.ColumnName).ToString = "FC" Then
                        excel.Cells(rowIndex + 1, colIndex) = CStr(enumCost.FC)
                    Else
                        excel.Cells(rowIndex + 1, colIndex) = ""
                    End If
                Else
                    excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
                End If
                excel.ActiveCell.NumberFormat = "@"
            Next
        Next

        wSheet.Columns.AutoFit()
        Dim strFileName As String = sdlgSave.FileName
        Dim blnFileOpen As Boolean = False
        Try
            Dim fileTemp As System.IO.FileStream = System.IO.File.OpenWrite(strFileName)
            fileTemp.Close()
        Catch ex As Exception
            blnFileOpen = False
        End Try

        If System.IO.File.Exists(strFileName) Then
            System.IO.File.Delete(strFileName)
        End If

        wBook.SaveAs(strFileName)
        excel.Workbooks.Open(strFileName)
        excel.Visible = True

        '//Release memory
        BGCommon.ExcelReleasememory(excel, wBook, wSheet)

        Me.Enabled = True

        ''showSystemMessage("Spreadsheet export successful")
    End Sub

    Private Sub cmdImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdImport.Click
        If Me.grvMaster.Columns.Count = 0 Or Me.grvMaster.Rows.Count = 0 Then
            Exit Sub
        End If

        Dim strFileName As String = String.Empty

        'Show dialog box
        Dim sdlgOpen As OpenFileDialog = New OpenFileDialog
        sdlgOpen.FileName = ""
        sdlgOpen.Filter = "Microsoft Excel Workbook (*.xls)|*.xls"

        Dim dlrConfirm As DialogResult = sdlgOpen.ShowDialog()
        If dlrConfirm.Equals(DialogResult.Cancel) Then
            Exit Sub
        End If

        Me.Enabled = False
        If sdlgOpen.FileName.Trim.Equals("") Then
            Exit Sub 'return
        End If

        Me.Cursor = Cursors.WaitCursor

        strFileName = Path.GetFullPath(sdlgOpen.FileName)

        Dim xConnStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                                    "Data Source=" & strFileName & ";" & _
                                    "Extended Properties=""Excel 12.0;HDR=Yes;ImportMixedTypes=Text;IMEX=1"";"
        Dim xConn As New OleDbConnection(xConnStr)
        Try
            xConn.Open()

            Dim xDA As New OleDbDataAdapter("SELECT * FROM [Sheet1$] WHERE BUDGET_ORDER_NO <> ''", xConn)
            Me.dtMasterBudgetOrder.Clear()
            Me.grvMaster.DataSource = Nothing
            xDA.Fill(Me.dtMasterBudgetOrder)

            'If Me.dtMasterBudgetOrder.Rows.Count > 0 Then
            '    For Each row As DataRow In Me.dtMasterBudgetOrder.Rows
            '        For Each col As DataColumn In Me.dtMasterBudgetOrder.Columns
            '            If col.ColumnName = "BUDGET_TYPE" Then
            '                If row(col.ColumnName).ToString = "E" Then
            '                    row(col.ColumnName) = "Expense"
            '                ElseIf row(col.ColumnName).ToString = "A" Then
            '                    row(col.ColumnName) = "Asset"
            '                Else
            '                    row(col.ColumnName) = ""
            '                End If
            '            ElseIf col.ColumnName = "EXPENSE_TYPE" Then
            '                If row(col.ColumnName).ToString = "1" Then
            '                    row(col.ColumnName) = "Labor Expense"
            '                ElseIf row(col.ColumnName).ToString = "2" Then
            '                    row(col.ColumnName) = "Variable Expense"
            '                ElseIf row(col.ColumnName).ToString = "3" Then
            '                    row(col.ColumnName) = "Fixed Expense"
            '                Else
            '                    row(col.ColumnName) = ""
            '                End If
            '            ElseIf col.ColumnName = "COST_TYPE" Then
            '                If row(col.ColumnName).ToString = "1" Then
            '                    row(col.ColumnName) = "Fixed Cost"
            '                ElseIf row(col.ColumnName).ToString = "2" Then
            '                    row(col.ColumnName) = "Variable Cost"
            '                Else
            '                    row(col.ColumnName) = ""
            '                End If
            '            ElseIf col.ColumnName = "COST" Then
            '                If row(col.ColumnName).ToString = "1" Then
            '                    row(col.ColumnName) = "Admin"
            '                ElseIf row(col.ColumnName).ToString = "2" Then
            '                    row(col.ColumnName) = "FC"
            '                Else
            '                    row(col.ColumnName) = ""
            '                End If
            '            End If
            '        Next
            '    Next
            'End If

            'If Me.dtMasterBudgetOrder.Rows.Count > 0 Then
            '    Me.grvMaster.DataSource = Me.dtMasterBudgetOrder
            'End If

            'showSystemMessage("Import from spreadsheet complete")

            displaySelectedItem()

            currentDS = "XLS"

        Catch ex As Exception
            showErrorMessage("Error importing from spreadsheet." & vbCrLf & ex.Message)
        Finally
            Me.Enabled = True
            xConn.Close()
        End Try

        saveData()

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub cmdFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFilter.Click
        FilterData()
    End Sub

    Private Sub cboBudgetTypeFilter_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboBudgetTypeFilter.SelectedIndexChanged
        If Me.cboBudgetTypeFilter.SelectedIndex > -1 Then
            SetBudgetTypeRelatedCtrlFilter()
        End If
    End Sub

    Private Sub cmdClearFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearFilter.Click
        clearFilter()
    End Sub
#End Region

End Class