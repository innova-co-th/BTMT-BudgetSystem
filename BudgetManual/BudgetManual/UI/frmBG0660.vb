Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.IO
Imports System.Data.OleDb

Public Class frmBG0660

#Region "Variable"
    Private myClsBG0660BL As New clsBG0660BL
    Private myClsBG0310BL As New clsBG0310BL
    Private dtTransferType As New DataTable
    Private dtTransferTypeFilter As New DataTable
    Private dtPeriodType As New DataTable
    Private dtOrder2 As DataTable
    Private dtOrder3 As DataTable
    Private fullLoad As Boolean = False
    Private dtAccount As DataTable
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

    Private Sub addDataColumn(ByVal dt As DataTable, ByVal dc As DataColumn)
        Dim colIndex As Integer = dt.Columns.IndexOf(dc)
        If colIndex < 0 Then
            dt.Columns.Add(dc)
        End If
    End Sub

    Private Sub LoadPeriodType()
        Try
            Me.cboPeriodType.Items.Clear()

            myClsBG0310BL.OpenPeriodFlg = "1"
            myClsBG0310BL.GetOpenPeriodList()

            If myClsBG0310BL.PeriodList IsNot Nothing AndAlso myClsBG0310BL.PeriodList.Rows.Count > 0 Then
                cboPeriodType.DisplayMember = "PERIOD_TYPE_NAME"
                cboPeriodType.ValueMember = "PERIOD_TYPE_ID"
                cboPeriodType.DataSource = myClsBG0310BL.PeriodList

                cboPeriodType.SelectedIndex = 0
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Detail by Account Code Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub setComboBoxData()
        '   Dim periodType As BGConstant.enumPeriodType
        Dim transferType As BGConstant.enumTransferType
        Dim drNew As DataRow

        If myClsBG0660BL.getBGOrderList Then
            dtOrder3 = myClsBG0660BL.DTResult.Clone
            dtOrder2 = myClsBG0660BL.DTResult.Clone

            For Each row As DataRow In myClsBG0660BL.DTResult.Rows
                dtOrder2.ImportRow(row)
                dtOrder3.ImportRow(row)
            Next

            fillComboBox(Me.cboBGOrderNo, myClsBG0660BL.DTResult, "BUDGET_ORDER_NO", "BUDGET_TEXT", False)
            fillComboBox(Me.cboFromOrderNo, dtOrder2, "BUDGET_ORDER_NO", "BUDGET_TEXT", False)
            fillComboBox(Me.cboToOrderNo, dtOrder3, "BUDGET_ORDER_NO", "BUDGET_TEXT", False)
        End If

        '// cboTransferType
        addDataColumn(dtTransferType, New DataColumn("KEY"))
        addDataColumn(dtTransferType, New DataColumn("VALUE"))
        For Each transferType In [Enum].GetValues(GetType(BGConstant.enumTransferType))
            drNew = dtTransferType.NewRow
            drNew("KEY") = CStr(transferType)
            Select Case transferType
                Case enumTransferType.FCtoADMIN
                    drNew("VALUE") = "FC to Admin"
                Case enumTransferType.ADMINtoFC
                    drNew("VALUE") = "Admin to FC"
            End Select
            dtTransferType.Rows.Add(drNew)
        Next
        fillComboBox(Me.cboTransferType, dtTransferType, "KEY", "VALUE", False)

        '// cboTransferTypeFilter
        addDataColumn(dtTransferTypeFilter, New DataColumn("KEY"))
        addDataColumn(dtTransferTypeFilter, New DataColumn("VALUE"))
        For Each transferType In [Enum].GetValues(GetType(BGConstant.enumTransferType))
            drNew = dtTransferTypeFilter.NewRow
            drNew("KEY") = CStr(transferType)
            Select Case transferType
                Case enumTransferType.FCtoADMIN
                    drNew("VALUE") = "FC to Admin"
                Case enumTransferType.ADMINtoFC
                    drNew("VALUE") = "Admin to FC"
            End Select
            dtTransferTypeFilter.Rows.Add(drNew)
        Next
        fillComboBox(Me.cboTransferTypeFilter, dtTransferTypeFilter, "KEY", "VALUE", True)

        LoadPeriodType()
        'addDataColumn(dtPeriodType, New DataColumn("KEY"))
        'addDataColumn(dtPeriodType, New DataColumn("VALUE"))
        'For Each periodType In [Enum].GetValues(GetType(BGConstant.enumPeriodType))
        '    drNew = dtPeriodType.NewRow
        '    Dim pType As String = CStr(periodType)
        '    drNew("KEY") = pType
        '    Select Case periodType
        '        Case enumPeriodType.OriginalBudget
        '            drNew("VALUE") = "Original Budget"
        '        Case enumPeriodType.EstimateBudget
        '            drNew("VALUE") = "Estimate Budget"
        '        Case enumPeriodType.ReviseBudget
        '            drNew("VALUE") = "Revise Budget"
        '    End Select
        '    dtPeriodType.Rows.Add(drNew)
        'Next
        'fillComboBox(Me.cboPeriodType, dtPeriodType, "KEY", "VALUE", False)

        If myClsBG0660BL.getAccountList Then
            dtAccount = myClsBG0660BL.DTResult
        End If
        fillComboBox(Me.cboAccountNo, dtAccount, "ACCOUNT_NO", "ACCOUNT_NAME_2", False)
    End Sub

    Private Sub loadGridData()
        Me.dtTCS.Clear()
        myClsBG0660BL.BudgetYear = Me.numYear.Value.ToString
        myClsBG0660BL.PeriodType = Me.cboPeriodType.SelectedValue.ToString
        myClsBG0660BL.ProjectNo = Me.numProjectNo.Value.ToString

        '//Filter
        myClsBG0660BL.BudgetOrderNoFilter = Me.txtOrderNoFilter.Text.Trim
        myClsBG0660BL.BudgetOrderNameFilter = Me.txtOrderNameFilter.Text.Trim
        myClsBG0660BL.AccountNoFilter = Me.txtAccountNoFilter.Text.Trim
        myClsBG0660BL.AccountNameFilter = Me.txtAccountNameFilter.Text.Trim
        myClsBG0660BL.TransferTypeFilter = Me.cboTransferTypeFilter.SelectedValue.ToString
        myClsBG0660BL.FromOrderNoFilter = Me.txtFromOrderNoFilter.Text.Trim
        myClsBG0660BL.ToOrderNoFilter = Me.txtToOrderNoFilter.Text.Trim

        If myClsBG0660BL.getDataList() Then
            Me.dtTCS = myClsBG0660BL.DTResult
            Me.grvMaster.DataSource = Me.dtTCS
        End If
        displaySelectedItem()
    End Sub

    Private Sub displaySelectedItem()
        Me.optAddByAccountNo.Enabled = False

        Me.optAddByOrderNo.Checked = True
        Me.cboBGOrderNo.Enabled = True

        Me.optAddByAccountNo.Checked = False
        Me.cboAccountNo.Enabled = False

        If Me.grvMaster.RowCount > 0 And Me.grvMaster.SelectedRows.Count = 1 Then
            findItemValue(Me.cboBGOrderNo, Me.grvMaster.SelectedRows(0).Cells(3).Value.ToString)
            findItemValue(Me.cboAccountNo, Me.grvMaster.SelectedRows(0).Cells(5).Value.ToString)
            findItemValue(Me.cboFromOrderNo, Me.grvMaster.SelectedRows(0).Cells(11).Value.ToString)
            findItemValue(Me.cboToOrderNo, Me.grvMaster.SelectedRows(0).Cells(14).Value.ToString)
            Me.cboBGOrderNo.Enabled = False

            If Me.grvMaster.SelectedRows(0).Cells(9).Value.ToString = "2" Then
                Me.cboTransferType.SelectedIndex = 1
            Else
                Me.cboTransferType.SelectedIndex = 0
            End If

            Me.txtTransferRate.Text = Me.grvMaster.SelectedRows(0).Cells(10).Value.ToString
        End If
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

    Private Sub clearData()
        Me.optAddByAccountNo.Enabled = True

        Me.optAddByOrderNo.Checked = True
        Me.cboBGOrderNo.Enabled = True
        Me.cboBGOrderNo.SelectedIndex = 0

        Me.optAddByAccountNo.Checked = False
        Me.cboAccountNo.Enabled = False
        Me.cboAccountNo.SelectedIndex = 0

        Me.cboFromOrderNo.SelectedIndex = 0
        Me.cboToOrderNo.SelectedIndex = 0
        Me.cboTransferType.SelectedIndex = 0
        Me.txtTransferRate.Text = ""
    End Sub

    Private Function checkData() As Boolean
        If optAddByOrderNo.Checked And Me.cboBGOrderNo.SelectedIndex < 0 Then
            showErrorMessage("Please select budger order number")
            Me.cboBGOrderNo.Focus()
            Return False
        ElseIf optAddByAccountNo.Checked And Me.cboAccountNo.SelectedIndex < 0 Then
            showErrorMessage("Please select account number")
            Me.cboAccountNo.Focus()
            Return False
        ElseIf Me.cboTransferType.SelectedIndex < 0 Then
            showErrorMessage("Please select transfer type")
            Me.cboTransferType.Focus()
            Return False
        ElseIf Me.txtTransferRate.Text = "" Then
            showErrorMessage("Please fill transfer rate")
            Me.txtTransferRate.Focus()
            Return False
        ElseIf Me.cboFromOrderNo.SelectedIndex < 0 Then
            showErrorMessage("Please select budget order number (From)")
            Me.cboFromOrderNo.Focus()
            Return False
        ElseIf Me.cboToOrderNo.SelectedIndex < 0 Then
            showErrorMessage("Please select budget order number (To)")
            Me.cboToOrderNo.Focus()
            Return False
        End If

        Return True
    End Function

    Private Function saveData(ByVal flag As String) As Boolean
        If Not checkData() Then

            Return False
        End If

        If flag = "save" And optAddByAccountNo.Checked Then
            myClsBG0660BL.AccountNo = Me.cboAccountNo.SelectedValue.ToString

            If myClsBG0660BL.getOrderByAccount Then
                For Each dr As DataRow In myClsBG0660BL.DTResult.Rows

                    myClsBG0660BL.BudgetOrderNo = CStr(dr("BUDGET_ORDER_NO"))
                    myClsBG0660BL.BudgetYear = Me.numYear.Value.ToString
                    myClsBG0660BL.PeriodType = Me.cboPeriodType.SelectedValue.ToString
                    myClsBG0660BL.TransferType = Me.cboTransferType.SelectedValue.ToString
                    myClsBG0660BL.TransferRate = Me.txtTransferRate.Text.Trim
                    myClsBG0660BL.FromOrderNo = Me.cboFromOrderNo.SelectedValue.ToString
                    myClsBG0660BL.ToOrderNo = Me.cboToOrderNo.SelectedValue.ToString
                    myClsBG0660BL.CreateUserID = p_strUserId
                    myClsBG0660BL.UpdateUserID = p_strUserId
                    myClsBG0660BL.ProjectNo = Me.numProjectNo.Value.ToString

                    myClsBG0660BL.saveIndividualData()
                Next

                showSystemMessage("Data successfully saved.")

                '// Write Transaction Log
                WriteTransactionLog(CStr(enumOperationCd.EditTransferCostMaster), "", "", "", "", "", "")

            End If

        Else
            myClsBG0660BL.BudgetOrderNo = Me.cboBGOrderNo.SelectedValue.ToString
            myClsBG0660BL.BudgetYear = Me.numYear.Value.ToString
            myClsBG0660BL.PeriodType = Me.cboPeriodType.SelectedValue.ToString
            myClsBG0660BL.TransferType = Me.cboTransferType.SelectedValue.ToString
            myClsBG0660BL.TransferRate = Me.txtTransferRate.Text.Trim
            myClsBG0660BL.FromOrderNo = Me.cboFromOrderNo.SelectedValue.ToString
            myClsBG0660BL.ToOrderNo = Me.cboToOrderNo.SelectedValue.ToString
            myClsBG0660BL.CreateUserID = p_strUserId
            myClsBG0660BL.UpdateUserID = p_strUserId
            myClsBG0660BL.ProjectNo = Me.numProjectNo.Value.ToString

            If flag = "save" Then
                If myClsBG0660BL.saveIndividualData() Then
                    showSystemMessage("Data successfully saved.")

                    '// Write Transaction Log
                    WriteTransactionLog(CStr(enumOperationCd.EditTransferCostMaster), "", "", "", "", "", "")

                End If
            Else
                If myClsBG0660BL.deleteData() Then
                    showSystemMessage("Data successfully deleted.")

                    '// Write Transaction Log
                    WriteTransactionLog(CStr(enumOperationCd.EditTransferCostMaster), "", "", "", "", "", "")

                End If
            End If
        End If

        '// Remember current state
        Dim intFirstRow As Integer
        Dim intFirstCol As Integer
        Dim intSelRow As Integer
        If grvMaster.FirstDisplayedCell IsNot Nothing Then
            intFirstRow = grvMaster.FirstDisplayedCell.RowIndex
            intFirstCol = grvMaster.FirstDisplayedCell.ColumnIndex
        End If
        If grvMaster.SelectedRows.Count > 0 Then
            intSelRow = grvMaster.SelectedRows(0).Index
        End If

        loadGridData()

        '// Select edited row
        If intFirstRow < grvMaster.Rows.Count And intFirstRow > 0 Then
            If grvMaster.Item(intFirstCol, intFirstRow) IsNot Nothing Then
                grvMaster.FirstDisplayedCell = grvMaster.Item(intFirstCol, intFirstRow)
            End If
        End If
        If intSelRow < grvMaster.Rows.Count Then
            If grvMaster.Rows(intSelRow) IsNot Nothing Then
                grvMaster.Rows(intSelRow).Selected = True
            End If
        End If

        Return True
    End Function

    Private Sub clearFilter()
        Me.txtOrderNoFilter.Text = ""
        Me.txtOrderNameFilter.Text = ""
        Me.txtAccountNoFilter.Text = ""
        Me.txtAccountNameFilter.Text = ""
        Me.cboTransferTypeFilter.SelectedIndex = 0
        Me.txtFromOrderNoFilter.Text = ""
        Me.txtToOrderNoFilter.Text = ""
    End Sub

#End Region

#Region "Control Event"

    Private Sub frmBG0660_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.numYear.Value = Now.Year
        setComboBoxData()
        clearData()
        fullLoad = True
        loadGridData()
    End Sub

    Private Sub numYear_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numYear.ValueChanged
        If Me.numYear.Value.ToString <> "" AndAlso _
            Me.cboPeriodType.SelectedIndex > -1 AndAlso _
            Me.numProjectNo.Value.ToString <> "" Then
            If fullLoad Then
                loadGridData()
            End If
        End If
    End Sub

    Private Sub cboPeriodType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPeriodType.SelectedIndexChanged
        If Me.numYear.Value.ToString <> "" AndAlso _
        Me.cboPeriodType.SelectedIndex > -1 AndAlso _
        Me.numProjectNo.Value.ToString <> "" Then
            If fullLoad Then
                loadGridData()
            End If
        End If
    End Sub

    Private Sub numProjectNo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numProjectNo.ValueChanged
        If Me.numYear.Value.ToString <> "" AndAlso _
        Me.cboPeriodType.SelectedIndex > -1 AndAlso _
        Me.numProjectNo.Value.ToString <> "" Then
            If fullLoad Then
                loadGridData()
            End If
        End If
    End Sub

    Private Sub grvMaster_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grvMaster.CellClick
        displaySelectedItem()
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        If saveData("save") = True Then

            '// Do you want to adjust budget data with this setting now?
            If MessageBox.Show("Do you want to [Re-Calculate] budget data with this Transfer Cost now?", Me.Text, MessageBoxButtons.YesNo, _
                       MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then

                Dim clsBG0200BL As New clsBG0200BL
                clsBG0200BL.BudgetYear = CStr(numYear.Value)
                clsBG0200BL.PeriodType = CStr(cboPeriodType.SelectedValue)
                clsBG0200BL.ProjectNo = numProjectNo.Value.ToString
                clsBG0200BL.BudgetType = BGConstant.P_BUDGET_TYPE_EXPENSE

                If clsBG0200BL.AdjustTransferCost() = True Then
                    MessageBox.Show("The budget data was adjusted", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    'MessageBox.Show("Can not adjust the budget data", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    MessageBox.Show("Not found budget data to adjust. (Only Rev No.2 or above can be adjusted)", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

            End If

        End If
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        clearData()
        Me.cboBGOrderNo.Focus()
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        If showConfirmMessage("Do you want to delete selected transfer data?") = Windows.Forms.DialogResult.Yes Then
            saveData("delete")
        End If
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub cmdImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdImport.Click
        Dim strFileName As String = String.Empty
        Dim tempDT As New DataTable

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
            Return
        End If

        strFileName = Path.GetFullPath(sdlgOpen.FileName)

        Dim xConnStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                                    "Data Source=" & strFileName & ";" & _
                                    "Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"";"
        Dim xConn As New OleDbConnection(xConnStr)
        Try
            xConn.Open()

            Dim xDA As New OleDbDataAdapter("SELECT * FROM [Sheet1$] WHERE BUDGET_ORDER_NO <> ''", xConn)
            tempDT.Clear()
            Me.grvMaster.DataSource = Nothing
            xDA.Fill(tempDT)
        Catch ex As Exception
            showErrorMessage("Error importing from spreadsheet." & vbCrLf & ex.Message)
        Finally
            Me.Enabled = True
            xConn.Close()
        End Try

        If tempDT.Rows.Count > 0 Then
            If showConfirmMessage("Do you want to import data from selected spreadsheet?") = Windows.Forms.DialogResult.Yes Then
                Dim budgetYear As String = tempDT.Rows(0).Item(0).ToString
                Dim periodType As String = tempDT.Rows(0).Item(1).ToString
                Dim projectNo As String = tempDT.Rows(0).Item(2).ToString

                Select Case periodType
                    Case "1"
                        periodType = "Original Budget"
                    Case "2"
                        periodType = "Estimate Budget"
                    Case "3"
                        periodType = "Revise Budget"
                    Case "10"
                        periodType = "MTP Budget"
                End Select

                If myClsBG0660BL.saveImportData(tempDT) Then
                    showSystemMessage("Spreadsheet data successfully imported to database" & vbCrLf & _
                                      "in year " & budgetYear & " period type [" & periodType & "] and Project No." & projectNo)

                    '// Write Transaction Log
                    WriteTransactionLog(CStr(enumOperationCd.EditTransferCostMaster), "", "", "", "", "", "")

                    Me.numYear.Value = CDec(budgetYear)
                    findItemValue(Me.cboPeriodType, periodType)
                    Me.numProjectNo.Value = CDec(projectNo)
                    loadGridData()
                End If
            End If
        Else
            showSystemMessage("No item was found in spreadsheet. No data were imported.")
        End If
        tempDT.Dispose()
    End Sub

    Private Sub cmdExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExport.Click
        If Me.dtTCS.Rows.Count = 0 Then
            showSystemMessage("No item was found on gridview. No data were exported.")
            Return
        End If

        'Show dialog box
        Dim sdlgSave As SaveFileDialog = New SaveFileDialog
        sdlgSave.FileName = "TransferCostMaster_" & Format(Date.Now, "yyyyMMdd")
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

        Dim dt As System.Data.DataTable = Me.dtTCS
        Dim dc As System.Data.DataColumn
        Dim dr As System.Data.DataRow
        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = 0

        For Each dc In dt.Columns

            If dc.ColumnName.Equals("BUDGET_YEAR") Or _
                dc.ColumnName.Equals("PERIOD_TYPE") Or _
                dc.ColumnName.Equals("PROJECT_NO") Or _
                dc.ColumnName.Equals("BUDGET_ORDER_NO") Or _
                dc.ColumnName.Equals("TRANSFER_TYPE") Or _
                dc.ColumnName.Equals("TRANSFER_RATE") Or _
                dc.ColumnName.Equals("FROM_ORDER_NO") Or _
                dc.ColumnName.Equals("TO_ORDER_NO") Or _
                dc.ColumnName.Equals("CREATE_USER_ID") Or _
                dc.ColumnName.Equals("CREATE_DATE") Or _
                dc.ColumnName.Equals("UPDATE_USER_ID") Or _
                dc.ColumnName.Equals("UPDATE_DATE") Then
                colIndex = colIndex + 1
                excel.Cells(1, colIndex) = dc.ColumnName
            End If

        Next

        For Each dr In dt.Rows
            rowIndex = rowIndex + 1
            colIndex = 0
            For Each dc In dt.Columns
                If dc.ColumnName.Equals("BUDGET_YEAR") Or _
                dc.ColumnName.Equals("PERIOD_TYPE") Or _
                dc.ColumnName.Equals("PROJECT_NO") Or _
                dc.ColumnName.Equals("BUDGET_ORDER_NO") Or _
                dc.ColumnName.Equals("TRANSFER_TYPE") Or _
                dc.ColumnName.Equals("TRANSFER_RATE") Or _
                dc.ColumnName.Equals("FROM_ORDER_NO") Or _
                dc.ColumnName.Equals("TO_ORDER_NO") Or _
                dc.ColumnName.Equals("CREATE_USER_ID") Or _
                dc.ColumnName.Equals("CREATE_DATE") Or _
                dc.ColumnName.Equals("UPDATE_USER_ID") Or _
                dc.ColumnName.Equals("UPDATE_DATE") Then
                    colIndex = colIndex + 1
                    excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
                End If
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

    Private Sub optAddByOrderNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optAddByOrderNo.Click
        If optAddByOrderNo.Checked Then
            cboAccountNo.Enabled = False
            cboBGOrderNo.Enabled = True
        End If
    End Sub

    Private Sub optAddByAccountNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optAddByAccountNo.Click
        If optAddByAccountNo.Checked Then
            cboAccountNo.Enabled = True
            cboBGOrderNo.Enabled = False
        End If
    End Sub

    Private Sub cmdFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFilter.Click
        loadGridData()
    End Sub

    Private Sub cmdClearFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearFilter.Click
        clearFilter()
    End Sub

#End Region

End Class