Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports CrystalDecisions.CrystalReports.Engine
Imports Microsoft.Office.Interop
Imports System.Drawing.Printing

Public Class frmBG0473

#Region "Variable"
    Const AC_M_COL As String = "AC_M"
    Const OB_M_COL As String = "OB_M"
    Const AC_H_COL As String = "ACC_AC_M"
    Const OB_H_COL As String = "ACC_OB_M"
    Const WB_M_COL As String = "WB_M"
    Const WB_H_COL As String = "ACC_WB_M"

    Private myClsBG0473BL As New clsBG0473BL
    Private myClsBG0310BL As New clsBG0310BL
    Private clsBG0400 As frmBG0400
    Private excelApp As Excel.Application
    Private missing As Object = System.Reflection.Missing.Value
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

    Private Function InitPage() As Boolean

        Try
            Me.numYear.Value = Now.Year

            Me.cboMonth.Text = Now.Month.ToString

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error " & Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Function

    Private Sub Print(ByVal blnShowPrintPreview As Boolean)
        Dim strReportName As String = String.Empty
        Try
            If Me.cboMonth.SelectedIndex = -1 Then
                MessageBox.Show("Please select Month!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Me.cboMonth.Focus()
                Me.cboMonth.SelectAll()
                Return
            End If

            Cursor = Cursors.WaitCursor

            myClsBG0473BL.BudgetYear = CStr(Me.numYear.Value)
            myClsBG0473BL.Month = Me.cboMonth.Text
            myClsBG0473BL.UserLevelId = p_intUserLevelId

            If myClsBG0473BL.getBudgetCompareData() Then

                Dim ds As DataSet = myClsBG0473BL.BudgetCompareData

                If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then

                    myClsBG0473BL.GetBudgetStatus()

                    strReportName = "RPT007-3.rpt"

                    If blnShowPrintPreview Then
                        'If clsBG0400 Is Nothing OrElse clsBG0400.IsDisposed Then
                        '    clsBG0400 = New frmBG0400()
                        'End If
                        If clsBG0400 IsNot Nothing Then
                            clsBG0400.Close()
                            clsBG0400.Dispose()
                        End If
                        clsBG0400 = New frmBG0400()
                        clsBG0400.MdiParent = p_frmBG0010
                        clsBG0400.ReportName = strReportName
                        clsBG0400.BudgetYear = Me.numYear.Value.ToString()
                        'clsBG0400.ParamPersonInCharge = False
                        clsBG0400.ReportType = "BudgetCompare"
                        clsBG0400.BudgetStatus = myClsBG0473BL.BudgetStatus
                        clsBG0400.Month = Me.cboMonth.Text

                        clsBG0400.DS = ds

                        clsBG0400.Show()
                        If clsBG0400.WindowState = FormWindowState.Minimized Then
                            clsBG0400.WindowState = FormWindowState.Normal
                        End If
                        clsBG0400.BringToFront()
                    Else
                        'Dim document As New Printing.PrintDocument ' This is the document to print
                        'Dim pageDialog2 As New PageSetupDialog ' This Dialog can set the paper size or kind
                        'Dim printDialog1 As New PrintDialog ' This is the dialog to setting the printer options
                        'Dim psize As Printing.PaperSize = Nothing

                        '' The parameter of Item method is any kind of paper size avaliable on the printer

                        'For i = 0 To document.PrinterSettings.PaperSizes.Count - 1

                        '    If document.PrinterSettings.PaperSizes.Item(i).Kind = PaperKind.A4 Then

                        '        psize = document.PrinterSettings.PaperSizes.Item(i)

                        '        Exit For

                        '    End If

                        'Next

                        'If psize Is Nothing Then
                        '    psize = document.PrinterSettings.PaperSizes.Item(0)
                        'End If

                        '' psize = document.PrinterSettings.PaperSizes.Item(6)

                        '' This line set the Page size of the document
                        'document.DefaultPageSettings.PaperSize = psize
                        'document.DefaultPageSettings.Landscape = True

                        '' This is for setting the page size on the page dialog
                        'pageDialog2.Document = document
                        'pageDialog2.PageSettings.PaperSize = psize
                        'pageDialog2.PageSettings.Landscape = True
                        '' This is for setting the page size on the printDialog
                        'printDialog1.Document = document
                        PrintDialog1.AllowSomePages = True
                        ' Allow the user to choose the page range he or she would
                        ' like to print.
                        ' PrintDialog1.AllowSomePages = True

                        ' Show the help button.
                        PrintDialog1.ShowHelp = True

                        Dim result As DialogResult = PrintDialog1.ShowDialog()

                        ' If the result is OK then print the document.
                        If (result = DialogResult.OK) Then

                            Dim rpt1 As ReportDocument = Nothing

                            rpt1 = New ReportDocument()
                            Dim reportPath As String = p_strAppPath & "\Reports\" & strReportName
                            rpt1.Load(reportPath)

                            'myClsBG0473BL.GetBudgetStatus()

                            'If myClsBG0473BL.BudgetStatus >= 5 Then
                            '    rpt1.ReportDefinition.ReportObjects("picAuth10").ObjectFormat.EnableSuppress = False
                            'Else
                            '    rpt1.ReportDefinition.ReportObjects("picAuth10").ObjectFormat.EnableSuppress = True
                            'End If

                            'If myClsBG0473BL.BudgetStatus >= 6 Then
                            '    rpt1.ReportDefinition.ReportObjects("picAuth3").ObjectFormat.EnableSuppress = False
                            'Else
                            '    rpt1.ReportDefinition.ReportObjects("picAuth3").ObjectFormat.EnableSuppress = True
                            'End If

                            rpt1.SetDataSource(ds)

                            rpt1.SetParameterValue("BUDGET_YEAR", Me.numYear.Value.ToString())
                            rpt1.SetParameterValue("HALF_BUDGET_YEAR", Me.numYear.Value.ToString().Substring(2, 2))
                            rpt1.SetParameterValue("MONTH", Me.cboMonth.Text)

                            rpt1.PrintOptions.PrinterName = PrintDialog1.PrinterSettings.PrinterName
                            rpt1.PrintToPrinter(PrintDialog1.PrinterSettings.Copies, _
                                                PrintDialog1.PrinterSettings.Collate, _
                                                PrintDialog1.PrinterSettings.FromPage, _
                                                PrintDialog1.PrinterSettings.ToPage)

                        End If
                    End If
                Else
                    MessageBox.Show("No data is available for viewing reports!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

                'Else
                '    MessageBox.Show("There are errors during the retrieved view reports!", "Detail by Account Code Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Cursor = Cursors.Default
    End Sub

    Private Function InsertExcelColumnData(ByRef dtColumns As DataTable, ByVal strYear As String, ByVal strMonth As String) As Boolean

        Dim dRow As DataRow
        Dim strDiffPeriod As String = BGCommon.GetBudgetCompareDiffPeriod(strMonth)

        dRow = dtColumns.NewRow()
        dRow("Column_Name") = "ACCOUNT_NO"
        dRow("Column_Title") = "Item"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow()
        dRow("Column_Name") = "ACCOUNT_NAME"
        dRow("Column_Title") = ""
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "OB_M" & strMonth
        dRow("Column_Title") = "Budget"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "AC_M" & strMonth
        dRow("Column_Title") = "Actual"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_MONTH"
        dRow("Column_Title") = strDiffPeriod
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERCENT_MONTH"
        dRow("Column_Title") = "%"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "ACC_OB_M" & strMonth
        dRow("Column_Title") = "Budget"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "ACC_AC_M" & strMonth
        dRow("Column_Title") = "Actual"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_HALF"
        dRow("Column_Title") = strDiffPeriod
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERCENT_HALF"
        dRow("Column_Title") = "%"
        dtColumns.Rows.Add(dRow)

        Return True

    End Function

    Private Function CalculateInvestments(ByVal strMonth As String, _
                                                  ByVal dsData As DataSet, _
                                                  ByVal intDataColumnIndex As Integer, _
                                                  ByRef drInvestments As DataRow) As Boolean
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object

        For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName

            If strColumnName.Equals(AC_M_COL & strMonth) OrElse _
                   strColumnName.Equals(OB_M_COL & strMonth) OrElse _
                   strColumnName.Equals(AC_H_COL & strMonth) OrElse _
                   strColumnName.Equals(OB_H_COL & strMonth) Then

                strExpression = "Sum(" + strColumnName + ")"
                strFilter = "BUDGET_TYPE = 'A'"
                returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
                drInvestments(strColumnName) = returnValue

            End If

        Next

        '// Diff Month
        drInvestments![DIFF_MONTH] = Convert.ToDecimal(Nz(drInvestments(AC_M_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drInvestments(OB_M_COL & strMonth), 0.0))
        drInvestments![PERCENT_MONTH] = CalPercent(Convert.ToDecimal(Nz(drInvestments(AC_M_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drInvestments(OB_M_COL & strMonth), 0.0)))

        '// Diff Half
        drInvestments![DIFF_HALF] = Convert.ToDecimal(Nz(drInvestments(AC_H_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drInvestments(OB_H_COL & strMonth), 0.0))
        drInvestments![PERCENT_HALF] = CalPercent(Convert.ToDecimal(Nz(drInvestments(AC_H_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drInvestments(OB_H_COL & strMonth), 0.0)))


        Return True
    End Function

    Private Function CalculateWorkingBudget(ByVal strMonth As String, _
                                                  ByVal dsData As DataSet, _
                                                  ByVal intDataColumnIndex As Integer, _
                                                  ByRef drWorkingBudget As DataRow) As Boolean
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object

        'For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

        '    Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName

        '    If strColumnName.Equals(OB_M_COL & strMonth) Then

        '        strExpression = "Sum(" + OB_M_COL + strMonth + ")"
        '        strFilter = "BUDGET_TYPE = 'W'"
        '        returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
        '        drWorkingBudget(strColumnName) = returnValue

        '    ElseIf strColumnName.Equals(OB_H_COL & strMonth) Then

        '        strExpression = "Sum(" + OB_H_COL + strMonth + ")"
        '        strFilter = "BUDGET_TYPE = 'W'"
        '        returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
        '        drWorkingBudget(strColumnName) = returnValue

        '    End If

        'Next
        For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName

            If strColumnName.Equals(AC_M_COL & strMonth) OrElse _
                   strColumnName.Equals(OB_M_COL & strMonth) OrElse _
                   strColumnName.Equals(AC_H_COL & strMonth) OrElse _
                   strColumnName.Equals(OB_H_COL & strMonth) Then

                strExpression = "Sum(" + strColumnName + ")"
                strFilter = "BUDGET_TYPE = 'W'"
                returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
                drWorkingBudget(strColumnName) = returnValue

            End If

        Next

        '// Diff Month
        drWorkingBudget![DIFF_MONTH] = Convert.ToDecimal(Nz(drWorkingBudget(AC_M_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drWorkingBudget(OB_M_COL & strMonth), 0.0))
        drWorkingBudget![PERCENT_MONTH] = CalPercent(Convert.ToDecimal(Nz(drWorkingBudget(AC_M_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drWorkingBudget(OB_M_COL & strMonth), 0.0)))

        '// Diff Half
        drWorkingBudget![DIFF_HALF] = Convert.ToDecimal(Nz(drWorkingBudget(AC_H_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drWorkingBudget(OB_H_COL & strMonth), 0.0))
        drWorkingBudget![PERCENT_HALF] = CalPercent(Convert.ToDecimal(Nz(drWorkingBudget(AC_H_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drWorkingBudget(OB_H_COL & strMonth), 0.0)))


        Return True
    End Function

    Private Function SetupCompareGroupbyData(ByVal strMonth As String, _
                                             ByVal dsData As DataSet, _
                                             ByVal strGroupColumnName As String, _
                                             ByVal strGroupColumnTitle As String, _
                                             ByVal intDataColumnIndex As Integer) As DataSet

        Dim dsResult As DataSet = New DataSet

        Dim drEmpty As DataRow
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object

        Dim strScript As String = strGroupColumnName

        Dim strExpr As String = "BUDGET_TYPE = 'E'"
        Dim strSort As String = strGroupColumnName + " ASC"

        '//Sort groups list by group column name
        Dim dtTmp As DataTable = dsData.Tables(0).Clone
        Dim arrTmp As DataRow() = dsData.Tables(0).Select(strExpr, strSort)
        For intTmp As Integer = 0 To arrTmp.Length - 1
            Dim drow(dtTmp.Columns.Count - 1) As Object
            arrTmp(intTmp).ItemArray.CopyTo(drow, 0)
            dtTmp.Rows.Add(drow)
        Next

        '//Get groups list
        Dim dtGroups As DataTable = dtTmp.DefaultView.ToTable(True, strScript)
        Dim intGroupCount As Integer = dtGroups.Rows.Count

        Dim dtResult As DataTable = dsData.Tables(0).Clone

        Dim drInvestments As DataRow = dtResult.NewRow
        Dim drTotalExpense As DataRow = dtResult.NewRow
        Dim drWorkingBudget As DataRow = dtResult.NewRow
        Dim drOutflowTotal As DataRow = dtResult.NewRow


        '//Calculate Investments
        CalculateInvestments(strMonth, dsData, intDataColumnIndex, drInvestments)

        '//Calculate Working Budget
        CalculateWorkingBudget(strMonth, dsData, intDataColumnIndex, drWorkingBudget)

        '//Calculate Total Expense
        CalculateTotalExpense(strMonth, dsData, intDataColumnIndex, drTotalExpense)


        Dim intGroupTotalIndex As Integer = 0
        For i As Integer = 0 To intGroupCount - 1

            strScript = strGroupColumnName + " = " + dtGroups.Rows(i)(0).ToString

            Dim arrRows As DataRow() = dtTmp.Select(strScript)

            For j As Integer = 0 To arrRows.Length - 1
                Dim drow(dtResult.Columns.Count - 1) As Object
                arrRows(j).ItemArray.CopyTo(drow, 0)
                dtResult.Rows.Add(drow)
            Next

            '//Calculate total for each group
            Dim drTotal As DataRow = dtResult.NewRow
            For k As Integer = intDataColumnIndex To dtResult.Columns.Count - 1

                Dim strColumnName As String = dtResult.Columns(k).ColumnName

                If strColumnName.Equals(AC_M_COL & strMonth) OrElse _
                 strColumnName.Equals(OB_M_COL & strMonth) OrElse _
                 strColumnName.Equals(AC_H_COL & strMonth) OrElse _
                 strColumnName.Equals(OB_H_COL & strMonth) Then

                    strExpression = "Sum(" + strColumnName + ")"
                    strFilter = strScript
                    returnValue = dtResult.Compute(strExpression, strFilter)
                    drTotal(dtResult.Columns(k).ColumnName) = returnValue

                End If

            Next
            '// Diff Month
            drTotal![DIFF_MONTH] = Convert.ToDecimal(Nz(drTotal(AC_M_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drTotal(OB_M_COL & strMonth), 0.0))
            drTotal![PERCENT_MONTH] = CalPercent(Convert.ToDecimal(Nz(drTotal(AC_M_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drTotal(OB_M_COL & strMonth), 0.0)))

            '// Diff Half
            drTotal![DIFF_HALF] = Convert.ToDecimal(Nz(drTotal(AC_H_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drTotal(OB_H_COL & strMonth), 0.0))
            drTotal![PERCENT_HALF] = CalPercent(Convert.ToDecimal(Nz(drTotal(AC_H_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drTotal(OB_H_COL & strMonth), 0.0)))

            '//Set Group header
            drTotal("ACCOUNT_NO") = GetGroupExpensesTitle(dtGroups.Rows(i)(0).ToString)

            '//Add total cost
            dtResult.Rows.InsertAt(drTotal, intGroupTotalIndex)

            '//Add one empty row
            drEmpty = dtResult.NewRow
            dtResult.Rows.Add(drEmpty)

            intGroupTotalIndex = intGroupTotalIndex + CInt(arrRows.Length) + 2

        Next
        '//Set Data to Account No.
        SetAccountNoText(drInvestments, drTotalExpense, drWorkingBudget, drOutflowTotal)

        '//Calculate Outflow Total 
        CalculateOutflowTotal(strMonth, dtResult, intDataColumnIndex, drTotalExpense, drWorkingBudget, drInvestments, drOutflowTotal)

        '//Add Investments
        dtResult.Rows.InsertAt(drInvestments, 0)

        '//Add one empty row
        drEmpty = dtResult.NewRow
        dtResult.Rows.InsertAt(drEmpty, 1)

        '//Add Total Expense
        dtResult.Rows.Add(drTotalExpense)

        '//Add Working Budget
        dtResult.Rows.Add(drWorkingBudget)

        '//Add one empty row
        drEmpty = dtResult.NewRow
        dtResult.Rows.Add(drEmpty)

        '//Add Outflow Total (Investment;Expenses)
        dtResult.Rows.Add(drOutflowTotal)

        '//Return data table
        dsResult.Tables.Add(dtResult)

        dtResult.TableName = "Budget Compare"

        Return dsResult
    End Function

    Private Function CalculateTotalExpense(ByVal strMonth As String, _
                                                  ByVal dsData As DataSet, _
                                                  ByVal intDataColumnIndex As Integer, _
                                                  ByRef drTotalExpense As DataRow) As Boolean
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object

        For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName

            If strColumnName.Equals(AC_M_COL & strMonth) OrElse _
                   strColumnName.Equals(OB_M_COL & strMonth) OrElse _
                   strColumnName.Equals(AC_H_COL & strMonth) OrElse _
                   strColumnName.Equals(OB_H_COL & strMonth) Then

                strExpression = "Sum(" + strColumnName + ")"
                strFilter = "BUDGET_TYPE = 'E'"
                returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
                drTotalExpense(strColumnName) = returnValue

            End If

        Next

        '// Diff Month
        drTotalExpense![DIFF_MONTH] = Convert.ToDecimal(Nz(drTotalExpense(AC_M_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drTotalExpense(OB_M_COL & strMonth), 0.0))
        drTotalExpense![PERCENT_MONTH] = CalPercent(Convert.ToDecimal(Nz(drTotalExpense(AC_M_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drTotalExpense(OB_M_COL & strMonth), 0.0)))

        '// Diff Half
        drTotalExpense![DIFF_HALF] = Convert.ToDecimal(Nz(drTotalExpense(AC_H_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drTotalExpense(OB_H_COL & strMonth), 0.0))
        drTotalExpense![PERCENT_HALF] = CalPercent(Convert.ToDecimal(Nz(drTotalExpense(AC_H_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drTotalExpense(OB_H_COL & strMonth), 0.0)))


        Return True
    End Function

    Private Function CalculateOutflowTotal(ByVal strMonth As String, _
                                           ByVal dt As DataTable, _
                                           ByVal intDataColumnIndex As Integer, _
                                           ByVal drTotalExpense As DataRow, _
                                           ByVal drWorkingBudget As DataRow, _
                                           ByVal drInvestments As DataRow, _
                                           ByRef drOutflowTotal As DataRow) As Boolean

        For k As Integer = intDataColumnIndex To dt.Columns.Count - 1
            Dim strColumnName As String = dt.Columns(k).ColumnName

            If strColumnName.Equals(AC_M_COL & strMonth) OrElse _
                   strColumnName.Equals(OB_M_COL & strMonth) OrElse _
                   strColumnName.Equals(AC_H_COL & strMonth) OrElse _
                   strColumnName.Equals(OB_H_COL & strMonth) Then

                drOutflowTotal(strColumnName) = Convert.ToDecimal(Nz(drTotalExpense(strColumnName), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget(strColumnName), 0.0)) + Convert.ToDecimal(Nz(drInvestments(strColumnName), 0.0))

            End If

        Next

        '// Diff Month
        drOutflowTotal![DIFF_MONTH] = Convert.ToDecimal(Nz(drOutflowTotal(AC_M_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drOutflowTotal(OB_M_COL & strMonth), 0.0))
        drOutflowTotal![PERCENT_MONTH] = CalPercent(Convert.ToDecimal(Nz(drOutflowTotal(AC_M_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drOutflowTotal(OB_M_COL & strMonth), 0.0)))

        '// Diff Half
        drOutflowTotal![DIFF_HALF] = Convert.ToDecimal(Nz(drOutflowTotal(AC_H_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drOutflowTotal(OB_H_COL & strMonth), 0.0))
        drOutflowTotal![PERCENT_HALF] = CalPercent(Convert.ToDecimal(Nz(drOutflowTotal(AC_H_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drOutflowTotal(OB_H_COL & strMonth), 0.0)))


        Return True
    End Function

    Private Function SetAccountNoText(ByRef drInvestments As DataRow, _
                                      ByRef drTotalExpense As DataRow, _
                                      ByRef drWorkingBudget As DataRow, _
                                      ByRef drOutflowTotal As DataRow) As Boolean

        Dim strColumnName As String = "ACCOUNT_NO"

        drInvestments(strColumnName) = "Investments"
        drTotalExpense(strColumnName) = "Total Expense"
        drWorkingBudget(strColumnName) = "Working Budget"
        drOutflowTotal(strColumnName) = "Outflow Total"
        Return True

    End Function

    Private Sub CalculateDiff(ByVal strMonth As String, ByRef dsData As DataSet)

        '//Add column 
        Dim col As DataColumn = New DataColumn()
        col.ColumnName = "DIFF_MONTH"
        col.DataType = Type.GetType("System.Decimal")
        col.DefaultValue = 0.0
        dsData.Tables(0).Columns.Add(col)

        col = New DataColumn()
        col.ColumnName = "PERCENT_MONTH"
        col.DataType = Type.GetType("System.Decimal")
        col.DefaultValue = 0.0
        dsData.Tables(0).Columns.Add(col)

        col = New DataColumn()
        col.ColumnName = "DIFF_HALF"
        col.DataType = Type.GetType("System.Decimal")
        col.DefaultValue = 0.0
        dsData.Tables(0).Columns.Add(col)

        col = New DataColumn()
        col.ColumnName = "PERCENT_HALF"
        col.DataType = Type.GetType("System.Decimal")
        col.DefaultValue = 0.0
        dsData.Tables(0).Columns.Add(col)

        For Each drData As DataRow In dsData.Tables(0).Rows
            '// diff month      : (AC-OB)
            drData![DIFF_MONTH] = Convert.ToDecimal(Nz(drData(AC_M_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drData(OB_M_COL & strMonth), 0.0))

            '// percent month   : (AC- OB)/OB
            drData![PERCENT_MONTH] = CalPercent(Convert.ToDecimal(Nz(drData(AC_M_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drData(OB_M_COL & strMonth), 0.0)))

            '// diff half       : (Accumulate AC- Accumulate OB)
            drData![DIFF_HALF] = Convert.ToDecimal(Nz(drData(AC_H_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drData(OB_H_COL & strMonth), 0.0))

            '// percent half    : (Accumulate AC- Accumulate OB)/Accumulate OB
            drData![PERCENT_HALF] = CalPercent(Convert.ToDecimal(Nz(drData(AC_H_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drData(OB_H_COL & strMonth), 0.0)))

        Next

    End Sub

    Private Function GeneratOriginalExcel(ByVal dsData As DataSet, ByVal dtColumns As DataTable, ByVal strMonth As String, ByVal strYear As String) As Boolean
        Dim blnRet As Boolean = False
        Dim rowStartIndex As Integer = 8
        Dim colStartIndex As Integer = 7
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim rng As Excel.Range = Nothing

        'excelApp = New Excel.Application()
        If excelApp Is Nothing Then
            excelApp = New Excel.Application
        End If

        'excelApp.Visible = False
        'excelApp.UserControl = False

        wb = excelApp.Workbooks.Add(missing)

        '//Delete Worksheets
        If wb.Worksheets.Count > 1 Then
            For i As Integer = 1 To wb.Worksheets.Count - 1
                CType(wb.Worksheets(i), Excel.Worksheet).Delete()
            Next
        End If

        For intSheetCount As Integer = 0 To dsData.Tables.Count - 1

            If intSheetCount <> 0 Then
                wb.Sheets.Add()
            End If

            ws = CType(wb.ActiveSheet, Excel.Worksheet)
            Dim strSheetName As String = dsData.Tables(intSheetCount).TableName
            ws.Name = strSheetName

            '//Setup DataColumn
            For i As Integer = 0 To dtColumns.Rows.Count - 1
                ws.Cells(colStartIndex, i + 1) = dtColumns.Rows(i)("Column_Title").ToString
                rng = ws.Range(ws.Cells(colStartIndex, i + 1), ws.Cells(colStartIndex, i + 1))
                rng.Font.Bold = True
                rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            Next

            Dim arrCols() As Integer

            arrCols = New Integer() {1}    '// Two Row Merge Col
            SetupCompareColumnsCells(ws, colStartIndex, CInt(strMonth), 1, 2, "Item", _
                                     arrCols, 3, 6, 7, 10, strYear)

            '//Setup Data
            For rowIndex As Integer = 0 To dsData.Tables(intSheetCount).Rows.Count - 1

                Dim row As DataRow = dsData.Tables(intSheetCount).Rows(rowIndex)

                If IsAccountNoEmpty(row) Then
                    Continue For
                End If

                For colIndex As Integer = 0 To dtColumns.Rows.Count - 1

                    Dim strColumnName As String = dtColumns.Rows(colIndex)("Column_Name").ToString
                    Dim col As DataColumn = dsData.Tables(intSheetCount).Columns(strColumnName)

                    '//Setup Font of Expense group is bold.
                    SetExpenseGroupBold(ws, strColumnName, row, col, rowIndex, rowStartIndex, colIndex, dtColumns.Rows.Count)


                    If col.DataType Is System.Type.GetType("System.DateTime") Then

                        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = (Convert.ToDateTime(row(col.ColumnName).ToString())).ToString("yyyy-MM-dd")
                        ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter

                    ElseIf col.DataType Is System.Type.GetType("System.String") Then

                        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = "'" + row(col.ColumnName).ToString()
                        ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter

                    ElseIf col.DataType Is System.Type.GetType("System.Decimal") Then

                        If row(col.ColumnName).ToString = String.Empty Then
                            row(col.ColumnName) = "0.00"
                        End If

                        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName)
                        ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).NumberFormat = "#,##0.00"

                    Else
                        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName).ToString()
                    End If

                Next
            Next

            Dim rowMax As Integer = dsData.Tables(intSheetCount).Rows.Count + colStartIndex
            Dim colMax As Integer = dtColumns.Rows.Count
            Dim intUnitPriceStart As Integer = 10

            '//Setup Investments Line
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).ClearContents()
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).MergeCells = True
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).Value = "Investments"
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, colMax)).Font.Bold = True

            '//Setup budget order name column to be left align
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowMax, 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            ws.Range(ws.Cells(rowStartIndex, 2), ws.Cells(rowMax, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup Total Expense Line
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).Value = "Total Expense"
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, colMax)).Font.Bold = True

            '//Setup Working Budget Line
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).Value = "Working Budget"
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Empry line
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).Value = ""
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup Outflow Total Line
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).Value = "Outflow Total (Investment;Expenses)"
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, colMax)).Font.Bold = True

            '// Set Borders
            rng = ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(rowMax, colMax))
            rng.Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.EntireColumn.AutoFit()

            '//Setup Wrap text for columns title
            ws.Range(ws.Cells(2, 3), ws.Cells(rowMax, 10)).Columns.ColumnWidth = 12

            '//Set Font
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(rowMax, colMax)).Font.Name = "Tahoma"
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(rowMax, colMax)).Font.Size = 10

            '//Setup Title & Title Font 
            SetupExcelTitle(ws, intUnitPriceStart, strMonth, strYear)

        Next

        '// Show excel
        excelApp.Visible = True
        '//Select the first worksheet in a workbook using the Excel Sheets collection
        CType(excelApp.Application.ActiveWorkbook.Sheets(1), Excel.Worksheet).Select()

        '//Release memory
        BGCommon.ExcelReleasememory(excelApp, wb, ws)

        blnRet = True

        Return blnRet
    End Function

    Private Sub SetExpenseGroupBold(ByVal ws As Excel.Worksheet, _
                                ByVal strColumnName As String, _
                                ByVal row As DataRow, _
                                ByVal col As DataColumn, _
                                ByVal rowIndex As Integer, _
                                ByVal rowStartIndex As Integer, _
                                ByVal colIndex As Integer, _
                                ByVal colMax As Integer)
        '//If a Column "ACCOUNT_NO" and Expense group.
        If strColumnName = "ACCOUNT_NO" Then
            If row(col).ToString() = P_EXPENSE_TYPE_LABOR _
            Or row(col).ToString() = P_EXPENSE_TYPE_VARIABLE _
            Or row(col).ToString() = P_EXPENSE_TYPE_FIXED Then
                ws.Range(excelApp.Cells(rowIndex + rowStartIndex, 1), excelApp.Cells(rowIndex + rowStartIndex, 2)).MergeCells = True
                ws.Range(excelApp.Cells(rowIndex + rowStartIndex, 1), excelApp.Cells(rowIndex + rowStartIndex, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colMax)).Font.Bold = True
            Else
                ws.Range(excelApp.Cells(rowIndex + rowStartIndex, 1), excelApp.Cells(rowIndex + rowStartIndex, 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colMax)).Font.Bold = False
            End If
        End If
    End Sub

    Private Function SetupExcelTitle(ByVal ws As Excel.Worksheet, ByVal intUnitPriceStart As Integer, ByVal strMonth As String, ByVal strYear As String) As Boolean
        Dim strSubTitle As String = "Summary by Account No : Budget Compare Year " + strYear

        ws.Range(ws.Cells(1, 1), ws.Cells(1, 4)).Font.Bold = True
        ws.Range(ws.Cells(1, 1), ws.Cells(1, 4)).Font.Size = 12
        ws.Range(ws.Cells(1, 1), ws.Cells(1, 4)).Font.Name = "Tahoma"
        ws.Range(ws.Cells(1, 1), ws.Cells(1, 4)).MergeCells = True
        ws.Range(ws.Cells(1, 1), ws.Cells(1, 4)).Value = "Bridgestone Tire Manufacturing (Thailand) Co.,Ltd."

        '//Setup subTitle  
        ws.Range(ws.Cells(2, 1), ws.Cells(2, 4)).Font.Bold = True
        ws.Range(ws.Cells(2, 1), ws.Cells(2, 4)).Font.Size = 11
        ws.Range(ws.Cells(2, 1), ws.Cells(2, 4)).Font.Name = "Tahoma"
        ws.Range(ws.Cells(2, 1), ws.Cells(2, 4)).MergeCells = True
        ws.Range(ws.Cells(2, 1), ws.Cells(2, 4)).Value = strSubTitle

        '//Setup unit price
        ws.Range(excelApp.Cells(4, intUnitPriceStart), excelApp.Cells(4, intUnitPriceStart)).Font.Name = "Tahoma"
        ws.Range(excelApp.Cells(4, intUnitPriceStart), excelApp.Cells(4, intUnitPriceStart)).Font.Bold = True
        ws.Range(excelApp.Cells(4, intUnitPriceStart), excelApp.Cells(4, intUnitPriceStart)).Font.Underline = True
        ws.Range(excelApp.Cells(4, intUnitPriceStart), excelApp.Cells(4, intUnitPriceStart)).Font.Size = 11
        ws.Range(excelApp.Cells(4, intUnitPriceStart), excelApp.Cells(4, intUnitPriceStart)).MergeCells = True
        ws.Range(excelApp.Cells(4, intUnitPriceStart), excelApp.Cells(4, intUnitPriceStart)).Value = "Unit : K.Baht"
        ws.Range(ws.Cells(4, intUnitPriceStart), ws.Cells(4, intUnitPriceStart)).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

    End Function

#End Region

#Region "Controls Event"

    Private Sub frmBG0473_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        If Not excelApp Is Nothing Then
            excelApp.Quit()
            excelApp = Nothing
        End If
    End Sub

    Private Sub frmBG473_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        InitPage()

    End Sub

    Private Sub cmdPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPreview.Click
        Print(True)
    End Sub

    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrint.Click
        Print(False)
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub cmdExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExcel.Click
        Try
            If Me.cboMonth.SelectedIndex = -1 Then
                MessageBox.Show("Please select Month!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Me.cboMonth.Focus()
                Me.cboMonth.SelectAll()
                Return
            End If

            Cursor = Cursors.WaitCursor

            Dim dsData As DataSet
            myClsBG0473BL.BudgetYear = CStr(Me.numYear.Value)
            myClsBG0473BL.Month = Me.cboMonth.Text
            myClsBG0473BL.UserLevelId = p_intUserLevelId

            If myClsBG0473BL.GetBudgetCompareData() Then

                dsData = myClsBG0473BL.BudgetCompareData

                If dsData IsNot Nothing AndAlso dsData.Tables(0).Rows.Count > 0 Then

                    myClsBG0473BL.GetBudgetStatus()

                    Dim strYear As String = Me.numYear.Value.ToString
                    Dim strMonth As String = Me.cboMonth.Text
                    '//Create output columns
                    Dim dtColumns As DataTable = CreateTableTemplate()
                    Dim dsGroups As DataSet = Nothing

                    InsertExcelColumnData(dtColumns, strYear, strMonth)
                    CalculateDiff(strMonth, dsData)

                    '//Dataset Groupby 
                    dsGroups = SetupCompareGroupbyData(strMonth, dsData, "EXPENSE_TYPE", "EXPENSE_TYPE", 4)
                    '//Generate excel
                    GeneratOriginalExcel(dsGroups, dtColumns, strMonth, strYear)

                Else
                    MessageBox.Show("No budget data found, please try it again.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Cursor = Cursors.Default
    End Sub

#End Region

End Class