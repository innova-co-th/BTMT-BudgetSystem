Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports CrystalDecisions.CrystalReports.Engine
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Drawing.Printing

Public Class frmBG0450

#Region "Variable"
    Private myClsBG0450BL As New clsBG0450BL
    Private myClsBG0310BL As New clsBG0310BL
    Private clsBG0400 As frmBG0400
    Private m_blnFormLoading As Boolean = False
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
    Private Sub LoadBudgetYear()
        Try
            Me.numYear.Value = CInt(Now.ToString("yyyy"))
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Summary By Applicant Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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

    'Private Sub LoadPeriodType()
    '    Try
    '        Me.cboPeriodType.Items.Clear()

    '        Dim dt As DataTable = New DataTable("PeriodType")
    '        Dim dc As DataColumn = dt.Columns.Add( _
    '            "PeriodTypeID", Type.GetType("System.Int32"))
    '        dc.AllowDBNull = False
    '        dc.Unique = True

    '        dt.Columns.Add("PeriodTypeName", Type.GetType("System.String"))

    '        Dim dr As DataRow = dt.NewRow()
    '        dt.Rows.Add(New Object() {enumPeriodType.OriginalBudget, "Original Budget"})
    '        dt.Rows.Add(New Object() {enumPeriodType.EstimateBudget, "Estimate Budget"})
    '        dt.Rows.Add(New Object() {enumPeriodType.ReviseBudget, "Revise Budget"})

    '        Me.cboPeriodType.DataSource = dt
    '        Me.cboPeriodType.DisplayMember = "PeriodTypeName"
    '        Me.cboPeriodType.ValueMember = "PeriodTypeID"

    '        dc = Nothing
    '        dr = Nothing
    '        dt = Nothing

    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message, "Summary By Applicant Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub

    Private Sub Print(ByVal blnShowPrintPreview As Boolean)
        Dim strReportName As String = String.Empty
        Try
            If Me.cboPeriodType.SelectedIndex = -1 Then
                MessageBox.Show("Please select a Period Type!", "Summary By Applicant Report", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Me.cboPeriodType.Focus()
                Me.cboPeriodType.SelectAll()
                Return
            End If

            If fncCheckRevNo() = False Then

                MessageBox.Show("No budget data found, please try it again.", "Summary By Applicant Report", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            'If fncCheckPrevRevNo() = False Then
            '    MessageBox.Show("No previous budget data found, please try it again.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            '    Exit Sub
            'End If

            Cursor = Cursors.WaitCursor

            myClsBG0450BL.BudgetYear = CStr(Me.numYear.Value)
            myClsBG0450BL.PeriodType = CStr(Me.cboPeriodType.SelectedValue)
            myClsBG0450BL.ProjectNo = Me.numProjectNo.Value.ToString
            myClsBG0450BL.UserLevelId = p_intUserLevelId
            If Me.cboRevNo.DataSource IsNot Nothing Then
                myClsBG0450BL.RevNo = Me.cboRevNo.SelectedValue.ToString
            End If

            myClsBG0450BL.PrevProjectNo = Me.numPrevProjectNo.Value.ToString
            If Me.cboPrevRevno.DataSource IsNot Nothing AndAlso _
                Me.cboPrevRevno.SelectedValue IsNot Nothing Then
                myClsBG0450BL.PrevRevNo = Me.cboPrevRevno.SelectedValue.ToString
            End If

            'myClsBG0450BL.MTPBudget = Me.chkShowMTP.Checked

            If myClsBG0450BL.getApplicantData() Then

                Dim ds As DataSet = myClsBG0450BL.ApplicantData

                If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then

                    myClsBG0450BL.GetBudgetStatus()

                    myClsBG0450BL.GetAuthImage()
                    ds.Tables.Add(myClsBG0450BL.AuthImage)

                    Select Case CType(Me.cboPeriodType.SelectedValue, enumPeriodType)
                        Case enumPeriodType.OriginalBudget
                            strReportName = "RPT005-1.rpt"
                        Case enumPeriodType.EstimateBudget
                            strReportName = "RPT005-2.rpt"
                        Case enumPeriodType.ReviseBudget
                            If Not chkShowMTP.Checked Then
                                strReportName = "RPT005-3.rpt"
                            Else
                                strReportName = "RPT005-4.rpt"
                            End If
                        Case enumPeriodType.MTPBudget
                            strReportName = "RPT005-5.rpt"
                    End Select

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
                        clsBG0400.ReportType = "SummaryByApplicant"
                        clsBG0400.BudgetStatus = myClsBG0450BL.BudgetStatus
                        clsBG0400.ProjectNo = Me.numProjectNo.Value.ToString

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

                            myClsBG0450BL.GetBudgetStatus()

                            If myClsBG0450BL.BudgetStatus >= 5 Then
                                rpt1.ReportDefinition.ReportObjects("picAuth10").ObjectFormat.EnableSuppress = False
                            Else
                                rpt1.ReportDefinition.ReportObjects("picAuth10").ObjectFormat.EnableSuppress = True
                            End If

                            If myClsBG0450BL.BudgetStatus >= 6 Then
                                rpt1.ReportDefinition.ReportObjects("picAuth3").ObjectFormat.EnableSuppress = False
                            Else
                                rpt1.ReportDefinition.ReportObjects("picAuth3").ObjectFormat.EnableSuppress = True
                            End If

                            rpt1.SetDataSource(ds)

                            rpt1.SetParameterValue("BUDGET_YEAR", Me.numYear.Value.ToString())
                            rpt1.SetParameterValue("HALF_BUDGET_YEAR", Me.numYear.Value.ToString().Substring(2, 2))
                            rpt1.SetParameterValue("PROJECT_NO", Me.numProjectNo.Value.ToString())

                            rpt1.PrintOptions.PrinterName = PrintDialog1.PrinterSettings.PrinterName
                            rpt1.PrintToPrinter(PrintDialog1.PrinterSettings.Copies, _
                                                PrintDialog1.PrinterSettings.Collate, _
                                                PrintDialog1.PrinterSettings.FromPage, _
                                                PrintDialog1.PrinterSettings.ToPage)

                        End If
                    End If
                Else
                    MessageBox.Show("No data is available for viewing reports!", "Summary By Applicant Report", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

                'Else
                '    MessageBox.Show("There are errors during the retrieved view reports!", "Detail by Account Code Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Summary By Applicant Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Cursor = Cursors.Default
    End Sub

    Private Function InsertOriginalColumnData(ByRef dtColumns As DataTable, _
                                             ByVal strYear As String) As Boolean
        Dim strHalfYear As String = strYear.Substring(2, 2)
        Dim row As DataRow

        Dim intYear As Integer = CInt(strYear)
        Dim strLastYear As String = CStr(intYear - 1)

        Dim strHalfLastYear As String = CStr(intYear - 1).Substring(2, 2)

        'MAX_REV.BUDGET_YEAR, MAX_REV.PERIOD_TYPE, '' AS BUDGET_ORDER_NO, 
        'MAX_REV.ACCOUNT_NO,
        row = dtColumns.NewRow()
        row("Column_Name") = "ACCOUNT_NO"
        row("Column_Title") = "Item"
        dtColumns.Rows.Add(row)

        ' MAX_REV.ACCOUNT_NAME,
        row = dtColumns.NewRow()
        row("Column_Name") = "ACCOUNT_NAME"
        row("Column_Title") = ""
        dtColumns.Rows.Add(row)

        ' MAX_REV.COST, MAX_REV.EXPENSE_TYPE,
        'MAX(MAX_REV.REV_NO) AS REV_NO,
        'SUM(ISNULL(ACTUAL_DATA.H1,0)) AS ACTUAL_1ST_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "ACTUAL_1ST_HALF"
        row("Column_Title") = "Actual 1st Half'" & strHalfLastYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(REVISE_BUDGET.H2,0)) AS REVISE_2ND_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_2ND_HALF"
        row("Column_Title") = "Estimate 2nd Half'" & strHalfLastYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M1, 0)) AS M1,
        row = dtColumns.NewRow()
        row("Column_Name") = "M1"
        row("Column_Title") = "Jan'" & strHalfYear
        dtColumns.Rows.Add(row)

        ' SUM(ISNULL(MASTER_DATA.M2, 0)) AS M2,
        row = dtColumns.NewRow()
        row("Column_Name") = "M2"
        row("Column_Title") = "Feb'" & strHalfYear
        dtColumns.Rows.Add(row)

        ' SUM(ISNULL(MASTER_DATA.M3, 0)) AS M3, 
        row = dtColumns.NewRow()
        row("Column_Name") = "M3"
        row("Column_Title") = "Mar'" & strHalfYear
        dtColumns.Rows.Add(row)


        'SUM(ISNULL(MASTER_DATA.M4, 0)) AS M4,
        row = dtColumns.NewRow()
        row("Column_Name") = "M4"
        row("Column_Title") = "Apr'" & strHalfYear
        dtColumns.Rows.Add(row)

        ' SUM(ISNULL(MASTER_DATA.M5, 0)) AS M5, 
        row = dtColumns.NewRow()
        row("Column_Name") = "M5"
        row("Column_Title") = "May'" & strHalfYear
        dtColumns.Rows.Add(row)


        'SUM(ISNULL(MASTER_DATA.M6, 0)) AS M6,
        row = dtColumns.NewRow()
        row("Column_Name") = "M6"
        row("Column_Title") = "Jun'" & strHalfYear
        dtColumns.Rows.Add(row)


        'SUM(ISNULL(MASTER_DATA.M7, 0)) AS M7, SUM(ISNULL(MASTER_DATA.M8, 0)) AS M8, SUM(ISNULL(MASTER_DATA.M9, 0)) AS M9, SUM(ISNULL(MASTER_DATA.M10, 0)) AS M10, SUM(ISNULL(MASTER_DATA.M11, 0)) AS M11, SUM(ISNULL(MASTER_DATA.M12, 0)) AS M12,
        'SUM(ISNULL(MASTER_DATA.M1, 0) + ISNULL(MASTER_DATA.M2, 0) + ISNULL(MASTER_DATA.M3, 0) + ISNULL(MASTER_DATA.M4, 0) + ISNULL(MASTER_DATA.M5, 0) + ISNULL(MASTER_DATA.M6, 0)) AS TOTAL_1ST_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "TOTAL_1ST_HALF"
        row("Column_Title") = "Total 1st Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M1, 0)) AS M1,
        row = dtColumns.NewRow()
        row("Column_Name") = "M7"
        row("Column_Title") = "Jul'" & strHalfYear
        dtColumns.Rows.Add(row)

        ' SUM(ISNULL(MASTER_DATA.M2, 0)) AS M2,
        row = dtColumns.NewRow()
        row("Column_Name") = "M8"
        row("Column_Title") = "Aug'" & strHalfYear
        dtColumns.Rows.Add(row)

        ' SUM(ISNULL(MASTER_DATA.M3, 0)) AS M3, 
        row = dtColumns.NewRow()
        row("Column_Name") = "M9"
        row("Column_Title") = "Sept'" & strHalfYear
        dtColumns.Rows.Add(row)


        'SUM(ISNULL(MASTER_DATA.M4, 0)) AS M4,
        row = dtColumns.NewRow()
        row("Column_Name") = "M10"
        row("Column_Title") = "Oct'" & strHalfYear
        dtColumns.Rows.Add(row)

        ' SUM(ISNULL(MASTER_DATA.M5, 0)) AS M5, 
        row = dtColumns.NewRow()
        row("Column_Name") = "M11"
        row("Column_Title") = "Nov'" & strHalfYear
        dtColumns.Rows.Add(row)


        'SUM(ISNULL(MASTER_DATA.M6, 0)) AS M6,
        row = dtColumns.NewRow()
        row("Column_Name") = "M12"
        row("Column_Title") = "Dec'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M7, 0) + ISNULL(MASTER_DATA.M8, 0) + ISNULL(MASTER_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11, 0) + ISNULL(MASTER_DATA.M12, 0)) AS TOTAL_2ND_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "TOTAL_2ND_HALF"
        row("Column_Title") = "Total 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M1, 0) + ISNULL(MASTER_DATA.M2, 0) + ISNULL(MASTER_DATA.M3, 0) + ISNULL(MASTER_DATA.M4, 0) + ISNULL(MASTER_DATA.M5, 0) + ISNULL(MASTER_DATA.M6, 0) + ISNULL(MASTER_DATA.M7, 0) + ISNULL(MASTER_DATA.M8, 0) + ISNULL(MASTER_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11, 0) + ISNULL(MASTER_DATA.M12, 0)) AS TOTAL_YEAR
        row = dtColumns.NewRow()
        row("Column_Name") = "TOTAL_YEAR"
        row("Column_Title") = "Total Year'" & strYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "MTP_RRT1"
        row("Column_Title") = "MTP " & CInt(strYear) - 1 & " Year'" & strYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "DIFF_MTP"
        row("Column_Title") = "Diff vs MTP" & CInt(strYear) - 1
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "TOTAL_LAST_YEAR"
        row("Column_Title") = "Total Year'" & CInt(strYear) - 1
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "DIFFERENCE"
        row("Column_Title") = "Diff vs Year'" & CInt(strYear) - 1
        dtColumns.Rows.Add(row)

 


        Return True
    End Function

    Private Function SetupOriginalGroupbyData(ByVal dsData As DataSet, _
                                           ByVal strGroupColumnName As String, _
                                           ByVal strGroupColumnTitle As String, _
                                           ByVal intDataColumnIndex As Integer) As DataSet

        Dim dsResult As DataSet = New DataSet

        Dim drEmpty As DataRow
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object

        Dim strScript As String = strGroupColumnName

        Dim dtGroups As DataTable = dsData.Tables(0).DefaultView.ToTable(True, strScript)
        Dim intGroupCount As Integer = dtGroups.Rows.Count

        For i As Integer = 0 To intGroupCount - 1

            Dim dtResult As DataTable = dsData.Tables(0).Clone

            '//Add column 
            Dim col As DataColumn = New DataColumn()
            col.ColumnName = "TOTAL_LAST_YEAR"
            col.DataType = Type.GetType("System.Decimal")
            col.DefaultValue = 0.0
            dtResult.Columns.Add(col)

            col = New DataColumn()
            col.ColumnName = "DIFFERENCE"
            col.DataType = Type.GetType("System.Decimal")
            col.DefaultValue = 0.0
            dtResult.Columns.Add(col)

            col = New DataColumn()
            col.ColumnName = "DIFF_MTP"
            col.DataType = Type.GetType("System.Decimal")
            col.DefaultValue = 0.0
            dtResult.Columns.Add(col)

            Dim drCost As DataRow = dtResult.NewRow
            Dim drTotalCost As DataRow = dtResult.NewRow
            Dim intGroupTotalIndex As Integer = 0

            strScript = strGroupColumnName + " = '" + dtGroups.Rows(i)(0).ToString & "'"
            Dim arrRows As DataRow() = dsData.Tables(0).Select(strScript)

            Dim strGroupColumnName2 As String = "EXPENSE_TYPE"
            '//Seperate dataset data into several datatables according to group no
            strScript = strGroupColumnName2
            Dim dtGroups2 As DataTable = dsData.Tables(0).DefaultView.ToTable(True, strScript)
            Dim intGroupCount2 As Integer = dtGroups2.Rows.Count

            For n As Integer = 0 To intGroupCount2 - 1

                Dim drTotalExpenses As DataRow = dtResult.NewRow

                strScript = strGroupColumnName + " = " + dtGroups.Rows(i)(0).ToString
                strScript &= " AND "
                strScript &= strGroupColumnName2 + " = " + dtGroups2.Rows(n)(0).ToString
                Dim arrRows2 As DataRow() = dsData.Tables(0).Select(strScript)

                If arrRows2.Length > 0 Then
                    For j As Integer = 0 To arrRows2.Length - 1
                        Dim drow(dtResult.Columns.Count - 1) As Object
                        arrRows2(j).ItemArray.CopyTo(drow, 0)
                        dtResult.Rows.Add(drow)
                    Next

                    '//Calculate Horizontal Column
                    For m As Integer = 0 To dtResult.Rows.Count - 1
                        For p As Integer = intDataColumnIndex To dtResult.Columns.Count - 1
                            Dim strColumnName As String = dtResult.Columns(p).ColumnName
                            If strColumnName = "TOTAL_LAST_YEAR" Then
                                '{DetailByAccountCode.ACTUAL_1ST_HALF} + {DetailByAccountCode.REVISE_2ND_HALF}
                                dtResult.Rows(m)![TOTAL_LAST_YEAR] = Convert.ToDecimal(Nz(dtResult.Rows(m)![ACTUAL_1ST_HALF], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![REVISE_2ND_HALF], 0.0))
                            ElseIf strColumnName = "DIFFERENCE" Then
                                '{@TotalYear} - {@TotalLastYear}
                                dtResult.Rows(m)![DIFFERENCE] = Convert.ToDecimal(Nz(dtResult.Rows(m)![TOTAL_YEAR], 0.0)) - Convert.ToDecimal(Nz(dtResult.Rows(m)![TOTAL_LAST_YEAR], 0.0))
                            ElseIf strColumnName = "DIFF_MTP" Then
                                '{@TotalYear} - {@TotalLastYear}
                                dtResult.Rows(m)![DIFF_MTP] = Convert.ToDecimal(Nz(dtResult.Rows(m)![TOTAL_YEAR], 0.0)) - Convert.ToDecimal(Nz(dtResult.Rows(m)![MTP_RRT1], 0.0))
                            End If

                        Next
                    Next
                    dtResult.AcceptChanges()

                    '//Calculate total for each group
                    For k As Integer = intDataColumnIndex To dtResult.Columns.Count - 1

                        Dim strColumnName As String = dtResult.Columns(k).ColumnName
                        strExpression = "Sum(" + strColumnName + ")"
                        strFilter = strGroupColumnName + " = " + dtGroups.Rows(i)(0).ToString
                        strFilter &= " AND "
                        strFilter &= strGroupColumnName2 + " = " + dtGroups2.Rows(n)(0).ToString
                        returnValue = dtResult.Compute(strExpression, strFilter)
                        drTotalExpenses(dtResult.Columns(k).ColumnName) = returnValue

                    Next
                    '//Setup Group header
                    drTotalExpenses("ACCOUNT_NO") = BGCommon.GetGroupExpensesTitle(arrRows2(0)(strGroupColumnName2).ToString)

                    '//Add total cost
                    dtResult.Rows.InsertAt(drTotalExpenses, intGroupTotalIndex)

                    '//Add one empty row
                    drEmpty = dtResult.NewRow
                    dtResult.Rows.Add(drEmpty)

                    'intGroupTotalIndex = intGroupTotalIndex + dtResult.Rows.Count
                    intGroupTotalIndex = intGroupTotalIndex + CInt(arrRows2.Length) + 2
                End If



            Next

            '//Calculate Total cost
            For l As Integer = intDataColumnIndex To dtResult.Columns.Count - 1
                Dim strColumnName As String = dtResult.Columns(l).ColumnName

                strExpression = "Sum(" + strColumnName + ")"
                strFilter = strGroupColumnName + " = " + dtGroups.Rows(i)(0).ToString
                returnValue = dtResult.Compute(strExpression, strFilter)
                drTotalCost(dtResult.Columns(l).ColumnName) = returnValue

            Next
            drTotalCost("ACCOUNT_NO") = "Total Cost"

            '//Add one empty row
            drEmpty = dtResult.NewRow
            dtResult.Rows.Add(drEmpty)

            '//Add total cost
            dtResult.Rows.Add(drTotalCost)

            '//Add on empty row at index 0
            drEmpty = dtResult.NewRow
            dtResult.Rows.InsertAt(drEmpty, 0)

            'dtResult.TableName = arrRows(0)(strGroupColumnName).ToString & " " & arrRows(0)(strGroupColumnTitle).ToString
            dtResult.TableName = BGCommon.GetGroupCostTitle(arrRows(0)(strGroupColumnName).ToString)

            '//Return data table
            dsResult.Tables.Add(dtResult)

        Next

        Return dsResult

    End Function

    Private Function GeneratOriginalExcel(ByVal dsData As DataSet, ByVal dtColumns As DataTable) As Boolean
        Dim blnRet As Boolean = False
        Dim rowStartIndex As Integer = 8
        Dim colStartIndex As Integer = 7
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim rng As Excel.Range = Nothing

        excelApp = New Excel.Application()

        'excelApp.Visible = False
        'excelApp.UserControl = False

        wb = excelApp.Workbooks.Add(missing)

        '//Delete Worksheets
        If wb.Worksheets.Count > 1 Then
            For i As Integer = 1 To wb.Worksheets.Count - 1
                CType(wb.Worksheets(i), Excel.Worksheet).Delete()
            Next
        End If

        '//Set Style Value < 0 please fill color "Red"
        Dim style As Excel.Style = excelApp.ActiveWorkbook.Styles.Add("NewStyle")
        style.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)

        For intSheetCount As Integer = 0 To dsData.Tables.Count - 1

            If intSheetCount <> 0 Then
                wb.Sheets.Add()
            End If

            ws = CType(wb.ActiveSheet, Excel.Worksheet)
            Dim strSheetName As String = dsData.Tables(intSheetCount).TableName '.Substring(0, 6)
            ws.Name = strSheetName

            '//Setup DataColumn
            For i As Integer = 0 To dtColumns.Rows.Count - 1
                ws.Cells(colStartIndex, i + 1) = dtColumns.Rows(i)("Column_Title").ToString
                rng = ws.Range(ws.Cells(colStartIndex, i + 1), ws.Cells(colStartIndex, i + 1))
                rng.Font.Bold = True
                rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            Next

            '//Merge two columns row
            MergeColumnsCells(ws, 1, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 3, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 4, colStartIndex - 1, colStartIndex)

            MergeColumnsCells(ws, 11, colStartIndex - 1, colStartIndex)

            MergeColumnsCells(ws, 18, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 19, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 20, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 21, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 22, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 23, colStartIndex - 1, colStartIndex)

            '//Setup Item
            ws.Cells(colStartIndex - 1, 1) = "Item"
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(colStartIndex - 1, 2)).MergeCells = True
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(colStartIndex - 1, 2)).Font.Bold = True
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(colStartIndex - 1, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            '//Setup Original Title
            ws.Cells(colStartIndex - 1, 5) = "1st Half'" & Me.numYear.Text.ToString().Substring(2, 2)
            ws.Range(ws.Cells(colStartIndex - 1, 5), ws.Cells(colStartIndex - 1, 10)).MergeCells = True
            ws.Range(ws.Cells(colStartIndex - 1, 5), ws.Cells(colStartIndex - 1, 10)).Font.Bold = True
            ws.Range(ws.Cells(colStartIndex - 1, 5), ws.Cells(colStartIndex - 1, 10)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            '//Setup Original Title
            ws.Cells(colStartIndex - 1, 12) = "2nd Half'" & Me.numYear.Text.ToString().Substring(2, 2)
            ws.Range(ws.Cells(colStartIndex - 1, 12), ws.Cells(colStartIndex - 1, 17)).MergeCells = True
            ws.Range(ws.Cells(colStartIndex - 1, 12), ws.Cells(colStartIndex - 1, 17)).Font.Bold = True
            ws.Range(ws.Cells(colStartIndex - 1, 12), ws.Cells(colStartIndex - 1, 17)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            '//Setup Data
            For rowIndex As Integer = 0 To dsData.Tables(intSheetCount).Rows.Count - 1

                Dim row As DataRow = dsData.Tables(intSheetCount).Rows(rowIndex)

                '//If the column is "ACCOUNT_NO" Empty.
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
                        'ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter

                    ElseIf col.DataType Is System.Type.GetType("System.Decimal") Then

                        If row(col.ColumnName).ToString = String.Empty Then
                            row(col.ColumnName) = "0.00"
                        End If

                        'If CDbl(row(col.ColumnName).ToString) = 0 Then
                        '    excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = "-"
                        '    ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                        'Else
                        '    excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName)
                        '    ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).NumberFormat = "#,###.00"
                        'End If

                        ''//Add by Max 01/10/2012
                        ''//Set Style Value < 0 please fill color "Red"
                        'If CDec(row(col.ColumnName)) < 0 Then
                        '    ws.Range(ws.Cells(rowIndex + rowStartIndex, colIndex + 1), ws.Cells(rowIndex + rowStartIndex, colIndex + 1)).Style = style
                        'End If
                        ''//End Add by Max 01/10/2012

                        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName)
                        ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).NumberFormat = "#,##0.00"

                    Else
                        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName).ToString()
                    End If

                Next
            Next


            Dim rowMax As Integer = dsData.Tables(intSheetCount).Rows.Count + colStartIndex
            Dim colMax As Integer = dtColumns.Rows.Count
            Dim intAuthorizeStart As Integer = 23
            'Dim intAuthorizeEnd As Integer

            '//Setup Cost title
            Dim strCostTitle As String = dsData.Tables(intSheetCount).TableName
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).ClearContents()
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).MergeCells = True
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).Value = strCostTitle
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup budget order name column to be left align
            'ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowMax, 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            ws.Range(ws.Cells(rowStartIndex, 2), ws.Cells(rowMax, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Empry line
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).Value = ""
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup total cost
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).ClearContents()
            'ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).Value = "Total " & strCostTitle
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '// Set Borders
            rng = ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(rowMax, colMax))
            rng.Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.EntireColumn.AutoFit()


            '//Set Font
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(rowMax, colMax)).Font.Name = "Tahoma"
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(rowMax, colMax)).Font.Size = 10


            '//Setup Title & Title Font 
            SetupExcelTitle(ws, intAuthorizeStart)

            '// Add by Max 27/09/2012
            ws.Range(ws.Cells(colStartIndex, 1), ws.Cells(rowMax, 1)).Columns.ColumnWidth = 10

            '//Set NumberFormat = "#,##0.00;[Red]-#,##0.00"
            colStartIndex = colStartIndex - 1
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).NumberFormat = "#,##0.00;[Red]-#,##0.00"

            '//Set Frame
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders.LineStyle = 1
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium

            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, 4)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 5), ws.Cells(rowMax, 10)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 11), ws.Cells(rowMax, 12)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 12), ws.Cells(rowMax, 13)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 14), ws.Cells(rowMax, 15)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium

            ws.Range(ws.Cells(2, 5), ws.Cells(rowMax, 10)).Columns.ColumnWidth = 12
            ws.Range(ws.Cells(2, 12), ws.Cells(rowMax, 17)).Columns.ColumnWidth = 12

            ws.Range(ws.Cells(2, 3), ws.Cells(rowMax, 4)).Columns.ColumnWidth = 13
            ws.Range(ws.Cells(2, 3), ws.Cells(rowMax, 4)).WrapText = True

            ws.Range(ws.Cells(2, 11), ws.Cells(rowMax, 11)).Columns.ColumnWidth = 13
            ws.Range(ws.Cells(2, 11), ws.Cells(rowMax, 11)).WrapText = True

            ws.Range(ws.Cells(2, 18), ws.Cells(rowMax, 23)).Columns.ColumnWidth = 13
            ws.Range(ws.Cells(2, 18), ws.Cells(rowMax, 23)).WrapText = True

            colStartIndex = colStartIndex + 1
            '// End Add by Max 27/09/2012

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

    Private Function InsertEstimateColumnData(ByRef dtColumns As DataTable, _
                                      ByVal strYear As String) As Boolean
        Dim strHalfYear As String = strYear.Substring(2, 2)
        Dim row As DataRow


        'SELECT
        'MAX_REV.BUDGET_YEAR, MAX_REV.PERIOD_TYPE, '' AS BUDGET_ORDER_NO, 
        'MAX_REV.ACCOUNT_NO, 
        row = dtColumns.NewRow()
        row("Column_Name") = "ACCOUNT_NO"
        row("Column_Title") = "Item"
        dtColumns.Rows.Add(row)

        'MAX_REV.ACCOUNT_NAME, 
        row = dtColumns.NewRow()
        row("Column_Name") = "ACCOUNT_NAME"
        row("Column_Title") = ""
        dtColumns.Rows.Add(row)

        'MAX_REV.COST, MAX_REV.EXPENSE_TYPE,
        'MAX(MAX_REV.REV_NO) AS REV_NO,
        'SUM(ISNULL(ACTUAL_DATA.H1,0)) AS ACTUAL_1ST_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "ACTUAL_1ST_HALF"
        row("Column_Title") = "Actual 1st Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(REVISE_BUDGET.H2,0)) AS REVISE_BUDGET_2ND_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_BUDGET_2ND_HALF"
        row("Column_Title") = "Original 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M1,0)) AS M1,
        'SUM(ISNULL(MASTER_DATA.M2,0)) AS M2,
        'SUM(ISNULL(MASTER_DATA.M3,0)) AS M3,
        'SUM(ISNULL(MASTER_DATA.M4,0)) AS M4,
        'SUM(ISNULL(MASTER_DATA.M5,0)) AS M5,
        'SUM(ISNULL(MASTER_DATA.M6,0)) AS M6,
        'SUM(ISNULL(MASTER_DATA.M7,0)) AS M7,
        'SUM(ISNULL(MASTER_DATA.M8,0)) AS M8,
        'SUM(ISNULL(MASTER_DATA.M9,0)) AS M9,
        'SUM(ISNULL(MASTER_DATA.M10,0)) AS M10,
        'SUM(ISNULL(MASTER_DATA.M11,0)) AS M11,
        'SUM(ISNULL(MASTER_DATA.M12,0)) AS M12,

        'SUM(ISNULL(ACTUAL_DATA.M7,0))AS ACTUAL_JUL,
        row = dtColumns.NewRow()
        row("Column_Name") = "ACTUAL_JUL"
        row("Column_Title") = "Jul'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ACTUAL_DATA.M8,0))AS ACTUAL_AUG,
        row = dtColumns.NewRow()
        row("Column_Name") = "ACTUAL_AUG"
        row("Column_Title") = "Aug'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ACTUAL_DATA.M9,0))AS ACTUAL_SEP,
        row = dtColumns.NewRow()
        row("Column_Name") = "ACTUAL_SEP"
        row("Column_Title") = "Sep'" & strHalfYear
        dtColumns.Rows.Add(row)


        'SUM(ISNULL(MASTER_DATA.M10,0)) AS ESTIMATE_OCT,
        row = dtColumns.NewRow()
        row("Column_Name") = "ESTIMATE_OCT"
        row("Column_Title") = "Oct'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M11,0)) AS ESTIMATE_NOV,
        row = dtColumns.NewRow()
        row("Column_Name") = "ESTIMATE_NOV"
        row("Column_Title") = "Nov'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M12,0)) AS ESTIMATE_DEC,
        row = dtColumns.NewRow()
        row("Column_Name") = "ESTIMATE_DEC"
        row("Column_Title") = "Dec'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ACTUAL_DATA.M7,0) + ISNULL(ACTUAL_DATA.M8,0) + ISNULL(ACTUAL_DATA.M9,0) + ISNULL(MASTER_DATA.M10,0) + ISNULL(MASTER_DATA.M11,0)+ ISNULL(MASTER_DATA.M12,0)) AS ESTIMATE_BUDGET_2ND_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "ESTIMATE_BUDGET_2ND_HALF"
        row("Column_Title") = "Estimate 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        '0 AS DIFFERENCE_2ND_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "DIFFERENCE_2ND_HALF"
        row("Column_Title") = "Diff 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        '0 AS ESTIMATE_BUDGET_TOTAL
        row = dtColumns.NewRow()
        row("Column_Name") = "ESTIMATE_BUDGET_TOTAL"
        row("Column_Title") = "Estimate Year'" & strYear
        dtColumns.Rows.Add(row)

        Return True
    End Function

    Private Function SetupEstimateGroupbyData(ByVal dsData As DataSet, _
                                              ByVal strGroupColumnName As String, _
                                              ByVal strGroupColumnTitle As String, _
                                              ByVal intDataColumnIndex As Integer) As DataSet

        Dim dsResult As DataSet = New DataSet

        Dim drEmpty As DataRow
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object

        Dim strScript As String = strGroupColumnName

        Dim dtGroups As DataTable = dsData.Tables(0).DefaultView.ToTable(True, strScript)
        Dim intGroupCount As Integer = dtGroups.Rows.Count

        For i As Integer = 0 To intGroupCount - 1

            Dim dtResult As DataTable = dsData.Tables(0).Clone

            Dim drCost As DataRow = dtResult.NewRow
            Dim drTotalCost As DataRow = dtResult.NewRow
            Dim intGroupTotalIndex As Integer = 0

            If CInt(dtGroups.Rows(i)(0)) = enumCost.ADMIN Then
                Debug.Print(CStr(enumCost.ADMIN))
            End If

            strScript = strGroupColumnName + " = '" + dtGroups.Rows(i)(0).ToString & "'"
            Dim arrRows As DataRow() = dsData.Tables(0).Select(strScript)

            Dim strGroupColumnName2 As String = "EXPENSE_TYPE"
            '//Seperate dataset data into several datatables according to group no
            strScript = strGroupColumnName2
            Dim dtGroups2 As DataTable = dsData.Tables(0).DefaultView.ToTable(True, strScript)
            Dim intGroupCount2 As Integer = dtGroups2.Rows.Count

            For n As Integer = 0 To intGroupCount2 - 1

                Dim drTotalExpenses As DataRow = dtResult.NewRow

                strScript = strGroupColumnName + " = " + dtGroups.Rows(i)(0).ToString
                strScript &= " AND "
                strScript &= strGroupColumnName2 + " = " + dtGroups2.Rows(n)(0).ToString
                Dim arrRows2 As DataRow() = dsData.Tables(0).Select(strScript)

                For j As Integer = 0 To arrRows2.Length - 1
                    Dim drow(dtResult.Columns.Count - 1) As Object
                    arrRows2(j).ItemArray.CopyTo(drow, 0)
                    dtResult.Rows.Add(drow)
                Next

                '//Calculate Horizontal Column
                For m As Integer = 0 To dtResult.Rows.Count - 1
                    For p As Integer = intDataColumnIndex To dtResult.Columns.Count - 1
                        Dim strColumnName As String = dtResult.Columns(p).ColumnName
                        If strColumnName = "DIFFERENCE_2ND_HALF" Then
                            'If strColumnName = "DIFFERENCE_2ND_HALF" Then
                            '{EstimateApplicant.ESTIMATE_BUDGET_2ND_HALF} - {EstimateApplicant.REVISE_BUDGET_2ND_HALF}
                            dtResult.Rows(m)![DIFFERENCE_2ND_HALF] = Convert.ToDecimal(Nz(dtResult.Rows(m)![ESTIMATE_BUDGET_2ND_HALF], 0.0)) - Convert.ToDecimal(Nz(dtResult.Rows(m)![REVISE_BUDGET_2ND_HALF], 0.0))
                        ElseIf strColumnName = "ESTIMATE_BUDGET_TOTAL" Then
                            'ElseIf strColumnName = "ESTIMATE_BUDGET_TOTAL" Then
                            '{EstimateApplicant.ACTUAL_1ST_HALF} + {EstimateApplicant.ESTIMATE_BUDGET_2ND_HALF}
                            dtResult.Rows(m)![ESTIMATE_BUDGET_TOTAL] = Convert.ToDecimal(Nz(dtResult.Rows(m)![ACTUAL_1ST_HALF], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![ESTIMATE_BUDGET_2ND_HALF], 0.0))
                        End If
                    Next
                Next
                dtResult.AcceptChanges()

                '//Calculate total for each group
                For k As Integer = intDataColumnIndex To dtResult.Columns.Count - 1

                    Dim strColumnName As String = dtResult.Columns(k).ColumnName
                    strExpression = "Sum(" + strColumnName + ")"
                    strFilter = strGroupColumnName + " = " + dtGroups.Rows(i)(0).ToString
                    strFilter &= " AND "
                    strFilter &= strGroupColumnName2 + " = " + dtGroups2.Rows(n)(0).ToString
                    returnValue = dtResult.Compute(strExpression, strFilter)
                    drTotalExpenses(dtResult.Columns(k).ColumnName) = returnValue

                Next
                '//Setup Group header
                drTotalExpenses("ACCOUNT_NO") = GetGroupExpensesTitle(arrRows2(0)(strGroupColumnName2).ToString)

                '//Add total cost
                dtResult.Rows.InsertAt(drTotalExpenses, intGroupTotalIndex)

                '//Add one empty row
                drEmpty = dtResult.NewRow
                dtResult.Rows.Add(drEmpty)

                'intGroupTotalIndex = intGroupTotalIndex + dtResult.Rows.Count
                intGroupTotalIndex = intGroupTotalIndex + CInt(arrRows2.Length) + 2

            Next

            '//Calculate Total cost
            For l As Integer = intDataColumnIndex To dtResult.Columns.Count - 1
                Dim strColumnName As String = dtResult.Columns(l).ColumnName
                strExpression = "Sum(" + strColumnName + ")"
                strFilter = strGroupColumnName + " = " + dtGroups.Rows(i)(0).ToString
                returnValue = dtResult.Compute(strExpression, strFilter)
                drTotalCost(dtResult.Columns(l).ColumnName) = returnValue
            Next
            drTotalCost("ACCOUNT_NO") = "Total Cost"

            '//Add one empty row
            drEmpty = dtResult.NewRow
            dtResult.Rows.Add(drEmpty)

            '//Add total cost
            dtResult.Rows.Add(drTotalCost)

            '//Add on empty row at index 0
            drEmpty = dtResult.NewRow
            dtResult.Rows.InsertAt(drEmpty, 0)

            Dim strGroupCostTitle As String = GetGroupCostTitle(arrRows(0)(strGroupColumnName).ToString)
            dtResult.TableName = strGroupCostTitle

            '//Return data table
            dsResult.Tables.Add(dtResult)

        Next

        Return dsResult

    End Function

    Private Function GeneratEstimateExcel(ByVal dsData As DataSet, ByVal dtColumns As DataTable) As Boolean
        Dim blnRet As Boolean = False
        Dim rowStartIndex As Integer = 8
        Dim colStartIndex As Integer = 7
        Dim wb As Excel.Workbook = Nothing
        Dim ws As Excel.Worksheet = Nothing
        Dim rng As Excel.Range = Nothing

        excelApp = New Excel.Application()

        'excelApp.Visible = False
        'excelApp.UserControl = False

        wb = excelApp.Workbooks.Add(missing)

        '//Delete Worksheets
        If wb.Worksheets.Count > 1 Then
            For i As Integer = 1 To wb.Worksheets.Count - 1
                CType(wb.Worksheets(i), Excel.Worksheet).Delete()
            Next
        End If

        '//Set Style Value < 0 please fill color "Red"
        Dim style As Excel.Style = excelApp.ActiveWorkbook.Styles.Add("NewStyle")
        style.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)

        For intSheetCount As Integer = 0 To dsData.Tables.Count - 1

            If intSheetCount <> 0 Then
                wb.Sheets.Add()
            End If

            ws = CType(wb.ActiveSheet, Excel.Worksheet)
            Dim strSheetName As String = dsData.Tables(intSheetCount).TableName '.Substring(0, 6)
            ws.Name = strSheetName

            '//Setup DataColumn
            For i As Integer = 0 To dtColumns.Rows.Count - 1
                ws.Cells(colStartIndex, i + 1) = dtColumns.Rows(i)("Column_Title").ToString
                ws.Range(ws.Cells(colStartIndex, i + 1), ws.Cells(colStartIndex, i + 1)).Font.Bold = True
                ws.Range(ws.Cells(colStartIndex, i + 1), ws.Cells(colStartIndex, i + 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            Next

            '//Merge two columns row
            MergeColumnsCells(ws, 1, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 3, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 4, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 11, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 12, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 13, colStartIndex - 1, colStartIndex)

            '//Setup Item
            ws.Cells(colStartIndex - 1, 1) = "Item"
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(colStartIndex - 1, 2)).MergeCells = True
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(colStartIndex - 1, 2)).Font.Bold = True
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(colStartIndex - 1, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            '//Setup Actual Title
            ws.Cells(colStartIndex - 1, 5) = "Actual"
            ws.Range(ws.Cells(colStartIndex - 1, 5), ws.Cells(colStartIndex - 1, 7)).MergeCells = True
            ws.Range(ws.Cells(colStartIndex - 1, 5), ws.Cells(colStartIndex - 1, 7)).Font.Bold = True
            ws.Range(ws.Cells(colStartIndex - 1, 5), ws.Cells(colStartIndex - 1, 7)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            '//Setup Estimate Title
            ws.Cells(colStartIndex - 1, 8) = "Estimate"
            ws.Range(ws.Cells(colStartIndex - 1, 8), ws.Cells(colStartIndex - 1, 10)).MergeCells = True
            ws.Range(ws.Cells(colStartIndex - 1, 8), ws.Cells(colStartIndex - 1, 10)).Font.Bold = True
            ws.Range(ws.Cells(colStartIndex - 1, 8), ws.Cells(colStartIndex - 1, 10)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            '//Setup Data
            For rowIndex As Integer = 0 To dsData.Tables(intSheetCount).Rows.Count - 1

                Dim row As DataRow = dsData.Tables(intSheetCount).Rows(rowIndex)

                '//If the column is "ACCOUNT_NO" Empty.
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
                        'ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter

                    ElseIf col.DataType Is System.Type.GetType("System.Decimal") Then

                        If row(col.ColumnName).ToString = String.Empty Then
                            row(col.ColumnName) = "0.00"
                        End If

                        'If CDbl(row(col.ColumnName).ToString) = 0 Then
                        '    excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = "-"
                        '    ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                        'Else
                        '    excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName)
                        '    ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).NumberFormat = "#,###.00"
                        'End If

                        ''//Add by Max 01/10/2012
                        ''//Set Style Value < 0 please fill color "Red"
                        'If CDec(row(col.ColumnName)) < 0 Then
                        '    ws.Range(ws.Cells(rowIndex + rowStartIndex, colIndex + 1), ws.Cells(rowIndex + rowStartIndex, colIndex + 1)).Style = style
                        'End If
                        ''//End Add by Max 01/10/2012

                        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName)
                        ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).NumberFormat = "#,##0.00"

                    Else
                        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName).ToString()
                    End If

                Next
            Next

            Dim rowMax As Integer = dsData.Tables(intSheetCount).Rows.Count + colStartIndex
            Dim colMax As Integer = dtColumns.Rows.Count
            Dim intAuthorizeStart As Integer = 13

            '//Setup Cost title
            Dim strCostTitle As String = dsData.Tables(intSheetCount).TableName
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).ClearContents()
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).MergeCells = True
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).Value = strCostTitle
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup budget order name column to be left align
            'ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowMax, 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            ws.Range(ws.Cells(rowStartIndex, 2), ws.Cells(rowMax, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Empry line
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).Value = ""
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup total cost
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).ClearContents()
            'ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).Value = "Total " & strCostTitle
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '// Set Borders
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(rowMax, colMax)).EntireColumn.AutoFit()

            '//Set Font
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(rowMax, colMax)).Font.Name = "Tahoma"
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(rowMax, colMax)).Font.Size = 10

            '//Setup Title & Title Font 
            SetupExcelTitle(ws, intAuthorizeStart)

            '// Add by Max 27/09/2012
            ws.Range(ws.Cells(colStartIndex, 1), ws.Cells(rowMax, 1)).Columns.ColumnWidth = 10

            '//Set NumberFormat = "#,##0.00;[Red]-#,##0.00"
            colStartIndex = colStartIndex - 1
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).NumberFormat = "#,##0.00;[Red]-#,##0.00"

            '//Set Frame
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders.LineStyle = 1
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium

            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, 4)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 5), ws.Cells(rowMax, 7)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 8), ws.Cells(rowMax, 10)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 11), ws.Cells(rowMax, 13)).Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium
            colStartIndex = colStartIndex + 1
            '// End Add by Max 27/09/2012

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

    Private Function InsertReviseColumnData(ByRef dtColumns As DataTable, _
                                            ByVal strYear As String) As Boolean

        Dim strHalfYear As String = strYear.Substring(2, 2)
        Dim row As DataRow

        '   SELECT
        'MAX_REV.BUDGET_YEAR, MAX_REV.PERIOD_TYPE, '' AS BUDGET_ORDER_NO, MAX_REV.ACCOUNT_NO, MAX_REV.ACCOUNT_NAME, MAX_REV.EXPENSE_TYPE, MAX_REV.COST,
        'MAX(MAX_REV.REV_NO) AS REV_NO,
        row = dtColumns.NewRow()
        row("Column_Name") = "ACCOUNT_NO"
        row("Column_Title") = ""
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "ACCOUNT_NAME"
        row("Column_Title") = ""
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M1, 0)) AS M1,
        'SUM(ISNULL(MASTER_DATA.M2, 0)) AS M2,
        'SUM(ISNULL(MASTER_DATA.M3, 0)) AS M3,
        'SUM(ISNULL(MASTER_DATA.M4, 0)) AS M4,
        'SUM(ISNULL(MASTER_DATA.M5, 0)) AS M5,
        'SUM(ISNULL(MASTER_DATA.M6, 0)) AS M6,


        'SUM(ISNULL(ORIGINAL_BUDGET.H1,0)) AS ORIGINAL_1ST_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "ORIGINAL_1ST_HALF"
        row("Column_Title") = "Original 1st Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ACTUAL_DATA.M1,0)) AS ACTUAL_JAN,
        row = dtColumns.NewRow()
        row("Column_Name") = "ACTUAL_JAN"
        row("Column_Title") = "Jan'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ACTUAL_DATA.M2,0)) AS ACTUAL_FEB,
        row = dtColumns.NewRow()
        row("Column_Name") = "ACTUAL_FEB"
        row("Column_Title") = "Feb'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ACTUAL_DATA.M3,0)) AS ACTUAL_MAR,
        row = dtColumns.NewRow()
        row("Column_Name") = "ACTUAL_MAR"
        row("Column_Title") = "Mar'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ESTIMATE_BUDGET.M4,0)) AS ESTIMATE_APR,
        row = dtColumns.NewRow()
        row("Column_Name") = "ESTIMATE_APR"
        row("Column_Title") = "Apr'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ESTIMATE_BUDGET.M5,0)) AS ESTIMATE_MAY,
        row = dtColumns.NewRow()
        row("Column_Name") = "ESTIMATE_MAY"
        row("Column_Title") = "May'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ESTIMATE_BUDGET.M6,0)) AS ESTIMATE_JUN,
        row = dtColumns.NewRow()
        row("Column_Name") = "ESTIMATE_JUN"
        row("Column_Title") = "Jun'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ACTUAL_DATA.M1,0) + ISNULL(ACTUAL_DATA.M2,0) + ISNULL(ACTUAL_DATA.M2,0) + ISNULL(ESTIMATE_BUDGET.M4,0) + ISNULL(ESTIMATE_BUDGET.M5,0) + ISNULL(ESTIMATE_BUDGET.M6,0)) AS ESTIMATE_1ST_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "ESTIMATE_1ST_HALF"
        row("Column_Title") = "Estimate 1st Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "DIFF_1ST_HALF"
        row("Column_Title") = "Diff 1st Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ORIGINAL_BUDGET.H2,0)) AS ORIGINAL_2ND_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "ORIGINAL_2ND_HALF"
        row("Column_Title") = "Original 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(row)


        'SUM(ISNULL(MASTER_DATA.M7, 0)) AS M7,
        row = dtColumns.NewRow()
        row("Column_Name") = "M7"
        row("Column_Title") = "Jul'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M8, 0)) AS M8,
        row = dtColumns.NewRow()
        row("Column_Name") = "M8"
        row("Column_Title") = "Aug'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M9, 0)) AS M9,
        row = dtColumns.NewRow()
        row("Column_Name") = "M9"
        row("Column_Title") = "Sep'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M10, 0)) AS M10,
        row = dtColumns.NewRow()
        row("Column_Name") = "M10"
        row("Column_Title") = "Oct'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M11, 0)) AS M11,
        row = dtColumns.NewRow()
        row("Column_Name") = "M11"
        row("Column_Title") = "Nov'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M12, 0)) AS M12,
        row = dtColumns.NewRow()
        row("Column_Name") = "M12"
        row("Column_Title") = "Dec'" & strHalfYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_2ND_HALF"
        row("Column_Title") = "Revise 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "DIFF_2ND_HALF"
        row("Column_Title") = "Diff 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_TOTAL_YEAR"
        row("Column_Title") = "Revise Year'" & strYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "DIFF_TOTAL_YEAR"
        row("Column_Title") = "Diff Year'" & strYear
        dtColumns.Rows.Add(row)


        'SUM(ISNULL(MASTER_DATA.RRT1 ,0)) AS RRT1,
        'SUM(ISNULL(MASTER_DATA.RRT2 ,0)) AS RRT2,
        'SUM(ISNULL(MASTER_DATA.RRT3 ,0)) AS RRT3,
        'SUM(ISNULL(MASTER_DATA.RRT4 ,0)) AS RRT4,
        'SUM(ISNULL(MASTER_DATA.RRT5 ,0)) AS RRT5

        'FROM (
        'SELECT BUDGET_DATA.BUDGET_ORDER_NO, BUDGET_DATA.BUDGET_YEAR, BUDGET_DATA.PERIOD_TYPE,
        'BUDGET_ORDER.ACCOUNT_NO, ACCOUNT.ACCOUNT_NAME,
        'BUDGET_ORDER.BUDGET_TYPE, BUDGET_ORDER.EXPENSE_TYPE, BUDGET_ORDER.COST,
        'MAX(BUDGET_DATA.REV_NO) AS REV_NO
        'FROM BG_T_BUDGET_DATA AS BUDGET_DATA
        'INNER JOIN BG_M_BUDGET_ORDER AS BUDGET_ORDER
        'ON BUDGET_ORDER.BUDGET_ORDER_NO = BUDGET_DATA.BUDGET_ORDER_NO
        'INNER JOIN BG_M_ACCOUNT AS ACCOUNT
        'ON ACCOUNT.ACCOUNT_NO = BUDGET_ORDER.ACCOUNT_NO
        'WHERE BUDGET_ORDER.BUDGET_TYPE = 'E' AND BUDGET_ORDER.ACTIVE_FLAG = 1 AND BUDGET_DATA.BUDGET_YEAR = @BudgetYear AND BUDGET_DATA.PERIOD_TYPE = @PeriodType

        'GROUP BY BUDGET_DATA.BUDGET_ORDER_NO, BUDGET_DATA.BUDGET_YEAR, BUDGET_DATA.PERIOD_TYPE,
        'BUDGET_ORDER.ACCOUNT_NO, ACCOUNT.ACCOUNT_NAME,
        'BUDGET_ORDER.BUDGET_TYPE, BUDGET_ORDER.EXPENSE_TYPE, BUDGET_ORDER.COST) AS MAX_REV

        'INNER JOIN BG_T_BUDGET_DATA AS MASTER_DATA
        'ON MAX_REV.BUDGET_YEAR = MASTER_DATA.BUDGET_YEAR AND MAX_REV.PERIOD_TYPE = MASTER_DATA.PERIOD_TYPE
        'AND MAX_REV.BUDGET_ORDER_NO = MASTER_DATA.BUDGET_ORDER_NO AND MAX_REV.REV_NO = MASTER_DATA.REV_NO

        'LEFT OUTER JOIN BG_T_UPLOAD_DATA AS ORIGINAL_BUDGET
        'ON MAX_REV.BUDGET_YEAR = ORIGINAL_BUDGET.BUDGET_YEAR
        'AND ORIGINAL_BUDGET.PERIOD_TYPE = @PeriodType
        'AND MAX_REV.BUDGET_ORDER_NO = ORIGINAL_BUDGET.BUDGET_ORDER_NO
        'AND ORIGINAL_BUDGET.DATA_TYPE = 1

        'LEFT OUTER JOIN BG_T_BUDGET_DATA AS ESTIMATE_BUDGET
        'ON MAX_REV.BUDGET_YEAR = ESTIMATE_BUDGET.BUDGET_YEAR
        'AND ESTIMATE_BUDGET.PERIOD_TYPE = @PeriodType
        'AND MAX_REV.BUDGET_ORDER_NO = ESTIMATE_BUDGET.BUDGET_ORDER_NO
        'AND MAX_REV.REV_NO = ESTIMATE_BUDGET.REV_NO

        'LEFT OUTER JOIN BG_T_UPLOAD_DATA AS ACTUAL_DATA
        'ON MAX_REV.BUDGET_YEAR = ACTUAL_DATA.BUDGET_YEAR
        'AND ACTUAL_DATA.PERIOD_TYPE = @PeriodType
        'AND MAX_REV.BUDGET_ORDER_NO = ACTUAL_DATA.BUDGET_ORDER_NO
        'AND ACTUAL_DATA.DATA_TYPE = 2

        'WHERE
        '( ISNULL(ACTUAL_DATA.H1,0) &lt;&gt; 0
        'OR ISNULL(ACTUAL_DATA.H2,0) &lt;&gt; 0
        'OR ISNULL(ACTUAL_DATA.M1,0) &lt;&gt; 0
        'OR ISNULL(ACTUAL_DATA.M2,0) &lt;&gt; 0
        'OR ISNULL(ACTUAL_DATA.M3,0) &lt;&gt; 0
        'OR ISNULL(MASTER_DATA.M4, 0) &lt;&gt; 0
        'OR ISNULL(MASTER_DATA.M5, 0) &lt;&gt; 0
        'OR ISNULL(MASTER_DATA.M6, 0) &lt;&gt; 0
        'OR ISNULL(MASTER_DATA.M7, 0) &lt;&gt; 0
        'OR ISNULL(MASTER_DATA.M8, 0) &lt;&gt; 0
        'OR ISNULL(MASTER_DATA.M9, 0) &lt;&gt; 0
        'OR ISNULL(MASTER_DATA.M10, 0) &lt;&gt; 0
        'OR ISNULL(MASTER_DATA.M11, 0) &lt;&gt; 0
        'OR ISNULL(MASTER_DATA.M12, 0) &lt;&gt; 0
        ')

        'GROUP BY MAX_REV.BUDGET_YEAR, MAX_REV.PERIOD_TYPE, MAX_REV.ACCOUNT_NO, MAX_REV.ACCOUNT_NAME, MAX_REV.EXPENSE_TYPE, MAX_REV.COST
        'ORDER BY MAX_REV.COST, MAX_REV.EXPENSE_TYPE, MAX_REV.ACCOUNT_NO


        Return True
    End Function

    Private Function InsertReviseMTPColumnData(ByRef dtColumns As DataTable, _
                                               ByVal strYear As String) As Boolean

        Dim strHalfYear As String = strYear.Substring(2, 2)
        Dim row As DataRow

        '   SELECT
        'MAX_REV.BUDGET_YEAR, MAX_REV.PERIOD_TYPE, '' AS BUDGET_ORDER_NO, MAX_REV.ACCOUNT_NO, MAX_REV.ACCOUNT_NAME, MAX_REV.EXPENSE_TYPE, MAX_REV.COST,
        'MAX(MAX_REV.REV_NO) AS REV_NO,
        row = dtColumns.NewRow()
        row("Column_Name") = "ACCOUNT_NO"
        row("Column_Title") = ""
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "ACCOUNT_NAME"
        row("Column_Title") = ""
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M1, 0)) AS M1,
        'SUM(ISNULL(MASTER_DATA.M2, 0)) AS M2,
        'SUM(ISNULL(MASTER_DATA.M3, 0)) AS M3,
        'SUM(ISNULL(MASTER_DATA.M4, 0)) AS M4,
        'SUM(ISNULL(MASTER_DATA.M5, 0)) AS M5,
        'SUM(ISNULL(MASTER_DATA.M6, 0)) AS M6,


        ''SUM(ISNULL(ORIGINAL_BUDGET.H1,0)) AS ORIGINAL_1ST_HALF,
        'row = dtColumns.NewRow()
        'row("Column_Name") = "ORIGINAL_1ST_HALF"
        'row("Column_Title") = "Original 1st Half'" & strHalfYear
        'dtColumns.Rows.Add(row)

        ''SUM(ISNULL(ACTUAL_DATA.M1,0)) AS ACTUAL_JAN,
        'row = dtColumns.NewRow()
        'row("Column_Name") = "ACTUAL_JAN"
        'row("Column_Title") = "Jan'" & strHalfYear
        'dtColumns.Rows.Add(row)

        ''SUM(ISNULL(ACTUAL_DATA.M2,0)) AS ACTUAL_FEB,
        'row = dtColumns.NewRow()
        'row("Column_Name") = "ACTUAL_FEB"
        'row("Column_Title") = "Feb'" & strHalfYear
        'dtColumns.Rows.Add(row)

        ''SUM(ISNULL(ACTUAL_DATA.M3,0)) AS ACTUAL_MAR,
        'row = dtColumns.NewRow()
        'row("Column_Name") = "ACTUAL_MAR"
        'row("Column_Title") = "Mar'" & strHalfYear
        'dtColumns.Rows.Add(row)

        ''SUM(ISNULL(ESTIMATE_BUDGET.M4,0)) AS ESTIMATE_APR,
        'row = dtColumns.NewRow()
        'row("Column_Name") = "ESTIMATE_APR"
        'row("Column_Title") = "Apr'" & strHalfYear
        'dtColumns.Rows.Add(row)

        ''SUM(ISNULL(ESTIMATE_BUDGET.M5,0)) AS ESTIMATE_MAY,
        'row = dtColumns.NewRow()
        'row("Column_Name") = "ESTIMATE_MAY"
        'row("Column_Title") = "May'" & strHalfYear
        'dtColumns.Rows.Add(row)

        ''SUM(ISNULL(ESTIMATE_BUDGET.M6,0)) AS ESTIMATE_JUN,
        'row = dtColumns.NewRow()
        'row("Column_Name") = "ESTIMATE_JUN"
        'row("Column_Title") = "Jun'" & strHalfYear
        'dtColumns.Rows.Add(row)

        ''SUM(ISNULL(ACTUAL_DATA.M1,0) + ISNULL(ACTUAL_DATA.M2,0) + ISNULL(ACTUAL_DATA.M2,0) + ISNULL(ESTIMATE_BUDGET.M4,0) + ISNULL(ESTIMATE_BUDGET.M5,0) + ISNULL(ESTIMATE_BUDGET.M6,0)) AS ESTIMATE_1ST_HALF,
        'row = dtColumns.NewRow()
        'row("Column_Name") = "ESTIMATE_1ST_HALF"
        'row("Column_Title") = "Estimate 1st Half'" & strHalfYear
        'dtColumns.Rows.Add(row)

        'row = dtColumns.NewRow()
        'row("Column_Name") = "DIFF_1ST_HALF"
        'row("Column_Title") = "Diff 1st Half'" & strHalfYear
        'dtColumns.Rows.Add(row)

        'SUM(ISNULL(ORIGINAL_BUDGET.H2,0)) AS ORIGINAL_2ND_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "ORIGINAL_2ND_HALF"
        row("Column_Title") = "Original 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(row)


        'SUM(ISNULL(MASTER_DATA.M7, 0)) AS M7,
        row = dtColumns.NewRow()
        row("Column_Name") = "M7"
        row("Column_Title") = "Jul'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M8, 0)) AS M8,
        row = dtColumns.NewRow()
        row("Column_Name") = "M8"
        row("Column_Title") = "Aug'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M9, 0)) AS M9,
        row = dtColumns.NewRow()
        row("Column_Name") = "M9"
        row("Column_Title") = "Sep'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M10, 0)) AS M10,
        row = dtColumns.NewRow()
        row("Column_Name") = "M10"
        row("Column_Title") = "Oct'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M11, 0)) AS M11,
        row = dtColumns.NewRow()
        row("Column_Name") = "M11"
        row("Column_Title") = "Nov'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M12, 0)) AS M12,
        row = dtColumns.NewRow()
        row("Column_Name") = "M12"
        row("Column_Title") = "Dec'" & strHalfYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_2ND_HALF"
        row("Column_Title") = "Revise 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "DIFF_2ND_HALF"
        row("Column_Title") = "Diff 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_TOTAL_YEAR"
        row("Column_Title") = "Revise Year'" & strYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "DIFF_TOTAL_YEAR"
        row("Column_Title") = "Diff Year'" & strYear
        dtColumns.Rows.Add(row)

        ''SUM(ISNULL(MASTER_DATA.RRT1 ,0)) AS RRT1,
        'row = dtColumns.NewRow()
        'row("Column_Name") = "RRT1"
        'row("Column_Title") = "Y" & CInt(strYear) + 1
        'dtColumns.Rows.Add(row)

        ''SUM(ISNULL(MASTER_DATA.RRT2 ,0)) AS RRT2,
        'row = dtColumns.NewRow()
        'row("Column_Name") = "RRT2"
        'row("Column_Title") = "Y" & CInt(strYear) + 2
        'dtColumns.Rows.Add(row)

        ''SUM(ISNULL(MASTER_DATA.RRT3 ,0)) AS RRT3,
        'row = dtColumns.NewRow()
        'row("Column_Name") = "RRT3"
        'row("Column_Title") = "Y" & CInt(strYear) + 3
        'dtColumns.Rows.Add(row)

        ''SUM(ISNULL(MASTER_DATA.RRT4 ,0)) AS RRT4,
        'row = dtColumns.NewRow()
        'row("Column_Name") = "RRT4"
        'row("Column_Title") = "Y" & CInt(strYear) + 4
        'dtColumns.Rows.Add(row)

        ''SUM(ISNULL(MASTER_DATA.RRT5 ,0)) AS RRT5
        'row = dtColumns.NewRow()
        'row("Column_Name") = "RRT5"
        'row("Column_Title") = "Y" & CInt(strYear) + 5
        'dtColumns.Rows.Add(row)


        Return True
    End Function

    Private Function SetupReviseGroupbyData(ByVal dsData As DataSet, _
                                            ByVal strGroupColumnName As String, _
                                            ByVal strGroupColumnTitle As String, _
                                            ByVal intDataColumnIndex As Integer, _
                                            ByVal blnMTPBudget As Boolean) As DataSet

        Dim dsResult As DataSet = New DataSet

        Dim drEmpty As DataRow
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object

        Dim strScript As String = strGroupColumnName

        Dim dtGroups As DataTable = dsData.Tables(0).DefaultView.ToTable(True, strScript)
        Dim intGroupCount As Integer = dtGroups.Rows.Count

        For i As Integer = 0 To intGroupCount - 1

            Dim dtResult As DataTable = dsData.Tables(0).Clone

            '//-- Begin Add by S.Watcharapong 2011-05-19 
            '//Add column 
            Dim col As DataColumn = New DataColumn()
            col.ColumnName = "DIFF_1ST_HALF"
            col.DataType = Type.GetType("System.Decimal")
            col.DefaultValue = 0.0
            dtResult.Columns.Add(col)

            col = New DataColumn()
            col.ColumnName = "REVISE_2ND_HALF"
            col.DataType = Type.GetType("System.Decimal")
            col.DefaultValue = 0.0
            dtResult.Columns.Add(col)

            col = New DataColumn()
            col.ColumnName = "DIFF_2ND_HALF"
            col.DataType = Type.GetType("System.Decimal")
            col.DefaultValue = 0.0
            dtResult.Columns.Add(col)

            col = New DataColumn()
            col.ColumnName = "REVISE_TOTAL_YEAR"
            col.DataType = Type.GetType("System.Decimal")
            col.DefaultValue = 0.0
            dtResult.Columns.Add(col)

            col = New DataColumn()
            col.ColumnName = "DIFF_TOTAL_YEAR"
            col.DataType = Type.GetType("System.Decimal")
            col.DefaultValue = 0.0
            dtResult.Columns.Add(col)
            '//-- End Add 2011-05-19

            Dim drCost As DataRow = dtResult.NewRow
            Dim drTotalCost As DataRow = dtResult.NewRow
            Dim intGroupTotalIndex As Integer = 0

            strScript = strGroupColumnName + " = '" + dtGroups.Rows(i)(0).ToString & "'"
            Dim arrRows As DataRow() = dsData.Tables(0).Select(strScript)

            Dim strGroupColumnName2 As String = "EXPENSE_TYPE"
            '//Seperate dataset data into several datatables according to group no
            strScript = strGroupColumnName2
            Dim dtGroups2 As DataTable = dsData.Tables(0).DefaultView.ToTable(True, strScript)
            Dim intGroupCount2 As Integer = dtGroups2.Rows.Count

            For n As Integer = 0 To intGroupCount2 - 1

                Dim drTotalExpenses As DataRow = dtResult.NewRow

                strScript = strGroupColumnName + " = " + dtGroups.Rows(i)(0).ToString
                strScript &= " AND "
                strScript &= strGroupColumnName2 + " = " + dtGroups2.Rows(n)(0).ToString
                Dim arrRows2 As DataRow() = dsData.Tables(0).Select(strScript)

                If arrRows2.Length > 0 Then
                    For j As Integer = 0 To arrRows2.Length - 1
                        Dim drow(dtResult.Columns.Count - 1) As Object
                        arrRows2(j).ItemArray.CopyTo(drow, 0)
                        dtResult.Rows.Add(drow)
                    Next

                    For m As Integer = 0 To dtResult.Rows.Count - 1
                        For p As Integer = intDataColumnIndex To dtResult.Columns.Count - 1
                            Dim strColumnName As String = dtResult.Columns(p).ColumnName
                            If strColumnName = "DIFF_1ST_HALF" Then
                                'Diff 1st Half'11= {ReviseApplicant.ESTIMATE_1ST_HALF} - {ReviseApplicant.ORIGINAL_1ST_HALF}
                                dtResult.Rows(m)![DIFF_1ST_HALF] = Convert.ToDecimal(Nz(dtResult.Rows(m)![ESTIMATE_1ST_HALF], 0.0)) - Convert.ToDecimal(Nz(dtResult.Rows(m)![ORIGINAL_1ST_HALF], 0.0))
                            ElseIf strColumnName = "REVISE_2ND_HALF" Then
                                'Revise 2nd Half'11 = {ReviseApplicant.M7} + {ReviseApplicant.M8} + {ReviseApplicant.M9} + {ReviseApplicant.M10} + {ReviseApplicant.M11} + {ReviseApplicant.M12}
                                dtResult.Rows(m)![REVISE_2ND_HALF] = Convert.ToDecimal(Nz(dtResult.Rows(m)![M7], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![M8], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![M9], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![M10], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![M11], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![M12], 0.0))
                            ElseIf strColumnName = "DIFF_2ND_HALF" Then
                                'Diff 2nd Half'11 = {@Revise2ndHalf} - {ReviseApplicant.ORIGINAL_2ND_HALF}
                                dtResult.Rows(m)![DIFF_2ND_HALF] = Convert.ToDecimal(Nz(dtResult.Rows(m)![REVISE_2ND_HALF], 0.0)) - Convert.ToDecimal(Nz(dtResult.Rows(m)![ORIGINAL_2ND_HALF], 0.0))
                            ElseIf strColumnName = "REVISE_TOTAL_YEAR" Then
                                'Revise Year'2011 = {ReviseApplicant.ESTIMATE_1ST_HALF} + {@Revise2ndHalf}
                                dtResult.Rows(m)![REVISE_TOTAL_YEAR] = Convert.ToDecimal(Nz(dtResult.Rows(m)![ESTIMATE_1ST_HALF], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![REVISE_2ND_HALF], 0.0))
                            ElseIf strColumnName = "DIFF_TOTAL_YEAR" Then
                                'Diff Year'2011 = {@Diff1stHalf} + {@Diff2ndHalf}
                                dtResult.Rows(m)![DIFF_TOTAL_YEAR] = Convert.ToDecimal(Nz(dtResult.Rows(m)![DIFF_1ST_HALF], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![DIFF_2ND_HALF], 0.0))
                            End If
                        Next
                    Next
                    dtResult.AcceptChanges()

                    '//Calculate total for each group
                    For k As Integer = intDataColumnIndex To dtResult.Columns.Count - 1
                        Dim strColumnName As String = dtResult.Columns(k).ColumnName
                        strExpression = "Sum(" + strColumnName + ")"
                        strFilter = strGroupColumnName + " = " + dtGroups.Rows(i)(0).ToString
                        strFilter &= " AND "
                        strFilter &= strGroupColumnName2 + " = " + dtGroups2.Rows(n)(0).ToString
                        returnValue = dtResult.Compute(strExpression, strFilter)
                        drTotalExpenses(dtResult.Columns(k).ColumnName) = returnValue
                    Next
                    '//Setup Group header
                    drTotalExpenses("ACCOUNT_NO") = GetGroupExpensesTitle(arrRows2(0)(strGroupColumnName2).ToString)

                    '//Add total cost
                    dtResult.Rows.InsertAt(drTotalExpenses, intGroupTotalIndex)

                    '//Add one empty row
                    drEmpty = dtResult.NewRow
                    dtResult.Rows.Add(drEmpty)

                    'intGroupTotalIndex = intGroupTotalIndex + dtResult.Rows.Count
                    intGroupTotalIndex = intGroupTotalIndex + CInt(arrRows2.Length) + 2
                End If



            Next

            '//Calculate Total cost
            For l As Integer = intDataColumnIndex To dtResult.Columns.Count - 1
                Dim strColumnName As String = dtResult.Columns(l).ColumnName

                strExpression = "Sum(" + strColumnName + ")"
                strFilter = strGroupColumnName + " = " + dtGroups.Rows(i)(0).ToString
                returnValue = dtResult.Compute(strExpression, strFilter)
                drTotalCost(dtResult.Columns(l).ColumnName) = returnValue

            Next
            drTotalCost("ACCOUNT_NO") = "Total Cost"

            '//Add one empty row
            drEmpty = dtResult.NewRow
            dtResult.Rows.Add(drEmpty)

            '//Add total cost
            dtResult.Rows.Add(drTotalCost)

            '//Add on empty row at index 0
            drEmpty = dtResult.NewRow
            dtResult.Rows.InsertAt(drEmpty, 0)

            Dim strTotalCostTitle As String = GetGroupCostTitle(arrRows(0)(strGroupColumnName).ToString)
            dtResult.TableName = strTotalCostTitle

            '//Return data table
            dsResult.Tables.Add(dtResult)

        Next

        Return dsResult

    End Function

    Private Function GeneratReviseExcel(ByVal dsData As DataSet, ByVal dtColumns As DataTable, ByVal blnMTPBudget As Boolean) As Boolean
        Dim blnRet As Boolean = False
        Dim rowStartIndex As Integer = 8
        Dim colStartIndex As Integer = 7
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim rng As Excel.Range = Nothing

        excelApp = New Excel.Application()

        'excelApp.Visible = False
        'excelApp.UserControl = False

        wb = excelApp.Workbooks.Add(missing)

        '//Delete Worksheets
        If wb.Worksheets.Count > 1 Then
            For i As Integer = 1 To wb.Worksheets.Count - 1
                CType(wb.Worksheets(i), Excel.Worksheet).Delete()
            Next
        End If

        '//Set Style Value < 0 please fill color "Red"
        Dim style As Excel.Style = excelApp.ActiveWorkbook.Styles.Add("NewStyle")
        style.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)

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

            '//Merge two columns row
            MergeColumnsCells(ws, 1, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 3, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 10, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 11, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 12, colStartIndex - 1, colStartIndex)
            If blnMTPBudget = False Then

                MergeColumnsCells(ws, 19, colStartIndex - 1, colStartIndex)
                MergeColumnsCells(ws, 20, colStartIndex - 1, colStartIndex)
                MergeColumnsCells(ws, 21, colStartIndex - 1, colStartIndex)
                MergeColumnsCells(ws, 22, colStartIndex - 1, colStartIndex)
            Else
                MergeColumnsCells(ws, 13, colStartIndex - 1, colStartIndex)
            End If
            '//Setup Item
            ws.Cells(colStartIndex - 1, 1) = "Item"
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(colStartIndex - 1, 2)).MergeCells = True
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(colStartIndex - 1, 2)).Font.Bold = True
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(colStartIndex - 1, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            If blnMTPBudget = False Then

                '//Setup Revise & Estimate Title
                ws.Cells(colStartIndex - 1, 4) = "Actual"
                ws.Range(ws.Cells(colStartIndex - 1, 4), ws.Cells(colStartIndex - 1, 6)).MergeCells = True
                ws.Range(ws.Cells(colStartIndex - 1, 4), ws.Cells(colStartIndex - 1, 6)).Font.Bold = True
                ws.Range(ws.Cells(colStartIndex - 1, 4), ws.Cells(colStartIndex - 1, 6)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

                ws.Cells(colStartIndex - 1, 7) = "Estimate"
                ws.Range(ws.Cells(colStartIndex - 1, 7), ws.Cells(colStartIndex - 1, 9)).MergeCells = True
                ws.Range(ws.Cells(colStartIndex - 1, 7), ws.Cells(colStartIndex - 1, 9)).Font.Bold = True
                ws.Range(ws.Cells(colStartIndex - 1, 7), ws.Cells(colStartIndex - 1, 9)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

                ws.Cells(colStartIndex - 1, 13) = "Revise"
                ws.Range(ws.Cells(colStartIndex - 1, 13), ws.Cells(colStartIndex - 1, 18)).MergeCells = True
                ws.Range(ws.Cells(colStartIndex - 1, 13), ws.Cells(colStartIndex - 1, 18)).Font.Bold = True
                ws.Range(ws.Cells(colStartIndex - 1, 13), ws.Cells(colStartIndex - 1, 18)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            Else
                ws.Cells(colStartIndex - 1, 4) = "Revise"
                ws.Range(ws.Cells(colStartIndex - 1, 4), ws.Cells(colStartIndex - 1, 9)).MergeCells = True
                ws.Range(ws.Cells(colStartIndex - 1, 4), ws.Cells(colStartIndex - 1, 9)).Font.Bold = True
                ws.Range(ws.Cells(colStartIndex - 1, 4), ws.Cells(colStartIndex - 1, 9)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            End If

            '//Setup Data
            For rowIndex As Integer = 0 To dsData.Tables(intSheetCount).Rows.Count - 1

                Dim row As DataRow = dsData.Tables(intSheetCount).Rows(rowIndex)

                '//If the column is "ACCOUNT_NO" Empty.
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
                        'ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter

                    ElseIf col.DataType Is System.Type.GetType("System.Decimal") Then

                        If row(col.ColumnName).ToString = String.Empty Then
                            row(col.ColumnName) = "0.00"
                        End If

                        'If CDbl(row(col.ColumnName).ToString) = 0 Then
                        '    excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = "-"
                        '    ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                        'Else
                        '    excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName)
                        '    ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).NumberFormat = "#,###.00"
                        'End If

                        ''//Add by Max 01/10/2012
                        ''//Set Style Value < 0 please fill color "Red"
                        'If CDec(row(col.ColumnName)) < 0 Then
                        '    ws.Range(ws.Cells(rowIndex + rowStartIndex, colIndex + 1), ws.Cells(rowIndex + rowStartIndex, colIndex + 1)).Style = style
                        'End If
                        ''//End Add by Max 01/10/2012

                        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName)
                        ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).NumberFormat = "#,##0.00"

                    Else
                        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName).ToString()
                    End If

                Next
            Next

            Dim rowMax As Integer = dsData.Tables(intSheetCount).Rows.Count + colStartIndex
            Dim colMax As Integer = dtColumns.Rows.Count
            Dim intAuthorizeStart As Integer
            'Dim intAuthorizeEnd As Integer
            Dim intFontStart As Integer
            Dim intFontEnd As Integer

            '//Setup Cost title
            Dim strCostTitle As String = dsData.Tables(intSheetCount).TableName
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).ClearContents()
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).MergeCells = True
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).Value = strCostTitle
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup budget order name column to be left align
            'ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowMax, 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            ws.Range(ws.Cells(rowStartIndex, 2), ws.Cells(rowMax, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Empty line
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).Value = ""
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup total cost
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).ClearContents()
            'ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).Value = "Total " & strCostTitle
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '// Set Borders
            rng = ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(rowMax, colMax))
            rng.Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.EntireColumn.AutoFit()

            ''// MTP Budget
            'If blnMTPBudget = True Then
            '    '// Set Header
            '    ws.Cells(colStartIndex - 1, 23) = "MTP Budget"
            '    ws.Range(ws.Cells(colStartIndex - 1, 23), ws.Cells(colStartIndex - 1, 27)).MergeCells = True
            '    ws.Range(ws.Cells(colStartIndex - 1, 23), ws.Cells(colStartIndex - 1, 27)).Font.Bold = True
            '    ws.Range(ws.Cells(colStartIndex - 1, 23), ws.Cells(colStartIndex - 1, 27)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            '    Dim xColumn As Excel.Range = CType(ws.Columns(23, Type.Missing), Excel.Range)
            '    xColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Type.Missing)
            '    excelApp.Range(excelApp.Cells(colStartIndex, 23), excelApp.Cells(rowMax, 23)).Borders.LineStyle = 0
            '    excelApp.Range(excelApp.Cells(colStartIndex, 23), excelApp.Cells(rowMax, 23)).ColumnWidth = 2
            '    excelApp.Range(excelApp.Cells(colStartIndex, 23), excelApp.Cells(rowMax, 23)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
            '    excelApp.Range(excelApp.Cells(colStartIndex, 23), excelApp.Cells(rowMax, 23)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
            '    excelApp.Range(excelApp.Cells(colStartIndex, 23), excelApp.Cells(rowMax, 23)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = 0
            '    excelApp.Range(excelApp.Cells(colStartIndex, 23), excelApp.Cells(rowMax, 23)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = 0
            '    excelApp.Range(excelApp.Cells(colStartIndex - 1, 23), excelApp.Cells(rowMax, 23)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = 0

            '    intAuthorizeStart = 28
            '    'intAuthorizeStart = 19
            '    intFontStart = 1
            '    intFontEnd = 28
            '    'intFontEnd = 19

            '    ''//Delete Column
            '    'rng = ws.Range(ws.Cells(colStartIndex - 1, 3), ws.Cells(rowMax, 11))
            '    'rng.EntireColumn.Delete(missing)
            '    'System.Runtime.InteropServices.Marshal.ReleaseComObject(rng)

            'Else
            'intAuthorizeStart = 22
            'intFontStart = 1
            'intFontEnd = colMax
            'End If

            intAuthorizeStart = 22
            intFontStart = 1
            intFontEnd = colMax

            '//Set Font
            ws.Range(ws.Cells(colStartIndex - 1, intFontStart), ws.Cells(rowMax, intFontEnd)).Font.Name = "Tahoma"
            ws.Range(ws.Cells(colStartIndex - 1, intFontStart), ws.Cells(rowMax, intFontEnd)).Font.Size = 10

            '//Setup Title & Title Font 
            SetupExcelTitle(ws, intAuthorizeStart)

            '// Add by Max 27/09/2012
            ws.Range(ws.Cells(colStartIndex, 1), ws.Cells(rowMax, 1)).Columns.ColumnWidth = 10

            '//Set NumberFormat = "#,##0.00;[Red]-#,##0.00"
            colStartIndex = colStartIndex - 1
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).NumberFormat = "#,##0.00;[Red]-#,##0.00"

            '//Set Frame
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders.LineStyle = 1
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium

            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, 3)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            If blnMTPBudget = False Then
                ws.Range(ws.Cells(colStartIndex, 4), ws.Cells(rowMax, 6)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium

            End If
            ws.Range(ws.Cells(colStartIndex, 7), ws.Cells(rowMax, 9)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 10), ws.Cells(rowMax, 12)).Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 12), ws.Cells(rowMax, 12)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium

            If blnMTPBudget = False Then
                ws.Range(ws.Cells(colStartIndex, 13), ws.Cells(rowMax, 18)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                ws.Range(ws.Cells(colStartIndex, 19), ws.Cells(rowMax, 22)).Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium
            End If

            colStartIndex = colStartIndex + 1
            '// End Add by Max 27/09/2012

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

    Private Function InsertMTPColumnData(ByRef dtColumns As DataTable, _
                                            ByVal strYear As String) As Boolean

        Dim strHalfYear As String = strYear.Substring(2, 2)
        Dim row As DataRow

        row = dtColumns.NewRow()
        row("Column_Name") = "ACCOUNT_NO"
        row("Column_Title") = ""
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "ACCOUNT_NAME"
        row("Column_Title") = ""
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "REVYEAR"
        row("Column_Title") = "Original Year'" & strYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow
        row("Column_Name") = "PrevRRT1"
        row("Column_Title") = "MTP" & CInt(strYear) - 1 & " Year'" & strYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "DiffYear"
        row("Column_Title") = "Diff Year'" & strYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow
        row("Column_Name") = "RRT1"
        row("Column_Title") = "MTP" & CInt(strYear) & " Year'" & CInt(strYear) + 1
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow
        row("Column_Name") = "PrevRRT2"
        row("Column_Title") = "MTP" & CInt(strYear) - 1 & " Year'" & CInt(strYear) + 1
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow
        row("Column_Name") = "RRT2"
        row("Column_Title") = "MTP" & CInt(strYear) & " Year'" & CInt(strYear) + 2
        dtColumns.Rows.Add(row)

        'row = dtColumns.NewRow
        'row("Column_Name") = "PrevRRT3"
        'row("Column_Title") = "MTP" & CInt(strYear) - 1 & " Year'" & CInt(strYear) + 2
        'dtColumns.Rows.Add(row)

        'row = dtColumns.NewRow
        'row("Column_Name") = "RRT3"
        'row("Column_Title") = "MTP" & CInt(strYear) & " Year'" & CInt(strYear) + 3
        'dtColumns.Rows.Add(row)

        'row = dtColumns.NewRow
        'row("Column_Name") = "PrevRRT4"
        'row("Column_Title") = "MTP" & CInt(strYear) - 1 & " Year'" & CInt(strYear) + 3
        'dtColumns.Rows.Add(row)

        'row = dtColumns.NewRow
        'row("Column_Name") = "RRT4"
        'row("Column_Title") = "MTP" & CInt(strYear) & " Year'" & CInt(strYear) + 4
        'dtColumns.Rows.Add(row)

        'row = dtColumns.NewRow
        'row("Column_Name") = "PrevRRT5"
        'row("Column_Title") = "MTP" & CInt(strYear) - 1 & " Year'" & CInt(strYear) + 4
        'dtColumns.Rows.Add(row)


        'row = dtColumns.NewRow
        'row("Column_Name") = "RRT5"
        'row("Column_Title") = "MTP" & CInt(strYear) & " Year'" & CInt(strYear) + 5
        'dtColumns.Rows.Add(row)


        Return True
    End Function

    Private Function SetupMTPGroupbyData(ByVal dsData As DataSet, _
                                            ByVal strGroupColumnName As String, _
                                            ByVal strGroupColumnTitle As String, _
                                            ByVal intDataColumnIndex As Integer) As DataSet

        Dim dsResult As DataSet = New DataSet

        Dim drEmpty As DataRow
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object

        Dim strScript As String = strGroupColumnName

        Dim dtGroups As DataTable = dsData.Tables(0).DefaultView.ToTable(True, strScript)
        Dim intGroupCount As Integer = dtGroups.Rows.Count

        For i As Integer = 0 To intGroupCount - 1

            Dim dtResult As DataTable = dsData.Tables(0).Clone

            ''//-- Begin Add by S.Watcharapong 2011-05-19 
            ''//Add column 
            'Dim col As DataColumn = New DataColumn()
            'col.ColumnName = "DIFF_1ST_HALF"
            'col.DataType = Type.GetType("System.Decimal")
            'col.DefaultValue = 0.0
            'dtResult.Columns.Add(col)

            'col = New DataColumn()
            'col.ColumnName = "REVISE_2ND_HALF"
            'col.DataType = Type.GetType("System.Decimal")
            'col.DefaultValue = 0.0
            'dtResult.Columns.Add(col)

            'col = New DataColumn()
            'col.ColumnName = "DIFF_2ND_HALF"
            'col.DataType = Type.GetType("System.Decimal")
            'col.DefaultValue = 0.0
            'dtResult.Columns.Add(col)

            'col = New DataColumn()
            'col.ColumnName = "REVISE_TOTAL_YEAR"
            'col.DataType = Type.GetType("System.Decimal")
            'col.DefaultValue = 0.0
            'dtResult.Columns.Add(col)

            'col = New DataColumn()
            'col.ColumnName = "DIFF_TOTAL_YEAR"
            'col.DataType = Type.GetType("System.Decimal")
            'col.DefaultValue = 0.0
            'dtResult.Columns.Add(col)
            ''//-- End Add 2011-05-19

            Dim drCost As DataRow = dtResult.NewRow
            Dim drTotalCost As DataRow = dtResult.NewRow
            Dim intGroupTotalIndex As Integer = 0

            strScript = strGroupColumnName + " = '" + dtGroups.Rows(i)(0).ToString & "'"
            Dim arrRows As DataRow() = dsData.Tables(0).Select(strScript)

            Dim strGroupColumnName2 As String = "EXPENSE_TYPE"
            '//Seperate dataset data into several datatables according to group no
            strScript = strGroupColumnName2
            Dim dtGroups2 As DataTable = dsData.Tables(0).DefaultView.ToTable(True, strScript)
            Dim intGroupCount2 As Integer = dtGroups2.Rows.Count

            For n As Integer = 0 To intGroupCount2 - 1

                Dim drTotalExpenses As DataRow = dtResult.NewRow

                strScript = strGroupColumnName + " = " + dtGroups.Rows(i)(0).ToString
                strScript &= " AND "
                strScript &= strGroupColumnName2 + " = " + dtGroups2.Rows(n)(0).ToString
                Dim arrRows2 As DataRow() = dsData.Tables(0).Select(strScript)


                If arrRows2.Length > 0 Then
                    For j As Integer = 0 To arrRows2.Length - 1
                        Dim drow(dtResult.Columns.Count - 1) As Object
                        arrRows2(j).ItemArray.CopyTo(drow, 0)
                        dtResult.Rows.Add(drow)
                    Next

                    'For m As Integer = 0 To dtResult.Rows.Count - 1
                    '    For p As Integer = intDataColumnIndex To dtResult.Columns.Count - 1
                    '        Dim strColumnName As String = dtResult.Columns(p).ColumnName
                    '        If strColumnName = "DIFF_1ST_HALF" Then
                    '            'Diff 1st Half'11= {ReviseApplicant.ESTIMATE_1ST_HALF} - {ReviseApplicant.ORIGINAL_1ST_HALF}
                    '            dtResult.Rows(m)![DIFF_1ST_HALF] = Convert.ToDecimal(Nz(dtResult.Rows(m)![ESTIMATE_1ST_HALF], 0.0)) - Convert.ToDecimal(Nz(dtResult.Rows(m)![ORIGINAL_1ST_HALF], 0.0))
                    '        ElseIf strColumnName = "REVISE_2ND_HALF" Then
                    '            'Revise 2nd Half'11 = {ReviseApplicant.M7} + {ReviseApplicant.M8} + {ReviseApplicant.M9} + {ReviseApplicant.M10} + {ReviseApplicant.M11} + {ReviseApplicant.M12}
                    '            dtResult.Rows(m)![REVISE_2ND_HALF] = Convert.ToDecimal(Nz(dtResult.Rows(m)![M7], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![M8], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![M9], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![M10], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![M11], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![M12], 0.0))
                    '        ElseIf strColumnName = "DIFF_2ND_HALF" Then
                    '            'Diff 2nd Half'11 = {@Revise2ndHalf} - {ReviseApplicant.ORIGINAL_2ND_HALF}
                    '            dtResult.Rows(m)![DIFF_2ND_HALF] = Convert.ToDecimal(Nz(dtResult.Rows(m)![REVISE_2ND_HALF], 0.0)) - Convert.ToDecimal(Nz(dtResult.Rows(m)![ORIGINAL_2ND_HALF], 0.0))
                    '        ElseIf strColumnName = "REVISE_TOTAL_YEAR" Then
                    '            'Revise Year'2011 = {ReviseApplicant.ESTIMATE_1ST_HALF} + {@Revise2ndHalf}
                    '            dtResult.Rows(m)![REVISE_TOTAL_YEAR] = Convert.ToDecimal(Nz(dtResult.Rows(m)![ESTIMATE_1ST_HALF], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![REVISE_2ND_HALF], 0.0))
                    '        ElseIf strColumnName = "DIFF_TOTAL_YEAR" Then
                    '            'Diff Year'2011 = {@Diff1stHalf} + {@Diff2ndHalf}
                    '            dtResult.Rows(m)![DIFF_TOTAL_YEAR] = Convert.ToDecimal(Nz(dtResult.Rows(m)![DIFF_1ST_HALF], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![DIFF_2ND_HALF], 0.0))
                    '        End If
                    '    Next
                    'Next
                    dtResult.AcceptChanges()

                    '//Calculate total for each group
                    For k As Integer = intDataColumnIndex To dtResult.Columns.Count - 1
                        Dim strColumnName As String = dtResult.Columns(k).ColumnName
                        strExpression = "Sum(" + strColumnName + ")"
                        strFilter = strGroupColumnName + " = " + dtGroups.Rows(i)(0).ToString
                        strFilter &= " AND "
                        strFilter &= strGroupColumnName2 + " = " + dtGroups2.Rows(n)(0).ToString
                        returnValue = dtResult.Compute(strExpression, strFilter)
                        drTotalExpenses(dtResult.Columns(k).ColumnName) = returnValue
                    Next
                    '//Setup Group header
                    drTotalExpenses("ACCOUNT_NO") = GetGroupExpensesTitle(arrRows2(0)(strGroupColumnName2).ToString)

                    '//Add total cost
                    dtResult.Rows.InsertAt(drTotalExpenses, intGroupTotalIndex)

                    '//Add one empty row
                    drEmpty = dtResult.NewRow
                    dtResult.Rows.Add(drEmpty)

                    'intGroupTotalIndex = intGroupTotalIndex + dtResult.Rows.Count
                    intGroupTotalIndex = intGroupTotalIndex + CInt(arrRows2.Length) + 2

                End If



            Next

            '//Calculate Total cost
            For l As Integer = intDataColumnIndex To dtResult.Columns.Count - 1
                Dim strColumnName As String = dtResult.Columns(l).ColumnName

                strExpression = "Sum(" + strColumnName + ")"
                strFilter = strGroupColumnName + " = " + dtGroups.Rows(i)(0).ToString
                returnValue = dtResult.Compute(strExpression, strFilter)
                drTotalCost(dtResult.Columns(l).ColumnName) = returnValue

            Next
            drTotalCost("ACCOUNT_NO") = "Total Cost"

            '//Add one empty row
            drEmpty = dtResult.NewRow
            dtResult.Rows.Add(drEmpty)

            '//Add total cost
            dtResult.Rows.Add(drTotalCost)

            '//Add on empty row at index 0
            drEmpty = dtResult.NewRow
            dtResult.Rows.InsertAt(drEmpty, 0)

            Dim strTotalCostTitle As String = GetGroupCostTitle(arrRows(0)(strGroupColumnName).ToString)
            dtResult.TableName = strTotalCostTitle

            '//Return data table
            dsResult.Tables.Add(dtResult)

        Next

        Return dsResult

    End Function

    Private Function GeneratMTPExcel(ByVal dsData As DataSet, ByVal dtColumns As DataTable) As Boolean
        Dim blnRet As Boolean = False
        Dim rowStartIndex As Integer = 8
        Dim colStartIndex As Integer = 7
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim rng As Excel.Range = Nothing

        excelApp = New Excel.Application()

        'excelApp.Visible = False
        'excelApp.UserControl = False

        wb = excelApp.Workbooks.Add(missing)

        '//Delete Worksheets
        If wb.Worksheets.Count > 1 Then
            For i As Integer = 1 To wb.Worksheets.Count - 1
                CType(wb.Worksheets(i), Excel.Worksheet).Delete()
            Next
        End If

        '//Set Style Value < 0 please fill color "Red"
        Dim style As Excel.Style = excelApp.ActiveWorkbook.Styles.Add("NewStyle")
        style.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)

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

            '//Merge two columns row
            MergeColumnsCells(ws, 1, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 3, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 4, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 5, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 6, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 7, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 8, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 9, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 10, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 11, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 12, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 13, colStartIndex - 1, colStartIndex)
            MergeColumnsCells(ws, 14, colStartIndex - 1, colStartIndex)

            '//Setup Item
            ws.Cells(colStartIndex - 1, 1) = "Item"
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(colStartIndex - 1, 2)).MergeCells = True
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(colStartIndex - 1, 2)).Font.Bold = True
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(colStartIndex - 1, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            '//Setup Data
            For rowIndex As Integer = 0 To dsData.Tables(intSheetCount).Rows.Count - 1

                Dim row As DataRow = dsData.Tables(intSheetCount).Rows(rowIndex)

                '//If the column is "ACCOUNT_NO" Empty.
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
                        'ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter

                    ElseIf col.DataType Is System.Type.GetType("System.Decimal") Then

                        If row(col.ColumnName).ToString = String.Empty Then
                            row(col.ColumnName) = "0.00"
                        End If

                        'If CDbl(row(col.ColumnName).ToString) = 0 Then
                        '    excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = "-"
                        '    ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                        'Else
                        '    excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName)
                        '    ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).NumberFormat = "#,###.00"
                        'End If

                        ''//Add by Max 01/10/2012
                        ''//Set Style Value < 0 please fill color "Red"
                        'If CDec(row(col.ColumnName)) < 0 Then
                        '    ws.Range(ws.Cells(rowIndex + rowStartIndex, colIndex + 1), ws.Cells(rowIndex + rowStartIndex, colIndex + 1)).Style = style
                        'End If
                        ''//End Add by Max 01/10/2012

                        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName)
                        ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).NumberFormat = "#,##0.00"

                    Else
                        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName).ToString()
                    End If

                Next
            Next

            Dim rowMax As Integer = dsData.Tables(intSheetCount).Rows.Count + colStartIndex
            Dim colMax As Integer = dtColumns.Rows.Count
            Dim intAuthorizeStart As Integer
            'Dim intAuthorizeEnd As Integer
            Dim intFontStart As Integer
            Dim intFontEnd As Integer

            '//Setup Cost title
            Dim strCostTitle As String = dsData.Tables(intSheetCount).TableName
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).ClearContents()
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).MergeCells = True
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).Value = strCostTitle
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup budget order name column to be left align
            'ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowMax, 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            ws.Range(ws.Cells(rowStartIndex, 2), ws.Cells(rowMax, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Empty line
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).Value = ""
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup total cost
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).ClearContents()
            'ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).Value = "Total " & strCostTitle
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '// Set Borders
            rng = ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(rowMax, colMax))
            rng.Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous
            rng.EntireColumn.AutoFit()


            intAuthorizeStart = 8
            intFontStart = 1
            intFontEnd = colMax

            '//Set Font
            ws.Range(ws.Cells(colStartIndex - 1, intFontStart), ws.Cells(rowMax, intFontEnd)).Font.Name = "Tahoma"
            ws.Range(ws.Cells(colStartIndex - 1, intFontStart), ws.Cells(rowMax, intFontEnd)).Font.Size = 10

            '//Setup Title & Title Font 
            SetupExcelTitle(ws, intAuthorizeStart)

            '// Add by Max 27/09/2012
            ws.Range(ws.Cells(colStartIndex, 1), ws.Cells(rowMax, 1)).Columns.ColumnWidth = 10

            '//Set NumberFormat = "#,##0.00;[Red]-#,##0.00"
            colStartIndex = colStartIndex - 1
            'ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).NumberFormat = "#,##0.00;[Red]-#,##0.00"
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, 4)).NumberFormat = "#,##0.00;[Red]-#,##0.00"
            ws.Range(ws.Cells(colStartIndex, 6), ws.Cells(rowMax, colMax)).NumberFormat = "#,##0.00;[Red]-#,##0.00"

            '//Set Frame
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders.LineStyle = 1
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium

            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, 4)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 5), ws.Cells(rowMax, 5)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 6), ws.Cells(rowMax, 8)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium

            '//Set font color
            ws.Range(ws.Cells(colStartIndex, 4), ws.Cells(rowMax, 5)).Font.Color = RGB(128, 128, 128)
            ws.Range(ws.Cells(colStartIndex, 7), ws.Cells(rowMax, 7)).Font.Color = RGB(128, 128, 128)
            ws.Range(ws.Cells(colStartIndex, 9), ws.Cells(rowMax, 9)).Font.Color = RGB(128, 128, 128)
            ws.Range(ws.Cells(colStartIndex, 11), ws.Cells(rowMax, 11)).Font.Color = RGB(128, 128, 128)
            ws.Range(ws.Cells(colStartIndex, 13), ws.Cells(rowMax, 13)).Font.Color = RGB(128, 128, 128)
            colStartIndex = colStartIndex + 1
            '// End Add by Max 27/09/2012

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

    Private Function SetupExcelTitle(ByVal ws As Excel.Worksheet, ByVal intUnitPriceStart As Integer) As Boolean
        Dim strSubTitle As String

        If Me.numProjectNo.Value.ToString <> "1" Then
            strSubTitle = "Summary by Applicant : " + Me.cboPeriodType.Text + " " + Me.numYear.Value.ToString + " (Project No." + Me.numProjectNo.Value.ToString + ")"
        Else
            strSubTitle = "Summary by Applicant : " + Me.cboPeriodType.Text + " " + Me.numYear.Value.ToString
        End If

        ws.Range(ws.Cells(1, 1), ws.Cells(1, 4)).Font.Bold = True
        ws.Range(ws.Cells(1, 1), ws.Cells(1, 4)).Font.Size = 12
        ws.Range(ws.Cells(1, 1), ws.Cells(1, 4)).Font.Name = "Tahoma"
        ws.Range(ws.Cells(1, 1), ws.Cells(1, 4)).MergeCells = True
        ws.Range(ws.Cells(1, 1), ws.Cells(1, 4)).Value = "Bridgestone Tire Manufacturing (Thailand) Co.,Ltd."

        '//Setup subTitle  
        ws.Range(ws.Cells(2, 1), ws.Cells(2, 5)).Font.Bold = True
        ws.Range(ws.Cells(2, 1), ws.Cells(2, 5)).Font.Size = 11
        ws.Range(ws.Cells(2, 1), ws.Cells(2, 5)).Font.Name = "Tahoma"
        ws.Range(ws.Cells(2, 1), ws.Cells(2, 5)).MergeCells = True
        ws.Range(ws.Cells(2, 1), ws.Cells(2, 5)).Value = strSubTitle

        ws.Range(ws.Cells(1, 1), ws.Cells(2, 5)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft '// Add by Max 27/09/2012

        ''//Setup GroupName  
        'ws.Range(excelApp.Cells(4, 1), excelApp.Cells(4, 3)).Font.Bold = True
        'ws.Range(excelApp.Cells(4, 1), excelApp.Cells(4, 3)).Font.Size = 12
        'ws.Range(excelApp.Cells(4, 1), excelApp.Cells(4, 3)).Font.Italic = True
        'ws.Range(excelApp.Cells(4, 1), excelApp.Cells(4, 3)).Font.Underline = True
        'ws.Range(excelApp.Cells(4, 1), excelApp.Cells(4, 3)).Font.Name = "Calibri"
        'ws.Range(excelApp.Cells(4, 1), excelApp.Cells(4, 3)).MergeCells = True
        'ws.Range(excelApp.Cells(4, 1), excelApp.Cells(4, 3)).Value = dsData.Tables(intSheetCount).TableName

        '//Setup unit price
        ws.Range(excelApp.Cells(4, intUnitPriceStart), excelApp.Cells(4, intUnitPriceStart)).Font.Name = "Tahoma"
        ws.Range(excelApp.Cells(4, intUnitPriceStart), excelApp.Cells(4, intUnitPriceStart)).Font.Bold = True
        ws.Range(excelApp.Cells(4, intUnitPriceStart), excelApp.Cells(4, intUnitPriceStart)).Font.Underline = True
        ws.Range(excelApp.Cells(4, intUnitPriceStart), excelApp.Cells(4, intUnitPriceStart)).Font.Size = 11
        ws.Range(excelApp.Cells(4, intUnitPriceStart), excelApp.Cells(4, intUnitPriceStart)).MergeCells = True
        ws.Range(excelApp.Cells(4, intUnitPriceStart), excelApp.Cells(4, intUnitPriceStart)).Value = "Unit : K.Baht"
        ws.Range(ws.Cells(4, intUnitPriceStart), ws.Cells(4, intUnitPriceStart)).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

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
            Or row(col).ToString() = P_EXPENSE_TYPE_FIXED _
            Or row(col).ToString() = "Total Cost" Then
                ws.Range(excelApp.Cells(rowIndex + rowStartIndex, 1), excelApp.Cells(rowIndex + rowStartIndex, 2)).MergeCells = True
                ws.Range(excelApp.Cells(rowIndex + rowStartIndex, 1), excelApp.Cells(rowIndex + rowStartIndex, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colMax)).Font.Bold = True
            Else
                ws.Range(excelApp.Cells(rowIndex + rowStartIndex, 1), excelApp.Cells(rowIndex + rowStartIndex, 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                ws.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colMax)).Font.Bold = False
            End If
        End If
    End Sub

    Private Sub LoadRevNo()
        If Me.cboRevNo.Visible = True Then

            If Me.cboPeriodType.SelectedIndex < 0 OrElse _
                Me.numProjectNo.Value <= 0 OrElse _
                Me.numYear.Value <= 0 Then

                Me.cboRevNo.DataSource = Nothing
                Exit Sub

            End If

            Dim strProjectNo = Me.numProjectNo.Value.ToString
            If Not strProjectNo Is Nothing And strProjectNo <> String.Empty And strProjectNo <> "System.Data.DataRowView" Then

                myClsBG0310BL.BudgetYear = Me.numYear.Value.ToString
                myClsBG0310BL.PeriodType = Me.cboPeriodType.SelectedValue.ToString
                myClsBG0310BL.ProjectNo = strProjectNo
                myClsBG0310BL.BudgetType = BGConstant.P_BUDGET_TYPE_EXPENSE

                If myClsBG0310BL.GetRevNo() = True Then
                    Me.cboRevNo.DisplayMember = "REV_NO"
                    Me.cboRevNo.ValueMember = "REV_NO"
                    Me.cboRevNo.DataSource = myClsBG0310BL.RevNoList
                Else
                    Me.cboRevNo.DataSource = Nothing
                End If
            Else
                Me.cboRevNo.DataSource = Nothing
            End If

        End If
    End Sub

    Private Function fncCheckRevNo() As Boolean
        Dim blnChkResult As Boolean = True

        If p_intUserLevelId = enumUserLevel.SystemAdministrator Then

            If Me.cboRevNo.DataSource Is Nothing OrElse _
                Me.cboRevNo.SelectedIndex < 0 Then
                blnChkResult = False
            End If

        End If

        Return blnChkResult
    End Function


#End Region

#Region "Control Event"

    Private Sub frmBG0450_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        If Not excelApp Is Nothing Then
            excelApp.Quit()
            excelApp = Nothing
        End If
    End Sub
    Private Sub frmBG0450_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        m_blnFormLoading = True
        LoadBudgetYear()
        LoadPeriodType()

        If p_intUserLevelId = enumUserLevel.SystemAdministrator Then
            Me.lblRevNo.Visible = True
            Me.cboRevNo.Visible = True
            LoadRevNo()

            'Me.lblPrevRevNo.Visible = True
            'Me.cboPrevRevno.Visible = True
            'LoadPrevRevNo()
        End If

        'If CType(cboPeriodType.SelectedValue, Integer) = CType(enumPeriodType.MTPBudget, Integer) Then
        '    Me.gbPrevYear.Text = "Previous Year"

        'ElseIf CType(cboPeriodType.SelectedValue, Integer) = CType(enumPeriodType.OriginalBudget, Integer) Then
        '    Me.gbPrevYear.Text = "MTP"

        'End If

        'If CType(cboPeriodType.SelectedValue, Integer) = CType(enumPeriodType.MTPBudget, Integer) OrElse _
        '        CType(cboPeriodType.SelectedValue, Integer) = CType(enumPeriodType.OriginalBudget, Integer) Then
        '    EnablePrev()
        'Else
        '    DisablePrev()
        'End If
        m_blnFormLoading = False
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

    Private Sub cboPeriodType_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPeriodType.SelectedValueChanged
        If m_blnFormLoading = False Then
            If CType(cboPeriodType.SelectedValue, Integer) = CType(enumPeriodType.ReviseBudget, Integer) Then
                chkShowMTP.Enabled = True
            Else
                chkShowMTP.Enabled = False
            End If

            'If CType(cboPeriodType.SelectedValue, Integer) = CType(enumPeriodType.MTPBudget, Integer) Then
            '    Me.gbPrevYear.Text = "Previous Year"

            'ElseIf CType(cboPeriodType.SelectedValue, Integer) = CType(enumPeriodType.OriginalBudget, Integer) Then
            '    Me.gbPrevYear.Text = "MTP"

            'End If

            'If CType(cboPeriodType.SelectedValue, Integer) = CType(enumPeriodType.MTPBudget, Integer) OrElse _
            '    CType(cboPeriodType.SelectedValue, Integer) = CType(enumPeriodType.OriginalBudget, Integer) Then
            '    EnablePrev()
            'Else
            '    DisablePrev()
            'End If

            LoadRevNo()

        End If
    End Sub

    Private Sub cmdExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExcel.Click
        Try
            If Me.cboPeriodType.SelectedIndex = -1 Then
                MessageBox.Show("Please select a Period Type!", "Summary By Applicant Report", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Me.cboPeriodType.Focus()
                Me.cboPeriodType.SelectAll()
                Return
            End If

            If fncCheckRevNo() = False Then

                MessageBox.Show("No budget data found, please try it again.", "Summary By Applicant Report", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            'If fncCheckPrevRevNo() = False Then
            '    MessageBox.Show("No previous budget data found, please try it again.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            '    Exit Sub
            'End If

            Cursor = Cursors.WaitCursor

            myClsBG0450BL.BudgetYear = CStr(Me.numYear.Value)
            myClsBG0450BL.PeriodType = CStr(Me.cboPeriodType.SelectedValue)
            myClsBG0450BL.ProjectNo = Me.numProjectNo.Value.ToString
            myClsBG0450BL.UserLevelId = p_intUserLevelId
            If Me.cboRevNo.DataSource IsNot Nothing Then
                myClsBG0450BL.RevNo = Me.cboRevNo.SelectedValue.ToString
            End If

            myClsBG0450BL.PrevProjectNo = Me.numPrevProjectNo.Value.ToString
            If Me.cboPrevRevno.DataSource IsNot Nothing AndAlso _
                Me.cboPrevRevno.SelectedValue IsNot Nothing Then
                myClsBG0450BL.PrevRevNo = Me.cboPrevRevno.SelectedValue.ToString
            End If

            'myClsBG0450BL.MTPBudget = Me.chkShowMTP.Checked

            If myClsBG0450BL.getApplicantData() Then

                Dim ds As DataSet = myClsBG0450BL.ApplicantData

                If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then

                    myClsBG0450BL.GetBudgetStatus()

                    myClsBG0450BL.GetAuthImage()
                    ds.Tables.Add(myClsBG0450BL.AuthImage)

                    Dim strYear As String = Me.numYear.Value.ToString
                    '//Create output columns
                    Dim dtColumns As DataTable = CreateTableTemplate()
                    Dim dsGroups As DataSet = Nothing

                    Select Case CType(Me.cboPeriodType.SelectedValue, enumPeriodType)
                        Case enumPeriodType.OriginalBudget
                            'strReportName = "RPT005-1.rpt"
                            InsertOriginalColumnData(dtColumns, strYear)
                            dsGroups = SetupOriginalGroupbyData(ds, "COST", "COST", 9)
                            GeneratOriginalExcel(dsGroups, dtColumns)

                        Case enumPeriodType.EstimateBudget
                            'strReportName = "RPT005-2.rpt"
                            InsertEstimateColumnData(dtColumns, strYear)
                            dsGroups = SetupEstimateGroupbyData(ds, "COST", "COST", 9)
                            GeneratEstimateExcel(dsGroups, dtColumns)

                        Case enumPeriodType.ReviseBudget
                            If Not chkShowMTP.Checked Then
                                'strReportName = "RPT005-3.rpt"
                                InsertReviseColumnData(dtColumns, strYear)
                                dsGroups = SetupReviseGroupbyData(ds, "COST", "COST", 9, False)
                                GeneratReviseExcel(dsGroups, dtColumns, False)
                            Else
                                'strReportName = "RPT005-4.rpt"
                                InsertReviseMTPColumnData(dtColumns, strYear)
                                dsGroups = SetupReviseGroupbyData(ds, "COST", "COST", 9, True)
                                GeneratReviseExcel(dsGroups, dtColumns, True)
                            End If

                        Case enumPeriodType.MTPBudget
                            InsertMTPColumnData(dtColumns, strYear)
                            dsGroups = SetupMTPGroupbyData(ds, "COST", "COST", 8)
                            GeneratMTPExcel(dsGroups, dtColumns)

                    End Select

                Else
                    MessageBox.Show("No budget data found, please try it again.", "Summary By Applicant Report", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Summary By Applicant Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Cursor = Cursors.Default
    End Sub

#End Region

    Private Sub numYear_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numYear.ValueChanged
        LoadRevNo()
        'LoadPrevRevNo()
    End Sub

    Private Sub numProjectNo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numProjectNo.ValueChanged
        LoadRevNo()
    End Sub

    'Private Sub numPrevProjectNo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numPrevProjectNo.ValueChanged
    '    LoadPrevRevNo()
    'End Sub

End Class