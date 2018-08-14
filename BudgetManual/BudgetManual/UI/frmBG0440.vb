Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports CrystalDecisions.CrystalReports.Engine
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class frmBG0440

#Region "Variable"
    Private myClsBG0440BL As New clsBG0440BL
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
            MessageBox.Show(ex.Message, "Summary By Account No Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
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
    '        MessageBox.Show(ex.Message, "Summary By Account No Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub

    Private Sub Print(ByVal blnShowPrintPreview As Boolean)
        Dim strReportName As String = String.Empty
        Try
            If Me.cboPeriodType.SelectedIndex = -1 Then
                MessageBox.Show("Please select a Period Type!", "Summary By Account No Report", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Me.cboPeriodType.Focus()
                Me.cboPeriodType.SelectAll()
                Return
            End If

            If fncCheckRevNo() = False Then

                MessageBox.Show("No budget data found, please try it again.", "Summary By Account No Report", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            'If fncCheckPrevRevNo() = False Then
            '    MessageBox.Show("No previous budget data found, please try it again.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            '    Exit Sub
            'End If

            Cursor = Cursors.WaitCursor

            myClsBG0440BL.BudgetYear = CStr(Me.numYear.Value)
            myClsBG0440BL.PeriodType = CStr(Me.cboPeriodType.SelectedValue)
            myClsBG0440BL.ProjectNo = Me.numProjectNo.Value.ToString
            'myClsBG0440BL.MTPBudget = Me.chkShowMTP.Checked
            myClsBG0440BL.UserLevelId = p_intUserLevelId
            If Me.cboRevNo.DataSource IsNot Nothing Then
                myClsBG0440BL.RevNo = Me.cboRevNo.SelectedValue.ToString
            End If

            myClsBG0440BL.PrevProjectNo = Me.numPrevProjectNo.Value.ToString
            If Me.cboPrevRevno.DataSource IsNot Nothing AndAlso _
                Me.cboPrevRevno.SelectedValue IsNot Nothing Then
                myClsBG0440BL.PrevRevNo = Me.cboPrevRevno.SelectedValue.ToString
            End If

            If myClsBG0440BL.getAccountData() Then

                Dim ds As DataSet = myClsBG0440BL.AccountCodeData

                If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then

                    myClsBG0440BL.GetBudgetStatus()

                    myClsBG0440BL.GetAuthImage()
                    ds.Tables.Add(myClsBG0440BL.AuthImage)

                    'Dim sngWorkingBG1 As Single = 0
                    'Dim sngWorkingBG2 As Single = 0
                    'myClsBG0440BL.getBudgetAdjust()
                    'If myClsBG0440BL.BudgetAdjustTable IsNot Nothing AndAlso myClsBG0440BL.BudgetAdjustTable.Rows.Count > 0 Then
                    '    sngWorkingBG1 = CSng(myClsBG0440BL.BudgetAdjustTable.Rows(0)![WORKING_BG1])
                    '    sngWorkingBG2 = CSng(myClsBG0440BL.BudgetAdjustTable.Rows(0)![WORKING_BG2])
                    'Else
                    '    sngWorkingBG1 = 0
                    '    sngWorkingBG2 = 0
                    'End If

                    Dim blnMTPBudget As Boolean = False
                    Select Case CType(Me.cboPeriodType.SelectedValue, enumPeriodType)
                        Case enumPeriodType.OriginalBudget
                            strReportName = "RPT004-1.rpt"
                        Case enumPeriodType.EstimateBudget
                            strReportName = "RPT004-2.rpt"
                        Case enumPeriodType.ReviseBudget
                            If Not chkShowMTP.Checked Then
                                strReportName = "RPT004-3.rpt"
                            Else
                                blnMTPBudget = True
                                strReportName = "RPT004-4.rpt"
                            End If
                        Case enumPeriodType.MTPBudget
                            strReportName = "RPT004-5.rpt"
                    End Select

                    Dim strExpr As String = "BUDGET_TYPE = 'A'"
                    Dim strSort As String = "ACCOUNT_NO ASC"
                    Dim dtTmp As DataTable = ds.Tables(0).Clone
                    If blnMTPBudget = True Then
                        Dim arrTmp As DataRow() = ds.Tables(0).Select(strExpr, strSort)
                        For intTmp As Integer = 0 To arrTmp.Length - 1
                            Dim drow(dtTmp.Columns.Count - 1) As Object
                            arrTmp(intTmp).ItemArray.CopyTo(drow, 0)
                            dtTmp.Rows.Add(drow)
                        Next
                    End If

                    If blnShowPrintPreview Then

                        If clsBG0400 IsNot Nothing Then
                            clsBG0400.Close()
                            clsBG0400.Dispose()
                        End If
                        clsBG0400 = New frmBG0400()
                        clsBG0400.MdiParent = p_frmBG0010
                        clsBG0400.ReportName = strReportName
                        clsBG0400.BudgetYear = Me.numYear.Value.ToString()
                        clsBG0400.ReportType = "SummaryByAccountNoReport"
                        clsBG0400.BudgetStatus = myClsBG0440BL.BudgetStatus
                        clsBG0400.ProjectNo = Me.numProjectNo.Value.ToString
                        'clsBG0400.WorkingBG1 = sngWorkingBG1
                        'clsBG0400.WorkingBG2 = sngWorkingBG2
                        If blnMTPBudget = True Then

                            clsBG0400.MTPBudget = True

                            If dtTmp.Rows.Count > 0 Then
                                clsBG0400.MTP_SUM1 = CDec(Nz(dtTmp.Rows(0)![INVESTMENT_RRT1], 0.0))
                                clsBG0400.MTP_SUM2 = CDec(Nz(dtTmp.Rows(0)![INVESTMENT_RRT2], 0.0))
                                clsBG0400.MTP_SUM3 = CDec(Nz(dtTmp.Rows(0)![INVESTMENT_RRT3], 0.0))
                                clsBG0400.MTP_SUM4 = CDec(Nz(dtTmp.Rows(0)![INVESTMENT_RRT4], 0.0))
                                clsBG0400.MTP_SUM5 = CDec(Nz(dtTmp.Rows(0)![INVESTMENT_RRT5], 0.0))
                            End If

                        End If


                        clsBG0400.DS = ds

                        clsBG0400.Show()
                        If clsBG0400.WindowState = FormWindowState.Minimized Then
                            clsBG0400.WindowState = FormWindowState.Normal
                        End If
                        clsBG0400.BringToFront()
                    Else
                        ' Allow the user to choose the page range he or she would
                        ' like to print.
                        PrintDialog1.AllowSomePages = True

                        ' Show the help button.
                        PrintDialog1.ShowHelp = True

                        Dim result As DialogResult = PrintDialog1.ShowDialog()

                        ' If the result is OK then print the document.
                        If (result = DialogResult.OK) Then

                            Dim rpt1 As ReportDocument = Nothing

                            rpt1 = New ReportDocument()
                            Dim reportPath As String = p_strAppPath & "\Reports\" & strReportName
                            rpt1.Load(reportPath)

                            myClsBG0440BL.GetBudgetStatus()

                            If myClsBG0440BL.BudgetStatus >= 5 Then
                                rpt1.ReportDefinition.ReportObjects("picAuth10").ObjectFormat.EnableSuppress = False
                            Else
                                rpt1.ReportDefinition.ReportObjects("picAuth10").ObjectFormat.EnableSuppress = True
                            End If

                            If myClsBG0440BL.BudgetStatus >= 6 Then
                                rpt1.ReportDefinition.ReportObjects("picAuth3").ObjectFormat.EnableSuppress = False
                            Else
                                rpt1.ReportDefinition.ReportObjects("picAuth3").ObjectFormat.EnableSuppress = True
                            End If

                            rpt1.SetDataSource(ds)

                            rpt1.SetParameterValue("BUDGET_YEAR", Me.numYear.Value.ToString())
                            rpt1.SetParameterValue("HALF_BUDGET_YEAR", Me.numYear.Value.ToString().Substring(2, 2))
                            rpt1.SetParameterValue("FC_COST", enumCost.FC)
                            rpt1.SetParameterValue("ADMIN_COST", enumCost.ADMIN)
                            rpt1.SetParameterValue("PROJECT_NO", Me.numProjectNo.Value.ToString)
                            'rpt1.SetParameterValue("WORKING_BG1", sngWorkingBG1)
                            'rpt1.SetParameterValue("WORKING_BG2", sngWorkingBG2)
                            If blnMTPBudget = True Then
                                rpt1.SetParameterValue("MTP_SUM1", CDec(Nz(dtTmp.Rows(0)![INVESTMENT_RRT1], 0.0)))
                                rpt1.SetParameterValue("MTP_SUM2", CDec(Nz(dtTmp.Rows(0)![INVESTMENT_RRT2], 0.0)))
                                rpt1.SetParameterValue("MTP_SUM3", CDec(Nz(dtTmp.Rows(0)![INVESTMENT_RRT3], 0.0)))
                                rpt1.SetParameterValue("MTP_SUM4", CDec(Nz(dtTmp.Rows(0)![INVESTMENT_RRT4], 0.0)))
                                rpt1.SetParameterValue("MTP_SUM5", CDec(Nz(dtTmp.Rows(0)![INVESTMENT_RRT5], 0.0)))
                            End If

                            rpt1.PrintOptions.PrinterName = PrintDialog1.PrinterSettings.PrinterName
                            rpt1.PrintToPrinter(PrintDialog1.PrinterSettings.Copies, _
                                                PrintDialog1.PrinterSettings.Collate, _
                                                PrintDialog1.PrinterSettings.FromPage, _
                                                PrintDialog1.PrinterSettings.ToPage)

                        End If
                    End If
                Else
                    MessageBox.Show("No data is available for viewing reports!", "Summary By Account No Report", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Summary By Account No Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
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


        '     SELECT MAX_REV.BUDGET_YEAR, MAX_REV.PERIOD_TYPE, 
        'MAX_REV.ACCOUNT_NO, 
        row = dtColumns.NewRow()
        row("Column_Name") = "ACCOUNT_NO"
        row("Column_Title") = "Item"
        dtColumns.Rows.Add(row)

        'MAX_REV.ACCOUNT_NAME, MAX_REV.BUDGET_TYPE, MAX_REV.EXPENSE_TYPE,
        row = dtColumns.NewRow()
        row("Column_Name") = "ACCOUNT_NAME"
        row("Column_Title") = ""
        dtColumns.Rows.Add(row)

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

        'SUM(ISNULL(MASTER_DATA.M1, 0)) AS M1, SUM(ISNULL(MASTER_DATA.M2, 0)) AS M2, SUM(ISNULL(MASTER_DATA.M3, 0)) AS M3, SUM(ISNULL(MASTER_DATA.M4, 0)) AS M4, SUM(ISNULL(MASTER_DATA.M5, 0)) AS M5, SUM(ISNULL(MASTER_DATA.M6, 0)) AS M6,
        row = dtColumns.NewRow()
        row("Column_Name") = "M1"
        row("Column_Title") = "Jan'" & strHalfYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "M2"
        row("Column_Title") = "Feb'" & strHalfYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "M3"
        row("Column_Title") = "Mar'" & strHalfYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "M4"
        row("Column_Title") = "Apr'" & strHalfYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "M5"
        row("Column_Title") = "May'" & strHalfYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "M6"
        row("Column_Title") = "Jun'" & strHalfYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "TOTAL_1ST_HALF"
        row("Column_Title") = "Total 1st Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "M7"
        row("Column_Title") = "Jul'" & strHalfYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "M8"
        row("Column_Title") = "Aug'" & strHalfYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "M9"
        row("Column_Title") = "Sept'" & strHalfYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "M10"
        row("Column_Title") = "Oct'" & strHalfYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "M11"
        row("Column_Title") = "Nov'" & strHalfYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "M12"
        row("Column_Title") = "Dec'" & strHalfYear
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "TOTAL_2ND_HALF"
        row("Column_Title") = "Total 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(row)

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

        'SUM(ISNULL(MASTER_DATA.M7, 0)) AS M7, SUM(ISNULL(MASTER_DATA.M8, 0)) AS M8, SUM(ISNULL(MASTER_DATA.M9, 0)) AS M9, SUM(ISNULL(MASTER_DATA.M10, 0)) AS M10, SUM(ISNULL(MASTER_DATA.M11, 0)) AS M11, SUM(ISNULL(MASTER_DATA.M12, 0)) AS M12,
        '0 AS INVESTMENT_ACTUAL_1ST_HALF,
        '0 AS INVESTMENT_REVISE_2ND_HALF,
        '0 AS INVESTMENT_M1,
        '0 AS INVESTMENT_M2,
        '0 AS INVESTMENT_M3,
        '0 AS INVESTMENT_M4,
        '0 AS INVESTMENT_M5,
        '0 AS INVESTMENT_M6,      
        '0 AS INVESTMENT_2ND_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @Admin THEN ISNULL(ACTUAL_DATA.H1,0) ELSE 0 END ) AS ADMIN_ACTUAL_1ST_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @FC THEN ISNULL(ACTUAL_DATA.H1,0) ELSE 0 END ) AS FC_ACTUAL_1ST_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @Admin THEN ISNULL(REVISE_BUDGET.H2,0) ELSE 0 END ) AS ADMIN_REVISE_2ND_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @FC THEN ISNULL(REVISE_BUDGET.H2,0) ELSE 0 END ) AS FC_REVISE_2ND_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @Admin THEN ISNULL(MASTER_DATA.M1,0) ELSE 0 END ) AS ADMIN_M1,
        'SUM(CASE WHEN MAX_REV.COST = @FC THEN ISNULL(MASTER_DATA.M1,0) ELSE 0 END ) AS FC_M1,
        'SUM(CASE WHEN MAX_REV.COST = @Admin THEN ISNULL(MASTER_DATA.M2,0) ELSE 0 END ) AS ADMIN_M2,
        'SUM(CASE WHEN MAX_REV.COST = @FC THEN ISNULL(MASTER_DATA.M2,0) ELSE 0 END ) AS FC_M2,
        'SUM(CASE WHEN MAX_REV.COST = @Admin THEN ISNULL(MASTER_DATA.M3,0) ELSE 0 END ) AS ADMIN_M3,
        'SUM(CASE WHEN MAX_REV.COST = @FC THEN ISNULL(MASTER_DATA.M3,0) ELSE 0 END ) AS FC_M3,
        'SUM(CASE WHEN MAX_REV.COST = @Admin THEN ISNULL(MASTER_DATA.M4,0) ELSE 0 END ) AS ADMIN_M4,
        'SUM(CASE WHEN MAX_REV.COST = @FC THEN ISNULL(MASTER_DATA.M4,0) ELSE 0 END ) AS FC_M4,
        'SUM(CASE WHEN MAX_REV.COST = @Admin THEN ISNULL(MASTER_DATA.M5,0) ELSE 0 END )AS ADMIN_M5,
        'SUM(CASE WHEN MAX_REV.COST = @FC THEN ISNULL(MASTER_DATA.M5,0) ELSE 0 END ) AS FC_M5,
        'SUM(CASE WHEN MAX_REV.COST = @Admin THEN ISNULL(MASTER_DATA.M6,0) ELSE 0 END )AS ADMIN_M6,
        'SUM(CASE WHEN MAX_REV.COST = @FC THEN ISNULL(MASTER_DATA.M6,0) ELSE 0 END ) AS FC_M6,
        'SUM(CASE WHEN MAX_REV.COST = @Admin THEN ISNULL(MASTER_DATA.M7,0) ELSE 0 END ) AS ADMIN_M7,
        'SUM(CASE WHEN MAX_REV.COST = @FC THEN ISNULL(MASTER_DATA.M7,0) ELSE 0 END )AS FC_M7,
        'SUM(CASE WHEN MAX_REV.COST = @Admin THEN ISNULL(MASTER_DATA.M8,0) ELSE 0 END )AS ADMIN_M8,
        'SUM(CASE WHEN MAX_REV.COST = @FC THEN ISNULL(MASTER_DATA.M8,0) ELSE 0 END )AS FC_M8,
        'SUM(CASE WHEN MAX_REV.COST = @Admin THEN ISNULL(MASTER_DATA.M9,0) ELSE 0 END )AS ADMIN_M9,
        'SUM(CASE WHEN MAX_REV.COST = @FC THEN ISNULL(MASTER_DATA.M9,0) ELSE 0 END ) AS FC_M9,
        'SUM(CASE WHEN MAX_REV.COST = @Admin THEN ISNULL(MASTER_DATA.M10,0) ELSE 0 END ) AS ADMIN_M10,
        'SUM(CASE WHEN MAX_REV.COST = @FC THEN ISNULL(MASTER_DATA.M10,0) ELSE 0 END ) AS FC_M10,
        'SUM(CASE WHEN MAX_REV.COST = @Admin THEN ISNULL(MASTER_DATA.M11,0) ELSE 0 END ) AS ADMIN_M11,
        'SUM(CASE WHEN MAX_REV.COST = @FC THEN ISNULL(MASTER_DATA.M11,0) ELSE 0 END ) AS FC_M11,
        'SUM(CASE WHEN MAX_REV.COST = @Admin THEN ISNULL(MASTER_DATA.M12,0) ELSE 0 END ) AS ADMIN_M12,
        'SUM(CASE WHEN MAX_REV.COST = @FC THEN ISNULL(MASTER_DATA.M12,0) ELSE 0 END ) AS FC_M12,
        'SUM(CASE WHEN MAX_REV.COST = @Admin THEN ISNULL(MASTER_DATA.M1,0) + ISNULL(MASTER_DATA.M2,0) + ISNULL(MASTER_DATA.M3,0) + ISNULL(MASTER_DATA.M4,0) + ISNULL(MASTER_DATA.M5,0) + ISNULL(MASTER_DATA.M6,0) ELSE 0 END ) AS ADMIN_1ST_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @FC THEN ISNULL(MASTER_DATA.M1,0) + ISNULL(MASTER_DATA.M2,0) + ISNULL(MASTER_DATA.M3,0) + ISNULL(MASTER_DATA.M4,0) + ISNULL(MASTER_DATA.M5,0) + ISNULL(MASTER_DATA.M6,0) ELSE 0 END ) AS FC_1ST_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @Admin THEN ISNULL(MASTER_DATA.M7,0) + ISNULL(MASTER_DATA.M8,0) + ISNULL(MASTER_DATA.M9,0) + ISNULL(MASTER_DATA.M10,0) + ISNULL(MASTER_DATA.M11,0) + ISNULL(MASTER_DATA.M12,0) ELSE 0 END ) AS ADMIN_2ND_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @FC THEN ISNULL(MASTER_DATA.M7,0) + ISNULL(MASTER_DATA.M8,0) + ISNULL(MASTER_DATA.M9,0) + ISNULL(MASTER_DATA.M10,0) + ISNULL(MASTER_DATA.M11,0) + ISNULL(MASTER_DATA.M12,0) ELSE 0 END ) AS FC_2ND_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @Admin THEN ISNULL(MASTER_DATA.M1,0) + ISNULL(MASTER_DATA.M2,0) + ISNULL(MASTER_DATA.M3,0) + ISNULL(MASTER_DATA.M4,0) + ISNULL(MASTER_DATA.M5,0) + ISNULL(MASTER_DATA.M6,0) + ISNULL(MASTER_DATA.M7,0) + ISNULL(MASTER_DATA.M8,0) + ISNULL(MASTER_DATA.M9,0) + ISNULL(MASTER_DATA.M10,0) + ISNULL(MASTER_DATA.M11,0) + ISNULL(MASTER_DATA.M12,0) ELSE 0 END ) AS ADMIN_TOTAL_YEAR,
        'SUM(CASE WHEN MAX_REV.COST = @FC THEN ISNULL(MASTER_DATA.M1,0) + ISNULL(MASTER_DATA.M2,0) + ISNULL(MASTER_DATA.M3,0) + ISNULL(MASTER_DATA.M4,0) + ISNULL(MASTER_DATA.M5,0) + ISNULL(MASTER_DATA.M6,0) + ISNULL(MASTER_DATA.M7,0) + ISNULL(MASTER_DATA.M8,0) + ISNULL(MASTER_DATA.M9,0) + ISNULL(MASTER_DATA.M10,0) + ISNULL(MASTER_DATA.M11,0) + ISNULL(MASTER_DATA.M12,0) ELSE 0 END ) AS FC_TOTAL_YEAR,
        '(SELECT ISNULL(WKH1,0) FROM BG_T_BUDGET_ADJUST WHERE BUDGET_YEAR=@BudgetYear AND PERIOD_TYPE=@PeriodType AND REV_NO= (SELECT MAX(REV_NO) AS REV_NO FROM BG_T_BUDGET_DATA WHERE BUDGET_YEAR = @BudgetYear AND PERIOD_TYPE = @PeriodType)) AS WB_ACTUAL_1ST_HALF,
        '(SELECT ISNULL(WKH2,0) FROM BG_T_BUDGET_ADJUST WHERE BUDGET_YEAR=@BudgetYear AND PERIOD_TYPE=@PeriodType AND REV_NO= (SELECT MAX(REV_NO) AS REV_NO FROM BG_T_BUDGET_DATA WHERE BUDGET_YEAR = @BudgetYear AND PERIOD_TYPE = @PeriodType)) AS WB_REVISE_2ND_HALF,
        'SUM(ISNULL(MIN_REV.M1,0) -   ISNULL(MASTER_DATA.M1,0)) AS WB_M1,
        'SUM(ISNULL(MIN_REV.M2,0) -   ISNULL(MASTER_DATA.M2,0)) AS WB_M2,
        'SUM(ISNULL(MIN_REV.M3,0) -   ISNULL(MASTER_DATA.M3,0)) AS WB_M3,
        'SUM(ISNULL(MIN_REV.M4,0) -   ISNULL(MASTER_DATA.M4,0)) AS WB_M4,
        'SUM(ISNULL(MIN_REV.M5,0) -   ISNULL(MASTER_DATA.M5,0)) AS WB_M5,
        'SUM(ISNULL(MIN_REV.M6,0) -   ISNULL(MASTER_DATA.M6,0)) AS WB_M6,
        'SUM(ISNULL(MIN_REV.M7,0) -   ISNULL(MASTER_DATA.M7,0)) AS WB_M7,
        'SUM(ISNULL(MIN_REV.M8,0) -   ISNULL(MASTER_DATA.M8,0)) AS WB_M8,
        'SUM(ISNULL(MIN_REV.M9,0) -   ISNULL(MASTER_DATA.M9,0)) AS WB_M9,
        'SUM(ISNULL(MIN_REV.M10,0) -   ISNULL(MASTER_DATA.M10,0)) AS WB_M10,
        'SUM(ISNULL(MIN_REV.M11,0) -   ISNULL(MASTER_DATA.M11,0)) AS WB_M11,
        'SUM(ISNULL(MIN_REV.M12,0) -   ISNULL(MASTER_DATA.M12,0)) AS WB_M12,
        '0 AS WB_TOTAL_1ST_HALF,
        '0 AS WB_TOTAL_2ND_HALF

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
        Dim drManufacturingCost As DataRow = dtResult.NewRow
        Dim drAdministrationCost As DataRow = dtResult.NewRow
        Dim drTotalExpense As DataRow = dtResult.NewRow
        Dim drWorkingBudget As DataRow = dtResult.NewRow
        Dim drOutflowTotal As DataRow = dtResult.NewRow

        '//Add column 
        Dim col As DataColumn = New DataColumn()
        col.ColumnName = "TOTAL_1ST_HALF"
        col.DataType = Type.GetType("System.Decimal")
        col.DefaultValue = 0.0
        dtResult.Columns.Add(col)

        col = New DataColumn()
        col.ColumnName = "TOTAL_2ND_HALF"
        col.DataType = Type.GetType("System.Decimal")
        col.DefaultValue = 0.0
        dtResult.Columns.Add(col)

        col = New DataColumn()
        col.ColumnName = "TOTAL_YEAR"
        col.DataType = Type.GetType("System.Decimal")
        col.DefaultValue = 0.0
        dtResult.Columns.Add(col)

        col = New DataColumn()
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

        '//Calculate Investments
        CalculateOriginalInvestments(dsData, intDataColumnIndex, drInvestments)

        '//Calculate Manufacturing cost
        CalculateOriginalManufacturingCost(dsData, intDataColumnIndex, drManufacturingCost)

        '//Calculate Administration cost
        CalculateOriginalAdministrationCost(dsData, intDataColumnIndex, drAdministrationCost)

        '//Calculate Working Budget
        CalculateOriginalWorkingBudget(dsData, intDataColumnIndex, drWorkingBudget)


        Dim intGroupTotalIndex As Integer = 0
        For i As Integer = 0 To intGroupCount - 1

            strScript = strGroupColumnName + " = " + dtGroups.Rows(i)(0).ToString

            Dim arrRows As DataRow() = dtTmp.Select(strScript)

            For j As Integer = 0 To arrRows.Length - 1
                Dim drow(dtResult.Columns.Count - 1) As Object
                arrRows(j).ItemArray.CopyTo(drow, 0)
                dtResult.Rows.Add(drow)
            Next

            '//Calculate Horizontal Column
            For m As Integer = 0 To dtResult.Rows.Count - 1
                For n As Integer = intDataColumnIndex To dtResult.Columns.Count - 1
                    Dim strColumnName As String = dtResult.Columns(n).ColumnName
                    If strColumnName = "TOTAL_1ST_HALF" Then
                        dtResult.Rows(m)![TOTAL_1ST_HALF] = Convert.ToDecimal(Nz(dtResult.Rows(m)![M1], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![M2], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![M3], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![M4], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![M5], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![M6], 0.0))
                    ElseIf strColumnName = "TOTAL_2ND_HALF" Then
                        dtResult.Rows(m)![TOTAL_2ND_HALF] = Convert.ToDecimal(Nz(dtResult.Rows(m)![M7], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![M8], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![M9], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![M10], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![M11], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![M12], 0.0))
                    ElseIf strColumnName = "TOTAL_YEAR" Then
                        dtResult.Rows(m)![TOTAL_YEAR] = Convert.ToDecimal(Nz(dtResult.Rows(m)![TOTAL_1ST_HALF], 0.0)) + Convert.ToDecimal(Nz(dtResult.Rows(m)![TOTAL_2ND_HALF], 0.0))
                    ElseIf strColumnName = "TOTAL_LAST_YEAR" Then
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
            Dim drTotal As DataRow = dtResult.NewRow
            For k As Integer = intDataColumnIndex To dtResult.Columns.Count - 1

                Dim strColumnName As String = dtResult.Columns(k).ColumnName
                strExpression = "Sum(" + strColumnName + ")"
                strFilter = strScript
                returnValue = dtResult.Compute(strExpression, strFilter)
                drTotal(dtResult.Columns(k).ColumnName) = returnValue

            Next
            '//Set Group header
            drTotal("ACCOUNT_NO") = GetGroupExpensesTitle(dtGroups.Rows(i)(0).ToString)

            '//Add total cost
            dtResult.Rows.InsertAt(drTotal, intGroupTotalIndex)

            '//Add one empty row
            drEmpty = dtResult.NewRow
            dtResult.Rows.Add(drEmpty)

            'dtResult.TableName = arrRows(0)(strGroupColumnName).ToString & " " & arrRows(0)(strGroupColumnTitle).ToString
            dtResult.TableName = "Original Budget"

            'intGroupTotalIndex = intGroupTotalIndex + dtResult.Rows.Count
            intGroupTotalIndex = intGroupTotalIndex + CInt(arrRows.Length) + 2

        Next
        '//Set Data to Account No.
        SetAccountNoText(drInvestments, drManufacturingCost, drAdministrationCost, drTotalExpense, drWorkingBudget, drOutflowTotal)

        '//Calculate Total Expense
        CalculateTotalExpense(dtResult, intDataColumnIndex, drManufacturingCost, drAdministrationCost, drTotalExpense)

        '//Calculate Outflow Total 
        CalculateOutflowTotal(dtResult, intDataColumnIndex, drTotalExpense, drWorkingBudget, drInvestments, drOutflowTotal)

        '//Add Investments
        dtResult.Rows.InsertAt(drInvestments, 0)

        '//Add one empty row
        drEmpty = dtResult.NewRow
        dtResult.Rows.InsertAt(drEmpty, 1)

        '//Add Manufacturing cost
        dtResult.Rows.Add(drManufacturingCost)

        '//Add Administration cost
        dtResult.Rows.Add(drAdministrationCost)

        '//Add Total Expense
        dtResult.Rows.Add(drTotalExpense)

        '//Add one empty row
        drEmpty = dtResult.NewRow
        dtResult.Rows.Add(drEmpty)

        '//Add Working Budget
        dtResult.Rows.Add(drWorkingBudget)

        '//Add Outflow Total (Investment;Expenses)
        dtResult.Rows.Add(drOutflowTotal)

        '//Return data table
        dsResult.Tables.Add(dtResult)

        Return dsResult
    End Function

    Private Function CalculateOriginalInvestments(ByVal dsData As DataSet, _
                                         ByVal intDataColumnIndex As Integer, _
                                         ByRef drInvestments As DataRow) As Boolean
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object

        For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName

            strExpression = "Sum(" + strColumnName + ")"
            strFilter = "BUDGET_TYPE = 'A'"
            returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
            drInvestments(dsData.Tables(0).Columns(k).ColumnName) = returnValue

            Select Case dsData.Tables(0).Columns(k).ColumnName
                Case "INVESTMENT_ACTUAL_1ST_HALF"
                    drInvestments("ACTUAL_1ST_HALF") = returnValue
                Case "INVESTMENT_REVISE_2ND_HALF"
                    drInvestments("REVISE_2ND_HALF") = returnValue

                Case "INVESTMENT_M1"
                    drInvestments("M1") = returnValue
                Case "INVESTMENT_M2"
                    drInvestments("M2") = returnValue
                Case "INVESTMENT_M3"
                    drInvestments("M3") = returnValue
                Case "INVESTMENT_M4"
                    drInvestments("M4") = returnValue
                Case "INVESTMENT_M5"
                    drInvestments("M5") = returnValue
                Case "INVESTMENT_M6"
                    drInvestments("M6") = returnValue

                Case "INVESTMENT_2ND_HALF"
                    drInvestments("TOTAL_2ND_HALF") = returnValue

                    '//Calculate
                    drInvestments("TOTAL_1ST_HALF") = Convert.ToDecimal(Nz(drInvestments("INVESTMENT_M1"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("INVESTMENT_M2"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("INVESTMENT_M3"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("INVESTMENT_M4"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("INVESTMENT_M5"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("INVESTMENT_M6"), 0.0))

                    '{@INVEST_1ST_HALF} + {@INVEST_2ND_HALF}
                    drInvestments("TOTAL_YEAR") = Convert.ToDecimal(Nz(drInvestments("TOTAL_1ST_HALF"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("TOTAL_2ND_HALF"), 0.0))

                    '{@INVEST_ACTUAL_1ST_HALF} + {@INVEST_REVISE_2ND_HALF}
                    drInvestments("TOTAL_LAST_YEAR") = Convert.ToDecimal(Nz(drInvestments("ACTUAL_1ST_HALF"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("REVISE_2ND_HALF"), 0.0))

                    '{@INVEST_TOTAL_YEAR} - {@INVEST_TOTAL_LAST_YEAR}
                    drInvestments("DIFFERENCE") = Convert.ToDecimal(Nz(drInvestments("TOTAL_YEAR"), 0.0)) - Convert.ToDecimal(Nz(drInvestments("TOTAL_LAST_YEAR"), 0.0))

                Case "MTP_RRT1"
                    drInvestments("DIFF_MTP") = Convert.ToDecimal(Nz(drInvestments("TOTAL_YEAR"), 0.0)) - Convert.ToDecimal(Nz(drInvestments("MTP_RRT1"), 0.0))


                Case "INVESTMENT_M7"
                    drInvestments("M7") = returnValue
                Case "INVESTMENT_M8"
                    drInvestments("M8") = returnValue
                Case "INVESTMENT_M9"
                    drInvestments("M9") = returnValue
                Case "INVESTMENT_M10"
                    drInvestments("M10") = returnValue
                Case "INVESTMENT_M11"
                    drInvestments("M11") = returnValue
                Case "INVESTMENT_M12"
                    drInvestments("M12") = returnValue

            End Select

        Next

        Return True
    End Function

    Private Function CalculateMTPInvestments(ByVal dsData As DataSet, _
                                             ByVal intDataColumnIndex As Integer, _
                                             ByRef drInvestments As DataRow) As Boolean
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object

        For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName

            strExpression = "Sum(" + strColumnName + ")"
            strFilter = "BUDGET_TYPE = 'A'"
            returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
            drInvestments(dsData.Tables(0).Columns(k).ColumnName) = returnValue

            Select Case dsData.Tables(0).Columns(k).ColumnName
                Case "REVYEAR"
                    drInvestments("REVYEAR") = returnValue
                Case "RRT1"
                    drInvestments("RRT1") = returnValue
                Case "RRT2"
                    drInvestments("RRT2") = returnValue
                Case "RRT3"
                    drInvestments("RRT3") = returnValue
                Case "RRT4"
                    drInvestments("RRT4") = returnValue
                Case "RRT5"
                    drInvestments("RRT5") = returnValue
                Case "PrevRRT1"
                    drInvestments("PrevRRT1") = returnValue
                Case "PrevRRT2"
                    drInvestments("PrevRRT2") = returnValue
                Case "PrevRRT3"
                    drInvestments("PrevRRT3") = returnValue
                Case "PrevRRT4"
                    drInvestments("PrevRRT4") = returnValue
                Case "PrevRRT5"
                    drInvestments("PrevRRT5") = returnValue
                Case "RevYear"
                    drInvestments("RevYear") = returnValue
                Case "DiffYear"
                    drInvestments("DiffYear") = (Convert.ToDecimal(Nz(drInvestments("RevYear"), 0.0)) - Convert.ToDecimal(Nz(drInvestments("PrevRRT1"), 0.0)))
            End Select

        Next

        Return True
    End Function

    Private Function CalculateOriginalManufacturingCost(ByVal dsData As DataSet, _
                                               ByVal intDataColumnIndex As Integer, _
                                               ByRef drManufacturingCost As DataRow) As Boolean
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object

        For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName

            strExpression = "Sum(" + strColumnName + ")"
            strFilter = "BUDGET_TYPE = 'E'"
            returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
            drManufacturingCost(dsData.Tables(0).Columns(k).ColumnName) = returnValue

            Select Case dsData.Tables(0).Columns(k).ColumnName
                Case "FC_ACTUAL_1ST_HALF"
                    drManufacturingCost("ACTUAL_1ST_HALF") = returnValue
                Case "FC_REVISE_2ND_HALF"
                    drManufacturingCost("REVISE_2ND_HALF") = returnValue
                Case "FC_M1"
                    drManufacturingCost("M1") = returnValue
                Case "FC_M2"
                    drManufacturingCost("M2") = returnValue
                Case "FC_M3"
                    drManufacturingCost("M3") = returnValue
                Case "FC_M4"
                    drManufacturingCost("M4") = returnValue
                Case "FC_M5"
                    drManufacturingCost("M5") = returnValue
                Case "FC_M6"
                    drManufacturingCost("M6") = returnValue

                Case "FC_1ST_HALF"
                    drManufacturingCost("TOTAL_1ST_HALF") = returnValue
                Case "FC_2ND_HALF"
                    drManufacturingCost("TOTAL_2ND_HALF") = returnValue
                Case "MTP_RRT1_FC"
                    drManufacturingCost("MTP_RRT1") = returnValue

                    drManufacturingCost("DIFF_MTP") = Convert.ToDecimal(Nz(drManufacturingCost("TOTAL_YEAR"), 0.0)) - Convert.ToDecimal(Nz(drManufacturingCost("MTP_RRT1"), 0.0))
                Case "FC_TOTAL_YEAR"
                    drManufacturingCost("TOTAL_YEAR") = returnValue

                    '//Calculate
                    '{@INVEST_ACTUAL_1ST_HALF} + {@INVEST_REVISE_2ND_HALF}
                    drManufacturingCost("TOTAL_LAST_YEAR") = Convert.ToDecimal(Nz(drManufacturingCost("FC_ACTUAL_1ST_HALF"), 0.0)) + Convert.ToDecimal(Nz(drManufacturingCost("FC_REVISE_2ND_HALF"), 0.0))

                    '{@INVEST_TOTAL_YEAR} - {@INVEST_TOTAL_LAST_YEAR}
                    drManufacturingCost("DIFFERENCE") = Convert.ToDecimal(Nz(drManufacturingCost("TOTAL_YEAR"), 0.0)) - Convert.ToDecimal(Nz(drManufacturingCost("TOTAL_LAST_YEAR"), 0.0))

                Case "FC_M7"
                    drManufacturingCost("M7") = returnValue
                Case "FC_M8"
                    drManufacturingCost("M8") = returnValue
                Case "FC_M9"
                    drManufacturingCost("M9") = returnValue
                Case "FC_M10"
                    drManufacturingCost("M10") = returnValue
                Case "FC_M11"
                    drManufacturingCost("M11") = returnValue
                Case "FC_M12"
                    drManufacturingCost("M12") = returnValue

            End Select

        Next

        Return True
    End Function

    Private Function CalculateOriginalAdministrationCost(ByVal dsData As DataSet, _
                                  ByVal intDataColumnIndex As Integer, _
                                  ByRef drAdministrationCost As DataRow) As Boolean
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object

        For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName

            strExpression = "Sum(" + strColumnName + ")"
            strFilter = "BUDGET_TYPE = 'E'"
            returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
            drAdministrationCost(dsData.Tables(0).Columns(k).ColumnName) = returnValue

            Select Case dsData.Tables(0).Columns(k).ColumnName
                Case "ADMIN_ACTUAL_1ST_HALF"
                    drAdministrationCost("ACTUAL_1ST_HALF") = returnValue
                Case "ADMIN_REVISE_2ND_HALF"
                    drAdministrationCost("REVISE_2ND_HALF") = returnValue
                Case "ADMIN_M1"
                    drAdministrationCost("M1") = returnValue
                Case "ADMIN_M2"
                    drAdministrationCost("M2") = returnValue
                Case "ADMIN_M3"
                    drAdministrationCost("M3") = returnValue
                Case "ADMIN_M4"
                    drAdministrationCost("M4") = returnValue
                Case "ADMIN_M5"
                    drAdministrationCost("M5") = returnValue
                Case "ADMIN_M6"
                    drAdministrationCost("M6") = returnValue

                Case "ADMIN_1ST_HALF"
                    drAdministrationCost("TOTAL_1ST_HALF") = returnValue
                Case "ADMIN_2ND_HALF"
                    drAdministrationCost("TOTAL_2ND_HALF") = returnValue
                Case "MTP_RRT1_ADMIN"
                    drAdministrationCost("MTP_RRT1") = returnValue

                    drAdministrationCost("DIFF_MTP") = Convert.ToDecimal(Nz(drAdministrationCost("TOTAL_YEAR"), 0.0)) - Convert.ToDecimal(Nz(drAdministrationCost("MTP_RRT1"), 0.0))
                Case "ADMIN_TOTAL_YEAR"
                    drAdministrationCost("TOTAL_YEAR") = returnValue

                    '//Calculate
                    '{@INVEST_ACTUAL_1ST_HALF} + {@INVEST_REVISE_2ND_HALF}
                    drAdministrationCost("TOTAL_LAST_YEAR") = Convert.ToDecimal(Nz(drAdministrationCost("ACTUAL_1ST_HALF"), 0.0)) + Convert.ToDecimal(Nz(drAdministrationCost("REVISE_2ND_HALF"), 0.0))

                    '{@INVEST_TOTAL_YEAR} - {@INVEST_TOTAL_LAST_YEAR}
                    drAdministrationCost("DIFFERENCE") = Convert.ToDecimal(Nz(drAdministrationCost("TOTAL_YEAR"), 0.0)) - Convert.ToDecimal(Nz(drAdministrationCost("TOTAL_LAST_YEAR"), 0.0))

                Case "ADMIN_M7"
                    drAdministrationCost("M7") = returnValue
                Case "ADMIN_M8"
                    drAdministrationCost("M8") = returnValue
                Case "ADMIN_M9"
                    drAdministrationCost("M9") = returnValue
                Case "ADMIN_M10"
                    drAdministrationCost("M10") = returnValue
                Case "ADMIN_M11"
                    drAdministrationCost("M11") = returnValue
                Case "ADMIN_M12"
                    drAdministrationCost("M12") = returnValue
            End Select

        Next

        Return True
    End Function

    Private Function CalculateOriginalWorkingBudget(ByVal dsData As DataSet, _
                                            ByVal intDataColumnIndex As Integer, _
                                            ByRef drWorkingBudget As DataRow) As Boolean
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object = Nothing

        For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName

            If strColumnName <> "MTP_RRT1" Then
                strExpression = "Sum(" + strColumnName + ")"
                strFilter = "BUDGET_TYPE = 'E'"
                returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
                drWorkingBudget(dsData.Tables(0).Columns(k).ColumnName) = returnValue
            End If
           

            Select Case dsData.Tables(0).Columns(k).ColumnName
                Case "WB_ACTUAL_1ST_HALF"
                    'drWorkingBudget("ACTUAL_1ST_HALF") = returnValue
                    drWorkingBudget("ACTUAL_1ST_HALF") = Convert.ToDecimal(Nz(dsData.Tables(0).Rows(2)![WB_ACTUAL_1ST_HALF], 0.0))
                Case "WB_REVISE_2ND_HALF"
                    'drWorkingBudget("REVISE_2ND_HALF") = returnValue
                    drWorkingBudget("REVISE_2ND_HALF") = Convert.ToDecimal(Nz(dsData.Tables(0).Rows(2)![WB_REVISE_2ND_HALF], 0.0))
                Case "WB_M1"
                    drWorkingBudget("M1") = returnValue
                Case "WB_M2"
                    drWorkingBudget("M2") = returnValue
                Case "WB_M3"
                    drWorkingBudget("M3") = returnValue
                Case "WB_M4"
                    drWorkingBudget("M4") = returnValue
                Case "WB_M5"
                    drWorkingBudget("M5") = returnValue
                Case "WB_M6"
                    drWorkingBudget("M6") = returnValue

                Case "WB_TOTAL_1ST_HALF"
                    drWorkingBudget("TOTAL_1ST_HALF") = Convert.ToDecimal(Nz(drWorkingBudget("WB_M1"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("WB_M2"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("WB_M3"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("WB_M4"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("WB_M5"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("WB_M6"), 0.0))

                Case "MTPWB"
                    strExpression = "Sum(" + strColumnName + ")"
                    strFilter = "BUDGET_TYPE = ''"
                    returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
                    drWorkingBudget("MTP_RRT1") = returnValue

                Case "WB_TOTAL_2ND_HALF"
                    drWorkingBudget("TOTAL_2ND_HALF") = Convert.ToDecimal(Nz(drWorkingBudget("WB_M7"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("WB_M8"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("WB_M9"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("WB_M10"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("WB_M11"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("WB_M12"), 0.0))

                    'Case "ADMIN_TOTAL_YEAR"
                    drWorkingBudget("TOTAL_YEAR") = Convert.ToDecimal(Nz(drWorkingBudget("TOTAL_1ST_HALF"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("TOTAL_2ND_HALF"), 0.0))

                    '{@INVEST_ACTUAL_1ST_HALF} + {@INVEST_REVISE_2ND_HALF}
                    drWorkingBudget("TOTAL_LAST_YEAR") = Convert.ToDecimal(Nz(drWorkingBudget("ACTUAL_1ST_HALF"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("REVISE_2ND_HALF"), 0.0))

                    '{@INVEST_TOTAL_YEAR} - {@INVEST_TOTAL_LAST_YEAR}
                    drWorkingBudget("DIFFERENCE") = Convert.ToDecimal(Nz(drWorkingBudget("TOTAL_YEAR"), 0.0)) - Convert.ToDecimal(Nz(drWorkingBudget("TOTAL_LAST_YEAR"), 0.0))

                    'drWorkingBudget("DIFF_MTP") = Convert.ToDecimal(Nz(drWorkingBudget("TOTAL_YEAR"), 0.0)) - Convert.ToDecimal(Nz(drWorkingBudget("WB_MTP_RRT1"), 0.0))
                    drWorkingBudget("DIFF_MTP") = Convert.ToDecimal(Nz(drWorkingBudget("TOTAL_YEAR"), 0.0)) - Convert.ToDecimal(Nz(drWorkingBudget("MTP_RRT1"), 0.0))

                Case "WB_M7"
                    drWorkingBudget("M7") = returnValue
                Case "WB_M8"
                    drWorkingBudget("M8") = returnValue
                Case "WB_M9"
                    drWorkingBudget("M9") = returnValue
                Case "WB_M10"
                    drWorkingBudget("M10") = returnValue
                Case "WB_M11"
                    drWorkingBudget("M11") = returnValue
                Case "WB_M12"
                    drWorkingBudget("M12") = returnValue
            End Select

        Next

        Return True
    End Function

    Private Function GeneratOriginalExcel(ByVal dsData As DataSet, ByVal dtColumns As DataTable) As Boolean
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

            '//Setup Revise & Estimate Title
            ws.Cells(colStartIndex - 1, 5) = "1st Half'" & Me.numYear.Text.ToString.Substring(2, 2)
            ws.Range(ws.Cells(colStartIndex - 1, 5), ws.Cells(colStartIndex - 1, 10)).MergeCells = True
            ws.Range(ws.Cells(colStartIndex - 1, 5), ws.Cells(colStartIndex - 1, 10)).Font.Bold = True
            ws.Range(ws.Cells(colStartIndex - 1, 5), ws.Cells(colStartIndex - 1, 10)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            '//Setup Revise & Estimate Title
            ws.Cells(colStartIndex - 1, 12) = "2nd Half'" & Me.numYear.Text.ToString.Substring(2, 2)
            ws.Range(ws.Cells(colStartIndex - 1, 12), ws.Cells(colStartIndex - 1, 17)).MergeCells = True
            ws.Range(ws.Cells(colStartIndex - 1, 12), ws.Cells(colStartIndex - 1, 17)).Font.Bold = True
            ws.Range(ws.Cells(colStartIndex - 1, 12), ws.Cells(colStartIndex - 1, 17)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

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

            '//Setup Investments Line
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).ClearContents()
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).MergeCells = True
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).Value = "Investments"
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, colMax)).Font.Bold = True '// Add by Max 27/09/2012

            '//Setup budget order name column to be left align
            'ws.Range(ws.Cells(rowStartIndex + 1, 1), ws.Cells(rowMax, 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            ws.Range(ws.Cells(rowStartIndex, 2), ws.Cells(rowMax, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup Manufacturing Cost Line
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, 2)).Value = "Manufacturing Cost"
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, colMax)).Font.Bold = True '// Add by Max 27/09/2012

            '//Setup Administration Cost Line
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, 2)).Value = "Administration Cost"
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, colMax)).Font.Bold = True '// Add by Max 27/09/2012

            '//Setup Total Expense Line
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).Value = "Total Expense"
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, colMax)).Font.Bold = True '// Add by Max 27/09/2012

            '//Empry line
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).Value = ""
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup Working Budget Line
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).Value = "Working Budget"
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
            ws.Range(ws.Cells(colStartIndex, 10), ws.Cells(rowMax, 11)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 11), ws.Cells(rowMax, 17)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 17), ws.Cells(rowMax, 18)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            'ws.Range(ws.Cells(colStartIndex, 13), ws.Cells(rowMax, 14)).Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium
            'ws.Range(ws.Cells(colStartIndex, 15), ws.Cells(rowMax, 16)).Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium

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

        '      SELECT MAX_REV.BUDGET_YEAR, MAX_REV.PERIOD_TYPE, MAX_REV.ACCOUNT_NO, MAX_REV.ACCOUNT_NAME,
        row = dtColumns.NewRow()
        row("Column_Name") = "ACCOUNT_NO"
        row("Column_Title") = "Item"
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "ACCOUNT_NAME"
        row("Column_Title") = ""
        dtColumns.Rows.Add(row)

        ' MAX_REV.BUDGET_TYPE, MAX_REV.EXPENSE_TYPE,

        'SUM(ISNULL(ACTUAL_DATA.H1,0)) AS ACTUAL_1ST_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "ACTUAL_1ST_HALF"
        row("Column_Title") = "Actual 1st Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(REVISE_BUDGET.H2,0)) AS REVISE_2ND_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_2ND_HALF"
        row("Column_Title") = "Original 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ACTUAL_DATA.M7, 0)) AS M7,
        row = dtColumns.NewRow()
        row("Column_Name") = "M7"
        row("Column_Title") = "Jul'" & strHalfYear
        dtColumns.Rows.Add(row)

        ' SUM(ISNULL(ACTUAL_DATA.M8, 0)) AS M8,
        row = dtColumns.NewRow()
        row("Column_Name") = "M8"
        row("Column_Title") = "Aug'" & strHalfYear
        dtColumns.Rows.Add(row)

        ' SUM(ISNULL(ACTUAL_DATA.M9, 0)) AS M9,
        row = dtColumns.NewRow()
        row("Column_Name") = "M9"
        row("Column_Title") = "Sep'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M10, 0)) AS M10,
        row = dtColumns.NewRow()
        row("Column_Name") = "M10"
        row("Column_Title") = "Oct'" & strHalfYear
        dtColumns.Rows.Add(row)

        ' SUM(ISNULL(MASTER_DATA.M11, 0)) AS M11, 
        row = dtColumns.NewRow()
        row("Column_Name") = "M11"
        row("Column_Title") = "Nov'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M12, 0)) AS M12,
        row = dtColumns.NewRow()
        row("Column_Name") = "M12"
        row("Column_Title") = "Dec'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ACTUAL_DATA.M7, 0) + ISNULL(ACTUAL_DATA.M8, 0) + ISNULL(ACTUAL_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11,0) + ISNULL(MASTER_DATA.M12, 0)) AS ESTIMATE_2ND_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "ESTIMATE_2ND_HALF"
        row("Column_Title") = "Estimate 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM((ISNULL(ACTUAL_DATA.M7, 0) + ISNULL(ACTUAL_DATA.M8, 0) + ISNULL(ACTUAL_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11,0) + ISNULL(MASTER_DATA.M12, 0)) - ISNULL(REVISE_BUDGET.H2,0)) AS DIFFERENCE_2ND_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "DIFFERENCE_2ND_HALF"
        row("Column_Title") = "Diff 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ACTUAL_DATA.H1,0) + ISNULL(ACTUAL_DATA.M7, 0) + ISNULL(ACTUAL_DATA.M8, 0) + ISNULL(ACTUAL_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11,0) + ISNULL(MASTER_DATA.M12, 0)) AS ESTIMATE_TOTAL_YEAR,
        row = dtColumns.NewRow()
        row("Column_Name") = "ESTIMATE_TOTAL_YEAR"
        row("Column_Title") = "Estimate Year'" & strYear
        dtColumns.Rows.Add(row)

        '0 AS INVESTMENT_ACTUAL_1ST_HALF,
        '0 AS INVESTMENT_REVISE_2ND_HALF,
        '0 AS INVESTMENT_ACTUAL_JUL,
        '0 AS INVESTMENT_ACTUAL_AUG,
        '0 AS INVESTMENT_ACTUAL_SEP,
        '0 AS INVESTMENT_ESTIMATE_OCT,
        '0 AS INVESTMENT_ESTIMATE_NOV,
        '0 AS INVESTMENT_ESTIMATE_DEC,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(ACTUAL_DATA.M7, 0) ELSE 0 END) AS ADMIN_M7,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(ACTUAL_DATA.M7, 0) ELSE 0 END) AS FC_M7,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(ACTUAL_DATA.M8, 0) ELSE 0 END) AS ADMIN_M8,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(ACTUAL_DATA.M8, 0) ELSE 0 END) AS FC_M8,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(ACTUAL_DATA.M9, 0) ELSE 0 END) AS ADMIN_M9,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(ACTUAL_DATA.M9, 0) ELSE 0 END) AS FC_M9,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(MASTER_DATA.M10, 0) ELSE 0 END) AS ADMIN_M10,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(MASTER_DATA.M10, 0) ELSE 0 END) AS FC_M10,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(MASTER_DATA.M11, 0) ELSE 0 END) AS ADMIN_M11,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(MASTER_DATA.M11, 0) ELSE 0 END) AS FC_M11,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(MASTER_DATA.M12, 0) ELSE 0 END) AS ADMIN_M12,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(MASTER_DATA.M12, 0) ELSE 0 END) AS FC_M12,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(ACTUAL_DATA.H1,0) ELSE 0 END) AS ACTUAL_ADMIN_1ST_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(ACTUAL_DATA.H1,0) ELSE 0 END) AS ACTUAL_FC_1ST_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(REVISE_BUDGET.H2,0) ELSE 0 END) AS REVISE_ADMIN_2ND_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(REVISE_BUDGET.H2,0) ELSE 0 END) AS REVISE_FC_2ND_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN (ISNULL(ACTUAL_DATA.M7, 0) + ISNULL(ACTUAL_DATA.M8, 0) + ISNULL(ACTUAL_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11,0) + ISNULL(MASTER_DATA.M12, 0)) ELSE 0 END) AS ESTIMATE_ADMIN_2ND_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN (ISNULL(ACTUAL_DATA.M7, 0) + ISNULL(ACTUAL_DATA.M8, 0) + ISNULL(ACTUAL_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11,0) + ISNULL(MASTER_DATA.M12, 0)) ELSE 0 END) AS ESTIMATE_FC_2ND_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN (ISNULL(ACTUAL_DATA.M7, 0) + ISNULL(ACTUAL_DATA.M8, 0) + ISNULL(ACTUAL_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11,0) + ISNULL(MASTER_DATA.M12, 0) - ISNULL(REVISE_BUDGET.H2,0)) ELSE 0 END) AS DIFFERENCE_ADMIN_2ND_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN (ISNULL(ACTUAL_DATA.M7, 0) + ISNULL(ACTUAL_DATA.M8, 0) + ISNULL(ACTUAL_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11,0) + ISNULL(MASTER_DATA.M12, 0) - ISNULL(REVISE_BUDGET.H2,0)) ELSE 0 END) AS DIFFERENCE_FC_2ND_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(ACTUAL_DATA.H1,0) + (ISNULL(ACTUAL_DATA.M7, 0) + ISNULL(ACTUAL_DATA.M8, 0) + ISNULL(ACTUAL_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11,0) + ISNULL(MASTER_DATA.M12, 0)) ELSE 0 END) AS ESTIMATE_ADMIN_TOTAL_YEAR,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(ACTUAL_DATA.H1,0) + (ISNULL(ACTUAL_DATA.M7, 0) + ISNULL(ACTUAL_DATA.M8, 0) + ISNULL(ACTUAL_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11,0) + ISNULL(MASTER_DATA.M12, 0)) ELSE 0 END) AS ESTIMATE_FC_TOTAL_YEAR,
        '(SELECT ISNULL(WKH1,0) FROM BG_T_BUDGET_ADJUST WHERE BUDGET_YEAR=@BudgetYear AND PERIOD_TYPE=@PeriodType AND REV_NO= (SELECT MAX(REV_NO) AS REV_NO FROM BG_T_BUDGET_DATA WHERE BUDGET_YEAR = @BudgetYear AND PERIOD_TYPE = @PeriodType)) AS WB_ACTUAL_1ST_HALF,
        '(SELECT ISNULL(WKH2,0) FROM BG_T_BUDGET_ADJUST WHERE BUDGET_YEAR=@BudgetYear AND PERIOD_TYPE=@PeriodType AND REV_NO= (SELECT MAX(REV_NO) AS REV_NO FROM BG_T_BUDGET_DATA WHERE BUDGET_YEAR = @BudgetYear AND PERIOD_TYPE = @PeriodType)) AS WB_REVISE_2ND_HALF,
        '0 AS WB_ACTUAL_M7,
        '0 AS WB_ACTUAL_M8,
        '0 AS WB_ACTUAL_M9,
        'SUM(ISNULL(MIN_REV.M10,0) - ISNULL(MASTER_DATA.M10,0)) AS WB_ESTIMATE_M10,
        'SUM(ISNULL(MIN_REV.M11,0) - ISNULL(MASTER_DATA.M11,0)) AS WB_ESTIMATE_M11,
        'SUM(ISNULL(MIN_REV.M12,0) - ISNULL(MASTER_DATA.M12,0)) AS WB_ESTIMATE_M12,
        '0 AS WB_ESTIMATE_2ND_HALF,
        '0 AS WB_DIFFERENCE_2ND_HALF,
        '0 AS WB_ESTIMATE_TOTAL_YEAR


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
        Dim drManufacturingCost As DataRow = dtResult.NewRow
        Dim drAdministrationCost As DataRow = dtResult.NewRow
        Dim drTotalExpense As DataRow = dtResult.NewRow
        Dim drWorkingBudget As DataRow = dtResult.NewRow
        Dim drOutflowTotal As DataRow = dtResult.NewRow

        '//Calculate Investments
        CalculateEstimateInvestments(dsData, intDataColumnIndex, drInvestments)

        '//Calculate Manufacturing cost
        CalculateEstimateManufacturingCost(dsData, intDataColumnIndex, drManufacturingCost)

        '//Calculate Administration cost
        CalculateEstimateAdministrationCost(dsData, intDataColumnIndex, drAdministrationCost)

        '//Calculate Working Budget
        CalculateEstimateWorkingBudget(dsData, intDataColumnIndex, drWorkingBudget)

        Dim intGroupTotalIndex As Integer = 0
        For i As Integer = 0 To intGroupCount - 1

            strScript = strGroupColumnName + " = " + dtGroups.Rows(i)(0).ToString

            Dim arrRows As DataRow() = dtTmp.Select(strScript)

            For j As Integer = 0 To arrRows.Length - 1
                Dim drow(dtResult.Columns.Count - 1) As Object
                arrRows(j).ItemArray.CopyTo(drow, 0)
                dtResult.Rows.Add(drow)
            Next

            ''//Calculate Horizontal Column
            'For m As Integer = 0 To dtResult.Rows.Count - 1
            '    For n As Integer = intDataColumnIndex To dtResult.Columns.Count - 1
            '        Dim strColumnName As String = dtResult.Columns(n).ColumnName
            '        If String.IsNullOrEmpty(dtResult.Rows(m)![M1].ToString) Then
            '            Continue For
            '        End If
            '        If strColumnName = "TOTAL_1ST_HALF" Then
            '            dtResult.Rows(m)![TOTAL_1ST_HALF] = Convert.ToDecimal(dtResult.Rows(m)![M1]) + Convert.ToDecimal(dtResult.Rows(m)![M2]) + Convert.ToDecimal(dtResult.Rows(m)![M3]) + Convert.ToDecimal(dtResult.Rows(m)![M4]) + Convert.ToDecimal(dtResult.Rows(m)![M5]) + Convert.ToDecimal(dtResult.Rows(m)![M6])
            '        ElseIf strColumnName = "TOTAL_2ND_HALF" Then
            '            dtResult.Rows(m)![TOTAL_2ND_HALF] = Convert.ToDecimal(dtResult.Rows(m)![M7]) + Convert.ToDecimal(dtResult.Rows(m)![M8]) + Convert.ToDecimal(dtResult.Rows(m)![M9]) + Convert.ToDecimal(dtResult.Rows(m)![M10]) + Convert.ToDecimal(dtResult.Rows(m)![M11]) + Convert.ToDecimal(dtResult.Rows(m)![M12])
            '        ElseIf strColumnName = "TOTAL_YEAR" Then
            '            dtResult.Rows(m)![TOTAL_YEAR] = Convert.ToDecimal(dtResult.Rows(m)![TOTAL_1ST_HALF]) + Convert.ToDecimal(dtResult.Rows(m)![TOTAL_2ND_HALF])
            '        ElseIf strColumnName = "TOTAL_LAST_YEAR" Then
            '            '{DetailByAccountCode.ACTUAL_1ST_HALF} + {DetailByAccountCode.REVISE_2ND_HALF}
            '            dtResult.Rows(m)![TOTAL_LAST_YEAR] = Convert.ToDecimal(dtResult.Rows(m)![ACTUAL_1ST_HALF]) + Convert.ToDecimal(dtResult.Rows(m)![REVISE_2ND_HALF])
            '        ElseIf strColumnName = "DIFFERENCE" Then
            '            '{@TotalYear} - {@TotalLastYear}
            '            dtResult.Rows(m)![DIFFERENCE] = Convert.ToDecimal(dtResult.Rows(m)![TOTAL_YEAR]) - Convert.ToDecimal(dtResult.Rows(m)![TOTAL_LAST_YEAR])
            '        End If

            '    Next
            'Next
            'dtResult.AcceptChanges()

            '//Calculate total for each group
            Dim drTotal As DataRow = dtResult.NewRow
            For k As Integer = intDataColumnIndex To dtResult.Columns.Count - 1

                Dim strColumnName As String = dtResult.Columns(k).ColumnName
                strExpression = "Sum(" + strColumnName + ")"
                strFilter = strScript
                returnValue = dtResult.Compute(strExpression, strFilter)
                drTotal(dtResult.Columns(k).ColumnName) = returnValue

            Next
            '//Set Group header
            drTotal("ACCOUNT_NO") = GetGroupExpensesTitle(dtGroups.Rows(i)(0).ToString)

            '//Add total cost
            dtResult.Rows.InsertAt(drTotal, intGroupTotalIndex)

            '//Add one empty row
            drEmpty = dtResult.NewRow
            dtResult.Rows.Add(drEmpty)

            'dtResult.TableName = arrRows(0)(strGroupColumnName).ToString & " " & arrRows(0)(strGroupColumnTitle).ToString
            dtResult.TableName = "Estimate Budget"

            'intGroupTotalIndex = intGroupTotalIndex + dtResult.Rows.Count
            intGroupTotalIndex = intGroupTotalIndex + CInt(arrRows.Length) + 2

        Next
        '//Set Data to Account No.
        SetAccountNoText(drInvestments, drManufacturingCost, drAdministrationCost, drTotalExpense, drWorkingBudget, drOutflowTotal)

        '//Calculate Total Expense
        CalculateTotalExpense(dtResult, intDataColumnIndex, drManufacturingCost, drAdministrationCost, drTotalExpense)

        '//Calculate Outflow Total 
        CalculateOutflowTotal(dtResult, intDataColumnIndex, drTotalExpense, drWorkingBudget, drInvestments, drOutflowTotal)

        '//Add Investments
        dtResult.Rows.InsertAt(drInvestments, 0)

        '//Add one empty row
        drEmpty = dtResult.NewRow
        dtResult.Rows.InsertAt(drEmpty, 1)

        '//Add Manufacturing cost
        dtResult.Rows.Add(drManufacturingCost)

        '//Add Administration cost
        dtResult.Rows.Add(drAdministrationCost)

        '//Add Total Expense
        dtResult.Rows.Add(drTotalExpense)

        '//Add one empty row
        drEmpty = dtResult.NewRow
        dtResult.Rows.Add(drEmpty)

        '//Add Working Budget
        dtResult.Rows.Add(drWorkingBudget)

        '//Add Outflow Total (Investment;Expenses)
        dtResult.Rows.Add(drOutflowTotal)

        '//Return data table
        dsResult.Tables.Add(dtResult)

        Return dsResult
    End Function

    Private Function CalculateEstimateInvestments(ByVal dsData As DataSet, _
                                        ByVal intDataColumnIndex As Integer, _
                                        ByRef drInvestments As DataRow) As Boolean
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object

        For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName

            strExpression = "Sum(" + strColumnName + ")"
            strFilter = "BUDGET_TYPE = 'A'"
            returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
            drInvestments(dsData.Tables(0).Columns(k).ColumnName) = returnValue

            Select Case dsData.Tables(0).Columns(k).ColumnName
                Case "INVESTMENT_ACTUAL_1ST_HALF"
                    drInvestments("ACTUAL_1ST_HALF") = returnValue
                Case "INVESTMENT_REVISE_2ND_HALF"
                    drInvestments("REVISE_2ND_HALF") = returnValue

                Case "INVESTMENT_ACTUAL_JUL"
                    drInvestments("M7") = returnValue
                Case "INVESTMENT_ACTUAL_AUG"
                    drInvestments("M8") = returnValue
                Case "INVESTMENT_ACTUAL_SEP"
                    drInvestments("M9") = returnValue
                Case "INVESTMENT_ESTIMATE_OCT"
                    drInvestments("M10") = returnValue
                Case "INVESTMENT_ESTIMATE_NOV"
                    drInvestments("M11") = returnValue
                Case "INVESTMENT_ESTIMATE_DEC"
                    drInvestments("M12") = returnValue

                    '//Calculate
                    'SUM(ISNULL(ACTUAL_DATA.M7, 0) + ISNULL(ACTUAL_DATA.M8, 0) + ISNULL(ACTUAL_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11,0) + ISNULL(MASTER_DATA.M12, 0)) AS ESTIMATE_2ND_HALF,
                    drInvestments("ESTIMATE_2ND_HALF") = Convert.ToDecimal(Nz(drInvestments("INVESTMENT_ACTUAL_JUL"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("INVESTMENT_ACTUAL_AUG"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("INVESTMENT_ACTUAL_SEP"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("INVESTMENT_ESTIMATE_OCT"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("INVESTMENT_ESTIMATE_NOV"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("INVESTMENT_ESTIMATE_DEC"), 0.0))

                    'SUM((ISNULL(ACTUAL_DATA.M7, 0) + ISNULL(ACTUAL_DATA.M8, 0) + ISNULL(ACTUAL_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11,0) + ISNULL(MASTER_DATA.M12, 0)) - ISNULL(REVISE_BUDGET.H2,0)) AS DIFFERENCE_2ND_HALF,
                    drInvestments("DIFFERENCE_2ND_HALF") = Convert.ToDecimal(Nz(drInvestments("ESTIMATE_2ND_HALF"), 0.0)) - Convert.ToDecimal(Nz(drInvestments("REVISE_2ND_HALF"), 0.0))

                    'SUM(ISNULL(ACTUAL_DATA.H1,0) + ISNULL(ACTUAL_DATA.M7, 0) + ISNULL(ACTUAL_DATA.M8, 0) + ISNULL(ACTUAL_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11,0) + ISNULL(MASTER_DATA.M12, 0)) AS ESTIMATE_TOTAL_YEAR,
                    drInvestments("ESTIMATE_TOTAL_YEAR") = Convert.ToDecimal(Nz(drInvestments("INVESTMENT_ACTUAL_1ST_HALF"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("INVESTMENT_ACTUAL_JUL"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("INVESTMENT_ACTUAL_AUG"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("INVESTMENT_ACTUAL_SEP"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("INVESTMENT_ESTIMATE_OCT"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("INVESTMENT_ESTIMATE_NOV"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("INVESTMENT_ESTIMATE_DEC"), 0.0))


            End Select

        Next

        Return True
    End Function

    Private Function CalculateEstimateManufacturingCost(ByVal dsData As DataSet, _
                                             ByVal intDataColumnIndex As Integer, _
                                             ByRef drManufacturingCost As DataRow) As Boolean
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object

        For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName

            strExpression = "Sum(" + strColumnName + ")"
            strFilter = "BUDGET_TYPE = 'E'"
            returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
            drManufacturingCost(dsData.Tables(0).Columns(k).ColumnName) = returnValue

            Select Case dsData.Tables(0).Columns(k).ColumnName
                Case "ACTUAL_FC_1ST_HALF"
                    drManufacturingCost("ACTUAL_1ST_HALF") = returnValue
                Case "REVISE_FC_2ND_HALF"
                    drManufacturingCost("REVISE_2ND_HALF") = returnValue
                Case "FC_M7"
                    drManufacturingCost("M7") = returnValue
                Case "FC_M8"
                    drManufacturingCost("M8") = returnValue
                Case "FC_M9"
                    drManufacturingCost("M9") = returnValue
                Case "FC_M10"
                    drManufacturingCost("M10") = returnValue
                Case "FC_M11"
                    drManufacturingCost("M11") = returnValue
                Case "FC_M12"
                    drManufacturingCost("M12") = returnValue

                Case "ESTIMATE_FC_2ND_HALF"
                    drManufacturingCost("ESTIMATE_2ND_HALF") = returnValue

                Case "DIFFERENCE_FC_2ND_HALF"
                    drManufacturingCost("DIFFERENCE_2ND_HALF") = returnValue
                Case "ESTIMATE_FC_TOTAL_YEAR"
                    drManufacturingCost("ESTIMATE_TOTAL_YEAR") = returnValue


            End Select

        Next

        Return True
    End Function

    Private Function CalculateEstimateAdministrationCost(ByVal dsData As DataSet, _
                              ByVal intDataColumnIndex As Integer, _
                              ByRef drAdministrationCost As DataRow) As Boolean
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object

        For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName

            strExpression = "Sum(" + strColumnName + ")"
            strFilter = "BUDGET_TYPE = 'E'"
            returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
            drAdministrationCost(dsData.Tables(0).Columns(k).ColumnName) = returnValue

            Select Case dsData.Tables(0).Columns(k).ColumnName
                Case "ACTUAL_ADMIN_1ST_HALF"
                    drAdministrationCost("ACTUAL_1ST_HALF") = returnValue
                Case "REVISE_ADMIN_2ND_HALF"
                    drAdministrationCost("REVISE_2ND_HALF") = returnValue
                Case "ADMIN_M7"
                    drAdministrationCost("M7") = returnValue
                Case "ADMIN_M8"
                    drAdministrationCost("M8") = returnValue
                Case "ADMIN_M9"
                    drAdministrationCost("M9") = returnValue
                Case "ADMIN_M10"
                    drAdministrationCost("M10") = returnValue
                Case "ADMIN_M11"
                    drAdministrationCost("M11") = returnValue
                Case "ADMIN_M12"
                    drAdministrationCost("M12") = returnValue

                Case "ESTIMATE_ADMIN_2ND_HALF"
                    drAdministrationCost("ESTIMATE_2ND_HALF") = returnValue
                Case "DIFFERENCE_ADMIN_2ND_HALF"
                    drAdministrationCost("DIFFERENCE_2ND_HALF") = returnValue
                Case "ESTIMATE_ADMIN_TOTAL_YEAR"
                    drAdministrationCost("ESTIMATE_TOTAL_YEAR") = returnValue

            End Select

        Next

        Return True
    End Function

    Private Function CalculateEstimateWorkingBudget(ByVal dsData As DataSet, _
                                            ByVal intDataColumnIndex As Integer, _
                                            ByRef drWorkingBudget As DataRow) As Boolean
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object

        For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName

            strExpression = "Sum(" + strColumnName + ")"
            strFilter = "BUDGET_TYPE = 'E'"
            returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
            drWorkingBudget(dsData.Tables(0).Columns(k).ColumnName) = returnValue

            Select Case dsData.Tables(0).Columns(k).ColumnName
                Case "WB_ACTUAL_1ST_HALF"
                    'drWorkingBudget("ACTUAL_1ST_HALF") = returnValue
                    drWorkingBudget("ACTUAL_1ST_HALF") = Convert.ToDecimal(Nz(dsData.Tables(0).Rows(0)![WB_ACTUAL_1ST_HALF], 0.0))
                Case "WB_REVISE_2ND_HALF"
                    'drWorkingBudget("REVISE_2ND_HALF") = returnValue
                    drWorkingBudget("REVISE_2ND_HALF") = Convert.ToDecimal(Nz(dsData.Tables(0).Rows(0)![WB_REVISE_2ND_HALF], 0.0))
                Case "WB_ACTUAL_M7"
                    drWorkingBudget("M7") = returnValue
                Case "WB_ACTUAL_M8"
                    drWorkingBudget("M8") = returnValue
                Case "WB_ACTUAL_M9"
                    drWorkingBudget("M9") = returnValue
                Case "WB_ESTIMATE_M10"
                    drWorkingBudget("M10") = returnValue
                Case "WB_ESTIMATE_M11"
                    drWorkingBudget("M11") = returnValue
                Case "WB_ESTIMATE_M12"
                    drWorkingBudget("M12") = returnValue
                Case "WB_ESTIMATE_2ND_HALF"
                    drWorkingBudget("ESTIMATE_2ND_HALF") = Convert.ToDecimal(Nz(drWorkingBudget("WB_ACTUAL_M7"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("WB_ACTUAL_M8"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("WB_ACTUAL_M9"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("WB_ESTIMATE_M10"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("WB_ESTIMATE_M11"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("WB_ESTIMATE_M12"), 0.0))
                Case "WB_DIFFERENCE_2ND_HALF"
                    drWorkingBudget("DIFFERENCE_2ND_HALF") = Convert.ToDecimal(Nz(drWorkingBudget("ESTIMATE_2ND_HALF"), 0.0)) - Convert.ToDecimal(Nz(drWorkingBudget("REVISE_2ND_HALF"), 0.0))
                Case "WB_ESTIMATE_TOTAL_YEAR"
                    drWorkingBudget("ESTIMATE_TOTAL_YEAR") = Convert.ToDecimal(Nz(drWorkingBudget("ACTUAL_1ST_HALF"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("ESTIMATE_2ND_HALF"), 0.0))
            End Select

        Next

        Return True
    End Function

    Private Function GeneratEstimateExcel(ByVal dsData As DataSet, ByVal dtColumns As DataTable) As Boolean
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
            'Dim intAuthorizeEnd As Integer

            '//Setup Investments Line
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).ClearContents()
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).MergeCells = True
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).Value = "Investments"
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, colMax)).Font.Bold = True '// Add by Max 27/09/2012

            '//Setup budget order name column to be left align
            'ws.Range(ws.Cells(rowStartIndex + 1, 1), ws.Cells(rowMax, 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            ws.Range(ws.Cells(rowStartIndex, 2), ws.Cells(rowMax, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup Manufacturing Cost Line
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, 2)).Value = "Manufacturing Cost"
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, colMax)).Font.Bold = True '// Add by Max 27/09/2012

            '//Setup Administration Cost Line
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, 2)).Value = "Administration Cost"
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, colMax)).Font.Bold = True '// Add by Max 27/09/2012

            '//Setup Total Expense Line
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).Value = "Total Expense"
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, colMax)).Font.Bold = True '// Add by Max 27/09/2012

            '//Empry line
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).Value = ""
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup Working Budget Line
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).Value = "Working Budget"
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup Outflow Total Line
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).Value = "Outflow Total (Investment;Expenses)"
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, colMax)).Font.Bold = True

            '//Setup Budget order number Title
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(colStartIndex, 2)).ClearContents()
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(colStartIndex, 2)).MergeCells = True
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(colStartIndex, 2)).Font.Bold = True
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(colStartIndex, 2)).Value = "Item"
            ws.Range(ws.Cells(colStartIndex - 1, 1), ws.Cells(colStartIndex, 2)).VerticalAlignment = Excel.XlHAlign.xlHAlignCenter

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

        'SELECT MAX_REV.BUDGET_YEAR, MAX_REV.PERIOD_TYPE, MAX_REV.ACCOUNT_NO , 
        row = dtColumns.NewRow()
        row("Column_Name") = "ACCOUNT_NO"
        row("Column_Title") = ""
        dtColumns.Rows.Add(row)

        'MAX_REV.ACCOUNT_NAME , MAX_REV.BUDGET_TYPE , MAX_REV.EXPENSE_TYPE,
        row = dtColumns.NewRow()
        row("Column_Name") = "ACCOUNT_NAME"
        row("Column_Title") = ""
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ORIGINAL_BUDGET.H1,0)) AS ORIGINAL_1ST_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "ORIGINAL_1ST_HALF"
        row("Column_Title") = "Original 1st Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ACTUAL_DATA.M1, 0)) AS ACTUAL_M1,
        row = dtColumns.NewRow()
        row("Column_Name") = "ACTUAL_M1"
        row("Column_Title") = "Jan'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ACTUAL_DATA.M2, 0)) AS ACTUAL_M2,
        row = dtColumns.NewRow()
        row("Column_Name") = "ACTUAL_M2"
        row("Column_Title") = "Feb'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ACTUAL_DATA.M3, 0)) AS ACTUAL_M3,
        row = dtColumns.NewRow()
        row("Column_Name") = "ACTUAL_M3"
        row("Column_Title") = "Mar'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ESTIMATE_BUDGET.M4, 0)) AS ESTIMATE_M4,
        row = dtColumns.NewRow()
        row("Column_Name") = "ESTIMATE_M4"
        row("Column_Title") = "Apr'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ESTIMATE_BUDGET.M5, 0)) AS ESTIMATE_M5,
        row = dtColumns.NewRow()
        row("Column_Name") = "ESTIMATE_M5"
        row("Column_Title") = "May'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ESTIMATE_BUDGET.M6, 0)) AS ESTIMATE_M6,
        row = dtColumns.NewRow()
        row("Column_Name") = "ESTIMATE_M6"
        row("Column_Title") = "Jun'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ACTUAL_DATA.M1, 0) + ISNULL(ACTUAL_DATA.M2, 0) + ISNULL(ACTUAL_DATA.M3, 0) + ISNULL(ESTIMATE_BUDGET.M4, 0) + ISNULL(ESTIMATE_BUDGET.M5, 0) + ISNULL(ESTIMATE_BUDGET.M6, 0)) AS ESTIMATE_1ST_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "ESTIMATE_1ST_HALF"
        row("Column_Title") = "Estimate 1st Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM((ISNULL(ACTUAL_DATA.M1, 0) + ISNULL(ACTUAL_DATA.M2, 0) + ISNULL(ACTUAL_DATA.M3, 0) + ISNULL(ESTIMATE_BUDGET.M4, 0) + ISNULL(ESTIMATE_BUDGET.M5, 0) + ISNULL(ESTIMATE_BUDGET.M6, 0)) - ISNULL(ORIGINAL_BUDGET.H1,0)) AS DIFF_1ST_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "DIFF_1ST_HALF"
        row("Column_Title") = "Diff 1st Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ORIGINAL_BUDGET.H2,0)) AS ORIGINAL_2ND_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "ORIGINAL_2ND_HALF"
        row("Column_Title") = "Original 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M7, 0)) AS REVISE_M7, 
        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_M7"
        row("Column_Title") = "Jul'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M8, 0)) AS REVISE_M8, 
        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_M8"
        row("Column_Title") = "Aug'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M9, 0)) AS REVISE_M9,
        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_M9"
        row("Column_Title") = "Sep'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M10, 0)) AS REVISE_M10, 
        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_M10"
        row("Column_Title") = "Oct'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M11, 0)) AS REVISE_M11, 
        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_M11"
        row("Column_Title") = "Nov'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M12, 0)) AS REVISE_M12,
        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_M12"
        row("Column_Title") = "Dec'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M7, 0) + ISNULL(MASTER_DATA.M8, 0) + ISNULL(MASTER_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11, 0) + ISNULL(MASTER_DATA.M12, 0)) AS REVISE_2ND_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_2ND_HALF"
        row("Column_Title") = "Revise 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M7, 0) + ISNULL(MASTER_DATA.M8, 0) + ISNULL(MASTER_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11, 0) + ISNULL(MASTER_DATA.M12, 0) - ISNULL(ORIGINAL_BUDGET.H2,0)) AS DIFF_2ND_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "DIFF_2ND_HALF"
        row("Column_Title") = "Diff 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ACTUAL_DATA.M1, 0) + ISNULL(ACTUAL_DATA.M2, 0) + ISNULL(ACTUAL_DATA.M3, 0) + ISNULL(ESTIMATE_BUDGET.M4, 0) + ISNULL(ESTIMATE_BUDGET.M5, 0) + ISNULL(ESTIMATE_BUDGET.M6, 0) + ISNULL(MASTER_DATA.M7, 0) + ISNULL(MASTER_DATA.M8, 0) + ISNULL(MASTER_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11, 0) + ISNULL(MASTER_DATA.M12, 0)) AS REVISE_TOTAL_YEAR,
        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_TOTAL_YEAR"
        row("Column_Title") = "Revise Year'" & strYear
        dtColumns.Rows.Add(row)

        'SUM((ISNULL(ACTUAL_DATA.M1, 0) + ISNULL(ACTUAL_DATA.M2, 0) + ISNULL(ACTUAL_DATA.M3, 0) + ISNULL(ESTIMATE_BUDGET.M4, 0) + ISNULL(ESTIMATE_BUDGET.M5, 0) + ISNULL(ESTIMATE_BUDGET.M6, 0)) - ISNULL(ORIGINAL_BUDGET.H1,0) + ISNULL(MASTER_DATA.M7, 0) + ISNULL(MASTER_DATA.M8, 0) + ISNULL(MASTER_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11, 0) + ISNULL(MASTER_DATA.M12, 0) - ISNULL(ORIGINAL_BUDGET.H2,0) ) AS DIFF_TOTAL_YEAR,
        row = dtColumns.NewRow()
        row("Column_Name") = "DIFF_TOTAL_YEAR"
        row("Column_Title") = "Diff Year'" & strYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.RRT1 ,0)) AS RRT1,
        'SUM(ISNULL(MASTER_DATA.RRT2 ,0)) AS RRT2,
        'SUM(ISNULL(MASTER_DATA.RRT3 ,0)) AS RRT3,
        'SUM(ISNULL(MASTER_DATA.RRT4 ,0)) AS RRT4,
        'SUM(ISNULL(MASTER_DATA.RRT5 ,0)) AS RRT5,
        '0 AS INVESTMENT_ORIGINAL_1ST_HALF,
        '0 AS INVESTMENT_ACTUAL_JAN,
        '0 AS INVESTMENT_ACTUAL_FEB,
        '0 AS INVESTMENT_ACTUAL_MAR,
        '0 AS INVESTMENT_ESTIMATE_APR,
        '0 AS INVESTMENT_ESTIMATE_MAY,
        '0 AS INVESTMENT_ESTIMATE_JUN,
        '0 AS INVESTMENT_ORIGINAL_2ND_HALF,
        '0 AS INVESTMENT_REVISE_JUL,
        '0 AS INVESTMENT_REVISE_AUG,
        '0 AS INVESTMENT_REVISE_SEP,
        '0 AS INVESTMENT_REVISE_OCT,
        '0 AS INVESTMENT_REVISE_NOV,
        '0 AS INVESTMENT_REVISE_DEC,
        '0 AS INVESTMENT_RRT1,
        '0 AS INVESTMENT_RRT2,
        '0 AS INVESTMENT_RRT3,
        '0 AS INVESTMENT_RRT4,
        '0 AS INVESTMENT_RRT5,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(ORIGINAL_BUDGET.H2,0) ELSE 0 END) AS ADMIN_ORIGINAL_1ST_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(ORIGINAL_BUDGET.H2,0) ELSE 0 END) AS FC_ORIGINAL_1ST_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(ACTUAL_DATA.M1, 0) ELSE 0 END) AS ADMIN_ACTUAL_M1,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(ACTUAL_DATA.M1, 0) ELSE 0 END) AS FC_ACTUAL_M1,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(ACTUAL_DATA.M2, 0) ELSE 0 END) AS ADMIN_ACTUAL_M2,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(ACTUAL_DATA.M2, 0) ELSE 0 END) AS FC_ACTUAL_M2,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(ACTUAL_DATA.M3, 0) ELSE 0 END) AS ADMIN_ACTUAL_M3,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(ACTUAL_DATA.M3, 0) ELSE 0 END) AS FC_ACTUAL_M3,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(ESTIMATE_BUDGET.M4, 0) ELSE 0 END) AS ADMIN_ESTIMATE_M4,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(ESTIMATE_BUDGET.M4, 0) ELSE 0 END) AS FC_ESTIMATE_M4,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(ESTIMATE_BUDGET.M5, 0) ELSE 0 END) AS ADMIN_ESTIMATE_M5,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(ESTIMATE_BUDGET.M5, 0) ELSE 0 END) AS FC_ESTIMATE_M5,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(ESTIMATE_BUDGET.M6, 0) ELSE 0 END) AS ADMIN_ESTIMATE_M6,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(ESTIMATE_BUDGET.M6, 0) ELSE 0 END) AS FC_ESTIMATE_M6,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(ACTUAL_DATA.M1, 0) + ISNULL(ACTUAL_DATA.M2, 0) + ISNULL(ACTUAL_DATA.M3, 0) + ISNULL(ESTIMATE_BUDGET.M4, 0) + ISNULL(ESTIMATE_BUDGET.M5, 0) + ISNULL(ESTIMATE_BUDGET.M6, 0) ELSE 0 END) AS ADMIN_ESTIMATE_1ST_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(ACTUAL_DATA.M1, 0) + ISNULL(ACTUAL_DATA.M2, 0) + ISNULL(ACTUAL_DATA.M3, 0) + ISNULL(ESTIMATE_BUDGET.M4, 0) + ISNULL(ESTIMATE_BUDGET.M5, 0) + ISNULL(ESTIMATE_BUDGET.M6, 0) ELSE 0 END) AS FC_ESTIMATE_1ST_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN (ISNULL(ACTUAL_DATA.M1, 0) + ISNULL(ACTUAL_DATA.M2, 0) + ISNULL(ACTUAL_DATA.M3, 0) + ISNULL(ESTIMATE_BUDGET.M4, 0) + ISNULL(ESTIMATE_BUDGET.M5, 0) + ISNULL(ESTIMATE_BUDGET.M6, 0) - ISNULL(ORIGINAL_BUDGET.H1,0)) ELSE 0 END)AS ADMIN_DIFF_1ST_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN (ISNULL(ACTUAL_DATA.M1, 0) + ISNULL(ACTUAL_DATA.M2, 0) + ISNULL(ACTUAL_DATA.M3, 0) + ISNULL(ESTIMATE_BUDGET.M4, 0) + ISNULL(ESTIMATE_BUDGET.M5, 0) + ISNULL(ESTIMATE_BUDGET.M6, 0) - ISNULL(ORIGINAL_BUDGET.H1,0)) ELSE 0 END) AS FC_DIFF_1ST_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(ORIGINAL_BUDGET.H2,0) ELSE 0 END) AS ADMIN_ORIGINAL_2ND_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(ORIGINAL_BUDGET.H2,0) ELSE 0 END) AS FC_ORIGINAL_2ND_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(MASTER_DATA.M7, 0) ELSE 0 END) AS ADMIN_REVISE_M7,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(MASTER_DATA.M7, 0) ELSE 0 END) AS FC_REVISE_M7,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(MASTER_DATA.M8, 0) ELSE 0 END) AS ADMIN_REVISE_M8,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(MASTER_DATA.M8, 0) ELSE 0 END) AS FC_REVISE_M8,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(MASTER_DATA.M9, 0) ELSE 0 END) AS ADMIN_REVISE_M9,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(MASTER_DATA.M9, 0) ELSE 0 END) AS FC_REVISE_M9,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(MASTER_DATA.M10, 0) ELSE 0 END) AS ADMIN_REVISE_M10,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(MASTER_DATA.M10, 0) ELSE 0 END) AS FC_REVISE_M10,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(MASTER_DATA.M11, 0) ELSE 0 END) AS ADMIN_REVISE_M11,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(MASTER_DATA.M11, 0) ELSE 0 END) AS FC_REVISE_M11,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(MASTER_DATA.M12, 0) ELSE 0 END) AS ADMIN_REVISE_M12,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(MASTER_DATA.M12, 0) ELSE 0 END) AS FC_REVISE_M12,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(MASTER_DATA.M7, 0) + ISNULL(MASTER_DATA.M8, 0) + ISNULL(MASTER_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11, 0) + ISNULL(MASTER_DATA.M12, 0) ELSE 0 END) AS ADMIN_REVISE_2ND_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(MASTER_DATA.M7, 0) + ISNULL(MASTER_DATA.M8, 0) + ISNULL(MASTER_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11, 0) + ISNULL(MASTER_DATA.M12, 0) ELSE 0 END) AS FC_REVISE_2ND_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN (ISNULL(MASTER_DATA.M7, 0) + ISNULL(MASTER_DATA.M8, 0) + ISNULL(MASTER_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11, 0) + ISNULL(MASTER_DATA.M12, 0) - ISNULL(ORIGINAL_BUDGET.H2,0)) ELSE 0 END) AS ADMIN_DIFF_2ND_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN (ISNULL(MASTER_DATA.M7, 0) + ISNULL(MASTER_DATA.M8, 0) + ISNULL(MASTER_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11, 0) + ISNULL(MASTER_DATA.M12, 0) - ISNULL(ORIGINAL_BUDGET.H2,0)) ELSE 0 END) AS FC_DIFF_2ND_HALF,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ( ISNULL(ACTUAL_DATA.M1, 0) + ISNULL(ACTUAL_DATA.M2, 0) + ISNULL(ACTUAL_DATA.M3, 0) + ISNULL(ESTIMATE_BUDGET.M4, 0) + ISNULL(ESTIMATE_BUDGET.M5, 0) + ISNULL(ESTIMATE_BUDGET.M6, 0) +  ISNULL(MASTER_DATA.M7, 0) + ISNULL(MASTER_DATA.M8, 0) + ISNULL(MASTER_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11, 0) + ISNULL(MASTER_DATA.M12, 0) ) ELSE 0 END) AS ADMIN_REVISE_TOTAL_YEAR,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ( ISNULL(ACTUAL_DATA.M1, 0) + ISNULL(ACTUAL_DATA.M2, 0) + ISNULL(ACTUAL_DATA.M3, 0) + ISNULL(ESTIMATE_BUDGET.M4, 0) + ISNULL(ESTIMATE_BUDGET.M5, 0) + ISNULL(ESTIMATE_BUDGET.M6, 0) +  ISNULL(MASTER_DATA.M7, 0) + ISNULL(MASTER_DATA.M8, 0) + ISNULL(MASTER_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11, 0) + ISNULL(MASTER_DATA.M12, 0) ) ELSE 0 END) AS FC_REVISE_TOTAL_YEAR,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN (ISNULL(ACTUAL_DATA.M1, 0) + ISNULL(ACTUAL_DATA.M2, 0) + ISNULL(ACTUAL_DATA.M3, 0) + ISNULL(ESTIMATE_BUDGET.M4, 0) + ISNULL(ESTIMATE_BUDGET.M5, 0) + ISNULL(ESTIMATE_BUDGET.M6, 0) - ISNULL(ORIGINAL_BUDGET.H1,0)) + (ISNULL(MASTER_DATA.M7, 0) + ISNULL(MASTER_DATA.M8, 0) + ISNULL(MASTER_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11, 0) + ISNULL(MASTER_DATA.M12, 0) - ISNULL(ORIGINAL_BUDGET.H2,0)) ELSE 0 END) AS ADMIN_DIFF_TOTAL_YEAR,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN (ISNULL(ACTUAL_DATA.M1, 0) + ISNULL(ACTUAL_DATA.M2, 0) + ISNULL(ACTUAL_DATA.M3, 0) + ISNULL(ESTIMATE_BUDGET.M4, 0) + ISNULL(ESTIMATE_BUDGET.M5, 0) + ISNULL(ESTIMATE_BUDGET.M6, 0) - ISNULL(ORIGINAL_BUDGET.H1,0)) + (ISNULL(MASTER_DATA.M7, 0) + ISNULL(MASTER_DATA.M8, 0) + ISNULL(MASTER_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11, 0) + ISNULL(MASTER_DATA.M12, 0) - ISNULL(ORIGINAL_BUDGET.H2,0)) ELSE 0 END) AS FC_DIFF_TOTAL_YEAR,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(MASTER_DATA.RRT1, 0) ELSE 0 END) AS ADMIN_RRT1,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(MASTER_DATA.RRT1, 0) ELSE 0 END) AS FC_RRT1,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(MASTER_DATA.RRT2, 0) ELSE 0 END) AS ADMIN_RRT2,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(MASTER_DATA.RRT2, 0) ELSE 0 END) AS FC_RRT2,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(MASTER_DATA.RRT3, 0) ELSE 0 END) AS ADMIN_RRT3,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(MASTER_DATA.RRT3, 0) ELSE 0 END) AS FC_RRT3,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(MASTER_DATA.RRT4, 0) ELSE 0 END) AS ADMIN_RRT4,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(MASTER_DATA.RRT4, 0) ELSE 0 END) AS FC_RRT4,
        'SUM(CASE WHEN MAX_REV.COST = @Admin  THEN ISNULL(MASTER_DATA.RRT5, 0) ELSE 0 END) AS ADMIN_RRT5,
        'SUM(CASE WHEN MAX_REV.COST = @FC  THEN ISNULL(MASTER_DATA.RRT5, 0) ELSE 0 END) AS FC_RRT5,
        '(SELECT ISNULL(WKH1,0) FROM BG_T_BUDGET_ADJUST WHERE BUDGET_YEAR=@BudgetYear AND PERIOD_TYPE=@PeriodType AND REV_NO= (SELECT MAX(REV_NO) AS REV_NO FROM BG_T_BUDGET_DATA WHERE BUDGET_YEAR = @BudgetYear AND PERIOD_TYPE = @PeriodType)) AS WB_ORIGINAL_1ST_HALF,
        '0 AS WB_ACTUAL_M1,
        '0 AS WB_ACTUAL_M2,
        '0 AS WB_ACTUAL_M3,
        'SUM(ISNULL(MIN_REV.M4,0) - ISNULL(ESTIMATE_BUDGET.M4, 0)) AS WB_ESTIMATE_M4,
        'SUM(ISNULL(MIN_REV.M5,0) - ISNULL(ESTIMATE_BUDGET.M5, 0)) AS WB_ESTIMATE_M5,
        'SUM(ISNULL(MIN_REV.M6,0) - ISNULL(ESTIMATE_BUDGET.M6, 0)) AS WB_ESTIMATE_M6,
        '0 AS WB_ESTIMATE_1ST_HALF,
        '0 AS WB_DIFF_1ST_HALF,
        '(SELECT ISNULL(WKH2,0) FROM BG_T_BUDGET_ADJUST WHERE BUDGET_YEAR=@BudgetYear AND PERIOD_TYPE=@PeriodType AND REV_NO= (SELECT MAX(REV_NO) AS REV_NO FROM BG_T_BUDGET_DATA WHERE BUDGET_YEAR = @BudgetYear AND PERIOD_TYPE = @PeriodType)) AS WB_ORIGINAL_2ND_HALF,
        'SUM(ISNULL(MIN_REV.M7,0) - ISNULL(MASTER_DATA.M7, 0)) AS WB_REVISE_M7,
        'SUM(ISNULL(MIN_REV.M8,0) - ISNULL(MASTER_DATA.M8, 0)) AS WB_REVISE_M8,
        'SUM(ISNULL(MIN_REV.M9,0) - ISNULL(MASTER_DATA.M9, 0)) AS WB_REVISE_M9,
        'SUM(ISNULL(MIN_REV.M10,0) - ISNULL(MASTER_DATA.M10, 0)) AS WB_REVISE_M10,
        'SUM(ISNULL(MIN_REV.M11,0) - ISNULL(MASTER_DATA.M11, 0)) AS WB_REVISE_M11,
        'SUM(ISNULL(MIN_REV.M12,0) - ISNULL(MASTER_DATA.M12, 0)) AS WB_REVISE_M12,
        '0 AS WB_REVISE_2ND_HALF,
        '0 AS WB_DIFF_2ND_HALF,
        '0 AS WB_REVISE_TOTAL_YEAR,
        '0 AS WB_DIFF_TOTAL_YEAR,
        '0 AS WB_RRT1,
        '0 AS WB_RRT2,
        '0 AS WB_RRT3,
        '0 AS WB_RRT4,
        '0 AS WB_RRT5

        Return True

    End Function

    Private Function InsertReviseMTPColumnData(ByRef dtColumns As DataTable, _
                                               ByVal strYear As String) As Boolean

        Dim strHalfYear As String = strYear.Substring(2, 2)
        Dim row As DataRow

        'SELECT MAX_REV.BUDGET_YEAR, MAX_REV.PERIOD_TYPE, MAX_REV.ACCOUNT_NO , 
        row = dtColumns.NewRow()
        row("Column_Name") = "ACCOUNT_NO"
        row("Column_Title") = ""
        dtColumns.Rows.Add(row)

        'MAX_REV.ACCOUNT_NAME , MAX_REV.BUDGET_TYPE , MAX_REV.EXPENSE_TYPE,
        row = dtColumns.NewRow()
        row("Column_Name") = "ACCOUNT_NAME"
        row("Column_Title") = ""
        dtColumns.Rows.Add(row)

        ''SUM(ISNULL(ORIGINAL_BUDGET.H1,0)) AS ORIGINAL_1ST_HALF,
        'row = dtColumns.NewRow()
        'row("Column_Name") = "ORIGINAL_1ST_HALF"
        'row("Column_Title") = "Original 1st Half'" & strHalfYear
        'dtColumns.Rows.Add(row)

        ''SUM(ISNULL(ACTUAL_DATA.M1, 0)) AS ACTUAL_M1,
        'row = dtColumns.NewRow()
        'row("Column_Name") = "ACTUAL_M1"
        'row("Column_Title") = "Jan'" & strHalfYear
        'dtColumns.Rows.Add(row)

        ''SUM(ISNULL(ACTUAL_DATA.M2, 0)) AS ACTUAL_M2,
        'row = dtColumns.NewRow()
        'row("Column_Name") = "ACTUAL_M2"
        'row("Column_Title") = "Feb'" & strHalfYear
        'dtColumns.Rows.Add(row)

        ''SUM(ISNULL(ACTUAL_DATA.M3, 0)) AS ACTUAL_M3,
        'row = dtColumns.NewRow()
        'row("Column_Name") = "ACTUAL_M3"
        'row("Column_Title") = "Mar'" & strHalfYear
        'dtColumns.Rows.Add(row)

        ''SUM(ISNULL(ESTIMATE_BUDGET.M4, 0)) AS ESTIMATE_M4,
        'row = dtColumns.NewRow()
        'row("Column_Name") = "ESTIMATE_M4"
        'row("Column_Title") = "Apr'" & strHalfYear
        'dtColumns.Rows.Add(row)

        ''SUM(ISNULL(ESTIMATE_BUDGET.M5, 0)) AS ESTIMATE_M5,
        'row = dtColumns.NewRow()
        'row("Column_Name") = "ESTIMATE_M5"
        'row("Column_Title") = "May'" & strHalfYear
        'dtColumns.Rows.Add(row)

        ''SUM(ISNULL(ESTIMATE_BUDGET.M6, 0)) AS ESTIMATE_M6,
        'row = dtColumns.NewRow()
        'row("Column_Name") = "ESTIMATE_M6"
        'row("Column_Title") = "Jun'" & strHalfYear
        'dtColumns.Rows.Add(row)

        ''SUM(ISNULL(ACTUAL_DATA.M1, 0) + ISNULL(ACTUAL_DATA.M2, 0) + ISNULL(ACTUAL_DATA.M3, 0) + ISNULL(ESTIMATE_BUDGET.M4, 0) + ISNULL(ESTIMATE_BUDGET.M5, 0) + ISNULL(ESTIMATE_BUDGET.M6, 0)) AS ESTIMATE_1ST_HALF,
        'row = dtColumns.NewRow()
        'row("Column_Name") = "ESTIMATE_1ST_HALF"
        'row("Column_Title") = "Estimate 1st Half'" & strHalfYear
        'dtColumns.Rows.Add(row)

        ''SUM((ISNULL(ACTUAL_DATA.M1, 0) + ISNULL(ACTUAL_DATA.M2, 0) + ISNULL(ACTUAL_DATA.M3, 0) + ISNULL(ESTIMATE_BUDGET.M4, 0) + ISNULL(ESTIMATE_BUDGET.M5, 0) + ISNULL(ESTIMATE_BUDGET.M6, 0)) - ISNULL(ORIGINAL_BUDGET.H1,0)) AS DIFF_1ST_HALF,
        'row = dtColumns.NewRow()
        'row("Column_Name") = "DIFF_1ST_HALF"
        'row("Column_Title") = "Diff 1st Half'" & strHalfYear
        'dtColumns.Rows.Add(row)

        'SUM(ISNULL(ORIGINAL_BUDGET.H2,0)) AS ORIGINAL_2ND_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "ORIGINAL_2ND_HALF"
        row("Column_Title") = "Original 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M7, 0)) AS REVISE_M7, 
        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_M7"
        row("Column_Title") = "Jul'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M8, 0)) AS REVISE_M8, 
        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_M8"
        row("Column_Title") = "Aug'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M9, 0)) AS REVISE_M9,
        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_M9"
        row("Column_Title") = "Sep'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M10, 0)) AS REVISE_M10, 
        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_M10"
        row("Column_Title") = "Oct'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M11, 0)) AS REVISE_M11, 
        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_M11"
        row("Column_Title") = "Nov'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M12, 0)) AS REVISE_M12,
        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_M12"
        row("Column_Title") = "Dec'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M7, 0) + ISNULL(MASTER_DATA.M8, 0) + ISNULL(MASTER_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11, 0) + ISNULL(MASTER_DATA.M12, 0)) AS REVISE_2ND_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_2ND_HALF"
        row("Column_Title") = "Revise 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(MASTER_DATA.M7, 0) + ISNULL(MASTER_DATA.M8, 0) + ISNULL(MASTER_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11, 0) + ISNULL(MASTER_DATA.M12, 0) - ISNULL(ORIGINAL_BUDGET.H2,0)) AS DIFF_2ND_HALF,
        row = dtColumns.NewRow()
        row("Column_Name") = "DIFF_2ND_HALF"
        row("Column_Title") = "Diff 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(row)

        'SUM(ISNULL(ACTUAL_DATA.M1, 0) + ISNULL(ACTUAL_DATA.M2, 0) + ISNULL(ACTUAL_DATA.M3, 0) + ISNULL(ESTIMATE_BUDGET.M4, 0) + ISNULL(ESTIMATE_BUDGET.M5, 0) + ISNULL(ESTIMATE_BUDGET.M6, 0) + ISNULL(MASTER_DATA.M7, 0) + ISNULL(MASTER_DATA.M8, 0) + ISNULL(MASTER_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11, 0) + ISNULL(MASTER_DATA.M12, 0)) AS REVISE_TOTAL_YEAR,
        row = dtColumns.NewRow()
        row("Column_Name") = "REVISE_TOTAL_YEAR"
        row("Column_Title") = "Revise Year'" & strYear
        dtColumns.Rows.Add(row)

        'SUM((ISNULL(ACTUAL_DATA.M1, 0) + ISNULL(ACTUAL_DATA.M2, 0) + ISNULL(ACTUAL_DATA.M3, 0) + ISNULL(ESTIMATE_BUDGET.M4, 0) + ISNULL(ESTIMATE_BUDGET.M5, 0) + ISNULL(ESTIMATE_BUDGET.M6, 0)) - ISNULL(ORIGINAL_BUDGET.H1,0) + ISNULL(MASTER_DATA.M7, 0) + ISNULL(MASTER_DATA.M8, 0) + ISNULL(MASTER_DATA.M9, 0) + ISNULL(MASTER_DATA.M10, 0) + ISNULL(MASTER_DATA.M11, 0) + ISNULL(MASTER_DATA.M12, 0) - ISNULL(ORIGINAL_BUDGET.H2,0) ) AS DIFF_TOTAL_YEAR,
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

        ''SUM(ISNULL(MASTER_DATA.RRT5 ,0)) AS RRT5,
        'row = dtColumns.NewRow()
        'row("Column_Name") = "RRT5"
        'row("Column_Title") = "Y" & CInt(strYear) + 5
        'dtColumns.Rows.Add(row)

        Return True

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
        row("Column_Title") = "Original Year'" & CInt(strYear) + 1
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow
        row("Column_Name") = "PrevRRT2"
        row("Column_Title") = "MTP" & CInt(strYear) - 1 & " Year'" & CInt(strYear) + 1
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "DiffYear"
        row("Column_Title") = "Diff Year'" & CInt(strYear) + 1
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow
        row("Column_Name") = "RRT2"
        row("Column_Title") = "MTP" & CInt(strYear) & " Year'" & CInt(strYear) + 2
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow
        row("Column_Name") = "PrevRRT3"
        row("Column_Title") = "MTP" & CInt(strYear) - 1 & " Year'" & CInt(strYear) + 2
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow()
        row("Column_Name") = "DIFFRRT1"
        row("Column_Title") = "Diff Year'" & CInt(strYear) + 2
        dtColumns.Rows.Add(row)

        row = dtColumns.NewRow
        row("Column_Name") = "RRT3"
        row("Column_Title") = "MTP" & CInt(strYear) & " Year'" & CInt(strYear) + 3
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
        Dim drManufacturingCost As DataRow = dtResult.NewRow
        Dim drAdministrationCost As DataRow = dtResult.NewRow
        Dim drTotalExpense As DataRow = dtResult.NewRow
        Dim drWorkingBudget As DataRow = dtResult.NewRow
        Dim drOutflowTotal As DataRow = dtResult.NewRow

        '//Calculate Investments
        CalculateReviseInvestments(dsData, intDataColumnIndex, blnMTPBudget, drInvestments)

        '//Calculate Manufacturing cost
        CalculateReviseManufacturingCost(dsData, intDataColumnIndex, blnMTPBudget, drManufacturingCost)

        '//Calculate Administration cost
        CalculateReviseAdministrationCost(dsData, intDataColumnIndex, blnMTPBudget, drAdministrationCost)

        '//Calculate Working Budget
        CalculateReviseWorkingBudget(dsData, intDataColumnIndex, blnMTPBudget, drWorkingBudget)

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
                strExpression = "Sum(" + strColumnName + ")"
                strFilter = strScript
                returnValue = dtResult.Compute(strExpression, strFilter)
                drTotal(dtResult.Columns(k).ColumnName) = returnValue

            Next
            '//Set Group header
            drTotal("ACCOUNT_NO") = GetGroupExpensesTitle(dtGroups.Rows(i)(0).ToString)

            '//Add total cost
            dtResult.Rows.InsertAt(drTotal, intGroupTotalIndex)

            '//Add one empty row
            drEmpty = dtResult.NewRow
            dtResult.Rows.Add(drEmpty)

            'dtResult.TableName = arrRows(0)(strGroupColumnName).ToString & " " & arrRows(0)(strGroupColumnTitle).ToString
            dtResult.TableName = "Revise Budget"

            'intGroupTotalIndex = intGroupTotalIndex + dtResult.Rows.Count
            intGroupTotalIndex = intGroupTotalIndex + CInt(arrRows.Length) + 2

        Next

        '//Set Data to Account No.
        SetAccountNoText(drInvestments, drManufacturingCost, drAdministrationCost, drTotalExpense, drWorkingBudget, drOutflowTotal)

        '//Calculate Total Expense
        CalculateTotalExpense(dtResult, intDataColumnIndex, drManufacturingCost, drAdministrationCost, drTotalExpense)

        '//Calculate Outflow Total 
        CalculateOutflowTotal(dtResult, intDataColumnIndex, drTotalExpense, drWorkingBudget, drInvestments, drOutflowTotal)

        '//Add Investments
        dtResult.Rows.InsertAt(drInvestments, 0)

        '//Add one empty row
        drEmpty = dtResult.NewRow
        dtResult.Rows.InsertAt(drEmpty, 1)

        '//Add Manufacturing cost
        dtResult.Rows.Add(drManufacturingCost)

        '//Add Administration cost
        dtResult.Rows.Add(drAdministrationCost)

        '//Add Total Expense
        dtResult.Rows.Add(drTotalExpense)

        '//Add one empty row
        drEmpty = dtResult.NewRow
        dtResult.Rows.Add(drEmpty)

        '//Add Working Budget
        dtResult.Rows.Add(drWorkingBudget)

        '//Add Outflow Total (Investment;Expenses)
        dtResult.Rows.Add(drOutflowTotal)

        '//Return data table
        dsResult.Tables.Add(dtResult)

        Return dsResult

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
        Dim drManufacturingCost As DataRow = dtResult.NewRow
        Dim drAdministrationCost As DataRow = dtResult.NewRow
        Dim drTotalExpense As DataRow = dtResult.NewRow
        Dim drWorkingBudget As DataRow = dtResult.NewRow
        Dim drOutflowTotal As DataRow = dtResult.NewRow

        '//Calculate Investments
        CalculateMTPInvestments(dsData, intDataColumnIndex, drInvestments)

        '//Calculate Manufacturing cost
        CalculateMTPManufacturingCost(dsData, intDataColumnIndex, drManufacturingCost)

        '//Calculate Administration cost
        CalculateMTPAdministrationCost(dsData, intDataColumnIndex, drAdministrationCost)

        '//Calculate Working Budget
        CalculateMTPWorkingBudget(dsData, intDataColumnIndex, drWorkingBudget)

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
                strExpression = "Sum(" + strColumnName + ")"
                strFilter = strScript
                returnValue = dtResult.Compute(strExpression, strFilter)
                drTotal(dtResult.Columns(k).ColumnName) = returnValue

            Next
            '//Set Group header
            drTotal("ACCOUNT_NO") = GetGroupExpensesTitle(dtGroups.Rows(i)(0).ToString)

            '//Add total cost
            dtResult.Rows.InsertAt(drTotal, intGroupTotalIndex)

            '//Add one empty row
            drEmpty = dtResult.NewRow
            dtResult.Rows.Add(drEmpty)

            'dtResult.TableName = arrRows(0)(strGroupColumnName).ToString & " " & arrRows(0)(strGroupColumnTitle).ToString
            dtResult.TableName = "MTP Budget"

            'intGroupTotalIndex = intGroupTotalIndex + dtResult.Rows.Count
            intGroupTotalIndex = intGroupTotalIndex + CInt(arrRows.Length) + 2

        Next

        '//Set Data to Account No.
        SetAccountNoText(drInvestments, drManufacturingCost, drAdministrationCost, drTotalExpense, drWorkingBudget, drOutflowTotal)

        '//Calculate Total Expense
        CalculateTotalExpense(dtResult, intDataColumnIndex, drManufacturingCost, drAdministrationCost, drTotalExpense)

        '//Calculate Outflow Total 
        CalculateOutflowTotal(dtResult, intDataColumnIndex, drTotalExpense, drWorkingBudget, drInvestments, drOutflowTotal)

        '//Add Investments
        dtResult.Rows.InsertAt(drInvestments, 0)

        '//Add one empty row
        drEmpty = dtResult.NewRow
        dtResult.Rows.InsertAt(drEmpty, 1)

        '//Add Manufacturing cost
        dtResult.Rows.Add(drManufacturingCost)

        '//Add Administration cost
        dtResult.Rows.Add(drAdministrationCost)

        '//Add Total Expense
        dtResult.Rows.Add(drTotalExpense)

        '//Add one empty row
        drEmpty = dtResult.NewRow
        dtResult.Rows.Add(drEmpty)

        '//Add Working Budget
        dtResult.Rows.Add(drWorkingBudget)

        '//Add Outflow Total (Investment;Expenses)
        dtResult.Rows.Add(drOutflowTotal)

        '//Return data table
        dsResult.Tables.Add(dtResult)

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

            '//Setup Revise & Estimate Title
            If blnMTPBudget = False Then

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

            '//Setup Investments Line
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).ClearContents()
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).MergeCells = True
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).Value = "Investments"
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, colMax)).Font.Bold = True '// Add by Max 27/09/2012

            '//Setup budget order name column to be left align
            'ws.Range(ws.Cells(rowStartIndex + 1, 1), ws.Cells(rowMax, 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            ws.Range(ws.Cells(rowStartIndex, 2), ws.Cells(rowMax, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup Manufacturing Cost Line
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, 2)).Value = "Manufacturing Cost"
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, colMax)).Font.Bold = True '// Add by Max 27/09/2012

            '//Setup Administration Cost Line
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, 2)).Value = "Administration Cost"
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, colMax)).Font.Bold = True '// Add by Max 27/09/2012

            '//Setup Total Expense Line
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).Value = "Total Expense"
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, colMax)).Font.Bold = True '// Add by Max 27/09/2012

            '//Empry line
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).Value = ""
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup Working Budget Line
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).Value = "Working Budget"
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

            '// MTP Budget
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
            '    intFontStart = 1
            '    intFontEnd = 28

            '    ''//Delete Column
            '    'rng = ws.Range(ws.Cells(colStartIndex - 1, 3), ws.Cells(rowMax, 11))
            '    'rng.EntireColumn.Delete(missing)
            '    'System.Runtime.InteropServices.Marshal.ReleaseComObject(rng)

            'Else
            '    intAuthorizeStart = 22
            '    intFontStart = 1
            '    intFontEnd = colMax
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

            '//Setup Investments Line
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).ClearContents()
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).MergeCells = True
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).Value = "Investments"
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            ws.Range(ws.Cells(rowStartIndex, 1), ws.Cells(rowStartIndex, colMax)).Font.Bold = True '// Add by Max 27/09/2012

            '//Setup budget order name column to be left align
            'ws.Range(ws.Cells(rowStartIndex + 1, 1), ws.Cells(rowMax, 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            ws.Range(ws.Cells(rowStartIndex, 2), ws.Cells(rowMax, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup Manufacturing Cost Line
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, 2)).Value = "Manufacturing Cost"
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            ws.Range(ws.Cells(rowMax - 5, 1), ws.Cells(rowMax - 5, colMax)).Font.Bold = True '// Add by Max 27/09/2012

            '//Setup Administration Cost Line
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, 2)).Value = "Administration Cost"
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            ws.Range(ws.Cells(rowMax - 4, 1), ws.Cells(rowMax - 4, colMax)).Font.Bold = True '// Add by Max 27/09/2012

            '//Setup Total Expense Line
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).Value = "Total Expense"
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            ws.Range(ws.Cells(rowMax - 3, 1), ws.Cells(rowMax - 3, colMax)).Font.Bold = True '// Add by Max 27/09/2012

            '//Empry line
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).Value = ""
            ws.Range(ws.Cells(rowMax - 2, 1), ws.Cells(rowMax - 2, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup Working Budget Line
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).Font.Bold = True
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).Value = "Working Budget"
            ws.Range(ws.Cells(rowMax - 1, 1), ws.Cells(rowMax - 1, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup Outflow Total Line
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).ClearContents()
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).MergeCells = True
            ws.Range(ws.Cells(rowMax, 1), ws.Cells(rowMax, 2)).Value = "Outflow Total"
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

            '// MTP Budget
            'If blnMTPBudget = True Then
            '// Set Header
            'ws.Cells(colStartIndex - 1, 15) = "MTP Budget"
            'ws.Range(ws.Cells(colStartIndex - 1, 7), ws.Cells(colStartIndex - 1, 15)).MergeCells = True
            'ws.Range(ws.Cells(colStartIndex - 1, 7), ws.Cells(colStartIndex - 1, 15)).Font.Bold = True
            'ws.Range(ws.Cells(colStartIndex - 1, 7), ws.Cells(colStartIndex - 1, 15)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            'Dim xColumn As Excel.Range = CType(ws.Columns(15, Type.Missing), Excel.Range)
            'xColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Type.Missing)
            'excelApp.Range(excelApp.Cells(colStartIndex, 15), excelApp.Cells(rowMax, 15)).Borders.LineStyle = 0
            'excelApp.Range(excelApp.Cells(colStartIndex, 15), excelApp.Cells(rowMax, 15)).ColumnWidth = 2
            'excelApp.Range(excelApp.Cells(colStartIndex, 15), excelApp.Cells(rowMax, 15)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
            'excelApp.Range(excelApp.Cells(colStartIndex, 15), excelApp.Cells(rowMax, 15)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
            'excelApp.Range(excelApp.Cells(colStartIndex, 15), excelApp.Cells(rowMax, 15)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = 0
            'excelApp.Range(excelApp.Cells(colStartIndex, 15), excelApp.Cells(rowMax, 15)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = 0
            'excelApp.Range(excelApp.Cells(colStartIndex - 1, 15), excelApp.Cells(rowMax, 15)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = 0

            intAuthorizeStart = 9
            intFontStart = 1
            intFontEnd = 14

            ''//Delete Column
            'rng = ws.Range(ws.Cells(colStartIndex - 1, 3), ws.Cells(rowMax, 11))
            'rng.EntireColumn.Delete(missing)
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(rng)

            'Else
            'intAuthorizeStart = 22
            'intFontStart = 1
            'intFontEnd = colMax
            'End If

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
            ws.Range(ws.Cells(colStartIndex, 6), ws.Cells(rowMax, 7)).NumberFormat = "#,##0.00;[Red]-#,##0.00"
            ws.Range(ws.Cells(colStartIndex, 9), ws.Cells(rowMax, colMax)).NumberFormat = "#,##0.00;[Red]-#,##0.00"

            '//Set Frame  
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders.LineStyle = 1
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium

            ws.Range(ws.Cells(colStartIndex, 3), ws.Cells(rowMax, 4)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 5), ws.Cells(rowMax, 5)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            ws.Range(ws.Cells(colStartIndex, 9), ws.Cells(rowMax, 9)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium

            '//Set font color
            ws.Range(ws.Cells(colStartIndex, 4), ws.Cells(rowMax, 5)).Font.Color = RGB(128, 128, 128)
            ws.Range(ws.Cells(colStartIndex, 7), ws.Cells(rowMax, 7)).Font.Color = RGB(128, 128, 128)
            ws.Range(ws.Cells(colStartIndex, 8), ws.Cells(rowMax, 8)).Font.Color = RGB(128, 128, 128)
            ws.Range(ws.Cells(colStartIndex, 11), ws.Cells(rowMax, 11)).Font.Color = RGB(128, 128, 128)
            ws.Range(ws.Cells(colStartIndex, 13), ws.Cells(rowMax, 13)).Font.Color = RGB(128, 128, 128)


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

    Private Function SetupExcelTitle(ByVal ws As Excel.Worksheet, ByVal intUnitPriceStart As Integer) As Boolean
        Dim strSubTitle As String

        If Me.numProjectNo.Value.ToString <> "1" Then
            strSubTitle = "Summary by Account No : " + Me.cboPeriodType.Text + " " + Me.numYear.Value.ToString + " (Project No." + Me.numProjectNo.Value.ToString + ")"
        Else
            strSubTitle = "Summary by Account No : " + Me.cboPeriodType.Text + " " + Me.numYear.Value.ToString
        End If


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

        ws.Range(ws.Cells(1, 1), ws.Cells(2, 4)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft '// Add by Max 27/09/2012

        '//Setup unit price
        ws.Range(excelApp.Cells(4, intUnitPriceStart), excelApp.Cells(4, intUnitPriceStart)).Font.Name = "Tahoma"
        ws.Range(excelApp.Cells(4, intUnitPriceStart), excelApp.Cells(4, intUnitPriceStart)).Font.Bold = True
        ws.Range(excelApp.Cells(4, intUnitPriceStart), excelApp.Cells(4, intUnitPriceStart)).Font.Underline = True
        ws.Range(excelApp.Cells(4, intUnitPriceStart), excelApp.Cells(4, intUnitPriceStart)).Font.Size = 11
        ws.Range(excelApp.Cells(4, intUnitPriceStart), excelApp.Cells(4, intUnitPriceStart)).MergeCells = True
        ws.Range(excelApp.Cells(4, intUnitPriceStart), excelApp.Cells(4, intUnitPriceStart)).WrapText = False
        ws.Range(excelApp.Cells(4, intUnitPriceStart), excelApp.Cells(4, intUnitPriceStart)).Value = "Unit : K.Baht"
        ws.Range(ws.Cells(4, intUnitPriceStart), ws.Cells(4, intUnitPriceStart)).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

    End Function

    Private Function CalculateTotalExpense(ByVal dt As DataTable, _
                                            ByVal intDataColumnIndex As Integer, _
                                            ByVal drManufacturingCost As DataRow, _
                                            ByVal drAdministrationCost As DataRow, _
                                            ByRef drTotalExpense As DataRow) As Boolean

        For k As Integer = intDataColumnIndex To dt.Columns.Count - 1
            drTotalExpense(dt.Columns(k).ColumnName) = Convert.ToDecimal(Nz(drManufacturingCost(dt.Columns(k).ColumnName), 0.0)) + Convert.ToDecimal(Nz(drAdministrationCost(dt.Columns(k).ColumnName), 0.0))
        Next

        Return True
    End Function

    Private Function CalculateOutflowTotal(ByVal dt As DataTable, _
                                          ByVal intDataColumnIndex As Integer, _
                                          ByVal drTotalExpense As DataRow, _
                                          ByVal drWorkingBudget As DataRow, _
                                          ByVal drInvestments As DataRow, _
                                          ByRef drOutflowTotal As DataRow) As Boolean

        For k As Integer = intDataColumnIndex To dt.Columns.Count - 1
            drOutflowTotal(dt.Columns(k).ColumnName) = Convert.ToDecimal(Nz(drTotalExpense(dt.Columns(k).ColumnName), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget(dt.Columns(k).ColumnName), 0.0)) + Convert.ToDecimal(Nz(drInvestments(dt.Columns(k).ColumnName), 0.0))
        Next

        Return True
    End Function

    Private Function CalculateReviseInvestments(ByVal dsData As DataSet, _
                                          ByVal intDataColumnIndex As Integer, _
                                          ByVal blnMTPBudget As Boolean, _
                                          ByRef drInvestments As DataRow) As Boolean
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object

        Dim strExpr As String = "BUDGET_TYPE = 'A'"
        Dim strSort As String = String.Empty '"ACCOUNT_NO ASC"

        '//Sort groups list by group column name
        Dim dtTmp As DataTable = dsData.Tables(0).Clone
        Dim arrTmp As DataRow() = dsData.Tables(0).Select(strExpr, strSort)
        For intTmp As Integer = 0 To arrTmp.Length - 1
            Dim drow(dtTmp.Columns.Count - 1) As Object
            arrTmp(intTmp).ItemArray.CopyTo(drow, 0)
            dtTmp.Rows.Add(drow)
        Next

        For k As Integer = intDataColumnIndex To dtTmp.Columns.Count - 1
            'For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dtTmp.Columns(k).ColumnName

            strExpression = "Sum(" + strColumnName + ")"
            'strFilter = "BUDGET_TYPE = 'A'"
            returnValue = dtTmp.Compute(strExpression, strFilter)
            drInvestments(dtTmp.Columns(k).ColumnName) = returnValue

            Select Case dtTmp.Columns(k).ColumnName
                Case "INVESTMENT_ORIGINAL_1ST_HALF"
                    drInvestments("ORIGINAL_1ST_HALF") = returnValue
                Case "INVESTMENT_ACTUAL_JAN"
                    drInvestments("ACTUAL_M1") = returnValue
                Case "INVESTMENT_ACTUAL_FEB"
                    drInvestments("ACTUAL_M2") = returnValue
                Case "INVESTMENT_ACTUAL_MAR"
                    drInvestments("ACTUAL_M3") = returnValue
                Case "INVESTMENT_ESTIMATE_APR"
                    drInvestments("ESTIMATE_M4") = returnValue
                Case "INVESTMENT_ESTIMATE_MAY"
                    drInvestments("ESTIMATE_M5") = returnValue
                Case "INVESTMENT_ESTIMATE_JUN"
                    drInvestments("ESTIMATE_M6") = returnValue
                Case "INVESTMENT_ORIGINAL_2ND_HALF"
                    drInvestments("ORIGINAL_2ND_HALF") = returnValue
                Case "INVESTMENT_REVISE_JUL"
                    drInvestments("REVISE_M7") = returnValue
                Case "INVESTMENT_REVISE_AUG"
                    drInvestments("REVISE_M8") = returnValue
                Case "INVESTMENT_REVISE_SEP"
                    drInvestments("REVISE_M9") = returnValue
                Case "INVESTMENT_REVISE_OCT"
                    drInvestments("REVISE_M10") = returnValue
                Case "INVESTMENT_REVISE_NOV"
                    drInvestments("REVISE_M11") = returnValue
                Case "INVESTMENT_REVISE_DEC"

                    drInvestments("REVISE_M12") = returnValue

                    '{@INVEST_ACTUAL_JAN} + {@INVEST_ACTUAL_FEB} + {@INVEST_ACTUAL_MAR} + {@INVEST_ESTIMATE_APR} + {@INVEST_ESTIMATE_MAY} + {@INVEST_ESTIMATE_JUN}
                    drInvestments("ESTIMATE_1ST_HALF") = Convert.ToDecimal(Nz(drInvestments("ACTUAL_M1"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("ACTUAL_M2"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("ACTUAL_M3"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("ESTIMATE_M4"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("ESTIMATE_M5"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("ESTIMATE_M6"), 0.0))

                    '{@INVEST_ESTIMATE_1ST_HALF} - {@INVEST_ORIGINAL_1ST_HALF}
                    drInvestments("DIFF_1ST_HALF") = Convert.ToDecimal(Nz(drInvestments("ESTIMATE_1ST_HALF"), 0.0)) - Convert.ToDecimal(Nz(drInvestments("ORIGINAL_1ST_HALF"), 0.0))

                    'Sum ({ReviseSummaryByAccountCode.INVESTMENT_REVISE_JUL})+Sum ({ReviseSummaryByAccountCode.INVESTMENT_REVISE_AUG})+Sum ({ReviseSummaryByAccountCode.INVESTMENT_REVISE_SEP})+Sum ({ReviseSummaryByAccountCode.INVESTMENT_REVISE_OCT})+Sum ({ReviseSummaryByAccountCode.INVESTMENT_REVISE_NOV})+Sum ({ReviseSummaryByAccountCode.INVESTMENT_REVISE_DEC})
                    drInvestments("REVISE_2ND_HALF") = Convert.ToDecimal(Nz(drInvestments("REVISE_M7"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("REVISE_M8"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("REVISE_M9"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("REVISE_M10"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("REVISE_M11"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("REVISE_M12"), 0.0))

                    '{@INVEST_REVISE_2ND_HALF} - {@INVEST_ORIGINAL_2ND_HALF}
                    drInvestments("DIFF_2ND_HALF") = Convert.ToDecimal(Nz(drInvestments("REVISE_2ND_HALF"), 0.0)) - Convert.ToDecimal(Nz(drInvestments("ORIGINAL_2ND_HALF"), 0.0))

                    '{@INVEST_ESTIMATE_1ST_HALF} + {@INVEST_REVISE_2ND_HALF}
                    'drInvestments("REVISE_TOTAL_YEAR") = Convert.ToDecimal(drInvestments("ACTUAL_M1")) + Convert.ToDecimal(drInvestments("ACTUAL_M2")) + Convert.ToDecimal(drInvestments("ACTUAL_M3")) + Convert.ToDecimal(drInvestments("ESTIMATE_M4")) + Convert.ToDecimal(drInvestments("ESTIMATE_M5")) + Convert.ToDecimal(drInvestments("ESTIMATE_M6")) + Convert.ToDecimal(drInvestments("REVISE_M7")) + Convert.ToDecimal(drInvestments("REVISE_M8")) + Convert.ToDecimal(drInvestments("REVISE_M9")) + Convert.ToDecimal(drInvestments("REVISE_M10")) + Convert.ToDecimal(drInvestments("REVISE_M11")) + Convert.ToDecimal(drInvestments("REVISE_M12"))
                    drInvestments("REVISE_TOTAL_YEAR") = Convert.ToDecimal(Nz(drInvestments("ESTIMATE_1ST_HALF"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("REVISE_2ND_HALF"), 0.0))

                    '{@INVEST_DIFF_1ST_HALF} + {@INVEST_DIFF_2ND_HALF}
                    'drInvestments("DIFF_TOTAL_YEAR") = Convert.ToDecimal(drInvestments("ACTUAL_M1")) + Convert.ToDecimal(drInvestments("ACTUAL_M2")) + Convert.ToDecimal(drInvestments("ACTUAL_M3")) + Convert.ToDecimal(drInvestments("ESTIMATE_M4")) + Convert.ToDecimal(drInvestments("ESTIMATE_M5")) + Convert.ToDecimal(drInvestments("ESTIMATE_M6")) + Convert.ToDecimal(drInvestments("REVISE_M7")) + Convert.ToDecimal(drInvestments("REVISE_M8")) + Convert.ToDecimal(drInvestments("REVISE_M9")) + Convert.ToDecimal(drInvestments("REVISE_M10")) + Convert.ToDecimal(drInvestments("REVISE_M11")) + Convert.ToDecimal(drInvestments("REVISE_M12")) - Convert.ToDecimal(drInvestments("ORIGINAL_2ND_HALF"))
                    drInvestments("DIFF_TOTAL_YEAR") = Convert.ToDecimal(Nz(drInvestments("DIFF_1ST_HALF"), 0.0)) + Convert.ToDecimal(Nz(drInvestments("DIFF_2ND_HALF"), 0.0))

            End Select

            '// MTP Budget
            If blnMTPBudget = True Then
                Select Case dtTmp.Columns(k).ColumnName
                    Case "INVESTMENT_RRT1"
                        'drInvestments("RRT1") = returnValue
                        drInvestments("RRT1") = Convert.ToDecimal(Nz(dtTmp.Rows(0)!INVESTMENT_RRT1, 0.0))
                    Case "INVESTMENT_RRT2"
                        'drInvestments("RRT2") = returnValue
                        drInvestments("RRT2") = Convert.ToDecimal(Nz(dtTmp.Rows(0)!INVESTMENT_RRT2, 0.0))
                    Case "INVESTMENT_RRT3"
                        'drInvestments("RRT3") = returnValue
                        drInvestments("RRT3") = Convert.ToDecimal(Nz(dtTmp.Rows(0)!INVESTMENT_RRT3, 0.0))
                    Case "INVESTMENT_RRT4"
                        'drInvestments("RRT4") = returnValue
                        drInvestments("RRT4") = Convert.ToDecimal(Nz(dtTmp.Rows(0)!INVESTMENT_RRT4, 0.0))
                    Case "INVESTMENT_RRT5"
                        'drInvestments("RRT5") = returnValue
                        drInvestments("RRT5") = Convert.ToDecimal(Nz(dtTmp.Rows(0)!INVESTMENT_RRT5, 0.0))
                End Select
            End If
        Next

        Return True
    End Function

    Private Function CalculateReviseManufacturingCost(ByVal dsData As DataSet, _
                                                ByVal intDataColumnIndex As Integer, _
                                                ByVal blnMTPBudget As Boolean, _
                                                ByRef drManufacturingCost As DataRow) As Boolean
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object

        For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName

            strExpression = "Sum(" + strColumnName + ")"
            strFilter = "BUDGET_TYPE = 'E'"
            returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
            drManufacturingCost(dsData.Tables(0).Columns(k).ColumnName) = returnValue

            Select Case dsData.Tables(0).Columns(k).ColumnName
                Case "FC_ORIGINAL_1ST_HALF"
                    drManufacturingCost("ORIGINAL_1ST_HALF") = returnValue
                Case "FC_ACTUAL_M1"
                    drManufacturingCost("ACTUAL_M1") = returnValue
                Case "FC_ACTUAL_M2"
                    drManufacturingCost("ACTUAL_M2") = returnValue
                Case "FC_ACTUAL_M3"
                    drManufacturingCost("ACTUAL_M3") = returnValue
                Case "FC_ESTIMATE_M4"
                    drManufacturingCost("ESTIMATE_M4") = returnValue
                Case "FC_ESTIMATE_M5"
                    drManufacturingCost("ESTIMATE_M5") = returnValue
                Case "FC_ESTIMATE_M6"
                    drManufacturingCost("ESTIMATE_M6") = returnValue
                Case "FC_ESTIMATE_1ST_HALF"
                    drManufacturingCost("ESTIMATE_1ST_HALF") = returnValue
                Case "FC_DIFF_1ST_HALF"
                    drManufacturingCost("DIFF_1ST_HALF") = returnValue
                Case "FC_ORIGINAL_2ND_HALF"
                    drManufacturingCost("ORIGINAL_2ND_HALF") = returnValue
                Case "FC_REVISE_M7"
                    drManufacturingCost("REVISE_M7") = returnValue
                Case "FC_REVISE_M8"
                    drManufacturingCost("REVISE_M8") = returnValue
                Case "FC_REVISE_M9"
                    drManufacturingCost("REVISE_M9") = returnValue
                Case "FC_REVISE_M10"
                    drManufacturingCost("REVISE_M10") = returnValue
                Case "FC_REVISE_M11"
                    drManufacturingCost("REVISE_M11") = returnValue
                Case "FC_REVISE_M12"
                    drManufacturingCost("REVISE_M12") = returnValue
                Case "FC_REVISE_2ND_HALF"
                    drManufacturingCost("REVISE_2ND_HALF") = returnValue
                Case "FC_DIFF_2ND_HALF"
                    drManufacturingCost("DIFF_2ND_HALF") = returnValue
                Case "FC_REVISE_TOTAL_YEAR"
                    drManufacturingCost("REVISE_TOTAL_YEAR") = returnValue
                Case "FC_DIFF_TOTAL_YEAR"
                    drManufacturingCost("DIFF_TOTAL_YEAR") = returnValue
            End Select

            '// MTP Budget
            If blnMTPBudget = True Then
                Select Case dsData.Tables(0).Columns(k).ColumnName
                    Case "FC_RRT1"
                        drManufacturingCost("RRT1") = returnValue
                    Case "FC_RRT2"
                        drManufacturingCost("RRT2") = returnValue
                    Case "FC_RRT3"
                        drManufacturingCost("RRT3") = returnValue
                    Case "FC_RRT4"
                        drManufacturingCost("RRT4") = returnValue
                    Case "FC_RRT5"
                        drManufacturingCost("RRT5") = returnValue
                End Select
            End If
        Next

        Return True
    End Function

    Private Function CalculateReviseAdministrationCost(ByVal dsData As DataSet, _
                                      ByVal intDataColumnIndex As Integer, _
                                       ByVal blnMTPBudget As Boolean, _
                                      ByRef drAdministrationCost As DataRow) As Boolean
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object

        For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName

            strExpression = "Sum(" + strColumnName + ")"
            strFilter = "BUDGET_TYPE = 'E'"
            returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
            drAdministrationCost(dsData.Tables(0).Columns(k).ColumnName) = returnValue

            Select Case dsData.Tables(0).Columns(k).ColumnName
                Case "ADMIN_ORIGINAL_1ST_HALF"
                    drAdministrationCost("ORIGINAL_1ST_HALF") = returnValue
                Case "ADMIN_ACTUAL_M1"
                    drAdministrationCost("ACTUAL_M1") = returnValue
                Case "ADMIN_ACTUAL_M2"
                    drAdministrationCost("ACTUAL_M2") = returnValue
                Case "ADMIN_ACTUAL_M3"
                    drAdministrationCost("ACTUAL_M3") = returnValue
                Case "ADMIN_ESTIMATE_M4"
                    drAdministrationCost("ESTIMATE_M4") = returnValue
                Case "ADMIN_ESTIMATE_M5"
                    drAdministrationCost("ESTIMATE_M5") = returnValue
                Case "ADMIN_ESTIMATE_M6"
                    drAdministrationCost("ESTIMATE_M6") = returnValue
                Case "ADMIN_ESTIMATE_1ST_HALF"
                    drAdministrationCost("ESTIMATE_1ST_HALF") = returnValue
                Case "ADMIN_DIFF_1ST_HALF"
                    drAdministrationCost("DIFF_1ST_HALF") = returnValue
                Case "ADMIN_ORIGINAL_2ND_HALF"
                    drAdministrationCost("ORIGINAL_2ND_HALF") = returnValue
                Case "ADMIN_REVISE_M7"
                    drAdministrationCost("REVISE_M7") = returnValue
                Case "ADMIN_REVISE_M8"
                    drAdministrationCost("REVISE_M8") = returnValue
                Case "ADMIN_REVISE_M9"
                    drAdministrationCost("REVISE_M9") = returnValue
                Case "ADMIN_REVISE_M10"
                    drAdministrationCost("REVISE_M10") = returnValue
                Case "ADMIN_REVISE_M11"
                    drAdministrationCost("REVISE_M11") = returnValue
                Case "ADMIN_REVISE_M12"
                    drAdministrationCost("REVISE_M12") = returnValue
                Case "ADMIN_REVISE_2ND_HALF"
                    drAdministrationCost("REVISE_2ND_HALF") = returnValue
                Case "ADMIN_DIFF_2ND_HALF"
                    drAdministrationCost("DIFF_2ND_HALF") = returnValue
                Case "ADMIN_REVISE_TOTAL_YEAR"
                    drAdministrationCost("REVISE_TOTAL_YEAR") = returnValue
                Case "ADMIN_DIFF_TOTAL_YEAR"
                    drAdministrationCost("DIFF_TOTAL_YEAR") = returnValue
            End Select

            '// MTP Budget
            If blnMTPBudget = True Then
                Select Case dsData.Tables(0).Columns(k).ColumnName
                    Case "ADMIN_RRT1"
                        drAdministrationCost("RRT1") = returnValue
                    Case "ADMIN_RRT2"
                        drAdministrationCost("RRT2") = returnValue
                    Case "ADMIN_RRT3"
                        drAdministrationCost("RRT3") = returnValue
                    Case "ADMIN_RRT4"
                        drAdministrationCost("RRT4") = returnValue
                    Case "ADMIN_RRT5"
                        drAdministrationCost("RRT5") = returnValue
                End Select
            End If
        Next

        Return True
    End Function

    Private Function CalculateReviseWorkingBudget(ByVal dsData As DataSet, _
                                            ByVal intDataColumnIndex As Integer, _
                                            ByVal blnMTPBudget As Boolean, _
                                            ByRef drWorkingBudget As DataRow) As Boolean
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object

        For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName

            strExpression = "Sum(" + strColumnName + ")"
            strFilter = "BUDGET_TYPE = 'E'"
            returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
            drWorkingBudget(dsData.Tables(0).Columns(k).ColumnName) = returnValue

            'If strColumnName = "WB_ORIGINAL_2ND_HALF" Then
            '    Debug.Print("WB_ORIGINAL_2ND_HALF")
            'End If

            'If strColumnName = "WB_REVISE_2ND_HALF" Then
            '    Debug.Print("WB_REVISE_2ND_HALF")
            'End If
            Select Case dsData.Tables(0).Columns(k).ColumnName
                Case "WB_ORIGINAL_1ST_HALF"
                    'drWorkingBudget("ORIGINAL_1ST_HALF") = returnValue                    
                    drWorkingBudget("ORIGINAL_1ST_HALF") = Convert.ToDecimal(Nz(dsData.Tables(0).Rows(0)![WB_ORIGINAL_1ST_HALF], 0.0))
                Case "WB_ACTUAL_M1"
                    drWorkingBudget("ACTUAL_M1") = returnValue
                Case "WB_ACTUAL_M2"
                    drWorkingBudget("ACTUAL_M2") = returnValue
                Case "WB_ACTUAL_M3"
                    drWorkingBudget("ACTUAL_M3") = returnValue
                Case "WB_ESTIMATE_M4"
                    drWorkingBudget("ESTIMATE_M4") = returnValue
                Case "WB_ESTIMATE_M5"
                    drWorkingBudget("ESTIMATE_M5") = returnValue
                Case "WB_ESTIMATE_M6"
                    drWorkingBudget("ESTIMATE_M6") = returnValue
                Case "WB_ESTIMATE_1ST_HALF"
                    'drWorkingBudget("ESTIMATE_1ST_HALF") = returnValue
                    drWorkingBudget("ESTIMATE_1ST_HALF") = Convert.ToDecimal(Nz(drWorkingBudget("ACTUAL_M1"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("ACTUAL_M2"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("ACTUAL_M3"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("ESTIMATE_M4"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("ESTIMATE_M5"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("ESTIMATE_M6"), 0.0))
                Case "WB_DIFF_1ST_HALF"
                    'drWorkingBudget("DIFF_1ST_HALF") = returnValue
                    drWorkingBudget("DIFF_1ST_HALF") = Convert.ToDecimal(Nz(drWorkingBudget("ACTUAL_M1"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("ACTUAL_M2"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("ACTUAL_M3"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("ESTIMATE_M4"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("ESTIMATE_M5"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("ESTIMATE_M6"), 0.0)) - Convert.ToDecimal(Nz(drWorkingBudget("ORIGINAL_1ST_HALF"), 0.0))
                Case "WB_ORIGINAL_2ND_HALF"
                    '  '{ReviseSummaryByAccountCode.WB_ORIGINAL_2ND_HALF}
                    'drWorkingBudget("ORIGINAL_2ND_HALF") = returnValue
                    drWorkingBudget("ORIGINAL_2ND_HALF") = Convert.ToDecimal(Nz(dsData.Tables(0).Rows(0)![WB_ORIGINAL_2ND_HALF], 0.0))
                Case "WB_REVISE_M7"
                    drWorkingBudget("REVISE_M7") = returnValue
                Case "WB_REVISE_M8"
                    drWorkingBudget("REVISE_M8") = returnValue
                Case "WB_REVISE_M9"
                    drWorkingBudget("REVISE_M9") = returnValue
                Case "WB_REVISE_M10"
                    drWorkingBudget("REVISE_M10") = returnValue
                Case "WB_REVISE_M11"
                    drWorkingBudget("REVISE_M11") = returnValue
                Case "WB_REVISE_M12"
                    drWorkingBudget("REVISE_M12") = returnValue
                Case "WB_REVISE_2ND_HALF"
                    '{@WBReviseM7} + {@WBReviseM8} + {@WBReviseM9} + {@WBReviseM10} + {@WBReviseM11} + {@WBReviseM12}
                    'drWorkingBudget("REVISE_2ND_HALF") = returnValue
                    drWorkingBudget("REVISE_2ND_HALF") = Convert.ToDecimal(Nz(drWorkingBudget("WB_REVISE_M7"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("WB_REVISE_M8"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("WB_REVISE_M9"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("REVISE_M10"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("WB_REVISE_M11"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("WB_REVISE_M12"), 0.0))
                Case "WB_DIFF_2ND_HALF"
                    '{@WBRevise2ndHalf} - {@WBOriginal2ndHalf}
                    'drWorkingBudget("DIFF_2ND_HALF") = returnValue
                    drWorkingBudget("DIFF_2ND_HALF") = Convert.ToDecimal(Nz(drWorkingBudget("REVISE_2ND_HALF"), 0.0)) - Convert.ToDecimal(Nz(drWorkingBudget("ORIGINAL_2ND_HALF"), 0.0))
                Case "WB_REVISE_TOTAL_YEAR"
                    '{@WBEstimate1stHalf} + {@WBRevise2ndHalf}
                    'drWorkingBudget("REVISE_TOTAL_YEAR") = returnValue
                    'drWorkingBudget("REVISE_TOTAL_YEAR") = Convert.ToDecimal(drWorkingBudget("ACTUAL_M1")) + Convert.ToDecimal(drWorkingBudget("ACTUAL_M2")) + Convert.ToDecimal(drWorkingBudget("ACTUAL_M3")) + Convert.ToDecimal(drWorkingBudget("ESTIMATE_M4")) + Convert.ToDecimal(drWorkingBudget("ESTIMATE_M5")) + Convert.ToDecimal(drWorkingBudget("ESTIMATE_M6")) + Convert.ToDecimal(drWorkingBudget("REVISE_M7")) + Convert.ToDecimal(drWorkingBudget("REVISE_M8")) + Convert.ToDecimal(drWorkingBudget("REVISE_M9")) + Convert.ToDecimal(drWorkingBudget("REVISE_M9")) + Convert.ToDecimal(drWorkingBudget("REVISE_M10")) + Convert.ToDecimal(drWorkingBudget("REVISE_M11")) + Convert.ToDecimal(drWorkingBudget("REVISE_M12"))
                    drWorkingBudget("REVISE_TOTAL_YEAR") = Convert.ToDecimal(Nz(drWorkingBudget("ESTIMATE_1ST_HALF"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("REVISE_2ND_HALF"), 0.0))
                Case "WB_DIFF_TOTAL_YEAR"
                    '{@WBDiff1stHalf} + {@WBDiff2ndHalf}
                    'drWorkingBudget("DIFF_TOTAL_YEAR") = returnValue
                    'drWorkingBudget("DIFF_TOTAL_YEAR") = Convert.ToDecimal(drWorkingBudget("REVISE_TOTAL_YEAR")) - Convert.ToDecimal(drWorkingBudget("ORIGINAL_2ND_HALF"))
                    drWorkingBudget("DIFF_TOTAL_YEAR") = Convert.ToDecimal(Nz(drWorkingBudget("DIFF_1ST_HALF"), 0.0)) + Convert.ToDecimal(Nz(drWorkingBudget("DIFF_2ND_HALF"), 0.0))
            End Select

            '//MTP Budget
            If blnMTPBudget = True Then
                Select Case dsData.Tables(0).Columns(k).ColumnName
                    Case "WB_RRT1"
                        drWorkingBudget("RRT1") = returnValue
                    Case "WB_RRT2"
                        drWorkingBudget("RRT2") = returnValue
                    Case "WB_RRT3"
                        drWorkingBudget("RRT3") = returnValue
                    Case "WB_RRT4"
                        drWorkingBudget("RRT4") = returnValue
                    Case "WB_RRT5"
                        drWorkingBudget("RRT5") = returnValue
                End Select
            End If
        Next

        Return True
    End Function

    Private Function CalculateMTPAdministrationCost(ByVal dsData As DataSet, _
                                      ByVal intDataColumnIndex As Integer, _
                                      ByRef drAdministrationCost As DataRow) As Boolean
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object

        For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName

            strExpression = "Sum(" + strColumnName + ")"
            strFilter = "BUDGET_TYPE = 'E'"
            returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
            drAdministrationCost(dsData.Tables(0).Columns(k).ColumnName) = returnValue


            Select Case dsData.Tables(0).Columns(k).ColumnName
                Case "ADMIN_RRT1"
                    drAdministrationCost("RRT1") = returnValue
                Case "ADMIN_RRT2"
                    drAdministrationCost("RRT2") = returnValue
                Case "ADMIN_RRT3"
                    drAdministrationCost("RRT3") = returnValue
                Case "ADMIN_RRT4"
                    drAdministrationCost("RRT4") = returnValue
                Case "ADMIN_RRT5"
                    drAdministrationCost("RRT5") = returnValue
                Case "ADMIN_REVYEAR"
                    drAdministrationCost("REVYEAR") = returnValue
                Case "ADMIN_PrevRRT1"
                    drAdministrationCost("PrevRRT1") = returnValue
                Case "ADMIN_PrevRRT2"
                    drAdministrationCost("PrevRRT2") = returnValue
                Case "ADMIN_PrevRRT3"
                    drAdministrationCost("PrevRRT3") = returnValue
                Case "ADMIN_PrevRRT4"
                    drAdministrationCost("PrevRRT4") = returnValue
                Case "ADMIN_PrevRRT5"
                    drAdministrationCost("PrevRRT5") = returnValue
                Case "ADMIN_DIFFYEAR"
                    drAdministrationCost("DiffYear") = returnValue
                Case "ADMIN_DIFFRRT1"
                    drAdministrationCost("DIFFRRT1") = returnValue



            End Select

        Next

        Return True
    End Function

    Private Function CalculateMTPManufacturingCost(ByVal dsData As DataSet, _
                                          ByVal intDataColumnIndex As Integer, _
                                          ByRef drManufacturingCost As DataRow) As Boolean
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object

        For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName

            strExpression = "Sum(" + strColumnName + ")"
            strFilter = "BUDGET_TYPE = 'E'"
            returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
            drManufacturingCost(dsData.Tables(0).Columns(k).ColumnName) = returnValue


            Select Case dsData.Tables(0).Columns(k).ColumnName
                Case "FC_RRT1"
                    drManufacturingCost("RRT1") = returnValue
                Case "FC_RRT2"
                    drManufacturingCost("RRT2") = returnValue
                Case "FC_RRT3"
                    drManufacturingCost("RRT3") = returnValue
                Case "FC_RRT4"
                    drManufacturingCost("RRT4") = returnValue
                Case "FC_RRT5"
                    drManufacturingCost("RRT5") = returnValue
                Case "FC_REVYEAR"
                    drManufacturingCost("REVYEAR") = returnValue
                Case "FC_PrevRRT1"
                    drManufacturingCost("PrevRRT1") = returnValue
                Case "FC_PrevRRT2"
                    drManufacturingCost("PrevRRT2") = returnValue
                Case "FC_PrevRRT3"
                    drManufacturingCost("PrevRRT3") = returnValue
                Case "FC_PrevRRT4"
                    drManufacturingCost("PrevRRT4") = returnValue
                Case "FC_PrevRRT5"
                    drManufacturingCost("PrevRRT5") = returnValue
                Case "FC_DIFFYEAR"
                    drManufacturingCost("DiffYear") = returnValue
                Case "FC_DIFFRRT1"
                    drManufacturingCost("DIFFRRT1") = returnValue

            End Select

        Next

        Return True
    End Function

    Private Function CalculateMTPWorkingBudget(ByVal dsData As DataSet, _
                                            ByVal intDataColumnIndex As Integer, _
                                            ByRef drWorkingBudget As DataRow) As Boolean
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object
        Dim strExpression2 As String

        For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName

            strExpression = "Sum(" + strColumnName + ")"
            strExpression2 = "Max(" + strColumnName + ")"

            strFilter = "BUDGET_TYPE = 'E'"
            returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
            drWorkingBudget(dsData.Tables(0).Columns(k).ColumnName) = returnValue

            Select Case dsData.Tables(0).Columns(k).ColumnName
                Case "WB_RRT1"
                    drWorkingBudget("RRT1") = returnValue
                Case "WB_RRT2"
                    drWorkingBudget("RRT2") = returnValue
                Case "WB_RRT3"
                    drWorkingBudget("RRT3") = returnValue
                Case "WB_RRT4"
                    drWorkingBudget("RRT4") = returnValue
                Case "WB_RRT5"
                    drWorkingBudget("RRT5") = returnValue
                Case "PREV_WB_RRT1"
                    drWorkingBudget("PrevRRT1") = dsData.Tables(0).Compute(strExpression2, strFilter)
                Case "PREV_WB_RRT2"
                    drWorkingBudget("PrevRRT2") = dsData.Tables(0).Compute(strExpression2, strFilter)
                Case "PREV_WB_RRT3"
                    drWorkingBudget("PrevRRT3") = dsData.Tables(0).Compute(strExpression2, strFilter)
                Case "PREV_WB_RRT4"
                    drWorkingBudget("PrevRRT4") = dsData.Tables(0).Compute(strExpression2, strFilter)
                Case "PREV_WB_RRT5"
                    drWorkingBudget("PrevRRT5") = dsData.Tables(0).Compute(strExpression2, strFilter)
                Case "WBRevYear"
                    drWorkingBudget("RevYear") = returnValue
                Case "WBDiffYear"
                    drWorkingBudget("DiffYear") = (Convert.ToDecimal(Nz(drWorkingBudget("RevYear"), 0.0)) - Convert.ToDecimal(Nz(drWorkingBudget("PrevRRT1"), 0.0)))
                Case "WBDIFFRRT1"
                    drWorkingBudget("DIFFRRT1") = returnValue

            End Select


        Next
        drWorkingBudget("DiffYear") = (Convert.ToDecimal(Nz(drWorkingBudget("RevYear"), 0.0)) - Convert.ToDecimal(Nz(drWorkingBudget("PrevRRT1"), 0.0)))
        Return True
    End Function

    Private Function SetAccountNoText(ByRef drInvestments As DataRow, _
                                      ByRef drManufacturingCost As DataRow, _
                                      ByRef drAdministrationCost As DataRow, _
                                      ByRef drTotalExpense As DataRow, _
                                      ByRef drWorkingBudget As DataRow, _
                                      ByRef drOutflowTotal As DataRow) As Boolean
        Dim strColumnName As String = "ACCOUNT_NO"

        drInvestments(strColumnName) = "Investments"
        drManufacturingCost(strColumnName) = P_FC_COST
        drAdministrationCost(strColumnName) = P_ADMIN_COST
        drTotalExpense(strColumnName) = "Total Expense"
        drWorkingBudget(strColumnName) = "Working Budget"
        drOutflowTotal(strColumnName) = "Outflow Total"
        Return True
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

    Private Sub frmBG0440_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        If Not excelApp Is Nothing Then
            excelApp.Quit()
            excelApp = Nothing
        End If
    End Sub

    Private Sub frmBG0440_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not excelApp Is Nothing Then
            excelApp.Quit()
            excelApp = Nothing
        End If
    End Sub

    Private Sub frmBG0440_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
                MessageBox.Show("Please select a Period Type!", "Summary By Account No Report", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Me.cboPeriodType.Focus()
                Me.cboPeriodType.SelectAll()
                Return
            End If

            If fncCheckRevNo() = False Then

                MessageBox.Show("No budget data found, please try it again.", "Summary By Account No Report", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            'If fncCheckPrevRevNo() = False Then
            '    MessageBox.Show("No previous budget data found, please try it again.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            '    Exit Sub
            'End If

            Cursor = Cursors.WaitCursor

            myClsBG0440BL.BudgetYear = CStr(Me.numYear.Value)
            myClsBG0440BL.PeriodType = CStr(Me.cboPeriodType.SelectedValue)
            myClsBG0440BL.ProjectNo = Me.numProjectNo.Value.ToString
            'myClsBG0440BL.MTPBudget = Me.chkShowMTP.Checked
            myClsBG0440BL.UserLevelId = p_intUserLevelId
            If Me.cboRevNo.DataSource IsNot Nothing Then
                myClsBG0440BL.RevNo = Me.cboRevNo.SelectedValue.ToString
            End If


            myClsBG0440BL.PrevProjectNo = Me.numPrevProjectNo.Value.ToString
            If Me.cboPrevRevno.DataSource IsNot Nothing AndAlso _
                Me.cboPrevRevno.SelectedValue IsNot Nothing Then
                myClsBG0440BL.PrevRevNo = Me.cboPrevRevno.SelectedValue.ToString
            End If

            If myClsBG0440BL.getAccountData() Then

                Dim ds As DataSet = myClsBG0440BL.AccountCodeData

                If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then

                    myClsBG0440BL.GetBudgetStatus()

                    myClsBG0440BL.GetAuthImage()
                    ds.Tables.Add(myClsBG0440BL.AuthImage)

                    Dim strYear As String = Me.numYear.Value.ToString
                    '//Create output columns
                    Dim dtColumns As DataTable = CreateTableTemplate()
                    Dim dsGroups As DataSet = Nothing

                    Select Case CType(Me.cboPeriodType.SelectedValue, enumPeriodType)
                        Case enumPeriodType.OriginalBudget
                            'strReportName = "RPT004-1.rpt"
                            '//Insert ColummData (Original)
                            InsertOriginalColumnData(dtColumns, strYear)
                            '//Dataset Groupby 
                            dsGroups = SetupOriginalGroupbyData(ds, "EXPENSE_TYPE", "EXPENSE_TYPE", 7)
                            '//Generate excel
                            GeneratOriginalExcel(dsGroups, dtColumns)

                        Case enumPeriodType.EstimateBudget
                            'strReportName = "RPT004-2.rpt"
                            '//Insert Columndata
                            InsertEstimateColumnData(dtColumns, strYear)
                            '//DataSet GroupBy (Estimate)
                            dsGroups = SetupEstimateGroupbyData(ds, "EXPENSE_TYPE", "EXPENSE_TYPE", 7)
                            '//Generate excel
                            GeneratEstimateExcel(dsGroups, dtColumns)

                        Case enumPeriodType.ReviseBudget
                            If Not chkShowMTP.Checked Then
                                'strReportName = "RPT004-3.rpt"
                                '//Insert ColumnData
                                InsertReviseColumnData(dtColumns, strYear)
                                '//DataSet GroupBy
                                dsGroups = SetupReviseGroupbyData(ds, "EXPENSE_TYPE", "EXPENSE_TYPE", 7, False)
                                '//Generate Excel
                                GeneratReviseExcel(dsGroups, dtColumns, False)
                            Else
                                'strReportName = "RPT004-4.rpt"
                                '//Insert ColumnData
                                InsertReviseMTPColumnData(dtColumns, strYear)
                                '//DataSet GroupBy
                                dsGroups = SetupReviseGroupbyData(ds, "EXPENSE_TYPE", "EXPENSE_TYPE", 7, True)
                                '//Generate Excel
                                GeneratReviseExcel(dsGroups, dtColumns, True)
                            End If

                        Case enumPeriodType.MTPBudget
                            '//Insert Columndata
                            InsertMTPColumnData(dtColumns, strYear)
                            '//DataSet GroupBy (Estimate)
                            dsGroups = SetupMTPGroupbyData(ds, "EXPENSE_TYPE", "EXPENSE_TYPE", 7)
                            '//Generate excel
                            GeneratMTPExcel(dsGroups, dtColumns)

                    End Select

                Else
                    MessageBox.Show("No budget data found, please try it again.", "Summary By Account No. Report", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Summary By Account No. Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Cursor = Cursors.Default
    End Sub

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

#End Region

End Class