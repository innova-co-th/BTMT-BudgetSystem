Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports Microsoft.Office.Interop
Imports System.Drawing.Printing

Public Class frmBG0420

#Region "Variable"
    Private myClsBG0420BL As New clsBG0420BL
    Private myClsBG0310BL As New clsBG0310BL
    Private clsBG0400 As frmBG0400
    Private excelApp As Excel.Application
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

    Private Function InitPage() As Boolean

        Try
            Me.numYear.Value = Now.Year
            Me.chkHideEstimate.Checked = False
            LoadPeriodType()
            'cboPeriodType.Items.Clear()
            'cboPeriodType.Items.Add("Original Budget")
            'cboPeriodType.Items.Add("Estimate Budget")
            'cboPeriodType.Items.Add("Forecast Budget")
            'cboPeriodType.SelectedIndex = 0

            If p_intUserLevelId = enumUserLevel.SystemAdministrator Then
                Me.lblRevNo.Visible = True
                Me.cboRevNo.Visible = True
                LoadRevNo()

                'Me.lblPrevRevNo.Visible = True
                'Me.cboPrevRevno.Visible = True
                'LoadPrevRevNo()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "RPT002", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try

    End Function

    Private Sub LoadRevNo()
        If Me.cboRevNo.Visible = True Then

            If Me.cboPeriodType.SelectedIndex < 0 OrElse _
                Me.numProjectNo.Value <= 0 OrElse _
                Me.numYear.Value <= 0 Then

                Me.cboRevNo.DataSource = Nothing
                Exit Sub

            End If

            '//Change BugetYear Parameter for MBP
            Dim intNumYear As Integer
            intNumYear = CInt(Me.numYear.Value.ToString)
            Dim strNumYear As String

            If CType(cboPeriodType.SelectedValue, Integer) = CType(enumPeriodType.MBPBudget, Integer) Then
                intNumYear = intNumYear - 1
            End If
            strNumYear = intNumYear.ToString
            '//Change BugetYear Parameter for MBP

            Dim strProjectNo = Me.numProjectNo.Value.ToString
            If Not strProjectNo Is Nothing And strProjectNo <> String.Empty And strProjectNo <> "System.Data.DataRowView" Then

                myClsBG0310BL.BudgetYear = strNumYear 'Me.numYear.Value.ToString
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

    Private Sub frmBG0420_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        If Not excelApp Is Nothing Then
            excelApp.Quit()
            excelApp = Nothing
        End If
    End Sub

    Private Sub frmBG0420_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not excelApp Is Nothing Then
            excelApp.Quit()
            excelApp = Nothing
        End If
    End Sub

    Private Sub frmBG0420_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        InitPage()
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrint.Click

        Try
            If fncCheckRevNo() = False Then

                MessageBox.Show("No budget data found, please try it again.", "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If


            PrintDialog1.AllowSomePages = True
            '     printDialog1.ShowDialog()

            If PrintDialog1.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then

                Me.Cursor = Cursors.WaitCursor
                Dim m_Report As ReportDocument = New ReportDocument()
                Dim reportPath As String = String.Empty

                Dim strPeriod As String = String.Empty

                Select Case CInt(Me.cboPeriodType.SelectedValue)
                    Case CType(enumPeriodType.OriginalBudget, Integer)
                        reportPath = p_strAppPath & "\Reports\RPT002-1.rpt"
                        strPeriod = "Original"
                        Exit Select
                    Case CType(enumPeriodType.EstimateBudget, Integer)
                        reportPath = p_strAppPath & "\Reports\RPT002-2.rpt"
                        strPeriod = "Estimate"
                        Exit Select
                    Case CType(enumPeriodType.ForecastBudget, Integer)
                        If Me.chkHideEstimate.Checked = True Then
                            reportPath = p_strAppPath & "\Reports\RPT002-4.rpt"
                        Else
                            reportPath = p_strAppPath & "\Reports\RPT002-3.rpt"
                        End If
                        strPeriod = "Forecast"
                        Exit Select
                    Case CType(enumPeriodType.MBPBudget, Integer)
                        reportPath = p_strAppPath & "\Reports\RPT002-5.rpt"
                        strPeriod = "Original"
                        Exit Select
                    Case Else
                        reportPath = p_strAppPath & "\Reports\RPT002-1.rpt"
                        strPeriod = "MBP"
                        Exit Select
                End Select

                m_Report.Load(reportPath)

                '//Change BugetYear Parameter for MBP
                Dim intNumYear As Integer
                intNumYear = CInt(Me.numYear.Value.ToString)
                Dim strNumYear As String

                If CType(cboPeriodType.SelectedValue, Integer) = CType(enumPeriodType.MBPBudget, Integer) Then
                    intNumYear = intNumYear - 1
                End If
                strNumYear = intNumYear.ToString
                '//Change BugetYear Parameter for MBP

                myClsBG0420BL.BudgetYear = strNumYear 'Me.numYear.Value.ToString
                myClsBG0420BL.PeriodType = (Me.cboPeriodType.SelectedValue).ToString
                myClsBG0420BL.MTPChecked = Me.chkHideEstimate.Checked
                myClsBG0420BL.ProjectNo = Me.numProjectNo.Value.ToString
                myClsBG0420BL.UserLevelId = p_intUserLevelId
                If Me.cboRevNo.DataSource IsNot Nothing Then
                    myClsBG0420BL.RevNo = Me.cboRevNo.SelectedValue.ToString
                End If

                myClsBG0420BL.PrevProjectNo = Me.numPrevProjectNo.Value.ToString
                If Me.cboPrevRevno.DataSource IsNot Nothing AndAlso _
                    Me.cboPrevRevno.SelectedValue IsNot Nothing Then
                    myClsBG0420BL.PrevRevNo = Me.cboPrevRevno.SelectedValue.ToString
                Else
                    myClsBG0420BL.PrevRevNo = String.Empty
                End If

                Dim ds As DataSet
                If myClsBG0420BL.GetBudgetData() = False Then
                    ds = Nothing
                Else
                    ds = myClsBG0420BL.BudgetData
                    If ds Is Nothing Or ds.Tables(0).Rows.Count = 0 Then
                        MessageBox.Show("No budget data found, please try it again.", "RPT006", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Me.Cursor = Cursors.Default
                        Return
                    End If
                End If
                m_Report.SetDataSource(ds)

                myClsBG0420BL.GetBudgetStatus()

                If myClsBG0420BL.BudgetStatus >= 5 Then
                    m_Report.ReportDefinition.ReportObjects("picAuth10").ObjectFormat.EnableSuppress = False
                Else
                    m_Report.ReportDefinition.ReportObjects("picAuth10").ObjectFormat.EnableSuppress = True
                End If

                If myClsBG0420BL.BudgetStatus >= 6 Then
                    m_Report.ReportDefinition.ReportObjects("picAuth3").ObjectFormat.EnableSuppress = False
                Else
                    m_Report.ReportDefinition.ReportObjects("picAuth3").ObjectFormat.EnableSuppress = True
                End If

                m_Report.SetParameterValue("PERIOD", strPeriod)
                'm_Report.SetParameterValue("BUDGET_YEAR", Me.numYear.Value.ToString)
                m_Report.SetParameterValue("BUDGET_YEAR", strNumYear)
                m_Report.SetParameterValue("HALF_BUDGET_YEAR", Me.numYear.Value.ToString.Substring(2, 2))
                m_Report.SetParameterValue("PROJECT_NO", Me.numProjectNo.Value.ToString)

                m_Report.PrintOptions.PrinterName = PrintDialog1.PrinterSettings.PrinterName
                m_Report.PrintToPrinter(PrintDialog1.PrinterSettings.Copies, _
                                        PrintDialog1.PrinterSettings.Collate, _
                                        PrintDialog1.PrinterSettings.FromPage, _
                                        PrintDialog1.PrinterSettings.ToPage)
                'Dim pt As Printing.PrintDocument

                Me.Cursor = Cursors.Default
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "RPT006", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.Cursor = Cursors.Default
            Return
        End Try

    End Sub

    Private Sub cmdPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPreview.Click

        Try
            If fncCheckRevNo() = False Then

                MessageBox.Show("No budget data found, please try it again.", "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            Me.Cursor = Cursors.WaitCursor

            If clsBG0400 IsNot Nothing Then
                clsBG0400.Close()
                clsBG0400.Dispose()
            End If
            clsBG0400 = New frmBG0400()

            '//Change BugetYear Parameter for MBP
            Dim intNumYear As Integer
            intNumYear = CInt(Me.numYear.Value.ToString)
            Dim strNumYear As String

            If CType(cboPeriodType.SelectedValue, Integer) = CType(enumPeriodType.MBPBudget, Integer) Then
                intNumYear = intNumYear - 1
            End If
            strNumYear = intNumYear.ToString
            '//Change BugetYear Parameter for MBP

            myClsBG0420BL.BudgetYear = strNumYear 'Me.numYear.Value.ToString
            myClsBG0420BL.PeriodType = (Me.cboPeriodType.SelectedValue).ToString
            myClsBG0420BL.MTPChecked = Me.chkHideEstimate.Checked
            myClsBG0420BL.ProjectNo = Me.numProjectNo.Value.ToString
            myClsBG0420BL.UserLevelId = p_intUserLevelId
            If Me.cboRevNo.DataSource IsNot Nothing Then
                myClsBG0420BL.RevNo = Me.cboRevNo.SelectedValue.ToString
            End If

            myClsBG0420BL.PrevProjectNo = Me.numPrevProjectNo.Value.ToString
            If Me.cboPrevRevno.DataSource IsNot Nothing AndAlso _
                Me.cboPrevRevno.SelectedValue IsNot Nothing Then
                myClsBG0420BL.PrevRevNo = Me.cboPrevRevno.SelectedValue.ToString
            Else
                myClsBG0420BL.PrevRevNo = String.Empty
            End If

            If myClsBG0420BL.GetBudgetData() = False Then
                clsBG0400.DS = Nothing
            Else
                clsBG0400.DS = myClsBG0420BL.BudgetData
                If clsBG0400.DS Is Nothing Or clsBG0400.DS.Tables(0).Rows.Count = 0 Then
                    MessageBox.Show("No budget data found, please try it again.", "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Me.Cursor = Cursors.Default
                    Return
                End If
            End If

            myClsBG0420BL.GetBudgetStatus()


            Dim strPeriod As String = String.Empty
            Select Case CInt(Me.cboPeriodType.SelectedValue)
                Case CType(enumPeriodType.OriginalBudget, Integer)
                    clsBG0400.ReportName = "RPT002-1.rpt"
                    strPeriod = "Original"
                    Exit Select
                Case CType(enumPeriodType.EstimateBudget, Integer)
                    clsBG0400.ReportName = "RPT002-2.rpt"
                    strPeriod = "Estimate"
                    Exit Select
                Case CType(enumPeriodType.ForecastBudget, Integer)
                    If Me.chkHideEstimate.Checked = True Then
                        clsBG0400.ReportName = "RPT002-4.rpt"
                    Else
                        clsBG0400.ReportName = "RPT002-3.rpt"
                    End If
                    strPeriod = "Forecast"
                    Exit Select
                Case CType(enumPeriodType.MBPBudget, Integer)
                    clsBG0400.ReportName = "RPT002-5.rpt"
                    strPeriod = "MBP"
                    Exit Select
            End Select

            'clsBG0400.ConfigureCrystalReports()
            clsBG0400.BudgetYear = strNumYear 'Me.numYear.Value.ToString
            clsBG0400.Period = strPeriod
            clsBG0400.ReportType = "SummaryByPersonInCharge"
            clsBG0400.BudgetStatus = myClsBG0420BL.BudgetStatus
            clsBG0400.ProjectNo = Me.numProjectNo.Value.ToString

            clsBG0400.MdiParent = p_frmBG0010
            clsBG0400.Show()

            If clsBG0400.WindowState = FormWindowState.Minimized Then
                clsBG0400.WindowState = FormWindowState.Normal
            End If
            clsBG0400.BringToFront()
            Me.Cursor = Cursors.Default

        Catch ex As Exception
            MessageBox.Show(ex.Message, "RPT002", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.Cursor = Cursors.Default
            Return
        End Try

    End Sub

    Private Sub cboPeriodType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPeriodType.SelectedIndexChanged

        If CInt(cboPeriodType.SelectedValue) = CType(enumPeriodType.ForecastBudget, Integer) Then
            Me.chkHideEstimate.Enabled = True
        Else
            Me.chkHideEstimate.Checked = False
            Me.chkHideEstimate.Enabled = False
        End If


        LoadRevNo()

    End Sub

    Private Sub cmdExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExcel.Click

        If fncCheckRevNo() = False Then

            MessageBox.Show("No budget data found, please try it again.", "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If



        Me.Cursor = Cursors.WaitCursor

        '//Change BugetYear Parameter for MBP
        Dim intNumYear As Integer
        intNumYear = CInt(Me.numYear.Value.ToString)
        Dim strNumYear As String

        If CType(cboPeriodType.SelectedValue, Integer) = CType(enumPeriodType.MBPBudget, Integer) Then
            intNumYear = intNumYear - 1
        End If
        strNumYear = intNumYear.ToString
        '//Change BugetYear Parameter for MBP

        '//Get Export Data
        Dim dsData As DataSet
        myClsBG0420BL.BudgetYear = strNumYear 'Me.numYear.Value.ToString
        myClsBG0420BL.PeriodType = (Me.cboPeriodType.SelectedValue).ToString
        myClsBG0420BL.MTPChecked = Me.chkHideEstimate.Checked
        myClsBG0420BL.ProjectNo = Me.numProjectNo.Value.ToString
        myClsBG0420BL.UserLevelId = p_intUserLevelId
        If Me.cboRevNo.DataSource IsNot Nothing Then
            myClsBG0420BL.RevNo = Me.cboRevNo.SelectedValue.ToString
        End If

        myClsBG0420BL.PrevProjectNo = Me.numPrevProjectNo.Value.ToString
        If Me.cboPrevRevno.DataSource IsNot Nothing AndAlso _
            Me.cboPrevRevno.SelectedValue IsNot Nothing Then
            myClsBG0420BL.PrevRevNo = Me.cboPrevRevno.SelectedValue.ToString
        Else
            myClsBG0420BL.PrevRevNo = String.Empty
        End If

        Dim strPeriod As String = cboPeriodType.Text
        strPeriod = strPeriod.Substring(0, strPeriod.IndexOf("Budget") - 1)

        If myClsBG0420BL.GetBudgetData() = False Then
            dsData = Nothing
            MessageBox.Show("No budget data found, please try it again.", "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Cursor = Cursors.Default
            Return
        Else
            dsData = myClsBG0420BL.BudgetData
            If dsData Is Nothing Or dsData.Tables(0).Rows.Count = 0 Then
                MessageBox.Show("No budget data found, please try it again.", "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Cursor = Cursors.Default
                Return
            End If
        End If

        Dim dtAuthorizeImages As DataTable = myClsBG0420BL.BudgetData.Tables(1)

        '//Create output columns
        Dim strYear As String = Me.numYear.Value.ToString
        Dim dtColumns As DataTable = CreateTableTemplate()

        Dim strPeriodType As String = cboPeriodType.Text
        Dim strProjectNo As String = Me.numProjectNo.Value.ToString

        Dim strSubTitle As String = String.Empty

        If strProjectNo <> "1" Then
            strSubTitle = "Summary by Person In Charge : " + strPeriodType + " " + strYear + " (Project No. " + strProjectNo + ")"
        Else
            strSubTitle = "Summary by Person In Charge : " + strPeriodType + " " + strYear
        End If

        Select Case CInt(Me.cboPeriodType.SelectedValue)

            Case CType(enumPeriodType.OriginalBudget, Integer)  '//Original 

                InsertOriginalColumnData(dtColumns, strYear)

                '//Create group data
                Dim intGroupFirstIndex As Integer = 0
                Dim intGroupSecondIndex As Integer = 0
                Dim dsGroups As DataSet = SetupPICSummaryGroupbyData(dsData, "PIC_SHOW_FLAG", "PIC_SHOW_FLAG", 6, False, intGroupFirstIndex, intGroupSecondIndex)

                '//Create Output Excel
                OutputExcel(dsGroups, dtColumns, False, strSubTitle, strYear, False, _
                            intGroupFirstIndex, intGroupSecondIndex, strPeriod)

            Case CType(enumPeriodType.EstimateBudget, Integer) '//Estimate

                InsertEstimateColumnData(dtColumns, strYear)

                '//Create group data
                Dim intGroupFirstIndex As Integer = 0
                Dim intGroupSecondIndex As Integer = 0
                Dim dsGroups As DataSet = SetupPICSummaryGroupbyData(dsData, "PIC_SHOW_FLAG", "PIC_SHOW_FLAG", 6, False, intGroupFirstIndex, intGroupSecondIndex)

                '//Create Output Excel
                OutputExcel(dsGroups, dtColumns, False, strSubTitle, strYear, False, _
                            intGroupFirstIndex, intGroupSecondIndex, strPeriod)

            Case CType(enumPeriodType.ForecastBudget, Integer)  '//Forecast

                If Me.chkHideEstimate.Checked = True Then
                    InsertMTPColumnData(dtColumns, strYear)
                Else
                    InsertForecastColumnData(dtColumns, strYear)
                End If

                '//Create group data
                Dim intGroupFirstIndex As Integer = 0
                Dim intGroupSecondIndex As Integer = 0
                Dim dsGroups As DataSet = SetupPICSummaryGroupbyData(dsData, "PIC_SHOW_FLAG", "PIC_SHOW_FLAG", 6, False, intGroupFirstIndex, intGroupSecondIndex)

                Dim bMTPCheck As Boolean = chkHideEstimate.Checked

                '//Create Output Excel
                OutputExcel(dsGroups, dtColumns, bMTPCheck, strSubTitle, strYear, False, _
                            intGroupFirstIndex, intGroupSecondIndex, strPeriod)

            Case CType(enumPeriodType.MBPBudget, Integer)  '//MTP
                'InsertMTPBudgetColumnData(dtColumns, strYear)
                InsertMTPBudgetColumnData(dtColumns, strNumYear)

                '//Create group data
                Dim intGroupFirstIndex As Integer = 0
                Dim intGroupSecondIndex As Integer = 0
                Dim dsGroups As DataSet = SetupPICSummaryGroupbyData(dsData, "PIC_SHOW_FLG", "PIC_SHOW_FLG", 6, False, intGroupFirstIndex, intGroupSecondIndex)

                '//Create Output Excel
                OutputExcel(dsGroups, dtColumns, False, strSubTitle, strYear, False, _
                            intGroupFirstIndex, intGroupSecondIndex, strPeriod)


        End Select

        Me.Cursor = Cursors.Default

    End Sub

    Private Function OutputExcel(ByVal dsData As DataSet, ByVal dtColumns As DataTable, ByVal bMTPCheck As Boolean, _
                                 ByVal strSubTitle As String, ByVal strYear As String, ByVal bShowGroupName As Boolean, _
                                 ByVal intGroupFirstIndex As Integer, ByVal intGroupSecondIndex As Integer, _
                                 ByVal strPeriod As String) As Boolean

        If excelApp Is Nothing Then
            excelApp = New Excel.Application
        End If

        Dim rowStartIndex As Integer = 10
        Dim colStartIndex As Integer = 9

        Dim xBk As Excel.Workbook = Nothing
        Dim xSt As Excel.Worksheet = Nothing

        xBk = excelApp.Workbooks.Add(System.Reflection.Missing.Value)
        If xBk.Worksheets.Count > 1 Then
            For i As Integer = 1 To xBk.Worksheets.Count - 1
                CType(xBk.Worksheets(i), Excel.Worksheet).Delete()
            Next
        End If


        '//Set Style Value < 0 please fill color "Red"
        Dim style As Excel.Style = excelApp.ActiveWorkbook.Styles.Add("NewStyle")
        ''style.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
        style.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)

        For intSheetCount As Integer = 0 To dsData.Tables.Count - 1

            If intSheetCount <> 0 Then
                xBk.Sheets.Add()
            End If

            xSt = CType(xBk.ActiveSheet, Excel.Worksheet)
            If dsData.Tables(intSheetCount).TableName = "REVISE_BUDGET" Then
                xSt.Name = "Forecast Budget"
            Else
                xSt.Name = dsData.Tables(intSheetCount).TableName
            End If


            '//Setup DataColumn
            For i As Integer = 0 To dtColumns.Rows.Count - 1
                xSt.Cells(colStartIndex, i + 1) = dtColumns.Rows(i)("Column_Title").ToString
                xSt.Range(xSt.Cells(colStartIndex, i + 1), xSt.Cells(colStartIndex, i + 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            Next

            Dim arrCols() As Integer

            If strPeriod = "Original" Then

                arrCols = New Integer() {1, 4, 5, 12, 19, 20, 21, 22, 23, 24}
                SetupOriginalColumnsCells(xSt, colStartIndex, 2, 3, "Person in Charge Section", arrCols, 6, 11, strYear, True, True, 13, 18)

            ElseIf strPeriod = "Estimate" Then

                arrCols = New Integer() {1, 4, 5, 12, 13, 14, 15, 16}
                SetupEstimateColumnsCells(xSt, colStartIndex, 2, 3, "Person in Charge Section", arrCols, 6, 8, 9, 11)

            ElseIf strPeriod = "Forecast" Then

                arrCols = New Integer() {1, 4, 11, 12, 13, 20, 21, 22, 23, 24}
                If bMTPCheck = True Then
                    arrCols = New Integer() {1, 4, 11, 12, 13, 14, 15}
                End If
                SetupForecastColumnsCells(xSt, colStartIndex, bMTPCheck, 2, 3, "Person in Charge Section", arrCols, 5, 7, 8, 10, 14, 19, 5, 10, 24, 28)

            ElseIf strPeriod = "MBP" Then

                arrCols = New Integer() {1, 4, 5, 6} '// Two Row Merge Col
                SetupMTPColumnsCells(xSt, colStartIndex, 2, 3, "Person in Charge Section", arrCols, 7, 10)

            End If

            '//Setup Data
            For rowIndex As Integer = 0 To dsData.Tables(intSheetCount).Rows.Count - 1

                Dim row As DataRow = dsData.Tables(intSheetCount).Rows(rowIndex)

                For colIndex As Integer = 0 To dtColumns.Rows.Count - 1

                    Dim strColumnName As String = dtColumns.Rows(colIndex)("Column_Name").ToString
                    Dim col As DataColumn = dsData.Tables(intSheetCount).Columns(strColumnName)

                    If col.DataType Is System.Type.GetType("System.DateTime") Then

                        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = (Convert.ToDateTime(row(col.ColumnName).ToString())).ToString("yyyy-MM-dd")
                        xSt.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter

                    ElseIf col.DataType Is System.Type.GetType("System.String") Then

                        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = "'" + row(col.ColumnName).ToString()
                        xSt.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter

                    ElseIf col.DataType Is System.Type.GetType("System.Decimal") Then

                        If row(col.ColumnName).ToString = String.Empty Then
                            row(col.ColumnName) = "0.00"
                        End If

                        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName)
                        xSt.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).NumberFormat = "#,##0.00"


                    Else
                        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName).ToString()
                    End If

                Next
            Next

            '//Add Title for excel 
            Dim strGroupName As String = dsData.Tables(intSheetCount).TableName
            Dim intUnitPriceStart As Integer
            Dim intUnitPriceEnd As Integer
            Dim intAuthorizeStart As Integer
            Dim intAuthorizeEnd As Integer
            Dim intImageIndex As Integer

            SetupTitleIndex(strPeriod, intUnitPriceStart, intUnitPriceEnd, intAuthorizeStart, intAuthorizeEnd, intImageIndex, bMTPCheck)

            Dim bAuthorizeTwoCols As Boolean = False
            If strPeriod = "Estimate" Then
                bAuthorizeTwoCols = True
            End If
            SetupExcelTitle(xSt, strSubTitle, strYear, bMTPCheck, intUnitPriceStart, intUnitPriceEnd, _
                            intAuthorizeStart, intAuthorizeEnd, intImageIndex, bShowGroupName, strGroupName, bAuthorizeTwoCols)

            Dim rowMax As Integer = dsData.Tables(intSheetCount).Rows.Count + colStartIndex
            Dim colMax As Integer = dtColumns.Rows.Count

            '//Setup budget order name column to be left align
            xSt.Range(xSt.Cells(rowStartIndex, 3), xSt.Cells(rowMax, 3)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup Total lines
            SetupTotalLines(xSt, rowMax - 1, "Expenses Total", "Left", 3, 3, 3, rowMax - 4, rowMax - 3, rowMax - 2, colMax, bMTPCheck)

            xSt.Range(xSt.Cells(rowMax - 4, 1), xSt.Cells(rowMax, 2)).ClearContents()
            xSt.Range(xSt.Cells(rowMax - 4, 1), xSt.Cells(rowMax, 2)).MergeCells = True

            xSt.Range(xSt.Cells(rowMax, 3), xSt.Cells(rowMax, colMax)).ClearContents()
            xSt.Range(xSt.Cells(rowMax, 3), xSt.Cells(rowMax, colMax)).MergeCells = True

            Dim strGroup1 As String = xSt.Range(xSt.Cells(rowStartIndex, 1), xSt.Cells(rowStartIndex, 1)).Value.ToString
            xSt.Range(xSt.Cells(rowStartIndex, 1), xSt.Cells(rowStartIndex + intGroupFirstIndex, 1)).ClearContents()
            xSt.Range(xSt.Cells(rowStartIndex, 1), xSt.Cells(rowStartIndex + intGroupFirstIndex, 1)).MergeCells = True
            xSt.Range(xSt.Cells(rowStartIndex, 1), xSt.Cells(rowStartIndex + intGroupFirstIndex, 1)).Value = strGroup1
            xSt.Range(xSt.Cells(rowStartIndex, 1), xSt.Cells(rowStartIndex + intGroupFirstIndex, 1)).Font.Bold = True
            xSt.Range(xSt.Cells(rowStartIndex, 1), xSt.Cells(rowStartIndex + intGroupFirstIndex, 1)).VerticalAlignment = Excel.XlVAlign.xlVAlignTop

            'xSt.Range(xSt.Cells(rowStartIndex, 1), xSt.Cells(rowStartIndex, colMax)).Font.Bold = True

            Dim intGroup2 As Integer = rowStartIndex + intGroupFirstIndex + 1
            Dim ObjGroup2 As Object = xSt.Range(xSt.Cells(intGroup2, 1), xSt.Cells(intGroup2, 1)).Value
            If Not IsNothing(ObjGroup2) Then
                Dim strGroup2 As String = ObjGroup2.ToString
                xSt.Range(xSt.Cells(intGroup2, 1), xSt.Cells(rowStartIndex + intGroupSecondIndex, 1)).ClearContents()
                xSt.Range(xSt.Cells(intGroup2, 1), xSt.Cells(rowStartIndex + intGroupSecondIndex, 1)).MergeCells = True
                xSt.Range(xSt.Cells(intGroup2, 1), xSt.Cells(rowStartIndex + intGroupSecondIndex, 1)).Value = strGroup2
                xSt.Range(xSt.Cells(intGroup2, 1), xSt.Cells(rowStartIndex + intGroupSecondIndex, 1)).Font.Bold = True
                xSt.Range(xSt.Cells(intGroup2, 1), xSt.Cells(rowStartIndex + intGroupSecondIndex, 1)).VerticalAlignment = Excel.XlVAlign.xlVAlignTop
            End If


            'xSt.Range(xSt.Cells(intGroup2, 1), xSt.Cells(intGroup2, colMax)).Font.Bold = True

            '//Setup sheet properly width
            xSt.Range(xSt.Cells(2, 1), xSt.Cells(rowMax, colMax)).Columns.AutoFit()

            '//Setup Columns Wrap text
            xSt.Range(xSt.Cells(2, 4), xSt.Cells(rowMax, 4)).Columns.ColumnWidth = 13
            xSt.Range(xSt.Cells(2, 4), xSt.Cells(rowMax, 4)).WrapText = True

            If strPeriod = "Original" Then

                xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 5)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 5)).WrapText = True

                xSt.Range(xSt.Cells(2, 6), xSt.Cells(rowMax, 11)).Columns.ColumnWidth = 12
                xSt.Range(xSt.Cells(2, 13), xSt.Cells(rowMax, 18)).Columns.ColumnWidth = 12

                xSt.Range(xSt.Cells(2, 12), xSt.Cells(rowMax, 12)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 12), xSt.Cells(rowMax, 12)).WrapText = True

                xSt.Range(xSt.Cells(2, 19), xSt.Cells(rowMax, 24)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 19), xSt.Cells(rowMax, 24)).WrapText = True


            ElseIf strPeriod = "Estimate" Then

                xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 5)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 5)).WrapText = True

                xSt.Range(xSt.Cells(2, 6), xSt.Cells(rowMax, 11)).Columns.ColumnWidth = 12

                xSt.Range(xSt.Cells(2, 12), xSt.Cells(rowMax, 16)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 12), xSt.Cells(rowMax, 16)).WrapText = True

            ElseIf strPeriod = "MBP" Then

                xSt.Range(xSt.Cells(2, 7), xSt.Cells(rowMax, 10)).Columns.ColumnWidth = 12
                xSt.Range(xSt.Cells(2, 7), xSt.Cells(rowMax, 10)).WrapText = True
            Else

                xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 10)).Columns.ColumnWidth = 12

                xSt.Range(xSt.Cells(2, 11), xSt.Cells(rowMax, 13)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 11), xSt.Cells(rowMax, 13)).WrapText = True

                xSt.Range(xSt.Cells(2, 14), xSt.Cells(rowMax, 19)).Columns.ColumnWidth = 12

                xSt.Range(xSt.Cells(2, 20), xSt.Cells(rowMax, 23)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 20), xSt.Cells(rowMax, 23)).WrapText = True

                If chkHideEstimate.Checked = True Then
                    xSt.Range(xSt.Cells(2, 24), xSt.Cells(rowMax, 28)).Columns.ColumnWidth = 12
                End If

            End If

            colStartIndex = colStartIndex - 1
            '//Setup Column Font 
            xSt.Range(xSt.Cells(colStartIndex, 1), xSt.Cells(colStartIndex + 1, colMax)).Font.Bold = True
            xSt.Range(xSt.Cells(colStartIndex, 1), xSt.Cells(rowMax, colMax)).Font.Name = "Tahoma"
            xSt.Range(xSt.Cells(colStartIndex, 1), xSt.Cells(rowMax, colMax)).Font.Size = 10

            '//Setup border
            xSt.Range(xSt.Cells(colStartIndex, 1), xSt.Cells(rowMax, colMax)).Borders.LineStyle = 1
            xSt.Range(xSt.Cells(colStartIndex, 1), xSt.Cells(rowMax, 1)).Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlThin
            xSt.Range(xSt.Cells(colStartIndex, 1), xSt.Cells(colStartIndex, colMax)).Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThin
            xSt.Range(xSt.Cells(colStartIndex, colMax), xSt.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThin
            xSt.Range(xSt.Cells(rowMax, 1), xSt.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThin

            '//Clear data last row on Group
            xSt.Range(xSt.Cells(rowStartIndex + intGroupFirstIndex, 4), xSt.Cells(rowStartIndex + intGroupFirstIndex, colMax)).ClearContents()
            xSt.Range(xSt.Cells(rowStartIndex + intGroupSecondIndex, 4), xSt.Cells(rowStartIndex + intGroupSecondIndex, colMax)).ClearContents()

            '//Set Bold first row on Group
            xSt.Range(xSt.Cells(rowStartIndex, 4), xSt.Cells(rowStartIndex, colMax)).Font.Bold = True
            xSt.Range(xSt.Cells(intGroup2, 4), xSt.Cells(intGroup2, colMax)).Font.Bold = True

            '//Set NumberFormat = "#,##0.00;[Red]-#,##0.00"
            If strPeriod = "MBP" Then
                xSt.Range(xSt.Cells(colStartIndex, 4), xSt.Cells(rowMax - 1, 5)).NumberFormat = "#,##0.00;[Red]-#,##0.00"
                xSt.Range(xSt.Cells(colStartIndex, 7), xSt.Cells(rowMax - 1, 8)).NumberFormat = "#,##0.00;[Red]-#,##0.00"
                xSt.Range(xSt.Cells(colStartIndex, 10), xSt.Cells(rowMax - 1, colMax)).NumberFormat = "#,##0.00;[Red]-#,##0.00"
            Else
                xSt.Range(xSt.Cells(colStartIndex, 4), xSt.Cells(rowMax - 1, colMax)).NumberFormat = "#,##0.00;[Red]-#,##0.00"
            End If


            '//Set Frame All
            xSt.Range(xSt.Cells(colStartIndex, 4), xSt.Cells(rowMax - 1, colMax)).Borders.LineStyle = 1
            xSt.Range(xSt.Cells(colStartIndex, 4), xSt.Cells(rowMax - 1, colMax)).Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium
            xSt.Range(xSt.Cells(colStartIndex, 4), xSt.Cells(rowMax - 1, colMax)).Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium
            xSt.Range(xSt.Cells(colStartIndex, 4), xSt.Cells(rowMax - 1, colMax)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            xSt.Range(xSt.Cells(colStartIndex, 4), xSt.Cells(rowMax - 1, colMax)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium

            '//Set Frame
            If strPeriod = "Original" Then

                xSt.Range(xSt.Cells(colStartIndex, 4), xSt.Cells(rowMax - 1, 5)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 6), xSt.Cells(rowMax - 1, 11)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 11), xSt.Cells(rowMax - 1, 12)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 12), xSt.Cells(rowMax - 1, 18)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 18), xSt.Cells(rowMax - 1, 19)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                'xSt.Range(xSt.Cells(colStartIndex, 14), xSt.Cells(rowMax - 1, 15)).Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium
                'xSt.Range(xSt.Cells(colStartIndex, 16), xSt.Cells(rowMax - 1, 17)).Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium

            ElseIf strPeriod = "Estimate" Then

                xSt.Range(xSt.Cells(colStartIndex, 4), xSt.Cells(rowMax - 1, 5)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 6), xSt.Cells(rowMax - 1, 8)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 9), xSt.Cells(rowMax - 1, 11)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 12), xSt.Cells(rowMax - 1, 14)).Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium

            ElseIf strPeriod = "Forecast" Then

                xSt.Range(xSt.Cells(colStartIndex, 4), xSt.Cells(rowMax - 1, 4)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                If bMTPCheck = False Then
                    xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax - 1, 7)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                End If
                xSt.Range(xSt.Cells(colStartIndex, 8), xSt.Cells(rowMax - 1, 10)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 11), xSt.Cells(rowMax - 1, 13)).Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 13), xSt.Cells(rowMax - 1, 13)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium

                If bMTPCheck = False Then
                    xSt.Range(xSt.Cells(colStartIndex, 14), xSt.Cells(rowMax - 1, 19)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                    xSt.Range(xSt.Cells(colStartIndex, 20), xSt.Cells(rowMax - 1, 24)).Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium
                End If

            ElseIf strPeriod = "MBP" Then

                xSt.Range(xSt.Cells(colStartIndex, 4), xSt.Cells(rowMax - 1, 5)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 6), xSt.Cells(rowMax - 1, 6)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium

                '//Set font color
                xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax - 1, 6)).Font.Color = RGB(128, 128, 128)
                xSt.Range(xSt.Cells(colStartIndex, 8), xSt.Cells(rowMax - 1, 8)).Font.Color = RGB(128, 128, 128)
                xSt.Range(xSt.Cells(colStartIndex, 9), xSt.Cells(rowMax - 1, 9)).Font.Color = RGB(128, 128, 128)
                'xSt.Range(xSt.Cells(colStartIndex, 12), xSt.Cells(rowMax - 1, 12)).Font.Color = RGB(128, 128, 128)
                'xSt.Range(xSt.Cells(colStartIndex, 14), xSt.Cells(rowMax - 1, 14)).Font.Color = RGB(128, 128, 128)

            End If
        Next

        '//Show Excel
        excelApp.Visible = True

        '//-- Begin Add by S.Watcharapong 2011-05-24
        '//Release memory
        BGCommon.ExcelReleasememory(excelApp, xBk, xSt)
        '//-- End Add 2011-05-24

        Return True

    End Function

    Private Function SetupTitleIndex(ByVal strPeriod As String, ByRef intUnitPriceStart As Integer, _
                                     ByRef intUnitPriceEnd As Integer, ByRef intAuthorizeStart As Integer, _
                                     ByRef intAuthorizeEnd As Integer, ByRef intImageIndex As Integer, ByVal bMTPCheck As Boolean) As Boolean

        Select Case strPeriod

            Case "Original"
                intUnitPriceStart = 24
                intUnitPriceEnd = 24

                intAuthorizeStart = 21
                intAuthorizeEnd = 22

                intImageIndex = 905

            Case "Estimate"
                intUnitPriceStart = 15
                intUnitPriceEnd = 16

                intAuthorizeStart = 10
                intAuthorizeEnd = 12

                intImageIndex = 795

            Case "Forecast"
                If bMTPCheck = False Then
                    intUnitPriceStart = 23
                    intUnitPriceEnd = 24
                Else
                    intUnitPriceStart = 14
                    intUnitPriceEnd = 15
                End If

                intAuthorizeStart = 20
                intAuthorizeEnd = 21

                intImageIndex = 1160

            Case "MBP"
                intUnitPriceStart = 9
                intUnitPriceEnd = 10

                intAuthorizeStart = 13
                intAuthorizeEnd = 14

                intImageIndex = 805

            Case Else
                Exit Select
        End Select

        Return True

    End Function

    Private Function InsertOriginalColumnData(ByRef dtColumns As DataTable, ByVal strYear As String) As Boolean

        Dim dRow As DataRow
        Dim strHalfYear As String = strYear.Substring(2, 2)

        Dim intYear As Integer = CInt(strYear)
        Dim strLastYear As String = CStr(intYear - 1)

        Dim strHalfLastYear As String = CStr(intYear - 1).Substring(2, 2)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "Group_Header"
        dRow("Column_Title") = "Description"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERSON_IN_CHARGE_NO"
        dRow("Column_Title") = "Person in Charge Section"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERSON_IN_CHARGE_NAME"
        dRow("Column_Title") = ""
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "ACTUAL_FIRST_HALF"
        dRow("Column_Title") = "Actual 1st Half'" & strHalfLastYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "REVISE_SECOND_HALF"
        dRow("Column_Title") = "Estimate 2nd Half'" & strHalfLastYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M1"
        dRow("Column_Title") = "Jan'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M2"
        dRow("Column_Title") = "Feb'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M3"
        dRow("Column_Title") = "Mar'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M4"
        dRow("Column_Title") = "Apr'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M5"
        dRow("Column_Title") = "May'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M6"
        dRow("Column_Title") = "Jun'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "FIRST_HALF_SUM"
        dRow("Column_Title") = "Total 1st Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M7"
        dRow("Column_Title") = "Jul'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M8"
        dRow("Column_Title") = "Aug'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M9"
        dRow("Column_Title") = "Sep'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M10"
        dRow("Column_Title") = "Oct'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M11"
        dRow("Column_Title") = "Nov'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M12"
        dRow("Column_Title") = "Dec'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "SECOND_HALF_SUM"
        dRow("Column_Title") = "Total 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "YEAR_SUM"
        dRow("Column_Title") = "Total Year'" & strYear
        dtColumns.Rows.Add(dRow)

        'dRow = dtColumns.NewRow()
        'dRow("Column_Name") = "MTP_RRT1"
        'dRow("Column_Title") = "MTP " & CInt(strYear) - 2 & " Year'" & strYear
        'dtColumns.Rows.Add(dRow)

        'dRow = dtColumns.NewRow()
        'dRow("Column_Name") = "DIFF_MTP"
        'dRow("Column_Title") = "Diff vs MTP" & CInt(strYear) - 2
        'dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow()
        dRow("Column_Name") = "MTP_RRT1"
        dRow("Column_Title") = "MBP " & CInt(strYear) - 1 & " Year'" & strYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow()
        dRow("Column_Name") = "DIFF_MTP"
        dRow("Column_Title") = "Diff vs MBP" & CInt(strYear) - 1
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "LAST_YEAR_SUM"
        dRow("Column_Title") = "Total Year'" & strLastYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow()
        dRow("Column_Name") = "DIFF_YEAR"
        dRow("Column_Title") = "Diff vs Year'" & CInt(strYear) - 1
        dtColumns.Rows.Add(dRow)

        Return True

    End Function

    Private Function InsertEstimateColumnData(ByRef dtColumns As DataTable, ByVal strYear As String) As Boolean

        Dim dRow As DataRow
        Dim strHalfYear As String = strYear.Substring(2, 2)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "Group_Header"
        dRow("Column_Title") = "Description"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERSON_IN_CHARGE_NO"
        dRow("Column_Title") = "Person in Charge Section"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERSON_IN_CHARGE_NAME"
        dRow("Column_Title") = ""
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "ACTUAL_FIRST_HALF"
        dRow("Column_Title") = "Actual 1st Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "REVISE_SECOND_HALF"
        dRow("Column_Title") = "Original 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M7"
        dRow("Column_Title") = "Jul'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M8"
        dRow("Column_Title") = "Aug'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M9"
        dRow("Column_Title") = "Sep'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M10"
        dRow("Column_Title") = "Oct'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M11"
        dRow("Column_Title") = "Nov'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M12"
        dRow("Column_Title") = "Dec'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "ESTIMATE_SECOND_HALF"
        dRow("Column_Title") = "Estimate 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFFERENCE_SECOND_HALF"
        dRow("Column_Title") = "Difference 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "ESTIMATE_TOTAL_YEAR"
        dRow("Column_Title") = "Estimate Year'" & strYear
        dtColumns.Rows.Add(dRow)


        dRow = dtColumns.NewRow
        dRow("Column_Name") = "ORIGINAL_FULL_YEAR"
        dRow("Column_Title") = "Original Year'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFFERENCE_ORIGINAL_FULL_YEAR"
        dRow("Column_Title") = "Difference Year'" & strHalfYear
        dtColumns.Rows.Add(dRow)


        Return True

    End Function

    Private Function InsertForecastColumnData(ByRef dtColumns As DataTable, ByVal strYear As String) As Boolean

        Dim dRow As DataRow
        Dim strHalfYear As String = strYear.Substring(2, 2)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "Group_Header"
        dRow("Column_Title") = "Description"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERSON_IN_CHARGE_NO"
        dRow("Column_Title") = "Person in Charge Section"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERSON_IN_CHARGE_NAME"
        dRow("Column_Title") = ""
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "ORIGINAL_FIRST_HALF"
        dRow("Column_Title") = "Original 1st Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M1"
        dRow("Column_Title") = "Jan'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M2"
        dRow("Column_Title") = "Feb'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M3"
        dRow("Column_Title") = "Mar'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M4"
        dRow("Column_Title") = "Apr'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M5"
        dRow("Column_Title") = "May'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M6"
        dRow("Column_Title") = "Jun'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "ESTIMATE_FIRST_HALF"
        dRow("Column_Title") = "Forecast 1st Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_FIRST_HALF"
        dRow("Column_Title") = "Diff 1st Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "ORIGINAL_SECOND_HALF"
        dRow("Column_Title") = "Original 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M7"
        dRow("Column_Title") = "Jul'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M8"
        dRow("Column_Title") = "Aug'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M9"
        dRow("Column_Title") = "Sep'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M10"
        dRow("Column_Title") = "Oct'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M11"
        dRow("Column_Title") = "Nov'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M12"
        dRow("Column_Title") = "Dec'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "REVISE_SECOND_HALF"
        dRow("Column_Title") = "Forecast 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_SECOND_HALF"
        dRow("Column_Title") = "Diff 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "REVISE_TOTAL_YEAR"
        dRow("Column_Title") = "Forecast Year'" & strYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "ORIGINAL_FULL_YEAR"
        dRow("Column_Title") = "Original Year'" & strYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_TOTAL_YEAR"
        dRow("Column_Title") = "Diff Year'" & strYear
        dtColumns.Rows.Add(dRow)

        Return True

    End Function

    Private Function InsertMTPColumnData(ByRef dtColumns As DataTable, ByVal strYear As String) As Boolean

        Dim dRow As DataRow
        Dim strHalfYear As String = strYear.Substring(2, 2)

        Dim intYear As Integer
        If strYear = String.Empty Then
            Return False
        Else
            intYear = CInt(strYear)
        End If

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "Group_Header"
        dRow("Column_Title") = "Description"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERSON_IN_CHARGE_NO"
        dRow("Column_Title") = "Person in Charge Section"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERSON_IN_CHARGE_NAME"
        dRow("Column_Title") = ""
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "ORIGINAL_SECOND_HALF"
        dRow("Column_Title") = "Original 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M7"
        dRow("Column_Title") = "Jul'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M8"
        dRow("Column_Title") = "Aug'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M9"
        dRow("Column_Title") = "Sep'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M10"
        dRow("Column_Title") = "Oct'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M11"
        dRow("Column_Title") = "Nov'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "M12"
        dRow("Column_Title") = "Dec'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "REVISE_SECOND_HALF"
        dRow("Column_Title") = "Forecast 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_SECOND_HALF"
        dRow("Column_Title") = "Diff 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "REVISE_TOTAL_YEAR"
        dRow("Column_Title") = "Forecast Year'" & strYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "ORIGINAL_FULL_YEAR"
        dRow("Column_Title") = "Original Year'" & strYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_TOTAL_YEAR"
        dRow("Column_Title") = "Diff Year'" & strYear
        dtColumns.Rows.Add(dRow)

        Return True
    End Function

    Private Function InsertMTPBudgetColumnData(ByRef dtColumns As DataTable, ByVal strYear As String) As Boolean

        Dim dRow As DataRow
        Dim strHalfYear As String = strYear.Substring(2, 2)
        Dim strPrevHalfYear As String = (CInt(strYear) + 1).ToString.Substring(2, 2)

        Dim intYear As Integer
        If strYear = String.Empty Then
            Return False
        Else
            intYear = CInt(strYear)
        End If

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "Group_Header"
        dRow("Column_Title") = "Description"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERSON_IN_CHARGE_NO"
        dRow("Column_Title") = "Person in Charge Section"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERSON_IN_CHARGE_NAME"
        dRow("Column_Title") = ""
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "REVISE_TOTAL_YEAR"
        dRow("Column_Title") = "Original Year'" & CStr(intYear + 1)
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PrevRRT2"
        'dRow("Column_Title") = "MTP" & intYear - 1 & " Year'" & CStr(intYear + 1)
        dRow("Column_Title") = "MBP" & intYear - 0 & " Year'" & CStr(intYear + 1)
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_TOTAL_YEAR"
        dRow("Column_Title") = "Diff Year'" & CStr(intYear + 1)
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "RRT2"
        'dRow("Column_Title") = "MTP" & intYear & " Year'" & CStr(intYear + 2)
        dRow("Column_Title") = "MBP" & (intYear + 1) & " Year'" & CStr(intYear + 2)
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PrevRRT3"
        'dRow("Column_Title") = "MTP" & intYear - 1 & " Year'" & CStr(intYear + 2)
        dRow("Column_Title") = "MBP" & intYear - 0 & " Year'" & CStr(intYear + 2)
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_PREV_YEAR"
        dRow("Column_Title") = "Diff Year'" & CStr(intYear + 2)
        dtColumns.Rows.Add(dRow)


        dRow = dtColumns.NewRow
        dRow("Column_Name") = "RRT3"
        'dRow("Column_Title") = "MTP" & intYear & " Year'" & CStr(intYear + 3)
        dRow("Column_Title") = "MTP" & (intYear + 1) & " Year'" & CStr(intYear + 3)
        dtColumns.Rows.Add(dRow)

        'dRow = dtColumns.NewRow
        'dRow("Column_Name") = "PrevRRT3"
        'dRow("Column_Title") = "MTP" & intYear - 1 & " Year'" & CStr(intYear + 2)
        'dtColumns.Rows.Add(dRow)

        'dRow = dtColumns.NewRow
        'dRow("Column_Name") = "RRT3"
        'dRow("Column_Title") = "MTP" & intYear & " Year'" & CStr(intYear + 3)
        'dtColumns.Rows.Add(dRow)

        'dRow = dtColumns.NewRow
        'dRow("Column_Name") = "PrevRRT4"
        'dRow("Column_Title") = "MTP" & intYear - 1 & " Year'" & CStr(intYear + 3)
        'dtColumns.Rows.Add(dRow)

        'dRow = dtColumns.NewRow
        'dRow("Column_Name") = "RRT4"
        'dRow("Column_Title") = "MTP" & intYear & " Year'" & CStr(intYear + 4)
        'dtColumns.Rows.Add(dRow)

        'dRow = dtColumns.NewRow
        'dRow("Column_Name") = "PrevRRT5"
        'dRow("Column_Title") = "MTP" & intYear - 1 & " Year'" & CStr(intYear + 4)
        'dtColumns.Rows.Add(dRow)


        'dRow = dtColumns.NewRow
        'dRow("Column_Name") = "RRT5"
        'dRow("Column_Title") = "MTP" & intYear & " Year'" & CStr(intYear + 5)
        'dtColumns.Rows.Add(dRow)

        Return True

    End Function

    Private Function SetupPICSummaryGroupbyData(ByVal dsData As DataSet, ByVal strGroupColumnName As String, _
                                                ByVal strGroupColumnTitle As String, ByVal intDataColumnIndex As Integer, _
                                                ByVal bShowGroupName As Boolean, ByRef intGroupFirstIndex As Integer, ByRef intGroupSecondIndex As Integer) As DataSet

        Dim dsResult As DataSet = New DataSet
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object
        Dim drEmpty As DataRow

        Dim strScript As String = strGroupColumnName

        Dim dtGroups As DataTable = dsData.Tables(0).DefaultView.ToTable(True, strScript)
        Dim intGroupCount As Integer = dtGroups.Rows.Count

        Dim dtResult As DataTable = dsData.Tables(0).Clone
        Dim dtTmp As DataTable = dsData.Tables(0).Clone

        Dim drFixcostTotal As DataRow = dtResult.NewRow
        Dim drVariablecostTotal As DataRow = dtResult.NewRow
        Dim drAllTotal As DataRow = dtResult.NewRow

        For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName

            strExpression = "Sum(" + strColumnName + ")"
            returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
            drAllTotal(dsData.Tables(0).Columns(k).ColumnName) = returnValue

            If strColumnName.IndexOf("FIXCOST") < 0 And strColumnName.IndexOf("VARIABLECOST") < 0 Then

                If strColumnName.IndexOf("SUM") > 0 And strColumnName <> "LAST_YEAR_SUM" Then
                    Dim intSumIndex As Integer = strColumnName.IndexOf("SUM")
                    strColumnName = strColumnName.Substring(0, intSumIndex - 1)
                End If

                strExpression = "Sum(" + strColumnName + "_FIXCOST)"
                returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
                drFixcostTotal(dsData.Tables(0).Columns(k).ColumnName) = returnValue

                strExpression = "Sum(" + strColumnName + "_VARIABLECOST)"
                returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
                drVariablecostTotal(dsData.Tables(0).Columns(k).ColumnName) = returnValue
            End If
        Next

        Dim intGroupTotalIndex As Integer = 0
        For i As Integer = 0 To intGroupCount - 1

            '//Seperate dataset data into several datatables according to group no
            strScript = strGroupColumnName + " = " + dtGroups.Rows(i)(0).ToString
            Dim arrRows As DataRow() = dsData.Tables(0).Select(strScript)

            dtTmp.Rows.Clear()
            For j As Integer = 0 To arrRows.Length - 1

                Dim drow(dtResult.Columns.Count - 1) As Object
                Dim dRowTmp(dtResult.Columns.Count - 1) As Object

                arrRows(j).ItemArray.CopyTo(drow, 0)
                arrRows(j).ItemArray.CopyTo(dRowTmp, 0)

                dtResult.Rows.Add(drow)
                dtTmp.Rows.Add(dRowTmp)
            Next

            '//Calculate total for each group
            Dim drTotal As DataRow = dtResult.NewRow
            For k As Integer = intDataColumnIndex To dtResult.Columns.Count - 1

                Dim strColumnName As String = dtResult.Columns(k).ColumnName
                strExpression = "Sum(" + strColumnName + ")"
                returnValue = dtTmp.Compute(strExpression, strFilter)
                drTotal(dtResult.Columns(k).ColumnName) = returnValue

            Next

            '//Add total cost
            dtResult.Rows.InsertAt(drTotal, intGroupTotalIndex)

            '//Add one empty row
            drEmpty = dtResult.NewRow
            dtResult.Rows.Add(drEmpty)

            If bShowGroupName = True Then
                dtResult.TableName = arrRows(0)(strGroupColumnName).ToString & " " & arrRows(0)(strGroupColumnTitle).ToString
            End If

            intGroupTotalIndex = dtResult.Rows.Count
            If i = 0 Then
                intGroupFirstIndex = intGroupTotalIndex - 1
            Else
                intGroupSecondIndex = intGroupTotalIndex - 1
            End If
        Next

        'intGroupLineIndex = intGroupTotalIndex - 1

        '//Add one empty row
        drEmpty = dtResult.NewRow
        dtResult.Rows.Add(drEmpty)

        '//Add variable cost total
        dtResult.Rows.Add(drVariablecostTotal)

        '//Add fixed cost total
        dtResult.Rows.Add(drFixcostTotal)

        '//Add All total cost
        dtResult.Rows.Add(drAllTotal)

        '//Add one empty row
        drEmpty = dtResult.NewRow
        dtResult.Rows.Add(drEmpty)

        '//Add one more column for group name
        Dim col As DataColumn = New DataColumn()
        col.ColumnName = "Group_Header"
        col.DataType = Type.GetType("System.String")
        dtResult.Columns.Add(col)

        dtResult.Rows(0)("Group_Header") = "Person in Charge"
        If intGroupCount > 1 Then
            dtResult.Rows(intGroupFirstIndex + 1)("Group_Header") = "Others"
        End If

        '//Return data table
        dsResult.Tables.Add(dtResult)
        Return dsResult

    End Function

    Private Sub numProjectNo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numProjectNo.ValueChanged

        LoadRevNo()
        'LoadPrevRevNo()
    End Sub

    Private Sub numYear_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numYear.ValueChanged

        LoadRevNo()

    End Sub

#End Region
End Class