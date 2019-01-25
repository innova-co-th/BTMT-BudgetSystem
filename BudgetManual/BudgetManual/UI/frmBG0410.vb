Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports Microsoft.Office.Interop
Imports System.Drawing.Printing

Public Class frmBG0410

#Region "Variable"
    Private myClsBG0410BL As New clsBG0410BL
    Private myClsBG0310BL As New clsBG0310BL
    Private clsBG0400 As frmBG0400
    Private excelApp As Excel.Application
    'Private dsRPT001 As DataSet
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
            Me.chkShowMTP.Checked = False

            LoadPeriodType()
            'cboPeriodType.Items.Clear()
            'cboPeriodType.Items.Add("Original Budget")
            'cboPeriodType.Items.Add("Estimate Budget")
            'cboPeriodType.Items.Add("Forecast Budget")
            'cboPeriodType.SelectedIndex = 0

            If p_intUserLevelId < enumUserLevel.GeneralManager Then
                If myClsBG0410BL.GetPersonInChargeList() = False Then
                    cboUserPIC.DataSource = Nothing
                Else
                    Dim dt As DataTable = myClsBG0410BL.PersonInCharge
                    Dim dr As DataRow = dt.NewRow
                    dr(0) = 0
                    dr(1) = ""
                    dr(6) = "All"
                    dt.Rows.InsertAt(dr, 0)

                    cboUserPIC.DataSource = myClsBG0410BL.PersonInCharge
                    cboUserPIC.DisplayMember = "PIC_NAME"
                    cboUserPIC.ValueMember = "PERSON_IN_CHARGE_NO"
                    cboUserPIC.SelectedIndex = 0
                End If
            Else
                myClsBG0410BL.PIC = p_strUserPIC
                If myClsBG0410BL.GetPersonInChargeList2() = False Then
                    cboUserPIC.DataSource = Nothing
                Else
                    cboUserPIC.DataSource = myClsBG0410BL.PersonInCharge
                    cboUserPIC.DisplayMember = "PIC_NAME"
                    cboUserPIC.ValueMember = "PERSON_IN_CHARGE_NO"
                    cboUserPIC.SelectedIndex = 0
                End If
            End If

            If p_intUserLevelId = enumUserLevel.SystemAdministrator Then
                Me.lblRevNo.Visible = True
                Me.cboRevNo.Visible = True
                LoadRevNo()

                'Me.lblPrevRevNo.Visible = True
                'Me.cboPrevRevno.Visible = True
                'LoadPrevRevNo()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
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

    Private Sub frmBG0410_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        If Not excelApp Is Nothing Then
            excelApp.Quit()
            excelApp = Nothing
        End If
    End Sub

    Private Sub frmBG0410_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not excelApp Is Nothing Then
            excelApp.Quit()
            excelApp = Nothing
        End If
    End Sub

    Private Sub frmBG0410_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        InitPage()
    End Sub

    Private Sub cmdPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPreview.Click

        Try

            If fncCheckRevNo() = False Then

                MessageBox.Show("No budget data found, please try it again.", "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            'If fncCheckPrevRevNo() = False Then
            '    MessageBox.Show("No previous budget data found, please try it again.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            '    Exit Sub
            'End If

            Me.Cursor = Cursors.WaitCursor

            If clsBG0400 IsNot Nothing Then
                clsBG0400.Close()
                clsBG0400.Dispose()
            End If
            clsBG0400 = New frmBG0400()
            myClsBG0410BL.BudgetYear = Me.numYear.Value.ToString
            myClsBG0410BL.PeriodType = (Me.cboPeriodType.SelectedValue).ToString
            myClsBG0410BL.PIC = Me.cboUserPIC.SelectedValue.ToString
            myClsBG0410BL.MTPChecked = Me.chkShowMTP.Checked
            myClsBG0410BL.ProjectNo = Me.numProjectNo.Value.ToString
            myClsBG0410BL.UserLevelId = p_intUserLevelId
            If Me.cboRevNo.DataSource IsNot Nothing Then
                myClsBG0410BL.RevNo = Me.cboRevNo.SelectedValue.ToString
            End If

            myClsBG0410BL.PrevProjectNo = Me.numPrevProjectNo.Value.ToString
            If Me.cboPrevRevno.DataSource IsNot Nothing AndAlso _
                Me.cboPrevRevno.SelectedValue IsNot Nothing Then
                myClsBG0410BL.PrevRevNo = Me.cboPrevRevno.SelectedValue.ToString
            Else
                myClsBG0410BL.PrevRevNo = String.Empty
            End If

            myClsBG0410BL.ShowZeroValue = Me.chkShowZeroValue.Checked

            If myClsBG0410BL.GetBudgetData() = False Then
                clsBG0400.DS = Nothing
            Else
                clsBG0400.DS = myClsBG0410BL.BudgetData
                If clsBG0400.DS Is Nothing Or clsBG0400.DS.Tables(0).Rows.Count = 0 Then
                    MessageBox.Show("No budget data found, please try it again.", "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Me.Cursor = Cursors.Default
                    Return
                End If
            End If

            myClsBG0410BL.GetBudgetStatus()


            Dim strPeriod As String = String.Empty
            Select Case CInt(Me.cboPeriodType.SelectedValue)
                Case CType(enumPeriodType.OriginalBudget, Integer)
                    clsBG0400.ReportName = "RPT001-1.rpt"
                    strPeriod = "Original"
                    Exit Select
                Case CType(enumPeriodType.EstimateBudget, Integer)
                    clsBG0400.ReportName = "RPT001-2.rpt"
                    strPeriod = "Estimate"
                    Exit Select
                Case CType(enumPeriodType.ForecastBudget, Integer)
                    If Me.chkShowMTP.Checked = True Then
                        clsBG0400.ReportName = "RPT001-4.rpt"
                    Else
                        clsBG0400.ReportName = "RPT001-3.rpt"
                    End If
                    strPeriod = "Forecast"
                    Exit Select
                Case CType(enumPeriodType.MTPBudget, Integer)
                    clsBG0400.ReportName = "RPT001-5.rpt"
                    strPeriod = "MTP"
                    Exit Select
            End Select

            'clsBG0400.ConfigureCrystalReports()
            clsBG0400.PIC = Me.cboUserPIC.Text.ToString
            clsBG0400.BudgetYear = Me.numYear.Value.ToString
            'clsBG0400.ParamPersonInCharge = True
            clsBG0400.Period = strPeriod
            clsBG0400.ReportType = "DetailByPersonInCharge"
            clsBG0400.BudgetStatus = myClsBG0410BL.BudgetStatus
            clsBG0400.ProjectNo = Me.numProjectNo.Value.ToString

            clsBG0400.MdiParent = p_frmBG0010
            clsBG0400.Show()

            If clsBG0400.WindowState = FormWindowState.Minimized Then
                clsBG0400.WindowState = FormWindowState.Normal
            End If
            clsBG0400.BringToFront()
            Me.Cursor = Cursors.Default

        Catch ex As Exception
            MessageBox.Show(ex.Message, "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.Cursor = Cursors.Default
            Return
        End Try

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
                Select Case CInt(Me.cboPeriodType.SelectedValue)
                    Case CType(enumPeriodType.OriginalBudget, Integer)
                        reportPath = p_strAppPath & "\Reports\RPT001-1.rpt"
                        Exit Select
                    Case CType(enumPeriodType.EstimateBudget, Integer)
                        reportPath = p_strAppPath & "\Reports\RPT001-2.rpt"
                        Exit Select
                    Case CType(enumPeriodType.ForecastBudget, Integer)
                        If Me.chkShowMTP.Checked = True Then
                            reportPath = p_strAppPath & "\Reports\RPT001-4.rpt"
                        Else
                            reportPath = p_strAppPath & "\Reports\RPT001-3.rpt"
                        End If
                        Exit Select
                    Case CType(enumPeriodType.MTPBudget, Integer)
                        reportPath = p_strAppPath & "\Reports\RPT001-5.rpt"
                        Exit Select
                    Case Else
                        reportPath = p_strAppPath & "\Reports\RPT001-1.rpt"
                        Exit Select
                End Select

                m_Report.Load(reportPath)

                myClsBG0410BL.BudgetYear = Me.numYear.Value.ToString
                myClsBG0410BL.PeriodType = (Me.cboPeriodType.SelectedValue).ToString
                myClsBG0410BL.PIC = Me.cboUserPIC.SelectedValue.ToString
                myClsBG0410BL.MTPChecked = Me.chkShowMTP.Checked
                myClsBG0410BL.ProjectNo = Me.numProjectNo.Value.ToString
                myClsBG0410BL.UserLevelId = p_intUserLevelId
                If Me.cboRevNo.DataSource IsNot Nothing Then
                    myClsBG0410BL.RevNo = Me.cboRevNo.SelectedValue.ToString
                End If

                myClsBG0410BL.PrevProjectNo = Me.numPrevProjectNo.Value.ToString
                If Me.cboPrevRevno.DataSource IsNot Nothing AndAlso _
                    Me.cboPrevRevno.SelectedValue IsNot Nothing Then
                    myClsBG0410BL.PrevRevNo = Me.cboPrevRevno.SelectedValue.ToString
                Else
                    myClsBG0410BL.PrevRevNo = String.Empty
                End If

                myClsBG0410BL.ShowZeroValue = Me.chkShowZeroValue.Checked

                Dim ds As DataSet
                If myClsBG0410BL.GetBudgetData() = False Then
                    ds = Nothing
                Else
                    ds = myClsBG0410BL.BudgetData
                    If ds Is Nothing Or ds.Tables(0).Rows.Count = 0 Then
                        MessageBox.Show("No budget data found, please try it again.", "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Me.Cursor = Cursors.Default
                        Return
                    End If
                End If
                m_Report.SetDataSource(ds)

                myClsBG0410BL.GetBudgetStatus()

                If myClsBG0410BL.BudgetStatus >= 5 Then
                    m_Report.ReportDefinition.ReportObjects("picAuth10").ObjectFormat.EnableSuppress = False
                Else
                    m_Report.ReportDefinition.ReportObjects("picAuth10").ObjectFormat.EnableSuppress = True
                End If

                If myClsBG0410BL.BudgetStatus >= 6 Then
                    m_Report.ReportDefinition.ReportObjects("picAuth3").ObjectFormat.EnableSuppress = False
                Else
                    m_Report.ReportDefinition.ReportObjects("picAuth3").ObjectFormat.EnableSuppress = True
                End If

                Dim strPeriod As String = String.Empty
                Select Case CInt(Me.cboPeriodType.SelectedValue)
                    Case CType(enumPeriodType.OriginalBudget, Integer)
                        strPeriod = "Original"
                    Case CType(enumPeriodType.EstimateBudget, Integer)
                        strPeriod = "Estimate"
                    Case CType(enumPeriodType.ForecastBudget, Integer)
                        strPeriod = "Forecast"
                    Case CType(enumPeriodType.MTPBudget, Integer)
                        strPeriod = "MTP"
                End Select

                m_Report.SetParameterValue("PERSON_IN_CHARGE_NM", Me.cboUserPIC.Text.ToString)
                m_Report.SetParameterValue("BUDGET_YEAR", Me.numYear.Value.ToString)
                m_Report.SetParameterValue("HALF_BUDGET_YEAR", Me.numYear.Value.ToString.Substring(2, 2))
                m_Report.SetParameterValue("PERIOD", strPeriod)
                m_Report.SetParameterValue("PROJECT_NO", Me.numProjectNo.Value.ToString)

                m_Report.PrintOptions.PrinterName = PrintDialog1.PrinterSettings.PrinterName

                '  m_Report.PrintOptions.PaperSize = PaperSize.PaperA4

                m_Report.PrintToPrinter(PrintDialog1.PrinterSettings.Copies, _
                                        PrintDialog1.PrinterSettings.Collate, _
                                        PrintDialog1.PrinterSettings.FromPage, _
                                        PrintDialog1.PrinterSettings.ToPage)
                'Dim pt As Printing.PrintDocument

                Me.Cursor = Cursors.Default
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.Cursor = Cursors.Default
            Return
        End Try

    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub cmdExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExcel.Click

        If fncCheckRevNo() = False Then

            MessageBox.Show("No budget data found, please try it again.", "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor

        '//Get Export Data
        Dim dsData As DataSet
        myClsBG0410BL.BudgetYear = Me.numYear.Value.ToString
        myClsBG0410BL.PeriodType = (Me.cboPeriodType.SelectedValue).ToString
        myClsBG0410BL.PIC = Me.cboUserPIC.SelectedValue.ToString
        myClsBG0410BL.MTPChecked = Me.chkShowMTP.Checked
        myClsBG0410BL.ProjectNo = Me.numProjectNo.Value.ToString
        myClsBG0410BL.UserLevelId = p_intUserLevelId
        If Me.cboRevNo.DataSource IsNot Nothing Then
            myClsBG0410BL.RevNo = Me.cboRevNo.SelectedValue.ToString
        End If

        myClsBG0410BL.PrevProjectNo = Me.numPrevProjectNo.Value.ToString
        If Me.cboPrevRevno.DataSource IsNot Nothing AndAlso _
            Me.cboPrevRevno.SelectedValue IsNot Nothing Then
            myClsBG0410BL.PrevRevNo = Me.cboPrevRevno.SelectedValue.ToString
        Else
            myClsBG0410BL.PrevRevNo = String.Empty
        End If

        Dim strPeriod As String = cboPeriodType.Text
        strPeriod = strPeriod.Substring(0, strPeriod.IndexOf("Budget") - 1)

        myClsBG0410BL.ShowZeroValue = Me.chkShowZeroValue.Checked

        If myClsBG0410BL.GetBudgetData() = False Then
            dsData = Nothing
            MessageBox.Show("No budget data found, please try it again.", "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Cursor = Cursors.Default
            Return
        Else
            dsData = myClsBG0410BL.BudgetData
            If dsData Is Nothing Or dsData.Tables(0).Rows.Count = 0 Then
                MessageBox.Show("No budget data found, please try it again.", "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Cursor = Cursors.Default
                Return
            End If
        End If

        Dim dtAuthorizeImages As DataTable = myClsBG0410BL.BudgetData.Tables(1)
        Dim strYear As String = Me.numYear.Value.ToString
        Dim dtColumns As DataTable = CreateTableTemplate()

        Dim strPeriodType As String = cboPeriodType.Text
        Dim strProjectNo As String = Me.numProjectNo.Value.ToString

        Dim strSubTitle As String = String.Empty

        If strProjectNo <> "1" Then
            strSubTitle = "Detail by Person In Charge : " + strPeriodType + " " + strYear + " (Project No. " + strProjectNo + ")"
        Else
            strSubTitle = "Detail by Person In Charge : " + strPeriodType + " " + strYear
        End If

        Select Case CInt(Me.cboPeriodType.SelectedValue)

            Case CType(enumPeriodType.OriginalBudget, Integer) '//Original

                InsertOriginalColumnData(dtColumns, strYear)

                '//Create group data
                Dim dsGroups As DataSet = SetupGroupbyData(dsData, "PERSON_IN_CHARGE", "PERSON_IN_CHARGE_NAME", 10, True)

                '//Create Output Excel
                OutputExcel(dsGroups, dtColumns, False, strSubTitle, strYear, True, strPeriod)

            Case CType(enumPeriodType.EstimateBudget, Integer) '//Estimate

                InsertEstimateColumnData(dtColumns, strYear)

                '//Create group data
                Dim dsGroups As DataSet = SetupGroupbyData(dsData, "PERSON_IN_CHARGE_NO", "PERSON_IN_CHARGE_NAME", 10, True)

                '//Create Output Excel
                OutputExcel(dsGroups, dtColumns, False, strSubTitle, strYear, True, strPeriod)

            Case CType(enumPeriodType.ForecastBudget, Integer) '//Forecast

                '//Create output columns
                If Me.chkShowMTP.Checked = True Then
                    InsertMTPColumnData(dtColumns, strYear)
                Else
                    InsertForecastColumnData(dtColumns, strYear)
                End If

                '//Create group data
                Dim dsGroups As DataSet = SetupGroupbyData(dsData, "PERSON_IN_CHARGE_NO", "PERSON_IN_CHARGE_NAME", 10, True)

                Dim bMTPCheck As Boolean = chkShowMTP.Checked
                '//Create Output Excel
                OutputExcel(dsGroups, dtColumns, bMTPCheck, strSubTitle, strYear, True, strPeriod)

            Case CType(enumPeriodType.MTPBudget, Integer) '//MTP

                InsertMTPColumnDataNew(dtColumns, strYear)

                '//Create group data
                Dim dsGroups As DataSet = SetupGroupbyData(dsData, "PERSON_IN_CHARGE_NO", "PERSON_IN_CHARGE_NAME", 11, True)

                '  Dim bMTPCheck As Boolean = chkShowMTP.Checked
                '//Create Output Excel
                OutputExcel(dsGroups, dtColumns, False, strSubTitle, strYear, True, strPeriod)


        End Select

        Me.Cursor = Cursors.Default

    End Sub

    Private Function InsertOriginalColumnData(ByRef dtColumns As DataTable, ByVal strYear As String) As Boolean

        Dim dRow As DataRow
        Dim strHalfYear As String = strYear.Substring(2, 2)

        Dim intYear As Integer = CInt(strYear)
        Dim strLastYear As String = CStr(intYear - 1)

        Dim strHalfLastYear As String = CStr(intYear - 1).Substring(2, 2)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "BUDGET_ORDER_NO"
        dRow("Column_Title") = "Budget Order Number & "
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "BUDGET_ORDER_NAME"
        dRow("Column_Title") = "Budget Name"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DEPT_NO"
        dRow("Column_Title") = "Dept."
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERSON_IN_CHARGE"
        dRow("Column_Title") = "Person in Charge"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "ACTUAL_FIRST_HALF"
        dRow("Column_Title") = "Actual 1st Half'" & strHalfLastYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "Forecast_SECOND_HALF"
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

        dRow = dtColumns.NewRow()
        dRow("Column_Name") = "MTP_RRT1"
        dRow("Column_Title") = "MTP " & CInt(strYear) - 2 & " Year'" & strYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow()
        dRow("Column_Name") = "DIFF_MTP"
        dRow("Column_Title") = "Diff vs MTP" & CInt(strYear) - 2
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "LAST_YEAR_SUM"
        dRow("Column_Title") = "Total Year'" & strLastYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_YEAR"
        dRow("Column_Title") = "Diff vs Year'" & CInt(strYear) - 1
        dtColumns.Rows.Add(dRow)

        Return True

    End Function

    Private Function InsertEstimateColumnData(ByRef dtColumns As DataTable, ByVal strYear As String) As Boolean

        Dim dRow As DataRow
        Dim strHalfYear As String = strYear.Substring(2, 2)

        Dim intYear As Integer = CInt(strYear)
        Dim strLastYear As String = CStr(intYear - 1)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "BUDGET_ORDER_NO"
        dRow("Column_Title") = "Budget Order Number & "
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "BUDGET_ORDER_NAME"
        dRow("Column_Title") = "Budget Name"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DEPT_NO"
        dRow("Column_Title") = "Dept."
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERSON_IN_CHARGE_NO"
        dRow("Column_Title") = "Person in Charge"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "ACTUAL_FIRST_HALF"
        dRow("Column_Title") = "Actual 1st Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "Forecast_SECOND_HALF"
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

        Return True

    End Function

    Private Function InsertForecastColumnData(ByRef dtColumns As DataTable, ByVal strYear As String) As Boolean

        Dim dRow As DataRow
        Dim strHalfYear As String = strYear.Substring(2, 2)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "BUDGET_ORDER_NO"
        dRow("Column_Title") = "Budget Order Number & "
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "BUDGET_ORDER_NAME"
        dRow("Column_Title") = "Budget Name"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DEPT_NO"
        dRow("Column_Title") = "Dept."
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERSON_IN_CHARGE_NO"
        dRow("Column_Title") = "Person in Charge"
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
        dRow("Column_Title") = "Estimate 1st Half'" & strHalfYear
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
        dRow("Column_Name") = "Forecast_SECOND_HALF"
        dRow("Column_Title") = "Forecast 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_SECOND_HALF"
        dRow("Column_Title") = "Diff 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "Forecast_TOTAL_YEAR"
        dRow("Column_Title") = "Forecast Year'" & strYear
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
        dRow("Column_Name") = "BUDGET_ORDER_NO"
        dRow("Column_Title") = "Budget Order Number & "
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "BUDGET_ORDER_NAME"
        dRow("Column_Title") = "Budget Name"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DEPT_NO"
        dRow("Column_Title") = "Dept."
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERSON_IN_CHARGE_NO"
        dRow("Column_Title") = "Person in Charge"
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
        dRow("Column_Name") = "Forecast_SECOND_HALF"
        dRow("Column_Title") = "Forecast 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_SECOND_HALF"
        dRow("Column_Title") = "Diff 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "Forecast_TOTAL_YEAR"
        dRow("Column_Title") = "Forecast Year'" & strYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_TOTAL_YEAR"
        dRow("Column_Title") = "Diff Year'" & strYear
        dtColumns.Rows.Add(dRow)

        Return True

    End Function

    Private Function InsertMTPColumnDataNew(ByRef dtColumns As DataTable, ByVal strYear As String) As Boolean

        Dim dRow As DataRow
        Dim strHalfYear As String = strYear.Substring(2, 2)

        Dim intYear As Integer
        If strYear = String.Empty Then
            Return False
        Else
            intYear = CInt(strYear)
        End If

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "BUDGET_ORDER_NO"
        dRow("Column_Title") = "Budget Order Number & "
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "BUDGET_ORDER_NAME"
        dRow("Column_Title") = "Budget Name"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DEPT_NO"
        dRow("Column_Title") = "Dept."
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERSON_IN_CHARGE_NO"
        dRow("Column_Title") = "Person in Charge"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "Forecast_TOTAL_YEAR"
        dRow("Column_Title") = "Original Year'" & CStr(intYear + 1)
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PrevRRT2"
        dRow("Column_Title") = "MTP" & intYear - 1 & " Year'" & CStr(intYear + 1)
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_TOTAL_YEAR"
        dRow("Column_Title") = "Diff Year'" & CStr(intYear + 1)
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "RRT2"
        dRow("Column_Title") = "MTP" & intYear & " Year'" & CStr(intYear + 2)
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PrevRRT3"
        dRow("Column_Title") = "MTP" & intYear - 1 & " Year'" & CStr(intYear + 2)
        dtColumns.Rows.Add(dRow)


        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_PREV_YEAR"
        dRow("Column_Title") = "Diff" & " Year'" & CStr(intYear + 2)
        dtColumns.Rows.Add(dRow)


        dRow = dtColumns.NewRow
        dRow("Column_Name") = "RRT3"
        dRow("Column_Title") = "MTP" & intYear & " Year'" & CStr(intYear + 3)
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

    Private Function OutputExcel(ByVal dsData As DataSet, ByVal dtColumns As DataTable, ByVal bMTPCheck As Boolean, _
                                 ByVal strSubTitle As String, ByVal strYear As String, ByVal bShowGroupName As Boolean, _
                                 ByVal strPeriod As String) As Boolean

        If excelApp Is Nothing Then
            excelApp = New Excel.Application
        End If

        Dim rowStartIndex As Integer
        Dim colStartIndex As Integer

        Dim xBk As Excel.Workbook = Nothing
        Dim xSt As Excel.Worksheet = Nothing

        Dim strHalfYear = strYear.Substring(2, 2)
        xBk = excelApp.Workbooks.Add(System.Reflection.Missing.Value)
        If xBk.Worksheets.Count > 1 Then
            For i As Integer = 1 To xBk.Worksheets.Count - 1
                CType(xBk.Worksheets(i), Excel.Worksheet).Delete()
            Next
        End If

        '//Set Style Value < 0 please fill color "Red"
        Dim style As Excel.Style = excelApp.ActiveWorkbook.Styles.Add("NewStyle")
        'style.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
        style.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)

        For intSheetCount As Integer = 0 To dsData.Tables.Count - 1

            If intSheetCount <> 0 Then
                xBk.Sheets.Add()
            End If

            rowStartIndex = 10
            colStartIndex = 9

            xSt = CType(xBk.ActiveSheet, Excel.Worksheet)
            Dim strTableName As String = dsData.Tables(intSheetCount).TableName
            Dim intIndexTblName As Integer = strTableName.IndexOf(" ")
            xSt.Name = strTableName.Substring(0, intIndexTblName)

            '//Setup DataColumn
            For i As Integer = 0 To dtColumns.Rows.Count - 1
                xSt.Cells(colStartIndex, i + 1) = dtColumns.Rows(i)("Column_Title").ToString
                xSt.Range(xSt.Cells(colStartIndex, i + 1), xSt.Cells(colStartIndex, i + 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            Next

            Dim arrCols() As Integer

            If strPeriod = "Original" Then

                arrCols = New Integer() {3, 4, 5, 6, 13, 20, 21, 22, 23, 24, 25}
                SetupOriginalColumnsCells(xSt, colStartIndex, 1, 2, "Budget Order Number & Budget Name", arrCols, 7, 12, strYear, True, True, 14, 19)

            ElseIf strPeriod = "Estimate" Then

                arrCols = New Integer() {3, 4, 5, 6, 13, 14, 15}
                SetupEstimateColumnsCells(xSt, colStartIndex, 1, 2, "Budget Order Number & Budget Name", arrCols, 7, 9, 10, 12)

            ElseIf strPeriod = "Forecast" Then
                arrCols = New Integer() {3, 4, 5, 12, 13, 14, 21, 22, 23, 24} '// Two Row Merge Col

                If bMTPCheck = True Then
                    arrCols = New Integer() {3, 4, 5, 12, 13, 14, 15} '// Two Row Merge Col
                End If

                SetupForecastColumnsCells(xSt, colStartIndex, bMTPCheck, 1, 2, "Budget Order Number & Budget Name", _
                                        arrCols, 6, 8, 9, 11, 15, 20, 6, 11, 25, 29)

            ElseIf strPeriod = "MTP" Then
                arrCols = New Integer() {3, 4, 5, 6, 7} '// Two Row Merge Col
                SetupMTPColumnsCells(xSt, colStartIndex, 1, 2, "Budget Order Number & Budget Name", _
                                        arrCols, 8, 11)
            End If

            '//Setup Data
            For rowIndex As Integer = 0 To dsData.Tables(intSheetCount).Rows.Count - 1

                Dim row As DataRow = dsData.Tables(intSheetCount).Rows(rowIndex)

                For colIndex As Integer = 0 To dtColumns.Rows.Count - 1

                    Dim strColumnName As String = dtColumns.Rows(colIndex)("Column_Name").ToString
                    Dim col As DataColumn = dsData.Tables(intSheetCount).Columns(strColumnName)

                    If col.DataType Is System.Type.GetType("System.DateTime") Then

                        xSt.Cells(rowIndex + rowStartIndex, colIndex + 1) = (Convert.ToDateTime(row(col.ColumnName).ToString())).ToString("yyyy-MM-dd")
                        xSt.Range(xSt.Cells(rowIndex + rowStartIndex, colIndex + 1), xSt.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter

                    ElseIf col.DataType Is System.Type.GetType("System.String") Then

                        xSt.Cells(rowIndex + rowStartIndex, colIndex + 1) = "'" + row(col.ColumnName).ToString()
                        xSt.Range(xSt.Cells(rowIndex + rowStartIndex, colIndex + 1), xSt.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter

                    ElseIf col.DataType Is System.Type.GetType("System.Decimal") Then

                        If row(col.ColumnName).ToString = String.Empty Then
                            row(col.ColumnName) = "0.00"
                        End If

                        '//Set Style Value < 0 please fill color "Red"
                        If CDec(row(col.ColumnName)) < 0 Then
                            xSt.Range(xSt.Cells(rowIndex + rowStartIndex, colIndex + 1), xSt.Cells(rowIndex + rowStartIndex, colIndex + 1)).Style = style
                        End If
                        '//End Add by Max 01/10/2012

                        xSt.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName)
                        xSt.Range(xSt.Cells(rowIndex + rowStartIndex, colIndex + 1), xSt.Cells(rowIndex + rowStartIndex, colIndex + 1)).NumberFormat = "#,##0.00"

                    Else
                        xSt.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName).ToString()
                    End If
                Next
            Next

            '//Add Title 
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
            xSt.Range(xSt.Cells(rowStartIndex, 2), xSt.Cells(rowMax, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup Total Lines
            SetupTotalLines(xSt, rowMax - 3, "Total", "Center", 1, 1, 4, rowMax - 2, rowMax - 1, rowMax, colMax, bMTPCheck)

            '//Setup sheet properly width
            xSt.Range(xSt.Cells(2, 1), xSt.Cells(rowMax, colMax)).Columns.AutoFit()

            '//Setup Wrap text for columns title
            If strPeriod = "Original" Then

                xSt.Range(xSt.Cells(2, 4), xSt.Cells(rowMax, 4)).Columns.ColumnWidth = 9
                xSt.Range(xSt.Cells(2, 4), xSt.Cells(rowMax, 4)).WrapText = True

                xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 6)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 6)).WrapText = True

                xSt.Range(xSt.Cells(2, 7), xSt.Cells(rowMax, 12)).Columns.ColumnWidth = 12

                xSt.Range(xSt.Cells(2, 14), xSt.Cells(rowMax, 19)).Columns.ColumnWidth = 12

                xSt.Range(xSt.Cells(2, 13), xSt.Cells(rowMax, 13)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 13), xSt.Cells(rowMax, 13)).WrapText = True

                xSt.Range(xSt.Cells(2, 20), xSt.Cells(rowMax, 25)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 20), xSt.Cells(rowMax, 25)).WrapText = True

            ElseIf strPeriod = "Estimate" Then

                xSt.Range(xSt.Cells(2, 4), xSt.Cells(rowMax, 4)).Columns.ColumnWidth = 9
                xSt.Range(xSt.Cells(2, 4), xSt.Cells(rowMax, 4)).WrapText = True

                xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 6)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 6)).WrapText = True

                xSt.Range(xSt.Cells(2, 7), xSt.Cells(rowMax, 12)).Columns.ColumnWidth = 12

                xSt.Range(xSt.Cells(2, 13), xSt.Cells(rowMax, 15)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 13), xSt.Cells(rowMax, 15)).WrapText = True
            ElseIf strPeriod = "Forecast" Then

                xSt.Range(xSt.Cells(2, 4), xSt.Cells(rowMax, 4)).Columns.ColumnWidth = 9
                xSt.Range(xSt.Cells(2, 4), xSt.Cells(rowMax, 4)).WrapText = True

                xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 5)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 5)).WrapText = True

                xSt.Range(xSt.Cells(2, 6), xSt.Cells(rowMax, 11)).Columns.ColumnWidth = 12

                xSt.Range(xSt.Cells(2, 12), xSt.Cells(rowMax, 14)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 12), xSt.Cells(rowMax, 14)).WrapText = True

                xSt.Range(xSt.Cells(2, 15), xSt.Cells(rowMax, 20)).Columns.ColumnWidth = 12

                xSt.Range(xSt.Cells(2, 21), xSt.Cells(rowMax, 24)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 21), xSt.Cells(rowMax, 24)).WrapText = True

            ElseIf strPeriod = "MTP" Then

                xSt.Range(xSt.Cells(2, 4), xSt.Cells(rowMax, 4)).Columns.ColumnWidth = 9
                xSt.Range(xSt.Cells(2, 4), xSt.Cells(rowMax, 4)).WrapText = True

                xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 7)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 7)).WrapText = True

                xSt.Range(xSt.Cells(2, 8), xSt.Cells(rowMax, 16)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 8), xSt.Cells(rowMax, 16)).WrapText = True
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


            '//Insert empty column
            If strPeriod = "MTP" AndAlso bMTPCheck = True Then
                '   SetupMTPEmptyColumn(xSt, colStartIndex, rowMax, colMax, 25, rowMax - 2, 1)
            ElseIf bMTPCheck = True Then
                'SetupMTPEmptyColumn(xSt, colStartIndex, rowMax, colMax, 25, rowMax - 2, 1)
            End If

            '//-- Add by Max 26/09/2012
            '//Set Frame All
            xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax, colMax)).Borders.LineStyle = 1
            xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium
            xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium
            xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium

            '//Set Frame
            If strPeriod = "Original" Then

                xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax, 6)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 7), xSt.Cells(rowMax, 12)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 12), xSt.Cells(rowMax, 13)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 13), xSt.Cells(rowMax, 19)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 19), xSt.Cells(rowMax, 20)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                'xSt.Range(xSt.Cells(colStartIndex, 15), xSt.Cells(rowMax, 16)).Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium
                'xSt.Range(xSt.Cells(colStartIndex, 17), xSt.Cells(rowMax, 18)).Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium

            ElseIf strPeriod = "Estimate" Then

                xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax, 6)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 7), xSt.Cells(rowMax, 9)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 10), xSt.Cells(rowMax, 12)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 13), xSt.Cells(rowMax, 15)).Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium

            ElseIf strPeriod = "Forecast" Then

                xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax, 5)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                If bMTPCheck = False Then
                    xSt.Range(xSt.Cells(colStartIndex, 6), xSt.Cells(rowMax, 8)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                End If
                xSt.Range(xSt.Cells(colStartIndex, 9), xSt.Cells(rowMax, 11)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 12), xSt.Cells(rowMax, 14)).Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 14), xSt.Cells(rowMax, 14)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium

                If bMTPCheck = False Then
                    xSt.Range(xSt.Cells(colStartIndex, 15), xSt.Cells(rowMax, 20)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                    xSt.Range(xSt.Cells(colStartIndex, 21), xSt.Cells(rowMax, 24)).Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium
                End If

                '//Merge empty line
                If bMTPCheck = True Then
                    xSt.Range(xSt.Cells(rowMax - 2, 1), xSt.Cells(rowMax - 2, colMax)).ClearContents()
                    xSt.Range(xSt.Cells(rowMax - 2, 1), xSt.Cells(rowMax - 2, colMax)).MergeCells = True
                End If

            ElseIf strPeriod = "MTP" Then


                xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax, 6)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 7), xSt.Cells(rowMax, 7)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium

                '//Set font color
                xSt.Range(xSt.Cells(colStartIndex, 6), xSt.Cells(rowMax, 7)).Font.Color = RGB(128, 128, 128)
                xSt.Range(xSt.Cells(colStartIndex, 9), xSt.Cells(rowMax, 9)).Font.Color = RGB(128, 128, 128)
                xSt.Range(xSt.Cells(colStartIndex, 10), xSt.Cells(rowMax, 10)).Font.Color = RGB(128, 128, 128)
                xSt.Range(xSt.Cells(colStartIndex, 13), xSt.Cells(rowMax, 13)).Font.Color = RGB(128, 128, 128)
                xSt.Range(xSt.Cells(colStartIndex, 15), xSt.Cells(rowMax, 15)).Font.Color = RGB(128, 128, 128)

                '//Set Format
                xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax, 6)).NumberFormat = "#,##0.00;[Red]-#,##0.00"
                xSt.Range(xSt.Cells(colStartIndex, 8), xSt.Cells(rowMax, 9)).NumberFormat = "#,##0.00;[Red]-#,##0.00"
                xSt.Range(xSt.Cells(colStartIndex, 11), xSt.Cells(rowMax, colMax)).NumberFormat = "#,##0.00;[Red]-#,##0.00"

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
                intUnitPriceEnd = 25

                intAuthorizeStart = 21
                intAuthorizeEnd = 22

                intImageIndex = 815

            Case "Estimate"

                intUnitPriceStart = 14
                intUnitPriceEnd = 15

                intAuthorizeStart = 11
                intAuthorizeEnd = 13

                intImageIndex = 755

            Case "Forecast"

                If bMTPCheck = True Then
                    intUnitPriceStart = 23
                    intUnitPriceEnd = 24
                Else
                    intUnitPriceStart = 23
                    intUnitPriceEnd = 24
                End If

                If bMTPCheck = True Then
                    intAuthorizeStart = 13
                    intAuthorizeEnd = 14
                Else
                    intAuthorizeStart = 21
                    intAuthorizeEnd = 22
                End If

                If bMTPCheck = True Then
                    intImageIndex = 875
                Else
                    intImageIndex = 1215
                End If

            Case "MTP"

                intUnitPriceStart = 10
                intUnitPriceEnd = 11

                intAuthorizeStart = 1
                intAuthorizeEnd = 2

                intImageIndex = 815

            Case Else
                Exit Select
        End Select

        Return True

    End Function

    Private Sub cboPeriodType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPeriodType.SelectedIndexChanged

        If CInt(cboPeriodType.SelectedValue) = CType(enumPeriodType.ForecastBudget, Integer) Then
            Me.chkShowMTP.Enabled = True
        Else
            Me.chkShowMTP.Checked = False
            Me.chkShowMTP.Enabled = False
        End If

        If CType(cboPeriodType.SelectedValue, Integer) = CType(enumPeriodType.MTPBudget, Integer) Then
            Me.gbPrevYear.Text = "Previous Year"

        ElseIf CType(cboPeriodType.SelectedValue, Integer) = CType(enumPeriodType.OriginalBudget, Integer) Then
            Me.gbPrevYear.Text = "MTP"

        End If

        LoadRevNo()
    End Sub

    Private Sub numYear_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numYear.ValueChanged
        LoadRevNo()
    End Sub

    Private Sub numProjectNo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numProjectNo.ValueChanged

        LoadRevNo()

    End Sub


    Private Sub numPrevProjectNo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numPrevProjectNo.ValueChanged

    End Sub

#End Region

End Class