Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports CrystalDecisions.CrystalReports.Engine
Imports Microsoft.Office.Interop
Imports System.Drawing.Printing

Public Class frmBG0460

#Region "Variable"
    Private myClsBG0460BL As New clsBG0460BL
    'Private myClsBG0410BL As New clsBG0410BL
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

                myClsBG0310BL.PeriodList.DefaultView.RowFilter = "PERIOD_TYPE_ID <> 10"

                cboPeriodType.DisplayMember = "PERIOD_TYPE_NAME"
                cboPeriodType.ValueMember = "PERIOD_TYPE_ID"
                cboPeriodType.DataSource = myClsBG0310BL.PeriodList.DefaultView

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
            
            'If p_intUserLevelId < enumUserLevel.GeneralManager Then
            '    If myClsBG0410BL.GetPersonInChargeList() = False Then
            '        cboUserPIC.DataSource = Nothing
            '    Else
            '        Dim dt As DataTable = myClsBG0410BL.PersonInCharge
            '        Dim dr As DataRow = dt.NewRow
            '        dr(0) = 0
            '        dr(1) = ""
            '        dr(6) = "All"
            '        dt.Rows.InsertAt(dr, 0)

            '        cboUserPIC.DataSource = myClsBG0410BL.PersonInCharge
            '        cboUserPIC.DisplayMember = "PIC_NAME"
            '        cboUserPIC.ValueMember = "PERSON_IN_CHARGE_NO"
            '        cboUserPIC.SelectedIndex = 0
            '    End If
            'Else
            '    myClsBG0410BL.PIC = p_strUserPIC
            '    If myClsBG0410BL.GetPersonInChargeList2() = False Then
            '        cboUserPIC.DataSource = Nothing
            '    Else
            '        Dim dt As DataTable = myClsBG0410BL.PersonInCharge
            '        Dim dr As DataRow = dt.NewRow
            '        dr(0) = 0
            '        dr(1) = ""
            '        dr(6) = "All"
            '        dt.Rows.InsertAt(dr, 0)

            '        cboUserPIC.DataSource = myClsBG0410BL.PersonInCharge
            '        cboUserPIC.DisplayMember = "PIC_NAME"
            '        cboUserPIC.ValueMember = "PERSON_IN_CHARGE_NO"
            '        cboUserPIC.SelectedIndex = 0
            '    End If
            'End If

            LoadPersonInCharge()

            If p_intUserLevelId = enumUserLevel.SystemAdministrator Then
                Me.lblRevNo.Visible = True
                Me.cboRevNo.Visible = True
                LoadRevNo()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
                myClsBG0310BL.BudgetType = BGConstant.P_BUDGET_TYPE_ASSET

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

    Private Sub LoadPersonInCharge()

        If Me.numYear.Value.ToString <> "" AndAlso _
           Me.cboPeriodType.SelectedIndex > -1 Then

            myClsBG0460BL.BudgetYear = Me.numYear.Value.ToString
            myClsBG0460BL.PeriodType = Me.cboPeriodType.SelectedValue.ToString

            If myClsBG0460BL.GetPersonInChargeList() = False Then
                cboUserPIC.DataSource = Nothing
            Else
                Dim dt As DataTable = myClsBG0460BL.PersonInCharge
                Dim dr As DataRow = dt.NewRow
                dr(0) = 0
                dr(1) = "All"
                dt.Rows.InsertAt(dr, 0)

                cboUserPIC.DataSource = myClsBG0460BL.PersonInCharge
                cboUserPIC.DisplayMember = "PIC_NAME"
                cboUserPIC.ValueMember = "PERSON_IN_CHARGE_NO"
                cboUserPIC.SelectedIndex = 0
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

    Private Sub frmBG0460_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        If Not excelApp Is Nothing Then
            excelApp.Quit()
            excelApp = Nothing
        End If
    End Sub

    Private Sub frmBG0460_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not excelApp Is Nothing Then
            excelApp.Quit()
            excelApp = Nothing
        End If
    End Sub

    Private Sub frmBG0460_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        InitPage()
    End Sub

    Private Sub cmdPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPreview.Click

        Try

            If fncCheckRevNo() = False Then

                MessageBox.Show("No budget data found, please try it again.", "RPT006", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            Me.Cursor = Cursors.WaitCursor

            If clsBG0400 IsNot Nothing Then
                clsBG0400.Close()
                clsBG0400.Dispose()
            End If
            clsBG0400 = New frmBG0400()
            myClsBG0460BL.BudgetYear = Me.numYear.Value.ToString
            myClsBG0460BL.PeriodType = (Me.cboPeriodType.SelectedValue).ToString
            myClsBG0460BL.MTPChecked = Me.chkShowMTP.Checked
            myClsBG0460BL.ProjectNo = Me.numProjectNo.Value.ToString
            myClsBG0460BL.UserLevelId = p_intUserLevelId
            If Me.cboRevNo.DataSource IsNot Nothing Then
                myClsBG0460BL.RevNo = Me.cboRevNo.SelectedValue.ToString
            End If
            myClsBG0460BL.PIC = Me.cboUserPIC.SelectedValue.ToString

            If myClsBG0460BL.GetBudgetStatus() = False Then
                MessageBox.Show("Load buget status failed.", "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Cursor = Cursors.Default
                Return
            End If

            If myClsBG0460BL.GetBudgetData() = False Then
                clsBG0400.DS = Nothing
                MessageBox.Show("Load buget data failed.", "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Cursor = Cursors.Default
                Return
            Else
                clsBG0400.DS = myClsBG0460BL.BudgetData
                If clsBG0400.DS Is Nothing Or clsBG0400.DS.Tables(0).Rows.Count = 0 Then
                    MessageBox.Show("No budget data found, please try it again.", "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Me.Cursor = Cursors.Default
                    Return
                End If
            End If

            Select Case CInt(Me.cboPeriodType.SelectedValue)
                Case CType(enumPeriodType.OriginalBudget, Integer)
                    clsBG0400.ReportName = "RPT006-1.rpt"
                    clsBG0400.Period = "Original"
                    Exit Select
                Case CType(enumPeriodType.EstimateBudget, Integer)
                    clsBG0400.ReportName = "RPT006-2.rpt"
                    clsBG0400.Period = "Estimate"
                    Exit Select
                Case CType(enumPeriodType.ReviseBudget, Integer)
                    If Me.chkShowMTP.Checked = True Then
                        clsBG0400.ReportName = "RPT006-4.rpt"
                    Else
                        clsBG0400.ReportName = "RPT006-3.rpt"
                    End If
                    clsBG0400.Period = "Revise"
                    Exit Select
            End Select

            'clsBG0400.ConfigureCrystalReports()
            clsBG0400.BudgetYear = Me.numYear.Value.ToString
            clsBG0400.ReportType = "SummarybyInvestment"
            clsBG0400.BudgetStatus = myClsBG0460BL.BudgetStatus
            clsBG0400.ProjectNo = Me.numProjectNo.Value.ToString

            clsBG0400.MdiParent = p_frmBG0010
            clsBG0400.Show()

            If clsBG0400.WindowState = FormWindowState.Minimized Then
                clsBG0400.WindowState = FormWindowState.Normal
            End If
            clsBG0400.BringToFront()
            Me.Cursor = Cursors.Default

        Catch ex As Exception
            MessageBox.Show(ex.Message, "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Cursor = Cursors.Default
            Return
        End Try

    End Sub

    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrint.Click

        Try

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


            If PrintDialog1.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then

                If fncCheckRevNo() = False Then

                    MessageBox.Show("No budget data found, please try it again.", "RPT006", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                Me.Cursor = Cursors.WaitCursor
                Dim m_Report As ReportDocument = New ReportDocument()
                Dim reportPath As String = String.Empty

                Dim strPeriod As String = String.Empty

                Select Case CInt(Me.cboPeriodType.SelectedValue)
                    Case CType(enumPeriodType.OriginalBudget, Integer)
                        reportPath = p_strAppPath & "\Reports\RPT006-1.rpt"
                        strPeriod = "Original"
                        Exit Select
                    Case CType(enumPeriodType.EstimateBudget, Integer)
                        reportPath = p_strAppPath & "\Reports\RPT006-2.rpt"
                        strPeriod = "Estimate"
                        Exit Select
                    Case CType(enumPeriodType.ReviseBudget, Integer)
                        If Me.chkShowMTP.Checked = True Then
                            reportPath = p_strAppPath & "\Reports\RPT006-4.rpt"
                        Else
                            reportPath = p_strAppPath & "\Reports\RPT006-3.rpt"
                        End If
                        strPeriod = "Revise"
                        Exit Select
                    Case Else
                        reportPath = p_strAppPath & "\Reports\RPT006-1.rpt"
                        strPeriod = "Original"
                        Exit Select
                End Select

                m_Report.Load(reportPath)

                myClsBG0460BL.BudgetYear = Me.numYear.Value.ToString
                myClsBG0460BL.PeriodType = (Me.cboPeriodType.SelectedValue).ToString
                myClsBG0460BL.MTPChecked = Me.chkShowMTP.Checked
                myClsBG0460BL.ProjectNo = Me.numProjectNo.Value.ToString
                myClsBG0460BL.UserLevelId = p_intUserLevelId
                If Me.cboRevNo.DataSource IsNot Nothing Then
                    myClsBG0460BL.RevNo = Me.cboRevNo.SelectedValue.ToString
                End If
                myClsBG0460BL.PIC = Me.cboUserPIC.SelectedValue.ToString

                Dim ds As DataSet
                If myClsBG0460BL.GetBudgetData() = False Then
                    ds = Nothing
                Else
                    ds = myClsBG0460BL.BudgetData
                    If ds Is Nothing Or ds.Tables(0).Rows.Count = 0 Then
                        MessageBox.Show("No budget data found, please try it again.", "RPT006", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Me.Cursor = Cursors.Default
                        Return
                    End If
                End If
                m_Report.SetDataSource(ds)

                myClsBG0460BL.GetBudgetStatus()
                If myClsBG0460BL.BudgetStatus >= 5 Then
                    m_Report.ReportDefinition.ReportObjects("picAuth10").ObjectFormat.EnableSuppress = False
                Else
                    m_Report.ReportDefinition.ReportObjects("picAuth10").ObjectFormat.EnableSuppress = True
                End If

                If myClsBG0460BL.BudgetStatus >= 6 Then
                    m_Report.ReportDefinition.ReportObjects("picAuth3").ObjectFormat.EnableSuppress = False
                Else
                    m_Report.ReportDefinition.ReportObjects("picAuth3").ObjectFormat.EnableSuppress = True
                End If

                m_Report.SetParameterValue("PERIOD", strPeriod)
                m_Report.SetParameterValue("BUDGET_YEAR", Me.numYear.Value.ToString)
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
            MessageBox.Show(ex.Message, "RPT006", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Cursor = Cursors.Default
            Return
        End Try

    End Sub

    Private Sub cboPeriodType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPeriodType.SelectedIndexChanged

        If CInt(cboPeriodType.SelectedValue) = CType(enumPeriodType.ReviseBudget, Integer) Then
            Me.chkShowMTP.Enabled = True
        Else
            Me.chkShowMTP.Checked = False
            Me.chkShowMTP.Enabled = False
        End If

        LoadRevNo()

        LoadPersonInCharge()

    End Sub

    Private Sub cmdExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExcel.Click

        If fncCheckRevNo() = False Then

            MessageBox.Show("No budget data found, please try it again.", "RPT006", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor

        '//Get Export Data
        Dim dsData As DataSet
        myClsBG0460BL.BudgetYear = Me.numYear.Value.ToString
        myClsBG0460BL.PeriodType = (Me.cboPeriodType.SelectedValue).ToString
        myClsBG0460BL.MTPChecked = Me.chkShowMTP.Checked
        myClsBG0460BL.ProjectNo = Me.numProjectNo.Value.ToString
        myClsBG0460BL.UserLevelId = p_intUserLevelId
        If Me.cboRevNo.DataSource IsNot Nothing Then
            myClsBG0460BL.RevNo = Me.cboRevNo.SelectedValue.ToString
        End If
        myClsBG0460BL.PIC = Me.cboUserPIC.SelectedValue.ToString

        Dim strPeriod As String = cboPeriodType.Text
        strPeriod = strPeriod.Substring(0, strPeriod.IndexOf("Budget") - 1)

        If myClsBG0460BL.GetBudgetData() = False Then
            dsData = Nothing
            MessageBox.Show("Load buget data failed.", "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Cursor = Cursors.Default
            Return
        Else
            dsData = myClsBG0460BL.BudgetData
            If dsData Is Nothing Or dsData.Tables(0).Rows.Count = 0 Then
                MessageBox.Show("No budget data found, please try it again.", "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Cursor = Cursors.Default
                Return
            End If
        End If

        Dim dtAuthorizeImages As DataTable = myClsBG0460BL.BudgetData.Tables(1)

        '//Create output columns
        Dim strYear As String = Me.numYear.Value.ToString
        Dim dtColumns As DataTable = CreateTableTemplate()

        Dim strPeriodType As String = cboPeriodType.Text
        Dim strProject As String = Me.numProjectNo.Value.ToString

        Dim strSubTitle As String = String.Empty

        If strProject <> "1" Then
            strSubTitle = "Summary by Investment : " + strPeriodType + " " + strYear + " (Project No." + strProject + ")"
        Else
            strSubTitle = "Summary by Investment : " + strPeriodType + " " + strYear
        End If


        Select Case CInt(Me.cboPeriodType.SelectedValue)

            Case CType(enumPeriodType.OriginalBudget, Integer)  '//Original

                InsertOriginalColumnData(dtColumns, strYear)

                '//Create group data
                Dim intGroupLineIndex As Integer = 0
                Dim arrGroupNames As String() = New String() {"ASSET_PROJECT", "ASSET_CATEGORY", "PERSON_IN_CHARGE_NO"}
                Dim arrGroupColumnTitles As String() = New String() {"ASSET_PROJECT_NAME", "ASSET_CATEGORY_NAME", "PERSON_IN_CHARGE_NAME"}
                Dim arrSecondGroups As Integer() = New Integer() {0}
                Dim arrFirstGroups As Integer() = New Integer() {0}
                Dim dsGroups As DataSet = SetupInvestmentSummaryGroupbyData(dsData, arrGroupNames, _
                                                                arrGroupColumnTitles, 10, arrSecondGroups, arrFirstGroups)

                '//Create Output Excel
                OutputExcel(dsGroups, dtColumns, False, strSubTitle, strYear, True, arrFirstGroups, arrSecondGroups, strPeriod)

            Case CType(enumPeriodType.EstimateBudget, Integer)  '//Estimate

                InsertEstimateColumnData(dtColumns, strYear)

                '//Create group data
                Dim intGroupLineIndex As Integer = 0
                Dim arrGroupNames As String() = New String() {"ASSET_PROJECT", "ASSET_CATEGORY", "PERSON_IN_CHARGE_NO"}
                Dim arrGroupColumnTitles As String() = New String() {"ASSET_PROJECT_NAME", "ASSET_CATEGORY_NAME", "PERSON_IN_CHARGE_NAME"}
                Dim arrSecondGroups As Integer() = New Integer() {0}
                Dim arrFirstGroups As Integer() = New Integer() {0}
                Dim dsGroups As DataSet = SetupInvestmentSummaryGroupbyData(dsData, arrGroupNames, _
                                                arrGroupColumnTitles, 10, arrSecondGroups, arrFirstGroups)

                '//Create Output Excel
                OutputExcel(dsGroups, dtColumns, False, strSubTitle, strYear, True, arrFirstGroups, arrSecondGroups, strPeriod)

            Case CType(enumPeriodType.ReviseBudget, Integer)  '//Revise

                If Me.chkShowMTP.Checked = True Then
                    InsertMTPColumnData(dtColumns, strYear)
                Else
                    InsertReviseColumnData(dtColumns, strYear)
                End If

                '//Create group data
                Dim intGroupLineIndex As Integer = 0
                Dim arrGroupNames As String() = New String() {"ASSET_PROJECT", "ASSET_CATEGORY", "PERSON_IN_CHARGE_NO"}
                Dim arrGroupColumnTitles As String() = New String() {"ASSET_PROJECT_NAME", "ASSET_CATEGORY_NAME", "PERSON_IN_CHARGE_NAME"}
                Dim arrSecondGroups As Integer() = New Integer() {0}
                Dim arrFirstGroups As Integer() = New Integer() {0}
                Dim dsGroups As DataSet = SetupInvestmentSummaryGroupbyData(dsData, arrGroupNames, _
                                                                    arrGroupColumnTitles, 10, arrSecondGroups, arrFirstGroups)

                Dim bMTPCheck As Boolean = chkShowMTP.Checked

                '//Create Output Excel
                OutputExcel(dsGroups, dtColumns, bMTPCheck, strSubTitle, strYear, True, arrFirstGroups, arrSecondGroups, strPeriod)

        End Select

        Me.Cursor = Cursors.Default

    End Sub

    Private Function OutputExcel(ByVal dsData As DataSet, ByVal dtColumns As DataTable, ByVal bMTPCheck As Boolean, _
                                 ByVal strSubTitle As String, ByVal strYear As String, ByVal bShowGroupName As Boolean, _
                                 ByVal arrFirstGroups As Integer(), ByVal arrSecondGroups As Integer(), _
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
            CType(xBk.Worksheets(1), Excel.Worksheet).Delete()
            CType(xBk.Worksheets(2), Excel.Worksheet).Delete()
        End If

        '//Set Style Value < 0 please fill color "Red"
        Dim style As Excel.Style = excelApp.ActiveWorkbook.Styles.Add("NewStyle")
        style.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)

        For intSheetCount As Integer = 0 To dsData.Tables.Count - 1

            If intSheetCount <> 0 Then
                xBk.Sheets.Add()
            End If

            xSt = CType(xBk.ActiveSheet, Excel.Worksheet)
            xSt.Name = dsData.Tables(intSheetCount).TableName

            '//Setup DataColumn
            For i As Integer = 0 To dtColumns.Rows.Count - 1
                xSt.Cells(colStartIndex, i + 1) = dtColumns.Rows(i)("Column_Title").ToString
                xSt.Range(xSt.Cells(colStartIndex, i + 1), xSt.Cells(colStartIndex, i + 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            Next

            Dim arrCols() As Integer
            If strPeriod = "Original" Then

                arrCols = New Integer() {1, 2, 3, 4, 5, 6, 13, 20, 21, 22, 23}
                SetupOriginalColumnsCells(xSt, colStartIndex, 2, 3, "", arrCols, 7, 12, strYear, False, True, 14, 19)

            ElseIf strPeriod = "Estimate" Then

                arrCols = New Integer() {1, 2, 3, 4, 5, 6, 13, 14, 15}
                SetupEstimateColumnsCells(xSt, colStartIndex, 2, 3, "", arrCols, 7, 9, 10, 12, False)

            Else

                arrCols = New Integer() {1, 2, 3, 4, 5, 12, 13, 14, 21, 22, 23, 24}
                If bMTPCheck = True Then
                    arrCols = New Integer() {1, 2, 3, 4, 5, 12, 13, 14, 15}
                End If
                SetupReviseColumnsCells(xSt, colStartIndex, bMTPCheck, 2, 3, "", arrCols, 6, 8, 9, 11, 15, 20, 6, 11, 25, 29, False)
            End If

            '//Setup Data
            For rowIndex As Integer = 0 To dsData.Tables(intSheetCount).Rows.Count - 1

                Dim row As DataRow = dsData.Tables(intSheetCount).Rows(rowIndex)

                '//If the column is "Group_Header" Empty.
                If IsGroupHeaderEmpty(row) Then
                    Continue For
                End If

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

                        'If row(col.ColumnName).ToString <> String.Empty Then
                        '    If CDbl(row(col.ColumnName).ToString) = 0 Then
                        '        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = "-"
                        '        xSt.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                        '    Else
                        '        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName)
                        '        xSt.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).NumberFormat = "#,##0.00"
                        '    End If
                        'End If

                        ''//Add by Max 01/10/2012
                        ''//Set Style Value < 0 please fill color "Red"
                        'If CDec(row(col.ColumnName)) < 0 Then
                        '    xSt.Range(xSt.Cells(rowIndex + rowStartIndex, colIndex + 1), xSt.Cells(rowIndex + rowStartIndex, colIndex + 1)).Style = style
                        'End If
                        ''//End Add by Max 01/10/2012

                        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName)
                        xSt.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).NumberFormat = "#,##0.00"

                    Else
                        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName).ToString()
                    End If

                Next
            Next

            '//Add Title for excel 
            Dim strGroupName As String = "Capital Investment"
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

            '//Setup group header column to be left align
            xSt.Range(xSt.Cells(rowStartIndex, 1), xSt.Cells(rowMax, 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup font bold for group total
            xSt.Range(xSt.Cells(rowStartIndex, 1), xSt.Cells(rowStartIndex, colMax)).Font.Bold = True
            xSt.Range(xSt.Cells(rowStartIndex + 1, 1), xSt.Cells(rowStartIndex + 1, colMax)).Font.Bold = True
            For i As Integer = 0 To arrSecondGroups.Length - 1
                Dim intIndexTmp As Integer = arrSecondGroups(i) + 2
                xSt.Range(xSt.Cells(rowStartIndex + intIndexTmp, 1), xSt.Cells(rowStartIndex + intIndexTmp, colMax)).Font.Bold = True
            Next

            For i As Integer = 0 To arrFirstGroups.Length - 1
                Dim intIndexTmp As Integer = arrFirstGroups(i) + 1
                xSt.Range(xSt.Cells(rowStartIndex + intIndexTmp, 1), xSt.Cells(rowStartIndex + intIndexTmp, colMax)).Font.Bold = True
            Next

            '//Setup sheet properly width
            xSt.Range(xSt.Cells(2, 1), xSt.Cells(rowMax, colMax)).Columns.EntireColumn.AutoFit()

            '//Setup Columns wrap text
            xSt.Range(xSt.Cells(2, 2), xSt.Cells(rowMax, 2)).Columns.ColumnWidth = 14
            xSt.Range(xSt.Cells(2, 2), xSt.Cells(rowMax, 2)).WrapText = True

            xSt.Range(xSt.Cells(2, 3), xSt.Cells(rowMax, 4)).Columns.ColumnWidth = 10
            xSt.Range(xSt.Cells(2, 3), xSt.Cells(rowMax, 4)).WrapText = True

            If strPeriod = "Original" Then

                xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 6)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 6)).WrapText = True

                xSt.Range(xSt.Cells(2, 7), xSt.Cells(rowMax, 12)).Columns.ColumnWidth = 12

                xSt.Range(xSt.Cells(2, 13), xSt.Cells(rowMax, 13)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 13), xSt.Cells(rowMax, 13)).WrapText = True

                xSt.Range(xSt.Cells(2, 14), xSt.Cells(rowMax, 19)).Columns.ColumnWidth = 12

                xSt.Range(xSt.Cells(2, 20), xSt.Cells(rowMax, 22)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 20), xSt.Cells(rowMax, 22)).WrapText = True

                xSt.Range(xSt.Cells(2, 23), xSt.Cells(rowMax, 23)).Columns.ColumnWidth = 15

            ElseIf strPeriod = "Estimate" Then

                xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 6)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 6)).WrapText = True

                xSt.Range(xSt.Cells(2, 7), xSt.Cells(rowMax, 12)).Columns.ColumnWidth = 12

                xSt.Range(xSt.Cells(2, 13), xSt.Cells(rowMax, 15)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 13), xSt.Cells(rowMax, 15)).WrapText = True

            Else
                xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 5)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 5)).WrapText = True

                xSt.Range(xSt.Cells(2, 6), xSt.Cells(rowMax, 11)).Columns.ColumnWidth = 12

                xSt.Range(xSt.Cells(2, 12), xSt.Cells(rowMax, 14)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 12), xSt.Cells(rowMax, 14)).WrapText = True

                xSt.Range(xSt.Cells(2, 15), xSt.Cells(rowMax, 20)).Columns.ColumnWidth = 12

                xSt.Range(xSt.Cells(2, 21), xSt.Cells(rowMax, 24)).Columns.ColumnWidth = 13
                xSt.Range(xSt.Cells(2, 21), xSt.Cells(rowMax, 24)).WrapText = True

                If chkShowMTP.Checked = True Then
                    xSt.Range(xSt.Cells(2, 25), xSt.Cells(rowMax, 29)).Columns.ColumnWidth = 12
                End If

            End If

            colStartIndex = colStartIndex - 1
            '//Setup Column Font 
            xSt.Range(excelApp.Cells(colStartIndex, 1), excelApp.Cells(colStartIndex + 1, colMax)).Font.Bold = True
            xSt.Range(xSt.Cells(colStartIndex, 1), xSt.Cells(rowMax, colMax)).Font.Name = "Tahoma"
            xSt.Range(xSt.Cells(colStartIndex, 1), xSt.Cells(rowMax, colMax)).Font.Size = 10

            '//Setup border
            xSt.Range(excelApp.Cells(colStartIndex, 1), excelApp.Cells(rowMax, colMax)).Borders.LineStyle = 1
            xSt.Range(excelApp.Cells(colStartIndex, 1), excelApp.Cells(rowMax, 1)).Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlThin
            xSt.Range(excelApp.Cells(colStartIndex, 1), excelApp.Cells(colStartIndex, colMax)).Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThin
            xSt.Range(excelApp.Cells(colStartIndex, colMax), excelApp.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThin
            xSt.Range(excelApp.Cells(rowMax, 1), excelApp.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThin

            '//Insert empty column
            'If bMTPCheck = True Then
            '    SetupMTPEmptyColumn(xSt, colStartIndex, rowMax, colMax, 25, rowMax, 1, False)
            'End If

            '//-- Add by Max 26/09/2012

            '//Set NumberFormat = "#,##0.00;[Red]-#,##0.00"
            xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax, colMax)).NumberFormat = "#,##0.00;[Red]-#,##0.00"

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
                xSt.Range(xSt.Cells(colStartIndex, 13), xSt.Cells(rowMax, 13)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 14), xSt.Cells(rowMax, 19)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 20), xSt.Cells(rowMax, 23)).Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium

            ElseIf strPeriod = "Estimate" Then

                xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax, 6)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 7), xSt.Cells(rowMax, 9)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 10), xSt.Cells(rowMax, 12)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
                xSt.Range(xSt.Cells(colStartIndex, 13), xSt.Cells(rowMax, 15)).Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium

            ElseIf strPeriod = "Revise" Then

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

                End If
                '//-- Edit Add by Max 26/09/2012

        Next

        '//Show Excel
        excelApp.Visible = True

        '//-- Begin Add by S.Watcharapong 2011-05-24
        '//Release memory
        BGCommon.ExcelReleasememory(excelApp, xBk, xSt)
        '//-- End Add 2011-05-24

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
        dRow("Column_Title") = "Item"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "BUDGET_ORDER_NO"
        dRow("Column_Title") = "New Budget Code"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "COST_CENTER"
        dRow("Column_Title") = "Cost Center"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERSON_IN_CHARGE_NO"
        dRow("Column_Title") = "Person in Charge"
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

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "LAST_YEAR_SUM"
        dRow("Column_Title") = "Total Year'" & strLastYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_YEAR"
        dRow("Column_Title") = "Difference"
        dtColumns.Rows.Add(dRow)

        Return True

    End Function

    Private Function InsertEstimateColumnData(ByRef dtColumns As DataTable, ByVal strYear As String) As Boolean

        Dim dRow As DataRow
        Dim strHalfYear As String = strYear.Substring(2, 2)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "Group_Header"
        dRow("Column_Title") = "Item"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "BUDGET_ORDER_NO"
        dRow("Column_Title") = "New Budget Code"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "COST_CENTER"
        dRow("Column_Title") = "Cost Center"
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
        dRow("Column_Name") = "DIFF_SECOND_HALF"
        dRow("Column_Title") = "Difference 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "ESTIMATE_TOTAL_YEAR"
        dRow("Column_Title") = "Estimate Year'" & strYear
        dtColumns.Rows.Add(dRow)

        Return True

    End Function

    Private Function InsertReviseColumnData(ByRef dtColumns As DataTable, ByVal strYear As String) As Boolean

        Dim dRow As DataRow
        Dim strHalfYear As String = strYear.Substring(2, 2)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "Group_Header"
        dRow("Column_Title") = "Item"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "BUDGET_ORDER_NO"
        dRow("Column_Title") = "New Budget Code"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "COST_CENTER"
        dRow("Column_Title") = "Cost Center"
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
        dRow("Column_Name") = "REVISE_SECOND_HALF"
        dRow("Column_Title") = "Revise 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_SECOND_HALF"
        dRow("Column_Title") = "Diff 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "REVISE_TOTAL_YEAR"
        dRow("Column_Title") = "Revise Year'" & strYear
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
        dRow("Column_Title") = "Item"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "BUDGET_ORDER_NO"
        dRow("Column_Title") = "New Budget Code"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "COST_CENTER"
        dRow("Column_Title") = "Cost Center"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERSON_IN_CHARGE_NO"
        dRow("Column_Title") = "Person in Charge"
        dtColumns.Rows.Add(dRow)

        'dRow = dtColumns.NewRow
        'dRow("Column_Name") = "ORIGINAL_FIRST_HALF"
        'dRow("Column_Title") = "Original 1st Half'" & strHalfYear
        'dtColumns.Rows.Add(dRow)

        'dRow = dtColumns.NewRow
        'dRow("Column_Name") = "M1"
        'dRow("Column_Title") = "Jan'" & strHalfYear
        'dtColumns.Rows.Add(dRow)

        'dRow = dtColumns.NewRow
        'dRow("Column_Name") = "M2"
        'dRow("Column_Title") = "Feb'" & strHalfYear
        'dtColumns.Rows.Add(dRow)

        'dRow = dtColumns.NewRow
        'dRow("Column_Name") = "M3"
        'dRow("Column_Title") = "Mar'" & strHalfYear
        'dtColumns.Rows.Add(dRow)

        'dRow = dtColumns.NewRow
        'dRow("Column_Name") = "M4"
        'dRow("Column_Title") = "Apr'" & strHalfYear
        'dtColumns.Rows.Add(dRow)

        'dRow = dtColumns.NewRow
        'dRow("Column_Name") = "M5"
        'dRow("Column_Title") = "May'" & strHalfYear
        'dtColumns.Rows.Add(dRow)

        'dRow = dtColumns.NewRow
        'dRow("Column_Name") = "M6"
        'dRow("Column_Title") = "Jun'" & strHalfYear
        'dtColumns.Rows.Add(dRow)

        'dRow = dtColumns.NewRow
        'dRow("Column_Name") = "ESTIMATE_FIRST_HALF"
        'dRow("Column_Title") = "Estimate 1st Half'" & strHalfYear
        'dtColumns.Rows.Add(dRow)

        'dRow = dtColumns.NewRow
        'dRow("Column_Name") = "DIFF_FIRST_HALF"
        'dRow("Column_Title") = "Diff 1st Half'" & strHalfYear
        'dtColumns.Rows.Add(dRow)

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
        dRow("Column_Title") = "Revise 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_SECOND_HALF"
        dRow("Column_Title") = "Diff 2nd Half'" & strHalfYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "REVISE_TOTAL_YEAR"
        dRow("Column_Title") = "Revise Year'" & strYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_TOTAL_YEAR"
        dRow("Column_Title") = "Diff Year'" & strYear
        dtColumns.Rows.Add(dRow)

        'dRow = dtColumns.NewRow
        'dRow("Column_Name") = "RRT1"
        'dRow("Column_Title") = "Y" & CStr(intYear + 1)
        'dtColumns.Rows.Add(dRow)

        'dRow = dtColumns.NewRow
        'dRow("Column_Name") = "RRT2"
        'dRow("Column_Title") = "Y" & CStr(intYear + 2)
        'dtColumns.Rows.Add(dRow)

        'dRow = dtColumns.NewRow
        'dRow("Column_Name") = "RRT3"
        'dRow("Column_Title") = "Y" & CStr(intYear + 3)
        'dtColumns.Rows.Add(dRow)

        'dRow = dtColumns.NewRow
        'dRow("Column_Name") = "RRT4"
        'dRow("Column_Title") = "Y" & CStr(intYear + 4)
        'dtColumns.Rows.Add(dRow)

        'dRow = dtColumns.NewRow
        'dRow("Column_Name") = "RRT5"
        'dRow("Column_Title") = "Y" & CStr(intYear + 5)
        'dtColumns.Rows.Add(dRow)

        Return True

    End Function

    Private Function SetupInvestmentSummaryGroupbyData(ByVal dsData As DataSet, ByVal arrGroupColumnName As String(), ByVal arrGroupColumnTitles As String(), _
                                                       ByVal intDataColumnIndex As Integer, ByRef arrSecondGroups() As Integer, _
                                                       ByRef arrFirstGroups() As Integer) As DataSet

        Dim dsResult As DataSet = New DataSet
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object
        Dim drEmpty As DataRow
        Dim strScript As String
        Dim intGroupTotalIndex As Integer
        Dim intSecondGroupTotalIndex As Integer
        Dim dtResult As DataTable = dsData.Tables(0).Clone
        Dim dtTmp As DataTable = dsData.Tables(0).Clone
        Dim dtSecondGroup As DataTable = dsData.Tables(0).Clone

        Dim drTotal As DataRow

        '//Add one more column for group name
        Dim col As DataColumn = New DataColumn()
        col.ColumnName = "Group_Header"
        col.DataType = Type.GetType("System.String")
        dtResult.Columns.Add(col)

        col = New DataColumn()
        col.ColumnName = "Group_Header"
        col.DataType = Type.GetType("System.String")
        dtTmp.Columns.Add(col)

        col = New DataColumn()
        col.ColumnName = "Group_Header"
        col.DataType = Type.GetType("System.String")
        dtSecondGroup.Columns.Add(col)

        Dim strFirstGroupColumnName As String = arrGroupColumnName(0)
        Dim strFirstGroupColumnTitle As String = arrGroupColumnTitles(0)

        strScript = strFirstGroupColumnName
        Dim dtGroups As DataTable = dsData.Tables(0).DefaultView.ToTable(True, strScript)
        Dim intGroupCount As Integer = dtGroups.Rows.Count

        '//Caculate Grand Total line
        Dim drAllTotal As DataRow = dtResult.NewRow
        For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName
            strExpression = "Sum(" + strColumnName + ")"
            returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
            drAllTotal(dsData.Tables(0).Columns(k).ColumnName) = returnValue
        Next

        intGroupTotalIndex = 0
        For i As Integer = 0 To intGroupCount - 1

            '//Seperate dataset data into several datatables according to group no
            If dtGroups.Rows(i)(0).ToString = String.Empty Then
                Continue For
            End If
            strScript = strFirstGroupColumnName + " = " + dtGroups.Rows(i)(0).ToString
            Dim arrRows As DataRow() = dsData.Tables(0).Select(strScript)

            dtSecondGroup.Rows.Clear()
            For j As Integer = 0 To arrRows.Length - 1

                Dim drow(dtResult.Columns.Count - 1) As Object
                Dim drow2(dtResult.Columns.Count - 1) As Object

                arrRows(j).ItemArray.CopyTo(drow, 0)
                arrRows(j).ItemArray.CopyTo(drow2, 0)

                dtResult.Rows.Add(drow)
                dtResult.Rows(dtResult.Rows.Count - 1)("Group_Header") = dtResult.Rows(dtResult.Rows.Count - 1)("BUDGET_ORDER_NAME")

                dtSecondGroup.Rows.Add(drow2)
            Next

            '//Calculate total for each group
            drTotal = dtResult.NewRow
            For k As Integer = intDataColumnIndex To dtSecondGroup.Columns.Count - 2

                Dim strColumnName As String = dtSecondGroup.Columns(k).ColumnName
                strExpression = "Sum(" + strColumnName + ")"
                returnValue = dtSecondGroup.Compute(strExpression, strFilter)
                drTotal(dtSecondGroup.Columns(k).ColumnName) = returnValue

            Next

            'For intIndex As Integer = 0 To intDataColumnIndex - 1
            '    drTotal(intIndex) = dtResult.Rows(intGroupTotalIndex)(intIndex)
            'Next
            drTotal("Group_Header") = dtSecondGroup.Rows(0)(strFirstGroupColumnTitle).ToString + " Capital Investment"

            'If bShowGroupName = True Then
            '    dtResult.TableName = arrRows(0)(strGroupColumnName).ToString & " " & arrRows(0)(strGroupColumnTitle).ToString
            'End If

            '//
            '//Calculate second group total line
            '//
            Dim strSecondGroupColumnName As String = arrGroupColumnName(1)
            Dim strSecondGroupColumnTitle As String = arrGroupColumnTitles(1)

            strScript = strSecondGroupColumnName
            Dim dtSecondGroups As DataTable = dtSecondGroup.DefaultView.ToTable(True, strScript)
            Dim intSecondGroupCount As Integer = dtSecondGroups.Rows.Count

            'ReDim arrSecondGroups(intSecondGroupCount - 1)

            intSecondGroupTotalIndex = intGroupTotalIndex
            Dim drSecondTotal As DataRow
            For j As Integer = 0 To intSecondGroupCount - 1

                If dtSecondGroups.Rows(j)(0).ToString = String.Empty Then
                    Continue For
                End If

                strScript = strSecondGroupColumnName + " = " + dtSecondGroups.Rows(j)(0).ToString
                Dim arrRows2 As DataRow() = dtSecondGroup.Select(strScript)
                dtTmp.Rows.Clear()

                For k As Integer = 0 To arrRows2.Length - 1
                    Dim drow(dtTmp.Columns.Count - 1) As Object
                    arrRows2(k).ItemArray.CopyTo(drow, 0)
                    dtTmp.Rows.Add(drow)
                Next

                '//Calculate total for each group
                drSecondTotal = dtResult.NewRow
                For k As Integer = intDataColumnIndex To dtTmp.Columns.Count - 2

                    Dim strColumnName As String = dtResult.Columns(k).ColumnName
                    strExpression = "Sum(" + strColumnName + ")"
                    returnValue = dtTmp.Compute(strExpression, strFilter)
                    drSecondTotal(dtResult.Columns(k).ColumnName) = returnValue

                Next

                '//
                '//Calculate third group records count
                '//
                Dim strThirdGroupColumnName As String = arrGroupColumnName(2)
                'Dim strThirdGroupColumnTitle As String = arrGroupColumnTitles(2)
                Dim intThirdGroupEmptyIndex As Integer = intSecondGroupTotalIndex

                strScript = strThirdGroupColumnName
                Dim dtThirdGroups As DataTable = dtTmp.DefaultView.ToTable(True, strScript)
                Dim intThirdGroupCount As Integer = dtThirdGroups.Rows.Count
                Dim intGroup3EmptyRows As Integer = 0

                If intThirdGroupCount > 1 Then

                    For intThirdIndex As Integer = 0 To intThirdGroupCount - 1
                        strScript = strThirdGroupColumnName + " = '" + dtThirdGroups.Rows(intThirdIndex)(0).ToString + "'"
                        Dim arrRows3 As DataRow() = dtTmp.Select(strScript)
                        If intThirdIndex <> intThirdGroupCount - 1 Then
                            drEmpty = dtResult.NewRow
                            dtResult.Rows.InsertAt(drEmpty, intThirdGroupEmptyIndex + arrRows3.Length + intGroup3EmptyRows)
                            intGroup3EmptyRows = intGroup3EmptyRows + 1
                            intThirdGroupEmptyIndex = intThirdGroupEmptyIndex + arrRows3.Length
                        End If
                    Next

                End If

                Dim intGroupTotalCount As Integer = arrSecondGroups.Length
                ReDim Preserve arrSecondGroups(intGroupTotalCount)
                arrSecondGroups(intGroupTotalCount - 1) = intSecondGroupTotalIndex

                '//Add total cost for group 2
                dtResult.Rows.InsertAt(drSecondTotal, intSecondGroupTotalIndex)
                dtResult.Rows(intSecondGroupTotalIndex)("Group_Header") = "(" + CStr(j + 1) + ") " + dtResult.Rows(intSecondGroupTotalIndex + 1)(strSecondGroupColumnTitle).ToString

                '//Add one empty row
                drEmpty = dtResult.NewRow
                dtResult.Rows.InsertAt(drEmpty, intSecondGroupTotalIndex + dtTmp.Rows.Count + 1 + intGroup3EmptyRows)

                intSecondGroupTotalIndex = intSecondGroupTotalIndex + dtTmp.Rows.Count + 2 + intGroup3EmptyRows

            Next

            Dim intTmp As Integer = arrFirstGroups.Length
            ReDim Preserve arrFirstGroups(intTmp)
            arrFirstGroups(intTmp - 1) = intGroupTotalIndex

            '//Add total cost for group 1
            dtResult.Rows.InsertAt(drTotal, intGroupTotalIndex)
            intGroupTotalIndex = intSecondGroupTotalIndex + 1
        Next

        '//Add All total cost
        dtResult.Rows.InsertAt(drAllTotal, 0)
        dtResult.Rows(0)("Group_Header") = "BTMT Capital Investment"

        '//Return data table
        dsResult.Tables.Add(dtResult)
        Return dsResult

    End Function

    Private Function SetupTitleIndex(ByVal strPeriod As String, ByRef intUnitPriceStart As Integer, _
                                     ByRef intUnitPriceEnd As Integer, ByRef intAuthorizeStart As Integer, _
                                     ByRef intAuthorizeEnd As Integer, ByRef intImageIndex As Integer, ByVal bMTPCheck As Boolean) As Boolean

        Select Case strPeriod

            Case "Original"

                intUnitPriceStart = 22
                intUnitPriceEnd = 23

                intAuthorizeStart = 20
                intAuthorizeEnd = 21

                intImageIndex = 1235

            Case "Estimate"

                intUnitPriceStart = 14
                intUnitPriceEnd = 15

                intAuthorizeStart = 11
                intAuthorizeEnd = 13

                intImageIndex = 795

            Case "Revise"

                'If bMTPCheck = True Then
                '    intUnitPriceStart = 28
                '    intUnitPriceEnd = 29
                'Else
                '    intUnitPriceStart = 23
                '    intUnitPriceEnd = 24
                'End If

                'If bMTPCheck = True Then
                '    intAuthorizeStart = 14
                '    intAuthorizeEnd = 15
                'Else
                '    intAuthorizeStart = 21
                '    intAuthorizeEnd = 22
                'End If

                'If bMTPCheck = True Then
                '    intImageIndex = 885
                'Else
                '    intImageIndex = 1220
                'End If

                intUnitPriceStart = 23
                intUnitPriceEnd = 24

                intAuthorizeStart = 21
                intAuthorizeEnd = 22

                intImageIndex = 1220

            Case Else
                Exit Select
        End Select

        Return True

    End Function

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub numYear_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numYear.ValueChanged
        LoadRevNo()

        LoadPersonInCharge()

    End Sub

    Private Sub numProjectNo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numProjectNo.ValueChanged
        LoadRevNo()
    End Sub
#End Region

End Class