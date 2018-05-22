Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports Microsoft.Office.Interop
Imports System.Drawing.Printing

Public Class frmBG0480

#Region "Variable"
    Private myClsBG0480BL As New clsBG0480BL
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
            'cboPeriodType.Items.Add("Revise Budget")
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

    'Private Sub LoadPrevRevNo()


    '    If Me.gbPrevYear.Enabled = True AndAlso Me.cboPrevRevno.Visible = True Then

    '        If Me.cboPeriodType.SelectedIndex < 0 OrElse _
    '            Me.numPrevProjectNo.Value <= 0 OrElse _
    '            Me.numYear.Value <= 0 Then

    '            Me.cboPrevRevno.DataSource = Nothing
    '            Exit Sub

    '        End If

    '        Dim strPeroidType As String = String.Empty
    '        If CType(cboPeriodType.SelectedValue, Integer) = CType(enumPeriodType.MTPBudget, Integer) Then
    '            strPeroidType = CStr(enumPeriodType.MTPBudget)
    '        ElseIf CType(cboPeriodType.SelectedValue, Integer) = CType(enumPeriodType.OriginalBudget, Integer) Then
    '            strPeroidType = CStr(enumPeriodType.MTPBudget)
    '        End If

    '        Dim strProjectNo = Me.numPrevProjectNo.Value.ToString
    '        If Not strProjectNo Is Nothing And strProjectNo <> String.Empty And strProjectNo <> "System.Data.DataRowView" Then

    '            myClsBG0310BL.BudgetYear = CStr(Me.numYear.Value - 1)
    '            myClsBG0310BL.PeriodType = strPeroidType
    '            myClsBG0310BL.ProjectNo = strProjectNo
    '            myClsBG0310BL.BudgetType = BGConstant.P_BUDGET_TYPE_EXPENSE

    '            If myClsBG0310BL.GetRevNo() = True Then
    '                Me.cboPrevRevno.DisplayMember = "REV_NO"
    '                Me.cboPrevRevno.ValueMember = "REV_NO"
    '                Me.cboPrevRevno.DataSource = myClsBG0310BL.RevNoList
    '            Else
    '                Me.cboPrevRevno.DataSource = Nothing
    '            End If
    '        Else
    '            Me.cboPrevRevno.DataSource = Nothing
    '        End If

    '    End If
    'End Sub

    'Private Function fncCheckPrevRevNo() As Boolean
    '    Dim blnChkResult As Boolean = True

    '    If CType(cboPeriodType.SelectedValue, Integer) = CType(enumPeriodType.MTPBudget, Integer) AndAlso _
    '        Me.gbPrevYear.Enabled = True AndAlso Me.cboPrevRevno.Visible = True Then

    '        If p_intUserLevelId = enumUserLevel.SystemAdministrator Then

    '            If Me.cboPrevRevno.DataSource Is Nothing OrElse _
    '                Me.cboPrevRevno.SelectedIndex < 0 Then
    '                blnChkResult = False
    '            End If

    '        End If

    '    End If

    '    Return blnChkResult
    'End Function

    'Private Sub EnablePrev()

    '    Me.gbPrevYear.Enabled = True
    '    Me.numPrevProjectNo.Enabled = True
    '    LoadPrevRevNo()

    'End Sub

    'Private Sub DisablePrev()

    '    Me.numPrevProjectNo.Value = 1
    '    Me.numPrevProjectNo.Enabled = False
    '    Me.cboPrevRevno.SelectedIndex = -1
    '    Me.gbPrevYear.Enabled = False

    'End Sub

#End Region

#Region "Control Event"

    Private Sub frmBG0480_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        If Not excelApp Is Nothing Then
            excelApp.Quit()
            excelApp = Nothing
        End If
    End Sub

    Private Sub frmBG0480_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not excelApp Is Nothing Then
            excelApp.Quit()
            excelApp = Nothing
        End If
    End Sub

    Private Sub frmBG0480_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
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
            myClsBG0480BL.BudgetYear = Me.numYear.Value.ToString
            myClsBG0480BL.PeriodType = (Me.cboPeriodType.SelectedValue).ToString
            myClsBG0480BL.PIC = Me.cboUserPIC.SelectedValue.ToString
            myClsBG0480BL.ProjectNo = Me.numProjectNo.Value.ToString
            If Me.cboRevNo.SelectedValue IsNot Nothing Then
                myClsBG0480BL.RevNo = Me.cboRevNo.SelectedValue.ToString
            Else
                myClsBG0480BL.RevNo = "1"
            End If
            myClsBG0480BL.BudgetType = "E"
            myClsBG0480BL.UserLevelId = p_intUserLevelId

            If Me.cboRevNo.DataSource IsNot Nothing Then
                myClsBG0480BL.RevNo = Me.cboRevNo.SelectedValue.ToString
            End If

            myClsBG0480BL.PrevProjectNo = Me.numPrevProjectNo.Value.ToString
            If Me.cboPrevRevno.DataSource IsNot Nothing AndAlso _
                Me.cboPrevRevno.SelectedValue IsNot Nothing Then
                myClsBG0480BL.PrevRevNo = Me.cboPrevRevno.SelectedValue.ToString
            Else
                myClsBG0480BL.PrevRevNo = String.Empty
            End If

            If myClsBG0480BL.GetCommentData() = False Then
                clsBG0400.DS = Nothing
                'Add
                MessageBox.Show("No budget data found, please try it again.", "RPT008", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Cursor = Cursors.Default
                Return
            Else
                clsBG0400.DS = myClsBG0480BL.BudgetData
                If clsBG0400.DS Is Nothing Or clsBG0400.DS.Tables(0).Rows.Count = 0 Then
                    MessageBox.Show("No budget data found, please try it again.", "RPT008", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Me.Cursor = Cursors.Default
                    Return
                End If
            End If


            myClsBG0480BL.GetBudgetStatus()

            Dim strPeriod As String = String.Empty
            clsBG0400.ReportName = "RPT008.rpt"
            Select Case CInt(Me.cboPeriodType.SelectedValue)
                Case CType(enumPeriodType.OriginalBudget, Integer)
                    'clsBG0400.ReportName = "RPT001-1.rpt"
                    strPeriod = "Original"
                    Exit Select
                Case CType(enumPeriodType.EstimateBudget, Integer)
                    'clsBG0400.ReportName = "RPT001-2.rpt"
                    strPeriod = "Estimate"
                    Exit Select
                Case CType(enumPeriodType.ReviseBudget, Integer)
                    'If Me.chkShowMTP.Checked = True Then
                    '    clsBG0400.ReportName = "RPT001-4.rpt"
                    'Else
                    '    clsBG0400.ReportName = "RPT001-3.rpt"
                    'End If
                    strPeriod = "Revise"
                    Exit Select
                Case CType(enumPeriodType.MTPBudget, Integer)
                    'clsBG0400.ReportName = "RPT001-5.rpt"
                    strPeriod = "MTP"
                    Exit Select
            End Select

            'clsBG0400.ConfigureCrystalReports()
            clsBG0400.PIC = Me.cboUserPIC.Text.ToString
            clsBG0400.BudgetYear = Me.numYear.Value.ToString
            'clsBG0400.ParamPersonInCharge = True
            clsBG0400.Period = strPeriod
            clsBG0400.ReportType = "CommentByPersonInCharge"
            clsBG0400.BudgetStatus = myClsBG0480BL.BudgetStatus
            clsBG0400.ProjectNo = Me.numProjectNo.Value.ToString

            clsBG0400.MdiParent = p_frmBG0010
            clsBG0400.Show()

            If clsBG0400.WindowState = FormWindowState.Minimized Then
                clsBG0400.WindowState = FormWindowState.Normal
            End If
            clsBG0400.BringToFront()
            Me.Cursor = Cursors.Default

        Catch ex As Exception
            MessageBox.Show(ex.Message, "RPT008", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.Cursor = Cursors.Default
            Return
        End Try

    End Sub

    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrint.Click

        Try

            If fncCheckRevNo() = False Then

                MessageBox.Show("No budget data found, please try it again.", "RPT008", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            'If fncCheckPrevRevNo() = False Then
            '    MessageBox.Show("No previous budget data found, please try it again.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            '    Exit Sub
            'End If

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
            '     printDialog1.ShowDialog()

            If PrintDialog1.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then

                Me.Cursor = Cursors.WaitCursor
                Dim m_Report As ReportDocument = New ReportDocument()
                Dim reportPath As String = String.Empty
                'Select Case CInt(Me.cboPeriodType.SelectedValue)
                '    Case CType(enumPeriodType.OriginalBudget, Integer)
                '        reportPath = p_strAppPath & "\Reports\RPT001-1.rpt"
                '        Exit Select
                '    Case CType(enumPeriodType.EstimateBudget, Integer)
                '        reportPath = p_strAppPath & "\Reports\RPT001-2.rpt"
                '        Exit Select
                '    Case CType(enumPeriodType.ReviseBudget, Integer)
                '        If Me.chkShowMTP.Checked = True Then
                '            reportPath = p_strAppPath & "\Reports\RPT001-4.rpt"
                '        Else
                '            reportPath = p_strAppPath & "\Reports\RPT001-3.rpt"
                '        End If
                '        Exit Select
                '    Case CType(enumPeriodType.MTPBudget, Integer)
                '        reportPath = p_strAppPath & "\Reports\RPT001-5.rpt"
                '        Exit Select
                '    Case Else
                '        reportPath = p_strAppPath & "\Reports\RPT001-1.rpt"
                '        Exit Select
                'End Select
                reportPath = p_strAppPath & "\Reports\RPT008.rpt"
                m_Report.Load(reportPath)

                myClsBG0480BL.BudgetYear = Me.numYear.Value.ToString
                myClsBG0480BL.PeriodType = (Me.cboPeriodType.SelectedValue).ToString
                myClsBG0480BL.PIC = Me.cboUserPIC.SelectedValue.ToString
                'myClsBG0410BL.MTPChecked = Me.chkShowMTP.Checked
                myClsBG0480BL.ProjectNo = Me.numProjectNo.Value.ToString
                myClsBG0480BL.UserLevelId = p_intUserLevelId
                If Me.cboRevNo.DataSource IsNot Nothing Then
                    myClsBG0480BL.RevNo = Me.cboRevNo.SelectedValue.ToString
                End If
                myClsBG0480BL.BudgetType = "E"

                myClsBG0480BL.PrevProjectNo = Me.numPrevProjectNo.Value.ToString
                If Me.cboPrevRevno.DataSource IsNot Nothing AndAlso _
                    Me.cboPrevRevno.SelectedValue IsNot Nothing Then
                    myClsBG0480BL.PrevRevNo = Me.cboPrevRevno.SelectedValue.ToString
                Else
                    myClsBG0480BL.PrevRevNo = String.Empty
                End If

                Dim ds As DataSet
                If myClsBG0480BL.GetCommentData() = False Then
                    ds = Nothing
                    'Add
                    MessageBox.Show("No budget data found, please try it again.", "RPT008", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Me.Cursor = Cursors.Default
                    Return
                Else
                    ds = myClsBG0480BL.BudgetData
                    If ds Is Nothing Or ds.Tables(0).Rows.Count = 0 Then
                        MessageBox.Show("No budget data found, please try it again.", "RPT008", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Me.Cursor = Cursors.Default
                        Return
                    End If
                End If
                m_Report.SetDataSource(ds)

                myClsBG0480BL.GetBudgetStatus()

                If myClsBG0480BL.BudgetStatus >= 5 Then
                    m_Report.ReportDefinition.ReportObjects("picAuth10").ObjectFormat.EnableSuppress = False
                Else
                    m_Report.ReportDefinition.ReportObjects("picAuth10").ObjectFormat.EnableSuppress = True
                End If

                If myClsBG0480BL.BudgetStatus >= 6 Then
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
                    Case CType(enumPeriodType.ReviseBudget, Integer)
                        strPeriod = "Revise"
                    Case CType(enumPeriodType.MTPBudget, Integer)
                        strPeriod = "MTP"
                End Select

                'm_Report.SetParameterValue("PERSON_IN_CHARGE_NM", Me.cboUserPIC.Text.ToString)
                m_Report.SetParameterValue("BUDGET_YEAR", Me.numYear.Value.ToString)
                'm_Report.SetParameterValue("HALF_BUDGET_YEAR", Me.numYear.Value.ToString.Substring(2, 2))
                m_Report.SetParameterValue("PERIOD", strPeriod)
                'm_Report.SetParameterValue("PROJECT_NO", Me.numProjectNo.Value.ToString)

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
            MessageBox.Show(ex.Message, "RPT008", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.Cursor = Cursors.Default
            Return
        End Try

    End Sub

    'Private Sub pd_QueryPageSettings(ByVal sender As Object, ByVal ev As QueryPageSettingsEventArgs)

    '    For i = 0 To ev.PageSettings.PrinterSettings.PaperSizes.Count - 1

    '        If ev.PageSettings.PrinterSettings.PaperSizes.Item(i).Kind = PaperKind.A3 Then

    '            ev.PageSettings.PaperSize = ev.PageSettings.PrinterSettings.PaperSizes.Item(i)

    '            Exit For

    '        End If

    '    Next

    '    ev.PageSettings.Landscape = True

    'End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub cmdExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExcel.Click

        If fncCheckRevNo() = False Then

            MessageBox.Show("No budget data found, please try it again.", "RPT008", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        'If fncCheckPrevRevNo() = False Then
        '    MessageBox.Show("No previous budget data found, please try it again.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
        '    Exit Sub
        'End If

        Me.Cursor = Cursors.WaitCursor

        '//Get Export Data
        Dim dsData As DataSet
        myClsBG0480BL.BudgetYear = Me.numYear.Value.ToString
        myClsBG0480BL.PeriodType = (Me.cboPeriodType.SelectedValue).ToString
        myClsBG0480BL.PIC = Me.cboUserPIC.SelectedValue.ToString
        'myClsBG0410BL.MTPChecked = Me.chkShowMTP.Checked
        myClsBG0480BL.ProjectNo = Me.numProjectNo.Value.ToString
        myClsBG0480BL.UserLevelId = p_intUserLevelId
        If Me.cboRevNo.DataSource IsNot Nothing Then
            myClsBG0480BL.RevNo = Me.cboRevNo.SelectedValue.ToString
        End If
        myClsBG0480BL.BudgetType = "E"

        myClsBG0480BL.PrevProjectNo = Me.numPrevProjectNo.Value.ToString
        If Me.cboPrevRevno.DataSource IsNot Nothing AndAlso _
            Me.cboPrevRevno.SelectedValue IsNot Nothing Then
            myClsBG0480BL.PrevRevNo = Me.cboPrevRevno.SelectedValue.ToString
        Else
            myClsBG0480BL.PrevRevNo = String.Empty
        End If

        Dim strPeriod As String = cboPeriodType.Text
        strPeriod = strPeriod.Substring(0, strPeriod.IndexOf("Budget") - 1)

        If myClsBG0480BL.GetCommentData() = False Then
            dsData = Nothing
            MessageBox.Show("No budget data found, please try it again.", "RPT008", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Cursor = Cursors.Default
            Return
        Else
            dsData = myClsBG0480BL.BudgetData
            If dsData Is Nothing Or dsData.Tables(0).Rows.Count = 0 Then
                MessageBox.Show("No budget data found, please try it again.", "RPT008", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Cursor = Cursors.Default
                Return
            End If
        End If

        Dim dtAuthorizeImages As DataTable = myClsBG0480BL.BudgetData.Tables(1)
        Dim strYear As String = Me.numYear.Value.ToString
        Dim dtColumns As DataTable = CreateTableTemplate()

        Dim strPeriodType As String = cboPeriodType.Text
        Dim strProjectNo As String = Me.numProjectNo.Value.ToString

        Dim strSubTitle As String = String.Empty

        'If strProjectNo <> "1" Then
        '    strSubTitle = "Detail by Person In Charge : " + strPeriodType + " " + strYear + " (Project No. " + strProjectNo + ")"
        'Else
        '    strSubTitle = "Detail by Person In Charge : " + strPeriodType + " " + strYear
        'End If

        strSubTitle = "Comment by Person In Charge : " + strPeriodType + " " + strYear

        Select Case CInt(Me.cboPeriodType.SelectedValue)

            Case CType(enumPeriodType.OriginalBudget, Integer) '//Original

                InsertCommentColumnData(dtColumns, strYear)

                '//Create group data
                Dim dsGroups As DataSet = SetupCommentGroupbyData(dsData, "PERSON_IN_CHARGE", "PERSON_IN_CHARGE_NAME", 5)

                '//Create Output Excel
                OutputExcel(dsGroups, dtColumns, False, strSubTitle, strYear, True, strPeriod)

            Case CType(enumPeriodType.EstimateBudget, Integer) '//Estimate

                InsertEstimateColumnData(dtColumns, strYear)

                '//Create group data
                Dim dsGroups As DataSet = SetupGroupbyData(dsData, "PERSON_IN_CHARGE_NO", "PERSON_IN_CHARGE_NAME", 10, True)

                '//Create Output Excel
                OutputExcel(dsGroups, dtColumns, False, strSubTitle, strYear, True, strPeriod)

            Case CType(enumPeriodType.ReviseBudget, Integer) '//Revise

                '//Create output columns
                If Me.chkShowMTP.Checked = True Then
                    InsertMTPColumnData(dtColumns, strYear)
                Else
                    InsertReviseColumnData(dtColumns, strYear)
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

    Private Function InsertCommentColumnData(ByRef dtColumns As DataTable, ByVal strYear As String) As Boolean

        Dim dRow As DataRow

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "MONTH"
        dRow("Column_Title") = "Month"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "COMMENT"
        dRow("Column_Title") = "Comment"
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

        Return True

    End Function

    Private Function InsertReviseColumnData(ByRef dtColumns As DataTable, ByVal strYear As String) As Boolean

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
        dRow("Column_Name") = "REVISE_TOTAL_YEAR"
        dRow("Column_Title") = "Original Year'" & strYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PrevRRT1"
        dRow("Column_Title") = "MTP" & intYear - 1 & " Year'" & CStr(intYear)
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_TOTAL_YEAR"
        dRow("Column_Title") = "Diff Year'" & strYear
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "RRT1"
        dRow("Column_Title") = "MTP" & intYear & " Year'" & CStr(intYear + 1)
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PrevRRT2"
        dRow("Column_Title") = "MTP" & intYear - 1 & " Year'" & CStr(intYear + 1)
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
        dRow("Column_Name") = "RRT3"
        dRow("Column_Title") = "MTP" & intYear & " Year'" & CStr(intYear + 3)
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PrevRRT4"
        dRow("Column_Title") = "MTP" & intYear - 1 & " Year'" & CStr(intYear + 3)
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "RRT4"
        dRow("Column_Title") = "MTP" & intYear & " Year'" & CStr(intYear + 4)
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PrevRRT5"
        dRow("Column_Title") = "MTP" & intYear - 1 & " Year'" & CStr(intYear + 4)
        dtColumns.Rows.Add(dRow)


        dRow = dtColumns.NewRow
        dRow("Column_Name") = "RRT5"
        dRow("Column_Title") = "MTP" & intYear & " Year'" & CStr(intYear + 5)
        dtColumns.Rows.Add(dRow)

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

        xBk = excelApp.Workbooks.Add(System.Reflection.Missing.Value)
        If xBk.Worksheets.Count > 1 Then
            For i As Integer = 1 To xBk.Worksheets.Count - 1
                CType(xBk.Worksheets(i), Excel.Worksheet).Delete()
                'CType(xBk.Worksheets(2), Excel.Worksheet).Delete()
            Next

        End If

        '//Set Style Value < 0 please fill color "Red"
        Dim style As Excel.Style = excelApp.ActiveWorkbook.Styles.Add("NewStyle")
        'style.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
        style.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)

        ''//Set Style Value Group fill Backgroud color "Gray"
        'Dim BGstyle As Excel.Style = excelApp.ActiveWorkbook.Styles.Add("NewStyle")
        'BGstyle.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)

        For intSheetCount As Integer = 0 To dsData.Tables.Count - 1

            If intSheetCount <> 0 Then
                xBk.Sheets.Add()
            End If

            rowStartIndex = 9
            colStartIndex = 8

            xSt = CType(xBk.ActiveSheet, Excel.Worksheet)
            Dim strTableName As String = dsData.Tables(intSheetCount).TableName
            Dim intIndexTblName As Integer = strTableName.IndexOf(" ")
            xSt.Name = strTableName.Substring(0, intIndexTblName)

            '//Setup DataColumn
            For i As Integer = 0 To dtColumns.Rows.Count - 1
                xSt.Cells(colStartIndex, i + 1) = dtColumns.Rows(i)("Column_Title").ToString
                xSt.Range(xSt.Cells(colStartIndex, i + 1), xSt.Cells(colStartIndex, i + 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

                If strPeriod = "MTP" Then
                    xSt.Cells(colStartIndex, 1) = "Year"
                End If
            Next


            '//Setup Data
            For rowIndex As Integer = 0 To dsData.Tables(intSheetCount).Rows.Count - 1

                Dim row As DataRow = dsData.Tables(intSheetCount).Rows(rowIndex)

                For colIndex As Integer = 0 To dtColumns.Rows.Count - 1

                    Dim strColumnName As String = dtColumns.Rows(colIndex)("Column_Name").ToString
                    Dim col As DataColumn = dsData.Tables(intSheetCount).Columns(strColumnName)

                    'Set BudgetOrder Style
                    If strColumnName = "COMMENT" Then
                        If String.IsNullOrEmpty(row("BUDGET_YEAR").ToString().Trim) Then
                            xSt.Range(xSt.Cells(rowIndex + rowStartIndex, colIndex), xSt.Cells(rowIndex + rowStartIndex, colIndex + 1)).Font.Bold = True
                            xSt.Range(xSt.Cells(rowIndex + rowStartIndex, colIndex), xSt.Cells(rowIndex + rowStartIndex, colIndex + 1)).MergeCells = True
                            xSt.Range(xSt.Cells(rowIndex + rowStartIndex, colIndex), xSt.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                            xSt.Range(xSt.Cells(rowIndex + rowStartIndex, colIndex), xSt.Cells(rowIndex + rowStartIndex, colIndex + 1)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)
                            Continue For
                        End If
                    End If

                    'Not output COMMENT EMPTY 
                    If (Not String.IsNullOrEmpty(row("BUDGET_YEAR").ToString().Trim)) And String.IsNullOrEmpty(row("COMMENT").ToString().Trim) Then
                        Continue For
                    End If

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

                        'If CDbl(row(col.ColumnName).ToString) = 0 Then
                        '    xSt.Cells(rowIndex + rowStartIndex, colIndex + 1) = "-"
                        '    xSt.Range(xSt.Cells(rowIndex + rowStartIndex, colIndex + 1), xSt.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                        'Else
                        '    xSt.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName)
                        '    xSt.Range(xSt.Cells(rowIndex + rowStartIndex, colIndex + 1), xSt.Cells(rowIndex + rowStartIndex, colIndex + 1)).NumberFormat = "#,##0.00"
                        'End If

                        '//Add by Max 01/10/2012
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

                    ''Set Group Header 
                    'If strColumnName = "COMMENT" Then
                    '    If String.IsNullOrEmpty(row(col.ColumnName).ToString().Trim) Then
                    '        xSt.Range(xSt.Cells(rowIndex + rowStartIndex, colIndex), xSt.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                    '        xSt.Range(xSt.Cells(rowIndex + rowStartIndex, colIndex), xSt.Cells(rowIndex + rowStartIndex, colIndex + 1)).Font.Bold = True
                    '        xSt.Range(xSt.Cells(rowIndex + rowStartIndex, colIndex), xSt.Cells(rowIndex + rowStartIndex, colIndex + 1)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)
                    '        xSt.Range(xSt.Cells(rowIndex + rowStartIndex, colIndex), xSt.Cells(rowIndex + rowStartIndex, colIndex + 1)).MergeCells = True
                    '    End If
                    'End If
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
            'Clear Unit
            xSt.Range(xSt.Cells(6, intUnitPriceStart), xSt.Cells(6, intUnitPriceEnd)).ClearContents()

            Dim rowMax As Integer = dsData.Tables(intSheetCount).Rows.Count + colStartIndex
            Dim colMax As Integer = dtColumns.Rows.Count

            '//Setup budget order name column to be left align
            xSt.Range(xSt.Cells(rowStartIndex, 2), xSt.Cells(rowMax, 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup Total Lines
            'SetupTotalLines(xSt, rowMax - 3, "Total", "Center", 1, 1, 4, rowMax - 2, rowMax - 1, rowMax, colMax, bMTPCheck)

            '//Setup sheet properly width
            xSt.Range(xSt.Cells(2, 1), xSt.Cells(rowMax, colMax)).Columns.AutoFit()

            '//Setup Wrap text for columns title
            'If strPeriod = "Original" Then

            '    xSt.Range(xSt.Cells(2, 4), xSt.Cells(rowMax, 4)).Columns.ColumnWidth = 9
            '    xSt.Range(xSt.Cells(2, 4), xSt.Cells(rowMax, 4)).WrapText = True

            '    xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 6)).Columns.ColumnWidth = 13
            '    xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 6)).WrapText = True

            '    xSt.Range(xSt.Cells(2, 7), xSt.Cells(rowMax, 12)).Columns.ColumnWidth = 12

            '    xSt.Range(xSt.Cells(2, 13), xSt.Cells(rowMax, 16)).Columns.ColumnWidth = 13
            '    xSt.Range(xSt.Cells(2, 13), xSt.Cells(rowMax, 16)).WrapText = True

            'ElseIf strPeriod = "Estimate" Then

            '    xSt.Range(xSt.Cells(2, 4), xSt.Cells(rowMax, 4)).Columns.ColumnWidth = 9
            '    xSt.Range(xSt.Cells(2, 4), xSt.Cells(rowMax, 4)).WrapText = True

            '    xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 6)).Columns.ColumnWidth = 13
            '    xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 6)).WrapText = True

            '    xSt.Range(xSt.Cells(2, 7), xSt.Cells(rowMax, 12)).Columns.ColumnWidth = 12

            '    xSt.Range(xSt.Cells(2, 13), xSt.Cells(rowMax, 15)).Columns.ColumnWidth = 13
            '    xSt.Range(xSt.Cells(2, 13), xSt.Cells(rowMax, 15)).WrapText = True
            'ElseIf strPeriod = "Revise" Then

            '    xSt.Range(xSt.Cells(2, 4), xSt.Cells(rowMax, 4)).Columns.ColumnWidth = 9
            '    xSt.Range(xSt.Cells(2, 4), xSt.Cells(rowMax, 4)).WrapText = True

            '    xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 5)).Columns.ColumnWidth = 13
            '    xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 5)).WrapText = True

            '    xSt.Range(xSt.Cells(2, 6), xSt.Cells(rowMax, 11)).Columns.ColumnWidth = 12

            '    xSt.Range(xSt.Cells(2, 12), xSt.Cells(rowMax, 14)).Columns.ColumnWidth = 13
            '    xSt.Range(xSt.Cells(2, 12), xSt.Cells(rowMax, 14)).WrapText = True

            '    xSt.Range(xSt.Cells(2, 15), xSt.Cells(rowMax, 20)).Columns.ColumnWidth = 12

            '    xSt.Range(xSt.Cells(2, 21), xSt.Cells(rowMax, 24)).Columns.ColumnWidth = 13
            '    xSt.Range(xSt.Cells(2, 21), xSt.Cells(rowMax, 24)).WrapText = True

            '    'If chkShowMTP.Checked = True Then
            '    '    xSt.Range(xSt.Cells(2, 25), xSt.Cells(rowMax, 29)).Columns.ColumnWidth = 12
            '    'End If

            'ElseIf strPeriod = "MTP" Then

            '    xSt.Range(xSt.Cells(2, 4), xSt.Cells(rowMax, 4)).Columns.ColumnWidth = 9
            '    xSt.Range(xSt.Cells(2, 4), xSt.Cells(rowMax, 4)).WrapText = True

            '    xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 7)).Columns.ColumnWidth = 13
            '    xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 7)).WrapText = True

            '    xSt.Range(xSt.Cells(2, 8), xSt.Cells(rowMax, 16)).Columns.ColumnWidth = 13
            '    xSt.Range(xSt.Cells(2, 8), xSt.Cells(rowMax, 16)).WrapText = True
            'End If

            'colStartIndex = colStartIndex - 1
            '//Setup Column Font 
            xSt.Range(xSt.Cells(colStartIndex, 1), xSt.Cells(colStartIndex, colMax)).Font.Bold = True
            xSt.Range(xSt.Cells(colStartIndex, 1), xSt.Cells(rowMax, colMax)).Font.Name = "Tahoma"
            xSt.Range(xSt.Cells(colStartIndex, 1), xSt.Cells(rowMax, colMax)).Font.Size = 10

            '//Setup border
            xSt.Range(xSt.Cells(colStartIndex, 1), xSt.Cells(rowMax, colMax)).Borders.LineStyle = 1
            xSt.Range(xSt.Cells(colStartIndex, 1), xSt.Cells(rowMax, 1)).Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlThin
            xSt.Range(xSt.Cells(colStartIndex, 1), xSt.Cells(colStartIndex, colMax)).Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThin
            xSt.Range(xSt.Cells(colStartIndex, colMax), xSt.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThin
            xSt.Range(xSt.Cells(rowMax, 1), xSt.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThin


            ''//Insert empty column
            'If strPeriod = "MTP" AndAlso bMTPCheck = True Then
            '    '   SetupMTPEmptyColumn(xSt, colStartIndex, rowMax, colMax, 25, rowMax - 2, 1)
            'ElseIf bMTPCheck = True Then
            '    'SetupMTPEmptyColumn(xSt, colStartIndex, rowMax, colMax, 25, rowMax - 2, 1)
            'End If

            '//-- Add by Max 26/09/2012
            '//Set Frame All
            'xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax, colMax)).Borders.LineStyle = 1
            'xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium
            'xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium
            'xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            'xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium

            ''//Set Frame
            'If strPeriod = "Original" Then

            '    xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax, 6)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            '    xSt.Range(xSt.Cells(colStartIndex, 7), xSt.Cells(rowMax, 12)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            '    xSt.Range(xSt.Cells(colStartIndex, 13), xSt.Cells(rowMax, 14)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            '    xSt.Range(xSt.Cells(colStartIndex, 15), xSt.Cells(rowMax, 16)).Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium
            '    xSt.Range(xSt.Cells(colStartIndex, 17), xSt.Cells(rowMax, 18)).Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium

            'ElseIf strPeriod = "Estimate" Then

            '    xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax, 6)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            '    xSt.Range(xSt.Cells(colStartIndex, 7), xSt.Cells(rowMax, 9)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            '    xSt.Range(xSt.Cells(colStartIndex, 10), xSt.Cells(rowMax, 12)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            '    xSt.Range(xSt.Cells(colStartIndex, 13), xSt.Cells(rowMax, 15)).Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium

            'ElseIf strPeriod = "Revise" Then

            '    xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax, 5)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            '    If bMTPCheck = False Then
            '        xSt.Range(xSt.Cells(colStartIndex, 6), xSt.Cells(rowMax, 8)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            '    End If
            '    xSt.Range(xSt.Cells(colStartIndex, 9), xSt.Cells(rowMax, 11)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            '    xSt.Range(xSt.Cells(colStartIndex, 12), xSt.Cells(rowMax, 14)).Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium
            '    xSt.Range(xSt.Cells(colStartIndex, 14), xSt.Cells(rowMax, 14)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium

            '    If bMTPCheck = False Then
            '        xSt.Range(xSt.Cells(colStartIndex, 15), xSt.Cells(rowMax, 20)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            '        xSt.Range(xSt.Cells(colStartIndex, 21), xSt.Cells(rowMax, 24)).Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium
            '    End If

            '    '//Merge empty line
            '    If bMTPCheck = True Then
            '        xSt.Range(xSt.Cells(rowMax - 2, 1), xSt.Cells(rowMax - 2, colMax)).ClearContents()
            '        xSt.Range(xSt.Cells(rowMax - 2, 1), xSt.Cells(rowMax - 2, colMax)).MergeCells = True
            '    End If

            'ElseIf strPeriod = "MTP" Then

            '    xSt.Range(xSt.Cells(colStartIndex, 5), xSt.Cells(rowMax, 6)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            '    xSt.Range(xSt.Cells(colStartIndex, 7), xSt.Cells(rowMax, 7)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium

            '    '//Set font color
            '    xSt.Range(xSt.Cells(colStartIndex, 6), xSt.Cells(rowMax, 7)).Font.Color = RGB(128, 128, 128)
            '    xSt.Range(xSt.Cells(colStartIndex, 9), xSt.Cells(rowMax, 9)).Font.Color = RGB(128, 128, 128)
            '    xSt.Range(xSt.Cells(colStartIndex, 11), xSt.Cells(rowMax, 11)).Font.Color = RGB(128, 128, 128)
            '    xSt.Range(xSt.Cells(colStartIndex, 13), xSt.Cells(rowMax, 13)).Font.Color = RGB(128, 128, 128)
            '    xSt.Range(xSt.Cells(colStartIndex, 15), xSt.Cells(rowMax, 15)).Font.Color = RGB(128, 128, 128)

            'End If
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

    Private Function SetupTitleIndex(ByVal strPeriod As String, ByRef intUnitPriceStart As Integer, _
                                     ByRef intUnitPriceEnd As Integer, ByRef intAuthorizeStart As Integer, _
                                     ByRef intAuthorizeEnd As Integer, ByRef intImageIndex As Integer, ByVal bMTPCheck As Boolean) As Boolean

        Select Case strPeriod

            Case "Original"

                intUnitPriceStart = 16
                intUnitPriceEnd = 17

                intAuthorizeStart = 13
                intAuthorizeEnd = 14

                intImageIndex = 815

            Case "Estimate"

                intUnitPriceStart = 14
                intUnitPriceEnd = 15

                intAuthorizeStart = 11
                intAuthorizeEnd = 13

                intImageIndex = 755

            Case "Revise"

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

                intUnitPriceStart = 15
                intUnitPriceEnd = 16

                intAuthorizeStart = 1
                intAuthorizeEnd = 2

                intImageIndex = 815

            Case Else
                Exit Select
        End Select

        Return True

    End Function

    Private Function SetupCommentGroupbyData(ByVal dsData As DataSet, _
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

            strScript = strGroupColumnName + " = '" + dtGroups.Rows(i)(0).ToString & "'"
            Dim arrRows As DataRow() = dsData.Tables(0).Select(strScript)

            Dim strGroupColumnName2 As String = "BUDGET_ORDER_NO"
            '//Seperate dataset data into several datatables according to group no
            strScript = strGroupColumnName2
            Dim dtGroups2 As DataTable = dsData.Tables(0).DefaultView.ToTable(True, strScript)
            Dim intGroupCount2 As Integer = dtGroups2.Rows.Count

            For n As Integer = 0 To intGroupCount2 - 1

                Dim drTotalExpenses As DataRow = dtResult.NewRow

                strScript = strGroupColumnName + " = '" + dtGroups.Rows(i)(0).ToString
                strScript &= "' AND "
                strScript &= strGroupColumnName2 + " = '" + dtGroups2.Rows(n)(0).ToString + "'"
                Dim arrRows2 As DataRow() = dsData.Tables(0).Select(strScript)

                If arrRows2.Length > 0 Then
                    For j As Integer = 0 To arrRows2.Length - 1
                        Dim drow(dtResult.Columns.Count - 1) As Object
                        arrRows2(j).ItemArray.CopyTo(drow, 0)
                        dtResult.Rows.Add(drow)
                    Next

                    '//Setup Group header
                    drTotalExpenses("MONTH") = arrRows2(0)(strGroupColumnName2).ToString + " " + arrRows2(0)("BUDGET_ORDER_NAME").ToString

                    '//Add total cost
                    dtResult.Rows.InsertAt(drTotalExpenses, intGroupTotalIndex)

                    ''//Add one empty row
                    'drEmpty = dtResult.NewRow
                    'dtResult.Rows.Add(drEmpty)

                    'intGroupTotalIndex = intGroupTotalIndex + dtResult.Rows.Count
                    intGroupTotalIndex = intGroupTotalIndex + CInt(arrRows2.Length) + 1
                End If



            Next

            ''//Calculate Total cost
            'For l As Integer = intDataColumnIndex To dtResult.Columns.Count - 1
            '    Dim strColumnName As String = dtResult.Columns(l).ColumnName

            '    strExpression = "Sum(" + strColumnName + ")"
            '    strFilter = strGroupColumnName + " = " + dtGroups.Rows(i)(0).ToString
            '    returnValue = dtResult.Compute(strExpression, strFilter)
            '    drTotalCost(dtResult.Columns(l).ColumnName) = returnValue

            'Next
            'drTotalCost("ACCOUNT_NO") = "Total Cost"

            ''//Add one empty row
            'drEmpty = dtResult.NewRow
            'dtResult.Rows.Add(drEmpty)

            ''//Add total cost
            'dtResult.Rows.Add(drTotalCost)

            ''//Add on empty row at index 0
            'drEmpty = dtResult.NewRow
            'dtResult.Rows.InsertAt(drEmpty, 0)

            dtResult.TableName = arrRows(0)(strGroupColumnName).ToString & " " & arrRows(0)(strGroupColumnTitle).ToString
            'dtResult.TableName = BGCommon.GetGroupCostTitle(arrRows(0)(strGroupColumnName).ToString)

            ''//Return data table
            dsResult.Tables.Add(dtResult)

        Next

        Return dsResult

    End Function

    Private Sub cboPeriodType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPeriodType.SelectedIndexChanged

        If CInt(cboPeriodType.SelectedValue) = CType(enumPeriodType.ReviseBudget, Integer) Then
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

        'If CType(cboPeriodType.SelectedValue, Integer) = CType(enumPeriodType.MTPBudget, Integer) OrElse _
        '    CType(cboPeriodType.SelectedValue, Integer) = CType(enumPeriodType.OriginalBudget, Integer) Then
        '    EnablePrev()
        'Else
        '    DisablePrev()
        'End If

        LoadRevNo()

    End Sub

    Private Sub numYear_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numYear.ValueChanged

        LoadRevNo()
        'LoadPrevRevNo()

    End Sub

    Private Sub numProjectNo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numProjectNo.ValueChanged

        LoadRevNo()

    End Sub


    Private Sub numPrevProjectNo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numPrevProjectNo.ValueChanged
        'LoadPrevRevNo()
    End Sub



#End Region

End Class