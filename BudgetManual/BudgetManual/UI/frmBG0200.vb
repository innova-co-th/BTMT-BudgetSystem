Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class frmBG0200

#Region "Variable"
    Private myClsBG0201BL As New clsBG0201BL

    Private Const GRIDVIEW_TOP1 As Integer = 203
    Private Const GRIDVIEW_TOP2 As Integer = 135
    Private Const GRID_HEIGHT_ADD As Integer = 55
    Private myClsBG0200BL As New clsBG0200BL
    Private myClsBG0310BL As New clsBG0310BL
    Private myBudgetKey As String = String.Empty
    Private myOperationCd As enumOperationCd
    Private myCurrRevNo As String = String.Empty
    Private myCurrReviseRevNo As String = String.Empty
    Private myFormLoadedFlg As Boolean = False
    Private myForceCloseFlg As Boolean = False
    Private myDataLoadingFlg As Boolean = False
    Private myDtAllData As DataTable = Nothing
    Private myBudgetDataChanged As Boolean = False
    Private myShowWarning As Boolean = False
    Private mySetGridValue As Boolean = False

    Private m_dtCheckMTP As DataTable = Nothing
    Private m_dtCheckMTPNew As DataTable = Nothing


    Private mydtBG1 As DataTable = Nothing
    Private mydtBG2 As DataTable = Nothing
    Private mydtBG3 As DataTable = Nothing
    Private mydtBG4 As DataTable = Nothing


    Private mydtBG1View As DataTable = Nothing
    Private mydtBG2View As DataTable = Nothing
    Private mydtBG3View As DataTable = Nothing
    Private mydtBG4View As DataTable = Nothing

    Private blnReInputByOrder As Boolean = False
#End Region

#Region "Property"

#Region "Budget Key"
    Property BudgetKey() As String
        Get
            Return myBudgetKey
        End Get
        Set(ByVal value As String)
            myBudgetKey = value
        End Set
    End Property
#End Region

#Region "Operation Code"
    Property OperationCd() As enumOperationCd
        Get
            Return myOperationCd
        End Get
        Set(ByVal value As enumOperationCd)
            myOperationCd = value
        End Set
    End Property
#End Region

#Region "Current Rev No"
    Property CurrRevNo() As String
        Get
            Return myCurrRevNo
        End Get
        Set(ByVal value As String)
            myCurrRevNo = value
        End Set
    End Property
#End Region

#Region "Current Revise Rev No"
    Property CurrReviseRevNo() As String
        Get
            Return myCurrReviseRevNo
        End Get
        Set(ByVal value As String)
            myCurrReviseRevNo = value
        End Set
    End Property
#End Region

#End Region

#Region "Overrides Function"
    Public Sub New(ByRef frmParent As Form, ByVal blnMaximize As Boolean, ByVal strFormTitle As String, ByVal strBudgetKey As String, ByVal intOperationCd As enumOperationCd)
        Try
            ' This call is required by the Windows Form Designer.
            InitializeComponent()

            ' Add any initialization after the InitializeComponent() call.
            Me.MdiParent = frmParent
            If blnMaximize Then
                Me.WindowState = FormWindowState.Maximized
            Else
                Me.WindowState = FormWindowState.Normal
            End If
            Me.Text = strFormTitle
            Me.BudgetKey = strBudgetKey
            Me.OperationCd = intOperationCd
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
#End Region

#Region "Function"

    Private Function GetBudgetYear() As String
        Try
            Return Me.BudgetKey.Substring(0, 4)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function GetPeriodType() As String
        Dim strPeriod As String = String.Empty

        Try
            If BudgetKey <> "" Then
                If BudgetKey.Substring(4, 2).IndexOf("0") = 0 Then
                    strPeriod = BudgetKey.Substring(5, 1)
                Else
                    strPeriod = BudgetKey.Substring(4, 2)
                End If
            End If
            
        Catch ex As Exception
            Throw ex
        End Try

        Return strPeriod
    End Function

    Private Function GetBudgetType() As String
        Try
            Return Me.BudgetKey.Substring(6, 1)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function GetProjectNo() As String
        Try
            Return Me.BudgetKey.Substring(7, Me.BudgetKey.Length - 7)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function GetMtpProjectNo() As String
        Dim strMtpProjectNo As String = String.Empty

        Try
            strMtpProjectNo = Me.numMtpProjectNo.Value.ToString()
        Catch ex As Exception
            Throw ex
        End Try

        Return strMtpProjectNo
    End Function

    Private Function GetMtpRevNo() As String
        Dim strMtpRevNo As String = String.Empty

        Try
            If Me.cboMtpRevno.DataSource IsNot Nothing AndAlso _
          Me.cboMtpRevno.SelectedValue IsNot Nothing Then
                strMtpRevNo = Me.cboMtpRevno.SelectedValue.ToString
            End If
        Catch ex As Exception
            Throw ex
        End Try

        Return strMtpRevNo
    End Function

    Private Sub ShowBudgetData()
        Dim dtmLoadTime As Date = Now

        Try

            Debug.Print(Now.ToString() & ": Begin LoadBudgetData")

            Me.Cursor = Cursors.WaitCursor

            myDataLoadingFlg = True

            '// Clear datagrid for check MTP 
            m_dtCheckMTP = Nothing
            m_dtCheckMTPNew = Nothing

            '// Load Transfer List
            LoadTransferList()

            '// Load Upload Data
            LoadUploadData()

            '// Clear controls
            ClearControls()

            '// Show Datagrid
            If ShowDatagrid() = True Then

                '// Show Upload Data
                ShowUploadData()

                '// Remember datagrid for check MTP 
                If Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then
                    m_dtCheckMTP = CType(grvBudget3.DataSource, DataTable).DefaultView.Table.Copy
                    '//-- End Edit 2011-05-27
                End If

                If Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then
                    m_dtCheckMTPNew = CType(grvBudget4.DataSource, DataTable).DefaultView.Table.Copy

                    '// Show Investment for MTP
                    ShowMTPInvestment()
                End If

                '// Show Total Working budget H1 & H2 and MTP Summary
                If Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput Then
                    If Me.GetBudgetType() = P_BUDGET_TYPE_EXPENSE Then
                        ShowWKH()
                    Else
                        ShowMTP_SUM()
                    End If
                End If

                '// Set Controls
                SetButtons()

                '// Calculate Total/Diff
                If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then
                    CalcOriginalBudget()

                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then
                    CalcEstimateBudget()

                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then
                    CalcReviseBudget(False)
                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then
                    CalcReviseMTPBudget()
                End If

                '// Highlight Datagrid
                If CDbl(Me.myCurrRevNo) > 1 And _
                (Me.OperationCd = enumOperationCd.AdjustBudget Or _
                Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput Or _
                 Me.OperationCd = enumOperationCd.Authorize1 Or _
                 Me.OperationCd = enumOperationCd.Authorize2) Then
                    '// Highlight Working Budget
                    HighlightWorkingBG()

                    '// Highlight Transfer Cost
                    HighlightTransferCost()
                End If

                If Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) And _
                    Me.GetBudgetType() = P_BUDGET_TYPE_EXPENSE Then
                    HighlightMTPValueAllNew()
                End If

                HighlightWorkingBGAndComment()
            End If

            myDataLoadingFlg = False
            myForceCloseFlg = False
            myBudgetDataChanged = False

            Me.Cursor = Cursors.Default

            Debug.Print(Now.ToString() & ": End LoadBudgetData")

            Debug.Print("Loading Time: " & DateDiff(DateInterval.Second, dtmLoadTime, Now).ToString("#,##0") & " sec(s)" & vbNewLine)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub ClearControls()
        Try
            '// Clear controls
            lblBudgetPeriod.Text = ""
            lblStatus.Text = ""
            lblRevNo.Text = ""
            txtRRT0.Text = ""
            lblRRT0.Text = ""
            txtRRT1.Text = ""
            lblRRT1.Text = ""
            lblRRT1p.Text = ""
            txtRRT2.Text = ""
            lblRRT2.Text = ""
            lblRRT2p.Text = ""
            txtRRT3.Text = ""
            lblRRT3.Text = ""
            lblRRT3p.Text = ""
            txtRRT4.Text = ""
            lblRRT4.Text = ""
            lblRRT4p.Text = ""
            txtRRT5.Text = ""
            lblRRT5.Text = ""
            lblRRT5p.Text = ""
            lblRefRevNo.Text = ""

            '// Show/Hide GridView
            If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then   '// Original Budget
                grvBudget1.Visible = True
                grvBudget2.Visible = False
                grvBudget3.Visible = False
                grvBudget4.Visible = False
                fraMTP.Visible = False

                pnlSummary.Visible = True
                pnlSummaryMTP.Visible = False

                If grvBudget1.Top <> GRIDVIEW_TOP2 Then
                    grvBudget1.Top = GRIDVIEW_TOP2
                End If

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then   '// Estimate Budget
                grvBudget1.Visible = False
                grvBudget2.Visible = True
                grvBudget3.Visible = False
                grvBudget4.Visible = False
                fraMTP.Visible = False

                pnlSummary.Visible = True
                pnlSummaryMTP.Visible = False

                If grvBudget2.Top <> GRIDVIEW_TOP2 Then
                    grvBudget2.Top = GRIDVIEW_TOP2
                End If

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then   '// Revise Budget
                grvBudget1.Visible = False
                grvBudget2.Visible = False
                grvBudget3.Visible = True
                grvBudget4.Visible = False

                pnlSummary.Visible = True
                pnlSummaryMTP.Visible = False

                fraMTP.Visible = False

                If grvBudget3.Top <> GRIDVIEW_TOP2 Then
                    grvBudget3.Top = GRIDVIEW_TOP2
                End If

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then   '// MTP Budget
                grvBudget1.Visible = False
                grvBudget2.Visible = False
                grvBudget3.Visible = False
                grvBudget4.Visible = True

                pnlSummary.Visible = False
                pnlSummaryMTP.Visible = True

                If Me.GetBudgetType() = P_BUDGET_TYPE_EXPENSE Then
                    fraMTP.Visible = True

                    If grvBudget4.Top <> GRIDVIEW_TOP1 Then
                        grvBudget4.Top = GRIDVIEW_TOP1
                        grvBudget4.Height = grvBudget4.Height - (GRIDVIEW_TOP1 - GRIDVIEW_TOP2)

                        If Me.OperationCd <> enumOperationCd.AdjustBudget Or Me.OperationCd <> enumOperationCd.AdjustBudgetDirectInput Then
                            grvBudget4.Height = grvBudget4.Height + GRID_HEIGHT_ADD
                        End If
                    End If
                Else
                    fraMTP.Visible = False

                    If grvBudget4.Top <> GRIDVIEW_TOP2 Then
                        grvBudget4.Top = GRIDVIEW_TOP2
                    End If
                End If
            End If

            '// Clear Datagrid's items
            Dim dt As DataTable
            If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then   '// Original Budget
                dt = CType(grvBudget1.DataSource, DataTable)
                If dt IsNot Nothing Then
                    dt.Rows.Clear()
                    grvBudget1.DataSource = dt
                End If

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then   '// Estimate Budget
                dt = CType(grvBudget2.DataSource, DataTable)
                If dt IsNot Nothing Then
                    dt.Rows.Clear()
                    grvBudget2.DataSource = dt
                End If

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then   '// Revise Budget
                dt = CType(grvBudget3.DataSource, DataTable)
                If dt IsNot Nothing Then
                    dt.Rows.Clear()
                    grvBudget3.DataSource = dt
                End If
            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then   '// MTP Budget
                dt = CType(grvBudget4.DataSource, DataTable)
                If dt IsNot Nothing Then
                    dt.Rows.Clear()
                    grvBudget4.AutoGenerateColumns = False
                    grvBudget4.DataSource = dt
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function ShowDatagrid() As Boolean
        Dim dtGrid As New DataTable
        Dim dc As DataColumn
        Dim dr As DataRow

        Try
            blnReInputByOrder = False

            Debug.Print(Now.ToString() & ": Begin ShowDatagrid")

            ShowDatagrid = False

            '// Check Budget Header
            myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
            myClsBG0200BL.PeriodType = Me.GetPeriodType()
            myClsBG0200BL.BudgetType = Me.GetBudgetType()
            myClsBG0200BL.UserPIC = CStr(cboPIC.SelectedValue)
            myClsBG0200BL.RevNo = Me.CurrRevNo
            myClsBG0200BL.UserId = p_strUserId
            myClsBG0200BL.OperationCd = Me.OperationCd
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            myClsBG0200BL.MtpProjectNo = Me.GetMtpProjectNo()
            myClsBG0200BL.MtpRevNo = Me.GetMtpRevNo()

            Dim blnInputHeader As Boolean = False
            If myClsBG0200BL.GetBudgetHeader() = False Then
                If Me.OperationCd = enumOperationCd.InputBudget Then
                    '// Create Budget data
                    myClsBG0200BL.RevNo = "1"
                    If myClsBG0200BL.CreateBudgetData() = False Then
                        MessageBox.Show("Can not load budget data!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        myForceCloseFlg = True

                        Exit Function
                    End If
                End If
            Else
                If Me.OperationCd = enumOperationCd.InputBudget Then
                    If myClsBG0200BL.Status = CStr(enumBudgetStatus.NewRecord) Then
                        blnInputHeader = True
                        myClsBG0200BL.RevNo = "1"
                        If myClsBG0200BL.SearchNewBudgetOrder() = True AndAlso myClsBG0200BL.OrderList.Rows.Count > 0 Then
                            '// Create Budget data
                            If myClsBG0200BL.CreateBudgetData2() = False Then
                                MessageBox.Show("Can not load budget data!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                                myForceCloseFlg = True

                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If

            '// Clear Record Num
            lblRecNum.Text = ""

            '// Load Budget Data
            Dim blnResult As Boolean = False
            Dim DtGetFromHeader As DataTable
            Dim dtClone As DataTable
            Dim dtDataReInput As DataTable
            If Me.OperationCd = enumOperationCd.InputBudget Then

                '--Get data FROM BG_T_BUDGET_DATA_REINPUT
                myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
                myClsBG0200BL.PeriodType = Me.GetPeriodType()
                myClsBG0200BL.BudgetType = Me.GetBudgetType()
                myClsBG0200BL.UserPIC = CStr(cboPIC.SelectedValue)
                myClsBG0200BL.RevNo = Me.CurrRevNo
                myClsBG0200BL.Status = CStr(enumBudgetStatus.NewRecord)
                myClsBG0200BL.ProjectNo = Me.GetProjectNo()
                dtDataReInput = myClsBG0200BL.GetBudGetDataReInput
                If Not dtDataReInput Is Nothing AndAlso dtDataReInput.Rows.Count > 0 Then
                    blnInputHeader = False
                Else
                    blnInputHeader = True
                End If

                If blnInputHeader = True Then
                    blnResult = myClsBG0200BL.GetBudgetData()
                Else
                    '---change to reinputbyorder
                    blnResult = myClsBG0200BL.GetBudgetData()
                    If Not myClsBG0200BL.BudgetList Is Nothing AndAlso myClsBG0200BL.BudgetList.Rows.Count > 0 Then
                        DtGetFromHeader = myClsBG0200BL.BudgetList
                        dtClone = DtGetFromHeader.Clone
                        Dim drSel() As DataRow
                        Dim drC As DataRow
                        '--Get data FROM BG_T_BUDGET_DATA_REINPUT

                        myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
                        myClsBG0200BL.PeriodType = Me.GetPeriodType()
                        myClsBG0200BL.BudgetType = Me.GetBudgetType()
                        myClsBG0200BL.UserPIC = CStr(cboPIC.SelectedValue)
                        myClsBG0200BL.RevNo = Me.CurrRevNo
                        myClsBG0200BL.Status = CStr(enumBudgetStatus.NewRecord)
                        myClsBG0200BL.ProjectNo = Me.GetProjectNo()

                        dtDataReInput = myClsBG0200BL.GetBudGetDataReInput
                        If Not dtDataReInput Is Nothing AndAlso dtDataReInput.Rows.Count > 0 Then
                            For i As Integer = 0 To dtDataReInput.Rows.Count - 1
                                drSel = Nothing
                                drSel = DtGetFromHeader.Select("BUDGET_ORDER_NO='" & CStr(dtDataReInput.Rows(i).Item("BUDGET_ORDER_NO")) & "'")
                                If Not drSel Is Nothing AndAlso drSel.Length > 0 Then
                                    drC = dtClone.NewRow

                                    For index As Integer = 0 To dtClone.Columns.Count - 1
                                        drC(index) = drSel(0).Item(index)
                                    Next
                                    dtClone.Rows.Add(drC)
                                End If
                            Next
                        End If
                        If Not dtClone Is Nothing AndAlso dtClone.Rows.Count > 0 Then
                            myClsBG0200BL.BudgetList = dtClone
                            blnResult = True
                            blnReInputByOrder = True
                        Else
                            blnResult = False
                        End If
                    Else
                        blnResult = False
                    End If
                End If
            Else
                If Me.OperationCd = enumOperationCd.ApproveBudget Then
                    'Check ReInputByOrder 
                    myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
                    myClsBG0200BL.PeriodType = Me.GetPeriodType()
                    myClsBG0200BL.BudgetType = Me.GetBudgetType()
                    myClsBG0200BL.UserPIC = CStr(cboPIC.SelectedValue)
                    myClsBG0200BL.RevNo = Me.CurrRevNo
                    myClsBG0200BL.Status = CStr(enumBudgetStatus.Submit)
                    myClsBG0200BL.ProjectNo = Me.GetProjectNo()

                    dtDataReInput = myClsBG0200BL.GetBudGetDataReInput
                    If Not dtDataReInput Is Nothing AndAlso dtDataReInput.Rows.Count > 0 Then
                        blnInputHeader = False
                    Else
                        blnInputHeader = True
                    End If

                    If blnInputHeader = True Then
                        blnResult = myClsBG0200BL.GetBudgetData()
                    Else
                        '---change to reinputbyorder
                        blnResult = myClsBG0200BL.GetBudgetData()
                        If Not myClsBG0200BL.BudgetList Is Nothing AndAlso myClsBG0200BL.BudgetList.Rows.Count > 0 Then
                            DtGetFromHeader = myClsBG0200BL.BudgetList
                            dtClone = DtGetFromHeader.Clone
                            Dim drSel() As DataRow
                            Dim drC As DataRow

                            If Not dtDataReInput Is Nothing AndAlso dtDataReInput.Rows.Count > 0 Then
                                For i As Integer = 0 To dtDataReInput.Rows.Count - 1
                                    drSel = Nothing
                                    drSel = DtGetFromHeader.Select("BUDGET_ORDER_NO='" & CStr(dtDataReInput.Rows(i).Item("BUDGET_ORDER_NO")) & "'")
                                    If Not drSel Is Nothing AndAlso drSel.Length > 0 Then
                                        drC = dtClone.NewRow

                                        For index As Integer = 0 To dtClone.Columns.Count - 1
                                            drC(index) = drSel(0).Item(index)
                                        Next
                                        dtClone.Rows.Add(drC)
                                    End If
                                Next
                            End If
                            If Not dtClone Is Nothing AndAlso dtClone.Rows.Count > 0 Then
                                myClsBG0200BL.BudgetList = dtClone
                                blnResult = True
                                blnReInputByOrder = True
                            Else
                                blnResult = False
                            End If
                        Else
                            blnResult = False
                        End If
                    End If
                Else
                    blnResult = myClsBG0200BL.GetBudgetData()
                End If
            End If

            If blnResult = False OrElse myClsBG0200BL.BudgetList.Rows.Count = 0 Then
                MessageBox.Show("Budget data not found!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                myForceCloseFlg = True

                Exit Function
            Else
                myForceCloseFlg = False

                '// Add Column
                dc = New DataColumn("Adjust", GetType(Boolean))
                dc.DefaultValue = False
                dtGrid.Columns.Add(dc)
                dc = New DataColumn("OrderNo", GetType(String))
                dtGrid.Columns.Add(dc)
                dc = New DataColumn("BudgetOrder", GetType(String))
                dtGrid.Columns.Add(dc)
                dc = New DataColumn("Account", GetType(String))
                dtGrid.Columns.Add(dc)
                dc = New DataColumn("CostType", GetType(String))
                dtGrid.Columns.Add(dc)
                dc = New DataColumn("Cost", GetType(String))
                dtGrid.Columns.Add(dc)
                dc = New DataColumn("Dept", GetType(String))
                dtGrid.Columns.Add(dc)
                dc = New DataColumn("Pic", GetType(String))
                dtGrid.Columns.Add(dc)

                If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then   '// Original Budget

                    dc = New DataColumn("IMP1", GetType(Double)) '// Prev. Actual H1
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("IMP2", GetType(Double)) '// Prev. Revise H2
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("M1", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M2", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M3", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M4", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M5", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M6", GetType(String))
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("Total1H", GetType(Double))
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("M7", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M8", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M9", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M10", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M11", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M12", GetType(String))
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("Total2H", GetType(Double))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("TotalY1", GetType(Double))
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("MTP_RRT1", GetType(Double))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("DiffMTP_RRT1", GetType(Double))
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("TotalY2", GetType(Double))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("Diff", GetType(Double))
                    dtGrid.Columns.Add(dc)

                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then   '// Estimate Budget

                    dc = New DataColumn("IMP1", GetType(Double)) '// Actual H1
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("IMP2", GetType(Double)) '// Revise H2
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("M7", GetType(Double))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M8", GetType(Double))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M9", GetType(Double))
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("M10", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M11", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M12", GetType(String))
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("Est2H", GetType(Double))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("Diff2H", GetType(Double))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("EstTotal", GetType(Double))
                    dtGrid.Columns.Add(dc)

                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then   '// Revise Budget

                    dc = New DataColumn("IMP1", GetType(Double)) '// Actual H1
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("M1", GetType(Double))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M2", GetType(Double))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M3", GetType(Double))
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("M4", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M5", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M6", GetType(String))
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("Est1H", GetType(Double))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("Diff1H", GetType(Double))
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("IMP2", GetType(Double)) '// Actual H2
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("M7", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M8", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M9", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M10", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M11", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("M12", GetType(String))
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("Rev2H", GetType(Double))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("Diff2H", GetType(Double))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("RevYear", GetType(Double))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("DiffYear", GetType(Double))
                    dtGrid.Columns.Add(dc)

                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then   '// MTP Budget

                    dc = New DataColumn("RevYear", GetType(Double))
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("PrevRRT1", GetType(Double))
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("PrevRRT2", GetType(Double))
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("PrevRRT3", GetType(Double))
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("PrevRRT4", GetType(Double))
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("PrevRRT5", GetType(Double))
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("DiffYear", GetType(Double))
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("DiffYear1", GetType(Double))
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("DiffYear2", GetType(Double))
                    dtGrid.Columns.Add(dc)

                    dc = New DataColumn("CALC1", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("CALC2", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("CALC3", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("CALC4", GetType(String))
                    dtGrid.Columns.Add(dc)
                    dc = New DataColumn("CALC5", GetType(String))
                    dtGrid.Columns.Add(dc)
                End If

                dc = New DataColumn("RRT1", GetType(String))
                dtGrid.Columns.Add(dc)
                dc = New DataColumn("RRT2", GetType(String))
                dtGrid.Columns.Add(dc)
                dc = New DataColumn("RRT3", GetType(String))
                dtGrid.Columns.Add(dc)
                dc = New DataColumn("RRT4", GetType(String))
                dtGrid.Columns.Add(dc)
                dc = New DataColumn("RRT5", GetType(String))
                dtGrid.Columns.Add(dc)

                dc = New DataColumn("Remarks", GetType(String))
                dtGrid.Columns.Add(dc)

                '// Prepare data table
                For Each drDat As DataRow In myClsBG0200BL.BudgetList.Rows
                    dr = dtGrid.NewRow

                    dr("Adjust") = CBool(IIf(CInt(drDat("WB_FLAG")) = 1, True, False))
                    dr("OrderNo") = CStr(drDat("BUDGET_ORDER_NO"))
                    dr("BudgetOrder") = CStr(drDat("BUDGET_ORDER_NO")) & " : " & CStr(drDat("BUDGET_ORDER_NAME"))
                    dr("Account") = CStr(drDat("ACCOUNT_NO"))
                    If CInt(Nz(drDat("COST_TYPE"), 0)) = enumCostType.FixedCost Then
                        dr("CostType") = "Fixed Cost"
                    ElseIf CInt(Nz(drDat("COST_TYPE"), 0)) = enumCostType.VariableCost Then
                        dr("CostType") = "Variable Cost"
                    End If
                    If CInt(Nz(drDat("COST"), 0)) = enumCost.ADMIN Then
                        dr("Cost") = "ADMIN"
                    ElseIf CInt(Nz(drDat("COST"), 0)) = enumCost.FC Then
                        dr("Cost") = "FC"
                    End If
                    dr("Dept") = CStr(drDat("DEPT_NO"))
                    dr("Pic") = CStr(drDat("PERSON_IN_CHARGE_NO"))

                    If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then   '// Original Budget
                        If CDec(Nz(drDat("M1"), 0)) <> 0 Then
                            dr("M1") = CDbl(drDat("M1")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("M2"), 0)) <> 0 Then
                            dr("M2") = CDbl(drDat("M2")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("M3"), 0)) <> 0 Then
                            dr("M3") = CDbl(drDat("M3")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("M4"), 0)) <> 0 Then
                            dr("M4") = CDbl(drDat("M4")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("M5"), 0)) <> 0 Then
                            dr("M5") = CDbl(drDat("M5")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("M6"), 0)) <> 0 Then
                            dr("M6") = CDbl(drDat("M6")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("M7"), 0)) <> 0 Then
                            dr("M7") = CDbl(drDat("M7")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("M8"), 0)) <> 0 Then
                            dr("M8") = CDbl(drDat("M8")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("M9"), 0)) <> 0 Then
                            dr("M9") = CDbl(drDat("M9")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("M10"), 0)) <> 0 Then
                            dr("M10") = CDbl(drDat("M10")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("M11"), 0)) <> 0 Then
                            dr("M11") = CDbl(drDat("M11")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("M12"), 0)) <> 0 Then
                            dr("M12") = CDbl(drDat("M12")).ToString("#,##0.00")
                        End If

                        If CDec(Nz(drDat("H1"), 0)) <> 0 Then
                            dr("Total1H") = CDbl(drDat("H1")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("H2"), 0)) <> 0 Then
                            dr("Total2H") = CDbl(drDat("H2")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("Y1"), 0)) <> 0 Then
                            dr("TotalY1") = CDbl(drDat("Y1")).ToString("#,##0.00")
                        End If

                        If CDec(Nz(drDat("IMP2"), 0)) <> 0 Then
                            dr("IMP2") = CDbl(drDat("IMP2")).ToString("#,##0.0000")
                        End If

                        If CDec(Nz(drDat("MTP_RRT1"), 0)) <> 0 Then
                            dr("MTP_RRT1") = CDbl(drDat("MTP_RRT1")).ToString("#,##0.00")
                        End If

                    ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then   '// Estimate Budget
                        If CDec(Nz(drDat("M10"), 0)) <> 0 Then
                            dr("M10") = CDbl(drDat("M10")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("M11"), 0)) <> 0 Then
                            dr("M11") = CDbl(drDat("M11")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("M12"), 0)) <> 0 Then
                            dr("M12") = CDbl(drDat("M12")).ToString("#,##0.00")
                        End If

                    ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then   '// Revise Budget
                        If CDec(Nz(drDat("M4"), 0)) <> 0 Then
                            dr("M4") = CDbl(drDat("M4")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("M5"), 0)) <> 0 Then
                            dr("M5") = CDbl(drDat("M5")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("M6"), 0)) <> 0 Then
                            dr("M6") = CDbl(drDat("M6")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("M7"), 0)) <> 0 Then
                            dr("M7") = CDbl(drDat("M7")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("M8"), 0)) <> 0 Then
                            dr("M8") = CDbl(drDat("M8")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("M9"), 0)) <> 0 Then
                            dr("M9") = CDbl(drDat("M9")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("M10"), 0)) <> 0 Then
                            dr("M10") = CDbl(drDat("M10")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("M11"), 0)) <> 0 Then
                            dr("M11") = CDbl(drDat("M11")).ToString("#,##0.00")
                        End If
                        If CDec(Nz(drDat("M12"), 0)) <> 0 Then
                            dr("M12") = CDbl(drDat("M12")).ToString("#,##0.00")
                        End If

                    ElseIf Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then   '// MTP Budget

                        If CDec(Nz(drDat("RevYear"), 0)) <> 0 Then
                            dr("RevYear") = CDbl(drDat("RevYear")).ToString("#,##0.00")
                        End If

                        If CDec(Nz(drDat("PrevRRT1"), 0)) <> 0 Then
                            dr("PrevRRT1") = CDbl(drDat("PrevRRT1")).ToString("#,##0.00")
                        End If

                        If CDec(Nz(drDat("PrevRRT2"), 0)) <> 0 Then
                            dr("PrevRRT2") = CDbl(drDat("PrevRRT2")).ToString("#,##0.00")
                        End If

                        If CDec(Nz(drDat("PrevRRT3"), 0)) <> 0 Then
                            dr("PrevRRT3") = CDbl(drDat("PrevRRT3")).ToString("#,##0.00")
                        End If

                        If CDec(Nz(drDat("PrevRRT4"), 0)) <> 0 Then
                            dr("PrevRRT4") = CDbl(drDat("PrevRRT4")).ToString("#,##0.00")
                        End If

                        If CDec(Nz(drDat("PrevRRT5"), 0)) <> 0 Then
                            dr("PrevRRT5") = CDbl(drDat("PrevRRT5")).ToString("#,##0.00")
                        End If
                    End If

                    If Not IsDBNull(drDat("RRT1")) Then
                        dr("RRT1") = CDbl(drDat("RRT1")).ToString("#,##0.00")
                    End If
                    If Not IsDBNull(drDat("RRT2")) Then
                        dr("RRT2") = CDbl(drDat("RRT2")).ToString("#,##0.00")
                    End If
                    If Not IsDBNull(drDat("RRT3")) Then
                        dr("RRT3") = CDbl(drDat("RRT3")).ToString("#,##0.00")
                    End If
                    If Not IsDBNull(drDat("RRT4")) Then
                        dr("RRT4") = CDbl(drDat("RRT4")).ToString("#,##0.00")
                    End If
                    If Not IsDBNull(drDat("RRT5")) Then
                        dr("RRT5") = CDbl(drDat("RRT5")).ToString("#,##0.00")
                    End If

                    dr("Remarks") = CStr(Nz(drDat("REMARKS")))

                    dtGrid.Rows.Add(dr)
                Next

                '// Load data into datagrid
                If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then   '// Original Budget

                    '// Bind Datasource
                    grvBudget1.AutoGenerateColumns = False
                    grvBudget1.DataSource = dtGrid

                    '// Set Column Headers
                    If Not grvBudget1.Columns("g1col6").HeaderText.Contains("'") Then
                        For i = 6 To 7
                            grvBudget1.Columns("g1col" & CStr(i)).HeaderText += "'" & (CInt(Mid(Me.BudgetKey, 3, 2)) - 1).ToString("00")
                        Next
                        For i = 8 To 16
                            grvBudget1.Columns("g1col" & CStr(i)).HeaderText += "'" & Mid(Me.BudgetKey, 3, 2)
                        Next
                        For i = 1 To 6
                            grvBudget1.Columns("g1colex" & CStr(i)).HeaderText += "'" & Mid(Me.BudgetKey, 3, 2)
                        Next
                        grvBudget1.Columns("g1col17").HeaderText += "'" & (CInt(Mid(Me.BudgetKey, 3, 2)) - 1).ToString("00")
                        grvBudget1.Columns("g1col18").HeaderText += "'" & (CInt(Mid(Me.BudgetKey, 3, 2)) - 1).ToString("00")
                        grvBudget1.Columns("g1col25").HeaderText += "'" & Mid(Me.BudgetKey, 3, 2)
                        grvBudget1.Columns("g1col25").HeaderText = grvBudget1.Columns("g1col25").HeaderText.Replace("@1", (CInt(Mid(Me.BudgetKey, 3, 2)) - 1).ToString("00"))
                        grvBudget1.Columns("g1col26").HeaderText += (CInt(Mid(Me.BudgetKey, 3, 2)) - 1).ToString("00")
                    End If

                    '// Show/Hide WK Column
                    If Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput Then
                        grvBudget1.Columns("g1Wk").Visible = True
                    Else
                        grvBudget1.Columns("g1Wk").Visible = False
                    End If
                    txtWk1.Text = myClsBG0200BL.WorkingBG(1)
                    txtWk2.Text = myClsBG0200BL.WorkingBG(2)

                    '// Show Record Num
                    lblRecNum.Text = grvBudget1.Rows.Count.ToString("#,##0") & " Record(s)"

                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then   '// Estimate Budget
                    '// Bind Datasource
                    grvBudget2.DataSource = dtGrid

                    '// Set Column Headers
                    If Not grvBudget2.Columns("g2col6").HeaderText.Contains("'") Then
                        For i = 6 To 16
                            grvBudget2.Columns("g2col" & CStr(i)).HeaderText += "'" & Mid(Me.BudgetKey, 3, 2)
                        Next
                    End If

                    '// Hide MTP Budget
                    For i = 17 To 21
                        grvBudget2.Columns("g2col" & CStr(i)).Visible = False
                    Next

                    '// Show/Hide WK Column
                    If Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput Then
                        grvBudget2.Columns("g2Wk").Visible = True
                    Else
                        grvBudget2.Columns("g2Wk").Visible = False
                    End If
                    txtWk1.Text = myClsBG0200BL.WorkingBG(1)
                    txtWk2.Text = myClsBG0200BL.WorkingBG(2)

                    '// Show Record Num
                    lblRecNum.Text = grvBudget2.Rows.Count.ToString("#,##0") & " Record(s)"

                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then   '// Revise Budget
                    '// Bind Datasource
                    grvBudget3.DataSource = dtGrid

                    '// Set Column Headers
                    If Not grvBudget3.Columns("g3col6").HeaderText.Contains("'") Then
                        For i = 6 To 25
                            grvBudget3.Columns("g3col" & CStr(i)).HeaderText += "'" & Mid(Me.BudgetKey, 3, 2)
                        Next
                    End If

                    For i = 26 To 30
                        grvBudget3.Columns("g3col" & CStr(i)).Visible = False
                    Next

                    '// Show/Hide WK Column
                    If Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput Then
                        grvBudget3.Columns("g3Wk").Visible = True
                    Else
                        grvBudget3.Columns("g3Wk").Visible = False
                    End If
                    txtWk1.Text = myClsBG0200BL.WorkingBG(1)
                    txtWk2.Text = myClsBG0200BL.WorkingBG(2)

                    '// Show Record Num
                    lblRecNum.Text = grvBudget3.Rows.Count.ToString("#,##0") & " Record(s)"

                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then   '// MTP Budget

                    '// Bind Datasource
                    grvBudget4.AutoGenerateColumns = False
                    grvBudget4.DataSource = dtGrid

                    '// Set Column Headers
                    If Not grvBudget4.Columns("g4col6").HeaderText.Contains("'") Then
                        For i = 6 To 8
                            'grvBudget4.Columns("g4col" & CStr(i)).HeaderText += "'" & Mid(Me.BudgetKey, 3, 2)

                            grvBudget4.Columns("g4col" & CStr(i)).HeaderText += "'" & CInt(Mid(Me.BudgetKey, 3, 2)) + 1

                            'grvBudget4.Columns("g4col" & CStr(i)).HeaderText = grvBudget4.Columns("g4col" & CStr(i)).HeaderText.Replace("@1", (CInt(Mid(Me.BudgetKey, 3, 2)) - 1).ToString("00"))

                            grvBudget4.Columns("g4col" & CStr(i)).HeaderText = grvBudget4.Columns("g4col" & CStr(i)).HeaderText.Replace("@1", (CInt(Mid(Me.BudgetKey, 3, 2)) - 1).ToString("00"))

                        Next

                        If CStr(Me.GetBudgetType()) = P_BUDGET_TYPE_EXPENSE Then
                            For i = 1 To 5
                                '// Current Year
                                'grvBudget4.Columns("g4col" & CStr(9 + ((i - 1) * 2))).HeaderText += "'" & (CInt(Mid(Me.BudgetKey, 3, 2)) + (i)).ToString("00")
                                grvBudget4.Columns("g4col" & CStr(9 + ((i - 1) * 2))).HeaderText += "'" & (CInt(Mid(Me.BudgetKey, 3, 2)) + (i + 1)).ToString("00")
                                grvBudget4.Columns("g4col" & CStr(9 + ((i - 1) * 2))).HeaderText = grvBudget4.Columns("g4col" & CStr(9 + ((i - 1) * 2))).HeaderText.Replace("@1", CInt(Mid(Me.BudgetKey, 3, 2)).ToString("00"))
                            Next

                            'grvBudget4.Columns("g4colDiff1").HeaderText += "'" & (CInt(Mid(Me.BudgetKey, 3, 2)) + (1)).ToString("00")
                            'grvBudget4.Columns("g4colDiff2").HeaderText += "'" & (CInt(Mid(Me.BudgetKey, 3, 2)) + (2)).ToString("00")

                            grvBudget4.Columns("g4colDiff1").HeaderText += "'" & (CInt(Mid(Me.BudgetKey, 3, 2)) + (2)).ToString("00")
                            grvBudget4.Columns("g4colDiff2").HeaderText += "'" & (CInt(Mid(Me.BudgetKey, 3, 2)) + (3)).ToString("00")

                            For i = 1 To 4
                                '// Previous Year
                                'grvBudget4.Columns("g4col" & CStr(10 + ((i - 1) * 2))).HeaderText += "'" & (CInt(Mid(Me.BudgetKey, 3, 2)) + (i)).ToString("00")
                                grvBudget4.Columns("g4col" & CStr(10 + ((i - 1) * 2))).HeaderText += "'" & (CInt(Mid(Me.BudgetKey, 3, 2)) + (i + 1)).ToString("00")
                                grvBudget4.Columns("g4col" & CStr(10 + ((i - 1) * 2))).HeaderText = grvBudget4.Columns("g4col" & CStr(10 + ((i - 1) * 2))).HeaderText.Replace("@1", (CInt(Mid(Me.BudgetKey, 3, 2)) - 1).ToString("00"))
                            Next
                        End If
                    End If

                    If CStr(Me.GetBudgetType()) = P_BUDGET_TYPE_EXPENSE Then
                        '// Show MTP Budget
                        For i = 9 To 12
                            grvBudget4.Columns("g4col" & CStr(i)).Visible = True
                        Next

                        '// Show MTP Budget CAL
                        'For i = 1 To 2
                        '    grvBudget4.Columns("g4ex0" & CStr(i)).Visible = True
                        'Next

                        '// Show RRT Header
                        txtRRT0.Text = CDbl(myClsBG0200BL.RRT(0)).ToString("#,##0")
                        txtRRT1.Text = CDbl(myClsBG0200BL.RRT(1)).ToString("#,##0")
                        txtRRT2.Text = CDbl(myClsBG0200BL.RRT(2)).ToString("#,##0")
                        txtRRT3.Text = CDbl(myClsBG0200BL.RRT(3)).ToString("#,##0")
                        txtRRT4.Text = CDbl(myClsBG0200BL.RRT(4)).ToString("#,##0")
                        txtRRT5.Text = CDbl(myClsBG0200BL.RRT(5)).ToString("#,##0")

                    Else
                        '// Hide MTP Budget
                        For i = 9 To 12
                            grvBudget4.Columns("g4col" & CStr(i)).Visible = False

                        Next
                        '// Hide MTP Budget CAL
                        For i = 1 To 2
                            grvBudget4.Columns("g4ex0" & CStr(i)).Visible = False
                        Next
                    End If

                    '// Show/Hide WK Column
                    If Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput Then
                        grvBudget4.Columns("g4Wk").Visible = True
                    Else
                        grvBudget4.Columns("g4Wk").Visible = False
                    End If
                    txtWk1.Text = myClsBG0200BL.WorkingBG(1)
                    txtWk2.Text = myClsBG0200BL.WorkingBG(2)

                    '// Show Record Num
                    lblRecNum.Text = grvBudget4.Rows.Count.ToString("#,##0") & " Record(s)"

                End If
            End If

            '// Set Read Only
            If Me.OperationCd = enumOperationCd.InputBudget Or (Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput) Then
                grvBudget1.ReadOnly = False
                grvBudget2.ReadOnly = False
                grvBudget3.ReadOnly = False
                grvBudget4.ReadOnly = False
            Else
                grvBudget1.ReadOnly = True
                grvBudget2.ReadOnly = True
                grvBudget3.ReadOnly = True
                grvBudget4.ReadOnly = True
            End If

            '// Set Filter Comboboxs
            SetFilterCombo()

            '// Show Reference Revision No
            lblRefRevNo.Text = myClsBG0200BL.RefRevNo

            '// Show Update History
            If IsDate(myClsBG0200BL.UpdateDate) Then
                lblUpdateDate.Text = "Last Updated: " & CDate(myClsBG0200BL.UpdateDate).ToString("yyyy/MM/dd HH:mm")
            Else
                lblUpdateDate.Text = ""
            End If
            If myClsBG0200BL.UpdateUser.Length > 0 Then
                lblUpdateUser.Text = "By " & myClsBG0200BL.UpdateUser
            Else
                lblUpdateUser.Text = ""
            End If

            '// Set Labels
            lblBudgetPeriod.Text = Mid(Me.BudgetKey, 1, 4)
            If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then      '// Original Budget
                lblBudgetPeriod.Text += " Original Budget"

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then  '// Estimate Budget
                lblBudgetPeriod.Text += " Estimate Budget"

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then    '// Revise Budget
                lblBudgetPeriod.Text += " Revise Budget"

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then    '// MTP Budget
                lblBudgetPeriod.Text += " MTP Budget"
            End If

            If Me.GetBudgetType() = P_BUDGET_TYPE_ASSET Then  '// Investment Budget
                lblBudgetPeriod.Text += " (Investment)"
            End If

            If Me.GetProjectNo <> "1" Then
                lblBudgetPeriod.Text += " " + Me.GetProjectNo
            End If


            lblRevNo.Text = myClsBG0200BL.RevNo

            If myClsBG0200BL.Status = CStr(enumBudgetStatus.Submit) Then
                lblStatus.Text = "Summitted"

            ElseIf myClsBG0200BL.Status = CStr(enumBudgetStatus.Approve) Then
                lblStatus.Text = "Approved"

            ElseIf myClsBG0200BL.Status = CStr(enumBudgetStatus.Adjust) Then
                lblStatus.Text = "Adjusted"

            ElseIf myClsBG0200BL.Status = CStr(enumBudgetStatus.Authorize1) Then
                lblStatus.Text = "Authorize 1"

            ElseIf myClsBG0200BL.Status = CStr(enumBudgetStatus.Authorize2) Then
                lblStatus.Text = "Authorize 2"

            Else
                lblStatus.Text = "New Record"
            End If

            lblRRT0.Text = "Y20" & Mid(Me.BudgetKey, 3, 2) & ":"
            lblRRT1.Text = "Y20" & (CInt(Mid(Me.BudgetKey, 3, 2)) + 1).ToString("00") & ":"
            lblRRT2.Text = "Y20" & (CInt(Mid(Me.BudgetKey, 3, 2)) + 2).ToString("00") & ":"
            lblRRT3.Text = "Y20" & (CInt(Mid(Me.BudgetKey, 3, 2)) + 3).ToString("00") & ":"
            lblRRT4.Text = "Y20" & (CInt(Mid(Me.BudgetKey, 3, 2)) + 4).ToString("00") & ":"
            lblRRT5.Text = "Y20" & (CInt(Mid(Me.BudgetKey, 3, 2)) + 5).ToString("00") & ":"

            Debug.Print(Now.ToString() & ": End ShowDatagrid")

            ShowDatagrid = True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub SetFilterCombo()
        Try
            Debug.Print(Now.ToString() & ": Begin SetFilterCombo")

            '// Save full data
            If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then      '// Original Budget
                '// Save full data
                myDtAllData = CType(grvBudget1.DataSource, DataTable)

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then  '// Estimate Budget
                '// Save full data
                myDtAllData = CType(grvBudget2.DataSource, DataTable)

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then    '// Revise Budget
                '// Save full data
                myDtAllData = CType(grvBudget3.DataSource, DataTable)

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then    '// MTP Budget
                '// Save full data
                myDtAllData = CType(grvBudget4.DataSource, DataTable)
            End If

            Dim dt As DataTable = Nothing
            Dim dr As DataRow = Nothing

            '// Set Account Filter Comboboxs
            dt = SelectDistinct(myDtAllData, New String() {"Account"})
            dr = dt.NewRow
            dr("Account") = "All"
            dt.Rows.InsertAt(dr, 0)
            cboAccount.DisplayMember = "Account"
            cboAccount.ValueMember = "Account"
            cboAccount.DataSource = dt

            '// Set Dept Filter Comboboxs
            dt = SelectDistinct(myDtAllData, New String() {"Dept"})
            dr = dt.NewRow
            dr("Dept") = "All"
            dt.Rows.InsertAt(dr, 0)
            cboDept.DisplayMember = "Dept"
            cboDept.ValueMember = "Dept"
            cboDept.DataSource = dt

            If Me.GetBudgetType() = P_BUDGET_TYPE_EXPENSE Then
                '// Set CostType Filter Comboboxs
                dt = SelectDistinct(myDtAllData, New String() {"CostType"})
                dr = dt.NewRow
                dr("CostType") = "All"
                dt.Rows.InsertAt(dr, 0)
                cboCostType.DisplayMember = "CostType"
                cboCostType.ValueMember = "CostType"
                cboCostType.DataSource = dt

                '// Set Cost Filter Comboboxs
                dt = SelectDistinct(myDtAllData, New String() {"Cost"})
                dr = dt.NewRow
                dr("Cost") = "All"
                dt.Rows.InsertAt(dr, 0)
                cboCost.DisplayMember = "Cost"
                cboCost.ValueMember = "Cost"
                cboCost.DataSource = dt
            Else
                '// Set CostType Filter Comboboxs
                cboCostType.Items.Clear()
                cboCostType.Items.Add("All")
                cboCostType.SelectedIndex = 0

                '// Set Cost Filter Comboboxs
                cboCost.Items.Clear()
                cboCost.Items.Add("All")
                cboCost.SelectedIndex = 0
            End If

            Debug.Print(Now.ToString() & ": End SetFilterCombo")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub FilterGridView()
        Dim grvTemp As DataGridView = Nothing
        Dim strNamePrefix As String = String.Empty

        Dim dtTemp As DataTable = Nothing
        Dim dtTemp1 As DataTable = Nothing

        Try
            If myDataLoadingFlg = True Then
                Exit Sub
            End If

            Debug.Print(Now.ToString() & ": Begin FilterGridView")
            Me.Cursor = Cursors.WaitCursor

            If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then      '// Original Budget
                If mydtBG1 Is Nothing Then
                    mydtBG1 = CType(grvBudget1.DataSource, DataTable).Copy
                    mydtBG1View = CType(grvBudget1.DataSource, DataTable).DefaultView.Table.Copy
                End If

                dtTemp = mydtBG1
                dtTemp1 = mydtBG1View
                grvTemp = grvBudget1
                strNamePrefix = "g1col"

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then  '// Estimate Budget

                If mydtBG2 Is Nothing Then
                    mydtBG2 = CType(grvBudget2.DataSource, DataTable).Copy
                    mydtBG2View = CType(grvBudget2.DataSource, DataTable).DefaultView.Table.Copy
                End If
                dtTemp = mydtBG2
                dtTemp1 = mydtBG2View
                grvTemp = grvBudget2
                strNamePrefix = "g2col"

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then    '// Revise Budget

                If mydtBG3 Is Nothing Then
                    mydtBG3 = CType(grvBudget3.DataSource, DataTable).Copy
                    mydtBG3View = CType(grvBudget3.DataSource, DataTable).DefaultView.Table.Copy
                End If
                dtTemp = mydtBG3
                dtTemp1 = mydtBG3View
                grvTemp = grvBudget3
                strNamePrefix = "g3col"

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then    '// MTP Budget

                If mydtBG4 Is Nothing Then
                    Dim dtGrvBudget4 As New DataTable
                    dtGrvBudget4 = CType(grvBudget4.DataSource, DataTable)
                    mydtBG4 = dtGrvBudget4.Clone
                    Dim drNew As DataRow
                    For Each dr As DataRow In dtGrvBudget4.Rows
                        drNew = mydtBG4.NewRow
                        drNew.ItemArray = dr.ItemArray
                        mydtBG4.Rows.Add(drNew)
                    Next
                    mydtBG4View = CType(grvBudget4.DataSource, DataTable).DefaultView.Table.Copy
                End If

                dtTemp = mydtBG4
                dtTemp1 = mydtBG4View
                grvTemp = grvBudget4
                strNamePrefix = "g4col"
            End If

            For Each row As DataRow In dtTemp.Rows
                For intcolumn As Integer = 0 To row.ItemArray.Length - 1
                    If row(intcolumn).ToString().Equals("0.00") Then
                        row(intcolumn) = Nothing
                    End If
                Next
            Next

            Dim dvTemp As DataView = New DataView(dtTemp)
            Dim strFilterAccount As String = ""
            Dim strFilterCost As String = ""
            Dim strFilterCostType As String = ""
            Dim strFilterDepartment As String = ""
            Dim strFilter As String = ""

            dvTemp.RowFilter = Nothing

            '// Show Record Num
            lblRecNum.Text = ""

            Dim intCounter As Integer = 0

            If grvTemp IsNot Nothing Then
                If cboAccount.SelectedIndex = 0 And cboCostType.SelectedIndex = 0 And _
                cboCost.SelectedIndex = 0 And cboDept.SelectedIndex = 0 Then
                    dvTemp.RowFilter = Nothing
                    grvTemp.DataSource = dvTemp.ToTable
                Else
                    If cboAccount.SelectedIndex > 0 Then
                        strFilterAccount = "Account ='" & cboAccount.Text & "'"
                    End If

                    If Me.GetBudgetType() = P_BUDGET_TYPE_EXPENSE Then
                        If cboCostType.SelectedIndex > 0 Then
                            strFilterCostType = "CostType ='" & cboCostType.Text & "'"
                        End If

                        If cboCost.SelectedIndex > 0 Then
                            strFilterCost = "Cost ='" & cboCost.Text & "'"
                        End If
                    End If

                    If cboDept.SelectedIndex > 0 Then
                        strFilterDepartment = "Dept = '" & cboDept.Text & "'"
                    End If

                    If strFilterAccount <> "" Then
                        strFilter = strFilter & " AND " & strFilterAccount
                    End If

                    If strFilterCostType <> "" Then
                        strFilter = strFilter & " AND " & strFilterCostType
                    End If

                    If strFilterCost <> "" Then
                        strFilter = strFilter & " AND " & strFilterCost
                    End If

                    If strFilterDepartment <> "" Then
                        strFilter = strFilter & " AND " & strFilterDepartment
                    End If

                    If strFilter <> "" AndAlso strFilter.StartsWith(" AND ") Then
                        strFilter = strFilter.Substring(4, strFilter.Length - 4)
                    End If

                    dvTemp.RowFilter = Nothing
                    dvTemp.RowFilter = strFilter
                    grvTemp.DataSource = dvTemp.ToTable
                End If

                If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then
                    Dim returnvalue As Object

                    lblSum1.Text = "Total 1st Half'" & Me.GetBudgetYear()
                    lblSum2.Text = "Actual 1st Half'" & CStr(CInt(Me.GetBudgetYear()) - 1)
                    lblSum3.Text = "Total 2nd Half'" & Me.GetBudgetYear()
                    lblSum4.Text = "Estimate 2nd Half'" & CStr(CInt(Me.GetBudgetYear()) - 1)
                    lblSum5.Text = "Diff 1st Half"
                    lblSum6.Text = "Diff 2st Half"
                    lblSum7.Text = "Total Year'" & Me.GetBudgetYear()
                    lblSum9.Text = "MTP" & (CInt(Mid(Me.BudgetKey, 3, 2)) - 1).ToString("00") & "Y" & Me.GetBudgetYear()
                    lblSum10.Text = "Diff vs MTP" & (CInt(Mid(Me.BudgetKey, 3, 2)) - 1).ToString("00")
                    lblSum11.Text = "Total Year'" & CStr(CInt(Me.GetBudgetYear()) - 1)
                    lblSum12.Text = "Diff vs Year'" & CStr(CInt(Me.GetBudgetYear()) - 1)

                    lblSum1Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
                    lblSum3Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
                    lblSum7Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
                    lblSum8.Visible = False
                    lblSum8Val.Visible = False

                    returnvalue = dtTemp1.Compute("sum(Total1H)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSum1Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSum1Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(IMP1)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSum2Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSum2Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(Total2H)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSum3Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSum3Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(IMP2)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSum4Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSum4Val.Text = "0.00"
                    End If

                    If lblSum1Val.Text.ToString <> "" And lblSum2Val.Text.ToString <> "" Then
                        lblSum5Val.Text = (CDbl(lblSum1Val.Text) - CDbl(lblSum2Val.Text)).ToString("#,##0.00")
                    Else
                        lblSum5Val.Text = "0.00"
                    End If

                    If lblSum3Val.Text.ToString <> "" And lblSum4Val.Text.ToString <> "" Then
                        lblSum6Val.Text = (CDbl(lblSum3Val.Text) - CDbl(lblSum4Val.Text)).ToString("#,##0.00")
                    Else
                        lblSum6Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(TotalY1)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSum7Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSum7Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(MTP_RRT1)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSum9Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSum9Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(DiffMTP_RRT1)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSum10Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSum10Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(TotalY2)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSum11Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSum11Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(Diff)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSum12Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSum12Val.Text = "0.00"
                    End If

                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then
                    Dim returnvalue As Object

                    lblSum1.Text = "Estimate 2nd Half'" & Me.GetBudgetYear()
                    lblSum2.Text = "Revise 2nd Half'" & Me.GetBudgetYear()
                    lblSum3.Text = "Actual 1st Half'" & Me.GetBudgetYear()
                    lblSum4.Text = "Diff 2nd Half'" & Me.GetBudgetYear()
                    lblSum5.Text = "Estimate Total Year'" & Me.GetBudgetYear()
                    lblSum6.Text = ""
                    lblSum7.Text = ""
                    lblSum8.Text = ""
                    lblSum9.Text = ""

                    lblSum1Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
                    lblSum5Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)

                    returnvalue = dtTemp1.Compute("sum(Est2H)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSum1Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSum1Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(IMP2)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSum2Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSum2Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(IMP1)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSum3Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSum3Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(Diff2H)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSum4Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSum4Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(EstTotal)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSum5Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSum5Val.Text = "0.00"
                    End If

                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then
                    Dim returnvalue As Object

                    lblSum1.Text = "Estimate 1st Half'" & Me.GetBudgetYear()
                    lblSum2.Text = "Original 1st Half'" & Me.GetBudgetYear()
                    lblSum3.Text = "Revise 2nd Half'" & Me.GetBudgetYear()
                    lblSum4.Text = "Original 2nd Half'" & Me.GetBudgetYear()
                    lblSum5.Text = "Diff 1st Half'" & Me.GetBudgetYear()
                    lblSum6.Text = "Diff 2nd Half'" & Me.GetBudgetYear()
                    lblSum7.Text = "Revise Year'" & Me.GetBudgetYear()
                    lblSum8.Text = "Original Year'" & Me.GetBudgetYear()
                    lblSum9.Text = "Diff Year'" & Me.GetBudgetYear()

                    lblSum1Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
                    lblSum3Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
                    lblSum7Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
                    lblSum9Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)

                    returnvalue = dtTemp1.Compute("sum(Est1H)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSum1Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSum1Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(IMP1)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSum2Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSum2Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(Rev2H)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSum3Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSum3Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(IMP2)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSum4Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSum4Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(Diff1H)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSum5Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSum5Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(Diff2H)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSum6Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSum6Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(RevYear)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSum7Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSum7Val.Text = "0.00"
                    End If

                    If lblSum2Val.Text.ToString <> "" And lblSum4Val.Text.ToString <> "" Then
                        lblSum8Val.Text = (CDbl(lblSum2Val.Text) + CDbl(lblSum4Val.Text)).ToString("#,##0.00")
                    Else
                        lblSum8Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(DiffYear)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSum9Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSum9Val.Text = "0.00"
                    End If

                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then
                    Dim returnvalue As Object

                    lblSumMTP1.Text = "MTP" & Me.GetBudgetYear()
                    lblSumMTP2.Text = "MTP" & CStr(CInt(Me.GetBudgetYear()) - 1)
                    lblSumMTP3.Text = "RB Y" & Me.GetBudgetYear()
                    lblSumMTP4.Text = "Y" & CStr(CInt(Me.GetBudgetYear()) + 1)
                    lblSumMTP5.Text = "Y" & CStr(CInt(Me.GetBudgetYear()) + 2)
                    lblSumMTP6.Text = "Y" & CStr(CInt(Me.GetBudgetYear()) + 3)
                    lblSumMTP7.Text = "Y" & CStr(CInt(Me.GetBudgetYear()) + 4)
                    lblSumMTP8.Text = "Y" & CStr(CInt(Me.GetBudgetYear()) + 5)

                    lblSumMTP3Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
                    lblSumMTP7Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
                    lblSumMTP5Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
                    lblSumMTP9Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
                    lblSumMTP11Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)

                    returnvalue = dtTemp1.Compute("sum(RevYear)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSumMTP1Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSumMTP1Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(PrevRRT1)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSumMTP2Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSumMTP2Val.Text = "0.00"
                    End If

                    dtTemp1.Columns.Add("temp", GetType(Double), "Convert(RRT1, 'System.Double')")
                    returnvalue = dtTemp1.Compute("sum(temp)", strFilter)
                    dtTemp1.Columns.Remove("temp")

                    If returnvalue.ToString <> "" Then
                        lblSumMTP3Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSumMTP3Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(PrevRRT2)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSumMTP4Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSumMTP4Val.Text = "0.00"
                    End If

                    dtTemp1.Columns.Add("temp", GetType(Double), "Convert(RRT2, 'System.Double')")
                    returnvalue = dtTemp1.Compute("sum(temp)", strFilter)
                    dtTemp1.Columns.Remove("temp")

                    If returnvalue.ToString <> "" Then
                        lblSumMTP5Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSumMTP5Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(PrevRRT3)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSumMTP6Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSumMTP6Val.Text = "0.00"
                    End If

                    dtTemp1.Columns.Add("temp", GetType(Double), "Convert(RRT3, 'System.Double')")
                    returnvalue = dtTemp1.Compute("sum(temp)", strFilter)
                    dtTemp1.Columns.Remove("temp")

                    If returnvalue.ToString <> "" Then
                        lblSumMTP7Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSumMTP7Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(PrevRRT4)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSumMTP8Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSumMTP8Val.Text = "0.00"
                    End If

                    dtTemp1.Columns.Add("temp", GetType(Double), "Convert(RRT4, 'System.Double')")
                    returnvalue = dtTemp1.Compute("sum(temp)", strFilter)
                    dtTemp1.Columns.Remove("temp")

                    If returnvalue.ToString <> "" Then
                        lblSumMTP9Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSumMTP9Val.Text = "0.00"
                    End If

                    returnvalue = dtTemp1.Compute("sum(PrevRRT5)", strFilter)
                    If returnvalue.ToString <> "" Then
                        lblSumMTP10Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSumMTP10Val.Text = "0.00"
                    End If

                    dtTemp1.Columns.Add("temp", GetType(Double), "Convert(RRT5, 'System.Double')")
                    returnvalue = dtTemp1.Compute("sum(temp)", strFilter)
                    dtTemp1.Columns.Remove("temp")

                    If returnvalue.ToString <> "" Then
                        lblSumMTP11Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                    Else
                        lblSumMTP11Val.Text = "0.00"
                    End If
                End If

                lblRecNum.Text = dvTemp.Count.ToString("#,##0") & " Record(s)"
            End If

            Me.Cursor = Cursors.Default

            Debug.Print(Now.ToString() & ": End FilterGridView")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub CalSum()
        Dim strNamePrefix As String = String.Empty
        Dim strFilterAccount As String = ""
        Dim strFilterCost As String = ""
        Dim strFilterCostType As String = ""
        Dim strFilterDepartment As String = ""
        Dim strFilter As String = ""

        Try
            If myDataLoadingFlg = True Then
                Exit Sub
            End If

            Debug.Print(Now.ToString() & ": Begin CalSum")

            Me.Cursor = Cursors.WaitCursor

            '// Show Record Num
            lblRecNum.Text = ""

            Dim intCounter As Integer = 0

            If cboAccount.SelectedIndex = 0 And cboCostType.SelectedIndex = 0 And _
            cboCost.SelectedIndex = 0 And cboDept.SelectedIndex = 0 Then

            Else
                If cboAccount.SelectedIndex > 0 Then
                    strFilterAccount = "Account ='" & cboAccount.Text & "'"
                End If

                If Me.GetBudgetType() = P_BUDGET_TYPE_EXPENSE Then
                    If cboCostType.SelectedIndex > 0 Then
                        strFilterCostType = "CostType ='" & cboCostType.Text & "'"
                    End If

                    If cboCost.SelectedIndex > 0 Then
                        strFilterCost = "Cost ='" & cboCost.Text & "'"
                    End If
                End If

                If cboDept.SelectedIndex > 0 Then
                    strFilterDepartment = "Dept = '" & cboDept.Text & "'"
                End If

                If strFilterAccount <> "" Then
                    strFilter = strFilter & " AND " & strFilterAccount
                End If

                If strFilterCostType <> "" Then
                    strFilter = strFilter & " AND " & strFilterCostType
                End If

                If strFilterCost <> "" Then
                    strFilter = strFilter & " AND " & strFilterCost
                End If

                If strFilterDepartment <> "" Then
                    strFilter = strFilter & " AND " & strFilterDepartment
                End If

                If strFilter <> "" AndAlso strFilter.StartsWith(" AND ") Then
                    strFilter = strFilter.Substring(4, strFilter.Length - 4)
                End If
            End If

            If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then
                Dim returnvalue As Object

                lblSum1.Text = "Total 1st Half'" & Me.GetBudgetYear()
                lblSum2.Text = "Actual 1st Half'" & CStr(CInt(Me.GetBudgetYear()) - 1)
                lblSum3.Text = "Total 2nd Half'" & Me.GetBudgetYear()
                lblSum4.Text = "Estimate 2nd Half'" & CStr(CInt(Me.GetBudgetYear()) - 1)
                lblSum5.Text = "Diff 1st Half"
                lblSum6.Text = "Diff 2st Half"
                lblSum7.Text = "Total Year'" & Me.GetBudgetYear()
                lblSum8.Text = "Total Year'" & CStr(CInt(Me.GetBudgetYear()) - 1)
                lblSum9.Text = "Diff"

                lblSum1Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
                lblSum3Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
                lblSum7Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
                lblSum9Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)

                returnvalue = CType(grvBudget1.DataSource, DataTable).DefaultView.Table.Compute("sum(Total1H)", strFilter)
                If returnvalue.ToString <> "" Then
                    lblSum1Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSum1Val.Text = "0.00"
                End If

                returnvalue = CType(grvBudget1.DataSource, DataTable).DefaultView.Table.Compute("sum(IMP1)", strFilter)
                If returnvalue.ToString <> "" Then
                    lblSum2Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSum2Val.Text = "0.00"
                End If

                returnvalue = CType(grvBudget1.DataSource, DataTable).DefaultView.Table.Compute("sum(Total2H)", strFilter)
                If returnvalue.ToString <> "" Then
                    lblSum3Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSum3Val.Text = "0.00"
                End If

                returnvalue = CType(grvBudget1.DataSource, DataTable).DefaultView.Table.Compute("sum(IMP2)", strFilter)
                If returnvalue.ToString <> "" Then
                    lblSum4Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSum4Val.Text = "0.00"
                End If

                If lblSum1Val.Text.ToString <> "" And lblSum2Val.Text.ToString <> "" Then
                    lblSum5Val.Text = (CDbl(lblSum1Val.Text) - CDbl(lblSum2Val.Text)).ToString("#,##0.00")
                Else
                    lblSum5Val.Text = "0.00"
                End If

                If lblSum3Val.Text.ToString <> "" And lblSum4Val.Text.ToString <> "" Then
                    lblSum6Val.Text = (CDbl(lblSum3Val.Text) - CDbl(lblSum4Val.Text)).ToString("#,##0.00")
                Else
                    lblSum6Val.Text = "0.00"
                End If

                returnvalue = CType(grvBudget1.DataSource, DataTable).DefaultView.Table.Compute("sum(TotalY1)", strFilter)
                If returnvalue.ToString <> "" Then
                    lblSum7Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSum7Val.Text = "0.00"
                End If

                returnvalue = CType(grvBudget1.DataSource, DataTable).DefaultView.Table.Compute("sum(TotalY2)", strFilter)
                If returnvalue.ToString <> "" Then
                    lblSum8Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSum8Val.Text = "0.00"
                End If

                returnvalue = CType(grvBudget1.DataSource, DataTable).DefaultView.Table.Compute("sum(Diff)", strFilter)
                If returnvalue.ToString <> "" Then
                    lblSum9Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSum9Val.Text = "0.00"
                End If

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then
                Dim returnvalue As Object

                lblSum1.Text = "Estimate 2nd Half'" & Me.GetBudgetYear()
                lblSum2.Text = "Revise 2nd Half'" & Me.GetBudgetYear()
                lblSum3.Text = "Actual 1st Half'" & Me.GetBudgetYear()
                lblSum4.Text = "Diff 2nd Half'" & Me.GetBudgetYear()
                lblSum5.Text = "Estimate Total Year'" & Me.GetBudgetYear()
                lblSum6.Text = ""
                lblSum7.Text = ""
                lblSum8.Text = ""
                lblSum9.Text = ""

                lblSum1Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
                lblSum5Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)

                returnvalue = CType(grvBudget2.DataSource, DataTable).DefaultView.Table.Compute("sum(Est2H)", strFilter)
                If returnvalue.ToString <> "" Then
                    lblSum1Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSum1Val.Text = "0.00"
                End If

                returnvalue = CType(grvBudget2.DataSource, DataTable).DefaultView.Table.Compute("sum(IMP2)", strFilter)
                If returnvalue.ToString <> "" Then
                    lblSum2Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSum2Val.Text = "0.00"
                End If

                returnvalue = CType(grvBudget2.DataSource, DataTable).DefaultView.Table.Compute("sum(IMP1)", strFilter)
                If returnvalue.ToString <> "" Then
                    lblSum3Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSum3Val.Text = "0.00"
                End If

                returnvalue = CType(grvBudget2.DataSource, DataTable).DefaultView.Table.Compute("sum(Diff2H)", strFilter)
                If returnvalue.ToString <> "" Then
                    lblSum4Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSum4Val.Text = "0.00"
                End If

                returnvalue = CType(grvBudget2.DataSource, DataTable).DefaultView.Table.Compute("sum(EstTotal)", strFilter)
                If returnvalue.ToString <> "" Then
                    lblSum5Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSum5Val.Text = "0.00"
                End If

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then
                Dim returnvalue As Object

                lblSum1.Text = "Estimate 1st Half'" & Me.GetBudgetYear()
                lblSum2.Text = "Original 1st Half'" & Me.GetBudgetYear()
                lblSum3.Text = "Revise 2nd Half'" & Me.GetBudgetYear()
                lblSum4.Text = "Original 2nd Half'" & Me.GetBudgetYear()
                lblSum5.Text = "Diff 1st Half'" & Me.GetBudgetYear()
                lblSum6.Text = "Diff 2nd Half'" & Me.GetBudgetYear()
                lblSum7.Text = "Revise Year'" & Me.GetBudgetYear()
                lblSum8.Text = "Original Year'" & Me.GetBudgetYear()
                lblSum9.Text = "Diff Year'" & Me.GetBudgetYear()

                lblSum1Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
                lblSum3Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
                lblSum7Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
                lblSum9Val.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)

                returnvalue = CType(grvBudget3.DataSource, DataTable).DefaultView.Table.Compute("sum(Est1H)", strFilter)
                If returnvalue.ToString <> "" Then
                    lblSum1Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSum1Val.Text = "0.00"
                End If

                returnvalue = CType(grvBudget3.DataSource, DataTable).DefaultView.Table.Compute("sum(IMP1)", strFilter)
                If returnvalue.ToString <> "" Then
                    lblSum2Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSum2Val.Text = "0.00"
                End If

                returnvalue = CType(grvBudget3.DataSource, DataTable).DefaultView.Table.Compute("sum(Rev2H)", strFilter)
                If returnvalue.ToString <> "" Then
                    lblSum3Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSum3Val.Text = "0.00"
                End If

                returnvalue = CType(grvBudget3.DataSource, DataTable).DefaultView.Table.Compute("sum(IMP2)", strFilter)
                If returnvalue.ToString <> "" Then
                    lblSum4Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSum4Val.Text = "0.00"
                End If

                returnvalue = CType(grvBudget3.DataSource, DataTable).DefaultView.Table.Compute("sum(Diff1H)", strFilter)
                If returnvalue.ToString <> "" Then
                    lblSum5Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSum5Val.Text = "0.00"
                End If

                returnvalue = CType(grvBudget3.DataSource, DataTable).DefaultView.Table.Compute("sum(Diff2H)", strFilter)
                If returnvalue.ToString <> "" Then
                    lblSum6Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSum6Val.Text = "0.00"
                End If

                returnvalue = CType(grvBudget3.DataSource, DataTable).DefaultView.Table.Compute("sum(RevYear)", strFilter)
                If returnvalue.ToString <> "" Then
                    lblSum7Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSum7Val.Text = "0.00"
                End If

                If lblSum2Val.Text.ToString <> "" And lblSum4Val.Text.ToString <> "" Then
                    lblSum8Val.Text = (CDbl(lblSum2Val.Text) + CDbl(lblSum4Val.Text)).ToString("#,##0.00")
                Else
                    lblSum8Val.Text = "0.00"
                End If

                returnvalue = CType(grvBudget3.DataSource, DataTable).DefaultView.Table.Compute("sum(DiffYear)", strFilter)
                If returnvalue.ToString <> "" Then
                    lblSum9Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSum9Val.Text = "0.00"
                End If

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then
                Dim returnvalue As Object
                Dim dtTemp1 As DataTable = CType(grvBudget4.DataSource, DataTable)

                dtTemp1.Columns.Add("temp", GetType(Double), "Convert(RRT1, 'System.Double')")
                returnvalue = dtTemp1.Compute("sum(temp)", strFilter)
                dtTemp1.Columns.Remove("temp")

                If returnvalue.ToString <> "" Then
                    lblSumMTP3Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSumMTP3Val.Text = "0.00"
                End If

                dtTemp1.Columns.Add("temp", GetType(Double), "Convert(RRT2, 'System.Double')")
                returnvalue = dtTemp1.Compute("sum(temp)", strFilter)
                dtTemp1.Columns.Remove("temp")

                If returnvalue.ToString <> "" Then
                    lblSumMTP5Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSumMTP5Val.Text = "0.00"
                End If

                dtTemp1.Columns.Add("temp", GetType(Double), "Convert(RRT3, 'System.Double')")
                returnvalue = dtTemp1.Compute("sum(temp)", strFilter)
                dtTemp1.Columns.Remove("temp")

                If returnvalue.ToString <> "" Then
                    lblSumMTP7Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSumMTP7Val.Text = "0.00"
                End If

                dtTemp1.Columns.Add("temp", GetType(Double), "Convert(RRT4, 'System.Double')")
                returnvalue = dtTemp1.Compute("sum(temp)", strFilter)
                dtTemp1.Columns.Remove("temp")

                If returnvalue.ToString <> "" Then
                    lblSumMTP9Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSumMTP9Val.Text = "0.00"
                End If

                dtTemp1.Columns.Add("temp", GetType(Double), "Convert(RRT5, 'System.Double')")
                returnvalue = dtTemp1.Compute("sum(temp)", strFilter)
                dtTemp1.Columns.Remove("temp")

                If returnvalue.ToString <> "" Then
                    lblSumMTP11Val.Text = CDbl(returnvalue).ToString("#,##0.00")
                Else
                    lblSumMTP11Val.Text = "0.00"
                End If

                If Me.GetBudgetType() = P_BUDGET_TYPE_EXPENSE Then
                    HighlightMTPValueAllNew()
                End If
            End If

            Me.Cursor = Cursors.Default

            Debug.Print(Now.ToString() & ": End CalSum")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub SetButtons()
        Try
            Debug.Print(Now.ToString() & ": Begin SetButtons")

            '// Enable buttons
            cmdSave.Enabled = True
            If Me.OperationCd = enumOperationCd.InputBudget Then
                If myClsBG0200BL.IsSubmitUser = True Then
                    cmdSubmit.Enabled = True
                Else
                    cmdSubmit.Enabled = False
                End If
            End If

            cmdApprove.Enabled = True
            cmdReject.Enabled = True
            cmdAuth1.Enabled = True
            cmdAuth2.Enabled = True
            cmdUpRev.Enabled = True
            cmdDelRev.Enabled = True
            cmdAdjust.Enabled = True
            cmdSubmit2.Enabled = True
            cmdReInput.Enabled = True

            Debug.Print(Now.ToString() & ": End SetButtons")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub LoadUploadData()
        Try
            If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then   '// Original Budget
                myClsBG0200BL.BudgetYear = CStr(CInt(Me.GetBudgetYear()) - 1)
            Else
                myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
            End If

            myClsBG0200BL.PeriodType = Me.GetPeriodType()
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()
            myClsBG0200BL.GetUploadData()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub LoadPICList()
        Try
            '// get Rev No.
            GetCurrentRevNo()

            '// load PIC Combo
            LoadComboPIC()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub LoadComboPIC()
        Try
            '// Show Person In Charge List
            myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
            myClsBG0200BL.PeriodType = Me.GetPeriodType()
            myClsBG0200BL.BudgetType = Me.GetBudgetType()
            myClsBG0200BL.RevNo = Me.CurrRevNo
            myClsBG0200BL.OperationCd = Me.OperationCd
            myClsBG0200BL.UserPIC = p_strUserPIC
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            myClsBG0200BL.GetPersonInChargeList()

            If myClsBG0200BL.PicList.Rows.Count > 0 Then
                If Me.OperationCd = enumOperationCd.AdjustBudget Or _
                Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput Or _
                Me.OperationCd = enumOperationCd.Authorize1 Or _
                Me.OperationCd = enumOperationCd.Authorize2 Then

                    '// Add Combobox's member
                    Dim dt As DataTable = myClsBG0200BL.PicList
                    Dim dr As DataRow = dt.NewRow
                    dr("PIC_NAME") = "All"
                    dr("PERSON_IN_CHARGE_NO") = "0000"
                    dt.Rows.InsertAt(dr, 0)

                    cboPIC.DisplayMember = "PIC_NAME"
                    cboPIC.ValueMember = "PERSON_IN_CHARGE_NO"
                    cboPIC.DataSource = dt

                ElseIf Me.OperationCd = enumOperationCd.ApproveBudget Then
                    '// Add Combobox's member
                    cboPIC.DisplayMember = "PIC_NAME"
                    cboPIC.ValueMember = "PERSON_IN_CHARGE_NO"
                    cboPIC.DataSource = myClsBG0200BL.PicList

                    '// Set Default Selection
                    If cboPIC.Items.Count > 0 Then
                        If cboPIC.FindString(p_strUserPIC) >= 0 Then
                            cboPIC.SelectedIndex = cboPIC.FindString(p_strUserPIC)
                            cboPIC.Text = cboPIC.SelectedText
                        End If
                    End If

                Else
                    '// Add Combobox's member
                    If p_strUserPIC = "0000" Or _
                    (Me.OperationCd = enumOperationCd.ViewBudget And (p_strUserPIC = "BTMT10" Or p_strUserPIC = "BTMT3")) Then
                        cboPIC.DisplayMember = "PIC_NAME"
                        cboPIC.ValueMember = "PERSON_IN_CHARGE_NO"
                        cboPIC.DataSource = myClsBG0200BL.PicList
                    Else
                        If myClsBG0200BL.GetChildPicList() = True Then
                            Dim strChildList As String = ""
                            For Each dr As DataRow In myClsBG0200BL.ChildPicList.Rows
                                If strChildList = "" Then
                                    strChildList = "'" & CStr(dr("PIC_CHILD_NO")) & "'"
                                Else
                                    strChildList = strChildList & ",'" & CStr(dr("PIC_CHILD_NO")) & "'"
                                End If
                            Next
                            myClsBG0200BL.PicList.DefaultView.RowFilter = "PERSON_IN_CHARGE_NO LIKE '" & p_strUserPIC & "%' OR " & _
                                                                          "PERSON_IN_CHARGE_NO IN (" & strChildList & ")"
                        Else
                            myClsBG0200BL.PicList.DefaultView.RowFilter = "PERSON_IN_CHARGE_NO LIKE '" & p_strUserPIC & "%'"
                        End If
                        cboPIC.DisplayMember = "PIC_NAME"
                        cboPIC.ValueMember = "PERSON_IN_CHARGE_NO"
                        cboPIC.DataSource = myClsBG0200BL.PicList.DefaultView.Table

                        '// Set Default Selection
                        If cboPIC.Items.Count > 0 Then
                            If cboPIC.FindString(p_strUserPIC) >= 0 Then
                                cboPIC.SelectedIndex = cboPIC.FindString(p_strUserPIC)
                                cboPIC.Text = cboPIC.SelectedText
                            End If
                        End If
                    End If
                End If
            Else
                cboPIC.DataSource = myClsBG0200BL.PicList
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub ReloadPICList(ByVal RevNo As String)
        Try
            '// set Current Rev No.
            Me.CurrRevNo = RevNo

            '// load PIC Combo
            LoadComboPIC()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub ShowWKH()
        Try
            myClsBG0200BL.BudgetYear = CStr(Me.GetBudgetYear())
            myClsBG0200BL.PeriodType = CStr(Me.GetPeriodType())
            myClsBG0200BL.RevNo = Me.CurrRevNo
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            If myClsBG0200BL.GetWKH() = True Then
                Me.txtWKH1.Text = CDbl(myClsBG0200BL.WKH1).ToString("#,##0.00")
                Me.txtWKH2.Text = CDbl(myClsBG0200BL.WKH2).ToString("#,##0.00")

                Me.txtWKRRT1.Text = CDbl(myClsBG0200BL.WKRRT1).ToString("#,##0.00")
                Me.txtWKRRT2.Text = CDbl(myClsBG0200BL.WKRRT2).ToString("#,##0.00")
                Me.txtWKRRT3.Text = CDbl(myClsBG0200BL.WKRRT3).ToString("#,##0.00")
                Me.txtWKRRT4.Text = CDbl(myClsBG0200BL.WKRRT4).ToString("#,##0.00")
                Me.txtWKRRT5.Text = CDbl(myClsBG0200BL.WKRRT5).ToString("#,##0.00")

                Me.txtMTPWB.Text = CDbl(myClsBG0200BL.MTPWB).ToString("#,##0.00")
            Else
                Me.txtWKH1.Text = ""
                Me.txtWKH2.Text = ""

                Me.txtWKRRT1.Text = ""
                Me.txtWKRRT2.Text = ""
                Me.txtWKRRT3.Text = ""
                Me.txtWKRRT4.Text = ""
                Me.txtWKRRT5.Text = ""

                Me.txtMTPWB.Text = ""
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub ShowMTP_SUM()
        Try
            myClsBG0200BL.BudgetYear = CStr(Me.GetBudgetYear())
            myClsBG0200BL.PeriodType = CStr(Me.GetPeriodType())
            myClsBG0200BL.RevNo = Me.CurrRevNo
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            If myClsBG0200BL.GetMTP_SUM() = True Then
                Me.txtMTP_SUM1.Text = CDbl(myClsBG0200BL.MTP_SUM1).ToString("#,##0.00")
                Me.txtMTP_SUM2.Text = CDbl(myClsBG0200BL.MTP_SUM2).ToString("#,##0.00")
                Me.txtMTP_SUM3.Text = CDbl(myClsBG0200BL.MTP_SUM3).ToString("#,##0.00")
                Me.txtMTP_SUM4.Text = CDbl(myClsBG0200BL.MTP_SUM4).ToString("#,##0.00")
                Me.txtMTP_SUM5.Text = CDbl(myClsBG0200BL.MTP_SUM5).ToString("#,##0.00")
            Else
                Me.txtMTP_SUM1.Text = ""
                Me.txtMTP_SUM2.Text = ""
                Me.txtMTP_SUM3.Text = ""
                Me.txtMTP_SUM4.Text = ""
                Me.txtMTP_SUM5.Text = ""
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub ShowMTPInvestment()
        Try
            myClsBG0200BL.BudgetYear = CStr(Me.GetBudgetYear())
            myClsBG0200BL.PeriodType = CStr(Me.GetPeriodType())
            myClsBG0200BL.RevNo = Me.CurrRevNo
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            If myClsBG0200BL.GetMTPInvestment() = True Then
                Me.txtMTPInv1.Text = CDbl(myClsBG0200BL.MTP_SUM1).ToString("#,##0.00")
                Me.txtMTPInv2.Text = CDbl(myClsBG0200BL.MTP_SUM2).ToString("#,##0.00")
                Me.txtMTPInv3.Text = CDbl(myClsBG0200BL.MTP_SUM3).ToString("#,##0.00")
                Me.txtMTPInv4.Text = CDbl(myClsBG0200BL.MTP_SUM4).ToString("#,##0.00")
                Me.txtMTPInv5.Text = CDbl(myClsBG0200BL.MTP_SUM5).ToString("#,##0.00")

                Me.txtPYInv1.Text = CDbl(myClsBG0200BL.MTP_PY_SUM1).ToString("#,##0.00")
                Me.txtPYInv2.Text = CDbl(myClsBG0200BL.MTP_PY_SUM2).ToString("#,##0.00")
                Me.txtPYInv3.Text = CDbl(myClsBG0200BL.MTP_PY_SUM3).ToString("#,##0.00")
                Me.txtPYInv4.Text = CDbl(myClsBG0200BL.MTP_PY_SUM4).ToString("#,##0.00")
                Me.txtPYInv5.Text = CDbl(myClsBG0200BL.MTP_PY_SUM5).ToString("#,##0.00")
            Else
                Me.txtMTPInv1.Text = ""
                Me.txtMTPInv2.Text = ""
                Me.txtMTPInv3.Text = ""
                Me.txtMTPInv4.Text = ""
                Me.txtMTPInv5.Text = ""

                Me.txtPYInv1.Text = ""
                Me.txtPYInv2.Text = ""
                Me.txtPYInv3.Text = ""
                Me.txtPYInv4.Text = ""
                Me.txtPYInv5.Text = ""
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub LoadTransferList()
        Try
            myClsBG0200BL.BudgetYear = CStr(Me.GetBudgetYear())
            myClsBG0200BL.PeriodType = CStr(Me.GetPeriodType())
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            myClsBG0200BL.GetTransferCost()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub ShowUploadData()
        Dim strOrderNo As String
        Dim drActualData As DataRow()
        Dim drBudgetData As DataRow()
        Dim drBudgetList As DataRow()
        Dim dtGridData As DataTable
        Dim blnGridChanged As Boolean = False

        Try
            myDataLoadingFlg = True

            Me.Cursor = Cursors.WaitCursor

            '// Show upload data if exists
            If myClsBG0200BL.UpDataList IsNot Nothing AndAlso myClsBG0200BL.UpDataList.Rows.Count > 0 Then
                Debug.Print(Now.ToString() & ": Begin ShowUploadData " & myClsBG0200BL.UpDataList.Rows.Count & " record(s)")

                If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then '// Original Budget
                    dtGridData = CType(grvBudget1.DataSource, DataTable).DefaultView.Table.Copy

                    If dtGridData IsNot Nothing Then
                        For i = 0 To dtGridData.Rows.Count - 1
                            strOrderNo = CStr(Nz(dtGridData.Rows(i).Item("OrderNo")))

                            drActualData = myClsBG0200BL.UpDataList.Select("DATA_TYPE = 2 AND BUDGET_ORDER_NO = '" & strOrderNo & "'")
                            drBudgetData = myClsBG0200BL.UpDataList.Select("DATA_TYPE = 1 AND BUDGET_ORDER_NO = '" & strOrderNo & "'")

                            '// Actual 1st Half
                            If drActualData.Count > 0 Then
                                If CDbl(Nz(drActualData(0).Item("H1"), 0)) = 0 Then
                                    dtGridData.Rows(i).Item("IMP1") = DBNull.Value
                                Else
                                    dtGridData.Rows(i).Item("IMP1") = drActualData(0).Item("H1")
                                End If

                                blnGridChanged = True
                            End If
                        Next
                    End If

                    '// Bind grid with changed data
                    If blnGridChanged = True Then
                        grvBudget1.DataSource = dtGridData
                    End If

                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then '// Estimate Budget
                    dtGridData = CType(grvBudget2.DataSource, DataTable).DefaultView.Table.Copy

                    If dtGridData IsNot Nothing Then
                        For i = 0 To dtGridData.Rows.Count - 1
                            strOrderNo = CStr(Nz(dtGridData.Rows(i).Item("OrderNo")))

                            drActualData = myClsBG0200BL.UpDataList.Select("DATA_TYPE = 2 AND BUDGET_ORDER_NO = '" & strOrderNo & "'")
                            drBudgetData = myClsBG0200BL.UpDataList.Select("DATA_TYPE = 1 AND BUDGET_ORDER_NO = '" & strOrderNo & "'")

                            '// Actual 1st Half
                            If drActualData.Count > 0 Then
                                If CDbl(Nz(drActualData(0).Item("H1"), 0)) = 0 Then
                                    dtGridData.Rows(i).Item("IMP1") = DBNull.Value
                                Else
                                    dtGridData.Rows(i).Item("IMP1") = drActualData(0).Item("H1")
                                End If

                                blnGridChanged = True
                            End If

                            '// Revise 2nd Half
                            If drBudgetData.Count > 0 Then

                                If CDbl(Nz(drBudgetData(0).Item("H2"), 0)) = 0 Then
                                    dtGridData.Rows(i).Item("IMP2") = DBNull.Value
                                Else
                                    dtGridData.Rows(i).Item("IMP2") = drBudgetData(0).Item("H2")
                                End If

                                blnGridChanged = True
                            End If

                            '// Actual Jul - Sep
                            If drActualData.Count > 0 Then
                                For j = 8 To 10
                                    If CDbl(Nz(drActualData(0).Item("M" & CStr(j - 1)), 0)) = 0 Then
                                        dtGridData.Rows(i).Item("M" & CStr(j - 1)) = DBNull.Value
                                    Else
                                        dtGridData.Rows(i).Item("M" & CStr(j - 1)) = drActualData(0).Item("M" & CStr(j - 1))
                                    End If
                                Next
                                blnGridChanged = True
                            End If

                            '// Estimate Oct - Dec
                            If drBudgetData.Count > 0 Then
                                drBudgetList = myClsBG0200BL.BudgetList.Select("BUDGET_ORDER_NO = '" & strOrderNo & "'")

                                For j = 11 To 13
                                    If IsDBNull(drBudgetList(0).Item("M" & CStr(j - 1))) Then

                                        If CDbl(Nz(drBudgetData(0).Item("M" & CStr(j - 1)), 0)) = 0 Then
                                            dtGridData.Rows(i).Item("M" & CStr(j - 1)) = DBNull.Value
                                        Else
                                            dtGridData.Rows(i).Item("M" & CStr(j - 1)) = drBudgetData(0).Item("M" & CStr(j - 1))
                                        End If

                                        blnGridChanged = True
                                    End If
                                Next
                            End If
                        Next
                    End If

                    '// Bind grid with changed data
                    If blnGridChanged = True Then
                        grvBudget2.DataSource = dtGridData
                    End If

                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then    '// Revise Budget
                    dtGridData = CType(grvBudget3.DataSource, DataTable).DefaultView.Table.Copy

                    If dtGridData IsNot Nothing Then
                        For i = 0 To dtGridData.Rows.Count - 1
                            strOrderNo = CStr(Nz(dtGridData.Rows(i).Item("OrderNo")))

                            drBudgetData = myClsBG0200BL.UpDataList.Select("DATA_TYPE = 1 AND BUDGET_ORDER_NO = '" & strOrderNo & "'")
                            drActualData = myClsBG0200BL.UpDataList.Select("DATA_TYPE = 2 AND BUDGET_ORDER_NO = '" & strOrderNo & "'")

                            '// Original 1st Half
                            If drBudgetData.Count > 0 Then
                                If CDbl(Nz(drBudgetData(0).Item("H1"), 0)) = 0 Then
                                    dtGridData.Rows(i).Item("IMP1") = DBNull.Value
                                Else
                                    dtGridData.Rows(i).Item("IMP1") = drBudgetData(0).Item("H1")
                                End If

                                blnGridChanged = True
                            End If

                            '// Actual Jan - Mar
                            If drActualData.Count > 0 Then
                                For j = 7 To 9
                                    If CDbl(Nz(drActualData(0).Item("M" & CStr(j - 6)), 0)) = 0 Then
                                        dtGridData.Rows(i).Item("M" & CStr(j - 6)) = DBNull.Value
                                    Else
                                        dtGridData.Rows(i).Item("M" & CStr(j - 6)) = drActualData(0).Item("M" & CStr(j - 6))
                                    End If
                                Next
                                blnGridChanged = True
                            End If

                            '// Estimate Apr - Jun
                            If drBudgetData.Count > 0 Then
                                drBudgetList = myClsBG0200BL.BudgetList.Select("BUDGET_ORDER_NO = '" & strOrderNo & "'")

                                For j = 10 To 12
                                    If IsDBNull(drBudgetList(0).Item("M" & CStr(j - 6))) Then

                                        If CDbl(Nz(drBudgetList(0).Item("M" & CStr(j - 6)), 0)) = 0 Then
                                            dtGridData.Rows(i).Item("M" & CStr(j - 6)) = DBNull.Value
                                        Else
                                            dtGridData.Rows(i).Item("M" & CStr(j - 6)) = drBudgetList(0).Item("M" & CStr(j - 6))
                                        End If

                                        blnGridChanged = True
                                    End If
                                Next
                            End If

                            '// Original 2nd Half
                            If drBudgetData.Count > 0 Then
                                If CDbl(Nz(drBudgetData(0).Item("H2"), 0)) = 0 Then
                                    dtGridData.Rows(i).Item("IMP2") = DBNull.Value
                                Else
                                    dtGridData.Rows(i).Item("IMP2") = drBudgetData(0).Item("H2")
                                End If

                                blnGridChanged = True
                            End If

                            '// Revise Jul - Dec
                            If drBudgetData.Count > 0 Then
                                drBudgetList = myClsBG0200BL.BudgetList.Select("BUDGET_ORDER_NO = '" & strOrderNo & "'")

                                For j = 16 To 21
                                    If IsDBNull(drBudgetList(0).Item("M" & CStr(j - 9))) Then

                                        If CDbl(Nz(drBudgetList(0).Item("M" & CStr(j - 9)), 0)) = 0 Then
                                            dtGridData.Rows(i).Item("M" & CStr(j - 9)) = DBNull.Value
                                        Else
                                            dtGridData.Rows(i).Item("M" & CStr(j - 9)) = drBudgetList(0).Item("M" & CStr(j - 9))
                                        End If

                                        blnGridChanged = True
                                    End If
                                Next
                            End If
                        Next
                    End If

                    '// Bind grid with changed data
                    If blnGridChanged = True Then
                        grvBudget3.DataSource = dtGridData
                    End If

                End If
            End If

            Me.Cursor = Cursors.Default

            myDataLoadingFlg = False

            Debug.Print(Now.ToString() & ": End ShowUploadData")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub CheckValidateOriginalBudget(ByVal intColumnIndex As Integer, ByVal intRowIndex As Integer)
        Try
            If myDataLoadingFlg = True Then
                Exit Sub
            End If

            '// Check Validation
            If ((CStr(Me.GetBudgetType()) = P_BUDGET_TYPE_EXPENSE And _
                (intColumnIndex = grvBudget1.Columns("g1col6").Index Or _
                intColumnIndex = grvBudget1.Columns("g1col7").Index Or _
                intColumnIndex = grvBudget1.Columns("g1col8").Index Or _
                intColumnIndex = grvBudget1.Columns("g1col9").Index Or _
                intColumnIndex = grvBudget1.Columns("g1col10").Index Or _
                intColumnIndex = grvBudget1.Columns("g1col11").Index Or _
                intColumnIndex = grvBudget1.Columns("g1col12").Index Or _
                intColumnIndex = grvBudget1.Columns("g1col13").Index Or _
                intColumnIndex = grvBudget1.Columns("g1colex1").Index Or _
                intColumnIndex = grvBudget1.Columns("g1colex2").Index Or _
                intColumnIndex = grvBudget1.Columns("g1colex3").Index Or _
                intColumnIndex = grvBudget1.Columns("g1colex4").Index Or _
                intColumnIndex = grvBudget1.Columns("g1colex5").Index Or _
                intColumnIndex = grvBudget1.Columns("g1colex6").Index))) Or _
                ((CStr(Me.GetBudgetType()) = P_BUDGET_TYPE_ASSET And _
                (intColumnIndex = grvBudget1.Columns("g1col6").Index Or _
                intColumnIndex = grvBudget1.Columns("g1col7").Index Or _
                intColumnIndex = grvBudget1.Columns("g1col8").Index Or _
                intColumnIndex = grvBudget1.Columns("g1col9").Index Or _
                intColumnIndex = grvBudget1.Columns("g1col10").Index Or _
                intColumnIndex = grvBudget1.Columns("g1col11").Index Or _
                intColumnIndex = grvBudget1.Columns("g1col12").Index Or _
                intColumnIndex = grvBudget1.Columns("g1col13").Index Or _
                intColumnIndex = grvBudget1.Columns("g1colex1").Index Or _
                intColumnIndex = grvBudget1.Columns("g1colex2").Index Or _
                intColumnIndex = grvBudget1.Columns("g1colex3").Index Or _
                intColumnIndex = grvBudget1.Columns("g1colex4").Index Or _
                intColumnIndex = grvBudget1.Columns("g1colex5").Index Or _
                intColumnIndex = grvBudget1.Columns("g1colex6").Index))) Then

                Dim objVal As Object = grvBudget1.Item(intColumnIndex, intRowIndex).Value
                If IsDBNull(objVal) Or objVal Is Nothing Then
                    Exit Sub
                End If

                '// Compute Expression
                If Not IsNumeric(objVal) Then
                    If CStr(objVal).Contains("+") Or CStr(objVal).Contains("-") Or _
                    CStr(objVal).Contains("*") Or CStr(objVal).Contains("/") Then
                        objVal = Equate(CStr(objVal))
                    End If
                End If

                '// Set Data Format
                mySetGridValue = True
                If Not IsNumeric(objVal) OrElse CDbl(objVal) = 0 Then
                    grvBudget1.Item(intColumnIndex, intRowIndex).Value = Nothing
                Else
                    grvBudget1.Item(intColumnIndex, intRowIndex).Value = CDbl(objVal).ToString("#,##0.00")
                End If
                mySetGridValue = False

                '// Auto Calculation
                CalcOriginalBudget(intRowIndex)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub CheckValidateEstimateBudget(ByVal intColumnIndex As Integer, ByVal intRowIndex As Integer)
        Try
            If myDataLoadingFlg = True Then
                Exit Sub
            End If

            '// Check Validation
            If intColumnIndex = grvBudget2.Columns("g2col6").Index Or _
                intColumnIndex = grvBudget2.Columns("g2col7").Index Or _
                intColumnIndex = grvBudget2.Columns("g2col8").Index Or _
                intColumnIndex = grvBudget2.Columns("g2col9").Index Or _
                intColumnIndex = grvBudget2.Columns("g2col10").Index Or _
                intColumnIndex = grvBudget2.Columns("g2col11").Index Or _
                intColumnIndex = grvBudget2.Columns("g2col12").Index Or _
                intColumnIndex = grvBudget2.Columns("g2col13").Index Then

                Dim objVal As Object = grvBudget2.Item(intColumnIndex, intRowIndex).Value
                If IsDBNull(objVal) Or objVal Is Nothing Then
                    Exit Sub
                End If

                '// Compute Expression
                If Not IsNumeric(objVal) Then
                    If CStr(objVal).Contains("+") Or CStr(objVal).Contains("-") Or _
                    CStr(objVal).Contains("*") Or CStr(objVal).Contains("/") Then
                        objVal = Equate(CStr(objVal))
                    End If
                End If

                '// Set Data Format
                mySetGridValue = True
                If Not IsNumeric(objVal) OrElse CDbl(objVal) = 0 Then
                    grvBudget2.Item(intColumnIndex, intRowIndex).Value = Nothing
                Else
                    grvBudget2.Item(intColumnIndex, intRowIndex).Value = CDbl(objVal).ToString("#,##0.00")
                End If
                mySetGridValue = False

                '// Auto Calculation
                CalcEstimateBudget(intRowIndex)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub CheckValidateReviseBudget(ByVal intColumnIndex As Integer, ByVal intRowIndex As Integer)
        Try
            If myDataLoadingFlg = True Then
                Exit Sub
            End If

            '// Check Validation
            If intColumnIndex = grvBudget3.Columns("g3col6").Index Or _
                intColumnIndex = grvBudget3.Columns("g3col7").Index Or _
                intColumnIndex = grvBudget3.Columns("g3col8").Index Or _
                intColumnIndex = grvBudget3.Columns("g3col9").Index Or _
                intColumnIndex = grvBudget3.Columns("g3col10").Index Or _
                intColumnIndex = grvBudget3.Columns("g3col11").Index Or _
                intColumnIndex = grvBudget3.Columns("g3col12").Index Or _
                intColumnIndex = grvBudget3.Columns("g3col15").Index Or _
                intColumnIndex = grvBudget3.Columns("g3col16").Index Or _
                intColumnIndex = grvBudget3.Columns("g3col17").Index Or _
                intColumnIndex = grvBudget3.Columns("g3col18").Index Or _
                intColumnIndex = grvBudget3.Columns("g3col19").Index Or _
                intColumnIndex = grvBudget3.Columns("g3col20").Index Or _
                intColumnIndex = grvBudget3.Columns("g3col21").Index Or _
                intColumnIndex = grvBudget3.Columns("g3col26").Index Or _
                intColumnIndex = grvBudget3.Columns("g3col27").Index Or _
                intColumnIndex = grvBudget3.Columns("g3col28").Index Or _
                intColumnIndex = grvBudget3.Columns("g3col29").Index Or _
                intColumnIndex = grvBudget3.Columns("g3col30").Index Then

                Dim objVal As Object = grvBudget3.Item(intColumnIndex, intRowIndex).Value
                If IsDBNull(objVal) Or objVal Is Nothing Then
                    Exit Sub
                End If

                '// Compute Expression
                If Not IsNumeric(objVal) Then
                    If CStr(objVal).Contains("+") Or CStr(objVal).Contains("-") Or _
                    CStr(objVal).Contains("*") Or CStr(objVal).Contains("/") Then
                        objVal = Equate(CStr(objVal))
                    End If
                End If

                '// Set Data Format
                mySetGridValue = True
                If Not IsNumeric(objVal) OrElse CDbl(objVal) = 0 Then
                    grvBudget3.Item(intColumnIndex, intRowIndex).Value = Nothing
                Else
                    grvBudget3.Item(intColumnIndex, intRowIndex).Value = CDbl(objVal).ToString("#,##0.00")
                End If
                mySetGridValue = False

                If intColumnIndex = grvBudget3.Columns("g3col6").Index Or _
                  intColumnIndex = grvBudget3.Columns("g3col7").Index Or _
                  intColumnIndex = grvBudget3.Columns("g3col8").Index Or _
                  intColumnIndex = grvBudget3.Columns("g3col9").Index Or _
                  intColumnIndex = grvBudget3.Columns("g3col10").Index Or _
                  intColumnIndex = grvBudget3.Columns("g3col11").Index Or _
                  intColumnIndex = grvBudget3.Columns("g3col12").Index Or _
                  intColumnIndex = grvBudget3.Columns("g3col15").Index Or _
                  intColumnIndex = grvBudget3.Columns("g3col16").Index Or _
                  intColumnIndex = grvBudget3.Columns("g3col17").Index Or _
                  intColumnIndex = grvBudget3.Columns("g3col18").Index Or _
                  intColumnIndex = grvBudget3.Columns("g3col19").Index Or _
                  intColumnIndex = grvBudget3.Columns("g3col20").Index Or _
                  intColumnIndex = grvBudget3.Columns("g3col21").Index Then
                    '// Auto Calculation
                    CalcReviseBudget(intRowIndex)
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub CheckValidateMTPBudget(ByVal intColumnIndex As Integer, ByVal intRowIndex As Integer)
        Try
            If myDataLoadingFlg = True Then
                Exit Sub
            End If

            '// Check Validation
            If intColumnIndex = grvBudget4.Columns("g4col9").Index Or _
                intColumnIndex = grvBudget4.Columns("g4col11").Index Or _
                intColumnIndex = grvBudget4.Columns("g4col13").Index Or _
                intColumnIndex = grvBudget4.Columns("g4col15").Index Or _
                intColumnIndex = grvBudget4.Columns("g4col17").Index Or _
                intColumnIndex = grvBudget4.Columns("g4ex01").Index Or _
                intColumnIndex = grvBudget4.Columns("g4ex02").Index Or _
                intColumnIndex = grvBudget4.Columns("g4ex03").Index Or _
                intColumnIndex = grvBudget4.Columns("g4ex04").Index Or _
                intColumnIndex = grvBudget4.Columns("g4ex05").Index Then

                Dim objVal As Object = grvBudget4.Item(intColumnIndex, intRowIndex).Value
                If IsDBNull(objVal) Or objVal Is Nothing Then
                    Exit Sub
                End If

                '// Compute Expression
                If Not IsNumeric(objVal) Then
                    If CStr(objVal).Contains("+") Or CStr(objVal).Contains("-") Or _
                    CStr(objVal).Contains("*") Or CStr(objVal).Contains("/") Then
                        objVal = Equate(CStr(objVal))
                    End If
                End If

                '// Set Data Format
                mySetGridValue = True
                If Not IsNumeric(objVal) OrElse CDbl(objVal) = 0 Then
                    grvBudget4.Item(intColumnIndex, intRowIndex).Value = Nothing
                Else
                    grvBudget4.Item(intColumnIndex, intRowIndex).Value = CDbl(objVal).ToString("#,##0.00")
                End If
                mySetGridValue = False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub CalcOriginalBudget()
        Dim dblTotal As Double

        Try
            Debug.Print(Now.ToString & ": Begin CalcOriginalBudget")

            Me.Cursor = Cursors.WaitCursor

            Dim dtDat As DataTable = CType(grvBudget1.DataSource, DataTable).DefaultView.Table.Copy

            Dim dblSumIMP1 As Double = 0D
            Dim dblSumIMP2 As Double = 0D
            '// Auto Calculation
            For intRowIndex As Integer = 0 To dtDat.Rows.Count - 1
                '// Calc Total last Year
                dblTotal = CDbl(Nz(dtDat.Rows(intRowIndex)![IMP1], 0)) + CDbl(Nz(dtDat.Rows(intRowIndex)![IMP2], 0))
                If dblTotal = 0.0 Then
                    dtDat.Rows(intRowIndex)![TotalY2] = DBNull.Value
                Else
                    dtDat.Rows(intRowIndex)![TotalY2] = dblTotal
                End If

                '// Calc Difference
                dblTotal = CDbl(Nz(dtDat.Rows(intRowIndex)![TotalY1], 0)) - CDbl(Nz(dtDat.Rows(intRowIndex)![TotalY2], 0))
                If dblTotal = 0.0 Then
                    dtDat.Rows(intRowIndex)![Diff] = DBNull.Value
                Else
                    dtDat.Rows(intRowIndex)![Diff] = dblTotal
                End If

                '// Calc Difference MTP_RRT1
                dblTotal = CDbl(Nz(dtDat.Rows(intRowIndex)![TotalY1], 0)) - CDbl(Nz(dtDat.Rows(intRowIndex)![MTP_RRT1], 0))
                If dblTotal = 0.0 Then
                    dtDat.Rows(intRowIndex)![DiffMTP_RRT1] = DBNull.Value
                Else
                    dtDat.Rows(intRowIndex)![DiffMTP_RRT1] = dblTotal
                End If
            Next

            grvBudget1.DataSource = dtDat
            FilterGridView()

            Me.Cursor = Cursors.Default
            Debug.Print(Now.ToString & ": End CalcOriginalBudget")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub CalcOriginalBudget(ByVal intRowIndex As Integer)
        Dim dblTotal As Double
        Dim dtDat As DataTable

        Try
            Debug.Print(Now.ToString & ": Begin CalcOriginalBudget2")
            Me.Cursor = Cursors.WaitCursor
            dtDat = CType(grvBudget1.DataSource, DataTable)

            '// Auto Calculation
            '// Calc Total 1st Half
            dblTotal = CDbl(Nz(grvBudget1.Item("g1col8", intRowIndex).Value, 0)) + _
                       CDbl(Nz(grvBudget1.Item("g1col9", intRowIndex).Value, 0)) + _
                       CDbl(Nz(grvBudget1.Item("g1col10", intRowIndex).Value, 0)) + _
                       CDbl(Nz(grvBudget1.Item("g1col11", intRowIndex).Value, 0)) + _
                       CDbl(Nz(grvBudget1.Item("g1col12", intRowIndex).Value, 0)) + _
                       CDbl(Nz(grvBudget1.Item("g1col13", intRowIndex).Value, 0))
            If dblTotal = 0 Then
                grvBudget1.Item("g1col14", intRowIndex).Value = DBNull.Value
            Else
                grvBudget1.Item("g1col14", intRowIndex).Value = dblTotal
            End If

            '// Calc Total 2nd Half
            dblTotal = CDbl(Nz(grvBudget1.Item("g1colex1", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget1.Item("g1colex2", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget1.Item("g1colex3", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget1.Item("g1colex4", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget1.Item("g1colex5", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget1.Item("g1colex6", intRowIndex).Value, 0))
            If dblTotal = 0 Then
                grvBudget1.Item("g1col15", intRowIndex).Value = DBNull.Value
            Else
                grvBudget1.Item("g1col15", intRowIndex).Value = dblTotal
            End If

            '// Calc Total Year
            dblTotal = CDbl(Nz(grvBudget1.Item("g1col14", intRowIndex).Value, 0)) + _
                       CDbl(Nz(grvBudget1.Item("g1col15", intRowIndex).Value, 0))
            If dblTotal = 0 Then
                grvBudget1.Item("g1col16", intRowIndex).Value = DBNull.Value
            Else
                grvBudget1.Item("g1col16", intRowIndex).Value = dblTotal
            End If

            '// Calc Total last Year
            dblTotal = CDbl(Nz(grvBudget1.Item("g1col6", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget1.Item("g1col7", intRowIndex).Value, 0))
            If dblTotal = 0 Then
                grvBudget1.Item("g1col17", intRowIndex).Value = DBNull.Value
            Else
                grvBudget1.Item("g1col17", intRowIndex).Value = dblTotal
            End If

            '// Calc Difference
            dblTotal = CDbl(Nz(grvBudget1.Item("g1col16", intRowIndex).Value, 0)) - _
                        CDbl(Nz(grvBudget1.Item("g1col17", intRowIndex).Value, 0))
            If dblTotal = 0 Then
                grvBudget1.Item("g1col18", intRowIndex).Value = DBNull.Value
            Else
                grvBudget1.Item("g1col18", intRowIndex).Value = dblTotal
            End If

            '// Calc Difference MTP_RRT1
            dblTotal = CDbl(Nz(grvBudget1.Item("g1col16", intRowIndex).Value, 0)) - _
                        CDbl(Nz(grvBudget1.Item("g1col25", intRowIndex).Value, 0))
            If dblTotal = 0 Then
                grvBudget1.Item("g1col26", intRowIndex).Value = DBNull.Value
            Else
                grvBudget1.Item("g1col26", intRowIndex).Value = dblTotal
            End If

            Me.Cursor = Cursors.Default
            Debug.Print(Now.ToString & ": End CalcOriginalBudget2")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub SumOriginalBudget(ByVal intRowIndex As Integer)
        Dim dblTotal As Double
        Dim dtDat As DataTable

        Try
            Debug.Print(Now.ToString & ": Begin CalcOriginalBudget2")
            Me.Cursor = Cursors.WaitCursor
            dtDat = CType(grvBudget1.DataSource, DataTable)

            '// Auto Calculation
            '// Calc Total 1st Half
            dblTotal = CDbl(Nz(grvBudget1.Item("g1col8", intRowIndex).Value, 0)) + _
                       CDbl(Nz(grvBudget1.Item("g1col9", intRowIndex).Value, 0)) + _
                       CDbl(Nz(grvBudget1.Item("g1col10", intRowIndex).Value, 0)) + _
                       CDbl(Nz(grvBudget1.Item("g1col11", intRowIndex).Value, 0)) + _
                       CDbl(Nz(grvBudget1.Item("g1col12", intRowIndex).Value, 0)) + _
                       CDbl(Nz(grvBudget1.Item("g1col13", intRowIndex).Value, 0))
            If dblTotal = 0 Then
                grvBudget1.Item("g1col14", intRowIndex).Value = DBNull.Value
            Else
                grvBudget1.Item("g1col14", intRowIndex).Value = dblTotal
            End If

            '// Calc Total 2nd Half
            dblTotal = CDbl(Nz(grvBudget1.Item("g1colex1", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget1.Item("g1colex2", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget1.Item("g1colex3", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget1.Item("g1colex4", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget1.Item("g1colex5", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget1.Item("g1colex6", intRowIndex).Value, 0))
            If dblTotal = 0 Then
                grvBudget1.Item("g1col15", intRowIndex).Value = DBNull.Value
            Else
                grvBudget1.Item("g1col15", intRowIndex).Value = dblTotal
            End If

            '// Calc Total Year
            dblTotal = CDbl(Nz(grvBudget1.Item("g1col14", intRowIndex).Value, 0)) + _
                       CDbl(Nz(grvBudget1.Item("g1col15", intRowIndex).Value, 0))
            If dblTotal = 0 Then
                grvBudget1.Item("g1col16", intRowIndex).Value = DBNull.Value
            Else
                grvBudget1.Item("g1col16", intRowIndex).Value = dblTotal
            End If

            '// Calc Total last Year
            dblTotal = CDbl(Nz(grvBudget1.Item("g1col6", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget1.Item("g1col7", intRowIndex).Value, 0))
            If dblTotal = 0 Then
                grvBudget1.Item("g1col17", intRowIndex).Value = DBNull.Value
            Else
                grvBudget1.Item("g1col17", intRowIndex).Value = dblTotal
            End If

            '// Calc Difference
            dblTotal = CDbl(Nz(grvBudget1.Item("g1col16", intRowIndex).Value, 0)) - _
                        CDbl(Nz(grvBudget1.Item("g1col17", intRowIndex).Value, 0))
            If dblTotal = 0 Then
                grvBudget1.Item("g1col18", intRowIndex).Value = DBNull.Value
            Else
                grvBudget1.Item("g1col18", intRowIndex).Value = dblTotal
            End If

            '// Calc Difference MTP_RRT1
            dblTotal = CDbl(Nz(grvBudget1.Item("g1col16", intRowIndex).Value, 0)) - _
                        CDbl(Nz(grvBudget1.Item("g1col25", intRowIndex).Value, 0))
            If dblTotal = 0 Then
                grvBudget1.Item("g1col26", intRowIndex).Value = DBNull.Value
            Else
                grvBudget1.Item("g1col26", intRowIndex).Value = dblTotal
            End If

            Me.Cursor = Cursors.Default

            Debug.Print(Now.ToString & ": End CalcOriginalBudget2")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub CalcEstimateBudget()
        Dim dblTotal As Double

        Try
            Debug.Print(Now.ToString & ": Begin CalcEstimateBudget")
            Me.Cursor = Cursors.WaitCursor
            Dim dtDat As DataTable = CType(grvBudget2.DataSource, DataTable).DefaultView.Table.Copy

            '// Auto Calculation
            For intRowIndex As Integer = 0 To dtDat.Rows.Count - 1
                '// Calc Estimate 2nd Half
                dblTotal = CDbl(Nz(dtDat.Rows(intRowIndex)![M7], 0)) + _
                            CDbl(Nz(dtDat.Rows(intRowIndex)![M8], 0)) + _
                            CDbl(Nz(dtDat.Rows(intRowIndex)![M9], 0)) + _
                            CDbl(Nz(dtDat.Rows(intRowIndex)![M10], 0)) + _
                            CDbl(Nz(dtDat.Rows(intRowIndex)![M11], 0)) + _
                            CDbl(Nz(dtDat.Rows(intRowIndex)![M12], 0))
                If dblTotal = 0 Then
                    dtDat.Rows(intRowIndex)![Est2H] = DBNull.Value
                Else
                    dtDat.Rows(intRowIndex)![Est2H] = dblTotal
                End If

                '// Calc Diff 2nd Half
                dblTotal = CDbl(Nz(dtDat.Rows(intRowIndex)![Est2H], 0)) - CDbl(Nz(dtDat.Rows(intRowIndex)![IMP2], 0))
                If dblTotal = 0 Then
                    dtDat.Rows(intRowIndex)![Diff2H] = DBNull.Value
                Else
                    dtDat.Rows(intRowIndex)![Diff2H] = dblTotal
                End If

                '// Calc Estimate Total Year
                dblTotal = CDbl(Nz(dtDat.Rows(intRowIndex)![IMP1], 0)) + CDbl(Nz(dtDat.Rows(intRowIndex)![Est2H], 0))
                If dblTotal = 0 Then
                    dtDat.Rows(intRowIndex)![EstTotal] = DBNull.Value
                Else
                    dtDat.Rows(intRowIndex)![EstTotal] = dblTotal
                End If
            Next

            grvBudget2.DataSource = dtDat
            FilterGridView()
            Me.Cursor = Cursors.Default
            Debug.Print(Now.ToString & ": End CalcEstimateBudget")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub CalcEstimateBudget(ByVal intRowIndex As Integer)
        Dim dblTotal As Double

        Try
            Debug.Print(Now.ToString & ": Begin CalcEstimateBudget2")
            Me.Cursor = Cursors.WaitCursor

            '// Auto Calculation
            '// Calc Estimate 2nd Half
            dblTotal = CDbl(Nz(grvBudget2.Item("g2col8", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget2.Item("g2col9", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget2.Item("g2col10", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget2.Item("g2col11", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget2.Item("g2col12", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget2.Item("g2col13", intRowIndex).Value, 0))
            If dblTotal = 0 Then
                grvBudget2.Item("g2col14", intRowIndex).Value = DBNull.Value
            Else
                grvBudget2.Item("g2col14", intRowIndex).Value = dblTotal
            End If

            '// Calc Diff 2nd Half
            dblTotal = CDbl(Nz(grvBudget2.Item("g2col14", intRowIndex).Value, 0)) - _
                        CDbl(Nz(grvBudget2.Item("g2col7", intRowIndex).Value, 0))
            If dblTotal = 0 Then
                grvBudget2.Item("g2col15", intRowIndex).Value = DBNull.Value
            Else
                grvBudget2.Item("g2col15", intRowIndex).Value = dblTotal
            End If

            '// Calc Estimate Total Year
            dblTotal = CDbl(Nz(grvBudget2.Item("g2col6", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget2.Item("g2col14", intRowIndex).Value, 0))
            If dblTotal = 0 Then
                grvBudget2.Item("g2col16", intRowIndex).Value = DBNull.Value
            Else
                grvBudget2.Item("g2col16", intRowIndex).Value = dblTotal
            End If

            Me.Cursor = Cursors.Default
            Debug.Print(Now.ToString & ": Begin CalcEstimateBudget2")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub CalcReviseMTPBudget()
        Dim dblTotal As Double
        Dim dblTotal1 As Double
        Dim dblTotal2 As Double

        Try
            Debug.Print(Now.ToString & ": Begin CalcReviseMTPBudget")
            Me.Cursor = Cursors.WaitCursor
            Dim dtDat As DataTable = CType(grvBudget4.DataSource, DataTable).DefaultView.Table.Copy

            '// Auto Calculation
            For intRowIndex As Integer = 0 To dtDat.Rows.Count - 1
                '// Calc Diff Year
                dblTotal = CDbl(Nz(dtDat.Rows(intRowIndex)![RevYear], 0)) - CDbl(Nz(dtDat.Rows(intRowIndex)![PrevRRT1], 0))
                If dblTotal = 0 Then
                    dtDat.Rows(intRowIndex)![DiffYear] = DBNull.Value
                Else
                    dtDat.Rows(intRowIndex)![DiffYear] = dblTotal
                End If

                '// Calc Diff Year 1 
                dblTotal1 = CDbl(Nz(dtDat.Rows(intRowIndex)![RRT1], 0)) - CDbl(Nz(dtDat.Rows(intRowIndex)![PrevRRT2], 0))
                If dblTotal1 = 0 Then
                    dtDat.Rows(intRowIndex)![DiffYear1] = DBNull.Value
                Else
                    dtDat.Rows(intRowIndex)![DiffYear1] = dblTotal1
                End If

                '// Calc Diff Year 2
                dblTotal2 = CDbl(Nz(dtDat.Rows(intRowIndex)![RRT2], 0)) - CDbl(Nz(dtDat.Rows(intRowIndex)![PrevRRT3], 0))
                If dblTotal2 = 0 Then
                    dtDat.Rows(intRowIndex)![DiffYear2] = DBNull.Value
                Else
                    dtDat.Rows(intRowIndex)![DiffYear2] = dblTotal2
                End If

            Next
            grvBudget4.AutoGenerateColumns = False
            grvBudget4.DataSource = dtDat

            '// Calc MTP Budget '---1. Change Input MTP Budget (Budget Journal) Comment 8/8/2018 
            If Me.GetBudgetType() = P_BUDGET_TYPE_EXPENSE Then
                For intRowIndex As Integer = 0 To dtDat.Rows.Count - 1
                    CalcMTPBudgetNew(intRowIndex)
                Next
            End If

            FilterGridView()

            Me.Cursor = Cursors.Default
            Debug.Print(Now.ToString & ": End CalcReviseMTPBudget")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub CalcReviseBudget(ByVal blnCalcMTP As Boolean)
        Dim dblTotal As Double

        Try
            Debug.Print(Now.ToString & ": Begin CalcReviseBudget")
            Me.Cursor = Cursors.WaitCursor
            Dim dtDat As DataTable = CType(grvBudget3.DataSource, DataTable).DefaultView.Table.Copy

            '// Auto Calculation
            For intRowIndex As Integer = 0 To dtDat.Rows.Count - 1
                '// Calc Estimate 1st Half
                dblTotal = CDbl(Nz(dtDat.Rows(intRowIndex)![M1], 0)) + _
                            CDbl(Nz(dtDat.Rows(intRowIndex)![M2], 0)) + _
                            CDbl(Nz(dtDat.Rows(intRowIndex)![M3], 0)) + _
                            CDbl(Nz(dtDat.Rows(intRowIndex)![M4], 0)) + _
                            CDbl(Nz(dtDat.Rows(intRowIndex)![M5], 0)) + _
                            CDbl(Nz(dtDat.Rows(intRowIndex)![M6], 0))
                If dblTotal = 0 Then
                    dtDat.Rows(intRowIndex)![Est1H] = DBNull.Value
                Else
                    dtDat.Rows(intRowIndex)![Est1H] = dblTotal
                End If

                '// Calc Diff 1st Half
                dblTotal = CDbl(Nz(dtDat.Rows(intRowIndex)![Est1H], 0)) - CDbl(Nz(dtDat.Rows(intRowIndex)![IMP1], 0))
                If dblTotal = 0 Then
                    dtDat.Rows(intRowIndex)![Diff1H] = DBNull.Value
                Else
                    dtDat.Rows(intRowIndex)![Diff1H] = dblTotal
                End If

                '// Calc Estimate 2nd Half
                dblTotal = CDbl(Nz(dtDat.Rows(intRowIndex)![M7], 0)) + _
                            CDbl(Nz(dtDat.Rows(intRowIndex)![M8], 0)) + _
                            CDbl(Nz(dtDat.Rows(intRowIndex)![M9], 0)) + _
                            CDbl(Nz(dtDat.Rows(intRowIndex)![M10], 0)) + _
                            CDbl(Nz(dtDat.Rows(intRowIndex)![M11], 0)) + _
                            CDbl(Nz(dtDat.Rows(intRowIndex)![M12], 0))
                If dblTotal = 0 Then
                    dtDat.Rows(intRowIndex)![Rev2H] = DBNull.Value
                Else
                    dtDat.Rows(intRowIndex)![Rev2H] = dblTotal
                End If

                '// Calc Diff 2nd Half
                dblTotal = CDbl(Nz(dtDat.Rows(intRowIndex)![Rev2H], 0)) - CDbl(Nz(dtDat.Rows(intRowIndex)![IMP2], 0))
                If dblTotal = 0 Then
                    dtDat.Rows(intRowIndex)![Diff2H] = DBNull.Value
                Else
                    dtDat.Rows(intRowIndex)![Diff2H] = dblTotal
                End If

                '// Calc Revise Year
                dblTotal = CDbl(Nz(dtDat.Rows(intRowIndex)![Est1H], 0)) + CDbl(Nz(dtDat.Rows(intRowIndex)![Rev2H], 0))
                If dblTotal = 0 Then
                    dtDat.Rows(intRowIndex)![RevYear] = DBNull.Value
                Else
                    dtDat.Rows(intRowIndex)![RevYear] = dblTotal
                End If

                '// Calc Diff Year
                dblTotal = CDbl(Nz(dtDat.Rows(intRowIndex)![Diff1H], 0)) + CDbl(Nz(dtDat.Rows(intRowIndex)![Diff2H], 0))
                If dblTotal = 0 Then
                    dtDat.Rows(intRowIndex)![DiffYear] = DBNull.Value
                Else
                    dtDat.Rows(intRowIndex)![DiffYear] = dblTotal
                End If
            Next

            grvBudget3.DataSource = dtDat
            FilterGridView()
            Me.Cursor = Cursors.Default
            Debug.Print(Now.ToString & ": End CalcReviseBudget")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub CalcReviseBudget(ByVal intRowIndex As Integer)
        Dim dblTotal As Double

        Try
            Debug.Print(Now.ToString & ": Begin CalcReviseBudget2")
            Me.Cursor = Cursors.WaitCursor

            '// Auto Calculation
            '// Calc Estimate 1st Half
            dblTotal = CDbl(Nz(grvBudget3.Item("g3col7", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget3.Item("g3col8", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget3.Item("g3col9", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget3.Item("g3col10", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget3.Item("g3col11", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget3.Item("g3col12", intRowIndex).Value, 0))
            If dblTotal = 0 Then
                grvBudget3.Item("g3col13", intRowIndex).Value = DBNull.Value
            Else
                grvBudget3.Item("g3col13", intRowIndex).Value = dblTotal
            End If

            '// Calc Diff 1st Half
            dblTotal = CDbl(Nz(grvBudget3.Item("g3col13", intRowIndex).Value, 0)) - _
                        CDbl(Nz(grvBudget3.Item("g3col6", intRowIndex).Value, 0))
            If dblTotal = 0 Then
                grvBudget3.Item("g3col14", intRowIndex).Value = DBNull.Value
            Else
                grvBudget3.Item("g3col14", intRowIndex).Value = dblTotal
            End If

            '// Calc Estimate 2nd Half
            dblTotal = CDbl(Nz(grvBudget3.Item("g3col16", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget3.Item("g3col17", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget3.Item("g3col18", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget3.Item("g3col19", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget3.Item("g3col20", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget3.Item("g3col21", intRowIndex).Value, 0))
            If dblTotal = 0 Then
                grvBudget3.Item("g3col22", intRowIndex).Value = DBNull.Value
            Else
                grvBudget3.Item("g3col22", intRowIndex).Value = dblTotal
            End If

            '// Calc Diff 2nd Half
            dblTotal = CDbl(Nz(grvBudget3.Item("g3col22", intRowIndex).Value, 0)) - _
                        CDbl(Nz(grvBudget3.Item("g3col15", intRowIndex).Value, 0))
            If dblTotal = 0 Then
                grvBudget3.Item("g3col23", intRowIndex).Value = DBNull.Value
            Else
                grvBudget3.Item("g3col23", intRowIndex).Value = dblTotal
            End If

            '// Calc Revise Year
            dblTotal = CDbl(Nz(grvBudget3.Item("g3col13", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget3.Item("g3col22", intRowIndex).Value, 0))
            If dblTotal = 0 Then
                grvBudget3.Item("g3col24", intRowIndex).Value = DBNull.Value
            Else
                grvBudget3.Item("g3col24", intRowIndex).Value = dblTotal
            End If

            '// Calc Diff Year
            dblTotal = CDbl(Nz(grvBudget3.Item("g3col14", intRowIndex).Value, 0)) + _
                        CDbl(Nz(grvBudget3.Item("g3col23", intRowIndex).Value, 0))
            If dblTotal = 0 Then
                grvBudget3.Item("g3col25", intRowIndex).Value = DBNull.Value
            Else
                grvBudget3.Item("g3col25", intRowIndex).Value = dblTotal
            End If

            Me.Cursor = Cursors.Default
            Debug.Print(Now.ToString & ": Begin CalcReviseBudget2")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub CalcMTPBudgetNew(ByVal intRow As Integer)
        Dim dblReviseYear As Double
        Dim strRRT(5) As String
        Dim intRRT As Integer
        Dim strOrderNo As String

        Try
            Debug.Print(Now.ToString & ": Begin CalcMTPBudget")
            mySetGridValue = True 'Add by Max
            strRRT(1) = lblRRT1p.Text.Replace("%", "")
            strRRT(2) = lblRRT2p.Text.Replace("%", "")
            strRRT(3) = lblRRT3p.Text.Replace("%", "")
            strRRT(4) = lblRRT4p.Text.Replace("%", "")
            strRRT(5) = lblRRT5p.Text.Replace("%", "")

            If IsNumeric(strRRT(1)) Or IsNumeric(strRRT(2)) Or IsNumeric(strRRT(3)) Or IsNumeric(strRRT(4)) Or IsNumeric(strRRT(5)) Then
                ''// get [Revise Year] value
                'dblReviseYear = CDbl(Nz(grvBudget4.Item("g4col6", intRow).Value, 0))
                'strOrderNo = CStr(grvBudget4.Item("OrderNo4", intRow).Value)

                'If dblReviseYear = 0 Then
                '    If (Me.OperationCd = enumOperationCd.InputBudget Or (Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput)) And _
                '        IsDBNull(m_dtCheckMTPNew.Select("OrderNo = '" & strOrderNo & "'")(0)![RRT1]) Then
                '        grvBudget4.Item("g4col9", intRow).Value = 0
                '    End If
                '    grvBudget4.Item("g4ex01", intRow).Value = 0

                '    If (Me.OperationCd = enumOperationCd.InputBudget Or (Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput)) And _
                '        IsDBNull(m_dtCheckMTPNew.Select("OrderNo = '" & strOrderNo & "'")(0)![RRT2]) Then
                '        grvBudget4.Item("g4col11", intRow).Value = 0
                '    End If
                '    grvBudget4.Item("g4ex02", intRow).Value = 0

                '    If (Me.OperationCd = enumOperationCd.InputBudget Or (Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput)) And _
                '        IsDBNull(m_dtCheckMTPNew.Select("OrderNo = '" & strOrderNo & "'")(0)![RRT3]) Then
                '        grvBudget4.Item("g4col13", intRow).Value = 0
                '    End If
                '    grvBudget4.Item("g4ex03", intRow).Value = 0

                '    If (Me.OperationCd = enumOperationCd.InputBudget Or (Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput)) And _
                '        IsDBNull(m_dtCheckMTPNew.Select("OrderNo = '" & strOrderNo & "'")(0)![RRT4]) Then
                '        grvBudget4.Item("g4col15", intRow).Value = 0
                '    End If
                '    grvBudget4.Item("g4ex04", intRow).Value = 0

                '    If (Me.OperationCd = enumOperationCd.InputBudget Or (Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput)) And _
                '        IsDBNull(m_dtCheckMTPNew.Select("OrderNo = '" & strOrderNo & "'")(0)![RRT5]) Then
                '        grvBudget4.Item("g4col17", intRow).Value = 0
                '    End If
                '    grvBudget4.Item("g4ex05", intRow).Value = 0

                'Else
                '    If CStr(Nz(grvBudget4.Item("g4col3", intRow).Value)) = "Fixed Cost" Then
                '        If (Me.OperationCd = enumOperationCd.InputBudget Or (Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput)) And _
                '        IsDBNull(m_dtCheckMTPNew.Select("OrderNo = '" & strOrderNo & "'")(0)![RRT1]) Then
                '            grvBudget4.Item("g4col9", intRow).Value = dblReviseYear
                '        End If
                '        grvBudget4.Item("g4ex01", intRow).Value = dblReviseYear

                '        If (Me.OperationCd = enumOperationCd.InputBudget Or (Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput)) And _
                '        IsDBNull(m_dtCheckMTPNew.Select("OrderNo = '" & strOrderNo & "'")(0)![RRT2]) Then
                '            grvBudget4.Item("g4col11", intRow).Value = dblReviseYear
                '        End If
                '        grvBudget4.Item("g4ex02", intRow).Value = dblReviseYear

                '        If (Me.OperationCd = enumOperationCd.InputBudget Or (Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput)) And _
                '        IsDBNull(m_dtCheckMTPNew.Select("OrderNo = '" & strOrderNo & "'")(0)![RRT3]) Then
                '            grvBudget4.Item("g4col13", intRow).Value = dblReviseYear
                '        End If
                '        grvBudget4.Item("g4ex03", intRow).Value = dblReviseYear

                '        If (Me.OperationCd = enumOperationCd.InputBudget Or (Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput)) And _
                '        IsDBNull(m_dtCheckMTPNew.Select("OrderNo = '" & strOrderNo & "'")(0)![RRT4]) Then
                '            grvBudget4.Item("g4col15", intRow).Value = dblReviseYear
                '        End If
                '        grvBudget4.Item("g4ex04", intRow).Value = dblReviseYear

                '        If (Me.OperationCd = enumOperationCd.InputBudget Or (Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput)) And _
                '        IsDBNull(m_dtCheckMTPNew.Select("OrderNo = '" & strOrderNo & "'")(0)![RRT5]) Then
                '            grvBudget4.Item("g4col17", intRow).Value = dblReviseYear
                '        End If
                '        grvBudget4.Item("g4ex05", intRow).Value = dblReviseYear

                '    Else
                '        If IsNumeric(strRRT(1)) Then
                '            intRRT = CInt(strRRT(1))
                '            If (Me.OperationCd = enumOperationCd.InputBudget Or (Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput)) And _
                '            IsDBNull(m_dtCheckMTPNew.Select("OrderNo = '" & strOrderNo & "'")(0)![RRT1]) Then
                '                grvBudget4.Item("g4col9", intRow).Value = dblReviseYear / 100 * intRRT
                '            End If
                '            grvBudget4.Item("g4ex01", intRow).Value = dblReviseYear / 100 * intRRT
                '        Else
                '            grvBudget4.Item("g4ex01", intRow).Value = 0
                '        End If

                '        If IsNumeric(strRRT(2)) Then
                '            intRRT = CInt(strRRT(2))
                '            If (Me.OperationCd = enumOperationCd.InputBudget Or (Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput)) And _
                '            IsDBNull(m_dtCheckMTPNew.Select("OrderNo = '" & strOrderNo & "'")(0)![RRT2]) Then
                '                grvBudget4.Item("g4col11", intRow).Value = dblReviseYear / 100 * intRRT
                '            End If
                '            grvBudget4.Item("g4ex02", intRow).Value = dblReviseYear / 100 * intRRT
                '        Else
                '            grvBudget4.Item("g4ex02", intRow).Value = 0
                '        End If

                '        If IsNumeric(strRRT(3)) Then
                '            intRRT = CInt(strRRT(3))
                '            If (Me.OperationCd = enumOperationCd.InputBudget Or (Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput)) And _
                '            IsDBNull(m_dtCheckMTPNew.Select("OrderNo = '" & strOrderNo & "'")(0)![RRT3]) Then
                '                grvBudget4.Item("g4col13", intRow).Value = dblReviseYear / 100 * intRRT
                '            End If
                '            grvBudget4.Item("g4ex03", intRow).Value = dblReviseYear / 100 * intRRT
                '        Else
                '            grvBudget4.Item("g4ex03", intRow).Value = 0
                '        End If

                '        If IsNumeric(strRRT(4)) Then
                '            intRRT = CInt(strRRT(4))
                '            If (Me.OperationCd = enumOperationCd.InputBudget Or (Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput)) And _
                '            IsDBNull(m_dtCheckMTPNew.Select("OrderNo = '" & strOrderNo & "'")(0)![RRT4]) Then
                '                grvBudget4.Item("g4col15", intRow).Value = dblReviseYear / 100 * intRRT
                '            End If
                '            grvBudget4.Item("g4ex04", intRow).Value = dblReviseYear / 100 * intRRT
                '        Else
                '            grvBudget4.Item("g4ex04", intRow).Value = 0
                '        End If

                '        If IsNumeric(strRRT(5)) Then
                '            intRRT = CInt(strRRT(5))
                '            If (Me.OperationCd = enumOperationCd.InputBudget Or (Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput)) And _
                '            IsDBNull(m_dtCheckMTPNew.Select("OrderNo = '" & strOrderNo & "'")(0)![RRT5]) Then
                '                grvBudget4.Item("g4col17", intRow).Value = dblReviseYear / 100 * intRRT
                '            End If
                '            grvBudget4.Item("g4ex05", intRow).Value = dblReviseYear / 100 * intRRT
                '        Else
                '            grvBudget4.Item("g4ex05", intRow).Value = 0
                '        End If
                '    End If
                'End If

                CheckValidateMTPBudget(grvBudget4.Columns("g4col9").Index, intRow)
                CheckValidateMTPBudget(grvBudget4.Columns("g4col11").Index, intRow)
                CheckValidateMTPBudget(grvBudget4.Columns("g4col13").Index, intRow)
                CheckValidateMTPBudget(grvBudget4.Columns("g4col15").Index, intRow)
                CheckValidateMTPBudget(grvBudget4.Columns("g4col17").Index, intRow)
                CheckValidateMTPBudget(grvBudget4.Columns("g4ex01").Index, intRow)
                CheckValidateMTPBudget(grvBudget4.Columns("g4ex02").Index, intRow)
                CheckValidateMTPBudget(grvBudget4.Columns("g4ex03").Index, intRow)
                CheckValidateMTPBudget(grvBudget4.Columns("g4ex04").Index, intRow)
                CheckValidateMTPBudget(grvBudget4.Columns("g4ex05").Index, intRow)
            End If

            mySetGridValue = False
            Debug.Print(Now.ToString & ": End CalcMTPBudget")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub HighlightMTPValue(ByVal intCol As Integer, ByVal intRow As Integer)
        Try
            If IsNumeric(myClsBG0200BL.MTPHighlightValue) Then
                If grvBudget3.Columns(intCol).Name = "g3col26" Or _
                    grvBudget3.Columns(intCol).Name = "g3col27" Or _
                    grvBudget3.Columns(intCol).Name = "g3col28" Or _
                    grvBudget3.Columns(intCol).Name = "g3col29" Or _
                    grvBudget3.Columns(intCol).Name = "g3col30" Then

                    If CDbl(Nz(grvBudget3.Item(intCol, intRow).Value, 0)) >= CDbl(myClsBG0200BL.MTPHighlightValue) Then
                        grvBudget3.Item(intCol, intRow).Style.BackColor = Color.Red
                        grvBudget3.Item(intCol, intRow).Style.SelectionBackColor = Color.LightCoral
                    Else
                        If intRow Mod 2 = 0 Then
                            grvBudget3.Item(intCol, intRow).Style.BackColor = grvBudget3.Columns(intCol).DefaultCellStyle.BackColor
                            grvBudget3.Item(intCol, intRow).Style.BackColor = grvBudget3.Columns(intCol).DefaultCellStyle.SelectionBackColor
                        Else
                            grvBudget3.Item(intCol, intRow).Style.BackColor = grvBudget3.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget3.Item(intCol, intRow).Style.BackColor = grvBudget3.AlternatingRowsDefaultCellStyle.SelectionBackColor
                        End If
                    End If

                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub HighlightMTPValueNew(ByVal intCol As Integer, ByVal intRow As Integer)
        Try
            If IsNumeric(myClsBG0200BL.MTPHighlightValue) Then
                If grvBudget4.Columns(intCol).Name = "g4col9" Or _
                    grvBudget4.Columns(intCol).Name = "g4col11" Or _
                    grvBudget4.Columns(intCol).Name = "g4col13" Or _
                    grvBudget4.Columns(intCol).Name = "g4col15" Or _
                    grvBudget4.Columns(intCol).Name = "g4col17" Then

                    If CDbl(Nz(grvBudget4.Item(intCol, intRow).Value, 0)) >= CDbl(myClsBG0200BL.MTPHighlightValue) Then
                        grvBudget4.Item(intCol, intRow).Style.BackColor = Color.Red
                        grvBudget4.Item(intCol, intRow).Style.SelectionBackColor = Color.LightCoral
                    Else
                        If intRow Mod 2 = 0 Then
                            grvBudget4.Item(intCol, intRow).Style.BackColor = grvBudget4.Columns(intCol).DefaultCellStyle.BackColor
                            grvBudget4.Item(intCol, intRow).Style.BackColor = grvBudget4.Columns(intCol).DefaultCellStyle.SelectionBackColor
                        Else
                            grvBudget4.Item(intCol, intRow).Style.BackColor = grvBudget4.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget4.Item(intCol, intRow).Style.BackColor = grvBudget4.AlternatingRowsDefaultCellStyle.SelectionBackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub HighlightMTPValueAll()
        Try
            If IsNumeric(myClsBG0200BL.MTPHighlightValue) Then
                For j = 0 To grvBudget3.RowCount - 1
                    For i = 26 To 30

                        If CDbl(Nz(grvBudget3.Item("g3col" & CStr(i), j).Value, 0)) >= CDbl(myClsBG0200BL.MTPHighlightValue) Then
                            grvBudget3.Item("g3col" & CStr(i), j).Style.BackColor = Color.Red
                            grvBudget3.Item("g3col" & CStr(i), j).Style.SelectionBackColor = Color.LightCoral
                        Else
                            If j Mod 2 = 0 Then
                                grvBudget3.Item("g3col" & CStr(i), j).Style.BackColor = grvBudget3.Columns("g3col" & CStr(i)).DefaultCellStyle.BackColor
                                grvBudget3.Item("g3col" & CStr(i), j).Style.BackColor = grvBudget3.Columns("g3col" & CStr(i)).DefaultCellStyle.SelectionBackColor
                            Else
                                grvBudget3.Item("g3col" & CStr(i), j).Style.BackColor = grvBudget3.AlternatingRowsDefaultCellStyle.BackColor
                                grvBudget3.Item("g3col" & CStr(i), j).Style.BackColor = grvBudget3.AlternatingRowsDefaultCellStyle.SelectionBackColor
                            End If
                        End If

                    Next
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub HighlightMTPValueAllNew()
        Try
            If IsNumeric(myClsBG0200BL.MTPHighlightValue) Then
                For j = 0 To grvBudget4.RowCount - 1
                    For i = 9 To 17 Step 2

                        If CDbl(Nz(grvBudget4.Item("g4col" & CStr(i), j).Value, 0)) >= CDbl(myClsBG0200BL.MTPHighlightValue) Then
                            grvBudget4.Item("g4col" & CStr(i), j).Style.BackColor = Color.Red
                            grvBudget4.Item("g4col" & CStr(i), j).Style.SelectionBackColor = Color.LightCoral
                        Else
                            If j Mod 2 = 0 Then
                                grvBudget4.Item("g4col" & CStr(i), j).Style.BackColor = grvBudget4.Columns("g4col" & CStr(i)).DefaultCellStyle.BackColor
                                grvBudget4.Item("g4col" & CStr(i), j).Style.BackColor = grvBudget4.Columns("g4col" & CStr(i)).DefaultCellStyle.SelectionBackColor
                            Else
                                grvBudget4.Item("g4col" & CStr(i), j).Style.BackColor = grvBudget4.AlternatingRowsDefaultCellStyle.BackColor
                                grvBudget4.Item("g4col" & CStr(i), j).Style.BackColor = grvBudget4.AlternatingRowsDefaultCellStyle.SelectionBackColor
                            End If
                        End If
                    Next

                    For i = 10 To 16 Step 2
                        If CDbl(Nz(grvBudget4.Item("g4col" & CStr(i), j).Value, 0)) >= CDbl(myClsBG0200BL.MTPHighlightValue) Then
                            grvBudget4.Item("g4col" & CStr(i), j).Style.BackColor = Color.Red
                            grvBudget4.Item("g4col" & CStr(i), j).Style.SelectionBackColor = Color.LightCoral
                        Else
                            If j Mod 2 = 0 Then
                                grvBudget4.Item("g4col" & CStr(i), j).Style.BackColor = grvBudget4.Columns("g4col" & CStr(i)).DefaultCellStyle.BackColor
                                grvBudget4.Item("g4col" & CStr(i), j).Style.BackColor = grvBudget4.Columns("g4col" & CStr(i)).DefaultCellStyle.SelectionBackColor
                            Else
                                grvBudget4.Item("g4col" & CStr(i), j).Style.BackColor = grvBudget4.AlternatingRowsDefaultCellStyle.BackColor
                                grvBudget4.Item("g4col" & CStr(i), j).Style.BackColor = grvBudget4.AlternatingRowsDefaultCellStyle.SelectionBackColor
                            End If
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' AdjustWorkingBG()
    ''' </summary>
    ''' <remarks>Last Updated: 2011/05/30 by S.Watcharapong</remarks>
    Private Sub AdjustWorkingBG()
        Dim dtRawData As DataTable = Nothing
        Dim dtRefData As DataTable = Nothing
        Dim dblWK1 As Double = 0
        Dim dblWK2 As Double = 0
        Dim dr As DataRow()
        Dim strOrderNo As String
        Dim dtmBeginTime As Date = Now

        Try
            Debug.Print(Now.ToString & ": Begin AdjustWorkingBG")

            myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
            myClsBG0200BL.PeriodType = Me.GetPeriodType()
            myClsBG0200BL.BudgetType = Me.GetBudgetType()
            myClsBG0200BL.UserPIC = CStr(cboPIC.SelectedValue)
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            myClsBG0200BL.MtpProjectNo = Me.GetMtpProjectNo()
            myClsBG0200BL.MtpRevNo = Me.GetMtpRevNo()

            '// Get Raw data
            myClsBG0200BL.RevNo = "1"
            If myClsBG0200BL.GetBudgetData() = False Then
                Exit Sub
            End If
            dtRawData = myClsBG0200BL.BudgetList

            '// Get Ref. data
            myClsBG0200BL.RevNo = CStr(IIf(IsNumeric(lblRefRevNo.Text), lblRefRevNo.Text, "1"))
            If myClsBG0200BL.GetBudgetData() = False Then
                Exit Sub
            End If
            dtRefData = myClsBG0200BL.BudgetList

            If IsNumeric(txtWk1.Text) And IsNumeric(txtWk2.Text) Then
                dblWK1 = CDbl(txtWk1.Text)
                dblWK2 = CDbl(txtWk2.Text)
            Else
                Exit Sub
            End If

            '// Set Working Budget
            If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then '// Original Budget
                For i = 0 To grvBudget1.RowCount - 1
                    strOrderNo = CStr(grvBudget1.Item("OrderNo1", i).Value)
                    dr = dtRawData.Select("BUDGET_ORDER_NO = '" & strOrderNo & "'")

                    If dr.Count > 0 Then
                        If CBool(grvBudget1.Item("g1wk", i).Value) = True Then
                            'grvBudget1.Item("g1Col8", i).Value = ((CDbl(Nz(dr(0).Item("M1"), 0)) / 100) * dblWK1).ToString("#,##0.00")
                            'grvBudget1.Item("g1Col9", i).Value = ((CDbl(Nz(dr(0).Item("M2"), 0)) / 100) * dblWK1).ToString("#,##0.00")
                            'grvBudget1.Item("g1Col10", i).Value = ((CDbl(Nz(dr(0).Item("M3"), 0)) / 100) * dblWK1).ToString("#,##0.00")
                            'grvBudget1.Item("g1Col11", i).Value = ((CDbl(Nz(dr(0).Item("M4"), 0)) / 100) * dblWK1).ToString("#,##0.00")
                            'grvBudget1.Item("g1Col12", i).Value = ((CDbl(Nz(dr(0).Item("M5"), 0)) / 100) * dblWK1).ToString("#,##0.00")
                            'grvBudget1.Item("g1Col13", i).Value = ((CDbl(Nz(dr(0).Item("M6"), 0)) / 100) * dblWK1).ToString("#,##0.00")

                            'grvBudget1.Item("g1Colex1", i).Value = ((CDbl(Nz(dr(0).Item("M7"), 0)) / 100) * dblWK2).ToString("#,##0.00")
                            'grvBudget1.Item("g1Colex2", i).Value = ((CDbl(Nz(dr(0).Item("M8"), 0)) / 100) * dblWK2).ToString("#,##0.00")
                            'grvBudget1.Item("g1Colex3", i).Value = ((CDbl(Nz(dr(0).Item("M9"), 0)) / 100) * dblWK2).ToString("#,##0.00")
                            'grvBudget1.Item("g1Colex4", i).Value = ((CDbl(Nz(dr(0).Item("M10"), 0)) / 100) * dblWK2).ToString("#,##0.00")
                            'grvBudget1.Item("g1Colex5", i).Value = ((CDbl(Nz(dr(0).Item("M11"), 0)) / 100) * dblWK2).ToString("#,##0.00")
                            'grvBudget1.Item("g1Colex6", i).Value = ((CDbl(Nz(dr(0).Item("M12"), 0)) / 100) * dblWK2).ToString("#,##0.00")
                            grvBudget1.Item("g1Col8", i).Value = 100
                            grvBudget1.Item("g1Col9", i).Value = 100
                            grvBudget1.Item("g1Col10", i).Value = 100
                            grvBudget1.Item("g1Col11", i).Value = 100
                            grvBudget1.Item("g1Col12", i).Value = 100
                            grvBudget1.Item("g1Col13", i).Value = 100

                            grvBudget1.Item("g1Colex1", i).Value = 100
                            grvBudget1.Item("g1Colex2", i).Value = 100
                            grvBudget1.Item("g1Colex3", i).Value = 100
                            grvBudget1.Item("g1Colex4", i).Value = 100
                            grvBudget1.Item("g1Colex5", i).Value = 100
                            grvBudget1.Item("g1Colex6", i).Value = 100
                        Else
                            dr = dtRefData.Select("BUDGET_ORDER_NO = '" & strOrderNo & "'")
                            If dr.Count = 0 Then
                                dr = dtRawData.Select("BUDGET_ORDER_NO = '" & strOrderNo & "'")
                            End If
                            grvBudget1.Item("g1Col8", i).Value = CDbl(Nz(dr(0).Item("M1"), 0)).ToString("#,##0.00")
                            grvBudget1.Item("g1Col9", i).Value = CDbl(Nz(dr(0).Item("M2"), 0)).ToString("#,##0.00")
                            grvBudget1.Item("g1Col10", i).Value = CDbl(Nz(dr(0).Item("M3"), 0)).ToString("#,##0.00")
                            grvBudget1.Item("g1Col11", i).Value = CDbl(Nz(dr(0).Item("M4"), 0)).ToString("#,##0.00")
                            grvBudget1.Item("g1Col12", i).Value = CDbl(Nz(dr(0).Item("M5"), 0)).ToString("#,##0.00")
                            grvBudget1.Item("g1Col13", i).Value = CDbl(Nz(dr(0).Item("M6"), 0)).ToString("#,##0.00")

                            grvBudget1.Item("g1Colex1", i).Value = CDbl(Nz(dr(0).Item("M7"), 0)).ToString("#,##0.00")
                            grvBudget1.Item("g1Colex2", i).Value = CDbl(Nz(dr(0).Item("M8"), 0)).ToString("#,##0.00")
                            grvBudget1.Item("g1Colex3", i).Value = CDbl(Nz(dr(0).Item("M9"), 0)).ToString("#,##0.00")
                            grvBudget1.Item("g1Colex4", i).Value = CDbl(Nz(dr(0).Item("M10"), 0)).ToString("#,##0.00")
                            grvBudget1.Item("g1Colex5", i).Value = CDbl(Nz(dr(0).Item("M11"), 0)).ToString("#,##0.00")
                            grvBudget1.Item("g1Colex6", i).Value = CDbl(Nz(dr(0).Item("M12"), 0)).ToString("#,##0.00")
                        End If
                    End If
                Next

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then '// Estimate Budget

                For i = 0 To grvBudget2.RowCount - 1
                    strOrderNo = CStr(grvBudget2.Item("OrderNo2", i).Value)
                    dr = dtRawData.Select("BUDGET_ORDER_NO = '" & strOrderNo & "'")

                    If dr.Count > 0 Then
                        If CBool(grvBudget2.Item("g2wk", i).Value) = True Then
                            grvBudget2.Item("g2Col11", i).Value = ((CDbl(Nz(dr(0).Item("M10"), 0)) / 100) * dblWK2).ToString("#,##0.00")
                            grvBudget2.Item("g2Col12", i).Value = ((CDbl(Nz(dr(0).Item("M11"), 0)) / 100) * dblWK2).ToString("#,##0.00")
                            grvBudget2.Item("g2Col13", i).Value = ((CDbl(Nz(dr(0).Item("M12"), 0)) / 100) * dblWK2).ToString("#,##0.00")
                        Else
                            dr = dtRefData.Select("BUDGET_ORDER_NO = '" & strOrderNo & "'")
                            If dr.Count = 0 Then
                                dr = dtRawData.Select("BUDGET_ORDER_NO = '" & strOrderNo & "'")
                            End If
                            grvBudget2.Item("g2Col11", i).Value = CDbl(Nz(dr(0).Item("M10"), 0)).ToString("#,##0.00")
                            grvBudget2.Item("g2Col12", i).Value = CDbl(Nz(dr(0).Item("M11"), 0)).ToString("#,##0.00")
                            grvBudget2.Item("g2Col13", i).Value = CDbl(Nz(dr(0).Item("M12"), 0)).ToString("#,##0.00")
                        End If
                    End If
                Next

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then '// Revise Budget

                For i = 0 To grvBudget3.RowCount - 1
                    strOrderNo = CStr(grvBudget3.Item("OrderNo3", i).Value)
                    dr = dtRawData.Select("BUDGET_ORDER_NO = '" & strOrderNo & "'")

                    If dr.Count > 0 Then
                        If CBool(grvBudget3.Item("g3wk", i).Value) = True Then
                            grvBudget3.Item("g3Col10", i).Value = ((CDbl(Nz(dr(0).Item("M4"), 0)) / 100) * dblWK1).ToString("#,##0.00")
                            grvBudget3.Item("g3Col11", i).Value = ((CDbl(Nz(dr(0).Item("M5"), 0)) / 100) * dblWK1).ToString("#,##0.00")
                            grvBudget3.Item("g3Col12", i).Value = ((CDbl(Nz(dr(0).Item("M6"), 0)) / 100) * dblWK1).ToString("#,##0.00")
                            grvBudget3.Item("g3Col16", i).Value = ((CDbl(Nz(dr(0).Item("M7"), 0)) / 100) * dblWK2).ToString("#,##0.00")
                            grvBudget3.Item("g3Col17", i).Value = ((CDbl(Nz(dr(0).Item("M8"), 0)) / 100) * dblWK2).ToString("#,##0.00")
                            grvBudget3.Item("g3Col18", i).Value = ((CDbl(Nz(dr(0).Item("M9"), 0)) / 100) * dblWK2).ToString("#,##0.00")
                            grvBudget3.Item("g3Col19", i).Value = ((CDbl(Nz(dr(0).Item("M10"), 0)) / 100) * dblWK2).ToString("#,##0.00")
                            grvBudget3.Item("g3Col20", i).Value = ((CDbl(Nz(dr(0).Item("M11"), 0)) / 100) * dblWK2).ToString("#,##0.00")
                            grvBudget3.Item("g3Col21", i).Value = ((CDbl(Nz(dr(0).Item("M12"), 0)) / 100) * dblWK2).ToString("#,##0.00")
                        Else
                            dr = dtRefData.Select("BUDGET_ORDER_NO = '" & strOrderNo & "'")

                            If dr.Count = 0 Then
                                dr = dtRawData.Select("BUDGET_ORDER_NO = '" & strOrderNo & "'")
                            End If
                            grvBudget3.Item("g3Col10", i).Value = CDbl(Nz(dr(0).Item("M4"), 0)).ToString("#,##0.00")
                            grvBudget3.Item("g3Col11", i).Value = CDbl(Nz(dr(0).Item("M5"), 0)).ToString("#,##0.00")
                            grvBudget3.Item("g3Col12", i).Value = CDbl(Nz(dr(0).Item("M6"), 0)).ToString("#,##0.00")
                            grvBudget3.Item("g3Col16", i).Value = CDbl(Nz(dr(0).Item("M7"), 0)).ToString("#,##0.00")
                            grvBudget3.Item("g3Col17", i).Value = CDbl(Nz(dr(0).Item("M8"), 0)).ToString("#,##0.00")
                            grvBudget3.Item("g3Col18", i).Value = CDbl(Nz(dr(0).Item("M9"), 0)).ToString("#,##0.00")
                            grvBudget3.Item("g3Col19", i).Value = CDbl(Nz(dr(0).Item("M10"), 0)).ToString("#,##0.00")
                            grvBudget3.Item("g3Col20", i).Value = CDbl(Nz(dr(0).Item("M11"), 0)).ToString("#,##0.00")
                            grvBudget3.Item("g3Col21", i).Value = CDbl(Nz(dr(0).Item("M12"), 0)).ToString("#,##0.00")
                        End If
                    End If
                Next

            End If

            Debug.Print(Now.ToString & ": End AdjustWorkingBG")
            Debug.Print("Calculate Time: " & DateDiff(DateInterval.Second, dtmBeginTime, Now) & " sec(s)")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub HighlightWorkingBG()
        Try
            Debug.Print(Now.ToString() & ": Begin HighlightWorkingBG")

            If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then  '//Original Budget
                For i = 0 To grvBudget1.RowCount - 1
                    If CBool(grvBudget1.Item("g1Wk", i).Value) = True Then
                        grvBudget1.Item("g1Col8", i).Style.BackColor = Color.OrangeRed
                        grvBudget1.Item("g1Col9", i).Style.BackColor = Color.OrangeRed
                        grvBudget1.Item("g1Col10", i).Style.BackColor = Color.OrangeRed
                        grvBudget1.Item("g1Col11", i).Style.BackColor = Color.OrangeRed
                        grvBudget1.Item("g1Col12", i).Style.BackColor = Color.OrangeRed
                        grvBudget1.Item("g1Col13", i).Style.BackColor = Color.OrangeRed
                        grvBudget1.Item("g1Colex1", i).Style.BackColor = Color.OrangeRed
                        grvBudget1.Item("g1Colex2", i).Style.BackColor = Color.OrangeRed
                        grvBudget1.Item("g1Colex3", i).Style.BackColor = Color.OrangeRed
                        grvBudget1.Item("g1Colex4", i).Style.BackColor = Color.OrangeRed
                        grvBudget1.Item("g1Colex5", i).Style.BackColor = Color.OrangeRed
                        grvBudget1.Item("g1Colex6", i).Style.BackColor = Color.OrangeRed
                    Else
                        If i Mod 2 = 0 Then
                            grvBudget1.Item("g1Col8", i).Style.BackColor = grvBudget1.Columns("g1Col8").DefaultCellStyle.BackColor
                            grvBudget1.Item("g1Col9", i).Style.BackColor = grvBudget1.Columns("g1Col9").DefaultCellStyle.BackColor
                            grvBudget1.Item("g1Col10", i).Style.BackColor = grvBudget1.Columns("g1Col10").DefaultCellStyle.BackColor
                            grvBudget1.Item("g1Col11", i).Style.BackColor = grvBudget1.Columns("g1Col11").DefaultCellStyle.BackColor
                            grvBudget1.Item("g1Col12", i).Style.BackColor = grvBudget1.Columns("g1Col12").DefaultCellStyle.BackColor
                            grvBudget1.Item("g1Col13", i).Style.BackColor = grvBudget1.Columns("g1Col13").DefaultCellStyle.BackColor
                            grvBudget1.Item("g1Colex1", i).Style.BackColor = grvBudget1.Columns("g1Colex1").DefaultCellStyle.BackColor
                            grvBudget1.Item("g1Colex2", i).Style.BackColor = grvBudget1.Columns("g1Colex2").DefaultCellStyle.BackColor
                            grvBudget1.Item("g1Colex3", i).Style.BackColor = grvBudget1.Columns("g1Colex3").DefaultCellStyle.BackColor
                            grvBudget1.Item("g1Colex4", i).Style.BackColor = grvBudget1.Columns("g1Colex4").DefaultCellStyle.BackColor
                            grvBudget1.Item("g1Colex5", i).Style.BackColor = grvBudget1.Columns("g1Colex5").DefaultCellStyle.BackColor
                            grvBudget1.Item("g1Colex6", i).Style.BackColor = grvBudget1.Columns("g1Colex5").DefaultCellStyle.BackColor

                        Else
                            grvBudget1.Item("g1Col8", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget1.Item("g1Col9", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget1.Item("g1Col10", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget1.Item("g1Col11", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget1.Item("g1Col12", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget1.Item("g1Col13", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget1.Item("g1Colex1", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget1.Item("g1Colex2", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget1.Item("g1Colex3", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget1.Item("g1Colex4", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget1.Item("g1Colex5", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget1.Item("g1Colex6", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                        End If
                    End If
                Next
            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then  '// Estimate Budget
                For i = 0 To grvBudget2.RowCount - 1
                    If CBool(grvBudget2.Item("g2Wk", i).Value) = True Then
                        grvBudget2.Item("g2Col11", i).Style.BackColor = Color.OrangeRed
                        grvBudget2.Item("g2Col12", i).Style.BackColor = Color.OrangeRed
                        grvBudget2.Item("g2Col13", i).Style.BackColor = Color.OrangeRed
                    Else
                        If i Mod 2 = 0 Then
                            grvBudget2.Item("g2Col11", i).Style.BackColor = grvBudget2.Columns("g2Col11").DefaultCellStyle.BackColor
                            grvBudget2.Item("g2Col12", i).Style.BackColor = grvBudget2.Columns("g2Col12").DefaultCellStyle.BackColor
                            grvBudget2.Item("g2Col13", i).Style.BackColor = grvBudget2.Columns("g2Col13").DefaultCellStyle.BackColor
                        Else
                            grvBudget2.Item("g2Col11", i).Style.BackColor = grvBudget2.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget2.Item("g2Col12", i).Style.BackColor = grvBudget2.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget2.Item("g2Col13", i).Style.BackColor = grvBudget2.AlternatingRowsDefaultCellStyle.BackColor
                        End If
                    End If
                Next
            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then    '// Revise Budget
                For i = 0 To grvBudget3.RowCount - 1
                    If CBool(grvBudget3.Item("g3Wk", i).Value) = True Then
                        grvBudget3.Item("g3Col10", i).Style.BackColor = Color.OrangeRed
                        grvBudget3.Item("g3Col11", i).Style.BackColor = Color.OrangeRed
                        grvBudget3.Item("g3Col12", i).Style.BackColor = Color.OrangeRed
                        grvBudget3.Item("g3Col16", i).Style.BackColor = Color.OrangeRed
                        grvBudget3.Item("g3Col17", i).Style.BackColor = Color.OrangeRed
                        grvBudget3.Item("g3Col18", i).Style.BackColor = Color.OrangeRed
                        grvBudget3.Item("g3Col19", i).Style.BackColor = Color.OrangeRed
                        grvBudget3.Item("g3Col20", i).Style.BackColor = Color.OrangeRed
                        grvBudget3.Item("g3Col21", i).Style.BackColor = Color.OrangeRed
                    Else
                        If i Mod 2 = 0 Then
                            grvBudget3.Item("g3Col10", i).Style.BackColor = grvBudget3.Columns("g3Col10").DefaultCellStyle.BackColor
                            grvBudget3.Item("g3Col11", i).Style.BackColor = grvBudget3.Columns("g3Col11").DefaultCellStyle.BackColor
                            grvBudget3.Item("g3Col12", i).Style.BackColor = grvBudget3.Columns("g3Col12").DefaultCellStyle.BackColor
                            grvBudget3.Item("g3Col16", i).Style.BackColor = grvBudget3.Columns("g3Col16").DefaultCellStyle.BackColor
                            grvBudget3.Item("g3Col17", i).Style.BackColor = grvBudget3.Columns("g3Col17").DefaultCellStyle.BackColor
                            grvBudget3.Item("g3Col18", i).Style.BackColor = grvBudget3.Columns("g3Col18").DefaultCellStyle.BackColor
                            grvBudget3.Item("g3Col19", i).Style.BackColor = grvBudget3.Columns("g3Col19").DefaultCellStyle.BackColor
                            grvBudget3.Item("g3Col20", i).Style.BackColor = grvBudget3.Columns("g3Col20").DefaultCellStyle.BackColor
                            grvBudget3.Item("g3Col21", i).Style.BackColor = grvBudget3.Columns("g3Col21").DefaultCellStyle.BackColor
                        Else
                            grvBudget3.Item("g3Col10", i).Style.BackColor = grvBudget3.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget3.Item("g3Col11", i).Style.BackColor = grvBudget3.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget3.Item("g3Col12", i).Style.BackColor = grvBudget3.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget3.Item("g3Col16", i).Style.BackColor = grvBudget3.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget3.Item("g3Col17", i).Style.BackColor = grvBudget3.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget3.Item("g3Col18", i).Style.BackColor = grvBudget3.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget3.Item("g3Col19", i).Style.BackColor = grvBudget3.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget3.Item("g3Col20", i).Style.BackColor = grvBudget3.AlternatingRowsDefaultCellStyle.BackColor
                            grvBudget3.Item("g3Col21", i).Style.BackColor = grvBudget3.AlternatingRowsDefaultCellStyle.BackColor
                        End If
                    End If
                Next
            End If

            Debug.Print(Now.ToString() & ": End HighlightWorkingBG")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub HighlightTransferCost()
        Dim strOrderNo As String
        Dim dr As DataRow()
        Dim dt As DataTable = Nothing

        Try
            If myClsBG0200BL.TransferList Is Nothing OrElse myClsBG0200BL.TransferList.Rows.Count = 0 Then
                Exit Sub
            End If

            Debug.Print(Now.ToString() & ": Begin HighlightTransferCost")
            '// Highlight Budget Order
            If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then      '// Original Budget
                For i = 0 To grvBudget1.RowCount - 1
                    strOrderNo = CStr(grvBudget1.Item("OrderNo1", i).Value)
                    dr = myClsBG0200BL.TransferList.Select("FROM_ORDER_NO = '" & strOrderNo & "'")
                    If dr.Count > 0 Then
                        '// Highlight 'From Order'
                        grvBudget1.Item("g1Col8", i).Style.BackColor = Color.LightSalmon
                        grvBudget1.Item("g1Col9", i).Style.BackColor = Color.LightSalmon
                        grvBudget1.Item("g1Col10", i).Style.BackColor = Color.LightSalmon
                        grvBudget1.Item("g1Col11", i).Style.BackColor = Color.LightSalmon
                        grvBudget1.Item("g1Col12", i).Style.BackColor = Color.LightSalmon
                        grvBudget1.Item("g1Col13", i).Style.BackColor = Color.LightSalmon
                        grvBudget1.Item("g1Colex1", i).Style.BackColor = Color.LightSalmon
                        grvBudget1.Item("g1Colex2", i).Style.BackColor = Color.LightSalmon
                        grvBudget1.Item("g1Colex3", i).Style.BackColor = Color.LightSalmon
                        grvBudget1.Item("g1Colex4", i).Style.BackColor = Color.LightSalmon
                        grvBudget1.Item("g1Colex5", i).Style.BackColor = Color.LightSalmon
                        grvBudget1.Item("g1Colex6", i).Style.BackColor = Color.LightSalmon
                    Else
                        dr = myClsBG0200BL.TransferList.Select("TO_ORDER_NO = '" & strOrderNo & "'")
                        If dr.Count > 0 Then
                            '// Highlight 'To Order'
                            grvBudget1.Item("g1Col8", i).Style.BackColor = Color.LightGreen
                            grvBudget1.Item("g1Col9", i).Style.BackColor = Color.LightGreen
                            grvBudget1.Item("g1Col10", i).Style.BackColor = Color.LightGreen
                            grvBudget1.Item("g1Col11", i).Style.BackColor = Color.LightGreen
                            grvBudget1.Item("g1Col12", i).Style.BackColor = Color.LightGreen
                            grvBudget1.Item("g1Col13", i).Style.BackColor = Color.LightGreen
                            grvBudget1.Item("g1Colex1", i).Style.BackColor = Color.LightGreen
                            grvBudget1.Item("g1Colex2", i).Style.BackColor = Color.LightGreen
                            grvBudget1.Item("g1Colex3", i).Style.BackColor = Color.LightGreen
                            grvBudget1.Item("g1Colex4", i).Style.BackColor = Color.LightGreen
                            grvBudget1.Item("g1Colex5", i).Style.BackColor = Color.LightGreen
                            grvBudget1.Item("g1Colex6", i).Style.BackColor = Color.LightGreen
                        End If
                    End If
                Next

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then '// Estimate Budget
                For i = 0 To grvBudget2.RowCount - 1

                    strOrderNo = CStr(grvBudget2.Item("OrderNo2", i).Value)
                    dr = myClsBG0200BL.TransferList.Select("FROM_ORDER_NO = '" & strOrderNo & "'")
                    If dr.Count > 0 Then
                        '// Highlight 'From Order'
                        grvBudget2.Item("g2Col11", i).Style.BackColor = Color.LightSalmon
                        grvBudget2.Item("g2Col12", i).Style.BackColor = Color.LightSalmon
                        grvBudget2.Item("g2Col13", i).Style.BackColor = Color.LightSalmon
                    Else
                        dr = myClsBG0200BL.TransferList.Select("TO_ORDER_NO = '" & strOrderNo & "'")
                        If dr.Count > 0 Then
                            '// Highlight 'To Order'
                            grvBudget2.Item("g2Col11", i).Style.BackColor = Color.LightGreen
                            grvBudget2.Item("g2Col12", i).Style.BackColor = Color.LightGreen
                            grvBudget2.Item("g2Col13", i).Style.BackColor = Color.LightGreen
                        End If
                    End If
                Next
            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then    '// Revise Budget
                For i = 0 To grvBudget3.RowCount - 1
                    strOrderNo = CStr(grvBudget3.Item("OrderNo3", i).Value)
                    dr = myClsBG0200BL.TransferList.Select("FROM_ORDER_NO = '" & strOrderNo & "'")
                    If dr.Count > 0 Then
                        '// Highlight 'From Order'
                        grvBudget3.Item("g3Col10", i).Style.BackColor = Color.LightSalmon
                        grvBudget3.Item("g3Col11", i).Style.BackColor = Color.LightSalmon
                        grvBudget3.Item("g3Col12", i).Style.BackColor = Color.LightSalmon
                        grvBudget3.Item("g3Col16", i).Style.BackColor = Color.LightSalmon
                        grvBudget3.Item("g3Col17", i).Style.BackColor = Color.LightSalmon
                        grvBudget3.Item("g3Col18", i).Style.BackColor = Color.LightSalmon
                        grvBudget3.Item("g3Col19", i).Style.BackColor = Color.LightSalmon
                        grvBudget3.Item("g3Col20", i).Style.BackColor = Color.LightSalmon
                        grvBudget3.Item("g3Col21", i).Style.BackColor = Color.LightSalmon
                    Else
                        dr = myClsBG0200BL.TransferList.Select("TO_ORDER_NO = '" & strOrderNo & "'")
                        If dr.Count > 0 Then
                            '// Highlight 'To Order'
                            grvBudget3.Item("g3Col10", i).Style.BackColor = Color.LightGreen
                            grvBudget3.Item("g3Col11", i).Style.BackColor = Color.LightGreen
                            grvBudget3.Item("g3Col12", i).Style.BackColor = Color.LightGreen
                            grvBudget3.Item("g3Col16", i).Style.BackColor = Color.LightGreen
                            grvBudget3.Item("g3Col17", i).Style.BackColor = Color.LightGreen
                            grvBudget3.Item("g3Col18", i).Style.BackColor = Color.LightGreen
                            grvBudget3.Item("g3Col19", i).Style.BackColor = Color.LightGreen
                            grvBudget3.Item("g3Col20", i).Style.BackColor = Color.LightGreen
                            grvBudget3.Item("g3Col21", i).Style.BackColor = Color.LightGreen
                        End If
                    End If
                Next
            End If

            Debug.Print(Now.ToString() & ": End HighlightTransferCost")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function SelectDistinct(ByVal SourceTable As DataTable, ByVal ParamArray FieldNames() As String) As DataTable
        Dim lastValues() As Object
        Dim newTable As DataTable

        Try
            If FieldNames Is Nothing OrElse FieldNames.Length = 0 Then
                Throw New ArgumentNullException("FieldNames")
            End If

            lastValues = New Object(FieldNames.Length - 1) {}
            newTable = New DataTable

            For Each field As String In FieldNames
                newTable.Columns.Add(field, GetType(String))
            Next

            For Each Row As DataRow In SourceTable.Select("", String.Join(", ", FieldNames))
                If Not FieldValuesAreEqual(lastValues, Row, FieldNames) Then
                    newTable.Rows.Add(CreateRowClone(Row, newTable.NewRow(), FieldNames))
                    SetLastValues(lastValues, Row, FieldNames)
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try

        Return newTable
    End Function

    Private Function FieldValuesAreEqual(ByVal lastValues() As Object, ByVal currentRow As DataRow, ByVal fieldNames() As String) As Boolean
        Dim areEqual As Boolean = True

        Try
            For i As Integer = 0 To fieldNames.Length - 1
                If lastValues(i) Is Nothing OrElse Not lastValues(i).Equals(currentRow(fieldNames(i))) Then
                    areEqual = False
                    Exit For
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try

        Return areEqual
    End Function

    Private Function CreateRowClone(ByVal sourceRow As DataRow, ByVal newRow As DataRow, ByVal fieldNames() As String) As DataRow
        Try
            For Each field As String In fieldNames
                newRow(field) = sourceRow(field)
            Next
        Catch ex As Exception
            Throw ex
        End Try

        Return newRow
    End Function

    Private Sub SetLastValues(ByVal lastValues() As Object, ByVal sourceRow As DataRow, ByVal fieldNames() As String)
        Try
            For i As Integer = 0 To fieldNames.Length - 1
                lastValues(i) = sourceRow(fieldNames(i))
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function UpRevision() As Boolean
        Try
            myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
            myClsBG0200BL.PeriodType = Me.GetPeriodType()
            myClsBG0200BL.BudgetType = Me.GetBudgetType()
            myClsBG0200BL.UserId = p_strUserId
            myClsBG0200BL.UserPIC = "0000"
            myClsBG0200BL.PicList = CType(cboPIC.DataSource, DataTable)
            myClsBG0200BL.RevNo = Me.CurrRevNo
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then          '// Original Budget
                myClsBG0200BL.BudgetList = CType(grvBudget1.DataSource, DataTable)

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then      '// Estimate Budget
                myClsBG0200BL.BudgetList = CType(grvBudget2.DataSource, DataTable)

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then        '// Revise Budget
                myClsBG0200BL.BudgetList = CType(grvBudget3.DataSource, DataTable)

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then        '// MTP Budget
                myClsBG0200BL.BudgetList = CType(grvBudget4.DataSource, DataTable)
            End If

            Return myClsBG0200BL.UpRevision()
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function DeleteRevision(ByVal RevNo As String) As Boolean
        Try
            myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
            myClsBG0200BL.PeriodType = Me.GetPeriodType()
            myClsBG0200BL.BudgetType = Me.GetBudgetType()
            myClsBG0200BL.RevNo = RevNo
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            Return myClsBG0200BL.DeleteRevision()

            HighlightWorkingBGAndComment()
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function DeleteBudgetData(ByVal UserPic As String, ByVal RevNo As String) As Boolean
        Try
            myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
            myClsBG0200BL.PeriodType = Me.GetPeriodType()
            myClsBG0200BL.BudgetType = Me.GetBudgetType()
            myClsBG0200BL.UserPIC = UserPic
            myClsBG0200BL.RevNo = RevNo
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            Return myClsBG0200BL.DeleteBudgetData()
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub LoadRevNoList()
        Try
            myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
            myClsBG0200BL.PeriodType = Me.GetPeriodType()
            myClsBG0200BL.BudgetType = Me.GetBudgetType()
            myClsBG0200BL.Status = CStr(enumBudgetStatus.Approve)
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            If myClsBG0200BL.GetRevNoList() = True Then
                myShowWarning = False
                cboRevNo.DisplayMember = "REV_NO"
                cboRevNo.ValueMember = "REV_NO"
                cboRevNo.DataSource = myClsBG0200BL.RevNoList
                myShowWarning = True
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub GetCurrentRevNo()
        Try
            If Me.OperationCd = enumOperationCd.AdjustBudget Or _
                    Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput Or _
                    Me.OperationCd = enumOperationCd.Authorize1 Or _
                    Me.OperationCd = enumOperationCd.Authorize2 Then
                '// Get Max Rev No.
                myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
                myClsBG0200BL.PeriodType = Me.GetPeriodType()
                myClsBG0200BL.BudgetType = Me.GetBudgetType()
                myClsBG0200BL.ProjectNo = Me.GetProjectNo

                If myClsBG0200BL.GetMaxRevNo() = True Then
                    Me.CurrRevNo = myClsBG0200BL.RevNo
                Else
                    Me.CurrRevNo = "1"
                End If
            ElseIf Me.OperationCd = enumOperationCd.ViewBudget Then
                '// Get status of Max Rev No.
                myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
                myClsBG0200BL.PeriodType = Me.GetPeriodType()
                myClsBG0200BL.BudgetType = Me.GetBudgetType()
                myClsBG0200BL.ProjectNo = Me.GetProjectNo()

                '// If status of max rev is more than "Adjusted" then return rev no = max rev, Otherwise rev no = 1 
                If myClsBG0200BL.GetMaxRevStatus() = True AndAlso CInt(myClsBG0200BL.Status) >= enumBudgetStatus.Adjust Then
                    Me.CurrRevNo = myClsBG0200BL.RevNo
                Else
                    Me.CurrRevNo = "1"
                End If
            Else
                Me.CurrRevNo = "1"
            End If

            If Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then
                myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
                myClsBG0200BL.PeriodType = CStr(enumPeriodType.ReviseBudget)
                myClsBG0200BL.BudgetType = Me.GetBudgetType()
                myClsBG0200BL.ProjectNo = Me.GetProjectNo

                If myClsBG0200BL.GetMaxRevNo() = True Then
                    CurrReviseRevNo = myClsBG0200BL.RevNo
                Else
                    CurrReviseRevNo = "1"
                End If
            Else
                CurrReviseRevNo = ""
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub LoadMtpRevNo()
        Try
            If Me.BudgetKey <> "" AndAlso Me.gbOB_MTP.Enabled = True AndAlso Me.cboMtpRevno.Visible = True Then
                Dim strProjectNo = Me.numMtpProjectNo.Value.ToString
                If Not strProjectNo Is Nothing And strProjectNo <> String.Empty And strProjectNo <> "System.Data.DataRowView" Then

                    myClsBG0310BL.BudgetYear = CStr(CInt(Me.GetBudgetYear()) - 1)
                    myClsBG0310BL.PeriodType = CStr(BGConstant.enumPeriodType.MTPBudget)
                    myClsBG0310BL.ProjectNo = strProjectNo
                    myClsBG0310BL.BudgetType = BGConstant.P_BUDGET_TYPE_EXPENSE

                    If myClsBG0310BL.GetRevNo() = True Then
                        Me.cboMtpRevno.DisplayMember = "REV_NO"
                        Me.cboMtpRevno.ValueMember = "REV_NO"
                        Me.cboMtpRevno.DataSource = myClsBG0310BL.RevNoList
                    Else
                        Me.cboMtpRevno.DataSource = Nothing
                    End If
                Else
                    Me.cboMtpRevno.DataSource = Nothing
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function chkBudgetOrderSelected() As Boolean
        Dim blnSelected As Boolean = False
        Dim strPeriodType As String = String.Empty

        Try
            Me.grvBudget1.EndEdit()
            Me.grvBudget2.EndEdit()
            Me.grvBudget3.EndEdit()
            Me.grvBudget4.EndEdit()

            strPeriodType = Me.GetPeriodType

            Select Case strPeriodType
                Case CStr(enumPeriodType.OriginalBudget)  '// OB
                    For Each row As DataGridViewRow In grvBudget1.Rows
                        blnSelected = DirectCast(grvBudget1(0, row.Index).Value, Boolean)
                        If blnSelected = True Then
                            Exit For
                        End If
                    Next
                Case CStr(enumPeriodType.EstimateBudget)  '// EB
                    For Each row As DataGridViewRow In grvBudget2.Rows
                        blnSelected = DirectCast(grvBudget2(0, row.Index).Value, Boolean)
                        If blnSelected = True Then
                            Exit For
                        End If
                    Next
                Case CStr(enumPeriodType.ReviseBudget)  '// RB
                    For Each row As DataGridViewRow In grvBudget3.Rows
                        blnSelected = DirectCast(grvBudget3(0, row.Index).Value, Boolean)
                        If blnSelected = True Then
                            Exit For
                        End If
                    Next
                Case CStr(enumPeriodType.MTPBudget)  '// MTP
                    For Each row As DataGridViewRow In grvBudget4.Rows
                        blnSelected = DirectCast(grvBudget4(0, row.Index).Value, Boolean)
                        If blnSelected = True Then
                            Exit For
                        End If
                    Next
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return blnSelected
    End Function

    Private Sub SetPropertiesDataGridView()
        Dim font As Font
        Dim style As DataGridViewCellStyle = New DataGridViewCellStyle()

        Try
            'Set "Original" Header Font Bold 
            font = New Font(grvBudget1.DefaultCellStyle.Font.FontFamily, 8, FontStyle.Bold)
            style.Font = font
            For Each column As DataGridViewColumn In grvBudget1.Columns
                If (column.Name = "g1col6" Or _
                    column.Name = "g1col7" Or _
                    column.Name = "g1col14" Or _
                    column.Name = "g1col15" Or _
                    column.Name = "g1col16" Or _
                    column.Name = "g1col17" Or _
                    column.Name = "g1col25") Then
                    column.HeaderCell.Style = style
                    column.Width = 80
                End If
            Next
            'Set Columns Font Bold
            grvBudget1.Columns("g1col6").DefaultCellStyle.Font = font
            grvBudget1.Columns("g1col7").DefaultCellStyle.Font = font
            grvBudget1.Columns("g1col14").DefaultCellStyle.Font = font
            grvBudget1.Columns("g1col15").DefaultCellStyle.Font = font
            grvBudget1.Columns("g1col16").DefaultCellStyle.Font = font
            grvBudget1.Columns("g1col17").DefaultCellStyle.Font = font
            grvBudget1.Columns("g1col25").DefaultCellStyle.Font = font

            'Set "Estimate" Header Font Bold 
            font = New Font(grvBudget2.DefaultCellStyle.Font.FontFamily, 8, FontStyle.Bold)
            style.Font = font
            For Each column As DataGridViewColumn In grvBudget2.Columns
                If (column.Name = "g2col6" Or _
                    column.Name = "g2col7" Or _
                    column.Name = "g2col14" Or _
                    column.Name = "g2col16") Then
                    column.HeaderCell.Style = style
                    column.Width = 80
                End If
            Next
            'Set Columns Font Bold
            grvBudget2.Columns("g2col6").DefaultCellStyle.Font = font
            grvBudget2.Columns("g2col7").DefaultCellStyle.Font = font
            grvBudget2.Columns("g2col14").DefaultCellStyle.Font = font
            grvBudget2.Columns("g2col16").DefaultCellStyle.Font = font

            'Set "Revise" Header Font Bold 
            font = New Font(grvBudget3.DefaultCellStyle.Font.FontFamily, 8, FontStyle.Bold)
            style.Font = font
            For Each column As DataGridViewColumn In grvBudget3.Columns
                If (column.Name = "g3col6" Or _
                    column.Name = "g3col13" Or _
                    column.Name = "g3col15" Or _
                    column.Name = "g3col22" Or _
                    column.Name = "g3col24") Then
                    column.HeaderCell.Style = style
                    column.Width = 80
                End If
            Next
            'Set Columns Font Bold
            grvBudget3.Columns("g3col6").DefaultCellStyle.Font = font
            grvBudget3.Columns("g3col13").DefaultCellStyle.Font = font
            grvBudget3.Columns("g3col15").DefaultCellStyle.Font = font
            grvBudget3.Columns("g3col22").DefaultCellStyle.Font = font
            grvBudget3.Columns("g3col24").DefaultCellStyle.Font = font

            'Set Header Font Bold MTP
            font = New Font(grvBudget4.DefaultCellStyle.Font.FontFamily, 8, FontStyle.Bold)
            style.Font = font
            For Each column As DataGridViewColumn In grvBudget4.Columns
                If (column.Name = "g4col6" Or _
                    column.Name = "g4col9" Or _
                    column.Name = "g4col11" Or _
                    column.Name = "g4col13" Or _
                    column.Name = "g4col15" Or _
                    column.Name = "g4col17") Then
                    column.HeaderCell.Style = style
                    column.Width = 80
                End If
            Next
            'Set Columns Font Bold
            grvBudget4.Columns("g4col6").DefaultCellStyle.Font = font
            grvBudget4.Columns("g4col9").DefaultCellStyle.Font = font
            grvBudget4.Columns("g4col11").DefaultCellStyle.Font = font
            grvBudget4.Columns("g4col13").DefaultCellStyle.Font = font
            grvBudget4.Columns("g4col15").DefaultCellStyle.Font = font
            grvBudget4.Columns("g4col17").DefaultCellStyle.Font = font
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Control Event"

    Private Sub frmBG0200_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            p_frmBG0200.RemoveAll(Function(f) f.BudgetKey = Me.BudgetKey And f.OperationCd = Me.OperationCd)
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmBG0200_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If (myForceCloseFlg = False And myBudgetDataChanged = True) AndAlso _
                    MessageBox.Show("Some changed data may not be saved." & vbNewLine & "Are you sure to close?", Me.Text, MessageBoxButtons.YesNo, _
                                   MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmBG0200_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            myFormLoadedFlg = False

            '// Load Config
            myClsBG0200BL.GetConfig()

            '// Load PIC ComboBox
            LoadPICList()

            '// Set control up to operation mode
            Select Case Me.OperationCd
                Case enumOperationCd.ViewBudget
                    Me.lblFormTitle.Text = "View Budget Journal"
                    Me.cmdApprove.Visible = False
                    Me.cmdReject.Visible = False
                    Me.cmdSubmit.Visible = False
                    Me.cmdSave.Visible = False
                    Me.cmdAuth1.Visible = False
                    Me.cmdAuth2.Visible = False
                    Me.cmdUpRev.Visible = False
                    Me.cmdDelRev.Visible = False
                    Me.cmdAdjust.Visible = False
                    Me.cmdSubmit2.Visible = False
                    Me.cmdReInput.Visible = False
                    Me.cmdReInputByOrder.Visible = False
                    Me.fraWk.Visible = False
                    Me.cboRevNo.Visible = False
                    Me.pnlWKH.Visible = False
                    Me.pnlMTP_SUM.Visible = False
                    Me.fraWBMTP.Visible = False
                    Me.pnlMTPInvestment.Visible = False
                    Me.grpMTPWB.Visible = False

                Case enumOperationCd.InputBudget
                    Me.lblFormTitle.Text = "Input Budget Journal"
                    Me.cmdApprove.Visible = False
                    Me.cmdReject.Visible = False
                    Me.cmdSubmit.Visible = True
                    Me.cmdSave.Visible = True
                    Me.cmdAuth1.Visible = False
                    Me.cmdAuth2.Visible = False
                    Me.cmdUpRev.Visible = False
                    Me.cmdDelRev.Visible = False
                    Me.cmdAdjust.Visible = False
                    Me.cmdSubmit2.Visible = False
                    Me.cmdReInput.Visible = False
                    Me.cmdReInputByOrder.Visible = False
                    Me.fraWk.Visible = False
                    Me.cboRevNo.Visible = False
                    Me.pnlWKH.Visible = False
                    Me.pnlMTP_SUM.Visible = False
                    Me.fraWBMTP.Visible = False
                    Me.pnlMTPInvestment.Visible = False
                    Me.grpMTPWB.Visible = False

                Case enumOperationCd.ApproveBudget
                    Me.lblFormTitle.Text = "Approve Budget Journal"
                    Me.cmdApprove.Visible = True
                    Me.cmdReject.Visible = True
                    Me.cmdSubmit.Visible = False
                    Me.cmdSave.Visible = False
                    Me.cmdAuth1.Visible = False
                    Me.cmdAuth2.Visible = False
                    Me.cmdUpRev.Visible = False
                    Me.cmdDelRev.Visible = False
                    Me.cmdAdjust.Visible = False
                    Me.cmdSubmit2.Visible = False
                    Me.cmdReInput.Visible = False
                    Me.cmdReInputByOrder.Visible = False
                    Me.fraWk.Visible = False
                    Me.cboRevNo.Visible = False
                    Me.pnlWKH.Visible = False
                    Me.pnlMTP_SUM.Visible = False
                    Me.fraWBMTP.Visible = False
                    Me.pnlMTPInvestment.Visible = False
                    Me.grpMTPWB.Visible = False

                Case enumOperationCd.AdjustBudget
                    Me.lblFormTitle.Text = "Adjust Budget Journal"
                    Me.cmdApprove.Visible = False
                    Me.cmdReject.Visible = True 'False
                    Me.cmdSubmit.Visible = False
                    Me.cmdSave.Visible = True
                    Me.cmdAuth1.Visible = False
                    Me.cmdAuth2.Visible = False
                    Me.cmdUpRev.Visible = True
                    Me.cmdDelRev.Visible = True
                    Me.cmdAdjust.Visible = True
                    Me.cmdSubmit2.Visible = True
                    Me.cmdReInput.Visible = True
                    Me.cmdReInputByOrder.Visible = True
                    Me.fraWk.Visible = True
                    If Me.GetBudgetType() = P_BUDGET_TYPE_EXPENSE Then
                        Me.pnlWKH.Visible = True
                        Me.pnlMTP_SUM.Visible = False
                    Else
                        Me.pnlWKH.Visible = False
                        Me.pnlMTP_SUM.Visible = False
                    End If
                    If Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then
                        Me.fraWBMTP.Visible = True
                        Me.pnlMTPInvestment.Visible = True
                        Me.pnlWKH.Visible = False
                    Else
                        Me.fraWBMTP.Visible = False
                        Me.pnlMTPInvestment.Visible = False
                    End If
                    If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then
                        Me.grpMTPWB.Visible = True
                    Else
                        Me.grpMTPWB.Visible = False
                    End If

                Case enumOperationCd.AdjustBudgetDirectInput
                    Me.lblFormTitle.Text = "Adjust Budget Journal (Direct Input)"
                    Me.cmdApprove.Visible = False
                    Me.cmdReject.Visible = False
                    Me.cmdSubmit.Visible = False
                    Me.cmdSave.Visible = True
                    Me.cmdAuth1.Visible = False
                    Me.cmdAuth2.Visible = False
                    Me.cmdUpRev.Visible = False
                    Me.cmdDelRev.Visible = False
                    Me.cmdAdjust.Visible = False
                    Me.cmdSubmit2.Visible = False
                    Me.cmdReInput.Visible = False
                    Me.cmdReInputByOrder.Visible = False
                    Me.fraWk.Visible = True
                    If Me.GetBudgetType() = P_BUDGET_TYPE_EXPENSE Then
                        Me.pnlWKH.Visible = True
                        Me.pnlMTP_SUM.Visible = False
                    Else
                        Me.pnlWKH.Visible = False
                        Me.pnlMTP_SUM.Visible = False
                    End If
                    If Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then
                        Me.fraWBMTP.Visible = True
                        Me.pnlMTPInvestment.Visible = True
                        Me.pnlWKH.Visible = False
                    Else
                        Me.fraWBMTP.Visible = False
                        Me.pnlMTPInvestment.Visible = False
                    End If
                    If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then
                        Me.grpMTPWB.Visible = True
                    Else
                        Me.grpMTPWB.Visible = False
                    End If

                Case enumOperationCd.Authorize1
                    Me.lblFormTitle.Text = "Authorize Budget Journal (AD)"
                    Me.cmdApprove.Visible = False
                    Me.cmdReject.Visible = True
                    Me.cmdSubmit.Visible = False
                    Me.cmdSave.Visible = False
                    Me.cmdAuth1.Visible = True
                    Me.cmdAuth2.Visible = False
                    Me.cmdUpRev.Visible = False
                    Me.cmdDelRev.Visible = False
                    Me.cmdAdjust.Visible = False
                    Me.cmdSubmit2.Visible = False
                    Me.cmdReInput.Visible = False
                    Me.cmdReInputByOrder.Visible = False
                    Me.fraWk.Visible = True
                    Me.cboRevNo.Visible = False
                    Me.pnlWKH.Visible = False
                    Me.pnlMTP_SUM.Visible = False
                    Me.fraWBMTP.Visible = False
                    Me.pnlMTPInvestment.Visible = False
                    Me.grpMTPWB.Visible = False

                Case enumOperationCd.Authorize2
                    Me.lblFormTitle.Text = "Authorize Budget Journal (MD)"
                    Me.cmdApprove.Visible = False
                    Me.cmdReject.Visible = True
                    Me.cmdSubmit.Visible = False
                    Me.cmdSave.Visible = False
                    Me.cmdAuth1.Visible = False
                    Me.cmdAuth2.Visible = True
                    Me.cmdUpRev.Visible = False
                    Me.cmdDelRev.Visible = False
                    Me.cmdAdjust.Visible = False
                    Me.cmdSubmit2.Visible = False
                    Me.cmdReInput.Visible = False
                    Me.cmdReInputByOrder.Visible = False
                    Me.fraWk.Visible = True
                    Me.cboRevNo.Visible = False
                    Me.pnlWKH.Visible = False
                    Me.pnlMTP_SUM.Visible = False
                    Me.fraWBMTP.Visible = False
                    Me.pnlMTPInvestment.Visible = False
                    Me.grpMTPWB.Visible = False
            End Select

            Me.gbOB_MTP.Visible = False
            If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) AndAlso _
                Me.GetBudgetType() = P_BUDGET_TYPE_EXPENSE Then
            Else
                Me.gbOB_MTP.Visible = False
            End If
            If Me.cboRevNo.Visible = True Then
                LoadRevNoList()
            End If

            myFormLoadedFlg = True
            Timer1.Enabled = True
            Me.SetPropertiesDataGridView()
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cboPIC_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPIC.SelectedIndexChanged
        Try
            If myFormLoadedFlg = True And cboPIC.SelectedIndex >= 0 Then
                ShowBudgetData()
                mydtBG1 = Nothing
                mydtBG2 = Nothing
                mydtBG3 = Nothing
                mydtBG4 = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cboAccount_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAccount.SelectedIndexChanged
        Try
            FilterGridView()
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cboCost_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCost.SelectedIndexChanged
        Try
            FilterGridView()
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cboCostType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCostType.SelectedIndexChanged
        Try
            FilterGridView()
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cboDept_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDept.SelectedIndexChanged
        FilterGridView()
    End Sub

    Private Sub cboRevNo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRevNo.SelectedIndexChanged
        Try
            If cboRevNo.Text = Me.CurrRevNo Then
                Exit Sub
            End If

            If myFormLoadedFlg = True And myShowWarning = True Then

                If MessageBox.Show("Some changed data may not be saved." & vbNewLine & "Are you sure to change Rev No.?", Me.Text, MessageBoxButtons.YesNo, _
                MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
                    ReloadPICList(cboRevNo.Text)
                Else
                    myFormLoadedFlg = False
                    cboRevNo.SelectedValue = Me.CurrRevNo
                    myFormLoadedFlg = True
                End If

            Else
                ReloadPICList(cboRevNo.Text)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cboMtpRevno_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboMtpRevno.SelectedIndexChanged
        Try
            If myFormLoadedFlg = True And cboPIC.SelectedIndex >= 0 Then
                ShowBudgetData()
                mydtBG1 = Nothing
                mydtBG2 = Nothing
                mydtBG3 = Nothing
                mydtBG4 = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub numMtpProjectNo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numMtpProjectNo.ValueChanged
        Try
            LoadMtpRevNo()
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub grvBudget1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grvBudget1.CellValueChanged
        Try
            If e.RowIndex >= 0 And mySetGridValue = False Then
                If grvBudget1.Columns(e.ColumnIndex).Name <> "g1Wk" Then
                    '// Validate input data
                    CheckValidateOriginalBudget(e.ColumnIndex, e.RowIndex)

                    CalSum()
                End If
            End If
            myBudgetDataChanged = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub grvBudget1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles grvBudget1.DataError
        Try
            e.ThrowException = False
            MessageBox.Show("the value must be a number", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub grvBudget1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles grvBudget1.KeyDown
        Try
            If e.Control AndAlso e.KeyCode = Keys.V Then
                Try
                    Dim s As String = Clipboard.GetText()
                    Dim lines() As String = s.Split(New String() {vbLf}, StringSplitOptions.None)
                    Dim iFail As Integer = 0
                    Dim iRow As Integer = grvBudget1.CurrentCell.RowIndex
                    Dim iCol As Integer = grvBudget1.CurrentCell.ColumnIndex
                    Dim oCell As DataGridViewCell
                    Dim i As Integer

                    For Each line As String In lines
                        If iRow < grvBudget1.RowCount AndAlso line.Length > 0 Then
                            Dim sCells() As String = line.Split(New String() {vbTab}, StringSplitOptions.None)
                            For i = 0 To sCells.GetLength(0) - 1
                                If iCol + i < Me.grvBudget1.ColumnCount Then
                                    oCell = grvBudget1(iCol + i, iRow)
                                    If Not oCell.ReadOnly Then
                                        If oCell.Value.ToString() <> sCells(i) Then
                                            oCell.Value = Convert.ChangeType(sCells(i), oCell.ValueType)
                                        End If
                                    Else
                                        iFail = iFail + 1
                                    End If
                                Else
                                    Exit For
                                End If
                            Next
                            iRow = iRow + 1
                        Else
                            Exit For
                        End If

                        If (iFail > 0) Then
                            MessageBox.Show(String.Format("{0} updates failed due to read only column setting", iFail))
                        End If
                    Next
                Catch ex As FormatException
                    MessageBox.Show("The data you pasted is in the wrong format for the cell")
                End Try
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub grvBudget1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grvBudget1.Sorted
        Try
            Debug.Print(Now.ToString() & ": Begin grvBudget1_Sorted")
            myDataLoadingFlg = True

            If CDbl(Me.myCurrRevNo) > 1 And _
            (Me.OperationCd = enumOperationCd.AdjustBudget Or _
             Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput Or _
             Me.OperationCd = enumOperationCd.Authorize1 Or _
             Me.OperationCd = enumOperationCd.Authorize2) Then
                '// Highlight Working Budget
                HighlightWorkingBG()

                '// Highlight Transfer Cost
                HighlightTransferCost()
            End If

            myDataLoadingFlg = False
            '// Filter Row
            FilterGridView()

            HighlightWorkingBGAndComment()

            Debug.Print(Now.ToString() & ": End grvBudget1_Sorted")
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub grvBudget2_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grvBudget2.CellValueChanged
        Try
            If e.RowIndex >= 0 And mySetGridValue = False Then
                If grvBudget2.Columns(e.ColumnIndex).Name <> "g2Wk" Then
                    '// Validate input data
                    CheckValidateEstimateBudget(e.ColumnIndex, e.RowIndex)

                    CalSum()
                End If
            End If
            myBudgetDataChanged = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub grvBudget2_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles grvBudget2.DataError
        Try
            e.ThrowException = False
            MessageBox.Show("the value must be a number", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub grvBudget2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles grvBudget2.KeyDown
        Try
            If e.Control AndAlso e.KeyCode = Keys.V Then
                Try
                    Dim s As String = Clipboard.GetText()
                    Dim lines() As String = s.Split(New String() {vbLf}, StringSplitOptions.None)
                    Dim iFail As Integer = 0
                    Dim iRow As Integer = grvBudget2.CurrentCell.RowIndex
                    Dim iCol As Integer = grvBudget2.CurrentCell.ColumnIndex
                    Dim oCell As DataGridViewCell
                    Dim i As Integer

                    For Each line As String In lines
                        If iRow < grvBudget2.RowCount AndAlso line.Length > 0 Then
                            Dim sCells() As String = line.Split(New String() {vbTab}, StringSplitOptions.None)
                            For i = 0 To sCells.GetLength(0) - 1
                                If iCol + i < Me.grvBudget2.ColumnCount Then
                                    oCell = grvBudget2(iCol + i, iRow)
                                    If Not oCell.ReadOnly Then
                                        If oCell.Value.ToString() <> sCells(i) Then
                                            oCell.Value = Convert.ChangeType(sCells(i), oCell.ValueType)
                                        End If
                                    Else
                                        iFail = iFail + 1
                                    End If
                                Else
                                    Exit For
                                End If
                            Next
                            iRow = iRow + 1
                        Else
                            Exit For
                        End If

                        If (iFail > 0) Then
                            MessageBox.Show(String.Format("{0} updates failed due to read only column setting", iFail))
                        End If
                    Next

                Catch ex As FormatException
                    MessageBox.Show("The data you pasted is in the wrong format for the cell")
                End Try
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub grvBudget2_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grvBudget2.Sorted
        Try
            Debug.Print(Now.ToString() & ": Begin grvBudget2_Sorted")
            myDataLoadingFlg = True

            If CDbl(Me.myCurrRevNo) > 1 And _
            (Me.OperationCd = enumOperationCd.AdjustBudget Or _
             Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput Or _
             Me.OperationCd = enumOperationCd.Authorize1 Or _
             Me.OperationCd = enumOperationCd.Authorize2) Then
                '// Highlight Working Budget
                HighlightWorkingBG()

                '// Highlight Transfer Cost
                HighlightTransferCost()
            End If

            myDataLoadingFlg = False
            '// Filter Row
            FilterGridView()
            Debug.Print(Now.ToString() & ": End grvBudget2_Sorted")
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub grvBudget3_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grvBudget3.CellValueChanged
        Try
            If e.RowIndex >= 0 And mySetGridValue = False Then
                If grvBudget3.Columns(e.ColumnIndex).Name <> "g3Wk" Then
                    '// Validate input data
                    CheckValidateReviseBudget(e.ColumnIndex, e.RowIndex)

                    CalSum()
                End If
            End If
            myBudgetDataChanged = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub grvBudget3_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles grvBudget3.DataError
        Try
            e.ThrowException = False
            MessageBox.Show("the value must be a number", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub grvBudget3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles grvBudget3.KeyDown
        Try
            If e.Control AndAlso e.KeyCode = Keys.V Then
                Try
                    Dim s As String = Clipboard.GetText()
                    Dim lines() As String = s.Split(New String() {vbLf}, StringSplitOptions.None)
                    Dim iFail As Integer = 0
                    Dim iRow As Integer = grvBudget3.CurrentCell.RowIndex
                    Dim iCol As Integer = grvBudget3.CurrentCell.ColumnIndex
                    Dim oCell As DataGridViewCell
                    Dim i As Integer

                    For Each line As String In lines
                        If iRow < grvBudget3.RowCount AndAlso line.Length > 0 Then
                            Dim sCells() As String = line.Split(New String() {vbTab}, StringSplitOptions.None)
                            For i = 0 To sCells.GetLength(0) - 1
                                If iCol + i < Me.grvBudget3.ColumnCount Then
                                    oCell = grvBudget3(iCol + i, iRow)
                                    If Not oCell.ReadOnly Then
                                        If oCell.Value.ToString() <> sCells(i) Then
                                            oCell.Value = Convert.ChangeType(sCells(i), oCell.ValueType)
                                        End If
                                    Else
                                        iFail = iFail + 1
                                    End If
                                Else
                                    Exit For
                                End If
                            Next
                            iRow = iRow + 1
                        Else
                            Exit For
                        End If

                        If (iFail > 0) Then
                            MessageBox.Show(String.Format("{0} updates failed due to read only column setting", iFail))
                        End If
                    Next
                Catch ex As FormatException
                    MessageBox.Show("The data you pasted is in the wrong format for the cell")
                End Try
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub grvBudget3_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grvBudget3.Sorted
        Try
            Debug.Print(Now.ToString() & ": Begin grvBudget3_Sorted")
            myDataLoadingFlg = True

            If CDbl(Me.myCurrRevNo) > 1 And _
            (Me.OperationCd = enumOperationCd.AdjustBudget Or _
             Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput Or _
             Me.OperationCd = enumOperationCd.Authorize1 Or _
             Me.OperationCd = enumOperationCd.Authorize2) Then
                '// Highlight Working Budget
                HighlightWorkingBG()

                '// Highlight Transfer Cost
                HighlightTransferCost()
            End If

            myDataLoadingFlg = False
            '// Filter Row
            FilterGridView()
            Debug.Print(Now.ToString() & ": End grvBudget3_Sorted")
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub grvBudget4_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grvBudget4.CellValueChanged
        Try
            If e.RowIndex > 0 And mySetGridValue = False Then
                If grvBudget4.Columns(e.ColumnIndex).Name <> "g4Wk" Then
                    '// Validate input data
                    CheckValidateMTPBudget(e.ColumnIndex, e.RowIndex)

                    '// Highlight Over MTP value
                    If Me.GetBudgetType() = P_BUDGET_TYPE_EXPENSE Then
                        HighlightMTPValueNew(e.ColumnIndex, e.RowIndex)
                    End If
                    CalSum()
                End If
            End If
            myBudgetDataChanged = True

        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub grvBudget4_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles grvBudget4.DataError
        Try
            e.ThrowException = False
            MessageBox.Show("the value must be a number", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub grvBudget4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles grvBudget4.KeyDown
        Try
            If e.Control AndAlso e.KeyCode = Keys.V Then
                Try
                    Dim s As String = Clipboard.GetText()
                    Dim lines() As String = s.Split(New String() {vbLf}, StringSplitOptions.None)
                    Dim iFail As Integer = 0
                    Dim iRow As Integer = grvBudget4.CurrentCell.RowIndex
                    Dim iCol As Integer = grvBudget4.CurrentCell.ColumnIndex
                    Dim oCell As DataGridViewCell
                    Dim i As Integer

                    For Each line As String In lines
                        If iRow < grvBudget4.RowCount AndAlso line.Length > 0 Then
                            Dim sCells() As String = line.Split(New String() {vbTab}, StringSplitOptions.None)
                            For i = 0 To sCells.GetLength(0) - 1
                                If iCol + i < Me.grvBudget4.ColumnCount Then
                                    oCell = grvBudget4(iCol + i, iRow)
                                    If Not oCell.ReadOnly Then
                                        If oCell.Value.ToString() <> sCells(i) Then
                                            oCell.Value = Convert.ChangeType(sCells(i), oCell.ValueType)
                                        End If
                                    Else
                                        iFail = iFail + 1
                                    End If

                                Else
                                    Exit For
                                End If
                            Next
                            iRow = iRow + 1
                        Else
                            Exit For
                        End If

                        If (iFail > 0) Then
                            MessageBox.Show(String.Format("{0} updates failed due to read only column setting", iFail))
                        End If
                    Next
                Catch ex As FormatException
                    MessageBox.Show("The data you pasted is in the wrong format for the cell")
                End Try
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub grvBudget4_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grvBudget4.Sorted
        Try
            Debug.Print(Now.ToString() & ": Begin grvBudget4_Sorted")
            myDataLoadingFlg = True

            If CDbl(Me.myCurrRevNo) > 1 And _
            (Me.OperationCd = enumOperationCd.AdjustBudget Or _
             Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput Or _
             Me.OperationCd = enumOperationCd.Authorize1 Or _
             Me.OperationCd = enumOperationCd.Authorize2) Then
            End If

            '// Highlight Over MTP value
            If Me.GetBudgetType() = P_BUDGET_TYPE_EXPENSE Then
                HighlightMTPValueAllNew()
            End If

            myDataLoadingFlg = False
            '// Filter Row
            FilterGridView()
            Debug.Print(Now.ToString() & ": End grvBudget4_Sorted")
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtRRT0_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRRT0.KeyPress, txtRRT1.KeyPress, txtRRT2.KeyPress, _
                                                                                                                    txtRRT3.KeyPress, txtRRT4.KeyPress, txtRRT5.KeyPress, _
                                                                                                                    txtWk1.KeyPress, txtWk2.KeyPress, _
                                                                                                                    txtWKRRT1.KeyPress, txtWKRRT2.KeyPress, _
                                                                                                                    txtWKRRT3.KeyPress, txtWKRRT4.KeyPress, txtWKRRT5.KeyPress, _
                                                                                                                    txtMTPInv1.KeyPress, txtMTPInv2.KeyPress, txtMTPInv3.KeyPress, txtMTPInv4.KeyPress, txtMTPInv5.KeyPress, _
                                                                                                                    txtPYInv1.KeyPress, txtPYInv2.KeyPress, txtPYInv3.KeyPress, txtPYInv4.KeyPress, txtPYInv5.KeyPress
        Try
            '// Set textbox accept number only
            If IsNumeric(e.KeyChar) Or Asc(e.KeyChar) = Keys.Back Or Asc(e.KeyChar) = Keys.Delete Or e.KeyChar = CChar(".") Then
                e.Handled = False
            Else
                e.Handled = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtRRT0_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRRT0.TextChanged
        Try
            If IsNumeric(txtRRT0.Text) AndAlso CDbl(txtRRT0.Text) > 0 Then
                If IsNumeric(txtRRT1.Text) AndAlso CDbl(txtRRT1.Text) > 0 Then
                    lblRRT1p.Text = ((CDbl(txtRRT1.Text) / CDbl(txtRRT0.Text)) * 100).ToString("0") & "%"
                Else
                    lblRRT1p.Text = ""
                End If
                If IsNumeric(txtRRT2.Text) AndAlso CDbl(txtRRT2.Text) > 0 Then
                    lblRRT2p.Text = ((CDbl(txtRRT2.Text) / CDbl(txtRRT0.Text)) * 100).ToString("0") & "%"
                Else
                    lblRRT2p.Text = ""
                End If
                If IsNumeric(txtRRT3.Text) AndAlso CDbl(txtRRT3.Text) > 0 Then
                    lblRRT3p.Text = ((CDbl(txtRRT3.Text) / CDbl(txtRRT0.Text)) * 100).ToString("0") & "%"
                Else
                    lblRRT3p.Text = ""
                End If
                If IsNumeric(txtRRT4.Text) AndAlso CDbl(txtRRT4.Text) > 0 Then
                    lblRRT4p.Text = ((CDbl(txtRRT4.Text) / CDbl(txtRRT0.Text)) * 100).ToString("0") & "%"
                Else
                    lblRRT4p.Text = ""
                End If
                If IsNumeric(txtRRT5.Text) AndAlso CDbl(txtRRT5.Text) > 0 Then
                    lblRRT5p.Text = ((CDbl(txtRRT5.Text) / CDbl(txtRRT0.Text)) * 100).ToString("0") & "%"
                Else
                    lblRRT5p.Text = ""
                End If
            Else
                lblRRT1p.Text = ""
                lblRRT2p.Text = ""
                lblRRT3p.Text = ""
                lblRRT4p.Text = ""
                lblRRT5p.Text = ""
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtRRT1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRRT1.TextChanged
        Try
            If IsNumeric(txtRRT1.Text) AndAlso IsNumeric(txtRRT0.Text) AndAlso CDbl(txtRRT0.Text) > 0 AndAlso CDbl(txtRRT1.Text) > 0 Then
                lblRRT1p.Text = ((CDbl(txtRRT1.Text) / CDbl(txtRRT0.Text)) * 100).ToString("0") & "%"
            Else
                lblRRT1p.Text = ""
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtRRT2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRRT2.TextChanged
        Try
            If IsNumeric(txtRRT2.Text) AndAlso IsNumeric(txtRRT0.Text) AndAlso CDbl(txtRRT0.Text) > 0 AndAlso CDbl(txtRRT2.Text) > 0 Then
                lblRRT2p.Text = ((CDbl(txtRRT2.Text) / CDbl(txtRRT0.Text)) * 100).ToString("0") & "%"
            Else
                lblRRT2p.Text = ""
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtRRT3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRRT3.TextChanged
        Try
            If IsNumeric(txtRRT3.Text) AndAlso IsNumeric(txtRRT0.Text) AndAlso CDbl(txtRRT0.Text) > 0 AndAlso CDbl(txtRRT3.Text) > 0 Then
                lblRRT3p.Text = ((CDbl(txtRRT3.Text) / CDbl(txtRRT0.Text)) * 100).ToString("0") & "%"
            Else
                lblRRT3p.Text = ""
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtRRT4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRRT4.TextChanged
        Try
            If IsNumeric(txtRRT4.Text) AndAlso IsNumeric(txtRRT0.Text) AndAlso CDbl(txtRRT0.Text) > 0 AndAlso CDbl(txtRRT4.Text) > 0 Then
                lblRRT4p.Text = ((CDbl(txtRRT4.Text) / CDbl(txtRRT0.Text)) * 100).ToString("0") & "%"
            Else
                lblRRT4p.Text = ""
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtRRT5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRRT5.TextChanged
        Try
            If IsNumeric(txtRRT5.Text) AndAlso IsNumeric(txtRRT0.Text) AndAlso CDbl(txtRRT0.Text) > 0 AndAlso CDbl(txtRRT5.Text) > 0 Then
                lblRRT5p.Text = ((CDbl(txtRRT5.Text) / CDbl(txtRRT0.Text)) * 100).ToString("0") & "%"
            Else
                lblRRT5p.Text = ""
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtWKH1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtWKH1.KeyPress, txtWKH2.KeyPress, _
                                                                                                                    txtMTP_SUM1.KeyPress, txtMTP_SUM2.KeyPress, _
                                                                                                                    txtMTP_SUM3.KeyPress, txtMTP_SUM4.KeyPress, _
                                                                                                                    txtMTP_SUM5.KeyPress, _
                                                                                                                    txtMTPWB.KeyPress
        Try
            '// Set textbox accept number only
            If IsNumeric(e.KeyChar) Or Asc(e.KeyChar) = Keys.Back Or Asc(e.KeyChar) = Keys.Delete Or _
            e.KeyChar = CChar(".") Or e.KeyChar = CChar("-") Then
                e.Handled = False
            Else
                e.Handled = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim strTmp As String

        Try
            If MessageBox.Show("Are you sure to save this budget journal?", Me.Text, MessageBoxButtons.YesNo, _
                                         MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            Me.grvBudget1.EndEdit()
            Me.grvBudget2.EndEdit()
            Me.grvBudget3.EndEdit()
            Me.grvBudget4.EndEdit()

            strTmp = txtWKH1.Text
            If strTmp <> "" Then
                txtWKH1.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If
            strTmp = txtWKH2.Text
            If strTmp <> "" Then
                txtWKH2.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If
            strTmp = txtMTP_SUM1.Text
            If strTmp <> "" Then
                txtMTP_SUM1.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If
            strTmp = txtMTP_SUM2.Text
            If strTmp <> "" Then
                txtMTP_SUM2.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If
            strTmp = txtMTP_SUM3.Text
            If strTmp <> "" Then
                txtMTP_SUM3.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If
            strTmp = txtMTP_SUM4.Text
            If strTmp <> "" Then
                txtMTP_SUM4.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If
            strTmp = txtMTP_SUM5.Text
            If strTmp <> "" Then
                txtMTP_SUM5.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If
            strTmp = txtWKRRT1.Text
            If strTmp <> "" Then
                txtWKRRT1.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If
            strTmp = txtWKRRT2.Text
            If strTmp <> "" Then
                txtWKRRT2.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If
            strTmp = txtWKRRT3.Text
            If strTmp <> "" Then
                txtWKRRT3.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If
            strTmp = txtWKRRT4.Text
            If strTmp <> "" Then
                txtWKRRT4.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If
            strTmp = txtWKRRT5.Text
            If strTmp <> "" Then
                txtWKRRT5.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If
            strTmp = txtMTPInv1.Text
            If strTmp <> "" Then
                txtMTPInv1.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If
            strTmp = txtMTPInv2.Text
            If strTmp <> "" Then
                txtMTPInv2.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If
            strTmp = txtMTPInv3.Text
            If strTmp <> "" Then
                txtMTPInv3.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If
            strTmp = txtMTPInv4.Text
            If strTmp <> "" Then
                txtMTPInv4.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If
            strTmp = txtMTPInv5.Text
            If strTmp <> "" Then
                txtMTPInv5.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If
            strTmp = txtPYInv1.Text
            If strTmp <> "" Then
                txtPYInv1.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If
            strTmp = txtPYInv2.Text
            If strTmp <> "" Then
                txtPYInv2.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If
            strTmp = txtPYInv3.Text
            If strTmp <> "" Then
                txtPYInv3.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If
            strTmp = txtPYInv4.Text
            If strTmp <> "" Then
                txtPYInv4.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If
            strTmp = txtPYInv5.Text
            If strTmp <> "" Then
                txtPYInv5.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If
            strTmp = txtMTPWB.Text
            If strTmp <> "" Then
                txtMTPWB.Text = CDbl(strTmp.Replace(",", "")).ToString("#,##0.00")
            End If

            '// Save budget journal
            '// Set Parameters
            myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
            myClsBG0200BL.PeriodType = Me.GetPeriodType()
            myClsBG0200BL.BudgetType = Me.GetBudgetType()
            myClsBG0200BL.UserId = p_strUserId
            myClsBG0200BL.UserPIC = CStr(cboPIC.SelectedValue)
            If Me.myOperationCd = enumOperationCd.AdjustBudget Or Me.myOperationCd = enumOperationCd.AdjustBudgetDirectInput Then
                myClsBG0200BL.Status = CStr(enumBudgetStatus.Approve)
            Else
                myClsBG0200BL.Status = CStr(enumBudgetStatus.NewRecord)
            End If
            myClsBG0200BL.RevNo = Me.CurrRevNo
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then          '// Original Budget
                myClsBG0200BL.BudgetList = CType(grvBudget1.DataSource, DataTable)

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then      '// Estimate Budget
                myClsBG0200BL.BudgetList = CType(grvBudget2.DataSource, DataTable)

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then        '// Revise Budget
                myClsBG0200BL.BudgetList = CType(grvBudget3.DataSource, DataTable)
            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then        '// MTP Budget
                myClsBG0200BL.BudgetList = CType(grvBudget4.DataSource, DataTable)
            End If

            If Me.CurrRevNo = "1" And Me.OperationCd = enumOperationCd.AdjustBudget Then
                Dim strDispRevNo As String
                strDispRevNo = Me.CurrRevNo

                '// Get Max rev
                GetCurrentRevNo()

                '// Check if max rev is not rev.1, then discard save.
                If strDispRevNo <> Me.CurrRevNo Then
                    '// Restore rev no
                    Me.CurrRevNo = strDispRevNo
                    MessageBox.Show("Can not save Budget journal", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                '// Transfer Cost
                myClsBG0200BL.AdjustTransferCost()
                LoadRevNoList()

                If Me.GetBudgetType() = P_BUDGET_TYPE_EXPENSE Then
                    '// Save WKH
                    myClsBG0200BL.WKH1 = CStr(IIf(IsNumeric(txtWKH1.Text), txtWKH1.Text, "0")).Replace(",", "")
                    myClsBG0200BL.WKH2 = CStr(IIf(IsNumeric(txtWKH2.Text), txtWKH2.Text, "0")).Replace(",", "")
                    myClsBG0200BL.WKRRT1 = CStr(IIf(IsNumeric(txtWKRRT1.Text), txtWKRRT1.Text, "0")).Replace(",", "")
                    myClsBG0200BL.WKRRT2 = CStr(IIf(IsNumeric(txtWKRRT2.Text), txtWKRRT2.Text, "0")).Replace(",", "")
                    myClsBG0200BL.WKRRT3 = CStr(IIf(IsNumeric(txtWKRRT3.Text), txtWKRRT3.Text, "0")).Replace(",", "")
                    myClsBG0200BL.WKRRT4 = CStr(IIf(IsNumeric(txtWKRRT4.Text), txtWKRRT4.Text, "0")).Replace(",", "")
                    myClsBG0200BL.WKRRT5 = CStr(IIf(IsNumeric(txtWKRRT5.Text), txtWKRRT5.Text, "0")).Replace(",", "")
                    myClsBG0200BL.MTPWB = CStr(IIf(IsNumeric(txtMTPWB.Text), txtMTPWB.Text, "0")).Replace(",", "")
                    myClsBG0200BL.RevNo = Me.CurrRevNo
                    myClsBG0200BL.SaveWKH()

                    If Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then
                        myClsBG0200BL.MTP_SUM1 = CStr(IIf(IsNumeric(txtMTPInv1.Text), txtMTPInv1.Text, "0")).Replace(",", "")
                        myClsBG0200BL.MTP_SUM2 = CStr(IIf(IsNumeric(txtMTPInv2.Text), txtMTPInv2.Text, "0")).Replace(",", "")
                        myClsBG0200BL.MTP_SUM3 = CStr(IIf(IsNumeric(txtMTPInv3.Text), txtMTPInv3.Text, "0")).Replace(",", "")
                        myClsBG0200BL.MTP_SUM4 = CStr(IIf(IsNumeric(txtMTPInv4.Text), txtMTPInv4.Text, "0")).Replace(",", "")
                        myClsBG0200BL.MTP_SUM5 = CStr(IIf(IsNumeric(txtMTPInv5.Text), txtMTPInv5.Text, "0")).Replace(",", "")
                        myClsBG0200BL.MTP_PY_SUM1 = CStr(IIf(IsNumeric(txtPYInv1.Text), txtPYInv1.Text, "0")).Replace(",", "")
                        myClsBG0200BL.MTP_PY_SUM2 = CStr(IIf(IsNumeric(txtPYInv2.Text), txtPYInv2.Text, "0")).Replace(",", "")
                        myClsBG0200BL.MTP_PY_SUM3 = CStr(IIf(IsNumeric(txtPYInv3.Text), txtPYInv3.Text, "0")).Replace(",", "")
                        myClsBG0200BL.MTP_PY_SUM4 = CStr(IIf(IsNumeric(txtPYInv4.Text), txtPYInv4.Text, "0")).Replace(",", "")
                        myClsBG0200BL.MTP_PY_SUM5 = CStr(IIf(IsNumeric(txtPYInv5.Text), txtPYInv5.Text, "0")).Replace(",", "")
                        myClsBG0200BL.SaveMTPInvestment()
                    End If
                Else
                    '// Save MTP_SUM
                    myClsBG0200BL.MTP_SUM1 = CStr(IIf(IsNumeric(txtMTP_SUM1.Text), txtMTP_SUM1.Text, "0")).Replace(",", "")
                    myClsBG0200BL.MTP_SUM2 = CStr(IIf(IsNumeric(txtMTP_SUM2.Text), txtMTP_SUM2.Text, "0")).Replace(",", "")
                    myClsBG0200BL.MTP_SUM3 = CStr(IIf(IsNumeric(txtMTP_SUM3.Text), txtMTP_SUM3.Text, "0")).Replace(",", "")
                    myClsBG0200BL.MTP_SUM4 = CStr(IIf(IsNumeric(txtMTP_SUM4.Text), txtMTP_SUM4.Text, "0")).Replace(",", "")
                    myClsBG0200BL.MTP_SUM5 = CStr(IIf(IsNumeric(txtMTP_SUM5.Text), txtMTP_SUM5.Text, "0")).Replace(",", "")
                    myClsBG0200BL.RevNo = Me.CurrRevNo
                    myClsBG0200BL.SaveMTP_SUM()
                End If

                '// Write Trans Log
                myClsBG0200BL.UserPIC = "0000"
                myClsBG0200BL.LogOperationCd = enumOperationCd.UpRevision
                myClsBG0200BL.WriteTransLog()

                MessageBox.Show("Budget journal was saved", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                '// Save Data
                Dim blnSave As Boolean = False
                If blnReInputByOrder = True Then
                    blnSave = myClsBG0200BL.SaveBudgetDataReInputByOrder()
                Else
                    blnSave = myClsBG0200BL.SaveBudgetData()
                End If
                If blnSave = True Then

                    If Me.myOperationCd = enumOperationCd.AdjustBudget Or Me.myOperationCd = enumOperationCd.AdjustBudgetDirectInput Then
                        If Me.GetBudgetType() = P_BUDGET_TYPE_EXPENSE Then
                            '// Save WKH
                            myClsBG0200BL.WKH1 = CStr(IIf(IsNumeric(txtWKH1.Text), txtWKH1.Text, "0")).Replace(",", "")
                            myClsBG0200BL.WKH2 = CStr(IIf(IsNumeric(txtWKH2.Text), txtWKH2.Text, "0")).Replace(",", "")
                            myClsBG0200BL.WKRRT1 = CStr(IIf(IsNumeric(txtWKRRT1.Text), txtWKRRT1.Text, "0")).Replace(",", "")
                            myClsBG0200BL.WKRRT2 = CStr(IIf(IsNumeric(txtWKRRT2.Text), txtWKRRT2.Text, "0")).Replace(",", "")
                            myClsBG0200BL.WKRRT3 = CStr(IIf(IsNumeric(txtWKRRT3.Text), txtWKRRT3.Text, "0")).Replace(",", "")
                            myClsBG0200BL.WKRRT4 = CStr(IIf(IsNumeric(txtWKRRT4.Text), txtWKRRT4.Text, "0")).Replace(",", "")
                            myClsBG0200BL.WKRRT5 = CStr(IIf(IsNumeric(txtWKRRT5.Text), txtWKRRT5.Text, "0")).Replace(",", "")
                            myClsBG0200BL.MTPWB = CStr(IIf(IsNumeric(txtMTPWB.Text), txtMTPWB.Text, "0")).Replace(",", "")
                            myClsBG0200BL.RevNo = Me.CurrRevNo
                            myClsBG0200BL.SaveWKH()

                            If Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then
                                myClsBG0200BL.MTP_SUM1 = CStr(IIf(IsNumeric(txtMTPInv1.Text), txtMTPInv1.Text, "0")).Replace(",", "")
                                myClsBG0200BL.MTP_SUM2 = CStr(IIf(IsNumeric(txtMTPInv2.Text), txtMTPInv2.Text, "0")).Replace(",", "")
                                myClsBG0200BL.MTP_SUM3 = CStr(IIf(IsNumeric(txtMTPInv3.Text), txtMTPInv3.Text, "0")).Replace(",", "")
                                myClsBG0200BL.MTP_SUM4 = CStr(IIf(IsNumeric(txtMTPInv4.Text), txtMTPInv4.Text, "0")).Replace(",", "")
                                myClsBG0200BL.MTP_SUM5 = CStr(IIf(IsNumeric(txtMTPInv5.Text), txtMTPInv5.Text, "0")).Replace(",", "")
                                myClsBG0200BL.MTP_PY_SUM1 = CStr(IIf(IsNumeric(txtPYInv1.Text), txtPYInv1.Text, "0")).Replace(",", "")
                                myClsBG0200BL.MTP_PY_SUM2 = CStr(IIf(IsNumeric(txtPYInv2.Text), txtPYInv2.Text, "0")).Replace(",", "")
                                myClsBG0200BL.MTP_PY_SUM3 = CStr(IIf(IsNumeric(txtPYInv3.Text), txtPYInv3.Text, "0")).Replace(",", "")
                                myClsBG0200BL.MTP_PY_SUM4 = CStr(IIf(IsNumeric(txtPYInv4.Text), txtPYInv4.Text, "0")).Replace(",", "")
                                myClsBG0200BL.MTP_PY_SUM5 = CStr(IIf(IsNumeric(txtPYInv5.Text), txtPYInv5.Text, "0")).Replace(",", "")
                                myClsBG0200BL.SaveMTPInvestment()
                            End If
                        Else
                            '// Save MTP SUM
                            myClsBG0200BL.MTP_SUM1 = CStr(IIf(IsNumeric(txtMTP_SUM1.Text), txtMTP_SUM1.Text, "0")).Replace(",", "")
                            myClsBG0200BL.MTP_SUM2 = CStr(IIf(IsNumeric(txtMTP_SUM2.Text), txtMTP_SUM2.Text, "0")).Replace(",", "")
                            myClsBG0200BL.MTP_SUM3 = CStr(IIf(IsNumeric(txtMTP_SUM3.Text), txtMTP_SUM3.Text, "0")).Replace(",", "")
                            myClsBG0200BL.MTP_SUM4 = CStr(IIf(IsNumeric(txtMTP_SUM4.Text), txtMTP_SUM4.Text, "0")).Replace(",", "")
                            myClsBG0200BL.MTP_SUM5 = CStr(IIf(IsNumeric(txtMTP_SUM5.Text), txtMTP_SUM5.Text, "0")).Replace(",", "")
                            myClsBG0200BL.RevNo = Me.CurrRevNo
                            myClsBG0200BL.SaveMTP_SUM()
                        End If
                    End If

                    '// Write Trans Log
                    If Me.OperationCd = enumOperationCd.AdjustBudget Then
                        myClsBG0200BL.LogOperationCd = enumOperationCd.AdjustBudget
                    ElseIf Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput Then
                        myClsBG0200BL.LogOperationCd = enumOperationCd.AdjustBudgetDirectInput
                    Else
                        myClsBG0200BL.LogOperationCd = enumOperationCd.InputBudget
                    End If

                    HighlightWorkingBGAndComment()
                    myClsBG0200BL.WriteTransLog()

                    '// Refresh side menu
                    p_frmBG0010.ShowBudgetMenu()

                    MessageBox.Show("Budget journal was saved", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("Can not save Budget journal", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmdSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSubmit.Click
        Try
            If MessageBox.Show("Are you sure to submit this budget journal?", Me.Text, MessageBoxButtons.YesNo, _
                                      MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            Me.grvBudget1.EndEdit()
            Me.grvBudget2.EndEdit()
            Me.grvBudget3.EndEdit()
            Me.grvBudget4.EndEdit()

            '// Submit budget journal
            '// Set Parameters
            myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
            myClsBG0200BL.PeriodType = Me.GetPeriodType()
            myClsBG0200BL.BudgetType = Me.GetBudgetType()
            myClsBG0200BL.UserId = p_strUserId
            myClsBG0200BL.UserPIC = CStr(cboPIC.SelectedValue)
            myClsBG0200BL.Status = CStr(enumBudgetStatus.Submit)
            myClsBG0200BL.RevNo = Me.CurrRevNo
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then          '// Original Budget
                myClsBG0200BL.BudgetList = CType(grvBudget1.DataSource, DataTable)

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then      '// Estimate Budget
                myClsBG0200BL.BudgetList = CType(grvBudget2.DataSource, DataTable)

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then        '// Revise Budget
                myClsBG0200BL.BudgetList = CType(grvBudget3.DataSource, DataTable)
            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then        '// MTP Budget
                myClsBG0200BL.BudgetList = CType(grvBudget4.DataSource, DataTable)
            End If
            Dim blnSave As Boolean = False
            If blnReInputByOrder = True Then
                blnSave = myClsBG0200BL.SaveBudgetDataReInputByOrder()
            Else
                blnSave = myClsBG0200BL.SaveBudgetData()
            End If

            '// Call Function
            If blnSave = True Then
                MessageBox.Show("Budget journal was submitted", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                '// Write Trans Log
                myClsBG0200BL.LogOperationCd = enumOperationCd.SubmitBudget
                myClsBG0200BL.WriteTransLog()

                '// Send auto mail
                If p_blnSendAutoMail Then
                    myClsBG0200BL.UserPIC = CStr(cboPIC.SelectedValue)
                    myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
                    myClsBG0200BL.PeriodType = Me.GetPeriodType()
                    myClsBG0200BL.BudgetType = Me.GetBudgetType()
                    myClsBG0200BL.OperationCd = enumOperationCd.SubmitBudget
                    myClsBG0200BL.ProjectNo = Me.GetProjectNo()

                    myClsBG0200BL.SendAutoMail()
                End If

                '// Refresh side menu
                p_frmBG0010.ShowBudgetMenu()
                myForceCloseFlg = True
                Me.Close()
            Else
                MessageBox.Show("Can not submit Budget journal", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmdReject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReject.Click
        Dim rtn As Boolean = False

        Try
            If CStr(cboPIC.SelectedValue) = "0000" Then
                If MessageBox.Show("Are you sure to reject all of this budget journal?", Me.Text, MessageBoxButtons.YesNo, _
                                   MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If
            Else
                If MessageBox.Show("Are you sure to reject this budget journal?", Me.Text, MessageBoxButtons.YesNo, _
                                   MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If
            End If

            '// Reject Budget journal
            '// Set Parameters
            myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
            myClsBG0200BL.PeriodType = Me.GetPeriodType()
            myClsBG0200BL.BudgetType = Me.GetBudgetType()
            myClsBG0200BL.UserId = p_strUserId
            myClsBG0200BL.UserPIC = CStr(cboPIC.SelectedValue)
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            '// Call Function
            If Me.OperationCd = enumOperationCd.ApproveBudget Then
                myClsBG0200BL.RevNo = "1"

                If blnReInputByOrder = True Then
                    rtn = myClsBG0200BL.SaveRejectBudgetDataReInputByOrder()
                Else
                    rtn = myClsBG0200BL.SaveRejectBudgetData()
                End If
            ElseIf Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput Then
                myClsBG0200BL.RevNo = "1"
                rtn = myClsBG0200BL.SaveRejectBudgetData2()

                If rtn = True Then
                    '// Clear revision which Rev No > 1
                    If CStr(cboPIC.SelectedValue) = "0000" Then
                        Dim dt As DataTable = CType(cboRevNo.DataSource, DataTable)
                        For i As Integer = 0 To dt.Rows.Count - 1
                            If CStr(dt.Rows(i).Item("REV_NO")) <> "1" Then
                                DeleteRevision(CStr(dt.Rows(i).Item("REV_NO")))
                            End If
                        Next
                    End If
                End If
            ElseIf Me.OperationCd = enumOperationCd.Authorize1 Or Me.OperationCd = enumOperationCd.Authorize2 Then
                myClsBG0200BL.RevNo = Me.CurrRevNo
                rtn = myClsBG0200BL.SaveRejectBudgetData3()
            End If

            If rtn = True Then
                MessageBox.Show("Budget journal was rejected", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                '// Write Trans Log
                If Me.OperationCd = enumOperationCd.ApproveBudget Then
                    myClsBG0200BL.LogOperationCd = enumOperationCd.RejectBudget1
                ElseIf Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput Then
                    myClsBG0200BL.LogOperationCd = enumOperationCd.RejectBudget2
                ElseIf Me.OperationCd = enumOperationCd.Authorize1 Or Me.OperationCd = enumOperationCd.Authorize2 Then
                    myClsBG0200BL.LogOperationCd = enumOperationCd.RejectBudget3
                End If
                myClsBG0200BL.WriteTransLog()

                '// Send auto mail
                If p_blnSendAutoMail Then
                    myClsBG0200BL.UserPIC = CStr(cboPIC.SelectedValue)
                    myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
                    myClsBG0200BL.PeriodType = Me.GetPeriodType()
                    myClsBG0200BL.BudgetType = Me.GetBudgetType()
                    myClsBG0200BL.ProjectNo = Me.GetProjectNo()

                    If Me.OperationCd = enumOperationCd.ApproveBudget Then
                        myClsBG0200BL.OperationCd = enumOperationCd.RejectBudget1
                    ElseIf Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput Then
                        myClsBG0200BL.OperationCd = enumOperationCd.RejectBudget2
                    ElseIf Me.OperationCd = enumOperationCd.Authorize1 Or Me.OperationCd = enumOperationCd.Authorize2 Then
                        myClsBG0200BL.OperationCd = enumOperationCd.RejectBudget3
                    End If

                    myClsBG0200BL.SendAutoMail()
                End If

                '// Refresh side menu
                p_frmBG0010.ShowBudgetMenu()
                myForceCloseFlg = True
                Me.Close()
            Else
                MessageBox.Show("Can not reject Budget journal", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmdApprove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdApprove.Click
        Try
            If MessageBox.Show("Are you sure to approve this budget journal?", Me.Text, MessageBoxButtons.YesNo, _
                           MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            '// Approve Budget journal
            '// Set Parameters
            myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
            myClsBG0200BL.PeriodType = Me.GetPeriodType()
            myClsBG0200BL.BudgetType = Me.GetBudgetType()
            myClsBG0200BL.UserId = p_strUserId
            myClsBG0200BL.UserPIC = CStr(cboPIC.SelectedValue)
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            '// Call Function
            Dim blnSave As Boolean = False

            If blnReInputByOrder = True Then
                If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then          '// Original Budget
                    myClsBG0200BL.BudgetList = CType(grvBudget1.DataSource, DataTable)
                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then      '// Estimate Budget
                    myClsBG0200BL.BudgetList = CType(grvBudget2.DataSource, DataTable)
                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then        '// Revise Budget
                    myClsBG0200BL.BudgetList = CType(grvBudget3.DataSource, DataTable)
                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then         '// MTP Budget
                    myClsBG0200BL.BudgetList = CType(grvBudget4.DataSource, DataTable)
                End If

                Dim CountSel As Integer = 0
                Dim drS As DataRow
                Dim dtSave As New DataTable
                dtSave.Columns.Add("BudgetYear", GetType(String))
                dtSave.Columns.Add("PeriodType", GetType(String))
                dtSave.Columns.Add("RevNo", GetType(String))
                dtSave.Columns.Add("ProjectNo", GetType(String))
                dtSave.Columns.Add("BudgetOrder", GetType(String))
                dtSave.Columns.Add("Status", GetType(String))

                If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then   '// Original Budget
                    'grvBudget1
                    For Each row As DataGridViewRow In grvBudget1.Rows
                        Dim isChecked As Boolean = DirectCast(grvBudget1(0, row.Index).Value, Boolean)

                        drS = dtSave.NewRow
                        drS("BudgetYear") = Me.GetBudgetYear()
                        drS("PeriodType") = Me.GetPeriodType()
                        drS("RevNo") = Me.CurrRevNo
                        drS("ProjectNo") = Me.GetProjectNo()
                        drS("BudgetOrder") = CStr(row.Cells("OrderNo1").Value)
                        drS("Status") = CStr(enumBudgetStatus.Submit)
                        dtSave.Rows.Add(drS)
                    Next
                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then   '// Estimate Budget
                    For Each row As DataGridViewRow In grvBudget2.Rows
                        Dim isChecked As Boolean = DirectCast(grvBudget2(0, row.Index).Value, Boolean)
                        drS = dtSave.NewRow
                        drS("BudgetYear") = Me.GetBudgetYear()
                        drS("PeriodType") = Me.GetPeriodType()
                        drS("RevNo") = Me.CurrRevNo
                        drS("ProjectNo") = Me.GetProjectNo()
                        drS("BudgetOrder") = CStr(row.Cells("OrderNo2").Value)
                        drS("Status") = CStr(enumBudgetStatus.Submit)
                        dtSave.Rows.Add(drS)
                    Next
                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then   '// Revise Budget        
                    For Each row As DataGridViewRow In grvBudget3.Rows
                        Dim isChecked As Boolean = DirectCast(grvBudget3(0, row.Index).Value, Boolean)
                        drS = dtSave.NewRow
                        drS("BudgetYear") = Me.GetBudgetYear()
                        drS("PeriodType") = Me.GetPeriodType()
                        drS("RevNo") = Me.CurrRevNo
                        drS("ProjectNo") = Me.GetProjectNo()
                        drS("BudgetOrder") = CStr(row.Cells("OrderNo3").Value)
                        drS("Status") = CStr(enumBudgetStatus.Submit)
                        dtSave.Rows.Add(drS)
                    Next
                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then   '// MTP Budget        
                    For Each row As DataGridViewRow In grvBudget4.Rows
                        Dim isChecked As Boolean = DirectCast(grvBudget4(0, row.Index).Value, Boolean)
                        drS = dtSave.NewRow
                        drS("BudgetYear") = Me.GetBudgetYear()
                        drS("PeriodType") = Me.GetPeriodType()
                        drS("RevNo") = Me.CurrRevNo
                        drS("ProjectNo") = Me.GetProjectNo()
                        drS("BudgetOrder") = CStr(row.Cells("OrderNo4").Value)
                        drS("Status") = CStr(enumBudgetStatus.Submit)
                        dtSave.Rows.Add(drS)
                    Next
                End If

                Dim conn As SqlConnection = Nothing
                Dim trans As SqlTransaction
                Dim success As Boolean = False

                conn = New SqlConnection(My.Settings.ConnStr)
                conn.Open()
                trans = conn.BeginTransaction()
                Try
                    myClsBG0200BL.dtSave = dtSave
                    If myClsBG0200BL.SaveApproveBudgetDataReInputByOrder(conn, trans) = True Then
                        success = True
                        blnSave = True
                    Else
                        success = False
                        blnSave = False
                    End If

                    If success Then
                        trans.Commit()

                        '// Write Transaction Log
                        WriteTransactionLog(CStr(enumOperationCd.AdjustBudgetDirectInput), "", "", "", "", "", "")

                        '// Refresh side menu
                        p_frmBG0010.ShowBudgetMenu()

                        myForceCloseFlg = True
                        Me.Close()
                    Else
                        trans.Rollback()
                    End If
                Catch ex As Exception
                    Throw ex
                End Try
            Else
                blnSave = myClsBG0200BL.SaveApproveBudgetData()
            End If

            If blnSave = True Then
                MessageBox.Show("Budget journal was approved", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                '// Write Trans Log
                myClsBG0200BL.RevNo = "1"
                myClsBG0200BL.LogOperationCd = enumOperationCd.ApproveBudget
                myClsBG0200BL.WriteTransLog()

                '// Send auto mail
                If p_blnSendAutoMail Then
                    myClsBG0200BL.UserPIC = CStr(cboPIC.SelectedValue)
                    myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
                    myClsBG0200BL.PeriodType = Me.GetPeriodType()
                    myClsBG0200BL.BudgetType = Me.GetBudgetType()
                    myClsBG0200BL.OperationCd = enumOperationCd.ApproveBudget
                    myClsBG0200BL.ProjectNo = Me.GetProjectNo()

                    myClsBG0200BL.SendAutoMail()
                End If

                '// Refresh side menu
                p_frmBG0010.ShowBudgetMenu()

                myForceCloseFlg = True
                Me.Close()
            Else
                MessageBox.Show("Can not approve Budget journal", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmdSubmit2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSubmit2.Click
        Try
            'Check ReInputByOrder 
            Dim dtDataReInput As DataTable
            myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
            myClsBG0200BL.PeriodType = Me.GetPeriodType()
            myClsBG0200BL.BudgetType = Me.GetBudgetType()
            myClsBG0200BL.UserPIC = "0000"
            myClsBG0200BL.RevNo = Me.CurrRevNo
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            dtDataReInput = myClsBG0200BL.GetBudGetDataReInputNoStatus
            If Not dtDataReInput Is Nothing AndAlso dtDataReInput.Rows.Count > 0 Then
                MessageBox.Show("Can not submit this budget journal. Some budget orders are in Re-Input process.)", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            '// Check Max Revision
            myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
            myClsBG0200BL.PeriodType = Me.GetPeriodType()
            myClsBG0200BL.BudgetType = Me.GetBudgetType()
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            myClsBG0200BL.GetMaxRevNo()
            If Me.CurrRevNo <> myClsBG0200BL.RevNo Then
                MessageBox.Show("You can submit Max Revision only.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
            If lblStatus.Text <> "Approved" Then
                MessageBox.Show("You can submit budget journal when status is approved only.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
            If MessageBox.Show("Are you sure to submit this budget journal?", Me.Text, MessageBoxButtons.YesNo, _
                               MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            Me.grvBudget1.EndEdit()
            Me.grvBudget2.EndEdit()
            Me.grvBudget3.EndEdit()
            Me.grvBudget4.EndEdit()

            '// Adjust Budget journal
            '// Set Parameters
            myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
            myClsBG0200BL.PeriodType = Me.GetPeriodType()
            myClsBG0200BL.BudgetType = Me.GetBudgetType()
            myClsBG0200BL.UserId = p_strUserId
            myClsBG0200BL.UserPIC = CStr(cboPIC.SelectedValue)
            myClsBG0200BL.Status = CStr(enumBudgetStatus.Approve)
            myClsBG0200BL.RevNo = Me.CurrRevNo
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then          '// Original Budget
                myClsBG0200BL.BudgetList = CType(grvBudget1.DataSource, DataTable)
            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then      '// Estimate Budget
                myClsBG0200BL.BudgetList = CType(grvBudget2.DataSource, DataTable)
            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then        '// Revise Budget
                myClsBG0200BL.BudgetList = CType(grvBudget3.DataSource, DataTable)
            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then         '// MTP Budget
                myClsBG0200BL.BudgetList = CType(grvBudget4.DataSource, DataTable)
            End If

            '// Call Function: Save Budget and Adjust Budget
            If myClsBG0200BL.SaveBudgetData() = True Then
                If myClsBG0200BL.SaveAdjustBudgetData() = True Then
                    MessageBox.Show("Budget journal was submitted", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                    '// Write Trans Log
                    myClsBG0200BL.LogOperationCd = enumOperationCd.SubmitBudget
                    myClsBG0200BL.WriteTransLog()

                    '// Send auto mail
                    If p_blnSendAutoMail Then
                        myClsBG0200BL.UserPIC = CStr(cboPIC.SelectedValue)
                        myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
                        myClsBG0200BL.PeriodType = Me.GetPeriodType()
                        myClsBG0200BL.BudgetType = Me.GetBudgetType()
                        If Me.OperationCd = enumOperationCd.AdjustBudget Then
                            myClsBG0200BL.OperationCd = enumOperationCd.AdjustBudget
                        ElseIf Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput Then
                            myClsBG0200BL.OperationCd = enumOperationCd.AdjustBudgetDirectInput
                        End If
                        myClsBG0200BL.ProjectNo = Me.GetProjectNo()

                        myClsBG0200BL.SendAutoMail()
                    End If

                    '// Refresh side menu
                    p_frmBG0010.ShowBudgetMenu()
                    myForceCloseFlg = True
                    Me.Close()
                Else
                    MessageBox.Show("Can not submit Budget journal", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Else
                MessageBox.Show("Can not submit Budget journal", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmdAuth1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAuth1.Click
        Try
            If MessageBox.Show("Are you sure to authorize this budget journal?", Me.Text, MessageBoxButtons.YesNo, _
                                       MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            '// AUTH1 Budget journal
            '// Set Parameters
            myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
            myClsBG0200BL.PeriodType = Me.GetPeriodType()
            myClsBG0200BL.BudgetType = Me.GetBudgetType()
            myClsBG0200BL.UserId = p_strUserId
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            '// Call Function
            If myClsBG0200BL.SaveAuth1BudgetData() = True Then
                MessageBox.Show("Budget journal was authorized", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                '// Write Trans Log
                myClsBG0200BL.UserPIC = "0000"
                myClsBG0200BL.RevNo = Me.CurrRevNo
                myClsBG0200BL.LogOperationCd = enumOperationCd.Authorize1
                myClsBG0200BL.WriteTransLog()

                '// Send auto mail
                If p_blnSendAutoMail Then
                    myClsBG0200BL.UserPIC = CStr(cboPIC.SelectedValue)
                    myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
                    myClsBG0200BL.PeriodType = Me.GetPeriodType()
                    myClsBG0200BL.BudgetType = Me.GetBudgetType()
                    myClsBG0200BL.OperationCd = enumOperationCd.Authorize1
                    myClsBG0200BL.ProjectNo = Me.GetProjectNo()

                    myClsBG0200BL.SendAutoMail()
                End If

                '// Refresh side menu
                p_frmBG0010.ShowBudgetMenu()
                myForceCloseFlg = True
                Me.Close()
            Else
                MessageBox.Show("Can not authorize Budget journal", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmdAuth2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAuth2.Click
        Try
            If MessageBox.Show("Are you sure to authorize this budget journal?", Me.Text, MessageBoxButtons.YesNo, _
                                       MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            '// AUTH2 Budget journal
            '// Set Parameters
            myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
            myClsBG0200BL.PeriodType = Me.GetPeriodType()
            myClsBG0200BL.BudgetType = Me.GetBudgetType()
            myClsBG0200BL.UserId = p_strUserId
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            '// Call Function
            If myClsBG0200BL.SaveAuth2BudgetData() = True Then
                MessageBox.Show("Budget journal was authorized", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                '// Write Trans Log
                myClsBG0200BL.UserPIC = "0000"
                myClsBG0200BL.RevNo = Me.CurrRevNo
                myClsBG0200BL.LogOperationCd = enumOperationCd.Authorize2
                myClsBG0200BL.WriteTransLog()

                '// Send auto mail
                If p_blnSendAutoMail Then
                    myClsBG0200BL.UserPIC = CStr(cboPIC.SelectedValue)
                    myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
                    myClsBG0200BL.PeriodType = Me.GetPeriodType()
                    myClsBG0200BL.BudgetType = Me.GetBudgetType()
                    myClsBG0200BL.OperationCd = enumOperationCd.Authorize2
                    myClsBG0200BL.ProjectNo = Me.GetProjectNo()

                    myClsBG0200BL.SendAutoMail()
                End If

                '// Refresh side menu
                p_frmBG0010.ShowBudgetMenu()
                myForceCloseFlg = True
                Me.Close()
            Else
                MessageBox.Show("Can not authorize Budget journal", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmdAdjust_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAdjust.Click
        Try
            If lblStatus.Text <> "Approved" Then
                MessageBox.Show("You can adjust working budget when status is approved only.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            If MessageBox.Show("Are you sure to adjust working budget to selected row?", Me.Text, MessageBoxButtons.YesNo, _
                              MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            Me.grvBudget1.EndEdit()
            Me.grvBudget2.EndEdit()
            Me.grvBudget3.EndEdit()
            Me.grvBudget4.EndEdit()

            'Me.grvBudget1.BeginEdit(True)

            Me.Cursor = Cursors.WaitCursor
            myDataLoadingFlg = False
            '// Adjust Working Budget
            AdjustWorkingBG()

            'Me.grvBudget1.EndEdit()
            'Me.grvBudget2.EndEdit()
            'Me.grvBudget3.EndEdit()
            'Me.grvBudget4.EndEdit()

            '// Calculate Total/Diff
            If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then
                CalcOriginalBudget()
            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then
                CalcEstimateBudget()
            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then
                CalcReviseBudget(False)
            End If

            '// Highlight Working Budget
            HighlightWorkingBG()
            '// Highlight Transfer Cost
            HighlightTransferCost()
            myDataLoadingFlg = False
            Me.Cursor = Cursors.Default
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmdUpRev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpRev.Click
        Try
            'Check ReInputByOrder 
            Dim dtDataReInput As DataTable
            myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
            myClsBG0200BL.PeriodType = Me.GetPeriodType()
            myClsBG0200BL.BudgetType = Me.GetBudgetType()
            myClsBG0200BL.UserPIC = "0000"
            myClsBG0200BL.RevNo = Me.CurrRevNo
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            dtDataReInput = myClsBG0200BL.GetBudGetDataReInputNoStatus
            If Not dtDataReInput Is Nothing AndAlso dtDataReInput.Rows.Count > 0 Then
                MessageBox.Show("Can not Up Revision this budget journal. Some budget orders are in Re-Input process.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            If MessageBox.Show("Are you sure to add new revision of this budget journal?", Me.Text, MessageBoxButtons.YesNo, _
                              MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            Me.grvBudget1.EndEdit()
            Me.grvBudget2.EndEdit()
            Me.grvBudget3.EndEdit()
            Me.grvBudget4.EndEdit()
            Me.Cursor = Cursors.WaitCursor

            '// Add new revision of budget journal
            If UpRevision() = True Then

                '// Transfer Cost
                myClsBG0200BL.AdjustTransferCost()

                If Me.cboAccount.Items.Count > 0 Then
                    Me.cboAccount.SelectedIndex = 0
                End If
                If Me.cboCost.Items.Count > 0 Then
                    Me.cboCost.SelectedIndex = 0
                End If
                If Me.cboCostType.Items.Count > 0 Then
                    Me.cboCostType.SelectedIndex = 0
                End If
                If Me.cboDept.Items.Count > 0 Then
                    Me.cboDept.SelectedIndex = 0
                End If

                Me.lblStatus.Text = "Approved"
                Me.lblRefRevNo.Text = Me.CurrRevNo()

                GetCurrentRevNo()
                LoadRevNoList()

                ReloadPICList(CStr(cboRevNo.SelectedValue))

                If Me.GetBudgetType() = P_BUDGET_TYPE_EXPENSE Then
                    '// Save WKH
                    myClsBG0200BL.WKH1 = CStr(IIf(IsNumeric(txtWKH1.Text), txtWKH1.Text, "0")).Replace(",", "")
                    myClsBG0200BL.WKH2 = CStr(IIf(IsNumeric(txtWKH2.Text), txtWKH2.Text, "0")).Replace(",", "")

                    myClsBG0200BL.WKRRT1 = CStr(IIf(IsNumeric(txtWKRRT1.Text), txtWKRRT1.Text, "0")).Replace(",", "")
                    myClsBG0200BL.WKRRT2 = CStr(IIf(IsNumeric(txtWKRRT2.Text), txtWKRRT2.Text, "0")).Replace(",", "")
                    myClsBG0200BL.WKRRT3 = CStr(IIf(IsNumeric(txtWKRRT3.Text), txtWKRRT3.Text, "0")).Replace(",", "")
                    myClsBG0200BL.WKRRT4 = CStr(IIf(IsNumeric(txtWKRRT4.Text), txtWKRRT4.Text, "0")).Replace(",", "")
                    myClsBG0200BL.WKRRT5 = CStr(IIf(IsNumeric(txtWKRRT5.Text), txtWKRRT5.Text, "0")).Replace(",", "")

                    myClsBG0200BL.MTPWB = CStr(IIf(IsNumeric(txtMTPWB.Text), txtMTPWB.Text, "0")).Replace(",", "")

                    myClsBG0200BL.RevNo = Me.CurrRevNo
                    myClsBG0200BL.SaveWKH()

                    If Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then

                        myClsBG0200BL.MTP_SUM1 = CStr(IIf(IsNumeric(txtMTPInv1.Text), txtMTPInv1.Text, "0")).Replace(",", "")
                        myClsBG0200BL.MTP_SUM2 = CStr(IIf(IsNumeric(txtMTPInv2.Text), txtMTPInv2.Text, "0")).Replace(",", "")
                        myClsBG0200BL.MTP_SUM3 = CStr(IIf(IsNumeric(txtMTPInv3.Text), txtMTPInv3.Text, "0")).Replace(",", "")
                        myClsBG0200BL.MTP_SUM4 = CStr(IIf(IsNumeric(txtMTPInv4.Text), txtMTPInv4.Text, "0")).Replace(",", "")
                        myClsBG0200BL.MTP_SUM5 = CStr(IIf(IsNumeric(txtMTPInv5.Text), txtMTPInv5.Text, "0")).Replace(",", "")

                        myClsBG0200BL.MTP_PY_SUM1 = CStr(IIf(IsNumeric(txtPYInv1.Text), txtPYInv1.Text, "0")).Replace(",", "")
                        myClsBG0200BL.MTP_PY_SUM2 = CStr(IIf(IsNumeric(txtPYInv2.Text), txtPYInv2.Text, "0")).Replace(",", "")
                        myClsBG0200BL.MTP_PY_SUM3 = CStr(IIf(IsNumeric(txtPYInv3.Text), txtPYInv3.Text, "0")).Replace(",", "")
                        myClsBG0200BL.MTP_PY_SUM4 = CStr(IIf(IsNumeric(txtPYInv4.Text), txtPYInv4.Text, "0")).Replace(",", "")
                        myClsBG0200BL.MTP_PY_SUM5 = CStr(IIf(IsNumeric(txtPYInv5.Text), txtPYInv5.Text, "0")).Replace(",", "")

                        myClsBG0200BL.SaveMTPInvestment()
                    End If

                Else
                    '// Save MTP SUM
                    myClsBG0200BL.MTP_SUM1 = CStr(IIf(IsNumeric(txtMTP_SUM1.Text), txtMTP_SUM1.Text, "0")).Replace(",", "")
                    myClsBG0200BL.MTP_SUM2 = CStr(IIf(IsNumeric(txtMTP_SUM2.Text), txtMTP_SUM2.Text, "0")).Replace(",", "")
                    myClsBG0200BL.MTP_SUM3 = CStr(IIf(IsNumeric(txtMTP_SUM3.Text), txtMTP_SUM3.Text, "0")).Replace(",", "")
                    myClsBG0200BL.MTP_SUM4 = CStr(IIf(IsNumeric(txtMTP_SUM4.Text), txtMTP_SUM4.Text, "0")).Replace(",", "")
                    myClsBG0200BL.MTP_SUM5 = CStr(IIf(IsNumeric(txtMTP_SUM5.Text), txtMTP_SUM5.Text, "0")).Replace(",", "")
                    myClsBG0200BL.RevNo = Me.CurrRevNo
                    myClsBG0200BL.SaveMTP_SUM()
                End If

                '// Write Trans Log
                myClsBG0200BL.LogOperationCd = enumOperationCd.UpRevision
                myClsBG0200BL.WriteTransLog()

                MessageBox.Show("Budget journal was saved", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else

                MessageBox.Show("Can not save Budget journal", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

            Me.Cursor = Cursors.Default
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmdDelRev_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDelRev.Click
        Try
            If Me.CurrRevNo = "1" Then
                MessageBox.Show("Can not delete the first revision of Budget journal", Me.Text, _
                                MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If

            If MessageBox.Show("Are you sure to delete this revision of budget journal?", Me.Text, MessageBoxButtons.YesNo, _
                              MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            Me.Cursor = Cursors.WaitCursor
            '// Delete revision of budget journal
            If DeleteRevision(Me.CurrRevNo) = True Then
                '// Write Trans Log
                myClsBG0200BL.UserPIC = "0000"
                myClsBG0200BL.RevNo = Me.CurrRevNo
                myClsBG0200BL.LogOperationCd = enumOperationCd.DelRevision
                myClsBG0200BL.WriteTransLog()

                If Me.cboAccount.Items.Count > 0 Then
                    Me.cboAccount.SelectedIndex = 0
                End If
                If Me.cboCost.Items.Count > 0 Then
                    Me.cboCost.SelectedIndex = 0
                End If
                If Me.cboCostType.Items.Count > 0 Then
                    Me.cboCostType.SelectedIndex = 0
                End If
                If Me.cboDept.Items.Count > 0 Then
                    Me.cboDept.SelectedIndex = 0
                End If

                GetCurrentRevNo()
                LoadRevNoList()
                MessageBox.Show("Budget journal was deleted", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Can not delete Budget journal", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

            Me.Cursor = Cursors.Default
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmdReInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReInput.Click
        Try
            If CStr(cboPIC.SelectedValue) = "0000" Then
                If MessageBox.Show("Are you sure to re-input [ALL] of this budget journal?", Me.Text, MessageBoxButtons.YesNo, _
                                   MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If
            Else
                If MessageBox.Show("Are you sure to re-input this budget journal?", Me.Text, MessageBoxButtons.YesNo, _
                                   MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If
            End If

            '// Reject Budget journal
            '// Set Parameters
            myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
            myClsBG0200BL.PeriodType = Me.GetPeriodType()
            myClsBG0200BL.BudgetType = Me.GetBudgetType()
            myClsBG0200BL.UserId = p_strUserId
            myClsBG0200BL.UserPIC = CStr(cboPIC.SelectedValue)
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            Dim rtn As Boolean = False
            Dim conn As SqlConnection = Nothing
            Dim trans As SqlTransaction
            Dim success As Boolean = False

            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()
            trans = conn.BeginTransaction()
            Try
                '// Call Function
                If Me.OperationCd = enumOperationCd.AdjustBudget Or Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput Then
                    myClsBG0200BL.RevNo = "1"
                    rtn = myClsBG0200BL.SaveRejectBudgetData4Tran(conn, trans)

                    If rtn = True Then
                        '// Set Parameters
                        myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
                        myClsBG0200BL.PeriodType = Me.GetPeriodType()
                        myClsBG0200BL.RevNo = "1"
                        myClsBG0200BL.BudgetType = Me.GetBudgetType()
                        myClsBG0200BL.UserPIC = CStr(cboPIC.SelectedValue)
                        myClsBG0200BL.ProjectNo = Me.GetProjectNo()

                        If myClsBG0200BL.DeleteBudgetDataReInputByOrder(conn, trans) = True Then
                            success = True
                        Else
                            success = False
                        End If
                    End If

                    If success Then
                        trans.Commit()

                        '// Write Transaction Log
                        WriteTransactionLog(CStr(enumOperationCd.ReInput), "", "", "", "", "", "")

                        '// Refresh side menu
                        p_frmBG0010.ShowBudgetMenu()

                        myForceCloseFlg = True
                        Me.Close()
                    Else
                        trans.Rollback()
                    End If

                    If success = True Then
                        '// Clear revision which Rev No > 1
                        If CStr(cboPIC.SelectedValue) = "0000" Then
                            myClsBG0200BL.GetMaxRevNo()
                            Dim intMaxRev As Integer = CInt(myClsBG0200BL.RevNo)
                            For i As Integer = 2 To intMaxRev
                                DeleteRevision(CStr(i))
                            Next
                        End If
                    End If
                End If
            Catch ex As Exception
                Throw ex
            End Try

            If rtn = True Then
                MessageBox.Show("Budget journal was rejected for re-input and set as new record", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                '// Write Trans Log
                myClsBG0200BL.LogOperationCd = enumOperationCd.ReInput
                myClsBG0200BL.WriteTransLog()

                '// Send auto mail
                If p_blnSendAutoMail Then
                    myClsBG0200BL.UserPIC = CStr(cboPIC.SelectedValue)
                    myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
                    myClsBG0200BL.PeriodType = Me.GetPeriodType()
                    myClsBG0200BL.BudgetType = Me.GetBudgetType()
                    myClsBG0200BL.OperationCd = enumOperationCd.RejectBudget2
                    myClsBG0200BL.ProjectNo = Me.GetProjectNo()

                    myClsBG0200BL.SendAutoMail()
                End If

                '// Refresh side menu
                p_frmBG0010.ShowBudgetMenu()
                myForceCloseFlg = True
                Me.Close()
            Else
                MessageBox.Show("Can not reject for re-input Budget journal", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmdReInputByOrder_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdReInputByOrder.Click
        Dim dtSave As New DataTable
        Dim drS As DataRow
        Dim CountSel As Integer = 0

        Try
            '// Check budget order selected
            If chkBudgetOrderSelected() = False Then
                MessageBox.Show("Please select Budget Order to Re-Input.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            '// Confirm dialog
            If MessageBox.Show("Are you sure to re-input by Budget Order No this budget journal?", Me.Text, MessageBoxButtons.YesNo, _
                               MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            Me.grvBudget1.EndEdit()
            Me.grvBudget2.EndEdit()
            Me.grvBudget3.EndEdit()
            Me.grvBudget4.EndEdit()

            dtSave.Columns.Add("BudgetYear", GetType(String))
            dtSave.Columns.Add("PeriodType", GetType(String))
            dtSave.Columns.Add("RevNo", GetType(String))
            dtSave.Columns.Add("ProjectNo", GetType(String))
            dtSave.Columns.Add("BudgetOrder", GetType(String))
            dtSave.Columns.Add("Status", GetType(String))
            If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then   '// Original Budget
                For Each row As DataGridViewRow In grvBudget1.Rows
                    Dim isChecked As Boolean = DirectCast(grvBudget1(0, row.Index).Value, Boolean)
                    If isChecked Then
                        drS = dtSave.NewRow
                        drS("BudgetYear") = Me.GetBudgetYear()
                        drS("PeriodType") = Me.GetPeriodType()
                        drS("RevNo") = "1"
                        drS("ProjectNo") = Me.GetProjectNo()
                        drS("BudgetOrder") = CStr(row.Cells("OrderNo1").Value)
                        drS("Status") = CStr(enumBudgetStatus.NewRecord)
                        dtSave.Rows.Add(drS)
                    End If
                Next
            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then   '// Estimate Budget
                For Each row As DataGridViewRow In grvBudget2.Rows
                    Dim isChecked As Boolean = DirectCast(grvBudget2(0, row.Index).Value, Boolean)
                    If isChecked Then
                        drS = dtSave.NewRow
                        drS("BudgetYear") = Me.GetBudgetYear()
                        drS("PeriodType") = Me.GetPeriodType()
                        drS("RevNo") = "1"
                        drS("ProjectNo") = Me.GetProjectNo()
                        drS("BudgetOrder") = CStr(row.Cells("OrderNo2").Value)
                        drS("Status") = CStr(enumBudgetStatus.NewRecord)
                        dtSave.Rows.Add(drS)
                    End If
                Next
            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then   '// Revise Budget        
                For Each row As DataGridViewRow In grvBudget3.Rows
                    Dim isChecked As Boolean = DirectCast(grvBudget3(0, row.Index).Value, Boolean)
                    If isChecked Then
                        drS = dtSave.NewRow
                        drS("BudgetYear") = Me.GetBudgetYear()
                        drS("PeriodType") = Me.GetPeriodType()
                        drS("RevNo") = "1"
                        drS("ProjectNo") = Me.GetProjectNo()
                        drS("BudgetOrder") = CStr(row.Cells("OrderNo3").Value)
                        drS("Status") = CStr(enumBudgetStatus.NewRecord)
                        dtSave.Rows.Add(drS)
                    End If
                Next
            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then   '// MTP Budget        
                For Each row As DataGridViewRow In grvBudget4.Rows
                    Dim isChecked As Boolean = DirectCast(grvBudget4(0, row.Index).Value, Boolean)
                    If isChecked Then
                        drS = dtSave.NewRow
                        drS("BudgetYear") = Me.GetBudgetYear()
                        drS("PeriodType") = Me.GetPeriodType()
                        drS("RevNo") = "1"
                        drS("ProjectNo") = Me.GetProjectNo()
                        drS("BudgetOrder") = CStr(row.Cells("OrderNo4").Value)
                        drS("Status") = CStr(enumBudgetStatus.NewRecord)
                        dtSave.Rows.Add(drS)
                    End If
                Next
            End If

            Dim conn As SqlConnection = Nothing
            Dim trans As SqlTransaction
            Dim success As Boolean = False

            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()
            trans = conn.BeginTransaction()

            Try
                myClsBG0200BL.dtSave = dtSave
                If myClsBG0200BL.SaveBudgetDataReInput(conn, trans) = True Then
                    success = True
                Else
                    success = False
                End If

                If success Then
                    trans.Commit()
                Else
                    trans.Rollback()
                End If
            Catch ex As Exception
                Throw ex
            End Try

            If success = True Then
                MessageBox.Show("Budget journal was rejected for re-input by Budget Order No and set as new record", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                '// Write Trans Log
                myClsBG0200BL.LogOperationCd = enumOperationCd.ReInputByOrder
                myClsBG0200BL.WriteTransLog()

                '// Send auto mail
                If p_blnSendAutoMail Then
                    myClsBG0200BL.UserPIC = CStr(cboPIC.SelectedValue)
                    myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
                    myClsBG0200BL.PeriodType = Me.GetPeriodType()
                    myClsBG0200BL.BudgetType = Me.GetBudgetType()
                    myClsBG0200BL.OperationCd = enumOperationCd.RejectBudget2
                    myClsBG0200BL.ProjectNo = Me.GetProjectNo()

                    myClsBG0200BL.SendAutoMail()
                End If

                '// Refresh side menu
                p_frmBG0010.ShowBudgetMenu()

                myForceCloseFlg = True
                Me.Close()
            Else
                MessageBox.Show("Can not reject for re-input by Budget Order No  Budget journal", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            Timer1.Enabled = False

            '// Reselect Combobox
            If cboPIC.Items.Count > 0 Then
                Dim i As Integer = cboPIC.SelectedIndex
                cboPIC.SelectedIndex = -1
                cboPIC.SelectedIndex = i
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function DeleteReInputByOrder() As Boolean
        Try
            '// Approve Budget journal
            '// Set Parameters
            myClsBG0200BL.BudgetYear = Me.GetBudgetYear()
            myClsBG0200BL.PeriodType = Me.GetPeriodType()
            myClsBG0200BL.BudgetType = Me.GetBudgetType()
            myClsBG0200BL.UserId = p_strUserId
            myClsBG0200BL.UserPIC = CStr(cboPIC.SelectedValue)
            myClsBG0200BL.ProjectNo = Me.GetProjectNo()

            '// Call Function
            Dim blnSave As Boolean = False
            If blnReInputByOrder = True Then
                If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then          '// Original Budget
                    myClsBG0200BL.BudgetList = CType(grvBudget1.DataSource, DataTable)
                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then      '// Estimate Budget
                    myClsBG0200BL.BudgetList = CType(grvBudget2.DataSource, DataTable)
                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then        '// Revise Budget
                    myClsBG0200BL.BudgetList = CType(grvBudget3.DataSource, DataTable)
                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then         '// MTP Budget
                    myClsBG0200BL.BudgetList = CType(grvBudget4.DataSource, DataTable)
                End If

                Dim CountSel As Integer = 0
                Dim drS As DataRow
                Dim dtSave As New DataTable
                dtSave.Columns.Add("BudgetYear", GetType(String))
                dtSave.Columns.Add("PeriodType", GetType(String))
                dtSave.Columns.Add("RevNo", GetType(String))
                dtSave.Columns.Add("ProjectNo", GetType(String))
                dtSave.Columns.Add("BudgetOrder", GetType(String))
                dtSave.Columns.Add("Status", GetType(String))
                If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then   '// Original Budget
                    For Each row As DataGridViewRow In grvBudget1.Rows
                        Dim isChecked As Boolean = DirectCast(grvBudget1(0, row.Index).Value, Boolean)

                        drS = dtSave.NewRow
                        drS("BudgetYear") = Me.GetBudgetYear()
                        drS("PeriodType") = Me.GetPeriodType()
                        drS("RevNo") = Me.CurrRevNo
                        drS("ProjectNo") = Me.GetProjectNo()
                        drS("BudgetOrder") = CStr(row.Cells("OrderNo1").Value)
                        drS("Status") = CStr(enumBudgetStatus.Submit)
                        dtSave.Rows.Add(drS)
                    Next
                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then   '// Estimate Budget
                    For Each row As DataGridViewRow In grvBudget2.Rows
                        Dim isChecked As Boolean = DirectCast(grvBudget2(0, row.Index).Value, Boolean)
                        drS = dtSave.NewRow
                        drS("BudgetYear") = Me.GetBudgetYear()
                        drS("PeriodType") = Me.GetPeriodType()
                        drS("RevNo") = Me.CurrRevNo
                        drS("ProjectNo") = Me.GetProjectNo()
                        drS("BudgetOrder") = CStr(row.Cells("OrderNo2").Value)
                        drS("Status") = CStr(enumBudgetStatus.Submit)
                        dtSave.Rows.Add(drS)
                    Next
                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then   '// Revise Budget        
                    For Each row As DataGridViewRow In grvBudget3.Rows
                        Dim isChecked As Boolean = DirectCast(grvBudget3(0, row.Index).Value, Boolean)
                        drS = dtSave.NewRow
                        drS("BudgetYear") = Me.GetBudgetYear()
                        drS("PeriodType") = Me.GetPeriodType()
                        drS("RevNo") = Me.CurrRevNo
                        drS("ProjectNo") = Me.GetProjectNo()
                        drS("BudgetOrder") = CStr(row.Cells("OrderNo3").Value)
                        drS("Status") = CStr(enumBudgetStatus.Submit)
                        dtSave.Rows.Add(drS)
                    Next
                ElseIf Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then   '// MTP Budget        
                    For Each row As DataGridViewRow In grvBudget4.Rows
                        Dim isChecked As Boolean = DirectCast(grvBudget4(0, row.Index).Value, Boolean)
                        drS = dtSave.NewRow
                        drS("BudgetYear") = Me.GetBudgetYear()
                        drS("PeriodType") = Me.GetPeriodType()
                        drS("RevNo") = Me.CurrRevNo
                        drS("ProjectNo") = Me.GetProjectNo()
                        drS("BudgetOrder") = CStr(row.Cells("OrderNo4").Value)
                        drS("Status") = CStr(enumBudgetStatus.Submit)
                        dtSave.Rows.Add(drS)
                    Next
                End If

                Dim conn As SqlConnection = Nothing
                Dim trans As SqlTransaction
                Dim success As Boolean = False

                conn = New SqlConnection(My.Settings.ConnStr)
                conn.Open()
                trans = conn.BeginTransaction()
                Try
                    myClsBG0200BL.dtSave = dtSave
                    If myClsBG0200BL.DeleteBudgetDataReInputByOrder(conn, trans) = True Then
                        success = True
                    Else
                        success = False
                    End If

                    If success Then
                        trans.Commit()

                        '// Write Transaction Log
                        WriteTransactionLog(CStr(enumOperationCd.AdjustBudgetDirectInput), "", "", "", "", "", "")

                        '// Refresh side menu
                        p_frmBG0010.ShowBudgetMenu()

                        myForceCloseFlg = True
                        Me.Close()
                    Else
                        trans.Rollback()
                    End If
                Catch ex As Exception
                    Throw ex
                End Try
            Else
                blnSave = myClsBG0200BL.SaveApproveBudgetData()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "BUDGET_COMMENT"
    Private Sub ShowCommentPopup(ByVal intNum As Integer, ByVal budgetOderSub As String())
        Dim budgetOrderNo As String = String.Empty
        Try
            p_frmBG0201 = New frmBG0201

            p_frmBG0201.BudgetYear = Me.GetBudgetYear()
            p_frmBG0201.PeriodType = Me.GetPeriodType()

            If Not budgetOderSub Is Nothing AndAlso budgetOderSub.Length > 0 Then
                budgetOrderNo = budgetOderSub(0).ToString
            End If
            p_frmBG0201.BudgetOrderNo = budgetOrderNo
            p_frmBG0201.RevNo = Me.CurrRevNo
            p_frmBG0201.ProjectNo = Me.GetProjectNo()
            If Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then
                p_frmBG0201.RRTNo = CStr(intNum)
            Else
                p_frmBG0201.MonthNo = CStr(intNum)
            End If

            p_frmBG0201.OperationCd = CStr(Me.OperationCd)

            If p_frmBG0201.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then

            End If

            If Me.OperationCd = enumOperationCd.InputBudget Or _
                     (Me.OperationCd = enumOperationCd.AdjustBudget And CInt(Me.CurrRevNo) > 1) Or _
                     (Me.OperationCd = enumOperationCd.AdjustBudgetDirectInput And CInt(Me.CurrRevNo) > 1) Then
                HighlightWorkingBGAndComment()
            End If

            p_frmBG0201.Dispose()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub grvBudget1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grvBudget1.CellDoubleClick
        'Original Budget
        Dim intMonth As Integer
        Dim ColumnName As String
        Dim blnFound As Boolean = False

        Try
            ColumnName = CStr(grvBudget1.Columns(e.ColumnIndex).Name)

            Select Case ColumnName
                Case "g1col8"
                    intMonth = 1
                    blnFound = True
                Case "g1col9"
                    intMonth = 2
                    blnFound = True
                Case "g1col10"
                    intMonth = 3
                    blnFound = True
                Case "g1col11"
                    intMonth = 4
                    blnFound = True
                Case "g1col12"
                    intMonth = 5
                    blnFound = True
                Case "g1col13"
                    intMonth = 6
                    blnFound = True
                Case "g1colex1"
                    intMonth = 7
                    blnFound = True
                Case "g1colex2"
                    intMonth = 8
                    blnFound = True
                Case "g1colex3"
                    intMonth = 9
                    blnFound = True
                Case "g1colex4"
                    intMonth = 10
                    blnFound = True
                Case "g1colex5"
                    intMonth = 11
                    blnFound = True
                Case "g1colex6"
                    intMonth = 12
                    blnFound = True
            End Select

            If blnFound = True Then
                Dim budgetOderSub() As String
                If e.RowIndex >= 0 Then
                    budgetOderSub = grvBudget1.Rows(e.RowIndex).Cells(2).Value.ToString.Split(CChar(" :"))
                    ShowCommentPopup(intMonth, budgetOderSub)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub grvBudget2_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grvBudget2.CellDoubleClick
        'Estimate Budget
        Dim intMonth As Integer
        Dim ColumnName As String
        Dim blnFound As Boolean = False

        Try
            ColumnName = CStr(grvBudget2.Columns(e.ColumnIndex).Name)

            Select Case ColumnName
                Case "g2col11"
                    intMonth = 10
                    blnFound = True
                Case "g2col12"
                    intMonth = 11
                    blnFound = True
                Case "g2col13"
                    intMonth = 12
                    blnFound = True
            End Select

            If blnFound = True Then
                Dim budgetOderSub() As String

                If e.RowIndex >= 0 Then
                    budgetOderSub = grvBudget2.Rows(e.RowIndex).Cells(2).Value.ToString.Split(CChar(" :"))
                    ShowCommentPopup(intMonth, budgetOderSub)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub grvBudget3_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grvBudget3.CellDoubleClick
        'Revise Budget
        Dim intMonth As Integer
        Dim ColumnName As String
        Dim blnFound As Boolean = False

        Try
            ColumnName = CStr(grvBudget3.Columns(e.ColumnIndex).Name)

            Select Case ColumnName
                Case "g3col10"
                    intMonth = 4
                    blnFound = True
                Case "g3col11"
                    intMonth = 5
                    blnFound = True
                Case "g3col12"
                    intMonth = 6
                    blnFound = True
                Case "g3col16"
                    intMonth = 7
                    blnFound = True
                Case "g3col17"
                    intMonth = 8
                    blnFound = True
                Case "g3col18"
                    intMonth = 9
                    blnFound = True
                Case "g3col19"
                    intMonth = 10
                    blnFound = True
                Case "g3col20"
                    intMonth = 11
                    blnFound = True
                Case "g3col21"
                    intMonth = 12
                    blnFound = True
            End Select

            If blnFound = True Then
                Dim budgetOderSub() As String

                If e.RowIndex >= 0 Then
                    budgetOderSub = grvBudget3.Rows(e.RowIndex).Cells(2).Value.ToString.Split(CChar(" :"))
                    ShowCommentPopup(intMonth, budgetOderSub)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub grvBudget4_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grvBudget4.CellDoubleClick
        Dim intRRT As Integer
        Dim ColumnName As String
        Dim blnFound As Boolean = False

        Try
            ColumnName = CStr(grvBudget4.Columns(e.ColumnIndex).Name)

            Select Case ColumnName
                Case "g4col9"
                    intRRT = 2
                    blnFound = True
                Case "g4col11"
                    intRRT = 3
                    blnFound = True
            End Select

            If blnFound = True Then
                Dim budgetOderSub() As String

                If e.RowIndex >= 0 Then
                    budgetOderSub = grvBudget4.Rows(e.RowIndex).Cells(2).Value.ToString.Split(CChar(" :"))
                    ShowCommentPopup(intRRT, budgetOderSub)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function GetComment(ByVal strOrderNo As String) As DataTable
        Dim result As DataTable = Nothing

        Try
            myClsBG0201BL.BudgetYear = Me.GetBudgetYear()
            myClsBG0201BL.PeriodType = Me.GetPeriodType()
            myClsBG0201BL.BudgetOrderNo = strOrderNo
            myClsBG0201BL.RevNo = Me.CurrRevNo
            myClsBG0201BL.ProjectNo = Me.GetProjectNo()

            If myClsBG0201BL.SearchComment AndAlso myClsBG0201BL.CommentList.Rows.Count > 0 Then
                result = myClsBG0201BL.CommentList
            End If
        Catch ex As Exception
            Throw ex
        End Try

        Return result
    End Function


    Private Sub HighlightWorkingBGAndComment()
        Try
            Debug.Print(Now.ToString() & ": Begin HighlightWorkingBGAndComment")
            Dim dt As DataTable = Nothing

            Dim strOrderNo As String
            '' // Hightlight Comment
            If Me.GetPeriodType() = CStr(enumPeriodType.OriginalBudget) Then '// Original Budget
                For i = 0 To grvBudget1.RowCount - 1
                    strOrderNo = CStr(grvBudget1.Item("OrderNo1", i).Value)
                    dt = GetComment(strOrderNo)

                    If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                        If i Mod 2 = 0 Then
                            If dt.Rows(0).Item("M1").ToString <> "" Then
                                grvBudget1.Item("g1col8", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget1.Item("g1col8", i).Style.BackColor = grvBudget1.Columns("g1Col8").DefaultCellStyle.BackColor
                            End If

                            If dt.Rows(0).Item("M2").ToString <> "" Then
                                grvBudget1.Item("g1col9", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget1.Item("g1col9", i).Style.BackColor = grvBudget1.Columns("g1Col9").DefaultCellStyle.BackColor
                            End If

                            If dt.Rows(0).Item("M3").ToString <> "" Then
                                grvBudget1.Item("g1col10", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget1.Item("g1col10", i).Style.BackColor = grvBudget1.Columns("g1Col10").DefaultCellStyle.BackColor
                            End If

                            If dt.Rows(0).Item("M4").ToString <> "" Then
                                grvBudget1.Item("g1col11", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget1.Item("g1col11", i).Style.BackColor = grvBudget1.Columns("g1Col11").DefaultCellStyle.BackColor
                            End If

                            If dt.Rows(0).Item("M5").ToString <> "" Then
                                grvBudget1.Item("g1col12", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget1.Item("g1col12", i).Style.BackColor = grvBudget1.Columns("g1Col12").DefaultCellStyle.BackColor
                            End If

                            If dt.Rows(0).Item("M6").ToString <> "" Then
                                grvBudget1.Item("g1col13", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget1.Item("g1col13", i).Style.BackColor = grvBudget1.Columns("g1col13").DefaultCellStyle.BackColor
                            End If

                            If dt.Rows(0).Item("M7").ToString <> "" Then
                                grvBudget1.Item("g1colex1", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget1.Item("g1colex1", i).Style.BackColor = grvBudget1.Columns("g1colex1").DefaultCellStyle.BackColor
                            End If

                            If dt.Rows(0).Item("M8").ToString <> "" Then
                                grvBudget1.Item("g1colex2", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget1.Item("g1colex2", i).Style.BackColor = grvBudget1.Columns("g1colex2").DefaultCellStyle.BackColor
                            End If

                            If dt.Rows(0).Item("M9").ToString <> "" Then
                                grvBudget1.Item("g1colex3", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget1.Item("g1colex3", i).Style.BackColor = grvBudget1.Columns("g1colex3").DefaultCellStyle.BackColor
                            End If

                            If dt.Rows(0).Item("M10").ToString <> "" Then
                                grvBudget1.Item("g1colex4", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget1.Item("g1colex4", i).Style.BackColor = grvBudget1.Columns("g1colex4").DefaultCellStyle.BackColor
                            End If

                            If dt.Rows(0).Item("M11").ToString <> "" Then
                                grvBudget1.Item("g1colex5", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget1.Item("g1colex5", i).Style.BackColor = grvBudget1.Columns("g1colex5").DefaultCellStyle.BackColor
                            End If

                            If dt.Rows(0).Item("M12").ToString <> "" Then
                                grvBudget1.Item("g1colex6", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget1.Item("g1colex6", i).Style.BackColor = grvBudget1.Columns("g1colex6").DefaultCellStyle.BackColor
                            End If

                        Else

                            If dt.Rows(0).Item("M1").ToString <> "" Then
                                grvBudget1.Item("g1col8", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget1.Item("g1col8", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            End If

                            If dt.Rows(0).Item("M2").ToString <> "" Then
                                grvBudget1.Item("g1col9", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget1.Item("g1col9", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            End If

                            If dt.Rows(0).Item("M3").ToString <> "" Then
                                grvBudget1.Item("g1col10", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget1.Item("g1col10", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            End If

                            If dt.Rows(0).Item("M4").ToString <> "" Then
                                grvBudget1.Item("g1col11", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget1.Item("g1col11", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            End If

                            If dt.Rows(0).Item("M5").ToString <> "" Then
                                grvBudget1.Item("g1col12", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget1.Item("g1col12", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            End If

                            If dt.Rows(0).Item("M6").ToString <> "" Then
                                grvBudget1.Item("g1col13", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget1.Item("g1col13", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            End If

                            If dt.Rows(0).Item("M7").ToString <> "" Then
                                grvBudget1.Item("g1colex1", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget1.Item("g1colex1", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            End If

                            If dt.Rows(0).Item("M8").ToString <> "" Then
                                grvBudget1.Item("g1colex2", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget1.Item("g1colex2", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            End If

                            If dt.Rows(0).Item("M9").ToString <> "" Then
                                grvBudget1.Item("g1colex3", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget1.Item("g1colex3", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            End If

                            If dt.Rows(0).Item("M10").ToString <> "" Then
                                grvBudget1.Item("g1colex4", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget1.Item("g1colex4", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            End If

                            If dt.Rows(0).Item("M11").ToString <> "" Then
                                grvBudget1.Item("g1colex5", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget1.Item("g1colex5", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            End If

                            If dt.Rows(0).Item("M12").ToString <> "" Then
                                grvBudget1.Item("g1colex6", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget1.Item("g1colex6", i).Style.BackColor = grvBudget1.AlternatingRowsDefaultCellStyle.BackColor
                            End If

                        End If
                    End If

                Next

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.EstimateBudget) Then '// EstimateBudget

                For i = 0 To grvBudget2.RowCount - 1
                    strOrderNo = CStr(grvBudget2.Item("OrderNo2", i).Value)
                    dt = GetComment(strOrderNo)
                    If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                        If i Mod 2 = 0 Then
                            If dt.Rows(0).Item("M10").ToString <> "" Then
                                grvBudget2.Item("g2col11", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget2.Item("g2col11", i).Style.BackColor = grvBudget2.Columns("g2col11").DefaultCellStyle.BackColor
                            End If
                            If dt.Rows(0).Item("M11").ToString <> "" Then
                                grvBudget2.Item("g2col12", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget2.Item("g2col12", i).Style.BackColor = grvBudget2.Columns("g2col12").DefaultCellStyle.BackColor
                            End If
                            If dt.Rows(0).Item("M12").ToString <> "" Then
                                grvBudget2.Item("g2col13", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget2.Item("g2col13", i).Style.BackColor = grvBudget2.Columns("g2col13").DefaultCellStyle.BackColor
                            End If
                        Else
                            If dt.Rows(0).Item("M10").ToString <> "" Then
                                grvBudget2.Item("g2col11", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget2.Item("g2col11", i).Style.BackColor = grvBudget2.AlternatingRowsDefaultCellStyle.BackColor
                            End If
                            If dt.Rows(0).Item("M11").ToString <> "" Then
                                grvBudget2.Item("g2col12", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget2.Item("g2col12", i).Style.BackColor = grvBudget2.AlternatingRowsDefaultCellStyle.BackColor
                            End If
                            If dt.Rows(0).Item("M12").ToString <> "" Then
                                grvBudget2.Item("g2col13", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget2.Item("g2col13", i).Style.BackColor = grvBudget2.AlternatingRowsDefaultCellStyle.BackColor
                            End If
                        End If
                    End If

                Next

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.ReviseBudget) Then    '// Revise Budget

                For i = 0 To grvBudget3.RowCount - 1
                    strOrderNo = CStr(grvBudget3.Item("OrderNo3", i).Value)
                    dt = GetComment(strOrderNo)
                    If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                        If i Mod 2 = 0 Then
                            If dt.Rows(0).Item("M4").ToString <> "" Then
                                grvBudget3.Item("g3col10", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget3.Item("g3col10", i).Style.BackColor = grvBudget3.Columns("g3col10").DefaultCellStyle.BackColor
                            End If
                            If dt.Rows(0).Item("M5").ToString <> "" Then
                                grvBudget3.Item("g3col11", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget3.Item("g3col11", i).Style.BackColor = grvBudget3.Columns("g3col11").DefaultCellStyle.BackColor
                            End If
                            If dt.Rows(0).Item("M6").ToString <> "" Then
                                grvBudget3.Item("g3col12", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget3.Item("g3col12", i).Style.BackColor = grvBudget3.Columns("g3col12").DefaultCellStyle.BackColor
                            End If
                            If dt.Rows(0).Item("M7").ToString <> "" Then
                                grvBudget3.Item("g3col16", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget3.Item("g3col16", i).Style.BackColor = grvBudget3.Columns("g3col16").DefaultCellStyle.BackColor
                            End If
                            If dt.Rows(0).Item("M8").ToString <> "" Then
                                grvBudget3.Item("g3col17", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget3.Item("g3col17", i).Style.BackColor = grvBudget3.Columns("g3col17").DefaultCellStyle.BackColor
                            End If
                            If dt.Rows(0).Item("M9").ToString <> "" Then
                                grvBudget3.Item("g3col18", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget3.Item("g3col18", i).Style.BackColor = grvBudget3.Columns("g3col18").DefaultCellStyle.BackColor
                            End If
                            If dt.Rows(0).Item("M10").ToString <> "" Then
                                grvBudget3.Item("g3col19", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget3.Item("g3col19", i).Style.BackColor = grvBudget3.Columns("g3col19").DefaultCellStyle.BackColor
                            End If
                            If dt.Rows(0).Item("M11").ToString <> "" Then
                                grvBudget3.Item("g3col20", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget3.Item("g3col20", i).Style.BackColor = grvBudget3.Columns("g3col20").DefaultCellStyle.BackColor
                            End If
                            If dt.Rows(0).Item("M12").ToString <> "" Then
                                grvBudget3.Item("g3col21", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget3.Item("g3col21", i).Style.BackColor = grvBudget3.Columns("g3col21").DefaultCellStyle.BackColor
                            End If
                        Else
                            If dt.Rows(0).Item("M4").ToString <> "" Then
                                grvBudget3.Item("g3col10", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget3.Item("g3col10", i).Style.BackColor = grvBudget3.AlternatingRowsDefaultCellStyle.BackColor
                            End If
                            If dt.Rows(0).Item("M5").ToString <> "" Then
                                grvBudget3.Item("g3col11", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget3.Item("g3col11", i).Style.BackColor = grvBudget3.AlternatingRowsDefaultCellStyle.BackColor
                            End If
                            If dt.Rows(0).Item("M6").ToString <> "" Then
                                grvBudget3.Item("g3col12", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget3.Item("g3col12", i).Style.BackColor = grvBudget3.AlternatingRowsDefaultCellStyle.BackColor
                            End If
                            If dt.Rows(0).Item("M7").ToString <> "" Then
                                grvBudget3.Item("g3col16", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget3.Item("g3col16", i).Style.BackColor = grvBudget3.AlternatingRowsDefaultCellStyle.BackColor
                            End If
                            If dt.Rows(0).Item("M8").ToString <> "" Then
                                grvBudget3.Item("g3col17", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget3.Item("g3col17", i).Style.BackColor = grvBudget3.AlternatingRowsDefaultCellStyle.BackColor
                            End If
                            If dt.Rows(0).Item("M9").ToString <> "" Then
                                grvBudget3.Item("g3col18", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget3.Item("g3col18", i).Style.BackColor = grvBudget3.AlternatingRowsDefaultCellStyle.BackColor
                            End If
                            If dt.Rows(0).Item("M10").ToString <> "" Then
                                grvBudget3.Item("g3col19", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget3.Item("g3col19", i).Style.BackColor = grvBudget3.AlternatingRowsDefaultCellStyle.BackColor
                            End If
                            If dt.Rows(0).Item("M11").ToString <> "" Then
                                grvBudget3.Item("g3col20", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget3.Item("g3col20", i).Style.BackColor = grvBudget3.AlternatingRowsDefaultCellStyle.BackColor
                            End If
                            If dt.Rows(0).Item("M12").ToString <> "" Then
                                grvBudget3.Item("g3col21", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget3.Item("g3col21", i).Style.BackColor = grvBudget3.AlternatingRowsDefaultCellStyle.BackColor
                            End If
                        End If
                    End If
                Next

            ElseIf Me.GetPeriodType() = CStr(enumPeriodType.MTPBudget) Then    '// MTP Budget
                For i = 0 To grvBudget4.RowCount - 1
                    strOrderNo = CStr(grvBudget4.Item("OrderNo4", i).Value)
                    dt = GetComment(strOrderNo)
                    If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                        If i Mod 2 = 0 Then
                            If dt.Rows(0).Item("RRT2").ToString <> "" Then
                                grvBudget4.Item("g4col9", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget4.Item("g4col9", i).Style.BackColor = grvBudget4.Columns("g4col9").DefaultCellStyle.BackColor
                            End If
                            If dt.Rows(0).Item("RRT3").ToString <> "" Then
                                grvBudget4.Item("g4col11", i).Style.BackColor = System.Drawing.Color.FromArgb(250, 233, 6) 'RGB(64, 221, 242)
                            Else
                                grvBudget4.Item("g4col11", i).Style.BackColor = grvBudget4.Columns("g4col11").DefaultCellStyle.BackColor
                            End If
                        Else
                            If dt.Rows(0).Item("RRT2").ToString <> "" Then
                                grvBudget4.Item("g4col9", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget4.Item("g4col9", i).Style.BackColor = grvBudget4.AlternatingRowsDefaultCellStyle.BackColor
                            End If
                            If dt.Rows(0).Item("RRT3").ToString <> "" Then
                                grvBudget4.Item("g4col11", i).Style.BackColor = System.Drawing.Color.FromArgb(64, 221, 242)  'Color.BlueViolet
                            Else
                                grvBudget4.Item("g4col11", i).Style.BackColor = grvBudget4.AlternatingRowsDefaultCellStyle.BackColor

                            End If
                        End If
                    End If
                Next
            End If

            Debug.Print(Now.ToString() & ": End HighlightWorkingBGAndComment")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class