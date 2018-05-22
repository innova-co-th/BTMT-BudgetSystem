Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Data
Imports System.Drawing.Printing

Public Class frmBG0400

#Region "Variable"
    Private m_Report As ReportDocument
    Private m_reportName As String = String.Empty
    Private m_DS As DataSet = Nothing
    Private myPIC As String = String.Empty
    Private strBudgetYear As String = String.Empty
    Private strProjectNo As String = String.Empty
    Private bParamPersonInCharge As Boolean = False
    Private myAccountNo As String = String.Empty
    Private m_strReportType As String = String.Empty
    Private myMonth As String = String.Empty
    Private myPeriod As String = String.Empty
    Private myBudgetStatus As Integer = 0
    Private myWorkingBG1 As Single = 0
    Private myWorkingBG2 As Single = 0
    Private myMTP_SUM1 As Decimal = 0
    Private myMTP_SUM2 As Decimal = 0
    Private myMTP_SUM3 As Decimal = 0
    Private myMTP_SUM4 As Decimal = 0
    Private myMTP_SUM5 As Decimal = 0
    Private myMTPBudget As Boolean = False
    'Private WithEvents ts As New ToolStrip
    'Private WithEvents tsbtn As New ToolStripButton
#End Region

#Region "Property"

    Public Property ReportName() As String
        Get
            Return m_reportName
        End Get
        Set(ByVal value As String)
            m_reportName = value
        End Set
    End Property

    Public Property DS() As DataSet
        Get
            Return m_DS
        End Get
        Set(ByVal value As DataSet)
            m_DS = value
        End Set
    End Property

    Public Property PIC() As String
        Get
            Return myPIC
        End Get
        Set(ByVal value As String)
            myPIC = value
        End Set
    End Property

    Public Property BudgetYear() As String
        Get
            Return strBudgetYear
        End Get
        Set(ByVal value As String)
            strBudgetYear = value
        End Set
    End Property

    Public Property ProjectNo() As String
        Get
            Return strProjectNo
        End Get
        Set(ByVal value As String)
            strProjectNo = value
        End Set
    End Property

    Public Property ParamPersonInCharge() As Boolean
        Get
            Return bParamPersonInCharge
        End Get
        Set(ByVal value As Boolean)
            bParamPersonInCharge = value
        End Set
    End Property

    Public Property AccountNo() As String
        Get
            Return myAccountNo
        End Get
        Set(ByVal value As String)
            myAccountNo = value
        End Set
    End Property

    Public Property ReportType() As String
        Get
            Return m_strReportType
        End Get
        Set(ByVal value As String)
            m_strReportType = value
        End Set
    End Property

    Public Property Month() As String
        Get
            Return myMonth
        End Get
        Set(ByVal value As String)
            myMonth = value
        End Set
    End Property

    Public Property Period() As String
        Get
            Return myPeriod
        End Get
        Set(ByVal value As String)
            myPeriod = value
        End Set
    End Property

    Public Property BudgetStatus() As Integer
        Get
            Return myBudgetStatus
        End Get
        Set(ByVal value As Integer)
            myBudgetStatus = value
        End Set
    End Property

    Public Property WorkingBG1() As Single
        Get
            Return myWorkingBG1
        End Get
        Set(ByVal value As Single)
            myWorkingBG1 = value
        End Set
    End Property
    Public Property WorkingBG2() As Single
        Get
            Return myWorkingBG2
        End Get
        Set(ByVal value As Single)
            myWorkingBG2 = value
        End Set
    End Property

    Public Property MTP_SUM1() As Decimal
        Get
            Return myMTP_SUM1
        End Get
        Set(ByVal value As Decimal)
            myMTP_SUM1 = value
        End Set
    End Property
    Public Property MTP_SUM2() As Decimal
        Get
            Return myMTP_SUM2
        End Get
        Set(ByVal value As Decimal)
            myMTP_SUM2 = value
        End Set
    End Property
    Public Property MTP_SUM3() As Decimal
        Get
            Return myMTP_SUM3
        End Get
        Set(ByVal value As Decimal)
            myMTP_SUM3 = value
        End Set
    End Property
    Public Property MTP_SUM4() As Decimal
        Get
            Return myMTP_SUM4
        End Get
        Set(ByVal value As Decimal)
            myMTP_SUM4 = value
        End Set
    End Property
    Public Property MTP_SUM5() As Decimal
        Get
            Return myMTP_SUM5
        End Get
        Set(ByVal value As Decimal)
            myMTP_SUM5 = value
        End Set
    End Property

    Public Property MTPBudget() As Boolean
        Get
            Return myMTPBudget
        End Get
        Set(ByVal value As Boolean)
            myMTPBudget = value
        End Set
    End Property

#End Region

#Region "Overrides Function"
    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

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

    Public Sub ConfigureCrystalReports()

        Try
            m_Report = New ReportDocument()
            Dim reportPath As String = p_strAppPath & "\Reports\" & ReportName
            m_Report.Load(reportPath)
            m_Report.SetDataSource(DS)

            'If Me.ParamPersonInCharge = True Then
            '    SetupPICReportParameter()
            'Else
            '    SetupAccountReportParameter()
            'End If
            If ReportType <> "BudgetCompare" Then
                SetupBudgetStatus()
            End If

            Select Case ReportType
                Case "DetailByPersonInCharge"
                    SetupPICReportParameter()

                Case "DetailByAccountCode"
                    SetupAccountReportParameter()

                Case "SummaryByApplicant"
                    SetupSummaryByApplicantReportParameter()

                Case "SummaryByPersonInCharge"
                    SetupSummarybyPICParameter()

                Case "SummarybyInvestment"
                    SetupSummarybyInvestmentReportParameter()

                Case "SummaryByAccountNoReport"
                    SetupSummaryByAccountNoReportParameter()

                Case "BudgetCompare"
                    SetupBudgetCompareReportParameter()

                Case "CommentByPersonInCharge"
                    SetupCommentByPICReportParameter()

            End Select

            '  m_Report.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperA3

            Me.CrystalReportViewer1.ReportSource = m_Report

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Budget Report", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try

    End Sub

    Public Sub SetupBudgetStatus()

        If Me.BudgetStatus >= 5 Then
            m_Report.ReportDefinition.ReportObjects("picAuth10").ObjectFormat.EnableSuppress = False
        Else
            m_Report.ReportDefinition.ReportObjects("picAuth10").ObjectFormat.EnableSuppress = True
        End If

        If Me.BudgetStatus >= 6 Then
            m_Report.ReportDefinition.ReportObjects("picAuth3").ObjectFormat.EnableSuppress = False
        Else
            m_Report.ReportDefinition.ReportObjects("picAuth3").ObjectFormat.EnableSuppress = True
        End If

    End Sub

    Public Sub SetupPICReportParameter()

        m_Report.SetParameterValue("PERSON_IN_CHARGE_NM", Me.PIC)
        m_Report.SetParameterValue("BUDGET_YEAR", Me.BudgetYear)
        m_Report.SetParameterValue("HALF_BUDGET_YEAR", Me.BudgetYear.Substring(2, 2))
        m_Report.SetParameterValue("PERIOD", Me.Period)
        m_Report.SetParameterValue("PROJECT_NO", Me.ProjectNo)

        'Dim intYear As Integer = CInt(Me.BudgetYear)
        'intYear = intYear - 1
        'Dim strHalfLastYear As String = intYear.ToString.Substring(2, 2)
        'm_Report.SetParameterValue("HALF_LAST_YEAR", strHalfLastYear)

    End Sub

    Public Sub SetupBudgetCompareReportParameter()

        'm_Report.SetParameterValue("PERSON_IN_CHARGE_NM", Me.PIC)
        m_Report.SetParameterValue("BUDGET_YEAR", Me.BudgetYear)
        m_Report.SetParameterValue("MONTH", Me.Month)
        m_Report.SetParameterValue("HALF_BUDGET_YEAR", Me.BudgetYear.Substring(2, 2))

    End Sub

    Public Sub SetupAccountReportParameter()

        m_Report.SetParameterValue("ACCOUNT_NO", Me.AccountNo)
        m_Report.SetParameterValue("BUDGET_YEAR", Me.BudgetYear)
        m_Report.SetParameterValue("HALF_BUDGET_YEAR", Me.BudgetYear.Substring(2, 2))
        m_Report.SetParameterValue("FC_COST", enumCost.FC)
        m_Report.SetParameterValue("ADMIN_COST", enumCost.ADMIN)
        m_Report.SetParameterValue("PROJECT_NO", Me.ProjectNo)
    End Sub

    Public Sub SetupSummaryByApplicantReportParameter()

        m_Report.SetParameterValue("BUDGET_YEAR", Me.BudgetYear)
        m_Report.SetParameterValue("HALF_BUDGET_YEAR", Me.BudgetYear.Substring(2, 2))
        m_Report.SetParameterValue("PROJECT_NO", Me.ProjectNo)

    End Sub

    Public Sub SetupSummaryByAccountNoReportParameter()

        m_Report.SetParameterValue("BUDGET_YEAR", Me.BudgetYear)
        m_Report.SetParameterValue("HALF_BUDGET_YEAR", Me.BudgetYear.Substring(2, 2))
        m_Report.SetParameterValue("FC_COST", enumCost.FC)
        m_Report.SetParameterValue("ADMIN_COST", enumCost.ADMIN)
        m_Report.SetParameterValue("PROJECT_NO", Me.ProjectNo)
        'm_Report.SetParameterValue("WORKING_BG1", WorkingBG1)
        'm_Report.SetParameterValue("WORKING_BG2", WorkingBG1)
        If MTPBudget = True Then
            m_Report.SetParameterValue("MTP_SUM1", Me.MTP_SUM1)
            m_Report.SetParameterValue("MTP_SUM2", Me.MTP_SUM2)
            m_Report.SetParameterValue("MTP_SUM3", Me.MTP_SUM3)
            m_Report.SetParameterValue("MTP_SUM4", Me.MTP_SUM4)
            m_Report.SetParameterValue("MTP_SUM5", Me.MTP_SUM5)
        End If
    End Sub

    Public Sub SetupSummarybyPICParameter()
        m_Report.SetParameterValue("PERIOD", Me.Period)
        'm_Report.SetParameterValue("MONTH", Me.Month)
        m_Report.SetParameterValue("BUDGET_YEAR", Me.BudgetYear)
        m_Report.SetParameterValue("HALF_BUDGET_YEAR", Me.BudgetYear.Substring(2, 2))
        m_Report.SetParameterValue("PROJECT_NO", Me.ProjectNo)
    End Sub

    Public Sub SetupSummarybyInvestmentReportParameter()
        m_Report.SetParameterValue("PERIOD", Me.Period)
        m_Report.SetParameterValue("BUDGET_YEAR", Me.BudgetYear)
        m_Report.SetParameterValue("HALF_BUDGET_YEAR", Me.BudgetYear.Substring(2, 2))
        m_Report.SetParameterValue("PROJECT_NO", Me.ProjectNo)
    End Sub

    Public Sub SetupCommentByPICReportParameter()
        m_Report.SetParameterValue("PERIOD", Me.Period)
        ''m_Report.SetParameterValue("MONTH", Me.Month)
        m_Report.SetParameterValue("BUDGET_YEAR", Me.BudgetYear)
        'm_Report.SetParameterValue("HALF_BUDGET_YEAR", Me.BudgetYear.Substring(2, 2))
        'm_Report.SetParameterValue("PROJECT_NO", Me.ProjectNo)
    End Sub

#End Region

#Region "Control Event"

    Private Sub frmBG0400_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConfigureCrystalReports()
    End Sub

#End Region

    'Private Sub CrystalReportViewer1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles CrystalReportViewer1.Load
    '    ts = CType(CrystalReportViewer1.Controls.Item(4), ToolStrip)
    '    If ts.Items.Count > 2 Then
    '        tsbtn = CType(ts.Items(1), ToolStripButton)
    '    End If
    'End Sub
    'Private Sub tsbtn_ButtonClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbtn.Click


    '    MsgBox("ปุ่มพิมพ์")


    'End Sub



    Private Sub btnPrintA4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintA4.Click

        Try

            Dim document As New Printing.PrintDocument ' This is the document to print
            Dim pageDialog2 As New PageSetupDialog ' This Dialog can set the paper size or kind
            Dim printDialog2 As New PrintDialog ' This is the dialog to setting the printer options
            Dim psize As Printing.PaperSize = Nothing

            ' The parameter of Item method is any kind of paper size avaliable on the printer

            For i = 0 To document.PrinterSettings.PaperSizes.Count - 1

                If document.PrinterSettings.PaperSizes.Item(i).Kind = PaperKind.A4 Then

                    psize = document.PrinterSettings.PaperSizes.Item(i)

                    Exit For

                End If

            Next

            If psize Is Nothing Then
                psize = document.PrinterSettings.PaperSizes.Item(0)
            End If

            ' psize = document.PrinterSettings.PaperSizes.Item(6)

            ' This line set the Page size of the document
            document.DefaultPageSettings.PaperSize = psize
            document.DefaultPageSettings.Landscape = True

            ' This is for setting the page size on the page dialog
            pageDialog2.Document = document
            pageDialog2.PageSettings.PaperSize = psize
            pageDialog2.PageSettings.Landscape = True
            ' This is for setting the page size on the printDialog
            printDialog2.Document = document
            printDialog2.AllowSomePages = True
            '     printDialog2.ShowDialog()

            If printDialog2.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor

                m_Report.PrintOptions.PrinterName = printDialog2.PrinterSettings.PrinterName

                '  m_Report.PrintOptions.PaperSize = PaperSize.PaperA4

                m_Report.PrintToPrinter(printDialog2.PrinterSettings.Copies, _
                                        printDialog2.PrinterSettings.Collate, _
                                        printDialog2.PrinterSettings.FromPage, _
                                        printDialog2.PrinterSettings.ToPage)
                'Dim pt As Printing.PrintDocument

                Me.Cursor = Cursors.Default

            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "RPT001", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.Cursor = Cursors.Default
            Return
        End Try

    End Sub
End Class