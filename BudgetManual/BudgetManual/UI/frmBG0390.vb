Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class frmBG0390

#Region "Variable"
    Private myClsBG0390BL As New clsBG0390BL
    Private myClsBG0310BL As New clsBG0310BL
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

    Private Function InitForm() As Boolean
        Me.numYear.Value = Now.Year

        Me.cboPeriodType.Items.Clear()

        myClsBG0310BL.OpenPeriodFlg = "1"
        myClsBG0310BL.GetOpenPeriodList()

        If myClsBG0310BL.PeriodList IsNot Nothing AndAlso myClsBG0310BL.PeriodList.Rows.Count > 0 Then
            cboPeriodType.DisplayMember = "PERIOD_TYPE_NAME"
            cboPeriodType.ValueMember = "PERIOD_TYPE_ID"
            cboPeriodType.DataSource = myClsBG0310BL.PeriodList

            cboPeriodType.SelectedIndex = 0
        End If
        'cboPeriodType.Items.Clear()
        'cboPeriodType.Items.Add("Original Budget")
        'cboPeriodType.Items.Add("Estimate Budget")
        'cboPeriodType.Items.Add("Forecast Budget")
        'cboPeriodType.SelectedIndex = 0
    End Function

#End Region

#Region "Control Event"
    Private Sub frmBG0390_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        InitForm()
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub cmdSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSearch.Click

        If optBudget.Checked Then

            myClsBG0390BL.BudgetYear = CStr(numYear.Value)
            Select Case CInt(cboPeriodType.SelectedValue)
                Case CType(enumPeriodType.OriginalBudget, Integer)
                    myClsBG0390BL.PeriodType = CStr(enumPeriodType.OriginalBudget)
                Case CType(enumPeriodType.EstimateBudget, Integer)
                    myClsBG0390BL.PeriodType = CStr(enumPeriodType.EstimateBudget)
                Case CType(enumPeriodType.ForecastBudget, Integer)
                    myClsBG0390BL.PeriodType = CStr(enumPeriodType.ForecastBudget)
                Case CType(enumPeriodType.MTPBudget, Integer)
                    myClsBG0390BL.PeriodType = CStr(enumPeriodType.MTPBudget)
            End Select
            myClsBG0390BL.ProjectNo = numProjectNo.Value.ToString


            If myClsBG0390BL.SearchTransLog() Then
                grvLog.DataSource = myClsBG0390BL.DTResult
            Else
                grvLog.DataSource = Nothing
            End If

        ElseIf optAccount.Checked Then

            myClsBG0390BL.FromDate = dtpFrom.Value.ToString("yyyy-MM-dd") & " 00:00:00"
            myClsBG0390BL.ToDate = dtpTo.Value.ToString("yyyy-MM-dd") & " 23:59:59"

            If myClsBG0390BL.SearchAdminLog() Then
                grvLog.DataSource = myClsBG0390BL.DTResult
            Else
                grvLog.DataSource = Nothing
            End If

        End If
    End Sub

    Private Sub optAccount_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles optAccount.Click
        If optAccount.Checked Then
            numYear.Enabled = False
            cboPeriodType.Enabled = False
            numProjectNo.Enabled = False
            dtpFrom.Enabled = True
            dtpTo.Enabled = True
        End If
    End Sub

    Private Sub optBudget_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles optBudget.Click
        If optBudget.Checked Then
            numYear.Enabled = True
            cboPeriodType.Enabled = True
            numProjectNo.Enabled = True
            dtpFrom.Enabled = False
            dtpTo.Enabled = False
        End If
    End Sub

#End Region

    Private Sub cboPeriodType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPeriodType.SelectedIndexChanged
        If cboPeriodType.SelectedIndex >= 0 Then

            If CStr(cboPeriodType.SelectedValue) = CStr(enumPeriodType.MTPBudget) Then
                Me.numProjectNo.Enabled = True
            Else
                Me.numProjectNo.Value = 1
                Me.numProjectNo.Enabled = False
            End If

        End If
    End Sub
End Class