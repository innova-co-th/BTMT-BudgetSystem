Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.IO
Imports System.Data.OleDb
Imports System.Text.RegularExpressions

Public Class frmBG0670

#Region "Variable"
    Private myClsBG0670BL As New clsBG0670BL
    Private blnIsInit As Boolean = False
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
            If myClsBG0670BL.GetBudgetYear() = True And Not myClsBG0670BL.dtResult Is Nothing And myClsBG0670BL.dtResult.Rows.Count > 0 Then
                Me.cboBudgetYear.DisplayMember = "BUDGET_YEAR"
                Me.cboBudgetYear.ValueMember = "BUDGET_YEAR"
                Me.cboBudgetYear.DataSource = myClsBG0670BL.dtResult
            Else
                Me.cboBudgetYear.DataSource = Nothing
                Me.cboPeriodType.DataSource = Nothing
                Me.cboRevNo.DataSource = Nothing
            End If

            initRefBudgetYear()
            initRefBudgetYear2()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "BG0670", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Private Sub initRefBudgetYear()

        Dim dtRefBudgetYear As DataTable

        If myClsBG0670BL.GetBudgetYear() = True And Not myClsBG0670BL.dtResult Is Nothing And myClsBG0670BL.dtResult.Rows.Count > 0 Then
            dtRefBudgetYear = myClsBG0670BL.dtResult
            Me.cboRefBudgetYear.DisplayMember = "BUDGET_YEAR"
            Me.cboRefBudgetYear.ValueMember = "BUDGET_YEAR"
            Me.cboRefBudgetYear.DataSource = dtRefBudgetYear
        Else
            Me.cboRefBudgetYear.DataSource = Nothing
            Me.cboRefPeriodType.DataSource = Nothing
            Me.cboRefRevNo.DataSource = Nothing
        End If

    End Sub

    Private Sub initRefBudgetYear2()

        Dim dtRefBudgetYear As DataTable

        If myClsBG0670BL.GetBudgetYear() = True And Not myClsBG0670BL.dtResult Is Nothing And myClsBG0670BL.dtResult.Rows.Count > 0 Then
            dtRefBudgetYear = myClsBG0670BL.dtResult
            Me.cboRefBudgetYear2.DisplayMember = "BUDGET_YEAR"
            Me.cboRefBudgetYear2.ValueMember = "BUDGET_YEAR"
            Me.cboRefBudgetYear2.DataSource = dtRefBudgetYear
        Else
            Me.cboRefBudgetYear2.DataSource = Nothing
            Me.cboRefPeriodType2.DataSource = Nothing
            Me.cboRefRevNo2.DataSource = Nothing
        End If

    End Sub

    Private Function SearchBudgetAdjust() As Boolean
        Try
            myClsBG0670BL.BudgetYear = Me.cboBudgetYear.SelectedValue.ToString
            myClsBG0670BL.PeriodType = Me.cboPeriodType.SelectedValue.ToString
            myClsBG0670BL.RevNo = Me.cboRevNo.SelectedValue.ToString
            myClsBG0670BL.ProjectNo = Me.cboProjectNo.SelectedValue.ToString
            If myClsBG0670BL.GetBudgetAdjust() = True AndAlso _
                Not myClsBG0670BL.dtResult Is Nothing AndAlso _
                myClsBG0670BL.dtResult.Rows.Count > 0 Then

                Me.txtFirstHalfWBudget.Text = myClsBG0670BL.dtResult.Rows(0)("WORKING_BG1").ToString
                Me.txtSecondHalfWBudget.Text = myClsBG0670BL.dtResult.Rows(0)("WORKING_BG2").ToString

                Me.txtRRT0.Text = myClsBG0670BL.dtResult.Rows(0)("RRT0").ToString
                Me.txtRRT1.Text = myClsBG0670BL.dtResult.Rows(0)("RRT1").ToString
                Me.txtRRT2.Text = myClsBG0670BL.dtResult.Rows(0)("RRT2").ToString
                Me.txtRRT3.Text = myClsBG0670BL.dtResult.Rows(0)("RRT3").ToString
                Me.txtRRT4.Text = myClsBG0670BL.dtResult.Rows(0)("RRT4").ToString
                Me.txtRRT5.Text = myClsBG0670BL.dtResult.Rows(0)("RRT5").ToString

                Dim intRRT0 As Integer = ConvertToInt(myClsBG0670BL.dtResult.Rows(0)("RRT0").ToString)
                Dim intRRT1 As Integer = ConvertToInt(myClsBG0670BL.dtResult.Rows(0)("RRT1").ToString)
                Dim intRRT2 As Integer = ConvertToInt(myClsBG0670BL.dtResult.Rows(0)("RRT2").ToString)
                Dim intRRT3 As Integer = ConvertToInt(myClsBG0670BL.dtResult.Rows(0)("RRT3").ToString)
                Dim intRRT4 As Integer = ConvertToInt(myClsBG0670BL.dtResult.Rows(0)("RRT4").ToString)
                Dim intRRT5 As Integer = ConvertToInt(myClsBG0670BL.dtResult.Rows(0)("RRT5").ToString)

                'Me.lblRRT1p.Text = CStr(Me.CalYearlyRate(intRRT0, intRRT1))
                'Me.lblRRT2p.Text = CStr(Me.CalYearlyRate(intRRT0, intRRT2))
                'Me.lblRRT3p.Text = CStr(Me.CalYearlyRate(intRRT0, intRRT3))
                'Me.lblRRT4p.Text = CStr(Me.CalYearlyRate(intRRT0, intRRT4))
                'Me.lblRRT5p.Text = CStr(Me.CalYearlyRate(intRRT0, intRRT5))

                '-- 2018/08/23
                Me.lblRRT1p.Text = CStr(Me.CalYearlyRate(intRRT1, intRRT1))
                Me.lblRRT2p.Text = CStr(Me.CalYearlyRate(intRRT1, intRRT2))
                Me.lblRRT3p.Text = CStr(Me.CalYearlyRate(intRRT1, intRRT3))
                Me.lblRRT4p.Text = CStr(Me.CalYearlyRate(intRRT1, intRRT4))
                Me.lblRRT5p.Text = CStr(Me.CalYearlyRate(intRRT1, intRRT5))

            Else
                Me.txtRRT0.Text = ""
                Me.txtRRT1.Text = ""
                Me.txtRRT2.Text = ""
                Me.txtRRT3.Text = ""
                Me.txtRRT4.Text = ""
                Me.txtRRT5.Text = ""

                Me.lblRRT1p.Text = ""
                Me.lblRRT2p.Text = ""
                Me.lblRRT3p.Text = ""
                Me.lblRRT4p.Text = ""
                Me.lblRRT5p.Text = ""

            End If

            Return True

        Catch ex As Exception
            MessageBox.Show(ex.Message, "BG0670", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Private Function SearchBudgetReference() As Boolean
        Try
            myClsBG0670BL.BudgetYear = Me.cboBudgetYear.SelectedValue.ToString
            myClsBG0670BL.PeriodType = Me.cboPeriodType.SelectedValue.ToString
            myClsBG0670BL.RevNo = Me.cboRevNo.SelectedValue.ToString
            myClsBG0670BL.ProjectNo = Me.cboProjectNo.SelectedValue.ToString

            '// Reference1
            If Me.cboRefPeriodType.SelectedValue IsNot Nothing Then

                myClsBG0670BL.RefPeriodType = Me.cboRefPeriodType.SelectedValue.ToString
                If myClsBG0670BL.GetBudgetReference() = True AndAlso _
                    Not myClsBG0670BL.dtReference Is Nothing AndAlso _
                    myClsBG0670BL.dtReference.Rows.Count > 0 Then

                    Me.cboRefProjectNo.SelectedValue = myClsBG0670BL.dtReference.Rows(0)("REF_PROJECT_NO").ToString
                    Me.cboRefRevNo.SelectedValue = myClsBG0670BL.dtReference.Rows(0)("REF_REV_NO").ToString

                End If

            End If

            '// Reference2
            If Me.cboRefPeriodType2.SelectedValue IsNot Nothing Then

                myClsBG0670BL.RefPeriodType = Me.cboRefPeriodType2.SelectedValue.ToString
                If myClsBG0670BL.GetBudgetReference() = True AndAlso _
                    Not myClsBG0670BL.dtReference Is Nothing AndAlso _
                    myClsBG0670BL.dtReference.Rows.Count > 0 Then

                    Me.cboRefProjectNo2.SelectedValue = myClsBG0670BL.dtReference.Rows(0)("REF_PROJECT_NO").ToString
                    Me.cboRefRevNo2.SelectedValue = myClsBG0670BL.dtReference.Rows(0)("REF_REV_NO").ToString

                End If

            End If

            Return True

        Catch ex As Exception
            MessageBox.Show(ex.Message, "BG0670", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Function ConvertToInt(ByVal strValue As String) As Integer
        If strValue Is Nothing Or strValue = String.Empty Then
            Return 0
        Else
            Return Convert.ToInt32(strValue)
        End If
    End Function

    Function CalYearlyRate(ByVal intLastYear As Integer, ByVal intCurrentYear As Integer) As Double
        Dim intTmp As Integer = intCurrentYear - intLastYear

        If intTmp = 0 Then
            Return 0
        Else
            If intLastYear = 0 Then
                Return 0
            Else
                Return Math.Round(intTmp / intLastYear, 2)
            End If
        End If
    End Function

    Function SaveBudgetAdjust() As Boolean
        myClsBG0670BL.BudgetYear = Me.cboBudgetYear.SelectedValue.ToString
        myClsBG0670BL.PeriodType = Me.cboPeriodType.SelectedValue.ToString
        myClsBG0670BL.RevNo = Me.cboRevNo.SelectedValue.ToString
        myClsBG0670BL.WorkingBG1 = CheckEmptyStr(Me.txtFirstHalfWBudget.Text)
        myClsBG0670BL.WorkingBG2 = CheckEmptyStr(Me.txtSecondHalfWBudget.Text)
        myClsBG0670BL.RRT0 = CheckEmptyStr(Me.txtRRT0.Text)
        myClsBG0670BL.RRT1 = CheckEmptyStr(Me.txtRRT1.Text)
        myClsBG0670BL.RRT2 = CheckEmptyStr(Me.txtRRT2.Text)
        myClsBG0670BL.RRT3 = CheckEmptyStr(Me.txtRRT3.Text)
        myClsBG0670BL.RRT4 = CheckEmptyStr(Me.txtRRT4.Text)
        myClsBG0670BL.RRT5 = CheckEmptyStr(Me.txtRRT5.Text)
        myClsBG0670BL.UpdateUserID = p_strUserId
        myClsBG0670BL.ProjectNo = Me.cboProjectNo.SelectedValue.ToString

        If Me.cboRefBudgetYear.DataSource IsNot Nothing AndAlso Me.cboRefBudgetYear.SelectedValue IsNot Nothing AndAlso _
            Me.cboRefPeriodType.DataSource IsNot Nothing AndAlso Me.cboRefPeriodType.SelectedValue IsNot Nothing AndAlso _
            Me.cboRefProjectNo.DataSource IsNot Nothing AndAlso Me.cboRefProjectNo.SelectedValue IsNot Nothing AndAlso _
            Me.cboRefRevNo.DataSource IsNot Nothing AndAlso Me.cboRefRevNo.SelectedValue IsNot Nothing Then

            myClsBG0670BL.RefBudgetYear = Me.cboRefBudgetYear.SelectedValue.ToString
            myClsBG0670BL.RefPeriodType = Me.cboRefPeriodType.SelectedValue.ToString
            myClsBG0670BL.RefProjectNo = Me.cboRefProjectNo.SelectedValue.ToString
            myClsBG0670BL.RefRevNo = Me.cboRefRevNo.SelectedValue.ToString

        End If

        If Me.cboRefBudgetYear2.DataSource IsNot Nothing AndAlso Me.cboRefBudgetYear2.SelectedValue IsNot Nothing AndAlso _
             Me.cboRefPeriodType2.DataSource IsNot Nothing AndAlso Me.cboRefPeriodType2.SelectedValue IsNot Nothing AndAlso _
             Me.cboRefProjectNo2.DataSource IsNot Nothing AndAlso Me.cboRefProjectNo2.SelectedValue IsNot Nothing AndAlso _
             Me.cboRefRevNo2.DataSource IsNot Nothing AndAlso Me.cboRefRevNo2.SelectedValue IsNot Nothing Then

            myClsBG0670BL.RefBudgetYear2 = Me.cboRefBudgetYear2.SelectedValue.ToString
            myClsBG0670BL.RefPeriodType2 = Me.cboRefPeriodType2.SelectedValue.ToString
            myClsBG0670BL.RefProjectNo2 = Me.cboRefProjectNo2.SelectedValue.ToString
            myClsBG0670BL.RefRevNo2 = Me.cboRefRevNo2.SelectedValue.ToString

        End If
        
        If myClsBG0670BL.UpdateBudgetAdjust() = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Function CheckEmptyStr(ByVal strValue As String) As String
        If strValue Is Nothing Or strValue = String.Empty Then
            Return "0"
        Else
            Return strValue
        End If
    End Function

#End Region

#Region "Control Event"

    Private Sub frmBG0670_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        InitPage()
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        If Me.SaveBudgetAdjust() = True Then
            MessageBox.Show("Budget adjust master update successfully.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            '// Write Transaction Log
            WriteTransactionLog(CStr(enumOperationCd.EditBudgetAdjustMaster), "", "", "", "", "", "")

            SearchBudgetAdjust()
            SearchBudgetReference()
        End If
    End Sub

    Private Sub cboBudgetYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboBudgetYear.SelectedIndexChanged
        If Me.cboBudgetYear.DataSource Is Nothing Then
            Return
        End If

        Dim strYear = Me.cboBudgetYear.SelectedValue.ToString
        If Not strYear Is Nothing And strYear <> String.Empty And strYear <> "System.Data.DataRowView" Then

            myClsBG0670BL.BudgetYear = strYear
            If myClsBG0670BL.GetBudgetPeriod() = True Then

                Me.cboPeriodType.DisplayMember = "PERIOD_NAME"
                Me.cboPeriodType.ValueMember = "PERIOD_TYPE"
                Me.cboPeriodType.DataSource = myClsBG0670BL.dtResult
            Else
                Me.cboPeriodType.DataSource = Nothing
                Me.cboProjectNo.DataSource = Nothing
                Me.cboRevNo.DataSource = Nothing
            End If
        Else
            Me.cboPeriodType.DataSource = Nothing
            Me.cboProjectNo.DataSource = Nothing
            Me.cboRevNo.DataSource = Nothing
        End If

    End Sub

    Private Sub cboPeriodType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPeriodType.SelectedIndexChanged
        If cboPeriodType.DataSource Is Nothing Then
            Return
        Else
            Dim strPeriod = Me.cboPeriodType.SelectedValue.ToString
            If Not strPeriod Is Nothing And strPeriod <> String.Empty And strPeriod <> "System.Data.DataRowView" Then

                myClsBG0670BL.BudgetYear = Me.cboBudgetYear.SelectedValue.ToString
                myClsBG0670BL.PeriodType = strPeriod

                If myClsBG0670BL.GetProjectNo() = True Then
                    Me.cboProjectNo.DisplayMember = "PROJECT_NO"
                    Me.cboProjectNo.ValueMember = "PROJECT_NO"
                    Me.cboProjectNo.DataSource = myClsBG0670BL.dtResult
                Else
                    Me.cboProjectNo.DataSource = Nothing
                    Me.cboRevNo.DataSource = Nothing
                End If
            Else
                Me.cboProjectNo.DataSource = Nothing
                Me.cboRevNo.DataSource = Nothing
            End If
        End If

        initReferenceBudget()
        initReferenceBudget2()
        SearchBudgetReference()
    End Sub

    Private Sub initReferenceBudget()

        'If blnIsInit = True Then
        '    Return
        'End If

        'blnIsInit = True
        If cboPeriodType.DataSource IsNot Nothing Then

            Dim strPeriod = Me.cboPeriodType.SelectedValue.ToString
            If Not strPeriod Is Nothing And strPeriod <> String.Empty And strPeriod <> "System.Data.DataRowView" Then

                If CInt(strPeriod) = BGConstant.enumPeriodType.OriginalBudget Then
                    grbReference.Enabled = True
                    Me.cboRefBudgetYear.SelectedValue = CInt(Me.cboBudgetYear.SelectedValue) - 1
                    Me.ComboBox1_SelectedIndexChanged("", Nothing)

                    'ElseIf CInt(strPeriod) = BGConstant.enumPeriodType.ForecastBudget Then
                    '    grbReference.Enabled = True
                    '    Me.cboRefBudgetYear.SelectedValue = CInt(Me.cboBudgetYear.SelectedValue)
                    '    Me.ComboBox1_SelectedIndexChanged("", Nothing)
                ElseIf CInt(strPeriod) = BGConstant.enumPeriodType.MBPBudget Then
                    grbReference.Enabled = True
                    Me.cboRefBudgetYear.SelectedValue = CInt(Me.cboBudgetYear.SelectedValue) + 1
                    Me.ComboBox1_SelectedIndexChanged("", Nothing)
                Else
                    grbReference.Enabled = False

                End If

            End If

        End If
        'blnIsInit = False
    End Sub

    Private Sub initReferenceBudget2()

        If cboPeriodType.DataSource IsNot Nothing Then

            Dim strPeriod = Me.cboPeriodType.SelectedValue.ToString
            If Not strPeriod Is Nothing And strPeriod <> String.Empty And strPeriod <> "System.Data.DataRowView" Then

                If CInt(strPeriod) = BGConstant.enumPeriodType.OriginalBudget Then
                    grbReference2.Enabled = True
                    Me.cboRefBudgetYear2.SelectedValue = CInt(Me.cboBudgetYear.SelectedValue) - 2
                    Me.cboRefBudgetYear2_SelectedIndexChanged("", Nothing)

                ElseIf CInt(strPeriod) = BGConstant.enumPeriodType.MBPBudget Then
                    grbReference2.Enabled = True
                    Me.cboRefBudgetYear2.SelectedValue = CInt(Me.cboBudgetYear.SelectedValue) - 1
                    Me.cboRefBudgetYear2_SelectedIndexChanged("", Nothing)
                Else
                    grbReference2.Enabled = False

                End If

            End If

        End If

    End Sub

    Private Sub cboProjectNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboProjectNo.SelectedIndexChanged
        If cboProjectNo.DataSource Is Nothing Then
            Return
        Else
            Dim strProjectNo = Me.cboProjectNo.SelectedValue.ToString
            If Not strProjectNo Is Nothing And strProjectNo <> String.Empty And strProjectNo <> "System.Data.DataRowView" Then

                myClsBG0670BL.BudgetYear = Me.cboBudgetYear.SelectedValue.ToString
                myClsBG0670BL.PeriodType = Me.cboPeriodType.SelectedValue.ToString
                myClsBG0670BL.ProjectNo = strProjectNo

                If myClsBG0670BL.GetRevNo() = True Then
                    Me.cboRevNo.DisplayMember = "REV_NO"
                    Me.cboRevNo.ValueMember = "REV_NO"
                    Me.cboRevNo.DataSource = myClsBG0670BL.dtResult
                Else
                    Me.cboRevNo.DataSource = Nothing
                End If
            Else
                Me.cboRevNo.DataSource = Nothing
            End If
        End If
    End Sub

    Private Sub cboRevNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboRevNo.SelectedIndexChanged
        If Me.cboRevNo.DataSource Is Nothing Then
            Return
        End If

        Dim strYear As String = Me.cboBudgetYear.SelectedValue.ToString
        Dim strPeriod As String = Me.cboPeriodType.SelectedValue.ToString
        Dim strProjectNo = Me.cboProjectNo.SelectedValue.ToString
        Dim strRevNo As String = Me.cboRevNo.SelectedValue.ToString

        If strYear <> "" AndAlso Not strYear Is Nothing AndAlso _
            strPeriod <> "" AndAlso Not strPeriod Is Nothing AndAlso _
            strProjectNo <> "" AndAlso Not strProjectNo Is Nothing AndAlso _
            strRevNo <> "" AndAlso Not strRevNo Is Nothing AndAlso strRevNo <> "System.Data.DataRowView" Then

            SearchBudgetAdjust()
            SearchBudgetReference()
        End If
    End Sub

    Private Sub txtFirstHalfWBudget_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFirstHalfWBudget.KeyPress
        If Char.IsControl(e.KeyChar) Then
            e.Handled = False
            Return
        End If

        Dim regexText As String = "^\d*\.?\d{0,2}$"
        Dim regex As Regex = New Regex(regexText)

        If Not regex.IsMatch(e.KeyChar) Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub txtSecondHalfWBudget_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSecondHalfWBudget.KeyPress
        If Char.IsControl(e.KeyChar) Then
            e.Handled = False
            Return
        End If

        Dim regexText As String = "^\d*\.?\d{0,2}$"
        Dim regex As Regex = New Regex(regexText)

        If Not regex.IsMatch(e.KeyChar) Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub txtRRT0_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRRT0.KeyPress
        If Char.IsDigit(e.KeyChar) Or Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub txtRRT1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRRT1.KeyPress
        If Char.IsDigit(e.KeyChar) Or Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub txtRRT2_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRRT2.KeyPress
        If Char.IsDigit(e.KeyChar) Or Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub txtRRT3_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRRT3.KeyPress
        If Char.IsDigit(e.KeyChar) Or Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub txtRRT4_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRRT4.KeyPress
        If Char.IsDigit(e.KeyChar) Or Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub txtRRT5_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRRT5.KeyPress
        If Char.IsDigit(e.KeyChar) Or Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub txtRRT0_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRRT0.Leave
        'Dim intLastYear As Integer = CInt(CheckEmptyStr(Me.txtRRT0.Text))
        Dim intLastYear As Integer = CInt(CheckEmptyStr(Me.txtRRT1.Text))
        Dim intCurrentYear As Integer = CInt(CheckEmptyStr(Me.txtRRT1.Text))

        Me.lblRRT1p.Text = CStr(CalYearlyRate(intLastYear, intCurrentYear))

        intCurrentYear = CInt(CheckEmptyStr(Me.txtRRT2.Text))
        Me.lblRRT2p.Text = CStr(CalYearlyRate(intLastYear, intCurrentYear))

        intCurrentYear = CInt(CheckEmptyStr(Me.txtRRT3.Text))
        Me.lblRRT3p.Text = CStr(CalYearlyRate(intLastYear, intCurrentYear))

        intCurrentYear = CInt(CheckEmptyStr(Me.txtRRT4.Text))
        Me.lblRRT4p.Text = CStr(CalYearlyRate(intLastYear, intCurrentYear))

        intCurrentYear = CInt(CheckEmptyStr(Me.txtRRT5.Text))
        Me.lblRRT5p.Text = CStr(CalYearlyRate(intLastYear, intCurrentYear))
    End Sub

    Private Sub txtRRT1_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRRT1.Leave
        Dim intLastYear As Integer = CInt(CheckEmptyStr(Me.txtRRT0.Text))
        Dim intCurrentYear As Integer = CInt(CheckEmptyStr(Me.txtRRT1.Text))

        Me.lblRRT1p.Text = CStr(CalYearlyRate(intLastYear, intCurrentYear))
    End Sub

    Private Sub txtRRT2_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRRT2.Leave
        'Dim intLastYear As Integer = CInt(CheckEmptyStr(Me.txtRRT0.Text))
        Dim intLastYear As Integer = CInt(CheckEmptyStr(Me.txtRRT1.Text))
        Dim intCurrentYear As Integer = CInt(CheckEmptyStr(Me.txtRRT2.Text))

        Me.lblRRT2p.Text = CStr(CalYearlyRate(intLastYear, intCurrentYear))
    End Sub

    Private Sub txtRRT3_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRRT3.Leave
        'Dim intLastYear As Integer = CInt(CheckEmptyStr(Me.txtRRT0.Text))
        Dim intLastYear As Integer = CInt(CheckEmptyStr(Me.txtRRT1.Text))
        Dim intCurrentYear As Integer = CInt(CheckEmptyStr(Me.txtRRT3.Text))

        Me.lblRRT3p.Text = CStr(CalYearlyRate(intLastYear, intCurrentYear))
    End Sub

    Private Sub txtRRT4_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRRT4.Leave
        'Dim intLastYear As Integer = CInt(CheckEmptyStr(Me.txtRRT0.Text))
        Dim intLastYear As Integer = CInt(CheckEmptyStr(Me.txtRRT1.Text))
        Dim intCurrentYear As Integer = CInt(CheckEmptyStr(Me.txtRRT4.Text))

        Me.lblRRT4p.Text = CStr(CalYearlyRate(intLastYear, intCurrentYear))
    End Sub

    Private Sub txtRRT5_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRRT5.Leave
        'Dim intLastYear As Integer = CInt(CheckEmptyStr(Me.txtRRT0.Text))
        Dim intLastYear As Integer = CInt(CheckEmptyStr(Me.txtRRT1.Text))
        Dim intCurrentYear As Integer = CInt(CheckEmptyStr(Me.txtRRT5.Text))

        Me.lblRRT5p.Text = CStr(CalYearlyRate(intLastYear, intCurrentYear))
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboRefBudgetYear.SelectedIndexChanged
        If Me.cboRefBudgetYear.DataSource Is Nothing Then
            Me.cboRefPeriodType.DataSource = Nothing
            Me.cboRefProjectNo.DataSource = Nothing
            Me.cboRefRevNo.DataSource = Nothing
            Return
        End If
        If Me.cboRefBudgetYear.SelectedValue Is Nothing Then
            Me.cboRefPeriodType.DataSource = Nothing
            Me.cboRefProjectNo.DataSource = Nothing
            Me.cboRefRevNo.DataSource = Nothing
            Return
        End If

        Dim strYear = Me.cboRefBudgetYear.SelectedValue.ToString
        If Not strYear Is Nothing And strYear <> String.Empty And strYear <> "System.Data.DataRowView" Then

            myClsBG0670BL.BudgetYear = strYear
            If myClsBG0670BL.GetBudgetPeriod() = True Then

                Me.cboRefPeriodType.DisplayMember = "PERIOD_NAME"
                Me.cboRefPeriodType.ValueMember = "PERIOD_TYPE"
                Me.cboRefPeriodType.DataSource = myClsBG0670BL.dtResult

                Dim strPeriod = Me.cboPeriodType.SelectedValue.ToString

                If CInt(strPeriod) = BGConstant.enumPeriodType.OriginalBudget Then

                    Me.cboRefPeriodType.SelectedValue = BGConstant.enumPeriodType.EstimateBudget
                    cboRefPeriodType_SelectedIndexChanged("", Nothing)

                    'ElseIf CInt(strPeriod) = BGConstant.enumPeriodType.ForecastBudget Then
                    '    Me.cboRefPeriodType.SelectedValue = BGConstant.enumPeriodType.OriginalBudget
                    '    cboRefPeriodType_SelectedIndexChanged("", Nothing)
                ElseIf CInt(strPeriod) = BGConstant.enumPeriodType.MBPBudget Then
                    Me.cboRefPeriodType.SelectedValue = BGConstant.enumPeriodType.OriginalBudget ' Edited by Kwang for Prototype No.4. Menu : Budget Adjust Master 'BGConstant.enumPeriodType.ForecastBudget
                    cboRefPeriodType_SelectedIndexChanged("", Nothing)
                Else
                    grbReference.Enabled = False

                End If
            Else
                Me.cboRefPeriodType.DataSource = Nothing
                Me.cboRefProjectNo.DataSource = Nothing
                Me.cboRefRevNo.DataSource = Nothing
            End If
        Else
            Me.cboRefPeriodType.DataSource = Nothing
            Me.cboRefProjectNo.DataSource = Nothing
            Me.cboRefRevNo.DataSource = Nothing
        End If
    End Sub

    Private Sub cboRefPeriodType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRefPeriodType.SelectedIndexChanged
        If cboRefPeriodType.DataSource Is Nothing Then
            Me.cboRefProjectNo.DataSource = Nothing
            Me.cboRefRevNo.DataSource = Nothing
            Return
        ElseIf cboRefPeriodType.SelectedValue Is Nothing Then
            Me.cboRefProjectNo.DataSource = Nothing
            Me.cboRefRevNo.DataSource = Nothing
            Return
        Else
            Dim strPeriod = Me.cboRefPeriodType.SelectedValue.ToString
            If Not strPeriod Is Nothing And strPeriod <> String.Empty And strPeriod <> "System.Data.DataRowView" Then

                myClsBG0670BL.BudgetYear = Me.cboRefBudgetYear.SelectedValue.ToString
                myClsBG0670BL.PeriodType = strPeriod

                If myClsBG0670BL.GetProjectNo() = True Then
                    Me.cboRefProjectNo.DisplayMember = "PROJECT_NO"
                    Me.cboRefProjectNo.ValueMember = "PROJECT_NO"
                    Me.cboRefProjectNo.DataSource = myClsBG0670BL.dtResult
                Else
                    Me.cboRefProjectNo.DataSource = Nothing
                    Me.cboRefRevNo.DataSource = Nothing
                End If
            Else
                Me.cboRefProjectNo.DataSource = Nothing
                Me.cboRefRevNo.DataSource = Nothing
            End If
        End If
    End Sub

    Private Sub cboRefProjectNo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRefProjectNo.SelectedIndexChanged
        If cboRefProjectNo.DataSource Is Nothing Then
            Return
        Else
            Dim strProjectNo = Me.cboRefProjectNo.SelectedValue.ToString
            If Not strProjectNo Is Nothing And strProjectNo <> String.Empty And strProjectNo <> "System.Data.DataRowView" Then

                myClsBG0670BL.BudgetYear = Me.cboRefBudgetYear.SelectedValue.ToString
                myClsBG0670BL.PeriodType = Me.cboRefPeriodType.SelectedValue.ToString
                myClsBG0670BL.ProjectNo = strProjectNo

                If myClsBG0670BL.GetRevNo() = True Then
                    Me.cboRefRevNo.DisplayMember = "REV_NO"
                    Me.cboRefRevNo.ValueMember = "REV_NO"
                    Me.cboRefRevNo.DataSource = myClsBG0670BL.dtResult
                Else
                    Me.cboRefRevNo.DataSource = Nothing
                End If
            Else
                Me.cboRefRevNo.DataSource = Nothing
            End If
        End If
    End Sub

#End Region

    Private Sub cboRefBudgetYear2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRefBudgetYear2.SelectedIndexChanged
        If Me.cboRefBudgetYear2.DataSource Is Nothing Then
            Me.cboRefPeriodType2.DataSource = Nothing
            Me.cboRefProjectNo2.DataSource = Nothing
            Me.cboRefRevNo2.DataSource = Nothing
            Return
        End If
        If Me.cboRefBudgetYear2.SelectedValue Is Nothing Then
            Me.cboRefPeriodType2.DataSource = Nothing
            Me.cboRefProjectNo2.DataSource = Nothing
            Me.cboRefRevNo2.DataSource = Nothing
            Return
        End If

        Dim strYear = Me.cboRefBudgetYear2.SelectedValue.ToString
        If Not strYear Is Nothing And strYear <> String.Empty And strYear <> "System.Data.DataRowView" Then

            myClsBG0670BL.BudgetYear = strYear
            If myClsBG0670BL.GetBudgetPeriod() = True Then

                Me.cboRefPeriodType2.DisplayMember = "PERIOD_NAME"
                Me.cboRefPeriodType2.ValueMember = "PERIOD_TYPE"
                Me.cboRefPeriodType2.DataSource = myClsBG0670BL.dtResult

                Dim strPeriod = Me.cboPeriodType.SelectedValue.ToString

                If CInt(strPeriod) = BGConstant.enumPeriodType.OriginalBudget Then
                    Me.cboRefPeriodType2.SelectedValue = BGConstant.enumPeriodType.MBPBudget
                    cboRefPeriodType2_SelectedIndexChanged("", Nothing)

                ElseIf CInt(strPeriod) = BGConstant.enumPeriodType.MBPBudget Then
                    Me.cboRefPeriodType2.SelectedValue = BGConstant.enumPeriodType.MBPBudget
                    cboRefPeriodType2_SelectedIndexChanged("", Nothing)
                Else

                End If
            Else
                Me.cboRefPeriodType2.DataSource = Nothing
                Me.cboRefProjectNo2.DataSource = Nothing
                Me.cboRefRevNo2.DataSource = Nothing
            End If
        Else
            Me.cboRefPeriodType2.DataSource = Nothing
            Me.cboRefProjectNo2.DataSource = Nothing
            Me.cboRefRevNo2.DataSource = Nothing
        End If
    End Sub

    Private Sub cboRefPeriodType2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRefPeriodType2.SelectedIndexChanged
        If cboRefPeriodType2.DataSource Is Nothing Then
            Me.cboRefProjectNo2.DataSource = Nothing
            Me.cboRefRevNo2.DataSource = Nothing
            Return
        ElseIf cboRefPeriodType2.SelectedValue Is Nothing Then
            Me.cboRefProjectNo2.DataSource = Nothing
            Me.cboRefRevNo2.DataSource = Nothing
            Return
        Else
            Dim strPeriod = Me.cboRefPeriodType2.SelectedValue.ToString
            If Not strPeriod Is Nothing And strPeriod <> String.Empty And strPeriod <> "System.Data.DataRowView" Then

                myClsBG0670BL.BudgetYear = Me.cboRefBudgetYear2.SelectedValue.ToString
                myClsBG0670BL.PeriodType = strPeriod

                If myClsBG0670BL.GetProjectNo() = True Then
                    Me.cboRefProjectNo2.DisplayMember = "PROJECT_NO"
                    Me.cboRefProjectNo2.ValueMember = "PROJECT_NO"
                    Me.cboRefProjectNo2.DataSource = myClsBG0670BL.dtResult
                Else
                    Me.cboRefProjectNo2.DataSource = Nothing
                    Me.cboRefRevNo2.DataSource = Nothing
                End If
            Else
                Me.cboRefProjectNo2.DataSource = Nothing
                Me.cboRefRevNo2.DataSource = Nothing
            End If
        End If
    End Sub

    Private Sub cboRefProjectNo2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRefProjectNo2.SelectedIndexChanged
        If cboRefProjectNo2.DataSource Is Nothing Then
            Return
        Else
            Dim strProjectNo = Me.cboRefProjectNo2.SelectedValue.ToString
            If Not strProjectNo Is Nothing And strProjectNo <> String.Empty And strProjectNo <> "System.Data.DataRowView" Then

                myClsBG0670BL.BudgetYear = Me.cboRefBudgetYear2.SelectedValue.ToString
                myClsBG0670BL.PeriodType = Me.cboRefPeriodType2.SelectedValue.ToString
                myClsBG0670BL.ProjectNo = strProjectNo

                If myClsBG0670BL.GetRevNo() = True Then
                    Me.cboRefRevNo2.DisplayMember = "REV_NO"
                    Me.cboRefRevNo2.ValueMember = "REV_NO"
                    Me.cboRefRevNo2.DataSource = myClsBG0670BL.dtResult
                Else
                    Me.cboRefRevNo2.DataSource = Nothing
                End If
            Else
                Me.cboRefRevNo2.DataSource = Nothing
            End If
        End If
    End Sub
End Class