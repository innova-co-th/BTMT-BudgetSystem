Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.SqlClient

Public Class clsBG0330BL

#Region "Variable"
    Private myDtResult As DataTable
    Private myPeriodList As DataTable
    Private myBudgetYear As String = String.Empty
    Private myProjectNo As String = String.Empty
    Private myPeriodType As String = String.Empty
    Private myUserId As String = String.Empty
    Private myAccountNo As String = String.Empty
    Private myPicNo As String = String.Empty
    Private myAddPicList As ArrayList = Nothing
    Private myAddAccountList As ArrayList = Nothing
#End Region

#Region "Property"

#Region "DtResult"
    Property DtResult() As DataTable
        Get
            Return myDtResult
        End Get
        Set(ByVal value As DataTable)
            myDtResult = value
        End Set
    End Property
#End Region

#Region "PeriodList"
    Property PeriodList() As DataTable
        Get
            Return myPeriodList
        End Get
        Set(ByVal value As DataTable)
            myPeriodList = value
        End Set
    End Property
#End Region

#Region "BudgetYear"
    Property BudgetYear() As String
        Get
            Return myBudgetYear
        End Get
        Set(ByVal value As String)
            myBudgetYear = value
        End Set
    End Property
#End Region

#Region "ProjectNo"
    Property ProjectNo() As String
        Get
            Return myProjectNo
        End Get
        Set(ByVal value As String)
            myProjectNo = value
        End Set
    End Property
#End Region

#Region "PeriodType"
    Property PeriodType() As String
        Get
            Return myPeriodType
        End Get
        Set(ByVal value As String)
            myPeriodType = value
        End Set
    End Property
#End Region

#Region "UserId"
    Property UserId() As String
        Get
            Return myUserId
        End Get
        Set(ByVal value As String)
            myUserId = value
        End Set
    End Property
#End Region

#Region "AccountNo"
    Property AccountNo() As String
        Get
            Return myAccountNo
        End Get
        Set(ByVal value As String)
            myAccountNo = value
        End Set
    End Property
#End Region

#Region "PicNo"
    Property PicNo() As String
        Get
            Return myPicNo
        End Get
        Set(ByVal value As String)
            myPicNo = value
        End Set
    End Property
#End Region

#Region "AddPicList"
    Property AddPicList() As ArrayList
        Get
            Return myAddPicList
        End Get
        Set(ByVal value As ArrayList)
            myAddPicList = value
        End Set
    End Property
#End Region

#Region "AddAccountList"
    Property AddAccountList() As ArrayList
        Get
            Return myAddAccountList
        End Get
        Set(ByVal value As ArrayList)
            myAddAccountList = value
        End Set
    End Property
#End Region

#End Region

#Region "Function"

    Public Function SearchBudgetPeriod() As Boolean
        Dim clsBG_T_BUDGET_PERIOD As New BG_T_BUDGET_PERIOD

        '// Call Function
        If clsBG_T_BUDGET_PERIOD.Select002() = True AndAlso _
        clsBG_T_BUDGET_PERIOD.dtResult.Rows.Count > 0 Then
            Me.PeriodList = clsBG_T_BUDGET_PERIOD.dtResult

            Return True
        Else
            Me.PeriodList = Nothing

            Return False
        End If
    End Function

    Public Function ReopenPeriod() As Boolean
        Dim clsBG_T_BUDGET_PERIOD As New BG_T_BUDGET_PERIOD

        '// Set Parameters
        clsBG_T_BUDGET_PERIOD.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_PERIOD.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_PERIOD.UserId = Me.UserId
        clsBG_T_BUDGET_PERIOD.CloseFlg = "0"
        clsBG_T_BUDGET_PERIOD.ProjectNo = Me.ProjectNo

        '// Call Function
        If clsBG_T_BUDGET_PERIOD.Update001() = True Then

            Return True
        Else
            Return False

        End If
    End Function

    Public Function DeleteReopenAccount() As Boolean
        Dim clsBG_T_ACCOUNT_REOPEN As New BG_T_ACCOUNT_REOPEN

        '// Set Parameters
        clsBG_T_ACCOUNT_REOPEN.BudgetYear = Me.BudgetYear
        clsBG_T_ACCOUNT_REOPEN.PeriodType = Me.PeriodType
        clsBG_T_ACCOUNT_REOPEN.AccountNo = Me.AccountNo
        clsBG_T_ACCOUNT_REOPEN.PicNo = Me.PicNo
        clsBG_T_ACCOUNT_REOPEN.ProjectNo = Me.ProjectNo

        '// Call Function
        If clsBG_T_ACCOUNT_REOPEN.Delete001() = True Then

            Return True
        Else
            Return False

        End If
    End Function

    Public Function DeleteAllReopenAccount() As Boolean
        Dim clsBG_T_ACCOUNT_REOPEN As New BG_T_ACCOUNT_REOPEN

        '// Set Parameters
        clsBG_T_ACCOUNT_REOPEN.BudgetYear = Me.BudgetYear
        clsBG_T_ACCOUNT_REOPEN.PeriodType = Me.PeriodType
        clsBG_T_ACCOUNT_REOPEN.ProjectNo = Me.ProjectNo

        '// Call Function
        If clsBG_T_ACCOUNT_REOPEN.Delete002() = True Then

            Return True
        Else
            Return False

        End If
    End Function

    Public Function AddReopenAccount() As Boolean
        If Me.AccountNo = "All" Then
            Return AddAllAccount()

        ElseIf Me.PicNo = "All" Then
            Return AddAllPic()

        Else
            Return AddAccount()
        End If
    End Function

    Private Function AddAccount() As Boolean
        Dim clsBG_T_ACCOUNT_REOPEN As New BG_T_ACCOUNT_REOPEN

        '// Set Parameters
        clsBG_T_ACCOUNT_REOPEN.BudgetYear = Me.BudgetYear
        clsBG_T_ACCOUNT_REOPEN.PeriodType = Me.PeriodType
        clsBG_T_ACCOUNT_REOPEN.AccountNo = Me.AccountNo
        clsBG_T_ACCOUNT_REOPEN.PicNo = Me.PicNo
        clsBG_T_ACCOUNT_REOPEN.UserId = Me.UserId
        clsBG_T_ACCOUNT_REOPEN.ProjectNo = Me.ProjectNo

        '// Call Function
        If clsBG_T_ACCOUNT_REOPEN.Insert001() = True Then

            Return True
        Else
            Return False

        End If
    End Function

    Private Function AddAllAccount() As Boolean
        If Me.AddAccountList Is Nothing OrElse Me.AddAccountList.Count = 0 Then

            Return False
        End If

        Dim clsBG_T_ACCOUNT_REOPEN As New BG_T_ACCOUNT_REOPEN
        Dim conn As SqlConnection
        Dim trans As SqlTransaction

        '// Create Connection
        conn = New SqlConnection(My.Settings.ConnStr)
        conn.Open()

        '// Begin Transaction
        trans = conn.BeginTransaction()

        Try
            '// Delete all account of selected PIC
            clsBG_T_ACCOUNT_REOPEN.BudgetYear = Me.BudgetYear
            clsBG_T_ACCOUNT_REOPEN.PeriodType = Me.PeriodType
            clsBG_T_ACCOUNT_REOPEN.PicNo = Me.PicNo
            clsBG_T_ACCOUNT_REOPEN.ProjectNo = Me.ProjectNo

            clsBG_T_ACCOUNT_REOPEN.Delete004(conn, trans)

            '// Insert all account of selected PIC
            For i = 0 To AddAccountList.Count - 1
                clsBG_T_ACCOUNT_REOPEN.AccountNo = Me.AddAccountList(i).ToString
                clsBG_T_ACCOUNT_REOPEN.UserId = Me.UserId

                If clsBG_T_ACCOUNT_REOPEN.Insert001(conn, trans) = False Then
                    Throw New Exception("Can not insert account!")
                End If
            Next

            '// Commit Transaction
            trans.Commit()

            '// Close Connection
            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception

            '// Rollback Transaction
            trans.Rollback()

            '// Close Connection
            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            MessageBox.Show("Error: " & ex.Message, My.Settings.ProgramTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)

            Return False
        End Try
    End Function

    Private Function AddAllPic() As Boolean
        If Me.AddPicList Is Nothing OrElse Me.AddPicList.Count = 0 Then

            Return False
        End If

        Dim clsBG_T_ACCOUNT_REOPEN As New BG_T_ACCOUNT_REOPEN
        Dim conn As SqlConnection
        Dim trans As SqlTransaction

        '// Create Connection
        conn = New SqlConnection(My.Settings.ConnStr)
        conn.Open()

        '// Begin Transaction
        trans = conn.BeginTransaction()

        Try
            '// Delete all PIC of selected account
            clsBG_T_ACCOUNT_REOPEN.BudgetYear = Me.BudgetYear
            clsBG_T_ACCOUNT_REOPEN.PeriodType = Me.PeriodType
            clsBG_T_ACCOUNT_REOPEN.AccountNo = Me.AccountNo
            clsBG_T_ACCOUNT_REOPEN.ProjectNo = Me.ProjectNo

            clsBG_T_ACCOUNT_REOPEN.Delete003(conn, trans)

            '// Insert all PIC of selected account
            For i = 0 To AddPicList.Count - 1
                clsBG_T_ACCOUNT_REOPEN.PicNo = Me.AddPicList(i).ToString
                clsBG_T_ACCOUNT_REOPEN.UserId = Me.UserId

                If clsBG_T_ACCOUNT_REOPEN.Insert001(conn, trans) = False Then
                    Throw New Exception("Can not insert account!")
                End If
            Next

            '// Commit Transaction
            trans.Commit()

            '// Close Connection
            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception

            '// Rollback Transaction
            trans.Rollback()

            '// Close Connection
            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            MessageBox.Show("Error: " & ex.Message, My.Settings.ProgramTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)

            Return False
        End Try
    End Function

    Public Function SearchReopenAccount() As Boolean
        Dim clsBG_T_ACCOUNT_REOPEN As New BG_T_ACCOUNT_REOPEN

        '// Set Parameters
        clsBG_T_ACCOUNT_REOPEN.BudgetYear = Me.BudgetYear
        clsBG_T_ACCOUNT_REOPEN.PeriodType = Me.PeriodType
        clsBG_T_ACCOUNT_REOPEN.ProjectNo = Me.ProjectNo

        '// Call Function
        If clsBG_T_ACCOUNT_REOPEN.Select001() = True AndAlso _
        clsBG_T_ACCOUNT_REOPEN.dtResult.Rows.Count > 0 Then
            Me.DtResult = clsBG_T_ACCOUNT_REOPEN.dtResult

            Return True
        Else
            Me.DtResult = Nothing

            Return False
        End If
    End Function

    Public Function GetAccountList() As Boolean
        Dim clsBG_M_ACCOUNT As New BG_M_ACCOUNT

        If clsBG_M_ACCOUNT.Select001 Then
            Me.DtResult = clsBG_M_ACCOUNT.DtResult

            Return True
        Else
            Me.DtResult = New DataTable

            Return False
        End If
    End Function

    Public Function GetAccountList2() As Boolean
        Dim clsBG_M_ACCOUNT As New BG_M_ACCOUNT

        clsBG_M_ACCOUNT.PicNo = Me.PicNo

        If clsBG_M_ACCOUNT.Select004 Then
            Me.DtResult = clsBG_M_ACCOUNT.DtResult

            Return True
        Else
            Me.DtResult = New DataTable

            Return False
        End If
    End Function

    Public Function GetPicList() As Boolean
        Dim clsBG_M_PERSON_IN_CHARGE As New BG_M_PERSON_IN_CHARGE

        clsBG_M_PERSON_IN_CHARGE.AccountNo = Me.AccountNo

        If clsBG_M_PERSON_IN_CHARGE.Select010 Then
            Me.DtResult = clsBG_M_PERSON_IN_CHARGE.DtResult

            Return True
        Else
            Me.DtResult = New DataTable

            Return False
        End If
    End Function

    Public Function GetAllPicList() As Boolean
        Dim clsBG_M_PERSON_IN_CHARGE As New BG_M_PERSON_IN_CHARGE

        If clsBG_M_PERSON_IN_CHARGE.Select002 Then
            Me.DtResult = clsBG_M_PERSON_IN_CHARGE.DtResult

            Return True
        Else
            Me.DtResult = New DataTable

            Return False
        End If
    End Function

#End Region

End Class
