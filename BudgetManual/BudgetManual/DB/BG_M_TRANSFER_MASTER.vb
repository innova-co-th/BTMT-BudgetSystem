Imports System.Data.SqlClient
Imports BudgetManual.BGCommon
Imports BudgetManual.BGConstant

Public Class BG_M_TRANSFER_MASTER

#Region "Variable"
    Private myBudgetYear As String = String.Empty
    Private myProjectNo As String = String.Empty
    Private myPeriodType As String = String.Empty
    Private myBudgetOrderNo As String = String.Empty
    Private myTransferType As String = String.Empty
    Private myTransferRate As String = String.Empty
    Private myFromOrderNo As String = String.Empty
    Private myToOrderNo As String = String.Empty
    Private myCreateUserID As String = String.Empty
    Private myCreateDate As String = String.Empty
    Private myUpdateUserID As String = String.Empty
    Private myDTResult As DataTable
    Private myRevNo As String = String.Empty

    Private myBudgetOrderNoFilter As String = String.Empty
    Private myBudgetOrderNameFilter As String = String.Empty
    Private myTransferTypeFilter As String = String.Empty
    Private myFromOrderNoFilter As String = String.Empty
    Private myToOrderNoFilter As String = String.Empty
    Private myAccountNoFilter As String = String.Empty
    Private myAccountNameFilter As String = String.Empty
#End Region

#Region "Properties"
    Public Property BudgetYear() As String
        Get
            Return myBudgetYear
        End Get
        Set(ByVal value As String)
            myBudgetYear = value
        End Set
    End Property
    Public Property ProjectNo() As String
        Get
            Return myProjectNo
        End Get
        Set(ByVal value As String)
            myProjectNo = value
        End Set
    End Property
    Public Property PeriodType() As String
        Get
            Return myPeriodType
        End Get
        Set(ByVal value As String)
            myPeriodType = value
        End Set
    End Property
    Public Property BudgetOrderNo() As String
        Get
            Return myBudgetOrderNo
        End Get
        Set(ByVal value As String)
            myBudgetOrderNo = value
        End Set
    End Property
    Public Property TransferType() As String
        Get
            Return myTransferType
        End Get
        Set(ByVal value As String)
            myTransferType = value
        End Set
    End Property
    Public Property TransferRate() As String
        Get
            Return myTransferRate
        End Get
        Set(ByVal value As String)
            myTransferRate = value
        End Set
    End Property
    Public Property FromOrderNo() As String
        Get
            Return myFromOrderNo
        End Get
        Set(ByVal value As String)
            myFromOrderNo = value
        End Set
    End Property
    Public Property ToOrderNo() As String
        Get
            Return myToOrderNo
        End Get
        Set(ByVal value As String)
            myToOrderNo = value
        End Set
    End Property
    Public Property CreateUserID() As String
        Get
            Return myCreateUserID
        End Get
        Set(ByVal value As String)
            myCreateUserID = value
        End Set
    End Property
    Public Property CreateDate() As String
        Get
            Return myCreateDate
        End Get
        Set(ByVal value As String)
            myCreateDate = value
        End Set
    End Property
    Public Property UpdateUserID() As String
        Get
            Return myUpdateUserID
        End Get
        Set(ByVal value As String)
            myUpdateUserID = value
        End Set
    End Property
    Public Property DTResult() As DataTable
        Get
            Return myDTResult
        End Get
        Set(ByVal value As DataTable)
            myDTResult = value
        End Set
    End Property
    Public Property RevNo() As String
        Get
            Return myRevNo
        End Get
        Set(ByVal value As String)
            myRevNo = value
        End Set
    End Property

    Public Property BudgetOrderNoFilter() As String
        Get
            Return myBudgetOrderNoFilter
        End Get
        Set(ByVal value As String)
            myBudgetOrderNoFilter = value
        End Set
    End Property
    Public Property BudgetOrderNameFilter() As String
        Get
            Return myBudgetOrderNameFilter
        End Get
        Set(ByVal value As String)
            myBudgetOrderNameFilter = value
        End Set
    End Property
    Public Property TransferTypeFilter() As String
        Get
            Return myTransferTypeFilter
        End Get
        Set(ByVal value As String)
            myTransferTypeFilter = value
        End Set
    End Property
    Public Property FromOrderNoFilter() As String
        Get
            Return myFromOrderNoFilter
        End Get
        Set(ByVal value As String)
            myFromOrderNoFilter = value
        End Set
    End Property
    Public Property ToOrderNoFilter() As String
        Get
            Return myToOrderNoFilter
        End Get
        Set(ByVal value As String)
            myToOrderNoFilter = value
        End Set
    End Property
    Public Property AccountNoFilter() As String
        Get
            Return myAccountNoFilter
        End Get
        Set(ByVal value As String)
            myAccountNoFilter = value
        End Set
    End Property
    Public Property AccountNameFilter() As String
        Get
            Return myAccountNameFilter
        End Get
        Set(ByVal value As String)
            myAccountNameFilter = value
        End Set
    End Property

#End Region

#Region "Function"

    Public Function select001() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String
        Dim strWhere As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_TRANSFER_MASTER", "SELECT001")
            strSQL = strSQL.Replace("@P_BUDGET_YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@P_PERIOD_TYPE", Me.PeriodType)
            strSQL = strSQL.Replace("@P_PROJECT_NO", Me.ProjectNo)

            '// (1) BudgetOrderNo
            If Not Me.BudgetOrderNoFilter.Equals(String.Empty) Then

                strWhere &= " AND TM.BUDGET_ORDER_NO LIKE '%" & Me.BudgetOrderNoFilter.Replace("'", "''") & "%' "

            End If

            '// (2) BudgetOrderName
            If Not Me.BudgetOrderNameFilter.Equals(String.Empty) Then

                strWhere &= " AND BO0.BUDGET_ORDER_NAME LIKE '%" & Me.BudgetOrderNameFilter.Replace("'", "''") & "%' "

            End If

            '// (3) AccountNo
            If Not Me.AccountNoFilter.Equals(String.Empty) Then

                strWhere &= " AND BO0.ACCOUNT_NO LIKE '%" & Me.AccountNoFilter.Replace("'", "''") & "%' "

            End If

            '// (4) AccountName
            If Not Me.AccountNameFilter.Equals(String.Empty) Then

                strWhere &= " AND ACC.ACCOUNT_NAME LIKE '%" & Me.AccountNameFilter.Replace("'", "''") & "%' "

            End If

            '// (5) TransferType
            If Not Me.TransferTypeFilter.Equals(String.Empty) Then

                strWhere &= " AND TM.TRANSFER_TYPE = " & Me.TransferTypeFilter

            End If

            '// (6) FromOrderNo
            If Not Me.FromOrderNoFilter.Equals(String.Empty) Then

                strWhere &= " AND TM.FROM_ORDER_NO LIKE '%" & Me.FromOrderNoFilter.Replace("'", "''") & "%' "

            End If

            '// (7) ToOrderNo
            If Not Me.ToOrderNoFilter.Equals(String.Empty) Then

                strWhere &= " AND TM.TO_ORDER_NO LIKE '%" & Me.ToOrderNoFilter.Replace("'", "''") & "%' "

            End If


            If strWhere.Equals(String.Empty) Then
                strSQL = strSQL.Replace("@Where", "")
            Else
                strSQL = strSQL.Replace("@Where", strWhere)
            End If

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.DTResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_TRANSFER_MASTER.Select001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False
        End Try
    End Function

    Public Function select002(ByVal pConn As SqlConnection, _
                              ByVal pTrans As SqlTransaction) As Boolean
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            Dim cmd As SqlCommand = pConn.CreateCommand

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_TRANSFER_MASTER", "SELECT002")
            strSQL = strSQL.Replace("@P_BUDGET_YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@P_PERIOD_TYPE", Me.PeriodType)
            strSQL = strSQL.Replace("@P_BUDGET_ORDER_NO", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@P_PROJECT_NO", Me.ProjectNo)

            cmd.CommandText = strSQL
            cmd.Transaction = pTrans

            da = New SqlDataAdapter(cmd)
            dt = New DataTable
            da.Fill(dt)

            Me.DTResult = dt

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_TRANSFER_MASTER.Select002] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            Return False
        End Try
    End Function

    Public Function Select003() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_TRANSFER_MASTER", "SELECT003")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.DTResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_TRANSFER_MASTER.Select003] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False
        End Try
    End Function

    Public Function Select004() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_TRANSFER_MASTER", "SELECT004")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.DTResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_TRANSFER_MASTER.Select004] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False
        End Try
    End Function

    Public Function Insert001(ByVal pConn As SqlConnection, _
                              ByVal pTrans As SqlTransaction) As Boolean
        Dim cmd As SqlCommand
        Dim strSQL As String
        Dim intRtn As Integer

        Try
            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_TRANSFER_MASTER", "INSERT001")
            strSQL = strSQL.Replace("@P_BUDGET_YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@P_PERIOD_TYPE", Me.PeriodType)
            strSQL = strSQL.Replace("@P_BUDGET_ORDER_NO", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@P_TRANSFER_TYPE", Me.TransferType)
            strSQL = strSQL.Replace("@P_TRANSFER_RATE", Me.TransferRate)
            strSQL = strSQL.Replace("@P_FROM_ORDER_NO", Me.FromOrderNo)
            strSQL = strSQL.Replace("@P_TO_ORDER_NO", Me.ToOrderNo)
            strSQL = strSQL.Replace("@P_CREATE_USER_ID", Me.CreateUserID)
            strSQL = strSQL.Replace("@P_CREATE_DATE", Me.CreateDate)
            strSQL = strSQL.Replace("@P_UPDATE_USER_ID", Me.UpdateUserID)
            strSQL = strSQL.Replace("@P_PROJECT_NO", Me.ProjectNo)

            cmd = New SqlCommand(strSQL, pConn, pTrans)
            intRtn = cmd.ExecuteNonQuery()

            If intRtn > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Insert001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Public Function Update001(ByVal pConn As SqlConnection, _
                              ByVal pTrans As SqlTransaction) As Boolean
        Dim cmd As SqlCommand
        Dim strSQL As String
        Dim intRtn As Integer

        Try

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_TRANSFER_MASTER", "UPDATE001")
            strSQL = strSQL.Replace("@P_BUDGET_YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@P_PERIOD_TYPE", Me.PeriodType)
            strSQL = strSQL.Replace("@P_BUDGET_ORDER_NO", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@P_TRANSFER_TYPE", Me.TransferType)
            strSQL = strSQL.Replace("@P_TRANSFER_RATE", Me.TransferRate)
            strSQL = strSQL.Replace("@P_FROM_ORDER_NO", Me.FromOrderNo)
            strSQL = strSQL.Replace("@P_TO_ORDER_NO", Me.ToOrderNo)
            strSQL = strSQL.Replace("@P_UPDATE_USER_ID", Me.UpdateUserID)
            strSQL = strSQL.Replace("@P_PROJECT_NO", Me.ProjectNo)

            cmd = New SqlCommand(strSQL, pConn, pTrans)
            intRtn = cmd.ExecuteNonQuery()

            If intRtn > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Update001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Public Function Delete001(ByVal pConn As SqlConnection, _
                              ByVal pTrans As SqlTransaction) As Boolean
        Dim cmd As SqlCommand
        Dim strSQL As String
        Dim intRtn As Integer

        Try

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_TRANSFER_MASTER", "DELETE001")
            strSQL = strSQL.Replace("@P_BUDGET_YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@P_PERIOD_TYPE", Me.PeriodType)
            strSQL = strSQL.Replace("@P_BUDGET_ORDER_NO", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@P_PROJECT_NO", Me.ProjectNo)

            cmd = New SqlCommand(strSQL, pConn, pTrans)
            intRtn = cmd.ExecuteNonQuery()

            If intRtn > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Delete001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

#End Region

End Class
