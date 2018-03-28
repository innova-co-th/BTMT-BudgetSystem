Imports System.Data.SqlClient
Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class BG_T_BUDGET_DATA_REINPUT

#Region "Variable"
    Private myDtResult As DataTable
    Private myUserPIC As String = String.Empty
    Private myUserId As String = String.Empty
    Private myBudgetYear As String = String.Empty
    Private myPeriodType As String = String.Empty
    Private myBudgetType As String = String.Empty
    Private myRevNo As String = String.Empty
    Private myItemStatus As String = String.Empty
    Private myRevNo2 As String = String.Empty
    Private myProjectNo As String = String.Empty
    Private myBudgetOrderNo As String = String.Empty
    Private myPersonInChargeNo As String = String.Empty
#End Region

#Region "Property"

#Region "dtResult"
    Property dtResult() As DataTable
        Get
            Return myDtResult
        End Get
        Set(ByVal value As DataTable)
            myDtResult = value
        End Set
    End Property
#End Region

#Region "UserPIC"
    Property UserPIC() As String
        Get
            Return myUserPIC
        End Get
        Set(ByVal value As String)
            myUserPIC = value
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

#Region "BudgetType"
    Property BudgetType() As String
        Get
            Return myBudgetType
        End Get
        Set(ByVal value As String)
            myBudgetType = value
        End Set
    End Property
#End Region

#Region "RevNo"
    Property RevNo() As String
        Get
            Return myRevNo
        End Get
        Set(ByVal value As String)
            myRevNo = value
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

#Region "RevNo2"
    Property RevNo2() As String
        Get
            Return myRevNo2
        End Get
        Set(ByVal value As String)
            myRevNo2 = value
        End Set
    End Property
#End Region

#Region "ItemStatus"
    Public Property ItemStatus() As String
        Get
            Return myItemStatus
        End Get
        Set(ByVal value As String)
            myItemStatus = value
        End Set
    End Property
#End Region

#Region "BudgetOrderNo"
    Property BudgetOrderNo() As String
        Get
            Return myBudgetOrderNo
        End Get
        Set(ByVal value As String)
            myBudgetOrderNo = value
        End Set
    End Property
#End Region

#Region "PersonInChargeNo"
    Property PersonInChargeNo() As String
        Get
            Return myPersonInChargeNo
        End Get
        Set(ByVal value As String)
            myPersonInChargeNo = value
        End Set
    End Property
#End Region


#End Region

#Region "Function"

#Region "Select001"
    ''' <summary>
    ''' Select all budget header data
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select001() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA_REINPUT", "SELECT001")
            strSQL = strSQL.Replace("@P_BUDGET_YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@P_PERIOD_TYPE", Me.PeriodType)
            strSQL = strSQL.Replace("@P_REV_NO", Me.RevNo)
            strSQL = strSQL.Replace("@P_PROJECT_NO", Me.ProjectNo)
            strSQL = strSQL.Replace("@P_PERSON_IN_CHARGE_NO", Me.ProjectNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA_REINPUT.Select001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function


#End Region

#Region "Select002"
    Public Function Select002() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA_REINPUT", "SELECT002")
            strSQL = strSQL.Replace("@P_BUDGET_YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@P_PERIOD_TYPE", Me.PeriodType)
            strSQL = strSQL.Replace("@P_BUDGET_TYPE", Me.BudgetType)
            'strSQL = strSQL.Replace("@UserPIC", Me.UserPIC)
            strSQL = strSQL.Replace("@P_REV_NO", Me.RevNo)
            strSQL = strSQL.Replace("@P_ITEM_STATUS", Me.ItemStatus)
            strSQL = strSQL.Replace("@P_PROJECT_NO", Me.ProjectNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_HEADER.Select004] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select003"
    Public Function Select003() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA_REINPUT", "SELECT003")
            strSQL = strSQL.Replace("@P_BUDGET_YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@P_PERIOD_TYPE", Me.PeriodType)
            strSQL = strSQL.Replace("@P_BUDGET_TYPE", Me.BudgetType)
            'strSQL = strSQL.Replace("@P_PERSON_IN_CHARGE_NO", Me.PersonInChargeNo)
            strSQL = strSQL.Replace("@P_PERSON_IN_CHARGE_NO", Me.UserPIC)
            strSQL = strSQL.Replace("@P_REV_NO", Me.RevNo)
            strSQL = strSQL.Replace("@P_ITEM_STATUS", Me.ItemStatus)
            strSQL = strSQL.Replace("@P_PROJECT_NO", Me.ProjectNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_HEADER.Select004] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select004"
    Public Function Select004() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA_REINPUT", "SELECT004")
            strSQL = strSQL.Replace("@P_BUDGET_YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@P_PERIOD_TYPE", Me.PeriodType)
            strSQL = strSQL.Replace("@P_BUDGET_TYPE", Me.BudgetType)
            'strSQL = strSQL.Replace("@UserPIC", Me.UserPIC)
            strSQL = strSQL.Replace("@P_REV_NO", Me.RevNo)
            'strSQL = strSQL.Replace("@P_ITEM_STATUS", Me.ItemStatus)
            strSQL = strSQL.Replace("@P_PROJECT_NO", Me.ProjectNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA_REINPUT.Select004] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select005"
    Public Function Select005() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA_REINPUT", "SELECT005")
            strSQL = strSQL.Replace("@P_BUDGET_YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@P_PERIOD_TYPE", Me.PeriodType)
            strSQL = strSQL.Replace("@P_BUDGET_TYPE", Me.BudgetType)
            strSQL = strSQL.Replace("@P_PERSON_IN_CHARGE_NO", Me.PersonInChargeNo)
            strSQL = strSQL.Replace("@P_REV_NO", Me.RevNo)
            'strSQL = strSQL.Replace("@P_ITEM_STATUS", Me.ItemStatus)
            strSQL = strSQL.Replace("@P_PROJECT_NO", Me.ProjectNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_HEADER.Select005] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Insert001"
    ''' <summary>
    ''' Insert001
    ''' </summary>
    ''' <param name="pConn">Connection</param>
    ''' <param name="pTrans">Transaction</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Insert001(Optional ByVal pConn As SqlConnection = Nothing, _
                              Optional ByVal pTrans As SqlTransaction = Nothing) As Boolean
        Dim conn As SqlConnection = Nothing
        Dim cmd As SqlCommand
        Dim strSQL As String
        Dim intRtn As Integer

        Try
            If pConn Is Nothing Or pTrans Is Nothing Then
                conn = New SqlConnection(My.Settings.ConnStr)
                conn.Open()
            End If

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA_REINPUT", "INSERT001")
            strSQL = strSQL.Replace("@P_BUDGET_YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@P_PERIOD_TYPE", Me.PeriodType)
            strSQL = strSQL.Replace("@P_REV_NO", Me.RevNo)
            strSQL = strSQL.Replace("@P_PROJECT_NO", Me.ProjectNo)
            strSQL = strSQL.Replace("@P_BUDGET_ORDER_NO", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@P_ITEM_STATUS", Me.ItemStatus)
            strSQL = strSQL.Replace("@P_CREATE_USER_ID", Me.UserId)

            If pConn IsNot Nothing And pTrans IsNot Nothing Then
                cmd = New SqlCommand(strSQL, pConn, pTrans)
                intRtn = cmd.ExecuteNonQuery()

            Else
                cmd = New SqlCommand(strSQL, conn)
                intRtn = cmd.ExecuteNonQuery()

                If conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If

            End If

            If intRtn > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA_REINPUT.Insert001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Update001"
    ''' <summary>
    ''' Update001
    ''' </summary>
    ''' <param name="pConn">Connection</param>
    ''' <param name="pTrans">Transaction</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Update001(Optional ByVal pConn As SqlConnection = Nothing, _
                              Optional ByVal pTrans As SqlTransaction = Nothing) As Boolean
        Dim conn As SqlConnection = Nothing
        Dim cmd As SqlCommand
        Dim strSQL As String
        Dim intRtn As Integer

        Try
            If pConn Is Nothing Or pTrans Is Nothing Then
                conn = New SqlConnection(My.Settings.ConnStr)
                conn.Open()
            End If

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA_REINPUT", "UPDATE001")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@UserPIC", Me.UserPIC)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@Status", Me.ItemStatus)
            strSQL = strSQL.Replace("@UserId", Me.UserId)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            If pConn IsNot Nothing And pTrans IsNot Nothing Then
                cmd = New SqlCommand(strSQL, pConn, pTrans)
                intRtn = cmd.ExecuteNonQuery()

            Else
                cmd = New SqlCommand(strSQL, conn)
                intRtn = cmd.ExecuteNonQuery()

                If conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If

            End If

            If intRtn > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA_REINPUT.Update001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Delete001"
    ''' <summary>
    ''' Delete Rev
    ''' </summary>
    ''' <param name="pConn">Connection</param>
    ''' <param name="pTrans">Transaction</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Delete001(Optional ByVal pConn As SqlConnection = Nothing, _
                              Optional ByVal pTrans As SqlTransaction = Nothing) As Boolean
        Dim conn As SqlConnection = Nothing
        Dim cmd As SqlCommand
        Dim strSQL As String
        Dim intRtn As Integer

        Try
            If pConn Is Nothing Or pTrans Is Nothing Then
                conn = New SqlConnection(My.Settings.ConnStr)
                conn.Open()
            End If

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA_REINPUT", "DELETE001")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@STATUS", Me.ItemStatus)

            If pConn IsNot Nothing And pTrans IsNot Nothing Then
                cmd = New SqlCommand(strSQL, pConn, pTrans)
                intRtn = cmd.ExecuteNonQuery()

            Else
                cmd = New SqlCommand(strSQL, conn)
                intRtn = cmd.ExecuteNonQuery()

                If conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If

            End If

            If intRtn >= 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA_REINPUT.Delete001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Delete002"
    ''' <summary>
    ''' Delete Rev
    ''' </summary>
    ''' <param name="pConn">Connection</param>
    ''' <param name="pTrans">Transaction</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Delete002(Optional ByVal pConn As SqlConnection = Nothing, _
                              Optional ByVal pTrans As SqlTransaction = Nothing) As Boolean
        Dim conn As SqlConnection = Nothing
        Dim cmd As SqlCommand
        Dim strSQL As String
        Dim intRtn As Integer

        Try
            If pConn Is Nothing Or pTrans Is Nothing Then
                conn = New SqlConnection(My.Settings.ConnStr)
                conn.Open()
            End If

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA_REINPUT", "DELETE002")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)


            If pConn IsNot Nothing And pTrans IsNot Nothing Then
                cmd = New SqlCommand(strSQL, pConn, pTrans)
                intRtn = cmd.ExecuteNonQuery()

            Else
                cmd = New SqlCommand(strSQL, conn)
                intRtn = cmd.ExecuteNonQuery()

                If conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If

            End If

            If intRtn >= 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA_REINPUT.Delete002] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Delete003"
    ''' <summary>
    ''' Delete Rev
    ''' </summary>
    ''' <param name="pConn">Connection</param>
    ''' <param name="pTrans">Transaction</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Delete003(Optional ByVal pConn As SqlConnection = Nothing, _
                              Optional ByVal pTrans As SqlTransaction = Nothing) As Boolean
        Dim conn As SqlConnection = Nothing
        Dim cmd As SqlCommand
        Dim strSQL As String
        Dim intRtn As Integer

        Try
            If pConn Is Nothing Or pTrans Is Nothing Then
                conn = New SqlConnection(My.Settings.ConnStr)
                conn.Open()
            End If

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA_REINPUT", "DELETE003")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@UserPIC", Me.PersonInChargeNo)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)


            If pConn IsNot Nothing And pTrans IsNot Nothing Then
                cmd = New SqlCommand(strSQL, pConn, pTrans)
                intRtn = cmd.ExecuteNonQuery()

            Else
                cmd = New SqlCommand(strSQL, conn)
                intRtn = cmd.ExecuteNonQuery()

                If conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If

            End If

            If intRtn >= 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA_REINPUT.Delete003] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function
#End Region

#End Region
End Class
