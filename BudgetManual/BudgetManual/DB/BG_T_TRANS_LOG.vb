Imports System.Data.SqlClient
Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class BG_T_TRANS_LOG

#Region "Variable"
    Private myDtResult As DataTable
    Private myUserId As String = String.Empty
    Private myBudgetYear As String = String.Empty
    Private myProjectNo As String = String.Empty
    Private myPeriodType As String = String.Empty
    Private myUserPIC As String = String.Empty
    Private myBudgetType As String = String.Empty
    Private myRevNo As String = String.Empty
    Private myOperationCd As String = String.Empty
    Private myFromDate As String = String.Empty
    Private myToDate As String = String.Empty
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

#Region "OperationCd"
    Property OperationCd() As String
        Get
            Return myOperationCd
        End Get
        Set(ByVal value As String)
            myOperationCd = value
        End Set
    End Property
#End Region

#Region "FromDate"
    Property FromDate() As String
        Get
            Return myFromDate
        End Get
        Set(ByVal value As String)
            myFromDate = value
        End Set
    End Property
#End Region

#Region "ToDate"
    Property ToDate() As String
        Get
            Return myToDate
        End Get
        Set(ByVal value As String)
            myToDate = value
        End Set
    End Property
#End Region

#End Region

#Region "Function"

#Region "Select001"
    ''' <summary>
    ''' Select001
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_TRANS_LOG", "SELECT001")

            da = New SqlDataAdapter(strSQL, conn)
            da.SelectCommand.Parameters.Add("@BudgetYear", SqlDbType.Decimal).Value = Me.BudgetYear
            da.SelectCommand.Parameters.Add("@PeriodType", SqlDbType.TinyInt).Value = Me.PeriodType
            da.SelectCommand.Parameters.Add("@ProjectNo", SqlDbType.TinyInt).Value = Me.ProjectNo

            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_TRANS_LOG.Select001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select002"
    ''' <summary>
    ''' Select001
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select002() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_TRANS_LOG", "SELECT002")

            da = New SqlDataAdapter(strSQL, conn)
            da.SelectCommand.Parameters.Add("@FromDate", SqlDbType.DateTime).Value = Me.FromDate
            da.SelectCommand.Parameters.Add("@ToDate", SqlDbType.DateTime).Value = Me.ToDate

            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_TRANS_LOG.Select002] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_TRANS_LOG", "INSERT001")

            cmd = New SqlCommand(strSQL, conn)
            cmd.Parameters.Add("@UserId", SqlDbType.VarChar).Value = Me.UserId
            cmd.Parameters.Add("@OperationCd", SqlDbType.TinyInt).Value = Me.OperationCd
            cmd.Parameters.Add("@BudgetYear", SqlDbType.Decimal).Value = IIf(IsNumeric(Me.BudgetYear), Me.BudgetYear, Now.Year)
            cmd.Parameters.Add("@PeriodType", SqlDbType.TinyInt).Value = IIf(IsNumeric(Me.PeriodType), Me.PeriodType, 0)
            cmd.Parameters.Add("@UserPIC", SqlDbType.VarChar).Value = Me.UserPIC
            cmd.Parameters.Add("@BudgetType", SqlDbType.VarChar).Value = Me.BudgetType
            cmd.Parameters.Add("@RevNo", SqlDbType.SmallInt).Value = IIf(IsNumeric(Me.RevNo), Me.RevNo, 0)
            cmd.Parameters.Add("@ProjectNo", SqlDbType.TinyInt).Value = IIf(IsNumeric(Me.ProjectNo), Me.ProjectNo, 0)

            intRtn = cmd.ExecuteNonQuery()

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            If intRtn > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show("[BG_T_TRANS_LOG.Insert001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#End Region

End Class

