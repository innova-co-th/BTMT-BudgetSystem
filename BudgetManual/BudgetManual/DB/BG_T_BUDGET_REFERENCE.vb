Imports System.Data.SqlClient
Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class BG_T_BUDGET_REFERENCE

#Region "Variable"
    Private myDtResult As DataTable

    Private myBudgetYear As String = String.Empty
    Private myProjectNo As String = String.Empty
    Private myPeriodType As String = String.Empty
    Private myRevNo As String = String.Empty
    Private myRefBudgetYear As String = String.Empty
    Private myRefProjectNo As String = String.Empty
    Private myRefPeriodType As String = String.Empty
    Private myRefRevNo As String = String.Empty
    Private myCreateUserID As String = String.Empty
    
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

#Region "BudgetYear"
    Public Property BudgetYear() As String
        Get
            Return myBudgetYear
        End Get
        Set(ByVal value As String)
            myBudgetYear = value
        End Set
    End Property
#End Region

#Region "ProjectNo"
    Public Property ProjectNo() As String
        Get
            Return myProjectNo
        End Get
        Set(ByVal value As String)
            myProjectNo = value
        End Set
    End Property
#End Region

#Region "PeriodType"
    Public Property PeriodType() As String
        Get
            Return myPeriodType
        End Get
        Set(ByVal value As String)
            myPeriodType = value
        End Set
    End Property
#End Region

#Region "RevNo"
    Public Property RevNo() As String
        Get
            Return myRevNo
        End Get
        Set(ByVal value As String)
            myRevNo = value
        End Set
    End Property
#End Region

#Region "RefBudgetYear"
    Public Property RefBudgetYear() As String
        Get
            Return myRefBudgetYear
        End Get
        Set(ByVal value As String)
            myRefBudgetYear = value
        End Set
    End Property
#End Region

#Region "RefProjectNo"
    Public Property RefProjectNo() As String
        Get
            Return myRefProjectNo
        End Get
        Set(ByVal value As String)
            myRefProjectNo = value
        End Set
    End Property
#End Region

#Region "RefPeriodType"
    Public Property RefPeriodType() As String
        Get
            Return myRefPeriodType
        End Get
        Set(ByVal value As String)
            myRefPeriodType = value
        End Set
    End Property
#End Region

#Region "RefRevNo"
    Public Property RefRevNo() As String
        Get
            Return myRefRevNo
        End Get
        Set(ByVal value As String)
            myRefRevNo = value
        End Set
    End Property
#End Region

#Region "CreateUserID"
    Public Property CreateUserID() As String
        Get
            Return myCreateUserID
        End Get
        Set(ByVal value As String)
            myCreateUserID = value
        End Set
    End Property
#End Region

#End Region

#Region "Function"

#Region "Select001"
    ''' <summary>
    ''' Select Reference
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_REFERENCE", "Select001")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@RevNo", CStr(IIf(IsNumeric(Me.RevNo), Me.RevNo, "0")))
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@RefPeriodType", Me.RefPeriodType)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception

            MessageBox.Show("[BG_T_BUDGET_REFERENCE.Select001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Select Reference
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_REFERENCE", "Select002")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@RefPeriodType", Me.RefPeriodType)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception

            MessageBox.Show("[BG_T_BUDGET_REFERENCE.Select002] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Insert Budget Adjust Master
    ''' </summary>
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_REFERENCE", "INSERT001")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@RefBudgetYear", Me.RefBudgetYear)
            strSQL = strSQL.Replace("@RefPeriodType", Me.RefPeriodType)
            strSQL = strSQL.Replace("@RefProjectNo", Me.RefProjectNo)
            strSQL = strSQL.Replace("@RefRevNo", Me.RefRevNo)
            strSQL = strSQL.Replace("@CreateUserID", Me.CreateUserID)
            strSQL = strSQL.Replace("@UpdateUserID", Me.CreateUserID)

            

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
            MessageBox.Show("[BG_T_BUDGET_REFERENCE.Insert001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Update exist data
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_REFERENCE", "UPDATE001")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@RefBudgetYear", Me.RefBudgetYear)
            strSQL = strSQL.Replace("@RefPeriodType", Me.RefPeriodType)
            strSQL = strSQL.Replace("@RefProjectNo", Me.RefProjectNo)
            strSQL = strSQL.Replace("@RefRevNo", Me.RefRevNo)
            strSQL = strSQL.Replace("@UpdateUserID", Me.CreateUserID)

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
            MessageBox.Show("[BG_T_BUDGET_REFERENCE.Update001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
