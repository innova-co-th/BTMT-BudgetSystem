Imports System.Data.SqlClient
Imports BudgetManual.BGCommon
Imports BudgetManual.BGConstant

Public Class BG_M_ACCOUNT

#Region "Variable"
    Private myDtResult As DataTable
    Private myAccountNo As String = String.Empty
    Private myPicNo As String = String.Empty
    Private myAccountName As String = String.Empty
    Private myCreateUserId As String = String.Empty
    Private myCreateDate As String = String.Empty
    Private myUpdateUserId As String = String.Empty
    Private myUpdateDate As String = String.Empty
#End Region

#Region "Property"
    Public Property DtResult() As DataTable
        Get
            Return myDtResult
        End Get
        Set(ByVal value As DataTable)
            myDtResult = value
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
    Public Property PicNo() As String
        Get
            Return myPicNo
        End Get
        Set(ByVal value As String)
            myPicNo = value
        End Set
    End Property
    Public Property AccountName() As String
        Get
            Return myAccountName
        End Get
        Set(ByVal value As String)
            myAccountName = value
        End Set
    End Property
    Public Property CreateUserId() As String
        Get
            Return myCreateUserId
        End Get
        Set(ByVal value As String)
            myCreateUserId = value
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
    Public Property UpdateUserId() As String
        Get
            Return myUpdateUserId
        End Get
        Set(ByVal value As String)
            myUpdateUserId = value
        End Set
    End Property
    Public Property UpdateDate() As String
        Get
            Return myUpdateDate
        End Get
        Set(ByVal value As String)
            myUpdateDate = value
        End Set
    End Property
#End Region

#Region "Function"

#Region "SELECT001"
    ''' <summary>
    ''' Query account data
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ACCOUNT", "SELECT001")

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.DtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_ACCOUNT.Select001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "SELECT002"
    ''' <summary>
    ''' Query account data
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select002() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String
        Dim strWhere As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ACCOUNT", "SELECT002")

            '// (1) Account No.
            If Not Me.AccountNo.Equals(String.Empty) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " ACCOUNT_NO LIKE '%" & Me.AccountNo.Replace("'", "''") & "%' "

            End If

            '// (2) Account Name.
            If Not Me.AccountName.Equals(String.Empty) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " ACCOUNT_NAME LIKE '%" & Me.AccountName.Replace("'", "''") & "%' "

            End If

            If strWhere.Equals(String.Empty) Then
                strSQL = strSQL.Replace("@Where", "")
            Else
                strSQL = strSQL.Replace("@Where", " WHERE " & strWhere)
            End If

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.DtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_ACCOUNT.Select002] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "SELECT003"
    ''' <summary>
    ''' Query account data
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select003(Optional ByVal pConn As SqlConnection = Nothing) As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try

            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ACCOUNT", "SELECT003")
            strSQL = strSQL.Replace("@No", Me.AccountNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.DtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_ACCOUNT.Select003] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "SELECT004"
    ''' <summary>
    ''' Query account data by PIC no
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select004(Optional ByVal pConn As SqlConnection = Nothing) As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try

            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ACCOUNT", "SELECT004")
            strSQL = strSQL.Replace("@PicNo", Me.PicNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.DtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_ACCOUNT.Select004] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ACCOUNT", "INSERT001")
            strSQL = strSQL.Replace("@No", Me.AccountNo)
            strSQL = strSQL.Replace("@Name", Me.AccountName.Replace("'", "''"))
            strSQL = strSQL.Replace("@UserId", Me.CreateUserId)

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
            MessageBox.Show("[BG_M_ACCOUNT.Insert001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ACCOUNT", "UPDATE001")
            strSQL = strSQL.Replace("@No", Me.AccountNo)
            strSQL = strSQL.Replace("@Name", Me.AccountName.Replace("'", "''"))
            strSQL = strSQL.Replace("@UserId", Me.UpdateUserId)

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
            MessageBox.Show("[BG_M_ACCOUNT.Update001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ACCOUNT", "DELETE001")
            strSQL = strSQL.Replace("@No", Me.AccountNo)

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
            MessageBox.Show("[BG_M_ACCOUNT.Delete001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
