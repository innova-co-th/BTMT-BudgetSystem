Imports System.Data.SqlClient
Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class BG_T_INFORMATION

#Region "Variable"
    Private myFileTitle As String = String.Empty
    Private myFilePath As String = String.Empty
    Private myUserId As String = String.Empty
    Private myFileNo As String = String.Empty
    Private myDTResult As DataTable
#End Region

#Region "Property"

    Public Property FileNo() As String
        Get
            Return myFileNo
        End Get
        Set(ByVal value As String)
            myFileNo = value
        End Set
    End Property
    Public Property FileTitle() As String
        Get
            Return myFileTitle
        End Get
        Set(ByVal value As String)
            myFileTitle = value
        End Set
    End Property
    Public Property FilePath() As String
        Get
            Return myFilePath
        End Get
        Set(ByVal value As String)
            myFilePath = value
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
    Public Property UserId() As String
        Get
            Return myUserId
        End Get
        Set(ByVal value As String)
            myUserId = value
        End Set
    End Property

#End Region

#Region "Function"
    Public Function Select001() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_INFORMATION", "SELECT001")

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.DTResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_INFORMATION.Select001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    Public Function Select002() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_INFORMATION", "SELECT002")
            strSQL = strSQL.Replace("@P_FILE_NO", Me.FileNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.DTResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_INFORMATION.Select002] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_INFORMATION", "INSERT001")
            strSQL = strSQL.Replace("@P_USER_ID", Me.UserId)
            strSQL = strSQL.Replace("@P_FILE_TITLE", Me.FileTitle.Replace("'", "''"))
            strSQL = strSQL.Replace("@P_FILE_PATH", Me.FilePath.Replace("'", "''"))

            cmd = New SqlCommand(strSQL, pConn, pTrans)
            intRtn = cmd.ExecuteNonQuery()

            If intRtn > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show("[BG_T_INFORMATION.Insert001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_INFORMATION", "UPDATE001")
            strSQL = strSQL.Replace("@P_USER_ID", Me.UserId)
            strSQL = strSQL.Replace("@P_FILE_TITLE", Me.FileTitle.Replace("'", "''"))
            strSQL = strSQL.Replace("@P_FILE_PATH", Me.FilePath.Replace("'", "''"))
            strSQL = strSQL.Replace("@P_FILE_NO", Me.FileNo)

            cmd = New SqlCommand(strSQL, pConn, pTrans)
            intRtn = cmd.ExecuteNonQuery()

            If intRtn > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show("[BG_T_INFORMATION.Update001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_INFORMATION", "DELETE001")
            strSQL = strSQL.Replace("@P_FILE_NO", Me.FileNo)

            cmd = New SqlCommand(strSQL, pConn, pTrans)
            'Dim cmd As SqlCommand = pConn.CreateCommand
            'cmd.CommandText = strSQL
            'cmd.Transaction = pTrans

            intRtn = cmd.ExecuteNonQuery()

            If intRtn > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show("[BG_T_INFORMATION.Delete001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False

        End Try
    End Function

#End Region

End Class
