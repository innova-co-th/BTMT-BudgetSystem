Imports System.Data.SqlClient
Imports BudgetManual.BGCommon
Imports BudgetManual.BGConstant

Public Class BG_M_ASSET_GROUP

#Region "Variable"
    Private mydtResult As DataTable
    Private myAssetGroupNo As String = String.Empty
    Private myAssetGroupName As String = String.Empty
    Private myAssetProjectNo As String = String.Empty
    Private myAssetProjectName As String = String.Empty
    Private myAssetCategoryNo As String = String.Empty
    Private myAssetCategoryName As String = String.Empty
    Private myCreateUserId As String = String.Empty
    Private myCreateDate As String = String.Empty
    Private myUpdateUserId As String = String.Empty
    Private myUpdateDate As String = String.Empty
#End Region

#Region "Properties"
    Public Property DtResult() As DataTable
        Get
            Return mydtResult
        End Get
        Set(ByVal value As DataTable)
            mydtResult = value
        End Set
    End Property
    Public Property AssetGroupNo() As String
        Get
            Return myAssetGroupNo
        End Get
        Set(ByVal value As String)
            myAssetGroupNo = value
        End Set
    End Property
    Public Property AssetGroupName() As String
        Get
            Return myAssetGroupName
        End Get
        Set(ByVal value As String)
            myAssetGroupName = value
        End Set
    End Property
    Public Property AssetProjectNo() As String
        Get
            Return myAssetProjectNo
        End Get
        Set(ByVal value As String)
            myAssetProjectNo = value
        End Set
    End Property
    Public Property AssetProjectName() As String
        Get
            Return myAssetProjectName
        End Get
        Set(ByVal value As String)
            myAssetProjectName = value
        End Set
    End Property
    Public Property AssetCategoryNo() As String
        Get
            Return myAssetCategoryNo
        End Get
        Set(ByVal value As String)
            myAssetCategoryNo = value
        End Set
    End Property
    Public Property AssetCategoryName() As String
        Get
            Return myAssetCategoryName
        End Get
        Set(ByVal value As String)
            myAssetCategoryName = value
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ASSET_GROUP", "SELECT001")

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.DtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_ASSET_GROUP.Select001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ASSET_GROUP", "SELECT002")

            '// (1) Asset Group No.
            If Not Me.AssetGroupNo.Equals(String.Empty) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " ASSET_GROUP_NO LIKE '%" & Me.AssetGroupNo.Replace("'", "''") & "%' "

            End If

            '// (2) Asset Group Name.
            If Not Me.AssetGroupName.Equals(String.Empty) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " ASSET_GROUP_NAME LIKE '%" & Me.AssetGroupName.Replace("'", "''") & "%' "

            End If

            '// (3) Asset Project N0.
            If Not Me.AssetProjectNo.Equals("0") Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " ASSET_PROJECT = " & Me.AssetProjectNo

            End If

            '// (4) Asset Category No.
            If Not Me.AssetCategoryNo.Equals("0") Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " ASSET_CATEGORY = " & Me.AssetCategoryNo

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
            MessageBox.Show("[BG_M_ASSET_GROUP.Select002] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ASSET_GROUP", "SELECT003")
            strSQL = strSQL.Replace("@No", Me.AssetGroupNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.DtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_ASSET_GROUP.Select003] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Query asset category
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select004(Optional ByVal pConn As SqlConnection = Nothing) As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String
        Dim strWhere As String = String.Empty

        Try

            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ASSET_GROUP", "SELECT004")


            '// (1) Category No.
            If Not Me.AssetCategoryNo.Equals(String.Empty) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " ASSET_CATEGORY LIKE '%" & Me.AssetCategoryNo.Replace("'", "''") & "%' "

            End If

            '// (2) Category Name.
            If Not Me.AssetCategoryName.Equals(String.Empty) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " ASSET_CATEGORY_TXT LIKE '%" & Me.AssetCategoryName.Replace("'", "''") & "%' "

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
            MessageBox.Show("[BG_M_ASSET_GROUP.Select004] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "SELECT005"
    ''' <summary>
    ''' Query asset project
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select005(Optional ByVal pConn As SqlConnection = Nothing) As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String
        Dim strWhere As String = String.Empty

        Try

            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ASSET_GROUP", "SELECT005")


            '// (1) Project No.
            If Not Me.AssetProjectNo.Equals(String.Empty) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " ASSET_PROJECT LIKE '%" & Me.AssetProjectNo.Replace("'", "''") & "%' "

            End If

            '// (2) Project Name.
            If Not Me.AssetProjectName.Equals(String.Empty) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " ASSET_PROJECT_TXT LIKE '%" & Me.AssetProjectName.Replace("'", "''") & "%' "

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
            MessageBox.Show("[BG_M_ASSET_GROUP.Select005] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "SELECT006"
    ''' <summary>
    ''' Query asset category data
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select006(Optional ByVal pConn As SqlConnection = Nothing) As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ASSET_GROUP", "SELECT006")
            strSQL = strSQL.Replace("@No", Me.AssetCategoryNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.DtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_ASSET_GROUP.Select006] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "SELECT007"
    ''' <summary>
    ''' Query asset project data
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select007(Optional ByVal pConn As SqlConnection = Nothing) As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ASSET_GROUP", "SELECT007")
            strSQL = strSQL.Replace("@No", Me.AssetProjectNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.DtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_ASSET_GROUP.Select007] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ASSET_GROUP", "INSERT001")
            strSQL = strSQL.Replace("@No", Me.AssetGroupNo)
            strSQL = strSQL.Replace("@Name", Me.AssetGroupName.Replace("'", "''"))
            strSQL = strSQL.Replace("@proj", Me.AssetProjectNo)
            strSQL = strSQL.Replace("@cate", Me.AssetCategoryNo)
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
            MessageBox.Show("[BG_M_ASSET_GROUP.Insert001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

#Region "Insert002"
    ''' <summary>
    ''' Insert002
    ''' </summary>
    ''' <param name="pConn">Connection</param>
    ''' <param name="pTrans">Transaction</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Insert002(Optional ByVal pConn As SqlConnection = Nothing, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ASSET_GROUP", "INSERT002")
            strSQL = strSQL.Replace("@No", Me.AssetCategoryNo)
            strSQL = strSQL.Replace("@Name", Me.AssetCategoryName.Replace("'", "''"))
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
            MessageBox.Show("[BG_M_ASSET_GROUP.Insert002] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

#Region "Insert003"
    ''' <summary>
    ''' Insert003
    ''' </summary>
    ''' <param name="pConn">Connection</param>
    ''' <param name="pTrans">Transaction</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Insert003(Optional ByVal pConn As SqlConnection = Nothing, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ASSET_GROUP", "INSERT003")
            strSQL = strSQL.Replace("@No", Me.AssetProjectNo)
            strSQL = strSQL.Replace("@Name", Me.AssetProjectName.Replace("'", "''"))
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
            MessageBox.Show("[BG_M_ASSET_GROUP.Insert003] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ASSET_GROUP", "UPDATE001")
            strSQL = strSQL.Replace("@No", Me.AssetGroupNo)
            strSQL = strSQL.Replace("@Name", Me.AssetGroupName.Replace("'", "''"))
            strSQL = strSQL.Replace("@proj", Me.AssetProjectNo)
            strSQL = strSQL.Replace("@cate", Me.AssetCategoryNo)
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
            MessageBox.Show("[BG_M_ASSET_GROUP.Update001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

#Region "Update002"
    ''' <summary>
    ''' Update002
    ''' </summary>
    ''' <param name="pConn">Connection</param>
    ''' <param name="pTrans">Transaction</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Update002(Optional ByVal pConn As SqlConnection = Nothing, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ASSET_GROUP", "UPDATE002")
            strSQL = strSQL.Replace("@No", Me.AssetCategoryNo)
            strSQL = strSQL.Replace("@Name", Me.AssetCategoryName.Replace("'", "''"))
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
            MessageBox.Show("[BG_M_ASSET_GROUP.Update002] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

#Region "Update003"
    ''' <summary>
    ''' Update003
    ''' </summary>
    ''' <param name="pConn">Connection</param>
    ''' <param name="pTrans">Transaction</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Update003(Optional ByVal pConn As SqlConnection = Nothing, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ASSET_GROUP", "UPDATE003")
            strSQL = strSQL.Replace("@No", Me.AssetProjectNo)
            strSQL = strSQL.Replace("@Name", Me.AssetProjectName.Replace("'", "''"))
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
            MessageBox.Show("[BG_M_ASSET_GROUP.Update003] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ASSET_GROUP", "DELETE001")
            strSQL = strSQL.Replace("@No", Me.AssetGroupNo)

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
            MessageBox.Show("[BG_M_ASSET_GROUP.Delete001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ASSET_GROUP", "DELETE002")
            strSQL = strSQL.Replace("@No", Me.AssetCategoryNo)

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
            MessageBox.Show("[BG_M_ASSET_GROUP.Delete002] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_ASSET_GROUP", "DELETE003")
            strSQL = strSQL.Replace("@No", Me.AssetProjectNo)

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
            MessageBox.Show("[BG_M_ASSET_GROUP.Delete003] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
