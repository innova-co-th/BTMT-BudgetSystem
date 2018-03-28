Imports System.Data.SqlClient
Imports BudgetManual.BGCommon
Imports BudgetManual.BGConstant

Public Class BG_M_USER

#Region "Variable"
    Private Const STRING_ALL As String = "All"

    Private myDtResult As DataTable
    Private myUserId As String = String.Empty
    Private myUserName As String = String.Empty
    Private myUserLevel As String = String.Empty
    Private myPassword As String = String.Empty
    Private myUserPIC As String = String.Empty
    Private myEmail As String = String.Empty
    Private myExpireFlg As String = String.Empty
    Private myUserId2 As String = String.Empty
    Private myUserPermissions As Boolean()
    Private myUserLevelName As String = String.Empty
    Private myUserPermissionsFilter As String = String.Empty
#End Region

#Region "Property"

#Region "UserLevelName"
    Property UserLevelName() As String
        Get
            Return myUserLevelName
        End Get
        Set(ByVal value As String)
            myUserLevelName = value
        End Set
    End Property
#End Region

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

#Region "UserName"
    Property UserName() As String
        Get
            Return myUserName
        End Get
        Set(ByVal value As String)
            myUserName = value
        End Set
    End Property
#End Region

#Region "UserLevel"
    Property UserLevel() As String
        Get
            Return myUserLevel
        End Get
        Set(ByVal value As String)
            myUserLevel = value
        End Set
    End Property
#End Region

#Region "Password"
    Property Password() As String
        Get
            Return myPassword
        End Get
        Set(ByVal value As String)
            myPassword = value
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

#Region "Email"
    Property Email() As String
        Get
            Return myEmail
        End Get
        Set(ByVal value As String)
            myEmail = value
        End Set
    End Property
#End Region

#Region "ExpireFlg"
    Property ExpireFlg() As String
        Get
            Return myExpireFlg
        End Get
        Set(ByVal value As String)
            myExpireFlg = value
        End Set
    End Property
#End Region

#Region "UserId2"
    Property UserId2() As String
        Get
            Return myUserId2
        End Get
        Set(ByVal value As String)
            myUserId2 = value
        End Set
    End Property
#End Region

#Region "UserPermissions"
    Property UserPermissions() As Boolean()
        Get
            Return myUserPermissions
        End Get
        Set(ByVal value As Boolean())
            myUserPermissions = value
        End Set
    End Property
#End Region

#Region "UserPermissions"
    Property UserPermissionsFilter() As String
        Get
            Return myUserPermissionsFilter
        End Get
        Set(ByVal value As String)
            myUserPermissionsFilter = value
        End Set
    End Property
#End Region

#End Region

#Region "Function"

#Region "Select001"
    ''' <summary>
    ''' Query user data by user id
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "SELECT001")
            strSQL = strSQL.Replace("@UserId", Me.UserId)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_USER.Select001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Query all user data
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "SELECT002")

            '// (1) UserID
            If Not Me.UserId.Equals(String.Empty) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If

                strWhere &= " US.USER_ID LIKE '%" & Me.UserId.Replace("'", "''") & "%' "

            End If

            '// (2) UserName
            If Not Me.UserName.Equals(String.Empty) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If

                strWhere &= " US.USER_NAME LIKE '%" & Me.UserName.Replace("'", "''") & "%' "

            End If

            '// (3) UserLevel
            If Not Me.UserLevel.Equals("") Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If

                strWhere &= " US.USER_LEVEL_ID = " & Me.UserLevel

            End If

            '// (4) UserPIC
            If Not Me.UserPIC.Equals("All") Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If

                strWhere &= " US.PERSON_IN_CHARGE_NO = '" & Me.UserPIC & "'"

            End If


            If strWhere.Equals(String.Empty) Then
                strSQL = strSQL.Replace("@Where", "")
            Else
                strSQL = strSQL.Replace("@Where", " WHERE " & strWhere)
            End If

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_USER.Select002] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select003"
    ''' <summary>
    ''' Query all user PIC
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select003() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "SELECT003")

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_USER.Select003] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select004"
    ''' <summary>
    ''' Query all user level
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select004() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "SELECT004")

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_USER.Select004] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select005"
    ''' <summary>
    ''' Query all user data
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select005() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "SELECT005")

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_USER.Select005] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select006"
    ''' <summary>
    ''' Select006
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select006() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "SELECT006")
            strSQL = strSQL.Replace("@UserPic", Me.UserPIC)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_USER.Select006] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select007"
    ''' <summary>
    ''' Select007
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select007() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "SELECT007")
            strSQL = strSQL.Replace("@UserPic", Me.UserPIC)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_USER.Select007] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select008"
    ''' <summary>
    ''' Select008
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select008() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "SELECT008")
            strSQL = strSQL.Replace("@UserPic", Me.UserPIC)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_USER.Select008] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select009"
    ''' <summary>
    ''' Select009
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select009() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "SELECT009")
            strSQL = strSQL.Replace("@UserPic", Me.UserPIC)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_USER.Select009] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select010"
    ''' <summary>
    ''' Select all approver
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select010() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "SELECT010")

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_USER.Select010] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select011"
    ''' <summary>
    ''' Select Login User
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select011() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "SELECT011")
            strSQL = strSQL.Replace("@UserPIC", Me.UserPIC)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_USER.Select011] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select012"
    ''' <summary>
    ''' Select Locked Pic
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select012() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "SELECT012")

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_USER.Select012] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

#End Region

#Region "Select013"
    ''' <summary>
    ''' Select User Level
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select013() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String
        Dim strWhere As String = String.Empty
        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "SELECT013")


            '// (1) UserLevelName
            If Not Me.UserLevel.Equals(String.Empty) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " USER_LEVEL_NAME LIKE '%" & Me.UserLevel.Replace("'", "''") & "%' "

            End If

            '// (2) UserPermission
            If Not Me.UserPermissionsFilter.Equals(STRING_ALL) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If

                Select Case Me.UserPermissionsFilter

                    Case BGConstant.enumPermissionCd.Adjust.ToString
                        strWhere &= " ADJUST = 'Y' "

                    Case BGConstant.enumPermissionCd.Approve.ToString
                        strWhere &= " APPROVE = 'Y' "

                    Case BGConstant.enumPermissionCd.Auth1.ToString
                        strWhere &= " AUTH1 = 'Y' "

                    Case BGConstant.enumPermissionCd.Auth2.ToString
                        strWhere &= " AUTH2 = 'Y' "

                    Case BGConstant.enumPermissionCd.Entry.ToString
                        strWhere &= " ENTRY = 'Y' "

                    Case BGConstant.enumPermissionCd.Export.ToString
                        strWhere &= " EXPORT = 'Y' "

                    Case BGConstant.enumPermissionCd.Import.ToString
                        strWhere &= " IMPORT = 'Y' "

                    Case BGConstant.enumPermissionCd.Master.ToString
                        strWhere &= " MASTER = 'Y' "

                    Case BGConstant.enumPermissionCd.Submit.ToString
                        strWhere &= " SUBMIT = 'Y' "

                    Case BGConstant.enumPermissionCd.System.ToString
                        strWhere &= " SYSTEM = 'Y' "

                    Case BGConstant.enumPermissionCd.View.ToString
                        strWhere &= " [VIEW] = 'Y' "

                    Case BGConstant.enumPermissionCd.DirectInput.ToString
                        strWhere &= " DIRECT_INPUT = 'Y' "

                    Case Else

                End Select

            End If

            If strWhere.Equals(String.Empty) Then
                strSQL = strSQL.Replace("@Where", "")
            Else
                strSQL = strSQL.Replace("@Where", " WHERE " & strWhere)
            End If


            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_USER.Select013] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

#End Region

#Region "Select015"
    ''' <summary>
    ''' Select Last User Level Id
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select015() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "SELECT015")

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_USER.Select015] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

#End Region

#Region "Select014"
    ''' <summary>
    ''' Select Login User
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select014() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "SELECT014")
            strSQL = strSQL.Replace("@UserID", Me.UserId)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_USER.Select014] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Update001"
    ''' <summary>
    ''' Update001
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Update001() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim cmd As SqlCommand
        Dim strSQL As String
        Dim intRtn As Integer

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "UPDATE001")
            strSQL = strSQL.Replace("@UserId", Me.UserId)
            strSQL = strSQL.Replace("@Password", Me.Password)

            cmd = New SqlCommand(strSQL, conn)
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
            MessageBox.Show("[BG_M_USER.Update001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Update002"
    ''' <summary>
    ''' Update002
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Update002(Optional ByVal pConn As SqlConnection = Nothing, _
                              Optional ByVal pTrans As SqlTransaction = Nothing) As Boolean
        Dim conn As SqlConnection = Nothing
        Dim cmd As SqlCommand
        Dim strSQL As String
        Dim intRtn As Integer

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "UPDATE002")
            strSQL = strSQL.Replace("@UserId1", Me.UserId)
            strSQL = strSQL.Replace("@UserLevel", Me.UserLevel)
            strSQL = strSQL.Replace("@UserName", Me.UserName.Replace("'", "''"))
            strSQL = strSQL.Replace("@Password", Me.Password)
            strSQL = strSQL.Replace("@UserPIC", Me.UserPIC)
            strSQL = strSQL.Replace("@Email", Me.Email)
            strSQL = strSQL.Replace("@ExpireFlg", Me.ExpireFlg)
            strSQL = strSQL.Replace("@UserId2", Me.UserId2)

            cmd = New SqlCommand(strSQL, conn)
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
            MessageBox.Show("[BG_M_USER.Update002] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Update003"
    ''' <summary>
    ''' Update User Level
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Update003(Optional ByVal pConn As SqlConnection = Nothing, _
                              Optional ByVal pTrans As SqlTransaction = Nothing) As Boolean
        Dim conn As SqlConnection = Nothing
        Dim cmd As SqlCommand
        Dim strSQL As String
        Dim intRtn As Integer

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "UPDATE003")
            strSQL = strSQL.Replace("@UserLevelName", Me.UserLevelName)
            strSQL = strSQL.Replace("@UserLevel", Me.UserLevel)
            strSQL = strSQL.Replace("@Entry", CStr(IIf(Me.UserPermissions(0), "Y", "N")))
            strSQL = strSQL.Replace("@Submit", CStr(IIf(Me.UserPermissions(1), "Y", "N")))
            strSQL = strSQL.Replace("@Approve", CStr(IIf(Me.UserPermissions(2), "Y", "N")))
            strSQL = strSQL.Replace("@Adjust", CStr(IIf(Me.UserPermissions(3), "Y", "N")))
            strSQL = strSQL.Replace("@Auth1", CStr(IIf(Me.UserPermissions(4), "Y", "N")))
            strSQL = strSQL.Replace("@Auth2", CStr(IIf(Me.UserPermissions(5), "Y", "N")))
            strSQL = strSQL.Replace("@Import", CStr(IIf(Me.UserPermissions(6), "Y", "N")))
            strSQL = strSQL.Replace("@Export", CStr(IIf(Me.UserPermissions(7), "Y", "N")))
            strSQL = strSQL.Replace("@Master", CStr(IIf(Me.UserPermissions(8), "Y", "N")))
            strSQL = strSQL.Replace("@System", CStr(IIf(Me.UserPermissions(9), "Y", "N")))
            strSQL = strSQL.Replace("@View", CStr(IIf(Me.UserPermissions(10), "Y", "N")))
            strSQL = strSQL.Replace("@DirectInput", CStr(IIf(Me.UserPermissions(11), "Y", "N")))
            strSQL = strSQL.Replace("@UserId", Me.UserId)


            cmd = New SqlCommand(strSQL, conn)
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
            MessageBox.Show("[BG_M_USER.Update003] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "INSERT001")
            strSQL = strSQL.Replace("@UserId1", Me.UserId)
            strSQL = strSQL.Replace("@UserLevel", Me.UserLevel)
            strSQL = strSQL.Replace("@UserName", Me.UserName.Replace("'", "''"))
            strSQL = strSQL.Replace("@Password", Me.Password)
            strSQL = strSQL.Replace("@UserPIC", Me.UserPIC)
            strSQL = strSQL.Replace("@Email", Me.Email)
            strSQL = strSQL.Replace("@ExpireFlg", Me.ExpireFlg)
            strSQL = strSQL.Replace("@UserId2", Me.UserId2)

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
            MessageBox.Show("[BG_M_USER.Insert001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "INSERT002")
            strSQL = strSQL.Replace("@UserPIC", Me.UserPIC)
            strSQL = strSQL.Replace("@UserId", Me.UserId)

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
            MessageBox.Show("[BG_M_USER.Insert002] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Add User Level
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Insert003(Optional ByVal pConn As SqlConnection = Nothing, _
                              Optional ByVal pTrans As SqlTransaction = Nothing) As Boolean
        Dim conn As SqlConnection = Nothing
        Dim cmd As SqlCommand
        Dim strSQL As String
        Dim intRtn As Integer

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "INSERT003")
            strSQL = strSQL.Replace("@UserLevelName", Me.UserLevelName)
            strSQL = strSQL.Replace("@UserLevel", Me.UserLevel)
            strSQL = strSQL.Replace("@Entry", CStr(IIf(Me.UserPermissions(0), "Y", "N")))
            strSQL = strSQL.Replace("@Submit", CStr(IIf(Me.UserPermissions(1), "Y", "N")))
            strSQL = strSQL.Replace("@Approve", CStr(IIf(Me.UserPermissions(2), "Y", "N")))
            strSQL = strSQL.Replace("@Adjust", CStr(IIf(Me.UserPermissions(3), "Y", "N")))
            strSQL = strSQL.Replace("@Auth1", CStr(IIf(Me.UserPermissions(4), "Y", "N")))
            strSQL = strSQL.Replace("@Auth2", CStr(IIf(Me.UserPermissions(5), "Y", "N")))
            strSQL = strSQL.Replace("@Import", CStr(IIf(Me.UserPermissions(6), "Y", "N")))
            strSQL = strSQL.Replace("@Export", CStr(IIf(Me.UserPermissions(7), "Y", "N")))
            strSQL = strSQL.Replace("@Master", CStr(IIf(Me.UserPermissions(8), "Y", "N")))
            strSQL = strSQL.Replace("@System", CStr(IIf(Me.UserPermissions(9), "Y", "N")))
            strSQL = strSQL.Replace("@View", CStr(IIf(Me.UserPermissions(10), "Y", "N")))
            strSQL = strSQL.Replace("@DirectInput", CStr(IIf(Me.UserPermissions(11), "Y", "N")))
            strSQL = strSQL.Replace("@UserId", Me.UserId)

            cmd = New SqlCommand(strSQL, conn)
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
            MessageBox.Show("[BG_M_USER.Insert003] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "DELETE001")
            strSQL = strSQL.Replace("@UserPIC", Me.UserPIC)
            strSQL = strSQL.Replace("@UserId", Me.UserId)

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
            MessageBox.Show("[BG_M_USER.Delete001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "DELETE002")
            strSQL = strSQL.Replace("@UserPIC", Me.UserPIC)

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
            MessageBox.Show("[BG_M_USER.Delete002] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_USER", "DELETE003")
            strSQL = strSQL.Replace("@UserLevel", Me.UserLevel)

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
            MessageBox.Show("[BG_M_USER.Delete003] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
