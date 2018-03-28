Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.SqlClient

Public Class clsBG0610BL

#Region "Variable"
    Private myUserList As DataTable
    Private myUserId As String = String.Empty
    Private myUserName As String = String.Empty
    Private myUserLevel As String = String.Empty
    Private myPassword As String = String.Empty
    Private myUserPIC As String = String.Empty
    Private myEmail As String = String.Empty
    Private myExpireFlg As String = String.Empty
    Private myUserId2 As String = String.Empty
    Private myUserIdFilter As String = String.Empty
    Private myUserNameFilter As String = String.Empty
    Private myUserLevelFilter As String = String.Empty
    Private myUserPICFilter As String = String.Empty
#End Region

#Region "Property"

#Region "UserList"
    Property UserList() As DataTable
        Get
            Return myUserList
        End Get
        Set(ByVal value As DataTable)
            myUserList = value
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

#Region "UserIdFilter"
    Property UserIdFilter() As String
        Get
            Return myUserIdFilter
        End Get
        Set(ByVal value As String)
            myUserIdFilter = value
        End Set
    End Property
#End Region

#Region "UserNameFilter"
    Property UserNameFilter() As String
        Get
            Return myUserNameFilter
        End Get
        Set(ByVal value As String)
            myUserNameFilter = value
        End Set
    End Property
#End Region

#Region "UserLevelFilter"
    Property UserLevelFilter() As String
        Get
            Return myUserLevelFilter
        End Get
        Set(ByVal value As String)
            myUserLevelFilter = value
        End Set
    End Property
#End Region

#Region "UserPICFilter"
    Property UserPICFilter() As String
        Get
            Return myUserPICFilter
        End Get
        Set(ByVal value As String)
            myUserPICFilter = value
        End Set
    End Property
#End Region

#End Region

#Region "Function"

    Public Function CreateNewUser() As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        '// Set Parameters
        clsBG_M_USER.UserId = Me.UserId
        clsBG_M_USER.UserLevel = Me.UserLevel
        clsBG_M_USER.UserName = Me.UserName
        clsBG_M_USER.Password = Me.Password
        clsBG_M_USER.UserPIC = Me.myUserPIC
        clsBG_M_USER.Email = Me.Email
        clsBG_M_USER.ExpireFlg = Me.ExpireFlg
        clsBG_M_USER.UserId2 = Me.UserId2

        '// Call Function
        If clsBG_M_USER.Insert001() = True Then

            Return True
        Else
            Return False

        End If
    End Function
    Public Function CreateNewUserImport(Optional ByVal pConn As SqlConnection = Nothing, _
                              Optional ByVal pTrans As SqlTransaction = Nothing) As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        '// Set Parameters
        clsBG_M_USER.UserId = Me.UserId
        clsBG_M_USER.UserLevel = Me.UserLevel
        clsBG_M_USER.UserName = Me.UserName
        clsBG_M_USER.Password = Me.Password
        clsBG_M_USER.UserPIC = Me.myUserPIC
        clsBG_M_USER.Email = Me.Email
        clsBG_M_USER.ExpireFlg = Me.ExpireFlg
        clsBG_M_USER.UserId2 = Me.UserId2

        '// Call Function
        If clsBG_M_USER.Insert001(pConn, pTrans) = True Then

            Return True
        Else
            Return False

        End If
    End Function

    Public Function UpdateUserData() As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        '// Set Parameters
        clsBG_M_USER.UserId = Me.UserId
        clsBG_M_USER.UserLevel = Me.UserLevel
        clsBG_M_USER.UserName = Me.UserName
        clsBG_M_USER.Password = Me.Password
        clsBG_M_USER.UserPIC = Me.myUserPIC
        clsBG_M_USER.Email = Me.Email
        clsBG_M_USER.ExpireFlg = Me.ExpireFlg
        clsBG_M_USER.UserId2 = Me.UserId2

        '// Call Function
        If clsBG_M_USER.Update002() = True Then

            Return True
        Else
            Return False

        End If
    End Function
    Public Function UpdateUserDataImport(Optional ByVal pConn As SqlConnection = Nothing, _
                              Optional ByVal pTrans As SqlTransaction = Nothing) As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        '// Set Parameters
        clsBG_M_USER.UserId = Me.UserId
        clsBG_M_USER.UserLevel = Me.UserLevel
        clsBG_M_USER.UserName = Me.UserName
        clsBG_M_USER.Password = Me.Password
        clsBG_M_USER.UserPIC = Me.myUserPIC
        clsBG_M_USER.Email = Me.Email
        clsBG_M_USER.ExpireFlg = Me.ExpireFlg
        clsBG_M_USER.UserId2 = Me.UserId2

        '// Call Function
        If clsBG_M_USER.Update002(pConn, pTrans) = True Then

            Return True
        Else
            Return False

        End If
    End Function

    Public Function SearchAllUser() As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        clsBG_M_USER.UserId = Me.UserIdFilter
        clsBG_M_USER.UserName = Me.UserNameFilter
        clsBG_M_USER.UserLevel = Me.UserLevelFilter
        clsBG_M_USER.UserPIC = Me.UserPICFilter

        '// Call Function
        If clsBG_M_USER.Select002() = True Then
            Me.UserList = clsBG_M_USER.dtResult

            Return True
        Else
            Me.UserList = Nothing

            Return False

        End If
    End Function

    Public Function CheckUserExist() As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        '// Set Parameter
        clsBG_M_USER.UserId = Me.UserId

        '// Call Function
        If clsBG_M_USER.Select001() = True AndAlso _
        clsBG_M_USER.dtResult.Rows.Count > 0 Then

            Return True
        Else
            Return False

        End If
    End Function

    Public Function SearchAllUserPIC() As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        '// Call Function
        If clsBG_M_USER.Select003() = True Then
            Me.UserList = clsBG_M_USER.dtResult

            Return True
        Else
            Me.UserList = Nothing

            Return False

        End If
    End Function

    Public Function SearchAllUserLevel() As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        '// Call Function
        If clsBG_M_USER.Select004() = True Then
            Me.UserList = clsBG_M_USER.dtResult

            Return True
        Else
            Me.UserList = Nothing

            Return False

        End If
    End Function

    Public Function SearchUserExcel() As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        '// Call Function
        If clsBG_M_USER.Select005() = True Then
            Me.UserList = clsBG_M_USER.dtResult

            Return True
        Else
            Me.UserList = Nothing

            Return False

        End If
    End Function
#End Region

End Class
