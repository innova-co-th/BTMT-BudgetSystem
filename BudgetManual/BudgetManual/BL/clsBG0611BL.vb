Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.SqlClient

Public Class clsBG0611BL

#Region "Variable"
    Private myUserLevelList As DataTable
    Private myUserLevelId As DataTable
    Private myUserLevel As String = String.Empty
    Private myUserPermissions() As Boolean
    Private myUserId As String = String.Empty
    Private myUserLevelName As String = String.Empty
    Private myUserLevelFilter As String = String.Empty
    Private myUserPermissionsFilter As String = String.Empty

#End Region

#Region "Property"

#Region "UserLevelList"
    Property UserLevelList() As DataTable
        Get
            Return myUserLevelList
        End Get
        Set(ByVal value As DataTable)
            myUserLevelList = value
        End Set
    End Property

#End Region

#Region "UserLevelId"
    Property UserLevelId() As DataTable
        Get
            Return myUserLevelId
        End Get
        Set(ByVal value As DataTable)
            myUserLevelId = value
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

#Region "UserPermissionsFilter"
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

    Public Function UpdateUserLevel() As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        '// Set Parameters
        clsBG_M_USER.UserLevel = Me.UserLevel
        clsBG_M_USER.UserPermissions = Me.UserPermissions
        clsBG_M_USER.UserId = Me.UserId
        clsBG_M_USER.UserLevelName = Me.UserLevelName

        '// Call Function
        If clsBG_M_USER.Update003() = True Then

            Return True
        Else
            Return False

        End If
    End Function

    Public Function AddUserLevel() As Boolean

        Dim clsBG_M_USER As New BG_M_USER

        '// Set Parameters
        clsBG_M_USER.UserLevel = Me.UserLevel
        clsBG_M_USER.UserPermissions = Me.UserPermissions
        clsBG_M_USER.UserId = Me.UserId
        clsBG_M_USER.UserLevelName = Me.UserLevelName

        '// Call Function
        If clsBG_M_USER.Insert003() = True Then

            Return True
        Else
            Return False

        End If

    End Function

    Public Function DeleteUserLevel() As Boolean

        Dim clsBG_M_USER As New BG_M_USER

        '// Set Parameters
        clsBG_M_USER.UserLevel = Me.UserLevel

        '// Call Function
        If clsBG_M_USER.Delete003() = True Then

            Return True
        Else
            Return False

        End If

    End Function

    Public Function GetUserLevelList() As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        clsBG_M_USER.UserLevel = Me.UserLevelFilter
        clsBG_M_USER.UserPermissionsFilter = Me.UserPermissionsFilter

        '// Call Function
        If clsBG_M_USER.Select013() = True Then
            Me.UserLevelList = clsBG_M_USER.dtResult

            Return True
        Else
            Me.UserLevelList = Nothing

            Return False

        End If
    End Function

    Public Function GetLastUserLevelId() As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        '// Call Function
        If clsBG_M_USER.Select015() = True Then
            Me.UserLevelId = clsBG_M_USER.dtResult

            Return True
        Else
            Me.UserLevelId = Nothing

            Return False

        End If
    End Function

#End Region

End Class
