Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class clsBG0000BL

#Region "Variable"
    Private myUserId As String = String.Empty
    Private myUserName As String = String.Empty
    Private myPassword As String = String.Empty
    Private myUserLevelId As Integer
    Private myUserLevelName As String = String.Empty
    Private myUserPIC As String = String.Empty
    Private myChildPicList As DataTable
    Private myChildPIC As String = String.Empty
#End Region

#Region "Property"

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

#Region "UserLevelId"
    Property UserLevelId() As Integer
        Get
            Return myUserLevelId
        End Get
        Set(ByVal value As Integer)
            myUserLevelId = value
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

#Region "UserPIC"
    Property ChildPIC() As String
        Get
            Return myChildPIC
        End Get
        Set(ByVal value As String)
            myChildPIC = value
        End Set
    End Property
#End Region

#Region "ChildPicList"
    Property ChildPicList() As DataTable
        Get
            Return myChildPicList
        End Get
        Set(ByVal value As DataTable)
            myChildPicList = value
        End Set
    End Property
#End Region

#End Region

#Region "Function"

    Public Function CheckLogin() As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        '// Set Function's Parameter
        clsBG_M_USER.UserId = Me.UserId

        '// Call Function
        If clsBG_M_USER.Select001() = False OrElse clsBG_M_USER.dtResult.Rows.Count = 0 Then

            Return False
        Else
            Dim dr As DataRow = clsBG_M_USER.dtResult.Rows(0)

            If CStr(dr("PASSWORD")) = Me.Password And CInt(dr("EXPIRE_FLAG")) = 0 Then
                Me.UserId = CStr(dr("USER_ID"))
                Me.UserName = CStr(dr("USER_NAME"))
                Me.UserLevelId = CShort(dr("USER_LEVEL_ID"))
                Me.UserLevelName = CStr(dr("USER_LEVEL_NAME"))
                Me.UserPIC = CStr(dr("PERSON_IN_CHARGE_NO"))

                Return True
            Else

                Return False
            End If
        End If
    End Function

    Public Function CheckUserLoggedIn() As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        '// Set Function's Parameter
        clsBG_M_USER.UserId = p_strUserId

        '// Call Function
        If clsBG_M_USER.Select014() = False OrElse clsBG_M_USER.dtResult.Rows.Count = 0 Then

            Return False
        Else

            Return True
        End If
    End Function

    Public Function CheckLockPIC() As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        '// Set Function's Parameter
        clsBG_M_USER.UserPIC = p_strUserPIC

        '// Call Function
        If clsBG_M_USER.Select011() = False OrElse clsBG_M_USER.dtResult.Rows.Count = 0 Then

            Return False
        Else
            Me.UserName = CStr(clsBG_M_USER.dtResult.Rows(0).Item("LOGIN_USER_NAME"))

            Return True
        End If
    End Function

    Public Function CheckLockChildPIC() As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        '// Set Function's Parameter
        clsBG_M_USER.UserPIC = Me.ChildPIC

        '// Call Function
        If clsBG_M_USER.Select011() = False OrElse clsBG_M_USER.dtResult.Rows.Count = 0 Then

            Return False
        Else
            Me.UserName = CStr(clsBG_M_USER.dtResult.Rows(0).Item("LOGIN_USER_NAME"))

            Return True
        End If
    End Function

    Public Function AddLockPIC() As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        '// Set Function's Parameter
        clsBG_M_USER.UserPIC = p_strUserPIC
        clsBG_M_USER.UserId = p_strUserId

        '// Call Function
        If clsBG_M_USER.Insert002() = True Then

            Return True
        Else
            Return False

        End If
    End Function

    Public Function GetChildPicList() As Boolean
        Dim clsBG_M_CHILD_PIC As New BG_M_CHILD_PIC

        clsBG_M_CHILD_PIC.ParentNo = Me.UserPIC

        If clsBG_M_CHILD_PIC.Select003 = True And clsBG_M_CHILD_PIC.DtResult.Rows.Count > 0 Then
            Me.ChildPicList = clsBG_M_CHILD_PIC.DtResult

            Return True
        Else
            Return False

        End If
    End Function

#End Region

End Class

