Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class clsBG0010BL

#Region "Variable"
    Private myInfoList As DataTable = Nothing
    Private myPeriodList As DataTable = Nothing
    Private myOrderList As DataTable = Nothing
    Private myUserInfo As DataTable = Nothing
    Private myBudgetYear As String = String.Empty
    Private myPeriodType As String = String.Empty
    Private myBudgetType As String = String.Empty
    Private myUserPIC As String = String.Empty
    Private myStatus As String = String.Empty
    Private myRevNo As String = String.Empty
    Private myProjectNo As String = String.Empty
    Private myUserPicList As DataTable = Nothing
    Private myViewAll As Boolean = False
#End Region

#Region "Property"

#Region "InfoList"
    Property InfoList() As DataTable
        Get
            Return myInfoList
        End Get
        Set(ByVal value As DataTable)
            myInfoList = value
        End Set
    End Property

#End Region

#Region "PeriodList"
    Property PeriodList() As DataTable
        Get
            Return myPeriodList
        End Get
        Set(ByVal value As DataTable)
            myPeriodList = value
        End Set
    End Property
#End Region

#Region "OrderList"
    Property OrderList() As DataTable
        Get
            Return myOrderList
        End Get
        Set(ByVal value As DataTable)
            myOrderList = value
        End Set
    End Property
#End Region

#Region "UserInfo"
    Property UserInfo() As DataTable
        Get
            Return myUserInfo
        End Get
        Set(ByVal value As DataTable)
            myUserInfo = value
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

#Region "Status"
    Property Status() As String
        Get
            Return myStatus
        End Get
        Set(ByVal value As String)
            myStatus = value
        End Set
    End Property
#End Region

#Region "Rev No"
    Property RevNo() As String
        Get
            Return myRevNo
        End Get
        Set(ByVal value As String)
            myRevNo = value
        End Set
    End Property
#End Region

#Region "UserPicList"
    Property UserPicList() As DataTable
        Get
            Return myUserPicList
        End Get
        Set(ByVal value As DataTable)
            myUserPicList = value
        End Set
    End Property
#End Region

#Region "ViewAll"
    Property ViewAll() As Boolean
        Get
            Return myViewAll
        End Get
        Set(ByVal value As Boolean)
            myViewAll = value
        End Set
    End Property
#End Region

#End Region

#Region "Function"
    Public Function GetOptions() As Boolean
        Dim clsBG_M_SETTINGS As New BG_M_SETTINGS

        '// Call Function
        If clsBG_M_SETTINGS.Select001() = True Then
            p_blnSendAutoMail = clsBG_M_SETTINGS.EnableAutoMail

            If p_blnSendAutoMail Then
                '// Initial Send Mail Module
                BGSendMail.SmtpServer = clsBG_M_SETTINGS.SmtpServer
                If IsNumeric(clsBG_M_SETTINGS.SmtpPort) Then
                    BGSendMail.SmtpPort = CInt(clsBG_M_SETTINGS.SmtpPort)
                Else
                    BGSendMail.SmtpPort = 25
                End If
                If clsBG_M_SETTINGS.UseAuthentication Then
                    BGSendMail.UseAuthentication = True
                    BGSendMail.SmtpUser = clsBG_M_SETTINGS.SmtpUser
                    BGSendMail.SmtpPassword = clsBG_M_SETTINGS.SmtpPassword
                Else
                    BGSendMail.UseAuthentication = False
                End If
                BGSendMail.FromAddress = clsBG_M_SETTINGS.FromAddr
            End If

            Return True
        Else
            p_blnSendAutoMail = False

            Return False
        End If
    End Function

    Public Function CheckBudgetDataStatus(Optional ByVal blnIgnoreAuthLevel As Boolean = False) As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER
        Dim blnRes As Boolean

        '// Set Parameters
        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_HEADER.UserPIC = Me.UserPIC
        clsBG_T_BUDGET_HEADER.RevNo = Me.RevNo
        clsBG_T_BUDGET_HEADER.Status = Me.Status
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

        '// Get Budget data

        If blnIgnoreAuthLevel = False Then
            If Me.UserPIC = "0000" Or Me.UserPIC = "BTMT10" Or Me.UserPIC = "BTMT3" Then
                blnRes = clsBG_T_BUDGET_HEADER.Select003() '// All budget data
            Else
                blnRes = clsBG_T_BUDGET_HEADER.Select004() '// Select by PIC
            End If

        Else
            If Me.UserPIC = "0000" Then
                blnRes = clsBG_T_BUDGET_HEADER.Select003() '// All budget data
            Else
                blnRes = clsBG_T_BUDGET_HEADER.Select004() '// Select by PIC
            End If
        End If

       

        If blnRes = True AndAlso clsBG_T_BUDGET_HEADER.dtResult.Rows.Count > 0 Then

            Return True
        Else
            Return False

        End If
    End Function

    Public Function CheckBudgetDataStatusReInputByOrder(Optional ByVal blnIgnoreAuthLevel As Boolean = False) As Boolean
        Dim clsBG_T_BUDGET_DATA_REINPUT As New BG_T_BUDGET_DATA_REINPUT
        Dim blnRes As Boolean

        '// Set Parameters
        clsBG_T_BUDGET_DATA_REINPUT.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_DATA_REINPUT.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_DATA_REINPUT.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_DATA_REINPUT.UserPIC = Me.UserPIC
        clsBG_T_BUDGET_DATA_REINPUT.RevNo = Me.RevNo
        clsBG_T_BUDGET_DATA_REINPUT.ItemStatus = Me.Status
        clsBG_T_BUDGET_DATA_REINPUT.ProjectNo = Me.ProjectNo

        '// Get Budget data

        If blnIgnoreAuthLevel = False Then
            If Me.UserPIC = "0000" Or Me.UserPIC = "BTMT10" Or Me.UserPIC = "BTMT3" Then
                blnRes = clsBG_T_BUDGET_DATA_REINPUT.Select002() '// All budget data
            Else
                blnRes = clsBG_T_BUDGET_DATA_REINPUT.Select003() '// Select by PIC
            End If

        Else
            If Me.UserPIC = "0000" Then
                blnRes = clsBG_T_BUDGET_DATA_REINPUT.Select002() '// All budget data
            Else
                blnRes = clsBG_T_BUDGET_DATA_REINPUT.Select003() '// Select by PIC
            End If
        End If

        If blnRes = True AndAlso clsBG_T_BUDGET_DATA_REINPUT.dtResult.Rows.Count > 0 Then

            Return True
        Else
            Return False

        End If
    End Function

    Public Function CheckBudgetDataExistView() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER
        Dim blnRes As Boolean

        '// Set Parameters
        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_HEADER.UserPIC = Me.UserPIC
        clsBG_T_BUDGET_HEADER.RevNo = Me.RevNo
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

        '// Get Budget data
        If Me.UserPIC = "0000" Or Me.UserPIC = "BTMT10" Or Me.UserPIC = "BTMT3" Then
            blnRes = clsBG_T_BUDGET_HEADER.Select001() '// All budget data
        Else
            blnRes = clsBG_T_BUDGET_HEADER.Select010() '// Select by PIC
        End If

        If blnRes = True AndAlso clsBG_T_BUDGET_HEADER.dtResult.Rows.Count > 0 Then

            Return True
        Else
            Return False

        End If
    End Function

    Public Function CheckBudgetDataExist() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER
        Dim blnRes As Boolean

        '// Set Parameters
        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_HEADER.UserPIC = Me.UserPIC
        clsBG_T_BUDGET_HEADER.RevNo = Me.RevNo
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

        '// Get Budget data
        If Me.UserPIC = "0000" Then
            blnRes = clsBG_T_BUDGET_HEADER.Select001() '// All budget data
        Else
            blnRes = clsBG_T_BUDGET_HEADER.Select002() '// Select by PIC
        End If

        If blnRes = True AndAlso clsBG_T_BUDGET_HEADER.dtResult.Rows.Count > 0 Then

            Return True
        Else
            Return False

        End If
    End Function

    Public Function CheckBudgetOrderMatch() As Boolean
        Dim clsBG_M_BUDGET_ORDER As New BG_M_BUDGET_ORDER
        Dim blnRes As Boolean

        '// Set Parameters
        clsBG_M_BUDGET_ORDER.BudgetType = Me.BudgetType
        clsBG_M_BUDGET_ORDER.UserPIC = Me.UserPIC

        '// Get Budget data
        If Me.UserPIC = "0000" Then
            blnRes = clsBG_M_BUDGET_ORDER.Select005() '// All budget data
        Else
            blnRes = clsBG_M_BUDGET_ORDER.Select006() '// Select by PIC
        End If

        If blnRes = True AndAlso clsBG_M_BUDGET_ORDER.dtResult.Rows.Count > 0 Then
            Me.UserPicList = clsBG_M_BUDGET_ORDER.dtResult

            Return True
        Else
            Me.UserPicList = New DataTable

            Return False
        End If
    End Function

    Public Function SearchOpenPeriod() As Boolean
        Dim clsBG_T_BUDGET_PERIOD As New BG_T_BUDGET_PERIOD

        clsBG_T_BUDGET_PERIOD.ViewAll = Me.ViewAll

        '// Call Function
        If clsBG_T_BUDGET_PERIOD.Select006() = False Then
            Me.PeriodList = Nothing

            Return False
        Else
            Me.PeriodList = clsBG_T_BUDGET_PERIOD.dtResult

            Return True
        End If
    End Function

    Public Function SearchViewPeriod() As Boolean
        Dim clsBG_T_BUDGET_PERIOD As New BG_T_BUDGET_PERIOD

        '// Set Parameter
        clsBG_T_BUDGET_PERIOD.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_PERIOD.ViewAll = Me.ViewAll

        '// Call Function
        If clsBG_T_BUDGET_PERIOD.Select004() = False Then
            Me.PeriodList = Nothing

            Return False
        Else
            Me.PeriodList = clsBG_T_BUDGET_PERIOD.dtResult

            Return True
        End If
    End Function

    Public Function SearchAllPeriod() As Boolean
        Dim clsBG_T_BUDGET_PERIOD As New BG_T_BUDGET_PERIOD

        '// Set Parameter
        clsBG_T_BUDGET_PERIOD.ViewAll = Me.ViewAll

        '// Call Function
        If clsBG_T_BUDGET_PERIOD.Select005() = False Then
            Me.PeriodList = Nothing

            Return False
        Else
            Me.PeriodList = clsBG_T_BUDGET_PERIOD.dtResult

            Return True
        End If
    End Function

    Public Function SearchInformation() As Boolean
        Dim clsBG_T_INFORMATION As New BG_T_INFORMATION

        '// Call Function
        If clsBG_T_INFORMATION.Select001() = False Then
            Me.InfoList = Nothing

            Return False
        Else
            Me.InfoList = clsBG_T_INFORMATION.DTResult

            Return True
        End If
    End Function

    Public Function GetUserInfo() As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        '// Set Parameter
        clsBG_M_USER.UserId = p_strUserId

        '// Call Function
        If clsBG_M_USER.Select001() = False OrElse clsBG_M_USER.dtResult.Rows.Count = 0 Then
            Me.UserInfo = Nothing

            Return False
        Else
            Me.UserInfo = clsBG_M_USER.dtResult

            Return True
        End If
    End Function

    Public Function HavePermission(ByVal pPermission As enumPermissionCd) As Boolean
        '// Get User info
        If Me.UserInfo Is Nothing OrElse Me.UserInfo.Rows.Count = 0 Then
            If GetUserInfo() = False Then
                Return False
            End If
        End If

        Dim dr As DataRow = Me.UserInfo.Rows(0)

        '// Check User's Permission
        Select Case pPermission
            Case enumPermissionCd.Entry
                If CStr(dr("ENTRY")) = "Y" Then
                    Return True
                Else
                    Return False
                End If

            Case enumPermissionCd.Submit
                If CStr(dr("SUBMIT")) = "Y" Then
                    Return True
                Else
                    Return False
                End If

            Case enumPermissionCd.Approve
                If CStr(dr("APPROVE")) = "Y" Then
                    Return True
                Else
                    Return False
                End If

            Case enumPermissionCd.Adjust
                If CStr(dr("ADJUST")) = "Y" Then
                    Return True
                Else
                    Return False
                End If

            Case enumPermissionCd.Auth1
                If CStr(dr("AUTH1")) = "Y" Then
                    Return True
                Else
                    Return False
                End If

            Case enumPermissionCd.Auth2
                If CStr(dr("AUTH2")) = "Y" Then
                    Return True
                Else
                    Return False
                End If

            Case enumPermissionCd.Import
                If CStr(dr("IMPORT")) = "Y" Then
                    Return True
                Else
                    Return False
                End If

            Case enumPermissionCd.Export
                If CStr(dr("EXPORT")) = "Y" Then
                    Return True
                Else
                    Return False
                End If

            Case enumPermissionCd.Master
                If CStr(dr("MASTER")) = "Y" Then
                    Return True
                Else
                    Return False
                End If

            Case enumPermissionCd.System
                If CStr(dr("SYSTEM")) = "Y" Then
                    Return True
                Else
                    Return False
                End If

            Case enumPermissionCd.View
                If CStr(dr("VIEW")) = "Y" Then
                    Return True
                Else
                    Return False
                End If
            Case enumPermissionCd.DirectInput
                If CStr(dr("DIRECT_INPUT")) = "Y" Then
                    Return True
                Else
                    Return False
                End If

            Case Else
                Return False

        End Select
    End Function

    Public Function GetMaxRevNo() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER

        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

        If clsBG_T_BUDGET_HEADER.Select006() = True Then
            Me.RevNo = clsBG_T_BUDGET_HEADER.RevNo

            Return True
        Else
            Me.RevNo = "1"

            Return False

        End If
    End Function

    Public Function ClearLockPIC() As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        '// Set Function's Parameter
        clsBG_M_USER.UserPIC = p_strUserPIC
        clsBG_M_USER.UserId = p_strUserId

        '// Call Function
        If clsBG_M_USER.Delete001() = True Then

            Return True
        Else
            Return False

        End If
    End Function

#End Region

End Class
