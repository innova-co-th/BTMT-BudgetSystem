Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.SqlClient

Public Class clsBG0390BL

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

#Region "DtResult"
    Public Property DTResult() As DataTable
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

    Public Function SearchTransLog() As Boolean
        Dim clsBG_T_TRANS_LOG As New BG_T_TRANS_LOG

        clsBG_T_TRANS_LOG.BudgetYear = Me.BudgetYear
        clsBG_T_TRANS_LOG.PeriodType = Me.PeriodType
        clsBG_T_TRANS_LOG.ProjectNo = Me.ProjectNo

        If clsBG_T_TRANS_LOG.Select001() = True Then
            Me.DTResult = clsBG_T_TRANS_LOG.dtResult

            Return True
        Else
            Me.DTResult = Nothing

            Return False
        End If
    End Function

    Public Function SearchAdminLog() As Boolean
        Dim clsBG_T_TRANS_LOG As New BG_T_TRANS_LOG

        clsBG_T_TRANS_LOG.FromDate = Me.FromDate
        clsBG_T_TRANS_LOG.ToDate = Me.ToDate

        If clsBG_T_TRANS_LOG.Select002() = True Then
            Me.DTResult = clsBG_T_TRANS_LOG.dtResult

            Return True
        Else
            Me.DTResult = Nothing

            Return False
        End If
    End Function

#End Region

End Class
