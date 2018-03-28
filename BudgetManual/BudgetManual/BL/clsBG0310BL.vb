Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class clsBG0310BL

#Region "Variable"
    Private myBudgetYear As String = String.Empty
    Private myPeriodType As String = String.Empty
    Private myUserId As String = String.Empty
    Private myProjectNo As String = String.Empty
    Private mydtPeriodList As DataTable
    Private myDtRevNoList As DataTable
    Private myOpenPeriodFlg As String = String.Empty
    Private myBudgetType As String = String.Empty
#End Region

#Region "Property"

#Region "OpenPeriod"
    Property OpenPeriodFlg() As String
        Get
            Return myOpenPeriodFlg
        End Get
        Set(ByVal value As String)
            myOpenPeriodFlg = value
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

#Region "PeriodList"
    Property PeriodList() As DataTable
        Get
            Return mydtPeriodList
        End Get
        Set(ByVal value As DataTable)
            mydtPeriodList = value
        End Set
    End Property

#End Region

#Region "RevNoList"
    Property RevNoList() As DataTable
        Get
            Return myDtRevNoList
        End Get
        Set(ByVal value As DataTable)
            myDtRevNoList = value
        End Set
    End Property

#End Region

#End Region

#Region "Function"

    Public Function CreateNewPeriod() As Boolean
        Dim clsBG_T_BUDGET_PERIOD As New BG_T_BUDGET_PERIOD

        '// Set Parameters
        clsBG_T_BUDGET_PERIOD.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_PERIOD.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_PERIOD.UserId = Me.UserId
        clsBG_T_BUDGET_PERIOD.ProjectNo = Me.ProjectNo

        '// Call Function: Select exist Period
        If clsBG_T_BUDGET_PERIOD.Select003() = False OrElse clsBG_T_BUDGET_PERIOD.dtResult.Rows.Count > 0 Then
            Return False
        End If

        '// Call Function: Insert New Period
        If clsBG_T_BUDGET_PERIOD.Insert001() = True Then

            Return True
        Else
            Return False

        End If
    End Function

    Public Function GetAllPeriodList() As Boolean
        Dim clsBG_M_PERIOD As New BG_M_PERIOD

        'clsBG_M_PERIOD.OpenFlg = Me.OpenPeriodFlg

        '// Call Function
        If clsBG_M_PERIOD.Select001() = True Then
            Me.PeriodList = clsBG_M_PERIOD.DtResult

            Return True
        Else
            Me.PeriodList = Nothing

            Return False

        End If
    End Function

    Public Function GetOpenPeriodList() As Boolean
        Dim clsBG_M_PERIOD As New BG_M_PERIOD

        clsBG_M_PERIOD.OpenFlg = Me.OpenPeriodFlg

        '// Call Function
        If clsBG_M_PERIOD.Select002() = True Then
            Me.PeriodList = clsBG_M_PERIOD.DtResult

            Return True
        Else
            Me.PeriodList = Nothing

            Return False

        End If
    End Function

    Public Function CheckReviseExist() As Boolean

        Dim clsBG_T_BUDGET_PERIOD As New BG_T_BUDGET_PERIOD

        '// Set Parameters
        clsBG_T_BUDGET_PERIOD.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_PERIOD.PeriodType = CStr(enumPeriodType.ReviseBudget)
        clsBG_T_BUDGET_PERIOD.ProjectNo = "1"

        '// Call Function: Select exist Period
        If clsBG_T_BUDGET_PERIOD.Select007() = False OrElse clsBG_T_BUDGET_PERIOD.dtResult.Rows.Count <= 0 Then
            Return False
        End If

        Return True

    End Function

    Public Function GetRevNo() As Boolean

        Dim clsBG_T_BUDGET_HEADER As BG_T_BUDGET_HEADER = New BG_T_BUDGET_HEADER()

        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo
        clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType

        If clsBG_T_BUDGET_HEADER.Select014() = False Then
            clsBG_T_BUDGET_HEADER = Nothing
            Me.RevNoList = Nothing
            Return False
        End If

        Me.RevNoList = clsBG_T_BUDGET_HEADER.dtResult
        clsBG_T_BUDGET_HEADER = Nothing
        Return True

    End Function

#End Region

End Class
