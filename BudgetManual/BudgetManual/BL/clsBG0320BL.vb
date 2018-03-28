Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class clsBG0320BL

#Region "Variable"
    Private myPeriodList As DataTable
    Private myBudgetYear As String = String.Empty
    Private myProjectNo As String = String.Empty
    Private myPeriodType As String = String.Empty
    Private myUserId As String = String.Empty
#End Region

#Region "Property"

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

#End Region

#Region "Function"

    Public Function SearchBudgetPeriod() As Boolean
        Dim clsBG_T_BUDGET_PERIOD As New BG_T_BUDGET_PERIOD

        '// Call Function
        If clsBG_T_BUDGET_PERIOD.Select001() = True AndAlso _
        clsBG_T_BUDGET_PERIOD.dtResult.Rows.Count > 0 Then
            Me.PeriodList = clsBG_T_BUDGET_PERIOD.dtResult

            Return True
        Else
            Me.PeriodList = Nothing

            Return False

        End If
    End Function

    Public Function ClosePeriod() As Boolean
        Dim clsBG_T_BUDGET_PERIOD As New BG_T_BUDGET_PERIOD

        '// Set Parameters
        clsBG_T_BUDGET_PERIOD.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_PERIOD.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_PERIOD.UserId = Me.UserId
        clsBG_T_BUDGET_PERIOD.CloseFlg = "1"
        clsBG_T_BUDGET_PERIOD.ProjectNo = Me.ProjectNo

        '// Call Function
        If clsBG_T_BUDGET_PERIOD.Update001() = True Then

            Return True
        Else
            Return False

        End If
    End Function

#End Region

End Class
