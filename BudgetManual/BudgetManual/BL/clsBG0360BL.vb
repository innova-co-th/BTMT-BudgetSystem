Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class clsBG0360BL

#Region "Variable"
    Private myDtResult As DataTable
    Private clsBG_M_ACCOUNT As BG_M_ACCOUNT
    Private clsBG_T_BUDGET_DATA As BG_T_BUDGET_DATA
    Private myBudgetYear As String = String.Empty
    Private myProjectNo As String = String.Empty
    Private myPeriodType As String = String.Empty
    Private myAccountNo As String = String.Empty
#End Region

#Region "Property"
#Region "DTResult"

    Public Property DtResult() As DataTable
        Get
            Return myDtResult
        End Get
        Set(ByVal value As DataTable)
            myDtResult = value
        End Set
    End Property

#End Region
#Region "AccountNo"
    Public Property AccountNo() As String
        Get
            Return myAccountNo
        End Get
        Set(ByVal value As String)
            myAccountNo = value
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
#End Region

#Region "Function"
    ''' <summary>
    ''' Get Account list
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getAccountList() As Boolean
        clsBG_M_ACCOUNT = New BG_M_ACCOUNT

        If clsBG_M_ACCOUNT.Select001 Then
            myDtResult = clsBG_M_ACCOUNT.DtResult
        Else
            myDtResult = New DataTable
        End If

        Return True
    End Function

    ''' <summary>
    ''' Get budget data to export
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getExportList() As Boolean
        clsBG_T_BUDGET_DATA = New BG_T_BUDGET_DATA

        clsBG_T_BUDGET_DATA.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_DATA.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_DATA.AccountNo = Me.AccountNo
        clsBG_T_BUDGET_DATA.ProjectNo = Me.ProjectNo

        If clsBG_T_BUDGET_DATA.Select006 Then
            myDtResult = clsBG_T_BUDGET_DATA.dtResult
        Else
            myDtResult = New DataTable
        End If

        Return True
    End Function
#End Region

End Class
