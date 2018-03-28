Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class clsBG0471BL

#Region "Variable"
    Private strBudgetYear As String = String.Empty
    Private strMonth As String = String.Empty
    Private strReportType As String = String.Empty
    Private strProjectNo As String = String.Empty
    Private strPIC As String = String.Empty
    Private dtPersonInCharge As DataTable = Nothing
    Private dsBudgetCompareData As DataSet = Nothing
    Private myBudgetStatus As Integer = 0
    Private myAuthImage As DataTable = Nothing
    Private myUserLevelId As Integer = 0I
    Private strRevNo As String = String.Empty
#End Region

#Region "Property"

    Property BudgetYear() As String
        Get
            Return strBudgetYear
        End Get
        Set(ByVal value As String)
            strBudgetYear = value
        End Set
    End Property

    Property Month() As String
        Get
            Return strMonth
        End Get
        Set(ByVal value As String)
            strMonth = value
        End Set
    End Property

    Property ReportType() As String
        Get
            Return strReportType
        End Get
        Set(ByVal value As String)
            strReportType = value
        End Set
    End Property

    Property ProjectNo() As String
        Get
            Return strProjectNo
        End Get
        Set(ByVal value As String)
            strProjectNo = value
        End Set
    End Property

    Property PIC() As String
        Get
            Return strPIC
        End Get
        Set(ByVal value As String)
            strPIC = value
        End Set
    End Property

    Property PersonInCharge() As DataTable
        Get
            Return dtPersonInCharge
        End Get
        Set(ByVal value As DataTable)
            dtPersonInCharge = value
        End Set
    End Property

    Property BudgetCompareData() As DataSet
        Get
            Return dsBudgetCompareData
        End Get
        Set(ByVal value As DataSet)
            dsBudgetCompareData = value
        End Set
    End Property

    Public Property BudgetStatus() As Integer
        Get
            Return myBudgetStatus
        End Get
        Set(ByVal value As Integer)
            myBudgetStatus = value
        End Set
    End Property

    Public Property AuthImage() As DataTable
        Get
            Return myAuthImage
        End Get
        Set(ByVal value As DataTable)
            myAuthImage = value
        End Set
    End Property

    Public Property UserLevelId() As Integer
        Get
            Return myUserLevelId
        End Get
        Set(ByVal value As Integer)
            myUserLevelId = value
        End Set
    End Property

    Property RevNo() As String
        Get
            Return strRevNo
        End Get
        Set(ByVal value As String)
            strRevNo = value
        End Set
    End Property

#End Region

#Region "Function"

    Public Function GetBudgetCompareData() As Boolean

        Dim ds As New DataSet
        Dim clsBG_T_UPLOAD_DATA As BG_T_UPLOAD_DATA = New BG_T_UPLOAD_DATA()

        clsBG_T_UPLOAD_DATA.BudgetYear = Me.BudgetYear

        If CInt(Me.Month) < 7 Then
            clsBG_T_UPLOAD_DATA.PeriodType = CStr(BGConstant.enumPeriodType.BudgetCompareVer10)
        Else
            clsBG_T_UPLOAD_DATA.PeriodType = CStr(BGConstant.enumPeriodType.BudgetCompareVer20)
        End If

        If clsBG_T_UPLOAD_DATA.Select004() = False Then
            Return False
        End If
        clsBG_T_UPLOAD_DATA.dtResult.TableName = "BUDGET_COMPARE_SUM_PIC"

        ds.Tables.Add(clsBG_T_UPLOAD_DATA.dtResult)

        If Me.GetAuthImage() = True Then
            ds.Tables.Add(Me.AuthImage)
        End If

        Me.BudgetCompareData = ds
        Return True

    End Function

    Public Function GetBudgetStatus() As Boolean

        Dim clsBG_T_BUDGET_HEADER As BG_T_BUDGET_HEADER = New BG_T_BUDGET_HEADER()

        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = CStr(enumPeriodType.OriginalBudget)
        clsBG_T_BUDGET_HEADER.BudgetType = "E"
        clsBG_T_BUDGET_HEADER.ProjectNo = "1"

        If clsBG_T_BUDGET_HEADER.Select007() = False Then
            clsBG_T_BUDGET_HEADER = Nothing
            Me.BudgetStatus = 0
            Return False
        End If

        Me.BudgetStatus = clsBG_T_BUDGET_HEADER.BudgetStatus
        clsBG_T_BUDGET_HEADER = Nothing
        Return True

    End Function

    Public Function GetAuthImage() As Boolean

        Dim clsBG_M_SETTINGS As BG_M_SETTINGS = New BG_M_SETTINGS()

        If clsBG_M_SETTINGS.Select002() = False Then
            clsBG_M_SETTINGS = Nothing
            Me.AuthImage = Nothing
            Return False
        End If

        Me.AuthImage = clsBG_M_SETTINGS.dtResult
        'Me.AuthImage.TableName = "BG_M_SETTING"
        clsBG_M_SETTINGS = Nothing
        Return True

    End Function

#End Region

End Class
