Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class clsBG0460BL

#Region "Variable"
    Private strBudgetYear As String = String.Empty
    Private strProjectNo As String = String.Empty
    Private strPeriodType As String = String.Empty
    Private bMTPChecked As Boolean = False
    Private dsBudgetData As DataSet = Nothing
    Private myBudgetStatus As Integer = 0
    Private myAuthImage As DataTable = Nothing
    Private myUserLevelId As Integer = 0I
    Private strRevNo As String = String.Empty
    Private strPIC As String = String.Empty
    Private dtPersonInCharge As DataTable = Nothing
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

    Property ProjectNo() As String
        Get
            Return strProjectNo
        End Get
        Set(ByVal value As String)
            strProjectNo = value
        End Set
    End Property

    Property PeriodType() As String
        Get
            Return strPeriodType
        End Get
        Set(ByVal value As String)
            strPeriodType = value
        End Set
    End Property

    Property MTPChecked() As Boolean
        Get
            Return bMTPChecked
        End Get
        Set(ByVal value As Boolean)
            bMTPChecked = value
        End Set
    End Property

    Property BudgetData() As DataSet
        Get
            Return dsBudgetData
        End Get
        Set(ByVal value As DataSet)
            dsBudgetData = value
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

#End Region

#Region "Function"

    Public Function GetBudgetData() As Boolean

        Dim ds As New DataSet
        Dim clsBG_T_BUDGET_DATA As BG_T_BUDGET_DATA = New BG_T_BUDGET_DATA()

        clsBG_T_BUDGET_DATA.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_DATA.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_DATA.MTPChecked = Me.MTPChecked
        clsBG_T_BUDGET_DATA.ProjectNo = Me.ProjectNo
        clsBG_T_BUDGET_DATA.UserPIC = Me.PIC

        If Me.UserLevelId = enumUserLevel.SystemAdministrator Then

            '// Admin user.
            clsBG_T_BUDGET_DATA.RevNo = Me.RevNo

            Select Case Me.PeriodType
                Case CStr(enumPeriodType.OriginalBudget)
                    If clsBG_T_BUDGET_DATA.Select012_4() = False Then
                        Return False
                    End If
                    clsBG_T_BUDGET_DATA.dtResult.TableName = "ORIGINAL_BUDGET"
                    Exit Select
                Case CStr(enumPeriodType.EstimateBudget)
                    If clsBG_T_BUDGET_DATA.Select012_5() = False Then
                        Return False
                    End If
                    clsBG_T_BUDGET_DATA.dtResult.TableName = "ESTIMATE_BUDGET"
                    Exit Select
                Case CStr(enumPeriodType.ReviseBudget)
                    If clsBG_T_BUDGET_DATA.Select012_6() = False Then
                        Return False
                    End If
                    clsBG_T_BUDGET_DATA.dtResult.TableName = "REVISE_BUDGET"
                    Exit Select
                Case Else
                    Return False
            End Select

        Else

            Select Case Me.PeriodType
                Case CStr(enumPeriodType.OriginalBudget)
                    If clsBG_T_BUDGET_DATA.Select012_1() = False Then
                        Return False
                    End If
                    clsBG_T_BUDGET_DATA.dtResult.TableName = "ORIGINAL_BUDGET"
                    Exit Select
                Case CStr(enumPeriodType.EstimateBudget)
                    If clsBG_T_BUDGET_DATA.Select012_2() = False Then
                        Return False
                    End If
                    clsBG_T_BUDGET_DATA.dtResult.TableName = "ESTIMATE_BUDGET"
                    Exit Select
                Case CStr(enumPeriodType.ReviseBudget)
                    If clsBG_T_BUDGET_DATA.Select012_3() = False Then
                        Return False
                    End If
                    clsBG_T_BUDGET_DATA.dtResult.TableName = "REVISE_BUDGET"
                    Exit Select
                Case Else
                    Return False
            End Select

        End If

        ds.Tables.Add(clsBG_T_BUDGET_DATA.dtResult)

        If Me.GetAuthImage() = True Then
            ds.Tables.Add(Me.AuthImage)
        End If

        Me.BudgetData = ds
        Return True

    End Function

    Public Function GetBudgetStatus() As Boolean

        Dim clsBG_T_BUDGET_HEADER As BG_T_BUDGET_HEADER = New BG_T_BUDGET_HEADER()

        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.BudgetType = "A"
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

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

    Public Function GetPersonInChargeList() As Boolean

        Dim clsBG_M_PERSON_IN_CHARGE As BG_M_PERSON_IN_CHARGE = New BG_M_PERSON_IN_CHARGE()

        clsBG_M_PERSON_IN_CHARGE.BudgetYear = Me.BudgetYear
        clsBG_M_PERSON_IN_CHARGE.PeriodType = Me.PeriodType

        If clsBG_M_PERSON_IN_CHARGE.Select012() = False Then
            clsBG_M_PERSON_IN_CHARGE = Nothing
            Return False
        End If

        Me.PersonInCharge = clsBG_M_PERSON_IN_CHARGE.DtResult
        clsBG_M_PERSON_IN_CHARGE = Nothing

        Return True

    End Function

End Class
