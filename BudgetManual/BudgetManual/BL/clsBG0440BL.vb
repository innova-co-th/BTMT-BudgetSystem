Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class clsBG0440BL

#Region "Variable"
    Private myAccountCodeData As DataSet = Nothing
    Private myBudgetYear As String = String.Empty
    Private myProjectNo As String = String.Empty
    Private myPrevProjectNo As String = String.Empty
    Private myPeriodType As String = String.Empty
    Private myBudgetStatus As Integer = 0
    Private myBudgetAdjustTable As DataTable = Nothing
    Private myAuthImage As DataTable = Nothing
    Private myInvestments As DataTable = Nothing
    Private myMTPBudget As Boolean = False
    Private myUserLevelId As Integer = 0I
    Private strRevNo As String = String.Empty
    Private strPrevRevNo As String = String.Empty
#End Region

#Region "Property"
    Public Property BudgetAdjustTable() As DataTable
        Get
            Return myBudgetAdjustTable
        End Get
        Set(ByVal value As DataTable)
            myBudgetAdjustTable = value
        End Set
    End Property
    Public Property AccountCodeData() As DataSet
        Get
            Return myAccountCodeData
        End Get
        Set(ByVal value As DataSet)
            myAccountCodeData = value
        End Set
    End Property
    Public Property BudgetYear() As String
        Get
            Return myBudgetYear
        End Get
        Set(ByVal value As String)
            myBudgetYear = value
        End Set
    End Property

    Public Property ProjectNo() As String
        Get
            Return myProjectNo
        End Get
        Set(ByVal value As String)
            myProjectNo = value
        End Set
    End Property

    Public Property PrevProjectNo() As String
        Get
            Return myPrevProjectNo
        End Get
        Set(ByVal value As String)
            myPrevProjectNo = value
        End Set
    End Property

    Public Property PeriodType() As String
        Get
            Return myPeriodType
        End Get
        Set(ByVal value As String)
            myPeriodType = value
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

    Public Property Investments() As DataTable
        Get
            Return myInvestments
        End Get
        Set(ByVal value As DataTable)
            myInvestments = value
        End Set
    End Property

    Public Property MTPBudget() As Boolean
        Get
            Return myMTPBudget
        End Get
        Set(ByVal value As Boolean)
            myMTPBudget = value
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

    Property PrevRevNo() As String
        Get
            Return strPrevRevNo
        End Get
        Set(ByVal value As String)
            strPrevRevNo = value
        End Set
    End Property
#End Region

#Region "Function"
    Public Function getAccountData() As Boolean
        Dim ds As New DataSet()
        Dim clsBG_T_BUDGET_DATA As BG_T_BUDGET_DATA = New BG_T_BUDGET_DATA()
        Dim clsBG_T_BUDGET_REFERENCE As BG_T_BUDGET_REFERENCE = New BG_T_BUDGET_REFERENCE

        clsBG_T_BUDGET_DATA.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_DATA.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_DATA.ProjectNo = Me.ProjectNo

        '// Get Reference Budget
        clsBG_T_BUDGET_DATA.RefBudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_DATA.RefPeriodType = "1"
        clsBG_T_BUDGET_DATA.RefProjectNo = "1"
        clsBG_T_BUDGET_DATA.RefRevNo = "1"
        clsBG_T_BUDGET_DATA.RefEstProjectNo = "1"
        clsBG_T_BUDGET_DATA.RefEstRevNo = "1"
        clsBG_T_BUDGET_DATA.RefRBProjectNo = "1"
        clsBG_T_BUDGET_DATA.RefRBRevNo = "1"
        clsBG_T_BUDGET_DATA.MtpProjectNo = "1"
        clsBG_T_BUDGET_DATA.MtpRevNo = "1"

        If Me.PeriodType = CStr(enumPeriodType.OriginalBudget) Then

            '// Ref. Estimate 
            clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
            clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
            clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
            clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
            clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(enumPeriodType.EstimateBudget)

            If Me.RevNo = "" Then
                If clsBG_T_BUDGET_REFERENCE.Select002 = False Then

                    Return False

                End If

            Else
                If clsBG_T_BUDGET_REFERENCE.Select001 = False Then

                    Return False

                End If

            End If

            If clsBG_T_BUDGET_REFERENCE.dtResult.Rows.Count > 0 Then

                clsBG_T_BUDGET_DATA.RefEstProjectNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_PROJECT_NO").ToString
                clsBG_T_BUDGET_DATA.RefEstRevNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_REV_NO").ToString

            End If

            '// Ref. MTP
            clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
            clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
            clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
            clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
            clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(enumPeriodType.MBPBudget)

            If Me.RevNo = "" Then
                If clsBG_T_BUDGET_REFERENCE.Select002 = False Then

                    Return False

                End If

            Else
                If clsBG_T_BUDGET_REFERENCE.Select001 = False Then

                    Return False

                End If

            End If

            If clsBG_T_BUDGET_REFERENCE.dtResult.Rows.Count > 0 Then

                clsBG_T_BUDGET_DATA.MtpProjectNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_PROJECT_NO").ToString
                clsBG_T_BUDGET_DATA.MtpRevNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_REV_NO").ToString

            Else
                clsBG_T_BUDGET_DATA.RefBudgetYear = CStr(CInt(Me.BudgetYear) - 2)
                clsBG_T_BUDGET_DATA.RefBudgetYear = CStr(enumPeriodType.MBPBudget)
            End If

        ElseIf Me.PeriodType = CStr(enumPeriodType.MBPBudget) Then

            '// Ref. Forecast  
            clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
            clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
            clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
            clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
            clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(enumPeriodType.OriginalBudget)


            If Me.RevNo = "" Then
                If clsBG_T_BUDGET_REFERENCE.Select002 = False Then

                    Return False

                End If

            Else
                If clsBG_T_BUDGET_REFERENCE.Select001 = False Then

                    Return False

                End If

            End If

            If clsBG_T_BUDGET_REFERENCE.dtResult.Rows.Count > 0 Then

                clsBG_T_BUDGET_DATA.RefRBProjectNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_PROJECT_NO").ToString
                clsBG_T_BUDGET_DATA.RefRBRevNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_REV_NO").ToString

            End If

            '// Ref. MTP Previous
            clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
            clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
            clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
            clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
            clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(enumPeriodType.MBPBudget)

            If Me.RevNo = "" Then
                If clsBG_T_BUDGET_REFERENCE.Select002 = False Then

                    Return False

                End If

            Else
                If clsBG_T_BUDGET_REFERENCE.Select001 = False Then

                    Return False

                End If

            End If

            If clsBG_T_BUDGET_REFERENCE.dtResult.Rows.Count > 0 Then

                clsBG_T_BUDGET_DATA.MtpProjectNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_PROJECT_NO").ToString
                clsBG_T_BUDGET_DATA.MtpRevNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_REV_NO").ToString

            End If

        End If

        If Me.UserLevelId = enumUserLevel.SystemAdministrator Then

            '// Admin user.
            clsBG_T_BUDGET_DATA.RevNo = Me.RevNo

            Select Case CType(Me.PeriodType, enumPeriodType)
                Case enumPeriodType.OriginalBudget
                    'clsBG_T_BUDGET_DATA.MtpProjectNo = Me.PrevProjectNo
                    'clsBG_T_BUDGET_DATA.MtpRevNo = Me.PrevRevNo
                    clsBG_T_BUDGET_DATA.TableName = "OriginalSummaryByAccountCode"
                    If clsBG_T_BUDGET_DATA.Select015_4() = False Then
                        Return False
                    End If

                Case enumPeriodType.EstimateBudget
                    clsBG_T_BUDGET_DATA.RefPeriodType = CStr(enumPeriodType.OriginalBudget)
                    clsBG_T_BUDGET_DATA.TableName = "EstimateSummaryByAccountCode"
                    If clsBG_T_BUDGET_DATA.Select015_5() = False Then
                        Return False
                    End If

                Case enumPeriodType.ForecastBudget
                    clsBG_T_BUDGET_DATA.TableName = "ReviseSummaryByAccountCode"
                    'clsBG_T_BUDGET_DATA.MTPBudget = Me.MTPBudget
                    If clsBG_T_BUDGET_DATA.Select015_6() = False Then
                        Return False
                    End If

                Case enumPeriodType.MBPBudget
                    clsBG_T_BUDGET_DATA.TableName = "MTPSummaryByAccountCode"
                    clsBG_T_BUDGET_DATA.Status = CStr(enumBudgetStatus.Approve)
                    clsBG_T_BUDGET_DATA.BudgetType = "E"
                    'clsBG_T_BUDGET_DATA.PrevProjectNo = Me.PrevProjectNo
                    'clsBG_T_BUDGET_DATA.PrevMTPRevNo = Me.PrevRevNo
                    If clsBG_T_BUDGET_DATA.Select029_2() = False Then
                        Return False
                    End If

            End Select

        Else

            Select Case CType(Me.PeriodType, enumPeriodType)
                Case enumPeriodType.OriginalBudget
                    'clsBG_T_BUDGET_DATA.MtpProjectNo = Me.PrevProjectNo
                    clsBG_T_BUDGET_DATA.TableName = "OriginalSummaryByAccountCode"
                    If clsBG_T_BUDGET_DATA.Select015_1() = False Then
                        Return False
                    End If

                Case enumPeriodType.EstimateBudget
                    clsBG_T_BUDGET_DATA.RefPeriodType = CStr(enumPeriodType.OriginalBudget)
                    clsBG_T_BUDGET_DATA.TableName = "EstimateSummaryByAccountCode"
                    If clsBG_T_BUDGET_DATA.Select015_2() = False Then
                        Return False
                    End If

                Case enumPeriodType.ForecastBudget
                    clsBG_T_BUDGET_DATA.TableName = "ReviseSummaryByAccountCode"
                    'clsBG_T_BUDGET_DATA.MTPBudget = Me.MTPBudget
                    If clsBG_T_BUDGET_DATA.Select015_3() = False Then
                        Return False
                    End If

                Case enumPeriodType.MBPBudget
                    clsBG_T_BUDGET_DATA.TableName = "MTPSummaryByAccountCode"
                    clsBG_T_BUDGET_DATA.Status = CStr(enumBudgetStatus.Approve)
                    clsBG_T_BUDGET_DATA.BudgetType = "E"
                    'clsBG_T_BUDGET_DATA.PrevProjectNo = Me.PrevProjectNo
                    If clsBG_T_BUDGET_DATA.Select029() = False Then
                        Return False
                    End If

            End Select

        End If

        ds = clsBG_T_BUDGET_DATA.DS

        Me.AccountCodeData = ds
        Return True

    End Function

    Public Function GetBudgetStatus() As Boolean

        Dim clsBG_T_BUDGET_HEADER As BG_T_BUDGET_HEADER = New BG_T_BUDGET_HEADER()

        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.BudgetType = "E"
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
            Return False
        End If

        Me.AuthImage = clsBG_M_SETTINGS.dtResult
        clsBG_M_SETTINGS = Nothing
        Return True

    End Function

    Public Function getBudgetAdjust() As Boolean
        Dim dt As New DataTable()
        Dim clsBG_T_BUDGET_DATA As BG_T_BUDGET_DATA = New BG_T_BUDGET_DATA()

        clsBG_T_BUDGET_DATA.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_DATA.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_DATA.ProjectNo = Me.ProjectNo

        If clsBG_T_BUDGET_DATA.Select018() = False Then
            Return False
        End If

        dt = clsBG_T_BUDGET_DATA.dtResult

        Me.BudgetAdjustTable = dt
        Return True

    End Function

#End Region

End Class
