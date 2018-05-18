Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class clsBG0480BL

#Region "Variable"
    Private strBudgetYear As String = String.Empty
    Private strProjectNo As String = String.Empty
    Private strPeriodType As String = String.Empty
    Private strPIC As String = String.Empty
    Private dtPersonInCharge As DataTable = Nothing
    Private dsBudgetData As DataSet = Nothing
    Private myBudgetStatus As Integer = 0
    Private myAuthImage As DataTable = Nothing
    Private myUserLevelId As Integer = 0I
    Private strRevNo As String = String.Empty
    Private strPrevProjectNo As String = String.Empty
    Private strPrevRevNo As String = String.Empty
    Private strBudgetType As String = String.Empty
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

    Property PrevProjectNo() As String
        Get
            Return strPrevProjectNo
        End Get
        Set(ByVal value As String)
            strPrevProjectNo = value
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

    Property PrevRevNo() As String
        Get
            Return strPrevRevNo
        End Get
        Set(ByVal value As String)
            strPrevRevNo = value
        End Set
    End Property

    Property BudgetType() As String
        Get
            Return strBudgetType
        End Get
        Set(ByVal value As String)
            strBudgetType = value
        End Set
    End Property

#End Region

#Region "Function"

    Public Function GetCommentData() As Boolean

        Dim ds As New DataSet
        Dim dtRearrange As New DataTable
        Dim BG_T_BUDGET_COMMENT As BG_T_BUDGET_COMMENT = New BG_T_BUDGET_COMMENT()
        'Dim clsBG_T_BUDGET_REFERENCE As BG_T_BUDGET_REFERENCE = New BG_T_BUDGET_REFERENCE

        BG_T_BUDGET_COMMENT.BudgetYear = Me.BudgetYear
        BG_T_BUDGET_COMMENT.PeriodType = Me.PeriodType
        BG_T_BUDGET_COMMENT.UserPIC = Me.PIC
        BG_T_BUDGET_COMMENT.ProjectNo = Me.ProjectNo
        BG_T_BUDGET_COMMENT.BudgetType = Me.BudgetType

        '// Get Reference Budget
        'clsBG_T_BUDGET_DATA.RefBudgetYear = "1"
        'clsBG_T_BUDGET_DATA.RefPeriodType = "1"
        'clsBG_T_BUDGET_DATA.RefProjectNo = "1"
        'clsBG_T_BUDGET_DATA.RefRevNo = "1"
        'clsBG_T_BUDGET_DATA.RefEstProjectNo = "1"
        'clsBG_T_BUDGET_DATA.RefEstRevNo = "1"
        'clsBG_T_BUDGET_DATA.RefRBProjectNo = "1"
        'clsBG_T_BUDGET_DATA.RefRBRevNo = "1"
        'clsBG_T_BUDGET_DATA.MtpProjectNo = "1"
        'clsBG_T_BUDGET_DATA.MtpRevNo = "1"

        'If Me.PeriodType = CStr(enumPeriodType.OriginalBudget) Then

        '    '// Ref. Estimate 
        '    clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
        '    clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
        '    clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
        '    clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
        '    clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(enumPeriodType.EstimateBudget)

        '    If Me.RevNo = "" Then
        '        If clsBG_T_BUDGET_REFERENCE.Select002 = False Then

        '            Return False

        '        End If

        '    Else
        '        If clsBG_T_BUDGET_REFERENCE.Select001 = False Then

        '            Return False

        '        End If

        '    End If

        '    If clsBG_T_BUDGET_REFERENCE.dtResult.Rows.Count > 0 Then

        '        clsBG_T_BUDGET_DATA.RefEstProjectNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_PROJECT_NO").ToString
        '        clsBG_T_BUDGET_DATA.RefEstRevNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_REV_NO").ToString

        '    End If

        '    '// Ref. MTP
        '    clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
        '    clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
        '    clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
        '    clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
        '    clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(enumPeriodType.MTPBudget)

        '    If Me.RevNo = "" Then
        '        If clsBG_T_BUDGET_REFERENCE.Select002 = False Then

        '            Return False

        '        End If

        '    Else
        '        If clsBG_T_BUDGET_REFERENCE.Select001 = False Then

        '            Return False

        '        End If

        '    End If

        '    If clsBG_T_BUDGET_REFERENCE.dtResult.Rows.Count > 0 Then

        '        'clsBG_T_BUDGET_DATA.RefBudgetYear = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_BUDGET_YEAR").ToString
        '        'clsBG_T_BUDGET_DATA.RefPeriodType = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_PERIOD_TYPE").ToString
        '        clsBG_T_BUDGET_DATA.MtpProjectNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_PROJECT_NO").ToString
        '        clsBG_T_BUDGET_DATA.MtpRevNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_REV_NO").ToString

        '    End If

        'ElseIf Me.PeriodType = CStr(enumPeriodType.MTPBudget) Then

        '    '// Ref. Revise  
        '    clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
        '    clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
        '    clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
        '    clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
        '    clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(enumPeriodType.ReviseBudget)


        '    If Me.RevNo = "" Then
        '        If clsBG_T_BUDGET_REFERENCE.Select002 = False Then

        '            Return False

        '        End If

        '    Else
        '        If clsBG_T_BUDGET_REFERENCE.Select001 = False Then

        '            Return False

        '        End If

        '    End If

        '    If clsBG_T_BUDGET_REFERENCE.dtResult.Rows.Count > 0 Then

        '        clsBG_T_BUDGET_DATA.RefRBProjectNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_PROJECT_NO").ToString
        '        clsBG_T_BUDGET_DATA.RefRBRevNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_REV_NO").ToString

        '    End If

        '    '// Ref. MTP Previous
        '    clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
        '    clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
        '    clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
        '    clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
        '    clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(enumPeriodType.MTPBudget)

        '    If Me.RevNo = "" Then
        '        If clsBG_T_BUDGET_REFERENCE.Select002 = False Then

        '            Return False

        '        End If

        '    Else
        '        If clsBG_T_BUDGET_REFERENCE.Select001 = False Then

        '            Return False

        '        End If

        '    End If

        '    If clsBG_T_BUDGET_REFERENCE.dtResult.Rows.Count > 0 Then

        '        'clsBG_T_BUDGET_DATA.RefBudgetYear = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_BUDGET_YEAR").ToString
        '        'clsBG_T_BUDGET_DATA.RefPeriodType = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_PERIOD_TYPE").ToString
        '        clsBG_T_BUDGET_DATA.MtpProjectNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_PROJECT_NO").ToString
        '        clsBG_T_BUDGET_DATA.MtpRevNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_REV_NO").ToString

        '    End If

        'End If


        If Me.UserLevelId = enumUserLevel.SystemAdministrator Then

            '// Admin user.
            BG_T_BUDGET_COMMENT.RevNo = Me.RevNo

            If BG_T_BUDGET_COMMENT.Select002() = False Then
                Return False
            End If

            dtRearrange = RearrangeDatatable(BG_T_BUDGET_COMMENT.dtResult)
            dtRearrange.TableName = "COMMENT_BY_PIC"
            'Select Case Me.PeriodType
            '    Case CStr(enumPeriodType.OriginalBudget)
            '        'clsBG_T_BUDGET_DATA.MtpProjectNo = Me.PrevProjectNo
            '        'clsBG_T_BUDGET_DATA.MtpRevNo = Me.PrevRevNo
            '        If clsBG_T_BUDGET_DATA.Select004_9() = False Then
            '            Return False
            '        End If
            '        clsBG_T_BUDGET_DATA.dtResult.TableName = "BUDGET_DATA"
            '        Exit Select
            '    Case CStr(enumPeriodType.EstimateBudget)
            '        clsBG_T_BUDGET_DATA.RefPeriodType = CStr(enumPeriodType.OriginalBudget)
            '        If clsBG_T_BUDGET_DATA.Select004_10() = False Then
            '            Return False
            '        End If
            '        clsBG_T_BUDGET_DATA.dtResult.TableName = "ESTIMATE_BUDGET"
            '        Exit Select
            '    Case CStr(enumPeriodType.ReviseBudget)
            '        If clsBG_T_BUDGET_DATA.Select004_11() = False Then
            '            Return False
            '        End If
            '        clsBG_T_BUDGET_DATA.dtResult.TableName = "REVISE_BUDGET"
            '        Exit Select

            '    Case CStr(enumPeriodType.MTPBudget)
            '        'clsBG_T_BUDGET_DATA.PrevProjectNo = Me.PrevProjectNo
            '        'clsBG_T_BUDGET_DATA.PrevMTPRevNo = Me.PrevRevNo
            '        clsBG_T_BUDGET_DATA.RefPeriodType = CStr(enumPeriodType.OriginalBudget)
            '        If clsBG_T_BUDGET_DATA.Select004_12() = False Then
            '            Return False
            '        End If
            '        clsBG_T_BUDGET_DATA.dtResult.TableName = "MTP_BUDGET"
            '        Exit Select

            '    Case Else
            '        Return False
            'End Select

        Else
            If BG_T_BUDGET_COMMENT.Select002() = False Then
                Return False
            End If

            dtRearrange = RearrangeDatatable(BG_T_BUDGET_COMMENT.dtResult)
            dtRearrange.TableName = "COMMENT_BY_PIC"

            'Select Case Me.PeriodType
            '    Case CStr(enumPeriodType.OriginalBudget)
            '        'clsBG_T_BUDGET_DATA.MtpProjectNo = Me.PrevProjectNo
            '        If clsBG_T_BUDGET_DATA.Select004_1() = False Then
            '            Return False
            '        End If
            '        clsBG_T_BUDGET_DATA.dtResult.TableName = "BUDGET_DATA"
            '        Exit Select
            '    Case CStr(enumPeriodType.EstimateBudget)
            '        clsBG_T_BUDGET_DATA.RefPeriodType = CStr(enumPeriodType.OriginalBudget)
            '        If clsBG_T_BUDGET_DATA.Select004_2() = False Then
            '            Return False
            '        End If
            '        clsBG_T_BUDGET_DATA.dtResult.TableName = "ESTIMATE_BUDGET"
            '        Exit Select
            '    Case CStr(enumPeriodType.ReviseBudget)
            '        If clsBG_T_BUDGET_DATA.Select004_3() = False Then
            '            Return False
            '        End If
            '        clsBG_T_BUDGET_DATA.dtResult.TableName = "REVISE_BUDGET"
            '        Exit Select

            '    Case CStr(enumPeriodType.MTPBudget)
            '        'clsBG_T_BUDGET_DATA.PrevProjectNo = Me.PrevProjectNo
            '        clsBG_T_BUDGET_DATA.RefPeriodType = CStr(enumPeriodType.OriginalBudget)
            '        If clsBG_T_BUDGET_DATA.Select004_6() = False Then
            '            Return False
            '        End If
            '        clsBG_T_BUDGET_DATA.dtResult.TableName = "MTP_BUDGET"
            '        Exit Select

            '    Case Else
            '        Return False
            'End Select

        End If

        ds.Tables.Add(dtRearrange)

        If Me.GetAuthImage() = True Then
            ds.Tables.Add(Me.AuthImage)
        End If

        Me.BudgetData = ds
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

    Private Function RearrangeDatatable(ByVal dt As DataTable) As DataTable
        Dim dtNew As New DataTable
        'Add Column
        dtNew = AddReportColumnData(dtNew)



        Return dtNew
    End Function

    Private Function AddReportColumnData(ByVal dt As DataTable) As DataTable

        dt.Columns.Add("BUDGET_YEAR", GetType(String))
        dt.Columns.Add("BUDGET_ORDER_NO", GetType(String))
        dt.Columns.Add("BUDGET_ORDER_NAME", GetType(String))
        dt.Columns.Add("PERSON_IN_CHARGE", GetType(String))
        dt.Columns.Add("PERSON_IN_CHARGE_NAME", GetType(String))
        dt.Columns.Add("MONTH", GetType(String))
        dt.Columns.Add("COMMENT", GetType(String))

        Return dt
    End Function

#End Region

End Class
