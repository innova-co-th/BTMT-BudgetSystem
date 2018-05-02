Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class clsBG0420BL

#Region "Variable"
    Private strBudgetYear As String = String.Empty
    Private strProjectNo As String = String.Empty
    Private strPeriodType As String = String.Empty
    Private strRevNo As String = String.Empty
    Private bMTPChecked As Boolean = False
    Private dsBudgetData As DataSet = Nothing
    Private myBudgetStatus As Integer = 0I
    Private myUserLevelId As Integer = 0I
    Private myAuthImage As DataTable = Nothing
    Private myDtResult As DataTable
    Private strPrevProjectNo As String = String.Empty
    Private strPrevRevNo As String = String.Empty
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

    Public Property UserLevelId() As Integer
        Get
            Return myUserLevelId
        End Get
        Set(ByVal value As Integer)
            myUserLevelId = value
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

#Region "dtResult"
    Property dtResult() As DataTable
        Get
            Return myDtResult
        End Get
        Set(ByVal value As DataTable)
            myDtResult = value
        End Set
    End Property
#End Region

#End Region

#Region "Function"

    Public Function GetBudgetData() As Boolean

        Dim ds As New DataSet
        Dim clsBG_T_BUDGET_DATA As BG_T_BUDGET_DATA = New BG_T_BUDGET_DATA()
        Dim clsBG_T_BUDGET_REFERENCE As BG_T_BUDGET_REFERENCE = New BG_T_BUDGET_REFERENCE

        clsBG_T_BUDGET_DATA.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_DATA.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_DATA.MTPChecked = Me.MTPChecked
        clsBG_T_BUDGET_DATA.ProjectNo = Me.ProjectNo

        '// Get Reference Budget
        clsBG_T_BUDGET_DATA.RefBudgetYear = "1"
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
            clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(enumPeriodType.MTPBudget)

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

        ElseIf Me.PeriodType = CStr(enumPeriodType.MTPBudget) Then

            '// Ref. Revise  
            clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
            clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
            clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
            clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
            clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(enumPeriodType.ReviseBudget)


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
            clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(enumPeriodType.MTPBudget)

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

            Select Case Me.PeriodType
                Case CStr(enumPeriodType.OriginalBudget)
                    'clsBG_T_BUDGET_DATA.MtpProjectNo = Me.PrevProjectNo
                    'clsBG_T_BUDGET_DATA.MtpRevNo = Me.PrevRevNo
                    If clsBG_T_BUDGET_DATA.Select014_4() = False Then
                        Return False
                    End If
                    clsBG_T_BUDGET_DATA.dtResult.TableName = "ORIGINAL_BUDGET"
                    Exit Select
                Case CStr(enumPeriodType.EstimateBudget)
                    clsBG_T_BUDGET_DATA.RefPeriodType = CStr(enumPeriodType.ReviseBudget)
                    If clsBG_T_BUDGET_DATA.Select014_5() = False Then
                        Return False
                    End If
                    clsBG_T_BUDGET_DATA.dtResult.TableName = "ESTIMATE_BUDGET"
                    Exit Select
                Case CStr(enumPeriodType.ReviseBudget)
                    If clsBG_T_BUDGET_DATA.Select014_6() = False Then
                        Return False
                    End If
                    clsBG_T_BUDGET_DATA.dtResult.TableName = "REVISE_BUDGET"
                    Exit Select

                Case CStr(enumPeriodType.MTPBudget)
                    'clsBG_T_BUDGET_DATA.PrevProjectNo = Me.PrevProjectNo
                    'clsBG_T_BUDGET_DATA.PrevMTPRevNo = Me.PrevRevNo
                    If clsBG_T_BUDGET_DATA.Select004_8() = False Then
                        Return False
                    End If
                    clsBG_T_BUDGET_DATA.dtResult.TableName = "MTP_BUDGET"
                    Exit Select

                Case Else
                    Return False
            End Select

        Else

            '// Other user.
            Select Case Me.PeriodType
                Case CStr(enumPeriodType.OriginalBudget)
                    'clsBG_T_BUDGET_DATA.MtpProjectNo = Me.PrevProjectNo
                    If clsBG_T_BUDGET_DATA.Select014_1() = False Then
                        Return False
                    End If
                    clsBG_T_BUDGET_DATA.dtResult.TableName = "ORIGINAL_BUDGET"
                    Exit Select
                Case CStr(enumPeriodType.EstimateBudget)
                    clsBG_T_BUDGET_DATA.RefPeriodType = CStr(enumPeriodType.ReviseBudget)
                    If clsBG_T_BUDGET_DATA.Select014_2() = False Then
                        Return False
                    End If
                    clsBG_T_BUDGET_DATA.dtResult.TableName = "ESTIMATE_BUDGET"
                    Exit Select
                Case CStr(enumPeriodType.ReviseBudget)
                    If clsBG_T_BUDGET_DATA.Select014_3() = False Then
                        Return False
                    End If
                    clsBG_T_BUDGET_DATA.dtResult.TableName = "REVISE_BUDGET"
                    Exit Select

                Case CStr(enumPeriodType.MTPBudget)
                    'clsBG_T_BUDGET_DATA.PrevProjectNo = Me.PrevProjectNo
                    If clsBG_T_BUDGET_DATA.Select004_7() = False Then
                        Return False
                    End If
                    clsBG_T_BUDGET_DATA.dtResult.TableName = "MTP_BUDGET"
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
