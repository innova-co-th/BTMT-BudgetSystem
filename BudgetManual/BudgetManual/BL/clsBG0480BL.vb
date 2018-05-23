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

        BG_T_BUDGET_COMMENT.BudgetYear = Me.BudgetYear
        BG_T_BUDGET_COMMENT.PeriodType = Me.PeriodType
        BG_T_BUDGET_COMMENT.UserPIC = Me.PIC
        BG_T_BUDGET_COMMENT.ProjectNo = Me.ProjectNo
        BG_T_BUDGET_COMMENT.BudgetType = Me.BudgetType


        If Me.UserLevelId = enumUserLevel.SystemAdministrator Then

            '// Admin user.
            BG_T_BUDGET_COMMENT.RevNo = Me.RevNo

            If BG_T_BUDGET_COMMENT.Select002_1() = False Then
                Return False
            End If

            dtRearrange = RearrangeDatatable(BG_T_BUDGET_COMMENT.dtResult)
            dtRearrange.TableName = "COMMENT_BY_PIC"         

        Else
            If BG_T_BUDGET_COMMENT.Select002_2() = False Then
                Return False
            End If

            dtRearrange = RearrangeDatatable(BG_T_BUDGET_COMMENT.dtResult)
            dtRearrange.TableName = "COMMENT_BY_PIC"

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
        Dim row As DataRow

        Try
            'Add Column
            dtNew = AddReportColumnData(dtNew)

            For Each dr As DataRow In dt.Rows
                For i As Integer = 6 To 22
                    If Not IsDBNull(dr(i)) Then

                        If Not String.IsNullOrEmpty(dr(i).ToString().Trim()) Then
                            row = dtNew.NewRow()
                            row("BUDGET_YEAR") = dr(0).ToString()
                            row("BUDGET_ORDER_NO") = dr(2).ToString()
                            row("BUDGET_ORDER_NAME") = dr(3).ToString()
                            row("PERSON_IN_CHARGE") = dr(4).ToString()
                            row("PERSON_IN_CHARGE_NAME") = dr(5).ToString()
                            row("MONTH") = GetMonth(dt.Columns(i).ColumnName)
                            row("COMMENT") = dr(i).ToString()

                            dtNew.Rows.Add(row)
                        End If
                    End If
                Next

            Next
        Catch ex As Exception
            dtNew = New DataTable
            'Add Column
            dtNew = AddReportColumnData(dtNew)
        End Try

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

    Private Function GetMonth(ByVal colName As String) As String

        Dim result As String
        Dim number As Integer
        Dim hasYear As Boolean = Int32.TryParse(Me.BudgetYear, number)

        Select Case colName
            Case "M1"
                result = "Jan"
            Case "M2"
                result = "Feb"
            Case "M3"
                result = "Mar"
            Case "M4"
                result = "Apr"
            Case "M5"
                result = "May"
            Case "M6"
                result = "Jun"
            Case "M7"
                result = "Jul"
            Case "M8"
                result = "Aug"
            Case "M9"
                result = "Sep"
            Case "M10"
                result = "Oct"
            Case "M11"
                result = "Nov"
            Case "M12"
                result = "Dec"
            Case "RRT1"
                If (hasYear) Then
                    result = (number + 1).ToString
                Else
                    result = " "
                End If
            Case "RRT2"
                If (hasYear) Then
                    result = (number + 2).ToString
                Else
                    result = " "
                End If
            Case "RRT3"
                If (hasYear) Then
                    result = (number + 3).ToString
                Else
                    result = " "
                End If
            Case "RRT4"
                If (hasYear) Then
                    result = (number + 4).ToString
                Else
                    result = " "
                End If
            Case "RRT5"
                If (hasYear) Then
                    result = (number + 5).ToString
                Else
                    result = " "
                End If
            Case Else
                result = " "
        End Select
        Return result
    End Function

#End Region

End Class
