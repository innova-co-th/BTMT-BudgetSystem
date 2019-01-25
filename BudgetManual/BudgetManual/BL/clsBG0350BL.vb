Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class clsBG0350BL

#Region "Variable"
    Private myBudgetYear As String = String.Empty
    Private myProjectNo As String = String.Empty
    Private myBudgetOrder As String = String.Empty
    Private myPeriodType As String = String.Empty
    Private myCheckPeriodType As String = String.Empty
    Private myDataType As String = String.Empty
    Private myDataList As ArrayList
    Private myRevNo As String = String.Empty
#End Region

#Region "Property"

#Region "BudgetYear"
    Public Property BudgetYear() As String
        Get
            Return myBudgetYear
        End Get
        Set(ByVal value As String)
            myBudgetYear = value
        End Set
    End Property
#End Region

#Region "ProjectNo"
    Public Property ProjectNo() As String
        Get
            Return myProjectNo
        End Get
        Set(ByVal value As String)
            myProjectNo = value
        End Set
    End Property
#End Region

#Region "PeriodType"
    Public Property PeriodType() As String
        Get
            Return myPeriodType
        End Get
        Set(ByVal value As String)
            myPeriodType = value
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

#Region "CheckPeriodType"
    Public Property CheckPeriodType() As String
        Get
            Return myCheckPeriodType
        End Get
        Set(ByVal value As String)
            myCheckPeriodType = value
        End Set
    End Property
#End Region

#Region "BudgetOrder"
    Public Property BudgetOrder() As String
        Get
            Return myBudgetOrder
        End Get
        Set(ByVal value As String)
            myBudgetOrder = value
        End Set
    End Property
#End Region

#Region "DataType"
    Public Property DataType() As String
        Get
            Return myDataType
        End Get
        Set(ByVal value As String)
            myDataType = value
        End Set
    End Property
#End Region

#Region "DataList"
    Public Property DataList() As ArrayList
        Get
            Return myDataList
        End Get
        Set(ByVal value As ArrayList)
            myDataList = value
        End Set
    End Property
#End Region

#End Region

#Region "Function"

    Public Function ImportData() As Boolean
        Dim rtn As Boolean = False

        If Me.PeriodType = CStr(enumPeriodType.EstimateBudget2) Or _
        Me.PeriodType = CStr(enumPeriodType.ForecastBudget2) Then

            '// Import Actual Data 
            '// -- Actual Data => Input Data Table
            If Me.DataType = CStr(enumUploadDataType.ActualData) Then

                '// Import Actual data to Budget Data table
                Dim clsBG_T_BUDGET_DATA As New BG_T_BUDGET_DATA

                clsBG_T_BUDGET_DATA.BudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_DATA.BudgetOrderNo = Me.BudgetOrder
                'clsBG_T_BUDGET_DATA.RevNo = "1"
                clsBG_T_BUDGET_DATA.RevNo = Me.RevNo
                clsBG_T_BUDGET_DATA.UserId = p_strUserId
                clsBG_T_BUDGET_DATA.ProjectNo = Me.ProjectNo

                If Me.PeriodType = CStr(enumPeriodType.EstimateBudget2) Then    '// Estimate Budget's Actual Oct
                    clsBG_T_BUDGET_DATA.PeriodType = CStr(enumPeriodType.EstimateBudget)

                    If clsBG_T_BUDGET_DATA.Select019() = True AndAlso clsBG_T_BUDGET_DATA.dtResult.Rows.Count > 0 Then
                        clsBG_T_BUDGET_DATA.M(10) = (CDbl(Me.DataList(9)) / 1000).ToString
                        rtn = clsBG_T_BUDGET_DATA.Update003()

                    Else
                        clsBG_T_BUDGET_DATA.M(10) = (CDbl(Me.DataList(9)) / 1000).ToString
                        rtn = clsBG_T_BUDGET_DATA.Insert003()

                    End If

                ElseIf Me.PeriodType = CStr(enumPeriodType.ForecastBudget2) Then  '// Forecast Budget's Actual Apr
                    clsBG_T_BUDGET_DATA.PeriodType = CStr(enumPeriodType.ForecastBudget)

                    If clsBG_T_BUDGET_DATA.Select019() = True AndAlso clsBG_T_BUDGET_DATA.dtResult.Rows.Count > 0 Then
                        clsBG_T_BUDGET_DATA.M(4) = (CDbl(Me.DataList(3)) / 1000).ToString
                        rtn = clsBG_T_BUDGET_DATA.Update004()

                    Else
                        clsBG_T_BUDGET_DATA.M(4) = (CDbl(Me.DataList(3)) / 1000).ToString
                        rtn = clsBG_T_BUDGET_DATA.Insert004()

                    End If

                End If

                Return rtn
            End If

        ElseIf Me.PeriodType = CStr(enumPeriodType.OriginalBudget3) Or _
        Me.PeriodType = CStr(enumPeriodType.EstimateBudget3) Or _
        Me.PeriodType = CStr(enumPeriodType.ForecastBudget3) Then

            '// Import Input Data
            '// -- Budget Data => Input Data Table
            If Me.DataType = CStr(enumUploadDataType.BudgetData) Then

                '// Import Budget data to Budget Data table
                Dim clsBG_T_BUDGET_DATA As New BG_T_BUDGET_DATA

                clsBG_T_BUDGET_DATA.BudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_DATA.BudgetOrderNo = Me.BudgetOrder
                'clsBG_T_BUDGET_DATA.RevNo = "1"
                clsBG_T_BUDGET_DATA.RevNo = Me.RevNo
                clsBG_T_BUDGET_DATA.UserId = p_strUserId
                clsBG_T_BUDGET_DATA.ProjectNo = Me.ProjectNo

                If Me.PeriodType = CStr(enumPeriodType.OriginalBudget3) Then '// Original Budget's Input Data

                    '//-- Begin Add 2011/04/29 by S.Watcharapong
                    '// Sum 2Half data if the order is expense type
                    Dim clsBG_M_BUDGET_ORDER As New BG_M_BUDGET_ORDER
                    clsBG_M_BUDGET_ORDER.BudgetOrderNo = Me.BudgetOrder

                    If clsBG_M_BUDGET_ORDER.Select016 AndAlso clsBG_M_BUDGET_ORDER.dtResult.Rows.Count > 0 Then
                        If CStr(clsBG_M_BUDGET_ORDER.dtResult.Rows(0)![BUDGET_TYPE]) = P_BUDGET_TYPE_EXPENSE Then
                            Me.DataList(6) = CDbl(Me.DataList(6))
                            Me.DataList(7) = CDbl(Me.DataList(7))
                            Me.DataList(8) = CDbl(Me.DataList(8))
                            Me.DataList(9) = CDbl(Me.DataList(9))
                            Me.DataList(10) = CDbl(Me.DataList(10))
                            Me.DataList(11) = CDbl(Me.DataList(11))
                        End If
                    End If
                    '//-- End Add 2011/04/29

                    clsBG_T_BUDGET_DATA.PeriodType = CStr(enumPeriodType.OriginalBudget)
                    clsBG_T_BUDGET_DATA.DataList = Me.DataList

                    If clsBG_T_BUDGET_DATA.Select019() = True AndAlso clsBG_T_BUDGET_DATA.dtResult.Rows.Count > 0 Then
                        rtn = clsBG_T_BUDGET_DATA.Update005()

                    Else
                        rtn = clsBG_T_BUDGET_DATA.Insert005()

                    End If

                ElseIf Me.PeriodType = CStr(enumPeriodType.EstimateBudget3) Then '// Estimate Budget's Input Data
                    clsBG_T_BUDGET_DATA.PeriodType = CStr(enumPeriodType.EstimateBudget)
                    clsBG_T_BUDGET_DATA.DataList = Me.DataList

                    If clsBG_T_BUDGET_DATA.Select019() = True AndAlso clsBG_T_BUDGET_DATA.dtResult.Rows.Count > 0 Then
                        rtn = clsBG_T_BUDGET_DATA.Update005()

                    Else
                        rtn = clsBG_T_BUDGET_DATA.Insert005()

                    End If

                ElseIf Me.PeriodType = CStr(enumPeriodType.ForecastBudget3) Then '// Forecast Budget's Input Data
                    clsBG_T_BUDGET_DATA.PeriodType = CStr(enumPeriodType.ForecastBudget)
                    clsBG_T_BUDGET_DATA.DataList = Me.DataList

                    If clsBG_T_BUDGET_DATA.Select019() = True AndAlso clsBG_T_BUDGET_DATA.dtResult.Rows.Count > 0 Then
                        rtn = clsBG_T_BUDGET_DATA.Update005()

                    Else
                        rtn = clsBG_T_BUDGET_DATA.Insert005()

                    End If

                End If

                Return rtn
            End If

        ElseIf Me.PeriodType = CStr(enumPeriodType.ForecastBudget4) Then '// Forecast Budget's MTP Data

            '// Import MTP Data
            '// -- MTP Data => Input Data Table
            If Me.DataType = CStr(enumUploadDataType.MTPData) Then

                '// Import MTP data to Budget Data table
                Dim clsBG_T_BUDGET_DATA As New BG_T_BUDGET_DATA

                clsBG_T_BUDGET_DATA.BudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_DATA.BudgetOrderNo = Me.BudgetOrder
                clsBG_T_BUDGET_DATA.RevNo = Me.RevNo
                clsBG_T_BUDGET_DATA.UserId = p_strUserId
                clsBG_T_BUDGET_DATA.ProjectNo = Me.ProjectNo

                clsBG_T_BUDGET_DATA.PeriodType = CStr(enumPeriodType.ForecastBudget)
                clsBG_T_BUDGET_DATA.DataList = Me.DataList

                If clsBG_T_BUDGET_DATA.Select019() = True AndAlso clsBG_T_BUDGET_DATA.dtResult.Rows.Count > 0 Then
                    rtn = clsBG_T_BUDGET_DATA.Update006()

                Else
                    rtn = False

                End If

                Return rtn
            End If
        ElseIf Me.PeriodType = CStr(enumPeriodType.MTPBudget) Then
            '// Import MTP Data
            '// -- MTP Data => Input Data Table
            If Me.DataType = CStr(enumUploadDataType.MTPData) Then

                '// Import MTP data to Budget Data table
                Dim clsBG_T_BUDGET_DATA As New BG_T_BUDGET_DATA

                clsBG_T_BUDGET_DATA.BudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_DATA.BudgetOrderNo = Me.BudgetOrder
                clsBG_T_BUDGET_DATA.RevNo = Me.RevNo
                clsBG_T_BUDGET_DATA.UserId = p_strUserId
                clsBG_T_BUDGET_DATA.ProjectNo = Me.ProjectNo

                clsBG_T_BUDGET_DATA.PeriodType = CStr(enumPeriodType.MTPBudget)
                clsBG_T_BUDGET_DATA.DataList = Me.DataList

                If clsBG_T_BUDGET_DATA.Select019() = True AndAlso clsBG_T_BUDGET_DATA.dtResult.Rows.Count > 0 Then
                    rtn = clsBG_T_BUDGET_DATA.Update006()

                Else
                    rtn = clsBG_T_BUDGET_DATA.Insert006()

                End If

                Return rtn
            End If



        Else
            '// Import Data from SAP
            '// -- Budget & Actual Data => Upload Data table

            '//-- Begin Add 2011/04/29 by S.Watcharapong
            If Me.PeriodType = CStr(enumPeriodType.OriginalBudget) Then
                '// Sum 2Half data if the order is expense type
                Dim clsBG_M_BUDGET_ORDER As New BG_M_BUDGET_ORDER
                clsBG_M_BUDGET_ORDER.BudgetOrderNo = Me.BudgetOrder

                If clsBG_M_BUDGET_ORDER.Select016 AndAlso clsBG_M_BUDGET_ORDER.dtResult.Rows.Count > 0 Then
                    If CStr(clsBG_M_BUDGET_ORDER.dtResult.Rows(0)![BUDGET_TYPE]) = P_BUDGET_TYPE_EXPENSE Then
                        Me.DataList(6) = CDbl(Me.DataList(6))
                        Me.DataList(7) = CDbl(Me.DataList(7))
                        Me.DataList(8) = CDbl(Me.DataList(8))
                        Me.DataList(9) = CDbl(Me.DataList(9))
                        Me.DataList(10) = CDbl(Me.DataList(10))
                        Me.DataList(11) = CDbl(Me.DataList(11))
                    End If
                End If
            End If
            '//-- End Add 2011/04/29

            Dim clsBG_T_UPLOAD_DATA As New BG_T_UPLOAD_DATA
            clsBG_T_UPLOAD_DATA.BudgetYear = Me.BudgetYear
            clsBG_T_UPLOAD_DATA.PeriodType = Me.PeriodType
            clsBG_T_UPLOAD_DATA.BudgetOrder = Me.BudgetOrder
            clsBG_T_UPLOAD_DATA.DataType = Me.DataType
            clsBG_T_UPLOAD_DATA.DataList = Me.DataList
            clsBG_T_UPLOAD_DATA.UserId = p_strUserId
            clsBG_T_UPLOAD_DATA.ProjectNo = Me.ProjectNo

            '// Check data exist
            If clsBG_T_UPLOAD_DATA.Select001() = True Then

                If clsBG_T_UPLOAD_DATA.dtResult.Rows.Count = 0 Then
                    '// Add new record
                    Return clsBG_T_UPLOAD_DATA.Insert001()

                Else
                    '// Update exists record
                    Return clsBG_T_UPLOAD_DATA.Update001()

                End If

            Else
                Return False

            End If
        End If
    End Function

    Public Function CheckPeroidExist() As Boolean
        Dim blnPeroidExist As Boolean = False

        Dim clsBG_T_BUDGET_PERIOD As New BG_T_BUDGET_PERIOD

        clsBG_T_BUDGET_PERIOD.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_PERIOD.PeriodType = Me.CheckPeriodType
        clsBG_T_BUDGET_PERIOD.ProjectNo = Me.ProjectNo

        If clsBG_T_BUDGET_PERIOD.Select003() = True Then

            If clsBG_T_BUDGET_PERIOD.dtResult.Rows.Count > 0 Then

                blnPeroidExist = True

            End If

        End If

        Return blnPeroidExist

    End Function

    Public Function CheckBudgetDataExist() As Boolean
        Dim blnPeroidExist As Boolean = False

        Dim clsBG_T_BUDGET_DATA As New BG_T_BUDGET_DATA

        clsBG_T_BUDGET_DATA.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_DATA.PeriodType = Me.CheckPeriodType
        clsBG_T_BUDGET_DATA.ProjectNo = Me.ProjectNo
        clsBG_T_BUDGET_DATA.RevNo = Me.RevNo

        If clsBG_T_BUDGET_DATA.Select031() = True Then

            If clsBG_T_BUDGET_DATA.dtResult.Rows.Count > 0 Then

                CheckBudgetDataExist = True

            End If

        End If

        Return CheckBudgetDataExist

    End Function

    Public Function CheckBudgetHeaderExist() As Boolean
        Dim blnHeaderExist As Boolean = False

        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER

        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.CheckPeriodType
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo
        clsBG_T_BUDGET_HEADER.RevNo = Me.RevNo
        clsBG_T_BUDGET_HEADER.BudgetOrderNo = Me.BudgetOrder

        If clsBG_T_BUDGET_HEADER.Select015 = True Then

            If clsBG_T_BUDGET_HEADER.dtResult.Rows.Count > 0 Then

                blnHeaderExist = True

            End If

        End If

        Return blnHeaderExist

    End Function

#End Region

End Class
