Imports System.Data.SqlClient
Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class clsBG0620BL

#Region "Variable"
    Private myDtResult As DataTable
    Private clsBG_M_ACCOUNT As BG_M_ACCOUNT
    Private clsBG_M_ASSET_GROUP As BG_M_ASSET_GROUP
    Private clsBG_M_DEPT As BG_M_DEPT
    Private clsBG_M_BUDGET_ORDER As BG_M_BUDGET_ORDER
    Private clsBG_M_PERSON_IN_CHARGE As BG_M_PERSON_IN_CHARGE
    Private clsBG_T_BUDGET_HEADER As BG_T_BUDGET_HEADER
    Private myBGOrderNo As String = String.Empty
    Private myBGOrderName As String = String.Empty
    Private myBGType As String = String.Empty
    Private myAccount As String = String.Empty
    Private myCostCenter As String = String.Empty
    Private myCostType As String = String.Empty
    Private myCost As String = String.Empty
    Private myAssetGroup As String = String.Empty
    Private myDepartment As String = String.Empty
    Private myPersonInCharge As String = String.Empty
    Private myActiveFlag As String = String.Empty
    Private myUpdateUserId As String = String.Empty
    Private myExpenseType As String = String.Empty
    Private myPICShowFlag As String = String.Empty
    Private myCreateUserId As String = String.Empty
    Private myCreateDate As String = String.Empty
    Private myRemarks As String = String.Empty
    Private myBGOrderNoFilter As String = String.Empty
    Private myBGOrderNameFilter As String = String.Empty
    Private myBGTypeFilter As String = String.Empty
    Private myAccountFilter As String = String.Empty
    Private myCostCenterFilter As String = String.Empty
    Private myCostTypeFilter As String = String.Empty
    Private myCostFilter As String = String.Empty
    Private myAssetGroupFilter As String = String.Empty
    Private myDepartmentFilter As String = String.Empty
    Private myPersonInChargeFilter As String = String.Empty
    Private myActiveFlagFilter As String = String.Empty
    Private myExpenseTypeFilter As String = String.Empty
#End Region

#Region "Property"
    Public Property DtResult() As DataTable
        Get
            Return myDtResult
        End Get
        Set(ByVal value As DataTable)
            myDtResult = value
        End Set
    End Property
    Public Property BudgetOrderNo() As String
        Get
            Return myBGOrderNo
        End Get
        Set(ByVal value As String)
            myBGOrderNo = value
        End Set
    End Property
    Public Property BudgetOrderName() As String
        Get
            Return myBGOrderName
        End Get
        Set(ByVal value As String)
            myBGOrderName = value
        End Set
    End Property
    Public Property BudgetType() As String
        Get
            Return myBGType
        End Get
        Set(ByVal value As String)
            myBGType = value
        End Set
    End Property
    Public Property Account() As String
        Get
            Return myAccount
        End Get
        Set(ByVal value As String)
            myAccount = value
        End Set
    End Property
    Public Property CostCenter() As String
        Get
            Return myCostCenter
        End Get
        Set(ByVal value As String)
            myCostCenter = value
        End Set
    End Property
    Public Property CostType() As String
        Get
            Return myCostType
        End Get
        Set(ByVal value As String)
            myCostType = value
        End Set
    End Property
    Public Property Cost() As String
        Get
            Return myCost
        End Get
        Set(ByVal value As String)
            myCost = value
        End Set
    End Property
    Public Property AssetGroup() As String
        Get
            Return myAssetGroup
        End Get
        Set(ByVal value As String)
            myAssetGroup = value
        End Set
    End Property
    Public Property Department() As String
        Get
            Return myDepartment
        End Get
        Set(ByVal value As String)
            myDepartment = value
        End Set
    End Property
    Public Property PersonInCharge() As String
        Get
            Return myPersonInCharge
        End Get
        Set(ByVal value As String)
            myPersonInCharge = value
        End Set
    End Property
    Public Property ActiveFlag() As String
        Get
            Return myActiveFlag
        End Get
        Set(ByVal value As String)
            myActiveFlag = value
        End Set
    End Property
    Public Property UpdateUserId() As String
        Get
            Return myUpdateUserId
        End Get
        Set(ByVal value As String)
            myUpdateUserId = value
        End Set
    End Property
    Public Property ExpenseType() As String
        Get
            Return myExpenseType
        End Get
        Set(ByVal value As String)
            myExpenseType = value
        End Set
    End Property
    Public Property PICShowFlag() As String
        Get
            Return myPICShowFlag
        End Get
        Set(ByVal value As String)
            myPICShowFlag = value
        End Set
    End Property
    Public Property CreateUserId() As String
        Get
            Return myCreateUserId
        End Get
        Set(ByVal value As String)
            myCreateUserId = value
        End Set
    End Property
    Public Property CreateDate() As String
        Get
            Return myCreateDate
        End Get
        Set(ByVal value As String)
            myCreateDate = value
        End Set
    End Property
    Public Property Remarks() As String
        Get
            Return myRemarks
        End Get
        Set(ByVal value As String)
            myRemarks = value
        End Set
    End Property

    Public Property BudgetOrderNoFilter() As String
        Get
            Return myBGOrderNoFilter
        End Get
        Set(ByVal value As String)
            myBGOrderNoFilter = value
        End Set
    End Property
    Public Property BudgetOrderNameFilter() As String
        Get
            Return myBGOrderNameFilter
        End Get
        Set(ByVal value As String)
            myBGOrderNameFilter = value
        End Set
    End Property
    Public Property BudgetTypeFilter() As String
        Get
            Return myBGTypeFilter
        End Get
        Set(ByVal value As String)
            myBGTypeFilter = value
        End Set
    End Property
    Public Property AccountFilter() As String
        Get
            Return myAccountFilter
        End Get
        Set(ByVal value As String)
            myAccountFilter = value
        End Set
    End Property
    Public Property CostCenterFilter() As String
        Get
            Return myCostCenterFilter
        End Get
        Set(ByVal value As String)
            myCostCenterFilter = value
        End Set
    End Property
    Public Property CostTypeFilter() As String
        Get
            Return myCostTypeFilter
        End Get
        Set(ByVal value As String)
            myCostTypeFilter = value
        End Set
    End Property
    Public Property CostFilter() As String
        Get
            Return myCostFilter
        End Get
        Set(ByVal value As String)
            myCostFilter = value
        End Set
    End Property
    Public Property AssetGroupFilter() As String
        Get
            Return myAssetGroupFilter
        End Get
        Set(ByVal value As String)
            myAssetGroupFilter = value
        End Set
    End Property
    Public Property DepartmentFilter() As String
        Get
            Return myDepartmentFilter
        End Get
        Set(ByVal value As String)
            myDepartmentFilter = value
        End Set
    End Property
    Public Property PersonInChargeFilter() As String
        Get
            Return myPersonInChargeFilter
        End Get
        Set(ByVal value As String)
            myPersonInChargeFilter = value
        End Set
    End Property
    Public Property ActiveFlagFilter() As String
        Get
            Return myActiveFlagFilter
        End Get
        Set(ByVal value As String)
            myActiveFlagFilter = value
        End Set
    End Property
    Public Property ExpenseTypeFilter() As String
        Get
            Return myExpenseTypeFilter
        End Get
        Set(ByVal value As String)
            myExpenseTypeFilter = value
        End Set
    End Property

#End Region

#Region "Function"

    ''' <summary>
    ''' Get Account list
    ''' </summary>
    ''' <returns></returns>

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
    ''' Get Department list
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getDepartmentList() As Boolean
        clsBG_M_DEPT = New BG_M_DEPT

        If clsBG_M_DEPT.Select001 Then
            myDtResult = clsBG_M_DEPT.DtResult
        Else
            myDtResult = New DataTable
        End If

        Return True
    End Function

    ''' <summary>
    ''' Get budget order list
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getBudgetOrderList() As Boolean
        clsBG_M_BUDGET_ORDER = New BG_M_BUDGET_ORDER

        clsBG_M_BUDGET_ORDER.BudgetOrderNo = Me.BudgetOrderNoFilter
        clsBG_M_BUDGET_ORDER.BudgetOrderName = Me.BudgetOrderNameFilter
        clsBG_M_BUDGET_ORDER.BudgetType = Me.BudgetTypeFilter
        clsBG_M_BUDGET_ORDER.Account = Me.AccountFilter
        clsBG_M_BUDGET_ORDER.CostCenter = Me.CostCenterFilter
        clsBG_M_BUDGET_ORDER.CostType = Me.CostTypeFilter
        clsBG_M_BUDGET_ORDER.Cost = Me.CostFilter
        clsBG_M_BUDGET_ORDER.AssetGroup = Me.AssetGroupFilter
        clsBG_M_BUDGET_ORDER.Department = Me.DepartmentFilter
        clsBG_M_BUDGET_ORDER.PersonInCharge = Me.PersonInChargeFilter
        clsBG_M_BUDGET_ORDER.ActiveFlag = Me.ActiveFlagFilter
        clsBG_M_BUDGET_ORDER.ExpenseType = Me.ExpenseTypeFilter

        If clsBG_M_BUDGET_ORDER.Select002 Then
            myDtResult = clsBG_M_BUDGET_ORDER.dtResult
        Else
            myDtResult = New DataTable
        End If

        Return True
    End Function

    ''' <summary>
    ''' Get asset group list
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getAssetGroupList() As Boolean
        clsBG_M_ASSET_GROUP = New BG_M_ASSET_GROUP

        If clsBG_M_ASSET_GROUP.Select001 Then
            myDtResult = clsBG_M_ASSET_GROUP.DtResult
        Else
            myDtResult = New DataTable
        End If

        Return True
    End Function

    ''' <summary>
    ''' Get person in charge list
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getPersonInChargeList() As Boolean
        clsBG_M_PERSON_IN_CHARGE = New BG_M_PERSON_IN_CHARGE

        If clsBG_M_PERSON_IN_CHARGE.Select001 Then
            myDtResult = clsBG_M_PERSON_IN_CHARGE.DtResult
        Else
            myDtResult = New DataTable
        End If

        Return True
    End Function

    ''' <summary>
    ''' Save budget order data
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function saveBudgetOrder() As Boolean
        Dim dtHeader As DataTable
        Dim dtAllHDActive As DataTable
        Dim conn As SqlConnection
        Dim trans As SqlTransaction
        Dim result As Boolean = False
        Dim strOldPIC As String

        clsBG_M_BUDGET_ORDER = New BG_M_BUDGET_ORDER
        clsBG_T_BUDGET_HEADER = New BG_T_BUDGET_HEADER

        clsBG_M_BUDGET_ORDER.BudgetOrderNo = Me.BudgetOrderNo
        If clsBG_M_BUDGET_ORDER.Select017 Then
            Me.DtResult = clsBG_M_BUDGET_ORDER.dtResult
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()
            trans = conn.BeginTransaction()

            Try
                clsBG_M_BUDGET_ORDER.BudgetOrderName = Me.BudgetOrderName
                clsBG_M_BUDGET_ORDER.BudgetType = Me.BudgetType
                clsBG_M_BUDGET_ORDER.Account = Me.Account
                clsBG_M_BUDGET_ORDER.CostCenter = Me.CostCenter
                clsBG_M_BUDGET_ORDER.CostType = Me.CostType
                clsBG_M_BUDGET_ORDER.Cost = Me.Cost
                clsBG_M_BUDGET_ORDER.AssetGroup = Me.AssetGroup
                clsBG_M_BUDGET_ORDER.Department = Me.Department
                clsBG_M_BUDGET_ORDER.PersonInCharge = Me.PersonInCharge
                clsBG_M_BUDGET_ORDER.ActiveFlag = Me.ActiveFlag
                clsBG_M_BUDGET_ORDER.ExpenseType = Me.ExpenseType
                clsBG_M_BUDGET_ORDER.PICShowFlag = Me.PICShowFlag
                clsBG_M_BUDGET_ORDER.CreateUserId = Me.CreateUserId
                clsBG_M_BUDGET_ORDER.UpdateUserId = Me.UpdateUserId
                clsBG_M_BUDGET_ORDER.Remarks = Me.Remarks

                If Me.DtResult.Rows.Count = 1 Then  '// Update data

                    strOldPIC = DtResult.Rows(0).Item("PERSON_IN_CHARGE_NO").ToString

                    If clsBG_M_BUDGET_ORDER.Update001(conn, trans) Then
                        result = True
                    End If

                    'Check Header by PIC 
                    dtHeader = Nothing
                    dtAllHDActive = Nothing



                    'Get All OLD Header 
                    clsBG_T_BUDGET_HEADER.UserPIC = strOldPIC
                    If clsBG_T_BUDGET_HEADER.Select017() Then
                        dtAllHDActive = clsBG_T_BUDGET_HEADER.dtResult
                    End If

                    clsBG_T_BUDGET_HEADER.UserPIC = Me.PersonInCharge
                    If clsBG_T_BUDGET_HEADER.Select016() Then
                        dtHeader = clsBG_T_BUDGET_HEADER.dtResult
                    End If

                    Dim drF() As DataRow
                    If Not dtHeader Is Nothing AndAlso dtHeader.Rows.Count > 0 Then
                        If Not dtAllHDActive Is Nothing AndAlso dtAllHDActive.Rows.Count > 0 Then
                            For i As Integer = 0 To dtAllHDActive.Rows.Count - 1
                                drF = Nothing

                                drF = dtHeader.Select("BUDGET_YEAR=" & dtAllHDActive.Rows(i).Item("BUDGET_YEAR").ToString & " AND PERIOD_TYPE=" & dtAllHDActive.Rows(i).Item("PERIOD_TYPE").ToString _
                                      & " AND PERSON_IN_CHARGE_NO='" & Me.PersonInCharge & "' AND BUDGET_TYPE='" & dtAllHDActive.Rows(i).Item("BUDGET_TYPE").ToString _
                                      & "' AND REV_NO=" & dtAllHDActive.Rows(i).Item("REV_NO").ToString)

                                If Not drF Is Nothing AndAlso drF.Length > 0 Then
                                    ' Not Insert 
                                Else
                                    '// Set Parameters
                                    clsBG_T_BUDGET_HEADER.BudgetYear = dtAllHDActive.Rows(i).Item("BUDGET_YEAR").ToString
                                    clsBG_T_BUDGET_HEADER.PeriodType = dtAllHDActive.Rows(i).Item("PERIOD_TYPE").ToString
                                    clsBG_T_BUDGET_HEADER.BudgetType = dtAllHDActive.Rows(i).Item("BUDGET_TYPE").ToString
                                    clsBG_T_BUDGET_HEADER.UserPIC = Me.PersonInCharge
                                    clsBG_T_BUDGET_HEADER.RevNo = dtAllHDActive.Rows(i).Item("REV_NO").ToString
                                    clsBG_T_BUDGET_HEADER.Status = CStr(enumBudgetStatus.NewRecord)
                                    clsBG_T_BUDGET_HEADER.UserId = Me.CreateUserId
                                    clsBG_T_BUDGET_HEADER.ProjectNo = dtAllHDActive.Rows(i).Item("PROJECT_NO").ToString
                                    clsBG_T_BUDGET_HEADER.RRT1 = dtAllHDActive.Rows(i).Item("RRT1").ToString
                                    clsBG_T_BUDGET_HEADER.RRT2 = dtAllHDActive.Rows(i).Item("RRT2").ToString
                                    clsBG_T_BUDGET_HEADER.RRT3 = dtAllHDActive.Rows(i).Item("RRT3").ToString
                                    clsBG_T_BUDGET_HEADER.RRT4 = dtAllHDActive.Rows(i).Item("RRT4").ToString
                                    clsBG_T_BUDGET_HEADER.RRT5 = dtAllHDActive.Rows(i).Item("RRT5").ToString
                                    clsBG_T_BUDGET_HEADER.WorkingBG1 = dtAllHDActive.Rows(i).Item("Working_BG1").ToString
                                    clsBG_T_BUDGET_HEADER.WorkingBG2 = dtAllHDActive.Rows(i).Item("Working_BG2").ToString


                                    '// Call Function: Insert Budget Header
                                    If clsBG_T_BUDGET_HEADER.Insert001(conn, trans) = False Then
                                        Throw New Exception("Can not insert budget header!")
                                    End If
                                End If


                            Next
                        End If
                    Else
                        If Not dtAllHDActive Is Nothing AndAlso dtAllHDActive.Rows.Count > 0 Then

                            For i As Integer = 0 To dtAllHDActive.Rows.Count - 1
                                '// Set Parameters
                                clsBG_T_BUDGET_HEADER.BudgetYear = dtAllHDActive.Rows(i).Item("BUDGET_YEAR").ToString
                                clsBG_T_BUDGET_HEADER.PeriodType = dtAllHDActive.Rows(i).Item("PERIOD_TYPE").ToString
                                clsBG_T_BUDGET_HEADER.BudgetType = dtAllHDActive.Rows(i).Item("BUDGET_TYPE").ToString
                                clsBG_T_BUDGET_HEADER.UserPIC = Me.PersonInCharge
                                clsBG_T_BUDGET_HEADER.RevNo = dtAllHDActive.Rows(i).Item("REV_NO").ToString
                                clsBG_T_BUDGET_HEADER.Status = CStr(enumBudgetStatus.NewRecord)
                                clsBG_T_BUDGET_HEADER.UserId = Me.CreateUserId
                                clsBG_T_BUDGET_HEADER.ProjectNo = dtAllHDActive.Rows(i).Item("PROJECT_NO").ToString
                                clsBG_T_BUDGET_HEADER.RRT1 = dtAllHDActive.Rows(i).Item("RRT1").ToString
                                clsBG_T_BUDGET_HEADER.RRT2 = dtAllHDActive.Rows(i).Item("RRT2").ToString
                                clsBG_T_BUDGET_HEADER.RRT3 = dtAllHDActive.Rows(i).Item("RRT3").ToString
                                clsBG_T_BUDGET_HEADER.RRT4 = dtAllHDActive.Rows(i).Item("RRT4").ToString
                                clsBG_T_BUDGET_HEADER.RRT5 = dtAllHDActive.Rows(i).Item("RRT5").ToString
                                clsBG_T_BUDGET_HEADER.WorkingBG1 = dtAllHDActive.Rows(i).Item("Working_BG1").ToString
                                clsBG_T_BUDGET_HEADER.WorkingBG2 = dtAllHDActive.Rows(i).Item("Working_BG2").ToString


                                '// Call Function: Insert Budget Header
                                If clsBG_T_BUDGET_HEADER.Insert001(conn, trans) = False Then
                                    Throw New Exception("Can not insert budget header!")
                                End If
                            Next

                        End If
                    End If

                Else                                '// Add data
                    If clsBG_M_BUDGET_ORDER.Insert001(conn, trans) Then
                        result = True
                    End If
                End If

                If result Then
                    trans.Commit()
                End If
            Catch ex As Exception
                trans.Rollback()
                showErrorMessage("Error: " & ex.Message)
            Finally
                If conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End Try
        Else
            result = False
        End If

        Return result
    End Function

    ''' <summary>
    ''' Delete budget order data
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function deleteBudgetOrder() As Boolean
        Dim conn As SqlConnection
        Dim trans As SqlTransaction

        conn = New SqlConnection(My.Settings.ConnStr)
        conn.Open()
        trans = conn.BeginTransaction()

        Try
            '// Delete budget data
            Dim clsBG_T_BUDGET_DATA As New BG_T_BUDGET_DATA

            clsBG_T_BUDGET_DATA.BudgetOrderNo = Me.BudgetOrderNo

            clsBG_T_BUDGET_DATA.Delete003(conn, trans)

            '// Delete budget order
            Dim clsBG_M_BUDGET_ORDER As New BG_M_BUDGET_ORDER

            clsBG_M_BUDGET_ORDER.BudgetOrderNo = Me.BudgetOrderNo

            clsBG_M_BUDGET_ORDER.Delete001(conn, trans)

            '// Commit Transaction
            trans.Commit()

            Return True

        Catch ex As Exception

            '// Rollback Transaction
            trans.Rollback()

            '// Close Connection
            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            MessageBox.Show("Error: " & ex.Message, My.Settings.ProgramTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)

            Return False
        End Try
    End Function

    ''' <summary>
    ''' Save budget order data (for imported data from excel)
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function saveBudgetOrderAll(ByVal conn As SqlConnection, _
                                       ByVal trans As SqlTransaction) As Boolean
        Dim result As Boolean = False

        clsBG_M_BUDGET_ORDER = New BG_M_BUDGET_ORDER

        clsBG_M_BUDGET_ORDER.BudgetOrderNo = Me.BudgetOrderNo
        If clsBG_M_BUDGET_ORDER.Select003 Then
            Me.DtResult = clsBG_M_BUDGET_ORDER.dtResult

            Try
                clsBG_M_BUDGET_ORDER.BudgetOrderName = Me.BudgetOrderName
                clsBG_M_BUDGET_ORDER.BudgetType = Me.BudgetType
                clsBG_M_BUDGET_ORDER.Account = Me.Account
                clsBG_M_BUDGET_ORDER.CostCenter = Me.CostCenter
                clsBG_M_BUDGET_ORDER.CostType = Me.CostType
                clsBG_M_BUDGET_ORDER.Cost = Me.Cost
                clsBG_M_BUDGET_ORDER.AssetGroup = Me.AssetGroup
                clsBG_M_BUDGET_ORDER.Department = Me.Department
                clsBG_M_BUDGET_ORDER.PersonInCharge = Me.PersonInCharge
                clsBG_M_BUDGET_ORDER.ActiveFlag = Me.ActiveFlag
                clsBG_M_BUDGET_ORDER.ExpenseType = Me.ExpenseType
                clsBG_M_BUDGET_ORDER.PICShowFlag = Me.PICShowFlag
                clsBG_M_BUDGET_ORDER.CreateUserId = Me.CreateUserId
                clsBG_M_BUDGET_ORDER.UpdateUserId = Me.UpdateUserId
                clsBG_M_BUDGET_ORDER.Remarks = Me.Remarks

                If Me.DtResult.Rows.Count = 1 Then  '// Update data
                    If clsBG_M_BUDGET_ORDER.Update001(conn, trans) Then
                        result = True
                    End If
                Else                                '// Add data
                    If clsBG_M_BUDGET_ORDER.Insert001(conn, trans) Then
                        result = True
                    End If
                End If
            Catch ex As Exception
                result = False
                showErrorMessage("Error: " & ex.Message)
            End Try
        Else
            result = False
        End If

        Return result
    End Function

    ''' <summary>
    ''' Get PIC Show Flag
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getPICShowFlag() As Boolean
        Dim clsBG_M_BUDGET_ORDER As New BG_M_BUDGET_ORDER

        clsBG_M_BUDGET_ORDER.PersonInCharge = Me.PersonInCharge

        If clsBG_M_BUDGET_ORDER.Select015 Then
            Me.PICShowFlag = clsBG_M_BUDGET_ORDER.PICShowFlag
        Else
            Me.PICShowFlag = "1"
        End If

        Return True
    End Function

#End Region

End Class
