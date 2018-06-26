Imports System.Data.SqlClient
Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class BG_M_BUDGET_ORDER

#Region "Variable"
    Private myDtResult As DataTable
    Private myUserPIC As String = String.Empty
    Private myBGOrderNo As String = String.Empty
    Private myBGOrderName As String = String.Empty
    Private myBGType As String = String.Empty
    Private myProjectNo As String = String.Empty
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
    Private myMonth As String = String.Empty
    Private myYear As String = String.Empty
    Private myPeriod As String = String.Empty
    Private myDS As DataSet = Nothing
    Private myBudgetYear As String = String.Empty
    Private myPeriodType As String = String.Empty
    Private myRevNo As String = String.Empty
    Private myStatus As String = String.Empty
    Private myRemarks As String = String.Empty
#End Region

#Region "Property"

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

#Region "UserPIC"
    Property UserPIC() As String
        Get
            Return myUserPIC
        End Get
        Set(ByVal value As String)
            myUserPIC = value
        End Set
    End Property
#End Region

#Region "Budget Order No."

    Public Property BudgetOrderNo() As String
        Get
            Return myBGOrderNo
        End Get
        Set(ByVal value As String)
            myBGOrderNo = value
        End Set
    End Property

#End Region

    Public Property BudgetOrderName() As String
        Get
            Return myBGOrderName
        End Get
        Set(ByVal value As String)
            myBGOrderName = value
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
    Public Property PeriodType() As String
        Get
            Return myPeriodType
        End Get
        Set(ByVal value As String)
            myPeriodType = value
        End Set
    End Property
    Public Property RevNo() As String
        Get
            Return myRevNo
        End Get
        Set(ByVal value As String)
            myRevNo = value
        End Set
    End Property
    Public Property Status() As String
        Get
            Return myStatus
        End Get
        Set(ByVal value As String)
            myStatus = value
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
    Public Property ProjectNo() As String
        Get
            Return myProjectNo
        End Get
        Set(ByVal value As String)
            myProjectNo = value
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
    Public Property Month() As String
        Get
            Return myMonth
        End Get
        Set(ByVal value As String)
            myMonth = value
        End Set
    End Property
    Public Property Year() As String
        Get
            Return myYear
        End Get
        Set(ByVal value As String)
            myYear = value
        End Set
    End Property
    Public Property Period() As String
        Get
            Return myPeriod
        End Get
        Set(ByVal value As String)
            myPeriod = value
        End Set
    End Property
#Region " DS "
    Property DS() As DataSet
        Get
            Return myDS
        End Get
        Set(ByVal value As DataSet)
            myDS = value
        End Set
    End Property

#End Region

#Region "Remarks"
    Property Remarks() As String
        Get
            Return myRemarks
        End Get
        Set(ByVal value As String)
            myRemarks = value
        End Set
    End Property

#End Region

#End Region

#Region "Function"

    Public Function Select001() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "SELECT001")
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@UserPIC", Me.UserPIC)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Select001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    Public Function Select002() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String
        Dim strWhere As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "SELECT002")

            '// (1) Budget Order No.
            If Not Me.BudgetOrderNo.Equals(String.Empty) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " BO.BUDGET_ORDER_NO LIKE '%" & Me.BudgetOrderNo.Replace("'", "''") & "%' "

            End If

            '// (2) Budget Order Name.
            If Not Me.BudgetOrderName.Equals(String.Empty) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " BO.BUDGET_ORDER_NAME LIKE '%" & Me.BudgetOrderName.Replace("'", "''") & "%' "

            End If

            '// (3) Budget Type
            If Not Me.BudgetType.Equals(String.Empty) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " BO.BUDGET_TYPE = '" & Me.BudgetType & "' "

            End If

            '// (4) Account
            If Not Me.Account.Equals(String.Empty) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " BO.ACCOUNT_NO = '" & Me.Account & "' "

            End If

            '// (5) Cost Center
            If Not Me.CostCenter.Equals(String.Empty) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " BO.COST_CENTER LIKE '%" & Me.CostCenter.Replace("'", "''") & "%' "

            End If

            '// (6) Expenst Type
            If Not Me.ExpenseType.Equals(String.Empty) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " BO.EXPENSE_TYPE = " & Me.ExpenseType & " "

            End If

            '// (7) Cont Type
            If Not Me.CostType.Equals(String.Empty) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " BO.COST_TYPE = " & Me.CostType & " "

            End If

            '// (8) Cost
            If Not Me.Cost.Equals(String.Empty) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " BO.COST = " & Me.Cost & " "

            End If

            '// (9) Asset Group
            If Not Me.AssetGroup.Equals(String.Empty) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " BO.ASSET_GROUP_NO = '" & Me.AssetGroup & "' "

            End If

            '// (10) Department
            If Not Me.Department.Equals(String.Empty) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " BO.DEPT_NO = '" & Me.Department & "' "

            End If

            '// (11) Person In Charge
            If Not Me.PersonInCharge.Equals(String.Empty) Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " BO.PERSON_IN_CHARGE_NO = '" & Me.PersonInCharge & "' "

            End If

            '// (12) Active
            If Me.ActiveFlag.Equals("1") Then

                If Not strWhere.Equals(String.Empty) Then
                    strWhere &= " AND "
                End If
                strWhere &= " BO.ACTIVE_FLAG = " & Me.ActiveFlag & " "

            End If

            If strWhere.Equals(String.Empty) Then
                strSQL = strSQL.Replace("@Where", "")
            Else
                strSQL = strSQL.Replace("@Where", " WHERE " & strWhere)
            End If

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Select002] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    Public Function Select003() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "SELECT003")
            strSQL = strSQL.Replace("@P_BG_ORDER_NO", Me.BudgetOrderNo.Trim)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Select003] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    Public Function Select004() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "SELECT004")

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Select004] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    Public Function Select005() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "SELECT005")
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Select005] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    Public Function Select006() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "SELECT006")
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@UserPIC", Me.UserPIC)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Select006] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    Public Function Select007() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "SELECT007")

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Select007] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    Public Function Select008() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "SELECT008_1")
            strSQL = strSQL.Replace("@Month", Me.Month)

            Dim intM As Integer = CInt(Month)
            Dim strbudget As String = String.Empty
            Dim stractual As String = String.Empty
            If intM > 1 Then
                For i As Integer = 1 To intM
                    If i = 1 Then
                        strbudget = ",SUM(ISNULL(D.M1,0.00)) "
                        stractual = ",SUM(ISNULL(U.M1,0.00)) "
                    Else
                        strbudget &= "+ SUM(ISNULL(D.M" & i & ",0.00)) "
                        stractual &= "+ SUM(ISNULL(U.M" & i & ",0.00)) "
                    End If
                Next
            Else
                strbudget = ",SUM(ISNULL(D.M1,0.00)) "
                stractual = ",SUM(ISNULL(U.M1,0.00)) "
            End If

            strbudget &= " AS SUM_BUDGET "
            stractual &= " AS SUM_ACTUAL "

            strSQL &= strbudget & stractual

            strSQL &= readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "SELECT008_2")
            stractual = String.Empty

            If intM > 1 Then
                For i As Integer = 1 To intM
                    If i = 1 Then
                        stractual = " SUM(M1) AS M1"
                    Else
                        stractual &= ", SUM(M" & i & ") AS M" & i
                    End If
                Next
            Else
                stractual = "SUM(M1) AS M1 "
            End If
            strSQL &= stractual

            strSQL &= readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "SELECT008_3")
            strSQL = strSQL.Replace("@year", Me.Year)
            strSQL = strSQL.Replace("@period", Me.Period)

            da = New SqlDataAdapter(strSQL, conn)
            DS = New DataSet

            da.Fill(DS, "BUDGET_MANAGEMENT")

            Me.DS = DS

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Select008] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' Select all PIC of Active Order 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select009() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "SELECT009")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@Status", Me.Status)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Select009] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' Select Related PIC of Active Order 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select010() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "SELECT010")
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@UserPIC", Me.UserPIC)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Select010] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    Public Function Select011() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "SELECT011")
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Select011] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    Public Function Select012() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "SELECT012")
            strSQL = strSQL.Replace("@AccountNo", Me.Account)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Select012] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    Public Function Select013() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "SELECT013")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@UserPIC", Me.UserPIC)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Select013] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    Public Function Select014() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "SELECT014")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Select014] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    Public Function Select015() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "SELECT015")
            strSQL = strSQL.Replace("@UserPIC", Me.PersonInCharge)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            If dt.Rows.Count > 0 Then
                Me.PICShowFlag = CStr(Nz(dt.Rows(0)!PIC_SHOW_FLAG, "1"))
            Else
                Me.PICShowFlag = "1"
            End If

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Select015] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    Public Function Select016() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String
        Dim rtn As Boolean = False

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "SELECT016")
            strSQL = strSQL.Replace("@OrderNo", Me.BudgetOrderNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            If dt.Rows.Count > 0 Then
                Me.dtResult = dt
                rtn = True
            Else
                Me.dtResult = Nothing
                rtn = False
            End If

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return rtn

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Select016] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    Public Function Select017() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "SELECT017")
            strSQL = strSQL.Replace("@P_BG_ORDER_NO", Me.BudgetOrderNo.Trim)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Select017] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    Public Function Update001(Optional ByVal pConn As SqlConnection = Nothing, _
                              Optional ByVal pTrans As SqlTransaction = Nothing) As Boolean
        Dim conn As SqlConnection = Nothing
        Dim cmd As SqlCommand
        Dim strSQL As String
        Dim intRtn As Integer

        Try
            If pConn Is Nothing Or pTrans Is Nothing Then
                conn = New SqlConnection(My.Settings.ConnStr)
                conn.Open()
            End If

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "UPDATE001")
            strSQL = strSQL.Replace("@P_BG_ORDER_NO", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@P_BG_ORDER_NAME", Me.BudgetOrderName.Replace("'", "''"))
            strSQL = strSQL.Replace("@P_BG_TYPE", Me.BudgetType)
            strSQL = strSQL.Replace("@P_ACCOUNT_NO", Me.Account)
            strSQL = strSQL.Replace("@P_COST_CENTER", Me.CostCenter)
            If Me.CostType = "" Then
                strSQL = strSQL.Replace("@P_COST_TYPE", "NULL")
            Else
                strSQL = strSQL.Replace("@P_COST_TYPE", Me.CostType)
            End If
            If Me.Cost = "" Then
                strSQL = strSQL.Replace("@P_COST", "NULL")
            Else
                strSQL = strSQL.Replace("@P_COST", Me.Cost)
            End If
            If Me.AssetGroup = "" Then
                strSQL = strSQL.Replace("@P_ASSET_GROUP_NO", "NULL")
            Else
                strSQL = strSQL.Replace("@P_ASSET_GROUP_NO", "'" & Me.AssetGroup & "'")
            End If
            strSQL = strSQL.Replace("@P_DEPT_NO", Me.Department)
            strSQL = strSQL.Replace("@P_PERSON_IN_CHARGE", Me.PersonInCharge)
            strSQL = strSQL.Replace("@P_ACTIVE_FLAG", Me.ActiveFlag)
            strSQL = strSQL.Replace("@P_EXPENSE_TYPE", Me.ExpenseType)
            strSQL = strSQL.Replace("@P_PIC_SHOW_FLAG", Me.PICShowFlag)
            strSQL = strSQL.Replace("@P_UPDATE_USER_ID", Me.UpdateUserId)
            strSQL = strSQL.Replace("@P_REMARKS", Me.Remarks)

            If pConn IsNot Nothing And pTrans IsNot Nothing Then
                cmd = New SqlCommand(strSQL, pConn, pTrans)
                intRtn = cmd.ExecuteNonQuery()

            Else
                cmd = New SqlCommand(strSQL, conn)
                intRtn = cmd.ExecuteNonQuery()

                If conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If

            End If

            If intRtn > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Update001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function

    Public Function Update002(Optional ByVal pConn As SqlConnection = Nothing, _
                              Optional ByVal pTrans As SqlTransaction = Nothing) As Boolean
        Dim conn As SqlConnection = Nothing
        Dim cmd As SqlCommand
        Dim strSQL As String
        Dim intRtn As Integer

        Try
            If pConn Is Nothing Or pTrans Is Nothing Then
                conn = New SqlConnection(My.Settings.ConnStr)
                conn.Open()
            End If

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "UPDATE002")
            strSQL = strSQL.Replace("@P_PERSON_IN_CHARGE", Me.PersonInCharge)
            strSQL = strSQL.Replace("@P_UPDATE_USER_ID", Me.UpdateUserId)
            strSQL = strSQL.Replace("@P_PIC_SHOW_FLAG", Me.PICShowFlag)

            If pConn IsNot Nothing And pTrans IsNot Nothing Then
                cmd = New SqlCommand(strSQL, pConn, pTrans)
                intRtn = cmd.ExecuteNonQuery()

            Else
                cmd = New SqlCommand(strSQL, conn)
                intRtn = cmd.ExecuteNonQuery()

                If conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If

            End If

            If intRtn > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Update002] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function

    Public Function Insert001(Optional ByVal pConn As SqlConnection = Nothing, _
                              Optional ByVal pTrans As SqlTransaction = Nothing) As Boolean
        Dim conn As SqlConnection = Nothing
        Dim cmd As SqlCommand
        Dim strSQL As String
        Dim intRtn As Integer

        Try
            If pConn Is Nothing Or pTrans Is Nothing Then
                conn = New SqlConnection(My.Settings.ConnStr)
                conn.Open()
            End If

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "INSERT001")
            strSQL = strSQL.Replace("@P_BG_ORDER_NO", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@P_BG_ORDER_NAME", Me.BudgetOrderName.Replace("'", "''"))
            strSQL = strSQL.Replace("@P_BG_TYPE", Me.BudgetType)
            strSQL = strSQL.Replace("@P_ACCOUNT_NO", Me.Account)
            strSQL = strSQL.Replace("@P_COST_CENTER", Me.CostCenter)
            If Me.CostType = "" Then
                strSQL = strSQL.Replace("@P_COST_TYPE", "NULL")
            Else
                strSQL = strSQL.Replace("@P_COST_TYPE", Me.CostType)
            End If
            If Me.Cost = "" Then
                strSQL = strSQL.Replace("@P_COST", "NULL")
            Else
                strSQL = strSQL.Replace("@P_COST", Me.Cost)
            End If
            If Me.AssetGroup = "" Then
                strSQL = strSQL.Replace("@P_ASSET_GROUP_NO", "NULL")
            Else
                strSQL = strSQL.Replace("@P_ASSET_GROUP_NO", "'" & Me.AssetGroup & "'")
            End If
            strSQL = strSQL.Replace("@P_DEPT_NO", Me.Department)
            strSQL = strSQL.Replace("@P_PERSON_IN_CHARGE", Me.PersonInCharge)
            strSQL = strSQL.Replace("@P_ACTIVE_FLAG", Me.ActiveFlag)
            strSQL = strSQL.Replace("@P_EXPENSE_TYPE", Me.ExpenseType)
            strSQL = strSQL.Replace("@P_PIC_SHOW_FLAG", Me.PICShowFlag)
            strSQL = strSQL.Replace("@P_CREATE_USER_ID", Me.CreateUserId)
            strSQL = strSQL.Replace("@P_REMARKS", Me.Remarks)

            If pConn IsNot Nothing And pTrans IsNot Nothing Then
                cmd = New SqlCommand(strSQL, pConn, pTrans)
                intRtn = cmd.ExecuteNonQuery()

            Else
                cmd = New SqlCommand(strSQL, conn)
                intRtn = cmd.ExecuteNonQuery()

                If conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If

            End If

            If intRtn > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Insert001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function

    Public Function Delete001(Optional ByVal pConn As SqlConnection = Nothing, _
                              Optional ByVal pTrans As SqlTransaction = Nothing) As Boolean
        Dim conn As SqlConnection = Nothing
        Dim cmd As SqlCommand
        Dim strSQL As String
        Dim intRtn As Integer

        Try
            If pConn Is Nothing Or pTrans Is Nothing Then
                conn = New SqlConnection(My.Settings.ConnStr)
                conn.Open()
            End If

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_BUDGET_ORDER", "DELETE001")
            strSQL = strSQL.Replace("@P_BG_ORDER_NO", Me.BudgetOrderNo.Trim)

            If pConn IsNot Nothing And pTrans IsNot Nothing Then
                cmd = New SqlCommand(strSQL, pConn, pTrans)
                intRtn = cmd.ExecuteNonQuery()
            Else
                cmd = New SqlCommand(strSQL, conn)
                intRtn = cmd.ExecuteNonQuery()

                If conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If

            End If

            If intRtn > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show("[BG_M_BUDGET_ORDER.Delete001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False
        End Try
    End Function

#End Region

End Class
