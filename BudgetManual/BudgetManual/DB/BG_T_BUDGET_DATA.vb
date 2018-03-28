Imports System.Data.SqlClient
Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class BG_T_BUDGET_DATA

#Region "Variable"
    Private Const STRING_ALL As String = "All"
    Private myDtResult As DataTable
    Private myDS As DataSet = Nothing
    Private myUserPIC As String = String.Empty
    Private myUserId As String = String.Empty
    Private myBudgetYear As String = String.Empty
    Private myPeriodType As String = String.Empty
    Private myMTPChecked As Boolean = False
    Private myBudgetType As String = String.Empty
    Private myProjectNo As String = String.Empty
    Private myPrevProjectNo As String = String.Empty
    Private myBudgetOrderNo As String = String.Empty
    Private myRevNo As String = String.Empty
    Private myBudgetData As DataRow
    Private myBudgetData2 As Double()
    Private myAccountNo As String = String.Empty
    Private myTableName As String = String.Empty
    Private myM(12) As String
    Private myRRT(5) As String
    Private myWorkingBG(2) As String
    Private myRemarks As String = String.Empty
    Private myRevNo2 As String = String.Empty
    Private myStatus As String = String.Empty
    Private myDataList As ArrayList
    Private myMTPBudget As Boolean = False
    Private myReviseRevNo As String = String.Empty
    Private myPrevMTPRevNo As String = String.Empty
    Private myMtpProjectNo As String = String.Empty
    Private myMtpRevNo As String = String.Empty
    Private myRefBudgetYear As String = String.Empty
    Private myRefPeriodType As String = String.Empty
    Private myRefProjectNo As String = String.Empty
    Private myRefRevNo As String = String.Empty
    Private myRefEstProjectNo As String = String.Empty
    Private myRefEstRevNo As String = String.Empty
    Private myRefRBProjectNo As String = String.Empty
    Private myRefRBRevNo As String = String.Empty

#End Region

#Region "Property"

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

#Region "UserId"
    Property UserId() As String
        Get
            Return myUserId
        End Get
        Set(ByVal value As String)
            myUserId = value
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

#Region "MTPChecked"
    Property MTPChecked() As Boolean
        Get
            Return myMTPChecked
        End Get
        Set(ByVal value As Boolean)
            myMTPChecked = value
        End Set
    End Property
#End Region

#Region "BudgetType"
    Property BudgetType() As String
        Get
            Return myBudgetType
        End Get
        Set(ByVal value As String)
            myBudgetType = value
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

#Region "MtpProjectNo"
    Property MtpProjectNo() As String
        Get
            Return myMtpProjectNo
        End Get
        Set(ByVal value As String)
            myMtpProjectNo = value
        End Set
    End Property
#End Region

#Region "MtpRevNo"
    Property MtpRevNo() As String
        Get
            Return myMtpRevNo
        End Get
        Set(ByVal value As String)
            myMtpRevNo = value
        End Set
    End Property
#End Region

#Region "PrevProjectNo"
    Property PrevProjectNo() As String
        Get
            Return myPrevProjectNo
        End Get
        Set(ByVal value As String)
            myPrevProjectNo = value
        End Set
    End Property
#End Region

#Region "BudgetOrderNo"
    Property BudgetOrderNo() As String
        Get
            Return myBudgetOrderNo
        End Get
        Set(ByVal value As String)
            myBudgetOrderNo = value
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

#Region "ReviseRevNo"
    Property ReviseRevNo() As String
        Get
            Return myReviseRevNo
        End Get
        Set(ByVal value As String)
            myReviseRevNo = value
        End Set
    End Property
#End Region

#Region "PrevMTPRevNo"
    Property PrevMTPRevNo() As String
        Get
            Return myPrevMTPRevNo
        End Get
        Set(ByVal value As String)
            myPrevMTPRevNo = value
        End Set
    End Property
#End Region

#Region "BudgetData"
    Property BudgetData() As DataRow
        Get
            Return myBudgetData
        End Get
        Set(ByVal value As DataRow)
            myBudgetData = value
        End Set
    End Property
#End Region

#Region "BudgetData2"
    Property BudgetData2() As Double()
        Get
            Return myBudgetData2
        End Get
        Set(ByVal value As Double())
            myBudgetData2 = value
        End Set
    End Property
#End Region

#Region " AccountNo "
    Public Property AccountNo() As String
        Get
            Return myAccountNo
        End Get
        Set(ByVal value As String)
            myAccountNo = value
        End Set
    End Property
#End Region

#Region " TableName "
    Public Property TableName() As String
        Get
            Return myTableName
        End Get
        Set(ByVal value As String)
            myTableName = value
        End Set
    End Property
#End Region

#Region "M"
    Property M() As String()
        Get
            Return myM
        End Get
        Set(ByVal value As String())
            myM = value
        End Set
    End Property
#End Region

#Region "RRT"
    Property RRT() As String()
        Get
            Return myRRT
        End Get
        Set(ByVal value As String())
            myRRT = value
        End Set
    End Property
#End Region

#Region "WorkingBG"
    Property WorkingBG() As String()
        Get
            Return myWorkingBG
        End Get
        Set(ByVal value As String())
            myWorkingBG = value
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

#Region "RevNo2"
    Property RevNo2() As String
        Get
            Return myRevNo2
        End Get
        Set(ByVal value As String)
            myRevNo2 = value
        End Set
    End Property
#End Region

#Region "Status"
    Property Status() As String
        Get
            Return myStatus
        End Get
        Set(ByVal value As String)
            myStatus = value
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

#Region "MTPBudget"
    Public Property MTPBudget() As Boolean
        Get
            Return myMTPBudget
        End Get
        Set(ByVal value As Boolean)
            myMTPBudget = value
        End Set
    End Property
#End Region

#Region "RefBudgetYear"
    Property RefBudgetYear() As String
        Get
            Return myRefBudgetYear
        End Get
        Set(ByVal value As String)
            myRefBudgetYear = value
        End Set
    End Property
#End Region

#Region "RefPeriodType"
    Property RefPeriodType() As String
        Get
            Return myRefPeriodType
        End Get
        Set(ByVal value As String)
            myRefPeriodType = value
        End Set
    End Property
#End Region

#Region "RefProjectNo"
    Property RefProjectNo() As String
        Get
            Return myRefProjectNo
        End Get
        Set(ByVal value As String)
            myRefProjectNo = value
        End Set
    End Property
#End Region

#Region "RefRevNo"
    Property RefRevNo() As String
        Get
            Return myRefRevNo
        End Get
        Set(ByVal value As String)
            myRefRevNo = value
        End Set
    End Property
#End Region

#Region "RefEstProjectNo"
    Property RefEstProjectNo() As String
        Get
            Return myRefEstProjectNo
        End Get
        Set(ByVal value As String)
            myRefEstProjectNo = value
        End Set
    End Property
#End Region

#Region "RefEstRevNo"
    Property RefEstRevNo() As String
        Get
            Return myRefEstRevNo
        End Get
        Set(ByVal value As String)
            myRefEstRevNo = value
        End Set
    End Property
#End Region

#Region "RefRBProjectNo"
    Property RefRBProjectNo() As String
        Get
            Return myRefRBProjectNo
        End Get
        Set(ByVal value As String)
            myRefRBProjectNo = value
        End Set
    End Property
#End Region

#Region "RefRBRevNo"
    Property RefRBRevNo() As String
        Get
            Return myRefRBRevNo
        End Get
        Set(ByVal value As String)
            myRefRBRevNo = value
        End Set
    End Property
#End Region

#End Region

#Region "Function"

#Region "Select001"
    ''' <summary>
    ''' Select budget data by PIC
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select001() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT001")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@UserPIC", Me.UserPIC)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            'strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            'strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefBudgetYear", Me.RefBudgetYear)
            strSQL = strSQL.Replace("@RefPeriodType", Me.RefPeriodType)
            strSQL = strSQL.Replace("@RefProjectNo", Me.RefProjectNo)
            strSQL = strSQL.Replace("@RefRevNo", Me.RefRevNo)

            strSQL = strSQL.Replace("@RefEstProjectNo", Me.RefEstProjectNo)
            strSQL = strSQL.Replace("@RefEstRevNo", Me.RefEstRevNo)


            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select001_1"
    ''' <summary>
    ''' Select budget data by PIC
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select001_1() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT001_1")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@UserPIC", Me.UserPIC)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            'strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            'strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefBudgetYear", Me.RefBudgetYear)
            strSQL = strSQL.Replace("@RefPeriodType", Me.RefPeriodType)
            strSQL = strSQL.Replace("@RefProjectNo", Me.RefProjectNo)
            strSQL = strSQL.Replace("@RefRevNo", Me.RefRevNo)

            strSQL = strSQL.Replace("@RefEstProjectNo", Me.RefEstProjectNo)
            strSQL = strSQL.Replace("@RefEstRevNo", Me.RefEstRevNo)


            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select002"
    ''' <summary>
    ''' Select all budget data
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select002() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT002")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            'strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            'strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefBudgetYear", Me.RefBudgetYear)
            strSQL = strSQL.Replace("@RefPeriodType", Me.RefPeriodType)
            strSQL = strSQL.Replace("@RefProjectNo", Me.RefProjectNo)
            strSQL = strSQL.Replace("@RefRevNo", Me.RefRevNo)

            strSQL = strSQL.Replace("@RefEstProjectNo", Me.RefEstProjectNo)
            strSQL = strSQL.Replace("@RefEstRevNo", Me.RefEstRevNo)


            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select002] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select003"
    ''' <summary>
    ''' Select Max Rev No.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select003() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT003")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@ProjectNo", Me.PeriodType)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select003] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select004"

    ''' <summary>
    ''' Select Budget data list for Report001-1
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select004_1() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT004_1")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECTNO", Me.ProjectNo)

            strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefEstProjectNo", Me.RefEstProjectNo)
            strSQL = strSQL.Replace("@RefEstRevNo", Me.RefEstRevNo)

            If Me.UserPIC = "0" Then
                strSQL = strSQL.Replace("@PICCONDITION", " ")
            Else
                strSQL = strSQL.Replace("@PICCONDITION", "AND BUDGET_ORDER.PERSON_IN_CHARGE_NO = '@PIC'")
                strSQL = strSQL.Replace("@PIC", Me.UserPIC)
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
            MessageBox.Show("[BG_T_BUDGET_DATA.Select004] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' Select Budget data list for Report001-2
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select004_2() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT004_2")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECTNO", Me.ProjectNo)

            If Me.UserPIC = "0" Then
                strSQL = strSQL.Replace("@PICCONDITION", " ")
            Else
                strSQL = strSQL.Replace("@PICCONDITION", "AND BUDGET_ORDER.PERSON_IN_CHARGE_NO = '@PIC'")
                strSQL = strSQL.Replace("@PIC", Me.UserPIC)
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
            MessageBox.Show("[BG_T_BUDGET_DATA.Select004] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' Select Budget data list for Report001-3
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select004_3() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT004_3")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECTNO", Me.ProjectNo)

            If Me.UserPIC = "0" Then
                strSQL = strSQL.Replace("@PICCONDITION", " ")
            Else
                strSQL = strSQL.Replace("@PICCONDITION", "AND BUDGET_ORDER.PERSON_IN_CHARGE_NO = '@PIC'")
                strSQL = strSQL.Replace("@PIC", Me.UserPIC)
            End If

            'If Me.MTPChecked = True Then
            '    Dim strMTPSql As String = " OR ISNULL(MASTER_DATA.RRT1, 0)<> 0 OR ISNULL(MASTER_DATA.RRT2, 0) <> 0 OR ISNULL(MASTER_DATA.RRT3, 0) <> 0 OR ISNULL(MASTER_DATA.RRT4, 0) <> 0 OR ISNULL(MASTER_DATA.RRT5, 0) <> 0"
            '    Dim intLen As Integer = strSQL.LastIndexOf(")")
            '    strSQL = strSQL.Insert(intLen, strMTPSql)
            'End If

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select004] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' Select Budget data list for Report001-5
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select004_6() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT004_6")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECTNO", Me.ProjectNo)
            'strSQL = strSQL.Replace("@PrevProjectNo", Me.PrevProjectNo)
            strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            If Me.UserPIC = "0" OrElse Me.UserPIC.Equals(String.Empty) Then
                strSQL = strSQL.Replace("@PICCONDITION", " ")
            Else
                strSQL = strSQL.Replace("@PICCONDITION", "AND BO.PERSON_IN_CHARGE_NO = '@PIC'")
                strSQL = strSQL.Replace("@PIC", Me.UserPIC)
            End If

            strSQL = strSQL.Replace("@RefRBProjectNo", Me.RefRBProjectNo)
            strSQL = strSQL.Replace("@RefRBRevNo", Me.RefRBRevNo)

            'If Me.MTPChecked = True Then
            '    Dim strMTPSql As String = " OR ISNULL(MASTER_DATA.RRT1, 0)<> 0 OR ISNULL(MASTER_DATA.RRT2, 0) <> 0 OR ISNULL(MASTER_DATA.RRT3, 0) <> 0 OR ISNULL(MASTER_DATA.RRT4, 0) <> 0 OR ISNULL(MASTER_DATA.RRT5, 0) <> 0"
            '    Dim intLen As Integer = strSQL.LastIndexOf(")")
            '    strSQL = strSQL.Insert(intLen, strMTPSql)
            'End If

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select004_6] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' Select Budget data list for Report002-5
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select004_7() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT004_7")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECTNO", Me.ProjectNo)
            'strSQL = strSQL.Replace("@PrevProjectNo", Me.PrevProjectNo)

            strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefRBProjectNo", Me.RefRBProjectNo)
            strSQL = strSQL.Replace("@RefRBRevNo", Me.RefRBRevNo)

            If Me.UserPIC = "0" OrElse Me.UserPIC.Equals(String.Empty) Then
                strSQL = strSQL.Replace("@PICCONDITION", " ")
            Else
                strSQL = strSQL.Replace("@PICCONDITION", "AND BO.PERSON_IN_CHARGE_NO = '@PIC'")
                strSQL = strSQL.Replace("@PIC", Me.UserPIC)
            End If

            'If Me.MTPChecked = True Then
            '    Dim strMTPSql As String = " OR ISNULL(MASTER_DATA.RRT1, 0)<> 0 OR ISNULL(MASTER_DATA.RRT2, 0) <> 0 OR ISNULL(MASTER_DATA.RRT3, 0) <> 0 OR ISNULL(MASTER_DATA.RRT4, 0) <> 0 OR ISNULL(MASTER_DATA.RRT5, 0) <> 0"
            '    Dim intLen As Integer = strSQL.LastIndexOf(")")
            '    strSQL = strSQL.Insert(intLen, strMTPSql)
            'End If

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select004_7] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' Select Budget data list for Report002-5
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select004_8() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT004_8")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECTNO", Me.ProjectNo)
            strSQL = strSQL.Replace("@REVNO", Me.RevNo)
            'strSQL = strSQL.Replace("@PrevProjectNo", Me.PrevProjectNo)
            'strSQL = strSQL.Replace("@PrevRevNo", Me.PrevMTPRevNo)
            strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefRBProjectNo", Me.RefRBProjectNo)
            strSQL = strSQL.Replace("@RefRBRevNo", Me.RefRBRevNo)


            If Me.UserPIC = "0" OrElse Me.UserPIC.Equals(String.Empty) Then
                strSQL = strSQL.Replace("@PICCONDITION", " ")
            Else
                strSQL = strSQL.Replace("@PICCONDITION", "AND BO.PERSON_IN_CHARGE_NO = '@PIC'")
                strSQL = strSQL.Replace("@PIC", Me.UserPIC)
            End If

            'If Me.MTPChecked = True Then
            '    Dim strMTPSql As String = " OR ISNULL(MASTER_DATA.RRT1, 0)<> 0 OR ISNULL(MASTER_DATA.RRT2, 0) <> 0 OR ISNULL(MASTER_DATA.RRT3, 0) <> 0 OR ISNULL(MASTER_DATA.RRT4, 0) <> 0 OR ISNULL(MASTER_DATA.RRT5, 0) <> 0"
            '    Dim intLen As Integer = strSQL.LastIndexOf(")")
            '    strSQL = strSQL.Insert(intLen, strMTPSql)
            'End If

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select004_8] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' Select Budget data list for Report001-1
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select004_9() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT004_9")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECTNO", Me.ProjectNo)
            strSQL = strSQL.Replace("@REVNO", Me.RevNo)

            strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefEstProjectNo", Me.RefEstProjectNo)
            strSQL = strSQL.Replace("@RefEstRevNo", Me.RefEstRevNo)

            If Me.UserPIC = "0" Then
                strSQL = strSQL.Replace("@PICCONDITION", " ")
            Else
                strSQL = strSQL.Replace("@PICCONDITION", "AND BUDGET_ORDER.PERSON_IN_CHARGE_NO = '@PIC'")
                strSQL = strSQL.Replace("@PIC", Me.UserPIC)
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
            MessageBox.Show("[BG_T_BUDGET_DATA.Select004] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' Select Budget data list for Report001-2
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select004_10() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT004_10")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECTNO", Me.ProjectNo)
            strSQL = strSQL.Replace("@REVNO", Me.RevNo)

            If Me.UserPIC = "0" Then
                strSQL = strSQL.Replace("@PICCONDITION", " ")
            Else
                strSQL = strSQL.Replace("@PICCONDITION", "AND BUDGET_ORDER.PERSON_IN_CHARGE_NO = '@PIC'")
                strSQL = strSQL.Replace("@PIC", Me.UserPIC)
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
            MessageBox.Show("[BG_T_BUDGET_DATA.Select004] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' Select Budget data list for Report001-3
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select004_11() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT004_11")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECTNO", Me.ProjectNo)
            strSQL = strSQL.Replace("@REVNO", Me.RevNo)

            If Me.UserPIC = "0" Then
                strSQL = strSQL.Replace("@PICCONDITION", " ")
            Else
                strSQL = strSQL.Replace("@PICCONDITION", "AND BUDGET_ORDER.PERSON_IN_CHARGE_NO = '@PIC'")
                strSQL = strSQL.Replace("@PIC", Me.UserPIC)
            End If

            'If Me.MTPChecked = True Then
            '    Dim strMTPSql As String = " OR ISNULL(MASTER_DATA.RRT1, 0)<> 0 OR ISNULL(MASTER_DATA.RRT2, 0) <> 0 OR ISNULL(MASTER_DATA.RRT3, 0) <> 0 OR ISNULL(MASTER_DATA.RRT4, 0) <> 0 OR ISNULL(MASTER_DATA.RRT5, 0) <> 0"
            '    Dim intLen As Integer = strSQL.LastIndexOf(")")
            '    strSQL = strSQL.Insert(intLen, strMTPSql)
            'End If

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select004] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' Select Budget data list for Report001-5
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select004_12() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT004_12")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECTNO", Me.ProjectNo)
            strSQL = strSQL.Replace("@REVNO", Me.RevNo)
            strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            If Me.UserPIC = "0" OrElse Me.UserPIC.Equals(String.Empty) Then
                strSQL = strSQL.Replace("@PICCONDITION", " ")
            Else
                strSQL = strSQL.Replace("@PICCONDITION", "AND BO.PERSON_IN_CHARGE_NO = '@PIC'")
                strSQL = strSQL.Replace("@PIC", Me.UserPIC)
            End If

            strSQL = strSQL.Replace("@RefRBProjectNo", Me.RefRBProjectNo)
            strSQL = strSQL.Replace("@RefRBRevNo", Me.RefRBRevNo)

            'If Me.MTPChecked = True Then
            '    Dim strMTPSql As String = " OR ISNULL(MASTER_DATA.RRT1, 0)<> 0 OR ISNULL(MASTER_DATA.RRT2, 0) <> 0 OR ISNULL(MASTER_DATA.RRT3, 0) <> 0 OR ISNULL(MASTER_DATA.RRT4, 0) <> 0 OR ISNULL(MASTER_DATA.RRT5, 0) <> 0"
            '    Dim intLen As Integer = strSQL.LastIndexOf(")")
            '    strSQL = strSQL.Insert(intLen, strMTPSql)
            'End If

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select004_12] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

#End Region

#Region "Select005"
    ''' <summary>
    ''' Select Budget data list for Report 003
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select005() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim strSQL As String
        Dim strDataTableName As String = String.Empty
        Dim strValueName As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            Dim strAccountNo As String = String.Empty 

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT005_1")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            'strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)

            strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefEstProjectNo", Me.RefEstProjectNo)
            strSQL = strSQL.Replace("@RefEstRevNo", Me.RefEstRevNo)

            If String.Equals(Me.AccountNo.Trim.ToUpper, STRING_ALL.ToUpper) Then
                strSQL = strSQL.Replace("@AccountNo", " ")
            Else
                strAccountNo = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT005_2")
                strAccountNo = strAccountNo.Replace("@AccountNo", Me.AccountNo)
                strSQL = strSQL.Replace("@AccountNo", strAccountNo)
            End If

            da = New SqlDataAdapter(strSQL, conn)
            ds = New DataSet

            da.Fill(ds, TableName)

            Me.DS = ds

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select005] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select005_3"
    ''' <summary>
    ''' Select Budget data list for Report 003
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select005_3() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim strSQL As String
        Dim strDataTableName As String = String.Empty
        Dim strValueName As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            Dim strAccountNo As String = String.Empty

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT005_3")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)

            strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefEstProjectNo", Me.RefEstProjectNo)
            strSQL = strSQL.Replace("@RefEstRevNo", Me.RefEstRevNo)

            If String.Equals(Me.AccountNo.Trim.ToUpper, STRING_ALL.ToUpper) Then
                strSQL = strSQL.Replace("@AccountNo", " ")
            Else
                strAccountNo = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT005_2")
                strAccountNo = strAccountNo.Replace("@AccountNo", Me.AccountNo)
                strSQL = strSQL.Replace("@AccountNo", strAccountNo)
            End If

            da = New SqlDataAdapter(strSQL, conn)
            ds = New DataSet

            da.Fill(ds, TableName)

            Me.DS = ds

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select005_3] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select006"
    ''' <summary>
    ''' Select budget data to export to SAP
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select006() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String = ""

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            If Me.AccountNo <> "" Then
                If Me.PeriodType = CStr(enumPeriodType.OriginalBudget) Then
                    strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT006_4")
                ElseIf Me.PeriodType = CStr(enumPeriodType.ReviseBudget) Then
                    strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT006_1")
                ElseIf Me.PeriodType = CStr(enumPeriodType.EstimateBudget) Then
                    strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT006_3")
                ElseIf Me.PeriodType = CStr(enumPeriodType.MTPBudget) Then
                    strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT006_8")
                End If

                strSQL = strSQL.Replace("@P_BUDGET_YEAR", Me.BudgetYear)
                strSQL = strSQL.Replace("@P_PERIOD_TYPE", Me.PeriodType)
                strSQL = strSQL.Replace("@P_ACCOUNT_NO", Me.AccountNo)
                strSQL = strSQL.Replace("@P_PROJECT_NO", Me.ProjectNo)
            Else
                If Me.PeriodType = CStr(enumPeriodType.OriginalBudget) Then
                    strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT006_6")
                ElseIf Me.PeriodType = CStr(enumPeriodType.ReviseBudget) Then
                    strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT006_2")
                ElseIf Me.PeriodType = CStr(enumPeriodType.EstimateBudget) Then
                    strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT006_5")
                ElseIf Me.PeriodType = CStr(enumPeriodType.MTPBudget) Then
                    strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT006_7")
                End If

                strSQL = strSQL.Replace("@P_BUDGET_YEAR", Me.BudgetYear)
                strSQL = strSQL.Replace("@P_PERIOD_TYPE", Me.PeriodType)
                strSQL = strSQL.Replace("@P_PROJECT_NO", Me.ProjectNo)
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
            MessageBox.Show("[BG_T_BUDGET_DATA.Select006] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

#End Region

#Region "Select007"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select007() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim strSQL As String
        Dim strDataTableName As String = String.Empty
        Dim strValueName As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            Dim strAccountNo As String = String.Empty

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT007_1")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@ReviseBudget", CStr(enumPeriodType.ReviseBudget))
            strSQL = strSQL.Replace("@ActualData", CStr(enumUploadDataType.ActualData))
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            'If Not String.Equals(Me.AccountNo.Trim.ToUpper, STRING_ALL.ToUpper) Then
            '    strAccountNo = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT007_2")
            '    strSQL += strAccountNo
            '    strSQL = strSQL.Replace("@AccountNo", Me.AccountNo)
            'End If
            If String.Equals(Me.AccountNo.Trim.ToUpper, STRING_ALL.ToUpper) Then
                strSQL = strSQL.Replace("@AccountNo", " ")
            Else
                strAccountNo = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT007_2")
                strAccountNo = strAccountNo.Replace("@AccountNo", Me.AccountNo)
                strSQL = strSQL.Replace("@AccountNo", strAccountNo)
            End If

            da = New SqlDataAdapter(strSQL, conn)
            ds = New DataSet

            da.Fill(ds, TableName)

            Me.DS = ds

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select007] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select007_3"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select007_3() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim strSQL As String
        Dim strDataTableName As String = String.Empty
        Dim strValueName As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            Dim strAccountNo As String = String.Empty

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT007_3")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@ReviseBudget", CStr(enumPeriodType.ReviseBudget))
            strSQL = strSQL.Replace("@ActualData", CStr(enumUploadDataType.ActualData))
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)

            If String.Equals(Me.AccountNo.Trim.ToUpper, STRING_ALL.ToUpper) Then
                strSQL = strSQL.Replace("@AccountNo", " ")
            Else
                strAccountNo = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT007_2")
                strAccountNo = strAccountNo.Replace("@AccountNo", Me.AccountNo)
                strSQL = strSQL.Replace("@AccountNo", strAccountNo)
            End If

            da = New SqlDataAdapter(strSQL, conn)
            ds = New DataSet

            da.Fill(ds, TableName)

            Me.DS = ds

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select007_3] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select008"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select008() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim strSQL As String
        Dim strDataTableName As String = String.Empty
        'Dim strValueName As String = String.Empty
        Dim strAccountNo As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            'If Not Me.MTPBudget Then
            '    strValueName = "SELECT008_1"
            'Else
            '    strValueName = "SELECT008_1_2"
            'End If

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT008_1")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@OriginalBudget", CStr(enumPeriodType.OriginalBudget))
            strSQL = strSQL.Replace("@EstimateBudget", CStr(enumPeriodType.EstimateBudget))
            'strSQL = strSQL.Replace("@ReviseBudget", CStr(enumPeriodType.ReviseBudget))
            strSQL = strSQL.Replace("@ActualData", CStr(enumUploadDataType.ActualData))
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            If String.Equals(Me.AccountNo.Trim.ToUpper, STRING_ALL.ToUpper) Then
                strSQL = strSQL.Replace("@AccountNo", " ")
            Else
                strAccountNo = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT008_2")
                strAccountNo = strAccountNo.Replace("@AccountNo", Me.AccountNo)
                strSQL = strSQL.Replace("@AccountNo", strAccountNo)
            End If

            da = New SqlDataAdapter(strSQL, conn)
            ds = New DataSet

            da.Fill(ds, TableName)

            Me.DS = ds

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select008] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select008_3"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select008_3() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim strSQL As String
        Dim strDataTableName As String = String.Empty
        'Dim strValueName As String = String.Empty
        Dim strAccountNo As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            'If Not Me.MTPBudget Then
            '    strValueName = "SELECT008_3_1"
            'Else
            '    strValueName = "SELECT008_3_1_2"
            'End If

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT008_3")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@OriginalBudget", CStr(enumPeriodType.OriginalBudget))
            strSQL = strSQL.Replace("@EstimateBudget", CStr(enumPeriodType.EstimateBudget))
            'strSQL = strSQL.Replace("@ReviseBudget", CStr(enumPeriodType.ReviseBudget))
            strSQL = strSQL.Replace("@ActualData", CStr(enumUploadDataType.ActualData))
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)

            If String.Equals(Me.AccountNo.Trim.ToUpper, STRING_ALL.ToUpper) Then
                strSQL = strSQL.Replace("@AccountNo", " ")
            Else
                strAccountNo = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT008_2")
                strAccountNo = strAccountNo.Replace("@AccountNo", Me.AccountNo)
                strSQL = strSQL.Replace("@AccountNo", strAccountNo)
            End If

            da = New SqlDataAdapter(strSQL, conn)
            ds = New DataSet

            da.Fill(ds, TableName)

            Me.DS = ds

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select008_3] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select009 "
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select009() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim strSQL As String
        Dim strDataTableName As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            Dim strAccountNo As String = String.Empty

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT009")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            'strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)

            strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefEstProjectNo", Me.RefEstProjectNo)
            strSQL = strSQL.Replace("@RefEstRevNo", Me.RefEstRevNo)

            da = New SqlDataAdapter(strSQL, conn)
            ds = New DataSet

            da.Fill(ds, TableName)

            Me.DS = ds

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select009] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select009_2 "
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select009_2() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim strSQL As String
        Dim strDataTableName As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            Dim strAccountNo As String = String.Empty

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT009_2")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)

            strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefEstProjectNo", Me.RefEstProjectNo)
            strSQL = strSQL.Replace("@RefEstRevNo", Me.RefEstRevNo)

            da = New SqlDataAdapter(strSQL, conn)
            ds = New DataSet

            da.Fill(ds, TableName)

            Me.DS = ds

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select009_2] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select010 "
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select010() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim strSQL As String
        Dim strDataTableName As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            Dim strAccountNo As String = String.Empty

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT010")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@ReviseBudget", CStr(enumPeriodType.ReviseBudget))
            strSQL = strSQL.Replace("@ActualData", CStr(enumUploadDataType.ActualData))
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            da = New SqlDataAdapter(strSQL, conn)
            ds = New DataSet

            da.Fill(ds, TableName)

            Me.DS = ds

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select010] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region
#Region "Select010_2 "
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select010_2() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim strSQL As String
        Dim strDataTableName As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            Dim strAccountNo As String = String.Empty

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT010_2")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@ReviseBudget", CStr(enumPeriodType.ReviseBudget))
            strSQL = strSQL.Replace("@ActualData", CStr(enumUploadDataType.ActualData))
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)

            da = New SqlDataAdapter(strSQL, conn)
            ds = New DataSet

            da.Fill(ds, TableName)

            Me.DS = ds

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select010_2] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select011 "
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select011() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim strSQL As String
        Dim strDataTableName As String = String.Empty
        'Dim strValueName As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            Dim strAccountNo As String = String.Empty

            'If Not Me.MTPBudget Then
            '    strValueName = "SELECT011"
            'Else
            '    strValueName = "SELECT011_1"
            'End If

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT011")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@OriginalBudget", CStr(enumPeriodType.OriginalBudget))
            strSQL = strSQL.Replace("@EstimateBudget", CStr(enumPeriodType.EstimateBudget))
            strSQL = strSQL.Replace("@ActualData", CStr(enumUploadDataType.ActualData))
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            da = New SqlDataAdapter(strSQL, conn)
            ds = New DataSet

            da.Fill(ds, TableName)

            Me.DS = ds

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select011] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

#End Region

#Region "Select011_2 "
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select011_2() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim strSQL As String
        Dim strDataTableName As String = String.Empty
        'Dim strValueName As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            Dim strAccountNo As String = String.Empty

            'If Not Me.MTPBudget Then
            '    strValueName = "SELECT011_2"
            'Else
            '    strValueName = "SELECT011_2_1"
            'End If

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT011_2")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@OriginalBudget", CStr(enumPeriodType.OriginalBudget))
            strSQL = strSQL.Replace("@EstimateBudget", CStr(enumPeriodType.EstimateBudget))
            strSQL = strSQL.Replace("@ActualData", CStr(enumUploadDataType.ActualData))
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)

            da = New SqlDataAdapter(strSQL, conn)
            ds = New DataSet

            da.Fill(ds, TableName)

            Me.DS = ds

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select011_2] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

#End Region

#Region "Select012"

    ''' <summary>
    ''' Select Budget data list for Report006-1
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select012_1() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT012_1")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECT_NO", Me.ProjectNo)

            If Me.UserPIC = "0" Then
                strSQL = strSQL.Replace("@PICCONDITION", " ")
            Else
                strSQL = strSQL.Replace("@PICCONDITION", "AND BUDGET_ORDER.PERSON_IN_CHARGE_NO = '@PIC'")
                strSQL = strSQL.Replace("@PIC", Me.UserPIC)
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
            MessageBox.Show("[BG_T_BUDGET_DATA.Select012_1] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' Select Budget data list for Report006-2
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select012_2() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT012_2")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECT_NO", Me.ProjectNo)

            If Me.UserPIC = "0" Then
                strSQL = strSQL.Replace("@PICCONDITION", " ")
            Else
                strSQL = strSQL.Replace("@PICCONDITION", "AND BUDGET_ORDER.PERSON_IN_CHARGE_NO = '@PIC'")
                strSQL = strSQL.Replace("@PIC", Me.UserPIC)
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
            MessageBox.Show("[BG_T_BUDGET_DATA.Select012_2] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' Select Budget data list for Report006-3
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select012_3() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT012_3")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECT_NO", Me.ProjectNo)

            If Me.UserPIC = "0" Then
                strSQL = strSQL.Replace("@PICCONDITION", " ")
            Else
                strSQL = strSQL.Replace("@PICCONDITION", "AND BUDGET_ORDER.PERSON_IN_CHARGE_NO = '@PIC'")
                strSQL = strSQL.Replace("@PIC", Me.UserPIC)
            End If

            'If Me.MTPChecked = True Then
            '    Dim strMTPSql As String = " OR ISNULL(MASTER_DATA.RRT1, 0)<> 0 OR ISNULL(MASTER_DATA.RRT2, 0) <> 0 OR ISNULL(MASTER_DATA.RRT3, 0) <> 0 OR ISNULL(MASTER_DATA.RRT4, 0) <> 0 OR ISNULL(MASTER_DATA.RRT5, 0) <> 0 "
            '    Dim intLen As Integer = strSQL.LastIndexOf("ORDER BY")
            '    strSQL = strSQL.Insert(intLen, strMTPSql)
            'End If

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select012_3] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try

    End Function


    ''' <summary>
    ''' Select Budget data list for Report006-1
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select012_4() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT012_4")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECT_NO", Me.ProjectNo)
            strSQL = strSQL.Replace("@REV_NO", Me.RevNo)

            If Me.UserPIC = "0" Then
                strSQL = strSQL.Replace("@PICCONDITION", " ")
            Else
                strSQL = strSQL.Replace("@PICCONDITION", "AND BUDGET_ORDER.PERSON_IN_CHARGE_NO = '@PIC'")
                strSQL = strSQL.Replace("@PIC", Me.UserPIC)
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
            MessageBox.Show("[BG_T_BUDGET_DATA.Select012_4] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' Select Budget data list for Report006-2
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select012_5() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT012_5")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECT_NO", Me.ProjectNo)
            strSQL = strSQL.Replace("@REV_NO", Me.RevNo)

            If Me.UserPIC = "0" Then
                strSQL = strSQL.Replace("@PICCONDITION", " ")
            Else
                strSQL = strSQL.Replace("@PICCONDITION", "AND BUDGET_ORDER.PERSON_IN_CHARGE_NO = '@PIC'")
                strSQL = strSQL.Replace("@PIC", Me.UserPIC)
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
            MessageBox.Show("[BG_T_BUDGET_DATA.Select012_5] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' Select Budget data list for Report006-3
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select012_6() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT012_6")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECT_NO", Me.ProjectNo)
            strSQL = strSQL.Replace("@REV_NO", Me.RevNo)

            If Me.UserPIC = "0" Then
                strSQL = strSQL.Replace("@PICCONDITION", " ")
            Else
                strSQL = strSQL.Replace("@PICCONDITION", "AND BUDGET_ORDER.PERSON_IN_CHARGE_NO = '@PIC'")
                strSQL = strSQL.Replace("@PIC", Me.UserPIC)
            End If

            'If Me.MTPChecked = True Then
            '    Dim strMTPSql As String = " OR ISNULL(MASTER_DATA.RRT1, 0)<> 0 OR ISNULL(MASTER_DATA.RRT2, 0) <> 0 OR ISNULL(MASTER_DATA.RRT3, 0) <> 0 OR ISNULL(MASTER_DATA.RRT4, 0) <> 0 OR ISNULL(MASTER_DATA.RRT5, 0) <> 0 "
            '    Dim intLen As Integer = strSQL.LastIndexOf("ORDER BY")
            '    strSQL = strSQL.Insert(intLen, strMTPSql)
            'End If

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select012_6] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try

    End Function

#End Region

#Region "Select013"
    ''' <summary>
    ''' Select all budget data by Status
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select013() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT013")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@Status", Me.Status)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            'strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            'strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefBudgetYear", Me.RefBudgetYear)
            strSQL = strSQL.Replace("@RefPeriodType", Me.RefPeriodType)
            strSQL = strSQL.Replace("@RefProjectNo", Me.RefProjectNo)
            strSQL = strSQL.Replace("@RefRevNo", Me.RefRevNo)

            strSQL = strSQL.Replace("@RefEstProjectNo", Me.RefEstProjectNo)
            strSQL = strSQL.Replace("@RefEstRevNo", Me.RefEstRevNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select013] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select014"

    ''' <summary>
    ''' Select Budget data list for Report002-1
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select014_1() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT014_1")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECTNO", Me.ProjectNo)

            'strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefEstProjectNo", Me.RefEstProjectNo)
            strSQL = strSQL.Replace("@RefEstRevNo", Me.RefEstRevNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select014_1] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' Select Budget data list for Report002-2
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select014_2() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT014_2")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECTNO", Me.ProjectNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select014_2] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' Select Budget data list for Report002-1
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select014_3() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT014_3")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECTNO", Me.ProjectNo)

            'If Me.MTPChecked = True Then
            '    Dim strMTPSql As String = " OR ISNULL(MASTER_DATA.RRT1, 0)<> 0 OR ISNULL(MASTER_DATA.RRT2, 0) <> 0 OR ISNULL(MASTER_DATA.RRT3, 0) <> 0 OR ISNULL(MASTER_DATA.RRT4, 0) <> 0 OR ISNULL(MASTER_DATA.RRT5, 0) <> 0 "
            '    Dim intLen As Integer = strSQL.LastIndexOf("GROUP BY")
            '    strSQL = strSQL.Insert(intLen, strMTPSql)
            'End If

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select014_3] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

#End Region

#Region "Select014 for Admin"

    ''' <summary>
    ''' Select Budget data list for Report002-1
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select014_4() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT014_4")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECTNO", Me.ProjectNo)
            strSQL = strSQL.Replace("@REVNO", Me.RevNo)

            strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefEstProjectNo", Me.RefEstProjectNo)
            strSQL = strSQL.Replace("@RefEstRevNo", Me.RefEstRevNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select014_4] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' Select Budget data list for Report002-2
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select014_5() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT014_5")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECTNO", Me.ProjectNo)
            strSQL = strSQL.Replace("@REVNO", Me.RevNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select014_5] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' Select Budget data list for Report002-1
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select014_6() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT014_6")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECTNO", Me.ProjectNo)
            strSQL = strSQL.Replace("@REVNO", Me.RevNo)

            'If Me.MTPChecked = True Then
            '    Dim strMTPSql As String = " OR ISNULL(MASTER_DATA.RRT1, 0)<> 0 OR ISNULL(MASTER_DATA.RRT2, 0) <> 0 OR ISNULL(MASTER_DATA.RRT3, 0) <> 0 OR ISNULL(MASTER_DATA.RRT4, 0) <> 0 OR ISNULL(MASTER_DATA.RRT5, 0) <> 0 "
            '    Dim intLen As Integer = strSQL.LastIndexOf("GROUP BY")
            '    strSQL = strSQL.Insert(intLen, strMTPSql)
            'End If

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select014_6] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

#End Region

#Region "Select015"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select015_1() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim strSQL As String
        Dim strDataTableName As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            Dim strAccountNo As String = String.Empty

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT015_1")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@Admin", Convert.ToString(BGConstant.enumCost.ADMIN))
            strSQL = strSQL.Replace("@FC", Convert.ToString(BGConstant.enumCost.FC))
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            'strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)

            strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefEstProjectNo", Me.RefEstProjectNo)
            strSQL = strSQL.Replace("@RefEstRevNo", Me.RefEstRevNo)

            da = New SqlDataAdapter(strSQL, conn)
            ds = New DataSet

            da.Fill(ds, TableName)

            Me.DS = ds

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select015_1] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select015_2() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim strSQL As String
        Dim strDataTableName As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            Dim strAccountNo As String = String.Empty

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT015_2")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@Admin", Convert.ToString(BGConstant.enumCost.ADMIN))
            strSQL = strSQL.Replace("@FC", Convert.ToString(BGConstant.enumCost.FC))
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            da = New SqlDataAdapter(strSQL, conn)
            ds = New DataSet

            da.Fill(ds, TableName)

            Me.DS = ds

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select015_2] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select015_3() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim strSQL As String
        Dim strDataTableName As String = String.Empty
        'Dim strValueName As String = String.Empty
        Dim strAccountNo As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            'If Not Me.MTPBudget Then
            '    strValueName = "SELECT015_3"
            'Else
            '    strValueName = "SELECT015_3_2"
            'End If

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT015_3")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@Admin", Convert.ToString(BGConstant.enumCost.ADMIN))
            strSQL = strSQL.Replace("@FC", Convert.ToString(BGConstant.enumCost.FC))
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            da = New SqlDataAdapter(strSQL, conn)
            ds = New DataSet

            da.Fill(ds, TableName)

            Me.DS = ds

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select015_3] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select015_4() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim strSQL As String
        Dim strDataTableName As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            Dim strAccountNo As String = String.Empty

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT015_4")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@Admin", Convert.ToString(BGConstant.enumCost.ADMIN))
            strSQL = strSQL.Replace("@FC", Convert.ToString(BGConstant.enumCost.FC))
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)

            strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefEstProjectNo", Me.RefEstProjectNo)
            strSQL = strSQL.Replace("@RefEstRevNo", Me.RefEstRevNo)

            da = New SqlDataAdapter(strSQL, conn)
            ds = New DataSet

            da.Fill(ds, TableName)

            Me.DS = ds

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select015_4] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select015_5() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim strSQL As String
        Dim strDataTableName As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            Dim strAccountNo As String = String.Empty

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT015_5")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@Admin", Convert.ToString(BGConstant.enumCost.ADMIN))
            strSQL = strSQL.Replace("@FC", Convert.ToString(BGConstant.enumCost.FC))
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)

            da = New SqlDataAdapter(strSQL, conn)
            ds = New DataSet

            da.Fill(ds, TableName)

            Me.DS = ds

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select015_5] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select015_6() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim strSQL As String
        Dim strDataTableName As String = String.Empty
        'Dim strValueName As String = String.Empty
        Dim strAccountNo As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            'If Not Me.MTPBudget Then
            '    strValueName = "SELECT015_6"
            'Else
            '    strValueName = "SELECT015_6_2"
            'End If

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT015_6")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@Admin", Convert.ToString(BGConstant.enumCost.ADMIN))
            strSQL = strSQL.Replace("@FC", Convert.ToString(BGConstant.enumCost.FC))
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)

            da = New SqlDataAdapter(strSQL, conn)
            ds = New DataSet

            da.Fill(ds, TableName)

            Me.DS = ds

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select015_6] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

#End Region

#Region "Select016"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select016() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim strSQL As String
        Dim strDataTableName As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            Dim strAccountNo As String = String.Empty

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT016")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@ReviseBudget", CStr(enumPeriodType.ReviseBudget))
            strSQL = strSQL.Replace("@ActualData", CStr(enumUploadDataType.ActualData))

            da = New SqlDataAdapter(strSQL, conn)
            ds = New DataSet

            da.Fill(ds, TableName)

            Me.DS = ds

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select016] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select017"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select017() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim strSQL As String
        Dim strDataTableName As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            Dim strAccountNo As String = String.Empty

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT017")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@ReviseBudget", CStr(enumPeriodType.ReviseBudget))
            strSQL = strSQL.Replace("@ActualData", CStr(enumUploadDataType.ActualData))

            da = New SqlDataAdapter(strSQL, conn)
            ds = New DataSet

            da.Fill(ds, TableName)

            Me.DS = ds

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select017] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select018"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select018() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String
        Dim strDataTableName As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            Dim strAccountNo As String = String.Empty

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT018")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
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
            MessageBox.Show("[BG_T_BUDGET_DATA.Select018] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select019"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select019() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String
        Dim strDataTableName As String = String.Empty

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            Dim strAccountNo As String = String.Empty

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT019")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
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
            MessageBox.Show("[BG_T_BUDGET_DATA.Select019] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select020"
    ''' <summary>
    ''' Select budget data by PIC (ReOpen Period)
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select020() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT020")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@UserPIC", Me.UserPIC)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            'strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            'strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefBudgetYear", Me.RefBudgetYear)
            strSQL = strSQL.Replace("@RefPeriodType", Me.RefPeriodType)
            strSQL = strSQL.Replace("@RefProjectNo", Me.RefProjectNo)
            strSQL = strSQL.Replace("@RefRevNo", Me.RefRevNo)

            strSQL = strSQL.Replace("@RefEstProjectNo", Me.RefEstProjectNo)
            strSQL = strSQL.Replace("@RefEstRevNo", Me.RefEstRevNo)


            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select020] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select021"
    ''' <summary>
    ''' Select all budget data
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select021() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT021")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            'strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            'strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefBudgetYear", Me.RefBudgetYear)
            strSQL = strSQL.Replace("@RefPeriodType", Me.RefPeriodType)
            strSQL = strSQL.Replace("@RefProjectNo", Me.RefProjectNo)
            strSQL = strSQL.Replace("@RefRevNo", Me.RefRevNo)

            strSQL = strSQL.Replace("@RefEstProjectNo", Me.RefEstProjectNo)
            strSQL = strSQL.Replace("@RefEstRevNo", Me.RefEstRevNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select021] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select022"
    ''' <summary>
    ''' Select total of budget data by budget order
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select022() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT022")
            strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select022] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select023"
    ''' <summary>
    ''' Select all budget data
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select023() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT023")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@ReviseRevNo", Me.ReviseRevNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@UserPIC", Me.UserPIC)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            'strSQL = strSQL.Replace("@PrevMTPRevNo", Me.PrevMTPRevNo)

            strSQL = strSQL.Replace("@RefBudgetYear", Me.RefBudgetYear)
            strSQL = strSQL.Replace("@RefPeriodType", Me.RefPeriodType)
            strSQL = strSQL.Replace("@RefProjectNo", Me.RefProjectNo)
            strSQL = strSQL.Replace("@RefRevNo", Me.RefRevNo)

            strSQL = strSQL.Replace("@RefRBProjectNo", Me.RefRBProjectNo)
            strSQL = strSQL.Replace("@RefRBRevNo", Me.RefRBRevNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select023] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select024"
    ''' <summary>
    ''' Select budget data by PIC (ReOpen Period)
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select024() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT024")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@ReviseRevNo", Me.ReviseRevNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            'strSQL = strSQL.Replace("@PrevMTPRevNo", Me.PrevMTPRevNo)

            strSQL = strSQL.Replace("@RefBudgetYear", Me.RefBudgetYear)
            strSQL = strSQL.Replace("@RefPeriodType", Me.RefPeriodType)
            strSQL = strSQL.Replace("@RefProjectNo", Me.RefProjectNo)
            strSQL = strSQL.Replace("@RefRevNo", Me.RefRevNo)

            strSQL = strSQL.Replace("@RefRBProjectNo", Me.RefRBProjectNo)
            strSQL = strSQL.Replace("@RefRBRevNo", Me.RefRBRevNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select024] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select025"
    ''' <summary>
    ''' Select budget data by PIC
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select025() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT025")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@ReviseRevNo", Me.ReviseRevNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@UserPIC", Me.UserPIC)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            'strSQL = strSQL.Replace("@PrevMTPRevNo", Me.PrevMTPRevNo)

            strSQL = strSQL.Replace("@RefBudgetYear", Me.RefBudgetYear)
            strSQL = strSQL.Replace("@RefPeriodType", Me.RefPeriodType)
            strSQL = strSQL.Replace("@RefProjectNo", Me.RefProjectNo)
            strSQL = strSQL.Replace("@RefRevNo", Me.RefRevNo)

            strSQL = strSQL.Replace("@RefRBProjectNo", Me.RefRBProjectNo)
            strSQL = strSQL.Replace("@RefRBRevNo", Me.RefRBRevNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select025] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select025_1"
    ''' <summary>
    ''' Select budget data by PIC
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select025_1() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT025_1")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@ReviseRevNo", Me.ReviseRevNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@UserPIC", Me.UserPIC)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            'strSQL = strSQL.Replace("@PrevMTPRevNo", Me.PrevMTPRevNo)

            strSQL = strSQL.Replace("@RefBudgetYear", Me.RefBudgetYear)
            strSQL = strSQL.Replace("@RefPeriodType", Me.RefPeriodType)
            strSQL = strSQL.Replace("@RefProjectNo", Me.RefProjectNo)
            strSQL = strSQL.Replace("@RefRevNo", Me.RefRevNo)

            strSQL = strSQL.Replace("@RefRBProjectNo", Me.RefRBProjectNo)
            strSQL = strSQL.Replace("@RefRBRevNo", Me.RefRBRevNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select025] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select026"
    ''' <summary>
    ''' Select all budget data
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select026() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT026")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@ReviseRevNo", Me.ReviseRevNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            'strSQL = strSQL.Replace("@PrevMTPRevNo", Me.PrevMTPRevNo)

            strSQL = strSQL.Replace("@RefBudgetYear", Me.RefBudgetYear)
            strSQL = strSQL.Replace("@RefPeriodType", Me.RefPeriodType)
            strSQL = strSQL.Replace("@RefProjectNo", Me.RefProjectNo)
            strSQL = strSQL.Replace("@RefRevNo", Me.RefRevNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select026] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select027"
    ''' <summary>
    ''' Select all budget data by Status
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select027() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT027")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@ReviseRevNo", Me.ReviseRevNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@Status", Me.Status)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            'strSQL = strSQL.Replace("@PrevMTPRevNo", Me.PrevMTPRevNo)

            strSQL = strSQL.Replace("@RefBudgetYear", Me.RefBudgetYear)
            strSQL = strSQL.Replace("@RefPeriodType", Me.RefPeriodType)
            strSQL = strSQL.Replace("@RefProjectNo", Me.RefProjectNo)
            strSQL = strSQL.Replace("@RefRevNo", Me.RefRevNo)

            strSQL = strSQL.Replace("@RefRBProjectNo", Me.RefRBProjectNo)
            strSQL = strSQL.Replace("@RefRBRevNo", Me.RefRBRevNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select027] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select028"
    ''' <summary>
    ''' Select all budget data by Status
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select028() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT028")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@Status", Me.Status)
            'strSQL = strSQL.Replace("@PrevProjectNo", Me.PrevProjectNo)
            strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefRBProjectNo", Me.RefRBProjectNo)
            strSQL = strSQL.Replace("@RefRBRevNo", Me.RefRBRevNo)

            If Me.AccountNo.Equals(String.Empty) OrElse _
                String.Equals(Me.AccountNo.Trim.ToUpper, STRING_ALL.ToUpper) Then
                strSQL = strSQL.Replace("@AccountNo", "")

            Else
                strSQL = strSQL.Replace("@AccountNo", " AND BO.ACCOUNT_NO = '" & Me.AccountNo & "'")

            End If

            da = New SqlDataAdapter(strSQL, conn)
            DS = New DataSet

            da.Fill(DS, TableName)

            Me.DS = DS

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select028] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select028_2"
    ''' <summary>
    ''' Select all budget data by Status
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select028_2() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT028_2")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@Status", Me.Status)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            'strSQL = strSQL.Replace("@PrevProjectNo", Me.PrevProjectNo)
            'strSQL = strSQL.Replace("@PrevRevNo", Me.PrevMTPRevNo)

            strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefRBProjectNo", Me.RefRBProjectNo)
            strSQL = strSQL.Replace("@RefRBRevNo", Me.RefRBRevNo)

            If Me.AccountNo.Equals(String.Empty) OrElse _
                String.Equals(Me.AccountNo.Trim.ToUpper, STRING_ALL.ToUpper) Then
                strSQL = strSQL.Replace("@AccountNo", "")

            Else
                strSQL = strSQL.Replace("@AccountNo", " AND BO.ACCOUNT_NO = '" & Me.AccountNo & "'")

            End If

            da = New SqlDataAdapter(strSQL, conn)
            DS = New DataSet

            da.Fill(DS, TableName)

            Me.DS = DS

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select028_2] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select029"
    ''' <summary>
    ''' Select all budget data by Status
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select029() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT029")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@Status", Me.Status)
            'strSQL = strSQL.Replace("@PrevProjectNo", Me.PrevProjectNo)

            strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefRBProjectNo", Me.RefRBProjectNo)
            strSQL = strSQL.Replace("@RefRBRevNo", Me.RefRBRevNo)

            da = New SqlDataAdapter(strSQL, conn)
            DS = New DataSet

            da.Fill(DS, TableName)

            Me.DS = DS

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select029] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select029_2"
    ''' <summary>
    ''' Select all budget data by Status
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select029_2() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT029_2")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@Status", Me.Status)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            'strSQL = strSQL.Replace("@PrevProjectNo", Me.PrevProjectNo)
            'strSQL = strSQL.Replace("@PrevRevNo", Me.PrevMTPRevNo)

            strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefRBProjectNo", Me.RefRBProjectNo)
            strSQL = strSQL.Replace("@RefRBRevNo", Me.RefRBRevNo)

            da = New SqlDataAdapter(strSQL, conn)
            DS = New DataSet

            da.Fill(DS, TableName)

            Me.DS = DS

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select029_2] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select030"
    ''' <summary>
    ''' Select all budget data by Status
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select030() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT030")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@Status", Me.Status)
            'strSQL = strSQL.Replace("@PrevProjectNo", Me.PrevProjectNo)

            strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefRBProjectNo", Me.RefRBProjectNo)
            strSQL = strSQL.Replace("@RefRBRevNo", Me.RefRBRevNo)

            da = New SqlDataAdapter(strSQL, conn)
            DS = New DataSet

            da.Fill(DS, TableName)

            Me.DS = DS

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select030] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select030_2"
    ''' <summary>
    ''' Select all budget data by Status
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select030_2() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT030_2")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@Status", Me.Status)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            'strSQL = strSQL.Replace("@PrevProjectNo", Me.PrevProjectNo)
            'strSQL = strSQL.Replace("@PrevRevNo", Me.PrevMTPRevNo)

            strSQL = strSQL.Replace("@MtpProjectNo", Me.MtpProjectNo)
            strSQL = strSQL.Replace("@MtpRevNo", Me.MtpRevNo)

            strSQL = strSQL.Replace("@RefRBProjectNo", Me.RefRBProjectNo)
            strSQL = strSQL.Replace("@RefRBRevNo", Me.RefRBRevNo)

            da = New SqlDataAdapter(strSQL, conn)
            DS = New DataSet

            da.Fill(DS, TableName)

            Me.DS = DS

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select030_2] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select031"
    ''' <summary>
    ''' Select all budget data by Status
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select031() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim strSQL As String
        Dim dt As DataTable

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "SELECT031")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)

            da = New SqlDataAdapter(strSQL, conn)
            DS = New DataSet
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Select031] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Insert001"
    ''' <summary>
    ''' Insert001
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "INSERT001")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@UserId", Me.UserId)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

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
            MessageBox.Show("[BG_T_BUDGET_DATA.Insert001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Insert002"
    ''' <summary>
    ''' Insert002
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Insert002(Optional ByVal pConn As SqlConnection = Nothing, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "INSERT002")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@M01", Me.M(1))
            strSQL = strSQL.Replace("@M02", Me.M(2))
            strSQL = strSQL.Replace("@M03", Me.M(3))
            strSQL = strSQL.Replace("@M04", Me.M(4))
            strSQL = strSQL.Replace("@M05", Me.M(5))
            strSQL = strSQL.Replace("@M06", Me.M(6))
            strSQL = strSQL.Replace("@M07", Me.M(7))
            strSQL = strSQL.Replace("@M08", Me.M(8))
            strSQL = strSQL.Replace("@M09", Me.M(9))
            strSQL = strSQL.Replace("@M10", Me.M(10))
            strSQL = strSQL.Replace("@M11", Me.M(11))
            strSQL = strSQL.Replace("@M12", Me.M(12))
            strSQL = strSQL.Replace("@RRT1", Me.RRT(1))
            strSQL = strSQL.Replace("@RRT2", Me.RRT(2))
            strSQL = strSQL.Replace("@RRT3", Me.RRT(3))
            strSQL = strSQL.Replace("@RRT4", Me.RRT(4))
            strSQL = strSQL.Replace("@RRT5", Me.RRT(5))
            strSQL = strSQL.Replace("@Remarks", Me.Remarks)
            strSQL = strSQL.Replace("@UserId", Me.UserId)

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
            MessageBox.Show("[BG_T_BUDGET_DATA.Insert002] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Insert003"
    ''' <summary>
    ''' Insert Estimate Actual Oct
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Insert003(Optional ByVal pConn As SqlConnection = Nothing, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "INSERT003")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@M10", Me.M(10))
            strSQL = strSQL.Replace("@UserId", Me.UserId)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

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
            MessageBox.Show("[BG_T_BUDGET_DATA.Insert003] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Insert004"
    ''' <summary>
    ''' Insert Revise Actual Apr
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Insert004(Optional ByVal pConn As SqlConnection = Nothing, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "INSERT004")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@M04", Me.M(4))
            strSQL = strSQL.Replace("@UserId", Me.UserId)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

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
            MessageBox.Show("[BG_T_BUDGET_DATA.Insert004] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Insert005"
    ''' <summary>
    ''' Insert data to table
    ''' </summary>
    ''' <param name="pConn">Connection</param>
    ''' <param name="pTrans">Transaction</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Insert005(Optional ByVal pConn As SqlConnection = Nothing, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "INSERT005")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@UserId", Me.UserId)
            strSQL = strSQL.Replace("@M01", (CDbl(Me.DataList(0)) / 1000).ToString)
            strSQL = strSQL.Replace("@M02", (CDbl(Me.DataList(1)) / 1000).ToString)
            strSQL = strSQL.Replace("@M03", (CDbl(Me.DataList(2)) / 1000).ToString)
            strSQL = strSQL.Replace("@M04", (CDbl(Me.DataList(3)) / 1000).ToString)
            strSQL = strSQL.Replace("@M05", (CDbl(Me.DataList(4)) / 1000).ToString)
            strSQL = strSQL.Replace("@M06", (CDbl(Me.DataList(5)) / 1000).ToString)
            strSQL = strSQL.Replace("@M07", (CDbl(Me.DataList(6)) / 1000).ToString)
            strSQL = strSQL.Replace("@M08", (CDbl(Me.DataList(7)) / 1000).ToString)
            strSQL = strSQL.Replace("@M09", (CDbl(Me.DataList(8)) / 1000).ToString)
            strSQL = strSQL.Replace("@M10", (CDbl(Me.DataList(9)) / 1000).ToString)
            strSQL = strSQL.Replace("@M11", (CDbl(Me.DataList(10)) / 1000).ToString)
            strSQL = strSQL.Replace("@M12", (CDbl(Me.DataList(11)) / 1000).ToString)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

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
            MessageBox.Show("[BG_T_UPLOAD_DATA.Insert005] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Insert006"
    ''' <summary>
    ''' Insert data to table
    ''' </summary>
    ''' <param name="pConn">Connection</param>
    ''' <param name="pTrans">Transaction</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Insert006(Optional ByVal pConn As SqlConnection = Nothing, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "INSERT006")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@UserId", Me.UserId)
            strSQL = strSQL.Replace("@RRT1", (CDbl(Me.DataList(0)) / 1000).ToString)
            strSQL = strSQL.Replace("@RRT2", (CDbl(Me.DataList(1)) / 1000).ToString)
            strSQL = strSQL.Replace("@RRT3", (CDbl(Me.DataList(2)) / 1000).ToString)
            strSQL = strSQL.Replace("@RRT4", (CDbl(Me.DataList(3)) / 1000).ToString)
            strSQL = strSQL.Replace("@RRT5", (CDbl(Me.DataList(4)) / 1000).ToString)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

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
            MessageBox.Show("[BG_T_UPLOAD_DATA.Insert005] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Update001"
    ''' <summary>
    ''' Update001
    ''' </summary>
    ''' <param name="pConn">Connection</param>
    ''' <param name="pTrans">Transaction</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "UPDATE001")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            If Me.PeriodType = CStr(enumPeriodType.OriginalBudget) Then
                strSQL = strSQL.Replace("@M01", CStr(Nz(Me.BudgetData("M1"), "0")).Replace(",", ""))
                strSQL = strSQL.Replace("@M02", CStr(Nz(Me.BudgetData("M2"), "0")).Replace(",", ""))
                strSQL = strSQL.Replace("@M03", CStr(Nz(Me.BudgetData("M3"), "0")).Replace(",", ""))
                strSQL = strSQL.Replace("@M04", CStr(Nz(Me.BudgetData("M4"), "0")).Replace(",", ""))
                strSQL = strSQL.Replace("@M05", CStr(Nz(Me.BudgetData("M5"), "0")).Replace(",", ""))
                strSQL = strSQL.Replace("@M06", CStr(Nz(Me.BudgetData("M6"), "0")).Replace(",", ""))
                If Me.BudgetType = P_BUDGET_TYPE_EXPENSE Then
                    strSQL = strSQL.Replace("@M07", CStr(Nz(Me.BudgetData("M7"), "0")).Replace(",", ""))
                    strSQL = strSQL.Replace("@M08", "0")
                    strSQL = strSQL.Replace("@M09", "0")
                    strSQL = strSQL.Replace("@M10", "0")
                    strSQL = strSQL.Replace("@M11", "0")
                    strSQL = strSQL.Replace("@M12", "0")
                Else
                    strSQL = strSQL.Replace("@M07", CStr(Nz(Me.BudgetData("M7"), "0")).Replace(",", ""))
                    strSQL = strSQL.Replace("@M08", CStr(Nz(Me.BudgetData("M8"), "0")).Replace(",", ""))
                    strSQL = strSQL.Replace("@M09", CStr(Nz(Me.BudgetData("M9"), "0")).Replace(",", ""))
                    strSQL = strSQL.Replace("@M10", CStr(Nz(Me.BudgetData("M10"), "0")).Replace(",", ""))
                    strSQL = strSQL.Replace("@M11", CStr(Nz(Me.BudgetData("M11"), "0")).Replace(",", ""))
                    strSQL = strSQL.Replace("@M12", CStr(Nz(Me.BudgetData("M12"), "0")).Replace(",", ""))
                End If

                strSQL = strSQL.Replace("@RRT1", "0")
                strSQL = strSQL.Replace("@RRT2", "0")
                strSQL = strSQL.Replace("@RRT3", "0")
                strSQL = strSQL.Replace("@RRT4", "0")
                strSQL = strSQL.Replace("@RRT5", "0")

            ElseIf Me.PeriodType = CStr(enumPeriodType.EstimateBudget) Then
                strSQL = strSQL.Replace("@M01", "0")
                strSQL = strSQL.Replace("@M02", "0")
                strSQL = strSQL.Replace("@M03", "0")
                strSQL = strSQL.Replace("@M04", "0")
                strSQL = strSQL.Replace("@M05", "0")
                strSQL = strSQL.Replace("@M06", "0")
                strSQL = strSQL.Replace("@M07", "0")
                strSQL = strSQL.Replace("@M08", "0")
                strSQL = strSQL.Replace("@M09", "0")
                strSQL = strSQL.Replace("@M10", CStr(Nz(Me.BudgetData("M10"), "0")).Replace(",", ""))
                strSQL = strSQL.Replace("@M11", CStr(Nz(Me.BudgetData("M11"), "0")).Replace(",", ""))
                strSQL = strSQL.Replace("@M12", CStr(Nz(Me.BudgetData("M12"), "0")).Replace(",", ""))

                strSQL = strSQL.Replace("@RRT1", "0")
                strSQL = strSQL.Replace("@RRT2", "0")
                strSQL = strSQL.Replace("@RRT3", "0")
                strSQL = strSQL.Replace("@RRT4", "0")
                strSQL = strSQL.Replace("@RRT5", "0")

            ElseIf Me.PeriodType = CStr(enumPeriodType.ReviseBudget) Then
                strSQL = strSQL.Replace("@M01", "0")
                strSQL = strSQL.Replace("@M02", "0")
                strSQL = strSQL.Replace("@M03", "0")
                strSQL = strSQL.Replace("@M04", CStr(Nz(Me.BudgetData("M4"), "0")).Replace(",", ""))
                strSQL = strSQL.Replace("@M05", CStr(Nz(Me.BudgetData("M5"), "0")).Replace(",", ""))
                strSQL = strSQL.Replace("@M06", CStr(Nz(Me.BudgetData("M6"), "0")).Replace(",", ""))
                strSQL = strSQL.Replace("@M07", CStr(Nz(Me.BudgetData("M7"), "0")).Replace(",", ""))
                strSQL = strSQL.Replace("@M08", CStr(Nz(Me.BudgetData("M8"), "0")).Replace(",", ""))
                strSQL = strSQL.Replace("@M09", CStr(Nz(Me.BudgetData("M9"), "0")).Replace(",", ""))
                strSQL = strSQL.Replace("@M10", CStr(Nz(Me.BudgetData("M10"), "0")).Replace(",", ""))
                strSQL = strSQL.Replace("@M11", CStr(Nz(Me.BudgetData("M11"), "0")).Replace(",", ""))
                strSQL = strSQL.Replace("@M12", CStr(Nz(Me.BudgetData("M12"), "0")).Replace(",", ""))


                strSQL = strSQL.Replace("@RRT1", "0")
                strSQL = strSQL.Replace("@RRT2", "0")
                strSQL = strSQL.Replace("@RRT3", "0")
                strSQL = strSQL.Replace("@RRT4", "0")
                strSQL = strSQL.Replace("@RRT5", "0")

                '// Comment by Ko
                'strSQL = strSQL.Replace("@RRT1", CStr(Nz(Me.BudgetData("RRT1"), "0")).Replace(",", ""))
                'strSQL = strSQL.Replace("@RRT2", CStr(Nz(Me.BudgetData("RRT2"), "0")).Replace(",", ""))
                'strSQL = strSQL.Replace("@RRT3", CStr(Nz(Me.BudgetData("RRT3"), "0")).Replace(",", ""))
                'strSQL = strSQL.Replace("@RRT4", CStr(Nz(Me.BudgetData("RRT4"), "0")).Replace(",", ""))
                'strSQL = strSQL.Replace("@RRT5", CStr(Nz(Me.BudgetData("RRT5"), "0")).Replace(",", ""))

            ElseIf Me.PeriodType = CStr(enumPeriodType.MTPBudget) Then

                strSQL = strSQL.Replace("@M01", "0")
                strSQL = strSQL.Replace("@M02", "0")
                strSQL = strSQL.Replace("@M03", "0")
                strSQL = strSQL.Replace("@M04", "0")
                strSQL = strSQL.Replace("@M05", "0")
                strSQL = strSQL.Replace("@M06", "0")
                strSQL = strSQL.Replace("@M07", "0")
                strSQL = strSQL.Replace("@M08", "0")
                strSQL = strSQL.Replace("@M09", "0")
                strSQL = strSQL.Replace("@M10", "0")
                strSQL = strSQL.Replace("@M11", "0")
                strSQL = strSQL.Replace("@M12", "0")

                strSQL = strSQL.Replace("@RRT1", CStr(Nz(Me.BudgetData("RRT1"), "0")).Replace(",", ""))
                strSQL = strSQL.Replace("@RRT2", CStr(Nz(Me.BudgetData("RRT2"), "0")).Replace(",", ""))
                strSQL = strSQL.Replace("@RRT3", CStr(Nz(Me.BudgetData("RRT3"), "0")).Replace(",", ""))
                strSQL = strSQL.Replace("@RRT4", CStr(Nz(Me.BudgetData("RRT4"), "0")).Replace(",", ""))
                strSQL = strSQL.Replace("@RRT5", CStr(Nz(Me.BudgetData("RRT5"), "0")).Replace(",", ""))


            End If

            strSQL = strSQL.Replace("@Remarks", CStr(Nz(Me.BudgetData("REMARKS"))).Replace("'", "_"))
            strSQL = strSQL.Replace("@WBFlag", CStr(IIf(CBool(Nz(Me.BudgetData("Adjust"), False)), 1, 0)))
            strSQL = strSQL.Replace("@UserId", Me.UserId)

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

            ''If intRtn > 0 Then
            Return True
            ''Else
            ''    Return False
            ''End If

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_DATA.Update001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Update002"
    ''' <summary>
    ''' Update Transfer Cost
    ''' </summary>
    ''' <param name="pConn">Connection</param>
    ''' <param name="pTrans">Transaction</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Update002(Optional ByVal pConn As SqlConnection = Nothing, _
                              Optional ByVal pTrans As SqlTransaction = Nothing) As Boolean
        Dim conn As SqlConnection = Nothing
        Dim cmd As SqlCommand
        Dim strSQL As String = String.Empty
        Dim intRtn As Integer

        Try
            If pConn Is Nothing Or pTrans Is Nothing Then
                conn = New SqlConnection(My.Settings.ConnStr)
                conn.Open()
            End If

            If Me.PeriodType = CStr(enumPeriodType.OriginalBudget) Then
                strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "UPDATE002-1")
                strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
                strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
                strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)
                strSQL = strSQL.Replace("@RevNo", Me.RevNo)
                strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

                strSQL = strSQL.Replace("@M01", CStr(Nz(Me.BudgetData2(0), "0")))
                strSQL = strSQL.Replace("@M02", CStr(Nz(Me.BudgetData2(1), "0")))
                strSQL = strSQL.Replace("@M03", CStr(Nz(Me.BudgetData2(2), "0")))
                strSQL = strSQL.Replace("@M04", CStr(Nz(Me.BudgetData2(3), "0")))
                strSQL = strSQL.Replace("@M05", CStr(Nz(Me.BudgetData2(4), "0")))
                strSQL = strSQL.Replace("@M06", CStr(Nz(Me.BudgetData2(5), "0")))
                strSQL = strSQL.Replace("@M07", CStr(Nz(Me.BudgetData2(6), "0")))
                strSQL = strSQL.Replace("@M08", CStr(Nz(Me.BudgetData2(7), "0")))
                strSQL = strSQL.Replace("@M09", CStr(Nz(Me.BudgetData2(8), "0")))
                strSQL = strSQL.Replace("@M10", CStr(Nz(Me.BudgetData2(9), "0")))
                strSQL = strSQL.Replace("@M11", CStr(Nz(Me.BudgetData2(10), "0")))
                strSQL = strSQL.Replace("@M12", CStr(Nz(Me.BudgetData2(11), "0")))

            ElseIf Me.PeriodType = CStr(enumPeriodType.EstimateBudget) Then
                strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "UPDATE002-2")
                strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
                strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
                strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)
                strSQL = strSQL.Replace("@RevNo", Me.RevNo)
                strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

                strSQL = strSQL.Replace("@M10", CStr(Nz(Me.BudgetData2(9), "0")))
                strSQL = strSQL.Replace("@M11", CStr(Nz(Me.BudgetData2(10), "0")))
                strSQL = strSQL.Replace("@M12", CStr(Nz(Me.BudgetData2(11), "0")))

            ElseIf Me.PeriodType = CStr(enumPeriodType.ReviseBudget) Then

                strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "UPDATE002-3")
                strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
                strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
                strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)
                strSQL = strSQL.Replace("@RevNo", Me.RevNo)
                strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

                strSQL = strSQL.Replace("@M04", CStr(Nz(Me.BudgetData2(3), "0")))
                strSQL = strSQL.Replace("@M05", CStr(Nz(Me.BudgetData2(4), "0")))
                strSQL = strSQL.Replace("@M06", CStr(Nz(Me.BudgetData2(5), "0")))
                strSQL = strSQL.Replace("@M07", CStr(Nz(Me.BudgetData2(6), "0")))
                strSQL = strSQL.Replace("@M08", CStr(Nz(Me.BudgetData2(7), "0")))
                strSQL = strSQL.Replace("@M09", CStr(Nz(Me.BudgetData2(8), "0")))
                strSQL = strSQL.Replace("@M10", CStr(Nz(Me.BudgetData2(9), "0")))
                strSQL = strSQL.Replace("@M11", CStr(Nz(Me.BudgetData2(10), "0")))
                strSQL = strSQL.Replace("@M12", CStr(Nz(Me.BudgetData2(11), "0")))

            ElseIf Me.PeriodType = CStr(enumPeriodType.MTPBudget) Then

                strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "UPDATE002-4")
                strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
                strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
                strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)
                strSQL = strSQL.Replace("@RevNo", Me.RevNo)
                strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

                strSQL = strSQL.Replace("@RRT1", CStr(Nz(Me.BudgetData2(0), "0")))
                strSQL = strSQL.Replace("@RRT2", CStr(Nz(Me.BudgetData2(1), "0")))
                strSQL = strSQL.Replace("@RRT3", CStr(Nz(Me.BudgetData2(2), "0")))
                strSQL = strSQL.Replace("@RRT4", CStr(Nz(Me.BudgetData2(3), "0")))
                strSQL = strSQL.Replace("@RRT5", CStr(Nz(Me.BudgetData2(4), "0")))



            End If

            strSQL = strSQL.Replace("@UserId", Me.UserId)

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
            MessageBox.Show("[BG_T_BUDGET_DATA.Update002] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Update003"
    ''' <summary>
    ''' Update Estimate Actual Oct
    ''' </summary>
    ''' <param name="pConn">Connection</param>
    ''' <param name="pTrans">Transaction</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Update003(Optional ByVal pConn As SqlConnection = Nothing, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "UPDATE003")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@M10", Me.M(10))
            strSQL = strSQL.Replace("@UserId", Me.UserId)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

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
            MessageBox.Show("[BG_T_BUDGET_DATA.Update003] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Update004"
    ''' <summary>
    ''' Update Revise Actual Apr
    ''' </summary>
    ''' <param name="pConn">Connection</param>
    ''' <param name="pTrans">Transaction</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Update004(Optional ByVal pConn As SqlConnection = Nothing, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "UPDATE004")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@M04", Me.M(4))
            strSQL = strSQL.Replace("@UserId", Me.UserId)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

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
            MessageBox.Show("[BG_T_BUDGET_DATA.Update004] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Update005"
    ''' <summary>
    ''' Update exist data
    ''' </summary>
    ''' <param name="pConn">Connection</param>
    ''' <param name="pTrans">Transaction</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Update005(Optional ByVal pConn As SqlConnection = Nothing, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "UPDATE005")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@UserId", Me.UserId)
            strSQL = strSQL.Replace("@M01", (CDbl(Me.DataList(0)) / 1000).ToString)
            strSQL = strSQL.Replace("@M02", (CDbl(Me.DataList(1)) / 1000).ToString)
            strSQL = strSQL.Replace("@M03", (CDbl(Me.DataList(2)) / 1000).ToString)
            strSQL = strSQL.Replace("@M04", (CDbl(Me.DataList(3)) / 1000).ToString)
            strSQL = strSQL.Replace("@M05", (CDbl(Me.DataList(4)) / 1000).ToString)
            strSQL = strSQL.Replace("@M06", (CDbl(Me.DataList(5)) / 1000).ToString)
            strSQL = strSQL.Replace("@M07", (CDbl(Me.DataList(6)) / 1000).ToString)
            strSQL = strSQL.Replace("@M08", (CDbl(Me.DataList(7)) / 1000).ToString)
            strSQL = strSQL.Replace("@M09", (CDbl(Me.DataList(8)) / 1000).ToString)
            strSQL = strSQL.Replace("@M10", (CDbl(Me.DataList(9)) / 1000).ToString)
            strSQL = strSQL.Replace("@M11", (CDbl(Me.DataList(10)) / 1000).ToString)
            strSQL = strSQL.Replace("@M12", (CDbl(Me.DataList(11)) / 1000).ToString)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

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
            MessageBox.Show("[BG_T_BUDGET_DATA.Update005] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Update006"
    ''' <summary>
    ''' Update exist data
    ''' </summary>
    ''' <param name="pConn">Connection</param>
    ''' <param name="pTrans">Transaction</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Update006(Optional ByVal pConn As SqlConnection = Nothing, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "UPDATE006")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@UserId", Me.UserId)
            strSQL = strSQL.Replace("@RRT1", (CDbl(Me.DataList(0)) / 1000).ToString)
            strSQL = strSQL.Replace("@RRT2", (CDbl(Me.DataList(1)) / 1000).ToString)
            strSQL = strSQL.Replace("@RRT3", (CDbl(Me.DataList(2)) / 1000).ToString)
            strSQL = strSQL.Replace("@RRT4", (CDbl(Me.DataList(3)) / 1000).ToString)
            strSQL = strSQL.Replace("@RRT5", (CDbl(Me.DataList(4)) / 1000).ToString)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

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
            MessageBox.Show("[BG_T_BUDGET_DATA.Update006] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Delete001"
    ''' <summary>
    ''' Delete Rev
    ''' </summary>
    ''' <param name="pConn">Connection</param>
    ''' <param name="pTrans">Transaction</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "DELETE001")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

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
            MessageBox.Show("[BG_T_BUDGET_DATA.Delete001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Delete002"
    ''' <summary>
    ''' Delete Specific PIC
    ''' </summary>
    ''' <param name="pConn">Connection</param>
    ''' <param name="pTrans">Transaction</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Delete002(Optional ByVal pConn As SqlConnection = Nothing, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "DELETE002")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetType", Me.BudgetType)
            strSQL = strSQL.Replace("@UserPIC", Me.UserPIC)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

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
            MessageBox.Show("[BG_T_BUDGET_DATA.Delete002] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Delete003"
    ''' <summary>
    ''' Delete Specific Budget Order
    ''' </summary>
    ''' <param name="pConn">Connection</param>
    ''' <param name="pTrans">Transaction</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Delete003(Optional ByVal pConn As SqlConnection = Nothing, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_DATA", "DELETE003")
            strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)

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
            MessageBox.Show("[BG_T_BUDGET_DATA.Delete003] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If pConn Is Nothing Or pTrans Is Nothing Then
                If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                    conn.Close()
                End If
            End If

            Return False

        End Try
    End Function
#End Region

#End Region

End Class

