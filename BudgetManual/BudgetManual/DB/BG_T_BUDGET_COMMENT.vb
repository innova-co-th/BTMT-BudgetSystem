Imports System.Data.SqlClient
Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Public Class BG_T_BUDGET_COMMENT

#Region "Variable"
    Private Const STRING_ALL As String = "All"

    Private myComment As String
    Private myBudgetYear As String
    Private myPeriodType As String
    Private myBudgetOrderNo As String
    Private myRevNo As String
    Private myProjectNo As String
    Private myMonthNo As String
    Private myRRTNo As String

    Private myCommentList As DataTable
    Private myCreateUserId As String

    Private myDtResult As DataTable
    Private myBudgetType As String = String.Empty
    Private myUserPIC As String = String.Empty

    Private myBudgetComment As DataRow


#End Region

#Region "Property"

#Region "Comment"
    Public Property Comment() As String
        Get
            Return myComment
        End Get
        Set(ByVal value As String)
            myComment = value
        End Set
    End Property
#End Region

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

#Region "BudgetOrderNo"
    Public Property BudgetOrderNo() As String
        Get
            Return myBudgetOrderNo
        End Get
        Set(ByVal value As String)
            myBudgetOrderNo = value
        End Set
    End Property
#End Region

#Region "RevNo"
    Public Property RevNo() As String
        Get
            Return myRevNo
        End Get
        Set(ByVal value As String)
            myRevNo = value
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

#Region "CommentList"
    Property CommentList() As DataTable
        Get
            Return myCommentList
        End Get
        Set(ByVal value As DataTable)
            myCommentList = value
        End Set
    End Property
#End Region

#Region "MonthNo"
    Public Property MonthNo() As String
        Get
            Return myMonthNo
        End Get
        Set(ByVal value As String)
            myMonthNo = value
        End Set
    End Property
#End Region

#Region "RRTNo"
    Public Property RRTNo() As String
        Get
            Return myRRTNo
        End Get
        Set(ByVal value As String)
            myRRTNo = value
        End Set
    End Property
#End Region

#Region "CreateUserId"
    Public Property CreateUserId() As String
        Get
            Return myCreateUserId
        End Get
        Set(ByVal value As String)
            myCreateUserId = value
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

#Region "BudgetComment"
    Property BudgetComment() As DataRow
        Get
            Return myBudgetComment
        End Get
        Set(ByVal value As DataRow)
            myBudgetComment = value
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_COMMENT", "SELECT001")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.CommentList = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_COMMENT.Select001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    Public Function Select002_1() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_COMMENT", "SELECT002")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECTNO", Me.ProjectNo)
            strSQL = strSQL.Replace("@BUDGETTYPE", Me.BudgetType)
            strSQL = strSQL.Replace("@REVNOCONDITION", "AND BC.REV_NO = @REVNO")
            strSQL = strSQL.Replace("@REVNO", Me.RevNo)

            If Me.UserPIC = "0" Then
                strSQL = strSQL.Replace("@PICCONDITION", " ")
            Else
                strSQL = strSQL.Replace("@PICCONDITION", "AND BO.PERSON_IN_CHARGE_NO = '@PIC'")
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
            MessageBox.Show("[BG_T_BUDGET_COMMENT.Select002] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function

    Public Function Select002_2() As Boolean

        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_COMMENT", "SELECT002")
            strSQL = strSQL.Replace("@REVNOCONDITION", "AND BC.REV_NO = (SELECT MAX(BD.REV_NO) FROM BG_T_BUDGET_COMMENT AS BD WHERE BD.BUDGET_YEAR = @YEAR AND BD.PERIOD_TYPE = @PERIOD AND BD.PROJECT_NO = @PROJECTNO)")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECTNO", Me.ProjectNo)
            strSQL = strSQL.Replace("@BUDGETTYPE", Me.BudgetType)
            'strSQL = strSQL.Replace("@REVNO", Me.RevNo)

            If Me.UserPIC = "0" Then
                strSQL = strSQL.Replace("@PICCONDITION", " ")
            Else
                strSQL = strSQL.Replace("@PICCONDITION", "AND BO.PERSON_IN_CHARGE_NO = '@PIC'")
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
            MessageBox.Show("[BG_T_BUDGET_COMMENT.Select002] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_COMMENT", "SELECT003")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.CommentList = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_COMMENT.Select003] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function


#Region "Update001"
    ''' <summary>
    ''' Update001
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Update001() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim cmd As SqlCommand
        Dim strSQL As String
        Dim intRtn As Integer

        Dim strCondition As String = String.Empty
        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_COMMENT", "UPDATE001")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            If Me.PeriodType = CStr(enumPeriodType.MTPBudget) Then
                Select Case Me.RRTNo
                    Case "1"
                        strCondition = " RRT1 = '" & Me.Comment & "'"
                    Case "2"
                        strCondition = " RRT2 = '" & Me.Comment & "'"
                    Case "3"
                        strCondition = " RRT3 = '" & Me.Comment & "'"
                    Case "4"
                        strCondition = " RRT4 = '" & Me.Comment & "'"
                    Case "5"
                        strCondition = " RRT5 = '" & Me.Comment & "'"
                End Select
            Else
                Select Case Me.MonthNo
                    Case "1"
                        strCondition = " M1 = '" & Me.Comment & "'"
                    Case "2"
                        strCondition = " M2 = '" & Me.Comment & "'"
                    Case "3"
                        strCondition = " M3 = '" & Me.Comment & "'"
                    Case "4"
                        strCondition = " M4 = '" & Me.Comment & "'"
                    Case "5"
                        strCondition = " M5 = '" & Me.Comment & "'"
                    Case "6"
                        strCondition = " M6 = '" & Me.Comment & "'"
                    Case "7"
                        strCondition = " M7 = '" & Me.Comment & "'"
                    Case "8"
                        strCondition = " M8 = '" & Me.Comment & "'"
                    Case "9"
                        strCondition = " M9 = '" & Me.Comment & "'"
                    Case "10"
                        strCondition = " M10 = '" & Me.Comment & "'"
                    Case "11"
                        strCondition = " M11 = '" & Me.Comment & "'"
                    Case "12"
                        strCondition = " M12 = '" & Me.Comment & "'"
                End Select
            End If

            strSQL = strSQL.Replace("@comment", strCondition)
            strSQL = strSQL.Replace("@UserId", Me.CreateUserId)


            cmd = New SqlCommand(strSQL, conn)
            intRtn = cmd.ExecuteNonQuery()

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            If intRtn > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_COMMENT.Update001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
        Dim strCondition As String = String.Empty

        Try
            If pConn Is Nothing Or pTrans Is Nothing Then
                conn = New SqlConnection(My.Settings.ConnStr)
                conn.Open()
            End If
            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_COMMENT", "INSERT001")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            If Me.PeriodType = CStr(enumPeriodType.MTPBudget) Then
                Select Case Me.RRTNo
                    Case "1"
                        strCondition = "RRT1"
                    Case "2"
                        strCondition = "RRT2"
                    Case "3"
                        strCondition = "RRT3"
                    Case "4"
                        strCondition = "RRT4"
                    Case "5"
                        strCondition = "RRT5"

                End Select
            Else
                Select Case Me.MonthNo
                    Case "1"
                        strCondition = "M1"
                    Case "2"
                        strCondition = "M2"
                    Case "3"
                        strCondition = "M3"
                    Case "4"
                        strCondition = "M4"
                    Case "5"
                        strCondition = "M5"
                    Case "6"
                        strCondition = "M6"
                    Case "7"
                        strCondition = "M7"
                    Case "8"
                        strCondition = "M8"
                    Case "9"
                        strCondition = "M9"
                    Case "10"
                        strCondition = "M10"
                    Case "11"
                        strCondition = "M11"
                    Case "12"
                        strCondition = "M12"
                End Select
            End If
           

            strSQL = strSQL.Replace("@Condition", strCondition)
            strSQL = strSQL.Replace("@Value", Me.Comment)
            strSQL = strSQL.Replace("@UserId", Me.CreateUserId)

            cmd = New SqlCommand(strSQL, conn)
            intRtn = cmd.ExecuteNonQuery()

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            If intRtn > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show("[BG_M_USER.Insert001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_COMMENT", "INSERT002")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetOrderNo", Me.BudgetOrderNo)
            strSQL = strSQL.Replace("@RevNo", Me.RevNo)
            strSQL = strSQL.Replace("@UserId", Me.CreateUserId)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            Dim dr As DataRow
            dr = Me.BudgetComment

            If Not dr Is Nothing Then
                strSQL = strSQL.Replace("@M01", dr.Item("M1").ToString)
                strSQL = strSQL.Replace("@M02", dr.Item("M2").ToString)
                strSQL = strSQL.Replace("@M03", dr.Item("M3").ToString)
                strSQL = strSQL.Replace("@M04", dr.Item("M4").ToString)
                strSQL = strSQL.Replace("@M05", dr.Item("M5").ToString)
                strSQL = strSQL.Replace("@M06", dr.Item("M6").ToString)
                strSQL = strSQL.Replace("@M07", dr.Item("M7").ToString)
                strSQL = strSQL.Replace("@M08", dr.Item("M8").ToString)
                strSQL = strSQL.Replace("@M09", dr.Item("M9").ToString)
                strSQL = strSQL.Replace("@M10", dr.Item("M10").ToString)
                strSQL = strSQL.Replace("@M11", dr.Item("M11").ToString)
                strSQL = strSQL.Replace("@M12", dr.Item("M12").ToString)

                strSQL = strSQL.Replace("@RRT1", dr.Item("RRT1").ToString)
                strSQL = strSQL.Replace("@RRT2", dr.Item("RRT2").ToString)
                strSQL = strSQL.Replace("@RRT3", dr.Item("RRT3").ToString)
                strSQL = strSQL.Replace("@RRT4", dr.Item("RRT4").ToString)
                strSQL = strSQL.Replace("@RRT5", dr.Item("RRT5").ToString)
            End If


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
            MessageBox.Show("[BG_T_BUDGET_COMMENT.Insert002] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_COMMENT", "DELETE001")
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
            MessageBox.Show("[BG_T_BUDGET_COMMENT.Delete001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
