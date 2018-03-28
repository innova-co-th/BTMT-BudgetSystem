Imports System.Data.SqlClient
Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class BG_T_BUDGET_PERIOD

#Region "Variable"
    Private myDtResult As DataTable
    Private myBudgetYear As String = String.Empty
    Private myPeriodType As String = String.Empty
    Private myHideFlag As String = String.Empty
    Private myUserId As String = String.Empty
    Private myCloseFlg As String = String.Empty
    Private myAccountNo As String = String.Empty
    Private myPicNo As String = String.Empty
    Private myProjectNo As String = String.Empty
    Private myViewAll As Boolean = False
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

#Region "HideFlag"
    Property HideFlag() As String
        Get
            Return myHideFlag
        End Get
        Set(ByVal value As String)
            myHideFlag = value
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

#Region "CloseFlg"
    Property CloseFlg() As String
        Get
            Return myCloseFlg
        End Get
        Set(ByVal value As String)
            myCloseFlg = value
        End Set
    End Property
#End Region

#Region "AccountNo"
    Property AccountNo() As String
        Get
            Return myAccountNo
        End Get
        Set(ByVal value As String)
            myAccountNo = value
        End Set
    End Property
#End Region

#Region "PicNo"
    Property PicNo() As String
        Get
            Return myPicNo
        End Get
        Set(ByVal value As String)
            myPicNo = value
        End Set
    End Property
#End Region

#Region "ViewAll"
    Property ViewAll() As Boolean
        Get
            Return myViewAll
        End Get
        Set(ByVal value As Boolean)
            myViewAll = value
        End Set
    End Property
#End Region

#End Region

#Region "Function"

#Region "Select001"
    ''' <summary>
    ''' Select Open Budget Period
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_PERIOD", "SELECT001")

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_PERIOD.Select001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Select Closed Budget Period
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_PERIOD", "SELECT002")

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_PERIOD.Select002] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Select Budget Period
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_PERIOD", "SELECT003")
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
            MessageBox.Show("[BG_T_BUDGET_PERIOD.Select003] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Select Budget Period for view
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select004() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_PERIOD", "SELECT004")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)

            If Me.ViewAll = False Then
                strSQL = strSQL.Replace("@ViewAll", " AND HIDE_FLAG = 0 ")
            Else
                strSQL = strSQL.Replace("@ViewAll", "")
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
            MessageBox.Show("[BG_T_BUDGET_PERIOD.Select004] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Select all Budget Period
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select005() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_PERIOD", "SELECT005")

            If Me.ViewAll = False Then
                strSQL = strSQL.Replace("@ViewAll", " WHERE HIDE_FLAG = 0 ")
            Else
                strSQL = strSQL.Replace("@ViewAll", "")
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
            MessageBox.Show("[BG_T_BUDGET_PERIOD.Select005] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Select Open/Reopen Budget Period
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select006() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_PERIOD", "SELECT006")

            If Me.ViewAll = False Then
                strSQL = strSQL.Replace("@ViewAll", " AND HIDE_FLAG = 0 ")
            Else
                strSQL = strSQL.Replace("@ViewAll", "")
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
            MessageBox.Show("[BG_T_BUDGET_PERIOD.Select006] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Select Budget Period
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select007() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_PERIOD", "SELECT007")
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
            MessageBox.Show("[BG_T_BUDGET_PERIOD.Select007] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Select View Budget Period
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select008() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_PERIOD", "SELECT008")

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_BUDGET_PERIOD.Select008] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select009"
    ''' <summary>
    ''' Select View Budget Period
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_PERIOD", "SELECT009")
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
            MessageBox.Show("[BG_T_BUDGET_PERIOD.Select009] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_PERIOD", "INSERT001")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
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
            MessageBox.Show("[BG_T_BUDGET_PERIOD.Insert001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_PERIOD", "UPDATE001")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@UserId", Me.UserId)
            strSQL = strSQL.Replace("@CloseFlg", Me.CloseFlg)
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
            MessageBox.Show("[BG_T_BUDGET_PERIOD.Update001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Update002
    ''' </summary>
    ''' <param name="pConn">Connection</param>
    ''' <param name="pTrans">Transaction</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_PERIOD", "UPDATE002")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@UserId", Me.UserId)
            strSQL = strSQL.Replace("@HideFlag", Me.HideFlag)
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
            MessageBox.Show("[BG_T_BUDGET_PERIOD.Update002] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
