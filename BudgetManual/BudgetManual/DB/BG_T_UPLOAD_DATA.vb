Imports System.Data.SqlClient
Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class BG_T_UPLOAD_DATA

#Region "Variable"
    Private Const STRING_ALL As String = "All"

    Private myDtResult As DataTable
    Private myBudgetYear As String = String.Empty
    Private myProjectNo As String = String.Empty
    Private myPeriodType As String = String.Empty
    Private myBudgetOrder As String = String.Empty
    Private myDataType As String = String.Empty
    Private myDataList As ArrayList
    Private myUserId As String = String.Empty
    Private myUserPIC As String = String.Empty
    Private myAccountNo As String = String.Empty
    Private myMonth As String = String.Empty
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

#Region " Month "
    Property Month() As String
        Get
            Return myMonth
        End Get
        Set(ByVal value As String)
            myMonth = value
        End Set
    End Property
#End Region

#End Region

#Region "Function"

#Region "Select001"
    ''' <summary>
    ''' Select specific data from table
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_UPLOAD_DATA", "SELECT001")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetOrder", Me.BudgetOrder)
            strSQL = strSQL.Replace("@DataType", Me.DataType)
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
            MessageBox.Show("[BG_T_UPLOAD_DATA.Select001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Select specific data from table
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_UPLOAD_DATA", "SELECT002")
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
            MessageBox.Show("[BG_T_UPLOAD_DATA.Select002] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Select Budget Compare list for Report007/007-2
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_UPLOAD_DATA", "SELECT003")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType1", Me.PeriodType)


            If Me.UserPIC.Equals(String.Empty) OrElse Me.UserPIC = "0" Then
                strSQL = strSQL.Replace("@PICCondition", " ")
            Else
                strSQL = strSQL.Replace("@PICCondition", "WHERE BUDGET_ORDER.PERSON_IN_CHARGE_NO = '@PIC'")
                strSQL = strSQL.Replace("@PIC", Me.UserPIC)
            End If

            If Me.AccountNo.Equals(String.Empty) OrElse String.Equals(Me.AccountNo.Trim.ToUpper, STRING_ALL.ToUpper) Then
                strSQL = strSQL.Replace("@ACCCondition", " ")
            Else
                strSQL = strSQL.Replace("@ACCCondition", "WHERE BUDGET_ORDER.ACCOUNT_NO = '@ACC'")
                strSQL = strSQL.Replace("@ACC", Me.AccountNo)
            End If

            strSQL = strSQL & " WHERE (OB_M" & Me.Month & " <> 0 " & _
                                    " OR AC_M" & Me.Month & " <> 0 " & _
                                    " OR ACC_OB_M" & Me.Month & " <> 0 " & _
                                    " OR ACC_AC_M" & Me.Month & " <> 0 )"

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_UPLOAD_DATA.Select003] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Select Budget Compare list for Report007-1
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_UPLOAD_DATA", "SELECT004")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType1", Me.PeriodType)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_UPLOAD_DATA.Select004] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Select Budget Compare list for Report007-3
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_UPLOAD_DATA", "SELECT005")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType1", Me.PeriodType)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_T_UPLOAD_DATA.Select005] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Select Budget Compare list for Report007-4
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_UPLOAD_DATA", "SELECT006")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType1", Me.PeriodType)

            If Me.UserPIC.Equals(String.Empty) OrElse Me.UserPIC = "0" Then
                strSQL = strSQL.Replace("@PICCondition", " ")
            Else
                strSQL = strSQL.Replace("@PICCondition", "AND BUDGET_ORDER.PERSON_IN_CHARGE_NO = '@PIC'")
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
            MessageBox.Show("[BG_T_UPLOAD_DATA.Select006] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Insert data to table
    ''' </summary>
    ''' <param name="pConn">Connection</param>
    ''' <param name="pTrans">Transaction</param>
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_UPLOAD_DATA", "INSERT001")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetOrder", Me.BudgetOrder)
            strSQL = strSQL.Replace("@DataType", Me.DataType)
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
            strSQL = strSQL.Replace("@H01", ((CDbl(Me.DataList(0)) + _
                                            CDbl(Me.DataList(1)) + _
                                            CDbl(Me.DataList(2)) + _
                                            CDbl(Me.DataList(3)) + _
                                            CDbl(Me.DataList(4)) + _
                                            CDbl(Me.DataList(5))) / 1000).ToString)
            strSQL = strSQL.Replace("@H02", ((CDbl(Me.DataList(6)) + _
                                            CDbl(Me.DataList(7)) + _
                                            CDbl(Me.DataList(8)) + _
                                            CDbl(Me.DataList(9)) + _
                                            CDbl(Me.DataList(10)) + _
                                            CDbl(Me.DataList(11))) / 1000).ToString)
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
            MessageBox.Show("[BG_T_UPLOAD_DATA.Insert001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Update exist data
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_UPLOAD_DATA", "UPDATE001")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@BudgetOrder", Me.BudgetOrder)
            strSQL = strSQL.Replace("@DataType", Me.DataType)
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
            strSQL = strSQL.Replace("@H01", ((CDbl(Me.DataList(0)) + _
                                            CDbl(Me.DataList(1)) + _
                                            CDbl(Me.DataList(2)) + _
                                            CDbl(Me.DataList(3)) + _
                                            CDbl(Me.DataList(4)) + _
                                            CDbl(Me.DataList(5))) / 1000).ToString)
            strSQL = strSQL.Replace("@H02", ((CDbl(Me.DataList(6)) + _
                                            CDbl(Me.DataList(7)) + _
                                            CDbl(Me.DataList(8)) + _
                                            CDbl(Me.DataList(9)) + _
                                            CDbl(Me.DataList(10)) + _
                                            CDbl(Me.DataList(11))) / 1000).ToString)
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
            MessageBox.Show("[BG_T_UPLOAD_DATA.Update001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

