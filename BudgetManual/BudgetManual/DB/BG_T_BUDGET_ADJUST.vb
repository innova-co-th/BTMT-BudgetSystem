Imports System.Data.SqlClient
Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class BG_T_BUDGET_ADJUST

#Region "Variable"
    Private myDtResult As DataTable
    Private myBudgetYear As String = String.Empty
    Private myProjectNo As String = String.Empty
    Private myPeriodType As String = String.Empty
    Private myRevNo As String = String.Empty
    Private myRRT0 As String = String.Empty
    Private myRRT1 As String = String.Empty
    Private myRRT2 As String = String.Empty
    Private myRRT3 As String = String.Empty
    Private myRRT4 As String = String.Empty
    Private myRRT5 As String = String.Empty
    Private myWorkingBG1 As String = String.Empty
    Private myWorkingBG2 As String = String.Empty
    Private myCreateUserID As String = String.Empty
    Private myUpdateUserID As String = String.Empty
    Private myUpdateDate As String = String.Empty
    Private myWKH1 As String = String.Empty
    Private myWKH2 As String = String.Empty
    Private myMTP_SUM1 As String = String.Empty
    Private myMTP_SUM2 As String = String.Empty
    Private myMTP_SUM3 As String = String.Empty
    Private myMTP_SUM4 As String = String.Empty
    Private myMTP_SUM5 As String = String.Empty
    Private myWKRRT1 As String = String.Empty
    Private myWKRRT2 As String = String.Empty
    Private myWKRRT3 As String = String.Empty
    Private myWKRRT4 As String = String.Empty
    Private myWKRRT5 As String = String.Empty
    Private myMTP_PY_SUM1 As String = String.Empty
    Private myMTP_PY_SUM2 As String = String.Empty
    Private myMTP_PY_SUM3 As String = String.Empty
    Private myMTP_PY_SUM4 As String = String.Empty
    Private myMTP_PY_SUM5 As String = String.Empty
    Private myMTPWB As String = String.Empty
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

#Region "RRT0"
    Public Property RRT0() As String
        Get
            Return myRRT0
        End Get
        Set(ByVal value As String)
            myRRT0 = value
        End Set
    End Property
#End Region

#Region "RRT1"
    Public Property RRT1() As String
        Get
            Return myRRT1
        End Get
        Set(ByVal value As String)
            myRRT1 = value
        End Set
    End Property
#End Region

#Region "RRT2"
    Public Property RRT2() As String
        Get
            Return myRRT2
        End Get
        Set(ByVal value As String)
            myRRT2 = value
        End Set
    End Property
#End Region

#Region "RRT3"
    Public Property RRT3() As String
        Get
            Return myRRT3
        End Get
        Set(ByVal value As String)
            myRRT3 = value
        End Set
    End Property
#End Region

#Region "RRT4"
    Public Property RRT4() As String
        Get
            Return myRRT4
        End Get
        Set(ByVal value As String)
            myRRT4 = value
        End Set
    End Property
#End Region

#Region "RRT5"
    Public Property RRT5() As String
        Get
            Return myRRT5
        End Get
        Set(ByVal value As String)
            myRRT5 = value
        End Set
    End Property
#End Region

#Region "WorkingBG1"
    Public Property WorkingBG1() As String
        Get
            Return myWorkingBG1
        End Get
        Set(ByVal value As String)
            myWorkingBG1 = value
        End Set
    End Property
#End Region

#Region "WorkingBG2"
    Public Property WorkingBG2() As String
        Get
            Return myWorkingBG2
        End Get
        Set(ByVal value As String)
            myWorkingBG2 = value
        End Set
    End Property
#End Region

#Region "CreateUserID"
    Public Property CreateUserID() As String
        Get
            Return myCreateUserID
        End Get
        Set(ByVal value As String)
            myCreateUserID = value
        End Set
    End Property
#End Region

#Region "UpdateUserID"
    Public Property UpdateUserID() As String
        Get
            Return myUpdateUserID
        End Get
        Set(ByVal value As String)
            myUpdateUserID = value
        End Set
    End Property
#End Region

#Region "UpdateDate"
    Public Property UpdateDate() As String
        Get
            Return myUpdateDate
        End Get
        Set(ByVal value As String)
            myUpdateDate = value
        End Set
    End Property
#End Region

#Region "WKH1"
    Property WKH1() As String
        Get
            Return myWKH1
        End Get
        Set(ByVal value As String)
            myWKH1 = value
        End Set
    End Property
#End Region

#Region "WKH2"
    Property WKH2() As String
        Get
            Return myWKH2
        End Get
        Set(ByVal value As String)
            myWKH2 = value
        End Set
    End Property
#End Region

#Region "MTP_SUM1"
    Property MTP_SUM1() As String
        Get
            Return myMTP_SUM1
        End Get
        Set(ByVal value As String)
            myMTP_SUM1 = value
        End Set
    End Property
#End Region

#Region "MTP_SUM2"
    Property MTP_SUM2() As String
        Get
            Return myMTP_SUM2
        End Get
        Set(ByVal value As String)
            myMTP_SUM2 = value
        End Set
    End Property
#End Region

#Region "MTP_SUM3"
    Property MTP_SUM3() As String
        Get
            Return myMTP_SUM3
        End Get
        Set(ByVal value As String)
            myMTP_SUM3 = value
        End Set
    End Property
#End Region

#Region "MTP_SUM4"
    Property MTP_SUM4() As String
        Get
            Return myMTP_SUM4
        End Get
        Set(ByVal value As String)
            myMTP_SUM4 = value
        End Set
    End Property
#End Region

#Region "MTP_SUM5"
    Property MTP_SUM5() As String
        Get
            Return myMTP_SUM5
        End Get
        Set(ByVal value As String)
            myMTP_SUM5 = value
        End Set
    End Property
#End Region


#Region "WKRRT1"
    Property WKRRT1() As String
        Get
            Return myWKRRT1
        End Get
        Set(ByVal value As String)
            myWKRRT1 = value
        End Set
    End Property
#End Region

#Region "WKRRT2"
    Property WKRRT2() As String
        Get
            Return myWKRRT2
        End Get
        Set(ByVal value As String)
            myWKRRT2 = value
        End Set
    End Property
#End Region

#Region "WKRRT3"
    Property WKRRT3() As String
        Get
            Return myWKRRT3
        End Get
        Set(ByVal value As String)
            myWKRRT3 = value
        End Set
    End Property
#End Region

#Region "WKRRT4"
    Property WKRRT4() As String
        Get
            Return myWKRRT4
        End Get
        Set(ByVal value As String)
            myWKRRT4 = value
        End Set
    End Property
#End Region

#Region "WKRRT5"
    Property WKRRT5() As String
        Get
            Return myWKRRT5
        End Get
        Set(ByVal value As String)
            myWKRRT5 = value
        End Set
    End Property
#End Region

#Region "MTPWB"
    Property MTPWB() As String
        Get
            Return myMTPWB
        End Get
        Set(ByVal value As String)
            myMTPWB = value
        End Set
    End Property
#End Region

#Region "MTP_PY_SUM1"
    Property MTP_PY_SUM1() As String
        Get
            Return myMTP_PY_SUM1
        End Get
        Set(ByVal value As String)
            myMTP_PY_SUM1 = value
        End Set
    End Property
#End Region

#Region "MTP_PY_SUM2"
    Property MTP_PY_SUM2() As String
        Get
            Return myMTP_PY_SUM2
        End Get
        Set(ByVal value As String)
            myMTP_PY_SUM2 = value
        End Set
    End Property
#End Region

#Region "MTP_PY_SUM3"
    Property MTP_PY_SUM3() As String
        Get
            Return myMTP_PY_SUM3
        End Get
        Set(ByVal value As String)
            myMTP_PY_SUM3 = value
        End Set
    End Property
#End Region

#Region "MTP_PY_SUM4"
    Property MTP_PY_SUM4() As String
        Get
            Return myMTP_PY_SUM4
        End Get
        Set(ByVal value As String)
            myMTP_PY_SUM4 = value
        End Set
    End Property
#End Region

#Region "MTP_PY_SUM5"
    Property MTP_PY_SUM5() As String
        Get
            Return myMTP_PY_SUM5
        End Get
        Set(ByVal value As String)
            myMTP_PY_SUM5 = value
        End Set
    End Property
#End Region

#End Region

#Region "Function"

#Region "Select001"
    ''' <summary>
    ''' Select budget year list
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_ADJUST", "SELECT001")

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception

            MessageBox.Show("[BG_T_BUDGET_ADJUST.Select001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Select period type list according budget year
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_ADJUST", "SELECT002")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception

            MessageBox.Show("[BG_T_BUDGET_ADJUST.Select002] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Select REV list according budget year, period type
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_ADJUST", "SELECT003")
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

            MessageBox.Show("[BG_T_BUDGET_ADJUST.Select003] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Select budget according budget year, period type, rev_no
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_ADJUST", "SELECT004")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@REV", CStr(IIf(IsNumeric(Me.RevNo), Me.RevNo, "0")))
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

            MessageBox.Show("[BG_T_BUDGET_ADJUST.Select004] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Select Working budget H1, H2
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_ADJUST", "SELECT005")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@REV", CStr(IIf(IsNumeric(Me.RevNo), Me.RevNo, "0")))
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

            MessageBox.Show("[BG_T_BUDGET_ADJUST.Select005] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Select MTP Summary
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_ADJUST", "SELECT006")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@REV", CStr(IIf(IsNumeric(Me.RevNo), Me.RevNo, "0")))
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

            MessageBox.Show("[BG_T_BUDGET_ADJUST.Select006] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Select REV list according budget year, period type
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_ADJUST", "SELECT007")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception

            MessageBox.Show("[BG_T_BUDGET_ADJUST.Select007] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Select MTP Summary
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_ADJUST", "SELECT008")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@REV", CStr(IIf(IsNumeric(Me.RevNo), Me.RevNo, "0")))
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

            MessageBox.Show("[BG_T_BUDGET_ADJUST.Select008] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Select Reference
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_ADJUST", "SELECT009")
            strSQL = strSQL.Replace("@BudgetYear", Me.BudgetYear)
            strSQL = strSQL.Replace("@PeriodType", Me.PeriodType)
            strSQL = strSQL.Replace("@RevNo", CStr(IIf(IsNumeric(Me.RevNo), Me.RevNo, "0")))
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

            MessageBox.Show("[BG_T_BUDGET_ADJUST.Select009] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try

    End Function

#End Region

#Region "Update001"
    ''' <summary>
    ''' Update Budget Adjust Master
    ''' </summary>
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_ADJUST", "UPDATE001")
            strSQL = strSQL.Replace("@RRT0", Me.RRT0)
            strSQL = strSQL.Replace("@RRT1", Me.RRT1)
            strSQL = strSQL.Replace("@RRT2", Me.RRT2)
            strSQL = strSQL.Replace("@RRT3", Me.RRT3)
            strSQL = strSQL.Replace("@RRT4", Me.RRT4)
            strSQL = strSQL.Replace("@RRT5", Me.RRT5)
            strSQL = strSQL.Replace("@WORKING_BG1", Me.WorkingBG1)
            strSQL = strSQL.Replace("@WORKING_BG2", Me.WorkingBG2)
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@REV", Me.RevNo)
            strSQL = strSQL.Replace("@UPDATE_USER_ID", Me.UpdateUserID)
            strSQL = strSQL.Replace("@PROJECTNO", Me.ProjectNo)

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

            MessageBox.Show("[BG_T_BUDGET_ADJUST.Update001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Update Working budget H1, H2 in report
    ''' </summary>
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_ADJUST", "UPDATE002")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@REV", Me.RevNo)
            strSQL = strSQL.Replace("@WKH1", Me.WKH1)
            strSQL = strSQL.Replace("@WKH2", Me.WKH2)
            strSQL = strSQL.Replace("@UPDATE_USER_ID", Me.UpdateUserID)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)
            strSQL = strSQL.Replace("@WKRRT1", Me.WKRRT1)
            strSQL = strSQL.Replace("@WKRRT2", Me.WKRRT2)
            strSQL = strSQL.Replace("@WKRRT3", Me.WKRRT3)
            strSQL = strSQL.Replace("@WKRRT4", Me.WKRRT4)
            strSQL = strSQL.Replace("@WKRRT5", Me.WKRRT5)
            strSQL = strSQL.Replace("@MTPWB", Me.MTPWB)

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

            MessageBox.Show("[BG_T_BUDGET_ADJUST.Update002] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Update Working budget MTP Summary in report
    ''' </summary>
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_ADJUST", "UPDATE003")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@REV", Me.RevNo)
            strSQL = strSQL.Replace("@MTP_SUM1", Me.MTP_SUM1)
            strSQL = strSQL.Replace("@MTP_SUM2", Me.MTP_SUM2)
            strSQL = strSQL.Replace("@MTP_SUM3", Me.MTP_SUM3)
            strSQL = strSQL.Replace("@MTP_SUM4", Me.MTP_SUM4)
            strSQL = strSQL.Replace("@MTP_SUM5", Me.MTP_SUM5)
            strSQL = strSQL.Replace("@UPDATE_USER_ID", Me.UpdateUserID)
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

            MessageBox.Show("[BG_T_BUDGET_ADJUST.Update003] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Update Investment for MTP budget
    ''' </summary>
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_ADJUST", "UPDATE004")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@REV", Me.RevNo)
            strSQL = strSQL.Replace("@MTP_SUM1", Me.MTP_SUM1)
            strSQL = strSQL.Replace("@MTP_SUM2", Me.MTP_SUM2)
            strSQL = strSQL.Replace("@MTP_SUM3", Me.MTP_SUM3)
            strSQL = strSQL.Replace("@MTP_SUM4", Me.MTP_SUM4)
            strSQL = strSQL.Replace("@MTP_SUM5", Me.MTP_SUM5)
            strSQL = strSQL.Replace("@MTP_PY_SUM1", Me.MTP_PY_SUM1)
            strSQL = strSQL.Replace("@MTP_PY_SUM2", Me.MTP_PY_SUM2)
            strSQL = strSQL.Replace("@MTP_PY_SUM3", Me.MTP_PY_SUM3)
            strSQL = strSQL.Replace("@MTP_PY_SUM4", Me.MTP_PY_SUM4)
            strSQL = strSQL.Replace("@MTP_PY_SUM5", Me.MTP_PY_SUM5)
            strSQL = strSQL.Replace("@UPDATE_USER_ID", Me.UpdateUserID)
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

            MessageBox.Show("[BG_T_BUDGET_ADJUST.Update004] Error: " & ex.Message, My.Settings.ProgramTitle, _
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

#Region "Insert001"
    ''' <summary>
    ''' Insert Budget Adjust Master
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_ADJUST", "INSERT001")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@REV", Me.RevNo)
            strSQL = strSQL.Replace("@RRT0", Me.RRT0)
            strSQL = strSQL.Replace("@RRT1", Me.RRT1)
            strSQL = strSQL.Replace("@RRT2", Me.RRT2)
            strSQL = strSQL.Replace("@RRT3", Me.RRT3)
            strSQL = strSQL.Replace("@RRT4", Me.RRT4)
            strSQL = strSQL.Replace("@RRT5", Me.RRT5)
            strSQL = strSQL.Replace("@WORKING_BG1", Me.WorkingBG1)
            strSQL = strSQL.Replace("@WORKING_BG2", Me.WorkingBG2)
            strSQL = strSQL.Replace("@CREATE_USER_ID", Me.CreateUserID)
            strSQL = strSQL.Replace("@UPDATE_USER_ID", Me.CreateUserID)
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
            MessageBox.Show("[BG_T_BUDGET_ADJUST.Insert001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Insert Budget Adjust Master2
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_ADJUST", "INSERT002")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@REV", Me.RevNo)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            strSQL = strSQL.Replace("@MTP_SUM1", Me.MTP_SUM1)
            strSQL = strSQL.Replace("@MTP_SUM2", Me.MTP_SUM2)
            strSQL = strSQL.Replace("@MTP_SUM3", Me.MTP_SUM3)
            strSQL = strSQL.Replace("@MTP_SUM4", Me.MTP_SUM4)
            strSQL = strSQL.Replace("@MTP_SUM5", Me.MTP_SUM5)

            strSQL = strSQL.Replace("@CREATE_USER_ID", Me.CreateUserID)
            strSQL = strSQL.Replace("@UPDATE_USER_ID", Me.CreateUserID)

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
            MessageBox.Show("[BG_T_BUDGET_ADJUST.Insert002] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Insert Budget Adjust Master2
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_ADJUST", "INSERT003")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@REV", Me.RevNo)
            strSQL = strSQL.Replace("@ProjectNo", Me.ProjectNo)

            strSQL = strSQL.Replace("@MTP_SUM1", Me.MTP_SUM1)
            strSQL = strSQL.Replace("@MTP_SUM2", Me.MTP_SUM2)
            strSQL = strSQL.Replace("@MTP_SUM3", Me.MTP_SUM3)
            strSQL = strSQL.Replace("@MTP_SUM4", Me.MTP_SUM4)
            strSQL = strSQL.Replace("@MTP_SUM5", Me.MTP_SUM5)

            strSQL = strSQL.Replace("@MTP_PY_SUM1", Me.MTP_PY_SUM1)
            strSQL = strSQL.Replace("@MTP_PY_SUM2", Me.MTP_PY_SUM2)
            strSQL = strSQL.Replace("@MTP_PY_SUM3", Me.MTP_PY_SUM3)
            strSQL = strSQL.Replace("@MTP_PY_SUM4", Me.MTP_PY_SUM4)
            strSQL = strSQL.Replace("@MTP_PY_SUM5", Me.MTP_PY_SUM5)

            strSQL = strSQL.Replace("@CREATE_USER_ID", Me.CreateUserID)
            strSQL = strSQL.Replace("@UPDATE_USER_ID", Me.CreateUserID)

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
            MessageBox.Show("[BG_T_BUDGET_ADJUST.Insert003] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Delete Adjust master
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_ADJUST", "DELETE001")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@REV", Me.RevNo)
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
            MessageBox.Show("[BG_T_BUDGET_ADJUST.Delete001] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
    ''' Delete Adjust master2
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_BUDGET_ADJUST", "DELETE002")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@REV", Me.RevNo)
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
            MessageBox.Show("[BG_T_BUDGET_ADJUST.Delete002] Error: " & ex.Message, My.Settings.ProgramTitle, _
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
