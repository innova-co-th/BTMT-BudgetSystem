Imports System.Data.SqlClient
Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class BG_T_COMMENT_DATA

#Region "Variable"
    Private Const STRING_ALL As String = "All"

    Private myDtResult As DataTable
    Private myBudgetYear As String = String.Empty
    Private myPeriodType As String = String.Empty
    Private myUserPIC As String = String.Empty
    Private myProjectNo As String = String.Empty
    Private myRevNo As String = String.Empty
    Private myBudgetType As String = String.Empty
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

#End Region

#Region "Function"
    ''' <summary>
    ''' Select Comment data list for Report008 All
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

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_T_COMMENT_DATA", "SELECT001")
            strSQL = strSQL.Replace("@YEAR", Me.BudgetYear)
            strSQL = strSQL.Replace("@PERIOD", Me.PeriodType)
            strSQL = strSQL.Replace("@PROJECTNO", Me.ProjectNo)
            strSQL = strSQL.Replace("@REVNO", Me.RevNo)
            strSQL = strSQL.Replace("@BUDGETTYPE", Me.BudgetType)

            If Me.UserPIC = "0" Then
                strSQL = strSQL.Replace("@PICCONDITION", " ")
                strSQL = strSQL.Replace("@REVNOCONDITION", "AND BC.REV_NO = @REVNO")
                strSQL = strSQL.Replace("@REVNO", Me.RevNo)
            Else
                strSQL = strSQL.Replace("@REVNOCONDITION", " ")
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
            MessageBox.Show("[BG_T_COMMENT_DATA.Select001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

End Class
