Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.SqlClient

Public Class clsBG0670BL

#Region "Variable"
    Private myDtResult As DataTable
    Private myDtReference As DataTable
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
    Private myUpdateUserID As String = String.Empty
    Private myUpdateDate As String = String.Empty
    Private myRefBudgetYear As String = String.Empty
    Private myRefProjectNo As String = String.Empty
    Private myRefPeriodType As String = String.Empty
    Private myRefRevNo As String = String.Empty
    Private myRefBudgetYear2 As String = String.Empty
    Private myRefProjectNo2 As String = String.Empty
    Private myRefPeriodType2 As String = String.Empty
    Private myRefRevNo2 As String = String.Empty
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

#Region "RefPeriodType"
    Public Property RefPeriodType() As String
        Get
            Return myRefPeriodType
        End Get
        Set(ByVal value As String)
            myRefPeriodType = value
        End Set
    End Property
#End Region

#Region "RefBudgetYear"
    Public Property RefBudgetYear() As String
        Get
            Return myRefBudgetYear
        End Get
        Set(ByVal value As String)
            myRefBudgetYear = value
        End Set
    End Property
#End Region

#Region "RefProjectNo"
    Public Property RefProjectNo() As String
        Get
            Return myRefProjectNo
        End Get
        Set(ByVal value As String)
            myRefProjectNo = value
        End Set
    End Property
#End Region

#Region "RefRevNo"
    Public Property RefRevNo() As String
        Get
            Return myRefRevNo
        End Get
        Set(ByVal value As String)
            myRefRevNo = value
        End Set
    End Property
#End Region

#Region "RefPeriodType2"
    Public Property RefPeriodType2() As String
        Get
            Return myRefPeriodType2
        End Get
        Set(ByVal value As String)
            myRefPeriodType2 = value
        End Set
    End Property
#End Region

#Region "RefBudgetYear2"
    Public Property RefBudgetYear2() As String
        Get
            Return myRefBudgetYear2
        End Get
        Set(ByVal value As String)
            myRefBudgetYear2 = value
        End Set
    End Property
#End Region

#Region "RefProjectNo2"
    Public Property RefProjectNo2() As String
        Get
            Return myRefProjectNo2
        End Get
        Set(ByVal value As String)
            myRefProjectNo2 = value
        End Set
    End Property
#End Region

#Region "RefRevNo2"
    Public Property RefRevNo2() As String
        Get
            Return myRefRevNo2
        End Get
        Set(ByVal value As String)
            myRefRevNo2 = value
        End Set
    End Property
#End Region


#Region "dtReference"
    Property dtReference() As DataTable
        Get
            Return myDtReference
        End Get
        Set(ByVal value As DataTable)
            myDtReference = value
        End Set
    End Property
#End Region

#End Region

#Region "Function"

    Public Function GetBudgetYear() As Boolean

        Dim clsBG_T_BUDGET_ADJUST As BG_T_BUDGET_ADJUST = New BG_T_BUDGET_ADJUST()

        If clsBG_T_BUDGET_ADJUST.Select001() = False Then
            clsBG_T_BUDGET_ADJUST = Nothing
            Me.dtResult = Nothing
            Return False
        End If

        Me.dtResult = clsBG_T_BUDGET_ADJUST.dtResult
        clsBG_T_BUDGET_ADJUST = Nothing
        Return True

    End Function

    Public Function GetBudgetPeriod() As Boolean

        Dim clsBG_T_BUDGET_ADJUST As BG_T_BUDGET_ADJUST = New BG_T_BUDGET_ADJUST()

        clsBG_T_BUDGET_ADJUST.BudgetYear = Me.BudgetYear

        If clsBG_T_BUDGET_ADJUST.Select002() = False Then
            clsBG_T_BUDGET_ADJUST = Nothing
            Me.dtResult = Nothing
            Return False
        End If

        Me.dtResult = clsBG_T_BUDGET_ADJUST.dtResult
        clsBG_T_BUDGET_ADJUST = Nothing
        Return True

    End Function

    Public Function GetProjectNo() As Boolean

        Dim clsBG_T_BUDGET_ADJUST As BG_T_BUDGET_ADJUST = New BG_T_BUDGET_ADJUST()

        clsBG_T_BUDGET_ADJUST.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_ADJUST.PeriodType = Me.PeriodType

        If clsBG_T_BUDGET_ADJUST.Select007() = False Then
            clsBG_T_BUDGET_ADJUST = Nothing
            Me.dtResult = Nothing
            Return False
        End If

        Me.dtResult = clsBG_T_BUDGET_ADJUST.dtResult
        clsBG_T_BUDGET_ADJUST = Nothing
        Return True

    End Function

    Public Function GetRevNo() As Boolean

        Dim clsBG_T_BUDGET_ADJUST As BG_T_BUDGET_ADJUST = New BG_T_BUDGET_ADJUST()

        clsBG_T_BUDGET_ADJUST.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_ADJUST.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_ADJUST.ProjectNo = Me.ProjectNo

        If clsBG_T_BUDGET_ADJUST.Select003() = False Then
            clsBG_T_BUDGET_ADJUST = Nothing
            Me.dtResult = Nothing
            Return False
        End If

        Me.dtResult = clsBG_T_BUDGET_ADJUST.dtResult
        clsBG_T_BUDGET_ADJUST = Nothing
        Return True

    End Function

    Public Function GetBudgetAdjust() As Boolean

        Dim clsBG_T_BUDGET_ADJUST As BG_T_BUDGET_ADJUST = New BG_T_BUDGET_ADJUST()

        clsBG_T_BUDGET_ADJUST.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_ADJUST.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_ADJUST.RevNo = Me.RevNo
        clsBG_T_BUDGET_ADJUST.ProjectNo = Me.ProjectNo

        If clsBG_T_BUDGET_ADJUST.Select004() = False Then
            clsBG_T_BUDGET_ADJUST = Nothing
            Me.dtResult = Nothing
            Return False
        End If

        Me.dtResult = clsBG_T_BUDGET_ADJUST.dtResult
        clsBG_T_BUDGET_ADJUST = Nothing
        Return True

    End Function

    Public Function GetBudgetReference() As Boolean

        Dim clsBG_T_BUDGET_REFERENCE As BG_T_BUDGET_REFERENCE = New BG_T_BUDGET_REFERENCE()

        clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
        clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
        clsBG_T_BUDGET_REFERENCE.RefPeriodType = Me.RefPeriodType

        If clsBG_T_BUDGET_REFERENCE.Select001() = False Then
            clsBG_T_BUDGET_REFERENCE = Nothing
            Me.dtReference = Nothing
            Return False
        End If

        Me.dtReference = clsBG_T_BUDGET_REFERENCE.dtResult
        clsBG_T_BUDGET_REFERENCE = Nothing
        Return True

    End Function

    Public Function UpdateBudgetAdjust() As Boolean

        Dim clsBG_T_BUDGET_ADJUST As New BG_T_BUDGET_ADJUST
        Dim clsBG_T_BUDGET_REFERENCE As New BG_T_BUDGET_REFERENCE

        Dim conn As SqlConnection
        Dim trans As SqlTransaction
        Dim bResult As Boolean = False

        conn = New SqlConnection(My.Settings.ConnStr)
        conn.Open()
        trans = conn.BeginTransaction()

        Try
            clsBG_T_BUDGET_ADJUST.BudgetYear = Me.BudgetYear
            clsBG_T_BUDGET_ADJUST.PeriodType = Me.PeriodType
            clsBG_T_BUDGET_ADJUST.RevNo = Me.RevNo
            clsBG_T_BUDGET_ADJUST.RRT0 = Me.RRT0
            clsBG_T_BUDGET_ADJUST.RRT1 = Me.RRT1
            clsBG_T_BUDGET_ADJUST.RRT2 = Me.RRT2
            clsBG_T_BUDGET_ADJUST.RRT3 = Me.RRT3
            clsBG_T_BUDGET_ADJUST.RRT4 = Me.RRT4
            clsBG_T_BUDGET_ADJUST.RRT5 = Me.RRT5
            clsBG_T_BUDGET_ADJUST.WorkingBG1 = Me.WorkingBG1
            clsBG_T_BUDGET_ADJUST.WorkingBG2 = Me.WorkingBG2
            clsBG_T_BUDGET_ADJUST.UpdateUserID = Me.UpdateUserID
            clsBG_T_BUDGET_ADJUST.UpdateDate = Now.ToString("yyyy-MM-dd HH:mm")
            clsBG_T_BUDGET_ADJUST.ProjectNo = Me.ProjectNo

            If clsBG_T_BUDGET_ADJUST.Update001(conn, trans) = True Then
                bResult = True
            End If

            '// Add(Update) Reference1
            If bResult = True Then
                If Me.RefBudgetYear <> "" AndAlso _
                    Me.RefPeriodType <> "" AndAlso _
                    Me.RefRevNo <> "" AndAlso _
                    Me.RefProjectNo <> "" Then

                    '// Check data exist
                    clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
                    clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
                    clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
                    clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
                    clsBG_T_BUDGET_REFERENCE.RefPeriodType = Me.RefPeriodType
                    clsBG_T_BUDGET_REFERENCE.RefBudgetYear = Me.RefBudgetYear
                    clsBG_T_BUDGET_REFERENCE.RefProjectNo = Me.RefProjectNo
                    clsBG_T_BUDGET_REFERENCE.RefRevNo = Me.RefRevNo
                    clsBG_T_BUDGET_REFERENCE.CreateUserID = Me.UpdateUserID
                    If clsBG_T_BUDGET_REFERENCE.Select001() = True Then

                        If clsBG_T_BUDGET_REFERENCE.dtResult.Rows.Count = 0 Then
                            '// Add new record
                            bResult = clsBG_T_BUDGET_REFERENCE.Insert001(conn, trans)

                        Else
                            '// Update exists record
                            bResult = clsBG_T_BUDGET_REFERENCE.Update001(conn, trans)

                        End If

                    Else
                        bResult = False

                    End If


                End If
            End If


            '// Add(Update) Reference2
            If bResult = True Then
                If Me.RefBudgetYear2 <> "" AndAlso _
                    Me.RefPeriodType2 <> "" AndAlso _
                    Me.RefRevNo2 <> "" AndAlso _
                    Me.RefProjectNo2 <> "" Then

                    '// Check data exist
                    clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
                    clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
                    clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
                    clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
                    clsBG_T_BUDGET_REFERENCE.RefPeriodType = Me.RefPeriodType2
                    clsBG_T_BUDGET_REFERENCE.RefBudgetYear = Me.RefBudgetYear2
                    clsBG_T_BUDGET_REFERENCE.RefProjectNo = Me.RefProjectNo2
                    clsBG_T_BUDGET_REFERENCE.RefRevNo = Me.RefRevNo2
                    clsBG_T_BUDGET_REFERENCE.CreateUserID = Me.UpdateUserID
                    If clsBG_T_BUDGET_REFERENCE.Select001() = True Then

                        If clsBG_T_BUDGET_REFERENCE.dtResult.Rows.Count = 0 Then
                            '// Add new record
                            bResult = clsBG_T_BUDGET_REFERENCE.Insert001(conn, trans)

                        Else
                            '// Update exists record
                            bResult = clsBG_T_BUDGET_REFERENCE.Update001(conn, trans)

                        End If

                    Else
                        bResult = False

                    End If


                End If
            End If






            If bResult Then
                trans.Commit()
            Else
                trans.Rollback()
            End If
        Catch ex As Exception
            trans.Rollback()
            showErrorMessage("Error clsBG0670BL.UpdateBudgetAdjust: " & ex.Message)
        Finally
            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If
        End Try

        Return bResult

    End Function

#End Region

End Class
