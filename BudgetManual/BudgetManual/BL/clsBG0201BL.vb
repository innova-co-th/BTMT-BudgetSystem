Imports System.Data.SqlClient
Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class clsBG0201BL

#Region "Variable"

    Private myBudgetYear As String
    Private myPeriodType As String
    Private myBudgetOrderNo As String
    Private myRevNo As String
    Private myProjectNo As String
    Private myMonthNo As String
    Private myRRTNo As String
    Private myUserId As String = String.Empty
    Private myComment As String = String.Empty

    Private myCommentList As DataTable
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

#Region "Comment"
    Property Comment() As String
        Get
            Return myComment
        End Get
        Set(ByVal value As String)
            myComment = value
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


#End Region

#Region "Function"
    Public Function SearchComment() As Boolean
        Dim clsBG_T_BUDGET_COMMENT As New BG_T_BUDGET_COMMENT
        Dim rtn As Boolean

        '// Set Parameters
        clsBG_T_BUDGET_COMMENT.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_COMMENT.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_COMMENT.BudgetOrderNo = Me.BudgetOrderNo
        clsBG_T_BUDGET_COMMENT.RevNo = Me.RevNo
        clsBG_T_BUDGET_COMMENT.ProjectNo = Me.ProjectNo


        rtn = clsBG_T_BUDGET_COMMENT.Select001()


        '// Call Function
        If rtn = True Then
            Me.CommentList = clsBG_T_BUDGET_COMMENT.CommentList

            Return True
        Else
            Me.CommentList = Nothing

            Return False
        End If
    End Function

    Public Function CreateNewComment() As Boolean
        Dim clsBG_T_BUDGET_COMMENT As New BG_T_BUDGET_COMMENT

        '// Set Parameters
        clsBG_T_BUDGET_COMMENT.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_COMMENT.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_COMMENT.BudgetOrderNo = Me.BudgetOrderNo
        clsBG_T_BUDGET_COMMENT.RevNo = Me.RevNo
        clsBG_T_BUDGET_COMMENT.ProjectNo = Me.ProjectNo
        clsBG_T_BUDGET_COMMENT.MonthNo = Me.MonthNo
        clsBG_T_BUDGET_COMMENT.RRTNo = Me.RRTNo
        clsBG_T_BUDGET_COMMENT.Comment = Me.Comment
        clsBG_T_BUDGET_COMMENT.CreateUserId = p_strUserId

        '// Call Function: Insert New Comment
        If clsBG_T_BUDGET_COMMENT.Insert001() = True Then

            Return True
        Else
            Return False

        End If
    End Function

    Public Function UpdateComment() As Boolean
        Dim clsBG_T_BUDGET_COMMENT As New BG_T_BUDGET_COMMENT

        '// Set Parameters
        clsBG_T_BUDGET_COMMENT.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_COMMENT.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_COMMENT.BudgetOrderNo = Me.BudgetOrderNo
        clsBG_T_BUDGET_COMMENT.RevNo = Me.RevNo
        clsBG_T_BUDGET_COMMENT.ProjectNo = Me.ProjectNo
        clsBG_T_BUDGET_COMMENT.MonthNo = Me.MonthNo
        clsBG_T_BUDGET_COMMENT.RRTNo = Me.RRTNo
        clsBG_T_BUDGET_COMMENT.Comment = Me.Comment
        clsBG_T_BUDGET_COMMENT.CreateUserId = p_strUserId

        '// Call Function: Insert New Comment
        If clsBG_T_BUDGET_COMMENT.Update001() = True Then

            Return True
        Else
            Return False

        End If
    End Function


#End Region

End Class
