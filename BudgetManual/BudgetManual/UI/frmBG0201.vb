Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Windows

Public Class frmBG0201

#Region "Variable"

    Private myClsBG0201BL As New clsBG0201BL

    Private myComment As String

    Private myBudgetYear As String
    Private myPeriodType As String
    Private myBudgetOrderNo As String
    Private myRevNo As String
    Private myProjectNo As String
    Private myMonthNo As String
    Private myRRTNo As String

    Private myOperationCd As String


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

#Region "OperationCd"
    Public Property OperationCd() As String
        Get
            Return myOperationCd
        End Get
        Set(ByVal value As String)
            myOperationCd = value
        End Set
    End Property
#End Region

#End Region

#Region "Control Event"
    Private Sub frmBG0201_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Dim dtSource As DataTable = Nothing
        Try
            'Initial all Controls
            'InitialControls()

            If Me.OperationCd = CStr(enumOperationCd.InputBudget) Or _
            (Me.OperationCd = CStr(enumOperationCd.AdjustBudget) And CInt(Me.RevNo) > 1) Or _
            (Me.OperationCd = CStr(enumOperationCd.AdjustBudgetDirectInput) And CInt(Me.RevNo) > 1) Then
                Me.cmdOK.Visible = True
            Else
                Me.cmdOK.Visible = False
            End If

            dtSource = GetComment()
            BindDatagrid(dtSource)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region

#Region "Function"
    Private Function GetComment() As DataTable
        Dim result As DataTable = Nothing

        Try
            myClsBG0201BL.BudgetYear = Me.BudgetYear
            myClsBG0201BL.PeriodType = Me.PeriodType
            myClsBG0201BL.BudgetOrderNo = Me.BudgetOrderNo
            myClsBG0201BL.RevNo = Me.RevNo
            myClsBG0201BL.ProjectNo = Me.ProjectNo

            If myClsBG0201BL.SearchComment AndAlso myClsBG0201BL.CommentList.Rows.Count > 0 Then
                result = myClsBG0201BL.CommentList
            End If
        Catch ex As Exception
            Throw ex
        End Try

        Return result
    End Function

    Private Sub BindDatagrid(ByVal pSource As DataTable)
        Try
            Me.Comment = ""
            If Me.PeriodType = CStr(enumPeriodType.MTPBudget) Then
                If Not pSource Is Nothing AndAlso pSource.Rows.Count > 0 Then
                    Select Case Me.RRTNo
                        Case "1"
                            Me.Comment = pSource.Rows(0).Item("RRT1").ToString
                        Case "2"
                            Me.Comment = pSource.Rows(0).Item("RRT2").ToString
                    End Select
                End If
            Else
                If Not pSource Is Nothing AndAlso pSource.Rows.Count > 0 Then
                    Select Case Me.MonthNo
                        Case "1"
                            Me.Comment = pSource.Rows(0).Item("M1").ToString
                        Case "2"
                            Me.Comment = pSource.Rows(0).Item("M2").ToString
                        Case "3"
                            Me.Comment = pSource.Rows(0).Item("M3").ToString
                        Case "4"
                            Me.Comment = pSource.Rows(0).Item("M4").ToString
                        Case "5"
                            Me.Comment = pSource.Rows(0).Item("M5").ToString
                        Case "6"
                            Me.Comment = pSource.Rows(0).Item("M6").ToString
                        Case "7"
                            Me.Comment = pSource.Rows(0).Item("M7").ToString
                        Case "8"
                            Me.Comment = pSource.Rows(0).Item("M8").ToString
                        Case "9"
                            Me.Comment = pSource.Rows(0).Item("M9").ToString
                        Case "10"
                            Me.Comment = pSource.Rows(0).Item("M10").ToString
                        Case "11"
                            Me.Comment = pSource.Rows(0).Item("M11").ToString
                        Case "12"
                            Me.Comment = pSource.Rows(0).Item("M12").ToString
                    End Select
                End If
            End If

            txtComment.Text = Me.Comment
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub cmdOK_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdOK.Click
        'Save Or Update 

        myClsBG0201BL.BudgetYear = Me.BudgetYear
        myClsBG0201BL.PeriodType = Me.PeriodType
        myClsBG0201BL.BudgetOrderNo = Me.BudgetOrderNo
        myClsBG0201BL.RevNo = Me.RevNo
        myClsBG0201BL.ProjectNo = Me.ProjectNo

        If Me.PeriodType = CStr(enumPeriodType.MTPBudget) Then
            myClsBG0201BL.RRTNo = Me.RRTNo
        Else
            myClsBG0201BL.MonthNo = Me.MonthNo
        End If


        myClsBG0201BL.Comment = txtComment.Text.Trim

        If myClsBG0201BL.SearchComment AndAlso myClsBG0201BL.CommentList.Rows.Count > 0 Then
            'Update

            If myClsBG0201BL.UpdateComment = True Then
                MessageBox.Show("Update comment completed", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                '// Write Transaction Log
                WriteTransactionLog(CStr(enumOperationCd.EditUserMaster), "", "", "", "", "", "")

                Me.Close()

            End If
        Else
            'Insert          

            If myClsBG0201BL.CreateNewComment = True Then
                MessageBox.Show("Add comment completed", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                '// Write Transaction Log
                WriteTransactionLog(CStr(enumOperationCd.EditUserMaster), "", "", "", "", "", "")

                Me.Close()

            End If

        End If

    End Sub

    Private Sub cmdClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdClose.Click
        Try

            DialogResult = DialogResult.Cancel
            Me.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region

  
End Class