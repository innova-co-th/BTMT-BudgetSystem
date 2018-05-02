Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Windows

Public Class frmBG0201

#Region "Variable"
    'Private myClsBG0201BL As New clsBG0201BL
    Private myComment As String = String.Empty
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

#End Region
#Region "Control Event"
    Private Sub frmBG0201_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Dim dtSource As DataTable = Nothing
        Try
            'Initial all Controls
            'InitialControls()

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

        'Try
        '    If myClsBG0201BL.GetAllRemark() AndAlso myClsBG0201BL.DtResult.Rows.Count > 0 Then
        '        result = myClsBG0201BL.DtResult
        '    End If
        'Catch ex As Exception
        '    Throw ex
        'End Try

        Return result
    End Function

    Private Sub BindDatagrid(ByVal pSource As DataTable)
        'Try
        '    grvBudget1.AutoGenerateColumns = False
        '    grvBudget1.DataSource = pSource
        'Catch ex As Exception
        '    Throw ex
        'End Try
    End Sub

    Private Sub cmdOK_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdOK.Click
        'Try
        '    If grvBudget1.SelectedCells.Count = 1 Then
        '        Remark = CStr(grvBudget1.SelectedCells(grvBudget1.SelectedCells.Count - 1).Value)
        '    End If
        '    DialogResult = DialogResult.OK
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message)
        'End Try
    End Sub

    Private Sub cmdClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdClose.Click
        Try
            Comment = ""
            DialogResult = DialogResult.Cancel
            Me.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region

  
End Class