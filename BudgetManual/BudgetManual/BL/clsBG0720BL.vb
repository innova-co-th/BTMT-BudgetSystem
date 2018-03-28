Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class clsBG0720BL

#Region "Variable"
    Private myPicList As DataTable
    Private myPicNo As String = String.Empty
#End Region

#Region "Property"

#Region "PicList"
    Property PicList() As DataTable
        Get
            Return myPicList
        End Get
        Set(ByVal value As DataTable)
            myPicList = value
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

#End Region

#Region "Function"

    Public Function GetPicList() As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        '// Call Function
        If clsBG_M_USER.select012() = True Then
            Me.PicList = clsBG_M_USER.dtResult

            Return True
        Else
            Me.PicList = Nothing

            Return False
        End If
    End Function

    Public Function UnlockPic() As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        '// Set parameters
        clsBG_M_USER.UserPIC = Me.PicNo

        '// Call Function
        If clsBG_M_USER.Delete002() = True Then

            Return True
        Else
            Return False

        End If
    End Function

#End Region

End Class
