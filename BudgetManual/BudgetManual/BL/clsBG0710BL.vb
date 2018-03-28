Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class clsBG0710BL

#Region "Variable"
    Private myPassword As String = String.Empty
#End Region

#Region "Property"

#Region "Password"
    Property Password() As String
        Get
            Return myPassword
        End Get
        Set(ByVal value As String)
            myPassword = value
        End Set
    End Property
#End Region

#End Region

#Region "Function"

    Public Function ChangePassword() As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        '// Set Parameters
        clsBG_M_USER.UserId = p_strUserId
        clsBG_M_USER.Password = Me.Password

        '// Call Function
        If clsBG_M_USER.Update001() = True Then

            Return True
        Else
            Return False

        End If
    End Function

#End Region

End Class
