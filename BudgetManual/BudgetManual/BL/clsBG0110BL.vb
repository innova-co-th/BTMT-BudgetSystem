Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class clsBG0110BL

#Region "Variable"

#End Region

#Region "Property"

#End Region

#Region "Function"
    Public Function GetHomeURL() As String
        Dim clsBG_M_SETTINGS As New BG_M_SETTINGS

        '// Call Function
        If clsBG_M_SETTINGS.Select001() = True Then

            Return clsBG_M_SETTINGS.HomeURL
        Else
            Return ""

        End If
    End Function
#End Region

End Class
