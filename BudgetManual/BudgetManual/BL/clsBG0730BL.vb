Imports System.Data.SqlClient
Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Public Class clsBG0730BL

#Region "Function"

    Public Function Backup() As Boolean

        Dim blnResult As Boolean
        'Dim da As DbAccess.IDbAccess = Common.clsCommon.AppDbAccess
        'Dim strSQL As String = ""

        'strSQL = "BACKUP DATABASE " & _
        '         Common.DbSetting.GetConfig.DbName & _
        '         " TO DISK='" & strBackupTempPath & "';"

        'Try
        '    If Common.clsCommon.AppDbAccess.IsConnected = False Then
        '        Common.clsCommon.AppDbAccess.Connect()
        '    End If

        '    Common.clsCommon.WriteSqlFile("Backup", strSQL)
        '    blnResult = Common.clsCommon.AppDbAccess.ExecuteSql(strSQL)
        '    If Not Common.clsCommon.AppDbAccess.GetLastError() Is Nothing Then
        '        Common.clsCommon.WriteErrorLog("Backup", Common.clsCommon.AppDbAccess.GetLastError())

        '    Else
        '    End If
        'Catch ex As Exception
        '    Common.clsCommon.WriteErrorLog("Backup", ex)

        'Finally
        '    If Common.clsCommon.AppDbAccess.IsConnected = True Then
        '        Common.clsCommon.AppDbAccess.Close()
        '    End If
        'End Try

        Return blnResult

    End Function

#End Region

End Class
