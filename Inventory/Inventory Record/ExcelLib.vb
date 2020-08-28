Imports System.Runtime.InteropServices

Public Class ExcelLib
#Region "Windows DLL"
    <DllImport("user32.dll")>
    Public Shared Function GetWindowThreadProcessId(ByVal hWnd As Integer, ByRef lpdwProcessId As Integer) As Integer
    End Function
#End Region

#Region "Excel"
    Public Shared Function ConvertDatatableToObject(oDt As DataTable) As Object
        Try
            Dim arr As Object(,) = New Object(oDt.Rows.Count - 1, oDt.Columns.Count - 1) {}
            For r As Integer = 0 To oDt.Rows.Count - 1
                Dim dr As DataRow = oDt.Rows(r)
                For c As Integer = 0 To oDt.Columns.Count - 1
                    If oDt.Columns(c).DataType.ToString() = "System.String" Then
                        'object val = "'" + dr[c];
                        'arr[r, c] = val;
                        If String.IsNullOrEmpty(dr(c).ToString()) Then
                            arr(r, c) = dr(c)
                        Else
                            Dim val As Object = "'" + dr(c)
                            arr(r, c) = val
                        End If
                    Else
                        arr(r, c) = dr(c)
                    End If
                Next
            Next

            Return arr

        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Public Shared Sub ReleaseObject(ByVal obj As Object)
        Try
            While (System.Runtime.InteropServices.Marshal.ReleaseComObject(obj) > 0)
            End While
        Catch
        Finally
            obj = Nothing
        End Try
    End Sub
#End Region
End Class
