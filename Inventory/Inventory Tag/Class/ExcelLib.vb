Imports System.Runtime.InteropServices
Imports Excel = Microsoft.Office.Interop.Excel

Public Class ExcelLib
#Region "Windows DLL"
    <DllImport("user32.dll")>
    Private Shared Function GetWindowThreadProcessId(ByVal hWnd As Integer, ByRef lpdwProcessId As Integer) As Integer
    End Function
#End Region

#Region "Center Screen"
    Public Shared Sub CenterForm(ByVal frm As Form, ByVal parent As Form)
        Dim x As Integer = parent.Left + (parent.Width - frm.Width) \ 2
        Dim y As Integer = parent.Top + (parent.Height - frm.Height) \ 2
        frm.Location = New Point(x, y)
    End Sub
#End Region

#Region "Excel"
    ''' <summary>
    ''' Import excel file to database
    ''' </summary>
    ''' <param name="pathOpenFile">Import file</param>
    ''' <param name="frmParent">Parent form</param>
    ''' <param name="DV">Data View</param>
    ''' <param name="tableName">Table name</param>
    ''' <param name="arrColumn">Array column</param>
    Public Shared Function Import(pathOpenFile As String, frmParent As Form, DV As DataView, tableName As String, arrColumn As String()) As DataTable
        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
        Dim xlApp As Excel.Application = New Excel.Application()
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim dtTemp As DataTable = New DataTable("TempData") 'Temporary datable

        Try
            'Read data in Excel
            xlWorkBook = xlApp.Workbooks.Open(pathOpenFile)
            xlWorkSheet = CType(xlWorkBook.Sheets(1), Excel.Worksheet) 'Get first sheet

            Dim numRows As Integer = xlWorkSheet.UsedRange.Rows.Count
            Dim numCols As Integer = xlWorkSheet.UsedRange.Columns.Count
            Dim xlRange As Excel.Range = xlWorkSheet.UsedRange 'Set range
            Dim arr As Object(,) = xlRange.Value2 'Load excel into array
            Dim dtRec As New DataTable(tableName)

            'Check number of record in Excel
            If numRows <= 1 Then
                'It have only header
                Throw New ApplicationException("No data for import!!")
            End If

            'Check format of import file
            If numCols <> arrColumn.Length Then
                'Error
                Throw New ApplicationException("Number columns of import is incorrect!!!")
            End If

            'Check column header of first row of array
            For i As Integer = 1 To numCols
                'Error
                If arr(1, i) <> arrColumn(i - 1) Then
                    Throw New ApplicationException("Column : " & arr(1, i) & " is incorrect!!!")
                End If
            Next

            'Create temporary datatable
            For i As Integer = 1 To numCols
                'Check second row of array
                If arr(2, i) Is Nothing Then
                    'If column is nothing value
                    dtTemp.Columns.Add(arr(1, i), GetType(String))
                Else
                    If arr(2, i).GetType().Equals(GetType(Double)) Then
                        dtTemp.Columns.Add(arr(1, i), GetType(Decimal))
                    ElseIf arr(2, i).GetType().Equals(GetType(Int32)) Then
                        dtTemp.Columns.Add(arr(1, i), GetType(Int32))
                    Else
                        dtTemp.Columns.Add(arr(1, i), GetType(String))
                    End If
                End If
            Next i

            'Convert Array to Datatable
            dtTemp = ConvertArrayToDatatable(arr, dtTemp)

            ReleaseObject(xlWorkSheet)
            xlWorkBook.Close(False)
            ReleaseObject(xlWorkBook)
        Catch ex As Exception
            MessageBox.Show("Import error" & vbCrLf & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            dtTemp = Nothing
        Finally
            xlApp.Quit()

            Dim pid As Integer = 0
            Dim a As Integer = GetWindowThreadProcessId(xlApp.Hwnd, pid)
            Dim p As Process = Process.GetProcessById(pid)
            p.Kill()

            ReleaseObject(xlApp)
            GC.Collect()
        End Try

        Return dtTemp
    End Function

    Private Shared Function ConvertDatatableToObject(oDt As DataTable) As Object
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

    Private Shared Function ConvertArrayToDatatable(arr As Object(,), dt As DataTable) As DataTable
        Try
            For i As Integer = 2 To arr.GetLength(0)
                Dim dr As DataRow = dt.NewRow()
                For j As Integer = 1 To arr.GetLength(1)
                    'Check third row when second row nothing
                    If IsNothing(arr(i, j)) Then
                        dr(j - 1) = ""
                    Else
                        If arr(i, j).GetType().Equals(GetType(Double)) Then
                            'Datatype Double
                            dr(j - 1) = CDec(arr(i, j))
                        Else
                            'Other
                            dr(j - 1) = arr(i, j)
                        End If
                    End If

                    'If arr(i, j).GetType().Equals(GetType(Double)) Then
                    '    'Datatype Double
                    '    dr(j - 1) = CDec(arr(i, j))
                    'Else
                    '    'Other
                    '    dr(j - 1) = arr(i, j)
                    'End If
                Next j
                dt.Rows.Add(dr)
            Next i

            Return dt
        Catch ex As Exception
            MessageBox.Show("Import error" & vbCrLf & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function

    Private Shared Sub ReleaseObject(ByVal obj As Object)
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
