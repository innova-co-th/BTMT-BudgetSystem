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
                Throw New ApplicationException("Number columns of import is incorrect!!!" & vbCrLf & "It have empty column " & numCols & " column(s). Please remove column header.")
            End If

            'Check column header
            For i As Integer = 1 To numCols
                'Error
                If arr(1, i) <> arrColumn(i - 1) Then
                    Throw New ApplicationException("Column : " & arr(1, i) & " is incorrect!!!")
                End If
            Next

            'Create temporary datatable
            For i As Integer = 1 To numCols
                'Check third row when second row nothing
                If IsNothing(arr(2, i)) Then
                    'If column is nothing value
                    dtTemp.Columns.Add(arr(1, i), GetType(String))
                Else
                    'Check second row
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
            dtTemp = ConvertArrayToDatatable(arr, dtTemp, tableName)

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

    ''' <summary>
    ''' Export data grid to export file
    ''' </summary>
    ''' <param name="frmParent">Parent Form</param>
    ''' <param name="DV">Data View</param>
    ''' <param name="tableName">table name</param>
    ''' <param name="arrColumn">Array column map with datatable</param>
    ''' <param name="arrColumnHeader">Array column header</param>
    Public Shared Sub Export(frmParent As Form, DV As DataView, tableName As String, arrColumn As String(), arrColumnHeader As String())
        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
        Dim xlApp As Excel.Application = New Excel.Application()
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Add()
        Dim xlWorkSheet As Excel.Worksheet = CType(xlWorkBook.Sheets("Sheet1"), Excel.Worksheet)
        Dim exportDialog As SaveFileDialog = New SaveFileDialog With {
            .Filter = System.Configuration.ConfigurationManager.AppSettings("DIALOG_FILE_EXT").ToString()
        }
        Dim pathSaveFile As String = String.Empty
        Dim frmOverlay As New Form()

        Try
            'Check number of record in DataView
            If DV.Count <= 0 Then
                'Error
                Throw New ApplicationException("No data for export!!")
            Else
                'Export excel
                If exportDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    Dim xlRange As Excel.Range
                    Dim misValue As Object = Type.Missing
                    Dim dtRec As DataTable = DV.ToTable(tableName) 'Get datatable by table name

                    'Get dataview into excel file
                    Using frm As New Exporting()
                        'Create loading of overlay
                        frmOverlay.StartPosition = FormStartPosition.Manual
                        frmOverlay.FormBorderStyle = FormBorderStyle.None
                        frmOverlay.Opacity = 0.5D
                        frmOverlay.BackColor = Color.Black
                        frmOverlay.WindowState = FormWindowState.Maximized
                        frmOverlay.TopMost = True
                        frmOverlay.Location = frmParent.Location
                        frmOverlay.ShowInTaskbar = False
                        frmOverlay.Show()
                        frm.Owner = frmOverlay
                        ExcelLib.CenterForm(frm, frmParent)
                        frm.Show()

                        pathSaveFile = exportDialog.FileName
                        xlWorkSheet.Name = tableName 'Set sheet name

                        Dim dtTemp As DataTable = New DataTable("TempData") 'Temporary datable

                        'Create temporary datatable


                        For Each col As String In arrColumn
                            'Check type of first row
                            If dtRec.Rows(0)(col).GetType().Equals(GetType(Decimal)) Then
                                dtTemp.Columns.Add(col, GetType(Decimal))
                            ElseIf dtRec.Columns.Item(col).GetType().Equals(GetType(Int32)) Then
                                dtTemp.Columns.Add(col, GetType(Int32))
                            Else
                                dtTemp.Columns.Add(col, GetType(String))
                            End If
                        Next

                        'Set header
                        For j As Integer = 1 To arrColumnHeader.Length
                            xlWorkSheet.Cells(1, j) = arrColumnHeader(j - 1) 'Excel start position at 1, Array start position at 0
                        Next

                        'Set font bold of header
                        xlRange = xlWorkSheet.Range(CType(xlWorkSheet.Cells(1, 1), Excel.Range), CType(xlWorkSheet.Cells(1, arrColumn.Length), Excel.Range))
                        xlRange.Font.Bold = True

                        'Set data
                        For i As Integer = 0 To dtRec.Rows.Count - 1
                            Dim drData As DataRow = dtTemp.NewRow()
                            For j As Integer = 0 To dtTemp.Columns.Count - 1
                                drData(j) = dtRec.Rows(i)(dtTemp.Columns(j).ColumnName)
                            Next
                            dtTemp.Rows.Add(drData)
                        Next i

                        'Set range for data
                        Dim c1 As Excel.Range = CType(xlWorkSheet.Cells(2, 1), Excel.Range)
                        Dim c2 As Excel.Range = CType(xlWorkSheet.Cells(2 + dtTemp.Rows.Count - 1, dtTemp.Columns.Count), Excel.Range)
                        xlRange = xlWorkSheet.Range(c1, c2)

                        'Convert DataTable to Array Object
                        xlRange.Value2 = ConvertDatatableToObject(dtTemp)

                        'Set autofit column
                        c1 = CType(xlWorkSheet.Cells(1, 1), Excel.Range)
                        c2 = CType(xlWorkSheet.Cells(1 + dtTemp.Rows.Count, dtTemp.Columns.Count), Excel.Range)
                        xlRange = xlWorkSheet.Range(c1, c2)
                        xlRange.Columns.AutoFit()

                        'Set off for display alerts
                        xlApp.DisplayAlerts = False
                        'Save excel
                        xlWorkBook.SaveAs(pathSaveFile, misValue, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, misValue, misValue, misValue, misValue, misValue)

                        MessageBox.Show("Export complete", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        frmOverlay.Dispose()
                    End Using 'Using frm
                End If 'If exportDialog.ShowDialog() = Windows.Forms.DialogResult.OK
            End If 'If GrdDV.Count <= 0

        Catch ex As Exception
            MessageBox.Show("Export error" & vbCrLf & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ReleaseObject(xlWorkSheet)
            xlWorkBook.Close(False)
            ReleaseObject(xlWorkBook)
            xlApp.Quit()
            frmOverlay.Dispose()

            Dim pid As Integer = 0
            Dim a As Integer = GetWindowThreadProcessId(xlApp.Hwnd, pid)
            Dim p As Process = Process.GetProcessById(pid)
            p.Kill()

            ReleaseObject(xlApp)
            GC.Collect()
        End Try
    End Sub

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

    ''' <summary>
    ''' Convert data from array to datatable
    ''' </summary>
    ''' <param name="arr">Data Source</param>
    ''' <param name="dt">Data Destination</param>
    ''' <param name="tableName">Type Code</param>
    ''' <returns></returns>
    Private Shared Function ConvertArrayToDatatable(arr As Object(,), dt As DataTable, tableName As String) As DataTable
        Try
            For i As Integer = 2 To arr.GetLength(0)
                Dim dr As DataRow = dt.NewRow()

                For j As Integer = 1 To arr.GetLength(1)
                    'Check nothing
                    If IsNothing(arr(i, j)) Then
                        If tableName.Equals("TBL_PreSemi") Then
                            'PreSemi type
                            If j = 4 Or j = 5 Or j = 6 Then
                                'Column Width, Length, N
                                dr(j - 1) = 0.0
                            End If
                        ElseIf tableName.Equals("TBL_Semi") Then
                            'Semi type
                        ElseIf tableName.Equals("TBL_GT") Then
                            'Green Tire
                            If j = 9 Then
                                'Column Length
                                dr(j - 1) = 0.0
                            End If
                        Else
                            'Other type (R/M, Pigment, Compound)
                            dr(j - 1) = ""
                        End If 'If tableName.Equals("TBL_PreSemi")
                    Else
                        If arr(i, j).GetType().Equals(GetType(Double)) Then
                            'Datatype Double
                            dr(j - 1) = CDec(arr(i, j))
                        Else
                            'Other
                            dr(j - 1) = arr(i, j)
                        End If
                    End If 'If IsNothing(arr(i, j))
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
