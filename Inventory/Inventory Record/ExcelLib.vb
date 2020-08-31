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
    ''' <param name="DG">Data Grid</param>
    ''' <param name="tableName">Table name</param>
    Public Shared Sub Import(pathOpenFile As String, frmParent As Form, DG As DataGrid, tableName As String)
        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
        Dim xlApp As Excel.Application = New Excel.Application()
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim frmOverlay As New Form()

        Try
            'Create loading of overlay
            Using frm As New Loading()
                frmOverlay.StartPosition = FormStartPosition.Manual
                frmOverlay.FormBorderStyle = FormBorderStyle.None
                frmOverlay.Opacity = 0.5D
                frmOverlay.BackColor = Color.Black
                frmOverlay.WindowState = FormWindowState.Maximized
                frmOverlay.TopMost = True
                frmOverlay.Location = frmParent.Location
                frmOverlay.ShowInTaskbar = False
                'frmOverlay.Show()
                frm.Owner = frmOverlay
                ExcelLib.CenterForm(frm, frmParent)
                frm.Show()

                'Read data in Excel
                xlWorkBook = xlApp.Workbooks.Open(pathOpenFile)
                xlWorkSheet = CType(xlWorkBook.Sheets(1), Excel.Worksheet) 'Get first sheet

                Dim numRows As Integer = xlWorkSheet.UsedRange.Rows.Count
                Dim numCols As Integer = xlWorkSheet.UsedRange.Columns.Count
                Dim xlRange As Excel.Range
                Dim arr As Object(,) = New Object(numRows, numCols) {}
                Dim dtRec As New DataTable(tableName)

                'Calculate the final column letter
                Dim finalColLetter As String = String.Empty
                Dim colCharset As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                Dim colCharsetLen As Integer = colCharset.Length

                If numCols > colCharsetLen Then
                    finalColLetter = colCharset.Substring((numCols - 1) \ colCharsetLen - 1, 1)
                End If

                finalColLetter += colCharset.Substring((numCols - 1) Mod colCharsetLen, 1)
                xlRange = xlWorkSheet.Range("A1:" & finalColLetter & CStr(numCols))
                arr(1, 1) = xlRange

                '    'display the cells value B2
                '    MsgBox(xlWorkSheet.Cells(2, 2).value)
                '    'edit the cell with new value
                '    xlWorkSheet.Cells(2, 2) = "http://vb.net-informations.com"

                ReleaseObject(xlWorkSheet)
                xlWorkBook.Close(False)
                ReleaseObject(xlWorkBook)

                frmOverlay.Dispose()
            End Using 'Using frm
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
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

    ''' <summary>
    ''' Export data grid to export file
    ''' </summary>
    ''' <param name="frmParent">Parent Form</param>
    ''' <param name="DG">Data Grid</param>
    ''' <param name="tableName">table name</param>
    Public Shared Sub Export(frmParent As Form, DG As DataGrid, tableName As String)
        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
        Dim xlApp As Excel.Application = New Excel.Application()
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Add()
        Dim xlWorkSheet As Excel.Worksheet = CType(xlWorkBook.Sheets("Sheet1"), Excel.Worksheet)
        Dim dtGridView As DataView = CType(DG.DataSource, DataView)
        Dim exportDialog As SaveFileDialog = New SaveFileDialog With {
            .Filter = System.Configuration.ConfigurationManager.AppSettings("DIALOG_FILE_EXT").ToString()
        }
        Dim pathSaveFile As String = String.Empty
        Dim frmOverlay As New Form()

        Try
            'Check number of record in DataGrid
            If dtGridView.Count <= 0 Then
                'Error
                MessageBox.Show("Export error" & vbCrLf & "No data for export!!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                'Export excel
                If exportDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    Dim xlRange As Excel.Range
                    Dim misValue As Object = Type.Missing
                    Dim dtRec As DataTable = dtGridView.ToTable(tableName)

                    'Create loading of overlay
                    Using frm As New Loading()
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
                        xlWorkSheet.Name = tableName

                        Dim dtHead As DataGridTableStyle = DG.TableStyles(0) 'Header as DataGridRM
                        Dim dtTemp As DataTable = New DataTable("TempData") 'Temporary datable

                        'Create temporary datatable
                        For j As Integer = 0 To dtHead.GridColumnStyles.OfType(Of DataGridColoredLine2).Count - 1
                            If dtRec.Rows(0)(dtHead.GridColumnStyles.Item(j).MappingName).GetType().Equals(GetType(Decimal)) Then
                                dtTemp.Columns.Add(dtHead.GridColumnStyles.Item(j).MappingName, GetType(Decimal))
                            Else
                                dtTemp.Columns.Add(dtHead.GridColumnStyles.Item(j).MappingName, GetType(String))
                            End If
                        Next

                        'Set header
                        For j As Integer = 1 To dtHead.GridColumnStyles.OfType(Of DataGridColoredLine2).Count
                            xlWorkSheet.Cells(1, j) = dtHead.GridColumnStyles.Item(j - 1).HeaderText
                        Next

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
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
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

    Private Shared Function ConvertArrayToDatatable(arr As Object(,)) As DataTable
        Try
            Dim dtRec As New DataTable()
            For i As Integer = 0 To arr.GetLength(0)
                For j As Integer = 0 To arr.GetLength(1)
                    MsgBox(arr(i, j).ToString())
                Next j
            Next i

            Return dtRec
        Catch ex As Exception
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
