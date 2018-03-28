Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Text

Public Class frmBG0681

#Region "Variable"
    Private myClsBG0681BL As New clsBG0681BL
    Private isInsert As Boolean = False
    Dim dtMaster As DataTable
#End Region

#Region "Overrides Function"
    Public Sub New(ByRef frmParent As Form, ByVal strFormName As String, ByVal blnMaximize As Boolean)
        '// This call is required by the Windows Form Designer.
        InitializeComponent()

        '// Add any initialization after the InitializeComponent() call.
        Me.MdiParent = frmParent
        If blnMaximize Then
            Me.WindowState = FormWindowState.Maximized
        Else
            Me.WindowState = FormWindowState.Normal
        End If
        Me.Text = strFormName
        Me.SearchDataGrid()
    End Sub
#End Region

#Region "Function"
    Public Sub setGridHeaderText()
        Me.grvMaster.Columns("ASSET_CATEGORY").HeaderText = "No"
        Me.grvMaster.Columns("ASSET_CATEGORY_TXT").HeaderText = "Asset category name"
        Me.grvMaster.Columns("CREATE_USER_ID").HeaderText = "Create user id"
        Me.grvMaster.Columns("CREATE_DATE").HeaderText = "Create date"
        Me.grvMaster.Columns("UPDATE_USER_ID").HeaderText = "Update user id"
        Me.grvMaster.Columns("UPDATE_DATE").HeaderText = "Update date"
    End Sub

    Public Sub setGridHeaderProperty()
        Me.grvMaster.Columns("ASSET_CATEGORY").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        Me.grvMaster.Columns("ASSET_CATEGORY").Width = 40
        Me.grvMaster.Columns("ASSET_CATEGORY").MinimumWidth = 40
        Me.grvMaster.Columns("ASSET_CATEGORY").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.grvMaster.Columns("ASSET_CATEGORY_TXT").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        Me.grvMaster.Columns("ASSET_CATEGORY_TXT").Width = 150
        Me.grvMaster.Columns("ASSET_CATEGORY_TXT").MinimumWidth = 150

        Me.grvMaster.Columns("CREATE_USER_ID").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        Me.grvMaster.Columns("CREATE_USER_ID").MinimumWidth = 100

        Me.grvMaster.Columns("CREATE_DATE").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        Me.grvMaster.Columns("CREATE_DATE").MinimumWidth = 100

        Me.grvMaster.Columns("UPDATE_USER_ID").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        Me.grvMaster.Columns("UPDATE_USER_ID").MinimumWidth = 100

        Me.grvMaster.Columns("UPDATE_DATE").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        Me.grvMaster.Columns("UPDATE_DATE").MinimumWidth = 100
    End Sub

    Public Sub SearchDataGrid()

        myClsBG0681BL.AssetCategoryFilter = Me.txtCategoryNoFilter.Text.Trim
        myClsBG0681BL.AssetCategoryTxtFilter = Me.txtCategoryNameFilter.Text.Trim

        If myClsBG0681BL.getAssetCategory() = True Then
            dtMaster = myClsBG0681BL.DtResult
            Me.grvMaster.DataSource = dtMaster

            setGridHeaderText()
            setGridHeaderProperty()
        Else
            MessageBox.Show("Error: Can not get Asset category information!", Me.Text, MessageBoxButtons.OK, _
                            MessageBoxIcon.Error)
        End If
    End Sub

    Public Sub setText(ByVal intRow As Integer)
        isInsert = False
        Me.txtCategoryNo.Enabled = False
        Me.txtCategoryNo.Text = CStr(Me.grvMaster.Rows(intRow).Cells("ASSET_CATEGORY").Value)
        Me.txtCategoryName.Text = CStr(Me.grvMaster.Rows(intRow).Cells("ASSET_CATEGORY_TXT").Value)
    End Sub

    Public Function checkData(ByVal strCategoryNo As String) As Boolean
        myClsBG0681BL.AssetCategory = strCategoryNo

        If myClsBG0681BL.checkData() = True Then
            If myClsBG0681BL.DtResult.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    Private Sub clearFilter()
        Me.txtCategoryNoFilter.Text = ""
        Me.txtCategoryNameFilter.Text = ""
    End Sub

#End Region

#Region "Control Event"

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Me.txtCategoryNo.Text = String.Empty
        Me.txtCategoryName.Text = String.Empty
        Me.txtCategoryNo.Enabled = True
        isInsert = True
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        '// Remember current state
        Dim intFirstRow As Integer = grvMaster.FirstDisplayedCell.RowIndex
        Dim intFirstCol As Integer = grvMaster.FirstDisplayedCell.ColumnIndex
        Dim intSelRow As Integer = grvMaster.SelectedCells(0).RowIndex
        Dim intSelCol As Integer = grvMaster.SelectedCells(0).ColumnIndex

        If isInsert = True Then
            If MessageBox.Show("Are you sure to create new Asset category?", Me.Text, MessageBoxButtons.YesNo, _
                          MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            myClsBG0681BL.AssetCategory = Me.txtCategoryNo.Text
            myClsBG0681BL.AssetCategoryTxt = Me.txtCategoryName.Text
            myClsBG0681BL.CreateUserId = p_strUserId

            If myClsBG0681BL.insertOneData() = True Then
                MessageBox.Show("Asset category was created", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                '// Write Transaction Log
                WriteTransactionLog(CStr(enumOperationCd.EditAssetCategoryMaster), "", "", "", "", "", "")

                SearchDataGrid()
            Else
                If myClsBG0681BL.DtResult.Rows.Count >= 1 Then
                    MessageBox.Show("Asset category" & Me.txtCategoryNo.Text & " is exist", Me.Text, MessageBoxButtons.OK, _
                                    MessageBoxIcon.Information)
                Else
                    MessageBox.Show("There are error between save Asset category", Me.Text, MessageBoxButtons.OK, _
                                    MessageBoxIcon.Error)
                End If
            End If
        Else
            If MessageBox.Show("Are you sure to update Asset category?", Me.Text, MessageBoxButtons.YesNo, _
                          MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            Dim dt As DataTable = CType(Me.grvMaster.DataSource, DataTable)
            dt.PrimaryKey = New DataColumn() {dt.Columns(0)}
            Dim dr As DataRow
            Dim findTheseVals(0) As Object

            findTheseVals(0) = Me.txtCategoryNo.Text
            dr = dt.Rows.Find(findTheseVals)

            myClsBG0681BL.AssetCategory = Me.txtCategoryNo.Text
            myClsBG0681BL.AssetCategoryTxt = Me.txtCategoryName.Text
            myClsBG0681BL.UpdateUserId = p_strUserId

            If myClsBG0681BL.UpdateOneData() = True Then
                MessageBox.Show("Asset category was updated", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                '// Write Transaction Log
                WriteTransactionLog(CStr(enumOperationCd.EditAssetCategoryMaster), "", "", "", "", "", "")

                SearchDataGrid()
            Else
                MessageBox.Show("There are error between save Asset category", Me.Text, MessageBoxButtons.OK, _
                                MessageBoxIcon.Error)
            End If
        End If

        '// Select edited row
        If intFirstRow < grvMaster.Rows.Count Then
            If grvMaster.Item(intFirstCol, intFirstRow) IsNot Nothing Then
                grvMaster.FirstDisplayedCell = grvMaster.Item(intFirstCol, intFirstRow)
            End If
        End If
        If intSelRow < grvMaster.Rows.Count Then
            If grvMaster.Item(intSelCol, intSelRow) IsNot Nothing Then
                grvMaster.Item(intSelCol, intSelRow).Selected = True
            End If
        End If

        setText(intSelRow)
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        If MessageBox.Show("Are you sure to delete Asset category?", Me.Text, MessageBoxButtons.YesNo, _
                      MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        myClsBG0681BL.AssetCategory = Me.txtCategoryNo.Text

        If myClsBG0681BL.DeleteData() = True Then
            MessageBox.Show("Asset category was deleted", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            SearchDataGrid()
        Else
            MessageBox.Show("There are error between delete Asset category", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub cmdImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdImport.Click
        Dim connectionString As String
        Dim FileToConvert As String
        Dim ds As New DataSet()
        Dim dt As New DataTable()

        Try
            Dim opDialog As New OpenFileDialog
            opDialog.Filter = "Microsoft Excel Workbook (*.xls)|*.xls"

            Dim dlrConfirm As DialogResult = opDialog.ShowDialog()
            If dlrConfirm.Equals(DialogResult.Cancel) Then
                Exit Sub
            End If

            FileToConvert = opDialog.FileName

            connectionString = "provider=Microsoft.ACE.OLEDB.12.0; " & _
                                "Data Source=" & FileToConvert & "; Extended Properties=Excel 12.0;"

            Dim connection As OleDbConnection = New OleDbConnection(connectionString)
            connection.Open()
            Dim adapter As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM [Sheet1$]", connection)

            adapter.Fill(ds)
            dt = ds.Tables(0)

            connection.Close()

        Catch ex As Exception
            MessageBox.Show("There are error between import file", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Dim conn As SqlConnection = Nothing
        Dim trans As SqlTransaction
        Dim success As Boolean = False

        conn = New SqlConnection(My.Settings.ConnStr)
        conn.Open()
        trans = conn.BeginTransaction()

        Try
            For Each row As DataRow In dt.Rows

                If checkData(row(0).ToString) = True Then
                    '// update
                    myClsBG0681BL.AssetCategory = row(0).ToString
                    myClsBG0681BL.AssetCategoryTxt = row(1).ToString
                    myClsBG0681BL.UpdateUserId = p_strUserId

                    If myClsBG0681BL.UpdateExcelData(conn, trans) = True Then
                        success = True
                    Else
                        success = False
                        Exit For
                    End If

                Else
                    '// insert
                    myClsBG0681BL.AssetCategory = row(0).ToString
                    myClsBG0681BL.AssetCategoryTxt = row(1).ToString
                    myClsBG0681BL.CreateUserId = p_strUserId

                    If myClsBG0681BL.insertExcelData(conn, trans) = True Then
                        success = True
                    Else
                        success = False
                        Exit For
                    End If
                End If
            Next

            If success Then
                MessageBox.Show("Import file completed", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                trans.Commit()

                '// Write Transaction Log
                WriteTransactionLog(CStr(enumOperationCd.EditAssetCategoryMaster), "", "", "", "", "", "")

                SearchDataGrid()
            Else
                MessageBox.Show("There are error between import file", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                trans.Rollback()
            End If

        Catch ex As Exception
            MessageBox.Show("There are error between import file", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            trans.Rollback()
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub cmdExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExport.Click
        If Me.grvMaster.Columns.Count = 0 Or Me.grvMaster.Rows.Count = 0 Then
            Exit Sub
        End If

        '// Show dialog box
        Dim sdlgSave As SaveFileDialog = New SaveFileDialog
        sdlgSave.FileName = "AssetCategoryMaster_" & Format(Date.Now, "yyyyMMdd")
        sdlgSave.Filter = "Microsoft Excel Workbook (*.xls)|*.xls"

        Dim dlrConfirm As DialogResult = sdlgSave.ShowDialog()
        If dlrConfirm.Equals(DialogResult.Cancel) Then
            Exit Sub
        End If

        Dim rowscount As Integer = dtMaster.Rows.Count
        Dim excel As New Microsoft.Office.Interop.Excel.Application
        Dim wBook As Microsoft.Office.Interop.Excel.Workbook
        Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet

        wBook = excel.Workbooks.Add()
        wSheet = CType(wBook.ActiveSheet(), Microsoft.Office.Interop.Excel.Worksheet)

        excel.Range("A1", "A" & (rowscount).ToString).NumberFormat = "@"

        Dim dc As System.Data.DataColumn
        Dim dr As System.Data.DataRow
        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = 0

        excel.Cells(1, 1) = "Asset category no"
        excel.Cells(1, 2) = "Asset category name"
        excel.Cells(1, 3) = "Create user id"
        excel.Cells(1, 4) = "Create date"
        excel.Cells(1, 5) = "Update user id"
        excel.Cells(1, 6) = "Update date"

        For Each dr In dtMaster.Rows
            rowIndex = rowIndex + 1
            colIndex = 0
            For Each dc In dtMaster.Columns
                colIndex = colIndex + 1
                excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
            Next
        Next

        wSheet.Columns.AutoFit()
        Dim strFileName As String = sdlgSave.FileName
        Dim blnFileOpen As Boolean = False
        Try
            Dim fileTemp As System.IO.FileStream = System.IO.File.OpenWrite(strFileName)
            fileTemp.Close()
        Catch ex As Exception
            blnFileOpen = False
        End Try

        If System.IO.File.Exists(strFileName) Then
            System.IO.File.Delete(strFileName)
        End If

        wBook.SaveAs(strFileName)
        excel.Workbooks.Open(strFileName)
        excel.Visible = True

        '//Release memory
        BGCommon.ExcelReleasememory(excel, wBook, wSheet)

        ''MessageBox.Show("Export file completed", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub grvMaster_RowEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grvMaster.RowEnter
        If Not Me.grvMaster Is Nothing Then
            setText(e.RowIndex)
        End If
    End Sub

    Private Sub cmdFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFilter.Click
        SearchDataGrid()
    End Sub

    Private Sub cmdClearFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearFilter.Click
        clearFilter()
    End Sub
#End Region

End Class