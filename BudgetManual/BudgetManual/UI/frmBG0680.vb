Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Text

Public Class frmBG0680

#Region "Variable"
    Private myClsBG0680BL As New clsBG0680BL
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
        loadComboFilter()
        Me.SearchDataGrid()
        Me.setCombo()
        Me.setAsset()
    End Sub
#End Region

#Region "Function"
    Public Sub setGridHeaderText()
        Me.grvMaster.Columns("ASSET_GROUP_NO").HeaderText = "Asset group no"
        Me.grvMaster.Columns("ASSET_GROUP_NAME").HeaderText = "Asset group name"
        Me.grvMaster.Columns("ASSET_PROJECT").HeaderText = "Asset project"
        Me.grvMaster.Columns("ASSET_PROJECT_TXT").HeaderText = "Asset project"
        Me.grvMaster.Columns("ASSET_CATEGORY").HeaderText = "Asset category"
        Me.grvMaster.Columns("ASSET_CATEGORY_TXT").HeaderText = "Asset category"
        Me.grvMaster.Columns("CREATE_USER_ID").HeaderText = "Create user id"
        Me.grvMaster.Columns("CREATE_DATE").HeaderText = "Create date"
        Me.grvMaster.Columns("UPDATE_USER_ID").HeaderText = "Update user id"
        Me.grvMaster.Columns("UPDATE_DATE").HeaderText = "Update date"
    End Sub

    Public Sub setGridHeaderProperty()
        Me.grvMaster.Columns("ASSET_GROUP_NO").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        Me.grvMaster.Columns("ASSET_GROUP_NO").MinimumWidth = 100

        Me.grvMaster.Columns("ASSET_GROUP_NAME").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        Me.grvMaster.Columns("ASSET_GROUP_NAME").MinimumWidth = 148

        Me.grvMaster.Columns("ASSET_PROJECT").Visible = False

        Me.grvMaster.Columns("ASSET_PROJECT_TXT").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        Me.grvMaster.Columns("ASSET_PROJECT_TXT").MinimumWidth = 100

        Me.grvMaster.Columns("ASSET_CATEGORY").Visible = False

        Me.grvMaster.Columns("ASSET_CATEGORY_TXT").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        Me.grvMaster.Columns("ASSET_CATEGORY_TXT").MinimumWidth = 100

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

        myClsBG0680BL.AssetGroupNoFilter = Me.txtAssetGroupNoFilter.Text.Trim
        myClsBG0680BL.AssetGroupNameFilter = Me.txtAssetGroupNameFilter.Text.Trim
        myClsBG0680BL.AssetProjFilter = Me.cboAssetProjectFilter.SelectedValue.ToString
        myClsBG0680BL.AssetCateFilter = Me.cboCategoryFilter.SelectedValue.ToString

        If myClsBG0680BL.searchDatagrid = True Then
            dtMaster = myClsBG0680BL.DtResult
            Dim intproj As Integer
            Dim intcate As Integer
            '//-- Edit 2011-05-11 by S.Watcharapong
            ''Dim strtmp As StringBuilder
            ''Dim strCat As String = String.Empty
            ''Dim intr As Integer = 0
            Dim dtCategory As DataTable = myClsBG0680BL.getAssetCategory()
            Dim dr As DataRow() = Nothing
            '//-- End Edit 2011-05-11
            '//-- Add 2011-05-11 by S.Watcharapong
            Dim dtProject As DataTable = myClsBG0680BL.getAssetProject()
            '//-- End Add 2011-05-11

            For i As Integer = 0 To dtMaster.Rows.Count - 1
                intproj = CInt(dtMaster.Rows(i).Item("ASSET_PROJECT"))
                intcate = CInt(dtMaster.Rows(i).Item("ASSET_CATEGORY"))

                '//-- Edit 2011-05-12 by S.Watcharapong
                ''dtMaster.Rows(i).Item("ASSET_PROJECT_TXT") = [Enum].GetName(GetType(enumAssetProject), intproj)
                dr = dtProject.Select("ASSET_PROJECT = " & intproj)
                If dr.Length > 0 Then
                    dtMaster.Rows(i).Item("ASSET_PROJECT_TXT") = dr(0)![ASSET_PROJECT_TXT]
                Else
                    dtMaster.Rows(i).Item("ASSET_PROJECT_TXT") = ""
                End If
                '//-- End Edit 2011-05-12

                '//-- Edit 2011-05-11 by S.Watcharapong
                ''strCat = [Enum].GetName(GetType(enumAssetCategory), intcate)
                ''strtmp = New StringBuilder(strCat)
                ''For j As Integer = 1 To strCat.Length - 1
                ''    If Char.IsUpper(strCat, j) Then
                ''        strtmp.Insert(j + intr, " ")
                ''        intr += 1
                ''    End If
                ''Next
                ''dtMaster.Rows(i).Item("ASSET_CATEGORY_TXT") = strtmp.ToString
                ''intr = 0
                dr = dtCategory.Select("ASSET_CATEGORY = " & intcate)
                If dr.Length > 0 Then
                    dtMaster.Rows(i).Item("ASSET_CATEGORY_TXT") = dr(0)![ASSET_CATEGORY_TXT]
                Else
                    dtMaster.Rows(i).Item("ASSET_CATEGORY_TXT") = ""
                End If
                '//-- End Edit 2011-05-11
            Next

            Me.grvMaster.DataSource = dtMaster

            setGridHeaderText()
            setGridHeaderProperty()
        Else
            MessageBox.Show("Error: Can not get Asset group information!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
    End Sub

    Public Function checkData(ByVal strAssetGroupNo As String) As Boolean
        myClsBG0680BL.AssetGroupNo = strAssetGroupNo

        If myClsBG0680BL.checkData() = True Then
            If myClsBG0680BL.DtResult.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    Public Sub setText(ByVal intRow As Integer)
        isInsert = False
        Me.txtAssetGroupNo.Enabled = False
        Me.txtAssetGroupNo.Text = CStr(Me.grvMaster.Rows(intRow).Cells("ASSET_GROUP_NO").Value)
        Me.txtAssetGroupName.Text = CStr(Me.grvMaster.Rows(intRow).Cells("ASSET_GROUP_NAME").Value)
        Me.cboAssetProject.SelectedItem = CStr(Me.grvMaster.Rows(intRow).Cells("ASSET_PROJECT_TXT").Value)
        Me.cboCategory.SelectedItem = CStr(Me.grvMaster.Rows(intRow).Cells("ASSET_CATEGORY_TXT").Value)
    End Sub

    Public Sub setCombo()
        '//-- Edit 2011-05-12 by S.Watcharapong
        ''Dim enumType As Type = GetType(enumAssetProject)
        ''Dim strProj() As String = [Enum].GetNames(enumType)

        ''For i As Integer = 0 To strProj.Length - 1
        ''    cboAssetProject.Items.Add(strProj(i))
        ''Next
        Dim dtProject As DataTable = myClsBG0680BL.getAssetProject()

        For Each dr As DataRow In dtProject.Rows
            cboAssetProject.Items.Add(CStr(dr![ASSET_PROJECT_TXT]))
        Next
        '//-- End Edit 2011-05-12
    End Sub

    Public Sub setAsset()
        '//-- Edit 2011-05-11 by S.Watcharapong
        ''Dim enumType As Type = GetType(enumAssetCategory)
        ''Dim strProj() As String = [Enum].GetNames(enumType)
        ''Dim strtmp As StringBuilder
        ''Dim intr As Integer = 0

        ''For i As Integer = 0 To strProj.Length - 1
        ''    strtmp = New StringBuilder(strProj(i))
        ''    For j As Integer = 1 To strProj(i).Length - 1
        ''        If Char.IsUpper(strProj(i), j) Then
        ''            strtmp.Insert(j + intr, " ")
        ''            intr += 1
        ''        End If
        ''    Next
        ''    cboCategory.Items.Add(strtmp.ToString)
        ''    intr = 0
        ''Next
        Dim dtCategory As DataTable = myClsBG0680BL.getAssetCategory()

        For Each dr As DataRow In dtCategory.Rows
            cboCategory.Items.Add(CStr(dr![ASSET_CATEGORY_TXT]))
        Next
        '//-- End Edit 2011-05-11
    End Sub

    Private Sub clearFilter()
        Me.txtAssetGroupNoFilter.Text = ""
        Me.txtAssetGroupNameFilter.Text = ""
        Me.cboAssetProjectFilter.SelectedIndex = 0
        Me.cboCategoryFilter.SelectedIndex = 0
    End Sub

    Private Sub loadComboFilter()
       
        '// cboAssetProjectFilter
        Dim dtProject As DataTable = myClsBG0680BL.getAssetProject()
        fillComboBox(Me.cboAssetProjectFilter, dtProject, "ASSET_PROJECT", "ASSET_PROJECT_TXT", True)

        '// cboAssetCategoryFilter
        Dim dtCategory As DataTable = myClsBG0680BL.getAssetCategory()
        fillComboBox(Me.cboCategoryFilter, dtCategory, "ASSET_CATEGORY", "ASSET_CATEGORY_TXT", True)
     
    End Sub

    Private Sub fillComboBox(ByVal cbo As ComboBox, _
                            ByVal dt As DataTable, _
                            ByVal keyColName As String, _
                            ByVal textColName As String, _
                            ByVal allowFirstEmptyItem As Boolean)
        Dim str As String = String.Empty
        cbo.Items.Clear()
        If dt.Rows.Count > 0 Then
            If allowFirstEmptyItem Then
                Dim dr As DataRow = dt.NewRow
                dr(keyColName) = "0"
                dr(textColName) = "All"
                dt.Rows.InsertAt(dr, 0)
            End If
            cbo.DataSource = dt
            cbo.ValueMember = keyColName
            cbo.DisplayMember = textColName
        End If
        cbo.SelectedIndex = 0
    End Sub

#End Region

#Region "Control Event"

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Me.txtAssetGroupNo.Text = String.Empty
        Me.txtAssetGroupName.Text = String.Empty
        Me.txtAssetGroupNo.Enabled = True
        Me.cboAssetProject.SelectedIndex = 0
        isInsert = True
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim strProj As String = CStr(Me.cboAssetProject.SelectedIndex + 1)
        Dim strCate As String = CStr(Me.cboCategory.SelectedIndex + 1)

        '// Remember current state
        Dim intFirstRow As Integer = grvMaster.FirstDisplayedCell.RowIndex
        Dim intFirstCol As Integer = grvMaster.FirstDisplayedCell.ColumnIndex
        Dim intSelRow As Integer = grvMaster.SelectedCells(0).RowIndex
        Dim intSelCol As Integer = grvMaster.SelectedCells(0).ColumnIndex

        If isInsert = True Then
            If MessageBox.Show("Are you sure to create new Asset group no?", Me.Text, MessageBoxButtons.YesNo, _
                          MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            Dim strAsset As String = Me.txtAssetGroupNo.Text
            myClsBG0680BL.AssetGroupNo = Me.txtAssetGroupNo.Text
            myClsBG0680BL.AssetGroupName = Me.txtAssetGroupName.Text
            myClsBG0680BL.AssetCate = strCate

            '//-- Edit 2011-05-12 by S.Watcharapong
            ''If strProj = [Enum].GetName(GetType(enumAssetProject), enumAssetProject.TBR) Then
            ''    myClsBG0680BL.AssetProj = CStr(enumAssetProject.TBR)
            ''ElseIf strProj = [Enum].GetName(GetType(enumAssetProject), enumAssetProject.ORR) Then
            ''    myClsBG0680BL.AssetProj = CStr(enumAssetProject.ORR)
            ''Else
            ''    myClsBG0680BL.AssetProj = CStr(enumAssetProject.PCT)
            ''End If
            myClsBG0680BL.AssetProj = strProj
            '//-- End Edit 2011-05-12

            myClsBG0680BL.CreateUserId = p_strUserId

            If myClsBG0680BL.insertOneData() = True Then
                MessageBox.Show("Asset group no was created", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                '// Write Transaction Log
                WriteTransactionLog(CStr(enumOperationCd.EditAssetGroupMaster), "", "", "", "", "", "")

                SearchDataGrid()
            Else
                If myClsBG0680BL.DtResult.Rows.Count >= 1 Then
                    MessageBox.Show("Asset group no " & Me.txtAssetGroupNo.Text & " is exist", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("There are error between save Asset group no", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        Else
            If MessageBox.Show("Are you sure to update Asset group no?", Me.Text, MessageBoxButtons.YesNo, _
                          MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            Dim dt As DataTable = CType(Me.grvMaster.DataSource, DataTable)
            dt.PrimaryKey = New DataColumn() {dt.Columns(0)}
            Dim dr As DataRow
            Dim findTheseVals(0) As Object

            findTheseVals(0) = Me.txtAssetGroupNo.Text
            dr = dt.Rows.Find(findTheseVals)

            myClsBG0680BL.AssetGroupNo = Me.txtAssetGroupNo.Text
            myClsBG0680BL.AssetGroupName = Me.txtAssetGroupName.Text
            myClsBG0680BL.AssetCate = strCate

            '//-- Edit 2011-05-12 by S.Watcharapong
            ''If strProj = [Enum].GetName(GetType(enumAssetProject), enumAssetProject.TBR) Then
            ''    myClsBG0680BL.AssetProj = CStr(enumAssetProject.TBR)
            ''ElseIf strProj = [Enum].GetName(GetType(enumAssetProject), enumAssetProject.ORR) Then
            ''    myClsBG0680BL.AssetProj = CStr(enumAssetProject.ORR)
            ''Else
            ''    myClsBG0680BL.AssetProj = CStr(enumAssetProject.PCT)
            ''End If
            myClsBG0680BL.AssetProj = strProj
            '//-- End Edit 2011-05-12

            myClsBG0680BL.UpdateUserId = p_strUserId

            If myClsBG0680BL.UpdateOneData() = True Then
                MessageBox.Show("Asset group no was updated", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                '// Write Transaction Log
                WriteTransactionLog(CStr(enumOperationCd.EditAssetGroupMaster), "", "", "", "", "", "")

                SearchDataGrid()
            Else
                MessageBox.Show("There are error between save Asset group no", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
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
        If MessageBox.Show("Are you sure to delete Asset group no?", Me.Text, MessageBoxButtons.YesNo, _
                      MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        myClsBG0680BL.AssetGroupNo = Me.txtAssetGroupNo.Text

        If myClsBG0680BL.DeleteData() = True Then
            MessageBox.Show("Asset group no was deleted", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            SearchDataGrid()
        Else
            MessageBox.Show("There are error between delete Asset group no", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
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
                    myClsBG0680BL.AssetGroupNo = row(0).ToString
                    myClsBG0680BL.AssetGroupName = row(1).ToString
                    myClsBG0680BL.AssetProj = row(2).ToString
                    myClsBG0680BL.AssetCate = row(3).ToString
                    myClsBG0680BL.UpdateUserId = p_strUserId

                    If myClsBG0680BL.UpdateExcelData(conn, trans) = True Then
                        success = True
                    Else
                        success = False
                        Exit For
                    End If

                Else
                    '// insert
                    myClsBG0680BL.AssetGroupNo = row(0).ToString
                    myClsBG0680BL.AssetGroupName = row(1).ToString
                    myClsBG0680BL.AssetProj = row(2).ToString
                    myClsBG0680BL.AssetCate = row(3).ToString
                    myClsBG0680BL.CreateUserId = p_strUserId

                    If myClsBG0680BL.insertExcelData(conn, trans) = True Then
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
                WriteTransactionLog(CStr(enumOperationCd.EditAssetGroupMaster), "", "", "", "", "", "")

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
        sdlgSave.FileName = "AssetGroupNoMaster_" & Format(Date.Now, "yyyyMMdd")
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

        excel.Cells(1, 1) = "Asset group no"
        excel.Cells(1, 2) = "Asset group name"
        excel.Cells(1, 3) = "Asset project"
        excel.Cells(1, 4) = "Asset category"
        excel.Cells(1, 5) = "Create user id"
        excel.Cells(1, 6) = "Create date"
        excel.Cells(1, 7) = "Update user id"
        excel.Cells(1, 8) = "Update date"

        For Each dr In dtMaster.Rows
            rowIndex = rowIndex + 1
            colIndex = 0
            For Each dc In dtMaster.Columns
                If dc.ColumnName <> "ASSET_PROJECT_TXT" AndAlso dc.ColumnName <> "ASSET_CATEGORY_TXT" Then
                    colIndex = colIndex + 1
                    excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
                End If
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