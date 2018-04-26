Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class frmBG0650

#Region "Variable"
    Private myClsBG0650BL As New clsBG0650BL
    Private isInsert As Boolean = False
#End Region

#Region "Overrides Function"
    Public Sub New(ByRef frmParent As Form, ByVal strFormName As String, ByVal blnMaximize As Boolean)
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
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
        Me.grvMaster.Columns("PERSON_IN_CHARGE_NO").HeaderText = "Person in charge no"
        Me.grvMaster.Columns("PERSON_IN_CHARGE_NAME").HeaderText = "Person in charge name"
        Me.grvMaster.Columns("CREATE_USER_ID").HeaderText = "Create user id"
        Me.grvMaster.Columns("CREATE_DATE").HeaderText = "Create date"
        Me.grvMaster.Columns("UPDATE_USER_ID").HeaderText = "Update user id"
        Me.grvMaster.Columns("UPDATE_DATE").HeaderText = "Update date"
    End Sub

    Public Sub setGridHeaderProperty()
        Me.grvMaster.Columns("PERSON_IN_CHARGE_NO").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        Me.grvMaster.Columns("PERSON_IN_CHARGE_NO").MinimumWidth = 100

        Me.grvMaster.Columns("PERSON_IN_CHARGE_NAME").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        Me.grvMaster.Columns("PERSON_IN_CHARGE_NAME").MinimumWidth = 248

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

        myClsBG0650BL.PersonNoFilter = Me.txtPicNoFilter.Text.Trim
        myClsBG0650BL.PersonNameFilter = Me.txtPicNameFilter.Text.Trim

        If myClsBG0650BL.searchDatagrid = True Then
            Me.grvMaster.DataSource = myClsBG0650BL.DtResult
            setGridHeaderText()
            setGridHeaderProperty()
        Else
            MessageBox.Show("Error: Can not get person in charge information!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
    End Sub

    Public Sub setText(ByVal intRow As Integer)
        isInsert = False
        Me.txtPicNo.Enabled = True
        Me.txtPicNo.Text = CStr(Me.grvMaster.Rows(intRow).Cells("PERSON_IN_CHARGE_NO").Value)
        Me.txtPicName.Text = CStr(Me.grvMaster.Rows(intRow).Cells("PERSON_IN_CHARGE_NAME").Value)
    End Sub

    Public Function checkData(ByVal strPersonNo As String) As Boolean

        myClsBG0650BL.PersonNo = strPersonNo

        If myClsBG0650BL.checkData() = True Then
            If myClsBG0650BL.DtResult.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If

        Else
            Return False
        End If

    End Function

    Private Sub clearFilter()
        Me.txtPicNoFilter.Text = ""
        Me.txtPicNameFilter.Text = ""
    End Sub

#End Region

#Region "Control Event"

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Me.txtPicNo.Text = String.Empty
        Me.txtPicName.Text = String.Empty
        Me.txtPicNo.Enabled = True
        isInsert = True
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        '// Remember current state
        Dim intFirstRow As Integer = grvMaster.FirstDisplayedCell.RowIndex
        Dim intFirstCol As Integer = grvMaster.FirstDisplayedCell.ColumnIndex
        Dim intSelRow As Integer = grvMaster.SelectedCells(0).RowIndex
        Dim intSelCol As Integer = grvMaster.SelectedCells(0).ColumnIndex

        If isInsert = True Then
            If MessageBox.Show("Are you sure to create new person in charge?", Me.Text, MessageBoxButtons.YesNo, _
                          MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            Dim strAccount As String = Me.txtPicNo.Text
            myClsBG0650BL.PersonNo = Me.txtPicNo.Text
            myClsBG0650BL.PersonName = Me.txtPicName.Text
            myClsBG0650BL.CreateUserId = p_strUserId

            If myClsBG0650BL.insertOneData() = True Then
                MessageBox.Show("Person in charge was created", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                '// Write Transaction Log
                WriteTransactionLog(CStr(enumOperationCd.EditPersonInChargeMaster), "", "", "", "", "", "")

                SearchDataGrid()

                Dim dt As DataTable = CType(Me.grvMaster.DataSource, DataTable)
                dt.PrimaryKey = New DataColumn() {dt.Columns(0)}
            Else
                If myClsBG0650BL.DtResult.Rows.Count >= 1 Then
                    MessageBox.Show("Person in charge " & Me.txtPicNo.Text & " is exist", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("There are error between save Person in charge", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        Else
            If MessageBox.Show("Are you sure to update person in charge?", Me.Text, MessageBoxButtons.YesNo, _
                          MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            Dim dt As DataTable = CType(Me.grvMaster.DataSource, DataTable)
            dt.PrimaryKey = New DataColumn() {dt.Columns(0)}

            myClsBG0650BL.PersonNo = Me.txtPicNo.Text
            myClsBG0650BL.PersonName = Me.txtPicName.Text
            myClsBG0650BL.UpdateUserId = p_strUserId

            If myClsBG0650BL.UpdateOneData() = True Then
                MessageBox.Show("Person in charge was updated", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                '// Write Transaction Log
                WriteTransactionLog(CStr(enumOperationCd.EditPersonInChargeMaster), "", "", "", "", "", "")

                SearchDataGrid()
            Else
                MessageBox.Show("There are error between save Person in charge", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
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
                setText(intSelRow)
            End If
        End If


    End Sub

    Private Sub grvMaster_RowEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grvMaster.RowEnter
        If Not Me.grvMaster Is Nothing Then
            setText(e.RowIndex)
        End If

    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        If MessageBox.Show("Are you sure to delete person in charge?", Me.Text, MessageBoxButtons.YesNo, _
                         MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        myClsBG0650BL.PersonNo = Me.txtPicNo.Text

        If myClsBG0650BL.DeleteData() = True Then
            MessageBox.Show("Person in charge was deleted", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            '// Write Transaction Log
            WriteTransactionLog(CStr(enumOperationCd.EditPersonInChargeMaster), "", "", "", "", "", "")

            SearchDataGrid()
        Else
            MessageBox.Show("There are error between delete Person in charge", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub cmdClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub cmdExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExport.Click
        If Me.grvMaster.Columns.Count = 0 Or Me.grvMaster.Rows.Count = 0 Then
            Exit Sub
        End If

        'Show dialog box
        Dim sdlgSave As SaveFileDialog = New SaveFileDialog
        sdlgSave.FileName = "PersonInChargeMaster_" & Format(Date.Now, "yyyyMMdd")
        sdlgSave.Filter = "Microsoft Excel Workbook (*.xls)|*.xls"

        Dim dlrConfirm As DialogResult = sdlgSave.ShowDialog()
        If dlrConfirm.Equals(DialogResult.Cancel) Then
            Exit Sub
        End If

        'Creating dataset to export
        Dim dset As New DataSet

        'add table to dataset
        dset.Tables.Add()

        'add column to that table
        For i As Integer = 0 To grvMaster.ColumnCount - 1
            dset.Tables(0).Columns.Add(grvMaster.Columns(i).HeaderText)
        Next

        'add rows to the table
        Dim dr1 As DataRow
        Dim rowscount As Integer = 0
        For rowscount = 0 To grvMaster.RowCount - 1
            dr1 = dset.Tables(0).NewRow
            For j As Integer = 0 To grvMaster.Columns.Count - 1
                dr1(j) = grvMaster.Rows(rowscount).Cells(j).Value
            Next
            dset.Tables(0).Rows.Add(dr1)
        Next

        Dim excel As New Microsoft.Office.Interop.Excel.Application
        Dim wBook As Microsoft.Office.Interop.Excel.Workbook
        Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet

        wBook = excel.Workbooks.Add()
        wSheet = CType(wBook.ActiveSheet(), Microsoft.Office.Interop.Excel.Worksheet)

        excel.Range("A1", "A" & (rowscount + 1).ToString).NumberFormat = "@"

        Dim dt As System.Data.DataTable = dset.Tables(0)
        Dim dc As System.Data.DataColumn
        Dim dr As System.Data.DataRow
        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = 0

        For Each dc In dt.Columns
            colIndex = colIndex + 1
            excel.Cells(1, colIndex) = dc.ColumnName
        Next

        For Each dr In dt.Rows
            rowIndex = rowIndex + 1
            colIndex = 0
            For Each dc In dt.Columns
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
    End Sub

    Private Sub cmdImport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdImport.Click

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
                    'update
                    myClsBG0650BL.PersonNo = row(0).ToString
                    myClsBG0650BL.PersonName = row(1).ToString
                    myClsBG0650BL.UpdateUserId = p_strUserId

                    If myClsBG0650BL.UpdateExcelData(conn, trans) = True Then
                        success = True
                    Else
                        success = False
                        Exit For
                    End If

                Else
                    'insert
                    myClsBG0650BL.PersonNo = row(0).ToString
                    myClsBG0650BL.PersonName = row(1).ToString
                    myClsBG0650BL.CreateUserId = p_strUserId

                    If myClsBG0650BL.insertExcelData(conn, trans) = True Then
                        success = True
                    Else
                        success = False
                        Exit For
                    End If
                End If
            Next

            If success Then
                trans.Commit()

                '// Write Transaction Log
                WriteTransactionLog(CStr(enumOperationCd.EditPersonInChargeMaster), "", "", "", "", "", "")

                SearchDataGrid()
            Else
                trans.Rollback()
            End If
        Catch ex As Exception
            MessageBox.Show("There are error between save file", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            trans.Rollback()
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub cmdFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFilter.Click
        SearchDataGrid()
    End Sub

    Private Sub cmdClearFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearFilter.Click
        clearFilter()
    End Sub

#End Region

End Class