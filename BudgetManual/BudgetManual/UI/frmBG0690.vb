Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Text

Public Class frmBG0690

#Region "Variable"
    Private myClsBG0690BL As New clsBG0690BL
    Private isInsert As Boolean = False
    Private OldChildNo As String = String.Empty
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

    End Sub
#End Region

#Region "Function"
    Public Sub setGridHeaderText()
        Me.grvMaster.Columns("PIC_PARENT_NO").HeaderText = "Pic parent no"
        Me.grvMaster.Columns("PERSON_IN_CHARGE_NAME_P").HeaderText = "Parent person in charge name"
        Me.grvMaster.Columns("PIC_CHILD_NO").HeaderText = "Pic child no"
        Me.grvMaster.Columns("PERSON_IN_CHARGE_NAME_C").HeaderText = "Child person in charge name"
        Me.grvMaster.Columns("CREATE_USER_ID").HeaderText = "Create user id"
        Me.grvMaster.Columns("CREATE_DATE").HeaderText = "Create date"
        Me.grvMaster.Columns("UPDATE_USER_ID").HeaderText = "Update user id"
        Me.grvMaster.Columns("UPDATE_DATE").HeaderText = "Update date"
    End Sub

    Public Sub setGridHeaderProperty()
        Me.grvMaster.Columns("PIC_PARENT_NO").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        Me.grvMaster.Columns("PIC_PARENT_NO").MinimumWidth = 90

        Me.grvMaster.Columns("PERSON_IN_CHARGE_NAME_P").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        Me.grvMaster.Columns("PERSON_IN_CHARGE_NAME_P").MinimumWidth = 160

        Me.grvMaster.Columns("PIC_CHILD_NO").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        Me.grvMaster.Columns("PIC_CHILD_NO").MinimumWidth = 90

        Me.grvMaster.Columns("PERSON_IN_CHARGE_NAME_C").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        Me.grvMaster.Columns("PERSON_IN_CHARGE_NAME_C").MinimumWidth = 150

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

        myClsBG0690BL.ParentNoFilter = Me.cboParentPICFilter.SelectedValue.ToString
        myClsBG0690BL.ChildNoFilter = Me.cboChildPICFilter.SelectedValue.ToString

        If myClsBG0690BL.searchDatagrid = True Then
            Me.grvMaster.DataSource = myClsBG0690BL.dtResult
            setGridHeaderText()
            setGridHeaderProperty()
        Else
            MessageBox.Show("Error: Can not get department information!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
    End Sub
    Public Sub SearchCombo()

        If myClsBG0690BL.searchCombo = True Then
            Me.cboParentPIC.DataSource = myClsBG0690BL.PicList
            Me.cboParentPIC.DisplayMember = "PERSON_IN_CHARGE_NAME"
            Me.cboParentPIC.ValueMember = "PERSON_IN_CHARGE_NO"
        Else
            Me.cboParentPIC.Items.Clear()
            Me.cboParentPIC.Items.Add("")
        End If

        If myClsBG0690BL.searchCombo = True Then
            Me.cboChildPIC.DataSource = myClsBG0690BL.PicList
            Me.cboChildPIC.DisplayMember = "PERSON_IN_CHARGE_NAME"
            Me.cboChildPIC.ValueMember = "PERSON_IN_CHARGE_NO"
        Else
            Me.cboChildPIC.Items.Clear()
            Me.cboChildPIC.Items.Add("")
        End If

        setText(0)
    End Sub

    Public Sub setText(ByVal intRow As Integer)
        isInsert = False
        Me.cboParentPIC.SelectedValue = CStr(Me.grvMaster.Rows(intRow).Cells("PIC_PARENT_NO").Value)
        Me.cboChildPIC.SelectedValue = CStr(Me.grvMaster.Rows(intRow).Cells("PIC_CHILD_NO").Value)
        OldChildNo = CStr(Me.grvMaster.Rows(intRow).Cells("PIC_CHILD_NO").Value)
    End Sub

    Public Sub loadComboFilter()
        Dim drNew As DataRow

        If myClsBG0690BL.searchCombo = True Then

            drNew = myClsBG0690BL.PicList.NewRow
            drNew("PERSON_IN_CHARGE_NO") = ""
            drNew("PERSON_IN_CHARGE_NAME") = "All"
            myClsBG0690BL.PicList.Rows.InsertAt(drNew, 0)

            Me.cboParentPICFilter.DataSource = myClsBG0690BL.PicList
            Me.cboParentPICFilter.DisplayMember = "PERSON_IN_CHARGE_NAME"
            Me.cboParentPICFilter.ValueMember = "PERSON_IN_CHARGE_NO"
        Else
            Me.cboParentPICFilter.Items.Clear()
            Me.cboParentPICFilter.Items.Add("")
        End If

        If myClsBG0690BL.searchCombo = True Then

            drNew = myClsBG0690BL.PicList.NewRow
            drNew("PERSON_IN_CHARGE_NO") = ""
            drNew("PERSON_IN_CHARGE_NAME") = "All"
            myClsBG0690BL.PicList.Rows.InsertAt(drNew, 0)

            Me.cboChildPICFilter.DataSource = myClsBG0690BL.PicList
            Me.cboChildPICFilter.DisplayMember = "PERSON_IN_CHARGE_NAME"
            Me.cboChildPICFilter.ValueMember = "PERSON_IN_CHARGE_NO"
        Else
            Me.cboChildPICFilter.Items.Clear()
            Me.cboChildPICFilter.Items.Add("")
        End If

    End Sub

    Private Sub clearFilter()
        Me.cboParentPICFilter.SelectedIndex = 0
        Me.cboChildPICFilter.SelectedIndex = 0
    End Sub

#End Region

#Region "Control Event"

    Private Sub frmBG0690_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadComboFilter()
        SearchDataGrid()
        SearchCombo()
        Me.cboParentPIC.Enabled = False
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub grvMaster_RowEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grvMaster.RowEnter
        If grvMaster.CurrentCell IsNot Nothing Then
            setText(e.RowIndex)
            Me.cboParentPIC.Enabled = False
            isInsert = False
        End If
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Me.cboParentPIC.Enabled = True
        Me.cboParentPIC.SelectedIndex = 0
        Me.cboChildPIC.SelectedIndex = 0
        isInsert = True
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        If MessageBox.Show("Are you sure to delete Child pic no?", Me.Text, MessageBoxButtons.YesNo, _
                    MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        myClsBG0690BL.ParentNo = CStr(Me.cboParentPIC.SelectedValue)
        myClsBG0690BL.ChildNo = CStr(Me.cboChildPIC.SelectedValue)

        If myClsBG0690BL.DeleteData() = True Then
            MessageBox.Show("Child pic  no was deleted", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            SearchDataGrid()
        Else
            MessageBox.Show("There are error between delete Child pic no", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        '// Remember current state
        Dim intFirstRow As Integer = grvMaster.FirstDisplayedCell.RowIndex
        Dim intFirstCol As Integer = grvMaster.FirstDisplayedCell.ColumnIndex
        Dim intSelRow As Integer = grvMaster.SelectedCells(0).RowIndex
        Dim intSelCol As Integer = grvMaster.SelectedCells(0).ColumnIndex

        If isInsert = True Then
            If MessageBox.Show("Are you sure to create new Child pic no?", Me.Text, MessageBoxButtons.YesNo, _
                          MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            myClsBG0690BL.ParentNo = CStr(Me.cboParentPIC.SelectedValue)
            myClsBG0690BL.ChildNo = CStr(Me.cboChildPIC.SelectedValue)
            myClsBG0690BL.UpdateUserId = p_strUserId

            If myClsBG0690BL.insertOneData() = True Then
                MessageBox.Show("Child pic no was created", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                '// Write Transaction Log
                WriteTransactionLog(CStr(enumOperationCd.EditChildPicMaster), "", "", "", "", "", "")

                SearchDataGrid()
            Else
                MessageBox.Show("There are error between save Child pic no", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Else
            If MessageBox.Show("Are you sure to update Child pic no?", Me.Text, MessageBoxButtons.YesNo, _
                          MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            Dim dt As DataTable = CType(Me.grvMaster.DataSource, DataTable)
            dt.PrimaryKey = New DataColumn() {dt.Columns(0), dt.Columns(2)}

            myClsBG0690BL.ParentNo = CStr(Me.cboParentPIC.SelectedValue)
            myClsBG0690BL.ChildNo = CStr(Me.cboChildPIC.SelectedValue)
            myClsBG0690BL.OldChildNo = Me.OldChildNo
            myClsBG0690BL.UpdateUserId = p_strUserId

            If myClsBG0690BL.UpdateOneData() = True Then
                MessageBox.Show("Child pic no was updated", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                '// Write Transaction Log
                WriteTransactionLog(CStr(enumOperationCd.EditChildPicMaster), "", "", "", "", "", "")

                SearchDataGrid()
            Else
                MessageBox.Show("There are error between save Child pic no", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
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

        Me.cboParentPIC.Enabled = False
    End Sub

    Private Sub cmdExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExport.Click
        If Me.grvMaster.Columns.Count = 0 Or Me.grvMaster.Rows.Count = 0 Then
            Exit Sub
        End If

        'Show dialog box
        Dim sdlgSave As SaveFileDialog = New SaveFileDialog
        sdlgSave.FileName = "ChildPersonInChargeMaster_" & Format(Date.Now, "yyyyMMdd")
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
        excel.Range("C1", "C" & (rowscount + 1).ToString).NumberFormat = "@"

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
            'delete all
            If myClsBG0690BL.DeleteAllData(conn, trans) = True Then
                For Each row As DataRow In dt.Rows

                    'insert
                    myClsBG0690BL.ParentNo = row(0).ToString
                    myClsBG0690BL.ChildNo = row(2).ToString
                    myClsBG0690BL.UpdateUserId = p_strUserId

                    If myClsBG0690BL.insertExcelData(conn, trans) = True Then
                        success = True
                    Else
                        success = False
                        Exit For
                    End If
                Next
            End If

            If success Then
                trans.Commit()

                '// Write Transaction Log
                WriteTransactionLog(CStr(enumOperationCd.EditChildPicMaster), "", "", "", "", "", "")

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