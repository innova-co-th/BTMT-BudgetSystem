Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class frmBG0610

#Region "Variable"
    Private myClsBG0610BL As New clsBG0610BL
    Private myOperationCd As Integer = OperationCd.AddNew
    Private myControlLoadingFlg As Boolean = False

#End Region

#Region "Enumeration"
    Private Enum OperationCd As Integer
        AddNew = 1
        Edit = 2
    End Enum
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
    Private Sub ClearInfo()
        txtUserID.Text = ""
        txtUserName.Text = ""
        txtPwd1.Text = ""
        txtPwd2.Text = ""
        txtEmail.Text = ""
        cboUserLevel.SelectedIndex = cboUserLevel.Items.Count - 1
        cboUserPIC.SelectedIndex = 0
        chkExpired.Checked = False
    End Sub

    Private Sub BeginAddUser()
        '// Clear all user info
        ClearInfo()

        '// Set focus for add new user id
        txtUserID.Enabled = True
        txtUserID.Focus()

        '// Set Operation Code
        myOperationCd = OperationCd.AddNew
    End Sub

    Private Sub BeginEditUser()
        '// Show User Info
        If grvMaster.Rows.Count > 0 Then

            Dim r As Integer = grvMaster.CurrentCell.RowIndex
            txtUserID.Text = CStr(grvMaster.Item(0, r).Value)
            txtUserID.Enabled = False
            txtUserName.Text = CStr(grvMaster.Item(1, r).Value)
            txtPwd1.Text = CStr(grvMaster.Item(2, r).Value)
            txtPwd2.Text = CStr(grvMaster.Item(2, r).Value)
            cboUserLevel.Text = CStr(Nz(grvMaster.Item(3, r).Value, ""))
            cboUserPIC.Text = CStr(grvMaster.Item(4, r).Value)
            txtEmail.Text = CStr(Nz(grvMaster.Item(5, r).Value))
            If CStr(grvMaster.Item(6, r).Value) = "Yes" Then
                chkExpired.Checked = True
            Else
                chkExpired.Checked = False
            End If

            '// Set Operation Code
            myOperationCd = OperationCd.Edit
        Else

            ClearInfo()

        End If

    End Sub

    Private Sub SaveNewData()
        If MessageBox.Show("Are you sure to add new user?", Me.Text, MessageBoxButtons.YesNo, _
                  MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        '// Check Input data
        If txtUserID.Text.Trim = "" Then
            MessageBox.Show("Please input User ID.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        myClsBG0610BL.UserId = txtUserID.Text.Trim
        If myClsBG0610BL.CheckUserExist = True Then
            MessageBox.Show("This User ID already exist.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        If txtUserName.Text.Trim = "" Then
            MessageBox.Show("Please input User Name.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        If txtPwd1.Text.Trim = "" Then
            MessageBox.Show("Please input Password.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        If txtPwd1.Text.Trim <> txtPwd2.Text Then
            MessageBox.Show("Password and Confirm password do not match.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        myControlLoadingFlg = True

        '// Set Parameters
        myClsBG0610BL.UserId = txtUserID.Text
        myClsBG0610BL.UserLevel = CStr(cboUserLevel.SelectedIndex)
        myClsBG0610BL.UserName = txtUserName.Text
        myClsBG0610BL.Password = txtPwd1.Text
        myClsBG0610BL.Email = txtEmail.Text
        If chkExpired.Checked = True Then
            myClsBG0610BL.ExpireFlg = "1"
        Else
            myClsBG0610BL.ExpireFlg = "0"
        End If
        myClsBG0610BL.UserId2 = p_strUserId

        '// Set Person In Charge
        If cboUserLevel.SelectedIndex = enumUserLevel.SystemAdministrator Then
            myClsBG0610BL.UserPIC = "0000"

        ElseIf cboUserLevel.SelectedIndex = enumUserLevel.AccountUser Then
            myClsBG0610BL.UserPIC = "210"

        ElseIf cboUserLevel.SelectedIndex = enumUserLevel.ManagingDirector Then
            myClsBG0610BL.UserPIC = "BTMT3"

        ElseIf cboUserLevel.SelectedIndex = enumUserLevel.AdminSaleDirector Then
            myClsBG0610BL.UserPIC = "BTMT10"

        Else
            myClsBG0610BL.UserPIC = CStr(cboUserPIC.SelectedValue)
        End If

        '// Call Function
        If myClsBG0610BL.CreateNewUser() = True Then
            MessageBox.Show("Add user completed", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            '// Write Transaction Log
            WriteTransactionLog(CStr(enumOperationCd.EditUserMaster), "", "", "", "", "", "")

        End If

        '// Load User List
        'If myClsBG0610BL.SearchAllUser = True Then
        '    grvMaster.DataSource = myClsBG0610BL.UserList
        '    If chkHideExpiredUser.Checked Then
        '        HideExpiredUser()
        '    End If
        'End If
        FilterData()

        '// Back to new user mode
        BeginAddUser()

        myControlLoadingFlg = False
    End Sub

    Private Sub SaveEditData()
        If MessageBox.Show("Are you sure to save user data?", Me.Text, MessageBoxButtons.YesNo, _
                  MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        '// Check Input data
        If txtUserID.Text.Trim = "" Then
            MessageBox.Show("Please input User ID.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        If txtUserName.Text.Trim = "" Then
            MessageBox.Show("Please input User Name.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        If txtPwd1.Text.Trim = "" Then
            MessageBox.Show("Please input Password.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        If txtPwd1.Text.Trim <> txtPwd2.Text.Trim Then
            MessageBox.Show("Password and Confirm password do not match.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        myControlLoadingFlg = True

        '// Set Parameters
        myClsBG0610BL.UserId = CStr(grvMaster.Item(0, grvMaster.CurrentCell.RowIndex).Value)
        myClsBG0610BL.UserLevel = CStr(cboUserLevel.SelectedIndex)
        myClsBG0610BL.UserName = txtUserName.Text
        myClsBG0610BL.Password = txtPwd1.Text

        '// Set Person In Charge
        If cboUserLevel.SelectedIndex = enumUserLevel.SystemAdministrator Then
            myClsBG0610BL.UserPIC = "0000"

        ElseIf cboUserLevel.SelectedIndex = enumUserLevel.AccountUser Then
            myClsBG0610BL.UserPIC = "210"

        ElseIf cboUserLevel.SelectedIndex = enumUserLevel.ManagingDirector Then
            myClsBG0610BL.UserPIC = "BTMT3"

        ElseIf cboUserLevel.SelectedIndex = enumUserLevel.AdminSaleDirector Then
            myClsBG0610BL.UserPIC = "BTMT10"

        Else
            myClsBG0610BL.UserPIC = CStr(cboUserPIC.SelectedValue)
        End If

        myClsBG0610BL.Email = txtEmail.Text
        If chkExpired.Checked = True Then
            myClsBG0610BL.ExpireFlg = "1"
        Else
            myClsBG0610BL.ExpireFlg = "0"
        End If
        myClsBG0610BL.UserId2 = p_strUserId

        '// Call Function
        If myClsBG0610BL.UpdateUserData() = True Then
            MessageBox.Show("Save user data completed", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            '// Write Transaction Log
            WriteTransactionLog(CStr(enumOperationCd.EditUserMaster), "", "", "", "", "", "")

        End If

        '// Remember current state
        Dim intFirstRow As Integer = grvMaster.FirstDisplayedCell.RowIndex
        Dim intFirstCol As Integer = grvMaster.FirstDisplayedCell.ColumnIndex
        Dim intSelRow As Integer = grvMaster.SelectedCells(0).RowIndex
        Dim intSelCol As Integer = grvMaster.SelectedCells(0).ColumnIndex

        '// Load User List
        'If myClsBG0610BL.SearchAllUser = True Then
        '    grvMaster.DataSource = myClsBG0610BL.UserList
        'End If
        FilterData()

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

        If chkHideExpiredUser.Checked Then
            HideExpiredUser()
        End If

        myControlLoadingFlg = False
    End Sub

    Private Sub SaveChange()
        If myOperationCd = OperationCd.AddNew Then
            '// Save new user info
            SaveNewData()
        Else
            '// Save changed user info
            SaveEditData()
        End If
    End Sub

    Private Sub HideExpiredUser()
        If myClsBG0610BL.UserList IsNot Nothing Then
            myClsBG0610BL.UserList.DefaultView.RowFilter = "EXPIRED = 'No'"
            grvMaster.DataSource = myClsBG0610BL.UserList.DefaultView
        End If
    End Sub

    Private Sub UnhideExpiredUser()
        If myClsBG0610BL.UserList IsNot Nothing Then
            myClsBG0610BL.UserList.DefaultView.RowFilter = ""
            grvMaster.DataSource = myClsBG0610BL.UserList.DefaultView
        End If
    End Sub

    Private Sub clearFilter()

        Me.txtUserIDFiter.Text = ""
        Me.txtUserNameFilter.Text = ""
        Me.cboUserLevelFilter.SelectedIndex = 0
        Me.cboUserPICFilter.SelectedIndex = 0
        Me.chkHideExpiredUser.Checked = True

    End Sub

    Private Sub loadCombo()
        '// Load User level
        If myClsBG0610BL.SearchAllUserLevel = True Then
            cboUserLevel.DataSource = myClsBG0610BL.UserList
            cboUserLevel.DisplayMember = "USER_LEVEL_NAME"
            cboUserLevel.ValueMember = "USER_LEVEL_ID"
        Else
            cboUserLevel.Items.Clear()
            cboUserLevel.Items.Add("")
        End If

        '// Load User PIC
        If myClsBG0610BL.SearchAllUserPIC = True Then
            '//ize edit 2015/05/11
            Dim dtRow0000 As DataRow = myClsBG0610BL.UserList.NewRow()
            dtRow0000("PERSON_IN_CHARGE_NO") = "0000"
            myClsBG0610BL.UserList.Rows.Add(dtRow0000)
            '//ize edit 2015/05/11
            cboUserPIC.DataSource = myClsBG0610BL.UserList
            cboUserPIC.DisplayMember = "PERSON_IN_CHARGE_NO"
            cboUserPIC.ValueMember = "PERSON_IN_CHARGE_NO"

        Else
            cboUserPIC.Items.Clear()
            cboUserPIC.Items.Add("")
        End If
    End Sub

    Private Sub loadComboFilter()
        Dim drNew As DataRow
        Dim dtUserLevel As DataTable

        '// Load User level
        If myClsBG0610BL.SearchAllUserLevel = True Then

            dtUserLevel = New DataTable
            dtUserLevel.Columns.Add("USER_LEVEL_ID", GetType(String))
            dtUserLevel.Columns.Add("USER_LEVEL_NAME", GetType(String))

            For Each dr As DataRow In myClsBG0610BL.UserList.Rows
                drNew = dtUserLevel.NewRow
                drNew("USER_LEVEL_ID") = CStr(dr("USER_LEVEL_ID"))
                drNew("USER_LEVEL_NAME") = CStr(dr("USER_LEVEL_NAME"))
                dtUserLevel.Rows.Add(drNew)
            Next

            drNew = dtUserLevel.NewRow
            drNew("USER_LEVEL_ID") = ""
            drNew("USER_LEVEL_NAME") = "All"
            dtUserLevel.Rows.InsertAt(drNew, 0)

            cboUserLevelFilter.DataSource = dtUserLevel
            cboUserLevelFilter.DisplayMember = "USER_LEVEL_NAME"
            cboUserLevelFilter.ValueMember = "USER_LEVEL_ID"
        Else
            cboUserLevelFilter.Items.Clear()
            cboUserLevelFilter.Items.Add("")
        End If

        '// Load User PIC
        If myClsBG0610BL.SearchAllUserPIC = True Then

            drNew = myClsBG0610BL.UserList.NewRow
            drNew("PERSON_IN_CHARGE_NO") = ""
            drNew("PERSON_IN_CHARGE_NO") = "All"
            myClsBG0610BL.UserList.Rows.InsertAt(drNew, 0)

            cboUserPICFilter.DataSource = myClsBG0610BL.UserList
            cboUserPICFilter.DisplayMember = "PERSON_IN_CHARGE_NO"
            cboUserPICFilter.ValueMember = "PERSON_IN_CHARGE_NO"
        Else
            cboUserPICFilter.Items.Clear()
            cboUserPICFilter.Items.Add("")
        End If
    End Sub

    Private Sub FilterData()

        myClsBG0610BL.UserIdFilter = Me.txtUserIDFiter.Text.Trim
        myClsBG0610BL.UserNameFilter = Me.txtUserNameFilter.Text.Trim
        myClsBG0610BL.UserLevelFilter = CStr(Me.cboUserLevelFilter.SelectedValue)

        '// Set Person In Charge
        If cboUserLevelFilter.SelectedIndex = enumUserLevel.SystemAdministrator + 1 Then
            myClsBG0610BL.UserPICFilter = "0000"

        ElseIf cboUserLevelFilter.SelectedIndex = enumUserLevel.AccountUser + 1 Then
            myClsBG0610BL.UserPICFilter = "210"

        ElseIf cboUserLevelFilter.SelectedIndex = enumUserLevel.ManagingDirector + 1 Then
            myClsBG0610BL.UserPICFilter = "BTMT3"

        ElseIf cboUserLevelFilter.SelectedIndex = enumUserLevel.AdminSaleDirector + 1 Then
            myClsBG0610BL.UserPICFilter = "BTMT10"

        Else
            myClsBG0610BL.UserPICFilter = CStr(cboUserPICFilter.SelectedValue)
        End If

        If myClsBG0610BL.SearchAllUser = True Then
            grvMaster.DataSource = myClsBG0610BL.UserList
            If chkHideExpiredUser.Checked Then
                HideExpiredUser()
            End If
        Else
            grvMaster.DataSource = Nothing
        End If

    End Sub
#End Region

#Region "Control Event"
    Private Sub frmBG0610_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        myControlLoadingFlg = True

        loadComboFilter()

        loadCombo()

        '// Load User List
        FilterData()

        '// Begin with Add new user mode
        BeginEditUser()

        '// Reset Flags
        myControlLoadingFlg = False
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub grvMaster_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grvMaster.CurrentCellChanged
        If grvMaster.CurrentCell IsNot Nothing And myControlLoadingFlg = False Then
            BeginEditUser()
        End If
    End Sub

    Private Sub cboUserLevel_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboUserLevel.SelectedIndexChanged
        If cboUserLevel.SelectedIndex > enumUserLevel.AdminSaleDirector Then
            cboUserPIC.Enabled = True
        Else
            cboUserPIC.Enabled = False
        End If
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        BeginAddUser()
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        SaveChange()
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

                myClsBG0610BL.UserId = CStr(row(0))

                If myClsBG0610BL.CheckUserExist = True Then
                    'update
                    myClsBG0610BL.UserId = row(0).ToString
                    myClsBG0610BL.UserLevel = row(1).ToString
                    myClsBG0610BL.UserName = row(2).ToString
                    myClsBG0610BL.Password = row(3).ToString
                    If row(4).ToString.Trim = "" Then
                        myClsBG0610BL.UserPIC = "0000"
                    Else
                        myClsBG0610BL.UserPIC = row(4).ToString
                    End If
                    myClsBG0610BL.Email = row(5).ToString
                    If row(6).ToString.Trim = "" Then
                        myClsBG0610BL.ExpireFlg = "0"
                    Else
                        myClsBG0610BL.ExpireFlg = row(6).ToString
                    End If

                    myClsBG0610BL.UserId2 = p_strUserId

                    If myClsBG0610BL.UpdateUserDataImport(conn, trans) = True Then
                        success = True
                    Else
                        success = False
                        Exit For
                    End If

                Else
                    'insert
                    myClsBG0610BL.UserId = row(0).ToString
                    myClsBG0610BL.UserLevel = row(1).ToString
                    myClsBG0610BL.UserName = row(2).ToString
                    myClsBG0610BL.Password = row(3).ToString
                    myClsBG0610BL.UserPIC = row(4).ToString
                    myClsBG0610BL.Email = row(5).ToString
                    myClsBG0610BL.ExpireFlg = row(6).ToString
                    myClsBG0610BL.UserId2 = p_strUserId

                    If myClsBG0610BL.CreateNewUserImport(conn, trans) = True Then
                        success = True
                    Else
                        success = False
                        Exit For
                    End If
                End If
            Next

            If success Then
                trans.Commit()
                '// Load User List
                'If myClsBG0610BL.SearchAllUser = True Then
                '    grvMaster.DataSource = myClsBG0610BL.UserList
                '    If chkHideExpiredUser.Checked Then
                '        HideExpiredUser()
                '    End If
                'Else
                '    grvMaster.DataSource = Nothing
                'End If
                FilterData()

                '// Write Transaction Log
                WriteTransactionLog(CStr(enumOperationCd.EditUserMaster), "", "", "", "", "", "")

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

    Private Sub cmdExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExport.Click
        If Me.grvMaster.Columns.Count = 0 Or Me.grvMaster.Rows.Count = 0 Then
            Exit Sub
        End If

        '// Show dialog box
        Dim sdlgSave As SaveFileDialog = New SaveFileDialog
        sdlgSave.FileName = "UserMaster_" & Format(Date.Now, "yyyyMMdd")
        sdlgSave.Filter = "Microsoft Excel Workbook (*.xls)|*.xls"

        Dim dlrConfirm As DialogResult = sdlgSave.ShowDialog()
        If dlrConfirm.Equals(DialogResult.Cancel) Then
            Exit Sub
        End If

        Dim dt As New DataTable
        If myClsBG0610BL.SearchUserExcel = True Then
            dt = myClsBG0610BL.UserList
        End If

        Dim excel As New Microsoft.Office.Interop.Excel.Application
        Dim wBook As Microsoft.Office.Interop.Excel.Workbook
        Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet

        wBook = excel.Workbooks.Add()
        wSheet = CType(wBook.ActiveSheet(), Microsoft.Office.Interop.Excel.Worksheet)

        excel.Range("A1", "A" & (dt.Rows.Count).ToString).NumberFormat = "@"
        excel.Range("C1", "C" & (dt.Rows.Count).ToString).NumberFormat = "@"
        excel.Range("D1", "D" & (dt.Rows.Count).ToString).NumberFormat = "@"
        excel.Range("E1", "E" & (dt.Rows.Count).ToString).NumberFormat = "@"

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

    Private Sub chkHideExpiredUser_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkHideExpiredUser.CheckedChanged
        If chkHideExpiredUser.Checked Then
            HideExpiredUser()
        Else
            UnhideExpiredUser()
        End If
    End Sub

    Private Sub cmdFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFilter.Click

        myControlLoadingFlg = True

        '// Load User List
        FilterData()

        '// Begin with Add new user mode
        BeginEditUser()

        '// Reset Flags
        myControlLoadingFlg = False

    End Sub

    Private Sub cmdClearFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearFilter.Click
        clearFilter()
    End Sub

#End Region


End Class