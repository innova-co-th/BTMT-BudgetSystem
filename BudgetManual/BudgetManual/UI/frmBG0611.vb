Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class frmBG0611

#Region "Variable"
    Private myClsBG0611BL As New clsBG0611BL
    Private myCurrUserLevel As Integer = -1
    Private myOperationCd As Integer = OperationCd.AddNew
    Private myControlLoadingFlg As Boolean = False

    Private Const STRING_ALL As String = "All"

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

    Private Sub BeginEditUserLevel()

        '// Show User Level Info
        Dim dtDat As DataTable = CType(grvMaster.DataSource, DataTable)
        If dtDat Is Nothing Then
            Exit Sub
        End If

        If grvMaster.CurrentCell Is Nothing Then
            Exit Sub
        End If

        Dim r As Integer = grvMaster.CurrentCell.RowIndex
        Me.txtUserLevel.Text = CStr(grvMaster.Item(1, r).Value)

        myCurrUserLevel = CInt(dtDat.Rows(r)![USER_LEVEL_ID])

        If CStr(grvMaster.Item(2, r).Value) = "Y" Then
            chkEntry.Checked = True
        Else
            chkEntry.Checked = False
        End If

        If CStr(grvMaster.Item(3, r).Value) = "Y" Then
            chkSubmit.Checked = True
        Else
            chkSubmit.Checked = False
        End If

        If CStr(grvMaster.Item(4, r).Value) = "Y" Then
            chkApprove.Checked = True
        Else
            chkApprove.Checked = False
        End If

        If CStr(grvMaster.Item(5, r).Value) = "Y" Then
            chkAdjust.Checked = True
        Else
            chkAdjust.Checked = False
        End If

        If CStr(grvMaster.Item(6, r).Value) = "Y" Then
            chkAuth1.Checked = True
        Else
            chkAuth1.Checked = False
        End If

        If CStr(grvMaster.Item(7, r).Value) = "Y" Then
            chkAuth2.Checked = True
        Else
            chkAuth2.Checked = False
        End If

        If CStr(grvMaster.Item(8, r).Value) = "Y" Then
            chkImport.Checked = True
        Else
            chkImport.Checked = False
        End If

        If CStr(grvMaster.Item(9, r).Value) = "Y" Then
            chkExport.Checked = True
        Else
            chkExport.Checked = False
        End If

        If CStr(grvMaster.Item(10, r).Value) = "Y" Then
            chkMaster.Checked = True
        Else
            chkMaster.Checked = False
        End If

        If CStr(grvMaster.Item(11, r).Value) = "Y" Then
            chkSystem.Checked = True
        Else
            chkSystem.Checked = False
        End If

        If CStr(grvMaster.Item(12, r).Value) = "Y" Then
            chkView.Checked = True
        Else
            chkView.Checked = False
        End If

        If CStr(grvMaster.Item(13, r).Value) = "Y" Then
            chkDirectInput.Checked = True
        Else
            chkDirectInput.Checked = False
        End If

        cmdDelete.Enabled = True

        '// Set Operation Code
        myOperationCd = OperationCd.Edit

    End Sub

    Private Sub SaveEditData()

        If MessageBox.Show("Are you sure to save the user level?", Me.Text, MessageBoxButtons.YesNo, _
                  MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        myControlLoadingFlg = True

        '// Set Parameters
        myClsBG0611BL.UserLevel = CStr(myCurrUserLevel)

        myClsBG0611BL.UserLevelName = Me.txtUserLevel.Text.Trim

        Dim arrPermission(11) As Boolean
        arrPermission(0) = chkEntry.Checked
        arrPermission(1) = chkSubmit.Checked
        arrPermission(2) = chkApprove.Checked
        arrPermission(3) = chkAdjust.Checked
        arrPermission(4) = chkAuth1.Checked
        arrPermission(5) = chkAuth2.Checked
        arrPermission(6) = chkImport.Checked
        arrPermission(7) = chkExport.Checked
        arrPermission(8) = chkMaster.Checked
        arrPermission(9) = chkSystem.Checked
        arrPermission(10) = chkView.Checked
        arrPermission(11) = chkDirectInput.Checked

        myClsBG0611BL.UserPermissions = arrPermission
        myClsBG0611BL.UserId = p_strUserId

        '// Call Function
        If myClsBG0611BL.UpdateUserLevel() = True Then
            MessageBox.Show("Save user level completed", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            '// Write Transaction Log
            WriteTransactionLog(CStr(enumOperationCd.EditUserMaster), "", "", "", "", "", "")

        End If

        '// Remember current state
        Dim intFirstRow As Integer = grvMaster.FirstDisplayedCell.RowIndex
        Dim intFirstCol As Integer = grvMaster.FirstDisplayedCell.ColumnIndex
        Dim intSelRow As Integer = grvMaster.SelectedCells(0).RowIndex
        Dim intSelCol As Integer = 1

        '// Load User List
        'myClsBG0611BL.GetUserLevelList()
        'grvMaster.DataSource = myClsBG0611BL.UserLevelList
        loadGridData()

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


        myControlLoadingFlg = False
    End Sub

    Private Sub SaveNewData()
        If MessageBox.Show("Are you sure to add new user level?", Me.Text, MessageBoxButtons.YesNo, _
              MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        '// Check Input data
        If txtUserLevel.Text.Trim = "" Then
            MessageBox.Show("Please input User Level.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        '// Get last User Level ID
        myClsBG0611BL.GetLastUserLevelId()


        myControlLoadingFlg = True

        '// Set Parameters
        If myClsBG0611BL.UserLevelId IsNot Nothing AndAlso myClsBG0611BL.UserLevelId.Rows.Count > 0 Then
            myClsBG0611BL.UserLevel = (CInt(myClsBG0611BL.UserLevelId.Rows(0).Item("USER_LEVEL_ID").ToString) + 1).ToString
        Else
            myClsBG0611BL.UserLevel = "0"
        End If

        myClsBG0611BL.UserLevelName = Me.txtUserLevel.Text.Trim

        Dim arrPermission(11) As Boolean
        arrPermission(0) = chkEntry.Checked
        arrPermission(1) = chkSubmit.Checked
        arrPermission(2) = chkApprove.Checked
        arrPermission(3) = chkAdjust.Checked
        arrPermission(4) = chkAuth1.Checked
        arrPermission(5) = chkAuth2.Checked
        arrPermission(6) = chkImport.Checked
        arrPermission(7) = chkExport.Checked
        arrPermission(8) = chkMaster.Checked
        arrPermission(9) = chkSystem.Checked
        arrPermission(10) = chkView.Checked
        arrPermission(11) = chkDirectInput.Checked

        myClsBG0611BL.UserPermissions = arrPermission
        myClsBG0611BL.UserId = p_strUserId

        '// Call Function
        If myClsBG0611BL.AddUserLevel() = True Then
            MessageBox.Show("Add user level completed", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            '// Write Transaction Log
            WriteTransactionLog(CStr(enumOperationCd.EditUserMaster), "", "", "", "", "", "")

        End If

        '// Load User List
        'myClsBG0611BL.GetUserLevelList()
        'grvMaster.DataSource = myClsBG0611BL.UserLevelList
        loadGridData()

        '// Back to new user mode
        BeginAddUserLevel()

        myControlLoadingFlg = False
    End Sub

    Private Sub ClearInfo()

        txtUserLevel.Text = ""
        chkEntry.Checked = False
        chkSubmit.Checked = False
        chkApprove.Checked = False
        chkAdjust.Checked = False
        chkAuth1.Checked = False
        chkAuth2.Checked = False
        chkImport.Checked = False
        chkExport.Checked = False
        chkMaster.Checked = False
        chkSystem.Checked = False
        chkView.Checked = False
        chkDirectInput.Checked = False

    End Sub

    Private Sub BeginAddUserLevel()
        '// Clear all user info
        ClearInfo()

        '// Set focus for add new user id
        txtUserLevel.Enabled = True
        txtUserLevel.Focus()

        cmdDelete.Enabled = False

        '// Set Operation Code
        myOperationCd = OperationCd.AddNew
    End Sub

    Private Sub loadComboFilter()

        Me.cboUserPermissionFilter.Items.Clear()
        Me.cboUserPermissionFilter.Items.Add(STRING_ALL)

        For Each userPermission In [Enum].GetValues(GetType(BGConstant.enumPermissionCd))
            Me.cboUserPermissionFilter.Items.Add(userPermission)
        Next

        Me.cboUserPermissionFilter.SelectedIndex = 0

    End Sub

    Private Sub clearFilter()

        Me.txtUserLevelFilter.Text = ""
        Me.cboUserPermissionFilter.SelectedIndex = 0

    End Sub

    Private Sub loadGridData()

        myClsBG0611BL.UserLevelFilter = Me.txtUserLevelFilter.Text.Trim
        myClsBG0611BL.UserPermissionsFilter = Me.cboUserPermissionFilter.SelectedItem.ToString

        myClsBG0611BL.GetUserLevelList()
        grvMaster.DataSource = myClsBG0611BL.UserLevelList
    End Sub

#End Region

#Region "Control Event"
    Private Sub frmBG0611_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        myControlLoadingFlg = True

        loadComboFilter()

        '// Load User Level List
        loadGridData()

        '// Begin with Add new user mode
        BeginEditUserLevel()

        '// Reset Flags
        myControlLoadingFlg = False
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub grvMaster_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grvMaster.CurrentCellChanged

        If grvMaster.CurrentCell IsNot Nothing And myControlLoadingFlg = False Then
            BeginEditUserLevel()
        End If

    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        If myOperationCd = OperationCd.AddNew Then
            '// Save new user info
            SaveNewData()
        Else
            '// Save changed user info
            SaveEditData()
        End If
    End Sub

#End Region

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click

        BeginAddUserLevel()


    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        If MessageBox.Show("Are you sure to delete user level?", Me.Text, MessageBoxButtons.YesNo, _
                MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        myClsBG0611BL.UserLevel = CStr(myCurrUserLevel)

        If myClsBG0611BL.DeleteUserLevel() = True Then
            MessageBox.Show("User level was deleted", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            '// Write Transaction Log
            WriteTransactionLog(CStr(enumOperationCd.EditDepartmentMaster), "", "", "", "", "", "")

            '// Load User List
            'myClsBG0611BL.GetUserLevelList()
            'grvMaster.DataSource = myClsBG0611BL.UserLevelList
            loadGridData()

        Else
            MessageBox.Show("There are error between delete user level", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

    End Sub

    Private Sub cmdFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFilter.Click
        loadGridData()
    End Sub

    Private Sub cmdClearFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearFilter.Click
        clearFilter()
    End Sub

End Class