Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.SqlClient
Imports System.IO

Public Class frmBG0380

#Region "Variable"
    Private myClsBG0380BL As New clsBG0380BL
    Private SharedURL As String = String.Empty
    Private fileName As String = String.Empty
    Private fileNo As String = String.Empty
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
    Private Sub displayGridData()
        If Me.grvMaster.RowCount > 0 And Me.grvMaster.SelectedRows.Count = 1 Then
            Me.txtFileTitle.Text = Me.grvMaster.SelectedRows(0).Cells(1).Value.ToString
            Me.txtFilePath.Text = Me.grvMaster.SelectedRows(0).Cells(2).Value.ToString
            Me.cmdBrowse.Enabled = False
            fileNo = Me.grvMaster.SelectedRows(0).Cells(0).Value.ToString
            fileName = Me.txtFilePath.Text.Replace(SharedURL & "\", "")
            Me.txtFilePath.ReadOnly = True
        End If
    End Sub
#End Region

#Region "Control Event"

    Private Sub frmBG0380_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.txtFileTitle.Text = ""
        Me.txtFilePath.Text = ""
        fileName = ""
        Me.txtFilePath.ReadOnly = False

        If myClsBG0380BL.getFileList Then
            Me.dtInformation = myClsBG0380BL.DTResult
            Me.grvMaster.DataSource = Me.dtInformation
        End If

        If myClsBG0380BL.getSharedUrl Then
            Me.SharedURL = myClsBG0380BL.SharedUrl
        End If

        displayGridData()
    End Sub

    Private Sub grvMaster_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grvMaster.CellClick
        displayGridData()
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Me.cmdBrowse.Enabled = True
        Me.txtFilePath.Text = ""
        Me.txtFileTitle.Text = ""
        fileNo = ""
        fileName = ""
        Me.txtFilePath.ReadOnly = False
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        If Me.txtFileTitle.Text.Trim = "" Then
            showErrorMessage("Please fill in file title")
            Me.txtFileTitle.Focus()
            Return
        ElseIf Me.txtFilePath.Text.Trim = "" Then
            showErrorMessage("Please select file to upload")
            Me.txtFilePath.Focus()
            Return
        ElseIf Not My.Computer.FileSystem.FileExists(Me.txtFilePath.Text.Trim) Then
            showErrorMessage("Specified file was not found.")
            Me.txtFilePath.Focus()
            Me.txtFilePath.Select(0, Me.txtFilePath.Text.Length)
            Return
        End If

        Dim filenames() As String
        filenames = Me.txtFilePath.Text.Split(CChar("\"))
        If filenames.Length > 1 Then
            fileName = filenames(filenames.Length - 1)
        End If

        Me.Enabled = False
        Me.Cursor = Cursors.WaitCursor

        myClsBG0380BL.FileNo = fileNo
        myClsBG0380BL.FileTitle = Me.txtFileTitle.Text

        If Me.SharedURL.Chars(Me.SharedURL.Length - 1).Equals("\") Then
            myClsBG0380BL.FilePath = Me.SharedURL & fileName
        Else
            myClsBG0380BL.FilePath = Me.SharedURL & "\" & fileName
        End If

        myClsBG0380BL.UserId = p_strUserId
        myClsBG0380BL.SourceFile = Me.txtFilePath.Text.Trim
        If myClsBG0380BL.saveData Then
            showSystemMessage("Information file was successfully saved.")

            '// Write Transaction Log
            WriteTransactionLog(CStr(enumOperationCd.EditInformation), "", "", "", "", "", "")

            If myClsBG0380BL.getFileList Then
                Me.dtInformation = myClsBG0380BL.DTResult
                Me.grvMaster.DataSource = Me.dtInformation
            End If

            If myClsBG0380BL.getSharedUrl Then
                Me.SharedURL = myClsBG0380BL.SharedUrl
            End If

            Me.txtFileTitle.Text = ""
            Me.txtFilePath.Text = ""
            fileName = ""

            displayGridData()

            p_frmBG0010.ShowInfoMenu()
        End If
        Me.Enabled = True
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub cmdBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowse.Click
        Dim strFileName As String = String.Empty

        'Show dialog box
        Dim sdlgOpen As OpenFileDialog = New OpenFileDialog
        sdlgOpen.FileName = ""
        sdlgOpen.Filter = "All files (*.*)|*.*"

        Dim dlrConfirm As DialogResult = sdlgOpen.ShowDialog()
        If dlrConfirm.Equals(DialogResult.Cancel) Then
            Exit Sub
        End If

        If sdlgOpen.FileName.Trim.Equals("") Then
            Return
        Else
            Me.txtFilePath.Text = Path.GetFullPath(sdlgOpen.FileName)
            fileName = sdlgOpen.FileName

            Dim filenames() As String
            filenames = Me.txtFilePath.Text.Split(CChar("\"))
            If filenames.Length > 1 Then
                fileName = filenames(filenames.Length - 1)
            End If
        End If

    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        If fileNo = "" Then
            showErrorMessage("Please select file to delete")
            Return
        End If

        If showConfirmMessage("Do you want to delete selected file?") = Windows.Forms.DialogResult.Yes Then
            myClsBG0380BL.FileNo = fileNo
            If myClsBG0380BL.deleteData Then
                showSystemMessage("Information file was successfully deleted.")

                If myClsBG0380BL.getFileList Then
                    Me.dtInformation = myClsBG0380BL.DTResult
                    Me.grvMaster.DataSource = Me.dtInformation
                End If

                If myClsBG0380BL.getSharedUrl Then
                    Me.SharedURL = myClsBG0380BL.SharedUrl
                End If

                Me.txtFileTitle.Text = ""
                Me.txtFilePath.Text = ""
                fileName = ""

                displayGridData()

                p_frmBG0010.ShowInfoMenu()
            End If
        End If
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

#End Region

End Class