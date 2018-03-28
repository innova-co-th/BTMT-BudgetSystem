Imports System.Data
Imports System.Data.OleDb
Imports System.Reflection
Imports System.Threading
Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class frmBG0730

#Region "Variable"

    Const FILE_BUDGET_DATA As String = "BUDGET_DATA_"
    Const FILE_UPLOAD_DATA As String = "UPLOAD_DATA_"
    Const FILE_MASTER_DATA As String = "MASTER_DATA_"

    Dim oScript As New clsScript
    Dim sFileName As String = String.Empty
    Dim strFileDate As String = String.Empty
    Dim strFilePath As String = String.Empty
    Dim strMessage As String = String.Empty
    Dim dtRestore As DataTable

    Dim BackupPath As String = GetCurrentPath() + "BackupTemp\"
    Dim RestorePath As String = GetCurrentPath() + "RestoreTemp\"
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

#Region "Backup"

    Private Sub cmdBackUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBackUp.Click

        If Me.chkBudget.Checked = False AndAlso _
             Me.chkUpload.Checked = False AndAlso _
              Me.chkMaster.Checked = False Then
            MessageBox.Show("Please select data to backup.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If


        If MessageBox.Show("Are you sure to backup?", Me.Text, MessageBoxButtons.YesNo, _
            MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then

            Try
                'If Me.ErrorProviderExtended1.ShowSummaryErrorMessage(Nothing) Then
                oScript.ConnectDatabaseWithRefresh(My.Settings.ServerName, My.Settings.DatabaseName, My.Settings.Username, My.Settings.Password)
                'If chkData.Checked = False And chkStructure.Checked = False Then
                '    MessageBox.Show("Please select structure or data to backup.", "Check", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                '    Exit Sub
                'End If
                ShowProgressBar()
                Panel1.Enabled = False
                ProgressBar1.Maximum = 8
                SetProgress(1)
                Me.Refresh()
                'sFileName = clsCommon.AskSaveAsFile()
                'sFileName = "C:\Documents and Settings\watchaporn\Desktop\BTMT\BTMT_SQLDMO_BG_M_USER.zip"

                strFileDate = Now.ToString("yyyyMMddHHmmss")
                strFilePath = p_strAppPath & "\DBBACKUP\"

                If (Not System.IO.Directory.Exists(strFilePath)) Then
                    System.IO.Directory.CreateDirectory(strFilePath)
                End If

                'If sFileName <> "" Then
                If strFilePath <> "" Then
                    'txtFileName.Text = sFileName
                    Dim t1 As New Thread(AddressOf Export_Database)
                    t1.Start()
                    'Export_Database()
                End If
                'End If
            Catch ex As Exception
                MessageBox.Show(ex.Message(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Finally
                Panel1.Enabled = True
                HideProgressBar()
            End Try

        End If

    End Sub

    'When process is running, disable controls on form
    Sub ShowProgressBar()
        SetControlPropertyValue(ProgressBar1, "Visible", True)

        SetControlPropertyValue(cmdClose, "Enabled", False)
        SetControlPropertyValue(cmdBackUp, "Enabled", False)
        SetControlPropertyValue(cmdOpenFile, "Enabled", False)
        SetControlPropertyValue(cmdRestore, "Enabled", False)

        SetControlPropertyValue(chkBudget, "Enabled", False)
        SetControlPropertyValue(chkMaster, "Enabled", False)
        SetControlPropertyValue(chkUpload, "Enabled", False)

    End Sub


    Sub SetProgress(ByVal iVal As Integer)
        'SetControlPropertyValue(lblPleasewait, "Visible", True)
        'SetControlPropertyValue(ProgressBar1, "Visible", True)
        SetControlPropertyValue(ProgressBar1, "Value", iVal)
        'ProgressBar1.Value = iVal
        'Me.Refresh()
    End Sub

    'Declare delegate for making thread safe calls 
    Delegate Sub SetControlValueCallback(ByVal oControl As Control, ByVal propName As String, ByVal propValue As Object)

    'Method which invokes thread safe call 
    Private Sub SetControlPropertyValue(ByVal oControl As Control, ByVal propName As String, ByVal propValue As Object)
        If oControl.InvokeRequired Then
            Dim d As New SetControlValueCallback(AddressOf SetControlPropertyValue)
            oControl.Invoke(d, New Object() {oControl, propName, propValue})
        Else
            Dim t As Type = oControl.[GetType]()
            Dim props As PropertyInfo() = t.GetProperties()
            For Each p As PropertyInfo In props
                If p.Name.ToUpper() = propName.ToUpper() Then
                    p.SetValue(oControl, propValue, Nothing)
                End If
            Next
        End If
    End Sub

    Sub Export_Database()

        Dim strTableName() As String
        Dim strSelectTable As String = String.Empty

        If Me.chkBudget.Checked = False AndAlso _
            Me.chkUpload.Checked = False AndAlso _
            Me.chkMaster.Checked = False Then

            MessageBox.Show("Cannot backup database.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub

        End If


        ShowProgressBar()


        'If My.Computer.FileSystem.DirectoryExists(BackupPath) Then
        '    My.Computer.FileSystem.DeleteDirectory(BackupPath, FileIO.DeleteDirectoryOption.DeleteAllContents)
        'End If
        'If Not My.Computer.FileSystem.DirectoryExists(BackupPath) Then
        '    My.Computer.FileSystem.CreateDirectory(BackupPath)
        'End If
        Dim DTScripts As New DataTable
        DTScripts.Columns.Add("ObjectName", "".GetType)
        DTScripts.Columns.Add("ObjectType", "".GetType)
        DTScripts.Columns.Add("ScriptSQL", "".GetType)
        'SetProgress(2)
        SetControlPropertyValue(lblObjectName, "Text", "Exporting tables...")

        If Me.chkBudget.Checked = True Then

            SetControlPropertyValue(lblObjectName, "Text", "Exporting Budget tables...")

            DTScripts.Rows.Clear()

            If My.Computer.FileSystem.DirectoryExists(BackupPath) Then
                My.Computer.FileSystem.DeleteDirectory(BackupPath, FileIO.DeleteDirectoryOption.DeleteAllContents)
            End If
            If Not My.Computer.FileSystem.DirectoryExists(BackupPath) Then
                My.Computer.FileSystem.CreateDirectory(BackupPath)
            End If

            strSelectTable = My.Settings.BudgetData
            strTableName = strSelectTable.Split(CChar(","))
            GenerateFiles(DTScripts, strTableName, True, "table", True)

            sFileName = strFilePath & FILE_BUDGET_DATA & strFileDate & ".zip"
            clsZip.ZipIT(BackupPath, sFileName)
            My.Computer.FileSystem.DeleteDirectory(BackupPath, FileIO.DeleteDirectoryOption.DeleteAllContents)

        End If

        If Me.chkUpload.Checked = True Then

            SetControlPropertyValue(lblObjectName, "Text", "Exporting Upload tables...")

            DTScripts.Rows.Clear()

            If My.Computer.FileSystem.DirectoryExists(BackupPath) Then
                My.Computer.FileSystem.DeleteDirectory(BackupPath, FileIO.DeleteDirectoryOption.DeleteAllContents)
            End If
            If Not My.Computer.FileSystem.DirectoryExists(BackupPath) Then
                My.Computer.FileSystem.CreateDirectory(BackupPath)
            End If

            strSelectTable = My.Settings.UploadData
            strTableName = strSelectTable.Split(CChar(","))
            GenerateFiles(DTScripts, strTableName, True, "table", True)

            sFileName = strFilePath & FILE_UPLOAD_DATA & strFileDate & ".zip"
            clsZip.ZipIT(BackupPath, sFileName)
            My.Computer.FileSystem.DeleteDirectory(BackupPath, FileIO.DeleteDirectoryOption.DeleteAllContents)

        End If

        If Me.chkMaster.Checked = True Then

            SetControlPropertyValue(lblObjectName, "Text", "Exporting Master tables...")

            DTScripts.Rows.Clear()

            If My.Computer.FileSystem.DirectoryExists(BackupPath) Then
                My.Computer.FileSystem.DeleteDirectory(BackupPath, FileIO.DeleteDirectoryOption.DeleteAllContents)
            End If
            If Not My.Computer.FileSystem.DirectoryExists(BackupPath) Then
                My.Computer.FileSystem.CreateDirectory(BackupPath)
            End If

            strSelectTable = My.Settings.MasterData
            strTableName = strSelectTable.Split(CChar(","))
            GenerateFiles(DTScripts, strTableName, True, "table", True)

            sFileName = strFilePath & FILE_MASTER_DATA & strFileDate & ".zip"
            clsZip.ZipIT(BackupPath, sFileName)
            My.Computer.FileSystem.DeleteDirectory(BackupPath, FileIO.DeleteDirectoryOption.DeleteAllContents)

        End If

        MessageBox.Show("Database backup complete", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
        HideProgressBar()
    End Sub

    Private Function GetCurrentPath() As String

        Dim sPath As String = System.Reflection.Assembly.GetExecutingAssembly.Location
        Dim n As Integer = InStrRev(sPath, "\")
        If (n > 0) Then
            sPath = Mid(sPath, 1, n)
        End If
        Return sPath

    End Function

    Sub GenerateFiles(ByRef DTScripts As DataTable, ByRef strTableName() As String, ByVal bExportData As Boolean, ByVal pObjectType As String, ByVal bGenerateScripts As Boolean)
        If strTableName IsNot Nothing AndAlso strTableName.Length > 0 Then

            SetControlPropertyValue(ProgressBar1, "Minimum", 0)
            SetControlPropertyValue(ProgressBar1, "Maximum", strTableName.Length)
            SetControlPropertyValue(ProgressBar1, "Value", 0)

            For i As Integer = 0 To strTableName.Length - 1
                'If DT.Rows(i)("Select") = True Then
                ''scriptfile.WriteLine("------------------<" + CStr(DT.Rows(i)(pObjectType + "NAME")) + ">----------------------")
                Dim ObjectName As String = CStr(strTableName(i))
                SetControlPropertyValue(lblStatus, "Text", "" + ObjectName + "")
                If bGenerateScripts Then
                    Dim sSQL As String = oScript.GenerateScript(My.Settings.ServerName, My.Settings.DatabaseName, My.Settings.Username, My.Settings.Password, pObjectType, ObjectName)
                    DTScripts.Rows.Add(New Object() {ObjectName, pObjectType, sSQL})
                End If
                If bExportData Then
                    SetControlPropertyValue(lblStatus, "Text", "Exporting data of '" + ObjectName + "'")
                    oScript.ExportData(My.Settings.ServerName, My.Settings.DatabaseName, My.Settings.Username, My.Settings.Password, ObjectName, "", "TOP 100 PERCENT *", BackupPath)
                End If
                'End If
                SetControlPropertyValue(ProgressBar1, "Value", i)
            Next

            DTScripts.TableName = "SQLScripts"
            DTScripts.WriteXml(BackupPath + "SQLScripts.xml")
        End If
    End Sub

#End Region

#Region "Restore"

    Sub Import_Click()
        If (txtBackupFile.Text = "") Then
            MessageBox.Show("Please select a backup file.", "Select File", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Enable_Form()
            Exit Sub
        End If
        Disable_Form()
        Try
            SetControlPropertyValue(lblStatus, "Text", "Checking database...")
            oScript.ConnectDatabaseWithRefresh(My.Settings.ServerName, My.Settings.DatabaseName, My.Settings.Username, My.Settings.Password)
            If (IsNothing(oScript.db)) Then
                SetControlPropertyValue(lblStatus, "Text", "Creating database...")
                oScript.CreateNewDatabase(My.Settings.ServerName, My.Settings.DatabaseName, My.Settings.Username, My.Settings.Password)
            Else
                'If (Not rdDrop.Checked And Not rdAppend.Checked) Then
                '    MessageBox.Show("This database already exist. Please specify in options if you want to drop existing database or append.", "Database exists", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                '    SetControlPropertyValue(lblStatus, "Text", "")
                '    'lblStatus.Text = ""
                '    Enable_Form()
                '    Exit Sub
                'End If
                'If (rdDrop.Checked) Then
                '    SetControlPropertyValue(lblStatus, "Text", "Dropping database...")
                '    'lblStatus.Text = "Dropping database..."
                '    oScript.DropDatabase(My.Settings.ServerName, My.Settings.DatabaseName, My.Settings.Username, My.Settings.Password)
                '    SetControlPropertyValue(lblStatus, "Text", "Creating database...")
                '    'lblStatus.Text = "Creating database..."
                '    oScript.CreateNewDatabase(My.Settings.ServerName, My.Settings.DatabaseName, My.Settings.Username, My.Settings.Password)
                'End If
            End If
            'If chkDropDB.Checked Then
            '    Dim ans = MessageBox.Show("You have selected to drop any existing database with name '" + txtDatabaseName.Text + "'" + vbNewLine + "Are you sure you want to continue?", "Drop Database", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            '    If ans = vbNo Then
            '        Exit Sub
            '    End If
            '    oScript.DropDatabase(My.Settings.ServerName, My.Settings.DatabaseName, My.Settings.Username, My.Settings.Password)
            'End If
            'If chkCreateDB.Checked Then
            '    oScript.CreateNewDatabase(My.Settings.ServerName, My.Settings.DatabaseName, My.Settings.Username, My.Settings.Password)
            'End If
            'If chkData.Checked = False And chkStructure.Checked = False Then
            '    MessageBox.Show("Please select structure or data to restore.", "Check", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            '    Exit Sub
            'End If
            ShowProgressBar()
            'SetControlPropertyValue(gbExport, "Enabled", False)
            '            gbExport.Enabled = False
            SetControlPropertyValue(ProgressBar1, "Maximum", 8)
            'ProgressBar1.Maximum = 8
            SetProgress(1)
            ' Me.Refresh()
            'Dim t1 As New Thread(AddressOf Import_Database)
            't1.Start()

            Import_Database()
        Catch ex As Exception
            MessageBox.Show(ex.Message(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'SetControlPropertyValue(gbExport, "Enabled", True)
            HideProgressBar()
            Enable_Form()
        End Try
        Enable_Form()
        Clear_Form()
    End Sub

    Sub Import_Database()
        Try
            'Restore_Objects(dgTables)
            'Restore_Objects(dgUDFs)
            'Restore_Objects(dgUDDs)
            'Restore_Objects(dgViews)
            'Restore_Objects(dgSps)
            'Restore_Objects(dgUsers)
            'If chkStructure.Checked Then
            Dim DS As New DataSet
            'lblStatus.Text = "Reading scripts..."
            SetControlPropertyValue(lblStatus, "Text", "Reading scripts...")
            DS.ReadXml(RestorePath + "SQLScripts.xml")
            If DS.Tables.Count > 0 Then
                Dim DT As DataTable = DS.Tables(0)
                Dim i As Integer
                DT.Columns.Add("Executed", False.GetType)
                DT.Columns.Add("Select", False.GetType)
                DT.Columns.Add("Status")
                For i = 0 To DT.Rows.Count - 1
                    DT.Rows(i)("Executed") = False
                    DT.Rows(i)("Select") = False
                    DT.Rows(i)("Status") = ""
                Next
                GetSelectedObjects(DT, dtRestore, "TableName")

                SetControlPropertyValue(lblStatus, "Text", "Creating objects...")
                oScript.ExecuteScriptWithDependency(DT, strMessage, ProgressBar1)
                'Me.txtMessage.Text = strMessage
                GetStatus(DT, dtRestore, "TableName")

            End If
            'End If
            'If chkData.Checked Then
            Restore_Data()
            'End If
            Try
                My.Computer.FileSystem.DeleteDirectory(RestorePath, FileIO.DeleteDirectoryOption.DeleteAllContents)
            Catch
            End Try
            MessageBox.Show("Database restore complete", "Restore Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        End Try
    End Sub

    Sub GetSelectedObjects(ByRef DTObjects As DataTable, ByRef dg As DataTable, ByRef ObjectNameColumn As String)
        Dim i As Integer
        If Not IsNothing(dg) Then
            Dim dt As DataTable = dg
            For i = 0 To DTObjects.Rows.Count - 1
                dt.DefaultView.RowFilter = ObjectNameColumn + "='" + CStr(DTObjects.Rows(i)("ObjectName")) + "'"
                If dt.DefaultView.Count > 0 Then
                    DTObjects.Rows(i)("Select") = dt.DefaultView(0)("Select")
                End If
            Next
            dt.DefaultView.RowFilter = Nothing
        End If
    End Sub

    Sub GetStatus(ByRef DTObjects As DataTable, ByRef dg As DataTable, ByRef ObjectNameColumn As String)
        Dim i As Integer
        If Not IsNothing(dg) Then
            Dim dt As DataTable = dg
            For i = 0 To dt.Rows.Count - 1
                DTObjects.DefaultView.RowFilter = "ObjectName='" + CStr(dt.Rows(i)(ObjectNameColumn)) + "'"
                If DTObjects.DefaultView.Count > 0 Then
                    dt.Rows(i)("status") = DTObjects.DefaultView(0)("status")
                    strMessage = strMessage & vbCrLf & dt.Rows(i)("status").ToString
                    'Me.txtMessage.Text = strMessage
                End If
            Next
            DTObjects.DefaultView.RowFilter = Nothing
        End If
    End Sub

    Sub Restore_Data()
        'lblStatus.Text = "Restoring Data..."
        SetControlPropertyValue(lblStatus, "Text", "Restoring Data...")
        Dim i As Integer
        If Not IsNothing(dtRestore) Then
            Dim DT As DataTable = dtRestore
            SetControlPropertyValue(ProgressBar1, "Value", 0)
            SetControlPropertyValue(ProgressBar1, "Minimum", 0)
            SetControlPropertyValue(ProgressBar1, "Maximum", DT.Rows.Count)
            For i = 0 To DT.Rows.Count - 1

                'If DT.Rows(i)("select") = True Then
                Try
                    SetControlPropertyValue(lblObjectName, "Text", DT.Rows(i)("TableName").ToString())
                    'lblObjectName.Text = DT.Rows(i)("TableName").ToString()
                    'DT.Rows(i)("status") += " " + oScript.ImportData(DT.Rows(i)("TableName"), RestorePath, chkDeleteExistingData.Checked)
                    DT.Rows(i)("status") = CStr(DT.Rows(i)("status")) + " " + oScript.ImportData(CStr(DT.Rows(i)("TableName")), RestorePath, True)
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
                'End If
                SetControlPropertyValue(ProgressBar1, "Value", i)
            Next
        End If
    End Sub

    Private Sub cmdOpenFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOpenFile.Click
        Dim sFile As String = AskOpenFile()
        If sFile <> "" Then
            txtBackupFile.Text = sFile
            'clsZip.UnzipIT(txtBackupFile.Text, txtBackupDir.Text)
            clsZip.UnzipIT(txtBackupFile.Text, RestorePath)
            Try
                ShowProgressBar()
                ProgressBar1.Maximum = 8
                SetProgress(1)
                Me.Refresh()
                Read_Extracted_Dir()
                'Format_All_Grids()
                cmdRestore.Enabled = True
            Catch ex As Exception
                MessageBox.Show(ex.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Finally
                HideProgressBar()
            End Try
        End If
    End Sub

    Private Function AskOpenFile() As String
        Dim fileD As New OpenFileDialog
        fileD.Filter = "Zip Files | *.zip"
        If fileD.ShowDialog() = DialogResult.OK Then
            Return fileD.FileName
        Else
            Return ""
        End If
    End Function

    Sub Read_Extracted_Dir()
        Try
            'oScript.ConnectDatabaseWithRefresh(txtServerName.Text, txtDatabaseName.Text, txtUsername.Text, txtPassword.Text)
            'If chkCreateDB.Checked Then
            '    oScript.DropDatabase(txtServerName.Text, txtDatabaseName.Text, txtUsername.Text, txtPassword.Text)
            '    oScript.CreateNewDatabase(txtServerName.Text, txtDatabaseName.Text, txtUsername.Text, txtPassword.Text)
            'End If

            SetProgress(3)
            Dim DS As New DataSet
            DS.ReadXml(RestorePath + "SQLScripts.xml")
            If DS.Tables.Count > 0 Then
                Dim DT As DataTable = DS.Tables(0)
                DT.DefaultView.RowFilter = "ObjectType='table'"
                Get_Objects(DT, Nothing, "TableName")
                'SetProgress(4)
                'DT.DefaultView.RowFilter = "ObjectType='view'"
                'Get_Objects(DT, dgViews, "ViewName")
                'SetProgress(5)
                'DT.DefaultView.RowFilter = "ObjectType='sp'"
                'Get_Objects(DT, dgSps, "SPName")
                'SetProgress(7)
                'DT.DefaultView.RowFilter = "ObjectType='udf'"
                'Get_Objects(DT, dgUDFs, "UDFName")
                'SetProgress(8)
                'DT.DefaultView.RowFilter = "ObjectType='udd'"
                'Get_Objects(DT, dgUDDs, "UDDName")
                'SetProgress(7)
                'DT.DefaultView.RowFilter = "ObjectType='user'"
                'Get_Objects(DT, dgUsers, "UserName")
                SetProgress(8)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message())
        End Try
    End Sub

    Sub Get_Objects(ByRef DT As DataTable, ByRef dg As DataGridView, ByVal ObjectTypeName As String)
        Dim i As Integer
        Dim DTObjects As New DataTable
        DTObjects.Columns.Add(ObjectTypeName, "".GetType)
        DTObjects.Columns.Add("Select", True.GetType)
        DTObjects.Columns.Add("ScriptSQL", "".GetType)
        DTObjects.Columns.Add("Status", "".GetType)
        For i = 0 To DT.DefaultView.Count - 1
            DTObjects.Rows.Add(New Object() {DT.DefaultView(i)("ObjectName"), True, DT.DefaultView(i)("ScriptSQL")})
        Next
        dtRestore = DTObjects
        'dg.DataSource = DTObjects
        'clsCommon.Format_Grid_Backup(dg)
    End Sub

    Private Sub cmdRestore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRestore.Click
        'If Me.ErrorProviderExtended1.ShowSummaryErrorMessage(Nothing) Then
        Dim t1 As New Thread(AddressOf Import_Click)
        t1.Start()
        'End If
    End Sub

#End Region

#Region "Function/Sub"

    'When process completed, enable controls on form
    Sub HideProgressBar()
        SetControlPropertyValue(ProgressBar1, "Visible", False)
        'SetControlPropertyValue(lblPleasewait, "Visible", False)
        SetControlPropertyValue(lblObjectName, "Text", "")
        SetControlPropertyValue(lblStatus, "Text", "")

        SetControlPropertyValue(cmdClose, "Enabled", True)
        SetControlPropertyValue(cmdBackUp, "Enabled", True)
        SetControlPropertyValue(cmdOpenFile, "Enabled", True)
        SetControlPropertyValue(cmdRestore, "Enabled", True)

        SetControlPropertyValue(chkBudget, "Enabled", True)
        SetControlPropertyValue(chkMaster, "Enabled", True)
        SetControlPropertyValue(chkUpload, "Enabled", True)
    End Sub

    Sub Enable_Form()

        SetControlPropertyValue(cmdClose, "Enabled", True)
        SetControlPropertyValue(cmdBackUp, "Enabled", True)
        SetControlPropertyValue(cmdOpenFile, "Enabled", True)
        SetControlPropertyValue(cmdRestore, "Enabled", True)

        SetControlPropertyValue(chkBudget, "Enabled", True)
        SetControlPropertyValue(chkMaster, "Enabled", True)
        SetControlPropertyValue(chkUpload, "Enabled", True)

    End Sub

    Sub Disable_Form()
        SetControlPropertyValue(cmdClose, "Enabled", False)
        SetControlPropertyValue(cmdBackUp, "Enabled", False)
        SetControlPropertyValue(cmdOpenFile, "Enabled", False)
        SetControlPropertyValue(cmdRestore, "Enabled", False)

        SetControlPropertyValue(chkBudget, "Enabled", False)
        SetControlPropertyValue(chkMaster, "Enabled", False)
        SetControlPropertyValue(chkUpload, "Enabled", False)
    End Sub

    Sub Clear_Form()
        SetControlPropertyValue(lblObjectName, "Text", "")
        SetControlPropertyValue(lblStatus, "Text", "")
        SetControlPropertyValue(ProgressBar1, "Value", 0)
        SetControlPropertyValue(txtBackupFile, "Text", "")
    End Sub

#End Region

#Region "Control Event"

    Private Sub frmBG0730_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        HideProgressBar()
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub
#End Region

End Class