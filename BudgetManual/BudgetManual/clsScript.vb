'   Database backup utility:
'   ========================
'   Copyright (C) 2007  Shabdar Ghata 
'   Email : ghata2002@gmail.com
'   URL : http://www.shabdar.org

'   This program is free software: you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or
'   (at your option) any later version.

'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.

'   You should have received a copy of the GNU General Public License
'   along with this program.  If not, see <http://www.gnu.org/licenses/>.

'   This program comes with ABSOLUTELY NO WARRANTY.

Imports System.Collections.Specialized
Imports System.Data.SqlClient
Imports System.Reflection
Public Class clsScript
    Dim i As Integer
    'Dim oServerSMO As Microsoft.SqlServer.Management.Smo.Server
    '*Dim oServer As New SQLServer2
    Dim objBCPExport As New SQLDMO.BulkCopy2
    Public db As SQLDMO.Database2
    'Public dbSMO As Microsoft.SqlServer.Management.Smo.Database
    'Function Generate_Object_Script() As String
    Dim oServer As New SQLDMO.SQLServer2
    'oServer.Connect("SHABDAR\SQLExpress", "safdar", "ghata")
    ''Dim db As SQLDMO.Database2 = oServer.Databases.Item("SMS")
    '    Return db.Tables.Item(1).Script()
    'End Function
    'This procedure will get all table names from a database
    Public Function ConnectDatabaseWithRefresh(ByVal pServer As String, ByVal pDatabase As String, ByVal pUserName As String, ByVal pPassword As String) As Boolean
        'Try
        'oServerSMO = New Microsoft.SqlServer.Management.Smo.Server
        oServer = New SQLDMO.SQLServer2

        oServer.EnableBcp = True
        oServer.Connect(pServer, pUserName, pPassword)
        'oServerSMO = New Microsoft.SqlServer.Management.Smo.Server(pServer)
        'oServerSMO.UserOptions.
        'oServer.Databases.Refresh()
        db = Nothing
        Try
            db = CType(oServer.Databases.Item(pDatabase), SQLDMO.Database2)
            db.DBOption.SelectIntoBulkCopy = True
            'dbSMO = oServerSMO.Databases(pDatabase)
        Catch
            'If database does not exists then ignore it.
        End Try
        Return True
        'Catch ex As Exception
        '    'MessageBox.Show(ex.ToString)
        '    Return False
        'End Try
    End Function
    Sub ConnectDatabase(ByVal pServer As String, ByVal pDatabase As String, ByVal pUserName As String, ByVal pPassword As String)
        'Try
        '    If oServer.ConnectionID <= 0 Then
        '    End If
        'Catch
        '    '''oServer.EnableBcp = True
        '    '''oServer.Connect(pServer, pUserName, pPassword)
        '    oServerSMO = New Microsoft.SqlServer.Management.Smo.Server(pServer)
        'End Try
        'If IsNothing(dbSMO) Then
        '    '''db = oServer.Databases.Item(pDatabase)
        '    dbSMO = oServerSMO.Databases(pDatabase)
        '    'db.DBOption.SelectIntoBulkCopy = True
        'End If
    End Sub
    Sub CreateNewDatabase(ByVal pServer As String, ByVal pDatabase As String, ByVal pUserName As String, ByVal pPassword As String)
        Dim CN As New SqlConnection("Data Source=" + pServer + ";Persist Security Info=True;User ID=" + pUserName + ";Password=" + pPassword)
        Dim COM As New SqlCommand("create database " + pDatabase, CN)
        CN.Open()
        COM.ExecuteNonQuery()
        CN.Close()
        'dbSMO.ExecuteNonQuery("create database " + pDBName)
        'oServerSMO.Databases.Refresh()
        'dbSMO = oServerSMO.Databases(pDatabase)
        'oServer.ExecuteImmediate("create database " + pDatabase, SQLDMO.SQLDMO_EXEC_TYPE.SQLDMOExec_Default)
        oServer.Databases.Refresh()
        db = CType(oServer.Databases.Item(pDatabase), SQLDMO.Database2)
    End Sub
    Sub DropDatabase(ByVal pServer As String, ByVal pDatabase As String, ByVal pUserName As String, ByVal pPassword As String)
        Try
            db = Nothing
            oServer.DisConnect()
            'oServer = Nothing
        Catch ex As Exception

        End Try
        Dim CN As New SqlConnection("Data Source=" + pServer + ";Persist Security Info=True;User ID=" + pUserName + ";Password=" + pPassword)

        Dim COM As New SqlCommand("DROP DATABASE " + pDatabase, CN)
        CN.Open()
        COM.ExecuteNonQuery()
        CN.Close()
        Try
            oServer.Connect()
        Catch ex As Exception

        End Try

        'dbSMO.ExecuteNonQuery("drop database " + pDBName)
        'dbSMO = Nothing

        'db = Nothing
        'oServer.ExecuteImmediate("drop database " + pDBName, SQLDMO_EXEC_TYPE.SQLDMOExec_Default)
    End Sub
    Sub ExportData(ByVal sServerName As String, ByVal sDatabaseName As String, ByVal sUserName As String, ByVal sPassword As String, ByVal sTableName As String, ByVal sCondition As String, ByVal sTotalRows As String, ByVal sBackUpDir As String)

        objBCPExport.IncludeIdentityValues = True
        objBCPExport.DataFilePath = sBackUpDir + sTableName + ".dat"
        objBCPExport.DataFileType = SQLDMO.SQLDMO_DATAFILE_TYPE.SQLDMODataFile_NativeFormat
        objBCPExport.UseExistingConnection = True

        'objBCP.DataFileType = SQLDMO_DATAFILE_TYPE.SQLDMODataFile_SpecialDelimitedChar
        'objBCP.ColumnDelimiter = vbTab
        'objBCP.RowDelimiter = vbCrLf
        'objBCP.MaximumErrorsBeforeAbort = 1
        'objBCPExport.UseBulkCopyOption = True
        'objBCPExport.ServerBCPKeepIdentity = True
        'objBCP.ServerBCPDataFileType = SQLDMO_SERVERBCP_DATAFILE_TYPE.SQLDMOBCPDataFile_WideNative
        'objBCP.ExportWideChar = True
        'objBCP.TableLock = True

        If InStr(sTotalRows, "TOP 100 PERCENT *", CompareMethod.Text) > 0 And sCondition.Trim.Equals("") Then
            'If db.Tables.Item(sTableName).Rows > 0 Then
            db.Tables.Item(sTableName).ExportData(objBCPExport)
            'End If
        Else
            Dim sSelectCommand As String = "select " + sTotalRows + " from " + sDatabaseName + ".." + sTableName
            If Trim(sCondition) <> "" Then
                sSelectCommand += " WHERE " + sCondition
            End If
            'Create a temporary view to hold data
            Try
                db.ExecuteImmediate("drop view Temp_View_For_Backup_001", SQLDMO.SQLDMO_EXEC_TYPE.SQLDMOExec_Default)
            Catch
            End Try
            db.ExecuteImmediate("create view Temp_View_For_Backup_001 as " + sSelectCommand, SQLDMO.SQLDMO_EXEC_TYPE.SQLDMOExec_Default)
            db.Views.Refresh()
            'db.Views.Item("Temp_View_For_Backup_001").ExportData(objBCP)
            'Export data to a file
            db.Views.Item("Temp_View_For_Backup_001").ExportData(objBCPExport)
            Try
                db.ExecuteImmediate("drop view Temp_View_For_Backup_001", SQLDMO.SQLDMO_EXEC_TYPE.SQLDMOExec_Default)
            Catch
            End Try
        End If
        If My.Computer.FileSystem.GetFileInfo(objBCPExport.DataFilePath).Length <= 0 Then
            'Delete 0 byte size files, to remove error when unzipping
            My.Computer.FileSystem.DeleteFile(objBCPExport.DataFilePath)
            ' Debug.Write("Length=0")
        End If
        'objBCPExport = Nothing
    End Sub
    Function ImportData(ByVal sTableName As String, ByVal sFilePath As String, ByVal bDeleteExistingData As Boolean) As String
        Try
            Dim objBCP As New SQLDMO.BulkCopy2
            objBCP.DataFilePath = sFilePath + sTableName + ".dat"
            If My.Computer.FileSystem.FileExists(objBCP.DataFilePath) Then
                'objBCP.ColumnDelimiter = vbTab
                objBCP.DataFileType = SQLDMO.SQLDMO_DATAFILE_TYPE.SQLDMODataFile_NativeFormat
                'objBCP.RowDelimiter = vbCrLf
                objBCP.UseExistingConnection = True
                objBCP.ServerBCPKeepIdentity = True
                objBCP.IncludeIdentityValues = True
                'objBCP.ExportWideChar = True
                'objBCP.TableLock = True
                'objBCP.UseBulkCopyOption = True
                'objBCP.ServerBCPDataFileType = SQLDMO_SERVERBCP_DATAFILE_TYPE.SQLDMOBCPDataFile_WideNative
                'objBCP.SetCodePage()
                If bDeleteExistingData Then
                    db.ExecuteImmediate("delete from " + sTableName)
                End If
                db.Tables.Refresh()
                db.Tables.Item(sTableName).ImportData(objBCP)
                objBCP = Nothing
                Return "Data imported."
            Else
                'Throw New Exception("File '" + objBCP.DataFilePath + "' does not exists")
                Return "No data found."
            End If
        Catch ex As Exception
            'MessageBox.Show(ex.Message)
            Return (ex.Message)
        End Try
        Return "Pending"
    End Function

    Function GetTableNames(ByVal pServer As String, ByVal pDatabase As String, ByVal pUserName As String, ByVal pPassword As String, ByVal pNoOfRows As String) As DataTable
        ConnectDatabase(pServer, pDatabase, pUserName, pPassword)
        Dim DT As New DataTable
        DT.Columns.Add("TableName", "".GetType)
        DT.Columns.Add("Select", True.GetType)
        DT.Columns.Add("TotalRows", "".GetType)
        DT.Columns.Add("Condition", "".GetType)
        DT.Columns.Add("BCPCommand", "".GetType)
        'DT.Columns.Add("Status", "".GetType)
        'Dim dtTables As DataTable = dbSMO.EnumObjects(DatabaseObjectTypes.Table)
        'dtTables.DefaultView.RowFilter = "schema<>'sys' and schema<>'INFORMATION_SCHEMA'"

        'For i = 0 To dtTables.DefaultView.Count - 1
        '    Dim sBCPCommand As String = "" 'Generate_BCPCommand(pServer, pDatabase, pUserName, pPassword, db.Tables.Item(i).Name, "", pNoOfRows)
        '    DT.Rows.Add(New Object() {dtTables.DefaultView.Item(i)("Name").ToString(), True, pNoOfRows, "", sBCPCommand})
        'Next

        For i As Integer = 1 To db.Tables.Count
            If Not db.Tables.Item(i).SystemObject Then
                Dim sBCPCommand As String = "" 'Generate_BCPCommand(pServer, pDatabase, pUserName, pPassword, db.Tables.Item(i).Name, "", pNoOfRows)
                DT.Rows.Add(New Object() {db.Tables.Item(i).Name, True, pNoOfRows, "", sBCPCommand})
            End If
        Next
        DT.TableName = "Tables"
        Return DT
    End Function
    'Function Generate_BCPImportCommand(ByVal sServerName As String, ByVal sDatabaseName As String, ByVal sUserName As String, ByVal sPassword As String, ByVal sTableName As String, ByVal sCondition As String, ByVal sTotalRows As String,byval )
    'Dim sCommand As String = " bcp ""select " + sTotalRows + " from " + sDatabaseName + ".dbo." + sTableName + """ queryout """ + txtBackupDir.Text + sTableName + ".dat"" -S " + sServerName + " -U " + sUserName + " -P " + sPassword + " " + txtBCPOptions.Text
    'Return sCommand
    'End Function
    Function GetViewNames(ByVal pServer As String, ByVal pDatabase As String, ByVal pUserName As String, ByVal pPassword As String) As DataTable
        ''Dim oServer As New SQLServer2
        ConnectDatabase(pServer, pDatabase, pUserName, pPassword)
        Dim DT As New DataTable
        DT.Columns.Add("ViewName", "".GetType)
        DT.Columns.Add("Select", True.GetType)
        'DT.Columns.Add("Status", "".GetType)
        For i As Integer = 1 To db.Views.Count
            If Not db.Views.Item(i).SystemObject Then
                DT.Rows.Add(New Object() {db.Views.Item(i).Name, True})
            End If
        Next
        'Dim dtTables As DataTable = dbSMO.EnumObjects(DatabaseObjectTypes.View)
        'dtTables.DefaultView.RowFilter = "schema<>'sys' and schema<>'INFORMATION_SCHEMA'"

        'For i = 0 To dtTables.DefaultView.Count - 1
        '    Dim sBCPCommand As String = "" 'Generate_BCPCommand(pServer, pDatabase, pUserName, pPassword, db.Tables.Item(i).Name, "", pNoOfRows)
        '    DT.Rows.Add(New Object() {dtTables.DefaultView.Item(i)("Name").ToString(), True})
        'Next

        DT.TableName = "Views"
        Return DT
    End Function
    Function GetUDFNames(ByVal pServer As String, ByVal pDatabase As String, ByVal pUserName As String, ByVal pPassword As String) As DataTable
        'Dim oServer As New SQLServer2
        ConnectDatabase(pServer, pDatabase, pUserName, pPassword)
        Dim DT As New DataTable
        DT.Columns.Add("UDFName", "".GetType)
        DT.Columns.Add("Select", True.GetType)
        For i As Integer = 1 To db.UserDefinedFunctions.Count
            If Not db.UserDefinedFunctions.Item(i).SystemObject Then
                DT.Rows.Add(New Object() {db.UserDefinedFunctions.Item(i).Name, True})
            End If
        Next
        'Dim dtTables As DataTable = dbSMO.EnumObjects(DatabaseObjectTypes.UserDefinedFunction)
        'dtTables.DefaultView.RowFilter = "schema<>'sys' and schema<>'INFORMATION_SCHEMA'"

        'For i = 0 To dtTables.DefaultView.Count - 1
        '    Dim sBCPCommand As String = "" 'Generate_BCPCommand(pServer, pDatabase, pUserName, pPassword, db.Tables.Item(i).Name, "", pNoOfRows)
        '    DT.Rows.Add(New Object() {dtTables.DefaultView.Item(i)("Name").ToString(), True})
        'Next

        Return DT
    End Function
    Function GetSPNames(ByVal pServer As String, ByVal pDatabase As String, ByVal pUserName As String, ByVal pPassword As String) As DataTable
        'Dim oServer As New SQLServer2
        ConnectDatabase(pServer, pDatabase, pUserName, pPassword)
        Dim DT As New DataTable
        DT.Columns.Add("SPName", "".GetType)
        DT.Columns.Add("Select", True.GetType)
        For i As Integer = 1 To db.StoredProcedures.Count
            If Not db.StoredProcedures.Item(i).SystemObject Then
                DT.Rows.Add(New Object() {db.StoredProcedures.Item(i).Name, True})
            End If
        Next
        'dbSMO.Views.
        'Dim dtTables As DataTable = dbSMO.EnumObjects(DatabaseObjectTypes.StoredProcedure)
        'dtTables.DefaultView.RowFilter = "schema<>'sys' and schema<>'INFORMATION_SCHEMA'"

        'For i = 0 To dtTables.DefaultView.Count - 1
        '    Dim sBCPCommand As String = "" 'Generate_BCPCommand(pServer, pDatabase, pUserName, pPassword, db.Tables.Item(i).Name, "", pNoOfRows)
        '    DT.Rows.Add(New Object() {dtTables.DefaultView.Item(i)("Name").ToString(), True})
        'Next
        Return DT
    End Function
    Function GetUDDNames(ByVal pServer As String, ByVal pDatabase As String, ByVal pUserName As String, ByVal pPassword As String) As DataTable
        'Dim oServer As New SQLServer2
        ConnectDatabase(pServer, pDatabase, pUserName, pPassword)
        Dim DT As New DataTable
        DT.Columns.Add("UDDName", "".GetType)
        DT.Columns.Add("Select", True.GetType)
        For i As Integer = 1 To db.UserDefinedDatatypes.Count
            DT.Rows.Add(New Object() {db.StoredProcedures.Item(i).Name, True})
        Next
        'Dim dtTables As DataTable = dbSMO.EnumObjects(DatabaseObjectTypes.UserDefinedDataType)
        'dtTables.DefaultView.RowFilter = "schema<>'sys' and schema<>'INFORMATION_SCHEMA'"

        'For i = 0 To dtTables.DefaultView.Count - 1
        '    Dim sBCPCommand As String = "" 'Generate_BCPCommand(pServer, pDatabase, pUserName, pPassword, db.Tables.Item(i).Name, "", pNoOfRows)
        '    DT.Rows.Add(New Object() {dtTables.DefaultView.Item(i)("Name").ToString(), True})
        'Next

        Return DT
    End Function
    Function GetUserNames(ByVal pServer As String, ByVal pDatabase As String, ByVal pUserName As String, ByVal pPassword As String) As DataTable
        'Dim oServer As New SQLServer2
        ConnectDatabase(pServer, pDatabase, pUserName, pPassword)
        Dim DT As New DataTable
        DT.Columns.Add("UserName", "".GetType)
        DT.Columns.Add("Select", True.GetType)
        For i As Integer = 1 To db.Users.Count
            If Not db.Users.Item(i).SystemObject Then
                If (db.Users.Item(i).Name.ToUpper() <> "GUEST" And db.Users.Item(i).Name.ToUpper() <> "INFORMATION_SCHEMA" And db.Users.Item(i).Name.ToUpper() <> "SYS") Then
                    DT.Rows.Add(New Object() {db.Users.Item(i).Name, True})
                End If
            End If
        Next

        'Dim dtTables As DataTable = dbSMO.EnumObjects(DatabaseObjectTypes.User)
        'dtTables.DefaultView.RowFilter = "schema<>'sys' and schema<>'INFORMATION_SCHEMA'"

        'For i = 0 To dtTables.DefaultView.Count - 1
        '    If (dtTables.DefaultView.Item(i)("Name").ToString().ToUpper() <> "GUEST" And dtTables.DefaultView.Item(i)("Name").ToString().ToUpper() <> "INFORMATION_SCHEMA" And dtTables.DefaultView.Item(i)("Name").ToString().ToUpper() <> "SYS" And dtTables.DefaultView.Item(i)("Name").ToString().ToUpper() <> "DBO") Then
        '        Dim sBCPCommand As String = "" 'Generate_BCPCommand(pServer, pDatabase, pUserName, pPassword, db.Tables.Item(i).Name, "", pNoOfRows)
        '        DT.Rows.Add(New Object() {dtTables.DefaultView.Item(i)("Name").ToString(), True})
        '    End If
        'Next

        Return DT
    End Function
    Function Generate_BCPBackupCommand(ByVal sServerName As String, ByVal sDatabaseName As String, ByVal sUserName As String, ByVal sPassword As String, ByVal sTableName As String, ByVal sCondition As String, ByVal sTotalRows As String, ByVal sBackUpDir As String, ByVal sBCPPath As String, ByVal sBCPOptions As String) As String
        Dim sCommand As String = "BCP ""select " + sTotalRows + " from " + sDatabaseName + ".dbo." + sTableName + """ queryout """ + sBackUpDir + sTableName + ".dat"" -S " + sServerName + " -U " + sUserName + " -P " + sPassword + " " + sBCPOptions
        Return sCommand
    End Function
    ',byref pObjectDropScript as String
    Function GenerateScript(ByVal pServer As String, ByVal pDatabase As String, ByVal pUserName As String, ByVal pPassword As String, ByVal pObjectType As String, ByVal pObjectName As String) As String
        Dim sSQL As String = ""
        If (pObjectType.ToLower() = "user") Then
            If Not IsNothing(db.Users.Item(pObjectName)) Then
                sSQL = db.Users.Item(pObjectName).Script(SQLDMO.SQLDMO_SCRIPT_TYPE.SQLDMOScript_Drops)
                sSQL = sSQL + db.Users.Item(pObjectName).Script()
            End If
        Else
            If Not IsNothing(db.GetObjectByName(pObjectName)) Then

                sSQL = db.GetObjectByName(pObjectName).Script(SQLDMO.SQLDMO_SCRIPT_TYPE.SQLDMOScript_Drops)
                sSQL = sSQL + db.GetObjectByName(pObjectName).Script()

            End If
        End If
        
        'Select Case pObjectType.ToLower
        '    Case "table"
        '        Dim sp As Table = dbSMO.Tables(pObjectName, dbSMO.UserName)
        '        Dim so As ScriptingOptions = New ScriptingOptions()
        '        so.ScriptDrops = True
        '        so.WithDependencies = True
        '        so.AllowSystemObjects = False
        '        so.ClusteredIndexes = True
        '        so.Default = True
        '        so.DriAll = True
        '        so.Indexes = True
        '        so.NoIdentities = False
        '        so.PrimaryObject = True
        '        so.Triggers = True
        '        so.IncludeIfNotExists = True
        '        Dim sc As StringCollection = sp.Script(so)
        '        Dim s As String
        '        sSQL = ""
        '        For Each s In sc
        '            sSQL += " " + s
        '            'Console.WriteLine(s)
        '        Next
        '        so.IncludeIfNotExists = False
        '        so.ScriptDrops = False
        '        sc = sp.Script(so)
        '        For Each s In sc
        '            sSQL += " " + s
        '            'Console.WriteLine(s)
        '        Next
        '    Case "view"
        '        Dim sp As View = dbSMO.Views(pObjectName, ".")
        '        Dim so As ScriptingOptions = New ScriptingOptions()
        '        so.ScriptDrops = True
        '        so.WithDependencies = True
        '        so.AllowSystemObjects = False
        '        so.ClusteredIndexes = True
        '        so.Default = True
        '        so.DriAll = True
        '        so.Indexes = True
        '        so.NoIdentities = False
        '        so.PrimaryObject = True
        '        so.Triggers = True
        '        so.IncludeIfNotExists = True
        '        Dim sc As StringCollection = sp.Script(so)
        '        Dim s As String
        '        sSQL = ""
        '        For Each s In sc
        '            sSQL += " " + s
        '            'Console.WriteLine(s)
        '        Next
        '        so.IncludeIfNotExists = False
        '        so.ScriptDrops = False
        '        sc = sp.Script(so)
        '        For Each s In sc
        '            sSQL += " " + s
        '            'Console.WriteLine(s)
        '        Next

        '    Case "sp"
        '        Dim sp As StoredProcedure = dbSMO.StoredProcedures(pObjectName, dbSMO.UserName)
        '        Dim so As ScriptingOptions = New ScriptingOptions()
        '        so.ScriptDrops = True
        '        so.WithDependencies = True
        '        so.AllowSystemObjects = False
        '        so.ClusteredIndexes = True
        '        so.Default = True
        '        so.DriAll = True
        '        so.Indexes = True
        '        so.NoIdentities = False
        '        so.PrimaryObject = True
        '        so.Triggers = True
        '        so.IncludeIfNotExists = True
        '        Dim sc As StringCollection = sp.Script(so)
        '        Dim s As String
        '        sSQL = ""
        '        For Each s In sc
        '            sSQL += " " + s
        '            'Console.WriteLine(s)
        '        Next
        '        so.IncludeIfNotExists = False
        '        so.ScriptDrops = False
        '        sc = sp.Script(so)
        '        For Each s In sc
        '            sSQL += " " + s
        '            'Console.WriteLine(s)
        '        Next
        '    Case "udf"
        '        Dim sp As UserDefinedFunction = dbSMO.UserDefinedFunctions(pObjectName, dbSMO.UserName)
        '        Dim so As ScriptingOptions = New ScriptingOptions()
        '        so.ScriptDrops = True
        '        so.WithDependencies = True
        '        so.AllowSystemObjects = False
        '        so.ClusteredIndexes = True
        '        so.Default = True
        '        so.DriAll = True
        '        so.Indexes = True
        '        so.NoIdentities = False
        '        so.PrimaryObject = True
        '        so.Triggers = True
        '        so.IncludeIfNotExists = True
        '        Dim sc As StringCollection = sp.Script(so)
        '        Dim s As String
        '        sSQL = ""
        '        For Each s In sc
        '            sSQL += " " + s
        '            'Console.WriteLine(s)
        '        Next
        '        so.IncludeIfNotExists = False
        '        so.ScriptDrops = False
        '        sc = sp.Script(so)
        '        For Each s In sc
        '            sSQL += " " + s
        '            'Console.WriteLine(s)
        '        Next
        '    Case "udd"
        '        Dim sp As UserDefinedDataType = dbSMO.UserDefinedDataTypes(pObjectName, dbSMO.UserName)
        '        Dim so As ScriptingOptions = New ScriptingOptions()
        '        so.ScriptDrops = True
        '        so.WithDependencies = True
        '        so.AllowSystemObjects = False
        '        so.ClusteredIndexes = True
        '        so.Default = True
        '        so.DriAll = True
        '        so.Indexes = True
        '        so.NoIdentities = False
        '        so.PrimaryObject = True
        '        so.Triggers = True
        '        so.IncludeIfNotExists = True
        '        Dim sc As StringCollection = sp.Script(so)
        '        Dim s As String
        '        sSQL = ""
        '        For Each s In sc
        '            sSQL += " " + s
        '            'Console.WriteLine(s)
        '        Next
        '        so.IncludeIfNotExists = False
        '        so.ScriptDrops = False
        '        sc = sp.Script(so)
        '        For Each s In sc
        '            sSQL += " " + s
        '            'Console.WriteLine(s)
        '        Next
        '    Case "user"
        '        Dim sp As User = dbSMO.Users(pObjectName)
        '        Dim so As ScriptingOptions = New ScriptingOptions()
        '        so.ScriptDrops = True
        '        'so.WithDependencies = True
        '        'so.AllowSystemObjects = False
        '        'so.ClusteredIndexes = True
        '        'so.Default = True
        '        'so.DriAll = True
        '        'so.Indexes = True
        '        'so.NoIdentities = False
        '        'so.PrimaryObject = True
        '        'so.Triggers = True
        '        'so.IncludeIfNotExists = True
        '        Dim sc As StringCollection = sp.Script(so)
        '        Dim s As String
        '        sSQL = ""
        '        For Each s In sc
        '            sSQL += " " + s
        '            'Console.WriteLine(s)
        '        Next
        '        so.IncludeIfNotExists = False
        '        so.ScriptDrops = False
        '        so.WithDependencies = True
        '        so.DriAll = False
        '        so.Permissions = True
        '        so.ExtendedProperties = False
        '        so.LoginSid = True
        '        so.WithDependencies = False
        '        sc = sp.Script(so)
        '        For Each s In sc
        '            sSQL += " " + s
        '            'Console.WriteLine(s)
        '        Next
        'End Select


        Return sSQL
    End Function
    'Function ExecuteScript(ByVal pScript As String) As String
    '    Try
    '        If Trim(pScript) <> "" Then
    '            db.ExecuteImmediate(pScript, SQLDMO_EXEC_TYPE.SQLDMOExec_ContinueOnError)
    '        End If
    '    Catch ex As Exception
    '        Return Find_DependantObject(ex.ToString)
    '    End Try
    '    Return ""
    'End Function
    Function ExecuteScriptWithDependency(ByRef DTObjects As DataTable, ByRef strMessage As String, ByRef progressbar1 As ProgressBar) As Boolean
        Dim sScript As String = ""
        Dim sObjectName As String = ""
        'progressbar1.Minimum = 0
        SetControlPropertyValue(progressbar1, "Minimum", 0)
        SetControlPropertyValue(progressbar1, "Maximum", DTObjects.Rows.Count)
        For i As Integer = 0 To DTObjects.Rows.Count - 1
            sObjectName = CStr(DTObjects.Rows(i)("ObjectName"))
            sScript = CStr(DTObjects.Rows(i)("ScriptSQL"))
            If sScript <> "" Then
                Try
                    strMessage = sObjectName
                    db.ExecuteImmediate(sScript)
                    'dbSMO.ExecuteNonQuery(sScript)
                    DTObjects.Rows(i)("Executed") = True
                    DTObjects.Rows(i)("Status") = "Created."
                Catch ex As Exception
                    DTObjects.Rows(i)("Executed") = False
                    DTObjects.Rows(i)("Status") = ex.Message
                End Try
            End If
            SetControlPropertyValue(progressbar1, "Value", i)
        Next
        Dim repeat As Integer
        For repeat = 0 To 10
            'Execute again with failed objects, this to ensure dependency. Repeat 10 times to include all dependency
            'DTObjects.DefaultView.RowFilter = "Executed=False"
            For i As Integer = 0 To DTObjects.Rows.Count - 1
                If CBool(DTObjects.Rows(i)("Executed")) = False Then
                    sObjectName = CStr(DTObjects.Rows(i)("ObjectName"))
                    sScript = CStr(DTObjects.Rows(i)("ScriptSQL"))
                    If sScript <> "" Then
                        Try
                            'dbSMO.ExecuteNonQuery(sScript)
                            db.ExecuteImmediate(sScript, SQLDMO.SQLDMO_EXEC_TYPE.SQLDMOExec_Default)
                            DTObjects.Rows(i)("Executed") = True
                            DTObjects.Rows(i)("Status") = "Created."
                        Catch ex As Exception
                            DTObjects.Rows(i)("Executed") = False
                            DTObjects.Rows(i)("Status") = ex.Message
                        End Try
                    End If
                End If
            Next
        Next
        Return True
    End Function

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
End Class
