Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.SqlClient

Public Class clsBG0380BL

#Region "Variable"
    Private clsBG_T_INFORMATION As BG_T_INFORMATION
    Private clsBG_M_SETTINGS As BG_M_SETTINGS
    Private myDtResult As DataTable
    Private mySharedUrl As String = String.Empty
    Private myFileTitle As String = String.Empty
    Private myFilePath As String = String.Empty
    Private myUserId As String = String.Empty
    Private myFileNo As String = String.Empty
    Private mySourceFile As String = String.Empty
#End Region

#Region "Property"
    Public Property DTResult() As DataTable
        Get
            Return myDtResult
        End Get
        Set(ByVal value As DataTable)
            myDtResult = value
        End Set
    End Property
    Public Property SharedUrl() As String
        Get
            Return mySharedUrl
        End Get
        Set(ByVal value As String)
            mySharedUrl = value
        End Set
    End Property
    Public Property FileNo() As String
        Get
            Return myFileNo
        End Get
        Set(ByVal value As String)
            myFileNo = value
        End Set
    End Property
    Public Property FileTitle() As String
        Get
            Return myFileTitle
        End Get
        Set(ByVal value As String)
            myFileTitle = value
        End Set
    End Property
    Public Property FilePath() As String
        Get
            Return myFilePath
        End Get
        Set(ByVal value As String)
            myFilePath = value
        End Set
    End Property
    Public Property UserId() As String
        Get
            Return myUserId
        End Get
        Set(ByVal value As String)
            myUserId = value
        End Set
    End Property
    Public Property SourceFile() As String
        Get
            Return mySourceFile
        End Get
        Set(ByVal value As String)
            mySourceFile = value
        End Set
    End Property
#End Region

#Region "Function"

    Public Function getFileList() As Boolean
        Dim result As Boolean = False
        clsBG_T_INFORMATION = New BG_T_INFORMATION

        If clsBG_T_INFORMATION.Select001 Then
            Me.DTResult = clsBG_T_INFORMATION.DTResult
            result = True
        Else
            result = False
        End If

        Return result
    End Function

    Public Function getSharedUrl() As Boolean
        Dim result As Boolean = False
        clsBG_M_SETTINGS = New BG_M_SETTINGS

        If clsBG_M_SETTINGS.Select001 Then
            Me.SharedUrl = clsBG_M_SETTINGS.SharedFolder
            result = True
        Else
            result = False
        End If

        Return result
    End Function

    Public Function saveData() As Boolean
        Dim success As Boolean = False
        Dim conn As New SqlConnection
        Dim trans As SqlTransaction
        conn.ConnectionString = My.Settings.ConnStr

        Me.FilePath = Me.FilePath.Replace("'", "_")
        clsBG_T_INFORMATION.FileNo = Me.FileNo
        clsBG_T_INFORMATION.FileTitle = Me.FileTitle
        clsBG_T_INFORMATION.FilePath = Me.FilePath
        clsBG_T_INFORMATION.UserId = Me.UserId


        If clsBG_T_INFORMATION.Select002 Then
            Me.DTResult = clsBG_T_INFORMATION.DTResult

            conn.Open()
            trans = conn.BeginTransaction
            Try
                If Me.DTResult.Rows.Count = 0 Then
                    If clsBG_T_INFORMATION.Insert001(conn, trans) Then
                        My.Computer.FileSystem.CopyFile(Me.SourceFile, Me.FilePath, True)
                        Dim fileDetail As IO.FileInfo = My.Computer.FileSystem.GetFileInfo(Me.FilePath)
                        fileDetail.IsReadOnly = True
                        success = True
                    Else
                        success = False
                    End If
                Else
                    If clsBG_T_INFORMATION.Update001(conn, trans) Then
                        success = True
                    Else
                        success = False
                    End If
                End If

                If success Then
                    trans.Commit()
                Else
                    trans.Rollback()
                End If
            Catch ex As Exception

            Finally
                conn.Close()
            End Try
        End If

        Return success
    End Function

    Public Function deleteData() As Boolean
        Dim success As Boolean = False
        Dim conn As New SqlConnection
        Dim trans As SqlTransaction
        conn.ConnectionString = My.Settings.ConnStr

        clsBG_T_INFORMATION.FileNo = Me.FileNo

        If clsBG_T_INFORMATION.Select002 Then
            Me.DTResult = clsBG_T_INFORMATION.DTResult

            conn.Open()
            trans = conn.BeginTransaction
            Try
                If clsBG_T_INFORMATION.Delete001(conn, trans) Then
                    success = True
                Else
                    success = False
                End If

                If success Then
                    trans.Commit()
                Else
                    trans.Rollback()
                End If
            Catch ex As Exception

            Finally
                conn.Close()
            End Try
        End If

        Return success
    End Function

#End Region

End Class
