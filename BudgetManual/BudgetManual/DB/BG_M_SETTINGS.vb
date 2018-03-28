Imports System.Data.SqlClient
Imports BudgetManual.BGCommon
Imports BudgetManual.BGConstant

Public Class BG_M_SETTINGS

#Region "Variable"
    Private myDtResult As DataTable
    Private myHomeURL As String = String.Empty
    Private mySharedFolder As String = String.Empty
    Private myUserId As String = String.Empty
    Private myHighLightMTP As String = String.Empty
    Private myFromAddr As String = String.Empty
    Private mySmtpServer As String = String.Empty
    Private mySmtpPort As String = String.Empty
    Private myUseAuthentication As Boolean = False
    Private mySmtpUser As String = String.Empty
    Private mySmtpPassword As String = String.Empty
    Private myEnableAutoMail As Boolean = False
    Private myAuthPic1 As Byte()
    Private myAuthPic2 As Byte()
#End Region

#Region "Property"

#Region "dtResult"
    Property dtResult() As DataTable
        Get
            Return myDtResult
        End Get
        Set(ByVal value As DataTable)
            myDtResult = value
        End Set
    End Property
#End Region

#Region "HomeURL"
    Property HomeURL() As String
        Get
            Return myHomeURL
        End Get
        Set(ByVal value As String)
            myHomeURL = value
        End Set
    End Property
#End Region

#Region "SharedFolder"
    Property SharedFolder() As String
        Get
            Return mySharedFolder
        End Get
        Set(ByVal value As String)
            mySharedFolder = value
        End Set
    End Property
#End Region

#Region "UserId"
    Property UserId() As String
        Get
            Return myUserId
        End Get
        Set(ByVal value As String)
            myUserId = value
        End Set
    End Property
#End Region

#Region "HighLightMTP"
    Public Property HighLightMTP() As String
        Get
            Return myHighLightMTP
        End Get
        Set(ByVal value As String)
            myHighLightMTP = value
        End Set
    End Property
#End Region

#Region "FromAddr"
    Public Property FromAddr() As String
        Get
            Return myFromAddr
        End Get
        Set(ByVal value As String)
            myFromAddr = value
        End Set
    End Property
#End Region

#Region "SmtpServer"
    Public Property SmtpServer() As String
        Get
            Return mySmtpServer
        End Get
        Set(ByVal value As String)
            mySmtpServer = value
        End Set
    End Property
#End Region

#Region "SmtpPort"
    Public Property SmtpPort() As String
        Get
            Return mySmtpPort
        End Get
        Set(ByVal value As String)
            mySmtpPort = value
        End Set
    End Property
#End Region

#Region "UseAuthentication "
    Public Property UseAuthentication() As Boolean
        Get
            Return myUseAuthentication
        End Get
        Set(ByVal value As Boolean)
            myUseAuthentication = value
        End Set
    End Property
#End Region

#Region "SmtpUser"
    Public Property SmtpUser() As String
        Get
            Return mySmtpUser
        End Get
        Set(ByVal value As String)
            mySmtpUser = value
        End Set
    End Property
#End Region

#Region "SmtpPassword"
    Public Property SmtpPassword() As String
        Get
            Return mySmtpPassword
        End Get
        Set(ByVal value As String)
            mySmtpPassword = value
        End Set
    End Property
#End Region

#Region "EnableAutoMail "
    Public Property EnableAutoMail() As Boolean
        Get
            Return myEnableAutoMail
        End Get
        Set(ByVal value As Boolean)
            myEnableAutoMail = value
        End Set
    End Property
#End Region

#Region "AuthPic1 "
    Public Property AuthPic1() As Byte()
        Get
            Return myAuthPic1
        End Get
        Set(ByVal value As Byte())
            myAuthPic1 = value
        End Set
    End Property
#End Region

#Region "AuthPic2 "
    Public Property AuthPic2() As Byte()
        Get
            Return myAuthPic2
        End Get
        Set(ByVal value As Byte())
            myAuthPic2 = value
        End Set
    End Property
#End Region

#End Region

#Region "Function"

#Region "Select001"
    ''' <summary>
    ''' Get Options
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select001() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_SETTINGS", "SELECT001")

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable
            da.Fill(dt)

            '// return value
            If dt.Rows.Count > 0 Then
                Me.dtResult = dt
                Me.HomeURL = CStr(Nz(dt.Rows(0).Item("HOME_URL")))
                Me.SharedFolder = CStr(Nz(dt.Rows(0).Item("SHARED_FOLDER")))
                Me.HighLightMTP = CStr(Nz(dt.Rows(0).Item("MTP_HIGHLIGHT_VALUE")))
                Me.FromAddr = CStr(Nz(dt.Rows(0).Item("MAIL_FROM_ADDR")))
                Me.SmtpServer = CStr(Nz(dt.Rows(0).Item("MAIL_SMTP_SVR")))
                Me.SmtpPort = CStr(Nz(dt.Rows(0).Item("MAIL_SMTP_PORT"), 25))
                Me.UseAuthentication = CBool(IIf(CStr(Nz(dt.Rows(0).Item("MAIL_USE_AUTH"), "N")) = "Y", True, False))
                Me.SmtpUser = CStr(Nz(dt.Rows(0).Item("MAIL_USER")))
                Me.SmtpPassword = CStr(Nz(dt.Rows(0).Item("MAIL_PWD")))
                Me.EnableAutoMail = CBool(IIf(CStr(Nz(dt.Rows(0).Item("MAIL_ENABLE"), "N")) = "Y", True, False))
                If Not IsDBNull(dt.Rows(0).Item("AUTH1_IMAGE")) Then
                    Me.AuthPic1 = CType(dt.Rows(0).Item("AUTH1_IMAGE"), Byte())
                Else
                    Me.AuthPic1 = Nothing
                End If
                If Not IsDBNull(dt.Rows(0).Item("AUTH2_IMAGE")) Then
                    Me.AuthPic2 = CType(dt.Rows(0).Item("AUTH2_IMAGE"), Byte())
                Else
                    Me.AuthPic2 = Nothing
                End If
            Else
                Me.dtResult = New DataTable
                Me.HomeURL = ""
                Me.SharedFolder = ""
                Me.HighLightMTP = ""
                Me.FromAddr = ""
                Me.SmtpServer = ""
                Me.SmtpPort = "25"
                Me.UseAuthentication = False
                Me.SmtpUser = ""
                Me.SmtpPassword = ""
                Me.EnableAutoMail = False
                Me.AuthPic1 = Nothing
                Me.AuthPic2 = Nothing
            End If

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_SETTINGS.Select001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Select002"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select002() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim strSQL As String

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_SETTINGS", "SELECT002")

            da = New SqlDataAdapter(strSQL, conn)
            dt = New DataTable("BG_M_SETTINGS")

            da.Fill(dt)

            Me.dtResult = dt

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("[BG_M_SETTINGS.Select002] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#Region "Update001"
    ''' <summary>
    ''' Save Options
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Update001() As Boolean
        Dim conn As SqlConnection = Nothing
        Dim cmd As SqlCommand
        Dim strSQL As String
        Dim intRtn As Integer

        Try
            conn = New SqlConnection(My.Settings.ConnStr)
            conn.Open()

            strSQL = readXMLConfig(p_strDataPath & My.Settings.SqlCmdFile, "BG_M_SETTINGS", "Update001")

            cmd = New SqlCommand(strSQL, conn)
            cmd.Parameters.Add("@HomeUrl", SqlDbType.NVarChar)
            cmd.Parameters("@HomeUrl").Value = Me.HomeURL
            cmd.Parameters.Add("@SharedFolder", SqlDbType.NVarChar)
            cmd.Parameters("@SharedFolder").Value = Me.SharedFolder.Replace("'", "''")
            cmd.Parameters.Add("@HighLightMTP", SqlDbType.Int)
            cmd.Parameters("@HighLightMTP").Value = Me.HighLightMTP
            cmd.Parameters.Add("@FromAddr", SqlDbType.VarChar)
            cmd.Parameters("@FromAddr").Value = Me.FromAddr
            cmd.Parameters.Add("@SmtpServer", SqlDbType.VarChar)
            cmd.Parameters("@SmtpServer").Value = Me.SmtpServer
            cmd.Parameters.Add("@SmtpPort", SqlDbType.TinyInt)
            cmd.Parameters("@SmtpPort").Value = Me.SmtpPort
            cmd.Parameters.Add("@UseAuth", SqlDbType.VarChar)
            If Me.UseAuthentication = True Then
                cmd.Parameters("@UseAuth").Value = "Y"
            Else
                cmd.Parameters("@UseAuth").Value = "N"
            End If
            cmd.Parameters.Add("@SmtpUser", SqlDbType.VarChar)
            cmd.Parameters("@SmtpUser").Value = Me.SmtpUser
            cmd.Parameters.Add("@SmtpPassword", SqlDbType.VarChar)
            cmd.Parameters("@SmtpPassword").Value = Me.SmtpPassword
            cmd.Parameters.Add("@EnableMail", SqlDbType.VarChar)
            If Me.EnableAutoMail = True Then
                cmd.Parameters("@EnableMail").Value = "Y"
            Else
                cmd.Parameters("@EnableMail").Value = "N"
            End If
            cmd.Parameters.Add("@UserId", SqlDbType.VarChar)
            cmd.Parameters("@UserId").Value = Me.UserId

            If Me.AuthPic1 Is Nothing Then
                cmd.Parameters.Add("@Auth1Pic", SqlDbType.Image, 0)
                cmd.Parameters("@Auth1Pic").Value = DBNull.Value
            Else
                cmd.Parameters.Add("@Auth1Pic", SqlDbType.Image, Me.AuthPic1.Length)
                cmd.Parameters("@Auth1Pic").Value = Me.AuthPic1
            End If

            If Me.AuthPic2 Is Nothing Then
                cmd.Parameters.Add("@Auth2Pic", SqlDbType.Image, 0)
                cmd.Parameters("@Auth2Pic").Value = DBNull.Value
            Else
                cmd.Parameters.Add("@Auth2Pic", SqlDbType.Image, Me.AuthPic2.Length)
                cmd.Parameters("@Auth2Pic").Value = Me.AuthPic2
            End If

            intRtn = cmd.ExecuteNonQuery()

            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            If intRtn > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show("[BG_M_SETTINGS.Update001] Error: " & ex.Message, My.Settings.ProgramTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

            If conn IsNot Nothing AndAlso conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return False

        End Try
    End Function
#End Region

#End Region

End Class