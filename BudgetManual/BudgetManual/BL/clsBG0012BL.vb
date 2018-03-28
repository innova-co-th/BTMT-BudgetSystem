Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class clsBG0012BL

#Region "Variable"
    Private myHomePageURL As String = String.Empty
    Private mySharedFolder As String = String.Empty
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

#Region "HomePageURL"
    Public Property HomePageURL() As String
        Get
            Return myHomePageURL
        End Get
        Set(ByVal value As String)
            myHomePageURL = value
        End Set
    End Property
#End Region

#Region "SharedFolder"
    Public Property SharedFolder() As String
        Get
            Return mySharedFolder
        End Get
        Set(ByVal value As String)
            mySharedFolder = value
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
    Public Function GetOptions() As Boolean
        Dim clsBG_M_SETTINGS As New BG_M_SETTINGS

        '// Call Function
        If clsBG_M_SETTINGS.Select001() = True Then
            Me.HomePageURL = clsBG_M_SETTINGS.HomeURL
            Me.SharedFolder = clsBG_M_SETTINGS.SharedFolder
            Me.HighLightMTP = clsBG_M_SETTINGS.HighLightMTP
            Me.FromAddr = clsBG_M_SETTINGS.FromAddr
            Me.SmtpServer = clsBG_M_SETTINGS.SmtpServer
            Me.SmtpPort = clsBG_M_SETTINGS.SmtpPort
            Me.UseAuthentication = clsBG_M_SETTINGS.UseAuthentication
            Me.SmtpUser = clsBG_M_SETTINGS.SmtpUser
            Me.SmtpPassword = clsBG_M_SETTINGS.SmtpPassword
            Me.EnableAutoMail = clsBG_M_SETTINGS.EnableAutoMail
            Me.AuthPic1 = clsBG_M_SETTINGS.AuthPic1
            Me.AuthPic2 = clsBG_M_SETTINGS.AuthPic2

            Return True
        Else
            Me.HomePageURL = ""
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

            Return False
        End If
    End Function

    Public Function UpdateOptions() As Boolean
        Dim clsBG_M_SETTINGS As New BG_M_SETTINGS

        '// Set Parameter
        clsBG_M_SETTINGS.HomeURL = Me.HomePageURL
        clsBG_M_SETTINGS.SharedFolder = Me.SharedFolder
        clsBG_M_SETTINGS.HighLightMTP = Me.HighLightMTP
        clsBG_M_SETTINGS.FromAddr = Me.FromAddr
        clsBG_M_SETTINGS.SmtpServer = Me.SmtpServer
        clsBG_M_SETTINGS.SmtpPort = Me.SmtpPort
        clsBG_M_SETTINGS.UseAuthentication = Me.UseAuthentication
        clsBG_M_SETTINGS.SmtpUser = Me.SmtpUser
        clsBG_M_SETTINGS.SmtpPassword = Me.SmtpPassword
        clsBG_M_SETTINGS.EnableAutoMail = Me.EnableAutoMail
        clsBG_M_SETTINGS.AuthPic1 = Me.AuthPic1
        clsBG_M_SETTINGS.AuthPic2 = Me.AuthPic2
        clsBG_M_SETTINGS.UserId = p_strUserId

        '// Call Function
        Return clsBG_M_SETTINGS.Update001()
    End Function
#End Region

End Class
