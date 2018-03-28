Imports System.Net
Imports System.Net.Mail

Public Class BGSendMail

#Region "Variable"
    Private Shared mySmtpServer As String
    Private Shared mySmtpPort As Integer
    Private Shared myUseAuthentication As Boolean
    Private Shared mySmtpUser As String
    Private Shared mySmtpPassword As String
    Private Shared myFromAddr As String
#End Region

#Region "Property"
    Public Shared Property SmtpServer() As String
        Get
            Return mySmtpServer
        End Get
        Set(ByVal value As String)
            mySmtpServer = value
        End Set
    End Property

    Public Shared Property SmtpPort() As Integer
        Get
            Return mySmtpPort
        End Get
        Set(ByVal value As Integer)
            mySmtpPort = value
        End Set
    End Property

    Public Shared Property UseAuthentication() As Boolean
        Get
            Return myUseAuthentication
        End Get
        Set(ByVal value As Boolean)
            myUseAuthentication = value
        End Set
    End Property

    Public Shared Property SmtpUser() As String
        Get
            Return mySmtpUser
        End Get
        Set(ByVal value As String)
            mySmtpUser = value
        End Set
    End Property

    Public Shared Property SmtpPassword() As String
        Get
            Return mySmtpPassword
        End Get
        Set(ByVal value As String)
            mySmtpPassword = value
        End Set
    End Property

    Public Shared Property FromAddress() As String
        Get
            Return myFromAddr
        End Get
        Set(ByVal value As String)
            myFromAddr = value
        End Set
    End Property

#End Region

#Region "Function"
    Public Shared Sub SendMessage(ByVal strToAddr As String, ByVal strSubject As String, ByVal strMessage As String)
        Dim objSmtpClient As SmtpClient
        Dim objMail As System.Net.Mail.MailMessage
        Dim objFromAddress As System.Net.Mail.MailAddress

        objSmtpClient = New SmtpClient(SmtpServer, SmtpPort)
        If UseAuthentication Then
            objSmtpClient.Credentials = New NetworkCredential(SmtpUser, SmtpPassword)
        End If
        ''objSmtpClient.DeliveryMethod = SmtpDeliveryMethod.Network

        objMail = New System.Net.Mail.MailMessage
        If FromAddress IsNot Nothing Then
            objFromAddress = New System.Net.Mail.MailAddress(FromAddress)
        Else
            Exit Sub
        End If

        Try
            objMail.From = objFromAddress
            Dim arrTo As String() = Split(strToAddr, ";")
            For Each strTo As String In arrTo
                If strTo.Trim <> "" Then
                    objMail.To.Add(strTo.Trim)
                End If
            Next
            objMail.Subject = strSubject
            objMail.Priority = MailPriority.Normal
            objMail.Body = strMessage
            objMail.IsBodyHtml = False
            objSmtpClient.Send(objMail)
            MessageBox.Show("Sent auto mail to: " & strToAddr, "Auto mail", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Auto mail", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            objSmtpClient = Nothing
            objMail = Nothing
            objFromAddress = Nothing
        End Try
    End Sub
#End Region

End Class

