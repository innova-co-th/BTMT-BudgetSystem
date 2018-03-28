Public Class frmBG0110

#Region "Variable"
    Private myClsBG0110BL As New clsBG0110BL
    Private myHomePageUrl As String
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
    Public Sub ReloadHome()
        If Not Me.IsDisposed Then
            '// Load URL setting
            myHomePageUrl = myClsBG0110BL.GetHomeURL()

            '// Goto Home page
            WebBrowser1.Navigate(myHomePageUrl)
        End If
    End Sub
#End Region

#Region "Control Event"
    Private Sub frmBG0110_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ReloadHome()
    End Sub

    Private Sub tsbHomePage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbHomePage.Click
        ReloadHome()
    End Sub

    Private Sub tsbClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbClose.Click
        Me.Close()
    End Sub

#End Region

End Class