#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Login.Common
#End Region

Public Class FrmChangePassword

#Region "Declare"
    Inherits System.Windows.Forms.Form
    Public Shared GrdDV As New DataView
    Protected Const TBL_Data As String = "TBL_Data"
    Protected DefaultGridBorderStyle As BorderStyle
    Public Shared cm As CurrencyManager
#End Region

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents BtmClose As System.Windows.Forms.Button
    Friend WithEvents TxtPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents BtmSave As System.Windows.Forms.Button
    Friend WithEvents TxtUser As System.Windows.Forms.TextBox
    Friend WithEvents TxtNewPassword As System.Windows.Forms.TextBox
    Friend WithEvents TxtConfirmPassword As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmChangePassword))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BtmClose = New System.Windows.Forms.Button
        Me.BtmSave = New System.Windows.Forms.Button
        Me.TxtPassword = New System.Windows.Forms.TextBox
        Me.TxtUser = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtNewPassword = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtConfirmPassword = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.TxtConfirmPassword)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.TxtNewPassword)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.BtmClose)
        Me.GroupBox1.Controls.Add(Me.BtmSave)
        Me.GroupBox1.Controls.Add(Me.TxtPassword)
        Me.GroupBox1.Controls.Add(Me.TxtUser)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(312, 152)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'BtmClose
        '
        Me.BtmClose.Image = CType(resources.GetObject("BtmClose.Image"), System.Drawing.Image)
        Me.BtmClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtmClose.Location = New System.Drawing.Point(232, 72)
        Me.BtmClose.Name = "BtmClose"
        Me.BtmClose.Size = New System.Drawing.Size(72, 56)
        Me.BtmClose.TabIndex = 5
        Me.BtmClose.Text = "Close"
        Me.BtmClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'BtmSave
        '
        Me.BtmSave.Image = CType(resources.GetObject("BtmSave.Image"), System.Drawing.Image)
        Me.BtmSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtmSave.Location = New System.Drawing.Point(232, 16)
        Me.BtmSave.Name = "BtmSave"
        Me.BtmSave.Size = New System.Drawing.Size(72, 56)
        Me.BtmSave.TabIndex = 4
        Me.BtmSave.Text = "Save"
        Me.BtmSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'TxtPassword
        '
        Me.TxtPassword.Location = New System.Drawing.Point(104, 56)
        Me.TxtPassword.Name = "TxtPassword"
        Me.TxtPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.TxtPassword.Size = New System.Drawing.Size(120, 20)
        Me.TxtPassword.TabIndex = 1
        Me.TxtPassword.Text = ""
        '
        'TxtUser
        '
        Me.TxtUser.Location = New System.Drawing.Point(104, 22)
        Me.TxtUser.Name = "TxtUser"
        Me.TxtUser.ReadOnly = True
        Me.TxtUser.Size = New System.Drawing.Size(120, 20)
        Me.TxtUser.TabIndex = 0
        Me.TxtUser.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 58)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 16)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Old Password"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "UserName"
        '
        'TxtNewPassword
        '
        Me.TxtNewPassword.Location = New System.Drawing.Point(104, 88)
        Me.TxtNewPassword.Name = "TxtNewPassword"
        Me.TxtNewPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.TxtNewPassword.Size = New System.Drawing.Size(120, 20)
        Me.TxtNewPassword.TabIndex = 2
        Me.TxtNewPassword.Text = ""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 90)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 16)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "New Password"
        '
        'TxtConfirmPassword
        '
        Me.TxtConfirmPassword.Location = New System.Drawing.Point(104, 120)
        Me.TxtConfirmPassword.Name = "TxtConfirmPassword"
        Me.TxtConfirmPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.TxtConfirmPassword.Size = New System.Drawing.Size(120, 20)
        Me.TxtConfirmPassword.TabIndex = 3
        Me.TxtConfirmPassword.Text = ""
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 122)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(104, 16)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Confirm Password"
        '
        'FrmChangePassword
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(336, 166)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmChangePassword"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Change Password"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "CONSTANT"
    Dim DT As New DataTable
    Dim StrSQL As String
    Dim oldrow As Integer
    Dim C1 As New SQLData()
#End Region

    Private Sub BtmClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtmClose.Click
        Me.Close()
    End Sub

    Private Sub BtmSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtmSave.Click
        If ChkData() Then
            If TxtConfirmPassword.Text = TxtNewPassword.Text Then
                Change()
            Else
                MsgBox("Invalid NewPassword ! Try again", MsgBoxStyle.Critical, "Change Password")
            End If
        Else
            MsgBox("Invalid OldPassword ! Try again", MsgBoxStyle.Critical, "Change Password")
        End If
    End Sub

#Region "CheckDataLogin"
    Private Function ChkData() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Try
            strSQL &= " select count(*) from TblUser "
            strSQL &= " where UserID  = '" & TxtUser.Text.Trim & "'"
            strSQL &= " and PasswordID  = '" & TxtPassword.Text.Trim & "'"
            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar
            If i <> 0 Then
                ChkData = True
            End If
            cmSQL = New SqlCommand(strSQL, cnSQL)
            cmSQL.ExecuteNonQuery()
            cnSQL.Close()

            cmSQL.Dispose()
            cnSQL.Dispose()
            '--------------------------------------------------------------------------------------
        Catch Exp As SqlException
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "SQL Error")
        Catch Exp As Exception
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "General Error")
        End Try
    End Function
    Sub Change()
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Try
            strSQL = String.Empty
            strSQL &= " Update TblUser "
            strSQL &= " Set PasswordID = '" & TxtNewPassword.Text.Trim & "'"
            strSQL &= " where UserID = '" & TxtUser.Text.Trim & "'"
            strSQL &= " and  PasswordID = '" & TxtPassword.Text.Trim & "'"

            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            cmSQL.ExecuteNonQuery()
            cnSQL.Close()

            cmSQL.Dispose()
            cnSQL.Dispose()
            '--------------------------------------------------------------------------------------
            MsgBox("Complete.", MsgBoxStyle.OkOnly, "Change Password")
            Me.Close()
        Catch Exp As SqlException
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "SQL Error")
        Catch Exp As Exception
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "General Error")
        End Try
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Sub
#End Region
End Class
