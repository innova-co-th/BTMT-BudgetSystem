#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Login.Common
#End Region

Public Class FrmAddUser

#Region "Declare"
    Inherits System.Windows.Forms.Form
    Public Shared GrdDV As New DataView
    Protected Const TBL_Data As String = "TBL_Data"
    Protected DefaultGridBorderStyle As BorderStyle
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
    Friend WithEvents TxtUser As System.Windows.Forms.TextBox
    Friend WithEvents CmbLevel As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents BtmSave As System.Windows.Forms.Button
    Friend WithEvents TxtEmpCode As System.Windows.Forms.TextBox
    Friend WithEvents BtmEmp As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmAddUser))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BtmEmp = New System.Windows.Forms.Button
        Me.BtmSave = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtEmpCode = New System.Windows.Forms.TextBox
        Me.BtmClose = New System.Windows.Forms.Button
        Me.TxtPassword = New System.Windows.Forms.TextBox
        Me.TxtUser = New System.Windows.Forms.TextBox
        Me.CmbLevel = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.BtmEmp)
        Me.GroupBox1.Controls.Add(Me.BtmSave)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.TxtEmpCode)
        Me.GroupBox1.Controls.Add(Me.BtmClose)
        Me.GroupBox1.Controls.Add(Me.TxtPassword)
        Me.GroupBox1.Controls.Add(Me.TxtUser)
        Me.GroupBox1.Controls.Add(Me.CmbLevel)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(352, 168)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'BtmEmp
        '
        Me.BtmEmp.Image = CType(resources.GetObject("BtmEmp.Image"), System.Drawing.Image)
        Me.BtmEmp.Location = New System.Drawing.Point(208, 86)
        Me.BtmEmp.Name = "BtmEmp"
        Me.BtmEmp.Size = New System.Drawing.Size(32, 25)
        Me.BtmEmp.TabIndex = 3
        '
        'BtmSave
        '
        Me.BtmSave.Image = CType(resources.GetObject("BtmSave.Image"), System.Drawing.Image)
        Me.BtmSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtmSave.Location = New System.Drawing.Point(256, 16)
        Me.BtmSave.Name = "BtmSave"
        Me.BtmSave.Size = New System.Drawing.Size(72, 56)
        Me.BtmSave.TabIndex = 5
        Me.BtmSave.Text = "Save"
        Me.BtmSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 90)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(64, 16)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "EmpCode"
        '
        'TxtEmpCode
        '
        Me.TxtEmpCode.Location = New System.Drawing.Point(88, 88)
        Me.TxtEmpCode.Name = "TxtEmpCode"
        Me.TxtEmpCode.Size = New System.Drawing.Size(120, 20)
        Me.TxtEmpCode.TabIndex = 2
        Me.TxtEmpCode.Text = ""
        '
        'BtmClose
        '
        Me.BtmClose.Image = CType(resources.GetObject("BtmClose.Image"), System.Drawing.Image)
        Me.BtmClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtmClose.Location = New System.Drawing.Point(256, 72)
        Me.BtmClose.Name = "BtmClose"
        Me.BtmClose.Size = New System.Drawing.Size(72, 56)
        Me.BtmClose.TabIndex = 6
        Me.BtmClose.Text = "Close"
        Me.BtmClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'TxtPassword
        '
        Me.TxtPassword.Location = New System.Drawing.Point(88, 56)
        Me.TxtPassword.Name = "TxtPassword"
        Me.TxtPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.TxtPassword.Size = New System.Drawing.Size(120, 20)
        Me.TxtPassword.TabIndex = 1
        Me.TxtPassword.Text = ""
        '
        'TxtUser
        '
        Me.TxtUser.Location = New System.Drawing.Point(88, 22)
        Me.TxtUser.Name = "TxtUser"
        Me.TxtUser.Size = New System.Drawing.Size(120, 20)
        Me.TxtUser.TabIndex = 0
        Me.TxtUser.Text = ""
        '
        'CmbLevel
        '
        Me.CmbLevel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CmbLevel.Items.AddRange(New Object() {"Viewer", "User", "Administrator"})
        Me.CmbLevel.Location = New System.Drawing.Point(88, 136)
        Me.CmbLevel.Name = "CmbLevel"
        Me.CmbLevel.Size = New System.Drawing.Size(152, 21)
        Me.CmbLevel.TabIndex = 4
        Me.CmbLevel.Text = "Select"
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label3.Location = New System.Drawing.Point(16, 138)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 16)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Level"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 58)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 16)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Password"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "UserName"
        '
        'FrmAddUser
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(368, 182)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmAddUser"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Create New User"
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

    Dim vBal As Boolean

    Function Check() As Boolean
        Check = True
        If TxtUser.Text <> "" Then
            TxtUser.Text = TxtUser.Text.ToUpper
        Else
            Check = False
        End If

        If TxtPassword.Text <> "" Then
        Else
            Check = False
        End If

        If TxtEmpCode.Text <> "" Then
        Else
            Check = False
        End If
        If Len(TxtEmpCode.Text) = 8 Then
        Else
            Check = False
        End If

        If CmbLevel.Text <> "Select" Then
        Else
            Check = False
        End If
    End Function

    Private Sub BtmClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtmClose.Click
        Me.Close()
    End Sub

    Private Sub BtmSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtmSave.Click
        If Check() Then
            CreateUser()
        Else
            MsgBox("Not Complete.", MsgBoxStyle.Exclamation, "Create User")
        End If
    End Sub

    Sub CreateUser()
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Try
            strSQL = " Insert  TblUser "
            strSQL &= " values ( '" & TxtUser.Text.Trim & "',"
            strSQL &= "'" & TxtPassword.Text.Trim & "',"
            strSQL &= "'" & TxtEmpCode.Text.Trim & "',"
            strSQL &= "'" & CmbLevel.Text.Trim & "')"

            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            cmSQL.ExecuteNonQuery()
            cnSQL.Close()

            cmSQL.Dispose()
            cnSQL.Dispose()
            Me.Close()
            '--------------------------------------------------------------------------------------
            MsgBox("Complete.", MsgBoxStyle.OkOnly, "Create User")
        Catch Exp As SqlException
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "SQL Error")
        Catch Exp As Exception
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "General Error")
        End Try
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Sub

    Private Sub BtmEmp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtmEmp.Click
        Dim frmEmp As New FrmEmployee
        frmEmp.ShowDialog()
        TxtEmpCode.Text = frmEmp.Code
        TxtUser.Text = frmEmp.UserName
    End Sub

    Private Sub TxtEmpCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtEmpCode.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case Else
                If Not IsNumeric(e.KeyChar) Then
                    e.Handled = True
                Else
                    If Len(sender.text) = 8 Then
                        If TxtEmpCode.SelectionLength = Len(TxtEmpCode.Text) Then
                        Else
                            e.Handled = True
                        End If
                    End If
                End If
        End Select
    End Sub

    Private Sub TxtPassword_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtPassword.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case Else
        End Select
    End Sub

    Private Sub TxtUser_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtUser.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                TxtUser.Text = TxtUser.Text.ToUpper
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case Else
        End Select
    End Sub

End Class
