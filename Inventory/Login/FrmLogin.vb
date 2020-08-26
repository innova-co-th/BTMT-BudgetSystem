#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports System.Text
Imports Login.Common
#End Region

Public Structure CellColor
    Public ForeG As Integer
    Public BackG As Integer
    Public LfItem As String
End Structure 'CellColor

Public Class FrmLogin

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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtPassword As System.Windows.Forms.TextBox
    Friend WithEvents BtmLogin As System.Windows.Forms.Button
    Friend WithEvents BtmClose As System.Windows.Forms.Button
    Friend WithEvents CHKPassword As System.Windows.Forms.CheckBox
    Friend WithEvents TxtUser As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmLogin))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.CHKPassword = New System.Windows.Forms.CheckBox
        Me.BtmClose = New System.Windows.Forms.Button
        Me.BtmLogin = New System.Windows.Forms.Button
        Me.TxtPassword = New System.Windows.Forms.TextBox
        Me.TxtUser = New System.Windows.Forms.TextBox
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
        Me.GroupBox1.Controls.Add(Me.CHKPassword)
        Me.GroupBox1.Controls.Add(Me.BtmClose)
        Me.GroupBox1.Controls.Add(Me.BtmLogin)
        Me.GroupBox1.Controls.Add(Me.TxtPassword)
        Me.GroupBox1.Controls.Add(Me.TxtUser)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(320, 144)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'CHKPassword
        '
        Me.CHKPassword.Location = New System.Drawing.Point(88, 88)
        Me.CHKPassword.Name = "CHKPassword"
        Me.CHKPassword.Size = New System.Drawing.Size(120, 24)
        Me.CHKPassword.TabIndex = 4
        Me.CHKPassword.Text = "Change Password"
        '
        'BtmClose
        '
        Me.BtmClose.Image = CType(resources.GetObject("BtmClose.Image"), System.Drawing.Image)
        Me.BtmClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtmClose.Location = New System.Drawing.Point(232, 72)
        Me.BtmClose.Name = "BtmClose"
        Me.BtmClose.Size = New System.Drawing.Size(72, 56)
        Me.BtmClose.TabIndex = 3
        Me.BtmClose.Text = "Close"
        Me.BtmClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'BtmLogin
        '
        Me.BtmLogin.Image = CType(resources.GetObject("BtmLogin.Image"), System.Drawing.Image)
        Me.BtmLogin.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtmLogin.Location = New System.Drawing.Point(232, 16)
        Me.BtmLogin.Name = "BtmLogin"
        Me.BtmLogin.Size = New System.Drawing.Size(72, 56)
        Me.BtmLogin.TabIndex = 2
        Me.BtmLogin.Text = "Login"
        Me.BtmLogin.TextAlign = System.Drawing.ContentAlignment.BottomCenter
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
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 56)
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
        'FrmLogin
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(334, 156)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmLogin"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Log On"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Shared HasColor As Boolean
    Public Shared HasEdited As Boolean
    Public Shared HasAut2Edit As Boolean
    Public Shared HasCommited As Boolean
    Public Shared EmpID As String
    Public Shared EmpName As String
    Public Shared LevelUsage As String

#Region "CONSTANT"
    Dim DT As New DataTable
    Dim oldrow As Integer
    Dim C1 As New SQLData()
#End Region

    Private Sub BtmClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtmClose.Click
        Me.Close()
    End Sub

    Private Sub BtmLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtmLogin.Click
        Dim frmchange As New FrmChangePassword
        If ChkData() Then
            If CHKPassword.Checked Then
                frmchange.TxtUser.Text = TxtUser.Text.Trim
                frmchange.ShowDialog()
                TxtUser.Text = String.Empty
                TxtPassword.Text = String.Empty
                CHKPassword.Checked = False
            Else
                LoadData()
                EmpID = GrdDV.Item(0).Row("EmpCode")
                EmpName = GrdDV.Item(0).Row("EmpName")
                LevelUsage = GrdDV.Item(0).Row("LevelUsage")
                Me.Close()
            End If
        Else
            MsgBox("Invalid UsageID ! Try again", MsgBoxStyle.Critical, "Log On")
            Exit Sub
        End If
    End Sub

#Region "CheckDataLogin"
    Private Function ChkData() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Dim sb As StringBuilder = New StringBuilder()
        Try
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            sb.AppendLine(" select count(*) from TblUser ")
            sb.AppendLine(" where UserID  = '" & TxtUser.Text.Trim & "'")
            sb.AppendLine(" and PasswordID  = '" & TxtPassword.Text.Trim & "'")
            strSQL = sb.ToString()

            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar()
            If i <> 0 Then
                ChkData = True
            End If
            cmSQL = New SqlCommand(strSQL, cnSQL)
            cmSQL.ExecuteNonQuery()
            cnSQL.Close()

            cmSQL.Dispose()
            cnSQL.Dispose()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            '--------------------------------------------------------------------------------------
        Catch Exp As SqlException
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "SQL Error")
        Catch Exp As Exception
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "General Error")
        End Try
    End Function
    Private Sub LoadData()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        Dim strSQL As String = String.Empty
        Dim sb As StringBuilder = New StringBuilder()
        sb.AppendLine(" select u.UserID, u.PasswordId, u.Empcode, u.LevelUsage")
        sb.AppendLine(" ,e.PersonFNameEng+' '+e.PersonLNameEng EmpName")
        sb.AppendLine(" ,e.CitizenID,e.EmployDate,e.DepartDate")
        sb.AppendLine(" from TBLUser u ")
        sb.AppendLine(" left outer join BTMTMaster..TblEmployee e ")
        sb.AppendLine(" on u.empcode = e.empcode ")
        sb.AppendLine(" where UserID  = '" & TxtUser.Text.Trim & "'")
        sb.AppendLine(" and PasswordID  = '" & TxtPassword.Text.Trim & "'")
        strSQL = sb.ToString()
    
        If Not DT Is Nothing Then
            If DT.Rows.Count >= 1 Then
                DT.Clear()
            End If
        End If

        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(strSQL, C1.Strcon)
            Dim CB As New SqlCommandBuilder(DA)
            DT = New DataTable
            DA.Fill(DT)
        Catch
            MsgBox("Can't Select Data,please Check Again.", MsgBoxStyle.Critical, "Load Data")
         Finally
        End Try
        '************************************
        DT.TableName = TBL_Data
        GrdDV = DT.DefaultView
        GrdDV.AllowNew = False
        GrdDV.AllowDelete = False
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    'Sub Login()
    '    If CmdSave.Text = "Save" Then
    '        Dim cnSQL As SqlConnection
    '        Dim cmSQL As SqlCommand
    '        Dim strSQL As String
    '        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
    '        Dim DA As SqlDataAdapter
    '        Try
    '            strSQL &= " Insert  TblUser "
    '            strSQL &= " values ( '" & TxtName.Text.Trim & " ',"
    '            strSQL &= PrepareStr(TxtDesc.Text.Trim) & ","
    '            strSQL &= PrepareStr(TxtSPrice.Text.Trim) & ","
    '            strSQL &= PrepareStr(TxtAPrice.Text.Trim) & ","
    '            strSQL &= PrepareStr("01") & ","
    '            strSQL &= PrepareStr(CmbUnit.SelectedValue) & " )"

    '            strSQL &= ""
    '            strSQL &= " Insert  TblGroup "
    '            strSQL &= " values ( '01',"
    '            strSQL &= PrepareStr(TxtName.Text.Trim) & ")"

    '            If TxtQty.Text <> 0 Then
    '                strSQL &= ""
    '                strSQL &= " Insert  TblQtyUnit "
    '                strSQL &= " values ( '" & TxtName.Text.Trim & " ',"
    '                strSQL &= PrepareStr(TxtRMQty.Text.Trim) & ","
    '                strSQL &= PrepareStr(CmbUnit.SelectedValue) & ","
    '                strSQL &= PrepareStr(TxtQty.Text.Trim) & ","
    '                strSQL &= PrepareStr("KG") & " )"
    '            Else
    '                strSQL &= ""
    '                strSQL &= " Insert  TblQtyUnit "
    '                strSQL &= " values ( '" & TxtName.Text.Trim & " ',"
    '                strSQL &= PrepareStr(CmbUnit.SelectedValue) & ","
    '                strSQL &= PrepareStr(0) & ","
    '                strSQL &= PrepareStr("KG") & " )"
    '            End If

    '            cnSQL = New SqlConnection(C1.Strcon)
    '            cnSQL.Open()
    '            cmSQL = New SqlCommand(strSQL, cnSQL)
    '            cmSQL.ExecuteNonQuery()
    '            cnSQL.Close()

    '            cmSQL.Dispose()
    '            cnSQL.Dispose()
    '            '--------------------------------------------------------------------------------------
    '            Me.Close()
    '        Catch Exp As SqlException
    '            MsgBox(Exp.Message, MsgBoxStyle.Critical, "SQL Error")
    '        Catch Exp As Exception
    '            MsgBox(Exp.Message, MsgBoxStyle.Critical, "General Error")
    '        End Try
    '        Me.Cursor = System.Windows.Forms.Cursors.Default()
    '    ElseIf CmdSave.Text = "Edit" Then
    '        Dim cnSQL As SqlConnection
    '        Dim cmSQL As SqlCommand
    '        Dim strSQL As String
    '        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
    '        Dim DA As SqlDataAdapter
    '        Try
    '            strSQL &= " Update TblUser"
    '            strSQL &= " set descName = '" & TxtDesc.Text.Trim & "'"
    '            strSQL &= " , StdPrice = '" & TxtSPrice.Text.Trim & "'"
    '            strSQL &= " , ActPrice = '" & TxtAPrice.Text.Trim & "'"
    '            strSQL &= " , Unit = '" & CmbUnit.SelectedValue & "'"
    '            strSQL &= " where RMCode = '" & TxtName.Text.Trim & "'"

    '            strSQL &= ""
    '            strSQL &= " Update TblQtyUnit"
    '            strSQL &= " set Qty = '" & TxtQty.Text.Trim & "'"
    '            strSQL &= " , RMQty = '" & TxtRMQty.Text.Trim & "'"
    '            strSQL &= " , UnitCode = '" & CmbUnit.SelectedValue & "'"
    '            strSQL &= " where RMCode = '" & TxtName.Text.Trim & "'"

    '            cnSQL = New SqlConnection(C1.Strcon)
    '            cnSQL.Open()
    '            cmSQL = New SqlCommand(strSQL, cnSQL)
    '            cmSQL.ExecuteNonQuery()
    '            cnSQL.Close()

    '            cmSQL.Dispose()
    '            cnSQL.Dispose()
    '            '--------------------------------------------------------------------------------------
    '            Me.Close()
    '        Catch Exp As SqlException
    '            MsgBox(Exp.Message, MsgBoxStyle.Critical, "SQL Error")
    '        Catch Exp As Exception
    '            MsgBox(Exp.Message, MsgBoxStyle.Critical, "General Error")
    '        End Try
    '        Me.Cursor = System.Windows.Forms.Cursors.Default()
    '    Else
    '    End If

    'End Sub
#End Region

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

    Private Sub TxtPassword_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtPassword.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                TxtPassword.Text = TxtPassword.Text.ToUpper
                e.Handled = True
                SendKeys.Send("{TAB}")
                BtmLogin.PerformClick()
            Case Else
        End Select
    End Sub

End Class
