#Region " Imports "
Imports System.Data
Imports System.Data.SqlClient
Imports Inventory_Record.Common
Imports Microsoft.VisualBasic
#End Region

Public Class FrmAddUnit
    Inherits System.Windows.Forms.Form
    Dim Vbal As Boolean
    Dim C1 As New SQLData("ACCINV")

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
    Friend WithEvents CmdSave As System.Windows.Forms.Button
    Friend WithEvents CmdClose As System.Windows.Forms.Button
    Friend WithEvents TxtName As System.Windows.Forms.TextBox
    Friend WithEvents TxtNo As System.Windows.Forms.TextBox
    Friend WithEvents lblerror As System.Windows.Forms.Label
    Friend WithEvents lblerror1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtCode As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmAddUnit))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtCode = New System.Windows.Forms.TextBox
        Me.lblerror1 = New System.Windows.Forms.Label
        Me.lblerror = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtName = New System.Windows.Forms.TextBox
        Me.TxtNo = New System.Windows.Forms.TextBox
        Me.CmdClose = New System.Windows.Forms.Button
        Me.CmdSave = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.TxtCode)
        Me.GroupBox1.Controls.Add(Me.lblerror1)
        Me.GroupBox1.Controls.Add(Me.lblerror)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.TxtName)
        Me.GroupBox1.Controls.Add(Me.TxtNo)
        Me.GroupBox1.Controls.Add(Me.CmdClose)
        Me.GroupBox1.Controls.Add(Me.CmdSave)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(338, 104)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 74)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 16)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Name2  :"
        '
        'TxtCode
        '
        Me.TxtCode.Location = New System.Drawing.Point(56, 46)
        Me.TxtCode.Name = "TxtCode"
        Me.TxtCode.Size = New System.Drawing.Size(56, 20)
        Me.TxtCode.TabIndex = 1
        Me.TxtCode.Text = ""
        '
        'lblerror1
        '
        Me.lblerror1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblerror1.ForeColor = System.Drawing.Color.Red
        Me.lblerror1.Location = New System.Drawing.Point(208, 72)
        Me.lblerror1.Name = "lblerror1"
        Me.lblerror1.Size = New System.Drawing.Size(16, 16)
        Me.lblerror1.TabIndex = 7
        Me.lblerror1.Text = "*"
        Me.lblerror1.Visible = False
        '
        'lblerror
        '
        Me.lblerror.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblerror.ForeColor = System.Drawing.Color.Red
        Me.lblerror.Location = New System.Drawing.Point(96, 24)
        Me.lblerror.Name = "lblerror"
        Me.lblerror.Size = New System.Drawing.Size(16, 16)
        Me.lblerror.TabIndex = 6
        Me.lblerror.Text = "*"
        Me.lblerror.Visible = False
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 16)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Name1  :"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 16)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Code  :"
        '
        'TxtName
        '
        Me.TxtName.Location = New System.Drawing.Point(56, 72)
        Me.TxtName.Name = "TxtName"
        Me.TxtName.Size = New System.Drawing.Size(144, 20)
        Me.TxtName.TabIndex = 2
        Me.TxtName.Text = ""
        '
        'TxtNo
        '
        Me.TxtNo.Enabled = False
        Me.TxtNo.Location = New System.Drawing.Point(56, 22)
        Me.TxtNo.Name = "TxtNo"
        Me.TxtNo.Size = New System.Drawing.Size(32, 20)
        Me.TxtNo.TabIndex = 0
        Me.TxtNo.Text = ""
        '
        'CmdClose
        '
        Me.CmdClose.Location = New System.Drawing.Point(248, 48)
        Me.CmdClose.Name = "CmdClose"
        Me.CmdClose.TabIndex = 4
        Me.CmdClose.Text = "Close"
        '
        'CmdSave
        '
        Me.CmdSave.Location = New System.Drawing.Point(248, 24)
        Me.CmdSave.Name = "CmdSave"
        Me.CmdSave.TabIndex = 3
        Me.CmdSave.Text = "Save"
        '
        'FrmAddUnit
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(354, 120)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmAddUnit"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Unit"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub FrmAddUnit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Vbal = True
    End Sub

    Private Sub CmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClose.Click
        Me.Close()
    End Sub

    Sub txt()
        If TxtNo.Text = "" Then
            lblerror.Visible = True
            Vbal = False
        Else
            lblerror.Visible = False
        End If

        If TxtName.Text = "" Then
            lblerror1.Visible = True
            Vbal = False
        Else
            lblerror1.Visible = False
        End If
        If TxtCode.Text = "" Then
            lblerror1.Visible = True
            Vbal = False
        Else
            lblerror1.Visible = False
        End If

    End Sub

    Private Sub CmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSave.Click
        txt()
        If Vbal Then
        Else
            Exit Sub
        End If
        If Me.Text = "Unit" Then
            Unit()
        Else
        End If
    End Sub

#Region "Unit"
    Sub Unit()
        If CmdSave.Text = "Save" Then
            Dim cnSQL As SqlConnection
            Dim cmSQL As SqlCommand
            Dim strSQL As String = String.Empty
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
            Try
                strSQL &= " Insert   TblUnit "
                strSQL &= " values ( '" & TxtNo.Text.Trim & "' , "
                strSQL &= " '" & TxtCode.Text.Trim & "' ,"
                strSQL &= " '" & TxtName.Text.Trim & "' )"
                cnSQL = New SqlConnection(C1.Strcon)
                cnSQL.Open()
                cmSQL = New SqlCommand(strSQL, cnSQL)
                cmSQL.ExecuteNonQuery()
                cnSQL.Close()

                cmSQL.Dispose()
                cnSQL.Dispose()
                '--------------------------------------------------------------------------------------
                Me.Close()
            Catch Exp As SqlException
                MsgBox(Exp.Message, MsgBoxStyle.Critical, "SQL Error")
            Catch Exp As Exception
                MsgBox(Exp.Message, MsgBoxStyle.Critical, "General Error")
            End Try
            Me.Cursor = System.Windows.Forms.Cursors.Default()
        ElseIf CmdSave.Text = "Edit" Then
            Dim cnSQL As SqlConnection
            Dim cmSQL As SqlCommand
            Dim strSQL As String = String.Empty
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
            Try
                strSQL &= " Update tblUnit"
                strSQL &= " set UnitName = '" & TxtName.Text.Trim & "'"
                strSQL &= " , ShortUnitName = '" & TxtCode.Text.Trim & "'"
                strSQL &= " where UnitCode = '" & TxtNo.Text.Trim & "'"
                cnSQL = New SqlConnection(C1.Strcon)
                cnSQL.Open()
                cmSQL = New SqlCommand(strSQL, cnSQL)
                cmSQL.ExecuteNonQuery()
                cnSQL.Close()

                cmSQL.Dispose()
                cnSQL.Dispose()
                '--------------------------------------------------------------------------------------
                Me.Close()
            Catch Exp As SqlException
                MsgBox(Exp.Message, MsgBoxStyle.Critical, "SQL Error")
            Catch Exp As Exception
                MsgBox(Exp.Message, MsgBoxStyle.Critical, "General Error")
            End Try
            Me.Cursor = System.Windows.Forms.Cursors.Default()
        Else
        End If
    End Sub

#End Region

    Private Sub TxtNo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtNo.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                SendKeys.Send("{TAB}")
                TxtNo.Text = TxtNo.Text.ToUpper
            Case Else
                If Len(sender.text) >= 2 Then
                    If TxtNo.SelectionLength = Len(TxtNo.Text) Then
                    Else
                        e.Handled = True
                    End If
                End If
        End Select
    End Sub

    Private Sub TxtCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCode.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                SendKeys.Send("{TAB}")
                TxtCode.Text = TxtCode.Text.ToUpper
            Case Else
        End Select
    End Sub

    Private Sub TxtName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtName.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                SendKeys.Send("{TAB}")
                TxtName.Text = TxtName.Text.ToUpper
            Case Else
        End Select
    End Sub
End Class
