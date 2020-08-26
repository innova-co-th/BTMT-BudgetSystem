#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Inventory_Record.Common
#End Region

Public Class FrmEditPigment

#Region "Declare"
    Inherits System.Windows.Forms.Form
    Dim GrdDV As New DataView
    Protected Const TBL_Pigment As String = "TBL_Pigment "

    Protected DefaultGridBorderStyle As BorderStyle
    Dim tNo As Double
    Dim dNo As Double
    Friend qNo As Double
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
    Friend WithEvents CmdSave As System.Windows.Forms.Button
    Friend WithEvents CmdClose As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents CmdDel As System.Windows.Forms.Button
    Friend WithEvents TxtQty As System.Windows.Forms.TextBox
    Friend WithEvents TxtRM As System.Windows.Forms.TextBox
    Friend WithEvents TxtCode As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtRev As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmEditPigment))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtQty = New System.Windows.Forms.TextBox
        Me.TxtRM = New System.Windows.Forms.TextBox
        Me.TxtCode = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.CmdSave = New System.Windows.Forms.Button
        Me.CmdClose = New System.Windows.Forms.Button
        Me.CmdDel = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtRev = New System.Windows.Forms.TextBox
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.TxtRev)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.TxtQty)
        Me.GroupBox1.Controls.Add(Me.TxtRM)
        Me.GroupBox1.Controls.Add(Me.TxtCode)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(306, 122)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(184, 90)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(48, 16)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "KG"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 90)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 16)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Qty"
        '
        'TxtQty
        '
        Me.TxtQty.Location = New System.Drawing.Point(96, 88)
        Me.TxtQty.Name = "TxtQty"
        Me.TxtQty.Size = New System.Drawing.Size(80, 20)
        Me.TxtQty.TabIndex = 4
        Me.TxtQty.Text = ""
        '
        'TxtRM
        '
        Me.TxtRM.Location = New System.Drawing.Point(96, 56)
        Me.TxtRM.Name = "TxtRM"
        Me.TxtRM.Size = New System.Drawing.Size(80, 20)
        Me.TxtRM.TabIndex = 3
        Me.TxtRM.Text = ""
        '
        'TxtCode
        '
        Me.TxtCode.Location = New System.Drawing.Point(96, 24)
        Me.TxtCode.Name = "TxtCode"
        Me.TxtCode.Size = New System.Drawing.Size(80, 20)
        Me.TxtCode.TabIndex = 2
        Me.TxtCode.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 58)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 16)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "R/M Code"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "MasterCode"
        '
        'CmdSave
        '
        Me.CmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdSave.Image = CType(resources.GetObject("CmdSave.Image"), System.Drawing.Image)
        Me.CmdSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdSave.Location = New System.Drawing.Point(160, 130)
        Me.CmdSave.Name = "CmdSave"
        Me.CmdSave.Size = New System.Drawing.Size(80, 56)
        Me.CmdSave.TabIndex = 1
        Me.CmdSave.Text = "Save"
        Me.CmdSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdClose
        '
        Me.CmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdClose.Image = CType(resources.GetObject("CmdClose.Image"), System.Drawing.Image)
        Me.CmdClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdClose.Location = New System.Drawing.Point(240, 130)
        Me.CmdClose.Name = "CmdClose"
        Me.CmdClose.Size = New System.Drawing.Size(72, 56)
        Me.CmdClose.TabIndex = 2
        Me.CmdClose.Text = "Close"
        Me.CmdClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdDel
        '
        Me.CmdDel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdDel.Image = CType(resources.GetObject("CmdDel.Image"), System.Drawing.Image)
        Me.CmdDel.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdDel.Location = New System.Drawing.Point(8, 130)
        Me.CmdDel.Name = "CmdDel"
        Me.CmdDel.Size = New System.Drawing.Size(80, 56)
        Me.CmdDel.TabIndex = 3
        Me.CmdDel.Text = "Delete"
        Me.CmdDel.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(184, 26)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(40, 16)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "Rev."
        '
        'TxtRev
        '
        Me.TxtRev.Location = New System.Drawing.Point(216, 24)
        Me.TxtRev.Name = "TxtRev"
        Me.TxtRev.Size = New System.Drawing.Size(32, 20)
        Me.TxtRev.TabIndex = 8
        Me.TxtRev.Text = ""
        '
        'FrmEditPigment
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(322, 192)
        Me.Controls.Add(Me.CmdDel)
        Me.Controls.Add(Me.CmdClose)
        Me.Controls.Add(Me.CmdSave)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmEditPigment"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "PIGMENT "
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "CONSTANT"
    Dim DT As New DataTable
    Dim StrSQL As String
    Dim oldrow As Integer
    Dim C1 As New SQLData("ACCINV")
#End Region

#Region "Pigment "
    Private Function ChkDataPigment() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Try
            strSQL &= " select count(*) from TBLMASTER  "
            '  strSQL &= " where Pigment code  = '" & TxtNo.Text.Trim & "'"
            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar
            If i <> 0 Then
                ChkDataPigment = True
            End If
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
    End Function
    Private Function iNoPigment() As Double
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim drSQL As SqlDataReader
        Dim strSQL As String
        Try

            ' Build Insert statement to insert 
            strSQL = "  SELECT   Qty "
            strSQL &= "  FROM   TBLPigment"
            strSQL &= "  Where   PigmentCode = '" & TxtCode.Text.Trim & "'"
            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            cmSQL.ExecuteScalar()
            drSQL = cmSQL.ExecuteReader()
            If drSQL.HasRows Then
                If drSQL.Read() Then
                    iNoPigment = CDbl(drSQL.Item("Qty").ToString())
                End If
            End If

            ' Close and Clean up objects
            drSQL.Close()
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
    Sub Pigment(ByVal Qty As Double)
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Try
            strSQL = " update TBLMASTER "
            strSQL &= " set Qty = '" & TxtQty.Text.Trim & "'"
            strSQL &= " where Mastercode = '" & TxtCode.Text.Trim & "'"
            strSQL &= " and  RMCode = '" & TxtRM.Text.Trim & "'"
            strSQL &= " and  Revision = '" & TxtRev.Text.Trim & "'"

            strSQL += ""
            strSQL += " Update TBLConvert "
            strSQL += " set SQty = '" & dNo & "'"
            strSQL += " Where Code =  '" & TxtCode.Text.Trim & "'"
            strSQL += " and Rev = '" & TxtRev.Text.Trim & "'"

            strSQL &= "  "
            strSQL &= " Update TBLPigment "
            strSQL &= " set Qty = '" & Qty & "'"
            strSQL &= " where PigmentCode = '" & TxtCode.Text.Trim & "'"
            strSQL &= " and  Revision = '" & TxtRev.Text.Trim & "'"

            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            cmSQL.ExecuteNonQuery()
            cnSQL.Close()

            cmSQL.Dispose()
            cnSQL.Dispose()
            Me.Close()
        Catch Exp As SqlException
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "SQL Error")
        Catch Exp As Exception
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "General Error")
        End Try
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Sub
#End Region

    Private Sub CmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClose.Click
        Me.Close()
    End Sub

    Private Sub CmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSave.Click
        dNo = tNo + TxtQty.Text.Trim - qNo

        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        msg = "Pigment Total  : " & dNo & "KG"  ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Pigment "   ' Define title.
        If TxtQty.Text.Trim = "" Then
            TxtQty.Focus()
            Exit Sub
        Else
        End If

        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            Pigment(dNo)
        Else
            Exit Sub
        End If
    End Sub


    Private Sub TxtQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtQty.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case 46
                If InStr(sender.text, ".") <> 0 Then
                    e.Handled = True
                End If

            Case Else
                Dim a As Integer = InStr(sender.text, ".")
                If a <> 0 And a = Len(sender.text.Trim) - 3 Then
                    If Len(sender.text.trim) <> sender.SelectionLength Then
                        e.Handled = True
                        Exit Sub
                    End If

                End If

                If Not IsNumeric(e.KeyChar) Then
                    e.Handled = True
                Else
                    If Len(sender.text) >= 8 Then
                        If Len(sender.text) = 8 Then
                            If CDbl(sender.text + e.KeyChar) > 999999 Then
                                e.Handled = True
                            End If
                        Else
                            e.Handled = True
                        End If
                    End If
                End If

        End Select
    End Sub

    Private Sub FrmEditPigment_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        tNo = iNoPigment()
    End Sub

    Private Sub CmdDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdDel.Click
        dNo = tNo - TxtQty.Text.Trim
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        msg = "Pigment Total  : " & dNo & "KG" ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Pigment "   ' Define title.
        If TxtQty.Text.Trim = "" Then
            TxtQty.Focus()
            Exit Sub
        Else
        End If

        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            Dim cnSQL As SqlConnection
            Dim cmSQL As SqlCommand
            Dim strSQL As String
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
            Try
                strSQL = " delete TBLMaster "
                strSQL &= " where MasterCode = '" & TxtCode.Text.Trim & "'"
                strSQL &= " and  RMCode = '" & TxtRM.Text.Trim & "'"
                strSQL &= " and  Revision = '" & TxtRev.Text.Trim & "'"

                strSQL &= "  "
                strSQL &= " Update TBLPigment "
                strSQL &= " set Qty = '" & dNo & "'"
                strSQL &= " where PigmentCode = '" & TxtCode.Text.Trim & "'"
                strSQL &= " and  Revision = '" & TxtRev.Text.Trim & "'"
                strSQL += ""
                strSQL += " Update TBLConvert "
                strSQL += " set SQty = '" & dNo & "'"
                strSQL += " Where Code =  '" & TxtCode.Text.Trim & "'"
                strSQL += " and Rev = '" & TxtRev.Text.Trim & "'"
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
End Class
