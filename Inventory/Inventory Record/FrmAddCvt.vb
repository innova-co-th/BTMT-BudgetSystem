#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Inventory_Record.Common
#End Region

Public Class FrmAddCvt
#Region "Inherits"
    Inherits System.Windows.Forms.Form
    Protected Const TBL_CVT As String = "TBL_CVT"
    Dim GrdDVType As New DataView
    Protected Const TBL_Type As String = "TBL_Type"
    Dim GrdDVRM As New DataView
    Protected Const TBL_RM As String = "TBL_RM"
    Dim GrdDVUnit1 As New DataView
    Protected Const TBL_Unit1 As String = "TBL_Unit1"
    Dim GrdDVUnit2 As New DataView
    Protected Const TBL_Unit2 As String = "TBL_Unit2"
    Dim GrdDVFinal As New DataView
    Protected Const TBL_Final As String = "TBL_Final"

    Protected DefaultGridBorderStyle As BorderStyle
    Dim GrdDV As New DataView
    Public Shared cmG As CurrencyManager
    Dim C1 As New SQLData("ACCINV")
    Public StrType, StrMaterial As String 'It is used in Inventory Tag
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
    Friend WithEvents CmbMaterial As System.Windows.Forms.ComboBox
    Friend WithEvents CmbType As System.Windows.Forms.ComboBox
    Friend WithEvents CheckBoxMaterial As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxType As System.Windows.Forms.CheckBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents CmbUnit2 As System.Windows.Forms.ComboBox
    Friend WithEvents CmbUnit1 As System.Windows.Forms.ComboBox
    Friend WithEvents Txtrev As System.Windows.Forms.TextBox
    Friend WithEvents lblRev As System.Windows.Forms.Label
    Friend WithEvents lblfinal As System.Windows.Forms.Label
    Friend WithEvents CmbFinal As System.Windows.Forms.ComboBox
    Friend WithEvents TxtQty2 As System.Windows.Forms.TextBox
    Friend WithEvents TxtQty1 As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmAddCvt))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Txtrev = New System.Windows.Forms.TextBox
        Me.lblRev = New System.Windows.Forms.Label
        Me.lblfinal = New System.Windows.Forms.Label
        Me.CmbFinal = New System.Windows.Forms.ComboBox
        Me.TxtQty2 = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.CmbUnit2 = New System.Windows.Forms.ComboBox
        Me.CmbUnit1 = New System.Windows.Forms.ComboBox
        Me.CmbMaterial = New System.Windows.Forms.ComboBox
        Me.CmbType = New System.Windows.Forms.ComboBox
        Me.CheckBoxMaterial = New System.Windows.Forms.CheckBox
        Me.CheckBoxType = New System.Windows.Forms.CheckBox
        Me.TxtQty1 = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.CmdSave = New System.Windows.Forms.Button
        Me.CmdClose = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.Txtrev)
        Me.GroupBox1.Controls.Add(Me.lblRev)
        Me.GroupBox1.Controls.Add(Me.lblfinal)
        Me.GroupBox1.Controls.Add(Me.CmbFinal)
        Me.GroupBox1.Controls.Add(Me.TxtQty2)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.CmbUnit2)
        Me.GroupBox1.Controls.Add(Me.CmbUnit1)
        Me.GroupBox1.Controls.Add(Me.CmbMaterial)
        Me.GroupBox1.Controls.Add(Me.CmbType)
        Me.GroupBox1.Controls.Add(Me.CheckBoxMaterial)
        Me.GroupBox1.Controls.Add(Me.CheckBoxType)
        Me.GroupBox1.Controls.Add(Me.TxtQty1)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(474, 154)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'Txtrev
        '
        Me.Txtrev.Location = New System.Drawing.Point(296, 48)
        Me.Txtrev.Name = "Txtrev"
        Me.Txtrev.Size = New System.Drawing.Size(56, 20)
        Me.Txtrev.TabIndex = 28
        Me.Txtrev.Text = ""
        '
        'lblRev
        '
        Me.lblRev.Location = New System.Drawing.Point(248, 50)
        Me.lblRev.Name = "lblRev"
        Me.lblRev.Size = New System.Drawing.Size(48, 16)
        Me.lblRev.TabIndex = 27
        Me.lblRev.Text = "Rev."
        '
        'lblfinal
        '
        Me.lblfinal.Location = New System.Drawing.Point(248, 18)
        Me.lblfinal.Name = "lblfinal"
        Me.lblfinal.Size = New System.Drawing.Size(40, 16)
        Me.lblfinal.TabIndex = 26
        Me.lblfinal.Text = "Final"
        '
        'CmbFinal
        '
        Me.CmbFinal.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmbFinal.Location = New System.Drawing.Point(296, 16)
        Me.CmbFinal.Name = "CmbFinal"
        Me.CmbFinal.Size = New System.Drawing.Size(120, 21)
        Me.CmbFinal.TabIndex = 25
        Me.CmbFinal.Text = "Select"
        '
        'TxtQty2
        '
        Me.TxtQty2.Location = New System.Drawing.Point(296, 112)
        Me.TxtQty2.Name = "TxtQty2"
        Me.TxtQty2.Size = New System.Drawing.Size(56, 20)
        Me.TxtQty2.TabIndex = 24
        Me.TxtQty2.Text = "1"
        Me.TxtQty2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(248, 114)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(48, 16)
        Me.Label5.TabIndex = 23
        Me.Label5.Text = "Qty"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 114)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 16)
        Me.Label4.TabIndex = 22
        Me.Label4.Text = "Unit (Small)"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 82)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 16)
        Me.Label3.TabIndex = 21
        Me.Label3.Text = "Unit (Big)"
        '
        'CmbUnit2
        '
        Me.CmbUnit2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmbUnit2.Enabled = False
        Me.CmbUnit2.Location = New System.Drawing.Point(112, 112)
        Me.CmbUnit2.Name = "CmbUnit2"
        Me.CmbUnit2.Size = New System.Drawing.Size(120, 21)
        Me.CmbUnit2.TabIndex = 20
        Me.CmbUnit2.Text = "Select"
        '
        'CmbUnit1
        '
        Me.CmbUnit1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmbUnit1.Location = New System.Drawing.Point(112, 80)
        Me.CmbUnit1.Name = "CmbUnit1"
        Me.CmbUnit1.Size = New System.Drawing.Size(120, 21)
        Me.CmbUnit1.TabIndex = 19
        Me.CmbUnit1.Text = "Select"
        '
        'CmbMaterial
        '
        Me.CmbMaterial.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmbMaterial.Enabled = False
        Me.CmbMaterial.Location = New System.Drawing.Point(112, 48)
        Me.CmbMaterial.Name = "CmbMaterial"
        Me.CmbMaterial.Size = New System.Drawing.Size(120, 21)
        Me.CmbMaterial.TabIndex = 18
        Me.CmbMaterial.Text = "Select"
        '
        'CmbType
        '
        Me.CmbType.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmbType.Enabled = False
        Me.CmbType.Location = New System.Drawing.Point(112, 16)
        Me.CmbType.Name = "CmbType"
        Me.CmbType.Size = New System.Drawing.Size(120, 21)
        Me.CmbType.TabIndex = 17
        Me.CmbType.Text = "Select"
        '
        'CheckBoxMaterial
        '
        Me.CheckBoxMaterial.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CheckBoxMaterial.Checked = True
        Me.CheckBoxMaterial.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxMaterial.Location = New System.Drawing.Point(16, 50)
        Me.CheckBoxMaterial.Name = "CheckBoxMaterial"
        Me.CheckBoxMaterial.Size = New System.Drawing.Size(96, 16)
        Me.CheckBoxMaterial.TabIndex = 16
        Me.CheckBoxMaterial.Text = "Material"
        '
        'CheckBoxType
        '
        Me.CheckBoxType.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CheckBoxType.Checked = True
        Me.CheckBoxType.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxType.Location = New System.Drawing.Point(16, 18)
        Me.CheckBoxType.Name = "CheckBoxType"
        Me.CheckBoxType.Size = New System.Drawing.Size(96, 16)
        Me.CheckBoxType.TabIndex = 15
        Me.CheckBoxType.Text = "TypeMaterial"
        '
        'TxtQty1
        '
        Me.TxtQty1.Location = New System.Drawing.Point(296, 80)
        Me.TxtQty1.Name = "TxtQty1"
        Me.TxtQty1.Size = New System.Drawing.Size(56, 20)
        Me.TxtQty1.TabIndex = 2
        Me.TxtQty1.Text = "1"
        Me.TxtQty1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(248, 82)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Qty"
        '
        'CmdSave
        '
        Me.CmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdSave.Image = CType(resources.GetObject("CmdSave.Image"), System.Drawing.Image)
        Me.CmdSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdSave.Location = New System.Drawing.Point(328, 162)
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
        Me.CmdClose.Location = New System.Drawing.Point(408, 162)
        Me.CmdClose.Name = "CmdClose"
        Me.CmdClose.Size = New System.Drawing.Size(72, 56)
        Me.CmdClose.TabIndex = 2
        Me.CmdClose.Text = "Close"
        Me.CmdClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'FrmAddCvt
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(490, 224)
        Me.Controls.Add(Me.CmdClose)
        Me.Controls.Add(Me.CmdSave)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmAddCvt"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ADD Unit Convert"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "CONSTANT"
    Friend vPrt As Boolean
    Friend key_code As String
    Dim str1, st, st_desc, st_type As String
    Dim dt As New DataTable
    Dim ws_prt As Boolean
    Dim oldrow, oldcol, Arow As Integer
    Dim TrxNo As String
    Dim dd(), idate As String
    Dim StrSQL, unitbig As String
#End Region

#Region "PrepareStr"
    Private Function PrepareStr(ByVal strValue As String) As String
        ' This function accepts a string and creates a string that can
        ' be used in a SQL statement by adding single quotes around
        ' it and handling empty values.
        If strValue.Trim() = "" Then
            Return "NULL"
        Else
            Return "'" & strValue.Trim() & "'"
        End If
    End Function

#End Region

#Region "COMBOBOX"
    Sub LoadType()
        Dim dtType As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = "SELECT  *  FROM  TBLType "
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtType = New DataTable
            DA.Fill(dtType)
        Catch
        Finally
        End Try
        dtType.TableName = TBL_Type
        GrdDVType = dtType.DefaultView
        '************************************
        CmbType.DisplayMember = "TypeName"
        CmbType.ValueMember = "TypeCode"
        CmbType.DataSource = dtType
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadUnit1()
        Dim dtUnit1 As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = "SELECT  *  FROM  TBLUnit "
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtUnit1 = New DataTable
            DA.Fill(dtUnit1)
        Catch
        Finally
        End Try
        dtUnit1.TableName = TBL_Unit1
        GrdDVUnit1 = dtUnit1.DefaultView
        '************************************
        CmbUnit1.DisplayMember = "UnitName"
        CmbUnit1.ValueMember = "UnitCode"
        CmbUnit1.DataSource = dtUnit1
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadUnit2()
        Dim dtUnit2 As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = "SELECT  *  FROM  TBLUnit "
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtUnit2 = New DataTable
            DA.Fill(dtUnit2)
        Catch
        Finally
        End Try
        dtUnit2.TableName = TBL_Unit2
        GrdDVUnit2 = dtUnit2.DefaultView
        '************************************
        CmbUnit2.DisplayMember = "UnitName"
        CmbUnit2.ValueMember = "UnitCode"
        CmbUnit2.DataSource = dtUnit2
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadMaterial()
        Dim dtRM As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = " select * from TBLGroup"
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtRM = New DataTable
            DA.Fill(dtRM)
        Catch
        Finally
        End Try
        dtRM.TableName = TBL_RM
        GrdDVRM = dtRM.DefaultView
        '************************************
        CmbMaterial.DisplayMember = "Code"
        CmbMaterial.ValueMember = "Code"
        CmbMaterial.DataSource = GrdDVRM
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub Loadfinal()
        Dim dtfinal As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = " select * from ("
        StrSQL += "  SELECT    distinct finalcompound Final,Compcode code,'03' Typecode FROM        "
        StrSQL += " TBLCompound  "
        StrSQL += " union"
        StrSQL += " SELECT    distinct Final ,semicode code,'05' Typecode  FROM TBLSemi "
        StrSQL += " union"
        StrSQL += " SELECT    distinct Final ,psemicode code,'04' Typecode  FROM TBLPreSemi "
        StrSQL += " union"
        StrSQL += " SELECT    distinct Final ,Tirecode code,'06' Typecode  FROM TBLGTHdr"
        StrSQL += " )Final"
        If CheckBoxType.Checked = True And CheckBoxMaterial.Checked = False Then
            StrSQL += " where Typecode = '" & CmbType.SelectedValue & "'"
        ElseIf CheckBoxType.Checked = False And CheckBoxMaterial.Checked = True Then
            StrSQL += " where code = '" & CmbMaterial.Text.Trim & "'"
        ElseIf CheckBoxType.Checked = True And CheckBoxMaterial.Checked = True Then
            StrSQL += " where Typecode = '" & CmbType.SelectedValue & "'"
            StrSQL += " and code = '" & CmbMaterial.Text.Trim & "'"
        Else : CheckBoxType.Checked = False And CheckBoxMaterial.Checked = False
            CmbFinal.Text = "Select"
            Me.Cursor = System.Windows.Forms.Cursors.Default
            Exit Sub
        End If
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtfinal = New DataTable
            DA.Fill(dtfinal)
        Catch
        Finally
        End Try
        dtfinal.TableName = TBL_Final
        GrdDVFinal = dtfinal.DefaultView
        '************************************
        CmbFinal.DisplayMember = "Final"
        CmbFinal.ValueMember = "Final"
        CmbFinal.DataSource = GrdDVFinal
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
#End Region

#Region "Combobox Change"
    Private Sub CheckBoxType_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxType.CheckedChanged
        If CheckBoxType.Checked = True Then
            CmbType.Enabled = True
            Loadfinal()
        Else
            CmbType.Enabled = False
        End If
    End Sub

    Private Sub CheckBoxMaterial_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxMaterial.CheckedChanged
        If CheckBoxMaterial.Checked = True Then
            CmbMaterial.Enabled = True
            Loadfinal()
        Else
            Txtrev.Text = ""
            CmbMaterial.Enabled = False
        End If
    End Sub

    Private Sub CmbType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbType.SelectedIndexChanged
        If CheckBoxType.Checked = True Then
            CmbType.Enabled = True
            Loadfinal()
        Else
            CmbType.Enabled = False
        End If
    End Sub

#End Region

    Private Sub TxtQty(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtQty1.KeyPress, TxtQty2.KeyPress
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
                If Not IsNumeric(e.KeyChar) Then
                    e.Handled = True
                Else

                End If
        End Select
    End Sub

    Private Sub CmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClose.Click
        Me.Close()
    End Sub

#Region "Add Convert"
    Private Function ChkData() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String
        Try
            strSQL = " select count(*) from TBLConvert"
            strSQL &= " where code  = '" & CmbMaterial.Text.Trim & "'"
            strSQL &= " and  Type  = '" & CmbType.SelectedValue & "'"
            strSQL &= " and  Rev  = '" & Txtrev.Text.Trim & "'"
            strSQL &= " and  UnitBig  = '" & CmbUnit1.SelectedValue & "'"
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
    Private Function ChkDataINV() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Try
            If CmbType.SelectedValue = "01" Then
                strSQL = " select count(*) from TBLRM"
                strSQL &= " where RMcode  = '" & CmbMaterial.Text.Trim & "'"
            ElseIf CmbType.SelectedValue = "02" Then
                strSQL = " select count(*) from TBLPigment "
                strSQL &= " where Pigmentcode  = '" & CmbMaterial.Text.Trim & "'"
                strSQL &= " and  Revision  = '" & Txtrev.Text.Trim & "'"
            ElseIf CmbType.SelectedValue = "03" Then
                strSQL = " select count(*) from TBLCompound "
                strSQL &= " where Compcode  = '" & CmbMaterial.Text.Trim & "'"
                strSQL &= " and  Revision  = '" & Txtrev.Text.Trim & "'"
            ElseIf CmbType.SelectedValue = "04" Then
                strSQL = " select count(*) from TBLPresemi "
                strSQL &= " where psemicode  = '" & CmbMaterial.Text.Trim & "'"
                strSQL &= " and  Revision  = '" & Txtrev.Text.Trim & "'"
            ElseIf CmbType.SelectedValue = "05" Then
                strSQL = " select count(*) from TBLsemi "
                strSQL &= " where semicode  = '" & CmbMaterial.Text.Trim & "'"
                strSQL &= " and  Revision  = '" & Txtrev.Text.Trim & "'"
            ElseIf CmbType.SelectedValue = "06" Then
                strSQL = " select count(*) from TBLGTHdr "
                strSQL &= " where Tirecode  = '" & CmbMaterial.Text.Trim & "'"
                strSQL &= " and  Rev = '" & Txtrev.Text.Trim & "'"
            End If

            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar
            If i <> 0 Then
                ChkDataINV = True
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
    Sub Convert()
        If Me.Text = "Save" Then
            Dim cnSQL As SqlConnection
            Dim cmSQL As SqlCommand
            Dim strSQL As String = String.Empty
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
            Try
                strSQL &= " Insert  TBLConvert "
                strSQL &= " values ( " & PrepareStr(CmbType.SelectedValue)
                strSQL &= " , " & PrepareStr(CmbFinal.Text.Trim)
                strSQL &= " , " & PrepareStr(CmbMaterial.Text.Trim)
                strSQL &= " , " & PrepareStr(Txtrev.Text.Trim)
                strSQL &= " , " & PrepareStr(CmbUnit1.SelectedValue)
                strSQL &= " , " & PrepareStr(CmbUnit2.SelectedValue)
                strSQL &= " , " & PrepareStr(TxtQty1.Text.Trim)
                strSQL &= " , " & PrepareStr(TxtQty2.Text.Trim)
                strSQL &= " ) "
                cnSQL = New SqlConnection(C1.Strcon)
                cnSQL.Open()
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
            Me.Cursor = System.Windows.Forms.Cursors.Default()
        ElseIf Me.Text = "Edit" Then
            Dim cnSQL As SqlConnection
            Dim cmSQL As SqlCommand
            Dim strSQL As String = String.Empty
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
            Try
                strSQL &= " Update  TBLConvert "
                strSQL &= " set UnitBig = " & PrepareStr(CmbUnit1.Text.Trim)
                strSQL &= " , Unitsmall = " & PrepareStr(CmbUnit2.Text.Trim)
                strSQL &= " , BQTY = " & PrepareStr(TxtQty1.Text.Trim)
                strSQL &= " , SQTY = " & PrepareStr(TxtQty2.Text.Trim)
                strSQL &= " Where  Type =  " & PrepareStr(CmbType.Text.Trim)
                strSQL &= " and code =  " & PrepareStr(CmbMaterial.Text.Trim)
                strSQL &= " and Unitbig =  " & PrepareStr(unitbig.Trim)
                If CmbType.Text.Trim = "01" Then
                ElseIf CmbType.Text.Trim = "02" Then
                Else
                    strSQL &= " and rev =  " & PrepareStr(Txtrev.Text.Trim)
                End If

                cnSQL = New SqlConnection(C1.Strcon)
                cnSQL.Open()
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
            Me.Cursor = System.Windows.Forms.Cursors.Default()
        Else
        End If

    End Sub
#End Region

    Private Sub FrmAddCvt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Me.Text = "Save" Then
            LoadType()
            LoadUnit1()
            LoadUnit2()
            LoadMaterial()
            Loadfinal()
            CmbType.SelectedValue = StrType.Trim
            CmbMaterial.Text = StrMaterial.Trim
            CmbUnit2.SelectedValue = "KG"
        Else
            unitbig = CmbUnit1.Text.Trim
        End If
    End Sub

    Private Sub CmbMaterial_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbMaterial.SelectedIndexChanged
        If CheckBoxMaterial.Checked = True Then
            CmbMaterial.Enabled = True
            Loadfinal()
        Else
            CmbMaterial.Enabled = False
        End If
    End Sub

    Private Sub CmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSave.Click
        If Me.Text = "Save" Then
            If CmbType.SelectedValue = "01" Then
            ElseIf CmbType.SelectedValue = "07" Then
            ElseIf CmbType.SelectedValue = "08" Then
            ElseIf CmbType.SelectedValue = "09" Then
            ElseIf CmbType.SelectedValue = "03" Then
                If Txtrev.Text.Trim = "" Or CmbFinal.Text.Trim = "" Then
                    CmbFinal.Focus()
                    MsgBox("Please Check Data.")
                    Exit Sub
                End If
            Else
                If Txtrev.Text.Trim = "" Then
                    Txtrev.Focus()
                    MsgBox("Please Check Data.")
                    Exit Sub
                End If
            End If

            If ChkDataINV() Then
            Else
                Dim msg1 As String
                Dim title1 As String
                Dim style1 As MsgBoxStyle
                Dim response1 As MsgBoxResult
                msg1 = "Don't have data. Do want you Save It." ' Define message.
                style1 = MsgBoxStyle.DefaultButton2 Or _
                   MsgBoxStyle.Information Or MsgBoxStyle.YesNo
                title1 = "Conver Unit Material"   ' Define title.
                ' Display message.
                response1 = MsgBox(msg1, style1, title1)
                If response1 = MsgBoxResult.Yes Then ' User chose Yes.
                Else
                    Exit Sub
                End If

            End If

            If ChkData() Then
                MsgBox("It's duplicate Data.")
                Exit Sub
            Else
            End If
        Else
        End If

        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        msg = "Convert New" ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Conver Unit Material"   ' Define title.

        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            Convert()
            MsgBox("Add New Convert Unit Complete.")
            Me.Close()
        Else
            Exit Sub
        End If
    End Sub
End Class
