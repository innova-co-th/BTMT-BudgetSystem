#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Inventory_Record.Common
#End Region

Public Class FrmCVT
    Inherits System.Windows.Forms.Form
    Protected Const TBL_CVT As String = "TBL_CVT"
    Dim GrdDVType As New DataView
    Protected Const TBL_Type As String = "TBL_Type"
     Dim GrdDVRM As New DataView
    Protected Const TBL_RM As String = "TBL_RM"

    Protected DefaultGridBorderStyle As BorderStyle
    Dim GrdDV As New DataView
    Public Shared cmG As CurrencyManager
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
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Lname As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents CmdDel As System.Windows.Forms.Button
    Friend WithEvents CmdEdit As System.Windows.Forms.Button
    Friend WithEvents CmdADD As System.Windows.Forms.Button
    Friend WithEvents Lmain As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DataGridCVT As System.Windows.Forms.DataGrid
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBoxType As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxMaterial As System.Windows.Forms.CheckBox
    Friend WithEvents CmbType As System.Windows.Forms.ComboBox
    Friend WithEvents CmbMaterial As System.Windows.Forms.ComboBox
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuView As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuClose As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmCVT))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.DataGridCVT = New System.Windows.Forms.DataGrid
        Me.Lname = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.CmdDel = New System.Windows.Forms.Button
        Me.CmdEdit = New System.Windows.Forms.Button
        Me.CmdADD = New System.Windows.Forms.Button
        Me.Lmain = New System.Windows.Forms.Label
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.CheckBoxType = New System.Windows.Forms.CheckBox
        Me.CheckBoxMaterial = New System.Windows.Forms.CheckBox
        Me.CmbType = New System.Windows.Forms.ComboBox
        Me.CmbMaterial = New System.Windows.Forms.ComboBox
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuClose = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuView = New System.Windows.Forms.MenuItem
        Me.Panel1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DataGridCVT, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.GroupBox2)
        Me.Panel1.Location = New System.Drawing.Point(8, 72)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(856, 328)
        Me.Panel1.TabIndex = 5
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.DataGridCVT)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(856, 328)
        Me.GroupBox2.TabIndex = 15
        Me.GroupBox2.TabStop = False
        '
        'DataGridCVT
        '
        Me.DataGridCVT.DataMember = ""
        Me.DataGridCVT.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridCVT.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridCVT.Location = New System.Drawing.Point(3, 16)
        Me.DataGridCVT.Name = "DataGridCVT"
        Me.DataGridCVT.Size = New System.Drawing.Size(850, 309)
        Me.DataGridCVT.TabIndex = 0
        '
        'Lname
        '
        Me.Lname.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Lname.Location = New System.Drawing.Point(16, 56)
        Me.Lname.Name = "Lname"
        Me.Lname.Size = New System.Drawing.Size(264, 16)
        Me.Lname.TabIndex = 8
        Me.Lname.Text = "Ratio :  Material(Big) / Material (Small)"
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.Controls.Add(Me.CmdDel)
        Me.Panel2.Controls.Add(Me.CmdEdit)
        Me.Panel2.Controls.Add(Me.CmdADD)
        Me.Panel2.Location = New System.Drawing.Point(16, 411)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(368, 88)
        Me.Panel2.TabIndex = 6
        '
        'CmdDel
        '
        Me.CmdDel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.CmdDel.Location = New System.Drawing.Point(16, 56)
        Me.CmdDel.Name = "CmdDel"
        Me.CmdDel.Size = New System.Drawing.Size(136, 23)
        Me.CmdDel.TabIndex = 2
        Me.CmdDel.Text = "Del Material"
        '
        'CmdEdit
        '
        Me.CmdEdit.Enabled = False
        Me.CmdEdit.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.CmdEdit.Location = New System.Drawing.Point(16, 32)
        Me.CmdEdit.Name = "CmdEdit"
        Me.CmdEdit.Size = New System.Drawing.Size(136, 23)
        Me.CmdEdit.TabIndex = 1
        Me.CmdEdit.Text = "Edit Material"
        '
        'CmdADD
        '
        Me.CmdADD.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.CmdADD.Location = New System.Drawing.Point(16, 8)
        Me.CmdADD.Name = "CmdADD"
        Me.CmdADD.Size = New System.Drawing.Size(136, 23)
        Me.CmdADD.TabIndex = 0
        Me.CmdADD.Text = "Add Material"
        '
        'Lmain
        '
        Me.Lmain.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Lmain.Font = New System.Drawing.Font("AngsanaUPC", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Lmain.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Lmain.Location = New System.Drawing.Point(208, 8)
        Me.Lmain.Name = "Lmain"
        Me.Lmain.Size = New System.Drawing.Size(424, 40)
        Me.Lmain.TabIndex = 9
        Me.Lmain.Text = "Convert  Unit  of  Material Stock"
        Me.Lmain.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.Controls.Add(Me.GroupBox1)
        Me.Panel3.Location = New System.Drawing.Point(8, 403)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(856, 104)
        Me.Panel3.TabIndex = 7
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(536, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(312, 88)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label4.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label4.Location = New System.Drawing.Point(16, 48)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(264, 16)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "2.  Unit Can't Duplicate"
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label2.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label2.Location = New System.Drawing.Point(16, 28)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(264, 16)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "1.  Material No. it's  identical"
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label1.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label1.Location = New System.Drawing.Point(16, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(264, 16)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Law of Material"
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(552, 8)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(32, 32)
        Me.PictureBox1.TabIndex = 10
        Me.PictureBox1.TabStop = False
        '
        'CheckBoxType
        '
        Me.CheckBoxType.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CheckBoxType.Location = New System.Drawing.Point(648, 10)
        Me.CheckBoxType.Name = "CheckBoxType"
        Me.CheckBoxType.Size = New System.Drawing.Size(96, 16)
        Me.CheckBoxType.TabIndex = 11
        Me.CheckBoxType.Text = "TypeMaterial"
        '
        'CheckBoxMaterial
        '
        Me.CheckBoxMaterial.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CheckBoxMaterial.Location = New System.Drawing.Point(648, 42)
        Me.CheckBoxMaterial.Name = "CheckBoxMaterial"
        Me.CheckBoxMaterial.Size = New System.Drawing.Size(96, 16)
        Me.CheckBoxMaterial.TabIndex = 12
        Me.CheckBoxMaterial.Text = "Material"
        '
        'CmbType
        '
        Me.CmbType.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmbType.Enabled = False
        Me.CmbType.Location = New System.Drawing.Point(744, 8)
        Me.CmbType.Name = "CmbType"
        Me.CmbType.Size = New System.Drawing.Size(120, 21)
        Me.CmbType.TabIndex = 13
        Me.CmbType.Text = "Select"
        '
        'CmbMaterial
        '
        Me.CmbMaterial.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmbMaterial.Enabled = False
        Me.CmbMaterial.Location = New System.Drawing.Point(744, 40)
        Me.CmbMaterial.Name = "CmbMaterial"
        Me.CmbMaterial.Size = New System.Drawing.Size(120, 21)
        Me.CmbMaterial.TabIndex = 14
        Me.CmbMaterial.Text = "Select"
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem3, Me.MenuItem1})
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 0
        Me.MenuItem3.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuClose})
        Me.MenuItem3.Text = "File"
        '
        'MenuClose
        '
        Me.MenuClose.Index = 0
        Me.MenuClose.Text = "Close"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 1
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuView})
        Me.MenuItem1.Text = "View"
        '
        'MenuView
        '
        Me.MenuView.Index = 0
        Me.MenuView.Shortcut = System.Windows.Forms.Shortcut.F5
        Me.MenuView.Text = "Refresh"
        '
        'FrmCVT
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(872, 516)
        Me.Controls.Add(Me.CmbMaterial)
        Me.Controls.Add(Me.CmbType)
        Me.Controls.Add(Me.CheckBoxMaterial)
        Me.Controls.Add(Me.CheckBoxType)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Lname)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Lmain)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmCVT"
        Me.Text = "Convert Material Unit"
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.DataGridCVT, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
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
    Dim StrSQL As String
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
    Sub LoadRMCode()
        Dim dtRM As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        If CheckBoxType.Checked = False Then
            StrSQL = "SELECT  *  FROM  TBLGroup "
        Else
            StrSQL = "SELECT  *  FROM  TBLGroup "
            StrSQL += " where Typecode = " & PrepareStr(CmbType.SelectedValue)
        End If
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
        CmbMaterial.DataSource = dtRM
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
#End Region

    Private Sub MenuClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub CmdADD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdADD.Click
        Dim fa As New FrmAddCvt
        fa.Text = "Save"
        fa.StrType = GrdDV.Item(oldrow).Row("Type")
        fa.StrMaterial = GrdDV.Item(oldrow).Row("Code")
        fa.ShowDialog()
        LoadGrd()
        change()
    End Sub

    Private Sub CmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdEdit.Click
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult

        msg = "Inventory Edit Material : " & GrdDV.Item(oldrow).Row("code") & " Unit : " & GrdDV.Item(oldrow).Row("ub") ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Inventory"   ' Define title.
        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            If ChkDataINV() Then
                Dim fa As New FrmAddCvt
                fa.Text = "Edit"
                fa.CmbType.Text = GrdDV.Item(oldrow).Row("Type")
                fa.CmbType.Enabled = False
                fa.CmbMaterial.Text = GrdDV.Item(oldrow).Row("code")
                fa.CmbMaterial.Enabled = False
                If GrdDV.Item(oldrow).Row("Type") = "01" Then
                    fa.CmbFinal.Text = ""
                    fa.CmbFinal.Enabled = False
                    fa.Txtrev.Text = ""
                    fa.Txtrev.Enabled = False
                ElseIf GrdDV.Item(oldrow).Row("Type") = "02" Then
                    fa.CmbFinal.Text = ""
                    fa.CmbFinal.Enabled = False
                    fa.Txtrev.Text = GrdDV.Item(oldrow).Row("rev")
                    fa.Txtrev.Enabled = False
                Else
                    fa.CmbFinal.Text = GrdDV.Item(oldrow).Row("final")
                    fa.CmbFinal.Enabled = False
                    fa.Txtrev.Text = GrdDV.Item(oldrow).Row("rev")
                    fa.Txtrev.Enabled = False
                End If
                fa.CmbUnit1.Text = GrdDV.Item(oldrow).Row("UnitBig")
                fa.CmbUnit2.Text = GrdDV.Item(oldrow).Row("Unitsmall")
                fa.TxtQty1.Text = GrdDV.Item(oldrow).Row("BQty")
                fa.TxtQty2.Text = GrdDV.Item(oldrow).Row("SQty")
                fa.ShowDialog()
                LoadGrd()
                change()
            Else
                MsgBox("Don't Edit. It's Link Master File.", MsgBoxStyle.OKOnly)
            End If
        Else
            Exit Sub
        End If
    End Sub

    Private Sub FrmCVT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadGrd()
    End Sub

#Region "Delete Convert"
    Private Sub CmdDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdDel.Click
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult

        msg = "Inventory Delete Material : " & GrdDV.Item(oldrow).Row("code") & " Unit : " & GrdDV.Item(oldrow).Row("ub") ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Inventory"   ' Define title.
        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            If ChkDataINV() Then
                Del()
                MsgBox("Complete", MsgBoxStyle.OKOnly)
                LoadGrd()
                change()
            Else
                MsgBox("Don't Delete. It's Link Master File.", MsgBoxStyle.OKOnly)
            End If
        Else
            Exit Sub
        End If
    End Sub
    Private Function ChkDataINV() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Try
            If GrdDV.Item(oldrow).Row("Type") = "01" Then
                strSQL = " select count(*) from TBLRM"
                strSQL &= " where RMcode  = '" & GrdDV.Item(oldrow).Row("Code") & "'"
                strSQL &= " and  Unit  = '" & GrdDV.Item(oldrow).Row("UnitBig") & "'"
                strSQL &= " and  Typecode  = '" & GrdDV.Item(oldrow).Row("Type") & "'"
            ElseIf GrdDV.Item(oldrow).Row("Type") = "02" Then
                If GrdDV.Item(oldrow).Row("UnitBig") = "BG" Then
                    ChkDataINV = False
                    Exit Function
                Else
                    ChkDataINV = True
                    Exit Function
                End If
            ElseIf GrdDV.Item(oldrow).Row("Type") = "03" Then
                If GrdDV.Item(oldrow).Row("UnitBig") = "BT" Then
                    ChkDataINV = False
                    Exit Function
                Else
                    ChkDataINV = True
                    Exit Function
                End If
            ElseIf GrdDV.Item(oldrow).Row("Type") = "04" Then
                If GrdDV.Item(oldrow).Row("UnitBig") = "UT" Then
                    ChkDataINV = False
                    Exit Function
                ElseIf GrdDV.Item(oldrow).Row("UnitBig") = "M " Then
                    ChkDataINV = False
                    Exit Function
                ElseIf GrdDV.Item(oldrow).Row("UnitBig") = "KG" Then
                    If ChkMaterial() Then
                        ChkDataINV = False
                        Exit Function
                    Else
                        ChkDataINV = True
                        Exit Function
                    End If
                Else
                    ChkDataINV = True
                    Exit Function
                End If
            ElseIf GrdDV.Item(oldrow).Row("Type") = "05" Then
                If GrdDV.Item(oldrow).Row("UnitBig") = "UT" Then
                    ChkDataINV = False
                    Exit Function
                ElseIf GrdDV.Item(oldrow).Row("UnitBig") = "M" Then
                    ChkDataINV = False
                    Exit Function
                Else
                    ChkDataINV = True
                    Exit Function
                End If
            ElseIf GrdDV.Item(oldrow).Row("Type") = "06" Then
                If GrdDV.Item(oldrow).Row("UnitBig") = "UT" Then
                    ChkDataINV = False
                    Exit Function
                Else
                    ChkDataINV = True
                    Exit Function
                End If
            End If

            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar
            If i = 0 Then
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
    Private Function ChkMaterial() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String
        Try
            strSQL = " select count(*) from TBLpresemi"
            strSQL &= " where psemicode  = '" & GrdDV.Item(oldrow).Row("Code") & "'"
            strSQL &= " and  revision  = '" & GrdDV.Item(oldrow).Row("Rev") & "'"
            strSQL &= " and  MaterialType  = '01'"
            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar
            If i <> 0 Then
                ChkMaterial = True
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

    Sub Del()
        Dim strsql As String = String.Empty
        Dim cn As New SqlConnection(C1.Strcon)
        Dim cmd As New SqlCommand(strsql, cn)
        cn.Open()
        Dim t1 As SqlTransaction = cn.BeginTransaction
        cmd.Transaction = t1
        Dim strDate() As String
        Dim strTime As String
        strDate = Split(Date.Now.ToShortDateString, "/")
        strTime = Date.Now.ToShortTimeString
        Try
            strsql = "Delete TBLConvert"
            strsql += " where Type = " & PrepareStr(GrdDV.Item(oldrow).Row("Type"))
            strsql += " and  code = " & PrepareStr(GrdDV.Item(oldrow).Row("Code"))
            strsql += " and   UnitBig = " & PrepareStr(GrdDV.Item(oldrow).Row("UnitBig"))
            If GrdDV.Item(oldrow).Row("Type") = "01" Then
            ElseIf GrdDV.Item(oldrow).Row("Type") = "06" Then
            Else
                strsql += " and  Rev = " & PrepareStr(GrdDV.Item(oldrow).Row("Rev"))
            End If
            cmd.CommandText = strsql
            cmd.ExecuteNonQuery()
            MsgBox("Delete Complete.", MsgBoxStyle.Information, "Inventory Record")
            t1.Commit()
        Catch
            t1.Rollback()
            MsgBox("Rollback data")
        Finally
            cn.Close()
        End Try
    End Sub
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

#Region "Function_Load"
    Private Sub LoadGrd()
        'โหลดข้อมูลจากฐานข้อมูล Sql
        Dim strsql As String
        Dim DA As SqlDataAdapter
        Try
            strsql = " select Type,Final,code,Rev,BQty,u1.shortUnitName ub"
            strsql += " ,c.UnitBig,SQty,u2.shortUnitName us ,c.Unitsmall from TBLConvert c"
            strsql += " left outer join "
            strsql += " TBLUnit u1"
            strsql += " on c.unitbig = u1.unitcode"
            strsql += " left outer join "
            strsql += " TBLUnit u2"
            strsql += " on c.unitsmall = u2.unitcode"
            strsql += " order by Type,Code,Rev,Ub  "
            DA = New SqlDataAdapter(strsql, C1.Strcon)
            Dim CB As New SqlCommandBuilder(DA)
            dt = New DataTable
            DA.Fill(dt)
        Catch
        Finally
        End Try
        '************************************
        dt.TableName = TBL_CVT
        GrdDV = dt.DefaultView
        GrdDV.AllowNew = False
        GrdDV.AllowDelete = False
        '************************************
        DataGridCVT.DataSource = GrdDV
        '************************************
        ResetTableStyle()

        With DataGridCVT
            .BackColor = Color.GhostWhite
            .BackgroundColor = Color.Honeydew
            .BorderStyle = BorderStyle.None
            .CaptionVisible = False
            .Font = New Font("Tahoma", 8.0!)
        End With

        ' Put as much of the formatting as possible here.
        Dim grdTableStyle1 As New DataGridTableStyle
        With grdTableStyle1
            .AlternatingBackColor = Color.MintCream
            .ForeColor = Color.MidnightBlue
            .GridLineColor = Color.RoyalBlue
            .HeaderBackColor = Color.GreenYellow
            .HeaderFont = New Font("Tahoma", 8.0!, FontStyle.Bold)
            .HeaderForeColor = Color.MediumBlue
            .SelectionBackColor = Color.Teal
            .SelectionForeColor = Color.PaleGreen
            .RowHeadersVisible = False

            '' Do not forget to set the MappingName property. 
            '' Without this, the DataGridTableStyle properties
            '' and any associated DataGridColumnStyle objects
            '' will have no effect.
            .MappingName = TBL_CVT
            .PreferredColumnWidth = 125
            .PreferredRowHeight = 15
        End With
        Dim grdColStyle1 As New DataGridColoredLine
        With grdColStyle1
            .HeaderText = "Material Code"
            .MappingName = "Code"
            .Width = 150
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle1_1 As New DataGridColoredLine
        With grdColStyle1_1
            .HeaderText = "Revision"
            .MappingName = "Rev"
            .NullText = ""
            .Width = 75
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle2 As New DataGridColoredLine
        With grdColStyle2
            .HeaderText = "รายการสินค้าคลัง"
            .MappingName = "descitem"
            .Width = 200
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle3 As New DataGridColoredLine
        With grdColStyle3
            .HeaderText = "Unit (Big)"
            .MappingName = "ub"
            .Width = 90
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle4 As New DataGridColoredLine
        With grdColStyle4
            .HeaderText = "Unit (Small)"
            .MappingName = "us"
            .Width = 90
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle8 As New DataGridColoredLine
        With grdColStyle8
            .HeaderText = "Qty "
            .MappingName = "BQty"
            .Width = 75
            .Format = "#,###,##0.0"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle9 As New DataGridColoredLine
        With grdColStyle9
            .HeaderText = "Qty "
            .MappingName = "SQTY"
            .Width = 75
            .ReadOnly = True
            .Format = "#,###,##0.000"
            .Alignment = HorizontalAlignment.Right
        End With
        grdTableStyle1.GridColumnStyles.AddRange _
    (New DataGridColumnStyle() _
    {grdColStyle1, grdColStyle1_1, grdColStyle8, grdColStyle3, grdColStyle9, grdColStyle4})

        DataGridCVT.TableStyles.Add(grdTableStyle1)
        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub
    Private Sub ResetTableStyle()
        ' Clear out the existing TableStyles and result default formatting.
        With DataGridCVT
            .BackgroundColor = SystemColors.InactiveCaptionText
            .CaptionText = ""
            .CaptionBackColor = SystemColors.ActiveCaption
            .TableStyles.Clear()
            .ResetAlternatingBackColor()
            .ResetBackColor()
            .ResetForeColor()
            .ResetGridLineColor()
            .ResetHeaderBackColor()
            .ResetHeaderFont()
            .ResetHeaderForeColor()
            .ResetSelectionBackColor()
            .ResetSelectionForeColor()
            .ResetText()
            .BorderStyle = DefaultGridBorderStyle
        End With
    End Sub
    Sub AddCol1()
        Dim col As New DataColumn
        With col
            .ColumnName = "Name"
            .DataType = GetType(String)
            .DefaultValue = "aaa"
            dt.Columns.Add(col)
        End With
    End Sub
    Sub AddDays()
        Dim i As Integer
        For i = 1 To 10
            Dim col As New DataColumn
            With col
                .ColumnName = "Day" + Format(i, "00")
                .DataType = GetType(Decimal)
                .DefaultValue = 0
                dt.Columns.Add(col)
            End With
        Next
    End Sub
#End Region

#Region "CurrentCell"
    Private Sub DataGrid1_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridCVT.CurrentCellChanged
        oldrow = DataGridCVT.CurrentCell.RowNumber
        oldcol = DataGridCVT.CurrentCell.ColumnNumber
    End Sub
#End Region

#Region "Combobox Change"
    Private Sub CheckBoxType_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxType.CheckedChanged
        If CheckBoxType.Checked = True Then
            CmbType.Enabled = True
            LoadType()
            If CheckBoxMaterial.Checked = True Then
                LoadRMCode()
            End If
        Else
            CmbType.Enabled = False
            CmbType.Text = "Select"
        End If
    End Sub

    Private Sub CheckBoxMaterial_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxMaterial.CheckedChanged
        If CheckBoxMaterial.Checked = True Then
            CmbMaterial.Enabled = True
            LoadRMCode()
        Else
            CmbMaterial.Enabled = False
            CmbMaterial.Text = "Select"
        End If
    End Sub

    Sub change()
        If CheckBoxType.Checked = True And CheckBoxMaterial.Checked = False Then
            GrdDV.RowFilter = " Type = " & PrepareStr(CmbType.SelectedValue)
            DataGridCVT.DataSource = GrdDV
        ElseIf CheckBoxType.Checked = False And CheckBoxMaterial.Checked = True Then
            GrdDV.RowFilter = "  code = " & PrepareStr(CmbMaterial.SelectedValue)
            DataGridCVT.DataSource = GrdDV
        ElseIf CheckBoxType.Checked = True And CheckBoxMaterial.Checked = True Then
            GrdDV.RowFilter = " Type = " & PrepareStr(CmbType.SelectedValue) _
                            & " and code = " & PrepareStr(CmbMaterial.SelectedValue)
            DataGridCVT.DataSource = GrdDV
        End If
    End Sub

    Private Sub CmbType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbType.SelectedIndexChanged
        If CheckBoxMaterial.Checked = True Then
            LoadRMCode()
        End If
        change()
    End Sub

    Private Sub CmbMaterial_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbMaterial.SelectedIndexChanged
        change()
    End Sub
#End Region

    Private Sub MenuClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuClose.Click
        Me.Close()
    End Sub

    Private Sub MenuView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuView.Click
        LoadGrd()
    End Sub
End Class
