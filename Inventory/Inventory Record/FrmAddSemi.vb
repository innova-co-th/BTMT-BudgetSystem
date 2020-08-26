#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Inventory_Record.Common
#End Region

Public Class FrmAddSemi

#Region "Declare"
    Inherits System.Windows.Forms.Form
    Public Shared GrdDV As New DataView
    Protected Const TBL_RM As String = "TBL_RM"
    Public Shared GrdDVType As New DataView
    Protected Const TBL_Type As String = "TBL_Type"
    Public Shared GrdDVTypeMaterial As New DataView
    Protected Const TBL_TypeMaterial As String = "TBL_TypeMaterial"
    Protected DefaultGridBorderStyle As BorderStyle
    Public Shared cm As CurrencyManager
    Dim iTotal, tTotal As Double
    Friend iCmb As String
    Friend txtType As String

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
    Friend WithEvents DataGridRM As System.Windows.Forms.DataGrid
    Friend WithEvents CmdSave As System.Windows.Forms.Button
    Friend WithEvents CmdClose As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtCode As System.Windows.Forms.TextBox
    Friend WithEvents CmdClear As System.Windows.Forms.Button
    Friend WithEvents lblError As System.Windows.Forms.Label
    Friend WithEvents CheckAll As System.Windows.Forms.CheckBox
    Friend WithEvents CmdSearch As System.Windows.Forms.Button
    Friend WithEvents CheckBoxType As System.Windows.Forms.CheckBox
    Friend WithEvents CmbType As System.Windows.Forms.ComboBox
    Friend WithEvents CmbMaterial As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtRev As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents CheckBoxSemi As System.Windows.Forms.CheckBox
    Friend WithEvents txtWidth As System.Windows.Forms.TextBox
    Friend WithEvents lblwidth As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtNum As System.Windows.Forms.TextBox
    Friend WithEvents TxtLenght As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmAddSemi))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.DataGridRM = New System.Windows.Forms.DataGrid
        Me.CmdSave = New System.Windows.Forms.Button
        Me.CmdClose = New System.Windows.Forms.Button
        Me.CmdSearch = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtCode = New System.Windows.Forms.TextBox
        Me.CmdClear = New System.Windows.Forms.Button
        Me.lblError = New System.Windows.Forms.Label
        Me.CheckAll = New System.Windows.Forms.CheckBox
        Me.CheckBoxType = New System.Windows.Forms.CheckBox
        Me.CmbType = New System.Windows.Forms.ComboBox
        Me.CmbMaterial = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtRev = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.CheckBoxSemi = New System.Windows.Forms.CheckBox
        Me.txtWidth = New System.Windows.Forms.TextBox
        Me.lblwidth = New System.Windows.Forms.Label
        Me.TxtNum = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtLenght = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridRM, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.DataGridRM)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(8, 64)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1014, 504)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        '
        'DataGridRM
        '
        Me.DataGridRM.DataMember = ""
        Me.DataGridRM.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridRM.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridRM.Location = New System.Drawing.Point(3, 16)
        Me.DataGridRM.Name = "DataGridRM"
        Me.DataGridRM.Size = New System.Drawing.Size(1008, 485)
        Me.DataGridRM.TabIndex = 0
        '
        'CmdSave
        '
        Me.CmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdSave.Image = CType(resources.GetObject("CmdSave.Image"), System.Drawing.Image)
        Me.CmdSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdSave.Location = New System.Drawing.Point(888, 570)
        Me.CmdSave.Name = "CmdSave"
        Me.CmdSave.Size = New System.Drawing.Size(64, 56)
        Me.CmdSave.TabIndex = 11
        Me.CmdSave.Text = "Save"
        Me.CmdSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdClose
        '
        Me.CmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdClose.Image = CType(resources.GetObject("CmdClose.Image"), System.Drawing.Image)
        Me.CmdClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdClose.Location = New System.Drawing.Point(952, 570)
        Me.CmdClose.Name = "CmdClose"
        Me.CmdClose.Size = New System.Drawing.Size(64, 56)
        Me.CmdClose.TabIndex = 12
        Me.CmdClose.Text = "Close"
        Me.CmdClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdSearch
        '
        Me.CmdSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdSearch.Image = CType(resources.GetObject("CmdSearch.Image"), System.Drawing.Image)
        Me.CmdSearch.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdSearch.Location = New System.Drawing.Point(940, 7)
        Me.CmdSearch.Name = "CmdSearch"
        Me.CmdSearch.Size = New System.Drawing.Size(80, 57)
        Me.CmdSearch.TabIndex = 8
        Me.CmdSearch.Text = "Search"
        Me.CmdSearch.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 16)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Semi (Material)"
        '
        'TxtCode
        '
        Me.TxtCode.Location = New System.Drawing.Point(112, 40)
        Me.TxtCode.Name = "TxtCode"
        Me.TxtCode.TabIndex = 1
        Me.TxtCode.Text = ""
        '
        'CmdClear
        '
        Me.CmdClear.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CmdClear.Image = CType(resources.GetObject("CmdClear.Image"), System.Drawing.Image)
        Me.CmdClear.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdClear.Location = New System.Drawing.Point(8, 570)
        Me.CmdClear.Name = "CmdClear"
        Me.CmdClear.Size = New System.Drawing.Size(80, 56)
        Me.CmdClear.TabIndex = 13
        Me.CmdClear.Text = "Clear"
        Me.CmdClear.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'lblError
        '
        Me.lblError.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblError.ForeColor = System.Drawing.Color.Red
        Me.lblError.Location = New System.Drawing.Point(216, 48)
        Me.lblError.Name = "lblError"
        Me.lblError.Size = New System.Drawing.Size(24, 8)
        Me.lblError.TabIndex = 8
        Me.lblError.Text = "***"
        Me.lblError.Visible = False
        '
        'CheckAll
        '
        Me.CheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CheckAll.Location = New System.Drawing.Point(764, 608)
        Me.CheckAll.Name = "CheckAll"
        Me.CheckAll.Size = New System.Drawing.Size(88, 16)
        Me.CheckAll.TabIndex = 10
        Me.CheckAll.Text = "ADD Check"
        '
        'CheckBoxType
        '
        Me.CheckBoxType.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CheckBoxType.Location = New System.Drawing.Point(740, 10)
        Me.CheckBoxType.Name = "CheckBoxType"
        Me.CheckBoxType.Size = New System.Drawing.Size(56, 16)
        Me.CheckBoxType.TabIndex = 7
        Me.CheckBoxType.Text = "Type"
        '
        'CmbType
        '
        Me.CmbType.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmbType.Enabled = False
        Me.CmbType.Location = New System.Drawing.Point(804, 8)
        Me.CmbType.Name = "CmbType"
        Me.CmbType.Size = New System.Drawing.Size(128, 21)
        Me.CmbType.TabIndex = 9
        '
        'CmbMaterial
        '
        Me.CmbMaterial.Location = New System.Drawing.Point(112, 8)
        Me.CmbMaterial.Name = "CmbMaterial"
        Me.CmbMaterial.Size = New System.Drawing.Size(152, 21)
        Me.CmbMaterial.TabIndex = 0
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 10)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(96, 16)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "Type Material"
        '
        'TxtRev
        '
        Me.TxtRev.Location = New System.Drawing.Point(328, 40)
        Me.TxtRev.Name = "TxtRev"
        Me.TxtRev.Size = New System.Drawing.Size(40, 20)
        Me.TxtRev.TabIndex = 2
        Me.TxtRev.Text = ""
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(248, 42)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 16)
        Me.Label3.TabIndex = 19
        Me.Label3.Text = "Semi  Rev."
        '
        'CheckBoxSemi
        '
        Me.CheckBoxSemi.Checked = True
        Me.CheckBoxSemi.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxSemi.Location = New System.Drawing.Point(272, 10)
        Me.CheckBoxSemi.Name = "CheckBoxSemi"
        Me.CheckBoxSemi.Size = New System.Drawing.Size(136, 16)
        Me.CheckBoxSemi.TabIndex = 20
        Me.CheckBoxSemi.Text = "Final Semi (Material)"
        '
        'txtWidth
        '
        Me.txtWidth.Location = New System.Drawing.Point(496, 8)
        Me.txtWidth.Name = "txtWidth"
        Me.txtWidth.Size = New System.Drawing.Size(56, 20)
        Me.txtWidth.TabIndex = 4
        Me.txtWidth.Text = ""
        Me.txtWidth.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblwidth
        '
        Me.lblwidth.Location = New System.Drawing.Point(456, 8)
        Me.lblwidth.Name = "lblwidth"
        Me.lblwidth.Size = New System.Drawing.Size(40, 16)
        Me.lblwidth.TabIndex = 23
        Me.lblwidth.Text = "Width"
        '
        'TxtNum
        '
        Me.TxtNum.Location = New System.Drawing.Point(408, 40)
        Me.TxtNum.Name = "TxtNum"
        Me.TxtNum.Size = New System.Drawing.Size(32, 20)
        Me.TxtNum.TabIndex = 3
        Me.TxtNum.Text = ""
        Me.TxtNum.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(384, 42)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(16, 16)
        Me.Label1.TabIndex = 26
        Me.Label1.Text = "N"
        '
        'TxtLenght
        '
        Me.TxtLenght.Location = New System.Drawing.Point(496, 40)
        Me.TxtLenght.Name = "TxtLenght"
        Me.TxtLenght.Size = New System.Drawing.Size(56, 20)
        Me.TxtLenght.TabIndex = 5
        Me.TxtLenght.Text = ""
        Me.TxtLenght.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(448, 40)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(48, 16)
        Me.Label6.TabIndex = 28
        Me.Label6.Text = "Lenght"
        '
        'FrmAddSemi
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1030, 632)
        Me.Controls.Add(Me.TxtLenght)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.TxtNum)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtWidth)
        Me.Controls.Add(Me.lblwidth)
        Me.Controls.Add(Me.CheckBoxSemi)
        Me.Controls.Add(Me.TxtRev)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.CmbMaterial)
        Me.Controls.Add(Me.CmbType)
        Me.Controls.Add(Me.CheckBoxType)
        Me.Controls.Add(Me.CheckAll)
        Me.Controls.Add(Me.lblError)
        Me.Controls.Add(Me.CmdClear)
        Me.Controls.Add(Me.TxtCode)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.CmdSearch)
        Me.Controls.Add(Me.CmdClose)
        Me.Controls.Add(Me.CmdSave)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmAddSemi"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Semi (Material)"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataGridRM, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "CONSTANT"
    Dim DT As New DataTable
    Dim StrSQL As String
    Dim oldrow As Integer
    Dim C1 As New SQLData("ACCINV")
#End Region

#Region "Delegate function"
    Public Shared Function MyGetSeqLine(ByVal row As Integer) As CellColor
        Dim c As CellColor
        c.ForeG = CInt(GrdDV.Item(row).Item(0))
        c.BackG = CInt(GrdDV.Item(row).Item(1))
        c.LfItem = Mid(GrdDV.Item(row).Item(3), 1, 4)
        Return c
    End Function
#End Region

#Region "COMBOBOX"
    Sub LoadType()
        Dim dtType As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        StrSQL = "SELECT  TypeCode,TypeName "
        StrSQL &= "  FROM  TblType  "
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
    Sub LoadTypeMaterial()
        Dim dtTypeMaterial As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        StrSQL = "SELECT  MaterialCode,MaterialName "
        StrSQL &= "  FROM  TblTypeMaterial  "
        StrSQL &= "  Where  descname like 'Semi' "
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtTypeMaterial = New DataTable
            DA.Fill(dtTypeMaterial)
        Catch
        Finally
        End Try
        dtTypeMaterial.TableName = TBL_TypeMaterial
        GrdDVTypeMaterial = dtTypeMaterial.DefaultView
        '************************************
        CmbMaterial.DisplayMember = "MaterialName"
        CmbMaterial.ValueMember = "MaterialCode"
        CmbMaterial.DataSource = dtTypeMaterial
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
#End Region

#Region "Function_Load"
    Private Sub LoadRM()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        If CmdSave.Text = "Edit" Then
            StrSQL = "  select * from "
            StrSQL &= "   ( "
            StrSQL &= "  select * from "
            StrSQL &= "  ("
            StrSQL &= "   select t.TypeCode,Typename,MasterCode,Revision as MRev,isnull(descname,'') as descName"
            StrSQL &= "   ,Code,RMRevision as Revision,Qty as Qty ,g.Unit"
            StrSQL &= "   FROM   "
            StrSQL &= "   ("
            StrSQL &= "   select  distinct Typecode ,MasterCode,isnull(Revision,'') Revision,isnull(rmcode,'') code"
            StrSQL &= "   ,rmRevision as RMRevision,Qty,Unit from TBLGroup t "
            StrSQL &= "  left outer join "
            StrSQL &= "  ( SELECT  distinct  MasterCode,Revision,rmcode,rmRevision,Qty,Unit"
            StrSQL &= "    FROM         TBLMaster"
            StrSQL &= "  ) yy"
            StrSQL &= "  on t.code = yy.MasterCode"
            StrSQL &= "  )g"
            StrSQL &= "  left outer join "
            StrSQL &= "   TBLTYPE t"
            StrSQL &= "   on g.typecode = t.typecode"
            StrSQL &= "   left outer join "
            StrSQL &= "   TBLRM r"
            StrSQL &= "  on r.rmcode = g.code"
            StrSQL &= "   ) PreSemi"
            StrSQL &= "  where typecode <>'02' and "
            StrSQL &= "   code not in "
            StrSQL &= "  ("
            StrSQL &= "   select code  from TBLGroup"
            StrSQL &= "   where code in ("
            StrSQL &= "   select CompCode from TBLcompound "
            StrSQL &= "   where Active = '0')) "
            StrSQL &= "  ) xxx"
            StrSQL &= " where Mastercode = '" & TxtCode.Text.Trim & "'and MRev = '" & TxtRev.Text.Trim & "'"
            StrSQL &= " and Typecode <>'01'     order by Typename,Code,descName"
        Else
            StrSQL = "  select Typecode,TypeName,descName,b.code,Qty,'g' Unit from"
            StrSQL &= "  (select t.Typecode ,TypeName,code from TBLType t"
            StrSQL &= "  left outer join TBLGroup  g"
            StrSQL &= "   on t.typecode=g.typecode"
            StrSQL &= "   )a"
            StrSQL &= "  left outer join "
            StrSQL &= "   ("
            StrSQL &= "  SELECT  distinct  Finalcompound code ,null DescName,compcode,0.00 Qty"
            StrSQL &= "   FROM         TBLCompound"
            StrSQL &= "   where Compcode not in "
            StrSQL &= "   ( select code from TblGroup where Typecode ='01')"
            StrSQL &= "   and active = 1"
            StrSQL &= "   union  SELECT  distinct  Final code ,null DescName,psemicode,0.00 Qty"
            StrSQL &= "  FROM         TBLpresemi"
            StrSQL &= "    where MaterialType   not in ( '01','02')"
            StrSQL &= "   and active = 1"
            StrSQL &= "   union"
            StrSQL &= "   select RMcode,DescName,RMcode, 0.00 Qty  from TblRM"
            StrSQL &= "  where descName like '%Steel%' or descName like '%Bead%' or descName like '%Nylon%' "
            StrSQL &= "   )b"
            StrSQL &= "    on a.code = b.compcode"
            StrSQL &= "   where b.code is not null"
            StrSQL &= "   order by descName desc,typecode desc ,b.code"
        End If

        If Not DT Is Nothing Then
            If DT.Rows.Count >= 1 Then
                DT.Clear()
            End If
        End If

        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CB As New SqlCommandBuilder(DA)
            DT = New DataTable
            DA.Fill(DT)
        Catch
            MsgBox("Can't Select Data.", MsgBoxStyle.Critical, "Load Data")
        Finally
        End Try
        '************************************
        DT.TableName = TBL_RM
        GrdDV = DT.DefaultView
        GrdDV.AllowNew = False
        GrdDV.AllowDelete = False
        '************************************
        DataGridRM.DataSource = GrdDV
        '************************************
        'Dim i As Integer
        'Dim c34 As String = Chr(34)
        'For i = 0 To dtReqInv.Columns.Count - 1
        '    Dim col As String = dtReqInv.Columns(i).ColumnName
        '    Dim coltype As String = dtReqInv.Columns(i).DataType.FullName
        '    coltype = coltype.Replace("System.", "")
        '    coltype = coltype.Replace("Int32", "integer")
        '    coltype = coltype.Replace("Int16", "integer")
        '    coltype = coltype.Replace("String", "string")
        '    coltype = coltype.Replace("Decimal", "decimal")
        '    Debug.WriteLine("<xs:element name=" & c34 & col.Trim & c34 & "  type= " & c34 & "xs:" & coltype & c34 & " minOccurs=" & c34 & "0" & c34 & "/>")
        'Next
        ResetTableStyle()

        With DataGridRM
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
            .AllowSorting = False

            '' Do not forget to set the MappingName property. 
            '' Without this, the DataGridTableStyle properties
            '' and any associated DataGridColumnStyle objects
            '' will have no effect.
            .MappingName = TBL_RM
            .PreferredColumnWidth = 125
            .PreferredRowHeight = 15
        End With
        Dim grdColStyle1 As New DataGridColoredLine2
        With grdColStyle1
            .HeaderText = "Typename"
            .MappingName = "Typename"
            .Width = 110
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle2 As New DataGridColoredLine2
        With grdColStyle2
            .HeaderText = ""
            .MappingName = "DescName"
            .NullText = ""
            .Width = 120
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle3 As New DataGridColoredLine2
        With grdColStyle3
            .HeaderText = "Code"
            .MappingName = "Code"
            .Width = 110
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        cm = CType(Me.BindingContext(DataGridRM.DataSource, DataGridRM.DataMember), CurrencyManager)
        Dim c As CheckRowHeader
        c = AddressOf CheckRowHeader

        Dim grdColStyle6 As New DataGridQtyBox(c)
        With grdColStyle6
            .HeaderText = "Q'ty/Unit"
            .MappingName = "Qty"
            .Format = "###,###.000"
            .Width = 80
            .Alignment = HorizontalAlignment.Right
            .NullText = ""
        End With
        cm = CType(Me.BindingContext(DataGridRM.DataSource, DataGridRM.DataMember), CurrencyManager)
        Dim d As CheckRow
        d = AddressOf CheckRow

        Dim grdColStyle5 As New DataGridUnitBox(d)
        With grdColStyle5
            .HeaderText = "(Unit)"
            .MappingName = "Unit"
            .Width = 100
            .Alignment = HorizontalAlignment.Center
            .NullText = ""
        End With

        grdTableStyle1.GridColumnStyles.AddRange _
(New DataGridColumnStyle() _
{grdColStyle3, grdColStyle1, grdColStyle2, grdColStyle6, grdColStyle5})

        DataGridRM.TableStyles.Add(grdTableStyle1)
        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub


    Public Shared Function CheckRowHeader(ByVal row As Integer) As Boolean
        Dim c As Boolean = False
        Try
            If GrdDV.Item(row).Item("Code").ToString.Trim = "" Then
                c = True
            Else
                c = False
            End If
        Catch ex As Exception
            c = False
        End Try

        Return c
    End Function

    Public Shared Function CheckRow(ByVal row As String) As Boolean
        Dim d As Boolean = False
        Try
            If GrdDV.Item(row).Item("Code").ToString.Trim = "" Then
                d = True
            Else
                d = False
            End If
        Catch ex As Exception
            d = False
        End Try

        Return d
    End Function

    Private Sub ResetTableStyle()
        ' Clear out the existing TableStyles and result default formatting.
        With DataGridRM
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
#End Region

    Private Sub FrmAddSemi_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadRM()
        If CmdSave.Text = "Edit" Then
            LoadTypeMaterial()
            CmbMaterial.Text = iCmb
        Else
            LoadType()
            LoadTypeMaterial()
        End If
        CmbMaterial.Text = txtType.Trim
    End Sub

    Private Sub CmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClose.Click
        Me.Close()
    End Sub

    Private Sub CmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSave.Click
        If TxtCode.Text.Trim = "" Then
            TxtCode.Focus()
            lblError.Visible = True
            Exit Sub
        Else
            lblError.Visible = False
        End If

        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        Dim aDr() As DataRow
        GrdDV.RowFilter = " Qty <> 0.000"
        aDr = GrdDV.Table.Select(GrdDV.RowFilter)
        iTotal = 0
        Dim dr As DataRow
        Dim qq, qsum As Double
        For Each dr In aDr
            With dr
                If IIf(.Item("Code") Is System.DBNull.Value, "", .Item("Code")) <> "" Then
                    If CmdSave.Text = "Save" Then
                        If CmbMaterial.SelectedValue = "13" Then
                            qq = .Item("Qty")
                        ElseIf CmbMaterial.SelectedValue = "14" Then
                            qq = .Item("Qty") / TxtNum.Text.Trim
                        Else
                            qq = ((.Item("Qty") / TxtLenght.Text.Trim) * 1000)
                        End If
                        iTotal = CSng(iTotal + qq)
                        tTotal = CSng(tTotal + .Item("Qty"))
                    Else
                        iTotal = CSng(iTotal + .Item("Qty"))
                        tTotal = CSng(tTotal + .Item("Qty"))
                    End If
                    qsum = qsum + .Item("Qty")
                End If
            End With
        Next
        If ChkData() Then
        Else
            MsgBox("It's Duplicate. Please CheckData again.", MsgBoxStyle.OKOnly, "Semi")
            Exit Sub
        End If

        msg = "Semi(Material) Total Qty :  " & qsum & " g.  Total Qty Per(Meter/Line) : " & CSng(iTotal) & " g." ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Semi(Material)"   ' Define title.
        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            If CheckBoxSemi.Checked Then
                UPSemi()
            End If
            RM()
        Else
            Exit Sub
        End If
    End Sub

#Region "Final"
    Private Function Chkfinal() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String
        Try
            strSQL = " Select count(*) from TBLSemi "
            strSQL += " Where final = " & PrepareStr(TxtCode.Text.Trim)
            strSQL += " and Active = " & PrepareStr(1)
            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar
            If i <> 0 Then
                Chkfinal = True
            End If
            cmSQL = New SqlCommand(strSQL, cnSQL)
            cmSQL.ExecuteNonQuery()
            cnSQL.Close()

            cmSQL.Dispose()
            cnSQL.Dispose()
            '--------------------------------------------------------------------------------------
        Catch Exp As SqlException
            ' MsgBox(Exp.Message, MsgBoxStyle.Critical, "SQL Error")
        Catch Exp As Exception
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "General Error")
        End Try
    End Function
    Sub UpComp()
        Dim strsql As String = String.Empty
        Dim cn As New SqlConnection(C1.Strcon)
        Dim cmd As New SqlCommand(strsql, cn)
        cn.Open()
        Dim t1 As SqlTransaction = cn.BeginTransaction
        cmd.Transaction = t1
        Dim strDate, SD() As String
        SD = Split(Now.ToShortDateString, "/")
        strDate = SD(2) + SD(1) + SD(0)
        Try
            strsql = "  "
            strsql += "update  TBLsemi "
            strsql += " set Active = " & PrepareStr(0)
            strsql += " , dateUp = " & PrepareStr(strDate)
            strsql += " Where  final = " & PrepareStr(TxtCode.Text.Trim)
            strsql += " and Active = " & PrepareStr(1)
            t1.Commit()
        Catch
            t1.Rollback()
            MsgBox("Rollback data")
        Finally
            cn.Close()
        End Try
    End Sub
#End Region

    Private Sub DataGridRM_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridRM.CurrentCellChanged
        oldrow = DataGridRM.CurrentCell.RowNumber
    End Sub

#Region "RM"
    Private Function ChkData() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String
        Try
            strSQL = " Select count(*) from TBLsemi "
            strSQL += " Where semicode = " & PrepareStr(TxtCode.Text.Trim)
            strSQL += " and  Revision  = " & PrepareStr(TxtRev.Text.Trim)

            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar
            If i = 0 Then
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
    Sub UPSemi()
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Try
            strSQL = " Update TblSemi"
            strSQL &= " set Active = 0"
            strSQL &= " where SemiCode = '" & TxtCode.Text.Trim & "'"

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
    End Sub
    Private Function iNo() As Long
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim drSQL As SqlDataReader
        Dim strSQL As String
        Try

            ' Build Insert statement to insert 
            strSQL = "  SELECT   top 1 Revision "
            strSQL &= "  FROM   TBLSemi"
            strSQL &= " Where semiCode  = '" & TxtCode.Text.Trim & "'"
            strSQL &= "  order by Revision desc"

            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            cmSQL.ExecuteScalar()
            drSQL = cmSQL.ExecuteReader()
            If drSQL.HasRows Then
                If drSQL.Read() Then
                    iNo = CInt(drSQL.Item("Revision").ToString())
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
    Sub RM()
        Dim strsql As String = String.Empty
        Dim cn As New SqlConnection(C1.Strcon)
        Dim cmd As New SqlCommand(strsql, cn)
        cn.Open()
        Dim t1 As SqlTransaction = cn.BeginTransaction
        cmd.Transaction = t1
        Dim strDate, str() As String
        str = Split(Now.Date.ToShortDateString, "/")
        strDate = str(2) + str(1) + str(0)
        If CmdSave.Text = "Save" Then
            Try
                Dim aDr() As DataRow
                GrdDV.RowFilter = " Qty <> 0"
                aDr = GrdDV.Table.Select(GrdDV.RowFilter)
                If UBound(aDr) < 0 Then
                    Exit Sub
                End If
                Dim dr As DataRow
                For Each dr In aDr
                    With dr
                        If IIf(.Item("Code") Is System.DBNull.Value, "", .Item("Code")) <> "" Then
                            strsql = "Insert TBLMaster"
                            strsql += " Values(" & PrepareStr(TxtCode.Text.Trim)
                            strsql += "," & PrepareStr(TxtRev.Text.Trim)
                            strsql += "," & PrepareStr(.Item("Code"))
                            strsql += "," & PrepareStr("")
                            If CmbMaterial.SelectedValue = "13" Then
                                strsql += "," & PrepareStr(CSng(.Item("Qty")))
                            ElseIf CmbMaterial.SelectedValue = "14" Then
                                strsql += "," & PrepareStr(CSng(.Item("Qty") / TxtNum.Text.Trim))
                            Else
                                strsql += "," & PrepareStr(CSng(((.Item("Qty") / TxtLenght.Text.Trim) * 1000)))
                            End If
                            strsql += "," & PrepareStr(.Item("Unit"))
                            'If CmbMaterial.SelectedValue = "10" Then
                            '    strsql += "," & PrepareStr(CSng(((.Item("Qty") / TxtLenght.Text.Trim) * 1000) / tTotal * 100))
                            'Else
                                strsql += "," & PrepareStr(CSng(.Item("Qty") / tTotal * 100))
                            'End If
                            strsql += ")"
                            cmd.CommandText = strsql
                            cmd.ExecuteNonQuery()
                        End If
                    End With
                Next

                Try
                    strsql = " Insert  TblGroup "
                    strsql += " values ( '05',"
                    strsql += PrepareStr(TxtCode.Text.Trim) & ")"

                    strsql += " "
                    strsql += "Insert  TBLSemi "
                    strsql += " values (" & PrepareStr(TxtCode.Text.Trim) & ","
                    strsql += PrepareStr(TxtCode.Text.Trim) & ","
                    strsql += PrepareStr(TxtRev.Text.Trim) & ","
                    strsql += PrepareStr(CmbMaterial.SelectedValue) & ","
                    strsql += PrepareStr(CSng(iTotal)) & ","
                    strsql += PrepareStr(txtWidth.Text.Trim) & ","
                    strsql += PrepareStr(TxtLenght.Text.Trim) & ","
                    strsql += PrepareStr(TxtNum.Text.Trim) & ","
                    If CheckBoxSemi.Checked = True Then
                        strsql += PrepareStr(1) & ","
                    Else
                        strsql += PrepareStr(0) & ","
                    End If
                    strsql += PrepareStr(strDate.Trim) & ","
                    If CmbMaterial.SelectedValue = "13" Then
                        strsql += PrepareStr(1) & ")"
                    ElseIf CmbMaterial.SelectedValue = "14" Then
                        strsql += PrepareStr(1) & ")"
                    ElseIf CmbMaterial.SelectedValue = "07" Then
                        strsql += PrepareStr(1) & ")"
                    ElseIf CmbMaterial.SelectedValue = "08" Then
                        strsql += PrepareStr(1) & ")"
                    ElseIf CmbMaterial.SelectedValue = "12" Then
                        strsql += PrepareStr(1) & ")"
                    ElseIf CmbMaterial.SelectedValue = "04" Then
                        strsql += PrepareStr(1) & ")"
                    Else
                        strsql += PrepareStr(2) & ")"
                    End If


                    strsql += ""
                    strsql += " Insert TBLConvert "
                    strsql += " Values('05'"
                    strsql += "," & PrepareStr(TxtCode.Text.Trim)
                    strsql += "," & PrepareStr(TxtCode.Text.Trim)
                    strsql += "," & PrepareStr(TxtRev.Text.Trim)
                    strsql += "," & PrepareStr("KG")
                    strsql += "," & PrepareStr("KG")
                    strsql += "," & PrepareStr(1)
                    strsql += "," & PrepareStr(1)
                    strsql += ")"

                    strsql += ""
                    strsql += " Insert TBLConvert "
                    strsql += " Values('05'"
                    strsql += "," & PrepareStr(TxtCode.Text.Trim)
                    strsql += "," & PrepareStr(TxtCode.Text.Trim)
                    strsql += "," & PrepareStr(TxtRev.Text.Trim)
                    If CmbMaterial.SelectedValue = "13" Then
                        strsql += "," & PrepareStr("UT")
                    ElseIf CmbMaterial.SelectedValue = "14" Then
                        strsql += "," & PrepareStr("UT")
                    Else
                        strsql += "," & PrepareStr("M")
                    End If
                    strsql += "," & PrepareStr("KG")
                    strsql += "," & PrepareStr(1)
                    strsql += "," & PrepareStr(CSng(iTotal / 1000))
                    strsql += ")"


                    cmd.CommandText = strsql
                    cmd.ExecuteNonQuery()
                Catch
                End Try

                t1.Commit()
            Catch
                t1.Rollback()
                MsgBox("Rollback data")
                Exit Sub
            Finally
                cn.Close()
            End Try
        ElseIf CmdSave.Text = "Edit" Then
            Try
                Dim aDr() As DataRow
                ' GrdDV.RowFilter = " PQty <> 0.000  and QTY <> 0"
                aDr = GrdDV.Table.Select()
                If UBound(aDr) < 0 Then
                    Exit Sub
                End If
                Dim dr As DataRow
                For Each dr In aDr
                    With dr
                        If IIf(.Item("Code") Is System.DBNull.Value, "", .Item("Code")) <> "" Then
                            strsql = "Update TBLMaster"
                            strsql += " set Qty = " & PrepareStr(.Item("Qty"))
                            strsql += " where Mastercode = " & PrepareStr(TxtCode.Text.Trim)
                            strsql += " and Revision = " & PrepareStr(TxtRev.Text.Trim)
                            strsql += " and RMCode = " & PrepareStr(.Item("Code"))
                            cmd.CommandText = strsql
                            cmd.ExecuteNonQuery()
                        End If
                    End With
                Next

                Try
                    strsql = " Update TBLSemi "
                    strsql += " set Qty = " & PrepareStr(CSng(iTotal))
                    If CheckBoxSemi.Checked = True Then
                        strsql += ", Active = " & PrepareStr(1)
                    Else
                        strsql += ", Active = " & PrepareStr(0)
                    End If
                    strsql += " Where semiCode = " & PrepareStr(TxtCode.Text.Trim)
                    strsql += " and Revision  = " & PrepareStr(TxtRev.Text.Trim)



                    If CmbMaterial.SelectedValue = "13" Then
                        strsql += ""
                        strsql += " Update TBLConvert "
                        strsql += " set SQty = " & PrepareStr(CSng(iTotal))
                        strsql += " where code = " & PrepareStr(TxtCode.Text.Trim)
                        strsql += " and Rev = " & PrepareStr(TxtRev.Text.Trim)
                        strsql += " and final = " & PrepareStr(TxtCode.Text.Trim)
                    ElseIf CmbMaterial.SelectedValue = "14" Then
                        strsql += ""
                        strsql += " Update TBLConvert "
                        strsql += " set SQty = " & PrepareStr(CSng(iTotal / 2))
                        strsql += " where code = " & PrepareStr(TxtCode.Text.Trim)
                        strsql += " and Rev = " & PrepareStr(TxtRev.Text.Trim)
                        strsql += " and final = " & PrepareStr(TxtCode.Text.Trim)
                    Else
                        strsql += ""
                        strsql += " Update TBLConvert "
                        strsql += " set SQty = " & PrepareStr(CSng(((iTotal / TxtLenght.Text.Trim) * 1000)))
                        strsql += " where code = " & PrepareStr(TxtCode.Text.Trim)
                        strsql += " and Rev = " & PrepareStr(TxtRev.Text.Trim)
                        strsql += " and final = " & PrepareStr(TxtCode.Text.Trim)
                    End If

                    cmd.CommandText = strsql
                    cmd.ExecuteNonQuery()
                Catch

                End Try

                t1.Commit()
                MsgBox("Update Complete.", MsgBoxStyle.Information, "Semi(Material)")
            Catch
                t1.Rollback()
                MsgBox("Rollback data")
                Exit Sub
            Finally
                cn.Close()
            End Try

        Else
        End If
        If CheckAll.Checked Then
            LoadRM()
            TxtCode.Text = ""
            txtWidth.Text = ""
            TxtLenght.Text = ""
            TxtNum.Text = ""
            TxtRev.Text = ""
        Else
            Me.Close()
        End If

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

    Private Sub TxtName_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case Else
        End Select

    End Sub

    Private Sub CmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClear.Click
        LoadRM()
    End Sub

    Private Sub CmdSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSearch.Click
        If CheckBoxType.Checked = True Then
            CmbType.Enabled = True
            GrdDV.RowFilter = " TypeCode like'%" & CmbType.SelectedValue & "%'"
        End If
        DataGridRM.DataSource = GrdDV
    End Sub

    Private Sub CheckBoxType_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxType.CheckedChanged
        If CheckBoxType.Checked = True Then
            CmbType.Enabled = True
        Else
            CmbType.Enabled = False
        End If
    End Sub

    Private Sub TxtCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCode.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                TxtCode.Text = TxtCode.Text.ToUpper
                Dim i As Integer
                If CmdSave.Text = "Save" Then
                    i = iNo() + 1
                    TxtRev.Text = Format(i, "000")
                Else
                End If
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case Else
        End Select
    End Sub

    Private Sub TxtRev_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtRev.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                SendKeys.Send("{TAB}")
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

    Private Sub TxtNum_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtNum.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                SendKeys.Send("{TAB}")
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

    Private Sub txtWidth_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtWidth.KeyPress
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

    Private Sub TxtLenght_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtLenght.KeyPress
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
End Class
