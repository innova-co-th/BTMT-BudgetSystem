#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports System.Text
Imports Inventory_Record.Common
Imports Inventory_Record.FrmMain
#End Region

Public Class FrmSemi

#Region "Declare"
    Inherits System.Windows.Forms.Form
    Dim GrdDV As New DataView
    Protected Const TBL_RM As String = "TBL_RM"
    Dim GrdDVPreSemi As New DataView
    Protected Const TBL_PreSemi As String = "TBL_PreSemi"
    Protected Const TBL_Semi As String = "TBL_Semi"
    Dim GrdDVType As New DataView
    Protected Const TBL_Type As String = "TBL_Type"
    Dim C1 As New SQLData("ACCINV")
    Friend WithEvents CmdImport As System.Windows.Forms.Button
    Friend WithEvents CmdExport As System.Windows.Forms.Button

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
    Friend WithEvents CmdSave As System.Windows.Forms.Button
    Friend WithEvents CmdClose As System.Windows.Forms.Button
    Friend WithEvents CmdEdit As System.Windows.Forms.Button
    Friend WithEvents DataGridCOM As System.Windows.Forms.DataGrid
    Friend WithEvents CheckBoxPreSemi As System.Windows.Forms.CheckBox
    Friend WithEvents CmbMaterial As System.Windows.Forms.ComboBox
    Friend WithEvents CheckBoxType As System.Windows.Forms.CheckBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents CmdDel As System.Windows.Forms.Button
    Friend WithEvents CheckBoxTire As System.Windows.Forms.CheckBox
    Friend WithEvents CmbTire As System.Windows.Forms.ComboBox
    Friend WithEvents ChkAvtive As System.Windows.Forms.CheckBox
    Friend WithEvents cmdActive As System.Windows.Forms.Button
    Friend WithEvents CmbSemi As System.Windows.Forms.ComboBox
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmSemi))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.DataGridCOM = New System.Windows.Forms.DataGrid()
        Me.CmdSave = New System.Windows.Forms.Button()
        Me.CmdClose = New System.Windows.Forms.Button()
        Me.CmdEdit = New System.Windows.Forms.Button()
        Me.CmbSemi = New System.Windows.Forms.ComboBox()
        Me.CheckBoxPreSemi = New System.Windows.Forms.CheckBox()
        Me.CmbMaterial = New System.Windows.Forms.ComboBox()
        Me.CheckBoxType = New System.Windows.Forms.CheckBox()
        Me.CmdDel = New System.Windows.Forms.Button()
        Me.CheckBoxTire = New System.Windows.Forms.CheckBox()
        Me.CmbTire = New System.Windows.Forms.ComboBox()
        Me.cmdActive = New System.Windows.Forms.Button()
        Me.ChkAvtive = New System.Windows.Forms.CheckBox()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.CmdImport = New System.Windows.Forms.Button()
        Me.CmdExport = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridCOM, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.DataGridCOM)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(8, 72)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1042, 520)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'DataGridCOM
        '
        Me.DataGridCOM.DataMember = ""
        Me.DataGridCOM.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridCOM.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridCOM.Location = New System.Drawing.Point(3, 16)
        Me.DataGridCOM.Name = "DataGridCOM"
        Me.DataGridCOM.Size = New System.Drawing.Size(1036, 501)
        Me.DataGridCOM.TabIndex = 0
        '
        'CmdSave
        '
        Me.CmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdSave.Image = CType(resources.GetObject("CmdSave.Image"), System.Drawing.Image)
        Me.CmdSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdSave.Location = New System.Drawing.Point(904, 594)
        Me.CmdSave.Name = "CmdSave"
        Me.CmdSave.Size = New System.Drawing.Size(72, 56)
        Me.CmdSave.TabIndex = 1
        Me.CmdSave.Text = "Add"
        Me.CmdSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdClose
        '
        Me.CmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdClose.Image = CType(resources.GetObject("CmdClose.Image"), System.Drawing.Image)
        Me.CmdClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdClose.Location = New System.Drawing.Point(976, 594)
        Me.CmdClose.Name = "CmdClose"
        Me.CmdClose.Size = New System.Drawing.Size(72, 56)
        Me.CmdClose.TabIndex = 2
        Me.CmdClose.Text = "Close"
        Me.CmdClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdEdit
        '
        Me.CmdEdit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdEdit.Image = CType(resources.GetObject("CmdEdit.Image"), System.Drawing.Image)
        Me.CmdEdit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdEdit.Location = New System.Drawing.Point(832, 594)
        Me.CmdEdit.Name = "CmdEdit"
        Me.CmdEdit.Size = New System.Drawing.Size(72, 56)
        Me.CmdEdit.TabIndex = 3
        Me.CmdEdit.Text = "Edit"
        Me.CmdEdit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.CmdEdit.Visible = False
        '
        'CmbSemi
        '
        Me.CmbSemi.Location = New System.Drawing.Point(144, 43)
        Me.CmbSemi.Name = "CmbSemi"
        Me.CmbSemi.Size = New System.Drawing.Size(152, 21)
        Me.CmbSemi.TabIndex = 7
        Me.CmbSemi.Text = "Select"
        '
        'CheckBoxPreSemi
        '
        Me.CheckBoxPreSemi.Location = New System.Drawing.Point(16, 45)
        Me.CheckBoxPreSemi.Name = "CheckBoxPreSemi"
        Me.CheckBoxPreSemi.Size = New System.Drawing.Size(128, 16)
        Me.CheckBoxPreSemi.TabIndex = 8
        Me.CheckBoxPreSemi.Text = "Semi  (Material)"
        '
        'CmbMaterial
        '
        Me.CmbMaterial.Enabled = False
        Me.CmbMaterial.Location = New System.Drawing.Point(144, 8)
        Me.CmbMaterial.Name = "CmbMaterial"
        Me.CmbMaterial.Size = New System.Drawing.Size(152, 21)
        Me.CmbMaterial.TabIndex = 18
        Me.CmbMaterial.Text = "Select"
        '
        'CheckBoxType
        '
        Me.CheckBoxType.Location = New System.Drawing.Point(16, 10)
        Me.CheckBoxType.Name = "CheckBoxType"
        Me.CheckBoxType.Size = New System.Drawing.Size(104, 16)
        Me.CheckBoxType.TabIndex = 20
        Me.CheckBoxType.Text = "Type Material"
        '
        'CmdDel
        '
        Me.CmdDel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CmdDel.Image = CType(resources.GetObject("CmdDel.Image"), System.Drawing.Image)
        Me.CmdDel.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdDel.Location = New System.Drawing.Point(8, 592)
        Me.CmdDel.Name = "CmdDel"
        Me.CmdDel.Size = New System.Drawing.Size(80, 56)
        Me.CmdDel.TabIndex = 21
        Me.CmdDel.Text = "Delete"
        Me.CmdDel.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CheckBoxTire
        '
        Me.CheckBoxTire.Location = New System.Drawing.Point(448, 8)
        Me.CheckBoxTire.Name = "CheckBoxTire"
        Me.CheckBoxTire.Size = New System.Drawing.Size(80, 16)
        Me.CheckBoxTire.TabIndex = 23
        Me.CheckBoxTire.Text = "Green Tire"
        Me.CheckBoxTire.Visible = False
        '
        'CmbTire
        '
        Me.CmbTire.Enabled = False
        Me.CmbTire.Location = New System.Drawing.Point(528, 8)
        Me.CmbTire.Name = "CmbTire"
        Me.CmbTire.Size = New System.Drawing.Size(152, 21)
        Me.CmbTire.TabIndex = 22
        Me.CmbTire.Text = "Select"
        Me.CmbTire.Visible = False
        '
        'cmdActive
        '
        Me.cmdActive.Image = CType(resources.GetObject("cmdActive.Image"), System.Drawing.Image)
        Me.cmdActive.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdActive.Location = New System.Drawing.Point(368, 8)
        Me.cmdActive.Name = "cmdActive"
        Me.cmdActive.Size = New System.Drawing.Size(72, 56)
        Me.cmdActive.TabIndex = 24
        Me.cmdActive.Text = " Active"
        Me.cmdActive.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ChkAvtive
        '
        Me.ChkAvtive.Checked = True
        Me.ChkAvtive.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkAvtive.Location = New System.Drawing.Point(304, 45)
        Me.ChkAvtive.Name = "ChkAvtive"
        Me.ChkAvtive.Size = New System.Drawing.Size(64, 16)
        Me.ChkAvtive.TabIndex = 25
        Me.ChkAvtive.Text = " Active"
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Shortcut = System.Windows.Forms.Shortcut.F5
        Me.MenuItem1.Text = "Load"
        '
        'CmdImport
        '
        Me.CmdImport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdImport.Image = CType(resources.GetObject("CmdImport.Image"), System.Drawing.Image)
        Me.CmdImport.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdImport.Location = New System.Drawing.Point(731, 594)
        Me.CmdImport.Name = "CmdImport"
        Me.CmdImport.Size = New System.Drawing.Size(72, 56)
        Me.CmdImport.TabIndex = 26
        Me.CmdImport.Text = "Import"
        Me.CmdImport.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdExport
        '
        Me.CmdExport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdExport.Image = CType(resources.GetObject("CmdExport.Image"), System.Drawing.Image)
        Me.CmdExport.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdExport.Location = New System.Drawing.Point(803, 594)
        Me.CmdExport.Name = "CmdExport"
        Me.CmdExport.Size = New System.Drawing.Size(72, 56)
        Me.CmdExport.TabIndex = 27
        Me.CmdExport.Text = "Export"
        Me.CmdExport.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'FrmSemi
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1058, 656)
        Me.Controls.Add(Me.CmdExport)
        Me.Controls.Add(Me.CmdImport)
        Me.Controls.Add(Me.cmdActive)
        Me.Controls.Add(Me.CheckBoxTire)
        Me.Controls.Add(Me.CmbTire)
        Me.Controls.Add(Me.CmdDel)
        Me.Controls.Add(Me.CheckBoxType)
        Me.Controls.Add(Me.CmbMaterial)
        Me.Controls.Add(Me.CheckBoxPreSemi)
        Me.Controls.Add(Me.CmbSemi)
        Me.Controls.Add(Me.CmdEdit)
        Me.Controls.Add(Me.CmdClose)
        Me.Controls.Add(Me.CmdSave)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ChkAvtive)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.MainMenu1
        Me.Name = "FrmSemi"
        Me.Text = "Semi (Material)"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataGridCOM, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "CONSTANT"
    Dim DT As New DataTable
    Dim StrSQL As String
    Dim oldrow As Integer
#End Region

#Region "Form Event"
    Private Sub FrmSemi_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadSemi() 'Load data for DataGrid
        LoadPSemi() 'Load data of Table Semi
        LoadType() 'Load data of Table TypeMaterial
        LoadTire()
        If CheckBoxType.Checked = False Then
            GrdDV.RowFilter = " "
            DataGridCOM.DataSource = GrdDV
        End If

        If CheckBoxPreSemi.Checked = False Then
            GrdDV.RowFilter = " "
            DataGridCOM.DataSource = GrdDV
        End If
    End Sub
#End Region

#Region "Function_Load"
    Private Sub LoadSemi()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        Dim sb As StringBuilder = New StringBuilder()
        sb.AppendLine("SELECT Final,Semicode,Revision,MaterialCode,MaterialName,QPU,Active,MName ")
        sb.AppendLine("   ,code,Mastercode,MRev,RMcode,RmRev,Qty,Qty*(n/cntn) Qty2,Unit,per,width,Length,num,Active,cn,wn,ln,n,aa,cntn ")
        sb.AppendLine("FROM (")
        sb.AppendLine("   SELECT Final,Semicode,Revision,MaterialCode,MaterialName,QPU,Active,MaterialName MName")
        sb.AppendLine("   ,Semicode+','+Revision  code, '' Mastercode, '' MRev, '' RMcode")
        sb.AppendLine("   ,'' RmRev, null Qty,null Unit,null per,width,Length,num,cn,width wn,Length ln,num n,Active aa,cn cntn")
        sb.AppendLine("   FROM TBLSemi s")
        sb.AppendLine("   LEFT OUTER JOIN TBLTypeMaterial t ON s.MaterialType = t.MaterialCode") 'Table Semi and TypeMaterial
        sb.AppendLine("   UNION")
        sb.AppendLine("   SELECT '' Final, '' Semicode, '' Revision, '' MaterialCode, '' MaterialName,null TQty, '' Active ,MaterialName MName ,")
        sb.AppendLine("     Mastercode+','+b.Revision code,Mastercode,b.Revision Mrev,RMcode,RmRevision RmRev,Qty,Unit,per")
        sb.AppendLine("     ,null width,null Length,null num,null cn,wn,ln,n,aa,cntn ")
        sb.AppendLine("    FROM (")
        sb.AppendLine("      SELECT Final,MaterialName,semicode,Revision,Width wn,Length ln,num n,Active aa,Cn cntn ")
        sb.AppendLine("      FROM TBLSemi s")
        sb.AppendLine("      LEFT OUTER JOIN TBLTypeMaterial t ON s.MaterialType = t.MaterialCode") 'Table Semi and TypeMaterial
        sb.AppendLine("    ) a")
        sb.AppendLine("    LEFT OUTER JOIN (")
        sb.AppendLine("      SELECT *")
        sb.AppendLine("      FROM TBLMaster")
        sb.AppendLine("      WHERE MasterCode in (  SELECT Semicode FROM TBLSemi )")
        sb.AppendLine("     ) b ON a.semicode+a.Revision = b.mastercode+b.Revision")
        sb.AppendLine(") xxx ")
        sb.AppendLine("WHERE Mname in ('BF (Upper,Lower,Center)', 'TREAD') ") 'Where MaterialName is BF and TREAD
        sb.AppendLine("UNION")
        sb.AppendLine("SELECT Final,Semicode,Revision,MaterialCode,MaterialName,QPU,Active,MName ")
        sb.AppendLine("   ,code,Mastercode,MRev,RMcode,RmRev,Qty,Qty*(n/cntn)*(ln/1000) Qty2,Unit,per,width,Length,num,Active,cn,wn,ln,n,aa,cntn ")
        sb.AppendLine("FROM (")
        sb.AppendLine("  SELECT Final,Semicode,Revision,MaterialCode,MaterialName,QPU,Active,MaterialName MName")
        sb.AppendLine("     ,Semicode+','+Revision  code,'' Mastercode,'' MRev,'' RMcode")
        sb.AppendLine("     ,'' RmRev,null Qty,null Unit,null per,width,Length,num,cn,width wn,Length ln,num n,Active aa,cn cntn")
        sb.AppendLine("  FROM TBLSemi s")
        sb.AppendLine("  LEFT OUTER JOIN TBLTypeMaterial t ON s.MaterialType = t.MaterialCode")
        sb.AppendLine("  UNION")
        sb.AppendLine("  SELECT ''Final,'' Semicode,'' Revision,'' MaterialCode,'' MaterialName,null TQty,'' Active ,MaterialName MName ,")
        sb.AppendLine("     Mastercode+','+b.Revision code,Mastercode,b.Revision Mrev,RMcode,RmRevision RmRev,Qty,Unit,per")
        sb.AppendLine("     ,null width,null Length,null num,null cn,wn,ln,n,aa,cntn ")
        sb.AppendLine("  FROM (")
        sb.AppendLine("    SELECT Final,MaterialName,semicode,Revision,Width wn,Length ln,num n,Active aa,Cn cntn ")
        sb.AppendLine("    FROM TBLSemi s")
        sb.AppendLine("    LEFT OUTER JOIN TBLTypeMaterial t ON s.MaterialType = t.MaterialCode") 'Table Semi and TypeMaterial
        sb.AppendLine("   ) a")
        sb.AppendLine("   LEFT OUTER JOIN (")
        sb.AppendLine("     SELECT * ")
        sb.AppendLine("     FROM TBLMaster ")
        sb.AppendLine("     WHERE MasterCode in ( SELECT Semicode FROM TBLSemi )")
        sb.AppendLine("    ) b ON a.semicode+a.Revision = b.mastercode+b.Revision")
        sb.AppendLine(") xxx ")
        sb.AppendLine("WHERE Mname not in  ('BF (Upper,Lower,Center)','TREAD')") 'Where MaterialName is not BF and TREAD
        sb.AppendLine("ORDER BY code,semicode DESC")
        StrSQL = sb.ToString()

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
        DataGridCOM.DataSource = GrdDV
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

        With DataGridCOM
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
            .HeaderText = "Material"
            .MappingName = "MaterialName"
            .Width = 135
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle2 As New DataGridColoredLine2
        With grdColStyle2
            .HeaderText = "Part No.(Material)"
            .MappingName = "SemiCode"
            .Width = 150
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle2_1 As New DataGridColoredLine2
        With grdColStyle2_1
            .HeaderText = "Rev."
            .MappingName = "Revision"
            .Width = 55
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle2_2 As New DataGridColoredLine2
        With grdColStyle2_2
            .HeaderText = "Final"
            .MappingName = "Final"
            .Width = 75
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle3 As New DataGridColoredLine2
        With grdColStyle3
            .HeaderText = "Code(Material)"
            .MappingName = "RmCode"
            .Width = 85
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle5 As New DataGridColoredLine2
        With grdColStyle5
            .HeaderText = "Weight (g/M)"
            .MappingName = "QPU"
            .NullText = ""
            .Width = 85
            .Format = "##,###,###.00"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle6 As New DataGridColoredLine2
        With grdColStyle6
            .HeaderText = "Unit"
            .MappingName = "Unit"
            .Width = 55
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle9_0 As New DataGridColoredLine2
        With grdColStyle9_0
            .HeaderText = "Weight Formula_"
            .MappingName = "Qty2"
            .NullText = ""
            .Width = 100
            .Format = "##,###,###.00"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle9 As New DataGridColoredLine2
        With grdColStyle9
            .HeaderText = "Weight Per Meter_"
            .MappingName = "Qty"
            .NullText = ""
            .Width = 110
            .Format = "##,###,###.00"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle10 As New DataGridColoredLine2
        With grdColStyle10
            .HeaderText = "Width"
            .MappingName = "Width"
            .NullText = ""
            .Width = 65
            .Format = "##,###,###.0"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle11 As New DataGridColoredLine2
        With grdColStyle11
            .HeaderText = "Length"
            .MappingName = "Length"
            .NullText = ""
            .Width = 65
            .Format = "##,###,###.0"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle12 As New DataGridColoredLine2
        With grdColStyle12
            .HeaderText = "N."
            .MappingName = "Num"
            .NullText = ""
            .Width = 45
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle13 As New DataGridColoredLine2
        With grdColStyle13
            .HeaderText = "CNT N "
            .MappingName = "cn"
            .NullText = ""
            .Width = 45
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle14 As New DataGridColoredLine2
        With grdColStyle14
            .HeaderText = "Active"
            .MappingName = "Active"
            .NullText = ""
            .Width = 45
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle15 As New DataGridColoredLine2
        With grdColStyle15
            .HeaderText = " % WT "
            .MappingName = "Per"
            .NullText = ""
            .Width = 45
            .Format = "#0.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        grdTableStyle1.GridColumnStyles.AddRange _
    (New DataGridColumnStyle() _
    {grdColStyle14, grdColStyle13, grdColStyle2, grdColStyle2_1, grdColStyle3,
    grdColStyle9_0, grdColStyle9, grdColStyle6, grdColStyle5, grdColStyle15, grdColStyle10,
    grdColStyle12, grdColStyle11})

        DataGridCOM.TableStyles.Add(grdTableStyle1)
        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub
    Private Sub ResetTableStyle()
        ' Clear out the existing TableStyles and result default formatting.
        With DataGridCOM
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

#Region "COMBOBOX"
    Sub LoadPSemi()
        Dim dtPSemi As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        If CheckBoxType.Checked Then
            StrSQL = "   SELECT  distinct Code,MaterialName"
            StrSQL &= "   FROM  TblGroup g"
            StrSQL &= "  left outer join "
            StrSQL &= "  ("
            StrSQL &= "  SELECT  semiCode,MaterialName"
            StrSQL &= "   FROM  TblSemi p"
            StrSQL &= "  left outer join  TblTypeMaterial t"
            StrSQL &= "  on p.MaterialType = t.MaterialCode"
            StrSQL &= "  )semi"
            StrSQL &= "  on g.code = semi.semicode"
            StrSQL &= "  where Typecode = '05'"
            StrSQL &= "  and  MaterialName = '" & CmbMaterial.Text.Trim & "'"
        Else
            StrSQL = "   SELECT distinct Code,MaterialName"
            StrSQL &= "   FROM  TblGroup g"
            StrSQL &= "  left outer join "
            StrSQL &= "  ("
            StrSQL &= "  SELECT  semiCode,MaterialName"
            StrSQL &= "   FROM  TblSemi p"
            StrSQL &= "  left outer join  TblTypeMaterial t"
            StrSQL &= "  on p.MaterialType = t.MaterialCode"
            StrSQL &= "  )semi"
            StrSQL &= "  on g.code = semi.semicode"
            StrSQL &= "  where Typecode = '05'"
        End If
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtPSemi = New DataTable
            DA.Fill(dtPSemi)
        Catch
        Finally
        End Try
        dtPSemi.TableName = TBL_PreSemi
        GrdDVPreSemi = dtPSemi.DefaultView
        '************************************
        CmbSemi.DisplayMember = "Code"
        CmbSemi.ValueMember = "Code"
        CmbSemi.DataSource = dtPSemi
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadType()
        Dim dtType As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        StrSQL = "SELECT  *  FROM  TBLTypeMaterial "
        StrSQL &= "  Where  descName like 'Semi' "

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
        CmbMaterial.DisplayMember = "MaterialName"
        CmbMaterial.ValueMember = "Materialcode"
        CmbMaterial.DataSource = dtType
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadTire()
        Dim dtType As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        StrSQL = "SELECT  *  FROM  TBLTire "

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
        CmbTire.DisplayMember = "TireCode"
        CmbTire.ValueMember = "TireCode"
        CmbTire.DataSource = dtType
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
#End Region

#Region "Control Event"
    Private Sub CmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClose.Click
        Me.Close()
    End Sub

    Private Sub CmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdEdit.Click
        Dim fAddSemi As New FrmAddSemi
        fAddSemi.CmdSave.Text = "Edit"
        If GrdDV.Item(oldrow).Row("Semicode") = "" Then
            Exit Sub
        Else
            fAddSemi.TxtCode.Text = GrdDV.Item(oldrow).Row("Semicode")
            fAddSemi.TxtCode.Enabled = False
            fAddSemi.TxtRev.Text = GrdDV.Item(oldrow).Row("Revision")
            fAddSemi.iCmb = GrdDV.Item(oldrow).Row("MaterialName")
            fAddSemi.TxtNum.Text = GrdDV.Item(oldrow).Row("Num")
            fAddSemi.TxtLenght.Text = GrdDV.Item(oldrow).Row("Length")
            If GrdDV.Item(oldrow).Row("MaterialCode") = "13" Then
                fAddSemi.txtWidth.Text = ""
            ElseIf GrdDV.Item(oldrow).Row("MaterialCode") = "14" Then
                fAddSemi.txtWidth.Text = ""
            Else
                fAddSemi.txtWidth.Text = GrdDV.Item(oldrow).Row("Width")
            End If
            fAddSemi.txtType = CmbMaterial.Text
            fAddSemi.ShowDialog()
            LoadSemi()
            CheckBox()
            oldrow = 0
        End If
    End Sub

    Private Sub CmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSave.Click
        Dim tType, Tcode As String
        Dim fAddSemi As New FrmAddSemi
        tType = CmbMaterial.Text
        fAddSemi.CmdSave.Text = "Save"
        fAddSemi.txtType = CmbMaterial.Text
        fAddSemi.ShowDialog()
        Tcode = fAddSemi.TxtCode.Text
        LoadSemi()
        LoadPSemi()
        LoadType()
        LoadTire()
        CmbMaterial.Text = tType
        CmbSemi.Text = Tcode
        CheckBox()
        oldrow = 0
    End Sub

    Private Sub DataGridCOM_CurrentCellChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGridCOM.CurrentCellChanged
        oldrow = DataGridCOM.CurrentCell.RowNumber
    End Sub

    Private Sub CmdExport_Click(sender As Object, e As EventArgs) Handles CmdExport.Click
        Dim arrColumn As String() = System.Configuration.ConfigurationManager.AppSettings("EXP_EXCEL_COLUMN_MASTER_SEMI").ToString().Split(New Char() {","c})
        Dim arrColumnHeader As String() = System.Configuration.ConfigurationManager.AppSettings("EXP_EXCEL_COLUMN_HEADER_MASTER_SEMI").ToString().Split(New Char() {","c})
        ExcelLib.Export(Me, GrdDV, TBL_RM, arrColumn, arrColumnHeader)
    End Sub

    Private Sub CmdImport_Click(sender As Object, e As EventArgs) Handles CmdImport.Click
        Dim arrColumn As String() = System.Configuration.ConfigurationManager.AppSettings("IMP_EXCEL_COLUMN_MASTER_SEMI").ToString().Split(New Char() {","c})
        Dim importDialog As OpenFileDialog = New OpenFileDialog With {
            .Filter = System.Configuration.ConfigurationManager.AppSettings("DIALOG_FILE_EXT").ToString()
        }
        Dim dtRec As DataTable
        Dim sb As New System.Text.StringBuilder()
        Dim frmOverlay As New Form()
        Dim totalQty As Double = 0
        Dim QPU As Double = 0

        If importDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            'Create Importing of overlay
            Dim frm As New Importing()
            frmOverlay.StartPosition = FormStartPosition.Manual
            frmOverlay.FormBorderStyle = FormBorderStyle.None
            frmOverlay.Opacity = 0.5D
            frmOverlay.BackColor = Color.Black
            frmOverlay.WindowState = FormWindowState.Maximized
            frmOverlay.TopMost = True
            frmOverlay.Location = Me.Location
            frmOverlay.ShowInTaskbar = False
            frmOverlay.Show()
            frm.Owner = frmOverlay
            ExcelLib.CenterForm(frm, Me)
            frm.Show()

            'Read excel file
            dtRec = ExcelLib.Import(importDialog.FileName, Me, GrdDV, TBL_Semi, arrColumn)

            'Save
            If dtRec IsNot Nothing Then
                Using cnSQL As New SqlConnection(C1.Strcon)
                    cnSQL.Open()
                    Dim cmSQL As SqlCommand = cnSQL.CreateCommand()
                    Dim trans As SqlTransaction = cnSQL.BeginTransaction("RMTransaction")

                    cmSQL.Connection = cnSQL
                    cmSQL.Transaction = trans

                    Try
                        'Set datetime
                        Dim strDate As String = DateTime.Now.ToString("yyyyMMdd", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))
                        Dim iTime As String = DateTime.Now.ToString("HHmm", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))
                        'Get Master
                        Dim dtTypeMaterial As DataTable = GetTypeMaterial() 'Table TBLTypeMaterial
                        Dim dtUnitCode As DataTable = GetUnitCode() 'Table TBLUnit

                        '//Sort Data from Excel
                        dtRec.DefaultView.Sort = "TypeMaterial DESC, Semi DESC, SemiRevision DESC"
                        dtRec = dtRec.DefaultView.ToTable

                        '//**Check All Import Data that all data still in TBLCompound, TBLMaster and TBLRM
                        If ChkRMCode_TypeMaterial_Unit(dtRec, dtTypeMaterial, dtUnitCode) = False Then
                            LoadSemi() 'ReQuery and set datagrid
                            frmOverlay.Dispose()
                            Exit Sub
                        End If

                        For i As Integer = 0 To dtRec.Rows.Count - 1
                            Dim strTypeMaterial As String = dtRec.Rows(i)("TypeMaterial").ToString().Trim()
                            Dim strTypeMaterialOriginal As String = String.Empty
                            Dim strSemi As String = dtRec.Rows(i)("Semi").ToString().Trim()
                            Dim strRevision As String = dtRec.Rows(i)("SemiRevision").ToString().Trim()
                            Dim strRMCode As String = dtRec.Rows(i)("RMCode").ToString().Trim()

                            'Check empty value of Semi and Revision
                            If strSemi.Equals(String.Empty) Or strRevision.Equals(String.Empty) Then
                                Throw New ApplicationException("Semi Code and Semi Revision is not empty.")
                            End If
                            If strRevision.Length > 3 Then
                                Throw New System.Exception("Revision must less than 4 digits.")
                            End If

                            'Get Type Material Master
                            Dim arrTypeMatCode As DataRow() = dtTypeMaterial.Select("MaterialName = '" & strTypeMaterial & "'")
                            strTypeMaterial = arrTypeMatCode(0)("MaterialCode")

                            If strTypeMaterial.Length > 0 And strSemi.Length > 0 And strRevision.Length > 0 And strRMCode.Length > 0 Then
                                Dim dblQty As Double = 0
                                Dim dblWidth As Double = 0
                                Dim dblLength As Double = 0
                                Dim intN As Integer = 0
                                Dim strUnit As String = dtRec.Rows(i)("Unit").ToString().Trim()
                                Dim GridRow As DataRow()        '//Grid Data
                                Dim ExcelRow As DataRow()       '//Excel Data

                                'Get Unit Master
                                Dim arrUnitCode As DataRow() = dtUnitCode.Select("ShortUnitName = '" & strUnit & "'")
                                strUnit = arrUnitCode(0)("ShortUnitName")

                                '//Check Qty input format as Number
                                If dtRec.Rows(i)("Qty").ToString().Length > 0 Then
                                    If Not Double.TryParse(dtRec.Rows(i)("Qty"), dblQty) Then
                                        Throw New System.Exception("Please input Qty data as Number")
                                    End If
                                Else
                                    Throw New System.Exception("Please input Qty data as Number")
                                End If

                                GridRow = DT.Select("Mastercode = '" & strSemi & "' AND MRev = '" & strRevision & "'")
                                If GridRow.Count > 0 Then
                                    'Found data in data grid
                                    'Check material type from data grid
                                    strTypeMaterialOriginal = GridRow(0)("MaterialCode").ToString()

                                    '//Check Width,Length and N as Number by Condition
                                    If GridRow(0)("MaterialCode").ToString() = "13" Then
                                        'TREAD
                                        '//Just QTY
                                        Double.TryParse(dtRec.Rows(i)("Width"), dblWidth)
                                        Double.TryParse(dtRec.Rows(i)("Length"), dblLength)
                                        Integer.TryParse(dtRec.Rows(i)("N"), intN)
                                    ElseIf GridRow(0)("MaterialCode").ToString() = "14" Then
                                        'BF (Upper,Lower,Center)
                                        If dtRec.Rows(i)("N").ToString().Length > 0 Then
                                            If Not Integer.TryParse(dtRec.Rows(i)("N"), intN) Then
                                                Throw New System.Exception("Please input N data as Number")
                                            End If
                                        Else
                                            Throw New System.Exception("Please input N data as Number")
                                        End If

                                        Double.TryParse(dtRec.Rows(i)("Width"), dblWidth)
                                        Double.TryParse(dtRec.Rows(i)("Length"), dblLength)
                                    Else
                                        'Other type (CUSSION,BODY PLY, BELT1-4,WIRE CHAFER,Nylon CHAFER,SIDE,INNERLINER,FLIPPER)
                                        If dtRec.Rows(i)("Length").ToString().Length > 0 Then
                                            If Not Double.TryParse(dtRec.Rows(i)("Length"), dblLength) Then
                                                Throw New System.Exception("Please input Length data as Number")
                                            End If
                                        Else
                                            Throw New System.Exception("Please input Length data as Number")
                                        End If

                                        Double.TryParse(dtRec.Rows(i)("Width"), dblWidth)
                                        Integer.TryParse(dtRec.Rows(i)("N"), intN)
                                    End If
                                Else
                                    'Not found data in data grid
                                    'Check material type from excel
                                    '//Check Width,Length and N as Number by Condition
                                    If strTypeMaterial = "13" Then
                                        'TREAD
                                        '//Just QTY
                                        Double.TryParse(dtRec.Rows(i)("Width"), dblWidth)
                                        Double.TryParse(dtRec.Rows(i)("Length"), dblLength)
                                        Integer.TryParse(dtRec.Rows(i)("N"), intN)
                                    ElseIf strTypeMaterial = "14" Then
                                        'BF (Upper,Lower,Center)
                                        If dtRec.Rows(i)("N").ToString().Length > 0 Then
                                            If Not Integer.TryParse(dtRec.Rows(i)("N"), intN) Then
                                                Throw New System.Exception("Please input N data as Number")
                                            End If
                                        Else
                                            Throw New System.Exception("Please input N data as Number")
                                        End If

                                        Double.TryParse(dtRec.Rows(i)("Width"), dblWidth)
                                        Double.TryParse(dtRec.Rows(i)("Length"), dblLength)
                                    Else
                                        'Other type
                                        If dtRec.Rows(i)("Length").ToString().Length > 0 Then
                                            If Not Double.TryParse(dtRec.Rows(i)("Length"), dblLength) Then
                                                Throw New System.Exception("Please input Length data as Number")
                                            End If
                                        Else
                                            Throw New System.Exception("Please input Length data as Number")
                                        End If

                                        Double.TryParse(dtRec.Rows(i)("Width"), dblWidth)
                                        Integer.TryParse(dtRec.Rows(i)("N"), intN)
                                    End If
                                End If 'If GridRow.Count > 0

                                '//For Check Data from above row on import file.
                                Dim chkSameSemiBefore As String = String.Empty
                                Dim chkSameRevisionBefore As String = String.Empty

                                If i > 0 Then
                                    'Not first row
                                    chkSameSemiBefore = dtRec.Rows(i - 1)("Semi").ToString().Trim()
                                    chkSameRevisionBefore = dtRec.Rows(i - 1)("SemiRevision").ToString().Trim()
                                Else
                                    chkSameSemiBefore = String.Empty
                                    chkSameRevisionBefore = String.Empty
                                End If

                                '//Summarize TotalQty, QPU
                                If strSemi <> chkSameSemiBefore Or strRevision <> chkSameRevisionBefore Then
                                    'Start new semi and revision in each group
                                    totalQty = 0
                                    QPU = 0
                                    GridRow = DT.Select("Mastercode = '" & strSemi & "' AND MRev = '" & strRevision & "'")

                                    If GridRow.Count > 0 Then
                                        '//Have Semi and Rev on DB so can use TypeMat from DB
                                        Dim sameTypeMat As String = GridRow(0)("MaterialCode").ToString()

                                        For j As Integer = 0 To GridRow.Count - 1
                                            'Summarize QPU from DB
                                            QPU = QPU + CalculateQPU(CDbl(GridRow(j)("Qty")), intN, dblLength, sameTypeMat)

                                            'Summarize TotalQty from DB
                                            totalQty = totalQty + GridRow(j)("Qty")
                                        Next j

                                        'Use Qty from ImportExcel if it found in ImportExcel
                                        ExcelRow = dtRec.Select("Semi = '" & strSemi & "' AND SemiRevision = '" & strRevision & "'")
                                        For j As Integer = 0 To ExcelRow.Count - 1
                                            GridRow = DT.Select("Mastercode = '" & strSemi & "' AND MRev = '" & strRevision & "' AND RMCode = '" & ExcelRow(j)("RMCode") & "'")

                                            If GridRow.Count > 0 Then
                                                'Remove value from DB
                                                QPU = QPU - CalculateQPU(CDbl(GridRow(0)("Qty")), intN, dblLength, sameTypeMat)
                                                totalQty = totalQty - GridRow(0)("Qty")
                                            End If

                                            'Summarize QPU from ImportExcel
                                            QPU = QPU + CalculateQPU(CDbl(ExcelRow(j)("Qty")), intN, dblLength, sameTypeMat)

                                            'Summarize Qty from ImportExcel
                                            totalQty = totalQty + CDbl(ExcelRow(j)("Qty"))
                                        Next j
                                    Else
                                        '//New PreSemi and Rev so use TypeMat from ImportExcel
                                        ExcelRow = dtRec.Select("Semi = '" & strSemi & "' AND SemiRevision = '" & strRevision & "'")
                                        For j As Integer = 0 To ExcelRow.Count - 1
                                            'Get Type Material Master
                                            Dim arrCheckTypeMatCodeExcel As DataRow() = dtTypeMaterial.Select("MaterialName = '" & ExcelRow(j)("TypeMaterial") & "'")
                                            Dim strCheckTypeMaterialCode As String = arrCheckTypeMatCodeExcel(0)("MaterialCode")

                                            'Summarize QPU
                                            QPU = QPU + CalculateQPU(CDbl(ExcelRow(j)("Qty")), intN, dblLength, strCheckTypeMaterialCode)

                                            'Summarize QPU
                                            totalQty = totalQty + CDbl(ExcelRow(j)("Qty"))
                                        Next j
                                    End If 'If GridRow.Count > 0
                                End If 'If strSemi <> chkSamePreSemiBefore Or strRevision <> chkSameRevisionBefore

                                '// 1.) Check Semi and Revision fron DB
                                '// 1.1) [NG] Insert TBLMASTER,TBLGroup,TBLSemi and TBLConvert
                                '// 1.2) [OK] Find Import RMCode on DB
                                '// 1.2.1) [NG] Insert TBLMASTER / Update TBLSemi,TBLConvert
                                '// 1.2.2) [OK] Compare QTY from Import file and DB
                                '// 1.2.2.1) [NG] Update TBLMASTER,TBLSemi,TBLConvert

                                '// 1.) Check PreSemi and Revision fron DB
                                GridRow = DT.Select("Mastercode = '" & strSemi & "' AND MRev = '" & strRevision & "'")

                                If GridRow.Count > 0 Then '// 1.2) [OK] Find Import RMCode on DB

                                    GridRow = DT.Select("Mastercode = '" & strSemi & "' AND MRev = '" & strRevision & "' AND RMCode = '" & strRMCode & "'")

                                    If GridRow.Count > 0 Then

                                        If CDbl(GridRow(0)("Qty")) <> dblQty Then '// 1.2.2) [OK] Compare QTY from Import file and DB
                                            '// 1.2.2.1) [NG] Update TBLMASTER,TBLPreSemi,TBLConvert
                                            '//Update TBLMASTER
                                            sb.Clear()
                                            sb.AppendLine(" Update TBLMASTER")
                                            sb.AppendLine(" Set ")
                                            sb.AppendLine(" Qty = ")

                                            If strTypeMaterialOriginal = "13" Then
                                                'TREAD
                                                sb.AppendLine(" '" & dblQty & "'")
                                            ElseIf strTypeMaterialOriginal = "14" Then
                                                'BF (Upper,Lower,Center)
                                                sb.AppendLine(" '" & (dblQty / intN) & "'")
                                            Else
                                                'Other type
                                                sb.AppendLine(" '" & (dblQty / dblLength * 1000) & "'")
                                            End If

                                            sb.AppendLine(", Per = '" & (dblQty * 100 / totalQty) & "' ")
                                            sb.AppendLine(" Where MasterCode = '" & strSemi & "' AND Revision = '" & strRevision & "' AND RMCode = '" & strRMCode & "' ")

                                            sb.AppendLine(" ")

                                            '//Update TBLPreSemi
                                            sb.AppendLine(" Update TBLSemi")
                                            sb.AppendLine(" Set ")
                                            sb.AppendLine(" QPU = '" & QPU & "'")
                                            sb.AppendLine(" Where SemiCode = '" & strSemi & "' AND Revision = '" & strRevision & "' ")

                                            sb.AppendLine(" ")

                                            '//Update TBLConvert
                                            sb.AppendLine(" Update TBLConvert")
                                            sb.AppendLine(" Set ")
                                            sb.AppendLine(" SQty = '" & QPU & "'")
                                            sb.AppendLine(" Where Code = '" & strSemi & "' AND Rev = '" & strRevision & "' ")
                                            If strTypeMaterialOriginal = "13" Then
                                                'TREAD
                                                sb.AppendLine(" AND UnitBig = 'UT'")
                                            ElseIf strTypeMaterialOriginal = "14" Then
                                                'BF (Upper,Lower,Center)
                                                sb.AppendLine(" AND UnitBig = 'UT'")
                                            Else
                                                'Other type
                                                sb.AppendLine(" AND UnitBig = 'M'")
                                            End If

                                            StrSQL = sb.ToString()
                                            cmSQL.CommandText = StrSQL
                                            cmSQL.ExecuteNonQuery()

                                            'Update Per in above query
                                            '//Update All Per in TBLMASTER***********
                                            'sb.Clear()
                                            'sb.AppendLine(" Update TBLMASTER")
                                            'sb.AppendLine(" Set ")
                                            'sb.AppendLine(" Per = Qty * 100 / " & totalQty)
                                            'sb.AppendLine(" Where MasterCode = '" & strSemi & "' AND Revision = '" & strRevision & "' ")
                                            'StrSQL = sb.ToString()
                                            'cmSQL.CommandText = StrSQL
                                            'cmSQL.ExecuteNonQuery()
                                        End If 'If CDbl(GridRow(0)("Qty")) <> dblQty

                                    Else '// 1.2.1) [NG] Insert TBLMASTER / Update TBLPreSemi,TBLConvert

                                        '//Insert TBLMASTER
                                        sb.Clear()
                                        sb.AppendLine(" Insert  TBLMASTER ")
                                        sb.AppendLine(" Values (")
                                        sb.AppendLine(" '" & strSemi & "', ")                                   'Column MasterCode
                                        sb.AppendLine(" '" & strRevision & "' , ")                              'Column Revision
                                        sb.AppendLine(" '" & strRMCode & "' , ")                                'Column RMCode
                                        sb.AppendLine(" NULL , ")                                               'Column RmRevision

                                        If strTypeMaterialOriginal = "13" Then                                  'Column Qty
                                            'TREAD
                                            sb.AppendLine(" '" & dblQty & "',")
                                        ElseIf strTypeMaterialOriginal = "14" Then
                                            'BF (Upper,Lower,Center)
                                            sb.AppendLine(" '" & (dblQty / intN) & "',")
                                        Else
                                            'Other type
                                            sb.AppendLine(" '" & ((dblQty / dblLength) * 1000) & "',")
                                        End If

                                        sb.AppendLine(" '" & strUnit & "' , ")                                  'Column Unit
                                        sb.AppendLine(" '" & (dblQty * 100 / totalQty) & "'")                 'Column Per
                                        sb.AppendLine(" )")

                                        sb.AppendLine(" ")

                                        '//Update TBLPreSemi
                                        sb.AppendLine(" Update TBLSemi")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" QPU = '" & QPU & "'")
                                        sb.AppendLine(" Where semiCode = '" & strSemi & "' AND Revision = '" & strRevision & "' ")

                                        sb.AppendLine(" ")

                                        '//Update TBLConvert
                                        sb.AppendLine(" Update TBLConvert")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" SQty = '" & totalQty & "'")
                                        sb.AppendLine(" Where Code = '" & strSemi & "' AND Rev = '" & strRevision & "' ")
                                        If strTypeMaterialOriginal = "13" Then
                                            'TREAD
                                            sb.AppendLine(" AND UnitBig = 'UT'")
                                        ElseIf strTypeMaterialOriginal = "14" Then
                                            'BF (Upper,Lower,Center)
                                            sb.AppendLine(" AND UnitBig = 'UT'")
                                        Else
                                            'Other type
                                            sb.AppendLine(" AND UnitBig = 'M'")
                                        End If

                                        StrSQL = sb.ToString()
                                        cmSQL.CommandText = StrSQL
                                        cmSQL.ExecuteNonQuery()

                                        'New record has already inserted and Other record has already updated
                                        '//Update All Per in TBLMASTER***********
                                        'sb.Clear()
                                        'sb.AppendLine(" Update TBLMASTER")
                                        'sb.AppendLine(" Set ")
                                        'sb.AppendLine(" Per = Qty * 100 / " & totalQty)
                                        'sb.AppendLine(" Where MasterCode = '" & strSemi & "' AND Revision = '" & strRevision & "' ")
                                        'StrSQL = sb.ToString()
                                        'cmSQL.CommandText = StrSQL
                                        'cmSQL.ExecuteNonQuery()

                                    End If

                                Else '// 1.1) [NG] Insert TBLMASTER,TBLGroup,TBLPreSemi and TBLConvert

                                    sb.Clear()

                                    If strSemi <> chkSameSemiBefore Or strRevision <> chkSameRevisionBefore Then
                                        'Semi and Revision is new group
                                        '//Insert TBLGroup
                                        sb.AppendLine(" Insert  TBLGroup ")
                                        sb.AppendLine(" Values (")
                                        sb.AppendLine("'05' , ")                'Column TypeCode
                                        sb.AppendLine("'" & strSemi & "'")      'Column Code
                                        sb.AppendLine(" )")

                                        sb.AppendLine(" ")

                                        '//Insert TBLSemi
                                        sb.AppendLine(" Insert TBLSemi ")
                                        sb.AppendLine(" Values (")
                                        sb.AppendLine(" '" & strSemi & "', ")           'Column Final
                                        sb.AppendLine(" '" & strSemi & "', ")           'Column SemiCode
                                        sb.AppendLine(" '" & strRevision & "', ")       'Column Revision
                                        sb.AppendLine(" '" & strTypeMaterial & "', ")   'Column MaterialType
                                        sb.AppendLine(" '" & QPU & "', ")               'Column QPU

                                        If dblWidth = 0 Then                            'Column Width
                                            sb.AppendLine(" NULL , ")
                                        Else
                                            sb.AppendLine(" '" & dblWidth & "' , ")
                                        End If

                                        If dblLength = 0 Then                           'Column Length
                                            sb.AppendLine(" NULL , ")
                                        Else
                                            sb.AppendLine(" '" & dblLength & "' , ")
                                        End If

                                        If intN = 0 Then                                'Column N
                                            sb.AppendLine(" NULL , ")
                                        Else
                                            sb.AppendLine(" '" & intN & "', ")
                                        End If

                                        sb.AppendLine(" '0' , ")                        'Column Active
                                        sb.AppendLine(" '" & strDate & "' , ")          'Column DateUp

                                        If strTypeMaterial = "13" Then                  'Column CN
                                            'TREAD
                                            sb.AppendLine(" '1' ")
                                        ElseIf strTypeMaterial = "14" Then
                                            'BF (Upper,Lower,Center)
                                            sb.AppendLine(" '1' ")
                                        ElseIf strTypeMaterial = "07" Then
                                            'BELT-3
                                            sb.AppendLine(" '1' ")
                                        ElseIf strTypeMaterial = "08" Then
                                            'BELT-4
                                            sb.AppendLine(" '1' ")
                                        ElseIf strTypeMaterial = "12" Then
                                            'INNERLINER
                                            sb.AppendLine(" '1' ")
                                        ElseIf strTypeMaterial = "04" Then
                                            'BODY PLY
                                            sb.AppendLine(" '1' ")
                                        Else
                                            'CUSSION,BELT-1,BELT-2,WIRE CHAFER,Nylon CHAFER,SIDE
                                            sb.AppendLine(" '2' ")
                                        End If

                                        sb.AppendLine(" )")

                                        sb.AppendLine(" ")

                                        '//Insert TblConvert #1
                                        sb.AppendLine(" Insert  TblConvert ")
                                        sb.AppendLine(" Values (")
                                        sb.AppendLine(" '05' , ")                   'Column Type
                                        sb.AppendLine(" '" & strSemi & "', ")       'Column Final
                                        sb.AppendLine(" '" & strSemi & "', ")       'Column Code
                                        sb.AppendLine(" '" & strRevision & "' , ")  'Column Rev
                                        sb.AppendLine(" 'KG' , ")                   'Column UnitBig
                                        sb.AppendLine(" 'KG' , ")                   'Column UnitSmall
                                        sb.AppendLine(" '1' , ")                    'Column BQty
                                        sb.AppendLine(" '1' ")                      'Column SQty
                                        sb.AppendLine(" )")

                                        '//Insert TblConvert #2
                                        sb.AppendLine(" ")

                                        sb.AppendLine(" Insert  TblConvert ")
                                        sb.AppendLine(" Values (")
                                        sb.AppendLine(" '05' , ")                       'Column Type
                                        sb.AppendLine(" '" & strSemi & "', ")           'Column Final
                                        sb.AppendLine(" '" & strSemi & "', ")           'Column Code
                                        sb.AppendLine(" '" & strRevision & "' , ")      'Column Rev

                                        If strTypeMaterial = "13" Then                  'Column UnitBig
                                            'TREAD
                                            sb.AppendLine(" 'UT' , ")
                                        ElseIf strTypeMaterial = "14" Then
                                            'BF (Upper,Lower,Center)
                                            sb.AppendLine(" 'UT' , ")
                                        Else
                                            sb.AppendLine(" 'M' , ")
                                        End If

                                        sb.AppendLine(" 'KG' , ")                       'Column UnitSmall
                                        sb.AppendLine(" '1' , ")                        'Column BQty
                                        sb.AppendLine(" '" & (QPU / 1000) & "' ")  'Column SQty

                                        sb.AppendLine(" )")
                                    End If 'If strSemi <> chkSameSemiBefore Or strRevision <> chkSameRevisionBefore

                                    '//Insert TblMaster
                                    sb.AppendLine(" Insert  TBLMASTER ")
                                    sb.AppendLine(" Values (")
                                    sb.AppendLine(" '" & strSemi & "', ")                                   'Column MasterCode
                                    sb.AppendLine(" '" & strRevision & "' , ")                              'Column Revision
                                    sb.AppendLine(" '" & strRMCode & "' , ")                                'Column RMCode
                                    sb.AppendLine(" NULL , ")                                               'Column RmRevision

                                    If strTypeMaterial = "13" Then                                          'Column Qty
                                        'TREAD
                                        sb.AppendLine(" '" & dblQty & "', ")
                                    ElseIf strTypeMaterial = "14" Then
                                        'BF (Upper,Lower,Center)
                                        sb.AppendLine(" '" & (dblQty / intN) & "', ")
                                    Else
                                        'Other type
                                        sb.AppendLine(" '" & ((dblQty / dblLength) * 1000) & "', ")
                                    End If

                                    sb.AppendLine(" '" & strUnit & "' , ")                                  'Column Unit
                                    sb.AppendLine(" '" & (dblQty * 100 / totalQty) & "'")                 'Column Per
                                    sb.AppendLine(" )")

                                    StrSQL = sb.ToString()
                                    cmSQL.CommandText = StrSQL
                                    cmSQL.ExecuteNonQuery()

                                End If 'If GridRow.Count > 0

                            End If 'If strTypeMaterial.Length > 0 And strSemi.Length > 0 And strRevision.Length > 0 And strRMCode.Length > 0
                        Next i

                        trans.Commit()
                        MessageBox.Show("Import complete", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Catch ex As SqlException
                        MsgBox("Import error" & vbCrLf & ex.Message, MsgBoxStyle.Critical, "SQL Error")
                        trans.Rollback()
                    Catch ex As Exception
                        MsgBox("Import error" & vbCrLf & ex.Message, MsgBoxStyle.Critical, "General Error")
                        trans.Rollback()
                    Finally
                        trans.Dispose()
                        cmSQL.Dispose()
                        cnSQL.Close()
                        cnSQL.Dispose()
                    End Try
                End Using 'Using cnSQL
            End If 'If dtRec IsNot Nothing Then

            LoadSemi() 'ReQuery and set datagrid
            frmOverlay.Dispose()
        End If 'If importDialog.ShowDialog() = Windows.Forms.DialogResult.OK
    End Sub
#End Region

#Region "Import"
    Private Function ChkRMCode_TypeMaterial_Unit(ByVal ImportTable As DataTable, dtTypeMaterial As DataTable, dtUnitCode As DataTable) As Boolean
        Dim cnSQLRM As SqlConnection
        Dim cmSQLRM As SqlCommand
        Dim strSQL As String = String.Empty
        Dim ret As Boolean = False
        Dim strRmcodeBefore As String = String.Empty
        Dim distinctImportTabale As New DataTable()
        Dim sb As New System.Text.StringBuilder()

        Try
            For x As Integer = 0 To ImportTable.Rows.Count - 1
                Dim strTypeMat As String = ImportTable.Rows(x)("TypeMaterial").ToString().Trim()
                Dim rmCode As String = ImportTable.Rows(x)("RMCode").ToString().Trim()
                Dim strUnit As String = ImportTable.Rows(x)("Unit").ToString().Trim()

                cnSQLRM = New SqlConnection(C1.Strcon)
                cnSQLRM.Open()

                '// 1.) Check RMCode
                If rmCode.Length > 0 Then
                    sb.Clear()
                    sb.AppendLine(" SELECT count(*)  ")
                    sb.AppendLine(" FROM ( ")
                    sb.AppendLine("   SELECT t.Typecode ,TypeName,code  ")
                    sb.AppendLine("   FROM TBLType t ")
                    sb.AppendLine("   LEFT OUTER JOIN TBLGroup  g   On t.typecode=g.typecode  ")
                    sb.AppendLine(" ) a ")
                    sb.AppendLine(" LEFT OUTER JOIN ( ")
                    sb.AppendLine("   SELECT distinct  Finalcompound code ,null DescName,compcode,0.00 Qty ")
                    sb.AppendLine("   FROM TBLCompound  ")
                    sb.AppendLine("   WHERE Compcode Not In ( SELECT code FROM TblGroup WHERE Typecode ='01') AND active = 1 ")
                    sb.AppendLine("   UNION ")
                    sb.AppendLine("   SELECT distinct  Final code ,null DescName,psemicode,0.00 Qty    ")
                    sb.AppendLine("   FROM TBLPreSemi  ")
                    sb.AppendLine("   WHERE MaterialType   Not In ( '01','02')  AND active = 1")
                    sb.AppendLine("   UNION ")
                    sb.AppendLine("   SELECT rmCode,DescName,RMcode, 0.00 Qty  ")
                    sb.AppendLine("   FROM TblRM  ")
                    sb.AppendLine("   WHERE descName Like '%Steel%' OR descName like '%Bead%' OR descName like '%Nylon%' ")
                    sb.AppendLine("   OR DescName LIKE '%Flipper%'") 'Support material type "Flipper"
                    sb.AppendLine(" ) b On a.code = b.compcode   ")
                    sb.AppendLine(" WHERE B.code Is Not null AND B.code = '" & rmCode & "' ")
                    strSQL = sb.ToString()

                    cmSQLRM = New SqlCommand(strSQL, cnSQLRM)
                    Dim i As Long = cmSQLRM.ExecuteScalar()
                    If i = 0 Then
                        cmSQLRM.Dispose()
                        Throw New System.Exception("This RM Code '" & rmCode & "' is not found from Master")
                    Else
                        cmSQLRM.Dispose()
                    End If
                Else
                    'Empty
                    ret = False
                    Throw New System.Exception("RM Code is not empty.")
                End If


                '// 2.) Check TypeMaterial
                If strTypeMat.Length = 0 Then
                    'Empty
                    Throw New System.Exception("Type Material is not empty.")
                End If

                Dim arrTypeMatCode As DataRow() = dtTypeMaterial.Select("MaterialName = '" & strTypeMat & "'")
                If arrTypeMatCode.Length = 0 Then
                    Throw New System.Exception("This TypeMaterial Code '" & strTypeMat & "' is not match Semi TypeMaterial")
                End If


                '// 3.) Check Unit
                If strUnit.Length = 0 Then
                    'Empty
                    Throw New System.Exception("Unit is not empty.")
                End If

                Dim arrUnitCode As DataRow() = dtUnitCode.Select("ShortUnitName = '" & strUnit & "'")
                If arrUnitCode.Length = 0 Then
                    Throw New System.Exception("This Unit Code '" & strUnit & "' is not match Unit Master")
                End If

                cnSQLRM.Close()
                cnSQLRM.Dispose()
            Next x

            ret = True
        Catch Exp As SqlException
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "SQL Error")
        Catch Exp As Exception
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "General Error")
        End Try

        Return ret
    End Function

    Private Function GetTypeMaterial() As DataTable
        Dim daSQL As SqlDataAdapter
        Dim strSQL As String = String.Empty
        Dim dt As New DataTable()
        Dim sb As New System.Text.StringBuilder()

        Try
            sb.AppendLine(" SELECT MaterialCode, MaterialName ")
            sb.AppendLine(" FROM TBLTypeMaterial ")
            sb.AppendLine(" WHERE descName like '%Semi%' ")
            strSQL = sb.ToString()
            daSQL = New SqlDataAdapter(strSQL, C1.Strcon)
            daSQL.Fill(dt)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "General Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return dt
    End Function

    Private Function GetUnitCode() As DataTable
        Dim daSQL As SqlDataAdapter
        Dim strSQL As String = String.Empty
        Dim dt As New DataTable()
        Dim sb As New System.Text.StringBuilder()

        Try
            sb.AppendLine(" SELECT UnitCode, ShortUnitName ")
            sb.AppendLine(" FROM TBLUnit ")
            strSQL = sb.ToString()
            daSQL = New SqlDataAdapter(strSQL, C1.Strcon)
            daSQL.Fill(dt)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "General Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return dt
    End Function

    Private Function CalculateQPU(qty As Double, n As Integer, length As Double, typeMaterial As String) As Double
        Dim QPU As Double = 0

        If typeMaterial = "13" Then
            'TREAD
            QPU = qty
        ElseIf typeMaterial = "14" Then
            'BF (Upper,Lower,Center)
            QPU = qty / n
        Else
            'Other type
            QPU = (qty / length) * 1000
        End If

        Return QPU
    End Function
#End Region

#Region "SelectData"
    Private Sub CmbSemi_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbSemi.SelectedIndexChanged
        CheckBox()
    End Sub

    Private Sub CheckBoxPreSemi_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxPreSemi.CheckedChanged
        CheckBox()
    End Sub

    Private Sub CheckBoxType_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxType.CheckedChanged
        CheckBox()
        If CheckBoxType.Checked = True Then
            LoadPSemi()
            CmbMaterial.Enabled = True
          Else
            CmbMaterial.Enabled = False
        End If
    End Sub

    Sub CheckBox()
        If CheckBoxPreSemi.Checked = True And CheckBoxType.Checked = True And CheckBoxTire.Checked = True And CheckBoxTire.Checked = False Then
            GrdDV.RowFilter = " MName like'%" & CmbMaterial.Text.Trim & "%'" _
                              & " and Code like'%" & CmbSemi.Text.Trim & "%'" _
                              & " and TCode like'%" & CmbTire.Text.Trim & "%'"
            DataGridCOM.DataSource = GrdDV
        ElseIf CheckBoxPreSemi.Checked = False And CheckBoxType.Checked = True And CheckBoxTire.Checked = False And ChkAvtive.Checked = False Then
            GrdDV.RowFilter = " MName like'%" & CmbMaterial.Text.Trim & "%'"
            DataGridCOM.DataSource = GrdDV
        ElseIf CheckBoxPreSemi.Checked = True And CheckBoxType.Checked = False And CheckBoxTire.Checked = False And ChkAvtive.Checked = False Then
            GrdDV.RowFilter = " Code like'%" & CmbSemi.Text.Trim & "%'"
            DataGridCOM.DataSource = GrdDV
        ElseIf CheckBoxPreSemi.Checked = False And CheckBoxType.Checked = True And CheckBoxTire.Checked = True And ChkAvtive.Checked = False Then
            GrdDV.RowFilter = " MName like'%" & CmbMaterial.Text.Trim & "%'" _
                              & " and TCode like'%" & CmbTire.Text.Trim & "%'"
            DataGridCOM.DataSource = GrdDV
        ElseIf CheckBoxPreSemi.Checked = True And CheckBoxType.Checked = True And CheckBoxTire.Checked = False And ChkAvtive.Checked = False Then
            GrdDV.RowFilter = " MName like'%" & CmbMaterial.Text.Trim & "%'" _
                              & " and Code like'%" & CmbSemi.Text.Trim & "%'"
            DataGridCOM.DataSource = GrdDV
        ElseIf CheckBoxPreSemi.Checked = True And CheckBoxType.Checked = False And CheckBoxTire.Checked = True And ChkAvtive.Checked = False Then
            GrdDV.RowFilter = " Code like'%" & CmbSemi.Text.Trim & "%'" _
                              & " and TCode like'%" & CmbTire.Text.Trim & "%'"
            DataGridCOM.DataSource = GrdDV
        ElseIf CheckBoxPreSemi.Checked = False And CheckBoxType.Checked = False And CheckBoxTire.Checked = True And ChkAvtive.Checked = False Then
            GrdDV.RowFilter = " TCode like'%" & CmbTire.Text.Trim & "%'"
            DataGridCOM.DataSource = GrdDV
            'edit
        ElseIf CheckBoxPreSemi.Checked = True And CheckBoxType.Checked = True And CheckBoxTire.Checked = True And ChkAvtive.Checked = True Then
            GrdDV.RowFilter = " MName like'%" & CmbMaterial.Text.Trim & "%'" _
                              & " and Code like'%" & CmbSemi.Text.Trim & "%'" _
                              & " and TCode like'%" & CmbTire.Text.Trim & "%'" _
                              & " and aa = '1'"
            DataGridCOM.DataSource = GrdDV
        ElseIf CheckBoxPreSemi.Checked = False And CheckBoxType.Checked = True And CheckBoxTire.Checked = False And ChkAvtive.Checked = True Then
            GrdDV.RowFilter = " MName like'%" & CmbMaterial.Text.Trim & "%'" _
                              & " and aa = '1'"
            DataGridCOM.DataSource = GrdDV
        ElseIf CheckBoxPreSemi.Checked = True And CheckBoxType.Checked = False And CheckBoxTire.Checked = False And ChkAvtive.Checked = True Then
            GrdDV.RowFilter = " Code like'%" & CmbSemi.Text.Trim & "%'" _
                              & " and aa = '1'"
            DataGridCOM.DataSource = GrdDV
        ElseIf CheckBoxPreSemi.Checked = False And CheckBoxType.Checked = True And CheckBoxTire.Checked = True And ChkAvtive.Checked = True Then
            GrdDV.RowFilter = " MName like'%" & CmbMaterial.Text.Trim & "%'" _
                              & " and TCode like'%" & CmbTire.Text.Trim & "%'" _
                              & " and aa = '1'"
            DataGridCOM.DataSource = GrdDV
        ElseIf CheckBoxPreSemi.Checked = True And CheckBoxType.Checked = True And CheckBoxTire.Checked = False And ChkAvtive.Checked = True Then
            GrdDV.RowFilter = " MName like'%" & CmbMaterial.Text.Trim & "%'" _
                              & " and Code like'%" & CmbSemi.Text.Trim & "%'" _
                              & " and aa = '1'"
            DataGridCOM.DataSource = GrdDV
        ElseIf CheckBoxPreSemi.Checked = True And CheckBoxType.Checked = False And CheckBoxTire.Checked = True And ChkAvtive.Checked = True Then
            GrdDV.RowFilter = " Code like'%" & CmbSemi.Text.Trim & "%'" _
                              & " and TCode like'%" & CmbTire.Text.Trim & "%'" _
                              & " and aa = '1'"
            DataGridCOM.DataSource = GrdDV
        ElseIf CheckBoxPreSemi.Checked = False And CheckBoxType.Checked = False And CheckBoxTire.Checked = True And ChkAvtive.Checked = True Then
            GrdDV.RowFilter = " TCode like'%" & CmbTire.Text.Trim & "%'" _
                              & " and aa = '1'"
            DataGridCOM.DataSource = GrdDV
        ElseIf CheckBoxPreSemi.Checked = False And CheckBoxType.Checked = False And CheckBoxTire.Checked = False And ChkAvtive.Checked = True Then
            GrdDV.RowFilter = " aa = '1'"
            DataGridCOM.DataSource = GrdDV

        Else : CheckBoxPreSemi.Checked = False And CheckBoxType.Checked = False And CheckBoxTire.Checked = False And ChkAvtive.Checked = False
            GrdDV.RowFilter = " "
            DataGridCOM.DataSource = GrdDV
        End If

         
    End Sub

    Private Sub CmbMaterial_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbMaterial.SelectedIndexChanged
        CheckBox()
        If CheckBoxType.Checked = True Then
            LoadPSemi()
            CmbMaterial.Enabled = True
        Else
            CmbMaterial.Enabled = False
        End If
    End Sub
#End Region

    Private Sub CmdDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdDel.Click

        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult

        msg = "Delete Semi(Material) : " & GrdDV.Item(oldrow).Row("semicode")  'Define message.
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Semi(Material)"   ' Define title.
        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            If ChkData() Then
                DelSemi()
                LoadSemi()
            Else
                MsgBox("Can't Delete. Please check Usage.", MsgBoxStyle.OKOnly, "Tire")
            End If
        Else
            Exit Sub
        End If
        CheckBox()
    End Sub

#Region "DelSemi"
    Private Function ChkData() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Try
            strSQL &= " select count(*) from TblSemi "
            strSQL &= " where semiCode  = '" & GrdDV.Item(oldrow).Row("SemiCode") & "'"
            strSQL &= " and Revision  = '" & GrdDV.Item(oldrow).Row("Revision") & "'"
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
    Private Function ChkDel() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Try
            strSQL &= " select count(*) from TblSemi "
            strSQL &= " where semiCode  = '" & GrdDV.Item(oldrow).Row("SemiCode") & "'"
            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar
            If i <> 0 Then
                ChkDel = True
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
    Sub DelSemi()
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Try
            strSQL = " Delete TblSemi"
            strSQL &= " where SemiCode = '" & GrdDV.Item(oldrow).Row("SemiCode") & "'"
            strSQL &= " and Revision = '" & GrdDV.Item(oldrow).Row("Revision") & "'"
            strSQL &= "  "
            strSQL &= " Delete TblMaster"
            strSQL &= " where Mastercode = '" & GrdDV.Item(oldrow).Row("SemiCode") & "'"
            strSQL &= " and Revision = '" & GrdDV.Item(oldrow).Row("Revision") & "'"
            strSQL &= "  "
            If ChkDel() Then
                strSQL &= " Delete TblGroup"
                strSQL &= " where Code = '" & GrdDV.Item(oldrow).Row("SemiCode") & "'"
            Else
                'Nothing
            End If
            strSQL &= "  "
            strSQL &= " Delete Tblconvert"
            strSQL &= " where code = '" & GrdDV.Item(oldrow).Row("SemiCode") & "'"
            strSQL &= " and Rev = '" & GrdDV.Item(oldrow).Row("Revision") & "'"
            If GrdDV.Item(oldrow).Row("MaterialCode") = "13" Then
                'TREAD
                'strSQL &= " and unitBig = 'LI'"
                strSQL &= " and unitBig = 'UT'"
            ElseIf GrdDV.Item(oldrow).Row("MaterialCode") = "14" Then
                'BF (Upper,Lower,Center)
                'strSQL &= " and unitBig = 'LI'"
                strSQL &= " and unitBig = 'UT'"
            Else
                strSQL &= " and unitBig = 'M'"
            End If

            strSQL &= "  "
            strSQL &= " Delete Tblconvert"
            strSQL &= " where code = '" & GrdDV.Item(oldrow).Row("SemiCode") & "'"
            strSQL &= " and Rev = '" & GrdDV.Item(oldrow).Row("Revision") & "'"
            strSQL &= " and unitBig = 'KG'"


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
#End Region

    Private Sub CheckBoxTire_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxTire.CheckedChanged
        CheckBox()
        If CheckBoxTire.Checked = True Then
            CmbTire.Enabled = True
        Else
            CmbTire.Enabled = False
        End If
    End Sub

    Private Sub CmbTire_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbTire.SelectedIndexChanged
        CheckBox()
    End Sub

#Region " Set Active"
    Private Sub cmdActive_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdActive.Click
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult

        If GrdDV.Count = 0 Then
            MessageBox.Show("Not found data which is actived. Please display data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If GrdDV.Item(oldrow).Row("semicode") = "" Then
            Exit Sub
        End If

        msg = "Change Active Semi(Material) : " & GrdDV.Item(oldrow).Row("semicode") _
        & "  Revision :" & GrdDV.Item(oldrow).Row("Revision") 'Define message.
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Semi(Material)"   ' Define title.
        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            If ChkData() Then
                UPSemi()
                LoadSemi()
                CheckBox()
                oldrow = 0
            Else
                MsgBox("Can't Delete. Please check Usage.", MsgBoxStyle.OkOnly, "PreSemi")
            End If
        Else
            Exit Sub
        End If
    End Sub
    Sub UPSemi()
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Try
            strSQL = " Update TblSemi"
            strSQL &= " set Active = 0"
            strSQL &= " where SemiCode = '" & GrdDV.Item(oldrow).Row("semicode") & "'"
            strSQL &= " "
            strSQL &= " Update TblSemi"
            strSQL &= " set Active = 1"
            strSQL &= " where SemiCode = '" & GrdDV.Item(oldrow).Row("semicode") & "'"
            strSQL &= " and Revision = '" & GrdDV.Item(oldrow).Row("Revision") & "'"

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

#End Region

    Private Sub ChkAvtive_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkAvtive.CheckedChanged
        CheckBox()
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        LoadSemi()
        CheckBox()
    End Sub
End Class
