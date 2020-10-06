#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Inventory_Record.Common
Imports Inventory_Record.FrmMain
#End Region

Public Class FrmPreSemi

#Region "Declare"
    Inherits System.Windows.Forms.Form
    Dim GrdDV As New DataView
    Protected Const TBL_RM As String = "TBL_RM"
    Dim GrdDVPreSemi As New DataView
    Protected Const TBL_PreSemi As String = "TBL_PreSemi"
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
    Friend WithEvents CmbPreSemi As System.Windows.Forms.ComboBox
    Friend WithEvents CheckBoxPreSemi As System.Windows.Forms.CheckBox
    Friend WithEvents CmbMaterial As System.Windows.Forms.ComboBox
    Friend WithEvents CheckBoxType As System.Windows.Forms.CheckBox
    Friend WithEvents cmdDel As System.Windows.Forms.Button
    Friend WithEvents DataGridCom As System.Windows.Forms.DataGrid
    Friend WithEvents ChkAvtive As System.Windows.Forms.CheckBox
    Friend WithEvents CmdActive As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmPreSemi))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.DataGridCom = New System.Windows.Forms.DataGrid()
        Me.CmdSave = New System.Windows.Forms.Button()
        Me.CmdClose = New System.Windows.Forms.Button()
        Me.CmdEdit = New System.Windows.Forms.Button()
        Me.CmbPreSemi = New System.Windows.Forms.ComboBox()
        Me.CheckBoxPreSemi = New System.Windows.Forms.CheckBox()
        Me.CmbMaterial = New System.Windows.Forms.ComboBox()
        Me.CheckBoxType = New System.Windows.Forms.CheckBox()
        Me.cmdDel = New System.Windows.Forms.Button()
        Me.ChkAvtive = New System.Windows.Forms.CheckBox()
        Me.CmdActive = New System.Windows.Forms.Button()
        Me.CmdImport = New System.Windows.Forms.Button()
        Me.CmdExport = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridCom, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.DataGridCom)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(8, 72)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1146, 504)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'DataGridCom
        '
        Me.DataGridCom.DataMember = ""
        Me.DataGridCom.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridCom.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridCom.Location = New System.Drawing.Point(3, 16)
        Me.DataGridCom.Name = "DataGridCom"
        Me.DataGridCom.Size = New System.Drawing.Size(1140, 485)
        Me.DataGridCom.TabIndex = 0
        '
        'CmdSave
        '
        Me.CmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdSave.Image = CType(resources.GetObject("CmdSave.Image"), System.Drawing.Image)
        Me.CmdSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdSave.Location = New System.Drawing.Point(1008, 578)
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
        Me.CmdClose.Location = New System.Drawing.Point(1080, 578)
        Me.CmdClose.Name = "CmdClose"
        Me.CmdClose.Size = New System.Drawing.Size(72, 56)
        Me.CmdClose.TabIndex = 2
        Me.CmdClose.Text = "Close"
        Me.CmdClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdEdit
        '
        Me.CmdEdit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdEdit.Enabled = False
        Me.CmdEdit.Image = CType(resources.GetObject("CmdEdit.Image"), System.Drawing.Image)
        Me.CmdEdit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdEdit.Location = New System.Drawing.Point(936, 578)
        Me.CmdEdit.Name = "CmdEdit"
        Me.CmdEdit.Size = New System.Drawing.Size(72, 56)
        Me.CmdEdit.TabIndex = 3
        Me.CmdEdit.Text = "Edit"
        Me.CmdEdit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmbPreSemi
        '
        Me.CmbPreSemi.Enabled = False
        Me.CmbPreSemi.Location = New System.Drawing.Point(144, 43)
        Me.CmbPreSemi.Name = "CmbPreSemi"
        Me.CmbPreSemi.Size = New System.Drawing.Size(152, 21)
        Me.CmbPreSemi.TabIndex = 7
        Me.CmbPreSemi.Text = "Select"
        '
        'CheckBoxPreSemi
        '
        Me.CheckBoxPreSemi.Location = New System.Drawing.Point(16, 45)
        Me.CheckBoxPreSemi.Name = "CheckBoxPreSemi"
        Me.CheckBoxPreSemi.Size = New System.Drawing.Size(128, 16)
        Me.CheckBoxPreSemi.TabIndex = 8
        Me.CheckBoxPreSemi.Text = "Pre Semi  (Material)"
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
        'cmdDel
        '
        Me.cmdDel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdDel.Image = CType(resources.GetObject("cmdDel.Image"), System.Drawing.Image)
        Me.cmdDel.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdDel.Location = New System.Drawing.Point(8, 576)
        Me.cmdDel.Name = "cmdDel"
        Me.cmdDel.Size = New System.Drawing.Size(80, 56)
        Me.cmdDel.TabIndex = 21
        Me.cmdDel.Text = "Delete"
        Me.cmdDel.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ChkAvtive
        '
        Me.ChkAvtive.Checked = True
        Me.ChkAvtive.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkAvtive.Location = New System.Drawing.Point(304, 45)
        Me.ChkAvtive.Name = "ChkAvtive"
        Me.ChkAvtive.Size = New System.Drawing.Size(64, 16)
        Me.ChkAvtive.TabIndex = 22
        Me.ChkAvtive.Text = " Active"
        '
        'CmdActive
        '
        Me.CmdActive.Image = CType(resources.GetObject("CmdActive.Image"), System.Drawing.Image)
        Me.CmdActive.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdActive.Location = New System.Drawing.Point(376, 8)
        Me.CmdActive.Name = "CmdActive"
        Me.CmdActive.Size = New System.Drawing.Size(80, 56)
        Me.CmdActive.TabIndex = 23
        Me.CmdActive.Text = " Active"
        Me.CmdActive.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdImport
        '
        Me.CmdImport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdImport.Image = CType(resources.GetObject("CmdImport.Image"), System.Drawing.Image)
        Me.CmdImport.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdImport.Location = New System.Drawing.Point(765, 578)
        Me.CmdImport.Name = "CmdImport"
        Me.CmdImport.Size = New System.Drawing.Size(72, 56)
        Me.CmdImport.TabIndex = 24
        Me.CmdImport.Text = "Import"
        Me.CmdImport.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdExport
        '
        Me.CmdExport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdExport.Image = CType(resources.GetObject("CmdExport.Image"), System.Drawing.Image)
        Me.CmdExport.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdExport.Location = New System.Drawing.Point(837, 578)
        Me.CmdExport.Name = "CmdExport"
        Me.CmdExport.Size = New System.Drawing.Size(72, 56)
        Me.CmdExport.TabIndex = 25
        Me.CmdExport.Text = "Export"
        Me.CmdExport.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'FrmPreSemi
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1162, 640)
        Me.Controls.Add(Me.CmdExport)
        Me.Controls.Add(Me.CmdImport)
        Me.Controls.Add(Me.CmdActive)
        Me.Controls.Add(Me.ChkAvtive)
        Me.Controls.Add(Me.cmdDel)
        Me.Controls.Add(Me.CheckBoxType)
        Me.Controls.Add(Me.CmbMaterial)
        Me.Controls.Add(Me.CheckBoxPreSemi)
        Me.Controls.Add(Me.CmbPreSemi)
        Me.Controls.Add(Me.CmdEdit)
        Me.Controls.Add(Me.CmdClose)
        Me.Controls.Add(Me.CmdSave)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmPreSemi"
        Me.Text = "Pre Semi (Material)"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataGridCom, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "CONSTANT"
    Dim DT As New DataTable
    Dim StrSQL As String
    Dim oldrow As Integer
#End Region

#Region "Form Event"
    Private Sub FrmPreSemi_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadSemi()
        LoadPreSemi()
        LoadType()

        'If CheckBoxType.Checked = False Then
        '    GrdDV.RowFilter = " "
        '    DataGridCom.DataSource = GrdDV
        'End If

        'If CheckBoxPreSemi.Checked = False Then
        '    GrdDV.RowFilter = " "
        '    DataGridCom.DataSource = GrdDV
        'End If
        CheckBox()
    End Sub
#End Region

#Region "Function_Load"
    Private Sub LoadSemi()
        Dim sb As New System.Text.StringBuilder()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sb.AppendLine("  SELECT Typecode,TypeName,Pcode,Revision,mm,MaterialType,MaterialName")
        sb.AppendLine("  ,TQty,N,Length,gmeter,Width,Gauge,Active,cn")
        sb.AppendLine("  ,code,Mastercode,MRev,RMCode,RmRev,Qty,Qty *(nm /cntn)*(lm/1000) Qty2,Unit")
        sb.AppendLine("  ,nm,lm,wm,QPU,Ac,cntn , per")
        sb.AppendLine("  FROM (")
        sb.AppendLine("    SELECT  Typecode,TypeName,Code Pcode,Revision,MaterialName mm,MaterialType,MaterialName")
        sb.AppendLine("    ,QPU TQty,N,Length,gmeter,Width,Gauge,Active,cn")
        sb.AppendLine("    ,code+','+Revision code,'' Mastercode,'' MRev,'' RMCode,'' RmRev,null Qty,'' Unit")
        sb.AppendLine("    ,isnull(N,1) nm,isnull(Length,1000) lm,isnull(Width,1000) wm,QPU ,Active Ac,cn cntn ,null per")
        sb.AppendLine("    FROM  (")
        sb.AppendLine("      SELECT Psemicode,Revision,MaterialType,MaterialName")
        sb.AppendLine("      ,QPU,N,Length,gmeter,width,gauge,Active,cn ")
        sb.AppendLine("      FROM  TBLPreSemi p")
        sb.AppendLine("      LEFT OUTER JOIN TBLTypeMaterial  t on p.MaterialType = t.MaterialCode")
        sb.AppendLine("    ) ps")
        sb.AppendLine("    LEFT OUTER JOIN  (")
        sb.AppendLine("      SELECT t.TypeCode,TypeName,Code ")
        sb.AppendLine("      FROM TBLType t")
        sb.AppendLine("      LEFT OUTER JOIN  TBLGroup g on t.typecode = g.typecode")
        sb.AppendLine("    ) gp on ps.psemicode = code")

        sb.AppendLine("    UNION")

        sb.AppendLine("    SELECT  '' Typecode,'' TypeName,'' Pcode,'' Revision,'' mm,Materialcode,MaterialName")
        sb.AppendLine("    ,null TQty,null N,NULL Length,null gmeter,null Width,null Gauge,null Acitve,null cn,")
        sb.AppendLine("    Mastercode+','+m.Revision Code, Mastercode,m.Revision MRev,RMCode,isnull(RMRevision,'') RmRev,Qty,Unit")
        sb.AppendLine("    ,isnull(N,1) nm,isnull(Length,1000) lm,isnull(Width,1000) wm,QPU,Active Ac,cn cntn  , per")
        sb.AppendLine("    FROM (")
        sb.AppendLine("      SELECT *")
        sb.AppendLine("      FROM TBLPreSemi pr")
        sb.AppendLine("      LEFT OUTER JOIN TBLTypeMaterial t on pr.MaterialType = t.Materialcode")
        sb.AppendLine("    ) p")
        sb.AppendLine("    LEFT OUTER JOIN (")
        sb.AppendLine("      SELECT * ")
        sb.AppendLine("      FROM TBLMaster")
        sb.AppendLine("      WHERE mastercode IN ( SELECT psemicode  FROM  TBLPresemi )") 'Get PreSemi in table TBLMaster
        sb.AppendLine("    ) m on p.psemicode+p.Revision = m.mastercode+m.Revision")
        sb.AppendLine("  ) aaaa")
        sb.AppendLine("  WHERE MaterialType  NOT IN (02)") 'Exclude Material Type COATED CORD

        sb.AppendLine("  UNION ")

        sb.AppendLine("  SELECT Typecode,TypeName,Pcode,Revision,mm,MaterialType,MaterialName")
        sb.AppendLine("  ,TQty,N,Length,gmeter,Width,Gauge,Active,cn")
        sb.AppendLine("  ,code,Mastercode,MRev,RMCode,RmRev,Qty,Qty *(nm /cntn)/(wm/1000) Qty2,Unit")
        sb.AppendLine("  ,nm,lm,wm,QPU,Ac,cntn ,per ")
        sb.AppendLine("  FROM (")
        sb.AppendLine("    SELECT  Typecode,TypeName,Code Pcode,Revision,MaterialName mm,MaterialType,MaterialName")
        sb.AppendLine("    ,QPU TQty,N,Length,gmeter,Width,Gauge,Active,cn")
        sb.AppendLine("    ,code+','+Revision code,'' Mastercode,'' MRev,'' RMCode,'' RmRev,null Qty,'' Unit")
        sb.AppendLine("    ,isnull(N,1) nm,isnull(Length,1000) lm,isnull(Width,1000) wm,QPU ,Active Ac,cn cntn  ,null per  ")
        sb.AppendLine("    FROM  (")
        sb.AppendLine("      SELECT Psemicode,Revision,MaterialType,MaterialName")
        sb.AppendLine("      ,QPU,N,Length,gmeter,width,gauge,Active,cn ")
        sb.AppendLine("      FROM  TBLPreSemi p")
        sb.AppendLine("      LEFT OUTER JOIN TBLTypeMaterial  t on p.MaterialType = t.MaterialCode")
        sb.AppendLine("    ) ps")
        sb.AppendLine("    LEFT OUTER JOIN (")
        sb.AppendLine("      SELECT t.TypeCode,TypeName,Code ")
        sb.AppendLine("      FROM TBLType t")
        sb.AppendLine("      LEFT OUTER JOIN  TBLGroup g on t.typecode = g.typecode")
        sb.AppendLine("    ) gp on ps.psemicode = code")

        sb.AppendLine("    UNION ")

        sb.AppendLine("    SELECT  '' Typecode,'' TypeName,'' Pcode,'' Revision,'' mm,Materialcode,MaterialName")
        sb.AppendLine("    ,null TQty,null N,NULL Length,null gmeter,null Width,null Gauge,null Acitve,null cn,")
        sb.AppendLine("    Mastercode+','+m.Revision Code, Mastercode,m.Revision MRev,RMCode,isnull(RMRevision,'') RmRev,Qty,Unit")
        sb.AppendLine("    ,isnull(N,1) nm,isnull(Length,1000) lm,isnull(Width,1000) wm,QPU,Active Ac,cn cntn,per ")
        sb.AppendLine("    FROM (")
        sb.AppendLine("      SELECT * ")
        sb.AppendLine("      FROM TBLPresemi pr")
        sb.AppendLine("      LEFT OUTER JOIN TBLTypeMaterial t on pr.MaterialType = t.Materialcode")
        sb.AppendLine("    ) p")
        sb.AppendLine("    LEFT OUTER JOIN (")
        sb.AppendLine("      SELECT *")
        sb.AppendLine("      FROM TBLMaster")
        sb.AppendLine("      WHERE mastercode in ( SELECT psemicode  FROM  TBLPresemi )") 'Get PreSemi in table TBLMaster
        sb.AppendLine("    ) m on p.psemicode+p.Revision = m.mastercode+m.Revision")
        sb.AppendLine("  ) aaaa")
        sb.AppendLine("  WHERE MaterialType  IN (02)") 'only Material Type COATED CORD
        sb.AppendLine("  ORDER BY code,Pcode DESC")
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
        DataGridCom.DataSource = GrdDV
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

        With DataGridCom
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
            .MappingName = "mm"
            .Width = 120
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle2 As New DataGridColoredLine2
        With grdColStyle2
            .HeaderText = "Pre Semi (Material)"
            .MappingName = "Pcode"
            .Width = 150
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle2_1 As New DataGridColoredLine2
        With grdColStyle2_1
            .HeaderText = "Rev."
            .MappingName = "Revision"
            .Width = 80
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle3 As New DataGridColoredLine2
        With grdColStyle3
            .HeaderText = "Code (Material)"
            .MappingName = "RmCode"
            .Width = 120
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle3_1 As New DataGridColoredLine2
        With grdColStyle3_1
            .HeaderText = "Rev."
            .MappingName = "RMRevision"
            .NullText = ""
            .Width = 80
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle5 As New DataGridColoredLine2
        With grdColStyle5
            .HeaderText = "Weight Per Meter_"
            .MappingName = "Qty"
            .Width = 120
            .NullText = ""
            .Format = "##,###,##0.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle5_1 As New DataGridColoredLine2
        With grdColStyle5_1
            .HeaderText = "Weight of Formula_"
            .MappingName = "Qty2"
            .Width = 120
            .NullText = ""
            .Format = "##,###,##0.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle6 As New DataGridColoredLine2
        With grdColStyle6
            .HeaderText = "Unit"
            .MappingName = "Unit"
            .Width = 45
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle7 As New DataGridColoredLine2
        With grdColStyle7
            .HeaderText = "Total Per Meter_"
            .MappingName = "TQty"
            .Width = 120
            .NullText = ""
            .Format = "##,###,##0.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle8 As New DataGridColoredLine2
        With grdColStyle8
            .HeaderText = "N"
            .MappingName = "n"
            .NullText = ""
            .Width = 35
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle9 As New DataGridColoredLine2
        With grdColStyle9
            .HeaderText = "Length(mm)"
            .MappingName = "Length"
            .NullText = ""
            .Width = 75
            .Format = "##,###,##0.00"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle10 As New DataGridColoredLine2
        With grdColStyle10
            .HeaderText = "g/meter"
            .MappingName = "gmeter"
            .NullText = ""
            .Width = 75
            .Format = "##,###,##0.00"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle11 As New DataGridColoredLine2
        With grdColStyle11
            .HeaderText = "Width (mm)"
            .MappingName = "width"
            .NullText = ""
            .Width = 75
            .Format = "##,###,##0.00"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle12 As New DataGridColoredLine2
        With grdColStyle12
            .HeaderText = "CNT N."
            .MappingName = "cn"
            .NullText = ""
            .Width = 75
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle13 As New DataGridColoredLine2
        With grdColStyle13
            .HeaderText = "Active"
            .MappingName = "Active"
            .NullText = ""
            .Width = 75
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle14 As New DataGridColoredLine2
        With grdColStyle14
            .HeaderText = " % WT "
            .MappingName = "Per"
            .NullText = ""
            .Width = 65
            .Format = "#0.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        grdTableStyle1.GridColumnStyles.AddRange _
(New DataGridColumnStyle() _
{grdColStyle13, grdColStyle12, grdColStyle2, grdColStyle2_1, grdColStyle3, grdColStyle3_1 _
, grdColStyle5_1, grdColStyle5, grdColStyle6, grdColStyle7, grdColStyle14,
grdColStyle11, grdColStyle8, grdColStyle9})

        DataGridCom.TableStyles.Add(grdTableStyle1)
        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub

    Private Sub ResetTableStyle()
        ' Clear out the existing TableStyles and result default formatting.
        With DataGridCom
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
    Sub LoadPreSemi()
        Dim sb As New System.Text.StringBuilder()
        Dim dtPreSemi As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sb.AppendLine("SELECT Code,MaterialName")
        sb.AppendLine("FROM  TblGroup g")
        sb.AppendLine("LEFT OUTER JOIN ( ")
        sb.AppendLine("  SELECT distinct PsemiCode,MaterialName")
        sb.AppendLine("  FROM  TblPreSemi p")
        sb.AppendLine("  LEFT OUTER JOIN  TblTypeMaterial t on p.MaterialType = t.MaterialCode")
        sb.AppendLine(") semi on g.code = semi.Psemicode")
        sb.AppendLine("WHERE Typecode = '04'")
        StrSQL = sb.ToString()

        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtPreSemi = New DataTable
            DA.Fill(dtPreSemi)
        Catch
        Finally
        End Try
        dtPreSemi.TableName = TBL_PreSemi
        GrdDVPreSemi = dtPreSemi.DefaultView
        '************************************
        CmbPreSemi.DisplayMember = "Code"
        CmbPreSemi.ValueMember = "Code"
        CmbPreSemi.DataSource = dtPreSemi
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadType()
        Dim sb As New System.Text.StringBuilder()
        Dim dtType As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sb.AppendLine("SELECT  *  ")
        sb.AppendLine("FROM  TBLTypeMaterial ")
        sb.AppendLine("WHERE  descname LIKE '%Presemi%'")
        StrSQL = sb.ToString()

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
        CmbMaterial.ValueMember = "MaterialName"
        CmbMaterial.DataSource = dtType
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
#End Region

#Region "Control Event"
    Private Sub CmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClose.Click
        Me.Close()
    End Sub

    Private Sub CmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdEdit.Click
        Dim faddPreSemi As New FrmAddPreSemi
        faddPreSemi.CmdSave.Text = "Edit"
        faddPreSemi.TxtCode.Text = GrdDV.Item(oldrow).Row("Pcode")
        faddPreSemi.TxtCode.Enabled = False
        faddPreSemi.TxtRev.Text = GrdDV.Item(oldrow).Row("Revision")
        faddPreSemi.iCmb = GrdDV.Item(oldrow).Row("MaterialName")
        If GrdDV.Item(oldrow).Row("Active") = 0 Then
            faddPreSemi.chkbal = False
        Else
            faddPreSemi.chkbal = True
        End If
        If GrdDV.Item(oldrow).Row("MaterialName") = "COATED CORD" Then
            faddPreSemi.TxtWidth.Text = GrdDV.Item(oldrow).Row("Width")
        Else
            faddPreSemi.TxtWidth.Text = ""
        End If
        faddPreSemi.ShowDialog()
        CheckBox()
        oldrow = 0
    End Sub

    Private Sub CmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSave.Click
        Dim faddPreSemi As New FrmAddPreSemi
        faddPreSemi.CmdSave.Text = "Save"
        faddPreSemi.iCmb = GrdDV.Item(oldrow).Row("MaterialName")
        faddPreSemi.ShowDialog()
        LoadSemi()
        CheckBox()
        oldrow = 0
    End Sub

    Private Sub DataGridCOM_CurrentCellChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGridCom.CurrentCellChanged
        oldrow = DataGridCom.CurrentCell.RowNumber
    End Sub

    Private Sub CmdActive_Click(sender As Object, e As EventArgs) Handles CmdActive.Click
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        If GrdDV.Item(oldrow).Row("Pcode") = "" Then
            Exit Sub
        End If
        msg = "Change Active Semi(Material) : " & GrdDV.Item(oldrow).Row("Pcode") _
        & "  Revision :" & GrdDV.Item(oldrow).Row("Revision") 'Define message.
        style = MsgBoxStyle.DefaultButton2 Or
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

    Private Sub CmdExport_Click(sender As Object, e As EventArgs) Handles CmdExport.Click
        Dim arrColumn As String() = System.Configuration.ConfigurationManager.AppSettings("EXP_EXCEL_COLUMN_MASTER_PRESEMI").ToString().Split(New Char() {","c})
        Dim arrColumnHeader As String() = System.Configuration.ConfigurationManager.AppSettings("EXP_EXCEL_COLUMN_HEADER_MASTER_PRESEMI").ToString().Split(New Char() {","c})
        ExcelLib.Export(Me, GrdDV, TBL_RM, arrColumn, arrColumnHeader)
    End Sub

    Private Sub CmdImport_Click(sender As Object, e As EventArgs) Handles CmdImport.Click
        Dim arrColumn As String() = System.Configuration.ConfigurationManager.AppSettings("IMP_EXCEL_COLUMN_MASTER_PRESEMI").ToString().Split(New Char() {","c})
        Dim importDialog As OpenFileDialog = New OpenFileDialog With {
            .Filter = System.Configuration.ConfigurationManager.AppSettings("DIALOG_FILE_EXT").ToString()
        }
        Dim dtRec As DataTable
        Dim sb As New System.Text.StringBuilder()
        Dim frmOverlay As New Form()
        Dim totalQty As Double = 0

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
            dtRec = ExcelLib.Import(importDialog.FileName, Me, GrdDV, TBL_RM, arrColumn)

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
                        dtRec.DefaultView.Sort = "TypeMaterial DESC, PreSemi DESC, PreSemiRevision DESC"
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
                            Dim strPreSemi As String = dtRec.Rows(i)("PreSemi").ToString().Trim()
                            Dim strRevision As String = dtRec.Rows(i)("PreSemiRevision").ToString().Trim()
                            Dim strRMCode As String = dtRec.Rows(i)("RMCode").ToString().Trim()

                            'Check Empty
                            If strPreSemi.Equals(String.Empty) Or strRevision.Equals(String.Empty) Then
                                Throw New ApplicationException("Presemi Code, Presemi Revision is not empty.")
                            End If

                            'Check Type Material Master
                            Dim arrTypeMatCode As DataRow() = dtTypeMaterial.Select("MaterialName = '" & strTypeMaterial & "'")
                            strTypeMaterial = arrTypeMatCode(0)("MaterialCode")

                            If strTypeMaterial.Length > 0 And strPreSemi.Length > 0 And strRevision.Length > 0 And strRMCode.Length > 0 Then
                                Dim dblQty As Double = 0
                                Dim dblWidth As Double = 0
                                Dim dblLength As Double = 0
                                Dim intN As Integer = 0
                                Dim strUnit As String = dtRec.Rows(i)("Unit").ToString().Trim()
                                Dim GridRow As DataRow()        '//Grid Data
                                Dim ExcelRow As DataRow()       '//Excel Data

                                '//Check Qty input format as Number
                                If dtRec.Rows(i)("Qty").ToString.Length > 0 Then
                                    If Not Double.TryParse(dtRec.Rows(i)("Qty"), dblQty) Then
                                        Throw New System.Exception("Please input Qty data as Number")
                                    End If
                                Else
                                    Throw New System.Exception("Please input Qty data as Number")
                                End If

                                GridRow = DT.Select("Mastercode = '" & strPreSemi & "' AND MRev = '" & strRevision & "'")
                                If GridRow.Count > 0 Then

                                    strTypeMaterialOriginal = GridRow(0)("MaterialType").ToString

                                    '//Check Width,Length and N as Number by Condition
                                    If GridRow(0)("MaterialType").ToString = "01" Then
                                        '//Just QTY
                                    ElseIf GridRow(0)("MaterialType").ToString = "02" Then
                                        If dtRec.Rows(i)("Width").ToString.Length > 0 Then
                                            If Not Double.TryParse(dtRec.Rows(i)("Width"), dblLength) Then
                                                Throw New System.Exception("Please input Width data as Number")
                                            End If
                                        Else
                                            Throw New System.Exception("Please input Width data as Number")
                                        End If
                                    ElseIf GridRow(0)("MaterialType").ToString = "16" Then
                                        If dtRec.Rows(i)("Length").ToString.Length > 0 Then
                                            If Not Double.TryParse(dtRec.Rows(i)("Length"), dblLength) Then
                                                Throw New System.Exception("Please input Length data as Number")
                                            End If
                                        Else
                                            Throw New System.Exception("Please input Length data as Number")
                                        End If
                                    ElseIf GridRow(0)("MaterialType").ToString = "19" Then
                                        If dtRec.Rows(i)("N").ToString.Length > 0 Then
                                            If Not Integer.TryParse(dtRec.Rows(i)("N"), intN) Then
                                                Throw New System.Exception("Please input N data as Number")
                                            End If
                                        Else
                                            Throw New System.Exception("Please input N data as Number")
                                        End If
                                    ElseIf GridRow(0)("MaterialType").ToString = "21" Then
                                        If dtRec.Rows(i)("Length").ToString.Length > 0 Then
                                            If Not Double.TryParse(dtRec.Rows(i)("Length"), dblLength) Then
                                                Throw New System.Exception("Please input Length data as Number")
                                            End If
                                        Else
                                            Throw New System.Exception("Please input Length data as Number")
                                        End If
                                    Else
                                        If dtRec.Rows(i)("Length").ToString.Length > 0 Then
                                            If Not Double.TryParse(dtRec.Rows(i)("Length"), dblLength) Then
                                                Throw New System.Exception("Please input Length data as Number")
                                            End If
                                        Else
                                            Throw New System.Exception("Please input Length data as Number")
                                        End If

                                        If dtRec.Rows(i)("N").ToString.Length > 0 Then
                                            If Not Integer.TryParse(dtRec.Rows(i)("N"), intN) Then
                                                Throw New System.Exception("Please input N data as Number")
                                            End If
                                        Else
                                            Throw New System.Exception("Please input N data as Number")
                                        End If
                                    End If
                                Else
                                    '//Check Width,Length and N as Number by Condition
                                    If strTypeMaterial = "01" Then
                                        '//Just QTY
                                    ElseIf strTypeMaterial = "02" Then
                                        If dtRec.Rows(i)("Width").ToString.Length > 0 Then
                                            If Not Double.TryParse(dtRec.Rows(i)("Width"), dblLength) Then
                                                Throw New System.Exception("Please input Width data as Number")
                                            End If
                                        Else
                                            Throw New System.Exception("Please input Width data as Number")
                                        End If
                                    ElseIf strTypeMaterial = "16" Then
                                        If dtRec.Rows(i)("Length").ToString.Length > 0 Then
                                            If Not Double.TryParse(dtRec.Rows(i)("Length"), dblLength) Then
                                                Throw New System.Exception("Please input Length data as Number")
                                            End If
                                        Else
                                            Throw New System.Exception("Please input Length data as Number")
                                        End If
                                    ElseIf strTypeMaterial = "19" Then
                                        If dtRec.Rows(i)("N").ToString.Length > 0 Then
                                            If Not Integer.TryParse(dtRec.Rows(i)("N"), intN) Then
                                                Throw New System.Exception("Please input N data as Number")
                                            End If
                                        Else
                                            Throw New System.Exception("Please input N data as Number")
                                        End If
                                    ElseIf strTypeMaterial = "21" Then
                                        If dtRec.Rows(i)("Length").ToString.Length > 0 Then
                                            If Not Double.TryParse(dtRec.Rows(i)("Length"), dblLength) Then
                                                Throw New System.Exception("Please input Length data as Number")
                                            End If
                                        Else
                                            Throw New System.Exception("Please input Length data as Number")
                                        End If
                                    Else
                                        If dtRec.Rows(i)("Length").ToString.Length > 0 Then
                                            If Not Double.TryParse(dtRec.Rows(i)("Length"), dblLength) Then
                                                Throw New System.Exception("Please input Length data as Number")
                                            End If
                                        Else
                                            Throw New System.Exception("Please input Length data as Number")
                                        End If

                                        If dtRec.Rows(i)("N").ToString.Length > 0 Then
                                            If Not Integer.TryParse(dtRec.Rows(i)("N"), intN) Then
                                                Throw New System.Exception("Please input N data as Number")
                                            End If
                                        Else
                                            Throw New System.Exception("Please input N data as Number")
                                        End If
                                    End If
                                End If

                                '//For Check Data from above row on import file.
                                Dim chkSamePreSemiBefore As String = String.Empty
                                Dim chkSameRevisionBefore As String = String.Empty
                                Dim chkSamePreSemiRevAbove As Boolean = False
                                If i > 0 Then
                                    chkSamePreSemiBefore = dtRec.Rows(i - 1)("PreSemi").ToString
                                    chkSameRevisionBefore = dtRec.Rows(i - 1)("PreSemiRevision").ToString
                                Else
                                    chkSamePreSemiBefore = ""
                                    chkSameRevisionBefore = ""
                                End If

                                If strPreSemi <> chkSamePreSemiBefore Or strRevision <> chkSameRevisionBefore Then
                                    totalQty = 0
                                    GridRow = DT.Select("Mastercode = '" & strPreSemi & "' AND MRev = '" & strRevision & "'")
                                    If GridRow.Count > 0 Then '//Have PreSemi and Rev on DB so can use TypeMat from DB
                                        Dim sameTypeMat As String = GridRow(0)("MaterialType").ToString
                                        For k As Integer = 0 To GridRow.Count - 1
                                            totalQty = totalQty + GridRow(k)("Qty")
                                        Next k

                                        ExcelRow = dtRec.Select("PreSemi = '" & strPreSemi & "' AND PreSemiRevision = '" & strRevision & "'")
                                        For j As Integer = 0 To ExcelRow.Count - 1
                                            GridRow = DT.Select("Mastercode = '" & strPreSemi & "' AND MRev = '" & strRevision & "' AND RMCode = '" & ExcelRow(j)("RMCode") & "'")

                                            If GridRow.Count > 0 Then
                                                totalQty = totalQty - GridRow(0)("Qty")
                                            End If

                                            If sameTypeMat = "01" Then
                                                totalQty = totalQty + CDbl(ExcelRow(j)("Qty"))
                                            ElseIf sameTypeMat = "02" Then
                                                totalQty = totalQty + ((CDbl(ExcelRow(j)("Qty")) * dblWidth) / 1000)
                                            ElseIf sameTypeMat = "19" Then
                                                totalQty = totalQty + (CDbl(ExcelRow(j)("Qty")) / intN)
                                            ElseIf sameTypeMat = "21" Then
                                                totalQty = totalQty + ((CDbl(ExcelRow(j)("Qty")) / dblLength) * 1000)
                                            ElseIf sameTypeMat = "16" Then
                                                totalQty = totalQty + ((CDbl(ExcelRow(j)("Qty")) / dblLength) * 1000)
                                            Else
                                                totalQty = totalQty + (((CDbl(ExcelRow(j)("Qty")) / dblLength) * 1000) / intN)
                                            End If
                                        Next j
                                    Else '//New PreSemi and Rev so use TypeMat from ImportExcel
                                        ExcelRow = dtRec.Select("PreSemi = '" & strPreSemi & "' AND PreSemiRevision = '" & strRevision & "'")
                                        For j As Integer = 0 To ExcelRow.Count - 1
                                            If ExcelRow(j)("TypeMaterial") = "01" Then
                                                totalQty = totalQty + CDbl(ExcelRow(j)("Qty"))
                                            ElseIf ExcelRow(j)("TypeMaterial") = "02" Then
                                                totalQty = totalQty + ((CDbl(ExcelRow(j)("Qty")) * dblWidth) / 1000)
                                            ElseIf ExcelRow(j)("TypeMaterial") = "19" Then
                                                totalQty = totalQty + (CDbl(ExcelRow(j)("Qty")) / intN)
                                            ElseIf ExcelRow(j)("TypeMaterial") = "21" Then
                                                totalQty = totalQty + ((CDbl(ExcelRow(j)("Qty")) / dblLength) * 1000)
                                            ElseIf ExcelRow(j)("TypeMaterial") = "16" Then
                                                totalQty = totalQty + ((CDbl(ExcelRow(j)("Qty")) / dblLength) * 1000)
                                            Else
                                                totalQty = totalQty + (((CDbl(ExcelRow(j)("Qty")) / dblLength) * 1000) / intN)
                                            End If
                                        Next j
                                    End If
                                Else
                                    chkSamePreSemiRevAbove = True
                                End If

                                '// 1.) Check PreSemi and Revision fron DB
                                '// 1.1) [NG] Insert TBLMASTER,TBLGroup,TBLPreSemi and TBLConvert
                                '// 1.2) [OK] Find Import RMCode on DB
                                '// 1.2.1) [NG] Insert TBLMASTER / Update TBLPreSemi,TBLConvert
                                '// 1.2.2) [OK] Compare QTY from Import file and DB
                                '// 1.2.2.1) [NG] Update TBLMASTER,TBLPreSemi,TBLConvert

                                '// 1.) Check PreSemi and Revision fron DB
                                GridRow = DT.Select("Mastercode = '" & strPreSemi & "' AND MRev = '" & strRevision & "'")

                                If GridRow.Count > 0 Or chkSamePreSemiRevAbove = True Then '// 1.2) [OK] Find Import RMCode on DB

                                    GridRow = DT.Select("Mastercode = '" & strPreSemi & "' AND MRev = '" & strRevision & "' AND RMCode = '" & strRMCode & "'")

                                    If GridRow.Count > 0 AndAlso CDbl(GridRow(0)("Qty")) <> dblQty Then '// 1.2.2) [OK] Compare QTY from Import file and DB

                                        '// 1.2.2.1) [NG] Update TBLMASTER,TBLPreSemi,TBLConvert
                                        '//Update TBLMASTER
                                        sb.Clear()
                                        sb.AppendLine(" Update TBLMASTER")
                                        sb.AppendLine(" Set ")

                                        If strTypeMaterialOriginal = "01" Then
                                            sb.AppendLine(" Qty = '" & dblQty & "' ")
                                        ElseIf strTypeMaterialOriginal = "02" Then
                                            sb.AppendLine(" Qty = '" & ((dblQty * dblWidth) / 1000) & "' ")
                                        ElseIf strTypeMaterialOriginal = "19" Then
                                            sb.AppendLine(" Qty = '" & (dblQty / intN) & "' ")
                                        ElseIf strTypeMaterialOriginal = "16" Then
                                            sb.AppendLine(" Qty = '" & ((dblQty / dblLength) * 1000) & "' ")
                                        Else
                                            sb.AppendLine(" Qty = '" & (((dblQty / dblLength) * 1000) / intN) & "' ")
                                        End If

                                        sb.AppendLine(" Where MasterCode = '" & strPreSemi & "' AND Revision = '" & strRevision & "' AND RMCode = '" & strRMCode & "' ")

                                        sb.AppendLine(" ")

                                        '//Update TBLPreSemi
                                        sb.AppendLine(" Update TBLPreSemi")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" QPU = '" & totalQty & "'")
                                        sb.AppendLine(" Where PSemiCode = '" & strPreSemi & "' AND Revision = '" & strRevision & "' ")

                                        sb.AppendLine(" ")

                                        '//Update TBLConvert
                                        sb.AppendLine(" Update TBLConvert")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" SQty = '" & totalQty & "'")
                                        sb.AppendLine(" Where Code = '" & strPreSemi & "' AND Rev = '" & strRevision & "' ")

                                        StrSQL = sb.ToString()
                                        cmSQL.CommandText = StrSQL
                                        cmSQL.ExecuteNonQuery()


                                        '//Update All Per in TBLMASTER***********
                                        sb.Clear()
                                        sb.AppendLine(" Update TBLMASTER")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" Per = Qty*(100/" & totalQty & ")")
                                        sb.AppendLine(" Where MasterCode = '" & strPreSemi & "' AND Revision = '" & strRevision & "' ")
                                        StrSQL = sb.ToString()
                                        cmSQL.CommandText = StrSQL
                                        cmSQL.ExecuteNonQuery()

                                    Else '// 1.2.1) [NG] Insert TBLMASTER / Update TBLPreSemi,TBLConvert

                                        '//Insert TBLMASTER
                                        sb.Clear()
                                        sb.AppendLine(" Insert  TBLMASTER ")
                                        sb.AppendLine(" Values (")
                                        sb.AppendLine(" '" & strPreSemi & "', ")                                    'Column MasterCode
                                        sb.AppendLine(" '" & strRevision & "' , ")                                  'Column Revision
                                        sb.AppendLine(" '" & strRMCode & "' , ")                                    'Column RMCode
                                        sb.AppendLine(" NULL , ")                                                   'Column RmRevision

                                        If strTypeMaterialOriginal = "01" Then                                              'Column Qty
                                            sb.AppendLine(" '" & dblQty & "', ")
                                        ElseIf strTypeMaterialOriginal = "02" Then
                                            sb.AppendLine(" '" & ((dblQty * dblWidth) / 1000) & "', ")
                                        ElseIf strTypeMaterialOriginal = "19" Then
                                            sb.AppendLine(" '" & (dblQty / intN) & "', ")
                                        ElseIf strTypeMaterialOriginal = "16" Then
                                            sb.AppendLine(" '" & ((dblQty / dblLength) * 1000) & "', ")
                                        Else
                                            sb.AppendLine(" '" & (((dblQty / dblLength) * 1000) / intN) & "', ")
                                        End If

                                        sb.AppendLine(" '" & strUnit & "' , ")                                      'Column Unit
                                        sb.AppendLine(" '" & ((dblQty * 100) / totalQty) & "'")                     'Column Per
                                        sb.AppendLine(" )")

                                        sb.AppendLine(" ")

                                        '//Update TBLPreSemi
                                        sb.AppendLine(" Update TBLPreSemi")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" QPU = '" & totalQty & "'")
                                        sb.AppendLine(" Where PSemiCode = '" & strPreSemi & "' AND Revision = '" & strRevision & "' ")

                                        sb.AppendLine(" ")

                                        '//Update TBLConvert
                                        sb.AppendLine(" Update TBLConvert")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" SQty = '" & totalQty & "'")
                                        sb.AppendLine(" Where Code = '" & strPreSemi & "' AND Rev = '" & strRevision & "' ")

                                        StrSQL = sb.ToString()
                                        cmSQL.CommandText = StrSQL
                                        cmSQL.ExecuteNonQuery()


                                        '//Update All Per in TBLMASTER***********
                                        sb.Clear()
                                        sb.AppendLine(" Update TBLMASTER")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" Per = Qty*(100/" & totalQty & ")")
                                        sb.AppendLine(" Where MasterCode = '" & strPreSemi & "' AND Revision = '" & strRevision & "' ")
                                        StrSQL = sb.ToString()
                                        cmSQL.CommandText = StrSQL
                                        cmSQL.ExecuteNonQuery()

                                    End If

                                Else '// 1.1) [NG] Insert TBLMASTER,TBLGroup,TBLPreSemi and TBLConvert

                                    '//Insert TBLGroup
                                    sb.Clear()
                                    sb.AppendLine(" Insert  TBLGroup ")
                                    sb.AppendLine(" Values (")
                                    sb.AppendLine("'04' , ")                    'Column TypeCode
                                    sb.AppendLine("'" & strPreSemi & "'")       'Column Code
                                    sb.AppendLine(" )")

                                    sb.AppendLine(" ")

                                    '//Insert TBLPreSemi
                                    sb.AppendLine(" Insert TBLPreSemi ")
                                    sb.AppendLine(" Values (")
                                    sb.AppendLine(" '" & strPreSemi & "', ")        'Column Final
                                    sb.AppendLine(" '" & strPreSemi & "', ")        'Column PsemiCode
                                    sb.AppendLine(" '" & strRevision & "', ")       'Column Revision
                                    sb.AppendLine(" '" & strTypeMaterial & "', ")   'Column MaterialType
                                    sb.AppendLine(" '" & totalQty & "', ")          'Column QPU

                                    If intN = 0 Then                                'Column N
                                        sb.AppendLine(" NULL , ")
                                    Else
                                        sb.AppendLine(" '" & intN & "', ")
                                    End If

                                    If dblLength = 0 Then                           'Column Length
                                        sb.AppendLine(" NULL , ")
                                    Else
                                        sb.AppendLine(" '" & dblLength & "' , ")
                                    End If

                                    sb.AppendLine(" NULL , ")                       'Column gmeter

                                    If dblWidth = 0 Then                            'Column Width
                                        sb.AppendLine(" NULL , ")
                                    Else
                                        sb.AppendLine(" '" & dblWidth & "' , ")
                                    End If

                                    sb.AppendLine(" NULL , ")                       'Column Gauge
                                    sb.AppendLine(" '0' , ")                        'Column Active
                                    sb.AppendLine(" '" & strDate & "' , ")          'Column DateUp

                                    If strTypeMaterial = "16" Then                  'Column CN
                                        sb.AppendLine(" '2' , ")
                                    Else
                                        sb.AppendLine(" '1' , ")
                                    End If

                                    sb.AppendLine(" )")

                                    sb.AppendLine(" ")

                                    '//Insert TblMaster
                                    sb.AppendLine(" Insert  TBLMASTER ")
                                    sb.AppendLine(" Values (")
                                    sb.AppendLine(" '" & strPreSemi & "', ")                                    'Column MasterCode
                                    sb.AppendLine(" '" & strRevision & "' , ")                                  'Column Revision
                                    sb.AppendLine(" '" & strRMCode & "' , ")                                    'Column RMCode
                                    sb.AppendLine(" NULL , ")                                                   'Column RmRevision

                                    If strTypeMaterial = "01" Then                                              'Column Qty
                                        sb.AppendLine(" '" & dblQty & "', ")
                                    ElseIf strTypeMaterial = "02" Then
                                        sb.AppendLine(" '" & ((dblQty * dblWidth) / 1000) & "', ")
                                    ElseIf strTypeMaterial = "19" Then
                                        sb.AppendLine(" '" & (dblQty / intN) & "', ")
                                    ElseIf strTypeMaterial = "16" Then
                                        sb.AppendLine(" '" & ((dblQty / dblLength) * 1000) & "', ")
                                    Else
                                        sb.AppendLine(" '" & (((dblQty / dblLength) * 1000) / intN) & "', ")
                                    End If

                                    sb.AppendLine(" '" & strUnit & "' , ")                                      'Column Unit
                                    sb.AppendLine(" '" & ((dblQty * 100) / totalQty) & "'")                     'Column Per
                                    sb.AppendLine(" )")

                                    sb.AppendLine(" ")

                                    '//Insert TblConvert #1
                                    sb.AppendLine(" Insert  TblConvert ")
                                    sb.AppendLine(" Values (")
                                    sb.AppendLine(" '04' , ")                   'Column Type
                                    sb.AppendLine(" '" & strPreSemi & "', ")    'Column Final
                                    sb.AppendLine(" '" & strPreSemi & "', ")    'Column Code
                                    sb.AppendLine(" '" & strRevision & "' , ")  'Column Rev
                                    sb.AppendLine(" 'KG' , ")                   'Column UnitBig
                                    sb.AppendLine(" 'KG' , ")                   'Column UnitSmall
                                    sb.AppendLine(" '1' , ")                    'Column BQty
                                    sb.AppendLine(" '1' ")                      'Column SQty
                                    sb.AppendLine(" )")

                                    '//Insert TblConvert #2
                                    If strTypeMaterial <> "01" Then
                                        sb.AppendLine(" ")

                                        sb.AppendLine(" Insert  TblConvert ")
                                        sb.AppendLine(" Values (")
                                        sb.AppendLine(" '04' , ")                   'Column Type
                                        sb.AppendLine(" '" & strPreSemi & "', ")    'Column Final
                                        sb.AppendLine(" '" & strPreSemi & "', ")    'Column Code
                                        sb.AppendLine(" '" & strRevision & "' , ")  'Column Rev

                                        If strTypeMaterial = "02" Then              'Column UnitBig
                                            sb.AppendLine(" 'M' , ")
                                        ElseIf strTypeMaterial = "19" Then
                                            sb.AppendLine(" 'UT' , ")
                                        Else
                                            sb.AppendLine(" 'M' , ")
                                        End If

                                        sb.AppendLine(" 'KG' , ")                   'Column UnitSmall
                                        sb.AppendLine(" '1' , ")                    'Column BQty

                                        If strUnit = "KG" Then                      'Column SQty
                                            sb.AppendLine(" '" & totalQty & "' ")
                                        Else
                                            sb.AppendLine(" '" & (totalQty / 1000) & "' ")
                                        End If

                                        sb.AppendLine(" )")
                                    End If

                                    StrSQL = sb.ToString()
                                    cmSQL.CommandText = StrSQL
                                    cmSQL.ExecuteNonQuery()

                                End If

                            End If
                        Next

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
                    sb.AppendLine(" Select COUNT(*)  ")
                    sb.AppendLine(" FROM (  ")
                    sb.AppendLine("   SELECT t.Typecode , TypeName, code")
                    sb.AppendLine("   FROM TBLType t  ")
                    sb.AppendLine("   LEFT OUTER JOIN TBLGroup g on t.typecode = g.typecode ")
                    sb.AppendLine(" ) a ")
                    sb.AppendLine(" LEFT OUTER JOIN ( ")
                    sb.AppendLine("   SELECT distinct  Finalcompound code , null DescName, compcode, 0.00 Qty ")
                    sb.AppendLine("   FROM TBLCompound ")
                    sb.AppendLine("   WHERE Compcode NOT IN ( SELECT code FROM TblGroup WHERE Typecode ='01') AND active = 1 ")
                    sb.AppendLine("   UNION ")
                    sb.AppendLine("   SELECT rmCode,DescName,RMcode, 0.00 Qty ")
                    sb.AppendLine("   FROM TblRM ")
                    sb.AppendLine("   WHERE descName Like '%Steel%' or descName like '%Bead%' ")
                    sb.AppendLine(" ) b on a.code = b.compcode ")
                    sb.AppendLine(" WHERE b.code Is Not null AND b.code = '" & rmCode & "' ")
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
                    Throw New System.Exception("RM Code is not empty.")
                End If


                '// 2.) Check TypeMaterial
                If strTypeMat.Length = 0 Then
                    'Empty
                    Throw New System.Exception("Type Material is not empty.")
                End If

                Dim arrTypeMatCode As DataRow() = dtTypeMaterial.Select("MaterialName = '" & strTypeMat & "'")
                If arrTypeMatCode.Length = 0 Then
                    Throw New System.Exception("This TypeMaterial Code '" & strTypeMat & "' is not match PreSemi TypeMaterial")
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
            sb.AppendLine(" WHERE descName like '%PreSemi%' ")
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
#End Region

#Region "SelectData"
    Private Sub CmbPreSemi_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbPreSemi.SelectedIndexChanged
        CheckBox()
    End Sub

    Private Sub CheckBoxPreSemi_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxPreSemi.CheckedChanged
        CheckBox()
    End Sub

    Private Sub CheckBoxType_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxType.CheckedChanged
        CheckBox()
    End Sub

    Sub CheckBox()
        If CheckBoxPreSemi.Checked = True And CheckBoxType.Checked = True And ChkAvtive.Checked = False Then
            GrdDVPreSemi.RowFilter = " MaterialName = '" & CmbMaterial.Text.Trim & "'"
            CmbPreSemi.DataSource = GrdDVPreSemi
            GrdDV.RowFilter = " MaterialName like'%" & CmbMaterial.Text.Trim & "%'" _
                              & " and Code like'%" & CmbPreSemi.Text.Trim & "%'"
            DataGridCom.DataSource = GrdDV
            CmbMaterial.Enabled = True
            CmbPreSemi.Enabled = True
        ElseIf CheckBoxPreSemi.Checked = False And CheckBoxType.Checked = True And ChkAvtive.Checked = False Then
            GrdDVPreSemi.RowFilter = " MaterialName = '" & CmbMaterial.Text.Trim & "'"
            CmbPreSemi.DataSource = GrdDVPreSemi
            GrdDV.RowFilter = " MaterialName like'%" & CmbMaterial.Text.Trim & "%'"
            DataGridCom.DataSource = GrdDV
            CmbMaterial.Enabled = True
            CmbPreSemi.Enabled = False
        ElseIf CheckBoxPreSemi.Checked = True And CheckBoxType.Checked = False And ChkAvtive.Checked = False Then
            GrdDVPreSemi.RowFilter = " "
            CmbPreSemi.DataSource = GrdDVPreSemi
            GrdDV.RowFilter = " Code like'%" & CmbPreSemi.Text.Trim & "%'"
            DataGridCom.DataSource = GrdDV
            CmbMaterial.Enabled = False
            CmbPreSemi.Enabled = True
        ElseIf CheckBoxPreSemi.Checked = False And CheckBoxType.Checked = False And ChkAvtive.Checked = False Then
            GrdDV.RowFilter = " "
            DataGridCom.DataSource = GrdDV
            CmbMaterial.Enabled = False
            CmbPreSemi.Enabled = False
        ElseIf CheckBoxPreSemi.Checked = True And CheckBoxType.Checked = True And ChkAvtive.Checked = True Then
            GrdDVPreSemi.RowFilter = " MaterialName = '" & CmbMaterial.Text.Trim & "'"
            CmbPreSemi.DataSource = GrdDVPreSemi
            GrdDV.RowFilter = " MaterialName like'%" & CmbMaterial.Text.Trim & "%'" _
                              & " and Code like'%" & CmbPreSemi.Text.Trim & "%'" _
                              & " and Ac = 1 "
            DataGridCom.DataSource = GrdDV
            CmbMaterial.Enabled = True
            CmbPreSemi.Enabled = True
        ElseIf CheckBoxPreSemi.Checked = False And CheckBoxType.Checked = True And ChkAvtive.Checked = True Then
            GrdDVPreSemi.RowFilter = " MaterialName = '" & CmbMaterial.Text.Trim & "'"
            CmbPreSemi.DataSource = GrdDVPreSemi
            GrdDV.RowFilter = " MaterialName like'%" & CmbMaterial.Text.Trim & "%'" _
                              & " and Ac = 1 "
            DataGridCom.DataSource = GrdDV
            CmbMaterial.Enabled = True
            CmbPreSemi.Enabled = False
        ElseIf CheckBoxPreSemi.Checked = True And CheckBoxType.Checked = False And ChkAvtive.Checked = True Then
            GrdDVPreSemi.RowFilter = " "
            CmbPreSemi.DataSource = GrdDVPreSemi
            GrdDV.RowFilter = " Code like'%" & CmbPreSemi.Text.Trim & "%'" _
                              & " and Ac = 1 "
            DataGridCom.DataSource = GrdDV
            CmbMaterial.Enabled = False
            CmbPreSemi.Enabled = True
        Else : CheckBoxPreSemi.Checked = False And CheckBoxType.Checked = False And ChkAvtive.Checked = True
            GrdDV.RowFilter = " Ac = 1 "
            DataGridCom.DataSource = GrdDV
            CmbMaterial.Enabled = False
            CmbPreSemi.Enabled = False
        End If
    End Sub

    Private Sub CmbMaterial_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbMaterial.SelectedIndexChanged
        CheckBox()
    End Sub
#End Region

    Private Sub cmdDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDel.Click
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        If GrdDV.Item(oldrow).Row("Pcode") = "" Then
            Exit Sub
        End If
        msg = "Delete Semi(Material) : " & GrdDV.Item(oldrow).Row("Pcode") _
           & "  Revision :" & GrdDV.Item(oldrow).Row("Revision") 'Define message.
        style = MsgBoxStyle.DefaultButton2 Or
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Semi(Material)"   ' Define title.
        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            If ChkData() Then
                DelSemi()
                LoadSemi()
                CheckBox()
            Else
                MsgBox("Can't Delete. Please check Usage.", MsgBoxStyle.OkOnly, "Tire")
            End If
        Else
            Exit Sub
        End If
    End Sub

#Region "DelSemi"
    Private Function ChkUPData() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Try
            strSQL &= " select count(*) from TblPreSemi"
            strSQL &= " where PsemiCode  = '" & GrdDV.Item(oldrow).Row("Pcode") & "'"
            strSQL &= " and Revision  = '" & GrdDV.Item(oldrow).Row("Revision") & "'"
            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar
            If i = 1 Then
                ChkUPData = True
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
    Private Function ChkData() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Try
            strSQL &= " select count(*) from TblPreSemi"
            strSQL &= " where PsemiCode  = '" & GrdDV.Item(oldrow).Row("Pcode") & "'"
            strSQL &= " and Revision  = '" & GrdDV.Item(oldrow).Row("Revision") & "'"
            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar
            If i = 1 Then
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
    Private Function ChkDataGroup() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Try
            strSQL &= " select count(*) from TblPreSemi"
            strSQL &= " where PsemiCode  = '" & GrdDV.Item(oldrow).Row("Pcode") & "'"
            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar
            If i = 1 Then
                ChkDataGroup = True
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
            strSQL = " Delete TblPreSemi"
            strSQL &= " where PSemiCode = '" & GrdDV.Item(oldrow).Row("Pcode") & "'"
            strSQL &= " and Revision = '" & GrdDV.Item(oldrow).Row("Revision") & "'"
            strSQL &= "  "
            strSQL &= " Delete TblMaster"
            strSQL &= " where Mastercode = '" & GrdDV.Item(oldrow).Row("Pcode") & "'"
            strSQL &= " and Revision = '" & GrdDV.Item(oldrow).Row("Revision") & "'"
            strSQL &= "  "
            strSQL &= " Delete Tblconvert"
            strSQL &= " where code = '" & GrdDV.Item(oldrow).Row("Pcode") & "'"
            strSQL &= " and Rev = '" & GrdDV.Item(oldrow).Row("Revision") & "'"
            If GrdDV.Item(oldrow).Row("MaterialType") = "13" Then
                strSQL &= " and unitBig = 'UT'"
            ElseIf GrdDV.Item(oldrow).Row("MaterialType") = "14" Then
                strSQL &= " and unitBig = 'UT'"
            ElseIf GrdDV.Item(oldrow).Row("MaterialType") = "01" Then
                strSQL &= " and unitBig = 'KG'"
            Else
                strSQL &= " and unitBig = 'M'"
            End If

            If ChkDataGroup() Then
                strSQL &= "  "
                strSQL &= " Delete TblGroup"
                strSQL &= " where Code = '" & GrdDV.Item(oldrow).Row("Pcode") & "'"
            Else
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
    End Sub
#End Region

    Sub UPSemi()
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Try
            strSQL = " Update TblPreSemi"
            strSQL &= " set Active = 0"
            strSQL &= " where PSemiCode = '" & GrdDV.Item(oldrow).Row("Pcode") & "'"
            strSQL &= " "
            strSQL &= " Update TblPreSemi"
            strSQL &= " set Active = 1"
            strSQL &= " where PSemiCode = '" & GrdDV.Item(oldrow).Row("Pcode") & "'"
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

    Private Sub ChkAvtive_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkAvtive.CheckedChanged
        CheckBox()
    End Sub

    Private Sub DataGridCom_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DataGridCom.KeyPress

    End Sub
End Class
