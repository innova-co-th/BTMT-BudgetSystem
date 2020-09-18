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

Public Class FrmCompound

#Region "Declare"
    Inherits System.Windows.Forms.Form
    Dim GrdDV As New DataView
    Protected Const TBL_RM As String = "TBL_RM"
    Dim GrdDVComp As New DataView
    Protected Const TBL_Comp As String = "TBL_Comp"
    Dim GrdDVGP As New DataView
    Protected Const TBL_Group As String = "TBL_Group"

    Protected DefaultGridBorderStyle As BorderStyle
    Dim C1 As New SQLData("ACCINV")
    Dim StrData As String
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
    Friend WithEvents CmbCompound As System.Windows.Forms.ComboBox
    Friend WithEvents CheckBoxCompoud As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxGP As System.Windows.Forms.CheckBox
    Friend WithEvents CmbGroup As System.Windows.Forms.ComboBox
    Friend WithEvents CmdDelete As System.Windows.Forms.Button
    Friend WithEvents cmdAvtive As System.Windows.Forms.Button
    Friend WithEvents ChkActive As System.Windows.Forms.CheckBox
    Friend WithEvents CmdImport As System.Windows.Forms.Button
    Friend WithEvents CmdExport As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmCompound))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.DataGridCOM = New System.Windows.Forms.DataGrid
        Me.CmdSave = New System.Windows.Forms.Button
        Me.CmdClose = New System.Windows.Forms.Button
        Me.CmdEdit = New System.Windows.Forms.Button
        Me.CmbCompound = New System.Windows.Forms.ComboBox
        Me.CheckBoxCompoud = New System.Windows.Forms.CheckBox
        Me.CheckBoxGP = New System.Windows.Forms.CheckBox
        Me.CmbGroup = New System.Windows.Forms.ComboBox
        Me.CmdDelete = New System.Windows.Forms.Button
        Me.cmdAvtive = New System.Windows.Forms.Button
        Me.ChkActive = New System.Windows.Forms.CheckBox
        Me.CmdImport = New System.Windows.Forms.Button
        Me.CmdExport = New System.Windows.Forms.Button
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
        Me.GroupBox1.Location = New System.Drawing.Point(8, 64)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(946, 408)
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
        Me.DataGridCOM.Size = New System.Drawing.Size(940, 389)
        Me.DataGridCOM.TabIndex = 0
        '
        'CmdSave
        '
        Me.CmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdSave.Image = CType(resources.GetObject("CmdSave.Image"), System.Drawing.Image)
        Me.CmdSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdSave.Location = New System.Drawing.Point(800, 474)
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
        Me.CmdClose.Location = New System.Drawing.Point(872, 474)
        Me.CmdClose.Name = "CmdClose"
        Me.CmdClose.Size = New System.Drawing.Size(70, 56)
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
        Me.CmdEdit.Location = New System.Drawing.Point(728, 474)
        Me.CmdEdit.Name = "CmdEdit"
        Me.CmdEdit.Size = New System.Drawing.Size(72, 56)
        Me.CmdEdit.TabIndex = 3
        Me.CmdEdit.Text = "Edit"
        Me.CmdEdit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.CmdEdit.Visible = False
        '
        'CmbCompound
        '
        Me.CmbCompound.Enabled = False
        Me.CmbCompound.Location = New System.Drawing.Point(88, 40)
        Me.CmbCompound.Name = "CmbCompound"
        Me.CmbCompound.Size = New System.Drawing.Size(152, 21)
        Me.CmbCompound.TabIndex = 7
        Me.CmbCompound.Text = "Select"
        '
        'CheckBoxCompoud
        '
        Me.CheckBoxCompoud.Location = New System.Drawing.Point(16, 42)
        Me.CheckBoxCompoud.Name = "CheckBoxCompoud"
        Me.CheckBoxCompoud.Size = New System.Drawing.Size(72, 16)
        Me.CheckBoxCompoud.TabIndex = 8
        Me.CheckBoxCompoud.Text = "Compoud"
        '
        'CheckBoxGP
        '
        Me.CheckBoxGP.Location = New System.Drawing.Point(16, 10)
        Me.CheckBoxGP.Name = "CheckBoxGP"
        Me.CheckBoxGP.Size = New System.Drawing.Size(72, 16)
        Me.CheckBoxGP.TabIndex = 10
        Me.CheckBoxGP.Text = "Group"
        '
        'CmbGroup
        '
        Me.CmbGroup.Enabled = False
        Me.CmbGroup.Location = New System.Drawing.Point(88, 8)
        Me.CmbGroup.Name = "CmbGroup"
        Me.CmbGroup.Size = New System.Drawing.Size(152, 21)
        Me.CmbGroup.TabIndex = 9
        Me.CmbGroup.Text = "Select"
        '
        'CmdDelete
        '
        Me.CmdDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CmdDelete.Image = CType(resources.GetObject("CmdDelete.Image"), System.Drawing.Image)
        Me.CmdDelete.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdDelete.Location = New System.Drawing.Point(8, 472)
        Me.CmdDelete.Name = "CmdDelete"
        Me.CmdDelete.Size = New System.Drawing.Size(80, 56)
        Me.CmdDelete.TabIndex = 11
        Me.CmdDelete.Text = "Del"
        Me.CmdDelete.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'cmdAvtive
        '
        Me.cmdAvtive.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdAvtive.Image = CType(resources.GetObject("cmdAvtive.Image"), System.Drawing.Image)
        Me.cmdAvtive.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdAvtive.Location = New System.Drawing.Point(872, 8)
        Me.cmdAvtive.Name = "cmdAvtive"
        Me.cmdAvtive.Size = New System.Drawing.Size(80, 56)
        Me.cmdAvtive.TabIndex = 24
        Me.cmdAvtive.Text = "Active"
        Me.cmdAvtive.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ChkActive
        '
        Me.ChkActive.Location = New System.Drawing.Point(248, 42)
        Me.ChkActive.Name = "ChkActive"
        Me.ChkActive.Size = New System.Drawing.Size(64, 16)
        Me.ChkActive.TabIndex = 25
        Me.ChkActive.Text = " Active"
        '
        'CmdImport
        '
        Me.CmdImport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdImport.Image = CType(resources.GetObject("CmdImport.Image"), System.Drawing.Image)
        Me.CmdImport.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdImport.Location = New System.Drawing.Point(624, 474)
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
        Me.CmdExport.Location = New System.Drawing.Point(696, 474)
        Me.CmdExport.Name = "CmdExport"
        Me.CmdExport.Size = New System.Drawing.Size(72, 56)
        Me.CmdExport.TabIndex = 27
        Me.CmdExport.Text = "Export"
        Me.CmdExport.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'FrmCompound
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(962, 536)
        Me.Controls.Add(Me.CmdExport)
        Me.Controls.Add(Me.CmdImport)
        Me.Controls.Add(Me.ChkActive)
        Me.Controls.Add(Me.cmdAvtive)
        Me.Controls.Add(Me.CmdDelete)
        Me.Controls.Add(Me.CheckBoxGP)
        Me.Controls.Add(Me.CmbGroup)
        Me.Controls.Add(Me.CheckBoxCompoud)
        Me.Controls.Add(Me.CmbCompound)
        Me.Controls.Add(Me.CmdEdit)
        Me.Controls.Add(Me.CmdClose)
        Me.Controls.Add(Me.CmdSave)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmCompound"
        Me.Text = "Compound (Weight)"
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
    Private Sub FrmCompound_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Loadgroup()
        LoadCompound()
        LoadCOM()
        SetTotal() 'Set number of items
    End Sub
#End Region

#Region "Function_Load"
    Private Sub LoadCOM()
        Dim sb As New System.Text.StringBuilder()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sb.AppendLine(" SELECT * ")
        sb.AppendLine(" FROM (")
        sb.AppendLine("   SELECT  Seq,FinalCompound,CompCode,Revision,FinalCompound cc,Qty TQty, Active, Active AC")
        sb.AppendLine("   ,CompCode+','+Revision Code, '' MasterCode, isnull(Revision,'') MRev,'' RMcode,'' RMRev,null Qty,'' Unit,null Per") 'Header from table TBLCompound
        sb.AppendLine("   , '' as Revision_No, '' as FinalCompound_Code, '' as Seq_No, '' as Compound_Code")
        sb.AppendLine("   FROM TBLCompound")
        sb.AppendLine("   UNION")
        sb.AppendLine("   SELECT null Seq,c.Finalcompound,'' CompCode,'' Revision,'' cc,null TQty,'' Active, c.Active AC")
        sb.AppendLine("   ,m.MasterCode+','+m.Revision Code, m.MasterCode, m.Revision, m.RMCode, isnull(m.RmRevision,'') RMRev, m.Qty, m.Unit, m.Per") 'Detail from table TBLMaster
        sb.AppendLine("   , m.Revision as Revision_No, c.FinalCompound as FinalCompound_Code, c.Seq as Seq_No, m.MasterCode as Compound_Code")
        sb.AppendLine("   FROM TBLCompound c ")
        sb.AppendLine("   LEFT OUTER JOIN (")
        sb.AppendLine("     SELECT * ")
        sb.AppendLine("     FROM TBLMASTER ")
        sb.AppendLine("     WHERE Mastercode in ( SELECT compcode FROM TBLCompound)")
        sb.AppendLine("   ) m on c.CompCode+c.Revision = m.Mastercode+m.Revision ")
        sb.AppendLine(" ) aa ")
        sb.AppendLine(" ORDER BY Code, Seq DESC")
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
        Dim grdColStyle0_0 As New DataGridColoredLine2
        With grdColStyle0_0
            .HeaderText = "Stage"
            .MappingName = "Seq"
            .Width = 50
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle0 As New DataGridColoredLine2
        With grdColStyle0
            .HeaderText = "Group"
            .MappingName = "cc"
            .Width = 110
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle1 As New DataGridColoredLine2
        With grdColStyle1
            .HeaderText = "Compound"
            .MappingName = "CompCode"
            .Width = 110
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle1_1 As New DataGridColoredLine2
        With grdColStyle1_1
            .HeaderText = "Rev."
            .MappingName = "Revision"
            .Width = 80
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle2 As New DataGridColoredLine2
        With grdColStyle2
            .HeaderText = "R/M Material & Compound"
            .MappingName = "RMCode"
            .Width = 180
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle2_1 As New DataGridColoredLine2
        With grdColStyle2_1
            .HeaderText = "Rev ."
            .MappingName = "RMRev"
            .NullText = ""
            .Width = 85
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle3 As New DataGridColoredLine2
        With grdColStyle3
            .HeaderText = "Total"
            .MappingName = "TQty"
            .Width = 75
            .Format = "##,###,###.000"
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle4 As New DataGridColoredLine2
        With grdColStyle4
            .HeaderText = "Qty"
            .MappingName = "Qty"
            .Width = 75
            .Format = "##,###,###.000"
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle5 As New DataGridColoredLine2
        With grdColStyle5
            .HeaderText = "Unit"
            .MappingName = "Unit"
            .Width = 75
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle6 As New DataGridColoredLine2
        With grdColStyle6
            .HeaderText = " % of WT"
            .MappingName = "Per"
            .Width = 75
            .Format = "#0.000"
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle7 As New DataGridColoredLine2
        With grdColStyle7
            .HeaderText = "Active"
            .MappingName = "Active"
            .Width = 75
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        grdTableStyle1.GridColumnStyles.AddRange _
    (New DataGridColumnStyle() _
    {grdColStyle7, grdColStyle0, grdColStyle0_0, grdColStyle1,
    grdColStyle1_1, grdColStyle2, grdColStyle2_1,
     grdColStyle4, grdColStyle5, grdColStyle3, grdColStyle6})

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
    Sub LoadCompound()
        Dim dtComp As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        StrSQL = "SELECT  Code "
        StrSQL &= "  FROM  TblGroup  where Typecode = '03'"

        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtComp = New DataTable
            DA.Fill(dtComp)
        Catch
        Finally
        End Try
        dtComp.TableName = TBL_Comp
        GrdDVComp = dtComp.DefaultView
        '************************************
        CmbCompound.DisplayMember = "Code"
        CmbCompound.ValueMember = "Code"
        CmbCompound.DataSource = dtComp
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub Loadgroup()
        Dim dtGroup As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        StrSQL = "SELECT     distinct FinalCompound "
        StrSQL &= "  FROM         TBLCompound"

        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtGroup = New DataTable
            DA.Fill(dtGroup)
        Catch
        Finally
        End Try
        dtGroup.TableName = TBL_Group
        GrdDVGP = dtGroup.DefaultView
        '************************************
        CmbGroup.DisplayMember = "FinalCompound"
        CmbGroup.ValueMember = "FinalCompound"
        CmbGroup.DataSource = dtGroup
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
#End Region

#Region "Control Event"
    Private Sub CmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClose.Click
        Me.Close()
    End Sub

    'Private Sub CmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdEdit.Click
    '    Dim faddCompound As New FrmAddCompound
    '    faddCompound.CmdSave.Text = "Edit"
    '    faddCompound.TCompound = GrdDV.Item(oldrow).Row("FinalCompound")
    '    faddCompound.TCode = GrdDV.Item(oldrow).Row("CompCode")
    '    faddCompound.TRev = GrdDV.Item(oldrow).Row("Revision")
    '    faddCompound.TStep = GrdDV.Item(oldrow).Row("seq")

    '    faddCompound.TxtCode.Text = GrdDV.Item(oldrow).Row("CompCode")
    '    faddCompound.TxtRev.Text = GrdDV.Item(oldrow).Row("Revision")
    '    faddCompound.TxtCompound.Text = GrdDV.Item(oldrow).Row("FinalCompound")
    '    faddCompound.txtStep.Text = GrdDV.Item(oldrow).Row("seq")
    '    faddCompound.ComboBoxPigment.Text = GrdDV.Item(oldrow).Row("mastercode")
    '    If GrdDV.Item(oldrow).Row("Active") = 1 Then
    '        faddCompound.CheckBoxFinalCompound.Checked = True
    '    End If
    '    If GrdDV.Item(oldrow).Row("CompCode") = "" Then
    '        Exit Sub
    '    Else
    '        faddCompound.ShowDialog()
    '        LoadCompound()
    '        LoadCOM()
    '    End If

    '    If CheckBoxCompoud.Checked = True And CheckBoxGP.Checked = True Then
    '        GrdDV.RowFilter = " Code like'%" & CmbCompound.Text.Trim & _
    '                          "%' and  Finalcompound like'%" & StrData & "%'"
    '        DataGridCOM.DataSource = GrdDV
    '        CmbCompound.Enabled = True
    '    ElseIf CheckBoxCompoud.Checked = True And CheckBoxGP.Checked = False Then
    '        GrdDV.RowFilter = " Code like'%" & CmbCompound.Text.Trim & "%'"
    '        DataGridCOM.DataSource = GrdDV
    '        CmbCompound.Enabled = True
    '    ElseIf CheckBoxCompoud.Checked = False And CheckBoxGP.Checked = True Then
    '        GrdDV.RowFilter = " Finalcompound like'%" & StrData & "%'"
    '        DataGridCOM.DataSource = GrdDV
    '        CmbCompound.Enabled = True
    '    Else
    '        GrdDV.RowFilter = " "
    '        DataGridCOM.DataSource = GrdDV
    '        CmbCompound.Enabled = False
    '    End If
    'End Sub

    Private Sub CmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSave.Click
        Dim faddCompound As New FrmAddCompound
        faddCompound.CmdSave.Text = "Save"
        faddCompound.ShowDialog()
        LoadCompound()
        LoadCOM()
        Loadgroup()
        SelectDatachange()
    End Sub

    Private Sub DataGridCOM_CurrentCellChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGridCOM.CurrentCellChanged
        oldrow = DataGridCOM.CurrentCell.RowNumber
    End Sub

    Private Sub CmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdDelete.Click
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        msg = "Delete Compound :" & GrdDV.Item(oldrow).Row("CompCode") ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Compound"   ' Define title.

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            If ChkData() Or ChkData2() Then
                MsgBox("It's have Usage , Can't Delete. Please contact IS.", MsgBoxStyle.Information, "Delete R/M ")
            Else
                DelCompound()
                LoadCOM()
                SelectDatachange()
                oldrow = 0
            End If
        Else
            Exit Sub
        End If

    End Sub

    Private Sub cmdAvtive_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAvtive.Click
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        If GrdDV.Item(oldrow).Row("compcode") = "" Then
            Exit Sub
        End If
        msg = "Change Active Semi(Material) : " & GrdDV.Item(oldrow).Row("Compcode") _
        & "  Revision :" & GrdDV.Item(oldrow).Row("Revision") 'Define message.
        style = MsgBoxStyle.DefaultButton2 Or
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Semi(Material)"   ' Define title.
        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            UPCompound()
            LoadCOM()
            SelectDatachange()
        Else
            Exit Sub
        End If
    End Sub

    Private Sub CmdExport_Click(sender As Object, e As EventArgs) Handles CmdExport.Click
        Dim arrColumn As String() = System.Configuration.ConfigurationManager.AppSettings("EXP_EXCEL_COLUMN_MASTER_COMPOUND_WEIGHT").ToString().Split(New Char() {","c})
        Dim arrColumnHeader As String() = System.Configuration.ConfigurationManager.AppSettings("EXP_EXCEL_COLUMN_HEADER_MASTER_COMPOUND_WEIGHT").ToString().Split(New Char() {","c})
        ExcelLib.Export(Me, GrdDV, TBL_RM, arrColumn, arrColumnHeader)
    End Sub

    Private Sub CmdImport_Click(sender As Object, e As EventArgs) Handles CmdImport.Click
        Dim arrColumn As String() = System.Configuration.ConfigurationManager.AppSettings("IMP_EXCEL_COLUMN_MASTER_COMPOUND_WEIGHT").ToString().Split(New Char() {","c})
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
                        Dim chkSamePigmentCode As String = String.Empty
                        'Dim totalQty As Double = 0

                        '//Sort Data from Excel
                        dtRec.DefaultView.Sort = "FinalCompound_Code DESC, Compound_Code DESC, Revision_No DESC"
                        dtRec = dtRec.DefaultView.ToTable

                        '//Check RMCode on Master
                        If ChkRMCodeMaster(dtRec) = False Then
                            Exit Sub
                        End If

                        For i As Integer = 0 To dtRec.Rows.Count - 1

                            '//Case 0 : Start With Check FinalCompoundCode Between Excel and Grid
                            '//Case 1 : [SAME] FinalCompoundCode on Grid
                            '//Case 1.1 : [SAME] FinalCompoundCode, CompoundCode and Revision
                            '//Case 1.1.1 : [SAME] FinalCompoundCode, CompoundCode, Revision and RMCode
                            '//Case 1.1.1.1 : [SAME] FinalCompoundCode, CompoundCode, Revision, RMCode but [DIFFERENCE] QTY
                            '//Case 1.1.2 : [SAME] FinalCompoundCode, CompoundCode, Revision but [DIFFERENCE] RMCode
                            '//Case 1.2 : [SAME] FinalCompoundCode but [DIFFERENCE] CompoundCode and Revision
                            '//Case 2 : [NEW] FinalCompoundCode on Grid
                            '//Case 2.1 : [NEW] FinalCompoundCode but [SAME] CompoundCode and Revision
                            '//Case 2.2 : [NEW] FinalCompoundCode, CompoundCode and Revision

                            Dim strFinalCompoundCode As String = dtRec.Rows(i)("FinalCompound_Code").ToString().Trim()
                            If strFinalCompoundCode.Length > 0 Then
                                Dim strCompoundCode As String = dtRec.Rows(i)("Compound_Code").ToString().Trim()
                                Dim strRevision As String = dtRec.Rows(i)("Revision_No").ToString().Trim()
                                Dim strRMCode As String = dtRec.Rows(i)("RMCode").ToString().Trim()
                                Dim dblRMQty As Double
                                Dim intSeq As Integer = dtRec.Rows(i)("Seq_No")
                                Dim GridRow As DataRow()        '//Grid Data
                                Dim ExcelRow As DataRow()       '//Excel Data

                                If dtRec.Rows(i)("Qty").ToString.Length > 0 Then
                                    If Not Double.TryParse(dtRec.Rows(i)("Qty"), dblRMQty) Then
                                        Throw New System.Exception("Please input Qty data as Number")
                                    End If
                                Else
                                    Throw New System.Exception("Please input Qty data as Number")
                                End If

                                '//Case Special [Check Duplicate CompoundCode and Revision on Excel]
                                Dim ImportTable As New DataTable
                                ExcelRow = dtRec.Select("Compound_Code = '" & strCompoundCode & "' AND Revision_No = '" & strRevision & "'")
                                ImportTable = ExcelRow.CopyToDataTable
                                ImportTable = ImportTable.DefaultView.ToTable(True, "FinalCompound_Code")
                                If ImportTable.Rows.Count > 1 Then
                                    '//Error because of This Compound and Revision duplicate with Other FinalCompoundCode!!!
                                    MsgBox("This Compound : " & strCompoundCode & " and Revision : " & strRevision & " is Duplicate on Import File!!!", MsgBoxStyle.OkOnly, "Import Compound")
                                    Exit Sub
                                End If

                                '//Get Data on above row on Excel ------------------------
                                Dim chkSameFinalCompoundCodeBefore As String = String.Empty
                                Dim chkSameCompoundCodeBefore As String = String.Empty
                                Dim chkSameRevisionBefore As String = String.Empty
                                If i > 0 Then
                                    chkSameFinalCompoundCodeBefore = dtRec.Rows(i - 1)("FinalCompound_Code").ToString
                                    chkSameCompoundCodeBefore = dtRec.Rows(i - 1)("Compound_Code").ToString
                                    chkSameRevisionBefore = dtRec.Rows(i - 1)("Revision_No").ToString
                                Else
                                    chkSameFinalCompoundCodeBefore = ""
                                    chkSameCompoundCodeBefore = ""
                                    chkSameRevisionBefore = ""
                                End If
                                '//-------------------------------------------------------

                                '//Sum QTY each Compound Code and Revision
                                If strFinalCompoundCode <> chkSameFinalCompoundCodeBefore Or strCompoundCode <> chkSameCompoundCodeBefore Or strRevision <> chkSameRevisionBefore Then
                                    totalQty = 0
                                    ExcelRow = dtRec.Select("FinalCompound_Code = '" & strFinalCompoundCode & "' AND Compound_Code = '" & strCompoundCode & "' AND Revision_No = '" & strRevision & "'")
                                    For j As Integer = 0 To ExcelRow.Count - 1
                                        totalQty = totalQty + ExcelRow(j)("Qty")
                                    Next j
                                End If

                                '//Case 0 : Start With Check FinalCompoundCode Between Excel and Grid
                                GridRow = DT.Select("FinalCompound_Code = '" & strFinalCompoundCode & "'")

                                If GridRow.Count > 0 Or strFinalCompoundCode = chkSameFinalCompoundCodeBefore Then '//Case 1 : [SAME] FinalCompoundCode on Grid

                                    '//Next Check CompoundCode abd Revision with same FinalCompoundCode on Grid
                                    GridRow = DT.Select("FinalCompound_Code = '" & strFinalCompoundCode & "' AND Compound_Code = '" & strCompoundCode & "' AND Revision_No = '" & strRevision & "'")
                                    Dim chkCompANDRevBefore As Boolean = False
                                    If strCompoundCode = chkSameCompoundCodeBefore And strRevision = chkSameRevisionBefore Then
                                        chkCompANDRevBefore = True
                                    End If

                                    If GridRow.Count > 0 Or chkCompANDRevBefore = True Then '//Case 1.1 : [SAME] FinalCompoundCode, CompoundCode and Revision

                                        '//Next Check RMCode with same FinalCompound, CompoundCode and Revision on Grid
                                        GridRow = DT.Select("FinalCompound_Code = '" & strFinalCompoundCode & "' AND Compound_Code = '" & strCompoundCode & "' AND Revision_No = '" & strRevision & "' AND RMCode = '" & strRMCode & "'")

                                        If GridRow.Count > 0 AndAlso CDbl(GridRow(0)("Qty")) <> dblRMQty Then '//Case 1.1.1 : [SAME] FinalCompoundCode, CompoundCode, Revision and RMCode

                                            '//Case 1.1.1.1 : [SAME] FinalCompoundCode, CompoundCode, Revision, RMCode but [DIFFERENCE] QTY
                                            '//Update TBLMASTER
                                            sb.Clear()
                                            sb.AppendLine(" Update TBLMASTER")
                                            sb.AppendLine(" Set ")
                                            sb.AppendLine(" Qty = '" & dblRMQty & "'")
                                            sb.AppendLine(" Where MasterCode = '" & strCompoundCode & "' AND Revision = '" & strRevision & "' AND RMCode = '" & strRMCode & "' ")
                                            StrSQL = sb.ToString()
                                            cmSQL.CommandText = StrSQL
                                            cmSQL.ExecuteNonQuery()

                                            '//Update All Per in TBLMASTER***********
                                            sb.Clear()
                                            sb.AppendLine(" Update TBLMASTER")
                                            sb.AppendLine(" Set ")
                                            sb.AppendLine(" Per = Qty*(100/" & totalQty & ")")
                                            sb.AppendLine(" Where MasterCode = '" & strCompoundCode & "' AND Revision = '" & strRevision & "' ")
                                            StrSQL = sb.ToString()
                                            cmSQL.CommandText = StrSQL
                                            cmSQL.ExecuteNonQuery()

                                            '//Update TBLCompound [Qty(totalQty), DateUp(strDate)]
                                            sb.Clear()
                                            sb.AppendLine(" Update TBLCompound")
                                            sb.AppendLine(" Set ")
                                            sb.AppendLine(" Qty = '" & totalQty & "'")
                                            sb.AppendLine(" , Seq = " & intSeq)
                                            sb.AppendLine(" , Dateup = '" & strDate & "'")
                                            sb.AppendLine(" Where FinalCompound = '" & strFinalCompoundCode & "' AND CompCode = '" & strCompoundCode & "' AND Revision = '" & strRevision & "' ")
                                            'StrSQL = sb.ToString()
                                            'cmSQL.CommandText = StrSQL
                                            'cmSQL.ExecuteNonQuery()

                                            sb.AppendLine(" ")

                                            '//Update TBLConvert [SQty(totalQty)]
                                            sb.AppendLine(" Update TblConvert")
                                            sb.AppendLine(" Set ")
                                            sb.AppendLine(" SQty = '" & totalQty & "'")
                                            sb.AppendLine(" Where Final = '" & strFinalCompoundCode & "' AND Code = '" & strCompoundCode & "' AND Rev = '" & strRevision & "' AND Type = '03'")

                                            StrSQL = sb.ToString()
                                            cmSQL.CommandText = StrSQL
                                            cmSQL.ExecuteNonQuery()

                                        Else '//Case 1.1.2 : [SAME] FinalCompoundCode, CompoundCode, Revision but [DIFFERENCE] RMCode
                                            '//Insert TBLMaster
                                            sb.Clear()
                                            sb.AppendLine(" Insert  TBLMASTER ")
                                            sb.AppendLine(" Values (")
                                            sb.AppendLine(" '" & strCompoundCode & "', ")               'Column MasterCode
                                            sb.AppendLine(" '" & strRevision & "', ")                   'Column Revision
                                            sb.AppendLine(" '" & strRMCode & "', ")                     'Column RMCode
                                            sb.AppendLine(" NULL , ")                                   'Column RmRevision
                                            sb.AppendLine(" '" & dblRMQty & "', ")                      'Column Qty
                                            sb.AppendLine("'KG', ")                                     'Column Unit
                                            sb.AppendLine(" '" & ((dblRMQty * 100) / totalQty) & "'")   'Column Per
                                            sb.AppendLine(" )")
                                            StrSQL = sb.ToString()
                                            cmSQL.CommandText = StrSQL
                                            cmSQL.ExecuteNonQuery()

                                            '//Update All Per in TBLMASTER***********
                                            sb.Clear()
                                            sb.AppendLine(" Update TBLMASTER")
                                            sb.AppendLine(" Set ")
                                            sb.AppendLine(" Per = Qty*(100/" & totalQty & ")")
                                            sb.AppendLine(" Where MasterCode = '" & strCompoundCode & "' AND Revision = '" & strRevision & "' ")
                                            StrSQL = sb.ToString()
                                            cmSQL.CommandText = StrSQL
                                            cmSQL.ExecuteNonQuery()

                                            '//Update TBLPigment [Qty(totalQty), DateUp(strDate)]
                                            sb.Clear()
                                            sb.AppendLine(" Update TBLCompound")
                                            sb.AppendLine(" Set ")
                                            sb.AppendLine(" Qty = '" & totalQty & "'")
                                            sb.AppendLine(" , Dateup = '" & strDate & "'")
                                            sb.AppendLine(" Where FinalCompound = '" & strFinalCompoundCode & "' AND CompCode = '" & strCompoundCode & "' AND Revision = '" & strRevision & "' ")

                                            sb.AppendLine(" ")

                                            '//Update TBLConvert [SQty(totalQty)]
                                            sb.AppendLine(" Update TblConvert")
                                            sb.AppendLine(" Set ")
                                            sb.AppendLine(" SQty = '" & totalQty & "'")
                                            sb.AppendLine(" Where Final = '" & strFinalCompoundCode & "' AND Code = '" & strCompoundCode & "' AND Rev = '" & strRevision & "' AND Type = '03'")

                                            StrSQL = sb.ToString()
                                            cmSQL.CommandText = StrSQL
                                            cmSQL.ExecuteNonQuery()

                                        End If

                                    Else '//Case 1.2 : [SAME] FinalCompoundCode but [DIFFERENCE] CompoundCode and Revision

                                        '//Insert TBLGroup
                                        sb.Clear()
                                        sb.AppendLine(" Insert  TBLGroup ")
                                        sb.AppendLine(" Values (")
                                        sb.AppendLine("'03' , ")                        'Column TypeCode
                                        sb.AppendLine("'" & strCompoundCode & "'")      'Column Code
                                        sb.AppendLine(" )")

                                        sb.AppendLine(" ")

                                        '//Insert TBLCompound
                                        sb.AppendLine(" Insert TBLCompound ")
                                        sb.AppendLine(" Values (")
                                        sb.AppendLine(intSeq & ", ")                        'Column Seq
                                        sb.AppendLine(" '" & strFinalCompoundCode & "', ")  'Column FinalCompound
                                        sb.AppendLine(" '" & strCompoundCode & "', ")       'Column CompCode
                                        sb.AppendLine(" '" & strRevision & "' , ")          'Column Revision
                                        sb.AppendLine(" NULL , ")                           'Column RHC
                                        sb.AppendLine(" NULL , ")                           'Column PER
                                        sb.AppendLine(" '" & totalQty & "' , ")             'Column Qty
                                        sb.AppendLine(" '0' , ")                            'Column Active
                                        sb.AppendLine(" '" & strDate & "' ")                'Column Dateup
                                        sb.AppendLine(" )")

                                        sb.AppendLine(" ")

                                        '//Insert TblConvert
                                        sb.AppendLine(" Insert  TblConvert ")
                                        sb.AppendLine(" Values (")
                                        sb.AppendLine("'02' , ")                            'Column Type
                                        sb.AppendLine(" '" & strFinalCompoundCode & "', ")  'Column Final
                                        sb.AppendLine(" '" & strCompoundCode & "', ")       'Column Code
                                        sb.AppendLine(" '" & strRevision & "' , ")          'Column Rev
                                        sb.AppendLine("'BT' , ")                            'Column UnitBig
                                        sb.AppendLine("'KG' , ")                            'Column UnitSmall
                                        sb.AppendLine("'1' , ")                             'Column BQty
                                        sb.AppendLine(" '" & totalQty & "' ")               'Column SQty
                                        sb.AppendLine(" )")

                                        sb.AppendLine(" ")

                                        '//Insert TblMaster
                                        sb.AppendLine(" Insert  TBLMASTER ")
                                        sb.AppendLine(" Values (")
                                        sb.AppendLine(" '" & strCompoundCode & "', ")        'Column MasterCode
                                        sb.AppendLine(" '" & strRevision & "' , ")           'Column Revision
                                        sb.AppendLine(" '" & strRMCode & "' , ")                 'Column RMCode
                                        sb.AppendLine(" NULL , ")                        'Column RmRevision
                                        sb.AppendLine(" '" & dblRMQty & "', ")                  'Column Qty
                                        sb.AppendLine(" 'KG', ")                                     'Column Unit
                                        sb.AppendLine(" '" & ((dblRMQty * 100) / totalQty) & "'")   'Column Per
                                        sb.AppendLine(" )")

                                        StrSQL = sb.ToString()
                                        cmSQL.CommandText = StrSQL
                                        cmSQL.ExecuteNonQuery()

                                    End If
                                Else '//Case 2 : [NEW] FinalCompoundCode on Grid

                                    '//Next check Duplicate CompoundCode and Revision on Grid
                                    GridRow = DT.Select("Compound_Code = '" & strCompoundCode & "' AND Revision_No = '" & strRevision & "'")

                                    If GridRow.Count > 0 Then '//Case 2.1 : [NEW] FinalCompoundCode but [SAME] CompoundCode and Revision [IN Other] FinalCompoundCode

                                        '//Error because of This Compound and Revision duplicate with Other FinalCompoundCode!!!
                                        MsgBox("This Compound : " & strCompoundCode & " and Revision : " & strRevision & " is Duplicate on Other Group!!!", MsgBoxStyle.OkOnly, "Import Compound")
                                        Exit Sub

                                    Else '//Case 2.2 : [NEW] FinalCompoundCode, CompoundCode and Revision

                                        '//Insert TBLGroup
                                        sb.Clear()
                                        sb.AppendLine(" Insert  TBLGroup ")
                                        sb.AppendLine(" Values (")
                                        sb.AppendLine("'03' , ")                        'Column TypeCode
                                        sb.AppendLine("'" & strCompoundCode & "'")      'Column Code
                                        sb.AppendLine(" )")

                                        sb.AppendLine(" ")

                                        '//Insert TBLCompound
                                        sb.AppendLine(" Insert TBLCompound ")
                                        sb.AppendLine(" Values (")
                                        sb.AppendLine(intSeq & ", ")                        'Column Seq
                                        sb.AppendLine(" '" & strFinalCompoundCode & "', ")  'Column FinalCompound
                                        sb.AppendLine(" '" & strCompoundCode & "', ")       'Column CompCode
                                        sb.AppendLine(" '" & strRevision & "' , ")          'Column Revision
                                        sb.AppendLine(" NULL , ")                           'Column RHC
                                        sb.AppendLine(" NULL , ")                           'Column PER
                                        sb.AppendLine(" '" & totalQty & "' , ")             'Column Qty
                                        sb.AppendLine(" '0' , ")                            'Column Active
                                        sb.AppendLine(" '" & strDate & "' ")                'Column Dateup
                                        sb.AppendLine(" )")

                                        sb.AppendLine(" ")

                                        '//Insert TblConvert
                                        sb.AppendLine(" Insert  TblConvert ")
                                        sb.AppendLine(" Values (")
                                        sb.AppendLine("'02' , ")                            'Column Type
                                        sb.AppendLine(" '" & strFinalCompoundCode & "', ")  'Column Final
                                        sb.AppendLine(" '" & strCompoundCode & "', ")       'Column Code
                                        sb.AppendLine(" '" & strRevision & "' , ")          'Column Rev
                                        sb.AppendLine("'BT' , ")                            'Column UnitBig
                                        sb.AppendLine("'KG' , ")                            'Column UnitSmall
                                        sb.AppendLine("'1' , ")                             'Column BQty
                                        sb.AppendLine(" '" & totalQty & "' ")               'Column SQty
                                        sb.AppendLine(" )")

                                        sb.AppendLine(" ")

                                        '//Insert TblMaster
                                        sb.AppendLine(" Insert  TBLMASTER ")
                                        sb.AppendLine(" Values (")
                                        sb.AppendLine(" '" & strCompoundCode & "', ")        'Column MasterCode
                                        sb.AppendLine(" '" & strRevision & "' , ")           'Column Revision
                                        sb.AppendLine(" '" & strRMCode & "' , ")                 'Column RMCode
                                        sb.AppendLine(" NULL , ")                        'Column RmRevision
                                        sb.AppendLine(" '" & dblRMQty & "', ")                  'Column Qty
                                        sb.AppendLine(" 'KG', ")                                     'Column Unit
                                        sb.AppendLine(" '" & ((dblRMQty * 100) / totalQty) & "'")   'Column Per
                                        sb.AppendLine(" )")

                                        StrSQL = sb.ToString()
                                        cmSQL.CommandText = StrSQL
                                        cmSQL.ExecuteNonQuery()

                                    End If
                                End If
                            End If
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

            LoadCOM() 'ReQuery and set datagrid
            frmOverlay.Dispose()
        End If 'If importDialog.ShowDialog() = Windows.Forms.DialogResult.OK
    End Sub
#End Region

#Region "Import"
    Private Function ChkRMCodeMaster(ByVal ImportTable As DataTable) As Boolean
        Dim cnSQLRM As SqlConnection
        Dim cmSQLRM As SqlCommand
        Dim strSQL As String = String.Empty
        Dim ret As Boolean = False
        Dim strRmcodeBefore As String = String.Empty
        Dim distinctImportTabale As New DataTable

        Try
            ImportTable.DefaultView.Sort = "RMCode DESC"
            ImportTable = ImportTable.DefaultView.ToTable
            ImportTable = ImportTable.DefaultView.ToTable(True, "RMCode")
            For x As Integer = 0 To ImportTable.Rows.Count - 1
                Dim rmCode As String = ImportTable.Rows(x)("RMCode").ToString().Trim()
                strSQL = ""

                If x = 0 Then
                    strRmcodeBefore = ""
                Else
                    strRmcodeBefore = ImportTable.Rows(x - 1)("RMCode").ToString().Trim()
                End If
                If rmCode.Length > 0 Then
                    If rmCode <> strRmcodeBefore Then
                        strSQL &= " SELECT count(*) "
                        strSQL &= " FROM TBLRM "
                        strSQL &= " WHERE RMcode  = '" & rmCode & "'"
                        cnSQLRM = New SqlConnection(C1.Strcon)
                        cnSQLRM.Open()
                        cmSQLRM = New SqlCommand(strSQL, cnSQLRM)
                        Dim i As Long = cmSQLRM.ExecuteScalar()
                        If i = 0 Then
                            cmSQLRM.Dispose()
                            cnSQLRM.Dispose()
                            Throw New System.Exception("This RM Code '" & rmCode & "' have no data on RM Master")
                        Else
                            cmSQLRM.Dispose()
                            cnSQLRM.Dispose()
                        End If
                    End If
                End If
            Next x

            ret = True
        Catch Exp As SqlException
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "SQL Error")
        Catch Exp As Exception
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "General Error")
        Finally
            cnSQLRM.Close()
            cmSQLRM.Dispose()
            cnSQLRM.Dispose()
        End Try

        Return ret
    End Function
#End Region

#Region "SelectData"
    Private Sub CheckBoxCompoud_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxCompoud.CheckedChanged
        If CheckBoxCompoud.Checked Then
            CmbCompound.Enabled = True
        Else
            CmbCompound.Enabled = False
        End If
        SelectDatachange()
    End Sub

    Private Sub ChkActive_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkActive.CheckedChanged
        SelectDatachange()
    End Sub

    Private Sub CmbCompound_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCompound.SelectedIndexChanged
        SelectDatachange()
    End Sub

    Private Sub CheckBoxGP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxGP.CheckedChanged
        If CheckBoxGP.Checked Then
            CmbGroup.Enabled = True
        Else
            CmbGroup.Enabled = False
        End If
        SelectDatachange()
    End Sub

    Private Sub CmbGroup_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbGroup.SelectedIndexChanged
        SelectDatachange()
    End Sub

    Sub SelectDatachange()
        Dim strView As String

        If CheckBoxCompoud.Checked = True And CheckBoxGP.Checked = True Then
            strView = " Code like'%" & CmbCompound.Text.Trim &
                              "%' and  Finalcompound like'%" & CmbGroup.Text.Trim & "%'"
        ElseIf CheckBoxCompoud.Checked = True And CheckBoxGP.Checked = False Then
            strView = " Code like'%" & CmbCompound.Text.Trim & "%'"
        ElseIf CheckBoxCompoud.Checked = False And CheckBoxGP.Checked = True Then
            strView = " Finalcompound like'%" & CmbGroup.Text.Trim & "%'"
        Else
            If ChkActive.Checked Then
                strView = "  ac  = '1' "
            Else
                strView = " "
            End If
            GrdDV.RowFilter = strView
            DataGridCOM.DataSource = GrdDV

            Exit Sub
        End If

        If ChkActive.Checked Then
            strView &= "  and ac  = '1' "
        Else
            strView &= " "
        End If
        GrdDV.RowFilter = strView
        DataGridCOM.DataSource = GrdDV

        SetTotal() 'Set number of items
    End Sub

#End Region

#Region "DelCompound"
    Private Function ChkData() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String
        Try
            strSQL = " select count(*) from TblCompound "
            strSQL &= " where Active  = '1'"
            strSQL &= " and  Compcode   = '" & GrdDV.Item(oldrow).Row("Compcode") & "'"
            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar
            If i = 0 Then
                ChkData = False
            Else
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
    Private Function ChkData2() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String
        Try
            strSQL = " select count(*) from TblMaster "
            strSQL &= " where RMCode  = '" & GrdDV.Item(oldrow).Row("Compcode") & "'"
            strSQL &= " and RMRevision  = '" & GrdDV.Item(oldrow).Row("Revision") & "'"
            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar
            If i <> 0 Then
                ChkData2 = True
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
        Dim strSQL As String
        Try
            strSQL = " select count(*) from TblCompound "
            strSQL &= " where CompCode  = '" & GrdDV.Item(oldrow).Row("CompCode") & "'"
            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar
            If i = 1 Then
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
    Sub DelCompound()
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Try
            strSQL = " Delete TblCompound"
            strSQL &= " where CompCode = '" & GrdDV.Item(oldrow).Row("CompCode") & "'"
            strSQL &= " and Revision = '" & GrdDV.Item(oldrow).Row("Revision") & "'"
            strSQL &= "  "
            strSQL &= " Delete TblMaster"
            strSQL &= " where Mastercode = '" & GrdDV.Item(oldrow).Row("CompCode") & "'"
            strSQL &= " and Revision = '" & GrdDV.Item(oldrow).Row("Revision") & "'"
            strSQL &= "  "
            strSQL &= " Delete TblRHCDtl"
            strSQL &= " where Mastercode = '" & GrdDV.Item(oldrow).Row("CompCode") & "'"
            strSQL &= " and Revision = '" & GrdDV.Item(oldrow).Row("Revision") & "'"
            strSQL &= "  "
            If ChkDel() Then
                strSQL &= " Delete TblGroup"
                strSQL &= " where Code = '" & GrdDV.Item(oldrow).Row("CompCode") & "'"
            Else
            End If
            strSQL &= "  "
            strSQL &= " Delete TblConvert"
            strSQL &= " where code = '" & GrdDV.Item(oldrow).Row("CompCode") & "'"
            strSQL &= " and Rev = '" & GrdDV.Item(oldrow).Row("Revision") & "'"
            strSQL &= " and UnitBig = 'BT'"
            strSQL &= "  "

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

    Sub UPCompound()
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Try
            strSQL = " Update TblCompound"
            strSQL &= " set Active = 0"
            strSQL &= " where FinalCompound = '" & GrdDV.Item(oldrow).Row("Finalcompound") & "'"
            strSQL &= " "
            strSQL &= " Update TblCompound"
            strSQL &= " set Active = 1"
            strSQL &= " where CompCode = '" & GrdDV.Item(oldrow).Row("compcode") & "'"
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

    Private Sub SetTotal()
        'Set total
        'Format: Form Text - xxx item(s)
        Dim frmTitle As String() = Me.Text.Split(New Char() {"-"c})
        Me.Text = frmTitle(0) & "- " & GrdDV.Count & " item(s)"
    End Sub
End Class
