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

Public Class FrmGreenTire

#Region "Declare"
    Inherits System.Windows.Forms.Form
    Dim GrdDV As New DataView
    Protected Const TBL_RM As String = "TBL_RM"
    Dim GrdDVPreSemi As New DataView
    Protected Const TBL_PreSemi As String = "TBL_PreSemi"
    Dim GrdDVType As New DataView
    Protected Const TBL_Type As String = "TBL_Type"
    Protected Const TBL_GT As String = "TBL_GT"
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
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents CmdDel As System.Windows.Forms.Button
    Friend WithEvents CheckBoxTire As System.Windows.Forms.CheckBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents CmbTire As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdView As System.Windows.Forms.Button
    Friend WithEvents txtsize As System.Windows.Forms.TextBox
    Friend WithEvents CHKActive As System.Windows.Forms.CheckBox
    Friend WithEvents cmdActive As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmGreenTire))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.DataGridCOM = New System.Windows.Forms.DataGrid()
        Me.CmdSave = New System.Windows.Forms.Button()
        Me.CmdClose = New System.Windows.Forms.Button()
        Me.CmdEdit = New System.Windows.Forms.Button()
        Me.CmbTire = New System.Windows.Forms.ComboBox()
        Me.CheckBoxTire = New System.Windows.Forms.CheckBox()
        Me.CmdDel = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtsize = New System.Windows.Forms.TextBox()
        Me.cmdView = New System.Windows.Forms.Button()
        Me.CHKActive = New System.Windows.Forms.CheckBox()
        Me.cmdActive = New System.Windows.Forms.Button()
        Me.CmdImport = New System.Windows.Forms.Button()
        Me.CmdExport = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridCOM, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.GroupBox1.Size = New System.Drawing.Size(1068, 575)
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
        Me.DataGridCOM.Size = New System.Drawing.Size(1062, 556)
        Me.DataGridCOM.TabIndex = 0
        '
        'CmdSave
        '
        Me.CmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdSave.Image = CType(resources.GetObject("CmdSave.Image"), System.Drawing.Image)
        Me.CmdSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdSave.Location = New System.Drawing.Point(844, 649)
        Me.CmdSave.Name = "CmdSave"
        Me.CmdSave.Size = New System.Drawing.Size(80, 56)
        Me.CmdSave.TabIndex = 1
        Me.CmdSave.Text = "Add"
        Me.CmdSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdClose
        '
        Me.CmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdClose.Image = CType(resources.GetObject("CmdClose.Image"), System.Drawing.Image)
        Me.CmdClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdClose.Location = New System.Drawing.Point(996, 649)
        Me.CmdClose.Name = "CmdClose"
        Me.CmdClose.Size = New System.Drawing.Size(75, 56)
        Me.CmdClose.TabIndex = 2
        Me.CmdClose.Text = "Close"
        Me.CmdClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdEdit
        '
        Me.CmdEdit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdEdit.Image = CType(resources.GetObject("CmdEdit.Image"), System.Drawing.Image)
        Me.CmdEdit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdEdit.Location = New System.Drawing.Point(924, 649)
        Me.CmdEdit.Name = "CmdEdit"
        Me.CmdEdit.Size = New System.Drawing.Size(75, 56)
        Me.CmdEdit.TabIndex = 3
        Me.CmdEdit.Text = "Edit"
        Me.CmdEdit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmbTire
        '
        Me.CmbTire.Enabled = False
        Me.CmbTire.Location = New System.Drawing.Point(128, 14)
        Me.CmbTire.Name = "CmbTire"
        Me.CmbTire.Size = New System.Drawing.Size(152, 21)
        Me.CmbTire.TabIndex = 7
        Me.CmbTire.Text = "Select"
        '
        'CheckBoxTire
        '
        Me.CheckBoxTire.Location = New System.Drawing.Point(16, 16)
        Me.CheckBoxTire.Name = "CheckBoxTire"
        Me.CheckBoxTire.Size = New System.Drawing.Size(56, 16)
        Me.CheckBoxTire.TabIndex = 8
        Me.CheckBoxTire.Text = "Tire"
        '
        'CmdDel
        '
        Me.CmdDel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CmdDel.Image = CType(resources.GetObject("CmdDel.Image"), System.Drawing.Image)
        Me.CmdDel.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdDel.Location = New System.Drawing.Point(8, 647)
        Me.CmdDel.Name = "CmdDel"
        Me.CmdDel.Size = New System.Drawing.Size(80, 56)
        Me.CmdDel.TabIndex = 21
        Me.CmdDel.Text = "Delete"
        Me.CmdDel.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(72, 8)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(48, 32)
        Me.PictureBox1.TabIndex = 22
        Me.PictureBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 16)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "Size"
        Me.Label1.Visible = False
        '
        'txtsize
        '
        Me.txtsize.Location = New System.Drawing.Point(128, 46)
        Me.txtsize.Name = "txtsize"
        Me.txtsize.Size = New System.Drawing.Size(152, 20)
        Me.txtsize.TabIndex = 24
        Me.txtsize.Visible = False
        '
        'cmdView
        '
        Me.cmdView.Image = CType(resources.GetObject("cmdView.Image"), System.Drawing.Image)
        Me.cmdView.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdView.Location = New System.Drawing.Point(352, 10)
        Me.cmdView.Name = "cmdView"
        Me.cmdView.Size = New System.Drawing.Size(72, 56)
        Me.cmdView.TabIndex = 25
        Me.cmdView.Text = "View"
        Me.cmdView.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CHKActive
        '
        Me.CHKActive.Location = New System.Drawing.Point(288, 16)
        Me.CHKActive.Name = "CHKActive"
        Me.CHKActive.Size = New System.Drawing.Size(56, 16)
        Me.CHKActive.TabIndex = 26
        Me.CHKActive.Text = "Avtice"
        '
        'cmdActive
        '
        Me.cmdActive.Image = CType(resources.GetObject("cmdActive.Image"), System.Drawing.Image)
        Me.cmdActive.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdActive.Location = New System.Drawing.Point(424, 10)
        Me.cmdActive.Name = "cmdActive"
        Me.cmdActive.Size = New System.Drawing.Size(72, 56)
        Me.cmdActive.TabIndex = 27
        Me.cmdActive.Text = "Active"
        Me.cmdActive.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdImport
        '
        Me.CmdImport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdImport.Image = CType(resources.GetObject("CmdImport.Image"), System.Drawing.Image)
        Me.CmdImport.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdImport.Location = New System.Drawing.Point(646, 649)
        Me.CmdImport.Name = "CmdImport"
        Me.CmdImport.Size = New System.Drawing.Size(80, 56)
        Me.CmdImport.TabIndex = 28
        Me.CmdImport.Text = "Import"
        Me.CmdImport.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdExport
        '
        Me.CmdExport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdExport.Image = CType(resources.GetObject("CmdExport.Image"), System.Drawing.Image)
        Me.CmdExport.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdExport.Location = New System.Drawing.Point(726, 649)
        Me.CmdExport.Name = "CmdExport"
        Me.CmdExport.Size = New System.Drawing.Size(80, 56)
        Me.CmdExport.TabIndex = 29
        Me.CmdExport.Text = "Export"
        Me.CmdExport.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'FrmGreenTire
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1084, 711)
        Me.Controls.Add(Me.CmdExport)
        Me.Controls.Add(Me.CmdImport)
        Me.Controls.Add(Me.CHKActive)
        Me.Controls.Add(Me.cmdView)
        Me.Controls.Add(Me.txtsize)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.CmdDel)
        Me.Controls.Add(Me.CheckBoxTire)
        Me.Controls.Add(Me.CmbTire)
        Me.Controls.Add(Me.CmdEdit)
        Me.Controls.Add(Me.CmdClose)
        Me.Controls.Add(Me.CmdSave)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmdActive)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(1100, 750)
        Me.Name = "FrmGreenTire"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Tire"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataGridCOM, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

#Region "CONSTANT"
    Dim DT As New DataTable
    Dim StrSQL As String
    Dim oldrow As Integer

#End Region

#Region "Form Event"
    Private Sub FrmGreenTire_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadTire()
        LoadTireGroup()
        'If CheckBoxType.Checked = False Then
        '    GrdDV.RowFilter = " "
        '    DataGridCOM.DataSource = GrdDV
        'End If

        'If CheckBoxTire.Checked = False Then
        '    GrdDV.RowFilter = " "
        '    DataGridCOM.DataSource = GrdDV
        'End If

    End Sub
#End Region

#Region "Function_Load"
    Private Sub LoadTire()
        Dim sb As New System.Text.StringBuilder()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sb.AppendLine("SELECT * ")
        sb.AppendLine("FROM (")
        sb.AppendLine("  SELECT final,TireSize,Round(Qty,1) TQty,")
        sb.AppendLine("  substring(DateUp,7,2)+'/'+substring(DateUp,5,2)+'/'")
        sb.AppendLine("  +substring(DateUp,1,4) dateup,TireSize TSize,Tirecode,Rev,null MaterialName")
        sb.AppendLine("  ,null Semicode,null Length,null number,null QTU,null Unit,Remark,Active,Active AC,'' as EachGreenTire, '' as EachRevision, '' as EachBSJ, '' as MaterialCode ")
        sb.AppendLine("  FROM TblGtHdr")
        sb.AppendLine("  UNION")
        sb.AppendLine("  SELECT null final,null TireSize,null TQty,null dateup,tiresize TSize,dt.Tirecode,dt.Rev,MaterialName")
        sb.AppendLine("  ,isnull(Semicode,'No Use') Semicode,Length,number, round(QTU,3) Qty, Unit ,null Remark ,null Active,Active AC, dt.TireCode as EachGreenTire, dt.Rev as EachRevision, TireSize as EachBSJ, MaterialCode ")
        sb.AppendLine("  FROM TblGtDtl dt")
        sb.AppendLine("  LEFT OUTER JOIN TBLTypeMaterial tm on dt.MaterialType = tm.MaterialCode")
        sb.AppendLine("  LEFT OUTER JOIN TBLGTHdr hd on dt.tirecode+dt.Rev = hd.Tirecode+hd.Rev")
        sb.AppendLine(") Tire")
        sb.AppendLine("ORDER BY Tirecode,Rev,Final DESC, MaterialName")
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

        Dim grdColStyle0 As New DataGridColoredLine2
        With grdColStyle0
            .HeaderText = "Tire"
            .MappingName = "final"
            .Width = 65
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle0_1 As New DataGridColoredLine2
        With grdColStyle0_1
            .HeaderText = "Rev"
            .MappingName = "Rev"
            .Width = 65
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle0_2 As New DataGridColoredLine2
        With grdColStyle0_2
            .HeaderText = "Active"
            .MappingName = "Active"
            .Width = 55
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle1 As New DataGridColoredLine2
        With grdColStyle1
            .HeaderText = "Size"
            .MappingName = "Tiresize"
            .NullText = ""
            .Width = 120
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle2 As New DataGridColoredLine2
        With grdColStyle2
            .HeaderText = "Material"
            .MappingName = "MaterialName"
            .NullText = ""
            .Width = 150
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle3 As New DataGridColoredLine2
        With grdColStyle3
            .HeaderText = "MaterialCode"
            .MappingName = "Semicode"
            .NullText = " "
            .Width = 95
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle5 As New DataGridColoredLine2
        With grdColStyle5
            .HeaderText = "Total Weight"
            .MappingName = "TQty"
            .NullText = ""
            .Width = 85
            .Format = "##,###,###.0"
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
        Dim grdColStyle9 As New DataGridColoredLine2
        With grdColStyle9
            .HeaderText = "Weight"
            .MappingName = "Qtu"
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
            .MappingName = "Number"
            .NullText = ""
            .Width = 45
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle13 As New DataGridColoredLine2
        With grdColStyle13
            .HeaderText = "Date"
            .MappingName = "DateUp"
            .NullText = ""
            .Width = 85
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle14 As New DataGridColoredLine2
        With grdColStyle14
            .HeaderText = "Remark"
            .MappingName = "Remark"
            .NullText = ""
            .Width = 85
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        grdTableStyle1.GridColumnStyles.AddRange _
    (New DataGridColumnStyle() _
    {grdColStyle0_2, grdColStyle0, grdColStyle0_1, grdColStyle1, grdColStyle2, grdColStyle3,
     grdColStyle12, grdColStyle11, grdColStyle9, grdColStyle6,
     grdColStyle5, grdColStyle13, grdColStyle14})

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
    Sub LoadTireGroup()
        Dim sb As New System.Text.StringBuilder()
        Dim dtPSemi As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sb.AppendLine("  SELECT Code,MaterialName")
        sb.AppendLine("  FROM  TblGroup g")
        sb.AppendLine("  LEFT OUTER JOIN (")
        sb.AppendLine("    SELECT  semiCode,MaterialName")
        sb.AppendLine("    FROM  TblSemi p")
        sb.AppendLine("    LEFT OUTER JOIN  TblTypeMaterial t on p.MaterialType = t.MaterialCode")
        sb.AppendLine("  ) semi on g.code = semi.semicode")
        sb.AppendLine("  WHERE Typecode = '06'")
        sb.AppendLine("  ORDER BY Code")
        StrSQL = sb.ToString()

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
        CmbTire.DisplayMember = "Code"
        CmbTire.ValueMember = "Code"
        CmbTire.DataSource = dtPSemi
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
#End Region

#Region "Control Event"
    Private Sub CmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClose.Click
        Me.Close()
    End Sub

    Private Sub CmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdEdit.Click
        Dim Strremark() As String
        Dim i As Integer
        Dim FAddGreenTire As New FrmAddGreenTire
        FAddGreenTire.Text = "Edit GreenTire"
        FAddGreenTire.CmdSave.Text = "Edit"

        If GrdDV.Item(oldrow).Row("Final").Equals(DBNull.Value) Or GrdDV.Item(oldrow).Row("Final").Equals(String.Empty) Then
            MessageBox.Show("Please select row which specify ""Tire"" value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        Else
            FAddGreenTire.TxtCode.Text = GrdDV.Item(oldrow).Row("Final")
            FAddGreenTire.TxtCode.Enabled = False
            FAddGreenTire.TxtRev.Text = GrdDV.Item(oldrow).Row("Rev")
            FAddGreenTire.CmbBSJCode.Enabled = False
            FAddGreenTire.BSJCode = GrdDV.Item(oldrow).Row("Tiresize")
            Strremark = Split(GrdDV.Item(oldrow).Row("Remark"), ",")
            FAddGreenTire.txtremark.Text = Strremark(0)
            FAddGreenTire.txtremark2.Text = Strremark(1)

            For i = 0 To 13
                If GrdDV.Item(oldrow + i).Row("MaterialName").Equals("BELT-1") Then
                    'Belt-1
                    FAddGreenTire.TxtB1_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                    FAddGreenTire.TxtB1_L.Text = GrdDV.Item(oldrow + i).Row("Length")
                    FAddGreenTire.B1code = GrdDV.Item(oldrow + i).Row("Semicode")
                End If
                If GrdDV.Item(oldrow + i).Row("MaterialName").Equals("BELT-2") Then
                    'Belt-2
                    FAddGreenTire.TxtB2_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                    FAddGreenTire.TxtB2_L.Text = GrdDV.Item(oldrow + i).Row("Length")
                    FAddGreenTire.B2code = GrdDV.Item(oldrow + i).Row("Semicode")
                End If
                If GrdDV.Item(oldrow + i).Row("MaterialName").Equals("BELT-3") Then
                    'Belt-3
                    FAddGreenTire.TxtB3_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                    FAddGreenTire.TxtB3_L.Text = GrdDV.Item(oldrow + i).Row("Length")
                    FAddGreenTire.B3code = GrdDV.Item(oldrow + i).Row("Semicode")
                End If
                If GrdDV.Item(oldrow + i).Row("MaterialName").Equals("BELT-4") Then
                    'Belt-4
                    FAddGreenTire.TxtB4_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                    FAddGreenTire.TxtB4_L.Text = GrdDV.Item(oldrow + i).Row("Length")
                    FAddGreenTire.B4code = GrdDV.Item(oldrow + i).Row("Semicode")
                End If
                If GrdDV.Item(oldrow + i).Row("MaterialName").Equals("BF (Upper,Lower,Center)") Then
                    'BF (Upper,Lower,Center)
                    FAddGreenTire.txtBF_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                    FAddGreenTire.BFcode = GrdDV.Item(oldrow + i).Row("Semicode")
                End If
                If GrdDV.Item(oldrow + i).Row("MaterialName").Equals("BODY PLY") Then
                    'Body Ply
                    FAddGreenTire.TxtBP_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                    FAddGreenTire.TxtBp_L.Text = GrdDV.Item(oldrow + i).Row("Length")
                    FAddGreenTire.Bpcode = GrdDV.Item(oldrow + i).Row("Semicode")
                End If
                If GrdDV.Item(oldrow + i).Row("MaterialName").Equals("CUSSION") Then
                    'Cussion
                    FAddGreenTire.TxtCu_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                    FAddGreenTire.TxtCU_L.Text = GrdDV.Item(oldrow + i).Row("Length")
                    FAddGreenTire.CUcode = GrdDV.Item(oldrow + i).Row("Semicode")
                End If
                If GrdDV.Item(oldrow + i).Row("MaterialName").Equals("FLIPPER") Then
                    'Flipper
                    If GrdDV.Item(oldrow + i).Row("Semicode") <> "No Use" Then
                        FAddGreenTire.CheckBoxFP.Checked = True
                        FAddGreenTire.TxtFP_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                        FAddGreenTire.TxtFP_L.Text = GrdDV.Item(oldrow + i).Row("Length")
                        FAddGreenTire.FPcode = GrdDV.Item(oldrow + i).Row("Semicode")
                    Else
                        'Nothing
                    End If
                End If
                If GrdDV.Item(oldrow + i).Row("MaterialName").Equals("INNERLINER") Then
                    'InnerLiner
                    FAddGreenTire.TxtIN_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                    FAddGreenTire.TxtIN_L.Text = GrdDV.Item(oldrow + i).Row("Length")
                    FAddGreenTire.INcode = GrdDV.Item(oldrow + i).Row("Semicode")
                End If
                If GrdDV.Item(oldrow + i).Row("MaterialName").Equals("Nylon CHAFER") Then
                    'Nylon Chafer
                    If GrdDV.Item(oldrow + i).Row("Semicode") <> "No Use" Then
                        FAddGreenTire.CheckBoxNy.Checked = True
                        FAddGreenTire.TxtNy_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                        FAddGreenTire.TxtNy_L.Text = GrdDV.Item(oldrow + i).Row("Length")
                        FAddGreenTire.NYcode = GrdDV.Item(oldrow + i).Row("Semicode")
                    Else
                        'Nothing
                    End If
                End If
                If GrdDV.Item(oldrow + i).Row("MaterialName").Equals("SIDE") Then
                    'Side
                    FAddGreenTire.TxtSD_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                    FAddGreenTire.TxtSD_L.Text = GrdDV.Item(oldrow + i).Row("Length")
                    FAddGreenTire.SDcode = GrdDV.Item(oldrow + i).Row("Semicode")
                End If
                If GrdDV.Item(oldrow + i).Row("MaterialName").Equals("TREAD") Then
                    'Tread
                    FAddGreenTire.TxtTT_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                    FAddGreenTire.TTcode = GrdDV.Item(oldrow + i).Row("Semicode")
                End If
                If GrdDV.Item(oldrow + i).Row("MaterialName").Equals("WIRE CHAFER") Then
                    'Wire Chafer
                    If GrdDV.Item(oldrow + i).Row("Semicode") <> "No Use" Then
                        FAddGreenTire.CheckBoxWf.Checked = True
                        FAddGreenTire.TxtWf_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                        FAddGreenTire.TxtWf_L.Text = GrdDV.Item(oldrow + i).Row("Length")
                        FAddGreenTire.Wfcode = GrdDV.Item(oldrow + i).Row("Semicode")
                    Else
                        'Nothing
                    End If

                    Exit For 'Exit loop because it is last record of Group
                End If
            Next i

            FAddGreenTire.ShowDialog()
            LoadTire()
            CheckBox()
            oldrow = 0
        End If
    End Sub

    Private Sub CmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSave.Click
        Dim ttype As String
        Dim FAddGreenTire As New FrmAddGreenTire
        ttype = CmbTire.Text.Trim()
        FAddGreenTire.CmdSave.Text = "Save"
        FAddGreenTire.ShowDialog()
        LoadTire()
        LoadTireGroup()
        CheckBox()
        oldrow = 0
        CmbTire.Text = ttype
    End Sub

    Private Sub DataGridCOM_CurrentCellChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGridCOM.CurrentCellChanged
        oldrow = DataGridCOM.CurrentCell.RowNumber
    End Sub

    Private Sub CmdExport_Click(sender As Object, e As EventArgs) Handles CmdExport.Click
        Dim arrColumn As String() = System.Configuration.ConfigurationManager.AppSettings("EXP_EXCEL_COLUMN_MASTER_GREENTIRE").ToString().Split(New Char() {","c})
        Dim arrColumnHeader As String() = System.Configuration.ConfigurationManager.AppSettings("EXP_EXCEL_COLUMN_HEADER_MASTER_GREENTIRE").ToString().Split(New Char() {","c})
        ExcelLib.Export(Me, GrdDV, TBL_RM, arrColumn, arrColumnHeader)
    End Sub

    Private Sub CmdImport_Click(sender As Object, e As EventArgs) Handles CmdImport.Click
        Dim arrColumn As String() = System.Configuration.ConfigurationManager.AppSettings("IMP_EXCEL_COLUMN_MASTER_GREENTIRE").ToString().Split(New Char() {","c})
        Dim importDialog As OpenFileDialog = New OpenFileDialog With {
            .Filter = System.Configuration.ConfigurationManager.AppSettings("DIALOG_FILE_EXT").ToString()
        }
        Dim dtRec As DataTable
        Dim sb As New System.Text.StringBuilder()
        Dim frmOverlay As New Form()
        Dim totalQty As Double = 0
        Dim QBF, QSide, QInnerLiner, QCussion, QTread, QBodyPly, QBelt1, QBelt2, QBelt3, QBelt4, QWireChafer, QNylonChafer, QFlipper As Double


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
            dtRec = ExcelLib.Import(importDialog.FileName, Me, GrdDV, TBL_GT, arrColumn)
            dtRec.Columns.Add("QTU", GetType(Double))

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
                        Dim dtTireSize As DataTable = GetTireSize() 'Table TBLTireSize

                        '//Sort Data from Excel
                        dtRec.DefaultView.Sort = "GreenTire DESC, Revision DESC, TypeMaterial DESC"
                        dtRec = dtRec.DefaultView.ToTable

                        '//**Check All Import Data that all data still in TBLCompound, TBLMaster and TBLRM
                        If ChkImportData_Correctly(dtRec, dtTypeMaterial, dtTireSize) = False Then
                            LoadTire() 'ReQuery and set datagrid
                            frmOverlay.Dispose()
                            Exit Sub
                        End If


                        For i As Integer = 0 To dtRec.Rows.Count - 1
                            Dim strGreentire As String = dtRec(i)("GreenTire").ToString().Trim()
                            Dim strRevision As String = dtRec(i)("Revision").ToString().Trim()
                            Dim strBSJ As String = dtRec(i)("BSJ").ToString().Trim()
                            Dim strRevisionBoss1st As String = dtRec(i)("RevisionBoss_1st").ToString().Trim()
                            Dim strRevisionBoss2nd As String = dtRec(i)("RevisionBoss_2nd").ToString().Trim()
                            Dim strTypeMaterial As String = dtRec.Rows(i)("TypeMaterial").ToString().Trim()
                            Dim GridRow As DataRow()        '//Grid Data
                            Dim ExcelRow As DataRow()       '//Excel Data

                            'Set Type Material
                            Dim arrTypeMatCode As DataRow() = dtTypeMaterial.Select("MaterialName = '" & strTypeMaterial & "'")
                            strTypeMaterial = arrTypeMatCode(0)("MaterialCode")

                            '//For Check Data from above row on import file.
                            Dim chkSameGreenTireBefore As String = String.Empty
                            Dim chkSameRevisionBefore As String = String.Empty
                            If i > 0 Then
                                chkSameGreenTireBefore = dtRec.Rows(i - 1)("GreenTire").ToString().Trim()
                                chkSameRevisionBefore = dtRec.Rows(i - 1)("Revision").ToString().Trim()
                            Else
                                chkSameGreenTireBefore = String.Empty
                                chkSameRevisionBefore = String.Empty
                            End If

                            GridRow = DT.Select("EachGreenTire = '" & strGreentire & "' AND EachRevision = '" & strRevision & "'")
                            If GridRow.Count > 0 Then '//Case Update
                                If strGreentire <> chkSameGreenTireBefore Or strRevision <> chkSameRevisionBefore Then

                                    '//Update TblGtDtl
                                    '// Tread and BF (Require, Need only Num) ------------------------------------------------------------------------------------------
                                    '// Tread
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'TREAD'")

                                    sb.Clear()
                                    sb.AppendLine(" Update TblGtDtl")
                                    sb.AppendLine(" Set ")
                                    sb.AppendLine(" Semicode = " & PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                    sb.AppendLine(" length = " & PrepareStr(ExcelRow(0)("Length")) & ", ")
                                    sb.AppendLine(" number = " & PrepareStr(ExcelRow(0)("Num")) & ", ")
                                    sb.AppendLine(" QTU = " & PrepareStr((ExcelRow(0)("QTU") * ExcelRow(0)("Num"))) & ", ")
                                    sb.AppendLine(" Dateup = " & PrepareStr(strDate))
                                    sb.AppendLine(" Where TireCode = '" & strGreentire & "' AND Rev = '" & strRevision & "' AND MaterialType = '13'")

                                    QTread = ExcelRow(0)("QTU")

                                    sb.AppendLine(" ")

                                    ''// BF
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'BF (Upper,Lower,Center)'")
                                    sb.AppendLine(" Update TblGtDtl")
                                    sb.AppendLine(" Set ")
                                    sb.AppendLine(" Semicode = " & PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                    sb.AppendLine(" length = " & PrepareStr(ExcelRow(0)("Length")) & ", ")
                                    sb.AppendLine(" number = " & PrepareStr(ExcelRow(0)("Num")) & ", ")
                                    sb.AppendLine(" QTU = " & PrepareStr((ExcelRow(0)("QTU") * ExcelRow(0)("Num"))) & ", ")
                                    sb.AppendLine(" Dateup = " & PrepareStr(strDate))
                                    sb.AppendLine(" Where TireCode = '" & strGreentire & "' AND Rev = '" & strRevision & "' AND MaterialType = '14'")

                                    QBF = ExcelRow(0)("QTU") * ExcelRow(0)("Num")
                                    ''//---------------------------------------------------------------------------------------------------------------------------------

                                    sb.AppendLine(" ")

                                    ''// Cussion, BodyPly, Belt-1, Belt-2, Belt-3, Belt-4, Side, InnerLiner (Require, Need Num and Length) ==============================
                                    ''// Cussion
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'CUSSION'")
                                    sb.AppendLine(" Update TblGtDtl")
                                    sb.AppendLine(" Set ")
                                    sb.AppendLine(" Semicode = " & PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                    sb.AppendLine(" length = " & PrepareStr(ExcelRow(0)("Length")) & ", ")
                                    sb.AppendLine(" number = " & PrepareStr(ExcelRow(0)("Num")) & ", ")
                                    sb.AppendLine(" QTU = " & PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ")
                                    sb.AppendLine(" Dateup = " & PrepareStr(strDate))
                                    sb.AppendLine(" Where TireCode = '" & strGreentire & "' AND Rev = '" & strRevision & "' AND MaterialType = '03'")

                                    QCussion = (ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000

                                    sb.AppendLine(" ")

                                    ''// BodyPly
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'BODY PLY'")
                                    sb.AppendLine(" Update TblGtDtl")
                                    sb.AppendLine(" Set ")
                                    sb.AppendLine(" Semicode = " & PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                    sb.AppendLine(" length = " & PrepareStr(ExcelRow(0)("Length")) & ", ")
                                    sb.AppendLine(" number = " & PrepareStr(ExcelRow(0)("Num")) & ", ")
                                    sb.AppendLine(" QTU = " & PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ")
                                    sb.AppendLine(" Dateup = " & PrepareStr(strDate))
                                    sb.AppendLine(" Where TireCode = '" & strGreentire & "' AND Rev = '" & strRevision & "' AND MaterialType = '04'")

                                    QBodyPly = (ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000

                                    sb.AppendLine(" ")

                                    ''// Belt-1
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'BELT-1'")
                                    sb.AppendLine(" Update TblGtDtl")
                                    sb.AppendLine(" Set ")
                                    sb.AppendLine(" Semicode = " & PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                    sb.AppendLine(" length = " & PrepareStr(ExcelRow(0)("Length")) & ", ")
                                    sb.AppendLine(" number = " & PrepareStr(ExcelRow(0)("Num")) & ", ")
                                    sb.AppendLine(" QTU = " & PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ")
                                    sb.AppendLine(" Dateup = " & PrepareStr(strDate))
                                    sb.AppendLine(" Where TireCode = '" & strGreentire & "' AND Rev = '" & strRevision & "' AND MaterialType = '05'")

                                    QBelt1 = (ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000

                                    sb.AppendLine(" ")

                                    ''// Belt-2
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'BELT-2'")
                                    sb.AppendLine(" Update TblGtDtl")
                                    sb.AppendLine(" Set ")
                                    sb.AppendLine(" Semicode = " & PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                    sb.AppendLine(" length = " & PrepareStr(ExcelRow(0)("Length")) & ", ")
                                    sb.AppendLine(" number = " & PrepareStr(ExcelRow(0)("Num")) & ", ")
                                    sb.AppendLine(" QTU = " & PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ")
                                    sb.AppendLine(" Dateup = " & PrepareStr(strDate))
                                    sb.AppendLine(" Where TireCode = '" & strGreentire & "' AND Rev = '" & strRevision & "' AND MaterialType = '06'")

                                    QBelt2 = (ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000

                                    sb.AppendLine(" ")

                                    ''// Belt-3
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'BELT-3'")
                                    sb.AppendLine(" Update TblGtDtl")
                                    sb.AppendLine(" Set ")
                                    sb.AppendLine(" Semicode = " & PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                    sb.AppendLine(" length = " & PrepareStr(ExcelRow(0)("Length")) & ", ")
                                    sb.AppendLine(" number = " & PrepareStr(ExcelRow(0)("Num")) & ", ")
                                    sb.AppendLine(" QTU = " & PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ")
                                    sb.AppendLine(" Dateup = " & PrepareStr(strDate))
                                    sb.AppendLine(" Where TireCode = '" & strGreentire & "' AND Rev = '" & strRevision & "' AND MaterialType = '07'")

                                    QBelt3 = (ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000

                                    sb.AppendLine(" ")

                                    ''// Belt-4
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'BELT-4'")
                                    sb.AppendLine(" Update TblGtDtl")
                                    sb.AppendLine(" Set ")
                                    sb.AppendLine(" Semicode = " & PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                    sb.AppendLine(" length = " & PrepareStr(ExcelRow(0)("Length")) & ", ")
                                    sb.AppendLine(" number = " & PrepareStr(ExcelRow(0)("Num")) & ", ")
                                    sb.AppendLine(" QTU = " & PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ")
                                    sb.AppendLine(" Dateup = " & PrepareStr(strDate))
                                    sb.AppendLine(" Where TireCode = '" & strGreentire & "' AND Rev = '" & strRevision & "' AND MaterialType = '08'")

                                    QBelt4 = (ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000

                                    sb.AppendLine(" ")

                                    ''// Side
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'SIDE'")
                                    sb.AppendLine(" Update TblGtDtl")
                                    sb.AppendLine(" Set ")
                                    sb.AppendLine(" Semicode = " & PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                    sb.AppendLine(" length = " & PrepareStr(ExcelRow(0)("Length")) & ", ")
                                    sb.AppendLine(" number = " & PrepareStr(ExcelRow(0)("Num")) & ", ")
                                    sb.AppendLine(" QTU = " & PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ")
                                    sb.AppendLine(" Dateup = " & PrepareStr(strDate))
                                    sb.AppendLine(" Where TireCode = '" & strGreentire & "' AND Rev = '" & strRevision & "' AND MaterialType = '11'")

                                    QSide = (ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000

                                    sb.AppendLine(" ")

                                    ''// InnerLiner
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'INNERLINER'")
                                    sb.AppendLine(" Update TblGtDtl")
                                    sb.AppendLine(" Set ")
                                    sb.AppendLine(" Semicode = " & PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                    sb.AppendLine(" length = " & PrepareStr(ExcelRow(0)("Length")) & ", ")
                                    sb.AppendLine(" number = " & PrepareStr(ExcelRow(0)("Num")) & ", ")
                                    sb.AppendLine(" QTU = " & PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ")
                                    sb.AppendLine(" Dateup = " & PrepareStr(strDate))
                                    sb.AppendLine(" Where TireCode = '" & strGreentire & "' AND Rev = '" & strRevision & "' AND MaterialType = '12'")

                                    QInnerLiner = (ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000
                                    ''//==================================================================================================================================

                                    ''// WireChafer, NylonChafer, Flipper (No Require, Need Num and Length) **************************************************************
                                    sb.AppendLine(" ")

                                    ''// WireChafer
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'WIRE CHAFER'")
                                    If ExcelRow.Count > 0 Then
                                        sb.AppendLine(" Update TblGtDtl")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" Semicode = " & PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                        sb.AppendLine(" length = " & PrepareStr(ExcelRow(0)("Length")) & ", ")
                                        sb.AppendLine(" number = " & PrepareStr(ExcelRow(0)("Num")) & ", ")
                                        sb.AppendLine(" QTU = " & PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ")
                                        sb.AppendLine(" Dateup = " & PrepareStr(strDate))
                                        sb.AppendLine(" Where TireCode = '" & strGreentire & "' AND Rev = '" & strRevision & "' AND MaterialType = '09'")

                                        QWireChafer = (ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000
                                    Else
                                        sb.AppendLine(" Update TblGtDtl")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" Semicode = " & PrepareStr("") & ", ")
                                        sb.AppendLine(" length = " & PrepareStr("") & ", ")
                                        sb.AppendLine(" number = " & PrepareStr("") & ", ")
                                        sb.AppendLine(" QTU = " & PrepareStr("") & ", ")
                                        sb.AppendLine(" Dateup = " & PrepareStr(strDate))
                                        sb.AppendLine(" Where TireCode = '" & strGreentire & "' AND Rev = '" & strRevision & "' AND MaterialType = '09'")

                                        QWireChafer = 0
                                    End If

                                    sb.AppendLine(" ")

                                    ''// NylonChafer
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'Nylon CHAFER'")
                                    If ExcelRow.Count > 0 Then
                                        sb.AppendLine(" Update TblGtDtl")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" Semicode = " & PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                        sb.AppendLine(" length = " & PrepareStr(ExcelRow(0)("Length")) & ", ")
                                        sb.AppendLine(" number = " & PrepareStr(ExcelRow(0)("Num")) & ", ")
                                        sb.AppendLine(" QTU = " & PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ")
                                        sb.AppendLine(" Dateup = " & PrepareStr(strDate))
                                        sb.AppendLine(" Where TireCode = '" & strGreentire & "' AND Rev = '" & strRevision & "' AND MaterialType = '10'")

                                        QNylonChafer = (ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000
                                    Else
                                        sb.AppendLine(" Update TblGtDtl")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" Semicode = " & PrepareStr("") & ", ")
                                        sb.AppendLine(" length = " & PrepareStr("") & ", ")
                                        sb.AppendLine(" number = " & PrepareStr("") & ", ")
                                        sb.AppendLine(" QTU = " & PrepareStr("") & ", ")
                                        sb.AppendLine(" Dateup = " & PrepareStr(strDate))
                                        sb.AppendLine(" Where TireCode = '" & strGreentire & "' AND Rev = '" & strRevision & "' AND MaterialType = '10'")

                                        QNylonChafer = 0
                                    End If

                                    sb.AppendLine(" ")

                                    ''// Flipper
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'FLIPPER'")
                                    GridRow = DT.Select("EachGreenTire = '" & strGreentire & "' AND EachRevision = '" & strRevision & "' AND MaterialCode = '22'")
                                    If ExcelRow.Count > 0 Then
                                        If GridRow.Count > 0 Then
                                            sb.AppendLine(" Update TblGtDtl")
                                            sb.AppendLine(" Set ")
                                            sb.AppendLine(" Semicode = " & PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                            sb.AppendLine(" length = " & PrepareStr(ExcelRow(0)("Length")) & ", ")
                                            sb.AppendLine(" number = " & PrepareStr(ExcelRow(0)("Num")) & ", ")
                                            sb.AppendLine(" QTU = " & PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ")
                                            sb.AppendLine(" Dateup = " & PrepareStr(strDate))
                                            sb.AppendLine(" Where TireCode = '" & strGreentire & "' AND Rev = '" & strRevision & "' AND MaterialType = '22'")
                                        Else
                                            sb.AppendLine(" Insert TblGtDtl ")
                                            sb.AppendLine(" Values ")
                                            sb.AppendLine(" (")
                                            sb.AppendLine(PrepareStr(strGreentire) & ", ")
                                            sb.AppendLine(PrepareStr(strRevision) & ", ")
                                            sb.AppendLine(PrepareStr("22") & ", ")
                                            sb.AppendLine(PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                            sb.AppendLine(PrepareStr(ExcelRow(0)("Length")) & ", ")
                                            sb.AppendLine(PrepareStr(ExcelRow(0)("Num")) & ", ")
                                            sb.AppendLine(PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ")
                                            sb.AppendLine(PrepareStr("g") & ", ")
                                            sb.AppendLine(PrepareStr(strDate))
                                            sb.AppendLine(") ")
                                        End If

                                        QFlipper = (ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000
                                    Else
                                        If GridRow.Count > 0 Then
                                            sb.AppendLine(" Update TblGtDtl")
                                            sb.AppendLine(" Set ")
                                            sb.AppendLine(" Semicode = " & PrepareStr("") & ", ")
                                            sb.AppendLine(" length = " & PrepareStr("") & ", ")
                                            sb.AppendLine(" number = " & PrepareStr("") & ", ")
                                            sb.AppendLine(" QTU = " & PrepareStr("") & ", ")
                                            sb.AppendLine(" Dateup = " & PrepareStr(strDate))
                                            sb.AppendLine(" Where TireCode = '" & strGreentire & "' AND Rev = '" & strRevision & "' AND MaterialType = '22'")
                                        Else
                                            sb.AppendLine(" Insert TblGtDtl ")
                                            sb.AppendLine(" Values ")
                                            sb.AppendLine(" (")
                                            sb.AppendLine(PrepareStr(strGreentire) & ", ")
                                            sb.AppendLine(PrepareStr(strRevision) & ", ")
                                            sb.AppendLine(PrepareStr("22") & ", ")
                                            sb.AppendLine(PrepareStr("") & ", ")
                                            sb.AppendLine(PrepareStr("") & ", ")
                                            sb.AppendLine(PrepareStr("") & ", ")
                                            sb.AppendLine(PrepareStr("") & ", ")
                                            sb.AppendLine(PrepareStr("g") & ", ")
                                            sb.AppendLine(PrepareStr(strDate))
                                            sb.AppendLine(") ")
                                        End If

                                        QFlipper = 0
                                    End If
                                    ''//**********************************************************************************************************************************

                                    '//Summarize QTU
                                    totalQty = QBF + QSide + QInnerLiner + QCussion + QTread + QBodyPly + QBelt1 + QBelt2 + QBelt3 + QBelt4 + QWireChafer + QNylonChafer + QFlipper

                                    sb.AppendLine(" ")

                                    '//Update TBLGTHdr
                                    sb.AppendLine(" Update TBLGTHdr")
                                    sb.AppendLine(" Set ")
                                    sb.AppendLine(" Qty = " & PrepareStr(totalQty) & ", ")
                                    sb.AppendLine(" remark = " & PrepareStr(strRevisionBoss1st + "," + strRevisionBoss2nd) & ", ")
                                    sb.AppendLine(" Dateup = " & PrepareStr(strDate))
                                    sb.AppendLine(" Where TireCode = '" & strGreentire & "' AND Rev = '" & strRevision & "'")

                                    sb.AppendLine(" ")

                                    '//Update TblConvert
                                    sb.AppendLine(" Update TblConvert")
                                    sb.AppendLine(" Set ")
                                    sb.AppendLine(" SQty = " & PrepareStr((totalQty / 1000)))
                                    sb.AppendLine(" Where Code = '" & strGreentire & "' AND Rev = '" & strRevision & "' AND UnitBig = 'UT'")

                                    StrSQL = sb.ToString()
                                    cmSQL.CommandText = StrSQL
                                    cmSQL.ExecuteNonQuery()

                                End If

                            Else '//Case Insert

                                If strGreentire <> chkSameGreenTireBefore Or strRevision <> chkSameRevisionBefore Then

                                    '//Insert TblGtDtl
                                    '// Tread and BF (Require, Need only Num) ------------------------------------------------------------------------------------------
                                    '// Tread
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'TREAD'")
                                    sb.Clear()
                                    sb.AppendLine(" Insert INTO TblGtDtl(TireCode,Rev,MaterialType,Semicode,length,number,QTU,Unit,Dateup) ")
                                    sb.AppendLine(" Values ")
                                    sb.AppendLine(" (")
                                    sb.AppendLine(PrepareStr(strGreentire) & ", ") 'TireCode
                                    sb.AppendLine(PrepareStr(strRevision) & ", ") 'Rev
                                    sb.AppendLine(PrepareStr("13") & ", ") 'MaterialType
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("SemiCode")) & ", ") 'Semicode
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("Length")) & ", ") 'length
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("Num")) & ", ") 'number
                                    sb.AppendLine(PrepareStr((ExcelRow(0)("QTU") * ExcelRow(0)("Num"))) & ", ") 'QTU
                                    sb.AppendLine(PrepareStr("g") & ", ") 'Unit
                                    sb.AppendLine(PrepareStr(strDate)) 'Dateup
                                    sb.AppendLine(") ")

                                    QTread = ExcelRow(0)("QTU")

                                    sb.AppendLine(", ")

                                    ''// BF
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'BF (Upper,Lower,Center)'")
                                    sb.AppendLine(" (")
                                    sb.AppendLine(PrepareStr(strGreentire) & ", ")
                                    sb.AppendLine(PrepareStr(strRevision) & ", ")
                                    sb.AppendLine(PrepareStr("14") & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("Length")) & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("Num")) & ", ")
                                    sb.AppendLine(PrepareStr((ExcelRow(0)("QTU") * ExcelRow(0)("Num"))) & ", ")
                                    sb.AppendLine(PrepareStr("g") & ", ")
                                    sb.AppendLine(PrepareStr(strDate))
                                    sb.AppendLine(") ")

                                    QBF = ExcelRow(0)("QTU") * ExcelRow(0)("Num")
                                    ''//---------------------------------------------------------------------------------------------------------------------------------

                                    ''// Cussion, BodyPly, Belt-1, Belt-2, Belt-3, Belt-4, Side, InnerLiner (Require, Need Num and Length) ==============================
                                    sb.AppendLine(", ")

                                    ''// Cussion
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'CUSSION'")
                                    sb.AppendLine(" (")
                                    sb.AppendLine(PrepareStr(strGreentire) & ", ")
                                    sb.AppendLine(PrepareStr(strRevision) & ", ")
                                    sb.AppendLine(PrepareStr("03") & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("Length")) & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("Num")) & ", ")
                                    sb.AppendLine(PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ")
                                    sb.AppendLine(PrepareStr("g") & ", ")
                                    sb.AppendLine(PrepareStr(strDate))
                                    sb.AppendLine(") ")

                                    QCussion = (ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000

                                    sb.AppendLine(", ")

                                    ''// BodyPly
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'BODY PLY'")
                                    sb.AppendLine(" (")
                                    sb.AppendLine(PrepareStr(strGreentire) & ", ")
                                    sb.AppendLine(PrepareStr(strRevision) & ", ")
                                    sb.AppendLine(PrepareStr("04") & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("Length")) & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("Num")) & ", ")
                                    sb.AppendLine(PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ")
                                    sb.AppendLine(PrepareStr("g") & ", ")
                                    sb.AppendLine(PrepareStr(strDate))
                                    sb.AppendLine(") ")

                                    QBodyPly = (ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000

                                    sb.AppendLine(", ")

                                    ''// Belt-1
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'BELT-1'")
                                    sb.AppendLine(" (")
                                    sb.AppendLine(PrepareStr(strGreentire) & ", ")
                                    sb.AppendLine(PrepareStr(strRevision) & ", ")
                                    sb.AppendLine(PrepareStr("05") & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("Length")) & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("Num")) & ", ")
                                    sb.AppendLine(PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ")
                                    sb.AppendLine(PrepareStr("g") & ", ")
                                    sb.AppendLine(PrepareStr(strDate))
                                    sb.AppendLine(") ")

                                    QBelt1 = (ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000

                                    sb.AppendLine(", ")

                                    ''// Belt-2
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'BELT-2'")
                                    sb.AppendLine(" (")
                                    sb.AppendLine(PrepareStr(strGreentire) & ", ")
                                    sb.AppendLine(PrepareStr(strRevision) & ", ")
                                    sb.AppendLine(PrepareStr("06") & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("Length")) & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("Num")) & ", ")
                                    sb.AppendLine(PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ")
                                    sb.AppendLine(PrepareStr("g") & ", ")
                                    sb.AppendLine(PrepareStr(strDate))
                                    sb.AppendLine(") ")

                                    QBelt2 = (ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000

                                    sb.AppendLine(", ")

                                    ''// Belt-3
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'BELT-3'")
                                    sb.AppendLine(" (")
                                    sb.AppendLine(PrepareStr(strGreentire) & ", ")
                                    sb.AppendLine(PrepareStr(strRevision) & ", ")
                                    sb.AppendLine(PrepareStr("07") & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("Length")) & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("Num")) & ", ")
                                    sb.AppendLine(PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ")
                                    sb.AppendLine(PrepareStr("g") & ", ")
                                    sb.AppendLine(PrepareStr(strDate))
                                    sb.AppendLine(") ")

                                    QBelt3 = (ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000

                                    sb.AppendLine(", ")

                                    ''// Belt-4
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'BELT-4'")
                                    sb.AppendLine(" (")
                                    sb.AppendLine(PrepareStr(strGreentire) & ", ")
                                    sb.AppendLine(PrepareStr(strRevision) & ", ")
                                    sb.AppendLine(PrepareStr("08") & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("Length")) & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("Num")) & ", ")
                                    sb.AppendLine(PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ")
                                    sb.AppendLine(PrepareStr("g") & ", ")
                                    sb.AppendLine(PrepareStr(strDate))
                                    sb.AppendLine(") ")

                                    QBelt4 = (ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000

                                    sb.AppendLine(", ")

                                    ''// Side
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'SIDE'")
                                    sb.AppendLine(" (")
                                    sb.AppendLine(PrepareStr(strGreentire) & ", ")
                                    sb.AppendLine(PrepareStr(strRevision) & ", ")
                                    sb.AppendLine(PrepareStr("11") & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("Length")) & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("Num")) & ", ")
                                    sb.AppendLine(PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ")
                                    sb.AppendLine(PrepareStr("g") & ", ")
                                    sb.AppendLine(PrepareStr(strDate))
                                    sb.AppendLine(") ")

                                    QSide = (ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000

                                    sb.AppendLine(", ")

                                    ''// InnerLiner
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'INNERLINER'")
                                    sb.AppendLine(" (")
                                    sb.AppendLine(PrepareStr(strGreentire) & ", ")
                                    sb.AppendLine(PrepareStr(strRevision) & ", ")
                                    sb.AppendLine(PrepareStr("12") & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("Length")) & ", ")
                                    sb.AppendLine(PrepareStr(ExcelRow(0)("Num")) & ", ")
                                    sb.AppendLine(PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ")
                                    sb.AppendLine(PrepareStr("g") & ", ")
                                    sb.AppendLine(PrepareStr(strDate))
                                    sb.AppendLine(") ")

                                    QInnerLiner = (ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000
                                    ''//==================================================================================================================================

                                    ''// WireChafer, NylonChafer, Flipper (No Require, Need Num and Length) **************************************************************
                                    sb.AppendLine(", ")

                                    ''// WireChafer
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'WIRE CHAFER'")
                                    If ExcelRow.Count > 0 Then
                                        sb.AppendLine(" (")
                                        sb.AppendLine(PrepareStr(strGreentire) & ", ")
                                        sb.AppendLine(PrepareStr(strRevision) & ", ")
                                        sb.AppendLine(PrepareStr("09") & ", ")
                                        sb.AppendLine(PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                        sb.AppendLine(PrepareStr(ExcelRow(0)("Length")) & ", ")
                                        sb.AppendLine(PrepareStr(ExcelRow(0)("Num")) & ", ")

                                        sb.AppendLine(PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ")
                                        sb.AppendLine(PrepareStr("g") & ", ")
                                        sb.AppendLine(PrepareStr(strDate))
                                        sb.AppendLine(") ")

                                        QWireChafer = (ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000
                                    Else
                                        sb.AppendLine(" (")
                                        sb.AppendLine(PrepareStr(strGreentire) & ", ")
                                        sb.AppendLine(PrepareStr(strRevision) & ", ")
                                        sb.AppendLine(PrepareStr("09") & ", ")
                                        sb.AppendLine(PrepareStr("") & ", ")
                                        sb.AppendLine(PrepareStr("") & ", ")
                                        sb.AppendLine(PrepareStr("") & ", ")

                                        sb.AppendLine(PrepareStr("") & ", ")
                                        sb.AppendLine(PrepareStr("g") & ", ")
                                        sb.AppendLine(PrepareStr(strDate))
                                        sb.AppendLine(") ")

                                        QWireChafer = 0
                                    End If


                                    sb.AppendLine(", ")

                                    ''// NylonChafer
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'Nylon CHAFER'")
                                    If ExcelRow.Count > 0 Then
                                        sb.AppendLine(" (")
                                        sb.AppendLine(PrepareStr(strGreentire) & ", ")
                                        sb.AppendLine(PrepareStr(strRevision) & ", ")
                                        sb.AppendLine(PrepareStr("10") & ", ")
                                        sb.AppendLine(PrepareStr(ExcelRow(0)("SemiCode")) & ", ")
                                        sb.AppendLine(PrepareStr(ExcelRow(0)("Length")) & ", ")
                                        sb.AppendLine(PrepareStr(ExcelRow(0)("Num")) & ", ")

                                        sb.AppendLine(PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ")
                                        sb.AppendLine(PrepareStr("g") & ", ")
                                        sb.AppendLine(PrepareStr(strDate))
                                        sb.AppendLine(") ")

                                        QNylonChafer = (ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000
                                    Else
                                        sb.AppendLine(" (")
                                        sb.AppendLine(PrepareStr(strGreentire) & ", ")
                                        sb.AppendLine(PrepareStr(strRevision) & ", ")
                                        sb.AppendLine(PrepareStr("10") & ", ")
                                        sb.AppendLine(PrepareStr("") & ", ")
                                        sb.AppendLine(PrepareStr("") & ", ")
                                        sb.AppendLine(PrepareStr("") & ", ")

                                        sb.AppendLine(PrepareStr("") & ", ")
                                        sb.AppendLine(PrepareStr("g") & ", ")
                                        sb.AppendLine(PrepareStr(strDate))
                                        sb.AppendLine(") ")

                                        QNylonChafer = 0
                                    End If

                                    sb.AppendLine(", ")

                                    ''// Flipper
                                    ExcelRow = dtRec.Select("GreenTire = '" & strGreentire & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'FLIPPER'")
                                    If ExcelRow.Count > 0 Then
                                        sb.AppendLine(" (")
                                        sb.AppendLine(PrepareStr(strGreentire) & ", ") 'TireCode
                                        sb.AppendLine(PrepareStr(strRevision) & ", ") 'Rev
                                        sb.AppendLine(PrepareStr("22") & ", ") 'MaterialType
                                        sb.AppendLine(PrepareStr(ExcelRow(0)("SemiCode")) & ", ") 'Semicode
                                        sb.AppendLine(PrepareStr(ExcelRow(0)("Length")) & ", ") 'length
                                        sb.AppendLine(PrepareStr(ExcelRow(0)("Num")) & ", ") 'number

                                        sb.AppendLine(PrepareStr(((ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000)) & ", ") 'QTU
                                        sb.AppendLine(PrepareStr("g") & ", ") 'Unit
                                        sb.AppendLine(PrepareStr(strDate)) 'Dateup
                                        sb.AppendLine(") ")

                                        QFlipper = (ExcelRow(0)("QTU") * ExcelRow(0)("Length")) / 1000
                                    Else
                                        sb.AppendLine(" (")
                                        sb.AppendLine(PrepareStr(strGreentire) & ", ") 'TireCode
                                        sb.AppendLine(PrepareStr(strRevision) & ", ") 'Rev
                                        sb.AppendLine(PrepareStr("22") & ", ") 'MaterialType
                                        sb.AppendLine(PrepareStr("") & ", ") 'Semicode
                                        sb.AppendLine(PrepareStr("") & ", ") 'length
                                        sb.AppendLine(PrepareStr("") & ", ") 'number

                                        sb.AppendLine(PrepareStr("") & ", ") 'QTU
                                        sb.AppendLine(PrepareStr("g") & ", ") 'Unit
                                        sb.AppendLine(PrepareStr(strDate)) 'Dateup
                                        sb.AppendLine(") ")

                                        QFlipper = 0
                                    End If

                                    StrSQL = sb.ToString()
                                    cmSQL.CommandText = StrSQL
                                    cmSQL.ExecuteNonQuery()
                                    ''//**********************************************************************************************************************************

                                    '//Summarize QTU
                                    totalQty = QBF + QSide + QInnerLiner + QCussion + QTread + QBodyPly + QBelt1 + QBelt2 + QBelt3 + QBelt4 + QWireChafer + QNylonChafer + QFlipper

                                    '//Insert TBLGroup
                                    sb.Clear()
                                    sb.AppendLine(" Insert  TBLGroup ")
                                    sb.AppendLine(" Values (")
                                    sb.AppendLine(PrepareStr("06") & ", ")      'Column TypeCode
                                    sb.AppendLine(PrepareStr(strGreentire))     'Column Code
                                    sb.AppendLine(" )")
                                    StrSQL = sb.ToString()
                                    cmSQL.CommandText = StrSQL
                                    cmSQL.ExecuteNonQuery()

                                    '//Insert TBLGTHdr
                                    sb.Clear()
                                    sb.AppendLine(" Insert  TBLGTHdr ")
                                    sb.AppendLine(" Values (")
                                    sb.AppendLine(PrepareStr(strGreentire) & ", ")                              'Column Final
                                    sb.AppendLine(PrepareStr(strGreentire) & ", ")                              'Column TireCide
                                    sb.AppendLine(PrepareStr(strRevision) & ", ")                               'Column Rev
                                    sb.AppendLine(PrepareStr(strBSJ) & ", ")                                    'Column TireSize
                                    sb.AppendLine(PrepareStr(totalQty) & ", ")                                  'Column Qty
                                    sb.AppendLine(PrepareStr(0) & ", ")                                         'Column Active
                                    sb.AppendLine(PrepareStr(strDate) & ", ")                                   'Column Dateup
                                    sb.AppendLine(PrepareStr(strRevisionBoss1st + "," + strRevisionBoss2nd))    'Column remark
                                    sb.AppendLine(" )")
                                    StrSQL = sb.ToString()
                                    cmSQL.CommandText = StrSQL
                                    cmSQL.ExecuteNonQuery()

                                    '//Insert TblConvert #1
                                    sb.Clear()
                                    sb.AppendLine(" Insert  TblConvert ")
                                    sb.AppendLine(" Values (")
                                    sb.AppendLine(PrepareStr("06") & ", ")          'Column Type
                                    sb.AppendLine(PrepareStr(strGreentire) & ", ")  'Column Final
                                    sb.AppendLine(PrepareStr(strGreentire) & ", ")  'Column Code
                                    sb.AppendLine(PrepareStr(strRevision) & ", ")   'Column Rev
                                    sb.AppendLine(PrepareStr("KG") & ", ")          'Column UnitBig
                                    sb.AppendLine(PrepareStr("KG") & ", ")          'Column UnitSmall
                                    sb.AppendLine(PrepareStr("1") & ", ")           'Column BQty
                                    sb.AppendLine(PrepareStr("1"))                  'Column SQty
                                    sb.AppendLine(" )")

                                    '//Insert TblConvert #2
                                    sb.AppendLine(" ")

                                    sb.AppendLine(" Insert  TblConvert ")
                                    sb.AppendLine(" Values (")
                                    sb.AppendLine(PrepareStr("06") & ", ")          'Column Type
                                    sb.AppendLine(PrepareStr(strGreentire) & ", ")  'Column Final
                                    sb.AppendLine(PrepareStr(strGreentire) & ", ")  'Column Code
                                    sb.AppendLine(PrepareStr(strRevision) & ", ")   'Column Rev
                                    sb.AppendLine(PrepareStr("UT") & ", ")          'Column UnitBig
                                    sb.AppendLine(PrepareStr("KG") & ", ")          'Column UnitSmall
                                    sb.AppendLine(PrepareStr("1") & ", ")           'Column BQty
                                    sb.AppendLine(PrepareStr((totalQty / 1000)))    'Column SQty
                                    sb.AppendLine(" )")

                                    StrSQL = sb.ToString()
                                    cmSQL.CommandText = StrSQL
                                    cmSQL.ExecuteNonQuery()
                                End If
                            End If 'If GridRow.Count > 0
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

            LoadTire() 'ReQuery and set datagrid
            frmOverlay.Dispose()
        End If 'If importDialog.ShowDialog() = Windows.Forms.DialogResult.OK
    End Sub
#End Region

#Region "Import"
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

    Private Function ChkImportData_Correctly(ByRef ImportTable As DataTable, dtTypeMaterial As DataTable, dtTireSize As DataTable) As Boolean
        Dim cnSQLRM As SqlConnection
        Dim cmSQLRM As SqlCommand
        Dim strSQL As String = String.Empty
        Dim ret As Boolean = False
        Dim strRmcodeBefore As String = String.Empty
        Dim QBF, QSide, QInnerLiner, QCussion, QTread, QBodyPly, QBelt1, QBelt2, QBelt3, QBelt4, QWireChafer, QNylonChafer, QFlipper As Double
        Dim distinctImportTabale As New DataTable
        Dim importRow As DataRow()

        Try
            'Check empty value
            For x As Integer = 0 To ImportTable.Rows.Count - 1
                Dim strGreenTireCode As String = ImportTable.Rows(x)("GreenTire").ToString().Trim()
                Dim strRevision As String = ImportTable.Rows(x)("Revision").ToString().Trim()
                Dim strBSJ As String = ImportTable.Rows(x)("BSJ").ToString().Trim() 'Tire Size
                Dim strRevisionBoss1st As String = ImportTable.Rows(x)("RevisionBoss_1st").ToString().Trim() 'Revision of Boss 1st
                Dim strRevisionBoss2nd As String = ImportTable.Rows(x)("RevisionBoss_2nd").ToString().Trim() 'Revision of Boss 2nd
                Dim strTypeMaterial As String = ImportTable.Rows(x)("TypeMaterial").ToString().Trim()
                Dim strSemiCode As String = ImportTable.Rows(x)("SemiCode").ToString().Trim()

                strSQL = String.Empty
                cnSQLRM = New SqlConnection(C1.Strcon)
                cnSQLRM.Open()

                '// 1.) Check GreeTireCode
                If strGreenTireCode.Length <= 0 Then
                    Throw New ApplicationException("Green Tire Code is not empty.")
                End If

                '// 2.) Check Revision
                If strRevision.Length <= 0 Then
                    Throw New System.Exception("Revision is not empty.")
                ElseIf strRevision.Length > 3 Then
                    Throw New System.Exception("Revision must less than 4 digits.")
                End If

                '// 3.) Check BSJ (TireSize)
                If strBSJ.Length <= 0 Then
                    Throw New System.Exception("BSJ (Tire Size) is not empty.")
                End If

                '// 4.) Check RevisionBoss 1st
                If strRevisionBoss1st.Length <= 0 Then
                    Throw New System.Exception("Revision of Boss 1st is not empty.")
                End If

                '// 5.) Check RevisionBoss 2nd
                If strRevisionBoss2nd.Length <= 0 Then
                    Throw New System.Exception("Revision of Boss 2nd is not empty.")
                End If

                '//6.) Check Type Material
                If strTypeMaterial.Equals(String.Empty) Then
                    Throw New ApplicationException("Type Material is not empty.")
                End If

                '7.) Check Semi Code
                If strSemiCode.Length <= 0 Then
                    Throw New ApplicationException("Semi Code is not empty.")
                End If
            Next x

            For x As Integer = 0 To ImportTable.Rows.Count - 1
                Dim strGreenTireCode As String = ImportTable.Rows(x)("GreenTire").ToString().Trim()
                Dim strRevision As String = ImportTable.Rows(x)("Revision").ToString().Trim()
                Dim strBSJ As String = ImportTable.Rows(x)("BSJ").ToString().Trim() 'Tire Size
                Dim strRevisionBoss1st As String = ImportTable.Rows(x)("RevisionBoss_1st").ToString().Trim() 'Revision of Boss 1st
                Dim strRevisionBoss2nd As String = ImportTable.Rows(x)("RevisionBoss_2nd").ToString().Trim() 'Revision of Boss 2nd
                Dim strTypeMaterial As String = ImportTable.Rows(x)("TypeMaterial").ToString().Trim()
                Dim dblNum As Double = 0.0

                strSQL = String.Empty
                cnSQLRM = New SqlConnection(C1.Strcon)
                cnSQLRM.Open()

                If strGreenTireCode.Length > 0 Then
                    '//For Check Data from above row on import file.
                    Dim chkSameGreenTireBefore As String = String.Empty
                    Dim chkSameRevisionBefore As String = String.Empty

                    If x > 0 Then
                        chkSameGreenTireBefore = ImportTable.Rows(x - 1)("GreenTire").ToString
                        chkSameRevisionBefore = ImportTable.Rows(x - 1)("Revision").ToString
                    Else
                        chkSameGreenTireBefore = String.Empty
                        chkSameRevisionBefore = String.Empty
                    End If

                    'Check Tire Size in Master
                    Dim arrTireSize As DataRow() = dtTireSize.Select("BSJCode = '" & strBSJ & "'")
                    If arrTireSize.Length = 0 Then
                        Throw New ApplicationException("BSJ Code: " & strBSJ & " is not found in master.")
                    End If

                    'Check Type Material in Master
                    Dim arrTypeMatCode As DataRow() = dtTypeMaterial.Select("MaterialName = '" & strTypeMaterial & "'")
                    If arrTypeMatCode.Length = 0 Then
                        Throw New ApplicationException("Material Code: " & strTypeMaterial & " is not found in master.")
                    End If

                    '// Check Each SemiCode in same group of GreenTire and Revision correctly
                    'For first GreenTire and Semi in each group
                    If strGreenTireCode <> chkSameGreenTireBefore Or strRevision <> chkSameRevisionBefore Then
                        '// Tread and BF (Require, Need only Num) ------------------------------------------------------------------------------------------
#Region "Material Type TREAD"
                        importRow = ImportTable.Select("GreenTire = '" & strGreenTireCode & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'TREAD'")
                        If importRow.Count > 0 Then
                            If importRow(0)("Num").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Num value of type 'Tread'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Num"), dblNum) Then
                                    Throw New System.Exception("Please input Num data as Number of type 'Tread'.")
                                End If
                            End If

                            'Check Semi in master
                            strSQL = " SELECT COUNT(*)  FROM  TblSemi  "
                            strSQL += " WHERE MaterialType = '13' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                            cnSQLRM = New SqlConnection(C1.Strcon)
                            cnSQLRM.Open()
                            cmSQLRM = New SqlCommand(strSQL, cnSQLRM)
                            Dim i As Long = cmSQLRM.ExecuteScalar()
                            If i = 0 Then
                                Throw New System.Exception("Please check correctly Tread '" & importRow(0)("SemiCode").ToString().Trim() & "'")
                            Else
                                'Get QPU
                                strSQL = " SELECT Round(QPU,4) QPU  FROM  TblSemi  "
                                strSQL += " WHERE MaterialType = '13' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                                cnSQLRM = New SqlConnection(C1.Strcon)
                                cnSQLRM.Open()
                                cmSQLRM = New SqlCommand(strSQL, cnSQLRM)

                                QTread = cmSQLRM.ExecuteScalar()
                                importRow(0)("QTU") = QTread

                                cmSQLRM.Dispose()
                                cnSQLRM.Dispose()
                            End If
                        Else
                            Throw New System.Exception("Green Tire: " & strGreenTireCode & " and Revision: " & strRevision & " does not found type 'Tread'")
                        End If
#End Region

#Region "Material Type BF"
                        importRow = ImportTable.Select("GreenTire = '" & strGreenTireCode & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'BF (Upper,Lower,Center)'")
                        If importRow.Count > 0 Then
                            If importRow(0)("Num").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Num value of type 'BF'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Num"), dblNum) Then
                                    Throw New System.Exception("Please input Num data as Number of type 'BF'.")
                                End If
                            End If

                            strSQL = " SELECT COUNT(*)  FROM  TblSemi  "
                            strSQL += " WHERE MaterialType = '14' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                            cnSQLRM = New SqlConnection(C1.Strcon)
                            cnSQLRM.Open()
                            cmSQLRM = New SqlCommand(strSQL, cnSQLRM)
                            Dim i As Long = cmSQLRM.ExecuteScalar()
                            If i = 0 Then
                                Throw New System.Exception("Please check correctly BF '" & importRow(0)("SemiCode").ToString().Trim() & "'")
                            Else
                                'Get QPU
                                strSQL = " SELECT Round(QPU,4) QPU  FROM  TblSemi  "
                                strSQL += " WHERE MaterialType = '14' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                                cnSQLRM = New SqlConnection(C1.Strcon)
                                cnSQLRM.Open()
                                cmSQLRM = New SqlCommand(strSQL, cnSQLRM)

                                QBF = cmSQLRM.ExecuteScalar()
                                importRow(0)("QTU") = QBF

                                cmSQLRM.Dispose()
                                cnSQLRM.Dispose()
                            End If
                        Else
                            Throw New System.Exception("Green Tire: " & strGreenTireCode & " and Revision: " & strRevision & " does not found type 'BF'")
                        End If
#End Region
                        '//---------------------------------------------------------------------------------------------------------------------------------

                        '// Cussion, BodyPly, Belt-1, Belt-2, Belt-3, Belt-4, Side, InnerLiner (Require, Need Num and Length) ==============================
#Region "Material Type Cussion"
                        importRow = ImportTable.Select("GreenTire = '" & strGreenTireCode & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'CUSSION'")
                        If importRow.Count > 0 Then
                            If importRow(0)("Num").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Num value of type 'Cussion'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Num"), dblNum) Then
                                    Throw New System.Exception("Please input Num data as Number of type 'Cussion'.")
                                End If
                            End If
                            If importRow(0)("Length").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Length value of type 'Cussion'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Length"), dblNum) Then
                                    Throw New System.Exception("Please input Length data as Number of type 'Cussion'.")
                                End If
                            End If

                            strSQL = " SELECT COUNT(*)  FROM  TblSemi  "
                            strSQL += " WHERE MaterialType = '03' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                            cnSQLRM = New SqlConnection(C1.Strcon)
                            cnSQLRM.Open()
                            cmSQLRM = New SqlCommand(strSQL, cnSQLRM)
                            Dim i As Long = cmSQLRM.ExecuteScalar()
                            If i = 0 Then
                                Throw New System.Exception("Please check correctly Cussion '" & importRow(0)("SemiCode").ToString().Trim() & "'")
                            Else
                                'Geet QPU
                                strSQL = " SELECT Round(QPU,4) QPU  FROM  TblSemi  "
                                strSQL += " WHERE MaterialType = '03' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                                cnSQLRM = New SqlConnection(C1.Strcon)
                                cnSQLRM.Open()
                                cmSQLRM = New SqlCommand(strSQL, cnSQLRM)

                                QCussion = cmSQLRM.ExecuteScalar()
                                importRow(0)("QTU") = QCussion

                                cmSQLRM.Dispose()
                                cnSQLRM.Dispose()
                            End If
                        Else
                            Throw New System.Exception("Green Tire: " & strGreenTireCode & " and Revision: " & strRevision & " does not found type 'Cussion'")
                        End If
#End Region

#Region "Material Type BodyPly"
                        importRow = ImportTable.Select("GreenTire = '" & strGreenTireCode & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'BODY PLY'")
                        If importRow.Count > 0 Then
                            If importRow(0)("Num").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Num value of type 'BodyPly'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Num"), dblNum) Then
                                    Throw New System.Exception("Please input Num data as Number of type 'BodyPly'.")
                                End If
                            End If
                            If importRow(0)("Length").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Length value of type 'BodyPly'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Length"), dblNum) Then
                                    Throw New System.Exception("Please input Length data as Number of type 'BodyPly'.")
                                End If
                            End If

                            strSQL = " SELECT COUNT(*)  FROM  TblSemi  "
                            strSQL += " WHERE MaterialType = '04' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                            cnSQLRM = New SqlConnection(C1.Strcon)
                            cnSQLRM.Open()
                            cmSQLRM = New SqlCommand(strSQL, cnSQLRM)
                            Dim i As Long = cmSQLRM.ExecuteScalar()
                            If i = 0 Then
                                Throw New System.Exception("Please check correctly BodyPly '" & importRow(0)("SemiCode").ToString().Trim() & "'")
                            Else
                                strSQL = " SELECT Round(QPU,4) QPU  FROM  TblSemi  "
                                strSQL += " WHERE MaterialType = '04' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                                cnSQLRM = New SqlConnection(C1.Strcon)
                                cnSQLRM.Open()
                                cmSQLRM = New SqlCommand(strSQL, cnSQLRM)

                                QBodyPly = cmSQLRM.ExecuteScalar()
                                importRow(0)("QTU") = QBodyPly

                                cmSQLRM.Dispose()
                                cnSQLRM.Dispose()
                            End If
                        Else
                            Throw New System.Exception("Green Tire: " & strGreenTireCode & " and Revision: " & strRevision & " does not found type 'BodyPly'")
                        End If
#End Region

#Region "Material Type Belt-1"
                        importRow = ImportTable.Select("GreenTire = '" & strGreenTireCode & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'BELT-1'")
                        If importRow.Count > 0 Then
                            If importRow(0)("Num").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Num value of type 'Belt-1'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Num"), dblNum) Then
                                    Throw New System.Exception("Please input Num data as Number of type 'Belt-1'.")
                                End If
                            End If
                            If importRow(0)("Length").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Length value of type 'Belt-1'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Length"), dblNum) Then
                                    Throw New System.Exception("Please input Length data as Number of 'Belt-2'.")
                                End If
                            End If

                            strSQL = " SELECT COUNT(*)  FROM  TblSemi  "
                            strSQL += " WHERE MaterialType = '05' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                            cnSQLRM = New SqlConnection(C1.Strcon)
                            cnSQLRM.Open()
                            cmSQLRM = New SqlCommand(strSQL, cnSQLRM)
                            Dim i As Long = cmSQLRM.ExecuteScalar()
                            If i = 0 Then
                                Throw New System.Exception("Please check correctly Belt-1 '" & importRow(0)("SemiCode").ToString().Trim() & "'")
                            Else
                                'Get QPU
                                strSQL = " SELECT Round(QPU,4) QPU  FROM  TblSemi  "
                                strSQL += " WHERE MaterialType = '05' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                                cnSQLRM = New SqlConnection(C1.Strcon)
                                cnSQLRM.Open()
                                cmSQLRM = New SqlCommand(strSQL, cnSQLRM)

                                QBelt1 = cmSQLRM.ExecuteScalar()
                                importRow(0)("QTU") = QBelt1

                                cmSQLRM.Dispose()
                                cnSQLRM.Dispose()
                            End If
                        Else
                            Throw New System.Exception("Green Tire: " & strGreenTireCode & " and Revision: " & strRevision & " does not found type 'Belt-1'")
                        End If
#End Region

#Region "Material Type Belt-2"
                        importRow = ImportTable.Select("GreenTire = '" & strGreenTireCode & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'BELT-2'")
                        If importRow.Count > 0 Then
                            If importRow(0)("Num").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Num value of type 'Belt-2'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Num"), dblNum) Then
                                    Throw New System.Exception("Please input Num data as Number of type 'Belt-2'.")
                                End If
                            End If
                            If importRow(0)("Length").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Length value of type 'Belt-2'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Length"), dblNum) Then
                                    Throw New System.Exception("Please input Length data as Number of type 'Belt-2'.")
                                End If
                            End If

                            strSQL = " SELECT COUNT(*)  FROM  TblSemi  "
                            strSQL += " WHERE MaterialType = '06' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                            cnSQLRM = New SqlConnection(C1.Strcon)
                            cnSQLRM.Open()
                            cmSQLRM = New SqlCommand(strSQL, cnSQLRM)
                            Dim i As Long = cmSQLRM.ExecuteScalar()
                            If i = 0 Then
                                Throw New System.Exception("Please check correctly Belt-2 '" & importRow(0)("SemiCode").ToString().Trim() & "'")
                            Else
                                'Get QPU
                                strSQL = " SELECT Round(QPU,4) QPU  FROM  TblSemi  "
                                strSQL += " WHERE MaterialType = '06' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                                cnSQLRM = New SqlConnection(C1.Strcon)
                                cnSQLRM.Open()
                                cmSQLRM = New SqlCommand(strSQL, cnSQLRM)

                                QBelt2 = cmSQLRM.ExecuteScalar()
                                importRow(0)("QTU") = QBelt2

                                cmSQLRM.Dispose()
                                cnSQLRM.Dispose()
                            End If
                        Else
                            Throw New System.Exception("Green Tire: " & strGreenTireCode & " and Revision: " & strRevision & " does not found type 'Belt-2'")
                        End If
#End Region

#Region "Material Type Belt-3"
                        importRow = ImportTable.Select("GreenTire = '" & strGreenTireCode & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'BELT-3'")
                        If importRow.Count > 0 Then
                            If importRow(0)("Num").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Num value of type 'Belt-3'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Num"), dblNum) Then
                                    Throw New System.Exception("Please input Num data as Number of type 'Belt-3'.")
                                End If
                            End If
                            If importRow(0)("Length").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Length value of type 'Belt-3'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Length"), dblNum) Then
                                    Throw New System.Exception("Please input Length data as Number of type 'Belt-3'.")
                                End If
                            End If

                            strSQL = " SELECT COUNT(*)  FROM  TblSemi  "
                            strSQL += " WHERE MaterialType = '07' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                            cnSQLRM = New SqlConnection(C1.Strcon)
                            cnSQLRM.Open()
                            cmSQLRM = New SqlCommand(strSQL, cnSQLRM)
                            Dim i As Long = cmSQLRM.ExecuteScalar()
                            If i = 0 Then
                                Throw New System.Exception("Please check correctly Belt-3 '" & importRow(0)("SemiCode").ToString().Trim() & "'")
                            Else
                                'Get QPU
                                strSQL = " SELECT Round(QPU,4) QPU  FROM  TblSemi  "
                                strSQL += " WHERE MaterialType = '07' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                                cnSQLRM = New SqlConnection(C1.Strcon)
                                cnSQLRM.Open()
                                cmSQLRM = New SqlCommand(strSQL, cnSQLRM)

                                QBelt3 = cmSQLRM.ExecuteScalar()
                                importRow(0)("QTU") = QBelt3

                                cmSQLRM.Dispose()
                                cnSQLRM.Dispose()
                            End If
                        Else
                            Throw New System.Exception("Green Tire: " & strGreenTireCode & " and Revision: " & strRevision & " does not found type 'Belt-3'")
                        End If
#End Region

#Region "Material Type Belt-4"
                        importRow = ImportTable.Select("GreenTire = '" & strGreenTireCode & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'BELT-4'")
                        If importRow.Count > 0 Then
                            If importRow(0)("Num").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Num value of type 'Belt-4'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Num"), dblNum) Then
                                    Throw New System.Exception("Please input Num data as Number of type 'Belt-4'.")
                                End If
                            End If
                            If importRow(0)("Length").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Length value of type 'Belt-4'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Length"), dblNum) Then
                                    Throw New System.Exception("Please input Length data as Number of type 'Belt-4'.")
                                End If
                            End If

                            strSQL = " SELECT COUNT(*)  FROM  TblSemi  "
                            strSQL += " WHERE MaterialType = '08' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                            cnSQLRM = New SqlConnection(C1.Strcon)
                            cnSQLRM.Open()
                            cmSQLRM = New SqlCommand(strSQL, cnSQLRM)
                            Dim i As Long = cmSQLRM.ExecuteScalar()
                            If i = 0 Then
                                Throw New System.Exception("Please check correctly Belt-4 '" & importRow(0)("SemiCode").ToString().Trim() & "'")
                            Else
                                'Get QPU
                                strSQL = " SELECT Round(QPU,4) QPU  FROM  TblSemi  "
                                strSQL += " WHERE MaterialType = '08' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                                cnSQLRM = New SqlConnection(C1.Strcon)
                                cnSQLRM.Open()
                                cmSQLRM = New SqlCommand(strSQL, cnSQLRM)

                                QBelt4 = cmSQLRM.ExecuteScalar()
                                importRow(0)("QTU") = QBelt4

                                cmSQLRM.Dispose()
                                cnSQLRM.Dispose()
                            End If
                        Else
                            Throw New System.Exception("Green Tire: " & strGreenTireCode & " and Revision: " & strRevision & " does not found type 'Belt-4'")
                        End If
#End Region

#Region "Material Type Side"
                        importRow = ImportTable.Select("GreenTire = '" & strGreenTireCode & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'SIDE'")
                        If importRow.Count > 0 Then
                            If importRow(0)("Num").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Num value of type 'Side'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Num"), dblNum) Then
                                    Throw New System.Exception("Please input Num data as Number of type 'Side'.")
                                End If
                            End If
                            If importRow(0)("Length").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Length value of type 'Side'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Length"), dblNum) Then
                                    Throw New System.Exception("Please input Length data as Number of type 'Side'.")
                                End If
                            End If

                            strSQL = " SELECT COUNT(*)  FROM  TblSemi  "
                            strSQL += " WHERE MaterialType = '11' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                            cnSQLRM = New SqlConnection(C1.Strcon)
                            cnSQLRM.Open()
                            cmSQLRM = New SqlCommand(strSQL, cnSQLRM)
                            Dim i As Long = cmSQLRM.ExecuteScalar()
                            If i = 0 Then
                                Throw New System.Exception("Please check correctly Side '" & importRow(0)("SemiCode").ToString().Trim() & "'")
                            Else
                                'Get QPU
                                strSQL = " SELECT Round(QPU,4) QPU  FROM  TblSemi  "
                                strSQL += " WHERE MaterialType = '11' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                                cnSQLRM = New SqlConnection(C1.Strcon)
                                cnSQLRM.Open()
                                cmSQLRM = New SqlCommand(strSQL, cnSQLRM)

                                QSide = cmSQLRM.ExecuteScalar()
                                importRow(0)("QTU") = QSide

                                cmSQLRM.Dispose()
                                cnSQLRM.Dispose()
                            End If
                        Else
                            Throw New System.Exception("Green Tire: " & strGreenTireCode & " and Revision: " & strRevision & " does not found type 'Side'")
                        End If
#End Region

#Region "Material Type InnerLiner"
                        importRow = ImportTable.Select("GreenTire = '" & strGreenTireCode & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'INNERLINER'")
                        If importRow.Count > 0 Then
                            If importRow(0)("Num").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Num value of type 'InnerLiner'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Num"), dblNum) Then
                                    Throw New System.Exception("Please input Num data as Number of type 'InnerLiner'.")
                                End If
                            End If
                            If importRow(0)("Length").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Length value of type 'InnerLiner'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Length"), dblNum) Then
                                    Throw New System.Exception("Please input Length data as Number of type 'InnerLiner'.")
                                End If
                            End If

                            strSQL = " SELECT COUNT(*)  FROM  TblSemi  "
                            strSQL += " WHERE MaterialType = '12' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                            cnSQLRM = New SqlConnection(C1.Strcon)
                            cnSQLRM.Open()
                            cmSQLRM = New SqlCommand(strSQL, cnSQLRM)
                            Dim i As Long = cmSQLRM.ExecuteScalar()
                            If i = 0 Then
                                Throw New System.Exception("Please check correctly InnerLiner '" & importRow(0)("SemiCode").ToString().Trim() & "'")
                            Else
                                'Get QPU
                                strSQL = " SELECT Round(QPU,4) QPU  FROM  TblSemi  "
                                strSQL += " WHERE MaterialType = '12' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                                cnSQLRM = New SqlConnection(C1.Strcon)
                                cnSQLRM.Open()
                                cmSQLRM = New SqlCommand(strSQL, cnSQLRM)

                                QInnerLiner = cmSQLRM.ExecuteScalar()
                                importRow(0)("QTU") = QInnerLiner

                                cmSQLRM.Dispose()
                                cnSQLRM.Dispose()
                            End If
                        Else
                            Throw New System.Exception("Green Tire: " & strGreenTireCode & " and Revision: " & strRevision & " does not found type 'InnerLiner'")
                        End If
#End Region
                        '//==================================================================================================================================

                        '// WireChafer, NylonChafer, Flipper (No Require, Need Num and Length) **************************************************************
#Region "Material Type WireChafer"
                        importRow = ImportTable.Select("GreenTire = '" & strGreenTireCode & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'WIRE CHAFER'")
                        If importRow.Count > 0 Then
                            If importRow(0)("Num").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Num value of type 'Wire Chafer'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Num"), dblNum) Then
                                    Throw New System.Exception("Please input Num data as Number of type 'Wire Chafer'.")
                                End If
                            End If
                            If importRow(0)("Length").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Length value of type 'Wire Chafer'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Length"), dblNum) Then
                                    Throw New System.Exception("Please input Length data as Number of type 'Wire Chafer'.")
                                End If
                            End If

                            strSQL = " SELECT COUNT(*)  FROM  TblSemi  "
                            strSQL += " WHERE MaterialType = '09' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                            cnSQLRM = New SqlConnection(C1.Strcon)
                            cnSQLRM.Open()
                            cmSQLRM = New SqlCommand(strSQL, cnSQLRM)
                            Dim i As Long = cmSQLRM.ExecuteScalar()
                            If i = 0 Then
                                Throw New System.Exception("Please check correctly Wire Chafer '" & importRow(0)("SemiCode").ToString().Trim() & "'")
                            Else
                                'Get QPU
                                strSQL = " SELECT Round(QPU,4) QPU  FROM  TblSemi  "
                                strSQL += " WHERE MaterialType = '09' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                                cnSQLRM = New SqlConnection(C1.Strcon)
                                cnSQLRM.Open()
                                cmSQLRM = New SqlCommand(strSQL, cnSQLRM)

                                QWireChafer = cmSQLRM.ExecuteScalar()
                                importRow(0)("QTU") = QWireChafer

                                cmSQLRM.Dispose()
                                cnSQLRM.Dispose()
                            End If
                        Else
                            '// Do nothing
                        End If
#End Region

#Region "Material Type NylonChafer"
                        '// NylonChafer
                        importRow = ImportTable.Select("GreenTire = '" & strGreenTireCode & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'Nylon CHAFER'")
                        If importRow.Count > 0 Then
                            If importRow(0)("Num").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Num value of type 'Nylon Chafer'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Num"), dblNum) Then
                                    Throw New System.Exception("Please input Num data as Number of type 'Nylon Chafer'.")
                                End If
                            End If
                            If importRow(0)("Length").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Length value of type 'Nylon Chafer'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Length"), dblNum) Then
                                    Throw New System.Exception("Please input Length data as Number of type 'Nylon Chafer'.")
                                End If
                            End If

                            strSQL = " SELECT COUNT(*)  FROM  TblSemi  "
                            strSQL += " WHERE MaterialType = '10' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                            cnSQLRM = New SqlConnection(C1.Strcon)
                            cnSQLRM.Open()
                            cmSQLRM = New SqlCommand(strSQL, cnSQLRM)
                            Dim i As Long = cmSQLRM.ExecuteScalar()
                            If i = 0 Then
                                Throw New System.Exception("Please check correctly Nylon Chafer '" & importRow(0)("SemiCode").ToString().Trim() & "'")
                            Else
                                'Get QPU
                                strSQL = " SELECT Round(QPU,4) QPU  FROM  TblSemi  "
                                strSQL += " WHERE MaterialType = '10' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                                cnSQLRM = New SqlConnection(C1.Strcon)
                                cnSQLRM.Open()
                                cmSQLRM = New SqlCommand(strSQL, cnSQLRM)

                                QNylonChafer = cmSQLRM.ExecuteScalar()
                                importRow(0)("QTU") = QNylonChafer

                                cmSQLRM.Dispose()
                                cnSQLRM.Dispose()
                            End If
                        Else
                            '// Do nothing
                        End If
#End Region

#Region "Material Type Flipper"
                        importRow = ImportTable.Select("GreenTire = '" & strGreenTireCode & "' AND Revision = '" & strRevision & "' AND TypeMaterial = 'FLIPPER'")
                        If importRow.Count > 0 Then
                            If importRow(0)("Num").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Num value of type 'Flipper'")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Num"), dblNum) Then
                                    Throw New System.Exception("Please input Num data as Number of type 'Flipper'.")
                                End If
                            End If
                            If importRow(0)("Length").ToString().Length <= 0 Then
                                Throw New System.Exception("Please check Length value of type 'Flipper'.")
                            Else
                                '//Check type number
                                If Not Double.TryParse(importRow(0)("Length"), dblNum) Then
                                    Throw New System.Exception("Please input Length data as Number of type 'Flipper'.")
                                End If
                            End If

                            strSQL = " SELECT COUNT(*)  FROM  TblSemi  "
                            strSQL += " WHERE MaterialType = '22' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                            cnSQLRM = New SqlConnection(C1.Strcon)
                            cnSQLRM.Open()
                            cmSQLRM = New SqlCommand(strSQL, cnSQLRM)
                            Dim i As Long = cmSQLRM.ExecuteScalar()
                            If i = 0 Then
                                Throw New System.Exception("Please check correctly Flipper '" & importRow(0)("SemiCode").ToString().Trim() & "'")
                            Else
                                'Get QPU
                                strSQL = " SELECT Round(QPU,4) QPU  FROM  TblSemi  "
                                strSQL += " WHERE MaterialType = '22' AND Active = '1' AND Final = '" & importRow(0)("SemiCode").ToString().Trim() & "' "
                                cnSQLRM = New SqlConnection(C1.Strcon)
                                cnSQLRM.Open()
                                cmSQLRM = New SqlCommand(strSQL, cnSQLRM)

                                QFlipper = cmSQLRM.ExecuteScalar()
                                importRow(0)("QTU") = QFlipper

                                cmSQLRM.Dispose()
                                cnSQLRM.Dispose()
                            End If
                        Else
                            '// Do nothing
                        End If
#End Region
                        '//**********************************************************************************************************************************
                    End If 'If strGreenTireCode <> chkSameGreenTireBefore Or strRevision <> chkSameRevisionBefore
                End If 'If strGreenTireCode.Length > 0

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

    Private Function GetTireSize() As DataTable
        Dim daSQL As SqlDataAdapter
        Dim strSQL As String = String.Empty
        Dim dt As New DataTable()
        Dim sb As New System.Text.StringBuilder()

        Try
            sb.AppendLine(" SELECT BSJCode FROM  TblTiresize ")
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
    Private Sub CmbTire_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbTire.SelectedIndexChanged
        CheckBox()
    End Sub

    Private Sub CheckBoxTire_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxTire.CheckedChanged
        CheckBox()
    End Sub

    Private Sub CheckBoxType_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        CheckBox()
    End Sub

    Sub CheckBox()
        If CheckBoxTire.Checked = True And txtsize.Text.Trim <> "" Then
            CmbTire.Enabled = True
            GrdDV.RowFilter = " Tirecode  like'%" & CmbTire.Text.Trim & "%'" _
                                   & " and Tsize like '%" & txtsize.Text.Trim & "%' "
            DataGridCOM.DataSource = GrdDV
        ElseIf CheckBoxTire.Checked = True And txtsize.Text.Trim = "" Then
            CmbTire.Enabled = True
            GrdDV.RowFilter = " Tirecode  like'%" & CmbTire.Text.Trim & "%'"
            DataGridCOM.DataSource = GrdDV
        ElseIf CheckBoxTire.Checked = False And txtsize.Text.Trim <> "" Then
            CmbTire.Enabled = False
            GrdDV.RowFilter = " Tsize like '%" & txtsize.Text.Trim & "%' "
            DataGridCOM.DataSource = GrdDV
        Else
            CmbTire.Enabled = False
            GrdDV.RowFilter = " "
            DataGridCOM.DataSource = GrdDV
        End If

        If CHKActive.Checked Then
            If CheckBoxTire.Checked = False And txtsize.Text.Trim = "" Then
                GrdDV.RowFilter &= "  ac = 1 "
                DataGridCOM.DataSource = GrdDV
            Else
                GrdDV.RowFilter &= "  and  ac = 1 "
                DataGridCOM.DataSource = GrdDV
            End If
        Else
            GrdDV.RowFilter &= " "
            DataGridCOM.DataSource = GrdDV
        End If
    End Sub

    Private Sub CmbMaterial_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        CheckBox()
    End Sub
#End Region

    Private Sub CmdDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdDel.Click

        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult

        msg = "Delete Green Tire : " & GrdDV.Item(oldrow).Row("final") _
            & " Revision : " & GrdDV.Item(oldrow).Row("Rev")  'Define message.
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Green Tire"   ' Define title.
        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            If ChkData() Then
                DelTire()
                LoadTire()
                CheckBox()
                oldrow = 0
            Else
                MsgBox("Can't Delete. Please check Usage.", MsgBoxStyle.OKOnly, "Tire")
            End If
        Else
            Exit Sub
        End If
    End Sub

#Region "Del"
    Private Function ChkData() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Try
            strSQL &= " select count(*) from TblGTHdr "
            strSQL &= " where Tirecode  = '" & GrdDV.Item(oldrow).Row("final") & "'"
            strSQL &= " and Rev  = '" & GrdDV.Item(oldrow).Row("Rev") & "'"
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
    Private Function ChkGP() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Try
            strSQL &= " select count(*) from TblGTHdr "
            strSQL &= " where Tirecode  = '" & GrdDV.Item(oldrow).Row("final") & "'"
            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar
            If i = 1 Then
                ChkGP = True
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
    Sub DelTire()
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Try
            strSQL = " Delete TblGTHdr"
            strSQL &= " where TireCode = '" & GrdDV.Item(oldrow).Row("final") & "'"
            strSQL &= " and Rev = '" & GrdDV.Item(oldrow).Row("Rev") & "'"
            strSQL &= "  "
            strSQL &= " Delete TblGTDtl"
            strSQL &= " where Tirecode = '" & GrdDV.Item(oldrow).Row("final") & "'"
            strSQL &= " and Rev = '" & GrdDV.Item(oldrow).Row("Rev") & "'"
            strSQL &= "  "
            strSQL &= " Delete TblConvert"
            strSQL &= " where code = '" & GrdDV.Item(oldrow).Row("final") & "'"
            strSQL &= " and Rev = '" & GrdDV.Item(oldrow).Row("Rev") & "'"
            strSQL &= "  "
            If ChkGP() Then
                strSQL &= " Delete TblGroup"
                strSQL &= " where Code = '" & GrdDV.Item(oldrow).Row("final") & "'"
                strSQL &= "  "
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

    Private Sub cmdView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdView.Click
        CheckBox()
    End Sub

    Private Sub CHKActive_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CHKActive.CheckedChanged
        CheckBox()
    End Sub

    Private Sub cmdActive_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdActive.Click
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult

        msg = "Change Active Green Tire : " & GrdDV.Item(oldrow).Row("final") _
             & " Revision : " & GrdDV.Item(oldrow).Row("Rev")  'Define message.
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Green Tire"   ' Define title.
        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            If ChkData() Then
                UpTire()
                LoadTire()
                CheckBox()
                oldrow = 0
            Else
                MsgBox("Can't Change. Please check Usage.", MsgBoxStyle.OKOnly, "Tire")
            End If
        Else
            Exit Sub
        End If
    End Sub
    Sub UpTire()
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Try
            strSQL = " Update TblGTHdr"
            strSQL &= " set Active = 0 "
            strSQL &= " where TireCode = '" & GrdDV.Item(oldrow).Row("final") & "'"
            strSQL &= "  "
            strSQL &= " Update TblGTHdr"
            strSQL &= " set Active = 1 "
            strSQL &= " where TireCode = '" & GrdDV.Item(oldrow).Row("final") & "'"
            strSQL &= " and Rev = '" & GrdDV.Item(oldrow).Row("Rev") & "'"

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

End Class
