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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.DataGridCOM = New System.Windows.Forms.DataGrid
        Me.CmdSave = New System.Windows.Forms.Button
        Me.CmdClose = New System.Windows.Forms.Button
        Me.CmdEdit = New System.Windows.Forms.Button
        Me.CmbTire = New System.Windows.Forms.ComboBox
        Me.CheckBoxTire = New System.Windows.Forms.CheckBox
        Me.CmdDel = New System.Windows.Forms.Button
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtsize = New System.Windows.Forms.TextBox
        Me.cmdView = New System.Windows.Forms.Button
        Me.CHKActive = New System.Windows.Forms.CheckBox
        Me.cmdActive = New System.Windows.Forms.Button
        Me.CmdImport = New System.Windows.Forms.Button
        Me.CmdExport = New System.Windows.Forms.Button
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
        Me.GroupBox1.Size = New System.Drawing.Size(824, 542)
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
        Me.DataGridCOM.Size = New System.Drawing.Size(818, 523)
        Me.DataGridCOM.TabIndex = 0
        '
        'CmdSave
        '
        Me.CmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdSave.Image = CType(resources.GetObject("CmdSave.Image"), System.Drawing.Image)
        Me.CmdSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdSave.Location = New System.Drawing.Point(600, 616)
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
        Me.CmdClose.Location = New System.Drawing.Point(752, 616)
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
        Me.CmdEdit.Location = New System.Drawing.Point(680, 616)
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
        Me.CmdDel.Location = New System.Drawing.Point(8, 614)
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
        Me.CmdImport.Location = New System.Drawing.Point(402, 616)
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
        Me.CmdExport.Location = New System.Drawing.Point(482, 616)
        Me.CmdExport.Name = "CmdExport"
        Me.CmdExport.Size = New System.Drawing.Size(80, 56)
        Me.CmdExport.TabIndex = 29
        Me.CmdExport.Text = "Export"
        Me.CmdExport.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'FrmGreenTire
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(840, 678)
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
        Me.Name = "FrmGreenTire"
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

#Region "Function_Load"
    Private Sub LoadTire()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = " select * from ("
        StrSQL &= " select final,TireSize,Round(Qty,1) TQty,"
        StrSQL &= "  substring(DateUp,7,2)+'/'+substring(DateUp,5,2)+'/'"
        StrSQL &= "  +substring(DateUp,1,4) dateup,TireSize TSize,Tirecode,Rev,null MaterialName"
        StrSQL &= "  ,null Semicode,null Length,null number,null QTU,null Unit,Remark,Active,Active AC from TblGtHdr"
        StrSQL &= "  union"
        StrSQL &= "  select null final,null TireSize,null TQty,null dateup,tiresize TSize,dt.Tirecode,dt.Rev,MaterialName"
        StrSQL &= "  ,isnull(Semicode,'No Use') Semicode,Length,number, round(QTU,3) Qty, Unit ,null Remark ,null Active,Active AC from TblGtDtl dt"
        StrSQL &= "   left outer join TBLTypeMaterial tm"
        StrSQL &= "   on dt.MaterialType = tm.MaterialCode"
        StrSQL &= " left outer join  TBLGTHdr hd"
        StrSQL &= " on dt.tirecode+dt.Rev = hd.Tirecode+hd.Rev"
        StrSQL &= "   ) Tire"
        StrSQL &= "   order by Tirecode,Rev,Final Desc"
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
    {grdColStyle0_2, grdColStyle0, grdColStyle0_1, grdColStyle1, grdColStyle2, grdColStyle3, _
     grdColStyle12, grdColStyle11, grdColStyle9, grdColStyle6, _
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
        Dim dtPSemi As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        StrSQL = "   SELECT Code,MaterialName"
        StrSQL &= "   FROM  TblGroup g"
        StrSQL &= "  left outer join "
        StrSQL &= "  ("
        StrSQL &= "  SELECT  semiCode,MaterialName"
        StrSQL &= "   FROM  TblSemi p"
        StrSQL &= "  left outer join  TblTypeMaterial t"
        StrSQL &= "  on p.MaterialType = t.MaterialCode"
        StrSQL &= "  )semi"
        StrSQL &= "  on g.code = semi.semicode"
        StrSQL &= "  where Typecode = '06'"
        StrSQL &= "  order by Code"
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

    Private Sub CmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClose.Click
        Me.Close()
    End Sub

    Private Sub CmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdEdit.Click
        Dim Strremark() As String
        Dim i As Integer
        Dim FAddGreenTire As New FrmAddGreenTire
        FAddGreenTire.Text = "Edit GreenTire"
        FAddGreenTire.CmdSave.Text = "Edit"
        If GrdDV.Item(oldrow).Row("Final") = "" Then
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

            For i = 0 To 12
                If i = 1 Then
                    FAddGreenTire.TxtB1_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                    FAddGreenTire.TxtB1_L.Text = GrdDV.Item(oldrow + i).Row("Length")
                    FAddGreenTire.B1code = GrdDV.Item(oldrow + i).Row("Semicode")
                End If
                If i = 2 Then
                    FAddGreenTire.TxtB2_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                    FAddGreenTire.TxtB2_L.Text = GrdDV.Item(oldrow + i).Row("Length")
                    FAddGreenTire.B2code = GrdDV.Item(oldrow + i).Row("Semicode")
                End If
                If i = 3 Then
                    FAddGreenTire.TxtB3_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                    FAddGreenTire.TxtB3_L.Text = GrdDV.Item(oldrow + i).Row("Length")
                    FAddGreenTire.B3code = GrdDV.Item(oldrow + i).Row("Semicode")
                End If
                If i = 4 Then
                    FAddGreenTire.TxtB4_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                    FAddGreenTire.TxtB4_L.Text = GrdDV.Item(oldrow + i).Row("Length")
                    FAddGreenTire.B4code = GrdDV.Item(oldrow + i).Row("Semicode")
                End If
                If i = 5 Then
                    FAddGreenTire.txtBF_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                    FAddGreenTire.BFcode = GrdDV.Item(oldrow + i).Row("Semicode")
                End If
                If i = 6 Then
                    FAddGreenTire.TxtBP_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                    FAddGreenTire.TxtBp_L.Text = GrdDV.Item(oldrow + i).Row("Length")
                    FAddGreenTire.Bpcode = GrdDV.Item(oldrow + i).Row("Semicode")
                End If
                If i = 7 Then
                    FAddGreenTire.TxtCu_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                    FAddGreenTire.TxtCU_L.Text = GrdDV.Item(oldrow + i).Row("Length")
                    FAddGreenTire.CUcode = GrdDV.Item(oldrow + i).Row("Semicode")
                End If
                If i = 8 Then
                    FAddGreenTire.TxtIN_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                    FAddGreenTire.TxtIN_L.Text = GrdDV.Item(oldrow + i).Row("Length")
                    FAddGreenTire.INcode = GrdDV.Item(oldrow + i).Row("Semicode")
                End If
                If i = 9 Then
                    If GrdDV.Item(oldrow + i).Row("Semicode") <> "No Use" Then
                        FAddGreenTire.CheckBoxNy.Checked = True
                        FAddGreenTire.TxtNy_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                        FAddGreenTire.TxtNy_L.Text = GrdDV.Item(oldrow + i).Row("Length")
                        FAddGreenTire.NYcode = GrdDV.Item(oldrow + i).Row("Semicode")
                    Else
                    End If
                End If
                If i = 10 Then
                    FAddGreenTire.TxtSD_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                    FAddGreenTire.TxtSD_L.Text = GrdDV.Item(oldrow + i).Row("Length")
                    FAddGreenTire.SDcode = GrdDV.Item(oldrow + i).Row("Semicode")
                End If
                If i = 11 Then
                    FAddGreenTire.TxtTT_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                    FAddGreenTire.TTcode = GrdDV.Item(oldrow + i).Row("Semicode")
                End If
                If i = 12 Then
                    If GrdDV.Item(oldrow + i).Row("Semicode") <> "No Use" Then
                        FAddGreenTire.CheckBoxWf.Checked = True
                        FAddGreenTire.TxtWf_N.Text = GrdDV.Item(oldrow + i).Row("Number")
                        FAddGreenTire.TxtWf_L.Text = GrdDV.Item(oldrow + i).Row("Length")
                        FAddGreenTire.Wfcode = GrdDV.Item(oldrow + i).Row("Semicode")
                    Else
                    End If
                End If
            Next

            FAddGreenTire.ShowDialog()
            LoadTire()
            CheckBox()
            oldrow = 0
        End If
    End Sub

    Private Sub CmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSave.Click
        Dim ttype As String
        Dim FAddGreenTire As New FrmAddGreenTire
        ttype = CmbTire.Text.Trim
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
