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

Public Class FrmCALMaster
#Region "Declare"
    Inherits System.Windows.Forms.Form
    Dim GrdDV As New DataView
    Protected Const TBL_Cal As String = "TBL_Cal"
    Dim GrdDVGP As New DataView
    Protected Const TBL_GP As String = "TBL_GP"
    Dim C1 As New SQLData("ACCINV")
    Dim vBal As Boolean
    Protected DefaultGridBorderStyle As BorderStyle
    Friend WithEvents ButtonExport As Button
#End Region
    Friend txtname As String

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
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DataGridCAL As System.Windows.Forms.DataGrid
    Friend WithEvents lblCal As System.Windows.Forms.Label
    Friend WithEvents ButtonClose As System.Windows.Forms.Button
    Friend WithEvents DateTime As System.Windows.Forms.DateTimePicker
    Friend WithEvents GroupCompound As System.Windows.Forms.GroupBox
    Friend WithEvents CheckCompound As System.Windows.Forms.CheckBox
    Friend WithEvents ComboBoxComp As System.Windows.Forms.ComboBox
    Friend WithEvents CheckCompGroup As System.Windows.Forms.CheckBox
    Friend WithEvents cmbType As System.Windows.Forms.ComboBox
    Friend WithEvents CheckType As System.Windows.Forms.CheckBox
    Friend WithEvents GType As System.Windows.Forms.GroupBox
    Friend WithEvents CmdView As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmCALMaster))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.DataGridCAL = New System.Windows.Forms.DataGrid()
        Me.ButtonClose = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.DateTime = New System.Windows.Forms.DateTimePicker()
        Me.CmdView = New System.Windows.Forms.Button()
        Me.lblCal = New System.Windows.Forms.Label()
        Me.GroupCompound = New System.Windows.Forms.GroupBox()
        Me.ComboBoxComp = New System.Windows.Forms.ComboBox()
        Me.CheckCompGroup = New System.Windows.Forms.CheckBox()
        Me.CheckCompound = New System.Windows.Forms.CheckBox()
        Me.GType = New System.Windows.Forms.GroupBox()
        Me.cmbType = New System.Windows.Forms.ComboBox()
        Me.CheckType = New System.Windows.Forms.CheckBox()
        Me.ButtonExport = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridCAL, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupCompound.SuspendLayout()
        Me.GType.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.DataGridCAL)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 112)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1168, 440)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'DataGridCAL
        '
        Me.DataGridCAL.DataMember = ""
        Me.DataGridCAL.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridCAL.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridCAL.Location = New System.Drawing.Point(3, 16)
        Me.DataGridCAL.Name = "DataGridCAL"
        Me.DataGridCAL.Size = New System.Drawing.Size(1162, 421)
        Me.DataGridCAL.TabIndex = 14
        '
        'ButtonClose
        '
        Me.ButtonClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonClose.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.ButtonClose.Image = CType(resources.GetObject("ButtonClose.Image"), System.Drawing.Image)
        Me.ButtonClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonClose.Location = New System.Drawing.Point(1104, 560)
        Me.ButtonClose.Name = "ButtonClose"
        Me.ButtonClose.Size = New System.Drawing.Size(72, 56)
        Me.ButtonClose.TabIndex = 10
        Me.ButtonClose.Text = "CLOSE"
        Me.ButtonClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label3.Location = New System.Drawing.Point(1016, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 16)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Date"
        '
        'DateTime
        '
        Me.DateTime.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DateTime.CustomFormat = "dd/MM/yyyy"
        Me.DateTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTime.Location = New System.Drawing.Point(1064, 14)
        Me.DateTime.Name = "DateTime"
        Me.DateTime.Size = New System.Drawing.Size(104, 20)
        Me.DateTime.TabIndex = 12
        '
        'CmdView
        '
        Me.CmdView.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdView.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.CmdView.Image = CType(resources.GetObject("CmdView.Image"), System.Drawing.Image)
        Me.CmdView.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdView.Location = New System.Drawing.Point(1096, 48)
        Me.CmdView.Name = "CmdView"
        Me.CmdView.Size = New System.Drawing.Size(75, 56)
        Me.CmdView.TabIndex = 11
        Me.CmdView.Text = "View"
        Me.CmdView.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'lblCal
        '
        Me.lblCal.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblCal.Location = New System.Drawing.Point(16, 16)
        Me.lblCal.Name = "lblCal"
        Me.lblCal.Size = New System.Drawing.Size(488, 32)
        Me.lblCal.TabIndex = 10
        Me.lblCal.Text = "Price  of"
        '
        'GroupCompound
        '
        Me.GroupCompound.Controls.Add(Me.ComboBoxComp)
        Me.GroupCompound.Controls.Add(Me.CheckCompGroup)
        Me.GroupCompound.Controls.Add(Me.CheckCompound)
        Me.GroupCompound.Location = New System.Drawing.Point(8, 48)
        Me.GroupCompound.Name = "GroupCompound"
        Me.GroupCompound.Size = New System.Drawing.Size(320, 64)
        Me.GroupCompound.TabIndex = 14
        Me.GroupCompound.TabStop = False
        Me.GroupCompound.Text = "Compound"
        Me.GroupCompound.Visible = False
        '
        'ComboBoxComp
        '
        Me.ComboBoxComp.Location = New System.Drawing.Point(136, 38)
        Me.ComboBoxComp.Name = "ComboBoxComp"
        Me.ComboBoxComp.Size = New System.Drawing.Size(152, 21)
        Me.ComboBoxComp.TabIndex = 2
        Me.ComboBoxComp.Text = "Select"
        '
        'CheckCompGroup
        '
        Me.CheckCompGroup.Location = New System.Drawing.Point(8, 40)
        Me.CheckCompGroup.Name = "CheckCompGroup"
        Me.CheckCompGroup.Size = New System.Drawing.Size(128, 16)
        Me.CheckCompGroup.TabIndex = 1
        Me.CheckCompGroup.Text = "Group Compound"
        '
        'CheckCompound
        '
        Me.CheckCompound.Location = New System.Drawing.Point(8, 18)
        Me.CheckCompound.Name = "CheckCompound"
        Me.CheckCompound.Size = New System.Drawing.Size(120, 16)
        Me.CheckCompound.TabIndex = 0
        Me.CheckCompound.Text = "Final Compound  "
        '
        'GType
        '
        Me.GType.Controls.Add(Me.cmbType)
        Me.GType.Controls.Add(Me.CheckType)
        Me.GType.Location = New System.Drawing.Point(8, 56)
        Me.GType.Name = "GType"
        Me.GType.Size = New System.Drawing.Size(360, 56)
        Me.GType.TabIndex = 15
        Me.GType.TabStop = False
        Me.GType.Text = "Material Type"
        Me.GType.Visible = False
        '
        'cmbType
        '
        Me.cmbType.Location = New System.Drawing.Point(80, 22)
        Me.cmbType.Name = "cmbType"
        Me.cmbType.Size = New System.Drawing.Size(152, 21)
        Me.cmbType.TabIndex = 2
        Me.cmbType.Text = "Select"
        '
        'CheckType
        '
        Me.CheckType.Location = New System.Drawing.Point(8, 24)
        Me.CheckType.Name = "CheckType"
        Me.CheckType.Size = New System.Drawing.Size(64, 16)
        Me.CheckType.TabIndex = 1
        Me.CheckType.Text = "Type"
        '
        'ButtonExport
        '
        Me.ButtonExport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonExport.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.ButtonExport.Image = CType(resources.GetObject("ButtonExport.Image"), System.Drawing.Image)
        Me.ButtonExport.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonExport.Location = New System.Drawing.Point(1030, 560)
        Me.ButtonExport.Name = "ButtonExport"
        Me.ButtonExport.Size = New System.Drawing.Size(72, 56)
        Me.ButtonExport.TabIndex = 16
        Me.ButtonExport.Text = "Export"
        Me.ButtonExport.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'FrmCALMaster
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1184, 630)
        Me.Controls.Add(Me.ButtonExport)
        Me.Controls.Add(Me.GroupCompound)
        Me.Controls.Add(Me.lblCal)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ButtonClose)
        Me.Controls.Add(Me.CmdView)
        Me.Controls.Add(Me.DateTime)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.GType)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmCALMaster"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Price Master (Material) -"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataGridCAL, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupCompound.ResumeLayout(False)
        Me.GType.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "CONSTANT"
    Dim DT As New DataTable
    Dim StrSQL As String
    Dim oldrow As Integer
#End Region

#Region "Form Event"
    Private Sub FrmCALMaster_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If txtname.Trim().Equals("RM") Then
            LoadRM()
        ElseIf txtname.Trim().Equals("Pigment") Then
            LoadPigment()
        ElseIf txtname.Trim().Equals("Compound") Then
            LoadCompound()
            LoadGroup()
            GroupCompound.Visible = True 'Show groupbox compound
        ElseIf txtname.Trim().Equals("PreSemi") Then
            LoadPresemi()
            LoadMaterialType()
            GType.Visible = True 'Show groupbox material type
        ElseIf txtname.Trim().Equals("Semi") Then
            Loadsemi()
            LoadMaterialType()
            GType.Visible = True 'Show groupbox material type
        ElseIf txtname.Trim().Equals("Green Tire") Then
            LoadTire()
        Else
            'Nothing
        End If

        vBal = False
        SetTotal() 'Set number of items
    End Sub
#End Region

#Region "Function_Load"
    Private Sub LoadRM()
        Dim sb As New System.Text.StringBuilder()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sb.AppendLine(" SELECT *")
        sb.AppendLine(" FROM TblRM r")
        sb.AppendLine(" LEFT OUTER JOIN ( ")
        sb.AppendLine("   SELECT Mastercode,StdPrice,ActPrice")
        sb.AppendLine("   ,substring(dateup,7,2)+'/'+substring(dateup,5,2)+'/'+substring(dateup,1,4) DateUp")
        sb.AppendLine("   ,substring(Timeup,1,2)+':'+substring(Timeup,3,2) Timeup")
        sb.AppendLine("   FROM TBLMasterPrice ")
        sb.AppendLine("   WHERE Typecode IN ('01','07','08','09')") 'R/M Material, IDM, Test and Other
        sb.AppendLine(" ) m on r.rmcode = m.mastercode")
        sb.AppendLine(" ORDER BY descName, mastercode")
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
        DT.TableName = TBL_Cal
        GrdDV = DT.DefaultView
        GrdDV.AllowNew = False
        GrdDV.AllowDelete = False
        '************************************
        DataGridCAL.DataSource = GrdDV
        '************************************
        'Dim i As Integer
        'Dim c34 As String = Chr(34)
        'For i = 0 To DT.Columns.Count - 1
        '    Dim col As String = DT.Columns(i).ColumnName
        '    Dim coltype As String = DT.Columns(i).DataType.FullName
        '    coltype = coltype.Replace("System.", "")
        '    coltype = coltype.Replace("Int32", "integer")
        '    coltype = coltype.Replace("Int16", "integer")
        '    coltype = coltype.Replace("String", "string")
        '    coltype = coltype.Replace("Decimal", "decimal")
        '    Debug.WriteLine("<xs:element name=" & c34 & col.Trim & c34 & "  type= " & c34 & "xs:" & coltype & c34 & " minOccurs=" & c34 & "0" & c34 & "/>")
        'Next
        ResetTableStyle()

        With DataGridCAL
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
            .MappingName = TBL_Cal
            .PreferredColumnWidth = 125
            .PreferredRowHeight = 15
        End With
        Dim grdColStyle1 As New DataGridColoredLine2
        With grdColStyle1
            .HeaderText = "Code"
            .MappingName = "Rmcode"
            .Width = 100
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle1_1 As New DataGridColoredLine2
        With grdColStyle1_1
            .HeaderText = "DescName"
            .MappingName = "DescName"
            .Width = 120
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle2 As New DataGridColoredLine2
        With grdColStyle2
            .HeaderText = "@ STD Price"
            .MappingName = "STDPrice"
            .Width = 100
            .Format = "##,###,##0.00"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle3 As New DataGridColoredLine2
        With grdColStyle3
            .HeaderText = "@ ACT Price"
            .MappingName = "ActPrice"
            .Width = 100
            .Format = "##,###,##0.00"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle4 As New DataGridColoredLine2
        With grdColStyle4
            .HeaderText = "Date"
            .MappingName = "dateup"
            .Width = 85
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle5 As New DataGridColoredLine2
        With grdColStyle5
            .HeaderText = "Time"
            .MappingName = "TimeUp"
            .Width = 85
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        grdTableStyle1.GridColumnStyles.AddRange _
    (New DataGridColumnStyle() _
    {grdColStyle1, grdColStyle1_1, grdColStyle2, grdColStyle3, grdColStyle4, grdColStyle5})

        DataGridCAL.TableStyles.Add(grdTableStyle1)
        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub
    Private Sub LoadPigment()
        Dim sb As New System.Text.StringBuilder()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sb.AppendLine(" SELECT *,bb.Qty*stdPrice STD,bb.Qty*actPrice Act")
        sb.AppendLine(" FROM (")
        sb.AppendLine("   SELECT Mastercode,StdPrice,ActPrice")
        sb.AppendLine("   ,substring(dateup,7,2)+'/'+substring(dateup,5,2)+'/'+substring(dateup,1,4) DateUp")
        sb.AppendLine("   ,substring(Timeup,1,2)+':'+substring(Timeup,3,2) Timeup")
        sb.AppendLine("   FROM TBLMasterPrice ")
        sb.AppendLine("   WHERE Typecode = '02'")
        sb.AppendLine(" ) aa")
        sb.AppendLine(" LEFT OUTER JOIN TBLPigment bb on aa.mastercode = bb.pigmentcode ")
        sb.AppendLine(" Order by Mastercode")
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
        DT.TableName = TBL_Cal
        GrdDV = DT.DefaultView
        GrdDV.AllowNew = False
        GrdDV.AllowDelete = False
        '************************************
        DataGridCAL.DataSource = GrdDV
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

        With DataGridCAL
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
            .MappingName = TBL_Cal
            .PreferredColumnWidth = 125
            .PreferredRowHeight = 15
        End With

        Dim grdColStyle1 As New DataGridColoredLine2
        With grdColStyle1
            .HeaderText = "Code (Material)"
            .MappingName = "Mastercode"
            .Width = 150
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle2 As New DataGridColoredLine2
        With grdColStyle2
            .HeaderText = "@ STD/KG"
            .MappingName = "StdPrice"
            .NullText = ""
            .Width = 115
            .Format = "##,###,###.0000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle3 As New DataGridColoredLine2
        With grdColStyle3
            .HeaderText = "@ ACT/KG"
            .MappingName = "ACTPrice"
            .NullText = ""
            .Width = 115
            .Format = "##,###,###.0000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle4 As New DataGridColoredLine2
        With grdColStyle4
            .HeaderText = "@ STD/BT"
            .MappingName = "Std"
            .NullText = ""
            .Width = 115
            .Format = "##,###,###.0000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle5 As New DataGridColoredLine2
        With grdColStyle5
            .HeaderText = "@ ACT/BT"
            .MappingName = "ACT"
            .NullText = ""
            .Width = 115
            .Format = "##,###,###.0000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle6 As New DataGridColoredLine2
        With grdColStyle6
            .HeaderText = "Date"
            .MappingName = "dateup"
            .Width = 85
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle7 As New DataGridColoredLine2
        With grdColStyle7
            .HeaderText = "Time"
            .MappingName = "TimeUp"
            .Width = 85
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        grdTableStyle1.GridColumnStyles.AddRange _
(New DataGridColumnStyle() _
{grdColStyle1, grdColStyle2, grdColStyle3 _
, grdColStyle4, grdColStyle5, grdColStyle6, grdColStyle7})

        DataGridCAL.TableStyles.Add(grdTableStyle1)
        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub
    Private Sub LoadCompound()
        Dim sb As New System.Text.StringBuilder()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sb.AppendLine(" SELECT aa.FinalCompound,aa.Compcode,aa.Revision,STD,ACT,aa.dateup,Timeup")
        sb.AppendLine(" ,aa.active, bb.Qty*STD STDBT,bb.Qty*ACT ACTBT")
        sb.AppendLine(" FROM (")
        sb.AppendLine("   SELECT    FinalCompound,c.Compcode,c.Revision,StdPrice STD,ActPrice ACT")
        sb.AppendLine("   ,substring(p.DateUp,7,2)+'/'+substring(p.DateUp,5,2)+'/'+substring(p.DateUp,1,4) dateup")
        sb.AppendLine("   ,substring(p.TimeUp,1,2)+':'+substring(p.TimeUp,3,2) Timeup, active")
        sb.AppendLine("   FROM         TBLCompound c")
        sb.AppendLine("   LEFT OUTER JOIN (")
        sb.AppendLine("     SELECT    *")
        sb.AppendLine("     FROM         TBLMASTERPRICE")
        sb.AppendLine("     WHERE Mastercode+revision IN ( SELECT Compcode+revision FROM TBLCompound )")
        sb.AppendLine("   ) p on c.Compcode+c.Revision = p.Mastercode+p.Revision ")
        sb.AppendLine(" ) aa")
        sb.AppendLine(" LEFT OUTER JOIN TBLcompound bb on aa.compcode+aa.Revision = bb.Compcode+bb.Revision")
        sb.AppendLine(" ORDER BY aa.Finalcompound,aa.Compcode ")
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
        DT.TableName = TBL_Cal
        GrdDV = DT.DefaultView
        GrdDV.AllowNew = False
        GrdDV.AllowDelete = False
        '************************************
        DataGridCAL.DataSource = GrdDV
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

        With DataGridCAL
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
            .MappingName = TBL_Cal
            .PreferredColumnWidth = 125
            .PreferredRowHeight = 15
        End With

        Dim grdColStyle0 As New DataGridColoredLine2
        With grdColStyle0
            .HeaderText = "Code (Material)"
            .MappingName = "FinalCompound"
            .Width = 150
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle1 As New DataGridColoredLine2
        With grdColStyle1
            .HeaderText = "CompoundCode"
            .MappingName = "Compcode"
            .Width = 150
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle2 As New DataGridColoredLine2
        With grdColStyle2
            .HeaderText = "Rev."
            .MappingName = "Revision"
            .Width = 80
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle3 As New DataGridColoredLine2
        With grdColStyle3
            .HeaderText = "@ STD/KG"
            .MappingName = "Std"
            .NullText = ""
            .Width = 115
            .Format = "##,###,##0.0000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle4 As New DataGridColoredLine2
        With grdColStyle4
            .HeaderText = "@ ACT/KG"
            .MappingName = "Act"
            .NullText = ""
            .Width = 115
            .Format = "##,###,##0.0000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle5 As New DataGridColoredLine2
        With grdColStyle5
            .HeaderText = "@ STD/BT"
            .MappingName = "StdBT"
            .NullText = ""
            .Width = 115
            .Format = "##,###,##0.0000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle6 As New DataGridColoredLine2
        With grdColStyle6
            .HeaderText = "@ ACT/BT"
            .MappingName = "ActBT"
            .NullText = ""
            .Width = 115
            .Format = "##,###,##0.0000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle7 As New DataGridColoredLine2
        With grdColStyle7
            .HeaderText = "Date"
            .MappingName = "DateUp"
            .Width = 80
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle8 As New DataGridColoredLine2
        With grdColStyle8
            .HeaderText = "Time"
            .MappingName = "TimeUp"
            .Width = 80
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        grdTableStyle1.GridColumnStyles.AddRange _
    (New DataGridColumnStyle() _
    {grdColStyle0, grdColStyle1, grdColStyle2, grdColStyle3 _
    , grdColStyle4, grdColStyle5, grdColStyle6, grdColStyle7, grdColStyle8})

        DataGridCAL.TableStyles.Add(grdTableStyle1)
        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub
    Private Sub LoadPresemi()
        Dim sb As New System.Text.StringBuilder()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sb.AppendLine(" SELECT *   ")
        sb.AppendLine(" FROM (")
        sb.AppendLine("   SELECT Pre.Mastercode, std, Act,Width WLN")
        sb.AppendLine("   ,round(std/(width/1000),4) STDM, round(Act/(width/1000),4) ACTM")
        sb.AppendLine("   ,Pre.DateUp,active,MaterialType MT")
        sb.AppendLine("   FROM ( ")
        sb.AppendLine("     SELECT    Mastercode,Revision,StdPrice std,ActPrice Act")
        sb.AppendLine("     ,substring(DateUp,7,2)+'/'+substring(DateUp,5,2)+'/'+substring(DateUp,1,4) dateup ")
        sb.AppendLine("     ,substring(TimeUp,1,2)+':'+substring(TimeUp,3,2) Timeup")
        sb.AppendLine("     FROM         TBLMASTERPRICE")
        sb.AppendLine("     WHERE Mastercode+revision IN ( SELECT    Psemicode+revision FROM         TBLPresemi)")
        sb.AppendLine("   ) pre")
        sb.AppendLine("   LEFT OUTER JOIN TBLPresemi se on pre.Mastercode+pre.Revision = se.psemicode+se.Revision")
        sb.AppendLine("   WHERE MaterialType IN ('02')")
        sb.AppendLine("   UNION ")
        sb.AppendLine("   SELECT Pre.Mastercode, std, Act,Length WLN")
        sb.AppendLine("   ,round(std,4) STDM, round(Act,4) ACTM")
        sb.AppendLine("   ,Pre.DateUp,active,MaterialType MT ")
        sb.AppendLine("   FROM (")
        sb.AppendLine("     SELECT    Mastercode,Revision,StdPrice std,ActPrice Act")
        sb.AppendLine("     ,substring(DateUp,7,2)+'/'+substring(DateUp,5,2)+'/'+substring(DateUp,1,4) dateup ")
        sb.AppendLine("     ,substring(TimeUp,1,2)+':'+substring(TimeUp,3,2) Timeup")
        sb.AppendLine("     FROM         TBLMASTERPRICE")
        sb.AppendLine("     WHERE Mastercode+revision IN ( SELECT    Psemicode+revision FROM         TBLPresemi )")
        sb.AppendLine("   ) pre")
        sb.AppendLine("   LEFT OUTER JOIN TBLPresemi se on pre.Mastercode+pre.Revision = se.psemicode+se.Revision")
        sb.AppendLine("   WHERE MaterialType IN ('01')")
        sb.AppendLine("   UNION ")
        sb.AppendLine("   SELECT Pre.Mastercode, std, Act,n WLN")
        sb.AppendLine("   ,round(std*n,4) STDM, round(Act*n,4) ACTM")
        sb.AppendLine("   ,Pre.DateUp,active,MaterialType")
        sb.AppendLine("   FROM ( ")
        sb.AppendLine("     SELECT    Mastercode,Revision,StdPrice std,ActPrice Act")
        sb.AppendLine("     ,substring(DateUp,7,2)+'/'+substring(DateUp,5,2)+'/'+substring(DateUp,1,4) dateup ")
        sb.AppendLine("     ,substring(TimeUp,1,2)+':'+substring(TimeUp,3,2) Timeup")
        sb.AppendLine("     FROM         TBLMASTERPRICE")
        sb.AppendLine("     WHERE Mastercode+revision IN ( SELECT    Psemicode+revision FROM         TBLPresemi )")
        sb.AppendLine("   ) pre")
        sb.AppendLine("   LEFT OUTER JOIN TBLPresemi se on pre.Mastercode+pre.Revision = se.psemicode+se.Revision")
        sb.AppendLine("   WHERE MaterialType IN ('19')")
        sb.AppendLine("   UNION ")
        sb.AppendLine("   SELECT Pre.Mastercode, std, Act,Length WLN")
        sb.AppendLine("   ,round(std*(Length/1000),3,1) STDM, round(Act*(Length/1000),3,1) ACTM")
        sb.AppendLine("   ,Pre.DateUp,active,MaterialType")
        sb.AppendLine("   FROM ( ")
        sb.AppendLine("     SELECT    Mastercode,Revision,StdPrice std,ActPrice Act")
        sb.AppendLine("     ,substring(DateUp,7,2)+'/'+substring(DateUp,5,2)+'/'+substring(DateUp,1,4) dateup ")
        sb.AppendLine("     ,substring(TimeUp,1,2)+':'+substring(TimeUp,3,2) Timeup")
        sb.AppendLine("     FROM         TBLMASTERPRICE")
        sb.AppendLine("     WHERE Mastercode+revision IN ( SELECT    Psemicode+revision FROM         TBLPresemi )")
        sb.AppendLine("   ) pre")
        sb.AppendLine("   LEFT OUTER JOIN TBLPresemi se on pre.Mastercode+pre.Revision = se.psemicode+se.Revision")
        sb.AppendLine("   WHERE MaterialType NOT IN ('19','02','01')")
        sb.AppendLine(" ) PM ")
        sb.AppendLine(" LEFT OUTER JOIN (")
        sb.AppendLine("   SELECT code, Rev, std stdKG, act actKG ")
        sb.AppendLine("   FROM (")
        sb.AppendLine("     SELECT code,Rev,MaterialType,n,cn,Qty,Round(std/qty,3,1) STD,Round(act/qty,3,1) ACT")
        sb.AppendLine("     FROM (")
        sb.AppendLine("       SELECT * ")
        sb.AppendLine("       FROM TBLPresemi ")
        sb.AppendLine("       WHERE active = 1")
        sb.AppendLine("     ) p")
        sb.AppendLine("     LEFT OUTER JOIN (")
        sb.AppendLine("       SELECT code,Rev,isnull(width,'1000')/1000 wt")
        sb.AppendLine("       ,Round(sum(Qty)/1000,3,1) Qty,Round(sum(STD),4,1)STD,sum(ACT) ACT ")
        sb.AppendLine("       FROM TBLMasterPriceRM")
        sb.AppendLine("       WHERE code+Rev IN (SELECT psemicode+Revision FROM TBLPresemi)")
        sb.AppendLine("       GROUP BY code,Rev,width,length")
        sb.AppendLine("     ) m on p.psemicode+p.Revision = m.code+m.rev")
        sb.AppendLine("     WHERE materialType IN ('19')")
        sb.AppendLine("   ) xx")
        sb.AppendLine("   UNION")
        sb.AppendLine("   SELECT code, Rev, round(sum(std)/1000,3,1) STD, round(sum(Act)/1000,3,1) ACT")
        sb.AppendLine("   FROM (")
        sb.AppendLine("     SELECT code,rev,materialType,rmcode,Qty*lt Qty,StdPrice*Qty*lt std,ActPrice*Qty*lt act ")
        sb.AppendLine("     FROM (")
        sb.AppendLine("       SELECT code,rev,materialType,p.length/1000 lt,n,cn,rmcode,Qty,stdprice,actprice  ")
        sb.AppendLine("       FROM TBLPresemi p")
        sb.AppendLine("       LEFT OUTER JOIN ( ")
        sb.AppendLine("         SELECT * ")
        sb.AppendLine("         FROM TBLMasterPriceRM")
        sb.AppendLine("         WHERE code+Rev IN (SELECT psemicode+Revision FROM TBLPresemi)")
        sb.AppendLine("       ) m on p.psemicode+p.Revision = m.code+m.Rev")
        sb.AppendLine("       WHERE materialtype NOT IN ('01','02','19') AND active = 1")
        sb.AppendLine("     ) xx")
        sb.AppendLine("   ) xxx")
        sb.AppendLine("   GROUP BY code,Rev")
        sb.AppendLine("   UNION")
        sb.AppendLine("   SELECT code,Rev,round(std*(wt)/Qty,4,1) STDKG,round(act*(wt)/Qty,4,1) ACTKG ")
        sb.AppendLine("   FROM (")
        sb.AppendLine("     SELECT code,Rev,MaterialType,Width/1000 wt,Qty,std/(Width/1000) std,act/(Width/1000) act")
        sb.AppendLine("     FROM (")
        sb.AppendLine("       SELECT * ")
        sb.AppendLine("       FROM TBLPresemi ")
        sb.AppendLine("       WHERE active = 1")
        sb.AppendLine("     ) p")
        sb.AppendLine("     LEFT OUTER JOIN (")
        sb.AppendLine("       SELECT code,Rev,isnull(width,'1000')/1000 wt")
        sb.AppendLine("       ,sum(Qty)/1000 Qty,sum(STD)STD,sum(ACT) ACT ")
        sb.AppendLine("       FROM TBLMasterPriceRM")
        sb.AppendLine("       WHERE code+Rev IN (SELECT psemicode+Revision FROM TBLPresemi)")
        sb.AppendLine("       GROUP BY code,Rev,width,length")
        sb.AppendLine("     ) m on p.psemicode+p.Revision = m.code+m.rev")
        sb.AppendLine("     WHERE materialType in ('02') ")
        sb.AppendLine("     UNION")
        sb.AppendLine("     SELECT code,Rev,MaterialType,Length/1000 wt,Qty,std/(Length/1000) std,act/(Length/1000) act")
        sb.AppendLine("     FROM (")
        sb.AppendLine("       SELECT * ")
        sb.AppendLine("       FROM TBLPresemi ")
        sb.AppendLine("     ) p")
        sb.AppendLine("     LEFT OUTER JOIN (")
        sb.AppendLine("       SELECT code,Rev,isnull(Length,'1000')/1000 wt")
        sb.AppendLine("       ,sum(Qty) Qty,sum(STD)STD,sum(ACT) ACT ")
        sb.AppendLine("       FROM TBLMasterPriceRM")
        sb.AppendLine("       WHERE code+Rev IN (SELECT psemicode+Revision FROM TBLPresemi) ")
        sb.AppendLine("       GROUP BY code,Rev,width,length")
        sb.AppendLine("     ) m on p.psemicode+p.Revision = m.code+m.rev")
        sb.AppendLine("     WHERE materialType IN ('01') ")
        sb.AppendLine("   ) xx ")
        sb.AppendLine(" ) KG on pm.mastercode = kg.code")
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
        DT.TableName = TBL_Cal
        GrdDV = DT.DefaultView
        GrdDV.AllowNew = False
        GrdDV.AllowDelete = False
        '************************************
        DataGridCAL.DataSource = GrdDV
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

        With DataGridCAL
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
            .MappingName = TBL_Cal
            .PreferredColumnWidth = 125
            .PreferredRowHeight = 15
        End With

        Dim grdColStyle1 As New DataGridColoredLine2
        With grdColStyle1
            .HeaderText = "Code (Material)"
            .MappingName = "MasterCode"
            .Width = 150
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle3 As New DataGridColoredLine2
        With grdColStyle3
            .HeaderText = "@ STD of Unit  "
            .MappingName = "Std"
            .NullText = ""
            .Width = 95
            .Format = "##,###,##0.0000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle4 As New DataGridColoredLine2
        With grdColStyle4
            .HeaderText = "@ ACT of Unit  "
            .MappingName = "Act"
            .NullText = ""
            .Width = 95
            .Format = "##,###,##0.0000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle5 As New DataGridColoredLine2
        With grdColStyle5
            .HeaderText = "Width(mm)/Length(mm)/Num "
            .MappingName = "WLN"
            .NullText = ""
            .Width = 200
            .Format = "#,###,##0.0"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle6 As New DataGridColoredLine2
        With grdColStyle6
            .HeaderText = "@ STD of Formula  "
            .MappingName = "StdM"
            .NullText = ""
            .Width = 120
            .Format = "##,###,##0.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle7 As New DataGridColoredLine2
        With grdColStyle7
            .HeaderText = "@ ACT of Formula  "
            .MappingName = "ActM"
            .NullText = ""
            .Width = 120
            .Format = "##,###,##0.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle8 As New DataGridColoredLine2
        With grdColStyle8
            .HeaderText = "Date"
            .MappingName = "DateUp"
            .Width = 80
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle9 As New DataGridColoredLine2
        With grdColStyle9
            .HeaderText = "Time"
            .MappingName = "TimeUp"
            .Width = 80
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle10 As New DataGridColoredLine2
        With grdColStyle10
            .HeaderText = "@STD of KG"
            .MappingName = "STDKG"
            .Width = 95
            .Format = "##,###,##0.0000"
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle11 As New DataGridColoredLine2
        With grdColStyle11
            .HeaderText = "@ACT of KG"
            .MappingName = "ACTKG"
            .Width = 95
            .Format = "##,###,##0.0000"
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        grdTableStyle1.GridColumnStyles.AddRange _
    (New DataGridColumnStyle() _
    {grdColStyle1, grdColStyle10, grdColStyle11, grdColStyle3, grdColStyle4,
        grdColStyle6, grdColStyle7, grdColStyle8, grdColStyle9, grdColStyle5})

        DataGridCAL.TableStyles.Add(grdTableStyle1)
        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub
    Private Sub Loadsemi()
        Dim sb As New System.Text.StringBuilder()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sb.AppendLine(" SELECT final,MT,Qpu,StdPrice,ActPrice,wln,std,act,dateup,timeup ,std/qpu*1000 stdKG,act/qpu*1000 actKG ")
        sb.AppendLine(" FROM (")
        sb.AppendLine("   SELECT se.Final,MaterialType MT,QPU*num QPU,StdPrice,ActPrice,num WLN")
        sb.AppendLine("   ,StdPrice*num STD,ActPrice*num ACT")
        sb.AppendLine("   ,substring(mp.DateUp,7,2)+'/'+substring(mp.DateUp,5,2)+'/'+substring(mp.DateUp,1,4) dateup ")
        sb.AppendLine("   ,substring(mp.TimeUp,1,2)+':'+substring(mp.TimeUp,3,2) Timeup")
        sb.AppendLine("   FROM TBLsemi se")
        sb.AppendLine("   LEFT OUTER JOIN ( ")
        sb.AppendLine("     SELECT * ")
        sb.AppendLine("     FROM TBLMasterPrice")
        sb.AppendLine("     WHERE Typecode = '05'")
        sb.AppendLine("   ) mp on se.semicode+se.Revision =mp.mastercode+mp.revision")
        sb.AppendLine("   WHERE active = 1 AND materialType IN ('13','14')")
        sb.AppendLine("   UNION ")
        sb.AppendLine("   SELECT se.Final,MaterialType,QPU*length/1000 Qpu,StdPrice,ActPrice,Length WLN")
        sb.AppendLine("   ,Round(StdPrice*(Length/1000),3) STD,Round(ActPrice*(Length/1000),3) ACT")
        sb.AppendLine("   ,substring(mp.DateUp,7,2)+'/'+substring(mp.DateUp,5,2)+'/'+substring(mp.DateUp,1,4) dateup ")
        sb.AppendLine("   ,substring(mp.TimeUp,1,2)+':'+substring(mp.TimeUp,3,2) Timeup")
        sb.AppendLine("   FROM TBLsemi se")
        sb.AppendLine("   LEFT OUTER JOIN ( ")
        sb.AppendLine("     SELECT * ")
        sb.AppendLine("     FROM TBLMasterPrice")
        sb.AppendLine("     WHERE Typecode = '05'")
        sb.AppendLine("   ) mp on se.semicode+se.Revision =mp.mastercode+mp.revision")
        sb.AppendLine("   WHERE active = 1 AND materialType NOT IN ('13','14','10')")
        sb.AppendLine(" ) xx")
        sb.AppendLine(" UNION ")
        sb.AppendLine(" SELECT final,materialType,Qpu,StdPrice,ActPrice,wln,std,act,dateup,timeup ,std/qpu*1000 stdKG,act/qpu*1000 actKG ")
        sb.AppendLine(" FROM (")
        sb.AppendLine("   SELECT se.Final,MaterialType,QPU*Length/1000 QPU,StdPrice,ActPrice,Length WLN")
        sb.AppendLine("   ,StdPrice*(QPU/1000)*(Length/1000) STD,ActPrice*Qpu/1000*Length/1000 ACT")
        sb.AppendLine("   ,substring(se.DateUp,7,2)+'/'+substring(se.DateUp,5,2)+'/'+substring(se.DateUp,1,4) dateup ")
        sb.AppendLine("   ,'00:00'Timeup")
        sb.AppendLine("   FROM TBLsemi se")
        sb.AppendLine("   LEFT OUTER JOIN ( ")
        sb.AppendLine("     SELECT * ")
        sb.AppendLine("     FROM TBLMasterPrice")
        sb.AppendLine("     WHERE Typecode = '05'")
        sb.AppendLine("   ) mp on se.semicode+se.Revision =mp.mastercode+mp.revision")
        sb.AppendLine("   WHERE active = 1 AND materialType IN ('10')")
        sb.AppendLine(" ) xx")
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
        DT.TableName = TBL_Cal
        GrdDV = DT.DefaultView
        GrdDV.AllowNew = False
        GrdDV.AllowDelete = False
        '************************************
        DataGridCAL.DataSource = GrdDV
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

        With DataGridCAL
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
            .MappingName = TBL_Cal
            .PreferredColumnWidth = 125
            .PreferredRowHeight = 15
        End With

        Dim grdColStyle1 As New DataGridColoredLine2
        With grdColStyle1
            .HeaderText = "Code (Material)"
            .MappingName = "Final"
            .Width = 150
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle3 As New DataGridColoredLine2
        With grdColStyle3
            .HeaderText = "@ STD of Unit  "
            .MappingName = "StdPrice"
            .NullText = ""
            .Width = 85
            .Format = "##,###,##0.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle4 As New DataGridColoredLine2
        With grdColStyle4
            .HeaderText = "@ ACT of Unit "
            .MappingName = "ActPrice"
            .NullText = ""
            .Width = 85
            .Format = "##,###,##0.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle5 As New DataGridColoredLine2
        With grdColStyle5
            .HeaderText = "Width(mm)/Length(mm)/Num "
            .MappingName = "WLN"
            .NullText = ""
            .Width = 200
            .Format = "#####0.0"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle6 As New DataGridColoredLine2
        With grdColStyle6
            .HeaderText = "@ STD by Material Unit  "
            .MappingName = "Std"
            .NullText = ""
            .Width = 150
            .Format = "##,###,##0.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle7 As New DataGridColoredLine2
        With grdColStyle7
            .HeaderText = "@ ACT by Material Unit  "
            .MappingName = "Act"
            .NullText = ""
            .Width = 150
            .Format = "##,###,##0.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle8 As New DataGridColoredLine2
        With grdColStyle8
            .HeaderText = "Date"
            .MappingName = "DateUp"
            .Width = 80
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle9 As New DataGridColoredLine2
        With grdColStyle9
            .HeaderText = "Time"
            .MappingName = "TimeUp"
            .Width = 80
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle10 As New DataGridColoredLine2
        With grdColStyle10
            .HeaderText = "@ STD of KG"
            .MappingName = "StdKG"
            .NullText = ""
            .Width = 85
            .Format = "##,###,##0.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle11 As New DataGridColoredLine2
        With grdColStyle11
            .HeaderText = "@ ACT of KG  "
            .MappingName = "ActKG"
            .NullText = ""
            .Width = 85
            .Format = "##,###,##0.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        grdTableStyle1.GridColumnStyles.AddRange _
    (New DataGridColumnStyle() _
    {grdColStyle1, grdColStyle10, grdColStyle11, grdColStyle3, grdColStyle4,
        grdColStyle6, grdColStyle7, grdColStyle8, grdColStyle9, grdColStyle5})

        DataGridCAL.TableStyles.Add(grdTableStyle1)
        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub
    Private Sub LoadTire()
        Dim sb As New System.Text.StringBuilder()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sb.AppendLine(" SELECT MasterCode,Revision,Tiresize,Typecode")
        sb.AppendLine(" ,(StdPrice/Qty)*1000 STDK,(ActPrice/Qty)*1000 ACTK")
        sb.AppendLine(" ,StdPrice,ActPrice,Qty,bb.dateup,bb.Timeup ")
        sb.AppendLine(" FROM (")
        sb.AppendLine("   SELECT * ")
        sb.AppendLine("   FROM TBLGTHdr ")
        sb.AppendLine("   WHERE active = 1 ")
        sb.AppendLine(" ) aa")
        sb.AppendLine(" LEFT OUTER JOIN ( ")
        sb.AppendLine("   SELECT Mastercode,Revision, Typecode,StdPrice,ActPrice")
        sb.AppendLine("   ,substring(DateUp,7,2)+'/'+substring(DateUp,5,2)+'/'+substring(DateUp,1,4) dateup ")
        sb.AppendLine("   ,substring(TimeUp,1,2)+':'+substring(TimeUp,3,2) Timeup")
        sb.AppendLine("   FROM TBLMasterPrice ")
        sb.AppendLine("   WHERE Typecode ='06'")
        sb.AppendLine(" ) bb on aa.Tirecode+Rev = bb.Mastercode+bb.Revision")
        sb.AppendLine(" ORDER BY Mastercode    ")
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
        DT.TableName = TBL_Cal
        GrdDV = DT.DefaultView
        GrdDV.AllowNew = False
        GrdDV.AllowDelete = False
        '************************************
        DataGridCAL.DataSource = GrdDV
        '************************************
        'Dim i As Integer
        'Dim c34 As String = Chr(34)
        'For i = 0 To DT.Columns.Count - 1
        '    Dim col As String = DT.Columns(i).ColumnName
        '    Dim coltype As String = DT.Columns(i).DataType.FullName
        '    coltype = coltype.Replace("System.", "")
        '    coltype = coltype.Replace("Int32", "integer")
        '    coltype = coltype.Replace("Int16", "integer")
        '    coltype = coltype.Replace("String", "string")
        '    coltype = coltype.Replace("Decimal", "decimal")
        '    Debug.WriteLine("<xs:element name=" & c34 & col.Trim & c34 & "  type= " & c34 & "xs:" & coltype & c34 & " minOccurs=" & c34 & "0" & c34 & "/>")
        'Next
        ResetTableStyle()

        With DataGridCAL
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
            .MappingName = TBL_Cal
            .PreferredColumnWidth = 125
            .PreferredRowHeight = 15
        End With
        Dim grdColStyle1 As New DataGridColoredLine2
        With grdColStyle1
            .HeaderText = "Code"
            .MappingName = "MasterCode"
            .Width = 75
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle1_1 As New DataGridColoredLine2
        With grdColStyle1_1
            .HeaderText = "Revision"
            .MappingName = "Revision"
            .Width = 75
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle1_2 As New DataGridColoredLine2
        With grdColStyle1_2
            .HeaderText = "Tiresize"
            .MappingName = "Tiresize"
            .Width = 95
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle2_0 As New DataGridColoredLine2
        With grdColStyle2_0
            .HeaderText = "@ STD KG"
            .MappingName = "STDk"
            .Width = 100
            .Format = "##,###,##0.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle2_1 As New DataGridColoredLine2
        With grdColStyle2_1
            .HeaderText = "@ ACT KG"
            .MappingName = "ACTk"
            .Width = 100
            .Format = "##,###,##0.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle3_0 As New DataGridColoredLine2
        With grdColStyle3_0
            .HeaderText = "@ STD Unit"
            .MappingName = "STDPrice"
            .Width = 100
            .Format = "##,###,##0.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle3_1 As New DataGridColoredLine2
        With grdColStyle3_1
            .HeaderText = "@ ACT Unit"
            .MappingName = "ActPrice"
            .Width = 100
            .Format = "##,###,##0.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle4_0 As New DataGridColoredLine2
        With grdColStyle4_0
            .HeaderText = "Weight(g.)"
            .MappingName = "Qty"
            .Width = 85
            .Format = "##,###,##0.00"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle4 As New DataGridColoredLine2
        With grdColStyle4
            .HeaderText = "Date"
            .MappingName = "dateup"
            .Width = 85
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle5 As New DataGridColoredLine2
        With grdColStyle5
            .HeaderText = "Time"
            .MappingName = "TimeUp"
            .Width = 85
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        grdTableStyle1.GridColumnStyles.AddRange _
    (New DataGridColumnStyle() _
    {grdColStyle1, grdColStyle1_1, grdColStyle1_2, grdColStyle2_0,
    grdColStyle2_1, grdColStyle3_0, grdColStyle3_1, grdColStyle4_0, grdColStyle4, grdColStyle5})

        DataGridCAL.TableStyles.Add(grdTableStyle1)
        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub

    Private Sub ResetTableStyle()
        ' Clear out the existing TableStyles and result default formatting.
        With DataGridCAL
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
    Sub LoadGroup()
        Dim dtGroup As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = " SELECT  distinct Finalcompound Final"
        StrSQL &= " FROM  TblCompound"
        ' StrSQL &= " where Active = 1 "
        ' StrSQL &= " order by finalcompound"
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtGroup = New DataTable
            DA.Fill(dtGroup)
        Catch
        Finally
        End Try
        dtGroup.TableName = TBL_GP
        GrdDVGP = dtGroup.DefaultView
        '************************************
        ComboBoxComp.DisplayMember = "Final"
        ComboBoxComp.ValueMember = "Final"
        ComboBoxComp.DataSource = dtGroup
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadMaterialType()
        Dim dt As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = " SELECT   Materialcode,MaterialName"
        StrSQL &= " FROM  TBLTypeMaterial where Descname like '" & txtname.Trim() & "'"
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dt = New DataTable
            DA.Fill(dt)
        Catch
        Finally
        End Try
        dt.TableName = TBL_GP
        GrdDVGP = dt.DefaultView
        '************************************
        cmbType.DisplayMember = "MaterialName"
        cmbType.ValueMember = "Materialcode"
        cmbType.DataSource = dt
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
#End Region

#Region "Control Event"
    Private Sub ButtonClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonClose.Click
        Me.Close()
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

#Region "Combobox"
    Private Sub CheckCompound_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckCompound.CheckedChanged
        SelectCompound()
    End Sub

    Private Sub CheckCompGroup_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckCompGroup.CheckedChanged
        SelectCompound()
    End Sub

    Sub SelectCompound()
        If CheckCompound.Checked = True And CheckCompGroup.Checked = True Then
            'If check Final and Group
            GrdDV.RowFilter = " Active = 1 and Finalcompound = '" & ComboBoxComp.Text.Trim & "'"
            DataGridCAL.DataSource = GrdDV
        ElseIf CheckCompound.Checked = True And CheckCompGroup.Checked = False Then
            'If check Final
            GrdDV.RowFilter = " Active = 1"
            DataGridCAL.DataSource = GrdDV
        ElseIf CheckCompound.Checked = False And CheckCompGroup.Checked = True Then
            'If check Group
            GrdDV.RowFilter = " Finalcompound = '" & ComboBoxComp.Text.Trim & "'"
            DataGridCAL.DataSource = GrdDV
        Else
            'If not check
            GrdDV.RowFilter = " "
            DataGridCAL.DataSource = GrdDV
        End If
    End Sub

    Private Sub ComboBoxComp_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxComp.SelectedIndexChanged
        SelectCompound()
    End Sub

#End Region


    Private Sub CheckType_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckType.CheckedChanged
        SelectSemi()
    End Sub

    Private Sub cmbType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbType.SelectedIndexChanged
        SelectSemi()
    End Sub

    ''' <summary>
    ''' For Material Type Semi or PreSemi
    ''' </summary>
    Sub SelectSemi()
        If CheckType.Checked = True Then
            GrdDV.RowFilter = " MT  = '" & cmbType.SelectedValue & "'"
            DataGridCAL.DataSource = GrdDV
        Else
            GrdDV.RowFilter = " "
            DataGridCAL.DataSource = GrdDV
        End If
    End Sub

    Private Sub SetTotal()
        'Set total
        'Format: Form Text - xxx item(s)
        Dim frmTitle As String() = Me.Text.Split(New Char() {"-"c})
        Me.Text = frmTitle(0) & "- " & GrdDV.Count & " item(s)"
    End Sub

    Private Sub CmdView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdView.Click
        If GroupCompound.Visible = True Then
            'Compound Type
            SelectCompound()
        End If

        If GType.Visible = True Then
            'PreSemi or Semi Type
            SelectSemi()
        End If

        If GType.Visible = True And GroupCompound.Visible = True Then
            GrdDV.RowFilter &= " AND dateUP LIKE '" & DateTime.Text.Trim() & "'"
            DataGridCAL.DataSource = GrdDV
        ElseIf GType.Visible = False And GroupCompound.Visible = False Then
            'Material Type R/M, Pigment and Green Tire
            GrdDV.RowFilter = " dateUP LIKE '" & DateTime.Text.Trim() & "'"
            DataGridCAL.DataSource = GrdDV
        ElseIf GType.Visible = False And GroupCompound.Visible = True Then
            'Material Type Compound
            If Not CheckCompound.Checked And Not CheckCompGroup.Checked Then
                'If not check
                GrdDV.RowFilter &= " dateUP LIKE '" & DateTime.Text.Trim() & "'"
            Else
                'If check Final or Group
                GrdDV.RowFilter &= " AND dateUP LIKE '" & DateTime.Text.Trim() & "'"
            End If

            DataGridCAL.DataSource = GrdDV
        ElseIf GType.Visible = True And GroupCompound.Visible = False Then
            'Material Type PreSemi and Semi
            GrdDV.RowFilter &= " AND dateUP LIKE '" & DateTime.Text.Trim() & "'"
            DataGridCAL.DataSource = GrdDV
        Else
            GrdDV.RowFilter &= " "
            DataGridCAL.DataSource = GrdDV
        End If
    End Sub
End Class
