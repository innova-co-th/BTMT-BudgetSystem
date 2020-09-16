#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Inventory_Record.Common
#End Region

Public Class FrmAddCompound

#Region "Declare"
    Inherits System.Windows.Forms.Form
    Public Shared GrdDV As New DataView
    Protected Const TBL_RM As String = "TBL_RM"
    Public Shared GrdDVComp As New DataView
    Protected Const TBL_Comp As String = "TBL_Comp"
    Protected DefaultGridBorderStyle As BorderStyle
    Public Shared cm As CurrencyManager
    Dim iTotal As Double
    Friend TCompound, TCode, TRev, TStep As String
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtName As System.Windows.Forms.TextBox
    Friend WithEvents CmdView As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtCode As System.Windows.Forms.TextBox
    Friend WithEvents CmdClear As System.Windows.Forms.Button
    Friend WithEvents lblError As System.Windows.Forms.Label
    Friend WithEvents CheckAll As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxPigment As System.Windows.Forms.CheckBox
    Friend WithEvents ComboBoxPigment As System.Windows.Forms.ComboBox
    Friend WithEvents TxtRev As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtCompound As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblerror2 As System.Windows.Forms.Label
    Friend WithEvents CheckBoxFinalCompound As System.Windows.Forms.CheckBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtStep As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtRmCode As System.Windows.Forms.TextBox
    Friend WithEvents CHKAll As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmAddCompound))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.DataGridRM = New System.Windows.Forms.DataGrid()
        Me.CmdSave = New System.Windows.Forms.Button()
        Me.CmdClose = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TxtName = New System.Windows.Forms.TextBox()
        Me.CmdView = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TxtCode = New System.Windows.Forms.TextBox()
        Me.CmdClear = New System.Windows.Forms.Button()
        Me.lblError = New System.Windows.Forms.Label()
        Me.CheckAll = New System.Windows.Forms.CheckBox()
        Me.CheckBoxPigment = New System.Windows.Forms.CheckBox()
        Me.ComboBoxPigment = New System.Windows.Forms.ComboBox()
        Me.TxtRev = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TxtCompound = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblerror2 = New System.Windows.Forms.Label()
        Me.CheckBoxFinalCompound = New System.Windows.Forms.CheckBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtStep = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TxtRmCode = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.CHKAll = New System.Windows.Forms.CheckBox()
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
        Me.GroupBox1.Location = New System.Drawing.Point(10, 102)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(912, 577)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        '
        'DataGridRM
        '
        Me.DataGridRM.DataMember = ""
        Me.DataGridRM.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridRM.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridRM.Location = New System.Drawing.Point(3, 18)
        Me.DataGridRM.Name = "DataGridRM"
        Me.DataGridRM.Size = New System.Drawing.Size(906, 556)
        Me.DataGridRM.TabIndex = 0
        '
        'CmdSave
        '
        Me.CmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdSave.Image = CType(resources.GetObject("CmdSave.Image"), System.Drawing.Image)
        Me.CmdSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdSave.Location = New System.Drawing.Point(728, 681)
        Me.CmdSave.Name = "CmdSave"
        Me.CmdSave.Size = New System.Drawing.Size(96, 65)
        Me.CmdSave.TabIndex = 10
        Me.CmdSave.Text = "Save"
        Me.CmdSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdClose
        '
        Me.CmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdClose.Image = CType(resources.GetObject("CmdClose.Image"), System.Drawing.Image)
        Me.CmdClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdClose.Location = New System.Drawing.Point(824, 681)
        Me.CmdClose.Name = "CmdClose"
        Me.CmdClose.Size = New System.Drawing.Size(96, 65)
        Me.CmdClose.TabIndex = 11
        Me.CmdClose.Text = "Close"
        Me.CmdClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.Location = New System.Drawing.Point(622, 46)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(77, 18)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "R/M Name "
        '
        'TxtName
        '
        Me.TxtName.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtName.Location = New System.Drawing.Point(708, 43)
        Me.TxtName.Name = "TxtName"
        Me.TxtName.Size = New System.Drawing.Size(116, 22)
        Me.TxtName.TabIndex = 3
        '
        'CmdView
        '
        Me.CmdView.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdView.Image = CType(resources.GetObject("CmdView.Image"), System.Drawing.Image)
        Me.CmdView.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdView.Location = New System.Drawing.Point(833, 38)
        Me.CmdView.Name = "CmdView"
        Me.CmdView.Size = New System.Drawing.Size(87, 65)
        Me.CmdView.TabIndex = 5
        Me.CmdView.Text = "View"
        Me.CmdView.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(10, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(76, 18)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Compound "
        '
        'TxtCode
        '
        Me.TxtCode.Location = New System.Drawing.Point(96, 9)
        Me.TxtCode.Name = "TxtCode"
        Me.TxtCode.Size = New System.Drawing.Size(120, 22)
        Me.TxtCode.TabIndex = 0
        '
        'CmdClear
        '
        Me.CmdClear.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CmdClear.Image = CType(resources.GetObject("CmdClear.Image"), System.Drawing.Image)
        Me.CmdClear.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdClear.Location = New System.Drawing.Point(10, 681)
        Me.CmdClear.Name = "CmdClear"
        Me.CmdClear.Size = New System.Drawing.Size(96, 65)
        Me.CmdClear.TabIndex = 12
        Me.CmdClear.Text = "Clear"
        Me.CmdClear.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'lblError
        '
        Me.lblError.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblError.ForeColor = System.Drawing.Color.Red
        Me.lblError.Location = New System.Drawing.Point(221, 16)
        Me.lblError.Name = "lblError"
        Me.lblError.Size = New System.Drawing.Size(29, 9)
        Me.lblError.TabIndex = 8
        Me.lblError.Text = "***"
        Me.lblError.Visible = False
        '
        'CheckAll
        '
        Me.CheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CheckAll.Location = New System.Drawing.Point(125, 728)
        Me.CheckAll.Name = "CheckAll"
        Me.CheckAll.Size = New System.Drawing.Size(134, 18)
        Me.CheckAll.TabIndex = 13
        Me.CheckAll.Text = "Show All"
        Me.CheckAll.Visible = False
        '
        'CheckBoxPigment
        '
        Me.CheckBoxPigment.Location = New System.Drawing.Point(19, 74)
        Me.CheckBoxPigment.Name = "CheckBoxPigment"
        Me.CheckBoxPigment.Size = New System.Drawing.Size(87, 28)
        Me.CheckBoxPigment.TabIndex = 6
        Me.CheckBoxPigment.Text = "Pigment "
        Me.CheckBoxPigment.Visible = False
        '
        'ComboBoxPigment
        '
        Me.ComboBoxPigment.Enabled = False
        Me.ComboBoxPigment.Location = New System.Drawing.Point(115, 74)
        Me.ComboBoxPigment.Name = "ComboBoxPigment"
        Me.ComboBoxPigment.Size = New System.Drawing.Size(144, 24)
        Me.ComboBoxPigment.TabIndex = 7
        Me.ComboBoxPigment.Text = "Select"
        Me.ComboBoxPigment.Visible = False
        '
        'TxtRev
        '
        Me.TxtRev.Location = New System.Drawing.Point(96, 37)
        Me.TxtRev.Name = "TxtRev"
        Me.TxtRev.Size = New System.Drawing.Size(48, 22)
        Me.TxtRev.TabIndex = 13
        Me.TxtRev.Text = "001"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(10, 39)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(76, 19)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = " Rev."
        '
        'TxtCompound
        '
        Me.TxtCompound.Location = New System.Drawing.Point(403, 9)
        Me.TxtCompound.Name = "TxtCompound"
        Me.TxtCompound.Size = New System.Drawing.Size(96, 22)
        Me.TxtCompound.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(278, 12)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(125, 18)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Group  Compound"
        '
        'lblerror2
        '
        Me.lblerror2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblerror2.ForeColor = System.Drawing.Color.Red
        Me.lblerror2.Location = New System.Drawing.Point(509, 16)
        Me.lblerror2.Name = "lblerror2"
        Me.lblerror2.Size = New System.Drawing.Size(29, 9)
        Me.lblerror2.TabIndex = 15
        Me.lblerror2.Text = "***"
        Me.lblerror2.Visible = False
        '
        'CheckBoxFinalCompound
        '
        Me.CheckBoxFinalCompound.Location = New System.Drawing.Point(394, 74)
        Me.CheckBoxFinalCompound.Name = "CheckBoxFinalCompound"
        Me.CheckBoxFinalCompound.Size = New System.Drawing.Size(134, 28)
        Me.CheckBoxFinalCompound.TabIndex = 8
        Me.CheckBoxFinalCompound.Text = "Final Compound"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(326, 74)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(48, 22)
        Me.TextBox1.TabIndex = 18
        Me.TextBox1.Visible = False
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(269, 76)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(48, 19)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = " Rev."
        Me.Label5.Visible = False
        '
        'txtStep
        '
        Me.txtStep.Location = New System.Drawing.Point(403, 37)
        Me.txtStep.Name = "txtStep"
        Me.txtStep.Size = New System.Drawing.Size(48, 22)
        Me.txtStep.TabIndex = 2
        Me.txtStep.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(346, 39)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(48, 19)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Stage"
        '
        'TxtRmCode
        '
        Me.TxtRmCode.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRmCode.Location = New System.Drawing.Point(708, 80)
        Me.TxtRmCode.Name = "TxtRmCode"
        Me.TxtRmCode.Size = New System.Drawing.Size(116, 22)
        Me.TxtRmCode.TabIndex = 4
        '
        'Label7
        '
        Me.Label7.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label7.Location = New System.Drawing.Point(622, 83)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(77, 18)
        Me.Label7.TabIndex = 22
        Me.Label7.Text = "R/M Code"
        '
        'CHKAll
        '
        Me.CHKAll.Location = New System.Drawing.Point(624, 695)
        Me.CHKAll.Name = "CHKAll"
        Me.CHKAll.Size = New System.Drawing.Size(106, 18)
        Me.CHKAll.TabIndex = 23
        Me.CHKAll.Text = "Add Check"
        '
        'FrmAddCompound
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(932, 753)
        Me.Controls.Add(Me.CHKAll)
        Me.Controls.Add(Me.TxtRmCode)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txtStep)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.CheckBoxFinalCompound)
        Me.Controls.Add(Me.lblerror2)
        Me.Controls.Add(Me.TxtCompound)
        Me.Controls.Add(Me.TxtRev)
        Me.Controls.Add(Me.TxtCode)
        Me.Controls.Add(Me.TxtName)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.ComboBoxPigment)
        Me.Controls.Add(Me.CheckBoxPigment)
        Me.Controls.Add(Me.CheckAll)
        Me.Controls.Add(Me.lblError)
        Me.Controls.Add(Me.CmdClear)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.CmdView)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CmdClose)
        Me.Controls.Add(Me.CmdSave)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmAddCompound"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "COMPOUND (MIXING)  "
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataGridRM, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

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

#Region "Function_Load"
    Private Sub LoadRM()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        If CmdSave.Text = "Save" Then
            StrSQL = " select * from "
            StrSQL &= " (select Typecode,'' mastercode,Code,Revision RMRev,''descName,Seq,FinalCompound,0.00 RMQty,0.00 Qty,'KG' Unitcode"
            StrSQL &= "  from tblGroup a"
            StrSQL &= "  left outer join "
            StrSQL &= " (   SELECT  seq,Finalcompound,Compcode,Revision "
            StrSQL &= "  FROM         TBLcompound"
            StrSQL &= " ) b"
            StrSQL &= " on a.code = b.compcode"
            StrSQL &= "  where Code in "
            StrSQL &= " ("
            StrSQL &= "  SELECT  Compcode "
            StrSQL &= "  FROM         TBLcompound"
            StrSQL &= " )"
            StrSQL &= " union "
            StrSQL &= " select Typecode,isnull(mastercode,'')Mastercode,Code,Revision,descName,Seq,FinalCompound"
            StrSQL &= " ,isnull(Qty,0.00) RMQTY ,isnull(Qty,0.00) Qty,'KG' Unitcode from "
            StrSQL &= " (select Typecode,Code,'' Revision,descName,1 Seq,'' FinalCompound"
            StrSQL &= "  from "
            StrSQL &= "  TBLGroup g"
            StrSQL &= "  left outer join "
            StrSQL &= "  ( select  RMCode,descName,0.00 QTY,'KG' Unit  "
            StrSQL &= "  FROM   TBLRM ) RM"
            StrSQL &= "  on g.code = rm.rmcode  "
            StrSQL &= " where Typecode ='01')aa"
            StrSQL &= " left outer join "
            StrSQL &= " (  SELECT    mastercode,RMCODE,QTY"
            StrSQL &= "  FROM         TBLMaster"
            If CheckBoxPigment.Checked = True Then
                StrSQL &= "   where MasterCode = '" & ComboBoxPigment.Text.Trim & "' "
            Else
                StrSQL &= "   where MasterCode = '' "
            End If
            StrSQL &= " ) bb  on aa.code = bb.rmcode"
            StrSQL &= " )xxx"
            StrSQL &= "  where FinalCompound ='" & TxtCompound.Text.Trim & "' or seq =1"
            StrSQL &= " order by typecode desc,Code"

        ElseIf CmdSave.Text = "Edit" Then
            'StrSQL = "  select Typecode,code,RMREV,isnull(descName,'')as descName"
            'StrSQL &= "   ,isnull(RMQty,'') as RMQty,isnull(BQty,'') as Qty, 'KG' as unitcode from "
            'StrSQL &= "   (select  b.RMCode,RMREV,isnull(descName,'') as descName,isnull(b.Qty,0.000) as RMQty"
            'StrSQL &= "   ,isnull(b.Qty,0.000) as BQTY  "
            'StrSQL &= "    FROM   (       "
            'StrSQL &= "   SELECT    RMCODE,isnull(RMRevision,'') RMREV,Qty"
            'StrSQL &= "   FROM         TBLMaster"
            'StrSQL &= "   where RMCode in "
            'StrSQL &= "  ("
            'StrSQL &= "  SELECT    RMCODE"
            'StrSQL &= "  FROM         TBLMaster"
            'StrSQL &= "   where MasterCode in ('" & TxtCode.Text.Trim & "')"
            'StrSQL &= "  ) and MasterCode = '" & TxtCode.Text.Trim & "')b"
            'StrSQL &= "   left outer join "
            'StrSQL &= "   TBLRM rm  "
            'StrSQL &= "   on rm.RMCode = b.RMCode"
            'StrSQL &= "   )bb"
            'StrSQL &= "   left outer join 	"
            'StrSQL &= "  (select * from tblGroup"
            'StrSQL &= "   )aa 	"
            'StrSQL &= "   on aa.code = bb.rmcode"
            'StrSQL &= "   order by typecode desc,aa.Code"
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
            .HeaderText = "Code"
            .MappingName = "Code"
            .Width = 100
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle1_1 As New DataGridColoredLine2
        With grdColStyle1_1
            .HeaderText = "Rev."
            .MappingName = "RMRev"
            .Width = 85
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle2 As New DataGridColoredLine2
        With grdColStyle2
            .HeaderText = "Name"
            .MappingName = "DescName"
            .Width = 110
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle3 As New DataGridColoredLine2
        With grdColStyle3
            .HeaderText = "Qty"
            .MappingName = "RMQty"
            .Width = 110
            .Format = "###,##0.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        cm = CType(Me.BindingContext(DataGridRM.DataSource, DataGridRM.DataMember), CurrencyManager)
        Dim c As CheckRowHeader
        c = AddressOf CheckRowHeader

        Dim grdColStyle5 As New DataGridQtyBox(c)
        With grdColStyle5
            .HeaderText = "RHC"
            .MappingName = "RHC"
            .Format = "###,##0.000"
            .Width = 110
            .Alignment = HorizontalAlignment.Center
            .NullText = ""
        End With
        Dim grdColStyle6 As New DataGridQtyBox(c)
        With grdColStyle6
            .HeaderText = "Qty(KG)"
            .MappingName = "Qty"
            .Format = "###,##0.000"
            .Width = 110
            .Alignment = HorizontalAlignment.Center
            .NullText = ""
        End With

        Dim grdColStyle7 As New DataGridColoredLine2
        With grdColStyle7
            .HeaderText = "Unit"
            .MappingName = "UnitCode"
            .Width = 80
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With

        grdTableStyle1.GridColumnStyles.AddRange _
(New DataGridColumnStyle() _
{grdColStyle1, grdColStyle1_1, grdColStyle2, grdColStyle5, grdColStyle6, grdColStyle7})

        DataGridRM.TableStyles.Add(grdTableStyle1)
        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub


    Public Shared Function CheckRowHeader(ByVal row As Integer) As Boolean
        Dim c As Boolean = False
        'Debug.WriteLine("st seq : " + CStr(GrdItemDv.Item(row).Item("st_seq")) + "   row : " + CStr(row))
        Try
            If GrdDV.Item(row).Item("item_no").ToString.Trim = "" Then
                c = True
            Else
                c = False
            End If
        Catch ex As Exception
            c = False
        End Try

        Return c
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

#Region "COMBOBOX"
    Sub LoadPigment()
        Dim dtComp As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        StrSQL = "SELECT  PigmentCode,Revision  from TblPigment"

        Dim C1 As New SQLData("ACCINV")
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
        ComboBoxPigment.DisplayMember = "PigmentCode"
        ComboBoxPigment.ValueMember = "Revision"
        ComboBoxPigment.DataSource = dtComp
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
#End Region

    Private Sub FrmAddCompound_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If CmdSave.Text = "Edit" Then
            CheckBoxPigment.Enabled = False
            TextBox1.Text = ComboBoxPigment.SelectedValue
        Else
            CheckAll.Visible = False
        End If
        LoadPigment()
    End Sub

    Private Sub CmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClose.Click
        Me.Close()
    End Sub

    Private Sub CmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSave.Click
        If ChkAddData() Then
        Else
            MsgBox("It's Duplicate.", MsgBoxStyle.OKOnly, "ADD COMPOUND")
            Exit Sub
        End If
        If TxtCode.Text.Trim = "" Then
            TxtCode.Focus()
            lblError.Visible = True
            Exit Sub
        Else
            TxtCode.Text = TxtCode.Text.ToUpper
            lblError.Visible = False
        End If

        If TxtCompound.Text.Trim = "" Then
            TxtCompound.Focus()
            lblerror2.Visible = True
            Exit Sub
        Else
            TxtCompound.Text = TxtCompound.Text.ToUpper
            lblerror2.Visible = False
        End If

        If txtStep.Text.Trim = "" Then
            txtStep.Focus()
            Exit Sub
        Else
        End If

        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        Dim aDr() As DataRow
        GrdDV.RowFilter = " Qty <> 0.000"
        aDr = GrdDV.Table.Select(GrdDV.RowFilter)
        iTotal = 0
        CheckAll.Checked = False
        Dim dr As DataRow
        For Each dr In aDr
            With dr
                If IIf(.Item("Code") Is System.DBNull.Value, "", .Item("Code")) <> "" Then
                    iTotal = iTotal + .Item("Qty")
                End If
            End With
        Next

        msg = "Compound Total :" & iTotal & "" ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Compound"   ' Define title.
        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            UpComp()
            RM()
            If CHKAll.Checked Then
                LoadRM()
                LoadPigment()
                TxtCompound.Text = TxtCompound.Text
                txtStep.Text = txtStep.Text + 1
                TxtCode.Text = ""
                TxtCode.Focus()
                TxtRev.Text = "001"
            Else
                Me.Close()
            End If
        Else
            Exit Sub
        End If
    End Sub

    Private Sub DataGridRM_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridRM.CurrentCellChanged
        oldrow = DataGridRM.CurrentCell.RowNumber
    End Sub

    Private Sub CmdView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdView.Click
        GrdDV.RowFilter = "  descname like'%" & TxtName.Text.Trim & "%'" _
                            & "  and Code like'%" & TxtRmCode.Text.Trim & "%'"
        DataGridRM.DataSource = GrdDV
    End Sub

#Region "RM"
    Private Function ChkAddData() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String
        Try
            strSQL = " Select count(*) from TBLMASTER "
            strSQL += " Where MasterCode = " & PrepareStr(TxtCode.Text.Trim)
            strSQL += " and Revision = " & PrepareStr(TxtRev.Text.Trim)
            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar
            If i = 0 Then
                ChkAddData = True
            End If
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
    Private Function ChkData() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String
        Try
            strSQL = " Select MasterCode from TBLMASTER "
            strSQL += " Where MasterCode = " & PrepareStr(ComboBoxPigment.Text.Trim)
            strSQL += " and Revision = " & PrepareStr(TxtRev.Text.Trim)
            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar
            If i <> 0 Then
                ChkData = True
            End If
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
    Private Function Chkfinal() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String
        Try
            strSQL = " Select count(*) from TBLcompound "
            strSQL += " Where finalcompound = " & PrepareStr(TxtCompound.Text.Trim)
            strSQL += " and Active = " & PrepareStr(1)
            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar
            If i <> 0 Then
                Chkfinal = True
            End If
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
            strsql += "update  TBLCompound "
            strsql += " set Active = " & PrepareStr(0)
            strsql += " , dateUp = " & PrepareStr(strDate)
            strsql += " Where  finalcompound = " & PrepareStr(TxtCompound.Text.Trim)
            cmd.CommandText = strsql
            cmd.ExecuteNonQuery()
            t1.Commit()
        Catch
            t1.Rollback()
            MsgBox("Rollback data")
        Finally
            cn.Close()
        End Try
    End Sub
    Private Function iNoRev() As Long
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim drSQL As SqlDataReader
        Dim strSQL As String
        Try

            ' Build Insert statement to insert 
            strSQL = "  SELECT   top 1 Revision "
            strSQL &= "  FROM   TBLCompound"
            strSQL &= " Where CompCode  = '" & TxtCode.Text.Trim & "'"
            strSQL &= "  order by Revision desc"

            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            cmSQL.ExecuteScalar()
            drSQL = cmSQL.ExecuteReader()
            If drSQL.HasRows Then
                If drSQL.Read() Then
                    iNoRev = CInt(drSQL.Item("Revision").ToString())
                End If
            End If

            ' Close and Clean up objects
            drSQL.Close()
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
    Private Function iNoSeq(ByVal Ccode As String, ByVal rev As String) As Long
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim drSQL As SqlDataReader
        Dim strSQL As String
        Try

            ' Build Insert statement to insert 
            strSQL = "    SELECT  seq"
            strSQL &= "  FROM  TBLCompound  "
            strSQL &= "  Where FinalCompound  = '" & Ccode & "' and Revision ='" & rev & "'"
            strSQL &= " order by seq desc"
            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            cmSQL.ExecuteScalar()
            drSQL = cmSQL.ExecuteReader()
            If drSQL.HasRows Then
                If drSQL.Read() Then
                    iNoSeq = CInt(drSQL.Item("seq").ToString())
                End If
            End If

            ' Close and Clean up objects
            drSQL.Close()
            cnSQL.Close()
            cmSQL.Dispose()
            cnSQL.Dispose()
            '--------------------------------------------------------------------------------------
        Catch Exp As SqlException
            '  MsgBox(Exp.Message, MsgBoxStyle.Critical, "SQL Error")
        Catch Exp As Exception
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "General Error")
        End Try
    End Function
    Private Function iqty(ByVal ii As String, ByVal jj As String) As Boolean
        If ii = jj Then
            iqty = False
        Else
            iqty = True
        End If
    End Function
    Sub RM()
        Dim strsql As String = String.Empty
        Dim tNo As Double
        Dim tSeq As Integer
        Dim cn As New SqlConnection(C1.Strcon)
        Dim cmd As New SqlCommand(strsql, cn)
        cn.Open()
        Dim t1 As SqlTransaction = cn.BeginTransaction
        cmd.Transaction = t1
        Dim strDate, SD() As String
        SD = Split(Now.ToShortDateString, "/")
        strDate = SD(2) + SD(1) + SD(0)
        If CmdSave.Text = "Save" Then
            Try
                Dim aDr() As DataRow
                GrdDV.RowFilter = " Qty <> 0.000"
                aDr = GrdDV.Table.Select(GrdDV.RowFilter)
                If UBound(aDr) < 0 Then
                    Exit Sub
                End If
                Dim dr As DataRow
                For Each dr In aDr
                    With dr
                        If IIf(.Item("Code") Is System.DBNull.Value, "", .Item("Code")) <> "" Then
                            Dim RMstr As String
                            RMstr = .Item("MasterCode")
                            If .Item("MasterCode") = "" Then
                                strsql = "Insert TBLMASTER "
                                strsql += " Values(" & PrepareStr(TxtCode.Text.Trim)
                                strsql += "," & PrepareStr(TxtRev.Text.Trim)
                                strsql += "," & PrepareStr(.Item("Code"))
                                strsql += "," & PrepareStr(.Item("RmRev"))
                                strsql += "," & PrepareStr(.Item("Qty"))
                                strsql += "," & PrepareStr(.Item("unitcode"))
                                strsql += "," & PrepareStr(CSng(.Item("Qty") / iTotal * 100))
                                strsql += ")"
                                cmd.CommandText = strsql
                                cmd.ExecuteNonQuery()
                            Else
                                tNo = tNo + .Item("Qty")
                            End If
                        End If
                    End With
                Next

                tSeq = txtStep.Text.Trim
                Try

                    strsql = "Insert  TblGroup "
                    strsql += " values ( '03',"
                    strsql += PrepareStr(TxtCode.Text.Trim) & ")"

                    strsql += "  "
                    strsql += "Insert  TBLCompound "
                    strsql += " Values(" & PrepareStr(tSeq)
                    strsql += "," & PrepareStr(TxtCompound.Text.Trim)
                    strsql += "," & PrepareStr(TxtCode.Text.Trim)
                    strsql += "," & PrepareStr(TxtRev.Text.Trim)
                    strsql += "," & PrepareStr("")
                    strsql += "," & PrepareStr("")
                    strsql += "," & PrepareStr(iTotal)
                    If CheckBoxFinalCompound.Checked Then
                        strsql += "," & PrepareStr(1)
                    Else
                        strsql += "," & PrepareStr(0)
                    End If
                    strsql += "," & PrepareStr(strDate)
                    strsql += ")"

                    If CheckBoxPigment.Checked = True Then
                        strsql += "  "
                        strsql += "Insert TBLMASTER "
                        strsql += " Values(" & PrepareStr(TxtCode.Text.Trim)
                        strsql += "," & PrepareStr(TxtRev.Text.Trim)
                        strsql += "," & PrepareStr(ComboBoxPigment.Text.Trim)
                        strsql += "," & PrepareStr(TextBox1.Text.Trim)
                        strsql += "," & PrepareStr(tNo)
                        strsql += "," & PrepareStr("KG")
                        strsql += "," & PrepareStr(CSng(tNo / iTotal * 100))
                        strsql += ")"
                    Else
                    End If

                    strsql &= ""
                    strsql &= " Insert TBLConvert "
                    strsql &= " Values('03'"
                    strsql &= "," & PrepareStr(TxtCompound.Text.Trim)
                    strsql &= "," & PrepareStr(TxtCode.Text.Trim)
                    strsql &= "," & PrepareStr(TxtRev.Text.Trim)
                    strsql &= "," & PrepareStr("BT")
                    strsql &= "," & PrepareStr("KG")
                    strsql &= "," & PrepareStr(1)
                    strsql &= "," & PrepareStr(iTotal)
                    strsql &= ")"

                    cmd.CommandText = strsql
                    cmd.ExecuteNonQuery()
                Catch
                End Try

                t1.Commit()
                MsgBox("Update Complete.", MsgBoxStyle.Information, "Compound Code")
            Catch
                t1.Rollback()
                MsgBox("Rollback data")
            Finally
                cn.Close()
            End Try
        ElseIf CmdSave.Text = "Edit" Then
            Try
                Dim aDr() As DataRow
                aDr = GrdDV.Table.Select()
                If UBound(aDr) < 0 Then
                    Exit Sub
                End If
                Dim dr As DataRow
                For Each dr In aDr
                    With dr
                        If IIf(.Item("Code") Is System.DBNull.Value, "", .Item("Code")) <> "" Then
                            If .Item("RMQty") = 0.0 And .Item("QTY") = 0 Then
                            Else
                                If iqty(ComboBoxPigment.Text.Trim, .Item("Code")) Then
                                    strsql = "Update TBLMASTER "
                                    strsql += " Set Qty = " & PrepareStr(.Item("Qty"))
                                    strsql += " Where MasterCode = " & PrepareStr(TCode.Trim)
                                    strsql += " and  RMCode = " & PrepareStr(.Item("Code"))
                                    strsql += " and  Revision = " & PrepareStr(TRev.Trim)
                                Else
                                    strsql = "Update TBLMASTER "
                                    strsql += " Set Qty = " & PrepareStr(.Item("Qty"))
                                    strsql += " Where MasterCode = " & PrepareStr(TCode.Trim)
                                    strsql += " and  RMCode = " & PrepareStr(.Item("Code"))
                                    strsql += " and  Revision = " & PrepareStr(TRev.Trim)
                                End If
                                cmd.CommandText = strsql
                                cmd.ExecuteNonQuery()
                            End If
                        End If
                    End With
                Next

                Dim uTotal As Double
                aDr = GrdDV.Table.Select()
                uTotal = 0
                CheckAll.Checked = False
                For Each dr In aDr
                    With dr
                        If IIf(.Item("Code") Is System.DBNull.Value, "", .Item("Code")) <> "" Then
                            uTotal = uTotal + .Item("Qty")
                        End If
                    End With
                Next

                Try

                    strsql = " Update  TblCompound "
                    strsql += " set Qty ='" & uTotal & "'"
                    strsql += ", Seq = " & PrepareStr(txtStep.Text.Trim)
                    If CheckBoxFinalCompound.Checked Then
                        strsql += ", Active = " & PrepareStr(1)
                    Else
                        strsql += ", Active =" & PrepareStr(0)
                    End If
                    strsql += " Where CompCode = " & PrepareStr(TCode.Trim)
                    strsql += " and  Revision = " & PrepareStr(TRev.Trim)

                    strsql += " "
                    strsql += "  Update TBLConvert "
                    strsql += " set  SQty = " & PrepareStr(iTotal)
                    strsql += " where Final = " & PrepareStr(TxtCompound.Text.Trim)
                    strsql += " and  code = " & PrepareStr(TxtCode.Text.Trim)
                    strsql += " and  Rev = " & PrepareStr(TxtRev.Text.Trim)

                    cmd.CommandText = strsql
                    cmd.ExecuteNonQuery()
                Catch
                End Try

                t1.Commit()
                MsgBox("Update Complete.", MsgBoxStyle.Information, "Compound Code")
            Catch
                t1.Rollback()
                MsgBox("Rollback data")
            Finally
                cn.Close()
            End Try

        Else
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

    Private Sub TxtName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtName.KeyPress
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
        CheckBoxPigment.Checked = False
        ComboBoxPigment.Enabled = False
        TxtCode.Text = ""
        TxtName.Text = ""
    End Sub

    Private Sub CheckAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckAll.CheckedChanged
        If CheckAll.Checked = True Then
            GrdDV.RowFilter = " RMQty <> '' "
            DataGridRM.DataSource = GrdDV
        Else
            GrdDV.RowFilter = " RMQty <> 0.000 "
            DataGridRM.DataSource = GrdDV
        End If
    End Sub

    Private Sub CheckBoxPigment_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxPigment.CheckedChanged
        LoadRM()
        If CheckBoxPigment.Checked = True Then
            ComboBoxPigment.Enabled = True
            TextBox1.Text = ComboBoxPigment.SelectedValue
        Else
            ComboBoxPigment.Enabled = False
            GrdDV.RowFilter = "  "
            TextBox1.Text = ""
        End If
    End Sub

    Private Sub ComboBoxPigment_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxPigment.SelectedIndexChanged
        LoadRM()
        If CheckBoxPigment.Checked = True Then
            TextBox1.Text = ComboBoxPigment.SelectedValue
        Else
            TextBox1.Text = ""
        End If
    End Sub

    Private Sub TxtCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCode.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                TxtCode.Text = TxtCode.Text.ToUpper
                'If CmdSave.Text = "Save" Then
                '    '  i = iNoRev() + 1
                '    ' TxtRev.Text = Format(i, "000")
                '    '  txtStep.Text = iNoSeq(TxtCode.Text.Trim, TxtRev.Text.Trim) + 1
                'Else
                'End If
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case Else
        End Select
    End Sub

    Private Sub TxtCompound_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCompound.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                TxtCompound.Text = TxtCompound.Text.ToUpper
                LoadRM()
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case Else
        End Select
    End Sub

    Private Sub txtStep_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtStep.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case Else
                If Not IsNumeric(e.KeyChar) Then
                    e.Handled = True
                Else
                    If Len(sender.text) >= 1 Then
                        If txtStep.SelectionLength = Len(txtStep.Text) Then
                        Else
                            e.Handled = True
                        End If
                    End If
                End If
        End Select
    End Sub


    Private Sub TxtRmCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtRmCode.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case Else
        End Select
    End Sub

   
End Class
