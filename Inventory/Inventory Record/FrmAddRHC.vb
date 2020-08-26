#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Inventory_Record.Common
#End Region

Public Class FrmAddRHC

#Region "Declare"
    Inherits System.Windows.Forms.Form
    Public Shared GrdDV As New DataView
    Protected Const TBL_RM As String = "TBL_RM"
    Public Shared GrdDVComp As New DataView
    Protected Const TBL_Comp As String = "TBL_Comp"
    Protected DefaultGridBorderStyle As BorderStyle
    Public Shared cm As CurrencyManager
    Dim iTotal, qTotal As Double
    Friend TCompound, TCode, TRev, TStep As String
    Friend CBal As Boolean
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
    Friend WithEvents TxtRev As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtCompound As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblerror2 As System.Windows.Forms.Label
    Friend WithEvents txtStep As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents PgBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents ChkCal As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmAddRHC))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.DataGridRM = New System.Windows.Forms.DataGrid
        Me.CmdSave = New System.Windows.Forms.Button
        Me.CmdClose = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtName = New System.Windows.Forms.TextBox
        Me.CmdView = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtCode = New System.Windows.Forms.TextBox
        Me.CmdClear = New System.Windows.Forms.Button
        Me.lblError = New System.Windows.Forms.Label
        Me.CheckAll = New System.Windows.Forms.CheckBox
        Me.TxtRev = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtCompound = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.lblerror2 = New System.Windows.Forms.Label
        Me.txtStep = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.PgBar1 = New System.Windows.Forms.ProgressBar
        Me.ChkCal = New System.Windows.Forms.CheckBox
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
        Me.GroupBox1.Size = New System.Drawing.Size(722, 456)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        '
        'DataGridRM
        '
        Me.DataGridRM.DataMember = ""
        Me.DataGridRM.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridRM.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridRM.Location = New System.Drawing.Point(3, 16)
        Me.DataGridRM.Name = "DataGridRM"
        Me.DataGridRM.Size = New System.Drawing.Size(716, 437)
        Me.DataGridRM.TabIndex = 0
        '
        'CmdSave
        '
        Me.CmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdSave.Image = CType(resources.GetObject("CmdSave.Image"), System.Drawing.Image)
        Me.CmdSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdSave.Location = New System.Drawing.Point(600, 562)
        Me.CmdSave.Name = "CmdSave"
        Me.CmdSave.Size = New System.Drawing.Size(64, 56)
        Me.CmdSave.TabIndex = 5
        Me.CmdSave.Text = "Save"
        Me.CmdSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdClose
        '
        Me.CmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdClose.Image = CType(resources.GetObject("CmdClose.Image"), System.Drawing.Image)
        Me.CmdClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdClose.Location = New System.Drawing.Point(664, 562)
        Me.CmdClose.Name = "CmdClose"
        Me.CmdClose.Size = New System.Drawing.Size(64, 56)
        Me.CmdClose.TabIndex = 6
        Me.CmdClose.Text = "Close"
        Me.CmdClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.Location = New System.Drawing.Point(480, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 16)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "R/M Name "
        '
        'TxtName
        '
        Me.TxtName.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtName.Location = New System.Drawing.Point(552, 8)
        Me.TxtName.Name = "TxtName"
        Me.TxtName.Size = New System.Drawing.Size(120, 20)
        Me.TxtName.TabIndex = 2
        Me.TxtName.Text = ""
        '
        'CmdView
        '
        Me.CmdView.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdView.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.CmdView.Location = New System.Drawing.Point(672, 7)
        Me.CmdView.Name = "CmdView"
        Me.CmdView.Size = New System.Drawing.Size(48, 23)
        Me.CmdView.TabIndex = 3
        Me.CmdView.Text = "View"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 10)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 16)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Compound "
        '
        'TxtCode
        '
        Me.TxtCode.Location = New System.Drawing.Point(80, 8)
        Me.TxtCode.Name = "TxtCode"
        Me.TxtCode.ReadOnly = True
        Me.TxtCode.TabIndex = 0
        Me.TxtCode.Text = ""
        '
        'CmdClear
        '
        Me.CmdClear.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CmdClear.Image = CType(resources.GetObject("CmdClear.Image"), System.Drawing.Image)
        Me.CmdClear.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdClear.Location = New System.Drawing.Point(8, 562)
        Me.CmdClear.Name = "CmdClear"
        Me.CmdClear.Size = New System.Drawing.Size(80, 56)
        Me.CmdClear.TabIndex = 7
        Me.CmdClear.Text = "Clear"
        Me.CmdClear.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'lblError
        '
        Me.lblError.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblError.ForeColor = System.Drawing.Color.Red
        Me.lblError.Location = New System.Drawing.Point(184, 14)
        Me.lblError.Name = "lblError"
        Me.lblError.Size = New System.Drawing.Size(24, 8)
        Me.lblError.TabIndex = 8
        Me.lblError.Text = "***"
        Me.lblError.Visible = False
        '
        'CheckAll
        '
        Me.CheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CheckAll.Location = New System.Drawing.Point(104, 602)
        Me.CheckAll.Name = "CheckAll"
        Me.CheckAll.Size = New System.Drawing.Size(112, 16)
        Me.CheckAll.TabIndex = 8
        Me.CheckAll.Text = "Show All"
        Me.CheckAll.Visible = False
        '
        'TxtRev
        '
        Me.TxtRev.Location = New System.Drawing.Point(80, 32)
        Me.TxtRev.Name = "TxtRev"
        Me.TxtRev.ReadOnly = True
        Me.TxtRev.Size = New System.Drawing.Size(40, 20)
        Me.TxtRev.TabIndex = 13
        Me.TxtRev.Text = ""
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 35)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 14)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = " Rev."
        '
        'TxtCompound
        '
        Me.TxtCompound.Location = New System.Drawing.Point(320, 8)
        Me.TxtCompound.Name = "TxtCompound"
        Me.TxtCompound.ReadOnly = True
        Me.TxtCompound.Size = New System.Drawing.Size(80, 20)
        Me.TxtCompound.TabIndex = 1
        Me.TxtCompound.Text = ""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(232, 10)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 16)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Final Compound"
        '
        'lblerror2
        '
        Me.lblerror2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblerror2.ForeColor = System.Drawing.Color.Red
        Me.lblerror2.Location = New System.Drawing.Point(408, 14)
        Me.lblerror2.Name = "lblerror2"
        Me.lblerror2.Size = New System.Drawing.Size(24, 8)
        Me.lblerror2.TabIndex = 15
        Me.lblerror2.Text = "***"
        Me.lblerror2.Visible = False
        '
        'txtStep
        '
        Me.txtStep.Location = New System.Drawing.Point(192, 32)
        Me.txtStep.Name = "txtStep"
        Me.txtStep.ReadOnly = True
        Me.txtStep.Size = New System.Drawing.Size(40, 20)
        Me.txtStep.TabIndex = 20
        Me.txtStep.Text = ""
        Me.txtStep.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(144, 35)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(40, 14)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Stage"
        '
        'PgBar1
        '
        Me.PgBar1.Location = New System.Drawing.Point(8, 528)
        Me.PgBar1.Name = "PgBar1"
        Me.PgBar1.Size = New System.Drawing.Size(720, 23)
        Me.PgBar1.TabIndex = 21
        Me.PgBar1.Visible = False
        '
        'ChkCal
        '
        Me.ChkCal.Location = New System.Drawing.Point(472, 592)
        Me.ChkCal.Name = "ChkCal"
        Me.ChkCal.Size = New System.Drawing.Size(120, 24)
        Me.ChkCal.TabIndex = 22
        Me.ChkCal.Text = "Calculate Percent"
        '
        'FrmAddRHC
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(738, 624)
        Me.Controls.Add(Me.ChkCal)
        Me.Controls.Add(Me.PgBar1)
        Me.Controls.Add(Me.txtStep)
        Me.Controls.Add(Me.TxtCompound)
        Me.Controls.Add(Me.TxtRev)
        Me.Controls.Add(Me.TxtCode)
        Me.Controls.Add(Me.TxtName)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.lblerror2)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.CheckAll)
        Me.Controls.Add(Me.lblError)
        Me.Controls.Add(Me.CmdClear)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.CmdView)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CmdClose)
        Me.Controls.Add(Me.CmdSave)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmAddRHC"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "COMPOUND (MIXING)  "
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
    Dim vbal As Boolean
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
            If txtStep.Text = "1" Then
                'stage1
                StrSQL = " select * from (SELECT    seq,FinalCompound Final,Compcode ,c.Revision Rev,"
                StrSQL &= " Active,RMCode code,0.000 RHC,M.Qty QTY,c.Qty TQTY,unit"
                StrSQL &= " FROM         TBLCompound C"
                StrSQL &= " left outer join  "
                StrSQL &= " (select * from TBLMaster) M"
                StrSQL &= " on C.compcode+C.Revision = M.MasterCode+M.Revision )zzz"
                StrSQL &= " where seq = 1 "
                StrSQL &= " and Final = '" & TCompound.Trim & "'"
                StrSQL &= " and compcode = '" & TCode.Trim & "'"
                StrSQL &= " and Rev = '" & TRev.Trim & "'"
            ElseIf txtStep.Text = "2" Then
                'stage2
                StrSQL = " select * from "
                StrSQL &= " (   	  select * from "
                StrSQL &= " (  select * from (SELECT    seq,FinalCompound Final,Compcode ,c.Revision Rev,"
                StrSQL &= "  Active,RMCode code"
                StrSQL &= "  ,0.000 RHC,M.Qty QTY,unit"
                StrSQL &= "   FROM         TBLCompound C"
                StrSQL &= "   left outer join  "
                StrSQL &= "   (select * from TBLMaster) M"
                StrSQL &= "   on C.compcode+C.Revision = M.MasterCode+M.Revision )zzz"
                StrSQL &= "   where seq = 2 and code  in (select rmcode from TBLRM))RM"
                StrSQL &= " union "
                ' -- Stage2"
                StrSQL &= " select * from   	  	"
                StrSQL &= " (  select S2.seq,S2.Final,S2.Compcode,S2.Rev,S2.Active,"
                StrSQL &= " S1.RMcode,S1.RHC,S1.RmQty,S2.Unit from  ( select * from "
                StrSQL &= " (SELECT    seq,FinalCompound Final,Compcode ,c.Revision Rev,"
                StrSQL &= "   Active,RMCode code,isnull(rmRevision,'') rmRev"
                StrSQL &= "  ,0.000 RHC,M.Qty QTY,c.Qty TQTY,unit"
                StrSQL &= "   FROM         TBLCompound C"
                StrSQL &= "   left outer join  "
                StrSQL &= "   (select * from TBLMaster) M"
                StrSQL &= "   on C.compcode+C.Revision = M.MasterCode+M.Revision )zzz"
                StrSQL &= "   where seq = 2 and code  in "
                StrSQL &= "   (select compcode from TBLCompound)) S2"
                StrSQL &= " left outer join"
                ' -- compound 1"
                StrSQL &= "  (select Comp,cRev,Rcode RMCode,0.00 RHC,round(RMQty,3) RMQTY"
                StrSQL &= " ,Round(CQty,3) TQty from CompoundStage2)S1"
                StrSQL &= " on S2.Compcode+S2.rev = S1.Comp+S1.cRev  ) Stage2"
                StrSQL &= " union "
                ' -- Pigment "
                StrSQL &= " select * from ("
                StrSQL &= " select cp.seq,cp.final,cp.compcode,cp.Rev,cp.Active, "
                StrSQL &= " pg.rmcode,0.000 RHC,pg.Qty,pg.unit"
                StrSQL &= " from  (   select * from "
                StrSQL &= " (SELECT    seq,FinalCompound Final,Compcode ,c.Revision Rev,"
                StrSQL &= "  Active,RMCode code,isnull(rmRevision,'') rmRev"
                StrSQL &= " ,0.000 RHC,M.Qty QTY,c.Qty TQTY,unit"
                StrSQL &= "  FROM         TBLCompound C"
                StrSQL &= "  left outer join  "
                StrSQL &= "  (select * from TBLMaster) M"
                StrSQL &= "  on C.compcode+C.Revision = M.MasterCode+M.Revision )zzz"
                StrSQL &= "  where seq = 2 and code   in "
                StrSQL &= "  (select Pigmentcode from TBLPigment))cp"
                StrSQL &= " left outer join "
                StrSQL &= " (select p.pigmentcode,p.Revision PRev,RMcode,M.Qty,m.Unit from TBLPigment  p"
                StrSQL &= " left outer join TBLMaster m"
                StrSQL &= " on p.pigmentcode+p.Revision = m.Mastercode+m.Revision)pg"
                StrSQL &= " on cp.code+cp.RMRev = pg.pigmentcode+pg.Prev )Pigment"
                StrSQL &= " )Stage2"
                StrSQL &= " where  Final = '" & TCompound.Trim & "'"
                StrSQL &= " and compcode = '" & TCode.Trim & "'"
                StrSQL &= " and Rev = '" & TRev.Trim & "'"
            ElseIf txtStep.Text = "3" Then
                '--stage2
                StrSQL = " select * from ( "
                StrSQL &= " select * from ( select * from "
                StrSQL &= "  (  select * from (SELECT    seq,FinalCompound Final,Compcode ,c.Revision Rev,"
                StrSQL &= "  Active,RMCode code"
                StrSQL &= "  ,0.000 RHC,M.Qty QTY,unit"
                StrSQL &= "   FROM         TBLCompound C"
                StrSQL &= "  left outer join  "
                StrSQL &= "   (select * from TBLMaster) M"
                StrSQL &= "    on C.compcode+C.Revision = M.MasterCode+M.Revision )zzz"
                StrSQL &= "    where seq = 3 and code  in (select rmcode from TBLRM))RM"
                StrSQL &= "   union "
                ' -- Pigment 
                StrSQL &= "   select * from ("
                StrSQL &= "   select cp.seq,cp.final,cp.compcode,cp.Rev,cp.Active, "
                StrSQL &= "   pg.rmcode,0.000 RHC,pg.Qty,pg.unit"
                StrSQL &= "   from  (   select * from "
                StrSQL &= "   (SELECT    seq,FinalCompound Final,Compcode ,c.Revision Rev,"
                StrSQL &= "    Active,RMCode code,isnull(rmRevision,'') rmRev"
                StrSQL &= "   ,0.000 RHC,M.Qty QTY,c.Qty TQTY,unit"
                StrSQL &= "    FROM         TBLCompound C"
                StrSQL &= "    left outer join  "
                StrSQL &= "    (select * from TBLMaster) M"
                StrSQL &= "    on C.compcode+C.Revision = M.MasterCode+M.Revision )zzz"
                StrSQL &= "    where seq = 3 and code   in "
                StrSQL &= "    (select Pigmentcode from TBLPigment))cp"
                StrSQL &= "   left outer join "
                StrSQL &= "   (select p.pigmentcode,p.Revision PRev,RMcode,M.Qty,m.Unit from TBLPigment  p"
                StrSQL &= "   left outer join TBLMaster m"
                StrSQL &= "   on p.pigmentcode+p.Revision = m.Mastercode+m.Revision)pg"
                StrSQL &= "   on cp.code+cp.RMRev = pg.pigmentcode+pg.Prev )Pigment"
                StrSQL &= "    )Stage2"
                StrSQL &= "  union "
                ' -- Stage3
                StrSQL &= "  select * from  ("
                StrSQL &= "  select r.seq,r.final,cm.Mastercode,cm.Revision RRev,c.Active"
                StrSQL &= "  ,r.rmcode,0.000 RHC,round((weight/c.Qty)*m.Qty,3) QTY,m.unit from "
                StrSQL &= " TBLMaster cm"
                StrSQL &= " Right outer join "
                StrSQL &= " TBLRHCDtl r"
                StrSQL &= " on cm.RMcode+cm.RMRevision = r.mastercode+r.Revision"
                StrSQL &= " left outer join "
                StrSQL &= "  TBLCompound c"
                StrSQL &= " on r.Mastercode+r.Revision = c.Compcode+c.Revision"
                StrSQL &= "  left outer join "
                StrSQL &= "  TBLMaster m"
                StrSQL &= "  on r.Mastercode+r.Revision = M.RMcode+M.RMRevision and cm.QTY = m.Qty"
                StrSQL &= "  where r.seq = 2  ) Stage3"
                StrSQL &= " )bbb"
                StrSQL &= " where  Final like '%" & TCompound.Trim & "%'"
                StrSQL &= " and compcode = '" & TCode.Trim & "'"
                StrSQL &= " and Rev = '" & TRev.Trim & "'"
            Else
                '--stage4-9
                StrSQL = "  select * from ( select * from "
                StrSQL &= "   (  select * from (SELECT    seq,FinalCompound Final,Compcode ,c.Revision Rev,"
                StrSQL &= "   Active,RMCode code"
                StrSQL &= "    ,0.000 RHC,M.Qty QTY,unit"
                StrSQL &= "    FROM         TBLCompound C"
                StrSQL &= "    left outer join  "
                StrSQL &= "    (select * from TBLMaster) M"
                StrSQL &= "     on C.compcode+C.Revision = M.MasterCode+M.Revision )zzz"
                StrSQL &= "     where seq = " & txtStep.Text.Trim & " and code  in (select rmcode from TBLRM))RM"
                StrSQL &= "    union "
                '             -- Pigment 
                StrSQL &= "      select * from ("
                StrSQL &= "      select cp.seq,cp.final,cp.compcode,cp.Rev,cp.Active, "
                StrSQL &= "       pg.rmcode,0.000 RHC,pg.Qty,pg.unit"
                StrSQL &= "   from  (   select * from "
                StrSQL &= "   (SELECT    seq,FinalCompound Final,Compcode ,c.Revision Rev,"
                StrSQL &= "    Active,RMCode code,isnull(rmRevision,'') rmRev"
                StrSQL &= "   ,0.000 RHC,M.Qty QTY,c.Qty TQTY,unit"
                StrSQL &= "    FROM         TBLCompound C"
                StrSQL &= "    left outer join  "
                StrSQL &= "    (select * from TBLMaster) M"
                StrSQL &= "    on C.compcode+C.Revision = M.MasterCode+M.Revision )zzz"
                StrSQL &= "    where seq = " & txtStep.Text.Trim & " and code   in "
                StrSQL &= "    (select Pigmentcode from TBLPigment))cp"
                StrSQL &= "   left outer join "
                StrSQL &= "   (select p.pigmentcode,p.Revision PRev,RMcode,M.Qty,m.Unit from TBLPigment  p"
                StrSQL &= "   left outer join TBLMaster m"
                StrSQL &= "   on p.pigmentcode+p.Revision = m.Mastercode+m.Revision)pg"
                StrSQL &= "   on cp.code+cp.RMRev = pg.pigmentcode+pg.Prev )Pigment"

                StrSQL &= "   union "
                '             -- StagetxtStep.text.trim
                StrSQL &= "        select * from   	  	"
                StrSQL &= "      ( "

                StrSQL &= "  select r.seq,r.final,cm.Mastercode,cm.Revision RRev,c.Active"
                StrSQL &= "  ,r.rmcode,0.000 RHC,round((weight/c.Qty)*m.Qty,3) QTY,m.unit from "
                StrSQL &= " TBLMaster cm"
                StrSQL &= "  Right outer join "
                StrSQL &= " TBLRHCDtl r"
                StrSQL &= " on cm.RMcode+cm.RMRevision = r.mastercode+r.Revision"
                StrSQL &= "  left outer join "
                StrSQL &= "  TBLCompound c"
                StrSQL &= "  on r.Mastercode+r.Revision = c.Compcode+c.Revision"
                StrSQL &= "  left outer join "
                StrSQL &= "  (select * from  TBLMaster where Mastercode = '" & TCode.Trim & "') m"
                StrSQL &= "  on r.Mastercode+r.Revision = M.RMcode+M.RMRevision"
                StrSQL &= "  where r.seq in ('1','2','3','4','5','6','7','8','9') )Stage4"
                StrSQL &= " )bbb"
                StrSQL &= " where compcode = '" & TCode.Trim & "'"
                StrSQL &= " and Rev = '" & TRev.Trim & "'"
                '   StrSQL &= " and  Final like '%" & TCompound.Trim & "%'"
            End If
        ElseIf CmdSave.Text = "Edit" Then
            StrSQL = " select seq,final,Mastercode Compcode,Revision Rev,RMCode code,"
            StrSQL &= "  Weight Qty ,RHC,'KG' Unit  from TBLRHCDtl"
            StrSQL &= " where  Final = '" & TCompound.Trim & "'"
            StrSQL &= " and Mastercode = '" & TCode.Trim & "'"
            StrSQL &= " and Revision = '" & TRev.Trim & "'"
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

        cm = CType(Me.BindingContext(DataGridRM.DataSource, DataGridRM.DataMember), CurrencyManager)
        Dim c As CheckRowHeader
        c = AddressOf CheckRowHeader

        Dim grdColStyle2 As New DataGridQtyBox(c)
        With grdColStyle2
            .HeaderText = "RHC"
            .MappingName = "RHC"
            .Format = "###,##0.000"
            .Width = 110
            .Alignment = HorizontalAlignment.Right
            .NullText = ""
        End With
        Dim grdColStyle3 As New DataGridQtyBox(c)
        With grdColStyle3
            .HeaderText = "Qty(KG)"
            .MappingName = "Qty"
            .Width = 110
            .Format = "###,##0.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle5 As New DataGridColoredLine2
        With grdColStyle5
            .HeaderText = "Unit"
            .MappingName = "unit"
            .Width = 80
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With

        grdTableStyle1.GridColumnStyles.AddRange _
(New DataGridColumnStyle() _
{grdColStyle1, grdColStyle3, grdColStyle2, grdColStyle5})

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

    Private Sub FrmAddRHC_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If CBal = True Then
            ChkCal.Checked = True
        Else
            ChkCal.Checked = False
        End If
        If CmdSave.Text = "Edit" Then
            TxtCode.Text = TCode
            TxtCompound.Text = TCompound
            TxtRev.Text = TRev
            txtStep.Text = TStep
            LoadRM()
        Else
            TxtCode.Text = TCode
            TxtCompound.Text = TCompound
            TxtRev.Text = TRev
            txtStep.Text = TStep
            LoadRM()
        End If
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

        Dim i As Integer
        TxtCode.Text = TxtCode.Text.ToUpper

        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        Dim aDr() As DataRow
        GrdDV.RowFilter = " Qty <> 0.000"
        aDr = GrdDV.Table.Select(GrdDV.RowFilter)
        iTotal = 0
        qTotal = 0
        CheckAll.Checked = False
        Dim dr As DataRow
        For Each dr In aDr
            With dr
                If IIf(.Item("Code") Is System.DBNull.Value, "", .Item("Code")) <> "" Then
                    iTotal = iTotal + .Item("RHC")
                    qTotal = qTotal + .Item("Qty")
                End If
            End With
        Next

        msg = "Compound Qty Total :" & qTotal ' Define message.
        msg += "   Compound RHC Total :" & iTotal
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Compound"   ' Define title.
        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            RM()
            PgBar1.Visible = True

            msg = "You Want to Calculate Percent of Compound.Click Yes ! "
            style = MsgBoxStyle.DefaultButton2 Or _
               MsgBoxStyle.Information Or MsgBoxStyle.YesNo
            title = "Compound"   ' Define title.
            ' Display message.
            '  response = MsgBox(msg, style, title)
            '' If response = MsgBoxResult.Yes Then ' User chose Yes.

            vbal = True

            If CBal = True Then

                ' Display the ProgressBar control.
                PgBar1.Visible = True
                ' Set Minimum to 1 to represent the first file being copied.
                PgBar1.Minimum = 0
                ' Set Maximum to the total number of files to copy.
                PgBar1.Maximum = aDr.Length
                ' Set the initial value of the ProgressBar.
                PgBar1.Value = 1
                ' Set the Step property to a value of 1 to represent each file being copied.
                PgBar1.Step = 1
                Dim RMstr, Qtystr, RHCstr As String
                For Each dr In aDr
                    With dr
                        If IIf(.Item("Code") Is System.DBNull.Value, "", .Item("Code")) <> "" Then
                            RMstr = .Item("Code")
                            Qtystr = .Item("Qty")
                            RHCstr = .Item("RHC")

                            CALPer(RMstr, Qtystr, RHCstr)
                            i = i + 1
                        End If
                    End With
                    PgBar1.PerformStep()
                Next

                CalTotalPercent(TxtCompound.Text.Trim, TxtCode.Text.Trim, TxtRev.Text.Trim)
            End If


            If vbal Then
                MsgBox("Cal Percent Complete.  " & i & " Record")
            Else
                MsgBox("Cal Percent not Complete.")
            End If
            '  End If
            PgBar1.PerformStep()
            Me.Close()
        Else
            Exit Sub
        End If
    End Sub

    Private Function CalPercent(ByVal final As String, ByVal Mastercode As String, _
  ByVal Rev As String, ByVal RMcode As String, ByVal Weight As Double, ByVal RHC As Double) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        CalPercent = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CalPercent"
        cmd2.Connection = cnn

        Dim sparam0 As SqlClient.SqlParameter
        sparam0 = New SqlClient.SqlParameter
        sparam0.ParameterName = "@Final"
        sparam0.SqlDbType = SqlDbType.Char
        sparam0.Size = 20
        sparam0.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam0)

        Dim sparam1 As SqlClient.SqlParameter
        sparam1 = New SqlClient.SqlParameter
        sparam1.ParameterName = "@MasterCode"
        sparam1.SqlDbType = SqlDbType.Char
        sparam1.Size = 20
        sparam1.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam1)

        Dim sparam2 As SqlClient.SqlParameter
        sparam2 = New SqlClient.SqlParameter
        sparam2.ParameterName = "@REV"
        sparam2.SqlDbType = SqlDbType.Char
        sparam2.Size = 3
        sparam2.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam2)

        Dim sparam3 As SqlClient.SqlParameter
        sparam3 = New SqlClient.SqlParameter
        sparam3.ParameterName = "@RMCode"
        sparam3.SqlDbType = SqlDbType.Char
        sparam3.Size = 20
        sparam3.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam3)

        Dim sparam4 As SqlClient.SqlParameter
        sparam4 = New SqlClient.SqlParameter
        sparam4.ParameterName = "@RHC"
        sparam4.SqlDbType = SqlDbType.Float
        '    sparam4.Size = 20
        sparam4.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam4)

        Dim sparam5 As SqlClient.SqlParameter
        sparam5 = New SqlClient.SqlParameter
        sparam5.ParameterName = "@Weight"
        sparam5.SqlDbType = SqlDbType.Float
        '    sparam5.Size = 20
        sparam5.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam5)

        Dim sparam6 As SqlClient.SqlParameter
        sparam6 = New SqlClient.SqlParameter
        sparam6.ParameterName = "@errID"
        sparam6.SqlDbType = SqlDbType.Char
        sparam6.Size = 4
        sparam6.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam6)

        Dim sparam7 As SqlClient.SqlParameter
        sparam7 = New SqlClient.SqlParameter
        sparam7.ParameterName = "@errMsg"
        sparam7.SqlDbType = SqlDbType.Char
        sparam7.Size = 40
        sparam7.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam7)

        Dim Reader As SqlClient.SqlDataReader
        cmd2.Parameters("@Final").Value = final.Trim
        cmd2.Parameters("@MasterCode").Value = Mastercode.Trim
        cmd2.Parameters("@Rev").Value = Rev.Trim
        cmd2.Parameters("@RMCode").Value = RMcode.Trim
        cmd2.Parameters("@Weight").Value = Weight
        cmd2.Parameters("@RHC").Value = RHC

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            CalPercent = True
        Catch ex As EvaluateException
            MsgBox(ex.Message, 48)
            CalPercent = False
        Catch ex As Exception
            ' MsgBox(ex.Message, 48)
            CalPercent = True
        End Try
        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Function

    Private Function CalTotalPercent(ByVal final As String, ByVal Mastercode As String, _
   ByVal Rev As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        CalTotalPercent = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CalTotalPercent"
        cmd2.Connection = cnn

        Dim sparam0 As SqlClient.SqlParameter
        sparam0 = New SqlClient.SqlParameter
        sparam0.ParameterName = "@Final"
        sparam0.SqlDbType = SqlDbType.Char
        sparam0.Size = 20
        sparam0.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam0)

        Dim sparam1 As SqlClient.SqlParameter
        sparam1 = New SqlClient.SqlParameter
        sparam1.ParameterName = "@MasterCode"
        sparam1.SqlDbType = SqlDbType.Char
        sparam1.Size = 20
        sparam1.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam1)

        Dim sparam2 As SqlClient.SqlParameter
        sparam2 = New SqlClient.SqlParameter
        sparam2.ParameterName = "@REV"
        sparam2.SqlDbType = SqlDbType.Char
        sparam2.Size = 3
        sparam2.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam2)

        Dim sparam3 As SqlClient.SqlParameter
        sparam3 = New SqlClient.SqlParameter
        sparam3.ParameterName = "@errID"
        sparam3.SqlDbType = SqlDbType.Char
        sparam3.Size = 4
        sparam3.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam3)

        Dim sparam4 As SqlClient.SqlParameter
        sparam4 = New SqlClient.SqlParameter
        sparam4.ParameterName = "@errMsg"
        sparam4.SqlDbType = SqlDbType.Char
        sparam4.Size = 40
        sparam4.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam4)

        Dim Reader As SqlClient.SqlDataReader
        cmd2.Parameters("@Final").Value = final.Trim
        cmd2.Parameters("@MasterCode").Value = Mastercode.Trim
        cmd2.Parameters("@Rev").Value = Rev.Trim

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            CalTotalPercent = True
        Catch ex As Exception
            MsgBox("การคำนวณผลรวมอาจผิดพลาดได้ กรุณาตรวจสอบ เปอร์เซ็นต์ผลรวมอีกครั้ง", 48)
            CalTotalPercent = False
        End Try
        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Function

    Private Sub DataGridRM_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridRM.CurrentCellChanged
        oldrow = DataGridRM.CurrentCell.RowNumber
    End Sub

    Private Sub CmdView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdView.Click
        GrdDV.RowFilter = "  descname like'%" & TxtName.Text.Trim & "%'"
        DataGridRM.DataSource = GrdDV
    End Sub

#Region "RM"
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
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "SQL Error")
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
        Dim cn As New SqlConnection(C1.Strcon)
        Dim cmd As New SqlCommand(strsql, cn)
        cn.Open()
        Dim t1 As SqlTransaction = cn.BeginTransaction
        cmd.Transaction = t1
        Dim strDate, SD() As String
        SD = Split(Date.Now.ToShortDateString, "/")
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
                Dim RMstr As String = String.Empty
                Dim CodeStr As String = String.Empty
                Dim Qtystr As String = String.Empty
                Dim RHCstr As String = String.Empty
                For Each dr In aDr
                    With dr
                        If IIf(.Item("Code") Is System.DBNull.Value, "", .Item("Code")) <> "" Then
                            RMstr = .Item("Code")
                            CodeStr = .Item("Code")
                            Qtystr = .Item("Qty")
                            RHCstr = .Item("RHC")
                        End If
                    End With
                    strsql = "Insert TBLRHCDtl "
                    strsql += " Values(" & PrepareStr(txtStep.Text.Trim)
                    strsql += "," & PrepareStr(TxtCompound.Text.Trim)
                    strsql += "," & PrepareStr(TxtCode.Text.Trim)
                    strsql += "," & PrepareStr(TxtRev.Text.Trim)
                    strsql += "," & PrepareStr(CodeStr)
                    strsql += "," & PrepareStr(Qtystr)
                    strsql += "," & PrepareStr(RHCstr)
                    strsql += "," & PrepareStr("")
                    strsql += "," & PrepareStr(strDate.Trim)
                    strsql += ")"
                    cmd.CommandText = strsql
                    cmd.ExecuteNonQuery()
                    strsql = ""
                Next

                Try
                    strsql += ""
                    strsql += "   Update TBLCompound "
                    strsql += " set RHC = " & PrepareStr(iTotal)
                    strsql += " , Qty = " & PrepareStr(qTotal)
                    strsql += " where FinalCompound = " & PrepareStr(TxtCompound.Text.Trim)
                    strsql += " and  Compcode = " & PrepareStr(TxtCode.Text.Trim)
                    strsql += " and  Revision = " & PrepareStr(TxtRev.Text.Trim)
                    cmd.CommandText = strsql
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
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
                            strsql = "Update TBLRHCDtl "
                            strsql += " Set Weight = " & PrepareStr(.Item("Qty"))
                            strsql += ", RHC =" & PrepareStr(.Item("RHC"))
                            strsql += " where Final = " & PrepareStr(TxtCompound.Text.Trim)
                            strsql += " and  Mastercode = " & PrepareStr(TxtCode.Text.Trim)
                            strsql += " and  Revision = " & PrepareStr(TxtRev.Text.Trim)
                            strsql += " and  RMcode = " & PrepareStr(.Item("code"))
                            strsql += " and  Weight = " & PrepareStr(.Item("Qty"))
                            cmd.CommandText = strsql
                            cmd.ExecuteNonQuery()
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

                    strsql = "  Update TBLCompound "
                    strsql += " set RHC = " & PrepareStr(iTotal)
                    strsql += " , Qty = " & PrepareStr(qTotal)
                    strsql += " where FinalCompound = " & PrepareStr(TxtCompound.Text.Trim)
                    strsql += " and  Compcode = " & PrepareStr(TxtCode.Text.Trim)
                    strsql += " and  Revision = " & PrepareStr(TxtRev.Text.Trim)
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
    Sub CALPer(ByVal rm As String, ByVal qty As String, ByVal rhc As String)
        Dim strsql As String = String.Empty
        Dim cn As New SqlConnection(C1.Strcon)
        Dim cmd As New SqlCommand(strsql, cn)
        cn.Open()
        Dim t1 As SqlTransaction = cn.BeginTransaction
        cmd.Transaction = t1
        Dim strDate, SD() As String
        SD = Split(Date.Now.ToShortDateString, "/")
        strDate = SD(2) + SD(1) + SD(0)
        Try
            'Dim aDr() As DataRow
            'GrdDV.RowFilter = " Qty <> 0.000"
            'aDr = GrdDV.Table.Select(GrdDV.RowFilter)
            'If UBound(aDr) < 0 Then
            '    Exit Sub
            'End If
            'Dim dr As DataRow
            'Dim RMstr, CodeStr, Qtystr, RHCstr, Finalstr, Revstr As String
            'Dim i As Integer
            'For Each dr In aDr
            '    With dr
            '        If IIf(.Item("Code") Is System.DBNull.Value, "", .Item("Code")) <> "" Then
            '            RMstr = .Item("Rmcode")
            '            CodeStr = .Item("Code")
            '            Qtystr = .Item("Qty")
            '            RHCstr = .Item("RHC")
            '            Finalstr = .Item("Final")
            '            Revstr = .Item("Rev")
            '        End If
            '    End With
            strsql = "   update TBLRHCDtl "
            strsql += " set Per = (     "
            strsql += " select Per from ( "
            strsql += " SELECT Seq,Final,CRev,xx.compcode,mm.Rev,mm.RMCode,mQty,mm.mRHC,Per"
            strsql += "  FROM   "
            strsql += " (    select seq,final,compcode,Revision REV,cRev,RHC,Active,RMcode,mRHC,mQty from  "
            strsql += "  (SELECT  seq,Finalcompound final,compcode,Revision,compcode+','+Revision CRev,RHC,Active from  TBLCompound) c"
            strsql += "  left outer join "
            strsql += "  ( select mastercode code,Revision Rev"
            strsql += "  ,mastercode+','+Revision MRev,RMcode,RHC mRHC,Weight mQty from   TBLRHCDtl"
            strsql += "    where Mastercode in "
            strsql += "  ( SELECT  compcode"
            strsql += "  FROM         TBLCompound))m"
            strsql += "  on c.CRev = m.MRev "
            strsql += " where Final = '" & TCompound & "'  and  Compcode = '" & TCode & "' and  Rev = '" & TRev & "' and RMcode = '" & rm & "'"
            strsql += " )mm"
            strsql += "   left outer join "
            strsql += "   (  select aa.compcode,aa.Rev,Rmcode rcode,mrhc,mQty Qty,aa.Per from "
            strsql += "   (SELECT     cc.Seq, cc.FinalCompound Final, cc.CompCode, cc.Revision REV, aa.RMCode,mQty, cc.RHC, aa.mRHC, ROUND(aa.mRHC / cc.RHC * 100, 3) AS per, cc.Active"
            strsql += "    FROM         (SELECT     Seq, FinalCompound, CompCode, Revision, Qty TQty, RHC, Active, CompCode + ',' + Revision Code"
            strsql += "     FROM          TBLCompound) cc LEFT OUTER JOIN"
            strsql += "     (SELECT     CRev, Code, Rev, RMCode, mRHC, mQty"
            strsql += "     FROM          (SELECT     seq, Finalcompound, compcode + ',' + Revision CRev, RHC, Active"
            strsql += "      FROM          TBLCompound) c LEFT OUTER JOIN"
            strsql += "   (SELECT     mastercode code, Revision Rev, mastercode + ',' + Revision MRev, RMcode, RHC mRHC, Weight mQty"
            strsql += "     FROM          TBLRHCDtl"
            strsql += "   WHERE      Mastercode IN"
            strsql += "   (SELECT     compcode"
            strsql += "   FROM          TBLCompound)) m ON c.CRev = m.MRev) aa ON cc.Code = aa.CRev"
            strsql += "   )aa)xx  on mm.CRev+mm.RMCode  = xx.compcode+','+xx.Rev+xx.rcode and  mm.mrhc = xx.mrhc and mQty =Qty"
            strsql += "   )zz "
            strsql += " where Final = '" & TCompound & "'  and  Compcode = '" & TCode & "' and  Rev = '" & TRev & "' and RMcode = '" & rm & "' and mRHC = '" & rhc & "' and mQty = '" & qty & "')"
            strsql += "  where Final = '" & TCompound & "'  and  Mastercode = '" & TCode & "' and  Revision = '" & TRev & "' and RMcode = '" & rm & "' and RHC = '" & rhc & "'  and Weight = '" & qty & "'"
            cmd.CommandText = strsql
            cmd.ExecuteNonQuery()
            'Next

            t1.Commit()
            'MsgBox("Calculate Complete.", MsgBoxStyle.Information, "Compound Code")
        Catch
            vbal = False
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

    Private Sub TxtName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtName.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case Else
                'If Not IsNumeric(e.KeyChar) Then
                '    e.Handled = True
                'Else
                'End If
        End Select

    End Sub

    Private Sub CmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClear.Click
        LoadRM()
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

    Private Sub TxtCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCode.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                TxtCode.Text = TxtCode.Text.ToUpper
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
                txtStep.Text = iNoSeq(TxtCompound.Text.Trim, TxtRev.Text.Trim) + 1
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

    Private Sub ChkCal_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkCal.CheckedChanged
        If ChkCal.Checked Then
            CBal = True
        Else
            CBal = False
        End If
    End Sub
End Class
