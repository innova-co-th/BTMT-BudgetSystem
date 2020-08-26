#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Inventory_Record.Common
#End Region

Public Class FrmAddGreenTire
#Region "Declare"
    Inherits System.Windows.Forms.Form
    Public Shared GrdDVBSJ As New DataView
    Protected Const TBL_BSJ As String = "TBL_BSJ"
    Public Shared GrdDVCU As New DataView
    Protected Const TBL_CU As String = "TBL_CU"
    Public Shared GrdDVSD As New DataView
    Protected Const TBL_SD As String = "TBL_SD"
    Public Shared GrdDVBF As New DataView
    Protected Const TBL_BF As String = "TBL_BF"
    Public Shared GrdDVTT As New DataView
    Protected Const TBL_TT As String = "TBL_TT"
    Public Shared GrdDVIN As New DataView
    Protected Const TBL_IN As String = "TBL_IN"
    Public Shared GrdDVNy As New DataView
    Protected Const TBL_Ny As String = "TBL_Ny"
    Public Shared GrdDVWf As New DataView
    Protected Const TBL_Wf As String = "TBL_Wf"
    Public Shared GrdDVB1 As New DataView
    Protected Const TBL_B1 As String = "TBL_B1"
    Public Shared GrdDVB2 As New DataView
    Protected Const TBL_B2 As String = "TBL_B2"
    Public Shared GrdDVB3 As New DataView
    Protected Const TBL_B3 As String = "TBL_B3"
    Public Shared GrdDVB4 As New DataView
    Protected Const TBL_B4 As String = "TBL_B4"
    Public Shared GrdDVBp As New DataView
    Protected Const TBL_Bp As String = "TBL_Bp"

    Public Shared GrdDV As New DataView
    Protected Const TBL_Tire As String = "TBL_Tire"
    Protected DefaultGridBorderStyle As BorderStyle
    Public Shared cm As CurrencyManager
    Dim iTotal As Double
    Friend iCmb As String
#End Region
    Friend TTcode, BFcode, Bpcode, CUcode, SDcode, INcode, _
            B1code, B2code, B3code, B4code, Wfcode, NYcode, BSJCode As String

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
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox8 As System.Windows.Forms.GroupBox
    Friend WithEvents TxtRev As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblError As System.Windows.Forms.Label
    Friend WithEvents TxtCode As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox10 As System.Windows.Forms.GroupBox
    Friend WithEvents GPTire As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox11 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox12 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox9 As System.Windows.Forms.GroupBox
    Friend WithEvents CmdClose As System.Windows.Forms.Button
    Friend WithEvents CmdSave As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtCU_L As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents cmbBF As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSD As System.Windows.Forms.ComboBox
    Friend WithEvents cmbIN As System.Windows.Forms.ComboBox
    Friend WithEvents cmbCU As System.Windows.Forms.ComboBox
    Friend WithEvents cmbTT As System.Windows.Forms.ComboBox
    Friend WithEvents CheckBoxWf As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxNy As System.Windows.Forms.CheckBox
    Friend WithEvents cmbWf As System.Windows.Forms.ComboBox
    Friend WithEvents cmbNy As System.Windows.Forms.ComboBox
    Friend WithEvents cmbB1 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbB2 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbB3 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbB4 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbBp As System.Windows.Forms.ComboBox
    Friend WithEvents TxtWf_L As System.Windows.Forms.TextBox
    Friend WithEvents TxtWf_N As System.Windows.Forms.TextBox
    Friend WithEvents txtBF_N As System.Windows.Forms.TextBox
    Friend WithEvents TxtB1_L As System.Windows.Forms.TextBox
    Friend WithEvents TxtB1_N As System.Windows.Forms.TextBox
    Friend WithEvents TxtB2_L As System.Windows.Forms.TextBox
    Friend WithEvents TxtB2_N As System.Windows.Forms.TextBox
    Friend WithEvents TxtB3_L As System.Windows.Forms.TextBox
    Friend WithEvents TxtB3_N As System.Windows.Forms.TextBox
    Friend WithEvents TxtB4_L As System.Windows.Forms.TextBox
    Friend WithEvents TxtB4_N As System.Windows.Forms.TextBox
    Friend WithEvents TxtIN_L As System.Windows.Forms.TextBox
    Friend WithEvents TxtIN_N As System.Windows.Forms.TextBox
    Friend WithEvents TxtSD_L As System.Windows.Forms.TextBox
    Friend WithEvents TxtSD_N As System.Windows.Forms.TextBox
    Friend WithEvents TxtCu_N As System.Windows.Forms.TextBox
    Friend WithEvents TxtTT_N As System.Windows.Forms.TextBox
    Friend WithEvents TxtNy_L As System.Windows.Forms.TextBox
    Friend WithEvents TxtNy_N As System.Windows.Forms.TextBox
    Friend WithEvents TxtBp_L As System.Windows.Forms.TextBox
    Friend WithEvents TxtBP_N As System.Windows.Forms.TextBox
    Friend WithEvents CheckBoxTire As System.Windows.Forms.CheckBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents CheckAll As System.Windows.Forms.CheckBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents CmbBSJCode As System.Windows.Forms.ComboBox
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents txtremark As System.Windows.Forms.TextBox
    Friend WithEvents txtremark2 As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmAddGreenTire))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.TxtWf_L = New System.Windows.Forms.TextBox
        Me.TxtWf_N = New System.Windows.Forms.TextBox
        Me.Label23 = New System.Windows.Forms.Label
        Me.Label24 = New System.Windows.Forms.Label
        Me.CheckBoxWf = New System.Windows.Forms.CheckBox
        Me.cmbWf = New System.Windows.Forms.ComboBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.txtBF_N = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.cmbBF = New System.Windows.Forms.ComboBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.TxtB1_L = New System.Windows.Forms.TextBox
        Me.TxtB1_N = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.cmbB1 = New System.Windows.Forms.ComboBox
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.TxtB2_L = New System.Windows.Forms.TextBox
        Me.TxtB2_N = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.cmbB2 = New System.Windows.Forms.ComboBox
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.TxtB3_L = New System.Windows.Forms.TextBox
        Me.TxtB3_N = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.cmbB3 = New System.Windows.Forms.ComboBox
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.TxtB4_L = New System.Windows.Forms.TextBox
        Me.TxtB4_N = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.Label22 = New System.Windows.Forms.Label
        Me.cmbB4 = New System.Windows.Forms.ComboBox
        Me.GroupBox7 = New System.Windows.Forms.GroupBox
        Me.TxtIN_L = New System.Windows.Forms.TextBox
        Me.TxtIN_N = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.cmbIN = New System.Windows.Forms.ComboBox
        Me.GroupBox8 = New System.Windows.Forms.GroupBox
        Me.TxtSD_L = New System.Windows.Forms.TextBox
        Me.TxtSD_N = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.cmbSD = New System.Windows.Forms.ComboBox
        Me.TxtRev = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.lblError = New System.Windows.Forms.Label
        Me.TxtCode = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.CheckBoxTire = New System.Windows.Forms.CheckBox
        Me.GPTire = New System.Windows.Forms.GroupBox
        Me.GroupBox10 = New System.Windows.Forms.GroupBox
        Me.TxtCU_L = New System.Windows.Forms.TextBox
        Me.TxtCu_N = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.cmbCU = New System.Windows.Forms.ComboBox
        Me.GroupBox11 = New System.Windows.Forms.GroupBox
        Me.cmbTT = New System.Windows.Forms.ComboBox
        Me.TxtTT_N = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.GroupBox12 = New System.Windows.Forms.GroupBox
        Me.TxtNy_L = New System.Windows.Forms.TextBox
        Me.TxtNy_N = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.Label26 = New System.Windows.Forms.Label
        Me.CheckBoxNy = New System.Windows.Forms.CheckBox
        Me.cmbNy = New System.Windows.Forms.ComboBox
        Me.GroupBox9 = New System.Windows.Forms.GroupBox
        Me.cmbBp = New System.Windows.Forms.ComboBox
        Me.TxtBp_L = New System.Windows.Forms.TextBox
        Me.TxtBP_N = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.CmdClose = New System.Windows.Forms.Button
        Me.CmdSave = New System.Windows.Forms.Button
        Me.CheckAll = New System.Windows.Forms.CheckBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtremark = New System.Windows.Forms.TextBox
        Me.CmbBSJCode = New System.Windows.Forms.ComboBox
        Me.txtremark2 = New System.Windows.Forms.TextBox
        Me.Label27 = New System.Windows.Forms.Label
        Me.Label28 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.GroupBox8.SuspendLayout()
        Me.GroupBox10.SuspendLayout()
        Me.GroupBox11.SuspendLayout()
        Me.GroupBox12.SuspendLayout()
        Me.GroupBox9.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TxtWf_L)
        Me.GroupBox1.Controls.Add(Me.TxtWf_N)
        Me.GroupBox1.Controls.Add(Me.Label23)
        Me.GroupBox1.Controls.Add(Me.Label24)
        Me.GroupBox1.Controls.Add(Me.CheckBoxWf)
        Me.GroupBox1.Controls.Add(Me.cmbWf)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(264, 56)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(232, 72)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Wire Chafer"
        '
        'TxtWf_L
        '
        Me.TxtWf_L.Location = New System.Drawing.Point(176, 38)
        Me.TxtWf_L.Name = "TxtWf_L"
        Me.TxtWf_L.Size = New System.Drawing.Size(48, 20)
        Me.TxtWf_L.TabIndex = 2
        Me.TxtWf_L.Text = ""
        '
        'TxtWf_N
        '
        Me.TxtWf_N.Location = New System.Drawing.Point(176, 14)
        Me.TxtWf_N.Name = "TxtWf_N"
        Me.TxtWf_N.Size = New System.Drawing.Size(48, 20)
        Me.TxtWf_N.TabIndex = 1
        Me.TxtWf_N.Text = ""
        '
        'Label23
        '
        Me.Label23.Location = New System.Drawing.Point(136, 40)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(40, 16)
        Me.Label23.TabIndex = 10
        Me.Label23.Text = "Length"
        '
        'Label24
        '
        Me.Label24.Location = New System.Drawing.Point(136, 16)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(32, 16)
        Me.Label24.TabIndex = 9
        Me.Label24.Text = "Num"
        '
        'CheckBoxWf
        '
        Me.CheckBoxWf.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CheckBoxWf.Location = New System.Drawing.Point(8, 16)
        Me.CheckBoxWf.Name = "CheckBoxWf"
        Me.CheckBoxWf.Size = New System.Drawing.Size(56, 16)
        Me.CheckBoxWf.TabIndex = 30
        Me.CheckBoxWf.Text = "Use"
        '
        'cmbWf
        '
        Me.cmbWf.Enabled = False
        Me.cmbWf.Location = New System.Drawing.Point(8, 40)
        Me.cmbWf.Name = "cmbWf"
        Me.cmbWf.Size = New System.Drawing.Size(120, 21)
        Me.cmbWf.TabIndex = 0
        Me.cmbWf.Text = "Select"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtBF_N)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.cmbBF)
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox2.Location = New System.Drawing.Point(16, 56)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(232, 48)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "BF (Upper-Lower-Center)"
        '
        'txtBF_N
        '
        Me.txtBF_N.Location = New System.Drawing.Point(176, 16)
        Me.txtBF_N.Name = "txtBF_N"
        Me.txtBF_N.Size = New System.Drawing.Size(48, 20)
        Me.txtBF_N.TabIndex = 1
        Me.txtBF_N.Text = ""
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(136, 18)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(32, 16)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Num"
        '
        'cmbBF
        '
        Me.cmbBF.Location = New System.Drawing.Point(8, 16)
        Me.cmbBF.Name = "cmbBF"
        Me.cmbBF.Size = New System.Drawing.Size(120, 21)
        Me.cmbBF.TabIndex = 0
        Me.cmbBF.Text = "Select"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.TxtB1_L)
        Me.GroupBox3.Controls.Add(Me.TxtB1_N)
        Me.GroupBox3.Controls.Add(Me.Label15)
        Me.GroupBox3.Controls.Add(Me.Label16)
        Me.GroupBox3.Controls.Add(Me.cmbB1)
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox3.Location = New System.Drawing.Point(264, 136)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(232, 72)
        Me.GroupBox3.TabIndex = 10
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Belt-1"
        '
        'TxtB1_L
        '
        Me.TxtB1_L.Location = New System.Drawing.Point(176, 38)
        Me.TxtB1_L.Name = "TxtB1_L"
        Me.TxtB1_L.Size = New System.Drawing.Size(48, 20)
        Me.TxtB1_L.TabIndex = 2
        Me.TxtB1_L.Text = ""
        '
        'TxtB1_N
        '
        Me.TxtB1_N.Location = New System.Drawing.Point(176, 14)
        Me.TxtB1_N.Name = "TxtB1_N"
        Me.TxtB1_N.Size = New System.Drawing.Size(48, 20)
        Me.TxtB1_N.TabIndex = 1
        Me.TxtB1_N.Text = ""
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(136, 40)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(40, 16)
        Me.Label15.TabIndex = 10
        Me.Label15.Text = "Length"
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(136, 16)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(32, 16)
        Me.Label16.TabIndex = 9
        Me.Label16.Text = "Num"
        '
        'cmbB1
        '
        Me.cmbB1.Location = New System.Drawing.Point(8, 16)
        Me.cmbB1.Name = "cmbB1"
        Me.cmbB1.Size = New System.Drawing.Size(120, 21)
        Me.cmbB1.TabIndex = 0
        Me.cmbB1.Text = "Select"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.TxtB2_L)
        Me.GroupBox4.Controls.Add(Me.TxtB2_N)
        Me.GroupBox4.Controls.Add(Me.Label17)
        Me.GroupBox4.Controls.Add(Me.Label18)
        Me.GroupBox4.Controls.Add(Me.cmbB2)
        Me.GroupBox4.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox4.Location = New System.Drawing.Point(264, 216)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(232, 72)
        Me.GroupBox4.TabIndex = 11
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Belt-2"
        '
        'TxtB2_L
        '
        Me.TxtB2_L.Location = New System.Drawing.Point(176, 38)
        Me.TxtB2_L.Name = "TxtB2_L"
        Me.TxtB2_L.Size = New System.Drawing.Size(48, 20)
        Me.TxtB2_L.TabIndex = 2
        Me.TxtB2_L.Text = ""
        '
        'TxtB2_N
        '
        Me.TxtB2_N.Location = New System.Drawing.Point(176, 14)
        Me.TxtB2_N.Name = "TxtB2_N"
        Me.TxtB2_N.Size = New System.Drawing.Size(48, 20)
        Me.TxtB2_N.TabIndex = 1
        Me.TxtB2_N.Text = ""
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(136, 40)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(40, 16)
        Me.Label17.TabIndex = 10
        Me.Label17.Text = "Length"
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(136, 16)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(32, 16)
        Me.Label18.TabIndex = 9
        Me.Label18.Text = "Num"
        '
        'cmbB2
        '
        Me.cmbB2.Location = New System.Drawing.Point(8, 16)
        Me.cmbB2.Name = "cmbB2"
        Me.cmbB2.Size = New System.Drawing.Size(120, 21)
        Me.cmbB2.TabIndex = 0
        Me.cmbB2.Text = "Select"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.TxtB3_L)
        Me.GroupBox5.Controls.Add(Me.TxtB3_N)
        Me.GroupBox5.Controls.Add(Me.Label19)
        Me.GroupBox5.Controls.Add(Me.Label20)
        Me.GroupBox5.Controls.Add(Me.cmbB3)
        Me.GroupBox5.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox5.Location = New System.Drawing.Point(264, 296)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(232, 72)
        Me.GroupBox5.TabIndex = 12
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Belt-3"
        '
        'TxtB3_L
        '
        Me.TxtB3_L.Location = New System.Drawing.Point(176, 38)
        Me.TxtB3_L.Name = "TxtB3_L"
        Me.TxtB3_L.Size = New System.Drawing.Size(48, 20)
        Me.TxtB3_L.TabIndex = 2
        Me.TxtB3_L.Text = ""
        '
        'TxtB3_N
        '
        Me.TxtB3_N.Location = New System.Drawing.Point(176, 14)
        Me.TxtB3_N.Name = "TxtB3_N"
        Me.TxtB3_N.Size = New System.Drawing.Size(48, 20)
        Me.TxtB3_N.TabIndex = 1
        Me.TxtB3_N.Text = ""
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(136, 40)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(40, 16)
        Me.Label19.TabIndex = 10
        Me.Label19.Text = "Length"
        '
        'Label20
        '
        Me.Label20.Location = New System.Drawing.Point(136, 16)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(32, 16)
        Me.Label20.TabIndex = 9
        Me.Label20.Text = "Num"
        '
        'cmbB3
        '
        Me.cmbB3.Location = New System.Drawing.Point(8, 16)
        Me.cmbB3.Name = "cmbB3"
        Me.cmbB3.Size = New System.Drawing.Size(120, 21)
        Me.cmbB3.TabIndex = 0
        Me.cmbB3.Text = "Select"
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.TxtB4_L)
        Me.GroupBox6.Controls.Add(Me.TxtB4_N)
        Me.GroupBox6.Controls.Add(Me.Label21)
        Me.GroupBox6.Controls.Add(Me.Label22)
        Me.GroupBox6.Controls.Add(Me.cmbB4)
        Me.GroupBox6.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox6.Location = New System.Drawing.Point(264, 376)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(232, 72)
        Me.GroupBox6.TabIndex = 13
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Belt-4"
        '
        'TxtB4_L
        '
        Me.TxtB4_L.Location = New System.Drawing.Point(176, 38)
        Me.TxtB4_L.Name = "TxtB4_L"
        Me.TxtB4_L.Size = New System.Drawing.Size(48, 20)
        Me.TxtB4_L.TabIndex = 2
        Me.TxtB4_L.Text = ""
        '
        'TxtB4_N
        '
        Me.TxtB4_N.Location = New System.Drawing.Point(176, 14)
        Me.TxtB4_N.Name = "TxtB4_N"
        Me.TxtB4_N.Size = New System.Drawing.Size(48, 20)
        Me.TxtB4_N.TabIndex = 1
        Me.TxtB4_N.Text = ""
        '
        'Label21
        '
        Me.Label21.Location = New System.Drawing.Point(136, 40)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(40, 16)
        Me.Label21.TabIndex = 10
        Me.Label21.Text = "Length"
        '
        'Label22
        '
        Me.Label22.Location = New System.Drawing.Point(136, 16)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(32, 16)
        Me.Label22.TabIndex = 9
        Me.Label22.Text = "Num"
        '
        'cmbB4
        '
        Me.cmbB4.Location = New System.Drawing.Point(8, 16)
        Me.cmbB4.Name = "cmbB4"
        Me.cmbB4.Size = New System.Drawing.Size(120, 21)
        Me.cmbB4.TabIndex = 0
        Me.cmbB4.Text = "Select"
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.TxtIN_L)
        Me.GroupBox7.Controls.Add(Me.TxtIN_N)
        Me.GroupBox7.Controls.Add(Me.Label9)
        Me.GroupBox7.Controls.Add(Me.Label10)
        Me.GroupBox7.Controls.Add(Me.cmbIN)
        Me.GroupBox7.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox7.Location = New System.Drawing.Point(16, 192)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(232, 72)
        Me.GroupBox7.TabIndex = 5
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "INNERLINER"
        '
        'TxtIN_L
        '
        Me.TxtIN_L.Location = New System.Drawing.Point(176, 38)
        Me.TxtIN_L.Name = "TxtIN_L"
        Me.TxtIN_L.Size = New System.Drawing.Size(48, 20)
        Me.TxtIN_L.TabIndex = 2
        Me.TxtIN_L.Text = ""
        '
        'TxtIN_N
        '
        Me.TxtIN_N.Location = New System.Drawing.Point(176, 14)
        Me.TxtIN_N.Name = "TxtIN_N"
        Me.TxtIN_N.Size = New System.Drawing.Size(48, 20)
        Me.TxtIN_N.TabIndex = 1
        Me.TxtIN_N.Text = ""
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(136, 40)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(40, 16)
        Me.Label9.TabIndex = 10
        Me.Label9.Text = "Length"
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(136, 16)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(32, 16)
        Me.Label10.TabIndex = 9
        Me.Label10.Text = "Num"
        '
        'cmbIN
        '
        Me.cmbIN.Location = New System.Drawing.Point(8, 16)
        Me.cmbIN.Name = "cmbIN"
        Me.cmbIN.Size = New System.Drawing.Size(120, 21)
        Me.cmbIN.TabIndex = 0
        Me.cmbIN.Text = "Select"
        '
        'GroupBox8
        '
        Me.GroupBox8.Controls.Add(Me.TxtSD_L)
        Me.GroupBox8.Controls.Add(Me.TxtSD_N)
        Me.GroupBox8.Controls.Add(Me.Label7)
        Me.GroupBox8.Controls.Add(Me.Label8)
        Me.GroupBox8.Controls.Add(Me.cmbSD)
        Me.GroupBox8.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox8.Location = New System.Drawing.Point(16, 112)
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.Size = New System.Drawing.Size(232, 72)
        Me.GroupBox8.TabIndex = 4
        Me.GroupBox8.TabStop = False
        Me.GroupBox8.Text = "Side"
        '
        'TxtSD_L
        '
        Me.TxtSD_L.Location = New System.Drawing.Point(176, 38)
        Me.TxtSD_L.Name = "TxtSD_L"
        Me.TxtSD_L.Size = New System.Drawing.Size(48, 20)
        Me.TxtSD_L.TabIndex = 2
        Me.TxtSD_L.Text = ""
        '
        'TxtSD_N
        '
        Me.TxtSD_N.Location = New System.Drawing.Point(176, 14)
        Me.TxtSD_N.Name = "TxtSD_N"
        Me.TxtSD_N.Size = New System.Drawing.Size(48, 20)
        Me.TxtSD_N.TabIndex = 1
        Me.TxtSD_N.Text = ""
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(136, 40)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(40, 16)
        Me.Label7.TabIndex = 10
        Me.Label7.Text = "Length"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(136, 16)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(32, 16)
        Me.Label8.TabIndex = 9
        Me.Label8.Text = "Num"
        '
        'cmbSD
        '
        Me.cmbSD.Location = New System.Drawing.Point(8, 16)
        Me.cmbSD.Name = "cmbSD"
        Me.cmbSD.Size = New System.Drawing.Size(120, 21)
        Me.cmbSD.TabIndex = 0
        Me.cmbSD.Text = "Select"
        '
        'TxtRev
        '
        Me.TxtRev.Location = New System.Drawing.Point(280, 16)
        Me.TxtRev.Name = "TxtRev"
        Me.TxtRev.Size = New System.Drawing.Size(40, 20)
        Me.TxtRev.TabIndex = 1
        Me.TxtRev.Text = ""
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(224, 18)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 16)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "Revision"
        '
        'lblError
        '
        Me.lblError.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblError.ForeColor = System.Drawing.Color.Red
        Me.lblError.Location = New System.Drawing.Point(192, 22)
        Me.lblError.Name = "lblError"
        Me.lblError.Size = New System.Drawing.Size(24, 8)
        Me.lblError.TabIndex = 23
        Me.lblError.Text = "***"
        Me.lblError.Visible = False
        '
        'TxtCode
        '
        Me.TxtCode.Location = New System.Drawing.Point(88, 16)
        Me.TxtCode.Name = "TxtCode"
        Me.TxtCode.TabIndex = 0
        Me.TxtCode.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 18)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 16)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "Green Tire"
        '
        'CheckBoxTire
        '
        Me.CheckBoxTire.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CheckBoxTire.Checked = True
        Me.CheckBoxTire.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxTire.Location = New System.Drawing.Point(616, 18)
        Me.CheckBoxTire.Name = "CheckBoxTire"
        Me.CheckBoxTire.Size = New System.Drawing.Size(112, 16)
        Me.CheckBoxTire.TabIndex = 20
        Me.CheckBoxTire.Text = "Final   GreenTire"
        '
        'GPTire
        '
        Me.GPTire.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GPTire.Location = New System.Drawing.Point(512, 144)
        Me.GPTire.Name = "GPTire"
        Me.GPTire.Size = New System.Drawing.Size(224, 304)
        Me.GPTire.TabIndex = 15
        Me.GPTire.TabStop = False
        Me.GPTire.Text = "Green Tire"
        Me.GPTire.Visible = False
        '
        'GroupBox10
        '
        Me.GroupBox10.Controls.Add(Me.TxtCU_L)
        Me.GroupBox10.Controls.Add(Me.TxtCu_N)
        Me.GroupBox10.Controls.Add(Me.Label4)
        Me.GroupBox10.Controls.Add(Me.Label1)
        Me.GroupBox10.Controls.Add(Me.cmbCU)
        Me.GroupBox10.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox10.Location = New System.Drawing.Point(16, 272)
        Me.GroupBox10.Name = "GroupBox10"
        Me.GroupBox10.Size = New System.Drawing.Size(232, 72)
        Me.GroupBox10.TabIndex = 6
        Me.GroupBox10.TabStop = False
        Me.GroupBox10.Text = "Cussion"
        '
        'TxtCU_L
        '
        Me.TxtCU_L.Location = New System.Drawing.Point(176, 40)
        Me.TxtCU_L.Name = "TxtCU_L"
        Me.TxtCU_L.Size = New System.Drawing.Size(48, 20)
        Me.TxtCU_L.TabIndex = 2
        Me.TxtCU_L.Text = ""
        '
        'TxtCu_N
        '
        Me.TxtCu_N.Location = New System.Drawing.Point(176, 16)
        Me.TxtCu_N.Name = "TxtCu_N"
        Me.TxtCu_N.Size = New System.Drawing.Size(48, 20)
        Me.TxtCu_N.TabIndex = 1
        Me.TxtCu_N.Text = ""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(136, 42)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(40, 16)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Length"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(136, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Num"
        '
        'cmbCU
        '
        Me.cmbCU.Location = New System.Drawing.Point(8, 16)
        Me.cmbCU.Name = "cmbCU"
        Me.cmbCU.Size = New System.Drawing.Size(120, 21)
        Me.cmbCU.TabIndex = 0
        Me.cmbCU.Text = "Select"
        '
        'GroupBox11
        '
        Me.GroupBox11.Controls.Add(Me.cmbTT)
        Me.GroupBox11.Controls.Add(Me.TxtTT_N)
        Me.GroupBox11.Controls.Add(Me.Label12)
        Me.GroupBox11.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox11.Location = New System.Drawing.Point(16, 352)
        Me.GroupBox11.Name = "GroupBox11"
        Me.GroupBox11.Size = New System.Drawing.Size(232, 48)
        Me.GroupBox11.TabIndex = 7
        Me.GroupBox11.TabStop = False
        Me.GroupBox11.Text = "Tread"
        '
        'cmbTT
        '
        Me.cmbTT.Location = New System.Drawing.Point(8, 14)
        Me.cmbTT.Name = "cmbTT"
        Me.cmbTT.Size = New System.Drawing.Size(120, 21)
        Me.cmbTT.TabIndex = 0
        Me.cmbTT.Text = "Select"
        '
        'TxtTT_N
        '
        Me.TxtTT_N.Location = New System.Drawing.Point(176, 14)
        Me.TxtTT_N.Name = "TxtTT_N"
        Me.TxtTT_N.Size = New System.Drawing.Size(48, 20)
        Me.TxtTT_N.TabIndex = 1
        Me.TxtTT_N.Text = ""
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(136, 16)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(32, 16)
        Me.Label12.TabIndex = 9
        Me.Label12.Text = "Num"
        '
        'GroupBox12
        '
        Me.GroupBox12.Controls.Add(Me.TxtNy_L)
        Me.GroupBox12.Controls.Add(Me.TxtNy_N)
        Me.GroupBox12.Controls.Add(Me.Label25)
        Me.GroupBox12.Controls.Add(Me.Label26)
        Me.GroupBox12.Controls.Add(Me.CheckBoxNy)
        Me.GroupBox12.Controls.Add(Me.cmbNy)
        Me.GroupBox12.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox12.Location = New System.Drawing.Point(512, 56)
        Me.GroupBox12.Name = "GroupBox12"
        Me.GroupBox12.Size = New System.Drawing.Size(224, 72)
        Me.GroupBox12.TabIndex = 14
        Me.GroupBox12.TabStop = False
        Me.GroupBox12.Text = "Nylon Chafer"
        '
        'TxtNy_L
        '
        Me.TxtNy_L.Location = New System.Drawing.Point(168, 38)
        Me.TxtNy_L.Name = "TxtNy_L"
        Me.TxtNy_L.Size = New System.Drawing.Size(48, 20)
        Me.TxtNy_L.TabIndex = 2
        Me.TxtNy_L.Text = ""
        '
        'TxtNy_N
        '
        Me.TxtNy_N.Location = New System.Drawing.Point(168, 14)
        Me.TxtNy_N.Name = "TxtNy_N"
        Me.TxtNy_N.Size = New System.Drawing.Size(48, 20)
        Me.TxtNy_N.TabIndex = 1
        Me.TxtNy_N.Text = ""
        '
        'Label25
        '
        Me.Label25.Location = New System.Drawing.Point(128, 40)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(40, 16)
        Me.Label25.TabIndex = 10
        Me.Label25.Text = "Length"
        '
        'Label26
        '
        Me.Label26.Location = New System.Drawing.Point(128, 16)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(32, 16)
        Me.Label26.TabIndex = 9
        Me.Label26.Text = "Num"
        '
        'CheckBoxNy
        '
        Me.CheckBoxNy.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CheckBoxNy.Location = New System.Drawing.Point(8, 16)
        Me.CheckBoxNy.Name = "CheckBoxNy"
        Me.CheckBoxNy.Size = New System.Drawing.Size(56, 16)
        Me.CheckBoxNy.TabIndex = 0
        Me.CheckBoxNy.Text = "use"
        '
        'cmbNy
        '
        Me.cmbNy.Enabled = False
        Me.cmbNy.Location = New System.Drawing.Point(8, 40)
        Me.cmbNy.Name = "cmbNy"
        Me.cmbNy.Size = New System.Drawing.Size(120, 21)
        Me.cmbNy.TabIndex = 0
        Me.cmbNy.Text = "Select"
        '
        'GroupBox9
        '
        Me.GroupBox9.Controls.Add(Me.cmbBp)
        Me.GroupBox9.Controls.Add(Me.TxtBp_L)
        Me.GroupBox9.Controls.Add(Me.TxtBP_N)
        Me.GroupBox9.Controls.Add(Me.Label13)
        Me.GroupBox9.Controls.Add(Me.Label14)
        Me.GroupBox9.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox9.Location = New System.Drawing.Point(16, 408)
        Me.GroupBox9.Name = "GroupBox9"
        Me.GroupBox9.Size = New System.Drawing.Size(232, 72)
        Me.GroupBox9.TabIndex = 8
        Me.GroupBox9.TabStop = False
        Me.GroupBox9.Text = "Body Ply"
        '
        'cmbBp
        '
        Me.cmbBp.Location = New System.Drawing.Point(8, 16)
        Me.cmbBp.Name = "cmbBp"
        Me.cmbBp.Size = New System.Drawing.Size(120, 21)
        Me.cmbBp.TabIndex = 0
        Me.cmbBp.Text = "Select"
        '
        'TxtBp_L
        '
        Me.TxtBp_L.Location = New System.Drawing.Point(176, 38)
        Me.TxtBp_L.Name = "TxtBp_L"
        Me.TxtBp_L.Size = New System.Drawing.Size(48, 20)
        Me.TxtBp_L.TabIndex = 2
        Me.TxtBp_L.Text = ""
        '
        'TxtBP_N
        '
        Me.TxtBP_N.Location = New System.Drawing.Point(176, 14)
        Me.TxtBP_N.Name = "TxtBP_N"
        Me.TxtBP_N.Size = New System.Drawing.Size(48, 20)
        Me.TxtBP_N.TabIndex = 1
        Me.TxtBP_N.Text = ""
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(136, 40)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(40, 16)
        Me.Label13.TabIndex = 10
        Me.Label13.Text = "Length"
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(136, 16)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(32, 16)
        Me.Label14.TabIndex = 9
        Me.Label14.Text = "Num"
        '
        'CmdClose
        '
        Me.CmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdClose.Image = CType(resources.GetObject("CmdClose.Image"), System.Drawing.Image)
        Me.CmdClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdClose.Location = New System.Drawing.Point(656, 456)
        Me.CmdClose.Name = "CmdClose"
        Me.CmdClose.Size = New System.Drawing.Size(80, 56)
        Me.CmdClose.TabIndex = 18
        Me.CmdClose.Text = "Close"
        Me.CmdClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdSave
        '
        Me.CmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdSave.Image = CType(resources.GetObject("CmdSave.Image"), System.Drawing.Image)
        Me.CmdSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdSave.Location = New System.Drawing.Point(576, 456)
        Me.CmdSave.Name = "CmdSave"
        Me.CmdSave.Size = New System.Drawing.Size(80, 56)
        Me.CmdSave.TabIndex = 17
        Me.CmdSave.Text = "Save"
        Me.CmdSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CheckAll
        '
        Me.CheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CheckAll.Location = New System.Drawing.Point(512, 496)
        Me.CheckAll.Name = "CheckAll"
        Me.CheckAll.Size = New System.Drawing.Size(56, 16)
        Me.CheckAll.TabIndex = 19
        Me.CheckAll.Text = "Add"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(336, 18)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(56, 16)
        Me.Label5.TabIndex = 27
        Me.Label5.Text = "BSJ Code"
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(288, 458)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(96, 16)
        Me.Label11.TabIndex = 29
        Me.Label11.Text = "Revision of Boss"
        '
        'txtremark
        '
        Me.txtremark.Location = New System.Drawing.Point(408, 456)
        Me.txtremark.Name = "txtremark"
        Me.txtremark.Size = New System.Drawing.Size(48, 20)
        Me.txtremark.TabIndex = 15
        Me.txtremark.Text = ""
        '
        'CmbBSJCode
        '
        Me.CmbBSJCode.Location = New System.Drawing.Point(400, 16)
        Me.CmbBSJCode.Name = "CmbBSJCode"
        Me.CmbBSJCode.Size = New System.Drawing.Size(121, 21)
        Me.CmbBSJCode.TabIndex = 2
        Me.CmbBSJCode.Text = "Select"
        '
        'txtremark2
        '
        Me.txtremark2.Location = New System.Drawing.Point(504, 456)
        Me.txtremark2.Name = "txtremark2"
        Me.txtremark2.Size = New System.Drawing.Size(48, 20)
        Me.txtremark2.TabIndex = 16
        Me.txtremark2.Text = ""
        '
        'Label27
        '
        Me.Label27.Location = New System.Drawing.Point(464, 458)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(32, 16)
        Me.Label27.TabIndex = 32
        Me.Label27.Text = "2 ND"
        '
        'Label28
        '
        Me.Label28.Location = New System.Drawing.Point(376, 458)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(32, 16)
        Me.Label28.TabIndex = 33
        Me.Label28.Text = "1 ST"
        '
        'FrmAddGreenTire
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(752, 518)
        Me.Controls.Add(Me.Label28)
        Me.Controls.Add(Me.Label27)
        Me.Controls.Add(Me.txtremark2)
        Me.Controls.Add(Me.CmbBSJCode)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.txtremark)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.CmdClose)
        Me.Controls.Add(Me.CmdSave)
        Me.Controls.Add(Me.GroupBox11)
        Me.Controls.Add(Me.GPTire)
        Me.Controls.Add(Me.CheckBoxTire)
        Me.Controls.Add(Me.TxtRev)
        Me.Controls.Add(Me.TxtCode)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.lblError)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.GroupBox8)
        Me.Controls.Add(Me.GroupBox7)
        Me.Controls.Add(Me.GroupBox6)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.GroupBox10)
        Me.Controls.Add(Me.GroupBox12)
        Me.Controls.Add(Me.GroupBox9)
        Me.Controls.Add(Me.CheckAll)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmAddGreenTire"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add GreenTire"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox8.ResumeLayout(False)
        Me.GroupBox10.ResumeLayout(False)
        Me.GroupBox11.ResumeLayout(False)
        Me.GroupBox12.ResumeLayout(False)
        Me.GroupBox9.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "CONSTANT"
    Dim DT As New DataTable
    Dim StrSQL As String
    Dim oldrow As Integer
    Dim C1 As New SQLData("ACCINV")
    Dim QTT, QBF, QBp, QCU, QSD, QIN, QB1, QB2, QB3, QB4, QWf, QNY As Double
#End Region

#Region "CMBBOX"
    Sub LoadBSJ()
        Dim dt As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = " SELECT  *  FROM  TblTiresize "
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dt = New DataTable
            DA.Fill(dt)
        Catch
        Finally
        End Try
        dt.TableName = TBL_BSJ
        GrdDVBSJ = dt.DefaultView
        '************************************
        CmbBSJCode.DisplayMember = "BSJCode"
        CmbBSJCode.ValueMember = "BSJCode"
        CmbBSJCode.DataSource = dt
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadCU()
        Dim dtCU As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = "   SELECT Final,Round(QPU,4) QPU  FROM  TblSemi  "
        StrSQL &= "  where MaterialType = '03' "
        StrSQL &= "  and active = '1' "
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtCU = New DataTable
            DA.Fill(dtCU)
        Catch
        Finally
        End Try
        dtCU.TableName = TBL_CU
        GrdDVCU = dtCU.DefaultView
        '************************************
        cmbCU.DisplayMember = "Final"
        cmbCU.ValueMember = "QPU"
        cmbCU.DataSource = dtCU
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadSD()
        Dim dt As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = "   SELECT Final,Round(QPU,4) QPU  FROM  TblSemi  "
        StrSQL &= "  where MaterialType = '11' "
        StrSQL &= "  and active = '1' "
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dt = New DataTable
            DA.Fill(dt)
        Catch
        Finally
        End Try
        dt.TableName = TBL_SD
        GrdDVSD = dt.DefaultView
        '************************************
        cmbSD.DisplayMember = "Final"
        cmbSD.ValueMember = "QPU"
        cmbSD.DataSource = dt
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadBF()
        Dim dtBF As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = "   SELECT Final,Round(QPU,4) QPU  FROM  TblSemi  "
        StrSQL &= "  where MaterialType = '14' "
        StrSQL &= "  and active = '1' "
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtBF = New DataTable
            DA.Fill(dtBF)
        Catch
        Finally
        End Try
        dtBF.TableName = TBL_BF
        GrdDVBF = dtBF.DefaultView
        '************************************
        cmbBF.DisplayMember = "Final"
        cmbBF.ValueMember = "QPU"
        cmbBF.DataSource = dtBF
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadIN()
        Dim dt As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = "   SELECT Final,Round(QPU,4) QPU  FROM  TblSemi  "
        StrSQL &= "  where MaterialType = '12' "
        StrSQL &= "  and active = '1' "
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dt = New DataTable
            DA.Fill(dt)
        Catch
        Finally
        End Try
        dt.TableName = TBL_IN
        GrdDVIN = dt.DefaultView
        '************************************
        cmbIN.DisplayMember = "Final"
        cmbIN.ValueMember = "QPU"
        cmbIN.DataSource = dt
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadTT()
        Dim dt As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = "   SELECT Final,Round(QPU,4) QPU  FROM  TblSemi  "
        StrSQL &= "  where MaterialType = '13' "
        StrSQL &= "  and active = '1' "
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dt = New DataTable
            DA.Fill(dt)
        Catch
        Finally
        End Try
        dt.TableName = TBL_TT
        GrdDVTT = dt.DefaultView
        '************************************
        cmbTT.DisplayMember = "Final"
        cmbTT.ValueMember = "QPU"
        cmbTT.DataSource = dt
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadWf()
        Dim dt As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = "   SELECT Final,Round(QPU,4) QPU  FROM  TblSemi  "
        StrSQL &= "  where MaterialType = '09' "
        StrSQL &= "  and active = '1' "
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dt = New DataTable
            DA.Fill(dt)
        Catch
        Finally
        End Try
        dt.TableName = TBL_Wf
        GrdDVWf = dt.DefaultView
        '************************************
        cmbWf.DisplayMember = "Final"
        cmbWf.ValueMember = "QPU"
        cmbWf.DataSource = dt
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadNy()
        Dim dt As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = "   SELECT Final,Round(QPU,4) QPU  FROM  TblSemi  "
        StrSQL &= "  where MaterialType = '10' "
        StrSQL &= "  and active = '1' "
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dt = New DataTable
            DA.Fill(dt)
        Catch
        Finally
        End Try
        dt.TableName = TBL_Ny
        GrdDVNy = dt.DefaultView
        '************************************
        cmbNy.DisplayMember = "Final"
        cmbNy.ValueMember = "QPU"
        cmbNy.DataSource = dt
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadB1()
        Dim dt As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = "   SELECT Final,Round(QPU,4) QPU  FROM  TblSemi  "
        StrSQL &= "  where MaterialType = '05' "
        StrSQL &= "  and active = '1' "
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dt = New DataTable
            DA.Fill(dt)
        Catch
        Finally
        End Try
        dt.TableName = TBL_B1
        GrdDVB1 = dt.DefaultView
        '************************************
        cmbB1.DisplayMember = "Final"
        cmbB1.ValueMember = "QPU"
        cmbB1.DataSource = dt
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadB2()
        Dim dt As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = "   SELECT Final,Round(QPU,4) QPU  FROM  TblSemi  "
        StrSQL &= "  where MaterialType = '06' "
        StrSQL &= "  and active = '1' "
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dt = New DataTable
            DA.Fill(dt)
        Catch
        Finally
        End Try
        dt.TableName = TBL_B2
        GrdDVB2 = dt.DefaultView
        '************************************
        cmbB2.DisplayMember = "Final"
        cmbB2.ValueMember = "QPU"
        cmbB2.DataSource = dt
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadB3()
        Dim dt As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = "   SELECT Final,Round(QPU,4) QPU  FROM  TblSemi  "
        StrSQL &= "  where MaterialType = '07' "
        StrSQL &= "  and active = '1' "
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dt = New DataTable
            DA.Fill(dt)
        Catch
        Finally
        End Try
        dt.TableName = TBL_B3
        GrdDVB3 = dt.DefaultView
        '************************************
        cmbB3.DisplayMember = "Final"
        cmbB3.ValueMember = "QPU"
        cmbB3.DataSource = dt
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadB4()
        Dim dt As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = "   SELECT Final,Round(QPU,4) QPU  FROM  TblSemi  "
        StrSQL &= "  where MaterialType = '08' "
        StrSQL &= "  and active = '1' "
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dt = New DataTable
            DA.Fill(dt)
        Catch
        Finally
        End Try
        dt.TableName = TBL_B4
        GrdDVB4 = dt.DefaultView
        '************************************
        cmbB4.DisplayMember = "Final"
        cmbB4.ValueMember = "QPU"
        cmbB4.DataSource = dt
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadBp()
        Dim dt As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = "   SELECT Final,Round(QPU,4) QPU  FROM  TblSemi  "
        StrSQL &= "  where MaterialType = '04' "
        StrSQL &= "  and active = '1' "
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dt = New DataTable
            DA.Fill(dt)
        Catch
        Finally
        End Try
        dt.TableName = TBL_Bp
        GrdDVBp = dt.DefaultView
        '************************************
        cmbBp.DisplayMember = "Final"
        cmbBp.ValueMember = "QPU"
        cmbBp.DataSource = dt
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
#End Region

    Private Sub FrmAddGreenTire_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadBSJ()
        LoadSD()
        LoadCU()
        LoadBF()
        LoadIN()
        LoadTT()
        LoadB1()
        LoadB2()
        LoadB3()
        LoadB4()
        LoadBp()
        chkWf()
        chkNy()
        If CmdSave.Text = "Save" Then
            SDcode = cmbSD.Text.Trim
            CUcode = cmbCU.Text.Trim
            BFcode = cmbBF.Text.Trim
            INcode = cmbIN.Text.Trim
            TTcode = cmbTT.Text.Trim
            B1code = cmbB1.Text.Trim
            B2code = cmbB2.Text.Trim
            B3code = cmbB3.Text.Trim
            B4code = cmbB4.Text.Trim
            Bpcode = cmbBp.Text.Trim
            Wfcode = cmbWf.Text.Trim
            NYcode = cmbNy.Text.Trim
        ElseIf CmdSave.Text = "Edit" Then
            TxtCode.Enabled = True
            cmbBF.Text = BFcode
            cmbSD.Text = SDcode
            cmbCU.Text = CUcode
            cmbIN.Text = INcode
            cmbTT.Text = TTcode
            cmbB1.Text = B1code
            cmbB2.Text = B2code
            cmbB3.Text = B3code
            cmbB4.Text = B4code
            cmbBp.Text = Bpcode
            cmbWf.Text = Wfcode
            cmbNy.Text = NYcode
            CmbBSJCode.Text = BSJCode
        End If
    End Sub

    Private Sub ChkListCU_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub CheckBoxWf_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxWf.CheckedChanged
        chkWf()
    End Sub

    Private Sub CheckBoxNy_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxNy.CheckedChanged
        chkNy()
    End Sub
    Sub chkWf()
        If CheckBoxWf.Checked Then
            cmbWf.Enabled = True
            LoadWf()
        Else
            cmbWf.Enabled = False
            cmbWf.Text = "Select"
        End If
    End Sub
    Sub chkNy()
        If CheckBoxNy.Checked Then
            cmbNy.Enabled = True
            LoadNy()
        Else
            cmbNy.Enabled = False
            cmbNy.Text = "Select"
        End If
    End Sub

#Region "checktxt"
    Sub checktxt()
        If txtremark.Text = "" Then
            MsgBox("Please check data again.", MsgBoxStyle.Exclamation)
            txtremark.Focus()
            Exit Sub
        End If
        If txtremark2.Text = "" Then
            MsgBox("Please check data again.", MsgBoxStyle.Exclamation)
            txtremark2.Focus()
            Exit Sub
        End If
        If TxtCode.Text = "" Then
            MsgBox("Please check data again.", MsgBoxStyle.Exclamation)
            TxtCode.Focus()
            Exit Sub
        End If
        If TxtRev.Text = "" Then
            MsgBox("Please check data again.", MsgBoxStyle.Exclamation)
            TxtRev.Focus()
            Exit Sub
        End If

        If cmbBF.Text <> "select" Then
            BFcode = cmbBF.Text.Trim
            If txtBF_N.Text = "" Then
                MsgBox("Please check data again.", MsgBoxStyle.Exclamation)
                txtBF_N.Focus()
                QBF = 0
                Exit Sub
            Else
                QBF = cmbBF.SelectedValue * txtBF_N.Text.Trim
            End If
        End If
        If cmbSD.Text <> "select" Then
            SDcode = cmbSD.Text.Trim
            If TxtSD_N.Text = "" Or TxtSD_L.Text = "" Then
                MsgBox("Please check data again.", MsgBoxStyle.Exclamation)
                TxtSD_N.Focus()
                QSD = 0
                Exit Sub
            Else
                QSD = (cmbSD.SelectedValue * (TxtSD_L.Text.Trim / 1000))
            End If
        End If
        If cmbIN.Text <> "select" Then
            INcode = cmbIN.Text.Trim
            If TxtIN_N.Text = "" Or TxtIN_L.Text = "" Then
                MsgBox("Please check data again.", MsgBoxStyle.Exclamation)
                TxtIN_N.Focus()
                QIN = 0
                Exit Sub
            Else
                QIN = (cmbIN.SelectedValue * (TxtIN_L.Text.Trim / 1000))
            End If
        End If
        If cmbCU.Text <> "select" Then
            CUcode = cmbCU.Text.Trim
            If TxtCu_N.Text = "" Or TxtCU_L.Text = "" Then
                MsgBox("Please check data again.", MsgBoxStyle.Exclamation)
                TxtCu_N.Focus()
                QCU = 0
                Exit Sub
            Else
                QCU = (cmbCU.SelectedValue * (TxtCU_L.Text.Trim / 1000))
            End If
        End If
        If cmbTT.Text <> "select" Then
            TTcode = cmbTT.Text.Trim
            If TxtTT_N.Text = "" Then
                MsgBox("Please check data again.", MsgBoxStyle.Exclamation)
                TxtTT_N.Focus()
                QTT = 0
                Exit Sub
            Else
                QTT = cmbTT.SelectedValue
            End If
        End If
        If cmbBp.Text <> "select" Then
            Bpcode = cmbBp.Text.Trim
            If TxtBP_N.Text = "" Or TxtBp_L.Text = "" Then
                MsgBox("Please check data again.", MsgBoxStyle.Exclamation)
                TxtBP_N.Focus()
                QBp = 0
                Exit Sub
            Else
                QBp = (cmbBp.SelectedValue * (TxtBp_L.Text.Trim / 1000))
            End If
        End If
        If cmbB1.Text <> "select" Then
            B1code = cmbB1.Text.Trim
            If TxtB1_N.Text = "" Or TxtB1_L.Text = "" Then
                MsgBox("Please check data again.", MsgBoxStyle.Exclamation)
                TxtB1_N.Focus()
                QB1 = 0
                Exit Sub
            Else
                QB1 = (cmbB1.SelectedValue * (TxtB1_L.Text.Trim / 1000))
            End If
        End If
        If cmbB2.Text <> "select" Then
            B2code = cmbB2.Text.Trim
            If TxtB2_N.Text = "" Or TxtB2_L.Text = "" Then
                MsgBox("Please check data again.", MsgBoxStyle.Exclamation)
                TxtB2_N.Focus()
                QB2 = 0
                Exit Sub
            Else
                QB2 = (cmbB2.SelectedValue * (TxtB2_L.Text.Trim / 1000))
            End If
        End If
        If cmbB3.Text <> "select" Then
            B3code = cmbB3.Text.Trim
            If TxtB3_N.Text = "" Or TxtB3_L.Text = "" Then
                MsgBox("Please check data again.", MsgBoxStyle.Exclamation)
                TxtB3_N.Focus()
                QB3 = 0
                Exit Sub
            Else
                QB3 = (cmbB3.SelectedValue * (TxtB3_L.Text.Trim / 1000))
            End If
        End If
        If cmbB4.Text <> "select" Then
            B4code = cmbB4.Text.Trim
            If TxtB4_N.Text = "" Or TxtB4_L.Text = "" Then
                MsgBox("Please check data again.", MsgBoxStyle.Exclamation)
                TxtB4_N.Focus()
                QB4 = 0
                Exit Sub
            Else
                QB4 = (cmbB4.SelectedValue * (TxtB4_L.Text.Trim / 1000))
            End If
        End If
        If CheckBoxWf.Checked Then
            Wfcode = cmbWf.Text.Trim
            If TxtWf_N.Text = "" Or TxtWf_L.Text = "" Then
                MsgBox("Please check data again.", MsgBoxStyle.Exclamation)
                TxtWf_N.Focus()
                QWf = 0
                Exit Sub
            Else
                QWf = (cmbWf.SelectedValue * (TxtWf_L.Text.Trim / 1000))
            End If
        Else
            Wfcode = ""
            QWf = 0
        End If
        If CheckBoxNy.Checked Then
            NYcode = cmbNy.Text.Trim
            If TxtNy_N.Text = "" Or TxtNy_L.Text = "" Then
                MsgBox("Please check data again.", MsgBoxStyle.Exclamation)
                TxtNy_N.Focus()
                QNY = 0
                Exit Sub
            Else
                QNY = (cmbNy.SelectedValue * (TxtNy_L.Text.Trim / 1000))
            End If
        Else
            NYcode = ""
            QNY = 0
        End If
    End Sub
#End Region

#Region "CMB"
    Private Sub cmbBF_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbBF.SelectedIndexChanged
        txtBF_N.Focus()
    End Sub
    Private Sub cmbSD_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSD.SelectedIndexChanged
        TxtSD_N.Focus()
    End Sub

    Private Sub cmbIN_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbIN.SelectedIndexChanged
        TxtIN_N.Focus()
    End Sub

    Private Sub cmbCU_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCU.SelectedIndexChanged
        TxtCu_N.Focus()
    End Sub

    Private Sub cmbTT_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbTT.SelectedIndexChanged
        TxtTT_N.Focus()
    End Sub

    Private Sub cmbBp_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbBp.SelectedIndexChanged
        TxtBP_N.Focus()
    End Sub

    Private Sub cmbWf_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbWf.SelectedIndexChanged
        TxtWf_N.Focus()
    End Sub

    Private Sub cmbB1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbB1.SelectedIndexChanged
        TxtB1_N.Focus()
    End Sub

    Private Sub cmbB2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbB2.SelectedIndexChanged
        TxtB2_N.Focus()
    End Sub

    Private Sub cmbB3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbB3.SelectedIndexChanged
        TxtB3_N.Focus()
    End Sub

    Private Sub cmbB4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbB4.SelectedIndexChanged
        TxtB4_N.Focus()
    End Sub

    Private Sub cmbNy_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbNy.SelectedIndexChanged
        TxtNy_N.Focus()
    End Sub

#End Region

    Private Sub CmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClose.Click
        Me.Close()
    End Sub

    Private Sub CmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSave.Click
        checktxt()

        iTotal = CSng(QTT + QBF + QBp + QCU + QSD + QIN + QB1 + QB2 + QB3 + QB4 + QWf + QNY)

        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        msg = "Green Tire Code : " & TxtCode.Text.Trim & " Weight Total : " & Format(CSng(iTotal / 1000), "##,##0.0000") & "KG." ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Add Green Tire"    ' Define title.
        If CmdSave.Text = "Save" Then
            If ChkData() = True Then
                MsgBox("It's Duplicate.", MsgBoxStyle.Critical, "Deprtment")
                Exit Sub
            Else
            End If
        Else
        End If

        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            If CmdSave.Text = "Save" Then
                If Tire_Hdr() Then
                    If Tire_Dtl() Then
                        UpTire()
                        If CheckAll.Checked Then
                            clear()
                        Else
                            Me.Close()
                        End If
                    End If
                End If
            Else
                If ChkEditData() Then
                    If Tire_Hdr() Then
                        If Tire_Dtl() Then
                            UpTire()
                            If CheckAll.Checked Then
                                clear()
                            Else
                                Me.Close()
                            End If
                        End If
                    End If
                Else
                    If UPTire_Hdr() Then
                        If UPTire_Dtl() Then
                            UpTire()
                            Me.Close()
                        Else
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            End If
        Else
            Exit Sub
        End If

    End Sub

#Region "KeyPress"
    Private Sub Txt_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCode.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                TxtCode.Text = TxtCode.Text.ToUpper
                Dim i As Integer
                If CmdSave.Text = "Save" Then
                    i = iNo() + 1
                    TxtRev.Text = Format(i, "000")

                    GrdDVBSJ.RowFilter = " tirecode Like  '%" & TxtCode.Text.Trim & "%'"
                    CmbBSJCode.DataSource = GrdDVBSJ
                Else
                End If
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case Else
        End Select
    End Sub
    Private Sub TxtM_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtB1_N.KeyPress _
   , TxtB2_N.KeyPress, TxtB3_N.KeyPress, TxtB4_N.KeyPress, txtBF_N.KeyPress, TxtBP_N.KeyPress, TxtCu_N.KeyPress, TxtIN_N.KeyPress _
   , TxtNy_N.KeyPress, TxtWf_N.KeyPress, TxtSD_N.KeyPress, TxtTT_N.KeyPress, TxtRev.KeyPress
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
    Private Sub TxtL_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtB1_L.KeyPress _
        , TxtB2_L.KeyPress, TxtB3_L.KeyPress, TxtB4_L.KeyPress, TxtBp_L.KeyPress, TxtCU_L.KeyPress, TxtIN_L.KeyPress _
        , TxtNy_L.KeyPress, TxtWf_L.KeyPress, TxtSD_L.KeyPress, txtremark.KeyPress, txtremark2.KeyPress
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
#End Region

#Region "RM"
    Private Function ChkGP() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Try
            strSQL = " Select count(*) from TBLGroup "
            strSQL += " Where Code = " & PrepareStr(TxtCode.Text.Trim)
            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar
            If i = 0 Then
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
    Private Function ChkEditData() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Try
            strSQL = " Select count(*) from TBLGTHdr "
            strSQL += " Where TireCode = " & PrepareStr(TxtCode.Text.Trim)
            strSQL += " and Rev = " & PrepareStr(TxtRev.Text.Trim)
            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar
            If i = 0 Then
                ChkEditData = True
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
            strSQL = " Select count(*) from TBLGTHdr "
            strSQL += " Where TireCode = " & PrepareStr(TxtCode.Text.Trim)
            strSQL += " and Rev = " & PrepareStr(TxtRev.Text.Trim)
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
    Private Function iNo() As Long
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim drSQL As SqlDataReader
        Dim strSQL As String
        Try

            ' Build Insert statement to insert 
            strSQL = "  SELECT   top 1 Rev "
            strSQL &= "  FROM   TBLGTHdr"
            strSQL &= " Where TireCode  = '" & TxtCode.Text.Trim & "'"
            strSQL &= "  order by Rev desc"

            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            cmSQL.ExecuteScalar()
            drSQL = cmSQL.ExecuteReader()
            If drSQL.HasRows Then
                If drSQL.Read() Then
                    iNo = CInt(drSQL.Item("Rev").ToString())
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
    Private Function Tire_Dtl() As Boolean
        Dim strsql As String = String.Empty
        Dim cn As New SqlConnection(C1.Strcon)
        Dim cmd As New SqlCommand(strsql, cn)
        cn.Open()
        Dim t1 As SqlTransaction = cn.BeginTransaction
        cmd.Transaction = t1
        Dim strDate, str() As String
        str = Split(Now.Date.ToShortDateString, "/")
        strDate = str(2) + str(1) + str(0)
        Dim i As Integer
        Try
            For i = 1 To 12
                strsql = "Insert TBLGTDtl"
                strsql += " Values(" & PrepareStr(TxtCode.Text.Trim)
                strsql += "," & PrepareStr(TxtRev.Text.Trim)
                If i = 1 Then
                    strsql += "," & PrepareStr("14")
                    strsql += "," & PrepareStr(BFcode)
                    strsql += "," & PrepareStr("")
                    strsql += "," & PrepareStr(txtBF_N.Text.Trim)
                    strsql += "," & PrepareStr(cmbBF.SelectedValue * txtBF_N.Text.Trim)
                ElseIf i = 2 Then
                    strsql += "," & PrepareStr("11")
                    strsql += "," & PrepareStr(SDcode)
                    strsql += "," & PrepareStr(TxtSD_L.Text.Trim)
                    strsql += "," & PrepareStr(TxtSD_N.Text.Trim)
                    strsql += "," & PrepareStr((cmbSD.SelectedValue * (TxtSD_L.Text.Trim / 1000)))
                ElseIf i = 3 Then
                    strsql += "," & PrepareStr("12")
                    strsql += "," & PrepareStr(INcode)
                    strsql += "," & PrepareStr(TxtIN_L.Text.Trim)
                    strsql += "," & PrepareStr(TxtIN_N.Text.Trim)
                    strsql += "," & PrepareStr((cmbIN.SelectedValue * (TxtIN_L.Text.Trim / 1000)))
                ElseIf i = 4 Then
                    strsql += "," & PrepareStr("03")
                    strsql += "," & PrepareStr(CUcode)
                    strsql += "," & PrepareStr(TxtCU_L.Text.Trim)
                    strsql += "," & PrepareStr(TxtCu_N.Text.Trim)
                    strsql += "," & PrepareStr((cmbCU.SelectedValue * (TxtCU_L.Text.Trim / 1000)))
                ElseIf i = 5 Then
                    strsql += "," & PrepareStr("13")
                    strsql += "," & PrepareStr(TTcode)
                    strsql += "," & PrepareStr("")
                    strsql += "," & PrepareStr(TxtTT_N.Text.Trim)
                    strsql += "," & PrepareStr((cmbTT.SelectedValue * TxtTT_N.Text.Trim))
                ElseIf i = 6 Then
                    strsql += "," & PrepareStr("04")
                    strsql += "," & PrepareStr(Bpcode)
                    strsql += "," & PrepareStr(TxtBp_L.Text.Trim)
                    strsql += "," & PrepareStr(TxtBP_N.Text.Trim)
                    strsql += "," & PrepareStr((cmbBp.SelectedValue * (TxtBp_L.Text.Trim / 1000)))
                ElseIf i = 7 Then
                    If CheckBoxWf.Checked Then
                        strsql += "," & PrepareStr("09")
                        strsql += "," & PrepareStr(Wfcode)
                        strsql += "," & PrepareStr(TxtWf_L.Text.Trim)
                        strsql += "," & PrepareStr(TxtWf_N.Text.Trim)
                        strsql += "," & PrepareStr((cmbWf.SelectedValue * (TxtWf_L.Text.Trim / 1000)))
                    Else
                        strsql += "," & PrepareStr("09")
                        strsql += "," & PrepareStr(Wfcode)
                        strsql += "," & PrepareStr(TxtWf_L.Text.Trim)
                        strsql += "," & PrepareStr(TxtWf_N.Text.Trim)
                        strsql += "," & PrepareStr("")
                    End If
                ElseIf i = 8 Then
                    strsql += "," & PrepareStr("05")
                    strsql += "," & PrepareStr(B1code)
                    strsql += "," & PrepareStr(TxtB1_L.Text.Trim)
                    strsql += "," & PrepareStr(TxtB1_N.Text.Trim)
                    strsql += "," & PrepareStr((cmbB1.SelectedValue * (TxtB1_L.Text.Trim / 1000)))
                ElseIf i = 9 Then
                    strsql += "," & PrepareStr("06")
                    strsql += "," & PrepareStr(B2code)
                    strsql += "," & PrepareStr(TxtB2_L.Text.Trim)
                    strsql += "," & PrepareStr(TxtB2_N.Text.Trim)
                    strsql += "," & PrepareStr((cmbB2.SelectedValue * (TxtB2_L.Text.Trim / 1000)))
                ElseIf i = 10 Then
                    strsql += "," & PrepareStr("07")
                    strsql += "," & PrepareStr(B3code)
                    strsql += "," & PrepareStr(TxtB3_L.Text.Trim)
                    strsql += "," & PrepareStr(TxtB3_N.Text.Trim)
                    strsql += "," & PrepareStr((cmbB3.SelectedValue * (TxtB3_L.Text.Trim / 1000)))
                ElseIf i = 11 Then
                    strsql += "," & PrepareStr("08")
                    strsql += "," & PrepareStr(B4code)
                    strsql += "," & PrepareStr(TxtB4_L.Text.Trim)
                    strsql += "," & PrepareStr(TxtB4_N.Text.Trim)
                    strsql += "," & PrepareStr((cmbB4.SelectedValue * (TxtB4_L.Text.Trim / 1000)))
                ElseIf i = 12 Then
                    If CheckBoxNy.Checked Then
                        strsql += "," & PrepareStr("10")
                        strsql += "," & PrepareStr(NYcode)
                        strsql += "," & PrepareStr(TxtNy_L.Text.Trim)
                        strsql += "," & PrepareStr(TxtNy_N.Text.Trim)
                        strsql += "," & PrepareStr((cmbNy.SelectedValue * (TxtNy_L.Text.Trim / 1000)))
                    Else
                        strsql += "," & PrepareStr("10")
                        strsql += "," & PrepareStr(NYcode)
                        strsql += "," & PrepareStr(TxtNy_L.Text.Trim)
                        strsql += "," & PrepareStr(TxtNy_N.Text.Trim)
                        strsql += "," & PrepareStr("")
                    End If
                End If
                strsql += "," & PrepareStr("g")
                strsql += "," & PrepareStr(strDate) & ")"
                cmd.CommandText = strsql
                cmd.ExecuteNonQuery()
            Next
            t1.Commit()
            Tire_Dtl = True
        Catch
            t1.Rollback()
            Tire_Dtl = False
            MsgBox("Rollback data")
            Exit Function
        Finally
            cn.Close()
        End Try
    End Function
    Private Function Tire_Hdr() As Boolean
        Dim strsql As String = String.Empty
        Dim cn As New SqlConnection(C1.Strcon)
        Dim cmd As New SqlCommand(strsql, cn)
        cn.Open()
        Dim t1 As SqlTransaction = cn.BeginTransaction
        cmd.Transaction = t1
        Dim strDate, str() As String
        str = Split(Now.Date.ToShortDateString, "/")
        strDate = str(2) + str(1) + str(0)
        Try
            strsql = " "

            If ChkGP() Then
                strsql += "Insert  TblGroup "
                strsql += " values ( '06',"
                strsql += PrepareStr(TxtCode.Text.Trim) & ")"
            End If

            strsql += " "
            strsql += "Insert  TBLGTHdr "
            strsql += " values (" & PrepareStr(TxtCode.Text.Trim) & ","
            strsql += PrepareStr(TxtCode.Text.Trim) & ","
            strsql += PrepareStr(TxtRev.Text.Trim) & ","
            strsql += PrepareStr(CmbBSJCode.Text.Trim) & ","
            strsql += PrepareStr(CSng(iTotal)) & ","
            If CheckBoxTire.Checked = True Then
                strsql += PrepareStr(1) & ","
            Else
                strsql += PrepareStr(0) & ","
            End If
            strsql += PrepareStr(strDate.Trim) & ","
            strsql += PrepareStr(txtremark.Text.Trim + "," + txtremark2.Text.Trim) & ")"

            strsql += ""
            strsql += " Insert TBLConvert "
            strsql += " Values('06'"
            strsql += "," & PrepareStr(TxtCode.Text.Trim)
            strsql += "," & PrepareStr(TxtCode.Text.Trim)
            strsql += "," & PrepareStr(TxtRev.Text.Trim)
            strsql += "," & PrepareStr("UT")
            strsql += "," & PrepareStr("KG")
            strsql += "," & PrepareStr(1)
            strsql += "," & PrepareStr(CSng(iTotal / 1000))
            strsql += ")"

            strsql += ""
            strsql += " Insert TBLConvert "
            strsql += " Values('06'"
            strsql += "," & PrepareStr(TxtCode.Text.Trim)
            strsql += "," & PrepareStr(TxtCode.Text.Trim)
            strsql += "," & PrepareStr(TxtRev.Text.Trim)
            strsql += "," & PrepareStr("KG")
            strsql += "," & PrepareStr("KG")
            strsql += "," & PrepareStr(1)
            strsql += "," & PrepareStr(1)
            strsql += ")"
            cmd.CommandText = strsql
            cmd.ExecuteNonQuery()
            t1.Commit()
            Tire_Hdr = True
        Catch
            Tire_Hdr = False
            t1.Rollback()
            MsgBox("Rollback data")
            Exit Function
        Finally
            cn.Close()
        End Try
    End Function
    Private Function UPTire_Hdr() As Boolean
        Dim strsql As String = String.Empty
        Dim cn As New SqlConnection(C1.Strcon)
        Dim cmd As New SqlCommand(strsql, cn)
        cn.Open()
        Dim t1 As SqlTransaction = cn.BeginTransaction
        cmd.Transaction = t1
        Dim strDate, str() As String
        str = Split(Now.Date.ToShortDateString, "/")
        strDate = str(2) + str(1) + str(0)
        If CmdSave.Text = "Edit" Then
            Try
                strsql = " "
                strsql += "Update  TBLGTHdr "
                strsql += " set Qty  = " & PrepareStr(CSng(iTotal)) & ","
                strsql += " Remark =  " & PrepareStr(txtremark.Text.Trim + "," + txtremark2.Text.Trim) & ","
                strsql += " Dateup =  " & PrepareStr(strDate.Trim)
                strsql += " Where Tirecode = " & PrepareStr(TxtCode.Text.Trim)
                strsql += " and Rev = " & PrepareStr(TxtRev.Text.Trim)

                strsql += ""
                strsql += " Update TBLConvert "
                strsql += " set SQty = " & PrepareStr(CSng(iTotal / 1000))
                strsql += " where code  = " & PrepareStr(TxtCode.Text.Trim)
                strsql += " and Rev = " & PrepareStr(TxtRev.Text.Trim)
                strsql += " and UnitBig = 'UT'"

                cmd.CommandText = strsql
                cmd.ExecuteNonQuery()
                t1.Commit()
                UPTire_Hdr = True
            Catch
                UPTire_Hdr = False
                t1.Rollback()
                MsgBox("Rollback data")
                Exit Function
            Finally
                cn.Close()
            End Try
        End If
    End Function
    Private Function UPTire_Dtl() As Boolean
        Dim strsql As String = String.Empty
        Dim cn As New SqlConnection(C1.Strcon)
        Dim cmd As New SqlCommand(strsql, cn)
        cn.Open()
        Dim t1 As SqlTransaction = cn.BeginTransaction
        cmd.Transaction = t1
        Dim strDate, str() As String
        str = Split(Now.Date.ToShortDateString, "/")
        strDate = str(2) + str(1) + str(0)
        If CmdSave.Text = "Edit" Then
            Dim i As Integer
            Try
                For i = 1 To 12
                    If i = 1 Then
                        strsql = "Update TBLGTDtl"
                        strsql += " set number = " & PrepareStr(txtBF_N.Text.Trim)
                        strsql += " , QTU = " & PrepareStr(cmbBF.SelectedValue * txtBF_N.Text.Trim)
                        strsql += " , Dateup =  " & PrepareStr(strDate)
                        strsql += " ,  semicode = " & PrepareStr(BFcode)
                        strsql += " where tirecode = " & PrepareStr(TxtCode.Text.Trim)
                        strsql += " and Rev = " & PrepareStr(TxtRev.Text.Trim)
                        strsql += " and MaterialType = " & PrepareStr("14")
                    ElseIf i = 2 Then
                        strsql = "Update TBLGTDtl"
                        strsql += " set Length = " & PrepareStr(TxtSD_L.Text.Trim)
                        strsql += " , number = " & PrepareStr(TxtSD_N.Text.Trim)
                        strsql += " , QTU = " & PrepareStr((cmbSD.SelectedValue * (TxtSD_L.Text.Trim / 1000)))
                        strsql += " , Dateup =  " & PrepareStr(strDate)
                        strsql += " ,  semicode = " & PrepareStr(SDcode)
                        strsql += " where tirecode = " & PrepareStr(TxtCode.Text.Trim)
                        strsql += " and Rev = " & PrepareStr(TxtRev.Text.Trim)
                        strsql += " and MaterialType = " & PrepareStr("11")
                    ElseIf i = 3 Then
                        strsql = "Update TBLGTDtl"
                        strsql += " set Length = " & PrepareStr(TxtIN_L.Text.Trim)
                        strsql += " , number = " & PrepareStr(TxtIN_N.Text.Trim)
                        strsql += " , QTU = " & PrepareStr((cmbIN.SelectedValue * (TxtIN_L.Text.Trim / 1000)))
                        strsql += " , Dateup =  " & PrepareStr(strDate)
                        strsql += " ,  semicode = " & PrepareStr(INcode)
                        strsql += " where tirecode = " & PrepareStr(TxtCode.Text.Trim)
                        strsql += " and Rev = " & PrepareStr(TxtRev.Text.Trim)
                        strsql += " and MaterialType = " & PrepareStr("12")
                    ElseIf i = 4 Then
                        strsql = "Update TBLGTDtl"
                        strsql += " set Length = " & PrepareStr(TxtCU_L.Text.Trim)
                        strsql += " , number = " & PrepareStr(TxtCu_N.Text.Trim)
                        strsql += " , QTU = " & PrepareStr((cmbCU.SelectedValue * (TxtCU_L.Text.Trim / 1000)))
                        strsql += " , Dateup =  " & PrepareStr(strDate)
                        strsql += " ,  semicode = " & PrepareStr(CUcode)
                        strsql += " where tirecode = " & PrepareStr(TxtCode.Text.Trim)
                        strsql += " and Rev = " & PrepareStr(TxtRev.Text.Trim)
                        strsql += " and MaterialType = " & PrepareStr("03")
                    ElseIf i = 5 Then
                        strsql = "Update TBLGTDtl"
                        strsql += " set number = " & PrepareStr(TxtTT_N.Text.Trim)
                        strsql += " , QTU = " & PrepareStr(cmbTT.SelectedValue * TxtTT_N.Text.Trim)
                        strsql += " , Dateup =  " & PrepareStr(strDate)
                        strsql += " ,  semicode = " & PrepareStr(TTcode)
                        strsql += " where tirecode = " & PrepareStr(TxtCode.Text.Trim)
                        strsql += " and Rev = " & PrepareStr(TxtRev.Text.Trim)
                        strsql += " and MaterialType = " & PrepareStr("13")
                    ElseIf i = 6 Then
                        strsql = "Update TBLGTDtl"
                        strsql += " set Length = " & PrepareStr(TxtBp_L.Text.Trim)
                        strsql += " , number = " & PrepareStr(TxtBP_N.Text.Trim)
                        strsql += " , QTU = " & PrepareStr((cmbBp.SelectedValue * (TxtBp_L.Text.Trim / 1000)))
                        strsql += " , Dateup =  " & PrepareStr(strDate)
                        strsql += " ,  semicode = " & PrepareStr(Bpcode)
                        strsql += " where tirecode = " & PrepareStr(TxtCode.Text.Trim)
                        strsql += " and Rev = " & PrepareStr(TxtRev.Text.Trim)
                        strsql += " and MaterialType = " & PrepareStr("04")
                    ElseIf i = 7 Then
                        If CheckBoxWf.Checked Then
                            strsql = "Update TBLGTDtl"
                            strsql += " set Length = " & PrepareStr(TxtWf_L.Text.Trim)
                            strsql += " , number = " & PrepareStr(TxtWf_N.Text.Trim)
                            strsql += " , QTU = " & PrepareStr((cmbWf.SelectedValue * (TxtWf_L.Text.Trim / 1000)))
                            strsql += " , Dateup =  " & PrepareStr(strDate)
                            strsql += " ,  semicode = " & PrepareStr(Wfcode)
                            strsql += " where tirecode = " & PrepareStr(TxtCode.Text.Trim)
                            strsql += " and Rev = " & PrepareStr(TxtRev.Text.Trim)
                            strsql += " and MaterialType = " & PrepareStr("09")
                        Else
                            strsql = "Update TBLGTDtl"
                            strsql += " set Length = " & PrepareStr(TxtWf_L.Text.Trim)
                            strsql += " , number = " & PrepareStr(TxtWf_N.Text.Trim)
                            strsql += " , QTU = " & PrepareStr("")
                            strsql += " , Dateup =  " & PrepareStr(strDate)
                            strsql += " ,  semicode = " & PrepareStr(Wfcode)
                            strsql += " where tirecode = " & PrepareStr(TxtCode.Text.Trim)
                            strsql += " and Rev = " & PrepareStr(TxtRev.Text.Trim)
                            strsql += " and MaterialType = " & PrepareStr("09")
                        End If
                    ElseIf i = 8 Then
                        strsql = "Update TBLGTDtl"
                        strsql += " set Length = " & PrepareStr(TxtB1_L.Text.Trim)
                        strsql += " , number = " & PrepareStr(TxtB1_N.Text.Trim)
                        strsql += " , QTU = " & PrepareStr((cmbB1.SelectedValue * (TxtB1_L.Text.Trim / 1000)))
                        strsql += " , Dateup =  " & PrepareStr(strDate)
                        strsql += " ,  semicode = " & PrepareStr(B1code)
                        strsql += " where tirecode = " & PrepareStr(TxtCode.Text.Trim)
                        strsql += " and Rev = " & PrepareStr(TxtRev.Text.Trim)
                        strsql += " and MaterialType = " & PrepareStr("05")
                    ElseIf i = 9 Then
                        strsql = "Update TBLGTDtl"
                        strsql += " set Length = " & PrepareStr(TxtB2_L.Text.Trim)
                        strsql += " , number = " & PrepareStr(TxtB2_N.Text.Trim)
                        strsql += " , QTU = " & PrepareStr((cmbB2.SelectedValue * (TxtB2_L.Text.Trim / 1000)))
                        strsql += " , Dateup =  " & PrepareStr(strDate)
                        strsql += " ,  semicode = " & PrepareStr(B2code)
                        strsql += " where tirecode = " & PrepareStr(TxtCode.Text.Trim)
                        strsql += " and Rev = " & PrepareStr(TxtRev.Text.Trim)
                        strsql += " and MaterialType = " & PrepareStr("06")
                    ElseIf i = 10 Then
                        strsql = "Update TBLGTDtl"
                        strsql += " set Length = " & PrepareStr(TxtB3_L.Text.Trim)
                        strsql += " , number = " & PrepareStr(TxtB3_N.Text.Trim)
                        strsql += " , QTU = " & PrepareStr((cmbB3.SelectedValue * (TxtB3_L.Text.Trim / 1000)))
                        strsql += " , Dateup =  " & PrepareStr(strDate)
                        strsql += " ,  semicode = " & PrepareStr(B3code)
                        strsql += " where tirecode = " & PrepareStr(TxtCode.Text.Trim)
                        strsql += " and Rev = " & PrepareStr(TxtRev.Text.Trim)
                        strsql += " and MaterialType = " & PrepareStr("07")
                    ElseIf i = 11 Then
                        strsql = "Update TBLGTDtl"
                        strsql += " set Length = " & PrepareStr(TxtB4_L.Text.Trim)
                        strsql += " , number = " & PrepareStr(TxtB4_N.Text.Trim)
                        strsql += " , QTU = " & PrepareStr((cmbB4.SelectedValue * (TxtB4_L.Text.Trim / 1000)))
                        strsql += " , Dateup =  " & PrepareStr(strDate)
                        strsql += " ,  semicode = " & PrepareStr(B4code)
                        strsql += " where tirecode = " & PrepareStr(TxtCode.Text.Trim)
                        strsql += " and Rev = " & PrepareStr(TxtRev.Text.Trim)
                        strsql += " and MaterialType = " & PrepareStr("08")
                    ElseIf i = 12 Then
                        If CheckBoxNy.Checked Then
                            strsql = "Update TBLGTDtl"
                            strsql += " set Length = " & PrepareStr(TxtNy_L.Text.Trim)
                            strsql += " , number = " & PrepareStr(TxtNy_N.Text.Trim)
                            strsql += " , QTU = " & PrepareStr((cmbNy.SelectedValue * (TxtNy_L.Text.Trim / 1000)))
                            strsql += " , Dateup =  " & PrepareStr(strDate)
                            strsql += " ,  semicode = " & PrepareStr(NYcode)
                            strsql += " where tirecode = " & PrepareStr(TxtCode.Text.Trim)
                            strsql += " and Rev = " & PrepareStr(TxtRev.Text.Trim)
                            strsql += " and MaterialType = " & PrepareStr("10")
                        Else
                            strsql = "Update TBLGTDtl"
                            strsql += " set Length = " & PrepareStr(TxtNy_L.Text.Trim)
                            strsql += " , number = " & PrepareStr(TxtNy_N.Text.Trim)
                            strsql += " , QTU = " & PrepareStr("")
                            strsql += " , Dateup =  " & PrepareStr(strDate)
                            strsql += " ,  semicode = " & PrepareStr(NYcode)
                            strsql += " where tirecode = " & PrepareStr(TxtCode.Text.Trim)
                            strsql += " and Rev = " & PrepareStr(TxtRev.Text.Trim)
                            strsql += " and MaterialType = " & PrepareStr("10")
                        End If
                    End If

                    cmd.CommandText = strsql
                    cmd.ExecuteNonQuery()
                Next
                t1.Commit()
                UPTire_Dtl = True
            Catch
                t1.Rollback()
                UPTire_Dtl = False
                MsgBox("Rollback data")
                Exit Function
            Finally
                cn.Close()
            End Try
        End If
    End Function
    Sub UpTire()
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Try
            strSQL = " Update TblGTHdr"
            strSQL &= " set Active = 0 "
            strSQL &= " where TireCode = '" & TxtCode.Text.Trim & "'"
            strSQL &= "  "
            strSQL &= " Update TblGTHdr"
            strSQL &= " set Active = 1 "
            strSQL &= " where TireCode = '" & TxtCode.Text.Trim & "'"
            strSQL &= " and Rev = '" & TxtRev.Text.Trim & "'"

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
    Sub clear()
        TxtCode.Text = ""
        TxtRev.Text = ""
        txtBF_N.Text = ""
        TxtTT_N.Text = ""
        TxtSD_N.Text = ""
        TxtSD_L.Text = ""
        TxtIN_N.Text = ""
        TxtIN_L.Text = ""
        TxtBP_N.Text = ""
        TxtBp_L.Text = ""
        TxtCu_N.Text = ""
        TxtCU_L.Text = ""
        TxtB1_N.Text = ""
        TxtB1_L.Text = ""
        TxtB2_N.Text = ""
        TxtB2_L.Text = ""
        TxtB3_N.Text = ""
        TxtB3_L.Text = ""
        TxtB4_N.Text = ""
        TxtB4_L.Text = ""
        TxtWf_N.Text = ""
        TxtWf_L.Text = ""
        TxtNy_N.Text = ""
        TxtNy_L.Text = ""
    End Sub
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
End Class
