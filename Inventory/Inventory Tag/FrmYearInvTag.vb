#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports System.Text
Imports Inventory_Tag.Common
Imports Inventory_Tag.FrmInvTag
#End Region

Public Class FrmYearInvTag

#Region "Declare"
    Inherits System.Windows.Forms.Form
    Dim GrdDV As New DataView
    Protected Const TBL_RM As String = "TBL_RM"
    Dim GrdDVComp As New DataView
    Protected Const TBL_Comp As String = "TBL_Comp"
    Dim GrdDVGP As New DataView
    Protected Const TBL_Group As String = "TBL_Group"
    Dim GrdDVLOC As New DataView
    Protected Const TBL_LOC As String = "TBL_LOC"
    Dim GrdDVUSER As New DataView
    Protected Const TBL_USER As String = "TBL_USER"
    Dim GrdDVType As New DataView
    Protected Const TBL_Type As String = "TBL_Type"
    Public Shared tb1 As New DataTable

    Protected DefaultGridBorderStyle As BorderStyle
    Dim C1 As New SQLData("ACCINV")
    Dim StrData As String
    Friend WithEvents ButtonImport As Button
    Friend Username As String
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
    Friend WithEvents ButtonDel As System.Windows.Forms.Button
    Friend WithEvents ButtonAdd As System.Windows.Forms.Button
    Friend WithEvents ButtonClose As System.Windows.Forms.Button
    Friend WithEvents DateYear As System.Windows.Forms.DateTimePicker
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents RdbDate As System.Windows.Forms.RadioButton
    Friend WithEvents RdbYear As System.Windows.Forms.RadioButton
    Friend WithEvents RdbMonth As System.Windows.Forms.RadioButton
    Friend WithEvents Datedate As System.Windows.Forms.DateTimePicker
    Friend WithEvents ButtonView As System.Windows.Forms.Button
    Friend WithEvents cmbLoc As System.Windows.Forms.ComboBox
    Friend WithEvents ChkLoc As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblTag As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ChkType As System.Windows.Forms.CheckBox
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents cmbType As System.Windows.Forms.ComboBox
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cmdPrint As System.Windows.Forms.Button
    Friend WithEvents ChkUser As System.Windows.Forms.CheckBox
    Friend WithEvents cmbUser As System.Windows.Forms.ComboBox
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    Friend WithEvents Datemonth As System.Windows.Forms.DateTimePicker
    Friend WithEvents CmdEdit As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Chktag As System.Windows.Forms.CheckBox
    Friend WithEvents TxtNo1 As System.Windows.Forms.TextBox
    Friend WithEvents TxtNo2 As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmYearInvTag))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.DataGrid1 = New System.Windows.Forms.DataGrid()
        Me.ButtonDel = New System.Windows.Forms.Button()
        Me.ButtonAdd = New System.Windows.Forms.Button()
        Me.ButtonClose = New System.Windows.Forms.Button()
        Me.DateYear = New System.Windows.Forms.DateTimePicker()
        Me.ButtonView = New System.Windows.Forms.Button()
        Me.RdbDate = New System.Windows.Forms.RadioButton()
        Me.RdbYear = New System.Windows.Forms.RadioButton()
        Me.RdbMonth = New System.Windows.Forms.RadioButton()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Datemonth = New System.Windows.Forms.DateTimePicker()
        Me.Datedate = New System.Windows.Forms.DateTimePicker()
        Me.cmbLoc = New System.Windows.Forms.ComboBox()
        Me.ChkLoc = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblName = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblTag = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ChkType = New System.Windows.Forms.CheckBox()
        Me.cmbType = New System.Windows.Forms.ComboBox()
        Me.lblType = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.ChkUser = New System.Windows.Forms.CheckBox()
        Me.cmbUser = New System.Windows.Forms.ComboBox()
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
        Me.CmdEdit = New System.Windows.Forms.Button()
        Me.Chktag = New System.Windows.Forms.CheckBox()
        Me.TxtNo1 = New System.Windows.Forms.TextBox()
        Me.TxtNo2 = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ButtonImport = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.DataGrid1)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 112)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(920, 440)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'DataGrid1
        '
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(3, 16)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.Size = New System.Drawing.Size(914, 421)
        Me.DataGrid1.TabIndex = 0
        '
        'ButtonDel
        '
        Me.ButtonDel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ButtonDel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.ButtonDel.Image = CType(resources.GetObject("ButtonDel.Image"), System.Drawing.Image)
        Me.ButtonDel.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonDel.Location = New System.Drawing.Point(8, 560)
        Me.ButtonDel.Name = "ButtonDel"
        Me.ButtonDel.Size = New System.Drawing.Size(72, 56)
        Me.ButtonDel.TabIndex = 11
        Me.ButtonDel.Text = "DEL"
        Me.ButtonDel.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ButtonAdd
        '
        Me.ButtonAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonAdd.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.ButtonAdd.Image = CType(resources.GetObject("ButtonAdd.Image"), System.Drawing.Image)
        Me.ButtonAdd.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonAdd.Location = New System.Drawing.Point(712, 560)
        Me.ButtonAdd.Name = "ButtonAdd"
        Me.ButtonAdd.Size = New System.Drawing.Size(72, 56)
        Me.ButtonAdd.TabIndex = 10
        Me.ButtonAdd.Text = "ADD"
        Me.ButtonAdd.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ButtonClose
        '
        Me.ButtonClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonClose.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.ButtonClose.Image = CType(resources.GetObject("ButtonClose.Image"), System.Drawing.Image)
        Me.ButtonClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonClose.Location = New System.Drawing.Point(856, 560)
        Me.ButtonClose.Name = "ButtonClose"
        Me.ButtonClose.Size = New System.Drawing.Size(72, 56)
        Me.ButtonClose.TabIndex = 12
        Me.ButtonClose.Text = "CLOSE"
        Me.ButtonClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'DateYear
        '
        Me.DateYear.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DateYear.CustomFormat = "yyyy"
        Me.DateYear.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateYear.Location = New System.Drawing.Point(80, 66)
        Me.DateYear.Name = "DateYear"
        Me.DateYear.ShowUpDown = True
        Me.DateYear.Size = New System.Drawing.Size(104, 20)
        Me.DateYear.TabIndex = 14
        '
        'ButtonView
        '
        Me.ButtonView.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonView.Image = CType(resources.GetObject("ButtonView.Image"), System.Drawing.Image)
        Me.ButtonView.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonView.Location = New System.Drawing.Point(768, 13)
        Me.ButtonView.Name = "ButtonView"
        Me.ButtonView.Size = New System.Drawing.Size(72, 56)
        Me.ButtonView.TabIndex = 13
        Me.ButtonView.Text = "View"
        Me.ButtonView.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'RdbDate
        '
        Me.RdbDate.Location = New System.Drawing.Point(16, 16)
        Me.RdbDate.Name = "RdbDate"
        Me.RdbDate.Size = New System.Drawing.Size(64, 24)
        Me.RdbDate.TabIndex = 16
        Me.RdbDate.Text = " Date"
        '
        'RdbYear
        '
        Me.RdbYear.Location = New System.Drawing.Point(16, 64)
        Me.RdbYear.Name = "RdbYear"
        Me.RdbYear.Size = New System.Drawing.Size(64, 24)
        Me.RdbYear.TabIndex = 18
        Me.RdbYear.Text = " Year"
        '
        'RdbMonth
        '
        Me.RdbMonth.Checked = True
        Me.RdbMonth.Location = New System.Drawing.Point(16, 40)
        Me.RdbMonth.Name = "RdbMonth"
        Me.RdbMonth.Size = New System.Drawing.Size(64, 24)
        Me.RdbMonth.TabIndex = 19
        Me.RdbMonth.TabStop = True
        Me.RdbMonth.Text = "Month"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Datemonth)
        Me.GroupBox2.Controls.Add(Me.Datedate)
        Me.GroupBox2.Controls.Add(Me.RdbDate)
        Me.GroupBox2.Controls.Add(Me.RdbMonth)
        Me.GroupBox2.Controls.Add(Me.RdbYear)
        Me.GroupBox2.Controls.Add(Me.DateYear)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(216, 100)
        Me.GroupBox2.TabIndex = 20
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Select Tag"
        '
        'Datemonth
        '
        Me.Datemonth.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Datemonth.CustomFormat = "MM/yyyy"
        Me.Datemonth.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.Datemonth.Location = New System.Drawing.Point(80, 40)
        Me.Datemonth.Name = "Datemonth"
        Me.Datemonth.ShowUpDown = True
        Me.Datemonth.Size = New System.Drawing.Size(104, 20)
        Me.Datemonth.TabIndex = 22
        '
        'Datedate
        '
        Me.Datedate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Datedate.CustomFormat = "dd/MM/yyyy"
        Me.Datedate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.Datedate.Location = New System.Drawing.Point(80, 16)
        Me.Datedate.Name = "Datedate"
        Me.Datedate.ShowUpDown = True
        Me.Datedate.Size = New System.Drawing.Size(104, 20)
        Me.Datedate.TabIndex = 21
        '
        'cmbLoc
        '
        Me.cmbLoc.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbLoc.Location = New System.Drawing.Point(568, 16)
        Me.cmbLoc.Name = "cmbLoc"
        Me.cmbLoc.Size = New System.Drawing.Size(184, 21)
        Me.cmbLoc.TabIndex = 21
        Me.cmbLoc.Text = "Select"
        '
        'ChkLoc
        '
        Me.ChkLoc.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ChkLoc.Location = New System.Drawing.Point(488, 18)
        Me.ChkLoc.Name = "ChkLoc"
        Me.ChkLoc.Size = New System.Drawing.Size(72, 16)
        Me.ChkLoc.TabIndex = 22
        Me.ChkLoc.Text = "Location"
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(88, 560)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(104, 32)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "Location :"
        '
        'lblName
        '
        Me.lblName.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblName.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblName.Location = New System.Drawing.Point(192, 560)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(232, 32)
        Me.lblName.TabIndex = 24
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(88, 592)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(136, 32)
        Me.Label2.TabIndex = 25
        Me.Label2.Text = "Record Tag : "
        '
        'lblTag
        '
        Me.lblTag.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblTag.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTag.Location = New System.Drawing.Point(224, 592)
        Me.lblTag.Name = "lblTag"
        Me.lblTag.Size = New System.Drawing.Size(104, 32)
        Me.lblTag.TabIndex = 26
        Me.lblTag.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(336, 592)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 32)
        Me.Label3.TabIndex = 27
        Me.Label3.Text = "Record"
        '
        'ChkType
        '
        Me.ChkType.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ChkType.Location = New System.Drawing.Point(488, 42)
        Me.ChkType.Name = "ChkType"
        Me.ChkType.Size = New System.Drawing.Size(56, 16)
        Me.ChkType.TabIndex = 29
        Me.ChkType.Text = "Type"
        '
        'cmbType
        '
        Me.cmbType.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbType.Location = New System.Drawing.Point(568, 40)
        Me.cmbType.Name = "cmbType"
        Me.cmbType.Size = New System.Drawing.Size(184, 21)
        Me.cmbType.TabIndex = 28
        Me.cmbType.Text = "Select"
        '
        'lblType
        '
        Me.lblType.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblType.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblType.Location = New System.Drawing.Point(488, 560)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(184, 32)
        Me.lblType.TabIndex = 31
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(424, 560)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(64, 32)
        Me.Label5.TabIndex = 30
        Me.Label5.Text = "Type:"
        '
        'cmdPrint
        '
        Me.cmdPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPrint.Image = CType(resources.GetObject("cmdPrint.Image"), System.Drawing.Image)
        Me.cmdPrint.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdPrint.Location = New System.Drawing.Point(848, 13)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(72, 56)
        Me.cmdPrint.TabIndex = 32
        Me.cmdPrint.Text = "Print"
        Me.cmdPrint.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ChkUser
        '
        Me.ChkUser.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ChkUser.Location = New System.Drawing.Point(488, 66)
        Me.ChkUser.Name = "ChkUser"
        Me.ChkUser.Size = New System.Drawing.Size(56, 16)
        Me.ChkUser.TabIndex = 34
        Me.ChkUser.Text = "User"
        '
        'cmbUser
        '
        Me.cmbUser.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbUser.Location = New System.Drawing.Point(568, 64)
        Me.cmbUser.Name = "cmbUser"
        Me.cmbUser.Size = New System.Drawing.Size(184, 21)
        Me.cmbUser.TabIndex = 33
        Me.cmbUser.Text = "Select"
        '
        'CmdEdit
        '
        Me.CmdEdit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdEdit.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.CmdEdit.Image = CType(resources.GetObject("CmdEdit.Image"), System.Drawing.Image)
        Me.CmdEdit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdEdit.Location = New System.Drawing.Point(784, 560)
        Me.CmdEdit.Name = "CmdEdit"
        Me.CmdEdit.Size = New System.Drawing.Size(72, 56)
        Me.CmdEdit.TabIndex = 35
        Me.CmdEdit.Text = "EDIT"
        Me.CmdEdit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Chktag
        '
        Me.Chktag.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Chktag.Location = New System.Drawing.Point(488, 90)
        Me.Chktag.Name = "Chktag"
        Me.Chktag.Size = New System.Drawing.Size(64, 16)
        Me.Chktag.TabIndex = 37
        Me.Chktag.Text = "TagNo"
        '
        'TxtNo1
        '
        Me.TxtNo1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNo1.Location = New System.Drawing.Point(568, 88)
        Me.TxtNo1.Name = "TxtNo1"
        Me.TxtNo1.Size = New System.Drawing.Size(72, 20)
        Me.TxtNo1.TabIndex = 38
        '
        'TxtNo2
        '
        Me.TxtNo2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNo2.Location = New System.Drawing.Point(680, 88)
        Me.TxtNo2.Name = "TxtNo2"
        Me.TxtNo2.Size = New System.Drawing.Size(72, 20)
        Me.TxtNo2.TabIndex = 39
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.Location = New System.Drawing.Point(648, 90)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(24, 16)
        Me.Label4.TabIndex = 40
        Me.Label4.Text = "to"
        '
        'ButtonImport
        '
        Me.ButtonImport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonImport.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.ButtonImport.Image = CType(resources.GetObject("ButtonImport.Image"), System.Drawing.Image)
        Me.ButtonImport.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonImport.Location = New System.Drawing.Point(620, 560)
        Me.ButtonImport.Name = "ButtonImport"
        Me.ButtonImport.Size = New System.Drawing.Size(72, 56)
        Me.ButtonImport.TabIndex = 41
        Me.ButtonImport.Text = "Import"
        Me.ButtonImport.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'FrmYearInvTag
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(936, 630)
        Me.Controls.Add(Me.ButtonImport)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TxtNo1)
        Me.Controls.Add(Me.Chktag)
        Me.Controls.Add(Me.CmdEdit)
        Me.Controls.Add(Me.ChkUser)
        Me.Controls.Add(Me.cmbUser)
        Me.Controls.Add(Me.cmdPrint)
        Me.Controls.Add(Me.lblType)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.ChkType)
        Me.Controls.Add(Me.cmbType)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.lblTag)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ChkLoc)
        Me.Controls.Add(Me.cmbLoc)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.ButtonView)
        Me.Controls.Add(Me.ButtonDel)
        Me.Controls.Add(Me.ButtonAdd)
        Me.Controls.Add(Me.ButtonClose)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.TxtNo2)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmYearInvTag"
        Me.Text = "Inventory Tag ( Year )"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

#Region "CONSTANT"
    Dim DT As New DataTable
    Dim StrSQLPRT As String
    Dim oldrow As Integer
#End Region

#Region "COMBOBOX"
    Sub LoadLoc()
        Dim dtLoc As DataTable = New DataTable()
        Dim strSQL As String = String.Empty
        Dim sb As StringBuilder = New StringBuilder()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sb.AppendLine("SELECT  * ")
        sb.AppendLine("  FROM  TBLDepartment  ")
        strSQL = sb.ToString()

        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(strSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtLoc = New DataTable
            DA.Fill(dtLoc)
        Catch
        Finally
        End Try
        dtLoc.TableName = TBL_LOC
        GrdDVLOC = dtLoc.DefaultView
        '************************************
        cmbLoc.DisplayMember = "DeptName"
        cmbLoc.ValueMember = "DeptCode"
        cmbLoc.DataSource = dtLoc
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadUser()
        Dim dtUser As DataTable = New DataTable()
        Dim strSQL As String = String.Empty
        Dim sb As StringBuilder = New StringBuilder()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sb.AppendLine("SELECT    empcode, personfnameeng + '  '+PersonlNameeng ename")
        sb.AppendLine("  FROM         BTMTMASTER..TblEmployee  ")
        sb.AppendLine("  where empcode in (select empcode from TblUser) ")
        strSQL = sb.ToString()

        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(strSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtUser = New DataTable
            DA.Fill(dtUser)
        Catch
        Finally
        End Try
        dtUser.TableName = TBL_USER
        GrdDVUSER = dtUser.DefaultView
        '************************************
        cmbUser.DisplayMember = "ename"
        cmbUser.ValueMember = "empcode"
        cmbUser.DataSource = dtUser
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadType()
        Dim dtType As DataTable = New DataTable()
        Dim strSQL As String = String.Empty
        Dim sb As StringBuilder = New StringBuilder()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sb.AppendLine("SELECT  * ")
        sb.AppendLine("  FROM  TBLType ")
        strSQL = sb.ToString()

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
        cmbType.DisplayMember = "TypeName"
        cmbType.ValueMember = "TypeCode"
        cmbType.DataSource = dtType
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
#End Region

#Region "Function_Load"
    Private Sub LoadCOM()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        Dim sb As StringBuilder = New StringBuilder()
        Dim strSQL As String = String.Empty
        sb.AppendLine(" select Tagno,code,Period,TrxYear,Typecode,TypeName,Location,DeptName,")
        sb.AppendLine(" Qty,Uom,UpdateDate,UpdateTime,UserID ")
        sb.AppendLine(" , PersonFnameEng+ ' '+ PersonLnameEng EName")
        sb.AppendLine(" , TrxDate,TrxTime,substring(TrxDate,7,2)+'/'+substring(TrxDate,5,2)+'/'+substring(TrxDate,1,4) dd ")
        sb.AppendLine(" , substring(UpdateDate,7,2)+'/'+substring(UpdateDate,5,2)+'/'+substring(UpdateDate,1,4) dd1 ")
        sb.AppendLine(" from (")
        sb.AppendLine("     select Tagno,code,Period,TrxYear,tx.Typecode,TypeName,Location,DeptName,")
        sb.AppendLine("     TrxDate,TrxTime,Qty,Uom,UserID,UpdateDate,UpdateTime ")
        sb.AppendLine("     from TBLTRX tx")
        sb.AppendLine("     left outer join TBLDepartment dp")
        sb.AppendLine("     on tx.location = dp.deptcode")
        sb.AppendLine("     left outer join TBLTYPE ty")
        sb.AppendLine("     on tx.typecode = ty.typecode ")
        sb.AppendLine(" ) trx")
        sb.AppendLine(" left outer join BTMTMASTER..TBLEmployee emp")
        sb.AppendLine(" on trx.UserID = emp.empcode")
        sb.AppendLine(" order by TagNo,trxYear,Location ")
        strSQL = sb.ToString()
        StrSQLPRT = strSQL
        If Not DT Is Nothing Then
            If DT.Rows.Count >= 1 Then
                DT.Clear()
            End If
        End If

        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(strSQL, C1.Strcon)
            Dim CB As New SqlCommandBuilder(DA)
            tb1 = New DataTable
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
        DataGrid1.DataSource = GrdDV
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

        With DataGrid1
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
            .HeaderBackColor = Color.Violet
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
            .HeaderText = "TagNo."
            .MappingName = "tagno"
            .Width = 50
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle0_1 As New DataGridColoredLine2
        With grdColStyle0_1
            .HeaderText = "Period"
            .MappingName = "Period"
            .Width = 50
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle0_2 As New DataGridColoredLine2
        With grdColStyle0_2
            .HeaderText = "TrxYear"
            .MappingName = "TrxYear"
            .Width = 50
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle0 As New DataGridColoredLine2
        With grdColStyle0
            .HeaderText = "TypeName"
            .MappingName = "typeName"
            .Width = 110
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle1 As New DataGridColoredLine2
        With grdColStyle1
            .HeaderText = "Material"
            .MappingName = "Code"
            .Width = 145
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle2 As New DataGridColoredLine2
        With grdColStyle2
            .HeaderText = "Location"
            .MappingName = "DeptName"
            .Width = 180
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle2_1 As New DataGridColoredLine2
        With grdColStyle2_1
            .HeaderText = "UserName"
            .MappingName = "EName"
            .NullText = ""
            .Width = 180
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle4 As New DataGridColoredLine2
        With grdColStyle4
            .HeaderText = "Qty"
            .MappingName = "Qty"
            .Width = 75
            .Format = "##,###,##0.00"
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle5 As New DataGridColoredLine2
        With grdColStyle5
            .HeaderText = "Unit"
            .MappingName = "Uom"
            .Width = 75
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle6_0 As New DataGridColoredLine2
        With grdColStyle6_0
            .HeaderText = "TRXDate"
            .MappingName = "dd"
            .Width = 75
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle6 As New DataGridColoredLine2
        With grdColStyle6
            .HeaderText = "I/P Date"
            .MappingName = "dd1"
            .Width = 75
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle7 As New DataGridColoredLine2
        With grdColStyle7
            .HeaderText = "Time"
            .MappingName = "UpdateTime"
            .Width = 75
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With

        grdTableStyle1.GridColumnStyles.AddRange _
    (New DataGridColumnStyle() _
    {grdColStyle2, grdColStyle0_1, grdColStyle0_2, grdColStyle0_0, grdColStyle1, _
    grdColStyle0, grdColStyle4, grdColStyle5, _
    grdColStyle6_0, grdColStyle2_1, grdColStyle6, grdColStyle7})

        DataGrid1.TableStyles.Add(grdTableStyle1)
        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub
    Private Sub ResetTableStyle()
        ' Clear out the existing TableStyles and result default formatting.
        With DataGrid1
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

    Private Sub FrmYearInvTag_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If ChkLoc.Checked = True Then
            LoadLoc()
        Else
            cmbLoc.Text = "Select"
        End If
        If ChkType.Checked = True Then
            LoadType()
        Else
            cmbType.Text = "Select"
        End If

        LoadCOM()
        GetTypeinv()
        GetBrand()
        GetLocation()
        selectData()
        lblTag.Text = GrdDV.Count
    End Sub

    Private Sub ButtonAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonAdd.Click
        Dim fadd As New FrmAdd
        fadd.ShowDialog()
        LoadCOM()
        ViewData()
    End Sub

    Private Sub ButtonDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonDel.Click
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult

        msg = "Inventory Delete TrxNo : " & GrdDV.Item(oldrow).Row("Tagno")  ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Inventory"   ' Define title.
        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            DelTRX()
            LoadCOM()
            ViewData()
        Else
            Exit Sub
        End If
        oldrow = 0

    End Sub

    Sub DelTRX()
        Dim strsql As String = String.Empty
        Dim cn As New SqlConnection(C1.Strcon)
        Dim cmd As New SqlCommand(strsql, cn)
        cn.Open()
        Dim t1 As SqlTransaction = cn.BeginTransaction
        cmd.Transaction = t1
        Dim strDate() As String
        Dim strTime As String
        strDate = Split(Date.Now.ToShortDateString, "/")
        strTime = Date.Now.ToShortTimeString
        Try
            strsql = "Delete TBLTRX"
            strsql += " where tagNo = " & PrepareStr(GrdDV.Item(oldrow).Row("Tagno"))
            strsql += " and  period = " & PrepareStr(GrdDV.Item(oldrow).Row("Period"))
            strsql += " and  trxyear = " & PrepareStr(GrdDV.Item(oldrow).Row("trxyear"))
            strsql += " and  Location = " & PrepareStr(GrdDV.Item(oldrow).Row("Location"))
            cmd.CommandText = strsql
            cmd.ExecuteNonQuery()
            MsgBox("Delete Complete.", MsgBoxStyle.Information, "Inventory Record")
            t1.Commit()
        Catch
            t1.Rollback()
            MsgBox("Rollback data")
        Finally
            cn.Close()
        End Try
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

    Private Sub ButtonClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonClose.Click
        Me.Close()
    End Sub

#Region "SelectDate"
    Sub selectData()
        Dim fm As New FrmInvTag
        Dim strDate, strMonth, str() As String
        Dim id As String
        If CurrentLevel.Trim = "Administrator" Then
            id = ""
        Else
            id = CurrentIDUser.Trim
        End If
        If RdbDate.Checked = True Then
            str = Split(Datedate.Text.Trim, "/")
            strDate = str(2) + str(1) + str(0)
            GrdDV.RowFilter = " TrxDate like  '" & strDate.Trim & "%'" _
                                & " and UserId like '%" & id.Trim & "%'"
            DataGrid1.DataSource = GrdDV
        ElseIf RdbMonth.Checked = True Then
            str = Split(Datemonth.Text.Trim, "/")
            strMonth = str(1) + str(0)
            GrdDV.RowFilter = " TrxDate like  '%" & strMonth.Trim & "%'" _
                                & " and UserId like '%" & id.Trim & "%'"
            DataGrid1.DataSource = GrdDV
        ElseIf RdbYear.Checked = True Then
            GrdDV.RowFilter = " TrxYear like  '%" & DateYear.Text.Trim & "%'" _
                                & " and UserId like '%" & id.Trim & "%'"
            DataGrid1.DataSource = GrdDV
        End If
    End Sub

    Private Sub DateYear_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateYear.ValueChanged
        selectData()
    End Sub
    Private Sub Datemonth_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Datemonth.ValueChanged
        selectData()
    End Sub
    Private Sub Datedate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Datedate.ValueChanged
        selectData()
    End Sub

    Sub ViewData()
        If ChkLoc.Checked = True And ChkType.Checked = True And ChkUser.Checked = False And Chktag.Checked = False Then
            selectData()
            GrdDV.RowFilter &= " and Location like '%" & cmbLoc.SelectedValue & "%'"
            GrdDV.RowFilter &= " and Typecode like '%" & cmbType.SelectedValue & "%'"
            DataGrid1.DataSource = GrdDV
            lblName.Text = cmbLoc.Text.Trim
            lblType.Text = cmbType.Text.Trim
            lblTag.Text = GrdDV.Count
        ElseIf ChkLoc.Checked = True And ChkType.Checked = True And ChkUser.Checked = True And Chktag.Checked = False Then
            selectData()
            GrdDV.RowFilter &= " and Location like '%" & cmbLoc.SelectedValue & "%'"
            GrdDV.RowFilter &= " and Typecode like '%" & cmbType.SelectedValue & "%'"
            GrdDV.RowFilter &= " and userid like '%" & cmbUser.SelectedValue & "%'"
            DataGrid1.DataSource = GrdDV
            lblName.Text = cmbLoc.Text.Trim
            lblType.Text = cmbType.Text.Trim
            lblTag.Text = GrdDV.Count
        ElseIf ChkLoc.Checked = True And ChkType.Checked = False And ChkUser.Checked = True And Chktag.Checked = False Then
            selectData()
            GrdDV.RowFilter &= " and Location like '%" & cmbLoc.SelectedValue & "%'"
            GrdDV.RowFilter &= " and userid like '%" & cmbUser.SelectedValue & "%'"
            DataGrid1.DataSource = GrdDV
            lblName.Text = cmbLoc.Text.Trim
            lblType.Text = ""
            lblTag.Text = GrdDV.Count
        ElseIf ChkLoc.Checked = True And ChkType.Checked = False And ChkUser.Checked = False And Chktag.Checked = False Then
            selectData()
            GrdDV.RowFilter &= " and Location like '%" & cmbLoc.SelectedValue & "%'"
            DataGrid1.DataSource = GrdDV
            lblName.Text = cmbLoc.Text.Trim
            lblType.Text = ""
            lblTag.Text = GrdDV.Count
        ElseIf ChkLoc.Checked = False And ChkType.Checked = True And ChkUser.Checked = False And Chktag.Checked = False Then
            selectData()
            GrdDV.RowFilter &= " and Typecode like '%" & cmbType.SelectedValue & "%'"
            DataGrid1.DataSource = GrdDV
            lblName.Text = ""
            lblType.Text = cmbType.Text.Trim
            lblTag.Text = GrdDV.Count
        ElseIf ChkLoc.Checked = False And ChkType.Checked = True And ChkUser.Checked = True And Chktag.Checked = False Then
            selectData()
            GrdDV.RowFilter &= " and Typecode like '%" & cmbType.SelectedValue & "%'"
            GrdDV.RowFilter &= " and userid like '%" & cmbUser.SelectedValue & "%'"
            DataGrid1.DataSource = GrdDV
            lblName.Text = ""
            lblType.Text = cmbType.Text.Trim
            lblTag.Text = GrdDV.Count
        ElseIf ChkLoc.Checked = False And ChkType.Checked = False And ChkUser.Checked = True And Chktag.Checked = False Then
            selectData()
            GrdDV.RowFilter &= " and userid like '%" & cmbUser.SelectedValue & "%'"
            DataGrid1.DataSource = GrdDV
            lblName.Text = ""
            lblType.Text = cmbType.Text.Trim
            lblTag.Text = GrdDV.Count
        ElseIf ChkLoc.Checked = True And ChkType.Checked = True And ChkUser.Checked = False And Chktag.Checked = True Then
            selectData()
            GrdDV.RowFilter &= " and Location like '%" & cmbLoc.SelectedValue & "%'"
            GrdDV.RowFilter &= " and Typecode like '%" & cmbType.SelectedValue & "%'"
            GrdDV.RowFilter &= " and tagno >='" & TxtNo1.Text.Trim & "'"
            GrdDV.RowFilter &= " and tagno <='" & TxtNo2.Text.Trim & "'"
            DataGrid1.DataSource = GrdDV
            lblName.Text = cmbLoc.Text.Trim
            lblType.Text = cmbType.Text.Trim
            lblTag.Text = GrdDV.Count
        ElseIf ChkLoc.Checked = True And ChkType.Checked = True And ChkUser.Checked = True And Chktag.Checked = True Then
            selectData()
            GrdDV.RowFilter &= " and Location like '%" & cmbLoc.SelectedValue & "%'"
            GrdDV.RowFilter &= " and Typecode like '%" & cmbType.SelectedValue & "%'"
            GrdDV.RowFilter &= " and userid like '%" & cmbUser.SelectedValue & "%'"
            GrdDV.RowFilter &= " and tagno >='" & TxtNo1.Text.Trim & "'"
            GrdDV.RowFilter &= " and tagno <='" & TxtNo2.Text.Trim & "'"
            DataGrid1.DataSource = GrdDV
            lblName.Text = cmbLoc.Text.Trim
            lblType.Text = cmbType.Text.Trim
            lblTag.Text = GrdDV.Count
        ElseIf ChkLoc.Checked = True And ChkType.Checked = False And ChkUser.Checked = True And Chktag.Checked = True Then
            selectData()
            GrdDV.RowFilter &= " and Location like '%" & cmbLoc.SelectedValue & "%'"
            GrdDV.RowFilter &= " and userid like '%" & cmbUser.SelectedValue & "%'"
            GrdDV.RowFilter &= " and tagno >='" & TxtNo1.Text.Trim & "'"
            GrdDV.RowFilter &= " and tagno <='" & TxtNo2.Text.Trim & "'"
            DataGrid1.DataSource = GrdDV
            lblName.Text = cmbLoc.Text.Trim
            lblType.Text = ""
            lblTag.Text = GrdDV.Count
        ElseIf ChkLoc.Checked = True And ChkType.Checked = False And ChkUser.Checked = False And Chktag.Checked = True Then
            selectData()
            GrdDV.RowFilter &= " and Location like '%" & cmbLoc.SelectedValue & "%'"
            GrdDV.RowFilter &= " and tagno >='" & TxtNo1.Text.Trim & "'"
            GrdDV.RowFilter &= " and tagno <='" & TxtNo2.Text.Trim & "'"
            DataGrid1.DataSource = GrdDV
            lblName.Text = cmbLoc.Text.Trim
            lblType.Text = ""
            lblTag.Text = GrdDV.Count
        ElseIf ChkLoc.Checked = False And ChkType.Checked = True And ChkUser.Checked = False And Chktag.Checked = True Then
            selectData()
            GrdDV.RowFilter &= " and Typecode like '%" & cmbType.SelectedValue & "%'"
            GrdDV.RowFilter &= " and tagno >='" & TxtNo1.Text.Trim & "'"
            GrdDV.RowFilter &= " and tagno <='" & TxtNo2.Text.Trim & "'"
            DataGrid1.DataSource = GrdDV
            lblName.Text = ""
            lblType.Text = cmbType.Text.Trim
            lblTag.Text = GrdDV.Count
        ElseIf ChkLoc.Checked = False And ChkType.Checked = True And ChkUser.Checked = True And Chktag.Checked = True Then
            selectData()
            GrdDV.RowFilter &= " and Typecode like '%" & cmbType.SelectedValue & "%'"
            GrdDV.RowFilter &= " and userid like '%" & cmbUser.SelectedValue & "%'"
            GrdDV.RowFilter &= " and tagno >='" & TxtNo1.Text.Trim & "'"
            GrdDV.RowFilter &= " and tagno <='" & TxtNo2.Text.Trim & "'"
            DataGrid1.DataSource = GrdDV
            lblName.Text = ""
            lblType.Text = cmbType.Text.Trim
            lblTag.Text = GrdDV.Count
        ElseIf ChkLoc.Checked = False And ChkType.Checked = False And ChkUser.Checked = True And Chktag.Checked = True Then
            selectData()
            GrdDV.RowFilter &= " and userid like '%" & cmbUser.SelectedValue & "%'"
            GrdDV.RowFilter &= " and tagno >='" & TxtNo1.Text.Trim & "'"
            GrdDV.RowFilter &= " and tagno <='" & TxtNo2.Text.Trim & "'"
            DataGrid1.DataSource = GrdDV
            lblName.Text = ""
            lblType.Text = cmbType.Text.Trim
            lblTag.Text = GrdDV.Count
        ElseIf ChkLoc.Checked = False And ChkType.Checked = False And ChkUser.Checked = False And Chktag.Checked = True Then
            selectData()
            GrdDV.RowFilter &= " and tagno >='" & TxtNo1.Text.Trim & "'"
            GrdDV.RowFilter &= " and tagno <='" & TxtNo2.Text.Trim & "'"
            DataGrid1.DataSource = GrdDV
            lblName.Text = ""
            lblType.Text = cmbType.Text.Trim
            lblTag.Text = GrdDV.Count
        Else
            selectData()
            cmbLoc.Text = "Select"
            lblTag.Text = GrdDV.Count
            lblName.Text = ""
            lblType.Text = ""
        End If
    End Sub
    Private Sub ButtonView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonView.Click
        ViewData()
    End Sub
#End Region

    Private Sub DataGrid1_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGrid1.DoubleClick
        Dim fed As New FrmEdit
        fed.TxtTagNo.Text = GrdDV.Item(oldrow).Row("Tagno")
        fed.TType = GrdDV.Item(oldrow).Row("TypeName")
        fed.TLoc = GrdDV.Item(oldrow).Row("DeptName")
        fed.TLocNo = GrdDV.Item(oldrow).Row("Location")
        fed.TRMCode = GrdDV.Item(oldrow).Row("code")
        fed.TxtQty.Text = GrdDV.Item(oldrow).Row("Qty")
        fed.TUnit = GrdDV.Item(oldrow).Row("Uom")
        fed.Ttime = Mid(GrdDV.Item(oldrow).Row("trxdate"), 7, 2) + "/" + _
                    Mid(GrdDV.Item(oldrow).Row("trxdate"), 5, 2) + "/" + _
                    Mid(GrdDV.Item(oldrow).Row("trxdate"), 1, 4)
        fed.ShowDialog()
        LoadCOM()
        ViewData()
    End Sub


    Private Sub DataGrid1_CurrentCellChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGrid1.CurrentCellChanged
        oldrow = DataGrid1.CurrentCell.RowNumber
    End Sub

    Private Sub ChkLoc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkLoc.CheckedChanged
        If ChkLoc.Checked = True Then
            LoadLoc()
            lblName.Text = cmbLoc.Text.Trim
        Else
            cmbLoc.Text = "Select"
            lblName.Text = ""
        End If
    End Sub

    Private Sub cmbLoc_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbLoc.SelectedIndexChanged
        If ChkLoc.Checked = True Then
            lblName.Text = cmbLoc.Text.Trim
        Else
            lblName.Text = ""
        End If
    End Sub

    Private Sub ChkType_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkType.CheckedChanged
        If ChkType.Checked = True Then
            LoadType()
            lblType.Text = cmbType.Text.Trim
        Else
            lblType.Text = ""
            cmbType.Text = "Select"
        End If
    End Sub

    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrint.Click
        Dim i As Integer
        Dim fRpt As New FrmPrt
        Dim aDr() As DataRow = GrdDV.Table.Select(GrdDV.RowFilter)
        Dim dr As DataRow
        Dim tbNew As DataTable
        tbNew = New DataTable
        tbNew = DT.Clone
        For Each dr In aDr
            Dim drNew As DataRow
            drNew = tbNew.NewRow
            For i = 0 To GrdDV.Table.Columns.Count - 1
                drNew(i) = dr(i)
            Next
            tbNew.Rows.Add(drNew)
        Next
        tbNew.AcceptChanges()
        Dim dt4prt As DataTable
        dt4prt = New DataTable
        dt4prt = tbNew

        'Dim c34 As String = Chr(34)
        'For i = 0 To dt4prt.Columns.Count - 1
        '    Dim col As String = dt4prt.Columns(i).ColumnName
        '    Dim coltype As String = dt4prt.Columns(i).DataType.FullName
        '    coltype = coltype.Replace("System.", "")
        '    coltype = coltype.Replace("Int32", "integer")
        '    coltype = coltype.Replace("Int16", "integer")
        '    coltype = coltype.Replace("String", "string")
        '    coltype = coltype.Replace("Decimal", "decimal")
        '    Debug.WriteLine("<xs:element name=" & c34 & col.Trim & c34 & "  type= " & c34 & "xs:" & coltype & c34 & " minOccurs=" & c34 & "0" & c34 & "/>")
        'Next
        fRpt.dt_new = tbNew
        fRpt.sUser = Username
        fRpt.ShowDialog()
    End Sub

#Region "Print"
    'Private Sub cmdprint_click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If IsInstallPrinter() = True Then
    '        If prDlg.ShowDialog = DialogResult.OK Then
    '            prDoc.Print()
    '        End If
    '    Else
    '        MsgBox("Cannot Printing. Please Install Printer Now.", MsgBoxStyle.OKCancel, MessageBoxIcon.Information)
    '    End If
    'End Sub

    'Private Sub StringPrint_Print(ByVal sender As Object, ByVal e As PrintPageEventArgs)
    '    AnyString(e.Graphics, lblName.Text, 200, 140)

    '    Dim i As Integer = 0
    '    Dim CurrentYPosition As Integer = 430
    '    Dim strColumn1 As String = ""
    '    Dim strColumn2 As String
    '    Dim strColumn3 As String
    '    Dim strColumn4 As String
    '    Dim strColumn5 As String
    '    Dim strColumn6 As String
    '    Dim strColumn7 As String
    '    Dim strColumn8 As String
    '    Dim strColumn9 As String
    '    Dim strColumn10 As String
    '    Dim strColumn11 As String


    'End Sub

    'Private Sub AnyString(ByVal g As Graphics, ByVal printString As String, ByVal xPos As Integer, ByVal yPos As Integer)
    '    Dim anyPoint As New PointF(xPos, yPos)
    '    g.DrawString(printString, usefont, Brushes.Black, anyPoint)
    'End Sub

    'Private Function IsInstallPrinter() As Boolean
    '    IsInstallPrinter = False

    '    If prDoc.PrinterSettings.PrinterName = "<no default Printer>" Then
    '        IsInstallPrinter = False
    '    Else
    '        IsInstallPrinter = True
    '    End If
    'End Function

#End Region

    Private Sub ChkUser_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkUser.CheckedChanged
        If ChkUser.Checked = True Then
            LoadUser()
        Else
            cmbUser.Text = "Select"
        End If
    End Sub

    Private Sub CmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdEdit.Click
        Dim fed As New FrmEdit
        fed.TxtTagNo.Text = GrdDV.Item(oldrow).Row("Tagno")
        fed.TType = GrdDV.Item(oldrow).Row("TypeName")
        fed.TLoc = GrdDV.Item(oldrow).Row("DeptName")
        fed.TLocNo = GrdDV.Item(oldrow).Row("Location")
        fed.TRMCode = GrdDV.Item(oldrow).Row("code")
        fed.TxtQty.Text = GrdDV.Item(oldrow).Row("Qty")
        fed.TUnit = GrdDV.Item(oldrow).Row("Uom")
        fed.Ttime = Mid(GrdDV.Item(oldrow).Row("trxdate"), 7, 2) + "/" + _
                    Mid(GrdDV.Item(oldrow).Row("trxdate"), 5, 2) + "/" + _
                    Mid(GrdDV.Item(oldrow).Row("trxdate"), 1, 4)
        fed.ShowDialog()
        LoadCOM()
        ViewData()
    End Sub

    Private Sub CHKTAG_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Chktag.CheckedChanged
        Dim i, j As Integer
        If Chktag.Checked Then
            If TxtNo1.Text = "" And TxtNo2.Text = "" Then
                TxtNo1.Text = 1
                TxtNo2.Text = 1
            End If
            i = TxtNo1.Text.Trim
            j = TxtNo2.Text.Trim
            TxtNo1.Text = Format(i, "0000")
            TxtNo2.Text = Format(j, "0000")
        Else
            TxtNo1.Text = ""
            TxtNo2.Text = ""
        End If
    End Sub

End Class
