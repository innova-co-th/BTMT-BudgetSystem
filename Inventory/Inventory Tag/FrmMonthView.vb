#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Inventory_Tag.Common
Imports Inventory_Tag.FrmInvTag
#End Region

Public Class FrmMonthView
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
    Dim GrdDVMType As New DataView
    Protected Const TBL_MType As String = "TBL_MType"
    Public Shared tb1 As New DataTable

    Protected DefaultGridBorderStyle As BorderStyle
    Dim C1 As New SQLData("ACCINV")
    Dim StrData As String
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
    Friend WithEvents cmbType As System.Windows.Forms.ComboBox
    Friend WithEvents cmbLoc As System.Windows.Forms.ComboBox
    Friend WithEvents DPeriod As System.Windows.Forms.DateTimePicker
    Friend WithEvents CmdView As System.Windows.Forms.Button
    Friend WithEvents CHKType As System.Windows.Forms.CheckBox
    Friend WithEvents CHKMonth As System.Windows.Forms.CheckBox
    Friend WithEvents SelectRaw As System.Windows.Forms.RadioButton
    Friend WithEvents cmbMType As System.Windows.Forms.ComboBox
    Friend WithEvents CHKMType As System.Windows.Forms.CheckBox
    Friend WithEvents CHKLoc As System.Windows.Forms.CheckBox
    Friend WithEvents CHKCODE As System.Windows.Forms.CheckBox
    Friend WithEvents TXTCODE As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmMonthView))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.TXTCODE = New System.Windows.Forms.TextBox
        Me.CHKCODE = New System.Windows.Forms.CheckBox
        Me.cmbMType = New System.Windows.Forms.ComboBox
        Me.CHKMType = New System.Windows.Forms.CheckBox
        Me.SelectRaw = New System.Windows.Forms.RadioButton
        Me.DPeriod = New System.Windows.Forms.DateTimePicker
        Me.CmdView = New System.Windows.Forms.Button
        Me.cmbType = New System.Windows.Forms.ComboBox
        Me.cmbLoc = New System.Windows.Forms.ComboBox
        Me.CHKLoc = New System.Windows.Forms.CheckBox
        Me.CHKType = New System.Windows.Forms.CheckBox
        Me.CHKMonth = New System.Windows.Forms.CheckBox
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.TXTCODE)
        Me.GroupBox1.Controls.Add(Me.CHKCODE)
        Me.GroupBox1.Controls.Add(Me.cmbMType)
        Me.GroupBox1.Controls.Add(Me.CHKMType)
        Me.GroupBox1.Controls.Add(Me.SelectRaw)
        Me.GroupBox1.Controls.Add(Me.DPeriod)
        Me.GroupBox1.Controls.Add(Me.CmdView)
        Me.GroupBox1.Controls.Add(Me.cmbType)
        Me.GroupBox1.Controls.Add(Me.cmbLoc)
        Me.GroupBox1.Controls.Add(Me.CHKLoc)
        Me.GroupBox1.Controls.Add(Me.CHKType)
        Me.GroupBox1.Controls.Add(Me.CHKMonth)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(288, 248)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'TXTCODE
        '
        Me.TXTCODE.Location = New System.Drawing.Point(88, 152)
        Me.TXTCODE.Name = "TXTCODE"
        Me.TXTCODE.Size = New System.Drawing.Size(184, 20)
        Me.TXTCODE.TabIndex = 73
        Me.TXTCODE.Text = ""
        '
        'CHKCODE
        '
        Me.CHKCODE.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.CHKCODE.Location = New System.Drawing.Point(16, 152)
        Me.CHKCODE.Name = "CHKCODE"
        Me.CHKCODE.Size = New System.Drawing.Size(72, 16)
        Me.CHKCODE.TabIndex = 72
        Me.CHKCODE.Text = "CODE"
        '
        'cmbMType
        '
        Me.cmbMType.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbMType.Location = New System.Drawing.Point(120, 88)
        Me.cmbMType.Name = "cmbMType"
        Me.cmbMType.Size = New System.Drawing.Size(152, 21)
        Me.cmbMType.TabIndex = 71
        Me.cmbMType.Text = "Select"
        '
        'CHKMType
        '
        Me.CHKMType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.CHKMType.Location = New System.Drawing.Point(16, 90)
        Me.CHKMType.Name = "CHKMType"
        Me.CHKMType.Size = New System.Drawing.Size(104, 16)
        Me.CHKMType.TabIndex = 70
        Me.CHKMType.Text = "MaterialType"
        '
        'SelectRaw
        '
        Me.SelectRaw.Checked = True
        Me.SelectRaw.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.SelectRaw.Location = New System.Drawing.Point(16, 184)
        Me.SelectRaw.Name = "SelectRaw"
        Me.SelectRaw.TabIndex = 68
        Me.SelectRaw.TabStop = True
        Me.SelectRaw.Text = "By  Raw"
        '
        'DPeriod
        '
        Me.DPeriod.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DPeriod.CustomFormat = "MM/yyyy"
        Me.DPeriod.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DPeriod.Location = New System.Drawing.Point(88, 24)
        Me.DPeriod.Name = "DPeriod"
        Me.DPeriod.ShowUpDown = True
        Me.DPeriod.Size = New System.Drawing.Size(80, 20)
        Me.DPeriod.TabIndex = 67
        '
        'CmdView
        '
        Me.CmdView.Image = CType(resources.GetObject("CmdView.Image"), System.Drawing.Image)
        Me.CmdView.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdView.Location = New System.Drawing.Point(192, 184)
        Me.CmdView.Name = "CmdView"
        Me.CmdView.Size = New System.Drawing.Size(75, 56)
        Me.CmdView.TabIndex = 31
        Me.CmdView.Text = "Report"
        Me.CmdView.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'cmbType
        '
        Me.cmbType.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbType.Location = New System.Drawing.Point(88, 56)
        Me.cmbType.Name = "cmbType"
        Me.cmbType.Size = New System.Drawing.Size(184, 21)
        Me.cmbType.TabIndex = 30
        Me.cmbType.Text = "Select"
        '
        'cmbLoc
        '
        Me.cmbLoc.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbLoc.Location = New System.Drawing.Point(88, 120)
        Me.cmbLoc.Name = "cmbLoc"
        Me.cmbLoc.Size = New System.Drawing.Size(184, 21)
        Me.cmbLoc.TabIndex = 29
        Me.cmbLoc.Text = "Select"
        '
        'CHKLoc
        '
        Me.CHKLoc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.CHKLoc.Location = New System.Drawing.Point(16, 122)
        Me.CHKLoc.Name = "CHKLoc"
        Me.CHKLoc.Size = New System.Drawing.Size(72, 16)
        Me.CHKLoc.TabIndex = 2
        Me.CHKLoc.Text = "Location"
        '
        'CHKType
        '
        Me.CHKType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.CHKType.Location = New System.Drawing.Point(16, 58)
        Me.CHKType.Name = "CHKType"
        Me.CHKType.Size = New System.Drawing.Size(56, 16)
        Me.CHKType.TabIndex = 1
        Me.CHKType.Text = "Type"
        '
        'CHKMonth
        '
        Me.CHKMonth.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.CHKMonth.Location = New System.Drawing.Point(16, 26)
        Me.CHKMonth.Name = "CHKMonth"
        Me.CHKMonth.Size = New System.Drawing.Size(64, 16)
        Me.CHKMonth.TabIndex = 0
        Me.CHKMonth.Text = "Month"
        '
        'FrmMonthView
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FloralWhite
        Me.ClientSize = New System.Drawing.Size(304, 262)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "FrmMonthView"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Monthly Report"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "CONSTANT"
    Dim DT As New DataTable
    Dim StrSQL As String
    Dim StrSQLPRT As String
    Dim oldrow As Integer
#End Region

#Region "COMBOBOX"
    Sub LoadLoc()
        Dim dtLoc As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        StrSQL = "SELECT  * "
        StrSQL &= "  FROM  TBLDepartment  "
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
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
    Sub LoadType()
        Dim dtType As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        StrSQL = "SELECT  * "
        StrSQL &= "  FROM  TBLType "
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
    Sub LoadMType()
        Dim dtmType As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        StrSQL = "SELECT  * "
        StrSQL &= "  FROM  TBLTypematerial "
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtmType = New DataTable
            DA.Fill(dtmType)
        Catch
        Finally
        End Try
        dtmType.TableName = TBL_MType
        GrdDVMType = dtmType.DefaultView
        '************************************
        cmbMType.DisplayMember = "MaterialName"
        cmbMType.ValueMember = "MaterialCode"
        cmbMType.DataSource = dtmType
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
#End Region

    Private Sub FrmMonthView_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadLoc()
        LoadType()
        LoadMType()
    End Sub

    Private Sub CmdView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdView.Click
        If SelectRaw.Checked Then
            Dim fview As New FrmReportScrap
            fview.sTrxPeriod = "ML"
            If CHKType.Checked Then
                fview.sType = cmbType.SelectedValue
            Else
                fview.sType = ""
            End If
           
            If CHKMonth.Checked Then
                Dim sdate() As String
                sdate = Split(DPeriod.Text.Trim, "/")
                fview.sPeriod1 = sdate(1) + sdate(0)
                fview.sPeriod2 = sdate(1) + sdate(0)
            Else
                fview.sPeriod1 = ""
                fview.sPeriod2 = ""
            End If
            If CHKLoc.Checked Then
                fview.sLoc = cmbLoc.SelectedValue
                fview.sSec = cmbLoc.Text.Trim
            Else
                fview.sLoc = ""
                fview.sSec = ""
            End If

            If CHKMType.Checked Then
                fview.sMType = cmbMType.SelectedValue
                fview.sName = cmbMType.Text.Trim
            Else
                fview.sMType = ""
                fview.sName = ""
            End If

            If CHKCODE.Checked Then
                fview.sCODE = TXTCODE.Text.Trim
            Else
                fview.sCODE = ""
            End If
            fview.sTrxPeriod = "ML"
            fview.ShowDialog()
        End If
        
    End Sub

End Class
