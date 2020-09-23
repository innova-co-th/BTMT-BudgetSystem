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

Public Class FrmView
    Inherits System.Windows.Forms.Form
#Region "Declare"
    Dim GrdDV As New DataView
    Dim GrdDVRM As New DataView
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
    Friend WithEvents GroupBoxCondition As System.Windows.Forms.GroupBox
    Friend WithEvents lblReport As System.Windows.Forms.Label
    Friend WithEvents lblTagNoFrom As System.Windows.Forms.Label
    Friend WithEvents lblTagNoTo As System.Windows.Forms.Label
    Friend WithEvents cmbSection1 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbType As System.Windows.Forms.ComboBox
    Friend WithEvents cmbTypeMaterial As System.Windows.Forms.ComboBox
    Friend WithEvents CmdClose As System.Windows.Forms.Button
    Friend WithEvents CmdView As System.Windows.Forms.Button
    Friend WithEvents DTPYear As System.Windows.Forms.DateTimePicker
    Friend WithEvents RBSec As System.Windows.Forms.RadioButton
    Friend WithEvents RBFrist As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBoxPeriod As System.Windows.Forms.GroupBox
    Friend WithEvents PanelPeriod As System.Windows.Forms.Panel
    Friend WithEvents CHKYear As System.Windows.Forms.RadioButton
    Friend WithEvents CHKSection As System.Windows.Forms.CheckBox
    Friend WithEvents CHKMatCode As System.Windows.Forms.CheckBox
    Friend WithEvents CHKTAG As System.Windows.Forms.CheckBox
    Friend WithEvents RBMat As System.Windows.Forms.RadioButton
    Friend WithEvents RBRAW As System.Windows.Forms.RadioButton
    Friend WithEvents lblType1 As System.Windows.Forms.Label
    Friend WithEvents cmbCode As System.Windows.Forms.ComboBox
    Friend WithEvents CHKMType As System.Windows.Forms.CheckBox
    Friend WithEvents CHKType As System.Windows.Forms.CheckBox
    Friend WithEvents TxtNo2 As System.Windows.Forms.TextBox
    Friend WithEvents TxtNo1 As System.Windows.Forms.TextBox
    Friend WithEvents CHKWIP As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmView))
        Me.GroupBoxCondition = New System.Windows.Forms.GroupBox()
        Me.CHKWIP = New System.Windows.Forms.CheckBox()
        Me.lblType1 = New System.Windows.Forms.Label()
        Me.CHKMType = New System.Windows.Forms.CheckBox()
        Me.CHKTAG = New System.Windows.Forms.CheckBox()
        Me.CHKMatCode = New System.Windows.Forms.CheckBox()
        Me.CHKType = New System.Windows.Forms.CheckBox()
        Me.CHKSection = New System.Windows.Forms.CheckBox()
        Me.RBMat = New System.Windows.Forms.RadioButton()
        Me.RBRAW = New System.Windows.Forms.RadioButton()
        Me.CmdClose = New System.Windows.Forms.Button()
        Me.CmdView = New System.Windows.Forms.Button()
        Me.cmbTypeMaterial = New System.Windows.Forms.ComboBox()
        Me.TxtNo2 = New System.Windows.Forms.TextBox()
        Me.TxtNo1 = New System.Windows.Forms.TextBox()
        Me.cmbCode = New System.Windows.Forms.ComboBox()
        Me.cmbType = New System.Windows.Forms.ComboBox()
        Me.cmbSection1 = New System.Windows.Forms.ComboBox()
        Me.lblTagNoTo = New System.Windows.Forms.Label()
        Me.lblTagNoFrom = New System.Windows.Forms.Label()
        Me.lblReport = New System.Windows.Forms.Label()
        Me.DTPYear = New System.Windows.Forms.DateTimePicker()
        Me.RBSec = New System.Windows.Forms.RadioButton()
        Me.RBFrist = New System.Windows.Forms.RadioButton()
        Me.GroupBoxPeriod = New System.Windows.Forms.GroupBox()
        Me.PanelPeriod = New System.Windows.Forms.Panel()
        Me.CHKYear = New System.Windows.Forms.RadioButton()
        Me.GroupBoxCondition.SuspendLayout()
        Me.GroupBoxPeriod.SuspendLayout()
        Me.PanelPeriod.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBoxCondition
        '
        Me.GroupBoxCondition.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBoxCondition.Controls.Add(Me.CHKWIP)
        Me.GroupBoxCondition.Controls.Add(Me.lblType1)
        Me.GroupBoxCondition.Controls.Add(Me.CHKMType)
        Me.GroupBoxCondition.Controls.Add(Me.CHKTAG)
        Me.GroupBoxCondition.Controls.Add(Me.CHKMatCode)
        Me.GroupBoxCondition.Controls.Add(Me.CHKType)
        Me.GroupBoxCondition.Controls.Add(Me.CHKSection)
        Me.GroupBoxCondition.Controls.Add(Me.RBMat)
        Me.GroupBoxCondition.Controls.Add(Me.RBRAW)
        Me.GroupBoxCondition.Controls.Add(Me.CmdClose)
        Me.GroupBoxCondition.Controls.Add(Me.CmdView)
        Me.GroupBoxCondition.Controls.Add(Me.cmbTypeMaterial)
        Me.GroupBoxCondition.Controls.Add(Me.TxtNo2)
        Me.GroupBoxCondition.Controls.Add(Me.TxtNo1)
        Me.GroupBoxCondition.Controls.Add(Me.cmbCode)
        Me.GroupBoxCondition.Controls.Add(Me.cmbType)
        Me.GroupBoxCondition.Controls.Add(Me.cmbSection1)
        Me.GroupBoxCondition.Controls.Add(Me.lblTagNoTo)
        Me.GroupBoxCondition.Controls.Add(Me.lblTagNoFrom)
        Me.GroupBoxCondition.Controls.Add(Me.lblReport)
        Me.GroupBoxCondition.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBoxCondition.Location = New System.Drawing.Point(8, 104)
        Me.GroupBoxCondition.Name = "GroupBoxCondition"
        Me.GroupBoxCondition.Size = New System.Drawing.Size(817, 238)
        Me.GroupBoxCondition.TabIndex = 1
        Me.GroupBoxCondition.TabStop = False
        '
        'CHKWIP
        '
        Me.CHKWIP.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.CHKWIP.Location = New System.Drawing.Point(344, 24)
        Me.CHKWIP.Name = "CHKWIP"
        Me.CHKWIP.Size = New System.Drawing.Size(72, 16)
        Me.CHKWIP.TabIndex = 45
        Me.CHKWIP.Text = " WIP"
        '
        'lblType1
        '
        Me.lblType1.Location = New System.Drawing.Point(344, 88)
        Me.lblType1.Name = "lblType1"
        Me.lblType1.Size = New System.Drawing.Size(296, 16)
        Me.lblType1.TabIndex = 44
        '
        'CHKMType
        '
        Me.CHKMType.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.CHKMType.Location = New System.Drawing.Point(320, 58)
        Me.CHKMType.Name = "CHKMType"
        Me.CHKMType.Size = New System.Drawing.Size(115, 16)
        Me.CHKMType.TabIndex = 43
        Me.CHKMType.Text = "Material  Type"
        Me.CHKMType.Visible = False
        '
        'CHKTAG
        '
        Me.CHKTAG.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.CHKTAG.Location = New System.Drawing.Point(16, 120)
        Me.CHKTAG.Name = "CHKTAG"
        Me.CHKTAG.Size = New System.Drawing.Size(80, 16)
        Me.CHKTAG.TabIndex = 8
        Me.CHKTAG.Text = "TagNo."
        '
        'CHKMatCode
        '
        Me.CHKMatCode.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.CHKMatCode.Location = New System.Drawing.Point(16, 88)
        Me.CHKMatCode.Name = "CHKMatCode"
        Me.CHKMatCode.Size = New System.Drawing.Size(107, 16)
        Me.CHKMatCode.TabIndex = 6
        Me.CHKMatCode.Text = "Material Code"
        '
        'CHKType
        '
        Me.CHKType.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.CHKType.Location = New System.Drawing.Point(16, 56)
        Me.CHKType.Name = "CHKType"
        Me.CHKType.Size = New System.Drawing.Size(131, 16)
        Me.CHKType.TabIndex = 3
        Me.CHKType.Text = "Type"
        '
        'CHKSection
        '
        Me.CHKSection.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.CHKSection.Location = New System.Drawing.Point(16, 24)
        Me.CHKSection.Name = "CHKSection"
        Me.CHKSection.Size = New System.Drawing.Size(80, 16)
        Me.CHKSection.TabIndex = 0
        Me.CHKSection.Text = "Section"
        '
        'RBMat
        '
        Me.RBMat.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.RBMat.Location = New System.Drawing.Point(320, 152)
        Me.RBMat.Name = "RBMat"
        Me.RBMat.Size = New System.Drawing.Size(144, 24)
        Me.RBMat.TabIndex = 13
        Me.RBMat.Text = "Report by  Material"
        '
        'RBRAW
        '
        Me.RBRAW.Checked = True
        Me.RBRAW.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.RBRAW.Location = New System.Drawing.Point(160, 152)
        Me.RBRAW.Name = "RBRAW"
        Me.RBRAW.Size = New System.Drawing.Size(160, 24)
        Me.RBRAW.TabIndex = 12
        Me.RBRAW.TabStop = True
        Me.RBRAW.Text = "Report by R/M Material"
        '
        'CmdClose
        '
        Me.CmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdClose.Image = CType(resources.GetObject("CmdClose.Image"), System.Drawing.Image)
        Me.CmdClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdClose.Location = New System.Drawing.Point(721, 165)
        Me.CmdClose.Name = "CmdClose"
        Me.CmdClose.Size = New System.Drawing.Size(80, 56)
        Me.CmdClose.TabIndex = 15
        Me.CmdClose.Text = "Close"
        Me.CmdClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdView
        '
        Me.CmdView.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdView.Image = CType(resources.GetObject("CmdView.Image"), System.Drawing.Image)
        Me.CmdView.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdView.Location = New System.Drawing.Point(641, 165)
        Me.CmdView.Name = "CmdView"
        Me.CmdView.Size = New System.Drawing.Size(80, 56)
        Me.CmdView.TabIndex = 14
        Me.CmdView.Text = "View"
        Me.CmdView.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'cmbTypeMaterial
        '
        Me.cmbTypeMaterial.Location = New System.Drawing.Point(464, 56)
        Me.cmbTypeMaterial.Name = "cmbTypeMaterial"
        Me.cmbTypeMaterial.Size = New System.Drawing.Size(136, 21)
        Me.cmbTypeMaterial.TabIndex = 5
        Me.cmbTypeMaterial.Text = "Select"
        Me.cmbTypeMaterial.Visible = False
        '
        'TxtNo2
        '
        Me.TxtNo2.Location = New System.Drawing.Point(376, 120)
        Me.TxtNo2.Name = "TxtNo2"
        Me.TxtNo2.Size = New System.Drawing.Size(96, 20)
        Me.TxtNo2.TabIndex = 10
        '
        'TxtNo1
        '
        Me.TxtNo1.Location = New System.Drawing.Point(160, 120)
        Me.TxtNo1.Name = "TxtNo1"
        Me.TxtNo1.Size = New System.Drawing.Size(96, 20)
        Me.TxtNo1.TabIndex = 9
        '
        'cmbCode
        '
        Me.cmbCode.Location = New System.Drawing.Point(160, 88)
        Me.cmbCode.Name = "cmbCode"
        Me.cmbCode.Size = New System.Drawing.Size(168, 21)
        Me.cmbCode.TabIndex = 7
        Me.cmbCode.Text = "Select"
        '
        'cmbType
        '
        Me.cmbType.Location = New System.Drawing.Point(160, 56)
        Me.cmbType.Name = "cmbType"
        Me.cmbType.Size = New System.Drawing.Size(136, 21)
        Me.cmbType.TabIndex = 4
        Me.cmbType.Text = "Select"
        '
        'cmbSection1
        '
        Me.cmbSection1.Location = New System.Drawing.Point(160, 22)
        Me.cmbSection1.Name = "cmbSection1"
        Me.cmbSection1.Size = New System.Drawing.Size(168, 21)
        Me.cmbSection1.TabIndex = 1
        Me.cmbSection1.Text = "Select"
        '
        'lblTagNoTo
        '
        Me.lblTagNoTo.Location = New System.Drawing.Point(344, 120)
        Me.lblTagNoTo.Name = "lblTagNoTo"
        Me.lblTagNoTo.Size = New System.Drawing.Size(24, 16)
        Me.lblTagNoTo.TabIndex = 15
        Me.lblTagNoTo.Text = "To"
        '
        'lblTagNoFrom
        '
        Me.lblTagNoFrom.Location = New System.Drawing.Point(120, 120)
        Me.lblTagNoFrom.Name = "lblTagNoFrom"
        Me.lblTagNoFrom.Size = New System.Drawing.Size(40, 16)
        Me.lblTagNoFrom.TabIndex = 12
        Me.lblTagNoFrom.Text = "from "
        '
        'lblReport
        '
        Me.lblReport.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblReport.Location = New System.Drawing.Point(32, 152)
        Me.lblReport.Name = "lblReport"
        Me.lblReport.Size = New System.Drawing.Size(80, 16)
        Me.lblReport.TabIndex = 11
        Me.lblReport.Text = "Report"
        '
        'DTPYear
        '
        Me.DTPYear.CustomFormat = "yyyy"
        Me.DTPYear.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DTPYear.Location = New System.Drawing.Point(144, 8)
        Me.DTPYear.Name = "DTPYear"
        Me.DTPYear.ShowUpDown = True
        Me.DTPYear.Size = New System.Drawing.Size(64, 20)
        Me.DTPYear.TabIndex = 41
        '
        'RBSec
        '
        Me.RBSec.Location = New System.Drawing.Point(72, 8)
        Me.RBSec.Name = "RBSec"
        Me.RBSec.Size = New System.Drawing.Size(64, 24)
        Me.RBSec.TabIndex = 40
        Me.RBSec.Text = "Second"
        '
        'RBFrist
        '
        Me.RBFrist.Checked = True
        Me.RBFrist.Location = New System.Drawing.Point(8, 8)
        Me.RBFrist.Name = "RBFrist"
        Me.RBFrist.Size = New System.Drawing.Size(48, 24)
        Me.RBFrist.TabIndex = 39
        Me.RBFrist.TabStop = True
        Me.RBFrist.Text = "Frist"
        '
        'GroupBoxPeriod
        '
        Me.GroupBoxPeriod.Controls.Add(Me.PanelPeriod)
        Me.GroupBoxPeriod.Controls.Add(Me.CHKYear)
        Me.GroupBoxPeriod.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBoxPeriod.Location = New System.Drawing.Point(8, 8)
        Me.GroupBoxPeriod.Name = "GroupBoxPeriod"
        Me.GroupBoxPeriod.Size = New System.Drawing.Size(675, 67)
        Me.GroupBoxPeriod.TabIndex = 0
        Me.GroupBoxPeriod.TabStop = False
        '
        'PanelPeriod
        '
        Me.PanelPeriod.Controls.Add(Me.RBFrist)
        Me.PanelPeriod.Controls.Add(Me.RBSec)
        Me.PanelPeriod.Controls.Add(Me.DTPYear)
        Me.PanelPeriod.Location = New System.Drawing.Point(160, 16)
        Me.PanelPeriod.Name = "PanelPeriod"
        Me.PanelPeriod.Size = New System.Drawing.Size(312, 40)
        Me.PanelPeriod.TabIndex = 44
        '
        'CHKYear
        '
        Me.CHKYear.Checked = True
        Me.CHKYear.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.CHKYear.Location = New System.Drawing.Point(16, 28)
        Me.CHKYear.Name = "CHKYear"
        Me.CHKYear.Size = New System.Drawing.Size(104, 16)
        Me.CHKYear.TabIndex = 0
        Me.CHKYear.TabStop = True
        Me.CHKYear.Text = "Period Year"
        '
        'FrmView
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FloralWhite
        Me.ClientSize = New System.Drawing.Size(832, 357)
        Me.Controls.Add(Me.GroupBoxPeriod)
        Me.Controls.Add(Me.GroupBoxCondition)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmView"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "PHYSICAL INVENTORY  REPORT"
        Me.GroupBoxCondition.ResumeLayout(False)
        Me.GroupBoxCondition.PerformLayout()
        Me.GroupBoxPeriod.ResumeLayout(False)
        Me.PanelPeriod.ResumeLayout(False)
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
    Sub LoadLoc1()
        Dim dtLoc As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        StrSQL = "SELECT  *  FROM  TBLDepartment  "
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
        cmbSection1.DisplayMember = "DeptName"
        cmbSection1.ValueMember = "DeptCode"
        cmbSection1.DataSource = dtLoc
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadType()
        Dim dtType As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        StrSQL = "SELECT  *  FROM  TBLType "
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
    Sub LoadMType(ByVal type As String)
        Dim dtMType As DataTable = New DataTable()
        Dim sb As New System.Text.StringBuilder()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sb.AppendLine("  SELECT distinct Materialcode,MaterialName,Typecode")
        sb.AppendLine("  FROM (")
        sb.AppendLine("    SELECT distinct psemicode code,rtrim(MaterialType) MaterialType,'04' Typecode")
        sb.AppendLine("    FROM TBLPresemi")
        sb.AppendLine("    WHERE active = '1'")
        sb.AppendLine("    UNION")
        sb.AppendLine("    SELECT distinct semicode code,rtrim(MaterialType) MaterialType,'05' Typecode")
        sb.AppendLine("    FROM TBLsemi")
        sb.AppendLine("    WHERE active = '1'")
        sb.AppendLine("  ) a")
        sb.AppendLine("  LEFT OUTER JOIN TBLTypeMaterial  b on a.MaterialType = b.Materialcode")

        If CHKType.Checked Then
            sb.AppendLine("  WHERE typecode = '" & type.Trim() & "'")
        End If

        sb.AppendLine("  ORDER BY Typecode")
        StrSQL = sb.ToString()

        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtMType = New DataTable
            DA.Fill(dtMType)
        Catch
        Finally
        End Try
        dtMType.TableName = TBL_MType
        GrdDVMType = dtMType.DefaultView
        '************************************
        cmbTypeMaterial.DisplayMember = "MaterialName"
        cmbTypeMaterial.ValueMember = "MaterialCode"
        cmbTypeMaterial.DataSource = dtMType
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadMaterial()
        Dim dtRM As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = " SELECT * FROM TBLGroup"
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtRM = New DataTable
            DA.Fill(dtRM)
        Catch
        Finally
        End Try
        dtRM.TableName = TBL_RM
        GrdDVRM = dtRM.DefaultView
        '************************************
        cmbCode.DisplayMember = "Code"
        cmbCode.ValueMember = "Code"
        cmbCode.DataSource = GrdDVRM
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
#End Region

#Region "Form Event"
    Private Sub FrmView_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Comment out by Beam 02-Sep-2020
        'DPeriod1.Value = Now.Date
        'DPeriod2.Value = Now.Date
    End Sub
#End Region

#Region "Control Event"
    Private Sub CmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClose.Click
        Me.Close()
    End Sub

    Private Sub CmdView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdView.Click
        '//Comment out by Beam 02-Sep-2020
        'If CHKMonth.Checked Then
        '    selectMonth()
        'End If
        If CHKYear.Checked Then
            selectYear()
        End If
    End Sub

    Private Sub CHKSection_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CHKSection.CheckedChanged
        If CHKSection.Checked Then
            CHKWIP.Checked = False
            LoadLoc1()
        Else
            cmbSection1.Text = "Select"
        End If

    End Sub

    Private Sub CHKWIP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CHKWIP.CheckedChanged
        If CHKWIP.Checked Then
            CHKSection.Checked = False
        End If
    End Sub

    Private Sub CHKType_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CHKType.CheckedChanged
        If CHKType.Checked Then
            LoadType()
        Else
            cmbType.Text = "Select"
            lblType1.Text = ""
        End If
    End Sub

    Private Sub cmbType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbType.SelectedIndexChanged
        lblType1.Text = cmbType.Text.Trim
        If CHKMatCode.Checked Then
            LoadMaterial()
            GrdDVRM.RowFilter = " typecode = '" & cmbType.SelectedValue & "'"
            cmbCode.DisplayMember = "Code"
            cmbCode.ValueMember = "Code"
            cmbCode.DataSource = GrdDVRM
        Else
            cmbCode.Text = "Select"
        End If

        If cmbType.SelectedValue = "04" Or
        cmbType.SelectedValue = "05" Then
            CHKMType.Visible = True
            cmbTypeMaterial.Visible = True
        Else
            CHKMType.Visible = False
            cmbTypeMaterial.Visible = False
        End If

        If CHKMType.Checked Then
            LoadMType(cmbType.SelectedValue)
        Else
            cmbTypeMaterial.Text = "Select"
        End If
    End Sub

    Private Sub CHKMType_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CHKMType.CheckedChanged
        If CHKMType.Checked Then
            CHKMatCode.Enabled = False
            cmbCode.Enabled = False
            LoadMType(cmbType.SelectedValue)
        Else
            CHKMatCode.Enabled = True
            cmbCode.Enabled = True
            cmbTypeMaterial.Text = "Select"
        End If
    End Sub

    Private Sub cmbTypeMaterial_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbTypeMaterial.SelectedIndexChanged
        If CHKMatCode.Checked Then
            LoadMaterial()
        Else
            cmbCode.Text = "Select"
        End If
    End Sub

    Private Sub CHKMatCode_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CHKMatCode.CheckedChanged
        If CHKMatCode.Checked Then
            LoadMaterial()
            If CHKType.Checked Then
                GrdDVRM.RowFilter = " typecode = '" & cmbType.SelectedValue & "'"
                cmbCode.DisplayMember = "Code"
                cmbCode.ValueMember = "Code"
                cmbCode.DataSource = GrdDVRM
            Else
            End If
        Else
            cmbCode.Text = "Select"
        End If
    End Sub

    Private Sub CHKTAG_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CHKTAG.CheckedChanged
        Dim i, j As Integer
        If CHKTAG.Checked Then
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

    '//Comment out by Beam 02-Sep-2020
    'Private Sub DPeriod2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If DPeriod2.ToString < DPeriod1.ToString Then
    '        MsgBox("Can't Select. Check Data Again.", MsgBoxStyle.OkOnly)
    '    Else
    '    End If
    'End Sub
#End Region

#Region "Sub"
    Sub selectMonth()
        Dim fview As New FrmReportScrap
        '//Comment out by Beam 02-Sep-2020
        'Select Period
        'If DPeriod2.ToString < DPeriod1.ToString Then
        '    MsgBox("Can't Select. Check Data Again.", MsgBoxStyle.OkOnly)
        '    Exit Sub
        'Else
        'End If

        'Adjust Tag 
        If CHKTAG.Checked Then
            Dim i, j As Integer
            i = TxtNo1.Text.Trim()
            j = TxtNo2.Text.Trim()
            TxtNo1.Text = Format(i, "0000")
            TxtNo2.Text = Format(j, "0000")
            fview.sTrx1 = Format(i, "0000")
            fview.sTrx2 = Format(j, "0000")
        Else
            fview.sTrx1 = ""
            fview.sTrx2 = ""
        End If

        '//Comment out by Beam 02-Sep-2020
        'Select format of report between by month or by half year
        'If CHKMonth.Checked Then
        '    fview.sTrxPeriod = "ML"
        'Else
        fview.sTrxPeriod = ""
        'End If

        'Report by R/M Material by display each material
        If RBRAW.Checked Then
            If CHKType.Checked Then
                fview.sType = cmbType.SelectedValue
            Else
                fview.sType = ""
            End If
            If CHKMType.Checked Then
                fview.sMType = cmbTypeMaterial.SelectedValue
                fview.sName = cmbTypeMaterial.Text.Trim
            Else
                fview.sMType = ""
                fview.sName = ""
            End If
            '//Comment out by Beam 02-Sep-2020
            'If CHKMonth.Checked Then
            '    Dim sdate1(), sdate2() As String
            '    sdate1 = Split(DPeriod1.Text.Trim, "/")
            '    sdate2 = Split(DPeriod2.Text.Trim, "/")
            '    fview.sPeriod1 = sdate1(1) + sdate1(0)
            '    fview.sPeriod2 = sdate2(1) + sdate2(0)
            'Else
            fview.sPeriod1 = ""
            fview.sPeriod2 = ""
            'End If
            If CHKSection.Checked Then
                fview.sLoc = cmbSection1.SelectedValue
                fview.sSec = cmbSection1.Text.Trim
            Else
                fview.sLoc = ""
                fview.sSec = ""
            End If
            If CHKMatCode.Checked Then
                fview.sCODE = cmbCode.Text.Trim
            Else
                fview.sCODE = ""
            End If
            If CHKTAG.Checked Then
                fview.sTag1 = TxtNo1.Text.Trim
                fview.sTag2 = TxtNo2.Text.Trim
            Else
                fview.sTag1 = ""
                fview.sTag2 = ""
            End If
            fview.Show()
        End If
        If RBMat.Checked Then
            MsgBox("don't Have Scarp Report.", MsgBoxStyle.OkOnly)
        End If

    End Sub
    Sub selectYear()
        Dim fview As New FrmPHYReport
        'Adjust Tag 
        If CHKTAG.Checked Then
            Dim i, j As Integer
            i = TxtNo1.Text.Trim
            j = TxtNo2.Text.Trim
            TxtNo1.Text = Format(i, "0000")
            TxtNo2.Text = Format(j, "0000")
            fview.sTrx1 = Format(i, "0000")
            fview.sTrx2 = Format(j, "0000")
        Else
            fview.sTrx1 = ""
            fview.sTrx2 = ""
        End If

        'Select format of report between by month or by half year
        If CHKYear.Checked Then
            fview.sTrxPeriod = "YL"
        Else
            fview.sTrxPeriod = ""
        End If

        'Report by R/M Material by display each material
        If RBRAW.Checked Then
            fview.sName = ""
            If CHKType.Checked Then
                fview.sType = cmbType.SelectedValue
                fview.sName = cmbType.Text.Trim
            Else
                fview.sType = ""
            End If
            If CHKMType.Checked Then
                fview.sMType = cmbTypeMaterial.SelectedValue
                fview.sName = cmbTypeMaterial.Text.Trim & "  "
            Else
                fview.sMType = ""
            End If
            If CHKType.Checked = False And CHKMType.Checked = False Then
                fview.sName = " All PROCESS"
            End If
            fview.sPeriod1 = ""
            fview.sPeriod2 = ""
            If RBFrist.Checked Then
                fview.sHeader = " 1st HALF'" & DTPYear.Text
                fview.sMonth = " JUNE '" & DTPYear.Text
                fview.sPeriod1 = DTPYear.Text.Trim & "01"
                fview.sPeriod2 = DTPYear.Text.Trim & "01"
            End If
            If RBSec.Checked Then
                fview.sHeader = " 2nd HALF'" & DTPYear.Text
                fview.sMonth = " DECEMBER '" & DTPYear.Text
                fview.sPeriod1 = DTPYear.Text & "02"
                fview.sPeriod2 = DTPYear.Text & "02"
            End If
            If CHKSection.Checked Then
                fview.sLoc = cmbSection1.SelectedValue
                fview.sSec = cmbSection1.Text.Trim
            Else
                fview.sLoc = ""
                fview.sSec = "ALL PRODUCTION"
            End If
            If CHKWIP.Checked Then
                fview.sLoc2 = "WIP"
                fview.sSec = "WIP"
            Else
                fview.sLoc2 = ""
            End If
            If CHKMatCode.Checked Then
                fview.sCODE = cmbCode.Text.Trim
            Else
                fview.sCODE = ""
            End If
            If CHKTAG.Checked Then
                fview.sTag1 = TxtNo1.Text.Trim
                fview.sTag2 = TxtNo2.Text.Trim
            Else
                fview.sTag1 = ""
                fview.sTag2 = ""
            End If
            fview.Show()
        End If

        Dim fmview As New FrmPHYReportMaterial
        'Select format of report between by month or by half year
        If CHKYear.Checked Then
            fmview.sTrxPeriod = "YL"
        Else
            fmview.sTrxPeriod = ""
        End If

        ' Report by material 
        If RBMat.Checked Then
            fmview.sName = ""
            If CHKType.Checked Then
                fmview.sType = cmbType.SelectedValue
                fmview.sName = cmbType.Text.Trim
            Else
                fmview.sType = ""
            End If
            If CHKMType.Checked Then
                fmview.sMType = cmbTypeMaterial.SelectedValue
                fmview.sName = cmbTypeMaterial.Text.Trim & "  "
            Else
                fmview.sMType = ""
            End If
            If CHKType.Checked = False And CHKMType.Checked = False Then
                fmview.sName &= " All PROCESS"
            End If
            fmview.sPeriod1 = ""
            fmview.sPeriod2 = ""
            If RBFrist.Checked Then
                fmview.sHeader = " 1St HALF'" & DTPYear.Text
                fmview.sMonth = " JUNE '" & DTPYear.Text
                fmview.sPeriod1 = DTPYear.Text & "01"
                fmview.sPeriod2 = DTPYear.Text & "01"
            End If
            If RBSec.Checked Then
                fmview.sHeader = " 2ND HALF'" & DTPYear.Text
                fmview.sMonth = " DECEMBER '" & DTPYear.Text
                fmview.sPeriod1 = DTPYear.Text & "02"
                fmview.sPeriod2 = DTPYear.Text & "02"
            End If
            If CHKSection.Checked Then
                fmview.sLoc = cmbSection1.SelectedValue
                fmview.sSec = cmbSection1.Text.Trim
                fmview.sIdSec = cmbSection1.SelectedValue
            Else
                fmview.sLoc = ""
                fmview.sSec = "ALL PRODUCTION"
                fmview.sIdSec = ""
            End If
            If CHKWIP.Checked Then
                fmview.sLoc2 = "WIP"
                fmview.sSec = "WIP"
            Else
                fmview.sLoc2 = ""
            End If
            If CHKMatCode.Checked Then
                fmview.sCODE = cmbCode.Text.Trim
            Else
                fmview.sCODE = ""
            End If
            If CHKTAG.Checked Then
                fmview.sTag1 = TxtNo1.Text.Trim
                fmview.sTag2 = TxtNo2.Text.Trim
            Else
                fmview.sTag1 = ""
                fmview.sTag2 = ""
            End If
            fmview.ShowDialog()
        End If
    End Sub
#End Region
End Class
