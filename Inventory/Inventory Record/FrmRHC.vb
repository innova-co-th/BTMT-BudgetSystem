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

Public Class FrmRHC

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
    Friend WithEvents CmdImport As System.Windows.Forms.Button
    Friend WithEvents CmdExport As System.Windows.Forms.Button
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ComboBoxStage As System.Windows.Forms.ComboBox
    Friend WithEvents CmdView As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmRHC))
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
        Me.ComboBoxStage = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.CmdView = New System.Windows.Forms.Button
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
        Me.GroupBox1.Size = New System.Drawing.Size(986, 504)
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
        Me.DataGridCOM.Size = New System.Drawing.Size(980, 485)
        Me.DataGridCOM.TabIndex = 0
        '
        'CmdSave
        '
        Me.CmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdSave.Image = CType(resources.GetObject("CmdSave.Image"), System.Drawing.Image)
        Me.CmdSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdSave.Location = New System.Drawing.Point(792, 568)
        Me.CmdSave.Name = "CmdSave"
        Me.CmdSave.Size = New System.Drawing.Size(64, 56)
        Me.CmdSave.TabIndex = 1
        Me.CmdSave.Text = "Add"
        Me.CmdSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdClose
        '
        Me.CmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdClose.Image = CType(resources.GetObject("CmdClose.Image"), System.Drawing.Image)
        Me.CmdClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdClose.Location = New System.Drawing.Point(920, 568)
        Me.CmdClose.Name = "CmdClose"
        Me.CmdClose.Size = New System.Drawing.Size(64, 56)
        Me.CmdClose.TabIndex = 2
        Me.CmdClose.Text = "Close"
        Me.CmdClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdEdit
        '
        Me.CmdEdit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdEdit.Image = CType(resources.GetObject("CmdEdit.Image"), System.Drawing.Image)
        Me.CmdEdit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdEdit.Location = New System.Drawing.Point(856, 568)
        Me.CmdEdit.Name = "CmdEdit"
        Me.CmdEdit.Size = New System.Drawing.Size(64, 56)
        Me.CmdEdit.TabIndex = 3
        Me.CmdEdit.Text = "Edit"
        Me.CmdEdit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmbCompound
        '
        Me.CmbCompound.Enabled = False
        Me.CmbCompound.Location = New System.Drawing.Point(96, 40)
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
        Me.CheckBoxGP.Location = New System.Drawing.Point(256, 10)
        Me.CheckBoxGP.Name = "CheckBoxGP"
        Me.CheckBoxGP.Size = New System.Drawing.Size(72, 16)
        Me.CheckBoxGP.TabIndex = 10
        Me.CheckBoxGP.Text = "Group"
        '
        'CmbGroup
        '
        Me.CmbGroup.Enabled = False
        Me.CmbGroup.Location = New System.Drawing.Point(328, 8)
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
        Me.CmdDelete.Location = New System.Drawing.Point(8, 568)
        Me.CmdDelete.Name = "CmdDelete"
        Me.CmdDelete.Size = New System.Drawing.Size(80, 56)
        Me.CmdDelete.TabIndex = 11
        Me.CmdDelete.Text = "Del"
        Me.CmdDelete.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.CmdDelete.Visible = False
        '
        'ComboBoxStage
        '
        Me.ComboBoxStage.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9"})
        Me.ComboBoxStage.Location = New System.Drawing.Point(96, 8)
        Me.ComboBoxStage.Name = "ComboBoxStage"
        Me.ComboBoxStage.Size = New System.Drawing.Size(72, 21)
        Me.ComboBoxStage.TabIndex = 13
        Me.ComboBoxStage.Text = "Select"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 16)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Stage"
        '
        'CmdView
        '
        Me.CmdView.Image = CType(resources.GetObject("CmdView.Image"), System.Drawing.Image)
        Me.CmdView.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdView.Location = New System.Drawing.Point(488, 8)
        Me.CmdView.Name = "CmdView"
        Me.CmdView.Size = New System.Drawing.Size(75, 56)
        Me.CmdView.TabIndex = 15
        Me.CmdView.Text = "View"
        Me.CmdView.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdImport
        '
        Me.CmdImport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdImport.Image = CType(resources.GetObject("CmdImport.Image"), System.Drawing.Image)
        Me.CmdImport.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdImport.Location = New System.Drawing.Point(636, 568)
        Me.CmdImport.Name = "CmdImport"
        Me.CmdImport.Size = New System.Drawing.Size(64, 56)
        Me.CmdImport.TabIndex = 16
        Me.CmdImport.Text = "Import"
        Me.CmdImport.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdExport
        '
        Me.CmdExport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdExport.Image = CType(resources.GetObject("CmdExport.Image"), System.Drawing.Image)
        Me.CmdExport.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdExport.Location = New System.Drawing.Point(700, 568)
        Me.CmdExport.Name = "CmdExport"
        Me.CmdExport.Size = New System.Drawing.Size(64, 56)
        Me.CmdExport.TabIndex = 17
        Me.CmdExport.Text = "Export"
        Me.CmdExport.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'FrmRHC
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1002, 632)
        Me.Controls.Add(Me.CmdExport)
        Me.Controls.Add(Me.CmdImport)
        Me.Controls.Add(Me.CmdView)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ComboBoxStage)
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
        Me.Name = "FrmRHC"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Compound  (C.RHC)"
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

#Region "Function_Load"
    Private Sub LoadCOM()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        StrSQL = "   select * from "
        StrSQL &= "  ( SELECT  Seq,FinalCompound,CompCode,Revision,FinalCompound cc,RHC,Qty TQty,Active"
        StrSQL &= "   ,CompCode+','+Revision Code"
        StrSQL &= "   ,'' MasterCode,isnull(Revision,'') MRev,'' RMcode,null mRHC,null Qty"
        StrSQL &= "    FROM         TBLCompound"
        StrSQL &= "   Union"
        StrSQL &= "    SELECT Seq,c.Finalcompound,'' code,'' rev,'' cc,null RHC,null TQty,'' Active"
        StrSQL &= "    ,CRev,Code,Rev,RMCode,mRHC,mQty"
        StrSQL &= "    FROM         "
        StrSQL &= "   (SELECT  seq,Finalcompound,compcode,Revision,compcode+','+Revision CRev,RHC,Active from  TBLCompound) c"
        StrSQL &= "    left outer join "
        StrSQL &= "    ( select mastercode code,Revision Rev"
        StrSQL &= "   ,mastercode+','+Revision MRev,RMcode,RHC mRHC,Weight mQty from   TBLRHCDtl"
        StrSQL &= "    where Mastercode in "
        StrSQL &= "    ( SELECT  compcode"
        StrSQL &= "     FROM         TBLCompound))m"
        StrSQL &= "    on c.CRev = m.MRev"
        StrSQL &= "     )aa"
        StrSQL &= "    where seq = " & ComboBoxStage.Text.Trim
        StrSQL &= "   order by code ,TQty desc"

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
            .HeaderText = "R/M Material"
            .MappingName = "RMCode"
            .NullText = ""
            .Width = 125
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle3 As New DataGridColoredLine2
        With grdColStyle3
            .HeaderText = "Total "
            .MappingName = "RHC"
            .Width = 75
            .Format = "##,###,##0.000"
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle3_1 As New DataGridColoredLine2
        With grdColStyle3_1
            .HeaderText = "Total "
            .MappingName = "TQTY"
            .Width = 75
            .Format = "##,###,##0.000"
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle4 As New DataGridColoredLine2
        With grdColStyle4
            .HeaderText = "RHC "
            .MappingName = "mRHC"
            .Width = 75
            .Format = "##,###,##0.000"
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle4_1 As New DataGridColoredLine2
        With grdColStyle4_1
            .HeaderText = "Qty "
            .MappingName = "Qty"
            .Width = 75
            .Format = "##,###,##0.000"
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        grdTableStyle1.GridColumnStyles.AddRange _
    (New DataGridColumnStyle() _
    {grdColStyle0, grdColStyle1, grdColStyle1_1, grdColStyle2, _
     grdColStyle4_1, grdColStyle3_1, grdColStyle4, grdColStyle3})

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

    Private Sub FrmRHC_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ComboBoxStage.SelectedIndex = 1
        LoadCOM()
        Loadgroup()
        LoadCompound()
    End Sub

    Private Sub CmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClose.Click
        Me.Close()
    End Sub

    Private Sub CmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdEdit.Click
        Dim fAddRHC As New FrmAddRHC
        fAddRHC.CmdSave.Text = "Edit"
        fAddRHC.TCompound = GrdDV.Item(oldrow).Row("FinalCompound")
        fAddRHC.TCode = GrdDV.Item(oldrow).Row("CompCode")
        fAddRHC.TRev = GrdDV.Item(oldrow).Row("Revision")
        fAddRHC.TStep = GrdDV.Item(oldrow).Row("seq")
        fAddRHC.CBal = False
        fAddRHC.ShowDialog()
        ' LoadCompound()
        ' Loadgroup()
        LoadCOM()
        'If CheckBoxCompoud.Checked = True And CheckBoxGP.Checked = True Then
        '    GrdDV.RowFilter = " Code like'%" & CmbCompound.Text.Trim & _
        '                      "%' and  Finalcompound like'%" & StrData & "%'"
        '    DataGridCOM.DataSource = GrdDV
        '    CmbCompound.Enabled = True
        'ElseIf CheckBoxCompoud.Checked = True And CheckBoxGP.Checked = False Then
        '    GrdDV.RowFilter = " Code like'%" & CmbCompound.Text.Trim & "%'"
        '    DataGridCOM.DataSource = GrdDV
        '    CmbCompound.Enabled = True
        'ElseIf CheckBoxCompoud.Checked = False And CheckBoxGP.Checked = True Then
        '    GrdDV.RowFilter = " Finalcompound like'%" & StrData & "%'"
        '    DataGridCOM.DataSource = GrdDV
        '    CmbCompound.Enabled = True
        'Else
        '    GrdDV.RowFilter = " "
        '    DataGridCOM.DataSource = GrdDV
        '    CmbCompound.Enabled = False
        'End If
    End Sub

    Private Sub CmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSave.Click
        Dim fAddRHC As New FrmAddRHC
        fAddRHC.CmdSave.Text = "Save"
        fAddRHC.TCompound = GrdDV.Item(oldrow).Row("FinalCompound")
        fAddRHC.TCode = GrdDV.Item(oldrow).Row("CompCode")
        fAddRHC.TRev = GrdDV.Item(oldrow).Row("Revision")
        fAddRHC.TStep = GrdDV.Item(oldrow).Row("seq")
        fAddRHC.CBal = True
        fAddRHC.ShowDialog()
        '   LoadCompound()
        '    Loadgroup()
        LoadCOM()
        chk()
    End Sub

    Sub chk()
        If CheckBoxCompoud.Checked = False And CheckBoxGP.Checked = False Then
            GrdDV.RowFilter = " seq = '" & ComboBoxStage.Text.Trim & "'"
            DataGridCOM.DataSource = GrdDV
        ElseIf CheckBoxCompoud.Checked = False And CheckBoxGP.Checked = True Then
            GrdDV.RowFilter = " Finalcompound like'%" & CmbGroup.Text.Trim & "%'"
            DataGridCOM.DataSource = GrdDV
            StrData = CmbGroup.Text.Trim
        ElseIf CheckBoxCompoud.Checked = True And CheckBoxGP.Checked = False Then
            GrdDV.RowFilter = " Code like'%" & CmbCompound.Text.Trim & "%'"
            DataGridCOM.DataSource = GrdDV
        Else : CheckBoxCompoud.Checked = True And CheckBoxGP.Checked = True
            GrdDV.RowFilter = " Finalcompound like'%" & CmbGroup.Text.Trim & "%'" _
                            & " and Code like'%" & CmbCompound.Text.Trim & "%'"
            DataGridCOM.DataSource = GrdDV
        End If
    End Sub
    Private Sub DataGridCOM_CurrentCellChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGridCOM.CurrentCellChanged
        oldrow = DataGridCOM.CurrentCell.RowNumber
    End Sub

    Private Sub CheckBoxCompoud_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxCompoud.CheckedChanged
        If CheckBoxCompoud.Checked = True And CheckBoxGP.Checked = True Then
            GrdDV.RowFilter = " Code like'%" & CmbCompound.Text.Trim & _
                              "%' and  Finalcompound like'%" & CmbGroup.Text.Trim & "%'"
            DataGridCOM.DataSource = GrdDV
            CmbCompound.Enabled = True
        ElseIf CheckBoxCompoud.Checked = True And CheckBoxGP.Checked = False Then
            GrdDV.RowFilter = " Code like'%" & CmbCompound.Text.Trim & "%'"
            DataGridCOM.DataSource = GrdDV
            CmbCompound.Enabled = True
        Else
            GrdDV.RowFilter = " "
            DataGridCOM.DataSource = GrdDV
            CmbCompound.Enabled = False
        End If
        chk()
    End Sub

    Private Sub CmbCompound_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCompound.SelectedIndexChanged
        If ComboBoxStage.Text <> "Select" Then
            LoadCOM()
        End If

        If CheckBoxCompoud.Checked = True And CheckBoxGP.Checked = True Then
            GrdDV.RowFilter = " Code like'%" & CmbCompound.Text.Trim & _
                              "%' and  Finalcompound like'%" & CmbGroup.Text.Trim & "%'"
            DataGridCOM.DataSource = GrdDV
            CmbCompound.Enabled = True
        ElseIf CheckBoxCompoud.Checked = True And CheckBoxGP.Checked = False Then
            GrdDV.RowFilter = " Code like'%" & CmbCompound.Text.Trim & "%'"
            DataGridCOM.DataSource = GrdDV
            CmbCompound.Enabled = True
        Else
            GrdDV.RowFilter = " "
            DataGridCOM.DataSource = GrdDV
            CmbCompound.Enabled = False
        End If

        chk()
        StrData = CmbGroup.Text.Trim

    End Sub

    Private Sub CheckBoxGP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxGP.CheckedChanged
        If CheckBoxGP.Checked = True Then
            GrdDV.RowFilter = " Finalcompound like'%" & CmbGroup.Text.Trim & "%'"
            DataGridCOM.DataSource = GrdDV
            CmbGroup.Enabled = True
        Else
            GrdDV.RowFilter = " "
            DataGridCOM.DataSource = GrdDV
            CmbGroup.Enabled = False
        End If
        StrData = CmbGroup.Text.Trim
        chk()
    End Sub

    Private Sub CmbGroup_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbGroup.SelectedIndexChanged
        chk()
        StrData = CmbGroup.Text.Trim
    End Sub


#Region "DelCompound"
    Private Function ChkData() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Try
            strSQL &= " select count(*) from TblMaster "
            strSQL &= " where RMCode  = '" & GrdDV.Item(oldrow).Row("CompCode") & "'"
            strSQL &= " and RMRevision  = '" & GrdDV.Item(oldrow).Row("Revision") & "'"
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
    Sub DelCompound()
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Try
            strSQL &= " Delete TBLRHCDtl"
            strSQL &= " where Mastercode = '" & GrdDV.Item(oldrow).Row("CompCode") & "'"
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
#End Region

    Private Sub CmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdDelete.Click
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        msg = "Delete Compound :" & GrdDV.Item(oldrow).Row("CompCode") ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Compound"   ' Define title.

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            If ChkData() Then
                MsgBox("It's have Usage , Can't Delete. Please contact IS.", MsgBoxStyle.Information, "Delete R/M ")
            Else
                DelCompound()
                LoadCompound()

                If CheckBoxGP.Checked = True Then
                    GrdDV.RowFilter = " Finalcompound like'%" & CmbGroup.Text.Trim & "%'"
                    DataGridCOM.DataSource = GrdDV
                    CmbGroup.Enabled = True
                Else
                    GrdDV.RowFilter = " "
                    DataGridCOM.DataSource = GrdDV
                    CmbGroup.Enabled = False
                End If
                StrData = CmbGroup.Text.Trim
            End If
        Else
            Exit Sub
        End If

    End Sub

    Private Sub ComboBoxStage_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxStage.SelectedIndexChanged
        If ComboBoxStage.Text <> "Select" Then
            LoadCOM()
            'Loadgroup()
            'LoadCompound()
            chk()
            StrData = CmbGroup.Text.Trim
        Else
        End If
    End Sub

    Private Sub CmdView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdView.Click
        If ComboBoxStage.Text <> "Select" Then
            LoadCOM()
            ' Loadgroup()
            'LoadCompound()
        Else
        End If

        chk()
        StrData = CmbGroup.Text.Trim

    End Sub
End Class