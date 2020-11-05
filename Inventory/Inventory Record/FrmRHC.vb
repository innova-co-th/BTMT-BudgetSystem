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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.DataGridCOM = New System.Windows.Forms.DataGrid()
        Me.CmdSave = New System.Windows.Forms.Button()
        Me.CmdClose = New System.Windows.Forms.Button()
        Me.CmdEdit = New System.Windows.Forms.Button()
        Me.CmbCompound = New System.Windows.Forms.ComboBox()
        Me.CheckBoxCompoud = New System.Windows.Forms.CheckBox()
        Me.CheckBoxGP = New System.Windows.Forms.CheckBox()
        Me.CmbGroup = New System.Windows.Forms.ComboBox()
        Me.CmdDelete = New System.Windows.Forms.Button()
        Me.ComboBoxStage = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CmdView = New System.Windows.Forms.Button()
        Me.CmdImport = New System.Windows.Forms.Button()
        Me.CmdExport = New System.Windows.Forms.Button()
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
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(1018, 671)
        Me.Name = "FrmRHC"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Compound  (C.RHC) -"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataGridCOM, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "CONSTANT"
    Dim DT As New DataTable() 'Compound Data which is filter by Seq
    Dim DTALL As New DataTable() 'Compound Data
    Dim StrSQL As String
    Dim oldrow As Integer
#End Region

#Region "Function_Load"
    Private Sub LoadCOM()
        Dim sb As New System.Text.StringBuilder()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        'Load all compound data
        sb.AppendLine("SELECT * ")
        sb.AppendLine("FROM ( ")
        sb.AppendLine("  SELECT  Seq,FinalCompound,CompCode,Revision,FinalCompound cc,RHC,Qty TQty,Active")
        sb.AppendLine("  ,CompCode+','+Revision Code")
        sb.AppendLine("  ,'' MasterCode,isnull(Revision,'') MRev,'' RMcode,null mRHC,null Qty,'' as FinalCompound_Code,'' as Revision_No,'' as Seq_No")
        sb.AppendLine("  FROM         TBLCompound")
        sb.AppendLine("  UNION")
        sb.AppendLine("  SELECT c.Seq,c.Finalcompound,'' code,'' rev,'' cc,null RHC,null TQty,'' Active")
        sb.AppendLine("  ,c.CRev,m.Code,m.Rev,m.RMCode,m.mRHC,m.mQty,c.FinalCompound as FinalCompound_Code,c.Revision as Revision_No,c.Seq as Seq_No")
        sb.AppendLine("  FROM (        ")
        sb.AppendLine("    SELECT  seq,Finalcompound,compcode,Revision,compcode+','+Revision CRev,RHC,Active")
        sb.AppendLine("    FROM  TBLCompound")
        sb.AppendLine("  ) c")
        sb.AppendLine("  LEFT OUTER JOIN ( ")
        sb.AppendLine("    SELECT MasterCode Code,Revision Rev")
        sb.AppendLine("    ,mastercode+','+Revision MRev,RMcode,RHC mRHC,Weight mQty")
        sb.AppendLine("    FROM   TBLRHCDtl")
        sb.AppendLine("    WHERE Mastercode in ( SELECT  compcode FROM TBLCompound)")
        sb.AppendLine("  ) m on c.CRev = m.MRev")
        sb.AppendLine(") aa ")
        sb.AppendLine("ORDER BY Code, TQty DESC")
        StrSQL = sb.ToString()

        If Not DTALL Is Nothing Then
            If DTALL.Rows.Count >= 1 Then
                DTALL.Clear()
            End If
        End If

        Dim DAALL As SqlDataAdapter
        Try
            DAALL = New SqlDataAdapter(StrSQL, C1.Strcon)
            DTALL = New DataTable()
            DAALL.Fill(DTALL)
        Catch
            MsgBox("Can't Select Data.", MsgBoxStyle.Critical, "Load Data")
        End Try

        DTALL.TableName = TBL_Comp

        'Load compound data by seq
        sb.Clear()
        sb.AppendLine("SELECT * ")
        sb.AppendLine("FROM ( ")
        sb.AppendLine("  SELECT  Seq,FinalCompound,CompCode,Revision,FinalCompound cc,RHC,Qty TQty,Active")
        sb.AppendLine("  ,CompCode+','+Revision Code")
        sb.AppendLine("  ,'' MasterCode,isnull(Revision,'') MRev,'' RMcode,null mRHC,null Qty,'' as FinalCompound_Code,'' as Revision_No,'' as Seq_No")
        sb.AppendLine("  FROM         TBLCompound")
        sb.AppendLine("  UNION")
        sb.AppendLine("  SELECT c.Seq,c.Finalcompound,'' code,'' rev,'' cc,null RHC,null TQty,'' Active")
        sb.AppendLine("  ,c.CRev,m.Code,m.Rev,m.RMCode,m.mRHC,m.mQty,c.FinalCompound as FinalCompound_Code,c.Revision as Revision_No,c.Seq as Seq_No")
        sb.AppendLine("  FROM (        ")
        sb.AppendLine("    SELECT  seq,Finalcompound,compcode,Revision,compcode+','+Revision CRev,RHC,Active")
        sb.AppendLine("    FROM  TBLCompound")
        sb.AppendLine("  ) c")
        sb.AppendLine("  LEFT OUTER JOIN ( ")
        sb.AppendLine("    SELECT MasterCode Code,Revision Rev")
        sb.AppendLine("    ,mastercode+','+Revision MRev,RMcode,RHC mRHC,Weight mQty")
        sb.AppendLine("    FROM   TBLRHCDtl")
        sb.AppendLine("    WHERE Mastercode in ( SELECT  compcode FROM TBLCompound)")
        sb.AppendLine("  ) m on c.CRev = m.MRev")
        sb.AppendLine(") aa ")
        sb.AppendLine("WHERE Seq = " & ComboBoxStage.Text.Trim())
        sb.AppendLine("ORDER BY Code, TQty DESC")
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
            'Nothing
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
    {grdColStyle0, grdColStyle1, grdColStyle1_1, grdColStyle2,
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

#Region "Form Event"
    Private Sub FrmRHC_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ComboBoxStage.SelectedIndex = 1
        LoadCOM()
        Loadgroup()
        LoadCompound()
        SetTotal() 'Set number of items
    End Sub
#End Region

#Region "Control Event"
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
        Chk()
    End Sub

    Private Sub DataGridCOM_CurrentCellChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGridCOM.CurrentCellChanged
        oldrow = DataGridCOM.CurrentCell.RowNumber
    End Sub

    Private Sub CheckBoxCompoud_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxCompoud.CheckedChanged
        If CheckBoxCompoud.Checked = True And CheckBoxGP.Checked = True Then
            GrdDV.RowFilter = " Code like'%" & CmbCompound.Text.Trim &
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
        Chk()
    End Sub

    Private Sub CmbCompound_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCompound.SelectedIndexChanged
        If ComboBoxStage.Text <> "Select" Then
            LoadCOM()
        End If

        If CheckBoxCompoud.Checked = True And CheckBoxGP.Checked = True Then
            GrdDV.RowFilter = " Code like'%" & CmbCompound.Text.Trim &
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

        Chk()
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
        Chk()
    End Sub

    Private Sub CmbGroup_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbGroup.SelectedIndexChanged
        Chk()
        StrData = CmbGroup.Text.Trim
    End Sub

    Private Sub CmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdDelete.Click
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        msg = "Delete Compound :" & GrdDV.Item(oldrow).Row("CompCode") ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or
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
            Chk()
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

        Chk()
        StrData = CmbGroup.Text.Trim

    End Sub

    Private Sub CmdExport_Click(sender As Object, e As EventArgs) Handles CmdExport.Click
        Dim arrColumn As String() = System.Configuration.ConfigurationManager.AppSettings("EXP_EXCEL_COLUMN_MASTER_COMPOUND_RHC").ToString().Split(New Char() {","c})
        Dim arrColumnHeader As String() = System.Configuration.ConfigurationManager.AppSettings("EXP_EXCEL_COLUMN_HEADER_MASTER_COMPOUND_RHC").ToString().Split(New Char() {","c})
        ExcelLib.Export(Me, GrdDV, TBL_RM, arrColumn, arrColumnHeader)
    End Sub

    Private Sub CmdImport_Click(sender As Object, e As EventArgs) Handles CmdImport.Click
        Dim arrColumn As String() = System.Configuration.ConfigurationManager.AppSettings("IMP_EXCEL_COLUMN_MASTER_COMPOUND_RHC").ToString().Split(New Char() {","c})
        Dim importDialog As OpenFileDialog = New OpenFileDialog With {
            .Filter = System.Configuration.ConfigurationManager.AppSettings("DIALOG_FILE_EXT").ToString()
        }
        Dim dtRec As DataTable
        Dim sb As New System.Text.StringBuilder()
        Dim frmOverlay As New Form()
        Dim totalQty As Double = 0
        Dim totalRHC As Double = 0

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
            dtRec = ExcelLib.Import(importDialog.FileName, Me, GrdDV, TBL_RM, arrColumn)

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
                        Dim chkSamePigmentCode As String = String.Empty

                        '//Sort Data from Excel
                        dtRec.DefaultView.Sort = "FinalCompound_Code DESC, Compound_Code DESC, Revision_No DESC"
                        dtRec = dtRec.DefaultView.ToTable

                        '//**Check All Import Data that all data still in TBLCompound, TBLMaster and TBLRM
                        If ChkRMCodeEachCompound(dtRec) = False Then
                            LoadCOM() 'ReQuery and set datagrid
                            frmOverlay.Dispose()
                            Exit Sub
                        End If

                        For i As Integer = 0 To dtRec.Rows.Count - 1
                            Dim strFinalCompoundCode As String = dtRec.Rows(i)("FinalCompound_Code").ToString().Trim()
                            Dim strCompoundCode As String = dtRec.Rows(i)("Compound_Code").ToString().Trim()
                            Dim strRevision As String = dtRec.Rows(i)("Revision_No").ToString().Trim()
                            Dim strRMCode As String = dtRec.Rows(i)("RMCode").ToString().Trim()

                            'Check empty
                            If strFinalCompoundCode.Length <= 0 Then
                                Throw New System.Exception("Please input Final Compound data.")
                            End If

                            If dtRec.Rows(i)("Seq_No").ToString().Trim().Length <= 0 Then
                                Throw New System.Exception("Please input Seq data.")
                            Else
                                Dim intSeqResult As Integer = 0
                                If Not Int32.TryParse(dtRec.Rows(i)("Seq_No"), intSeqResult) Then
                                    Throw New System.Exception("Please input Seq data as Number.")
                                End If
                            End If

                            'Declare and convert seq to integer
                            Dim intSeq As Integer = dtRec.Rows(i)("Seq_No")

                            'Check empty
                            If strFinalCompoundCode.Length <= 0 Then
                                Throw New System.Exception("Please input Final Compound data.")
                            End If

                            If strCompoundCode.Length <= 0 Then
                                Throw New System.Exception("Please input Compound data.")
                            End If

                            If strRevision.Length <= 0 Then
                                Throw New System.Exception("Please input Revision data.")
                            ElseIf strRevision.Length > 3 Then
                                Throw New System.Exception("Revision data must less than 4 digits.")
                            End If

                            If strFinalCompoundCode.Length > 0 Then
                                Dim dblRMQty As Double
                                Dim dblRHC As Double

                                If dtRec.Rows(i)("Qty").ToString().Length > 0 Then
                                    If Not Double.TryParse(dtRec.Rows(i)("Qty"), dblRMQty) Then
                                        Throw New System.Exception("Please input Qty data as Number")
                                    End If
                                Else
                                    Throw New System.Exception("Please input Qty data as Number")
                                End If
                                If dtRec.Rows(i)("RHC").ToString().Length > 0 Then
                                    If Not Double.TryParse(dtRec.Rows(i)("RHC"), dblRHC) Then
                                        Throw New System.Exception("Please input RHC data as Number")
                                    End If
                                Else
                                    Throw New System.Exception("Please input RHC data as Number")
                                End If

                                Dim GridRow As DataRow()        '//Grid Data
                                Dim ExcelRow As DataRow()       '//Excel Data

                                '//Get Data on above row on Excel ------------------------
                                Dim chkSameFinalCompoundCodeBefore As String = String.Empty
                                Dim chkSameCompoundCodeBefore As String = String.Empty
                                Dim chkSameRevisionBefore As String = String.Empty
                                If i > 0 Then
                                    chkSameFinalCompoundCodeBefore = dtRec.Rows(i - 1)("FinalCompound_Code").ToString
                                    chkSameCompoundCodeBefore = dtRec.Rows(i - 1)("Compound_Code").ToString
                                    chkSameRevisionBefore = dtRec.Rows(i - 1)("Revision_No").ToString
                                Else
                                    chkSameFinalCompoundCodeBefore = ""
                                    chkSameCompoundCodeBefore = ""
                                    chkSameRevisionBefore = ""
                                End If
                                '//-------------------------------------------------------

                                '//Sum QTY each Compound Code and Revision
                                If strFinalCompoundCode <> chkSameFinalCompoundCodeBefore Or strCompoundCode <> chkSameCompoundCodeBefore Or strRevision <> chkSameRevisionBefore Then
                                    totalQty = 0
                                    totalRHC = 0
                                    GridRow = DTALL.Select("FinalCompound_Code = '" & strFinalCompoundCode & "' AND MasterCode = '" & strCompoundCode & "' AND Revision_No = '" & strRevision & "'")
                                    For k As Integer = 0 To GridRow.Count - 1
                                        totalQty = totalQty + GridRow(k)("Qty")
                                        totalRHC = totalRHC + GridRow(k)("mRHC")
                                    Next k

                                    ExcelRow = dtRec.Select("FinalCompound_Code = '" & strFinalCompoundCode & "' AND Compound_Code = '" & strCompoundCode & "' AND Revision_No = '" & strRevision & "'")
                                    For j As Integer = 0 To ExcelRow.Count - 1
                                        GridRow = DTALL.Select("FinalCompound_Code = '" & strFinalCompoundCode & "' AND MasterCode = '" & strCompoundCode & "' AND Revision_No = '" & strRevision & "'AND RMCode = '" & ExcelRow(j)("RMCode") & "'")

                                        If GridRow.Count > 0 Then
                                            totalQty = totalQty - GridRow(0)("Qty")
                                            totalRHC = totalRHC - GridRow(0)("mRHC")
                                        End If

                                        totalQty = totalQty + CDbl(ExcelRow(j)("Qty"))
                                        totalRHC = totalRHC + CDbl(ExcelRow(j)("RHC"))
                                    Next j
                                End If

                                '//Case 1 : Start With Check FinalCompoundCode, CompoundCode and Revision on Grid (TBLRHCDtl, TBLCompound)
                                '//Case 1.1 : [Found] Final,Comp,Rev So Next Check RMCode
                                '//Case 1.1.1 : [Found] RMCode So Next Check RHC and QTY
                                '//Case 1.1.1.1 : Do not change any RHC and QTY So Finish
                                '//Case 1.1.1.2 : Found some value changed RHC and QTY So next Update TBLRHCDtl, TBLCompound
                                '//Case 1.1.2 : [Found No] RMCode 
                                '//Case 1.2 : [Found No] Final,Comp,Rev on TBLRHCDtl So Next Add data to TBLRHCDtl

                                '//Case 1 : Start With Check FinalCompoundCode, CompoundCode and Revision on Grid (TBLRHCDtl, TBLCompound)
                                GridRow = DTALL.Select("FinalCompound_Code = '" & strFinalCompoundCode & "' AND MasterCode = '" & strCompoundCode & "' AND Revision_No = '" & strRevision & "'")

                                If GridRow.Count > 0 Then '//Case 1.1 : [Found] Final,Comp,Rev So Next Check RMCode

                                    GridRow = DTALL.Select("FinalCompound_Code = '" & strFinalCompoundCode & "' AND MasterCode = '" & strCompoundCode & "' AND Revision_No = '" & strRevision & "'AND RMCode = '" & strRMCode & "'")

                                    If GridRow.Count > 0 Then '//Case 1.1.1 : [Found] RMCode So Next Check RHC and QTY

                                        '//Case 1.1.1.2 : Found some value changed RHC and QTY So next Update TBLRHCDtl, TBLCompound
                                        If CDbl(GridRow(0)("Qty")) = dblRMQty Then
                                            If CDbl(GridRow(0)("mRHC")) <> dblRHC Then

                                                '//Update TBLRHCDtl (RHC each RMCode)
                                                sb.Clear()
                                                sb.AppendLine(" Update TBLRHCDtl")
                                                sb.AppendLine(" Set ")
                                                sb.AppendLine(" Weight = '" & dblRMQty & "'")
                                                sb.AppendLine(" , RHC = '" & dblRHC & "'")
                                                sb.AppendLine(" , Dateup = '" & strDate & "'")
                                                sb.AppendLine(" Where Final = '" & strFinalCompoundCode & "' AND MasterCode = '" & strCompoundCode & "' AND Revision = '" & strRevision & "' AND RMCode = '" & strRMCode & "'")

                                                sb.AppendLine(" ")

                                                '//Update TBLCompound ((RHC each RMCode))
                                                sb.AppendLine(" Update TBLCompound")
                                                sb.AppendLine(" Set ")
                                                sb.AppendLine(" Qty = '" & totalQty & "'")
                                                sb.AppendLine(" , RHC = '" & totalRHC & "'")
                                                sb.AppendLine(" , Dateup = '" & strDate & "'")
                                                sb.AppendLine(" Where FinalCompound = '" & strFinalCompoundCode & "' AND CompCode = '" & strCompoundCode & "' AND Revision = '" & strRevision & "' ")
                                                StrSQL = sb.ToString()
                                                cmSQL.CommandText = StrSQL
                                                cmSQL.ExecuteNonQuery()

                                            End If
                                        Else

                                            '//Update TBLRHCDtl (QTY total, RHC total)
                                            sb.Clear()
                                            sb.AppendLine(" Update TBLRHCDtl")
                                            sb.AppendLine(" Set ")
                                            sb.AppendLine(" Weight = '" & dblRMQty & "'")
                                            sb.AppendLine(" , RHC = '" & dblRHC & "'")
                                            sb.AppendLine(" , Dateup = '" & strDate & "'")
                                            sb.AppendLine(" Where Final = '" & strFinalCompoundCode & "' AND MasterCode = '" & strCompoundCode & "' AND Revision = '" & strRevision & "' AND RMCode = '" & strRMCode & "'")

                                            sb.AppendLine(" ")

                                            '//Update TBLCompound (QTY total, RHC total)
                                            sb.AppendLine(" Update TBLCompound")
                                            sb.AppendLine(" Set ")
                                            sb.AppendLine(" Qty = '" & totalQty & "'")
                                            sb.AppendLine(" , RHC = '" & totalRHC & "'")
                                            sb.AppendLine(" , Dateup = '" & strDate & "'")
                                            sb.AppendLine(" Where FinalCompound = '" & strFinalCompoundCode & "' AND CompCode = '" & strCompoundCode & "' AND Revision = '" & strRevision & "' ")

                                            sb.AppendLine(" ")

                                            '//Update TBLMASTER (QTY each RMCode)
                                            sb.AppendLine(" Update TBLMASTER")
                                            sb.AppendLine(" Set ")
                                            sb.AppendLine(" Qty = '" & dblRMQty & "'")
                                            sb.AppendLine(" Where MasterCode = '" & strCompoundCode & "' AND Revision = '" & strRevision & "' AND RMCode = '" & strRMCode & "' ")

                                            sb.AppendLine(" ")

                                            '//Update TBLConvert [SQty(totalQty)]
                                            sb.AppendLine(" Update TblConvert")
                                            sb.AppendLine(" Set ")
                                            sb.AppendLine(" SQty = '" & totalQty & "'")
                                            sb.AppendLine(" Where Final = '" & strFinalCompoundCode & "' AND Code = '" & strCompoundCode & "' AND Rev = '" & strRevision & "' AND Type = '03'")
                                            StrSQL = sb.ToString()
                                            cmSQL.CommandText = StrSQL
                                            cmSQL.ExecuteNonQuery()

                                            '//Update All Per in TBLMASTER***********
                                            sb.Clear()
                                            sb.AppendLine(" Update TBLMASTER")
                                            sb.AppendLine(" Set ")
                                            sb.AppendLine(" Per = Qty * 100 / " & totalQty)
                                            sb.AppendLine(" Where MasterCode = '" & strCompoundCode & "' AND Revision = '" & strRevision & "' ")
                                            StrSQL = sb.ToString()
                                            cmSQL.CommandText = StrSQL
                                            cmSQL.ExecuteNonQuery()

                                        End If
                                    Else '//Case 1.1.2 : [Found No] RMCode 

                                        '//Add TBLRHCDtl (RHC each RMCode)
                                        sb.Clear()
                                        sb.AppendLine(" Insert TBLRHCDtl ")
                                        sb.AppendLine(" Values (")
                                        sb.AppendLine(intSeq & ", ")                       'Column Seq
                                        sb.AppendLine(" '" & strFinalCompoundCode & "', ") 'Column Final
                                        sb.AppendLine(" '" & strCompoundCode & "', ")      'Column MasterCode
                                        sb.AppendLine(" '" & strRevision & "', ")          'Column Revision
                                        sb.AppendLine(" '" & strRMCode & "', ")            'Column RMCode
                                        sb.AppendLine(" '" & dblRMQty & "', ")             'Column Weight
                                        sb.AppendLine(" '" & dblRHC & "', ")               'Column RHC
                                        sb.AppendLine(" NULL, ")                           'Column PER
                                        sb.AppendLine(" '" & strDate & "' ")               'Column Dateup
                                        sb.AppendLine(" )")
                                        StrSQL = sb.ToString()
                                        cmSQL.CommandText = StrSQL
                                        cmSQL.ExecuteNonQuery()

                                        '//Update TBLCompound (QTY total, RHC total)
                                        sb.Clear()
                                        sb.AppendLine(" Update TBLCompound")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" Qty = '" & totalQty & "'")
                                        sb.AppendLine(" , RHC = '" & totalRHC & "'")
                                        sb.AppendLine(" , Dateup = '" & strDate & "'")
                                        sb.AppendLine(" Where FinalCompound = '" & strFinalCompoundCode & "' AND CompCode = '" & strCompoundCode & "' AND Revision = '" & strRevision & "' ")

                                        sb.AppendLine(" ")

                                        '//Update TBLMASTER (QTY each RMCode)
                                        sb.AppendLine(" Update TBLMASTER")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" Qty = '" & dblRMQty & "'")
                                        sb.AppendLine(" Where MasterCode = '" & strCompoundCode & "' AND Revision = '" & strRevision & "' AND RMCode = '" & strRMCode & "' ")

                                        sb.AppendLine(" ")

                                        '//Update TBLConvert [SQty(totalQty)]
                                        sb.AppendLine(" Update TblConvert")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" SQty = '" & totalQty & "'")
                                        sb.AppendLine(" Where Final = '" & strFinalCompoundCode & "' AND Code = '" & strCompoundCode & "' AND Rev = '" & strRevision & "' AND Type = '03'")
                                        StrSQL = sb.ToString()
                                        cmSQL.CommandText = StrSQL
                                        cmSQL.ExecuteNonQuery()

                                        '//Update All Per in TBLMASTER***********
                                        sb.Clear()
                                        sb.AppendLine(" Update TBLMASTER")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" Per = Qty * 100 / " & totalQty)
                                        sb.AppendLine(" Where MasterCode = '" & strCompoundCode & "' AND Revision = '" & strRevision & "' ")
                                        StrSQL = sb.ToString()
                                        cmSQL.CommandText = StrSQL
                                        cmSQL.ExecuteNonQuery()

                                    End If

                                Else '//Case 1.2 : [Found No] Final,Comp,Rev on TBLRHCDtl So Next Add data to TBLRHCDtl

                                    '//Add TBLRHCDtl (RHC each RMCode)
                                    sb.Clear()
                                    sb.AppendLine(" Insert TBLRHCDtl ")
                                    sb.AppendLine(" Values (")
                                    sb.AppendLine(intSeq & ", ")                           'Column Seq
                                    sb.AppendLine(" '" & strFinalCompoundCode & "', ")     'Column Final
                                    sb.AppendLine(" '" & strCompoundCode & "', ")          'Column MasterCode
                                    sb.AppendLine(" '" & strRevision & "', ")              'Column Revision
                                    sb.AppendLine(" '" & strRMCode & "', ")                'Column RMCode
                                    sb.AppendLine(" '" & dblRMQty & "', ")                 'Column Weight
                                    sb.AppendLine(" '" & dblRHC & "', ")                   'Column RHC
                                    sb.AppendLine(" NULL, ")                               'Column PER
                                    sb.AppendLine(" '" & strDate & "' ")                   'Column Dateup
                                    sb.AppendLine(" )")
                                    StrSQL = sb.ToString()
                                    cmSQL.CommandText = StrSQL
                                    cmSQL.ExecuteNonQuery()

                                    '//Update TBLCompound (QTY total, RHC total)
                                    sb.Clear()
                                    sb.AppendLine(" Update TBLCompound")
                                    sb.AppendLine(" Set ")
                                    sb.AppendLine(" Qty = '" & totalQty & "'")
                                    sb.AppendLine(" , RHC = '" & totalRHC & "'")
                                    sb.AppendLine(" , Dateup = '" & strDate & "'")
                                    sb.AppendLine(" Where FinalCompound = '" & strFinalCompoundCode & "' AND CompCode = '" & strCompoundCode & "' AND Revision = '" & strRevision & "' ")

                                    sb.AppendLine(" ")

                                    '//Update TBLMASTER (QTY each RMCode)
                                    sb.AppendLine(" Update TBLMASTER")
                                    sb.AppendLine(" Set ")
                                    sb.AppendLine(" Qty = '" & dblRMQty & "'")
                                    sb.AppendLine(" Where MasterCode = '" & strCompoundCode & "' AND Revision = '" & strRevision & "' AND RMCode = '" & strRMCode & "' ")

                                    sb.AppendLine(" ")

                                    '//Update TBLConvert [SQty(totalQty)]
                                    sb.AppendLine(" Update TblConvert")
                                    sb.AppendLine(" Set ")
                                    sb.AppendLine(" SQty = '" & totalQty & "'")
                                    sb.AppendLine(" Where Final = '" & strFinalCompoundCode & "' AND Code = '" & strCompoundCode & "' AND Rev = '" & strRevision & "' AND Type = '03'")
                                    StrSQL = sb.ToString()
                                    cmSQL.CommandText = StrSQL
                                    cmSQL.ExecuteNonQuery()

                                    '//Update All Per in TBLMASTER***********
                                    sb.Clear()
                                    sb.AppendLine(" Update TBLMASTER")
                                    sb.AppendLine(" Set ")
                                    sb.AppendLine(" Per = Qty * 100 / " & totalQty)
                                    sb.AppendLine(" Where MasterCode = '" & strCompoundCode & "' AND Revision = '" & strRevision & "' ")
                                    StrSQL = sb.ToString()
                                    cmSQL.CommandText = StrSQL
                                    cmSQL.ExecuteNonQuery()
                                End If
                            End If 'If strFinalCompoundCode.Length
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

                'Calculate Percent and Total Percent
                Using cnSQL As New SqlConnection(C1.Strcon)
                    cnSQL.Open()
                    Dim cmSQL As SqlCommand = cnSQL.CreateCommand()
                    Dim trans As SqlTransaction = cnSQL.BeginTransaction("RMTransaction")

                    cmSQL.Connection = cnSQL
                    cmSQL.Transaction = trans

                    Try
                        For i As Integer = 0 To dtRec.Rows.Count - 1
                            Dim strFinalCompoundCode As String = dtRec.Rows(i)("FinalCompound_Code").ToString().Trim()
                            Dim strCompoundCode As String = dtRec.Rows(i)("Compound_Code").ToString().Trim()
                            Dim strRevision As String = dtRec.Rows(i)("Revision_No").ToString().Trim()
                            Dim strRMCode As String = dtRec.Rows(i)("RMCode").ToString().Trim()
                            Dim dblRMQty As Double = dtRec.Rows(i)("Qty")
                            Dim dblRHC As Double = dtRec.Rows(i)("RHC")

                            Dim chkSameFinalCompoundCodeBefore As String = String.Empty
                            Dim chkSameCompoundCodeBefore As String = String.Empty
                            Dim chkSameRevisionBefore As String = String.Empty
                            If i > 0 Then
                                chkSameFinalCompoundCodeBefore = dtRec.Rows(i - 1)("FinalCompound_Code").ToString()
                                chkSameCompoundCodeBefore = dtRec.Rows(i - 1)("Compound_Code").ToString()
                                chkSameRevisionBefore = dtRec.Rows(i - 1)("Revision_No").ToString()
                            Else
                                chkSameFinalCompoundCodeBefore = dtRec.Rows(i)("FinalCompound_Code").ToString()
                                chkSameCompoundCodeBefore = dtRec.Rows(i)("Compound_Code").ToString()
                                chkSameRevisionBefore = dtRec.Rows(i)("Revision_No").ToString()
                            End If

                            'Calculate Percent
                            sb.Clear()
                            sb.AppendLine(" UPDATE TBLRHCDtl ")
                            sb.AppendLine(" SET Per = (     ")
                            sb.AppendLine("   SELECT Per FROM ( ")
                            sb.AppendLine("     SELECT Seq,Final,CRev,xx.compcode,mm.Rev,mm.RMCode,mQty,mm.mRHC,Per")
                            sb.AppendLine("     FROM (  ")
                            sb.AppendLine("       SELECT seq,final,compcode,Revision REV,cRev,RHC,Active,RMcode,mRHC,mQty ")
                            sb.AppendLine("       FROM (")
                            sb.AppendLine("         SELECT  seq, Finalcompound final, compcode, Revision, compcode +','+Revision CRev,RHC,Active ")
                            sb.AppendLine("         FROM  TBLCompound")
                            sb.AppendLine("       ) c")
                            sb.AppendLine("       LEFT OUTER JOIN ( ")
                            sb.AppendLine("         SELECT mastercode code,Revision Rev,mastercode+','+Revision MRev,RMcode,RHC mRHC,Weight mQty ")
                            sb.AppendLine("         FROM   TBLRHCDtl ")
                            sb.AppendLine("         WHERE Mastercode in (SELECT  compcode FROM  TBLCompound)")
                            sb.AppendLine("       ) m on c.CRev = m.MRev ")
                            sb.AppendLine("       WHERE Final = '" & strFinalCompoundCode & "'  AND  Compcode = '" & strCompoundCode & "' and  Rev = '" & strRevision & "' and RMCode = '" & strRMCode & "'")
                            sb.AppendLine("     ) mm ")
                            sb.AppendLine("     LEFT OUTER JOIN ( ")
                            sb.AppendLine("       SELECT aa.compcode,aa.Rev,Rmcode rcode,mrhc,mQty Qty,aa.Per ")
                            sb.AppendLine("       FROM (")
                            sb.AppendLine("         SELECT  cc.Seq, cc.FinalCompound Final, cc.CompCode, cc.Revision REV, aa.RMCode,mQty, cc.RHC, aa.mRHC, ROUND(aa.mRHC / cc.RHC * 100, 3) AS per, cc.Active ")
                            sb.AppendLine("         FROM (")
                            sb.AppendLine("           SELECT     Seq, FinalCompound, CompCode, Revision, Qty TQty, RHC, Active, CompCode + ',' + Revision Code")
                            sb.AppendLine("           FROM       TBLCompound")
                            sb.AppendLine("         ) cc ")
                            sb.AppendLine("         LEFT OUTER JOIN (")
                            sb.AppendLine("           SELECT     CRev, Code, Rev, RMCode, mRHC, mQty")
                            sb.AppendLine("           FROM  (")
                            sb.AppendLine("             SELECT     seq, Finalcompound, compcode + ',' + Revision CRev, RHC, Active")
                            sb.AppendLine("             FROM       TBLCompound")
                            sb.AppendLine("           ) c ")
                            sb.AppendLine("           LEFT OUTER JOIN ( ")
                            sb.AppendLine("             SELECT     mastercode code, Revision Rev, mastercode + ',' + Revision MRev, RMcode, RHC mRHC, Weight mQty")
                            sb.AppendLine("             FROM          TBLRHCDtl ")
                            sb.AppendLine("             WHERE      Mastercode In (SELECT     compcode FROM          TBLCompound) ")
                            sb.AppendLine("           ) m ON c.CRev = m.MRev")
                            sb.AppendLine("         ) aa ON cc.Code = aa.CRev")
                            sb.AppendLine("       ) aa ")
                            sb.AppendLine("     ) xx on mm.CRev+mm.RMCode  = xx.compcode+','+xx.Rev+xx.rcode and  mm.mrhc = xx.mrhc and mQty = Qty ")
                            sb.AppendLine("   ) zz ")
                            sb.AppendLine("   WHERE Final = '" & strFinalCompoundCode & "'  and  Compcode = '" & strCompoundCode & "' and  Rev = '" & strRevision & "' and RMcode = '" & strRMCode & "' and mRHC = '" & dblRHC & "' and mQty = '" & dblRMQty & "'")
                            sb.AppendLine(" )")
                            sb.AppendLine(" WHERE Final = '" & strFinalCompoundCode & "'  and  Mastercode = '" & strCompoundCode & "' and  Revision = '" & strRevision & "' and RMcode = '" & strRMCode & "' and RHC = '" & dblRHC & "'  and Weight = '" & dblRMQty & "'")
                            StrSQL = sb.ToString()
                            cmSQL.CommandText = StrSQL
                            cmSQL.ExecuteNonQuery()

                            If strFinalCompoundCode <> chkSameFinalCompoundCodeBefore Or strCompoundCode <> chkSameCompoundCodeBefore Or strRevision <> chkSameRevisionBefore Then
                                'Calculate total percent when it is last row of group
                                sb.Clear()
                                sb.AppendLine("UPDATE TBLcompound ")
                                sb.AppendLine("SET Per = (")
                                sb.AppendLine("  SELECT SUM(Per) Per ")
                                sb.AppendLine("  FROM TBLRHCDtl ")
                                sb.AppendLine("  WHERE Final = '" & chkSameFinalCompoundCodeBefore & "' and  Mastercode = '" & chkSameCompoundCodeBefore & "' and  Revision = '" & chkSameRevisionBefore & "' ")
                                sb.AppendLine("  GROUP BY Final,Mastercode,Revision")
                                sb.AppendLine(") ")
                                sb.AppendLine("WHERE FinalCompound = '" & chkSameFinalCompoundCodeBefore & "'  and  Compcode = '" & chkSameCompoundCodeBefore & "' and  Revision = '" & chkSameRevisionBefore & "'")
                                StrSQL = sb.ToString()
                                cmSQL.CommandText = StrSQL
                                cmSQL.ExecuteNonQuery()
                            End If
                        Next i

                        'Calculate total percent of last group
                        sb.Clear()
                        sb.AppendLine("UPDATE TBLcompound ")
                        sb.AppendLine("SET Per = (")
                        sb.AppendLine("  SELECT SUM(Per) Per ")
                        sb.AppendLine("  FROM TBLRHCDtl ")
                        sb.AppendLine("  WHERE Final = '" & dtRec.Rows(dtRec.Rows.Count - 1)("FinalCompound_Code").ToString() & "' and  Mastercode = '" & dtRec.Rows(dtRec.Rows.Count - 1)("Compound_Code").ToString() & "' and  Revision = '" & dtRec.Rows(dtRec.Rows.Count - 1)("Revision_No").ToString() & "' ")
                        sb.AppendLine("  GROUP BY Final,Mastercode,Revision")
                        sb.AppendLine(") ")
                        sb.AppendLine("WHERE FinalCompound = '" & dtRec.Rows(dtRec.Rows.Count - 1)("FinalCompound_Code").ToString() & "'  and  Compcode = '" & dtRec.Rows(dtRec.Rows.Count - 1)("Compound_Code").ToString() & "' and  Revision = '" & dtRec.Rows(dtRec.Rows.Count - 1)("Revision_No").ToString() & "'")
                        StrSQL = sb.ToString()
                        cmSQL.CommandText = StrSQL
                        cmSQL.ExecuteNonQuery()

                        trans.Commit()
                        MessageBox.Show("Calculate complete", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Catch ex As SqlException
                        MsgBox("Calculate error" & vbCrLf & ex.Message, MsgBoxStyle.Critical, "SQL Error")
                        trans.Rollback()
                    Catch ex As Exception
                        MsgBox("Calculate error" & vbCrLf & ex.Message, MsgBoxStyle.Critical, "General Error")
                        trans.Rollback()
                    Finally
                        trans.Dispose()
                        cmSQL.Dispose()
                        cnSQL.Close()
                        cnSQL.Dispose()
                    End Try
                End Using 'Using cnSQL
            End If 'If dtRec IsNot Nothing Then

            LoadCOM() 'ReQuery and set datagrid
            frmOverlay.Dispose()
        End If 'If importDialog.ShowDialog() = Windows.Forms.DialogResult.OK
    End Sub
#End Region

#Region "Import"
    Private Function ChkRMCodeEachCompound(ByVal ImportTable As DataTable) As Boolean
        Dim cnSQLRM As SqlConnection
        Dim cmSQLRM As SqlCommand
        Dim strSQL As String = String.Empty
        Dim ret As Boolean = False
        Dim strRmcodeBefore As String = String.Empty
        Dim distinctImportTabale As New DataTable()
        Dim sb As New System.Text.StringBuilder()

        Try
            For x As Integer = 0 To ImportTable.Rows.Count - 1
                Dim rmCode As String = ImportTable.Rows(x)("RMCode").ToString().Trim()
                Dim Final As String = ImportTable.Rows(x)("FinalCompound_Code").ToString().Trim()
                Dim Comp As String = ImportTable.Rows(x)("Compound_Code").ToString().Trim()
                Dim Rev As String = ImportTable.Rows(x)("Revision_No").ToString().Trim()
                sb.Clear()

                'Check empty
                'If Final.Equals(String.Empty) Or Comp.Equals(String.Empty) Or Rev.Equals(String.Empty) Then
                '    ret = False
                '    Throw New ApplicationException("FinalCompoundCode Code, Compound Code, Revision is not empty.")
                'End If

                If rmCode.Length > 0 Then
                    sb.AppendLine(" SELECT COUNT(*) ")
                    sb.AppendLine(" FROM (")
                    sb.AppendLine("   SELECT c.seq,c.FinalCompound,Compcode,c.Revision,RMCode")
                    sb.AppendLine("   FROM TBLCompound C ")
                    sb.AppendLine("   LEFT OUTER JOIN (")
                    sb.AppendLine("     SELECT * ")
                    sb.AppendLine("     FROM TBLMaster")
                    sb.AppendLine("   ) M on C.compcode+C.Revision = M.MasterCode+M.Revision ")
                    sb.AppendLine(" ) zzz")
                    sb.AppendLine(" WHERE RMCode IN (SELECT rmcode FROM TBLRM) ")
                    sb.AppendLine(" AND FinalCompound = '" & Final & "' AND CompCode = '" & Comp & "' AND Revision = '" & Rev & "' AND RMCode = '" & rmCode & "' ")
                    strSQL = sb.ToString()

                    cnSQLRM = New SqlConnection(C1.Strcon)
                    cnSQLRM.Open()
                    cmSQLRM = New SqlCommand(strSQL, cnSQLRM)
                    Dim i As Long = cmSQLRM.ExecuteScalar()
                    If i = 0 Then
                        cmSQLRM.Dispose()
                        cnSQLRM.Dispose()
                        Throw New System.Exception("This Group """ & Final & """, Compound """ & Comp & """, Revision """ & Rev & """ and RMCode """ & rmCode & """ is not found in master.")
                    Else
                        cmSQLRM.Dispose()
                        cnSQLRM.Dispose()
                    End If

                    cnSQLRM.Close()
                Else
                    'Check empty
                    If rmCode.Equals(String.Empty) Then
                        Throw New ApplicationException("RM Code is not empty.")
                    End If
                End If
            Next x

            ret = True
        Catch Exp As SqlException
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "SQL Error")
        Catch Exp As Exception
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "General Error")
        End Try

        Return ret
    End Function

    Private Function CalTotalPercent(ByVal final As String, ByVal Mastercode As String, ByVal Rev As String) As Boolean
        'Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Dim ret As Boolean = False

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
            ret = True
        Catch ex As Exception
            MsgBox("It may fail for calculation sum total. Please verify percent of sum total again.", 48)
            ret = False
        End Try
        cnn.Close()

        Return ret
        'Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Function
#End Region

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

    Sub Chk()
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

        SetTotal() 'Set number of items
    End Sub

    Private Sub SetTotal()
        'Set total
        'Format: Form Text - xxx item(s)
        Dim frmTitle As String() = Me.Text.Split(New Char() {"-"c})
        Me.Text = frmTitle(0) & "- " & GrdDV.Count & " item(s)"
    End Sub
#End Region
End Class
