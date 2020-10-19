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

Public Class FrmPerRHC

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
    Friend WithEvents CmdClose As System.Windows.Forms.Button
    Friend WithEvents CmdEdit As System.Windows.Forms.Button
    Friend WithEvents DataGridCOM As System.Windows.Forms.DataGrid
    Friend WithEvents CheckBoxGP As System.Windows.Forms.CheckBox
    Friend WithEvents CmbGroup As System.Windows.Forms.ComboBox
    Friend WithEvents CmdDelete As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ComboBoxStage As System.Windows.Forms.ComboBox
    Friend WithEvents RbP100 As System.Windows.Forms.RadioButton
    Friend WithEvents RbNP100 As System.Windows.Forms.RadioButton
    Friend WithEvents CheckBoxFinal As System.Windows.Forms.CheckBox
    Friend WithEvents RbAll As System.Windows.Forms.RadioButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmPerRHC))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.DataGridCOM = New System.Windows.Forms.DataGrid()
        Me.CmdClose = New System.Windows.Forms.Button()
        Me.CmdEdit = New System.Windows.Forms.Button()
        Me.CheckBoxGP = New System.Windows.Forms.CheckBox()
        Me.CmbGroup = New System.Windows.Forms.ComboBox()
        Me.CmdDelete = New System.Windows.Forms.Button()
        Me.ComboBoxStage = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.RbP100 = New System.Windows.Forms.RadioButton()
        Me.RbNP100 = New System.Windows.Forms.RadioButton()
        Me.CheckBoxFinal = New System.Windows.Forms.CheckBox()
        Me.RbAll = New System.Windows.Forms.RadioButton()
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
        Me.GroupBox1.Size = New System.Drawing.Size(874, 504)
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
        Me.DataGridCOM.Size = New System.Drawing.Size(868, 485)
        Me.DataGridCOM.TabIndex = 0
        '
        'CmdClose
        '
        Me.CmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdClose.Image = CType(resources.GetObject("CmdClose.Image"), System.Drawing.Image)
        Me.CmdClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdClose.Location = New System.Drawing.Point(802, 570)
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
        Me.CmdEdit.Location = New System.Drawing.Point(730, 570)
        Me.CmdEdit.Name = "CmdEdit"
        Me.CmdEdit.Size = New System.Drawing.Size(75, 56)
        Me.CmdEdit.TabIndex = 3
        Me.CmdEdit.Text = "Edit"
        Me.CmdEdit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
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
        '
        'ComboBoxStage
        '
        Me.ComboBoxStage.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9"})
        Me.ComboBoxStage.Location = New System.Drawing.Point(96, 8)
        Me.ComboBoxStage.Name = "ComboBoxStage"
        Me.ComboBoxStage.Size = New System.Drawing.Size(120, 21)
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
        'RbP100
        '
        Me.RbP100.Location = New System.Drawing.Point(72, 40)
        Me.RbP100.Name = "RbP100"
        Me.RbP100.Size = New System.Drawing.Size(64, 24)
        Me.RbP100.TabIndex = 15
        Me.RbP100.Text = "100 %"
        '
        'RbNP100
        '
        Me.RbNP100.Location = New System.Drawing.Point(136, 40)
        Me.RbNP100.Name = "RbNP100"
        Me.RbNP100.Size = New System.Drawing.Size(80, 24)
        Me.RbNP100.TabIndex = 16
        Me.RbNP100.Text = "<>  100 %"
        '
        'CheckBoxFinal
        '
        Me.CheckBoxFinal.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CheckBoxFinal.Location = New System.Drawing.Point(256, 44)
        Me.CheckBoxFinal.Name = "CheckBoxFinal"
        Me.CheckBoxFinal.Size = New System.Drawing.Size(112, 16)
        Me.CheckBoxFinal.TabIndex = 17
        Me.CheckBoxFinal.Text = "Final Compound"
        '
        'RbAll
        '
        Me.RbAll.Checked = True
        Me.RbAll.Location = New System.Drawing.Point(16, 40)
        Me.RbAll.Name = "RbAll"
        Me.RbAll.Size = New System.Drawing.Size(56, 24)
        Me.RbAll.TabIndex = 18
        Me.RbAll.TabStop = True
        Me.RbAll.Text = "All"
        '
        'CmdImport
        '
        Me.CmdImport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdImport.Image = CType(resources.GetObject("CmdImport.Image"), System.Drawing.Image)
        Me.CmdImport.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdImport.Location = New System.Drawing.Point(553, 570)
        Me.CmdImport.Name = "CmdImport"
        Me.CmdImport.Size = New System.Drawing.Size(75, 56)
        Me.CmdImport.TabIndex = 19
        Me.CmdImport.Text = "Import"
        Me.CmdImport.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdExport
        '
        Me.CmdExport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdExport.Image = CType(resources.GetObject("CmdExport.Image"), System.Drawing.Image)
        Me.CmdExport.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdExport.Location = New System.Drawing.Point(627, 570)
        Me.CmdExport.Name = "CmdExport"
        Me.CmdExport.Size = New System.Drawing.Size(75, 56)
        Me.CmdExport.TabIndex = 20
        Me.CmdExport.Text = "Export"
        Me.CmdExport.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'FrmPerRHC
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(890, 632)
        Me.Controls.Add(Me.CmdExport)
        Me.Controls.Add(Me.CmdImport)
        Me.Controls.Add(Me.RbAll)
        Me.Controls.Add(Me.CheckBoxFinal)
        Me.Controls.Add(Me.RbNP100)
        Me.Controls.Add(Me.RbP100)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ComboBoxStage)
        Me.Controls.Add(Me.CmdDelete)
        Me.Controls.Add(Me.CheckBoxGP)
        Me.Controls.Add(Me.CmbGroup)
        Me.Controls.Add(Me.CmdEdit)
        Me.Controls.Add(Me.CmdClose)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(906, 671)
        Me.Name = "FrmPerRHC"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Compound  ( % ) -"
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
        Dim sb As New System.Text.StringBuilder()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sb.AppendLine(" SELECT    seq,finalcompound,Compcode,Revision,Qty,RHC,Per,per Tper,Active")
        sb.AppendLine(" ,finalcompound Final,Compcode+Revision CRev,null rmcode,null mQty,null mRHC,null mPer, CompCode as mCompCode, Revision as mRev")
        sb.AppendLine(" FROM     TBLCompound  ")
        sb.AppendLine(" UNION")
        sb.AppendLine(" SELECT  seq,null finalcompound,null Compcode,null Revision,null Qty,null RHC,null Per,Tper,Active")
        sb.AppendLine(" ,final,mastercode+Revision MRev,rmcode,weight mQty,RHC mRHC,per mPer, mCompCode, mRev ")
        sb.AppendLine(" FROM (")
        sb.AppendLine("   SELECT dt.seq,dt.final,dt.mastercode,dt.revision,dt.rmcode,dt.weight,dt.RHC,dt.per,")
        sb.AppendLine("   c.Per Tper,c.Active, dt.MasterCode as mCompCode, dt.Revision as mRev")
        sb.AppendLine("   FROM         TBLRHCDtl dt")
        sb.AppendLine("   LEFT OUTER JOIN TBLcompound c on dt.final+dt.Mastercode+dt.Revision = c.Finalcompound+c.compcode+c.revision ")
        sb.AppendLine(" ) xxx")
        sb.AppendLine(" ORDER BY Crev ,finalcompound DESC ,RMcode")
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
            .MappingName = "Compcode"
            .NullText = ""
            .Width = 110
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle1_1 As New DataGridColoredLine2
        With grdColStyle1_1
            .HeaderText = "Rev."
            .MappingName = "Revision"
            .NullText = ""
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
            .MappingName = "QTY"
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
            .MappingName = "mQty"
            .Width = 75
            .Format = "##,###,##0.000"
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle5 As New DataGridColoredLine2
        With grdColStyle5
            .HeaderText = " % (Percent) "
            .MappingName = "mPER"
            .Width = 75
            .Format = "##,###,##0.000"
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle5_1 As New DataGridColoredLine2
        With grdColStyle5_1
            .HeaderText = "Total "
            .MappingName = "PER"
            .Width = 75
            .Format = "##,###,##0.000"
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        grdTableStyle1.GridColumnStyles.AddRange _
    (New DataGridColumnStyle() _
    {grdColStyle0, grdColStyle1, grdColStyle1_1, grdColStyle2,
     grdColStyle4_1, grdColStyle3_1, grdColStyle4, grdColStyle3, grdColStyle5, grdColStyle5_1})

        DataGridCOM.TableStyles.Add(grdTableStyle1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        '  ResetTableStyle()

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

#Region "Form Event"
    Private Sub FrmPerRHC_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Loadgroup()
        LoadCOM()
        SetTotal() 'Set number of items
    End Sub
#End Region

#Region "COMBOBOX"
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

#Region "Control Event"
    Private Sub CmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClose.Click
        Me.Close()
    End Sub

Private Sub CmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdEdit.Click

        Dim fAddpRHC As New FrmAddPerRHC
        fAddpRHC.CmdSave.Text = "Edit"
        fAddpRHC.TCompound = GrdDV.Item(oldrow).Row("Final")
        fAddpRHC.TCode = GrdDV.Item(oldrow).Row("mCompCode")
        fAddpRHC.TRev = GrdDV.Item(oldrow).Row("mRev")
        fAddpRHC.TStep = GrdDV.Item(oldrow).Row("seq")
        fAddpRHC.ShowDialog()
        LoadCOM()
        Selectcompcound()
    End Sub

    Private Sub DataGridCOM_CurrentCellChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGridCOM.CurrentCellChanged
        oldrow = DataGridCOM.CurrentCell.RowNumber
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
            End If
        Else
            Exit Sub
        End If

    End Sub

    Private Sub ComboBoxStage_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxStage.SelectedIndexChanged
        Selectcompcound()
    End Sub

    Private Sub CheckBoxGP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxGP.CheckedChanged
        If CheckBoxGP.Checked = True Then
            CmbGroup.Enabled = True
        Else
            CmbGroup.Enabled = False
        End If
        Selectcompcound()
    End Sub

    Private Sub CmbGroup_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbGroup.SelectedIndexChanged
        Selectcompcound()
    End Sub

    Private Sub CheckBoxFinal_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxFinal.CheckedChanged
        Selectcompcound()
    End Sub

    Private Sub RbAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RbAll.CheckedChanged
        Selectcompcound()
    End Sub

    Private Sub RbP100_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RbP100.CheckedChanged
        Selectcompcound()
    End Sub

    Private Sub RbNP100_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RbNP100.CheckedChanged
        Selectcompcound()
    End Sub

    Private Sub CmdExport_Click(sender As Object, e As EventArgs) Handles CmdExport.Click
        Dim arrColumn As String() = System.Configuration.ConfigurationManager.AppSettings("EXP_EXCEL_COLUMN_MASTER_COMPOUND_RHC_PER").ToString().Split(New Char() {","c})
        Dim arrColumnHeader As String() = System.Configuration.ConfigurationManager.AppSettings("EXP_EXCEL_COLUMN_HEADER_MASTER_COMPOUND_RHC_PER").ToString().Split(New Char() {","c})
        ExcelLib.Export(Me, GrdDV, TBL_RM, arrColumn, arrColumnHeader)
    End Sub

    Private Sub CmdImport_Click(sender As Object, e As EventArgs) Handles CmdImport.Click
        Dim arrColumn As String() = System.Configuration.ConfigurationManager.AppSettings("IMP_EXCEL_COLUMN_MASTER_COMPOUND_RHC_PER").ToString().Split(New Char() {","c})
        Dim importDialog As OpenFileDialog = New OpenFileDialog With {
            .Filter = System.Configuration.ConfigurationManager.AppSettings("DIALOG_FILE_EXT").ToString()
        }
        Dim dtRec As DataTable
        Dim sb As New System.Text.StringBuilder()
        Dim frmOverlay As New Form()
        Dim totalPer As Double = 0
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
                        If ChkAvailableFromTBLRHCDtl(dtRec) = False Then
                            LoadCOM() 'ReQuery and set datagrid
                            frmOverlay.Dispose()
                            Exit Sub
                        End If

                        For i As Integer = 0 To dtRec.Rows.Count - 1
                            Dim strFinalCompoundCode As String = dtRec.Rows(i)("FinalCompound_Code").ToString().Trim()
                            Dim strCompoundCode As String = dtRec.Rows(i)("Compound_Code").ToString().Trim()
                            Dim strRevision As String = dtRec.Rows(i)("Revision_No").ToString().Trim()
                            Dim strRMCode As String = dtRec.Rows(i)("RMCode").ToString().Trim()

                            If strFinalCompoundCode.Length <= 0 Then
                                Throw New System.Exception("Please input Final Compound data.")
                            End If

                            If strCompoundCode.Length <= 0 Then
                                Throw New System.Exception("Please input Compound data.")
                            End If

                            If strRevision.Length <= 0 Then
                                Throw New System.Exception("Please input Revision data.")
                            End If

                            If strFinalCompoundCode.Length > 0 Then
                                Dim dblPer As Double
                                Dim dblRHC As Double

                                If dtRec.Rows(i)("Percent").ToString.Length > 0 Then
                                    If Not Double.TryParse(dtRec.Rows(i)("Percent"), dblPer) Then
                                        Throw New System.Exception("Please input Percent data as Number")
                                    End If
                                Else
                                    Throw New System.Exception("Please input Percent data as Number")
                                End If
                                If dtRec.Rows(i)("RHC").ToString.Length > 0 Then
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
                                    totalPer = 0
                                    totalRHC = 0
                                    GridRow = DT.Select("Final = '" & strFinalCompoundCode & "' AND mCompCode = '" & strCompoundCode & "' AND mRev = '" & strRevision & "'")
                                    For k As Integer = 0 To GridRow.Count - 1
                                        If Not GridRow(k)("mPer") Is DBNull.Value Then
                                            totalPer = totalPer + GridRow(k)("mPer")
                                        End If

                                        If Not GridRow(k)("mRHC") Is DBNull.Value Then
                                            totalRHC = totalRHC + GridRow(k)("mRHC")
                                        End If

                                        'totalPer = totalPer + GridRow(k)("mPer")
                                        'totalRHC = totalRHC + GridRow(k)("mRHC")

                                    Next k

                                    ExcelRow = dtRec.Select("FinalCompound_Code = '" & strFinalCompoundCode & "' AND Compound_Code = '" & strCompoundCode & "' AND Revision_No = '" & strRevision & "'")
                                    For j As Integer = 0 To ExcelRow.Count - 1
                                        GridRow = DT.Select("Final = '" & strFinalCompoundCode & "' AND mCompCode = '" & strCompoundCode & "' AND mRev = '" & strRevision & "' AND rmcode = '" & ExcelRow(j)("RMCode") & "'")

                                        If GridRow.Count > 0 Then
                                            If Not GridRow(0)("mPer") Is DBNull.Value Then
                                                totalPer = totalPer - GridRow(0)("mPer")
                                            End If

                                            If Not GridRow(0)("mRHC") Is DBNull.Value Then
                                                totalRHC = totalRHC - GridRow(0)("mRHC")
                                            End If

                                            'totalPer = totalPer - GridRow(0)("mPer")
                                            'totalRHC = totalRHC - GridRow(0)("mRHC")
                                        End If

                                        totalPer = totalPer + CDbl(ExcelRow(j)("Percent"))
                                        totalRHC = totalRHC + CDbl(ExcelRow(j)("RHC"))
                                    Next j
                                End If

                                '//Case 1 : Start With Check FinalCompoundCode, CompoundCode, Revision and RMCode on Grid (TBLRHCDtl, TBLCompound)
                                GridRow = DT.Select("Final = '" & strFinalCompoundCode & "' AND mCompCode = '" & strCompoundCode & "' AND mRev = '" & strRevision & "' AND rmcode = '" & strRMCode & "'")

                                If GridRow.Count > 0 Then '//Case 1.1 : [Found] Check Null of RHC and Per form Excel // Then check RHC and Per same as Grid or not?
                                    '//Update TBLRHCDtl (RHC and Per each RMCode)
                                    sb.Clear()
                                    sb.AppendLine(" Update TBLRHCDtl")
                                    sb.AppendLine(" Set ")
                                    sb.AppendLine(" PER = '" & totalPer & "'")
                                    sb.AppendLine(" , RHC = '" & totalRHC & "'")
                                    sb.AppendLine(" , Dateup = '" & strDate & "'")
                                    sb.AppendLine(" Where Final = '" & strFinalCompoundCode & "' AND MasterCode = '" & strCompoundCode & "' AND Revision = '" & strRevision & "' AND RMCode = '" & strRMCode & "'")

                                    sb.AppendLine(" ")

                                    '//Update TBLCompound ((totalRHC and TotalPer each RMCode))
                                    sb.AppendLine(" Update TBLCompound")
                                    sb.AppendLine(" Set ")
                                    sb.AppendLine(" PER = '" & totalPer & "'")
                                    sb.AppendLine(" , RHC = '" & totalRHC & "'")
                                    sb.AppendLine(" , Dateup = '" & strDate & "'")
                                    sb.AppendLine(" Where FinalCompound = '" & strFinalCompoundCode & "' AND CompCode = '" & strCompoundCode & "' AND Revision = '" & strRevision & "' ")
                                    StrSQL = sb.ToString()
                                    cmSQL.CommandText = StrSQL
                                    cmSQL.ExecuteNonQuery()
                                End If
                            End If 'If strFinalCompoundCode.Length > 0
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
            End If 'If dtRec IsNot Nothing Then

            LoadCOM() 'ReQuery and set datagrid
            frmOverlay.Dispose()
        End If 'If importDialog.ShowDialog() = Windows.Forms.DialogResult.OK
    End Sub
#End Region

#Region "Import"
    Private Function ChkAvailableFromTBLRHCDtl(ByVal ImportTable As DataTable) As Boolean
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

                'If Final.Equals(String.Empty) Or Comp.Equals(String.Empty) Or Rev.Equals(String.Empty) Then
                '    ret = False
                '    Throw New ApplicationException("FinalCompoundCode Code, Compound Code and Revision is not empty.")
                'End If

                If rmCode.Length > 0 Then
                    sb.AppendLine(" SELECT COUNT(*) ")
                    sb.AppendLine(" FROM ( ")
                    sb.AppendLine("   SELECT R.Final,R.MasterCode,R.Revision,RMCode ")
                    sb.AppendLine("   FROM TBLRHCDtl as R ")
                    sb.AppendLine("   LEFT OUTER JOIN TBLCompound as C on R.Final+R.MasterCode+R.Revision = C.FinalCompound+C.CompCode+C.Revision ")
                    sb.AppendLine(" ) RC ")
                    sb.AppendLine(" WHERE Final = '" & Final & "' and MasterCode = '" & Comp & "' and Revision = '" & Rev & "' and RMCode = '" & rmCode & "' ")
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
        Dim strSQL As String
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Try
            strSQL = " Delete TblCompound"
            strSQL &= " where CompCode = '" & GrdDV.Item(oldrow).Row("CompCode") & "'"
            strSQL &= " and Revision = '" & GrdDV.Item(oldrow).Row("Revision") & "'"
            strSQL &= "  "
            strSQL &= " Delete TblMaster"
            strSQL &= " where Mastercode = '" & GrdDV.Item(oldrow).Row("CompCode") & "'"
            strSQL &= " and Revision = '" & GrdDV.Item(oldrow).Row("Revision") & "'"
            strSQL &= "  "
            strSQL &= " Delete TblGroup"
            strSQL &= " where Code = '" & GrdDV.Item(oldrow).Row("CompCode") & "'"

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

    Sub Selectcompcound()
        Dim StrSearch As String
        Dim k As Integer = 0
        If CheckBoxFinal.Checked = True Then
            StrSearch = " Active  =  '1'"
            'RbAll.Checked = True
            ComboBoxStage.Text = "Select"
        Else
            StrSearch = " "
        End If

        If RbAll.Checked = True Then
            If CheckBoxFinal.Checked = True Then
                StrSearch += "and" & " Tper <> 0"
            Else
                StrSearch += " Tper <> 0"
            End If
            RbP100.Checked = False
            RbNP100.Checked = False
        Else
            StrSearch += " "
        End If

        If RbP100.Checked = True Then
            If CheckBoxFinal.Checked = True Then
                StrSearch += "and" & "  Tper ='100.000'"
            Else
                StrSearch += "  Tper ='100.000'"
            End If
            RbAll.Checked = False
            RbNP100.Checked = False
        Else
            StrSearch += " "
        End If

        If RbNP100.Checked = True Then
            If CheckBoxFinal.Checked = True Then
                StrSearch += "and" & "  Tper <> '100.000'"
            Else
                StrSearch += "  Tper <> '100.000'"
            End If
            RbP100.Checked = False
            RbAll.Checked = False
        Else
            StrSearch += " "
        End If

        If CheckBoxGP.Checked = True Then
            If CheckBoxFinal.Checked = False Then
                If RbAll.Checked = False And RbP100.Checked = False _
              And RbNP100.Checked = False And CheckBoxGP.Checked = False Then
                    StrSearch += " Final like'%" & CmbGroup.Text.Trim & "%'"
                Else
                    StrSearch += "and" & " Final like'%" & CmbGroup.Text.Trim & "%'"
                End If
            Else
                StrSearch += "and" & " Final like'%" & CmbGroup.Text.Trim & "%'"
            End If
        Else
            StrSearch += " "
        End If

        If ComboBoxStage.Text <> "Select" Then
            If CheckBoxFinal.Checked = False Then
                If RbAll.Checked = False And RbP100.Checked = False _
                And RbNP100.Checked = False And CheckBoxGP.Checked = False Then
                    StrSearch += " seq = '" & ComboBoxStage.Text.Trim & "'"
                Else
                    StrSearch += "and" & " seq = '" & ComboBoxStage.Text.Trim & "'"
                End If
            Else
                StrSearch += "and" & " seq = '" & ComboBoxStage.Text.Trim & "'"
            End If
        Else
            StrSearch += " "
        End If

        GrdDV.RowFilter = StrSearch
        DataGridCOM.DataSource = GrdDV
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
