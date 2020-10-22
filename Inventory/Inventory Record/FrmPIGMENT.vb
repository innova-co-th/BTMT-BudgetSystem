#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Inventory_Record.Common
#End Region

Public Class FrmPIGMENT

#Region "Declare"
    Inherits System.Windows.Forms.Form
    Public Shared GrdDV As New DataView
    Protected Const TBL_PIGMENT As String = "TBL_PIGMENT"
    Public Shared GrdDVPigment As New DataView 'Combobox Pigment
    Dim C1 As New SQLData("ACCINV")

    Protected DefaultGridBorderStyle As BorderStyle
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
    Friend WithEvents DataGridPigment As System.Windows.Forms.DataGrid
    Friend WithEvents CmdSave As System.Windows.Forms.Button
    Friend WithEvents CmdClose As System.Windows.Forms.Button
    Friend WithEvents CmdEdit As System.Windows.Forms.Button
    Friend WithEvents CmbPigment As System.Windows.Forms.ComboBox
    Friend WithEvents CheckBoxPigment As System.Windows.Forms.CheckBox
    Friend WithEvents CmdDelete As System.Windows.Forms.Button
    Friend WithEvents CmdImport As System.Windows.Forms.Button
    Friend WithEvents CmdExport As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmPIGMENT))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.DataGridPigment = New System.Windows.Forms.DataGrid()
        Me.CmdSave = New System.Windows.Forms.Button()
        Me.CmdClose = New System.Windows.Forms.Button()
        Me.CmdEdit = New System.Windows.Forms.Button()
        Me.CmbPigment = New System.Windows.Forms.ComboBox()
        Me.CheckBoxPigment = New System.Windows.Forms.CheckBox()
        Me.CmdDelete = New System.Windows.Forms.Button()
        Me.CmdImport = New System.Windows.Forms.Button()
        Me.CmdExport = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridPigment, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.DataGridPigment)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(8, 48)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(762, 499)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'DataGridPigment
        '
        Me.DataGridPigment.DataMember = ""
        Me.DataGridPigment.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridPigment.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridPigment.Location = New System.Drawing.Point(3, 16)
        Me.DataGridPigment.Name = "DataGridPigment"
        Me.DataGridPigment.Size = New System.Drawing.Size(756, 480)
        Me.DataGridPigment.TabIndex = 0
        '
        'CmdSave
        '
        Me.CmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdSave.Image = CType(resources.GetObject("CmdSave.Image"), System.Drawing.Image)
        Me.CmdSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdSave.Location = New System.Drawing.Point(616, 549)
        Me.CmdSave.Name = "CmdSave"
        Me.CmdSave.Size = New System.Drawing.Size(72, 56)
        Me.CmdSave.TabIndex = 1
        Me.CmdSave.Text = "Add"
        Me.CmdSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdClose
        '
        Me.CmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdClose.Image = CType(resources.GetObject("CmdClose.Image"), System.Drawing.Image)
        Me.CmdClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdClose.Location = New System.Drawing.Point(688, 549)
        Me.CmdClose.Name = "CmdClose"
        Me.CmdClose.Size = New System.Drawing.Size(75, 56)
        Me.CmdClose.TabIndex = 2
        Me.CmdClose.Text = "Close"
        Me.CmdClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdEdit
        '
        Me.CmdEdit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdEdit.Enabled = False
        Me.CmdEdit.Image = CType(resources.GetObject("CmdEdit.Image"), System.Drawing.Image)
        Me.CmdEdit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdEdit.Location = New System.Drawing.Point(544, 549)
        Me.CmdEdit.Name = "CmdEdit"
        Me.CmdEdit.Size = New System.Drawing.Size(72, 56)
        Me.CmdEdit.TabIndex = 3
        Me.CmdEdit.Text = "Edit"
        Me.CmdEdit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.CmdEdit.Visible = False
        '
        'CmbPigment
        '
        Me.CmbPigment.Enabled = False
        Me.CmbPigment.Location = New System.Drawing.Point(104, 16)
        Me.CmbPigment.Name = "CmbPigment"
        Me.CmbPigment.Size = New System.Drawing.Size(112, 21)
        Me.CmbPigment.TabIndex = 5
        '
        'CheckBoxPigment
        '
        Me.CheckBoxPigment.Location = New System.Drawing.Point(16, 18)
        Me.CheckBoxPigment.Name = "CheckBoxPigment"
        Me.CheckBoxPigment.Size = New System.Drawing.Size(80, 16)
        Me.CheckBoxPigment.TabIndex = 6
        Me.CheckBoxPigment.Text = "PIGMENT"
        '
        'CmdDelete
        '
        Me.CmdDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CmdDelete.Image = CType(resources.GetObject("CmdDelete.Image"), System.Drawing.Image)
        Me.CmdDelete.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdDelete.Location = New System.Drawing.Point(8, 547)
        Me.CmdDelete.Name = "CmdDelete"
        Me.CmdDelete.Size = New System.Drawing.Size(80, 56)
        Me.CmdDelete.TabIndex = 7
        Me.CmdDelete.Text = "Del"
        Me.CmdDelete.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdImport
        '
        Me.CmdImport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdImport.Image = CType(resources.GetObject("CmdImport.Image"), System.Drawing.Image)
        Me.CmdImport.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdImport.Location = New System.Drawing.Point(443, 550)
        Me.CmdImport.Name = "CmdImport"
        Me.CmdImport.Size = New System.Drawing.Size(72, 56)
        Me.CmdImport.TabIndex = 8
        Me.CmdImport.Text = "Import"
        Me.CmdImport.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdExport
        '
        Me.CmdExport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdExport.Image = CType(resources.GetObject("CmdExport.Image"), System.Drawing.Image)
        Me.CmdExport.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdExport.Location = New System.Drawing.Point(515, 550)
        Me.CmdExport.Name = "CmdExport"
        Me.CmdExport.Size = New System.Drawing.Size(72, 56)
        Me.CmdExport.TabIndex = 9
        Me.CmdExport.Text = "Export"
        Me.CmdExport.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'FrmPIGMENT
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(778, 611)
        Me.Controls.Add(Me.CmdExport)
        Me.Controls.Add(Me.CmdImport)
        Me.Controls.Add(Me.CmdDelete)
        Me.Controls.Add(Me.CheckBoxPigment)
        Me.Controls.Add(Me.CmbPigment)
        Me.Controls.Add(Me.CmdEdit)
        Me.Controls.Add(Me.CmdClose)
        Me.Controls.Add(Me.CmdSave)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(794, 650)
        Me.Name = "FrmPIGMENT"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "PIGMENT (R/M Material) -"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataGridPigment, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "CONSTANT"
    Dim DT As New DataTable
    Dim StrSQL As String
    Dim oldrow As Integer
#End Region

#Region "COMBOBOX"
    Sub LoadCmbPigment()
        Dim dtComp As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        StrSQL = "SELECT  Code "
        StrSQL &= "  FROM  TblGroup  where Typecode = '02'"

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
        dtComp.TableName = TBL_PIGMENT
        GrdDVPigment = dtComp.DefaultView
        '************************************
        CmbPigment.DisplayMember = "Code"
        CmbPigment.ValueMember = "Code"
        CmbPigment.DataSource = dtComp
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
#End Region

#Region "Form Event"
    Private Sub FrmPIGMENT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadCmbPigment() 'Get Pigment Code
        LoadPIGMENT()
        SetTotal() 'Set number of items
    End Sub
#End Region

#Region "Function_Load"
    Private Sub LoadPIGMENT()
        Dim sb As New System.Text.StringBuilder()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sb.AppendLine(" SELECT aa.Code, aa.PigmentCode, aa.Revision, aa.Qty, aa.Unit, aa.Per")
        sb.AppendLine(" ,aa.PType, aa.PigmentCode + ',' + aa.Revision as PCode, aa.rmCode, aa.RmRevision, aa.RmQty, aa.RMUnit, aa.TypeCode, aa.EachRevision, aa.EachPigmentCode")
        sb.AppendLine(" FROM (")
        sb.AppendLine("   SELECT Code, p.PigmentCode, p.Revision, p.Qty, p.Unit, null Per")
        sb.AppendLine("   ,'H' PType, p.PigmentCode + ',' + p.Revision as PCode,'' rmCode,'' RmRevision, null RmQty,'' RMUnit, H.Typecode, '' as EachRevision, '' as EachPigmentCode") 'Header from table TBLPigment
        sb.AppendLine("   FROM TBLPigment p ")
        sb.AppendLine("   LEFT OUTER JOIN (")
        sb.AppendLine("     SELECT t.Typecode,t.TypeName,G.Code,'H' PType")
        sb.AppendLine("     FROM TBLType t")
        sb.AppendLine("     LEFT OUTER JOIN TBLGroup G on t.TypeCode = g.TypeCode ")
        sb.AppendLine("     WHERE t.TypeCode = '02'")
        sb.AppendLine("   ) H on p.PIGMENTCode = H.Code")
        sb.AppendLine(" ) aa") 'Header
        sb.AppendLine(" UNION")
        sb.AppendLine(" SELECT * ")
        sb.AppendLine(" FROM (")
        sb.AppendLine("   SELECT '' as code, '' as PigmentCode,'' Revision,null Qty,'' Unit, t.Per")
        sb.AppendLine("   ,'D' PType, t.MasterCode + ',' + Revision as PCode, t.RMCode, t.RmRevision, t.Qty RMQty, t.Unit RMUnit, g.TypeCode, t.Revision as EachRevision, t.MasterCode as EachPigmentCode") 'Detail from table TBLMaster
        sb.AppendLine("   FROM TBLMASTER t")
        sb.AppendLine("   LEFT OUTER JOIN TBLGroup G on t.MasterCode = g.Code")
        sb.AppendLine("   WHERE g.TypeCode = '02'")
        sb.AppendLine(" ) bb") 'Detail
        sb.AppendLine(" ORDER BY PCode, PType DESC")
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
        DT.TableName = TBL_PIGMENT
        GrdDV = DT.DefaultView
        GrdDV.AllowNew = False
        GrdDV.AllowDelete = False
        '************************************
        DataGridPigment.DataSource = GrdDV
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

        With DataGridPigment
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
            .MappingName = TBL_PIGMENT
            .PreferredColumnWidth = 125
            .PreferredRowHeight = 15
        End With

        Dim d As delegateGetRowSeq
        d = AddressOf MyGetSeqLine
        'Dim grdColStyle0 As New DataGridColoredLine
        'With grdColStyle0
        '    .HeaderText = "TypeName"
        '    .MappingName = "Typename"
        '    .Width = 110
        '    .ReadOnly = True
        '    .Alignment = HorizontalAlignment.Center
        'End With
        Dim grdColStyle1 As New DataGridColoredLine
        With grdColStyle1
            .HeaderText = "PIGMENT"
            .MappingName = "Pigmentcode"
            .Width = 80
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle2 As New DataGridColoredLine2
        With grdColStyle2
            .HeaderText = "R/M Material"
            .MappingName = "RMCode"
            .Width = 100
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle3 As New DataGridColoredLine2
        With grdColStyle3
            .HeaderText = "Qty "
            .MappingName = "RMQty"
            .Width = 65
            .NullText = ""
            .Format = "##,###,###.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle4 As New DataGridColoredLine2
        With grdColStyle4
            .HeaderText = "Unit "
            .MappingName = "RMUnit"
            .NullText = ""
            .Width = 75
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle5 As New DataGridColoredLine
        With grdColStyle5
            .HeaderText = "Total "
            .MappingName = "Qty"
            .Width = 65
            .NullText = ""
            .Format = "##,###,###.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle6 As New DataGridColoredLine
        With grdColStyle6
            .HeaderText = "Unit "
            .MappingName = "Unit"
            .NullText = ""
            .Width = 75
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle7 As New DataGridColoredLine
        With grdColStyle7
            .HeaderText = "Per %"
            .MappingName = "Per"
            .NullText = ""
            .Format = "#0.000"
            .Width = 75
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        grdTableStyle1.GridColumnStyles.AddRange _
        (New DataGridColumnStyle() _
        {grdColStyle1, grdColStyle2, grdColStyle3, grdColStyle4,
        grdColStyle5, grdColStyle6, grdColStyle7})

        DataGridPigment.TableStyles.Add(grdTableStyle1)
        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub

#Region "Delegate function"
    Public Shared Function MyGetSeqLine(ByVal row As Integer) As CellColor
        Dim c As CellColor
        c.ForeG = CInt(GrdDV.Item(row).Item(0))
        c.BackG = CInt(GrdDV.Item(row).Item(1))
        c.LfItem = Mid(GrdDV.Item(row).Item(3), 1, 4)
        Return c
    End Function
#End Region

    Private Sub ResetTableStyle()

        ' Clear out the existing TableStyles and result default formatting.
        With DataGridPigment
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

#Region "Control Event"
    Private Sub CmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClose.Click
        Me.Close()
    End Sub

    Private Sub CmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdEdit.Click
        Dim faddPigment As New FrmAddPigment
        faddPigment.CmdSave.Text = "Edit"
        faddPigment.TxtCode.Text = GrdDV.Item(oldrow).Row("Pigmentcode")
        faddPigment.TxtRev.Text = GrdDV.Item(oldrow).Row("Revision")
        If GrdDV.Item(oldrow).Row("Pigmentcode") = "" Then
            Exit Sub
        Else
            faddPigment.ShowDialog()
            LoadPIGMENT()
        End If
        Changedata()
    End Sub

    Private Sub CmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSave.Click
        Dim faddPigment As New FrmAddPigment
        faddPigment.CmdSave.Text = "Save"
        faddPigment.ShowDialog()
        LoadCmbPigment()
        LoadPIGMENT()
        Changedata()
    End Sub

    Private Sub DataGridPigment_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridPigment.CurrentCellChanged
        oldrow = DataGridPigment.CurrentCell.RowNumber
    End Sub

    Private Sub DataGridPigment_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridPigment.DoubleClick
        'Dim fEditPigment As New FrmEditPigment
        'Dim str() As String
        'str = Split(GrdDV.Item(oldrow).Row("PCode"), ",")
        'fEditPigment.TxtCode.Text = str(0)
        'fEditPigment.TxtCode.Enabled = False
        'fEditPigment.TxtRM.Text = GrdDV.Item(oldrow).Row("RMCode")
        'fEditPigment.TxtRM.Enabled = False
        'If GrdDV.Item(oldrow).Row("RMCode") = "" Then
        '    Exit Sub
        'Else
        '    fEditPigment.TxtQty.Text = GrdDV.Item(oldrow).Row("RMQty")
        '    fEditPigment.TxtRev.Enabled = False
        'End If
        'fEditPigment.TxtRev.Text = str(1)
        'fEditPigment.qNo = GrdDV.Item(oldrow).Row("RMQty")
        'fEditPigment.ShowDialog()
        'LoadPIGMENT()
        'If CheckBoxPigment.Checked = True Then
        '    GrdDV.RowFilter = " Pigmentcode = '" & CmbPigment.Text.Trim & "'"
        '    DataGridPigment.DataSource = GrdDV
        '    CmbPigment.Enabled = True
        'Else
        '    GrdDV.RowFilter = ""
        '    DataGridPigment.DataSource = GrdDV
        '    CmbPigment.Enabled = False
        'End If
    End Sub

    Private Sub CheckBoxPigment_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxPigment.CheckedChanged
        LoadCmbPigment()
        Changedata()
    End Sub

    Private Sub CmbPigment_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbPigment.SelectedIndexChanged
        Changedata()
    End Sub

    Private Sub CmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdDelete.Click
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        msg = "Delete pigmant :" & GrdDV.Item(oldrow).Row("PigmentCode") ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Pigment "   ' Define title.

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            If ChkData() Then
                MsgBox("It's have Usage , Can't Delete. Please contact IS.", MsgBoxStyle.Information, "Delete R/M ")
            Else
                DelPigment()
            End If
        Else
            Exit Sub
        End If
        LoadPIGMENT()
        Changedata()
    End Sub

    Private Sub CmdImport_Click(sender As Object, e As EventArgs) Handles CmdImport.Click
        Dim arrColumn As String() = System.Configuration.ConfigurationManager.AppSettings("IMP_EXCEL_COLUMN_MASTER_PIGMENT").ToString().Split(New Char() {","c})
        Dim importDialog As OpenFileDialog = New OpenFileDialog With {
            .Filter = System.Configuration.ConfigurationManager.AppSettings("DIALOG_FILE_EXT").ToString()
        }
        Dim dtRec As DataTable
        Dim sb As New System.Text.StringBuilder()
        Dim frmOverlay As New Form()
        Dim totalQty As Double = 0

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
            dtRec = ExcelLib.Import(importDialog.FileName, Me, GrdDV, TBL_PIGMENT, arrColumn)

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
                        'Dim totalQty As Double = 0

                        '//Sort Data from Excel
                        dtRec.DefaultView.Sort = "EachPigmentCode DESC, EachRevision DESC"
                        dtRec = dtRec.DefaultView.ToTable

                        '//Check RMCode on Master
                        If ChkRMCodeMaster(dtRec) = False Then
                            LoadPIGMENT() 'ReQuery and set datagrid
                            frmOverlay.Dispose()
                            Exit Sub
                        End If

                        For i As Integer = 0 To dtRec.Rows.Count - 1
                            Dim strEachPigmentCode As String = dtRec.Rows(i)("EachPigmentCode").ToString().Trim()

                            If strEachPigmentCode.Length <= 0 Then
                                Throw New System.Exception("Please input EachPigmentCode data.")
                            End If
                            If strEachPigmentCode.Length > 0 Then
                                Dim strEachRevision As String = dtRec.Rows(i)("EachRevision").ToString().Trim()
                                Dim strRMCode As String = dtRec.Rows(i)("rmCode").ToString().Trim()
                                Dim dblRMQty As Double

                                'Check empty
                                If strEachRevision.Length = 0 Then
                                    Throw New System.Exception("Please input EachRevision data.")
                                ElseIf strEachRevision.Length > 3 Then
                                    Throw New System.Exception("EachRevision data must less than 4 digits.")
                                End If

                                'Check RmQty
                                If dtRec.Rows(i)("RmQty").ToString().Length > 0 Then
                                    If Not Double.TryParse(dtRec.Rows(i)("RmQty"), dblRMQty) Then
                                        Throw New System.Exception("Please input Qty data as Number.")
                                    End If
                                Else
                                    Throw New System.Exception("Please input Qty data as Number.")
                                End If

                                Dim DTRow As DataRow()          '//Grid Data

                                '//Check is same {Pigmentcode} as above excel row
                                Dim chkSameEachPigmentCodeBefore As String = String.Empty
                                Dim chkSameEachRevisionBefore As String = String.Empty
                                If i > 0 Then
                                    chkSameEachPigmentCodeBefore = dtRec.Rows(i - 1)("EachPigmentCode").ToString()
                                    chkSameEachRevisionBefore = dtRec.Rows(i - 1)("EachRevision").ToString()
                                Else
                                    chkSameEachPigmentCodeBefore = String.Empty
                                    chkSameEachRevisionBefore = String.Empty
                                End If

                                'Filter data in data grid
                                DTRow = DT.Select("EachPigmentCode = '" & strEachPigmentCode & "' AND EachRevision = '" & strEachRevision & "' ")

                                '//Sum QTY each PigmentCode and Revision
                                If strEachRevision <> chkSameEachPigmentCodeBefore Or strEachRevision <> chkSameEachRevisionBefore Then
                                    'First record of each pigment and each reivision
                                    totalQty = 0

                                    If DTRow.Length > 0 Then
                                        'Found data in data grid
                                        'Summarize QTY in data grid
                                        For j As Integer = 0 To DTRow.Length - 1
                                            'Check exist in excel
                                            Dim drRecRMCode As DataRow() = dtRec.Select("EachPigmentCode = '" & strEachPigmentCode & "' AND EachRevision = '" & strEachRevision & "' AND rmCode = '" & DTRow(j)("RMCode") & "' ")
                                            If drRecRMCode.Length > 0 Then
                                                'Found RMCode in excel
                                                'Summarize QTY from excel
                                                totalQty = totalQty + drRecRMCode(0)("RmQty")
                                            Else
                                                'Summarize QTY from data grid
                                                totalQty = totalQty + Convert.ToDouble(DTRow(j)("RMQty"))
                                            End If
                                        Next j

                                        'Summarize QTY in excel which is not in data grid
                                        Dim drRecExcel As DataRow() = dtRec.Select("EachPigmentCode = '" & strEachPigmentCode & "' AND EachRevision = '" & strEachRevision & "' ")
                                        For j As Integer = 0 To drRecExcel.Length - 1
                                            'Check not exist in data grid
                                            Dim drRecRMCode As DataRow() = DT.Select("EachPigmentCode = '" & strEachPigmentCode & "' AND EachRevision = '" & strEachRevision & "' AND RMCode = '" & drRecExcel(j)("rmCode") & "' ")
                                            If drRecRMCode.Length = 0 Then
                                                totalQty = totalQty + drRecExcel(j)("RmQty")
                                            End If
                                        Next j
                                    Else
                                        'New data
                                        'Filter data in excel
                                        Dim drRecExcel As DataRow() = dtRec.Select("EachPigmentCode = '" & strEachPigmentCode & "' AND EachRevision = '" & strEachRevision & "' ")
                                        For j As Integer = 0 To drRecExcel.Length - 1
                                            totalQty = totalQty + drRecExcel(j)("RmQty")
                                        Next j
                                    End If 'If DTRow.Length > 0
                                End If 'If strEachRevision <> chkSameEachPigmentCodeBefore Or strEachRevision <> chkSameEachRevisionBefore

                                '//Check matching between Excel and Grid
                                If DTRow.Count > 0 Then 'Have data on Grid
                                    '//Check RMCode each PigmentCode and Revision.
                                    DTRow = DT.Select("EachPigmentCode = '" & strEachPigmentCode & "' AND EachRevision = '" & strEachRevision & "' AND rmCode = '" & strRMCode & "' ")
                                    If DTRow.Count > 0 AndAlso CDbl(DTRow(0)("RmQty")) <> dblRMQty Then
                                        '//Update TBLMaster [Qty(dblRMQty), Per(dblRMQty * 100) / totalQty)]
                                        sb.Clear()
                                        sb.AppendLine(" Update TBLMASTER")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" Qty = '" & dblRMQty & "'")
                                        sb.AppendLine(" , Per = '" & (dblRMQty * 100 / totalQty) & "'")
                                        sb.AppendLine(" Where MasterCode = " & PrepareStr(strEachPigmentCode) & " AND Revision = " & PrepareStr(strEachRevision) & " AND RMCode = " & PrepareStr(strRMCode))
                                        StrSQL = sb.ToString()
                                        cmSQL.CommandText = StrSQL
                                        cmSQL.ExecuteNonQuery()

                                        '//Update All Per in TBLMASTER***********
                                        sb.Clear()
                                        sb.AppendLine(" Update TBLMASTER")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" Per = Qty * 100 / " & totalQty)
                                        sb.AppendLine(" Where MasterCode = " & PrepareStr(strEachPigmentCode) & " AND Revision = " & PrepareStr(strEachRevision))
                                        StrSQL = sb.ToString()
                                        cmSQL.CommandText = StrSQL
                                        cmSQL.ExecuteNonQuery()

                                        '//Update TBLPigment [Qty(totalQty), DateUp(strDate)]
                                        sb.Clear()
                                        sb.AppendLine(" Update TBLPigment")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" Qty = '" & totalQty & "'")
                                        sb.AppendLine(" , Dateup = '" & strDate & "'")
                                        sb.AppendLine(" Where PIGMENTCode = " & PrepareStr(strEachPigmentCode) & " AND Revision = " & PrepareStr(strEachRevision))
                                        'StrSQL = sb.ToString()
                                        'cmSQL.CommandText = StrSQL
                                        'cmSQL.ExecuteNonQuery()

                                        sb.AppendLine(" ")

                                        '//Update TBLConvert [SQty(totalQty)]
                                        sb.AppendLine(" Update TblConvert")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" SQty = '" & totalQty & "'")
                                        sb.AppendLine(" Where Code = " & PrepareStr(strEachPigmentCode) & " AND Rev = " & PrepareStr(strEachRevision) & " AND Type = '02'")

                                        StrSQL = sb.ToString()
                                        cmSQL.CommandText = StrSQL
                                        cmSQL.ExecuteNonQuery()

                                    ElseIf DTRow.Count <= 0 Then
                                        '//Insert TBLMaster
                                        sb.Clear()
                                        sb.AppendLine(" Insert  TBLMASTER ")
                                        sb.AppendLine(" Values (")
                                        sb.AppendLine(PrepareStr(strEachPigmentCode) & ", ")        'Column MasterCode
                                        sb.AppendLine(PrepareStr(strEachRevision) & ", ")           'Column Revision
                                        sb.AppendLine(PrepareStr(strRMCode) & ", ")                 'Column RMCode
                                        sb.AppendLine(PrepareStr("") & ", ")                        'Column RmRevision
                                        sb.AppendLine(PrepareStr(dblRMQty) & ", ")                  'Column Qty
                                        sb.AppendLine("'KG', ")                                     'Column Unit
                                        sb.AppendLine(" '" & (dblRMQty * 100 / totalQty) & "'")   'Column Per
                                        sb.AppendLine(" )")
                                        StrSQL = sb.ToString()
                                        cmSQL.CommandText = StrSQL
                                        cmSQL.ExecuteNonQuery()

                                        '//Update All Per in TBLMASTER***********
                                        sb.Clear()
                                        sb.AppendLine(" Update TBLMASTER")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" Per = Qty * 100 / " & totalQty)
                                        sb.AppendLine(" Where MasterCode = " & PrepareStr(strEachPigmentCode) & " AND Revision = " & PrepareStr(strEachRevision))
                                        StrSQL = sb.ToString()
                                        cmSQL.CommandText = StrSQL
                                        cmSQL.ExecuteNonQuery()

                                        '//Update TBLPigment [Qty(totalQty), DateUp(strDate)]
                                        sb.Clear()
                                        sb.AppendLine(" Update TBLPigment")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" Qty = '" & totalQty & "'")
                                        sb.AppendLine(" , Dateup = '" & strDate & "'")
                                        sb.AppendLine(" Where PIGMENTCode = " & PrepareStr(strEachPigmentCode) & " AND Revision = " & PrepareStr(strEachRevision))
                                        StrSQL = sb.ToString()
                                        cmSQL.CommandText = StrSQL
                                        cmSQL.ExecuteNonQuery()

                                        sb.AppendLine(" ")

                                        '//Update TBLConvert [SQty(totalQty)]
                                        sb.AppendLine(" Update TblConvert")
                                        sb.AppendLine(" Set ")
                                        sb.AppendLine(" SQty = '" & totalQty & "'")
                                        sb.AppendLine(" Where Code = " & PrepareStr(strEachPigmentCode) & " AND Rev = " & PrepareStr(strEachRevision) & " AND Type = '02'")

                                        StrSQL = sb.ToString()
                                        cmSQL.CommandText = StrSQL
                                        cmSQL.ExecuteNonQuery()
                                    End If
                                Else 'Have no data on Grid
                                    If strEachPigmentCode = chkSameEachPigmentCodeBefore Then
                                        '//Check is same {Revision} as above excel row
                                        If strEachRevision = chkSameEachRevisionBefore Then
                                            '//Insert TBLMaster
                                            sb.Clear()
                                            sb.AppendLine(" Insert  TBLMASTER ")
                                            sb.AppendLine(" Values (")
                                            sb.AppendLine(PrepareStr(strEachPigmentCode) & ", ")        'Column MasterCode
                                            sb.AppendLine(PrepareStr(strEachRevision) & ", ")           'Column Revision
                                            sb.AppendLine(PrepareStr(strRMCode) & ", ")                 'Column RMCode
                                            sb.AppendLine(PrepareStr("") & ", ")                        'Column RmRevision
                                            sb.AppendLine(PrepareStr(dblRMQty) & ", ")                  'Column Qty
                                            sb.AppendLine("'KG', ")                                     'Column Unit
                                            sb.AppendLine(" '" & (dblRMQty * 100 / totalQty) & "'")   'Column Per
                                            sb.AppendLine(" )")
                                            StrSQL = sb.ToString()
                                            cmSQL.CommandText = StrSQL
                                            cmSQL.ExecuteNonQuery()

                                            '//Update All Per in TBLMASTER***********
                                            sb.Clear()
                                            sb.AppendLine(" Update TBLMASTER")
                                            sb.AppendLine(" Set ")
                                            sb.AppendLine(" Per = Qty * 100 / " & totalQty)
                                            sb.AppendLine(" Where MasterCode = " & PrepareStr(strEachPigmentCode) & " AND Revision = " & PrepareStr(strEachRevision))
                                            StrSQL = sb.ToString()
                                            cmSQL.CommandText = StrSQL
                                            cmSQL.ExecuteNonQuery()
                                        Else
                                            '//Insert TBLPigment
                                            sb.Clear()
                                            sb.AppendLine(" Insert TBLPigment ")
                                            sb.AppendLine(" Values (")
                                            sb.AppendLine(PrepareStr(strEachPigmentCode) & ", ")     'Column PIGMENTCode
                                            sb.AppendLine(PrepareStr(strEachRevision) & ", ")        'Column Revision
                                            sb.AppendLine(" '" & totalQty & "', ")       'Column Qty
                                            sb.AppendLine("'KG' , ")                     'Column Unit
                                            sb.AppendLine(" '" & strDate & "' ")         'Column Dateup
                                            sb.AppendLine(" )")

                                            sb.AppendLine(" ")

                                            '//Insert TBLConvert
                                            sb.AppendLine(" Insert TblConvert ")
                                            sb.AppendLine(" Values (")
                                            sb.AppendLine("'02' , ")                     'Column Type
                                            sb.AppendLine(PrepareStr("") & " , ")        'Column Final
                                            sb.AppendLine(PrepareStr(strEachPigmentCode) & " , ")    'Column Code
                                            sb.AppendLine(PrepareStr(strEachRevision) & " , ")       'Column Rev
                                            sb.AppendLine("'BT' , ")                     'Column UnitBig
                                            sb.AppendLine("'KG' , ")                     'Column UnitSmall
                                            sb.AppendLine("'1' , ")                      'Column BQty
                                            sb.AppendLine(" '" & totalQty & "' ")        'Column SQty
                                            sb.AppendLine(" )")

                                            sb.AppendLine(" ")

                                            '//Insert TBLMASTER
                                            sb.AppendLine(" Insert  TBLMASTER ")
                                            sb.AppendLine(" Values (")
                                            sb.AppendLine(PrepareStr(strEachPigmentCode) & ", ")        'Column MasterCode
                                            sb.AppendLine(PrepareStr(strEachRevision) & ", ")           'Column Revision
                                            sb.AppendLine(PrepareStr(strRMCode) & ", ")                 'Column RMCode
                                            sb.AppendLine(PrepareStr("") & ", ")                        'Column RmRevision
                                            sb.AppendLine(PrepareStr(dblRMQty) & ", ")                  'Column Qty
                                            sb.AppendLine("'KG', ")                                     'Column Unit
                                            sb.AppendLine(" '" & (dblRMQty * 100 / totalQty) & "'")   'Column Per
                                            sb.AppendLine(" )")

                                            StrSQL = sb.ToString()
                                            cmSQL.CommandText = StrSQL
                                            cmSQL.ExecuteNonQuery()
                                        End If
                                    Else 'Have no data (pigmentcode and revision) on any Table
                                        '//Insert TBLGroup
                                        sb.Clear()
                                        sb.AppendLine(" Insert  TBLGroup ")
                                        sb.AppendLine(" Values (")
                                        sb.AppendLine("'02' , ")                 'Column TypeCode
                                        sb.AppendLine(PrepareStr(strEachPigmentCode))       'Column Code
                                        sb.AppendLine(" )")

                                        sb.AppendLine(" ")

                                        '//Insert TBLPigment
                                        sb.AppendLine(" Insert  TBLPigment ")
                                        sb.AppendLine(" Values (")
                                        sb.AppendLine(PrepareStr(strEachPigmentCode) & ", ")     'Column PIGMENTCode
                                        sb.AppendLine(PrepareStr(strEachRevision) & ", ")        'Column Revision
                                        sb.AppendLine(" '" & totalQty & "', ")       'Column Qty
                                        sb.AppendLine("'KG' , ")                     'Column Unit
                                        sb.AppendLine(" '" & strDate & "' ")         'Column Dateup
                                        sb.AppendLine(" )")

                                        sb.AppendLine(" ")

                                        '//Insert TblConvert
                                        sb.AppendLine(" Insert  TblConvert ")
                                        sb.AppendLine(" Values (")
                                        sb.AppendLine("'02' , ")                     'Column Type
                                        sb.AppendLine(PrepareStr("") & " , ")        'Column Final
                                        sb.AppendLine(PrepareStr(strEachPigmentCode) & " , ")    'Column Code
                                        sb.AppendLine(PrepareStr(strEachRevision) & " , ")       'Column Rev
                                        sb.AppendLine("'BT' , ")                     'Column UnitBig
                                        sb.AppendLine("'KG' , ")                     'Column UnitSmall
                                        sb.AppendLine("'1' , ")                      'Column BQty
                                        sb.AppendLine(" '" & totalQty & "' ")        'Column SQty
                                        sb.AppendLine(" )")

                                        sb.AppendLine(" ")

                                        '//Insert TblMaster
                                        sb.AppendLine(" Insert  TBLMASTER ")
                                        sb.AppendLine(" Values (")
                                        sb.AppendLine(PrepareStr(strEachPigmentCode) & ", ")        'Column MasterCode
                                        sb.AppendLine(PrepareStr(strEachRevision) & ", ")           'Column Revision
                                        sb.AppendLine(PrepareStr(strRMCode) & ", ")                 'Column RMCode
                                        sb.AppendLine(PrepareStr("") & ", ")                        'Column RmRevision
                                        sb.AppendLine(PrepareStr(dblRMQty) & ", ")                  'Column Qty
                                        sb.AppendLine("'KG', ")                                     'Column Unit
                                        sb.AppendLine(" '" & (dblRMQty * 100 / totalQty) & "'")   'Column Per
                                        sb.AppendLine(" )")

                                        StrSQL = sb.ToString()
                                        cmSQL.CommandText = StrSQL
                                        cmSQL.ExecuteNonQuery()
                                    End If
                                End If 'If DTRow.Count > 0
                            End If 'If strEachPigmentCode.Length > 0
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

            LoadPIGMENT() 'ReQuery and set datagrid
            'View() 'Filter by condition
            frmOverlay.Dispose()
        End If 'If importDialog.ShowDialog() = Windows.Forms.DialogResult.OK
    End Sub

    Private Sub CmdExport_Click(sender As Object, e As EventArgs) Handles CmdExport.Click
        Dim arrColumn As String() = System.Configuration.ConfigurationManager.AppSettings("EXP_EXCEL_COLUMN_MASTER_PIGMENT").ToString().Split(New Char() {","c})
        Dim arrColumnHeader As String() = System.Configuration.ConfigurationManager.AppSettings("EXP_EXCEL_COLUMN_HEADER_MASTER_PIGMENT").ToString().Split(New Char() {","c})
        ExcelLib.Export(Me, GrdDV, TBL_PIGMENT, arrColumn, arrColumnHeader)
    End Sub
#End Region

#Region "DelPigment"
    Private Function ChkData() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Try
            strSQL &= " select count(*) from TblMaster "
            strSQL &= " where RMCode  = '" & GrdDV.Item(oldrow).Row("PigmentCode") & "'"
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
    Private Function ChkDel() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Try
            strSQL &= " select count(*) from TblPigment "
            strSQL &= " where PigmentCode  = '" & GrdDV.Item(oldrow).Row("PigmentCode") & "'"
            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar
            If i = 1 Then
                ChkDel = True
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
    Sub DelPigment()
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Try
            strSQL = " Delete TblPigment"
            strSQL &= " where PigmentCode = '" & GrdDV.Item(oldrow).Row("PigmentCode") & "'"
            strSQL &= " and Revision = '" & GrdDV.Item(oldrow).Row("Revision") & "'"
            strSQL &= "  "
            strSQL &= " Delete TblMaster"
            strSQL &= " where Mastercode = '" & GrdDV.Item(oldrow).Row("PigmentCode") & "'"
            strSQL &= " and Revision = '" & GrdDV.Item(oldrow).Row("Revision") & "'"
            strSQL &= "  "
            If ChkDel() Then
                strSQL &= " Delete TblGroup"
                strSQL &= " where Code = '" & GrdDV.Item(oldrow).Row("PigmentCode") & "'"
            Else
            End If
            strSQL &= "  "
            strSQL &= " Delete TblConvert"
            strSQL &= " where Code = '" & GrdDV.Item(oldrow).Row("PigmentCode") & "'"
            strSQL &= " and Rev = '" & GrdDV.Item(oldrow).Row("Revision") & "'"
            strSQL &= "  "

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

    Private Sub Changedata()
        If CheckBoxPigment.Checked = True Then
            GrdDV.RowFilter = " Pcode like '%" & CmbPigment.Text.Trim & "%'"
            DataGridPigment.DataSource = GrdDV
            CmbPigment.Enabled = True
        Else
            GrdDV.RowFilter = ""
            DataGridPigment.DataSource = GrdDV
            CmbPigment.Enabled = False
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

#Region "Import"
    Private Function ChkRMCodeMaster(ByVal ImportTable As DataTable) As Boolean
        Dim cnSQLRM As SqlConnection
        Dim cmSQLRM As SqlCommand
        Dim strSQL As String = String.Empty
        Dim ret As Boolean = False
        Dim strRmcodeBefore As String = String.Empty 'Previous rmCode
        Dim distinctImportTabale As New DataTable()

        Try
            ImportTable.DefaultView.Sort = "rmCode DESC" 'Sort datatable
            ImportTable = ImportTable.DefaultView.ToTable()
            ImportTable = ImportTable.DefaultView.ToTable(True, "rmCode") 'Distinct rmCode

            For x As Integer = 0 To ImportTable.Rows.Count - 1
                Dim rmCode As String = ImportTable.Rows(x)("rmCode").ToString().Trim()
                strSQL = String.Empty

                If x = 0 Then
                    'First record
                    strRmcodeBefore = String.Empty
                Else
                    strRmcodeBefore = ImportTable.Rows(x - 1)("rmCode").ToString().Trim()
                End If

                If rmCode.Length > 0 Then
                    If rmCode <> strRmcodeBefore Then
                        'Previous rmCode and current rmCode not equal
                        strSQL &= " SELECT count(*) "
                        strSQL &= " FROM TBLRM "
                        strSQL &= " WHERE RMcode  = '" & rmCode & "'"
                        cnSQLRM = New SqlConnection(C1.Strcon)
                        cnSQLRM.Open()
                        cmSQLRM = New SqlCommand(strSQL, cnSQLRM)
                        Dim i As Long = cmSQLRM.ExecuteScalar()

                        If i = 0 Then
                            'Not found in Master
                            cmSQLRM.Dispose()
                            cnSQLRM.Dispose()
                            Throw New System.Exception("This RM Code '" & rmCode & "' have no data on RM Master")
                        Else
                            cmSQLRM.Dispose()
                            cnSQLRM.Dispose()
                        End If

                        cnSQLRM.Close()
                    End If
                Else
                    ret = False
                    Throw New ApplicationException("RM Code is not empty.")
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
