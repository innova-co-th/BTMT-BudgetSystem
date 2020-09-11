#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Inventory_Record.Common
#End Region

Public Class FrmRM

#Region "Declare"
    Inherits System.Windows.Forms.Form
    Dim GrdDV As New DataView
    Protected Const TBL_RM As String = "TBL_RM"
    Dim GrdDVType As New DataView
    Protected Const TBL_Type As String = "TBL_Type"
    Friend WithEvents CmdImport As System.Windows.Forms.Button
    Friend WithEvents CmdExport As System.Windows.Forms.Button

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
    Friend WithEvents GbData As System.Windows.Forms.GroupBox
    Friend WithEvents DataGridRM As System.Windows.Forms.DataGrid
    Friend WithEvents CmdSave As System.Windows.Forms.Button
    Friend WithEvents CmdClose As System.Windows.Forms.Button
    Friend WithEvents CmdEdit As System.Windows.Forms.Button
    Friend WithEvents LblName As System.Windows.Forms.Label
    Friend WithEvents TxtName As System.Windows.Forms.TextBox
    Friend WithEvents CmdView As System.Windows.Forms.Button
    Friend WithEvents BtDel As System.Windows.Forms.Button
    Friend WithEvents LblCode As System.Windows.Forms.Label
    Friend WithEvents Txtcode As System.Windows.Forms.TextBox
    Friend WithEvents ChkType As System.Windows.Forms.CheckBox
    Friend WithEvents CmbType As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmRM))
        Me.GbData = New System.Windows.Forms.GroupBox()
        Me.DataGridRM = New System.Windows.Forms.DataGrid()
        Me.CmdSave = New System.Windows.Forms.Button()
        Me.CmdClose = New System.Windows.Forms.Button()
        Me.CmdEdit = New System.Windows.Forms.Button()
        Me.LblName = New System.Windows.Forms.Label()
        Me.TxtName = New System.Windows.Forms.TextBox()
        Me.CmdView = New System.Windows.Forms.Button()
        Me.BtDel = New System.Windows.Forms.Button()
        Me.Txtcode = New System.Windows.Forms.TextBox()
        Me.LblCode = New System.Windows.Forms.Label()
        Me.ChkType = New System.Windows.Forms.CheckBox()
        Me.CmbType = New System.Windows.Forms.ComboBox()
        Me.CmdImport = New System.Windows.Forms.Button()
        Me.CmdExport = New System.Windows.Forms.Button()
        Me.GbData.SuspendLayout()
        CType(Me.DataGridRM, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GbData
        '
        Me.GbData.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GbData.Controls.Add(Me.DataGridRM)
        Me.GbData.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GbData.Location = New System.Drawing.Point(8, 72)
        Me.GbData.Name = "GbData"
        Me.GbData.Size = New System.Drawing.Size(880, 430)
        Me.GbData.TabIndex = 4
        Me.GbData.TabStop = False
        '
        'DataGridRM
        '
        Me.DataGridRM.DataMember = ""
        Me.DataGridRM.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridRM.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridRM.Location = New System.Drawing.Point(3, 16)
        Me.DataGridRM.Name = "DataGridRM"
        Me.DataGridRM.Size = New System.Drawing.Size(874, 411)
        Me.DataGridRM.TabIndex = 0
        '
        'CmdSave
        '
        Me.CmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdSave.Image = CType(resources.GetObject("CmdSave.Image"), System.Drawing.Image)
        Me.CmdSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdSave.Location = New System.Drawing.Point(656, 504)
        Me.CmdSave.Name = "CmdSave"
        Me.CmdSave.Size = New System.Drawing.Size(80, 56)
        Me.CmdSave.TabIndex = 5
        Me.CmdSave.Text = "Add"
        Me.CmdSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdClose
        '
        Me.CmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdClose.Image = CType(resources.GetObject("CmdClose.Image"), System.Drawing.Image)
        Me.CmdClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdClose.Location = New System.Drawing.Point(808, 504)
        Me.CmdClose.Name = "CmdClose"
        Me.CmdClose.Size = New System.Drawing.Size(75, 56)
        Me.CmdClose.TabIndex = 7
        Me.CmdClose.Text = "Close"
        Me.CmdClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdEdit
        '
        Me.CmdEdit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdEdit.Image = CType(resources.GetObject("CmdEdit.Image"), System.Drawing.Image)
        Me.CmdEdit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdEdit.Location = New System.Drawing.Point(736, 504)
        Me.CmdEdit.Name = "CmdEdit"
        Me.CmdEdit.Size = New System.Drawing.Size(75, 56)
        Me.CmdEdit.TabIndex = 6
        Me.CmdEdit.Text = "Edit"
        Me.CmdEdit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'LblName
        '
        Me.LblName.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblName.Location = New System.Drawing.Point(558, 16)
        Me.LblName.Name = "LblName"
        Me.LblName.Size = New System.Drawing.Size(96, 16)
        Me.LblName.TabIndex = 4
        Me.LblName.Text = "R/M  DescName "
        '
        'TxtName
        '
        Me.TxtName.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtName.Location = New System.Drawing.Point(654, 16)
        Me.TxtName.Name = "TxtName"
        Me.TxtName.Size = New System.Drawing.Size(120, 20)
        Me.TxtName.TabIndex = 1
        '
        'CmdView
        '
        Me.CmdView.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdView.Image = CType(resources.GetObject("CmdView.Image"), System.Drawing.Image)
        Me.CmdView.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdView.Location = New System.Drawing.Point(790, 12)
        Me.CmdView.Name = "CmdView"
        Me.CmdView.Size = New System.Drawing.Size(72, 56)
        Me.CmdView.TabIndex = 3
        Me.CmdView.Text = "View"
        Me.CmdView.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'BtDel
        '
        Me.BtDel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BtDel.Image = CType(resources.GetObject("BtDel.Image"), System.Drawing.Image)
        Me.BtDel.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtDel.Location = New System.Drawing.Point(16, 502)
        Me.BtDel.Name = "BtDel"
        Me.BtDel.Size = New System.Drawing.Size(80, 56)
        Me.BtDel.TabIndex = 7
        Me.BtDel.Text = "Delete"
        Me.BtDel.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Txtcode
        '
        Me.Txtcode.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txtcode.Location = New System.Drawing.Point(654, 48)
        Me.Txtcode.Name = "Txtcode"
        Me.Txtcode.Size = New System.Drawing.Size(88, 20)
        Me.Txtcode.TabIndex = 2
        '
        'LblCode
        '
        Me.LblCode.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblCode.Location = New System.Drawing.Point(590, 48)
        Me.LblCode.Name = "LblCode"
        Me.LblCode.Size = New System.Drawing.Size(64, 16)
        Me.LblCode.TabIndex = 8
        Me.LblCode.Text = "R/M Code"
        '
        'ChkType
        '
        Me.ChkType.Location = New System.Drawing.Point(16, 12)
        Me.ChkType.Name = "ChkType"
        Me.ChkType.Size = New System.Drawing.Size(80, 24)
        Me.ChkType.TabIndex = 10
        Me.ChkType.Text = "TypeName"
        '
        'CmbType
        '
        Me.CmbType.Location = New System.Drawing.Point(96, 14)
        Me.CmbType.Name = "CmbType"
        Me.CmbType.Size = New System.Drawing.Size(128, 21)
        Me.CmbType.TabIndex = 0
        Me.CmbType.Text = "TypaName"
        '
        'CmdImport
        '
        Me.CmdImport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdImport.Image = CType(resources.GetObject("CmdImport.Image"), System.Drawing.Image)
        Me.CmdImport.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdImport.Location = New System.Drawing.Point(482, 504)
        Me.CmdImport.Name = "CmdImport"
        Me.CmdImport.Size = New System.Drawing.Size(75, 56)
        Me.CmdImport.TabIndex = 11
        Me.CmdImport.Text = "Import"
        Me.CmdImport.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdExport
        '
        Me.CmdExport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdExport.Image = CType(resources.GetObject("CmdExport.Image"), System.Drawing.Image)
        Me.CmdExport.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdExport.Location = New System.Drawing.Point(557, 504)
        Me.CmdExport.Name = "CmdExport"
        Me.CmdExport.Size = New System.Drawing.Size(75, 56)
        Me.CmdExport.TabIndex = 12
        Me.CmdExport.Text = "Export"
        Me.CmdExport.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'FrmRM
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(896, 566)
        Me.Controls.Add(Me.CmdExport)
        Me.Controls.Add(Me.CmdImport)
        Me.Controls.Add(Me.CmbType)
        Me.Controls.Add(Me.ChkType)
        Me.Controls.Add(Me.BtDel)
        Me.Controls.Add(Me.CmdView)
        Me.Controls.Add(Me.TxtName)
        Me.Controls.Add(Me.LblName)
        Me.Controls.Add(Me.CmdEdit)
        Me.Controls.Add(Me.CmdClose)
        Me.Controls.Add(Me.CmdSave)
        Me.Controls.Add(Me.GbData)
        Me.Controls.Add(Me.Txtcode)
        Me.Controls.Add(Me.LblCode)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(912, 605)
        Me.Name = "FrmRM"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "R/M  WarehouseStock -"
        Me.GbData.ResumeLayout(False)
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

#Region "COMBOBOX"
    Sub LoadCmbType()
        Dim dt As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        StrSQL = "SELECT  TypeCode,Typename "
        StrSQL &= "  FROM  TblType "
        StrSQL &= "  Where Typecode not in ('02','03','04','05','06') "

        Dim C1 As New SQLData("ACCINV")
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dt = New DataTable
            DA.Fill(dt)
        Catch
        Finally
        End Try
        dt.TableName = TBL_Type
        GrdDVType = dt.DefaultView
        '************************************
        CmbType.DisplayMember = "Typename"
        CmbType.ValueMember = "TypeCode"
        CmbType.DataSource = dt
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
#End Region

#Region "Form Event"
    Private Sub FrmRM_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadCmbType()
        LoadRM()
        SetTotal() 'Set number of items
    End Sub
#End Region

#Region "Function_Load"
    Private Sub LoadRM()
        Dim sb As New System.Text.StringBuilder()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sb.AppendLine("SELECT rm.TypeCode, t.TypeName, rm.RMCode, rm.DescName, rm.StdPrice, rm.ActPrice,")
        sb.AppendLine(" ut.UnitCode, ut.shortUnitName, ut.UnitName,")
        sb.AppendLine(" isnull(qt.RMQty,'0.000') as Qty ,isnull(qt.Qty,'0.00') as Qunit, qt.QUnit as unit ")
        sb.AppendLine("FROM TBLRM rm ")
        sb.AppendLine("LEFT OUTER JOIN TBLUNIT ut on rm.Unit = ut.UnitCode ")
        sb.AppendLine("LEFT OUTER JOIN TBLQTYUNIT qt on rm.RMCode = qt.RMCode ")
        sb.AppendLine("LEFT OUTER JOIN TBLType t on rm.TypeCode = t.TypeCode ")
        sb.AppendLine("ORDER BY t.Typename, rm.DescName, rm.Rmcode")
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
        Dim grdColStyle0 As New DataGridColoredLine2
        With grdColStyle0
            .HeaderText = "Type"
            .MappingName = "TypeName"
            .Width = 75
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle1 As New DataGridColoredLine2
        With grdColStyle1
            .HeaderText = "Code"
            .MappingName = "RMCode"
            .Width = 100
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle2 As New DataGridColoredLine2
        With grdColStyle2
            .HeaderText = "Name"
            .MappingName = "DescName"
            .Width = 120
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle3 As New DataGridColoredLine2
        With grdColStyle3
            .HeaderText = "@Std Price (KG)"
            .MappingName = "StdPrice"
            .Width = 100
            .Format = "##,###,###.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle4 As New DataGridColoredLine2
        With grdColStyle4
            .HeaderText = "@Act Price (KG)"
            .MappingName = "ActPrice"
            .Width = 100
            .Format = "##,###,###.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle5_0 As New DataGridColoredLine2
        With grdColStyle5_0
            .HeaderText = "Qty"
            .MappingName = "Qty"
            .Width = 40
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle5 As New DataGridColoredLine2
        With grdColStyle5
            .HeaderText = "Unit"
            .MappingName = "shortUnitName"
            .Width = 65
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle6 As New DataGridColoredLine2
        With grdColStyle6
            .HeaderText = "Qty"
            .MappingName = "Qunit"
            .Width = 80
            .Format = "##,###,###.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle7 As New DataGridColoredLine2
        With grdColStyle7
            .HeaderText = "Unit"
            .MappingName = "Unit"
            .Width = 65
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With

        'Arrange column
        grdTableStyle1.GridColumnStyles.AddRange _
    (New DataGridColumnStyle() _
    {grdColStyle0, grdColStyle1, grdColStyle2, grdColStyle5_0, grdColStyle5,
    grdColStyle6, grdColStyle7, grdColStyle3, grdColStyle4})
        'Set column in DataGrid
        DataGridRM.TableStyles.Add(grdTableStyle1)
        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub

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

#Region "RM"
    Private Function ChkData() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Try
            strSQL &= " select count(*) from TblMaster "
            strSQL &= " where RMcode  = '" & GrdDV.Item(oldrow).Row("RMCode") & "'"
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
    Sub DelRM()
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Try
            '1.Table TblRM
            strSQL = " Delete TblRM"
            strSQL &= " where RMCode = '" & GrdDV.Item(oldrow).Row("RMCode") & "'"
            '2.Table TblQtyUnit
            strSQL &= "  "
            strSQL &= " Delete TblQtyUnit"
            strSQL &= " where RMCode = '" & GrdDV.Item(oldrow).Row("RMCode") & "'"
            '3.Table TblGrup
            strSQL &= "  "
            strSQL &= " Delete TblGroup"
            strSQL &= " where Code = '" & GrdDV.Item(oldrow).Row("RMCode") & "'"
            '4.Table TblConvert
            strSQL &= "  "
            strSQL &= " Delete TblConvert"
            strSQL &= " where Code = '" & GrdDV.Item(oldrow).Row("RMCode") & "'"

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

    Sub View()
        If ChkType.Checked = True Then
            GrdDV.RowFilter = " descname like'%" & TxtName.Text.Trim & "%'" _
                            & " and RMcode like'%" & Txtcode.Text.Trim & "%'" _
                            & " and Typecode like'%" & CmbType.SelectedValue & "%'"
            DataGridRM.DataSource = GrdDV
        Else
            GrdDV.RowFilter = " descname like'%" & TxtName.Text.Trim & "%'" _
                                       & " and RMcode like'%" & Txtcode.Text.Trim & "%'"
            DataGridRM.DataSource = GrdDV
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

#Region "Control Event"
    Private Sub CmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClose.Click
        Me.Close()
    End Sub

    Private Sub CmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdEdit.Click
        Dim faddRM As New FrmAddRM
        faddRM.CmdSave.Text = "Edit"
        faddRM.TxtName.Text = GrdDV.Item(oldrow).Row("RMCode")
        faddRM.TxtName.Enabled = False
        faddRM.TxtDesc.Text = GrdDV.Item(oldrow).Row("DescName")
        faddRM.TxtSPrice.Text = GrdDV.Item(oldrow).Row("StdPrice")
        faddRM.TxtAPrice.Text = GrdDV.Item(oldrow).Row("ActPrice")
        faddRM.unittext = GrdDV.Item(oldrow).Row("Unitname")
        faddRM.unitcode = GrdDV.Item(oldrow).Row("Unitcode")
        faddRM.TxtQty.Text = GrdDV.Item(oldrow).Row("QUnit")
        faddRM.TxtRMQty.Text = GrdDV.Item(oldrow).Row("QTY")
        faddRM.ShowDialog()
        LoadRM()
        View()
    End Sub

    Private Sub CmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSave.Click
        Dim faddRM As New FrmAddRM
        faddRM.CmdSave.Text = "Save"
        faddRM.ShowDialog()
        LoadRM()
        View()
    End Sub

    Private Sub CmdView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdView.Click
        View()
    End Sub

    Private Sub BtDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtDel.Click
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        msg = "Delete R/M Meterial :" & GrdDV.Item(oldrow).Row("RMCode") ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "R/M Warehouse"   ' Define title.

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            If ChkData() Then
                MsgBox("It's have Usage , Can't Delete. Please contact IS.", MsgBoxStyle.Information, "Delete R/M ")
            Else
                DelRM()
            End If
        Else
            Exit Sub
        End If
        LoadRM()
        GrdDV.RowFilter = " descname like'%" & TxtName.Text.Trim & "%'"
        DataGridRM.DataSource = GrdDV
        SetTotal() 'Set number of items
    End Sub

    Private Sub DataGridRM_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridRM.CurrentCellChanged
        oldrow = DataGridRM.CurrentCell.RowNumber
    End Sub

    Private Sub DataGridRM_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridRM.DoubleClick
        Dim faddRM As New FrmAddRM
        faddRM.CmdSave.Text = "Edit"
        faddRM.TxtName.Text = GrdDV.Item(oldrow).Row("RMCode")
        faddRM.TxtName.Enabled = False
        faddRM.TxtDesc.Text = GrdDV.Item(oldrow).Row("DescName")
        faddRM.TxtSPrice.Text = GrdDV.Item(oldrow).Row("StdPrice")
        faddRM.TxtAPrice.Text = GrdDV.Item(oldrow).Row("ActPrice")
        faddRM.unittext = GrdDV.Item(oldrow).Row("Unitname")
        faddRM.unitcode = GrdDV.Item(oldrow).Row("Unitcode")
        faddRM.TxtQty.Text = GrdDV.Item(oldrow).Row("QUnit")
        faddRM.TxtRMQty.Text = GrdDV.Item(oldrow).Row("QTY")
        faddRM.ShowDialog()
        LoadRM()
        GrdDV.RowFilter = " descname like'%" & TxtName.Text.Trim & "%'"
        DataGridRM.DataSource = GrdDV
        SetTotal() 'Set number of items
    End Sub

    Private Sub ChkType_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkType.CheckedChanged
        View()
    End Sub

    Private Sub CmbType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbType.SelectedIndexChanged
        View()
    End Sub

    Private Sub Txtcode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Txtcode.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                Txtcode.Text = Txtcode.Text.ToUpper
                SendKeys.Send("{TAB}")
            Case 32 ' space bar
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case Else
        End Select
    End Sub

    Private Sub TxtName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtName.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                TxtName.Text = Txtcode.Text.ToUpper
                SendKeys.Send("{TAB}")
            Case 32 ' space bar
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case Else
        End Select
    End Sub

    Private Sub CmdImport_Click(sender As Object, e As EventArgs) Handles CmdImport.Click
        Dim arrColumn As String() = System.Configuration.ConfigurationManager.AppSettings("EXCEL_COLUMN_MASTER_RM").ToString().Split(New Char() {","c})
        Dim importDialog As OpenFileDialog = New OpenFileDialog With {
            .Filter = System.Configuration.ConfigurationManager.AppSettings("DIALOG_FILE_EXT").ToString()
        }
        Dim dtRec As DataTable
        Dim sb As New System.Text.StringBuilder()
        Dim frmOverlay As New Form()

        If importDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            'Create loading of overlay
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

                        For i As Integer = 0 To dtRec.Rows.Count - 1
                            'Check RMCode
                            Dim rmCode As String = dtRec.Rows(i)("RMCode").ToString().Trim()
                            Dim strTypeCode As String = PrepareStr(dtRec.Rows(i)("TypeCode").ToString().Trim())
                            Dim unitCode As String = PrepareStr(dtRec.Rows(i)("UnitCode").ToString().Trim())
                            Dim qty As String = PrepareStr(dtRec.Rows(i)("Qty"))
                            Dim qUnit As String = PrepareStr(dtRec.Rows(i)("Qunit"))
                            Dim isExists As Boolean = ChkDataImport(rmCode)

                            'Refer sub RM() of FrmAddRM
                            If Not isExists Then
                                'Insert
                                sb.Clear()
                                '1.Table TblRM
                                sb.AppendLine(" Insert  TblRM ")
                                sb.AppendLine(" Values (")
                                sb.AppendLine(" '" & rmCode & "',") 'Column RMCode(PK)
                                sb.AppendLine(PrepareStr(dtRec.Rows(i)("DescName").ToString().Trim()) & ",") 'Column DescName
                                sb.AppendLine(PrepareStr(dtRec.Rows(i)("StdPrice")) & ",") 'Column StdPrice
                                sb.AppendLine(PrepareStr(dtRec.Rows(i)("ActPrice")) & ",") 'Column ActPrice
                                sb.AppendLine(strTypeCode & ",") 'Column TypeCode
                                sb.AppendLine(unitCode & ",") 'Column Unit
                                sb.AppendLine(PrepareStr(strDate) & ",") 'Column UpdateDate
                                sb.AppendLine(PrepareStr(iTime)) 'Column UpdateTime
                                sb.AppendLine(" )")
                                StrSQL = sb.ToString()
                                cmSQL.CommandText = StrSQL
                                cmSQL.ExecuteNonQuery()

                                sb.Clear()
                                '2.Table TblGroup
                                sb.AppendLine(" Insert  TblGroup ")
                                sb.AppendLine(" Values (")
                                sb.AppendLine(strTypeCode & ",") 'Column TypeCode
                                sb.AppendLine(" '" & rmCode & "'") 'Column Code
                                sb.AppendLine(" )")
                                StrSQL = sb.ToString()
                                cmSQL.CommandText = StrSQL
                                cmSQL.ExecuteNonQuery()

                                sb.Clear()
                                '3.Table TblConvert
                                sb.AppendLine(" Insert TBLConvert ")
                                sb.AppendLine(" Values (")
                                sb.AppendLine(strTypeCode & ",") 'Column Type
                                sb.AppendLine(PrepareStr(String.Empty) & ",") 'Column Final
                                sb.AppendLine(" '" & rmCode & "',") 'Column Code
                                sb.AppendLine(PrepareStr(String.Empty) & ",") 'Column Rev
                                sb.AppendLine(unitCode & ",") 'Column UnitBig
                                sb.AppendLine(PrepareStr("KG") & ",") 'Column UnitSmall
                                sb.AppendLine(qty & ",") 'Column BQty
                                sb.AppendLine(qUnit) 'Column SQty
                                sb.AppendLine(" )")
                                StrSQL = sb.ToString()
                                cmSQL.CommandText = StrSQL
                                cmSQL.ExecuteNonQuery()

                                sb.Clear()
                                '4.Table TblQtyUnit
                                If CDec(dtRec.Rows(i)("Qunit")) <> 0 Then
                                    sb.AppendLine(" Insert  TblQtyUnit ")
                                    sb.AppendLine(" Values (")
                                    sb.AppendLine(" '" & rmCode & "',") 'Column RMCode(PK)
                                    sb.AppendLine(qty & ",") 'Column RMQty
                                    sb.AppendLine(unitCode & ",") 'Column UnitCode
                                    sb.AppendLine(qUnit & ",") 'Column Qty
                                    sb.AppendLine(PrepareStr("KG") & ",") 'Column QUnit
                                    sb.AppendLine(PrepareStr(strDate) & ",") 'Column UpdateDate
                                    sb.AppendLine(PrepareStr(iTime)) 'UpdateTime
                                    sb.AppendLine(" )")
                                Else
                                    sb.AppendLine(" Insert  TblQtyUnit ")
                                    sb.AppendLine(" Values (")
                                    sb.AppendLine(" '" & rmCode & "',")
                                    sb.AppendLine(qty & ",")
                                    sb.AppendLine(unitCode & ",")
                                    sb.AppendLine("0,") 'Force 0
                                    sb.AppendLine(PrepareStr("KG") & ",")
                                    sb.AppendLine(PrepareStr(strDate) & ",")
                                    sb.AppendLine(PrepareStr(iTime))
                                    sb.AppendLine(" )")
                                End If

                                StrSQL = sb.ToString()
                                cmSQL.CommandText = StrSQL
                                cmSQL.ExecuteNonQuery()
                            Else
                                'Update
                                sb.Clear()
                                '1.Table TblRM
                                sb.AppendLine(" Update TblRM")
                                sb.AppendLine(" Set ")
                                sb.AppendLine(" descName = '" & dtRec.Rows(i)("DescName").ToString().Trim() & "'")
                                sb.AppendLine(" , StdPrice = '" & dtRec.Rows(i)("StdPrice") & "'")
                                sb.AppendLine(" , ActPrice = '" & dtRec.Rows(i)("ActPrice") & "'")
                                sb.AppendLine(" , Unit = '" & dtRec.Rows(i)("UnitCode").ToString().Trim() & "'")
                                sb.AppendLine(" Where RMCode = '" & rmCode & "'")
                                StrSQL = sb.ToString()
                                cmSQL.CommandText = StrSQL
                                cmSQL.ExecuteNonQuery()

                                sb.Clear()
                                '2.Table TblQtyUnit
                                sb.AppendLine(" Update TblQtyUnit")
                                sb.AppendLine(" Set ")
                                sb.AppendLine(" Qty = '" & dtRec.Rows(i)("Qunit") & "'")
                                sb.AppendLine(" , RMQty = '" & dtRec.Rows(i)("Qty") & "'")
                                sb.AppendLine(" , UnitCode = '" & dtRec.Rows(i)("UnitCode").ToString().Trim() & "'")
                                sb.AppendLine(" Where RMCode = '" & rmCode & "'")
                                StrSQL = sb.ToString()
                                cmSQL.CommandText = StrSQL
                                cmSQL.ExecuteNonQuery()

                                sb.Clear()
                                '3.Table TblConvert
                                sb.AppendLine(" Update TblConvert")
                                sb.AppendLine(" Set ")
                                sb.AppendLine(" SQty = '" & dtRec.Rows(i)("Qunit") & "'")
                                sb.AppendLine(" , BQty = '" & dtRec.Rows(i)("Qty") & "'")
                                sb.AppendLine(" , UnitBig = '" & dtRec.Rows(i)("UnitCode").ToString().Trim() & "'")
                                sb.AppendLine(" Where Code = '" & rmCode & "'")
                                sb.AppendLine(" And  UnitBig = '" & dtRec.Rows(i)("UnitCode").ToString() & "'")
                                StrSQL = sb.ToString()
                                cmSQL.CommandText = StrSQL
                                cmSQL.ExecuteNonQuery()
                            End If ' If Not isExists
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

            LoadRM() 'ReQuery and set datagrid
            View() 'Filter by condition
            frmOverlay.Dispose()
        End If 'If importDialog.ShowDialog() = Windows.Forms.DialogResult.OK
    End Sub

    Private Sub CmdExport_Click(sender As Object, e As EventArgs) Handles CmdExport.Click
        Dim arrColumn As String() = System.Configuration.ConfigurationManager.AppSettings("EXCEL_COLUMN_MASTER_RM").ToString().Split(New Char() {","c})
        ExcelLib.Export(Me, GrdDV, TBL_RM, arrColumn)
    End Sub
#End Region

#Region "Import"
    Private Function ChkDataImport(rmCode As String) As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Dim ret As Boolean = False

        Try
            strSQL &= " SELECT count(*) "
            strSQL &= " FROM TblRM "
            strSQL &= " WHERE RMcode  = '" & rmCode & "'"
            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            Dim i As Long = cmSQL.ExecuteScalar()
            If i <> 0 Then
                ret = True
            End If

            cmSQL.Dispose()
            cnSQL.Dispose()
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
