#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Inventory_Record.Common
Imports Excel = Microsoft.Office.Interop.Excel
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents DataGridRM As System.Windows.Forms.DataGrid
    Friend WithEvents CmdSave As System.Windows.Forms.Button
    Friend WithEvents CmdClose As System.Windows.Forms.Button
    Friend WithEvents CmdEdit As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtName As System.Windows.Forms.TextBox
    Friend WithEvents CmdView As System.Windows.Forms.Button
    Friend WithEvents BtDel As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Txtcode As System.Windows.Forms.TextBox
    Friend WithEvents ChkType As System.Windows.Forms.CheckBox
    Friend WithEvents CmbType As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmRM))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.DataGridRM = New System.Windows.Forms.DataGrid()
        Me.CmdSave = New System.Windows.Forms.Button()
        Me.CmdClose = New System.Windows.Forms.Button()
        Me.CmdEdit = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TxtName = New System.Windows.Forms.TextBox()
        Me.CmdView = New System.Windows.Forms.Button()
        Me.BtDel = New System.Windows.Forms.Button()
        Me.Txtcode = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ChkType = New System.Windows.Forms.CheckBox()
        Me.CmbType = New System.Windows.Forms.ComboBox()
        Me.CmdImport = New System.Windows.Forms.Button()
        Me.CmdExport = New System.Windows.Forms.Button()
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
        Me.GroupBox1.Location = New System.Drawing.Point(8, 72)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(880, 430)
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
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.Location = New System.Drawing.Point(558, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 16)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "R/M  DescName "
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
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.Location = New System.Drawing.Point(590, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 16)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "R/M Code"
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
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CmdEdit)
        Me.Controls.Add(Me.CmdClose)
        Me.Controls.Add(Me.CmdSave)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Txtcode)
        Me.Controls.Add(Me.Label2)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmRM"
        Me.Text = "R/M  WarehouseStock"
        Me.GroupBox1.ResumeLayout(False)
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
    Dim DialogFileExtension As String = System.Configuration.ConfigurationManager.AppSettings("DIALOG_FILE_EXT").ToString()
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
    End Sub
#End Region

#Region "Function_Load"
    Private Sub LoadRM()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        StrSQL = "   SELECT   rm.Typecode,Typename,rm.RMCode,descName,stdPrice,ActPrice,ut.unitcode, " &
                " shortUnitName,UnitName,isnull(RMQty,'0.000') as Qty ,isnull(Qty,'0.00') as Qunit " &
                " ,Qunit as unit  FROM   TBLRM rm " &
                " left outer join  TBLUNIT ut " &
                " on rm.unit = ut.unitcode " &
                " left outer join  TBLQTYUNIT qt " &
                "  on rm.RMCode = qt.RMCode" &
                " left outer join " &
                " TBLType t " &
                " on rm.Typecode = t.Typecode " &
                " order by Typename,descName,rm. Rmcode"

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
            strSQL = " Delete TblRM"
            strSQL &= " where RMCode = '" & GrdDV.Item(oldrow).Row("RMCode") & "'"
            strSQL &= "  "
            strSQL &= " Delete TblQtyUnit"
            strSQL &= " where RMCode = '" & GrdDV.Item(oldrow).Row("RMCode") & "'"
            strSQL &= "  "
            strSQL &= " Delete TblGroup"
            strSQL &= " where Code = '" & GrdDV.Item(oldrow).Row("RMCode") & "'"
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
#End Region

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
    End Sub

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
        Dim xlApp As Excel.Application = New Excel.Application()
        Dim importDialog As OpenFileDialog = New OpenFileDialog With {
            .Filter = DialogFileExtension
        }
        Dim openFile As String = String.Empty

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
            Return
        End If

        If importDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim xlWorkBook As Excel.Workbook
            Dim xlWorkSheet As Excel.Worksheet

            openFile = importDialog.FileName
            xlWorkBook = xlApp.Workbooks.Open(openFile)
            xlWorkSheet = xlWorkBook.Worksheets("sheet1")
            'display the cells value B2
            MsgBox(xlWorkSheet.Cells(2, 2).value)
            'edit the cell with new value
            xlWorkSheet.Cells(2, 2) = "http://vb.net-informations.com"
            xlWorkBook.Close()

            ExcelLib.ReleaseObject(xlWorkBook)
            ExcelLib.ReleaseObject(xlWorkSheet)
        End If

        xlApp.Quit()
        ExcelLib.ReleaseObject(xlApp)
    End Sub

    Private Sub CmdExport_Click(sender As Object, e As EventArgs) Handles CmdExport.Click
        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
        Dim xlApp As Excel.Application = New Excel.Application()
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Add()
        Dim xlWorkSheet As Excel.Worksheet = CType(xlWorkBook.Sheets("Sheet1"), Excel.Worksheet)
        Dim exportDialog As SaveFileDialog = New SaveFileDialog With {
            .Filter = DialogFileExtension
        }
        Dim pathSaveFile As String = String.Empty

        Try
            'Check number of record in DataGrid
            If GrdDV.Count <= 0 Then
                'Error
                MessageBox.Show("Export error" & vbCrLf & "No data for export!!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                'Export excel
                If exportDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    Dim xlRange As Excel.Range
                    Dim misValue As Object = Type.Missing
                    Dim dtRec As DataTable = GrdDV.ToTable(TBL_RM)

                    pathSaveFile = exportDialog.FileName
                    xlWorkSheet.Name = "Master RM"

                    Dim dtHead As DataGridTableStyle = DataGridRM.TableStyles(0) 'Header as DataGridRM
                    Dim dtTemp As DataTable = New DataTable("RM Temp") 'Temporary datable

                    'Create temporary datatable
                    For j As Integer = 0 To dtHead.GridColumnStyles.OfType(Of DataGridColoredLine2).Count - 1
                        If dtRec.Rows(0)(dtHead.GridColumnStyles.Item(j).MappingName).GetType().Equals(GetType(Decimal)) Then
                            dtTemp.Columns.Add(dtHead.GridColumnStyles.Item(j).MappingName, GetType(Decimal))
                        Else
                            dtTemp.Columns.Add(dtHead.GridColumnStyles.Item(j).MappingName, GetType(String))
                        End If
                    Next

                    'Set header
                    For j As Integer = 1 To dtHead.GridColumnStyles.OfType(Of DataGridColoredLine2).Count
                        xlWorkSheet.Cells(1, j) = dtHead.GridColumnStyles.Item(j - 1).HeaderText
                    Next

                    'Set data
                    For i As Integer = 0 To dtRec.Rows.Count - 1
                        Dim drData As DataRow = dtTemp.NewRow()
                        For j As Integer = 0 To dtTemp.Columns.Count - 1
                            drData(j) = dtRec.Rows(i)(dtTemp.Columns(j).ColumnName)
                        Next
                        dtTemp.Rows.Add(drData)
                    Next i

                    'Set range for data
                    Dim c1 As Excel.Range = CType(xlWorkSheet.Cells(2, 1), Excel.Range)
                    Dim c2 As Excel.Range = CType(xlWorkSheet.Cells(2 + dtTemp.Rows.Count - 1, dtTemp.Columns.Count), Excel.Range)
                    xlRange = xlWorkSheet.Range(c1, c2)

                    'Convert DataTable to Array Object
                    xlRange.Value2 = ExcelLib.ConvertDatatableToObject(dtTemp)

                    'Set autofit column
                    c1 = CType(xlWorkSheet.Cells(1, 1), Excel.Range)
                    c2 = CType(xlWorkSheet.Cells(1 + dtTemp.Rows.Count, dtTemp.Columns.Count), Excel.Range)
                    xlRange = xlWorkSheet.Range(c1, c2)
                    xlRange.Columns.AutoFit()

                    'Set off for display alerts
                    xlApp.DisplayAlerts = False
                    'Save excel
                    xlWorkBook.SaveAs(pathSaveFile, misValue, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, misValue, misValue, misValue, misValue, misValue)

                    MessageBox.Show("Export complete", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If 'If GrdDV.Count <= 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ExcelLib.ReleaseObject(xlWorkSheet)
            xlWorkBook.Close(False)
            ExcelLib.ReleaseObject(xlWorkBook)
            xlApp.Quit()

            Dim pid As Integer = 0
            Dim a As Integer = ExcelLib.GetWindowThreadProcessId(xlApp.Hwnd, pid)
            Dim p As Process = Process.GetProcessById(pid)
            p.Kill()

            ExcelLib.ReleaseObject(xlApp)
            GC.Collect()
        End Try
    End Sub
#End Region
End Class
