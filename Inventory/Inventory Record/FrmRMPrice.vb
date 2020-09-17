#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Inventory_Record.Common
#End Region

Public Class FrmRMPrice

#Region "Declare"
    Inherits System.Windows.Forms.Form
    Public Shared GrdDV As New DataView
    Protected Const TBL_RM As String = "TBL_RM"
    Dim GrdDVType As New DataView
    Protected Const TBL_Type As String = "TBL_Type"

    Protected DefaultGridBorderStyle As BorderStyle
    Friend WithEvents CmdImport As System.Windows.Forms.Button
    Friend WithEvents CmdExport As System.Windows.Forms.Button
    Public Shared cm As CurrencyManager
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtName As System.Windows.Forms.TextBox
    Friend WithEvents CmdView As System.Windows.Forms.Button
    Friend WithEvents CmbType As System.Windows.Forms.ComboBox
    Friend WithEvents ChkType As System.Windows.Forms.CheckBox
    Friend WithEvents Txtcode As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmRMPrice))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.DataGridRM = New System.Windows.Forms.DataGrid()
        Me.CmdSave = New System.Windows.Forms.Button()
        Me.CmdClose = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TxtName = New System.Windows.Forms.TextBox()
        Me.CmdView = New System.Windows.Forms.Button()
        Me.CmbType = New System.Windows.Forms.ComboBox()
        Me.ChkType = New System.Windows.Forms.CheckBox()
        Me.Txtcode = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
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
        Me.GroupBox1.Size = New System.Drawing.Size(936, 464)
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
        Me.DataGridRM.Size = New System.Drawing.Size(930, 445)
        Me.DataGridRM.TabIndex = 0
        '
        'CmdSave
        '
        Me.CmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdSave.Image = CType(resources.GetObject("CmdSave.Image"), System.Drawing.Image)
        Me.CmdSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdSave.Location = New System.Drawing.Point(782, 544)
        Me.CmdSave.Name = "CmdSave"
        Me.CmdSave.Size = New System.Drawing.Size(80, 56)
        Me.CmdSave.TabIndex = 5
        Me.CmdSave.Text = "Save"
        Me.CmdSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdClose
        '
        Me.CmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdClose.Image = CType(resources.GetObject("CmdClose.Image"), System.Drawing.Image)
        Me.CmdClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdClose.Location = New System.Drawing.Point(862, 544)
        Me.CmdClose.Name = "CmdClose"
        Me.CmdClose.Size = New System.Drawing.Size(80, 56)
        Me.CmdClose.TabIndex = 6
        Me.CmdClose.Text = "Close"
        Me.CmdClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.Location = New System.Drawing.Point(600, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 16)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "R/M  DescName "
        '
        'TxtName
        '
        Me.TxtName.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtName.Location = New System.Drawing.Point(696, 7)
        Me.TxtName.Name = "TxtName"
        Me.TxtName.Size = New System.Drawing.Size(120, 20)
        Me.TxtName.TabIndex = 1
        '
        'CmdView
        '
        Me.CmdView.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdView.Image = CType(resources.GetObject("CmdView.Image"), System.Drawing.Image)
        Me.CmdView.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdView.Location = New System.Drawing.Point(848, 7)
        Me.CmdView.Name = "CmdView"
        Me.CmdView.Size = New System.Drawing.Size(80, 57)
        Me.CmdView.TabIndex = 3
        Me.CmdView.Text = "View"
        Me.CmdView.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmbType
        '
        Me.CmbType.Location = New System.Drawing.Point(96, 7)
        Me.CmbType.Name = "CmbType"
        Me.CmbType.Size = New System.Drawing.Size(128, 21)
        Me.CmbType.TabIndex = 0
        Me.CmbType.Text = "TypaName"
        '
        'ChkType
        '
        Me.ChkType.Location = New System.Drawing.Point(16, 5)
        Me.ChkType.Name = "ChkType"
        Me.ChkType.Size = New System.Drawing.Size(80, 24)
        Me.ChkType.TabIndex = 15
        Me.ChkType.Text = "TypeName"
        '
        'Txtcode
        '
        Me.Txtcode.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txtcode.Location = New System.Drawing.Point(696, 44)
        Me.Txtcode.Name = "Txtcode"
        Me.Txtcode.Size = New System.Drawing.Size(88, 20)
        Me.Txtcode.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.Location = New System.Drawing.Point(632, 46)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 16)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "R/M Code"
        '
        'CmdImport
        '
        Me.CmdImport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdImport.Image = CType(resources.GetObject("CmdImport.Image"), System.Drawing.Image)
        Me.CmdImport.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdImport.Location = New System.Drawing.Point(587, 544)
        Me.CmdImport.Name = "CmdImport"
        Me.CmdImport.Size = New System.Drawing.Size(80, 56)
        Me.CmdImport.TabIndex = 16
        Me.CmdImport.Text = "Import"
        Me.CmdImport.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdExport
        '
        Me.CmdExport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdExport.Image = CType(resources.GetObject("CmdExport.Image"), System.Drawing.Image)
        Me.CmdExport.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdExport.Location = New System.Drawing.Point(667, 544)
        Me.CmdExport.Name = "CmdExport"
        Me.CmdExport.Size = New System.Drawing.Size(80, 56)
        Me.CmdExport.TabIndex = 17
        Me.CmdExport.Text = "Export"
        Me.CmdExport.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'FrmRMPrice
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(952, 606)
        Me.Controls.Add(Me.CmdExport)
        Me.Controls.Add(Me.CmdImport)
        Me.Controls.Add(Me.CmbType)
        Me.Controls.Add(Me.ChkType)
        Me.Controls.Add(Me.Txtcode)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.CmdView)
        Me.Controls.Add(Me.TxtName)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CmdClose)
        Me.Controls.Add(Me.CmdSave)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(968, 645)
        Me.Name = "FrmRMPrice"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "R/M  WarehouseStock -"
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
#End Region

#Region "Delegate function"
    Public Shared Function MyGetSeqLine(ByVal row As Integer) As CellColor
        Dim c As CellColor
        c.ForeG = CInt(GrdDV.Item(row).Item(0))
        c.BackG = CInt(GrdDV.Item(row).Item(1))
        c.LfItem = Mid(GrdDV.Item(row).Item(3), 1, 4)
        Return c
    End Function
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

#Region "Function_Load"
    Private Sub LoadRM()
        Dim sb As New System.Text.StringBuilder()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sb.AppendLine(" SELECT  rm.Typecode,TypeName,rm.RMCode,descName,stdPrice,ActPrice,ut.unitcode,  ")
        sb.AppendLine(" ut.shortUnitName,ut.UnitName,1 as Qty ,")
        sb.AppendLine(" stdPrice as SPrice,ActPrice as APrice")
        sb.AppendLine(" FROM   TBLRM rm  ")
        sb.AppendLine(" LEFT OUTER JOIN TBLUNIT ut on rm.unit = ut.unitcode  ")
        sb.AppendLine(" LEFT OUTER JOIN TBLTYPE t on rm.Typecode = t.Typecode ")
        sb.AppendLine(" ORDER BY TypeName,descName,rm. Rmcode")
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
            .HeaderText = "TypeName"
            .MappingName = "TypeName"
            .Width = 85
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle1 As New DataGridColoredLine2
        With grdColStyle1
            .HeaderText = "Code"
            .MappingName = "RMCode"
            .Width = 75
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With
        Dim grdColStyle2 As New DataGridColoredLine2
        With grdColStyle2
            .HeaderText = "Name"
            .MappingName = "DescName"
            .Width = 110
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With
        Dim grdColStyle0_6 As New DataGridColoredLine2
        With grdColStyle0_6
            .HeaderText = "@ STD Price  "
            .MappingName = "STDPrice"
            .Width = 80
            .Format = "##,###,###.00"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        Dim grdColStyle0_7 As New DataGridColoredLine2
        With grdColStyle0_7
            .HeaderText = "@ ACT Price  "
            .MappingName = "ActPrice"
            .Width = 80
            .Format = "##,###,###.00"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With
        cm = CType(Me.BindingContext(DataGridRM.DataSource, DataGridRM.DataMember), CurrencyManager)
        Dim c As CheckRowHeader
        c = AddressOf CheckRowHeader
        Dim grdColStyle6 As New DataGridQtyBox(c)
        With grdColStyle6
            .HeaderText = "New Price  "
            .MappingName = "SPrice"
            .Format = "###,###.00"
            .Width = 80
            .Alignment = HorizontalAlignment.Right
            .NullText = ""
        End With
        Dim grdColStyle7 As New DataGridQtyBox(c)
        With grdColStyle7
            .HeaderText = "New Price  "
            .MappingName = "APrice"
            .Format = "###,###.00"
            .Width = 80
            .Alignment = HorizontalAlignment.Right
            .NullText = ""
        End With

        grdTableStyle1.GridColumnStyles.AddRange _
    (New DataGridColumnStyle() _
    {grdColStyle0, grdColStyle1, grdColStyle2,
 grdColStyle0_6, grdColStyle6, grdColStyle0_7, grdColStyle7})

        DataGridRM.TableStyles.Add(grdTableStyle1)
        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub


    Public Shared Function CheckRowHeader(ByVal row As Integer) As Boolean
        Dim c As Boolean = False
        'Debug.WriteLine("st seq : " + CStr(GrdItemDv.Item(row).Item("st_seq")) + "   row : " + CStr(row))
        Try
            If GrdDV.Item(row).Item("item_no").ToString.Trim = "" Then
                c = True
            Else
                c = False
            End If
        Catch ex As Exception
            c = False
        End Try

        Return c
    End Function


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

#Region "Form Event"
    Private Sub FrmRMPrice_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadRM()
        LoadCmbType()
        SetTotal() 'Set number of items
    End Sub
#End Region

#Region "Control Event"
    Private Sub CmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClose.Click
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        msg = "Do you want to Close. R/M Meterial Price." ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Price(KG) "   ' Define title.
        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            Me.Close()
        Else
            Exit Sub
        End If
    End Sub

    Private Sub CmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSave.Click
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        msg = "R/M Meterial Price" ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Price (KG) "   ' Define title.
        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            RM()
            View()
            MsgBox("Update Complete.", MsgBoxStyle.Information, "Price")
        Else
            Exit Sub
        End If
    End Sub

    Private Sub DataGridRM_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridRM.CurrentCellChanged
        oldrow = DataGridRM.CurrentCell.RowNumber
    End Sub


    Private Sub CmdView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdView.Click
        View()
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
                TxtName.Text = TxtName.Text.ToUpper
                SendKeys.Send("{TAB}")
            Case 32 ' space bar
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case Else
        End Select
    End Sub
#End Region
#Region "RM"
    Sub RM()
        Dim strsql As String = String.Empty
        Dim cn As New SqlConnection(C1.Strcon)
        Dim cmd As New SqlCommand(strsql, cn)
        cn.Open()
        Dim t1 As SqlTransaction = cn.BeginTransaction
        cmd.Transaction = t1
        Dim strDate As String
        Dim strTime As String
        Dim ddate() As String
        ddate = Split(Now.Date.ToShortDateString, "/")
        strDate = ddate(2) + ddate(1) + ddate(0)
        strTime = Format(Now.TimeOfDay.Hours, "00") & Format(Now.TimeOfDay.Minutes, "00")

        Try
            Dim aDr() As DataRow
            GrdDV.RowFilter = " "
            aDr = GrdDV.Table.Select(GrdDV.RowFilter)
            If UBound(aDr) < 0 Then
                Exit Sub
            End If
            Dim dr As DataRow
            For Each dr In aDr
                With dr
                    If IIf(.Item("RMCode") Is System.DBNull.Value, "", .Item("RMCode")) <> "" Then
                        If .Item("StdPrice") <> .Item("SPrice") Or .Item("ActPrice") <> .Item("APrice") Then
                            strsql = "UPDATE TblRM SET StdPrice=" & CStr(.Item("SPrice"))
                            strsql += " , ActPrice=" & CStr(.Item("APrice"))
                            strsql += " , Updatedate =" & CStr(strDate)
                            strsql += " , UpdateTime =" & CStr(strTime)
                            strsql += " where   RMCode =" & PrepareStr(.Item("RMCode"))
                            cmd.CommandText = strsql
                            cmd.ExecuteNonQuery()
                        End If
                    End If
                End With
            Next
            t1.Commit()
        Catch
            t1.Rollback()
            MsgBox("Rollback data")
        Finally
            cn.Close()
        End Try
    End Sub
#End Region

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

#Region "Export"
    Private Sub CmdExport_Click(sender As Object, e As EventArgs) Handles CmdExport.Click
        Dim arrColumn As String() = System.Configuration.ConfigurationManager.AppSettings("EXP_EXCEL_COLUMN_MASTER_RM_PRICE").ToString().Split(New Char() {","c})
        Dim arrColumnHeader As String() = System.Configuration.ConfigurationManager.AppSettings("EXP_EXCEL_COLUMN_HEADER_MASTER_RM_PRICE").ToString().Split(New Char() {","c})
        ExcelLib.Export(Me, GrdDV, TBL_RM, arrColumn, arrColumnHeader)
    End Sub
#End Region

#Region "Import"
    Private Sub CmdImport_Click(sender As Object, e As EventArgs) Handles CmdImport.Click
        Dim arrColumn As String() = System.Configuration.ConfigurationManager.AppSettings("IMP_EXCEL_COLUMN_MASTER_RM_PRICE").ToString().Split(New Char() {","c})
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
                            'Dim strTypeCode As String = PrepareStr(dtRec.Rows(i)("TypeCode").ToString().Trim())
                            'Dim unitCode As String = PrepareStr(dtRec.Rows(i)("UnitCode").ToString().Trim())
                            'Dim stdPrice As String = PrepareStr(dtRec.Rows(i)("StdPrice").ToString().Trim())
                            'Dim acrPrice As String = PrepareStr(dtRec.Rows(i)("ActPrice").ToString().Trim())
                            Dim isExists As Boolean = ChkDataImport(rmCode)

                            'Refer sub RM() of FrmAddRM
                            If isExists Then
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
#End Region
End Class
