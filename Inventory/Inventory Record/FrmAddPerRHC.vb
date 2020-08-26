#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Inventory_Record.Common
#End Region

Public Class FrmAddPerRHC

#Region "Declare"
    Inherits System.Windows.Forms.Form
    Public Shared GrdDV As New DataView
    Protected Const TBL_RM As String = "TBL_RM"
    Public Shared GrdDVComp As New DataView
    Protected Const TBL_Comp As String = "TBL_Comp"
    Protected DefaultGridBorderStyle As BorderStyle
    Public Shared cm As CurrencyManager
    Dim iTotal, qTotal, pTotal As Double
    Friend TCompound, TCode, TRev, TStep As String
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
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtCode As System.Windows.Forms.TextBox
    Friend WithEvents CmdClear As System.Windows.Forms.Button
    Friend WithEvents lblError As System.Windows.Forms.Label
    Friend WithEvents CheckAll As System.Windows.Forms.CheckBox
    Friend WithEvents TxtRev As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtCompound As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblerror2 As System.Windows.Forms.Label
    Friend WithEvents txtStep As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmAddPerRHC))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.DataGridRM = New System.Windows.Forms.DataGrid
        Me.CmdSave = New System.Windows.Forms.Button
        Me.CmdClose = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtName = New System.Windows.Forms.TextBox
        Me.CmdView = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtCode = New System.Windows.Forms.TextBox
        Me.CmdClear = New System.Windows.Forms.Button
        Me.lblError = New System.Windows.Forms.Label
        Me.CheckAll = New System.Windows.Forms.CheckBox
        Me.TxtRev = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtCompound = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.lblerror2 = New System.Windows.Forms.Label
        Me.txtStep = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
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
        Me.GroupBox1.Location = New System.Drawing.Point(8, 64)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(762, 496)
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
        Me.DataGridRM.Size = New System.Drawing.Size(756, 477)
        Me.DataGridRM.TabIndex = 0
        '
        'CmdSave
        '
        Me.CmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdSave.Image = CType(resources.GetObject("CmdSave.Image"), System.Drawing.Image)
        Me.CmdSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdSave.Location = New System.Drawing.Point(608, 562)
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
        Me.CmdClose.Location = New System.Drawing.Point(688, 562)
        Me.CmdClose.Name = "CmdClose"
        Me.CmdClose.Size = New System.Drawing.Size(80, 56)
        Me.CmdClose.TabIndex = 6
        Me.CmdClose.Text = "Close"
        Me.CmdClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.Location = New System.Drawing.Point(520, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 16)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "R/M Name "
        '
        'TxtName
        '
        Me.TxtName.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtName.Location = New System.Drawing.Point(592, 8)
        Me.TxtName.Name = "TxtName"
        Me.TxtName.Size = New System.Drawing.Size(120, 20)
        Me.TxtName.TabIndex = 2
        Me.TxtName.Text = ""
        '
        'CmdView
        '
        Me.CmdView.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdView.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.CmdView.Location = New System.Drawing.Point(712, 7)
        Me.CmdView.Name = "CmdView"
        Me.CmdView.Size = New System.Drawing.Size(48, 23)
        Me.CmdView.TabIndex = 3
        Me.CmdView.Text = "View"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 10)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 16)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Compound "
        '
        'TxtCode
        '
        Me.TxtCode.Location = New System.Drawing.Point(80, 8)
        Me.TxtCode.Name = "TxtCode"
        Me.TxtCode.ReadOnly = True
        Me.TxtCode.TabIndex = 0
        Me.TxtCode.Text = ""
        '
        'CmdClear
        '
        Me.CmdClear.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CmdClear.Image = CType(resources.GetObject("CmdClear.Image"), System.Drawing.Image)
        Me.CmdClear.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdClear.Location = New System.Drawing.Point(8, 562)
        Me.CmdClear.Name = "CmdClear"
        Me.CmdClear.Size = New System.Drawing.Size(80, 56)
        Me.CmdClear.TabIndex = 7
        Me.CmdClear.Text = "Clear"
        Me.CmdClear.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'lblError
        '
        Me.lblError.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblError.ForeColor = System.Drawing.Color.Red
        Me.lblError.Location = New System.Drawing.Point(184, 14)
        Me.lblError.Name = "lblError"
        Me.lblError.Size = New System.Drawing.Size(24, 8)
        Me.lblError.TabIndex = 8
        Me.lblError.Text = "***"
        Me.lblError.Visible = False
        '
        'CheckAll
        '
        Me.CheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CheckAll.Location = New System.Drawing.Point(104, 602)
        Me.CheckAll.Name = "CheckAll"
        Me.CheckAll.Size = New System.Drawing.Size(112, 16)
        Me.CheckAll.TabIndex = 8
        Me.CheckAll.Text = "Show All"
        Me.CheckAll.Visible = False
        '
        'TxtRev
        '
        Me.TxtRev.Location = New System.Drawing.Point(80, 32)
        Me.TxtRev.Name = "TxtRev"
        Me.TxtRev.ReadOnly = True
        Me.TxtRev.Size = New System.Drawing.Size(40, 20)
        Me.TxtRev.TabIndex = 13
        Me.TxtRev.Text = ""
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 34)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 16)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = " Rev."
        '
        'TxtCompound
        '
        Me.TxtCompound.Location = New System.Drawing.Point(320, 8)
        Me.TxtCompound.Name = "TxtCompound"
        Me.TxtCompound.ReadOnly = True
        Me.TxtCompound.Size = New System.Drawing.Size(80, 20)
        Me.TxtCompound.TabIndex = 1
        Me.TxtCompound.Text = ""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(232, 10)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 16)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Final Compound"
        '
        'lblerror2
        '
        Me.lblerror2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblerror2.ForeColor = System.Drawing.Color.Red
        Me.lblerror2.Location = New System.Drawing.Point(408, 14)
        Me.lblerror2.Name = "lblerror2"
        Me.lblerror2.Size = New System.Drawing.Size(24, 8)
        Me.lblerror2.TabIndex = 15
        Me.lblerror2.Text = "***"
        Me.lblerror2.Visible = False
        '
        'txtStep
        '
        Me.txtStep.Location = New System.Drawing.Point(192, 32)
        Me.txtStep.Name = "txtStep"
        Me.txtStep.ReadOnly = True
        Me.txtStep.Size = New System.Drawing.Size(40, 20)
        Me.txtStep.TabIndex = 20
        Me.txtStep.Text = ""
        Me.txtStep.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(152, 34)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(32, 16)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Step"
        '
        'FrmAddPerRHC
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(778, 624)
        Me.Controls.Add(Me.txtStep)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.lblerror2)
        Me.Controls.Add(Me.TxtCompound)
        Me.Controls.Add(Me.TxtRev)
        Me.Controls.Add(Me.TxtCode)
        Me.Controls.Add(Me.TxtName)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.CheckAll)
        Me.Controls.Add(Me.lblError)
        Me.Controls.Add(Me.CmdClear)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.CmdView)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CmdClose)
        Me.Controls.Add(Me.CmdSave)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmAddPerRHC"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "COMPOUND (MIXING)  "
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataGridRM, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

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

#Region "Function_Load"
    Private Sub LoadRM()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        StrSQL = "   SELECT   seq,final,mastercode,Revision"
        StrSQL &= " ,rmcode,weight mQty,RHC mRHC,per mPer"
        StrSQL &= "  FROM         TBLRHCDtl "
        StrSQL &= " where  Final = '" & TCompound.Trim & "'"
        StrSQL &= " and mastercode = '" & TCode.Trim & "'"
        StrSQL &= " and Revision = '" & TRev.Trim & "' "

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
        Dim grdColStyle1 As New DataGridColoredLine2
        With grdColStyle1
            .HeaderText = "Code"
            .MappingName = "RMCode"
            .Width = 100
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With

        cm = CType(Me.BindingContext(DataGridRM.DataSource, DataGridRM.DataMember), CurrencyManager)
        Dim c As CheckRowHeader
        c = AddressOf CheckRowHeader

        Dim grdColStyle2 As New DataGridQtyBox(c)
        With grdColStyle2
            .HeaderText = "RHC"
            .MappingName = "mRHC"
            .Format = "###,##0.000"
            .Width = 110
            .Alignment = HorizontalAlignment.Right
            .NullText = ""
        End With
        Dim grdColStyle3 As New DataGridQtyBox(c)
        With grdColStyle3
            .HeaderText = "%(Percent)"
            .MappingName = "mPer"
            .Width = 110
            .Format = "###,##0.000"
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Right
        End With

        grdTableStyle1.GridColumnStyles.AddRange _
(New DataGridColumnStyle() _
{grdColStyle1, grdColStyle2, grdColStyle3})

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

    Private Sub FrmAddPerRHC_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If CmdSave.Text = "Edit" Then
            TxtCode.Text = TCode
            TxtCompound.Text = TCompound
            TxtRev.Text = TRev
            txtStep.Text = TStep
            LoadRM()
        Else
            TxtCode.Text = TCode
            TxtCompound.Text = TCompound
            TxtRev.Text = TRev
            txtStep.Text = TStep
            LoadRM()
        End If
    End Sub

    Private Sub CmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClose.Click
        Me.Close()
    End Sub

    Private Sub CmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSave.Click
        If TxtCode.Text.Trim = "" Then
            TxtCode.Focus()
            lblError.Visible = True
            Exit Sub
        Else
            TxtCode.Text = TxtCode.Text.ToUpper
            lblError.Visible = False
        End If

        If TxtCompound.Text.Trim = "" Then
            TxtCompound.Focus()
            lblerror2.Visible = True
            Exit Sub
        Else
            TxtCompound.Text = TxtCompound.Text.ToUpper
            lblerror2.Visible = False
        End If

        TxtCode.Text = TxtCode.Text.ToUpper

        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        Dim aDr() As DataRow
        GrdDV.RowFilter = " mPer <> 0.000"
        aDr = GrdDV.Table.Select(GrdDV.RowFilter)
        iTotal = 0
        qTotal = 0
        CheckAll.Checked = False
        Dim dr As DataRow
        For Each dr In aDr
            With dr
                If IIf(.Item("RMCode") Is System.DBNull.Value, "", .Item("RMCode")) <> "" Then
                    iTotal = iTotal + .Item("mPer")
                End If
            End With
        Next

        msg = "Compound % Total :" & iTotal
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Compound"   ' Define title.
        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            RM()
            Me.Close()
        Else
            Exit Sub
        End If
    End Sub

    Private Sub DataGridRM_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridRM.CurrentCellChanged
        oldrow = DataGridRM.CurrentCell.RowNumber
    End Sub


    Private Sub CmdView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdView.Click
        GrdDV.RowFilter = "  descname like'%" & TxtName.Text.Trim & "%'"
        DataGridRM.DataSource = GrdDV
    End Sub

#Region "RM"
    Private Function iNoSeq(ByVal Ccode As String, ByVal rev As String) As Long
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim drSQL As SqlDataReader
        Dim strSQL As String
        Try

            ' Build Insert statement to insert 
            strSQL = "    SELECT  seq"
            strSQL &= "  FROM  TBLCompound  "
            strSQL &= "  Where FinalCompound  = '" & Ccode & "' and Revision ='" & rev & "'"
            strSQL &= " order by seq desc"
            cnSQL = New SqlConnection(C1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            cmSQL.ExecuteScalar()
            drSQL = cmSQL.ExecuteReader()
            If drSQL.HasRows Then
                If drSQL.Read() Then
                    iNoSeq = CInt(drSQL.Item("seq").ToString())
                End If
            End If

            ' Close and Clean up objects
            drSQL.Close()
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

    Private Function iqty(ByVal ii As String, ByVal jj As String) As Boolean
        If ii = jj Then
            iqty = False
        Else
            iqty = True
        End If
    End Function
    Sub RM()
        Dim strsql As String = String.Empty
        Dim cn As New SqlConnection(C1.Strcon)
        Dim cmd As New SqlCommand(strsql, cn)
        cn.Open()
        Dim t1 As SqlTransaction = cn.BeginTransaction
        cmd.Transaction = t1
        Dim strDate, SD() As String
        SD = Split(Date.Now.ToShortDateString, "/")
        strDate = SD(2) + SD(1) + SD(0)
        Try
            Dim aDr() As DataRow
            GrdDV.RowFilter = " mPer <> 0.000"
            aDr = GrdDV.Table.Select(GrdDV.RowFilter)
            If UBound(aDr) < 0 Then
                Exit Sub
            End If
            Dim dr As DataRow
            For Each dr In aDr
                With dr
                    If IIf(.Item("RMCode") Is System.DBNull.Value, "", .Item("RMCode")) <> "" Then
                        Dim RMstr As String
                        RMstr = .Item("RMCode")
                        strsql = "Update TBLRHCDtl "
                        strsql += " set Per =" & PrepareStr(.Item("mPer"))
                        strsql += " where Final = " & PrepareStr(TxtCompound.Text.Trim)
                        strsql += " and  Mastercode = " & PrepareStr(TxtCode.Text.Trim)
                        strsql += " and  Revision = " & PrepareStr(TxtRev.Text.Trim)
                        strsql += " and  RMCode = " & PrepareStr(.Item("RMCode"))
                        strsql += " and  RHC = " & PrepareStr(.Item("mRHC"))
                        cmd.CommandText = strsql
                        cmd.ExecuteNonQuery()
                    End If
                End With
            Next
            Try
                strsql = "  Update TBLCompound "
                strsql += " set Per = " & PrepareStr(iTotal)
                strsql += " where FinalCompound = " & PrepareStr(TxtCompound.Text.Trim)
                strsql += " and  Compcode = " & PrepareStr(TxtCode.Text.Trim)
                strsql += " and  Revision = " & PrepareStr(TxtRev.Text.Trim)
                cmd.CommandText = strsql
                cmd.ExecuteNonQuery()
            Catch ex As Exception
            End Try

            t1.Commit()
            MsgBox("Update Complete.", MsgBoxStyle.Information, "Compound Code")
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

    Private Sub TxtName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtName.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case Else
                'If Not IsNumeric(e.KeyChar) Then
                '    e.Handled = True
                'Else
                'End If
        End Select

    End Sub

    Private Sub CmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClear.Click
        LoadRM()
        TxtCode.Text = ""
        TxtName.Text = ""
    End Sub

    Private Sub CheckAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckAll.CheckedChanged
        If CheckAll.Checked = True Then
            GrdDV.RowFilter = " RMQty <> '' "
            DataGridRM.DataSource = GrdDV
        Else
            GrdDV.RowFilter = " RMQty <> 0.000 "
            DataGridRM.DataSource = GrdDV
        End If
    End Sub

    Private Sub TxtCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCode.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                TxtCode.Text = TxtCode.Text.ToUpper
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case Else
        End Select
    End Sub

    Private Sub TxtCompound_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCompound.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                TxtCompound.Text = TxtCompound.Text.ToUpper
                txtStep.Text = iNoSeq(TxtCompound.Text.Trim, TxtRev.Text.Trim) + 1
                LoadRM()
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case Else
        End Select
    End Sub

    Private Sub txtStep_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtStep.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case Else
                If Not IsNumeric(e.KeyChar) Then
                    e.Handled = True
                Else
                    If Len(sender.text) >= 1 Then
                        If txtStep.SelectionLength = Len(txtStep.Text) Then
                        Else
                            e.Handled = True
                        End If
                    End If
                End If
        End Select
    End Sub

End Class
