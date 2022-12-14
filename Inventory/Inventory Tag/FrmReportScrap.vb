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

Public Class FrmReportScrap
#Region "Declare"
    Inherits System.Windows.Forms.Form
    Dim GrdDV As New DataView
    Protected Const TBL_RM As String = "TBL_RM"
    Public Shared tb1 As New DataTable

    Protected DefaultGridBorderStyle As BorderStyle
    Dim C1 As New SQLData("ACCINV")
    Dim StrData As String
    Friend Username As String
    Friend sType As String
    Friend sMType, sName As String
    Friend sPeriod1, sPeriod2 As String
    Friend sLoc, sSec As String
    Friend sCODE As String
    Friend sTrxPeriod As String
    Friend sTrx1, sTrx2 As String
    Friend sTag1, sTag2 As String
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
    Friend WithEvents DGView As System.Windows.Forms.DataGrid
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LblTotal As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lblstd As System.Windows.Forms.Label
    Friend WithEvents lblatc As System.Windows.Forms.Label
    Friend WithEvents CmdPrint As System.Windows.Forms.Button
    Friend WithEvents txtAdj As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents pbLoading As PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmReportScrap))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.pbLoading = New System.Windows.Forms.PictureBox()
        Me.DGView = New System.Windows.Forms.DataGrid()
        Me.CmdPrint = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtAdj = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.LblTotal = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lblstd = New System.Windows.Forms.Label()
        Me.lblatc = New System.Windows.Forms.Label()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.GroupBox1.SuspendLayout()
        CType(Me.pbLoading, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DGView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.pbLoading)
        Me.GroupBox1.Controls.Add(Me.DGView)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 72)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(712, 600)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'pbLoading
        '
        Me.pbLoading.Image = CType(resources.GetObject("pbLoading.Image"), System.Drawing.Image)
        Me.pbLoading.Location = New System.Drawing.Point(6, 45)
        Me.pbLoading.Name = "pbLoading"
        Me.pbLoading.Size = New System.Drawing.Size(150, 150)
        Me.pbLoading.TabIndex = 1
        Me.pbLoading.TabStop = False
        '
        'DGView
        '
        Me.DGView.BackgroundColor = System.Drawing.Color.LightGray
        Me.DGView.DataMember = ""
        Me.DGView.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DGView.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DGView.Location = New System.Drawing.Point(3, 16)
        Me.DGView.Name = "DGView"
        Me.DGView.PreferredColumnWidth = 95
        Me.DGView.ReadOnly = True
        Me.DGView.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.DGView.Size = New System.Drawing.Size(706, 581)
        Me.DGView.TabIndex = 0
        '
        'CmdPrint
        '
        Me.CmdPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdPrint.Image = CType(resources.GetObject("CmdPrint.Image"), System.Drawing.Image)
        Me.CmdPrint.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdPrint.Location = New System.Drawing.Point(640, 8)
        Me.CmdPrint.Name = "CmdPrint"
        Me.CmdPrint.Size = New System.Drawing.Size(75, 56)
        Me.CmdPrint.TabIndex = 1
        Me.CmdPrint.Text = "Print"
        Me.CmdPrint.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label1.Location = New System.Drawing.Point(72, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 16)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "ADJUST"
        '
        'txtAdj
        '
        Me.txtAdj.Location = New System.Drawing.Point(144, 48)
        Me.txtAdj.Name = "txtAdj"
        Me.txtAdj.Size = New System.Drawing.Size(88, 20)
        Me.txtAdj.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label2.Location = New System.Drawing.Point(72, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 16)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "TOTAL"
        '
        'LblTotal
        '
        Me.LblTotal.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.LblTotal.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.LblTotal.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.LblTotal.Location = New System.Drawing.Point(144, 16)
        Me.LblTotal.Name = "LblTotal"
        Me.LblTotal.Size = New System.Drawing.Size(152, 16)
        Me.LblTotal.TabIndex = 5
        Me.LblTotal.Text = "0"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label4.Location = New System.Drawing.Point(304, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 16)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "AMOUNT"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label5.Location = New System.Drawing.Point(304, 48)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(56, 16)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "AMOUNT"
        '
        'lblstd
        '
        Me.lblstd.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblstd.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.lblstd.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblstd.Location = New System.Drawing.Point(376, 16)
        Me.lblstd.Name = "lblstd"
        Me.lblstd.Size = New System.Drawing.Size(152, 16)
        Me.lblstd.TabIndex = 9
        Me.lblstd.Text = "0"
        '
        'lblatc
        '
        Me.lblatc.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblatc.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.lblatc.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblatc.Location = New System.Drawing.Point(376, 48)
        Me.lblatc.Name = "lblatc"
        Me.lblatc.Size = New System.Drawing.Size(152, 16)
        Me.lblatc.TabIndex = 10
        Me.lblatc.Text = "0"
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        Me.BackgroundWorker1.WorkerSupportsCancellation = True
        '
        'FrmReportScrap
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ClientSize = New System.Drawing.Size(728, 686)
        Me.Controls.Add(Me.lblatc)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.LblTotal)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtAdj)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CmdPrint)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.lblstd)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmReportScrap"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Report"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.pbLoading, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DGView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

#Region "CONSTANT"
    Dim DT As New DataTable
    Dim StrSQL As String
    Dim oldrow As Integer
#End Region

#Region "Propeties"
    ''' <summary>
    ''' Flag for check loading
    ''' </summary>
    ''' <returns>Boolean</returns>
    Public Property IsLoad As Boolean
#End Region

#Region "Function_Load"
    Private Sub LoadData()
        StrSQL &= " select  rm.rmcode,Round(isnull(QQty,0),2) TOTAL,round(StdPrice,4) StdPrice "
        StrSQL &= " ,Round(isnull(QQty,0)*StdPrice,2) SAMOUNT,round(ActPrice,4) ActPrice,Round(isnull(QQty,0)*ActPrice,2) AAMOUNT from "
        StrSQL &= "  TBLRM rm"
        StrSQL &= " left outer join "
        StrSQL &= "  ("
        StrSQL &= "  select  rmcode,Sum(QQty) qqty  from ("
        StrSQL &= "  select * "
        StrSQL &= "  from ScarpRmTag1 "
        StrSQL &= "  where period  = '" & sTrxPeriod.Trim() & "' "
        If sType <> "" Then
            StrSQL &= "  and Typecode = '" & sType.Trim() & "' "
        End If
        If sLoc <> "" Then
            StrSQL &= "  and Loc = ' " & sLoc.Trim() & "' "
        End If
        If sPeriod1 <> "" Then
            StrSQL &= "   and trxyear > = '" & sPeriod1.Trim() & "' "
        End If
        If sPeriod2 <> "" Then
            StrSQL &= "  and trxyear < = '" & sPeriod2.Trim() & "' "
        End If
        If sMType <> "" Then
            StrSQL &= "  and MaterialType = '" & sMType.Trim() & " ' "
        End If
        If sCODE <> "" Then
            StrSQL &= "  and CODE = '" & sCODE.Trim() & "' "
        End If
        If sTag1 <> "" And sTag2 <> "" Then
            StrSQL &= "  and Tagno >= '" & sTag1.Trim() & "' "
            StrSQL &= "  and Tagno <= '" & sTag2.Trim() & "' "
        End If

        StrSQL &= " union"
        StrSQL &= " select * "
        StrSQL &= " from ScarpRmTag2  "
        StrSQL &= " where period  = '" & sTrxPeriod.Trim() & "' "
        If sType <> "" Then
            StrSQL &= "  and Typecode = '" & sType.Trim() & "' "
        End If
        If sLoc <> "" Then
            StrSQL &= " and Loc = '" & sLoc.Trim() & "' "
        End If
        If sPeriod1 <> "" Then
            StrSQL &= "  and trxyear > = '" & sPeriod1.Trim() & "' "
        End If
        If sPeriod2 <> "" Then
            StrSQL &= "  and trxyear < = '" & sPeriod2.Trim() & "' "
        End If
        If sMType <> "" Then
            StrSQL &= "  and MaterialType = '" & sMType.Trim() & "' "
        End If
        If sCODE <> "" Then
            StrSQL &= "  and CODE = '" & sCODE.Trim() & "' "
        End If
        If sTag1 <> "" And sTag2 <> "" Then
            StrSQL &= "   and Tagno >= '" & sTag1.Trim() & "' "
            StrSQL &= "   and Tagno <= '" & sTag2.Trim() & "' "
        End If


        StrSQL &= "  union"
        StrSQL &= "  select * "
        StrSQL &= "  from ScarpRmTag3 "
        StrSQL &= "  where period  = '" & sTrxPeriod.Trim() & "' "
        If sType <> "" Then
            StrSQL &= "  and Typecode = '" & sType.Trim() & "' "
        End If
        If sLoc <> "" Then
            StrSQL &= " and Loc = '" & sLoc.Trim() & " ' "
        End If
        If sPeriod1 <> "" Then
            StrSQL &= "  and trxyear > = '" & sPeriod1.Trim() & " ' "
        End If
        If sPeriod2 <> "" Then
            StrSQL &= "  and trxyear < = '" & sPeriod2.Trim() & "' "
        End If
        If sMType <> "" Then
            StrSQL &= "  and MaterialType = '" & sMType.Trim() & "' "
        End If
        If sCODE <> "" Then
            StrSQL &= "  and CODE = '" & sCODE.Trim() & "' "
        End If
        If sTag1 <> "" And sTag2 <> "" Then
            StrSQL &= "   and Tagno >= '" & sTag1.Trim() & "' "
            StrSQL &= "   and Tagno <= '" & sTag2.Trim() & "' "
        End If

        StrSQL &= "  union"
        StrSQL &= "  select * "
        StrSQL &= "  from ScarpRmTag4 "
        StrSQL &= "  where period  = '" & sTrxPeriod.Trim() & "' "
        If sType <> "" Then
            StrSQL &= "  and Typecode = '" & sType.Trim() & "' "
        End If
        If sLoc <> "" Then
            StrSQL &= " and Loc = '" & sLoc.Trim() & " ' "
        End If
        If sPeriod1 <> "" Then
            StrSQL &= "  and trxyear > = '" & sPeriod1.Trim() & " ' "
        End If
        If sPeriod2 <> "" Then
            StrSQL &= "  and trxyear < = '" & sPeriod2.Trim() & "' "
        End If
        If sMType <> "" Then
            StrSQL &= "  and MaterialType = '" & sMType.Trim() & "' "
        End If
        If sCODE <> "" Then
            StrSQL &= "  and CODE = '" & sCODE.Trim() & "' "
        End If
        If sTag1 <> "" And sTag2 <> "" Then
            StrSQL &= "   and Tagno >= '" & sTag1.Trim() & "' "
            StrSQL &= "   and Tagno <= '" & sTag2.Trim() & "' "
        End If
        StrSQL &= "   )xx"
        StrSQL &= "   group by rmcode  ) yy"
        StrSQL &= "  on yy.rmcode = rm.rmcode "

        StrSQL &= "  where (ActPrice+QQty <> 0.00 or ActPrice <> 0.00) "
        StrSQL &= "  order by rm.rmcode "
        If Not DT Is Nothing Then
            If DT.Rows.Count >= 1 Then
                DT.Clear()
            End If
        End If

        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CB As New SqlCommandBuilder(DA)
            tb1 = New DataTable
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
        UpdateDataGrid()
        '************************************
        'Dim i As Integer
        'Dim c34 As String = Chr(34)
        'For i = 0 To DT.Columns.Count - 1
        '    Dim col As String = DT.Columns(i).ColumnName
        '    Dim coltype As String = DT.Columns(i).DataType.FullName
        '    coltype = coltype.Replace("System.", "")
        '    coltype = coltype.Replace("Int32", "integer")
        '    coltype = coltype.Replace("Int16", "integer")
        '    coltype = coltype.Replace("String", "string")
        '    coltype = coltype.Replace("Decimal", "decimal")
        '    Debug.WriteLine("<xs:element name=" & c34 & col.Trim & c34 & "  type= " & c34 & "xs:" & coltype & c34 & " minOccurs=" & c34 & "0" & c34 & "/>")
        'Next
    End Sub

#End Region

#Region "Form Event"
    Private Sub FrmReportScrap_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        IsLoad = True
    End Sub

    Private Sub FrmReportScrap_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        If IsLoad Then
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            IsLoad = False
            CmdPrint.Enabled = False
            DGView.Visible = False
            pbLoading.Visible = True

            'Set center
            Dim x As Integer = (DGView.Width \ 2) - (pbLoading.Width \ 2)
            Dim y As Integer = (DGView.Height \ 2) - (pbLoading.Height \ 2)
            pbLoading.Location = New Point(x, y)

            BackgroundWorker1.RunWorkerAsync()
        End If
    End Sub
#End Region

#Region "Control Event"
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        'Report
        LoadData()
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        DGView.DataSource = GrdDV
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Dim itotal, istotal, iatotal As Double

        For i As Integer = 0 To GrdDV.Count - 1
            itotal = itotal + GrdDV.Item(i).Row("Total")
            istotal = istotal + GrdDV.Item(i).Row("SAMOUNT")
            iatotal = iatotal + GrdDV.Item(i).Row("AAMOUNT")
        Next

        LblTotal.Text = Format(CDbl(itotal), "###,###,###,###,##0.00")
        lblstd.Text = Format(CDbl(istotal), "###,###,###,###,##0.00")
        lblatc.Text = Format(CDbl(iatotal), "###,###,###,###,##0.00")

        pbLoading.Visible = False
        DGView.Visible = True
        CmdPrint.Enabled = True
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub UpdateDataGrid()
        BackgroundWorker1.ReportProgress(100)
    End Sub

    Private Sub CmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdPrint.Click
        Dim i As Integer
        Dim fRpt As New FrmViewScrap
        Dim aDr() As DataRow = GrdDV.Table.Select(GrdDV.RowFilter)
        Dim dr As DataRow
        Dim tbNew As DataTable
        tbNew = New DataTable
        tbNew = DT.Clone
        For Each dr In aDr
            Dim drNew As DataRow
            drNew = tbNew.NewRow
            For i = 0 To GrdDV.Table.Columns.Count - 1
                drNew(i) = dr(i)
            Next
            tbNew.Rows.Add(drNew)
        Next
        tbNew.AcceptChanges()
        Dim dt4prt As DataTable
        dt4prt = New DataTable
        dt4prt = tbNew

        fRpt.dt_new = GrdDV.Table
        fRpt.sUser = Username
        fRpt.sAdj = txtAdj.Text.Trim()
        fRpt.sCODE = sCODE.Trim()
        fRpt.sName = sName.Trim()
        fRpt.sSec = sSec.Trim()
        fRpt.ShowDialog()
    End Sub
#End Region
End Class
