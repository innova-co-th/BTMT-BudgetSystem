#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports System.Globalization
Imports Inventory_Tag.Common
Imports Inventory_Tag.FrmInvTag
#End Region

Public Class FrmAdd
    Inherits System.Windows.Forms.Form
   

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
    Friend WithEvents cmdShow As System.Windows.Forms.Button
    Friend WithEvents GrdItem As System.Windows.Forms.DataGrid
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmbType As System.Windows.Forms.ComboBox
    Friend WithEvents cmbCode As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmbLoc As System.Windows.Forms.ComboBox
    Friend WithEvents DateTime1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents sbar As System.Windows.Forms.StatusBar
    Friend WithEvents MsgPanel As System.Windows.Forms.StatusBarPanel
    Friend WithEvents CurrentUserPanel As System.Windows.Forms.StatusBarPanel
    Friend WithEvents DateTimePanel As System.Windows.Forms.StatusBarPanel
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtTagNo As System.Windows.Forms.TextBox
    Friend WithEvents RB1 As System.Windows.Forms.RadioButton
    Friend WithEvents RB2 As System.Windows.Forms.RadioButton
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtRemark As System.Windows.Forms.TextBox
    Friend WithEvents txtcode As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RB3 As System.Windows.Forms.RadioButton
    Friend WithEvents DPeriod As System.Windows.Forms.DateTimePicker
    Friend WithEvents yPeriod As System.Windows.Forms.DateTimePicker
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmAdd))
        Me.cmdShow = New System.Windows.Forms.Button
        Me.cmbType = New System.Windows.Forms.ComboBox
        Me.GrdItem = New System.Windows.Forms.DataGrid
        Me.cmdOK = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmbCode = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.cmbLoc = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.DateTime1 = New System.Windows.Forms.DateTimePicker
        Me.sbar = New System.Windows.Forms.StatusBar
        Me.MsgPanel = New System.Windows.Forms.StatusBarPanel
        Me.CurrentUserPanel = New System.Windows.Forms.StatusBarPanel
        Me.DateTimePanel = New System.Windows.Forms.StatusBarPanel
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtTagNo = New System.Windows.Forms.TextBox
        Me.RB1 = New System.Windows.Forms.RadioButton
        Me.RB2 = New System.Windows.Forms.RadioButton
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtRemark = New System.Windows.Forms.TextBox
        Me.txtcode = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.RB3 = New System.Windows.Forms.RadioButton
        Me.DPeriod = New System.Windows.Forms.DateTimePicker
        Me.yPeriod = New System.Windows.Forms.DateTimePicker
        CType(Me.GrdItem, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MsgPanel, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CurrentUserPanel, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DateTimePanel, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdShow
        '
        Me.cmdShow.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdShow.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.cmdShow.Location = New System.Drawing.Point(231, 144)
        Me.cmdShow.Name = "cmdShow"
        Me.cmdShow.Size = New System.Drawing.Size(184, 40)
        Me.cmdShow.TabIndex = 4
        Me.cmdShow.Text = "แสดงรายการ (F5)"
        Me.cmdShow.Visible = False
        '
        'cmbType
        '
        Me.cmbType.AccessibleDescription = ""
        Me.cmbType.AccessibleName = ""
        Me.cmbType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbType.ItemHeight = 13
        Me.cmbType.Location = New System.Drawing.Point(168, 40)
        Me.cmbType.Name = "cmbType"
        Me.cmbType.Size = New System.Drawing.Size(184, 21)
        Me.cmbType.TabIndex = 1
        '
        'GrdItem
        '
        Me.GrdItem.CaptionVisible = False
        Me.GrdItem.DataMember = ""
        Me.GrdItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GrdItem.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.GrdItem.Location = New System.Drawing.Point(3, 16)
        Me.GrdItem.Name = "GrdItem"
        Me.GrdItem.Size = New System.Drawing.Size(666, 117)
        Me.GrdItem.TabIndex = 0
        '
        'cmdOK
        '
        Me.cmdOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOK.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdOK.Location = New System.Drawing.Point(544, 254)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(64, 24)
        Me.cmdOK.TabIndex = 5
        Me.cmdOK.Text = "Save"
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdCancel.Location = New System.Drawing.Point(616, 254)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(64, 24)
        Me.cmdCancel.TabIndex = 7
        Me.cmdCancel.Text = "Close"
        '
        'cmbCode
        '
        Me.cmbCode.AccessibleDescription = ""
        Me.cmbCode.AccessibleName = ""
        Me.cmbCode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCode.ItemHeight = 13
        Me.cmbCode.Location = New System.Drawing.Point(400, 40)
        Me.cmbCode.Name = "cmbCode"
        Me.cmbCode.Size = New System.Drawing.Size(120, 21)
        Me.cmbCode.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(360, 43)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 14)
        Me.Label1.TabIndex = 51
        Me.Label1.Text = "Code"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(136, 43)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(32, 14)
        Me.Label2.TabIndex = 49
        Me.Label2.Text = "Type"
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label3.Location = New System.Drawing.Point(8, 11)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 14)
        Me.Label3.TabIndex = 52
        Me.Label3.Text = "Location"
        '
        'cmbLoc
        '
        Me.cmbLoc.AccessibleDescription = ""
        Me.cmbLoc.AccessibleName = ""
        Me.cmbLoc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmbLoc.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbLoc.ItemHeight = 13
        Me.cmbLoc.Location = New System.Drawing.Point(64, 8)
        Me.cmbLoc.Name = "cmbLoc"
        Me.cmbLoc.Size = New System.Drawing.Size(216, 21)
        Me.cmbLoc.TabIndex = 14
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.Location = New System.Drawing.Point(528, 10)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(40, 16)
        Me.Label4.TabIndex = 55
        Me.Label4.Text = "Date"
        '
        'DateTime1
        '
        Me.DateTime1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DateTime1.CustomFormat = "dd/MM/yyyy"
        Me.DateTime1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTime1.Location = New System.Drawing.Point(568, 8)
        Me.DateTime1.Name = "DateTime1"
        Me.DateTime1.Size = New System.Drawing.Size(96, 20)
        Me.DateTime1.TabIndex = 10
        '
        'sbar
        '
        Me.sbar.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.sbar.Dock = System.Windows.Forms.DockStyle.None
        Me.sbar.Location = New System.Drawing.Point(0, 288)
        Me.sbar.Name = "sbar"
        Me.sbar.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.MsgPanel, Me.CurrentUserPanel, Me.DateTimePanel})
        Me.sbar.ShowPanels = True
        Me.sbar.Size = New System.Drawing.Size(712, 22)
        Me.sbar.TabIndex = 56
        '
        'MsgPanel
        '
        Me.MsgPanel.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
        Me.MsgPanel.Icon = CType(resources.GetObject("MsgPanel.Icon"), System.Drawing.Icon)
        Me.MsgPanel.Width = 366
        '
        'CurrentUserPanel
        '
        Me.CurrentUserPanel.Alignment = System.Windows.Forms.HorizontalAlignment.Center
        Me.CurrentUserPanel.BorderStyle = System.Windows.Forms.StatusBarPanelBorderStyle.Raised
        Me.CurrentUserPanel.Width = 120
        '
        'DateTimePanel
        '
        Me.DateTimePanel.BorderStyle = System.Windows.Forms.StatusBarPanelBorderStyle.Raised
        Me.DateTimePanel.Width = 210
        '
        'CheckBox1
        '
        Me.CheckBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CheckBox1.Location = New System.Drawing.Point(448, 256)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(72, 24)
        Me.CheckBox1.TabIndex = 8
        Me.CheckBox1.Text = "Continue"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 42)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(48, 16)
        Me.Label5.TabIndex = 58
        Me.Label5.Text = "Tag No."
        '
        'TxtTagNo
        '
        Me.TxtTagNo.Location = New System.Drawing.Point(64, 40)
        Me.TxtTagNo.Name = "TxtTagNo"
        Me.TxtTagNo.Size = New System.Drawing.Size(64, 20)
        Me.TxtTagNo.TabIndex = 0
        Me.TxtTagNo.Text = ""
        '
        'RB1
        '
        Me.RB1.Checked = True
        Me.RB1.Location = New System.Drawing.Point(528, 34)
        Me.RB1.Name = "RB1"
        Me.RB1.Size = New System.Drawing.Size(80, 16)
        Me.RB1.TabIndex = 9
        Me.RB1.TabStop = True
        Me.RB1.Text = "1 st Year"
        '
        'RB2
        '
        Me.RB2.Location = New System.Drawing.Point(528, 56)
        Me.RB2.Name = "RB2"
        Me.RB2.Size = New System.Drawing.Size(80, 16)
        Me.RB2.TabIndex = 10
        Me.RB2.Text = "2 nd Year"
        '
        'Label6
        '
        Me.Label6.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label6.Location = New System.Drawing.Point(8, 240)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(56, 14)
        Me.Label6.TabIndex = 62
        Me.Label6.Text = "Remark"
        '
        'TxtRemark
        '
        Me.TxtRemark.Location = New System.Drawing.Point(64, 240)
        Me.TxtRemark.Multiline = True
        Me.TxtRemark.Name = "TxtRemark"
        Me.TxtRemark.Size = New System.Drawing.Size(376, 40)
        Me.TxtRemark.TabIndex = 6
        Me.TxtRemark.Text = ""
        '
        'txtcode
        '
        Me.txtcode.Location = New System.Drawing.Point(400, 72)
        Me.txtcode.Name = "txtcode"
        Me.txtcode.Size = New System.Drawing.Size(120, 20)
        Me.txtcode.TabIndex = 2
        Me.txtcode.Text = ""
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(360, 72)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(40, 14)
        Me.Label7.TabIndex = 63
        Me.Label7.Text = "Code"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GrdItem)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 96)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(672, 136)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        '
        'RB3
        '
        Me.RB3.Location = New System.Drawing.Point(528, 80)
        Me.RB3.Name = "RB3"
        Me.RB3.Size = New System.Drawing.Size(88, 16)
        Me.RB3.TabIndex = 11
        Me.RB3.Text = "Month/Year"
        '
        'DPeriod
        '
        Me.DPeriod.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DPeriod.CustomFormat = "MM/yyyy"
        Me.DPeriod.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DPeriod.Location = New System.Drawing.Point(608, 78)
        Me.DPeriod.Name = "DPeriod"
        Me.DPeriod.ShowUpDown = True
        Me.DPeriod.Size = New System.Drawing.Size(72, 20)
        Me.DPeriod.TabIndex = 13
        '
        'yPeriod
        '
        Me.yPeriod.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.yPeriod.CustomFormat = "yyyy"
        Me.yPeriod.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.yPeriod.Location = New System.Drawing.Point(608, 32)
        Me.yPeriod.Name = "yPeriod"
        Me.yPeriod.ShowUpDown = True
        Me.yPeriod.Size = New System.Drawing.Size(72, 20)
        Me.yPeriod.TabIndex = 12
        '
        'FrmAdd
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(694, 316)
        Me.Controls.Add(Me.yPeriod)
        Me.Controls.Add(Me.DPeriod)
        Me.Controls.Add(Me.RB3)
        Me.Controls.Add(Me.txtcode)
        Me.Controls.Add(Me.TxtRemark)
        Me.Controls.Add(Me.TxtTagNo)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.RB1)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.RB2)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.sbar)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.DateTime1)
        Me.Controls.Add(Me.cmbLoc)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cmdShow)
        Me.Controls.Add(Me.cmbType)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmbCode)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "FrmAdd"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ADD Inventory Tag"
        CType(Me.GrdItem, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MsgPanel, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CurrentUserPanel, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DateTimePanel, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Public Shared GrdItemDv As New DataView
    Protected Const TBL_ITEM As String = "TBL_Item"
    Dim bFormLoad As Boolean
    Dim oldCol As Long
    Dim oldRow As Long
    Dim c1 As New SQLData("ACCINV")
    Public Shared aRow As Long
    Public Shared cm As CurrencyManager
    Dim TrxNo As String
    Dim dd(), idate As String

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub FrmAdd_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dd = Split(DateTime1.Text.Trim, "/")
        idate = dd(2)
        PopulateTypeComboType()
        PopulateLocation()
        sbar.Panels(0).Text = Date.Now
        sbar.Panels(2).Text = "Usage Name : " + CurrentName.Trim
    End Sub

    Sub trxRecord()
        Dim i As Integer
        Dim istr As String
        dd = Split(DateTime1.Text.Trim, "/")
        idate = dd(2)
        If cmdOK.Text = "Save" Then
            i = iNo(idate) + 1
            istr = Format(i, "00000")
            TrxNo = istr
        Else
        End If
        sbar.Panels(1).Text = TrxNo
    End Sub

    Private Function iNo(ByVal idate As String) As Long
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim drSQL As SqlDataReader
        Dim strSQL As String
        Try

            ' Build Insert statement to insert 
            strSQL = "  SELECT   top 1 TrxNo "
            strSQL &= "  FROM   TblTrx"
            strSQL &= "  where TrxYear = '" & idate & "'"
            strSQL &= "  order by TrxNo desc"

            cnSQL = New SqlConnection(c1.Strcon)
            cnSQL.Open()
            cmSQL = New SqlCommand(strSQL, cnSQL)
            cmSQL.ExecuteScalar()
            drSQL = cmSQL.ExecuteReader()
            If drSQL.HasRows Then
                If drSQL.Read() Then
                    iNo = CLng(drSQL.Item("TrxNo").ToString())
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

#Region "Populate Combo"

    Private Sub PopulateTypeComboType()
        Dim objListType As ListType
        Dim dr As DataRow
        cmbType.Items.Clear()
        Dim i As Integer = 0
        For Each dr In Vwtype.Table.Select(Vwtype.RowFilter)
            objListType = New ListType(RTrim(dr.Item("Typecode").ToString()) + " - " + RTrim(dr.Item("TypeName").ToString()), i)
            cmbType.Items.Add(objListType)
            i += 1
        Next
        cmbType.SelectedIndex = 0
    End Sub

    Private Sub PopulateTypeComboCode()
        VwCode.RowFilter = "Typecode='" & Mid(cmbType.Text, 1, 2) & "'"
        Dim objListType As ListType
        Dim dr As DataRow
        cmbCode.Items.Clear()
        Dim i As Integer = 0
        For Each dr In VwCode.Table.Select(VwCode.RowFilter)
            objListType = New ListType(RTrim(dr.Item("code").ToString()), i)
            cmbCode.Items.Add(objListType)
            i += 1
        Next

        'Modify error when cmbCode have no any item (By Beam 31-Aug-2020)------
        'cmbCode.SelectedIndex = 0
        If cmbCode.Items.Count > 0 Then
            cmbCode.SelectedIndex = 0
        End If
        '----------------------------------------------------------------------
    End Sub
    Private Sub PopulateLocation()
        Dim objListType As ListType
        Dim dr As DataRow
        cmbLoc.Items.Clear()
        Dim i As Integer = 0
        For Each dr In VwLoc.Table.Select(VwLoc.RowFilter)
            objListType = New ListType(RTrim(dr.Item("DeptCode").ToString()) + " - " + RTrim(dr.Item("DeptName").ToString()), i)
            cmbLoc.Items.Add(objListType)
            i += 1
        Next
        cmbLoc.SelectedIndex = 0
    End Sub
#End Region

#Region "Combo Change"

    Private Sub cmbType_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbType.Leave
        If cmbCode.Text.Trim = "" Then
            PopulateTypeComboCode()
        End If
    End Sub
    Private Sub cmbType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbType.SelectedIndexChanged
        PopulateTypeComboCode()

        'Modify error when cmbCode have no any item (By Beam 31-Aug-2020)------
        If cmbCode.Items.Count <= 0 Then
            switch()
            ShowData()
        End If
        '----------------------------------------------------------------------
    End Sub

    Private Sub cmbCode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCode.SelectedIndexChanged
        switch()
        ShowData()
    End Sub

    Sub switch()
        Try
            GrdItem.CurrentCell = New DataGridCell(0, 0)
        Catch
        End Try
        GrdItem.Visible = False
        cmdShow.Visible = True
    End Sub

#End Region

#Region "Buttom Show Grid"

    'Private Sub cmdShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdShow.Click
    '    ShowData()
    'End Sub
    Private Sub ShowData()
        If cmbType.Text.Trim = "" Or cmbCode.Text.Trim = "" Then
            MsgBox("เลือกเงื่อนไขก่อน")
            Exit Sub
        End If
        cmdShow.Visible = False
        GrdItem.Visible = True
        setGrid()
        GrdItemDv.RowFilter = " Typecode = '" & Mid(cmbType.Text.Trim, 1, 2) & "' and code = '" & cmbCode.Text.Trim & "'"
        GrdItem.DataSource = GrdItemDv
        oldRow = 0
        oldCol = 0
        GrdItem.Focus()
    End Sub

#End Region

#Region "Set grid "
    Private Sub setGrid()
        Dim dt As DataTable = New DataTable()
        Dim StrSQL As String
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = " select  distinct Typecode,code,QTY,Unit,shortUnitName,UnitName from ("
        StrSQL &= "  SELECT    cc.Typecode,cc.code,isnull(Revision,'')Rev,0.00 QTY,UnitBig Unit"
        StrSQL &= "  FROM         "
        StrSQL &= " (SELECT distinct TYpecode, MasterCode, Revision "
        StrSQL &= " FROM         TBLMASTER m"
        StrSQL &= " Right outer join "
        StrSQL &= " TBLGroup  g"
        StrSQL &= " on m.Mastercode = g.code )gg"
        StrSQL &= "   right outer join "
        StrSQL &= " ("
        StrSQL &= " select distinct Type Typecode,c.code,UnitBig from TBLConvert c"
        StrSQL &= " left outer join  TBLGroup g"
        StrSQL &= " on c.code = g.code"
        StrSQL &= " ) cc"
        StrSQL &= " on gg.MasterCode = cc.code"
        StrSQL &= "  Union"
        StrSQL &= " SELECT  distinct Typecode,code,isnull(Revision,'') Rev,0.00 Qty,'KG' Unit"
        StrSQL &= " FROM         TBLMASTER m"
        StrSQL &= " Right outer join "
        StrSQL &= " TBLGroup  g"
        StrSQL &= " on m.Mastercode = g.code"
        StrSQL &= " )Item"
        StrSQL &= " left outer join"
        StrSQL &= "  TblUnit "
        StrSQL &= "  on Item.unit = TblUnit.UnitCode"
        StrSQL &= " order by code,ShortUnitName"
        If Not dt Is Nothing Then
            If dt.Rows.Count >= 1 Then
                dt.Clear()
            End If
        End If

        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, c1.Strcon)
            Dim CB As New SqlCommandBuilder(DA)
            dt = New DataTable
            DA.Fill(dt)
        Catch
            MsgBox("Can't Select Data.", MsgBoxStyle.Critical, "Load Data")
        Finally
        End Try
        '************************************
        dt.TableName = TBL_ITEM
        GrdItemDv = dt.DefaultView
        GrdItemDv.AllowNew = False
        GrdItemDv.AllowDelete = False
        '************************************
        GrdItem.DataSource = GrdItemDv
        '************************************
        ResetTableStyleBf()

        With GrdItem
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
            .HeaderBackColor = Color.Orchid
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
            .MappingName = TBL_ITEM
            .PreferredColumnWidth = 125
            .PreferredRowHeight = 15
        End With
        Dim grdColStyle1 As New DataGridColoredLine2
        With grdColStyle1
            .HeaderText = "Material"
            .MappingName = "Code"
            .Width = 180
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Center
        End With

        cm = CType(Me.BindingContext(GrdItem.DataSource, GrdItem.DataMember), CurrencyManager)
        Dim c As CheckRowHeader
        c = AddressOf CheckRowHeader

        Dim grdColStyle6 As New DataGridQtyBox(c)
        With grdColStyle6
            .HeaderText = "Qty"
            .MappingName = "Qty"
            .Format = "###,###.000"
            .Width = 80
            .Alignment = HorizontalAlignment.Center
            .NullText = ""
        End With

        Dim grdColStyle5 As New DataGridColoredLine2
        With grdColStyle5
            .HeaderText = "Unit"
            .MappingName = "ShortUnitName"
            .Width = 120
            .NullText = ""
            .ReadOnly = True
            .Alignment = HorizontalAlignment.Left
        End With

        grdTableStyle1.GridColumnStyles.AddRange _
(New DataGridColumnStyle() _
{grdColStyle1, grdColStyle6, grdColStyle5})

        GrdItem.TableStyles.Add(grdTableStyle1)
        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub

    Private Sub ResetTableStyleBf()
        ' Clear out the existing TableStyles and result default formatting.
        With GrdItem
            .BackgroundColor = SystemColors.InactiveCaptionText
            '.CaptionText = ""
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
            '.BorderStyle = DefaultGridBorderStyle
        End With
    End Sub
#End Region

#Region "Delegate"

    Public Shared Function CheckRowHeader(ByVal row As Integer) As Boolean
        Dim c As Boolean = False
        'Debug.WriteLine("st seq : " + CStr(GrdDV.Item(row).Item("st_seq")) + "   row : " + CStr(row))
        If GrdItemDv.Item(row).Item("Code").ToString.Trim = "" Then
            c = True
        Else
            c = False
        End If
        Return c
    End Function

#End Region

#Region "Key Enter"

    Private Sub KeyEnterToNext(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbType.KeyPress
        Select Case Asc(e.KeyChar)
            Case 13
                e.Handled = True
                txtcode.Focus()
                '    SendKeys.Send("{TAB}")
            Case 43 '+
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case 45 '-
                e.Handled = True
                SendKeys.Send("+{TAB}")
            Case 32 ' space bar
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case Else
        End Select
    End Sub
    Private Sub KeyEnterToNext2(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbCode.KeyPress
        Select Case Asc(e.KeyChar)
            Case 13
                ShowData()
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case 43 '+
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case 45 '-
                e.Handled = True
                SendKeys.Send("+{TAB}")
            Case 32 ' space bar
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case Else
        End Select
    End Sub
    Private Sub txtcode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtcode.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                cmbCode.Text = txtcode.Text
                txtcode.Text = txtcode.Text.ToUpper

                If txtcode.Text.Trim <> "" Then
                    If cmbCode.Text.Trim <> txtcode.Text.Trim Then
                        MsgBox("Please check data again. ", MsgBoxStyle.OKOnly, "TAG")
                        Exit Sub
                    End If
                End If

                SendKeys.Send("{TAB}")
            Case 32 ' space bar
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case Else
        End Select
    End Sub
    Private Sub TxtTagNo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtTagNo.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case 32 ' space bar
                e.Handled = True
                SendKeys.Send("{TAB}")
            Case Else
        End Select
    End Sub

#End Region

#Region "Grid Item Event"
    Private Sub GrdItem_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GrdItem.CurrentCellChanged
        'Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        Dim xRow As Long = GrdItem.CurrentCell.RowNumber
        Dim xCol As Long = GrdItem.CurrentCell.ColumnNumber
        If xRow = aRow And oldRow = aRow - 1 Then
            ''cbCount.Focus()
            cmdOK.Focus()
            Exit Sub
        Else
            If xRow = aRow - 1 And xCol > oldCol And xRow <> 0 Then
                cmdOK.Focus()
                'oldCol = 2
                Exit Sub
            End If
        End If
    End Sub

    Private Sub GrdItem_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles GrdItem.MouseDown
        If e.Button = MouseButtons.Left And GrdItem.CurrentRowIndex >= 0 Then
            'clear_select(GrdItem)
            'GrdItem.CurrentCell = CellZero 'New DataGridCell(0, 0)
            Dim hit As System.Windows.Forms.DataGrid.HitTestInfo
            hit = GrdItem.HitTest(e.X, e.Y)
            If hit.Row >= aRow And hit.Row <> aRow - 1 Then
                'GrdItem.Select(hit.Row)
                GrdItem.CurrentRowIndex = hit.Row
                If GrdItem.CurrentCell.ColumnNumber = hit.Column Then
                    If hit.Column = 2 Then
                        GrdItem.CurrentCell = New DataGridCell(hit.Row, 2)
                    Else
                        GrdItem.CurrentCell = New DataGridCell(hit.Row, 2)
                    End If
                Else
                    GrdItem.CurrentCell = New DataGridCell(hit.Row, 2)
                End If
            Else
                If hit.Row = aRow - 1 Then
                    GrdItem.CurrentCell = New DataGridCell(aRow - 1, 2)
                    oldRow = aRow - 1
                    oldCol = 2
                    sender = e.Empty
                End If
            End If
        Else
            Exit Sub
        End If
    End Sub

    Private Sub GrdItem_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles GrdItem.Enter
        Try
            If oldRow = aRow - 1 And oldRow >= 1 Then
                oldRow -= 1
            End If
            GrdItem.CurrentCell = New DataGridCell(oldRow, 2)
        Catch
        End Try
        oldRow = GrdItem.CurrentCell.RowNumber
        oldCol = GrdItem.CurrentCell.ColumnNumber
    End Sub


#End Region

#Region "Ok add line "

    'Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    '    Dim Row As DataRow
    '    Dim dv As New DataView
    '    dv.Table = tb2
    '    dv.RowFilter = "qty<>0"
    '    For Each Row In dv.Table.Select(dv.RowFilter)
    '        Dim a_row As DataRow
    '        a_row = tb2.NewRow
    '        a_row.Item("seq") = Row.Item("seq")
    '        a_row.Item("LeftItem") = Row.Item("LeftItem")
    '        a_row.Item("Code") = Row.Item("Code")
    '        a_row.Item("loc") = Row.Item("loc")
    '        a_row.Item("desc1") = Row.Item("desc1")
    '        a_row.Item("qty") = Row.Item("qty")
    '        a_row.Item("uom") = Row.Item("uom")
    '        a_row.Item("descrip") = Row.Item("descrip")
    '        tb2.Rows.Add(a_row)
    '    Next
    '    tb2.AcceptChanges()
    '    Me.Hide()
    'End Sub

#End Region

#Region "Form Keydown F5 & Esc"

    Private Sub frmAdd_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown, cmbCode.KeyDown, cmbType.KeyDown

        If e.KeyCode = Keys.F5 Then
            ShowData()
        End If
        If e.KeyValue = 27 Then
            Me.Hide()
        End If
    End Sub

#End Region

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        Dim tno() As String
        Dim k As Integer
        If txtcode.Text.Trim <> "" Then
            If cmbCode.Text.Trim <> txtcode.Text.Trim Then
                MsgBox("Please check data again. ", MsgBoxStyle.OKOnly, "TAG")
                Exit Sub
            End If
        End If

        If TxtTagNo.Text = "" Then
            TxtTagNo.Focus()
            MsgBox("Please Check Data Again.", MsgBoxStyle.OKCancel)
            Exit Sub
        Else
            tno = Split(TxtTagNo.Text.Trim, "-")
            If tno.Length = 1 Then
                k = TxtTagNo.Text.Trim
                TxtTagNo.Text = Format(k, "0000")
            ElseIf tno.Length = 2 Then
                k = tno(0)
                TxtTagNo.Text = Format(k, "0000") + "-" + tno(1)
            End If
        End If
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult

        msg = "Inventory Record TrxNo : " & TxtTagNo.Text.Trim  ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Inventory"   ' Define title.
        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            TRX()
        Else
            Exit Sub
        End If
    End Sub

#Region "TRX Function"
    Private Function ChkData(ByVal PGCode As String, ByVal RMCode As String, ByVal iCode As String) As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Try
            'strSQL = " Select count(*) from TBLMaster "
            'strSQL += " Where MasterCode = " & PrepareStr(PGCode)
            'strSQL += " and  RMCode = " & PrepareStr(RMCode)
            'strSQL += " and  Revision = " & PrepareStr(iCode)
            cnSQL = New SqlConnection(c1.Strcon)
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
            Me.Close()
        Catch Exp As SqlException
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "SQL Error")
        Catch Exp As Exception
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "General Error")
        End Try
    End Function

    Sub TRX()
        Dim strsql As String = String.Empty
        Dim cn As New SqlConnection(c1.Strcon)
        Dim cmd As New SqlCommand(strsql, cn)
        cn.Open()
        Dim t1 As SqlTransaction = cn.BeginTransaction
        cmd.Transaction = t1
        'Dim strDate(), strDate2(), strDate3() As String    'Comment Beam 28-Aug-2020
        Dim strDate2(), strDate3() As String                'Add Beam 28-Aug-2020
        Dim strTime As String

        '//Modify date format Beam 28-Aug-2020 --------------------
        'strDate = Split(Date.Now.ToShortDateString, "/")
        'strTime = Date.Now.ToShortTimeString

        Dim strDateEN As Date
        strDateEN = Date.Now.ToShortDateString.ToString(New CultureInfo("en-US"))
        Dim strTrxDate As String = String.Empty
        strTrxDate = strDateEN.Year.ToString("D4") & strDateEN.Month.ToString("D2") & strDateEN.Day.ToString("D2")

        strTime = Date.Now.ToString("HH:mm:ss")
        '//End Modify date format Beam 28-Aug-2020 ----------------

        strDate2 = Split(DateTime1.Text.Trim, "/")
        Dim Period As String
        Dim Ytrx As Integer
        If RB1.Checked = True Then
            Ytrx = "01"
            Period = "YL"
            strDate3 = Split(yPeriod.Value.Month & "/" & yPeriod.Value.Year, "/")
        ElseIf RB2.Checked = True Then
            Ytrx = "02"
            Period = "YL"
            strDate3 = Split(yPeriod.Value.Month & "/" & yPeriod.Value.Year, "/")
        Else : RB3.Checked = True
            Ytrx = Now.Date.Month
            Period = "ML"
            strDate3 = Split(DPeriod.Text.Trim, "/")
        End If

        Try
            Dim aDr() As DataRow
            GrdItemDv.RowFilter = " Qty <> 0.00"
            aDr = GrdItemDv.Table.Select(GrdItemDv.RowFilter)
            If UBound(aDr) < 0 Then
                Exit Sub
            End If
            Dim dr As DataRow
            For Each dr In aDr
                With dr
                    If IIf(.Item("Code") Is System.DBNull.Value, "", .Item("Code")) <> "" Then
                        strsql = "Insert TBLTRX"
                        strsql += " Values("
                        strsql += PrepareStr(TxtTagNo.Text.Trim)
                        strsql += "," & PrepareStr(.Item("Code"))
                        strsql += "," & PrepareStr(Period)
                        strsql += "," & PrepareStr(strDate3(1) + Format(Ytrx, "#00"))
                        strsql += "," & PrepareStr(Mid(cmbType.Text.Trim, 1, 2))
                        strsql += "," & PrepareStr(Mid(cmbLoc.Text.Trim, 1, 4))
                        strsql += "," & PrepareStr(.Item("Qty"))
                        strsql += "," & PrepareStr(.Item("Unit"))

                        'strsql += "," & PrepareStr(strDate(2) + strDate(1) + strDate(0))   'Comment by Beam 28-Aug-2020
                        strsql += "," & PrepareStr(strTrxDate)                              'Change to this by Beam 28-Aug-2020

                        strsql += "," & PrepareStr(strTime)
                        strsql += "," & PrepareStr(CurrentIDUser.Trim)
                        strsql += "," & PrepareStr(TxtRemark.Text.Trim)
                        strsql += "," & PrepareStr(strDate2(2) + strDate2(1) + strDate2(0))
                        strsql += "," & PrepareStr(strTime)
                        strsql += ")"
                        cmd.CommandText = strsql
                        cmd.ExecuteNonQuery()
                        MsgBox("Update Complete.", MsgBoxStyle.Information, "Inventory Record")
                        If CheckBox1.Checked = True Then
                            switch()
                        Else
                            Me.Close()
                        End If
                    End If
                End With
            Next
            t1.Commit()
            Dim tno(), n As String
            Dim k As Integer
            tno = Split(TxtTagNo.Text.Trim, "-")
            If tno.Length = 1 Then
                TxtTagNo.Text = TxtTagNo.Text.Trim + 1
            ElseIf tno.Length = 2 Then
                k = tno(0)
                n = tno(1) + 1
                TxtTagNo.Text = Format(k, "0000") + "-" + n
            End If
        Catch
            t1.Rollback()
            MsgBox("It's Duplicate.Check Data Again.")
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

    Private Sub cmbCode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbCode.Click
        txtcode.Text = ""
    End Sub

    Private Sub cmdShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdShow.Click
        switch()
        ShowData()
    End Sub

    Private Sub GrdItem_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles GrdItem.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                cmdOK.Focus()
                'Case 32 ' space bar
                ' e.Handled = True
            Case Else
        End Select
    End Sub
End Class
