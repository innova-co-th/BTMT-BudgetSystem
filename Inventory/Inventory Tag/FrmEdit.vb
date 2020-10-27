#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Inventory_Tag.Common
Imports Inventory_Tag.FrmInvTag
Imports Inventory_Record.FrmCVT
#End Region

Public Class FrmEdit

#Region "Declare"
    Inherits System.Windows.Forms.Form
    Dim GrdDVType As New DataView
    Protected Const TBL_Type As String = "TBL_Type"
    Dim GrdDVLOC As New DataView
    Protected Const TBL_LOC As String = "TBL_LOC"
    Dim GrdDVRM As New DataView
    Protected Const TBL_RM As String = "TBL_RM"
    Dim GrdDVUOM As New DataView
    Protected Const TBL_UOM As String = "TBL_UOM"

    Public Shared GrdItemDv As New DataView
    Protected Const TBL_ITEM As String = "TBL_Item"
    Dim bFormLoad As Boolean
    Dim oldCol As Long
    Dim oldRow As Long
    Public Shared aRow As Long
    Public Shared cm As CurrencyManager
    Private ReadOnly cult As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

    Friend TType, TLoc, TLocNo, TRMCode, TUnit, Ttime As String
    Dim C1 As New SQLData("ACCINV")
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
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents lblCode As System.Windows.Forms.Label
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents cmbType As System.Windows.Forms.ComboBox
    Friend WithEvents cmbCode As System.Windows.Forms.ComboBox
    Friend WithEvents lblLocation As System.Windows.Forms.Label
    Friend WithEvents lblDate As System.Windows.Forms.Label
    Friend WithEvents cmbLoc As System.Windows.Forms.ComboBox
    Friend WithEvents DateTime1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents sbar As System.Windows.Forms.StatusBar
    Friend WithEvents MsgPanel As System.Windows.Forms.StatusBarPanel
    Friend WithEvents CurrentUserPanel As System.Windows.Forms.StatusBarPanel
    Friend WithEvents DateTimePanel As System.Windows.Forms.StatusBarPanel
    Friend WithEvents lblTagNo As System.Windows.Forms.Label
    Friend WithEvents TxtTagNo As System.Windows.Forms.TextBox
    Friend WithEvents lblQty As System.Windows.Forms.Label
    Friend WithEvents lblUnit As System.Windows.Forms.Label
    Friend WithEvents cmbUnit As System.Windows.Forms.ComboBox
    Friend WithEvents TxtQty As System.Windows.Forms.TextBox
    Friend WithEvents cmdUnit As System.Windows.Forms.Button
    Friend WithEvents Timer1 As Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmEdit))
        Me.cmbType = New System.Windows.Forms.ComboBox()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmbCode = New System.Windows.Forms.ComboBox()
        Me.lblCode = New System.Windows.Forms.Label()
        Me.lblType = New System.Windows.Forms.Label()
        Me.lblLocation = New System.Windows.Forms.Label()
        Me.cmbLoc = New System.Windows.Forms.ComboBox()
        Me.lblDate = New System.Windows.Forms.Label()
        Me.DateTime1 = New System.Windows.Forms.DateTimePicker()
        Me.sbar = New System.Windows.Forms.StatusBar()
        Me.MsgPanel = New System.Windows.Forms.StatusBarPanel()
        Me.CurrentUserPanel = New System.Windows.Forms.StatusBarPanel()
        Me.DateTimePanel = New System.Windows.Forms.StatusBarPanel()
        Me.lblTagNo = New System.Windows.Forms.Label()
        Me.TxtTagNo = New System.Windows.Forms.TextBox()
        Me.TxtQty = New System.Windows.Forms.TextBox()
        Me.lblQty = New System.Windows.Forms.Label()
        Me.lblUnit = New System.Windows.Forms.Label()
        Me.cmbUnit = New System.Windows.Forms.ComboBox()
        Me.cmdUnit = New System.Windows.Forms.Button()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        CType(Me.MsgPanel, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CurrentUserPanel, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DateTimePanel, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmbType
        '
        Me.cmbType.AccessibleDescription = ""
        Me.cmbType.AccessibleName = ""
        Me.cmbType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbType.ItemHeight = 13
        Me.cmbType.Location = New System.Drawing.Point(168, 8)
        Me.cmbType.Name = "cmbType"
        Me.cmbType.Size = New System.Drawing.Size(184, 21)
        Me.cmbType.TabIndex = 0
        '
        'cmdOK
        '
        Me.cmdOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOK.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdOK.Location = New System.Drawing.Point(342, 126)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(80, 24)
        Me.cmdOK.TabIndex = 3
        Me.cmdOK.Text = "Save"
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdCancel.Location = New System.Drawing.Point(430, 126)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(80, 24)
        Me.cmdCancel.TabIndex = 4
        Me.cmdCancel.Text = "Close"
        '
        'cmbCode
        '
        Me.cmbCode.AccessibleDescription = ""
        Me.cmbCode.AccessibleName = ""
        Me.cmbCode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCode.ItemHeight = 13
        Me.cmbCode.Location = New System.Drawing.Point(64, 40)
        Me.cmbCode.Name = "cmbCode"
        Me.cmbCode.Size = New System.Drawing.Size(120, 21)
        Me.cmbCode.TabIndex = 1
        '
        'lblCode
        '
        Me.lblCode.Location = New System.Drawing.Point(8, 43)
        Me.lblCode.Name = "lblCode"
        Me.lblCode.Size = New System.Drawing.Size(40, 14)
        Me.lblCode.TabIndex = 51
        Me.lblCode.Text = "Code"
        '
        'lblType
        '
        Me.lblType.Location = New System.Drawing.Point(120, 11)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(40, 14)
        Me.lblType.TabIndex = 49
        Me.lblType.Text = "Type"
        '
        'lblLocation
        '
        Me.lblLocation.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblLocation.Location = New System.Drawing.Point(8, 137)
        Me.lblLocation.Name = "lblLocation"
        Me.lblLocation.Size = New System.Drawing.Size(56, 14)
        Me.lblLocation.TabIndex = 52
        Me.lblLocation.Text = "Location"
        '
        'cmbLoc
        '
        Me.cmbLoc.AccessibleDescription = ""
        Me.cmbLoc.AccessibleName = ""
        Me.cmbLoc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmbLoc.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbLoc.ItemHeight = 13
        Me.cmbLoc.Location = New System.Drawing.Point(64, 134)
        Me.cmbLoc.Name = "cmbLoc"
        Me.cmbLoc.Size = New System.Drawing.Size(216, 21)
        Me.cmbLoc.TabIndex = 53
        '
        'lblDate
        '
        Me.lblDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblDate.Location = New System.Drawing.Point(368, 10)
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Size = New System.Drawing.Size(40, 16)
        Me.lblDate.TabIndex = 55
        Me.lblDate.Text = "Date"
        '
        'DateTime1
        '
        Me.DateTime1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DateTime1.CustomFormat = "dd/MM/yyyy"
        Me.DateTime1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTime1.Location = New System.Drawing.Point(408, 8)
        Me.DateTime1.Name = "DateTime1"
        Me.DateTime1.Size = New System.Drawing.Size(96, 20)
        Me.DateTime1.TabIndex = 54
        '
        'sbar
        '
        Me.sbar.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.sbar.Dock = System.Windows.Forms.DockStyle.None
        Me.sbar.Location = New System.Drawing.Point(0, 160)
        Me.sbar.Name = "sbar"
        Me.sbar.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.MsgPanel, Me.CurrentUserPanel, Me.DateTimePanel})
        Me.sbar.ShowPanels = True
        Me.sbar.Size = New System.Drawing.Size(536, 22)
        Me.sbar.TabIndex = 56
        '
        'MsgPanel
        '
        Me.MsgPanel.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
        Me.MsgPanel.Icon = CType(resources.GetObject("MsgPanel.Icon"), System.Drawing.Icon)
        Me.MsgPanel.Name = "MsgPanel"
        Me.MsgPanel.Width = 189
        '
        'CurrentUserPanel
        '
        Me.CurrentUserPanel.Alignment = System.Windows.Forms.HorizontalAlignment.Center
        Me.CurrentUserPanel.BorderStyle = System.Windows.Forms.StatusBarPanelBorderStyle.Raised
        Me.CurrentUserPanel.Name = "CurrentUserPanel"
        Me.CurrentUserPanel.Width = 120
        '
        'DateTimePanel
        '
        Me.DateTimePanel.BorderStyle = System.Windows.Forms.StatusBarPanelBorderStyle.Raised
        Me.DateTimePanel.Name = "DateTimePanel"
        Me.DateTimePanel.Width = 210
        '
        'lblTagNo
        '
        Me.lblTagNo.Location = New System.Drawing.Point(8, 10)
        Me.lblTagNo.Name = "lblTagNo"
        Me.lblTagNo.Size = New System.Drawing.Size(48, 16)
        Me.lblTagNo.TabIndex = 58
        Me.lblTagNo.Text = "Tag No."
        '
        'TxtTagNo
        '
        Me.TxtTagNo.Location = New System.Drawing.Point(64, 8)
        Me.TxtTagNo.Name = "TxtTagNo"
        Me.TxtTagNo.ReadOnly = True
        Me.TxtTagNo.Size = New System.Drawing.Size(48, 20)
        Me.TxtTagNo.TabIndex = 59
        '
        'TxtQty
        '
        Me.TxtQty.Location = New System.Drawing.Point(64, 72)
        Me.TxtQty.Name = "TxtQty"
        Me.TxtQty.Size = New System.Drawing.Size(48, 20)
        Me.TxtQty.TabIndex = 61
        '
        'lblQty
        '
        Me.lblQty.Location = New System.Drawing.Point(8, 74)
        Me.lblQty.Name = "lblQty"
        Me.lblQty.Size = New System.Drawing.Size(48, 16)
        Me.lblQty.TabIndex = 60
        Me.lblQty.Text = "Qty"
        '
        'lblUnit
        '
        Me.lblUnit.Location = New System.Drawing.Point(8, 106)
        Me.lblUnit.Name = "lblUnit"
        Me.lblUnit.Size = New System.Drawing.Size(48, 16)
        Me.lblUnit.TabIndex = 62
        Me.lblUnit.Text = "Unit"
        '
        'cmbUnit
        '
        Me.cmbUnit.AccessibleDescription = ""
        Me.cmbUnit.AccessibleName = ""
        Me.cmbUnit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbUnit.ItemHeight = 13
        Me.cmbUnit.Location = New System.Drawing.Point(64, 104)
        Me.cmbUnit.Name = "cmbUnit"
        Me.cmbUnit.Size = New System.Drawing.Size(72, 21)
        Me.cmbUnit.TabIndex = 63
        '
        'cmdUnit
        '
        Me.cmdUnit.Location = New System.Drawing.Point(136, 103)
        Me.cmdUnit.Name = "cmdUnit"
        Me.cmdUnit.Size = New System.Drawing.Size(24, 23)
        Me.cmdUnit.TabIndex = 64
        Me.cmdUnit.Text = "..."
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'FrmEdit
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(518, 188)
        Me.Controls.Add(Me.cmdUnit)
        Me.Controls.Add(Me.cmbUnit)
        Me.Controls.Add(Me.lblUnit)
        Me.Controls.Add(Me.TxtQty)
        Me.Controls.Add(Me.lblQty)
        Me.Controls.Add(Me.TxtTagNo)
        Me.Controls.Add(Me.lblTagNo)
        Me.Controls.Add(Me.sbar)
        Me.Controls.Add(Me.lblDate)
        Me.Controls.Add(Me.DateTime1)
        Me.Controls.Add(Me.cmbLoc)
        Me.Controls.Add(Me.lblLocation)
        Me.Controls.Add(Me.cmbType)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmbCode)
        Me.Controls.Add(Me.lblCode)
        Me.Controls.Add(Me.lblType)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "FrmEdit"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Edit Invtory Tag"
        CType(Me.MsgPanel, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CurrentUserPanel, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DateTimePanel, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

#Region "CONSTANT"
    Dim TrxNo As String
    Dim idate As String
    Dim DT As New DataTable
    Dim StrSQL As String
#End Region

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

#Region "Form Event"
    Private Sub FrmEdit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadType() 'Set Type
        LoadLoc() 'Set Department
        LoadRMCode() 'Load data in table TBLGrouop and TBLCompound
        cmbType.Text = TType.Trim() 'TypeName
        cmbLoc.Text = TLoc.Trim() 'DeptName
        cmbCode.Text = TRMCode.Trim() 'Code
        cmbUnit.Text = TUnit.Trim() 'Unit

        Dim dd As DateTime = Convert.ToDateTime(DateTime1.Value, cult)
        idate = dd.ToString("yyyy") 'Get year

        sbar.Panels(2).Text = "Usage Name : " + CurrentName.Trim() 'Display User
        Timer1.Enabled = True 'Start Time
    End Sub
#End Region

#Region "COMBOBOX"
    Sub LoadType()
        Dim dtType As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = "SELECT  *  FROM  TBLType "
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtType = New DataTable
            DA.Fill(dtType)
        Catch
        Finally
        End Try
        dtType.TableName = TBL_Type
        GrdDVType = dtType.DefaultView
        '************************************
        cmbType.DisplayMember = "TypeName"
        cmbType.ValueMember = "TypeCode"
        cmbType.DataSource = dtType
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadLoc()
        Dim dtLoc As DataTable = New DataTable()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSQL = "SELECT  *  FROM  TBLDepartment "
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtLoc = New DataTable
            DA.Fill(dtLoc)
        Catch
        Finally
        End Try
        dtLoc.TableName = TBL_LOC
        GrdDVLOC = dtLoc.DefaultView
        '************************************
        cmbLoc.DisplayMember = "DeptName"
        cmbLoc.ValueMember = "DeptCode"
        cmbLoc.DataSource = dtLoc
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadRMCode()
        Dim dtRM As DataTable = New DataTable()
        Dim sb As New System.Text.StringBuilder()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sb.AppendLine(" SELECT   *  FROM  TblGroup")
        sb.AppendLine(" UNION")
        sb.AppendLine(" SELECT '03' Typecode,Finalcompound code")
        sb.AppendLine(" FROM TBLCompound")
        sb.AppendLine(" WHERE active = 1")
        sb.AppendLine(" ORDER BY Typecode,code")
        StrSQL = sb.ToString()

        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtRM = New DataTable
            DA.Fill(dtRM)
        Catch
        Finally
        End Try
        dtRM.TableName = TBL_RM
        GrdDVRM = dtRM.DefaultView
        '************************************
        cmbCode.DisplayMember = "Code"
        cmbCode.ValueMember = "Code"
        cmbCode.DataSource = dtRM
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Sub LoadUom()
        Dim dtUom As DataTable = New DataTable()
        Dim sb As New System.Text.StringBuilder()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sb.AppendLine(" SELECT  unitcode,ShortUnitName ")
        sb.AppendLine(" FROM  TBLUnit ")
        sb.AppendLine(" WHERE unitcode IN (SELECT  UnitBig  FROM  TBLConvert  where  code ='" & cmbCode.Text.Trim & "')")
        StrSQL = sb.ToString()

        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtUom = New DataTable
            DA.Fill(dtUom)
        Catch
        Finally
        End Try
        dtUom.TableName = TBL_UOM
        GrdDVUOM = dtUom.DefaultView
        '************************************
        cmbUnit.DisplayMember = "ShortUnitName"
        cmbUnit.ValueMember = "unitcode"
        cmbUnit.DataSource = dtUom
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
#End Region

#Region "Combo Change"



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
    Private Sub KeyEnterToNext(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbCode.KeyPress, cmbType.KeyPress
        Select Case Asc(e.KeyChar)
            Case 13
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
        End Select
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

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        If TxtTagNo.Text = "" Then
            TxtTagNo.Focus()
            MsgBox("Please Check Data Again.", MsgBoxStyle.OkCancel)
            Exit Sub
        Else
            Dim k As Integer
            k = TxtTagNo.Text.Trim()
            TxtTagNo.Text = Format(k, "0000")
        End If

        Dim msg As String = "Inventory Record TrxNo : " & TxtTagNo.Text.Trim  ' Define message.
        Dim title As String = "Inventory"   ' Define title.
        Dim style As MsgBoxStyle = MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        Dim response As MsgBoxResult

        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            TRX()
        Else
            Exit Sub
        End If
    End Sub
#Region "TRX Function"
    Sub TRX()
        Dim strsql As String = String.Empty
        Dim cn As New SqlConnection(C1.Strcon)
        Dim cmd As New SqlCommand(strsql, cn)
        cn.Open()
        Dim t1 As SqlTransaction = cn.BeginTransaction
        cmd.Transaction = t1
        Dim strDate(), strDate1() As String
        Dim strTime As String
        strDate = Split(DateTime1.Text, "/")
        strDate1 = Split(Now.Date.ToShortDateString, "/")
        strTime = Date.Now.ToShortTimeString

        Try
            strsql = "update TBLTRX"
            strsql += " set code =" & PrepareStr(cmbCode.SelectedValue) 'Code
            strsql += ", Qty = " & PrepareStr(TxtQty.Text.Trim) 'Qty
            strsql += ", Typecode = " & PrepareStr(cmbType.SelectedValue) 'TypeCode
            strsql += ", Location = " & PrepareStr(cmbLoc.SelectedValue) 'Location
            strsql += ", UserID = " & PrepareStr(CurrentIDUser.Trim) 'UserID
            strsql += ", Trxdate = " & PrepareStr(strDate(2) + strDate(1) + strDate(0)) 'TrxDate
            strsql += ",TrxTime = " & PrepareStr(strTime) 'TrxTime
            strsql += ",Updatedate = " & PrepareStr(strDate1(2) + strDate1(1) + strDate1(0)) 'UpdateDate
            strsql += ",UpdateTime = " & PrepareStr(strTime) 'UpdateTime
            strsql += ",UOM = " & PrepareStr(cmbUnit.SelectedValue) 'Uom
            strsql += " where tagNo = " & PrepareStr(TxtTagNo.Text.Trim)
            strsql += " and Location = " & PrepareStr(TLocNo.Trim)
            strsql += " and code = " & PrepareStr(TRMCode.Trim)
            cmd.CommandText = strsql
            cmd.ExecuteNonQuery()
            MsgBox("Update Complete.", MsgBoxStyle.Information, "Inventory Record")
            Me.Close()
            t1.Commit()
            TxtTagNo.Text = TxtTagNo.Text + 1
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

    Private Sub TxtTagNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtTagNo.TextChanged

    End Sub
    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQty.TextChanged

    End Sub

    Private Sub cmbCode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCode.SelectedIndexChanged
        LoadUom()
    End Sub

    Private Sub cmdUnit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUnit.Click
        Dim fcvt As New Inventory_Record.FrmAddCvt
        fcvt.Text = "Save"
        fcvt.StrType = cmbType.SelectedValue.ToString()
        fcvt.StrMaterial = cmbCode.SelectedValue.ToString()
        fcvt.ShowDialog()
        LoadType()
        LoadLoc()
        LoadRMCode()
        cmbType.Text = TType.Trim
        cmbLoc.Text = TLoc.Trim
        cmbCode.Text = TRMCode.Trim
        cmbUnit.Text = TUnit.Trim
        Dim dd As DateTime = Convert.ToDateTime(DateTime1.Value, cult)
        idate = dd.ToString("yyyy") 'Get year
        sbar.Panels(0).Text = Date.Now
        sbar.Panels(2).Text = "Usage Name : " + CurrentName.Trim
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        sbar.Panels(0).Text = Date.Now.ToString("dd/MM/yyyy HH:mm:ss", cult) 'Display DateTime
    End Sub
End Class
