#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Inventory_Record.Common
#End Region

Public Class FrmAddRM

#Region "Declare"
    Inherits System.Windows.Forms.Form
    Dim GrdDV As New DataView
    Protected Const TBL_RM As String = "TBL_RM"
    Dim GrdDVUnit As New DataView
    Protected Const TBL_Unit As String = "TBL_Unit"
    Dim GrdDVType As New DataView
    Protected Const TBL_Type As String = "TBL_Type"
    Friend unittext As String
    Friend unitcode As String
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
    Friend WithEvents CmdSave As System.Windows.Forms.Button
    Friend WithEvents CmdClose As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents CmbUnit As System.Windows.Forms.ComboBox
    Friend WithEvents TxtAPrice As System.Windows.Forms.TextBox
    Friend WithEvents TxtSPrice As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TxtQty As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TxtRMQty As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents cmbType As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmAddRM))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cmbType = New System.Windows.Forms.ComboBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.TxtRMQty = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.TxtQty = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.CmbUnit = New System.Windows.Forms.ComboBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtAPrice = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtSPrice = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtDesc = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtName = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.CmdSave = New System.Windows.Forms.Button
        Me.CmdClose = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.cmbType)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.TxtRMQty)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.TxtQty)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.CmbUnit)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.TxtAPrice)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.TxtSPrice)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.TxtDesc)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.TxtName)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(306, 242)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'cmbType
        '
        Me.cmbType.Location = New System.Drawing.Point(56, 208)
        Me.cmbType.Name = "cmbType"
        Me.cmbType.Size = New System.Drawing.Size(128, 21)
        Me.cmbType.TabIndex = 20
        Me.cmbType.Text = "Select"
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(16, 210)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(40, 16)
        Me.Label11.TabIndex = 19
        Me.Label11.Text = "Type"
        '
        'TxtRMQty
        '
        Me.TxtRMQty.Location = New System.Drawing.Point(56, 152)
        Me.TxtRMQty.Name = "TxtRMQty"
        Me.TxtRMQty.Size = New System.Drawing.Size(56, 20)
        Me.TxtRMQty.TabIndex = 18
        Me.TxtRMQty.Text = "1"
        Me.TxtRMQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(24, 154)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(24, 16)
        Me.Label10.TabIndex = 17
        Me.Label10.Text = "Qty"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(160, 184)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(40, 16)
        Me.Label9.TabIndex = 16
        Me.Label9.Text = "KG"
        '
        'TxtQty
        '
        Me.TxtQty.Location = New System.Drawing.Point(104, 182)
        Me.TxtQty.Name = "TxtQty"
        Me.TxtQty.Size = New System.Drawing.Size(48, 20)
        Me.TxtQty.TabIndex = 15
        Me.TxtQty.Text = "0.000"
        Me.TxtQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 184)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 16)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Q'ty/Pack"
        '
        'CmbUnit
        '
        Me.CmbUnit.Location = New System.Drawing.Point(168, 152)
        Me.CmbUnit.Name = "CmbUnit"
        Me.CmbUnit.Size = New System.Drawing.Size(128, 21)
        Me.CmbUnit.TabIndex = 13
        Me.CmbUnit.Text = "Select"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(160, 120)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(40, 16)
        Me.Label8.TabIndex = 12
        Me.Label8.Text = "Baht"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(160, 88)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(40, 16)
        Me.Label7.TabIndex = 11
        Me.Label7.Text = "Baht"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(120, 154)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(40, 16)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "Unit"
        '
        'TxtAPrice
        '
        Me.TxtAPrice.Location = New System.Drawing.Point(104, 120)
        Me.TxtAPrice.Name = "TxtAPrice"
        Me.TxtAPrice.Size = New System.Drawing.Size(48, 20)
        Me.TxtAPrice.TabIndex = 9
        Me.TxtAPrice.Text = "0.00"
        Me.TxtAPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(16, 120)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(88, 16)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Price  Actual"
        '
        'TxtSPrice
        '
        Me.TxtSPrice.Location = New System.Drawing.Point(104, 88)
        Me.TxtSPrice.Name = "TxtSPrice"
        Me.TxtSPrice.Size = New System.Drawing.Size(48, 20)
        Me.TxtSPrice.TabIndex = 7
        Me.TxtSPrice.Text = "0.00"
        Me.TxtSPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 88)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(96, 16)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Price Standard"
        '
        'TxtDesc
        '
        Me.TxtDesc.Location = New System.Drawing.Point(64, 56)
        Me.TxtDesc.Name = "TxtDesc"
        Me.TxtDesc.Size = New System.Drawing.Size(224, 20)
        Me.TxtDesc.TabIndex = 5
        Me.TxtDesc.Text = ""
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 56)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 16)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Desc."
        '
        'TxtName
        '
        Me.TxtName.Location = New System.Drawing.Point(64, 24)
        Me.TxtName.Name = "TxtName"
        Me.TxtName.Size = New System.Drawing.Size(88, 20)
        Me.TxtName.TabIndex = 3
        Me.TxtName.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 16)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Code"
        '
        'CmdSave
        '
        Me.CmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdSave.Image = CType(resources.GetObject("CmdSave.Image"), System.Drawing.Image)
        Me.CmdSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdSave.Location = New System.Drawing.Point(160, 250)
        Me.CmdSave.Name = "CmdSave"
        Me.CmdSave.Size = New System.Drawing.Size(80, 56)
        Me.CmdSave.TabIndex = 1
        Me.CmdSave.Text = "Save"
        Me.CmdSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CmdClose
        '
        Me.CmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdClose.Image = CType(resources.GetObject("CmdClose.Image"), System.Drawing.Image)
        Me.CmdClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CmdClose.Location = New System.Drawing.Point(240, 250)
        Me.CmdClose.Name = "CmdClose"
        Me.CmdClose.Size = New System.Drawing.Size(72, 56)
        Me.CmdClose.TabIndex = 2
        Me.CmdClose.Text = "Close"
        Me.CmdClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'FrmAddRM
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(322, 312)
        Me.Controls.Add(Me.CmdClose)
        Me.Controls.Add(Me.CmdSave)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmAddRM"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "R/M  "
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "CONSTANT"
    Dim DT As New DataTable
    Dim dtUnit As New DataTable
    Dim StrSQL As String
    Dim oldrow As Integer
    Dim C1 As New SQLData("ACCINV")
    Dim strDate, idate, iMonth, iYear, iTime, STime() As String
    Dim im As Integer
#End Region

#Region "COMBOBOX"
    Sub LoadCmbUnit()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        StrSQL = "SELECT  UnitCode,Unitname "
        StrSQL &= "  FROM  TblUnit "

        Dim C1 As New SQLData("ACCINV")
        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSQL, C1.Strcon)
            Dim CBu As New SqlCommandBuilder(DA)
            dtUnit = New DataTable
            DA.Fill(dtUnit)
        Catch
        Finally
        End Try
        dtUnit.TableName = TBL_Unit
        GrdDVUnit = dtUnit.DefaultView
        '************************************
        CmbUnit.DisplayMember = "Unitname"
        CmbUnit.ValueMember = "UnitCode"
        CmbUnit.DataSource = dtUnit
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
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
        cmbType.DisplayMember = "Typename"
        cmbType.ValueMember = "TypeCode"
        cmbType.DataSource = dt
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
#End Region

#Region "RM"
    Private Function ChkData() As Boolean
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Try
            strSQL &= " select count(*) from TblRM "
            strSQL &= " where RMcode  = '" & TxtName.Text.Trim & "'"
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
            Me.Close()
        Catch Exp As SqlException
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "SQL Error")
        Catch Exp As Exception
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "General Error")
        End Try
    End Function
    Sub RM()
        If CmdSave.Text = "Save" Then
            Dim cnSQL As SqlConnection
            Dim cmSQL As SqlCommand
            Dim strSQL As String = String.Empty
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
            Try
                strSQL &= " Insert  TblRM "
                strSQL &= " values ( '" & TxtName.Text.Trim & " ',"
                strSQL &= PrepareStr(TxtDesc.Text.Trim) & ","
                strSQL &= PrepareStr(TxtSPrice.Text.Trim) & ","
                strSQL &= PrepareStr(TxtAPrice.Text.Trim) & ","
                strSQL &= PrepareStr(cmbType.SelectedValue) & ","
                strSQL &= PrepareStr(CmbUnit.SelectedValue) & ","
                strSQL &= PrepareStr(strDate) & ","
                strSQL &= PrepareStr(iTime) & " )"

                strSQL &= ""
                strSQL &= " Insert  TblGroup "
                strSQL &= " values ( " & PrepareStr(cmbType.SelectedValue) & ","
                strSQL &= PrepareStr(TxtName.Text.Trim) & ")"

                strSQL &= ""
                strSQL &= " Insert TBLConvert "
                strSQL &= " Values(" & PrepareStr(cmbType.SelectedValue)
                strSQL &= "," & PrepareStr("")
                strSQL &= "," & PrepareStr(TxtName.Text.Trim)
                strSQL &= "," & PrepareStr("")
                strSQL &= "," & PrepareStr(CmbUnit.SelectedValue)
                strSQL &= "," & PrepareStr("KG")
                strSQL &= "," & PrepareStr(TxtRMQty.Text.Trim)
                strSQL &= "," & PrepareStr(TxtQty.Text.Trim)
                strSQL &= ")"

                If TxtQty.Text <> 0 Then
                    strSQL &= ""
                    strSQL &= " Insert  TblQtyUnit "
                    strSQL &= " values ( '" & TxtName.Text.Trim & " ',"
                    strSQL &= PrepareStr(TxtRMQty.Text.Trim) & ","
                    strSQL &= PrepareStr(CmbUnit.SelectedValue) & ","
                    strSQL &= PrepareStr(TxtQty.Text.Trim) & ","
                    strSQL &= PrepareStr("KG") & ","
                    strSQL &= PrepareStr(strDate) & ","
                    strSQL &= PrepareStr(iTime) & " )"
                Else
                    strSQL &= ""
                    strSQL &= " Insert  TblQtyUnit "
                    strSQL &= " values ( '" & TxtName.Text.Trim & " ',"
                    strSQL &= PrepareStr(TxtRMQty.Text.Trim) & ","
                    strSQL &= PrepareStr(CmbUnit.SelectedValue) & ","
                    strSQL &= PrepareStr(0) & ","
                    strSQL &= PrepareStr("KG") & " ,"
                    strSQL &= PrepareStr(strDate) & ","
                    strSQL &= PrepareStr(iTime) & " )"
                End If

                cnSQL = New SqlConnection(C1.Strcon)
                cnSQL.Open()
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
            Me.Cursor = System.Windows.Forms.Cursors.Default()
        ElseIf CmdSave.Text = "Edit" Then
            Dim cnSQL As SqlConnection
            Dim cmSQL As SqlCommand
            Dim strSQL As String = String.Empty
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
            Try
                strSQL &= " Update TblRM"
                strSQL &= " set descName = '" & TxtDesc.Text.Trim & "'"
                strSQL &= " , StdPrice = '" & TxtSPrice.Text.Trim & "'"
                strSQL &= " , ActPrice = '" & TxtAPrice.Text.Trim & "'"
                strSQL &= " , Unit = '" & CmbUnit.SelectedValue & "'"
                strSQL &= " where RMCode = '" & TxtName.Text.Trim & "'"

                strSQL &= ""
                strSQL &= " Update TblQtyUnit"
                strSQL &= " set Qty = '" & TxtQty.Text.Trim & "'"
                strSQL &= " , RMQty = '" & TxtRMQty.Text.Trim & "'"
                strSQL &= " , UnitCode = '" & CmbUnit.SelectedValue & "'"
                strSQL &= " where RMCode = '" & TxtName.Text.Trim & "'"

                strSQL &= ""
                strSQL &= " Update TblConvert"
                strSQL &= " set SQty = '" & TxtQty.Text.Trim & "'"
                strSQL &= " , BQty = '" & TxtRMQty.Text.Trim & "'"
                strSQL &= " , UnitBig = '" & CmbUnit.SelectedValue & "'"
                strSQL &= " where Code = '" & TxtName.Text.Trim & "'"
                strSQL &= " and  UnitBig = '" & unitcode.Trim & "'"

                cnSQL = New SqlConnection(C1.Strcon)
                cnSQL.Open()
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
            Me.Cursor = System.Windows.Forms.Cursors.Default()
        Else
        End If

    End Sub
#End Region

    Private Sub CmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClose.Click
        Me.Close()
    End Sub
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

    Private Sub CmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSave.Click
        Dim i, j As Double
        i = TxtSPrice.Text.Trim
        j = TxtAPrice.Text.Trim
        TxtSPrice.Text = Format(i, "##,###,###.00")
        TxtAPrice.Text = Format(j, "##,###,###.00")
        TxtDesc.Text = TxtDesc.Text.ToUpper
        TxtName.Text = TxtName.Text.ToUpper
        idate = Date.Now.Day
        im = Date.Now.Month
        iMonth = Format(im, "00")
        iYear = Date.Now.Year
        STime = Split(Date.Now.ToShortTimeString, ":")
        strDate = iYear + iMonth + idate
        iTime = STime(0) + STime(1)
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        msg = "R/M Meterial" ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "R/M Warehouse"   ' Define title.
        If TxtName.Text.Trim = "" Then
            TxtName.Focus()
            Exit Sub
        Else
        End If

        If CmdSave.Text = "Save" Then
            If ChkData() = True Then
                MsgBox("It's Duplicate.", MsgBoxStyle.Critical, "Deprtment")
                TxtName.Focus()
                Exit Sub
            Else
            End If
        Else
        End If

        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            RM()
        Else
            Exit Sub
        End If
    End Sub

    Private Sub FrmAddRM_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadCmbUnit()
        LoadCmbType()
        If CmdSave.Text = "Save" Then
        Else
            CmbUnit.Text = unittext.Trim
            cmbType.Enabled = False
        End If
    End Sub

    Private Sub TxtSPrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtSPrice.KeyPress
        Dim i As Double
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                SendKeys.Send("{TAB}")
                i = TxtSPrice.Text.Trim
                TxtSPrice.Text = Format(i, "##,###,###.00")
            Case Else
        End Select
    End Sub

    Private Sub TxtAPrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtAPrice.KeyPress
        Dim i As Double
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                SendKeys.Send("{TAB}")
                i = TxtAPrice.Text.Trim
                TxtAPrice.Text = Format(i, "##,###,###.00")
            Case Else
        End Select
    End Sub

    Private Sub TxtName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtName.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                SendKeys.Send("{TAB}")
                TxtName.Text = TxtName.Text.ToUpper
            Case Else
        End Select
    End Sub

    Private Sub TxtDesc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtDesc.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
                e.Handled = True
                SendKeys.Send("{TAB}")
                TxtDesc.Text = TxtDesc.Text.ToUpper
            Case Else
        End Select
    End Sub
End Class
