#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Inventory_Record.Common
#End Region

Public Class Calculate
    Inherits System.Windows.Forms.Form
    Dim C1 As New SQLData("ACCINV")

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
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CKPigment As System.Windows.Forms.CheckBox
    Friend WithEvents CKRM As System.Windows.Forms.CheckBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents CKCompound As System.Windows.Forms.CheckBox
    Friend WithEvents CKPresemi As System.Windows.Forms.CheckBox
    Friend WithEvents CKSemi As System.Windows.Forms.CheckBox
    Friend WithEvents CKGT As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Calculate))
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.Button1 = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.CKGT = New System.Windows.Forms.CheckBox
        Me.CKSemi = New System.Windows.Forms.CheckBox
        Me.CKPresemi = New System.Windows.Forms.CheckBox
        Me.CKCompound = New System.Windows.Forms.CheckBox
        Me.CKRM = New System.Windows.Forms.CheckBox
        Me.CKPigment = New System.Windows.Forms.CheckBox
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(8, 184)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(312, 23)
        Me.ProgressBar1.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(112, 136)
        Me.Button1.Name = "Button1"
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Calculate"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CKGT)
        Me.GroupBox1.Controls.Add(Me.CKSemi)
        Me.GroupBox1.Controls.Add(Me.CKPresemi)
        Me.GroupBox1.Controls.Add(Me.CKCompound)
        Me.GroupBox1.Controls.Add(Me.CKRM)
        Me.GroupBox1.Controls.Add(Me.CKPigment)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(312, 168)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        '
        'CKGT
        '
        Me.CKGT.Checked = True
        Me.CKGT.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CKGT.Location = New System.Drawing.Point(176, 88)
        Me.CKGT.Name = "CKGT"
        Me.CKGT.TabIndex = 5
        Me.CKGT.Text = "Green Tire"
        '
        'CKSemi
        '
        Me.CKSemi.Checked = True
        Me.CKSemi.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CKSemi.Location = New System.Drawing.Point(176, 56)
        Me.CKSemi.Name = "CKSemi"
        Me.CKSemi.TabIndex = 4
        Me.CKSemi.Text = "Semi"
        '
        'CKPresemi
        '
        Me.CKPresemi.Checked = True
        Me.CKPresemi.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CKPresemi.Location = New System.Drawing.Point(176, 24)
        Me.CKPresemi.Name = "CKPresemi"
        Me.CKPresemi.TabIndex = 3
        Me.CKPresemi.Text = "Presemi"
        '
        'CKCompound
        '
        Me.CKCompound.Checked = True
        Me.CKCompound.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CKCompound.Location = New System.Drawing.Point(32, 88)
        Me.CKCompound.Name = "CKCompound"
        Me.CKCompound.TabIndex = 2
        Me.CKCompound.Text = "Compound"
        '
        'CKRM
        '
        Me.CKRM.Checked = True
        Me.CKRM.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CKRM.Location = New System.Drawing.Point(32, 24)
        Me.CKRM.Name = "CKRM"
        Me.CKRM.TabIndex = 1
        Me.CKRM.Text = "R/M (Material)"
        '
        'CKPigment
        '
        Me.CKPigment.Checked = True
        Me.CKPigment.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CKPigment.Location = New System.Drawing.Point(32, 56)
        Me.CKPigment.Name = "CKPigment"
        Me.CKPigment.TabIndex = 0
        Me.CKPigment.Text = "Pigment"
        '
        'Calculate
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Linen
        Me.ClientSize = New System.Drawing.Size(328, 214)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ProgressBar1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Calculate"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Calculate"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim StrDate, Str(), StrTime As String
    Dim strType As String
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If CKRM.Checked And CKPigment.Checked And CKCompound.Checked And CKPresemi.Checked And CKSemi.Checked And CKGT.Checked Then
            DelPrice(strType)
        End If

        If Progress() Then
            MsgBox(" Calculate Complete. ", MsgBoxStyle.OKOnly, "Calculate")
            Me.Close()
        Else
            MsgBox(" Calculate Not Complete. Please Check Data. Calculate by Process ", MsgBoxStyle.OKOnly, "Calculate")
        End If
    End Sub
    Sub CKdata()
        If CKRM.Checked Then
            strType &= "'01'"
        End If
        If CKPigment.Checked Then
            strType &= "'02'"
        End If
        If CKCompound.Checked Then
            strType &= "'03'"
        End If
        If CKPresemi.Checked Then
            strType &= "'04'"
        End If
        If CKSemi.Checked Then
            strType &= "'05'"
        End If
        If CKGT.Checked Then
            strType &= "'06'"
        End If
    End Sub
    Private Sub Calculate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Str = Split(Now.Date.ToShortDateString, "/")
        StrDate = Str(2) & Str(1) & Str(0)
        StrTime = Format(Now.TimeOfDay.Hours, "00") & Format(Now.TimeOfDay.Minutes, "00")
    End Sub

    Function Progress(ByVal ParamArray filenames As String()) As Boolean
        Progress = False
        ' Display the ProgressBar control.
        ProgressBar1.Visible = True
        ' Set Minimum to 1 to represent the first file being copied.
        ProgressBar1.Minimum = 0
        ' Set Maximum to the total number of files to copy.
        ProgressBar1.Maximum = 100
        ' Set the initial value of the ProgressBar.
        ProgressBar1.Value = 1
        ' Set the Step property to a value of 1 to represent each file being copied.
        ProgressBar1.Step = 1

        ' Loop through all files to copy.
        Dim x As Integer
        For x = 1 To 100
            ProgressBar1.PerformStep()
            If x = 1 Then
                If CKRM.Checked Then
                    CalRM(StrDate, StrTime)
                Else
                    x = 10
                End If
            Else
            End If

            If x = 10 Then
                If CKPigment.Checked Then
                    CalPigment(StrDate, StrTime)
                Else
                    x = 20
                End If
            End If

            If x = 20 Then
                If CKCompound.Checked Then
                    CalCompound(StrDate, StrTime)
                    CalCompound2(StrDate, StrTime)
                Else
                    x = 30
                End If
            End If

            If x = 30 Then
                If CKPresemi.Checked Then
                    CalCoatedcord2(StrDate, StrTime)
                    CalCoatedcord3(StrDate, StrTime)
                    CalCoatedcord(StrDate, StrTime)
                Else
                    x = 40
                End If
            End If

            If x = 40 Then
                If CKPresemi.Checked Then
                    PreSemi(StrDate, StrTime)
                    SteelCord(StrDate, StrTime)
                Else
                    x = 50
                End If
            End If

            If x = 50 Then
                If CKSemi.Checked Then
                    Tread(StrDate, StrTime)
                    BF(StrDate, StrTime)
                    BF2(StrDate, StrTime)
                Else
                    x = 70
                End If
            End If

            If x = 70 Then
                If CKSemi.Checked Then
                    BElT(StrDate, StrTime) 'Type Material:Belt-1, Belt-2, Belt-3, Belt-4 to Table TBLMASTERPRICE
                    BElT2(StrDate, StrTime) 'Type Material:Belt-1, Belt-2, Belt-3, Belt-4 to Table TBLMASTERPRICERM
                    SIC(StrDate, StrTime) 'Type Material:Cussion, Side, Innerliner
                    Chafer(StrDate, StrTime) 'Type Material:Body Ply, Wire Chafer, Nylon Chafer to Table TBLMASTERPRICE
                    Chafer2(StrDate, StrTime) 'Type Material:Body Ply, Wire Chafer, Nylon Chafer to Table TBLMASTERPRICERM
                Else
                    x = 90
                End If
            End If

            If x = 90 Then
                If CKGT.Checked Then
                    GT(StrDate, StrTime)
                Else
                    x = 100
                End If
            End If
            If x = 100 Then
                Progress = True
            Else
                Progress = False
            End If
        Next x
    End Function

#Region "Call Store"

#Region "CalRM Price"
    Private Function CalRM(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        CalRM = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CALRM"
        cmd2.Connection = cnn

        Dim sparam0 As SqlClient.SqlParameter
        sparam0 = New SqlClient.SqlParameter
        sparam0.ParameterName = "@Date"
        sparam0.SqlDbType = SqlDbType.Char
        sparam0.Size = 8
        sparam0.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam0)

        Dim sparam1 As SqlClient.SqlParameter
        sparam1 = New SqlClient.SqlParameter
        sparam1.ParameterName = "@Time"
        sparam1.SqlDbType = SqlDbType.Char
        sparam1.Size = 8
        sparam1.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam1)

        Dim sparam2 As SqlClient.SqlParameter
        sparam2 = New SqlClient.SqlParameter
        sparam2.ParameterName = "@errID"
        sparam2.SqlDbType = SqlDbType.Char
        sparam2.Size = 4
        sparam2.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam2)

        Dim sparam3 As SqlClient.SqlParameter
        sparam3 = New SqlClient.SqlParameter
        sparam3.ParameterName = "@errMsg"
        sparam3.SqlDbType = SqlDbType.Char
        sparam3.Size = 40
        sparam3.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam3)

        Dim Reader As SqlClient.SqlDataReader
        cmd2.Parameters("@Date").Value = dateup.Trim
        cmd2.Parameters("@Time").Value = Timeup.Trim

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            CalRM = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            CalRM = False
        End Try
        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Function
#End Region

#Region "Pigment CAL"
    Private Function CalPigment(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        CalPigment = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CALPigment"
        cmd2.Connection = cnn

        Dim sparam0 As SqlClient.SqlParameter
        sparam0 = New SqlClient.SqlParameter
        sparam0.ParameterName = "@Date"
        sparam0.SqlDbType = SqlDbType.Char
        sparam0.Size = 8
        sparam0.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam0)

        Dim sparam1 As SqlClient.SqlParameter
        sparam1 = New SqlClient.SqlParameter
        sparam1.ParameterName = "@Time"
        sparam1.SqlDbType = SqlDbType.Char
        sparam1.Size = 8
        sparam1.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam1)

        Dim sparam2 As SqlClient.SqlParameter
        sparam2 = New SqlClient.SqlParameter
        sparam2.ParameterName = "@errID"
        sparam2.SqlDbType = SqlDbType.Char
        sparam2.Size = 4
        sparam2.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam2)

        Dim sparam3 As SqlClient.SqlParameter
        sparam3 = New SqlClient.SqlParameter
        sparam3.ParameterName = "@errMsg"
        sparam3.SqlDbType = SqlDbType.Char
        sparam3.Size = 40
        sparam3.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam3)

        Dim Reader As SqlClient.SqlDataReader
        cmd2.Parameters("@Date").Value = dateup.Trim
        cmd2.Parameters("@Time").Value = Timeup.Trim

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            CalPigment = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            CalPigment = False
        End Try
        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Function
#End Region

#Region "Calcompound Price"
    Private Function CalCompound(ByVal dateup As String, ByVal timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        CalCompound = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CALCompound"
        cmd2.Connection = cnn

        Dim sparam0 As SqlClient.SqlParameter
        sparam0 = New SqlClient.SqlParameter
        sparam0.ParameterName = "@Date"
        sparam0.SqlDbType = SqlDbType.Char
        sparam0.Size = 8
        sparam0.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam0)

        Dim sparam1 As SqlClient.SqlParameter
        sparam1 = New SqlClient.SqlParameter
        sparam1.ParameterName = "@Time"
        sparam1.SqlDbType = SqlDbType.Char
        sparam1.Size = 8
        sparam1.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam1)

        Dim sparam2 As SqlClient.SqlParameter
        sparam2 = New SqlClient.SqlParameter
        sparam2.ParameterName = "@errID"
        sparam2.SqlDbType = SqlDbType.Char
        sparam2.Size = 4
        sparam2.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam2)

        Dim sparam3 As SqlClient.SqlParameter
        sparam3 = New SqlClient.SqlParameter
        sparam3.ParameterName = "@errMsg"
        sparam3.SqlDbType = SqlDbType.Char
        sparam3.Size = 40
        sparam3.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam3)

        Dim Reader As SqlClient.SqlDataReader
        cmd2.Parameters("@Date").Value = dateup.Trim
        cmd2.Parameters("@Time").Value = timeup.Trim

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            CalCompound = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            CalCompound = False
        End Try
        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Function
    Private Function CalCompound2(ByVal dateup As String, ByVal timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        CalCompound2 = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CalCompound2"
        cmd2.Connection = cnn

        Dim sparam0 As SqlClient.SqlParameter
        sparam0 = New SqlClient.SqlParameter
        sparam0.ParameterName = "@Date"
        sparam0.SqlDbType = SqlDbType.Char
        sparam0.Size = 8
        sparam0.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam0)

        Dim sparam1 As SqlClient.SqlParameter
        sparam1 = New SqlClient.SqlParameter
        sparam1.ParameterName = "@Time"
        sparam1.SqlDbType = SqlDbType.Char
        sparam1.Size = 8
        sparam1.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam1)

        Dim sparam2 As SqlClient.SqlParameter
        sparam2 = New SqlClient.SqlParameter
        sparam2.ParameterName = "@errID"
        sparam2.SqlDbType = SqlDbType.Char
        sparam2.Size = 4
        sparam2.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam2)

        Dim sparam3 As SqlClient.SqlParameter
        sparam3 = New SqlClient.SqlParameter
        sparam3.ParameterName = "@errMsg"
        sparam3.SqlDbType = SqlDbType.Char
        sparam3.Size = 40
        sparam3.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam3)

        Dim Reader As SqlClient.SqlDataReader
        cmd2.Parameters("@Date").Value = dateup.Trim
        cmd2.Parameters("@Time").Value = timeup.Trim

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            CalCompound2 = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            CalCompound2 = False
        End Try
        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Function
#End Region

#Region "CalCoatedcord"
    Private Function CalCoatedcord(ByVal dateup As String, ByVal timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        CalCoatedcord = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CALCoatedCord"
        cmd2.Connection = cnn

        Dim sparam0 As SqlClient.SqlParameter
        sparam0 = New SqlClient.SqlParameter
        sparam0.ParameterName = "@Date"
        sparam0.SqlDbType = SqlDbType.Char
        sparam0.Size = 8
        sparam0.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam0)

        Dim sparam1 As SqlClient.SqlParameter
        sparam1 = New SqlClient.SqlParameter
        sparam1.ParameterName = "@Time"
        sparam1.SqlDbType = SqlDbType.Char
        sparam1.Size = 8
        sparam1.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam1)

        Dim sparam2 As SqlClient.SqlParameter
        sparam2 = New SqlClient.SqlParameter
        sparam2.ParameterName = "@errID"
        sparam2.SqlDbType = SqlDbType.Char
        sparam2.Size = 4
        sparam2.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam2)

        Dim sparam3 As SqlClient.SqlParameter
        sparam3 = New SqlClient.SqlParameter
        sparam3.ParameterName = "@errMsg"
        sparam3.SqlDbType = SqlDbType.Char
        sparam3.Size = 40
        sparam3.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam3)

        Dim Reader As SqlClient.SqlDataReader
        cmd2.Parameters("@Date").Value = dateup.Trim
        cmd2.Parameters("@Time").Value = timeup.Trim

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            CalCoatedcord = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            CalCoatedcord = False
        End Try
        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Function
    Private Function CalCoatedcord2(ByVal dateup As String, ByVal timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CalCoatedcord2"
        cmd2.Connection = cnn

        Dim sparam0 As SqlClient.SqlParameter
        sparam0 = New SqlClient.SqlParameter
        sparam0.ParameterName = "@Date"
        sparam0.SqlDbType = SqlDbType.Char
        sparam0.Size = 8
        sparam0.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam0)

        Dim sparam1 As SqlClient.SqlParameter
        sparam1 = New SqlClient.SqlParameter
        sparam1.ParameterName = "@Time"
        sparam1.SqlDbType = SqlDbType.Char
        sparam1.Size = 8
        sparam1.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam1)

        Dim sparam2 As SqlClient.SqlParameter
        sparam2 = New SqlClient.SqlParameter
        sparam2.ParameterName = "@errID"
        sparam2.SqlDbType = SqlDbType.Char
        sparam2.Size = 4
        sparam2.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam2)

        Dim sparam3 As SqlClient.SqlParameter
        sparam3 = New SqlClient.SqlParameter
        sparam3.ParameterName = "@errMsg"
        sparam3.SqlDbType = SqlDbType.Char
        sparam3.Size = 40
        sparam3.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam3)

        Dim Reader As SqlClient.SqlDataReader
        cmd2.Parameters("@Date").Value = dateup.Trim
        cmd2.Parameters("@Time").Value = timeup.Trim

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
        Catch ex As Exception
            MsgBox(ex.Message, 48)
        End Try
        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()

        Return True
    End Function
    Private Function CalCoatedcord3(ByVal dateup As String, ByVal timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CalCoatedcord3"
        cmd2.Connection = cnn

        Dim sparam0 As SqlClient.SqlParameter
        sparam0 = New SqlClient.SqlParameter
        sparam0.ParameterName = "@Date"
        sparam0.SqlDbType = SqlDbType.Char
        sparam0.Size = 8
        sparam0.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam0)

        Dim sparam1 As SqlClient.SqlParameter
        sparam1 = New SqlClient.SqlParameter
        sparam1.ParameterName = "@Time"
        sparam1.SqlDbType = SqlDbType.Char
        sparam1.Size = 8
        sparam1.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam1)

        Dim sparam2 As SqlClient.SqlParameter
        sparam2 = New SqlClient.SqlParameter
        sparam2.ParameterName = "@errID"
        sparam2.SqlDbType = SqlDbType.Char
        sparam2.Size = 4
        sparam2.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam2)

        Dim sparam3 As SqlClient.SqlParameter
        sparam3 = New SqlClient.SqlParameter
        sparam3.ParameterName = "@errMsg"
        sparam3.SqlDbType = SqlDbType.Char
        sparam3.Size = 40
        sparam3.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam3)

        Dim Reader As SqlClient.SqlDataReader
        cmd2.Parameters("@Date").Value = dateup.Trim
        cmd2.Parameters("@Time").Value = timeup.Trim

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
        Catch ex As Exception
            MsgBox(ex.Message, 48)
        End Try
        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()

        Return True
    End Function
#End Region

#Region "CalPreSemi Price"
    Private Function PreSemi(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        PreSemi = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CALPreSemi"
        cmd2.Connection = cnn

        Dim sparam0 As SqlClient.SqlParameter
        sparam0 = New SqlClient.SqlParameter
        sparam0.ParameterName = "@Date"
        sparam0.SqlDbType = SqlDbType.Char
        sparam0.Size = 8
        sparam0.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam0)

        Dim sparam1 As SqlClient.SqlParameter
        sparam1 = New SqlClient.SqlParameter
        sparam1.ParameterName = "@Time"
        sparam1.SqlDbType = SqlDbType.Char
        sparam1.Size = 8
        sparam1.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam1)

        Dim sparam2 As SqlClient.SqlParameter
        sparam2 = New SqlClient.SqlParameter
        sparam2.ParameterName = "@errID"
        sparam2.SqlDbType = SqlDbType.Char
        sparam2.Size = 4
        sparam2.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam2)

        Dim sparam3 As SqlClient.SqlParameter
        sparam3 = New SqlClient.SqlParameter
        sparam3.ParameterName = "@errMsg"
        sparam3.SqlDbType = SqlDbType.Char
        sparam3.Size = 40
        sparam3.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam3)

        Dim Reader As SqlClient.SqlDataReader
        cmd2.Parameters("@Date").Value = dateup.Trim
        cmd2.Parameters("@Time").Value = Timeup.Trim

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            PreSemi = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            PreSemi = False
        End Try
        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Function
#End Region

#Region "CalPreSemi Price"
    Private Function SteelCord(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        SteelCord = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CALSteelCord"
        cmd2.Connection = cnn

        Dim sparam0 As SqlClient.SqlParameter
        sparam0 = New SqlClient.SqlParameter
        sparam0.ParameterName = "@Date"
        sparam0.SqlDbType = SqlDbType.Char
        sparam0.Size = 8
        sparam0.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam0)

        Dim sparam1 As SqlClient.SqlParameter
        sparam1 = New SqlClient.SqlParameter
        sparam1.ParameterName = "@Time"
        sparam1.SqlDbType = SqlDbType.Char
        sparam1.Size = 8
        sparam1.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam1)

        Dim sparam2 As SqlClient.SqlParameter
        sparam2 = New SqlClient.SqlParameter
        sparam2.ParameterName = "@errID"
        sparam2.SqlDbType = SqlDbType.Char
        sparam2.Size = 4
        sparam2.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam2)

        Dim sparam3 As SqlClient.SqlParameter
        sparam3 = New SqlClient.SqlParameter
        sparam3.ParameterName = "@errMsg"
        sparam3.SqlDbType = SqlDbType.Char
        sparam3.Size = 40
        sparam3.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam3)

        Dim Reader As SqlClient.SqlDataReader
        cmd2.Parameters("@Date").Value = dateup.Trim
        cmd2.Parameters("@Time").Value = Timeup.Trim

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            SteelCord = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            SteelCord = False
        End Try
        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Function
#End Region

#Region "CalSemi Price"
#Region "TT"
    Private Function Tread(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Tread = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CALTT"
        cmd2.Connection = cnn

        Dim sparam0 As SqlClient.SqlParameter
        sparam0 = New SqlClient.SqlParameter
        sparam0.ParameterName = "@Date"
        sparam0.SqlDbType = SqlDbType.Char
        sparam0.Size = 8
        sparam0.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam0)

        Dim sparam1 As SqlClient.SqlParameter
        sparam1 = New SqlClient.SqlParameter
        sparam1.ParameterName = "@Time"
        sparam1.SqlDbType = SqlDbType.Char
        sparam1.Size = 8
        sparam1.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam1)

        Dim sparam2 As SqlClient.SqlParameter
        sparam2 = New SqlClient.SqlParameter
        sparam2.ParameterName = "@errID"
        sparam2.SqlDbType = SqlDbType.Char
        sparam2.Size = 4
        sparam2.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam2)

        Dim sparam3 As SqlClient.SqlParameter
        sparam3 = New SqlClient.SqlParameter
        sparam3.ParameterName = "@errMsg"
        sparam3.SqlDbType = SqlDbType.Char
        sparam3.Size = 40
        sparam3.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam3)

        Dim Reader As SqlClient.SqlDataReader
        cmd2.Parameters("@Date").Value = dateup.Trim
        cmd2.Parameters("@Time").Value = Timeup.Trim

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            Tread = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            Tread = False
        End Try
        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Function
#End Region

#Region "BF"
    Private Function BF(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        BF = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CALBF"
        cmd2.Connection = cnn

        Dim sparam0 As SqlClient.SqlParameter
        sparam0 = New SqlClient.SqlParameter
        sparam0.ParameterName = "@Date"
        sparam0.SqlDbType = SqlDbType.Char
        sparam0.Size = 8
        sparam0.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam0)

        Dim sparam1 As SqlClient.SqlParameter
        sparam1 = New SqlClient.SqlParameter
        sparam1.ParameterName = "@Time"
        sparam1.SqlDbType = SqlDbType.Char
        sparam1.Size = 8
        sparam1.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam1)

        Dim sparam2 As SqlClient.SqlParameter
        sparam2 = New SqlClient.SqlParameter
        sparam2.ParameterName = "@errID"
        sparam2.SqlDbType = SqlDbType.Char
        sparam2.Size = 4
        sparam2.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam2)

        Dim sparam3 As SqlClient.SqlParameter
        sparam3 = New SqlClient.SqlParameter
        sparam3.ParameterName = "@errMsg"
        sparam3.SqlDbType = SqlDbType.Char
        sparam3.Size = 40
        sparam3.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam3)

        Dim Reader As SqlClient.SqlDataReader
        cmd2.Parameters("@Date").Value = dateup.Trim
        cmd2.Parameters("@Time").Value = Timeup.Trim

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            BF = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            BF = False
        End Try
        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Function
    Private Function BF2(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        BF2 = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CALBF2"
        cmd2.Connection = cnn

        Dim sparam0 As SqlClient.SqlParameter
        sparam0 = New SqlClient.SqlParameter
        sparam0.ParameterName = "@Date"
        sparam0.SqlDbType = SqlDbType.Char
        sparam0.Size = 8
        sparam0.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam0)

        Dim sparam1 As SqlClient.SqlParameter
        sparam1 = New SqlClient.SqlParameter
        sparam1.ParameterName = "@Time"
        sparam1.SqlDbType = SqlDbType.Char
        sparam1.Size = 8
        sparam1.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam1)

        Dim sparam2 As SqlClient.SqlParameter
        sparam2 = New SqlClient.SqlParameter
        sparam2.ParameterName = "@errID"
        sparam2.SqlDbType = SqlDbType.Char
        sparam2.Size = 4
        sparam2.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam2)

        Dim sparam3 As SqlClient.SqlParameter
        sparam3 = New SqlClient.SqlParameter
        sparam3.ParameterName = "@errMsg"
        sparam3.SqlDbType = SqlDbType.Char
        sparam3.Size = 40
        sparam3.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam3)

        Dim Reader As SqlClient.SqlDataReader
        cmd2.Parameters("@Date").Value = dateup.Trim
        cmd2.Parameters("@Time").Value = Timeup.Trim

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            BF2 = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            BF2 = False
        End Try
        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Function
#End Region

#Region "BElT 1-4"
    Private Function BElT(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        BElT = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CALBElT"
        cmd2.Connection = cnn

        Dim sparam0 As SqlClient.SqlParameter
        sparam0 = New SqlClient.SqlParameter
        sparam0.ParameterName = "@Date"
        sparam0.SqlDbType = SqlDbType.Char
        sparam0.Size = 8
        sparam0.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam0)

        Dim sparam1 As SqlClient.SqlParameter
        sparam1 = New SqlClient.SqlParameter
        sparam1.ParameterName = "@Time"
        sparam1.SqlDbType = SqlDbType.Char
        sparam1.Size = 8
        sparam1.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam1)

        Dim sparam2 As SqlClient.SqlParameter
        sparam2 = New SqlClient.SqlParameter
        sparam2.ParameterName = "@errID"
        sparam2.SqlDbType = SqlDbType.Char
        sparam2.Size = 4
        sparam2.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam2)

        Dim sparam3 As SqlClient.SqlParameter
        sparam3 = New SqlClient.SqlParameter
        sparam3.ParameterName = "@errMsg"
        sparam3.SqlDbType = SqlDbType.Char
        sparam3.Size = 40
        sparam3.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam3)

        Dim Reader As SqlClient.SqlDataReader
        cmd2.Parameters("@Date").Value = dateup.Trim
        cmd2.Parameters("@Time").Value = Timeup.Trim

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            BElT = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            BElT = False
        End Try
        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Function
    Private Function BElT2(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        BElT2 = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CALBElT2"
        cmd2.Connection = cnn

        Dim sparam0 As SqlClient.SqlParameter
        sparam0 = New SqlClient.SqlParameter
        sparam0.ParameterName = "@Date"
        sparam0.SqlDbType = SqlDbType.Char
        sparam0.Size = 8
        sparam0.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam0)

        Dim sparam1 As SqlClient.SqlParameter
        sparam1 = New SqlClient.SqlParameter
        sparam1.ParameterName = "@Time"
        sparam1.SqlDbType = SqlDbType.Char
        sparam1.Size = 8
        sparam1.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam1)

        Dim sparam2 As SqlClient.SqlParameter
        sparam2 = New SqlClient.SqlParameter
        sparam2.ParameterName = "@errID"
        sparam2.SqlDbType = SqlDbType.Char
        sparam2.Size = 4
        sparam2.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam2)

        Dim sparam3 As SqlClient.SqlParameter
        sparam3 = New SqlClient.SqlParameter
        sparam3.ParameterName = "@errMsg"
        sparam3.SqlDbType = SqlDbType.Char
        sparam3.Size = 40
        sparam3.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam3)

        Dim Reader As SqlClient.SqlDataReader
        cmd2.Parameters("@Date").Value = dateup.Trim
        cmd2.Parameters("@Time").Value = Timeup.Trim

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            BElT2 = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            BElT2 = False
        End Try
        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Function
#End Region

#Region "SIDE CUSSION INNERLINER"
    Private Function SIC(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        SIC = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CALSIC"
        cmd2.Connection = cnn

        Dim sparam0 As SqlClient.SqlParameter
        sparam0 = New SqlClient.SqlParameter
        sparam0.ParameterName = "@Date"
        sparam0.SqlDbType = SqlDbType.Char
        sparam0.Size = 8
        sparam0.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam0)

        Dim sparam1 As SqlClient.SqlParameter
        sparam1 = New SqlClient.SqlParameter
        sparam1.ParameterName = "@Time"
        sparam1.SqlDbType = SqlDbType.Char
        sparam1.Size = 8
        sparam1.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam1)

        Dim sparam2 As SqlClient.SqlParameter
        sparam2 = New SqlClient.SqlParameter
        sparam2.ParameterName = "@errID"
        sparam2.SqlDbType = SqlDbType.Char
        sparam2.Size = 4
        sparam2.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam2)

        Dim sparam3 As SqlClient.SqlParameter
        sparam3 = New SqlClient.SqlParameter
        sparam3.ParameterName = "@errMsg"
        sparam3.SqlDbType = SqlDbType.Char
        sparam3.Size = 40
        sparam3.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam3)

        Dim Reader As SqlClient.SqlDataReader
        cmd2.Parameters("@Date").Value = dateup.Trim
        cmd2.Parameters("@Time").Value = Timeup.Trim

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            SIC = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            SIC = False
        End Try
        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Function
#End Region

#Region "Wire Chafer,Nylon Chafer,Body Ply"
    Private Function Chafer(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Chafer = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CALCHAFER"
        cmd2.Connection = cnn

        Dim sparam0 As SqlClient.SqlParameter
        sparam0 = New SqlClient.SqlParameter
        sparam0.ParameterName = "@Date"
        sparam0.SqlDbType = SqlDbType.Char
        sparam0.Size = 8
        sparam0.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam0)

        Dim sparam1 As SqlClient.SqlParameter
        sparam1 = New SqlClient.SqlParameter
        sparam1.ParameterName = "@Time"
        sparam1.SqlDbType = SqlDbType.Char
        sparam1.Size = 8
        sparam1.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam1)

        Dim sparam2 As SqlClient.SqlParameter
        sparam2 = New SqlClient.SqlParameter
        sparam2.ParameterName = "@errID"
        sparam2.SqlDbType = SqlDbType.Char
        sparam2.Size = 4
        sparam2.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam2)

        Dim sparam3 As SqlClient.SqlParameter
        sparam3 = New SqlClient.SqlParameter
        sparam3.ParameterName = "@errMsg"
        sparam3.SqlDbType = SqlDbType.Char
        sparam3.Size = 40
        sparam3.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam3)

        Dim Reader As SqlClient.SqlDataReader
        cmd2.Parameters("@Date").Value = dateup.Trim
        cmd2.Parameters("@Time").Value = Timeup.Trim

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            Chafer = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            Chafer = False
        End Try
        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Function
    Private Function Chafer2(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Chafer2 = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CALCHAFER2"
        cmd2.Connection = cnn

        Dim sparam0 As SqlClient.SqlParameter
        sparam0 = New SqlClient.SqlParameter
        sparam0.ParameterName = "@Date"
        sparam0.SqlDbType = SqlDbType.Char
        sparam0.Size = 8
        sparam0.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam0)

        Dim sparam1 As SqlClient.SqlParameter
        sparam1 = New SqlClient.SqlParameter
        sparam1.ParameterName = "@Time"
        sparam1.SqlDbType = SqlDbType.Char
        sparam1.Size = 8
        sparam1.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam1)

        Dim sparam2 As SqlClient.SqlParameter
        sparam2 = New SqlClient.SqlParameter
        sparam2.ParameterName = "@errID"
        sparam2.SqlDbType = SqlDbType.Char
        sparam2.Size = 4
        sparam2.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam2)

        Dim sparam3 As SqlClient.SqlParameter
        sparam3 = New SqlClient.SqlParameter
        sparam3.ParameterName = "@errMsg"
        sparam3.SqlDbType = SqlDbType.Char
        sparam3.Size = 40
        sparam3.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam3)

        Dim Reader As SqlClient.SqlDataReader
        cmd2.Parameters("@Date").Value = dateup.Trim
        cmd2.Parameters("@Time").Value = Timeup.Trim

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            Chafer2 = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            Chafer2 = False
        End Try
        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Function
#End Region
#End Region

#Region "CALGT Price"
    Private Function GT(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        GT = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CALGT"
        cmd2.Connection = cnn

        Dim sparam0 As SqlClient.SqlParameter
        sparam0 = New SqlClient.SqlParameter
        sparam0.ParameterName = "@Date"
        sparam0.SqlDbType = SqlDbType.Char
        sparam0.Size = 8
        sparam0.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam0)

        Dim sparam1 As SqlClient.SqlParameter
        sparam1 = New SqlClient.SqlParameter
        sparam1.ParameterName = "@Time"
        sparam1.SqlDbType = SqlDbType.Char
        sparam1.Size = 8
        sparam1.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam1)

        Dim sparam2 As SqlClient.SqlParameter
        sparam2 = New SqlClient.SqlParameter
        sparam2.ParameterName = "@errID"
        sparam2.SqlDbType = SqlDbType.Char
        sparam2.Size = 4
        sparam2.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam2)

        Dim sparam3 As SqlClient.SqlParameter
        sparam3 = New SqlClient.SqlParameter
        sparam3.ParameterName = "@errMsg"
        sparam3.SqlDbType = SqlDbType.Char
        sparam3.Size = 40
        sparam3.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam3)

        Dim Reader As SqlClient.SqlDataReader
        cmd2.Parameters("@Date").Value = dateup.Trim
        cmd2.Parameters("@Time").Value = Timeup.Trim

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            GT = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            GT = False
        End Try
        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Function
#End Region

#End Region

    Sub DelPrice(ByVal strType As String)
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Try
            strSQL = "   Delete TBLMasterPrice"
            strSQL += ""
            strSQL += "   Delete TBLMasterPriceRM"

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
End Class
