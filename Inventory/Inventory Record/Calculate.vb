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
    Dim StrDate, StrTime As String
    Dim StrType As String

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
    Friend WithEvents CmdCalculate As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CKPigment As System.Windows.Forms.CheckBox
    Friend WithEvents CKRM As System.Windows.Forms.CheckBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents CKCompound As System.Windows.Forms.CheckBox
    Friend WithEvents CKPresemi As System.Windows.Forms.CheckBox
    Friend WithEvents CKSemi As System.Windows.Forms.CheckBox
    Friend WithEvents CKGT As System.Windows.Forms.CheckBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Calculate))
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.CmdCalculate = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.CKGT = New System.Windows.Forms.CheckBox()
        Me.CKSemi = New System.Windows.Forms.CheckBox()
        Me.CKPresemi = New System.Windows.Forms.CheckBox()
        Me.CKCompound = New System.Windows.Forms.CheckBox()
        Me.CKRM = New System.Windows.Forms.CheckBox()
        Me.CKPigment = New System.Windows.Forms.CheckBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
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
        'CmdCalculate
        '
        Me.CmdCalculate.Location = New System.Drawing.Point(112, 136)
        Me.CmdCalculate.Name = "CmdCalculate"
        Me.CmdCalculate.Size = New System.Drawing.Size(75, 23)
        Me.CmdCalculate.TabIndex = 1
        Me.CmdCalculate.Text = "Calculate"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CKGT)
        Me.GroupBox1.Controls.Add(Me.CKSemi)
        Me.GroupBox1.Controls.Add(Me.CKPresemi)
        Me.GroupBox1.Controls.Add(Me.CKCompound)
        Me.GroupBox1.Controls.Add(Me.CKRM)
        Me.GroupBox1.Controls.Add(Me.CKPigment)
        Me.GroupBox1.Controls.Add(Me.CmdCalculate)
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
        Me.CKGT.Size = New System.Drawing.Size(104, 24)
        Me.CKGT.TabIndex = 5
        Me.CKGT.Text = "Green Tire"
        '
        'CKSemi
        '
        Me.CKSemi.Checked = True
        Me.CKSemi.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CKSemi.Location = New System.Drawing.Point(176, 56)
        Me.CKSemi.Name = "CKSemi"
        Me.CKSemi.Size = New System.Drawing.Size(104, 24)
        Me.CKSemi.TabIndex = 4
        Me.CKSemi.Text = "Semi"
        '
        'CKPresemi
        '
        Me.CKPresemi.Checked = True
        Me.CKPresemi.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CKPresemi.Location = New System.Drawing.Point(176, 24)
        Me.CKPresemi.Name = "CKPresemi"
        Me.CKPresemi.Size = New System.Drawing.Size(104, 24)
        Me.CKPresemi.TabIndex = 3
        Me.CKPresemi.Text = "Presemi"
        '
        'CKCompound
        '
        Me.CKCompound.Checked = True
        Me.CKCompound.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CKCompound.Location = New System.Drawing.Point(32, 88)
        Me.CKCompound.Name = "CKCompound"
        Me.CKCompound.Size = New System.Drawing.Size(104, 24)
        Me.CKCompound.TabIndex = 2
        Me.CKCompound.Text = "Compound"
        '
        'CKRM
        '
        Me.CKRM.Checked = True
        Me.CKRM.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CKRM.Location = New System.Drawing.Point(32, 24)
        Me.CKRM.Name = "CKRM"
        Me.CKRM.Size = New System.Drawing.Size(104, 24)
        Me.CKRM.TabIndex = 1
        Me.CKRM.Text = "R/M (Material)"
        '
        'CKPigment
        '
        Me.CKPigment.Checked = True
        Me.CKPigment.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CKPigment.Location = New System.Drawing.Point(32, 56)
        Me.CKPigment.Name = "CKPigment"
        Me.CKPigment.Size = New System.Drawing.Size(104, 24)
        Me.CKPigment.TabIndex = 0
        Me.CKPigment.Text = "Pigment"
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        Me.BackgroundWorker1.WorkerSupportsCancellation = True
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
        Me.MaximizeBox = False
        Me.Name = "Calculate"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Calculate"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Form Event"
    Private Sub Calculate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Get datetime
        StrDate = DateTime.Now.ToString("yyyyMMdd", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))
        StrTime = DateTime.Now.ToString("HHmm", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))
    End Sub
#End Region

#Region "Control Event"
    Private Sub CmdCalculate_Click(sender As Object, e As EventArgs) Handles CmdCalculate.Click
        If CKRM.Checked And CKPigment.Checked And CKCompound.Checked And CKPresemi.Checked And CKSemi.Checked And CKGT.Checked Then
            'If if is selected all
            DelPrice() 'Delete data in Table TBLMasterPrice and TBLMasterPriceRM
        End If

        'Start to process
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

        'Call method DoWork
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        CmdCalculate.Enabled = False
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        e.Result = Progress()
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        'Update step
        If e.UserState Then
            'PerformStep
            ProgressBar1.PerformStep()
        Else
            'Skip
            ProgressBar1.Value = e.ProgressPercentage
        End If
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        CmdCalculate.Enabled = True
        Me.Cursor = System.Windows.Forms.Cursors.Default

        If e.Result Then
            MessageBox.Show(" Calculate Complete. ", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Close()
        Else
            MessageBox.Show(" Calculate Not Complete. Please Check Data. Calculate by Process ", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub
#End Region

#Region "Call Store"

#Region "CalRM Price"
    Private Function CalRM(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Dim ret As Boolean = False
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
        cmd2.Parameters("@Date").Value = dateup.Trim()
        cmd2.Parameters("@Time").Value = Timeup.Trim()

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            ret = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            ret = False
        End Try

        cnn.Close()
        Return ret
    End Function
#End Region

#Region "CalPigment Price"
    Private Function CalPigment(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Dim ret As Boolean = False
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
        cmd2.Parameters("@Date").Value = dateup.Trim()
        cmd2.Parameters("@Time").Value = Timeup.Trim()

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            ret = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            ret = False
        End Try

        cnn.Close()
        Return ret
    End Function
#End Region

#Region "CalCompound Price"
    Private Function CalCompound(ByVal dateup As String, ByVal timeup As String) As Boolean
        Dim ret As Boolean = False
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
        cmd2.Parameters("@Date").Value = dateup.Trim()
        cmd2.Parameters("@Time").Value = timeup.Trim()

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            ret = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            ret = False
        End Try

        cnn.Close()
        Return ret
    End Function
    Private Function CalCompound2(ByVal dateup As String, ByVal timeup As String) As Boolean
        Dim ret As Boolean = False
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
        cmd2.Parameters("@Date").Value = dateup.Trim()
        cmd2.Parameters("@Time").Value = timeup.Trim()

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            ret = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            ret = False
        End Try

        cnn.Close()
        Return ret
    End Function
#End Region

#Region "CalCoatedcord (PreSemi Price)"
    Private Function CalCoatedcord(ByVal dateup As String, ByVal timeup As String) As Boolean
        Dim ret As Boolean = False
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
        cmd2.Parameters("@Date").Value = dateup.Trim()
        cmd2.Parameters("@Time").Value = timeup.Trim()

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            ret = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            ret = False
        End Try

        cnn.Close()
        Return ret
    End Function
    Private Function CalCoatedcord2(ByVal dateup As String, ByVal timeup As String) As Boolean
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
        cmd2.Parameters("@Date").Value = dateup.Trim()
        cmd2.Parameters("@Time").Value = timeup.Trim()

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
        Catch ex As Exception
            MsgBox(ex.Message, 48)
        End Try

        cnn.Close()
        Return True
    End Function
    Private Function CalCoatedcord3(ByVal dateup As String, ByVal timeup As String) As Boolean
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
        cmd2.Parameters("@Date").Value = dateup.Trim()
        cmd2.Parameters("@Time").Value = timeup.Trim()

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
        Catch ex As Exception
            MsgBox(ex.Message, 48)
        End Try

        cnn.Close()
        Return True
    End Function
#End Region

#Region "CalPreSemi Price"
    Private Function PreSemi(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Dim ret As Boolean = False
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
        cmd2.Parameters("@Date").Value = dateup.Trim()
        cmd2.Parameters("@Time").Value = Timeup.Trim()

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            ret = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            ret = False
        End Try

        cnn.Close()
        Return ret
    End Function
#End Region

#Region "CalSteelCord (PreSemi Price)"
    Private Function SteelCord(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Dim ret As Boolean = False
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
        cmd2.Parameters("@Date").Value = dateup.Trim()
        cmd2.Parameters("@Time").Value = Timeup.Trim()

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            ret = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            ret = False
        End Try

        cnn.Close()
        Return ret
    End Function
#End Region

#Region "CalSemi Price"
#Region "TREAD Semi Price"
    Private Function Tread(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Dim ret As Boolean = False
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
        cmd2.Parameters("@Date").Value = dateup.Trim()
        cmd2.Parameters("@Time").Value = Timeup.Trim()

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            ret = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            ret = False
        End Try

        cnn.Close()
        Return ret
    End Function
#End Region

#Region "BF Semi Price"
    Private Function BF(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Dim ret As Boolean = False
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
        cmd2.Parameters("@Date").Value = dateup.Trim()
        cmd2.Parameters("@Time").Value = Timeup.Trim()

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            ret = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            ret = False
        End Try

        cnn.Close()
        Return ret
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

#Region "BELT 1-4 Semi Price"
    Private Function BElT(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Dim ret As Boolean = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CALBELT"
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
        cmd2.Parameters("@Date").Value = dateup.Trim()
        cmd2.Parameters("@Time").Value = Timeup.Trim()

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            ret = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            ret = False
        End Try

        cnn.Close()
        Return ret
    End Function
    Private Function BElT2(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Dim ret As Boolean = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CALBELT2"
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
        cmd2.Parameters("@Date").Value = dateup.Trim()
        cmd2.Parameters("@Time").Value = Timeup.Trim()

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            ret = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            ret = False
        End Try

        cnn.Close()
        Return ret
    End Function
#End Region

#Region "SIDE CUSSION INNERLINER Semi Price"
    Private Function SIC(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Dim ret As Boolean = False
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
        cmd2.Parameters("@Date").Value = dateup.Trim()
        cmd2.Parameters("@Time").Value = Timeup.Trim()

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            ret = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            ret = False
        End Try

        cnn.Close()
        Return ret
    End Function
#End Region

#Region "WIRE CHAFER,Nylon CHAFER,BODY PLY Semi Price"
    Private Function Chafer(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Dim ret As Boolean = False
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
        cmd2.Parameters("@Date").Value = dateup.Trim()
        cmd2.Parameters("@Time").Value = Timeup.Trim()

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            ret = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            ret = False
        End Try

        cnn.Close()
        Return ret
    End Function
    Private Function Chafer2(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Dim ret As Boolean = False
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
            ret = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            ret = False
        End Try

        cnn.Close()
        Return ret
    End Function
#End Region
#End Region 'CalSemi Price

#Region "CALGT Price"
    Private Function GT(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Dim ret As Boolean = False
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
        cmd2.Parameters("@Date").Value = dateup.Trim()
        cmd2.Parameters("@Time").Value = Timeup.Trim()

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            ret = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            ret = False
        End Try

        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()
        Return ret
    End Function
#End Region

#End Region 'Call Store

#Region "Function"
    Sub CKdata()
        If CKRM.Checked Then
            StrType &= "'01'"
        End If
        If CKPigment.Checked Then
            StrType &= "'02'"
        End If
        If CKCompound.Checked Then
            StrType &= "'03'"
        End If
        If CKPresemi.Checked Then
            StrType &= "'04'"
        End If
        If CKSemi.Checked Then
            StrType &= "'05'"
        End If
        If CKGT.Checked Then
            StrType &= "'06'"
        End If
    End Sub

    Function Progress(ByVal ParamArray filenames As String()) As Boolean
        Dim ret As Boolean = False

        ' Loop through all files to copy.
        For x As Integer = 1 To 100
            UpdateProgress(x, False)

            If x = 1 Then
                'R/M Material
                If CKRM.Checked Then
                    CalRM(StrDate, StrTime) 'Call Store Procedure CALRM
                Else
                    x = 10 'Skip to next process
                    UpdateProgress(x, False)
                End If
            End If

            If x = 10 Then
                'Pigment
                If CKPigment.Checked Then
                    CalPigment(StrDate, StrTime) 'Call Store Procedure CALPigment
                Else
                    x = 20 'Skip to next process
                    UpdateProgress(x, False)
                End If
            End If

            If x = 20 Then
                'Compound
                If CKCompound.Checked Then
                    CalCompound(StrDate, StrTime) 'Call Store Procedure CALCompound (Insert Table TBLMasterPriceRM)
                    CalCompound2(StrDate, StrTime) 'Call Store Procedure CALCompound2 (Insert Table TBLMasterPrice)
                Else
                    x = 30 'Skip to next process
                    UpdateProgress(x, False)
                End If
            End If

            If x = 30 Then
                'Pre Semi
                If CKPresemi.Checked Then
                    CalCoatedcord2(StrDate, StrTime) 'Call Store Procedure CALCoatedCord2 (Insert Table TBLMasterPriceRM Exclude TBLRM)
                    CalCoatedcord3(StrDate, StrTime) 'Call Store Procedure CALCoatedCord3 (Insert Table TBLMasterPriceRM Include TBLRM)
                    CalCoatedcord(StrDate, StrTime) 'Call Store Procedure CALCoatedCord (Insert Table TBLMasterPrice)
                Else
                    x = 40 'Skip to next process
                    UpdateProgress(x, False)
                End If
            End If

            If x = 40 Then
                'Pre Semi
                If CKPresemi.Checked Then
                    PreSemi(StrDate, StrTime) 'Call Store Procedure CALPreSemi (Material Type is not STEEL CORD and COATED CORD)
                    SteelCord(StrDate, StrTime) 'Call Store Procedure CALSteelCord (Material Type is STEEL CORD)
                Else
                    x = 50 'Skip to next process
                    UpdateProgress(x, False)
                End If
            End If

            If x = 50 Then
                'Semi
                If CKSemi.Checked Then
                    Tread(StrDate, StrTime) 'Call Store Procedure CALTT (Material Type is TREAD)
                    BF(StrDate, StrTime) 'Call Store Procedure CALBF (Material Type is BF)(Insert Table TBLMasterPrice)
                    BF2(StrDate, StrTime) 'Call Store Procedure CALBF2 (Material Type is BF)(Insert Table TBLMasterPriceRM)
                Else
                    x = 70 'Skip to next process
                    UpdateProgress(x, False)
                End If
            End If

            If x = 70 Then
                'Semi
                If CKSemi.Checked Then
                    BElT(StrDate, StrTime) 'Call Store Procedure CALBELT (Type Material:BELT-1, BELT-2, BELT-3, BELT-4 to Table TBLMASTERPRICE)
                    BElT2(StrDate, StrTime) 'Call Store Procedure CALBELT2 (Type Material:BELT-1, BELT-2, BELT-3, BELT-4 to Table TBLMASTERPRICERM)
                    SIC(StrDate, StrTime) 'Call Store Procedure CALSIC (Type Material:CUSSION, SIDE, INNERLINER)
                    Chafer(StrDate, StrTime) 'Call Store Procedure (Type Material:BODY PLY, WIRE CHAFER, Nylon CHAFER to Table TBLMASTERPRICE)
                    Chafer2(StrDate, StrTime) 'Call Store Procedure (Type Material:BODY PLY, WIRE CHAFER, Nylon CHAFER to Table TBLMASTERPRICERM)
                    'Call Store Procedure (Type Material:Flipper)
                Else
                    x = 90 'Skip to next process
                    UpdateProgress(x, False)
                End If
            End If

            If x = 90 Then
                'Green Tire
                If CKGT.Checked Then
                    GT(StrDate, StrTime) 'Call Store Procedure CALGT
                Else
                    x = 100
                    UpdateProgress(x, False)
                End If
            End If

            If x = 100 Then
                ret = True
            Else
                ret = False
            End If
        Next x

        Return ret
    End Function

    Protected Sub UpdateProgress(x As Integer, isPerformStep As Boolean)
        BackgroundWorker1.ReportProgress(x, isPerformStep)
    End Sub

    Sub DelPrice()
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim strSQL As String = String.Empty
        Dim sb As New System.Text.StringBuilder()

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Try
            sb.AppendLine("   Delete TBLMasterPrice")
            sb.AppendLine("   Delete TBLMasterPriceRM")
            strSQL = sb.ToString()

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
End Class
