Public Class frmPvRptTag
    Inherits System.Windows.Forms.Form

    Dim sOutAs As String
    Friend sDesc As String
    Friend sNo As String
    Friend sDt As String

    Friend dt2 As DataTable
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
    Friend WithEvents V2 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents C1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.C1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'C1
        '
        Me.C1.ActiveViewIndex = -1
        Me.C1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.C1.Location = New System.Drawing.Point(0, 0)
        Me.C1.Name = "C1"
        Me.C1.ReportSource = Nothing
        Me.C1.Size = New System.Drawing.Size(680, 438)
        Me.C1.TabIndex = 0
        '
        'frmPvRptTag
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(680, 438)
        Me.Controls.Add(Me.C1)
        Me.Name = "frmPvRptTag"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "พิมพ์ใบนับ"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim ds2 As New DataSet("Dataset1")

    Private Sub frmPrintOut_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Me.Text += " " & sDt.Trim
        'dt2.TableName = "TBL_RM"
        'ds2.Tables.Add(dt2)
        'sOutAs = " " '& sDt.Trim & " เลขที่ " & sNo.Trim & "  " & sDesc.Trim
        '   Dim r1 As New CrystalReport1
        '  r1.SetDataSource(ds2)
        '        r1.SetParameterValue("sHeader", sOutAs)
        '   r1.SetParameterValue("sNo", sNo)
        '   r1.SetParameterValue("sDt", sDt)
        ' r1.SetParameterValue("sFooter", "หมายเหตุ : " & sDesc)

        'C1.ReportSource = r1
    End Sub

    Private Sub frmPrintOut_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        ds2.Tables.Remove("retdoc")
        ds2.Dispose()
    End Sub

End Class
