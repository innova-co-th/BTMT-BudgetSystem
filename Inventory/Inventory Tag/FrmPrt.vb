#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Inventory_Tag.Common
#End Region

Public Class FrmPrt
    Inherits System.Windows.Forms.Form
    Friend sDesc As String
    Friend sUser As String
    Friend sDt As String
    Friend sOutAs As String
    Friend dt_new As DataTable
 
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
    Friend WithEvents Viewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmPrt))
        Me.Viewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'Viewer1
        '
        Me.Viewer1.ActiveViewIndex = -1
        Me.Viewer1.DisplayGroupTree = False
        Me.Viewer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Viewer1.Location = New System.Drawing.Point(0, 0)
        Me.Viewer1.Name = "Viewer1"
        Me.Viewer1.ReportSource = Nothing
        Me.Viewer1.Size = New System.Drawing.Size(720, 478)
        Me.Viewer1.TabIndex = 0
        '
        'FrmPrt
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(720, 478)
        Me.Controls.Add(Me.Viewer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmPrt"
        Me.Text = "Print Preview"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim ds2 As New DataSet("Dataset1")

    Private Sub FrmPrt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dt_new.TableName = "TBL_TAG"
        ds2.Tables.Add(dt_new)
        sOutAs = " Physical Check Work in Process by Tag."
        sDt = Now.Date
        sDesc = ""
        Dim r1 As New Report1
        r1.SetDataSource(ds2)
        r1.SetParameterValue("sHeader", sOutAs)
        r1.SetParameterValue("sUser", sUser)
        'r1.SetParameterValue("sFooter", "หมายเหตุ : " & sDesc)
        Viewer1.ReportSource = r1
    End Sub

    Private Sub frmPrintOut_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        ds2.Tables.Remove("TBL_TAG")
        ds2.Dispose()
    End Sub

End Class
