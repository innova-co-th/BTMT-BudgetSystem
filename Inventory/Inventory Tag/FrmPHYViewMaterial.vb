#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Inventory_Tag.Common
#End Region

Public Class FrmPHYViewMaterial
#Region "Declare"
    Inherits System.Windows.Forms.Form
    Dim GrdDV As New DataView
    Protected Const TBL_RM As String = "TBL_RM"
    Public Shared tb1 As New DataTable
    Protected DefaultGridBorderStyle As BorderStyle
    Dim C1 As New SQLData("ACCINV")
    Dim StrData As String
    Friend sDesc As String
    Friend sAdj As String
    Friend sUser As String
    Friend sDt As String
    Friend sOutAs As String
    Friend sCODE, sName, sSec, sHeader, sMonth, sIdSec As String

    Friend dt_new As DataTable
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
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmPHYViewMaterial))
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = -1
        'Me.CrystalReportViewer1.DisplayGroupTree = False 'It is obsolete
        Me.CrystalReportViewer1.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        Me.CrystalReportViewer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrystalReportViewer1.Location = New System.Drawing.Point(0, 0)
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.ReportSource = Nothing
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(744, 334)
        Me.CrystalReportViewer1.TabIndex = 0
        '
        'FrmPHYViewMaterial
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(744, 334)
        Me.Controls.Add(Me.CrystalReportViewer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmPHYViewMaterial"
        Me.Text = "View Physical Report By Material"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim ds2 As New DataSet("Dataset3")


    Private Sub FrmPHYViewMaterial_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dt_new.TableName = "TBL_Trx"
        ds2.Tables.Add(dt_new)
        sOutAs = " Physical Check Work in Process by Tag."
        sDt = Now.Date
        sDesc = ""
        Dim r1 As New PHYReportMaterial
        r1.SetDataSource(ds2)
       r1.SetParameterValue("strCODE", sCODE)
        r1.SetParameterValue("sName", sName)
        r1.SetParameterValue("sSection", sSec)
        r1.SetParameterValue("sIdsec", sIdSec)
        r1.SetParameterValue("sHeader", sHeader)
        r1.SetParameterValue("sMonth", sMonth)
        CrystalReportViewer1.ReportSource = r1
    End Sub

    Private Sub FrmPHYViewMaterial_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        ds2.Tables.Remove("TBL_Trx")
        ds2.Dispose()
    End Sub
End Class
