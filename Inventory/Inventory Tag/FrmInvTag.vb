#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Inventory_Tag.Common
Imports Login
#End Region

Public Structure CellColor
    Public ForeG As Integer
    Public BackG As Integer
    Public LfItem As String
End Structure 'CellColor

Public Class FrmInvTag
    Inherits System.Windows.Forms.Form
#Region "Varialble"
    Public Shared HasColor As Boolean
    Public Shared HasEdited As Boolean
    Public Shared HasAut2Edit As Boolean
    Public Shared HasCommited As Boolean
    Public Shared CurrentIDUser As String = String.Empty 'Variable for EmpCode
    Public Shared CurrentName As String = String.Empty 'Varaible for PersonFNameEng and PersonLNameEng
    Public Shared CurrentLevel As String = String.Empty 'Variable for LevelUsage
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
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuHorizontal As System.Windows.Forms.MenuItem
    Friend WithEvents MenuVertical As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCasCade As System.Windows.Forms.MenuItem
    Friend WithEvents MenuArrange As System.Windows.Forms.MenuItem
    Friend WithEvents MenuClose As System.Windows.Forms.MenuItem
    Friend WithEvents MenuInvTag As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuAbout As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuReport As System.Windows.Forms.MenuItem
    Friend WithEvents MonthView As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmInvTag))
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.MenuClose = New System.Windows.Forms.MenuItem()
        Me.MenuItem2 = New System.Windows.Forms.MenuItem()
        Me.MenuInvTag = New System.Windows.Forms.MenuItem()
        Me.MenuItem9 = New System.Windows.Forms.MenuItem()
        Me.MenuReport = New System.Windows.Forms.MenuItem()
        Me.MenuItem5 = New System.Windows.Forms.MenuItem()
        Me.MonthView = New System.Windows.Forms.MenuItem()
        Me.MenuItem10 = New System.Windows.Forms.MenuItem()
        Me.MenuHorizontal = New System.Windows.Forms.MenuItem()
        Me.MenuVertical = New System.Windows.Forms.MenuItem()
        Me.MenuCasCade = New System.Windows.Forms.MenuItem()
        Me.MenuArrange = New System.Windows.Forms.MenuItem()
        Me.MenuItem11 = New System.Windows.Forms.MenuItem()
        Me.MenuItem3 = New System.Windows.Forms.MenuItem()
        Me.MenuAbout = New System.Windows.Forms.MenuItem()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2, Me.MenuItem9, Me.MenuItem10, Me.MenuItem11})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuClose})
        Me.MenuItem1.Text = "File"
        '
        'MenuClose
        '
        Me.MenuClose.Index = 0
        Me.MenuClose.Shortcut = System.Windows.Forms.Shortcut.AltF4
        Me.MenuClose.Text = "Close"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuInvTag})
        Me.MenuItem2.Text = "Record Tag"
        '
        'MenuInvTag
        '
        Me.MenuInvTag.Index = 0
        Me.MenuInvTag.Text = " Inventory Tag"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 2
        Me.MenuItem9.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuReport, Me.MenuItem5, Me.MonthView})
        Me.MenuItem9.Text = "Report "
        '
        'MenuReport
        '
        Me.MenuReport.Index = 0
        Me.MenuReport.Text = "Physical  Report"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 1
        Me.MenuItem5.Text = "-"
        '
        'MonthView
        '
        Me.MonthView.Index = 2
        Me.MonthView.Text = "Scarp  Report"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 3
        Me.MenuItem10.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuHorizontal, Me.MenuVertical, Me.MenuCasCade, Me.MenuArrange})
        Me.MenuItem10.Text = "View"
        '
        'MenuHorizontal
        '
        Me.MenuHorizontal.Index = 0
        Me.MenuHorizontal.Text = "Tile Horizontal"
        '
        'MenuVertical
        '
        Me.MenuVertical.Index = 1
        Me.MenuVertical.Text = "Tile Vertical"
        '
        'MenuCasCade
        '
        Me.MenuCasCade.Index = 2
        Me.MenuCasCade.Text = "CasCade"
        '
        'MenuArrange
        '
        Me.MenuArrange.Index = 3
        Me.MenuArrange.Text = "Arrange Icons"
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 4
        Me.MenuItem11.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem3, Me.MenuAbout})
        Me.MenuItem11.Text = "Help"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 0
        Me.MenuItem3.Text = "-"
        '
        'MenuAbout
        '
        Me.MenuAbout.Index = 1
        Me.MenuAbout.Text = "About Inventory Tag"
        '
        'FrmInvTag
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(792, 449)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.Menu = Me.MainMenu1
        Me.Name = "FrmInvTag"
        Me.Text = "Inventory Tag"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub MenuHorizontal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuHorizontal.Click
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub MenuVertical_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuVertical.Click
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub MenuCasCade_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCasCade.Click
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub MenuArrange_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuArrange.Click
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub MenuClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuClose.Click
        Me.Close()
    End Sub

    Private Sub FrmInvTag_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim flog As New FrmLogin
        flog.ShowDialog()
        If flog.EmpID <> String.Empty Then
            Me.WindowState = FormWindowState.Maximized
            CurrentIDUser = Login.FrmLogin.EmpID
            CurrentName = Login.FrmLogin.EmpName
            CurrentLevel = Login.FrmLogin.LevelUsage

            Dim fYTag As New FrmYearInvTag
            fYTag.MdiParent = Me
            fYTag.Username = CurrentName
            fYTag.Show()
            Me.LayoutMdi(MdiLayout.TileHorizontal)
        Else
            If closeprogram() Then
                Me.Close()
            End If
        End If
    End Sub
    Function closeprogram() As Boolean
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        Dim flog As New FrmLogin
        closeprogram = False
        msg = "Inventory Record Closing Program"  ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Inventory"   ' Define title.
        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            closeprogram = True
            Me.Close()
        Else
            flog.ShowDialog()
            If flog.EmpID <> "" Then
                Me.WindowState = FormWindowState.Maximized
                CurrentIDUser = Login.FrmLogin.EmpID
                CurrentName = Login.FrmLogin.EmpName
                CurrentLevel = Login.FrmLogin.LevelUsage
            Else
                If closeprogram() Then
                    Me.Close()
                End If
            End If
        End If
    End Function
    Private Sub MenuInvTag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuInvTag.Click
        Dim fYTag As New FrmYearInvTag
        fYTag.MdiParent = Me
        fYTag.Username = CurrentName
        fYTag.Show()
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub MenuReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuReport.Click
        Dim fv As New FrmView
        fv.ShowDialog()
    End Sub


    Private Sub MonthView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MonthView.Click
        Dim fmv As New FrmMonthView
        fmv.ShowDialog()
    End Sub

    Private Sub MenuAbout_Click(sender As Object, e As EventArgs) Handles MenuAbout.Click
        Dim titleAttr As System.Reflection.AssemblyTitleAttribute = System.Reflection.Assembly.GetExecutingAssembly().GetCustomAttributes(GetType(System.Reflection.AssemblyTitleAttribute), False)(0)
        Dim version As String = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()
        MsgBox(titleAttr.Title & " " & version, MsgBoxStyle.Information, titleAttr.Title)
    End Sub
End Class
