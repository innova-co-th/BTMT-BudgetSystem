#Region "Import"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.Drawing.Image
Imports System.Drawing.Printing
Imports Inventory_Record.Common
Imports Inventory_Tag
Imports Login
#End Region

Public Structure CellColor
    Public ForeG As Integer
    Public BackG As Integer
    Public LfItem As String
End Structure 'CellColor

Public Class FrmMain
    Inherits System.Windows.Forms.Form
#Region "Variable"
    Public Shared HasColor As Boolean
    Public Shared HasEdited As Boolean
    Public Shared HasAut2Edit As Boolean
    Public Shared HasCommited As Boolean
    Dim C1 As New SQLData() 'Connect Database
    Friend CurrentIDUser As String = String.Empty 'Variable for EmpCode
    Friend CurrentName As String = String.Empty 'Variable for PersonFNameEng and PersonLNameEng
    Friend CurrentLevel As String = String.Empty 'Variable for LevelUsage

    Public Shared GrdDV As New DataView
    Protected Const TBL_RM As String = "TBL_RM"
    Protected DefaultGridBorderStyle As BorderStyle
    Public Shared cm As CurrencyManager
    Dim iTotal As Double
    Dim DT As New DataTable
    Dim StrSql As String = String.Empty
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
    Friend WithEvents MenuFile As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuRM As System.Windows.Forms.MenuItem
    Friend WithEvents MenuSemi As System.Windows.Forms.MenuItem
    Friend WithEvents MenuDept As System.Windows.Forms.MenuItem
    Friend WithEvents MenuHelp As System.Windows.Forms.MenuItem
    Friend WithEvents MenuAbout As System.Windows.Forms.MenuItem
    Friend WithEvents MenuUpdate As System.Windows.Forms.MenuItem
    Friend WithEvents MenuUnit As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuRMPrice As System.Windows.Forms.MenuItem
    Friend WithEvents MenuRMQty As System.Windows.Forms.MenuItem
    Friend WithEvents MenuView As System.Windows.Forms.MenuItem
    Friend WithEvents MenuHorizontal As System.Windows.Forms.MenuItem
    Friend WithEvents MenuVertical As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCasCade As System.Windows.Forms.MenuItem
    Friend WithEvents MenuArrange As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem12 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCompound As System.Windows.Forms.MenuItem
    Friend WithEvents MenuTireCode As System.Windows.Forms.MenuItem
    Friend WithEvents MenuType As System.Windows.Forms.MenuItem
    Friend WithEvents MenuPIGMENT As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuConvertUnit As System.Windows.Forms.MenuItem
    Friend WithEvents MenuPreSemi As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuTypeMaterial As System.Windows.Forms.MenuItem
    Friend WithEvents MenuInv As System.Windows.Forms.MenuItem
    Friend WithEvents MenuClose As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCAL As System.Windows.Forms.MenuItem
    Friend WithEvents MenuPricePigment As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCALPigment As System.Windows.Forms.MenuItem
    Friend WithEvents MenuPriceCompound As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCALCompound As System.Windows.Forms.MenuItem
    Friend WithEvents MenuPriceRM As System.Windows.Forms.MenuItem
    Friend WithEvents MenuPricePreSemi As System.Windows.Forms.MenuItem
    Friend WithEvents MenuPriceSemi As System.Windows.Forms.MenuItem
    Friend WithEvents MenuPriceTire As System.Windows.Forms.MenuItem
    Friend WithEvents MenuConvert As System.Windows.Forms.MenuItem
    Friend WithEvents MenuPrice As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem32 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuRecord As System.Windows.Forms.MenuItem
    Friend WithEvents MenuAddUser As System.Windows.Forms.MenuItem
    Friend WithEvents MenuMaster As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCalculate As System.Windows.Forms.MenuItem
    Friend WithEvents MenuReport As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem13 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem31 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuRHC As System.Windows.Forms.MenuItem
    Friend WithEvents MenuPer As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem14 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuPCOM As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCalRM As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCalCompoundStage1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCalCompoundStage2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCALPreSemi As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCalCoatedCord As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCALSteel As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCALALLPreSemi As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCalSemi As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCalTT As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCalBF As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCalBelt As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCalSIC As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCalBW As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCalNF As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCalGT As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem16 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmMain))
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.MenuFile = New System.Windows.Forms.MenuItem()
        Me.MenuClose = New System.Windows.Forms.MenuItem()
        Me.MenuMaster = New System.Windows.Forms.MenuItem()
        Me.MenuItem7 = New System.Windows.Forms.MenuItem()
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.MenuRM = New System.Windows.Forms.MenuItem()
        Me.MenuItem12 = New System.Windows.Forms.MenuItem()
        Me.MenuRMPrice = New System.Windows.Forms.MenuItem()
        Me.MenuRMQty = New System.Windows.Forms.MenuItem()
        Me.MenuPIGMENT = New System.Windows.Forms.MenuItem()
        Me.MenuItem31 = New System.Windows.Forms.MenuItem()
        Me.MenuCompound = New System.Windows.Forms.MenuItem()
        Me.MenuRHC = New System.Windows.Forms.MenuItem()
        Me.MenuPer = New System.Windows.Forms.MenuItem()
        Me.MenuPreSemi = New System.Windows.Forms.MenuItem()
        Me.MenuSemi = New System.Windows.Forms.MenuItem()
        Me.MenuTireCode = New System.Windows.Forms.MenuItem()
        Me.MenuItem8 = New System.Windows.Forms.MenuItem()
        Me.MenuUnit = New System.Windows.Forms.MenuItem()
        Me.MenuItem9 = New System.Windows.Forms.MenuItem()
        Me.MenuConvertUnit = New System.Windows.Forms.MenuItem()
        Me.MenuItem10 = New System.Windows.Forms.MenuItem()
        Me.MenuType = New System.Windows.Forms.MenuItem()
        Me.MenuItem11 = New System.Windows.Forms.MenuItem()
        Me.MenuTypeMaterial = New System.Windows.Forms.MenuItem()
        Me.MenuItem6 = New System.Windows.Forms.MenuItem()
        Me.MenuDept = New System.Windows.Forms.MenuItem()
        Me.MenuItem32 = New System.Windows.Forms.MenuItem()
        Me.MenuAddUser = New System.Windows.Forms.MenuItem()
        Me.MenuCalculate = New System.Windows.Forms.MenuItem()
        Me.MenuItem5 = New System.Windows.Forms.MenuItem()
        Me.MenuItem16 = New System.Windows.Forms.MenuItem()
        Me.MenuCAL = New System.Windows.Forms.MenuItem()
        Me.MenuCalRM = New System.Windows.Forms.MenuItem()
        Me.MenuCALPigment = New System.Windows.Forms.MenuItem()
        Me.MenuCALCompound = New System.Windows.Forms.MenuItem()
        Me.MenuCalCompoundStage1 = New System.Windows.Forms.MenuItem()
        Me.MenuCalCompoundStage2 = New System.Windows.Forms.MenuItem()
        Me.MenuCALPreSemi = New System.Windows.Forms.MenuItem()
        Me.MenuCALSteel = New System.Windows.Forms.MenuItem()
        Me.MenuCalCoatedCord = New System.Windows.Forms.MenuItem()
        Me.MenuCALALLPreSemi = New System.Windows.Forms.MenuItem()
        Me.MenuCalSemi = New System.Windows.Forms.MenuItem()
        Me.MenuCalTT = New System.Windows.Forms.MenuItem()
        Me.MenuCalBF = New System.Windows.Forms.MenuItem()
        Me.MenuCalBelt = New System.Windows.Forms.MenuItem()
        Me.MenuCalSIC = New System.Windows.Forms.MenuItem()
        Me.MenuCalBW = New System.Windows.Forms.MenuItem()
        Me.MenuCalGT = New System.Windows.Forms.MenuItem()
        Me.MenuConvert = New System.Windows.Forms.MenuItem()
        Me.MenuItem14 = New System.Windows.Forms.MenuItem()
        Me.MenuPCOM = New System.Windows.Forms.MenuItem()
        Me.MenuPrice = New System.Windows.Forms.MenuItem()
        Me.MenuPriceRM = New System.Windows.Forms.MenuItem()
        Me.MenuPricePigment = New System.Windows.Forms.MenuItem()
        Me.MenuPriceCompound = New System.Windows.Forms.MenuItem()
        Me.MenuPricePreSemi = New System.Windows.Forms.MenuItem()
        Me.MenuPriceSemi = New System.Windows.Forms.MenuItem()
        Me.MenuPriceTire = New System.Windows.Forms.MenuItem()
        Me.MenuInv = New System.Windows.Forms.MenuItem()
        Me.MenuRecord = New System.Windows.Forms.MenuItem()
        Me.MenuReport = New System.Windows.Forms.MenuItem()
        Me.MenuItem13 = New System.Windows.Forms.MenuItem()
        Me.MenuView = New System.Windows.Forms.MenuItem()
        Me.MenuHorizontal = New System.Windows.Forms.MenuItem()
        Me.MenuVertical = New System.Windows.Forms.MenuItem()
        Me.MenuCasCade = New System.Windows.Forms.MenuItem()
        Me.MenuArrange = New System.Windows.Forms.MenuItem()
        Me.MenuHelp = New System.Windows.Forms.MenuItem()
        Me.MenuUpdate = New System.Windows.Forms.MenuItem()
        Me.MenuAbout = New System.Windows.Forms.MenuItem()
        Me.MenuCalNF = New System.Windows.Forms.MenuItem()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuFile, Me.MenuMaster, Me.MenuCalculate, Me.MenuPrice, Me.MenuInv, Me.MenuReport, Me.MenuView, Me.MenuHelp})
        '
        'MenuFile
        '
        Me.MenuFile.Index = 0
        Me.MenuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuClose})
        Me.MenuFile.Text = "File"
        '
        'MenuClose
        '
        Me.MenuClose.Index = 0
        Me.MenuClose.Shortcut = System.Windows.Forms.Shortcut.AltF4
        Me.MenuClose.Text = "Close"
        '
        'MenuMaster
        '
        Me.MenuMaster.Index = 1
        Me.MenuMaster.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem7, Me.MenuItem8, Me.MenuItem10, Me.MenuItem6, Me.MenuDept, Me.MenuItem32, Me.MenuAddUser})
        Me.MenuMaster.Text = "Master"
        Me.MenuMaster.Visible = False
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 0
        Me.MenuItem7.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuPIGMENT, Me.MenuItem31, Me.MenuPreSemi, Me.MenuSemi, Me.MenuTireCode})
        Me.MenuItem7.Text = "Material"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuRM, Me.MenuItem12, Me.MenuRMPrice, Me.MenuRMQty})
        Me.MenuItem1.Text = "R/M WareHouse"
        '
        'MenuRM
        '
        Me.MenuRM.Index = 0
        Me.MenuRM.Text = "R/M (Material)"
        '
        'MenuItem12
        '
        Me.MenuItem12.Index = 1
        Me.MenuItem12.Text = "-"
        '
        'MenuRMPrice
        '
        Me.MenuRMPrice.Index = 2
        Me.MenuRMPrice.Text = "R/M Price"
        '
        'MenuRMQty
        '
        Me.MenuRMQty.Index = 3
        Me.MenuRMQty.Text = "R/M Qty (KG)"
        Me.MenuRMQty.Visible = False
        '
        'MenuPIGMENT
        '
        Me.MenuPIGMENT.Index = 1
        Me.MenuPIGMENT.Text = "PIGMENT"
        '
        'MenuItem31
        '
        Me.MenuItem31.Index = 2
        Me.MenuItem31.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuCompound, Me.MenuRHC, Me.MenuPer})
        Me.MenuItem31.Text = "Compound"
        '
        'MenuCompound
        '
        Me.MenuCompound.Index = 0
        Me.MenuCompound.Text = "Weight"
        '
        'MenuRHC
        '
        Me.MenuRHC.Index = 1
        Me.MenuRHC.Text = "RHC "
        '
        'MenuPer
        '
        Me.MenuPer.Index = 2
        Me.MenuPer.Text = "% (Percent)"
        '
        'MenuPreSemi
        '
        Me.MenuPreSemi.Index = 3
        Me.MenuPreSemi.Text = "Pre Semi (Material)"
        '
        'MenuSemi
        '
        Me.MenuSemi.Index = 4
        Me.MenuSemi.Text = "Semi (Material)"
        '
        'MenuTireCode
        '
        Me.MenuTireCode.Index = 5
        Me.MenuTireCode.Text = "Green Tire "
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 1
        Me.MenuItem8.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuUnit, Me.MenuItem9, Me.MenuConvertUnit})
        Me.MenuItem8.Text = "Unit"
        '
        'MenuUnit
        '
        Me.MenuUnit.Index = 0
        Me.MenuUnit.Text = "Unit"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 1
        Me.MenuItem9.Text = "-"
        '
        'MenuConvertUnit
        '
        Me.MenuConvertUnit.Index = 2
        Me.MenuConvertUnit.Text = "Convert Unit"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 2
        Me.MenuItem10.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuType, Me.MenuItem11, Me.MenuTypeMaterial})
        Me.MenuItem10.Text = "Type"
        Me.MenuItem10.Visible = False
        '
        'MenuType
        '
        Me.MenuType.Index = 0
        Me.MenuType.Text = "Type"
        Me.MenuType.Visible = False
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 1
        Me.MenuItem11.Text = "-"
        Me.MenuItem11.Visible = False
        '
        'MenuTypeMaterial
        '
        Me.MenuTypeMaterial.Index = 2
        Me.MenuTypeMaterial.Text = "TypeMaterial"
        Me.MenuTypeMaterial.Visible = False
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 3
        Me.MenuItem6.Text = "BSJcode"
        '
        'MenuDept
        '
        Me.MenuDept.Index = 4
        Me.MenuDept.Text = "Cost Center"
        '
        'MenuItem32
        '
        Me.MenuItem32.Index = 5
        Me.MenuItem32.Text = "-"
        '
        'MenuAddUser
        '
        Me.MenuAddUser.Index = 6
        Me.MenuAddUser.Text = "New User"
        '
        'MenuCalculate
        '
        Me.MenuCalculate.Index = 2
        Me.MenuCalculate.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem5, Me.MenuItem16, Me.MenuCAL, Me.MenuConvert, Me.MenuItem14})
        Me.MenuCalculate.Text = "Calculate"
        Me.MenuCalculate.Visible = False
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 0
        Me.MenuItem5.Text = "CAL ALL Master "
        '
        'MenuItem16
        '
        Me.MenuItem16.Index = 1
        Me.MenuItem16.Text = "-"
        '
        'MenuCAL
        '
        Me.MenuCAL.Index = 2
        Me.MenuCAL.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuCalRM, Me.MenuCALPigment, Me.MenuCALCompound, Me.MenuCALPreSemi, Me.MenuCalSemi, Me.MenuCalGT})
        Me.MenuCAL.Text = "CAL Master Price "
        '
        'MenuCalRM
        '
        Me.MenuCalRM.Index = 0
        Me.MenuCalRM.Text = "RM (Row Material)"
        '
        'MenuCALPigment
        '
        Me.MenuCALPigment.Index = 1
        Me.MenuCALPigment.Text = "Pigment"
        '
        'MenuCALCompound
        '
        Me.MenuCALCompound.Index = 2
        Me.MenuCALCompound.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuCalCompoundStage1, Me.MenuCalCompoundStage2})
        Me.MenuCALCompound.Text = "Compound"
        '
        'MenuCalCompoundStage1
        '
        Me.MenuCalCompoundStage1.Index = 0
        Me.MenuCalCompoundStage1.Text = "Stage  1  "
        '
        'MenuCalCompoundStage2
        '
        Me.MenuCalCompoundStage2.Index = 1
        Me.MenuCalCompoundStage2.Text = "Stage  2 "
        '
        'MenuCALPreSemi
        '
        Me.MenuCALPreSemi.Index = 3
        Me.MenuCALPreSemi.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuCALSteel, Me.MenuCalCoatedCord, Me.MenuCALALLPreSemi})
        Me.MenuCALPreSemi.Text = "PreSemi"
        '
        'MenuCALSteel
        '
        Me.MenuCALSteel.Index = 0
        Me.MenuCALSteel.Text = "Steel CORD"
        '
        'MenuCalCoatedCord
        '
        Me.MenuCalCoatedCord.Index = 1
        Me.MenuCalCoatedCord.Text = "COATED CORD"
        '
        'MenuCALALLPreSemi
        '
        Me.MenuCALALLPreSemi.Index = 2
        Me.MenuCALALLPreSemi.Text = "All PreSemi"
        '
        'MenuCalSemi
        '
        Me.MenuCalSemi.Index = 4
        Me.MenuCalSemi.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuCalTT, Me.MenuCalBF, Me.MenuCalBelt, Me.MenuCalSIC, Me.MenuCalBW, Me.MenuCalNF})
        Me.MenuCalSemi.Text = "Semi"
        '
        'MenuCalTT
        '
        Me.MenuCalTT.Index = 0
        Me.MenuCalTT.Text = "Tread"
        '
        'MenuCalBF
        '
        Me.MenuCalBF.Index = 1
        Me.MenuCalBF.Text = "BF (Up-Low-Center)"
        '
        'MenuCalBelt
        '
        Me.MenuCalBelt.Index = 2
        Me.MenuCalBelt.Text = "Belt"
        '
        'MenuCalSIC
        '
        Me.MenuCalSIC.Index = 3
        Me.MenuCalSIC.Text = "Side,Innerliner,Cussion"
        '
        'MenuCalBW
        '
        Me.MenuCalBW.Index = 4
        Me.MenuCalBW.Text = "Wire Chafer,Body Ply"
        '
        'MenuCalGT
        '
        Me.MenuCalGT.Index = 5
        Me.MenuCalGT.Text = "Green Tire"
        '
        'MenuConvert
        '
        Me.MenuConvert.Index = 3
        Me.MenuConvert.Text = "Convert (Material)"
        Me.MenuConvert.Visible = False
        '
        'MenuItem14
        '
        Me.MenuItem14.Index = 4
        Me.MenuItem14.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuPCOM})
        Me.MenuItem14.Text = "CAL Percent "
        Me.MenuItem14.Visible = False
        '
        'MenuPCOM
        '
        Me.MenuPCOM.Index = 0
        Me.MenuPCOM.Text = "Compound"
        '
        'MenuPrice
        '
        Me.MenuPrice.Index = 3
        Me.MenuPrice.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuPriceRM, Me.MenuPricePigment, Me.MenuPriceCompound, Me.MenuPricePreSemi, Me.MenuPriceSemi, Me.MenuPriceTire})
        Me.MenuPrice.Text = "Price"
        '
        'MenuPriceRM
        '
        Me.MenuPriceRM.Index = 0
        Me.MenuPriceRM.Text = "R/M Material"
        '
        'MenuPricePigment
        '
        Me.MenuPricePigment.Index = 1
        Me.MenuPricePigment.Text = "Pigment"
        '
        'MenuPriceCompound
        '
        Me.MenuPriceCompound.Index = 2
        Me.MenuPriceCompound.Text = "Compound"
        '
        'MenuPricePreSemi
        '
        Me.MenuPricePreSemi.Index = 3
        Me.MenuPricePreSemi.Text = "Presemi"
        '
        'MenuPriceSemi
        '
        Me.MenuPriceSemi.Index = 4
        Me.MenuPriceSemi.Text = "Semi"
        '
        'MenuPriceTire
        '
        Me.MenuPriceTire.Index = 5
        Me.MenuPriceTire.Text = "Green Tire"
        '
        'MenuInv
        '
        Me.MenuInv.Index = 4
        Me.MenuInv.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuRecord})
        Me.MenuInv.Text = "Record"
        Me.MenuInv.Visible = False
        '
        'MenuRecord
        '
        Me.MenuRecord.Index = 0
        Me.MenuRecord.Text = "Record Tag "
        '
        'MenuReport
        '
        Me.MenuReport.Index = 5
        Me.MenuReport.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem13})
        Me.MenuReport.Text = "Report"
        Me.MenuReport.Visible = False
        '
        'MenuItem13
        '
        Me.MenuItem13.Index = 0
        Me.MenuItem13.Text = "Report "
        '
        'MenuView
        '
        Me.MenuView.Index = 6
        Me.MenuView.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuHorizontal, Me.MenuVertical, Me.MenuCasCade, Me.MenuArrange})
        Me.MenuView.Text = "View"
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
        'MenuHelp
        '
        Me.MenuHelp.Index = 7
        Me.MenuHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuUpdate, Me.MenuAbout})
        Me.MenuHelp.Text = "Help"
        '
        'MenuUpdate
        '
        Me.MenuUpdate.Index = 0
        Me.MenuUpdate.Text = "Check for Update"
        Me.MenuUpdate.Visible = False
        '
        'MenuAbout
        '
        Me.MenuAbout.Index = 1
        Me.MenuAbout.Text = "About Inventory Record"
        '
        'MenuCalNF
        '
        Me.MenuCalNF.Index = 5
        Me.MenuCalNF.Text = "Nylon Chafer,Flipper"
        '
        'FrmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(672, 366)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.Menu = Me.MainMenu1
        Me.Name = "FrmMain"
        Me.Text = "Inventory Record System"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Delegate Sub DisposeDelegate()

    Private Sub FrmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Check splash screen
        If My.Application.SplashScreen IsNot Nothing Then
            Dim splashScreenDispose As New DisposeDelegate(AddressOf My.Application.SplashScreen.Dispose)
            My.Application.SplashScreen.Invoke(splashScreenDispose)
            Me.Activate() 'Focus main screen
        End If

        'Authentication
        Dim flog As New FrmLogin
        flog.ShowDialog()
        If flog.EmpIDValue <> String.Empty Then
            Me.WindowState = FormWindowState.Maximized
            CurrentIDUser = Login.FrmLogin.EmpID
            CurrentName = Login.FrmLogin.EmpName
            CurrentLevel = Login.FrmLogin.LevelUsage
        Else
            Me.Close()
        End If

        'Check level usage
        Select Case CurrentLevel
            Case "Administrator"
                MenuCalculate.Visible = True
                MenuMaster.Visible = True
            Case "User"
                'Nothing
            Case Else
                Me.Close()
        End Select

    End Sub

    Private Sub MenuDept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuDept.Click
        Dim fdept As New FrmDept
        fdept.MdiParent = Me
        fdept.Show()
    End Sub

    Private Sub MenuRM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuRM.Click
        Dim fRM As New FrmRM
        '  fRM.MdiParent = Me
        fRM.Show()
    End Sub

    Private Sub MenuUnit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuUnit.Click
        Dim fUnit As New FrmUnit
        fUnit.MdiParent = Me
        fUnit.Show()
    End Sub

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

    Private Sub MenuAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuAbout.Click
        Dim titleAttr As System.Reflection.AssemblyTitleAttribute = System.Reflection.Assembly.GetExecutingAssembly().GetCustomAttributes(GetType(System.Reflection.AssemblyTitleAttribute), False)(0)
        Dim version As String = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()
        MsgBox(titleAttr.Title & " " & version, MsgBoxStyle.Information, titleAttr.Title)
    End Sub

    Private Sub MenuPrice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuRMPrice.Click
        Dim fpriceRM As New FrmRMPrice
        ' fpriceRM.MdiParent = Me
        fpriceRM.Show()
    End Sub

    Private Sub MenuQty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuRMQty.Click
        Dim fRMQty As New FrmRMQty
        ' fRMQty.MdiParent = Me
        fRMQty.Show()
    End Sub

    Private Sub MenuType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuType.Click
        Dim fType As New FrmType
        ' fType.MdiParent = Me
        fType.Show()
    End Sub

    Private Sub MenuPIGMENT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuPIGMENT.Click
        Dim fPigment As New FrmPIGMENT
        ' fPigment.MdiParent = Me
        fPigment.Show()
    End Sub

    Private Sub MenuCompound_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCompound.Click
        Dim fCompound As New FrmCompound
        '  fCompound.MdiParent = Me
        fCompound.Show()
    End Sub

    Private Sub MenuPreSemi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuPreSemi.Click
        Dim fPreSemi As New FrmPreSemi
        ' fPreSemi.MdiParent = Me
        fPreSemi.Show()
    End Sub

    Private Sub MenuSemi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuSemi.Click
        Dim fSemi As New FrmSemi
        '   fSemi.MdiParent = Me
        fSemi.Show()
    End Sub

    Private Sub MenuTypeMaterial_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuTypeMaterial.Click
        Dim fTypeMaterial As New FrmTypeMaterial
        fTypeMaterial.MdiParent = Me 'Within MDI Form
        fTypeMaterial.Show()
    End Sub

    'Private Sub MenuLoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuLoc.Click
    '    'Dim floc As New FrmLoc
    '    'floc.MdiParent = Me
    '    'floc.Show()
    'End Sub

    Private Sub MenuTireCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuTireCode.Click
        Dim fgtire As New FrmGreenTire
        ' fgtire.MdiParent = Me
        fgtire.Show()
    End Sub

    Private Sub MenuClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuClose.Click
        Me.Close()
    End Sub

#Region "PriceAccount"
    Private Sub MenuPriceCompound_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuPriceCompound.Click
        Dim fcal As New FrmCALMaster
        fcal.lblCal.Text += "  Compound"
        fcal.txtname = "Compound"
        fcal.MdiParent = Me
        fcal.Show()
    End Sub

    Private Sub MenuPriceRM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuPriceRM.Click
        Dim fcal As New FrmCALMaster
        fcal.MdiParent = Me
        fcal.lblCal.Text += "  R/M (Material)"
        fcal.txtname = "RM"
        fcal.Show()
    End Sub

    Private Sub MenuPricePreSemi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuPricePreSemi.Click
        Dim fcal As New FrmCALMaster
        fcal.MdiParent = Me
        fcal.lblCal.Text += "  PreSemi (Material)"
        fcal.txtname = "PreSemi"
        fcal.Show()
    End Sub

    Private Sub MenuPriceSemi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuPriceSemi.Click
        Dim fcal As New FrmCALMaster
        fcal.MdiParent = Me
        fcal.lblCal.Text += "  Semi (Material)"
        fcal.txtname = "Semi"
        fcal.Show()
    End Sub

    Private Sub MenuPriceTire_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuPriceTire.Click
        Dim fcal As New FrmCALMaster
        fcal.MdiParent = Me
        fcal.lblCal.Text += "  Green Tire"
        fcal.txtname = "Green Tire"
        fcal.Show()
    End Sub

    Private Sub MenuPricePigment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuPricePigment.Click
        Dim fcal As New FrmCALMaster
        fcal.MdiParent = Me
        fcal.lblCal.Text += "  Pigment"
        fcal.txtname = "Pigment"
        fcal.Show()
    End Sub

#End Region

    Private Sub MenuConvert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuConvert.Click

    End Sub

    Private Sub MenuRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuRecord.Click
        Dim ftag As New FrmYearInvTag
        Dim finvtag As New FrmInvTag
        ftag.MdiParent = Me
        finvtag.CurrentIDUserValue = Login.FrmLogin.EmpID
        finvtag.CurrentNameValue = Login.FrmLogin.EmpName
        finvtag.CurrentLevelValue = Login.FrmLogin.LevelUsage
        ftag.Show()
    End Sub

    Private Sub MenuAddUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuAddUser.Click
        Dim frmadduser As New FrmAddUser
        '   frmadduser.MdiParent = Me
        frmadduser.Show()
    End Sub

    Private Sub MenuItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem13.Click
        Dim fRpt As New frmPvRptTag
        fRpt.ShowDialog()
    End Sub


    Private Sub MenuRHC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuRHC.Click
        Dim fRHC As New FrmRHC
        '  fRHC.MdiParent = Me
        fRHC.Show()
    End Sub

    Private Sub MenuPer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuPer.Click
        Dim fpRHC As New FrmPerRHC
        ' fpRHC.MdiParent = Me
        fpRHC.Show()
    End Sub
    Private Sub LoadRM()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSql = "select final,mastercode,Revision,RMcode,Weight,RHC from TBLRHCDtl"
        StrSql += "  where  per Is null"
        If Not DT Is Nothing Then
            If DT.Rows.Count >= 1 Then
                DT.Clear()
            End If
        End If

        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSql, C1.Strcon)
            Dim CB As New SqlCommandBuilder(DA)
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
        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub

#Region "Percentcompound"
    Private Sub Loadcompound()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        StrSql = " select distinct final,Mastercode,Revision from TBLRHCDtl"
        If Not DT Is Nothing Then
            If DT.Rows.Count >= 1 Then
                DT.Clear()
            End If
        End If

        Dim DA As SqlDataAdapter
        Try
            DA = New SqlDataAdapter(StrSql, C1.Strcon)
            Dim CB As New SqlCommandBuilder(DA)
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
        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub

    Private Sub MenuPCOM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuPCOM.Click
        Dim vbal As Boolean
        vbal = True
        LoadRM()
        Dim dr As DataRow
        Dim aDr() As DataRow
        Dim i As Integer
        GrdDV.RowFilter = " RHC <> 0.000"
        aDr = GrdDV.Table.Select(GrdDV.RowFilter)

        For Each dr In aDr
            With dr
                If IIf(.Item("RMcode") Is System.DBNull.Value, "", .Item("RMcode")) <> "" Then

                    If CalPercent(.Item("Final"), .Item("MasterCode") _
                    , .Item("Revision"), .Item("RMcode"), .Item("Weight"), .Item("RHC")) Then
                        i = i + 1

                    Else
                        vbal = False
                    End If
                End If
            End With
        Next

        Loadcompound()
        GrdDV.RowFilter = " "
        aDr = GrdDV.Table.Select(GrdDV.RowFilter)
        For Each dr In aDr
            With dr
                If IIf(.Item("MasterCode") Is System.DBNull.Value, "", .Item("MasterCode")) <> "" Then
                    CalTotalPercent(.Item("Final"), .Item("MasterCode"), .Item("Revision"))
                End If
            End With
        Next

        If vbal Then
            MsgBox("Cal Percent Complete.  " & i & " Record")

        Else
            MsgBox("Cal Percent not Complete.")
        End If

    End Sub

    Private Function CalPercent(ByVal final As String, ByVal Mastercode As String,
  ByVal Rev As String, ByVal RMcode As String, ByVal Weight As Double, ByVal RHC As Double) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        CalPercent = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CalPercent"
        cmd2.Connection = cnn

        Dim sparam0 As SqlClient.SqlParameter
        sparam0 = New SqlClient.SqlParameter
        sparam0.ParameterName = "@Final"
        sparam0.SqlDbType = SqlDbType.Char
        sparam0.Size = 20
        sparam0.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam0)

        Dim sparam1 As SqlClient.SqlParameter
        sparam1 = New SqlClient.SqlParameter
        sparam1.ParameterName = "@MasterCode"
        sparam1.SqlDbType = SqlDbType.Char
        sparam1.Size = 20
        sparam1.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam1)

        Dim sparam2 As SqlClient.SqlParameter
        sparam2 = New SqlClient.SqlParameter
        sparam2.ParameterName = "@REV"
        sparam2.SqlDbType = SqlDbType.Char
        sparam2.Size = 3
        sparam2.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam2)

        Dim sparam3 As SqlClient.SqlParameter
        sparam3 = New SqlClient.SqlParameter
        sparam3.ParameterName = "@RMCode"
        sparam3.SqlDbType = SqlDbType.Char
        sparam3.Size = 20
        sparam3.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam3)

        Dim sparam4 As SqlClient.SqlParameter
        sparam4 = New SqlClient.SqlParameter
        sparam4.ParameterName = "@RHC"
        sparam4.SqlDbType = SqlDbType.Float
        '    sparam4.Size = 20
        sparam4.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam4)

        Dim sparam5 As SqlClient.SqlParameter
        sparam5 = New SqlClient.SqlParameter
        sparam5.ParameterName = "@Weight"
        sparam5.SqlDbType = SqlDbType.Float
        '    sparam5.Size = 20
        sparam5.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam5)

        Dim sparam6 As SqlClient.SqlParameter
        sparam6 = New SqlClient.SqlParameter
        sparam6.ParameterName = "@errID"
        sparam6.SqlDbType = SqlDbType.Char
        sparam6.Size = 4
        sparam6.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam6)

        Dim sparam7 As SqlClient.SqlParameter
        sparam7 = New SqlClient.SqlParameter
        sparam7.ParameterName = "@errMsg"
        sparam7.SqlDbType = SqlDbType.Char
        sparam7.Size = 40
        sparam7.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam7)

        Dim Reader As SqlClient.SqlDataReader
        cmd2.Parameters("@Final").Value = final.Trim
        cmd2.Parameters("@MasterCode").Value = Mastercode.Trim
        cmd2.Parameters("@Rev").Value = Rev.Trim
        cmd2.Parameters("@RMCode").Value = RMcode.Trim
        cmd2.Parameters("@Weight").Value = Weight
        cmd2.Parameters("@RHC").Value = RHC

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            CalPercent = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            CalPercent = False
        End Try
        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Function

    Private Function CalTotalPercent(ByVal final As String, ByVal Mastercode As String,
   ByVal Rev As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        CalTotalPercent = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CalTotalPercent"
        cmd2.Connection = cnn

        Dim sparam0 As SqlClient.SqlParameter
        sparam0 = New SqlClient.SqlParameter
        sparam0.ParameterName = "@Final"
        sparam0.SqlDbType = SqlDbType.Char
        sparam0.Size = 20
        sparam0.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam0)

        Dim sparam1 As SqlClient.SqlParameter
        sparam1 = New SqlClient.SqlParameter
        sparam1.ParameterName = "@MasterCode"
        sparam1.SqlDbType = SqlDbType.Char
        sparam1.Size = 20
        sparam1.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam1)

        Dim sparam2 As SqlClient.SqlParameter
        sparam2 = New SqlClient.SqlParameter
        sparam2.ParameterName = "@REV"
        sparam2.SqlDbType = SqlDbType.Char
        sparam2.Size = 3
        sparam2.Direction = ParameterDirection.Input
        cmd2.Parameters.Add(sparam2)

        Dim sparam3 As SqlClient.SqlParameter
        sparam3 = New SqlClient.SqlParameter
        sparam3.ParameterName = "@errID"
        sparam3.SqlDbType = SqlDbType.Char
        sparam3.Size = 4
        sparam3.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam3)

        Dim sparam4 As SqlClient.SqlParameter
        sparam4 = New SqlClient.SqlParameter
        sparam4.ParameterName = "@errMsg"
        sparam4.SqlDbType = SqlDbType.Char
        sparam4.Size = 40
        sparam4.Direction = ParameterDirection.Output
        cmd2.Parameters.Add(sparam4)

        Dim Reader As SqlClient.SqlDataReader
        cmd2.Parameters("@Final").Value = final.Trim
        cmd2.Parameters("@MasterCode").Value = Mastercode.Trim
        cmd2.Parameters("@Rev").Value = Rev.Trim

        cnn.Open()
        Try
            Reader = cmd2.ExecuteReader()
            CalTotalPercent = True
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            CalTotalPercent = False
        End Try
        cnn.Close()
        Me.Cursor = System.Windows.Forms.Cursors.Default()
    End Function
#End Region

#Region "CalRM Price"
    Private Sub MenuCalRM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCalRM.Click
        Dim msg As String = "Calculate RM Price " ' Define message.
        Dim title As String = "Calculate"   ' Define title.
        Dim style As MsgBoxStyle = MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        Dim response As MsgBoxResult

        ' Display message.
        response = MsgBox(msg, style, title)

        If response = MsgBoxResult.Yes Then ' User chose Yes.
            Dim StrDate, StrTime As String

            'Get datetime
            StrDate = DateTime.Now.ToString("yyyyMMdd", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))
            StrTime = DateTime.Now.ToString("HHmm", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))

            If CalRM(StrDate, StrTime) Then
                MessageBox.Show("Update RM Price .", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Don't Update RM Price .", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Else
            Exit Sub
        End If
    End Sub
    Private Function CalRM(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
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
        Me.Cursor = System.Windows.Forms.Cursors.Default()
        Return ret
    End Function
#End Region

#Region "CalPigment Price"
    Private Sub MenuCALPigment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCALPigment.Click
        Dim msg As String = "Calculate Pigment Price " ' Define message.
        Dim title As String = "Calculate"   ' Define title.
        Dim style As MsgBoxStyle = MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        Dim response As MsgBoxResult

        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            Dim StrDate, StrTime As String

            'Get datetime
            StrDate = DateTime.Now.ToString("yyyyMMdd", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))
            StrTime = DateTime.Now.ToString("HHmm", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))

            If CalPigment(StrDate, StrTime) Then
                MessageBox.Show("Update Pigment Price .", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Don't Update Pigment Price .", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Else
            Exit Sub
        End If
    End Sub
    Private Function CalPigment(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
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
        Me.Cursor = System.Windows.Forms.Cursors.Default()
        Return ret
    End Function
#End Region

#Region "CalCompound Price"
    Private Function CalCompound(ByVal dateup As String, ByVal timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
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
        Me.Cursor = System.Windows.Forms.Cursors.Default()
        Return ret
    End Function
    Private Function CalCompound2(ByVal dateup As String, ByVal timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
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
        Me.Cursor = System.Windows.Forms.Cursors.Default()
        Return ret
    End Function

    Private Sub MenuCalCompoundStage1_Click(sender As Object, e As EventArgs) Handles MenuCalCompoundStage1.Click
        Dim msg As String = "Calculate Compound Stage-1 Price " ' Define message.
        Dim title As String = "Calculate"   ' Define title.
        Dim style As MsgBoxStyle = MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        Dim response As MsgBoxResult

        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            Dim StrDate, StrTime As String
            'Get datetime
            StrDate = DateTime.Now.ToString("yyyyMMdd", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))
            StrTime = DateTime.Now.ToString("HHmm", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))

            'Call Store Procedure CALCompound (Insert Table TBLMasterPriceRM)
            If CalCompound(StrDate, StrTime) Then
                MessageBox.Show("Update Compound Price .", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Don't Update Compound Price .", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Else
            Exit Sub
        End If
    End Sub

    Private Sub MenuCalCompoundStage2_Click(sender As Object, e As EventArgs) Handles MenuCalCompoundStage2.Click
        Dim msg As String = "Calculate Compound Stage-2 Price " ' Define message.
        Dim title As String = "Calculate"   ' Define title.
        Dim style As MsgBoxStyle = MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        Dim response As MsgBoxResult

        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            Dim StrDate, StrTime As String
            'Get datetime
            StrDate = DateTime.Now.ToString("yyyyMMdd", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))
            StrTime = DateTime.Now.ToString("HHmm", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))

            'Call Store Procedure CALCompound2 (Insert Table TBLMasterPrice)
            If CalCompound2(StrDate, StrTime) Then
                MessageBox.Show("Update Compound Price .", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Don't Update Compound Price .", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Else
            Exit Sub
        End If
    End Sub
#End Region

#Region "CalPreSemi Price (CoatedCord)"
    Private Sub MenuCalCoatedCord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCalCoatedCord.Click
        Dim msg As String = "Calculate Coated Cord Price " ' Define message.
        Dim title As String = "Calculate"   ' Define title.
        Dim style As MsgBoxStyle = MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        Dim response As MsgBoxResult

        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            Dim StrDate, StrTime As String
            'Get datetime
            StrDate = DateTime.Now.ToString("yyyyMMdd", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))
            StrTime = DateTime.Now.ToString("HHmm", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))

            CalCoatedcord2(StrDate, StrTime) 'Call Store Procedure CALCoatedCord2 (Insert Table TBLMasterPriceRM Exclude TBLRM)
            CalCoatedcord3(StrDate, StrTime) 'Call Store Procedure CALCoatedCord3 (Insert Table TBLMasterPriceRM Include TBLRM)

            'Call Store Procedure CALCoatedCord (Insert Table TBLMasterPrice)
            If CalCoatedcord(StrDate, StrTime) Then
                MessageBox.Show("Update Coated Cord Price .", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Don't Update Coated Cord Price .", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Else
            Exit Sub
        End If

    End Sub
    Private Function CalCoatedcord(ByVal dateup As String, ByVal timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
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
        Me.Cursor = System.Windows.Forms.Cursors.Default()
        Return ret
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
        cmd2.Parameters("@Date").Value = dateup.Trim()
        cmd2.Parameters("@Time").Value = timeup.Trim()

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
        cmd2.Parameters("@Date").Value = dateup.Trim()
        cmd2.Parameters("@Time").Value = timeup.Trim()

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
    Private Sub MenuCALALLPreSemi_Click(sender As Object, e As EventArgs) Handles MenuCALALLPreSemi.Click
        Dim msg As String = "Calculate PreSemi Price " ' Define message.
        Dim title As String = "Calculate"   ' Define title.
        Dim style As MsgBoxStyle = MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        Dim response As MsgBoxResult

        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            Dim StrDate, StrTime As String
            'Get datetime
            StrDate = DateTime.Now.ToString("yyyyMMdd", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))
            StrTime = DateTime.Now.ToString("HHmm", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))

            'Call Store Procedure CALPreSemi (Material Type is not STEEL CORD and COATED CORD)
            If PreSemi(StrDate, StrTime) Then
                MessageBox.Show("Update PreSemi Price.", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Don't Update PreSemi Price.", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Else
            Exit Sub
        End If
    End Sub
    Private Function PreSemi(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
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
        Me.Cursor = System.Windows.Forms.Cursors.Default()
        Return ret
    End Function
#End Region

#Region "CalPreSemi Price (SteelCord)"
    Private Sub MenuCALSteel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCALSteel.Click
        Dim msg As String = "Calculate SteelCord Price " ' Define message.
        Dim title As String = "Calculate"   ' Define title.
        Dim style As MsgBoxStyle = MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        Dim response As MsgBoxResult

        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            Dim StrDate, StrTime As String
            'Get datetime
            StrDate = DateTime.Now.ToString("yyyyMMdd", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))
            StrTime = DateTime.Now.ToString("HHmm", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))

            'Call Store Procedure CALSteelCord (Material Type is STEEL CORD)
            If SteelCord(StrDate, StrTime) Then
                MessageBox.Show("Update SteelCord Price.", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Don't Update SteelCord Price.", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Else
            Exit Sub
        End If

    End Sub
    Private Function SteelCord(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
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
        Me.Cursor = System.Windows.Forms.Cursors.Default()
        Return ret
    End Function
#End Region

#Region "CalSemi Price"
#Region "TT"
    Private Sub MenuCalTT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCalTT.Click
        Dim msg As String = "Calculate Tread Price " ' Define message.
        Dim title As String = "Calculate"   ' Define title.
        Dim style As MsgBoxStyle = MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        Dim response As MsgBoxResult

        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            Dim StrDate, StrTime As String
            'Get datetime
            StrDate = DateTime.Now.ToString("yyyyMMdd", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))
            StrTime = DateTime.Now.ToString("HHmm", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))

            'Call Store Procedure CALTT (Material Type is TREAD)
            If Tread(StrDate, StrTime) Then
                MessageBox.Show("Update Tread Price.", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Don't Update Tread Price.", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Else
            Exit Sub
        End If

    End Sub
    Private Function Tread(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
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
        Me.Cursor = System.Windows.Forms.Cursors.Default()
        Return ret
    End Function
#End Region

#Region "BF"
    Private Sub MenuCalBF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCalBF.Click
        Dim msg As String = "Calculate BF Price " ' Define message.
        Dim title As String = "Calculate"   ' Define title.
        Dim style As MsgBoxStyle = MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        Dim response As MsgBoxResult

        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            Dim StrDate, StrTime As String
            'Get datetime
            StrDate = DateTime.Now.ToString("yyyyMMdd", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))
            StrTime = DateTime.Now.ToString("HHmm", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))

            'Call Store Procedure CALBF (Material Type is BF)(Insert Table TBLMasterPrice)
            If BF(StrDate, StrTime) Then
                'Nothing
            Else
                Exit Sub
            End If

            'Call Store Procedure CALBF2 (Material Type is BF)(Insert Table TBLMasterPriceRM)
            If BF2(StrDate, StrTime) Then
                MessageBox.Show("Update BF Price .", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                MessageBox.Show("Update BF Price. Not Complete", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Else
            Exit Sub
        End If

    End Sub
    Private Function BF(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
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
        Me.Cursor = System.Windows.Forms.Cursors.Default()
        Return ret
    End Function
    Private Function BF2(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Dim ret As Boolean = False
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

#Region "BElT 1-4"
    Private Sub MenuCalBelt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCalBelt.Click
        Dim msg As String = "Calculate BELT 1-4  Price " ' Define message.
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult

        style = MsgBoxStyle.DefaultButton2 Or
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Calculate"   ' Define title.
        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            Dim StrDate, StrTime As String
            'Get datetime
            StrDate = DateTime.Now.ToString("yyyyMMdd", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))
            StrTime = DateTime.Now.ToString("HHmm", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))

            'Call Store Procedure CALBELT (Type Material:BELT-1, BELT-2, BELT-3, BELT-4 to Table TBLMASTERPRICE)
            If BElT(StrDate, StrTime) Then
                'Nothing
            Else
                Exit Sub
            End If

            'Call Store Procedure CALBELT2 (Type Material:BELT-1, BELT-2, BELT-3, BELT-4 to Table TBLMASTERPRICERM)
            If BElT2(StrDate, StrTime) Then
                MessageBox.Show("Update BELT 1-4 Price.", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Update BELT 1-4 Price. Not Complete", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Else
            Exit Sub
        End If

    End Sub
    Private Function BElT(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Dim ret As Boolean = False
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
    Private Function BElT2(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Dim ret As Boolean = False
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

#Region "SIDE CUSSION INNERLINER"
    Private Sub MenuCalSIC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCalSIC.Click
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult

        msg = "Calculate SIDE,CUSSION,INNERLINER Price " ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Calculate"   ' Define title.
        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            Dim StrDate, StrTime As String
            'Get datetime
            StrDate = DateTime.Now.ToString("yyyyMMdd", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))
            StrTime = DateTime.Now.ToString("HHmm", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))

            'Call Store Procedure CALSIC (Type Material:CUSSION, SIDE, INNERLINER)
            If SIC(StrDate, StrTime) Then
                MessageBox.Show("Update SIDE,CUSSION,INNERLINER Price.", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Update SIDE,CUSSION,INNERLINER Price. Not Complete", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Else
            Exit Sub
        End If

    End Sub
    Private Function SIC(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
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
        Me.Cursor = System.Windows.Forms.Cursors.Default()
        Return ret
    End Function
#End Region

#Region "Wire Chafer,Body Ply"
    Private Sub MenuCalBW_Click(sender As Object, e As EventArgs) Handles MenuCalBW.Click
        Dim msg As String = "Calculate Wire Chafer,Body Ply Price " ' Define message.
        Dim title As String = "Calculate"   ' Define title.
        Dim style As MsgBoxStyle = MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        Dim response As MsgBoxResult

        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            Dim StrDate, StrTime As String
            'Get datetime
            StrDate = DateTime.Now.ToString("yyyyMMdd", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))
            StrTime = DateTime.Now.ToString("HHmm", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))

            'Call Store Procedure (Type Material:BODY PLY, WIRE CHAFER to Table TBLMASTERPRICE)
            If BW(StrDate, StrTime) Then
                'Nothing
            Else
                Exit Sub
            End If

            'Call Store Procedure (Type Material:BODY PLY, WIRE CHAFER to Table TBLMASTERPRICERM)
            If BW2(StrDate, StrTime) Then
                MessageBox.Show("Update Wire Chafer,Body Ply Price", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Update Wire Chafer,Body Ply Price. Not Complete", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Else
            Exit Sub
        End If
    End Sub

    Private Function BW(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Dim ret As Boolean = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CALBW"
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
    Private Function BW2(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Dim ret As Boolean = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CALBW2"
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

#Region "Nylon Chafer,Flipper"
    Private Sub MenuCalNF_Click(sender As Object, e As EventArgs) Handles MenuCalNF.Click
        Dim msg As String = "Calculate Nylon Chafer,Flipper Price " ' Define message.
        Dim title As String = "Calculate"   ' Define title.
        Dim style As MsgBoxStyle = MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        Dim response As MsgBoxResult

        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            Dim StrDate, StrTime As String
            'Get datetime
            StrDate = DateTime.Now.ToString("yyyyMMdd", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))
            StrTime = DateTime.Now.ToString("HHmm", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))

            'Call Store Procedure (Type Material:Nylon CHAFER, FLIPPER to Table TBLMASTERPRICE)
            If NF(StrDate, StrTime) Then
                'Nothing
            Else
                Exit Sub
            End If

            'Call Store Procedure (Type Material:Nylon CHAFER,FLIPPER to Table TBLMASTERPRICERM)
            If NF2(StrDate, StrTime) Then
                MessageBox.Show("Update Nylon Chafer,Flipper Price", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Update Nylon Chafer,Flipper Price. Not Complete", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Else
            Exit Sub
        End If
    End Sub

    Private Function NF(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Dim ret As Boolean = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CALNF"
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
    Private Function NF2(ByVal dateup As String, ByVal Timeup As String) As Boolean
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor()
        Dim ret As Boolean = False
        Dim cnn As New SqlConnection(C1.Strcon)

        Dim cmd2 As SqlClient.SqlCommand
        cmd2 = New SqlClient.SqlCommand
        cmd2.CommandTimeout = 0
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.CommandText = "CALNF2"
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
#End Region

#Region "CALGT Price"
    Private Sub MenuCalGT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCalGT.Click
        Dim msg As String = "Calculate Green Tire Price " ' Define message.
        Dim title As String = "Calculate"   ' Define title.
        Dim style As MsgBoxStyle = MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        Dim response As MsgBoxResult

        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then ' User chose Yes.
            Dim StrDate, StrTime As String
            'Get datetime
            StrDate = DateTime.Now.ToString("yyyyMMdd", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))
            StrTime = DateTime.Now.ToString("HHmm", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))

            'Call Store Procedure CALGT
            If GT(StrDate, StrTime) Then
                MessageBox.Show("Update Green Tire Price.", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Don't Update Green Tire Price.", "Calculate", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Else
            Exit Sub
        End If

    End Sub
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
        Me.Cursor = System.Windows.Forms.Cursors.Default()
        Return ret
    End Function
#End Region

    Private Sub MenuConvertUnit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuConvertUnit.Click
        Dim fc As New FrmCVT
        fc.MdiParent = Me
        fc.Show()
    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        Dim frmc As New Calculate
        frmc.ShowDialog()
    End Sub

    Private Sub MenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem6.Click
        Dim fbsj As New FrmBSJ
        fbsj.MdiParent = Me
        fbsj.Show()
    End Sub
End Class
