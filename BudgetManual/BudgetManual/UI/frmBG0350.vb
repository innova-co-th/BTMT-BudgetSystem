Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class frmBG0350

#Region "Variable"
    Private myClsBG0350BL As New clsBG0350BL
    Private myClsBG0310BL As New clsBG0310BL
#End Region

#Region "Overrides Function"
    Public Sub New(ByRef frmParent As Form, ByVal strFormName As String, ByVal blnMaximize As Boolean)
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.MdiParent = frmParent
        If blnMaximize Then
            Me.WindowState = FormWindowState.Maximized
        Else
            Me.WindowState = FormWindowState.Normal
        End If
        Me.Text = strFormName
    End Sub
#End Region


#Region "Function"

    Public Sub LoadPeriodList()
        '// Initialize controls


        myClsBG0310BL.GetAllPeriodList()

        If myClsBG0310BL.PeriodList IsNot Nothing AndAlso myClsBG0310BL.PeriodList.Rows.Count > 0 Then

            myClsBG0310BL.PeriodList.DefaultView.RowFilter = "PERIOD_TYPE_ID <> 9"

            cboPeriodType.DisplayMember = "PERIOD_TYPE_NAME"
            cboPeriodType.ValueMember = "PERIOD_TYPE_ID"
            cboPeriodType.DataSource = myClsBG0310BL.PeriodList

            cboPeriodType.SelectedIndex = 0
        End If

    End Sub

#End Region

#Region "Control Event"

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub cmdBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowse.Click
        '// Open File dialog for select data file
        If OpenFileDialog1.ShowDialog(Me) <> Windows.Forms.DialogResult.Cancel Then
            txtFilePath.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub cmdImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdImport.Click
        Const FIRST_BUDGET_DATA_COL As Integer = 16
        Const FIRST_ACTUAL_DATA_COL As Integer = 17
        Const FIRST_MTP_DATA_COL As Integer = 52
        Dim srFileReader As System.IO.StreamReader
        Dim strInputLine As String
        Dim blnHeaderRow As Boolean = True
        Dim arrDataList As String() = Nothing
        Dim strTmpYear As String = String.Empty
        Dim strTmpMonth As String = String.Empty
        Dim DataList As ArrayList
        Dim n As Integer
        Dim lngCounter As Long = 0
        Dim blnPeroidExist As Boolean = False
        Dim blnBudgetDataExist As Boolean = False
        Dim lngTotalCounter As Long = 0

        Me.Cursor = Cursors.WaitCursor

        Try
            '// Open File dialog for select data file
            If My.Computer.FileSystem.FileExists(OpenFileDialog1.FileName) = True Then

                If OpenFileDialog1.SafeFileName <> cboPeriodType.Text & ".txt" Then
                    MessageBox.Show("The selected text file is invalid!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

                    Me.Cursor = Cursors.Default
                    Exit Sub
                End If

                '// Open text file
                srFileReader = System.IO.File.OpenText(OpenFileDialog1.FileName)

                '// Read text line by line
                strInputLine = srFileReader.ReadLine()
                Do Until strInputLine Is Nothing

                    Debug.Print("Read Data File: " & strInputLine)

                    '// Split text line
                    arrDataList = Split(strInputLine, vbTab)
                    For i = 0 To arrDataList.Length - 1
                        If Trim(arrDataList(i)) = "-" Then
                            arrDataList(i) = "0"
                        End If
                    Next

                    '// Get information Header
                    If blnHeaderRow = True Then
                        blnHeaderRow = False

                        '// Check content of text file
                        strTmpYear = arrDataList(FIRST_BUDGET_DATA_COL).Substring(0, 4)
                        strTmpMonth = arrDataList(FIRST_BUDGET_DATA_COL).Substring(5, 2)
                        If Not IsNumeric(strTmpYear) Or Not IsNumeric(strTmpMonth) Or strTmpMonth <> "01" Then
                            MessageBox.Show("The selected text file is invalid!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End If

                        '// Read text next line
                        strInputLine = srFileReader.ReadLine()

                        Continue Do
                    End If

                    '// ----------------- Add Budget Order No's format check (space only) ------- (+)
                    lngTotalCounter += 1

                    If Me.fncCheckBudgetOrderNo(arrDataList(0)) = False Then

                        '// Read text next line
                        strInputLine = srFileReader.ReadLine()

                        Continue Do

                    End If
                    '// ----------------- Add Order budget No's format check (space only) ------- (-)

                    '// Import Budget Data
                    '// --Set parameters
                    myClsBG0350BL.BudgetYear = strTmpYear
                    myClsBG0350BL.PeriodType = CStr(cboPeriodType.SelectedValue)
                    myClsBG0350BL.BudgetOrder = arrDataList(0)
                    myClsBG0350BL.DataType = CStr(enumUploadDataType.BudgetData)
                    myClsBG0350BL.ProjectNo = numProjectNo.Value.ToString
                    myClsBG0350BL.RevNo = numRev.Value.ToString

                    '// Check peroid existed
                    '// If peroid not exist, then show error message.
                    If blnPeroidExist = False Then

                        If myClsBG0350BL.PeriodType = CStr(enumPeriodType.BudgetCompareVer10) Or _
                            myClsBG0350BL.PeriodType = CStr(enumPeriodType.BudgetCompareVer20) Then

                            blnPeroidExist = True

                        Else

                            If myClsBG0350BL.PeriodType = CStr(enumPeriodType.EstimateBudget2) Or _
                                                   myClsBG0350BL.PeriodType = CStr(enumPeriodType.EstimateBudget3) Then

                                myClsBG0350BL.CheckPeriodType = CStr(enumPeriodType.EstimateBudget)

                            ElseIf myClsBG0350BL.PeriodType = CStr(enumPeriodType.ForecastBudget2) Or _
                            myClsBG0350BL.PeriodType = CStr(enumPeriodType.ForecastBudget3) Or _
                            myClsBG0350BL.PeriodType = CStr(enumPeriodType.ForecastBudget4) Then

                                myClsBG0350BL.CheckPeriodType = CStr(enumPeriodType.ForecastBudget)

                            ElseIf myClsBG0350BL.PeriodType = CStr(enumPeriodType.OriginalBudget3) Then

                                myClsBG0350BL.CheckPeriodType = CStr(enumPeriodType.OriginalBudget)

                            Else

                                myClsBG0350BL.CheckPeriodType = CStr(cboPeriodType.SelectedValue)

                            End If

                            If myClsBG0350BL.CheckPeroidExist() = False Then
                                MessageBox.Show("The selected peroid not exist!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

                                Me.Cursor = Cursors.Default
                                Exit Sub

                            Else
                                blnPeroidExist = True
                            End If

                        End If

                    End If

                    If blnBudgetDataExist = False Then

                        If myClsBG0350BL.PeriodType = CStr(enumPeriodType.BudgetCompareVer10) Or _
                            myClsBG0350BL.PeriodType = CStr(enumPeriodType.BudgetCompareVer20) Then

                            blnPeroidExist = True

                        Else

                            '// Check Budget header existed.
                            If myClsBG0350BL.CheckBudgetHeaderExist() = False Then

                                MessageBox.Show("The budget header of selected budget data not exist!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

                                Me.Cursor = Cursors.Default
                                Exit Sub

                            End If

                            '// Check Budget data existed.
                            If myClsBG0350BL.CheckBudgetDataExist() = False Then
                                MessageBox.Show("The selected budget data not exist!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

                                Me.Cursor = Cursors.Default
                                Exit Sub

                            Else
                                blnPeroidExist = True
                            End If

                        End If

                    End If

                    DataList = New ArrayList

                    n = FIRST_BUDGET_DATA_COL
                    DataList.Add(arrDataList(n)) '// M1
                    n += 3
                    DataList.Add(arrDataList(n)) '// M2
                    n += 3
                    DataList.Add(arrDataList(n)) '// M3
                    n += 3
                    DataList.Add(arrDataList(n)) '// M4
                    n += 3
                    DataList.Add(arrDataList(n)) '// M5
                    n += 3
                    DataList.Add(arrDataList(n)) '// M6
                    n += 3
                    DataList.Add(arrDataList(n)) '// M7
                    n += 3
                    DataList.Add(arrDataList(n)) '// M8
                    n += 3
                    DataList.Add(arrDataList(n)) '// M9
                    n += 3
                    DataList.Add(arrDataList(n)) '// M10
                    n += 3
                    DataList.Add(arrDataList(n)) '// M11
                    n += 3
                    DataList.Add(arrDataList(n)) '// M12
                    n = 8
                    DataList.Add(arrDataList(n)) '// H1
                    n = 12
                    DataList.Add(arrDataList(n)) '// H2

                    '// --Import data to database
                    myClsBG0350BL.DataList = DataList
                    myClsBG0350BL.ImportData()

                    '// Import Actual Data
                    '// --Set parameters
                    myClsBG0350BL.BudgetYear = strTmpYear
                    myClsBG0350BL.PeriodType = CStr(cboPeriodType.SelectedValue)
                    myClsBG0350BL.BudgetOrder = arrDataList(0)
                    myClsBG0350BL.DataType = CStr(enumUploadDataType.ActualData)
                    myClsBG0350BL.ProjectNo = numProjectNo.Value.ToString
                    DataList = New ArrayList

                    n = FIRST_ACTUAL_DATA_COL
                    DataList.Add(arrDataList(n)) '// M1
                    n += 3
                    DataList.Add(arrDataList(n)) '// M2
                    n += 3
                    DataList.Add(arrDataList(n)) '// M3
                    n += 3
                    DataList.Add(arrDataList(n)) '// M4
                    n += 3
                    DataList.Add(arrDataList(n)) '// M5
                    n += 3
                    DataList.Add(arrDataList(n)) '// M6
                    n += 3
                    DataList.Add(arrDataList(n)) '// M7
                    n += 3
                    DataList.Add(arrDataList(n)) '// M8
                    n += 3
                    DataList.Add(arrDataList(n)) '// M9
                    n += 3
                    DataList.Add(arrDataList(n)) '// M10
                    n += 3
                    DataList.Add(arrDataList(n)) '// M11
                    n += 3
                    DataList.Add(arrDataList(n)) '// M12
                    n = 9
                    DataList.Add(arrDataList(n)) '// H1
                    n = 13
                    DataList.Add(arrDataList(n)) '// H2

                    '// --Import data to database
                    myClsBG0350BL.DataList = DataList
                    myClsBG0350BL.ImportData()

                    If arrDataList.Length >= FIRST_MTP_DATA_COL + 2 Then
                        '// Import MTP Data
                        '// --Set parameters
                        myClsBG0350BL.BudgetYear = strTmpYear
                        myClsBG0350BL.PeriodType = CStr(cboPeriodType.SelectedValue)
                        myClsBG0350BL.BudgetOrder = arrDataList(0)
                        myClsBG0350BL.DataType = CStr(enumUploadDataType.MTPData)
                        myClsBG0350BL.ProjectNo = numProjectNo.Value.ToString
                        DataList = New ArrayList

                        n = FIRST_MTP_DATA_COL
                        DataList.Add(arrDataList(n)) '// RRT1
                        n += 1
                        DataList.Add(arrDataList(n)) '// RRT2
                        n += 1
                        DataList.Add(arrDataList(n)) '// RRT3
                        'n += 1
                        'DataList.Add(arrDataList(n)) '// RRT4
                        'n += 1
                        'DataList.Add(arrDataList(n)) '// RRT5

                        '// --Import data to database
                        myClsBG0350BL.DataList = DataList
                        myClsBG0350BL.ImportData()
                    End If

                    lngCounter += 1

                    '// Read text next line
                    strInputLine = srFileReader.ReadLine()
                Loop

                '// Close text file
                srFileReader.Close()

                MessageBox.Show(lngCounter.ToString("#,##0") & " record(s) of " & lngTotalCounter.ToString("#,##0") & " record(s) was imported.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                '// Write Transaction Log
                WriteTransactionLog(CStr(enumOperationCd.ImportData), myClsBG0350BL.BudgetYear, myClsBG0350BL.PeriodType, "", myClsBG0350BL.DataType, "", myClsBG0350BL.ProjectNo)

            Else
                MessageBox.Show("The selected file not exist.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If

        Catch ex As Exception
            MessageBox.Show("[Import Data] Error: " & ex.Message, My.Settings.ProgramTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Me.Cursor = Cursors.Default

    End Sub

    Private Sub frmBG0350_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' cboPeriodType.SelectedIndex = 0
        LoadPeriodList()
    End Sub

    Private Function fncCheckBudgetOrderNo(ByVal strBudgetOrderNo As String) As Boolean

        '// Check first character is space, or not.
        If strBudgetOrderNo.Substring(0, 1) = " " Then

            Return False

        End If

        Return True
    End Function

#End Region

    Private Sub cboPeriodType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPeriodType.SelectedIndexChanged
        If cboPeriodType.SelectedIndex >= 0 Then

            If CStr(cboPeriodType.SelectedValue) = CStr(enumPeriodType.MBPBudget) Then
                Me.numProjectNo.Enabled = True
            Else
                Me.numProjectNo.Value = 1
                Me.numProjectNo.Enabled = False
            End If

        End If
    End Sub
End Class