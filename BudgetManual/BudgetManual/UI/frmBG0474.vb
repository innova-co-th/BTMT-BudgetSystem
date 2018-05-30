Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports CrystalDecisions.CrystalReports.Engine
Imports Microsoft.Office.Interop
Imports System.Drawing.Printing

Public Class frmBG0474

#Region "Variable"
    Const AC_M_COL As String = "AC_M"
    Const OB_M_COL As String = "OB_M"
    Const AC_H_COL As String = "ACC_AC_M"
    Const OB_H_COL As String = "ACC_OB_M"
    Const AC_Y_COL As String = "ACC_AC_Y"
    Const OB_Y_COL As String = "ACC_OB_Y"

    Private myClsBG0474BL As New clsBG0474BL
    Private myClsBG0310BL As New clsBG0310BL
    Private clsBG0400 As frmBG0400
    Private excelApp As Excel.Application
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

    Private Function InitPage() As Boolean

        Try
            Me.numYear.Value = Now.Year

            Me.cboMonth.Text = Now.Month.ToString

            LoadPersonInCharge()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error " & Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Function

    Private Sub Print(ByVal blnShowPrintPreview As Boolean)
        Dim strReportName As String = String.Empty
        Try
            If Me.cboMonth.SelectedIndex = -1 Then
                MessageBox.Show("Please select Month!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Me.cboMonth.Focus()
                Me.cboMonth.SelectAll()
                Return
            End If

            Cursor = Cursors.WaitCursor

            myClsBG0474BL.BudgetYear = CStr(Me.numYear.Value)
            myClsBG0474BL.Month = Me.cboMonth.Text
            myClsBG0474BL.PIC = Me.cboUserPIC.SelectedValue.ToString
            myClsBG0474BL.UserLevelId = p_intUserLevelId

            If myClsBG0474BL.getBudgetCompareData() Then

                Dim ds As DataSet = myClsBG0474BL.BudgetCompareData

                If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then

                    myClsBG0474BL.GetBudgetStatus()

                    strReportName = "RPT007-4.rpt"

                    If blnShowPrintPreview Then
                        'If clsBG0400 Is Nothing OrElse clsBG0400.IsDisposed Then
                        '    clsBG0400 = New frmBG0400()
                        'End If
                        If clsBG0400 IsNot Nothing Then
                            clsBG0400.Close()
                            clsBG0400.Dispose()
                        End If
                        clsBG0400 = New frmBG0400()
                        clsBG0400.MdiParent = p_frmBG0010
                        clsBG0400.ReportName = strReportName
                        clsBG0400.BudgetYear = Me.numYear.Value.ToString()
                        'clsBG0400.ParamPersonInCharge = False
                        clsBG0400.ReportType = "BudgetCompare"
                        clsBG0400.BudgetStatus = myClsBG0474BL.BudgetStatus
                        clsBG0400.Month = Me.cboMonth.Text

                        clsBG0400.DS = ds

                        clsBG0400.Show()
                        If clsBG0400.WindowState = FormWindowState.Minimized Then
                            clsBG0400.WindowState = FormWindowState.Normal
                        End If
                        clsBG0400.BringToFront()
                    Else
                        'Dim document As New Printing.PrintDocument ' This is the document to print
                        'Dim pageDialog2 As New PageSetupDialog ' This Dialog can set the paper size or kind
                        'Dim printDialog1 As New PrintDialog ' This is the dialog to setting the printer options
                        'Dim psize As Printing.PaperSize = Nothing

                        '' The parameter of Item method is any kind of paper size avaliable on the printer

                        'For i = 0 To document.PrinterSettings.PaperSizes.Count - 1

                        '    If document.PrinterSettings.PaperSizes.Item(i).Kind = PaperKind.A4 Then

                        '        psize = document.PrinterSettings.PaperSizes.Item(i)

                        '        Exit For

                        '    End If

                        'Next

                        'If psize Is Nothing Then
                        '    psize = document.PrinterSettings.PaperSizes.Item(0)
                        'End If

                        '' psize = document.PrinterSettings.PaperSizes.Item(6)

                        '' This line set the Page size of the document
                        'document.DefaultPageSettings.PaperSize = psize
                        'document.DefaultPageSettings.Landscape = True

                        '' This is for setting the page size on the page dialog
                        'pageDialog2.Document = document
                        'pageDialog2.PageSettings.PaperSize = psize
                        'pageDialog2.PageSettings.Landscape = True
                        '' This is for setting the page size on the printDialog
                        'printDialog1.Document = document
                        PrintDialog1.AllowSomePages = True
                        ' Allow the user to choose the page range he or she would
                        ' like to print.
                        ' PrintDialog1.AllowSomePages = True

                        ' Show the help button.
                        PrintDialog1.ShowHelp = True

                        Dim result As DialogResult = PrintDialog1.ShowDialog()

                        ' If the result is OK then print the document.
                        If (result = DialogResult.OK) Then

                            Dim rpt1 As ReportDocument = Nothing

                            rpt1 = New ReportDocument()
                            Dim reportPath As String = p_strAppPath & "\Reports\" & strReportName
                            rpt1.Load(reportPath)

                            'myClsBG0474BL.GetBudgetStatus()

                            'If myClsBG0474BL.BudgetStatus >= 5 Then
                            '    rpt1.ReportDefinition.ReportObjects("picAuth10").ObjectFormat.EnableSuppress = False
                            'Else
                            '    rpt1.ReportDefinition.ReportObjects("picAuth10").ObjectFormat.EnableSuppress = True
                            'End If

                            'If myClsBG0474BL.BudgetStatus >= 6 Then
                            '    rpt1.ReportDefinition.ReportObjects("picAuth3").ObjectFormat.EnableSuppress = False
                            'Else
                            '    rpt1.ReportDefinition.ReportObjects("picAuth3").ObjectFormat.EnableSuppress = True
                            'End If

                            rpt1.SetDataSource(ds)

                            rpt1.SetParameterValue("BUDGET_YEAR", Me.numYear.Value.ToString())
                            rpt1.SetParameterValue("HALF_BUDGET_YEAR", Me.numYear.Value.ToString().Substring(2, 2))
                            rpt1.SetParameterValue("MONTH", Me.cboMonth.Text)

                            rpt1.PrintOptions.PrinterName = PrintDialog1.PrinterSettings.PrinterName
                            rpt1.PrintToPrinter(PrintDialog1.PrinterSettings.Copies, _
                                                PrintDialog1.PrinterSettings.Collate, _
                                                PrintDialog1.PrinterSettings.FromPage, _
                                                PrintDialog1.PrinterSettings.ToPage)

                        End If
                    End If
                Else
                    MessageBox.Show("No data is available for viewing reports!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

                'Else
                '    MessageBox.Show("There are errors during the retrieved view reports!", "Detail by Account Code Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Cursor = Cursors.Default
    End Sub

    Private Sub LoadPersonInCharge()

        If Me.numYear.Value.ToString <> "" Then

            myClsBG0474BL.BudgetYear = Me.numYear.Value.ToString

            If myClsBG0474BL.GetPersonInChargeList() = False Then
                cboUserPIC.DataSource = Nothing
            Else
                Dim dt As DataTable = myClsBG0474BL.PersonInCharge
                Dim dr As DataRow = dt.NewRow
                dr(0) = 0
                dr(1) = "All"
                dt.Rows.InsertAt(dr, 0)

                cboUserPIC.DataSource = dt
                cboUserPIC.DisplayMember = "PIC_NAME"
                cboUserPIC.ValueMember = "PERSON_IN_CHARGE_NO"
                If cboUserPIC.Items.Count > 0 Then
                    cboUserPIC.SelectedIndex = 0
                End If

            End If

        End If

    End Sub

    Private Function InsertExcelColumnData(ByRef dtColumns As DataTable, ByVal strYear As String, ByVal strMonth As String) As Boolean

        Dim dRow As DataRow
        Dim strHalfYear As String = strYear.Substring(2, 2)

        Dim intYear As Integer = CInt(strYear)
        Dim strLastYear As String = CStr(intYear - 1)

        Dim strHalfLastYear As String = CStr(intYear - 1).Substring(2, 2)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "Group_Header"
        dRow("Column_Title") = "Item"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "BUDGET_ORDER_NO"
        dRow("Column_Title") = "New Budget Code"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "COST_CENTER"
        dRow("Column_Title") = "Cost Center"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERSON_IN_CHARGE_NO"
        dRow("Column_Title") = "Person in Charge"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "OB_M" & strMonth
        dRow("Column_Title") = "Budget"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "AC_M" & strMonth
        dRow("Column_Title") = "Actual"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_MONTH"
        dRow("Column_Title") = "Variance"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "ACC_OB_M" & strMonth
        dRow("Column_Title") = "Budget"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "ACC_AC_M" & strMonth
        dRow("Column_Title") = "Actual"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_HALF"
        dRow("Column_Title") = "Variance"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "ACC_OB_Y" & strMonth
        dRow("Column_Title") = "Budget"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "ACC_AC_Y" & strMonth
        dRow("Column_Title") = "Actual"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_YEAR"
        dRow("Column_Title") = "Variance"
        dtColumns.Rows.Add(dRow)

        Return True

    End Function

    Private Sub CalculateDiff(ByVal strMonth As String, ByRef dsData As DataSet)

        '//Add column 
        Dim col As DataColumn = New DataColumn()
        col.ColumnName = "DIFF_MONTH"
        col.DataType = Type.GetType("System.Decimal")
        col.DefaultValue = 0.0
        dsData.Tables(0).Columns.Add(col)

        col = New DataColumn()
        col.ColumnName = "DIFF_HALF"
        col.DataType = Type.GetType("System.Decimal")
        col.DefaultValue = 0.0
        dsData.Tables(0).Columns.Add(col)

        col = New DataColumn()
        col.ColumnName = "DIFF_YEAR"
        col.DataType = Type.GetType("System.Decimal")
        col.DefaultValue = 0.0
        dsData.Tables(0).Columns.Add(col)

        For Each drData As DataRow In dsData.Tables(0).Rows

            '// diff month      : (AC-OB)
            drData![DIFF_MONTH] = Convert.ToDecimal(Nz(drData(AC_M_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drData(OB_M_COL & strMonth), 0.0))

            '// diff half       : (Accumulate AC- Accumulate OB)
            drData![DIFF_HALF] = Convert.ToDecimal(Nz(drData(AC_H_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drData(OB_H_COL & strMonth), 0.0))

            '// diff year       : (Accumulate AC- Accumulate OB)
            drData![DIFF_YEAR] = Convert.ToDecimal(Nz(drData(AC_Y_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drData(OB_Y_COL & strMonth), 0.0))

        Next

    End Sub

    Private Function SetupInvestmentSummaryGroupbyData(ByVal strMonth As String, _
                                                       ByVal dsData As DataSet, ByVal arrGroupColumnName As String(), ByVal arrGroupColumnTitles As String(), _
                                                      ByVal intDataColumnIndex As Integer, ByRef arrSecondGroups() As Integer, _
                                                      ByRef arrFirstGroups() As Integer) As DataSet

        Dim dsResult As DataSet = New DataSet
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object
        Dim drEmpty As DataRow
        Dim strScript As String
        Dim intGroupTotalIndex As Integer
        Dim intSecondGroupTotalIndex As Integer
        Dim dtResult As DataTable = dsData.Tables(0).Clone
        Dim dtTmp As DataTable = dsData.Tables(0).Clone
        Dim dtSecondGroup As DataTable = dsData.Tables(0).Clone

        Dim drTotal As DataRow

        '//Add one more column for group name
        Dim col As DataColumn = New DataColumn()
        col.ColumnName = "Group_Header"
        col.DataType = Type.GetType("System.String")
        dtResult.Columns.Add(col)

        col = New DataColumn()
        col.ColumnName = "Group_Header"
        col.DataType = Type.GetType("System.String")
        dtTmp.Columns.Add(col)

        col = New DataColumn()
        col.ColumnName = "Group_Header"
        col.DataType = Type.GetType("System.String")
        dtSecondGroup.Columns.Add(col)

        Dim strFirstGroupColumnName As String = arrGroupColumnName(0)
        Dim strFirstGroupColumnTitle As String = arrGroupColumnTitles(0)

        strScript = strFirstGroupColumnName
        Dim dtGroups As DataTable = dsData.Tables(0).DefaultView.ToTable(True, strScript)
        Dim intGroupCount As Integer = dtGroups.Rows.Count

        '//Caculate Grand Total line
        Dim drAllTotal As DataRow = dtResult.NewRow

        For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName

            If strColumnName.Equals(AC_M_COL & strMonth) OrElse _
                  strColumnName.Equals(OB_M_COL & strMonth) OrElse _
                  strColumnName.Equals(AC_H_COL & strMonth) OrElse _
                  strColumnName.Equals(OB_H_COL & strMonth) OrElse _
                  strColumnName.Equals(AC_Y_COL & strMonth) OrElse _
                  strColumnName.Equals(OB_Y_COL & strMonth) Then

                strExpression = "Sum(" + strColumnName + ")"
                returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
                drAllTotal(strColumnName) = returnValue

            End If

        Next

        intGroupTotalIndex = 0
        For i As Integer = 0 To intGroupCount - 1

            '//Seperate dataset data into several datatables according to group no
            If dtGroups.Rows(i)(0).ToString = String.Empty Then
                Continue For
            End If
            strScript = strFirstGroupColumnName + " = " + dtGroups.Rows(i)(0).ToString
            Dim arrRows As DataRow() = dsData.Tables(0).Select(strScript, arrGroupColumnName(1) & ", " & arrGroupColumnName(2))

            dtSecondGroup.Rows.Clear()
            For j As Integer = 0 To arrRows.Length - 1

                Dim drow(dtResult.Columns.Count - 1) As Object
                Dim drow2(dtResult.Columns.Count - 1) As Object

                arrRows(j).ItemArray.CopyTo(drow, 0)
                arrRows(j).ItemArray.CopyTo(drow2, 0)

                dtResult.Rows.Add(drow)
                dtResult.Rows(dtResult.Rows.Count - 1)("Group_Header") = dtResult.Rows(dtResult.Rows.Count - 1)("BUDGET_ORDER_NAME")

                dtSecondGroup.Rows.Add(drow2)
            Next

            '//Calculate total for each group
            drTotal = dtResult.NewRow
            For k As Integer = intDataColumnIndex To dtSecondGroup.Columns.Count - 2

                Dim strColumnName As String = dtSecondGroup.Columns(k).ColumnName

                If strColumnName.Equals(AC_M_COL & strMonth) OrElse _
                    strColumnName.Equals(OB_M_COL & strMonth) OrElse _
                    strColumnName.Equals(AC_H_COL & strMonth) OrElse _
                    strColumnName.Equals(OB_H_COL & strMonth) OrElse _
                    strColumnName.Equals(AC_Y_COL & strMonth) OrElse _
                    strColumnName.Equals(OB_Y_COL & strMonth) Then

                    strExpression = "Sum(" + strColumnName + ")"
                    returnValue = dtSecondGroup.Compute(strExpression, strFilter)
                    drTotal(strColumnName) = returnValue

                End If

            Next

            'For intIndex As Integer = 0 To intDataColumnIndex - 1
            '    drTotal(intIndex) = dtResult.Rows(intGroupTotalIndex)(intIndex)
            'Next
            drTotal("Group_Header") = dtSecondGroup.Rows(0)(strFirstGroupColumnTitle).ToString + " Capital Investment"

            'If bShowGroupName = True Then
            '    dtResult.TableName = arrRows(0)(strGroupColumnName).ToString & " " & arrRows(0)(strGroupColumnTitle).ToString
            'End If

            '//
            '//Calculate second group total line
            '//
            Dim strSecondGroupColumnName As String = arrGroupColumnName(1)
            Dim strSecondGroupColumnTitle As String = arrGroupColumnTitles(1)

            strScript = strSecondGroupColumnName
            Dim dtSecondGroups As DataTable = dtSecondGroup.DefaultView.ToTable(True, strScript)
            Dim intSecondGroupCount As Integer = dtSecondGroups.Rows.Count

            'ReDim arrSecondGroups(intSecondGroupCount - 1)

            intSecondGroupTotalIndex = intGroupTotalIndex
            Dim drSecondTotal As DataRow
            For j As Integer = 0 To intSecondGroupCount - 1

                If dtSecondGroups.Rows(j)(0).ToString = String.Empty Then
                    Continue For
                End If

                strScript = strSecondGroupColumnName + " = " + dtSecondGroups.Rows(j)(0).ToString
                Dim arrRows2 As DataRow() = dtSecondGroup.Select(strScript, arrGroupColumnName(2))
                dtTmp.Rows.Clear()

                For k As Integer = 0 To arrRows2.Length - 1
                    Dim drow(dtTmp.Columns.Count - 1) As Object
                    arrRows2(k).ItemArray.CopyTo(drow, 0)
                    dtTmp.Rows.Add(drow)
                Next

                '//Calculate total for each group
                drSecondTotal = dtResult.NewRow
                For k As Integer = intDataColumnIndex To dtTmp.Columns.Count - 2

                    Dim strColumnName As String = dtResult.Columns(k).ColumnName

                    If strColumnName.Equals(AC_M_COL & strMonth) OrElse _
                       strColumnName.Equals(OB_M_COL & strMonth) OrElse _
                       strColumnName.Equals(AC_H_COL & strMonth) OrElse _
                       strColumnName.Equals(OB_H_COL & strMonth) OrElse _
                       strColumnName.Equals(AC_Y_COL & strMonth) OrElse _
                       strColumnName.Equals(OB_Y_COL & strMonth) Then

                        strExpression = "Sum(" + strColumnName + ")"
                        returnValue = dtTmp.Compute(strExpression, strFilter)
                        drSecondTotal(strColumnName) = returnValue

                    End If
                Next

                ''//
                ''//Calculate third group records count
                ''//
                'Dim strThirdGroupColumnName As String = arrGroupColumnName(2)
                'Dim intThirdGroupEmptyIndex As Integer = intSecondGroupTotalIndex

                'strScript = strThirdGroupColumnName
                'Dim dtThirdGroups As DataTable = dtTmp.DefaultView.ToTable(True, strScript)
                'Dim intThirdGroupCount As Integer = dtThirdGroups.Rows.Count
                'Dim intGroup3EmptyRows As Integer = 0

                'If intThirdGroupCount > 1 Then

                '    For intThirdIndex As Integer = 0 To intThirdGroupCount - 1
                '        strScript = strThirdGroupColumnName + " = '" + dtThirdGroups.Rows(intThirdIndex)(0).ToString + "'"
                '        Dim arrRows3 As DataRow() = dtTmp.Select(strScript)
                '        If intThirdIndex <> intThirdGroupCount - 1 Then
                '            drEmpty = dtResult.NewRow
                '            dtResult.Rows.InsertAt(drEmpty, intThirdGroupEmptyIndex + arrRows3.Length + intGroup3EmptyRows)
                '            intGroup3EmptyRows = intGroup3EmptyRows + 1
                '            intThirdGroupEmptyIndex = intThirdGroupEmptyIndex + arrRows3.Length
                '        End If
                '    Next

                'End If

                Dim intGroupTotalCount As Integer = arrSecondGroups.Length
                ReDim Preserve arrSecondGroups(intGroupTotalCount)
                arrSecondGroups(intGroupTotalCount - 1) = intSecondGroupTotalIndex

                '// diff month
                drSecondTotal![DIFF_MONTH] = Convert.ToDecimal(Nz(drSecondTotal(AC_M_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drSecondTotal(OB_M_COL & strMonth), 0.0))

                '// diff half
                drSecondTotal![DIFF_HALF] = Convert.ToDecimal(Nz(drSecondTotal(AC_H_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drSecondTotal(OB_H_COL & strMonth), 0.0))

                '// diff year
                drSecondTotal![DIFF_YEAR] = Convert.ToDecimal(Nz(drSecondTotal(AC_Y_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drSecondTotal(OB_Y_COL & strMonth), 0.0))


                '//Add total cost for group 2
                dtResult.Rows.InsertAt(drSecondTotal, intSecondGroupTotalIndex)
                dtResult.Rows(intSecondGroupTotalIndex)("Group_Header") = "(" + CStr(j + 1) + ") " + dtResult.Rows(intSecondGroupTotalIndex + 1)(strSecondGroupColumnTitle).ToString

                '//Add one empty row
                drEmpty = dtResult.NewRow
                dtResult.Rows.InsertAt(drEmpty, intSecondGroupTotalIndex + dtTmp.Rows.Count + 1)

                intSecondGroupTotalIndex = intSecondGroupTotalIndex + dtTmp.Rows.Count + 2

            Next

            Dim intTmp As Integer = arrFirstGroups.Length
            ReDim Preserve arrFirstGroups(intTmp)
            arrFirstGroups(intTmp - 1) = intGroupTotalIndex

            '// diff month
            drTotal![DIFF_MONTH] = Convert.ToDecimal(Nz(drTotal(AC_M_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drTotal(OB_M_COL & strMonth), 0.0))

            '// diff half
            drTotal![DIFF_HALF] = Convert.ToDecimal(Nz(drTotal(AC_H_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drTotal(OB_H_COL & strMonth), 0.0))

            '// diff year
            drTotal![DIFF_YEAR] = Convert.ToDecimal(Nz(drTotal(AC_Y_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drTotal(OB_Y_COL & strMonth), 0.0))


            '//Add total cost for group 1
            dtResult.Rows.InsertAt(drTotal, intGroupTotalIndex)
            intGroupTotalIndex = intSecondGroupTotalIndex + 1
        Next

        '// diff month
        drAllTotal![DIFF_MONTH] = Convert.ToDecimal(Nz(drAllTotal(AC_M_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drAllTotal(OB_M_COL & strMonth), 0.0))

        '// diff half
        drAllTotal![DIFF_HALF] = Convert.ToDecimal(Nz(drAllTotal(AC_H_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drAllTotal(OB_H_COL & strMonth), 0.0))

        '// diff year
        drAllTotal![DIFF_YEAR] = Convert.ToDecimal(Nz(drAllTotal(AC_Y_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drAllTotal(OB_Y_COL & strMonth), 0.0))


        '//Add All total cost
        dtResult.Rows.InsertAt(drAllTotal, 0)
        dtResult.Rows(0)("Group_Header") = "BTMT Capital Investment"

        '//Return data table
        dsResult.Tables.Add(dtResult)
        Return dsResult

    End Function

    Private Function OutputExcel(ByVal dsData As DataSet, ByVal dtColumns As DataTable, ByVal bMTPCheck As Boolean, _
                                 ByVal strSubTitle As String, ByVal strYear As String, ByVal bShowGroupName As Boolean, _
                                 ByVal arrFirstGroups As Integer(), ByVal arrSecondGroups As Integer(), _
                                 ByVal strMonth As String) As Boolean

        If excelApp Is Nothing Then
            excelApp = New Excel.Application
        End If

        Dim rowStartIndex As Integer = 10
        Dim colStartIndex As Integer = 9

        Dim xBk As Excel.Workbook = Nothing
        Dim xSt As Excel.Worksheet = Nothing

        xBk = excelApp.Workbooks.Add(System.Reflection.Missing.Value)
        If xBk.Worksheets.Count > 1 Then
            For i As Integer = 1 To xBk.Worksheets.Count - 1
                CType(xBk.Worksheets(i), Excel.Worksheet).Delete()
            Next
        End If

        For intSheetCount As Integer = 0 To dsData.Tables.Count - 1

            If intSheetCount <> 0 Then
                xBk.Sheets.Add()
            End If

            xSt = CType(xBk.ActiveSheet, Excel.Worksheet)
            xSt.Name = dsData.Tables(intSheetCount).TableName

            '//Setup DataColumn
            For i As Integer = 0 To dtColumns.Rows.Count - 1
                xSt.Cells(colStartIndex, i + 1) = dtColumns.Rows(i)("Column_Title").ToString
                xSt.Range(xSt.Cells(colStartIndex, i + 1), xSt.Cells(colStartIndex, i + 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            Next

            Dim arrCols() As Integer

            arrCols = New Integer() {1, 2, 3, 4}
            SetupCompareColumnsCells(xSt, colStartIndex, CInt(strMonth), 1, 1, "Item", arrCols, 5, 7, 8, 10, strYear, True, True, 11, 13)


            '//Setup Data
            For rowIndex As Integer = 0 To dsData.Tables(intSheetCount).Rows.Count - 1

                Dim row As DataRow = dsData.Tables(intSheetCount).Rows(rowIndex)

                '//If the column is "Group_Header" Empty.
                If IsGroupHeaderEmpty(row) Then
                    Continue For
                End If

                For colIndex As Integer = 0 To dtColumns.Rows.Count - 1

                    Dim strColumnName As String = dtColumns.Rows(colIndex)("Column_Name").ToString
                    Dim col As DataColumn = dsData.Tables(intSheetCount).Columns(strColumnName)

                    If col.DataType Is System.Type.GetType("System.DateTime") Then

                        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = (Convert.ToDateTime(row(col.ColumnName).ToString())).ToString("yyyy-MM-dd")
                        xSt.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter

                    ElseIf col.DataType Is System.Type.GetType("System.String") Then

                        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = "'" + row(col.ColumnName).ToString()
                        xSt.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter

                    ElseIf col.DataType Is System.Type.GetType("System.Decimal") Then

                        If row(col.ColumnName).ToString = String.Empty Then
                            row(col.ColumnName) = "0.00"
                        End If

                        'If row(col.ColumnName).ToString <> String.Empty Then
                        '    If CDbl(row(col.ColumnName).ToString) = 0 Then
                        '        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = "-"
                        '        xSt.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                        '    Else
                        '        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName)
                        '        xSt.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).NumberFormat = "#,##0.00"
                        '    End If
                        'End If

                        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName)
                        xSt.Range(excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1), excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1)).NumberFormat = "#,##0.00"

                    Else
                        excelApp.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName).ToString()
                    End If

                Next
            Next

            '//Add Title for excel 
            Dim strGroupName As String = "Capital Investment"
            Dim intUnitPriceStart As Integer
            Dim intUnitPriceEnd As Integer
            Dim intAuthorizeStart As Integer
            Dim intAuthorizeEnd As Integer
            Dim intImageIndex As Integer

            SetupTitleIndex(intUnitPriceStart, intUnitPriceEnd, intAuthorizeStart, intAuthorizeEnd, intImageIndex)

            Dim bAuthorizeTwoCols As Boolean = False
            

            SetupExcelTitle(xSt, strSubTitle, strYear, bMTPCheck, intUnitPriceStart, intUnitPriceEnd, _
                            intAuthorizeStart, intAuthorizeEnd, intImageIndex, bShowGroupName, strGroupName, bAuthorizeTwoCols)

            Dim rowMax As Integer = dsData.Tables(intSheetCount).Rows.Count + colStartIndex
            Dim colMax As Integer = dtColumns.Rows.Count

            '//Setup group header column to be left align
            xSt.Range(xSt.Cells(rowStartIndex, 1), xSt.Cells(rowMax, 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup font bold for group total
            xSt.Range(xSt.Cells(rowStartIndex, 1), xSt.Cells(rowStartIndex, colMax)).Font.Bold = True
            xSt.Range(xSt.Cells(rowStartIndex + 1, 1), xSt.Cells(rowStartIndex + 1, colMax)).Font.Bold = True
            For i As Integer = 0 To arrSecondGroups.Length - 1
                Dim intIndexTmp As Integer = arrSecondGroups(i) + 2
                xSt.Range(xSt.Cells(rowStartIndex + intIndexTmp, 1), xSt.Cells(rowStartIndex + intIndexTmp, colMax)).Font.Bold = True
            Next

            For i As Integer = 0 To arrFirstGroups.Length - 1
                Dim intIndexTmp As Integer = arrFirstGroups(i) + 1
                xSt.Range(xSt.Cells(rowStartIndex + intIndexTmp, 1), xSt.Cells(rowStartIndex + intIndexTmp, colMax)).Font.Bold = True
            Next

            '//Setup sheet properly width
            xSt.Range(xSt.Cells(2, 1), xSt.Cells(rowMax, colMax)).Columns.EntireColumn.AutoFit()

            '//Setup Columns wrap text
            xSt.Range(xSt.Cells(2, 2), xSt.Cells(rowMax, 2)).Columns.ColumnWidth = 14
            xSt.Range(xSt.Cells(2, 2), xSt.Cells(rowMax, 2)).WrapText = True

            xSt.Range(xSt.Cells(2, 3), xSt.Cells(rowMax, 4)).Columns.ColumnWidth = 10
            xSt.Range(xSt.Cells(2, 3), xSt.Cells(rowMax, 4)).WrapText = True

            '// 
            xSt.Range(xSt.Cells(2, 5), xSt.Cells(rowMax, 13)).Columns.ColumnWidth = 12
           

            colStartIndex = colStartIndex - 1
            '//Setup Column Font 
            xSt.Range(excelApp.Cells(colStartIndex, 1), excelApp.Cells(colStartIndex + 1, colMax)).Font.Bold = True
            xSt.Range(xSt.Cells(colStartIndex, 1), xSt.Cells(rowMax, colMax)).Font.Name = "Tahoma"
            xSt.Range(xSt.Cells(colStartIndex, 1), xSt.Cells(rowMax, colMax)).Font.Size = 10

            '//Setup border
            xSt.Range(excelApp.Cells(colStartIndex, 1), excelApp.Cells(rowMax, colMax)).Borders.LineStyle = 1
            xSt.Range(excelApp.Cells(colStartIndex, 1), excelApp.Cells(rowMax, 1)).Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlThin
            xSt.Range(excelApp.Cells(colStartIndex, 1), excelApp.Cells(colStartIndex, colMax)).Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThin
            xSt.Range(excelApp.Cells(colStartIndex, colMax), excelApp.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThin
            xSt.Range(excelApp.Cells(rowMax, 1), excelApp.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThin

        
        Next

        '//Show Excel
        excelApp.Visible = True

        '//-- Begin Add by S.Watcharapong 2011-05-24
        '//Release memory
        BGCommon.ExcelReleasememory(excelApp, xBk, xSt)
        '//-- End Add 2011-05-24

        Return True

    End Function

    Private Function SetupTitleIndex(ByRef intUnitPriceStart As Integer, _
                                     ByRef intUnitPriceEnd As Integer, ByRef intAuthorizeStart As Integer, _
                                     ByRef intAuthorizeEnd As Integer, ByRef intImageIndex As Integer) As Boolean

        intUnitPriceStart = 22
        intUnitPriceEnd = 23

        intAuthorizeStart = 20
        intAuthorizeEnd = 21

        intImageIndex = 1235

        Return True

    End Function

#End Region

#Region "Controls Event"

    Private Sub frmBG0474_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        If Not excelApp Is Nothing Then
            excelApp.Quit()
            excelApp = Nothing
        End If
    End Sub

    Private Sub frmBG474_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        InitPage()

    End Sub

    Private Sub cmdPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPreview.Click
        Print(True)
    End Sub

    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrint.Click
        Print(False)
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub numYear_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numYear.ValueChanged

        LoadPersonInCharge()

    End Sub

    Private Sub cmdExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExcel.Click

        If Me.cboMonth.SelectedIndex = -1 Then
            MessageBox.Show("Please select Month!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.cboMonth.Focus()
            Me.cboMonth.SelectAll()
            Return
        End If

        Cursor = Cursors.WaitCursor

        Dim dsData As DataSet
        myClsBG0474BL.BudgetYear = CStr(Me.numYear.Value)
        myClsBG0474BL.Month = Me.cboMonth.Text
        myClsBG0474BL.PIC = Me.cboUserPIC.SelectedValue.ToString
        myClsBG0474BL.UserLevelId = p_intUserLevelId

        If myClsBG0474BL.GetBudgetCompareData() = False Then
            dsData = Nothing
            MessageBox.Show("Load buget data failed.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Cursor = Cursors.Default
            Return
        Else
            dsData = myClsBG0474BL.BudgetCompareData
            If dsData Is Nothing Or dsData.Tables(0).Rows.Count = 0 Then
                MessageBox.Show("No budget data found, please try it again.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Cursor = Cursors.Default
                Return
            End If
        End If

        Dim dtAuthorizeImages As DataTable = myClsBG0474BL.BudgetCompareData.Tables(1)

        '//Create output columns
        Dim strYear As String = Me.numYear.Value.ToString
        Dim strMonth As String = Me.cboMonth.Text
        Dim dtColumns As DataTable = CreateTableTemplate()

        Dim strSubTitle As String = "Summary by Investment : Budget Compare Year " + strYear

        InsertExcelColumnData(dtColumns, strYear, strMonth)
        CalculateDiff(strMonth, dsData)

        '//Create group data
        Dim intGroupLineIndex As Integer = 0
        Dim arrGroupNames As String() = New String() {"ASSET_PROJECT", "ASSET_CATEGORY", "PERSON_IN_CHARGE_NO"}
        Dim arrGroupColumnTitles As String() = New String() {"ASSET_PROJECT_TXT", "ASSET_CATEGORY_TXT", "PERSON_IN_CHARGE_NO"}
        Dim arrSecondGroups As Integer() = New Integer() {0}
        Dim arrFirstGroups As Integer() = New Integer() {0}
        Dim dsGroups As DataSet = SetupInvestmentSummaryGroupbyData(strMonth, dsData, arrGroupNames, _
                                                        arrGroupColumnTitles, 11, arrSecondGroups, arrFirstGroups)

        '//Create Output Excel
        OutputExcel(dsGroups, dtColumns, False, strSubTitle, strYear, True, arrFirstGroups, arrSecondGroups, strMonth)

        Me.Cursor = Cursors.Default

    End Sub
#End Region

End Class