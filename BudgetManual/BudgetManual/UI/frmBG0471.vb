Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports CrystalDecisions.CrystalReports.Engine
Imports Microsoft.Office.Interop
Imports System.Drawing.Printing

Public Class frmBG0471

#Region "Variable"
    Const AC_M_COL As String = "AC_M"
    Const OB_M_COL As String = "OB_M"
    Const OB_FC_M_COL As String = "OB_FC_M"
    Const OB_VC_M_COL As String = "OB_VC_M"
    Const AC_FC_M_COL As String = "AC_FC_M"
    Const AC_VC_M_COL As String = "AC_VC_M"
    Const AC_H_COL As String = "ACC_AC_M"
    Const OB_H_COL As String = "ACC_OB_M"
    Const AC_FC_H_COL As String = "ACC_AC_FC_M"
    Const OB_FC_H_COL As String = "ACC_OB_FC_M"
    Const AC_VC_H_COL As String = "ACC_AC_VC_M"
    Const OB_VC_H_COL As String = "ACC_OB_VC_M"

    Private myClsBG0471BL As New clsBG0471BL
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

            myClsBG0471BL.BudgetYear = CStr(Me.numYear.Value)
            myClsBG0471BL.Month = Me.cboMonth.Text
            myClsBG0471BL.UserLevelId = p_intUserLevelId

            If myClsBG0471BL.getBudgetCompareData() Then

                Dim ds As DataSet = myClsBG0471BL.BudgetCompareData

                If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then

                    myClsBG0471BL.GetBudgetStatus()

                    strReportName = "RPT007-1.rpt"

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
                        clsBG0400.BudgetStatus = myClsBG0471BL.BudgetStatus
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

                            'myClsBG0471BL.GetBudgetStatus()

                            'If myClsBG0471BL.BudgetStatus >= 5 Then
                            '    rpt1.ReportDefinition.ReportObjects("picAuth10").ObjectFormat.EnableSuppress = False
                            'Else
                            '    rpt1.ReportDefinition.ReportObjects("picAuth10").ObjectFormat.EnableSuppress = True
                            'End If

                            'If myClsBG0471BL.BudgetStatus >= 6 Then
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

    Private Function InsertExcelColumnData(ByRef dtColumns As DataTable, ByVal strYear As String, ByVal strMonth As String) As Boolean

        Dim dRow As DataRow
        Dim strDiffPeriod As String = BGCommon.GetBudgetCompareDiffPeriod(strMonth)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "Group_Header"
        dRow("Column_Title") = "Description"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERSON_IN_CHARGE"
        dRow("Column_Title") = "Person in Charge Section"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERSON_IN_CHARGE_NAME"
        dRow("Column_Title") = ""
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
        dRow("Column_Title") = strDiffPeriod
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERCENT_MONTH"
        dRow("Column_Title") = "%"
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
        dRow("Column_Title") = strDiffPeriod
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERCENT_HALF"
        dRow("Column_Title") = "%"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_VARIABLE"
        dRow("Column_Title") = "Variable Cost"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERCENT_VARIABLE"
        dRow("Column_Title") = "%"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "DIFF_FIXED"
        dRow("Column_Title") = "Fixed Cost"
        dtColumns.Rows.Add(dRow)

        dRow = dtColumns.NewRow
        dRow("Column_Name") = "PERCENT_FIXED"
        dRow("Column_Title") = "%"
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
        col.ColumnName = "PERCENT_MONTH"
        col.DataType = Type.GetType("System.Decimal")
        col.DefaultValue = 0.0
        dsData.Tables(0).Columns.Add(col)

        col = New DataColumn()
        col.ColumnName = "DIFF_HALF"
        col.DataType = Type.GetType("System.Decimal")
        col.DefaultValue = 0.0
        dsData.Tables(0).Columns.Add(col)

        col = New DataColumn()
        col.ColumnName = "PERCENT_HALF"
        col.DataType = Type.GetType("System.Decimal")
        col.DefaultValue = 0.0
        dsData.Tables(0).Columns.Add(col)

        col = New DataColumn()
        col.ColumnName = "DIFF_VARIABLE"
        col.DataType = Type.GetType("System.Decimal")
        col.DefaultValue = 0.0
        dsData.Tables(0).Columns.Add(col)

        col = New DataColumn()
        col.ColumnName = "PERCENT_VARIABLE"
        col.DataType = Type.GetType("System.Decimal")
        col.DefaultValue = 0.0
        dsData.Tables(0).Columns.Add(col)

        col = New DataColumn()
        col.ColumnName = "DIFF_FIXED"
        col.DataType = Type.GetType("System.Decimal")
        col.DefaultValue = 0.0
        dsData.Tables(0).Columns.Add(col)

        col = New DataColumn()
        col.ColumnName = "PERCENT_FIXED"
        col.DataType = Type.GetType("System.Decimal")
        col.DefaultValue = 0.0
        dsData.Tables(0).Columns.Add(col)

        For Each drData As DataRow In dsData.Tables(0).Rows
            '// diff month      : (AC-OB)
            drData![DIFF_MONTH] = Convert.ToDecimal(Nz(drData(AC_M_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drData(OB_M_COL & strMonth), 0.0))

            '// percent month   : (AC- OB)/OB
            drData![PERCENT_MONTH] = CalPercent(Convert.ToDecimal(Nz(drData(AC_M_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drData(OB_M_COL & strMonth), 0.0)))

            '// diff half       : (Accumulate AC- Accumulate OB)
            drData![DIFF_HALF] = Convert.ToDecimal(Nz(drData(AC_H_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drData(OB_H_COL & strMonth), 0.0))

            '// percent half    : (Accumulate AC- Accumulate OB)/Accumulate OB
            drData![PERCENT_HALF] = CalPercent(Convert.ToDecimal(Nz(drData(AC_H_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drData(OB_H_COL & strMonth), 0.0)))

            '// diff variable   : (Accumulate AC- Accumulate OB)
            drData![DIFF_VARIABLE] = Convert.ToDecimal(Nz(drData(AC_VC_H_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drData(OB_VC_H_COL & strMonth), 0.0))

            '// percent variable: (Accumulate AC- Accumulate OB)/Accumulate OB
            drData![PERCENT_VARIABLE] = CalPercent(Convert.ToDecimal(Nz(drData(AC_VC_H_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drData(OB_VC_H_COL & strMonth), 0.0)))

            '// diff fixed      : (Accumulate AC- Accumulate OB)
            drData![DIFF_FIXED] = Convert.ToDecimal(Nz(drData(AC_FC_H_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drData(OB_FC_H_COL & strMonth), 0.0))

            '// percent fixed   : (Accumulate AC- Accumulate OB)/Accumulate OB
            drData![PERCENT_FIXED] = CalPercent(Convert.ToDecimal(Nz(drData(AC_FC_H_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drData(OB_FC_H_COL & strMonth), 0.0)))

        Next

    End Sub

    Private Function SetupCompareGroupbyData(ByVal strMonth As String, _
                                             ByVal dsData As DataSet, _
                                             ByVal strGroupColumnName As String, _
                                             ByVal strGroupColumnTitle As String, _
                                             ByVal intDataColumnIndex As Integer, _
                                             ByVal bShowGroupName As Boolean, _
                                             ByRef intGroupFirstIndex As Integer, _
                                             ByRef intGroupSecondIndex As Integer) As DataSet

        Dim dsResult As DataSet = New DataSet
        Dim strExpression As String
        Dim strFilter As String = String.Empty
        Dim returnValue As Object
        Dim drEmpty As DataRow

        Dim strScript As String = strGroupColumnName
        'Dim strGroupbyScript As String = "Group by PERSON_IN_CHARGE_NO"
        'Dim arrGroups As DataRow() = dsData.Tables(0).Select(strScript)

        Dim dtGroups As DataTable = dsData.Tables(0).DefaultView.ToTable(True, strScript)
        Dim intGroupCount As Integer = dtGroups.Rows.Count

        Dim dtResult As DataTable = dsData.Tables(0).Clone
        Dim dtTmp As DataTable = dsData.Tables(0).Clone

        Dim drFixcostTotal As DataRow = dtResult.NewRow
        Dim drVariablecostTotal As DataRow = dtResult.NewRow
        Dim drAllTotal As DataRow = dtResult.NewRow

        For k As Integer = intDataColumnIndex To dsData.Tables(0).Columns.Count - 1

            Dim strColumnName As String = dsData.Tables(0).Columns(k).ColumnName

            strExpression = "Sum(" + strColumnName + ")"
            strFilter = ""
            returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
            drAllTotal(dsData.Tables(0).Columns(k).ColumnName) = returnValue

            If strColumnName.Equals(OB_M_COL & strMonth) Then

                strExpression = "Sum(" + OB_FC_M_COL + strMonth + ")"
                returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
                drFixcostTotal(dsData.Tables(0).Columns(k).ColumnName) = returnValue

                strExpression = "Sum(" + OB_VC_M_COL + strMonth + ")"
                returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
                drVariablecostTotal(dsData.Tables(0).Columns(k).ColumnName) = returnValue

            ElseIf strColumnName.Equals(AC_M_COL & strMonth) Then

                strExpression = "Sum(" + AC_FC_M_COL + strMonth + ")"
                returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
                drFixcostTotal(dsData.Tables(0).Columns(k).ColumnName) = returnValue

                strExpression = "Sum(" + AC_VC_M_COL + strMonth + ")"
                returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
                drVariablecostTotal(dsData.Tables(0).Columns(k).ColumnName) = returnValue

            ElseIf strColumnName.Equals(OB_H_COL & strMonth) Then

                strExpression = "Sum(" + OB_FC_H_COL + strMonth + ")"
                returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
                drFixcostTotal(dsData.Tables(0).Columns(k).ColumnName) = returnValue

                strExpression = "Sum(" + OB_VC_H_COL + strMonth + ")"
                returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
                drVariablecostTotal(dsData.Tables(0).Columns(k).ColumnName) = returnValue

            ElseIf strColumnName.Equals(AC_H_COL & strMonth) Then

                strExpression = "Sum(" + AC_FC_H_COL + strMonth + ")"
                returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
                drFixcostTotal(dsData.Tables(0).Columns(k).ColumnName) = returnValue

                strExpression = "Sum(" + AC_VC_H_COL + strMonth + ")"
                returnValue = dsData.Tables(0).Compute(strExpression, strFilter)
                drVariablecostTotal(dsData.Tables(0).Columns(k).ColumnName) = returnValue

            End If

            
        Next

        Dim intGroupTotalIndex As Integer = 0
        For i As Integer = 0 To intGroupCount - 1

            '//Seperate dataset data into several datatables according to group no
            strScript = strGroupColumnName + " = " + dtGroups.Rows(i)(0).ToString
            Dim arrRows As DataRow() = dsData.Tables(0).Select(strScript)

            dtTmp.Rows.Clear()
            For j As Integer = 0 To arrRows.Length - 1

                Dim drow(dtResult.Columns.Count - 1) As Object
                Dim dRowTmp(dtResult.Columns.Count - 1) As Object

                arrRows(j).ItemArray.CopyTo(drow, 0)
                arrRows(j).ItemArray.CopyTo(dRowTmp, 0)

                dtResult.Rows.Add(drow)
                dtTmp.Rows.Add(dRowTmp)
            Next

            '//Calculate total for each group
            Dim drTotal As DataRow = dtResult.NewRow
            For k As Integer = intDataColumnIndex To dtResult.Columns.Count - 1

                Dim strColumnName As String = dtResult.Columns(k).ColumnName
                strExpression = "Sum(" + strColumnName + ")"
                returnValue = dtTmp.Compute(strExpression, strFilter)
                drTotal(dtResult.Columns(k).ColumnName) = returnValue

                '// Diff Total
                drTotal![DIFF_MONTH] = Convert.ToDecimal(Nz(drTotal(AC_M_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drTotal(OB_M_COL & strMonth), 0.0))
                drTotal![PERCENT_MONTH] = CalPercent(Convert.ToDecimal(Nz(drTotal(AC_M_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drTotal(OB_M_COL & strMonth), 0.0)))

                drTotal![DIFF_HALF] = Convert.ToDecimal(Nz(drTotal(AC_H_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drTotal(OB_H_COL & strMonth), 0.0))
                drTotal![PERCENT_HALF] = CalPercent(Convert.ToDecimal(Nz(drTotal(AC_H_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drTotal(OB_H_COL & strMonth), 0.0)))

                drTotal![DIFF_VARIABLE] = Convert.ToDecimal(Nz(drTotal(AC_VC_H_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drTotal(OB_VC_H_COL & strMonth), 0.0))
                drTotal![PERCENT_VARIABLE] = CalPercent(Convert.ToDecimal(Nz(drTotal(AC_VC_H_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drTotal(OB_VC_H_COL & strMonth), 0.0)))

                drTotal![DIFF_FIXED] = Convert.ToDecimal(Nz(drTotal(AC_FC_H_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drTotal(OB_FC_H_COL & strMonth), 0.0))
                drTotal![PERCENT_FIXED] = CalPercent(Convert.ToDecimal(Nz(drTotal(AC_FC_H_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drTotal(OB_FC_H_COL & strMonth), 0.0)))

            Next

            '//Add total cost
            dtResult.Rows.InsertAt(drTotal, intGroupTotalIndex)

            '//Add one empty row
            drEmpty = dtResult.NewRow
            dtResult.Rows.Add(drEmpty)

            If bShowGroupName = True Then
                dtResult.TableName = arrRows(0)(strGroupColumnName).ToString & " " & arrRows(0)(strGroupColumnTitle).ToString
            End If

            intGroupTotalIndex = dtResult.Rows.Count
            If i = 0 Then
                intGroupFirstIndex = intGroupTotalIndex - 1
            Else
                intGroupSecondIndex = intGroupTotalIndex - 1
            End If
        Next

        'intGroupLineIndex = intGroupTotalIndex - 1

        '// Diff Total Variable
        drVariablecostTotal![DIFF_MONTH] = Convert.ToDecimal(Nz(drVariablecostTotal(AC_M_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drVariablecostTotal(OB_M_COL & strMonth), 0.0))
        drVariablecostTotal![PERCENT_MONTH] = CalPercent(Convert.ToDecimal(Nz(drVariablecostTotal(AC_M_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drVariablecostTotal(OB_M_COL & strMonth), 0.0)))

        drVariablecostTotal![DIFF_HALF] = Convert.ToDecimal(Nz(drVariablecostTotal(AC_H_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drVariablecostTotal(OB_H_COL & strMonth), 0.0))
        drVariablecostTotal![PERCENT_HALF] = CalPercent(Convert.ToDecimal(Nz(drVariablecostTotal(AC_H_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drVariablecostTotal(OB_H_COL & strMonth), 0.0)))

        '// Diff Total Fixed
        drFixcostTotal![DIFF_MONTH] = Convert.ToDecimal(Nz(drFixcostTotal(AC_M_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drFixcostTotal(OB_M_COL & strMonth), 0.0))
        drFixcostTotal![PERCENT_MONTH] = CalPercent(Convert.ToDecimal(Nz(drFixcostTotal(AC_M_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drFixcostTotal(OB_M_COL & strMonth), 0.0)))

        drFixcostTotal![DIFF_HALF] = Convert.ToDecimal(Nz(drFixcostTotal(AC_H_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drFixcostTotal(OB_H_COL & strMonth), 0.0))
        drFixcostTotal![PERCENT_HALF] = CalPercent(Convert.ToDecimal(Nz(drFixcostTotal(AC_H_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drFixcostTotal(OB_H_COL & strMonth), 0.0)))

        '// Diff Total
        drAllTotal![DIFF_MONTH] = Convert.ToDecimal(Nz(drAllTotal(AC_M_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drAllTotal(OB_M_COL & strMonth), 0.0))
        drAllTotal![PERCENT_MONTH] = CalPercent(Convert.ToDecimal(Nz(drAllTotal(AC_M_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drAllTotal(OB_M_COL & strMonth), 0.0)))

        drAllTotal![DIFF_HALF] = Convert.ToDecimal(Nz(drAllTotal(AC_H_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drAllTotal(OB_H_COL & strMonth), 0.0))
        drAllTotal![PERCENT_HALF] = CalPercent(Convert.ToDecimal(Nz(drAllTotal(AC_H_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drAllTotal(OB_H_COL & strMonth), 0.0)))

        drAllTotal![DIFF_VARIABLE] = Convert.ToDecimal(Nz(drAllTotal(AC_VC_H_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drAllTotal(OB_VC_H_COL & strMonth), 0.0))
        drAllTotal![PERCENT_VARIABLE] = CalPercent(Convert.ToDecimal(Nz(drAllTotal(AC_VC_H_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drAllTotal(OB_VC_H_COL & strMonth), 0.0)))

        drAllTotal![DIFF_FIXED] = Convert.ToDecimal(Nz(drAllTotal(AC_FC_H_COL & strMonth), 0.0)) - Convert.ToDecimal(Nz(drAllTotal(OB_FC_H_COL & strMonth), 0.0))
        drAllTotal![PERCENT_FIXED] = CalPercent(Convert.ToDecimal(Nz(drAllTotal(AC_FC_H_COL & strMonth), 0.0)), Convert.ToDecimal(Nz(drAllTotal(OB_FC_H_COL & strMonth), 0.0)))


        '//Add one empty row
        drEmpty = dtResult.NewRow
        dtResult.Rows.Add(drEmpty)

        '//Add variable cost total
        dtResult.Rows.Add(drVariablecostTotal)

        '//Add fixed cost total
        dtResult.Rows.Add(drFixcostTotal)

        '//Add All total cost
        dtResult.Rows.Add(drAllTotal)

        '//Add one empty row
        drEmpty = dtResult.NewRow
        dtResult.Rows.Add(drEmpty)

        '//Add one more column for group name
        Dim col As DataColumn = New DataColumn()
        col.ColumnName = "Group_Header"
        col.DataType = Type.GetType("System.String")
        dtResult.Columns.Add(col)

        dtResult.Rows(0)("Group_Header") = "Person in Charge"
        If intGroupCount > 1 Then
            dtResult.Rows(intGroupFirstIndex + 1)("Group_Header") = "Others"
        End If

        '//Return data table
        dsResult.Tables.Add(dtResult)
        Return dsResult

    End Function

    Private Function OutputExcel(ByVal dsData As DataSet, ByVal dtColumns As DataTable, ByVal bMTPCheck As Boolean, _
                               ByVal strSubTitle As String, ByVal strYear As String, ByVal strMonth As String, _
                               ByVal bShowGroupName As Boolean, ByVal intGroupFirstIndex As Integer, ByVal intGroupSecondIndex As Integer) As Boolean

        If excelApp Is Nothing Then
            excelApp = New Excel.Application
        End If

        Dim rowStartIndex As Integer = 10
        Dim colStartIndex As Integer = 9

        Dim xBk As Excel.Workbook = Nothing
        Dim xSt As Excel.Worksheet = Nothing

        Dim strHalfYear = strYear.Substring(2, 2)
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

            arrCols = New Integer() {1}    '// Two Row Merge Col
            SetupCompareColumnsCells(xSt, colStartIndex, CInt(strMonth), 2, 3, "Person in Charge Section", _
                                     arrCols, 4, 7, 8, 15, strYear)


            '//Setup Data
            For rowIndex As Integer = 0 To dsData.Tables(intSheetCount).Rows.Count - 1

                Dim row As DataRow = dsData.Tables(intSheetCount).Rows(rowIndex)

                For colIndex As Integer = 0 To dtColumns.Rows.Count - 1

                    Dim strColumnName As String = dtColumns.Rows(colIndex)("Column_Name").ToString
                    Dim col As DataColumn = dsData.Tables(intSheetCount).Columns(strColumnName)

                    If col.DataType Is System.Type.GetType("System.DateTime") Then

                        xSt.Cells(rowIndex + rowStartIndex, colIndex + 1) = (Convert.ToDateTime(row(col.ColumnName).ToString())).ToString("yyyy-MM-dd")
                        xSt.Range(xSt.Cells(rowIndex + rowStartIndex, colIndex + 1), xSt.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter

                    ElseIf col.DataType Is System.Type.GetType("System.String") Then

                        xSt.Cells(rowIndex + rowStartIndex, colIndex + 1) = "'" + row(col.ColumnName).ToString()
                        xSt.Range(xSt.Cells(rowIndex + rowStartIndex, colIndex + 1), xSt.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter

                    ElseIf col.DataType Is System.Type.GetType("System.Decimal") Then

                        If row(col.ColumnName).ToString = String.Empty Then
                            row(col.ColumnName) = "0.00"
                        End If

                        'If CDbl(row(col.ColumnName).ToString) = 0 Then
                        '    xSt.Cells(rowIndex + rowStartIndex, colIndex + 1) = "-"
                        '    xSt.Range(xSt.Cells(rowIndex + rowStartIndex, colIndex + 1), xSt.Cells(rowIndex + rowStartIndex, colIndex + 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                        'Else
                        '    xSt.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName)
                        '    xSt.Range(xSt.Cells(rowIndex + rowStartIndex, colIndex + 1), xSt.Cells(rowIndex + rowStartIndex, colIndex + 1)).NumberFormat = "#,##0.00"
                        'End If

                        xSt.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName)
                        xSt.Range(xSt.Cells(rowIndex + rowStartIndex, colIndex + 1), xSt.Cells(rowIndex + rowStartIndex, colIndex + 1)).NumberFormat = "#,##0.00"

                    Else
                        xSt.Cells(rowIndex + rowStartIndex, colIndex + 1) = row(col.ColumnName).ToString()
                    End If

                Next
            Next

            '//Add Title 
            Dim strGroupName As String = dsData.Tables(intSheetCount).TableName
            Dim intUnitPriceStart As Integer
            Dim intUnitPriceEnd As Integer
            Dim intAuthorizeStart As Integer
            Dim intAuthorizeEnd As Integer
            Dim intImageIndex As Integer

            SetupTitleIndex(intUnitPriceStart, intUnitPriceEnd, intAuthorizeStart, intAuthorizeEnd, intImageIndex, bMTPCheck)

            Dim bAuthorizeTwoCols As Boolean = False

            SetupExcelTitle(xSt, strSubTitle, strYear, bMTPCheck, intUnitPriceStart, intUnitPriceEnd, _
                            intAuthorizeStart, intAuthorizeEnd, intImageIndex, bShowGroupName, strGroupName, bAuthorizeTwoCols)

            Dim rowMax As Integer = dsData.Tables(intSheetCount).Rows.Count + colStartIndex
            Dim colMax As Integer = dtColumns.Rows.Count

            '//Setup budget order name column to be left align
            xSt.Range(xSt.Cells(rowStartIndex, 3), xSt.Cells(rowMax, 3)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            '//Setup Total lines
            SetupTotalLines(xSt, rowMax - 1, "Expenses Total", "Left", 3, 3, 3, rowMax - 4, rowMax - 3, rowMax - 2, colMax, bMTPCheck)

            xSt.Range(xSt.Cells(rowMax - 4, 1), xSt.Cells(rowMax, 2)).ClearContents()
            xSt.Range(xSt.Cells(rowMax - 4, 1), xSt.Cells(rowMax, 2)).MergeCells = True

            If bMTPCheck = False Then
                xSt.Range(xSt.Cells(rowMax, 3), xSt.Cells(rowMax, colMax)).ClearContents()
                xSt.Range(xSt.Cells(rowMax, 3), xSt.Cells(rowMax, colMax)).MergeCells = True
            End If

            Dim strGroup1 As String = xSt.Range(xSt.Cells(rowStartIndex, 1), xSt.Cells(rowStartIndex, 1)).Value.ToString
            xSt.Range(xSt.Cells(rowStartIndex, 1), xSt.Cells(rowStartIndex + intGroupFirstIndex, 1)).ClearContents()
            xSt.Range(xSt.Cells(rowStartIndex, 1), xSt.Cells(rowStartIndex + intGroupFirstIndex, 1)).MergeCells = True
            xSt.Range(xSt.Cells(rowStartIndex, 1), xSt.Cells(rowStartIndex + intGroupFirstIndex, 1)).Value = strGroup1
            xSt.Range(xSt.Cells(rowStartIndex, 1), xSt.Cells(rowStartIndex + intGroupFirstIndex, 1)).Font.Bold = True
            xSt.Range(xSt.Cells(rowStartIndex, 1), xSt.Cells(rowStartIndex + intGroupFirstIndex, 1)).VerticalAlignment = Excel.XlVAlign.xlVAlignTop

            'xSt.Range(xSt.Cells(rowStartIndex, 1), xSt.Cells(rowStartIndex, colMax)).Font.Bold = True

            Dim intGroup2 As Integer = rowStartIndex + intGroupFirstIndex + 1
            Dim strGroup2 As String = xSt.Range(xSt.Cells(intGroup2, 1), xSt.Cells(intGroup2, 1)).Value.ToString
            xSt.Range(xSt.Cells(intGroup2, 1), xSt.Cells(rowStartIndex + intGroupSecondIndex, 1)).ClearContents()
            xSt.Range(xSt.Cells(intGroup2, 1), xSt.Cells(rowStartIndex + intGroupSecondIndex, 1)).MergeCells = True
            xSt.Range(xSt.Cells(intGroup2, 1), xSt.Cells(rowStartIndex + intGroupSecondIndex, 1)).Value = strGroup2
            xSt.Range(xSt.Cells(intGroup2, 1), xSt.Cells(rowStartIndex + intGroupSecondIndex, 1)).Font.Bold = True
            xSt.Range(xSt.Cells(intGroup2, 1), xSt.Cells(rowStartIndex + intGroupSecondIndex, 1)).VerticalAlignment = Excel.XlVAlign.xlVAlignTop

            '//Setup sheet properly width
            xSt.Range(xSt.Cells(2, 1), xSt.Cells(rowMax, colMax)).Columns.AutoFit()


            '//Setup Wrap text for columns title
            xSt.Range(xSt.Cells(2, 4), xSt.Cells(rowMax, 15)).Columns.ColumnWidth = 12


            colStartIndex = colStartIndex - 1
            '//Setup Column Font 
            xSt.Range(xSt.Cells(colStartIndex, 1), xSt.Cells(colStartIndex + 1, colMax)).Font.Bold = True
            xSt.Range(xSt.Cells(colStartIndex, 1), xSt.Cells(rowMax, colMax)).Font.Name = "Tahoma"
            xSt.Range(xSt.Cells(colStartIndex, 1), xSt.Cells(rowMax, colMax)).Font.Size = 10

            '//Setup border
            xSt.Range(xSt.Cells(colStartIndex, 1), xSt.Cells(rowMax, colMax)).Borders.LineStyle = 1
            xSt.Range(xSt.Cells(colStartIndex, 1), xSt.Cells(rowMax, 1)).Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlThin
            xSt.Range(xSt.Cells(colStartIndex, 1), xSt.Cells(colStartIndex, colMax)).Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThin
            xSt.Range(xSt.Cells(colStartIndex, colMax), xSt.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThin
            xSt.Range(xSt.Cells(rowMax, 1), xSt.Cells(rowMax, colMax)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThin

            '//Clear data last row on Group
            xSt.Range(xSt.Cells(rowStartIndex + intGroupFirstIndex, 4), xSt.Cells(rowStartIndex + intGroupFirstIndex, colMax)).ClearContents()
            xSt.Range(xSt.Cells(rowStartIndex + intGroupSecondIndex, 4), xSt.Cells(rowStartIndex + intGroupSecondIndex, colMax)).ClearContents()

            '//Set Bold first row on Group
            xSt.Range(xSt.Cells(rowStartIndex, 4), xSt.Cells(rowStartIndex, colMax)).Font.Bold = True
            xSt.Range(xSt.Cells(intGroup2, 4), xSt.Cells(intGroup2, colMax)).Font.Bold = True
        Next

        '//Show Excel
        excelApp.Visible = True


        '//Release memory
        BGCommon.ExcelReleasememory(excelApp, xBk, xSt)


        Return True

    End Function

    Private Function SetupTitleIndex(ByRef intUnitPriceStart As Integer, _
                                     ByRef intUnitPriceEnd As Integer, ByRef intAuthorizeStart As Integer, _
                                     ByRef intAuthorizeEnd As Integer, ByRef intImageIndex As Integer, ByVal bMTPCheck As Boolean) As Boolean

        intUnitPriceStart = 14
        intUnitPriceEnd = 15

        intAuthorizeStart = 13
        intAuthorizeEnd = 14

        intImageIndex = 815

        Return True

    End Function

#End Region

#Region "Controls Event"

    Private Sub frmBG0471_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        If Not excelApp Is Nothing Then
            excelApp.Quit()
            excelApp = Nothing
        End If
    End Sub

    Private Sub frmBG471_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

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

    Private Sub cmdExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExcel.Click
        If Me.cboMonth.SelectedIndex = -1 Then
            MessageBox.Show("Please select Month!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.cboMonth.Focus()
            Me.cboMonth.SelectAll()
            Return
        End If

        Cursor = Cursors.WaitCursor

        Dim dsData As DataSet
        myClsBG0471BL.BudgetYear = CStr(Me.numYear.Value)
        myClsBG0471BL.Month = Me.cboMonth.Text
        myClsBG0471BL.UserLevelId = p_intUserLevelId

        If myClsBG0471BL.GetBudgetCompareData() = False Then
            dsData = Nothing
            MessageBox.Show("No budget data found, please try it again.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Cursor = Cursors.Default
            Return
        Else
            dsData = myClsBG0471BL.BudgetCompareData
            If dsData Is Nothing Or dsData.Tables(0).Rows.Count = 0 Then
                MessageBox.Show("No budget data found, please try it again.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Cursor = Cursors.Default
                Return
            End If
        End If

        Dim dtAuthorizeImages As DataTable = myClsBG0471BL.BudgetCompareData.Tables(1)
        Dim strYear As String = Me.numYear.Value.ToString
        Dim strMonth As String = Me.cboMonth.Text
        Dim dtColumns As DataTable = CreateTableTemplate()

        'Dim strPeriodType As String = cboPeriodType.Text
        'Dim strProjectNo As String = Me.numProjectNo.Value.ToString

        Dim strSubTitle As String = "Summary by Person In Charge : Budget Compare Year " + strYear

        InsertExcelColumnData(dtColumns, strYear, strMonth)
        CalculateDiff(strMonth, dsData)

        '//Create group data
        Dim intGroupFirstIndex As Integer = 0
        Dim intGroupSecondIndex As Integer = 0
        Dim dsGroups As DataSet = SetupCompareGroupbyData(strMonth, dsData, "PIC_SHOW_FLAG", "PIC_SHOW_FLAG", 3, False, intGroupFirstIndex, intGroupSecondIndex)

        '//Create Output Excel
        OutputExcel(dsGroups, dtColumns, False, strSubTitle, strYear, strMonth, False, intGroupFirstIndex, intGroupSecondIndex)


        Me.Cursor = Cursors.Default
    End Sub

#End Region

End Class