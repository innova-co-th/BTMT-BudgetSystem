Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.IO

Public Class frmBG0360

#Region "Variable"
    Private myClsBG0360BL As New clsBG0360BL
    Private myClsBG0310BL As New clsBG0310BL
    Private dtPeriodType As New DataTable
    Private dtAccount As New DataTable
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

    Private Sub LoadPeriodType()
        Try
            Me.cboPeriodType.Items.Clear()

            myClsBG0310BL.OpenPeriodFlg = "1"
            myClsBG0310BL.GetOpenPeriodList()

            If myClsBG0310BL.PeriodList IsNot Nothing AndAlso myClsBG0310BL.PeriodList.Rows.Count > 0 Then
                cboPeriodType.DisplayMember = "PERIOD_TYPE_NAME"
                cboPeriodType.ValueMember = "PERIOD_TYPE_ID"
                cboPeriodType.DataSource = myClsBG0310BL.PeriodList

                cboPeriodType.SelectedIndex = 0
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Detail by Account Code Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub init()
        'Dim periodType As BGConstant.enumPeriodType
        'Dim drTemp As DataRow

        'addDataColumn(dtPeriodType, New DataColumn("KEY"))
        'addDataColumn(dtPeriodType, New DataColumn("VALUE"))

        'drTemp = dtPeriodType.NewRow
        'drTemp("KEY") = "1"
        'drTemp("VALUE") = "Original Budget"
        'dtPeriodType.Rows.Add(drTemp)

        'drTemp = dtPeriodType.NewRow
        'drTemp("KEY") = "2"
        'drTemp("VALUE") = "Estimate Budget"
        'dtPeriodType.Rows.Add(drTemp)

        'drTemp = dtPeriodType.NewRow
        'drTemp("KEY") = "3"
        'drTemp("VALUE") = "Forecast Budget"
        'dtPeriodType.Rows.Add(drTemp)

        'fillComboBox(Me.cboPeriodType, dtPeriodType, "KEY", "VALUE", False)
        LoadPeriodType()


        '// Add item for account list
        If myClsBG0360BL.getAccountList Then
            dtAccount = myClsBG0360BL.DtResult
        End If
        ''fillComboBox(Me.cboAccountNo, dtAccount, "ACCOUNT_NO", "ACCOUNT_NAME_2", False)
        fillComboBox(Me.cboAccountNo, dtAccount, "ACCOUNT_NO", "ACCOUNT_NAME_2", True)

    End Sub

    Private Sub fillComboBox(ByVal cbo As ComboBox, _
                         ByVal dt As DataTable, _
                         ByVal keyColName As String, _
                         ByVal textColName As String, _
                         ByVal allowFirstEmptyItem As Boolean)
        Dim str As String = String.Empty

        cbo.Items.Clear()

        If dt.Rows.Count > 0 Then
            If allowFirstEmptyItem Then
                Dim dr As DataRow = dt.NewRow
                dr(keyColName) = ""
                dr(textColName) = "All"
                dt.Rows.InsertAt(dr, 0)
            End If
            cbo.DataSource = dt
            cbo.ValueMember = keyColName
            cbo.DisplayMember = textColName
        End If
        cbo.SelectedIndex = 0
    End Sub

    Private Sub addDataColumn(ByVal dt As DataTable, ByVal dc As DataColumn)
        Dim colIndex As Integer = dt.Columns.IndexOf(dc)
        If colIndex < 0 Then
            dt.Columns.Add(dc)
        End If
    End Sub
#End Region

#Region "Control Event"

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExport.Click
        Dim fileName As String = String.Empty
        Dim saveDialog As New SaveFileDialog
        Dim dtExport As New DataTable

        If Me.cboAccountNo.SelectedValue.ToString <> "" Then
            fileName = Me.cboAccountNo.SelectedValue.ToString
        Else
            fileName = "All"
        End If

        With saveDialog
            .FileName = "Export_" & _
                        fileName & "_" & _
                        Date.Now.ToString("yyyyMMdd") & ".txt"
            .Filter = "Text Files (*.txt)|*.txt"
        End With

        saveDialog.CheckPathExists = True

        If saveDialog.ShowDialog() = Windows.Forms.DialogResult.Cancel Then
            Exit Sub
        End If

        If saveDialog.FileName.Trim = "" Then
            Exit Sub
        End If

        Me.Enabled = False
        Me.Cursor = Cursors.WaitCursor

        fileName = System.IO.Path.GetFullPath(saveDialog.FileName)
        Dim strWtr As StreamWriter
        Dim colCount As Integer = 0

        Try
            myClsBG0360BL.BudgetYear = Me.numYear.Value.ToString
            myClsBG0360BL.PeriodType = Me.cboPeriodType.SelectedValue.ToString
            myClsBG0360BL.AccountNo = Me.cboAccountNo.SelectedValue.ToString
            myClsBG0360BL.ProjectNo = Me.numProjectNo.Value.ToString

            If myClsBG0360BL.getExportList Then
                dtExport = myClsBG0360BL.DtResult

                If dtExport.Rows.Count > 0 Then
                    strWtr = File.CreateText(fileName)

                    strWtr.Write("1" & vbTab)
                    For Each dc As DataColumn In dtExport.Columns
                        colCount += 1
                        If colCount = dtExport.Columns.Count Then
                            strWtr.Write(dc.ColumnName & vbNewLine)
                        Else
                            strWtr.Write(dc.ColumnName & vbTab)
                        End If
                    Next

                    colCount = 0
                    For Each dr As DataRow In dtExport.Rows
                        strWtr.Write("0" & vbTab)
                        For Each dc As DataColumn In dtExport.Columns
                            colCount += 1
                            '//-- Begin Edit 2011/04/29 by S.Watcharapong
                            If colCount = dtExport.Columns.Count Then
                                ''strWtr.Write(dr(dc.ColumnName).ToString & vbNewLine)
                                '//-- Begin Edit 2011/07/19 by S.Watcharapong 
                                If dc.ColumnName.StartsWith("Budget_") Or dc.ColumnName.StartsWith("MTP_") Then 'And Not IsDBNull(dr(dc.ColumnName)) Then
                                    '//-- End Edit 2011/07/19 
                                    strWtr.Write((CDbl(Nz(dr(dc.ColumnName), 0)) * 1000).ToString("0.0000") & vbNewLine)
                                Else
                                    strWtr.Write(dr(dc.ColumnName).ToString & vbNewLine)
                                End If
                            Else
                                ''strWtr.Write(dr(dc.ColumnName).ToString & vbTab)
                                '//-- Begin Edit 2011/07/19 by S.Watcharapong
                                If dc.ColumnName.StartsWith("Budget_") Or dc.ColumnName.StartsWith("MTP_") Then 'And Not IsDBNull(dr(dc.ColumnName)) Then
                                    '//-- End Edit 2011/07/19
                                    strWtr.Write((CDbl(Nz(dr(dc.ColumnName), 0)) * 1000).ToString("0.0000") & vbTab)
                                Else
                                    strWtr.Write(dr(dc.ColumnName).ToString & vbTab)
                                End If
                            End If
                            '//-- End Edit 2011/04/29
                        Next
                        colCount = 0
                    Next

                    strWtr.Close()

                    showSystemMessage("Export completed." & vbCrLf & "Exported file was stored at " & fileName)

                    '// Write Transaction Log
                    WriteTransactionLog(CStr(enumOperationCd.ExportData), myClsBG0360BL.BudgetYear, myClsBG0360BL.PeriodType, "", "", "", myClsBG0360BL.ProjectNo)

                Else
                    showSystemMessage("No item found with specified conditions. No items were exported.")
                End If
            End If

        Catch ex As Exception
            showErrorMessage("Export failed." & vbCrLf & ex.Message)
        End Try

        Me.Enabled = True
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub frmBG0360_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        init()
    End Sub

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