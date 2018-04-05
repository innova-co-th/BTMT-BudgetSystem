Imports BudgetManual.BGConstant
Imports Microsoft.Office.Interop
Imports System.IO

Public Class BGCommon

#Region "Variable"

    '// Public Variables
    Public Shared p_strUserId As String = String.Empty
    Public Shared p_strUserName As String = String.Empty
    Public Shared p_intUserLevelId As Integer
    Public Shared p_intUserLevelName As String = String.Empty
    Public Shared p_strUserPIC As String = String.Empty
    Public Shared p_blnSendAutoMail As Boolean = False
    Public Shared p_strAutoMailFromAddr As String = String.Empty
    Public Shared p_strDataPath As String = String.Empty
    Public Shared p_strAppPath As String = String.Empty
    '//-- Begin Edit 2011/08/25 S.Watcharapong
    Public Shared p_blnReadOnlyMode As Boolean = False
    '//-- End Edit 2011/08/25

    '// Screens
    Public Shared p_frmBG0010 As frmBG0010 = Nothing
    Public Shared p_frmBG0110 As frmBG0110 = Nothing
    Public Shared p_frmBG0200 As New List(Of frmBG0200)
    Public Shared p_frmBG0201 As New frmBG0201
    Public Shared p_frmBG0310 As frmBG0310 = Nothing
    Public Shared p_frmBG0320 As frmBG0320 = Nothing
    Public Shared p_frmBG0330 As frmBG0330 = Nothing
    Public Shared p_frmBG0350 As frmBG0350 = Nothing
    Public Shared p_frmBG0360 As frmBG0360 = Nothing
    Public Shared p_frmBG0380 As frmBG0380 = Nothing
    Public Shared p_frmBG0390 As frmBG0390 = Nothing
    Public Shared p_frmBG0395 As frmBG0395 = Nothing
    Public Shared p_frmBG0401 As frmBG0401 = Nothing
    Public Shared p_frmBG0410 As frmBG0410 = Nothing
    Public Shared p_frmBG0420 As frmBG0420 = Nothing
    Public Shared p_frmBG0430 As frmBG0430 = Nothing
    Public Shared p_frmBG0440 As frmBG0440 = Nothing
    Public Shared p_frmBG0450 As frmBG0450 = Nothing
    Public Shared p_frmBG0460 As frmBG0460 = Nothing
    Public Shared p_frmBG0470 As frmBG0470 = Nothing
    Public Shared p_frmBG0471 As frmBG0471 = Nothing
    Public Shared p_frmBG0472 As frmBG0472 = Nothing
    Public Shared p_frmBG0473 As frmBG0473 = Nothing
    Public Shared p_frmBG0474 As frmBG0474 = Nothing
    Public Shared p_frmBG0480 As frmBG0480 = Nothing
    Public Shared p_frmBG0610 As frmBG0610 = Nothing
    Public Shared p_frmBG0611 As frmBG0611 = Nothing
    Public Shared p_frmBG0620 As frmBG0620 = Nothing
    Public Shared p_frmBG0630 As frmBG0630 = Nothing
    Public Shared p_frmBG0640 As frmBG0640 = Nothing
    Public Shared p_frmBG0650 As frmBG0650 = Nothing
    Public Shared p_frmBG0660 As frmBG0660 = Nothing
    Public Shared p_frmBG0670 As frmBG0670 = Nothing
    Public Shared p_frmBG0680 As frmBG0680 = Nothing
    Public Shared p_frmBG0681 As frmBG0681 = Nothing
    Public Shared p_frmBG0682 As frmBG0682 = Nothing
    Public Shared p_frmBG0690 As frmBG0690 = Nothing
    Public Shared p_frmBG0710 As frmBG0710 = Nothing
    Public Shared p_frmBG0720 As frmBG0720 = Nothing
    Public Shared p_frmBG0730 As frmBG0730 = Nothing

#End Region

#Region "Function"
    ''' <summary>
    ''' Convert object to empty string if is DBNULL
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Shared Function Nz(ByVal objValue As Object, Optional ByVal strDefault As Object = "") As Object
        If objValue Is Nothing OrElse IsDBNull(objValue) Then
            Return strDefault
        Else
            Return CStr(objValue)
        End If
    End Function

    ''' <summary>
    ''' readXMLConfig
    ''' </summary>
    ''' <param name="FilePath"></param>
    ''' <param name="SectionName"></param>
    ''' <param name="ValueName"></param>
    ''' <param name="DefaultValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function readXMLConfig(ByVal FilePath As String, ByVal SectionName As String, _
                                         ByVal ValueName As String, Optional ByVal DefaultValue As String = "") As String
        Dim ds As New Data.DataSet
        Dim dt As Data.DataTable = Nothing
        Dim strRetValue As String = DefaultValue

        Try
            ds.ReadXml(FilePath)
            dt = ds.Tables(SectionName)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Try
                    strRetValue = CStr(dt.Rows(0).Item(ValueName))
                Catch ex As Exception
                    strRetValue = DefaultValue
                End Try
            End If
            Return strRetValue

        Catch ex As Exception
            Debug.Print("[readXMLConfig] Error: " & ex.Message)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' SaveXMLConfig
    ''' </summary>
    ''' <param name="FilePath"></param>
    ''' <param name="SectionName"></param>
    ''' <param name="ValueName"></param>
    ''' <param name="Value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function SaveXMLConfig(ByVal FilePath As String, ByVal SectionName As String, _
                                         ByVal ValueName As String, ByVal Value As String) As Boolean
        Dim ds As New DataSet

        Try
            ds.ReadXml(FilePath)
            If ds.Tables(SectionName).Columns.Contains(ValueName) = False Then
                ds.Tables(SectionName).Columns.Add(ValueName)
            End If
            ds.Tables(SectionName).Rows(0).Item(ValueName) = Value

            ds.WriteXml(FilePath)
            Return True

        Catch ex As Exception
            Debug.Print("[saveXMLConfig] Error: " & ex.Message)
            Return False
        End Try
    End Function

    Public Shared Function Equate(ByVal str As String) As String
        Dim formula As String = "({0})"
        Dim expr As String = String.Format(formula, str)

        Try
            Return New DataTable().Compute(expr, "").ToString()

        Catch ex As Exception
            Return Nothing

        End Try
    End Function

    ''' <summary>
    ''' Show general program message with information icon
    ''' </summary>
    ''' <param name="message">Message to display</param>
    ''' <remarks></remarks>
    Public Shared Sub showSystemMessage(ByVal message As String)
        MessageBox.Show(message, My.Settings.ProgramTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ''' <summary>
    ''' Show general program message with error icon
    ''' </summary>
    ''' <param name="message">Error message to display</param>
    ''' <remarks></remarks>
    Public Shared Sub showErrorMessage(ByVal message As String)
        MessageBox.Show(message, My.Settings.ProgramTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    ''' <summary>
    ''' Show confirmation message, return Yes or No
    ''' </summary>
    ''' <param name="message">Confirmation message to display</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function showConfirmMessage(ByVal message As String) As Windows.Forms.DialogResult
        Dim result As Windows.Forms.DialogResult
        result = MessageBox.Show(message, My.Settings.ProgramTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        Return result
    End Function

    Public Shared Sub WriteTransactionLog(ByVal LogOperationCd As String, ByVal BudgetYear As String, ByVal PeriodType As String, _
                                    ByVal UserPIC As String, ByVal BudgetType As String, ByVal RevNo As String, ByVal ProjectNo As String)
        Dim clsBG_T_TRANS_LOG As New BG_T_TRANS_LOG

        clsBG_T_TRANS_LOG.UserId = p_strUserId
        clsBG_T_TRANS_LOG.OperationCd = LogOperationCd
        clsBG_T_TRANS_LOG.BudgetYear = BudgetYear
        clsBG_T_TRANS_LOG.PeriodType = PeriodType
        clsBG_T_TRANS_LOG.UserPIC = UserPIC
        clsBG_T_TRANS_LOG.BudgetType = BudgetType
        clsBG_T_TRANS_LOG.RevNo = RevNo
        clsBG_T_TRANS_LOG.ProjectNo = ProjectNo

        clsBG_T_TRANS_LOG.Insert001()
    End Sub

    Public Shared Function SetupGroupbyData(ByVal dsData As DataSet, ByVal strGroupColumnName As String, ByVal strGroupColumnTitle As String, ByVal intDataColumnIndex As Integer, ByVal bShowGroupName As Boolean) As DataSet

        Dim dsResult As DataSet = New DataSet

        Dim strScript As String = strGroupColumnName
        'Dim strGroupbyScript As String = "Group by PERSON_IN_CHARGE_NO"
        'Dim arrGroups As DataRow() = dsData.Tables(0).Select(strScript)

        Dim strSort As String = strGroupColumnName + " DESC"

        '//Get groups list
        Dim dtTmp As DataTable = dsData.Tables(0).DefaultView.ToTable(True, strScript)

        '//Sort groups list by group column name
        Dim dtGroups As DataTable = dtTmp.Clone
        Dim arrTmp As DataRow() = dtTmp.Select("", strSort)
        For intTmp As Integer = 0 To arrTmp.Length - 1
            Dim drow(dtGroups.Columns.Count - 1) As Object
            arrTmp(intTmp).ItemArray.CopyTo(drow, 0)
            dtGroups.Rows.Add(drow)
        Next

        Dim intGroupCount As Integer = dtGroups.Rows.Count
        For i As Integer = 0 To intGroupCount - 1

            '//Seperate dataset data into several datatables according to group no
            strScript = strGroupColumnName + " = '" + dtGroups.Rows(i)(0).ToString & "'"
            Dim arrRows As DataRow() = dsData.Tables(0).Select(strScript)

            Dim dtGroupTmp As DataTable = dsData.Tables(0).Clone
            For j As Integer = 0 To arrRows.Length - 1
                Dim drow(dtGroupTmp.Columns.Count - 1) As Object
                arrRows(j).ItemArray.CopyTo(drow, 0)
                dtGroupTmp.Rows.Add(drow)
            Next

            '//Calculate total for each group
            Dim strExpression As String
            Dim strFilter As String = String.Empty

            Dim drTotal As DataRow = dtGroupTmp.NewRow
            Dim drFixcostTotal As DataRow = dtGroupTmp.NewRow
            Dim drVariablecostTotal As DataRow = dtGroupTmp.NewRow

            Dim returnValue As Object
            For k As Integer = intDataColumnIndex To dtGroupTmp.Columns.Count - 1

                Dim strColumnName As String = dtGroupTmp.Columns(k).ColumnName
                strExpression = "Sum(" + strColumnName + ")"
                returnValue = dtGroupTmp.Compute(strExpression, strFilter)
                drTotal(dtGroupTmp.Columns(k).ColumnName) = returnValue

                If strColumnName.IndexOf("FIXCOST") < 0 And strColumnName.IndexOf("VARIABLECOST") < 0 Then

                    If strColumnName.IndexOf("SUM") > 0 And strColumnName <> "LAST_YEAR_SUM" Then
                        Dim intSumIndex As Integer = strColumnName.IndexOf("SUM")
                        strColumnName = strColumnName.Substring(0, intSumIndex - 1)
                    End If
                    strExpression = "Sum(" + strColumnName + "_FIXCOST)"
                    returnValue = dtGroupTmp.Compute(strExpression, strFilter)
                    drFixcostTotal(dtGroupTmp.Columns(k).ColumnName) = returnValue

                    strExpression = "Sum(" + strColumnName + "_VARIABLECOST)"
                    returnValue = dtGroupTmp.Compute(strExpression, strFilter)
                    drVariablecostTotal(dtGroupTmp.Columns(k).ColumnName) = returnValue
                End If

            Next

            '//Add total cost
            dtGroupTmp.Rows.Add(drTotal)

            '//Add one empty row
            Dim drEmpty As DataRow = dtGroupTmp.NewRow
            dtGroupTmp.Rows.Add(drEmpty)

            '//Add variable cost total
            dtGroupTmp.Rows.Add(drVariablecostTotal)

            '//Add fixed cost total
            dtGroupTmp.Rows.Add(drFixcostTotal)

            If bShowGroupName = True Then
                dtGroupTmp.TableName = arrRows(0)(strGroupColumnName).ToString & " " & arrRows(0)(strGroupColumnTitle).ToString
            End If

            dsResult.Tables.Add(dtGroupTmp)
        Next

        Return dsResult

    End Function

    Public Shared Function SetupOriginalColumnsCells(ByRef xSt As Excel.Worksheet, ByVal colStartIndex As Integer, _
                                    ByVal intColMergeStart As Integer, ByVal intColMergeEnd As Integer, _
                                    ByVal strColValue As String, ByVal arrColRowList() As Integer, _
                                    ByVal intFirstHalfStart As Integer, ByVal intFirstHalfEnd As Integer, _
                                    ByVal strYear As String, Optional ByVal bMergeTwoColumn As Boolean = True, _
                                    Optional ByVal bMergeSecondHalfColumn As Boolean = False, _
                                    Optional ByVal intSecondHalfStart As Integer = 1, _
                                    Optional ByVal intSecondHalfEnd As Integer = 1) As Boolean

        Dim strHalfYear As String = strYear.Substring(2, 2)
        If bMergeTwoColumn = True Then
            '//Merge Column1 & Column2
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).ClearContents()
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).MergeCells = True
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).Value = strColValue
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        End If

        '//Merge two columns row
        For i As Integer = 0 To arrColRowList.Length - 1
            MergeColumnsCells(xSt, arrColRowList(i), colStartIndex - 1, colStartIndex)
        Next

        '//Setup first half year title
        xSt.Cells(colStartIndex - 1, intFirstHalfStart) = "1st Half'" & strHalfYear
        xSt.Range(xSt.Cells(colStartIndex - 1, intFirstHalfStart), xSt.Cells(colStartIndex - 1, intFirstHalfEnd)).MergeCells = True
        xSt.Range(xSt.Cells(colStartIndex - 1, intFirstHalfStart), xSt.Cells(colStartIndex - 1, intFirstHalfEnd)).Font.Bold = True
        xSt.Range(xSt.Cells(colStartIndex - 1, intFirstHalfStart), xSt.Cells(colStartIndex - 1, intFirstHalfEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

        If bMergeSecondHalfColumn = True Then
            xSt.Cells(colStartIndex - 1, intSecondHalfStart) = "2nd Half'" & strHalfYear
            xSt.Range(xSt.Cells(colStartIndex - 1, intSecondHalfStart), xSt.Cells(colStartIndex - 1, intSecondHalfEnd)).MergeCells = True
            xSt.Range(xSt.Cells(colStartIndex - 1, intSecondHalfStart), xSt.Cells(colStartIndex - 1, intSecondHalfEnd)).Font.Bold = True
            xSt.Range(xSt.Cells(colStartIndex - 1, intSecondHalfStart), xSt.Cells(colStartIndex - 1, intSecondHalfEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        End If

        Return True

    End Function

    Public Shared Function SetupEstimateColumnsCells(ByRef xSt As Excel.Worksheet, ByVal colStartIndex As Integer, _
                                    ByVal intColMergeStart As Integer, ByVal intColMergeEnd As Integer, _
                                    ByVal strColValue As String, ByVal arrColRowList() As Integer, _
                                    ByVal intActualStart As Integer, ByVal intActualEnd As Integer, _
                                    ByVal intEstimateStart As Integer, ByVal intEstimateEnd As Integer, _
                                    Optional ByVal bMergeTwoColumn As Boolean = True) As Boolean

        If bMergeTwoColumn = True Then
            '//Merge Column1 & Column2
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).ClearContents()
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).MergeCells = True
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).Value = strColValue
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        End If

        '//Merge two columns row
        For i As Integer = 0 To arrColRowList.Length - 1
            MergeColumnsCells(xSt, arrColRowList(i), colStartIndex - 1, colStartIndex)
        Next

        '//Setup Actual & Estimate Title
        xSt.Cells(colStartIndex - 1, intActualStart) = "Actual"
        xSt.Range(xSt.Cells(colStartIndex - 1, intActualStart), xSt.Cells(colStartIndex - 1, intActualEnd)).MergeCells = True
        xSt.Range(xSt.Cells(colStartIndex - 1, intActualStart), xSt.Cells(colStartIndex - 1, intActualEnd)).Font.Bold = True
        xSt.Range(xSt.Cells(colStartIndex - 1, intActualStart), xSt.Cells(colStartIndex - 1, intActualEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

        xSt.Cells(colStartIndex - 1, intEstimateStart) = "Estimate"
        xSt.Range(xSt.Cells(colStartIndex - 1, intEstimateStart), xSt.Cells(colStartIndex - 1, intEstimateEnd)).MergeCells = True
        xSt.Range(xSt.Cells(colStartIndex - 1, intEstimateStart), xSt.Cells(colStartIndex - 1, intEstimateEnd)).Font.Bold = True
        xSt.Range(xSt.Cells(colStartIndex - 1, intEstimateStart), xSt.Cells(colStartIndex - 1, intEstimateEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

        Return True

    End Function

    Public Shared Function SetupReviseColumnsCells(ByRef xSt As Excel.Worksheet, ByVal colStartIndex As Integer, ByVal bMTPCheck As Boolean, _
                                    ByVal intColMergeStart As Integer, ByVal intColMergeEnd As Integer, _
                                    ByVal strColValue As String, ByVal arrColRowList() As Integer, _
                                    ByVal intActualStart As Integer, ByVal intActualEnd As Integer, _
                                    ByVal intEstimateStart As Integer, ByVal intEstimateEnd As Integer, _
                                    ByVal intReviseStart As Integer, ByVal intReviseEnd As Integer, _
                                    ByVal intMTPReviseStart As Integer, ByVal intMTPReviseEnd As Integer, _
                                    ByVal intMTPStart As Integer, ByVal intMTPEnd As Integer, Optional ByVal bMergeTwoColumn As Boolean = True) As Boolean

        If bMergeTwoColumn = True Then
            '//Merge Column1 & Column2
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).ClearContents()
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).MergeCells = True
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).Value = strColValue
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        End If

        '//Merge two columns row
        For i As Integer = 0 To arrColRowList.Length - 1
            MergeColumnsCells(xSt, arrColRowList(i), colStartIndex - 1, colStartIndex)
        Next
        If bMTPCheck = False Then

            '//Setup Actual, Estimate & Revise Title
            xSt.Cells(colStartIndex - 1, intActualStart) = "Actual"
            xSt.Range(xSt.Cells(colStartIndex - 1, intActualStart), xSt.Cells(colStartIndex - 1, intActualEnd)).MergeCells = True
            xSt.Range(xSt.Cells(colStartIndex - 1, intActualStart), xSt.Cells(colStartIndex - 1, intActualEnd)).Font.Bold = True
            xSt.Range(xSt.Cells(colStartIndex - 1, intActualStart), xSt.Cells(colStartIndex - 1, intActualEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            xSt.Cells(colStartIndex - 1, intEstimateStart) = "Estimate"
            xSt.Range(xSt.Cells(colStartIndex - 1, intEstimateStart), xSt.Cells(colStartIndex - 1, intEstimateEnd)).MergeCells = True
            xSt.Range(xSt.Cells(colStartIndex - 1, intEstimateStart), xSt.Cells(colStartIndex - 1, intEstimateEnd)).Font.Bold = True
            xSt.Range(xSt.Cells(colStartIndex - 1, intEstimateStart), xSt.Cells(colStartIndex - 1, intEstimateEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            xSt.Cells(colStartIndex - 1, intReviseStart) = "Revise"
            xSt.Range(xSt.Cells(colStartIndex - 1, intReviseStart), xSt.Cells(colStartIndex - 1, intReviseEnd)).MergeCells = True
            xSt.Range(xSt.Cells(colStartIndex - 1, intReviseStart), xSt.Cells(colStartIndex - 1, intReviseEnd)).Font.Bold = True
            xSt.Range(xSt.Cells(colStartIndex - 1, intReviseStart), xSt.Cells(colStartIndex - 1, intReviseEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        Else

            xSt.Cells(colStartIndex - 1, intActualStart) = "Revise"
            xSt.Range(xSt.Cells(colStartIndex - 1, intActualStart), xSt.Cells(colStartIndex - 1, intEstimateEnd)).MergeCells = True
            xSt.Range(xSt.Cells(colStartIndex - 1, intActualStart), xSt.Cells(colStartIndex - 1, intEstimateEnd)).Font.Bold = True
            xSt.Range(xSt.Cells(colStartIndex - 1, intActualStart), xSt.Cells(colStartIndex - 1, intEstimateEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

        End If

        If bMTPCheck = True Then

            'xSt.Cells(colStartIndex - 1, intMTPReviseStart) = "Revise"
            'xSt.Range(xSt.Cells(colStartIndex - 1, intMTPReviseStart), xSt.Cells(colStartIndex - 1, intMTPReviseEnd)).MergeCells = True
            'xSt.Range(xSt.Cells(colStartIndex - 1, intMTPReviseStart), xSt.Cells(colStartIndex - 1, intMTPReviseEnd)).Font.Bold = True
            'xSt.Range(xSt.Cells(colStartIndex - 1, intMTPReviseStart), xSt.Cells(colStartIndex - 1, intMTPReviseEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            'xSt.Cells(colStartIndex - 1, intMTPStart) = "MTP Budget"
            'xSt.Range(xSt.Cells(colStartIndex - 1, intMTPStart), xSt.Cells(colStartIndex - 1, intMTPEnd)).MergeCells = True
            'xSt.Range(xSt.Cells(colStartIndex - 1, intMTPStart), xSt.Cells(colStartIndex - 1, intMTPEnd)).Font.Bold = True
            'xSt.Range(xSt.Cells(colStartIndex - 1, intMTPStart), xSt.Cells(colStartIndex - 1, intMTPEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

        End If

        Return True

    End Function

    Public Shared Function SetupMTPColumnsCells(ByRef xSt As Excel.Worksheet, ByVal colStartIndex As Integer, _
                                   ByVal intColMergeStart As Integer, ByVal intColMergeEnd As Integer, _
                                   ByVal strColValue As String, ByVal arrColRowList() As Integer, _
                                   ByVal intMTPStart As Integer, ByVal intMTPEnd As Integer, Optional ByVal bMergeTwoColumn As Boolean = True) As Boolean

        If bMergeTwoColumn = True Then
            '//Merge Column1 & Column2
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).ClearContents()
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).MergeCells = True
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).Value = strColValue
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        End If

        '//Merge two columns row
        For i As Integer = 0 To arrColRowList.Length - 1
            MergeColumnsCells(xSt, arrColRowList(i), colStartIndex - 1, colStartIndex)
        Next

        xSt.Cells(colStartIndex - 1, intMTPStart) = "MTP Budget"
        xSt.Range(xSt.Cells(colStartIndex - 1, intMTPStart), xSt.Cells(colStartIndex - 1, intMTPEnd)).MergeCells = True
        xSt.Range(xSt.Cells(colStartIndex - 1, intMTPStart), xSt.Cells(colStartIndex - 1, intMTPEnd)).Font.Bold = True
        xSt.Range(xSt.Cells(colStartIndex - 1, intMTPStart), xSt.Cells(colStartIndex - 1, intMTPEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

    End Function

    Public Shared Function SetupCompareColumnsCells(ByRef xSt As Excel.Worksheet, ByVal colStartIndex As Integer, _
                                    ByVal intMonth As Integer, _
                                    ByVal intColMergeStart As Integer, ByVal intColMergeEnd As Integer, _
                                    ByVal strColValue As String, ByVal arrColRowList() As Integer, _
                                    ByVal intMonthStart As Integer, ByVal intMonthEnd As Integer, _
                                    ByVal intHalfStart As Integer, ByVal intHalfEnd As Integer, _
                                    ByVal strYear As String, Optional ByVal bMergeTwoColumn As Boolean = True, _
                                    Optional ByVal bMergeYearColumn As Boolean = False, _
                                    Optional ByVal intYearStart As Integer = 1, _
                                    Optional ByVal intYearEnd As Integer = 1) As Boolean

        Dim strMonthName As String = MonthName(intMonth)
        Dim strAccMonth As String = String.Empty

        If bMergeTwoColumn = True Then
            '//Merge Column1 & Column2
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).ClearContents()
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).MergeCells = True
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).Font.Bold = True
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).Value = strColValue
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            xSt.Range(xSt.Cells(colStartIndex - 1, intColMergeStart), xSt.Cells(colStartIndex, intColMergeEnd)).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        End If

        '//Merge two columns row
        For i As Integer = 0 To arrColRowList.Length - 1
            MergeColumnsCells(xSt, arrColRowList(i), colStartIndex - 1, colStartIndex)
        Next

        '//Setup month title
        xSt.Cells(colStartIndex - 1, intMonthStart) = strMonthName & "'" & strYear
        xSt.Range(xSt.Cells(colStartIndex - 1, intMonthStart), xSt.Cells(colStartIndex - 1, intMonthEnd)).MergeCells = True
        xSt.Range(xSt.Cells(colStartIndex - 1, intMonthStart), xSt.Cells(colStartIndex - 1, intMonthEnd)).Font.Bold = True
        xSt.Range(xSt.Cells(colStartIndex - 1, intMonthStart), xSt.Cells(colStartIndex - 1, intMonthEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

        '//Setup half year title
        If intMonth <= 6 Then
            'If intMonth = 1 Then
            '    strAccMonth = MonthName(intMonth)
            'Else
            '    strAccMonth = MonthName(1) & "-" & MonthName(intMonth)
            'End If

            'xSt.Cells(colStartIndex - 1, intHalfStart) = "Accumulation 1st Half'" & strYear & " (" & strAccMonth & "'" & strYear & " )"
            xSt.Cells(colStartIndex - 1, intHalfStart) = "Accumulation 1st Half'" & strYear 
        Else
            'If intMonth = 7 Then
            '    strAccMonth = MonthName(intMonth)
            'Else
            '    strAccMonth = MonthName(7) & "-" & MonthName(intMonth)
            'End If

            'xSt.Cells(colStartIndex - 1, intHalfStart) = "Accumulation 2nd Half'" & strYear & " (" & strAccMonth & "'" & strYear & " )"
            xSt.Cells(colStartIndex - 1, intHalfStart) = "Accumulation 2nd Half'" & strYear
        End If
        xSt.Range(xSt.Cells(colStartIndex - 1, intHalfStart), xSt.Cells(colStartIndex - 1, intHalfEnd)).MergeCells = True
        xSt.Range(xSt.Cells(colStartIndex - 1, intHalfStart), xSt.Cells(colStartIndex - 1, intHalfEnd)).Font.Bold = True
        xSt.Range(xSt.Cells(colStartIndex - 1, intHalfStart), xSt.Cells(colStartIndex - 1, intHalfEnd)).WrapText = True
        xSt.Range(xSt.Cells(colStartIndex - 1, intHalfStart), xSt.Cells(colStartIndex - 1, intHalfEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

        '//Setup year title
        If bMergeYearColumn = True Then
            xSt.Cells(colStartIndex - 1, intYearStart) = "Total Year End"
            xSt.Range(xSt.Cells(colStartIndex - 1, intYearStart), xSt.Cells(colStartIndex - 1, intYearEnd)).MergeCells = True
            xSt.Range(xSt.Cells(colStartIndex - 1, intYearStart), xSt.Cells(colStartIndex - 1, intYearEnd)).Font.Bold = True
            xSt.Range(xSt.Cells(colStartIndex - 1, intYearStart), xSt.Cells(colStartIndex - 1, intYearEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        End If

        Return True

    End Function

    Public Shared Function SetupExcelTitle(ByRef xSt As Excel.Worksheet, ByVal strSubTitle As String, _
                                           ByVal strYear As String, ByVal bMTPCheck As Boolean, _
                                           ByVal intUnitPriceStart As Integer, ByVal intUnitPriceEnd As Integer, _
                                           ByVal intAuthorizeStart As Integer, ByVal intAuthorizeEnd As Integer, _
                                           ByVal intImageLogoIndex As Integer, ByVal bShowGroupName As Boolean, _
                                           Optional ByVal strGroupName As String = "", _
                                           Optional ByVal bAuthorize1twocols As Boolean = False) As Boolean

        '//Setup Title & Title Font 
        xSt.Range(xSt.Cells(2, 1), xSt.Cells(2, 5)).Font.Name = "Tahoma"
        xSt.Range(xSt.Cells(2, 1), xSt.Cells(2, 5)).Font.Size = 11
        xSt.Range(xSt.Cells(2, 1), xSt.Cells(2, 5)).Font.Bold = True
        xSt.Range(xSt.Cells(2, 1), xSt.Cells(2, 5)).MergeCells = True
        xSt.Range(xSt.Cells(2, 1), xSt.Cells(2, 5)).Value = "Bridgestone Tire Manufacturing (Thailand) Co.,Ltd."

        '//Setup subTitle  
        xSt.Range(xSt.Cells(3, 1), xSt.Cells(3, 5)).Font.Name = "Tahoma"
        xSt.Range(xSt.Cells(3, 1), xSt.Cells(3, 5)).Font.Size = 11
        xSt.Range(xSt.Cells(3, 1), xSt.Cells(3, 5)).Font.Bold = True
        xSt.Range(xSt.Cells(3, 1), xSt.Cells(3, 5)).MergeCells = True
        xSt.Range(xSt.Cells(3, 1), xSt.Cells(3, 5)).Value = strSubTitle

        xSt.Range(xSt.Cells(2, 1), xSt.Cells(3, 5)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

        If bShowGroupName = True Then
            '//Setup GroupName  
            'xSt.Range(excelApp.Cells(6, 1), excelApp.Cells(6, 3)).Font.Italic = True
            'xSt.Range(excelApp.Cells(6, 1), excelApp.Cells(6, 3)).Font.Underline = True
            xSt.Range(xSt.Cells(6, 1), xSt.Cells(6, 3)).Font.Name = "Tahoma"
            xSt.Range(xSt.Cells(6, 1), xSt.Cells(6, 3)).Font.Bold = True
            xSt.Range(xSt.Cells(6, 1), xSt.Cells(6, 3)).Font.Size = 11
            xSt.Range(xSt.Cells(6, 1), xSt.Cells(6, 3)).MergeCells = True
            xSt.Range(xSt.Cells(6, 1), xSt.Cells(6, 3)).Value = strGroupName
            xSt.Range(xSt.Cells(6, 1), xSt.Cells(6, 3)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        End If

        '//Setup unit price        
        'xSt.Range(excelApp.Cells(6, 23), excelApp.Cells(6, 24)).Font.Italic = True
        xSt.Range(xSt.Cells(6, intUnitPriceStart), xSt.Cells(6, intUnitPriceEnd)).Font.Name = "Tahoma"
        xSt.Range(xSt.Cells(6, intUnitPriceStart), xSt.Cells(6, intUnitPriceEnd)).Font.Underline = True
        xSt.Range(xSt.Cells(6, intUnitPriceStart), xSt.Cells(6, intUnitPriceEnd)).Font.Bold = True
        xSt.Range(xSt.Cells(6, intUnitPriceStart), xSt.Cells(6, intUnitPriceEnd)).Font.Size = 11
        xSt.Range(xSt.Cells(6, intUnitPriceStart), xSt.Cells(6, intUnitPriceEnd)).MergeCells = True
        xSt.Range(xSt.Cells(6, intUnitPriceStart), xSt.Cells(6, intUnitPriceEnd)).Value = "Unit : K.Baht"
        xSt.Range(xSt.Cells(6, intUnitPriceStart), xSt.Cells(6, intUnitPriceEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

        ''//Add authorize image
        'If bAuthorize1twocols = False Then
        '    xSt.Range(xSt.Cells(2, intAuthorizeStart), xSt.Cells(2, intAuthorizeStart)).Font.Name = "Tahoma"
        '    xSt.Range(xSt.Cells(2, intAuthorizeStart), xSt.Cells(2, intAuthorizeStart)).Font.Size = 11
        '    xSt.Range(xSt.Cells(2, intAuthorizeStart), xSt.Cells(2, intAuthorizeStart)).Font.Bold = True
        '    xSt.Range(xSt.Cells(2, intAuthorizeStart), xSt.Cells(2, intAuthorizeStart)).Value = "BTMT10"
        '    xSt.Range(xSt.Cells(2, intAuthorizeStart), xSt.Cells(2, intAuthorizeStart)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        'Else
        '    xSt.Range(xSt.Cells(2, intAuthorizeStart), xSt.Cells(2, intAuthorizeStart + 1)).ClearContents()
        '    xSt.Range(xSt.Cells(2, intAuthorizeStart), xSt.Cells(2, intAuthorizeStart + 1)).MergeCells = True
        '    xSt.Range(xSt.Cells(2, intAuthorizeStart), xSt.Cells(2, intAuthorizeStart + 1)).Font.Name = "Tahoma"
        '    xSt.Range(xSt.Cells(2, intAuthorizeStart), xSt.Cells(2, intAuthorizeStart + 1)).Font.Size = 11
        '    xSt.Range(xSt.Cells(2, intAuthorizeStart), xSt.Cells(2, intAuthorizeStart + 1)).Font.Bold = True
        '    xSt.Range(xSt.Cells(2, intAuthorizeStart), xSt.Cells(2, intAuthorizeStart + 1)).Value = "BTMT10"
        '    xSt.Range(xSt.Cells(2, intAuthorizeStart), xSt.Cells(2, intAuthorizeStart + 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        'End If

        'xSt.Range(xSt.Cells(2, intAuthorizeEnd), xSt.Cells(2, intAuthorizeEnd)).Font.Name = "Tahoma"
        'xSt.Range(xSt.Cells(2, intAuthorizeEnd), xSt.Cells(2, intAuthorizeEnd)).Font.Size = 11
        'xSt.Range(xSt.Cells(2, intAuthorizeEnd), xSt.Cells(2, intAuthorizeEnd)).Font.Bold = True
        'xSt.Range(xSt.Cells(2, intAuthorizeEnd), xSt.Cells(2, intAuthorizeEnd)).Value = "BTMT3"
        'xSt.Range(xSt.Cells(2, intAuthorizeEnd), xSt.Cells(2, intAuthorizeEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

        'If bAuthorize1twocols = False Then
        '    xSt.Range(xSt.Cells(3, intAuthorizeStart), xSt.Cells(6, intAuthorizeStart)).MergeCells = True
        'Else
        '    xSt.Range(xSt.Cells(3, intAuthorizeStart), xSt.Cells(6, intAuthorizeStart + 1)).MergeCells = True
        'End If

        'xSt.Range(xSt.Cells(3, intAuthorizeEnd), xSt.Cells(6, intAuthorizeEnd)).MergeCells = True
        'xSt.Range(xSt.Cells(2, intAuthorizeStart), xSt.Cells(6, intAuthorizeEnd)).Borders.LineStyle = 1

        '//Add Logo
        'Dim strPath As String = p_strAppPath
        'Dim intIndex As Integer = strPath.IndexOf("bin")
        'strPath = p_strAppPath.Substring(0, intIndex - 1)
        'Dim imgPath As String = strPath & "\Images\bridgestone_logo.jpg"

        'Dim bytes() As Byte
        'Dim ms As MemoryStream
        'Dim img As Image
        'Dim objImage1 As Object = dtAuthorizeImages.Rows(0)(0)
        'If Not objImage1 Is DBNull.Value Then
        '    bytes = CType(objImage1, Byte())
        '    ms = New MemoryStream(bytes)
        '    ms.Position = 0
        '    img = Image.FromStream(ms)
        '    ms.Close()

        '    xSt.Paste(xSt.Range(xSt.Cells(3, intAuthorizeStart), xSt.Cells(6, intAuthorizeStart)), img)
        'End If

        'Dim objImage2 As Object = dtAuthorizeImages.Rows(0)(1)
        'If Not objImage2 Is DBNull.Value Then
        '    bytes = CType(objImage2, Byte())
        '    ms = New MemoryStream(bytes)
        '    ms.Position = 0
        '    img = Image.FromStream(ms)
        '    ms.Close()
        'End If

        'Dim img As Image = Image.FromFile(imgPath)
        'xSt.Paste(xSt.Range(excelApp.Cells(2, 10), excelApp.Cells(10, 24)), img)
        '//Microsoft.Office.Core.MsoTriState.msoFalse
        '//Microsoft.Office.Core.MsoTriState.msoTrue
        'xSt.Shapes.AddPicture(imgPath, 0, 1, intImageLogoIndex, 10, 150, 24)

    End Function

    Public Shared Function MergeColumnsCells(ByVal xSt As Excel.Worksheet, ByVal colNumber As Integer, ByVal colStart As Integer, ByVal colEnd As Integer) As Boolean

        xSt.Range(xSt.Cells(colStart, colNumber), xSt.Cells(colEnd, colNumber)).MergeCells = True
        xSt.Range(xSt.Cells(colStart, colNumber), xSt.Cells(colEnd, colNumber)).WrapText = True
        xSt.Range(xSt.Cells(colStart, colNumber), xSt.Cells(colEnd, colNumber)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        xSt.Range(xSt.Cells(colStart, colNumber), xSt.Cells(colEnd, colNumber)).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

        Return True

    End Function

    Public Shared Function SetupTotalLines(ByRef xSt As Excel.Worksheet, _
                                           ByVal intTotalIndex As Integer, ByVal strTotalTitle As String, _
                                           ByVal strTotalAlign As String, ByVal intEmptyStart As Integer, _
                                           ByVal intTotalStart As Integer, ByVal intTotalEnd As Integer, _
                                           ByVal intEmptyIndex As Integer, ByVal intVariableIndex As Integer, _
                                           ByVal intFixedIndex As Integer, ByVal colMax As Integer, _
                                           ByVal bMTPCheck As Boolean) As Boolean

        '//Setup Total Line
        xSt.Range(xSt.Cells(intTotalIndex, intTotalStart), xSt.Cells(intTotalIndex, intTotalEnd)).ClearContents()
        xSt.Range(xSt.Cells(intTotalIndex, intTotalStart), xSt.Cells(intTotalIndex, intTotalEnd)).Font.Bold = True
        xSt.Range(xSt.Cells(intTotalIndex, intTotalStart), xSt.Cells(intTotalIndex, intTotalEnd)).MergeCells = True
        xSt.Range(xSt.Cells(intTotalIndex, intTotalStart), xSt.Cells(intTotalIndex, intTotalEnd)).Value = strTotalTitle
        xSt.Range(xSt.Cells(intTotalIndex, intTotalEnd), xSt.Cells(intTotalIndex, colMax)).Font.Bold = True

        If strTotalAlign = "Center" Then
            xSt.Range(xSt.Cells(intTotalIndex, intTotalStart), xSt.Cells(intTotalIndex, intTotalEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        Else
            xSt.Range(xSt.Cells(intTotalIndex, intTotalStart), xSt.Cells(intTotalIndex, intTotalEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        End If

        '//Merge empty line
        'If bMTPCheck = False Then
        '    xSt.Range(xSt.Cells(intEmptyIndex, intEmptyStart), xSt.Cells(intEmptyIndex, colMax)).ClearContents()
        '    xSt.Range(xSt.Cells(intEmptyIndex, intEmptyStart), xSt.Cells(intEmptyIndex, colMax)).MergeCells = True
        'End If
        xSt.Range(xSt.Cells(intEmptyIndex, intEmptyStart), xSt.Cells(intEmptyIndex, colMax)).ClearContents()
        xSt.Range(xSt.Cells(intEmptyIndex, intEmptyStart), xSt.Cells(intEmptyIndex, colMax)).MergeCells = True

        '//Setup Total VariableCost Line
        xSt.Range(xSt.Cells(intVariableIndex, intTotalStart), xSt.Cells(intVariableIndex, intTotalEnd)).ClearContents()
        xSt.Range(xSt.Cells(intVariableIndex, intTotalStart), xSt.Cells(intVariableIndex, intTotalEnd)).Font.Bold = True
        xSt.Range(xSt.Cells(intVariableIndex, intTotalStart), xSt.Cells(intVariableIndex, intTotalEnd)).MergeCells = True
        xSt.Range(xSt.Cells(intVariableIndex, intTotalStart), xSt.Cells(intVariableIndex, intTotalEnd)).Value = "Variable Cost"
        xSt.Range(xSt.Cells(intVariableIndex, intTotalStart), xSt.Cells(intVariableIndex, intTotalEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

        '//Setup Total FixedCost Line
        xSt.Range(xSt.Cells(intFixedIndex, intTotalStart), xSt.Cells(intFixedIndex, intTotalEnd)).ClearContents()
        xSt.Range(xSt.Cells(intFixedIndex, intTotalStart), xSt.Cells(intFixedIndex, intTotalEnd)).Font.Bold = True
        xSt.Range(xSt.Cells(intFixedIndex, intTotalStart), xSt.Cells(intFixedIndex, intTotalEnd)).MergeCells = True
        xSt.Range(xSt.Cells(intFixedIndex, intTotalStart), xSt.Cells(intFixedIndex, intTotalEnd)).Value = "Fixed Cost"
        xSt.Range(xSt.Cells(intFixedIndex, intTotalStart), xSt.Cells(intFixedIndex, intTotalEnd)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

    End Function

    Public Shared Function SetupMTPEmptyColumn(ByRef xSt As Excel.Worksheet, _
                                               ByVal colStartIndex As Integer, ByVal rowMax As Integer, ByVal colMax As Integer, _
                                               ByVal intColIndex As Integer, ByVal intEmptyLineIndex As Integer, _
                                               ByVal intEmptyLineStart As Integer, Optional ByVal bLastLineEmpty As Boolean = False) As Boolean

        Dim xColumn As Excel.Range = CType(xSt.Columns(intColIndex, Type.Missing), Excel.Range)
        xColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Type.Missing)
        xSt.Range(xSt.Cells(colStartIndex, intColIndex), xSt.Cells(rowMax, intColIndex)).Borders.LineStyle = 0
        xSt.Range(xSt.Cells(colStartIndex, intColIndex), xSt.Cells(rowMax, intColIndex)).ColumnWidth = 2
        xSt.Range(xSt.Cells(colStartIndex, intColIndex), xSt.Cells(rowMax, intColIndex)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
        xSt.Range(xSt.Cells(colStartIndex, intColIndex), xSt.Cells(rowMax, intColIndex)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1

        xSt.Range(xSt.Cells(intEmptyLineIndex, intEmptyLineStart), xSt.Cells(intEmptyLineIndex, intColIndex - 1)).ClearContents()
        xSt.Range(xSt.Cells(intEmptyLineIndex, intEmptyLineStart), xSt.Cells(intEmptyLineIndex, intColIndex - 1)).MergeCells = True
        xSt.Range(xSt.Cells(intEmptyLineIndex, intColIndex + 1), xSt.Cells(intEmptyLineIndex, colMax + 1)).ClearContents()
        xSt.Range(xSt.Cells(intEmptyLineIndex, intColIndex + 1), xSt.Cells(intEmptyLineIndex, colMax + 1)).MergeCells = True

        If bLastLineEmpty = True Then
            xSt.Range(xSt.Cells(rowMax, intEmptyLineStart), xSt.Cells(rowMax, intColIndex - 1)).ClearContents()
            xSt.Range(xSt.Cells(rowMax, intEmptyLineStart), xSt.Cells(rowMax, intColIndex - 1)).MergeCells = True
            xSt.Range(xSt.Cells(rowMax, intColIndex + 1), xSt.Cells(rowMax, colMax + 1)).ClearContents()
            xSt.Range(xSt.Cells(rowMax, intColIndex + 1), xSt.Cells(rowMax, colMax + 1)).MergeCells = True
        End If

        Return True

    End Function

    Public Shared Function CreateTableTemplate() As DataTable

        Dim dtTable As DataTable = New DataTable()

        Dim col As DataColumn

        col = New DataColumn()
        col.ColumnName = "Column_Name"
        col.DataType = Type.GetType("System.String")
        dtTable.Columns.Add(col)

        col = New DataColumn()
        col.ColumnName = "Column_Title"
        col.DataType = Type.GetType("System.String")
        dtTable.Columns.Add(col)

        Return dtTable

    End Function

    Public Shared Function GetGroupExpensesTitle(ByVal strExpenses As String) As String
        Dim strRet As String = String.Empty
        Select Case strExpenses
            Case CStr(enumExpenseType.LaborExpense)
                strRet = P_EXPENSE_TYPE_LABOR '"Labor Expenses"
            Case CStr(enumExpenseType.VariableExpense)
                strRet = P_EXPENSE_TYPE_VARIABLE  '"Variable Expense"
            Case CStr(enumExpenseType.FixedExpense)
                strRet = P_EXPENSE_TYPE_FIXED  '"Fixed Expense"
        End Select
        Return strRet
    End Function

    Public Shared Function GetGroupCostTitle(ByVal strCost As String) As String
        Dim strRet As String = String.Empty
        Select Case strCost
            Case CStr(enumCost.FC)
                strRet = P_FC_COST '"Manufacturing Cost"
            Case CStr(enumCost.ADMIN)
                strRet = P_ADMIN_COST '"Administration Cost"
        End Select
        Return strRet
    End Function

    Public Shared Function IsAccountNoEmpty(ByVal row As DataRow) As Boolean
        If row("ACCOUNT_NO").ToString = String.Empty Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Function IsGroupHeaderEmpty(ByVal row As DataRow) As Boolean
        If row("Group_Header").ToString = String.Empty Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Function ExcelReleasememory(ByRef excelApp As Excel.Application, _
                                               ByRef workBook As Excel.Workbook, _
                                               ByRef workSheet As Excel.Worksheet) As Boolean

        '// Release memory
        System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheet)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(workBook)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
        GC.SuppressFinalize(workSheet)
        GC.SuppressFinalize(workBook)
        GC.SuppressFinalize(excelApp)
        GC.Collect()
        workSheet = Nothing
        workBook = Nothing
        excelApp = Nothing

        Return True
    End Function

    Public Shared Function CalPercent(ByVal Num1 As Decimal, ByVal Num2 As Decimal) As Decimal
        Dim decResult As Decimal = 0D

        If Num2 <> 0 Then
            decResult = (Num1 - Num2) / Num2 * 100
        End If

        Return decResult

    End Function

    Public Shared Function GetBudgetCompareDiffPeriod(ByVal pMonth As String) As String
        Dim strDiffPeriod As String = String.Empty

        If pMonth = "1" OrElse _
            pMonth = "2" OrElse _
            pMonth = "3" OrElse _
            pMonth = "4" OrElse _
            pMonth = "5" OrElse _
            pMonth = "6" Then

            strDiffPeriod = "Diff vs OB"

        Else

            strDiffPeriod = "Diff vs OB"

        End If

        Return strDiffPeriod

    End Function

#End Region

End Class
