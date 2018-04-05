Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class frmBG0010

#Region "Variable"
    Private myClsBG0010BL As New clsBG0010BL()
#End Region

#Region "Function"
    Public Sub ShowSideMenu()
        '// ShowSide Menu up to User's Permission
        '// Clear Menu
        trvMenu.Nodes("nodMenu").Nodes.Clear()

        '// [Home Page] Menu
        ShowHomeMenu()

        '// [Information] Menu
        ShowInfoMenu()

        '// [Budget Jouenal] Menu
        ShowBudgetMenu()

        '// [Budget Reports] Menu
        ShowBudgeReportsMenu()

        '// [Budget Compare Reports] Menu
        ShowBudgeCompareReportsMenu()

        '// [Account Tools] Menu
        If myClsBG0010BL.HavePermission(enumPermissionCd.Adjust) = True Or _
        myClsBG0010BL.HavePermission(enumPermissionCd.Import) = True Or _
        myClsBG0010BL.HavePermission(enumPermissionCd.Export) = True Or _
        myClsBG0010BL.HavePermission(enumPermissionCd.System) = True Then
            ShowAccountMenu()
        End If

        '// [Master Data] Menu
        If myClsBG0010BL.HavePermission(enumPermissionCd.Master) = True Then
            ShowMasterMenu()
        End If

        '// [Misc] Menu
        ShowMiscMenu()

        '//[Direct Input]
        If myClsBG0010BL.HavePermission(enumPermissionCd.DirectInput) = True Then
            ShowDirectInputMenu()
        End If

        '// Expand Menu
        trvMenu.Nodes("nodMenu").Expand()
        trvMenu.Nodes("nodMenu").EnsureVisible()
    End Sub

    Private Sub ShowHomeMenu()
        '// Create Node if do not exist.
        If trvMenu.Nodes("nodMenu").Nodes("nodHome") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes.Add("nodHome", "Home Page", "home.ico", "home.ico")
        End If
    End Sub

    Public Sub ShowInfoMenu()
        Dim strIcon As String = String.Empty

        '// Create Node if do not exist.
        If trvMenu.Nodes("nodMenu").Nodes("nodInfo") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes.Add("nodInfo", "Information", "info.ico", "info.ico")
        Else
            trvMenu.Nodes("nodMenu").Nodes("nodInfo").Nodes.Clear()
        End If

        '// Show Information File Node
        If myClsBG0010BL.SearchInformation() = True Then

            For Each dr As DataRow In myClsBG0010BL.InfoList.Rows
                '// Set Node Icon
                If CStr(dr("FILE_PATH")).ToLower.EndsWith("xls") Or _
                CStr(dr("FILE_PATH")).ToLower.EndsWith("xlsx") Then
                    strIcon = "excel.ico"

                ElseIf CStr(dr("FILE_PATH")).ToLower.EndsWith("doc") Or _
                CStr(dr("FILE_PATH")).ToLower.EndsWith("docx") Then
                    strIcon = "word.ico"

                ElseIf CStr(dr("FILE_PATH")).ToLower.EndsWith("jpg") Or _
                CStr(dr("FILE_PATH")).ToLower.EndsWith("png") Or _
                CStr(dr("FILE_PATH")).ToLower.EndsWith("gif") Or _
                CStr(dr("FILE_PATH")).ToLower.EndsWith("bmp") Then
                    strIcon = "image.ico"

                ElseIf CStr(dr("FILE_PATH")).ToLower.EndsWith("ppt") Then
                    strIcon = "powerpoint.ico"

                ElseIf CStr(dr("FILE_PATH")).ToLower.EndsWith("pdf") Then
                    strIcon = "pdf.ico"

                Else
                    strIcon = "none.ico"

                End If

                '// Add Info Node
                trvMenu.Nodes("nodMenu").Nodes("nodInfo").Nodes.Add("nodF" & CStr(dr("FILE_NO")), CStr(dr("FILE_TITLE")), strIcon, strIcon)
            Next

            '// Expand Menu
            trvMenu.Nodes("nodMenu").Nodes("nodInfo").ExpandAll()
        End If
    End Sub

    Public Sub ShowBudgetMenu()
        '// Create Node if do not exist.
        If trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes.Add("nodBudgetJournal", "Budget Journal", "folder.ico", "folder.ico")
        Else
            trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes.Clear()
        End If

        myClsBG0010BL.ViewAll = myClsBG0010BL.HavePermission(enumPermissionCd.View)

        '// [View Budget Journal] Menu
        ShowViewBudgetMenu()

        If p_blnReadOnlyMode = False Then
            '// [Input Budget Journal] Menu
            If myClsBG0010BL.HavePermission(enumPermissionCd.Entry) = True Then
                ShowInputBudgetMenu()
            End If

            '// [Approve Budget Journal] Menu
            If myClsBG0010BL.HavePermission(enumPermissionCd.Approve) = True Then
                ShowApproveBudgetMenu()
            End If

            '// [Adjust Budget Journal] Menu
            If myClsBG0010BL.HavePermission(enumPermissionCd.Adjust) = True Then
                ShowAdjustBudgetMenu()
            End If

            '// [Authorize Budget Journal] Menu
            If myClsBG0010BL.HavePermission(enumPermissionCd.Auth1) = True Then
                ShowAuth1BudgetMenu()
            End If
            If myClsBG0010BL.HavePermission(enumPermissionCd.Auth2) = True Then
                ShowAuth2BudgetMenu()
            End If
        End If

        '// Expand Menu
        trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").ExpandAll()
        If trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodViewBudget") IsNot Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodViewBudget").Collapse()
        End If
    End Sub

    Private Sub ShowViewBudgetMenu()
        Dim strYear As String = String.Empty
        Dim strPeriodId As String = String.Empty
        Dim strPeriodName As String = String.Empty
        Dim strProjectNo As String = String.Empty
        Dim strProjectName As String = String.Empty

        '// Create Node if do not exist.
        If trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodViewBudget") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes.Add("nodViewBudget", "View Budget", _
                                                                         "folder.ico", "folder.ico")
        End If

        myClsBG0010BL.BudgetYear = CStr(Year(Now))

        '// Load Budget Period
        If myClsBG0010BL.SearchViewPeriod() = True AndAlso myClsBG0010BL.PeriodList.Rows.Count > 0 Then
            For Each drPd As DataRow In myClsBG0010BL.PeriodList.Rows

                '// Create Budget Key
                strYear = CStr(drPd("BUDGET_YEAR"))
                strPeriodId = CStr(drPd("PERIOD_TYPE"))
                strProjectNo = CStr(drPd("PROJECT_NO"))

                If strPeriodId = CStr(enumPeriodType.OriginalBudget) Then
                    strPeriodName = "Original Budget"
                ElseIf strPeriodId = CStr(enumPeriodType.EstimateBudget) Then
                    strPeriodName = "Estimate Budget"
                ElseIf strPeriodId = CStr(enumPeriodType.ReviseBudget) Then
                    strPeriodName = "Revise Budget"
                ElseIf strPeriodId = CStr(enumPeriodType.MTPBudget) Then
                    strPeriodName = "MTP Budget"
                End If

                '// Set common parameters
                myClsBG0010BL.BudgetYear = strYear
                myClsBG0010BL.PeriodType = strPeriodId
                myClsBG0010BL.UserPIC = p_strUserPIC
                myClsBG0010BL.ProjectNo = strProjectNo

                If strProjectNo.Equals("1") Then
                    strProjectName = String.Empty
                Else
                    strProjectName = " (Project " & strProjectNo & ")"
                End If


                '// 1) Add View Budget Node (Expense)
                myClsBG0010BL.BudgetType = P_BUDGET_TYPE_EXPENSE

                ''myClsBG0010BL.GetMaxRevNo()
                myClsBG0010BL.RevNo = "1"

                If myClsBG0010BL.CheckBudgetDataExistView() = True Then
                    trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodViewBudget").Nodes.Add( _
                        strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_EXPENSE & strProjectNo, strYear & " " & strPeriodName & strProjectName, "journal.ico", "journal.ico")
                End If

                '// 2) Add View Budget Node (Asset)
                myClsBG0010BL.BudgetType = P_BUDGET_TYPE_ASSET

                ''myClsBG0010BL.GetMaxRevNo()
                myClsBG0010BL.RevNo = "1"

                If strPeriodId <> CStr(enumPeriodType.MTPBudget) AndAlso myClsBG0010BL.CheckBudgetDataExistView() = True Then
                    trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodViewBudget").Nodes.Add( _
                        strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_ASSET & strProjectNo, strYear & " " & strPeriodName & " (Investment)" & strProjectName, "journal.ico", "journal.ico")
                End If

            Next
        End If
    End Sub

    Private Sub ShowInputBudgetMenu()
        Dim strYear As String = String.Empty
        Dim strPeriodId As String = String.Empty
        Dim strPeriodName As String = String.Empty
        Dim strProjectNo As String = String.Empty
        Dim strProjectName As String = String.Empty

        '// Create Node if do not exist.
        If trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodInputBudget") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes.Add("nodInputBudget", "Input Budget", _
                                                                         "folder.ico", "folder.ico")
        End If

        '// Load Budget Period
        If myClsBG0010BL.SearchOpenPeriod() = True AndAlso myClsBG0010BL.PeriodList.Rows.Count > 0 Then
            For Each drPd As DataRow In myClsBG0010BL.PeriodList.Rows

                '// Create Budget Key
                strYear = CStr(drPd("BUDGET_YEAR"))
                strPeriodId = CStr(drPd("PERIOD_TYPE"))
                strProjectNo = CStr(drPd("PROJECT_NO"))

                If strPeriodId = CStr(enumPeriodType.OriginalBudget) Then
                    strPeriodName = "Original Budget"
                ElseIf strPeriodId = CStr(enumPeriodType.EstimateBudget) Then
                    strPeriodName = "Estimate Budget"
                ElseIf strPeriodId = CStr(enumPeriodType.ReviseBudget) Then
                    strPeriodName = "Revise Budget"
                ElseIf strPeriodId = CStr(enumPeriodType.MTPBudget) Then
                    strPeriodName = "MTP Budget"
                End If

                '// Set common parameters
                myClsBG0010BL.BudgetYear = strYear
                myClsBG0010BL.PeriodType = strPeriodId
                myClsBG0010BL.Status = CStr(enumBudgetStatus.NewRecord)
                myClsBG0010BL.RevNo = "1"
                myClsBG0010BL.ProjectNo = strProjectNo

                If strProjectNo.Equals("1") Then
                    strProjectName = String.Empty
                Else
                    strProjectName = " (Project " & strProjectNo & ")"
                End If

                '// 1) Add Input Budget Node (Expense)
                myClsBG0010BL.UserPIC = p_strUserPIC
                myClsBG0010BL.BudgetType = P_BUDGET_TYPE_EXPENSE

                If myClsBG0010BL.CheckBudgetOrderMatch() = True Then
                    For Each dr As DataRow In myClsBG0010BL.UserPicList.Rows
                        myClsBG0010BL.UserPIC = CStr(dr("PERSON_IN_CHARGE_NO"))

                        If myClsBG0010BL.CheckBudgetDataExist() = False Then
                            trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodInputBudget").Nodes.Add( _
                                strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_EXPENSE & strProjectNo, strYear & " " & strPeriodName & strProjectName, _
                                    "journal.ico", "journal.ico")
                            Exit For

                        ElseIf myClsBG0010BL.CheckBudgetDataStatus() = True Then
                            trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodInputBudget").Nodes.Add( _
                                strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_EXPENSE & strProjectNo, strYear & " " & strPeriodName & strProjectName, _
                                    "journal.ico", "journal.ico")
                            Exit For

                        ElseIf myClsBG0010BL.CheckBudgetDataStatusReInputByOrder() = True Then
                            trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodInputBudget").Nodes.Add( _
                                strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_EXPENSE & strProjectNo, strYear & " " & strPeriodName & strProjectName, _
                                    "journal.ico", "journal.ico")
                            Exit For

                        End If

                    Next
                End If

                '// 2) Add Input Budget Node (Asset)
                myClsBG0010BL.UserPIC = p_strUserPIC
                myClsBG0010BL.BudgetType = P_BUDGET_TYPE_ASSET

                If strPeriodId <> CStr(enumPeriodType.MTPBudget) AndAlso myClsBG0010BL.CheckBudgetOrderMatch() = True Then
                    For Each dr As DataRow In myClsBG0010BL.UserPicList.Rows
                        myClsBG0010BL.UserPIC = CStr(dr("PERSON_IN_CHARGE_NO"))

                        If myClsBG0010BL.CheckBudgetDataExist() = False Then
                            trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodInputBudget").Nodes.Add( _
                                strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_ASSET & strProjectNo, strYear & " " & strPeriodName & " (Investment)" & strProjectName, _
                                    "journal.ico", "journal.ico")
                            Exit For

                        ElseIf myClsBG0010BL.CheckBudgetDataStatus() = True Then
                            trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodInputBudget").Nodes.Add( _
                                strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_ASSET & strProjectNo, strYear & " " & strPeriodName & " (Investment)" & strProjectName, _
                                    "journal.ico", "journal.ico")
                            Exit For

                        ElseIf myClsBG0010BL.CheckBudgetDataStatusReInputByOrder() = True Then
                            trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodInputBudget").Nodes.Add( _
                                strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_EXPENSE & strProjectNo, strYear & " " & strPeriodName & strProjectName, _
                                    "journal.ico", "journal.ico")
                            Exit For

                        End If
                    Next
                End If

            Next
        End If
    End Sub

    Private Sub ShowApproveBudgetMenu()
        Dim strYear As String = String.Empty
        Dim strPeriodId As String = String.Empty
        Dim strPeriodName As String = String.Empty
        Dim strProjectNo As String = String.Empty
        Dim strProjectName As String = String.Empty

        '// Create Node if do not exist.
        If trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodApproveBudget") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes.Add("nodApproveBudget", "Approve Budget", _
                                                                         "folder.ico", "folder.ico")
        End If

        '// Load Budget Period
        If myClsBG0010BL.SearchOpenPeriod() = True AndAlso myClsBG0010BL.PeriodList.Rows.Count > 0 Then
            For Each drPd As DataRow In myClsBG0010BL.PeriodList.Rows

                '// Create Budget Key
                strYear = CStr(drPd("BUDGET_YEAR"))
                strPeriodId = CStr(drPd("PERIOD_TYPE"))
                strProjectNo = CStr(drPd("PROJECT_NO"))
                If strPeriodId = CStr(enumPeriodType.OriginalBudget) Then
                    strPeriodName = "Original Budget"
                ElseIf strPeriodId = CStr(enumPeriodType.EstimateBudget) Then
                    strPeriodName = "Estimate Budget"
                ElseIf strPeriodId = CStr(enumPeriodType.ReviseBudget) Then
                    strPeriodName = "Revise Budget"
                ElseIf strPeriodId = CStr(enumPeriodType.MTPBudget) Then
                    strPeriodName = "MTP Budget"
                End If

                '// Set common parameters
                myClsBG0010BL.BudgetYear = strYear
                myClsBG0010BL.PeriodType = strPeriodId
                myClsBG0010BL.UserPIC = p_strUserPIC
                myClsBG0010BL.Status = CStr(enumBudgetStatus.Submit)
                myClsBG0010BL.RevNo = "1"
                myClsBG0010BL.ProjectNo = strProjectNo

                If strProjectNo.Equals("1") Then
                    strProjectName = String.Empty
                Else
                    strProjectName = " (Project " & strProjectNo & ")"
                End If

                '// 1) Add Approve Budget Node (Expense)
                myClsBG0010BL.BudgetType = P_BUDGET_TYPE_EXPENSE

                If myClsBG0010BL.CheckBudgetDataStatus(True) = True Then
                    trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodApproveBudget").Nodes.Add( _
                        strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_EXPENSE & strProjectNo, strYear & " " & strPeriodName & strProjectName, _
                            "journal.ico", "journal.ico")

                Else

                    '// Add Approve From ReINPUTBYORDERNO
                    If myClsBG0010BL.CheckBudgetDataStatusReInputByOrder(True) = True Then
                        trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodApproveBudget").Nodes.Add( _
                            strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_EXPENSE & strProjectNo, strYear & " " & strPeriodName & strProjectName, _
                                "journal.ico", "journal.ico")
                    End If

                End If


                '// 2) Add Approve Budget Node (Asset)
                myClsBG0010BL.BudgetType = P_BUDGET_TYPE_ASSET

                If strPeriodId <> CStr(enumPeriodType.MTPBudget) AndAlso myClsBG0010BL.CheckBudgetDataStatus(True) = True Then
                    trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodApproveBudget").Nodes.Add( _
                        strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_ASSET & strProjectNo, strYear & " " & strPeriodName & " (Investment)" & strProjectName, _
                            "journal.ico", "journal.ico")

                Else

                    '// Add Approve From ReINPUTBYORDERNO
                    If myClsBG0010BL.CheckBudgetDataStatusReInputByOrder(True) = True Then
                        trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodApproveBudget").Nodes.Add( _
                            strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_EXPENSE & strProjectNo, strYear & " " & strPeriodName & strProjectName, _
                                "journal.ico", "journal.ico")
                    End If

                End If

            Next
        End If
    End Sub

    Private Sub ShowAdjustBudgetMenu()
        Dim strYear As String = String.Empty
        Dim strPeriodId As String = String.Empty
        Dim strPeriodName As String = String.Empty
        Dim strProjectNo As String = String.Empty
        Dim strProjectName As String = String.Empty

        '// Create Node if do not exist.
        If trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodAdjustBudget") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes.Add("nodAdjustBudget", "Adjust Budget Journal", _
                                                                         "folder.ico", "folder.ico")
        End If

        '// Load Budget Period
        If myClsBG0010BL.SearchAllPeriod() = True AndAlso myClsBG0010BL.PeriodList.Rows.Count > 0 Then
            For Each drPd As DataRow In myClsBG0010BL.PeriodList.Rows

                '// Create Budget Key
                strYear = CStr(drPd("BUDGET_YEAR"))
                strPeriodId = CStr(drPd("PERIOD_TYPE"))
                strProjectNo = CStr(drPd("PROJECT_NO"))

                If strPeriodId = CStr(enumPeriodType.OriginalBudget) Then
                    strPeriodName = "Original Budget"
                ElseIf strPeriodId = CStr(enumPeriodType.EstimateBudget) Then
                    strPeriodName = "Estimate Budget"
                ElseIf strPeriodId = CStr(enumPeriodType.ReviseBudget) Then
                    strPeriodName = "Revise Budget"
                ElseIf strPeriodId = CStr(enumPeriodType.MTPBudget) Then
                    strPeriodName = "MTP Budget"
                End If

                '// Set common parameters
                myClsBG0010BL.BudgetYear = strYear
                myClsBG0010BL.PeriodType = strPeriodId
                myClsBG0010BL.UserPIC = p_strUserPIC
                myClsBG0010BL.ProjectNo = strProjectNo

                If strProjectNo.Equals("1") Then
                    strProjectName = String.Empty
                Else
                    strProjectName = " (Project " & strProjectNo & ")"
                End If

                '// 1) Add Adjust Budget Node (Expense)
                myClsBG0010BL.BudgetType = P_BUDGET_TYPE_EXPENSE
                myClsBG0010BL.GetMaxRevNo()

                myClsBG0010BL.Status = CStr(enumBudgetStatus.Approve)

                If myClsBG0010BL.CheckBudgetDataStatus() = True Then
                    trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodAdjustBudget").Nodes.Add( _
                        strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_EXPENSE & strProjectNo, strYear & " " & strPeriodName & strProjectName, _
                            "journal.ico", "journal.ico")
                Else
                    myClsBG0010BL.Status = CStr(enumBudgetStatus.Authorize2)

                    If myClsBG0010BL.CheckBudgetDataStatus() = True Then
                        trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodAdjustBudget").Nodes.Add( _
                            strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_EXPENSE & strProjectNo, strYear & " " & strPeriodName & strProjectName, _
                                "journal.ico", "journal.ico")
                    End If
                End If

                '// 2) Add Adjust Budget Node (Asset)
                myClsBG0010BL.BudgetType = P_BUDGET_TYPE_ASSET
                myClsBG0010BL.GetMaxRevNo()

                myClsBG0010BL.Status = CStr(enumBudgetStatus.Approve)

                If strPeriodId <> CStr(enumPeriodType.MTPBudget) Then

                    If myClsBG0010BL.CheckBudgetDataStatus() = True Then
                        trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodAdjustBudget").Nodes.Add( _
                            strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_ASSET & strProjectNo, strYear & " " & strPeriodName & " (Investment)" & strProjectName, _
                                "journal.ico", "journal.ico")
                    Else
                        myClsBG0010BL.Status = CStr(enumBudgetStatus.Authorize2)

                        If myClsBG0010BL.CheckBudgetDataStatus() = True Then
                            trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodAdjustBudget").Nodes.Add( _
                                strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_ASSET & strProjectNo, strYear & " " & strPeriodName & " (Investment)" & strProjectName, _
                                    "journal.ico", "journal.ico")
                        End If
                    End If

                End If

            Next
        End If
    End Sub

    Private Sub ShowAdjustBudgetForDirectMenu()
        Dim strYear As String = String.Empty
        Dim strPeriodId As String = String.Empty
        Dim strPeriodName As String = String.Empty
        Dim strProjectNo As String = String.Empty
        Dim strProjectName As String = String.Empty

        '// Create Node if do not exist.
        If trvMenu.Nodes("nodMenu").Nodes("nodDirectInput").Nodes("nodAdjustBudget2") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodDirectInput").Nodes.Add("nodAdjustBudget2", "Adjust Budget Journal", _
                                                                         "folder.ico", "folder.ico")
        End If

        '// Load Budget Period
        If myClsBG0010BL.SearchAllPeriod() = True AndAlso myClsBG0010BL.PeriodList.Rows.Count > 0 Then
            For Each drPd As DataRow In myClsBG0010BL.PeriodList.Rows

                '// Create Budget Key
                strYear = CStr(drPd("BUDGET_YEAR"))
                strPeriodId = CStr(drPd("PERIOD_TYPE"))
                strProjectNo = CStr(drPd("PROJECT_NO"))

                If strPeriodId = CStr(enumPeriodType.OriginalBudget) Then
                    strPeriodName = "Original Budget"
                ElseIf strPeriodId = CStr(enumPeriodType.EstimateBudget) Then
                    strPeriodName = "Estimate Budget"
                ElseIf strPeriodId = CStr(enumPeriodType.ReviseBudget) Then
                    strPeriodName = "Revise Budget"
                ElseIf strPeriodId = CStr(enumPeriodType.MTPBudget) Then
                    strPeriodName = "MTP Budget"
                End If

                '// Set common parameters
                myClsBG0010BL.BudgetYear = strYear
                myClsBG0010BL.PeriodType = strPeriodId
                myClsBG0010BL.UserPIC = p_strUserPIC
                myClsBG0010BL.ProjectNo = strProjectNo

                If strProjectNo.Equals("1") Then
                    strProjectName = String.Empty
                Else
                    strProjectName = " (Project " & strProjectNo & ")"
                End If

                '// 1) Add Adjust Budget Node (Expense)
                myClsBG0010BL.BudgetType = P_BUDGET_TYPE_EXPENSE
                myClsBG0010BL.GetMaxRevNo()

                myClsBG0010BL.Status = CStr(enumBudgetStatus.Approve)

                If myClsBG0010BL.CheckBudgetDataStatus() = True Then
                    trvMenu.Nodes("nodMenu").Nodes("nodDirectInput").Nodes("nodAdjustBudget2").Nodes.Add( _
                        strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_EXPENSE & strProjectNo, strYear & " " & strPeriodName & strProjectName, _
                            "journal.ico", "journal.ico")
                Else
                    myClsBG0010BL.Status = CStr(enumBudgetStatus.Authorize2)

                    If myClsBG0010BL.CheckBudgetDataStatus() = True Then
                        trvMenu.Nodes("nodMenu").Nodes("nodDirectInput").Nodes("nodAdjustBudget2").Nodes.Add( _
                            strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_EXPENSE & strProjectNo, strYear & " " & strPeriodName & strProjectName, _
                                "journal.ico", "journal.ico")
                    End If
                End If

                '// 2) Add Adjust Budget Node (Asset)
                myClsBG0010BL.BudgetType = P_BUDGET_TYPE_ASSET
                myClsBG0010BL.GetMaxRevNo()

                myClsBG0010BL.Status = CStr(enumBudgetStatus.Approve)

                If strPeriodId <> CStr(enumPeriodType.MTPBudget) Then

                    If myClsBG0010BL.CheckBudgetDataStatus() = True Then
                        trvMenu.Nodes("nodMenu").Nodes("nodDirectInput").Nodes("nodAdjustBudget2").Nodes.Add( _
                            strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_ASSET & strProjectNo, strYear & " " & strPeriodName & " (Investment)" & strProjectName, _
                                "journal.ico", "journal.ico")
                    Else
                        myClsBG0010BL.Status = CStr(enumBudgetStatus.Authorize2)

                        If myClsBG0010BL.CheckBudgetDataStatus() = True Then
                            trvMenu.Nodes("nodMenu").Nodes("nodDirectInput").Nodes("nodAdjustBudget2").Nodes.Add( _
                                strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_ASSET & strProjectNo, strYear & " " & strPeriodName & " (Investment)" & strProjectName, _
                                    "journal.ico", "journal.ico")
                        End If
                    End If

                End If

            Next
        End If
    End Sub
    Private Sub ShowAuth1BudgetMenu()
        Dim strYear As String = String.Empty
        Dim strPeriodId As String = String.Empty
        Dim strPeriodName As String = String.Empty
        Dim strProjectNo As String = String.Empty
        Dim strProjectName As String = String.Empty

        '// Create Node if do not exist.
        If trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodAuth1Budget") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes.Add("nodAuth1Budget", "Authorize Budget (AD)", _
                                                                         "folder.ico", "folder.ico")
        End If

        '// Load Budget Period
        If myClsBG0010BL.SearchAllPeriod() = True AndAlso myClsBG0010BL.PeriodList.Rows.Count > 0 Then
            For Each drPd As DataRow In myClsBG0010BL.PeriodList.Rows

                '// Create Budget Key
                strYear = CStr(drPd("BUDGET_YEAR"))
                strPeriodId = CStr(drPd("PERIOD_TYPE"))
                strProjectNo = CStr(drPd("PROJECT_NO"))
                If strPeriodId = CStr(enumPeriodType.OriginalBudget) Then
                    strPeriodName = "Original Budget"
                ElseIf strPeriodId = CStr(enumPeriodType.EstimateBudget) Then
                    strPeriodName = "Estimate Budget"
                ElseIf strPeriodId = CStr(enumPeriodType.ReviseBudget) Then
                    strPeriodName = "Revise Budget"
                ElseIf strPeriodId = CStr(enumPeriodType.MTPBudget) Then
                    strPeriodName = "MTP Budget"
                End If

                '// [AUTH1] Node
                '// Set common parameters
                myClsBG0010BL.BudgetYear = strYear
                myClsBG0010BL.PeriodType = strPeriodId
                myClsBG0010BL.UserPIC = p_strUserPIC
                myClsBG0010BL.Status = CStr(enumBudgetStatus.Adjust)
                myClsBG0010BL.ProjectNo = strProjectNo

                If strProjectNo.Equals("1") Then
                    strProjectName = String.Empty
                Else
                    strProjectName = " (Project " & strProjectNo & ")"
                End If

                '// 1) Add Adjust Budget Node (Expense)
                myClsBG0010BL.BudgetType = P_BUDGET_TYPE_EXPENSE
                myClsBG0010BL.GetMaxRevNo()

                If myClsBG0010BL.CheckBudgetDataStatus() = True Then
                    trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodAuth1Budget").Nodes.Add( _
                        strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_EXPENSE & strProjectNo, strYear & " " & strPeriodName & strProjectName, _
                            "journal.ico", "journal.ico")
                End If

                '// 2) Add Adjust Budget Node (Asset)
                myClsBG0010BL.BudgetType = P_BUDGET_TYPE_ASSET
                myClsBG0010BL.GetMaxRevNo()

                If strPeriodId <> CStr(enumPeriodType.MTPBudget) AndAlso myClsBG0010BL.CheckBudgetDataStatus() = True Then
                    trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodAuth1Budget").Nodes.Add( _
                        strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_ASSET & strProjectNo, strYear & " " & strPeriodName & " (Investment)" & strProjectName, _
                            "journal.ico", "journal.ico")
                End If

            Next
        End If
    End Sub

    Private Sub ShowAuth2BudgetMenu()
        Dim strYear As String = String.Empty
        Dim strPeriodId As String = String.Empty
        Dim strPeriodName As String = String.Empty
        Dim strProjectNo As String = String.Empty
        Dim strProjectName As String = String.Empty

        '// Create Node if do not exist.
        If trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodAuth2Budget") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes.Add("nodAuth2Budget", "Authorize Budget (MD)", _
                                                                         "folder.ico", "folder.ico")
        End If

        '// Load Budget Period
        If myClsBG0010BL.SearchAllPeriod() = True AndAlso myClsBG0010BL.PeriodList.Rows.Count > 0 Then
            For Each drPd As DataRow In myClsBG0010BL.PeriodList.Rows

                '// Create Budget Key
                strYear = CStr(drPd("BUDGET_YEAR"))
                strPeriodId = CStr(drPd("PERIOD_TYPE"))
                strProjectNo = CStr(drPd("PROJECT_NO"))
                If strPeriodId = CStr(enumPeriodType.OriginalBudget) Then
                    strPeriodName = "Original Budget"
                ElseIf strPeriodId = CStr(enumPeriodType.EstimateBudget) Then
                    strPeriodName = "Estimate Budget"
                ElseIf strPeriodId = CStr(enumPeriodType.ReviseBudget) Then
                    strPeriodName = "Revise Budget"
                ElseIf strPeriodId = CStr(enumPeriodType.MTPBudget) Then
                    strPeriodName = "MTP Budget"
                End If

                '// [AUTH1] Node
                '// Set common parameters
                myClsBG0010BL.BudgetYear = strYear
                myClsBG0010BL.PeriodType = strPeriodId
                myClsBG0010BL.UserPIC = p_strUserPIC
                myClsBG0010BL.Status = CStr(enumBudgetStatus.Authorize1)
                myClsBG0010BL.ProjectNo = strProjectNo

                If strProjectNo.Equals("1") Then
                    strProjectName = String.Empty
                Else
                    strProjectName = " (Project " & strProjectNo & ")"
                End If

                '// 1) Add Adjust Budget Node (Expense)
                myClsBG0010BL.BudgetType = P_BUDGET_TYPE_EXPENSE
                myClsBG0010BL.GetMaxRevNo()

                If myClsBG0010BL.CheckBudgetDataStatus() = True Then
                    trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodAuth2Budget").Nodes.Add( _
                        strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_EXPENSE & strProjectNo, strYear & " " & strPeriodName & strProjectName, _
                            "journal.ico", "journal.ico")
                End If

                '// 2) Add Adjust Budget Node (Asset)
                myClsBG0010BL.BudgetType = P_BUDGET_TYPE_ASSET
                myClsBG0010BL.GetMaxRevNo()

                If strPeriodId <> CStr(enumPeriodType.MTPBudget) AndAlso myClsBG0010BL.CheckBudgetDataStatus() = True Then
                    trvMenu.Nodes("nodMenu").Nodes("nodBudgetJournal").Nodes("nodAuth2Budget").Nodes.Add( _
                        strYear & CInt(strPeriodId).ToString("00") & P_BUDGET_TYPE_ASSET & strProjectNo, strYear & " " & strPeriodName & " (Investment)" & strProjectName, _
                            "journal.ico", "journal.ico")
                End If

            Next
        End If
    End Sub

    Public Sub ShowAccountMenu()
        '// Create Node if do not exist.
        If trvMenu.Nodes("nodMenu").Nodes("nodAccountTools") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes.Add("nodAccountTools", "Account Tools", "folder.ico", "opened_folder.ico")
        End If

        '//[Open New Period] Menu
        If myClsBG0010BL.HavePermission(enumPermissionCd.Adjust) = True Then
            If trvMenu.Nodes("nodMenu").Nodes("nodAccountTools").Nodes("nodOpenPeriod") Is Nothing Then
                trvMenu.Nodes("nodMenu").Nodes("nodAccountTools").Nodes.Add("nodOpenPeriod", "Open New Budget Period", _
                                                                            "window.ico", "window.ico")
            End If
        End If

        '//[Close Period] Menu
        If myClsBG0010BL.HavePermission(enumPermissionCd.Adjust) = True Then
            If trvMenu.Nodes("nodMenu").Nodes("nodAccountTools").Nodes("nodClosePeriod") Is Nothing Then
                trvMenu.Nodes("nodMenu").Nodes("nodAccountTools").Nodes.Add("nodClosePeriod", "Close Budget Period", _
                                                                            "window.ico", "window.ico")
            End If
        End If

        '//[View Period] Menu
        If myClsBG0010BL.HavePermission(enumPermissionCd.Adjust) = True Then
            If trvMenu.Nodes("nodMenu").Nodes("nodAccountTools").Nodes("nodViewPeriod") Is Nothing Then
                trvMenu.Nodes("nodMenu").Nodes("nodAccountTools").Nodes.Add("nodViewPeriod", "View Budget Period", _
                                                                            "window.ico", "window.ico")
            End If
        End If

        '//[Reopen Period] Menu
        If myClsBG0010BL.HavePermission(enumPermissionCd.Adjust) = True Then
            If trvMenu.Nodes("nodMenu").Nodes("nodAccountTools").Nodes("nodReopenPeriod") Is Nothing Then
                trvMenu.Nodes("nodMenu").Nodes("nodAccountTools").Nodes.Add("nodReopenPeriod", "Re-Open Budget Period", _
                                                                            "window.ico", "window.ico")
            End If
        End If

        '//[Import Budget Data] Menu
        If myClsBG0010BL.HavePermission(enumPermissionCd.Import) = True Then
            If trvMenu.Nodes("nodMenu").Nodes("nodAccountTools").Nodes("nodImportBudget") Is Nothing Then
                trvMenu.Nodes("nodMenu").Nodes("nodAccountTools").Nodes.Add("nodImportBudget", "Import Data from SAP", _
                                                                            "window.ico", "window.ico")
            End If
        End If

        '//[Export File to SAP] Menu
        If myClsBG0010BL.HavePermission(enumPermissionCd.Export) = True Then
            If trvMenu.Nodes("nodMenu").Nodes("nodAccountTools").Nodes("nodExportFile") Is Nothing Then
                trvMenu.Nodes("nodMenu").Nodes("nodAccountTools").Nodes.Add("nodExportFile", "Export File to SAP", _
                                                                            "window.ico", "window.ico")
            End If
        End If

        '//[Add/Remove Information] Menu
        If myClsBG0010BL.HavePermission(enumPermissionCd.System) = True Then
            If trvMenu.Nodes("nodMenu").Nodes("nodAccountTools").Nodes("nodAddRemoveInfo") Is Nothing Then
                trvMenu.Nodes("nodMenu").Nodes("nodAccountTools").Nodes.Add("nodAddRemoveInfo", "Add / Remove Information", _
                                                                            "window.ico", "window.ico")
            End If
        End If

        '//[View Transaction Log] Menu
        If myClsBG0010BL.HavePermission(enumPermissionCd.System) = True Then
            If trvMenu.Nodes("nodMenu").Nodes("nodAccountTools").Nodes("nodViewTransLog") Is Nothing Then
                trvMenu.Nodes("nodMenu").Nodes("nodAccountTools").Nodes.Add("nodViewTransLog", "View Transaction Log", _
                                                                            "window.ico", "window.ico")
            End If
        End If

        '// [Unlock Person in charge] Menu
        If p_intUserLevelId = enumUserLevel.SystemAdministrator Then
            If trvMenu.Nodes("nodMenu").Nodes("nodAccountTools").Nodes("nodUnlockPic") Is Nothing Then
                trvMenu.Nodes("nodMenu").Nodes("nodAccountTools").Nodes.Add("nodUnlockPic", "Unlock Person In Charge", _
                                                                            "window.ico", "window.ico")
            End If
        End If

        '// [Database Backup] Menu
        If p_intUserLevelId = enumUserLevel.SystemAdministrator Then
            If trvMenu.Nodes("nodMenu").Nodes("nodAccountTools").Nodes("nodDBBackup") Is Nothing Then
                trvMenu.Nodes("nodMenu").Nodes("nodAccountTools").Nodes.Add("nodDBBackup", "Backup / Restore Data", _
                                                                            "window.ico", "window.ico")
            End If
        End If
    End Sub

    Private Sub ShowBudgeReportsMenu()
        '// Create Node if do not exist.
        If trvMenu.Nodes("nodMenu").Nodes("nodReports") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes.Add("nodReports", "Budget Reports", "folder.ico", "folder.ico")
        End If

        If p_intUserLevelId = enumUserLevel.SystemAdministrator Then
            '// [Report Option] Menu
            If myClsBG0010BL.HavePermission(enumPermissionCd.Master) = True Then
                If trvMenu.Nodes("nodMenu").Nodes("nodReports").Nodes("nodReportOption") Is Nothing Then
                    trvMenu.Nodes("nodMenu").Nodes("nodReports").Nodes.Add("nodReportOption", "Report Options", _
                                                                              "window.ico", "window.ico")
                End If
            End If
        End If

        '// [Detail by Person In Charge] Menu (RPT001)
        If trvMenu.Nodes("nodMenu").Nodes("nodReports").Nodes("nodRPT001") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodReports").Nodes.Add("nodRPT001", "Detail by Person In Charge", _
                                                                      "document.ico", "document.ico")
        End If

        If p_intUserLevelId < enumUserLevel.GeneralManager Then
            '// [Summary by Person In Charge] Menu (RPT002)
            If trvMenu.Nodes("nodMenu").Nodes("nodReports").Nodes("nodRPT002") Is Nothing Then
                trvMenu.Nodes("nodMenu").Nodes("nodReports").Nodes.Add("nodRPT002", "Summary by Person In Charge", _
                                                                          "document.ico", "document.ico")
            End If

            '// [Detail by Account No] Menu (RPT003)
            If trvMenu.Nodes("nodMenu").Nodes("nodReports").Nodes("nodRPT003") Is Nothing Then
                trvMenu.Nodes("nodMenu").Nodes("nodReports").Nodes.Add("nodRPT003", "Detail by Account No", _
                                                                          "document.ico", "document.ico")
            End If

            '// [Summary by Account No] Menu (RPT004)
            If trvMenu.Nodes("nodMenu").Nodes("nodReports").Nodes("nodRPT004") Is Nothing Then
                trvMenu.Nodes("nodMenu").Nodes("nodReports").Nodes.Add("nodRPT004", "Summary by Account No", _
                                                                          "document.ico", "document.ico")
            End If

            '// [Summary by Applicant] Menu (RPT005)
            If trvMenu.Nodes("nodMenu").Nodes("nodReports").Nodes("nodRPT005") Is Nothing Then
                trvMenu.Nodes("nodMenu").Nodes("nodReports").Nodes.Add("nodRPT005", "Summary by Applicant", _
                                                                          "document.ico", "document.ico")
            End If

         
        End If

        If p_intUserLevelId < enumUserLevel.NormalUser Then
            '// [Summary by Investment] Menu (RPT006)
            If trvMenu.Nodes("nodMenu").Nodes("nodReports").Nodes("nodRPT006") Is Nothing Then
                trvMenu.Nodes("nodMenu").Nodes("nodReports").Nodes.Add("nodRPT006", "Summary by Investment", _
                                                                          "document.ico", "document.ico")
            End If
        End If


        '// [Comment by Person In Charge Report] Menu (RPT008)
        If trvMenu.Nodes("nodMenu").Nodes("nodReports").Nodes("nodRPT008") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodReports").Nodes.Add("nodRPT008", "Comment by Person In Charge Report", _
                                                                      "document.ico", "document.ico")
        End If



    End Sub

    Private Sub ShowBudgeCompareReportsMenu()
        '// Create Node if do not exist.
        If trvMenu.Nodes("nodMenu").Nodes("nodCompareReports") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes.Add("nodCompareReports", "Budget Compare Reports", "folder.ico", "folder.ico")
        End If

        '// [Detail by Person In Charge] Menu (RPT007)
        If trvMenu.Nodes("nodMenu").Nodes("nodCompareReports").Nodes("nodRPT007") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodCompareReports").Nodes.Add("nodRPT007", "Detail by Person In Charge", _
                                                                      "document.ico", "document.ico")
        End If

        If p_intUserLevelId < enumUserLevel.GeneralManager Then
            '// [Summary by Person In Charge] Menu (RPT007_1)
            If trvMenu.Nodes("nodMenu").Nodes("nodCompareReports").Nodes("nodRPT007_1") Is Nothing Then
                trvMenu.Nodes("nodMenu").Nodes("nodCompareReports").Nodes.Add("nodRPT007_1", "Summary by Person In Charge", _
                                                                          "document.ico", "document.ico")
            End If

            '// [Detail by Account No] Menu (RPT007_2)
            If trvMenu.Nodes("nodMenu").Nodes("nodCompareReports").Nodes("nodRPT007_2") Is Nothing Then
                trvMenu.Nodes("nodMenu").Nodes("nodCompareReports").Nodes.Add("nodRPT007_2", "Detail by Account No", _
                                                                          "document.ico", "document.ico")
            End If

            '// [Summary by Account No] Menu (RPT007_3)
            If trvMenu.Nodes("nodMenu").Nodes("nodCompareReports").Nodes("nodRPT007_3") Is Nothing Then
                trvMenu.Nodes("nodMenu").Nodes("nodCompareReports").Nodes.Add("nodRPT007_3", "Summary by Account No", _
                                                                          "document.ico", "document.ico")
            End If

        End If

        If p_intUserLevelId < enumUserLevel.NormalUser Then
            '// [Summary by Investment] Menu (RPT007_4)
            If trvMenu.Nodes("nodMenu").Nodes("nodCompareReports").Nodes("nodRPT007_4") Is Nothing Then
                trvMenu.Nodes("nodMenu").Nodes("nodCompareReports").Nodes.Add("nodRPT007_4", "Summary by Investment", _
                                                                          "document.ico", "document.ico")
            End If
        End If

    End Sub

    Private Sub ShowMasterMenu()
        '// Create Node if do not exist.
        If trvMenu.Nodes("nodMenu").Nodes("nodMasterData") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes.Add("nodMasterData", "Master Data", "folder.ico", "folder.ico")
        End If

        '// [Budget Order Master] Menu
        If trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes("nodOrderMaster") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes.Add("nodOrderMaster", "Budget Order Master", _
                                                                      "window.ico", "window.ico")
        End If

        '// [Account Master] Menu
        If trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes("nodAccountMaster") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes.Add("nodAccountMaster", "Account Master", _
                                                                      "window.ico", "window.ico")
        End If

        '// [Department Master] Menu
        If trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes("nodDeptMaster") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes.Add("nodDeptMaster", "Department Master", _
                                                                      "window.ico", "window.ico")
        End If

        '// [Person In Charge Master] Menu
        If trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes("nodPICMaster") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes.Add("nodPICMaster", "Person In Charge Master", _
                                                                      "window.ico", "window.ico")
        End If

        '// [Transfer Cost Master] Menu
        If trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes("nodTransferMaster") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes.Add("nodTransferMaster", "Transfer Cost Master", _
                                                                      "window.ico", "window.ico")
        End If

        '// [Budget Adjust Master] Menu
        If trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes("nodAdjustMaster") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes.Add("nodAdjustMaster", "Budget Adjust Master", _
                                                                      "window.ico", "window.ico")
        End If

        '// [Asset Group Master] Menu
        If trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes("nodAssetGroupMaster") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes.Add("nodAssetGroupMaster", "Asset Group Master", _
                                                                      "window.ico", "window.ico")
        End If

        '// [Asset Category Master] Menu
        If trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes("nodAssetCategoryMaster") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes.Add("nodAssetCategoryMaster", "Asset Category Master", _
                                                                      "window.ico", "window.ico")
        End If

        '// [Asset Project Master] Menu
        If trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes("nodAssetProjectMaster") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes.Add("nodAssetProjectMaster", "Asset Project Master", _
                                                                      "window.ico", "window.ico")
        End If

        '// [Child PIC Master] Menu
        If trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes("nodChildPICMaster") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes.Add("nodChildPICMaster", "Child PIC Master", _
                                                                      "window.ico", "window.ico")
        End If

        '// [User Master] Menu
        If trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes("nodUserMaster") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes.Add("nodUserMaster", "User Master", _
                                                                      "window.ico", "window.ico")
        End If

        '// [User Level Master] Menu
        If trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes("nodUserLevelMaster") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodMasterData").Nodes.Add("nodUserLevelMaster", "User Level Master", _
                                                                      "window.ico", "window.ico")
        End If
    End Sub

    Private Sub ShowMiscMenu()
        '// Create Node if do not exist.
        If trvMenu.Nodes("nodMenu").Nodes("nodMisc") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes.Add("nodMisc", "Misc", "folder.ico", "folder.ico")
        End If

        '// [Change Password] Menu
        If trvMenu.Nodes("nodMenu").Nodes("nodMisc").Nodes("nodUserPassword") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes("nodMisc").Nodes.Add("nodUserPassword", "Change Password", _
                                                                "window.ico", "window.ico")
        End If
    End Sub

    Private Sub ShowDirectInputMenu()
        '// Create Node if do not exist.
        If trvMenu.Nodes("nodMenu").Nodes("nodDirectInput") Is Nothing Then
            trvMenu.Nodes("nodMenu").Nodes.Add("nodDirectInput", "Direct Input", "folder.ico", "folder.ico")
        Else
            trvMenu.Nodes("nodMenu").Nodes("nodDirectInput").Nodes.Clear()
        End If

        myClsBG0010BL.ViewAll = myClsBG0010BL.HavePermission(enumPermissionCd.View)

        ''// [View Budget Journal] Menu
        'ShowViewBudgetMenu()

        If p_blnReadOnlyMode = False Then
            ''// [Input Budget Journal] Menu
            'If myClsBG0010BL.HavePermission(enumPermissionCd.Entry) = True Then
            '    ShowInputBudgetMenu()
            'End If

            ''// [Approve Budget Journal] Menu
            'If myClsBG0010BL.HavePermission(enumPermissionCd.Approve) = True Then
            '    ShowApproveBudgetMenu()
            'End If

            '// [Adjust Budget Journal] Menu
            If myClsBG0010BL.HavePermission(enumPermissionCd.DirectInput) = True Then
                ShowAdjustBudgetForDirectMenu()
            End If

            ''// [Authorize Budget Journal] Menu
            'If myClsBG0010BL.HavePermission(enumPermissionCd.Auth1) = True Then
            '    ShowAuth1BudgetMenu()
            'End If
            'If myClsBG0010BL.HavePermission(enumPermissionCd.Auth2) = True Then
            '    ShowAuth2BudgetMenu()
            'End If
        End If

        ''// Expand Menu
        'trvMenu.Nodes("nodMenu").Nodes("nodDirectInput").ExpandAll()
        'If trvMenu.Nodes("nodMenu").Nodes("nodDirectInput").Nodes("nodViewBudget2") IsNot Nothing Then
        '    trvMenu.Nodes("nodMenu").Nodes("nodDirectInput").Nodes("nodViewBudget2").Collapse()
        'End If
    End Sub

    Public Sub EnableMenuStrip()
        '// Hide Tools Menu if not administrator level.
        '// --Load User Permission
        If myClsBG0010BL.GetUserInfo() = True Then
            If CStr(myClsBG0010BL.UserInfo.Rows(0).Item("SYSTEM")) = "Y" Then
                ToolsMenu.Enabled = True
            End If
        Else
            MessageBox.Show("Error: Can not get user information!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub OpenMenuFromTree(ByVal nodSelNode As TreeNode)
        If nodSelNode.Parent IsNot Nothing Then

            If nodSelNode.Name = "nodHome" Then

                '// Open Home Screen
                If p_frmBG0110 Is Nothing OrElse p_frmBG0110.IsDisposed Then
                    p_frmBG0110 = New frmBG0110(Me, "Home Page", True)
                    p_frmBG0110.Show()
                Else
                    If p_frmBG0110.WindowState = FormWindowState.Minimized Then
                        p_frmBG0110.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0110.BringToFront()
                End If

            ElseIf nodSelNode.Parent.Name = "nodInfo" Then

                '// Open Information File
                If myClsBG0010BL.InfoList IsNot Nothing Then
                    Try
                        Dim dr As DataRow = myClsBG0010BL.InfoList.Select("FILE_NO = " & nodSelNode.Name.Substring(4))(0)
                        Process.Start(CStr(dr("FILE_PATH")))

                    Catch ex As Exception
                        MessageBox.Show("Error: Can not open file!" & vbNewLine & ex.Message, Me.Text, _
                                        MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If

            ElseIf nodSelNode.Parent.Name = "nodViewBudget" Then

                '// Open View Budget Screen
                Dim myFrmBG0200 As frmBG0200 = p_frmBG0200.FirstOrDefault(Function(f) f.BudgetKey = nodSelNode.Name And _
                                                                            f.OperationCd = enumOperationCd.ViewBudget)
                If myFrmBG0200 Is Nothing Then
                    '// Create new when do not exist.
                    myFrmBG0200 = New frmBG0200(Me, True, "View " & nodSelNode.Text, nodSelNode.Name, enumOperationCd.ViewBudget)
                    myFrmBG0200.Show()

                    '// Add to screen list
                    p_frmBG0200.Add(myFrmBG0200)
                Else
                    '// Bring to front when exist.
                    If myFrmBG0200.WindowState = FormWindowState.Minimized Then
                        myFrmBG0200.WindowState = FormWindowState.Normal
                    End If
                    myFrmBG0200.BringToFront()
                End If

            ElseIf nodSelNode.Parent.Name = "nodInputBudget" Then

                '// Open Input Budget Screen
                Dim myFrmBG0200 As frmBG0200 = p_frmBG0200.FirstOrDefault(Function(f) f.BudgetKey = nodSelNode.Name And _
                                                                            f.OperationCd = enumOperationCd.InputBudget)
                If myFrmBG0200 Is Nothing Then
                    '// Create new when do not exist.
                    myFrmBG0200 = New frmBG0200(Me, True, "Input " & nodSelNode.Text, nodSelNode.Name, enumOperationCd.InputBudget)
                    myFrmBG0200.Show()

                    '// Add to screen list
                    p_frmBG0200.Add(myFrmBG0200)
                Else
                    '// Bring to front when exist.
                    If myFrmBG0200.WindowState = FormWindowState.Minimized Then
                        myFrmBG0200.WindowState = FormWindowState.Normal
                    End If
                    myFrmBG0200.BringToFront()
                End If

            ElseIf nodSelNode.Parent.Name = "nodApproveBudget" Then

                '// Open Approve Budget Screen
                Dim myFrmBG0200 As frmBG0200 = p_frmBG0200.FirstOrDefault(Function(f) f.BudgetKey = nodSelNode.Name And _
                                                                            f.OperationCd = enumOperationCd.ApproveBudget)
                If myFrmBG0200 Is Nothing Then
                    '// Create new when do not exist.
                    myFrmBG0200 = New frmBG0200(Me, True, "Approve " & nodSelNode.Text, nodSelNode.Name, enumOperationCd.ApproveBudget)
                    myFrmBG0200.Show()

                    '// Add to screen list
                    p_frmBG0200.Add(myFrmBG0200)
                Else
                    '// Bring to front when exist.
                    If myFrmBG0200.WindowState = FormWindowState.Minimized Then
                        myFrmBG0200.WindowState = FormWindowState.Normal
                    End If
                    myFrmBG0200.BringToFront()
                End If

            ElseIf nodSelNode.Parent.Name = "nodAdjustBudget" Then

                '// Open Approve Budget Screen
                Dim myFrmBG0200 As frmBG0200 = p_frmBG0200.FirstOrDefault(Function(f) f.BudgetKey = nodSelNode.Name And _
                                                                            f.OperationCd = enumOperationCd.AdjustBudget)
                If myFrmBG0200 Is Nothing Then
                    '// Create new when do not exist.
                    myFrmBG0200 = New frmBG0200(Me, True, "Adjust " & nodSelNode.Text, nodSelNode.Name, enumOperationCd.AdjustBudget)
                    myFrmBG0200.Show()

                    '// Add to screen list
                    p_frmBG0200.Add(myFrmBG0200)
                Else
                    '// Bring to front when exist.
                    If myFrmBG0200.WindowState = FormWindowState.Minimized Then
                        myFrmBG0200.WindowState = FormWindowState.Normal
                    End If
                    myFrmBG0200.BringToFront()
                End If

            ElseIf nodSelNode.Parent.Name = "nodAuth1Budget" Then

                '// Open Authorize1 Budget Screen
                Dim myFrmBG0200 As frmBG0200 = p_frmBG0200.FirstOrDefault(Function(f) f.BudgetKey = nodSelNode.Name And _
                                                                                        f.OperationCd = enumOperationCd.Authorize1)
                If myFrmBG0200 Is Nothing Then
                    '// Create new when do not exist.
                    myFrmBG0200 = New frmBG0200(Me, True, "Authorize " & nodSelNode.Text & " (AD)", nodSelNode.Name, _
                                                enumOperationCd.Authorize1)
                    myFrmBG0200.Show()

                    '// Add to screen list
                    p_frmBG0200.Add(myFrmBG0200)
                Else
                    '// Bring to front when exist.
                    If myFrmBG0200.WindowState = FormWindowState.Minimized Then
                        myFrmBG0200.WindowState = FormWindowState.Normal
                    End If
                    myFrmBG0200.BringToFront()
                End If

            ElseIf nodSelNode.Parent.Name = "nodAuth2Budget" Then

                '// Open Authorize2 Budget Screen
                Dim myFrmBG0200 As frmBG0200 = p_frmBG0200.FirstOrDefault(Function(f) f.BudgetKey = nodSelNode.Name And _
                                                                                        f.OperationCd = enumOperationCd.Authorize2)
                If myFrmBG0200 Is Nothing Then
                    '// Create new when do not exist.
                    myFrmBG0200 = New frmBG0200(Me, True, "Authorize " & nodSelNode.Text & " (MD)", nodSelNode.Name, _
                                                enumOperationCd.Authorize2)
                    myFrmBG0200.Show()

                    '// Add to screen list
                    p_frmBG0200.Add(myFrmBG0200)
                Else
                    '// Bring to front when exist.
                    If myFrmBG0200.WindowState = FormWindowState.Minimized Then
                        myFrmBG0200.WindowState = FormWindowState.Normal
                    End If
                    myFrmBG0200.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodImportBudget" Then

                '// Open Import Data From SAP Screen
                If p_frmBG0350 Is Nothing OrElse p_frmBG0350.IsDisposed Then
                    p_frmBG0350 = New frmBG0350(Me, "Import Data from SAP", True)
                    p_frmBG0350.Show()
                Else
                    If p_frmBG0350.WindowState = FormWindowState.Minimized Then
                        p_frmBG0350.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0350.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodExportFile" Then

                '// Open Export File To SAP Screen
                If p_frmBG0360 Is Nothing OrElse p_frmBG0360.IsDisposed Then
                    p_frmBG0360 = New frmBG0360(Me, "Export File To SAP", True)
                    p_frmBG0360.Show()
                Else
                    If p_frmBG0360.WindowState = FormWindowState.Minimized Then
                        p_frmBG0360.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0360.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodAddRemoveInfo" Then

                '// Open Add/Remove Information Screen
                If p_frmBG0380 Is Nothing OrElse p_frmBG0380.IsDisposed Then
                    p_frmBG0380 = New frmBG0380(Me, "Add/Remove Information", True)
                    p_frmBG0380.Show()
                Else
                    If p_frmBG0380.WindowState = FormWindowState.Minimized Then
                        p_frmBG0380.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0380.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodViewTransLog" Then

                '// View Transaction Log Screen
                If p_frmBG0390 Is Nothing OrElse p_frmBG0390.IsDisposed Then
                    p_frmBG0390 = New frmBG0390(Me, "View Transaction Log", True)
                    p_frmBG0390.Show()
                Else
                    If p_frmBG0390.WindowState = FormWindowState.Minimized Then
                        p_frmBG0390.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0390.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodReportOption" Then

                '// Open Report Options Screen
                If p_frmBG0401 Is Nothing OrElse p_frmBG0401.IsDisposed Then
                    p_frmBG0401 = New frmBG0401(Me, "Report Options", True)
                    p_frmBG0401.Show()
                Else
                    If p_frmBG0401.WindowState = FormWindowState.Minimized Then
                        p_frmBG0401.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0401.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodRPT001" Then

                '// Open Detail by PIC Report Screen
                If p_frmBG0410 Is Nothing OrElse p_frmBG0410.IsDisposed Then
                    p_frmBG0410 = New frmBG0410(Me, "Detail by Person In Charge Report", True)
                    p_frmBG0410.Show()
                Else
                    If p_frmBG0410.WindowState = FormWindowState.Minimized Then
                        p_frmBG0410.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0410.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodRPT002" Then

                '// Open Summary by PIC Report Screen
                If p_frmBG0420 Is Nothing OrElse p_frmBG0420.IsDisposed Then
                    p_frmBG0420 = New frmBG0420(Me, "Summary by Person In Charge Report", True)
                    p_frmBG0420.Show()
                Else
                    If p_frmBG0420.WindowState = FormWindowState.Minimized Then
                        p_frmBG0420.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0420.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodRPT003" Then

                '// Open Detail by Account No Report Screen
                If p_frmBG0430 Is Nothing OrElse p_frmBG0430.IsDisposed Then
                    p_frmBG0430 = New frmBG0430(Me, "Detail by Account No Report", True)
                    p_frmBG0430.Show()
                Else
                    If p_frmBG0430.WindowState = FormWindowState.Minimized Then
                        p_frmBG0430.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0430.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodRPT004" Then

                '// Open Summary by Account No Report Screen
                If p_frmBG0440 Is Nothing OrElse p_frmBG0440.IsDisposed Then
                    p_frmBG0440 = New frmBG0440(Me, "Summary by Account No Report", True)
                    p_frmBG0440.Show()
                Else
                    If p_frmBG0440.WindowState = FormWindowState.Minimized Then
                        p_frmBG0440.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0440.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodRPT005" Then

                '// Open Summary by Applicant Report Screen
                If p_frmBG0450 Is Nothing OrElse p_frmBG0450.IsDisposed Then
                    p_frmBG0450 = New frmBG0450(Me, "Summary by Applicant Report", True)
                    p_frmBG0450.Show()
                Else
                    If p_frmBG0450.WindowState = FormWindowState.Minimized Then
                        p_frmBG0450.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0450.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodRPT006" Then

                '// Open Summary by Investment Report Screen
                If p_frmBG0460 Is Nothing OrElse p_frmBG0460.IsDisposed Then
                    p_frmBG0460 = New frmBG0460(Me, "Summary by Investment Report", True)
                    p_frmBG0460.Show()
                Else
                    If p_frmBG0460.WindowState = FormWindowState.Minimized Then
                        p_frmBG0460.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0460.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodRPT007" Then

                '// Open Budget Compare Report Screen(Detail by PIC)
                If p_frmBG0470 Is Nothing OrElse p_frmBG0470.IsDisposed Then
                    p_frmBG0470 = New frmBG0470(Me, "Detail by Person In Charge Report (Budget Compare)", True)
                    p_frmBG0470.Show()
                Else
                    If p_frmBG0470.WindowState = FormWindowState.Minimized Then
                        p_frmBG0470.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0470.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodRPT007_1" Then

                '// Open Budget Compare Report Screen(Summary by PIC)
                If p_frmBG0471 Is Nothing OrElse p_frmBG0471.IsDisposed Then
                    p_frmBG0471 = New frmBG0471(Me, "Summary by Person In Charge Report (Budget Compare)", True)
                    p_frmBG0471.Show()
                Else
                    If p_frmBG0471.WindowState = FormWindowState.Minimized Then
                        p_frmBG0471.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0471.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodRPT007_2" Then

                '// Open Budget Compare Report Screen(Detail by AccountNo)
                If p_frmBG0472 Is Nothing OrElse p_frmBG0472.IsDisposed Then
                    p_frmBG0472 = New frmBG0472(Me, "Detail by Account No Report (Budget Compare)", True)
                    p_frmBG0472.Show()
                Else
                    If p_frmBG0472.WindowState = FormWindowState.Minimized Then
                        p_frmBG0472.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0472.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodRPT007_3" Then

                '// Open Budget Compare Report Screen(Summary by AccountNo)
                If p_frmBG0473 Is Nothing OrElse p_frmBG0473.IsDisposed Then
                    p_frmBG0473 = New frmBG0473(Me, "Summary by Account No Report (Budget Compare)", True)
                    p_frmBG0473.Show()
                Else
                    If p_frmBG0473.WindowState = FormWindowState.Minimized Then
                        p_frmBG0473.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0473.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodRPT007_4" Then

                '// Open Budget Compare Report Screen(Summary by Investment)
                If p_frmBG0474 Is Nothing OrElse p_frmBG0474.IsDisposed Then
                    p_frmBG0474 = New frmBG0474(Me, "Summary by Investment Report (Budget Compare)", True)
                    p_frmBG0474.Show()
                Else
                    If p_frmBG0474.WindowState = FormWindowState.Minimized Then
                        p_frmBG0474.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0474.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodRPT008" Then

                '// Open Comment by Person In Charge Report
                If p_frmBG0480 Is Nothing OrElse p_frmBG0480.IsDisposed Then
                    p_frmBG0480 = New frmBG0480(Me, "Comment by Person In Charge Report", True)
                    p_frmBG0480.Show()
                Else
                    If p_frmBG0480.WindowState = FormWindowState.Minimized Then
                        p_frmBG0480.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0480.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodOpenPeriod" Then

                '// Open New Period Screen
                If p_frmBG0310 Is Nothing OrElse p_frmBG0310.IsDisposed Then
                    p_frmBG0310 = New frmBG0310(Me, "Open New Budget Period", True)
                    p_frmBG0310.Show()
                Else
                    If p_frmBG0310.WindowState = FormWindowState.Minimized Then
                        p_frmBG0310.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0310.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodClosePeriod" Then

                '// Open Close Period Screen
                If p_frmBG0320 Is Nothing OrElse p_frmBG0320.IsDisposed Then
                    p_frmBG0320 = New frmBG0320(Me, "Close Budget Period", True)
                    p_frmBG0320.Show()
                Else
                    If p_frmBG0320.WindowState = FormWindowState.Minimized Then
                        p_frmBG0320.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0320.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodViewPeriod" Then

                '// Open View Period Screen
                If p_frmBG0395 Is Nothing OrElse p_frmBG0395.IsDisposed Then
                    p_frmBG0395 = New frmBG0395(Me, "View Budget Period", True)
                    p_frmBG0395.Show()
                Else
                    If p_frmBG0395.WindowState = FormWindowState.Minimized Then
                        p_frmBG0395.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0395.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodReopenPeriod" Then

                '// Open Reopen Period Screen
                If p_frmBG0330 Is Nothing OrElse p_frmBG0330.IsDisposed Then
                    p_frmBG0330 = New frmBG0330(Me, "Re-Open Budget Period", True)
                    p_frmBG0330.Show()
                Else
                    If p_frmBG0330.WindowState = FormWindowState.Minimized Then
                        p_frmBG0330.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0330.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodOrderMaster" Then

                '// Open Order Master Screen
                If p_frmBG0620 Is Nothing OrElse p_frmBG0620.IsDisposed Then
                    p_frmBG0620 = New frmBG0620(Me, "Budget Order Master", True)
                    p_frmBG0620.Show()
                Else
                    If p_frmBG0620.WindowState = FormWindowState.Minimized Then
                        p_frmBG0620.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0620.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodAccountMaster" Then

                '// Open Account Master Screen
                If p_frmBG0630 Is Nothing OrElse p_frmBG0630.IsDisposed Then
                    p_frmBG0630 = New frmBG0630(Me, "Account Master", True)
                    p_frmBG0630.Show()
                Else
                    If p_frmBG0630.WindowState = FormWindowState.Minimized Then
                        p_frmBG0630.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0630.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodDeptMaster" Then

                '// Open Dept Master Screen
                If p_frmBG0640 Is Nothing OrElse p_frmBG0640.IsDisposed Then
                    p_frmBG0640 = New frmBG0640(Me, "Department Master", True)
                    p_frmBG0640.Show()
                Else
                    If p_frmBG0640.WindowState = FormWindowState.Minimized Then
                        p_frmBG0640.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0640.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodPICMaster" Then

                '// Open PIC Master Screen
                If p_frmBG0650 Is Nothing OrElse p_frmBG0650.IsDisposed Then
                    p_frmBG0650 = New frmBG0650(Me, "Person In Charge Master", True)
                    p_frmBG0650.Show()
                Else
                    If p_frmBG0650.WindowState = FormWindowState.Minimized Then
                        p_frmBG0650.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0650.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodTransferMaster" Then

                '// Open Transfer Cost Master Screen
                If p_frmBG0660 Is Nothing OrElse p_frmBG0660.IsDisposed Then
                    p_frmBG0660 = New frmBG0660(Me, "Transfer Cost Master", True)
                    p_frmBG0660.Show()
                Else
                    If p_frmBG0660.WindowState = FormWindowState.Minimized Then
                        p_frmBG0660.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0660.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodAdjustMaster" Then

                '// Open Asset Group Master Screen
                If p_frmBG0670 Is Nothing OrElse p_frmBG0670.IsDisposed Then
                    p_frmBG0670 = New frmBG0670(Me, "Budget Adjust Master", True)
                    p_frmBG0670.Show()
                Else
                    If p_frmBG0670.WindowState = FormWindowState.Minimized Then
                        p_frmBG0670.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0670.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodAssetGroupMaster" Then

                '// Open Asset Group Master Screen
                If p_frmBG0680 Is Nothing OrElse p_frmBG0680.IsDisposed Then
                    p_frmBG0680 = New frmBG0680(Me, "Asset Group Master", True)
                    p_frmBG0680.Show()
                Else
                    If p_frmBG0680.WindowState = FormWindowState.Minimized Then
                        p_frmBG0680.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0680.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodAssetCategoryMaster" Then

                '// Open Asset Category Master Screen
                If p_frmBG0681 Is Nothing OrElse p_frmBG0681.IsDisposed Then
                    p_frmBG0681 = New frmBG0681(Me, "Asset Category Master", True)
                    p_frmBG0681.Show()
                Else
                    If p_frmBG0681.WindowState = FormWindowState.Minimized Then
                        p_frmBG0681.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0681.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodAssetProjectMaster" Then

                '// Open Asset Project Master Screen
                If p_frmBG0682 Is Nothing OrElse p_frmBG0682.IsDisposed Then
                    p_frmBG0682 = New frmBG0682(Me, "Asset Project Master", True)
                    p_frmBG0682.Show()
                Else
                    If p_frmBG0682.WindowState = FormWindowState.Minimized Then
                        p_frmBG0682.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0682.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodChildPICMaster" Then

                '// Open Child PIC Master Screen
                If p_frmBG0690 Is Nothing OrElse p_frmBG0690.IsDisposed Then
                    p_frmBG0690 = New frmBG0690(Me, "Child Person In Charge Master", True)
                    p_frmBG0690.Show()
                Else
                    If p_frmBG0690.WindowState = FormWindowState.Minimized Then
                        p_frmBG0690.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0690.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodUserMaster" Then

                '// Open User Master Screen
                If p_frmBG0610 Is Nothing OrElse p_frmBG0610.IsDisposed Then
                    p_frmBG0610 = New frmBG0610(Me, "User Master", True)
                    p_frmBG0610.Show()
                Else
                    If p_frmBG0610.WindowState = FormWindowState.Minimized Then
                        p_frmBG0610.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0610.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodUserLevelMaster" Then

                '// Open User Level Master Screen
                If p_frmBG0611 Is Nothing OrElse p_frmBG0611.IsDisposed Then
                    p_frmBG0611 = New frmBG0611(Me, "User Level Master", True)
                    p_frmBG0611.Show()
                Else
                    If p_frmBG0611.WindowState = FormWindowState.Minimized Then
                        p_frmBG0611.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0611.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodUserPassword" Then

                '// Open User Password Screen
                If p_frmBG0710 Is Nothing OrElse p_frmBG0710.IsDisposed Then
                    p_frmBG0710 = New frmBG0710(Me, "User Password", True)
                    p_frmBG0710.Show()
                Else
                    If p_frmBG0710.WindowState = FormWindowState.Minimized Then
                        p_frmBG0710.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0710.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodUnlockPic" Then

                '// Open Unlock Pic Screen
                If p_frmBG0720 Is Nothing OrElse p_frmBG0720.IsDisposed Then
                    p_frmBG0720 = New frmBG0720(Me, "Unlock Person In Charge", True)
                    p_frmBG0720.Show()
                Else
                    If p_frmBG0720.WindowState = FormWindowState.Minimized Then
                        p_frmBG0720.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0720.BringToFront()
                End If

            ElseIf nodSelNode.Name = "nodDBBackup" Then

                '// Open Unlock Pic Screen
                If p_frmBG0730 Is Nothing OrElse p_frmBG0730.IsDisposed Then
                    p_frmBG0730 = New frmBG0730(Me, "Backup / Restore Data", True)
                    p_frmBG0730.Show()
                Else
                    If p_frmBG0730.WindowState = FormWindowState.Minimized Then
                        p_frmBG0730.WindowState = FormWindowState.Normal
                    End If
                    p_frmBG0730.BringToFront()
                End If

            ElseIf nodSelNode.Parent.Name = "nodAdjustBudget2" Then

                '// Open Approve Budget Screen
                Dim myFrmBG0200 As frmBG0200 = p_frmBG0200.FirstOrDefault(Function(f) f.BudgetKey = nodSelNode.Name And _
                                                                            f.OperationCd = enumOperationCd.AdjustBudgetDirectInput)
                If myFrmBG0200 Is Nothing Then
                    '// Create new when do not exist.
                    myFrmBG0200 = New frmBG0200(Me, True, "Adjust " & nodSelNode.Text, nodSelNode.Name, enumOperationCd.AdjustBudgetDirectInput)
                    myFrmBG0200.Show()

                    '// Add to screen list
                    p_frmBG0200.Add(myFrmBG0200)
                Else
                    '// Bring to front when exist.
                    If myFrmBG0200.WindowState = FormWindowState.Minimized Then
                        myFrmBG0200.WindowState = FormWindowState.Normal
                    End If
                    myFrmBG0200.BringToFront()
                End If

            End If
        End If
    End Sub

#End Region

#Region "Control Event"
    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CascadeToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TileVerticalToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TileHorizontalToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CloseAllToolStripMenuItem.Click
        '// Close all child forms of the parent.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private Sub frmBG0010_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If e.Cancel = True Then
            Exit Sub
        End If

        If MessageBox.Show("Are you sure to exit?", My.Settings.ProgramTitle, _
                           MessageBoxButtons.OKCancel, MessageBoxIcon.Question, _
                           MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Cancel Then
            e.Cancel = True
            Exit Sub
        End If

        ''// Save application settings
        ''My.Settings.Save()

        '// Unload PIC
        If p_strUserPIC <> "0000" Then
            myClsBG0010BL.ClearLockPIC()
        End If
    End Sub

    Private Sub frmBG0010_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '// Get Program Options
        myClsBG0010BL.GetOptions()

        '// Set Controls Attributes
        Me.Text = My.Settings.ProgramTitle
        Me.trvMenu.ExpandAll()
        Me.trvMenu.Focus()
        Me.lblToolbarStatus.Text = "Ready"
        Me.lblUserName.Text = "User Name: " & p_strUserName
        Me.lblPIC.Text = "PIC: " & p_strUserPIC
        Me.lblProgramVersion.Text = "Version: " & Application.ProductVersion

        '// Load Side Menu
        ShowSideMenu()

        '// Open Home Screen
        p_frmBG0110 = New frmBG0110(Me, "Home Page", True)
        p_frmBG0110.Show()

        '// Enable/Disable Menu
        EnableMenuStrip()

        '// Show Current Mode
        If p_blnReadOnlyMode = True Then
            lblCurrMode.Text = "[Read-Only Mode]"
            lblCurrMode.ForeColor = Color.Red
        Else
            lblCurrMode.Text = "" '"[Normal Mode]"
        End If
    End Sub

    Private Sub trvMenu_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles trvMenu.KeyDown
        If e.KeyCode = Keys.Enter Then
            OpenMenuFromTree(trvMenu.SelectedNode)
        End If
    End Sub

    Private Sub trvMenu_NodeMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles trvMenu.NodeMouseDoubleClick
        If e.Button = Windows.Forms.MouseButtons.Left Then
            OpenMenuFromTree(e.Node)
        End If
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        Dim myFrmBG0011 As New frmBG0011()
        myFrmBG0011.ShowDialog()
    End Sub

    Private Sub OptionsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptionsToolStripMenuItem.Click
        Dim myFrmBG0012 As New frmBG0012("Options")
        myFrmBG0012.ShowDialog()
    End Sub

    Private Sub trvMenu_AfterCollapse(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles trvMenu.AfterCollapse
        If e.Node.Name <> "nodInfo" And e.Node.Name <> "nodFavorites" Then
            e.Node.ImageKey = "folder.ico"
            e.Node.SelectedImageKey = "folder.ico"
        End If
    End Sub

    Private Sub trvMenu_AfterExpand(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles trvMenu.AfterExpand
        If e.Node.Name <> "nodInfo" And e.Node.Name <> "nodFavorites" Then
            e.Node.ImageKey = "opened_folder.ico"
            e.Node.SelectedImageKey = "opened_folder.ico"
        End If
    End Sub

#End Region

End Class
