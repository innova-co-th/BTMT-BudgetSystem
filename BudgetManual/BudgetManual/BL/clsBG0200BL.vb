Imports System.Data.SqlClient
Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class clsBG0200BL

#Region "Variable"
    Private myBudgetYear As String = String.Empty
    Private myPeriodType As String = String.Empty
    Private myBudgetType As String = String.Empty
    Private myProjectNo As String = String.Empty
    Private myBudgetOrder As String = String.Empty
    Private myUserPIC As String = String.Empty
    Private myUserId As String = String.Empty
    Private myOrderList As DataTable
    Private myBudgetList As DataTable
    Private myStatus As String = String.Empty
    Private myRRT(5) As String
    Private myPicList As DataTable
    Private myChildPicList As DataTable
    Private myUpDataList As DataTable
    Private myMTPHighlightValue As String = String.Empty
    Private myWorkingBG(2) As String
    Private myRevNo As String = String.Empty
    Private myRevNo2 As String = String.Empty
    Private myTransferList As DataTable
    Private myOperationCd As Integer
    Private myAutoMailList As DataTable
    Private myRefRevNo As String
    Private myRevNoList As DataTable
    Private myUpdateUser As String = String.Empty
    Private myUpdateDate As String = String.Empty
    Private myLogOperationCd As Integer
    Private myWKH1 As String = String.Empty
    Private myWKH2 As String = String.Empty
    Private myMTP_SUM1 As String = String.Empty
    Private myMTP_SUM2 As String = String.Empty
    Private myMTP_SUM3 As String = String.Empty
    Private myMTP_SUM4 As String = String.Empty
    Private myMTP_SUM5 As String = String.Empty
    Private myReviseRevNo As String = String.Empty
    Private myPrevMTPRevNo As String = String.Empty
    Private myWKRRT1 As String = String.Empty
    Private myWKRRT2 As String = String.Empty
    Private myWKRRT3 As String = String.Empty
    Private myWKRRT4 As String = String.Empty
    Private myWKRRT5 As String = String.Empty
    Private myMTP_PY_SUM1 As String = String.Empty
    Private myMTP_PY_SUM2 As String = String.Empty
    Private myMTP_PY_SUM3 As String = String.Empty
    Private myMTP_PY_SUM4 As String = String.Empty
    Private myMTP_PY_SUM5 As String = String.Empty
    Private myMtpProjectNo As String = String.Empty
    Private myMtpRevNo As String = String.Empty
    Private myMTPWB As String = String.Empty
    Private myDtSave As DataTable = Nothing
#End Region

#Region "Property"

#Region "BudgetYear"
    Property BudgetYear() As String
        Get
            Return myBudgetYear
        End Get
        Set(ByVal value As String)
            myBudgetYear = value
        End Set
    End Property
#End Region

#Region "PeriodType"
    Property PeriodType() As String
        Get
            Return myPeriodType
        End Get
        Set(ByVal value As String)
            myPeriodType = value
        End Set
    End Property
#End Region

#Region "BudgetType"
    Property BudgetType() As String
        Get
            Return myBudgetType
        End Get
        Set(ByVal value As String)
            myBudgetType = value
        End Set
    End Property
#End Region

#Region "ProjectNo"
    Property ProjectNo() As String
        Get
            Return myProjectNo
        End Get
        Set(ByVal value As String)
            myProjectNo = value
        End Set
    End Property
#End Region

#Region "BudgetOrder"
    Property BudgetOrder() As String
        Get
            Return myBudgetOrder
        End Get
        Set(ByVal value As String)
            myBudgetOrder = value
        End Set
    End Property
#End Region

#Region "UserPIC"
    Property UserPIC() As String
        Get
            Return myUserPIC
        End Get
        Set(ByVal value As String)
            myUserPIC = value
        End Set
    End Property
#End Region

#Region "UserId"
    Property UserId() As String
        Get
            Return myUserId
        End Get
        Set(ByVal value As String)
            myUserId = value
        End Set
    End Property
#End Region

#Region "OrderList"
    Property OrderList() As DataTable
        Get
            Return myOrderList
        End Get
        Set(ByVal value As DataTable)
            myOrderList = value
        End Set
    End Property
#End Region

#Region "BudgetList"
    Property BudgetList() As DataTable
        Get
            Return myBudgetList
        End Get
        Set(ByVal value As DataTable)
            myBudgetList = value
        End Set
    End Property
#End Region

#Region "Status"
    Property Status() As String
        Get
            Return myStatus
        End Get
        Set(ByVal value As String)
            myStatus = value
        End Set
    End Property
#End Region

#Region "RRT"
    Property RRT() As String()
        Get
            Return myRRT
        End Get
        Set(ByVal value As String())
            myRRT = value
        End Set
    End Property
#End Region

#Region "PicList"
    Property PicList() As DataTable
        Get
            Return myPicList
        End Get
        Set(ByVal value As DataTable)
            myPicList = value
        End Set
    End Property
#End Region

#Region "ChildPicList"
    Property ChildPicList() As DataTable
        Get
            Return myChildPicList
        End Get
        Set(ByVal value As DataTable)
            myChildPicList = value
        End Set
    End Property
#End Region

#Region "UpDataList"
    Property UpDataList() As DataTable
        Get
            Return myUpDataList
        End Get
        Set(ByVal value As DataTable)
            myUpDataList = value
        End Set
    End Property
#End Region

#Region "MTPHighlightValue"
    Property MTPHighlightValue() As String
        Get
            Return myMtpHighlightValue
        End Get
        Set(ByVal value As String)
            myMtpHighlightValue = value
        End Set
    End Property
#End Region

#Region "WorkingBG"
    Property WorkingBG() As String()
        Get
            Return myWorkingBG
        End Get
        Set(ByVal value As String())
            myWorkingBG = value
        End Set
    End Property
#End Region

#Region "RevNo"
    Property RevNo() As String
        Get
            Return myRevNo
        End Get
        Set(ByVal value As String)
            myRevNo = value
        End Set
    End Property
#End Region

#Region "ReviseRevNo"
    Property ReviseRevNo() As String
        Get
            Return myReviseRevNo
        End Get
        Set(ByVal value As String)
            myReviseRevNo = value
        End Set
    End Property
#End Region

#Region "PrevMTPRevNo"
    Property PrevMTPRevNo() As String
        Get
            Return myPrevMTPRevNo
        End Get
        Set(ByVal value As String)
            myPrevMTPRevNo = value
        End Set
    End Property
#End Region

#Region "TransferList"
    Property TransferList() As DataTable
        Get
            Return myTransferList
        End Get
        Set(ByVal value As DataTable)
            myTransferList = value
        End Set
    End Property
#End Region

#Region "OperationCd"
    Property OperationCd() As Integer
        Get
            Return myOperationCd
        End Get
        Set(ByVal value As Integer)
            myOperationCd = value
        End Set
    End Property
#End Region

#Region "AutoMailList"
    Property AutoMailList() As DataTable
        Get
            Return myAutoMailList
        End Get
        Set(ByVal value As DataTable)
            myAutoMailList = value
        End Set
    End Property
#End Region

#Region "RefRevNo"
    Property RefRevNo() As String
        Get
            Return myRefRevNo
        End Get
        Set(ByVal value As String)
            myRefRevNo = value
        End Set
    End Property
#End Region

#Region "RevNoList"
    Property RevNoList() As DataTable
        Get
            Return myRevNoList
        End Get
        Set(ByVal value As DataTable)
            myRevNoList = value
        End Set
    End Property
#End Region

#Region "UpdateUser"
    Property UpdateUser() As String
        Get
            Return myUpdateUser
        End Get
        Set(ByVal value As String)
            myUpdateUser = value
        End Set
    End Property
#End Region

#Region "UpdateDate"
    Property UpdateDate() As String
        Get
            Return myUpdateDate
        End Get
        Set(ByVal value As String)
            myUpdateDate = value
        End Set
    End Property
#End Region

#Region "LogOperationCd"
    Property LogOperationCd() As Integer
        Get
            Return myLogOperationCd
        End Get
        Set(ByVal value As Integer)
            myLogOperationCd = value
        End Set
    End Property
#End Region

#Region "RevNo2"
    Property RevNo2() As String
        Get
            Return myRevNo2
        End Get
        Set(ByVal value As String)
            myRevNo2 = value
        End Set
    End Property
#End Region

#Region "WKH1"
    Property WKH1() As String
        Get
            Return myWKH1
        End Get
        Set(ByVal value As String)
            myWKH1 = value
        End Set
    End Property
#End Region

#Region "WKH2"
    Property WKH2() As String
        Get
            Return myWKH2
        End Get
        Set(ByVal value As String)
            myWKH2 = value
        End Set
    End Property
#End Region

#Region "MTP_SUM1"
    Property MTP_SUM1() As String
        Get
            Return myMTP_SUM1
        End Get
        Set(ByVal value As String)
            myMTP_SUM1 = value
        End Set
    End Property
#End Region

#Region "MTP_SUM2"
    Property MTP_SUM2() As String
        Get
            Return myMTP_SUM2
        End Get
        Set(ByVal value As String)
            myMTP_SUM2 = value
        End Set
    End Property
#End Region

#Region "MTP_SUM3"
    Property MTP_SUM3() As String
        Get
            Return myMTP_SUM3
        End Get
        Set(ByVal value As String)
            myMTP_SUM3 = value
        End Set
    End Property
#End Region

#Region "MTP_SUM4"
    Property MTP_SUM4() As String
        Get
            Return myMTP_SUM4
        End Get
        Set(ByVal value As String)
            myMTP_SUM4 = value
        End Set
    End Property
#End Region

#Region "MTP_SUM5"
    Property MTP_SUM5() As String
        Get
            Return myMTP_SUM5
        End Get
        Set(ByVal value As String)
            myMTP_SUM5 = value
        End Set
    End Property
#End Region

#Region "WKRRT1"
    Property WKRRT1() As String
        Get
            Return myWKRRT1
        End Get
        Set(ByVal value As String)
            myWKRRT1 = value
        End Set
    End Property
#End Region

#Region "WKRRT2"
    Property WKRRT2() As String
        Get
            Return myWKRRT2
        End Get
        Set(ByVal value As String)
            myWKRRT2 = value
        End Set
    End Property
#End Region

#Region "WKRRT3"
    Property WKRRT3() As String
        Get
            Return myWKRRT3
        End Get
        Set(ByVal value As String)
            myWKRRT3 = value
        End Set
    End Property
#End Region

#Region "WKRRT4"
    Property WKRRT4() As String
        Get
            Return myWKRRT4
        End Get
        Set(ByVal value As String)
            myWKRRT4 = value
        End Set
    End Property
#End Region

#Region "WKRRT5"
    Property WKRRT5() As String
        Get
            Return myWKRRT5
        End Get
        Set(ByVal value As String)
            myWKRRT5 = value
        End Set
    End Property
#End Region

#Region "MTP_PY_SUM1"
    Property MTP_PY_SUM1() As String
        Get
            Return myMTP_PY_SUM1
        End Get
        Set(ByVal value As String)
            myMTP_PY_SUM1 = value
        End Set
    End Property
#End Region

#Region "MTP_PY_SUM2"
    Property MTP_PY_SUM2() As String
        Get
            Return myMTP_PY_SUM2
        End Get
        Set(ByVal value As String)
            myMTP_PY_SUM2 = value
        End Set
    End Property
#End Region

#Region "MTP_PY_SUM3"
    Property MTP_PY_SUM3() As String
        Get
            Return myMTP_PY_SUM3
        End Get
        Set(ByVal value As String)
            myMTP_PY_SUM3 = value
        End Set
    End Property
#End Region

#Region "MTP_PY_SUM4"
    Property MTP_PY_SUM4() As String
        Get
            Return myMTP_PY_SUM4
        End Get
        Set(ByVal value As String)
            myMTP_PY_SUM4 = value
        End Set
    End Property
#End Region

#Region "MTP_PY_SUM5"
    Property MTP_PY_SUM5() As String
        Get
            Return myMTP_PY_SUM5
        End Get
        Set(ByVal value As String)
            myMTP_PY_SUM5 = value
        End Set
    End Property
#End Region

#Region "MTPWB"
    Property MTPWB() As String
        Get
            Return myMTPWB
        End Get
        Set(ByVal value As String)
            myMTPWB = value
        End Set
    End Property
#End Region

#Region "MtpProjectNo"
    Property MtpProjectNo() As String
        Get
            Return myMtpProjectNo
        End Get
        Set(ByVal value As String)
            myMtpProjectNo = value
        End Set
    End Property
#End Region

#Region "MtpRevNo"
    Property MtpRevNo() As String
        Get
            Return myMtpRevNo
        End Get
        Set(ByVal value As String)
            myMtpRevNo = value
        End Set
    End Property
#End Region

#Region "dtSave"
    Property dtSave() As DataTable
        Get
            Return myDtSave
        End Get
        Set(ByVal value As DataTable)
            myDtSave = value
        End Set
    End Property
#End Region

#End Region

#Region "Function"

    Private Function SearchBudgetOrder() As Boolean
        Dim clsBG_M_BUDGET_ORDER As New BG_M_BUDGET_ORDER
        Dim rtn As Boolean

        '// Set Parameters
        clsBG_M_BUDGET_ORDER.BudgetType = Me.BudgetType
        clsBG_M_BUDGET_ORDER.UserPIC = Me.UserPIC

        If Me.UserPIC = "0000" Then
            rtn = clsBG_M_BUDGET_ORDER.Select011()
        Else
            rtn = clsBG_M_BUDGET_ORDER.Select001()
        End If

        '// Call Function
        If rtn = True Then
            Me.OrderList = clsBG_M_BUDGET_ORDER.dtResult

            Return True
        Else
            Me.OrderList = Nothing

            Return False
        End If
    End Function

    Public Function SearchNewBudgetOrder() As Boolean
        Dim clsBG_M_BUDGET_ORDER As New BG_M_BUDGET_ORDER
        Dim rtn As Boolean

        '// Set Parameters
        clsBG_M_BUDGET_ORDER.BudgetYear = Me.BudgetYear
        clsBG_M_BUDGET_ORDER.PeriodType = Me.PeriodType
        clsBG_M_BUDGET_ORDER.RevNo = Me.RevNo
        clsBG_M_BUDGET_ORDER.BudgetType = Me.BudgetType
        clsBG_M_BUDGET_ORDER.UserPIC = Me.UserPIC
        clsBG_M_BUDGET_ORDER.ProjectNo = Me.ProjectNo

        If Me.UserPIC = "0000" Then
            rtn = clsBG_M_BUDGET_ORDER.Select014()
        Else
            rtn = clsBG_M_BUDGET_ORDER.Select013()
        End If

        '// Call Function
        If rtn = True Then
            Me.OrderList = clsBG_M_BUDGET_ORDER.dtResult

            Return True
        Else
            Me.OrderList = Nothing

            Return False
        End If
    End Function

    Public Function GetBudgetHeader() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER
        Dim blnRtn As Boolean

        '// Set Parameters
        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_HEADER.UserPIC = Me.UserPIC
        clsBG_T_BUDGET_HEADER.RevNo = Me.RevNo
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

        '// Call Function: Select Budget Header by PIC
        If Me.OperationCd = enumOperationCd.AdjustBudget Or _
        Me.myOperationCd = enumOperationCd.AdjustBudgetDirectInput Or _
        Me.OperationCd = enumOperationCd.Authorize1 Or _
        Me.OperationCd = enumOperationCd.Authorize2 Then
            If Me.UserPIC = "0000" Then
                blnRtn = clsBG_T_BUDGET_HEADER.Select008()
            Else
                blnRtn = clsBG_T_BUDGET_HEADER.Select002()
            End If
        Else
            If Me.UserPIC = "0000" Then
                blnRtn = clsBG_T_BUDGET_HEADER.Select001()
            Else
                blnRtn = clsBG_T_BUDGET_HEADER.Select002()
            End If
        End If

        If blnRtn = True AndAlso clsBG_T_BUDGET_HEADER.dtResult.Rows.Count > 0 Then
            Me.Status = CStr(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("STATUS"))
            Me.RRT(0) = CStr(Nz(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("RRT0"), 0))
            Me.RRT(1) = CStr(Nz(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("RRT1"), 0))
            Me.RRT(2) = CStr(Nz(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("RRT2"), 0))
            Me.RRT(3) = CStr(Nz(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("RRT3"), 0))
            Me.RRT(4) = CStr(Nz(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("RRT4"), 0))
            Me.RRT(5) = CStr(Nz(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("RRT5"), 0))
            Me.WorkingBG(1) = CStr(Nz(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("WORKING_BG1"), 0))
            Me.WorkingBG(2) = CStr(Nz(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("WORKING_BG2"), 0))
            Me.RefRevNo = CStr(Nz(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("REF_REV_NO"), ""))

            Me.UpdateUser = CStr(Nz(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("UPDATE_USER_NAME")))
            Me.UpdateDate = CStr(Nz(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("UPDATE_DATE")))

            Return True
        Else
            Return False

        End If
    End Function

    Public Function GetBudgetHeaderBYSTATUS() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER
        Dim blnRtn As Boolean

        '// Set Parameters
        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_HEADER.UserPIC = Me.UserPIC
        clsBG_T_BUDGET_HEADER.RevNo = Me.RevNo
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo
        clsBG_T_BUDGET_HEADER.Status = Me.Status

        '// Call Function: Select Budget Header by PIC
        If Me.OperationCd = enumOperationCd.AdjustBudget Or _
        Me.myOperationCd = enumOperationCd.AdjustBudgetDirectInput Or _
        Me.OperationCd = enumOperationCd.Authorize1 Or _
        Me.OperationCd = enumOperationCd.Authorize2 Then
            If Me.UserPIC = "0000" Then
                blnRtn = clsBG_T_BUDGET_HEADER.Select008_1()
            Else
                blnRtn = clsBG_T_BUDGET_HEADER.Select002_1()
            End If
        Else
            If Me.UserPIC = "0000" Then
                blnRtn = clsBG_T_BUDGET_HEADER.Select001_1()
            Else
                blnRtn = clsBG_T_BUDGET_HEADER.Select002_1()
            End If
        End If

        If blnRtn = True AndAlso clsBG_T_BUDGET_HEADER.dtResult.Rows.Count > 0 Then
            Me.Status = CStr(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("STATUS"))
            Me.RRT(0) = CStr(Nz(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("RRT0"), 0))
            Me.RRT(1) = CStr(Nz(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("RRT1"), 0))
            Me.RRT(2) = CStr(Nz(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("RRT2"), 0))
            Me.RRT(3) = CStr(Nz(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("RRT3"), 0))
            Me.RRT(4) = CStr(Nz(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("RRT4"), 0))
            Me.RRT(5) = CStr(Nz(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("RRT5"), 0))
            Me.WorkingBG(1) = CStr(Nz(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("WORKING_BG1"), 0))
            Me.WorkingBG(2) = CStr(Nz(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("WORKING_BG2"), 0))
            Me.RefRevNo = CStr(Nz(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("REF_REV_NO"), ""))

            Me.UpdateUser = CStr(Nz(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("UPDATE_USER_NAME")))
            Me.UpdateDate = CStr(Nz(clsBG_T_BUDGET_HEADER.dtResult.Rows(0).Item("UPDATE_DATE")))

            Return True
        Else
            Return False

        End If
    End Function

    Public Function GetBudgetData() As Boolean
        Dim clsBG_T_BUDGET_DATA As New BG_T_BUDGET_DATA
        Dim clsBG_T_BUDGET_REFERENCE As New BG_T_BUDGET_REFERENCE

        Dim blnRtn As Boolean = False

        '// Set Parameters
        clsBG_T_BUDGET_DATA.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_DATA.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_DATA.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_DATA.UserPIC = Me.UserPIC
        clsBG_T_BUDGET_DATA.RevNo = Me.RevNo
        clsBG_T_BUDGET_DATA.ProjectNo = Me.ProjectNo

        '// Get Reference Budget
        clsBG_T_BUDGET_DATA.RefBudgetYear = "1"
        clsBG_T_BUDGET_DATA.RefPeriodType = "1"
        clsBG_T_BUDGET_DATA.RefProjectNo = "1"
        clsBG_T_BUDGET_DATA.RefRevNo = "1"
        clsBG_T_BUDGET_DATA.RefEstProjectNo = "1"
        clsBG_T_BUDGET_DATA.RefEstRevNo = "1"
        clsBG_T_BUDGET_DATA.RefRBProjectNo = "1"
        clsBG_T_BUDGET_DATA.RefRBRevNo = "1"
        clsBG_T_BUDGET_DATA.MtpProjectNo = "1"
        clsBG_T_BUDGET_DATA.MtpRevNo = "1"

        If Me.BudgetType <> BGConstant.P_BUDGET_TYPE_ASSET Then

            If Me.PeriodType = CStr(enumPeriodType.OriginalBudget) Then

                '// Ref. Estimate 
                clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
                clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
                clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
                clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(enumPeriodType.EstimateBudget)

                blnRtn = clsBG_T_BUDGET_REFERENCE.Select001()

                If blnRtn = False Then

                    Me.BudgetList = Nothing

                    Return False

                End If

                If clsBG_T_BUDGET_REFERENCE.dtResult.Rows.Count > 0 Then

                    clsBG_T_BUDGET_DATA.RefEstProjectNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_PROJECT_NO").ToString
                    clsBG_T_BUDGET_DATA.RefEstRevNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_REV_NO").ToString

                End If

                '// Ref. MTP
                clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
                clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
                clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
                clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(enumPeriodType.MTPBudget)

                blnRtn = clsBG_T_BUDGET_REFERENCE.Select001()

                If blnRtn = False Then

                    Me.BudgetList = Nothing

                    Return False

                End If

                If clsBG_T_BUDGET_REFERENCE.dtResult.Rows.Count > 0 Then

                    clsBG_T_BUDGET_DATA.RefBudgetYear = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_BUDGET_YEAR").ToString
                    clsBG_T_BUDGET_DATA.RefPeriodType = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_PERIOD_TYPE").ToString
                    clsBG_T_BUDGET_DATA.RefProjectNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_PROJECT_NO").ToString
                    clsBG_T_BUDGET_DATA.RefRevNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_REV_NO").ToString

                End If

            ElseIf Me.PeriodType = CStr(enumPeriodType.MTPBudget) Then

                '// Ref. Revise  
                clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
                clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
                clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
                clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(enumPeriodType.ReviseBudget)

                blnRtn = clsBG_T_BUDGET_REFERENCE.Select001()

                If blnRtn = False Then

                    Me.BudgetList = Nothing

                    Return False

                End If

                If clsBG_T_BUDGET_REFERENCE.dtResult.Rows.Count > 0 Then

                    clsBG_T_BUDGET_DATA.RefRBProjectNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_PROJECT_NO").ToString
                    clsBG_T_BUDGET_DATA.RefRBRevNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_REV_NO").ToString

                End If

                '// Ref. MTP Previous
                clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
                clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
                clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
                clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(enumPeriodType.MTPBudget)

                blnRtn = clsBG_T_BUDGET_REFERENCE.Select001()

                If blnRtn = False Then

                    Me.BudgetList = Nothing

                    Return False

                End If

                If clsBG_T_BUDGET_REFERENCE.dtResult.Rows.Count > 0 Then

                    clsBG_T_BUDGET_DATA.RefBudgetYear = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_BUDGET_YEAR").ToString
                    clsBG_T_BUDGET_DATA.RefPeriodType = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_PERIOD_TYPE").ToString
                    clsBG_T_BUDGET_DATA.RefProjectNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_PROJECT_NO").ToString
                    clsBG_T_BUDGET_DATA.RefRevNo = clsBG_T_BUDGET_REFERENCE.dtResult.Rows(0)("REF_REV_NO").ToString

                End If

            End If

        End If


        '// Call Function: Select Budget Header by PIC
        If Me.OperationCd = enumOperationCd.AdjustBudget Or _
        Me.myOperationCd = enumOperationCd.AdjustBudgetDirectInput Or _
        Me.OperationCd = enumOperationCd.Authorize1 Or _
        Me.OperationCd = enumOperationCd.Authorize2 Then
            clsBG_T_BUDGET_DATA.Status = CStr(enumBudgetStatus.Approve)

            If Me.PeriodType = CStr(enumPeriodType.MTPBudget) Then
                If Me.UserPIC = "0000" Then
                    blnRtn = clsBG_T_BUDGET_DATA.Select027()
                Else
                    blnRtn = clsBG_T_BUDGET_DATA.Select025()
                End If
            Else
                If Me.UserPIC = "0000" Then
                    blnRtn = clsBG_T_BUDGET_DATA.Select013()
                Else
                    blnRtn = clsBG_T_BUDGET_DATA.Select001()
                End If
            End If


        ElseIf Me.OperationCd = enumOperationCd.ViewBudget Then

            If Me.PeriodType = CStr(enumPeriodType.MTPBudget) Then
                If Me.UserPIC = "0000" Then
                    blnRtn = clsBG_T_BUDGET_DATA.Select026()
                Else
                    blnRtn = clsBG_T_BUDGET_DATA.Select025_1()
                End If
            Else
                If Me.UserPIC = "0000" Then
                    blnRtn = clsBG_T_BUDGET_DATA.Select002()
                Else
                    blnRtn = clsBG_T_BUDGET_DATA.Select001_1()
                End If

            End If

        Else

            If Me.PeriodType = CStr(enumPeriodType.MTPBudget) Then
                If Me.UserPIC = "0000" Then
                    blnRtn = clsBG_T_BUDGET_DATA.Select024()
                Else
                    blnRtn = clsBG_T_BUDGET_DATA.Select023()
                End If
            Else
                If Me.UserPIC = "0000" Then
                    blnRtn = clsBG_T_BUDGET_DATA.Select021()
                Else
                    blnRtn = clsBG_T_BUDGET_DATA.Select020()
                End If
            End If

        End If

        If blnRtn = True Then
            Me.BudgetList = clsBG_T_BUDGET_DATA.dtResult
            '--Check ReInput 

            Return True
        Else
            Me.BudgetList = Nothing

            Return False
        End If
    End Function

    Public Function CreateBudgetData() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER
        Dim clsBG_T_BUDGET_DATA As New BG_T_BUDGET_DATA
        Dim clsBG_T_BUDGET_ADJUST As New BG_T_BUDGET_ADJUST
        Dim clsBG_T_BUDGET_REFERENCE As New BG_T_BUDGET_REFERENCE
        Dim conn As SqlConnection
        Dim trans As SqlTransaction
        Dim blnHv As Boolean

        '// Query Related budget order
        If SearchBudgetOrder() = False OrElse Me.OrderList.Rows.Count = 0 Then
            Return False
        End If

        '// Create Connection
        conn = New SqlConnection(My.Settings.ConnStr)
        conn.Open()

        '// Begin Transaction
        trans = conn.BeginTransaction()

        Try
            '// Add Budget Data Header
            If Me.UserPIC = "0000" Then
                For i = 1 To Me.PicList.Rows.Count - 1 '// Skip "0000"
                    '// Set Parameters
                    clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
                    clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
                    clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
                    clsBG_T_BUDGET_HEADER.UserPIC = CStr(PicList.Rows(i).Item("PERSON_IN_CHARGE_NO"))
                    clsBG_T_BUDGET_HEADER.RevNo = Me.RevNo
                    clsBG_T_BUDGET_HEADER.Status = CStr(enumBudgetStatus.NewRecord)
                    clsBG_T_BUDGET_HEADER.UserId = Me.UserId
                    clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

                    '// Call Function: Insert Budget Header
                    If clsBG_T_BUDGET_HEADER.Insert001(conn, trans) = False Then
                        Throw New Exception("Can not insert budget header!")
                    End If
                Next
            Else
                '// Set Parameters
                clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
                clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
                clsBG_T_BUDGET_HEADER.UserPIC = Me.UserPIC
                clsBG_T_BUDGET_HEADER.RevNo = Me.RevNo
                clsBG_T_BUDGET_HEADER.Status = CStr(enumBudgetStatus.NewRecord)
                clsBG_T_BUDGET_HEADER.UserId = Me.UserId
                clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

                '// Call Function: Insert Budget Header
                If clsBG_T_BUDGET_HEADER.Insert001(conn, trans) = False Then
                    Throw New Exception("Can not insert budget header!")
                End If
            End If

            For Each dr As DataRow In Me.OrderList.Rows
                blnHv = False
                '// Add Budget Data Detail
                '// Set Parameters
                clsBG_T_BUDGET_DATA.BudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_DATA.PeriodType = Me.PeriodType
                clsBG_T_BUDGET_DATA.BudgetOrderNo = CStr(dr("BUDGET_ORDER_NO"))
                clsBG_T_BUDGET_DATA.RevNo = Me.RevNo
                clsBG_T_BUDGET_DATA.UserId = Me.UserId
                clsBG_T_BUDGET_DATA.ProjectNo = Me.ProjectNo

                '//Select 
                blnHv = clsBG_T_BUDGET_DATA.Select032()
                If blnHv = True Then
                    ' Exist : Not Insert 
                Else
                    '// Call Function
                    If clsBG_T_BUDGET_DATA.Insert001(conn, trans) = False Then
                        Throw New Exception("Can not insert budget data!")
                    End If
                End If
            Next

            Dim dtPreDat As DataTable = Nothing

            '// Add Budget Data Header
            clsBG_T_BUDGET_ADJUST.BudgetYear = Me.BudgetYear
            clsBG_T_BUDGET_ADJUST.PeriodType = Me.PeriodType
            clsBG_T_BUDGET_ADJUST.ProjectNo = Me.ProjectNo

            '//-- Select Prev Rev data
            clsBG_T_BUDGET_ADJUST.RevNo = Me.RevNo2

            If clsBG_T_BUDGET_ADJUST.Select004() = True And clsBG_T_BUDGET_ADJUST.dtResult.Rows.Count > 0 Then
                dtPreDat = clsBG_T_BUDGET_ADJUST.dtResult
            Else
                dtPreDat = Nothing
            End If

            '//-- Insert adjust data
            clsBG_T_BUDGET_ADJUST.RevNo = Me.RevNo
            clsBG_T_BUDGET_ADJUST.CreateUserID = Me.UserId

            If dtPreDat IsNot Nothing Then
                Dim dr As DataRow = dtPreDat.Rows(0)
                clsBG_T_BUDGET_ADJUST.RRT0 = CStr(Nz(dr("RRT0"), 100))
                clsBG_T_BUDGET_ADJUST.RRT1 = CStr(Nz(dr("RRT1"), 100))
                clsBG_T_BUDGET_ADJUST.RRT2 = CStr(Nz(dr("RRT2"), 100))
                clsBG_T_BUDGET_ADJUST.RRT3 = CStr(Nz(dr("RRT3"), 100))
                clsBG_T_BUDGET_ADJUST.RRT4 = CStr(Nz(dr("RRT4"), 100))
                clsBG_T_BUDGET_ADJUST.RRT5 = CStr(Nz(dr("RRT5"), 100))
                clsBG_T_BUDGET_ADJUST.WorkingBG1 = CStr(Nz(dr("WORKING_BG1"), 100))
                clsBG_T_BUDGET_ADJUST.WorkingBG2 = CStr(Nz(dr("WORKING_BG2"), 100))
            Else
                clsBG_T_BUDGET_ADJUST.RRT0 = "100"
                clsBG_T_BUDGET_ADJUST.RRT1 = "100"
                clsBG_T_BUDGET_ADJUST.RRT2 = "100"
                clsBG_T_BUDGET_ADJUST.RRT3 = "100"
                clsBG_T_BUDGET_ADJUST.RRT4 = "100"
                clsBG_T_BUDGET_ADJUST.RRT5 = "100"
                clsBG_T_BUDGET_ADJUST.WorkingBG1 = "100"
                clsBG_T_BUDGET_ADJUST.WorkingBG2 = "100"
            End If

            '//-- Select Current Rev data
            If clsBG_T_BUDGET_ADJUST.Select004() = True And clsBG_T_BUDGET_ADJUST.dtResult.Rows.Count = 0 Then
                If clsBG_T_BUDGET_ADJUST.Insert001(conn, trans) = False Then
                    Throw New Exception("Can not insert budget adjust master!")
                End If
            End If


            '//-- Insert Budget Reference (Added by Kate 2013/04/30) ------------(+)

            If CInt(Me.PeriodType) = BGConstant.enumPeriodType.OriginalBudget Then
                '// Add Estimate 2nd half(Previous year)
                clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
                clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
                clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
                clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(BGConstant.enumPeriodType.EstimateBudget)
                clsBG_T_BUDGET_REFERENCE.RefBudgetYear = CStr(CInt(Me.BudgetYear) - 1)
                clsBG_T_BUDGET_REFERENCE.RefProjectNo = "1"
                clsBG_T_BUDGET_REFERENCE.RefRevNo = "1"
                clsBG_T_BUDGET_REFERENCE.CreateUserID = Me.UserId

                '// Check data exist
                If clsBG_T_BUDGET_REFERENCE.Select001() = True Then

                    If clsBG_T_BUDGET_REFERENCE.dtResult.Rows.Count = 0 Then
                        '// Add new record
                        If clsBG_T_BUDGET_REFERENCE.Insert001(conn, trans) = False Then
                            Throw New Exception("Can not insert budget reference!")
                        End If
                    End If
                Else
                    Throw New Exception("Can not insert budget reference!")
                End If


                '// Add MTP(Previous year)
                clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
                clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
                clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
                clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(BGConstant.enumPeriodType.MTPBudget)
                clsBG_T_BUDGET_REFERENCE.RefBudgetYear = CStr(CInt(Me.BudgetYear) - 1)
                clsBG_T_BUDGET_REFERENCE.RefProjectNo = "1"
                clsBG_T_BUDGET_REFERENCE.RefRevNo = "1"
                clsBG_T_BUDGET_REFERENCE.CreateUserID = Me.UserId

                '// Check data exist
                If clsBG_T_BUDGET_REFERENCE.Select001() = True Then

                    If clsBG_T_BUDGET_REFERENCE.dtResult.Rows.Count = 0 Then
                        '// Add new record
                        If clsBG_T_BUDGET_REFERENCE.Insert001(conn, trans) = False Then
                            Throw New Exception("Can not insert budget reference!")
                        End If
                    End If
                Else
                    Throw New Exception("Can not insert budget reference!")
                End If

            ElseIf CInt(Me.PeriodType) = BGConstant.enumPeriodType.MTPBudget Then
                '// Add Revise (Same year)
                clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
                clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
                clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
                clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(BGConstant.enumPeriodType.ReviseBudget)
                clsBG_T_BUDGET_REFERENCE.RefBudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_REFERENCE.RefProjectNo = "1"
                clsBG_T_BUDGET_REFERENCE.RefRevNo = "1"
                clsBG_T_BUDGET_REFERENCE.CreateUserID = Me.UserId

                '// Check data exist
                If clsBG_T_BUDGET_REFERENCE.Select001() = True Then

                    If clsBG_T_BUDGET_REFERENCE.dtResult.Rows.Count = 0 Then
                        '// Add new record
                        If clsBG_T_BUDGET_REFERENCE.Insert001(conn, trans) = False Then
                            Throw New Exception("Can not insert budget reference!")
                        End If
                    End If
                Else
                    Throw New Exception("Can not insert budget reference!")
                End If


                '// Add MTP(Previous year)
                clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
                clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
                clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
                clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(BGConstant.enumPeriodType.MTPBudget)
                clsBG_T_BUDGET_REFERENCE.RefBudgetYear = CStr(CInt(Me.BudgetYear) - 1)
                clsBG_T_BUDGET_REFERENCE.RefProjectNo = "1"
                clsBG_T_BUDGET_REFERENCE.RefRevNo = "1"
                clsBG_T_BUDGET_REFERENCE.CreateUserID = Me.UserId

                '// Check data exist
                If clsBG_T_BUDGET_REFERENCE.Select001() = True Then

                    If clsBG_T_BUDGET_REFERENCE.dtResult.Rows.Count = 0 Then
                        '// Add new record
                        If clsBG_T_BUDGET_REFERENCE.Insert001(conn, trans) = False Then
                            Throw New Exception("Can not insert budget reference!")
                        End If
                    End If
                Else
                    Throw New Exception("Can not insert budget reference!")
                End If

            End If

            '// Commit Transaction
            trans.Commit()

            '// Close Connection
            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception

            '// Rollback Transaction
            trans.Rollback()

            '// Close Connection
            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            MessageBox.Show("Error: " & ex.Message, My.Settings.ProgramTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)

            Return False
        End Try
    End Function

    Public Function CreateBudgetData2() As Boolean
        Dim clsBG_T_BUDGET_DATA As New BG_T_BUDGET_DATA
        Dim conn As SqlConnection
        Dim trans As SqlTransaction
        Dim dtNewOrder As DataTable

        If Me.OrderList.Rows.Count = 0 Then
            Exit Function
        End If

        dtNewOrder = Me.OrderList

        '// Create Connection
        conn = New SqlConnection(My.Settings.ConnStr)
        conn.Open()

        '// Begin Transaction
        trans = conn.BeginTransaction()

        Try
            For Each dr As DataRow In dtNewOrder.Rows
                '// Add Budget Data Detail
                '// Set Parameters
                clsBG_T_BUDGET_DATA.BudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_DATA.PeriodType = Me.PeriodType
                clsBG_T_BUDGET_DATA.BudgetOrderNo = CStr(dr("BUDGET_ORDER_NO"))
                clsBG_T_BUDGET_DATA.RevNo = Me.RevNo
                clsBG_T_BUDGET_DATA.UserId = Me.UserId
                clsBG_T_BUDGET_DATA.ProjectNo = Me.ProjectNo

                '// Call Function
                If clsBG_T_BUDGET_DATA.Insert001(conn, trans) = False Then
                    Throw New Exception("Can not insert budget data!")
                End If
            Next

            '// Commit Transaction
            trans.Commit()

            '// Close Connection
            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception

            '// Rollback Transaction
            trans.Rollback()

            '// Close Connection
            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            MessageBox.Show("Error: " & ex.Message, My.Settings.ProgramTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)

            Return False
        End Try
    End Function

    Public Function CreateBudgetData3(ByRef Conn As SqlConnection, ByRef Trans As SqlTransaction) As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER
        Dim clsBG_T_BUDGET_DATA As New BG_T_BUDGET_DATA
        Dim clsBG_T_BUDGET_ADJUST As New BG_T_BUDGET_ADJUST
        Dim clsBG_T_BUDGET_REFERENCE As New BG_T_BUDGET_REFERENCE
        Dim clsBG_T_BUDGET_COMMENT As New BG_T_BUDGET_COMMENT
        ''Dim conn As SqlConnection
        ''Dim trans As SqlTransaction

        '// Query Related budget order
        If SearchBudgetOrder() = False OrElse Me.OrderList.Rows.Count = 0 Then
            Return False
        End If

        '// Add Budget Data Header
        If Me.UserPIC = "0000" Then
            For i = 1 To Me.PicList.Rows.Count - 1 '// Skip "0000"
                '// Set Parameters
                clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
                clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
                clsBG_T_BUDGET_HEADER.UserPIC = CStr(PicList.Rows(i).Item("PERSON_IN_CHARGE_NO"))
                clsBG_T_BUDGET_HEADER.RevNo = Me.RevNo
                clsBG_T_BUDGET_HEADER.Status = CStr(enumBudgetStatus.NewRecord)
                clsBG_T_BUDGET_HEADER.UserId = Me.UserId
                clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

                '// Call Function: Insert Budget Header
                If clsBG_T_BUDGET_HEADER.Insert001(Conn, Trans) = False Then
                    ''Throw New Exception("Can not insert budget header!")
                    Return False
                End If
            Next
        Else
            '// Set Parameters
            clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
            clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
            clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
            clsBG_T_BUDGET_HEADER.UserPIC = Me.UserPIC
            clsBG_T_BUDGET_HEADER.RevNo = Me.RevNo
            clsBG_T_BUDGET_HEADER.Status = CStr(enumBudgetStatus.NewRecord)
            clsBG_T_BUDGET_HEADER.UserId = Me.UserId
            clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

            '// Call Function: Insert Budget Header
            If clsBG_T_BUDGET_HEADER.Insert001(Conn, Trans) = False Then
                ''Throw New Exception("Can not insert budget header!")
                Return False
            End If
        End If

        Dim dtComment As DataTable
        For Each dr As DataRow In Me.OrderList.Rows
            '// Add Budget Data Detail
            '// Set Parameters
            clsBG_T_BUDGET_DATA.BudgetYear = Me.BudgetYear
            clsBG_T_BUDGET_DATA.PeriodType = Me.PeriodType
            clsBG_T_BUDGET_DATA.BudgetOrderNo = CStr(dr("BUDGET_ORDER_NO"))
            clsBG_T_BUDGET_DATA.RevNo = Me.RevNo
            clsBG_T_BUDGET_DATA.UserId = Me.UserId
            clsBG_T_BUDGET_DATA.ProjectNo = Me.ProjectNo

            '// Call Function
            If clsBG_T_BUDGET_DATA.Insert001(Conn, Trans) = False Then
                ''Throw New Exception("Can not insert budget data!")
                Return False
            End If


            '//Add Budget Comment
            clsBG_T_BUDGET_COMMENT.BudgetYear = Me.BudgetYear
            clsBG_T_BUDGET_COMMENT.PeriodType = Me.PeriodType
            clsBG_T_BUDGET_COMMENT.BudgetOrderNo = CStr(dr("BUDGET_ORDER_NO"))
            clsBG_T_BUDGET_COMMENT.RevNo = Me.RevNo2
            clsBG_T_BUDGET_COMMENT.CreateUserId = Me.UserId
            clsBG_T_BUDGET_COMMENT.ProjectNo = Me.ProjectNo

            If clsBG_T_BUDGET_COMMENT.Select001 Then
                If Not clsBG_T_BUDGET_COMMENT.CommentList Is Nothing AndAlso clsBG_T_BUDGET_COMMENT.CommentList.Rows.Count > 0 Then
                    dtComment = clsBG_T_BUDGET_COMMENT.CommentList

                    clsBG_T_BUDGET_COMMENT.BudgetComment = dtComment.Rows(0)
                    clsBG_T_BUDGET_COMMENT.RevNo = Me.RevNo

                    '// Call Function
                    If clsBG_T_BUDGET_COMMENT.Insert002(Conn, Trans) = False Then
                        ''Throw New Exception("Can not insert budget data!")
                        Return False
                    End If

                End If
            End If


        Next


        Dim dtPreDat As DataTable = Nothing

        '// Add Budget Data Header
        clsBG_T_BUDGET_ADJUST.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_ADJUST.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_ADJUST.ProjectNo = Me.ProjectNo

        '//-- Select Prev Rev data
        clsBG_T_BUDGET_ADJUST.RevNo = Me.RevNo2

        If clsBG_T_BUDGET_ADJUST.Select004() = True And clsBG_T_BUDGET_ADJUST.dtResult.Rows.Count > 0 Then
            dtPreDat = clsBG_T_BUDGET_ADJUST.dtResult
        Else
            dtPreDat = Nothing
        End If

        '//-- Insert adjust data
        clsBG_T_BUDGET_ADJUST.RevNo = Me.RevNo
        clsBG_T_BUDGET_ADJUST.CreateUserID = Me.UserId

        If dtPreDat IsNot Nothing Then
            Dim dr As DataRow = dtPreDat.Rows(0)
            clsBG_T_BUDGET_ADJUST.RRT0 = CStr(Nz(dr("RRT0"), 100))
            clsBG_T_BUDGET_ADJUST.RRT1 = CStr(Nz(dr("RRT1"), 100))
            clsBG_T_BUDGET_ADJUST.RRT2 = CStr(Nz(dr("RRT2"), 100))
            clsBG_T_BUDGET_ADJUST.RRT3 = CStr(Nz(dr("RRT3"), 100))
            clsBG_T_BUDGET_ADJUST.RRT4 = CStr(Nz(dr("RRT4"), 100))
            clsBG_T_BUDGET_ADJUST.RRT5 = CStr(Nz(dr("RRT5"), 100))
            clsBG_T_BUDGET_ADJUST.WorkingBG1 = CStr(Nz(dr("WORKING_BG1"), 100))
            clsBG_T_BUDGET_ADJUST.WorkingBG2 = CStr(Nz(dr("WORKING_BG2"), 100))
        Else
            clsBG_T_BUDGET_ADJUST.RRT0 = "100"
            clsBG_T_BUDGET_ADJUST.RRT1 = "100"
            clsBG_T_BUDGET_ADJUST.RRT2 = "100"
            clsBG_T_BUDGET_ADJUST.RRT3 = "100"
            clsBG_T_BUDGET_ADJUST.RRT4 = "100"
            clsBG_T_BUDGET_ADJUST.RRT5 = "100"
            clsBG_T_BUDGET_ADJUST.WorkingBG1 = "100"
            clsBG_T_BUDGET_ADJUST.WorkingBG2 = "100"
        End If

        '//-- Select Current Rev data
        If clsBG_T_BUDGET_ADJUST.Select004() = True And clsBG_T_BUDGET_ADJUST.dtResult.Rows.Count = 0 Then
            If clsBG_T_BUDGET_ADJUST.Insert001(Conn, Trans) = False Then
                ''Throw New Exception("Can not insert budget adjust master!")
                Return False
            End If
        End If

        If CInt(Me.PeriodType) = BGConstant.enumPeriodType.OriginalBudget Then
            '// Add Estimate 2nd half(Previous year)
            clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
            clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
            clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
            clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
            clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(BGConstant.enumPeriodType.EstimateBudget)
            clsBG_T_BUDGET_REFERENCE.RefBudgetYear = CStr(CInt(Me.BudgetYear) - 1)
            clsBG_T_BUDGET_REFERENCE.RefProjectNo = "1"
            clsBG_T_BUDGET_REFERENCE.RefRevNo = "1"
            clsBG_T_BUDGET_REFERENCE.CreateUserID = Me.UserId

            '// Check data exist
            If clsBG_T_BUDGET_REFERENCE.Select001() = True Then

                If clsBG_T_BUDGET_REFERENCE.dtResult.Rows.Count = 0 Then
                    '// Add new record
                    If clsBG_T_BUDGET_REFERENCE.Insert001(Conn, Trans) = False Then
                        'Throw New Exception("Can not insert budget reference!")
                        Return False
                    End If
                End If
            Else
                'Throw New Exception("Can not insert budget reference!")
                Return False
            End If

            '// Add MTP(Previous year)
            clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
            clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
            clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
            clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
            clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(BGConstant.enumPeriodType.MTPBudget)
            clsBG_T_BUDGET_REFERENCE.RefBudgetYear = CStr(CInt(Me.BudgetYear) - 1)
            clsBG_T_BUDGET_REFERENCE.RefProjectNo = "1"
            clsBG_T_BUDGET_REFERENCE.RefRevNo = "1"
            clsBG_T_BUDGET_REFERENCE.CreateUserID = Me.UserId

            '// Check data exist
            If clsBG_T_BUDGET_REFERENCE.Select001() = True Then

                If clsBG_T_BUDGET_REFERENCE.dtResult.Rows.Count = 0 Then
                    '// Add new record
                    If clsBG_T_BUDGET_REFERENCE.Insert001(Conn, Trans) = False Then
                        'Throw New Exception("Can not insert budget reference!")
                        Return False
                    End If
                End If
            Else
                'Throw New Exception("Can not insert budget reference!")
                Return False
            End If

        ElseIf CInt(Me.PeriodType) = BGConstant.enumPeriodType.MTPBudget Then
            '// Add Revise (Same year)
            clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
            clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
            clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
            clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
            clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(BGConstant.enumPeriodType.ReviseBudget)
            clsBG_T_BUDGET_REFERENCE.RefBudgetYear = Me.BudgetYear
            clsBG_T_BUDGET_REFERENCE.RefProjectNo = "1"
            clsBG_T_BUDGET_REFERENCE.RefRevNo = "1"
            clsBG_T_BUDGET_REFERENCE.CreateUserID = Me.UserId

            '// Check data exist
            If clsBG_T_BUDGET_REFERENCE.Select001() = True Then

                If clsBG_T_BUDGET_REFERENCE.dtResult.Rows.Count = 0 Then
                    '// Add new record
                    If clsBG_T_BUDGET_REFERENCE.Insert001(Conn, Trans) = False Then
                        'Throw New Exception("Can not insert budget reference!")
                        Return False
                    End If
                End If
            Else
                'Throw New Exception("Can not insert budget reference!")
                Return False
            End If


            '// Add MTP(Previous year)
            clsBG_T_BUDGET_REFERENCE.BudgetYear = Me.BudgetYear
            clsBG_T_BUDGET_REFERENCE.PeriodType = Me.PeriodType
            clsBG_T_BUDGET_REFERENCE.ProjectNo = Me.ProjectNo
            clsBG_T_BUDGET_REFERENCE.RevNo = Me.RevNo
            clsBG_T_BUDGET_REFERENCE.RefPeriodType = CStr(BGConstant.enumPeriodType.MTPBudget)
            clsBG_T_BUDGET_REFERENCE.RefBudgetYear = CStr(CInt(Me.BudgetYear) - 1)
            clsBG_T_BUDGET_REFERENCE.RefProjectNo = "1"
            clsBG_T_BUDGET_REFERENCE.RefRevNo = "1"
            clsBG_T_BUDGET_REFERENCE.CreateUserID = Me.UserId

            '// Check data exist
            If clsBG_T_BUDGET_REFERENCE.Select001() = True Then

                If clsBG_T_BUDGET_REFERENCE.dtResult.Rows.Count = 0 Then
                    '// Add new record
                    If clsBG_T_BUDGET_REFERENCE.Insert001(Conn, Trans) = False Then
                        'Throw New Exception("Can not insert budget reference!")
                        Return False
                    End If
                End If
            Else
                'Throw New Exception("Can not insert budget reference!")
                Return False
            End If

        End If

       
        Return True
    End Function

    Public Function SaveBudgetData() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER
        Dim clsBG_T_BUDGET_DATA As New BG_T_BUDGET_DATA
        Dim conn As SqlConnection
        Dim trans As SqlTransaction

        '// Create Connection
        conn = New SqlConnection(My.Settings.ConnStr)
        conn.Open()

        '// Begin Transaction
        trans = conn.BeginTransaction()

        Try
            '// Update Budget Data Header
            If Me.UserPIC = "0000" Then
                For i = 1 To Me.PicList.Rows.Count - 1 '// Skip "0000"
                    '// Set Parameters
                    clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
                    clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
                    clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
                    clsBG_T_BUDGET_HEADER.UserPIC = CStr(PicList.Rows(i).Item("PERSON_IN_CHARGE_NO"))
                    clsBG_T_BUDGET_HEADER.RevNo = Me.RevNo
                    clsBG_T_BUDGET_HEADER.Status = Me.Status
                    clsBG_T_BUDGET_HEADER.UserId = Me.UserId
                    clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

                    '// Call Function
                    If clsBG_T_BUDGET_HEADER.Update001(conn, trans) = False Then
                        Throw New Exception("Can not update Budget header")
                    End If
                Next
            Else
                '// Set Parameters
                clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
                clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
                clsBG_T_BUDGET_HEADER.UserPIC = Me.UserPIC
                clsBG_T_BUDGET_HEADER.RevNo = Me.RevNo
                clsBG_T_BUDGET_HEADER.Status = Me.Status
                clsBG_T_BUDGET_HEADER.UserId = Me.UserId
                clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

                '// Call Function
                If clsBG_T_BUDGET_HEADER.Update001(conn, trans) = False Then
                    Throw New Exception("Can not update Budget header")
                End If
            End If

            '// Update Budget Data Detail
            For Each dr As DataRow In Me.BudgetList.Rows

                '// Set Parameters
                clsBG_T_BUDGET_DATA.BudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_DATA.PeriodType = Me.PeriodType
                clsBG_T_BUDGET_DATA.BudgetType = Me.BudgetType
                clsBG_T_BUDGET_DATA.BudgetOrderNo = CStr(dr("OrderNo"))
                clsBG_T_BUDGET_DATA.RevNo = Me.RevNo
                clsBG_T_BUDGET_DATA.BudgetData = dr
                clsBG_T_BUDGET_DATA.UserId = Me.UserId
                clsBG_T_BUDGET_DATA.ProjectNo = Me.ProjectNo

                '// Call Function
                If clsBG_T_BUDGET_DATA.Update001(conn, trans) = False Then
                    Throw New Exception("Can not update Budget data")
                End If

            Next

            '// Commit Transaction
            trans.Commit()

            '// Close Connection
            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception

            '// Rollback Transaction
            trans.Rollback()

            '// Close Connection
            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            MessageBox.Show("Error: " & ex.Message, My.Settings.ProgramTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)

            Return False
        End Try
    End Function
    Public Function SaveBudgetDataReInputByOrder() As Boolean
        'Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER
        Dim clsBG_T_BUDGET_DATA_REINPUT As New BG_T_BUDGET_DATA_REINPUT
        Dim clsBG_T_BUDGET_DATA As New BG_T_BUDGET_DATA
        Dim conn As SqlConnection
        Dim trans As SqlTransaction

        '// Create Connection
        conn = New SqlConnection(My.Settings.ConnStr)
        conn.Open()

        '// Begin Transaction
        trans = conn.BeginTransaction()

        Try
            '// Update Budget Data Header
            If Me.UserPIC = "0000" Then
                For i = 1 To Me.PicList.Rows.Count - 1 '// Skip "0000"
                    '// Set Parameters
                    clsBG_T_BUDGET_DATA_REINPUT.BudgetYear = Me.BudgetYear
                    clsBG_T_BUDGET_DATA_REINPUT.PeriodType = Me.PeriodType
                    clsBG_T_BUDGET_DATA_REINPUT.BudgetType = Me.BudgetType
                    clsBG_T_BUDGET_DATA_REINPUT.UserPIC = CStr(PicList.Rows(i).Item("PERSON_IN_CHARGE_NO"))
                    clsBG_T_BUDGET_DATA_REINPUT.RevNo = Me.RevNo
                    clsBG_T_BUDGET_DATA_REINPUT.ItemStatus = Me.Status
                    clsBG_T_BUDGET_DATA_REINPUT.UserId = Me.UserId
                    clsBG_T_BUDGET_DATA_REINPUT.ProjectNo = Me.ProjectNo

                    '// Call Function
                    If clsBG_T_BUDGET_DATA_REINPUT.Update001(conn, trans) = False Then
                        Throw New Exception("Can not update Budget header")
                    End If
                Next
            Else
                '// Set Parameters
                clsBG_T_BUDGET_DATA_REINPUT.BudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_DATA_REINPUT.PeriodType = Me.PeriodType
                clsBG_T_BUDGET_DATA_REINPUT.BudgetType = Me.BudgetType
                clsBG_T_BUDGET_DATA_REINPUT.UserPIC = Me.UserPIC
                clsBG_T_BUDGET_DATA_REINPUT.RevNo = Me.RevNo
                clsBG_T_BUDGET_DATA_REINPUT.ItemStatus = Me.Status
                clsBG_T_BUDGET_DATA_REINPUT.UserId = Me.UserId
                clsBG_T_BUDGET_DATA_REINPUT.ProjectNo = Me.ProjectNo

                '// Call Function
                If clsBG_T_BUDGET_DATA_REINPUT.Update001(conn, trans) = False Then
                    Throw New Exception("Can not update Budget header")
                End If
            End If

            '// Update Budget Data Detail
            For Each dr As DataRow In Me.BudgetList.Rows

                '// Set Parameters
                clsBG_T_BUDGET_DATA.BudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_DATA.PeriodType = Me.PeriodType
                clsBG_T_BUDGET_DATA.BudgetType = Me.BudgetType
                clsBG_T_BUDGET_DATA.BudgetOrderNo = CStr(dr("OrderNo"))
                clsBG_T_BUDGET_DATA.RevNo = Me.RevNo
                clsBG_T_BUDGET_DATA.BudgetData = dr
                clsBG_T_BUDGET_DATA.UserId = Me.UserId
                clsBG_T_BUDGET_DATA.ProjectNo = Me.ProjectNo

                '// Call Function
                If clsBG_T_BUDGET_DATA.Update001(conn, trans) = False Then
                    Throw New Exception("Can not update Budget data")
                End If

            Next

            '// Commit Transaction
            trans.Commit()

            '// Close Connection
            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception

            '// Rollback Transaction
            trans.Rollback()

            '// Close Connection
            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            MessageBox.Show("Error: " & ex.Message, My.Settings.ProgramTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)

            Return False
        End Try
    End Function
    Public Function SaveSubmitBudgetData() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER

        '// Update Budget Data Header
        '// Set Parameters
        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_HEADER.UserPIC = Me.UserPIC
        clsBG_T_BUDGET_HEADER.RevNo = "1"
        clsBG_T_BUDGET_HEADER.Status = CStr(enumBudgetStatus.Submit)
        clsBG_T_BUDGET_HEADER.UserId = Me.UserId
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

        '// Call Function: Update budget status
        If clsBG_T_BUDGET_HEADER.Update002() = False Then

            Return False
        Else
            Return True

        End If
    End Function

    Public Function SaveRejectBudgetData() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER

        '// Update Budget Data Header
        '// Set Parameters
        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_HEADER.UserPIC = Me.UserPIC
        clsBG_T_BUDGET_HEADER.RevNo = "1"
        clsBG_T_BUDGET_HEADER.Status = CStr(enumBudgetStatus.NewRecord)
        clsBG_T_BUDGET_HEADER.UserId = Me.UserId
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

        '// Call Function: Update budget status
        If clsBG_T_BUDGET_HEADER.Update002() = False Then

            Return False
        Else
            Return True

        End If
    End Function
    Public Function SaveRejectBudgetDataReInputByOrder() As Boolean
        Dim clsBG_T_BUDGET_DATA_REINPUT As New BG_T_BUDGET_DATA_REINPUT

        '// Update Budget Data Header
        '// Set Parameters
        clsBG_T_BUDGET_DATA_REINPUT.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_DATA_REINPUT.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_DATA_REINPUT.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_DATA_REINPUT.UserPIC = Me.UserPIC
        clsBG_T_BUDGET_DATA_REINPUT.RevNo = "1"
        clsBG_T_BUDGET_DATA_REINPUT.ItemStatus = CStr(enumBudgetStatus.NewRecord)
        clsBG_T_BUDGET_DATA_REINPUT.UserId = Me.UserId
        clsBG_T_BUDGET_DATA_REINPUT.ProjectNo = Me.ProjectNo

        '// Call Function: Update budget status
        If clsBG_T_BUDGET_DATA_REINPUT.Update001() = False Then

            Return False
        Else
            Return True

        End If
    End Function
    Public Function SaveRejectBudgetData2() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER
        Dim rtn As Boolean

        '// Update Budget Data Header
        '// Set Parameters
        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_HEADER.UserPIC = Me.UserPIC
        clsBG_T_BUDGET_HEADER.RevNo = "1"
        clsBG_T_BUDGET_HEADER.Status = CStr(enumBudgetStatus.Submit)
        clsBG_T_BUDGET_HEADER.UserId = Me.UserId
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

        '// Call Function: Update budget status
        If Me.UserPIC = "0000" Then
            rtn = clsBG_T_BUDGET_HEADER.Update003()
        Else
            rtn = clsBG_T_BUDGET_HEADER.Update002()
        End If
        If rtn = False Then

            Return False
        Else
            Return True

        End If
    End Function

    Public Function SaveRejectBudgetData3() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER
        Dim rtn As Boolean

        '// Update Budget Data Header
        '// Set Parameters
        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_HEADER.UserPIC = Me.UserPIC
        clsBG_T_BUDGET_HEADER.RevNo = Me.RevNo
        clsBG_T_BUDGET_HEADER.Status = CStr(enumBudgetStatus.Approve)
        clsBG_T_BUDGET_HEADER.UserId = Me.UserId
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

        '// Call Function: Update budget status
        ''If Me.UserPIC = "0000" Then
        rtn = clsBG_T_BUDGET_HEADER.Update003()
        ''Else
        ''    rtn = clsBG_T_BUDGET_HEADER.Update002()
        ''End If
        If rtn = False Then

            Return False
        Else
            Return True

        End If
    End Function

    Public Function SaveRejectBudgetData4() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER
        Dim rtn As Boolean

        '// Update Budget Data Header
        '// Set Parameters
        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_HEADER.UserPIC = Me.UserPIC
        clsBG_T_BUDGET_HEADER.RevNo = "1"
        clsBG_T_BUDGET_HEADER.Status = CStr(enumBudgetStatus.NewRecord)
        clsBG_T_BUDGET_HEADER.UserId = Me.UserId
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

        '// Call Function: Update budget status
        If Me.UserPIC = "0000" Then
            rtn = clsBG_T_BUDGET_HEADER.Update003()
        Else
            rtn = clsBG_T_BUDGET_HEADER.Update002()
        End If
        If rtn = False Then

            Return False
        Else
            Return True

        End If
    End Function

    Public Function SaveRejectBudgetData4Tran(ByVal pConn As SqlConnection, _
                               ByVal pTrans As SqlTransaction) As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER
        Dim rtn As Boolean

        '// Update Budget Data Header
        '// Set Parameters
        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_HEADER.UserPIC = Me.UserPIC
        clsBG_T_BUDGET_HEADER.RevNo = "1"
        clsBG_T_BUDGET_HEADER.Status = CStr(enumBudgetStatus.NewRecord)
        clsBG_T_BUDGET_HEADER.UserId = Me.UserId
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

        Dim blnS As Boolean = False
        '// Call Function: Update budget status
        If Me.UserPIC = "0000" Then
            If clsBG_T_BUDGET_HEADER.Update003(pConn, pTrans) = True Then
                blnS = True
            Else
                blnS = False
                Return blnS
            End If
        Else

            If clsBG_T_BUDGET_HEADER.Update002(pConn, pTrans) = True Then
                blnS = True
            Else
                blnS = False
                Return blnS
            End If
        End If

        Return blnS
      
    End Function
    Public Function SaveApproveBudgetData() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER

        '// Update Budget Data Header
        '// Set Parameters
        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_HEADER.UserPIC = Me.UserPIC
        clsBG_T_BUDGET_HEADER.RevNo = "1"
        clsBG_T_BUDGET_HEADER.Status = CStr(enumBudgetStatus.Approve)
        clsBG_T_BUDGET_HEADER.UserId = Me.UserId
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

        '// Call Function: Update budget status
        If clsBG_T_BUDGET_HEADER.Update002() = False Then

            Return False
        Else
            Return True

        End If
    End Function
    Public Function DeleteBudgetDataReInputByOrder(ByVal pConn As SqlConnection, _
                               ByVal pTrans As SqlTransaction) As Boolean
        Dim clsBG_T_BUDGET_DATA_REINPUT As New BG_T_BUDGET_DATA_REINPUT

        Dim blnS As Boolean = True

        clsBG_T_BUDGET_DATA_REINPUT.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_DATA_REINPUT.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_DATA_REINPUT.RevNo = Me.RevNo
        clsBG_T_BUDGET_DATA_REINPUT.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_DATA_REINPUT.PersonInChargeNo = Me.UserPIC
        clsBG_T_BUDGET_DATA_REINPUT.ProjectNo = Me.ProjectNo

        If clsBG_T_BUDGET_DATA_REINPUT.PersonInChargeNo = "0000" Then
            If clsBG_T_BUDGET_DATA_REINPUT.Delete002(pConn, pTrans) = True Then
                blnS = True
            Else
                blnS = False
            End If
        Else
            If clsBG_T_BUDGET_DATA_REINPUT.Delete003(pConn, pTrans) = True Then
                blnS = True
            Else
                blnS = False
            End If
        End If

  

        Return blnS

    End Function

    Public Function SaveApproveBudgetDataReInputByOrder(ByVal pConn As SqlConnection, _
                               ByVal pTrans As SqlTransaction) As Boolean
        Dim clsBG_T_BUDGET_DATA_REINPUT As New BG_T_BUDGET_DATA_REINPUT

        Dim blnS As Boolean = True
        If Not Me.dtSave Is Nothing AndAlso dtSave.Rows.Count > 0 Then

            For i As Integer = 0 To dtSave.Rows.Count - 1
                If blnS = True Then
                    clsBG_T_BUDGET_DATA_REINPUT.BudgetYear = dtSave.Rows(i).Item("BudgetYear").ToString
                    clsBG_T_BUDGET_DATA_REINPUT.PeriodType = dtSave.Rows(i).Item("PeriodType").ToString
                    clsBG_T_BUDGET_DATA_REINPUT.RevNo = dtSave.Rows(i).Item("RevNo").ToString
                    clsBG_T_BUDGET_DATA_REINPUT.ProjectNo = dtSave.Rows(i).Item("ProjectNo").ToString
                    clsBG_T_BUDGET_DATA_REINPUT.BudgetOrderNo = dtSave.Rows(i).Item("BudgetOrder").ToString
                    clsBG_T_BUDGET_DATA_REINPUT.ItemStatus = dtSave.Rows(i).Item("Status").ToString
                    'clsBG_T_BUDGET_DATA_REINPUT.UserId = p_strUserId
                    If clsBG_T_BUDGET_DATA_REINPUT.Delete001(pConn, pTrans) = True Then
                        'Return True
                        blnS = True
                    Else
                        'Return False
                        blnS = False
                        Exit For
                    End If
                End If

            Next

            Return blnS
        End If
    End Function

    Public Function SaveAdjustBudgetData() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER

        '// Update Budget Data Header
        '// Set Parameters
        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_HEADER.RevNo = Me.RevNo
        clsBG_T_BUDGET_HEADER.Status = CStr(enumBudgetStatus.Adjust)
        clsBG_T_BUDGET_HEADER.UserId = Me.UserId
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

        '// Call Function: Update budget status
        If clsBG_T_BUDGET_HEADER.Update003() = False Then

            Return False
        Else
            Return True

        End If
    End Function

    Public Function SaveAuth1BudgetData() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER

        '// Update Budget Data Header
        '// Set Parameters
        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_HEADER.RevNo = Me.RevNo
        clsBG_T_BUDGET_HEADER.Status = CStr(enumBudgetStatus.Authorize1)
        clsBG_T_BUDGET_HEADER.UserId = Me.UserId
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

        '// Call Function: Update budget status
        If clsBG_T_BUDGET_HEADER.Update003() = False Then

            Return False
        Else
            Return True

        End If
    End Function

    Public Function SaveAuth2BudgetData() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER

        '// Update Budget Data Header
        '// Set Parameters
        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_HEADER.RevNo = Me.RevNo
        clsBG_T_BUDGET_HEADER.Status = CStr(enumBudgetStatus.Authorize2)
        clsBG_T_BUDGET_HEADER.UserId = Me.UserId
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

        '// Call Function: Update budget status
        If clsBG_T_BUDGET_HEADER.Update003() = False Then

            Return False
        Else
            Return True

        End If
    End Function

    Public Function SaveBudgetDataReInput(ByVal pConn As SqlConnection, _
                               ByVal pTrans As SqlTransaction) As Boolean
        Dim clsBG_T_BUDGET_DATA_REINPUT As New BG_T_BUDGET_DATA_REINPUT

        Dim blnS As Boolean = True
        If Not Me.dtSave Is Nothing AndAlso dtSave.Rows.Count > 0 Then

            For i As Integer = 0 To dtSave.Rows.Count - 1
                If blnS = True Then
                    clsBG_T_BUDGET_DATA_REINPUT.BudgetYear = dtSave.Rows(i).Item("BudgetYear").ToString
                    clsBG_T_BUDGET_DATA_REINPUT.PeriodType = dtSave.Rows(i).Item("PeriodType").ToString
                    clsBG_T_BUDGET_DATA_REINPUT.RevNo = dtSave.Rows(i).Item("RevNo").ToString
                    clsBG_T_BUDGET_DATA_REINPUT.ProjectNo = dtSave.Rows(i).Item("ProjectNo").ToString
                    clsBG_T_BUDGET_DATA_REINPUT.BudgetOrderNo = dtSave.Rows(i).Item("BudgetOrder").ToString
                    clsBG_T_BUDGET_DATA_REINPUT.ItemStatus = dtSave.Rows(i).Item("Status").ToString
                    clsBG_T_BUDGET_DATA_REINPUT.UserId = p_strUserId
                    If clsBG_T_BUDGET_DATA_REINPUT.Insert001(pConn, pTrans) = True Then
                        'Return True
                        blnS = True
                    Else
                        'Return False
                        blnS = False
                        Exit For
                    End If
                End If
              
            Next

            Return blnS
        End If
    End Function

    Public Function GetRevNoList() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER

        '// Set Parameters
        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

        If clsBG_T_BUDGET_HEADER.Select009 = True Then
            Me.RevNoList = clsBG_T_BUDGET_HEADER.dtResult

            Return True
        Else
            Me.RevNoList = New DataTable

            Return False
        End If
    End Function

    Public Function IsSubmitUser() As Boolean
        Dim clsBG_M_USER As New BG_M_USER

        '// Set Parameter
        clsBG_M_USER.UserId = p_strUserId

        '// Call Function: Select User Permission
        If clsBG_M_USER.Select001() = False OrElse _
        clsBG_M_USER.dtResult.Rows.Count = 0 OrElse _
        CStr(clsBG_M_USER.dtResult.Rows(0).Item("SUBMIT")) <> "Y" Then

            Return False
        Else
            Return True
        End If
    End Function

    Public Function GetPersonInChargeList() As Boolean

        If Me.OperationCd = enumOperationCd.InputBudget Then  '// Input Budget
            '// Get all related PIC
            Dim clsBG_M_BUDGET_ORDER As New BG_M_BUDGET_ORDER

            clsBG_M_BUDGET_ORDER.BudgetYear = Me.BudgetYear
            clsBG_M_BUDGET_ORDER.PeriodType = Me.PeriodType
            clsBG_M_BUDGET_ORDER.BudgetType = Me.BudgetType
            clsBG_M_BUDGET_ORDER.RevNo = "1"
            clsBG_M_BUDGET_ORDER.Status = CStr(enumBudgetStatus.NewRecord)
            clsBG_M_BUDGET_ORDER.ProjectNo = Me.ProjectNo

            If clsBG_M_BUDGET_ORDER.Select009() Then
                Me.PicList = clsBG_M_BUDGET_ORDER.dtResult
            Else
                Me.PicList = New DataTable
            End If

        Else
            '// Get PIC of exists data
            Dim clsBG_M_PERSON_IN_CHARGE As New BG_M_PERSON_IN_CHARGE

            clsBG_M_PERSON_IN_CHARGE.BudgetYear = Me.BudgetYear
            clsBG_M_PERSON_IN_CHARGE.PeriodType = Me.PeriodType
            clsBG_M_PERSON_IN_CHARGE.BudgetType = Me.BudgetType
            clsBG_M_PERSON_IN_CHARGE.RevNo = Me.RevNo
            clsBG_M_PERSON_IN_CHARGE.ProjectNo = Me.ProjectNo

            If Me.OperationCd = enumOperationCd.ApproveBudget Then '// Approve Budget

                clsBG_M_PERSON_IN_CHARGE.Status = CStr(enumBudgetStatus.Submit)
                clsBG_M_PERSON_IN_CHARGE.UserPIC = Me.UserPIC

                If Me.UserPIC = "0000" Then
                    If clsBG_M_PERSON_IN_CHARGE.Select006_1 Then
                        Me.PicList = clsBG_M_PERSON_IN_CHARGE.DtResult
                    Else
                        Me.PicList = New DataTable
                    End If
                Else
                    If clsBG_M_PERSON_IN_CHARGE.Select007_1 Then
                        Me.PicList = clsBG_M_PERSON_IN_CHARGE.DtResult
                    Else
                        Me.PicList = New DataTable
                    End If
                End If

            ElseIf Me.OperationCd = enumOperationCd.AdjustBudget Or _
                Me.myOperationCd = enumOperationCd.AdjustBudgetDirectInput Then '// Adjust Budget

                clsBG_M_PERSON_IN_CHARGE.Status = CStr(enumBudgetStatus.Approve)

                ''If clsBG_M_PERSON_IN_CHARGE.Select006 Then
                If clsBG_M_PERSON_IN_CHARGE.Select011 Then
                    Me.PicList = clsBG_M_PERSON_IN_CHARGE.DtResult
                Else
                    Me.PicList = New DataTable
                End If

            ElseIf Me.OperationCd = enumOperationCd.Authorize1 Then '// Authorize1

                clsBG_M_PERSON_IN_CHARGE.Status = CStr(enumBudgetStatus.Adjust)

                If clsBG_M_PERSON_IN_CHARGE.Select006 Then
                    Me.PicList = clsBG_M_PERSON_IN_CHARGE.DtResult
                Else
                    Me.PicList = New DataTable
                End If

            ElseIf Me.OperationCd = enumOperationCd.Authorize2 Then '// Authorize2

                clsBG_M_PERSON_IN_CHARGE.Status = CStr(enumBudgetStatus.Authorize1)

                If clsBG_M_PERSON_IN_CHARGE.Select006 Then
                    Me.PicList = clsBG_M_PERSON_IN_CHARGE.DtResult
                Else
                    Me.PicList = New DataTable
                End If

            Else
                If clsBG_M_PERSON_IN_CHARGE.Select005 Then
                    Me.PicList = clsBG_M_PERSON_IN_CHARGE.DtResult
                Else
                    Me.PicList = New DataTable
                End If
            End If
        End If

        Return True

    End Function

    Public Function GetChildPicList() As Boolean
        Dim clsBG_M_CHILD_PIC As New BG_M_CHILD_PIC

        clsBG_M_CHILD_PIC.ParentNo = Me.UserPIC

        If clsBG_M_CHILD_PIC.Select003 = True And clsBG_M_CHILD_PIC.DtResult.Rows.Count > 0 Then
            Me.ChildPicList = clsBG_M_CHILD_PIC.DtResult

            Return True
        Else
            Return False

        End If
    End Function

    Public Function GetUploadData() As Boolean
        Dim clsBG_T_UPLOAD_DATA As New BG_T_UPLOAD_DATA

        clsBG_T_UPLOAD_DATA.BudgetYear = Me.BudgetYear
        clsBG_T_UPLOAD_DATA.PeriodType = Me.PeriodType
        clsBG_T_UPLOAD_DATA.ProjectNo = Me.ProjectNo

        If clsBG_T_UPLOAD_DATA.Select002 = True Then
            Me.UpDataList = clsBG_T_UPLOAD_DATA.dtResult

            Return True
        Else
            Return False

        End If
    End Function

    Public Function GetConfig() As Boolean
        Dim clsBG_M_SETTINGS As New BG_M_SETTINGS

        If clsBG_M_SETTINGS.Select001 = True Then
            Me.MTPHighlightValue = clsBG_M_SETTINGS.HighLightMTP

            Return True
        Else
            Return False

        End If
    End Function

    Public Function GetMaxRevNo() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER

        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

        If clsBG_T_BUDGET_HEADER.Select006() = True Then
            Me.RevNo = clsBG_T_BUDGET_HEADER.RevNo

            Return True
        Else
            Me.RevNo = "1"

            Return False

        End If
    End Function

    Public Function GetReviseMaxRevNo() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER

        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = CStr(enumPeriodType.ReviseBudget)
        clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_HEADER.ProjectNo = "1"
        clsBG_T_BUDGET_HEADER.UserPIC = Me.UserPIC

        If Me.UserPIC = "0000" Then
            If clsBG_T_BUDGET_HEADER.Select006() = True Then
                Me.ReviseRevNo = clsBG_T_BUDGET_HEADER.RevNo

                Return True
            Else
                Me.ReviseRevNo = "1"

                Return False

            End If
        Else
            If clsBG_T_BUDGET_HEADER.Select012() = True Then
                Me.ReviseRevNo = clsBG_T_BUDGET_HEADER.RevNo

                Return True
            Else
                Me.ReviseRevNo = "1"

                Return False

            End If
        End If

     
    End Function

    Public Function GetPrevMTPMaxRevNo() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER

        clsBG_T_BUDGET_HEADER.BudgetYear = (CInt(Me.BudgetYear) - 1).ToString
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo
        clsBG_T_BUDGET_HEADER.UserPIC = Me.UserPIC

        If Me.UserPIC = "0000" Then
            If clsBG_T_BUDGET_HEADER.Select006() = True Then
                Me.PrevMTPRevNo = clsBG_T_BUDGET_HEADER.RevNo

                Return True
            Else
                Me.PrevMTPRevNo = "1"

                Return False

            End If
        Else
            If clsBG_T_BUDGET_HEADER.Select012() = True Then
                Me.PrevMTPRevNo = clsBG_T_BUDGET_HEADER.RevNo

                Return True
            Else
                Me.PrevMTPRevNo = "1"

                Return False

            End If
        End If

    End Function

    Public Function GetMaxRevStatus() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER

        clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
        clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

        If clsBG_T_BUDGET_HEADER.Select011() = True Then
            Me.Status = clsBG_T_BUDGET_HEADER.Status
            Me.RevNo = clsBG_T_BUDGET_HEADER.RevNo

            Return True
        Else
            Me.Status = "0"
            Me.RevNo = "1"

            Return False

        End If
    End Function

    Public Function DeleteRevision() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER
        Dim clsBG_T_BUDGET_DATA As New BG_T_BUDGET_DATA
        Dim clsBG_T_BUDGET_ADJUST As New BG_T_BUDGET_ADJUST
        Dim clsBG_T_BUDGET_DATA_REINPUT As New BG_T_BUDGET_DATA_REINPUT
        Dim clsBG_T_BUDGET_COMMENT As New BG_T_BUDGET_COMMENT
        Dim conn As SqlConnection
        Dim trans As SqlTransaction

        '// Create Connection
        conn = New SqlConnection(My.Settings.ConnStr)
        conn.Open()

        '// Begin Transaction
        trans = conn.BeginTransaction()

        Try
            clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
            clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
            clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
            clsBG_T_BUDGET_HEADER.RevNo = Me.RevNo
            clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

            clsBG_T_BUDGET_HEADER.Delete001(conn, trans)

            clsBG_T_BUDGET_DATA.BudgetYear = Me.BudgetYear
            clsBG_T_BUDGET_DATA.PeriodType = Me.PeriodType
            clsBG_T_BUDGET_DATA.BudgetType = Me.BudgetType
            clsBG_T_BUDGET_DATA.RevNo = Me.RevNo
            clsBG_T_BUDGET_DATA.ProjectNo = Me.ProjectNo

            clsBG_T_BUDGET_DATA.Delete001(conn, trans)

            '//Delete BUDGET_COMMENT 
            clsBG_T_BUDGET_COMMENT.BudgetYear = Me.BudgetYear
            clsBG_T_BUDGET_COMMENT.PeriodType = Me.PeriodType
            clsBG_T_BUDGET_COMMENT.BudgetType = Me.BudgetType
            clsBG_T_BUDGET_COMMENT.RevNo = Me.RevNo
            clsBG_T_BUDGET_COMMENT.ProjectNo = Me.ProjectNo
            clsBG_T_BUDGET_COMMENT.Delete001(conn, trans)


            clsBG_T_BUDGET_ADJUST.BudgetYear = Me.BudgetYear
            clsBG_T_BUDGET_ADJUST.PeriodType = Me.PeriodType
            clsBG_T_BUDGET_ADJUST.RevNo = Me.RevNo
            clsBG_T_BUDGET_ADJUST.ProjectNo = Me.ProjectNo

            If Me.BudgetType = P_BUDGET_TYPE_EXPENSE Then
                clsBG_T_BUDGET_ADJUST.Delete001(conn, trans)
            Else
                clsBG_T_BUDGET_ADJUST.Delete002(conn, trans)
            End If

            clsBG_T_BUDGET_DATA_REINPUT.BudgetYear = Me.BudgetYear
            clsBG_T_BUDGET_DATA_REINPUT.PeriodType = Me.PeriodType
            clsBG_T_BUDGET_DATA_REINPUT.BudgetType = Me.BudgetType
            clsBG_T_BUDGET_DATA_REINPUT.RevNo = Me.RevNo
            clsBG_T_BUDGET_DATA_REINPUT.ProjectNo = Me.ProjectNo

            clsBG_T_BUDGET_DATA_REINPUT.Delete002(conn, trans)


            '// Commit Transaction
            trans.Commit()

            '// Close Connection
            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception

            '// Rollback Transaction
            trans.Rollback()

            '// Close Connection
            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            MessageBox.Show("Error: " & ex.Message, My.Settings.ProgramTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)

            Return False
        End Try
    End Function

    Public Function DeleteBudgetData() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER
        Dim clsBG_T_BUDGET_DATA As New BG_T_BUDGET_DATA
        Dim conn As SqlConnection
        Dim trans As SqlTransaction

        '// Create Connection
        conn = New SqlConnection(My.Settings.ConnStr)
        conn.Open()

        '// Begin Transaction
        trans = conn.BeginTransaction()

        Try
            clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
            clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
            clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
            clsBG_T_BUDGET_HEADER.UserPIC = Me.UserPIC
            clsBG_T_BUDGET_HEADER.RevNo = Me.RevNo
            clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo

            clsBG_T_BUDGET_HEADER.Delete002(conn, trans)

            clsBG_T_BUDGET_DATA.BudgetYear = Me.BudgetYear
            clsBG_T_BUDGET_DATA.PeriodType = Me.PeriodType
            clsBG_T_BUDGET_DATA.BudgetType = Me.BudgetType
            clsBG_T_BUDGET_DATA.UserPIC = Me.UserPIC
            clsBG_T_BUDGET_DATA.RevNo = Me.RevNo
            clsBG_T_BUDGET_DATA.ProjectNo = Me.ProjectNo

            clsBG_T_BUDGET_DATA.Delete002(conn, trans)

            '// Commit Transaction
            trans.Commit()

            '// Close Connection
            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception

            '// Rollback Transaction
            trans.Rollback()

            '// Close Connection
            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            MessageBox.Show("Error: " & ex.Message, My.Settings.ProgramTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)

            Return False
        End Try
    End Function

    Public Function UpRevision() As Boolean
        Dim clsBG_T_BUDGET_HEADER As New BG_T_BUDGET_HEADER
        Dim clsBG_T_BUDGET_DATA As New BG_T_BUDGET_DATA
        Dim clsBG_T_BUDGET_COMMENT As New BG_T_BUDGET_COMMENT
        Dim conn As SqlConnection
        Dim trans As SqlTransaction
        Dim dtRawDat As DataTable = Nothing

        '// Get rev no.
        Me.RevNo2 = Me.RevNo
        GetMaxRevNo()
        Me.RevNo = CStr(CInt(Me.RevNo) + 1)

        '// Create Connection
        conn = New SqlConnection(My.Settings.ConnStr)
        conn.Open()

        '// Begin Transaction
        trans = conn.BeginTransaction()

        '// Create budget header & budget data & budget adjust
        If CreateBudgetData3(conn, trans) = False Then
            Throw New Exception("Can not insert Budget header")
        End If

        Try
            '// Update Budget Data Header
            For i = 1 To Me.PicList.Rows.Count - 1 '// skip "0000"
                '// Set Parameters
                clsBG_T_BUDGET_HEADER.BudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_HEADER.PeriodType = Me.PeriodType
                clsBG_T_BUDGET_HEADER.BudgetType = Me.BudgetType
                clsBG_T_BUDGET_HEADER.ProjectNo = Me.ProjectNo
                clsBG_T_BUDGET_HEADER.UserPIC = CStr(PicList.Rows(i).Item("PERSON_IN_CHARGE_NO"))
                clsBG_T_BUDGET_HEADER.RevNo = Me.RevNo
                clsBG_T_BUDGET_HEADER.Status = CStr(enumBudgetStatus.Approve)
                clsBG_T_BUDGET_HEADER.UserId = Me.UserId
                clsBG_T_BUDGET_HEADER.RevNo2 = Me.RevNo2

                '// Call Function
                If clsBG_T_BUDGET_HEADER.Update004(conn, trans) = False Then
                    Throw New Exception("Can not update Budget header")
                End If
            Next

            '// Update Budget Data Detail
            For Each dr As DataRow In Me.BudgetList.Rows
                '// Set Parameters
                clsBG_T_BUDGET_DATA.BudgetYear = Me.BudgetYear
                clsBG_T_BUDGET_DATA.PeriodType = Me.PeriodType
                clsBG_T_BUDGET_DATA.ProjectNo = Me.ProjectNo
                clsBG_T_BUDGET_DATA.BudgetOrderNo = CStr(dr("OrderNo"))
                clsBG_T_BUDGET_DATA.RevNo = Me.RevNo
                clsBG_T_BUDGET_DATA.BudgetData = dr ''dtRawDat.Rows(0) 'dr
                clsBG_T_BUDGET_DATA.UserId = Me.UserId

                '// Call Function
                If clsBG_T_BUDGET_DATA.Update001(conn, trans) = False Then
                    Throw New Exception("Can not update Budget data")
                End If

            Next

            '// Commit Transaction
            trans.Commit()

            '// Close Connection
            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            Return True

        Catch ex As Exception

            '// Rollback Transaction
            trans.Rollback()

            '// Close Connection
            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If

            MessageBox.Show("Error: " & ex.Message, My.Settings.ProgramTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)

            Return False
        End Try
    End Function

    Public Function GetWKH() As Boolean
        Dim clsBG_T_BUDGET_ADJUST As New BG_T_BUDGET_ADJUST

        clsBG_T_BUDGET_ADJUST.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_ADJUST.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_ADJUST.RevNo = Me.RevNo
        clsBG_T_BUDGET_ADJUST.ProjectNo = Me.ProjectNo

        If clsBG_T_BUDGET_ADJUST.Select005() = True And clsBG_T_BUDGET_ADJUST.dtResult.Rows.Count > 0 Then
            Me.WKH1 = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!WKH1, 0))
            Me.WKH2 = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!WKH2, 0))

            Me.WKRRT1 = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!WKRRT1, 0))
            Me.WKRRT2 = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!WKRRT2, 0))
            Me.WKRRT3 = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!WKRRT3, 0))
            Me.WKRRT4 = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!WKRRT4, 0))
            Me.WKRRT5 = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!WKRRT5, 0))

            Me.MTPWB = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!MTPWB, 0))
            Return True
        Else
            Me.WKH1 = ""
            Me.WKH2 = ""

            Me.WKRRT1 = ""
            Me.WKRRT2 = ""
            Me.WKRRT3 = ""
            Me.WKRRT4 = ""
            Me.WKRRT5 = ""

            Me.MTPWB = ""
            Return False
        End If
    End Function

    Public Function GetMTP_SUM() As Boolean
        Dim clsBG_T_BUDGET_ADJUST As New BG_T_BUDGET_ADJUST

        clsBG_T_BUDGET_ADJUST.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_ADJUST.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_ADJUST.RevNo = Me.RevNo
        clsBG_T_BUDGET_ADJUST.ProjectNo = Me.ProjectNo

        If clsBG_T_BUDGET_ADJUST.Select006() = True And clsBG_T_BUDGET_ADJUST.dtResult.Rows.Count > 0 Then
            Me.MTP_SUM1 = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!MTP_SUM1, 0))
            Me.MTP_SUM2 = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!MTP_SUM2, 0))
            Me.MTP_SUM3 = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!MTP_SUM3, 0))
            Me.MTP_SUM4 = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!MTP_SUM4, 0))
            Me.MTP_SUM5 = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!MTP_SUM5, 0))

            Return True
        Else
            Me.MTP_SUM1 = ""
            Me.MTP_SUM2 = ""
            Me.MTP_SUM3 = ""
            Me.MTP_SUM4 = ""
            Me.MTP_SUM5 = ""

            Return False
        End If
    End Function

    Public Function GetMTPInvestment() As Boolean
        Dim clsBG_T_BUDGET_ADJUST As New BG_T_BUDGET_ADJUST

        clsBG_T_BUDGET_ADJUST.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_ADJUST.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_ADJUST.RevNo = Me.RevNo
        clsBG_T_BUDGET_ADJUST.ProjectNo = Me.ProjectNo

        If clsBG_T_BUDGET_ADJUST.Select008() = True And clsBG_T_BUDGET_ADJUST.dtResult.Rows.Count > 0 Then
            Me.MTP_SUM1 = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!MTP_SUM1, 0))
            Me.MTP_SUM2 = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!MTP_SUM2, 0))
            Me.MTP_SUM3 = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!MTP_SUM3, 0))
            Me.MTP_SUM4 = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!MTP_SUM4, 0))
            Me.MTP_SUM5 = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!MTP_SUM5, 0))

            Me.MTP_PY_SUM1 = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!MTP_PY_SUM1, 0))
            Me.MTP_PY_SUM2 = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!MTP_PY_SUM2, 0))
            Me.MTP_PY_SUM3 = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!MTP_PY_SUM3, 0))
            Me.MTP_PY_SUM4 = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!MTP_PY_SUM4, 0))
            Me.MTP_PY_SUM5 = CStr(Nz(clsBG_T_BUDGET_ADJUST.dtResult.Rows(0)!MTP_PY_SUM5, 0))

            Return True
        Else
            Me.MTP_SUM1 = ""
            Me.MTP_SUM2 = ""
            Me.MTP_SUM3 = ""
            Me.MTP_SUM4 = ""
            Me.MTP_SUM5 = ""

            Me.MTP_PY_SUM1 = ""
            Me.MTP_PY_SUM2 = ""
            Me.MTP_PY_SUM3 = ""
            Me.MTP_PY_SUM4 = ""
            Me.MTP_PY_SUM5 = ""

            Return False
        End If
    End Function

    Public Function SaveWKH() As Boolean
        Dim clsBG_T_BUDGET_ADJUST As New BG_T_BUDGET_ADJUST

        clsBG_T_BUDGET_ADJUST.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_ADJUST.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_ADJUST.RevNo = Me.RevNo
        clsBG_T_BUDGET_ADJUST.WKH1 = Me.WKH1
        clsBG_T_BUDGET_ADJUST.WKH2 = Me.WKH2
        clsBG_T_BUDGET_ADJUST.WKRRT1 = Me.WKRRT1
        clsBG_T_BUDGET_ADJUST.WKRRT2 = Me.WKRRT2
        clsBG_T_BUDGET_ADJUST.WKRRT3 = Me.WKRRT3
        clsBG_T_BUDGET_ADJUST.WKRRT4 = Me.WKRRT4
        clsBG_T_BUDGET_ADJUST.WKRRT5 = Me.WKRRT5
        clsBG_T_BUDGET_ADJUST.MTPWB = Me.MTPWB
        clsBG_T_BUDGET_ADJUST.UpdateUserID = Me.UserId
        clsBG_T_BUDGET_ADJUST.ProjectNo = Me.ProjectNo

        If clsBG_T_BUDGET_ADJUST.Update002() = True Then

            Return True
        Else

            Return False
        End If
    End Function

    Public Function SaveMTP_SUM() As Boolean
        Dim clsBG_T_BUDGET_ADJUST As New BG_T_BUDGET_ADJUST

        clsBG_T_BUDGET_ADJUST.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_ADJUST.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_ADJUST.RevNo = Me.RevNo
        clsBG_T_BUDGET_ADJUST.MTP_SUM1 = Me.MTP_SUM1
        clsBG_T_BUDGET_ADJUST.MTP_SUM2 = Me.MTP_SUM2
        clsBG_T_BUDGET_ADJUST.MTP_SUM3 = Me.MTP_SUM3
        clsBG_T_BUDGET_ADJUST.MTP_SUM4 = Me.MTP_SUM4
        clsBG_T_BUDGET_ADJUST.MTP_SUM5 = Me.MTP_SUM5
        clsBG_T_BUDGET_ADJUST.UpdateUserID = Me.UserId
        clsBG_T_BUDGET_ADJUST.ProjectNo = Me.ProjectNo

        If clsBG_T_BUDGET_ADJUST.Select006() = True AndAlso clsBG_T_BUDGET_ADJUST.dtResult.Rows.Count > 0 Then
            If clsBG_T_BUDGET_ADJUST.Update003() = True Then

                Return True
            Else
                Return False
            End If

        Else
            If clsBG_T_BUDGET_ADJUST.Insert002() = True Then

                Return True
            Else
                Return False
            End If
        End If
    End Function

    Public Function SaveMTPInvestment() As Boolean
        Dim clsBG_T_BUDGET_ADJUST As New BG_T_BUDGET_ADJUST

        clsBG_T_BUDGET_ADJUST.BudgetYear = Me.BudgetYear
        clsBG_T_BUDGET_ADJUST.PeriodType = Me.PeriodType
        clsBG_T_BUDGET_ADJUST.RevNo = Me.RevNo

        clsBG_T_BUDGET_ADJUST.MTP_SUM1 = Me.MTP_SUM1
        clsBG_T_BUDGET_ADJUST.MTP_SUM2 = Me.MTP_SUM2
        clsBG_T_BUDGET_ADJUST.MTP_SUM3 = Me.MTP_SUM3
        clsBG_T_BUDGET_ADJUST.MTP_SUM4 = Me.MTP_SUM4
        clsBG_T_BUDGET_ADJUST.MTP_SUM5 = Me.MTP_SUM5

        clsBG_T_BUDGET_ADJUST.MTP_PY_SUM1 = Me.MTP_PY_SUM1
        clsBG_T_BUDGET_ADJUST.MTP_PY_SUM2 = Me.MTP_PY_SUM2
        clsBG_T_BUDGET_ADJUST.MTP_PY_SUM3 = Me.MTP_PY_SUM3
        clsBG_T_BUDGET_ADJUST.MTP_PY_SUM4 = Me.MTP_PY_SUM4
        clsBG_T_BUDGET_ADJUST.MTP_PY_SUM5 = Me.MTP_PY_SUM5

        clsBG_T_BUDGET_ADJUST.UpdateUserID = Me.UserId
        clsBG_T_BUDGET_ADJUST.ProjectNo = Me.ProjectNo

        If clsBG_T_BUDGET_ADJUST.Select008() = True AndAlso clsBG_T_BUDGET_ADJUST.dtResult.Rows.Count > 0 Then
            If clsBG_T_BUDGET_ADJUST.Update004() = True Then

                Return True
            Else
                Return False
            End If

        Else
            If clsBG_T_BUDGET_ADJUST.Insert003() = True Then

                Return True
            Else
                Return False
            End If
        End If
    End Function

    Public Function GetTransferCost() As Boolean
        Dim clsBG_M_TRANSFER_MASTER As New BG_M_TRANSFER_MASTER

        clsBG_M_TRANSFER_MASTER.BudgetYear = Me.BudgetYear
        clsBG_M_TRANSFER_MASTER.PeriodType = Me.PeriodType
        clsBG_M_TRANSFER_MASTER.ProjectNo = Me.ProjectNo

        If clsBG_M_TRANSFER_MASTER.Select003() = True Then
            Me.TransferList = clsBG_M_TRANSFER_MASTER.DTResult

            Return True
        Else
            Me.TransferList = New DataTable

            Return False

        End If
    End Function

    Public Function GetTransferCost2(ByVal BudgetYear As String, ByVal PeriodType As String, ByVal RevNo As String, ByVal ProjectNo As String) As Boolean
        Dim clsBG_M_TRANSFER_MASTER As New BG_M_TRANSFER_MASTER

        clsBG_M_TRANSFER_MASTER.BudgetYear = BudgetYear
        clsBG_M_TRANSFER_MASTER.PeriodType = PeriodType
        clsBG_M_TRANSFER_MASTER.RevNo = RevNo
        clsBG_M_TRANSFER_MASTER.ProjectNo = ProjectNo

        If clsBG_M_TRANSFER_MASTER.Select004() = True Then
            Me.TransferList = clsBG_M_TRANSFER_MASTER.DTResult

            Return True
        Else
            Me.TransferList = New DataTable

            Return False
        End If
    End Function

    Public Function SendAutoMail() As Boolean
        Dim strFromAddr As String = String.Empty
        Dim strToAddr As String = String.Empty
        Dim strSubject As String = String.Empty
        Dim strMessage As String = String.Empty
        Dim strPeriodType As String = String.Empty
        Dim strBudgetType As String = String.Empty
        Dim strPic As String = String.Empty
        Dim strStatus As String = String.Empty
        Dim dtRecipients As DataTable = Nothing

        '// Get From Address
        strFromAddr = p_strAutoMailFromAddr

        '// Create recipient list
        If Me.OperationCd = enumOperationCd.SubmitBudget Then
            dtRecipients = GetApproverMail(Me.UserPIC)

        ElseIf Me.OperationCd = enumOperationCd.RejectBudget1 Then
            dtRecipients = GetSubmitterMail(Me.UserPIC)

        ElseIf Me.OperationCd = enumOperationCd.ApproveBudget Then
            dtRecipients = GetAdminMail()

        ElseIf Me.OperationCd = enumOperationCd.RejectBudget2 Then
            dtRecipients = GetApproverMail(Me.UserPIC)

        ElseIf Me.OperationCd = enumOperationCd.AdjustBudget Then
            dtRecipients = GetAuth1Mail(Me.UserPIC)

        ElseIf Me.OperationCd = enumOperationCd.Authorize1 Then
            dtRecipients = GetAuth2Mail(Me.UserPIC)

        ElseIf Me.OperationCd = enumOperationCd.Authorize2 Then
            dtRecipients = GetAdminMail()

        ElseIf Me.OperationCd = enumOperationCd.RejectBudget3 Then
            dtRecipients = GetAdminMail()
        End If

        strToAddr = ""
        If dtRecipients IsNot Nothing Then
            For Each dr As DataRow In dtRecipients.Rows
                If CStr(Nz(dr("EMAIL"), "")) <> "" Then
                    strToAddr &= CStr(dr("EMAIL")) & "; "
                End If
            Next
        End If
        If strToAddr.Trim.Length = 0 Then
            '// exit function if on recipient
            Return True
        End If

        '// Create Budget Information
        If Me.PeriodType = CStr(enumPeriodType.OriginalBudget) Then
            strPeriodType = "Original Budget"

        ElseIf Me.PeriodType = CStr(enumPeriodType.EstimateBudget) Then
            strPeriodType = "Estimate Budget"

        ElseIf Me.PeriodType = CStr(enumPeriodType.ReviseBudget) Then
            strPeriodType = "Revise Budget"

        ElseIf Me.PeriodType = CStr(enumPeriodType.MTPBudget) Then
            strPeriodType = "MTP Budget"

        End If

        If Me.BudgetType = P_BUDGET_TYPE_EXPENSE Then
            strBudgetType = "(Expense)"

        ElseIf Me.BudgetType = P_BUDGET_TYPE_ASSET Then
            strBudgetType = "(Investment)"
        End If

        '// Create Subject & Status
        If Me.OperationCd = enumOperationCd.SubmitBudget Then
            strSubject = "[Budget System] Submit " & Me.BudgetYear & " " & strPeriodType & " " & Me.ProjectNo
            strStatus = " was been submitted."

        ElseIf Me.OperationCd = enumOperationCd.RejectBudget1 Then
            strSubject = "[Budget System] Reject " & Me.BudgetYear & " " & strPeriodType & " " & Me.ProjectNo
            strStatus = " was been rejected."

        ElseIf Me.OperationCd = enumOperationCd.ApproveBudget Then
            strSubject = "[Budget System] Approve " & Me.BudgetYear & " " & strPeriodType & " " & Me.ProjectNo
            strStatus = " was been approved."

        ElseIf Me.OperationCd = enumOperationCd.RejectBudget2 Then
            strSubject = "[Budget System] Reject " & Me.BudgetYear & " " & strPeriodType & " " & Me.ProjectNo
            strStatus = " was been rejected."

        ElseIf Me.OperationCd = enumOperationCd.AdjustBudget Then
            strSubject = "[Budget System] Adjust " & Me.BudgetYear & " " & strPeriodType & " " & Me.ProjectNo
            strStatus = " was been adjusted."

        ElseIf Me.OperationCd = enumOperationCd.Authorize1 Then
            strSubject = "[Budget System] Authorize (AUTH1) " & Me.BudgetYear & " " & strPeriodType & " " & Me.ProjectNo
            strStatus = " was been authorized (AUTH1)."

        ElseIf Me.OperationCd = enumOperationCd.Authorize2 Then
            strSubject = "[Budget System] Authorize (AUTH2) " & Me.BudgetYear & " " & strPeriodType & " " & Me.ProjectNo
            strStatus = " was been authorized (AUTH2)."
        End If

        '// Create send message
        strMessage = Me.BudgetYear & " " & strPeriodType & " " & Me.ProjectNo & " " & strBudgetType & strStatus & vbNewLine & vbNewLine
        strMessage &= "By" & vbNewLine
        strMessage &= "User Name: " & p_strUserName & vbNewLine
        strMessage &= "PIC: " & Me.UserPIC & vbNewLine & vbNewLine
        strMessage &= "*This is auto mail from " & My.Settings.ProgramTitle & "." & vbNewLine

        '// Send auto mail
        If strMessage.Trim.Length > 0 Then
            BGSendMail.SendMessage(strToAddr, strSubject, strMessage)

            Return True
        Else
            Return False
        End If
    End Function

    Private Function GetAdminMail() As DataTable
        Dim clsBG_M_USER As New BG_M_USER
        Dim dt As DataTable = Nothing

        clsBG_M_USER.UserPIC = "0000"

        If clsBG_M_USER.Select006 = True Then
            dt = clsBG_M_USER.dtResult
        End If

        Return dt
    End Function

    Private Function GetAuth1Mail(ByVal UserPic As String) As DataTable
        Dim clsBG_M_USER As New BG_M_USER
        Dim dt As DataTable = Nothing

        clsBG_M_USER.UserPIC = "BTMT10"

        If clsBG_M_USER.Select006 = True Then
            dt = clsBG_M_USER.dtResult
        End If

        Return dt
    End Function

    Private Function GetAuth2Mail(ByVal UserPic As String) As DataTable
        Dim clsBG_M_USER As New BG_M_USER
        Dim dt As DataTable = Nothing

        clsBG_M_USER.UserPIC = "BTMT3"

        If clsBG_M_USER.Select006 = True Then
            dt = clsBG_M_USER.dtResult
        End If

        Return dt
    End Function

    Private Function GetApproverMail(ByVal UserPic As String) As DataTable
        Dim clsBG_M_USER As New BG_M_USER
        Dim dt As DataTable = Nothing

        If UserPic = "0000" Then
            If clsBG_M_USER.Select010 = True Then
                dt = clsBG_M_USER.dtResult
            End If
        Else
            clsBG_M_USER.UserPIC = UserPic

            If clsBG_M_USER.Select008 = True Then
                dt = clsBG_M_USER.dtResult
            End If
        End If

        Return dt
    End Function

    Private Function GetSubmitterMail(ByVal UserPic As String) As DataTable
        Dim clsBG_M_USER As New BG_M_USER
        Dim dt As DataTable = Nothing

        If UserPic = "BTMT3" Or UserPic = "BTMT10" Then
            clsBG_M_USER.UserPIC = UserPic
        Else
            clsBG_M_USER.UserPIC = Mid(UserPic, 1, 3)
        End If

        If clsBG_M_USER.Select009 = True Then
            dt = clsBG_M_USER.dtResult
        End If

        Return dt
    End Function

    Private Function UpdateTransferCost(ByVal IncreaseFlag As Boolean, _
                                          ByVal BudgetYear As String, ByVal PeriodType As String, _
                                          ByVal RevNo As String, ByVal UserId As String, _
                                          ByVal OrderNo As String, ByVal arrDat As Double(), _
                                          ByVal ProjectNo As String) As Boolean
        Dim clsBG_T_BUDGET_DATA As New BG_T_BUDGET_DATA

        clsBG_T_BUDGET_DATA.BudgetYear = BudgetYear
        clsBG_T_BUDGET_DATA.PeriodType = PeriodType
        clsBG_T_BUDGET_DATA.BudgetOrderNo = OrderNo
        clsBG_T_BUDGET_DATA.RevNo = RevNo
        clsBG_T_BUDGET_DATA.BudgetData2 = arrDat
        clsBG_T_BUDGET_DATA.UserId = UserId
        clsBG_T_BUDGET_DATA.ProjectNo = ProjectNo

        If clsBG_T_BUDGET_DATA.Update002() = True Then

            Return True
        Else

            Return False
        End If
    End Function

    Private Function GetRawdata(ByVal BudgetYear As String, ByVal PeriodType As String, _
                                ByVal RevNo As String, ByVal OrderNo As String, ByVal ProjectNo As String) As DataTable
        Dim clsBG_T_BUDGET_DATA As New BG_T_BUDGET_DATA
        Dim dt As DataTable = Nothing

        clsBG_T_BUDGET_DATA.BudgetYear = BudgetYear
        clsBG_T_BUDGET_DATA.PeriodType = PeriodType
        clsBG_T_BUDGET_DATA.BudgetOrderNo = OrderNo
        clsBG_T_BUDGET_DATA.RevNo = RevNo '' "1"
        clsBG_T_BUDGET_DATA.ProjectNo = ProjectNo

        If clsBG_T_BUDGET_DATA.Select019 = True Then
            dt = clsBG_T_BUDGET_DATA.dtResult
        Else
            dt = New DataTable
        End If

        Return dt
    End Function

    ''' <summary>
    ''' AdjustTransferCost()
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Last Updated: 2011/05/30 by S.Watcharapong</remarks>
    Public Function AdjustTransferCost() As Boolean
        Dim dblTransValue(11) As Double
        Dim dtRawDat As DataTable = Nothing
        Dim arrToOrders As New ArrayList
        Dim arrFromOrders As New ArrayList

        Me.UserId = p_strUserId

        If GetMaxRevNo() = False Then
            Return False
        End If

      
        If Me.RevNo = "1" Then
            Return False
        End If
       
        '// Get Transfer Data
        If GetTransferCost2(Me.BudgetYear, Me.PeriodType, Me.RevNo, Me.ProjectNo) = False Then
            Return False
        End If

        '     Dim intTempRev As Integer
        Dim drData() As DataRow
        Dim idx As Integer
        For Each dr As DataRow In Me.TransferList.Rows

            If CInt(dr("CAL_FROM_FLG").ToString) = 0 Then

                drData = TransferList.Select("FROM_ORDER_NO ='" & CStr(dr("FROM_ORDER_NO")) & "'")

                If drData.Length > 0 Then

                    ReDim dblTransValue(11)

                    For idx = 0 To drData.Length - 1
                        If Me.PeriodType = CStr(enumPeriodType.ReviseBudget) Then '// Revise Budget transfer cost on Apr - Dec only.
                            For i = 3 To 11
                                dblTransValue(i) = dblTransValue(i) - _
                                                    (CDbl(Nz(drData(idx).Item("M" & (i + 1)), 0))) / 100 * CDbl(drData(idx).Item("TRANSFER_RATE"))
                            Next
                            '    UpdateTransferCost(False, Me.BudgetYear, Me.PeriodType, Me.RevNo, Me.UserId, CStr(dr("FROM_ORDER_NO")), dblTransValue)

                        ElseIf Me.PeriodType = CStr(enumPeriodType.EstimateBudget) Then '// Estimate Budget transfer cost on Oct - Dec only.
                            For i = 9 To 11

                                dblTransValue(i) = dblTransValue(i) - _
                                                                        (CDbl(Nz(drData(idx).Item("M" & (i + 1)), 0))) / 100 * CDbl(drData(idx).Item("TRANSFER_RATE"))

                            Next

                            System.Diagnostics.Debug.Print("Rev." & RevNo & "Main No." & drData(idx).Item("MAIN_BUDGET_ORDER_NO").ToString & " FromNo." & CStr(drData(idx).Item("FROM_ORDER_NO")) & "M11 : " & dblTransValue(10).ToString & " : " & drData(idx).Item("M11").ToString)

                        ElseIf Me.PeriodType = CStr(enumPeriodType.OriginalBudget) Then '// Original Budget transfer cost on all month.
                            For i = 0 To 11
                                dblTransValue(i) = dblTransValue(i) - _
                                                    (CDbl(Nz(drData(idx).Item("M" & (i + 1)), 0))) / 100 * CDbl(drData(idx).Item("TRANSFER_RATE"))
                            Next

                        ElseIf Me.PeriodType = CStr(enumPeriodType.MTPBudget) Then
                            For i = 0 To 4
                                dblTransValue(i) = dblTransValue(i) - _
                                                    (CDbl(Nz(drData(idx).Item("RRT" & (i + 1)), 0))) / 100 * CDbl(drData(idx).Item("TRANSFER_RATE"))
                            Next
                        End If
                    Next

                    UpdateTransferCost(False, Me.BudgetYear, Me.PeriodType, Me.RevNo, Me.UserId, CStr(dr("FROM_ORDER_NO")), dblTransValue, Me.ProjectNo)
                End If

                ' Update Flag
                For i = 0 To drData.Length - 1

                    drData(i).Item("CAL_FROM_FLG") = 1

                Next
            End If
        Next

        For Each dr As DataRow In Me.TransferList.Rows
            If CInt(dr("CAL_TO_FLG").ToString) = 0 Then
                drData = TransferList.Select("TO_ORDER_NO ='" & CStr(dr("TO_ORDER_NO")) & "'")

                If drData.Length > 0 Then

                    ReDim dblTransValue(11)

                    For idx = 0 To drData.Length - 1

                        If Me.PeriodType = CStr(enumPeriodType.ReviseBudget) Then '// Revise Budget transfer cost on Apr - Dec only.
                            For i = 3 To 11
                                dblTransValue(i) = dblTransValue(i) + _
                                                    (CDbl(Nz(drData(idx).Item("M" & (i + 1)), 0))) / 100 * CDbl(drData(idx).Item("TRANSFER_RATE"))
                            Next

                        ElseIf Me.PeriodType = CStr(enumPeriodType.EstimateBudget) Then '// Estimate Budget transfer cost on Oct - Dec only.
                            For i = 9 To 11

                                dblTransValue(i) = dblTransValue(i) + _
                                                                        (CDbl(Nz(drData(idx).Item("M" & (i + 1)), 0))) / 100 * CDbl(drData(idx).Item("TRANSFER_RATE"))

                            Next

                            System.Diagnostics.Debug.Print("Rev." & RevNo & "Main No." & drData(idx).Item("MAIN_BUDGET_ORDER_NO").ToString & " FromNo." & CStr(drData(idx).Item("TO_ORDER_NO")) & "M11 : " & dblTransValue(10).ToString & " : " & drData(idx).Item("M11").ToString)

                        ElseIf Me.PeriodType = CStr(enumPeriodType.OriginalBudget) Then '// Original Budget transfer cost on all month.
                            For i = 0 To 11
                                dblTransValue(i) = dblTransValue(i) + _
                                                    (CDbl(Nz(drData(idx).Item("M" & (i + 1)), 0))) / 100 * CDbl(drData(idx).Item("TRANSFER_RATE"))
                            Next

                        ElseIf Me.PeriodType = CStr(enumPeriodType.MTPBudget) Then
                            For i = 0 To 4
                                dblTransValue(i) = dblTransValue(i) + _
                                                    (CDbl(Nz(drData(idx).Item("RRT" & (i + 1)), 0))) / 100 * CDbl(drData(idx).Item("TRANSFER_RATE"))
                            Next

                        End If

                    Next

                    UpdateTransferCost(False, Me.BudgetYear, Me.PeriodType, Me.RevNo, Me.UserId, CStr(dr("TO_ORDER_NO")), dblTransValue, Me.ProjectNo)

                End If
                ' Update Flag
                For i = 0 To drData.Length - 1

                    drData(i).Item("CAL_TO_FLG") = 1

                Next
            End If
        Next

        Return True
    End Function

    Public Sub WriteTransLog()
        WriteTransactionLog(CStr(Me.LogOperationCd), Me.BudgetYear, Me.PeriodType, Me.UserPIC, Me.BudgetType, Me.RevNo, Me.ProjectNo)
    End Sub

    Public Function GetBudGetDataReInput() As DataTable
        Dim clsBG_T_BUDGET_DATA_REINPUT As New BG_T_BUDGET_DATA_REINPUT
        Dim dt As DataTable = Nothing

        clsBG_T_BUDGET_DATA_REINPUT.BudgetYear = BudgetYear
        clsBG_T_BUDGET_DATA_REINPUT.PeriodType = PeriodType
        clsBG_T_BUDGET_DATA_REINPUT.BudgetType = BudgetType
        clsBG_T_BUDGET_DATA_REINPUT.UserPIC = UserPIC
        clsBG_T_BUDGET_DATA_REINPUT.RevNo = RevNo
        clsBG_T_BUDGET_DATA_REINPUT.ItemStatus = Status
        clsBG_T_BUDGET_DATA_REINPUT.ProjectNo = ProjectNo

        If clsBG_T_BUDGET_DATA_REINPUT.PersonInChargeNo = "0000" Then
            If clsBG_T_BUDGET_DATA_REINPUT.Select002 = True Then
                dt = clsBG_T_BUDGET_DATA_REINPUT.dtResult
            Else
                dt = New DataTable
            End If
        Else
            If clsBG_T_BUDGET_DATA_REINPUT.Select003 = True Then
                dt = clsBG_T_BUDGET_DATA_REINPUT.dtResult
            Else
                dt = New DataTable
            End If
        End If
       
        Return dt
    End Function

    Public Function GetBudGetDataReInputNoStatus() As DataTable
        Dim clsBG_T_BUDGET_DATA_REINPUT As New BG_T_BUDGET_DATA_REINPUT
        Dim dt As DataTable = Nothing

        clsBG_T_BUDGET_DATA_REINPUT.BudgetYear = BudgetYear
        clsBG_T_BUDGET_DATA_REINPUT.PeriodType = PeriodType
        clsBG_T_BUDGET_DATA_REINPUT.BudgetType = BudgetType
        clsBG_T_BUDGET_DATA_REINPUT.PersonInChargeNo = UserPIC
        clsBG_T_BUDGET_DATA_REINPUT.RevNo = RevNo
        clsBG_T_BUDGET_DATA_REINPUT.ProjectNo = ProjectNo

        If clsBG_T_BUDGET_DATA_REINPUT.PersonInChargeNo = "0000" Then
            If clsBG_T_BUDGET_DATA_REINPUT.Select004 = True Then
                dt = clsBG_T_BUDGET_DATA_REINPUT.dtResult
            Else
                dt = New DataTable
            End If
        Else
            If clsBG_T_BUDGET_DATA_REINPUT.Select005 = True Then
                dt = clsBG_T_BUDGET_DATA_REINPUT.dtResult
            Else
                dt = New DataTable
            End If
        End If

        Return dt

    End Function
#End Region

End Class
