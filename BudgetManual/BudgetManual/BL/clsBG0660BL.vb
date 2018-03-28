Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.SqlClient

Public Class clsBG0660BL

#Region "Variable"
    Private clsBG_M_BUDGET_ORDER As BG_M_BUDGET_ORDER
    Private clsBG_M_TRANSFER_MASTER As BG_M_TRANSFER_MASTER
    Private myBudgetYear As String = String.Empty
    Private myProjectNo As String = String.Empty
    Private myPeriodType As String = String.Empty
    Private myBudgetOrderNo As String = String.Empty
    Private myTransferType As String = String.Empty
    Private myTransferRate As String = String.Empty
    Private myFromOrderNo As String = String.Empty
    Private myToOrderNo As String = String.Empty
    Private myCreateUserID As String = String.Empty
    Private myCreateDate As String = String.Empty
    Private myUpdateUserID As String = String.Empty
    Private myDTResult As DataTable
    Private myAccountNo As String = String.Empty

    Private myBudgetOrderNoFilter As String = String.Empty
    Private myBudgetOrderNameFilter As String = String.Empty
    Private myTransferTypeFilter As String = String.Empty
    Private myFromOrderNoFilter As String = String.Empty
    Private myToOrderNoFilter As String = String.Empty
    Private myAccountNoFilter As String = String.Empty
    Private myAccountNameFilter As String = String.Empty
#End Region

#Region "Property"
    Public Property DTResult() As DataTable
        Get
            Return myDTResult
        End Get
        Set(ByVal value As DataTable)
            myDTResult = value
        End Set
    End Property
    Public Property BudgetYear() As String
        Get
            Return myBudgetYear
        End Get
        Set(ByVal value As String)
            myBudgetYear = value
        End Set
    End Property
    Public Property ProjectNo() As String
        Get
            Return myProjectNo
        End Get
        Set(ByVal value As String)
            myProjectNo = value
        End Set
    End Property
    Public Property PeriodType() As String
        Get
            Return myPeriodType
        End Get
        Set(ByVal value As String)
            myPeriodType = value
        End Set
    End Property
    Public Property BudgetOrderNo() As String
        Get
            Return myBudgetOrderNo
        End Get
        Set(ByVal value As String)
            myBudgetOrderNo = value
        End Set
    End Property
    Public Property TransferType() As String
        Get
            Return myTransferType
        End Get
        Set(ByVal value As String)
            myTransferType = value
        End Set
    End Property
    Public Property TransferRate() As String
        Get
            Return myTransferRate
        End Get
        Set(ByVal value As String)
            myTransferRate = value
        End Set
    End Property
    Public Property FromOrderNo() As String
        Get
            Return myFromOrderNo
        End Get
        Set(ByVal value As String)
            myFromOrderNo = value
        End Set
    End Property
    Public Property ToOrderNo() As String
        Get
            Return myToOrderNo
        End Get
        Set(ByVal value As String)
            myToOrderNo = value
        End Set
    End Property
    Public Property CreateUserID() As String
        Get
            Return myCreateUserID
        End Get
        Set(ByVal value As String)
            myCreateUserID = value
        End Set
    End Property
    Public Property CreateDate() As String
        Get
            Return myCreateDate
        End Get
        Set(ByVal value As String)
            myCreateDate = value
        End Set
    End Property
    Public Property UpdateUserID() As String
        Get
            Return myUpdateUserID
        End Get
        Set(ByVal value As String)
            myUpdateUserID = value
        End Set
    End Property
    Public Property AccountNo() As String
        Get
            Return myAccountNo
        End Get
        Set(ByVal value As String)
            myAccountNo = value
        End Set
    End Property


    Public Property BudgetOrderNoFilter() As String
        Get
            Return myBudgetOrderNoFilter
        End Get
        Set(ByVal value As String)
            myBudgetOrderNoFilter = value
        End Set
    End Property
    Public Property BudgetOrderNameFilter() As String
        Get
            Return myBudgetOrderNameFilter
        End Get
        Set(ByVal value As String)
            myBudgetOrderNameFilter = value
        End Set
    End Property
    Public Property TransferTypeFilter() As String
        Get
            Return myTransferTypeFilter
        End Get
        Set(ByVal value As String)
            myTransferTypeFilter = value
        End Set
    End Property
    Public Property FromOrderNoFilter() As String
        Get
            Return myFromOrderNoFilter
        End Get
        Set(ByVal value As String)
            myFromOrderNoFilter = value
        End Set
    End Property
    Public Property ToOrderNoFilter() As String
        Get
            Return myToOrderNoFilter
        End Get
        Set(ByVal value As String)
            myToOrderNoFilter = value
        End Set
    End Property
    Public Property AccountNoFilter() As String
        Get
            Return myAccountNoFilter
        End Get
        Set(ByVal value As String)
            myAccountNoFilter = value
        End Set
    End Property
    Public Property AccountNameFilter() As String
        Get
            Return myAccountNameFilter
        End Get
        Set(ByVal value As String)
            myAccountNameFilter = value
        End Set
    End Property
#End Region

#Region "Function"

    Private Function dateConvert(ByVal strDate As String) As String
        Dim myDate As Date
        If strDate <> "" Then
            myDate = CDate(strDate)
        Else
            myDate = Date.Now
        End If

        Return myDate.ToString("yyyy-MM-dd HH:mm:ss")
    End Function

    Public Function getBGOrderList() As Boolean
        Dim result As Boolean = False

        clsBG_M_BUDGET_ORDER = New BG_M_BUDGET_ORDER

        If clsBG_M_BUDGET_ORDER.Select007 Then
            myDTResult = clsBG_M_BUDGET_ORDER.dtResult
            result = True
        Else
            myDTResult = New DataTable
        End If

        Return result
    End Function

    Public Function getAccountList() As Boolean
        Dim result As Boolean = False

        Dim clsBG_M_ACCOUNT As New BG_M_ACCOUNT

        If clsBG_M_ACCOUNT.Select001 Then
            myDTResult = clsBG_M_ACCOUNT.DtResult
            result = True
        Else
            myDTResult = New DataTable
        End If

        Return result
    End Function

    Public Function getOrderByAccount() As Boolean
        Dim result As Boolean = False

        Dim clsBG_M_BUDGET_ORDER As New BG_M_BUDGET_ORDER

        clsBG_M_BUDGET_ORDER.Account = Me.AccountNo

        If clsBG_M_BUDGET_ORDER.Select012 Then
            myDTResult = clsBG_M_BUDGET_ORDER.dtResult
            result = True
        Else
            myDTResult = New DataTable
        End If

        Return result
    End Function

    Public Function getDataList() As Boolean
        Dim result As Boolean = False
        clsBG_M_TRANSFER_MASTER = New BG_M_TRANSFER_MASTER

        clsBG_M_TRANSFER_MASTER.BudgetYear = Me.BudgetYear
        clsBG_M_TRANSFER_MASTER.PeriodType = Me.PeriodType
        clsBG_M_TRANSFER_MASTER.ProjectNo = Me.ProjectNo

        clsBG_M_TRANSFER_MASTER.BudgetOrderNoFilter = Me.BudgetOrderNoFilter
        clsBG_M_TRANSFER_MASTER.BudgetOrderNameFilter = Me.BudgetOrderNameFilter
        clsBG_M_TRANSFER_MASTER.AccountNoFilter = Me.AccountNoFilter
        clsBG_M_TRANSFER_MASTER.AccountNameFilter = Me.AccountNameFilter
        clsBG_M_TRANSFER_MASTER.TransferTypeFilter = Me.TransferTypeFilter
        clsBG_M_TRANSFER_MASTER.FromOrderNoFilter = Me.FromOrderNoFilter
        clsBG_M_TRANSFER_MASTER.ToOrderNoFilter = Me.ToOrderNoFilter

        If clsBG_M_TRANSFER_MASTER.select001 Then
            Me.DTResult = clsBG_M_TRANSFER_MASTER.DTResult
            result = True
        Else
            Me.DTResult = New DataTable
        End If

        Return result
    End Function

    Public Function saveIndividualData() As Boolean
        Dim success As Boolean = False
        clsBG_M_TRANSFER_MASTER = New BG_M_TRANSFER_MASTER

        clsBG_M_TRANSFER_MASTER.BudgetOrderNo = Me.BudgetOrderNo
        clsBG_M_TRANSFER_MASTER.BudgetYear = Me.BudgetYear
        clsBG_M_TRANSFER_MASTER.PeriodType = Me.PeriodType
        clsBG_M_TRANSFER_MASTER.TransferType = Me.TransferType
        clsBG_M_TRANSFER_MASTER.TransferRate = Me.TransferRate
        clsBG_M_TRANSFER_MASTER.FromOrderNo = Me.FromOrderNo
        clsBG_M_TRANSFER_MASTER.ToOrderNo = Me.ToOrderNo
        clsBG_M_TRANSFER_MASTER.CreateUserID = Me.CreateUserID
        clsBG_M_TRANSFER_MASTER.CreateDate = "GETDATE()"
        clsBG_M_TRANSFER_MASTER.UpdateUserID = Me.UpdateUserID
        clsBG_M_TRANSFER_MASTER.ProjectNo = Me.ProjectNo

        Dim conn As New SqlConnection
        Dim trans As SqlTransaction
        conn.ConnectionString = My.Settings.ConnStr

        conn.Open()
        trans = conn.BeginTransaction
        Try
            If clsBG_M_TRANSFER_MASTER.select002(conn, trans) Then
                If CInt(clsBG_M_TRANSFER_MASTER.DTResult.Rows(0).Item(0).ToString) = 0 Then
                    'Add
                    If clsBG_M_TRANSFER_MASTER.Insert001(conn, trans) Then
                        success = True
                    End If
                Else
                    'Update
                    If clsBG_M_TRANSFER_MASTER.Update001(conn, trans) Then
                        success = True
                    End If
                End If
            End If

            If success Then
                trans.Commit()
            Else
                trans.Rollback()
            End If
        Catch ex As Exception
            'showErrorMessage(ex.Message)
        Finally
            conn.Close()
        End Try

        Return success
    End Function

    Public Function saveImportData(ByVal dt As DataTable) As Boolean
        Dim success As Boolean = False
        clsBG_M_TRANSFER_MASTER = New BG_M_TRANSFER_MASTER

        Dim conn As New SqlConnection
        Dim trans As SqlTransaction
        conn.ConnectionString = My.Settings.ConnStr

        conn.Open()
        trans = conn.BeginTransaction
        Try
            For Each row As DataRow In dt.Rows
                clsBG_M_TRANSFER_MASTER.BudgetOrderNo = row.Item("BUDGET_ORDER_NO").ToString
                clsBG_M_TRANSFER_MASTER.BudgetYear = row.Item("BUDGET_YEAR").ToString
                clsBG_M_TRANSFER_MASTER.PeriodType = row.Item("PERIOD_TYPE").ToString
                clsBG_M_TRANSFER_MASTER.TransferType = row.Item("TRANSFER_TYPE").ToString
                clsBG_M_TRANSFER_MASTER.TransferRate = row.Item("TRANSFER_RATE").ToString
                clsBG_M_TRANSFER_MASTER.FromOrderNo = row.Item("FROM_ORDER_NO").ToString
                clsBG_M_TRANSFER_MASTER.ToOrderNo = row.Item("TO_ORDER_NO").ToString
                If row.Item("CREATE_USER_ID").ToString = "" Then
                    clsBG_M_TRANSFER_MASTER.CreateUserID = p_strUserId
                Else
                    clsBG_M_TRANSFER_MASTER.CreateUserID = row.Item("CREATE_USER_ID").ToString
                End If
                If row.Item("CREATE_DATE").ToString = "" Then
                    clsBG_M_TRANSFER_MASTER.CreateDate = "GETDATE()"
                Else
                    clsBG_M_TRANSFER_MASTER.CreateDate = "'" & dateConvert(row.Item("CREATE_DATE").ToString) & "'"
                End If
                clsBG_M_TRANSFER_MASTER.UpdateUserID = p_strUserId
                clsBG_M_TRANSFER_MASTER.ProjectNo = row.Item("PROJECT_NO").ToString

                If clsBG_M_TRANSFER_MASTER.select002(conn, trans) Then
                    If CInt(clsBG_M_TRANSFER_MASTER.DTResult.Rows(0).Item(0).ToString) = 0 Then
                        'Add
                        If clsBG_M_TRANSFER_MASTER.Insert001(conn, trans) Then
                            success = True
                        Else
                            success = False
                            Exit For
                        End If
                    Else
                        'Update
                        If clsBG_M_TRANSFER_MASTER.Update001(conn, trans) Then
                            success = True
                        Else
                            success = False
                            Exit For
                        End If
                    End If
                End If
            Next

            If success Then
                trans.Commit()
            Else
                trans.Rollback()
            End If
        Catch ex As Exception

        Finally
            conn.Close()
        End Try

        Return success
    End Function

    Public Function deleteData() As Boolean
        Dim success As Boolean = False
        clsBG_M_TRANSFER_MASTER = New BG_M_TRANSFER_MASTER

        clsBG_M_TRANSFER_MASTER.BudgetOrderNo = Me.BudgetOrderNo
        clsBG_M_TRANSFER_MASTER.BudgetYear = Me.BudgetYear
        clsBG_M_TRANSFER_MASTER.PeriodType = Me.PeriodType
        clsBG_M_TRANSFER_MASTER.ProjectNo = Me.ProjectNo

        Dim conn As New SqlConnection
        Dim trans As SqlTransaction
        conn.ConnectionString = My.Settings.ConnStr

        conn.Open()
        trans = conn.BeginTransaction
        Try
            If clsBG_M_TRANSFER_MASTER.select002(conn, trans) Then
                If CInt(clsBG_M_TRANSFER_MASTER.DTResult.Rows(0).Item(0).ToString) = 1 Then
                    'Delete
                    If clsBG_M_TRANSFER_MASTER.Delete001(conn, trans) Then
                        success = True
                    End If
                End If
            End If

            If success Then
                trans.Commit()
            Else
                trans.Rollback()
            End If
        Catch ex As Exception
            'showErrorMessage(ex.Message)
        Finally
            conn.Close()
        End Try

        Return success
    End Function

#End Region

End Class
