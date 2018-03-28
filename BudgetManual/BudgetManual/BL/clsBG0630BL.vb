Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.SqlClient

Public Class clsBG0630BL

#Region "Variable"
    Private myDtResult As DataTable
    Private myAccountNo As String = String.Empty
    Private myAccountName As String = String.Empty
    Private myCreateUserId As String = String.Empty
    Private myUpdateUserId As String = String.Empty
    Private myAccountNoFilter As String = String.Empty
    Private myAccountNameFilter As String = String.Empty
#End Region

#Region "Property"
    Property DtResult() As DataTable
        Get
            Return myDtResult
        End Get
        Set(ByVal value As DataTable)
            myDtResult = value
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
    Public Property AccountName() As String
        Get
            Return myAccountName
        End Get
        Set(ByVal value As String)
            myAccountName = value
        End Set
    End Property
    Public Property CreateUserId() As String
        Get
            Return myCreateUserId
        End Get
        Set(ByVal value As String)
            myCreateUserId = value
        End Set
    End Property
    Public Property UpdateUserId() As String
        Get
            Return myUpdateUserId
        End Get
        Set(ByVal value As String)
            myUpdateUserId = value
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
    Public Function searchDatagrid() As Boolean
        Dim clsBG_M_ACCOUNT As New BG_M_ACCOUNT

        clsBG_M_ACCOUNT.AccountNo = Me.AccountNoFilter
        clsBG_M_ACCOUNT.AccountName = Me.AccountNameFilter

        If clsBG_M_ACCOUNT.Select002 = True AndAlso _
        clsBG_M_ACCOUNT.DtResult.Rows.Count >= 1 Then
            Me.DtResult = clsBG_M_ACCOUNT.DtResult
            Return True
        Else
            Return False
        End If
    End Function
    Public Function checkData(Optional ByVal pConn As SqlConnection = Nothing) As Boolean
        Dim clsBG_M_ACCOUNT As New BG_M_ACCOUNT

        clsBG_M_ACCOUNT.AccountNo = Me.AccountNo
        If pConn Is Nothing Then
            If clsBG_M_ACCOUNT.Select003 = True AndAlso _
                    clsBG_M_ACCOUNT.DtResult.Rows.Count >= 1 Then
                Me.DtResult = clsBG_M_ACCOUNT.DtResult
                Return True
            Else
                Return False
            End If
        Else
            If clsBG_M_ACCOUNT.Select003(pConn) = True AndAlso _
                   clsBG_M_ACCOUNT.DtResult.Rows.Count >= 1 Then
                Me.DtResult = clsBG_M_ACCOUNT.DtResult
                Return True
            Else
                Return False
            End If
        End If

    End Function

    Public Function insertOneData() As Boolean
        Dim clsBG_M_ACCOUNT As New BG_M_ACCOUNT

        clsBG_M_ACCOUNT.AccountNo = Me.AccountNo

        If clsBG_M_ACCOUNT.Select003 = True AndAlso _
        clsBG_M_ACCOUNT.DtResult.Rows.Count >= 1 Then
            Me.DtResult = clsBG_M_ACCOUNT.DtResult
            Return False
        End If

        clsBG_M_ACCOUNT.AccountName = Me.AccountName
        clsBG_M_ACCOUNT.CreateUserId = Me.CreateUserId

        If clsBG_M_ACCOUNT.Insert001 = True Then
            Return True
        Else
            Return False
        End If

    End Function
    Public Function insertExcelData(ByVal pConn As SqlConnection, _
                               ByVal pTrans As SqlTransaction) As Boolean
        Dim clsBG_M_ACCOUNT As New BG_M_ACCOUNT

        clsBG_M_ACCOUNT.AccountNo = Me.AccountNo
        clsBG_M_ACCOUNT.AccountName = Me.AccountName
        clsBG_M_ACCOUNT.CreateUserId = Me.CreateUserId

        If clsBG_M_ACCOUNT.Insert001(pConn, pTrans) = True Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function UpdateOneData() As Boolean
        Dim clsBG_M_ACCOUNT As New BG_M_ACCOUNT

        clsBG_M_ACCOUNT.AccountNo = Me.AccountNo
        clsBG_M_ACCOUNT.AccountName = Me.AccountName
        clsBG_M_ACCOUNT.UpdateUserId = Me.UpdateUserId

        If clsBG_M_ACCOUNT.Update001 = True Then
            Return True
        Else
            Return False
        End If

    End Function
    Public Function UpdateExcelData(ByVal pConn As SqlConnection, _
                               ByVal pTrans As SqlTransaction) As Boolean
        Dim clsBG_M_ACCOUNT As New BG_M_ACCOUNT

        clsBG_M_ACCOUNT.AccountNo = Me.AccountNo
        clsBG_M_ACCOUNT.AccountName = Me.AccountName
        clsBG_M_ACCOUNT.UpdateUserId = Me.UpdateUserId

        If clsBG_M_ACCOUNT.Update001(pConn, pTrans) = True Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function DeleteData() As Boolean
        Dim clsBG_M_ACCOUNT As New BG_M_ACCOUNT

        clsBG_M_ACCOUNT.AccountNo = Me.AccountNo

        If clsBG_M_ACCOUNT.Delete001 = True Then
            Return True
        Else
            Return False
        End If

    End Function
#End Region

End Class
