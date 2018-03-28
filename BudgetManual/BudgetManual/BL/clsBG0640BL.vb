Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.SqlClient

Public Class clsBG0640BL

#Region "Variable"
    Private myDtResult As DataTable
    Private myDeptNo As String
    Private myDeptName As String
    Private myCreateUserId As String
    Private myUpdateUserId As String
    Private myDeptNoFilter As String
    Private myDeptNameFilter As String
#End Region

#Region "Property"
    Public Property DtResult() As DataTable
        Get
            Return mydtResult
        End Get
        Set(ByVal value As DataTable)
            mydtResult = value
        End Set
    End Property
    Public Property DeptNo() As String
        Get
            Return myDeptNo
        End Get
        Set(ByVal value As String)
            myDeptNo = value
        End Set
    End Property
    Public Property DeptName() As String
        Get
            Return myDeptName
        End Get
        Set(ByVal value As String)
            myDeptName = value
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
    Public Property DeptNoFilter() As String
        Get
            Return myDeptNoFilter
        End Get
        Set(ByVal value As String)
            myDeptNoFilter = value
        End Set
    End Property
    Public Property DeptNameFilter() As String
        Get
            Return myDeptNameFilter
        End Get
        Set(ByVal value As String)
            myDeptNameFilter = value
        End Set
    End Property
#End Region

#Region "Function"
    Public Function searchDatagrid() As Boolean
        Dim clsBG_M_DEPT As New BG_M_DEPT

        clsBG_M_DEPT.DeptNo = Me.DeptNoFilter
        clsBG_M_DEPT.DeptName = Me.DeptNameFilter

        If clsBG_M_DEPT.Select002 = True AndAlso _
        clsBG_M_DEPT.DtResult.Rows.Count >= 1 Then
            Me.DtResult = clsBG_M_DEPT.DtResult
            Return True
        Else
            Return False
        End If
    End Function
    Public Function checkData() As Boolean
        Dim clsBG_M_DEPT As New BG_M_DEPT

        clsBG_M_DEPT.DeptNo = Me.DeptNo

        If clsBG_M_DEPT.Select003 = True AndAlso _
                clsBG_M_DEPT.DtResult.Rows.Count >= 1 Then
            Me.DtResult = clsBG_M_DEPT.DtResult
            Return True
        Else
            Return False
        End If

    End Function
    Public Function insertOneData() As Boolean
        Dim clsBG_M_DEPT As New BG_M_DEPT

        clsBG_M_DEPT.DeptNo = Me.DeptNo

        If clsBG_M_DEPT.Select003 = True AndAlso _
        clsBG_M_DEPT.DtResult.Rows.Count >= 1 Then
            Me.DtResult = clsBG_M_DEPT.DtResult
            Return False
        End If

        clsBG_M_DEPT.DeptName = Me.DeptName
        clsBG_M_DEPT.CreateUserId = Me.CreateUserId

        If clsBG_M_DEPT.Insert001 = True Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function insertExcelData(ByVal pConn As SqlConnection, _
                             ByVal pTrans As SqlTransaction) As Boolean
        Dim clsBG_M_DEPT As New BG_M_DEPT

        clsBG_M_DEPT.DeptNo = Me.DeptNo
        clsBG_M_DEPT.DeptName = Me.DeptName
        clsBG_M_DEPT.CreateUserId = Me.CreateUserId

        If clsBG_M_DEPT.Insert001(pConn, pTrans) = True Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function UpdateOneData() As Boolean
        Dim clsBG_M_DEPT As New BG_M_DEPT

        clsBG_M_DEPT.DeptNo = Me.DeptNo
        clsBG_M_DEPT.DeptName = Me.DeptName
        clsBG_M_DEPT.UpdateUserId = Me.UpdateUserId

        If clsBG_M_DEPT.Update001 = True Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function UpdateExcelData(ByVal pConn As SqlConnection, _
                             ByVal pTrans As SqlTransaction) As Boolean
        Dim clsBG_M_DEPT As New BG_M_DEPT

        clsBG_M_DEPT.DeptNo = Me.DeptNo
        clsBG_M_DEPT.DeptName = Me.DeptName
        clsBG_M_DEPT.UpdateUserId = Me.UpdateUserId

        If clsBG_M_DEPT.Update001(pConn, pTrans) = True Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function DeleteData() As Boolean
        Dim clsBG_M_DEPT As New BG_M_DEPT

        clsBG_M_DEPT.DeptNo = Me.DeptNo

        If clsBG_M_DEPT.Delete001 = True Then
            Return True
        Else
            Return False
        End If

    End Function
#End Region

End Class
