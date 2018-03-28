Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.SqlClient

Public Class clsBG0650BL

#Region "Variable"
    Private myDtResult As DataTable
    Private myPersonNo As String = String.Empty
    Private myPersonName As String = String.Empty
    Private myPersonNoFilter As String = String.Empty
    Private myPersonNameFilter As String = String.Empty
    Private myCreateUserId As String = String.Empty
    Private myUpdateUserId As String = String.Empty
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
    Public Property PersonNo() As String
        Get
            Return myPersonNo
        End Get
        Set(ByVal value As String)
            myPersonNo = value
        End Set
    End Property
    Public Property PersonName() As String
        Get
            Return myPersonName
        End Get
        Set(ByVal value As String)
            myPersonName = value
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
    Public Property PersonNoFilter() As String
        Get
            Return myPersonNoFilter
        End Get
        Set(ByVal value As String)
            myPersonNoFilter = value
        End Set
    End Property
    Public Property PersonNameFilter() As String
        Get
            Return myPersonNameFilter
        End Get
        Set(ByVal value As String)
            myPersonNameFilter = value
        End Set
    End Property
#End Region

#Region "Function"
    Public Function searchDatagrid() As Boolean
        Dim clsBG_M_PERSON_IN_CHARGE As New BG_M_PERSON_IN_CHARGE

        clsBG_M_PERSON_IN_CHARGE.PersonNo = Me.PersonNoFilter
        clsBG_M_PERSON_IN_CHARGE.PersonName = Me.PersonNameFilter

        If clsBG_M_PERSON_IN_CHARGE.Select002 = True AndAlso _
        clsBG_M_PERSON_IN_CHARGE.DtResult.Rows.Count >= 1 Then
            Me.DtResult = clsBG_M_PERSON_IN_CHARGE.DtResult
            Return True
        Else
            Return False
        End If
    End Function
    Public Function checkData() As Boolean
        Dim clsBG_M_PERSON_IN_CHARGE As New BG_M_PERSON_IN_CHARGE

        clsBG_M_PERSON_IN_CHARGE.PersonNo = Me.PersonNo

        If clsBG_M_PERSON_IN_CHARGE.Select003 = True AndAlso _
                clsBG_M_PERSON_IN_CHARGE.DtResult.Rows.Count >= 1 Then
            Me.DtResult = clsBG_M_PERSON_IN_CHARGE.DtResult
            Return True
        Else
            Return False
        End If

    End Function
    Public Function insertOneData() As Boolean
        Dim clsBG_M_PERSON_IN_CHARGE As New BG_M_PERSON_IN_CHARGE

        clsBG_M_PERSON_IN_CHARGE.PersonNo = Me.PersonNo

        If clsBG_M_PERSON_IN_CHARGE.Select003 = True AndAlso _
        clsBG_M_PERSON_IN_CHARGE.DtResult.Rows.Count >= 1 Then
            Me.DtResult = clsBG_M_PERSON_IN_CHARGE.DtResult
            Return False
        End If

        clsBG_M_PERSON_IN_CHARGE.PersonName = Me.PersonName
        clsBG_M_PERSON_IN_CHARGE.CreateUserId = Me.CreateUserId

        If clsBG_M_PERSON_IN_CHARGE.Insert001 = True Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function insertExcelData(ByVal pConn As SqlConnection, _
                             ByVal pTrans As SqlTransaction) As Boolean
        Dim clsBG_M_PERSON_IN_CHARGE As New BG_M_PERSON_IN_CHARGE

        clsBG_M_PERSON_IN_CHARGE.PersonNo = Me.PersonNo
        clsBG_M_PERSON_IN_CHARGE.PersonName = Me.PersonName
        clsBG_M_PERSON_IN_CHARGE.CreateUserId = Me.CreateUserId

        If clsBG_M_PERSON_IN_CHARGE.Insert001(pConn, pTrans) = True Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function UpdateOneData() As Boolean
        Dim clsBG_M_PERSON_IN_CHARGE As New BG_M_PERSON_IN_CHARGE

        clsBG_M_PERSON_IN_CHARGE.PersonNo = Me.PersonNo
        clsBG_M_PERSON_IN_CHARGE.PersonName = Me.PersonName
        clsBG_M_PERSON_IN_CHARGE.UpdateUserId = Me.UpdateUserId

        If clsBG_M_PERSON_IN_CHARGE.Update001 = True Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function UpdateExcelData(ByVal pConn As SqlConnection, _
                             ByVal pTrans As SqlTransaction) As Boolean
        Dim clsBG_M_PERSON_IN_CHARGE As New BG_M_PERSON_IN_CHARGE

        clsBG_M_PERSON_IN_CHARGE.PersonNo = Me.PersonNo
        clsBG_M_PERSON_IN_CHARGE.PersonName = Me.PersonName
        clsBG_M_PERSON_IN_CHARGE.UpdateUserId = Me.UpdateUserId

        If clsBG_M_PERSON_IN_CHARGE.Update001(pConn, pTrans) = True Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function DeleteData() As Boolean
        Dim clsBG_M_PERSON_IN_CHARGE As New BG_M_PERSON_IN_CHARGE

        clsBG_M_PERSON_IN_CHARGE.PersonNo = Me.PersonNo

        If clsBG_M_PERSON_IN_CHARGE.Delete001 = True Then
            Return True
        Else
            Return False
        End If

    End Function
#End Region

End Class
