Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.SqlClient

Public Class clsBG0680BL

#Region "Variable"
    Private mydtResult As DataTable
    Private myAssetGroupNo As String
    Private myAssetGroupName As String
    Private myAssetProj As String
    Private myAssetCate As String
    Private myCreateUserId As String
    Private myCreateDate As String
    Private myUpdateUserId As String
    Private myUpdateDate As String
    Private myAssetGroupNoFilter As String = String.Empty
    Private myAssetGroupNameFilter As String = String.Empty
    Private myAssetProjFilter As String = String.Empty
    Private myAssetCateFilter As String = String.Empty
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
    Public Property AssetGroupNo() As String
        Get
            Return myAssetGroupNo
        End Get
        Set(ByVal value As String)
            myAssetGroupNo = value
        End Set
    End Property
    Public Property AssetGroupName() As String
        Get
            Return myAssetGroupName
        End Get
        Set(ByVal value As String)
            myAssetGroupName = value
        End Set
    End Property
    Public Property AssetProj() As String
        Get
            Return myAssetProj
        End Get
        Set(ByVal value As String)
            myAssetProj = value
        End Set
    End Property
    Public Property AssetCate() As String
        Get
            Return myAssetCate
        End Get
        Set(ByVal value As String)
            myAssetCate = value
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
    Public Property CreateDate() As String
        Get
            Return myCreateDate
        End Get
        Set(ByVal value As String)
            myCreateDate = value
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
    Public Property UpdateDate() As String
        Get
            Return myUpdateDate
        End Get
        Set(ByVal value As String)
            myUpdateDate = value
        End Set
    End Property
    Public Property AssetGroupNoFilter() As String
        Get
            Return myAssetGroupNoFilter
        End Get
        Set(ByVal value As String)
            myAssetGroupNoFilter = value
        End Set
    End Property
    Public Property AssetGroupNameFilter() As String
        Get
            Return myAssetGroupNameFilter
        End Get
        Set(ByVal value As String)
            myAssetGroupNameFilter = value
        End Set
    End Property
    Public Property AssetProjFilter() As String
        Get
            Return myAssetProjFilter
        End Get
        Set(ByVal value As String)
            myAssetProjFilter = value
        End Set
    End Property
    Public Property AssetCateFilter() As String
        Get
            Return myAssetCateFilter
        End Get
        Set(ByVal value As String)
            myAssetCateFilter = value
        End Set
    End Property
#End Region

#Region "Function"
    Public Function searchDatagrid() As Boolean
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        clsBG_M_ASSET_GROUP.AssetGroupNo = Me.AssetGroupNoFilter
        clsBG_M_ASSET_GROUP.AssetGroupName = Me.AssetGroupNameFilter
        clsBG_M_ASSET_GROUP.AssetProjectNo = Me.AssetProjFilter
        clsBG_M_ASSET_GROUP.AssetCategoryNo = Me.AssetCateFilter

        If clsBG_M_ASSET_GROUP.Select002 = True AndAlso _
        clsBG_M_ASSET_GROUP.DtResult.Rows.Count >= 1 Then
            Me.DtResult = clsBG_M_ASSET_GROUP.DtResult
            Return True
        Else
            Return False
        End If
    End Function

    Public Function getAssetCategory() As DataTable
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        If clsBG_M_ASSET_GROUP.Select004 = True AndAlso _
        clsBG_M_ASSET_GROUP.DtResult.Rows.Count >= 1 Then
            Return clsBG_M_ASSET_GROUP.DtResult
        Else
            Return Nothing
        End If
    End Function

    Public Function getAssetProject() As DataTable
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        If clsBG_M_ASSET_GROUP.Select005 = True AndAlso _
        clsBG_M_ASSET_GROUP.DtResult.Rows.Count >= 1 Then
            Return clsBG_M_ASSET_GROUP.DtResult
        Else
            Return Nothing
        End If
    End Function

    Public Function checkData() As Boolean
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        clsBG_M_ASSET_GROUP.AssetGroupNo = Me.AssetGroupNo

        If clsBG_M_ASSET_GROUP.Select003 = True AndAlso _
                clsBG_M_ASSET_GROUP.DtResult.Rows.Count >= 1 Then
            Me.DtResult = clsBG_M_ASSET_GROUP.DtResult
            Return True
        Else
            Return False
        End If
    End Function

    Public Function insertOneData() As Boolean
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        clsBG_M_ASSET_GROUP.AssetGroupNo = Me.AssetGroupNo

        If clsBG_M_ASSET_GROUP.Select003 = True AndAlso _
        clsBG_M_ASSET_GROUP.DtResult.Rows.Count >= 1 Then
            Me.DtResult = clsBG_M_ASSET_GROUP.DtResult
            Return False
        End If

        clsBG_M_ASSET_GROUP.AssetGroupName = Me.AssetGroupName
        clsBG_M_ASSET_GROUP.AssetProjectNo = Me.AssetProj
        clsBG_M_ASSET_GROUP.AssetCategoryNo = Me.AssetCate
        clsBG_M_ASSET_GROUP.CreateUserId = Me.CreateUserId

        If clsBG_M_ASSET_GROUP.Insert001 = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function insertExcelData(ByVal pConn As SqlConnection, _
                             ByVal pTrans As SqlTransaction) As Boolean
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        clsBG_M_ASSET_GROUP.AssetGroupNo = Me.AssetGroupNo
        clsBG_M_ASSET_GROUP.AssetGroupName = Me.AssetGroupName
        clsBG_M_ASSET_GROUP.AssetProjectNo = Me.AssetProj
        clsBG_M_ASSET_GROUP.AssetCategoryNo = Me.AssetCate
        clsBG_M_ASSET_GROUP.CreateUserId = Me.CreateUserId

        If clsBG_M_ASSET_GROUP.Insert001(pConn, pTrans) = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function UpdateOneData() As Boolean
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        clsBG_M_ASSET_GROUP.AssetGroupNo = Me.AssetGroupNo
        clsBG_M_ASSET_GROUP.AssetGroupName = Me.AssetGroupName
        clsBG_M_ASSET_GROUP.AssetProjectNo = Me.AssetProj
        clsBG_M_ASSET_GROUP.AssetCategoryNo = Me.AssetCate
        clsBG_M_ASSET_GROUP.UpdateUserId = Me.UpdateUserId

        If clsBG_M_ASSET_GROUP.Update001 = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function UpdateExcelData(ByVal pConn As SqlConnection, _
                             ByVal pTrans As SqlTransaction) As Boolean
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        clsBG_M_ASSET_GROUP.AssetGroupNo = Me.AssetGroupNo
        clsBG_M_ASSET_GROUP.AssetGroupName = Me.AssetGroupName
        clsBG_M_ASSET_GROUP.AssetProjectNo = Me.AssetProj
        clsBG_M_ASSET_GROUP.AssetCategoryNo = Me.AssetCate
        clsBG_M_ASSET_GROUP.UpdateUserId = Me.UpdateUserId

        If clsBG_M_ASSET_GROUP.Update001(pConn, pTrans) = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function DeleteData() As Boolean
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        clsBG_M_ASSET_GROUP.AssetGroupNo = Me.AssetGroupNo

        If clsBG_M_ASSET_GROUP.Delete001 = True Then
            Return True
        Else
            Return False
        End If

    End Function
#End Region

End Class
