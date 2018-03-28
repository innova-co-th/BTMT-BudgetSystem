Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.SqlClient

Public Class clsBG0681BL

#Region "Variable"
    Private mydtResult As DataTable
    Private myAssetCategory As String
    Private myAssetCategoryTxt As String
    Private myCreateUserId As String
    Private myCreateDate As String
    Private myUpdateUserId As String
    Private myUpdateDate As String
    Private myAssetCategoryFilter As String = String.Empty
    Private myAssetCategoryTxtFilter As String = String.Empty
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
    Public Property AssetCategory() As String
        Get
            Return myAssetCategory
        End Get
        Set(ByVal value As String)
            myAssetCategory = value
        End Set
    End Property
    Public Property AssetCategoryTxt() As String
        Get
            Return myAssetCategoryTxt
        End Get
        Set(ByVal value As String)
            myAssetCategoryTxt = value
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
    Public Property AssetCategoryFilter() As String
        Get
            Return myAssetCategoryFilter
        End Get
        Set(ByVal value As String)
            myAssetCategoryFilter = value
        End Set
    End Property
    Public Property AssetCategoryTxtFilter() As String
        Get
            Return myAssetCategoryTxtFilter
        End Get
        Set(ByVal value As String)
            myAssetCategoryTxtFilter = value
        End Set
    End Property
#End Region

#Region "Function"

    Public Function getAssetCategory() As Boolean
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        clsBG_M_ASSET_GROUP.AssetCategoryNo = Me.AssetCategoryFilter
        clsBG_M_ASSET_GROUP.AssetCategoryName = Me.AssetCategoryTxtFilter

        If clsBG_M_ASSET_GROUP.Select004 = True AndAlso _
        clsBG_M_ASSET_GROUP.DtResult.Rows.Count >= 1 Then
            Me.DtResult = clsBG_M_ASSET_GROUP.DtResult
            Return True
        Else
            Return False
        End If
    End Function

    Public Function checkData() As Boolean
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        clsBG_M_ASSET_GROUP.AssetCategoryNo = Me.AssetCategory

        If clsBG_M_ASSET_GROUP.Select006 = True AndAlso _
                clsBG_M_ASSET_GROUP.DtResult.Rows.Count >= 1 Then
            Me.DtResult = clsBG_M_ASSET_GROUP.DtResult
            Return True
        Else
            Return False
        End If
    End Function

    Public Function insertOneData() As Boolean
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        clsBG_M_ASSET_GROUP.AssetCategoryNo = Me.AssetCategory

        If clsBG_M_ASSET_GROUP.Select006 = True AndAlso _
        clsBG_M_ASSET_GROUP.DtResult.Rows.Count >= 1 Then
            Me.DtResult = clsBG_M_ASSET_GROUP.DtResult
            Return False
        End If

        clsBG_M_ASSET_GROUP.AssetCategoryName = Me.AssetCategoryTxt
        clsBG_M_ASSET_GROUP.CreateUserId = Me.CreateUserId

        If clsBG_M_ASSET_GROUP.Insert002 = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function insertExcelData(ByVal pConn As SqlConnection, _
                             ByVal pTrans As SqlTransaction) As Boolean
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        clsBG_M_ASSET_GROUP.AssetCategoryNo = Me.AssetCategory
        clsBG_M_ASSET_GROUP.AssetCategoryName = Me.AssetCategoryTxt
        clsBG_M_ASSET_GROUP.CreateUserId = Me.CreateUserId

        If clsBG_M_ASSET_GROUP.Insert002(pConn, pTrans) = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function UpdateOneData() As Boolean
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        clsBG_M_ASSET_GROUP.AssetCategoryNo = Me.AssetCategory
        clsBG_M_ASSET_GROUP.AssetCategoryName = Me.AssetCategoryTxt
        clsBG_M_ASSET_GROUP.UpdateUserId = Me.UpdateUserId

        If clsBG_M_ASSET_GROUP.Update002 = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function UpdateExcelData(ByVal pConn As SqlConnection, _
                                     ByVal pTrans As SqlTransaction) As Boolean
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        clsBG_M_ASSET_GROUP.AssetCategoryNo = Me.AssetCategory
        clsBG_M_ASSET_GROUP.AssetCategoryName = Me.AssetCategoryTxt
        clsBG_M_ASSET_GROUP.UpdateUserId = Me.UpdateUserId

        If clsBG_M_ASSET_GROUP.Update002(pConn, pTrans) = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function DeleteData() As Boolean
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        clsBG_M_ASSET_GROUP.AssetCategoryNo = Me.AssetCategory

        If clsBG_M_ASSET_GROUP.Delete002 = True Then
            Return True
        Else
            Return False
        End If

    End Function
#End Region

End Class
