Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.SqlClient

Public Class clsBG0682BL

#Region "Variable"
    Private mydtResult As DataTable
    Private myAssetProject As String
    Private myAssetProjectTxt As String
    Private myCreateUserId As String
    Private myCreateDate As String
    Private myUpdateUserId As String
    Private myUpdateDate As String
    Private myAssetProjectFilter As String = String.Empty
    Private myAssetProjectTxtFilter As String = String.Empty
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
    Public Property AssetProject() As String
        Get
            Return myAssetProject
        End Get
        Set(ByVal value As String)
            myAssetProject = value
        End Set
    End Property
    Public Property AssetProjectTxt() As String
        Get
            Return myAssetProjectTxt
        End Get
        Set(ByVal value As String)
            myAssetProjectTxt = value
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
    Public Property AssetProjectFilter() As String
        Get
            Return myAssetProjectFilter
        End Get
        Set(ByVal value As String)
            myAssetProjectFilter = value
        End Set
    End Property
    Public Property AssetProjectTxtFilter() As String
        Get
            Return myAssetProjectTxtFilter
        End Get
        Set(ByVal value As String)
            myAssetProjectTxtFilter = value
        End Set
    End Property
#End Region

#Region "Function"

    Public Function getAssetProject() As Boolean
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        clsBG_M_ASSET_GROUP.AssetProjectNo = Me.AssetProjectFilter
        clsBG_M_ASSET_GROUP.AssetProjectName = Me.AssetProjectTxtFilter

        If clsBG_M_ASSET_GROUP.Select005 = True AndAlso _
        clsBG_M_ASSET_GROUP.DtResult.Rows.Count >= 1 Then
            Me.DtResult = clsBG_M_ASSET_GROUP.DtResult
            Return True
        Else
            Return False
        End If
    End Function

    Public Function checkData() As Boolean
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        clsBG_M_ASSET_GROUP.AssetProjectNo = Me.AssetProject

        If clsBG_M_ASSET_GROUP.Select007 = True AndAlso _
                clsBG_M_ASSET_GROUP.DtResult.Rows.Count >= 1 Then
            Me.DtResult = clsBG_M_ASSET_GROUP.DtResult
            Return True
        Else
            Return False
        End If
    End Function

    Public Function insertOneData() As Boolean
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        clsBG_M_ASSET_GROUP.AssetProjectNo = Me.AssetProject

        If clsBG_M_ASSET_GROUP.Select007 = True AndAlso _
        clsBG_M_ASSET_GROUP.DtResult.Rows.Count >= 1 Then
            Me.DtResult = clsBG_M_ASSET_GROUP.DtResult
            Return False
        End If

        clsBG_M_ASSET_GROUP.AssetProjectName = Me.AssetProjectTxt
        clsBG_M_ASSET_GROUP.CreateUserId = Me.CreateUserId

        If clsBG_M_ASSET_GROUP.Insert003 = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function insertExcelData(ByVal pConn As SqlConnection, _
                                    ByVal pTrans As SqlTransaction) As Boolean
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        clsBG_M_ASSET_GROUP.AssetProjectNo = Me.AssetProject
        clsBG_M_ASSET_GROUP.AssetProjectName = Me.AssetProjectTxt
        clsBG_M_ASSET_GROUP.CreateUserId = Me.CreateUserId

        If clsBG_M_ASSET_GROUP.Insert003(pConn, pTrans) = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function UpdateOneData() As Boolean
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        clsBG_M_ASSET_GROUP.AssetProjectNo = Me.AssetProject
        clsBG_M_ASSET_GROUP.AssetProjectName = Me.AssetProjectTxt
        clsBG_M_ASSET_GROUP.UpdateUserId = Me.UpdateUserId

        If clsBG_M_ASSET_GROUP.Update003 = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function UpdateExcelData(ByVal pConn As SqlConnection, _
                                     ByVal pTrans As SqlTransaction) As Boolean
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        clsBG_M_ASSET_GROUP.AssetProjectNo = Me.AssetProject
        clsBG_M_ASSET_GROUP.AssetProjectName = Me.AssetProjectTxt
        clsBG_M_ASSET_GROUP.UpdateUserId = Me.UpdateUserId

        If clsBG_M_ASSET_GROUP.Update003(pConn, pTrans) = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function DeleteData() As Boolean
        Dim clsBG_M_ASSET_GROUP As New BG_M_ASSET_GROUP

        clsBG_M_ASSET_GROUP.AssetProjectNo = Me.AssetProject

        If clsBG_M_ASSET_GROUP.Delete003 = True Then
            Return True
        Else
            Return False
        End If

    End Function
#End Region

End Class
