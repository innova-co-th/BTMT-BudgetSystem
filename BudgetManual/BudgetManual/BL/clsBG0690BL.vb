Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.SqlClient

Public Class clsBG0690BL

#Region "Variable"
    Private mydtResult As DataTable
    Private myPicList As DataTable
    Private myParentNo As String
    Private myChildNo As String
    Private myUpdateUserId As String
    Private myOldChildNo As String
    Private myParentNoFilter As String
    Private myChildNoFilter As String
#End Region

#Region "Property"

#Region "dtResult"
    Public Property dtResult() As DataTable
        Get
            Return mydtResult
        End Get
        Set(ByVal value As DataTable)
            mydtResult = value
        End Set
    End Property
#End Region

#Region "PicList"
    Public Property PicList() As DataTable
        Get
            Return myPicList
        End Get
        Set(ByVal value As DataTable)
            myPicList = value
        End Set
    End Property
#End Region

    Public Property ParentNo() As String
        Get
            Return myParentNo
        End Get
        Set(ByVal value As String)
            myParentNo = value
        End Set
    End Property

    Public Property ChildNo() As String
        Get
            Return myChildNo
        End Get
        Set(ByVal value As String)
            myChildNo = value
        End Set
    End Property

    Public Property OldChildNo() As String
        Get
            Return myOldChildNo
        End Get
        Set(ByVal value As String)
            myOldChildNo = value
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
    Public Property ParentNoFilter() As String
        Get
            Return myParentNoFilter
        End Get
        Set(ByVal value As String)
            myParentNoFilter = value
        End Set
    End Property

    Public Property ChildNoFilter() As String
        Get
            Return myChildNoFilter
        End Get
        Set(ByVal value As String)
            myChildNoFilter = value
        End Set
    End Property
#End Region

#Region "Function"

    Public Function searchDatagrid() As Boolean
        Dim clsBG_M_CHILD_PIC As New BG_M_CHILD_PIC

        clsBG_M_CHILD_PIC.ParentNo = Me.ParentNoFilter
        clsBG_M_CHILD_PIC.ChildNo = Me.ChildNoFilter

        If clsBG_M_CHILD_PIC.Select001 = True AndAlso clsBG_M_CHILD_PIC.DtResult.Rows.Count > 0 Then
            Me.dtResult = clsBG_M_CHILD_PIC.DtResult

            Return True
        Else

            Return False
        End If
    End Function

    Public Function searchCombo() As Boolean
        Dim clsBG_M_PERSON_IN_CHARGE As New BG_M_PERSON_IN_CHARGE

        If clsBG_M_PERSON_IN_CHARGE.Select008 = True AndAlso clsBG_M_PERSON_IN_CHARGE.DtResult.Rows.Count > 0 Then
            Me.PicList = clsBG_M_PERSON_IN_CHARGE.DtResult

            Return True
        Else

            Return False
        End If
    End Function

    Public Function insertOneData() As Boolean
        Dim clsBG_M_CHILD_PIC As New BG_M_CHILD_PIC

        clsBG_M_CHILD_PIC.ParentNo = Me.ParentNo
        clsBG_M_CHILD_PIC.ChildNo = Me.ChildNo
        clsBG_M_CHILD_PIC.UpdateUserId = Me.UpdateUserId

        If clsBG_M_CHILD_PIC.Select002 = True AndAlso clsBG_M_CHILD_PIC.DtResult.Rows.Count > 0 Then
            Me.dtResult = clsBG_M_CHILD_PIC.DtResult

            Return False
        End If

        If clsBG_M_CHILD_PIC.Insert001 = True Then

            Return True
        Else

            Return False
        End If
    End Function

    Public Function insertExcelData(ByVal pConn As SqlConnection, _
                                    ByVal pTrans As SqlTransaction) As Boolean
        Dim clsBG_M_CHILD_PIC As New BG_M_CHILD_PIC

        clsBG_M_CHILD_PIC.ParentNo = Me.ParentNo
        clsBG_M_CHILD_PIC.ChildNo = Me.ChildNo
        clsBG_M_CHILD_PIC.UpdateUserId = Me.UpdateUserId

        If clsBG_M_CHILD_PIC.Insert001(pConn, pTrans) = True Then

            Return True
        Else

            Return False
        End If
    End Function

    Public Function UpdateOneData() As Boolean
        Dim clsBG_M_CHILD_PIC As New BG_M_CHILD_PIC

        clsBG_M_CHILD_PIC.ParentNo = Me.ParentNo
        clsBG_M_CHILD_PIC.ChildNo = Me.ChildNo
        clsBG_M_CHILD_PIC.OldChildNo = Me.OldChildNo
        clsBG_M_CHILD_PIC.UpdateUserId = Me.UpdateUserId

        If clsBG_M_CHILD_PIC.Select002 = True AndAlso clsBG_M_CHILD_PIC.DtResult.Rows.Count > 0 Then
            Me.dtResult = clsBG_M_CHILD_PIC.DtResult

            Return False
        End If

        If clsBG_M_CHILD_PIC.Update001 = True Then

            Return True
        Else

            Return False
        End If
    End Function

    Public Function DeleteData() As Boolean
        Dim clsBG_M_CHILD_PIC As New BG_M_CHILD_PIC

        clsBG_M_CHILD_PIC.ParentNo = Me.ParentNo
        clsBG_M_CHILD_PIC.ChildNo = Me.ChildNo

        If clsBG_M_CHILD_PIC.Delete001 = True Then

            Return True
        Else

            Return False
        End If
    End Function

    Public Function DeleteAllData(ByVal pConn As SqlConnection, _
                             ByVal pTrans As SqlTransaction) As Boolean
        Dim clsBG_M_CHILD_PIC As New BG_M_CHILD_PIC

        If clsBG_M_CHILD_PIC.Delete002(pConn, pTrans) = True Then

            Return True
        Else

            Return False
        End If
    End Function
#End Region

End Class
