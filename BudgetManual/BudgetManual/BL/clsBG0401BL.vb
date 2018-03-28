Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon
Imports System.Data.SqlClient

Public Class clsBG0401BL

#Region "Variable"
    Private myDtResult As DataTable
    Private myDtLst As DataTable
    Private myUpdateUserId As String = String.Empty
    Private myPICflg As String
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
    Property DtLst() As DataTable
        Get
            Return myDtLst
        End Get
        Set(ByVal value As DataTable)
            myDtLst = value
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
    Public Property PICflg() As String
        Get
            Return myPICflg
        End Get
        Set(ByVal value As String)
            myPICflg = value
        End Set
    End Property
#End Region

#Region "Function"
    Public Function searchListBox() As Boolean
        Dim clsBG_M_BUDGET_ORDER As New BG_M_BUDGET_ORDER

        If clsBG_M_BUDGET_ORDER.Select004 = True AndAlso _
        clsBG_M_BUDGET_ORDER.dtResult.Rows.Count >= 1 Then
            Me.DtResult = clsBG_M_BUDGET_ORDER.dtResult
            Return True
        Else
            Return False
        End If
    End Function

    Public Function updateDataList() As Boolean
        Dim clsBG_M_BUDGET_ORDER As New BG_M_BUDGET_ORDER
        Dim conn As SqlConnection
        Dim trans As SqlTransaction
        Dim result As Boolean = False

        conn = New SqlConnection(My.Settings.ConnStr)
        conn.Open()
        trans = conn.BeginTransaction()

        Try

            For Each row As DataRow In Me.DtLst.Rows

                clsBG_M_BUDGET_ORDER.PersonInCharge = CStr(row("PERSON_IN_CHARGE_NO"))
                clsBG_M_BUDGET_ORDER.PICShowFlag = Me.PICflg
                clsBG_M_BUDGET_ORDER.UpdateUserId = Me.UpdateUserId

                If clsBG_M_BUDGET_ORDER.Update002(conn, trans) Then
                    result = True
                End If
            Next

            If result Then
                trans.Commit()
            Else
                trans.Rollback()
            End If
        Catch ex As Exception
            trans.Rollback()
            showErrorMessage("Error: " & ex.Message)
        Finally
            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If
        End Try

        Return result

    End Function
#End Region

End Class
