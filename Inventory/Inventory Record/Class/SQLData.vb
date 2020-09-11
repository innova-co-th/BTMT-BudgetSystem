Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports Inventory_Record.Common
Imports System.Configuration

#Region "SQLData"
Public Class SQLData

    Public Strcon As String

    Public Sub New()
        'Database ACCINV
        Strcon = ConfigurationManager.ConnectionStrings("INV").ToString()
    End Sub
    Public Sub New(ByVal DBName As String)
        'Database ACCINV
        If DBName.Equals("ACCINV") Then
            Strcon = ConfigurationManager.ConnectionStrings("INV").ToString()
        End If
        'Database BTMTMASTER
        If DBName.Equals("BTMTMASTER") Then
            Strcon = ConfigurationManager.ConnectionStrings("BTMT").ToString()
        End If
    End Sub


    Public Function GetDataset(ByVal Strsql As String) As DataSet
        Dim DA As New SqlDataAdapter(Strsql, Strcon)
        Dim DS As New DataSet
        Try
            DA.Fill(DS)
        Catch ex As Exception
            MsgBox(ex.Message, 48)
        Finally
        End Try
        Return DS
    End Function

    Public Function GetDataTable(ByVal Strsql As String) As DataTable
        Dim DA As New SqlDataAdapter(Strsql, Strcon)
        Dim DT As New DataTable
        Try
            DA.Fill(DT)
        Catch ex As Exception
            MsgBox(ex.Message, 48)
        Finally

        End Try
        Return DT
    End Function

End Class
#End Region