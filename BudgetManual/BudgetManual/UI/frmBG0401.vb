Imports BudgetManual.BGConstant
Imports BudgetManual.BGCommon

Public Class frmBG0401

#Region "Variable"
    Private myClsBG0401BL As New clsBG0401BL
    Dim dtPic As DataTable
    Dim dtOther As DataTable
#End Region

#Region "Overrides Function"
    Public Sub New(ByRef frmParent As Form, ByVal strFormName As String, ByVal blnMaximize As Boolean)
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.MdiParent = frmParent
        If blnMaximize Then
            Me.WindowState = FormWindowState.Maximized
        Else
            Me.WindowState = FormWindowState.Normal
        End If
        Me.Text = strFormName
       
        setList()

    End Sub

#End Region

#Region "Function"
    Private Function setTable(ByVal strSelect As String, ByVal dtTmp As DataTable) As DataTable
        Dim dt As DataTable
        Dim dr As DataRow()

        dt = dtTmp.Clone
        dr = dtTmp.Select(strSelect)

        For Each row As DataRow In dr
            dt.ImportRow(row)
        Next

        Return dt

    End Function

    Private Sub setList()
        Dim dtTmp As New DataTable

        If myClsBG0401BL.searchListBox = True Then
            dtTmp = myClsBG0401BL.DtResult
            If dtTmp.Rows.Count > 0 Then
                Me.dtPic = setTable("PIC_SHOW_FLAG=1", dtTmp)
                Me.dtOther = setTable("PIC_SHOW_FLAG=0", dtTmp)

                lstPIC.DataSource = dtPic
                lstPIC.ValueMember = "PERSON_IN_CHARGE_NO"
                lstPIC.DisplayMember = "PIC_DISPLAY"

                lstOther.DataSource = dtOther
                lstOther.ValueMember = "PERSON_IN_CHARGE_NO"
                lstOther.DisplayMember = "PIC_DISPLAY"
            End If

        End If
    End Sub

#End Region

#Region "Control Event"

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        myClsBG0401BL.UpdateUserId = p_strUserId

        myClsBG0401BL.DtLst = Me.dtPic
        myClsBG0401BL.PICflg = "1"
        If Not Me.dtPic Is Nothing AndAlso Me.dtPic.Rows.Count > 0 Then
            If Not myClsBG0401BL.updateDataList Then
                Exit Sub
            End If
        End If


        myClsBG0401BL.DtLst = Me.dtOther
        myClsBG0401BL.PICflg = "0"
        If Not Me.dtOther Is Nothing AndAlso Me.dtOther.Rows.Count > 0 Then
            If Not myClsBG0401BL.updateDataList Then
                Exit Sub
            End If
        End If

        MessageBox.Show("Report Options was updated", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
        setList()
    End Sub

    Private Sub cmdToPIC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdToPIC.Click

        If lstOther.SelectedIndex < 0 Then
            Exit Sub
        End If
        Dim strNo As String = CStr(Me.lstOther.SelectedValue)
        Dim dr As DataRow() = dtOther.Select("PERSON_IN_CHARGE_NO='" & strNo & "'")

        For Each r As DataRow In dr
            dtPic.ImportRow(r)
            dtOther.Rows.Remove(r)
        Next
        dtOther.AcceptChanges()

        Me.lstPIC.DataSource = dtPic
        Me.lstPIC.ValueMember = "PERSON_IN_CHARGE_NO"
        Me.lstPIC.DisplayMember = "PIC_DISPLAY"

        Me.lstOther.DataSource = dtOther
        Me.lstOther.ValueMember = "PERSON_IN_CHARGE_NO"
        Me.lstOther.DisplayMember = "PIC_DISPLAY"

    End Sub

    Private Sub cmdToOther_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdToOther.Click

        If lstPIC.SelectedIndex < 0 Then
            Exit Sub
        End If

        Dim strNo As String = CStr(Me.lstPIC.SelectedValue)
        Dim dr As DataRow() = dtPic.Select("PERSON_IN_CHARGE_NO='" & strNo & "'")

        For Each r As DataRow In dr
            dtOther.ImportRow(r)
            dtPic.Rows.Remove(r)
        Next
        dtPic.AcceptChanges()

        Me.lstOther.DataSource = dtOther
        Me.lstOther.ValueMember = "PERSON_IN_CHARGE_NO"
        Me.lstOther.DisplayMember = "PIC_DISPLAY"

        Me.lstPIC.DataSource = dtPic
        Me.lstPIC.ValueMember = "PERSON_IN_CHARGE_NO"
        Me.lstPIC.DisplayMember = "PIC_DISPLAY"

    End Sub

#End Region

End Class