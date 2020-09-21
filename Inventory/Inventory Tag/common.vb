Imports System.Security.Cryptography
Imports System.text
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports Inventory_Tag.FrmInvTag

Public Class Common
#Region "Public Table & View"
    Public Shared TB_Type As New DataTable
    Public Shared Vwtype As New DataView

    Public Shared TB_Code As New DataTable
    Public Shared VwCode As New DataView

    Public Shared TB_Loc As New DataTable
    Public Shared VwLoc As New DataView
#End Region

    Public Shared Sub GetBrand()
        Dim strsql As String = String.Empty
        Dim sb As New StringBuilder()
        Dim c1 As New SQLData("ACCINV")
        sb.AppendLine(" SELECT   *  ")
        sb.AppendLine(" FROM  TblGroup")
        sb.AppendLine(" UNION")
        sb.AppendLine(" SELECT '03' Typecode,Finalcompound code")
        sb.AppendLine(" FROM TBLCompound")
        sb.AppendLine(" WHERE active = 1")
        sb.AppendLine(" ORDER BY Typecode,code")

        strsql = sb.ToString()
        TB_Code = c1.GetDataset(strsql).Tables(0)
        VwCode.Table = TB_Code
        VwCode.Sort = "code"
    End Sub

    Public Shared Sub GetTypeinv()
        Dim strsql As String = String.Empty
        Dim c1 As New SQLData("ACCINV")
        strsql += "SELECT * FROM TBLTYPE "
        TB_Type = c1.GetDataset(strsql).Tables(0)
        Vwtype.Table = TB_Type
        Vwtype.Sort = "TypeCode"
    End Sub

    Public Shared Sub GetLocation()
        Dim strsql As String = String.Empty
        Dim c1 As New SQLData("ACCINV")
        strsql += "SELECT * FROM TBLDepartment "
        TB_Loc = c1.GetDataset(strsql).Tables(0)
        VwLoc.Table = TB_Loc
        VwLoc.Sort = "DeptCode"
    End Sub

End Class
#Region "Key for RSA - Do not Edit"

Public Class RSAKEY
    Public Const KPublic As String = "<RSAKeyValue><Modulus>tux6I1dMhAcdNT8NHs1l5PhhR9RV9EZRaQh0n4ojvQYpyfOQQt6M1N1TaUSIklqvNkJPixqxoE9pdhVfUvaU0yEN7vy5moQoDDEo0pd1dKtTkWInSNYMafGe7QFke0EMVFif9qJnhyU+aXyjNWLjSMCpqx+WHo0eM8VxWKIFb6c=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>"
    Public Const KPrivate As String = "<RSAKeyValue><Modulus>tux6I1dMhAcdNT8NHs1l5PhhR9RV9EZRaQh0n4ojvQYpyfOQQt6M1N1TaUSIklqvNkJPixqxoE9pdhVfUvaU0yEN7vy5moQoDDEo0pd1dKtTkWInSNYMafGe7QFke0EMVFif9qJnhyU+aXyjNWLjSMCpqx+WHo0eM8VxWKIFb6c=</Modulus><Exponent>AQAB</Exponent><P>8QK8H248+vT3Ie7Nj8MBV9/vNpRTeMsT1AhaRWBalL6rZxoaJIiWGrTqC1/7M86Usw6ZYO+PEq2MIvvR5EIgbQ==</P><Q>wkzoXr4Q+ze6e+az1XGP+QzYfBXTCq/1Rach3Kdp4Sp7cMbeL5J3hWmnzeUTl/mBHZyMqK+787yof/5JoTYL4w==</Q><DP>nHNpCYI3RbWlg7qQaGVvRssQbz7EHOK/QWIWr3iH9Iz9mVVBaTvdLQMJ905cNFpC/yVX/awlFTvhf4g2zVT71Q==</DP><DQ>Ot70lSg/mu5yuXHYUTa8abiDq20taZKQ3U7birDK+udVSYFn9sAJKMovhsn+2tBFV8SENeQxLZOe9lEE3Cy1Aw==</DQ><InverseQ>T20e2fOjB1dydOjlRO9WJGs+25JMw6d4G7H9uXnHa80kY/8HLvLOXRUIia/MsW17sFwiGvIoflkxlN2Bmsv4ug==</InverseQ><D>qpWEoQh8NnNb7ZfK6HqrFwf50D5XmeEpckWMXGs6QMBKoCYe1f0sYCW172kV40XmNzdHbnWKR/FGa/QqXPfOeC9bUaUKp+WQDTtuaz3t/RcE2/hnPWa6ZBF7HUtu7pgpSfnSIt1pu7lFpJaY6mvn1WyxcfK1pzTrDnZZtAe4RKE=</D></RSAKeyValue>"
End Class

#End Region

Public Class DataGridColoredTextBoxColumn
    Inherits DataGridTextBoxColumn
    'Fields
    'Constructors
    'Events
    'Methods
    Private column As Integer ' column where this columnstyle is located...
    Public Sub New()
        ' nothing
    End Sub
    Protected Overloads Overrides Sub Paint(ByVal g As Graphics, ByVal bounds As Rectangle, ByVal source As CurrencyManager, ByVal rowNum As Integer, ByVal backBrush As Brush, ByVal foreBrush As Brush, ByVal alignToRight As Boolean)
        Try
            Dim grid As DataGrid = Me.DataGridTableStyle.DataGrid
            Dim bT0 As Boolean
            Dim bT1 As Boolean
            Dim bT2 As Boolean
            'first time set the column properly
            If column = -2 Then
                Dim i As Integer
                i = Me.DataGridTableStyle.GridColumnStyles.IndexOf(Me)
                If i > -1 Then
                    column = i
                End If

                'If CType(CType(CType(source.Current, Object), System.Data.DataRowView).Row, System.Data.DataRow).ItemArray(21) = True Then
                '    backBrush = New LinearGradientBrush(bounds, Color.Red, Color.Green, LinearGradientMode.BackwardDiagonal)
                '    foreBrush = New SolidBrush(Color.Blue)
                'End If
                'Else
                '    If column = -9 Then
                '        backBrush = New LinearGradientBrush(bounds, Color.Red, Color.Green, LinearGradientMode.BackwardDiagonal)
                '        foreBrush = New SolidBrush(Color.Blue)
                '    End If
                bT0 = CType(CType(CType(source.Current, Object), System.Data.DataRowView).Row, System.Data.DataRow).ItemArray(20)
                bT1 = CType(CType(CType(source.Current, Object), System.Data.DataRowView).Row, System.Data.DataRow).ItemArray(21)
                bT2 = CType(CType(CType(source.Current, Object), System.Data.DataRowView).Row, System.Data.DataRow).ItemArray(22)
                If bT0 Or bT1 Or bT2 Then
                    If source.Position = rowNum Then
                        backBrush = New LinearGradientBrush(bounds, Color.AntiqueWhite, Color.FloralWhite, LinearGradientMode.BackwardDiagonal)
                        foreBrush = New SolidBrush(Color.Blue)
                    End If
                Else
                    'Color.FromArgb(255, 200, 200) Color.FromArgb(128, 20, 20)
                    If grid.CurrentRowIndex = rowNum And grid.CurrentCell.ColumnNumber = column Then
                        backBrush = New LinearGradientBrush(bounds, Color.White, Color.YellowGreen, LinearGradientMode.Vertical)
                        foreBrush = New SolidBrush(Color.Blue)
                    End If
                End If
            End If
        Catch ex As Exception
            ' empty catch 
        Finally
            ' make sure the base class gets called to do the drawing with
            ' the possibly changed brushes
            MyBase.Paint(g, bounds, source, rowNum, backBrush, foreBrush, alignToRight)
        End Try

    End Sub

    Protected Overloads Overrides Sub Edit(ByVal source As System.Windows.Forms.CurrencyManager, ByVal rowNum As Integer, ByVal bounds As System.Drawing.Rectangle, ByVal [readOnly] As Boolean, ByVal instantText As String, ByVal cellIsVisible As Boolean)
        'do nothing...
    End Sub


End Class

Public Class DataGridColoredLine2
    Inherits DataGridTextBoxColumn
    'Fields
    'Constructors
    'Events
    'Methods
    Private column As Integer ' column where this columnstyle is located...
    Public Sub New()
        ' nothing
    End Sub
    Protected Overloads Overrides Sub Paint(ByVal g As Graphics, ByVal bounds As Rectangle, ByVal source As CurrencyManager, ByVal rowNum As Integer, ByVal backBrush As Brush, ByVal foreBrush As Brush, ByVal alignToRight As Boolean)
        Try
            Dim grid As DataGrid = Me.DataGridTableStyle.DataGrid
            If rowNum = grid.CurrentRowIndex Then
                backBrush = New LinearGradientBrush(bounds, Color.LightYellow, Color.Violet, LinearGradientMode.Vertical)
                foreBrush = New SolidBrush(Color.Black)
            Else
                foreBrush = New SolidBrush(Color.Black)
            End If

            'foreBrush = New SolidBrush(Color.Black)
            'Dim dblQty As Double
            'dblQty = CType(CType(CType(grid.DataSource, Object), System.Data.DataView).Table, System.Data.DataTable).Rows(rowNum).Item(column)
            'foreBrush = New SolidBrush(Color.Green)
        Catch ex As Exception
            ' empty catch 
        Finally
            ' make sure the base class gets called to do the drawing with
            ' the possibly changed brushes
            MyBase.Paint(g, bounds, source, rowNum, backBrush, foreBrush, alignToRight)
        End Try
    End Sub
    Protected Overloads Overrides Sub Edit(ByVal source As System.Windows.Forms.CurrencyManager, ByVal rowNum As Integer, ByVal bounds As System.Drawing.Rectangle, ByVal [readOnly] As Boolean, ByVal instantText As String, ByVal cellIsVisible As Boolean)
        'do nothing...
    End Sub
End Class

Public Delegate Function CheckRowHeader(ByVal row As Integer) As Boolean
Public Class DataGridQtyBox
    Inherits DataGridTextBoxColumn
    Public WithEvents TBox As TextBox
    Private WithEvents _source As CurrencyManager
    Private _rowNum As Integer
    Private _isEditing As Boolean
    Private _getRowEntry As CheckRowHeader

    Shared Sub New()
        'Warning: Implementation not found
    End Sub
    Public Sub New(ByVal cre As CheckRowHeader)
        MyBase.New()
        _getRowEntry = cre
        _source = Nothing
        _isEditing = False
        TBox = New TextBox
        AddHandler TBox.Leave, New EventHandler(AddressOf LeaveTBox)
        AddHandler TBox.TextChanged, New EventHandler(AddressOf TboxStartEditing)
    End Sub

    Protected Overloads Overrides Sub Edit(ByVal source As CurrencyManager, ByVal rowNum As Integer, ByVal bounds As Rectangle, ByVal readOnly1 As Boolean, ByVal instantText As String, ByVal cellIsVisible As Boolean)
        If _getRowEntry(rowNum) Then
        Else
            MyBase.Edit(source, rowNum, bounds, readOnly1, instantText, cellIsVisible)
            _rowNum = rowNum
            _source = source
            TBox.Parent = Me.TextBox.Parent
            TBox.Location = Me.TextBox.Location
            TBox.Size = New Size(Me.TextBox.Size.Width - 1, TBox.Size.Height - 1)
            TBox.Location = Me.TextBox.Location
            TBox.Text = Me.TextBox.Text
            Me.TextBox.Visible = False
            TBox.Visible = True
            TBox.BackColor = Color.White
            TBox.ForeColor = Color.Black

            TBox.BringToFront()
            TBox.Focus()
        End If
    End Sub
    Protected Overloads Overrides Function Commit(ByVal dataSource As CurrencyManager, ByVal rowNum As Integer) As Boolean
        If _getRowEntry(rowNum) Then
        Else
            If _isEditing Then
                'HasEdited = True
                'If rowNum <> 3 Then
                Try
                    SetColumnValueAtRow(dataSource, rowNum, CDbl(IIf(TBox.Text = "", "0", TBox.Text)))
                Catch ex As Exception
                    MsgBox("ผิดพลาด Commit")
                End Try

                'Else
                '    SendKeys.Send("{TAB}")
                'End If
                TBox.Hide()
                Exit Function
            End If
        End If
        Return True
    End Function
    Private Sub TboxStartEditing(ByVal sender As Object, ByVal e As EventArgs)
        If _getRowEntry(_rowNum) Then
            _isEditing = False
        Else
            _isEditing = True
            MyBase.ColumnStartedEditing(sender)
        End If
    End Sub
    Private Sub LeaveTBox(ByVal sender As Object, ByVal e As EventArgs)
        If _getRowEntry(_rowNum) Then
            _source.Position = _rowNum
        Else
            Try
                SetColumnValueAtRow(_source, _rowNum, CDbl(IIf(TBox.Text = "", "0", TBox.Text)))
                'HasEdited = True
            Catch
            End Try
            _isEditing = False
        End If
        TBox.Hide()
    End Sub
    Private Sub TBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TBox.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8 ' back space
            Case 13
                e.Handled = True

                If _getRowEntry(_rowNum + 1) Then
                    SendKeys.Send("{TAB}")
                Else
                    SendKeys.Send("{DOWN}")
                End If
            Case 43 '+
                e.Handled = True

                SendKeys.Send("{DOWN}")
            Case 45 '-
                e.Handled = True

                SendKeys.Send("{UP}")
            Case 46
                If InStr(sender.text, ".") <> 0 Then
                    e.Handled = True
                End If

            Case Else
                'If Not HasAut2Edit Then
                '    e.Handled = True
                '    Exit Sub
                'End If

                Dim a As Integer = InStr(sender.text, ".")
                If a <> 0 And a = Len(sender.text.Trim) - 3 Then
                    If Len(sender.text.trim) <> sender.SelectionLength Then
                        e.Handled = True
                        Exit Sub
                    End If

                End If

                If Not IsNumeric(e.KeyChar) Then
                    e.Handled = True
                Else
                    If Len(sender.text) >= 8 Then
                        If Len(sender.text) = 8 Then
                            If CDbl(sender.text + e.KeyChar) > 999999 Then
                                e.Handled = True
                            End If
                        Else
                            e.Handled = True
                        End If
                    End If
                End If


        End Select
    End Sub


End Class

Public Delegate Function CheckRow(ByVal row As String) As Boolean
Public Class DataGridUnitBox
    Inherits DataGridTextBoxColumn
    Public WithEvents TBox As TextBox
    Private WithEvents _source As CurrencyManager
    Private _rowNum As Integer
    Private _isEditing As Boolean
    Private _getRowEntry As CheckRow

    Shared Sub New()
        'Warning: Implementation not found
    End Sub
    Public Sub New(ByVal cre As CheckRow)
        MyBase.New()
        _getRowEntry = cre
        _source = Nothing
        _isEditing = False
        TBox = New TextBox
        AddHandler TBox.Leave, New EventHandler(AddressOf LeaveTBox)
        AddHandler TBox.TextChanged, New EventHandler(AddressOf TboxStartEditing)
    End Sub

    Protected Overloads Overrides Sub Edit(ByVal source As CurrencyManager, ByVal rowNum As Integer, ByVal bounds As Rectangle, ByVal readOnly1 As Boolean, ByVal instantText As String, ByVal cellIsVisible As Boolean)
        If _getRowEntry(rowNum) Then
        Else
            MyBase.Edit(source, rowNum, bounds, readOnly1, instantText, cellIsVisible)
            _rowNum = rowNum
            _source = source
            TBox.Parent = Me.TextBox.Parent
            TBox.Location = Me.TextBox.Location
            TBox.Size = New Size(Me.TextBox.Size.Width - 1, TBox.Size.Height - 1)
            TBox.Location = Me.TextBox.Location
            TBox.Text = Me.TextBox.Text
            Me.TextBox.Visible = False
            TBox.Visible = True
            TBox.BackColor = Color.White
            TBox.ForeColor = Color.Black

            TBox.BringToFront()
            TBox.Focus()
        End If
    End Sub
    Protected Overloads Overrides Function Commit(ByVal dataSource As CurrencyManager, ByVal rowNum As Integer) As Boolean
        If _getRowEntry(rowNum) Then
        Else
            If _isEditing Then
                'HasEdited = True
                'If rowNum <> 3 Then
                Try
                    SetColumnValueAtRow(dataSource, rowNum, CStr(IIf(TBox.Text = "", "0", TBox.Text)))
                Catch ex As Exception
                    MsgBox("ผิดพลาด Commit")
                End Try

                'Else
                '    SendKeys.Send("{TAB}")
                'End If
                TBox.Hide()
                Exit Function
            End If
        End If
        Return True
    End Function
    Private Sub TboxStartEditing(ByVal sender As Object, ByVal e As EventArgs)
        If _getRowEntry(_rowNum) Then
            _isEditing = False
        Else
            _isEditing = True
            MyBase.ColumnStartedEditing(sender)
        End If
    End Sub
    Private Sub LeaveTBox(ByVal sender As Object, ByVal e As EventArgs)
        If _getRowEntry(_rowNum) Then
            _source.Position = _rowNum
        Else
            Try
                SetColumnValueAtRow(_source, _rowNum, CStr(IIf(TBox.Text = "", "0", TBox.Text)))
                'HasEdited = True
            Catch
            End Try
            _isEditing = False
        End If
        TBox.Hide()
    End Sub
    Private Sub TBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TBox.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8 ' back space
            Case 13
                e.Handled = True

                If _getRowEntry(_rowNum + 1) Then
                    SendKeys.Send("{TAB}")
                Else
                    SendKeys.Send("{DOWN}")
                End If
            Case 43 '+
                e.Handled = True

                SendKeys.Send("{DOWN}")
            Case 45 '-
                e.Handled = True

                SendKeys.Send("{UP}")
            Case 46
                'If InStr(sender.text, ".") <> 0 Then
                '    e.Handled = True
                'End If

            Case Else
                'If Not HasAut2Edit Then
                '    e.Handled = True
                '    Exit Sub
                'End If

                'Dim a As Integer = InStr(sender.text, ".")
                'If a <> 0 And a = Len(sender.text.Trim) - 3 Then
                '    If Len(sender.text.trim) <> sender.SelectionLength Then
                '        e.Handled = True
                '        Exit Sub
                '    End If

                'End If

                'If Not IsNumeric(e.KeyChar) Then
                '    e.Handled = True
                'Else
                '    If Len(sender.text) >= 8 Then
                '        If Len(sender.text) = 8 Then
                '            If CDbl(sender.text + e.KeyChar) > 999999 Then
                '                e.Handled = True
                '            End If
                '        Else
                '            e.Handled = True
                '        End If
                '    End If
                'End If


        End Select
    End Sub


End Class

Public Delegate Function delegateGetRowSeq(ByVal row As Integer) As CellColor
Public Class DataGridColoredParent
    Inherits DataGridTextBoxColumn
    'Fields
    'Constructors
    'Events
    'Methods
    Private _GetRowSeq As delegateGetRowSeq
    Private column As Integer ' column where this columnstyle is located...
    Private adjCol As Integer
    Public Sub New()
        ' nothing
    End Sub
    Public Sub New(ByVal GetRowSeq As delegateGetRowSeq)
        _GetRowSeq = GetRowSeq
    End Sub
    'Public Sub New(ByVal adj As Integer)
    '    ' nothing
    '    adjCol = adj
    'End Sub
    Protected Overloads Overrides Sub Paint(ByVal g As Graphics, ByVal bounds As Rectangle, ByVal source As CurrencyManager, ByVal rowNum As Integer, ByVal backBrush As Brush, ByVal foreBrush As Brush, ByVal alignToRight As Boolean)
        Try
            Dim c As CellColor
            c = _GetRowSeq(rowNum)
            Dim grid As DataGrid = Me.DataGridTableStyle.DataGrid
            If rowNum = grid.CurrentRowIndex Then
                'backBrush = New LinearGradientBrush(bounds, Color.Gold, Color.Orange, LinearGradientMode.Vertical)
                backBrush = SelColorBrush(c.LfItem, bounds)
            Else

            End If

            If c.ForeG = 0 Then
                foreBrush = New SolidBrush(Color.Black)
                If rowNum = grid.CurrentRowIndex Then
                    'backBrush = New LinearGradientBrush(bounds, Color.Gold, Color.Orange, LinearGradientMode.Vertical)
                Else
                    If Not HasColor Then
                        'backBrush = New LinearGradientBrush(bounds, Color.Azure, Color.LightCyan, LinearGradientMode.Vertical)
                        If c.BackG = 0 Then
                            'backBrush = New LinearGradientBrush(bounds, Color.LightYellow, Color.YellowGreen, LinearGradientMode.Vertical)
                            backBrush = New LinearGradientBrush(bounds, Color.AliceBlue, Color.Lavender, LinearGradientMode.Vertical)
                        Else
                            'backBrush = New LinearGradientBrush(bounds, Color.LavenderBlush, Color.Plum, LinearGradientMode.Vertical)
                            backBrush = New LinearGradientBrush(bounds, Color.Linen, Color.MistyRose, LinearGradientMode.Vertical)
                        End If
                    Else
                        backBrush = SelColorBrush(c.LfItem, bounds)
                    End If
                End If
            Else
                foreBrush = New SolidBrush(Color.Transparent)
            End If
            'If 1 = 1 Then 
            '    foreBrush = New SolidBrush(Color.Black) 
            'Else 
            '    foreBrush = New SolidBrush(Color.Transparent) 
            'End If 


        Catch ex As Exception
            ' empty catch 
            'MsgBox(ex.Message)
        Finally
            ' make sure the base class gets called to do the drawing with
            ' the possibly changed brushes
            MyBase.Paint(g, bounds, source, rowNum, backBrush, foreBrush, alignToRight)
        End Try
    End Sub
    Protected Function SelColorBrush(ByVal strPrefixItem As String, ByRef bounds As Rectangle) As Brush
        Dim backBrush As Brush
        If Not HasColor Then
            backBrush = New LinearGradientBrush(bounds, Color.LightYellow, Color.GreenYellow, LinearGradientMode.Vertical)
        Else
            Select Case strPrefixItem
                Case "0101"
                    backBrush = New LinearGradientBrush(bounds, Color.SeaShell, Color.Tomato, LinearGradientMode.Vertical)
                Case "0102"
                    backBrush = New LinearGradientBrush(bounds, Color.Moccasin, Color.Orange, LinearGradientMode.Vertical)
                Case "0103"
                    backBrush = New LinearGradientBrush(bounds, Color.LavenderBlush, Color.DeepPink, LinearGradientMode.Vertical)
                Case "0104"
                    backBrush = New LinearGradientBrush(bounds, Color.LavenderBlush, Color.Plum, LinearGradientMode.Vertical)
                Case "0105"
                    backBrush = New LinearGradientBrush(bounds, Color.LightYellow, Color.YellowGreen, LinearGradientMode.Vertical)
                Case "0106"
                    backBrush = New LinearGradientBrush(bounds, Color.MistyRose, Color.Red, LinearGradientMode.Vertical)
                    ' ไลท์ 
                Case "0209"
                    backBrush = New LinearGradientBrush(bounds, Color.White, Color.SkyBlue, LinearGradientMode.Vertical)
                Case "0211"
                    backBrush = New LinearGradientBrush(bounds, Color.PaleGoldenrod, Color.PaleGreen, LinearGradientMode.Vertical)
                Case "0409" 'ธรรมชาติ            
                    backBrush = New LinearGradientBrush(bounds, Color.PowderBlue, Color.RoyalBlue, LinearGradientMode.Vertical)
                Case "0451" 'สตรอเบอรี่          
                    backBrush = New LinearGradientBrush(bounds, Color.Snow, Color.DeepPink, LinearGradientMode.Vertical)
                Case "0454" 'วุ้นมะพร้าว         
                    backBrush = New LinearGradientBrush(bounds, Color.GreenYellow, Color.MediumSeaGreen, LinearGradientMode.Vertical)
                Case "0456" 'ผลไม้รวม
                    backBrush = New LinearGradientBrush(bounds, Color.Orange, Color.OrangeRed, LinearGradientMode.Vertical)
                Case "0457" ' เผือก               
                    backBrush = New LinearGradientBrush(bounds, Color.Gainsboro, Color.Orchid, LinearGradientMode.Vertical)
                Case "0458" ' ธัญญาหาร                        
                    backBrush = New LinearGradientBrush(bounds, Color.Yellow, Color.GreenYellow, LinearGradientMode.Vertical)
                Case "รวม " ' Total Line         
                    backBrush = New LinearGradientBrush(bounds, Color.AliceBlue, Color.DodgerBlue, LinearGradientMode.Vertical)
                Case Else
                    backBrush = New LinearGradientBrush(bounds, Color.White, Color.PowderBlue, LinearGradientMode.Vertical)
            End Select
        End If
        Return backBrush
    End Function
    Protected Overloads Overrides Sub Edit(ByVal source As System.Windows.Forms.CurrencyManager, ByVal rowNum As Integer, ByVal bounds As System.Drawing.Rectangle, ByVal [readOnly] As Boolean, ByVal instantText As String, ByVal cellIsVisible As Boolean)
        'do nothing...
    End Sub
End Class

Public Delegate Function delegateNegValue(ByVal row As Integer) As Boolean
Public Class DataGridColoredLine
    Inherits DataGridTextBoxColumn
    'Fields
    'Constructors
    'Events
    'Methods
    Private _GetColNeg As delegateNegValue
    Private column As Integer ' column where this columnstyle is located...
    Public Sub New()
        ' nothing
    End Sub
    Public Sub New(ByVal GetColNeg As delegateNegValue)
        _GetColNeg = GetColNeg
    End Sub
    Protected Overloads Overrides Sub Paint(ByVal g As Graphics, ByVal bounds As Rectangle, ByVal source As CurrencyManager, ByVal rowNum As Integer, ByVal backBrush As Brush, ByVal foreBrush As Brush, ByVal alignToRight As Boolean)
        Try
            Dim grid As DataGrid = Me.DataGridTableStyle.DataGrid
            If rowNum = grid.CurrentRowIndex Then
                backBrush = New LinearGradientBrush(bounds, Color.LightYellow, Color.GreenYellow, LinearGradientMode.Vertical)
                foreBrush = New SolidBrush(Color.LightYellow)
            End If
            foreBrush = New SolidBrush(Color.Black)
            ' เช็คค่าติดลบให้เป็นตัวหนังสือสีแดง
            'Dim Neg As Boolean
            'Neg = _GetColNeg(rowNum)
            'Dim dblQty As Double
            'dblQty = CType(CType(CType(grid.DataSource, Object), System.Data.DataView).Table, System.Data.DataTable).Rows(rowNum).Item(column)
            'If Neg Then
            'foreBrush = New SolidBrush(Color.Red)
            'Else
            '    foreBrush = New SolidBrush(Color.Green)
            'End If
        Catch ex As Exception
            ' empty catch 
        Finally
            ' make sure the base class gets called to do the drawing with
            ' the possibly changed brushes
            MyBase.Paint(g, bounds, source, rowNum, backBrush, foreBrush, alignToRight)
        End Try
    End Sub
    Protected Overloads Overrides Sub Edit(ByVal source As System.Windows.Forms.CurrencyManager, ByVal rowNum As Integer, ByVal bounds As System.Drawing.Rectangle, ByVal [readOnly] As Boolean, ByVal instantText As String, ByVal cellIsVisible As Boolean)
        'do nothing...
    End Sub


End Class




