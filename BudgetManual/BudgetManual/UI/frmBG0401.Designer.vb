<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBG0401
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBG0401))
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.lblFormTitle = New System.Windows.Forms.Label
        Me.fraPIC = New System.Windows.Forms.GroupBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.cmdToOther = New System.Windows.Forms.Button
        Me.cmdToPIC = New System.Windows.Forms.Button
        Me.lstOther = New System.Windows.Forms.ListBox
        Me.lstPIC = New System.Windows.Forms.ListBox
        Me.fraPIC.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdSave
        '
        Me.cmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSave.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdSave.Location = New System.Drawing.Point(456, 316)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(75, 23)
        Me.cmdSave.TabIndex = 4
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'cmdClose
        '
        Me.cmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdClose.Location = New System.Drawing.Point(537, 316)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 23)
        Me.cmdClose.TabIndex = 5
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'lblFormTitle
        '
        Me.lblFormTitle.AutoSize = True
        Me.lblFormTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblFormTitle.Location = New System.Drawing.Point(12, 9)
        Me.lblFormTitle.Name = "lblFormTitle"
        Me.lblFormTitle.Size = New System.Drawing.Size(140, 24)
        Me.lblFormTitle.TabIndex = 11
        Me.lblFormTitle.Text = "Report Option"
        '
        'fraPIC
        '
        Me.fraPIC.Controls.Add(Me.Label2)
        Me.fraPIC.Controls.Add(Me.Label1)
        Me.fraPIC.Controls.Add(Me.cmdToOther)
        Me.fraPIC.Controls.Add(Me.cmdToPIC)
        Me.fraPIC.Controls.Add(Me.lstOther)
        Me.fraPIC.Controls.Add(Me.lstPIC)
        Me.fraPIC.Location = New System.Drawing.Point(12, 46)
        Me.fraPIC.Name = "fraPIC"
        Me.fraPIC.Size = New System.Drawing.Size(600, 264)
        Me.fraPIC.TabIndex = 12
        Me.fraPIC.TabStop = False
        Me.fraPIC.Text = "Summary by Person In Charge Report"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(21, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(92, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Person In Charge:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(323, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(36, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Other:"
        '
        'cmdToOther
        '
        Me.cmdToOther.Location = New System.Drawing.Point(280, 114)
        Me.cmdToOther.Name = "cmdToOther"
        Me.cmdToOther.Size = New System.Drawing.Size(40, 23)
        Me.cmdToOther.TabIndex = 2
        Me.cmdToOther.Text = ">>"
        Me.cmdToOther.UseVisualStyleBackColor = True
        '
        'cmdToPIC
        '
        Me.cmdToPIC.Location = New System.Drawing.Point(280, 143)
        Me.cmdToPIC.Name = "cmdToPIC"
        Me.cmdToPIC.Size = New System.Drawing.Size(40, 23)
        Me.cmdToPIC.TabIndex = 1
        Me.cmdToPIC.Text = "<<"
        Me.cmdToPIC.UseVisualStyleBackColor = True
        '
        'lstOther
        '
        Me.lstOther.FormattingEnabled = True
        Me.lstOther.Location = New System.Drawing.Point(326, 40)
        Me.lstOther.Name = "lstOther"
        Me.lstOther.Size = New System.Drawing.Size(250, 199)
        Me.lstOther.TabIndex = 3
        '
        'lstPIC
        '
        Me.lstPIC.FormattingEnabled = True
        Me.lstPIC.Location = New System.Drawing.Point(24, 40)
        Me.lstPIC.Name = "lstPIC"
        Me.lstPIC.Size = New System.Drawing.Size(250, 199)
        Me.lstPIC.TabIndex = 0
        '
        'frmBG0401
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(622, 350)
        Me.Controls.Add(Me.fraPIC)
        Me.Controls.Add(Me.lblFormTitle)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.cmdSave)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "frmBG0401"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "frmBG0401"
        Me.fraPIC.ResumeLayout(False)
        Me.fraPIC.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents lblFormTitle As System.Windows.Forms.Label
    Friend WithEvents fraPIC As System.Windows.Forms.GroupBox
    Friend WithEvents cmdToOther As System.Windows.Forms.Button
    Friend WithEvents cmdToPIC As System.Windows.Forms.Button
    Friend WithEvents lstOther As System.Windows.Forms.ListBox
    Friend WithEvents lstPIC As System.Windows.Forms.ListBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
