<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Calculation
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Calculation))
        Me.pbLoading = New System.Windows.Forms.PictureBox()
        Me.lblImporting = New System.Windows.Forms.Label()
        CType(Me.pbLoading, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pbLoading
        '
        Me.pbLoading.Image = CType(resources.GetObject("pbLoading.Image"), System.Drawing.Image)
        Me.pbLoading.Location = New System.Drawing.Point(0, 0)
        Me.pbLoading.Margin = New System.Windows.Forms.Padding(0)
        Me.pbLoading.Name = "pbLoading"
        Me.pbLoading.Size = New System.Drawing.Size(220, 220)
        Me.pbLoading.TabIndex = 0
        Me.pbLoading.TabStop = False
        '
        'lblImporting
        '
        Me.lblImporting.AutoSize = True
        Me.lblImporting.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblImporting.ForeColor = System.Drawing.SystemColors.ControlDark
        Me.lblImporting.Location = New System.Drawing.Point(45, 221)
        Me.lblImporting.Name = "lblImporting"
        Me.lblImporting.Size = New System.Drawing.Size(128, 31)
        Me.lblImporting.TabIndex = 1
        Me.lblImporting.Text = "Calculate"
        '
        'Calculation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(220, 260)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblImporting)
        Me.Controls.Add(Me.pbLoading)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Calculation"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Loading"
        Me.TopMost = True
        CType(Me.pbLoading, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents pbLoading As PictureBox
    Friend WithEvents lblImporting As Label
End Class
