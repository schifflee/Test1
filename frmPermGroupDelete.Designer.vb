<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPermGroupDelete
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPermGroupDelete))
        Me.dgvPerm = New System.Windows.Forms.DataGridView()
        Me.lbl1 = New System.Windows.Forms.Label()
        Me.lbl2 = New System.Windows.Forms.Label()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdClipboard = New System.Windows.Forms.Button()
        CType(Me.dgvPerm, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvPerm
        '
        Me.dgvPerm.BackgroundColor = System.Drawing.Color.White
        Me.dgvPerm.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvPerm.Location = New System.Drawing.Point(24, 102)
        Me.dgvPerm.Name = "dgvPerm"
        Me.dgvPerm.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvPerm.Size = New System.Drawing.Size(465, 149)
        Me.dgvPerm.TabIndex = 0
        '
        'lbl1
        '
        Me.lbl1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl1.ForeColor = System.Drawing.Color.Blue
        Me.lbl1.Location = New System.Drawing.Point(21, 9)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(468, 53)
        Me.lbl1.TabIndex = 1
        Me.lbl1.Text = "Label1"
        Me.lbl1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lbl2
        '
        Me.lbl2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl2.ForeColor = System.Drawing.Color.Blue
        Me.lbl2.Location = New System.Drawing.Point(21, 266)
        Me.lbl2.Name = "lbl2"
        Me.lbl2.Size = New System.Drawing.Size(468, 41)
        Me.lbl2.TabIndex = 2
        Me.lbl2.Text = "Label1"
        Me.lbl2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOK.CausesValidation = False
        Me.cmdOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.Blue
        Me.cmdOK.Location = New System.Drawing.Point(221, 310)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(65, 25)
        Me.cmdOK.TabIndex = 91
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'cmdClipboard
        '
        Me.cmdClipboard.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdClipboard.CausesValidation = False
        Me.cmdClipboard.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClipboard.ForeColor = System.Drawing.Color.Blue
        Me.cmdClipboard.Location = New System.Drawing.Point(325, 71)
        Me.cmdClipboard.Name = "cmdClipboard"
        Me.cmdClipboard.Size = New System.Drawing.Size(164, 25)
        Me.cmdClipboard.TabIndex = 92
        Me.cmdClipboard.Text = "&Copy to Clipboard"
        Me.cmdClipboard.UseVisualStyleBackColor = False
        '
        'frmPermGroupDelete
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(525, 352)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmdClipboard)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.lbl2)
        Me.Controls.Add(Me.lbl1)
        Me.Controls.Add(Me.dgvPerm)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmPermGroupDelete"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Permissions Group cannot be deleted..."
        CType(Me.dgvPerm, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvPerm As System.Windows.Forms.DataGridView
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Friend WithEvents lbl2 As System.Windows.Forms.Label
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdClipboard As System.Windows.Forms.Button
End Class
