<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmShowSymbol
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
        Me.lbxSymbol = New System.Windows.Forms.ListBox()
        Me.txtSymbol = New System.Windows.Forms.TextBox()
        Me.cmdSymbol = New System.Windows.Forms.Button()
        Me.dgvFieldCodes = New System.Windows.Forms.DataGridView()
        Me.lblSymbol = New System.Windows.Forms.Label()
        CType(Me.dgvFieldCodes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lbxSymbol
        '
        Me.lbxSymbol.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lbxSymbol.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbxSymbol.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbxSymbol.ItemHeight = 19
        Me.lbxSymbol.Location = New System.Drawing.Point(24, 101)
        Me.lbxSymbol.Name = "lbxSymbol"
        Me.lbxSymbol.Size = New System.Drawing.Size(63, 401)
        Me.lbxSymbol.TabIndex = 93
        '
        'txtSymbol
        '
        Me.txtSymbol.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSymbol.Location = New System.Drawing.Point(23, 68)
        Me.txtSymbol.Name = "txtSymbol"
        Me.txtSymbol.Size = New System.Drawing.Size(64, 29)
        Me.txtSymbol.TabIndex = 102
        Me.txtSymbol.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'cmdSymbol
        '
        Me.cmdSymbol.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSymbol.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdSymbol.Location = New System.Drawing.Point(16, 4)
        Me.cmdSymbol.Name = "cmdSymbol"
        Me.cmdSymbol.Size = New System.Drawing.Size(79, 61)
        Me.cmdSymbol.TabIndex = 144
        Me.cmdSymbol.Text = "Hide Symbol Copy"
        Me.cmdSymbol.UseVisualStyleBackColor = True
        '
        'dgvFieldCodes
        '
        Me.dgvFieldCodes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvFieldCodes.Location = New System.Drawing.Point(50, 490)
        Me.dgvFieldCodes.Name = "dgvFieldCodes"
        Me.dgvFieldCodes.Size = New System.Drawing.Size(69, 51)
        Me.dgvFieldCodes.TabIndex = 143
        Me.dgvFieldCodes.Visible = False
        '
        'lblSymbol
        '
        Me.lblSymbol.Location = New System.Drawing.Point(13, 509)
        Me.lblSymbol.Name = "lblSymbol"
        Me.lblSymbol.Size = New System.Drawing.Size(85, 200)
        Me.lblSymbol.TabIndex = 145
        Me.lblSymbol.Text = "(nbh = nonbreaking hyphen" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "nbsp = nonbreaking space)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'frmShowSymbol
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(116, 718)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblSymbol)
        Me.Controls.Add(Me.lbxSymbol)
        Me.Controls.Add(Me.txtSymbol)
        Me.Controls.Add(Me.dgvFieldCodes)
        Me.Controls.Add(Me.cmdSymbol)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "frmShowSymbol"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Symbols..."
        Me.TopMost = True
        CType(Me.dgvFieldCodes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lbxSymbol As System.Windows.Forms.ListBox
    Friend WithEvents txtSymbol As System.Windows.Forms.TextBox
    Friend WithEvents cmdSymbol As System.Windows.Forms.Button
    Friend WithEvents dgvFieldCodes As System.Windows.Forms.DataGridView
    Friend WithEvents lblSymbol As System.Windows.Forms.Label
End Class
