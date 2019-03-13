<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConfigSDCmpds
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmConfigSDCmpds))
        Me.dgvCmpd = New System.Windows.Forms.DataGridView()
        Me.dgvSource = New System.Windows.Forms.DataGridView()
        Me.pan1 = New System.Windows.Forms.Panel()
        Me.cmdCancel1 = New System.Windows.Forms.Button()
        Me.cmdOK1 = New System.Windows.Forms.Button()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.cmdRemove = New System.Windows.Forms.Button()
        Me.lblD = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.dgvCmpd, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pan1.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvCmpd
        '
        Me.dgvCmpd.BackgroundColor = System.Drawing.Color.White
        Me.dgvCmpd.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCmpd.Location = New System.Drawing.Point(24, 36)
        Me.dgvCmpd.Name = "dgvCmpd"
        Me.dgvCmpd.ReadOnly = True
        Me.dgvCmpd.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvCmpd.Size = New System.Drawing.Size(165, 249)
        Me.dgvCmpd.TabIndex = 151
        '
        'dgvSource
        '
        Me.dgvSource.BackgroundColor = System.Drawing.Color.White
        Me.dgvSource.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSource.Location = New System.Drawing.Point(278, 36)
        Me.dgvSource.Name = "dgvSource"
        Me.dgvSource.ReadOnly = True
        Me.dgvSource.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvSource.Size = New System.Drawing.Size(425, 249)
        Me.dgvSource.TabIndex = 152
        '
        'pan1
        '
        Me.pan1.Controls.Add(Me.cmdCancel1)
        Me.pan1.Controls.Add(Me.cmdOK1)
        Me.pan1.Location = New System.Drawing.Point(24, 291)
        Me.pan1.Name = "pan1"
        Me.pan1.Size = New System.Drawing.Size(182, 45)
        Me.pan1.TabIndex = 153
        '
        'cmdCancel1
        '
        Me.cmdCancel1.CausesValidation = False
        Me.cmdCancel1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel1.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel1.Location = New System.Drawing.Point(96, 10)
        Me.cmdCancel1.Name = "cmdCancel1"
        Me.cmdCancel1.Size = New System.Drawing.Size(80, 35)
        Me.cmdCancel1.TabIndex = 1
        Me.cmdCancel1.Text = "&Cancel"
        Me.cmdCancel1.UseVisualStyleBackColor = True
        '
        'cmdOK1
        '
        Me.cmdOK1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK1.ForeColor = System.Drawing.Color.Blue
        Me.cmdOK1.Location = New System.Drawing.Point(0, 10)
        Me.cmdOK1.Name = "cmdOK1"
        Me.cmdOK1.Size = New System.Drawing.Size(80, 35)
        Me.cmdOK1.TabIndex = 0
        Me.cmdOK1.Text = "&OK"
        Me.cmdOK1.UseVisualStyleBackColor = True
        '
        'cmdAdd
        '
        Me.cmdAdd.ForeColor = System.Drawing.Color.Blue
        Me.cmdAdd.Location = New System.Drawing.Point(195, 59)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(77, 27)
        Me.cmdAdd.TabIndex = 154
        Me.cmdAdd.Text = "<- Add"
        Me.cmdAdd.UseVisualStyleBackColor = True
        '
        'cmdRemove
        '
        Me.cmdRemove.ForeColor = System.Drawing.Color.Red
        Me.cmdRemove.Location = New System.Drawing.Point(195, 92)
        Me.cmdRemove.Name = "cmdRemove"
        Me.cmdRemove.Size = New System.Drawing.Size(77, 27)
        Me.cmdRemove.TabIndex = 155
        Me.cmdRemove.Text = "Remove ->"
        Me.cmdRemove.UseVisualStyleBackColor = True
        '
        'lblD
        '
        Me.lblD.AutoSize = True
        Me.lblD.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblD.Location = New System.Drawing.Point(21, 18)
        Me.lblD.Name = "lblD"
        Me.lblD.Size = New System.Drawing.Size(167, 15)
        Me.lblD.TabIndex = 156
        Me.lblD.Text = "Configured Compound(s)"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(275, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(237, 15)
        Me.Label1.TabIndex = 157
        Me.Label1.Text = "Source List of Available Compounds"
        '
        'frmConfigSDCmpds
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(730, 354)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblD)
        Me.Controls.Add(Me.cmdRemove)
        Me.Controls.Add(Me.cmdAdd)
        Me.Controls.Add(Me.pan1)
        Me.Controls.Add(Me.dgvSource)
        Me.Controls.Add(Me.dgvCmpd)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmConfigSDCmpds"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Configure Assay Compounds"
        CType(Me.dgvCmpd, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pan1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvCmpd As System.Windows.Forms.DataGridView
    Friend WithEvents dgvSource As System.Windows.Forms.DataGridView
    Friend WithEvents pan1 As System.Windows.Forms.Panel
    Friend WithEvents cmdCancel1 As System.Windows.Forms.Button
    Friend WithEvents cmdOK1 As System.Windows.Forms.Button
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents cmdRemove As System.Windows.Forms.Button
    Friend WithEvents lblD As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
