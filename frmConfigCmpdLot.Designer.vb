<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConfigCmpdLot
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmConfigCmpdLot))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblD = New System.Windows.Forms.Label()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.pan1 = New System.Windows.Forms.Panel()
        Me.cmdCancel1 = New System.Windows.Forms.Button()
        Me.cmdOK1 = New System.Windows.Forms.Button()
        Me.dgvSource = New System.Windows.Forms.DataGridView()
        Me.dgvLot = New System.Windows.Forms.DataGridView()
        Me.cmdRemove = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtCmpd = New System.Windows.Forms.TextBox()
        Me.lblNone = New System.Windows.Forms.Label()
        Me.pan1.SuspendLayout()
        CType(Me.dgvSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvLot, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(263, 36)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(261, 15)
        Me.Label1.TabIndex = 164
        Me.Label1.Text = "Source List of Available Compound Lots"
        '
        'lblD
        '
        Me.lblD.AutoSize = True
        Me.lblD.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblD.Location = New System.Drawing.Point(9, 36)
        Me.lblD.Name = "lblD"
        Me.lblD.Size = New System.Drawing.Size(101, 15)
        Me.lblD.TabIndex = 163
        Me.lblD.Text = "Configured Lot"
        '
        'cmdAdd
        '
        Me.cmdAdd.ForeColor = System.Drawing.Color.Blue
        Me.cmdAdd.Location = New System.Drawing.Point(183, 54)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(77, 27)
        Me.cmdAdd.TabIndex = 161
        Me.cmdAdd.Text = "<- Add"
        Me.cmdAdd.UseVisualStyleBackColor = True
        '
        'pan1
        '
        Me.pan1.Controls.Add(Me.cmdCancel1)
        Me.pan1.Controls.Add(Me.cmdOK1)
        Me.pan1.Location = New System.Drawing.Point(12, 286)
        Me.pan1.Name = "pan1"
        Me.pan1.Size = New System.Drawing.Size(182, 45)
        Me.pan1.TabIndex = 160
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
        'dgvSource
        '
        Me.dgvSource.BackgroundColor = System.Drawing.Color.White
        Me.dgvSource.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSource.Location = New System.Drawing.Point(266, 54)
        Me.dgvSource.MultiSelect = False
        Me.dgvSource.Name = "dgvSource"
        Me.dgvSource.ReadOnly = True
        Me.dgvSource.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvSource.Size = New System.Drawing.Size(437, 226)
        Me.dgvSource.TabIndex = 159
        '
        'dgvLot
        '
        Me.dgvLot.BackgroundColor = System.Drawing.Color.White
        Me.dgvLot.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvLot.Location = New System.Drawing.Point(12, 54)
        Me.dgvLot.Name = "dgvLot"
        Me.dgvLot.ReadOnly = True
        Me.dgvLot.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvLot.Size = New System.Drawing.Size(165, 226)
        Me.dgvLot.TabIndex = 158
        '
        'cmdRemove
        '
        Me.cmdRemove.ForeColor = System.Drawing.Color.Red
        Me.cmdRemove.Location = New System.Drawing.Point(183, 87)
        Me.cmdRemove.Name = "cmdRemove"
        Me.cmdRemove.Size = New System.Drawing.Size(77, 27)
        Me.cmdRemove.TabIndex = 162
        Me.cmdRemove.Text = "Remove ->"
        Me.cmdRemove.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(9, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(140, 15)
        Me.Label2.TabIndex = 165
        Me.Label2.Text = "Selected Compound:"
        '
        'txtCmpd
        '
        Me.txtCmpd.Enabled = False
        Me.txtCmpd.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCmpd.Location = New System.Drawing.Point(155, 6)
        Me.txtCmpd.Name = "txtCmpd"
        Me.txtCmpd.Size = New System.Drawing.Size(536, 22)
        Me.txtCmpd.TabIndex = 166
        '
        'lblNone
        '
        Me.lblNone.AutoSize = True
        Me.lblNone.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNone.ForeColor = System.Drawing.Color.Red
        Me.lblNone.Location = New System.Drawing.Point(263, 283)
        Me.lblNone.Name = "lblNone"
        Me.lblNone.Size = New System.Drawing.Size(341, 16)
        Me.lblNone.TabIndex = 167
        Me.lblNone.Text = "No Lots have been configured for this compound"
        '
        'frmConfigCmpdLot
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(715, 339)
        Me.Controls.Add(Me.lblNone)
        Me.Controls.Add(Me.txtCmpd)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblD)
        Me.Controls.Add(Me.cmdRemove)
        Me.Controls.Add(Me.cmdAdd)
        Me.Controls.Add(Me.pan1)
        Me.Controls.Add(Me.dgvSource)
        Me.Controls.Add(Me.dgvLot)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmConfigCmpdLot"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Configure Compound Lot Information"
        Me.pan1.ResumeLayout(False)
        CType(Me.dgvSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvLot, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblD As System.Windows.Forms.Label
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents pan1 As System.Windows.Forms.Panel
    Friend WithEvents cmdCancel1 As System.Windows.Forms.Button
    Friend WithEvents cmdOK1 As System.Windows.Forms.Button
    Friend WithEvents dgvSource As System.Windows.Forms.DataGridView
    Friend WithEvents dgvLot As System.Windows.Forms.DataGridView
    Friend WithEvents cmdRemove As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtCmpd As System.Windows.Forms.TextBox
    Friend WithEvents lblNone As System.Windows.Forms.Label
End Class
