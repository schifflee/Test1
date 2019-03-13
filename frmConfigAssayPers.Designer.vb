<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConfigAssayPers
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmConfigAssayPers))
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdRemovePI = New System.Windows.Forms.Button()
        Me.dgvPI = New System.Windows.Forms.DataGridView()
        Me.dgvSource = New System.Windows.Forms.DataGridView()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.lblPI = New System.Windows.Forms.Label()
        Me.cmdAddPI = New System.Windows.Forms.Button()
        Me.pan1 = New System.Windows.Forms.Panel()
        Me.dgvAnal = New System.Windows.Forms.DataGridView()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmdRemoveAnal = New System.Windows.Forms.Button()
        Me.cmdAddAnal = New System.Windows.Forms.Button()
        CType(Me.dgvPI, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pan1.SuspendLayout()
        CType(Me.dgvAnal, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdCancel
        '
        Me.cmdCancel.CausesValidation = False
        Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(96, 10)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(80, 35)
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(262, 42)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(226, 15)
        Me.Label1.TabIndex = 174
        Me.Label1.Text = "Source List of Available Personnel"
        '
        'cmdRemovePI
        '
        Me.cmdRemovePI.ForeColor = System.Drawing.Color.Red
        Me.cmdRemovePI.Location = New System.Drawing.Point(182, 93)
        Me.cmdRemovePI.Name = "cmdRemovePI"
        Me.cmdRemovePI.Size = New System.Drawing.Size(77, 27)
        Me.cmdRemovePI.TabIndex = 172
        Me.cmdRemovePI.Text = "Remove ->"
        Me.cmdRemovePI.UseVisualStyleBackColor = True
        '
        'dgvPI
        '
        Me.dgvPI.BackgroundColor = System.Drawing.Color.White
        Me.dgvPI.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvPI.Location = New System.Drawing.Point(10, 60)
        Me.dgvPI.Name = "dgvPI"
        Me.dgvPI.ReadOnly = True
        Me.dgvPI.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvPI.Size = New System.Drawing.Size(166, 121)
        Me.dgvPI.TabIndex = 168
        '
        'dgvSource
        '
        Me.dgvSource.BackgroundColor = System.Drawing.Color.White
        Me.dgvSource.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSource.Location = New System.Drawing.Point(265, 60)
        Me.dgvSource.Name = "dgvSource"
        Me.dgvSource.ReadOnly = True
        Me.dgvSource.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvSource.Size = New System.Drawing.Size(442, 273)
        Me.dgvSource.TabIndex = 169
        '
        'cmdOK
        '
        Me.cmdOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.Blue
        Me.cmdOK.Location = New System.Drawing.Point(0, 10)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(80, 35)
        Me.cmdOK.TabIndex = 0
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'lblPI
        '
        Me.lblPI.AutoSize = True
        Me.lblPI.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPI.Location = New System.Drawing.Point(7, 42)
        Me.lblPI.Name = "lblPI"
        Me.lblPI.Size = New System.Drawing.Size(111, 15)
        Me.lblPI.TabIndex = 173
        Me.lblPI.Text = "Configured PI(s)"
        '
        'cmdAddPI
        '
        Me.cmdAddPI.ForeColor = System.Drawing.Color.Blue
        Me.cmdAddPI.Location = New System.Drawing.Point(182, 60)
        Me.cmdAddPI.Name = "cmdAddPI"
        Me.cmdAddPI.Size = New System.Drawing.Size(77, 27)
        Me.cmdAddPI.TabIndex = 171
        Me.cmdAddPI.Text = "<- Add"
        Me.cmdAddPI.UseVisualStyleBackColor = True
        '
        'pan1
        '
        Me.pan1.Controls.Add(Me.cmdCancel)
        Me.pan1.Controls.Add(Me.cmdOK)
        Me.pan1.Location = New System.Drawing.Point(12, 378)
        Me.pan1.Name = "pan1"
        Me.pan1.Size = New System.Drawing.Size(182, 45)
        Me.pan1.TabIndex = 170
        '
        'dgvAnal
        '
        Me.dgvAnal.BackgroundColor = System.Drawing.Color.White
        Me.dgvAnal.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvAnal.Location = New System.Drawing.Point(10, 212)
        Me.dgvAnal.Name = "dgvAnal"
        Me.dgvAnal.ReadOnly = True
        Me.dgvAnal.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvAnal.Size = New System.Drawing.Size(166, 121)
        Me.dgvAnal.TabIndex = 178
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(7, 194)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(143, 15)
        Me.Label2.TabIndex = 179
        Me.Label2.Text = "Configured Analyst(s)"
        '
        'cmdRemoveAnal
        '
        Me.cmdRemoveAnal.ForeColor = System.Drawing.Color.Red
        Me.cmdRemoveAnal.Location = New System.Drawing.Point(182, 245)
        Me.cmdRemoveAnal.Name = "cmdRemoveAnal"
        Me.cmdRemoveAnal.Size = New System.Drawing.Size(77, 27)
        Me.cmdRemoveAnal.TabIndex = 181
        Me.cmdRemoveAnal.Text = "Remove ->"
        Me.cmdRemoveAnal.UseVisualStyleBackColor = True
        '
        'cmdAddAnal
        '
        Me.cmdAddAnal.ForeColor = System.Drawing.Color.Blue
        Me.cmdAddAnal.Location = New System.Drawing.Point(182, 212)
        Me.cmdAddAnal.Name = "cmdAddAnal"
        Me.cmdAddAnal.Size = New System.Drawing.Size(77, 27)
        Me.cmdAddAnal.TabIndex = 180
        Me.cmdAddAnal.Text = "<- Add"
        Me.cmdAddAnal.UseVisualStyleBackColor = True
        '
        'frmConfigAssayPers
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(727, 455)
        Me.Controls.Add(Me.cmdRemoveAnal)
        Me.Controls.Add(Me.cmdAddAnal)
        Me.Controls.Add(Me.dgvAnal)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmdRemovePI)
        Me.Controls.Add(Me.dgvPI)
        Me.Controls.Add(Me.dgvSource)
        Me.Controls.Add(Me.lblPI)
        Me.Controls.Add(Me.cmdAddPI)
        Me.Controls.Add(Me.pan1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmConfigAssayPers"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Configure Assay Personnel"
        CType(Me.dgvPI, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pan1.ResumeLayout(False)
        CType(Me.dgvAnal, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdRemovePI As System.Windows.Forms.Button
    Friend WithEvents dgvPI As System.Windows.Forms.DataGridView
    Friend WithEvents dgvSource As System.Windows.Forms.DataGridView
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents lblPI As System.Windows.Forms.Label
    Friend WithEvents cmdAddPI As System.Windows.Forms.Button
    Friend WithEvents pan1 As System.Windows.Forms.Panel
    Friend WithEvents dgvAnal As System.Windows.Forms.DataGridView
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmdRemoveAnal As System.Windows.Forms.Button
    Friend WithEvents cmdAddAnal As System.Windows.Forms.Button
End Class
