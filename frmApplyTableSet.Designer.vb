<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmApplyTableSet
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
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.gbChoose = New System.Windows.Forms.GroupBox()
        Me.rbTemplate = New System.Windows.Forms.RadioButton()
        Me.rbStudy = New System.Windows.Forms.RadioButton()
        Me.dgvSource = New System.Windows.Forms.DataGridView()
        Me.gbChoose.SuspendLayout()
        CType(Me.dgvSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancel.FlatAppearance.BorderSize = 0
        Me.cmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCancel.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(127, 458)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(3, 5, 3, 5)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(92, 38)
        Me.cmdCancel.TabIndex = 126
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdOK
        '
        Me.cmdOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdOK.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOK.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.cmdOK.FlatAppearance.BorderSize = 0
        Me.cmdOK.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdOK.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.Blue
        Me.cmdOK.Location = New System.Drawing.Point(12, 458)
        Me.cmdOK.Margin = New System.Windows.Forms.Padding(3, 5, 3, 5)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(92, 38)
        Me.cmdOK.TabIndex = 127
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'gbChoose
        '
        Me.gbChoose.Controls.Add(Me.rbStudy)
        Me.gbChoose.Controls.Add(Me.rbTemplate)
        Me.gbChoose.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbChoose.Location = New System.Drawing.Point(12, 12)
        Me.gbChoose.Name = "gbChoose"
        Me.gbChoose.Size = New System.Drawing.Size(185, 84)
        Me.gbChoose.TabIndex = 131
        Me.gbChoose.TabStop = False
        Me.gbChoose.Text = "Choose a Source..."
        '
        'rbTemplate
        '
        Me.rbTemplate.AutoSize = True
        Me.rbTemplate.Checked = True
        Me.rbTemplate.Location = New System.Drawing.Point(18, 26)
        Me.rbTemplate.Name = "rbTemplate"
        Me.rbTemplate.Size = New System.Drawing.Size(124, 21)
        Me.rbTemplate.TabIndex = 0
        Me.rbTemplate.TabStop = True
        Me.rbTemplate.Text = "Report Template"
        Me.rbTemplate.UseVisualStyleBackColor = True
        '
        'rbStudy
        '
        Me.rbStudy.AutoSize = True
        Me.rbStudy.Location = New System.Drawing.Point(18, 53)
        Me.rbStudy.Name = "rbStudy"
        Me.rbStudy.Size = New System.Drawing.Size(58, 21)
        Me.rbStudy.TabIndex = 1
        Me.rbStudy.Text = "Study"
        Me.rbStudy.UseVisualStyleBackColor = True
        '
        'dgvSource
        '
        Me.dgvSource.AllowUserToAddRows = False
        Me.dgvSource.AllowUserToDeleteRows = False
        Me.dgvSource.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSource.Location = New System.Drawing.Point(12, 102)
        Me.dgvSource.Name = "dgvSource"
        Me.dgvSource.Size = New System.Drawing.Size(207, 348)
        Me.dgvSource.TabIndex = 132
        '
        'frmApplyTableSet
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(238, 520)
        Me.ControlBox = False
        Me.Controls.Add(Me.dgvSource)
        Me.Controls.Add(Me.gbChoose)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "frmApplyTableSet"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Apply a table set..."
        Me.gbChoose.ResumeLayout(False)
        Me.gbChoose.PerformLayout()
        CType(Me.dgvSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents gbChoose As System.Windows.Forms.GroupBox
    Friend WithEvents rbStudy As System.Windows.Forms.RadioButton
    Friend WithEvents rbTemplate As System.Windows.Forms.RadioButton
    Friend WithEvents dgvSource As System.Windows.Forms.DataGridView
End Class
