<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWordStatementsActiveTemplates
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.cmdCancelEdit = New System.Windows.Forms.Button()
        Me.cmdDone = New System.Windows.Forms.Button()
        Me.dgvReportStatements = New System.Windows.Forms.DataGridView()
        Me.gbActive = New System.Windows.Forms.GroupBox()
        Me.rbInactive = New System.Windows.Forms.RadioButton()
        Me.rbActive = New System.Windows.Forms.RadioButton()
        Me.rbAll = New System.Windows.Forms.RadioButton()
        Me.lbl1 = New System.Windows.Forms.Label()
        Me.lblNote = New System.Windows.Forms.Label()
        CType(Me.dgvReportStatements, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbActive.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdCancelEdit
        '
        Me.cmdCancelEdit.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancelEdit.CausesValidation = False
        Me.cmdCancelEdit.FlatAppearance.BorderSize = 0
        Me.cmdCancelEdit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCancelEdit.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancelEdit.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.cmdCancelEdit.Location = New System.Drawing.Point(104, 412)
        Me.cmdCancelEdit.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCancelEdit.Name = "cmdCancelEdit"
        Me.cmdCancelEdit.Size = New System.Drawing.Size(83, 61)
        Me.cmdCancelEdit.TabIndex = 58
        Me.cmdCancelEdit.Text = "&Cancel"
        Me.cmdCancelEdit.UseVisualStyleBackColor = False
        '
        'cmdDone
        '
        Me.cmdDone.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdDone.FlatAppearance.BorderSize = 0
        Me.cmdDone.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdDone.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDone.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdDone.Location = New System.Drawing.Point(14, 412)
        Me.cmdDone.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdDone.Name = "cmdDone"
        Me.cmdDone.Size = New System.Drawing.Size(83, 61)
        Me.cmdDone.TabIndex = 57
        Me.cmdDone.Text = "&Done"
        Me.cmdDone.UseVisualStyleBackColor = False
        '
        'dgvReportStatements
        '
        Me.dgvReportStatements.AllowUserToAddRows = False
        Me.dgvReportStatements.AllowUserToDeleteRows = False
        Me.dgvReportStatements.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvReportStatements.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvReportStatements.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvReportStatements.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvReportStatements.Location = New System.Drawing.Point(14, 139)
        Me.dgvReportStatements.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvReportStatements.MultiSelect = False
        Me.dgvReportStatements.Name = "dgvReportStatements"
        Me.dgvReportStatements.Size = New System.Drawing.Size(384, 269)
        Me.dgvReportStatements.TabIndex = 59
        '
        'gbActive
        '
        Me.gbActive.Controls.Add(Me.rbInactive)
        Me.gbActive.Controls.Add(Me.rbActive)
        Me.gbActive.Controls.Add(Me.rbAll)
        Me.gbActive.Location = New System.Drawing.Point(14, 8)
        Me.gbActive.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbActive.Name = "gbActive"
        Me.gbActive.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbActive.Size = New System.Drawing.Size(384, 102)
        Me.gbActive.TabIndex = 60
        Me.gbActive.TabStop = False
        Me.gbActive.Text = "Show Templates"
        '
        'rbInactive
        '
        Me.rbInactive.AutoSize = True
        Me.rbInactive.Location = New System.Drawing.Point(22, 73)
        Me.rbInactive.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbInactive.Name = "rbInactive"
        Me.rbInactive.Size = New System.Drawing.Size(104, 21)
        Me.rbInactive.TabIndex = 2
        Me.rbInactive.Text = "Show Inactive"
        Me.rbInactive.UseVisualStyleBackColor = True
        '
        'rbActive
        '
        Me.rbActive.AutoSize = True
        Me.rbActive.Location = New System.Drawing.Point(22, 49)
        Me.rbActive.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbActive.Name = "rbActive"
        Me.rbActive.Size = New System.Drawing.Size(95, 21)
        Me.rbActive.TabIndex = 1
        Me.rbActive.Text = "Show Active"
        Me.rbActive.UseVisualStyleBackColor = True
        '
        'rbAll
        '
        Me.rbAll.AutoSize = True
        Me.rbAll.Checked = True
        Me.rbAll.Location = New System.Drawing.Point(22, 25)
        Me.rbAll.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbAll.Name = "rbAll"
        Me.rbAll.Size = New System.Drawing.Size(75, 21)
        Me.rbAll.TabIndex = 0
        Me.rbAll.TabStop = True
        Me.rbAll.Text = "Show All"
        Me.rbAll.UseVisualStyleBackColor = True
        '
        'lbl1
        '
        Me.lbl1.AutoSize = True
        Me.lbl1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lbl1.Location = New System.Drawing.Point(14, 119)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(263, 17)
        Me.lbl1.TabIndex = 61
        Me.lbl1.Text = "Double-click entry to Activate/Deactivate"
        '
        'lblNote
        '
        Me.lblNote.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNote.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblNote.Location = New System.Drawing.Point(193, 412)
        Me.lblNote.Name = "lblNote"
        Me.lblNote.Size = New System.Drawing.Size(205, 60)
        Me.lblNote.TabIndex = 62
        Me.lblNote.Text = "Note: This action" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "is not audit trailed"
        Me.lblNote.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmWordStatementsActiveTemplates
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(416, 481)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblNote)
        Me.Controls.Add(Me.lbl1)
        Me.Controls.Add(Me.gbActive)
        Me.Controls.Add(Me.dgvReportStatements)
        Me.Controls.Add(Me.cmdCancelEdit)
        Me.Controls.Add(Me.cmdDone)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "frmWordStatementsActiveTemplates"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Activate/Deactive Report Templates..."
        CType(Me.dgvReportStatements, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbActive.ResumeLayout(False)
        Me.gbActive.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdCancelEdit As System.Windows.Forms.Button
    Friend WithEvents cmdDone As System.Windows.Forms.Button
    Friend WithEvents dgvReportStatements As System.Windows.Forms.DataGridView
    Friend WithEvents gbActive As System.Windows.Forms.GroupBox
    Friend WithEvents rbInactive As System.Windows.Forms.RadioButton
    Friend WithEvents rbActive As System.Windows.Forms.RadioButton
    Friend WithEvents rbAll As System.Windows.Forms.RadioButton
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Friend WithEvents lblNote As System.Windows.Forms.Label
End Class
