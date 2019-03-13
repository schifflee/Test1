<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmClearStudy
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
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdApply = New System.Windows.Forms.Button()
        Me.dgvStudies = New System.Windows.Forms.DataGridView()
        Me.lbl1 = New System.Windows.Forms.Label()
        Me.lblProjects = New System.Windows.Forms.Label()
        Me.dgvProjects = New System.Windows.Forms.DataGridView()
        Me.lblStudies = New System.Windows.Forms.Label()
        Me.txtFilter = New System.Windows.Forms.TextBox()
        Me.lblFilter = New System.Windows.Forms.Label()
        Me.txtProjectID = New System.Windows.Forms.TextBox()
        Me.lblPID = New System.Windows.Forms.Label()
        Me.lblSID = New System.Windows.Forms.Label()
        Me.txtStudyID = New System.Windows.Forms.TextBox()
        Me.cmdClear = New System.Windows.Forms.Button()
        Me.panStudies = New System.Windows.Forms.Panel()
        Me.lblReason = New System.Windows.Forms.Label()
        Me.txtReason = New System.Windows.Forms.TextBox()
        CType(Me.dgvStudies, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvProjects, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panStudies.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancel.CausesValidation = False
        Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(820, 13)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(102, 59)
        Me.cmdCancel.TabIndex = 95
        Me.cmdCancel.TabStop = False
        Me.cmdCancel.Text = "&Go Back"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'cmdApply
        '
        Me.cmdApply.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdApply.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdApply.CausesValidation = False
        Me.cmdApply.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdApply.ForeColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(227, Byte), Integer))
        Me.cmdApply.Location = New System.Drawing.Point(820, 641)
        Me.cmdApply.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdApply.Name = "cmdApply"
        Me.cmdApply.Size = New System.Drawing.Size(102, 59)
        Me.cmdApply.TabIndex = 94
        Me.cmdApply.TabStop = False
        Me.cmdApply.Text = "&Execute..."
        Me.cmdApply.UseVisualStyleBackColor = False
        '
        'dgvStudies
        '
        Me.dgvStudies.AllowUserToAddRows = False
        Me.dgvStudies.AllowUserToDeleteRows = False
        Me.dgvStudies.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvStudies.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvStudies.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.Padding = New System.Windows.Forms.Padding(0, 3, 0, 3)
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvStudies.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvStudies.Location = New System.Drawing.Point(426, 136)
        Me.dgvStudies.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvStudies.MultiSelect = False
        Me.dgvStudies.Name = "dgvStudies"
        Me.dgvStudies.ReadOnly = True
        Me.dgvStudies.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvStudies.Size = New System.Drawing.Size(496, 497)
        Me.dgvStudies.TabIndex = 96
        Me.dgvStudies.TabStop = False
        '
        'lbl1
        '
        Me.lbl1.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl1.ForeColor = System.Drawing.Color.Blue
        Me.lbl1.Location = New System.Drawing.Point(12, 0)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(775, 89)
        Me.lbl1.TabIndex = 97
        Me.lbl1.Text = "Select a study to clear from the StudyDoc database, then click the Execute button" & _
    "." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "The data in the Watson database is not affected."
        Me.lbl1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblProjects
        '
        Me.lblProjects.AutoSize = True
        Me.lblProjects.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProjects.Location = New System.Drawing.Point(12, 117)
        Me.lblProjects.Name = "lblProjects"
        Me.lblProjects.Size = New System.Drawing.Size(121, 16)
        Me.lblProjects.TabIndex = 99
        Me.lblProjects.Text = "Watson Projects"
        '
        'dgvProjects
        '
        Me.dgvProjects.AllowUserToAddRows = False
        Me.dgvProjects.AllowUserToDeleteRows = False
        Me.dgvProjects.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dgvProjects.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvProjects.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvProjects.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvProjects.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle4.Padding = New System.Windows.Forms.Padding(0, 3, 0, 3)
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvProjects.DefaultCellStyle = DataGridViewCellStyle4
        Me.dgvProjects.Location = New System.Drawing.Point(12, 136)
        Me.dgvProjects.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvProjects.MultiSelect = False
        Me.dgvProjects.Name = "dgvProjects"
        Me.dgvProjects.ReadOnly = True
        Me.dgvProjects.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvProjects.Size = New System.Drawing.Size(405, 497)
        Me.dgvProjects.TabIndex = 98
        Me.dgvProjects.TabStop = False
        '
        'lblStudies
        '
        Me.lblStudies.AutoSize = True
        Me.lblStudies.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStudies.Location = New System.Drawing.Point(0, 20)
        Me.lblStudies.Name = "lblStudies"
        Me.lblStudies.Size = New System.Drawing.Size(116, 16)
        Me.lblStudies.TabIndex = 100
        Me.lblStudies.Text = "Watson Studies"
        '
        'txtFilter
        '
        Me.txtFilter.Location = New System.Drawing.Point(253, 8)
        Me.txtFilter.Name = "txtFilter"
        Me.txtFilter.Size = New System.Drawing.Size(135, 25)
        Me.txtFilter.TabIndex = 0
        '
        'lblFilter
        '
        Me.lblFilter.AutoSize = True
        Me.lblFilter.Location = New System.Drawing.Point(117, 2)
        Me.lblFilter.Name = "lblFilter"
        Me.lblFilter.Size = New System.Drawing.Size(135, 34)
        Me.lblFilter.TabIndex = 101
        Me.lblFilter.Text = "Filter for Study Name:" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(wildcard built in)"
        '
        'txtProjectID
        '
        Me.txtProjectID.Location = New System.Drawing.Point(714, 16)
        Me.txtProjectID.Name = "txtProjectID"
        Me.txtProjectID.Size = New System.Drawing.Size(73, 25)
        Me.txtProjectID.TabIndex = 103
        Me.txtProjectID.TabStop = False
        Me.txtProjectID.Visible = False
        '
        'lblPID
        '
        Me.lblPID.AutoSize = True
        Me.lblPID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPID.Location = New System.Drawing.Point(632, 20)
        Me.lblPID.Name = "lblPID"
        Me.lblPID.Size = New System.Drawing.Size(76, 16)
        Me.lblPID.TabIndex = 104
        Me.lblPID.Text = "Project ID"
        Me.lblPID.Visible = False
        '
        'lblSID
        '
        Me.lblSID.AutoSize = True
        Me.lblSID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSID.Location = New System.Drawing.Point(632, 51)
        Me.lblSID.Name = "lblSID"
        Me.lblSID.Size = New System.Drawing.Size(66, 16)
        Me.lblSID.TabIndex = 106
        Me.lblSID.Text = "Study ID"
        Me.lblSID.Visible = False
        '
        'txtStudyID
        '
        Me.txtStudyID.Location = New System.Drawing.Point(714, 47)
        Me.txtStudyID.Name = "txtStudyID"
        Me.txtStudyID.Size = New System.Drawing.Size(73, 25)
        Me.txtStudyID.TabIndex = 105
        Me.txtStudyID.TabStop = False
        Me.txtStudyID.Visible = False
        '
        'cmdClear
        '
        Me.cmdClear.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdClear.CausesValidation = False
        Me.cmdClear.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdClear.ForeColor = System.Drawing.Color.Red
        Me.cmdClear.Location = New System.Drawing.Point(397, 7)
        Me.cmdClear.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdClear.Name = "cmdClear"
        Me.cmdClear.Size = New System.Drawing.Size(102, 25)
        Me.cmdClear.TabIndex = 107
        Me.cmdClear.TabStop = False
        Me.cmdClear.Text = "C&lear Filter"
        Me.cmdClear.UseVisualStyleBackColor = False
        '
        'panStudies
        '
        Me.panStudies.Controls.Add(Me.txtFilter)
        Me.panStudies.Controls.Add(Me.lblFilter)
        Me.panStudies.Controls.Add(Me.cmdClear)
        Me.panStudies.Controls.Add(Me.lblStudies)
        Me.panStudies.Location = New System.Drawing.Point(423, 96)
        Me.panStudies.Name = "panStudies"
        Me.panStudies.Size = New System.Drawing.Size(504, 37)
        Me.panStudies.TabIndex = 108
        Me.panStudies.TabStop = True
        '
        'lblReason
        '
        Me.lblReason.ForeColor = System.Drawing.Color.Red
        Me.lblReason.Location = New System.Drawing.Point(14, 641)
        Me.lblReason.Name = "lblReason"
        Me.lblReason.Size = New System.Drawing.Size(105, 36)
        Me.lblReason.TabIndex = 109
        Me.lblReason.Text = "Enter reason for study deletion:"
        '
        'txtReason
        '
        Me.txtReason.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtReason.Location = New System.Drawing.Point(125, 641)
        Me.txtReason.Multiline = True
        Me.txtReason.Name = "txtReason"
        Me.txtReason.Size = New System.Drawing.Size(686, 59)
        Me.txtReason.TabIndex = 110
        '
        'frmClearStudy
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(932, 713)
        Me.ControlBox = False
        Me.Controls.Add(Me.txtReason)
        Me.Controls.Add(Me.lblReason)
        Me.Controls.Add(Me.panStudies)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.lblSID)
        Me.Controls.Add(Me.txtStudyID)
        Me.Controls.Add(Me.lblPID)
        Me.Controls.Add(Me.txtProjectID)
        Me.Controls.Add(Me.lblProjects)
        Me.Controls.Add(Me.dgvProjects)
        Me.Controls.Add(Me.lbl1)
        Me.Controls.Add(Me.dgvStudies)
        Me.Controls.Add(Me.cmdApply)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "frmClearStudy"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Clear Study..."
        CType(Me.dgvStudies, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvProjects, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panStudies.ResumeLayout(False)
        Me.panStudies.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout

End Sub
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdApply As System.Windows.Forms.Button
    Friend WithEvents dgvStudies As System.Windows.Forms.DataGridView
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Friend WithEvents lblProjects As System.Windows.Forms.Label
    Friend WithEvents dgvProjects As System.Windows.Forms.DataGridView
    Friend WithEvents lblStudies As System.Windows.Forms.Label
    Friend WithEvents txtFilter As System.Windows.Forms.TextBox
    Friend WithEvents lblFilter As System.Windows.Forms.Label
    Friend WithEvents txtProjectID As System.Windows.Forms.TextBox
    Friend WithEvents lblPID As System.Windows.Forms.Label
    Friend WithEvents lblSID As System.Windows.Forms.Label
    Friend WithEvents txtStudyID As System.Windows.Forms.TextBox
    Friend WithEvents cmdClear As System.Windows.Forms.Button
    Friend WithEvents panStudies As System.Windows.Forms.Panel
    Friend WithEvents lblReason As System.Windows.Forms.Label
    Friend WithEvents txtReason As System.Windows.Forms.TextBox
End Class
