<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAnalyticalRunSummary
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAnalyticalRunSummary))
        Me.dgvAnalRunSummary = New System.Windows.Forms.DataGridView()
        Me.lblAnalyticalRunSummary = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.panAnalRunChoices = New System.Windows.Forms.Panel()
        Me.lblAnalRunReportOptions = New System.Windows.Forms.Label()
        Me.chkPSAE = New System.Windows.Forms.CheckBox()
        Me.chkNoRegrPerformed = New System.Windows.Forms.CheckBox()
        Me.chkRegrPerformed = New System.Windows.Forms.CheckBox()
        Me.chkRejected = New System.Windows.Forms.CheckBox()
        Me.chkAccepted = New System.Windows.Forms.CheckBox()
        Me.chkAll = New System.Windows.Forms.CheckBox()
        Me.cmdExit = New System.Windows.Forms.Button()
        CType(Me.dgvAnalRunSummary, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panAnalRunChoices.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvAnalRunSummary
        '
        Me.dgvAnalRunSummary.AllowUserToAddRows = False
        Me.dgvAnalRunSummary.AllowUserToDeleteRows = False
        Me.dgvAnalRunSummary.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvAnalRunSummary.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvAnalRunSummary.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvAnalRunSummary.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvAnalRunSummary.Location = New System.Drawing.Point(12, 94)
        Me.dgvAnalRunSummary.Name = "dgvAnalRunSummary"
        Me.dgvAnalRunSummary.ReadOnly = True
        Me.dgvAnalRunSummary.Size = New System.Drawing.Size(1097, 469)
        Me.dgvAnalRunSummary.TabIndex = 0
        '
        'lblAnalyticalRunSummary
        '
        Me.lblAnalyticalRunSummary.BackColor = System.Drawing.Color.Transparent
        Me.lblAnalyticalRunSummary.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAnalyticalRunSummary.ForeColor = System.Drawing.Color.Blue
        Me.lblAnalyticalRunSummary.Location = New System.Drawing.Point(8, 9)
        Me.lblAnalyticalRunSummary.Name = "lblAnalyticalRunSummary"
        Me.lblAnalyticalRunSummary.Size = New System.Drawing.Size(218, 20)
        Me.lblAnalyticalRunSummary.TabIndex = 97
        Me.lblAnalyticalRunSummary.Text = "Analytical Run Summary"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(983, 16)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(105, 45)
        Me.Button1.TabIndex = 100
        Me.Button1.Text = "Test"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'panAnalRunChoices
        '
        Me.panAnalRunChoices.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panAnalRunChoices.Controls.Add(Me.lblAnalRunReportOptions)
        Me.panAnalRunChoices.Controls.Add(Me.chkPSAE)
        Me.panAnalRunChoices.Controls.Add(Me.chkNoRegrPerformed)
        Me.panAnalRunChoices.Controls.Add(Me.chkRegrPerformed)
        Me.panAnalRunChoices.Controls.Add(Me.chkRejected)
        Me.panAnalRunChoices.Controls.Add(Me.chkAccepted)
        Me.panAnalRunChoices.Controls.Add(Me.chkAll)
        Me.panAnalRunChoices.Location = New System.Drawing.Point(228, 3)
        Me.panAnalRunChoices.Name = "panAnalRunChoices"
        Me.panAnalRunChoices.Size = New System.Drawing.Size(354, 87)
        Me.panAnalRunChoices.TabIndex = 160
        '
        'lblAnalRunReportOptions
        '
        Me.lblAnalRunReportOptions.AutoSize = True
        Me.lblAnalRunReportOptions.Location = New System.Drawing.Point(9, 3)
        Me.lblAnalRunReportOptions.Name = "lblAnalRunReportOptions"
        Me.lblAnalRunReportOptions.Size = New System.Drawing.Size(206, 13)
        Me.lblAnalRunReportOptions.TabIndex = 6
        Me.lblAnalRunReportOptions.Text = "Include the following run/regression types:" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'chkPSAE
        '
        Me.chkPSAE.AutoSize = True
        Me.chkPSAE.Location = New System.Drawing.Point(172, 61)
        Me.chkPSAE.Name = "chkPSAE"
        Me.chkPSAE.Size = New System.Drawing.Size(120, 17)
        Me.chkPSAE.TabIndex = 5
        Me.chkPSAE.Text = "Include PSAE Runs"
        Me.chkPSAE.UseVisualStyleBackColor = True
        '
        'chkNoRegrPerformed
        '
        Me.chkNoRegrPerformed.AutoSize = True
        Me.chkNoRegrPerformed.Location = New System.Drawing.Point(172, 42)
        Me.chkNoRegrPerformed.Name = "chkNoRegrPerformed"
        Me.chkNoRegrPerformed.Size = New System.Drawing.Size(149, 17)
        Me.chkNoRegrPerformed.TabIndex = 4
        Me.chkNoRegrPerformed.Text = "NO Regression Performed"
        Me.chkNoRegrPerformed.UseVisualStyleBackColor = True
        '
        'chkRegrPerformed
        '
        Me.chkRegrPerformed.AutoSize = True
        Me.chkRegrPerformed.Location = New System.Drawing.Point(172, 22)
        Me.chkRegrPerformed.Name = "chkRegrPerformed"
        Me.chkRegrPerformed.Size = New System.Drawing.Size(130, 17)
        Me.chkRegrPerformed.TabIndex = 3
        Me.chkRegrPerformed.Text = "Regression Performed"
        Me.chkRegrPerformed.UseVisualStyleBackColor = True
        '
        'chkRejected
        '
        Me.chkRejected.AutoSize = True
        Me.chkRejected.Location = New System.Drawing.Point(12, 61)
        Me.chkRejected.Name = "chkRejected"
        Me.chkRejected.Size = New System.Drawing.Size(125, 17)
        Me.chkRejected.TabIndex = 2
        Me.chkRejected.Text = "Rejected Regression"
        Me.chkRejected.UseVisualStyleBackColor = True
        '
        'chkAccepted
        '
        Me.chkAccepted.AutoSize = True
        Me.chkAccepted.Location = New System.Drawing.Point(12, 42)
        Me.chkAccepted.Name = "chkAccepted"
        Me.chkAccepted.Size = New System.Drawing.Size(128, 17)
        Me.chkAccepted.TabIndex = 1
        Me.chkAccepted.Text = "Accepted Regression"
        Me.chkAccepted.UseVisualStyleBackColor = True
        '
        'chkAll
        '
        Me.chkAll.AutoSize = True
        Me.chkAll.Location = New System.Drawing.Point(12, 22)
        Me.chkAll.Name = "chkAll"
        Me.chkAll.Size = New System.Drawing.Size(113, 17)
        Me.chkAll.TabIndex = 0
        Me.chkAll.Text = "All Analytical Runs"
        Me.chkAll.UseVisualStyleBackColor = True
        '
        'cmdExit
        '
        Me.cmdExit.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdExit.FlatAppearance.BorderSize = 0
        Me.cmdExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdExit.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdExit.ForeColor = System.Drawing.Color.Red
        Me.cmdExit.Location = New System.Drawing.Point(12, 36)
        Me.cmdExit.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(79, 33)
        Me.cmdExit.TabIndex = 161
        Me.cmdExit.Text = "G&o Back"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'frmAnalyticalRunSummary
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1121, 575)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.panAnalRunChoices)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.lblAnalyticalRunSummary)
        Me.Controls.Add(Me.dgvAnalRunSummary)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmAnalyticalRunSummary"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Analytical Run Summary"
        CType(Me.dgvAnalRunSummary, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panAnalRunChoices.ResumeLayout(False)
        Me.panAnalRunChoices.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvAnalRunSummary As System.Windows.Forms.DataGridView
    Friend WithEvents lblAnalyticalRunSummary As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents panAnalRunChoices As System.Windows.Forms.Panel
    Friend WithEvents lblAnalRunReportOptions As System.Windows.Forms.Label
    Friend WithEvents chkPSAE As System.Windows.Forms.CheckBox
    Friend WithEvents chkNoRegrPerformed As System.Windows.Forms.CheckBox
    Friend WithEvents chkRegrPerformed As System.Windows.Forms.CheckBox
    Friend WithEvents chkRejected As System.Windows.Forms.CheckBox
    Friend WithEvents chkAccepted As System.Windows.Forms.CheckBox
    Friend WithEvents chkAll As System.Windows.Forms.CheckBox
    Friend WithEvents cmdExit As System.Windows.Forms.Button
End Class
