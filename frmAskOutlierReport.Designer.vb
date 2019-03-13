<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAskOutlierReport
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAskOutlierReport))
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.gb1 = New System.Windows.Forms.GroupBox()
        Me.gb2 = New System.Windows.Forms.GroupBox()
        Me.rbSelected = New System.Windows.Forms.RadioButton()
        Me.rbAll = New System.Windows.Forms.RadioButton()
        Me.rbDetailed = New System.Windows.Forms.RadioButton()
        Me.rbSummary = New System.Windows.Forms.RadioButton()
        Me.gb1.SuspendLayout()
        Me.gb2.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdOK
        '
        Me.cmdOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.Blue
        Me.cmdOK.Location = New System.Drawing.Point(14, 213)
        Me.cmdOK.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(85, 54)
        Me.cmdOK.TabIndex = 0
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(146, 213)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(85, 54)
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'gb1
        '
        Me.gb1.Controls.Add(Me.gb2)
        Me.gb1.Controls.Add(Me.rbDetailed)
        Me.gb1.Controls.Add(Me.rbSummary)
        Me.gb1.Location = New System.Drawing.Point(14, 29)
        Me.gb1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gb1.Name = "gb1"
        Me.gb1.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gb1.Size = New System.Drawing.Size(217, 177)
        Me.gb1.TabIndex = 2
        Me.gb1.TabStop = False
        Me.gb1.Text = "Report Type"
        '
        'gb2
        '
        Me.gb2.Controls.Add(Me.rbSelected)
        Me.gb2.Controls.Add(Me.rbAll)
        Me.gb2.Enabled = False
        Me.gb2.Location = New System.Drawing.Point(31, 78)
        Me.gb2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gb2.Name = "gb2"
        Me.gb2.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gb2.Size = New System.Drawing.Size(155, 78)
        Me.gb2.TabIndex = 2
        Me.gb2.TabStop = False
        '
        'rbSelected
        '
        Me.rbSelected.AutoSize = True
        Me.rbSelected.Location = New System.Drawing.Point(7, 43)
        Me.rbSelected.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbSelected.Name = "rbSelected"
        Me.rbSelected.Size = New System.Drawing.Size(111, 21)
        Me.rbSelected.TabIndex = 2
        Me.rbSelected.Text = "Selected Table"
        Me.rbSelected.UseVisualStyleBackColor = True
        '
        'rbAll
        '
        Me.rbAll.AutoSize = True
        Me.rbAll.Checked = True
        Me.rbAll.Location = New System.Drawing.Point(7, 13)
        Me.rbAll.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbAll.Name = "rbAll"
        Me.rbAll.Size = New System.Drawing.Size(82, 21)
        Me.rbAll.TabIndex = 1
        Me.rbAll.TabStop = True
        Me.rbAll.Text = "All Tables"
        Me.rbAll.UseVisualStyleBackColor = True
        '
        'rbDetailed
        '
        Me.rbDetailed.AutoSize = True
        Me.rbDetailed.Location = New System.Drawing.Point(7, 55)
        Me.rbDetailed.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbDetailed.Name = "rbDetailed"
        Me.rbDetailed.Size = New System.Drawing.Size(189, 21)
        Me.rbDetailed.TabIndex = 1
        Me.rbDetailed.Text = "Summary + Detailed Report"
        Me.rbDetailed.UseVisualStyleBackColor = True
        '
        'rbSummary
        '
        Me.rbSummary.AutoSize = True
        Me.rbSummary.Checked = True
        Me.rbSummary.Location = New System.Drawing.Point(7, 25)
        Me.rbSummary.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbSummary.Name = "rbSummary"
        Me.rbSummary.Size = New System.Drawing.Size(154, 21)
        Me.rbSummary.TabIndex = 0
        Me.rbSummary.TabStop = True
        Me.rbSummary.Text = "Summary Report Only"
        Me.rbSummary.UseVisualStyleBackColor = True
        '
        'frmAskOutlierReport
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(247, 280)
        Me.ControlBox = False
        Me.Controls.Add(Me.gb1)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "frmAskOutlierReport"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " Generate Outlier Report"
        Me.gb1.ResumeLayout(False)
        Me.gb1.PerformLayout()
        Me.gb2.ResumeLayout(False)
        Me.gb2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents gb1 As System.Windows.Forms.GroupBox
    Friend WithEvents rbDetailed As System.Windows.Forms.RadioButton
    Friend WithEvents rbSummary As System.Windows.Forms.RadioButton
    Friend WithEvents gb2 As System.Windows.Forms.GroupBox
    Friend WithEvents rbSelected As System.Windows.Forms.RadioButton
    Friend WithEvents rbAll As System.Windows.Forms.RadioButton
End Class
