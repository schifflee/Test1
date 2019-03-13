<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmApplyDataToGroups
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
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmApplyDataToGroups))
        Me.pan1 = New System.Windows.Forms.Panel()
        Me.cmdCancel1 = New System.Windows.Forms.Button()
        Me.cmdOK1 = New System.Windows.Forms.Button()
        Me.dgvGroupSummary = New System.Windows.Forms.DataGridView()
        Me.dgvFrom = New System.Windows.Forms.DataGridView()
        Me.dgvTo = New System.Windows.Forms.DataGridView()
        Me.lblFrom = New System.Windows.Forms.Label()
        Me.lblTo = New System.Windows.Forms.Label()
        Me.cmdAddFrom = New System.Windows.Forms.Button()
        Me.cmdRemoveFrom = New System.Windows.Forms.Button()
        Me.cmdRemoveTo = New System.Windows.Forms.Button()
        Me.cmdAddTo = New System.Windows.Forms.Button()
        Me.cmdAddAll = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.pan1.SuspendLayout()
        CType(Me.dgvGroupSummary, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvFrom, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvTo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pan1
        '
        Me.pan1.Controls.Add(Me.cmdCancel1)
        Me.pan1.Controls.Add(Me.cmdOK1)
        Me.pan1.Location = New System.Drawing.Point(27, 491)
        Me.pan1.Name = "pan1"
        Me.pan1.Size = New System.Drawing.Size(305, 45)
        Me.pan1.TabIndex = 143
        '
        'cmdCancel1
        '
        Me.cmdCancel1.CausesValidation = False
        Me.cmdCancel1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel1.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel1.Location = New System.Drawing.Point(155, 7)
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
        Me.cmdOK1.Location = New System.Drawing.Point(69, 6)
        Me.cmdOK1.Name = "cmdOK1"
        Me.cmdOK1.Size = New System.Drawing.Size(80, 35)
        Me.cmdOK1.TabIndex = 0
        Me.cmdOK1.Text = "&OK"
        Me.cmdOK1.UseVisualStyleBackColor = True
        '
        'dgvGroupSummary
        '
        Me.dgvGroupSummary.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvGroupSummary.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvGroupSummary.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvGroupSummary.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvGroupSummary.Location = New System.Drawing.Point(27, 94)
        Me.dgvGroupSummary.Name = "dgvGroupSummary"
        Me.dgvGroupSummary.ReadOnly = True
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvGroupSummary.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvGroupSummary.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvGroupSummary.Size = New System.Drawing.Size(95, 391)
        Me.dgvGroupSummary.TabIndex = 161
        '
        'dgvFrom
        '
        Me.dgvFrom.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvFrom.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.dgvFrom.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvFrom.DefaultCellStyle = DataGridViewCellStyle5
        Me.dgvFrom.Location = New System.Drawing.Point(246, 94)
        Me.dgvFrom.Name = "dgvFrom"
        Me.dgvFrom.ReadOnly = True
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvFrom.RowHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.dgvFrom.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvFrom.Size = New System.Drawing.Size(95, 182)
        Me.dgvFrom.TabIndex = 162
        '
        'dgvTo
        '
        Me.dgvTo.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvTo.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle7
        Me.dgvTo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvTo.DefaultCellStyle = DataGridViewCellStyle8
        Me.dgvTo.Location = New System.Drawing.Point(246, 304)
        Me.dgvTo.Name = "dgvTo"
        Me.dgvTo.ReadOnly = True
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvTo.RowHeadersDefaultCellStyle = DataGridViewCellStyle9
        Me.dgvTo.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvTo.Size = New System.Drawing.Size(95, 182)
        Me.dgvTo.TabIndex = 163
        '
        'lblFrom
        '
        Me.lblFrom.ForeColor = System.Drawing.Color.Blue
        Me.lblFrom.Location = New System.Drawing.Point(130, 95)
        Me.lblFrom.Name = "lblFrom"
        Me.lblFrom.Size = New System.Drawing.Size(113, 31)
        Me.lblFrom.TabIndex = 166
        Me.lblFrom.Text = "Select Route to copy data FROM -->"
        '
        'lblTo
        '
        Me.lblTo.ForeColor = System.Drawing.Color.Red
        Me.lblTo.Location = New System.Drawing.Point(130, 304)
        Me.lblTo.Name = "lblTo"
        Me.lblTo.Size = New System.Drawing.Size(102, 31)
        Me.lblTo.TabIndex = 167
        Me.lblTo.Text = "Select Route to copy data TO -->"
        '
        'cmdAddFrom
        '
        Me.cmdAddFrom.ForeColor = System.Drawing.Color.Blue
        Me.cmdAddFrom.Location = New System.Drawing.Point(133, 129)
        Me.cmdAddFrom.Name = "cmdAddFrom"
        Me.cmdAddFrom.Size = New System.Drawing.Size(70, 22)
        Me.cmdAddFrom.TabIndex = 170
        Me.cmdAddFrom.Text = "Add ->"
        Me.cmdAddFrom.UseVisualStyleBackColor = True
        '
        'cmdRemoveFrom
        '
        Me.cmdRemoveFrom.ForeColor = System.Drawing.Color.Red
        Me.cmdRemoveFrom.Location = New System.Drawing.Point(170, 157)
        Me.cmdRemoveFrom.Name = "cmdRemoveFrom"
        Me.cmdRemoveFrom.Size = New System.Drawing.Size(70, 22)
        Me.cmdRemoveFrom.TabIndex = 171
        Me.cmdRemoveFrom.Text = "<- Remove"
        Me.cmdRemoveFrom.UseVisualStyleBackColor = True
        '
        'cmdRemoveTo
        '
        Me.cmdRemoveTo.ForeColor = System.Drawing.Color.Red
        Me.cmdRemoveTo.Location = New System.Drawing.Point(170, 394)
        Me.cmdRemoveTo.Name = "cmdRemoveTo"
        Me.cmdRemoveTo.Size = New System.Drawing.Size(70, 22)
        Me.cmdRemoveTo.TabIndex = 173
        Me.cmdRemoveTo.Text = "<- Clear All"
        Me.cmdRemoveTo.UseVisualStyleBackColor = True
        '
        'cmdAddTo
        '
        Me.cmdAddTo.ForeColor = System.Drawing.Color.Blue
        Me.cmdAddTo.Location = New System.Drawing.Point(133, 338)
        Me.cmdAddTo.Name = "cmdAddTo"
        Me.cmdAddTo.Size = New System.Drawing.Size(70, 22)
        Me.cmdAddTo.TabIndex = 172
        Me.cmdAddTo.Text = "Add ->"
        Me.cmdAddTo.UseVisualStyleBackColor = True
        '
        'cmdAddAll
        '
        Me.cmdAddAll.ForeColor = System.Drawing.Color.Green
        Me.cmdAddAll.Location = New System.Drawing.Point(133, 366)
        Me.cmdAddAll.Name = "cmdAddAll"
        Me.cmdAddAll.Size = New System.Drawing.Size(70, 22)
        Me.cmdAddAll.TabIndex = 174
        Me.cmdAddAll.Text = "Add All ->"
        Me.cmdAddAll.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(24, 78)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(47, 13)
        Me.Label1.TabIndex = 175
        Me.Label1.Text = "Routes"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(243, 79)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(66, 13)
        Me.Label2.TabIndex = 176
        Me.Label2.Text = "Copy From"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(243, 288)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(54, 13)
        Me.Label3.TabIndex = 177
        Me.Label3.Text = "Copy To"
        '
        'lblTitle
        '
        Me.lblTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.Color.Blue
        Me.lblTitle.Location = New System.Drawing.Point(12, 9)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(346, 59)
        Me.lblTitle.TabIndex = 178
        Me.lblTitle.Text = "Routes"
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'frmApplyDataToGroups
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(370, 547)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dgvGroupSummary)
        Me.Controls.Add(Me.cmdAddAll)
        Me.Controls.Add(Me.dgvTo)
        Me.Controls.Add(Me.cmdRemoveTo)
        Me.Controls.Add(Me.cmdAddTo)
        Me.Controls.Add(Me.cmdRemoveFrom)
        Me.Controls.Add(Me.cmdAddFrom)
        Me.Controls.Add(Me.dgvFrom)
        Me.Controls.Add(Me.lblTo)
        Me.Controls.Add(Me.lblFrom)
        Me.Controls.Add(Me.pan1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmApplyDataToGroups"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Apply Data To Groups..."
        Me.pan1.ResumeLayout(False)
        CType(Me.dgvGroupSummary, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvFrom, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvTo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents pan1 As System.Windows.Forms.Panel
    Friend WithEvents cmdCancel1 As System.Windows.Forms.Button
    Friend WithEvents cmdOK1 As System.Windows.Forms.Button
    Friend WithEvents dgvGroupSummary As System.Windows.Forms.DataGridView
    Friend WithEvents dgvFrom As System.Windows.Forms.DataGridView
    Friend WithEvents dgvTo As System.Windows.Forms.DataGridView
    Friend WithEvents lblFrom As System.Windows.Forms.Label
    Friend WithEvents lblTo As System.Windows.Forms.Label
    Friend WithEvents cmdAddFrom As System.Windows.Forms.Button
    Friend WithEvents cmdRemoveFrom As System.Windows.Forms.Button
    Friend WithEvents cmdRemoveTo As System.Windows.Forms.Button
    Friend WithEvents cmdAddTo As System.Windows.Forms.Button
    Friend WithEvents cmdAddAll As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblTitle As System.Windows.Forms.Label
End Class
