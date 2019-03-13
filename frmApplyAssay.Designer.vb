<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmApplyAssay
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmApplyAssay))
        Me.dgvStudy = New System.Windows.Forms.DataGridView()
        Me.txtStudyFilter = New System.Windows.Forms.TextBox()
        Me.dgvAssay = New System.Windows.Forms.DataGridView()
        Me.gbxFilter = New System.Windows.Forms.GroupBox()
        Me.rbStudyName = New System.Windows.Forms.RadioButton()
        Me.rbStudyNumber = New System.Windows.Forms.RadioButton()
        Me.lblFilter = New System.Windows.Forms.Label()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.lblAssay = New System.Windows.Forms.Label()
        Me.txtAssayName = New System.Windows.Forms.TextBox()
        Me.panTemplate = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblTemplate = New System.Windows.Forms.Label()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.chkApplyTemplate = New System.Windows.Forms.CheckBox()
        Me.panNew = New System.Windows.Forms.Panel()
        CType(Me.dgvStudy, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvAssay, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbxFilter.SuspendLayout()
        Me.panTemplate.SuspendLayout()
        Me.panNew.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvStudy
        '
        Me.dgvStudy.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvStudy.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvStudy.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvStudy.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvStudy.Location = New System.Drawing.Point(11, 148)
        Me.dgvStudy.MultiSelect = False
        Me.dgvStudy.Name = "dgvStudy"
        Me.dgvStudy.ReadOnly = True
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvStudy.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvStudy.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvStudy.Size = New System.Drawing.Size(223, 270)
        Me.dgvStudy.TabIndex = 139
        '
        'txtStudyFilter
        '
        Me.txtStudyFilter.Location = New System.Drawing.Point(11, 107)
        Me.txtStudyFilter.Name = "txtStudyFilter"
        Me.txtStudyFilter.Size = New System.Drawing.Size(223, 20)
        Me.txtStudyFilter.TabIndex = 140
        '
        'dgvAssay
        '
        Me.dgvAssay.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvAssay.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.dgvAssay.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvAssay.DefaultCellStyle = DataGridViewCellStyle5
        Me.dgvAssay.Location = New System.Drawing.Point(240, 148)
        Me.dgvAssay.MultiSelect = False
        Me.dgvAssay.Name = "dgvAssay"
        Me.dgvAssay.ReadOnly = True
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvAssay.RowHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.dgvAssay.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvAssay.Size = New System.Drawing.Size(153, 270)
        Me.dgvAssay.TabIndex = 141
        '
        'gbxFilter
        '
        Me.gbxFilter.Controls.Add(Me.rbStudyName)
        Me.gbxFilter.Controls.Add(Me.rbStudyNumber)
        Me.gbxFilter.Location = New System.Drawing.Point(11, 33)
        Me.gbxFilter.Name = "gbxFilter"
        Me.gbxFilter.Size = New System.Drawing.Size(202, 46)
        Me.gbxFilter.TabIndex = 142
        Me.gbxFilter.TabStop = False
        Me.gbxFilter.Text = "Filter and Sort Studies By:"
        '
        'rbStudyName
        '
        Me.rbStudyName.AutoSize = True
        Me.rbStudyName.Checked = True
        Me.rbStudyName.Location = New System.Drawing.Point(104, 19)
        Me.rbStudyName.Name = "rbStudyName"
        Me.rbStudyName.Size = New System.Drawing.Size(83, 17)
        Me.rbStudyName.TabIndex = 1
        Me.rbStudyName.TabStop = True
        Me.rbStudyName.Text = "Study Name"
        Me.rbStudyName.UseVisualStyleBackColor = True
        '
        'rbStudyNumber
        '
        Me.rbStudyNumber.AutoSize = True
        Me.rbStudyNumber.Location = New System.Drawing.Point(6, 19)
        Me.rbStudyNumber.Name = "rbStudyNumber"
        Me.rbStudyNumber.Size = New System.Drawing.Size(92, 17)
        Me.rbStudyNumber.TabIndex = 0
        Me.rbStudyNumber.Text = "Study Number"
        Me.rbStudyNumber.UseVisualStyleBackColor = True
        '
        'lblFilter
        '
        Me.lblFilter.AutoSize = True
        Me.lblFilter.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFilter.Location = New System.Drawing.Point(8, 89)
        Me.lblFilter.Name = "lblFilter"
        Me.lblFilter.Size = New System.Drawing.Size(96, 15)
        Me.lblFilter.TabIndex = 143
        Me.lblFilter.Text = "Filter Studies:"
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancel.CausesValidation = False
        Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(341, 48)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(65, 25)
        Me.cmdCancel.TabIndex = 144
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'lblAssay
        '
        Me.lblAssay.AutoSize = True
        Me.lblAssay.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAssay.Location = New System.Drawing.Point(3, 3)
        Me.lblAssay.Name = "lblAssay"
        Me.lblAssay.Size = New System.Drawing.Size(159, 15)
        Me.lblAssay.TabIndex = 147
        Me.lblAssay.Text = "Enter New Assay Name:"
        '
        'txtAssayName
        '
        Me.txtAssayName.Location = New System.Drawing.Point(6, 21)
        Me.txtAssayName.Name = "txtAssayName"
        Me.txtAssayName.Size = New System.Drawing.Size(223, 20)
        Me.txtAssayName.TabIndex = 146
        '
        'panTemplate
        '
        Me.panTemplate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panTemplate.Controls.Add(Me.Label2)
        Me.panTemplate.Controls.Add(Me.Label1)
        Me.panTemplate.Controls.Add(Me.lblTemplate)
        Me.panTemplate.Controls.Add(Me.dgvStudy)
        Me.panTemplate.Controls.Add(Me.txtStudyFilter)
        Me.panTemplate.Controls.Add(Me.dgvAssay)
        Me.panTemplate.Controls.Add(Me.gbxFilter)
        Me.panTemplate.Controls.Add(Me.lblFilter)
        Me.panTemplate.Location = New System.Drawing.Point(12, 88)
        Me.panTemplate.Name = "panTemplate"
        Me.panTemplate.Size = New System.Drawing.Size(406, 437)
        Me.panTemplate.TabIndex = 148
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Blue
        Me.Label2.Location = New System.Drawing.Point(237, 130)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(119, 15)
        Me.Label2.TabIndex = 150
        Me.Label2.Text = "Choose an Assay:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 130)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(55, 15)
        Me.Label1.TabIndex = 149
        Me.Label1.Text = "Studies"
        '
        'lblTemplate
        '
        Me.lblTemplate.AutoSize = True
        Me.lblTemplate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTemplate.ForeColor = System.Drawing.Color.Blue
        Me.lblTemplate.Location = New System.Drawing.Point(8, 9)
        Me.lblTemplate.Name = "lblTemplate"
        Me.lblTemplate.Size = New System.Drawing.Size(214, 15)
        Me.lblTemplate.TabIndex = 148
        Me.lblTemplate.Text = "Choose an Assay as a Template:"
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.Blue
        Me.cmdOK.Location = New System.Drawing.Point(341, 16)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(65, 25)
        Me.cmdOK.TabIndex = 149
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'chkApplyTemplate
        '
        Me.chkApplyTemplate.AutoSize = True
        Me.chkApplyTemplate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkApplyTemplate.Location = New System.Drawing.Point(6, 47)
        Me.chkApplyTemplate.Name = "chkApplyTemplate"
        Me.chkApplyTemplate.Size = New System.Drawing.Size(136, 17)
        Me.chkApplyTemplate.TabIndex = 150
        Me.chkApplyTemplate.Text = "Apply a Template..."
        Me.chkApplyTemplate.UseVisualStyleBackColor = True
        '
        'panNew
        '
        Me.panNew.Controls.Add(Me.lblAssay)
        Me.panNew.Controls.Add(Me.chkApplyTemplate)
        Me.panNew.Controls.Add(Me.txtAssayName)
        Me.panNew.Location = New System.Drawing.Point(24, 9)
        Me.panNew.Name = "panNew"
        Me.panNew.Size = New System.Drawing.Size(235, 73)
        Me.panNew.TabIndex = 151
        '
        'frmApplyAssay
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LemonChiffon
        Me.ClientSize = New System.Drawing.Size(439, 535)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.panNew)
        Me.Controls.Add(Me.panTemplate)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmApplyAssay"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Apply Existing Assay..."
        CType(Me.dgvStudy, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvAssay, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbxFilter.ResumeLayout(False)
        Me.gbxFilter.PerformLayout()
        Me.panTemplate.ResumeLayout(False)
        Me.panTemplate.PerformLayout()
        Me.panNew.ResumeLayout(False)
        Me.panNew.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvStudy As System.Windows.Forms.DataGridView
    Friend WithEvents txtStudyFilter As System.Windows.Forms.TextBox
    Friend WithEvents dgvAssay As System.Windows.Forms.DataGridView
    Friend WithEvents gbxFilter As System.Windows.Forms.GroupBox
    Friend WithEvents rbStudyName As System.Windows.Forms.RadioButton
    Friend WithEvents rbStudyNumber As System.Windows.Forms.RadioButton
    Friend WithEvents lblFilter As System.Windows.Forms.Label
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents lblAssay As System.Windows.Forms.Label
    Friend WithEvents txtAssayName As System.Windows.Forms.TextBox
    Friend WithEvents panTemplate As System.Windows.Forms.Panel
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents chkApplyTemplate As System.Windows.Forms.CheckBox
    Friend WithEvents panNew As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblTemplate As System.Windows.Forms.Label
End Class
