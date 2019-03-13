<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFieldCodes
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFieldCodes))
        Me.panFC = New System.Windows.Forms.Panel()
        Me.dgvFC = New System.Windows.Forms.DataGridView()
        Me.lblGroup = New System.Windows.Forms.Label()
        Me.cbxGroup = New System.Windows.Forms.ComboBox()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.txtFilterFC = New System.Windows.Forms.TextBox()
        Me.lblFieldCode = New System.Windows.Forms.Label()
        Me.lblDescr = New System.Windows.Forms.Label()
        Me.txtFilterDescr = New System.Windows.Forms.TextBox()
        Me.lbl1 = New System.Windows.Forms.Label()
        Me.lblTable = New System.Windows.Forms.Label()
        Me.txtFilterTable = New System.Windows.Forms.TextBox()
        Me.gbFilter = New System.Windows.Forms.GroupBox()
        Me.lblSelection = New System.Windows.Forms.Label()
        Me.rbReportItems = New System.Windows.Forms.RadioButton()
        Me.rbNone = New System.Windows.Forms.RadioButton()
        Me.rbStudySpecific = New System.Windows.Forms.RadioButton()
        Me.cmdCopyAll = New System.Windows.Forms.Button()
        Me.gbCopyAll = New System.Windows.Forms.GroupBox()
        Me.rbWithoutLabels = New System.Windows.Forms.RadioButton()
        Me.rbWithLabels = New System.Windows.Forms.RadioButton()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.gbxlblReportTemplateFCstatus = New System.Windows.Forms.GroupBox()
        Me.lblReportTemplateFCstatus = New System.Windows.Forms.Label()
        Me.lblCount = New System.Windows.Forms.Label()
        Me.txtCount = New System.Windows.Forms.TextBox()
        Me.panFC.SuspendLayout()
        CType(Me.dgvFC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbFilter.SuspendLayout()
        Me.gbCopyAll.SuspendLayout()
        Me.gbxlblReportTemplateFCstatus.SuspendLayout()
        Me.SuspendLayout()
        '
        'panFC
        '
        Me.panFC.Controls.Add(Me.dgvFC)
        Me.panFC.Location = New System.Drawing.Point(14, 131)
        Me.panFC.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panFC.Name = "panFC"
        Me.panFC.Size = New System.Drawing.Size(1001, 399)
        Me.panFC.TabIndex = 2
        '
        'dgvFC
        '
        Me.dgvFC.AllowUserToAddRows = False
        Me.dgvFC.AllowUserToDeleteRows = False
        Me.dgvFC.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvFC.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvFC.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvFC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvFC.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvFC.Location = New System.Drawing.Point(0, 0)
        Me.dgvFC.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvFC.Name = "dgvFC"
        Me.dgvFC.ReadOnly = True
        Me.dgvFC.RowHeadersWidth = 25
        DataGridViewCellStyle2.Padding = New System.Windows.Forms.Padding(0, 2, 0, 2)
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvFC.RowsDefaultCellStyle = DataGridViewCellStyle2
        Me.dgvFC.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvFC.Size = New System.Drawing.Size(1001, 399)
        Me.dgvFC.TabIndex = 100
        '
        'lblGroup
        '
        Me.lblGroup.AutoSize = True
        Me.lblGroup.Location = New System.Drawing.Point(540, 29)
        Me.lblGroup.Name = "lblGroup"
        Me.lblGroup.Size = New System.Drawing.Size(95, 17)
        Me.lblGroup.TabIndex = 5
        Me.lblGroup.Text = "Filter by Group"
        '
        'cbxGroup
        '
        Me.cbxGroup.DropDownWidth = 500
        Me.cbxGroup.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxGroup.FormattingEnabled = True
        Me.cbxGroup.Location = New System.Drawing.Point(635, 22)
        Me.cbxGroup.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxGroup.MaxDropDownItems = 25
        Me.cbxGroup.Name = "cbxGroup"
        Me.cbxGroup.Size = New System.Drawing.Size(308, 23)
        Me.cbxGroup.TabIndex = 3
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.OK_Button.FlatAppearance.BorderSize = 0
        Me.OK_Button.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.OK_Button.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OK_Button.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.OK_Button.Location = New System.Drawing.Point(547, 64)
        Me.OK_Button.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(78, 41)
        Me.OK_Button.TabIndex = 4
        Me.OK_Button.Text = "&OK"
        Me.OK_Button.UseVisualStyleBackColor = True
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.FlatAppearance.BorderSize = 0
        Me.Cancel_Button.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Cancel_Button.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cancel_Button.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.Cancel_Button.Location = New System.Drawing.Point(632, 64)
        Me.Cancel_Button.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(78, 41)
        Me.Cancel_Button.TabIndex = 5
        Me.Cancel_Button.Text = "&Cancel"
        Me.Cancel_Button.UseVisualStyleBackColor = True
        '
        'txtFilterFC
        '
        Me.txtFilterFC.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.txtFilterFC.Location = New System.Drawing.Point(145, 24)
        Me.txtFilterFC.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtFilterFC.Name = "txtFilterFC"
        Me.txtFilterFC.Size = New System.Drawing.Size(214, 25)
        Me.txtFilterFC.TabIndex = 0
        '
        'lblFieldCode
        '
        Me.lblFieldCode.AutoSize = True
        Me.lblFieldCode.Location = New System.Drawing.Point(13, 28)
        Me.lblFieldCode.Name = "lblFieldCode"
        Me.lblFieldCode.Size = New System.Drawing.Size(138, 17)
        Me.lblFieldCode.TabIndex = 9
        Me.lblFieldCode.Text = "Filter by Field Code 1: "
        '
        'lblDescr
        '
        Me.lblDescr.AutoSize = True
        Me.lblDescr.Location = New System.Drawing.Point(13, 76)
        Me.lblDescr.Name = "lblDescr"
        Me.lblDescr.Size = New System.Drawing.Size(135, 17)
        Me.lblDescr.TabIndex = 11
        Me.lblDescr.Text = "Filter by Description:  "
        '
        'txtFilterDescr
        '
        Me.txtFilterDescr.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.txtFilterDescr.Location = New System.Drawing.Point(145, 76)
        Me.txtFilterDescr.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtFilterDescr.Name = "txtFilterDescr"
        Me.txtFilterDescr.Size = New System.Drawing.Size(214, 25)
        Me.txtFilterDescr.TabIndex = 2
        '
        'lbl1
        '
        Me.lbl1.AutoSize = True
        Me.lbl1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl1.ForeColor = System.Drawing.Color.ForestGreen
        Me.lbl1.Location = New System.Drawing.Point(174, 5)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(126, 15)
        Me.lbl1.TabIndex = 12
        Me.lbl1.Text = "No wild cards needed"
        '
        'lblTable
        '
        Me.lblTable.AutoSize = True
        Me.lblTable.Location = New System.Drawing.Point(13, 53)
        Me.lblTable.Name = "lblTable"
        Me.lblTable.Size = New System.Drawing.Size(134, 17)
        Me.lblTable.TabIndex = 14
        Me.lblTable.Text = "Filter by Field Code 2:"
        '
        'txtFilterTable
        '
        Me.txtFilterTable.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.txtFilterTable.Location = New System.Drawing.Point(145, 50)
        Me.txtFilterTable.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtFilterTable.Name = "txtFilterTable"
        Me.txtFilterTable.Size = New System.Drawing.Size(214, 25)
        Me.txtFilterTable.TabIndex = 1
        '
        'gbFilter
        '
        Me.gbFilter.Controls.Add(Me.lblSelection)
        Me.gbFilter.Controls.Add(Me.rbReportItems)
        Me.gbFilter.Controls.Add(Me.rbNone)
        Me.gbFilter.Controls.Add(Me.rbStudySpecific)
        Me.gbFilter.Location = New System.Drawing.Point(367, 5)
        Me.gbFilter.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbFilter.Name = "gbFilter"
        Me.gbFilter.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbFilter.Size = New System.Drawing.Size(164, 120)
        Me.gbFilter.TabIndex = 15
        Me.gbFilter.TabStop = False
        Me.gbFilter.Text = "Hi-level Filter"
        '
        'lblSelection
        '
        Me.lblSelection.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSelection.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblSelection.Location = New System.Drawing.Point(6, 20)
        Me.lblSelection.Name = "lblSelection"
        Me.lblSelection.Size = New System.Drawing.Size(143, 37)
        Me.lblSelection.TabIndex = 3
        Me.lblSelection.Text = "Selection can increase performance"
        '
        'rbReportItems
        '
        Me.rbReportItems.AutoSize = True
        Me.rbReportItems.Location = New System.Drawing.Point(9, 73)
        Me.rbReportItems.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbReportItems.Name = "rbReportItems"
        Me.rbReportItems.Size = New System.Drawing.Size(101, 21)
        Me.rbReportItems.TabIndex = 1
        Me.rbReportItems.Text = "Report Items"
        Me.rbReportItems.UseVisualStyleBackColor = True
        '
        'rbNone
        '
        Me.rbNone.AutoSize = True
        Me.rbNone.Checked = True
        Me.rbNone.Location = New System.Drawing.Point(9, 54)
        Me.rbNone.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbNone.Name = "rbNone"
        Me.rbNone.Size = New System.Drawing.Size(58, 21)
        Me.rbNone.TabIndex = 0
        Me.rbNone.TabStop = True
        Me.rbNone.Text = "None"
        Me.rbNone.UseVisualStyleBackColor = True
        '
        'rbStudySpecific
        '
        Me.rbStudySpecific.AutoSize = True
        Me.rbStudySpecific.Location = New System.Drawing.Point(9, 93)
        Me.rbStudySpecific.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbStudySpecific.Name = "rbStudySpecific"
        Me.rbStudySpecific.Size = New System.Drawing.Size(142, 21)
        Me.rbStudySpecific.TabIndex = 2
        Me.rbStudySpecific.Text = "Table-Specific Items"
        Me.rbStudySpecific.UseVisualStyleBackColor = True
        '
        'cmdCopyAll
        '
        Me.cmdCopyAll.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.cmdCopyAll.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCopyAll.FlatAppearance.BorderSize = 0
        Me.cmdCopyAll.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCopyAll.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCopyAll.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdCopyAll.Location = New System.Drawing.Point(7, 17)
        Me.cmdCopyAll.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCopyAll.Name = "cmdCopyAll"
        Me.cmdCopyAll.Size = New System.Drawing.Size(96, 56)
        Me.cmdCopyAll.TabIndex = 16
        Me.cmdCopyAll.Text = "Copy &All"
        Me.cmdCopyAll.UseVisualStyleBackColor = True
        '
        'gbCopyAll
        '
        Me.gbCopyAll.Controls.Add(Me.cmdCopyAll)
        Me.gbCopyAll.Controls.Add(Me.rbWithoutLabels)
        Me.gbCopyAll.Controls.Add(Me.rbWithLabels)
        Me.gbCopyAll.Location = New System.Drawing.Point(734, 49)
        Me.gbCopyAll.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbCopyAll.Name = "gbCopyAll"
        Me.gbCopyAll.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbCopyAll.Size = New System.Drawing.Size(237, 80)
        Me.gbCopyAll.TabIndex = 17
        Me.gbCopyAll.TabStop = False
        '
        'rbWithoutLabels
        '
        Me.rbWithoutLabels.AutoSize = True
        Me.rbWithoutLabels.Location = New System.Drawing.Point(113, 50)
        Me.rbWithoutLabels.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbWithoutLabels.Name = "rbWithoutLabels"
        Me.rbWithoutLabels.Size = New System.Drawing.Size(112, 21)
        Me.rbWithoutLabels.TabIndex = 18
        Me.rbWithoutLabels.Text = "Without Labels"
        Me.rbWithoutLabels.UseVisualStyleBackColor = True
        '
        'rbWithLabels
        '
        Me.rbWithLabels.AutoSize = True
        Me.rbWithLabels.Checked = True
        Me.rbWithLabels.Location = New System.Drawing.Point(113, 21)
        Me.rbWithLabels.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbWithLabels.Name = "rbWithLabels"
        Me.rbWithLabels.Size = New System.Drawing.Size(93, 21)
        Me.rbWithLabels.TabIndex = 17
        Me.rbWithLabels.TabStop = True
        Me.rbWithLabels.Text = "With Labels"
        Me.rbWithLabels.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.Button1.Location = New System.Drawing.Point(1162, 5)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(78, 41)
        Me.Button1.TabIndex = 18
        Me.Button1.Text = "button"
        Me.Button1.UseVisualStyleBackColor = False
        Me.Button1.Visible = False
        '
        'gbxlblReportTemplateFCstatus
        '
        Me.gbxlblReportTemplateFCstatus.AutoSize = True
        Me.gbxlblReportTemplateFCstatus.Controls.Add(Me.lblReportTemplateFCstatus)
        Me.gbxlblReportTemplateFCstatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 3.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbxlblReportTemplateFCstatus.Location = New System.Drawing.Point(951, 29)
        Me.gbxlblReportTemplateFCstatus.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbxlblReportTemplateFCstatus.Name = "gbxlblReportTemplateFCstatus"
        Me.gbxlblReportTemplateFCstatus.Padding = New System.Windows.Forms.Padding(0)
        Me.gbxlblReportTemplateFCstatus.Size = New System.Drawing.Size(492, 100)
        Me.gbxlblReportTemplateFCstatus.TabIndex = 47
        Me.gbxlblReportTemplateFCstatus.TabStop = False
        '
        'lblReportTemplateFCstatus
        '
        Me.lblReportTemplateFCstatus.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReportTemplateFCstatus.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblReportTemplateFCstatus.Location = New System.Drawing.Point(3, 8)
        Me.lblReportTemplateFCstatus.Name = "lblReportTemplateFCstatus"
        Me.lblReportTemplateFCstatus.Size = New System.Drawing.Size(481, 68)
        Me.lblReportTemplateFCstatus.TabIndex = 44
        Me.lblReportTemplateFCstatus.Text = "Text loaded at frm Load"
        '
        'lblCount
        '
        Me.lblCount.AutoSize = True
        Me.lblCount.Location = New System.Drawing.Point(13, 112)
        Me.lblCount.Name = "lblCount"
        Me.lblCount.Size = New System.Drawing.Size(147, 17)
        Me.lblCount.TabIndex = 49
        Me.lblCount.Text = "Number of Field Codes:"
        '
        'txtCount
        '
        Me.txtCount.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.txtCount.Location = New System.Drawing.Point(160, 104)
        Me.txtCount.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtCount.Name = "txtCount"
        Me.txtCount.ReadOnly = True
        Me.txtCount.Size = New System.Drawing.Size(70, 25)
        Me.txtCount.TabIndex = 48
        Me.txtCount.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'frmFieldCodes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1452, 619)
        Me.ControlBox = False
        Me.Controls.Add(Me.txtFilterTable)
        Me.Controls.Add(Me.txtFilterFC)
        Me.Controls.Add(Me.lblCount)
        Me.Controls.Add(Me.txtCount)
        Me.Controls.Add(Me.txtFilterDescr)
        Me.Controls.Add(Me.gbCopyAll)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.gbxlblReportTemplateFCstatus)
        Me.Controls.Add(Me.cbxGroup)
        Me.Controls.Add(Me.gbFilter)
        Me.Controls.Add(Me.lblTable)
        Me.Controls.Add(Me.lbl1)
        Me.Controls.Add(Me.lblDescr)
        Me.Controls.Add(Me.lblFieldCode)
        Me.Controls.Add(Me.OK_Button)
        Me.Controls.Add(Me.Cancel_Button)
        Me.Controls.Add(Me.lblGroup)
        Me.Controls.Add(Me.panFC)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmFieldCodes"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "  Choose Field Code"
        Me.panFC.ResumeLayout(False)
        CType(Me.dgvFC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbFilter.ResumeLayout(False)
        Me.gbFilter.PerformLayout()
        Me.gbCopyAll.ResumeLayout(False)
        Me.gbCopyAll.PerformLayout()
        Me.gbxlblReportTemplateFCstatus.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents panFC As System.Windows.Forms.Panel
    Friend WithEvents lblGroup As System.Windows.Forms.Label
    Friend WithEvents cbxGroup As System.Windows.Forms.ComboBox
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents txtFilterFC As System.Windows.Forms.TextBox
    Friend WithEvents lblFieldCode As System.Windows.Forms.Label
    Friend WithEvents lblDescr As System.Windows.Forms.Label
    Friend WithEvents txtFilterDescr As System.Windows.Forms.TextBox
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Friend WithEvents lblTable As System.Windows.Forms.Label
    Friend WithEvents txtFilterTable As System.Windows.Forms.TextBox
    Friend WithEvents dgvFC As System.Windows.Forms.DataGridView
    Friend WithEvents gbFilter As System.Windows.Forms.GroupBox
    Friend WithEvents rbStudySpecific As System.Windows.Forms.RadioButton
    Friend WithEvents rbReportItems As System.Windows.Forms.RadioButton
    Friend WithEvents rbNone As System.Windows.Forms.RadioButton
    Friend WithEvents lblSelection As System.Windows.Forms.Label
    Friend WithEvents cmdCopyAll As System.Windows.Forms.Button
    Friend WithEvents gbCopyAll As System.Windows.Forms.GroupBox
    Friend WithEvents rbWithoutLabels As System.Windows.Forms.RadioButton
    Friend WithEvents rbWithLabels As System.Windows.Forms.RadioButton
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents gbxlblReportTemplateFCstatus As System.Windows.Forms.GroupBox
    Friend WithEvents lblReportTemplateFCstatus As System.Windows.Forms.Label
    Friend WithEvents lblCount As System.Windows.Forms.Label
    Friend WithEvents txtCount As System.Windows.Forms.TextBox

End Class
