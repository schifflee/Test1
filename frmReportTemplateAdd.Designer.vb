<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmReportTemplateAdd
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
        Dim DataGridViewCellStyle29 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle30 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle31 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle32 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.panList = New System.Windows.Forms.Panel()
        Me.lblVersions = New System.Windows.Forms.Label()
        Me.lblReports = New System.Windows.Forms.Label()
        Me.dgvReportStatements = New System.Windows.Forms.DataGridView()
        Me.dgvVersions = New System.Windows.Forms.DataGridView()
        Me.lblEnter = New System.Windows.Forms.Label()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.gbChoice = New System.Windows.Forms.GroupBox()
        Me.cmdBrowse = New System.Windows.Forms.Button()
        Me.txtFilePath = New System.Windows.Forms.TextBox()
        Me.rbTemplate = New System.Windows.Forms.RadioButton()
        Me.rbDocument = New System.Windows.Forms.RadioButton()
        Me.rbBlank = New System.Windows.Forms.RadioButton()
        Me.lblWarning = New System.Windows.Forms.Label()
        Me.panList.SuspendLayout()
        CType(Me.dgvReportStatements, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvVersions, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbChoice.SuspendLayout()
        Me.SuspendLayout()
        '
        'panList
        '
        Me.panList.Controls.Add(Me.lblVersions)
        Me.panList.Controls.Add(Me.lblReports)
        Me.panList.Controls.Add(Me.dgvReportStatements)
        Me.panList.Controls.Add(Me.dgvVersions)
        Me.panList.Location = New System.Drawing.Point(48, 190)
        Me.panList.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panList.Name = "panList"
        Me.panList.Size = New System.Drawing.Size(454, 387)
        Me.panList.TabIndex = 67
        '
        'lblVersions
        '
        Me.lblVersions.AutoSize = True
        Me.lblVersions.Location = New System.Drawing.Point(0, 195)
        Me.lblVersions.Name = "lblVersions"
        Me.lblVersions.Size = New System.Drawing.Size(116, 17)
        Me.lblVersions.TabIndex = 69
        Me.lblVersions.Text = "Template Versions"
        '
        'lblReports
        '
        Me.lblReports.AutoSize = True
        Me.lblReports.Location = New System.Drawing.Point(0, 0)
        Me.lblReports.Name = "lblReports"
        Me.lblReports.Size = New System.Drawing.Size(112, 17)
        Me.lblReports.TabIndex = 67
        Me.lblReports.Text = "Report Templates"
        '
        'dgvReportStatements
        '
        Me.dgvReportStatements.AllowUserToAddRows = False
        Me.dgvReportStatements.AllowUserToDeleteRows = False
        Me.dgvReportStatements.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle29.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle29.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle29.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle29.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle29.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle29.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle29.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvReportStatements.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle29
        Me.dgvReportStatements.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle30.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle30.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle30.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle30.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle30.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle30.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle30.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvReportStatements.DefaultCellStyle = DataGridViewCellStyle30
        Me.dgvReportStatements.Location = New System.Drawing.Point(0, 21)
        Me.dgvReportStatements.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvReportStatements.MultiSelect = False
        Me.dgvReportStatements.Name = "dgvReportStatements"
        Me.dgvReportStatements.Size = New System.Drawing.Size(451, 166)
        Me.dgvReportStatements.TabIndex = 1
        '
        'dgvVersions
        '
        Me.dgvVersions.AllowUserToAddRows = False
        Me.dgvVersions.AllowUserToDeleteRows = False
        Me.dgvVersions.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle31.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle31.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle31.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle31.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle31.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle31.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle31.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvVersions.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle31
        Me.dgvVersions.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle32.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle32.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle32.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle32.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle32.NullValue = "No Comments"
        DataGridViewCellStyle32.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle32.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle32.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvVersions.DefaultCellStyle = DataGridViewCellStyle32
        Me.dgvVersions.Location = New System.Drawing.Point(0, 216)
        Me.dgvVersions.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvVersions.MultiSelect = False
        Me.dgvVersions.Name = "dgvVersions"
        Me.dgvVersions.Size = New System.Drawing.Size(451, 166)
        Me.dgvVersions.TabIndex = 2
        '
        'lblEnter
        '
        Me.lblEnter.AutoSize = True
        Me.lblEnter.Location = New System.Drawing.Point(12, 15)
        Me.lblEnter.Name = "lblEnter"
        Me.lblEnter.Size = New System.Drawing.Size(182, 17)
        Me.lblEnter.TabIndex = 68
        Me.lblEnter.Text = "Enter Report Template Name:"
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(200, 12)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(294, 25)
        Me.txtName.TabIndex = 0
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancel.CausesValidation = False
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.FlatAppearance.BorderSize = 0
        Me.cmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCancel.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(424, 18)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 33)
        Me.cmdCancel.TabIndex = 72
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOK.FlatAppearance.BorderSize = 0
        Me.cmdOK.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdOK.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdOK.Location = New System.Drawing.Point(343, 18)
        Me.cmdOK.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(75, 33)
        Me.cmdOK.TabIndex = 71
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Gainsboro
        Me.Button1.FlatAppearance.BorderSize = 0
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Button1.Location = New System.Drawing.Point(445, 13)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 33)
        Me.Button1.TabIndex = 73
        Me.Button1.Text = "Test"
        Me.Button1.UseVisualStyleBackColor = False
        Me.Button1.Visible = False
        '
        'gbChoice
        '
        Me.gbChoice.Controls.Add(Me.lblWarning)
        Me.gbChoice.Controls.Add(Me.cmdBrowse)
        Me.gbChoice.Controls.Add(Me.txtFilePath)
        Me.gbChoice.Controls.Add(Me.cmdCancel)
        Me.gbChoice.Controls.Add(Me.cmdOK)
        Me.gbChoice.Controls.Add(Me.rbTemplate)
        Me.gbChoice.Controls.Add(Me.rbDocument)
        Me.gbChoice.Controls.Add(Me.rbBlank)
        Me.gbChoice.Controls.Add(Me.panList)
        Me.gbChoice.Location = New System.Drawing.Point(15, 49)
        Me.gbChoice.Name = "gbChoice"
        Me.gbChoice.Size = New System.Drawing.Size(505, 589)
        Me.gbChoice.TabIndex = 74
        Me.gbChoice.TabStop = False
        Me.gbChoice.Text = "Create new template based on:"
        '
        'cmdBrowse
        '
        Me.cmdBrowse.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdBrowse.FlatAppearance.BorderSize = 0
        Me.cmdBrowse.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdBrowse.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowse.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdBrowse.Location = New System.Drawing.Point(223, 39)
        Me.cmdBrowse.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdBrowse.Name = "cmdBrowse"
        Me.cmdBrowse.Size = New System.Drawing.Size(75, 33)
        Me.cmdBrowse.TabIndex = 72
        Me.cmdBrowse.Text = "&Browse..."
        Me.cmdBrowse.UseVisualStyleBackColor = False
        '
        'txtFilePath
        '
        Me.txtFilePath.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFilePath.Location = New System.Drawing.Point(48, 100)
        Me.txtFilePath.Multiline = True
        Me.txtFilePath.Name = "txtFilePath"
        Me.txtFilePath.Size = New System.Drawing.Size(451, 56)
        Me.txtFilePath.TabIndex = 69
        '
        'rbTemplate
        '
        Me.rbTemplate.AutoSize = True
        Me.rbTemplate.Location = New System.Drawing.Point(29, 162)
        Me.rbTemplate.Name = "rbTemplate"
        Me.rbTemplate.Size = New System.Drawing.Size(172, 21)
        Me.rbTemplate.TabIndex = 2
        Me.rbTemplate.Text = "Existing Report Template"
        Me.rbTemplate.UseVisualStyleBackColor = True
        '
        'rbDocument
        '
        Me.rbDocument.AutoSize = True
        Me.rbDocument.Location = New System.Drawing.Point(29, 51)
        Me.rbDocument.Name = "rbDocument"
        Me.rbDocument.Size = New System.Drawing.Size(188, 21)
        Me.rbDocument.TabIndex = 1
        Me.rbDocument.Text = "Existing Microsoft Word file"
        Me.rbDocument.UseVisualStyleBackColor = True
        '
        'rbBlank
        '
        Me.rbBlank.AutoSize = True
        Me.rbBlank.Checked = True
        Me.rbBlank.Location = New System.Drawing.Point(29, 24)
        Me.rbBlank.Name = "rbBlank"
        Me.rbBlank.Size = New System.Drawing.Size(119, 21)
        Me.rbBlank.TabIndex = 0
        Me.rbBlank.TabStop = True
        Me.rbBlank.Text = "Blank Document"
        Me.rbBlank.UseVisualStyleBackColor = True
        '
        'lblWarning
        '
        Me.lblWarning.AutoSize = True
        Me.lblWarning.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWarning.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblWarning.Location = New System.Drawing.Point(48, 80)
        Me.lblWarning.Name = "lblWarning"
        Me.lblWarning.Size = New System.Drawing.Size(315, 17)
        Me.lblWarning.TabIndex = 73
        Me.lblWarning.Text = "Note that chosen Word document must be closed"
        '
        'frmReportTemplateAdd
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(539, 652)
        Me.ControlBox = False
        Me.Controls.Add(Me.gbChoice)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.lblEnter)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "frmReportTemplateAdd"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add a new Report Template..."
        Me.panList.ResumeLayout(False)
        Me.panList.PerformLayout()
        CType(Me.dgvReportStatements, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvVersions, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbChoice.ResumeLayout(False)
        Me.gbChoice.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents panList As System.Windows.Forms.Panel
    Friend WithEvents lblVersions As System.Windows.Forms.Label
    Friend WithEvents lblReports As System.Windows.Forms.Label
    Friend WithEvents dgvReportStatements As System.Windows.Forms.DataGridView
    Friend WithEvents dgvVersions As System.Windows.Forms.DataGridView
    Friend WithEvents lblEnter As System.Windows.Forms.Label
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents gbChoice As System.Windows.Forms.GroupBox
    Friend WithEvents cmdBrowse As System.Windows.Forms.Button
    Friend WithEvents txtFilePath As System.Windows.Forms.TextBox
    Friend WithEvents rbTemplate As System.Windows.Forms.RadioButton
    Friend WithEvents rbDocument As System.Windows.Forms.RadioButton
    Friend WithEvents rbBlank As System.Windows.Forms.RadioButton
    Friend WithEvents lblWarning As System.Windows.Forms.Label
End Class
