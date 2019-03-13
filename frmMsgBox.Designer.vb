<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMsgBox
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMsgBox))
        Me.lblText = New System.Windows.Forms.Label()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.gb1 = New System.Windows.Forms.GroupBox()
        Me.chkReadOnlyTables = New System.Windows.Forms.CheckBox()
        Me.chkExcludeHeaderFooter = New System.Windows.Forms.CheckBox()
        Me.chkExcludeCoverPages = New System.Windows.Forms.CheckBox()
        Me.chkExcludeEntireTableTitle = New System.Windows.Forms.CheckBox()
        Me.chkExcludeTableTitles = New System.Windows.Forms.CheckBox()
        Me.chkExcludeTableNumbers = New System.Windows.Forms.CheckBox()
        Me.chkShortSampleName = New System.Windows.Forms.CheckBox()
        Me.chkHyperlink = New System.Windows.Forms.CheckBox()
        Me.lblGlobal = New System.Windows.Forms.Label()
        Me.chkDisableWarnings = New System.Windows.Forms.CheckBox()
        Me.chkVerbose = New System.Windows.Forms.CheckBox()
        Me.chkDoPDF = New System.Windows.Forms.CheckBox()
        Me.chkWatermark = New System.Windows.Forms.CheckBox()
        Me.chkFormulas = New System.Windows.Forms.CheckBox()
        Me.cmdSaveSelections = New System.Windows.Forms.Button()
        Me.lblReportTemplate = New System.Windows.Forms.Label()
        Me.cbxReportTemplate = New System.Windows.Forms.ComboBox()
        Me.lblPDF = New System.Windows.Forms.Label()
        Me.lblWord = New System.Windows.Forms.Label()
        Me.lblWatermark = New System.Windows.Forms.Label()
        Me.lblPermissions = New System.Windows.Forms.Label()
        Me.chkAdvSettings = New System.Windows.Forms.CheckBox()
        Me.pan1 = New System.Windows.Forms.Panel()
        Me.lblForcePDF = New System.Windows.Forms.Label()
        Me.panAnalytes = New System.Windows.Forms.Panel()
        Me.lblAnalytes = New System.Windows.Forms.Label()
        Me.cmdDeselect = New System.Windows.Forms.Button()
        Me.cmdSelect = New System.Windows.Forms.Button()
        Me.lbxAnalytes = New System.Windows.Forms.CheckedListBox()
        Me.gb1.SuspendLayout()
        Me.pan1.SuspendLayout()
        Me.panAnalytes.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblText
        '
        Me.lblText.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblText.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblText.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblText.Location = New System.Drawing.Point(3, 3)
        Me.lblText.Name = "lblText"
        Me.lblText.Size = New System.Drawing.Size(685, 76)
        Me.lblText.TabIndex = 0
        Me.lblText.Text = "Label1"
        Me.lblText.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOK.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.cmdOK.FlatAppearance.BorderSize = 0
        Me.cmdOK.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.Blue
        Me.cmdOK.Location = New System.Drawing.Point(67, 61)
        Me.cmdOK.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(85, 43)
        Me.cmdOK.TabIndex = 0
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancel.FlatAppearance.BorderSize = 0
        Me.cmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(218, 63)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(85, 41)
        Me.cmdCancel.TabIndex = 13
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'gb1
        '
        Me.gb1.Controls.Add(Me.chkReadOnlyTables)
        Me.gb1.Controls.Add(Me.chkExcludeHeaderFooter)
        Me.gb1.Controls.Add(Me.chkExcludeCoverPages)
        Me.gb1.Controls.Add(Me.chkExcludeEntireTableTitle)
        Me.gb1.Controls.Add(Me.chkExcludeTableTitles)
        Me.gb1.Controls.Add(Me.chkExcludeTableNumbers)
        Me.gb1.Controls.Add(Me.chkShortSampleName)
        Me.gb1.Controls.Add(Me.chkHyperlink)
        Me.gb1.Controls.Add(Me.lblGlobal)
        Me.gb1.Controls.Add(Me.chkDisableWarnings)
        Me.gb1.Controls.Add(Me.chkVerbose)
        Me.gb1.Controls.Add(Me.chkDoPDF)
        Me.gb1.Controls.Add(Me.chkWatermark)
        Me.gb1.Controls.Add(Me.chkFormulas)
        Me.gb1.Location = New System.Drawing.Point(26, 221)
        Me.gb1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gb1.Name = "gb1"
        Me.gb1.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gb1.Size = New System.Drawing.Size(372, 323)
        Me.gb1.TabIndex = 3
        Me.gb1.TabStop = False
        Me.gb1.Text = "Select Report Options"
        '
        'chkReadOnlyTables
        '
        Me.chkReadOnlyTables.AutoSize = True
        Me.chkReadOnlyTables.Location = New System.Drawing.Point(9, 123)
        Me.chkReadOnlyTables.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkReadOnlyTables.Name = "chkReadOnlyTables"
        Me.chkReadOnlyTables.Size = New System.Drawing.Size(172, 21)
        Me.chkReadOnlyTables.TabIndex = 5
        Me.chkReadOnlyTables.Text = "Create Read-Only Tables"
        Me.chkReadOnlyTables.UseVisualStyleBackColor = True
        '
        'chkExcludeHeaderFooter
        '
        Me.chkExcludeHeaderFooter.AutoSize = True
        Me.chkExcludeHeaderFooter.Location = New System.Drawing.Point(29, 284)
        Me.chkExcludeHeaderFooter.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkExcludeHeaderFooter.Name = "chkExcludeHeaderFooter"
        Me.chkExcludeHeaderFooter.Size = New System.Drawing.Size(260, 21)
        Me.chkExcludeHeaderFooter.TabIndex = 10
        Me.chkExcludeHeaderFooter.Text = "Exclude Example Section Header/Footer"
        Me.chkExcludeHeaderFooter.UseVisualStyleBackColor = True
        '
        'chkExcludeCoverPages
        '
        Me.chkExcludeCoverPages.AutoSize = True
        Me.chkExcludeCoverPages.Location = New System.Drawing.Point(29, 259)
        Me.chkExcludeCoverPages.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkExcludeCoverPages.Name = "chkExcludeCoverPages"
        Me.chkExcludeCoverPages.Size = New System.Drawing.Size(241, 21)
        Me.chkExcludeCoverPages.TabIndex = 9
        Me.chkExcludeCoverPages.Text = "Exclude Example Section Cover Page"
        Me.chkExcludeCoverPages.UseVisualStyleBackColor = True
        '
        'chkExcludeEntireTableTitle
        '
        Me.chkExcludeEntireTableTitle.AutoSize = True
        Me.chkExcludeEntireTableTitle.Location = New System.Drawing.Point(29, 234)
        Me.chkExcludeEntireTableTitle.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkExcludeEntireTableTitle.Name = "chkExcludeEntireTableTitle"
        Me.chkExcludeEntireTableTitle.Size = New System.Drawing.Size(172, 21)
        Me.chkExcludeEntireTableTitle.TabIndex = 8
        Me.chkExcludeEntireTableTitle.Text = "Exclude Entire Table Title"
        Me.chkExcludeEntireTableTitle.UseVisualStyleBackColor = True
        '
        'chkExcludeTableTitles
        '
        Me.chkExcludeTableTitles.AutoSize = True
        Me.chkExcludeTableTitles.Location = New System.Drawing.Point(29, 209)
        Me.chkExcludeTableTitles.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkExcludeTableTitles.Name = "chkExcludeTableTitles"
        Me.chkExcludeTableTitles.Size = New System.Drawing.Size(204, 21)
        Me.chkExcludeTableTitles.TabIndex = 7
        Me.chkExcludeTableTitles.Text = "Exclude Caption in Table Titles"
        Me.chkExcludeTableTitles.UseVisualStyleBackColor = True
        '
        'chkExcludeTableNumbers
        '
        Me.chkExcludeTableNumbers.AutoSize = True
        Me.chkExcludeTableNumbers.Location = New System.Drawing.Point(29, 184)
        Me.chkExcludeTableNumbers.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkExcludeTableNumbers.Name = "chkExcludeTableNumbers"
        Me.chkExcludeTableNumbers.Size = New System.Drawing.Size(249, 21)
        Me.chkExcludeTableNumbers.TabIndex = 6
        Me.chkExcludeTableNumbers.Text = "Exclude Table Numbers in Table Titles"
        Me.chkExcludeTableNumbers.UseVisualStyleBackColor = True
        '
        'chkShortSampleName
        '
        Me.chkShortSampleName.AutoSize = True
        Me.chkShortSampleName.Location = New System.Drawing.Point(147, 143)
        Me.chkShortSampleName.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkShortSampleName.Name = "chkShortSampleName"
        Me.chkShortSampleName.Size = New System.Drawing.Size(204, 21)
        Me.chkShortSampleName.TabIndex = 8
        Me.chkShortSampleName.Text = "Use shortened Sample Names"
        Me.chkShortSampleName.UseVisualStyleBackColor = True
        Me.chkShortSampleName.Visible = False
        '
        'chkHyperlink
        '
        Me.chkHyperlink.Checked = True
        Me.chkHyperlink.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkHyperlink.Location = New System.Drawing.Point(198, 50)
        Me.chkHyperlink.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkHyperlink.Name = "chkHyperlink"
        Me.chkHyperlink.Size = New System.Drawing.Size(239, 30)
        Me.chkHyperlink.TabIndex = 1
        Me.chkHyperlink.Text = "Automatically create table and figure hyperlinks"
        Me.chkHyperlink.UseVisualStyleBackColor = True
        Me.chkHyperlink.Visible = False
        '
        'lblGlobal
        '
        Me.lblGlobal.AutoSize = True
        Me.lblGlobal.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGlobal.ForeColor = System.Drawing.Color.Blue
        Me.lblGlobal.Location = New System.Drawing.Point(6, 157)
        Me.lblGlobal.Name = "lblGlobal"
        Me.lblGlobal.Size = New System.Drawing.Size(92, 13)
        Me.lblGlobal.TabIndex = 7
        Me.lblGlobal.Text = "Global formats:"
        '
        'chkDisableWarnings
        '
        Me.chkDisableWarnings.AutoSize = True
        Me.chkDisableWarnings.Location = New System.Drawing.Point(9, 24)
        Me.chkDisableWarnings.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkDisableWarnings.Name = "chkDisableWarnings"
        Me.chkDisableWarnings.Size = New System.Drawing.Size(218, 21)
        Me.chkDisableWarnings.TabIndex = 1
        Me.chkDisableWarnings.Text = "Disable Warnings and Messages"
        Me.chkDisableWarnings.UseVisualStyleBackColor = True
        '
        'chkVerbose
        '
        Me.chkVerbose.AutoSize = True
        Me.chkVerbose.Location = New System.Drawing.Point(9, 73)
        Me.chkVerbose.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkVerbose.Name = "chkVerbose"
        Me.chkVerbose.Size = New System.Drawing.Size(296, 21)
        Me.chkVerbose.TabIndex = 3
        Me.chkVerbose.Text = "Verbose (show Word doc during preparation)"
        Me.chkVerbose.UseVisualStyleBackColor = True
        '
        'chkDoPDF
        '
        Me.chkDoPDF.AutoSize = True
        Me.chkDoPDF.Location = New System.Drawing.Point(9, 98)
        Me.chkDoPDF.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkDoPDF.Name = "chkDoPDF"
        Me.chkDoPDF.Size = New System.Drawing.Size(324, 21)
        Me.chkDoPDF.TabIndex = 4
        Me.chkDoPDF.Text = "Create as PDF (must of Word(TM) 2007 or greater)"
        Me.chkDoPDF.UseVisualStyleBackColor = True
        '
        'chkWatermark
        '
        Me.chkWatermark.AutoSize = True
        Me.chkWatermark.Location = New System.Drawing.Point(9, 49)
        Me.chkWatermark.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkWatermark.Name = "chkWatermark"
        Me.chkWatermark.Size = New System.Drawing.Size(144, 21)
        Me.chkWatermark.TabIndex = 2
        Me.chkWatermark.Text = "Include a watermark"
        Me.chkWatermark.UseVisualStyleBackColor = True
        '
        'chkFormulas
        '
        Me.chkFormulas.AutoSize = True
        Me.chkFormulas.Location = New System.Drawing.Point(234, 16)
        Me.chkFormulas.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkFormulas.Name = "chkFormulas"
        Me.chkFormulas.Size = New System.Drawing.Size(256, 21)
        Me.chkFormulas.TabIndex = 0
        Me.chkFormulas.Text = "Automatically format chemical formulas"
        Me.chkFormulas.UseVisualStyleBackColor = True
        Me.chkFormulas.Visible = False
        '
        'cmdSaveSelections
        '
        Me.cmdSaveSelections.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSaveSelections.ForeColor = System.Drawing.Color.Blue
        Me.cmdSaveSelections.Location = New System.Drawing.Point(316, 177)
        Me.cmdSaveSelections.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdSaveSelections.Name = "cmdSaveSelections"
        Me.cmdSaveSelections.Size = New System.Drawing.Size(82, 51)
        Me.cmdSaveSelections.TabIndex = 14
        Me.cmdSaveSelections.TabStop = False
        Me.cmdSaveSelections.Text = "&Save Settings"
        Me.cmdSaveSelections.UseVisualStyleBackColor = True
        '
        'lblReportTemplate
        '
        Me.lblReportTemplate.AutoSize = True
        Me.lblReportTemplate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReportTemplate.ForeColor = System.Drawing.Color.Blue
        Me.lblReportTemplate.Location = New System.Drawing.Point(0, 4)
        Me.lblReportTemplate.Name = "lblReportTemplate"
        Me.lblReportTemplate.Size = New System.Drawing.Size(207, 13)
        Me.lblReportTemplate.TabIndex = 6
        Me.lblReportTemplate.Text = "Use the following Report Template:"
        '
        'cbxReportTemplate
        '
        Me.cbxReportTemplate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxReportTemplate.FormattingEnabled = True
        Me.cbxReportTemplate.IntegralHeight = False
        Me.cbxReportTemplate.Location = New System.Drawing.Point(0, 25)
        Me.cbxReportTemplate.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxReportTemplate.Name = "cbxReportTemplate"
        Me.cbxReportTemplate.Size = New System.Drawing.Size(368, 24)
        Me.cbxReportTemplate.TabIndex = 11
        '
        'lblPDF
        '
        Me.lblPDF.AutoSize = True
        Me.lblPDF.Location = New System.Drawing.Point(23, 98)
        Me.lblPDF.Name = "lblPDF"
        Me.lblPDF.Size = New System.Drawing.Size(195, 17)
        Me.lblPDF.TabIndex = 4
        Me.lblPDF.Text = "User is allowed to generate PDF"
        '
        'lblWord
        '
        Me.lblWord.AutoSize = True
        Me.lblWord.Location = New System.Drawing.Point(23, 119)
        Me.lblWord.Name = "lblWord"
        Me.lblWord.Size = New System.Drawing.Size(344, 17)
        Me.lblWord.TabIndex = 5
        Me.lblWord.Text = "User is allowed to generate Microsoft(R) Word document"
        '
        'lblWatermark
        '
        Me.lblWatermark.AutoSize = True
        Me.lblWatermark.Location = New System.Drawing.Point(23, 140)
        Me.lblWatermark.Name = "lblWatermark"
        Me.lblWatermark.Size = New System.Drawing.Size(201, 17)
        Me.lblWatermark.TabIndex = 6
        Me.lblWatermark.Text = "User is forced to use watermarks"
        '
        'lblPermissions
        '
        Me.lblPermissions.AutoSize = True
        Me.lblPermissions.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPermissions.ForeColor = System.Drawing.Color.Blue
        Me.lblPermissions.Location = New System.Drawing.Point(23, 79)
        Me.lblPermissions.Name = "lblPermissions"
        Me.lblPermissions.Size = New System.Drawing.Size(131, 13)
        Me.lblPermissions.TabIndex = 7
        Me.lblPermissions.Text = "Permissions Summary:"
        '
        'chkAdvSettings
        '
        Me.chkAdvSettings.AutoSize = True
        Me.chkAdvSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAdvSettings.ForeColor = System.Drawing.Color.Blue
        Me.chkAdvSettings.Location = New System.Drawing.Point(27, 194)
        Me.chkAdvSettings.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkAdvSettings.Name = "chkAdvSettings"
        Me.chkAdvSettings.Size = New System.Drawing.Size(164, 17)
        Me.chkAdvSettings.TabIndex = 15
        Me.chkAdvSettings.Text = "View Advanced Settings"
        Me.chkAdvSettings.UseVisualStyleBackColor = True
        '
        'pan1
        '
        Me.pan1.Controls.Add(Me.cmdOK)
        Me.pan1.Controls.Add(Me.lblReportTemplate)
        Me.pan1.Controls.Add(Me.cbxReportTemplate)
        Me.pan1.Controls.Add(Me.cmdCancel)
        Me.pan1.Location = New System.Drawing.Point(27, 549)
        Me.pan1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pan1.Name = "pan1"
        Me.pan1.Size = New System.Drawing.Size(371, 106)
        Me.pan1.TabIndex = 16
        Me.pan1.TabStop = True
        '
        'lblForcePDF
        '
        Me.lblForcePDF.AutoSize = True
        Me.lblForcePDF.Location = New System.Drawing.Point(23, 161)
        Me.lblForcePDF.Name = "lblForcePDF"
        Me.lblForcePDF.Size = New System.Drawing.Size(251, 17)
        Me.lblForcePDF.TabIndex = 17
        Me.lblForcePDF.Text = "User is forced to create document as PDF"
        '
        'panAnalytes
        '
        Me.panAnalytes.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.panAnalytes.Controls.Add(Me.lblAnalytes)
        Me.panAnalytes.Controls.Add(Me.cmdDeselect)
        Me.panAnalytes.Controls.Add(Me.cmdSelect)
        Me.panAnalytes.Controls.Add(Me.lbxAnalytes)
        Me.panAnalytes.Location = New System.Drawing.Point(411, 171)
        Me.panAnalytes.Name = "panAnalytes"
        Me.panAnalytes.Size = New System.Drawing.Size(277, 381)
        Me.panAnalytes.TabIndex = 22
        '
        'lblAnalytes
        '
        Me.lblAnalytes.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAnalytes.Location = New System.Drawing.Point(3, 68)
        Me.lblAnalytes.Name = "lblAnalytes"
        Me.lblAnalytes.Size = New System.Drawing.Size(266, 17)
        Me.lblAnalytes.TabIndex = 25
        Me.lblAnalytes.Text = "Select Analytes"
        Me.lblAnalytes.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'cmdDeselect
        '
        Me.cmdDeselect.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDeselect.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDeselect.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.cmdDeselect.Location = New System.Drawing.Point(169, 6)
        Me.cmdDeselect.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdDeselect.Name = "cmdDeselect"
        Me.cmdDeselect.Size = New System.Drawing.Size(100, 51)
        Me.cmdDeselect.TabIndex = 24
        Me.cmdDeselect.Text = "&Deselect All Analytes"
        Me.cmdDeselect.UseVisualStyleBackColor = True
        '
        'cmdSelect
        '
        Me.cmdSelect.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSelect.ForeColor = System.Drawing.Color.Blue
        Me.cmdSelect.Location = New System.Drawing.Point(3, 6)
        Me.cmdSelect.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdSelect.Name = "cmdSelect"
        Me.cmdSelect.Size = New System.Drawing.Size(100, 51)
        Me.cmdSelect.TabIndex = 23
        Me.cmdSelect.Text = "&Select All Analytes"
        Me.cmdSelect.UseVisualStyleBackColor = True
        '
        'lbxAnalytes
        '
        Me.lbxAnalytes.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbxAnalytes.FormattingEnabled = True
        Me.lbxAnalytes.Location = New System.Drawing.Point(3, 88)
        Me.lbxAnalytes.Name = "lbxAnalytes"
        Me.lbxAnalytes.Size = New System.Drawing.Size(266, 284)
        Me.lbxAnalytes.TabIndex = 22
        '
        'frmMsgBox
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(699, 673)
        Me.ControlBox = False
        Me.Controls.Add(Me.panAnalytes)
        Me.Controls.Add(Me.cmdSaveSelections)
        Me.Controls.Add(Me.lblForcePDF)
        Me.Controls.Add(Me.pan1)
        Me.Controls.Add(Me.chkAdvSettings)
        Me.Controls.Add(Me.lblPermissions)
        Me.Controls.Add(Me.lblWatermark)
        Me.Controls.Add(Me.lblWord)
        Me.Controls.Add(Me.lblPDF)
        Me.Controls.Add(Me.gb1)
        Me.Controls.Add(Me.lblText)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "frmMsgBox"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Prepare a Report..."
        Me.gb1.ResumeLayout(False)
        Me.gb1.PerformLayout()
        Me.pan1.ResumeLayout(False)
        Me.pan1.PerformLayout()
        Me.panAnalytes.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblText As System.Windows.Forms.Label
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents gb1 As System.Windows.Forms.GroupBox
    Friend WithEvents chkFormulas As System.Windows.Forms.CheckBox
    Friend WithEvents chkHyperlink As System.Windows.Forms.CheckBox
    Friend WithEvents chkWatermark As System.Windows.Forms.CheckBox
    Friend WithEvents chkDoPDF As System.Windows.Forms.CheckBox
    Friend WithEvents chkVerbose As System.Windows.Forms.CheckBox
    Friend WithEvents cbxReportTemplate As System.Windows.Forms.ComboBox
    Friend WithEvents lblReportTemplate As System.Windows.Forms.Label
    Friend WithEvents chkDisableWarnings As System.Windows.Forms.CheckBox
    Friend WithEvents chkExcludeEntireTableTitle As System.Windows.Forms.CheckBox
    Friend WithEvents chkExcludeTableTitles As System.Windows.Forms.CheckBox
    Friend WithEvents chkExcludeTableNumbers As System.Windows.Forms.CheckBox
    Friend WithEvents chkShortSampleName As System.Windows.Forms.CheckBox
    Friend WithEvents lblGlobal As System.Windows.Forms.Label
    Friend WithEvents chkExcludeCoverPages As System.Windows.Forms.CheckBox
    Friend WithEvents chkExcludeHeaderFooter As System.Windows.Forms.CheckBox
    Friend WithEvents lblPDF As System.Windows.Forms.Label
    Friend WithEvents lblWord As System.Windows.Forms.Label
    Friend WithEvents lblWatermark As System.Windows.Forms.Label
    Friend WithEvents lblPermissions As System.Windows.Forms.Label
    Friend WithEvents chkReadOnlyTables As System.Windows.Forms.CheckBox
    Friend WithEvents cmdSaveSelections As System.Windows.Forms.Button
    Friend WithEvents chkAdvSettings As System.Windows.Forms.CheckBox
    Friend WithEvents pan1 As System.Windows.Forms.Panel
    Friend WithEvents lblForcePDF As System.Windows.Forms.Label
    Friend WithEvents panAnalytes As System.Windows.Forms.Panel
    Friend WithEvents lblAnalytes As System.Windows.Forms.Label
    Friend WithEvents cmdDeselect As System.Windows.Forms.Button
    Friend WithEvents cmdSelect As System.Windows.Forms.Button
    Friend WithEvents lbxAnalytes As System.Windows.Forms.CheckedListBox
End Class
