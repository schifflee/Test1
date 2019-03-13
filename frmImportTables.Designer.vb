<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmImportTables
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmImportTables))
        Me.dgvReportTableConfiguration = New System.Windows.Forms.DataGridView()
        Me.gbRTC = New System.Windows.Forms.GroupBox()
        Me.rbShowAllRTConfig = New System.Windows.Forms.RadioButton()
        Me.rbShowIncludedRTConfig = New System.Windows.Forms.RadioButton()
        Me.cbxStudy = New System.Windows.Forms.ComboBox()
        Me.lblStudy = New System.Windows.Forms.Label()
        Me.txtValue = New System.Windows.Forms.TextBox()
        Me.dgvAdded = New System.Windows.Forms.DataGridView()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.cmdRemove = New System.Windows.Forms.Button()
        Me.lblSource = New System.Windows.Forms.Label()
        Me.lblAdded = New System.Windows.Forms.Label()
        Me.cmdExitNoSave = New System.Windows.Forms.Button()
        Me.cmdExitSave = New System.Windows.Forms.Button()
        Me.dgvTblProps = New System.Windows.Forms.DataGridView()
        Me.panTables = New System.Windows.Forms.Panel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rbAll = New System.Windows.Forms.RadioButton()
        Me.rbIncl = New System.Windows.Forms.RadioButton()
        Me.gbChoose = New System.Windows.Forms.GroupBox()
        Me.rbStudy = New System.Windows.Forms.RadioButton()
        Me.rbTemplate = New System.Windows.Forms.RadioButton()
        CType(Me.dgvReportTableConfiguration, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbRTC.SuspendLayout()
        CType(Me.dgvAdded, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvTblProps, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panTables.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.gbChoose.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvReportTableConfiguration
        '
        Me.dgvReportTableConfiguration.AllowUserToAddRows = False
        Me.dgvReportTableConfiguration.AllowUserToDeleteRows = False
        Me.dgvReportTableConfiguration.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dgvReportTableConfiguration.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvReportTableConfiguration.BackgroundColor = System.Drawing.Color.White
        Me.dgvReportTableConfiguration.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvReportTableConfiguration.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvReportTableConfiguration.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.Padding = New System.Windows.Forms.Padding(0, 2, 0, 2)
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvReportTableConfiguration.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvReportTableConfiguration.Location = New System.Drawing.Point(12, 109)
        Me.dgvReportTableConfiguration.Name = "dgvReportTableConfiguration"
        Me.dgvReportTableConfiguration.ReadOnly = True
        Me.dgvReportTableConfiguration.RowHeadersWidth = 25
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvReportTableConfiguration.RowsDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvReportTableConfiguration.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvReportTableConfiguration.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvReportTableConfiguration.Size = New System.Drawing.Size(432, 406)
        Me.dgvReportTableConfiguration.TabIndex = 100
        '
        'gbRTC
        '
        Me.gbRTC.Controls.Add(Me.rbShowAllRTConfig)
        Me.gbRTC.Controls.Add(Me.rbShowIncludedRTConfig)
        Me.gbRTC.Location = New System.Drawing.Point(12, 3)
        Me.gbRTC.Name = "gbRTC"
        Me.gbRTC.Size = New System.Drawing.Size(212, 60)
        Me.gbRTC.TabIndex = 132
        Me.gbRTC.TabStop = False
        '
        'rbShowAllRTConfig
        '
        Me.rbShowAllRTConfig.AutoSize = True
        Me.rbShowAllRTConfig.BackColor = System.Drawing.Color.Transparent
        Me.rbShowAllRTConfig.Location = New System.Drawing.Point(7, 31)
        Me.rbShowAllRTConfig.Name = "rbShowAllRTConfig"
        Me.rbShowAllRTConfig.Size = New System.Drawing.Size(117, 21)
        Me.rbShowAllRTConfig.TabIndex = 1
        Me.rbShowAllRTConfig.Text = "Show All Tables"
        Me.rbShowAllRTConfig.UseVisualStyleBackColor = False
        '
        'rbShowIncludedRTConfig
        '
        Me.rbShowIncludedRTConfig.AutoSize = True
        Me.rbShowIncludedRTConfig.BackColor = System.Drawing.Color.Transparent
        Me.rbShowIncludedRTConfig.Checked = True
        Me.rbShowIncludedRTConfig.Location = New System.Drawing.Point(7, 9)
        Me.rbShowIncludedRTConfig.Name = "rbShowIncludedRTConfig"
        Me.rbShowIncludedRTConfig.Size = New System.Drawing.Size(180, 21)
        Me.rbShowIncludedRTConfig.TabIndex = 0
        Me.rbShowIncludedRTConfig.TabStop = True
        Me.rbShowIncludedRTConfig.Text = "Show only Included Tables"
        Me.rbShowIncludedRTConfig.UseVisualStyleBackColor = False
        '
        'cbxStudy
        '
        Me.cbxStudy.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.cbxStudy.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cbxStudy.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxStudy.Location = New System.Drawing.Point(186, 103)
        Me.cbxStudy.MaxDropDownItems = 20
        Me.cbxStudy.Name = "cbxStudy"
        Me.cbxStudy.Size = New System.Drawing.Size(360, 25)
        Me.cbxStudy.TabIndex = 0
        '
        'lblStudy
        '
        Me.lblStudy.AutoSize = True
        Me.lblStudy.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStudy.Location = New System.Drawing.Point(12, 110)
        Me.lblStudy.Name = "lblStudy"
        Me.lblStudy.Size = New System.Drawing.Size(168, 17)
        Me.lblStudy.TabIndex = 134
        Me.lblStudy.Text = "Choose a Report Template:"
        '
        'txtValue
        '
        Me.txtValue.Location = New System.Drawing.Point(612, 53)
        Me.txtValue.Name = "txtValue"
        Me.txtValue.Size = New System.Drawing.Size(87, 23)
        Me.txtValue.TabIndex = 135
        Me.txtValue.Visible = False
        '
        'dgvAdded
        '
        Me.dgvAdded.AllowUserToAddRows = False
        Me.dgvAdded.AllowUserToDeleteRows = False
        Me.dgvAdded.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvAdded.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvAdded.BackgroundColor = System.Drawing.Color.White
        Me.dgvAdded.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvAdded.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.dgvAdded.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.Padding = New System.Windows.Forms.Padding(0, 2, 0, 2)
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvAdded.DefaultCellStyle = DataGridViewCellStyle5
        Me.dgvAdded.Location = New System.Drawing.Point(552, 109)
        Me.dgvAdded.Name = "dgvAdded"
        Me.dgvAdded.ReadOnly = True
        Me.dgvAdded.RowHeadersWidth = 25
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvAdded.RowsDefaultCellStyle = DataGridViewCellStyle6
        Me.dgvAdded.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvAdded.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvAdded.Size = New System.Drawing.Size(366, 406)
        Me.dgvAdded.TabIndex = 136
        '
        'cmdAdd
        '
        Me.cmdAdd.ForeColor = System.Drawing.Color.Blue
        Me.cmdAdd.Location = New System.Drawing.Point(450, 157)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(96, 52)
        Me.cmdAdd.TabIndex = 137
        Me.cmdAdd.Text = "--->" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Add"
        Me.cmdAdd.UseVisualStyleBackColor = True
        '
        'cmdRemove
        '
        Me.cmdRemove.ForeColor = System.Drawing.Color.Red
        Me.cmdRemove.Location = New System.Drawing.Point(450, 216)
        Me.cmdRemove.Name = "cmdRemove"
        Me.cmdRemove.Size = New System.Drawing.Size(96, 52)
        Me.cmdRemove.TabIndex = 138
        Me.cmdRemove.Text = "<---" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Remove"
        Me.cmdRemove.UseVisualStyleBackColor = True
        '
        'lblSource
        '
        Me.lblSource.AutoSize = True
        Me.lblSource.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.lblSource.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblSource.Location = New System.Drawing.Point(9, 89)
        Me.lblSource.Name = "lblSource"
        Me.lblSource.Size = New System.Drawing.Size(93, 17)
        Me.lblSource.TabIndex = 139
        Me.lblSource.Text = "Source Tables"
        '
        'lblAdded
        '
        Me.lblAdded.AutoSize = True
        Me.lblAdded.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAdded.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblAdded.Location = New System.Drawing.Point(549, 72)
        Me.lblAdded.Name = "lblAdded"
        Me.lblAdded.Size = New System.Drawing.Size(224, 34)
        Me.lblAdded.TabIndex = 140
        Me.lblAdded.Text = "ExistingTables in Underlying Study" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(Added rows are orange)"
        '
        'cmdExitNoSave
        '
        Me.cmdExitNoSave.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExitNoSave.ForeColor = System.Drawing.Color.Red
        Me.cmdExitNoSave.Location = New System.Drawing.Point(297, 28)
        Me.cmdExitNoSave.Name = "cmdExitNoSave"
        Me.cmdExitNoSave.Size = New System.Drawing.Size(80, 48)
        Me.cmdExitNoSave.TabIndex = 141
        Me.cmdExitNoSave.Text = "&Cancel"
        Me.cmdExitNoSave.UseVisualStyleBackColor = True
        '
        'cmdExitSave
        '
        Me.cmdExitSave.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExitSave.ForeColor = System.Drawing.Color.Blue
        Me.cmdExitSave.Location = New System.Drawing.Point(211, 27)
        Me.cmdExitSave.Name = "cmdExitSave"
        Me.cmdExitSave.Size = New System.Drawing.Size(80, 49)
        Me.cmdExitSave.TabIndex = 142
        Me.cmdExitSave.Text = "&Import"
        Me.cmdExitSave.UseVisualStyleBackColor = True
        '
        'dgvTblProps
        '
        Me.dgvTblProps.AllowUserToAddRows = False
        Me.dgvTblProps.AllowUserToDeleteRows = False
        Me.dgvTblProps.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.dgvTblProps.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvTblProps.BackgroundColor = System.Drawing.Color.White
        Me.dgvTblProps.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvTblProps.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle7
        Me.dgvTblProps.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvTblProps.Location = New System.Drawing.Point(677, 5)
        Me.dgvTblProps.Name = "dgvTblProps"
        Me.dgvTblProps.ReadOnly = True
        Me.dgvTblProps.RowHeadersWidth = 25
        DataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvTblProps.RowsDefaultCellStyle = DataGridViewCellStyle8
        Me.dgvTblProps.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvTblProps.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvTblProps.Size = New System.Drawing.Size(194, 54)
        Me.dgvTblProps.TabIndex = 143
        Me.dgvTblProps.Visible = False
        '
        'panTables
        '
        Me.panTables.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.panTables.Controls.Add(Me.GroupBox1)
        Me.panTables.Controls.Add(Me.gbRTC)
        Me.panTables.Controls.Add(Me.lblSource)
        Me.panTables.Controls.Add(Me.dgvReportTableConfiguration)
        Me.panTables.Controls.Add(Me.cmdRemove)
        Me.panTables.Controls.Add(Me.lblAdded)
        Me.panTables.Controls.Add(Me.cmdAdd)
        Me.panTables.Controls.Add(Me.dgvAdded)
        Me.panTables.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.panTables.Location = New System.Drawing.Point(2, 133)
        Me.panTables.Name = "panTables"
        Me.panTables.Size = New System.Drawing.Size(924, 524)
        Me.panTables.TabIndex = 144
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rbAll)
        Me.GroupBox1.Controls.Add(Me.rbIncl)
        Me.GroupBox1.Location = New System.Drawing.Point(552, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(212, 60)
        Me.GroupBox1.TabIndex = 141
        Me.GroupBox1.TabStop = False
        '
        'rbAll
        '
        Me.rbAll.AutoSize = True
        Me.rbAll.BackColor = System.Drawing.Color.Transparent
        Me.rbAll.Location = New System.Drawing.Point(7, 31)
        Me.rbAll.Name = "rbAll"
        Me.rbAll.Size = New System.Drawing.Size(117, 21)
        Me.rbAll.TabIndex = 1
        Me.rbAll.Text = "Show All Tables"
        Me.rbAll.UseVisualStyleBackColor = False
        '
        'rbIncl
        '
        Me.rbIncl.AutoSize = True
        Me.rbIncl.BackColor = System.Drawing.Color.Transparent
        Me.rbIncl.Checked = True
        Me.rbIncl.Location = New System.Drawing.Point(7, 9)
        Me.rbIncl.Name = "rbIncl"
        Me.rbIncl.Size = New System.Drawing.Size(180, 21)
        Me.rbIncl.TabIndex = 0
        Me.rbIncl.TabStop = True
        Me.rbIncl.Text = "Show only Included Tables"
        Me.rbIncl.UseVisualStyleBackColor = False
        '
        'gbChoose
        '
        Me.gbChoose.Controls.Add(Me.rbStudy)
        Me.gbChoose.Controls.Add(Me.rbTemplate)
        Me.gbChoose.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbChoose.Location = New System.Drawing.Point(14, 14)
        Me.gbChoose.Name = "gbChoose"
        Me.gbChoose.Size = New System.Drawing.Size(191, 79)
        Me.gbChoose.TabIndex = 145
        Me.gbChoose.TabStop = False
        Me.gbChoose.Text = "Choose a Source..."
        '
        'rbStudy
        '
        Me.rbStudy.AutoSize = True
        Me.rbStudy.Location = New System.Drawing.Point(21, 51)
        Me.rbStudy.Name = "rbStudy"
        Me.rbStudy.Size = New System.Drawing.Size(58, 21)
        Me.rbStudy.TabIndex = 1
        Me.rbStudy.Text = "Study"
        Me.rbStudy.UseVisualStyleBackColor = True
        '
        'rbTemplate
        '
        Me.rbTemplate.AutoSize = True
        Me.rbTemplate.Checked = True
        Me.rbTemplate.Location = New System.Drawing.Point(21, 24)
        Me.rbTemplate.Name = "rbTemplate"
        Me.rbTemplate.Size = New System.Drawing.Size(124, 21)
        Me.rbTemplate.TabIndex = 0
        Me.rbTemplate.TabStop = True
        Me.rbTemplate.Text = "Report Template"
        Me.rbTemplate.UseVisualStyleBackColor = True
        '
        'frmImportTables
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(935, 662)
        Me.ControlBox = False
        Me.Controls.Add(Me.gbChoose)
        Me.Controls.Add(Me.dgvTblProps)
        Me.Controls.Add(Me.cmdExitSave)
        Me.Controls.Add(Me.cmdExitNoSave)
        Me.Controls.Add(Me.txtValue)
        Me.Controls.Add(Me.lblStudy)
        Me.Controls.Add(Me.cbxStudy)
        Me.Controls.Add(Me.panTables)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmImportTables"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Import Tables..."
        CType(Me.dgvReportTableConfiguration, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbRTC.ResumeLayout(False)
        Me.gbRTC.PerformLayout()
        CType(Me.dgvAdded, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvTblProps, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panTables.ResumeLayout(False)
        Me.panTables.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.gbChoose.ResumeLayout(False)
        Me.gbChoose.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvReportTableConfiguration As System.Windows.Forms.DataGridView
    Friend WithEvents gbRTC As System.Windows.Forms.GroupBox
    Friend WithEvents rbShowAllRTConfig As System.Windows.Forms.RadioButton
    Friend WithEvents rbShowIncludedRTConfig As System.Windows.Forms.RadioButton
    Friend WithEvents cbxStudy As System.Windows.Forms.ComboBox
    Friend WithEvents lblStudy As System.Windows.Forms.Label
    Friend WithEvents txtValue As System.Windows.Forms.TextBox
    Friend WithEvents dgvAdded As System.Windows.Forms.DataGridView
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents cmdRemove As System.Windows.Forms.Button
    Friend WithEvents lblSource As System.Windows.Forms.Label
    Friend WithEvents lblAdded As System.Windows.Forms.Label
    Friend WithEvents cmdExitNoSave As System.Windows.Forms.Button
    Friend WithEvents cmdExitSave As System.Windows.Forms.Button
    Friend WithEvents dgvTblProps As System.Windows.Forms.DataGridView
    Friend WithEvents panTables As System.Windows.Forms.Panel
    Friend WithEvents gbChoose As System.Windows.Forms.GroupBox
    Friend WithEvents rbStudy As System.Windows.Forms.RadioButton
    Friend WithEvents rbTemplate As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rbAll As System.Windows.Forms.RadioButton
    Friend WithEvents rbIncl As System.Windows.Forms.RadioButton
End Class
