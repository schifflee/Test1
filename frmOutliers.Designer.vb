<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOutliers
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOutliers))
        Me.dgvTables = New System.Windows.Forms.DataGridView()
        Me.dgvResults = New System.Windows.Forms.DataGridView()
        Me.gbStatsMethod = New System.Windows.Forms.GroupBox()
        Me.txtStdDev = New System.Windows.Forms.TextBox()
        Me.lblCL = New System.Windows.Forms.Label()
        Me.cbxCL = New System.Windows.Forms.ComboBox()
        Me.lblDixonTable = New System.Windows.Forms.LinkLabel()
        Me.lblGrubbsTable = New System.Windows.Forms.LinkLabel()
        Me.lblStdDev = New System.Windows.Forms.Label()
        Me.rbStdDev = New System.Windows.Forms.RadioButton()
        Me.rbDixon = New System.Windows.Forms.RadioButton()
        Me.rbGrubbs = New System.Windows.Forms.RadioButton()
        Me.lblLegend = New System.Windows.Forms.Label()
        Me.dgvAnalytes = New System.Windows.Forms.DataGridView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dgvSummary = New System.Windows.Forms.DataGridView()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.dgvAllSummary = New System.Windows.Forms.DataGridView()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.cmdReport = New System.Windows.Forms.Button()
        Me.panGTable = New System.Windows.Forms.Panel()
        Me.lblSource = New System.Windows.Forms.Label()
        Me.lblGTable = New System.Windows.Forms.Label()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.dgvGTable = New System.Windows.Forms.DataGridView()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.lblFormLoadProgress = New System.Windows.Forms.Label()
        Me.panFormLoadProgress = New System.Windows.Forms.Panel()
        Me.lblSortSummaryAll = New System.Windows.Forms.Label()
        Me.cbxSortSummaryAll = New System.Windows.Forms.ComboBox()
        Me.Button1 = New System.Windows.Forms.Button()
        CType(Me.dgvTables, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvResults, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbStatsMethod.SuspendLayout()
        CType(Me.dgvAnalytes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvSummary, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvAllSummary, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panGTable.SuspendLayout()
        CType(Me.dgvGTable, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panFormLoadProgress.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvTables
        '
        Me.dgvTables.AllowUserToAddRows = False
        Me.dgvTables.AllowUserToDeleteRows = False
        Me.dgvTables.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvTables.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvTables.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvTables.Location = New System.Drawing.Point(17, 42)
        Me.dgvTables.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvTables.Name = "dgvTables"
        Me.dgvTables.ReadOnly = True
        Me.dgvTables.Size = New System.Drawing.Size(398, 643)
        Me.dgvTables.TabIndex = 0
        '
        'dgvResults
        '
        Me.dgvResults.AllowUserToAddRows = False
        Me.dgvResults.AllowUserToDeleteRows = False
        Me.dgvResults.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvResults.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.dgvResults.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvResults.Location = New System.Drawing.Point(422, 42)
        Me.dgvResults.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvResults.Name = "dgvResults"
        Me.dgvResults.ReadOnly = True
        Me.dgvResults.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.dgvResults.Size = New System.Drawing.Size(566, 453)
        Me.dgvResults.TabIndex = 1
        '
        'gbStatsMethod
        '
        Me.gbStatsMethod.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbStatsMethod.Controls.Add(Me.txtStdDev)
        Me.gbStatsMethod.Controls.Add(Me.lblCL)
        Me.gbStatsMethod.Controls.Add(Me.cbxCL)
        Me.gbStatsMethod.Controls.Add(Me.lblDixonTable)
        Me.gbStatsMethod.Controls.Add(Me.lblGrubbsTable)
        Me.gbStatsMethod.Controls.Add(Me.lblStdDev)
        Me.gbStatsMethod.Controls.Add(Me.rbStdDev)
        Me.gbStatsMethod.Controls.Add(Me.rbDixon)
        Me.gbStatsMethod.Controls.Add(Me.rbGrubbs)
        Me.gbStatsMethod.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbStatsMethod.Location = New System.Drawing.Point(995, 254)
        Me.gbStatsMethod.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbStatsMethod.Name = "gbStatsMethod"
        Me.gbStatsMethod.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbStatsMethod.Size = New System.Drawing.Size(233, 229)
        Me.gbStatsMethod.TabIndex = 2
        Me.gbStatsMethod.TabStop = False
        Me.gbStatsMethod.Text = "Stats Method"
        '
        'txtStdDev
        '
        Me.txtStdDev.Enabled = False
        Me.txtStdDev.Location = New System.Drawing.Point(115, 123)
        Me.txtStdDev.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtStdDev.Name = "txtStdDev"
        Me.txtStdDev.Size = New System.Drawing.Size(46, 21)
        Me.txtStdDev.TabIndex = 3
        Me.txtStdDev.Text = "2"
        Me.txtStdDev.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblCL
        '
        Me.lblCL.Location = New System.Drawing.Point(136, 158)
        Me.lblCL.Name = "lblCL"
        Me.lblCL.Size = New System.Drawing.Size(81, 36)
        Me.lblCL.TabIndex = 8
        Me.lblCL.Text = "Confidence Level"
        Me.lblCL.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'cbxCL
        '
        Me.cbxCL.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxCL.FormattingEnabled = True
        Me.cbxCL.Location = New System.Drawing.Point(139, 197)
        Me.cbxCL.Name = "cbxCL"
        Me.cbxCL.Size = New System.Drawing.Size(75, 23)
        Me.cbxCL.TabIndex = 7
        '
        'lblDixonTable
        '
        Me.lblDixonTable.AutoSize = True
        Me.lblDixonTable.Location = New System.Drawing.Point(34, 187)
        Me.lblDixonTable.Name = "lblDixonTable"
        Me.lblDixonTable.Size = New System.Drawing.Size(100, 15)
        Me.lblDixonTable.TabIndex = 6
        Me.lblDixonTable.TabStop = True
        Me.lblDixonTable.Text = "View Crit R Table"
        '
        'lblGrubbsTable
        '
        Me.lblGrubbsTable.AutoSize = True
        Me.lblGrubbsTable.Location = New System.Drawing.Point(34, 55)
        Me.lblGrubbsTable.Name = "lblGrubbsTable"
        Me.lblGrubbsTable.Size = New System.Drawing.Size(98, 15)
        Me.lblGrubbsTable.TabIndex = 5
        Me.lblGrubbsTable.TabStop = True
        Me.lblGrubbsTable.Text = "View Crit Z Table"
        '
        'lblStdDev
        '
        Me.lblStdDev.AutoSize = True
        Me.lblStdDev.Enabled = False
        Me.lblStdDev.Location = New System.Drawing.Point(36, 127)
        Me.lblStdDev.Name = "lblStdDev"
        Me.lblStdDev.Size = New System.Drawing.Size(62, 15)
        Me.lblStdDev.TabIndex = 4
        Me.lblStdDev.Text = "Std Dev #:"
        '
        'rbStdDev
        '
        Me.rbStdDev.AutoSize = True
        Me.rbStdDev.Location = New System.Drawing.Point(12, 90)
        Me.rbStdDev.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbStdDev.Name = "rbStdDev"
        Me.rbStdDev.Size = New System.Drawing.Size(129, 19)
        Me.rbStdDev.TabIndex = 2
        Me.rbStdDev.Text = "Standard Deviation"
        Me.rbStdDev.UseVisualStyleBackColor = True
        '
        'rbDixon
        '
        Me.rbDixon.AutoSize = True
        Me.rbDixon.Location = New System.Drawing.Point(12, 158)
        Me.rbDixon.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbDixon.Name = "rbDixon"
        Me.rbDixon.Size = New System.Drawing.Size(122, 19)
        Me.rbDixon.TabIndex = 1
        Me.rbDixon.Text = "Dixon's Q Text (2)"
        Me.rbDixon.UseVisualStyleBackColor = True
        '
        'rbGrubbs
        '
        Me.rbGrubbs.AutoSize = True
        Me.rbGrubbs.Checked = True
        Me.rbGrubbs.Location = New System.Drawing.Point(12, 26)
        Me.rbGrubbs.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbGrubbs.Name = "rbGrubbs"
        Me.rbGrubbs.Size = New System.Drawing.Size(109, 19)
        Me.rbGrubbs.TabIndex = 0
        Me.rbGrubbs.TabStop = True
        Me.rbGrubbs.Text = "Grubbs Test (1)"
        Me.rbGrubbs.UseVisualStyleBackColor = True
        '
        'lblLegend
        '
        Me.lblLegend.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblLegend.AutoSize = True
        Me.lblLegend.Location = New System.Drawing.Point(14, 866)
        Me.lblLegend.MaximumSize = New System.Drawing.Size(933, 131)
        Me.lblLegend.MinimumSize = New System.Drawing.Size(933, 0)
        Me.lblLegend.Name = "lblLegend"
        Me.lblLegend.Size = New System.Drawing.Size(933, 17)
        Me.lblLegend.TabIndex = 3
        Me.lblLegend.Text = "Label1"
        '
        'dgvAnalytes
        '
        Me.dgvAnalytes.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvAnalytes.BackgroundColor = System.Drawing.Color.White
        Me.dgvAnalytes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvAnalytes.Location = New System.Drawing.Point(995, 42)
        Me.dgvAnalytes.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvAnalytes.Name = "dgvAnalytes"
        Me.dgvAnalytes.Size = New System.Drawing.Size(239, 196)
        Me.dgvAnalytes.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(419, 504)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(352, 15)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Summary Stats Method Results for the Selected Table"
        '
        'dgvSummary
        '
        Me.dgvSummary.AllowUserToAddRows = False
        Me.dgvSummary.AllowUserToDeleteRows = False
        Me.dgvSummary.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvSummary.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvSummary.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSummary.Location = New System.Drawing.Point(422, 523)
        Me.dgvSummary.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvSummary.Name = "dgvSummary"
        Me.dgvSummary.ReadOnly = True
        Me.dgvSummary.Size = New System.Drawing.Size(566, 162)
        Me.dgvSummary.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(419, 25)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(288, 15)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Stats Method Results for the Selected Table"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(992, 25)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 15)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Analytes"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(14, 25)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(50, 15)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Tables"
        '
        'dgvAllSummary
        '
        Me.dgvAllSummary.AllowUserToAddRows = False
        Me.dgvAllSummary.AllowUserToDeleteRows = False
        Me.dgvAllSummary.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvAllSummary.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvAllSummary.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.dgvAllSummary.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvAllSummary.Location = New System.Drawing.Point(17, 719)
        Me.dgvAllSummary.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvAllSummary.Name = "dgvAllSummary"
        Me.dgvAllSummary.ReadOnly = True
        Me.dgvAllSummary.Size = New System.Drawing.Size(1211, 145)
        Me.dgvAllSummary.TabIndex = 11
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.Label5.Location = New System.Drawing.Point(14, 700)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(295, 15)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Summary Stats Method Results for All Tables"
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblStatus.Location = New System.Drawing.Point(673, 700)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(59, 15)
        Me.lblStatus.TabIndex = 13
        Me.lblStatus.Text = "Status..."
        Me.lblStatus.Visible = False
        '
        'cmdReport
        '
        Me.cmdReport.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdReport.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReport.ForeColor = System.Drawing.Color.Blue
        Me.cmdReport.Location = New System.Drawing.Point(1125, 536)
        Me.cmdReport.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdReport.Name = "cmdReport"
        Me.cmdReport.Size = New System.Drawing.Size(104, 68)
        Me.cmdReport.TabIndex = 14
        Me.cmdReport.Text = "Generate &Report..."
        Me.cmdReport.UseVisualStyleBackColor = True
        '
        'panGTable
        '
        Me.panGTable.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panGTable.Controls.Add(Me.lblSource)
        Me.panGTable.Controls.Add(Me.lblGTable)
        Me.panGTable.Controls.Add(Me.cmdClose)
        Me.panGTable.Controls.Add(Me.dgvGTable)
        Me.panGTable.Location = New System.Drawing.Point(994, 580)
        Me.panGTable.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panGTable.Name = "panGTable"
        Me.panGTable.Size = New System.Drawing.Size(156, 137)
        Me.panGTable.TabIndex = 16
        Me.panGTable.Visible = False
        '
        'lblSource
        '
        Me.lblSource.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSource.Location = New System.Drawing.Point(3, 84)
        Me.lblSource.Name = "lblSource"
        Me.lblSource.Size = New System.Drawing.Size(146, 51)
        Me.lblSource.TabIndex = 11
        Me.lblSource.Text = "Source"
        '
        'lblGTable
        '
        Me.lblGTable.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblGTable.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGTable.Location = New System.Drawing.Point(3, 4)
        Me.lblGTable.Name = "lblGTable"
        Me.lblGTable.Size = New System.Drawing.Size(65, 38)
        Me.lblGTable.TabIndex = 10
        Me.lblGTable.Text = "Tables"
        Me.lblGTable.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'cmdClose
        '
        Me.cmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClose.ForeColor = System.Drawing.Color.Blue
        Me.cmdClose.Location = New System.Drawing.Point(78, 46)
        Me.cmdClose.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(73, 34)
        Me.cmdClose.TabIndex = 1
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'dgvGTable
        '
        Me.dgvGTable.AllowUserToAddRows = False
        Me.dgvGTable.AllowUserToDeleteRows = False
        Me.dgvGTable.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvGTable.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvGTable.BackgroundColor = System.Drawing.Color.White
        Me.dgvGTable.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvGTable.Location = New System.Drawing.Point(3, 46)
        Me.dgvGTable.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvGTable.Name = "dgvGTable"
        Me.dgvGTable.ReadOnly = True
        Me.dgvGTable.Size = New System.Drawing.Size(65, 34)
        Me.dgvGTable.TabIndex = 0
        '
        'cmdExit
        '
        Me.cmdExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdExit.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.cmdExit.ForeColor = System.Drawing.Color.Red
        Me.cmdExit.Location = New System.Drawing.Point(1125, 496)
        Me.cmdExit.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(104, 33)
        Me.cmdExit.TabIndex = 108
        Me.cmdExit.Text = "G&o Back"
        Me.cmdExit.UseVisualStyleBackColor = False
        '
        'lblProgress
        '
        Me.lblProgress.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblProgress.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProgress.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblProgress.Location = New System.Drawing.Point(705, 758)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(485, 104)
        Me.lblProgress.TabIndex = 109
        Me.lblProgress.Text = "Progress..."
        Me.lblProgress.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblFormLoadProgress
        '
        Me.lblFormLoadProgress.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFormLoadProgress.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblFormLoadProgress.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFormLoadProgress.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblFormLoadProgress.Location = New System.Drawing.Point(0, 0)
        Me.lblFormLoadProgress.Name = "lblFormLoadProgress"
        Me.lblFormLoadProgress.Size = New System.Drawing.Size(299, 166)
        Me.lblFormLoadProgress.TabIndex = 110
        Me.lblFormLoadProgress.Text = "Building tables..."
        Me.lblFormLoadProgress.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'panFormLoadProgress
        '
        Me.panFormLoadProgress.Controls.Add(Me.lblFormLoadProgress)
        Me.panFormLoadProgress.Location = New System.Drawing.Point(607, 94)
        Me.panFormLoadProgress.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panFormLoadProgress.Name = "panFormLoadProgress"
        Me.panFormLoadProgress.Size = New System.Drawing.Size(299, 166)
        Me.panFormLoadProgress.TabIndex = 111
        '
        'lblSortSummaryAll
        '
        Me.lblSortSummaryAll.AutoSize = True
        Me.lblSortSummaryAll.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSortSummaryAll.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblSortSummaryAll.Location = New System.Drawing.Point(401, 700)
        Me.lblSortSummaryAll.Name = "lblSortSummaryAll"
        Me.lblSortSummaryAll.Size = New System.Drawing.Size(55, 15)
        Me.lblSortSummaryAll.TabIndex = 112
        Me.lblSortSummaryAll.Text = "Sort by:"
        '
        'cbxSortSummaryAll
        '
        Me.cbxSortSummaryAll.FormattingEnabled = True
        Me.cbxSortSummaryAll.Location = New System.Drawing.Point(472, 692)
        Me.cbxSortSummaryAll.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSortSummaryAll.Name = "cbxSortSummaryAll"
        Me.cbxSortSummaryAll.Size = New System.Drawing.Size(142, 25)
        Me.cbxSortSummaryAll.TabIndex = 113
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(1125, 615)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(84, 33)
        Me.Button1.TabIndex = 114
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'frmOutliers
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1248, 903)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.cbxSortSummaryAll)
        Me.Controls.Add(Me.lblSortSummaryAll)
        Me.Controls.Add(Me.panFormLoadProgress)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.panGTable)
        Me.Controls.Add(Me.cmdReport)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.dgvAllSummary)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dgvSummary)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dgvAnalytes)
        Me.Controls.Add(Me.lblLegend)
        Me.Controls.Add(Me.gbStatsMethod)
        Me.Controls.Add(Me.dgvResults)
        Me.Controls.Add(Me.dgvTables)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmOutliers"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = " Evaluate Outliers"
        CType(Me.dgvTables, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvResults, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbStatsMethod.ResumeLayout(False)
        Me.gbStatsMethod.PerformLayout()
        CType(Me.dgvAnalytes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvSummary, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvAllSummary, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panGTable.ResumeLayout(False)
        CType(Me.dgvGTable, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panFormLoadProgress.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvTables As System.Windows.Forms.DataGridView
    Friend WithEvents dgvResults As System.Windows.Forms.DataGridView
    Friend WithEvents gbStatsMethod As System.Windows.Forms.GroupBox
    Friend WithEvents rbGrubbs As System.Windows.Forms.RadioButton
    Friend WithEvents lblLegend As System.Windows.Forms.Label
    Friend WithEvents dgvAnalytes As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dgvSummary As System.Windows.Forms.DataGridView
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dgvAllSummary As System.Windows.Forms.DataGridView
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents cmdReport As System.Windows.Forms.Button
    Friend WithEvents rbDixon As System.Windows.Forms.RadioButton
    Friend WithEvents lblStdDev As System.Windows.Forms.Label
    Friend WithEvents txtStdDev As System.Windows.Forms.TextBox
    Friend WithEvents rbStdDev As System.Windows.Forms.RadioButton
    Friend WithEvents lblGrubbsTable As System.Windows.Forms.LinkLabel
    Friend WithEvents panGTable As System.Windows.Forms.Panel
    Friend WithEvents dgvGTable As System.Windows.Forms.DataGridView
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents lblGTable As System.Windows.Forms.Label
    Friend WithEvents lblDixonTable As System.Windows.Forms.LinkLabel
    Friend WithEvents lblSource As System.Windows.Forms.Label
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents lblFormLoadProgress As System.Windows.Forms.Label
    Friend WithEvents panFormLoadProgress As System.Windows.Forms.Panel
    Friend WithEvents lblSortSummaryAll As System.Windows.Forms.Label
    Friend WithEvents cbxSortSummaryAll As System.Windows.Forms.ComboBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents cbxCL As System.Windows.Forms.ComboBox
    Friend WithEvents lblCL As System.Windows.Forms.Label
End Class
