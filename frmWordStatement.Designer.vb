<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWordStatement
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
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmWordStatement))
        Me.panRefresh = New System.Windows.Forms.Panel()
        Me.lblRefresh = New System.Windows.Forms.Label()
        Me.cmdRefresh = New System.Windows.Forms.Button()
        Me.pan2 = New System.Windows.Forms.Panel()
        Me.lblReadOnly = New System.Windows.Forms.Label()
        Me.cmdPrint1 = New System.Windows.Forms.Button()
        Me.cmdOpenPDF = New System.Windows.Forms.Button()
        Me.cmdExit2 = New System.Windows.Forms.Button()
        Me.cmdWord = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.cmdWord1 = New System.Windows.Forms.Button()
        Me.chkCreateNew = New System.Windows.Forms.CheckBox()
        Me.pan1 = New System.Windows.Forms.Panel()
        Me.panButtons = New System.Windows.Forms.Panel()
        Me.cmdFieldCode = New System.Windows.Forms.Button()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.cmdAddStatement = New System.Windows.Forms.Button()
        Me.cmdPDF = New System.Windows.Forms.Button()
        Me.cmdDeactivateTemplates = New System.Windows.Forms.Button()
        Me.cmdEditTitle = New System.Windows.Forms.Button()
        Me.panList = New System.Windows.Forms.Panel()
        Me.cmdShow = New System.Windows.Forms.Button()
        Me.lblVersions = New System.Windows.Forms.Label()
        Me.lblReports = New System.Windows.Forms.Label()
        Me.dgvReportStatements = New System.Windows.Forms.DataGridView()
        Me.dgvVersions = New System.Windows.Forms.DataGridView()
        Me.panSave = New System.Windows.Forms.Panel()
        Me.cmdOpenExisting = New System.Windows.Forms.Button()
        Me.cmdInsertNew = New System.Windows.Forms.Button()
        Me.cmdCompareDocs = New System.Windows.Forms.Button()
        Me.cmdCancelEdit = New System.Windows.Forms.Button()
        Me.cmdEdit = New System.Windows.Forms.Button()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.panEditReports = New System.Windows.Forms.Panel()
        Me.cmdSaveStatements = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.lblEditTitles = New System.Windows.Forms.Label()
        Me.CHARTITLE = New System.Windows.Forms.TextBox()
        Me.cmdEditStatements = New System.Windows.Forms.Button()
        Me.lblDGV = New System.Windows.Forms.Label()
        Me.panEdraw = New System.Windows.Forms.Panel()
        Me.ov1 = New AxEDOfficeLib.AxEDOffice()
        Me.cmiFieldCode = New System.Windows.Forms.ToolStripMenuItem()
        Me.cmsfrmWord = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.lblSection = New System.Windows.Forms.Label()
        Me.ofd1 = New System.Windows.Forms.OpenFileDialog()
        Me.tSave = New System.Windows.Forms.Timer(Me.components)
        Me.lblBlink = New System.Windows.Forms.Label()
        Me.panRefresh.SuspendLayout()
        Me.pan2.SuspendLayout()
        Me.pan1.SuspendLayout()
        Me.panButtons.SuspendLayout()
        Me.panList.SuspendLayout()
        CType(Me.dgvReportStatements, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvVersions, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panSave.SuspendLayout()
        Me.panEditReports.SuspendLayout()
        Me.panEdraw.SuspendLayout()
        CType(Me.ov1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.cmsfrmWord.SuspendLayout()
        Me.SuspendLayout()
        '
        'panRefresh
        '
        Me.panRefresh.Controls.Add(Me.lblRefresh)
        Me.panRefresh.Controls.Add(Me.cmdRefresh)
        Me.panRefresh.Location = New System.Drawing.Point(834, 485)
        Me.panRefresh.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panRefresh.Name = "panRefresh"
        Me.panRefresh.Size = New System.Drawing.Size(250, 119)
        Me.panRefresh.TabIndex = 65
        '
        'lblRefresh
        '
        Me.lblRefresh.AutoSize = True
        Me.lblRefresh.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRefresh.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.lblRefresh.Location = New System.Drawing.Point(41, 3)
        Me.lblRefresh.MaximumSize = New System.Drawing.Size(174, 191)
        Me.lblRefresh.MinimumSize = New System.Drawing.Size(174, 65)
        Me.lblRefresh.Name = "lblRefresh"
        Me.lblRefresh.Size = New System.Drawing.Size(174, 65)
        Me.lblRefresh.TabIndex = 35
        Me.lblRefresh.Text = "Click Refresh if Word object menus seem to be frozen"
        Me.lblRefresh.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.lblRefresh.Visible = False
        '
        'cmdRefresh
        '
        Me.cmdRefresh.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdRefresh.CausesValidation = False
        Me.cmdRefresh.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdRefresh.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRefresh.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdRefresh.Location = New System.Drawing.Point(62, 78)
        Me.cmdRefresh.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdRefresh.Name = "cmdRefresh"
        Me.cmdRefresh.Size = New System.Drawing.Size(120, 37)
        Me.cmdRefresh.TabIndex = 30
        Me.cmdRefresh.Text = "&Refresh"
        Me.cmdRefresh.UseVisualStyleBackColor = False
        Me.cmdRefresh.Visible = False
        '
        'pan2
        '
        Me.pan2.Controls.Add(Me.lblReadOnly)
        Me.pan2.Controls.Add(Me.cmdPrint1)
        Me.pan2.Controls.Add(Me.cmdOpenPDF)
        Me.pan2.Controls.Add(Me.cmdExit2)
        Me.pan2.Controls.Add(Me.cmdWord)
        Me.pan2.Location = New System.Drawing.Point(277, 410)
        Me.pan2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pan2.Name = "pan2"
        Me.pan2.Size = New System.Drawing.Size(251, 176)
        Me.pan2.TabIndex = 60
        Me.pan2.Visible = False
        '
        'lblReadOnly
        '
        Me.lblReadOnly.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReadOnly.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblReadOnly.Location = New System.Drawing.Point(3, 141)
        Me.lblReadOnly.Name = "lblReadOnly"
        Me.lblReadOnly.Size = New System.Drawing.Size(243, 31)
        Me.lblReadOnly.TabIndex = 62
        Me.lblReadOnly.Text = "View-Only Mode"
        Me.lblReadOnly.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdPrint1
        '
        Me.cmdPrint1.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdPrint1.CausesValidation = False
        Me.cmdPrint1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdPrint1.FlatAppearance.BorderSize = 0
        Me.cmdPrint1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdPrint1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPrint1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdPrint1.Location = New System.Drawing.Point(3, 71)
        Me.cmdPrint1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdPrint1.Name = "cmdPrint1"
        Me.cmdPrint1.Size = New System.Drawing.Size(120, 61)
        Me.cmdPrint1.TabIndex = 57
        Me.cmdPrint1.Text = "&Print"
        Me.cmdPrint1.UseVisualStyleBackColor = False
        '
        'cmdOpenPDF
        '
        Me.cmdOpenPDF.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOpenPDF.FlatAppearance.BorderSize = 0
        Me.cmdOpenPDF.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdOpenPDF.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOpenPDF.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdOpenPDF.Location = New System.Drawing.Point(126, 4)
        Me.cmdOpenPDF.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdOpenPDF.Name = "cmdOpenPDF"
        Me.cmdOpenPDF.Size = New System.Drawing.Size(120, 61)
        Me.cmdOpenPDF.TabIndex = 56
        Me.cmdOpenPDF.Text = "Open as PDF"
        Me.cmdOpenPDF.UseVisualStyleBackColor = False
        Me.cmdOpenPDF.Visible = False
        '
        'cmdExit2
        '
        Me.cmdExit2.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdExit2.CausesValidation = False
        Me.cmdExit2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdExit2.FlatAppearance.BorderSize = 0
        Me.cmdExit2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdExit2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.cmdExit2.Location = New System.Drawing.Point(126, 71)
        Me.cmdExit2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdExit2.Name = "cmdExit2"
        Me.cmdExit2.Size = New System.Drawing.Size(120, 61)
        Me.cmdExit2.TabIndex = 30
        Me.cmdExit2.Text = "&Go Back"
        Me.cmdExit2.UseVisualStyleBackColor = False
        '
        'cmdWord
        '
        Me.cmdWord.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdWord.FlatAppearance.BorderSize = 0
        Me.cmdWord.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdWord.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdWord.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdWord.Location = New System.Drawing.Point(3, 4)
        Me.cmdWord.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdWord.Name = "cmdWord"
        Me.cmdWord.Size = New System.Drawing.Size(120, 61)
        Me.cmdWord.TabIndex = 26
        Me.cmdWord.Text = "Open in &Word"
        Me.cmdWord.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(142, 102)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(58, 44)
        Me.Button1.TabIndex = 57
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'cmdWord1
        '
        Me.cmdWord1.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdWord1.Enabled = False
        Me.cmdWord1.FlatAppearance.BorderSize = 0
        Me.cmdWord1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdWord1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdWord1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdWord1.Location = New System.Drawing.Point(0, 2)
        Me.cmdWord1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdWord1.Name = "cmdWord1"
        Me.cmdWord1.Size = New System.Drawing.Size(78, 45)
        Me.cmdWord1.TabIndex = 55
        Me.cmdWord1.Text = "Open in &Word"
        Me.cmdWord1.UseVisualStyleBackColor = True
        '
        'chkCreateNew
        '
        Me.chkCreateNew.AutoSize = True
        Me.chkCreateNew.Location = New System.Drawing.Point(136, 155)
        Me.chkCreateNew.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkCreateNew.Name = "chkCreateNew"
        Me.chkCreateNew.Size = New System.Drawing.Size(95, 21)
        Me.chkCreateNew.TabIndex = 58
        Me.chkCreateNew.Text = "Create New"
        Me.chkCreateNew.UseVisualStyleBackColor = True
        Me.chkCreateNew.Visible = False
        '
        'pan1
        '
        Me.pan1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.pan1.AutoScroll = True
        Me.pan1.Controls.Add(Me.panButtons)
        Me.pan1.Controls.Add(Me.cmdEditTitle)
        Me.pan1.Controls.Add(Me.panList)
        Me.pan1.Controls.Add(Me.Button1)
        Me.pan1.Controls.Add(Me.panSave)
        Me.pan1.Location = New System.Drawing.Point(1, 2)
        Me.pan1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pan1.Name = "pan1"
        Me.pan1.Size = New System.Drawing.Size(255, 720)
        Me.pan1.TabIndex = 59
        '
        'panButtons
        '
        Me.panButtons.Controls.Add(Me.cmdFieldCode)
        Me.panButtons.Controls.Add(Me.cmdPrint)
        Me.panButtons.Controls.Add(Me.cmdAddStatement)
        Me.panButtons.Controls.Add(Me.cmdPDF)
        Me.panButtons.Controls.Add(Me.cmdWord1)
        Me.panButtons.Controls.Add(Me.cmdDeactivateTemplates)
        Me.panButtons.Location = New System.Drawing.Point(6, 0)
        Me.panButtons.Name = "panButtons"
        Me.panButtons.Size = New System.Drawing.Size(248, 98)
        Me.panButtons.TabIndex = 67
        '
        'cmdFieldCode
        '
        Me.cmdFieldCode.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdFieldCode.Enabled = False
        Me.cmdFieldCode.FlatAppearance.BorderSize = 0
        Me.cmdFieldCode.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdFieldCode.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdFieldCode.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.cmdFieldCode.Location = New System.Drawing.Point(168, 51)
        Me.cmdFieldCode.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdFieldCode.Name = "cmdFieldCode"
        Me.cmdFieldCode.Size = New System.Drawing.Size(78, 45)
        Me.cmdFieldCode.TabIndex = 25
        Me.cmdFieldCode.Text = "Enter &Field Code..."
        Me.cmdFieldCode.UseVisualStyleBackColor = True
        '
        'cmdPrint
        '
        Me.cmdPrint.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdPrint.Enabled = False
        Me.cmdPrint.FlatAppearance.BorderSize = 0
        Me.cmdPrint.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdPrint.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPrint.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdPrint.Location = New System.Drawing.Point(168, 2)
        Me.cmdPrint.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(78, 45)
        Me.cmdPrint.TabIndex = 69
        Me.cmdPrint.Text = "&Print"
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdAddStatement
        '
        Me.cmdAddStatement.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdAddStatement.FlatAppearance.BorderSize = 0
        Me.cmdAddStatement.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdAddStatement.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddStatement.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.cmdAddStatement.Location = New System.Drawing.Point(0, 51)
        Me.cmdAddStatement.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdAddStatement.Name = "cmdAddStatement"
        Me.cmdAddStatement.Size = New System.Drawing.Size(78, 45)
        Me.cmdAddStatement.TabIndex = 36
        Me.cmdAddStatement.Text = "Create &New Template..."
        Me.cmdAddStatement.UseVisualStyleBackColor = True
        '
        'cmdPDF
        '
        Me.cmdPDF.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdPDF.Enabled = False
        Me.cmdPDF.FlatAppearance.BorderSize = 0
        Me.cmdPDF.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdPDF.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPDF.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdPDF.Location = New System.Drawing.Point(84, 2)
        Me.cmdPDF.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdPDF.Name = "cmdPDF"
        Me.cmdPDF.Size = New System.Drawing.Size(78, 45)
        Me.cmdPDF.TabIndex = 68
        Me.cmdPDF.Text = "Open as &PDF"
        Me.cmdPDF.UseVisualStyleBackColor = True
        '
        'cmdDeactivateTemplates
        '
        Me.cmdDeactivateTemplates.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdDeactivateTemplates.FlatAppearance.BorderSize = 0
        Me.cmdDeactivateTemplates.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdDeactivateTemplates.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDeactivateTemplates.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdDeactivateTemplates.Location = New System.Drawing.Point(84, 51)
        Me.cmdDeactivateTemplates.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdDeactivateTemplates.Name = "cmdDeactivateTemplates"
        Me.cmdDeactivateTemplates.Size = New System.Drawing.Size(78, 45)
        Me.cmdDeactivateTemplates.TabIndex = 67
        Me.cmdDeactivateTemplates.Text = "&Template Status..."
        Me.cmdDeactivateTemplates.UseVisualStyleBackColor = True
        '
        'cmdEditTitle
        '
        Me.cmdEditTitle.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdEditTitle.FlatAppearance.BorderSize = 0
        Me.cmdEditTitle.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdEditTitle.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdEditTitle.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdEditTitle.Location = New System.Drawing.Point(6, 102)
        Me.cmdEditTitle.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdEditTitle.Name = "cmdEditTitle"
        Me.cmdEditTitle.Size = New System.Drawing.Size(115, 45)
        Me.cmdEditTitle.TabIndex = 70
        Me.cmdEditTitle.Text = "&Edit Template Title..."
        Me.cmdEditTitle.UseVisualStyleBackColor = True
        '
        'panList
        '
        Me.panList.Controls.Add(Me.cmdShow)
        Me.panList.Controls.Add(Me.lblVersions)
        Me.panList.Controls.Add(Me.lblReports)
        Me.panList.Controls.Add(Me.dgvReportStatements)
        Me.panList.Controls.Add(Me.dgvVersions)
        Me.panList.Location = New System.Drawing.Point(6, 151)
        Me.panList.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panList.Name = "panList"
        Me.panList.Size = New System.Drawing.Size(246, 387)
        Me.panList.TabIndex = 66
        '
        'cmdShow
        '
        Me.cmdShow.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdShow.FlatAppearance.BorderSize = 0
        Me.cmdShow.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdShow.Location = New System.Drawing.Point(142, 190)
        Me.cmdShow.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdShow.Name = "cmdShow"
        Me.cmdShow.Size = New System.Drawing.Size(92, 26)
        Me.cmdShow.TabIndex = 67
        Me.cmdShow.Text = "Show"
        Me.cmdShow.UseVisualStyleBackColor = True
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
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvReportStatements.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvReportStatements.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvReportStatements.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvReportStatements.Location = New System.Drawing.Point(0, 21)
        Me.dgvReportStatements.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvReportStatements.MultiSelect = False
        Me.dgvReportStatements.Name = "dgvReportStatements"
        Me.dgvReportStatements.Size = New System.Drawing.Size(233, 166)
        Me.dgvReportStatements.TabIndex = 33
        '
        'dgvVersions
        '
        Me.dgvVersions.AllowUserToAddRows = False
        Me.dgvVersions.AllowUserToDeleteRows = False
        Me.dgvVersions.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvVersions.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvVersions.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle4.NullValue = "No Comments"
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvVersions.DefaultCellStyle = DataGridViewCellStyle4
        Me.dgvVersions.Location = New System.Drawing.Point(0, 216)
        Me.dgvVersions.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvVersions.MultiSelect = False
        Me.dgvVersions.Name = "dgvVersions"
        Me.dgvVersions.Size = New System.Drawing.Size(233, 166)
        Me.dgvVersions.TabIndex = 68
        '
        'panSave
        '
        Me.panSave.Controls.Add(Me.cmdOpenExisting)
        Me.panSave.Controls.Add(Me.cmdInsertNew)
        Me.panSave.Controls.Add(Me.chkCreateNew)
        Me.panSave.Controls.Add(Me.cmdCompareDocs)
        Me.panSave.Controls.Add(Me.cmdCancelEdit)
        Me.panSave.Controls.Add(Me.cmdEdit)
        Me.panSave.Controls.Add(Me.cmdSave)
        Me.panSave.Controls.Add(Me.cmdExit)
        Me.panSave.Location = New System.Drawing.Point(6, 544)
        Me.panSave.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panSave.Name = "panSave"
        Me.panSave.Size = New System.Drawing.Size(246, 165)
        Me.panSave.TabIndex = 66
        '
        'cmdOpenExisting
        '
        Me.cmdOpenExisting.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOpenExisting.Enabled = False
        Me.cmdOpenExisting.FlatAppearance.BorderSize = 0
        Me.cmdOpenExisting.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdOpenExisting.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOpenExisting.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdOpenExisting.Location = New System.Drawing.Point(6, 56)
        Me.cmdOpenExisting.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdOpenExisting.Name = "cmdOpenExisting"
        Me.cmdOpenExisting.Size = New System.Drawing.Size(110, 45)
        Me.cmdOpenExisting.TabIndex = 142
        Me.cmdOpenExisting.Text = "Open Existing Word Doc"
        Me.cmdOpenExisting.UseVisualStyleBackColor = False
        '
        'cmdInsertNew
        '
        Me.cmdInsertNew.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdInsertNew.Enabled = False
        Me.cmdInsertNew.FlatAppearance.BorderSize = 0
        Me.cmdInsertNew.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdInsertNew.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdInsertNew.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdInsertNew.Location = New System.Drawing.Point(121, 56)
        Me.cmdInsertNew.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdInsertNew.Name = "cmdInsertNew"
        Me.cmdInsertNew.Size = New System.Drawing.Size(110, 45)
        Me.cmdInsertNew.TabIndex = 67
        Me.cmdInsertNew.Text = "Clear Document"
        Me.cmdInsertNew.UseVisualStyleBackColor = False
        '
        'cmdCompareDocs
        '
        Me.cmdCompareDocs.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCompareDocs.FlatAppearance.BorderSize = 0
        Me.cmdCompareDocs.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCompareDocs.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCompareDocs.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdCompareDocs.Location = New System.Drawing.Point(6, 107)
        Me.cmdCompareDocs.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCompareDocs.Name = "cmdCompareDocs"
        Me.cmdCompareDocs.Size = New System.Drawing.Size(110, 45)
        Me.cmdCompareDocs.TabIndex = 141
        Me.cmdCompareDocs.Text = "&Compare Versions"
        Me.cmdCompareDocs.UseVisualStyleBackColor = False
        '
        'cmdCancelEdit
        '
        Me.cmdCancelEdit.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancelEdit.CausesValidation = False
        Me.cmdCancelEdit.Enabled = False
        Me.cmdCancelEdit.FlatAppearance.BorderSize = 0
        Me.cmdCancelEdit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCancelEdit.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancelEdit.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.cmdCancelEdit.Location = New System.Drawing.Point(84, 7)
        Me.cmdCancelEdit.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCancelEdit.Name = "cmdCancelEdit"
        Me.cmdCancelEdit.Size = New System.Drawing.Size(69, 45)
        Me.cmdCancelEdit.TabIndex = 56
        Me.cmdCancelEdit.Text = "&Cancel"
        Me.cmdCancelEdit.UseVisualStyleBackColor = False
        '
        'cmdEdit
        '
        Me.cmdEdit.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdEdit.FlatAppearance.BorderSize = 0
        Me.cmdEdit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdEdit.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdEdit.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdEdit.Location = New System.Drawing.Point(6, 7)
        Me.cmdEdit.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.Size = New System.Drawing.Size(69, 45)
        Me.cmdEdit.TabIndex = 55
        Me.cmdEdit.Text = "Check &Out"
        Me.cmdEdit.UseVisualStyleBackColor = False
        '
        'cmdSave
        '
        Me.cmdSave.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdSave.Enabled = False
        Me.cmdSave.FlatAppearance.BorderSize = 0
        Me.cmdSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdSave.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdSave.Location = New System.Drawing.Point(162, 7)
        Me.cmdSave.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(69, 45)
        Me.cmdSave.TabIndex = 57
        Me.cmdSave.Text = "Check &In"
        Me.cmdSave.UseVisualStyleBackColor = False
        '
        'cmdExit
        '
        Me.cmdExit.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdExit.CausesValidation = False
        Me.cmdExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdExit.FlatAppearance.BorderSize = 0
        Me.cmdExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdExit.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.cmdExit.Location = New System.Drawing.Point(121, 107)
        Me.cmdExit.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(110, 45)
        Me.cmdExit.TabIndex = 58
        Me.cmdExit.Text = "G&o Back"
        Me.cmdExit.UseVisualStyleBackColor = False
        '
        'panEditReports
        '
        Me.panEditReports.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panEditReports.Controls.Add(Me.cmdSaveStatements)
        Me.panEditReports.Controls.Add(Me.cmdCancel)
        Me.panEditReports.Controls.Add(Me.lblEditTitles)
        Me.panEditReports.Controls.Add(Me.CHARTITLE)
        Me.panEditReports.Controls.Add(Me.cmdEditStatements)
        Me.panEditReports.Location = New System.Drawing.Point(835, 626)
        Me.panEditReports.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panEditReports.Name = "panEditReports"
        Me.panEditReports.Size = New System.Drawing.Size(249, 96)
        Me.panEditReports.TabIndex = 45
        Me.panEditReports.Visible = False
        '
        'cmdSaveStatements
        '
        Me.cmdSaveStatements.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdSaveStatements.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdSaveStatements.Enabled = False
        Me.cmdSaveStatements.FlatAppearance.BorderSize = 0
        Me.cmdSaveStatements.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdSaveStatements.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSaveStatements.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdSaveStatements.Location = New System.Drawing.Point(163, 56)
        Me.cmdSaveStatements.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdSaveStatements.Name = "cmdSaveStatements"
        Me.cmdSaveStatements.Size = New System.Drawing.Size(75, 33)
        Me.cmdSaveStatements.TabIndex = 42
        Me.cmdSaveStatements.Text = "&Save"
        Me.cmdSaveStatements.UseVisualStyleBackColor = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancel.CausesValidation = False
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Enabled = False
        Me.cmdCancel.FlatAppearance.BorderSize = 0
        Me.cmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCancel.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(82, 56)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 33)
        Me.cmdCancel.TabIndex = 40
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'lblEditTitles
        '
        Me.lblEditTitles.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEditTitles.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.lblEditTitles.Location = New System.Drawing.Point(5, 4)
        Me.lblEditTitles.MinimumSize = New System.Drawing.Size(167, 0)
        Me.lblEditTitles.Name = "lblEditTitles"
        Me.lblEditTitles.Size = New System.Drawing.Size(239, 17)
        Me.lblEditTitles.TabIndex = 38
        Me.lblEditTitles.Text = "Edit Template Title"
        Me.lblEditTitles.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CHARTITLE
        '
        Me.CHARTITLE.Enabled = False
        Me.CHARTITLE.Location = New System.Drawing.Point(2, 25)
        Me.CHARTITLE.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.CHARTITLE.Name = "CHARTITLE"
        Me.CHARTITLE.ReadOnly = True
        Me.CHARTITLE.Size = New System.Drawing.Size(241, 25)
        Me.CHARTITLE.TabIndex = 41
        '
        'cmdEditStatements
        '
        Me.cmdEditStatements.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdEditStatements.FlatAppearance.BorderSize = 0
        Me.cmdEditStatements.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdEditStatements.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdEditStatements.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdEditStatements.Location = New System.Drawing.Point(2, 56)
        Me.cmdEditStatements.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdEditStatements.Name = "cmdEditStatements"
        Me.cmdEditStatements.Size = New System.Drawing.Size(75, 33)
        Me.cmdEditStatements.TabIndex = 39
        Me.cmdEditStatements.Text = "&Edit"
        Me.cmdEditStatements.UseVisualStyleBackColor = False
        '
        'lblDGV
        '
        Me.lblDGV.AutoSize = True
        Me.lblDGV.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDGV.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblDGV.Location = New System.Drawing.Point(318, 670)
        Me.lblDGV.MaximumSize = New System.Drawing.Size(250, 191)
        Me.lblDGV.Name = "lblDGV"
        Me.lblDGV.Size = New System.Drawing.Size(247, 34)
        Me.lblDGV.TabIndex = 37
        Me.lblDGV.Text = "Save must be clicked before a new row may be selected"
        Me.lblDGV.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.lblDGV.Visible = False
        '
        'panEdraw
        '
        Me.panEdraw.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.panEdraw.Controls.Add(Me.ov1)
        Me.panEdraw.Location = New System.Drawing.Point(262, 44)
        Me.panEdraw.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panEdraw.Name = "panEdraw"
        Me.panEdraw.Size = New System.Drawing.Size(839, 228)
        Me.panEdraw.TabIndex = 64
        '
        'ov1
        '
        Me.ov1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ov1.Enabled = True
        Me.ov1.Location = New System.Drawing.Point(0, 0)
        Me.ov1.Name = "ov1"
        Me.ov1.OcxState = CType(resources.GetObject("ov1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ov1.Size = New System.Drawing.Size(836, 228)
        Me.ov1.TabIndex = 67
        '
        'cmiFieldCode
        '
        Me.cmiFieldCode.Name = "cmiFieldCode"
        Me.cmiFieldCode.Size = New System.Drawing.Size(171, 22)
        Me.cmiFieldCode.Text = "Insert Field Code..."
        '
        'cmsfrmWord
        '
        Me.cmsfrmWord.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.cmiFieldCode})
        Me.cmsfrmWord.Name = "cmsfrmWord"
        Me.cmsfrmWord.Size = New System.Drawing.Size(172, 26)
        '
        'lblSection
        '
        Me.lblSection.AutoSize = True
        Me.lblSection.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblSection.Font = New System.Drawing.Font("Segoe UI", 10.0!)
        Me.lblSection.ForeColor = System.Drawing.Color.White
        Me.lblSection.Location = New System.Drawing.Point(596, 686)
        Me.lblSection.MaximumSize = New System.Drawing.Size(773, 131)
        Me.lblSection.Name = "lblSection"
        Me.lblSection.Size = New System.Drawing.Size(216, 19)
        Me.lblSection.TabIndex = 55
        Me.lblSection.Text = "Report Templates for the Section: "
        Me.lblSection.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.lblSection.Visible = False
        '
        'ofd1
        '
        Me.ofd1.FileName = "OpenFileDialog1"
        '
        'tSave
        '
        '
        'lblBlink
        '
        Me.lblBlink.Font = New System.Drawing.Font("Segoe UI", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBlink.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblBlink.Location = New System.Drawing.Point(432, 311)
        Me.lblBlink.Name = "lblBlink"
        Me.lblBlink.Size = New System.Drawing.Size(632, 213)
        Me.lblBlink.TabIndex = 66
        Me.lblBlink.Text = "The document must be Checked Out in order to access the Report Template."
        Me.lblBlink.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmWordStatement
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1113, 734)
        Me.Controls.Add(Me.pan2)
        Me.Controls.Add(Me.lblDGV)
        Me.Controls.Add(Me.lblBlink)
        Me.Controls.Add(Me.panEditReports)
        Me.Controls.Add(Me.panRefresh)
        Me.Controls.Add(Me.lblSection)
        Me.Controls.Add(Me.panEdraw)
        Me.Controls.Add(Me.pan1)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "frmWordStatement"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmWordStatement_Copy"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.panRefresh.ResumeLayout(False)
        Me.panRefresh.PerformLayout()
        Me.pan2.ResumeLayout(False)
        Me.pan1.ResumeLayout(False)
        Me.panButtons.ResumeLayout(False)
        Me.panList.ResumeLayout(False)
        Me.panList.PerformLayout()
        CType(Me.dgvReportStatements, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvVersions, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panSave.ResumeLayout(False)
        Me.panSave.PerformLayout()
        Me.panEditReports.ResumeLayout(False)
        Me.panEditReports.PerformLayout()
        Me.panEdraw.ResumeLayout(False)
        CType(Me.ov1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.cmsfrmWord.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents panRefresh As System.Windows.Forms.Panel
    Friend WithEvents lblRefresh As System.Windows.Forms.Label
    Friend WithEvents cmdRefresh As System.Windows.Forms.Button
    Friend WithEvents pan2 As System.Windows.Forms.Panel
    Friend WithEvents cmdOpenPDF As System.Windows.Forms.Button
    Friend WithEvents cmdExit2 As System.Windows.Forms.Button
    Friend WithEvents cmdWord As System.Windows.Forms.Button
    Friend WithEvents cmdWord1 As System.Windows.Forms.Button
    Friend WithEvents chkCreateNew As System.Windows.Forms.CheckBox
    Friend WithEvents pan1 As System.Windows.Forms.Panel
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents dgvReportStatements As System.Windows.Forms.DataGridView
    Friend WithEvents panEditReports As System.Windows.Forms.Panel
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents lblEditTitles As System.Windows.Forms.Label
    Friend WithEvents CHARTITLE As System.Windows.Forms.TextBox
    Friend WithEvents cmdEditStatements As System.Windows.Forms.Button
    Friend WithEvents cmdFieldCode As System.Windows.Forms.Button
    Friend WithEvents cmdAddStatement As System.Windows.Forms.Button
    Friend WithEvents lblDGV As System.Windows.Forms.Label
    Friend WithEvents panEdraw As System.Windows.Forms.Panel
    Friend WithEvents cmiFieldCode As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cmsfrmWord As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents lblSection As System.Windows.Forms.Label
    Friend WithEvents panSave As System.Windows.Forms.Panel
    Friend WithEvents panList As System.Windows.Forms.Panel
    'Friend WithEvents ov1 As AxOfficeViewer.AxOfficeViewer
    Friend WithEvents Button1 As System.Windows.Forms.Button
    'Friend WithEvents ov1 As AxEDOfficeLib.AxEDOffice
    Friend WithEvents ofd1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents cmdCancelEdit As System.Windows.Forms.Button
    Friend WithEvents cmdEdit As System.Windows.Forms.Button
    Friend WithEvents cmdSaveStatements As System.Windows.Forms.Button
    Friend WithEvents tSave As System.Windows.Forms.Timer
    Friend WithEvents lblBlink As System.Windows.Forms.Label
    Friend WithEvents cmdCompareDocs As System.Windows.Forms.Button
    Friend WithEvents lblVersions As System.Windows.Forms.Label
    Friend WithEvents dgvVersions As System.Windows.Forms.DataGridView
    Friend WithEvents lblReports As System.Windows.Forms.Label
    Friend WithEvents cmdShow As System.Windows.Forms.Button
    Friend WithEvents cmdOpenExisting As System.Windows.Forms.Button
    Friend WithEvents cmdInsertNew As System.Windows.Forms.Button
    Friend WithEvents cmdDeactivateTemplates As System.Windows.Forms.Button
    Friend WithEvents cmdPDF As System.Windows.Forms.Button
    Friend WithEvents cmdPrint1 As System.Windows.Forms.Button
    Friend WithEvents cmdPrint As System.Windows.Forms.Button
    Friend WithEvents lblReadOnly As System.Windows.Forms.Label
    Friend WithEvents cmdEditTitle As System.Windows.Forms.Button
    Friend WithEvents panButtons As System.Windows.Forms.Panel
    Friend WithEvents ov1 As AxEDOfficeLib.AxEDOffice
    'Friend WithEvents ov1 As AxEDOfficeLib.AxEDOffice
End Class
