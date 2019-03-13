<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDocumentCompare
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDocumentCompare))
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.cbxWRT1 = New System.Windows.Forms.ComboBox()
        Me.lbl1 = New System.Windows.Forms.Label()
        Me.panH1 = New System.Windows.Forms.Panel()
        Me.cbxFR1 = New System.Windows.Forms.ComboBox()
        Me.panH2 = New System.Windows.Forms.Panel()
        Me.cbxFR2 = New System.Windows.Forms.ComboBox()
        Me.lbl2 = New System.Windows.Forms.Label()
        Me.cbxWRT2 = New System.Windows.Forms.ComboBox()
        Me.sc1 = New System.Windows.Forms.SplitContainer()
        Me.ovDC1 = New AxEDOfficeLib.AxEDOffice()
        Me.cmdPaste = New System.Windows.Forms.Button()
        Me.lblNewDoc = New System.Windows.Forms.Label()
        Me.ovDC2 = New AxEDOfficeLib.AxEDOffice()
        Me.cmdCopy = New System.Windows.Forms.Button()
        Me.lblCompareWith = New System.Windows.Forms.Label()
        Me.panProgress1 = New System.Windows.Forms.Panel()
        Me.lblProgress1 = New System.Windows.Forms.Label()
        Me.txtLoadedDocDescription = New System.Windows.Forms.TextBox()
        Me.lblLoadedDoc = New System.Windows.Forms.Label()
        Me.txtStudyID = New System.Windows.Forms.TextBox()
        Me.txtProjectID = New System.Windows.Forms.TextBox()
        Me.lblProject = New System.Windows.Forms.Label()
        Me.txtProject = New System.Windows.Forms.TextBox()
        Me.dgvStudies = New System.Windows.Forms.DataGridView()
        Me.txtStudy = New System.Windows.Forms.TextBox()
        Me.lblStudy = New System.Windows.Forms.Label()
        Me.panStudy = New System.Windows.Forms.Panel()
        Me.cmdBrowseStudy = New System.Windows.Forms.Button()
        Me.dgvProjects = New System.Windows.Forms.DataGridView()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.panWT = New System.Windows.Forms.Panel()
        Me.lblReadOnly = New System.Windows.Forms.Label()
        Me.lblWordTemplate = New System.Windows.Forms.Label()
        Me.txtWordTemplate = New System.Windows.Forms.TextBox()
        Me.txtWSID = New System.Windows.Forms.TextBox()
        Me.pan2 = New System.Windows.Forms.Panel()
        Me.ovDC = New AxEDOfficeLib.AxEDOffice()
        Me.pan1 = New System.Windows.Forms.Panel()
        Me.chkRPane1 = New System.Windows.Forms.CheckBox()
        Me.panProgress = New System.Windows.Forms.Panel()
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.cmdCompare = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.gbCompare = New System.Windows.Forms.GroupBox()
        Me.chkRPane = New System.Windows.Forms.CheckBox()
        Me.rbCompare = New System.Windows.Forms.RadioButton()
        Me.rbSbyS = New System.Windows.Forms.RadioButton()
        Me.panSC1 = New System.Windows.Forms.Panel()
        Me.txtComparedDocDescription = New System.Windows.Forms.TextBox()
        Me.lblComparedDoc = New System.Windows.Forms.Label()
        Me.txtDescr = New System.Windows.Forms.TextBox()
        Me.lblDescr = New System.Windows.Forms.Label()
        Me.txtReportTitle = New System.Windows.Forms.TextBox()
        Me.lblReportTitle = New System.Windows.Forms.Label()
        Me.txtReportNumber = New System.Windows.Forms.TextBox()
        Me.lblReportNumber = New System.Windows.Forms.Label()
        Me.panSave = New System.Windows.Forms.Panel()
        Me.chkShowHilite = New System.Windows.Forms.CheckBox()
        Me.panOpen = New System.Windows.Forms.Panel()
        Me.cmdWord = New System.Windows.Forms.Button()
        Me.cmdOpenPDF = New System.Windows.Forms.Button()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.cmdInsertDocument = New System.Windows.Forms.Button()
        Me.panEdit = New System.Windows.Forms.Panel()
        Me.lblFinalReportLocked = New System.Windows.Forms.Label()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdEdit = New System.Windows.Forms.Button()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.lblInstructions01 = New System.Windows.Forms.Label()
        Me.panOptions = New System.Windows.Forms.Panel()
        Me.gbLoad = New System.Windows.Forms.GroupBox()
        Me.cmdClearCompare = New System.Windows.Forms.Button()
        Me.rbLoadCompare = New System.Windows.Forms.RadioButton()
        Me.rbLoad = New System.Windows.Forms.RadioButton()
        Me.cmdCompareSection = New System.Windows.Forms.Button()
        Me.cmdCompareFinalReport = New System.Windows.Forms.Button()
        Me.lblSection = New System.Windows.Forms.Label()
        Me.lblFinalReport = New System.Windows.Forms.Label()
        Me.dgvSections = New System.Windows.Forms.DataGridView()
        Me.dgvFinalReports = New System.Windows.Forms.DataGridView()
        Me.gbSaveType = New System.Windows.Forms.GroupBox()
        Me.rbSection = New System.Windows.Forms.RadioButton()
        Me.rbFinalReport = New System.Windows.Forms.RadioButton()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.panH1.SuspendLayout()
        Me.panH2.SuspendLayout()
        CType(Me.sc1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.sc1.Panel1.SuspendLayout()
        Me.sc1.Panel2.SuspendLayout()
        Me.sc1.SuspendLayout()
        CType(Me.ovDC1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ovDC2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panProgress1.SuspendLayout()
        CType(Me.dgvStudies, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panStudy.SuspendLayout()
        CType(Me.dgvProjects, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panWT.SuspendLayout()
        Me.pan2.SuspendLayout()
        CType(Me.ovDC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pan1.SuspendLayout()
        Me.panProgress.SuspendLayout()
        Me.gbCompare.SuspendLayout()
        Me.panSC1.SuspendLayout()
        Me.panSave.SuspendLayout()
        Me.panOpen.SuspendLayout()
        Me.panEdit.SuspendLayout()
        Me.gbLoad.SuspendLayout()
        CType(Me.dgvSections, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvFinalReports, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbSaveType.SuspendLayout()
        Me.SuspendLayout()
        '
        'cbxWRT1
        '
        Me.cbxWRT1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxWRT1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxWRT1.FormattingEnabled = True
        Me.cbxWRT1.Location = New System.Drawing.Point(0, 29)
        Me.cbxWRT1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxWRT1.Name = "cbxWRT1"
        Me.cbxWRT1.Size = New System.Drawing.Size(343, 23)
        Me.cbxWRT1.TabIndex = 0
        '
        'lbl1
        '
        Me.lbl1.AutoSize = True
        Me.lbl1.BackColor = System.Drawing.Color.Transparent
        Me.lbl1.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.lbl1.ForeColor = System.Drawing.Color.Black
        Me.lbl1.Location = New System.Drawing.Point(0, 8)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(251, 17)
        Me.lbl1.TabIndex = 1
        Me.lbl1.Text = "Choose A Word Report Template Version"
        '
        'panH1
        '
        Me.panH1.Controls.Add(Me.cbxFR1)
        Me.panH1.Controls.Add(Me.lbl1)
        Me.panH1.Controls.Add(Me.cbxWRT1)
        Me.panH1.Location = New System.Drawing.Point(3, 4)
        Me.panH1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panH1.Name = "panH1"
        Me.panH1.Size = New System.Drawing.Size(353, 57)
        Me.panH1.TabIndex = 3
        '
        'cbxFR1
        '
        Me.cbxFR1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxFR1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxFR1.FormattingEnabled = True
        Me.cbxFR1.Location = New System.Drawing.Point(183, 4)
        Me.cbxFR1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxFR1.Name = "cbxFR1"
        Me.cbxFR1.Size = New System.Drawing.Size(343, 23)
        Me.cbxFR1.TabIndex = 2
        '
        'panH2
        '
        Me.panH2.Controls.Add(Me.cbxFR2)
        Me.panH2.Controls.Add(Me.lbl2)
        Me.panH2.Controls.Add(Me.cbxWRT2)
        Me.panH2.Location = New System.Drawing.Point(364, 4)
        Me.panH2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panH2.Name = "panH2"
        Me.panH2.Size = New System.Drawing.Size(353, 57)
        Me.panH2.TabIndex = 4
        '
        'cbxFR2
        '
        Me.cbxFR2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxFR2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxFR2.FormattingEnabled = True
        Me.cbxFR2.Location = New System.Drawing.Point(101, 4)
        Me.cbxFR2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxFR2.Name = "cbxFR2"
        Me.cbxFR2.Size = New System.Drawing.Size(343, 23)
        Me.cbxFR2.TabIndex = 2
        '
        'lbl2
        '
        Me.lbl2.AutoSize = True
        Me.lbl2.BackColor = System.Drawing.Color.Transparent
        Me.lbl2.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.lbl2.ForeColor = System.Drawing.Color.Black
        Me.lbl2.Location = New System.Drawing.Point(0, 8)
        Me.lbl2.Name = "lbl2"
        Me.lbl2.Size = New System.Drawing.Size(251, 17)
        Me.lbl2.TabIndex = 1
        Me.lbl2.Text = "Choose A Word Report Template Version"
        '
        'cbxWRT2
        '
        Me.cbxWRT2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxWRT2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxWRT2.FormattingEnabled = True
        Me.cbxWRT2.Location = New System.Drawing.Point(0, 29)
        Me.cbxWRT2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxWRT2.Name = "cbxWRT2"
        Me.cbxWRT2.Size = New System.Drawing.Size(343, 23)
        Me.cbxWRT2.TabIndex = 0
        '
        'sc1
        '
        Me.sc1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.sc1.Location = New System.Drawing.Point(0, 135)
        Me.sc1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.sc1.Name = "sc1"
        '
        'sc1.Panel1
        '
        Me.sc1.Panel1.Controls.Add(Me.ovDC1)
        Me.sc1.Panel1.Controls.Add(Me.cmdPaste)
        Me.sc1.Panel1.Controls.Add(Me.lblNewDoc)
        '
        'sc1.Panel2
        '
        Me.sc1.Panel2.Controls.Add(Me.ovDC2)
        Me.sc1.Panel2.Controls.Add(Me.cmdCopy)
        Me.sc1.Panel2.Controls.Add(Me.lblCompareWith)
        Me.sc1.Size = New System.Drawing.Size(727, 300)
        Me.sc1.SplitterDistance = 358
        Me.sc1.SplitterWidth = 5
        Me.sc1.TabIndex = 6
        '
        'ovDC1
        '
        Me.ovDC1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ovDC1.Enabled = True
        Me.ovDC1.Location = New System.Drawing.Point(2, 36)
        Me.ovDC1.Name = "ovDC1"
        Me.ovDC1.OcxState = CType(resources.GetObject("ovDC1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ovDC1.Size = New System.Drawing.Size(355, 262)
        Me.ovDC1.TabIndex = 73
        '
        'cmdPaste
        '
        Me.cmdPaste.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdPaste.FlatAppearance.BorderSize = 0
        Me.cmdPaste.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdPaste.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPaste.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.cmdPaste.Location = New System.Drawing.Point(126, 5)
        Me.cmdPaste.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdPaste.Name = "cmdPaste"
        Me.cmdPaste.Size = New System.Drawing.Size(161, 24)
        Me.cmdPaste.TabIndex = 72
        Me.cmdPaste.Text = "&Paste from Clipboard..."
        Me.cmdPaste.UseVisualStyleBackColor = True
        '
        'lblNewDoc
        '
        Me.lblNewDoc.AutoSize = True
        Me.lblNewDoc.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.lblNewDoc.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNewDoc.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblNewDoc.Location = New System.Drawing.Point(2, 13)
        Me.lblNewDoc.Name = "lblNewDoc"
        Me.lblNewDoc.Size = New System.Drawing.Size(128, 15)
        Me.lblNewDoc.TabIndex = 4
        Me.lblNewDoc.Text = "Loaded Document:"
        '
        'ovDC2
        '
        Me.ovDC2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ovDC2.Enabled = True
        Me.ovDC2.Location = New System.Drawing.Point(2, 36)
        Me.ovDC2.Name = "ovDC2"
        Me.ovDC2.OcxState = CType(resources.GetObject("ovDC2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ovDC2.Size = New System.Drawing.Size(355, 262)
        Me.ovDC2.TabIndex = 74
        '
        'cmdCopy
        '
        Me.cmdCopy.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCopy.FlatAppearance.BorderSize = 0
        Me.cmdCopy.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCopy.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCopy.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.cmdCopy.Location = New System.Drawing.Point(97, 5)
        Me.cmdCopy.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCopy.Name = "cmdCopy"
        Me.cmdCopy.Size = New System.Drawing.Size(260, 24)
        Me.cmdCopy.TabIndex = 73
        Me.cmdCopy.Text = "&Paste Selection to Loaded Document..."
        Me.cmdCopy.UseVisualStyleBackColor = True
        '
        'lblCompareWith
        '
        Me.lblCompareWith.AutoSize = True
        Me.lblCompareWith.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.lblCompareWith.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompareWith.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblCompareWith.Location = New System.Drawing.Point(0, 13)
        Me.lblCompareWith.Name = "lblCompareWith"
        Me.lblCompareWith.Size = New System.Drawing.Size(101, 15)
        Me.lblCompareWith.TabIndex = 3
        Me.lblCompareWith.Text = "Compare With:"
        '
        'panProgress1
        '
        Me.panProgress1.Controls.Add(Me.lblProgress1)
        Me.panProgress1.Location = New System.Drawing.Point(1068, 221)
        Me.panProgress1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panProgress1.Name = "panProgress1"
        Me.panProgress1.Size = New System.Drawing.Size(251, 67)
        Me.panProgress1.TabIndex = 19
        '
        'lblProgress1
        '
        Me.lblProgress1.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProgress1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.lblProgress1.Location = New System.Drawing.Point(3, 0)
        Me.lblProgress1.Name = "lblProgress1"
        Me.lblProgress1.Size = New System.Drawing.Size(202, 42)
        Me.lblProgress1.TabIndex = 0
        Me.lblProgress1.Text = "Label4"
        Me.lblProgress1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtLoadedDocDescription
        '
        Me.txtLoadedDocDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLoadedDocDescription.Location = New System.Drawing.Point(126, 78)
        Me.txtLoadedDocDescription.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtLoadedDocDescription.Name = "txtLoadedDocDescription"
        Me.txtLoadedDocDescription.ReadOnly = True
        Me.txtLoadedDocDescription.Size = New System.Drawing.Size(595, 25)
        Me.txtLoadedDocDescription.TabIndex = 12
        '
        'lblLoadedDoc
        '
        Me.lblLoadedDoc.AutoSize = True
        Me.lblLoadedDoc.BackColor = System.Drawing.Color.Transparent
        Me.lblLoadedDoc.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.lblLoadedDoc.ForeColor = System.Drawing.Color.Black
        Me.lblLoadedDoc.Location = New System.Drawing.Point(2, 81)
        Me.lblLoadedDoc.Name = "lblLoadedDoc"
        Me.lblLoadedDoc.Size = New System.Drawing.Size(82, 17)
        Me.lblLoadedDoc.TabIndex = 2
        Me.lblLoadedDoc.Text = "Loaded Doc:"
        '
        'txtStudyID
        '
        Me.txtStudyID.Location = New System.Drawing.Point(13, 418)
        Me.txtStudyID.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtStudyID.Name = "txtStudyID"
        Me.txtStudyID.Size = New System.Drawing.Size(107, 25)
        Me.txtStudyID.TabIndex = 8
        Me.txtStudyID.Text = "txtStudyID"
        Me.txtStudyID.Visible = False
        '
        'txtProjectID
        '
        Me.txtProjectID.Location = New System.Drawing.Point(13, 389)
        Me.txtProjectID.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtProjectID.Name = "txtProjectID"
        Me.txtProjectID.Size = New System.Drawing.Size(107, 25)
        Me.txtProjectID.TabIndex = 7
        Me.txtProjectID.Text = "txtProjectID"
        Me.txtProjectID.Visible = False
        '
        'lblProject
        '
        Me.lblProject.AutoSize = True
        Me.lblProject.Location = New System.Drawing.Point(1, 0)
        Me.lblProject.Name = "lblProject"
        Me.lblProject.Size = New System.Drawing.Size(51, 17)
        Me.lblProject.TabIndex = 5
        Me.lblProject.Text = "Project:"
        '
        'txtProject
        '
        Me.txtProject.Location = New System.Drawing.Point(1, 21)
        Me.txtProject.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtProject.Multiline = True
        Me.txtProject.Name = "txtProject"
        Me.txtProject.ReadOnly = True
        Me.txtProject.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtProject.Size = New System.Drawing.Size(275, 79)
        Me.txtProject.TabIndex = 5
        '
        'dgvStudies
        '
        Me.dgvStudies.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvStudies.Location = New System.Drawing.Point(14, 642)
        Me.dgvStudies.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvStudies.Name = "dgvStudies"
        Me.dgvStudies.Size = New System.Drawing.Size(278, 167)
        Me.dgvStudies.TabIndex = 5
        Me.dgvStudies.UseWaitCursor = True
        Me.dgvStudies.Visible = False
        '
        'txtStudy
        '
        Me.txtStudy.Location = New System.Drawing.Point(1, 139)
        Me.txtStudy.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtStudy.Multiline = True
        Me.txtStudy.Name = "txtStudy"
        Me.txtStudy.ReadOnly = True
        Me.txtStudy.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtStudy.Size = New System.Drawing.Size(275, 79)
        Me.txtStudy.TabIndex = 7
        '
        'lblStudy
        '
        Me.lblStudy.AutoSize = True
        Me.lblStudy.Location = New System.Drawing.Point(2, 118)
        Me.lblStudy.Name = "lblStudy"
        Me.lblStudy.Size = New System.Drawing.Size(43, 17)
        Me.lblStudy.TabIndex = 8
        Me.lblStudy.Text = "Study:"
        '
        'panStudy
        '
        Me.panStudy.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panStudy.Controls.Add(Me.cmdBrowseStudy)
        Me.panStudy.Controls.Add(Me.lblProject)
        Me.panStudy.Controls.Add(Me.txtStudy)
        Me.panStudy.Controls.Add(Me.txtProject)
        Me.panStudy.Controls.Add(Me.lblStudy)
        Me.panStudy.Location = New System.Drawing.Point(13, 84)
        Me.panStudy.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panStudy.Name = "panStudy"
        Me.panStudy.Size = New System.Drawing.Size(283, 228)
        Me.panStudy.TabIndex = 5
        Me.panStudy.Visible = False
        '
        'cmdBrowseStudy
        '
        Me.cmdBrowseStudy.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseStudy.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdBrowseStudy.Location = New System.Drawing.Point(192, 4)
        Me.cmdBrowseStudy.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdBrowseStudy.Name = "cmdBrowseStudy"
        Me.cmdBrowseStudy.Size = New System.Drawing.Size(85, 56)
        Me.cmdBrowseStudy.TabIndex = 5
        Me.cmdBrowseStudy.Text = "Select Study..."
        Me.cmdBrowseStudy.UseVisualStyleBackColor = True
        Me.cmdBrowseStudy.Visible = False
        '
        'dgvProjects
        '
        Me.dgvProjects.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvProjects.Location = New System.Drawing.Point(13, 452)
        Me.dgvProjects.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvProjects.Name = "dgvProjects"
        Me.dgvProjects.Size = New System.Drawing.Size(279, 167)
        Me.dgvProjects.TabIndex = 9
        Me.dgvProjects.UseWaitCursor = True
        Me.dgvProjects.Visible = False
        '
        'cmdExit
        '
        Me.cmdExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.cmdExit.Location = New System.Drawing.Point(13, 13)
        Me.cmdExit.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(128, 50)
        Me.cmdExit.TabIndex = 10
        Me.cmdExit.Text = "&Go Back"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'panWT
        '
        Me.panWT.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panWT.Controls.Add(Me.lblReadOnly)
        Me.panWT.Controls.Add(Me.lblWordTemplate)
        Me.panWT.Controls.Add(Me.txtWordTemplate)
        Me.panWT.Location = New System.Drawing.Point(687, 183)
        Me.panWT.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panWT.Name = "panWT"
        Me.panWT.Size = New System.Drawing.Size(278, 132)
        Me.panWT.TabIndex = 4
        '
        'lblReadOnly
        '
        Me.lblReadOnly.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.lblReadOnly.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReadOnly.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblReadOnly.Location = New System.Drawing.Point(3, 82)
        Me.lblReadOnly.Name = "lblReadOnly"
        Me.lblReadOnly.Size = New System.Drawing.Size(270, 47)
        Me.lblReadOnly.TabIndex = 69
        Me.lblReadOnly.Text = "The documents shown here are read-only"
        Me.lblReadOnly.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblWordTemplate
        '
        Me.lblWordTemplate.AutoSize = True
        Me.lblWordTemplate.BackColor = System.Drawing.Color.Transparent
        Me.lblWordTemplate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWordTemplate.ForeColor = System.Drawing.Color.Black
        Me.lblWordTemplate.Location = New System.Drawing.Point(1, 0)
        Me.lblWordTemplate.Name = "lblWordTemplate"
        Me.lblWordTemplate.Size = New System.Drawing.Size(128, 15)
        Me.lblWordTemplate.TabIndex = 5
        Me.lblWordTemplate.Text = "Active Word Template:"
        '
        'txtWordTemplate
        '
        Me.txtWordTemplate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWordTemplate.Location = New System.Drawing.Point(3, 20)
        Me.txtWordTemplate.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtWordTemplate.Multiline = True
        Me.txtWordTemplate.Name = "txtWordTemplate"
        Me.txtWordTemplate.ReadOnly = True
        Me.txtWordTemplate.Size = New System.Drawing.Size(270, 59)
        Me.txtWordTemplate.TabIndex = 4
        '
        'txtWSID
        '
        Me.txtWSID.Location = New System.Drawing.Point(13, 359)
        Me.txtWSID.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtWSID.Name = "txtWSID"
        Me.txtWSID.Size = New System.Drawing.Size(107, 25)
        Me.txtWSID.TabIndex = 11
        Me.txtWSID.Text = "txtWSID"
        Me.txtWSID.Visible = False
        '
        'pan2
        '
        Me.pan2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pan2.AutoScroll = True
        Me.pan2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pan2.Controls.Add(Me.ovDC)
        Me.pan2.Location = New System.Drawing.Point(0, 86)
        Me.pan2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pan2.Name = "pan2"
        Me.pan2.Size = New System.Drawing.Size(5214, 54)
        Me.pan2.TabIndex = 12
        '
        'ovDC
        '
        Me.ovDC.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ovDC.Enabled = True
        Me.ovDC.Location = New System.Drawing.Point(0, 0)
        Me.ovDC.Name = "ovDC"
        Me.ovDC.OcxState = CType(resources.GetObject("ovDC.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ovDC.Size = New System.Drawing.Size(5212, 52)
        Me.ovDC.TabIndex = 74
        '
        'pan1
        '
        Me.pan1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pan1.AutoScroll = True
        Me.pan1.Controls.Add(Me.chkRPane1)
        Me.pan1.Controls.Add(Me.panProgress)
        Me.pan1.Controls.Add(Me.cmdCompare)
        Me.pan1.Controls.Add(Me.panH1)
        Me.pan1.Controls.Add(Me.pan2)
        Me.pan1.Controls.Add(Me.panH2)
        Me.pan1.Location = New System.Drawing.Point(595, 13)
        Me.pan1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pan1.Name = "pan1"
        Me.pan1.Size = New System.Drawing.Size(707, 162)
        Me.pan1.TabIndex = 13
        '
        'chkRPane1
        '
        Me.chkRPane1.AutoSize = True
        Me.chkRPane1.Checked = True
        Me.chkRPane1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkRPane1.Location = New System.Drawing.Point(3, 65)
        Me.chkRPane1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkRPane1.Name = "chkRPane1"
        Me.chkRPane1.Size = New System.Drawing.Size(134, 21)
        Me.chkRPane1.TabIndex = 17
        Me.chkRPane1.Text = "Show Review Pane"
        Me.chkRPane1.UseVisualStyleBackColor = True
        '
        'panProgress
        '
        Me.panProgress.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panProgress.Controls.Add(Me.lblProgress)
        Me.panProgress.Location = New System.Drawing.Point(1193, 4)
        Me.panProgress.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panProgress.Name = "panProgress"
        Me.panProgress.Size = New System.Drawing.Size(158, 66)
        Me.panProgress.TabIndex = 14
        '
        'lblProgress
        '
        Me.lblProgress.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProgress.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.lblProgress.Location = New System.Drawing.Point(55, 9)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(202, 42)
        Me.lblProgress.TabIndex = 0
        Me.lblProgress.Text = "lblProgress"
        Me.lblProgress.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdCompare
        '
        Me.cmdCompare.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCompare.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdCompare.Location = New System.Drawing.Point(724, 5)
        Me.cmdCompare.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCompare.Name = "cmdCompare"
        Me.cmdCompare.Size = New System.Drawing.Size(85, 56)
        Me.cmdCompare.TabIndex = 13
        Me.cmdCompare.Text = "&Compare"
        Me.cmdCompare.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(127, 369)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(222, 17)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "StudyDoc ID_TBLWORDSTATEMENTS"
        Me.Label1.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(127, 399)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(120, 17)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "Watson PROJECTID"
        Me.Label2.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(127, 428)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(107, 17)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "Watson STUDYID"
        Me.Label3.Visible = False
        '
        'gbCompare
        '
        Me.gbCompare.Controls.Add(Me.chkRPane)
        Me.gbCompare.Controls.Add(Me.rbCompare)
        Me.gbCompare.Controls.Add(Me.rbSbyS)
        Me.gbCompare.Enabled = False
        Me.gbCompare.Location = New System.Drawing.Point(93, 23)
        Me.gbCompare.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbCompare.Name = "gbCompare"
        Me.gbCompare.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbCompare.Size = New System.Drawing.Size(180, 106)
        Me.gbCompare.TabIndex = 14
        Me.gbCompare.TabStop = False
        Me.gbCompare.Text = "Compare Type"
        '
        'chkRPane
        '
        Me.chkRPane.AutoSize = True
        Me.chkRPane.Checked = True
        Me.chkRPane.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkRPane.Enabled = False
        Me.chkRPane.Location = New System.Drawing.Point(37, 76)
        Me.chkRPane.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkRPane.Name = "chkRPane"
        Me.chkRPane.Size = New System.Drawing.Size(134, 21)
        Me.chkRPane.TabIndex = 16
        Me.chkRPane.Text = "Show Review Pane"
        Me.chkRPane.UseVisualStyleBackColor = True
        Me.chkRPane.Visible = False
        '
        'rbCompare
        '
        Me.rbCompare.AutoSize = True
        Me.rbCompare.Location = New System.Drawing.Point(14, 50)
        Me.rbCompare.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbCompare.Name = "rbCompare"
        Me.rbCompare.Size = New System.Drawing.Size(117, 21)
        Me.rbCompare.TabIndex = 15
        Me.rbCompare.Text = "Word Compare"
        Me.rbCompare.UseVisualStyleBackColor = True
        '
        'rbSbyS
        '
        Me.rbSbyS.AutoSize = True
        Me.rbSbyS.Checked = True
        Me.rbSbyS.Location = New System.Drawing.Point(14, 25)
        Me.rbSbyS.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbSbyS.Name = "rbSbyS"
        Me.rbSbyS.Size = New System.Drawing.Size(98, 21)
        Me.rbSbyS.TabIndex = 14
        Me.rbSbyS.TabStop = True
        Me.rbSbyS.Text = "Side by Side"
        Me.rbSbyS.UseVisualStyleBackColor = True
        '
        'panSC1
        '
        Me.panSC1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.panSC1.AutoScroll = True
        Me.panSC1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panSC1.Controls.Add(Me.txtComparedDocDescription)
        Me.panSC1.Controls.Add(Me.lblComparedDoc)
        Me.panSC1.Controls.Add(Me.txtDescr)
        Me.panSC1.Controls.Add(Me.lblDescr)
        Me.panSC1.Controls.Add(Me.txtReportTitle)
        Me.panSC1.Controls.Add(Me.txtLoadedDocDescription)
        Me.panSC1.Controls.Add(Me.lblReportTitle)
        Me.panSC1.Controls.Add(Me.lblLoadedDoc)
        Me.panSC1.Controls.Add(Me.txtReportNumber)
        Me.panSC1.Controls.Add(Me.lblReportNumber)
        Me.panSC1.Controls.Add(Me.sc1)
        Me.panSC1.Location = New System.Drawing.Point(590, 342)
        Me.panSC1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panSC1.Name = "panSC1"
        Me.panSC1.Size = New System.Drawing.Size(729, 441)
        Me.panSC1.TabIndex = 17
        '
        'txtComparedDocDescription
        '
        Me.txtComparedDocDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtComparedDocDescription.Location = New System.Drawing.Point(126, 107)
        Me.txtComparedDocDescription.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtComparedDocDescription.Name = "txtComparedDocDescription"
        Me.txtComparedDocDescription.ReadOnly = True
        Me.txtComparedDocDescription.Size = New System.Drawing.Size(595, 25)
        Me.txtComparedDocDescription.TabIndex = 22
        '
        'lblComparedDoc
        '
        Me.lblComparedDoc.AutoSize = True
        Me.lblComparedDoc.BackColor = System.Drawing.Color.Transparent
        Me.lblComparedDoc.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.lblComparedDoc.ForeColor = System.Drawing.Color.Black
        Me.lblComparedDoc.Location = New System.Drawing.Point(2, 110)
        Me.lblComparedDoc.Name = "lblComparedDoc"
        Me.lblComparedDoc.Size = New System.Drawing.Size(100, 17)
        Me.lblComparedDoc.TabIndex = 20
        Me.lblComparedDoc.Text = "Compared Doc:"
        '
        'txtDescr
        '
        Me.txtDescr.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDescr.Location = New System.Drawing.Point(551, 50)
        Me.txtDescr.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtDescr.Name = "txtDescr"
        Me.txtDescr.ReadOnly = True
        Me.txtDescr.Size = New System.Drawing.Size(170, 25)
        Me.txtDescr.TabIndex = 16
        Me.txtDescr.Visible = False
        '
        'lblDescr
        '
        Me.lblDescr.AutoSize = True
        Me.lblDescr.BackColor = System.Drawing.Color.Transparent
        Me.lblDescr.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.lblDescr.ForeColor = System.Drawing.Color.Black
        Me.lblDescr.Location = New System.Drawing.Point(393, 53)
        Me.lblDescr.Name = "lblDescr"
        Me.lblDescr.Size = New System.Drawing.Size(152, 17)
        Me.lblDescr.TabIndex = 15
        Me.lblDescr.Text = "Loaded Doc Description:"
        Me.lblDescr.Visible = False
        '
        'txtReportTitle
        '
        Me.txtReportTitle.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtReportTitle.Location = New System.Drawing.Point(126, 4)
        Me.txtReportTitle.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtReportTitle.Multiline = True
        Me.txtReportTitle.Name = "txtReportTitle"
        Me.txtReportTitle.ReadOnly = True
        Me.txtReportTitle.Size = New System.Drawing.Size(595, 42)
        Me.txtReportTitle.TabIndex = 14
        Me.txtReportTitle.Text = "txtReportTitle"
        '
        'lblReportTitle
        '
        Me.lblReportTitle.BackColor = System.Drawing.Color.Transparent
        Me.lblReportTitle.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReportTitle.ForeColor = System.Drawing.Color.Black
        Me.lblReportTitle.Location = New System.Drawing.Point(2, 4)
        Me.lblReportTitle.Name = "lblReportTitle"
        Me.lblReportTitle.Size = New System.Drawing.Size(117, 25)
        Me.lblReportTitle.TabIndex = 13
        Me.lblReportTitle.Text = "Report Title:"
        '
        'txtReportNumber
        '
        Me.txtReportNumber.Location = New System.Drawing.Point(126, 50)
        Me.txtReportNumber.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtReportNumber.Name = "txtReportNumber"
        Me.txtReportNumber.ReadOnly = True
        Me.txtReportNumber.Size = New System.Drawing.Size(226, 25)
        Me.txtReportNumber.TabIndex = 12
        Me.txtReportNumber.Text = "txtReportNumber"
        '
        'lblReportNumber
        '
        Me.lblReportNumber.BackColor = System.Drawing.Color.Transparent
        Me.lblReportNumber.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.lblReportNumber.ForeColor = System.Drawing.Color.Black
        Me.lblReportNumber.Location = New System.Drawing.Point(2, 53)
        Me.lblReportNumber.Name = "lblReportNumber"
        Me.lblReportNumber.Size = New System.Drawing.Size(117, 25)
        Me.lblReportNumber.TabIndex = 7
        Me.lblReportNumber.Text = "Study Name:"
        '
        'panSave
        '
        Me.panSave.AutoScroll = True
        Me.panSave.Controls.Add(Me.chkShowHilite)
        Me.panSave.Controls.Add(Me.panOpen)
        Me.panSave.Controls.Add(Me.panEdit)
        Me.panSave.Controls.Add(Me.lblInstructions01)
        Me.panSave.Controls.Add(Me.panOptions)
        Me.panSave.Controls.Add(Me.gbLoad)
        Me.panSave.Controls.Add(Me.cmdCompareSection)
        Me.panSave.Controls.Add(Me.cmdCompareFinalReport)
        Me.panSave.Controls.Add(Me.lblSection)
        Me.panSave.Controls.Add(Me.lblFinalReport)
        Me.panSave.Controls.Add(Me.dgvSections)
        Me.panSave.Controls.Add(Me.dgvFinalReports)
        Me.panSave.Location = New System.Drawing.Point(298, 13)
        Me.panSave.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panSave.Name = "panSave"
        Me.panSave.Size = New System.Drawing.Size(286, 847)
        Me.panSave.TabIndex = 61
        Me.panSave.Visible = False
        '
        'chkShowHilite
        '
        Me.chkShowHilite.AutoSize = True
        Me.chkShowHilite.Checked = True
        Me.chkShowHilite.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowHilite.Location = New System.Drawing.Point(14, 320)
        Me.chkShowHilite.Name = "chkShowHilite"
        Me.chkShowHilite.Size = New System.Drawing.Size(167, 21)
        Me.chkShowHilite.TabIndex = 16
        Me.chkShowHilite.Text = "Hilight Editable Sections"
        Me.chkShowHilite.UseVisualStyleBackColor = True
        Me.chkShowHilite.Visible = False
        '
        'panOpen
        '
        Me.panOpen.Controls.Add(Me.cmdWord)
        Me.panOpen.Controls.Add(Me.cmdOpenPDF)
        Me.panOpen.Controls.Add(Me.cmdPrint)
        Me.panOpen.Controls.Add(Me.cmdInsertDocument)
        Me.panOpen.Location = New System.Drawing.Point(0, 0)
        Me.panOpen.Name = "panOpen"
        Me.panOpen.Size = New System.Drawing.Size(263, 103)
        Me.panOpen.TabIndex = 62
        '
        'cmdWord
        '
        Me.cmdWord.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdWord.FlatAppearance.BorderSize = 0
        Me.cmdWord.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdWord.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdWord.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdWord.Location = New System.Drawing.Point(0, 0)
        Me.cmdWord.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdWord.Name = "cmdWord"
        Me.cmdWord.Size = New System.Drawing.Size(128, 50)
        Me.cmdWord.TabIndex = 26
        Me.cmdWord.Text = "Open in &Word"
        Me.cmdWord.UseVisualStyleBackColor = True
        '
        'cmdOpenPDF
        '
        Me.cmdOpenPDF.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOpenPDF.FlatAppearance.BorderSize = 0
        Me.cmdOpenPDF.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdOpenPDF.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOpenPDF.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdOpenPDF.Location = New System.Drawing.Point(131, 0)
        Me.cmdOpenPDF.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdOpenPDF.Name = "cmdOpenPDF"
        Me.cmdOpenPDF.Size = New System.Drawing.Size(128, 50)
        Me.cmdOpenPDF.TabIndex = 27
        Me.cmdOpenPDF.Text = "Open as PDF"
        Me.cmdOpenPDF.UseVisualStyleBackColor = True
        '
        'cmdPrint
        '
        Me.cmdPrint.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdPrint.FlatAppearance.BorderSize = 0
        Me.cmdPrint.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdPrint.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPrint.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdPrint.Location = New System.Drawing.Point(0, 53)
        Me.cmdPrint.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(128, 50)
        Me.cmdPrint.TabIndex = 30
        Me.cmdPrint.Text = "&Print"
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdInsertDocument
        '
        Me.cmdInsertDocument.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdInsertDocument.FlatAppearance.BorderSize = 0
        Me.cmdInsertDocument.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdInsertDocument.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdInsertDocument.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdInsertDocument.Location = New System.Drawing.Point(131, 53)
        Me.cmdInsertDocument.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdInsertDocument.Name = "cmdInsertDocument"
        Me.cmdInsertDocument.Size = New System.Drawing.Size(128, 50)
        Me.cmdInsertDocument.TabIndex = 29
        Me.cmdInsertDocument.Text = "&Insert Document"
        Me.cmdInsertDocument.UseVisualStyleBackColor = True
        Me.cmdInsertDocument.Visible = False
        '
        'panEdit
        '
        Me.panEdit.Controls.Add(Me.lblFinalReportLocked)
        Me.panEdit.Controls.Add(Me.cmdCancel)
        Me.panEdit.Controls.Add(Me.cmdEdit)
        Me.panEdit.Controls.Add(Me.cmdSave)
        Me.panEdit.Location = New System.Drawing.Point(0, 111)
        Me.panEdit.Name = "panEdit"
        Me.panEdit.Size = New System.Drawing.Size(263, 62)
        Me.panEdit.TabIndex = 62
        '
        'lblFinalReportLocked
        '
        Me.lblFinalReportLocked.AutoSize = True
        Me.lblFinalReportLocked.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFinalReportLocked.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblFinalReportLocked.Location = New System.Drawing.Point(130, 37)
        Me.lblFinalReportLocked.Name = "lblFinalReportLocked"
        Me.lblFinalReportLocked.Size = New System.Drawing.Size(130, 17)
        Me.lblFinalReportLocked.TabIndex = 16
        Me.lblFinalReportLocked.Text = "Final Report Locked"
        Me.lblFinalReportLocked.Visible = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancel.Enabled = False
        Me.cmdCancel.FlatAppearance.BorderSize = 0
        Me.cmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCancel.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(0, 31)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(128, 28)
        Me.cmdCancel.TabIndex = 72
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdEdit
        '
        Me.cmdEdit.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdEdit.FlatAppearance.BorderSize = 0
        Me.cmdEdit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdEdit.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdEdit.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdEdit.Location = New System.Drawing.Point(0, 0)
        Me.cmdEdit.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.Size = New System.Drawing.Size(128, 28)
        Me.cmdEdit.TabIndex = 71
        Me.cmdEdit.Text = "&Edit"
        Me.cmdEdit.UseVisualStyleBackColor = True
        '
        'cmdSave
        '
        Me.cmdSave.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdSave.Enabled = False
        Me.cmdSave.FlatAppearance.BorderSize = 0
        Me.cmdSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdSave.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.cmdSave.Location = New System.Drawing.Point(131, 0)
        Me.cmdSave.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(128, 28)
        Me.cmdSave.TabIndex = 30
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'lblInstructions01
        '
        Me.lblInstructions01.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.lblInstructions01.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInstructions01.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblInstructions01.Location = New System.Drawing.Point(0, 314)
        Me.lblInstructions01.Name = "lblInstructions01"
        Me.lblInstructions01.Size = New System.Drawing.Size(279, 47)
        Me.lblInstructions01.TabIndex = 68
        Me.lblInstructions01.Text = "After document is saved, user may perform further actions."
        Me.lblInstructions01.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'panOptions
        '
        Me.panOptions.Location = New System.Drawing.Point(10, 776)
        Me.panOptions.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panOptions.Name = "panOptions"
        Me.panOptions.Size = New System.Drawing.Size(260, 65)
        Me.panOptions.TabIndex = 70
        '
        'gbLoad
        '
        Me.gbLoad.Controls.Add(Me.cmdClearCompare)
        Me.gbLoad.Controls.Add(Me.gbCompare)
        Me.gbLoad.Controls.Add(Me.rbLoadCompare)
        Me.gbLoad.Controls.Add(Me.rbLoad)
        Me.gbLoad.Location = New System.Drawing.Point(0, 182)
        Me.gbLoad.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbLoad.Name = "gbLoad"
        Me.gbLoad.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbLoad.Size = New System.Drawing.Size(279, 133)
        Me.gbLoad.TabIndex = 69
        Me.gbLoad.TabStop = False
        Me.gbLoad.Text = "Load or Compare"
        Me.gbLoad.Visible = False
        '
        'cmdClearCompare
        '
        Me.cmdClearCompare.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdClearCompare.FlatAppearance.BorderSize = 0
        Me.cmdClearCompare.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdClearCompare.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClearCompare.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdClearCompare.Location = New System.Drawing.Point(14, 75)
        Me.cmdClearCompare.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdClearCompare.Name = "cmdClearCompare"
        Me.cmdClearCompare.Size = New System.Drawing.Size(73, 52)
        Me.cmdClearCompare.TabIndex = 72
        Me.cmdClearCompare.Text = "&Clear" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Compare"
        Me.cmdClearCompare.UseVisualStyleBackColor = True
        '
        'rbLoadCompare
        '
        Me.rbLoadCompare.AutoSize = True
        Me.rbLoadCompare.Location = New System.Drawing.Point(14, 52)
        Me.rbLoadCompare.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbLoadCompare.Name = "rbLoadCompare"
        Me.rbLoadCompare.Size = New System.Drawing.Size(80, 21)
        Me.rbLoadCompare.TabIndex = 15
        Me.rbLoadCompare.Text = "Compare"
        Me.rbLoadCompare.UseVisualStyleBackColor = True
        '
        'rbLoad
        '
        Me.rbLoad.AutoSize = True
        Me.rbLoad.Checked = True
        Me.rbLoad.Location = New System.Drawing.Point(14, 26)
        Me.rbLoad.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbLoad.Name = "rbLoad"
        Me.rbLoad.Size = New System.Drawing.Size(55, 21)
        Me.rbLoad.TabIndex = 14
        Me.rbLoad.TabStop = True
        Me.rbLoad.Text = "Load"
        Me.rbLoad.UseVisualStyleBackColor = True
        '
        'cmdCompareSection
        '
        Me.cmdCompareSection.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCompareSection.Enabled = False
        Me.cmdCompareSection.FlatAppearance.BorderSize = 0
        Me.cmdCompareSection.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCompareSection.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCompareSection.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdCompareSection.Location = New System.Drawing.Point(159, 566)
        Me.cmdCompareSection.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCompareSection.Name = "cmdCompareSection"
        Me.cmdCompareSection.Size = New System.Drawing.Size(120, 30)
        Me.cmdCompareSection.TabIndex = 67
        Me.cmdCompareSection.Text = "Load-->"
        Me.cmdCompareSection.UseVisualStyleBackColor = True
        '
        'cmdCompareFinalReport
        '
        Me.cmdCompareFinalReport.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCompareFinalReport.Enabled = False
        Me.cmdCompareFinalReport.FlatAppearance.BorderSize = 0
        Me.cmdCompareFinalReport.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCompareFinalReport.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCompareFinalReport.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdCompareFinalReport.Location = New System.Drawing.Point(159, 361)
        Me.cmdCompareFinalReport.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCompareFinalReport.Name = "cmdCompareFinalReport"
        Me.cmdCompareFinalReport.Size = New System.Drawing.Size(120, 30)
        Me.cmdCompareFinalReport.TabIndex = 66
        Me.cmdCompareFinalReport.Text = "Load-->"
        Me.cmdCompareFinalReport.UseVisualStyleBackColor = True
        '
        'lblSection
        '
        Me.lblSection.AutoSize = True
        Me.lblSection.BackColor = System.Drawing.Color.Transparent
        Me.lblSection.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.lblSection.ForeColor = System.Drawing.Color.Black
        Me.lblSection.Location = New System.Drawing.Point(0, 579)
        Me.lblSection.Name = "lblSection"
        Me.lblSection.Size = New System.Drawing.Size(96, 17)
        Me.lblSection.TabIndex = 62
        Me.lblSection.Text = "Sections/Drafts"
        '
        'lblFinalReport
        '
        Me.lblFinalReport.AutoSize = True
        Me.lblFinalReport.BackColor = System.Drawing.Color.Transparent
        Me.lblFinalReport.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.lblFinalReport.ForeColor = System.Drawing.Color.Black
        Me.lblFinalReport.Location = New System.Drawing.Point(0, 374)
        Me.lblFinalReport.Name = "lblFinalReport"
        Me.lblFinalReport.Size = New System.Drawing.Size(132, 17)
        Me.lblFinalReport.TabIndex = 61
        Me.lblFinalReport.Text = "Final Report Versions"
        '
        'dgvSections
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvSections.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvSections.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSections.Location = New System.Drawing.Point(0, 598)
        Me.dgvSections.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvSections.Name = "dgvSections"
        Me.dgvSections.Size = New System.Drawing.Size(279, 167)
        Me.dgvSections.TabIndex = 60
        '
        'dgvFinalReports
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvFinalReports.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.dgvFinalReports.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvFinalReports.Location = New System.Drawing.Point(0, 393)
        Me.dgvFinalReports.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvFinalReports.Name = "dgvFinalReports"
        Me.dgvFinalReports.Size = New System.Drawing.Size(279, 167)
        Me.dgvFinalReports.TabIndex = 59
        '
        'gbSaveType
        '
        Me.gbSaveType.Controls.Add(Me.rbSection)
        Me.gbSaveType.Controls.Add(Me.rbFinalReport)
        Me.gbSaveType.Location = New System.Drawing.Point(987, 242)
        Me.gbSaveType.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbSaveType.Name = "gbSaveType"
        Me.gbSaveType.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbSaveType.Size = New System.Drawing.Size(150, 97)
        Me.gbSaveType.TabIndex = 63
        Me.gbSaveType.TabStop = False
        Me.gbSaveType.Text = "Save Type"
        Me.gbSaveType.Visible = False
        '
        'rbSection
        '
        Me.rbSection.AutoSize = True
        Me.rbSection.Checked = True
        Me.rbSection.Location = New System.Drawing.Point(13, 46)
        Me.rbSection.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbSection.Name = "rbSection"
        Me.rbSection.Size = New System.Drawing.Size(102, 21)
        Me.rbSection.TabIndex = 1
        Me.rbSection.TabStop = True
        Me.rbSection.Text = "Section/Draft"
        Me.rbSection.UseVisualStyleBackColor = True
        '
        'rbFinalReport
        '
        Me.rbFinalReport.AutoSize = True
        Me.rbFinalReport.Location = New System.Drawing.Point(13, 21)
        Me.rbFinalReport.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbFinalReport.Name = "rbFinalReport"
        Me.rbFinalReport.Size = New System.Drawing.Size(96, 21)
        Me.rbFinalReport.TabIndex = 0
        Me.rbFinalReport.Text = "Final Report"
        Me.rbFinalReport.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(177, 11)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(70, 40)
        Me.Button1.TabIndex = 62
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'frmDocumentCompare
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1386, 867)
        Me.ControlBox = False
        Me.Controls.Add(Me.gbSaveType)
        Me.Controls.Add(Me.panProgress1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.panSave)
        Me.Controls.Add(Me.panWT)
        Me.Controls.Add(Me.panSC1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.pan1)
        Me.Controls.Add(Me.txtWSID)
        Me.Controls.Add(Me.dgvProjects)
        Me.Controls.Add(Me.dgvStudies)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.txtStudyID)
        Me.Controls.Add(Me.panStudy)
        Me.Controls.Add(Me.txtProjectID)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmDocumentCompare"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Compare Documents..."
        Me.panH1.ResumeLayout(False)
        Me.panH1.PerformLayout()
        Me.panH2.ResumeLayout(False)
        Me.panH2.PerformLayout()
        Me.sc1.Panel1.ResumeLayout(False)
        Me.sc1.Panel1.PerformLayout()
        Me.sc1.Panel2.ResumeLayout(False)
        Me.sc1.Panel2.PerformLayout()
        CType(Me.sc1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.sc1.ResumeLayout(False)
        CType(Me.ovDC1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ovDC2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panProgress1.ResumeLayout(False)
        CType(Me.dgvStudies, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panStudy.ResumeLayout(False)
        Me.panStudy.PerformLayout()
        CType(Me.dgvProjects, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panWT.ResumeLayout(False)
        Me.panWT.PerformLayout()
        Me.pan2.ResumeLayout(False)
        CType(Me.ovDC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pan1.ResumeLayout(False)
        Me.pan1.PerformLayout()
        Me.panProgress.ResumeLayout(False)
        Me.gbCompare.ResumeLayout(False)
        Me.gbCompare.PerformLayout()
        Me.panSC1.ResumeLayout(False)
        Me.panSC1.PerformLayout()
        Me.panSave.ResumeLayout(False)
        Me.panSave.PerformLayout()
        Me.panOpen.ResumeLayout(False)
        Me.panEdit.ResumeLayout(False)
        Me.panEdit.PerformLayout()
        Me.gbLoad.ResumeLayout(False)
        Me.gbLoad.PerformLayout()
        CType(Me.dgvSections, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvFinalReports, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbSaveType.ResumeLayout(False)
        Me.gbSaveType.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout

End Sub
    Friend WithEvents cbxWRT1 As System.Windows.Forms.ComboBox
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Friend WithEvents panH1 As System.Windows.Forms.Panel
    Friend WithEvents cbxFR1 As System.Windows.Forms.ComboBox
    Friend WithEvents panH2 As System.Windows.Forms.Panel
    Friend WithEvents cbxFR2 As System.Windows.Forms.ComboBox
    Friend WithEvents lbl2 As System.Windows.Forms.Label
    Friend WithEvents cbxWRT2 As System.Windows.Forms.ComboBox
    Friend WithEvents sc1 As System.Windows.Forms.SplitContainer
    'Friend WithEvents ovDC1 As AxEDOfficeLib.AxEDOffice
    'Friend WithEvents ovDC2 As AxEDOfficeLib.AxEDOffice
    Friend WithEvents txtProjectID As System.Windows.Forms.TextBox
    Friend WithEvents txtStudyID As System.Windows.Forms.TextBox
    Friend WithEvents lblProject As System.Windows.Forms.Label
    Friend WithEvents txtProject As System.Windows.Forms.TextBox
    Friend WithEvents dgvStudies As System.Windows.Forms.DataGridView
    Friend WithEvents txtStudy As System.Windows.Forms.TextBox
    Friend WithEvents lblStudy As System.Windows.Forms.Label
    Friend WithEvents panStudy As System.Windows.Forms.Panel
    Friend WithEvents cmdBrowseStudy As System.Windows.Forms.Button
    Friend WithEvents dgvProjects As System.Windows.Forms.DataGridView
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents panWT As System.Windows.Forms.Panel
    Friend WithEvents lblWordTemplate As System.Windows.Forms.Label
    Friend WithEvents txtWordTemplate As System.Windows.Forms.TextBox
    Friend WithEvents txtWSID As System.Windows.Forms.TextBox
    Friend WithEvents pan2 As System.Windows.Forms.Panel
    'Friend WithEvents ovDC As AxEDOfficeLib.AxEDOffice
    Friend WithEvents pan1 As System.Windows.Forms.Panel
    Friend WithEvents cmdCompare As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents gbCompare As System.Windows.Forms.GroupBox
    Friend WithEvents rbCompare As System.Windows.Forms.RadioButton
    Friend WithEvents rbSbyS As System.Windows.Forms.RadioButton
    Friend WithEvents panProgress As System.Windows.Forms.Panel
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents lblLoadedDoc As System.Windows.Forms.Label
    Friend WithEvents lblCompareWith As System.Windows.Forms.Label
    Friend WithEvents panSC1 As System.Windows.Forms.Panel
    Friend WithEvents txtReportTitle As System.Windows.Forms.TextBox
    Friend WithEvents lblReportTitle As System.Windows.Forms.Label
    Friend WithEvents txtReportNumber As System.Windows.Forms.TextBox
    Friend WithEvents lblReportNumber As System.Windows.Forms.Label
    Friend WithEvents txtLoadedDocDescription As System.Windows.Forms.TextBox
    Friend WithEvents panSave As System.Windows.Forms.Panel
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdOpenPDF As System.Windows.Forms.Button
    Friend WithEvents cmdWord As System.Windows.Forms.Button
    Friend WithEvents dgvSections As System.Windows.Forms.DataGridView
    Friend WithEvents dgvFinalReports As System.Windows.Forms.DataGridView
    Friend WithEvents lblSection As System.Windows.Forms.Label
    Friend WithEvents lblFinalReport As System.Windows.Forms.Label
    Friend WithEvents gbSaveType As System.Windows.Forms.GroupBox
    Friend WithEvents rbSection As System.Windows.Forms.RadioButton
    Friend WithEvents rbFinalReport As System.Windows.Forms.RadioButton
    Friend WithEvents lblNewDoc As System.Windows.Forms.Label
    Friend WithEvents cmdCompareFinalReport As System.Windows.Forms.Button
    Friend WithEvents txtDescr As System.Windows.Forms.TextBox
    Friend WithEvents lblDescr As System.Windows.Forms.Label
    Friend WithEvents gbLoad As System.Windows.Forms.GroupBox
    Friend WithEvents rbLoadCompare As System.Windows.Forms.RadioButton
    Friend WithEvents rbLoad As System.Windows.Forms.RadioButton
    Friend WithEvents panProgress1 As System.Windows.Forms.Panel
    Friend WithEvents lblProgress1 As System.Windows.Forms.Label
    Friend WithEvents chkRPane As System.Windows.Forms.CheckBox
    Friend WithEvents chkRPane1 As System.Windows.Forms.CheckBox
    Friend WithEvents txtComparedDocDescription As System.Windows.Forms.TextBox
    Friend WithEvents lblComparedDoc As System.Windows.Forms.Label
    Friend WithEvents lblInstructions01 As System.Windows.Forms.Label
    Friend WithEvents cmdCompareSection As System.Windows.Forms.Button
    Friend WithEvents cmdInsertDocument As System.Windows.Forms.Button
    Friend WithEvents panOptions As System.Windows.Forms.Panel
    Friend WithEvents cmdPrint As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdEdit As System.Windows.Forms.Button
    Friend WithEvents panEdit As System.Windows.Forms.Panel
    Friend WithEvents panOpen As System.Windows.Forms.Panel
    Friend WithEvents chkShowHilite As System.Windows.Forms.CheckBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents cmdClearCompare As System.Windows.Forms.Button
    Friend WithEvents cmdPaste As System.Windows.Forms.Button
    Friend WithEvents cmdCopy As System.Windows.Forms.Button
    Friend WithEvents lblReadOnly As System.Windows.Forms.Label
    Friend WithEvents lblFinalReportLocked As System.Windows.Forms.Label
    Friend WithEvents ovDC1 As AxEDOfficeLib.AxEDOffice
    Friend WithEvents ovDC2 As AxEDOfficeLib.AxEDOffice
    Friend WithEvents ovDC As AxEDOfficeLib.AxEDOffice
End Class
