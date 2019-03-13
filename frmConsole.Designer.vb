<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConsole
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmConsole))
        Me.panButtons = New System.Windows.Forms.Panel()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.lblResults = New System.Windows.Forms.Label()
        Me.cmdResults = New System.Windows.Forms.Button()
        Me.lblStudyDesigner = New System.Windows.Forms.Label()
        Me.cmdStudyDesigner = New System.Windows.Forms.Button()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.mnuMain = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuAbout = New System.Windows.Forms.ToolStripMenuItem()
        Me.cmdAbout = New System.Windows.Forms.Button()
        Me.gbxLogo = New System.Windows.Forms.GroupBox()
        Me.picGubbs = New System.Windows.Forms.PictureBox()
        Me.panDashboard = New System.Windows.Forms.Panel()
        Me.gbPB = New System.Windows.Forms.GroupBox()
        Me.dgvDashboard = New System.Windows.Forms.DataGridView()
        Me.lblTotalGuWuStudiesCount = New System.Windows.Forms.Label()
        Me.lblSort = New System.Windows.Forms.Label()
        Me.gbReportDashboardSortOrder = New System.Windows.Forms.GroupBox()
        Me.rbDESCReportDB = New System.Windows.Forms.RadioButton()
        Me.rbASCReportDB = New System.Windows.Forms.RadioButton()
        Me.cbxSortDashboard = New System.Windows.Forms.ComboBox()
        Me.rbTotalGuWuStudies = New System.Windows.Forms.RadioButton()
        Me.pbTotalGuWuStudies = New System.Windows.Forms.ProgressBar()
        Me.lblpbTotalGuWuStudies = New System.Windows.Forms.Label()
        Me.rbFinalReports = New System.Windows.Forms.RadioButton()
        Me.lblTotalOpenStudiesCount = New System.Windows.Forms.Label()
        Me.lblInProgressCount = New System.Windows.Forms.Label()
        Me.rbDraftReports = New System.Windows.Forms.RadioButton()
        Me.lblFinalCount = New System.Windows.Forms.Label()
        Me.pbFinalReport = New System.Windows.Forms.ProgressBar()
        Me.lblInDraftCount = New System.Windows.Forms.Label()
        Me.pbTotalOpenStudies = New System.Windows.Forms.ProgressBar()
        Me.rbInProgressStudies = New System.Windows.Forms.RadioButton()
        Me.pbInProgressReport = New System.Windows.Forms.ProgressBar()
        Me.pbDraftReport = New System.Windows.Forms.ProgressBar()
        Me.rbTotalOpenStudies = New System.Windows.Forms.RadioButton()
        Me.lblpbTotalOpenStudies = New System.Windows.Forms.Label()
        Me.lblpbInProgressReport = New System.Windows.Forms.Label()
        Me.lblpbDraftReport = New System.Windows.Forms.Label()
        Me.lblpbFinalReport = New System.Windows.Forms.Label()
        Me.lblDashboardTitle = New System.Windows.Forms.Label()
        Me.gbStats = New System.Windows.Forms.GroupBox()
        Me.rbAbsolute = New System.Windows.Forms.RadioButton()
        Me.rbPercent = New System.Windows.Forms.RadioButton()
        Me.lblExit = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RadioButton3 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.panBody = New System.Windows.Forms.Panel()
        Me.cmdChangePassword = New System.Windows.Forms.Button()
        Me.cmdLogin = New System.Windows.Forms.Button()
        Me.lblAuditTrail = New System.Windows.Forms.Label()
        Me.lblConfig = New System.Windows.Forms.Label()
        Me.lblReportWriter = New System.Windows.Forms.Label()
        Me.lbl1b = New System.Windows.Forms.Label()
        Me.lbl1a = New System.Windows.Forms.TextBox()
        Me.lbl1 = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.cmdAuditTrail = New System.Windows.Forms.Button()
        Me.cmdConfig = New System.Windows.Forms.Button()
        Me.cmdReportWriter = New System.Windows.Forms.Button()
        Me.lbl2 = New System.Windows.Forms.Label()
        Me.panButtons.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        CType(Me.picGubbs, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panDashboard.SuspendLayout()
        Me.gbPB.SuspendLayout()
        CType(Me.dgvDashboard, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbReportDashboardSortOrder.SuspendLayout()
        Me.gbStats.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.panBody.SuspendLayout()
        Me.SuspendLayout()
        '
        'panButtons
        '
        Me.panButtons.Controls.Add(Me.Button4)
        Me.panButtons.Controls.Add(Me.Button3)
        Me.panButtons.Controls.Add(Me.Button2)
        Me.panButtons.Location = New System.Drawing.Point(328, 452)
        Me.panButtons.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panButtons.Name = "panButtons"
        Me.panButtons.Size = New System.Drawing.Size(189, 208)
        Me.panButtons.TabIndex = 32
        Me.panButtons.Visible = False
        '
        'Button4
        '
        Me.Button4.FlatAppearance.BorderSize = 0
        Me.Button4.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Button4.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button4.ForeColor = System.Drawing.Color.Black
        Me.Button4.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Button4.Location = New System.Drawing.Point(23, 124)
        Me.Button4.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(127, 61)
        Me.Button4.TabIndex = 8
        Me.Button4.Text = "Flat no border, mouseover, colorclick"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.FlatAppearance.BorderSize = 0
        Me.Button3.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.Button3.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.ForeColor = System.Drawing.Color.Black
        Me.Button3.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Button3.Location = New System.Drawing.Point(12, 68)
        Me.Button3.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(127, 48)
        Me.Button3.TabIndex = 7
        Me.Button3.Text = "Flat no border, mouseover"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.Button2.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Button2.Location = New System.Drawing.Point(12, 17)
        Me.Button2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(127, 35)
        Me.Button2.TabIndex = 6
        Me.Button2.Text = "Button Flat"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'cmdExit
        '
        Me.cmdExit.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdExit.FlatAppearance.BorderSize = 0
        Me.cmdExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdExit.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ForeColor = System.Drawing.Color.Firebrick
        Me.cmdExit.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cmdExit.Location = New System.Drawing.Point(26, 395)
        Me.cmdExit.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(572, 42)
        Me.cmdExit.TabIndex = 9
        Me.cmdExit.Text = "E&xit"
        Me.cmdExit.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'lblResults
        '
        Me.lblResults.AutoSize = True
        Me.lblResults.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold)
        Me.lblResults.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.lblResults.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblResults.Location = New System.Drawing.Point(112, 682)
        Me.lblResults.Name = "lblResults"
        Me.lblResults.Size = New System.Drawing.Size(131, 25)
        Me.lblResults.TabIndex = 16
        Me.lblResults.Text = "Study &Results"
        Me.lblResults.Visible = False
        '
        'cmdResults
        '
        Me.cmdResults.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cmdResults.Location = New System.Drawing.Point(58, 681)
        Me.cmdResults.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdResults.Name = "cmdResults"
        Me.cmdResults.Size = New System.Drawing.Size(47, 35)
        Me.cmdResults.TabIndex = 17
        Me.cmdResults.UseVisualStyleBackColor = True
        Me.cmdResults.Visible = False
        '
        'lblStudyDesigner
        '
        Me.lblStudyDesigner.AutoSize = True
        Me.lblStudyDesigner.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold)
        Me.lblStudyDesigner.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.lblStudyDesigner.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblStudyDesigner.Location = New System.Drawing.Point(112, 635)
        Me.lblStudyDesigner.Name = "lblStudyDesigner"
        Me.lblStudyDesigner.Size = New System.Drawing.Size(148, 25)
        Me.lblStudyDesigner.TabIndex = 2
        Me.lblStudyDesigner.Text = "&Study Designer"
        Me.lblStudyDesigner.Visible = False
        '
        'cmdStudyDesigner
        '
        Me.cmdStudyDesigner.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.cmdStudyDesigner.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cmdStudyDesigner.Location = New System.Drawing.Point(58, 635)
        Me.cmdStudyDesigner.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdStudyDesigner.Name = "cmdStudyDesigner"
        Me.cmdStudyDesigner.Size = New System.Drawing.Size(47, 35)
        Me.cmdStudyDesigner.TabIndex = 3
        Me.cmdStudyDesigner.UseVisualStyleBackColor = True
        Me.cmdStudyDesigner.Visible = False
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuMain})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Padding = New System.Windows.Forms.Padding(7, 3, 0, 3)
        Me.MenuStrip1.Size = New System.Drawing.Size(1278, 27)
        Me.MenuStrip1.TabIndex = 17
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'mnuMain
        '
        Me.mnuMain.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAbout})
        Me.mnuMain.Name = "mnuMain"
        Me.mnuMain.Size = New System.Drawing.Size(53, 21)
        Me.mnuMain.Text = "Menu"
        '
        'mnuAbout
        '
        Me.mnuAbout.Name = "mnuAbout"
        Me.mnuAbout.Size = New System.Drawing.Size(120, 22)
        Me.mnuAbout.Text = "About..."
        '
        'cmdAbout
        '
        Me.cmdAbout.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdAbout.BackColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdAbout.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdAbout.ForeColor = System.Drawing.Color.White
        Me.cmdAbout.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cmdAbout.Location = New System.Drawing.Point(561, 4)
        Me.cmdAbout.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdAbout.Name = "cmdAbout"
        Me.cmdAbout.Size = New System.Drawing.Size(84, 35)
        Me.cmdAbout.TabIndex = 18
        Me.cmdAbout.Text = "A&bout..."
        Me.cmdAbout.UseVisualStyleBackColor = False
        '
        'gbxLogo
        '
        Me.gbxLogo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbxLogo.Location = New System.Drawing.Point(324, 3409)
        Me.gbxLogo.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbxLogo.Name = "gbxLogo"
        Me.gbxLogo.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbxLogo.Size = New System.Drawing.Size(188, 64)
        Me.gbxLogo.TabIndex = 19
        Me.gbxLogo.TabStop = False
        Me.gbxLogo.Text = "A product of..."
        Me.gbxLogo.Visible = False
        '
        'picGubbs
        '
        Me.picGubbs.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.picGubbs.Image = CType(resources.GetObject("picGubbs.Image"), System.Drawing.Image)
        Me.picGubbs.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.picGubbs.Location = New System.Drawing.Point(778, 3443)
        Me.picGubbs.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.picGubbs.Name = "picGubbs"
        Me.picGubbs.Size = New System.Drawing.Size(245, 63)
        Me.picGubbs.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picGubbs.TabIndex = 0
        Me.picGubbs.TabStop = False
        '
        'panDashboard
        '
        Me.panDashboard.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.panDashboard.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panDashboard.Controls.Add(Me.cmdAbout)
        Me.panDashboard.Controls.Add(Me.gbPB)
        Me.panDashboard.Controls.Add(Me.lblDashboardTitle)
        Me.panDashboard.Controls.Add(Me.gbStats)
        Me.panDashboard.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.panDashboard.Location = New System.Drawing.Point(615, 42)
        Me.panDashboard.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panDashboard.Name = "panDashboard"
        Me.panDashboard.Size = New System.Drawing.Size(651, 699)
        Me.panDashboard.TabIndex = 20
        '
        'gbPB
        '
        Me.gbPB.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbPB.Controls.Add(Me.dgvDashboard)
        Me.gbPB.Controls.Add(Me.lblTotalGuWuStudiesCount)
        Me.gbPB.Controls.Add(Me.lblSort)
        Me.gbPB.Controls.Add(Me.gbReportDashboardSortOrder)
        Me.gbPB.Controls.Add(Me.cbxSortDashboard)
        Me.gbPB.Controls.Add(Me.rbTotalGuWuStudies)
        Me.gbPB.Controls.Add(Me.pbTotalGuWuStudies)
        Me.gbPB.Controls.Add(Me.lblpbTotalGuWuStudies)
        Me.gbPB.Controls.Add(Me.rbFinalReports)
        Me.gbPB.Controls.Add(Me.lblTotalOpenStudiesCount)
        Me.gbPB.Controls.Add(Me.lblInProgressCount)
        Me.gbPB.Controls.Add(Me.rbDraftReports)
        Me.gbPB.Controls.Add(Me.lblFinalCount)
        Me.gbPB.Controls.Add(Me.pbFinalReport)
        Me.gbPB.Controls.Add(Me.lblInDraftCount)
        Me.gbPB.Controls.Add(Me.pbTotalOpenStudies)
        Me.gbPB.Controls.Add(Me.rbInProgressStudies)
        Me.gbPB.Controls.Add(Me.pbInProgressReport)
        Me.gbPB.Controls.Add(Me.pbDraftReport)
        Me.gbPB.Controls.Add(Me.rbTotalOpenStudies)
        Me.gbPB.Controls.Add(Me.lblpbTotalOpenStudies)
        Me.gbPB.Controls.Add(Me.lblpbInProgressReport)
        Me.gbPB.Controls.Add(Me.lblpbDraftReport)
        Me.gbPB.Controls.Add(Me.lblpbFinalReport)
        Me.gbPB.Location = New System.Drawing.Point(20, 111)
        Me.gbPB.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbPB.Name = "gbPB"
        Me.gbPB.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbPB.Size = New System.Drawing.Size(613, 579)
        Me.gbPB.TabIndex = 21
        Me.gbPB.TabStop = False
        '
        'dgvDashboard
        '
        Me.dgvDashboard.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvDashboard.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDashboard.Location = New System.Drawing.Point(12, 328)
        Me.dgvDashboard.Name = "dgvDashboard"
        Me.dgvDashboard.Size = New System.Drawing.Size(592, 244)
        Me.dgvDashboard.TabIndex = 34
        '
        'lblTotalGuWuStudiesCount
        '
        Me.lblTotalGuWuStudiesCount.AutoSize = True
        Me.lblTotalGuWuStudiesCount.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblTotalGuWuStudiesCount.Location = New System.Drawing.Point(317, 31)
        Me.lblTotalGuWuStudiesCount.Name = "lblTotalGuWuStudiesCount"
        Me.lblTotalGuWuStudiesCount.Size = New System.Drawing.Size(15, 17)
        Me.lblTotalGuWuStudiesCount.TabIndex = 33
        Me.lblTotalGuWuStudiesCount.Text = "0"
        Me.lblTotalGuWuStudiesCount.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblSort
        '
        Me.lblSort.AutoSize = True
        Me.lblSort.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblSort.Location = New System.Drawing.Point(9, 294)
        Me.lblSort.Name = "lblSort"
        Me.lblSort.Size = New System.Drawing.Size(53, 17)
        Me.lblSort.TabIndex = 23
        Me.lblSort.Text = "Sort by:"
        Me.lblSort.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'gbReportDashboardSortOrder
        '
        Me.gbReportDashboardSortOrder.Controls.Add(Me.rbDESCReportDB)
        Me.gbReportDashboardSortOrder.Controls.Add(Me.rbASCReportDB)
        Me.gbReportDashboardSortOrder.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbReportDashboardSortOrder.Location = New System.Drawing.Point(249, 272)
        Me.gbReportDashboardSortOrder.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbReportDashboardSortOrder.Name = "gbReportDashboardSortOrder"
        Me.gbReportDashboardSortOrder.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbReportDashboardSortOrder.Size = New System.Drawing.Size(154, 44)
        Me.gbReportDashboardSortOrder.TabIndex = 24
        Me.gbReportDashboardSortOrder.TabStop = False
        '
        'rbDESCReportDB
        '
        Me.rbDESCReportDB.AutoSize = True
        Me.rbDESCReportDB.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.rbDESCReportDB.Location = New System.Drawing.Point(73, 16)
        Me.rbDESCReportDB.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbDESCReportDB.Name = "rbDESCReportDB"
        Me.rbDESCReportDB.Size = New System.Drawing.Size(57, 21)
        Me.rbDESCReportDB.TabIndex = 1
        Me.rbDESCReportDB.Text = "DESC"
        Me.rbDESCReportDB.UseVisualStyleBackColor = True
        '
        'rbASCReportDB
        '
        Me.rbASCReportDB.AutoSize = True
        Me.rbASCReportDB.Checked = True
        Me.rbASCReportDB.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.rbASCReportDB.Location = New System.Drawing.Point(6, 16)
        Me.rbASCReportDB.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbASCReportDB.Name = "rbASCReportDB"
        Me.rbASCReportDB.Size = New System.Drawing.Size(49, 21)
        Me.rbASCReportDB.TabIndex = 0
        Me.rbASCReportDB.TabStop = True
        Me.rbASCReportDB.Text = "ASC"
        Me.rbASCReportDB.UseVisualStyleBackColor = True
        '
        'cbxSortDashboard
        '
        Me.cbxSortDashboard.FormattingEnabled = True
        Me.cbxSortDashboard.Location = New System.Drawing.Point(62, 291)
        Me.cbxSortDashboard.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSortDashboard.Name = "cbxSortDashboard"
        Me.cbxSortDashboard.Size = New System.Drawing.Size(179, 25)
        Me.cbxSortDashboard.TabIndex = 22
        '
        'rbTotalGuWuStudies
        '
        Me.rbTotalGuWuStudies.AutoSize = True
        Me.rbTotalGuWuStudies.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.rbTotalGuWuStudies.Location = New System.Drawing.Point(29, 31)
        Me.rbTotalGuWuStudies.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbTotalGuWuStudies.Name = "rbTotalGuWuStudies"
        Me.rbTotalGuWuStudies.Size = New System.Drawing.Size(14, 13)
        Me.rbTotalGuWuStudies.TabIndex = 32
        Me.rbTotalGuWuStudies.UseVisualStyleBackColor = True
        '
        'pbTotalGuWuStudies
        '
        Me.pbTotalGuWuStudies.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbTotalGuWuStudies.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.pbTotalGuWuStudies.Location = New System.Drawing.Point(222, 21)
        Me.pbTotalGuWuStudies.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pbTotalGuWuStudies.Name = "pbTotalGuWuStudies"
        Me.pbTotalGuWuStudies.Size = New System.Drawing.Size(383, 37)
        Me.pbTotalGuWuStudies.TabIndex = 30
        '
        'lblpbTotalGuWuStudies
        '
        Me.lblpbTotalGuWuStudies.AutoSize = True
        Me.lblpbTotalGuWuStudies.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblpbTotalGuWuStudies.Location = New System.Drawing.Point(52, 31)
        Me.lblpbTotalGuWuStudies.Name = "lblpbTotalGuWuStudies"
        Me.lblpbTotalGuWuStudies.Size = New System.Drawing.Size(142, 17)
        Me.lblpbTotalGuWuStudies.TabIndex = 31
        Me.lblpbTotalGuWuStudies.Text = "Total StudyDoc Studies"
        Me.lblpbTotalGuWuStudies.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'rbFinalReports
        '
        Me.rbFinalReports.AutoSize = True
        Me.rbFinalReports.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.rbFinalReports.Location = New System.Drawing.Point(29, 73)
        Me.rbFinalReports.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbFinalReports.Name = "rbFinalReports"
        Me.rbFinalReports.Size = New System.Drawing.Size(14, 13)
        Me.rbFinalReports.TabIndex = 29
        Me.rbFinalReports.UseVisualStyleBackColor = True
        '
        'lblTotalOpenStudiesCount
        '
        Me.lblTotalOpenStudiesCount.AutoSize = True
        Me.lblTotalOpenStudiesCount.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblTotalOpenStudiesCount.Location = New System.Drawing.Point(317, 143)
        Me.lblTotalOpenStudiesCount.Name = "lblTotalOpenStudiesCount"
        Me.lblTotalOpenStudiesCount.Size = New System.Drawing.Size(15, 17)
        Me.lblTotalOpenStudiesCount.TabIndex = 23
        Me.lblTotalOpenStudiesCount.Text = "0"
        Me.lblTotalOpenStudiesCount.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblInProgressCount
        '
        Me.lblInProgressCount.AutoSize = True
        Me.lblInProgressCount.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblInProgressCount.Location = New System.Drawing.Point(317, 187)
        Me.lblInProgressCount.Name = "lblInProgressCount"
        Me.lblInProgressCount.Size = New System.Drawing.Size(15, 17)
        Me.lblInProgressCount.TabIndex = 22
        Me.lblInProgressCount.Text = "0"
        Me.lblInProgressCount.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'rbDraftReports
        '
        Me.rbDraftReports.AutoSize = True
        Me.rbDraftReports.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.rbDraftReports.Location = New System.Drawing.Point(29, 230)
        Me.rbDraftReports.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbDraftReports.Name = "rbDraftReports"
        Me.rbDraftReports.Size = New System.Drawing.Size(14, 13)
        Me.rbDraftReports.TabIndex = 28
        Me.rbDraftReports.UseVisualStyleBackColor = True
        '
        'lblFinalCount
        '
        Me.lblFinalCount.AutoSize = True
        Me.lblFinalCount.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblFinalCount.Location = New System.Drawing.Point(317, 75)
        Me.lblFinalCount.Name = "lblFinalCount"
        Me.lblFinalCount.Size = New System.Drawing.Size(15, 17)
        Me.lblFinalCount.TabIndex = 25
        Me.lblFinalCount.Text = "0"
        Me.lblFinalCount.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pbFinalReport
        '
        Me.pbFinalReport.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbFinalReport.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.pbFinalReport.Location = New System.Drawing.Point(222, 65)
        Me.pbFinalReport.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pbFinalReport.Name = "pbFinalReport"
        Me.pbFinalReport.Size = New System.Drawing.Size(383, 35)
        Me.pbFinalReport.TabIndex = 3
        '
        'lblInDraftCount
        '
        Me.lblInDraftCount.AutoSize = True
        Me.lblInDraftCount.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblInDraftCount.Location = New System.Drawing.Point(317, 231)
        Me.lblInDraftCount.Name = "lblInDraftCount"
        Me.lblInDraftCount.Size = New System.Drawing.Size(15, 17)
        Me.lblInDraftCount.TabIndex = 24
        Me.lblInDraftCount.Text = "0"
        Me.lblInDraftCount.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pbTotalOpenStudies
        '
        Me.pbTotalOpenStudies.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbTotalOpenStudies.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.pbTotalOpenStudies.Location = New System.Drawing.Point(222, 132)
        Me.pbTotalOpenStudies.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pbTotalOpenStudies.Name = "pbTotalOpenStudies"
        Me.pbTotalOpenStudies.Size = New System.Drawing.Size(383, 37)
        Me.pbTotalOpenStudies.TabIndex = 0
        '
        'rbInProgressStudies
        '
        Me.rbInProgressStudies.AutoSize = True
        Me.rbInProgressStudies.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.rbInProgressStudies.Location = New System.Drawing.Point(29, 187)
        Me.rbInProgressStudies.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbInProgressStudies.Name = "rbInProgressStudies"
        Me.rbInProgressStudies.Size = New System.Drawing.Size(14, 13)
        Me.rbInProgressStudies.TabIndex = 27
        Me.rbInProgressStudies.UseVisualStyleBackColor = True
        '
        'pbInProgressReport
        '
        Me.pbInProgressReport.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbInProgressReport.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.pbInProgressReport.Location = New System.Drawing.Point(222, 177)
        Me.pbInProgressReport.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pbInProgressReport.Name = "pbInProgressReport"
        Me.pbInProgressReport.Size = New System.Drawing.Size(383, 37)
        Me.pbInProgressReport.TabIndex = 1
        '
        'pbDraftReport
        '
        Me.pbDraftReport.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbDraftReport.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.pbDraftReport.Location = New System.Drawing.Point(222, 221)
        Me.pbDraftReport.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pbDraftReport.Name = "pbDraftReport"
        Me.pbDraftReport.Size = New System.Drawing.Size(383, 37)
        Me.pbDraftReport.TabIndex = 2
        '
        'rbTotalOpenStudies
        '
        Me.rbTotalOpenStudies.AutoSize = True
        Me.rbTotalOpenStudies.Checked = True
        Me.rbTotalOpenStudies.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.rbTotalOpenStudies.Location = New System.Drawing.Point(29, 143)
        Me.rbTotalOpenStudies.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbTotalOpenStudies.Name = "rbTotalOpenStudies"
        Me.rbTotalOpenStudies.Size = New System.Drawing.Size(14, 13)
        Me.rbTotalOpenStudies.TabIndex = 26
        Me.rbTotalOpenStudies.TabStop = True
        Me.rbTotalOpenStudies.UseVisualStyleBackColor = True
        '
        'lblpbTotalOpenStudies
        '
        Me.lblpbTotalOpenStudies.AutoSize = True
        Me.lblpbTotalOpenStudies.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblpbTotalOpenStudies.Location = New System.Drawing.Point(52, 143)
        Me.lblpbTotalOpenStudies.Name = "lblpbTotalOpenStudies"
        Me.lblpbTotalOpenStudies.Size = New System.Drawing.Size(119, 17)
        Me.lblpbTotalOpenStudies.TabIndex = 4
        Me.lblpbTotalOpenStudies.Text = "Total Open Studies"
        Me.lblpbTotalOpenStudies.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblpbInProgressReport
        '
        Me.lblpbInProgressReport.AutoSize = True
        Me.lblpbInProgressReport.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblpbInProgressReport.Location = New System.Drawing.Point(52, 187)
        Me.lblpbInProgressReport.Name = "lblpbInProgressReport"
        Me.lblpbInProgressReport.Size = New System.Drawing.Size(120, 17)
        Me.lblpbInProgressReport.TabIndex = 5
        Me.lblpbInProgressReport.Text = "In Progress Studies"
        Me.lblpbInProgressReport.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblpbDraftReport
        '
        Me.lblpbDraftReport.AutoSize = True
        Me.lblpbDraftReport.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblpbDraftReport.Location = New System.Drawing.Point(52, 230)
        Me.lblpbDraftReport.Name = "lblpbDraftReport"
        Me.lblpbDraftReport.Size = New System.Drawing.Size(101, 17)
        Me.lblpbDraftReport.TabIndex = 6
        Me.lblpbDraftReport.Text = "Reports in Draft"
        Me.lblpbDraftReport.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblpbFinalReport
        '
        Me.lblpbFinalReport.AutoSize = True
        Me.lblpbFinalReport.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblpbFinalReport.Location = New System.Drawing.Point(52, 73)
        Me.lblpbFinalReport.Name = "lblpbFinalReport"
        Me.lblpbFinalReport.Size = New System.Drawing.Size(84, 17)
        Me.lblpbFinalReport.TabIndex = 7
        Me.lblpbFinalReport.Text = "Reports Final"
        Me.lblpbFinalReport.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblDashboardTitle
        '
        Me.lblDashboardTitle.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblDashboardTitle.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold)
        Me.lblDashboardTitle.ForeColor = System.Drawing.Color.White
        Me.lblDashboardTitle.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblDashboardTitle.Location = New System.Drawing.Point(-1, -1)
        Me.lblDashboardTitle.Name = "lblDashboardTitle"
        Me.lblDashboardTitle.Size = New System.Drawing.Size(293, 43)
        Me.lblDashboardTitle.TabIndex = 21
        Me.lblDashboardTitle.Text = "Report Metric Dashboard"
        Me.lblDashboardTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'gbStats
        '
        Me.gbStats.Controls.Add(Me.rbAbsolute)
        Me.gbStats.Controls.Add(Me.rbPercent)
        Me.gbStats.Location = New System.Drawing.Point(20, 48)
        Me.gbStats.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbStats.Name = "gbStats"
        Me.gbStats.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbStats.Size = New System.Drawing.Size(223, 61)
        Me.gbStats.TabIndex = 9
        Me.gbStats.TabStop = False
        Me.gbStats.Text = "Show Stats..."
        '
        'rbAbsolute
        '
        Me.rbAbsolute.AutoSize = True
        Me.rbAbsolute.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.rbAbsolute.Location = New System.Drawing.Point(128, 26)
        Me.rbAbsolute.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbAbsolute.Name = "rbAbsolute"
        Me.rbAbsolute.Size = New System.Drawing.Size(77, 21)
        Me.rbAbsolute.TabIndex = 1
        Me.rbAbsolute.Text = "Absolute"
        Me.rbAbsolute.UseVisualStyleBackColor = True
        '
        'rbPercent
        '
        Me.rbPercent.AutoSize = True
        Me.rbPercent.Checked = True
        Me.rbPercent.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.rbPercent.Location = New System.Drawing.Point(20, 26)
        Me.rbPercent.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbPercent.Name = "rbPercent"
        Me.rbPercent.Size = New System.Drawing.Size(91, 21)
        Me.rbPercent.TabIndex = 0
        Me.rbPercent.TabStop = True
        Me.rbPercent.Text = "Percentage"
        Me.rbPercent.UseVisualStyleBackColor = True
        '
        'lblExit
        '
        Me.lblExit.AutoSize = True
        Me.lblExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblExit.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblExit.ForeColor = System.Drawing.Color.Firebrick
        Me.lblExit.Location = New System.Drawing.Point(523, 452)
        Me.lblExit.Name = "lblExit"
        Me.lblExit.Size = New System.Drawing.Size(45, 25)
        Me.lblExit.TabIndex = 31
        Me.lblExit.Text = "E&xit"
        Me.lblExit.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RadioButton3)
        Me.GroupBox1.Controls.Add(Me.RadioButton2)
        Me.GroupBox1.Controls.Add(Me.RadioButton1)
        Me.GroupBox1.Location = New System.Drawing.Point(115, 452)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(174, 137)
        Me.GroupBox1.TabIndex = 32
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "GroupBox1"
        Me.GroupBox1.Visible = False
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Checked = True
        Me.RadioButton3.Location = New System.Drawing.Point(7, 47)
        Me.RadioButton3.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(70, 21)
        Me.RadioButton3.TabIndex = 34
        Me.RadioButton3.TabStop = True
        Me.RadioButton3.Text = "Normal"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(7, 107)
        Me.RadioButton2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(152, 21)
        Me.RadioButton2.TabIndex = 33
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "Flat button no border"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(7, 77)
        Me.RadioButton1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(160, 21)
        Me.RadioButton1.TabIndex = 32
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Flat button with border"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'panBody
        '
        Me.panBody.Controls.Add(Me.cmdChangePassword)
        Me.panBody.Controls.Add(Me.GroupBox1)
        Me.panBody.Controls.Add(Me.panButtons)
        Me.panBody.Controls.Add(Me.cmdLogin)
        Me.panBody.Controls.Add(Me.lblAuditTrail)
        Me.panBody.Controls.Add(Me.lblExit)
        Me.panBody.Controls.Add(Me.lblConfig)
        Me.panBody.Controls.Add(Me.lblReportWriter)
        Me.panBody.Controls.Add(Me.lbl1b)
        Me.panBody.Controls.Add(Me.lbl1a)
        Me.panBody.Controls.Add(Me.lbl1)
        Me.panBody.Controls.Add(Me.Button1)
        Me.panBody.Controls.Add(Me.cmdAuditTrail)
        Me.panBody.Controls.Add(Me.cmdExit)
        Me.panBody.Controls.Add(Me.cmdConfig)
        Me.panBody.Controls.Add(Me.cmdReportWriter)
        Me.panBody.Controls.Add(Me.lbl2)
        Me.panBody.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.panBody.ForeColor = System.Drawing.Color.Blue
        Me.panBody.Location = New System.Drawing.Point(6, 42)
        Me.panBody.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panBody.Name = "panBody"
        Me.panBody.Size = New System.Drawing.Size(603, 583)
        Me.panBody.TabIndex = 16
        '
        'cmdChangePassword
        '
        Me.cmdChangePassword.BackColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdChangePassword.CausesValidation = False
        Me.cmdChangePassword.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdChangePassword.ForeColor = System.Drawing.Color.White
        Me.cmdChangePassword.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cmdChangePassword.Location = New System.Drawing.Point(405, 290)
        Me.cmdChangePassword.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdChangePassword.Name = "cmdChangePassword"
        Me.cmdChangePassword.Size = New System.Drawing.Size(192, 42)
        Me.cmdChangePassword.TabIndex = 1
        Me.cmdChangePassword.Text = "Change &Password"
        Me.cmdChangePassword.UseVisualStyleBackColor = False
        '
        'cmdLogin
        '
        Me.cmdLogin.BackColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdLogin.CausesValidation = False
        Me.cmdLogin.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdLogin.ForeColor = System.Drawing.Color.White
        Me.cmdLogin.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cmdLogin.Location = New System.Drawing.Point(405, 233)
        Me.cmdLogin.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdLogin.Name = "cmdLogin"
        Me.cmdLogin.Size = New System.Drawing.Size(192, 42)
        Me.cmdLogin.TabIndex = 0
        Me.cmdLogin.Text = "&Log On"
        Me.cmdLogin.UseVisualStyleBackColor = False
        '
        'lblAuditTrail
        '
        Me.lblAuditTrail.AutoSize = True
        Me.lblAuditTrail.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblAuditTrail.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAuditTrail.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.lblAuditTrail.Location = New System.Drawing.Point(385, 307)
        Me.lblAuditTrail.Name = "lblAuditTrail"
        Me.lblAuditTrail.Size = New System.Drawing.Size(105, 25)
        Me.lblAuditTrail.TabIndex = 30
        Me.lblAuditTrail.Text = "Audit &Trail"
        Me.lblAuditTrail.Visible = False
        '
        'lblConfig
        '
        Me.lblConfig.AutoSize = True
        Me.lblConfig.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblConfig.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblConfig.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.lblConfig.Location = New System.Drawing.Point(385, 234)
        Me.lblConfig.Name = "lblConfig"
        Me.lblConfig.Size = New System.Drawing.Size(145, 25)
        Me.lblConfig.TabIndex = 29
        Me.lblConfig.Text = "&Administration"
        Me.lblConfig.Visible = False
        '
        'lblReportWriter
        '
        Me.lblReportWriter.AutoSize = True
        Me.lblReportWriter.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblReportWriter.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReportWriter.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.lblReportWriter.Location = New System.Drawing.Point(385, 275)
        Me.lblReportWriter.Name = "lblReportWriter"
        Me.lblReportWriter.Size = New System.Drawing.Size(135, 25)
        Me.lblReportWriter.TabIndex = 28
        Me.lblReportWriter.Text = "&Report Writer"
        Me.lblReportWriter.Visible = False
        '
        'lbl1b
        '
        Me.lbl1b.AutoSize = True
        Me.lbl1b.Font = New System.Drawing.Font("Century Gothic", 14.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl1b.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lbl1b.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lbl1b.Location = New System.Drawing.Point(176, 27)
        Me.lbl1b.Name = "lbl1b"
        Me.lbl1b.Size = New System.Drawing.Size(35, 23)
        Me.lbl1b.TabIndex = 15
        Me.lbl1b.Text = "TM"
        '
        'lbl1a
        '
        Me.lbl1a.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.lbl1a.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lbl1a.Font = New System.Drawing.Font("Century Gothic", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl1a.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lbl1a.Location = New System.Drawing.Point(115, 27)
        Me.lbl1a.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.lbl1a.Name = "lbl1a"
        Me.lbl1a.ReadOnly = True
        Me.lbl1a.Size = New System.Drawing.Size(77, 40)
        Me.lbl1a.TabIndex = 27
        Me.lbl1a.Text = "Doc"
        '
        'lbl1
        '
        Me.lbl1.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.lbl1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lbl1.Font = New System.Drawing.Font("Century Gothic", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.lbl1.Location = New System.Drawing.Point(5, 27)
        Me.lbl1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.ReadOnly = True
        Me.lbl1.Size = New System.Drawing.Size(112, 40)
        Me.lbl1.TabIndex = 26
        Me.lbl1.Text = "Study"
        Me.lbl1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Button1
        '
        Me.Button1.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Button1.Location = New System.Drawing.Point(525, 10)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(72, 55)
        Me.Button1.TabIndex = 18
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'cmdAuditTrail
        '
        Me.cmdAuditTrail.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdAuditTrail.FlatAppearance.BorderSize = 0
        Me.cmdAuditTrail.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdAuditTrail.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAuditTrail.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdAuditTrail.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cmdAuditTrail.Location = New System.Drawing.Point(26, 290)
        Me.cmdAuditTrail.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdAuditTrail.Name = "cmdAuditTrail"
        Me.cmdAuditTrail.Size = New System.Drawing.Size(255, 42)
        Me.cmdAuditTrail.TabIndex = 22
        Me.cmdAuditTrail.Text = "Audit &Trail"
        Me.cmdAuditTrail.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAuditTrail.UseVisualStyleBackColor = True
        '
        'cmdConfig
        '
        Me.cmdConfig.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdConfig.FlatAppearance.BorderSize = 0
        Me.cmdConfig.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdConfig.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdConfig.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdConfig.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cmdConfig.Location = New System.Drawing.Point(26, 233)
        Me.cmdConfig.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdConfig.Name = "cmdConfig"
        Me.cmdConfig.Size = New System.Drawing.Size(255, 42)
        Me.cmdConfig.TabIndex = 7
        Me.cmdConfig.Text = "&Administration"
        Me.cmdConfig.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdConfig.UseVisualStyleBackColor = True
        '
        'cmdReportWriter
        '
        Me.cmdReportWriter.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdReportWriter.FlatAppearance.BorderSize = 0
        Me.cmdReportWriter.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdReportWriter.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReportWriter.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdReportWriter.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cmdReportWriter.Location = New System.Drawing.Point(26, 139)
        Me.cmdReportWriter.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdReportWriter.Name = "cmdReportWriter"
        Me.cmdReportWriter.Size = New System.Drawing.Size(572, 42)
        Me.cmdReportWriter.TabIndex = 5
        Me.cmdReportWriter.Text = "&Report Writer"
        Me.cmdReportWriter.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdReportWriter.UseVisualStyleBackColor = True
        '
        'lbl2
        '
        Me.lbl2.AutoSize = True
        Me.lbl2.Font = New System.Drawing.Font("Century Gothic", 14.25!, System.Drawing.FontStyle.Bold)
        Me.lbl2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lbl2.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lbl2.Location = New System.Drawing.Point(26, 84)
        Me.lbl2.Name = "lbl2"
        Me.lbl2.Size = New System.Drawing.Size(228, 23)
        Me.lbl2.TabIndex = 7
        Me.lbl2.Text = "Report Writing Manager"
        Me.lbl2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmConsole
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1278, 753)
        Me.Controls.Add(Me.panBody)
        Me.Controls.Add(Me.picGubbs)
        Me.Controls.Add(Me.panDashboard)
        Me.Controls.Add(Me.gbxLogo)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.lblStudyDesigner)
        Me.Controls.Add(Me.cmdStudyDesigner)
        Me.Controls.Add(Me.lblResults)
        Me.Controls.Add(Me.cmdResults)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "frmConsole"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "StudyDoc(TM) Report Writing Manager"
        Me.panButtons.ResumeLayout(False)
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.picGubbs, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panDashboard.ResumeLayout(False)
        Me.gbPB.ResumeLayout(False)
        Me.gbPB.PerformLayout()
        CType(Me.dgvDashboard, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbReportDashboardSortOrder.ResumeLayout(False)
        Me.gbReportDashboardSortOrder.PerformLayout()
        Me.gbStats.ResumeLayout(False)
        Me.gbStats.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.panBody.ResumeLayout(False)
        Me.panBody.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblStudyDesigner As System.Windows.Forms.Label
    Friend WithEvents cmdStudyDesigner As System.Windows.Forms.Button
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents mnuMain As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAbout As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lblResults As System.Windows.Forms.Label
    Friend WithEvents cmdResults As System.Windows.Forms.Button
    Friend WithEvents cmdAbout As System.Windows.Forms.Button
    Friend WithEvents gbxLogo As System.Windows.Forms.GroupBox
    Friend WithEvents picGubbs As System.Windows.Forms.PictureBox
    Friend WithEvents panDashboard As System.Windows.Forms.Panel
    Friend WithEvents pbDraftReport As System.Windows.Forms.ProgressBar
    Friend WithEvents pbInProgressReport As System.Windows.Forms.ProgressBar
    Friend WithEvents pbTotalOpenStudies As System.Windows.Forms.ProgressBar
    Friend WithEvents pbFinalReport As System.Windows.Forms.ProgressBar
    Friend WithEvents lblpbDraftReport As System.Windows.Forms.Label
    Friend WithEvents lblpbInProgressReport As System.Windows.Forms.Label
    Friend WithEvents lblpbTotalOpenStudies As System.Windows.Forms.Label
    Friend WithEvents lblpbFinalReport As System.Windows.Forms.Label
    Friend WithEvents gbStats As System.Windows.Forms.GroupBox
    Friend WithEvents rbAbsolute As System.Windows.Forms.RadioButton
    Friend WithEvents rbPercent As System.Windows.Forms.RadioButton
    Friend WithEvents lblDashboardTitle As System.Windows.Forms.Label
    Friend WithEvents lblInProgressCount As System.Windows.Forms.Label
    Friend WithEvents lblFinalCount As System.Windows.Forms.Label
    Friend WithEvents lblInDraftCount As System.Windows.Forms.Label
    Friend WithEvents lblTotalOpenStudiesCount As System.Windows.Forms.Label
    Friend WithEvents gbPB As System.Windows.Forms.GroupBox
    Friend WithEvents rbFinalReports As System.Windows.Forms.RadioButton
    Friend WithEvents rbDraftReports As System.Windows.Forms.RadioButton
    Friend WithEvents rbInProgressStudies As System.Windows.Forms.RadioButton
    Friend WithEvents rbTotalOpenStudies As System.Windows.Forms.RadioButton
    Friend WithEvents lblSort As System.Windows.Forms.Label
    Friend WithEvents cbxSortDashboard As System.Windows.Forms.ComboBox
    Friend WithEvents gbReportDashboardSortOrder As System.Windows.Forms.GroupBox
    Friend WithEvents lblTotalGuWuStudiesCount As System.Windows.Forms.Label
    Friend WithEvents rbTotalGuWuStudies As System.Windows.Forms.RadioButton
    Friend WithEvents pbTotalGuWuStudies As System.Windows.Forms.ProgressBar
    Friend WithEvents lblpbTotalGuWuStudies As System.Windows.Forms.Label
    Friend WithEvents lblExit As System.Windows.Forms.Label
    Friend WithEvents panButtons As System.Windows.Forms.Panel
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton3 As System.Windows.Forms.RadioButton
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents panBody As System.Windows.Forms.Panel
    Friend WithEvents lblAuditTrail As System.Windows.Forms.Label
    Friend WithEvents lblConfig As System.Windows.Forms.Label
    Friend WithEvents lblReportWriter As System.Windows.Forms.Label
    Friend WithEvents lbl1b As System.Windows.Forms.Label
    Friend WithEvents lbl1a As System.Windows.Forms.TextBox
    Friend WithEvents lbl1 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents cmdAuditTrail As System.Windows.Forms.Button
    Friend WithEvents cmdChangePassword As System.Windows.Forms.Button
    Friend WithEvents cmdLogin As System.Windows.Forms.Button
    Friend WithEvents cmdConfig As System.Windows.Forms.Button
    Friend WithEvents cmdReportWriter As System.Windows.Forms.Button
    Friend WithEvents lbl2 As System.Windows.Forms.Label
    Friend WithEvents rbDESCReportDB As System.Windows.Forms.RadioButton
    Friend WithEvents rbASCReportDB As System.Windows.Forms.RadioButton
    Friend WithEvents dgvDashboard As System.Windows.Forms.DataGridView
End Class
