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
        Me.components = New System.ComponentModel.Container
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmWordStatement))
        Me.cmiFieldCode = New System.Windows.Forms.ToolStripMenuItem
        Me.dgvReportStatements = New System.Windows.Forms.DataGridView
        Me.cmdFieldCode = New System.Windows.Forms.Button
        Me.cmdExit = New System.Windows.Forms.Button
        Me.cmdOpen = New System.Windows.Forms.Button
        Me.cmsfrmWord = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdRefresh = New System.Windows.Forms.Button
        Me.lblRefresh = New System.Windows.Forms.Label
        Me.cmdAddStatement = New System.Windows.Forms.Button
        Me.lblDGV = New System.Windows.Forms.Label
        Me.lblEditTitles = New System.Windows.Forms.Label
        Me.cmdEditStatements = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.CHARTITLE = New System.Windows.Forms.TextBox
        Me.lblSection = New System.Windows.Forms.Label
        Me.lblStatus = New System.Windows.Forms.Label
        Me.lblStatus2 = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.chkCreateNew = New System.Windows.Forms.CheckBox
        Me.pan1 = New System.Windows.Forms.Panel
        Me.cmdWord1 = New System.Windows.Forms.Button
        Me.lblAfter = New System.Windows.Forms.Label
        Me.lblBefore = New System.Windows.Forms.Label
        Me.pan2 = New System.Windows.Forms.Panel
        Me.cmdExit2 = New System.Windows.Forms.Button
        Me.cmdWord = New System.Windows.Forms.Button
        Me.lblEdit = New System.Windows.Forms.Label
        Me.txtBefore = New System.Windows.Forms.TextBox
        Me.txtAfter = New System.Windows.Forms.TextBox
        Me.panEdraw = New System.Windows.Forms.Panel
        Me.ov1 = New AxOfficeViewer.AxOfficeViewer
        Me.panRefresh = New System.Windows.Forms.Panel
        Me.cmdOpenPDF = New System.Windows.Forms.Button
        CType(Me.dgvReportStatements, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.cmsfrmWord.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.pan1.SuspendLayout()
        Me.pan2.SuspendLayout()
        Me.panEdraw.SuspendLayout()
        CType(Me.ov1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panRefresh.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmiFieldCode
        '
        Me.cmiFieldCode.Name = "cmiFieldCode"
        Me.cmiFieldCode.Size = New System.Drawing.Size(171, 22)
        Me.cmiFieldCode.Text = "Insert Field Code..."
        '
        'dgvReportStatements
        '
        Me.dgvReportStatements.AllowUserToAddRows = False
        Me.dgvReportStatements.AllowUserToDeleteRows = False
        Me.dgvReportStatements.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvReportStatements.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvReportStatements.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvReportStatements.DefaultCellStyle = DataGridViewCellStyle4
        Me.dgvReportStatements.Location = New System.Drawing.Point(2, 184)
        Me.dgvReportStatements.Name = "dgvReportStatements"
        Me.dgvReportStatements.Size = New System.Drawing.Size(214, 146)
        Me.dgvReportStatements.TabIndex = 33
        '
        'cmdFieldCode
        '
        Me.cmdFieldCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdFieldCode.ForeColor = System.Drawing.Color.Green
        Me.cmdFieldCode.Location = New System.Drawing.Point(109, 3)
        Me.cmdFieldCode.Name = "cmdFieldCode"
        Me.cmdFieldCode.Size = New System.Drawing.Size(103, 47)
        Me.cmdFieldCode.TabIndex = 25
        Me.cmdFieldCode.Text = "Enter &Field Code..."
        Me.cmdFieldCode.UseVisualStyleBackColor = True
        '
        'cmdExit
        '
        Me.cmdExit.CausesValidation = False
        Me.cmdExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ForeColor = System.Drawing.Color.Red
        Me.cmdExit.Location = New System.Drawing.Point(53, 434)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(103, 47)
        Me.cmdExit.TabIndex = 29
        Me.cmdExit.Text = "E&xit"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'cmdOpen
        '
        Me.cmdOpen.Location = New System.Drawing.Point(246, 584)
        Me.cmdOpen.Name = "cmdOpen"
        Me.cmdOpen.Size = New System.Drawing.Size(91, 36)
        Me.cmdOpen.TabIndex = 26
        Me.cmdOpen.Text = "&Open"
        Me.cmdOpen.UseVisualStyleBackColor = True
        Me.cmdOpen.Visible = False
        '
        'cmsfrmWord
        '
        Me.cmsfrmWord.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.cmiFieldCode})
        Me.cmsfrmWord.Name = "cmsfrmWord"
        Me.cmsfrmWord.Size = New System.Drawing.Size(172, 26)
        '
        'cmdSave
        '
        Me.cmdSave.Enabled = False
        Me.cmdSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ForeColor = System.Drawing.Color.Blue
        Me.cmdSave.Location = New System.Drawing.Point(53, 381)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(103, 47)
        Me.cmdSave.TabIndex = 27
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'cmdRefresh
        '
        Me.cmdRefresh.CausesValidation = False
        Me.cmdRefresh.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdRefresh.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRefresh.ForeColor = System.Drawing.Color.Purple
        Me.cmdRefresh.Location = New System.Drawing.Point(53, 60)
        Me.cmdRefresh.Name = "cmdRefresh"
        Me.cmdRefresh.Size = New System.Drawing.Size(103, 28)
        Me.cmdRefresh.TabIndex = 30
        Me.cmdRefresh.Text = "&Refresh"
        Me.cmdRefresh.UseVisualStyleBackColor = True
        '
        'lblRefresh
        '
        Me.lblRefresh.AutoSize = True
        Me.lblRefresh.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRefresh.ForeColor = System.Drawing.Color.Blue
        Me.lblRefresh.Location = New System.Drawing.Point(35, 2)
        Me.lblRefresh.MaximumSize = New System.Drawing.Size(149, 146)
        Me.lblRefresh.MinimumSize = New System.Drawing.Size(149, 50)
        Me.lblRefresh.Name = "lblRefresh"
        Me.lblRefresh.Size = New System.Drawing.Size(149, 50)
        Me.lblRefresh.TabIndex = 35
        Me.lblRefresh.Text = "Click Refresh if Word object menus seem to be frozen"
        Me.lblRefresh.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'cmdAddStatement
        '
        Me.cmdAddStatement.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddStatement.ForeColor = System.Drawing.Color.Black
        Me.cmdAddStatement.Location = New System.Drawing.Point(53, 56)
        Me.cmdAddStatement.Name = "cmdAddStatement"
        Me.cmdAddStatement.Size = New System.Drawing.Size(103, 47)
        Me.cmdAddStatement.TabIndex = 36
        Me.cmdAddStatement.Text = "Create &New Statement..."
        Me.cmdAddStatement.UseVisualStyleBackColor = True
        '
        'lblDGV
        '
        Me.lblDGV.AutoSize = True
        Me.lblDGV.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDGV.ForeColor = System.Drawing.Color.Red
        Me.lblDGV.Location = New System.Drawing.Point(3, 333)
        Me.lblDGV.MaximumSize = New System.Drawing.Size(214, 146)
        Me.lblDGV.Name = "lblDGV"
        Me.lblDGV.Size = New System.Drawing.Size(207, 26)
        Me.lblDGV.TabIndex = 37
        Me.lblDGV.Text = "Save must be clicked before a new row may be selected"
        Me.lblDGV.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.lblDGV.Visible = False
        '
        'lblEditTitles
        '
        Me.lblEditTitles.AutoSize = True
        Me.lblEditTitles.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEditTitles.ForeColor = System.Drawing.Color.Blue
        Me.lblEditTitles.Location = New System.Drawing.Point(2, 3)
        Me.lblEditTitles.MaximumSize = New System.Drawing.Size(149, 146)
        Me.lblEditTitles.MinimumSize = New System.Drawing.Size(143, 0)
        Me.lblEditTitles.Name = "lblEditTitles"
        Me.lblEditTitles.Size = New System.Drawing.Size(143, 13)
        Me.lblEditTitles.TabIndex = 38
        Me.lblEditTitles.Text = "Edit Statement Titles"
        Me.lblEditTitles.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'cmdEditStatements
        '
        Me.cmdEditStatements.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdEditStatements.ForeColor = System.Drawing.Color.Blue
        Me.cmdEditStatements.Location = New System.Drawing.Point(2, 43)
        Me.cmdEditStatements.Name = "cmdEditStatements"
        Me.cmdEditStatements.Size = New System.Drawing.Size(64, 25)
        Me.cmdEditStatements.TabIndex = 39
        Me.cmdEditStatements.Text = "&Edit"
        Me.cmdEditStatements.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.CausesValidation = False
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Enabled = False
        Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(81, 43)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(64, 25)
        Me.cmdCancel.TabIndex = 40
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'CHARTITLE
        '
        Me.CHARTITLE.Enabled = False
        Me.CHARTITLE.Location = New System.Drawing.Point(2, 19)
        Me.CHARTITLE.Name = "CHARTITLE"
        Me.CHARTITLE.ReadOnly = True
        Me.CHARTITLE.Size = New System.Drawing.Size(207, 20)
        Me.CHARTITLE.TabIndex = 41
        '
        'lblSection
        '
        Me.lblSection.AutoSize = True
        Me.lblSection.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSection.ForeColor = System.Drawing.Color.Blue
        Me.lblSection.Location = New System.Drawing.Point(222, 3)
        Me.lblSection.MaximumSize = New System.Drawing.Size(663, 100)
        Me.lblSection.Name = "lblSection"
        Me.lblSection.Size = New System.Drawing.Size(173, 13)
        Me.lblSection.TabIndex = 42
        Me.lblSection.Text = "Report Statements for the Section: "
        Me.lblSection.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(446, 596)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(37, 13)
        Me.lblStatus.TabIndex = 43
        Me.lblStatus.Text = "Status"
        Me.lblStatus.Visible = False
        '
        'lblStatus2
        '
        Me.lblStatus2.AutoSize = True
        Me.lblStatus2.Location = New System.Drawing.Point(374, 607)
        Me.lblStatus2.Name = "lblStatus2"
        Me.lblStatus2.Size = New System.Drawing.Size(37, 13)
        Me.lblStatus2.TabIndex = 44
        Me.lblStatus2.Text = "Status"
        Me.lblStatus2.Visible = False
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.cmdCancel)
        Me.Panel1.Controls.Add(Me.lblEditTitles)
        Me.Panel1.Controls.Add(Me.CHARTITLE)
        Me.Panel1.Controls.Add(Me.cmdEditStatements)
        Me.Panel1.Location = New System.Drawing.Point(2, 109)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(214, 74)
        Me.Panel1.TabIndex = 45
        '
        'chkCreateNew
        '
        Me.chkCreateNew.AutoSize = True
        Me.chkCreateNew.Location = New System.Drawing.Point(423, 633)
        Me.chkCreateNew.Name = "chkCreateNew"
        Me.chkCreateNew.Size = New System.Drawing.Size(82, 17)
        Me.chkCreateNew.TabIndex = 46
        Me.chkCreateNew.Text = "Create New"
        Me.chkCreateNew.UseVisualStyleBackColor = True
        Me.chkCreateNew.Visible = False
        '
        'pan1
        '
        Me.pan1.Controls.Add(Me.cmdWord1)
        Me.pan1.Controls.Add(Me.lblAfter)
        Me.pan1.Controls.Add(Me.lblBefore)
        Me.pan1.Controls.Add(Me.cmdExit)
        Me.pan1.Controls.Add(Me.cmdSave)
        Me.pan1.Controls.Add(Me.dgvReportStatements)
        Me.pan1.Controls.Add(Me.Panel1)
        Me.pan1.Controls.Add(Me.cmdFieldCode)
        Me.pan1.Controls.Add(Me.cmdAddStatement)
        Me.pan1.Controls.Add(Me.lblDGV)
        Me.pan1.Location = New System.Drawing.Point(0, 0)
        Me.pan1.Name = "pan1"
        Me.pan1.Size = New System.Drawing.Size(219, 484)
        Me.pan1.TabIndex = 47
        '
        'cmdWord1
        '
        Me.cmdWord1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdWord1.ForeColor = System.Drawing.Color.Green
        Me.cmdWord1.Location = New System.Drawing.Point(3, 3)
        Me.cmdWord1.Name = "cmdWord1"
        Me.cmdWord1.Size = New System.Drawing.Size(103, 47)
        Me.cmdWord1.TabIndex = 55
        Me.cmdWord1.Text = "Open in &Word"
        Me.cmdWord1.UseVisualStyleBackColor = True
        '
        'lblAfter
        '
        Me.lblAfter.AutoSize = True
        Me.lblAfter.Location = New System.Drawing.Point(117, 365)
        Me.lblAfter.Name = "lblAfter"
        Me.lblAfter.Size = New System.Drawing.Size(39, 13)
        Me.lblAfter.TabIndex = 54
        Me.lblAfter.Text = "Label2"
        Me.lblAfter.Visible = False
        '
        'lblBefore
        '
        Me.lblBefore.AutoSize = True
        Me.lblBefore.Location = New System.Drawing.Point(2, 365)
        Me.lblBefore.Name = "lblBefore"
        Me.lblBefore.Size = New System.Drawing.Size(39, 13)
        Me.lblBefore.TabIndex = 53
        Me.lblBefore.Text = "Label2"
        Me.lblBefore.Visible = False
        '
        'pan2
        '
        Me.pan2.Controls.Add(Me.cmdOpenPDF)
        Me.pan2.Controls.Add(Me.cmdExit2)
        Me.pan2.Controls.Add(Me.cmdWord)
        Me.pan2.Location = New System.Drawing.Point(576, 403)
        Me.pan2.Name = "pan2"
        Me.pan2.Size = New System.Drawing.Size(219, 136)
        Me.pan2.TabIndex = 48
        Me.pan2.Visible = False
        '
        'cmdExit2
        '
        Me.cmdExit2.CausesValidation = False
        Me.cmdExit2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdExit2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit2.ForeColor = System.Drawing.Color.Red
        Me.cmdExit2.Location = New System.Drawing.Point(60, 72)
        Me.cmdExit2.Name = "cmdExit2"
        Me.cmdExit2.Size = New System.Drawing.Size(103, 47)
        Me.cmdExit2.TabIndex = 30
        Me.cmdExit2.Text = "E&xit"
        Me.cmdExit2.UseVisualStyleBackColor = True
        '
        'cmdWord
        '
        Me.cmdWord.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdWord.ForeColor = System.Drawing.Color.Green
        Me.cmdWord.Location = New System.Drawing.Point(60, 19)
        Me.cmdWord.Name = "cmdWord"
        Me.cmdWord.Size = New System.Drawing.Size(103, 47)
        Me.cmdWord.TabIndex = 26
        Me.cmdWord.Text = "Open in &Word"
        Me.cmdWord.UseVisualStyleBackColor = True
        '
        'lblEdit
        '
        Me.lblEdit.AutoSize = True
        Me.lblEdit.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEdit.ForeColor = System.Drawing.Color.Red
        Me.lblEdit.Location = New System.Drawing.Point(33, 578)
        Me.lblEdit.MaximumSize = New System.Drawing.Size(134, 75)
        Me.lblEdit.MinimumSize = New System.Drawing.Size(134, 75)
        Me.lblEdit.Name = "lblEdit"
        Me.lblEdit.Size = New System.Drawing.Size(134, 75)
        Me.lblEdit.TabIndex = 49
        Me.lblEdit.Text = "Non-Edit Mode"
        Me.lblEdit.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblEdit.Visible = False
        '
        'txtBefore
        '
        Me.txtBefore.Location = New System.Drawing.Point(8, 618)
        Me.txtBefore.Name = "txtBefore"
        Me.txtBefore.Size = New System.Drawing.Size(108, 20)
        Me.txtBefore.TabIndex = 50
        Me.txtBefore.Visible = False
        '
        'txtAfter
        '
        Me.txtAfter.Location = New System.Drawing.Point(8, 644)
        Me.txtAfter.Name = "txtAfter"
        Me.txtAfter.Size = New System.Drawing.Size(108, 20)
        Me.txtAfter.TabIndex = 51
        Me.txtAfter.Visible = False
        '
        'panEdraw
        '
        Me.panEdraw.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.panEdraw.Controls.Add(Me.ov1)
        Me.panEdraw.Location = New System.Drawing.Point(225, 22)
        Me.panEdraw.Name = "panEdraw"
        Me.panEdraw.Size = New System.Drawing.Size(710, 202)
        Me.panEdraw.TabIndex = 52
        '
        'ov1
        '
        Me.ov1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ov1.Enabled = True
        Me.ov1.Location = New System.Drawing.Point(0, 0)
        Me.ov1.Name = "ov1"
        Me.ov1.OcxState = CType(resources.GetObject("ov1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ov1.Size = New System.Drawing.Size(710, 202)
        Me.ov1.TabIndex = 0
        '
        'panRefresh
        '
        Me.panRefresh.Controls.Add(Me.lblRefresh)
        Me.panRefresh.Controls.Add(Me.cmdRefresh)
        Me.panRefresh.Location = New System.Drawing.Point(0, 487)
        Me.panRefresh.Name = "panRefresh"
        Me.panRefresh.Size = New System.Drawing.Size(219, 91)
        Me.panRefresh.TabIndex = 53
        '
        'cmdOpenPDF
        '
        Me.cmdOpenPDF.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOpenPDF.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.cmdOpenPDF.Location = New System.Drawing.Point(113, 19)
        Me.cmdOpenPDF.Name = "cmdOpenPDF"
        Me.cmdOpenPDF.Size = New System.Drawing.Size(103, 47)
        Me.cmdOpenPDF.TabIndex = 56
        Me.cmdOpenPDF.Text = "Open PDF"
        Me.cmdOpenPDF.UseVisualStyleBackColor = True
        Me.cmdOpenPDF.Visible = False
        '
        'frmWordStatement
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.MistyRose
        Me.ClientSize = New System.Drawing.Size(940, 677)
        Me.ControlBox = False
        Me.Controls.Add(Me.panRefresh)
        Me.Controls.Add(Me.txtBefore)
        Me.Controls.Add(Me.panEdraw)
        Me.Controls.Add(Me.txtAfter)
        Me.Controls.Add(Me.lblEdit)
        Me.Controls.Add(Me.chkCreateNew)
        Me.Controls.Add(Me.pan2)
        Me.Controls.Add(Me.pan1)
        Me.Controls.Add(Me.lblStatus2)
        Me.Controls.Add(Me.cmdOpen)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.lblSection)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmWordStatement"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmWordStatement"
        CType(Me.dgvReportStatements, System.ComponentModel.ISupportInitialize).EndInit()
        Me.cmsfrmWord.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.pan1.ResumeLayout(False)
        Me.pan1.PerformLayout()
        Me.pan2.ResumeLayout(False)
        Me.panEdraw.ResumeLayout(False)
        CType(Me.ov1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panRefresh.ResumeLayout(False)
        Me.panRefresh.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmiFieldCode As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents dgvReportStatements As System.Windows.Forms.DataGridView
    Friend WithEvents cmdFieldCode As System.Windows.Forms.Button
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents cmdOpen As System.Windows.Forms.Button
    Friend WithEvents cmsfrmWord As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdRefresh As System.Windows.Forms.Button
    Friend WithEvents lblRefresh As System.Windows.Forms.Label
    Friend WithEvents cmdAddStatement As System.Windows.Forms.Button
    Friend WithEvents lblDGV As System.Windows.Forms.Label
    Friend WithEvents lblEditTitles As System.Windows.Forms.Label
    Friend WithEvents cmdEditStatements As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents CHARTITLE As System.Windows.Forms.TextBox
    Friend WithEvents lblSection As System.Windows.Forms.Label
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents lblStatus2 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents chkCreateNew As System.Windows.Forms.CheckBox
    Friend WithEvents pan1 As System.Windows.Forms.Panel
    Friend WithEvents pan2 As System.Windows.Forms.Panel
    Friend WithEvents cmdWord As System.Windows.Forms.Button
    Friend WithEvents cmdExit2 As System.Windows.Forms.Button
    Friend WithEvents lblEdit As System.Windows.Forms.Label
    Friend WithEvents txtBefore As System.Windows.Forms.TextBox
    Friend WithEvents txtAfter As System.Windows.Forms.TextBox
    Friend WithEvents panEdraw As System.Windows.Forms.Panel
    Friend WithEvents ov1 As AxOfficeViewer.AxOfficeViewer
    Friend WithEvents lblAfter As System.Windows.Forms.Label
    Friend WithEvents lblBefore As System.Windows.Forms.Label
    Friend WithEvents panRefresh As System.Windows.Forms.Panel
    Friend WithEvents cmdWord1 As System.Windows.Forms.Button
    Friend WithEvents cmdOpenPDF As System.Windows.Forms.Button
End Class
