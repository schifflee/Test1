<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWordCompare
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
        Me.lblDoc1 = New System.Windows.Forms.Label()
        Me.lblDoc2 = New System.Windows.Forms.Label()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.txtDoc1 = New System.Windows.Forms.TextBox()
        Me.txtDoc2 = New System.Windows.Forms.TextBox()
        Me.gbSettings = New System.Windows.Forms.GroupBox()
        Me.chkMoves = New System.Windows.Forms.CheckBox()
        Me.chkComments = New System.Windows.Forms.CheckBox()
        Me.chkFormatting = New System.Windows.Forms.CheckBox()
        Me.chkCase = New System.Windows.Forms.CheckBox()
        Me.chkWhiteSpace = New System.Windows.Forms.CheckBox()
        Me.chkTables = New System.Windows.Forms.CheckBox()
        Me.chkHeaders = New System.Windows.Forms.CheckBox()
        Me.chkFootnotes = New System.Windows.Forms.CheckBox()
        Me.chkTextboxes = New System.Windows.Forms.CheckBox()
        Me.chkFields = New System.Windows.Forms.CheckBox()
        Me.txtDocName2 = New System.Windows.Forms.TextBox()
        Me.txtDocName1 = New System.Windows.Forms.TextBox()
        Me.cmdSaveSettings = New System.Windows.Forms.Button()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.gbSettings.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblDoc1
        '
        Me.lblDoc1.AutoSize = True
        Me.lblDoc1.Location = New System.Drawing.Point(34, 110)
        Me.lblDoc1.Name = "lblDoc1"
        Me.lblDoc1.Size = New System.Drawing.Size(118, 17)
        Me.lblDoc1.TabIndex = 0
        Me.lblDoc1.Text = "Loaded Document:"
        '
        'lblDoc2
        '
        Me.lblDoc2.AutoSize = True
        Me.lblDoc2.Location = New System.Drawing.Point(34, 143)
        Me.lblDoc2.Name = "lblDoc2"
        Me.lblDoc2.Size = New System.Drawing.Size(136, 17)
        Me.lblDoc2.TabIndex = 1
        Me.lblDoc2.Text = "Compared Document:"
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOK.CausesValidation = False
        Me.cmdOK.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.cmdOK.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.Blue
        Me.cmdOK.Location = New System.Drawing.Point(34, 357)
        Me.cmdOK.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(79, 33)
        Me.cmdOK.TabIndex = 3
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancel.CausesValidation = False
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.cmdCancel.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(144, 357)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(79, 33)
        Me.cmdCancel.TabIndex = 4
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'txtDoc1
        '
        Me.txtDoc1.BackColor = System.Drawing.Color.White
        Me.txtDoc1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDoc1.Location = New System.Drawing.Point(37, 397)
        Me.txtDoc1.Name = "txtDoc1"
        Me.txtDoc1.ReadOnly = True
        Me.txtDoc1.Size = New System.Drawing.Size(390, 25)
        Me.txtDoc1.TabIndex = 5
        Me.txtDoc1.Visible = False
        '
        'txtDoc2
        '
        Me.txtDoc2.BackColor = System.Drawing.Color.White
        Me.txtDoc2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDoc2.Location = New System.Drawing.Point(70, 397)
        Me.txtDoc2.Name = "txtDoc2"
        Me.txtDoc2.ReadOnly = True
        Me.txtDoc2.Size = New System.Drawing.Size(390, 25)
        Me.txtDoc2.TabIndex = 6
        Me.txtDoc2.Visible = False
        '
        'gbSettings
        '
        Me.gbSettings.Controls.Add(Me.cmdSaveSettings)
        Me.gbSettings.Controls.Add(Me.chkFields)
        Me.gbSettings.Controls.Add(Me.chkTextboxes)
        Me.gbSettings.Controls.Add(Me.chkFootnotes)
        Me.gbSettings.Controls.Add(Me.chkHeaders)
        Me.gbSettings.Controls.Add(Me.chkTables)
        Me.gbSettings.Controls.Add(Me.chkWhiteSpace)
        Me.gbSettings.Controls.Add(Me.chkCase)
        Me.gbSettings.Controls.Add(Me.chkFormatting)
        Me.gbSettings.Controls.Add(Me.chkComments)
        Me.gbSettings.Controls.Add(Me.chkMoves)
        Me.gbSettings.Location = New System.Drawing.Point(34, 170)
        Me.gbSettings.Name = "gbSettings"
        Me.gbSettings.Size = New System.Drawing.Size(390, 169)
        Me.gbSettings.TabIndex = 7
        Me.gbSettings.TabStop = False
        Me.gbSettings.Text = "Comparison Settings"
        '
        'chkMoves
        '
        Me.chkMoves.AutoSize = True
        Me.chkMoves.Checked = True
        Me.chkMoves.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMoves.Location = New System.Drawing.Point(130, 24)
        Me.chkMoves.Name = "chkMoves"
        Me.chkMoves.Size = New System.Drawing.Size(66, 21)
        Me.chkMoves.TabIndex = 0
        Me.chkMoves.Text = "Moves"
        Me.chkMoves.UseVisualStyleBackColor = True
        Me.chkMoves.Visible = False
        '
        'chkComments
        '
        Me.chkComments.AutoSize = True
        Me.chkComments.Location = New System.Drawing.Point(6, 24)
        Me.chkComments.Name = "chkComments"
        Me.chkComments.Size = New System.Drawing.Size(89, 21)
        Me.chkComments.TabIndex = 1
        Me.chkComments.Text = "Comments"
        Me.chkComments.UseVisualStyleBackColor = True
        '
        'chkFormatting
        '
        Me.chkFormatting.AutoSize = True
        Me.chkFormatting.Location = New System.Drawing.Point(6, 51)
        Me.chkFormatting.Name = "chkFormatting"
        Me.chkFormatting.Size = New System.Drawing.Size(90, 21)
        Me.chkFormatting.TabIndex = 2
        Me.chkFormatting.Text = "Formatting"
        Me.chkFormatting.UseVisualStyleBackColor = True
        '
        'chkCase
        '
        Me.chkCase.AutoSize = True
        Me.chkCase.Checked = True
        Me.chkCase.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCase.Location = New System.Drawing.Point(6, 78)
        Me.chkCase.Name = "chkCase"
        Me.chkCase.Size = New System.Drawing.Size(109, 21)
        Me.chkCase.TabIndex = 3
        Me.chkCase.Text = "Case Changes"
        Me.chkCase.UseVisualStyleBackColor = True
        '
        'chkWhiteSpace
        '
        Me.chkWhiteSpace.AutoSize = True
        Me.chkWhiteSpace.Checked = True
        Me.chkWhiteSpace.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkWhiteSpace.Location = New System.Drawing.Point(6, 105)
        Me.chkWhiteSpace.Name = "chkWhiteSpace"
        Me.chkWhiteSpace.Size = New System.Drawing.Size(99, 21)
        Me.chkWhiteSpace.TabIndex = 4
        Me.chkWhiteSpace.Text = "White Space"
        Me.chkWhiteSpace.UseVisualStyleBackColor = True
        '
        'chkTables
        '
        Me.chkTables.AutoSize = True
        Me.chkTables.Checked = True
        Me.chkTables.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkTables.Location = New System.Drawing.Point(6, 132)
        Me.chkTables.Name = "chkTables"
        Me.chkTables.Size = New System.Drawing.Size(65, 21)
        Me.chkTables.TabIndex = 5
        Me.chkTables.Text = "Tables"
        Me.chkTables.UseVisualStyleBackColor = True
        '
        'chkHeaders
        '
        Me.chkHeaders.AutoSize = True
        Me.chkHeaders.Checked = True
        Me.chkHeaders.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkHeaders.Location = New System.Drawing.Point(192, 24)
        Me.chkHeaders.Name = "chkHeaders"
        Me.chkHeaders.Size = New System.Drawing.Size(150, 21)
        Me.chkHeaders.TabIndex = 6
        Me.chkHeaders.Text = "Headers and Footers"
        Me.chkHeaders.UseVisualStyleBackColor = True
        '
        'chkFootnotes
        '
        Me.chkFootnotes.AutoSize = True
        Me.chkFootnotes.Checked = True
        Me.chkFootnotes.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkFootnotes.Location = New System.Drawing.Point(192, 51)
        Me.chkFootnotes.Name = "chkFootnotes"
        Me.chkFootnotes.Size = New System.Drawing.Size(169, 21)
        Me.chkFootnotes.TabIndex = 7
        Me.chkFootnotes.Text = "Footnotes and Endnotes"
        Me.chkFootnotes.UseVisualStyleBackColor = True
        '
        'chkTextboxes
        '
        Me.chkTextboxes.AutoSize = True
        Me.chkTextboxes.Checked = True
        Me.chkTextboxes.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkTextboxes.Location = New System.Drawing.Point(192, 78)
        Me.chkTextboxes.Name = "chkTextboxes"
        Me.chkTextboxes.Size = New System.Drawing.Size(86, 21)
        Me.chkTextboxes.TabIndex = 8
        Me.chkTextboxes.Text = "Textboxes"
        Me.chkTextboxes.UseVisualStyleBackColor = True
        '
        'chkFields
        '
        Me.chkFields.AutoSize = True
        Me.chkFields.Checked = True
        Me.chkFields.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkFields.Location = New System.Drawing.Point(192, 105)
        Me.chkFields.Name = "chkFields"
        Me.chkFields.Size = New System.Drawing.Size(60, 21)
        Me.chkFields.TabIndex = 9
        Me.chkFields.Text = "Fields"
        Me.chkFields.UseVisualStyleBackColor = True
        '
        'txtDocName2
        '
        Me.txtDocName2.BackColor = System.Drawing.Color.White
        Me.txtDocName2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDocName2.Location = New System.Drawing.Point(176, 139)
        Me.txtDocName2.Name = "txtDocName2"
        Me.txtDocName2.ReadOnly = True
        Me.txtDocName2.Size = New System.Drawing.Size(306, 25)
        Me.txtDocName2.TabIndex = 9
        '
        'txtDocName1
        '
        Me.txtDocName1.BackColor = System.Drawing.Color.White
        Me.txtDocName1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDocName1.Location = New System.Drawing.Point(176, 108)
        Me.txtDocName1.Name = "txtDocName1"
        Me.txtDocName1.ReadOnly = True
        Me.txtDocName1.Size = New System.Drawing.Size(306, 25)
        Me.txtDocName1.TabIndex = 8
        '
        'cmdSaveSettings
        '
        Me.cmdSaveSettings.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdSaveSettings.CausesValidation = False
        Me.cmdSaveSettings.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.cmdSaveSettings.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSaveSettings.ForeColor = System.Drawing.Color.Blue
        Me.cmdSaveSettings.Location = New System.Drawing.Point(282, 105)
        Me.cmdSaveSettings.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdSaveSettings.Name = "cmdSaveSettings"
        Me.cmdSaveSettings.Size = New System.Drawing.Size(79, 47)
        Me.cmdSaveSettings.TabIndex = 10
        Me.cmdSaveSettings.Text = "&Save Settings"
        Me.cmdSaveSettings.UseVisualStyleBackColor = False
        '
        'lblTitle
        '
        Me.lblTitle.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.Location = New System.Drawing.Point(34, 20)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(448, 85)
        Me.lblTitle.TabIndex = 10
        Me.lblTitle.Text = "The two documents will be opened in Microsoft Word as a Review - Compare Document" & _
    "s document"
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'frmWordCompare
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(512, 424)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.txtDocName2)
        Me.Controls.Add(Me.txtDocName1)
        Me.Controls.Add(Me.gbSettings)
        Me.Controls.Add(Me.txtDoc2)
        Me.Controls.Add(Me.txtDoc1)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.lblDoc2)
        Me.Controls.Add(Me.lblDoc1)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "frmWordCompare"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Word Compare"
        Me.gbSettings.ResumeLayout(False)
        Me.gbSettings.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblDoc1 As System.Windows.Forms.Label
    Friend WithEvents lblDoc2 As System.Windows.Forms.Label
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents txtDoc1 As System.Windows.Forms.TextBox
    Friend WithEvents txtDoc2 As System.Windows.Forms.TextBox
    Friend WithEvents gbSettings As System.Windows.Forms.GroupBox
    Friend WithEvents chkFields As System.Windows.Forms.CheckBox
    Friend WithEvents chkTextboxes As System.Windows.Forms.CheckBox
    Friend WithEvents chkFootnotes As System.Windows.Forms.CheckBox
    Friend WithEvents chkHeaders As System.Windows.Forms.CheckBox
    Friend WithEvents chkTables As System.Windows.Forms.CheckBox
    Friend WithEvents chkWhiteSpace As System.Windows.Forms.CheckBox
    Friend WithEvents chkCase As System.Windows.Forms.CheckBox
    Friend WithEvents chkFormatting As System.Windows.Forms.CheckBox
    Friend WithEvents chkComments As System.Windows.Forms.CheckBox
    Friend WithEvents chkMoves As System.Windows.Forms.CheckBox
    Friend WithEvents txtDocName2 As System.Windows.Forms.TextBox
    Friend WithEvents txtDocName1 As System.Windows.Forms.TextBox
    Friend WithEvents cmdSaveSettings As System.Windows.Forms.Button
    Friend WithEvents lblTitle As System.Windows.Forms.Label
End Class
