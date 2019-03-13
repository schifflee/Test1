<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFinalReportPrintOptions
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
        Me.gbChoice = New System.Windows.Forms.GroupBox()
        Me.rbReadOnly = New System.Windows.Forms.RadioButton()
        Me.rbAsIs = New System.Windows.Forms.RadioButton()
        Me.gbFooters = New System.Windows.Forms.GroupBox()
        Me.rbNone = New System.Windows.Forms.RadioButton()
        Me.rbText = New System.Windows.Forms.RadioButton()
        Me.gbLocation = New System.Windows.Forms.GroupBox()
        Me.rbFooter = New System.Windows.Forms.RadioButton()
        Me.rbCenter = New System.Windows.Forms.RadioButton()
        Me.rbWaterMark = New System.Windows.Forms.RadioButton()
        Me.gbLabel = New System.Windows.Forms.GroupBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.cmdSelectAll = New System.Windows.Forms.Button()
        Me.chkDTReported = New System.Windows.Forms.CheckBox()
        Me.chkOwner = New System.Windows.Forms.CheckBox()
        Me.chkGenerator = New System.Windows.Forms.CheckBox()
        Me.chkID = New System.Windows.Forms.CheckBox()
        Me.chkDTCreated = New System.Windows.Forms.CheckBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.panRest = New System.Windows.Forms.Panel()
        Me.lblWatermark = New System.Windows.Forms.Label()
        Me.gbChoice.SuspendLayout()
        Me.gbFooters.SuspendLayout()
        Me.gbLocation.SuspendLayout()
        Me.gbLabel.SuspendLayout()
        Me.panRest.SuspendLayout()
        Me.SuspendLayout()
        '
        'gbChoice
        '
        Me.gbChoice.Controls.Add(Me.rbReadOnly)
        Me.gbChoice.Controls.Add(Me.rbAsIs)
        Me.gbChoice.Location = New System.Drawing.Point(12, 23)
        Me.gbChoice.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbChoice.Name = "gbChoice"
        Me.gbChoice.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbChoice.Size = New System.Drawing.Size(294, 92)
        Me.gbChoice.TabIndex = 0
        Me.gbChoice.TabStop = False
        Me.gbChoice.Text = "Choose a Document Security option..."
        '
        'rbReadOnly
        '
        Me.rbReadOnly.AutoSize = True
        Me.rbReadOnly.Location = New System.Drawing.Point(24, 55)
        Me.rbReadOnly.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbReadOnly.Name = "rbReadOnly"
        Me.rbReadOnly.Size = New System.Drawing.Size(218, 21)
        Me.rbReadOnly.TabIndex = 1
        Me.rbReadOnly.Text = "Generate document as read-only"
        Me.rbReadOnly.UseVisualStyleBackColor = True
        '
        'rbAsIs
        '
        Me.rbAsIs.AutoSize = True
        Me.rbAsIs.Checked = True
        Me.rbAsIs.Location = New System.Drawing.Point(24, 26)
        Me.rbAsIs.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbAsIs.Name = "rbAsIs"
        Me.rbAsIs.Size = New System.Drawing.Size(171, 21)
        Me.rbAsIs.TabIndex = 0
        Me.rbAsIs.TabStop = True
        Me.rbAsIs.Text = "Generate document as is"
        Me.rbAsIs.UseVisualStyleBackColor = True
        '
        'gbFooters
        '
        Me.gbFooters.Controls.Add(Me.rbNone)
        Me.gbFooters.Controls.Add(Me.rbText)
        Me.gbFooters.Controls.Add(Me.gbLocation)
        Me.gbFooters.Controls.Add(Me.rbWaterMark)
        Me.gbFooters.Location = New System.Drawing.Point(0, 0)
        Me.gbFooters.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbFooters.Name = "gbFooters"
        Me.gbFooters.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbFooters.Size = New System.Drawing.Size(320, 87)
        Me.gbFooters.TabIndex = 1
        Me.gbFooters.TabStop = False
        Me.gbFooters.Text = "Choose a Document Label option..."
        '
        'rbNone
        '
        Me.rbNone.AutoSize = True
        Me.rbNone.Checked = True
        Me.rbNone.Location = New System.Drawing.Point(24, 23)
        Me.rbNone.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbNone.Name = "rbNone"
        Me.rbNone.Size = New System.Drawing.Size(58, 21)
        Me.rbNone.TabIndex = 0
        Me.rbNone.TabStop = True
        Me.rbNone.Text = "None"
        Me.rbNone.UseVisualStyleBackColor = True
        '
        'rbText
        '
        Me.rbText.AutoSize = True
        Me.rbText.Location = New System.Drawing.Point(101, 26)
        Me.rbText.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbText.Name = "rbText"
        Me.rbText.Size = New System.Drawing.Size(114, 21)
        Me.rbText.TabIndex = 3
        Me.rbText.Text = "Text (In Footer)"
        Me.rbText.UseVisualStyleBackColor = True
        Me.rbText.Visible = False
        '
        'gbLocation
        '
        Me.gbLocation.Controls.Add(Me.rbFooter)
        Me.gbLocation.Controls.Add(Me.rbCenter)
        Me.gbLocation.Location = New System.Drawing.Point(232, 14)
        Me.gbLocation.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbLocation.Name = "gbLocation"
        Me.gbLocation.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbLocation.Size = New System.Drawing.Size(147, 80)
        Me.gbLocation.TabIndex = 2
        Me.gbLocation.TabStop = False
        Me.gbLocation.Text = "Choose a Watermark location"
        Me.gbLocation.Visible = False
        '
        'rbFooter
        '
        Me.rbFooter.AutoSize = True
        Me.rbFooter.Location = New System.Drawing.Point(24, 55)
        Me.rbFooter.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbFooter.Name = "rbFooter"
        Me.rbFooter.Size = New System.Drawing.Size(64, 21)
        Me.rbFooter.TabIndex = 1
        Me.rbFooter.TabStop = True
        Me.rbFooter.Text = "Footer"
        Me.rbFooter.UseVisualStyleBackColor = True
        '
        'rbCenter
        '
        Me.rbCenter.AutoSize = True
        Me.rbCenter.Checked = True
        Me.rbCenter.Location = New System.Drawing.Point(24, 26)
        Me.rbCenter.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbCenter.Name = "rbCenter"
        Me.rbCenter.Size = New System.Drawing.Size(142, 21)
        Me.rbCenter.TabIndex = 0
        Me.rbCenter.TabStop = True
        Me.rbCenter.Text = "Center of document"
        Me.rbCenter.UseVisualStyleBackColor = True
        '
        'rbWaterMark
        '
        Me.rbWaterMark.AutoSize = True
        Me.rbWaterMark.Location = New System.Drawing.Point(24, 52)
        Me.rbWaterMark.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbWaterMark.Name = "rbWaterMark"
        Me.rbWaterMark.Size = New System.Drawing.Size(90, 21)
        Me.rbWaterMark.TabIndex = 1
        Me.rbWaterMark.Text = "Watermark"
        Me.rbWaterMark.UseVisualStyleBackColor = True
        '
        'gbLabel
        '
        Me.gbLabel.Controls.Add(Me.lblWatermark)
        Me.gbLabel.Controls.Add(Me.Button1)
        Me.gbLabel.Controls.Add(Me.cmdSelectAll)
        Me.gbLabel.Controls.Add(Me.chkDTReported)
        Me.gbLabel.Controls.Add(Me.chkOwner)
        Me.gbLabel.Controls.Add(Me.chkGenerator)
        Me.gbLabel.Controls.Add(Me.chkID)
        Me.gbLabel.Controls.Add(Me.chkDTCreated)
        Me.gbLabel.Location = New System.Drawing.Point(0, 99)
        Me.gbLabel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbLabel.Name = "gbLabel"
        Me.gbLabel.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbLabel.Size = New System.Drawing.Size(320, 156)
        Me.gbLabel.TabIndex = 2
        Me.gbLabel.TabStop = False
        Me.gbLabel.Text = "Choose Document Watermari Content"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Gainsboro
        Me.Button1.FlatAppearance.BorderSize = 0
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.Button1.Location = New System.Drawing.Point(191, 99)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(103, 28)
        Me.Button1.TabIndex = 73
        Me.Button1.Text = "Deselect All"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'cmdSelectAll
        '
        Me.cmdSelectAll.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdSelectAll.FlatAppearance.BorderSize = 0
        Me.cmdSelectAll.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdSelectAll.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSelectAll.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdSelectAll.Location = New System.Drawing.Point(191, 65)
        Me.cmdSelectAll.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdSelectAll.Name = "cmdSelectAll"
        Me.cmdSelectAll.Size = New System.Drawing.Size(103, 28)
        Me.cmdSelectAll.TabIndex = 72
        Me.cmdSelectAll.Text = "Select All"
        Me.cmdSelectAll.UseVisualStyleBackColor = True
        '
        'chkDTReported
        '
        Me.chkDTReported.AutoSize = True
        Me.chkDTReported.Location = New System.Drawing.Point(215, 61)
        Me.chkDTReported.Name = "chkDTReported"
        Me.chkDTReported.Size = New System.Drawing.Size(265, 21)
        Me.chkDTReported.TabIndex = 2
        Me.chkDTReported.Text = "Date/Time Stamp: Report Printed/Output"
        Me.chkDTReported.UseVisualStyleBackColor = True
        Me.chkDTReported.Visible = False
        '
        'chkOwner
        '
        Me.chkOwner.AutoSize = True
        Me.chkOwner.Location = New System.Drawing.Point(24, 119)
        Me.chkOwner.Name = "chkOwner"
        Me.chkOwner.Size = New System.Drawing.Size(128, 21)
        Me.chkOwner.TabIndex = 4
        Me.chkOwner.Text = "Document Owner"
        Me.chkOwner.UseVisualStyleBackColor = True
        '
        'chkGenerator
        '
        Me.chkGenerator.AutoSize = True
        Me.chkGenerator.Location = New System.Drawing.Point(24, 92)
        Me.chkGenerator.Name = "chkGenerator"
        Me.chkGenerator.Size = New System.Drawing.Size(149, 21)
        Me.chkGenerator.TabIndex = 3
        Me.chkGenerator.Text = "Document Generator"
        Me.chkGenerator.UseVisualStyleBackColor = True
        '
        'chkID
        '
        Me.chkID.AutoSize = True
        Me.chkID.Location = New System.Drawing.Point(24, 65)
        Me.chkID.Name = "chkID"
        Me.chkID.Size = New System.Drawing.Size(102, 21)
        Me.chkID.TabIndex = 0
        Me.chkID.Text = "Document ID"
        Me.chkID.UseVisualStyleBackColor = True
        '
        'chkDTCreated
        '
        Me.chkDTCreated.AutoSize = True
        Me.chkDTCreated.Location = New System.Drawing.Point(215, 34)
        Me.chkDTCreated.Name = "chkDTCreated"
        Me.chkDTCreated.Size = New System.Drawing.Size(225, 21)
        Me.chkDTCreated.TabIndex = 1
        Me.chkDTCreated.Text = "Date/Time Stamp: Report Created"
        Me.chkDTCreated.UseVisualStyleBackColor = True
        Me.chkDTCreated.Visible = False
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
        Me.cmdCancel.Location = New System.Drawing.Point(167, 264)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 33)
        Me.cmdCancel.TabIndex = 4
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
        Me.cmdOK.Location = New System.Drawing.Point(33, 263)
        Me.cmdOK.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(75, 33)
        Me.cmdOK.TabIndex = 3
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'panRest
        '
        Me.panRest.Controls.Add(Me.gbFooters)
        Me.panRest.Controls.Add(Me.cmdCancel)
        Me.panRest.Controls.Add(Me.gbLabel)
        Me.panRest.Controls.Add(Me.cmdOK)
        Me.panRest.Location = New System.Drawing.Point(12, 122)
        Me.panRest.Name = "panRest"
        Me.panRest.Size = New System.Drawing.Size(334, 313)
        Me.panRest.TabIndex = 43
        '
        'lblWatermark
        '
        Me.lblWatermark.Location = New System.Drawing.Point(21, 23)
        Me.lblWatermark.Name = "lblWatermark"
        Me.lblWatermark.Size = New System.Drawing.Size(283, 39)
        Me.lblWatermark.TabIndex = 5
        Me.lblWatermark.Text = "Watermark contains 'DRAFT', Date/Time Stamp, plus additional options chosen below" & _
    ""
        '
        'frmFinalReportPrintOptions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(359, 447)
        Me.ControlBox = False
        Me.Controls.Add(Me.panRest)
        Me.Controls.Add(Me.gbChoice)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "frmFinalReportPrintOptions"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Word/Print Options..."
        Me.gbChoice.ResumeLayout(False)
        Me.gbChoice.PerformLayout()
        Me.gbFooters.ResumeLayout(False)
        Me.gbFooters.PerformLayout()
        Me.gbLocation.ResumeLayout(False)
        Me.gbLocation.PerformLayout()
        Me.gbLabel.ResumeLayout(False)
        Me.gbLabel.PerformLayout()
        Me.panRest.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbChoice As System.Windows.Forms.GroupBox
    Friend WithEvents rbReadOnly As System.Windows.Forms.RadioButton
    Friend WithEvents rbAsIs As System.Windows.Forms.RadioButton
    Friend WithEvents gbFooters As System.Windows.Forms.GroupBox
    Friend WithEvents rbNone As System.Windows.Forms.RadioButton
    Friend WithEvents gbLabel As System.Windows.Forms.GroupBox
    Friend WithEvents chkOwner As System.Windows.Forms.CheckBox
    Friend WithEvents chkGenerator As System.Windows.Forms.CheckBox
    Friend WithEvents chkID As System.Windows.Forms.CheckBox
    Friend WithEvents chkDTCreated As System.Windows.Forms.CheckBox
    Friend WithEvents gbLocation As System.Windows.Forms.GroupBox
    Friend WithEvents rbFooter As System.Windows.Forms.RadioButton
    Friend WithEvents rbCenter As System.Windows.Forms.RadioButton
    Friend WithEvents rbText As System.Windows.Forms.RadioButton
    Friend WithEvents rbWaterMark As System.Windows.Forms.RadioButton
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents panRest As System.Windows.Forms.Panel
    Friend WithEvents chkDTReported As System.Windows.Forms.CheckBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents cmdSelectAll As System.Windows.Forms.Button
    Friend WithEvents lblWatermark As System.Windows.Forms.Label
End Class
