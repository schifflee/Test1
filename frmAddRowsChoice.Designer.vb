<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAddRowsChoice
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
        Me.gbRecovery = New System.Windows.Forms.GroupBox()
        Me.rbQC = New System.Windows.Forms.RadioButton()
        Me.rbRS = New System.Windows.Forms.RadioButton()
        Me.rbPES = New System.Windows.Forms.RadioButton()
        Me.chkIgnore = New System.Windows.Forms.CheckBox()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.lblH = New System.Windows.Forms.Label()
        Me.gbULots = New System.Windows.Forms.GroupBox()
        Me.rbLot10 = New System.Windows.Forms.RadioButton()
        Me.rbLot9 = New System.Windows.Forms.RadioButton()
        Me.rbLot8 = New System.Windows.Forms.RadioButton()
        Me.rbLot7 = New System.Windows.Forms.RadioButton()
        Me.rbLot6 = New System.Windows.Forms.RadioButton()
        Me.rbLot5 = New System.Windows.Forms.RadioButton()
        Me.rbLot4 = New System.Windows.Forms.RadioButton()
        Me.rbLot3 = New System.Windows.Forms.RadioButton()
        Me.rbLot2 = New System.Windows.Forms.RadioButton()
        Me.rbLot1 = New System.Windows.Forms.RadioButton()
        Me.gbStockSoln = New System.Windows.Forms.GroupBox()
        Me.rbNewStockSoln = New System.Windows.Forms.RadioButton()
        Me.rbOldStockSoln = New System.Windows.Forms.RadioButton()
        Me.gbSpikingSoln = New System.Windows.Forms.GroupBox()
        Me.rbNewSpikingSoln = New System.Windows.Forms.RadioButton()
        Me.rbOldSpikingSoln = New System.Windows.Forms.RadioButton()
        Me.gbCarryOver = New System.Windows.Forms.GroupBox()
        Me.rbBlank = New System.Windows.Forms.RadioButton()
        Me.rbULOQ = New System.Windows.Forms.RadioButton()
        Me.rbLLOQ = New System.Windows.Forms.RadioButton()
        Me.gbLongTermQC = New System.Windows.Forms.GroupBox()
        Me.rbLTQCLTA = New System.Windows.Forms.RadioButton()
        Me.rbLTQCT0 = New System.Windows.Forms.RadioButton()
        Me.gbAdHocStabComp = New System.Windows.Forms.GroupBox()
        Me.txtAdHocStabComp = New System.Windows.Forms.TextBox()
        Me.lblAdHocStabComp = New System.Windows.Forms.Label()
        Me.gbRecovery.SuspendLayout()
        Me.gbULots.SuspendLayout()
        Me.gbStockSoln.SuspendLayout()
        Me.gbSpikingSoln.SuspendLayout()
        Me.gbCarryOver.SuspendLayout()
        Me.gbLongTermQC.SuspendLayout()
        Me.gbAdHocStabComp.SuspendLayout()
        Me.SuspendLayout()
        '
        'gbRecovery
        '
        Me.gbRecovery.Controls.Add(Me.rbQC)
        Me.gbRecovery.Controls.Add(Me.rbRS)
        Me.gbRecovery.Controls.Add(Me.rbPES)
        Me.gbRecovery.Location = New System.Drawing.Point(142, 175)
        Me.gbRecovery.Name = "gbRecovery"
        Me.gbRecovery.Size = New System.Drawing.Size(275, 83)
        Me.gbRecovery.TabIndex = 0
        Me.gbRecovery.TabStop = False
        Me.gbRecovery.Text = "Recovery"
        '
        'rbQC
        '
        Me.rbQC.AutoSize = True
        Me.rbQC.Location = New System.Drawing.Point(143, 51)
        Me.rbQC.Name = "rbQC"
        Me.rbQC.Size = New System.Drawing.Size(101, 21)
        Me.rbQC.TabIndex = 2
        Me.rbQC.TabStop = True
        Me.rbQC.Text = "QC Standard"
        Me.rbQC.UseVisualStyleBackColor = True
        '
        'rbRS
        '
        Me.rbRS.AutoSize = True
        Me.rbRS.Location = New System.Drawing.Point(23, 51)
        Me.rbRS.Name = "rbRS"
        Me.rbRS.Size = New System.Drawing.Size(130, 21)
        Me.rbRS.TabIndex = 1
        Me.rbRS.TabStop = True
        Me.rbRS.Text = "Recovery Solution"
        Me.rbRS.UseVisualStyleBackColor = True
        '
        'rbPES
        '
        Me.rbPES.AutoSize = True
        Me.rbPES.Location = New System.Drawing.Point(23, 24)
        Me.rbPES.Name = "rbPES"
        Me.rbPES.Size = New System.Drawing.Size(199, 21)
        Me.rbPES.TabIndex = 0
        Me.rbPES.TabStop = True
        Me.rbPES.Text = "Post-Extraction Spike Solution"
        Me.rbPES.UseVisualStyleBackColor = True
        '
        'chkIgnore
        '
        Me.chkIgnore.AutoSize = True
        Me.chkIgnore.Location = New System.Drawing.Point(198, 148)
        Me.chkIgnore.Name = "chkIgnore"
        Me.chkIgnore.Size = New System.Drawing.Size(155, 21)
        Me.chkIgnore.TabIndex = 1
        Me.chkIgnore.Text = "Ignore choice for now"
        Me.chkIgnore.UseVisualStyleBackColor = True
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOK.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.cmdOK.FlatAppearance.BorderSize = 0
        Me.cmdOK.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdOK.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.Blue
        Me.cmdOK.Location = New System.Drawing.Point(142, 276)
        Me.cmdOK.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(79, 33)
        Me.cmdOK.TabIndex = 126
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.cmdCancel.FlatAppearance.BorderSize = 0
        Me.cmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCancel.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(285, 276)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(79, 33)
        Me.cmdCancel.TabIndex = 127
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'lblH
        '
        Me.lblH.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblH.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblH.Location = New System.Drawing.Point(24, 21)
        Me.lblH.Name = "lblH"
        Me.lblH.Size = New System.Drawing.Size(550, 106)
        Me.lblH.TabIndex = 128
        Me.lblH.Text = "This table requires additional assignments for the sample set." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Please make a s" & _
    "election among the options below." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "If you wish to make the assignments later, " & _
    "check the 'Ignore' checkbox."
        Me.lblH.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'gbULots
        '
        Me.gbULots.Controls.Add(Me.rbLot10)
        Me.gbULots.Controls.Add(Me.rbLot9)
        Me.gbULots.Controls.Add(Me.rbLot8)
        Me.gbULots.Controls.Add(Me.rbLot7)
        Me.gbULots.Controls.Add(Me.rbLot6)
        Me.gbULots.Controls.Add(Me.rbLot5)
        Me.gbULots.Controls.Add(Me.rbLot4)
        Me.gbULots.Controls.Add(Me.rbLot3)
        Me.gbULots.Controls.Add(Me.rbLot2)
        Me.gbULots.Controls.Add(Me.rbLot1)
        Me.gbULots.Location = New System.Drawing.Point(618, 21)
        Me.gbULots.Name = "gbULots"
        Me.gbULots.Size = New System.Drawing.Size(132, 303)
        Me.gbULots.TabIndex = 129
        Me.gbULots.TabStop = False
        Me.gbULots.Text = "Unique Lots"
        '
        'rbLot10
        '
        Me.rbLot10.AutoSize = True
        Me.rbLot10.Location = New System.Drawing.Point(23, 267)
        Me.rbLot10.Name = "rbLot10"
        Me.rbLot10.Size = New System.Drawing.Size(62, 21)
        Me.rbLot10.TabIndex = 9
        Me.rbLot10.TabStop = True
        Me.rbLot10.Text = "Lot 10"
        Me.rbLot10.UseVisualStyleBackColor = True
        '
        'rbLot9
        '
        Me.rbLot9.AutoSize = True
        Me.rbLot9.Location = New System.Drawing.Point(23, 240)
        Me.rbLot9.Name = "rbLot9"
        Me.rbLot9.Size = New System.Drawing.Size(55, 21)
        Me.rbLot9.TabIndex = 8
        Me.rbLot9.TabStop = True
        Me.rbLot9.Text = "Lot 9"
        Me.rbLot9.UseVisualStyleBackColor = True
        '
        'rbLot8
        '
        Me.rbLot8.AutoSize = True
        Me.rbLot8.Location = New System.Drawing.Point(23, 213)
        Me.rbLot8.Name = "rbLot8"
        Me.rbLot8.Size = New System.Drawing.Size(55, 21)
        Me.rbLot8.TabIndex = 7
        Me.rbLot8.TabStop = True
        Me.rbLot8.Text = "Lot 8"
        Me.rbLot8.UseVisualStyleBackColor = True
        '
        'rbLot7
        '
        Me.rbLot7.AutoSize = True
        Me.rbLot7.Location = New System.Drawing.Point(23, 186)
        Me.rbLot7.Name = "rbLot7"
        Me.rbLot7.Size = New System.Drawing.Size(55, 21)
        Me.rbLot7.TabIndex = 6
        Me.rbLot7.TabStop = True
        Me.rbLot7.Text = "Lot 7"
        Me.rbLot7.UseVisualStyleBackColor = True
        '
        'rbLot6
        '
        Me.rbLot6.AutoSize = True
        Me.rbLot6.Location = New System.Drawing.Point(23, 159)
        Me.rbLot6.Name = "rbLot6"
        Me.rbLot6.Size = New System.Drawing.Size(55, 21)
        Me.rbLot6.TabIndex = 5
        Me.rbLot6.TabStop = True
        Me.rbLot6.Text = "Lot 6"
        Me.rbLot6.UseVisualStyleBackColor = True
        '
        'rbLot5
        '
        Me.rbLot5.AutoSize = True
        Me.rbLot5.Location = New System.Drawing.Point(23, 132)
        Me.rbLot5.Name = "rbLot5"
        Me.rbLot5.Size = New System.Drawing.Size(55, 21)
        Me.rbLot5.TabIndex = 4
        Me.rbLot5.TabStop = True
        Me.rbLot5.Text = "Lot 5"
        Me.rbLot5.UseVisualStyleBackColor = True
        '
        'rbLot4
        '
        Me.rbLot4.AutoSize = True
        Me.rbLot4.Location = New System.Drawing.Point(23, 105)
        Me.rbLot4.Name = "rbLot4"
        Me.rbLot4.Size = New System.Drawing.Size(55, 21)
        Me.rbLot4.TabIndex = 3
        Me.rbLot4.TabStop = True
        Me.rbLot4.Text = "Lot 4"
        Me.rbLot4.UseVisualStyleBackColor = True
        '
        'rbLot3
        '
        Me.rbLot3.AutoSize = True
        Me.rbLot3.Location = New System.Drawing.Point(23, 78)
        Me.rbLot3.Name = "rbLot3"
        Me.rbLot3.Size = New System.Drawing.Size(55, 21)
        Me.rbLot3.TabIndex = 2
        Me.rbLot3.TabStop = True
        Me.rbLot3.Text = "Lot 3"
        Me.rbLot3.UseVisualStyleBackColor = True
        '
        'rbLot2
        '
        Me.rbLot2.AutoSize = True
        Me.rbLot2.Location = New System.Drawing.Point(23, 51)
        Me.rbLot2.Name = "rbLot2"
        Me.rbLot2.Size = New System.Drawing.Size(55, 21)
        Me.rbLot2.TabIndex = 1
        Me.rbLot2.TabStop = True
        Me.rbLot2.Text = "Lot 2"
        Me.rbLot2.UseVisualStyleBackColor = True
        '
        'rbLot1
        '
        Me.rbLot1.AutoSize = True
        Me.rbLot1.Location = New System.Drawing.Point(23, 24)
        Me.rbLot1.Name = "rbLot1"
        Me.rbLot1.Size = New System.Drawing.Size(55, 21)
        Me.rbLot1.TabIndex = 0
        Me.rbLot1.TabStop = True
        Me.rbLot1.Text = "Lot 1"
        Me.rbLot1.UseVisualStyleBackColor = True
        '
        'gbStockSoln
        '
        Me.gbStockSoln.Controls.Add(Me.rbNewStockSoln)
        Me.gbStockSoln.Controls.Add(Me.rbOldStockSoln)
        Me.gbStockSoln.Location = New System.Drawing.Point(783, 21)
        Me.gbStockSoln.Name = "gbStockSoln"
        Me.gbStockSoln.Size = New System.Drawing.Size(275, 83)
        Me.gbStockSoln.TabIndex = 130
        Me.gbStockSoln.TabStop = False
        Me.gbStockSoln.Text = "Stock Solution Stability"
        '
        'rbNewStockSoln
        '
        Me.rbNewStockSoln.AutoSize = True
        Me.rbNewStockSoln.Location = New System.Drawing.Point(23, 51)
        Me.rbNewStockSoln.Name = "rbNewStockSoln"
        Me.rbNewStockSoln.Size = New System.Drawing.Size(138, 21)
        Me.rbNewStockSoln.TabIndex = 2
        Me.rbNewStockSoln.TabStop = True
        Me.rbNewStockSoln.Text = "New Stock Solution"
        Me.rbNewStockSoln.UseVisualStyleBackColor = True
        '
        'rbOldStockSoln
        '
        Me.rbOldStockSoln.AutoSize = True
        Me.rbOldStockSoln.Location = New System.Drawing.Point(23, 24)
        Me.rbOldStockSoln.Name = "rbOldStockSoln"
        Me.rbOldStockSoln.Size = New System.Drawing.Size(200, 21)
        Me.rbOldStockSoln.TabIndex = 0
        Me.rbOldStockSoln.TabStop = True
        Me.rbOldStockSoln.Text = "Old or Original Stock Solution"
        Me.rbOldStockSoln.UseVisualStyleBackColor = True
        '
        'gbSpikingSoln
        '
        Me.gbSpikingSoln.Controls.Add(Me.rbNewSpikingSoln)
        Me.gbSpikingSoln.Controls.Add(Me.rbOldSpikingSoln)
        Me.gbSpikingSoln.Location = New System.Drawing.Point(783, 118)
        Me.gbSpikingSoln.Name = "gbSpikingSoln"
        Me.gbSpikingSoln.Size = New System.Drawing.Size(275, 83)
        Me.gbSpikingSoln.TabIndex = 131
        Me.gbSpikingSoln.TabStop = False
        Me.gbSpikingSoln.Text = "Spiking Solution Stability"
        '
        'rbNewSpikingSoln
        '
        Me.rbNewSpikingSoln.AutoSize = True
        Me.rbNewSpikingSoln.Location = New System.Drawing.Point(23, 51)
        Me.rbNewSpikingSoln.Name = "rbNewSpikingSoln"
        Me.rbNewSpikingSoln.Size = New System.Drawing.Size(149, 21)
        Me.rbNewSpikingSoln.TabIndex = 2
        Me.rbNewSpikingSoln.TabStop = True
        Me.rbNewSpikingSoln.Text = "New Spiking Solution"
        Me.rbNewSpikingSoln.UseVisualStyleBackColor = True
        '
        'rbOldSpikingSoln
        '
        Me.rbOldSpikingSoln.AutoSize = True
        Me.rbOldSpikingSoln.Location = New System.Drawing.Point(23, 24)
        Me.rbOldSpikingSoln.Name = "rbOldSpikingSoln"
        Me.rbOldSpikingSoln.Size = New System.Drawing.Size(211, 21)
        Me.rbOldSpikingSoln.TabIndex = 0
        Me.rbOldSpikingSoln.TabStop = True
        Me.rbOldSpikingSoln.Text = "Old or Original Spiking Solution"
        Me.rbOldSpikingSoln.UseVisualStyleBackColor = True
        '
        'gbCarryOver
        '
        Me.gbCarryOver.Controls.Add(Me.rbBlank)
        Me.gbCarryOver.Controls.Add(Me.rbULOQ)
        Me.gbCarryOver.Controls.Add(Me.rbLLOQ)
        Me.gbCarryOver.Location = New System.Drawing.Point(783, 216)
        Me.gbCarryOver.Name = "gbCarryOver"
        Me.gbCarryOver.Size = New System.Drawing.Size(275, 108)
        Me.gbCarryOver.TabIndex = 132
        Me.gbCarryOver.TabStop = False
        Me.gbCarryOver.Text = "Carry Over"
        '
        'rbBlank
        '
        Me.rbBlank.AutoSize = True
        Me.rbBlank.Location = New System.Drawing.Point(23, 78)
        Me.rbBlank.Name = "rbBlank"
        Me.rbBlank.Size = New System.Drawing.Size(56, 21)
        Me.rbBlank.TabIndex = 3
        Me.rbBlank.TabStop = True
        Me.rbBlank.Text = "Blank"
        Me.rbBlank.UseVisualStyleBackColor = True
        '
        'rbULOQ
        '
        Me.rbULOQ.AutoSize = True
        Me.rbULOQ.Location = New System.Drawing.Point(23, 51)
        Me.rbULOQ.Name = "rbULOQ"
        Me.rbULOQ.Size = New System.Drawing.Size(61, 21)
        Me.rbULOQ.TabIndex = 2
        Me.rbULOQ.TabStop = True
        Me.rbULOQ.Text = "ULOQ"
        Me.rbULOQ.UseVisualStyleBackColor = True
        '
        'rbLLOQ
        '
        Me.rbLLOQ.AutoSize = True
        Me.rbLLOQ.Location = New System.Drawing.Point(23, 24)
        Me.rbLLOQ.Name = "rbLLOQ"
        Me.rbLLOQ.Size = New System.Drawing.Size(58, 21)
        Me.rbLLOQ.TabIndex = 0
        Me.rbLLOQ.TabStop = True
        Me.rbLLOQ.Text = "LLOQ"
        Me.rbLLOQ.UseVisualStyleBackColor = True
        '
        'gbLongTermQC
        '
        Me.gbLongTermQC.Controls.Add(Me.rbLTQCLTA)
        Me.gbLongTermQC.Controls.Add(Me.rbLTQCT0)
        Me.gbLongTermQC.Location = New System.Drawing.Point(783, 340)
        Me.gbLongTermQC.Name = "gbLongTermQC"
        Me.gbLongTermQC.Size = New System.Drawing.Size(275, 83)
        Me.gbLongTermQC.TabIndex = 133
        Me.gbLongTermQC.TabStop = False
        Me.gbLongTermQC.Text = "Long Term QC Storage Stability"
        '
        'rbLTQCLTA
        '
        Me.rbLTQCLTA.AutoSize = True
        Me.rbLTQCLTA.Location = New System.Drawing.Point(23, 51)
        Me.rbLTQCLTA.Name = "rbLTQCLTA"
        Me.rbLTQCLTA.Size = New System.Drawing.Size(139, 21)
        Me.rbLTQCLTA.TabIndex = 2
        Me.rbLTQCLTA.TabStop = True
        Me.rbLTQCLTA.Text = "Long Term Analysis"
        Me.rbLTQCLTA.UseVisualStyleBackColor = True
        '
        'rbLTQCT0
        '
        Me.rbLTQCT0.AutoSize = True
        Me.rbLTQCT0.Location = New System.Drawing.Point(23, 24)
        Me.rbLTQCT0.Name = "rbLTQCT0"
        Me.rbLTQCT0.Size = New System.Drawing.Size(98, 21)
        Me.rbLTQCT0.TabIndex = 0
        Me.rbLTQCT0.TabStop = True
        Me.rbLTQCT0.Text = "T(0) Analysis"
        Me.rbLTQCT0.UseVisualStyleBackColor = True
        '
        'gbAdHocStabComp
        '
        Me.gbAdHocStabComp.Controls.Add(Me.txtAdHocStabComp)
        Me.gbAdHocStabComp.Controls.Add(Me.lblAdHocStabComp)
        Me.gbAdHocStabComp.Location = New System.Drawing.Point(421, 364)
        Me.gbAdHocStabComp.Name = "gbAdHocStabComp"
        Me.gbAdHocStabComp.Size = New System.Drawing.Size(275, 124)
        Me.gbAdHocStabComp.TabIndex = 0
        Me.gbAdHocStabComp.TabStop = False
        Me.gbAdHocStabComp.Text = "Ad Hoc Stability Comparison"
        '
        'txtAdHocStabComp
        '
        Me.txtAdHocStabComp.Location = New System.Drawing.Point(22, 89)
        Me.txtAdHocStabComp.Name = "txtAdHocStabComp"
        Me.txtAdHocStabComp.Size = New System.Drawing.Size(231, 25)
        Me.txtAdHocStabComp.TabIndex = 0
        '
        'lblAdHocStabComp
        '
        Me.lblAdHocStabComp.Location = New System.Drawing.Point(6, 27)
        Me.lblAdHocStabComp.Name = "lblAdHocStabComp"
        Me.lblAdHocStabComp.Size = New System.Drawing.Size(263, 56)
        Me.lblAdHocStabComp.TabIndex = 0
        Me.lblAdHocStabComp.Text = "This table requires two sets of data that need to be labeled." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Enter a label for " & _
    "this set of data."
        Me.lblAdHocStabComp.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'frmAddRowsChoice
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1110, 521)
        Me.ControlBox = False
        Me.Controls.Add(Me.gbAdHocStabComp)
        Me.Controls.Add(Me.gbLongTermQC)
        Me.Controls.Add(Me.gbCarryOver)
        Me.Controls.Add(Me.gbSpikingSoln)
        Me.Controls.Add(Me.gbStockSoln)
        Me.Controls.Add(Me.gbULots)
        Me.Controls.Add(Me.lblH)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.chkIgnore)
        Me.Controls.Add(Me.gbRecovery)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "frmAddRowsChoice"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Choose Additional Sample Assignments..."
        Me.gbRecovery.ResumeLayout(False)
        Me.gbRecovery.PerformLayout()
        Me.gbULots.ResumeLayout(False)
        Me.gbULots.PerformLayout()
        Me.gbStockSoln.ResumeLayout(False)
        Me.gbStockSoln.PerformLayout()
        Me.gbSpikingSoln.ResumeLayout(False)
        Me.gbSpikingSoln.PerformLayout()
        Me.gbCarryOver.ResumeLayout(False)
        Me.gbCarryOver.PerformLayout()
        Me.gbLongTermQC.ResumeLayout(False)
        Me.gbLongTermQC.PerformLayout()
        Me.gbAdHocStabComp.ResumeLayout(False)
        Me.gbAdHocStabComp.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gbRecovery As System.Windows.Forms.GroupBox
    Friend WithEvents rbQC As System.Windows.Forms.RadioButton
    Friend WithEvents rbRS As System.Windows.Forms.RadioButton
    Friend WithEvents rbPES As System.Windows.Forms.RadioButton
    Friend WithEvents chkIgnore As System.Windows.Forms.CheckBox
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents lblH As System.Windows.Forms.Label
    Friend WithEvents gbULots As System.Windows.Forms.GroupBox
    Friend WithEvents rbLot1 As System.Windows.Forms.RadioButton
    Friend WithEvents rbLot10 As System.Windows.Forms.RadioButton
    Friend WithEvents rbLot9 As System.Windows.Forms.RadioButton
    Friend WithEvents rbLot8 As System.Windows.Forms.RadioButton
    Friend WithEvents rbLot7 As System.Windows.Forms.RadioButton
    Friend WithEvents rbLot6 As System.Windows.Forms.RadioButton
    Friend WithEvents rbLot5 As System.Windows.Forms.RadioButton
    Friend WithEvents rbLot4 As System.Windows.Forms.RadioButton
    Friend WithEvents rbLot3 As System.Windows.Forms.RadioButton
    Friend WithEvents rbLot2 As System.Windows.Forms.RadioButton
    Friend WithEvents gbStockSoln As System.Windows.Forms.GroupBox
    Friend WithEvents rbNewStockSoln As System.Windows.Forms.RadioButton
    Friend WithEvents rbOldStockSoln As System.Windows.Forms.RadioButton
    Friend WithEvents gbSpikingSoln As System.Windows.Forms.GroupBox
    Friend WithEvents rbNewSpikingSoln As System.Windows.Forms.RadioButton
    Friend WithEvents rbOldSpikingSoln As System.Windows.Forms.RadioButton
    Friend WithEvents gbCarryOver As System.Windows.Forms.GroupBox
    Friend WithEvents rbBlank As System.Windows.Forms.RadioButton
    Friend WithEvents rbULOQ As System.Windows.Forms.RadioButton
    Friend WithEvents rbLLOQ As System.Windows.Forms.RadioButton
    Friend WithEvents gbLongTermQC As System.Windows.Forms.GroupBox
    Friend WithEvents rbLTQCLTA As System.Windows.Forms.RadioButton
    Friend WithEvents rbLTQCT0 As System.Windows.Forms.RadioButton
    Friend WithEvents gbAdHocStabComp As System.Windows.Forms.GroupBox
    Friend WithEvents txtAdHocStabComp As System.Windows.Forms.TextBox
    Friend WithEvents lblAdHocStabComp As System.Windows.Forms.Label
End Class
