<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConfigCompliance
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmConfigCompliance))
        Me.panCompliance = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblReasonForChange = New System.Windows.Forms.Label()
        Me.lblMeaningOfSig = New System.Windows.Forms.Label()
        Me.dgvRFC = New System.Windows.Forms.DataGridView()
        Me.dgvMOS = New System.Windows.Forms.DataGridView()
        Me.gbAuditTrail = New System.Windows.Forms.GroupBox()
        Me.rbAuditTrailOff = New System.Windows.Forms.RadioButton()
        Me.gbESig = New System.Windows.Forms.GroupBox()
        Me.panESigOptions = New System.Windows.Forms.Panel()
        Me.chkReasonFreeForm = New System.Windows.Forms.CheckBox()
        Me.chkSigFreeForm = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.gbUserIDType = New System.Windows.Forms.GroupBox()
        Me.rbUserIDChoice = New System.Windows.Forms.RadioButton()
        Me.rbOnlyLoggedOn = New System.Windows.Forms.RadioButton()
        Me.chkReasonForChange = New System.Windows.Forms.CheckBox()
        Me.chkMeaningOfSign = New System.Windows.Forms.CheckBox()
        Me.rbESigOff = New System.Windows.Forms.RadioButton()
        Me.rbESigOn = New System.Windows.Forms.RadioButton()
        Me.rbAuditTrailOn = New System.Windows.Forms.RadioButton()
        Me.panCompliance.SuspendLayout()
        CType(Me.dgvRFC, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvMOS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbAuditTrail.SuspendLayout()
        Me.gbESig.SuspendLayout()
        Me.panESigOptions.SuspendLayout()
        Me.gbUserIDType.SuspendLayout()
        Me.SuspendLayout()
        '
        'panCompliance
        '
        Me.panCompliance.Controls.Add(Me.Label2)
        Me.panCompliance.Controls.Add(Me.lblReasonForChange)
        Me.panCompliance.Controls.Add(Me.lblMeaningOfSig)
        Me.panCompliance.Controls.Add(Me.dgvRFC)
        Me.panCompliance.Controls.Add(Me.dgvMOS)
        Me.panCompliance.Controls.Add(Me.gbAuditTrail)
        Me.panCompliance.Location = New System.Drawing.Point(12, 12)
        Me.panCompliance.Name = "panCompliance"
        Me.panCompliance.Size = New System.Drawing.Size(424, 422)
        Me.panCompliance.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Blue
        Me.Label2.Location = New System.Drawing.Point(4, 5)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(209, 19)
        Me.Label2.TabIndex = 136
        Me.Label2.Text = "Compliance Configuration"
        '
        'lblReasonForChange
        '
        Me.lblReasonForChange.AutoSize = True
        Me.lblReasonForChange.Location = New System.Drawing.Point(308, 254)
        Me.lblReasonForChange.Name = "lblReasonForChange"
        Me.lblReasonForChange.Size = New System.Drawing.Size(102, 13)
        Me.lblReasonForChange.TabIndex = 5
        Me.lblReasonForChange.Text = "Reason For Change"
        '
        'lblMeaningOfSig
        '
        Me.lblMeaningOfSig.AutoSize = True
        Me.lblMeaningOfSig.Location = New System.Drawing.Point(308, 23)
        Me.lblMeaningOfSig.Name = "lblMeaningOfSig"
        Me.lblMeaningOfSig.Size = New System.Drawing.Size(108, 13)
        Me.lblMeaningOfSig.TabIndex = 4
        Me.lblMeaningOfSig.Text = "Meaning of Signature"
        '
        'dgvRFC
        '
        Me.dgvRFC.AllowUserToAddRows = False
        Me.dgvRFC.AllowUserToDeleteRows = False
        Me.dgvRFC.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvRFC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvRFC.Location = New System.Drawing.Point(308, 270)
        Me.dgvRFC.Name = "dgvRFC"
        Me.dgvRFC.Size = New System.Drawing.Size(190, 209)
        Me.dgvRFC.TabIndex = 3
        '
        'dgvMOS
        '
        Me.dgvMOS.AllowUserToAddRows = False
        Me.dgvMOS.AllowUserToDeleteRows = False
        Me.dgvMOS.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvMOS.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvMOS.Location = New System.Drawing.Point(308, 39)
        Me.dgvMOS.Name = "dgvMOS"
        Me.dgvMOS.Size = New System.Drawing.Size(190, 209)
        Me.dgvMOS.TabIndex = 2
        '
        'gbAuditTrail
        '
        Me.gbAuditTrail.Controls.Add(Me.rbAuditTrailOff)
        Me.gbAuditTrail.Controls.Add(Me.gbESig)
        Me.gbAuditTrail.Controls.Add(Me.rbAuditTrailOn)
        Me.gbAuditTrail.Location = New System.Drawing.Point(15, 39)
        Me.gbAuditTrail.Name = "gbAuditTrail"
        Me.gbAuditTrail.Size = New System.Drawing.Size(287, 401)
        Me.gbAuditTrail.TabIndex = 1
        Me.gbAuditTrail.TabStop = False
        Me.gbAuditTrail.Text = "Audit Trail Options"
        '
        'rbAuditTrailOff
        '
        Me.rbAuditTrailOff.AutoSize = True
        Me.rbAuditTrailOff.Checked = True
        Me.rbAuditTrailOff.Location = New System.Drawing.Point(89, 19)
        Me.rbAuditTrailOff.Name = "rbAuditTrailOff"
        Me.rbAuditTrailOff.Size = New System.Drawing.Size(39, 17)
        Me.rbAuditTrailOff.TabIndex = 2
        Me.rbAuditTrailOff.TabStop = True
        Me.rbAuditTrailOff.Text = "Off"
        Me.rbAuditTrailOff.UseVisualStyleBackColor = True
        '
        'gbESig
        '
        Me.gbESig.Controls.Add(Me.panESigOptions)
        Me.gbESig.Controls.Add(Me.rbESigOff)
        Me.gbESig.Controls.Add(Me.rbESigOn)
        Me.gbESig.Location = New System.Drawing.Point(17, 55)
        Me.gbESig.Name = "gbESig"
        Me.gbESig.Size = New System.Drawing.Size(244, 310)
        Me.gbESig.TabIndex = 0
        Me.gbESig.TabStop = False
        Me.gbESig.Text = "Electronic Signature Options"
        '
        'panESigOptions
        '
        Me.panESigOptions.Controls.Add(Me.chkReasonFreeForm)
        Me.panESigOptions.Controls.Add(Me.chkSigFreeForm)
        Me.panESigOptions.Controls.Add(Me.Label1)
        Me.panESigOptions.Controls.Add(Me.gbUserIDType)
        Me.panESigOptions.Controls.Add(Me.chkReasonForChange)
        Me.panESigOptions.Controls.Add(Me.chkMeaningOfSign)
        Me.panESigOptions.Location = New System.Drawing.Point(15, 51)
        Me.panESigOptions.Name = "panESigOptions"
        Me.panESigOptions.Size = New System.Drawing.Size(203, 223)
        Me.panESigOptions.TabIndex = 2
        '
        'chkReasonFreeForm
        '
        Me.chkReasonFreeForm.AutoSize = True
        Me.chkReasonFreeForm.Location = New System.Drawing.Point(34, 192)
        Me.chkReasonFreeForm.Name = "chkReasonFreeForm"
        Me.chkReasonFreeForm.Size = New System.Drawing.Size(164, 17)
        Me.chkReasonFreeForm.TabIndex = 4
        Me.chkReasonFreeForm.Text = "Restrict to dropdown choices"
        Me.chkReasonFreeForm.UseVisualStyleBackColor = True
        '
        'chkSigFreeForm
        '
        Me.chkSigFreeForm.AutoSize = True
        Me.chkSigFreeForm.Location = New System.Drawing.Point(34, 136)
        Me.chkSigFreeForm.Name = "chkSigFreeForm"
        Me.chkSigFreeForm.Size = New System.Drawing.Size(164, 17)
        Me.chkSigFreeForm.TabIndex = 3
        Me.chkSigFreeForm.Text = "Restrict to dropdown choices"
        Me.chkSigFreeForm.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 92)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(107, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Include the following:"
        '
        'gbUserIDType
        '
        Me.gbUserIDType.Controls.Add(Me.rbUserIDChoice)
        Me.gbUserIDType.Controls.Add(Me.rbOnlyLoggedOn)
        Me.gbUserIDType.Location = New System.Drawing.Point(15, 12)
        Me.gbUserIDType.Name = "gbUserIDType"
        Me.gbUserIDType.Size = New System.Drawing.Size(139, 68)
        Me.gbUserIDType.TabIndex = 2
        Me.gbUserIDType.TabStop = False
        Me.gbUserIDType.Text = "User ID Option"
        '
        'rbUserIDChoice
        '
        Me.rbUserIDChoice.AutoSize = True
        Me.rbUserIDChoice.Location = New System.Drawing.Point(6, 42)
        Me.rbUserIDChoice.Name = "rbUserIDChoice"
        Me.rbUserIDChoice.Size = New System.Drawing.Size(128, 17)
        Me.rbUserIDChoice.TabIndex = 2
        Me.rbUserIDChoice.Text = "Allow Choice of Users"
        Me.rbUserIDChoice.UseVisualStyleBackColor = True
        '
        'rbOnlyLoggedOn
        '
        Me.rbOnlyLoggedOn.AutoSize = True
        Me.rbOnlyLoggedOn.Checked = True
        Me.rbOnlyLoggedOn.Location = New System.Drawing.Point(6, 19)
        Me.rbOnlyLoggedOn.Name = "rbOnlyLoggedOn"
        Me.rbOnlyLoggedOn.Size = New System.Drawing.Size(127, 17)
        Me.rbOnlyLoggedOn.TabIndex = 1
        Me.rbOnlyLoggedOn.TabStop = True
        Me.rbOnlyLoggedOn.Text = "Only Logged On User"
        Me.rbOnlyLoggedOn.UseVisualStyleBackColor = True
        '
        'chkReasonForChange
        '
        Me.chkReasonForChange.AutoSize = True
        Me.chkReasonForChange.Location = New System.Drawing.Point(15, 169)
        Me.chkReasonForChange.Name = "chkReasonForChange"
        Me.chkReasonForChange.Size = New System.Drawing.Size(121, 17)
        Me.chkReasonForChange.TabIndex = 2
        Me.chkReasonForChange.Text = "Reason For Change"
        Me.chkReasonForChange.UseVisualStyleBackColor = True
        '
        'chkMeaningOfSign
        '
        Me.chkMeaningOfSign.AutoSize = True
        Me.chkMeaningOfSign.Location = New System.Drawing.Point(15, 113)
        Me.chkMeaningOfSign.Name = "chkMeaningOfSign"
        Me.chkMeaningOfSign.Size = New System.Drawing.Size(129, 17)
        Me.chkMeaningOfSign.TabIndex = 1
        Me.chkMeaningOfSign.Text = "Meaning Of Signature"
        Me.chkMeaningOfSign.UseVisualStyleBackColor = True
        '
        'rbESigOff
        '
        Me.rbESigOff.AutoSize = True
        Me.rbESigOff.Checked = True
        Me.rbESigOff.Location = New System.Drawing.Point(72, 28)
        Me.rbESigOff.Name = "rbESigOff"
        Me.rbESigOff.Size = New System.Drawing.Size(39, 17)
        Me.rbESigOff.TabIndex = 1
        Me.rbESigOff.TabStop = True
        Me.rbESigOff.Text = "Off"
        Me.rbESigOff.UseVisualStyleBackColor = True
        '
        'rbESigOn
        '
        Me.rbESigOn.AutoSize = True
        Me.rbESigOn.Location = New System.Drawing.Point(15, 28)
        Me.rbESigOn.Name = "rbESigOn"
        Me.rbESigOn.Size = New System.Drawing.Size(39, 17)
        Me.rbESigOn.TabIndex = 0
        Me.rbESigOn.Text = "On"
        Me.rbESigOn.UseVisualStyleBackColor = True
        '
        'rbAuditTrailOn
        '
        Me.rbAuditTrailOn.AutoSize = True
        Me.rbAuditTrailOn.Location = New System.Drawing.Point(17, 19)
        Me.rbAuditTrailOn.Name = "rbAuditTrailOn"
        Me.rbAuditTrailOn.Size = New System.Drawing.Size(39, 17)
        Me.rbAuditTrailOn.TabIndex = 1
        Me.rbAuditTrailOn.Text = "On"
        Me.rbAuditTrailOn.UseVisualStyleBackColor = True
        '
        'frmConfigCompliance
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(758, 508)
        Me.Controls.Add(Me.panCompliance)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmConfigCompliance"
        Me.ShowInTaskbar = False
        Me.Text = "Compliance Configuration"
        Me.panCompliance.ResumeLayout(False)
        Me.panCompliance.PerformLayout()
        CType(Me.dgvRFC, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvMOS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbAuditTrail.ResumeLayout(False)
        Me.gbAuditTrail.PerformLayout()
        Me.gbESig.ResumeLayout(False)
        Me.gbESig.PerformLayout()
        Me.panESigOptions.ResumeLayout(False)
        Me.panESigOptions.PerformLayout()
        Me.gbUserIDType.ResumeLayout(False)
        Me.gbUserIDType.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents panCompliance As System.Windows.Forms.Panel
    Friend WithEvents gbAuditTrail As System.Windows.Forms.GroupBox
    Friend WithEvents gbESig As System.Windows.Forms.GroupBox
    Friend WithEvents rbESigOn As System.Windows.Forms.RadioButton
    Friend WithEvents rbAuditTrailOff As System.Windows.Forms.RadioButton
    Friend WithEvents rbAuditTrailOn As System.Windows.Forms.RadioButton
    Friend WithEvents rbESigOff As System.Windows.Forms.RadioButton
    Friend WithEvents panESigOptions As System.Windows.Forms.Panel
    Friend WithEvents chkReasonForChange As System.Windows.Forms.CheckBox
    Friend WithEvents chkMeaningOfSign As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents gbUserIDType As System.Windows.Forms.GroupBox
    Friend WithEvents rbUserIDChoice As System.Windows.Forms.RadioButton
    Friend WithEvents rbOnlyLoggedOn As System.Windows.Forms.RadioButton
    Friend WithEvents dgvRFC As System.Windows.Forms.DataGridView
    Friend WithEvents dgvMOS As System.Windows.Forms.DataGridView
    Friend WithEvents chkSigFreeForm As System.Windows.Forms.CheckBox
    Friend WithEvents chkReasonFreeForm As System.Windows.Forms.CheckBox
    Friend WithEvents lblReasonForChange As System.Windows.Forms.Label
    Friend WithEvents lblMeaningOfSig As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
