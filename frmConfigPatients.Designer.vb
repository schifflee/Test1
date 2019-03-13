<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConfigPatients
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmConfigPatients))
        Me.dgvGroupSummary = New System.Windows.Forms.DataGridView()
        Me.dgvGroupTimePoints = New System.Windows.Forms.DataGridView()
        Me.dgvPatients = New System.Windows.Forms.DataGridView()
        Me.gbxBleeds = New System.Windows.Forms.GroupBox()
        Me.rbSerialNon = New System.Windows.Forms.RadioButton()
        Me.rbSerial = New System.Windows.Forms.RadioButton()
        Me.lblGroupSummary = New System.Windows.Forms.Label()
        Me.lblPatients = New System.Windows.Forms.Label()
        Me.lblTimePoints = New System.Windows.Forms.Label()
        Me.txtBaseName = New System.Windows.Forms.TextBox()
        Me.gbxNameType = New System.Windows.Forms.GroupBox()
        Me.rbBothNames = New System.Windows.Forms.RadioButton()
        Me.rbUnique = New System.Windows.Forms.RadioButton()
        Me.rbIncrement = New System.Windows.Forms.RadioButton()
        Me.lblBaseName = New System.Windows.Forms.Label()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.cmdRemove = New System.Windows.Forms.Button()
        Me.panTimePoints = New System.Windows.Forms.Panel()
        Me.cmdAddTP = New System.Windows.Forms.Button()
        Me.pan1 = New System.Windows.Forms.Panel()
        Me.cmdCancel1 = New System.Windows.Forms.Button()
        Me.cmdOK1 = New System.Windows.Forms.Button()
        Me.lblUnique = New System.Windows.Forms.Label()
        Me.gbxIncrement = New System.Windows.Forms.GroupBox()
        Me.rbGroup = New System.Windows.Forms.RadioButton()
        Me.rbEntire = New System.Windows.Forms.RadioButton()
        Me.lbltxtUniqueID = New System.Windows.Forms.Label()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.cmdEdit = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        CType(Me.dgvGroupSummary, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvGroupTimePoints, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvPatients, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbxBleeds.SuspendLayout()
        Me.gbxNameType.SuspendLayout()
        Me.panTimePoints.SuspendLayout()
        Me.pan1.SuspendLayout()
        Me.gbxIncrement.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvGroupSummary
        '
        Me.dgvGroupSummary.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvGroupSummary.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvGroupSummary.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvGroupSummary.Location = New System.Drawing.Point(12, 64)
        Me.dgvGroupSummary.Name = "dgvGroupSummary"
        Me.dgvGroupSummary.Size = New System.Drawing.Size(100, 369)
        Me.dgvGroupSummary.TabIndex = 0
        '
        'dgvGroupTimePoints
        '
        Me.dgvGroupTimePoints.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvGroupTimePoints.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.dgvGroupTimePoints.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvGroupTimePoints.Location = New System.Drawing.Point(66, 16)
        Me.dgvGroupTimePoints.MultiSelect = False
        Me.dgvGroupTimePoints.Name = "dgvGroupTimePoints"
        Me.dgvGroupTimePoints.Size = New System.Drawing.Size(100, 369)
        Me.dgvGroupTimePoints.TabIndex = 1
        '
        'dgvPatients
        '
        Me.dgvPatients.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvPatients.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvPatients.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvPatients.Location = New System.Drawing.Point(308, 64)
        Me.dgvPatients.MultiSelect = False
        Me.dgvPatients.Name = "dgvPatients"
        Me.dgvPatients.Size = New System.Drawing.Size(346, 369)
        Me.dgvPatients.TabIndex = 2
        '
        'gbxBleeds
        '
        Me.gbxBleeds.Controls.Add(Me.rbSerialNon)
        Me.gbxBleeds.Controls.Add(Me.rbSerial)
        Me.gbxBleeds.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbxBleeds.Location = New System.Drawing.Point(135, 12)
        Me.gbxBleeds.Name = "gbxBleeds"
        Me.gbxBleeds.Size = New System.Drawing.Size(164, 38)
        Me.gbxBleeds.TabIndex = 177
        Me.gbxBleeds.TabStop = False
        Me.gbxBleeds.Text = "Bleed Type"
        '
        'rbSerialNon
        '
        Me.rbSerialNon.AutoSize = True
        Me.rbSerialNon.Location = New System.Drawing.Point(69, 16)
        Me.rbSerialNon.Name = "rbSerialNon"
        Me.rbSerialNon.Size = New System.Drawing.Size(74, 17)
        Me.rbSerialNon.TabIndex = 1
        Me.rbSerialNon.Text = "Non-Serial"
        Me.rbSerialNon.UseVisualStyleBackColor = True
        '
        'rbSerial
        '
        Me.rbSerial.AutoSize = True
        Me.rbSerial.Checked = True
        Me.rbSerial.Location = New System.Drawing.Point(12, 16)
        Me.rbSerial.Name = "rbSerial"
        Me.rbSerial.Size = New System.Drawing.Size(51, 17)
        Me.rbSerial.TabIndex = 0
        Me.rbSerial.TabStop = True
        Me.rbSerial.Text = "Serial"
        Me.rbSerial.UseVisualStyleBackColor = True
        '
        'lblGroupSummary
        '
        Me.lblGroupSummary.AutoSize = True
        Me.lblGroupSummary.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGroupSummary.Location = New System.Drawing.Point(9, 48)
        Me.lblGroupSummary.Name = "lblGroupSummary"
        Me.lblGroupSummary.Size = New System.Drawing.Size(95, 13)
        Me.lblGroupSummary.TabIndex = 178
        Me.lblGroupSummary.Text = "Group Summary"
        '
        'lblPatients
        '
        Me.lblPatients.AutoSize = True
        Me.lblPatients.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPatients.Location = New System.Drawing.Point(305, 48)
        Me.lblPatients.Name = "lblPatients"
        Me.lblPatients.Size = New System.Drawing.Size(53, 13)
        Me.lblPatients.TabIndex = 179
        Me.lblPatients.Text = "Patients"
        '
        'lblTimePoints
        '
        Me.lblTimePoints.AutoSize = True
        Me.lblTimePoints.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTimePoints.Location = New System.Drawing.Point(63, 0)
        Me.lblTimePoints.Name = "lblTimePoints"
        Me.lblTimePoints.Size = New System.Drawing.Size(73, 13)
        Me.lblTimePoints.TabIndex = 180
        Me.lblTimePoints.Text = "Time Points"
        '
        'txtBaseName
        '
        Me.txtBaseName.Location = New System.Drawing.Point(135, 197)
        Me.txtBaseName.Name = "txtBaseName"
        Me.txtBaseName.Size = New System.Drawing.Size(151, 20)
        Me.txtBaseName.TabIndex = 181
        '
        'gbxNameType
        '
        Me.gbxNameType.Controls.Add(Me.rbBothNames)
        Me.gbxNameType.Controls.Add(Me.rbUnique)
        Me.gbxNameType.Controls.Add(Me.rbIncrement)
        Me.gbxNameType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbxNameType.Location = New System.Drawing.Point(135, 56)
        Me.gbxNameType.Name = "gbxNameType"
        Me.gbxNameType.Size = New System.Drawing.Size(164, 108)
        Me.gbxNameType.TabIndex = 182
        Me.gbxNameType.TabStop = False
        Me.gbxNameType.Text = "Naming Convention"
        '
        'rbBothNames
        '
        Me.rbBothNames.Location = New System.Drawing.Point(12, 62)
        Me.rbBothNames.Name = "rbBothNames"
        Me.rbBothNames.Size = New System.Drawing.Size(142, 40)
        Me.rbBothNames.TabIndex = 2
        Me.rbBothNames.Text = "Incremented Base Name AND Unique ID"
        Me.rbBothNames.UseVisualStyleBackColor = True
        '
        'rbUnique
        '
        Me.rbUnique.AutoSize = True
        Me.rbUnique.Location = New System.Drawing.Point(12, 39)
        Me.rbUnique.Name = "rbUnique"
        Me.rbUnique.Size = New System.Drawing.Size(73, 17)
        Me.rbUnique.TabIndex = 1
        Me.rbUnique.Text = "Unique ID"
        Me.rbUnique.UseVisualStyleBackColor = True
        '
        'rbIncrement
        '
        Me.rbIncrement.AutoSize = True
        Me.rbIncrement.Checked = True
        Me.rbIncrement.Location = New System.Drawing.Point(12, 16)
        Me.rbIncrement.Name = "rbIncrement"
        Me.rbIncrement.Size = New System.Drawing.Size(142, 17)
        Me.rbIncrement.TabIndex = 0
        Me.rbIncrement.TabStop = True
        Me.rbIncrement.Text = "Incremented Base Name"
        Me.rbIncrement.UseVisualStyleBackColor = True
        '
        'lblBaseName
        '
        Me.lblBaseName.AutoSize = True
        Me.lblBaseName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBaseName.Location = New System.Drawing.Point(132, 181)
        Me.lblBaseName.Name = "lblBaseName"
        Me.lblBaseName.Size = New System.Drawing.Size(75, 13)
        Me.lblBaseName.TabIndex = 183
        Me.lblBaseName.Text = "Base Name:"
        '
        'cmdAdd
        '
        Me.cmdAdd.Enabled = False
        Me.cmdAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAdd.ForeColor = System.Drawing.Color.Blue
        Me.cmdAdd.Location = New System.Drawing.Point(207, 272)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(79, 26)
        Me.cmdAdd.TabIndex = 185
        Me.cmdAdd.Text = "Add ->"
        Me.cmdAdd.UseVisualStyleBackColor = True
        '
        'cmdRemove
        '
        Me.cmdRemove.Enabled = False
        Me.cmdRemove.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRemove.ForeColor = System.Drawing.Color.Red
        Me.cmdRemove.Location = New System.Drawing.Point(207, 304)
        Me.cmdRemove.Name = "cmdRemove"
        Me.cmdRemove.Size = New System.Drawing.Size(79, 26)
        Me.cmdRemove.TabIndex = 186
        Me.cmdRemove.Text = "<- Remove"
        Me.cmdRemove.UseVisualStyleBackColor = True
        '
        'panTimePoints
        '
        Me.panTimePoints.Controls.Add(Me.cmdAddTP)
        Me.panTimePoints.Controls.Add(Me.dgvGroupTimePoints)
        Me.panTimePoints.Controls.Add(Me.lblTimePoints)
        Me.panTimePoints.Location = New System.Drawing.Point(660, 48)
        Me.panTimePoints.Name = "panTimePoints"
        Me.panTimePoints.Size = New System.Drawing.Size(169, 394)
        Me.panTimePoints.TabIndex = 187
        Me.panTimePoints.Visible = False
        '
        'cmdAddTP
        '
        Me.cmdAddTP.Enabled = False
        Me.cmdAddTP.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddTP.ForeColor = System.Drawing.Color.Blue
        Me.cmdAddTP.Location = New System.Drawing.Point(3, 133)
        Me.cmdAddTP.Name = "cmdAddTP"
        Me.cmdAddTP.Size = New System.Drawing.Size(57, 26)
        Me.cmdAddTP.TabIndex = 186
        Me.cmdAddTP.Text = "<- Add"
        Me.cmdAddTP.UseVisualStyleBackColor = True
        '
        'pan1
        '
        Me.pan1.Controls.Add(Me.cmdCancel1)
        Me.pan1.Controls.Add(Me.cmdOK1)
        Me.pan1.Location = New System.Drawing.Point(135, 369)
        Me.pan1.Name = "pan1"
        Me.pan1.Size = New System.Drawing.Size(231, 45)
        Me.pan1.TabIndex = 188
        Me.pan1.Visible = False
        '
        'cmdCancel1
        '
        Me.cmdCancel1.CausesValidation = False
        Me.cmdCancel1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel1.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel1.Location = New System.Drawing.Point(89, 9)
        Me.cmdCancel1.Name = "cmdCancel1"
        Me.cmdCancel1.Size = New System.Drawing.Size(80, 35)
        Me.cmdCancel1.TabIndex = 1
        Me.cmdCancel1.Text = "&Cancel"
        Me.cmdCancel1.UseVisualStyleBackColor = True
        '
        'cmdOK1
        '
        Me.cmdOK1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK1.ForeColor = System.Drawing.Color.Blue
        Me.cmdOK1.Location = New System.Drawing.Point(0, 9)
        Me.cmdOK1.Name = "cmdOK1"
        Me.cmdOK1.Size = New System.Drawing.Size(80, 35)
        Me.cmdOK1.TabIndex = 0
        Me.cmdOK1.Text = "&OK"
        Me.cmdOK1.UseVisualStyleBackColor = True
        '
        'lblUnique
        '
        Me.lblUnique.AutoSize = True
        Me.lblUnique.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUnique.Location = New System.Drawing.Point(132, 222)
        Me.lblUnique.Name = "lblUnique"
        Me.lblUnique.Size = New System.Drawing.Size(68, 13)
        Me.lblUnique.TabIndex = 183
        Me.lblUnique.Text = "Unique ID:"
        '
        'gbxIncrement
        '
        Me.gbxIncrement.Controls.Add(Me.rbGroup)
        Me.gbxIncrement.Controls.Add(Me.rbEntire)
        Me.gbxIncrement.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbxIncrement.Location = New System.Drawing.Point(456, 341)
        Me.gbxIncrement.Name = "gbxIncrement"
        Me.gbxIncrement.Size = New System.Drawing.Size(164, 92)
        Me.gbxIncrement.TabIndex = 190
        Me.gbxIncrement.TabStop = False
        Me.gbxIncrement.Text = "Increment Base Name Style"
        Me.gbxIncrement.Visible = False
        '
        'rbGroup
        '
        Me.rbGroup.Location = New System.Drawing.Point(12, 53)
        Me.rbGroup.Name = "rbGroup"
        Me.rbGroup.Size = New System.Drawing.Size(143, 34)
        Me.rbGroup.TabIndex = 1
        Me.rbGroup.Text = "Increment Base Name within Group"
        Me.rbGroup.UseVisualStyleBackColor = True
        '
        'rbEntire
        '
        Me.rbEntire.Checked = True
        Me.rbEntire.Location = New System.Drawing.Point(12, 16)
        Me.rbEntire.Name = "rbEntire"
        Me.rbEntire.Size = New System.Drawing.Size(142, 31)
        Me.rbEntire.TabIndex = 0
        Me.rbEntire.TabStop = True
        Me.rbEntire.Text = "Increment Base Name through entire Assay"
        Me.rbEntire.UseVisualStyleBackColor = True
        '
        'lbltxtUniqueID
        '
        Me.lbltxtUniqueID.AutoSize = True
        Me.lbltxtUniqueID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbltxtUniqueID.ForeColor = System.Drawing.Color.Blue
        Me.lbltxtUniqueID.Location = New System.Drawing.Point(141, 241)
        Me.lbltxtUniqueID.Name = "lbltxtUniqueID"
        Me.lbltxtUniqueID.Size = New System.Drawing.Size(151, 13)
        Me.lbltxtUniqueID.TabIndex = 191
        Me.lbltxtUniqueID.Text = "Edit Unique ID in table ->"
        '
        'cmdSave
        '
        Me.cmdSave.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdSave.Enabled = False
        Me.cmdSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ForeColor = System.Drawing.Color.ForestGreen
        Me.cmdSave.Location = New System.Drawing.Point(625, 12)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(65, 25)
        Me.cmdSave.TabIndex = 195
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.UseVisualStyleBackColor = False
        '
        'cmdEdit
        '
        Me.cmdEdit.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdEdit.CausesValidation = False
        Me.cmdEdit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdEdit.ForeColor = System.Drawing.Color.Blue
        Me.cmdEdit.Location = New System.Drawing.Point(555, 12)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.Size = New System.Drawing.Size(65, 25)
        Me.cmdEdit.TabIndex = 193
        Me.cmdEdit.Text = "&Edit"
        Me.cmdEdit.UseVisualStyleBackColor = False
        '
        'cmdExit
        '
        Me.cmdExit.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdExit.CausesValidation = False
        Me.cmdExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ForeColor = System.Drawing.Color.Red
        Me.cmdExit.Location = New System.Drawing.Point(764, 12)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(65, 25)
        Me.cmdExit.TabIndex = 192
        Me.cmdExit.Text = "E&xit"
        Me.cmdExit.UseVisualStyleBackColor = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancel.CausesValidation = False
        Me.cmdCancel.Enabled = False
        Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(694, 12)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(65, 25)
        Me.cmdCancel.TabIndex = 194
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'frmConfigPatients
        '
        Me.AcceptButton = Me.cmdAdd
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(837, 444)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdEdit)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.lbltxtUniqueID)
        Me.Controls.Add(Me.lblUnique)
        Me.Controls.Add(Me.txtBaseName)
        Me.Controls.Add(Me.lblBaseName)
        Me.Controls.Add(Me.gbxIncrement)
        Me.Controls.Add(Me.pan1)
        Me.Controls.Add(Me.panTimePoints)
        Me.Controls.Add(Me.cmdRemove)
        Me.Controls.Add(Me.cmdAdd)
        Me.Controls.Add(Me.gbxNameType)
        Me.Controls.Add(Me.lblPatients)
        Me.Controls.Add(Me.lblGroupSummary)
        Me.Controls.Add(Me.gbxBleeds)
        Me.Controls.Add(Me.dgvPatients)
        Me.Controls.Add(Me.dgvGroupSummary)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmConfigPatients"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Configure Assay Patients"
        CType(Me.dgvGroupSummary, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvGroupTimePoints, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvPatients, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbxBleeds.ResumeLayout(False)
        Me.gbxBleeds.PerformLayout()
        Me.gbxNameType.ResumeLayout(False)
        Me.gbxNameType.PerformLayout()
        Me.panTimePoints.ResumeLayout(False)
        Me.panTimePoints.PerformLayout()
        Me.pan1.ResumeLayout(False)
        Me.gbxIncrement.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvGroupSummary As System.Windows.Forms.DataGridView
    Friend WithEvents dgvGroupTimePoints As System.Windows.Forms.DataGridView
    Friend WithEvents dgvPatients As System.Windows.Forms.DataGridView
    Friend WithEvents gbxBleeds As System.Windows.Forms.GroupBox
    Friend WithEvents rbSerialNon As System.Windows.Forms.RadioButton
    Friend WithEvents rbSerial As System.Windows.Forms.RadioButton
    Friend WithEvents lblGroupSummary As System.Windows.Forms.Label
    Friend WithEvents lblPatients As System.Windows.Forms.Label
    Friend WithEvents lblTimePoints As System.Windows.Forms.Label
    Friend WithEvents txtBaseName As System.Windows.Forms.TextBox
    Friend WithEvents gbxNameType As System.Windows.Forms.GroupBox
    Friend WithEvents rbUnique As System.Windows.Forms.RadioButton
    Friend WithEvents rbIncrement As System.Windows.Forms.RadioButton
    Friend WithEvents lblBaseName As System.Windows.Forms.Label
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents cmdRemove As System.Windows.Forms.Button
    Friend WithEvents panTimePoints As System.Windows.Forms.Panel
    Friend WithEvents pan1 As System.Windows.Forms.Panel
    Friend WithEvents cmdCancel1 As System.Windows.Forms.Button
    Friend WithEvents cmdOK1 As System.Windows.Forms.Button
    Friend WithEvents cmdAddTP As System.Windows.Forms.Button
    Friend WithEvents lblUnique As System.Windows.Forms.Label
    Friend WithEvents rbBothNames As System.Windows.Forms.RadioButton
    Friend WithEvents gbxIncrement As System.Windows.Forms.GroupBox
    Friend WithEvents rbGroup As System.Windows.Forms.RadioButton
    Friend WithEvents rbEntire As System.Windows.Forms.RadioButton
    Friend WithEvents lbltxtUniqueID As System.Windows.Forms.Label
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdEdit As System.Windows.Forms.Button
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
End Class
