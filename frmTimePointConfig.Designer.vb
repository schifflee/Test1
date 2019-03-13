<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTimePointConfig
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
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTimePointConfig))
        Me.dgvTimePoints = New System.Windows.Forms.DataGridView()
        Me.dgvTimepointSets = New System.Windows.Forms.DataGridView()
        Me.txtMin = New System.Windows.Forms.TextBox()
        Me.lblMin = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtHour = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtDay = New System.Windows.Forms.TextBox()
        Me.cmdRemove = New System.Windows.Forms.Button()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.dgvTP = New System.Windows.Forms.DataGridView()
        Me.cmdDeactivate = New System.Windows.Forms.Button()
        Me.gbxShow = New System.Windows.Forms.GroupBox()
        Me.rbInactive = New System.Windows.Forms.RadioButton()
        Me.rbActive = New System.Windows.Forms.RadioButton()
        Me.rbAll = New System.Windows.Forms.RadioButton()
        Me.panSets = New System.Windows.Forms.Panel()
        Me.lblTP = New System.Windows.Forms.Label()
        Me.cmdActivate = New System.Windows.Forms.Button()
        Me.cmdDelete = New System.Windows.Forms.Button()
        Me.cmdApply = New System.Windows.Forms.Button()
        Me.panActive = New System.Windows.Forms.Panel()
        Me.cmdSaveSet = New System.Windows.Forms.Button()
        Me.pan1 = New System.Windows.Forms.Panel()
        Me.cmdCancel1 = New System.Windows.Forms.Button()
        Me.cmdOK1 = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        CType(Me.dgvTimePoints, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvTimepointSets, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvTP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbxShow.SuspendLayout()
        Me.panSets.SuspendLayout()
        Me.panActive.SuspendLayout()
        Me.pan1.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvTimePoints
        '
        Me.dgvTimePoints.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvTimePoints.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvTimePoints.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvTimePoints.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvTimePoints.Location = New System.Drawing.Point(207, 32)
        Me.dgvTimePoints.Name = "dgvTimePoints"
        Me.dgvTimePoints.ReadOnly = True
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvTimePoints.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvTimePoints.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvTimePoints.Size = New System.Drawing.Size(73, 345)
        Me.dgvTimePoints.TabIndex = 163
        Me.dgvTimePoints.TabStop = False
        '
        'dgvTimepointSets
        '
        Me.dgvTimepointSets.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvTimepointSets.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.dgvTimepointSets.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvTimepointSets.DefaultCellStyle = DataGridViewCellStyle5
        Me.dgvTimepointSets.Location = New System.Drawing.Point(102, 76)
        Me.dgvTimepointSets.MultiSelect = False
        Me.dgvTimepointSets.Name = "dgvTimepointSets"
        Me.dgvTimepointSets.ReadOnly = True
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvTimepointSets.RowHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.dgvTimepointSets.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvTimepointSets.Size = New System.Drawing.Size(97, 351)
        Me.dgvTimepointSets.TabIndex = 164
        Me.dgvTimepointSets.TabStop = False
        '
        'txtMin
        '
        Me.txtMin.Location = New System.Drawing.Point(11, 32)
        Me.txtMin.Name = "txtMin"
        Me.txtMin.Size = New System.Drawing.Size(60, 20)
        Me.txtMin.TabIndex = 0
        Me.txtMin.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblMin
        '
        Me.lblMin.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMin.Location = New System.Drawing.Point(11, 9)
        Me.lblMin.Name = "lblMin"
        Me.lblMin.Size = New System.Drawing.Size(60, 20)
        Me.lblMin.TabIndex = 166
        Me.lblMin.Text = "Minutes"
        Me.lblMin.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(77, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 20)
        Me.Label1.TabIndex = 168
        Me.Label1.Text = "Hours"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'txtHour
        '
        Me.txtHour.Location = New System.Drawing.Point(77, 32)
        Me.txtHour.Name = "txtHour"
        Me.txtHour.Size = New System.Drawing.Size(60, 20)
        Me.txtHour.TabIndex = 1
        Me.txtHour.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(143, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 20)
        Me.Label2.TabIndex = 170
        Me.Label2.Text = "Days"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'txtDay
        '
        Me.txtDay.Location = New System.Drawing.Point(143, 32)
        Me.txtDay.Name = "txtDay"
        Me.txtDay.Size = New System.Drawing.Size(60, 20)
        Me.txtDay.TabIndex = 2
        Me.txtDay.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'cmdRemove
        '
        Me.cmdRemove.ForeColor = System.Drawing.Color.Red
        Me.cmdRemove.Location = New System.Drawing.Point(133, 91)
        Me.cmdRemove.Name = "cmdRemove"
        Me.cmdRemove.Size = New System.Drawing.Size(70, 22)
        Me.cmdRemove.TabIndex = 173
        Me.cmdRemove.Text = "<- Remove"
        Me.cmdRemove.UseVisualStyleBackColor = True
        '
        'cmdAdd
        '
        Me.cmdAdd.ForeColor = System.Drawing.Color.Blue
        Me.cmdAdd.Location = New System.Drawing.Point(133, 63)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(70, 22)
        Me.cmdAdd.TabIndex = 172
        Me.cmdAdd.Text = "Add ->"
        Me.cmdAdd.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(204, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(76, 20)
        Me.Label3.TabIndex = 174
        Me.Label3.Text = "Time Points"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(99, 52)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 20)
        Me.Label4.TabIndex = 175
        Me.Label4.Text = "Time Point Sets"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(202, 52)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(76, 20)
        Me.Label5.TabIndex = 177
        Me.Label5.Text = "Time Points"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'dgvTP
        '
        Me.dgvTP.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvTP.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle7
        Me.dgvTP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvTP.DefaultCellStyle = DataGridViewCellStyle8
        Me.dgvTP.Location = New System.Drawing.Point(205, 76)
        Me.dgvTP.MultiSelect = False
        Me.dgvTP.Name = "dgvTP"
        Me.dgvTP.ReadOnly = True
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvTP.RowHeadersDefaultCellStyle = DataGridViewCellStyle9
        Me.dgvTP.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvTP.Size = New System.Drawing.Size(73, 351)
        Me.dgvTP.TabIndex = 176
        Me.dgvTP.TabStop = False
        '
        'cmdDeactivate
        '
        Me.cmdDeactivate.ForeColor = System.Drawing.Color.Maroon
        Me.cmdDeactivate.Location = New System.Drawing.Point(16, 249)
        Me.cmdDeactivate.Name = "cmdDeactivate"
        Me.cmdDeactivate.Size = New System.Drawing.Size(80, 22)
        Me.cmdDeactivate.TabIndex = 179
        Me.cmdDeactivate.Text = "Deactivate ->"
        Me.cmdDeactivate.UseVisualStyleBackColor = True
        '
        'gbxShow
        '
        Me.gbxShow.Controls.Add(Me.rbInactive)
        Me.gbxShow.Controls.Add(Me.rbActive)
        Me.gbxShow.Controls.Add(Me.rbAll)
        Me.gbxShow.Location = New System.Drawing.Point(16, 119)
        Me.gbxShow.Name = "gbxShow"
        Me.gbxShow.Size = New System.Drawing.Size(80, 96)
        Me.gbxShow.TabIndex = 180
        Me.gbxShow.TabStop = False
        Me.gbxShow.Text = "Show"
        '
        'rbInactive
        '
        Me.rbInactive.AutoSize = True
        Me.rbInactive.Location = New System.Drawing.Point(7, 68)
        Me.rbInactive.Name = "rbInactive"
        Me.rbInactive.Size = New System.Drawing.Size(63, 17)
        Me.rbInactive.TabIndex = 2
        Me.rbInactive.Text = "Inactive"
        Me.rbInactive.UseVisualStyleBackColor = True
        '
        'rbActive
        '
        Me.rbActive.AutoSize = True
        Me.rbActive.Checked = True
        Me.rbActive.Location = New System.Drawing.Point(7, 45)
        Me.rbActive.Name = "rbActive"
        Me.rbActive.Size = New System.Drawing.Size(55, 17)
        Me.rbActive.TabIndex = 1
        Me.rbActive.TabStop = True
        Me.rbActive.Text = "Active"
        Me.rbActive.UseVisualStyleBackColor = True
        '
        'rbAll
        '
        Me.rbAll.AutoSize = True
        Me.rbAll.Location = New System.Drawing.Point(7, 22)
        Me.rbAll.Name = "rbAll"
        Me.rbAll.Size = New System.Drawing.Size(36, 17)
        Me.rbAll.TabIndex = 0
        Me.rbAll.Text = "All"
        Me.rbAll.UseVisualStyleBackColor = True
        '
        'panSets
        '
        Me.panSets.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panSets.Controls.Add(Me.lblTP)
        Me.panSets.Controls.Add(Me.cmdActivate)
        Me.panSets.Controls.Add(Me.cmdDelete)
        Me.panSets.Controls.Add(Me.cmdApply)
        Me.panSets.Controls.Add(Me.dgvTimepointSets)
        Me.panSets.Controls.Add(Me.dgvTP)
        Me.panSets.Controls.Add(Me.Label5)
        Me.panSets.Controls.Add(Me.Label4)
        Me.panSets.Controls.Add(Me.gbxShow)
        Me.panSets.Controls.Add(Me.cmdDeactivate)
        Me.panSets.Location = New System.Drawing.Point(387, 71)
        Me.panSets.Name = "panSets"
        Me.panSets.Size = New System.Drawing.Size(296, 436)
        Me.panSets.TabIndex = 181
        '
        'lblTP
        '
        Me.lblTP.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTP.ForeColor = System.Drawing.Color.Blue
        Me.lblTP.Location = New System.Drawing.Point(0, 0)
        Me.lblTP.Name = "lblTP"
        Me.lblTP.Size = New System.Drawing.Size(295, 52)
        Me.lblTP.TabIndex = 184
        Me.lblTP.Text = "NOTE: All actions performed in the Time Point Saved Sets pane are final and canno" & _
    "t be reversed."
        Me.lblTP.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'cmdActivate
        '
        Me.cmdActivate.ForeColor = System.Drawing.Color.Blue
        Me.cmdActivate.Location = New System.Drawing.Point(16, 221)
        Me.cmdActivate.Name = "cmdActivate"
        Me.cmdActivate.Size = New System.Drawing.Size(80, 22)
        Me.cmdActivate.TabIndex = 183
        Me.cmdActivate.Text = "Activate ->"
        Me.cmdActivate.UseVisualStyleBackColor = True
        '
        'cmdDelete
        '
        Me.cmdDelete.ForeColor = System.Drawing.Color.Red
        Me.cmdDelete.Location = New System.Drawing.Point(16, 277)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(80, 22)
        Me.cmdDelete.TabIndex = 182
        Me.cmdDelete.Text = "Delete ->"
        Me.cmdDelete.UseVisualStyleBackColor = True
        '
        'cmdApply
        '
        Me.cmdApply.ForeColor = System.Drawing.Color.Blue
        Me.cmdApply.Location = New System.Drawing.Point(16, 76)
        Me.cmdApply.Name = "cmdApply"
        Me.cmdApply.Size = New System.Drawing.Size(80, 37)
        Me.cmdApply.TabIndex = 181
        Me.cmdApply.Text = "<-- Apply to Timepoints"
        Me.cmdApply.UseVisualStyleBackColor = True
        '
        'panActive
        '
        Me.panActive.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panActive.Controls.Add(Me.cmdSaveSet)
        Me.panActive.Controls.Add(Me.dgvTimePoints)
        Me.panActive.Controls.Add(Me.pan1)
        Me.panActive.Controls.Add(Me.txtMin)
        Me.panActive.Controls.Add(Me.lblMin)
        Me.panActive.Controls.Add(Me.Label3)
        Me.panActive.Controls.Add(Me.txtHour)
        Me.panActive.Controls.Add(Me.cmdRemove)
        Me.panActive.Controls.Add(Me.Label1)
        Me.panActive.Controls.Add(Me.cmdAdd)
        Me.panActive.Controls.Add(Me.txtDay)
        Me.panActive.Controls.Add(Me.Label2)
        Me.panActive.Location = New System.Drawing.Point(15, 71)
        Me.panActive.Name = "panActive"
        Me.panActive.Size = New System.Drawing.Size(294, 436)
        Me.panActive.TabIndex = 182
        '
        'cmdSaveSet
        '
        Me.cmdSaveSet.ForeColor = System.Drawing.Color.Blue
        Me.cmdSaveSet.Location = New System.Drawing.Point(123, 176)
        Me.cmdSaveSet.Name = "cmdSaveSet"
        Me.cmdSaveSet.Size = New System.Drawing.Size(80, 37)
        Me.cmdSaveSet.TabIndex = 182
        Me.cmdSaveSet.Text = "Save as Timepoint Set"
        Me.cmdSaveSet.UseVisualStyleBackColor = True
        '
        'pan1
        '
        Me.pan1.Controls.Add(Me.cmdCancel1)
        Me.pan1.Controls.Add(Me.cmdOK1)
        Me.pan1.Location = New System.Drawing.Point(3, 383)
        Me.pan1.Name = "pan1"
        Me.pan1.Size = New System.Drawing.Size(231, 45)
        Me.pan1.TabIndex = 185
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
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(12, 52)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(90, 16)
        Me.Label6.TabIndex = 183
        Me.Label6.Text = "Time Points"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(384, 52)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(166, 16)
        Me.Label7.TabIndex = 184
        Me.Label7.Text = "Time Point Saved Sets"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'lblTitle
        '
        Me.lblTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.Color.Blue
        Me.lblTitle.Location = New System.Drawing.Point(138, 9)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(545, 34)
        Me.lblTitle.TabIndex = 186
        Me.lblTitle.Text = "Time Points"
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Blue
        Me.Label8.Location = New System.Drawing.Point(12, 9)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(128, 34)
        Me.Label8.TabIndex = 187
        Me.Label8.Text = "Time Points For:"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'frmTimePointConfig
        '
        Me.AcceptButton = Me.cmdAdd
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(700, 528)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.panActive)
        Me.Controls.Add(Me.panSets)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmTimePointConfig"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Time Point Configuration"
        CType(Me.dgvTimePoints, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvTimepointSets, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvTP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbxShow.ResumeLayout(False)
        Me.gbxShow.PerformLayout()
        Me.panSets.ResumeLayout(False)
        Me.panActive.ResumeLayout(False)
        Me.panActive.PerformLayout()
        Me.pan1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvTimePoints As System.Windows.Forms.DataGridView
    Friend WithEvents dgvTimepointSets As System.Windows.Forms.DataGridView
    Friend WithEvents txtMin As System.Windows.Forms.TextBox
    Friend WithEvents lblMin As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtHour As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtDay As System.Windows.Forms.TextBox
    Friend WithEvents cmdRemove As System.Windows.Forms.Button
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dgvTP As System.Windows.Forms.DataGridView
    Friend WithEvents cmdDeactivate As System.Windows.Forms.Button
    Friend WithEvents gbxShow As System.Windows.Forms.GroupBox
    Friend WithEvents rbInactive As System.Windows.Forms.RadioButton
    Friend WithEvents rbActive As System.Windows.Forms.RadioButton
    Friend WithEvents rbAll As System.Windows.Forms.RadioButton
    Friend WithEvents panSets As System.Windows.Forms.Panel
    Friend WithEvents cmdApply As System.Windows.Forms.Button
    Friend WithEvents panActive As System.Windows.Forms.Panel
    Friend WithEvents cmdSaveSet As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents pan1 As System.Windows.Forms.Panel
    Friend WithEvents cmdCancel1 As System.Windows.Forms.Button
    Friend WithEvents cmdOK1 As System.Windows.Forms.Button
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents cmdDelete As System.Windows.Forms.Button
    Friend WithEvents cmdActivate As System.Windows.Forms.Button
    Friend WithEvents lblTP As System.Windows.Forms.Label
End Class
