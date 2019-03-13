<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmIncSamplesAssignSamples
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmIncSamplesAssignSamples))
        Me.dgvDesignSample = New System.Windows.Forms.DataGridView()
        Me.dgvAllInjections = New System.Windows.Forms.DataGridView()
        Me.dgvIncSamplesOrig = New System.Windows.Forms.DataGridView()
        Me.panISR = New System.Windows.Forms.Panel()
        Me.panSave = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.panISRdata = New System.Windows.Forms.Panel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lbltxtNS3 = New System.Windows.Forms.Label()
        Me.txtNS3 = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rbAllMatchedSamples = New System.Windows.Forms.RadioButton()
        Me.rbSourceMatched = New System.Windows.Forms.RadioButton()
        Me.rbSourceRepeated = New System.Windows.Forms.RadioButton()
        Me.rbSourceAll = New System.Windows.Forms.RadioButton()
        Me.panDesignSample = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lbltxtNS1 = New System.Windows.Forms.Label()
        Me.txtNS1 = New System.Windows.Forms.TextBox()
        Me.gbOriginal = New System.Windows.Forms.GroupBox()
        Me.optRepeatsO = New System.Windows.Forms.RadioButton()
        Me.rbAllO = New System.Windows.Forms.RadioButton()
        Me.panISRi = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.panISRButtons = New System.Windows.Forms.Panel()
        Me.cmdRemoveISR = New System.Windows.Forms.Button()
        Me.cmdAddISR = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtNS4 = New System.Windows.Forms.TextBox()
        Me.dgvIncSamplesISR = New System.Windows.Forms.DataGridView()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.rbAllAssigned = New System.Windows.Forms.RadioButton()
        Me.rbAssignedMatch = New System.Windows.Forms.RadioButton()
        Me.lblISR = New System.Windows.Forms.Label()
        Me.panO = New System.Windows.Forms.Panel()
        Me.panOButtons = New System.Windows.Forms.Panel()
        Me.cmdRemoveO = New System.Windows.Forms.Button()
        Me.cmdAddOrig = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtNS2 = New System.Windows.Forms.TextBox()
        Me.lblO = New System.Windows.Forms.Label()
        CType(Me.dgvDesignSample, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvAllInjections, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvIncSamplesOrig, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panISR.SuspendLayout()
        Me.panSave.SuspendLayout()
        Me.panISRdata.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.panDesignSample.SuspendLayout()
        Me.gbOriginal.SuspendLayout()
        Me.panISRi.SuspendLayout()
        Me.panISRButtons.SuspendLayout()
        CType(Me.dgvIncSamplesISR, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.panO.SuspendLayout()
        Me.panOButtons.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvDesignSample
        '
        Me.dgvDesignSample.AllowUserToAddRows = False
        Me.dgvDesignSample.AllowUserToDeleteRows = False
        Me.dgvDesignSample.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvDesignSample.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvDesignSample.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvDesignSample.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDesignSample.Location = New System.Drawing.Point(5, 77)
        Me.dgvDesignSample.Name = "dgvDesignSample"
        Me.dgvDesignSample.ReadOnly = True
        Me.dgvDesignSample.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvDesignSample.Size = New System.Drawing.Size(381, 179)
        Me.dgvDesignSample.TabIndex = 0
        '
        'dgvAllInjections
        '
        Me.dgvAllInjections.AllowUserToAddRows = False
        Me.dgvAllInjections.AllowUserToDeleteRows = False
        Me.dgvAllInjections.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvAllInjections.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvAllInjections.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.dgvAllInjections.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvAllInjections.Location = New System.Drawing.Point(6, 78)
        Me.dgvAllInjections.Name = "dgvAllInjections"
        Me.dgvAllInjections.ReadOnly = True
        Me.dgvAllInjections.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvAllInjections.Size = New System.Drawing.Size(379, 178)
        Me.dgvAllInjections.TabIndex = 1
        '
        'dgvIncSamplesOrig
        '
        Me.dgvIncSamplesOrig.AllowUserToAddRows = False
        Me.dgvIncSamplesOrig.AllowUserToDeleteRows = False
        Me.dgvIncSamplesOrig.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvIncSamplesOrig.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvIncSamplesOrig.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvIncSamplesOrig.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvIncSamplesOrig.Location = New System.Drawing.Point(6, 60)
        Me.dgvIncSamplesOrig.Name = "dgvIncSamplesOrig"
        Me.dgvIncSamplesOrig.ReadOnly = True
        Me.dgvIncSamplesOrig.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvIncSamplesOrig.Size = New System.Drawing.Size(380, 275)
        Me.dgvIncSamplesOrig.TabIndex = 2
        '
        'panISR
        '
        Me.panISR.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.panISR.Controls.Add(Me.panSave)
        Me.panISR.Controls.Add(Me.panISRdata)
        Me.panISR.Controls.Add(Me.panDesignSample)
        Me.panISR.Controls.Add(Me.panISRi)
        Me.panISR.Controls.Add(Me.panO)
        Me.panISR.Location = New System.Drawing.Point(12, 12)
        Me.panISR.Name = "panISR"
        Me.panISR.Size = New System.Drawing.Size(797, 664)
        Me.panISR.TabIndex = 3
        '
        'panSave
        '
        Me.panSave.Controls.Add(Me.cmdCancel)
        Me.panSave.Controls.Add(Me.cmdSave)
        Me.panSave.Location = New System.Drawing.Point(3, 2)
        Me.panSave.Name = "panSave"
        Me.panSave.Size = New System.Drawing.Size(258, 57)
        Me.panSave.TabIndex = 14
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.cmdCancel.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(89, 3)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(74, 51)
        Me.cmdCancel.TabIndex = 5
        Me.cmdCancel.Text = "&Cancel and Exit"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdSave
        '
        Me.cmdSave.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.cmdSave.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ForeColor = System.Drawing.Color.Blue
        Me.cmdSave.Location = New System.Drawing.Point(3, 3)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(80, 51)
        Me.cmdSave.TabIndex = 4
        Me.cmdSave.Text = "&Save and Exit"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'panISRdata
        '
        Me.panISRdata.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.panISRdata.Controls.Add(Me.Label3)
        Me.panISRdata.Controls.Add(Me.lbltxtNS3)
        Me.panISRdata.Controls.Add(Me.txtNS3)
        Me.panISRdata.Controls.Add(Me.dgvAllInjections)
        Me.panISRdata.Controls.Add(Me.GroupBox1)
        Me.panISRdata.Location = New System.Drawing.Point(402, 60)
        Me.panISRdata.Name = "panISRdata"
        Me.panISRdata.Size = New System.Drawing.Size(392, 263)
        Me.panISRdata.TabIndex = 13
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Blue
        Me.Label3.Location = New System.Drawing.Point(3, 59)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(149, 16)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Matched ISR Source"
        '
        'lbltxtNS3
        '
        Me.lbltxtNS3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbltxtNS3.AutoSize = True
        Me.lbltxtNS3.Location = New System.Drawing.Point(313, 24)
        Me.lbltxtNS3.Name = "lbltxtNS3"
        Me.lbltxtNS3.Size = New System.Drawing.Size(72, 13)
        Me.lbltxtNS3.TabIndex = 5
        Me.lbltxtNS3.Text = "# of Samples:"
        '
        'txtNS3
        '
        Me.txtNS3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNS3.BackColor = System.Drawing.Color.White
        Me.txtNS3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtNS3.Location = New System.Drawing.Point(353, 41)
        Me.txtNS3.Name = "txtNS3"
        Me.txtNS3.ReadOnly = True
        Me.txtNS3.Size = New System.Drawing.Size(32, 20)
        Me.txtNS3.TabIndex = 4
        Me.txtNS3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rbAllMatchedSamples)
        Me.GroupBox1.Controls.Add(Me.rbSourceMatched)
        Me.GroupBox1.Controls.Add(Me.rbSourceRepeated)
        Me.GroupBox1.Controls.Add(Me.rbSourceAll)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 6)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(263, 55)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Show Data"
        '
        'rbAllMatchedSamples
        '
        Me.rbAllMatchedSamples.AutoSize = True
        Me.rbAllMatchedSamples.Checked = True
        Me.rbAllMatchedSamples.Location = New System.Drawing.Point(122, 34)
        Me.rbAllMatchedSamples.Name = "rbAllMatchedSamples"
        Me.rbAllMatchedSamples.Size = New System.Drawing.Size(124, 17)
        Me.rbAllMatchedSamples.TabIndex = 3
        Me.rbAllMatchedSamples.TabStop = True
        Me.rbAllMatchedSamples.Text = "All Matched Samples"
        Me.rbAllMatchedSamples.UseVisualStyleBackColor = True
        '
        'rbSourceMatched
        '
        Me.rbSourceMatched.AutoSize = True
        Me.rbSourceMatched.Location = New System.Drawing.Point(6, 34)
        Me.rbSourceMatched.Name = "rbSourceMatched"
        Me.rbSourceMatched.Size = New System.Drawing.Size(110, 17)
        Me.rbSourceMatched.TabIndex = 1
        Me.rbSourceMatched.Text = "Matched Samples"
        Me.rbSourceMatched.UseVisualStyleBackColor = True
        '
        'rbSourceRepeated
        '
        Me.rbSourceRepeated.AutoSize = True
        Me.rbSourceRepeated.Location = New System.Drawing.Point(6, 15)
        Me.rbSourceRepeated.Name = "rbSourceRepeated"
        Me.rbSourceRepeated.Size = New System.Drawing.Size(185, 17)
        Me.rbSourceRepeated.TabIndex = 0
        Me.rbSourceRepeated.Text = "Samples that have been repeated"
        Me.rbSourceRepeated.UseVisualStyleBackColor = True
        '
        'rbSourceAll
        '
        Me.rbSourceAll.Location = New System.Drawing.Point(172, 5)
        Me.rbSourceAll.Name = "rbSourceAll"
        Me.rbSourceAll.Size = New System.Drawing.Size(123, 34)
        Me.rbSourceAll.TabIndex = 2
        Me.rbSourceAll.Text = "All Eligible Acquired Samples"
        Me.rbSourceAll.UseVisualStyleBackColor = True
        Me.rbSourceAll.Visible = False
        '
        'panDesignSample
        '
        Me.panDesignSample.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.panDesignSample.Controls.Add(Me.Label1)
        Me.panDesignSample.Controls.Add(Me.lbltxtNS1)
        Me.panDesignSample.Controls.Add(Me.txtNS1)
        Me.panDesignSample.Controls.Add(Me.dgvDesignSample)
        Me.panDesignSample.Controls.Add(Me.gbOriginal)
        Me.panDesignSample.Location = New System.Drawing.Point(3, 60)
        Me.panDesignSample.Name = "panDesignSample"
        Me.panDesignSample.Size = New System.Drawing.Size(392, 263)
        Me.panDesignSample.TabIndex = 12
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Blue
        Me.Label1.Location = New System.Drawing.Point(3, 59)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(211, 16)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Original Observations Source"
        '
        'lbltxtNS1
        '
        Me.lbltxtNS1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbltxtNS1.AutoSize = True
        Me.lbltxtNS1.Location = New System.Drawing.Point(276, 45)
        Me.lbltxtNS1.Name = "lbltxtNS1"
        Me.lbltxtNS1.Size = New System.Drawing.Size(72, 13)
        Me.lbltxtNS1.TabIndex = 3
        Me.lbltxtNS1.Text = "# of Samples:"
        '
        'txtNS1
        '
        Me.txtNS1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNS1.BackColor = System.Drawing.Color.White
        Me.txtNS1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtNS1.Location = New System.Drawing.Point(354, 41)
        Me.txtNS1.Name = "txtNS1"
        Me.txtNS1.ReadOnly = True
        Me.txtNS1.Size = New System.Drawing.Size(32, 20)
        Me.txtNS1.TabIndex = 2
        Me.txtNS1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'gbOriginal
        '
        Me.gbOriginal.Controls.Add(Me.optRepeatsO)
        Me.gbOriginal.Controls.Add(Me.rbAllO)
        Me.gbOriginal.Location = New System.Drawing.Point(6, 5)
        Me.gbOriginal.Name = "gbOriginal"
        Me.gbOriginal.Size = New System.Drawing.Size(224, 55)
        Me.gbOriginal.TabIndex = 1
        Me.gbOriginal.TabStop = False
        Me.gbOriginal.Text = "Show Data"
        '
        'optRepeatsO
        '
        Me.optRepeatsO.AutoSize = True
        Me.optRepeatsO.Checked = True
        Me.optRepeatsO.Location = New System.Drawing.Point(6, 15)
        Me.optRepeatsO.Name = "optRepeatsO"
        Me.optRepeatsO.Size = New System.Drawing.Size(185, 17)
        Me.optRepeatsO.TabIndex = 0
        Me.optRepeatsO.TabStop = True
        Me.optRepeatsO.Text = "Samples that have been repeated"
        Me.optRepeatsO.UseVisualStyleBackColor = True
        '
        'rbAllO
        '
        Me.rbAllO.AutoSize = True
        Me.rbAllO.Location = New System.Drawing.Point(6, 34)
        Me.rbAllO.Name = "rbAllO"
        Me.rbAllO.Size = New System.Drawing.Size(126, 17)
        Me.rbAllO.TabIndex = 1
        Me.rbAllO.Text = "All Reported Samples"
        Me.rbAllO.UseVisualStyleBackColor = True
        '
        'panISRi
        '
        Me.panISRi.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.panISRi.Controls.Add(Me.Button1)
        Me.panISRi.Controls.Add(Me.panISRButtons)
        Me.panISRi.Controls.Add(Me.Label4)
        Me.panISRi.Controls.Add(Me.txtNS4)
        Me.panISRi.Controls.Add(Me.dgvIncSamplesISR)
        Me.panISRi.Controls.Add(Me.GroupBox2)
        Me.panISRi.Controls.Add(Me.lblISR)
        Me.panISRi.Location = New System.Drawing.Point(402, 323)
        Me.panISRi.Name = "panISRi"
        Me.panISRi.Size = New System.Drawing.Size(392, 338)
        Me.panISRi.TabIndex = 11
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(297, 8)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(65, 26)
        Me.Button1.TabIndex = 10
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'panISRButtons
        '
        Me.panISRButtons.Controls.Add(Me.cmdRemoveISR)
        Me.panISRButtons.Controls.Add(Me.cmdAddISR)
        Me.panISRButtons.Location = New System.Drawing.Point(150, 46)
        Me.panISRButtons.Name = "panISRButtons"
        Me.panISRButtons.Size = New System.Drawing.Size(71, 57)
        Me.panISRButtons.TabIndex = 9
        '
        'cmdRemoveISR
        '
        Me.cmdRemoveISR.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.cmdRemoveISR.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRemoveISR.ForeColor = System.Drawing.Color.Red
        Me.cmdRemoveISR.Location = New System.Drawing.Point(37, 3)
        Me.cmdRemoveISR.Name = "cmdRemoveISR"
        Me.cmdRemoveISR.Size = New System.Drawing.Size(30, 51)
        Me.cmdRemoveISR.TabIndex = 7
        Me.cmdRemoveISR.Text = "RemoveISR"
        Me.cmdRemoveISR.UseVisualStyleBackColor = True
        '
        'cmdAddISR
        '
        Me.cmdAddISR.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.cmdAddISR.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddISR.ForeColor = System.Drawing.Color.Blue
        Me.cmdAddISR.Location = New System.Drawing.Point(4, 3)
        Me.cmdAddISR.Name = "cmdAddISR"
        Me.cmdAddISR.Size = New System.Drawing.Size(30, 51)
        Me.cmdAddISR.TabIndex = 6
        Me.cmdAddISR.Text = "AddISR"
        Me.cmdAddISR.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(275, 44)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "# of Samples:"
        '
        'txtNS4
        '
        Me.txtNS4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNS4.BackColor = System.Drawing.Color.White
        Me.txtNS4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtNS4.Location = New System.Drawing.Point(353, 39)
        Me.txtNS4.Name = "txtNS4"
        Me.txtNS4.ReadOnly = True
        Me.txtNS4.Size = New System.Drawing.Size(32, 20)
        Me.txtNS4.TabIndex = 7
        Me.txtNS4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'dgvIncSamplesISR
        '
        Me.dgvIncSamplesISR.AllowUserToAddRows = False
        Me.dgvIncSamplesISR.AllowUserToDeleteRows = False
        Me.dgvIncSamplesISR.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvIncSamplesISR.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvIncSamplesISR.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.dgvIncSamplesISR.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvIncSamplesISR.Location = New System.Drawing.Point(6, 60)
        Me.dgvIncSamplesISR.Name = "dgvIncSamplesISR"
        Me.dgvIncSamplesISR.ReadOnly = True
        Me.dgvIncSamplesISR.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvIncSamplesISR.Size = New System.Drawing.Size(379, 275)
        Me.dgvIncSamplesISR.TabIndex = 3
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.rbAllAssigned)
        Me.GroupBox2.Controls.Add(Me.rbAssignedMatch)
        Me.GroupBox2.Location = New System.Drawing.Point(6, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(245, 38)
        Me.GroupBox2.TabIndex = 6
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Show Data"
        '
        'rbAllAssigned
        '
        Me.rbAllAssigned.AutoSize = True
        Me.rbAllAssigned.Checked = True
        Me.rbAllAssigned.Location = New System.Drawing.Point(6, 15)
        Me.rbAllAssigned.Name = "rbAllAssigned"
        Me.rbAllAssigned.Size = New System.Drawing.Size(125, 17)
        Me.rbAllAssigned.TabIndex = 0
        Me.rbAllAssigned.TabStop = True
        Me.rbAllAssigned.Text = "All Assigned Samples"
        Me.rbAllAssigned.UseVisualStyleBackColor = True
        '
        'rbAssignedMatch
        '
        Me.rbAssignedMatch.AutoSize = True
        Me.rbAssignedMatch.Location = New System.Drawing.Point(137, 15)
        Me.rbAssignedMatch.Name = "rbAssignedMatch"
        Me.rbAssignedMatch.Size = New System.Drawing.Size(91, 17)
        Me.rbAssignedMatch.TabIndex = 1
        Me.rbAssignedMatch.Text = "Only Matched"
        Me.rbAssignedMatch.UseVisualStyleBackColor = True
        '
        'lblISR
        '
        Me.lblISR.AutoSize = True
        Me.lblISR.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblISR.ForeColor = System.Drawing.Color.Blue
        Me.lblISR.Location = New System.Drawing.Point(3, 41)
        Me.lblISR.Name = "lblISR"
        Me.lblISR.Size = New System.Drawing.Size(96, 16)
        Me.lblISR.TabIndex = 5
        Me.lblISR.Text = "Matched ISR"
        '
        'panO
        '
        Me.panO.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.panO.Controls.Add(Me.panOButtons)
        Me.panO.Controls.Add(Me.Label2)
        Me.panO.Controls.Add(Me.txtNS2)
        Me.panO.Controls.Add(Me.dgvIncSamplesOrig)
        Me.panO.Controls.Add(Me.lblO)
        Me.panO.Location = New System.Drawing.Point(3, 323)
        Me.panO.Name = "panO"
        Me.panO.Size = New System.Drawing.Size(392, 338)
        Me.panO.TabIndex = 10
        '
        'panOButtons
        '
        Me.panOButtons.Controls.Add(Me.cmdRemoveO)
        Me.panOButtons.Controls.Add(Me.cmdAddOrig)
        Me.panOButtons.Location = New System.Drawing.Point(187, 3)
        Me.panOButtons.Name = "panOButtons"
        Me.panOButtons.Size = New System.Drawing.Size(71, 57)
        Me.panOButtons.TabIndex = 7
        '
        'cmdRemoveO
        '
        Me.cmdRemoveO.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.cmdRemoveO.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRemoveO.ForeColor = System.Drawing.Color.Red
        Me.cmdRemoveO.Location = New System.Drawing.Point(37, 3)
        Me.cmdRemoveO.Name = "cmdRemoveO"
        Me.cmdRemoveO.Size = New System.Drawing.Size(30, 51)
        Me.cmdRemoveO.TabIndex = 5
        Me.cmdRemoveO.Text = "RemoveO"
        Me.cmdRemoveO.UseVisualStyleBackColor = True
        '
        'cmdAddOrig
        '
        Me.cmdAddOrig.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.cmdAddOrig.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddOrig.ForeColor = System.Drawing.Color.Blue
        Me.cmdAddOrig.Location = New System.Drawing.Point(4, 3)
        Me.cmdAddOrig.Name = "cmdAddOrig"
        Me.cmdAddOrig.Size = New System.Drawing.Size(30, 51)
        Me.cmdAddOrig.TabIndex = 4
        Me.cmdAddOrig.Text = "AddO"
        Me.cmdAddOrig.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(276, 44)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "# of Samples:"
        '
        'txtNS2
        '
        Me.txtNS2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNS2.BackColor = System.Drawing.Color.White
        Me.txtNS2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtNS2.Location = New System.Drawing.Point(354, 39)
        Me.txtNS2.Name = "txtNS2"
        Me.txtNS2.ReadOnly = True
        Me.txtNS2.Size = New System.Drawing.Size(32, 20)
        Me.txtNS2.TabIndex = 5
        Me.txtNS2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblO
        '
        Me.lblO.AutoSize = True
        Me.lblO.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblO.ForeColor = System.Drawing.Color.Blue
        Me.lblO.Location = New System.Drawing.Point(3, 41)
        Me.lblO.Name = "lblO"
        Me.lblO.Size = New System.Drawing.Size(158, 16)
        Me.lblO.TabIndex = 4
        Me.lblO.Text = "Original Observations"
        '
        'frmIncSamplesAssignSamples
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(821, 688)
        Me.Controls.Add(Me.panISR)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimizeBox = False
        Me.Name = "frmIncSamplesAssignSamples"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "Incurred Sample Repeats Assigned Samples"
        CType(Me.dgvDesignSample, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvAllInjections, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvIncSamplesOrig, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panISR.ResumeLayout(False)
        Me.panSave.ResumeLayout(False)
        Me.panISRdata.ResumeLayout(False)
        Me.panISRdata.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.panDesignSample.ResumeLayout(False)
        Me.panDesignSample.PerformLayout()
        Me.gbOriginal.ResumeLayout(False)
        Me.gbOriginal.PerformLayout()
        Me.panISRi.ResumeLayout(False)
        Me.panISRi.PerformLayout()
        Me.panISRButtons.ResumeLayout(False)
        CType(Me.dgvIncSamplesISR, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.panO.ResumeLayout(False)
        Me.panO.PerformLayout()
        Me.panOButtons.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvDesignSample As System.Windows.Forms.DataGridView
    Friend WithEvents dgvAllInjections As System.Windows.Forms.DataGridView
    Friend WithEvents dgvIncSamplesOrig As System.Windows.Forms.DataGridView
    Friend WithEvents panISR As System.Windows.Forms.Panel
    Friend WithEvents dgvIncSamplesISR As System.Windows.Forms.DataGridView
    Friend WithEvents cmdRemoveISR As System.Windows.Forms.Button
    Friend WithEvents cmdAddISR As System.Windows.Forms.Button
    Friend WithEvents cmdRemoveO As System.Windows.Forms.Button
    Friend WithEvents cmdAddOrig As System.Windows.Forms.Button
    Friend WithEvents panISRi As System.Windows.Forms.Panel
    Friend WithEvents panO As System.Windows.Forms.Panel
    Friend WithEvents panISRdata As System.Windows.Forms.Panel
    Friend WithEvents panDesignSample As System.Windows.Forms.Panel
    Friend WithEvents lblISR As System.Windows.Forms.Label
    Friend WithEvents lblO As System.Windows.Forms.Label
    Friend WithEvents gbOriginal As System.Windows.Forms.GroupBox
    Friend WithEvents rbAllO As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rbSourceRepeated As System.Windows.Forms.RadioButton
    Friend WithEvents rbSourceAll As System.Windows.Forms.RadioButton
    Friend WithEvents optRepeatsO As System.Windows.Forms.RadioButton
    Friend WithEvents rbSourceMatched As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents rbAllAssigned As System.Windows.Forms.RadioButton
    Friend WithEvents rbAssignedMatch As System.Windows.Forms.RadioButton
    Friend WithEvents lbltxtNS1 As System.Windows.Forms.Label
    Friend WithEvents txtNS1 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtNS2 As System.Windows.Forms.TextBox
    Friend WithEvents lbltxtNS3 As System.Windows.Forms.Label
    Friend WithEvents txtNS3 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtNS4 As System.Windows.Forms.TextBox
    Friend WithEvents panOButtons As System.Windows.Forms.Panel
    Friend WithEvents panISRButtons As System.Windows.Forms.Panel
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents rbAllMatchedSamples As System.Windows.Forms.RadioButton
    Friend WithEvents panSave As System.Windows.Forms.Panel
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
