<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmApplyTemplateTables
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
        Me.cmdExitSave = New System.Windows.Forms.Button()
        Me.cmdExitNoSave = New System.Windows.Forms.Button()
        Me.lblExisting = New System.Windows.Forms.Label()
        Me.lblProposed = New System.Windows.Forms.Label()
        Me.lblTemplates = New System.Windows.Forms.Label()
        Me.dgvS = New System.Windows.Forms.DataGridView()
        Me.dgvT = New System.Windows.Forms.DataGridView()
        Me.dgvD = New System.Windows.Forms.DataGridView()
        Me.lbl1 = New System.Windows.Forms.Label()
        Me.lbl2 = New System.Windows.Forms.Label()
        CType(Me.dgvS, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvT, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvD, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdExitSave
        '
        Me.cmdExitSave.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExitSave.ForeColor = System.Drawing.Color.Blue
        Me.cmdExitSave.Location = New System.Drawing.Point(368, 11)
        Me.cmdExitSave.Name = "cmdExitSave"
        Me.cmdExitSave.Size = New System.Drawing.Size(80, 49)
        Me.cmdExitSave.TabIndex = 144
        Me.cmdExitSave.Text = "&Apply"
        Me.cmdExitSave.UseVisualStyleBackColor = True
        '
        'cmdExitNoSave
        '
        Me.cmdExitNoSave.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExitNoSave.ForeColor = System.Drawing.Color.Red
        Me.cmdExitNoSave.Location = New System.Drawing.Point(454, 12)
        Me.cmdExitNoSave.Name = "cmdExitNoSave"
        Me.cmdExitNoSave.Size = New System.Drawing.Size(80, 48)
        Me.cmdExitNoSave.TabIndex = 143
        Me.cmdExitNoSave.Text = "&Cancel"
        Me.cmdExitNoSave.UseVisualStyleBackColor = True
        '
        'lblExisting
        '
        Me.lblExisting.AutoSize = True
        Me.lblExisting.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblExisting.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblExisting.Location = New System.Drawing.Point(12, 135)
        Me.lblExisting.Name = "lblExisting"
        Me.lblExisting.Size = New System.Drawing.Size(101, 17)
        Me.lblExisting.TabIndex = 145
        Me.lblExisting.Text = "Existing Tables"
        '
        'lblProposed
        '
        Me.lblProposed.AutoSize = True
        Me.lblProposed.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProposed.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblProposed.Location = New System.Drawing.Point(537, 135)
        Me.lblProposed.Name = "lblProposed"
        Me.lblProposed.Size = New System.Drawing.Size(110, 17)
        Me.lblProposed.TabIndex = 146
        Me.lblProposed.Text = "Proposed Tables"
        '
        'lblTemplates
        '
        Me.lblTemplates.AutoSize = True
        Me.lblTemplates.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTemplates.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblTemplates.Location = New System.Drawing.Point(377, 135)
        Me.lblTemplates.Name = "lblTemplates"
        Me.lblTemplates.Size = New System.Drawing.Size(111, 17)
        Me.lblTemplates.TabIndex = 148
        Me.lblTemplates.Text = "Study Templates"
        '
        'dgvS
        '
        Me.dgvS.AllowUserToAddRows = False
        Me.dgvS.AllowUserToDeleteRows = False
        Me.dgvS.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dgvS.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvS.BackgroundColor = System.Drawing.Color.White
        Me.dgvS.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvS.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvS.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.Padding = New System.Windows.Forms.Padding(0, 2, 0, 2)
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvS.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvS.Location = New System.Drawing.Point(15, 155)
        Me.dgvS.Name = "dgvS"
        Me.dgvS.ReadOnly = True
        Me.dgvS.RowHeadersWidth = 25
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvS.RowsDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvS.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvS.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvS.Size = New System.Drawing.Size(359, 380)
        Me.dgvS.TabIndex = 149
        '
        'dgvT
        '
        Me.dgvT.AllowUserToAddRows = False
        Me.dgvT.AllowUserToDeleteRows = False
        Me.dgvT.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dgvT.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvT.BackgroundColor = System.Drawing.Color.White
        Me.dgvT.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvT.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.dgvT.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.Padding = New System.Windows.Forms.Padding(0, 2, 0, 2)
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvT.DefaultCellStyle = DataGridViewCellStyle5
        Me.dgvT.Location = New System.Drawing.Point(380, 155)
        Me.dgvT.MultiSelect = False
        Me.dgvT.Name = "dgvT"
        Me.dgvT.ReadOnly = True
        Me.dgvT.RowHeadersWidth = 25
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvT.RowsDefaultCellStyle = DataGridViewCellStyle6
        Me.dgvT.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvT.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvT.Size = New System.Drawing.Size(154, 380)
        Me.dgvT.TabIndex = 150
        '
        'dgvD
        '
        Me.dgvD.AllowUserToAddRows = False
        Me.dgvD.AllowUserToDeleteRows = False
        Me.dgvD.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvD.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvD.BackgroundColor = System.Drawing.Color.White
        Me.dgvD.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvD.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle7
        Me.dgvD.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle8.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle8.Padding = New System.Windows.Forms.Padding(0, 2, 0, 2)
        DataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvD.DefaultCellStyle = DataGridViewCellStyle8
        Me.dgvD.Location = New System.Drawing.Point(540, 155)
        Me.dgvD.MultiSelect = False
        Me.dgvD.Name = "dgvD"
        Me.dgvD.ReadOnly = True
        Me.dgvD.RowHeadersWidth = 25
        DataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvD.RowsDefaultCellStyle = DataGridViewCellStyle9
        Me.dgvD.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvD.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvD.Size = New System.Drawing.Size(308, 380)
        Me.dgvD.TabIndex = 151
        '
        'lbl1
        '
        Me.lbl1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl1.Location = New System.Drawing.Point(12, 12)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(350, 108)
        Me.lbl1.TabIndex = 152
        Me.lbl1.Text = "1. Choose a Study Template to view tables belonging to this template." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "2. Click" & _
    " Apply to apply these tables to the underlying study."
        '
        'lbl2
        '
        Me.lbl2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl2.Location = New System.Drawing.Point(537, 77)
        Me.lbl2.Name = "lbl2"
        Me.lbl2.Size = New System.Drawing.Size(308, 58)
        Me.lbl2.TabIndex = 153
        Me.lbl2.Text = "The Word Report Template associated with the chosen Study Template will be assign" & _
    "ed to the underlying study"
        '
        'frmApplyTemplateTables
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(860, 547)
        Me.ControlBox = False
        Me.Controls.Add(Me.lbl2)
        Me.Controls.Add(Me.lbl1)
        Me.Controls.Add(Me.dgvD)
        Me.Controls.Add(Me.dgvT)
        Me.Controls.Add(Me.dgvS)
        Me.Controls.Add(Me.lblTemplates)
        Me.Controls.Add(Me.lblProposed)
        Me.Controls.Add(Me.lblExisting)
        Me.Controls.Add(Me.cmdExitSave)
        Me.Controls.Add(Me.cmdExitNoSave)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "frmApplyTemplateTables"
        Me.ShowInTaskbar = False
        Me.Text = "Apply Study Template Tables..."
        CType(Me.dgvS, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvT, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvD, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdExitSave As System.Windows.Forms.Button
    Friend WithEvents cmdExitNoSave As System.Windows.Forms.Button
    Friend WithEvents lblExisting As System.Windows.Forms.Label
    Friend WithEvents lblProposed As System.Windows.Forms.Label
    Friend WithEvents lblTemplates As System.Windows.Forms.Label
    Friend WithEvents dgvS As System.Windows.Forms.DataGridView
    Friend WithEvents dgvT As System.Windows.Forms.DataGridView
    Friend WithEvents dgvD As System.Windows.Forms.DataGridView
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Friend WithEvents lbl2 As System.Windows.Forms.Label
End Class
