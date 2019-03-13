<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmViewAllGroups
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
        Me.lbxGroupTables = New System.Windows.Forms.ListBox()
        Me.lblGroupTables = New System.Windows.Forms.Label()
        Me.dgvFormGroups = New System.Windows.Forms.DataGridView()
        Me.lblGroupTableContent = New System.Windows.Forms.Label()
        Me.cmdCopyAll = New System.Windows.Forms.Button()
        Me.lblTitle = New System.Windows.Forms.Label()
        CType(Me.dgvFormGroups, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lbxGroupTables
        '
        Me.lbxGroupTables.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbxGroupTables.FormattingEnabled = True
        Me.lbxGroupTables.ItemHeight = 17
        Me.lbxGroupTables.Location = New System.Drawing.Point(17, 110)
        Me.lbxGroupTables.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.lbxGroupTables.Name = "lbxGroupTables"
        Me.lbxGroupTables.Size = New System.Drawing.Size(166, 191)
        Me.lbxGroupTables.TabIndex = 0
        '
        'lblGroupTables
        '
        Me.lblGroupTables.AutoSize = True
        Me.lblGroupTables.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGroupTables.Location = New System.Drawing.Point(14, 91)
        Me.lblGroupTables.Name = "lblGroupTables"
        Me.lblGroupTables.Size = New System.Drawing.Size(103, 16)
        Me.lblGroupTables.TabIndex = 1
        Me.lblGroupTables.Text = "Group Tables"
        '
        'dgvFormGroups
        '
        Me.dgvFormGroups.AllowUserToAddRows = False
        Me.dgvFormGroups.AllowUserToDeleteRows = False
        Me.dgvFormGroups.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvFormGroups.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvFormGroups.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvFormGroups.BackgroundColor = System.Drawing.Color.White
        Me.dgvFormGroups.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvFormGroups.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvFormGroups.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvFormGroups.Location = New System.Drawing.Point(188, 110)
        Me.dgvFormGroups.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvFormGroups.Name = "dgvFormGroups"
        Me.dgvFormGroups.ReadOnly = True
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvFormGroups.RowHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.dgvFormGroups.RowHeadersWidth = 25
        DataGridViewCellStyle3.Padding = New System.Windows.Forms.Padding(0, 3, 0, 3)
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvFormGroups.RowsDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvFormGroups.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvFormGroups.Size = New System.Drawing.Size(709, 531)
        Me.dgvFormGroups.TabIndex = 156
        '
        'lblGroupTableContent
        '
        Me.lblGroupTableContent.AutoSize = True
        Me.lblGroupTableContent.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGroupTableContent.Location = New System.Drawing.Point(184, 91)
        Me.lblGroupTableContent.Name = "lblGroupTableContent"
        Me.lblGroupTableContent.Size = New System.Drawing.Size(159, 16)
        Me.lblGroupTableContent.TabIndex = 157
        Me.lblGroupTableContent.Text = "Group Table Contents"
        '
        'cmdCopyAll
        '
        Me.cmdCopyAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCopyAll.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCopyAll.FlatAppearance.BorderSize = 0
        Me.cmdCopyAll.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCopyAll.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCopyAll.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdCopyAll.Location = New System.Drawing.Point(772, 47)
        Me.cmdCopyAll.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCopyAll.Name = "cmdCopyAll"
        Me.cmdCopyAll.Size = New System.Drawing.Size(125, 55)
        Me.cmdCopyAll.TabIndex = 159
        Me.cmdCopyAll.Text = "&Copy All Tables to Clipboard"
        Me.cmdCopyAll.UseVisualStyleBackColor = False
        '
        'lblTitle
        '
        Me.lblTitle.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.Location = New System.Drawing.Point(0, 0)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(911, 44)
        Me.lblTitle.TabIndex = 160
        Me.lblTitle.Text = "This window intended for administrative troubleshooting"
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmViewAllGroups
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(911, 654)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.cmdCopyAll)
        Me.Controls.Add(Me.lblGroupTableContent)
        Me.Controls.Add(Me.dgvFormGroups)
        Me.Controls.Add(Me.lblGroupTables)
        Me.Controls.Add(Me.lbxGroupTables)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmViewAllGroups"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "View All Group Information..."
        CType(Me.dgvFormGroups, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lbxGroupTables As System.Windows.Forms.ListBox
    Friend WithEvents lblGroupTables As System.Windows.Forms.Label
    Friend WithEvents dgvFormGroups As System.Windows.Forms.DataGridView
    Friend WithEvents lblGroupTableContent As System.Windows.Forms.Label
    Friend WithEvents cmdCopyAll As System.Windows.Forms.Button
    Friend WithEvents lblTitle As System.Windows.Forms.Label
End Class
