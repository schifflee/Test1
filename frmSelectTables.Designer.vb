<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSelectTables
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSelectTables))
        Me.lbxAnalytes = New System.Windows.Forms.ListBox()
        Me.gbAll = New System.Windows.Forms.GroupBox()
        Me.rbDeselect = New System.Windows.Forms.RadioButton()
        Me.rbSelect = New System.Windows.Forms.RadioButton()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.gbColRow = New System.Windows.Forms.GroupBox()
        Me.rbRows = New System.Windows.Forms.RadioButton()
        Me.rbColumns = New System.Windows.Forms.RadioButton()
        Me.gbAll.SuspendLayout()
        Me.gbColRow.SuspendLayout()
        Me.SuspendLayout()
        '
        'lbxAnalytes
        '
        Me.lbxAnalytes.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbxAnalytes.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbxAnalytes.FormattingEnabled = True
        Me.lbxAnalytes.ItemHeight = 20
        Me.lbxAnalytes.Location = New System.Drawing.Point(134, 168)
        Me.lbxAnalytes.Name = "lbxAnalytes"
        Me.lbxAnalytes.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.lbxAnalytes.Size = New System.Drawing.Size(428, 284)
        Me.lbxAnalytes.TabIndex = 0
        '
        'gbAll
        '
        Me.gbAll.Controls.Add(Me.rbDeselect)
        Me.gbAll.Controls.Add(Me.rbSelect)
        Me.gbAll.Location = New System.Drawing.Point(12, 247)
        Me.gbAll.Name = "gbAll"
        Me.gbAll.Size = New System.Drawing.Size(116, 73)
        Me.gbAll.TabIndex = 1
        Me.gbAll.TabStop = False
        Me.gbAll.Text = "Select or DeSelect"
        '
        'rbDeselect
        '
        Me.rbDeselect.AutoSize = True
        Me.rbDeselect.Location = New System.Drawing.Point(12, 42)
        Me.rbDeselect.Name = "rbDeselect"
        Me.rbDeselect.Size = New System.Drawing.Size(86, 17)
        Me.rbDeselect.TabIndex = 4
        Me.rbDeselect.Text = "De-Select All"
        Me.rbDeselect.UseVisualStyleBackColor = True
        '
        'rbSelect
        '
        Me.rbSelect.AutoSize = True
        Me.rbSelect.Checked = True
        Me.rbSelect.Location = New System.Drawing.Point(12, 19)
        Me.rbSelect.Name = "rbSelect"
        Me.rbSelect.Size = New System.Drawing.Size(69, 17)
        Me.rbSelect.TabIndex = 3
        Me.rbSelect.TabStop = True
        Me.rbSelect.Text = "Select All"
        Me.rbSelect.UseVisualStyleBackColor = True
        '
        'lblTitle
        '
        Me.lblTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.Color.Blue
        Me.lblTitle.Location = New System.Drawing.Point(12, 9)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(550, 156)
        Me.lblTitle.TabIndex = 2
        Me.lblTitle.Text = "lblTitle"
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOK.CausesValidation = False
        Me.cmdOK.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.cmdOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.Blue
        Me.cmdOK.Location = New System.Drawing.Point(24, 326)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(86, 34)
        Me.cmdOK.TabIndex = 6
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancel.CausesValidation = False
        Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(24, 378)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(86, 34)
        Me.cmdCancel.TabIndex = 7
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'gbColRow
        '
        Me.gbColRow.Controls.Add(Me.rbRows)
        Me.gbColRow.Controls.Add(Me.rbColumns)
        Me.gbColRow.Location = New System.Drawing.Point(12, 168)
        Me.gbColRow.Name = "gbColRow"
        Me.gbColRow.Size = New System.Drawing.Size(116, 73)
        Me.gbColRow.TabIndex = 8
        Me.gbColRow.TabStop = False
        Me.gbColRow.Text = "Act On"
        '
        'rbRows
        '
        Me.rbRows.AutoSize = True
        Me.rbRows.Location = New System.Drawing.Point(12, 42)
        Me.rbRows.Name = "rbRows"
        Me.rbRows.Size = New System.Drawing.Size(52, 17)
        Me.rbRows.TabIndex = 4
        Me.rbRows.Text = "Rows"
        Me.rbRows.UseVisualStyleBackColor = True
        '
        'rbColumns
        '
        Me.rbColumns.AutoSize = True
        Me.rbColumns.Checked = True
        Me.rbColumns.Location = New System.Drawing.Point(12, 19)
        Me.rbColumns.Name = "rbColumns"
        Me.rbColumns.Size = New System.Drawing.Size(65, 17)
        Me.rbColumns.TabIndex = 3
        Me.rbColumns.TabStop = True
        Me.rbColumns.Text = "Columns"
        Me.rbColumns.UseVisualStyleBackColor = True
        '
        'frmSelectTables
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(574, 464)
        Me.ControlBox = False
        Me.Controls.Add(Me.gbColRow)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.gbAll)
        Me.Controls.Add(Me.lbxAnalytes)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmSelectTables"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Select/De-Select All"
        Me.gbAll.ResumeLayout(False)
        Me.gbAll.PerformLayout()
        Me.gbColRow.ResumeLayout(False)
        Me.gbColRow.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lbxAnalytes As System.Windows.Forms.ListBox
    Friend WithEvents gbAll As System.Windows.Forms.GroupBox
    Friend WithEvents rbDeselect As System.Windows.Forms.RadioButton
    Friend WithEvents rbSelect As System.Windows.Forms.RadioButton
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents gbColRow As System.Windows.Forms.GroupBox
    Friend WithEvents rbRows As System.Windows.Forms.RadioButton
    Friend WithEvents rbColumns As System.Windows.Forms.RadioButton
End Class
