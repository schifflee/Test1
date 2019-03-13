<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAutoAssignBegin
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAutoAssignBegin))
        Me.lbl1 = New System.Windows.Forms.Label()
        Me.gbChoice = New System.Windows.Forms.GroupBox()
        Me.rbSelectedTable = New System.Windows.Forms.RadioButton()
        Me.rbOnlyEmpty = New System.Windows.Forms.RadioButton()
        Me.rbOverwrite = New System.Windows.Forms.RadioButton()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.gbChoice.SuspendLayout()
        Me.SuspendLayout()
        '
        'lbl1
        '
        Me.lbl1.Location = New System.Drawing.Point(12, 22)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(461, 167)
        Me.lbl1.TabIndex = 0
        Me.lbl1.Text = resources.GetString("lbl1.Text")
        Me.lbl1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'gbChoice
        '
        Me.gbChoice.Controls.Add(Me.rbSelectedTable)
        Me.gbChoice.Controls.Add(Me.rbOnlyEmpty)
        Me.gbChoice.Controls.Add(Me.rbOverwrite)
        Me.gbChoice.Location = New System.Drawing.Point(92, 171)
        Me.gbChoice.Name = "gbChoice"
        Me.gbChoice.Size = New System.Drawing.Size(293, 137)
        Me.gbChoice.TabIndex = 1
        Me.gbChoice.TabStop = False
        '
        'rbSelectedTable
        '
        Me.rbSelectedTable.AutoSize = True
        Me.rbSelectedTable.Checked = True
        Me.rbSelectedTable.Location = New System.Drawing.Point(19, 78)
        Me.rbSelectedTable.Name = "rbSelectedTable"
        Me.rbSelectedTable.Size = New System.Drawing.Size(245, 38)
        Me.rbSelectedTable.TabIndex = 2
        Me.rbSelectedTable.TabStop = True
        Me.rbSelectedTable.Text = "Assign samples only to selected table" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(will overwrite existing assignments)"
        Me.rbSelectedTable.UseVisualStyleBackColor = True
        '
        'rbOnlyEmpty
        '
        Me.rbOnlyEmpty.AutoSize = True
        Me.rbOnlyEmpty.Location = New System.Drawing.Point(19, 24)
        Me.rbOnlyEmpty.Name = "rbOnlyEmpty"
        Me.rbOnlyEmpty.Size = New System.Drawing.Size(242, 21)
        Me.rbOnlyEmpty.TabIndex = 1
        Me.rbOnlyEmpty.Text = "Assign samples only to empty tables."
        Me.rbOnlyEmpty.UseVisualStyleBackColor = True
        '
        'rbOverwrite
        '
        Me.rbOverwrite.AutoSize = True
        Me.rbOverwrite.Location = New System.Drawing.Point(19, 51)
        Me.rbOverwrite.Name = "rbOverwrite"
        Me.rbOverwrite.Size = New System.Drawing.Size(258, 21)
        Me.rbOverwrite.TabIndex = 0
        Me.rbOverwrite.Text = "Overwrite all existing assigned samples."
        Me.rbOverwrite.UseVisualStyleBackColor = True
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOK.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.cmdOK.FlatAppearance.BorderSize = 0
        Me.cmdOK.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdOK.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.Blue
        Me.cmdOK.Location = New System.Drawing.Point(134, 338)
        Me.cmdOK.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(79, 33)
        Me.cmdOK.TabIndex = 127
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
        Me.cmdCancel.Location = New System.Drawing.Point(254, 338)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(79, 33)
        Me.cmdCancel.TabIndex = 128
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'frmAutoAssignBegin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(491, 405)
        Me.ControlBox = False
        Me.Controls.Add(Me.gbChoice)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.lbl1)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "frmAutoAssignBegin"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Continue?"
        Me.gbChoice.ResumeLayout(False)
        Me.gbChoice.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Friend WithEvents gbChoice As System.Windows.Forms.GroupBox
    Friend WithEvents rbOnlyEmpty As System.Windows.Forms.RadioButton
    Friend WithEvents rbOverwrite As System.Windows.Forms.RadioButton
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents rbSelectedTable As System.Windows.Forms.RadioButton
End Class
