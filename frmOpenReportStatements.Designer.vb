<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOpenReportStatements
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOpenReportStatements))
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rbGuWu = New System.Windows.Forms.RadioButton()
        Me.rbClient = New System.Windows.Forms.RadioButton()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.rbEditMode = New System.Windows.Forms.RadioButton()
        Me.rbReadOnly = New System.Windows.Forms.RadioButton()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancel.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(225, 109)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(81, 30)
        Me.cmdCancel.TabIndex = 114
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.Blue
        Me.cmdOK.Location = New System.Drawing.Point(16, 109)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(81, 30)
        Me.cmdOK.TabIndex = 113
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rbGuWu)
        Me.GroupBox1.Controls.Add(Me.rbClient)
        Me.GroupBox1.Location = New System.Drawing.Point(87, 47)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(150, 61)
        Me.GroupBox1.TabIndex = 115
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Choose Statement Document..."
        Me.GroupBox1.Visible = False
        '
        'rbGuWu
        '
        Me.rbGuWu.AutoSize = True
        Me.rbGuWu.Checked = True
        Me.rbGuWu.Location = New System.Drawing.Point(22, 42)
        Me.rbGuWu.Name = "rbGuWu"
        Me.rbGuWu.Size = New System.Drawing.Size(147, 17)
        Me.rbGuWu.TabIndex = 1
        Me.rbGuWu.TabStop = True
        Me.rbGuWu.Text = "GuWu Report Statements"
        Me.rbGuWu.UseVisualStyleBackColor = True
        '
        'rbClient
        '
        Me.rbClient.AutoSize = True
        Me.rbClient.Location = New System.Drawing.Point(22, 19)
        Me.rbClient.Name = "rbClient"
        Me.rbClient.Size = New System.Drawing.Size(169, 17)
        Me.rbClient.TabIndex = 0
        Me.rbClient.Text = "User/Client Report Statements"
        Me.rbClient.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.rbEditMode)
        Me.GroupBox2.Controls.Add(Me.GroupBox1)
        Me.GroupBox2.Controls.Add(Me.rbReadOnly)
        Me.GroupBox2.Location = New System.Drawing.Point(16, 12)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(290, 79)
        Me.GroupBox2.TabIndex = 116
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Edit Mode or Read-Only..."
        '
        'rbEditMode
        '
        Me.rbEditMode.AutoSize = True
        Me.rbEditMode.Location = New System.Drawing.Point(22, 47)
        Me.rbEditMode.Name = "rbEditMode"
        Me.rbEditMode.Size = New System.Drawing.Size(73, 17)
        Me.rbEditMode.TabIndex = 2
        Me.rbEditMode.Text = "Edit Mode"
        Me.rbEditMode.UseVisualStyleBackColor = True
        '
        'rbReadOnly
        '
        Me.rbReadOnly.AutoSize = True
        Me.rbReadOnly.Checked = True
        Me.rbReadOnly.Location = New System.Drawing.Point(22, 19)
        Me.rbReadOnly.Name = "rbReadOnly"
        Me.rbReadOnly.Size = New System.Drawing.Size(105, 17)
        Me.rbReadOnly.TabIndex = 1
        Me.rbReadOnly.TabStop = True
        Me.rbReadOnly.Text = "Read-Only Mode"
        Me.rbReadOnly.UseVisualStyleBackColor = True
        '
        'frmOpenReportStatements
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(333, 159)
        Me.ControlBox = False
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmOpenReportStatements"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "   Open Report Statements..."
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents rbGuWu As System.Windows.Forms.RadioButton
    Friend WithEvents rbClient As System.Windows.Forms.RadioButton
    Friend WithEvents rbEditMode As System.Windows.Forms.RadioButton
    Friend WithEvents rbReadOnly As System.Windows.Forms.RadioButton
End Class
