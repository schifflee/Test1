<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAuditTrailPrintChoice
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAuditTrailPrintChoice))
        Me.gbChoice = New System.Windows.Forms.GroupBox()
        Me.rbShort = New System.Windows.Forms.RadioButton()
        Me.rbLong = New System.Windows.Forms.RadioButton()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.gbChoice.SuspendLayout()
        Me.SuspendLayout()
        '
        'gbChoice
        '
        Me.gbChoice.Controls.Add(Me.rbShort)
        Me.gbChoice.Controls.Add(Me.rbLong)
        Me.gbChoice.Location = New System.Drawing.Point(34, 20)
        Me.gbChoice.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbChoice.Name = "gbChoice"
        Me.gbChoice.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbChoice.Size = New System.Drawing.Size(218, 131)
        Me.gbChoice.TabIndex = 0
        Me.gbChoice.TabStop = False
        '
        'rbShort
        '
        Me.rbShort.AutoSize = True
        Me.rbShort.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbShort.Location = New System.Drawing.Point(20, 77)
        Me.rbShort.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbShort.Name = "rbShort"
        Me.rbShort.Size = New System.Drawing.Size(160, 24)
        Me.rbShort.TabIndex = 2
        Me.rbShort.TabStop = True
        Me.rbShort.Text = "Print Short Version"
        Me.rbShort.UseVisualStyleBackColor = True
        '
        'rbLong
        '
        Me.rbLong.AutoSize = True
        Me.rbLong.Checked = True
        Me.rbLong.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbLong.Location = New System.Drawing.Point(20, 25)
        Me.rbLong.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbLong.Name = "rbLong"
        Me.rbLong.Size = New System.Drawing.Size(157, 24)
        Me.rbLong.TabIndex = 1
        Me.rbLong.TabStop = True
        Me.rbLong.Text = "Print Long Version"
        Me.rbLong.UseVisualStyleBackColor = True
        '
        'cmdOK
        '
        Me.cmdOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.Blue
        Me.cmdOK.Location = New System.Drawing.Point(34, 174)
        Me.cmdOK.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(85, 48)
        Me.cmdOK.TabIndex = 1
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(167, 174)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(85, 48)
        Me.cmdCancel.TabIndex = 2
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'frmAuditTrailPrintChoice
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(289, 255)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.gbChoice)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "frmAuditTrailPrintChoice"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Print..."
        Me.gbChoice.ResumeLayout(False)
        Me.gbChoice.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbChoice As System.Windows.Forms.GroupBox
    Friend WithEvents rbShort As System.Windows.Forms.RadioButton
    Friend WithEvents rbLong As System.Windows.Forms.RadioButton
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
End Class
