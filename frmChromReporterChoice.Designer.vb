<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChromReporterChoice
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
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.lblRDB = New System.Windows.Forms.Label()
        Me.txtRDB = New System.Windows.Forms.TextBox()
        Me.cmdBrowseRDB = New System.Windows.Forms.Button()
        Me.cmdBrowseDirectory = New System.Windows.Forms.Button()
        Me.txtDestinationPath = New System.Windows.Forms.TextBox()
        Me.lblDestinationPath = New System.Windows.Forms.Label()
        Me.txtWordFileName = New System.Windows.Forms.TextBox()
        Me.lblEnterWordName = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOK.CausesValidation = False
        Me.cmdOK.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.cmdOK.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.Blue
        Me.cmdOK.Location = New System.Drawing.Point(12, 359)
        Me.cmdOK.Margin = New System.Windows.Forms.Padding(3, 5, 3, 5)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(92, 43)
        Me.cmdOK.TabIndex = 3
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancel.CausesValidation = False
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.cmdCancel.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(131, 359)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(3, 5, 3, 5)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(92, 43)
        Me.cmdCancel.TabIndex = 4
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'lblRDB
        '
        Me.lblRDB.AutoSize = True
        Me.lblRDB.Location = New System.Drawing.Point(9, 62)
        Me.lblRDB.Name = "lblRDB"
        Me.lblRDB.Size = New System.Drawing.Size(248, 17)
        Me.lblRDB.TabIndex = 5
        Me.lblRDB.Text = "1.  Browse to a Sciex Analyst TM .rdb file:"
        '
        'txtRDB
        '
        Me.txtRDB.Location = New System.Drawing.Point(12, 82)
        Me.txtRDB.Multiline = True
        Me.txtRDB.Name = "txtRDB"
        Me.txtRDB.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtRDB.Size = New System.Drawing.Size(652, 47)
        Me.txtRDB.TabIndex = 0
        '
        'cmdBrowseRDB
        '
        Me.cmdBrowseRDB.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdBrowseRDB.CausesValidation = False
        Me.cmdBrowseRDB.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.cmdBrowseRDB.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseRDB.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.cmdBrowseRDB.Location = New System.Drawing.Point(263, 36)
        Me.cmdBrowseRDB.Margin = New System.Windows.Forms.Padding(3, 5, 3, 5)
        Me.cmdBrowseRDB.Name = "cmdBrowseRDB"
        Me.cmdBrowseRDB.Size = New System.Drawing.Size(104, 43)
        Me.cmdBrowseRDB.TabIndex = 7
        Me.cmdBrowseRDB.TabStop = False
        Me.cmdBrowseRDB.Text = "&Browse to .rdb..."
        Me.cmdBrowseRDB.UseVisualStyleBackColor = False
        '
        'cmdBrowseDirectory
        '
        Me.cmdBrowseDirectory.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdBrowseDirectory.CausesValidation = False
        Me.cmdBrowseDirectory.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.cmdBrowseDirectory.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseDirectory.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.cmdBrowseDirectory.Location = New System.Drawing.Point(269, 258)
        Me.cmdBrowseDirectory.Margin = New System.Windows.Forms.Padding(3, 5, 3, 5)
        Me.cmdBrowseDirectory.Name = "cmdBrowseDirectory"
        Me.cmdBrowseDirectory.Size = New System.Drawing.Size(104, 43)
        Me.cmdBrowseDirectory.TabIndex = 10
        Me.cmdBrowseDirectory.TabStop = False
        Me.cmdBrowseDirectory.Text = "&Browse to directory..."
        Me.cmdBrowseDirectory.UseVisualStyleBackColor = False
        '
        'txtDestinationPath
        '
        Me.txtDestinationPath.Location = New System.Drawing.Point(12, 304)
        Me.txtDestinationPath.Multiline = True
        Me.txtDestinationPath.Name = "txtDestinationPath"
        Me.txtDestinationPath.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtDestinationPath.Size = New System.Drawing.Size(652, 47)
        Me.txtDestinationPath.TabIndex = 2
        '
        'lblDestinationPath
        '
        Me.lblDestinationPath.AutoSize = True
        Me.lblDestinationPath.Location = New System.Drawing.Point(9, 284)
        Me.lblDestinationPath.Name = "lblDestinationPath"
        Me.lblDestinationPath.Size = New System.Drawing.Size(254, 17)
        Me.lblDestinationPath.TabIndex = 8
        Me.lblDestinationPath.Text = "3.  Browse to a Word file destination path:"
        '
        'txtWordFileName
        '
        Me.txtWordFileName.Location = New System.Drawing.Point(12, 185)
        Me.txtWordFileName.Name = "txtWordFileName"
        Me.txtWordFileName.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtWordFileName.Size = New System.Drawing.Size(652, 25)
        Me.txtWordFileName.TabIndex = 1
        '
        'lblEnterWordName
        '
        Me.lblEnterWordName.Location = New System.Drawing.Point(9, 149)
        Me.lblEnterWordName.Name = "lblEnterWordName"
        Me.lblEnterWordName.Size = New System.Drawing.Size(655, 33)
        Me.lblEnterWordName.TabIndex = 11
        Me.lblEnterWordName.Text = "2.  Enter a name for the Word file that will be generated" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(cannot contain specia" & _
    "l characters or spaces or periods):"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Gainsboro
        Me.Button1.CausesValidation = False
        Me.Button1.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.Button1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.Button1.Location = New System.Drawing.Point(560, 14)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 5, 3, 5)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(104, 43)
        Me.Button1.TabIndex = 12
        Me.Button1.TabStop = False
        Me.Button1.Text = "Test"
        Me.Button1.UseVisualStyleBackColor = False
        Me.Button1.Visible = False
        '
        'frmChromReporterChoice
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(682, 435)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.txtWordFileName)
        Me.Controls.Add(Me.lblEnterWordName)
        Me.Controls.Add(Me.cmdBrowseDirectory)
        Me.Controls.Add(Me.txtDestinationPath)
        Me.Controls.Add(Me.lblDestinationPath)
        Me.Controls.Add(Me.cmdBrowseRDB)
        Me.Controls.Add(Me.txtRDB)
        Me.Controls.Add(Me.lblRDB)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "frmChromReporterChoice"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ChromReporter Configuration..."
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents lblRDB As System.Windows.Forms.Label
    Friend WithEvents txtRDB As System.Windows.Forms.TextBox
    Friend WithEvents cmdBrowseRDB As System.Windows.Forms.Button
    Friend WithEvents cmdBrowseDirectory As System.Windows.Forms.Button
    Friend WithEvents txtDestinationPath As System.Windows.Forms.TextBox
    Friend WithEvents lblDestinationPath As System.Windows.Forms.Label
    Friend WithEvents txtWordFileName As System.Windows.Forms.TextBox
    Friend WithEvents lblEnterWordName As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
