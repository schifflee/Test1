<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAddUser
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAddUser))
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.lblUserID = New System.Windows.Forms.Label()
        Me.txtUserID = New System.Windows.Forms.TextBox()
        Me.txtPswd = New System.Windows.Forms.TextBox()
        Me.lblPswd = New System.Windows.Forms.Label()
        Me.panUserID = New System.Windows.Forms.Panel()
        Me.txtConfirm = New System.Windows.Forms.TextBox()
        Me.lblConfirm = New System.Windows.Forms.Label()
        Me.panUserName = New System.Windows.Forms.Panel()
        Me.lblLastName = New System.Windows.Forms.Label()
        Me.txtLastName = New System.Windows.Forms.TextBox()
        Me.lblMiddleName = New System.Windows.Forms.Label()
        Me.txtMiddleName = New System.Windows.Forms.TextBox()
        Me.lblFirstName = New System.Windows.Forms.Label()
        Me.txtFirstName = New System.Windows.Forms.TextBox()
        Me.panUserID.SuspendLayout()
        Me.panUserName.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(259, 149)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(76, 33)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOK.CausesValidation = False
        Me.cmdOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.Blue
        Me.cmdOK.Location = New System.Drawing.Point(174, 149)
        Me.cmdOK.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(76, 33)
        Me.cmdOK.TabIndex = 2
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'lblUserID
        '
        Me.lblUserID.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUserID.Location = New System.Drawing.Point(1, 4)
        Me.lblUserID.Name = "lblUserID"
        Me.lblUserID.Size = New System.Drawing.Size(152, 26)
        Me.lblUserID.TabIndex = 96
        Me.lblUserID.Text = "Enter New UserID:"
        '
        'txtUserID
        '
        Me.txtUserID.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.txtUserID.Location = New System.Drawing.Point(160, 4)
        Me.txtUserID.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtUserID.Name = "txtUserID"
        Me.txtUserID.Size = New System.Drawing.Size(271, 25)
        Me.txtUserID.TabIndex = 0
        '
        'txtPswd
        '
        Me.txtPswd.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.txtPswd.Location = New System.Drawing.Point(160, 41)
        Me.txtPswd.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtPswd.Name = "txtPswd"
        Me.txtPswd.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPswd.Size = New System.Drawing.Size(271, 25)
        Me.txtPswd.TabIndex = 1
        '
        'lblPswd
        '
        Me.lblPswd.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPswd.Location = New System.Drawing.Point(1, 41)
        Me.lblPswd.Name = "lblPswd"
        Me.lblPswd.Size = New System.Drawing.Size(131, 26)
        Me.lblPswd.TabIndex = 98
        Me.lblPswd.Text = "Enter Password:"
        '
        'panUserID
        '
        Me.panUserID.Controls.Add(Me.txtConfirm)
        Me.panUserID.Controls.Add(Me.lblConfirm)
        Me.panUserID.Controls.Add(Me.txtPswd)
        Me.panUserID.Controls.Add(Me.lblUserID)
        Me.panUserID.Controls.Add(Me.lblPswd)
        Me.panUserID.Controls.Add(Me.txtUserID)
        Me.panUserID.Location = New System.Drawing.Point(14, 16)
        Me.panUserID.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panUserID.Name = "panUserID"
        Me.panUserID.Size = New System.Drawing.Size(450, 126)
        Me.panUserID.TabIndex = 0
        Me.panUserID.TabStop = True
        '
        'txtConfirm
        '
        Me.txtConfirm.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.txtConfirm.Location = New System.Drawing.Point(160, 77)
        Me.txtConfirm.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtConfirm.Name = "txtConfirm"
        Me.txtConfirm.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtConfirm.Size = New System.Drawing.Size(271, 25)
        Me.txtConfirm.TabIndex = 2
        '
        'lblConfirm
        '
        Me.lblConfirm.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblConfirm.Location = New System.Drawing.Point(1, 77)
        Me.lblConfirm.Name = "lblConfirm"
        Me.lblConfirm.Size = New System.Drawing.Size(152, 26)
        Me.lblConfirm.TabIndex = 100
        Me.lblConfirm.Text = "Confirm Password:"
        '
        'panUserName
        '
        Me.panUserName.Controls.Add(Me.lblLastName)
        Me.panUserName.Controls.Add(Me.txtLastName)
        Me.panUserName.Controls.Add(Me.lblMiddleName)
        Me.panUserName.Controls.Add(Me.txtMiddleName)
        Me.panUserName.Controls.Add(Me.lblFirstName)
        Me.panUserName.Controls.Add(Me.txtFirstName)
        Me.panUserName.Location = New System.Drawing.Point(14, 235)
        Me.panUserName.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panUserName.Name = "panUserName"
        Me.panUserName.Size = New System.Drawing.Size(450, 118)
        Me.panUserName.TabIndex = 1
        Me.panUserName.TabStop = True
        '
        'lblLastName
        '
        Me.lblLastName.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.lblLastName.Location = New System.Drawing.Point(1, 77)
        Me.lblLastName.Name = "lblLastName"
        Me.lblLastName.Size = New System.Drawing.Size(131, 26)
        Me.lblLastName.TabIndex = 102
        Me.lblLastName.Text = "Last Name:"
        '
        'txtLastName
        '
        Me.txtLastName.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.txtLastName.Location = New System.Drawing.Point(160, 77)
        Me.txtLastName.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtLastName.Name = "txtLastName"
        Me.txtLastName.Size = New System.Drawing.Size(271, 25)
        Me.txtLastName.TabIndex = 2
        '
        'lblMiddleName
        '
        Me.lblMiddleName.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.lblMiddleName.Location = New System.Drawing.Point(1, 41)
        Me.lblMiddleName.Name = "lblMiddleName"
        Me.lblMiddleName.Size = New System.Drawing.Size(131, 26)
        Me.lblMiddleName.TabIndex = 100
        Me.lblMiddleName.Text = "Middle Name:"
        '
        'txtMiddleName
        '
        Me.txtMiddleName.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.txtMiddleName.Location = New System.Drawing.Point(160, 41)
        Me.txtMiddleName.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtMiddleName.Name = "txtMiddleName"
        Me.txtMiddleName.Size = New System.Drawing.Size(271, 25)
        Me.txtMiddleName.TabIndex = 1
        '
        'lblFirstName
        '
        Me.lblFirstName.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.lblFirstName.Location = New System.Drawing.Point(1, 4)
        Me.lblFirstName.Name = "lblFirstName"
        Me.lblFirstName.Size = New System.Drawing.Size(131, 26)
        Me.lblFirstName.TabIndex = 98
        Me.lblFirstName.Text = "First Name:"
        '
        'txtFirstName
        '
        Me.txtFirstName.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.txtFirstName.Location = New System.Drawing.Point(160, 4)
        Me.txtFirstName.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtFirstName.Name = "txtFirstName"
        Me.txtFirstName.Size = New System.Drawing.Size(271, 25)
        Me.txtFirstName.TabIndex = 0
        '
        'frmAddUser
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(478, 400)
        Me.ControlBox = False
        Me.Controls.Add(Me.panUserName)
        Me.Controls.Add(Me.panUserID)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "frmAddUser"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "frmAddUser"
        Me.panUserID.ResumeLayout(False)
        Me.panUserID.PerformLayout()
        Me.panUserName.ResumeLayout(False)
        Me.panUserName.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents lblUserID As System.Windows.Forms.Label
    Friend WithEvents txtUserID As System.Windows.Forms.TextBox
    Friend WithEvents txtPswd As System.Windows.Forms.TextBox
    Friend WithEvents lblPswd As System.Windows.Forms.Label
    Friend WithEvents panUserID As System.Windows.Forms.Panel
    Friend WithEvents panUserName As System.Windows.Forms.Panel
    Friend WithEvents lblLastName As System.Windows.Forms.Label
    Friend WithEvents txtLastName As System.Windows.Forms.TextBox
    Friend WithEvents lblMiddleName As System.Windows.Forms.Label
    Friend WithEvents txtMiddleName As System.Windows.Forms.TextBox
    Friend WithEvents lblFirstName As System.Windows.Forms.Label
    Friend WithEvents txtFirstName As System.Windows.Forms.TextBox
    Friend WithEvents txtConfirm As System.Windows.Forms.TextBox
    Friend WithEvents lblConfirm As System.Windows.Forms.Label
End Class
