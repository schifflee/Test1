<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLDAP
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
        Me.panExisting = New System.Windows.Forms.Panel()
        Me.lblExisting = New System.Windows.Forms.Label()
        Me.lbxLDAP = New System.Windows.Forms.ListBox()
        Me.cmdOK1 = New System.Windows.Forms.Button()
        Me.cmdExit1 = New System.Windows.Forms.Button()
        Me.panTest = New System.Windows.Forms.Panel()
        Me.lblUse = New System.Windows.Forms.Label()
        Me.txtFilter = New System.Windows.Forms.TextBox()
        Me.lblAdmin = New System.Windows.Forms.Label()
        Me.lblFilter = New System.Windows.Forms.Label()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdRetrieve = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.chkSaveCreds = New System.Windows.Forms.CheckBox()
        Me.txtUserID = New System.Windows.Forms.TextBox()
        Me.lblUserID = New System.Windows.Forms.Label()
        Me.lblPswd = New System.Windows.Forms.Label()
        Me.txtPswd = New System.Windows.Forms.TextBox()
        Me.chkShowPswd = New System.Windows.Forms.CheckBox()
        Me.lblNetworkAccounts = New System.Windows.Forms.Label()
        Me.dgvUsers = New System.Windows.Forms.DataGridView()
        Me.txtStatus = New System.Windows.Forms.RichTextBox()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.lblLDAP = New System.Windows.Forms.Label()
        Me.txtLDAP = New System.Windows.Forms.TextBox()
        Me.panExisting.SuspendLayout()
        Me.panTest.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.dgvUsers, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'panExisting
        '
        Me.panExisting.Controls.Add(Me.lblExisting)
        Me.panExisting.Controls.Add(Me.lbxLDAP)
        Me.panExisting.Controls.Add(Me.cmdOK1)
        Me.panExisting.Controls.Add(Me.cmdExit1)
        Me.panExisting.Location = New System.Drawing.Point(12, 13)
        Me.panExisting.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panExisting.Name = "panExisting"
        Me.panExisting.Size = New System.Drawing.Size(388, 253)
        Me.panExisting.TabIndex = 0
        '
        'lblExisting
        '
        Me.lblExisting.AutoSize = True
        Me.lblExisting.Location = New System.Drawing.Point(3, 11)
        Me.lblExisting.Name = "lblExisting"
        Me.lblExisting.Size = New System.Drawing.Size(255, 17)
        Me.lblExisting.TabIndex = 137
        Me.lblExisting.Text = "Pick an existing LDAP address from the list"
        '
        'lbxLDAP
        '
        Me.lbxLDAP.FormattingEnabled = True
        Me.lbxLDAP.ItemHeight = 17
        Me.lbxLDAP.Location = New System.Drawing.Point(4, 34)
        Me.lbxLDAP.Name = "lbxLDAP"
        Me.lbxLDAP.Size = New System.Drawing.Size(350, 174)
        Me.lbxLDAP.TabIndex = 136
        '
        'cmdOK1
        '
        Me.cmdOK1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdOK1.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOK1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK1.ForeColor = System.Drawing.Color.Blue
        Me.cmdOK1.Location = New System.Drawing.Point(4, 214)
        Me.cmdOK1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdOK1.Name = "cmdOK1"
        Me.cmdOK1.Size = New System.Drawing.Size(105, 33)
        Me.cmdOK1.TabIndex = 135
        Me.cmdOK1.Text = "&OK"
        Me.cmdOK1.UseVisualStyleBackColor = False
        '
        'cmdExit1
        '
        Me.cmdExit1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdExit1.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdExit1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit1.ForeColor = System.Drawing.Color.Firebrick
        Me.cmdExit1.Location = New System.Drawing.Point(130, 214)
        Me.cmdExit1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdExit1.Name = "cmdExit1"
        Me.cmdExit1.Size = New System.Drawing.Size(105, 33)
        Me.cmdExit1.TabIndex = 134
        Me.cmdExit1.Text = "&Cancel"
        Me.cmdExit1.UseVisualStyleBackColor = False
        '
        'panTest
        '
        Me.panTest.Controls.Add(Me.lblUse)
        Me.panTest.Controls.Add(Me.txtFilter)
        Me.panTest.Controls.Add(Me.lblAdmin)
        Me.panTest.Controls.Add(Me.lblFilter)
        Me.panTest.Controls.Add(Me.cmdOK)
        Me.panTest.Controls.Add(Me.cmdRetrieve)
        Me.panTest.Controls.Add(Me.cmdExit)
        Me.panTest.Controls.Add(Me.Panel1)
        Me.panTest.Controls.Add(Me.lblNetworkAccounts)
        Me.panTest.Controls.Add(Me.dgvUsers)
        Me.panTest.Controls.Add(Me.txtStatus)
        Me.panTest.Controls.Add(Me.lblStatus)
        Me.panTest.Controls.Add(Me.lblLDAP)
        Me.panTest.Controls.Add(Me.txtLDAP)
        Me.panTest.Location = New System.Drawing.Point(421, 13)
        Me.panTest.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panTest.Name = "panTest"
        Me.panTest.Size = New System.Drawing.Size(586, 670)
        Me.panTest.TabIndex = 1
        '
        'lblUse
        '
        Me.lblUse.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUse.ForeColor = System.Drawing.Color.Chocolate
        Me.lblUse.Location = New System.Drawing.Point(17, 602)
        Me.lblUse.Name = "lblUse"
        Me.lblUse.Size = New System.Drawing.Size(550, 27)
        Me.lblUse.TabIndex = 139
        Me.lblUse.Text = "Choose a Network Account, then click OK"
        '
        'txtFilter
        '
        Me.txtFilter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFilter.Location = New System.Drawing.Point(297, 385)
        Me.txtFilter.Name = "txtFilter"
        Me.txtFilter.Size = New System.Drawing.Size(274, 25)
        Me.txtFilter.TabIndex = 138
        '
        'lblAdmin
        '
        Me.lblAdmin.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAdmin.ForeColor = System.Drawing.Color.Chocolate
        Me.lblAdmin.Location = New System.Drawing.Point(143, 102)
        Me.lblAdmin.Name = "lblAdmin"
        Me.lblAdmin.Size = New System.Drawing.Size(429, 75)
        Me.lblAdmin.TabIndex = 134
        Me.lblAdmin.Text = "If 'Retrieve Users' results in an error in the Status box below:" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "    -  enter Ne" & _
    "twork Admin Credentials below" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "       or" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "    -  modify LDAP Server address abov" & _
    "e"
        '
        'lblFilter
        '
        Me.lblFilter.AutoSize = True
        Me.lblFilter.Location = New System.Drawing.Point(168, 393)
        Me.lblFilter.Name = "lblFilter"
        Me.lblFilter.Size = New System.Drawing.Size(123, 17)
        Me.lblFilter.TabIndex = 137
        Me.lblFilter.Text = "Filter by Last Name:"
        '
        'cmdOK
        '
        Me.cmdOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdOK.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.Blue
        Me.cmdOK.Location = New System.Drawing.Point(21, 633)
        Me.cmdOK.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(105, 33)
        Me.cmdOK.TabIndex = 133
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'cmdRetrieve
        '
        Me.cmdRetrieve.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdRetrieve.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRetrieve.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdRetrieve.Location = New System.Drawing.Point(21, 102)
        Me.cmdRetrieve.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdRetrieve.Name = "cmdRetrieve"
        Me.cmdRetrieve.Size = New System.Drawing.Size(113, 50)
        Me.cmdRetrieve.TabIndex = 129
        Me.cmdRetrieve.Text = "&Retrieve Users"
        Me.cmdRetrieve.UseVisualStyleBackColor = False
        '
        'cmdExit
        '
        Me.cmdExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdExit.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ForeColor = System.Drawing.Color.Firebrick
        Me.cmdExit.Location = New System.Drawing.Point(143, 633)
        Me.cmdExit.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(105, 33)
        Me.cmdExit.TabIndex = 132
        Me.cmdExit.Text = "&Cancel"
        Me.cmdExit.UseVisualStyleBackColor = False
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.chkSaveCreds)
        Me.Panel1.Controls.Add(Me.txtUserID)
        Me.Panel1.Controls.Add(Me.lblUserID)
        Me.Panel1.Controls.Add(Me.lblPswd)
        Me.Panel1.Controls.Add(Me.txtPswd)
        Me.Panel1.Controls.Add(Me.chkShowPswd)
        Me.Panel1.Location = New System.Drawing.Point(20, 184)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(551, 103)
        Me.Panel1.TabIndex = 133
        '
        'chkSaveCreds
        '
        Me.chkSaveCreds.Location = New System.Drawing.Point(257, 61)
        Me.chkSaveCreds.Name = "chkSaveCreds"
        Me.chkSaveCreds.Size = New System.Drawing.Size(223, 39)
        Me.chkSaveCreds.TabIndex = 135
        Me.chkSaveCreds.Text = "Save credentials while Administration window is open"
        Me.chkSaveCreds.UseVisualStyleBackColor = True
        '
        'txtUserID
        '
        Me.txtUserID.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUserID.Location = New System.Drawing.Point(123, 2)
        Me.txtUserID.Name = "txtUserID"
        Me.txtUserID.Size = New System.Drawing.Size(334, 25)
        Me.txtUserID.TabIndex = 1
        '
        'lblUserID
        '
        Me.lblUserID.AutoSize = True
        Me.lblUserID.Location = New System.Drawing.Point(0, 5)
        Me.lblUserID.Name = "lblUserID"
        Me.lblUserID.Size = New System.Drawing.Size(117, 17)
        Me.lblUserID.TabIndex = 0
        Me.lblUserID.Text = "Network Admin ID:"
        '
        'lblPswd
        '
        Me.lblPswd.AutoSize = True
        Me.lblPswd.Location = New System.Drawing.Point(0, 33)
        Me.lblPswd.Name = "lblPswd"
        Me.lblPswd.Size = New System.Drawing.Size(120, 17)
        Me.lblPswd.TabIndex = 1
        Me.lblPswd.Text = "Network Password:"
        '
        'txtPswd
        '
        Me.txtPswd.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPswd.Location = New System.Drawing.Point(123, 30)
        Me.txtPswd.Name = "txtPswd"
        Me.txtPswd.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPswd.Size = New System.Drawing.Size(334, 25)
        Me.txtPswd.TabIndex = 2
        '
        'chkShowPswd
        '
        Me.chkShowPswd.AutoSize = True
        Me.chkShowPswd.Location = New System.Drawing.Point(123, 70)
        Me.chkShowPswd.Name = "chkShowPswd"
        Me.chkShowPswd.Size = New System.Drawing.Size(118, 21)
        Me.chkShowPswd.TabIndex = 132
        Me.chkShowPswd.Text = "Show Password"
        Me.chkShowPswd.UseVisualStyleBackColor = True
        '
        'lblNetworkAccounts
        '
        Me.lblNetworkAccounts.AutoSize = True
        Me.lblNetworkAccounts.Location = New System.Drawing.Point(18, 393)
        Me.lblNetworkAccounts.Name = "lblNetworkAccounts"
        Me.lblNetworkAccounts.Size = New System.Drawing.Size(116, 17)
        Me.lblNetworkAccounts.TabIndex = 136
        Me.lblNetworkAccounts.Text = "Network Accounts:"
        '
        'dgvUsers
        '
        Me.dgvUsers.AllowUserToAddRows = False
        Me.dgvUsers.AllowUserToDeleteRows = False
        Me.dgvUsers.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvUsers.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvUsers.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvUsers.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvUsers.Location = New System.Drawing.Point(20, 413)
        Me.dgvUsers.MultiSelect = False
        Me.dgvUsers.Name = "dgvUsers"
        Me.dgvUsers.ReadOnly = True
        Me.dgvUsers.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvUsers.Size = New System.Drawing.Size(551, 176)
        Me.dgvUsers.TabIndex = 135
        '
        'txtStatus
        '
        Me.txtStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtStatus.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtStatus.Location = New System.Drawing.Point(20, 310)
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.Size = New System.Drawing.Size(552, 63)
        Me.txtStatus.TabIndex = 133
        Me.txtStatus.Text = ""
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(17, 290)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(174, 17)
        Me.lblStatus.TabIndex = 131
        Me.lblStatus.Text = "LDAP Communication Status:"
        '
        'lblLDAP
        '
        Me.lblLDAP.AutoSize = True
        Me.lblLDAP.Location = New System.Drawing.Point(17, 11)
        Me.lblLDAP.Name = "lblLDAP"
        Me.lblLDAP.Size = New System.Drawing.Size(133, 17)
        Me.lblLDAP.TabIndex = 5
        Me.lblLDAP.Text = "LDAP Server address:"
        '
        'txtLDAP
        '
        Me.txtLDAP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLDAP.Location = New System.Drawing.Point(20, 31)
        Me.txtLDAP.Multiline = True
        Me.txtLDAP.Name = "txtLDAP"
        Me.txtLDAP.Size = New System.Drawing.Size(552, 66)
        Me.txtLDAP.TabIndex = 0
        '
        'frmLDAP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1019, 708)
        Me.ControlBox = False
        Me.Controls.Add(Me.panTest)
        Me.Controls.Add(Me.panExisting)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "frmLDAP"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "LDAP Actions..."
        Me.panExisting.ResumeLayout(False)
        Me.panExisting.PerformLayout()
        Me.panTest.ResumeLayout(False)
        Me.panTest.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.dgvUsers, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents panExisting As System.Windows.Forms.Panel
    Friend WithEvents panTest As System.Windows.Forms.Panel
    Friend WithEvents lblLDAP As System.Windows.Forms.Label
    Friend WithEvents txtLDAP As System.Windows.Forms.TextBox
    Friend WithEvents txtPswd As System.Windows.Forms.TextBox
    Friend WithEvents txtUserID As System.Windows.Forms.TextBox
    Friend WithEvents lblPswd As System.Windows.Forms.Label
    Friend WithEvents lblUserID As System.Windows.Forms.Label
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents cmdRetrieve As System.Windows.Forms.Button
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents chkShowPswd As System.Windows.Forms.CheckBox
    Friend WithEvents txtStatus As System.Windows.Forms.RichTextBox
    Friend WithEvents lblAdmin As System.Windows.Forms.Label
    Friend WithEvents lblNetworkAccounts As System.Windows.Forms.Label
    Friend WithEvents dgvUsers As System.Windows.Forms.DataGridView
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdOK1 As System.Windows.Forms.Button
    Friend WithEvents cmdExit1 As System.Windows.Forms.Button
    Friend WithEvents txtFilter As System.Windows.Forms.TextBox
    Friend WithEvents lblFilter As System.Windows.Forms.Label
    Friend WithEvents lblExisting As System.Windows.Forms.Label
    Friend WithEvents lbxLDAP As System.Windows.Forms.ListBox
    Friend WithEvents chkSaveCreds As System.Windows.Forms.CheckBox
    Friend WithEvents lblUse As System.Windows.Forms.Label
End Class
