<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmESig
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmESig))
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.txtUserID = New System.Windows.Forms.TextBox()
        Me.lblPassword = New System.Windows.Forms.Label()
        Me.lblUserID = New System.Windows.Forms.Label()
        Me.lblMOS = New System.Windows.Forms.Label()
        Me.cbxMOS = New System.Windows.Forms.ComboBox()
        Me.cbxRFC = New System.Windows.Forms.ComboBox()
        Me.lblRFC = New System.Windows.Forms.Label()
        Me.txtMOS = New System.Windows.Forms.TextBox()
        Me.txtRFC = New System.Windows.Forms.TextBox()
        Me.panMOS = New System.Windows.Forms.Panel()
        Me.panMOSE = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblMOSE = New System.Windows.Forms.Label()
        Me.panCred = New System.Windows.Forms.Panel()
        Me.txtUserName = New System.Windows.Forms.TextBox()
        Me.lblUserName = New System.Windows.Forms.Label()
        Me.panRFC = New System.Windows.Forms.Panel()
        Me.panRFCE = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblRFCE = New System.Windows.Forms.Label()
        Me.panOK = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.chkRFC = New System.Windows.Forms.CheckBox()
        Me.chkMOS = New System.Windows.Forms.CheckBox()
        Me.chkRRFC = New System.Windows.Forms.CheckBox()
        Me.chkRMOS = New System.Windows.Forms.CheckBox()
        Me.lblTest = New System.Windows.Forms.Label()
        Me.pan1 = New System.Windows.Forms.Panel()
        Me.rbESigOn = New System.Windows.Forms.RadioButton()
        Me.panMOS.SuspendLayout
        Me.panMOSE.SuspendLayout
        Me.panCred.SuspendLayout
        Me.panRFC.SuspendLayout
        Me.panRFCE.SuspendLayout
        Me.panOK.SuspendLayout
        Me.pan1.SuspendLayout
        Me.SuspendLayout
        '
        'txtPassword
        '
        Me.txtPassword.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPassword.Font = New System.Drawing.Font("Segoe UI", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.txtPassword.Location = New System.Drawing.Point(220, 75)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(254, 23)
        Me.txtPassword.TabIndex = 2
        Me.txtPassword.UseSystemPasswordChar = True
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancel.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(130, 11)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(97, 39)
        Me.cmdCancel.TabIndex = 9
        Me.cmdCancel.TabStop = False
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOK.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.Blue
        Me.cmdOK.Location = New System.Drawing.Point(12, 11)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(97, 39)
        Me.cmdOK.TabIndex = 8
        Me.cmdOK.TabStop = False
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'txtUserID
        '
        Me.txtUserID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUserID.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUserID.Location = New System.Drawing.Point(220, 9)
        Me.txtUserID.Name = "txtUserID"
        Me.txtUserID.Size = New System.Drawing.Size(254, 23)
        Me.txtUserID.TabIndex = 0
        Me.txtUserID.TabStop = False
        '
        'lblPassword
        '
        Me.lblPassword.AutoSize = True
        Me.lblPassword.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPassword.Location = New System.Drawing.Point(8, 77)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(208, 21)
        Me.lblPassword.TabIndex = 6
        Me.lblPassword.Text = "Enter StudyDoc Password:"
        '
        'lblUserID
        '
        Me.lblUserID.AutoSize = True
        Me.lblUserID.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUserID.Location = New System.Drawing.Point(8, 12)
        Me.lblUserID.Name = "lblUserID"
        Me.lblUserID.Size = New System.Drawing.Size(191, 21)
        Me.lblUserID.TabIndex = 5
        Me.lblUserID.Text = "Enter StudyDoc User ID:"
        '
        'lblMOS
        '
        Me.lblMOS.AutoSize = True
        Me.lblMOS.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMOS.Location = New System.Drawing.Point(8, 6)
        Me.lblMOS.Name = "lblMOS"
        Me.lblMOS.Size = New System.Drawing.Size(180, 21)
        Me.lblMOS.TabIndex = 10
        Me.lblMOS.Text = "Meaning of Signature:"
        '
        'cbxMOS
        '
        Me.cbxMOS.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxMOS.FormattingEnabled = True
        Me.cbxMOS.Location = New System.Drawing.Point(220, 6)
        Me.cbxMOS.Name = "cbxMOS"
        Me.cbxMOS.Size = New System.Drawing.Size(254, 21)
        Me.cbxMOS.TabIndex = 4
        '
        'cbxRFC
        '
        Me.cbxRFC.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxRFC.FormattingEnabled = True
        Me.cbxRFC.Location = New System.Drawing.Point(220, 10)
        Me.cbxRFC.Name = "cbxRFC"
        Me.cbxRFC.Size = New System.Drawing.Size(254, 21)
        Me.cbxRFC.TabIndex = 6
        '
        'lblRFC
        '
        Me.lblRFC.AutoSize = True
        Me.lblRFC.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRFC.Location = New System.Drawing.Point(8, 10)
        Me.lblRFC.Name = "lblRFC"
        Me.lblRFC.Size = New System.Drawing.Size(159, 21)
        Me.lblRFC.TabIndex = 12
        Me.lblRFC.Text = "Reason For Change:"
        '
        'txtMOS
        '
        Me.txtMOS.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMOS.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMOS.Location = New System.Drawing.Point(212, 23)
        Me.txtMOS.Multiline = True
        Me.txtMOS.Name = "txtMOS"
        Me.txtMOS.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtMOS.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtMOS.Size = New System.Drawing.Size(254, 63)
        Me.txtMOS.TabIndex = 5
        Me.txtMOS.UseSystemPasswordChar = True
        '
        'txtRFC
        '
        Me.txtRFC.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtRFC.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRFC.Location = New System.Drawing.Point(212, 23)
        Me.txtRFC.Multiline = True
        Me.txtRFC.Name = "txtRFC"
        Me.txtRFC.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtRFC.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtRFC.Size = New System.Drawing.Size(254, 63)
        Me.txtRFC.TabIndex = 7
        Me.txtRFC.UseSystemPasswordChar = True
        '
        'panMOS
        '
        Me.panMOS.Controls.Add(Me.panMOSE)
        Me.panMOS.Controls.Add(Me.lblMOS)
        Me.panMOS.Controls.Add(Me.cbxMOS)
        Me.panMOS.Location = New System.Drawing.Point(3, 118)
        Me.panMOS.Name = "panMOS"
        Me.panMOS.Size = New System.Drawing.Size(479, 123)
        Me.panMOS.TabIndex = 16
        Me.panMOS.TabStop = True
        '
        'panMOSE
        '
        Me.panMOSE.Controls.Add(Me.Label1)
        Me.panMOSE.Controls.Add(Me.txtMOS)
        Me.panMOSE.Controls.Add(Me.lblMOSE)
        Me.panMOSE.Location = New System.Drawing.Point(8, 28)
        Me.panMOSE.Name = "panMOSE"
        Me.panMOSE.Size = New System.Drawing.Size(471, 89)
        Me.panMOSE.TabIndex = 19
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(209, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(78, 13)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Free Form Text"
        '
        'lblMOSE
        '
        Me.lblMOSE.ForeColor = System.Drawing.Color.Black
        Me.lblMOSE.Location = New System.Drawing.Point(3, 23)
        Me.lblMOSE.Name = "lblMOSE"
        Me.lblMOSE.Size = New System.Drawing.Size(193, 60)
        Me.lblMOSE.TabIndex = 11
        Me.lblMOSE.Text = "lblMOSE"
        '
        'panCred
        '
        Me.panCred.Controls.Add(Me.txtUserName)
        Me.panCred.Controls.Add(Me.lblUserName)
        Me.panCred.Controls.Add(Me.txtPassword)
        Me.panCred.Controls.Add(Me.lblUserID)
        Me.panCred.Controls.Add(Me.lblPassword)
        Me.panCred.Controls.Add(Me.txtUserID)
        Me.panCred.Location = New System.Drawing.Point(3, 3)
        Me.panCred.Name = "panCred"
        Me.panCred.Size = New System.Drawing.Size(479, 109)
        Me.panCred.TabIndex = 0
        Me.panCred.TabStop = True
        '
        'txtUserName
        '
        Me.txtUserName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUserName.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUserName.Location = New System.Drawing.Point(220, 42)
        Me.txtUserName.Name = "txtUserName"
        Me.txtUserName.ReadOnly = True
        Me.txtUserName.Size = New System.Drawing.Size(254, 23)
        Me.txtUserName.TabIndex = 1
        Me.txtUserName.TabStop = False
        '
        'lblUserName
        '
        Me.lblUserName.AutoSize = True
        Me.lblUserName.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUserName.Location = New System.Drawing.Point(8, 44)
        Me.lblUserName.Name = "lblUserName"
        Me.lblUserName.Size = New System.Drawing.Size(176, 21)
        Me.lblUserName.TabIndex = 9
        Me.lblUserName.Text = "StudyDoc User Name:"
        '
        'panRFC
        '
        Me.panRFC.Controls.Add(Me.panRFCE)
        Me.panRFC.Controls.Add(Me.lblRFC)
        Me.panRFC.Controls.Add(Me.cbxRFC)
        Me.panRFC.Location = New System.Drawing.Point(3, 241)
        Me.panRFC.Name = "panRFC"
        Me.panRFC.Size = New System.Drawing.Size(479, 130)
        Me.panRFC.TabIndex = 18
        Me.panRFC.TabStop = True
        '
        'panRFCE
        '
        Me.panRFCE.Controls.Add(Me.Label2)
        Me.panRFCE.Controls.Add(Me.lblRFCE)
        Me.panRFCE.Controls.Add(Me.txtRFC)
        Me.panRFCE.Location = New System.Drawing.Point(8, 32)
        Me.panRFCE.Name = "panRFCE"
        Me.panRFCE.Size = New System.Drawing.Size(471, 89)
        Me.panRFCE.TabIndex = 19
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(209, 7)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(78, 13)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Free Form Text"
        '
        'lblRFCE
        '
        Me.lblRFCE.ForeColor = System.Drawing.Color.Black
        Me.lblRFCE.Location = New System.Drawing.Point(3, 23)
        Me.lblRFCE.Name = "lblRFCE"
        Me.lblRFCE.Size = New System.Drawing.Size(193, 60)
        Me.lblRFCE.TabIndex = 13
        Me.lblRFCE.Text = "lblRFCE"
        '
        'panOK
        '
        Me.panOK.Controls.Add(Me.Button1)
        Me.panOK.Controls.Add(Me.cmdCancel)
        Me.panOK.Controls.Add(Me.cmdOK)
        Me.panOK.Location = New System.Drawing.Point(3, 377)
        Me.panOK.Name = "panOK"
        Me.panOK.Size = New System.Drawing.Size(259, 53)
        Me.panOK.TabIndex = 19
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(109, 0)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(29, 29)
        Me.Button1.TabIndex = 25
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'chkRFC
        '
        Me.chkRFC.AutoSize = True
        Me.chkRFC.Location = New System.Drawing.Point(371, 377)
        Me.chkRFC.Name = "chkRFC"
        Me.chkRFC.Size = New System.Drawing.Size(121, 17)
        Me.chkRFC.TabIndex = 20
        Me.chkRFC.Text = "Reason For Change"
        Me.chkRFC.UseVisualStyleBackColor = True
        Me.chkRFC.Visible = False
        '
        'chkMOS
        '
        Me.chkMOS.AutoSize = True
        Me.chkMOS.Location = New System.Drawing.Point(268, 377)
        Me.chkMOS.Name = "chkMOS"
        Me.chkMOS.Size = New System.Drawing.Size(97, 17)
        Me.chkMOS.TabIndex = 21
        Me.chkMOS.Text = "Meaning of Sig"
        Me.chkMOS.UseVisualStyleBackColor = True
        Me.chkMOS.Visible = False
        '
        'chkRRFC
        '
        Me.chkRRFC.AutoSize = True
        Me.chkRRFC.Location = New System.Drawing.Point(371, 400)
        Me.chkRRFC.Name = "chkRRFC"
        Me.chkRRFC.Size = New System.Drawing.Size(148, 17)
        Me.chkRRFC.TabIndex = 22
        Me.chkRRFC.Text = "RFC Restrict to dropdown"
        Me.chkRRFC.UseVisualStyleBackColor = True
        Me.chkRRFC.Visible = False
        '
        'chkRMOS
        '
        Me.chkRMOS.AutoSize = True
        Me.chkRMOS.Location = New System.Drawing.Point(268, 400)
        Me.chkRMOS.Name = "chkRMOS"
        Me.chkRMOS.Size = New System.Drawing.Size(136, 17)
        Me.chkRMOS.TabIndex = 23
        Me.chkRMOS.Text = "Restrict Meaning of Sig"
        Me.chkRMOS.UseVisualStyleBackColor = True
        Me.chkRMOS.Visible = False
        '
        'lblTest
        '
        Me.lblTest.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTest.ForeColor = System.Drawing.Color.Blue
        Me.lblTest.Location = New System.Drawing.Point(12, 9)
        Me.lblTest.Name = "lblTest"
        Me.lblTest.Size = New System.Drawing.Size(474, 105)
        Me.lblTest.TabIndex = 25
        Me.lblTest.Text = "Label3"
        Me.lblTest.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'pan1
        '
        Me.pan1.Controls.Add(Me.rbESigOn)
        Me.pan1.Controls.Add(Me.panCred)
        Me.pan1.Controls.Add(Me.panMOS)
        Me.pan1.Controls.Add(Me.chkRRFC)
        Me.pan1.Controls.Add(Me.panRFC)
        Me.pan1.Controls.Add(Me.panOK)
        Me.pan1.Controls.Add(Me.chkRMOS)
        Me.pan1.Controls.Add(Me.chkRFC)
        Me.pan1.Controls.Add(Me.chkMOS)
        Me.pan1.Location = New System.Drawing.Point(7, 117)
        Me.pan1.Name = "pan1"
        Me.pan1.Size = New System.Drawing.Size(488, 443)
        Me.pan1.TabIndex = 0
        '
        'rbESigOn
        '
        Me.rbESigOn.AutoSize = true
        Me.rbESigOn.Location = New System.Drawing.Point(268, 423)
        Me.rbESigOn.Name = "rbESigOn"
        Me.rbESigOn.Size = New System.Drawing.Size(145, 17)
        Me.rbESigOn.TabIndex = 24
        Me.rbESigOn.Text = "On (Require ESig prompt)"
        Me.rbESigOn.UseVisualStyleBackColor = true
        Me.rbESigOn.Visible = false
        '
        'frmESig
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 13!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229,Byte),Integer), CType(CType(239,Byte),Integer), CType(CType(249,Byte),Integer))
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(505, 590)
        Me.Controls.Add(Me.pan1)
        Me.Controls.Add(Me.lblTest)
        Me.Icon = CType(resources.GetObject("$this.Icon"),System.Drawing.Icon)
        Me.MaximizeBox = false
        Me.MinimizeBox = false
        Me.Name = "frmESig"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Electronic Signature"
        Me.panMOS.ResumeLayout(false)
        Me.panMOS.PerformLayout
        Me.panMOSE.ResumeLayout(false)
        Me.panMOSE.PerformLayout
        Me.panCred.ResumeLayout(false)
        Me.panCred.PerformLayout
        Me.panRFC.ResumeLayout(false)
        Me.panRFC.PerformLayout
        Me.panRFCE.ResumeLayout(false)
        Me.panRFCE.PerformLayout
        Me.panOK.ResumeLayout(false)
        Me.pan1.ResumeLayout(false)
        Me.pan1.PerformLayout
        Me.ResumeLayout(false)

End Sub
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents txtUserID As System.Windows.Forms.TextBox
    Friend WithEvents lblPassword As System.Windows.Forms.Label
    Friend WithEvents lblUserID As System.Windows.Forms.Label
    Friend WithEvents lblMOS As System.Windows.Forms.Label
    Friend WithEvents cbxMOS As System.Windows.Forms.ComboBox
    Friend WithEvents cbxRFC As System.Windows.Forms.ComboBox
    Friend WithEvents lblRFC As System.Windows.Forms.Label
    Friend WithEvents txtMOS As System.Windows.Forms.TextBox
    Friend WithEvents txtRFC As System.Windows.Forms.TextBox
    Friend WithEvents panMOS As System.Windows.Forms.Panel
    Friend WithEvents panCred As System.Windows.Forms.Panel
    Friend WithEvents panRFC As System.Windows.Forms.Panel
    Friend WithEvents txtUserName As System.Windows.Forms.TextBox
    Friend WithEvents lblUserName As System.Windows.Forms.Label
    Friend WithEvents lblMOSE As System.Windows.Forms.Label
    Friend WithEvents lblRFCE As System.Windows.Forms.Label
    Friend WithEvents panMOSE As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents panRFCE As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents panOK As System.Windows.Forms.Panel
    Friend WithEvents chkRFC As System.Windows.Forms.CheckBox
    Friend WithEvents chkMOS As System.Windows.Forms.CheckBox
    Friend WithEvents chkRRFC As System.Windows.Forms.CheckBox
    Friend WithEvents chkRMOS As System.Windows.Forms.CheckBox
    Friend WithEvents lblTest As System.Windows.Forms.Label
    Friend WithEvents pan1 As System.Windows.Forms.Panel
    Friend WithEvents rbESigOn As System.Windows.Forms.RadioButton
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
