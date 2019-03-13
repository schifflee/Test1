<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCode
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCode))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.txtDeEncrypt = New System.Windows.Forms.TextBox()
        Me.txtEncrypt = New System.Windows.Forms.TextBox()
        Me.cmdEncrypt = New System.Windows.Forms.Button()
        Me.cmdDeEncrypt = New System.Windows.Forms.Button()
        Me.txtEncryptO = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtAscii = New System.Windows.Forms.TextBox()
        Me.lblAsc = New System.Windows.Forms.Label()
        Me.txtAsciiO = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtDeEncryptO = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(33, 45)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(53, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Password"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(35, 129)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(71, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "De-encrypted"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(33, 77)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(55, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Encrypted"
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(122, 42)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Size = New System.Drawing.Size(449, 20)
        Me.txtPassword.TabIndex = 3
        '
        'txtDeEncrypt
        '
        Me.txtDeEncrypt.Location = New System.Drawing.Point(122, 126)
        Me.txtDeEncrypt.Name = "txtDeEncrypt"
        Me.txtDeEncrypt.Size = New System.Drawing.Size(449, 20)
        Me.txtDeEncrypt.TabIndex = 4
        '
        'txtEncrypt
        '
        Me.txtEncrypt.Location = New System.Drawing.Point(122, 74)
        Me.txtEncrypt.Name = "txtEncrypt"
        Me.txtEncrypt.Size = New System.Drawing.Size(449, 20)
        Me.txtEncrypt.TabIndex = 5
        '
        'cmdEncrypt
        '
        Me.cmdEncrypt.Location = New System.Drawing.Point(213, 270)
        Me.cmdEncrypt.Name = "cmdEncrypt"
        Me.cmdEncrypt.Size = New System.Drawing.Size(75, 23)
        Me.cmdEncrypt.TabIndex = 6
        Me.cmdEncrypt.Text = "Encrypt"
        Me.cmdEncrypt.UseVisualStyleBackColor = True
        '
        'cmdDeEncrypt
        '
        Me.cmdDeEncrypt.Location = New System.Drawing.Point(213, 305)
        Me.cmdDeEncrypt.Name = "cmdDeEncrypt"
        Me.cmdDeEncrypt.Size = New System.Drawing.Size(75, 23)
        Me.cmdDeEncrypt.TabIndex = 7
        Me.cmdDeEncrypt.Text = "De-Encrypt"
        Me.cmdDeEncrypt.UseVisualStyleBackColor = True
        '
        'txtEncryptO
        '
        Me.txtEncryptO.Location = New System.Drawing.Point(122, 178)
        Me.txtEncryptO.Name = "txtEncryptO"
        Me.txtEncryptO.Size = New System.Drawing.Size(449, 20)
        Me.txtEncryptO.TabIndex = 9
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(35, 181)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(66, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Encrypted O"
        '
        'txtAscii
        '
        Me.txtAscii.Location = New System.Drawing.Point(122, 100)
        Me.txtAscii.Name = "txtAscii"
        Me.txtAscii.Size = New System.Drawing.Size(449, 20)
        Me.txtAscii.TabIndex = 11
        '
        'lblAsc
        '
        Me.lblAsc.AutoSize = True
        Me.lblAsc.Location = New System.Drawing.Point(35, 103)
        Me.lblAsc.Name = "lblAsc"
        Me.lblAsc.Size = New System.Drawing.Size(29, 13)
        Me.lblAsc.TabIndex = 10
        Me.lblAsc.Text = "Ascii"
        '
        'txtAsciiO
        '
        Me.txtAsciiO.Location = New System.Drawing.Point(122, 204)
        Me.txtAsciiO.Name = "txtAsciiO"
        Me.txtAsciiO.Size = New System.Drawing.Size(449, 20)
        Me.txtAsciiO.TabIndex = 13
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(35, 207)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(40, 13)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Ascii O"
        '
        'txtDeEncryptO
        '
        Me.txtDeEncryptO.Location = New System.Drawing.Point(122, 230)
        Me.txtDeEncryptO.Name = "txtDeEncryptO"
        Me.txtDeEncryptO.Size = New System.Drawing.Size(449, 20)
        Me.txtDeEncryptO.TabIndex = 15
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(35, 233)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(82, 13)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "De-encrypted O"
        '
        'frmCode
        '
        Me.AcceptButton = Me.cmdEncrypt
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(615, 335)
        Me.Controls.Add(Me.txtDeEncryptO)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txtAsciiO)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtAscii)
        Me.Controls.Add(Me.lblAsc)
        Me.Controls.Add(Me.txtEncryptO)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cmdDeEncrypt)
        Me.Controls.Add(Me.cmdEncrypt)
        Me.Controls.Add(Me.txtEncrypt)
        Me.Controls.Add(Me.txtDeEncrypt)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmCode"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents txtDeEncrypt As System.Windows.Forms.TextBox
    Friend WithEvents txtEncrypt As System.Windows.Forms.TextBox
    Friend WithEvents cmdEncrypt As System.Windows.Forms.Button
    Friend WithEvents cmdDeEncrypt As System.Windows.Forms.Button
    Friend WithEvents txtEncryptO As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtAscii As System.Windows.Forms.TextBox
    Friend WithEvents lblAsc As System.Windows.Forms.Label
    Friend WithEvents txtAsciiO As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtDeEncryptO As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
End Class
