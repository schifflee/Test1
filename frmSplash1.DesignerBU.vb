<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSplash1
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSplash1))
        Me.pan1 = New System.Windows.Forms.Panel
        Me.lblR = New System.Windows.Forms.Label
        Me.lblR1 = New System.Windows.Forms.Label
        Me.pb1 = New System.Windows.Forms.ProgressBar
        Me.lblBoot = New System.Windows.Forms.Label
        Me.l = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.lbl2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.pb2 = New System.Windows.Forms.ProgressBar
        Me.lblErr = New System.Windows.Forms.Label
        Me.txtMax = New System.Windows.Forms.TextBox
        Me.txtValue = New System.Windows.Forms.TextBox
        Me.AxOfficeViewer1 = New AxOfficeViewer.AxOfficeViewer
        Me.pan1.SuspendLayout()
        CType(Me.AxOfficeViewer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pan1
        '
        Me.pan1.BackColor = System.Drawing.Color.Goldenrod
        Me.pan1.BackgroundImage = CType(resources.GetObject("pan1.BackgroundImage"), System.Drawing.Image)
        Me.pan1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.pan1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pan1.Controls.Add(Me.lblR)
        Me.pan1.Controls.Add(Me.lblR1)
        Me.pan1.Controls.Add(Me.pb1)
        Me.pan1.Controls.Add(Me.lblBoot)
        Me.pan1.Controls.Add(Me.l)
        Me.pan1.Controls.Add(Me.Label3)
        Me.pan1.Controls.Add(Me.lbl2)
        Me.pan1.Controls.Add(Me.Label1)
        Me.pan1.Controls.Add(Me.Label4)
        Me.pan1.Location = New System.Drawing.Point(162, 148)
        Me.pan1.Name = "pan1"
        Me.pan1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pan1.Size = New System.Drawing.Size(479, 344)
        Me.pan1.TabIndex = 108
        Me.pan1.Visible = False
        '
        'lblR
        '
        Me.lblR.AutoSize = True
        Me.lblR.BackColor = System.Drawing.Color.Transparent
        Me.lblR.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblR.Location = New System.Drawing.Point(438, 241)
        Me.lblR.Name = "lblR"
        Me.lblR.Size = New System.Drawing.Size(24, 23)
        Me.lblR.TabIndex = 115
        Me.lblR.Text = "R"
        Me.lblR.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblR1
        '
        Me.lblR1.AutoSize = True
        Me.lblR1.BackColor = System.Drawing.Color.Transparent
        Me.lblR1.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblR1.Location = New System.Drawing.Point(359, 57)
        Me.lblR1.Name = "lblR1"
        Me.lblR1.Size = New System.Drawing.Size(15, 15)
        Me.lblR1.TabIndex = 114
        Me.lblR1.Text = "R"
        Me.lblR1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pb1
        '
        Me.pb1.BackColor = System.Drawing.Color.Maroon
        Me.pb1.Cursor = System.Windows.Forms.Cursors.Default
        Me.pb1.ForeColor = System.Drawing.Color.Gold
        Me.pb1.Location = New System.Drawing.Point(7, 309)
        Me.pb1.Name = "pb1"
        Me.pb1.Size = New System.Drawing.Size(210, 17)
        Me.pb1.TabIndex = 112
        '
        'lblBoot
        '
        Me.lblBoot.BackColor = System.Drawing.Color.Transparent
        Me.lblBoot.Font = New System.Drawing.Font("Tahoma", 14.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBoot.ForeColor = System.Drawing.Color.Purple
        Me.lblBoot.Location = New System.Drawing.Point(6, 272)
        Me.lblBoot.Name = "lblBoot"
        Me.lblBoot.Size = New System.Drawing.Size(146, 29)
        Me.lblBoot.TabIndex = 111
        Me.lblBoot.Text = "....booting up"
        Me.lblBoot.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'l
        '
        Me.l.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.l.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.l.Location = New System.Drawing.Point(152, 285)
        Me.l.Name = "l"
        Me.l.Size = New System.Drawing.Size(65, 16)
        Me.l.TabIndex = 113
        Me.l.Text = "Label5"
        Me.l.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.l.UseWaitCursor = True
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 14.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Maroon
        Me.Label3.Location = New System.Drawing.Point(273, 301)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(202, 29)
        Me.Label3.TabIndex = 110
        Me.Label3.Text = "A Product of Gubbs Inc"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'lbl2
        '
        Me.lbl2.BackColor = System.Drawing.Color.Transparent
        Me.lbl2.Font = New System.Drawing.Font("Tahoma", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl2.ForeColor = System.Drawing.Color.DarkGoldenrod
        Me.lbl2.Location = New System.Drawing.Point(154, 72)
        Me.lbl2.Name = "lbl2"
        Me.lbl2.Size = New System.Drawing.Size(307, 76)
        Me.lbl2.TabIndex = 4
        Me.lbl2.Text = "Gubbs Watson(TM) Report Writing Solution"
        Me.lbl2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 15.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Maroon
        Me.Label1.Location = New System.Drawing.Point(3, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(107, 24)
        Me.Label1.TabIndex = 108
        Me.Label1.Text = "Welcome to"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Rockwell Extra Bold", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Maroon
        Me.Label4.Location = New System.Drawing.Point(118, 3)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(260, 76)
        Me.Label4.TabIndex = 109
        Me.Label4.Text = "GuWu"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'pb2
        '
        Me.pb2.BackColor = System.Drawing.Color.Maroon
        Me.pb2.ForeColor = System.Drawing.Color.Gold
        Me.pb2.Location = New System.Drawing.Point(162, 119)
        Me.pb2.Name = "pb2"
        Me.pb2.Size = New System.Drawing.Size(479, 23)
        Me.pb2.TabIndex = 110
        Me.pb2.Visible = False
        '
        'lblErr
        '
        Me.lblErr.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblErr.ForeColor = System.Drawing.Color.Maroon
        Me.lblErr.Location = New System.Drawing.Point(12, 9)
        Me.lblErr.Name = "lblErr"
        Me.lblErr.Size = New System.Drawing.Size(788, 106)
        Me.lblErr.TabIndex = 109
        Me.lblErr.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtMax
        '
        Me.txtMax.Location = New System.Drawing.Point(688, 170)
        Me.txtMax.Name = "txtMax"
        Me.txtMax.Size = New System.Drawing.Size(50, 20)
        Me.txtMax.TabIndex = 111
        Me.txtMax.Visible = False
        '
        'txtValue
        '
        Me.txtValue.Location = New System.Drawing.Point(688, 201)
        Me.txtValue.Name = "txtValue"
        Me.txtValue.Size = New System.Drawing.Size(50, 20)
        Me.txtValue.TabIndex = 112
        Me.txtValue.Visible = False
        '
        'AxOfficeViewer1
        '
        Me.AxOfficeViewer1.Enabled = True
        Me.AxOfficeViewer1.Location = New System.Drawing.Point(679, 74)
        Me.AxOfficeViewer1.Name = "AxOfficeViewer1"
        Me.AxOfficeViewer1.OcxState = CType(resources.GetObject("AxOfficeViewer1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxOfficeViewer1.Size = New System.Drawing.Size(75, 23)
        Me.AxOfficeViewer1.TabIndex = 113
        '
        'frmSplash1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LemonChiffon
        Me.ClientSize = New System.Drawing.Size(812, 660)
        Me.Controls.Add(Me.AxOfficeViewer1)
        Me.Controls.Add(Me.txtValue)
        Me.Controls.Add(Me.txtMax)
        Me.Controls.Add(Me.pb2)
        Me.Controls.Add(Me.lblErr)
        Me.Controls.Add(Me.pan1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSplash1"
        Me.Text = "   GuWu start up..."
        Me.pan1.ResumeLayout(False)
        Me.pan1.PerformLayout()
        CType(Me.AxOfficeViewer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents pan1 As System.Windows.Forms.Panel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents pb1 As System.Windows.Forms.ProgressBar
    Friend WithEvents lblBoot As System.Windows.Forms.Label
    Friend WithEvents l As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lbl2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblR As System.Windows.Forms.Label
    Friend WithEvents lblR1 As System.Windows.Forms.Label
    Friend WithEvents pb2 As System.Windows.Forms.ProgressBar
    Friend WithEvents lblErr As System.Windows.Forms.Label
    Friend WithEvents txtMax As System.Windows.Forms.TextBox
    Friend WithEvents txtValue As System.Windows.Forms.TextBox
    Friend WithEvents AxOfficeViewer1 As AxOfficeViewer.AxOfficeViewer
End Class
