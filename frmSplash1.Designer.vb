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
        Me.pan1 = New System.Windows.Forms.Panel()
        Me.lbl1 = New System.Windows.Forms.Label()
        Me.lbl1b = New System.Windows.Forms.Label()
        Me.lbl2 = New System.Windows.Forms.Label()
        Me.picGubbs = New System.Windows.Forms.PictureBox()
        Me.pb1 = New System.Windows.Forms.ProgressBar()
        Me.lblBoot = New System.Windows.Forms.Label()
        Me.lblC = New System.Windows.Forms.Label()
        Me.lbl1a = New System.Windows.Forms.Label()
        Me.pb2 = New System.Windows.Forms.ProgressBar()
        Me.lblErr = New System.Windows.Forms.Label()
        Me.txtMax = New System.Windows.Forms.TextBox()
        Me.txtValue = New System.Windows.Forms.TextBox()
        Me.pan1.SuspendLayout()
        CType(Me.picGubbs, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pan1
        '
        Me.pan1.AutoScroll = True
        Me.pan1.BackColor = System.Drawing.Color.Goldenrod
        Me.pan1.BackgroundImage = CType(resources.GetObject("pan1.BackgroundImage"), System.Drawing.Image)
        Me.pan1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.pan1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pan1.Controls.Add(Me.lbl1)
        Me.pan1.Controls.Add(Me.lbl1b)
        Me.pan1.Controls.Add(Me.lbl2)
        Me.pan1.Controls.Add(Me.picGubbs)
        Me.pan1.Controls.Add(Me.pb1)
        Me.pan1.Controls.Add(Me.lblBoot)
        Me.pan1.Controls.Add(Me.lblC)
        Me.pan1.Controls.Add(Me.lbl1a)
        Me.pan1.Location = New System.Drawing.Point(186, 95)
        Me.pan1.Name = "pan1"
        Me.pan1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pan1.Size = New System.Drawing.Size(479, 361)
        Me.pan1.TabIndex = 108
        Me.pan1.Visible = False
        '
        'lbl1
        '
        Me.lbl1.AutoEllipsis = True
        Me.lbl1.AutoSize = True
        Me.lbl1.BackColor = System.Drawing.Color.Transparent
        Me.lbl1.Font = New System.Drawing.Font("Century Gothic", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(25, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.lbl1.Location = New System.Drawing.Point(24, 157)
        Me.lbl1.Margin = New System.Windows.Forms.Padding(0)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(202, 77)
        Me.lbl1.TabIndex = 118
        Me.lbl1.Text = "Study"
        Me.lbl1.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'lbl1b
        '
        Me.lbl1b.AutoSize = True
        Me.lbl1b.BackColor = System.Drawing.Color.Transparent
        Me.lbl1b.Font = New System.Drawing.Font("Century Gothic", 14.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl1b.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lbl1b.Location = New System.Drawing.Point(344, 163)
        Me.lbl1b.Name = "lbl1b"
        Me.lbl1b.Size = New System.Drawing.Size(35, 23)
        Me.lbl1b.TabIndex = 115
        Me.lbl1b.Text = "TM"
        Me.lbl1b.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'lbl2
        '
        Me.lbl2.AutoEllipsis = True
        Me.lbl2.AutoSize = True
        Me.lbl2.BackColor = System.Drawing.Color.Transparent
        Me.lbl2.Font = New System.Drawing.Font("Century Gothic", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(25, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.lbl2.Location = New System.Drawing.Point(65, 252)
        Me.lbl2.Name = "lbl2"
        Me.lbl2.Size = New System.Drawing.Size(190, 19)
        Me.lbl2.TabIndex = 4
        Me.lbl2.Text = "Report Writing Manager"
        '
        'picGubbs
        '
        Me.picGubbs.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.picGubbs.Image = CType(resources.GetObject("picGubbs.Image"), System.Drawing.Image)
        Me.picGubbs.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.picGubbs.Location = New System.Drawing.Point(372, 313)
        Me.picGubbs.Name = "picGubbs"
        Me.picGubbs.Size = New System.Drawing.Size(105, 27)
        Me.picGubbs.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picGubbs.TabIndex = 113
        Me.picGubbs.TabStop = False
        '
        'pb1
        '
        Me.pb1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.pb1.Cursor = System.Windows.Forms.Cursors.Default
        Me.pb1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.pb1.Location = New System.Drawing.Point(-1, 324)
        Me.pb1.Name = "pb1"
        Me.pb1.Size = New System.Drawing.Size(280, 17)
        Me.pb1.TabIndex = 112
        '
        'lblBoot
        '
        Me.lblBoot.BackColor = System.Drawing.Color.Transparent
        Me.lblBoot.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBoot.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblBoot.Location = New System.Drawing.Point(3, 294)
        Me.lblBoot.Name = "lblBoot"
        Me.lblBoot.Size = New System.Drawing.Size(170, 27)
        Me.lblBoot.TabIndex = 111
        Me.lblBoot.Text = "Starting application..."
        Me.lblBoot.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'lblC
        '
        Me.lblC.BackColor = System.Drawing.Color.White
        Me.lblC.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblC.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblC.Location = New System.Drawing.Point(285, 324)
        Me.lblC.Name = "lblC"
        Me.lblC.Size = New System.Drawing.Size(65, 16)
        Me.lblC.TabIndex = 113
        Me.lblC.Text = "Label5"
        Me.lblC.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblC.UseWaitCursor = True
        '
        'lbl1a
        '
        Me.lbl1a.AutoEllipsis = True
        Me.lbl1a.AutoSize = True
        Me.lbl1a.BackColor = System.Drawing.Color.Transparent
        Me.lbl1a.Font = New System.Drawing.Font("Century Gothic", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl1a.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lbl1a.Location = New System.Drawing.Point(209, 157)
        Me.lbl1a.Margin = New System.Windows.Forms.Padding(0)
        Me.lbl1a.Name = "lbl1a"
        Me.lbl1a.Size = New System.Drawing.Size(160, 77)
        Me.lbl1a.TabIndex = 119
        Me.lbl1a.Text = "Doc"
        Me.lbl1a.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'pb2
        '
        Me.pb2.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.pb2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.pb2.Location = New System.Drawing.Point(162, 66)
        Me.pb2.Name = "pb2"
        Me.pb2.Size = New System.Drawing.Size(479, 23)
        Me.pb2.TabIndex = 110
        Me.pb2.Visible = False
        '
        'lblErr
        '
        Me.lblErr.AutoSize = True
        Me.lblErr.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblErr.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblErr.Location = New System.Drawing.Point(12, 9)
        Me.lblErr.Name = "lblErr"
        Me.lblErr.Size = New System.Drawing.Size(79, 13)
        Me.lblErr.TabIndex = 109
        Me.lblErr.Text = "Establishing..."
        Me.lblErr.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'txtMax
        '
        Me.txtMax.Location = New System.Drawing.Point(683, 118)
        Me.txtMax.Name = "txtMax"
        Me.txtMax.Size = New System.Drawing.Size(50, 20)
        Me.txtMax.TabIndex = 111
        Me.txtMax.Visible = False
        '
        'txtValue
        '
        Me.txtValue.Location = New System.Drawing.Point(683, 149)
        Me.txtValue.Name = "txtValue"
        Me.txtValue.Size = New System.Drawing.Size(50, 20)
        Me.txtValue.TabIndex = 112
        Me.txtValue.Visible = False
        '
        'frmSplash1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(812, 660)
        Me.Controls.Add(Me.txtValue)
        Me.Controls.Add(Me.txtMax)
        Me.Controls.Add(Me.pb2)
        Me.Controls.Add(Me.lblErr)
        Me.Controls.Add(Me.pan1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSplash1"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "   StudyDoc start up..."
        Me.pan1.ResumeLayout(False)
        Me.pan1.PerformLayout()
        CType(Me.picGubbs, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents pan1 As System.Windows.Forms.Panel
    Friend WithEvents pb1 As System.Windows.Forms.ProgressBar
    Friend WithEvents lblBoot As System.Windows.Forms.Label
    Friend WithEvents lblC As System.Windows.Forms.Label
    Friend WithEvents lbl2 As System.Windows.Forms.Label
    Friend WithEvents lbl1b As System.Windows.Forms.Label
    Friend WithEvents pb2 As System.Windows.Forms.ProgressBar
    Friend WithEvents lblErr As System.Windows.Forms.Label
    Friend WithEvents txtMax As System.Windows.Forms.TextBox
    Friend WithEvents txtValue As System.Windows.Forms.TextBox
    Friend WithEvents picGubbs As System.Windows.Forms.PictureBox
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Friend WithEvents lbl1a As System.Windows.Forms.Label
End Class
