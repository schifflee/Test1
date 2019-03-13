<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMarginHeaders
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMarginHeaders))
        Me.gbMargin = New System.Windows.Forms.GroupBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.panParameters = New System.Windows.Forms.Panel()
        Me.gbFooter = New System.Windows.Forms.GroupBox()
        Me.rbFooterIsText = New System.Windows.Forms.RadioButton()
        Me.rbFooterIsFigure = New System.Windows.Forms.RadioButton()
        Me.charFooterVert = New System.Windows.Forms.ComboBox()
        Me.numFooterAbsPosHoriz = New System.Windows.Forms.TextBox()
        Me.charFooterHoriz = New System.Windows.Forms.ComboBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.numFooterAbsPosVert = New System.Windows.Forms.TextBox()
        Me.gbHeader = New System.Windows.Forms.GroupBox()
        Me.rbHeaderIsText = New System.Windows.Forms.RadioButton()
        Me.rbHeaderIsFigure = New System.Windows.Forms.RadioButton()
        Me.numHeaderAbsPosHoriz = New System.Windows.Forms.TextBox()
        Me.lbl1 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.numHeaderAbsPosVert = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblhp = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.charHeaderHoriz = New System.Windows.Forms.ComboBox()
        Me.charHeaderVert = New System.Windows.Forms.ComboBox()
        Me.rbMarginCustom = New System.Windows.Forms.RadioButton()
        Me.rbMarginGuWu = New System.Windows.Forms.RadioButton()
        Me.gbMargin.SuspendLayout()
        Me.panParameters.SuspendLayout()
        Me.gbFooter.SuspendLayout()
        Me.gbHeader.SuspendLayout()
        Me.SuspendLayout()
        '
        'gbMargin
        '
        Me.gbMargin.Controls.Add(Me.cmdCancel)
        Me.gbMargin.Controls.Add(Me.cmdOK)
        Me.gbMargin.Controls.Add(Me.panParameters)
        Me.gbMargin.Controls.Add(Me.rbMarginCustom)
        Me.gbMargin.Controls.Add(Me.rbMarginGuWu)
        Me.gbMargin.Location = New System.Drawing.Point(14, 16)
        Me.gbMargin.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbMargin.Name = "gbMargin"
        Me.gbMargin.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbMargin.Size = New System.Drawing.Size(446, 705)
        Me.gbMargin.TabIndex = 0
        Me.gbMargin.TabStop = False
        Me.gbMargin.Text = "Parameters"
        '
        'cmdCancel
        '
        Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(229, 636)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(77, 48)
        Me.cmdCancel.TabIndex = 4
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdOK
        '
        Me.cmdOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.Blue
        Me.cmdOK.Location = New System.Drawing.Point(124, 635)
        Me.cmdOK.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(77, 48)
        Me.cmdOK.TabIndex = 3
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'panParameters
        '
        Me.panParameters.Controls.Add(Me.gbFooter)
        Me.panParameters.Controls.Add(Me.gbHeader)
        Me.panParameters.Location = New System.Drawing.Point(45, 97)
        Me.panParameters.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panParameters.Name = "panParameters"
        Me.panParameters.Size = New System.Drawing.Size(323, 523)
        Me.panParameters.TabIndex = 2
        '
        'gbFooter
        '
        Me.gbFooter.Controls.Add(Me.rbFooterIsText)
        Me.gbFooter.Controls.Add(Me.rbFooterIsFigure)
        Me.gbFooter.Controls.Add(Me.charFooterVert)
        Me.gbFooter.Controls.Add(Me.numFooterAbsPosHoriz)
        Me.gbFooter.Controls.Add(Me.charFooterHoriz)
        Me.gbFooter.Controls.Add(Me.Label12)
        Me.gbFooter.Controls.Add(Me.Label7)
        Me.gbFooter.Controls.Add(Me.Label11)
        Me.gbFooter.Controls.Add(Me.Label8)
        Me.gbFooter.Controls.Add(Me.Label10)
        Me.gbFooter.Controls.Add(Me.Label9)
        Me.gbFooter.Controls.Add(Me.numFooterAbsPosVert)
        Me.gbFooter.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbFooter.ForeColor = System.Drawing.Color.Blue
        Me.gbFooter.Location = New System.Drawing.Point(3, 262)
        Me.gbFooter.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbFooter.Name = "gbFooter"
        Me.gbFooter.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbFooter.Size = New System.Drawing.Size(311, 254)
        Me.gbFooter.TabIndex = 2
        Me.gbFooter.TabStop = False
        Me.gbFooter.Text = "Footer"
        '
        'rbFooterIsText
        '
        Me.rbFooterIsText.AutoSize = True
        Me.rbFooterIsText.Checked = True
        Me.rbFooterIsText.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.rbFooterIsText.ForeColor = System.Drawing.Color.Black
        Me.rbFooterIsText.Location = New System.Drawing.Point(31, 49)
        Me.rbFooterIsText.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbFooterIsText.Name = "rbFooterIsText"
        Me.rbFooterIsText.Size = New System.Drawing.Size(105, 21)
        Me.rbFooterIsText.TabIndex = 20
        Me.rbFooterIsText.TabStop = True
        Me.rbFooterIsText.Text = "Footer is Text"
        Me.rbFooterIsText.UseVisualStyleBackColor = True
        '
        'rbFooterIsFigure
        '
        Me.rbFooterIsFigure.AutoSize = True
        Me.rbFooterIsFigure.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.rbFooterIsFigure.ForeColor = System.Drawing.Color.Black
        Me.rbFooterIsFigure.Location = New System.Drawing.Point(31, 24)
        Me.rbFooterIsFigure.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbFooterIsFigure.Name = "rbFooterIsFigure"
        Me.rbFooterIsFigure.Size = New System.Drawing.Size(117, 21)
        Me.rbFooterIsFigure.TabIndex = 19
        Me.rbFooterIsFigure.Text = "Footer is Figure"
        Me.rbFooterIsFigure.UseVisualStyleBackColor = True
        '
        'charFooterVert
        '
        Me.charFooterVert.FormattingEnabled = True
        Me.charFooterVert.Location = New System.Drawing.Point(172, 216)
        Me.charFooterVert.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.charFooterVert.Name = "charFooterVert"
        Me.charFooterVert.Size = New System.Drawing.Size(122, 25)
        Me.charFooterVert.TabIndex = 15
        '
        'numFooterAbsPosHoriz
        '
        Me.numFooterAbsPosHoriz.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.numFooterAbsPosHoriz.Location = New System.Drawing.Point(172, 100)
        Me.numFooterAbsPosHoriz.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.numFooterAbsPosHoriz.Name = "numFooterAbsPosHoriz"
        Me.numFooterAbsPosHoriz.Size = New System.Drawing.Size(122, 25)
        Me.numFooterAbsPosHoriz.TabIndex = 10
        Me.numFooterAbsPosHoriz.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'charFooterHoriz
        '
        Me.charFooterHoriz.FormattingEnabled = True
        Me.charFooterHoriz.Location = New System.Drawing.Point(172, 129)
        Me.charFooterHoriz.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.charFooterHoriz.Name = "charFooterHoriz"
        Me.charFooterHoriz.Size = New System.Drawing.Size(122, 25)
        Me.charFooterHoriz.TabIndex = 12
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.Label12.ForeColor = System.Drawing.Color.Black
        Me.Label12.Location = New System.Drawing.Point(61, 104)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(109, 17)
        Me.Label12.TabIndex = 9
        Me.Label12.Text = "Absolute Position"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(61, 220)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(43, 17)
        Me.Label7.TabIndex = 18
        Me.Label7.Text = "below"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Label11.ForeColor = System.Drawing.Color.Black
        Me.Label11.Location = New System.Drawing.Point(26, 79)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(73, 17)
        Me.Label11.TabIndex = 11
        Me.Label11.Text = "Horizontal"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(61, 133)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(89, 17)
        Me.Label8.TabIndex = 17
        Me.Label8.Text = "to the right of"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.Label10.ForeColor = System.Drawing.Color.Black
        Me.Label10.Location = New System.Drawing.Point(61, 191)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(109, 17)
        Me.Label10.TabIndex = 13
        Me.Label10.Text = "Absolute Position"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Label9.ForeColor = System.Drawing.Color.Black
        Me.Label9.Location = New System.Drawing.Point(26, 166)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(55, 17)
        Me.Label9.TabIndex = 16
        Me.Label9.Text = "Vertical"
        '
        'numFooterAbsPosVert
        '
        Me.numFooterAbsPosVert.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.numFooterAbsPosVert.Location = New System.Drawing.Point(172, 187)
        Me.numFooterAbsPosVert.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.numFooterAbsPosVert.Name = "numFooterAbsPosVert"
        Me.numFooterAbsPosVert.Size = New System.Drawing.Size(122, 25)
        Me.numFooterAbsPosVert.TabIndex = 14
        Me.numFooterAbsPosVert.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'gbHeader
        '
        Me.gbHeader.Controls.Add(Me.rbHeaderIsText)
        Me.gbHeader.Controls.Add(Me.rbHeaderIsFigure)
        Me.gbHeader.Controls.Add(Me.numHeaderAbsPosHoriz)
        Me.gbHeader.Controls.Add(Me.lbl1)
        Me.gbHeader.Controls.Add(Me.Label1)
        Me.gbHeader.Controls.Add(Me.Label2)
        Me.gbHeader.Controls.Add(Me.numHeaderAbsPosVert)
        Me.gbHeader.Controls.Add(Me.Label3)
        Me.gbHeader.Controls.Add(Me.lblhp)
        Me.gbHeader.Controls.Add(Me.Label4)
        Me.gbHeader.Controls.Add(Me.charHeaderHoriz)
        Me.gbHeader.Controls.Add(Me.charHeaderVert)
        Me.gbHeader.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbHeader.ForeColor = System.Drawing.Color.Blue
        Me.gbHeader.Location = New System.Drawing.Point(3, 4)
        Me.gbHeader.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbHeader.Name = "gbHeader"
        Me.gbHeader.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbHeader.Size = New System.Drawing.Size(311, 254)
        Me.gbHeader.TabIndex = 1
        Me.gbHeader.TabStop = False
        Me.gbHeader.Text = "Header"
        '
        'rbHeaderIsText
        '
        Me.rbHeaderIsText.AutoSize = True
        Me.rbHeaderIsText.Checked = True
        Me.rbHeaderIsText.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.rbHeaderIsText.ForeColor = System.Drawing.Color.Black
        Me.rbHeaderIsText.Location = New System.Drawing.Point(31, 49)
        Me.rbHeaderIsText.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbHeaderIsText.Name = "rbHeaderIsText"
        Me.rbHeaderIsText.Size = New System.Drawing.Size(110, 21)
        Me.rbHeaderIsText.TabIndex = 9
        Me.rbHeaderIsText.TabStop = True
        Me.rbHeaderIsText.Text = "Header is Text"
        Me.rbHeaderIsText.UseVisualStyleBackColor = True
        '
        'rbHeaderIsFigure
        '
        Me.rbHeaderIsFigure.AutoSize = True
        Me.rbHeaderIsFigure.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.rbHeaderIsFigure.ForeColor = System.Drawing.Color.Black
        Me.rbHeaderIsFigure.Location = New System.Drawing.Point(31, 24)
        Me.rbHeaderIsFigure.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbHeaderIsFigure.Name = "rbHeaderIsFigure"
        Me.rbHeaderIsFigure.Size = New System.Drawing.Size(133, 21)
        Me.rbHeaderIsFigure.TabIndex = 8
        Me.rbHeaderIsFigure.Text = "Header is a Figure"
        Me.rbHeaderIsFigure.UseVisualStyleBackColor = True
        '
        'numHeaderAbsPosHoriz
        '
        Me.numHeaderAbsPosHoriz.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.numHeaderAbsPosHoriz.Location = New System.Drawing.Point(172, 100)
        Me.numHeaderAbsPosHoriz.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.numHeaderAbsPosHoriz.Name = "numHeaderAbsPosHoriz"
        Me.numHeaderAbsPosHoriz.Size = New System.Drawing.Size(122, 25)
        Me.numHeaderAbsPosHoriz.TabIndex = 1
        Me.numHeaderAbsPosHoriz.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lbl1
        '
        Me.lbl1.AutoSize = True
        Me.lbl1.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.lbl1.ForeColor = System.Drawing.Color.Black
        Me.lbl1.Location = New System.Drawing.Point(61, 104)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(109, 17)
        Me.lbl1.TabIndex = 0
        Me.lbl1.Text = "Absolute Position"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(26, 79)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Horizontal"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(61, 191)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(109, 17)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Absolute Position"
        '
        'numHeaderAbsPosVert
        '
        Me.numHeaderAbsPosVert.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.numHeaderAbsPosVert.Location = New System.Drawing.Point(172, 187)
        Me.numHeaderAbsPosVert.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.numHeaderAbsPosVert.Name = "numHeaderAbsPosVert"
        Me.numHeaderAbsPosVert.Size = New System.Drawing.Size(122, 25)
        Me.numHeaderAbsPosVert.TabIndex = 3
        Me.numHeaderAbsPosVert.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(26, 166)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(55, 17)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Vertical"
        '
        'lblhp
        '
        Me.lblhp.AutoSize = True
        Me.lblhp.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.lblhp.ForeColor = System.Drawing.Color.Black
        Me.lblhp.Location = New System.Drawing.Point(61, 133)
        Me.lblhp.Name = "lblhp"
        Me.lblhp.Size = New System.Drawing.Size(89, 17)
        Me.lblhp.TabIndex = 6
        Me.lblhp.Text = "to the right of"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(61, 220)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(43, 17)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "below"
        '
        'charHeaderHoriz
        '
        Me.charHeaderHoriz.FormattingEnabled = True
        Me.charHeaderHoriz.Location = New System.Drawing.Point(172, 129)
        Me.charHeaderHoriz.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.charHeaderHoriz.Name = "charHeaderHoriz"
        Me.charHeaderHoriz.Size = New System.Drawing.Size(122, 25)
        Me.charHeaderHoriz.TabIndex = 2
        '
        'charHeaderVert
        '
        Me.charHeaderVert.FormattingEnabled = True
        Me.charHeaderVert.Location = New System.Drawing.Point(172, 216)
        Me.charHeaderVert.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.charHeaderVert.Name = "charHeaderVert"
        Me.charHeaderVert.Size = New System.Drawing.Size(122, 25)
        Me.charHeaderVert.TabIndex = 4
        '
        'rbMarginCustom
        '
        Me.rbMarginCustom.AutoSize = True
        Me.rbMarginCustom.Location = New System.Drawing.Point(20, 67)
        Me.rbMarginCustom.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbMarginCustom.Name = "rbMarginCustom"
        Me.rbMarginCustom.Size = New System.Drawing.Size(234, 21)
        Me.rbMarginCustom.TabIndex = 1
        Me.rbMarginCustom.Text = "Manually set placement parameters"
        Me.rbMarginCustom.UseVisualStyleBackColor = True
        '
        'rbMarginGuWu
        '
        Me.rbMarginGuWu.AutoSize = True
        Me.rbMarginGuWu.Checked = True
        Me.rbMarginGuWu.Location = New System.Drawing.Point(20, 37)
        Me.rbMarginGuWu.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbMarginGuWu.Name = "rbMarginGuWu"
        Me.rbMarginGuWu.Size = New System.Drawing.Size(416, 21)
        Me.rbMarginGuWu.TabIndex = 0
        Me.rbMarginGuWu.TabStop = True
        Me.rbMarginGuWu.Text = "Allow StudyDoc to determine header/footer placement parameters"
        Me.rbMarginGuWu.UseVisualStyleBackColor = True
        '
        'frmMarginHeaders
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(472, 738)
        Me.ControlBox = False
        Me.Controls.Add(Me.gbMargin)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "frmMarginHeaders"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Set Landscape Heater/Footer Right/Left Margin Parameters..."
        Me.gbMargin.ResumeLayout(False)
        Me.gbMargin.PerformLayout()
        Me.panParameters.ResumeLayout(False)
        Me.gbFooter.ResumeLayout(False)
        Me.gbFooter.PerformLayout()
        Me.gbHeader.ResumeLayout(False)
        Me.gbHeader.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbMargin As System.Windows.Forms.GroupBox
    Friend WithEvents rbMarginCustom As System.Windows.Forms.RadioButton
    Friend WithEvents rbMarginGuWu As System.Windows.Forms.RadioButton
    Friend WithEvents panParameters As System.Windows.Forms.Panel
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents charHeaderVert As System.Windows.Forms.ComboBox
    Friend WithEvents charHeaderHoriz As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblhp As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents numHeaderAbsPosVert As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents numHeaderAbsPosHoriz As System.Windows.Forms.TextBox
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Friend WithEvents gbHeader As System.Windows.Forms.GroupBox
    Friend WithEvents rbHeaderIsText As System.Windows.Forms.RadioButton
    Friend WithEvents rbHeaderIsFigure As System.Windows.Forms.RadioButton
    Friend WithEvents charFooterVert As System.Windows.Forms.ComboBox
    Friend WithEvents charFooterHoriz As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents numFooterAbsPosVert As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents numFooterAbsPosHoriz As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents gbFooter As System.Windows.Forms.GroupBox
    Friend WithEvents rbFooterIsText As System.Windows.Forms.RadioButton
    Friend WithEvents rbFooterIsFigure As System.Windows.Forms.RadioButton
End Class
