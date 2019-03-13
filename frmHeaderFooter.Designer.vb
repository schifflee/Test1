<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmHeaderFooter
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmHeaderFooter))
        Me.chkDiffFirstPage = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.chkIncludeLogo = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cmsRHeader = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.cmiFieldCodes = New System.Windows.Forms.ToolStripMenuItem()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.cmdEdit = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.CHARHLT = New System.Windows.Forms.TextBox()
        Me.CHARHLB = New System.Windows.Forms.TextBox()
        Me.CHARFLB = New System.Windows.Forms.TextBox()
        Me.CHARFLT = New System.Windows.Forms.TextBox()
        Me.CHARHRT = New System.Windows.Forms.TextBox()
        Me.CHARHRB = New System.Windows.Forms.TextBox()
        Me.CHARFRB = New System.Windows.Forms.TextBox()
        Me.CHARFRT = New System.Windows.Forms.TextBox()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.cmsRHeader.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkDiffFirstPage
        '
        Me.chkDiffFirstPage.AutoSize = True
        Me.chkDiffFirstPage.Location = New System.Drawing.Point(6, 117)
        Me.chkDiffFirstPage.Name = "chkDiffFirstPage"
        Me.chkDiffFirstPage.Size = New System.Drawing.Size(116, 17)
        Me.chkDiffFirstPage.TabIndex = 8
        Me.chkDiffFirstPage.Text = "Different First Page"
        Me.chkDiffFirstPage.UseVisualStyleBackColor = True
        Me.chkDiffFirstPage.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.Color.Red
        Me.Label1.Location = New System.Drawing.Point(128, 118)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(407, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "(If checked, only 1st page will have header items contained in GuWu Word template" & _
    ")"
        Me.Label1.Visible = False
        '
        'chkIncludeLogo
        '
        Me.chkIncludeLogo.AutoSize = True
        Me.chkIncludeLogo.Location = New System.Drawing.Point(6, 140)
        Me.chkIncludeLogo.Name = "chkIncludeLogo"
        Me.chkIncludeLogo.Size = New System.Drawing.Size(155, 17)
        Me.chkIncludeLogo.TabIndex = 9
        Me.chkIncludeLogo.Text = "Include logo on every page"
        Me.chkIncludeLogo.UseVisualStyleBackColor = True
        Me.chkIncludeLogo.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ForeColor = System.Drawing.Color.Red
        Me.Label2.Location = New System.Drawing.Point(167, 141)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(215, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "(Logo must appear in GuWu Word template)"
        Me.Label2.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(3, 159)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 16)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Header"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(3, 367)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 16)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Footer"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Blue
        Me.Label5.Location = New System.Drawing.Point(694, 160)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(211, 15)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "Right-click to insert Field Codes"
        '
        'cmsRHeader
        '
        Me.cmsRHeader.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.cmiFieldCodes})
        Me.cmsRHeader.Name = "cmsRHeader"
        Me.cmsRHeader.Size = New System.Drawing.Size(172, 26)
        '
        'cmiFieldCodes
        '
        Me.cmiFieldCodes.Name = "cmiFieldCodes"
        Me.cmiFieldCodes.Size = New System.Drawing.Size(171, 22)
        Me.cmiFieldCodes.Text = "Insert Field Code..."
        '
        'cmdSave
        '
        Me.cmdSave.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdSave.Enabled = False
        Me.cmdSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ForeColor = System.Drawing.Color.ForestGreen
        Me.cmdSave.Location = New System.Drawing.Point(696, 3)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(68, 25)
        Me.cmdSave.TabIndex = 11
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.UseVisualStyleBackColor = False
        '
        'cmdEdit
        '
        Me.cmdEdit.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdEdit.CausesValidation = False
        Me.cmdEdit.Enabled = False
        Me.cmdEdit.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.cmdEdit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdEdit.ForeColor = System.Drawing.Color.Blue
        Me.cmdEdit.Location = New System.Drawing.Point(626, 3)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.Size = New System.Drawing.Size(68, 25)
        Me.cmdEdit.TabIndex = 10
        Me.cmdEdit.Text = "&Edit"
        Me.cmdEdit.UseVisualStyleBackColor = False
        '
        'cmdExit
        '
        Me.cmdExit.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdExit.CausesValidation = False
        Me.cmdExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ForeColor = System.Drawing.Color.Red
        Me.cmdExit.Location = New System.Drawing.Point(837, 3)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(68, 25)
        Me.cmdExit.TabIndex = 13
        Me.cmdExit.Text = "E&xit"
        Me.cmdExit.UseVisualStyleBackColor = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancel.CausesValidation = False
        Me.cmdCancel.Enabled = False
        Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(766, 3)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(68, 25)
        Me.cmdCancel.TabIndex = 12
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Blue
        Me.Label6.Location = New System.Drawing.Point(696, 368)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(211, 15)
        Me.Label6.TabIndex = 95
        Me.Label6.Text = "Right-click to insert Field Codes"
        '
        'CHARHLT
        '
        Me.CHARHLT.ContextMenuStrip = Me.cmsRHeader
        Me.CHARHLT.Location = New System.Drawing.Point(6, 178)
        Me.CHARHLT.MinimumSize = New System.Drawing.Size(450, 80)
        Me.CHARHLT.Multiline = True
        Me.CHARHLT.Name = "CHARHLT"
        Me.CHARHLT.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.CHARHLT.Size = New System.Drawing.Size(450, 80)
        Me.CHARHLT.TabIndex = 0
        Me.CHARHLT.WordWrap = False
        '
        'CHARHLB
        '
        Me.CHARHLB.ContextMenuStrip = Me.cmsRHeader
        Me.CHARHLB.Location = New System.Drawing.Point(6, 257)
        Me.CHARHLB.MinimumSize = New System.Drawing.Size(450, 80)
        Me.CHARHLB.Multiline = True
        Me.CHARHLB.Name = "CHARHLB"
        Me.CHARHLB.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.CHARHLB.Size = New System.Drawing.Size(450, 80)
        Me.CHARHLB.TabIndex = 2
        Me.CHARHLB.WordWrap = False
        '
        'CHARFLB
        '
        Me.CHARFLB.ContextMenuStrip = Me.cmsRHeader
        Me.CHARFLB.Location = New System.Drawing.Point(6, 465)
        Me.CHARFLB.MinimumSize = New System.Drawing.Size(450, 80)
        Me.CHARFLB.Multiline = True
        Me.CHARFLB.Name = "CHARFLB"
        Me.CHARFLB.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.CHARFLB.Size = New System.Drawing.Size(450, 80)
        Me.CHARFLB.TabIndex = 6
        Me.CHARFLB.WordWrap = False
        '
        'CHARFLT
        '
        Me.CHARFLT.ContextMenuStrip = Me.cmsRHeader
        Me.CHARFLT.Location = New System.Drawing.Point(6, 386)
        Me.CHARFLT.MinimumSize = New System.Drawing.Size(450, 80)
        Me.CHARFLT.Multiline = True
        Me.CHARFLT.Name = "CHARFLT"
        Me.CHARFLT.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.CHARFLT.Size = New System.Drawing.Size(450, 80)
        Me.CHARFLT.TabIndex = 4
        Me.CHARFLT.WordWrap = False
        '
        'CHARHRT
        '
        Me.CHARHRT.ContextMenuStrip = Me.cmsRHeader
        Me.CHARHRT.Location = New System.Drawing.Point(455, 178)
        Me.CHARHRT.MinimumSize = New System.Drawing.Size(450, 80)
        Me.CHARHRT.Multiline = True
        Me.CHARHRT.Name = "CHARHRT"
        Me.CHARHRT.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.CHARHRT.Size = New System.Drawing.Size(450, 80)
        Me.CHARHRT.TabIndex = 1
        Me.CHARHRT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.CHARHRT.WordWrap = False
        '
        'CHARHRB
        '
        Me.CHARHRB.ContextMenuStrip = Me.cmsRHeader
        Me.CHARHRB.Location = New System.Drawing.Point(455, 257)
        Me.CHARHRB.MinimumSize = New System.Drawing.Size(450, 80)
        Me.CHARHRB.Multiline = True
        Me.CHARHRB.Name = "CHARHRB"
        Me.CHARHRB.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.CHARHRB.Size = New System.Drawing.Size(450, 80)
        Me.CHARHRB.TabIndex = 3
        Me.CHARHRB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.CHARHRB.WordWrap = False
        '
        'CHARFRB
        '
        Me.CHARFRB.ContextMenuStrip = Me.cmsRHeader
        Me.CHARFRB.Location = New System.Drawing.Point(455, 465)
        Me.CHARFRB.MinimumSize = New System.Drawing.Size(450, 80)
        Me.CHARFRB.Multiline = True
        Me.CHARFRB.Name = "CHARFRB"
        Me.CHARFRB.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.CHARFRB.Size = New System.Drawing.Size(450, 80)
        Me.CHARFRB.TabIndex = 7
        Me.CHARFRB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.CHARFRB.WordWrap = False
        '
        'CHARFRT
        '
        Me.CHARFRT.ContextMenuStrip = Me.cmsRHeader
        Me.CHARFRT.Location = New System.Drawing.Point(455, 386)
        Me.CHARFRT.MinimumSize = New System.Drawing.Size(450, 80)
        Me.CHARFRT.Multiline = True
        Me.CHARFRT.Name = "CHARFRT"
        Me.CHARFRT.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.CHARFRT.Size = New System.Drawing.Size(450, 80)
        Me.CHARFRT.TabIndex = 5
        Me.CHARFRT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.CHARFRT.WordWrap = False
        '
        'lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.Color.Blue
        Me.lblTitle.Location = New System.Drawing.Point(3, 32)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(49, 16)
        Me.lblTitle.TabIndex = 104
        Me.lblTitle.Text = "Label7"
        '
        'frmHeaderFooter
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(914, 554)
        Me.ControlBox = False
        Me.Controls.Add(Me.CHARFRT)
        Me.Controls.Add(Me.CHARHRB)
        Me.Controls.Add(Me.CHARFRB)
        Me.Controls.Add(Me.CHARHRT)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.CHARFLB)
        Me.Controls.Add(Me.CHARHLB)
        Me.Controls.Add(Me.CHARFLT)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.CHARHLT)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cmdEdit)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.chkIncludeLogo)
        Me.Controls.Add(Me.chkDiffFirstPage)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmHeaderFooter"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " Configure Report Header/Footer"
        Me.cmsRHeader.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents chkDiffFirstPage As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents chkIncludeLogo As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdEdit As System.Windows.Forms.Button
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmsRHeader As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents cmiFieldCodes As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents CHARHLT As System.Windows.Forms.TextBox
    Friend WithEvents CHARHLB As System.Windows.Forms.TextBox
    Friend WithEvents CHARFLB As System.Windows.Forms.TextBox
    Friend WithEvents CHARFLT As System.Windows.Forms.TextBox
    Friend WithEvents CHARHRT As System.Windows.Forms.TextBox
    Friend WithEvents CHARHRB As System.Windows.Forms.TextBox
    Friend WithEvents CHARFRB As System.Windows.Forms.TextBox
    Friend WithEvents CHARFRT As System.Windows.Forms.TextBox
    Friend WithEvents lblTitle As System.Windows.Forms.Label
End Class
