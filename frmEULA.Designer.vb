<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEULA
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEULA))
        Me.txtTrial = New System.Windows.Forms.TextBox()
        Me.txtEULA = New System.Windows.Forms.TextBox()
        Me.txtSupport = New System.Windows.Forms.TextBox()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
        Me.rtxEULA = New Global.GooWoo.RichTextBoxPrintCtrl.RichTextBoxPrintCtrl()
        Me.SuspendLayout()
        '
        'txtTrial
        '
        Me.txtTrial.Location = New System.Drawing.Point(14, 671)
        Me.txtTrial.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtTrial.Name = "txtTrial"
        Me.txtTrial.Size = New System.Drawing.Size(157, 25)
        Me.txtTrial.TabIndex = 1
        Me.txtTrial.Text = resources.GetString("txtTrial.Text")
        Me.txtTrial.Visible = False
        '
        'txtEULA
        '
        Me.txtEULA.Location = New System.Drawing.Point(212, 671)
        Me.txtEULA.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtEULA.Name = "txtEULA"
        Me.txtEULA.Size = New System.Drawing.Size(157, 25)
        Me.txtEULA.TabIndex = 2
        Me.txtEULA.Text = resources.GetString("txtEULA.Text")
        Me.txtEULA.Visible = False
        '
        'txtSupport
        '
        Me.txtSupport.Location = New System.Drawing.Point(399, 671)
        Me.txtSupport.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtSupport.Name = "txtSupport"
        Me.txtSupport.Size = New System.Drawing.Size(157, 25)
        Me.txtSupport.TabIndex = 3
        Me.txtSupport.Visible = False
        '
        'cmdClose
        '
        Me.cmdClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClose.ForeColor = System.Drawing.Color.Blue
        Me.cmdClose.Location = New System.Drawing.Point(743, 671)
        Me.cmdClose.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(106, 35)
        Me.cmdClose.TabIndex = 10
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'cmdPrint
        '
        Me.cmdPrint.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPrint.ForeColor = System.Drawing.Color.Blue
        Me.cmdPrint.Location = New System.Drawing.Point(856, 671)
        Me.cmdPrint.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(106, 35)
        Me.cmdPrint.TabIndex = 11
        Me.cmdPrint.Text = "&Print"
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'PrintDialog1
        '
        Me.PrintDialog1.Document = Me.PrintDocument1
        Me.PrintDialog1.UseEXDialog = True
        '
        'PrintDocument1
        '
        '
        'rtxEULA
        '
        Me.rtxEULA.BackColor = System.Drawing.Color.White
        Me.rtxEULA.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.rtxEULA.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rtxEULA.Location = New System.Drawing.Point(14, 16)
        Me.rtxEULA.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rtxEULA.Name = "rtxEULA"
        Me.rtxEULA.Size = New System.Drawing.Size(948, 647)
        Me.rtxEULA.TabIndex = 12
        Me.rtxEULA.Text = ""
        '
        'frmEULA
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(976, 714)
        Me.Controls.Add(Me.rtxEULA)
        Me.Controls.Add(Me.cmdPrint)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.txtSupport)
        Me.Controls.Add(Me.txtEULA)
        Me.Controls.Add(Me.txtTrial)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmEULA"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "frmEULA"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtTrial As System.Windows.Forms.TextBox
    Friend WithEvents txtEULA As System.Windows.Forms.TextBox
    Friend WithEvents txtSupport As System.Windows.Forms.TextBox
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents cmdPrint As System.Windows.Forms.Button
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    Friend WithEvents PrintDocument1 As System.Drawing.Printing.PrintDocument
    Friend WithEvents rtxEULA As Global.GooWoo.RichTextBoxPrintCtrl.RichTextBoxPrintCtrl
End Class
