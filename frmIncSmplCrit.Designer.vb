<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmIncSmplCrit
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmIncSmplCrit))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.gbChoice = New System.Windows.Forms.GroupBox()
        Me.rb2 = New System.Windows.Forms.RadioButton()
        Me.rb1 = New System.Windows.Forms.RadioButton()
        Me.NUMISCRIT1 = New System.Windows.Forms.TextBox()
        Me.NUMISCRIT1LEVEL = New System.Windows.Forms.TextBox()
        Me.NUMISCRIT2 = New System.Windows.Forms.TextBox()
        Me.lblNUMISCRIT1 = New System.Windows.Forms.Label()
        Me.lblNUMISCRIT1LEVEL = New System.Windows.Forms.Label()
        Me.lblNUMISCRIT2 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.gbChoice.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(233, 174)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.ForeColor = System.Drawing.Color.Blue
        Me.OK_Button.Location = New System.Drawing.Point(3, 3)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 23)
        Me.OK_Button.TabIndex = 3
        Me.OK_Button.Text = "OK"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.CausesValidation = False
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.ForeColor = System.Drawing.Color.Red
        Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
        Me.Cancel_Button.TabIndex = 4
        Me.Cancel_Button.Text = "Cancel"
        '
        'gbChoice
        '
        Me.gbChoice.Controls.Add(Me.rb2)
        Me.gbChoice.Controls.Add(Me.rb1)
        Me.gbChoice.Location = New System.Drawing.Point(12, 12)
        Me.gbChoice.Name = "gbChoice"
        Me.gbChoice.Size = New System.Drawing.Size(209, 47)
        Me.gbChoice.TabIndex = 1
        Me.gbChoice.TabStop = False
        Me.gbChoice.Text = "Choose Number of levels to evaluate"
        '
        'rb2
        '
        Me.rb2.AutoSize = True
        Me.rb2.CausesValidation = False
        Me.rb2.Location = New System.Drawing.Point(98, 19)
        Me.rb2.Name = "rb2"
        Me.rb2.Size = New System.Drawing.Size(80, 17)
        Me.rb2.TabIndex = 1
        Me.rb2.Text = "Two Levels"
        Me.rb2.UseVisualStyleBackColor = True
        '
        'rb1
        '
        Me.rb1.AutoSize = True
        Me.rb1.CausesValidation = False
        Me.rb1.Checked = True
        Me.rb1.Location = New System.Drawing.Point(6, 19)
        Me.rb1.Name = "rb1"
        Me.rb1.Size = New System.Drawing.Size(74, 17)
        Me.rb1.TabIndex = 0
        Me.rb1.TabStop = True
        Me.rb1.Text = "One Level"
        Me.rb1.UseVisualStyleBackColor = True
        '
        'NUMISCRIT1
        '
        Me.NUMISCRIT1.Location = New System.Drawing.Point(296, 84)
        Me.NUMISCRIT1.Name = "NUMISCRIT1"
        Me.NUMISCRIT1.Size = New System.Drawing.Size(83, 20)
        Me.NUMISCRIT1.TabIndex = 0
        Me.NUMISCRIT1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'NUMISCRIT1LEVEL
        '
        Me.NUMISCRIT1LEVEL.Location = New System.Drawing.Point(296, 110)
        Me.NUMISCRIT1LEVEL.Name = "NUMISCRIT1LEVEL"
        Me.NUMISCRIT1LEVEL.Size = New System.Drawing.Size(83, 20)
        Me.NUMISCRIT1LEVEL.TabIndex = 1
        Me.NUMISCRIT1LEVEL.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.NUMISCRIT1LEVEL.Visible = False
        '
        'NUMISCRIT2
        '
        Me.NUMISCRIT2.Location = New System.Drawing.Point(296, 136)
        Me.NUMISCRIT2.Name = "NUMISCRIT2"
        Me.NUMISCRIT2.Size = New System.Drawing.Size(83, 20)
        Me.NUMISCRIT2.TabIndex = 2
        Me.NUMISCRIT2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.NUMISCRIT2.Visible = False
        '
        'lblNUMISCRIT1
        '
        Me.lblNUMISCRIT1.AutoSize = True
        Me.lblNUMISCRIT1.Location = New System.Drawing.Point(9, 87)
        Me.lblNUMISCRIT1.Name = "lblNUMISCRIT1"
        Me.lblNUMISCRIT1.Size = New System.Drawing.Size(280, 13)
        Me.lblNUMISCRIT1.TabIndex = 5
        Me.lblNUMISCRIT1.Text = "Enter Level One Incurred Sample Acceptance Criteria (%):"
        '
        'lblNUMISCRIT1LEVEL
        '
        Me.lblNUMISCRIT1LEVEL.AutoSize = True
        Me.lblNUMISCRIT1LEVEL.Location = New System.Drawing.Point(9, 113)
        Me.lblNUMISCRIT1LEVEL.Name = "lblNUMISCRIT1LEVEL"
        Me.lblNUMISCRIT1LEVEL.Size = New System.Drawing.Size(202, 13)
        Me.lblNUMISCRIT1LEVEL.TabIndex = 6
        Me.lblNUMISCRIT1LEVEL.Text = "Enter % of BQL for Level One Evaluation:"
        Me.lblNUMISCRIT1LEVEL.Visible = False
        '
        'lblNUMISCRIT2
        '
        Me.lblNUMISCRIT2.AutoSize = True
        Me.lblNUMISCRIT2.Location = New System.Drawing.Point(9, 139)
        Me.lblNUMISCRIT2.Name = "lblNUMISCRIT2"
        Me.lblNUMISCRIT2.Size = New System.Drawing.Size(281, 13)
        Me.lblNUMISCRIT2.TabIndex = 7
        Me.lblNUMISCRIT2.Text = "Enter Level Two Incurred Sample Acceptance Criteria (%):"
        Me.lblNUMISCRIT2.Visible = False
        '
        'frmIncSmplCrit
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(403, 215)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblNUMISCRIT2)
        Me.Controls.Add(Me.lblNUMISCRIT1LEVEL)
        Me.Controls.Add(Me.lblNUMISCRIT1)
        Me.Controls.Add(Me.NUMISCRIT2)
        Me.Controls.Add(Me.NUMISCRIT1LEVEL)
        Me.Controls.Add(Me.NUMISCRIT1)
        Me.Controls.Add(Me.gbChoice)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmIncSmplCrit"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "  Enter Incurred Sample Criteria..."
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.gbChoice.ResumeLayout(False)
        Me.gbChoice.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents gbChoice As System.Windows.Forms.GroupBox
    Friend WithEvents rb2 As System.Windows.Forms.RadioButton
    Friend WithEvents rb1 As System.Windows.Forms.RadioButton
    Friend WithEvents NUMISCRIT1 As System.Windows.Forms.TextBox
    Friend WithEvents NUMISCRIT1LEVEL As System.Windows.Forms.TextBox
    Friend WithEvents NUMISCRIT2 As System.Windows.Forms.TextBox
    Friend WithEvents lblNUMISCRIT1 As System.Windows.Forms.Label
    Friend WithEvents lblNUMISCRIT1LEVEL As System.Windows.Forms.Label
    Friend WithEvents lblNUMISCRIT2 As System.Windows.Forms.Label

End Class
