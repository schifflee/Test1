<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPeriodTemp
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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.OK_Button = New System.Windows.Forms.Button
        Me.Cancel_Button = New System.Windows.Forms.Button
        Me.panCycles = New System.Windows.Forms.Panel
        Me.txtCycles = New System.Windows.Forms.TextBox
        Me.lblCycles = New System.Windows.Forms.Label
        Me.panTP = New System.Windows.Forms.Panel
        Me.txtTemp = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtTF = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtTP = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TableLayoutPanel1.SuspendLayout()
        Me.panCycles.SuspendLayout()
        Me.panTP.SuspendLayout()
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
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(277, 173)
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
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "&OK"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.ForeColor = System.Drawing.Color.Red
        Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "&Cancel"
        '
        'panCycles
        '
        Me.panCycles.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panCycles.Controls.Add(Me.txtCycles)
        Me.panCycles.Controls.Add(Me.lblCycles)
        Me.panCycles.Location = New System.Drawing.Point(12, 22)
        Me.panCycles.Name = "panCycles"
        Me.panCycles.Size = New System.Drawing.Size(171, 36)
        Me.panCycles.TabIndex = 1
        '
        'txtCycles
        '
        Me.txtCycles.Location = New System.Drawing.Point(127, 6)
        Me.txtCycles.Name = "txtCycles"
        Me.txtCycles.Size = New System.Drawing.Size(30, 20)
        Me.txtCycles.TabIndex = 1
        Me.txtCycles.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblCycles
        '
        Me.lblCycles.AutoSize = True
        Me.lblCycles.Location = New System.Drawing.Point(3, 9)
        Me.lblCycles.Name = "lblCycles"
        Me.lblCycles.Size = New System.Drawing.Size(118, 13)
        Me.lblCycles.TabIndex = 0
        Me.lblCycles.Text = "Enter number of cycles:"
        '
        'panTP
        '
        Me.panTP.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panTP.Controls.Add(Me.txtTemp)
        Me.panTP.Controls.Add(Me.Label3)
        Me.panTP.Controls.Add(Me.txtTF)
        Me.panTP.Controls.Add(Me.Label2)
        Me.panTP.Controls.Add(Me.txtTP)
        Me.panTP.Controls.Add(Me.Label1)
        Me.panTP.Location = New System.Drawing.Point(12, 74)
        Me.panTP.Name = "panTP"
        Me.panTP.Size = New System.Drawing.Size(408, 90)
        Me.panTP.TabIndex = 2
        '
        'txtTemp
        '
        Me.txtTemp.Location = New System.Drawing.Point(333, 59)
        Me.txtTemp.Name = "txtTemp"
        Me.txtTemp.Size = New System.Drawing.Size(70, 20)
        Me.txtTemp.TabIndex = 5
        Me.txtTemp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 62)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(322, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Enter temperature (e.g. -70 deg C, Room Temp, Refrigerated, etc.):"
        '
        'txtTF
        '
        Me.txtTF.Location = New System.Drawing.Point(333, 33)
        Me.txtTF.Name = "txtTF"
        Me.txtTF.Size = New System.Drawing.Size(70, 20)
        Me.txtTF.TabIndex = 3
        Me.txtTF.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 36)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(201, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Enter time frame (e.g. Hours, Days, etc.) :"
        '
        'txtTP
        '
        Me.txtTP.Location = New System.Drawing.Point(333, 9)
        Me.txtTP.Name = "txtTP"
        Me.txtTP.Size = New System.Drawing.Size(70, 20)
        Me.txtTP.TabIndex = 1
        Me.txtTP.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(176, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Enter time period (e.g. 1, 5, 10, etc):"
        '
        'frmPeriodTemp
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.MistyRose
        Me.ClientSize = New System.Drawing.Size(435, 214)
        Me.ControlBox = False
        Me.Controls.Add(Me.panTP)
        Me.Controls.Add(Me.panCycles)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPeriodTemp"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Enter Period Temp or Cycles"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.panCycles.ResumeLayout(False)
        Me.panCycles.PerformLayout()
        Me.panTP.ResumeLayout(False)
        Me.panTP.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents panCycles As System.Windows.Forms.Panel
    Friend WithEvents lblCycles As System.Windows.Forms.Label
    Friend WithEvents txtCycles As System.Windows.Forms.TextBox
    Friend WithEvents panTP As System.Windows.Forms.Panel
    Friend WithEvents txtTF As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtTP As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtTemp As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label

End Class
