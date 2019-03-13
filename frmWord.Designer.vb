<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWord
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmWord))
        Me.cmdOpen = New System.Windows.Forms.Button
        Me.panWord = New System.Windows.Forms.Panel
        Me.afrWord = New AxDSOFramer.AxFramerControl
        Me.cmdFieldCode = New System.Windows.Forms.Button
        Me.cmdExit = New System.Windows.Forms.Button
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmsfrmWord = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.cmiFieldCode = New System.Windows.Forms.ToolStripMenuItem
        Me.panWdWB = New System.Windows.Forms.Panel
        Me.wbFrmWd = New System.Windows.Forms.WebBrowser
        Me.cmdRefresh = New System.Windows.Forms.Button
        Me.panWord.SuspendLayout()
        CType(Me.afrWord, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.cmsfrmWord.SuspendLayout()
        Me.panWdWB.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdOpen
        '
        Me.cmdOpen.Location = New System.Drawing.Point(12, 500)
        Me.cmdOpen.Name = "cmdOpen"
        Me.cmdOpen.Size = New System.Drawing.Size(91, 36)
        Me.cmdOpen.TabIndex = 1
        Me.cmdOpen.Text = "&Open"
        Me.cmdOpen.UseVisualStyleBackColor = True
        Me.cmdOpen.Visible = False
        '
        'panWord
        '
        Me.panWord.Controls.Add(Me.afrWord)
        Me.panWord.Location = New System.Drawing.Point(128, 7)
        Me.panWord.Name = "panWord"
        Me.panWord.Size = New System.Drawing.Size(700, 465)
        Me.panWord.TabIndex = 2
        Me.panWord.Visible = False
        '
        'afrWord
        '
        Me.afrWord.Dock = System.Windows.Forms.DockStyle.Fill
        Me.afrWord.Enabled = True
        Me.afrWord.Location = New System.Drawing.Point(0, 0)
        Me.afrWord.Name = "afrWord"
        Me.afrWord.OcxState = CType(resources.GetObject("afrWord.OcxState"), System.Windows.Forms.AxHost.State)
        Me.afrWord.Size = New System.Drawing.Size(700, 465)
        Me.afrWord.TabIndex = 1
        Me.afrWord.TabStop = False
        Me.afrWord.Visible = False
        '
        'cmdFieldCode
        '
        Me.cmdFieldCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdFieldCode.ForeColor = System.Drawing.Color.Green
        Me.cmdFieldCode.Location = New System.Drawing.Point(12, 7)
        Me.cmdFieldCode.Name = "cmdFieldCode"
        Me.cmdFieldCode.Size = New System.Drawing.Size(91, 47)
        Me.cmdFieldCode.TabIndex = 0
        Me.cmdFieldCode.Text = "Enter &Field Code..."
        Me.cmdFieldCode.UseVisualStyleBackColor = True
        '
        'cmdExit
        '
        Me.cmdExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ForeColor = System.Drawing.Color.Red
        Me.cmdExit.Location = New System.Drawing.Point(12, 195)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(91, 47)
        Me.cmdExit.TabIndex = 2
        Me.cmdExit.Text = "&Cancel"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'cmdSave
        '
        Me.cmdSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ForeColor = System.Drawing.Color.Blue
        Me.cmdSave.Location = New System.Drawing.Point(12, 60)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(91, 47)
        Me.cmdSave.TabIndex = 1
        Me.cmdSave.Text = "&Save and Close"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'cmsfrmWord
        '
        Me.cmsfrmWord.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.cmiFieldCode})
        Me.cmsfrmWord.Name = "cmsfrmWord"
        Me.cmsfrmWord.Size = New System.Drawing.Size(169, 26)
        '
        'cmiFieldCode
        '
        Me.cmiFieldCode.Name = "cmiFieldCode"
        Me.cmiFieldCode.Size = New System.Drawing.Size(168, 22)
        Me.cmiFieldCode.Text = "Insert Field Code..."
        '
        'panWdWB
        '
        Me.panWdWB.Controls.Add(Me.wbFrmWd)
        Me.panWdWB.Location = New System.Drawing.Point(320, 500)
        Me.panWdWB.Name = "panWdWB"
        Me.panWdWB.Size = New System.Drawing.Size(399, 137)
        Me.panWdWB.TabIndex = 3
        '
        'wbFrmWd
        '
        Me.wbFrmWd.Dock = System.Windows.Forms.DockStyle.Fill
        Me.wbFrmWd.Location = New System.Drawing.Point(0, 0)
        Me.wbFrmWd.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbFrmWd.Name = "wbFrmWd"
        Me.wbFrmWd.Size = New System.Drawing.Size(399, 137)
        Me.wbFrmWd.TabIndex = 0
        '
        'cmdRefresh
        '
        Me.cmdRefresh.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdRefresh.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRefresh.ForeColor = System.Drawing.Color.Purple
        Me.cmdRefresh.Location = New System.Drawing.Point(12, 113)
        Me.cmdRefresh.Name = "cmdRefresh"
        Me.cmdRefresh.Size = New System.Drawing.Size(91, 47)
        Me.cmdRefresh.TabIndex = 4
        Me.cmdRefresh.Text = "&Refresh"
        Me.cmdRefresh.UseVisualStyleBackColor = True
        '
        'frmWord
        '
        Me.AcceptButton = Me.cmdSave
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.CancelButton = Me.cmdExit
        Me.ClientSize = New System.Drawing.Size(928, 711)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmdRefresh)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.panWdWB)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdFieldCode)
        Me.Controls.Add(Me.panWord)
        Me.Controls.Add(Me.cmdOpen)
        Me.Name = "frmWord"
        Me.ShowInTaskbar = False
        Me.Text = "ReportStatement"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.panWord.ResumeLayout(False)
        CType(Me.afrWord, System.ComponentModel.ISupportInitialize).EndInit()
        Me.cmsfrmWord.ResumeLayout(False)
        Me.panWdWB.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdOpen As System.Windows.Forms.Button
    Friend WithEvents panWord As System.Windows.Forms.Panel
    Friend WithEvents afrWord As AxDSOFramer.AxFramerControl
    Friend WithEvents cmdFieldCode As System.Windows.Forms.Button
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmsfrmWord As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents cmiFieldCode As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents panWdWB As System.Windows.Forms.Panel
    Friend WithEvents wbFrmWd As System.Windows.Forms.WebBrowser
    Friend WithEvents cmdRefresh As System.Windows.Forms.Button
End Class
