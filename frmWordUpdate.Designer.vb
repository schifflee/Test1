<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWordUpdate
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmWordUpdate))
        Me.pan1 = New System.Windows.Forms.Panel
        'Me.afr1 = New AxDSOFramer.AxFramerControl
        Me.cmdBrowse = New System.Windows.Forms.Button
        Me.cmdWordStatements = New System.Windows.Forms.Button
        Me.cmdReportHeaders = New System.Windows.Forms.Button
        Me.lblStatus = New System.Windows.Forms.Label
        Me.pan2 = New System.Windows.Forms.Panel
        Me.wb1 = New System.Windows.Forms.WebBrowser
        Me.cmdReportStatements = New System.Windows.Forms.Button
        Me.lblPath = New System.Windows.Forms.Label
        Me.cmdDirectories = New System.Windows.Forms.Button
        Me.cmdStoreXML = New System.Windows.Forms.Button
        Me.cmdFileStream = New System.Windows.Forms.Button
        Me.cmdRetrieveDB = New System.Windows.Forms.Button
        Me.cmdIndDBUpdate = New System.Windows.Forms.Button
        Me.cmdPopulateBLOB = New System.Windows.Forms.Button
        Me.chkJustView = New System.Windows.Forms.CheckBox
        Me.pan1.SuspendLayout()
        CType(Me.afr1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pan2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pan1
        '
        Me.pan1.Controls.Add(Me.afr1)
        Me.pan1.Location = New System.Drawing.Point(214, 0)
        Me.pan1.Name = "pan1"
        Me.pan1.Size = New System.Drawing.Size(534, 738)
        Me.pan1.TabIndex = 1
        '
        'afr1
        '
        Me.afr1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.afr1.Enabled = True
        Me.afr1.Location = New System.Drawing.Point(0, 0)
        Me.afr1.Name = "afr1"
        Me.afr1.OcxState = CType(resources.GetObject("afr1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.afr1.Size = New System.Drawing.Size(534, 738)
        Me.afr1.TabIndex = 1
        '
        'cmdBrowse
        '
        Me.cmdBrowse.Location = New System.Drawing.Point(12, 6)
        Me.cmdBrowse.Name = "cmdBrowse"
        Me.cmdBrowse.Size = New System.Drawing.Size(114, 44)
        Me.cmdBrowse.TabIndex = 3
        Me.cmdBrowse.Text = "Browse..."
        Me.cmdBrowse.UseVisualStyleBackColor = True
        '
        'cmdWordStatements
        '
        Me.cmdWordStatements.Location = New System.Drawing.Point(12, 106)
        Me.cmdWordStatements.Name = "cmdWordStatements"
        Me.cmdWordStatements.Size = New System.Drawing.Size(114, 44)
        Me.cmdWordStatements.TabIndex = 4
        Me.cmdWordStatements.Text = "WordStatements"
        Me.cmdWordStatements.UseVisualStyleBackColor = True
        '
        'cmdReportHeaders
        '
        Me.cmdReportHeaders.Location = New System.Drawing.Point(12, 156)
        Me.cmdReportHeaders.Name = "cmdReportHeaders"
        Me.cmdReportHeaders.Size = New System.Drawing.Size(114, 44)
        Me.cmdReportHeaders.TabIndex = 5
        Me.cmdReportHeaders.Text = "Report Headers"
        Me.cmdReportHeaders.UseVisualStyleBackColor = True
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(24, 628)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(46, 13)
        Me.lblStatus.TabIndex = 6
        Me.lblStatus.Text = "Status..."
        '
        'pan2
        '
        Me.pan2.Controls.Add(Me.wb1)
        Me.pan2.Location = New System.Drawing.Point(754, 225)
        Me.pan2.Name = "pan2"
        Me.pan2.Size = New System.Drawing.Size(151, 148)
        Me.pan2.TabIndex = 7
        '
        'wb1
        '
        Me.wb1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.wb1.Location = New System.Drawing.Point(0, 0)
        Me.wb1.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wb1.Name = "wb1"
        Me.wb1.Size = New System.Drawing.Size(151, 148)
        Me.wb1.TabIndex = 0
        '
        'cmdReportStatements
        '
        Me.cmdReportStatements.Location = New System.Drawing.Point(12, 206)
        Me.cmdReportStatements.Name = "cmdReportStatements"
        Me.cmdReportStatements.Size = New System.Drawing.Size(114, 44)
        Me.cmdReportStatements.TabIndex = 8
        Me.cmdReportStatements.Text = "Report Statements"
        Me.cmdReportStatements.UseVisualStyleBackColor = True
        '
        'lblPath
        '
        Me.lblPath.AutoSize = True
        Me.lblPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPath.Location = New System.Drawing.Point(799, 76)
        Me.lblPath.Name = "lblPath"
        Me.lblPath.Size = New System.Drawing.Size(38, 18)
        Me.lblPath.TabIndex = 9
        Me.lblPath.Text = "Path"
        '
        'cmdDirectories
        '
        Me.cmdDirectories.Location = New System.Drawing.Point(754, 490)
        Me.cmdDirectories.Name = "cmdDirectories"
        Me.cmdDirectories.Size = New System.Drawing.Size(114, 44)
        Me.cmdDirectories.TabIndex = 10
        Me.cmdDirectories.Text = "Statement Directories"
        Me.cmdDirectories.UseVisualStyleBackColor = True
        Me.cmdDirectories.Visible = False
        '
        'cmdStoreXML
        '
        Me.cmdStoreXML.Location = New System.Drawing.Point(754, 540)
        Me.cmdStoreXML.Name = "cmdStoreXML"
        Me.cmdStoreXML.Size = New System.Drawing.Size(114, 44)
        Me.cmdStoreXML.TabIndex = 11
        Me.cmdStoreXML.Text = "StoreXML"
        Me.cmdStoreXML.UseVisualStyleBackColor = True
        Me.cmdStoreXML.Visible = False
        '
        'cmdFileStream
        '
        Me.cmdFileStream.Location = New System.Drawing.Point(12, 302)
        Me.cmdFileStream.Name = "cmdFileStream"
        Me.cmdFileStream.Size = New System.Drawing.Size(114, 44)
        Me.cmdFileStream.TabIndex = 12
        Me.cmdFileStream.Text = "File Stream"
        Me.cmdFileStream.UseVisualStyleBackColor = True
        '
        'cmdRetrieveDB
        '
        Me.cmdRetrieveDB.Location = New System.Drawing.Point(12, 422)
        Me.cmdRetrieveDB.Name = "cmdRetrieveDB"
        Me.cmdRetrieveDB.Size = New System.Drawing.Size(114, 44)
        Me.cmdRetrieveDB.TabIndex = 13
        Me.cmdRetrieveDB.Text = "Retrieve from DB"
        Me.cmdRetrieveDB.UseVisualStyleBackColor = True
        '
        'cmdIndDBUpdate
        '
        Me.cmdIndDBUpdate.Location = New System.Drawing.Point(12, 352)
        Me.cmdIndDBUpdate.Name = "cmdIndDBUpdate"
        Me.cmdIndDBUpdate.Size = New System.Drawing.Size(114, 44)
        Me.cmdIndDBUpdate.TabIndex = 14
        Me.cmdIndDBUpdate.Text = "Individual DB Update"
        Me.cmdIndDBUpdate.UseVisualStyleBackColor = True
        '
        'cmdPopulateBLOB
        '
        Me.cmdPopulateBLOB.Location = New System.Drawing.Point(12, 472)
        Me.cmdPopulateBLOB.Name = "cmdPopulateBLOB"
        Me.cmdPopulateBLOB.Size = New System.Drawing.Size(114, 44)
        Me.cmdPopulateBLOB.TabIndex = 15
        Me.cmdPopulateBLOB.Text = "Populate BLOB table"
        Me.cmdPopulateBLOB.UseVisualStyleBackColor = True
        '
        'chkJustView
        '
        Me.chkJustView.AutoSize = True
        Me.chkJustView.Location = New System.Drawing.Point(27, 399)
        Me.chkJustView.Name = "chkJustView"
        Me.chkJustView.Size = New System.Drawing.Size(71, 17)
        Me.chkJustView.TabIndex = 16
        Me.chkJustView.Text = "Just View"
        Me.chkJustView.UseVisualStyleBackColor = True
        '
        'frmWordUpdate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(976, 736)
        Me.Controls.Add(Me.chkJustView)
        Me.Controls.Add(Me.cmdPopulateBLOB)
        Me.Controls.Add(Me.cmdIndDBUpdate)
        Me.Controls.Add(Me.cmdRetrieveDB)
        Me.Controls.Add(Me.cmdFileStream)
        Me.Controls.Add(Me.cmdStoreXML)
        Me.Controls.Add(Me.cmdDirectories)
        Me.Controls.Add(Me.lblPath)
        Me.Controls.Add(Me.cmdReportStatements)
        Me.Controls.Add(Me.pan2)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.cmdReportHeaders)
        Me.Controls.Add(Me.cmdWordStatements)
        Me.Controls.Add(Me.cmdBrowse)
        Me.Controls.Add(Me.pan1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmWordUpdate"
        Me.Text = "Update Word DataStore Components"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pan1.ResumeLayout(False)
        CType(Me.afr1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pan2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents pan1 As System.Windows.Forms.Panel
    'Friend WithEvents afr1 As AxDSOFramer.AxFramerControl
    Friend WithEvents cmdBrowse As System.Windows.Forms.Button
    Friend WithEvents cmdWordStatements As System.Windows.Forms.Button
    Friend WithEvents cmdReportHeaders As System.Windows.Forms.Button
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents pan2 As System.Windows.Forms.Panel
    Friend WithEvents cmdReportStatements As System.Windows.Forms.Button
    Friend WithEvents lblPath As System.Windows.Forms.Label
    Friend WithEvents cmdDirectories As System.Windows.Forms.Button
    Friend WithEvents cmdStoreXML As System.Windows.Forms.Button
    Friend WithEvents wb1 As System.Windows.Forms.WebBrowser
    Friend WithEvents cmdFileStream As System.Windows.Forms.Button
    Friend WithEvents cmdRetrieveDB As System.Windows.Forms.Button
    Friend WithEvents cmdIndDBUpdate As System.Windows.Forms.Button
    Friend WithEvents cmdPopulateBLOB As System.Windows.Forms.Button
    Friend WithEvents chkJustView As System.Windows.Forms.CheckBox
End Class
