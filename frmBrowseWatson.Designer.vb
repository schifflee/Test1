<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBrowseWatson
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBrowseWatson))
        Me.dgvProjects = New System.Windows.Forms.DataGridView()
        Me.dgvStudies = New System.Windows.Forms.DataGridView()
        Me.lblProjects = New System.Windows.Forms.Label()
        Me.lblStudies = New System.Windows.Forms.Label()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.txtProjectID = New System.Windows.Forms.TextBox()
        Me.txtStudyID = New System.Windows.Forms.TextBox()
        Me.panHeader = New System.Windows.Forms.Panel()
        Me.txtFilter = New System.Windows.Forms.TextBox()
        Me.cmdClear = New System.Windows.Forms.Button()
        Me.lblFilter = New System.Windows.Forms.Label()
        CType(Me.dgvProjects, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvStudies, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panHeader.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvProjects
        '
        Me.dgvProjects.AllowUserToAddRows = False
        Me.dgvProjects.AllowUserToDeleteRows = False
        Me.dgvProjects.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dgvProjects.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvProjects.BackgroundColor = System.Drawing.Color.White
        Me.dgvProjects.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvProjects.Location = New System.Drawing.Point(14, 58)
        Me.dgvProjects.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvProjects.MultiSelect = False
        Me.dgvProjects.Name = "dgvProjects"
        Me.dgvProjects.ReadOnly = True
        Me.dgvProjects.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvProjects.Size = New System.Drawing.Size(405, 777)
        Me.dgvProjects.TabIndex = 0
        '
        'dgvStudies
        '
        Me.dgvStudies.AllowUserToAddRows = False
        Me.dgvStudies.AllowUserToDeleteRows = False
        Me.dgvStudies.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvStudies.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader
        Me.dgvStudies.BackgroundColor = System.Drawing.Color.White
        Me.dgvStudies.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvStudies.Location = New System.Drawing.Point(426, 58)
        Me.dgvStudies.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvStudies.MultiSelect = False
        Me.dgvStudies.Name = "dgvStudies"
        Me.dgvStudies.ReadOnly = True
        Me.dgvStudies.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvStudies.Size = New System.Drawing.Size(764, 777)
        Me.dgvStudies.TabIndex = 1
        '
        'lblProjects
        '
        Me.lblProjects.AutoSize = True
        Me.lblProjects.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProjects.Location = New System.Drawing.Point(14, 39)
        Me.lblProjects.Name = "lblProjects"
        Me.lblProjects.Size = New System.Drawing.Size(121, 16)
        Me.lblProjects.TabIndex = 2
        Me.lblProjects.Text = "Watson Projects"
        '
        'lblStudies
        '
        Me.lblStudies.AutoSize = True
        Me.lblStudies.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStudies.Location = New System.Drawing.Point(0, 25)
        Me.lblStudies.Name = "lblStudies"
        Me.lblStudies.Size = New System.Drawing.Size(116, 16)
        Me.lblStudies.TabIndex = 3
        Me.lblStudies.Text = "Watson Studies"
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.Color.DarkGray
        Me.cmdOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdOK.Location = New System.Drawing.Point(123, 4)
        Me.cmdOK.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(97, 35)
        Me.cmdOK.TabIndex = 92
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancel.CausesValidation = False
        Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(227, 4)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(97, 35)
        Me.cmdCancel.TabIndex = 93
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'txtProjectID
        '
        Me.txtProjectID.Location = New System.Drawing.Point(97, 9)
        Me.txtProjectID.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtProjectID.Name = "txtProjectID"
        Me.txtProjectID.Size = New System.Drawing.Size(111, 25)
        Me.txtProjectID.TabIndex = 94
        Me.txtProjectID.Visible = False
        '
        'txtStudyID
        '
        Me.txtStudyID.Location = New System.Drawing.Point(227, 10)
        Me.txtStudyID.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtStudyID.Name = "txtStudyID"
        Me.txtStudyID.Size = New System.Drawing.Size(111, 25)
        Me.txtStudyID.TabIndex = 95
        Me.txtStudyID.Visible = False
        '
        'panHeader
        '
        Me.panHeader.Controls.Add(Me.txtFilter)
        Me.panHeader.Controls.Add(Me.cmdClear)
        Me.panHeader.Controls.Add(Me.lblFilter)
        Me.panHeader.Controls.Add(Me.cmdCancel)
        Me.panHeader.Controls.Add(Me.lblStudies)
        Me.panHeader.Controls.Add(Me.cmdOK)
        Me.panHeader.Location = New System.Drawing.Point(426, 14)
        Me.panHeader.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panHeader.Name = "panHeader"
        Me.panHeader.Size = New System.Drawing.Size(743, 44)
        Me.panHeader.TabIndex = 96
        '
        'txtFilter
        '
        Me.txtFilter.Location = New System.Drawing.Point(472, 14)
        Me.txtFilter.Name = "txtFilter"
        Me.txtFilter.Size = New System.Drawing.Size(135, 25)
        Me.txtFilter.TabIndex = 95
        '
        'cmdClear
        '
        Me.cmdClear.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdClear.CausesValidation = False
        Me.cmdClear.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdClear.ForeColor = System.Drawing.Color.Red
        Me.cmdClear.Location = New System.Drawing.Point(613, 14)
        Me.cmdClear.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdClear.Name = "cmdClear"
        Me.cmdClear.Size = New System.Drawing.Size(98, 25)
        Me.cmdClear.TabIndex = 96
        Me.cmdClear.Text = "C&lear Filter"
        Me.cmdClear.UseVisualStyleBackColor = False
        '
        'lblFilter
        '
        Me.lblFilter.AutoSize = True
        Me.lblFilter.Location = New System.Drawing.Point(339, 8)
        Me.lblFilter.Name = "lblFilter"
        Me.lblFilter.Size = New System.Drawing.Size(135, 34)
        Me.lblFilter.TabIndex = 94
        Me.lblFilter.Text = "Filter for Study Name:" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(wildcard built in)"
        '
        'frmBrowseWatson
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1204, 850)
        Me.Controls.Add(Me.panHeader)
        Me.Controls.Add(Me.txtStudyID)
        Me.Controls.Add(Me.txtProjectID)
        Me.Controls.Add(Me.lblProjects)
        Me.Controls.Add(Me.dgvStudies)
        Me.Controls.Add(Me.dgvProjects)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmBrowseWatson"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Retrieve Watson Study..."
        CType(Me.dgvProjects, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvStudies, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panHeader.ResumeLayout(False)
        Me.panHeader.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvProjects As System.Windows.Forms.DataGridView
    Friend WithEvents dgvStudies As System.Windows.Forms.DataGridView
    Friend WithEvents lblProjects As System.Windows.Forms.Label
    Friend WithEvents lblStudies As System.Windows.Forms.Label
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents txtProjectID As System.Windows.Forms.TextBox
    Friend WithEvents txtStudyID As System.Windows.Forms.TextBox
    Friend WithEvents panHeader As System.Windows.Forms.Panel
    Friend WithEvents txtFilter As System.Windows.Forms.TextBox
    Friend WithEvents lblFilter As System.Windows.Forms.Label
    Friend WithEvents cmdClear As System.Windows.Forms.Button
End Class
