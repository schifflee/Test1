<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUpdateCheck
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmUpdateCheck))
        Me.pan1 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.pgTable = New System.Windows.Forms.ProgressBar()
        Me.pgOverall = New System.Windows.Forms.ProgressBar()
        Me.panInsert = New System.Windows.Forms.Panel()
        Me.lblInsert = New System.Windows.Forms.Label()
        Me.pgInsert = New System.Windows.Forms.ProgressBar()
        Me.panIndex = New System.Windows.Forms.Panel()
        Me.lblIndex = New System.Windows.Forms.Label()
        Me.pgIndex = New System.Windows.Forms.ProgressBar()
        Me.panCreate = New System.Windows.Forms.Panel()
        Me.lblCreate = New System.Windows.Forms.Label()
        Me.pgCreate = New System.Windows.Forms.ProgressBar()
        Me.lblTable = New System.Windows.Forms.Label()
        Me.lbl1 = New System.Windows.Forms.Label()
        Me.cmdYes = New System.Windows.Forms.Button()
        Me.cmdNo = New System.Windows.Forms.Button()
        Me.pan1.SuspendLayout()
        Me.panInsert.SuspendLayout()
        Me.panIndex.SuspendLayout()
        Me.panCreate.SuspendLayout()
        Me.SuspendLayout()
        '
        'pan1
        '
        Me.pan1.Controls.Add(Me.Label1)
        Me.pan1.Controls.Add(Me.pgTable)
        Me.pan1.Controls.Add(Me.pgOverall)
        Me.pan1.Controls.Add(Me.panInsert)
        Me.pan1.Controls.Add(Me.panIndex)
        Me.pan1.Controls.Add(Me.panCreate)
        Me.pan1.Controls.Add(Me.lblTable)
        Me.pan1.Location = New System.Drawing.Point(92, 489)
        Me.pan1.Name = "pan1"
        Me.pan1.Size = New System.Drawing.Size(580, 82)
        Me.pan1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.Label1.Location = New System.Drawing.Point(228, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(105, 17)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Overall Progress"
        '
        'pgTable
        '
        Me.pgTable.Location = New System.Drawing.Point(6, 85)
        Me.pgTable.Name = "pgTable"
        Me.pgTable.Size = New System.Drawing.Size(563, 22)
        Me.pgTable.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.pgTable.TabIndex = 8
        Me.pgTable.Visible = False
        '
        'pgOverall
        '
        Me.pgOverall.Location = New System.Drawing.Point(6, 33)
        Me.pgOverall.Name = "pgOverall"
        Me.pgOverall.Size = New System.Drawing.Size(563, 22)
        Me.pgOverall.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.pgOverall.TabIndex = 7
        '
        'panInsert
        '
        Me.panInsert.Controls.Add(Me.lblInsert)
        Me.panInsert.Controls.Add(Me.pgInsert)
        Me.panInsert.Location = New System.Drawing.Point(3, 217)
        Me.panInsert.Name = "panInsert"
        Me.panInsert.Size = New System.Drawing.Size(571, 46)
        Me.panInsert.TabIndex = 6
        Me.panInsert.Visible = False
        '
        'lblInsert
        '
        Me.lblInsert.AutoSize = True
        Me.lblInsert.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInsert.Location = New System.Drawing.Point(4, 3)
        Me.lblInsert.Name = "lblInsert"
        Me.lblInsert.Size = New System.Drawing.Size(113, 16)
        Me.lblInsert.TabIndex = 8
        Me.lblInsert.Text = "Inserting Records"
        '
        'pgInsert
        '
        Me.pgInsert.Location = New System.Drawing.Point(4, 21)
        Me.pgInsert.Name = "pgInsert"
        Me.pgInsert.Size = New System.Drawing.Size(563, 22)
        Me.pgInsert.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.pgInsert.TabIndex = 7
        '
        'panIndex
        '
        Me.panIndex.Controls.Add(Me.lblIndex)
        Me.panIndex.Controls.Add(Me.pgIndex)
        Me.panIndex.Location = New System.Drawing.Point(3, 165)
        Me.panIndex.Name = "panIndex"
        Me.panIndex.Size = New System.Drawing.Size(571, 46)
        Me.panIndex.TabIndex = 5
        Me.panIndex.Visible = False
        '
        'lblIndex
        '
        Me.lblIndex.AutoSize = True
        Me.lblIndex.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblIndex.Location = New System.Drawing.Point(4, 3)
        Me.lblIndex.Name = "lblIndex"
        Me.lblIndex.Size = New System.Drawing.Size(108, 16)
        Me.lblIndex.TabIndex = 7
        Me.lblIndex.Text = "Creating Indexes"
        '
        'pgIndex
        '
        Me.pgIndex.Location = New System.Drawing.Point(4, 21)
        Me.pgIndex.Name = "pgIndex"
        Me.pgIndex.Size = New System.Drawing.Size(563, 22)
        Me.pgIndex.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.pgIndex.TabIndex = 6
        '
        'panCreate
        '
        Me.panCreate.Controls.Add(Me.lblCreate)
        Me.panCreate.Controls.Add(Me.pgCreate)
        Me.panCreate.Location = New System.Drawing.Point(3, 113)
        Me.panCreate.Name = "panCreate"
        Me.panCreate.Size = New System.Drawing.Size(571, 46)
        Me.panCreate.TabIndex = 4
        Me.panCreate.Visible = False
        '
        'lblCreate
        '
        Me.lblCreate.AutoSize = True
        Me.lblCreate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCreate.Location = New System.Drawing.Point(4, 3)
        Me.lblCreate.Name = "lblCreate"
        Me.lblCreate.Size = New System.Drawing.Size(163, 16)
        Me.lblCreate.TabIndex = 4
        Me.lblCreate.Text = "Creating Table and Fields"
        '
        'pgCreate
        '
        Me.pgCreate.Location = New System.Drawing.Point(3, 21)
        Me.pgCreate.Name = "pgCreate"
        Me.pgCreate.Size = New System.Drawing.Size(563, 22)
        Me.pgCreate.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.pgCreate.TabIndex = 3
        '
        'lblTable
        '
        Me.lblTable.AutoSize = True
        Me.lblTable.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTable.Location = New System.Drawing.Point(7, 66)
        Me.lblTable.Name = "lblTable"
        Me.lblTable.Size = New System.Drawing.Size(110, 16)
        Me.lblTable.TabIndex = 0
        Me.lblTable.Text = "Evaluating table: "
        Me.lblTable.Visible = False
        '
        'lbl1
        '
        Me.lbl1.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl1.Location = New System.Drawing.Point(11, 9)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(753, 426)
        Me.lbl1.TabIndex = 1
        Me.lbl1.Text = "lbl1"
        Me.lbl1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'cmdYes
        '
        Me.cmdYes.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdYes.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.cmdYes.FlatAppearance.BorderSize = 0
        Me.cmdYes.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdYes.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdYes.ForeColor = System.Drawing.Color.Blue
        Me.cmdYes.Location = New System.Drawing.Point(312, 451)
        Me.cmdYes.Name = "cmdYes"
        Me.cmdYes.Size = New System.Drawing.Size(61, 31)
        Me.cmdYes.TabIndex = 2
        Me.cmdYes.Text = "Yes"
        Me.cmdYes.UseVisualStyleBackColor = True
        '
        'cmdNo
        '
        Me.cmdNo.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdNo.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.cmdNo.FlatAppearance.BorderSize = 0
        Me.cmdNo.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdNo.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdNo.ForeColor = System.Drawing.Color.Red
        Me.cmdNo.Location = New System.Drawing.Point(411, 452)
        Me.cmdNo.Name = "cmdNo"
        Me.cmdNo.Size = New System.Drawing.Size(61, 31)
        Me.cmdNo.TabIndex = 3
        Me.cmdNo.Text = "No"
        Me.cmdNo.UseVisualStyleBackColor = True
        '
        'frmUpdateCheck
        '
        Me.AcceptButton = Me.cmdYes
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(776, 596)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmdNo)
        Me.Controls.Add(Me.cmdYes)
        Me.Controls.Add(Me.lbl1)
        Me.Controls.Add(Me.pan1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmUpdateCheck"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Update StudyDoc"
        Me.pan1.ResumeLayout(False)
        Me.pan1.PerformLayout()
        Me.panInsert.ResumeLayout(False)
        Me.panInsert.PerformLayout()
        Me.panIndex.ResumeLayout(False)
        Me.panIndex.PerformLayout()
        Me.panCreate.ResumeLayout(False)
        Me.panCreate.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pan1 As System.Windows.Forms.Panel
    Friend WithEvents lblTable As System.Windows.Forms.Label
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Friend WithEvents cmdYes As System.Windows.Forms.Button
    Friend WithEvents panInsert As System.Windows.Forms.Panel
    Friend WithEvents lblInsert As System.Windows.Forms.Label
    Friend WithEvents pgInsert As System.Windows.Forms.ProgressBar
    Friend WithEvents panIndex As System.Windows.Forms.Panel
    Friend WithEvents lblIndex As System.Windows.Forms.Label
    Friend WithEvents pgIndex As System.Windows.Forms.ProgressBar
    Friend WithEvents panCreate As System.Windows.Forms.Panel
    Friend WithEvents lblCreate As System.Windows.Forms.Label
    Friend WithEvents pgCreate As System.Windows.Forms.ProgressBar
    Friend WithEvents pgOverall As System.Windows.Forms.ProgressBar
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents pgTable As System.Windows.Forms.ProgressBar
    Friend WithEvents cmdNo As System.Windows.Forms.Button
End Class
