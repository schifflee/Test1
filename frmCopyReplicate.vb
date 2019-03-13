Option Compare Text

Public Class frmCopyReplicate
    Inherits System.Windows.Forms.Form
    Public boolStop As Boolean = True

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents lblAddRep As System.Windows.Forms.Label
    Friend WithEvents lbxAnalytesCopiedFrom As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lbxAnalytesCopiedTo As System.Windows.Forms.ListBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCopyReplicate))
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.lblAddRep = New System.Windows.Forms.Label()
        Me.lbxAnalytesCopiedFrom = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lbxAnalytesCopiedTo = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(240, 250)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(88, 40)
        Me.cmdCancel.TabIndex = 7
        Me.cmdCancel.Text = "&Cancel"
        '
        'cmdOK
        '
        Me.cmdOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.Blue
        Me.cmdOK.Location = New System.Drawing.Point(136, 250)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(88, 40)
        Me.cmdOK.TabIndex = 6
        Me.cmdOK.Text = "&OK"
        '
        'lblAddRep
        '
        Me.lblAddRep.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAddRep.Location = New System.Drawing.Point(24, 16)
        Me.lblAddRep.Name = "lblAddRep"
        Me.lblAddRep.Size = New System.Drawing.Size(204, 47)
        Me.lblAddRep.TabIndex = 5
        Me.lblAddRep.Text = "Choose an analyte from which data will be copied:"
        Me.lblAddRep.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'lbxAnalytesCopiedFrom
        '
        Me.lbxAnalytesCopiedFrom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbxAnalytesCopiedFrom.Location = New System.Drawing.Point(24, 64)
        Me.lbxAnalytesCopiedFrom.Name = "lbxAnalytesCopiedFrom"
        Me.lbxAnalytesCopiedFrom.Size = New System.Drawing.Size(204, 171)
        Me.lbxAnalytesCopiedFrom.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(240, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(204, 47)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Choose an analyte to which data will be copied:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'lbxAnalytesCopiedTo
        '
        Me.lbxAnalytesCopiedTo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbxAnalytesCopiedTo.Location = New System.Drawing.Point(240, 66)
        Me.lbxAnalytesCopiedTo.Name = "lbxAnalytesCopiedTo"
        Me.lbxAnalytesCopiedTo.Size = New System.Drawing.Size(204, 171)
        Me.lbxAnalytesCopiedTo.TabIndex = 8
        '
        'frmCopyReplicate
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(480, 308)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lbxAnalytesCopiedTo)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.lblAddRep)
        Me.Controls.Add(Me.lbxAnalytesCopiedFrom)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmCopyReplicate"
        Me.ShowInTaskbar = False
        Me.Text = "Copy Replicate..."
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        boolStop = False
        Me.Visible = False

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        boolStop = True
        Me.Visible = False
    End Sub
End Class
