Option Compare Text

Public Class frmAddQARow
    Inherits System.Windows.Forms.Form
    Public boolCancel As Boolean
    Public char1 As String
    Public numID As Int64

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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lbxQACriticalPhase As System.Windows.Forms.ListBox
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents lbxID As System.Windows.Forms.ListBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAddQARow))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lbxQACriticalPhase = New System.Windows.Forms.ListBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.lbxID = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(24, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(224, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Choose a Critical Phase:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'lbxQACriticalPhase
        '
        Me.lbxQACriticalPhase.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbxQACriticalPhase.ItemHeight = 17
        Me.lbxQACriticalPhase.Location = New System.Drawing.Point(24, 32)
        Me.lbxQACriticalPhase.Name = "lbxQACriticalPhase"
        Me.lbxQACriticalPhase.Size = New System.Drawing.Size(224, 155)
        Me.lbxQACriticalPhase.TabIndex = 1
        '
        'cmdCancel
        '
        Me.cmdCancel.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(192, 208)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(56, 32)
        Me.cmdCancel.TabIndex = 2
        Me.cmdCancel.Text = "&Cancel"
        '
        'cmdOK
        '
        Me.cmdOK.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.Blue
        Me.cmdOK.Location = New System.Drawing.Point(24, 208)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(56, 32)
        Me.cmdOK.TabIndex = 3
        Me.cmdOK.Text = "&OK"
        '
        'lbxID
        '
        Me.lbxID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbxID.ItemHeight = 17
        Me.lbxID.Location = New System.Drawing.Point(96, 208)
        Me.lbxID.Name = "lbxID"
        Me.lbxID.Size = New System.Drawing.Size(64, 36)
        Me.lbxID.TabIndex = 4
        Me.lbxID.Visible = False
        '
        'frmAddQARow
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 18)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(272, 261)
        Me.Controls.Add(Me.lbxID)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.lbxQACriticalPhase)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAddQARow"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add New QA Table Row..."
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        boolCancel = True
        Me.Visible = False

    End Sub

    Private Sub frmAddQARow_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        boolCancel = False

        If Me.lbxQACriticalPhase.Items.Count > 1 Then

            'select first item
            Me.lbxQACriticalPhase.SelectedIndex = 0

        End If

    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        Call DoOK()

    End Sub

    Sub DoOK()
        Dim var1
        Dim int1 As Short

        boolCancel = False
        var1 = Me.lbxQACriticalPhase.SelectedItem
        If var1 = Nothing Then
            MsgBox("An item must be selected.", MsgBoxStyle.Information, "An item must be selected...")
        Else
            char1 = var1
            int1 = Me.lbxQACriticalPhase.SelectedIndex
            numID = Me.lbxID.Items(int1)
            Me.Visible = False
        End If

    End Sub

    Private Sub lbxQACriticalPhase_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbxQACriticalPhase.DoubleClick

        Call DoOK()

    End Sub


End Class
