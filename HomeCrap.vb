Public Class Home
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents lbxTab1 As System.Windows.Forms.ListBox
    Friend WithEvents cmdUpdateProject As System.Windows.Forms.Button
    Friend WithEvents lblcbxStudies As System.Windows.Forms.Label
    Friend WithEvents cbxStudies As System.Windows.Forms.ComboBox
    Friend WithEvents cmdPrepareReport As System.Windows.Forms.Button
    Friend WithEvents charDateofLastWatsonStudyModification As System.Windows.Forms.TextBox
    Friend WithEvents charDateofLastReportGeneration As System.Windows.Forms.TextBox
    Friend WithEvents charDateofLastSave As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Home))
        Me.cmdUpdateProject = New System.Windows.Forms.Button
        Me.lblcbxStudies = New System.Windows.Forms.Label
        Me.cbxStudies = New System.Windows.Forms.ComboBox
        Me.cmdPrepareReport = New System.Windows.Forms.Button
        Me.lbxTab1 = New System.Windows.Forms.ListBox
        Me.charDateofLastWatsonStudyModification = New System.Windows.Forms.TextBox
        Me.charDateofLastReportGeneration = New System.Windows.Forms.TextBox
        Me.charDateofLastSave = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'cmdUpdateProject
        '
        Me.cmdUpdateProject.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUpdateProject.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(64, Byte), CType(0, Byte))
        Me.cmdUpdateProject.Location = New System.Drawing.Point(5, 285)
        Me.cmdUpdateProject.Name = "cmdUpdateProject"
        Me.cmdUpdateProject.Size = New System.Drawing.Size(96, 32)
        Me.cmdUpdateProject.TabIndex = 8
        Me.cmdUpdateProject.Text = "Update Project"
        '
        'lblcbxStudies
        '
        Me.lblcbxStudies.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblcbxStudies.Location = New System.Drawing.Point(35, 390)
        Me.lblcbxStudies.Name = "lblcbxStudies"
        Me.lblcbxStudies.Size = New System.Drawing.Size(160, 16)
        Me.lblcbxStudies.TabIndex = 6
        Me.lblcbxStudies.Text = "Select a Watson Study:"
        '
        'cbxStudies
        '
        Me.cbxStudies.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxStudies.Location = New System.Drawing.Point(-5, 440)
        Me.cbxStudies.Name = "cbxStudies"
        Me.cbxStudies.Size = New System.Drawing.Size(160, 21)
        Me.cbxStudies.Sorted = True
        Me.cbxStudies.TabIndex = 4
        '
        'cmdPrepareReport
        '
        Me.cmdPrepareReport.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPrepareReport.ForeColor = System.Drawing.Color.Blue
        Me.cmdPrepareReport.Location = New System.Drawing.Point(10, 345)
        Me.cmdPrepareReport.Name = "cmdPrepareReport"
        Me.cmdPrepareReport.Size = New System.Drawing.Size(96, 32)
        Me.cmdPrepareReport.TabIndex = 5
        Me.cmdPrepareReport.Text = "Generate Report"
        '
        'lbxTab1
        '
        Me.lbxTab1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbxTab1.Location = New System.Drawing.Point(5, 10)
        Me.lbxTab1.Name = "lbxTab1"
        Me.lbxTab1.Size = New System.Drawing.Size(136, 236)
        Me.lbxTab1.TabIndex = 5
        '
        'charDateofLastWatsonStudyModification
        '
        Me.charDateofLastWatsonStudyModification.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.charDateofLastWatsonStudyModification.Location = New System.Drawing.Point(248, 70)
        Me.charDateofLastWatsonStudyModification.Name = "charDateofLastWatsonStudyModification"
        Me.charDateofLastWatsonStudyModification.Size = New System.Drawing.Size(88, 20)
        Me.charDateofLastWatsonStudyModification.TabIndex = 5
        Me.charDateofLastWatsonStudyModification.Text = "charDateofLastWatsonStudyModification"
        '
        'charDateofLastReportGeneration
        '
        Me.charDateofLastReportGeneration.BackColor = System.Drawing.Color.Red
        Me.charDateofLastReportGeneration.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.charDateofLastReportGeneration.Location = New System.Drawing.Point(230, 45)
        Me.charDateofLastReportGeneration.Name = "charDateofLastReportGeneration"
        Me.charDateofLastReportGeneration.Size = New System.Drawing.Size(88, 20)
        Me.charDateofLastReportGeneration.TabIndex = 4
        Me.charDateofLastReportGeneration.Text = "charDateofLastReportGeneration"
        '
        'charDateofLastSave
        '
        Me.charDateofLastSave.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.charDateofLastSave.Location = New System.Drawing.Point(248, 20)
        Me.charDateofLastSave.Name = "charDateofLastSave"
        Me.charDateofLastSave.Size = New System.Drawing.Size(88, 20)
        Me.charDateofLastSave.TabIndex = 3
        Me.charDateofLastSave.Text = "charDateofLastSave"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(208, 16)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Date"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(176, 16)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Date"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(176, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Date"
        '
        'Home
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1028, 533)
        Me.Controls.Add(Me.lbxTab1)
        Me.Controls.Add(Me.cmdUpdateProject)
        Me.Controls.Add(Me.cmdPrepareReport)
        Me.Controls.Add(Me.cbxStudies)
        Me.Controls.Add(Me.lblcbxStudies)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Home"
        Me.Text = "GooWoo Home"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

#End Region



    Private Sub Home_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim var1, var2
        Dim Count1 As Integer
        Dim str1 As String
        Dim ct1 As Integer
        Exit Sub

        'fill lbxTab1
        ct1 = 6
        For Count1 = 1 To 6
            Select Case Count1
                Case 1
                    str1 = "Home"
                Case 2
                    str1 = "Data"
                Case 3
                    str1 = "Analytical Run Summary"
                Case 4
                    str1 = "Summary Table"
                Case 5
                    str1 = "Report Table Configuration"
                Case 6
                    str1 = "Blank"
            End Select
            Me.lbxTab1.Items.Add(str1)
            If Count1 > Me.tab1.TabPages.Count Then
            Else
                Me.tab1.TabPages(Count1 - 1).Text = ""
            End If

            Me.WindowState = FormWindowState.Maximized

        Next
    End Sub


    Private Sub lbxTab1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbxTab1.SelectedIndexChanged
        Dim int1 As Integer
        Dim int2 As Integer
        Dim Count1 As Integer

        'record selected row number
        int1 = Me.lbxTab1.SelectedIndex
        int2 = Me.tab1.TabPages.Count
        'select appropriate tab
        If int1 > Me.tab1.TabPages.Count - 1 Then
        Else
            Me.tab1.SelectedTab = Me.tab1.TabPages.Item(int1)
            'hide all tabs
        End If




    End Sub


End Class
