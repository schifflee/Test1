Option Compare Text


Public Class frmSplash1

    Public pbMax As Short = 10
    Public pbVal As Short = 1

    'Private Sub frmSplash1_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
    '    Me.Refresh()

    'End Sub


    Private Sub frmSplash1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Dim boolDemo As Boolean
        Dim dt1 As Date
        Dim dt2 As Date

        '20160617 LEE:
        'do not run ControlDefaults on Splash
        'Call ControlDefaults(Me)


        boolDemo = False

        dt2 = #9/1/2019#

        Cursor.Current = Cursors.WaitCursor
        Dim str1, str2 As String
        Dim w, h, w1, h1
        Dim strV As String = GetVersion()

        str1 = "LABIntegrity StudyDoc" & ChrW(8482) & " v" & strV & " Start Up..."
        Me.Text = str1

        'Dim intP1 As Int32
        'Dim intP2 As Int32

        'intP1 = 27
        'intP2 = 157

        'Dim lbl1 As New System.Windows.Forms.Label
        'lbl1.Margin = New System.Windows.Forms.Padding(0, 0, 0, 0, FALSE)
        'lbl1.Location = New Point(intP1, intP2)
        'lbl1.Text = "Study"
        'lbl1.AutoSize = True
        'lbl1.Font = New Font("Century Gothic", 48, FontStyle.Bold)
        'lbl1.BringToFront()
        'lbl1.BackColor = Color.Transparent
        'lbl1.ForeColor = Color.FromArgb(25, 35, 79)

        'Dim lbl2 As New System.Windows.Forms.Label
        'lbl2.Margin = New System.Windows.Forms.Padding(0, 0, 0, 0, FALSE)
        'lbl2.Location = New Point(intP1, intP2)
        'lbl2.Left = lbl1.Left + lbl1.Width + 2
        'lbl2.Text = "Doc" ' & ChrW(8482)
        'lbl2.AutoSize = True
        'lbl2.Font = New Font("Century Gothic", 48, FontStyle.Bold)
        'lbl2.BringToFront()
        'lbl2.BackColor = Color.Transparent
        'lbl2.ForeColor = Color.FromArgb(49, 112, 193) '49, 112, 193

        'Dim lbl3 As New System.Windows.Forms.Label
        'lbl3.Margin = New System.Windows.Forms.Padding(0, 0, 0, 0, FALSE)
        'lbl3.Location = New Point(intP1, intP2 - 5)
        'lbl3.Text = ChrW(8482)
        'lbl3.AutoSize = True
        'lbl3.Font = New Font("Century Gothic", 36, FontStyle.Bold)
        'lbl3.BringToFront()
        'lbl3.BackColor = Color.Transparent
        'lbl3.ForeColor = Color.FromArgb(49, 112, 193) '49, 112, 193

        ''Me.pan1.Controls.Add(lbl1)
        ''Me.pan1.Controls.Add(lbl2)
        ''Me.pan1.Controls.Add(lbl3)

        'lbl1.AutoSize = True
        'lbl2.AutoSize = True
        'lbl3.AutoSize = True

        'lbl2.Left = lbl1.Left + lbl1.Width + 2
        'lbl3.Left = lbl2.Left + lbl2.Width + 2

        ''Me.Label2.Visible = False
        ''Me.Label3.Visible = False
        ''Me.lblR.Visible = False

        'Me.Label2.Margin = New System.Windows.Forms.Padding(0, 0, 0, 0, FALSE)
        'Me.Label3.Margin = New System.Windows.Forms.Padding(0, 0, 0, 0, FALSE)
        'Me.lblR.Margin = New System.Windows.Forms.Padding(0, 0, 0, 0, FALSE)

        'Me.Label3.Left = Me.Label2.Left + Me.Label2.Width
        'Me.lblR.Left = Me.Label3.Left + Me.Label3.Width


        'this.label1.Margin = new System.Windows.Forms.Padding(3,3,2,2);<br/>   this.label1.Padding = new System.Windows.Forms.Padding(4, 2, 3, 

        'Note: For using a network service account, look in help for networkcredential

        'w = My.Computer.Screen.WorkingArea.Width
        'h = My.Computer.Screen.WorkingArea.Height

        'Me.Width = w '* 0.9
        'Me.Height = h ' * 0.9
        'Me.Left = (w - Me.Width) / 2
        'Me.Top = (h - Me.Height) / 2

        Call SetFormPos(Me)

        'Me.WindowState = FormWindowState.Maximized


        Me.SuspendLayout()

        'Me.lbl2.TextAlign = ContentAlignment.MiddleRight
        'str1 = "Gubbs Watson" & ChrW(8482) & Chr(10) & "Report Writing Manager"
        str1 = GetStudyDocHeader(True) & " v" & strV
        Me.lbl2.Text = str1

        'center pan1
        w = Me.Width 'My.Computer.Screen.WorkingArea.Width
        h = Me.Height 'My.Computer.Screen.WorkingArea.Height
        w1 = Me.pan1.Width
        h1 = Me.pan1.Height
        Dim xCenter, yCenter As Single
        xCenter = ((w - w1) / 2)
        yCenter = ((h - h1) / 2)
        Me.pan1.Location = New System.Drawing.Point(xCenter, yCenter)

        'center pb1
        w1 = Me.pb2.Width
        h1 = Me.pb2.Height
        xCenter = ((w - w1) / 2)
        ' NIck Change
        ' yCenter = Me.pan1.Top - Me.pb2.Height - 10 '((h - h1) / 2)
        yCenter = Me.pan1.Bottom + (Me.pb2.Height / 2) + 5 '5 pixels below splashscreen
        Me.pb2.Location = New System.Drawing.Point(xCenter, yCenter)

        'center lblerr
        w1 = Me.lblErr.Width
        h1 = Me.lblErr.Height
        ' Nick Change: 
        'xCenter = ((w - w1) / 2)
        'yCenter = Me.pb2.Top - Me.lblErr.Height - 10 '((h - h1) / 2)
        xCenter = 20
        yCenter = h - h1 - 150  '50 gets it high enough to be seen.
        Me.lblErr.Location = New System.Drawing.Point(xCenter, yCenter)

        'dumb
        Me.lblErr.Left = Me.pan1.Left
        Me.lblErr.Top = Me.pan1.Top + Me.pan1.Height + 7

        'Me.lblErr.Text = "...Establishing communication with the Watson" & ChrW(8482) & " database..."
        'Me.lblErr.Visible = True

        Me.pan1.Visible = True
        'Me.pb1.Visible = True
        'Me.lblErr.Visible = True

        Me.ResumeLayout()

        'lblR.Text = ChrW(8482)



        If boolDemo Then
            'evaluate dt1 and dt2
            dt1 = Now
            If dt1 > dt2 Then
                str1 = "StudyDoc evaluation period expired on " & Format(dt2, "MMMM dd, yyyy") & "."
                MsgBox(str1, MsgBoxStyle.Information, "Evaluation period expired...")
                End
            Else
                str1 = "StudyDoc evaluation period expires on " & Format(dt2, "MMMM dd, yyyy") & "."
                MsgBox(str1, MsgBoxStyle.Information, "Evaluation period expires on...")
            End If
        End If

    End Sub

    Private Sub txtMax_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMax.TextChanged
        Dim var1

        var1 = CLng(Me.txtMax.Text)
        If var1 = 10000 Then
            Me.pb2.Visible = False
            Me.Refresh()
        Else
            Me.pb2.Maximum = var1
            Me.pb2.Visible = True
            Me.pb2.Refresh()
        End If
    End Sub

    Private Sub txtValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtValue.TextChanged
        Dim var1

        var1 = CLng(Me.txtValue.Text)

        If var1 = 0 Then
            Me.pb2.Visible = False
            Me.Refresh()

        Else
            If var1 > Me.pb2.Maximum Then
                Me.pb2.Maximum = CInt(Me.pb2.Maximum * 1.25)
            End If
            Me.pb2.Value = var1
            Me.pb2.Refresh()
        End If
    End Sub

    Private Sub pan1_Paint(sender As Object, e As PaintEventArgs) Handles pan1.Paint

    End Sub



    Private Sub Label4_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub lbl2_Click(sender As System.Object, e As System.EventArgs) Handles lbl2.Click

    End Sub
End Class