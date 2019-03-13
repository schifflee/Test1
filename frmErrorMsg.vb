Option Compare Text

Public Class frmErrorMsg
    Public pbMax As Short = 10
    Public pbVal As Short = 1

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        Me.Visible = False

    End Sub

    Private Sub frmErrorMsg_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim str1 As String

        Me.pb1.ForeColor = Color.Gold
        Me.pb1.Refresh()

        Cursor.Current = Cursors.WaitCursor

        If boolFormLoad Then 'locate on top of splash

            Me.SuspendLayout()

            Dim t, l, t1, l1, t2, l2
            Dim h, w, h1, w1
            t = St
            l = Sl
            w = Sw
            w1 = Me.Width

            Me.Top = t - Me.Height
            Me.Left = l - ((w1 - w) / 2)

            Me.ResumeLayout()

        End If

    End Sub

    Sub RunTimer()

        Exit Sub

        Me.pb1.Maximum = pbMax
        Me.pb1.Value = pbVal
        Me.pb1.Refresh()
        'Me.TimerE.Enabled = True
        Me.Refresh()
    End Sub

    Private Sub TimerE_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerE.Tick
        Exit Sub

        Me.pbVal = Me.pbVal + 1
        If Me.pbVal > Me.pbMax Then
            Me.pbVal = 1
        End If
        Me.pb1.Value = Me.pbVal
        Me.pb1.Refresh()
    End Sub
End Class