Public Class frmAutoAssignBegin

    Public boolCancel As Boolean = True


    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click

        boolCancel = True
        Me.Visible = False

    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click

        boolCancel = False
        Me.Visible = False

    End Sub
End Class