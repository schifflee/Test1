Public Class frmSaveFinalDraft

    Public boolCancel As Boolean = True


    Private Sub frmSaveFinalDraft_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click

        boolCancel = False
        Me.Visible = False

    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click

        boolCancel = True
        Me.Visible = False

    End Sub
End Class