Public Class frmPasswordInput

    Public boolCancel As Boolean = False

    Private Sub frmPasswordInput_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call ControlDefaults(Me)

    End Sub

    Private Sub cmdEdit_Click(sender As Object, e As EventArgs) Handles cmdEdit.Click

        boolCancel = False

        Me.Visible = False

    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click

        boolCancel = True

        Me.Visible = False

    End Sub

    Private Sub chkShow_CheckedChanged(sender As Object, e As EventArgs) Handles chkShow.CheckedChanged

        Call PWD()

    End Sub

    Sub PWD()

        If Me.chkShow.Checked Then
            Me.txtPassword.PasswordChar = ""
        Else
            Me.txtPassword.PasswordChar = "*"
        End If

    End Sub
End Class