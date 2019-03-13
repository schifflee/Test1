Public Class SampleIForm1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        My.Forms.LoginForm1.ShowDialog()
        ' Check if the user was authenticated.
        If My.User.IsAuthenticated Then
            Me.Label1.Text = "Authenticated " & My.User.Name
        Else
            Me.Label1.Text = "User not authenticated"
        End If

        If My.User.IsInRole(ApplicationServices.BuiltInRole.Administrator) Then
            Me.Label2.Text = "User is an Administrator"
        Else
            Me.Label2.Text = "User is not an Administrator"
        End If


    End Sub
End Class