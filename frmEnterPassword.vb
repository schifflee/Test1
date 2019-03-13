Option Compare Text

Public Class frmEnterPassword
    Public boolCancel As Boolean = True
    Private Sub frmEnterPassword_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        boolCancel = True
        Dim var1, var2
        Dim frm As New frmAdministration
        var1 = frm.cmdEnterPassword.Left + frm.pan1.Left + frm.Left
        var2 = frm.cmdEnterPassword.Top + frm.pan1.Top + frm.Top

        Me.Top = var2 + frm.cmdEnterPassword.Height
        Me.Left = var1

        frm.Dispose()

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        boolCancel = True
        Me.Visible = False
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        boolCancel = False
        'ensure password is not blank
        Dim var1
        var1 = Me.txtPassword.Text
        If Len(NZ(var1, "")) = 0 Then
            MsgBox("Password cannot be blank.", MsgBoxStyle.Information, "Password cannot be blank...")
            Exit Sub
        End If
        Me.Visible = False
    End Sub
End Class