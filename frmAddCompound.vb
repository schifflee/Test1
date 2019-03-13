Option Compare Text

Public Class frmAddCompound
    Public boolCancel As Boolean = True
    Private Sub frmAddCompound_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        boolCancel = True

    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        boolCancel = False
        Me.Visible = False
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        boolCancel = True
        Me.Visible = False
    End Sub

    Private Sub txtName_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtName.Validating
        Dim str1 As String
        Dim str2 As String

        str1 = Me.txtName.Text
        If Len(Trim(str1)) = 0 Then
            str2 = "Analyte Name cannot be blank."
            MsgBox(str2, MsgBoxStyle.Information, "Entry cannot be blank..")
            Me.txtName.Select()
        End If

    End Sub
End Class