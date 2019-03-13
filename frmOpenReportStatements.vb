Option Compare Text

Public Class frmOpenReportStatements
    Public boolCancel As Boolean = True

    Private Sub frmOpenReportStatements_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.rbGuWu.Checked = True
        Me.rbReadOnly.Checked = True
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        boolCancel = False
        Me.Visible = False

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        boolCancel = True
        Me.Visible = False

    End Sub
End Class