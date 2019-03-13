Option Compare Text

Public Class frmAskOutlierReport
    Public boolCancel As Boolean = True

    Private Sub frmAskOutlierReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call ControlDefaults(Me)

    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        boolCancel = False
        Me.Close()

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub rbDetailed_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbDetailed.CheckedChanged
        If Me.rbDetailed.Checked Then
            Me.gb2.Enabled = True
        Else
            Me.gb2.Enabled = False
        End If
    End Sub
End Class