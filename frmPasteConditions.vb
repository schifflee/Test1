Public Class frmPasteConditions

    Public boolCancel = True

    Private Sub frmPasteConditions_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'center pancmd
        Dim a, b, c, d

        a = Me.Width
        b = Me.panCmd.Width

        Me.panCmd.Left = (a / 2) - (b / 2)

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