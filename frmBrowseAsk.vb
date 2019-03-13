Option Compare Text

Public Class frmBrowseAsk
    Public boolAddPath As Boolean = True
    Public boolGo As Boolean = False

    Private Sub frmBrowseAsk_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call ControlDefaults(Me)

        boolAddPath = True


    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        boolGo = False
        If Me.rbAddPath.Checked Then
            boolAddPath = True
            boolGo = True
        ElseIf Me.rbRemovePath.Checked Then
            boolAddPath = False
            boolGo = True
        End If

        Me.Visible = False

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        boolGo = False
        Me.Visible = False

    End Sub
End Class