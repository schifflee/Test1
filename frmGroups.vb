Option Compare Text

Public Class frmGroups

    Private Sub frmGroups_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        Dim str1 As String
        str1 = "Show _C[n] Groups"
        frmH.cmdShowGroups.Text = str1

    End Sub

    Private Sub frmGroups_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call dgvSetGroups(Me.dgvFormGroups)

    End Sub

    Private Sub dgvFormGroups_Resize(sender As Object, e As EventArgs) Handles dgvFormGroups.Resize

        Dim dgv As DataGridView = Me.dgvFormGroups
        dgv.AutoResizeColumns()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim dgv As DataGridView = Me.dgvFormGroups
        Dim Count1 As Short

        For Count1 = 1 To dgv.Columns.Count
            dgv.Columns(Count1 - 1).Visible = True
        Next

    End Sub
End Class