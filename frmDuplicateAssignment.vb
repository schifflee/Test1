Option Compare Text

Public Class frmDuplicateAssignment
    Public boolCancel As Boolean = True

    Private Sub frmDuplicateAssignment_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        boolCancel = True

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        boolCancel = True
        Me.Visible = False
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        Dim sr As DataGridViewSelectedRowCollection
        Dim dgv As DataGridView

        'ensure one or more rows are selected
        sr = Me.dgvAnalytes.SelectedRows

        If sr.Count = 0 Then
            MsgBox("Please choose at least one analyte.", MsgBoxStyle.Information, "Choose at least one analyte...")
            GoTo end1
        End If

        boolCancel = False
        Me.Visible = False

end1:

    End Sub
End Class