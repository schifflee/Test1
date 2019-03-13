Option Compare Text

Public Class frmPermGroupDelete

    Private Sub cmdOK_Click(sender As System.Object, e As System.EventArgs) Handles cmdOK.Click

        Me.Dispose()

    End Sub

    Private Sub cmdClipboard_Click(sender As System.Object, e As System.EventArgs) Handles cmdClipboard.Click

        Dim dgv As DataGridView = Me.dgvPerm
        Dim strM As String

        Try
            dgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText

            'select all rows
            Dim Count1 As Int32
            For Count1 = 0 To dgv.Rows.Count - 1
                dgv.Rows(Count1).Selected = True
            Next

            Clipboard.SetDataObject(dgv.GetClipboardContent())
        Catch ex As Exception
            strM = "There was a problem pasting the table to the clipboard."
            strM = strM & ChrW(10) & ChrW(10) & ex.Message
            MsgBox(strM, vbInformation, "Problem...")
        End Try



    End Sub
End Class