Option Compare Text

Public Class frmAbort

    Private Sub cmdReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReport.Click

        Dim int1 As Short
        Dim str1 As String

        str1 = "Do you really wish to abort this report?"
        int1 = MsgBox(str1, MsgBoxStyle.OkCancel, "Abort report?")
        If int1 = 1 Then 'abort
            Try
                wdAbort.Quit(False)
            Catch ex As Exception

            End Try
        End If


    End Sub
End Class