Option Compare Text

Public Class frmReportErrMsg
    Inherits System.Windows.Forms.Form

    Private Sub frmReportErrMsg_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call ControlDefaults(Me)

        Dim str1 As String

        str1 = "The following field code anomolies were detected during the preparation of this report. "
        str1 = str1 & "If so desired, click on the Copy to Clipboard button to copy/paste this information to "
        str1 = str1 & "a document (e.g. Word, Excel)  for further reference."
        Me.lblTitle.Text = str1

        str1 = "Copy to Clipboard generates a semicolon-delimited text. For best viewing results, paste into "
        str1 = str1 & "Microsoft(R) Excel and select semicolon-delimited when prompted by the Excel text input wizard."
        Me.lbl2.Text = str1

        Me.lblTitle.Visible = True
        Me.lbl2.Visible = True
        Me.cmdCopyToClipboard.Visible = True

        If ctArrReportNA = 0 Then

            str1 = "There were no field code anomolies found in this report." & ChrW(10) & ChrW(10)
            str1 = str1 & "This report has all field codes satisfied."
            Me.lblTitle.Text = str1

            Me.lbl2.Visible = False

            Me.cmdCopyToClipboard.Visible = False


        End If

        Me.cmdExit.Focus()


    End Sub

    Private Sub cmdCopyToClipboard_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopyToClipboard.Click
        'first select entire table
        Dim str1 As String
        Dim var1
        Dim int1 As Short
        Dim int2 As Short
        Dim Count1 As Short
        Dim Count2 As Short

        Me.dgvReportErrMsg.SelectAll()
        If Me.dgvReportErrMsg.GetCellCount(DataGridViewElementStates.Selected) > 0 Then

            Try
                ' Add the selection to the clipboard.
                var1 = Me.dgvReportErrMsg.GetClipboardContent()
                str1 = Me.txtReportTitle.Text
                str1 = Me.txtReportTitle.Text & Chr(10)
                int1 = Me.dgvReportErrMsg.Columns.Count
                'enter column headings
                For Count1 = 0 To int1 - 2
                    var1 = Me.dgvReportErrMsg.Columns(Count1).Name
                    str1 = str1 & var1 & ChrW(9)
                Next
                var1 = Me.dgvReportErrMsg.Columns(int1 - 1).Name
                str1 = str1 & var1 & Chr(10)
                'enter grid data
                int2 = Me.dgvReportErrMsg.Rows.Count
                For Count1 = 0 To int2 - 1
                    For Count2 = 0 To int1 - 2
                        var1 = Me.dgvReportErrMsg.Item(Count2, Count1).Value
                        str1 = str1 & NZ(var1, "") & ChrW(9)
                    Next
                    var1 = Me.dgvReportErrMsg.Item(int1 - 1, Count1).Value
                    str1 = str1 & NZ(var1, "") & Chr(10)
                Next

                Clipboard.SetDataObject(str1, True)
                'Clipboard.SetDataObject(Me.dgvReportErrMsg.GetClipboardContent())

            Catch ex As System.Runtime.InteropServices.ExternalException
                MsgBox("The Clipboard could not be accessed. Please try again.")
            End Try

        End If

    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Visible = False

    End Sub

    Private Sub dgvReportErrMsg_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvReportErrMsg.MouseEnter
        Me.dgvReportErrMsg.Focus()
    End Sub
End Class