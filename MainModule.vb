Module MainModule
    <STAThread()> Sub Main()
        'install the event handler
        AddHandler System.Windows.Forms.Application.ThreadException, AddressOf Application_ThreadException

        System.Windows.Forms.Application.Run(New frmHome_01)

    End Sub

    Sub Application_ThreadException(ByVal sender As Object, ByVal e As System.Threading.ThreadExceptionEventArgs)
        Try
            Dim msg As String
            msg = "An error occurred:" & Chr(10) & Chr(10)
            msg = msg & e.Exception.Message & Chr(10) & Chr(10)
            msg = msg & e.Exception.StackTrace & Chr(10) & Chr(10)
            msg = msg & "The application startup will continue, but this error msg should be reported to you GuWu administrator."

            Dim Result As DialogResult = MessageBox.Show(msg, "Application Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'If Result = DialogResult.Abort Then
            '    System.Windows.Forms.Application.Exit()
            'End If

        Catch ex As Exception
            Try
                MsgBox("The application will be terminated.", MsgBoxStyle.Critical, "Fatal error...")

            Finally
                System.Windows.Forms.Application.Exit()

            End Try

        End Try
    End Sub
End Module
