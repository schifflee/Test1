Namespace My

    ' The following events are availble for MyApplication:
    ' 
    ' Startup: Raised when the application starts, before the startup form is created.
    ' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally.
    ' UnhandledException: Raised if the application encounters an unhandled exception.
    ' StartupNextInstance: Raised when launching a single-instance application and the application is already active. 
    ' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected.
    Partial Friend Class MyApplication

        Private Sub MyApplication_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown

            Try
                'frmH.wbRBS.Navigate("about:blank")
            Catch ex As Exception

            End Try

        End Sub

        Private Sub MyApplication_Startup(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.ApplicationServices.StartupEventArgs) Handles Me.Startup
            Dim var1

            var1 = 1 'for debugging purposes


        End Sub

        Private Sub MyApplication_UnhandledException( _
            ByVal sender As Object, _
            ByVal e As Microsoft.VisualBasic.ApplicationServices.UnhandledExceptionEventArgs _
        ) Handles Me.UnhandledException
            My.Application.Log.WriteException(e.Exception, _
                TraceEventType.Critical, _
                "Unhandled Exception.")
            Dim strM As String

            strM = "An unhandled exception was detected. This exception may or may not cause StudyDoc to shutdown."
            strM = strM & Chr(10) & "It is recommended that this notice be reported to your StudyDoc system administrator."
            strM = strM & Chr(10) & Chr(10)
            strM = strM & "Exception descriptions:"
            strM = strM & e.Exception.ToString
            MsgBox(strM, MsgBoxStyle.Information, "Unhandled Application Exception...")

        End Sub



    End Class

End Namespace

