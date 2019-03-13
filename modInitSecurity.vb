Option Compare Text

Module modInitSecurity


    Sub main()
        ' InitSecurity
        '
        ' This routine calls CoInitializeSecurity() and MUST be called at the
        '  very beginning of the application (from a "Sub Main" routine). This routine 
        ' needs() ' to be called in order to open Analyst data files 
        '(using the Explore Data ' Objects) which reside on network drives.
        '
        ' The return value is an error code indicating success (if zero) or failure ' (if non-zero). However note that this routine will always fail when running 
        ' from within the VB IDE - it only works from the stand-alone application.
        Dim myApp As Security
        myApp = New Security

        'Application.Run(New frmConsole())

        System.Windows.Forms.Application.Run(New frmConsole)


        ' Create the MyApplicationContext, that derives from ApplicationContext,
        ' that manages when the application should exit.

        'Dim context As MyApplicationContext = New MyApplicationContext()

        ' Run the application with the specific context. It will exit when
        ' all forms are closed.
        'Application.Run(context)



    End Sub

    Public Class MyApplicationContext
        Inherits ApplicationContext

        Private formCount As Integer
        Private form1 As frmConsole 'AppForm1
        'Private form2 As frmStatus ' AppForm2

        Private form1Position As System.Drawing.Rectangle
        Private form2Position As System.Drawing.Rectangle

        'Private userData As FileStream

        Public Sub New()
            MyBase.New()
            formCount = 0

            ' Handle the ApplicationExit event to know when the application is exiting.
            AddHandler System.Windows.Forms.Application.ApplicationExit, AddressOf OnApplicationExit


            ' Create both application forms and handle the Closed event
            ' to know when both forms are closed.
            'form1 = New AppForm1()
            form1 = New frmConsole

            'AddHandler form1.Closed, AddressOf OnFormClosed
            'AddHandler form1.Closing, AddressOf OnFormClosing
            formCount = formCount + 1

            'form2 = New AppForm2()
            'form2 = New frmStatus

            'AddHandler form2.Closed, AddressOf OnFormClosed
            'AddHandler form2.Closing, AddressOf OnFormClosing
            formCount = formCount + 1

            ' Get the form positions based upon the user specific data.
            'If (ReadFormDataFromFile()) Then
            '    ' If the data was read from the file, set the form
            '    ' positions manually.
            '    form1.StartPosition = FormStartPosition.Manual
            '    form2.StartPosition = FormStartPosition.Manual

            '    form1.Bounds = form1Position
            '    form2.Bounds = form2Position
            'End If

            ' Show both forms.
            form1.Show()
            'form2.Show()
        End Sub

        Private Sub OnApplicationExit(ByVal sender As Object, ByVal e As EventArgs)

        End Sub


        'Private Sub OnFormClosing(ByVal sender As Object, ByVal e As CancelEventArgs)


        '    ' When a form is closing, remember the form position so it
        '    '' can be saved in the user data file.
        '    'If TypeOf sender Is AppForm1 Then
        '    '    form1Position = CType(sender, Form).Bounds
        '    'ElseIf TypeOf sender Is AppForm2 Then
        '    '    form2Position = CType(sender, Form).Bounds
        '    'End If

        '    If TypeOf sender Is frmConsole Then
        '        form1Position = CType(sender, Form).Bounds
        '    ElseIf TypeOf sender Is frmStatus Then
        '        form2Position = CType(sender, Form).Bounds
        '    End If

        'End Sub

        Private Sub OnFormClosed(ByVal sender As Object, ByVal e As EventArgs)
            ' When a form is closed, decrement the count of open forms.

            ' When the count gets to 0, exit the app by calling
            ' ExitThread().
            formCount = formCount - 1
            If (formCount = 0) Then
                ExitThread()
            End If
        End Sub

    End Class
End Module
