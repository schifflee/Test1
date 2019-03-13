Public Class frmFinalReportPrintOptions

    Public boolCancel As Boolean = True

    Private Sub rbNone_CheckedChanged(sender As Object, e As EventArgs) Handles rbNone.CheckedChanged

        Call DoEnable()

    End Sub

    Private Sub frmFinalReportPrintOptions_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call FormSize()

        Call FinalReportPrintOptionsToolTips()

    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click

        boolCancel = False

        Me.Visible = False

    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click

        boolCancel = True

        Me.Visible = False

    End Sub

    Sub DoEnable()


        If Me.rbNone.Checked Then

            Me.gbLabel.Enabled = False
            Me.gbLocation.Enabled = False

        Else

            If Me.rbWaterMark.Checked Then
                Me.gbLocation.Enabled = True
            Else
                Me.gbLocation.Enabled = False
            End If

            Me.gbLabel.Enabled = True

        End If

    End Sub

    Private Sub rbWaterMark_CheckedChanged(sender As Object, e As EventArgs) Handles rbWaterMark.CheckedChanged

        Call DoEnable()

    End Sub

    Private Sub rbText_CheckedChanged(sender As Object, e As EventArgs) Handles rbText.CheckedChanged

        Call DoEnable()

    End Sub

    Sub FormSize()

        If Me.gbChoice.Visible Then
        Else

            Me.panRest.Top = Me.gbChoice.Top

            Dim bw As Int16 = (Me.Width - Me.ClientSize.Width) / 2 'form border width
            Dim tbh As Int16 = Me.Height - Me.ClientSize.Height - 2 * bw 'titlebar height

            Me.Height = Me.panRest.Top + Me.panRest.Height + bw + tbh + 15

        End If

    End Sub

    Sub FinalReportPrintOptionsToolTips()

        ' Create the ToolTip and associate with the Form container.
        Dim toolTip1 As New ToolTip()
        Dim str1 As String

        ' Set up the delays for the ToolTip.
        'toolTip1.AutoPopDelay = 5000
        'toolTip1.InitialDelay = 250
        'toolTip1.ReshowDelay = 50

        toolTip1.AutomaticDelay = intToolTipDelay
        'toolTip1.UseFading = False
        'tooltip1.
        'toolTip1.BackColor = Color.Goldenrod
        'toolTip1.IsBalloon = True
        ' Force the ToolTip text to be displayed whether or not the form is active.
        toolTip1.ShowAlways = True

        Try
            'Set mode buttons
            str1 = "StudyDoc ID"
            toolTip1.SetToolTip(Me.chkID, str1)

            str1 = "Date on which report was created in StudyDoc"
            toolTip1.SetToolTip(Me.chkDTCreated, str1)

            str1 = "Date on which report is printed, opened in Word, or opened in PDF"
            toolTip1.SetToolTip(Me.chkDTReported, str1)

            str1 = "The user name of the person who printed, opened in Word, or opened in PDF"
            toolTip1.SetToolTip(Me.chkGenerator, str1)

            str1 = "The user name of the owner of the document within StudyDoc"
            toolTip1.SetToolTip(Me.chkOwner, str1)

          
        Catch ex As Exception

        End Try

    End Sub

    Private Sub cmdSelectAll_Click(sender As Object, e As EventArgs) Handles cmdSelectAll.Click

        Call CheckUnCheck(True)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Call CheckUnCheck(False)

    End Sub

    Sub CheckUnCheck(boolCheck As Boolean)

        Me.chkID.Checked = boolCheck
        Me.chkDTCreated.Checked = False
        Me.chkDTReported.Checked = False
        Me.chkGenerator.Checked = boolCheck
        Me.chkOwner.Checked = boolCheck

    End Sub

    
End Class