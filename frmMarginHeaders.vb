Option Compare Text

Public Class frmMarginHeaders

    Public boolCancel As Boolean = True


    Private Sub rbMarginGuWu_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbMarginGuWu.CheckedChanged

        Call ParamsVis()

    End Sub


    Private Sub rbMarginCustom_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbMarginCustom.CheckedChanged

        Call ParamsVis()

    End Sub

    Sub ParamsVis()

        If Me.rbMarginCustom.Checked Then
            Me.panParameters.Enabled = True
        Else
            Me.panParameters.Enabled = False
        End If

    End Sub

    Sub FormLoad()



    End Sub

    Sub SetBools()

        'Legend
        'Public boolFlipHeaderAuto As Boolean = False
        'Public boolHeaderIsText As Boolean = True
        'Public boolFooterIsText As Boolean = True

        If Me.rbMarginGuWu.Checked Then
            boolFlipHeaderAuto = True
        Else
            boolFlipHeaderAuto = False
        End If

        If Me.rbHeaderIsText.Checked Then
            boolHeaderIsText = True
        Else
            boolHeaderIsText = False
        End If

        If Me.rbFooterIsText.Checked Then
            boolFooterIsText = True
        Else
            boolFooterIsText = False
        End If

    End Sub

    Private Sub cmdOK_Click(sender As System.Object, e As System.EventArgs) Handles cmdOK.Click

        boolCancel = False

        Me.Visible = False

    End Sub

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click

        boolCancel = True

        Me.Visible = False

    End Sub

    Function ValChar(var1) As Boolean

        ValChar = True

        Dim strM As String

        If IsNumeric(var1) Then
            ValChar = False
        Else
            ValChar = True
            strM = "Entry must by numeric"
            MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
        End If

    End Function

    Private Sub frmMarginHeaders_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        Call ControlDefaults(Me)

        'populate cbxHorizontal Header
        Me.charHeaderHoriz.Items.Clear()
        Me.charHeaderHoriz.Items.Add("Margin")
        Me.charHeaderHoriz.Items.Add("Page")
        Me.charHeaderHoriz.Items.Add("Column")
        Me.charHeaderHoriz.Items.Add("Character")
        Me.charHeaderHoriz.Items.Add("Left Margin")
        Me.charHeaderHoriz.Items.Add("Right Margin")
        Me.charHeaderHoriz.Items.Add("Inside Margin")
        Me.charHeaderHoriz.Items.Add("Outside Margin")

        Me.charHeaderHoriz.SelectedIndex = 2


        'populate cbxVertical Header
        Me.charHeaderVert.Items.Clear()
        Me.charHeaderVert.Items.Add("Margin")
        Me.charHeaderVert.Items.Add("Page")
        Me.charHeaderVert.Items.Add("Paragraph")
        Me.charHeaderVert.Items.Add("Line")
        Me.charHeaderVert.Items.Add("Top Margin")
        Me.charHeaderVert.Items.Add("Bottom Margin")
        Me.charHeaderVert.Items.Add("Inside Margin")
        Me.charHeaderVert.Items.Add("Outside Margin")

        Me.charHeaderHoriz.SelectedIndex = 2

        'default horiz pos:  7.79
        Me.numHeaderAbsPosHoriz.Text = "7.79"

        'default vert pos:  0.04
        Me.numHeaderAbsPosVert.Text = "0.04"

        'populate cbxHorizontal Footer
        Me.charFooterHoriz.Items.Clear()
        Me.charFooterVert.Items.Add("Margin")
        Me.charFooterVert.Items.Add("Page")
        Me.charFooterVert.Items.Add("Column")
        Me.charFooterVert.Items.Add("Character")
        Me.charFooterVert.Items.Add("Left Margin")
        Me.charFooterVert.Items.Add("Right Margin")
        Me.charFooterVert.Items.Add("Inside Margin")
        Me.charFooterVert.Items.Add("Outside Margin")

        Me.charFooterVert.SelectedIndex = 2


        'populate cbxVertical Footer
        Me.charFooterVert.Items.Clear()
        Me.charFooterVert.Items.Add("Margin")
        Me.charFooterVert.Items.Add("Page")
        Me.charFooterVert.Items.Add("Paragraph")
        Me.charFooterVert.Items.Add("Line")
        Me.charFooterVert.Items.Add("Top Margin")
        Me.charFooterVert.Items.Add("Bottom Margin")
        Me.charFooterVert.Items.Add("Inside Margin")
        Me.charFooterVert.Items.Add("Outside Margin")

        Me.charFooterVert.SelectedIndex = 2

        'default horiz pos:  7.79
        Me.numFooterAbsPosHoriz.Text = "7.79"

        'default vert pos:  0.04
        Me.numFooterAbsPosVert.Text = "0.04"

    End Sub


    Private Sub numHeaderAbsPosHoriz_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles numHeaderAbsPosHoriz.Validating

        Dim var1

        var1 = Me.numHeaderAbsPosHoriz.Text

        e.Cancel = ValChar(var1)

    End Sub

    Private Sub numHeaderAbsPosVert_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles numHeaderAbsPosVert.Validating

        Dim var1

        var1 = Me.numHeaderAbsPosVert.Text

        e.Cancel = ValChar(var1)

    End Sub

    Private Sub numFooterAbsPosHoriz_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles numFooterAbsPosHoriz.Validating

        Dim var1

        var1 = Me.numFooterAbsPosHoriz.Text

        e.Cancel = ValChar(var1)

    End Sub


    Private Sub numFooterAbsPosVert_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles numFooterAbsPosVert.Validating

        Dim var1

        var1 = Me.numFooterAbsPosVert.Text

        e.Cancel = ValChar(var1)

    End Sub
End Class