Option Compare Text

Imports System.Windows.Forms

Public Class frmIncSmplCrit
    Public boolCancel As Boolean = True

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        'Me.DialogResult = System.Windows.Forms.DialogResult.OK

        'first check validation
        Dim boolC As Boolean
        boolC = CheckValidation()
        If boolC Then
            Exit Sub
        End If

        boolCancel = False
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        boolCancel = True
        Me.Close()
    End Sub

    Private Sub rb2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb2.CheckedChanged
        Call UpdateView()
    End Sub

    Sub UpdateView()
        If Me.rb2.Checked Then
            Me.lblNUMISCRIT1.Text = "Enter Level One Incurred Sample Acceptance Criteria (%):"
            Me.lblNUMISCRIT1LEVEL.Visible = True
            Me.lblNUMISCRIT2.Visible = True
            Me.NUMISCRIT1LEVEL.Visible = True
            Me.NUMISCRIT2.Visible = True
        Else
            Me.lblNUMISCRIT1.Text = "Enter Incurred Sample Acceptance Criteria (%):"
            Me.lblNUMISCRIT1LEVEL.Visible = False
            Me.lblNUMISCRIT2.Visible = False
            Me.NUMISCRIT1LEVEL.Visible = False
            Me.NUMISCRIT2.Visible = False

            Me.NUMISCRIT1LEVEL.Text = Nothing
            Me.NUMISCRIT2.Text = Nothing

        End If
    End Sub

    Function CheckValidation() As Boolean
        Dim var1
        Dim strM As String
        Dim boolM As Boolean


        boolM = False
        CheckValidation = False

        'crit1 must have data
        var1 = Me.NUMISCRIT1.Text

        If IsNumeric(var1) Then 'OK
            If var1 <= 0 Then
                boolM = True
            End If
        Else
            boolM = True
        End If

        If boolM Then
            strM = "Entry must be numeric > 0"
            MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
            CheckValidation = True
            Me.NUMISCRIT1.Focus()
            GoTo end1
        End If

        If Me.rb1.Checked Then

        Else

            var1 = Me.NUMISCRIT1LEVEL.Text
            If IsNumeric(var1) Then 'OK
                If var1 <= 0 Then
                    boolM = True
                End If
            Else
                boolM = True
            End If

            If boolM Then
                strM = "Entry must be numeric > 0"
                MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
                CheckValidation = True
                Me.NUMISCRIT1LEVEL.Focus()
                GoTo end1
            End If

            var1 = Me.NUMISCRIT2.Text
            If IsNumeric(var1) Then 'OK
                If var1 <= 0 Then
                    boolM = True
                End If
            Else
                boolM = True
            End If

            If boolM Then
                strM = "Entry must be numeric > 0"
                MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
                CheckValidation = True
                Me.NUMISCRIT2.Focus()
                GoTo end1
            End If

        End If

end1:

    End Function

    Private Sub NUMISCRIT1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUMISCRIT1.Validating
        'must be numeric or blank
        Dim var1
        Dim strM As String
        Dim boolM As Boolean

        boolM = False

        var1 = Me.NUMISCRIT1.Text

        If IsNumeric(var1) Then 'OK
            If var1 <= 0 Then
                boolM = True
            End If
        Else
            boolM = True
        End If

        If boolM Then
            strM = "Entry must be numeric > 0"
            MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
            e.Cancel = True
        End If
end1:
    End Sub

    Private Sub lblNUMISCRIT2_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles lblNUMISCRIT2.Validating
        'must be numeric or blank
        Dim var1
        Dim strM As String
        Dim boolM As Boolean

        If Me.rb1.Checked Then
            GoTo end1
        End If

        boolM = False

        var1 = Me.NUMISCRIT2.Text

        If IsNumeric(var1) Then 'OK
            If var1 <= 0 Then
                boolM = True
            End If
        Else
            boolM = True
        End If

        If boolM Then
            strM = "Entry must be numeric > 0"
            MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
            e.Cancel = True
        End If
end1:
    End Sub

    Private Sub lblNUMISCRIT1LEVEL_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles lblNUMISCRIT1LEVEL.Validating
        'must be numeric or blank
        Dim var1
        Dim strM As String
        Dim boolM As Boolean

        If Me.rb1.Checked Then
            GoTo end1
        End If

        boolM = False

        var1 = Me.NUMISCRIT1LEVEL.Text

        If IsNumeric(var1) Then 'OK
            If var1 <= 0 Then
                boolM = True
            End If
        Else
            boolM = True
        End If

        If boolM Then
            strM = "Entry must be numeric > 0"
            MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
            e.Cancel = True
        End If
end1:
    End Sub
End Class
