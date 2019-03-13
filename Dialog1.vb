Imports System.Windows.Forms


Public Class frmPeriodTemp
    Public boolCancel As Boolean = True

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        boolCancel = False
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel

        Me.txtCycles.Text = ""
        Me.txtTP.Text = ""
        Me.txtTF.Text = ""
        Me.txtTemp.Text = ""

        boolCancel = True

        Me.Close()
    End Sub

    Private Sub frmPeriodTemp_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Private Sub txtTP_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtTP.Validating
        'entry must be number
        Dim var1
        Dim str1 As String
        Dim boolM As Boolean

        str1 = "Entry must be integer greater 0"
        var1 = Me.txtTP.Text
        If Len(var1) = 0 Then
            Exit Sub
        End If

        boolM = True
        If IsNumeric(var1) Then
            If IsInt(var1) Then
                'must be greater than 0
                If var1 < 1 Then
                    boolM = True
                Else
                    boolM = False
                End If
            Else
                boolM = True
            End If

        Else
            boolM = True
        End If

        If boolM Then
            Me.txtTP.Focus()
            MsgBox(str1, MsgBoxStyle.Information, "Invalid entry...")
        End If
    End Sub

    Private Sub txtCycles_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCycles.Validating
        'entry must be number
        Dim var1
        Dim str1 As String
        Dim boolM As Boolean

        str1 = "Entry must be integer greater 0"
        var1 = Me.txtCycles.Text
        If Len(var1) = 0 Then
            Exit Sub
        End If
        boolM = True
        If IsNumeric(var1) Then
            If IsInt(var1) Then
                'must be greater than 0
                If var1 < 1 Then
                    boolM = True
                Else
                    boolM = False
                End If
            Else
                boolM = True
            End If

        Else
            boolM = True
        End If

        If boolM Then
            Me.txtCycles.Focus()
            MsgBox(str1, MsgBoxStyle.Information, "Invalid entry...")
        End If
    End Sub
End Class
