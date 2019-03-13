Public Class frmShowSymbol

    Public frm As Form

    Private Sub frmShowSymbol_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call FilllbxSymbol(Me.lbxSymbol, Me.lblSymbol)

    End Sub

    Private Sub cmdSymbol_Click(sender As Object, e As EventArgs) Handles cmdSymbol.Click

        Me.Dispose()

    End Sub

    Private Sub lbxSymbol_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbxSymbol.SelectedIndexChanged

        Dim var1

        var1 = Me.lbxSymbol.SelectedItem
        If StrComp(var1, "nbh", CompareMethod.Text) = 0 Then
            var1 = ChrW(2011) 'NBH
        ElseIf StrComp(var1, "nbsp", CompareMethod.Text) = 0 Then
            var1 = ChrW(160)
        ElseIf StrComp(var1, "CR", CompareMethod.Text) = 0 Then
            var1 = ChrW(10)
        Else
        End If
        Me.txtSymbol.Select()
        Me.txtSymbol.Text = var1
        'SendKeys.Send("+{END}")

        Me.txtSymbol.SelectAll()


    End Sub
End Class