Option Compare Text

Public Class frmCode

    Private Sub cmdEncrypt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEncrypt.Click

        Dim str1 As String
        Dim var1, var2, var3
        Dim Count1 As Short
        Dim int1 As Short
        Dim varP, varPA

        str1 = Me.txtPassword.Text.ToString
        varP = Coding(str1, True)
        Me.txtEncrypt.Text = varP.ToString

        'record ascii code
        varPA = Decode(varP, False)
        Me.txtAscii.Text = varPA

        'return decrypted
        str1 = Me.txtEncrypt.Text
        var1 = Coding(str1, False)
        Me.txtDeEncrypt.Text = var1.ToString

        Exit Sub

        'record varP in Oracle
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        tbl = tblUserAccounts
        strF = "id_tblUserAccounts = 1"
        rows = tbl.Select(strF)
        rows(0).BeginEdit()
        rows(0).Item("charPassword") = varPA
        rows(0).EndEdit()

        If boolGuWuOracle Then
            ta_tblUserAccounts.Update(tblUserAccounts)
            ta_tblUserAccounts.Fill(tblUserAccounts)
        ElseIf boolGuWuAccess Then
            ta_tblUserAccountsAcc.Update(tblUserAccounts)
            ta_tblUserAccountsAcc.Fill(tblUserAccounts)
        ElseIf boolGuWuSQLServer Then
            ta_tblUserAccountsSQLServer.Update(tblUserAccounts)
            ta_tblUserAccountsSQLServer.Fill(tblUserAccounts)
        End If

        'retrieve Oracle value
        Erase rows
        Dim int2 As Short
        Dim int3 As Short
        Dim intS As Short
        Dim intE As Short
        Dim intE1 As Short

        rows = tbl.Select(strF)
        varPA = NZ(rows(0).Item("charpassword"), "0")
        Me.txtEncryptO.Text = Decode(varPA, True)

        'record ascii code Oracle
        'var1 = Me.txtEncryptO.Text
        'Me.txtEncryptO.Text = Coding(var1, False)

        'return decrypted Oracle
        str1 = Me.txtEncryptO.Text
        var1 = Coding(str1, False)
        Me.txtDeEncryptO.Text = var1.ToString


    End Sub

    Private Sub cmdDeEncrypt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeEncrypt.Click

        Dim str1 As String
        Dim var1

        str1 = Me.txtEncrypt.Text
        var1 = Coding(str1, False)
        Me.txtDeEncrypt.Text = var1.ToString


    End Sub
End Class