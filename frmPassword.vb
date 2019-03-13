Public Class frmLogon
    Public boolCancel As Boolean
    Public idP As Long
    Public idU As Long

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        boolCancel = False
        Dim varU, varP, var1
        Dim tbl As DataTable
        Dim tblP As DataTable
        Dim row() As DataRow
        Dim rowP() As DataRow
        'Dim frmH As New frmHome_01
        Dim strF As String
        Dim boolGo As Boolean
        Dim strM As String
        Dim strM1 As String
        Dim str3 As String
        Dim int1 As Short

        varU = Me.txtUserID.Text
        varP = Me.txtPassword.Text
        boolGo = False

        tbl = frmH.tblUserAccounts
        int1 = tbl.Rows.Count
        tblP = frmH.tblPersonnel
        strF = "charUserID = '" & NZ(varU, "") & "'"
        row = tbl.Select(strF)
        strM = "Message"
        strM1 = "Message"
        If row.Length = 0 Then
            strM = "Invalid login credentials"
            strM1 = "Invalid login credentials..."
        Else
            'record idu
            idU = row(0).Item("id_tblUserAccounts")
            'check to see if userid is active
            If row(0).Item("boolActive") Then
                'check to see if user is active
                idP = row(0).Item("id_tblPersonnel")
                strF = "id_tblPersonnel = " & idP & " AND boolActive = " & True
                rowP = tblP.Select(strF)
                If rowP.Length = 0 Then
                    str3 = rowP(0).Item("charFirstName") & NZ(rowP(0).Item("charMiddleName"), "") & rowP(0).Item("charLastName")
                    strM = "UserID is associated with inactive user " & str3 & "."
                    strM1 = "Inactive User..."
                Else
                    boolGo = True
                End If
            Else
                strM = "Inactive UserID."
                strM1 = "Inactive UserID..."
            End If

        End If

        If boolGo Then
            Me.Visible = False
        Else
            MsgBox(strM, MsgBoxStyle.Information, strM1)
        End If

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        boolCancel = True
        Me.Visible = False
    End Sub
End Class