Option Compare Text

Public Class frmPasswordChange
    Public boolCancel As Boolean
    Public boolFromAdmin As Boolean = False
    Public strPswd As String
    Public boolFromChgPswd As Boolean = False

    Private Sub frmPasswordChange_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call ControlDefaults(Me)

        'these variables will be assigned by the calling routine
        'boolFromAdmin = False
        'boolFromChgPswd = False

        boolCancel = True
        Dim var1, var2

        'var1 = frmH.cmdChangePassword.Left
        'var2 = frmH.cmdChangePassword.Top

        'Me.Top = var2
        'Me.Left = var1 + frmH.cmdChangePassword.Width

    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        Dim var1, var2, var3, var4, varP1, varP2
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim str1 As String
        Dim varPswd
        Dim dt As Date
        Dim int1 As Short
        Dim int2 As Short
        Dim strM As String
        Dim strM1 As String
        Dim Count1 As Short
        Dim boolGo As Boolean
        Dim tblC As System.Data.DataTable
        Dim rowC() As DataRow
        Dim tblPW As System.Data.DataTable
        Dim rowPW() As DataRow
        Dim tblPH As System.Data.DataTable
        Dim rowPH() As DataRow
        Dim varPE1, varPE2, varPO
        Dim boolS As Boolean
        Dim strME As String

        tbl = tblUserAccounts
        strF = "id_tblUserAccounts = " & Me.txtID.Text ' frmH.id_tblUserAccounts
        rows = tbl.Select(strF)
        tblC = tblConfiguration
        boolS = False

        boolFromAdmin = Me.chkFromAdmin.Checked
        boolFromChgPswd = Me.chkFromChgPswd.Checked


        'Dim strME As String
        strME = "Password change unsuccessful"

        'ensure old password is correct
        var1 = NZ(rows(0).Item("charPassword"), "")
        If Len(var1) = 0 Then
            varPE1 = ""
            varPE2 = ""
        Else
            varPE1 = Decode(rows(0).Item("charPassword"), True)
            'decrypt
            varPE2 = Coding(varPE1, False)
        End If
        varPO = NZ(Me.txtOldPassword.Text, "")
        boolGo = True
        boolCancel = True

        If boolGo And boolFromAdmin = False Then

            If StrComp(varPO, varPE2, CompareMethod.Binary) = 0 Then
            Else
                boolGo = False
                strM = "An incorrect old password has been entered."
                strM1 = "Incorrect Password..."
                Me.txtOldPassword.Select()
            End If
        End If

        If boolGo Then
        Else
            MsgBox(strM, MsgBoxStyle.Information, strM1)
            Exit Sub
        End If

        'ensure new and confirm password match
        varP1 = Me.txtNewPassword.Text
        varP2 = Me.txtConfirmPassword.Text
        boolGo = True
        If boolGo Then

            If StrComp(varP1, varP2, CompareMethod.Binary) = 0 Then
            Else
                boolGo = False
                strM = "New and confirm passwords don't match."
                strM1 = "Passwords don't match..."
                Me.txtNewPassword.Text = ""
                Me.txtConfirmPassword.Text = ""
                Me.txtNewPassword.Select()
            End If

        End If

        If boolGo Then
        Else
            MsgBox(strM, MsgBoxStyle.Information, strM1)
            Exit Sub
        End If

        'now check to ensure that new password is not the same as old password
        If boolGo Then
            If StrComp(varP2, varPE2, CompareMethod.Binary) = 0 Then
                boolGo = False
                strM = "New and old passwords must not match."
                strM1 = "Passwords are identical..."
                Me.txtNewPassword.Text = ""
                Me.txtConfirmPassword.Text = ""
                Me.txtNewPassword.Select()
            End If
        End If

        If boolGo Then
        Else
            MsgBox(strM, MsgBoxStyle.Information, strM1)
            Exit Sub
        End If

        If boolGo Then 'now check for password length
            strF = "charConfigCategory = 'Password Settings' AND charConfigTitle = 'Minimum password length'"
            rowC = tblC.Select(strF)
            var1 = rowC(0).Item("charConfigValue")
            If Len(varP2) < var1 Then
                strM = "Password length must be " & var1 & " characters or more." ', MsgBoxStyle.Information, "Invalid entry...")
                strM1 = "Invalid entry..."
                boolGo = False
                Me.txtNewPassword.Text = ""
                Me.txtConfirmPassword.Text = ""
                Me.txtNewPassword.Select()
            End If
        End If

        If boolGo Then
        Else
            MsgBox(strM, MsgBoxStyle.Information, strM1)
            Exit Sub
        End If

        If boolGo Then 'check for password integrity

            Dim ctI As Short
            ctI = 0
            strF = "charConfigCategory = 'Password Settings' AND charConfigTitle = 'Enforce password integrity (TRUE or FALSE)'"
            Erase rowC
            rowC = tblC.Select(strF)
            var1 = rowC(0).Item("charConfigValue") 'boolean
            If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then 'check for password integrity
                Dim int3 As Short
                Dim int4 As Short
                int1 = Len(varP2)
                'boolGo = False
                For Count1 = 1 To int1 'check for one upper case letter
                    str1 = Mid(varP2, Count1, 1).ToString
                    var2 = AscW(str1)
                    If var2 > 64 And var2 < 91 Then 'is upper case
                        ctI = ctI + 1
                        Exit For
                    End If
                Next
                For Count1 = 1 To int1 'check for one lower case letter
                    str1 = Mid(varP2, Count1, 1).ToString
                    var2 = AscW(str1)
                    If var2 > 96 And var2 < 123 Then 'is lower case
                        ctI = ctI + 1
                        Exit For
                    End If
                Next
                For Count1 = 1 To int1 'check for one digit
                    str1 = Mid(varP2, Count1, 1).ToString
                    var2 = AscW(str1)
                    If var2 > 47 And var2 < 58 Then 'is capital
                        ctI = ctI + 1
                        Exit For
                    End If
                Next
                For Count1 = 1 To int1 'check for one non-alphanumeric character
                    str1 = Mid(varP2, Count1, 1).ToString
                    var2 = AscW(str1)
                    If (var2 < 48 And var2 > 58) Then 'is not numeric
                        If (var2 < 65 And var2 > 90) Then 'is not capital alpha
                            If (var2 < 97 And var2 > 122) Then 'is not small alpha
                                ctI = ctI + 1
                                Exit For
                            End If
                        End If
                    End If
                Next

                If ctI > 2 Then 'good
                Else
                    str1 = "This password violates the password integrity policy configured by your StudyDoc Administrator. The password integrity criteria are the following:"
                    str1 = str1 & Chr(10) & Chr(10)
                    str1 = str1 & "Password integrity policy enforces that a password content meets at least three of the four following criteria:"
                    str1 = str1 & Chr(10) & Chr(10)
                    str1 = str1 & "      An upper case letter" & Chr(10)
                    str1 = str1 & "      A lower case letter" & Chr(10)
                    str1 = str1 & "      A digit" & Chr(10)
                    str1 = str1 & "      A non-alphanumeric character"
                    strM = str1
                    strM1 = "Password integrity policy violation..."
                    boolGo = False
                    Me.txtNewPassword.Text = ""
                    Me.txtConfirmPassword.Text = ""
                    Me.txtNewPassword.Select()
                End If
            End If
        End If

        If boolGo Then
        Else
            MsgBox(strM, MsgBoxStyle.Information, strM1)
            Exit Sub
        End If

        If boolGo Then 'check for password change history restriction
            Dim strS As String
            tblPH = tblPasswordHistory
            strF = "id_tblUserAccounts = " & Me.txtID.Text
            strS = "dtPassword DESC"
            rowPH = tblPH.Select(strF, strS)
            strF = "charConfigCategory = 'Password Settings' AND charConfigTitle = 'Password history restriction'"
            Erase rowC
            rowC = tblC.Select(strF) '), strS)
            var1 = rowC(0).Item("charConfigValue") 'integer value giving number of passwords to save
            If rowPH.Length > CInt(var1) Then 'look at only last bunch
                int1 = var1
            Else
                int1 = rowPH.Length
            End If
            For Count1 = 0 To int1 - 1
                var1 = NZ(rowPH(Count1).Item("charPassword"), "aaa")
                var2 = Decode(var1, True)
                var3 = Coding(var2, False)
                If StrComp(var3, varP2, CompareMethod.Binary) = 0 Then
                    boolGo = False
                    strM = "This password can't be used. It's use violates the configured Password History Restriction policy."
                    strM1 = "Password History Restriction policy violation..."
                    Me.txtNewPassword.Text = ""
                    Me.txtConfirmPassword.Text = ""
                    Me.txtNewPassword.Select()
                    Exit For
                End If
            Next

        End If

        If boolGo Then
            boolCancel = False
        Else
            MsgBox(strM, MsgBoxStyle.Information, strM1)
            Exit Sub
        End If

        If boolCancel Then
        Else 'update password history

            strPswd = PasswordUnEncrypt(varP2.ToString) ' Decode(Coding(varP2, True), False)

            gPswd = varP2

            'If boolFromAdmin Or boolFromChgPswd Then 'write to database
            If boolFromAdmin Then
                Me.Visible = False
                boolS = False

            Else

                dt = Now

                Call SavePswdHistory(CLng(Me.txtID.Text), strPswd, dt)

                ''update password
                Dim dv As System.Data.DataView
                dv = tblUserAccounts.DefaultView
                strF = "id_tblUserAccounts = " & Me.txtID.Text ' frmH.id_tblUserAccounts
                dv.RowFilter = strF
                dv(0).BeginEdit()
                dv(0).Item("charPassword") = strPswd 'Decode(Coding(varP2, True), False)
                dv(0).Item("dtTimeStamp") = dt
                dv(0).EndEdit()

                'clear audittrailtemp
                tblAuditTrailTemp.Clear()
                idSE = 0
                Call FillAuditTrailTemp(tblUserAccounts)
                If boolGuWuOracle Then
                    Try
                        ta_tblUserAccounts.Update(tblUserAccounts)
                        boolS = True
                        strME = "Password Change Successful"
                    Catch ex As DBConcurrencyException
                        'ds2005.TBLUSERACCOUNTS.Merge('ds2005.TBLUSERACCOUNTS, True)
                        boolS = True
                        strME = "Password Change Successful"
                    End Try
                ElseIf boolGuWuAccess Then
                    Try
                        ta_tblUserAccountsAcc.Update(tblUserAccounts)
                        boolS = True
                        strME = "Password Change Successful"
                    Catch ex As DBConcurrencyException
                        'ds2005Acc.TBLUSERACCOUNTS.Merge('ds2005Acc.TBLUSERACCOUNTS, True)
                        boolS = True
                        strME = "Password Change Successful"
                    End Try
                ElseIf boolGuWuSQLServer Then
                    Try
                        ta_tblUserAccountsSQLServer.Update(tblUserAccounts)
                        boolS = True
                        strME = "Password Change Successful"
                    Catch ex As DBConcurrencyException
                        'ds2005Acc.TBLUSERACCOUNTS.Merge('ds2005Acc.TBLUSERACCOUNTS, True)
                        boolS = True
                        strME = "Password Change Successful"
                    End Try
                End If

                'record tblaudittrailtemp
                Call RecordAuditTrail(False, Now)

                boolS = True
                strME = "Password Change Successful"

            End If

            If boolS Then
                Me.Visible = False
                MsgBox(strME, MsgBoxStyle.Information, "Password change successful...")
            End If
        End If

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        boolCancel = True
        Me.Visible = False
    End Sub

    Private Sub txtOldPassword_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtOldPassword.Validating
        Dim tbl As System.Data.DataTable
        Dim strF As String
        Dim rows() As DataRow
        Dim varPO, varPE1, varPE2
        Dim boolGo As Boolean
        Dim strM As String
        Dim strM1 As String

        If Me.chkFromAdmin.Checked Then
            Exit Sub
        End If

        tbl = tblUserAccounts
        strF = "id_tblUserAccounts = " & Me.txtID.Text ' frmH.id_tblUserAccounts
        rows = tbl.Select(strF)

        'ensure old password is correct
        varPE1 = rows(0).Item("charPassword")
        varPE1 = Decode(varPE1, True)
        varPE2 = Coding(varPE1, False)
        varPO = NZ(Me.txtOldPassword.Text, "")

        If StrComp(varPE2, varPO, CompareMethod.Binary) = 0 Then
        Else
            e.Cancel = True
            strM = "An incorrect old password has been entered."
            strM1 = "Incorrect Password..."
            Me.txtOldPassword.Select()
            MsgBox(strM, MsgBoxStyle.Information, strM1)
        End If
    End Sub
End Class