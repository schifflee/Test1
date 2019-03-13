Option Compare Text

Public Class frmESig

    Public boolCancel As Boolean = True
    Public idP As Long
    Public idU As Long
    Public intAtt As Short = 0
    Public intAttMax As Short = 0
    Public intMinutesMax As Short = 0
    Public tUserID As String
    Public tUserName As String
    Public boolTest As Boolean = False


    Sub Establish_intAttMax()

        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String

        tbl = tblConfiguration
        strF = "charConfigCategory = 'Password Settings' AND charConfigTitle = 'Number of login attempts allowed'"
        rows = tbl.Select(strF)
        intAttMax = CInt(rows(0).Item("charConfigValue"))

        strF = "charConfigCategory = 'Password Settings' AND charConfigTitle = 'Password change restriction (minutes)'"
        rows = tbl.Select(strF)
        intMinutesMax = CInt(rows(0).Item("charConfigValue"))


    End Sub

    Private Sub frmESig_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call ControlDefaults(Me)

        If gboolLDAP Then
            'Me.txtUserID.Text = gNetAcct
            Me.lblUserID.Text = "Enter Network User ID:"
            Me.lblPassword.Text = "Enter Network Password:"
        Else
            'Me.txtUserID.Text = gUserID
            Me.lblUserID.Text = "Enter StudyDoc User ID:"
            Me.lblPassword.Text = "Enter StudyDoc Password:"
        End If

        Dim str1 As String

        str1 = "the eSig window is configured as Silent Audit Trail." & ChrW(10)
        str1 = str1 & "The default Meaning of Signature (if configured to be shown)" & ChrW(10)
        str1 = str1 & "and Reason for Change value (if configured to be shown) that will be recorded in the Audit Trail are shown below."

        Me.lblTest.Text = str1

        str1 = "The combination of the Meaning of Signature dropdown box and the Free Form text box will be recorded as 'Meaning of Signature'"
        Me.lblMOSE.Text = str1

        str1 = "The combination of the Reason for Change dropdown box and the Free Form text box will be recorded as 'Reason for Change'"
        Me.lblRFCE.Text = str1

        Call Establish_intAttMax()

        If boolTest Then
            Call EvalComplianceTest()
        Else
            Call EvalCompliance()
        End If

        Call PlaceC()

        Me.txtUserID.Focus()

    End Sub


    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        'check if password is correct

        'Dim strUID As String
        'Dim idUID As String
        'Dim tblUID as System.Data.Datatable = tblUserAccounts
        'Dim rowsT() As DataRow
        'Dim strF As String

        'strUID = Me.txtUserID.Text
        'strF = "CHARUSERID = '" & strUID & "'"
        'rowsT = tblUID.Select(strF)

        Dim strF As String
        Dim varU, varP, var1, var2, var3
        Dim tbl As System.Data.DataTable
        Dim tblP As System.Data.DataTable
        Dim tblC As System.Data.DataTable
        Dim row() As DataRow
        Dim rowP() As DataRow
        'Dim frmH As New frmHome_01
        Dim frmCP As New frmPasswordChange
        Dim boolGo As Boolean
        Dim strM As String
        Dim strM1 As String
        Dim str3 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim Count1 As Short
        Dim varPE1, varPE2


        'User ID can be either a StudyDoc userid or a network account

  
        varP = Me.txtPassword.Text
        If Len(varP) = 0 Then
            boolGo = False
            strM = "Blank password not allowed"
            strM1 = "Invalid login credentials..."
            MsgBox(strM, MsgBoxStyle.Information, strM1)

            GoTo end1
        End If
        boolGo = False

        varU = Me.txtUserID.Text

        tUserID = varU

        strM = "Message"
        strM1 = "Message"

        tbl = tblUserAccounts
        tblP = tblPersonnel
        tblC = tblConfiguration

        If gboolLDAP Then
            strF = "CHARNETWORKACCOUNT = '" & NZ(varU, "") & "'"
        Else
            strF = "charUserID = '" & NZ(varU, "") & "'"
        End If

        row = tbl.Select(strF)
        '20180707 LEE:
        'Need to check old value
        row = tbl.Select(strF, "", DataViewRowState.ModifiedOriginal)
        If row.Length = 0 Then
            '20180802 LEE:
            'check for current
            row = tbl.Select(strF, "", DataViewRowState.CurrentRows)
        End If
        If row.Length = 0 Then
            strM = "Invalid ESig credentials"
            strM1 = "Invalid ESig credentials..."
            boolGo = False
            MsgBox(strM, MsgBoxStyle.Information, strM1)
            Me.txtPassword.Text = ""
            Me.txtPassword.Select()

            GoTo end1
        Else
            tUserID = row(0).Item("CHARUSERID")
            'record idu
            idU = row(0).Item("id_tblUserAccounts")
            'check to see if userid is associated with an active user
            'check to see if user is active
            idP = row(0).Item("id_tblPersonnel")
            strF = "id_tblPersonnel = " & idP & " AND boolActive = -1" ' & True
            rowP = tblP.Select(strF)
            If rowP.Length = 0 Then
                'str3 = rowP(0).Item("charFirstName") & NZ(rowP(0).Item("charMiddleName"), "") & rowP(0).Item("charLastName")
                strM = "UserID is associated with inactive user " & str3 & "."
                strM1 = "Inactive User..."
                boolGo = False

                GoTo end1

            Else

                'Don't need to do this anymore.
                'Dim strA As String
                'Dim strB As String
                'Dim strC As String

                'strA = NZ(rowP(0).Item("charFirstName"), "")
                'strB = NZ(rowP(0).Item("charMiddleName"), "")
                'strC = NZ(rowP(0).Item("charLastName"), "")
                'If Len(strB) = 0 Then
                '    tUserName = strA & " " & strC
                'Else
                '    tUserName = strA & " " & strB & " " & strC
                'End If

                tUserName = gUserName
                boolGo = True
            End If

            If row(0).Item("boolActive") = 0 Then 'check to see if user id is active
                strM = "Inactive UserID."
                strM1 = "Inactive UserID..."
                boolGo = False
                MsgBox(strM, MsgBoxStyle.Information, strM1)
                Me.txtPassword.Text = ""
                Me.txtPassword.Select()

                GoTo end1

            Else 'account is active
                boolGo = True
            End If

        End If

        'now check for correct password
        If boolGo Then


            '20180722 LEE:
            'unneccesarily complex evalution
            'just use binary

            varPE2 = gPswd

            If StrComp(varP, varPE2, CompareMethod.Binary) = 0 Then

            Else
                boolGo = False
                strM = "Invalid ESig credentials"
                strM1 = "Invalid ESig credentials..."
                MsgBox(strM, MsgBoxStyle.Information, strM1)
                Me.txtPassword.Text = ""
                Me.txtPassword.Select()
            End If

            ''varPE1 = NZ(row(0).Item("charPassword"), 0)
            ''varPE1 = Decode(varPE1, True)
            ''varPE2 = Coding(varPE1, False)

            'varPE2 = gPswd
            'int1 = Len(NZ(varPE2, ""))
            'int2 = Len(NZ(varP, ""))
            ''20180722 LEE:
            ''Incorrect logic. Should be <>
            ''If int1 = int2 Then 'continue
            'If int1 <> int2 Then 'continue
            '    boolGo = False
            '    strM = "Invalid ESig credentials"
            '    strM1 = "Invalid ESig credentials..."
            '    MsgBox(strM, MsgBoxStyle.Information, strM1)
            '    Me.txtPassword.Text = ""
            '    Me.txtPassword.Select()

            'Else
            '    For Count1 = 1 To int1
            '        var2 = AscW(Mid(varPE2, Count1, 1))
            '        var3 = AscW(Mid(varP, Count1, 1))
            '        If var2 = var3 Then 'continue
            '        Else
            '            boolGo = False
            '            strM = "Invalid ESig credentials"
            '            strM1 = "Invalid ESig credentials..."
            '            MsgBox(strM, MsgBoxStyle.Information, strM1)
            '            Me.txtPassword.Text = ""
            '            Me.txtPassword.Select()
            '            Exit For
            '        End If
            '    Next

            '    var1 = var1

            'End If
            

        Else

            boolGo = False
            strM = "Invalid ESig credentials"
            strM1 = "Invalid ESig credentials..."
            MsgBox(strM, MsgBoxStyle.Information, strM1)
            Me.txtPassword.Text = ""
            Me.txtPassword.Select()

            GoTo end1

        End If

        Dim boolPC As Short
        boolPC = 0
        If boolGo Then 'check to ensure user isn't locked out
            boolPC = NZ(row(0).Item("boolAccountIsLockedOut"), 0)
            If boolPC = -1 Then
                boolGo = False
                MsgBox("This account is locked out. Please contact your StudyDoc Administrator.", MsgBoxStyle.Information, "Account is locked out...")
            Else
                boolGo = True
            End If
        End If

end1:

        If boolGo Then
            tUserName = gUserName 'FindUserName(idP)
            boolCancel = False
            Call GetRFC()
            Call GetMOS()
            Me.Visible = False
        Else
            intAtt = intAtt + 1
            If intAtt >= intAttMax Then
                MsgBox("You have exceeded the maximum number of password attempts allowed.", MsgBoxStyle.Information, "Maxed out credential attempts...")
                boolCancel = True
                Me.Visible = False
            End If

        End If


    End Sub

    Sub GetRFC()

        Dim var1, var2
        Dim str1 As String

        Dim intRFC As Short

        intRFC = tblConfigCompliance.Rows(0).Item("BOOLREASONFORCHANGE")

        If intRFC = 0 Then
            str1 = "[Reason For Change option disabled]"
        Else
            var1 = NZ(Me.cbxRFC.Text, "")
            var2 = NZ(Me.txtRFC.Text, "")

            If Me.cbxRFC.Visible Then
                If Me.txtRFC.Visible Then
                    If Len(var2) = 0 Then
                        str1 = var1
                    Else
                        str1 = var1 & ". " & var2
                    End If
                Else
                    str1 = var1
                End If
            Else
                str1 = ""
            End If
        End If


        strRFC = str1

    End Sub

    Sub GetMOS()

        Dim var1, var2
        Dim str1 As String

        Dim intMOS As Short

        intMOS = tblConfigCompliance.Rows(0).Item("BOOLMEANINGOFSIG")

        If intMOS = 0 Then
            str1 = "[Meaning of Signature option disabled]"
        Else
            var1 = NZ(Me.cbxMOS.Text, "")
            var2 = NZ(Me.txtMOS.Text, "")

            If Me.cbxMOS.Visible Then
                If Me.txtMOS.Visible Then
                    If Len(var2) = 0 Then
                        str1 = var1
                    Else
                        str1 = var1 & ". " & var2
                    End If
                Else
                    str1 = var1
                End If
            Else
                str1 = ""
            End If
        End If


        strMOS = str1

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        boolCancel = True
        Me.Visible = False

    End Sub

    Sub EvalCompliance()


        Call FillRFC_MOS()


        Dim dtbl As System.Data.DataTable = tblConfigCompliance
        Dim var1

        'show meaning of sig
        var1 = dtbl.Rows(0).Item("BOOLMEANINGOFSIG")
        If var1 = 0 Then
            Me.panMOS.Visible = False
        Else
            Me.panMOS.Visible = True
        End If

        'show reason for change
        var1 = dtbl.Rows(0).Item("BOOLREASONFORCHANGE")
        If var1 = 0 Then
            Me.panRFC.Visible = False
        Else
            Me.panRFC.Visible = True
        End If

        'evaluate userid
        var1 = dtbl.Rows(0).Item("BOOLLOGGEDONUSER")
        'If var1 = 0 Then
        '    'Me.txtUserID.Text = ""
        '    Me.txtUserID.Visible = False
        '    'Me.txtUserName.Text = ""
        'Else
        '    Me.txtUserID.Text = gUserID
        '    Me.txtUserID.ReadOnly = True
        '    Me.txtUserID.Visible = True
        '    Me.txtUserName.Text = gUserName
        'End If

        '20170403 LEE: Logic is that only logged on user can enter credentials

        Me.txtUserID.Text = gUserID
        Me.txtUserID.ReadOnly = True
        Me.txtUserID.Visible = True
        Me.txtUserName.Text = gUserName
        Me.txtUserName.ReadOnly = True

        Me.txtPassword.Focus()

        'evaluate MOS choice
        var1 = dtbl.Rows(0).Item("BOOLRESTRICTSIG")
        Me.txtMOS.Text = ""
        If var1 = 0 Then
            Me.panMOSE.Visible = True
        Else
            Me.panMOSE.Visible = False
        End If

        'evaluate MOS choice
        var1 = dtbl.Rows(0).Item("BOOLRESTRICTREASON")
        Me.txtRFC.Text = ""
        If var1 = 0 Then
            Me.panRFCE.Visible = True
        Else
            Me.panRFCE.Visible = False
        End If


    End Sub

    Sub EvalComplianceTest()

        Call FillRFC_MOS()

        Dim dtbl As System.Data.DataTable = tblConfigCompliance
        Dim var1

        'show meaning of sig
        'var1 = dtbl.Rows(0).Item("BOOLMEANINGOFSIG")
        If Me.chkMOS.Checked = False Then
            Me.panMOS.Visible = False
        Else
            Me.panMOS.Visible = True

            Me.txtMOS.Text = ""
            If Me.chkRMOS.Checked Then
                Me.panMOSE.Visible = False
            Else
                Me.panMOSE.Visible = True
            End If

        End If

        'show reason for change
        'var1 = dtbl.Rows(0).Item("BOOLREASONFORCHANGE")
        If Me.chkRFC.Checked = False Then
            Me.panRFC.Visible = False
        Else
            Me.panRFC.Visible = True

            Me.txtRFC.Text = ""
            If Me.chkRRFC.Checked Then
                Me.panRFCE.Visible = False
            Else
                Me.panRFCE.Visible = True
            End If

        End If

    End Sub

    Sub PlaceC()

        Dim boolRFC As Boolean
        Dim boolRFCE As Boolean
        Dim boolMOS As Boolean
        Dim boolMOSE As Boolean

        If boolTest Then
            'boolRFC = Me.chkRFC.Checked
            'boolMOS = Me.chkMOS.Checked
            If Me.rbESigOn.Checked Then
                Me.pan1.Top = Me.lblTest.Top
                Me.lblTest.Visible = False
            Else
                Me.pan1.Top = Me.lblTest.Top + Me.lblTest.Height + 10
                Me.lblTest.Visible = True
            End If
        Else
            'boolRFC = Me.panRFC.Visible
            'boolMOS = Me.panMOS.Visible
            Me.pan1.Top = Me.lblTest.Top
            Me.lblTest.Visible = False
        End If

        boolRFC = Me.panRFC.Visible
        boolRFCE = Me.panRFCE.Visible

        boolMOS = Me.panMOS.Visible
        boolMOSE = Me.panMOSE.Visible

        If boolMOS Then

            If boolMOSE Then

                If boolRFC Then

                    If boolRFCE Then
                        Me.panOK.Top = Me.panRFC.Top + Me.panRFC.Height + 10
                    Else
                        Me.panOK.Top = Me.panRFC.Top + (Me.cbxRFC.Top + Me.cbxRFC.Height) + 10
                    End If

                Else

                    Me.panOK.Top = Me.panMOS.Top + Me.panMOS.Height + 10

                End If

            Else

                Me.panRFC.Top = Me.panMOS.Top + (Me.cbxMOS.Top + Me.cbxMOS.Height)

                If boolRFC Then

                    If boolRFCE Then
                        Me.panOK.Top = Me.panRFC.Top + Me.panRFC.Height
                    Else
                        Me.panOK.Top = Me.panRFC.Top + (Me.cbxRFC.Top + Me.cbxRFC.Height) + 10
                    End If

                Else

                    Me.panOK.Top = Me.panMOS.Top + (Me.cbxMOS.Top + Me.cbxMOS.Height) + 10

                End If

            End If

        Else

            Me.panRFC.Top = Me.panMOS.Top

            If boolRFC Then
                If boolRFCE Then
                    Me.panOK.Top = Me.panRFC.Top + Me.panRFC.Height + 10
                Else
                    Me.panOK.Top = Me.panRFC.Top + (Me.cbxRFC.Top + Me.cbxRFC.Height) + 10
                End If
            Else

                Me.panOK.Top = Me.panCred.Top + Me.panCred.Height + 10
            End If

        End If

        Me.pan1.Height = Me.panOK.Top + Me.panOK.Height + 50

        Me.Height = Me.pan1.Top + Me.pan1.Height ' + 50

        Me.panMOS.BringToFront()
        Me.panRFC.BringToFront()
        Me.panOK.BringToFront()


    End Sub

    Function FindUserName(ByVal idP As Int64)

        FindUserName = "[NA]"

        Dim rows() As DataRow
        Dim strF As String
        Dim rowsP() As DataRow
        Dim var1, var2, var3, var4

        Erase rowsP
        strF = "ID_TBLPERSONNEL = " & idP
        rowsP = tblPersonnel.Select(strF)
        If rowsP.Length = 0 Then
            var4 = "[NA]"
        Else
            var1 = NZ(rowsP(0).Item("CHARFIRSTNAME"), "")
            var2 = NZ(rowsP(0).Item("CHARMIDDLENAME"), "")
            var3 = NZ(rowsP(0).Item("CHARLASTNAME"), "")

            If Len(var2) = 0 Then
                var4 = var1 & " " & var3
            Else
                'If Len(var2) = 1 Then
                '    var4 = var1 & " " & var2 & ". " & var3
                'Else
                '    var4 = var1 & " " & var2 & " " & var3
                'End If
                var4 = var1 & " " & var2 & " " & var3
            End If
        End If

        FindUserName = CStr(var4)

    End Function

    Sub FillRFC_MOS()

        Dim strF As String
        Dim strS As String
        Dim str1 As String
        Dim Count1 As Short
        Dim rows() As System.Data.DataRow
        Dim var1, var2
        Dim dtbl As System.Data.DataTable
        Dim boolHit As Boolean

        strF = "INTORDER > -100"
        strS = "INTORDER ASC"

        'do Reason For Change
        dtbl = tblReasonForChange

        rows = dtbl.Select(strF, strS)

        Me.cbxRFC.Items.Clear()
        'Me.cbxRFC.Items.Add("")
        For Count1 = 0 To rows.Length - 1
            var1 = NZ(dtbl.Rows(Count1).Item("CHARREASONFORCHANGE"), "")
            Me.cbxRFC.Items.Add(CStr(var1))
        Next

        'find default setting
        boolHit = False
        For Count1 = 0 To rows.Length - 1
            var1 = NZ(dtbl.Rows(Count1).Item("BOOLDEFAULT"), 0)
            var2 = dtbl.Rows(Count1).Item("CHARREASONFORCHANGE")
            If var1 = 0 Then
            Else
                Me.cbxRFC.SelectedIndex = Count1 ' + 1 'account for added ""
                boolHit = True
                Exit For
            End If
        Next

        If boolHit Then
        Else
            Me.cbxRFC.SelectedIndex = 0
        End If

        'do Meaning of Sig
        dtbl = tblMeaningOfSig

        Erase rows

        rows = dtbl.Select(strF, strS)

        Me.cbxMOS.Items.Clear()
        'Me.cbxMOS.Items.Add("")
        For Count1 = 0 To rows.Length - 1
            var1 = NZ(dtbl.Rows(Count1).Item("CHARMEANINGOFSIG"), "")
            Me.cbxMOS.Items.Add(CStr(var1))
        Next

        'find default setting
        boolHit = False
        For Count1 = 0 To rows.Length - 1
            var1 = NZ(dtbl.Rows(Count1).Item("BOOLDEFAULT"), 0)
            var2 = dtbl.Rows(Count1).Item("CHARMEANINGOFSIG")
            If var1 = 0 Then
            Else
                Try
                    Me.cbxMOS.SelectedIndex = Count1 ' + 1 'account for added ""
                Catch ex As Exception
                    str1 = ex.Message
                    str1 = str1
                End Try

                boolHit = True
                Exit For
            End If
        Next

        If boolHit Then
        Else
            Me.cbxMOS.SelectedIndex = 0
        End If




    End Sub

    Private Sub txtMOS_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMOS.TextChanged

    End Sub

    Private Sub txtMOS_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtMOS.Validating

        Dim str1 As String

        str1 = Me.txtMOS.Text
        If Len(str1) > 200 Then
            Dim strE As String
            strE = "This entry must be 200 characters or less"
            MsgBox(strE, MsgBoxStyle.Information, "Invalid entry...")
            e.Cancel = True
        End If

    End Sub

    Private Sub panRFC_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles panRFC.Validating

        Dim str1 As String

        str1 = Me.txtRFC.Text
        If Len(str1) > 200 Then
            Dim strE As String
            strE = "This entry must be 200 characters or less"
            MsgBox(strE, MsgBoxStyle.Information, "Invalid entry...")
            e.Cancel = True
        End If

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        Dim var1, var2, var3, var4

        var1 = Me.panMOS.Top
        var2 = Me.panRFC.Top
        var3 = Me.panOK.Top

        var4 = "MOS: " & var1 & ChrW(10) & "RFC: " & var2 & ChrW(10) & "OK: " & var3
        MsgBox(var4)

    End Sub
End Class