Option Compare Text

Public Class frmLogon
    Public boolCancel As Boolean
    Public idP As Int64
    'Public idU As Int64  declared in modConsserverd. Needed for Logout save
    Public idPerm As Int64
    Public strPerm As String = ""
    Public intAtt As Short = 0
    Public intAttMax As Short = 0
    Public intMinutesMax As Short = 0
    Dim intATTUB As Short = 7
    Dim arrAtt(intATTUB, 1) As Object
    '1=UserIDAtt, 2=id_Userid (-1 if none), 3=dtAttempt, 4=boolSucUID (-1 or 0), 5=boolSucPswd (-1 or 0), 6=boolSucLogin (-1 or 0), 7=Comments

    Dim con As New ADODB.Connection
    Dim rsLogIn As New ADODB.Recordset
    Dim dtRes As Date


    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click


        boolCancel = False
        Dim varU, varP, var1, var2, var3
        Dim tbl As System.Data.DataTable
        Dim tblP As System.Data.DataTable
        Dim tblC As System.Data.DataTable
        Dim tblPerm As System.Data.DataTable
        Dim rowP() As DataRow
        Dim rowPerm() As DataRow
        'Dim frmH As New frmHome_01
        Dim frmCP As New frmPasswordChange
        Dim strF As String
        Dim boolGo As Boolean
        Dim strM As String
        Dim strM1 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim Count1 As Short
        Dim varPE1, varPE2
        Dim tUserID As String
        Dim tUserName As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String

        Dim boolDBSave As Boolean = False
        Dim boolSaveLogin As Boolean = False
        Dim boolPWC As Boolean = False 'password has been changed

        Dim boolClearIDs As Boolean = False

        Dim ts As TimeSpan
        Dim tsd As TimeSpan

        Dim boolE As Boolean = False

        Dim dt As Date
        Dim rows() As DataRow
        Dim rowsC() As DataRow
        'Dim dtRes As Date
        Dim dtRes1 As Date
        Dim bool As Boolean

        Dim dtNow As Date = Now
        Dim dtLastAttempt As Date
        Dim intMinutes As Short

        'query the tblLogIn database table

        Dim boolAttMinutesBad As Boolean = False

        Dim strAdmin As String = "aaAdmin"
        Dim boolAdmin As Boolean = False

        'Dim arrAtt(6, intAttMax) this is redim'd at form load
        ''1=UserIDAtt, 2=id_Userid (-1 if none), 3=dtAttempt, 4=boolSucUID (-1 or 0), 5=boolSucPswd (-1 or 0), 6=boolSucLogin (-1 or 0),
        Dim boolSucUID As Boolean = False
        Dim boolSucPswd As Boolean = False

        gboolUseWatson = False
        gboolLDAP = False
        gNetAcct = ""
        gWatsonAcct = ""


        'first set all permissions to false
        Call SetPermissions(False)


        'Note: Don't have to check for max attempts here
        'at the end of this routine, attempts are checked

        'date comes from tblLogin

        intMinutes = DateDiff(DateInterval.Minute, dtRes, dtNow)

        dtRes1 = DateAdd(DateInterval.Minute, -CDbl(intMinutesMax), dtNow)
        ts = dtNow - dtRes1
        tsd = dtNow - dtRes

        Dim intAA As Short
        Try
            intAA = rsLogIn.RecordCount
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        If rsLogIn.RecordCount > 0 Then
            'check last entry
            rsLogIn.MoveLast()
            var1 = rsLogIn.Fields("BOOLSUCCESSLOGIN").Value
            If var1 = 0 Then
                boolAttMinutesBad = True
            Else
                boolAttMinutesBad = False
            End If

        Else 'this means that everything is OK
            boolAttMinutesBad = False
        End If

        If boolAttMinutesBad Then
            strM = "StudyDoc login restriction policy states that login attempts are restricted for " & intMinutesMax & " minutes after being locked out of the system."
            strM = strM & ChrW(10) & ChrW(10)
            strM = strM & "User must wait " & intMinutes & " minutes to attempt another login."
            MsgBox(strM, vbInformation, "Invalid action...")
            Me.txtPassword.Clear()
            GoTo end1
        End If

        frmCP.chkFromAdmin.Checked = False
        frmCP.chkFromChgPswd.Checked = True

        tbl = tblUserAccounts
        tblP = tblPersonnel
        tblC = tblConfiguration
        tblPerm = tblPermissions

        varU = Me.txtUserID.Text
        varP = Me.txtPassword.Text

        If StrComp(strAdmin, varU, CompareMethod.Text) = 0 Then
            boolAdmin = True
        End If

        'check to see if LDAP is to be used
        'note that CHARNETWORKACCOUNT must have an entry

        gboolLDAP = UseLDAP(varU.ToString)

        If gboolLDAP Then
        Else
            intAtt = intAtt + 1
        End If

        If gboolLDAP Or boolAdmin Then
        Else
            'check to see if this is a StudyDoc user and if that user is supposed to be using LDAP
            strF = "CHARUSERID = '" & varU & "'"
            Dim rowsA() As DataRow = tblUserAccounts.Select(strF)

            If rowsA.Length = 0 Then
                'keep going
            Else

                Dim boolConfigLDAP As Boolean = False
                str2 = NZ(rowsA(0).Item("CHARLDAP"), "")
                If Len(str2) > 0 Then
                    boolConfigLDAP = True
                End If

                str1 = NZ(rowsA(0).Item("CHARNETWORKACCOUNT"), "")

                Select Case INTWINAUTH
                    Case 1 'LDAP
                        If Len(str1) = 0 Then

                        Else
                            If boolConfigLDAP Then
                                strM = "This StudyDoc user is configured to use LDAP Network Account login."
                                strM = strM & ChrW(10) & ChrW(10)
                                strM = strM & "Please use your Network Account to login."
                                MsgBox(strM, vbInformation, "Invalid action...")
                                Me.txtPassword.Clear()
                                GoTo end1
                            End If
                        End If
                    Case 2, 3 'Non-LDAP, ADVAPI32
                        If Len(str1) = 0 Then
                            'continue
                        Else

                            Select Case INTWINAUTH
                                Case 2
                                    str1 = "Non-LDAP"
                                Case 3
                                    str1 = "ADVAPI32"
                            End Select

                            strM = "This StudyDoc user is configured to use " & str1 & " Network Account login."
                            strM = strM & ChrW(10) & ChrW(10)
                            strM = strM & "Please use your Network Account to login."
                            MsgBox(strM, vbInformation, "Invalid action...")
                            Me.txtPassword.Clear()
                            GoTo end1
                        End If
                End Select

            End If
        End If


        If Len(varU) = 0 Then

            boolGo = False
            strM = "User Name cannot be blank"
            strM1 = "Invalid logon credentials..."

            MsgBox(strM, vbInformation, strM1)

            strF = "charUserID = '-1'"
            rows = tbl.Select(strF)



            dtNow = Now
            '1=UserIDAtt, 2=id_Userid (-1 if none), 3=dtAttempt, 4=boolSucUID (-1 or 0), 5=boolSucPswd (-1 or 0), 6=boolSucLogin (-1 or 0), 7=Comments
            arrAtt(1, intAtt) = varU
            arrAtt(2, intAtt) = -1
            arrAtt(3, intAtt) = dtNow
            arrAtt(4, intAtt) = 0
            arrAtt(5, intAtt) = 0
            arrAtt(6, intAtt) = 0
            arrAtt(7, intAtt) = strM

            boolSaveLogin = True

            GoTo end1

        Else

            If gboolLDAP And boolAdmin = False Then
                strF = "CHARNETWORKACCOUNT = '" & NZ(varU, "") & "'"
            Else
                strF = "charUserID = '" & NZ(varU, "") & "'"
            End If


            Try
                rows = tbl.Select(strF)
            Catch ex As Exception

                strM1 = "There was a problem interpreting the entered User ID."
                'strM = strM1 & ChrW(10) & ChrW(10) & "It is probable that the entry contains the forbidden character 'apostrophe'."
                MsgBox(strM, vbInformation, "Invalid entry...")
                dtNow = Now
                '1=UserIDAtt, 2=id_Userid (-1 if none), 3=dtAttempt, 4=boolSucUID (-1 or 0), 5=boolSucPswd (-1 or 0), 6=boolSucLogin (-1 or 0), 7=Comments
                arrAtt(1, intAtt) = "Problem"
                arrAtt(2, intAtt) = -1
                arrAtt(3, intAtt) = dtNow
                arrAtt(4, intAtt) = 0
                arrAtt(5, intAtt) = 0
                arrAtt(6, intAtt) = 0
                arrAtt(7, intAtt) = strM1

                boolSaveLogin = False ' True

                Me.txtPassword.Clear()

                GoTo end1

            End Try


        End If


        If Len(varP) = 0 Then

            boolGo = False
            strM = "Password cannot be blank"
            strM1 = "Invalid logon credentials..."

            MsgBox(strM, vbInformation, strM1)

            dtNow = Now
            '1=UserIDAtt, 2=id_Userid (-1 if none), 3=dtAttempt, 4=boolSucUID (-1 or 0), 5=boolSucPswd (-1 or 0), 6=boolSucLogin (-1 or 0), 7=Comments
            arrAtt(1, intAtt) = varU
            arrAtt(2, intAtt) = -1
            arrAtt(3, intAtt) = dtNow
            If rows.Length = 0 Then
                arrAtt(4, intAtt) = 0
            Else
                arrAtt(4, intAtt) = -1
            End If
            arrAtt(5, intAtt) = 0
            arrAtt(6, intAtt) = 0
            arrAtt(7, intAtt) = strM

            boolE = True
            GoTo end1

        End If
        boolGo = False

        strM = "Message"
        strM1 = "Message"

        If rows.Length = 0 Then

            strM = "Invalid logon credentials"
            strM1 = "Invalid logon credentials..."
            boolGo = False
            boolE = True

            MsgBox(strM, MsgBoxStyle.Information, strM1)

            Me.txtPassword.Text = ""
            Me.txtPassword.Select()

            dtNow = Now
            '1=UserIDAtt, 2=id_Userid (-1 if none), 3=dtAttempt, 4=boolSucUID (-1 or 0), 5=boolSucPswd (-1 or 0), 6=boolSucLogin (-1 or 0),  7=Comments
            arrAtt(1, intAtt) = varU
            arrAtt(2, intAtt) = -1
            arrAtt(3, intAtt) = dtNow
            arrAtt(4, intAtt) = 0
            arrAtt(5, intAtt) = 0
            arrAtt(6, intAtt) = 0
            arrAtt(7, intAtt) = strM

            GoTo end1

        Else

            If gboolLDAP And boolAdmin = False Then
                tUserID = rows(0).Item("CHARNETWORKACCOUNT")
                gNetAcct = tUserID
            Else
                tUserID = rows(0).Item("CHARUSERID")
                gNetAcct = ""
            End If

            'record idu
            idU = rows(0).Item("id_tblUserAccounts")
            idPerm = rows(0).Item("ID_TBLPERMISSIONS")
            rowPerm = tblPerm.Select("ID_TBLPERMISSIONS = " & idPerm)
            strPerm = rowPerm(0).Item("CHARPERMISSIONSNAME")
            'check to see if userid is associated with an active user
            'check to see if user is active
            idP = rows(0).Item("id_tblPersonnel")
            strF = "id_tblPersonnel = " & idP & " AND boolActive = -1" ' & True
            rowP = tblP.Select(strF)
            If rowP.Length = 0 And boolAdmin = False Then
                'str3 = rowP(0).Item("charFirstName") & NZ(rowP(0).Item("charMiddleName"), "") & rowP(0).Item("charLastName")
                strM = "UserID is associated with inactive user " & str3 & "."
                strM1 = "Inactive User..."
                boolGo = False

                dtNow = Now
                '1=UserIDAtt, 2=id_Userid (-1 if none), 3=dtAttempt, 4=boolSucUID (-1 or 0), 5=boolSucPswd (-1 or 0), 6=boolSucLogin (-1 or 0), 7=Comments
                arrAtt(1, intAtt) = varU
                arrAtt(2, intAtt) = idU
                arrAtt(3, intAtt) = dtNow
                arrAtt(4, intAtt) = -1
                arrAtt(5, intAtt) = -1
                arrAtt(6, intAtt) = 0
                arrAtt(7, intAtt) = strM

                GoTo end1

            Else
                Dim strA As String
                Dim strB As String
                Dim strC As String

                strA = NZ(rowP(0).Item("charFirstName"), "")
                strB = NZ(rowP(0).Item("charMiddleName"), "")
                strC = NZ(rowP(0).Item("charLastName"), "")
                If Len(strB) = 0 Then
                    tUserName = strA & " " & strC
                Else
                    tUserName = strA & " " & strB & " " & strC
                End If

                boolGo = True
            End If

            var1 = rows(0).Item("boolActive")
            If rows(0).Item("boolActive") = 0 And boolAdmin = False Then 'check to see if user id is active; 0 = Inactive

                strM = "Inactive UserID."
                strM1 = "Inactive UserID..."
                boolGo = False
                boolE = True

                MsgBox(strM, MsgBoxStyle.Information, strM1)

                Me.txtPassword.Text = ""
                Me.txtPassword.Select()

                dtNow = Now
                '1=UserIDAtt, 2=id_Userid (-1 if none), 3=dtAttempt, 4=boolSucUID (-1 or 0), 5=boolSucPswd (-1 or 0), 6=boolSucLogin (-1 or 0), 7=Comments
                arrAtt(1, intAtt) = varU
                arrAtt(2, intAtt) = idU
                arrAtt(3, intAtt) = dtNow
                arrAtt(4, intAtt) = -1
                arrAtt(5, intAtt) = -1
                arrAtt(6, intAtt) = 0
                arrAtt(7, intAtt) = strM

                GoTo end2

            Else 'account is active
                boolGo = True
            End If

        End If

        'now check for correct password

        Dim boolAuth As Boolean
        Dim strLDAP As String


        If boolGo Then

            If gboolLDAP And boolAdmin = False Then

                boolCountLogin = True

                strLDAP = NZ(rows(0).Item("CHARLDAP"), "")

                Call GetINTWINAUTH()

                Select Case INTWINAUTH
                    Case 1
                        boolAuth = AuthenticateUserLDAP(strLDAP, tUserID.ToString, varP.ToString)
                    Case 2, 3
                        boolAuth = AuthenticateUser(tUserID.ToString, varP.ToString)
                End Select


                If boolCountLogin Then
                    intAtt = intAtt + 1
                End If

                If boolAuth Then
                Else
                    varPE2 = ""
                End If

            Else
                varPE1 = NZ(rows(0).Item("charPassword"), 0)
                varPE1 = Decode(varPE1, True)
                varPE2 = Coding(varPE1, False)
            End If



            If StrComp(varPE2, varP, CompareMethod.Binary) = 0 Or boolAuth Then

            Else
                boolGo = False
                strM = "Invalid logon credentials"
                strM1 = "Invalid logon credentials..."
                boolE = True

                If gboolLDAP Then

                Else
                    MsgBox(strM, MsgBoxStyle.Information, strM1)
                End If

                Me.txtPassword.Text = ""
                Me.txtPassword.Select()

                dtNow = Now
                '1=UserIDAtt, 2=id_Userid (-1 if none), 3=dtAttempt, 4=boolSucUID (-1 or 0), 5=boolSucPswd (-1 or 0), 6=boolSucLogin (-1 or 0), 7=Comments
                arrAtt(1, intAtt) = varU
                arrAtt(2, intAtt) = idU
                arrAtt(3, intAtt) = dtNow
                arrAtt(4, intAtt) = -1
                arrAtt(5, intAtt) = 0
                arrAtt(6, intAtt) = 0
                arrAtt(7, intAtt) = strM

                GoTo end1

            End If

        End If


        Dim boolPC As Short
        boolPC = 0
        If boolGo Then 'check to ensure user isn't locked out
            boolPC = NZ(rows(0).Item("boolAccountIsLockedOut"), 0)
            If boolPC = -1 And boolAdmin = False Then
                boolGo = False
                strM = "This account is locked out. Please contact your StudyDoc Administrator."
                boolE = True
                MsgBox(strM, MsgBoxStyle.Information, "Account is locked out...")

                dtNow = Now
                '1=UserIDAtt, 2=id_Userid (-1 if none), 3=dtAttempt, 4=boolSucUID (-1 or 0), 5=boolSucPswd (-1 or 0), 6=boolSucLogin (-1 or 0), 7=Comments
                arrAtt(1, intAtt) = varU
                arrAtt(2, intAtt) = idU
                arrAtt(3, intAtt) = dtNow
                arrAtt(4, intAtt) = -1
                arrAtt(5, intAtt) = -1
                arrAtt(6, intAtt) = -1
                arrAtt(7, intAtt) = strM

                'clear id's
                idU = -1
                idP = -1

                boolSaveLogin = False ' True

                GoTo end2
            Else
                boolGo = True
            End If
        End If

        Dim num1 As Double

        'check password items if Windows Authentication = false

        If boolGo Then 'record global variables at this point

            gUserID = rows(0).Item("CHARUSERID")
            gPswd = varP
            'GNETACCT is already assigned
        End If


        If boolGo And gboolLDAP = False Then

            boolPC = NZ(rows(0).Item("boolPasswordNeverExpires"), 0)
            If boolPC = -1 Or boolAdmin Then 'password never expires
                boolGo = True
            Else

                If boolGo Then 'check to see if user's password has expired

                    strF = "charConfigCategory = 'Password Settings' AND charConfigTitle = 'Password expiration period (days)'"
                    Dim rowC() As DataRow
                    Dim dtExp As Date
                    rowC = tblC.Select(strF)
                    num1 = rowC(0).Item("charConfigValue")
                    dtExp = DateAdd(DateInterval.Day, num1, rows(0).Item("dtTimeStamp"))
                    ts = dtExp - dtNow
                    var1 = ts.Days
                    var2 = ts.Hours
                    var3 = ts.Minutes
                    boolPC = rows(0).Item("boolUserCannotChangePassword")
                    int2 = 0 '6=Yes, 7=No

                    If dtExp < dtNow Then 'password has expired
                        If boolPC = -1 Then
                            strM = "Your password expired " & dtExp.ToString & ", but you do not have permission to change your password."
                            MsgBox(strM & Chr(10) & Chr(10) & "Please notify the StudyDoc Administrator if you wish to change your password.", MsgBoxStyle.Information, "Change your password...")
                        Else
                            int2 = MsgBox("Your password expired " & dtExp.ToString & ". You must change your password. Do you wish to change your password now?", MsgBoxStyle.YesNo, "Change your password...")
                        End If
                        boolGo = False
                    ElseIf var2 <= 24 And var1 = 0 Then 'password expires today
                        If boolPC = -1 Then
                            strM = "Your password expires in less than 24 hours on " & dtExp.ToString & ", but you do not have permission to change your password."
                            MsgBox(strM & Chr(10) & Chr(10) & "Please notify the StudyDoc Administrator if you wish to change your password.", MsgBoxStyle.Information, "Change your password...")
                            int2 = 7
                        Else
                            int2 = MsgBox("Your password expires in less then 24 hours on " & dtExp.ToString & ". Do you wish to change your password now?", MsgBoxStyle.YesNo, "Change your password...")
                        End If
                    ElseIf var1 <= 10 Then 'notify user
                        If boolPC = -1 Then
                            strM = "Your password expires on " & dtExp.ToString & ", but you do not have permission to change your password."
                            MsgBox(strM & Chr(10) & Chr(10) & "Please notify the StudyDoc Administrator if you wish to change your password.", MsgBoxStyle.Information, "Change your password...")
                            int2 = 7
                        Else
                            int2 = MsgBox("Your password expires on " & dtExp.ToString & ". Do you wish to change your password now?", MsgBoxStyle.YesNo, "Change your password...")
                        End If
                    Else
                        int2 = 7
                    End If

                    If int2 = 0 Then 'continue
                        boolGo = False

                        dtNow = Now
                        '1=UserIDAtt, 2=id_Userid (-1 if none), 3=dtAttempt, 4=boolSucUID (-1 or 0), 5=boolSucPswd (-1 or 0), 6=boolSucLogin (-1 or 0), 7=Comments
                        arrAtt(1, intAtt) = varU
                        arrAtt(2, intAtt) = idU
                        arrAtt(3, intAtt) = dtNow
                        arrAtt(4, intAtt) = -1
                        arrAtt(5, intAtt) = -1
                        arrAtt(6, intAtt) = -1
                        arrAtt(7, intAtt) = strM

                        GoTo end1

                    ElseIf int2 = 6 Then 'change password

                        frmCP.txtID.Text = rows(0).Item("id_tblUserAccounts")
                        frmCP.boolFromChgPswd = True
                        Dim vud, vun
                        vud = gUserID
                        vun = gUserName
                        gUserID = tUserID
                        gUserName = tUserName

                        Me.Visible = False

                        frmCP.ShowDialog()
                        If frmCP.boolCancel Then
                            boolGo = False
                        Else
                            boolGo = True
                        End If
                        gUserID = vud
                        gUserName = vun
                        Dim boolGo1 As Boolean = False
                        If frmCP.boolCancel Then
                            boolGo1 = False
                        Else
                            boolGo1 = True
                        End If
                        frmCP.Visible = False
                        frmCP.Dispose()
                        If boolGo1 Then
                            boolPWC = True
                            GoTo end1
                        Else
                            Me.Visible = True
                            boolPC = rows(0).Item("boolChangePasswordAtNextLogon")
                            dtNow = Now
                            boolPWC = True
                            If boolPC = -1 Then 'change to false
                                rows(0).BeginEdit()
                                rows(0).Item("boolChangePasswordAtNextLogon") = 0 'False
                                'Do not need to record date of password change
                                'already recorded in Change Password window
                                'rows(0).Item("DTTIMESTAMP") = dtNow
                                rows(0).EndEdit()
                                boolDBSave = True

                            End If

                        End If

                    ElseIf int2 = 7 Then 'don't change password

                    End If

                End If

                If boolGo Then 'check to see if user must change his/herpassword due to checkbox

                    'record global params
                    gUserID = tUserID
                    gUserName = tUserName

                    boolPC = rows(0).Item("boolChangePasswordAtNextLogon")
                    If boolPC = -1 And boolAdmin = False Then 'force user to change password
                        MsgBox("You must change your password at this logon.", MsgBoxStyle.Information, "Change password...")
                        frmCP.txtID.Text = rows(0).Item("id_tblUserAccounts")
                        frmCP.boolFromChgPswd = True

                        Dim vud, vun
                        vud = gUserID
                        vun = gUserName
                        gUserID = tUserID
                        gUserName = tUserName

                        Me.Visible = False

                        frmCP.ShowDialog()

                        If frmCP.boolCancel Then
                            boolGo = False
                        Else
                            boolGo = True
                            rows(0).BeginEdit()
                            rows(0).Item("boolChangePasswordAtNextLogon") = 0 'False
                            rows(0).EndEdit()

                            boolDBSave = True

                        End If
                        frmCP.Visible = False
                        If boolGo Then
                            boolPWC = True

                            'gPswd global variable set in frmChangePassword


                        Else
                            Me.Visible = True
                        End If

                    End If

                End If

            End If

        Else
            boolGo = True

        End If 'boolGo And gboolLDAP = False w

end1:


        If boolGo Then

            Me.Visible = False
            Me.TopMost = False


            dtNow = Now
            '1=UserIDAtt, 2=id_Userid (-1 if none), 3=dtAttempt, 4=boolSucUID (-1 or 0), 5=boolSucPswd (-1 or 0), 6=boolSucLogin (-1 or 0), 7=Comments
            arrAtt(1, intAtt) = varU
            arrAtt(2, intAtt) = idU
            arrAtt(3, intAtt) = dtNow
            arrAtt(4, intAtt) = -1
            arrAtt(5, intAtt) = -1
            arrAtt(6, intAtt) = -1
            If gboolLDAP Then
                strM = "Successful login (Network Account)."
            Else
                strM = "Successful login."
            End If

            If boolPWC Then
                strM = strM & " (Password changed in this login)."
            End If
            arrAtt(7, intAtt) = strM

            gUserID = varU

            boolSaveLogin = True


            'set all permissions 
            Call SetPermissions(True)

            var1 = var1

        Else

            '****
            'check for timeout
            'If intAtt >= intAttMax Then
            If intAtt = intAttMax Then

                strM1 = "The maximum number of allowed logon attempts (" & intAttMax & ") has been surpassed."
                strM = strM1 & ChrW(10) & ChrW(10)
                strM = strM & "Login attempts to StudyDoc will be prohibited for " & intMinutesMax & " minutes on this computer ." & ChrW(10) & ChrW(10)
                strM = strM & "Exiting..."
                MsgBox(strM, vbInformation, "Maxed out logon attempts...")

                gUserID = ""
                gPswd = ""


                boolDBSave = True

                boolCancel = True
                Me.Visible = False

                dtNow = Now
                '1=UserIDAtt, 2=id_Userid (-1 if none), 3=dtAttempt, 4=boolSucUID (-1 or 0), 5=boolSucPswd (-1 or 0), 6=boolSucLogin (-1 or 0), 7=Comments
                '1 - 6 should have been filled out earlier
                'arrAtt(1, intAtt) = varU
                'arrAtt(2, intAtt) = -1
                'arrAtt(3, intAtt) = dtNow
                'arrAtt(4, intAtt) = 0
                'arrAtt(5, intAtt) = 0
                'arrAtt(6, intAtt) = 0
                str1 = arrAtt(7, intAtt)
                arrAtt(7, intAtt) = str1 & " " & strM

                'clear id's
                idU = -1
                idP = -1

                boolSaveLogin = True

                GoTo end2

            End If


        End If

        '****

end2:



        If boolDBSave Then

            If boolGuWuOracle Then
                Try
                    ta_tblUserAccounts.Update(tblUserAccounts)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLUSERACCOUNTS.Merge('ds2005.TBLUSERACCOUNTS, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblUserAccountsAcc.Update(tblUserAccounts)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLUSERACCOUNTS.Merge('ds2005Acc.TBLUSERACCOUNTS, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblUserAccountsSQLServer.Update(tblUserAccounts)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLUSERACCOUNTS.Merge('ds2005Acc.TBLUSERACCOUNTS, True)
                End Try
            End If

        End If


        If boolSaveLogin Then
            'save login attempt
            Try
                Call SaveLoginAttempt(arrAtt, con, intAtt, False, False)
            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try

        Else

            If intAtt = intAttMax - 1 Then
                strM = "Please note that user has one remaining login attempt."
                MsgBox(strM, vbInformation, "Note...")
            End If

        End If




end3:

        frmCP.Dispose()


    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        boolCancel = True
        Me.Visible = False
    End Sub

    Private Sub frmLogon_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        Try
            rsLogIn.Close()
            rsLogIn = Nothing
        Catch ex As Exception

        End Try


        Try
            con.Close()
            con = Nothing
        Catch ex As Exception

        End Try


    End Sub


    Private Sub frmLogon_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call ControlDefaults(Me)

        Dim str1 As String
        Dim str2 As String
        Dim strSQL As String

        str1 = "LABIntegrity StudyDoc" & ChrW(8482) & " Logon..."
        Me.Text = str1

        Dim var1, var2
        boolCancel = True

        'var1 = frmH.cmdChangePassword.Left
        'var2 = frmH.cmdChangePassword.Top

        'If boolFormLoad Then
        'Else
        '    Me.Top = var2
        '    Me.Left = var1 + frmH.cmdChangePassword.Width
        'End If

        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String

        tbl = tblConfiguration
        strF = "charConfigCategory = 'Password Settings' AND charConfigTitle = 'Number of login attempts allowed'"
        rows = tbl.Select(strF)
        intAttMax = CInt(rows(0).Item("charConfigValue"))

        ReDim arrAtt(intATTUB, intAttMax)

        strF = "charConfigCategory = 'Password Settings' AND charConfigTitle = 'Password change restriction (minutes)'"
        rows = tbl.Select(strF)
        intMinutesMax = CInt(rows(0).Item("charConfigValue"))

        Me.cmdOK.Text = "&OK"

        Me.Focus()
        Me.txtUserID.Focus()
        Me.TopMost = True
        'SendKeys.Send("a")
        SendKeys.Send("%(A)")
        'Me.txtUserID.Text = "a"

        Dim constr As String
        If boolGuWuAccess Then
            constr = constrIni
        ElseIf boolGuWuSQLServer Then
            constr = constrIni ' "Provider=SQLOLEDB;" & constrIni
        ElseIf boolGuWuOracle Then
            constr = constrIni
        End If
        Try
            con.Open(constr)
        Catch ex As Exception
            var1 = ex.Message
        End Try

        Try
            dtRes = DateAdd(DateInterval.Minute, -CDbl(intMinutesMax), Now)
        Catch ex As Exception
            var1 = ex.Message
        End Try

        ''console.writeline(CStr(dtRes))

        str1 = "SELECT TBLLOGIN.* FROM TBLLOGIN "

        str2 = "WHERE CHARCOMPUTER = '" & gWorkstation & "' AND DTATTEMPT > " & ReturnDate(dtRes) & ";" ' AND BOOLSUCCESSLOGIN = 0"


        strSQL = str1 & str2
        rsLogIn.CursorLocation = CursorLocationEnum.adUseClient
        Try
            rsLogIn.Open(strSQL, con, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
            MsgBox(ex.Message & ChrW(10) & str2 & ChrW(10) & "rsLogin.Open")
        End Try

        rsLogIn.ActiveConnection = Nothing



    End Sub

End Class