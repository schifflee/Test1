Option Compare Text

Public Class frmAddUser

    Public boolCancel As Boolean = True
    Public strForm As String = ""
    Public idP As Int64


    Private Sub frmAddUser_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Me.panUserName.Top = Me.panUserID.Top
        Me.panUserName.Left = Me.panUserID.Left

        Dim a

        a = Me.cmdOK.Top + Me.cmdOK.Height + 50
        Me.Height = a

    End Sub


    Sub FormLoad()

        Select Case strForm
            Case "UserID"
                Me.panUserID.Visible = True
                Me.panUserName.Visible = False
            Case "UserName"
                Me.panUserName.Visible = True
                Me.panUserID.Visible = False
        End Select

        Dim a, b, c, d

        a = Me.cmdOK.Top + Me.cmdOK.Height + 50
        Me.Height = a

    End Sub

    Private Sub cmdOK_Click(sender As System.Object, e As System.EventArgs) Handles cmdOK.Click

        If ValidateEntries() Then
        Else
            Exit Sub
        End If

        boolCancel = False
        Me.Visible = False

    End Sub

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click

        boolCancel = True
        Me.Visible = False

    End Sub

    Function ValidateEntries() As Boolean

        ValidateEntries = False

        Dim strM As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strUN As String
        Dim strF As String
        Dim strF1 As String
        Dim rowC() As DataRow
        Dim var1, var2, var3, var4
        Dim varP2
        Dim strM1 As String
        Dim strP As String
        Dim Count1 As Int32
        Dim strUID As String
        Dim id As Int64

        Select Case strForm
            Case "UserID"
                'check UserID
                strUID = Me.txtUserID.Text
                If Len(strUID) = 0 Then
                    strM = "UserID cannot be blank."
                    GoTo end1
                End If

                'make sure entry is unique for this userid
                Dim rows() As DataRow

                strF = "ID_TBLPERSONNEL = " & idP
                rows = tblUserAccounts.Select(strF)
                'don't filter
                'don't allow repeats throughout entire table
                For Count1 = 0 To tblUserAccounts.Rows.Count - 1
                    str1 = tblUserAccounts.Rows(Count1).Item("CHARUSERID")
                    If StrComp(str1, strUID, CompareMethod.Text) = 0 Then
                        id = tblUserAccounts.Rows(Count1).Item("ID_TBLPERSONNEL")
                        strF1 = "ID_TBLPERSONNEL = " & id
                        Dim rowsA() As DataRow = tblPersonnel.Select(strF1)
                        str1 = rowsA(0).Item("CHARFIRSTNAME")
                        str2 = NZ(rowsA(0).Item("CHARMIDDLENAME"), "")
                        str3 = rowsA(0).Item("CHARLASTNAME")
                        If Len(NZ(str2, "")) = 0 Then
                            str4 = str1 & " " & str3
                        Else
                            str4 = str1 & " " & str2 & " " & str3
                        End If
                        strM = "User ID '" & strUID & "' already exists associated with the User Name '" & str4 & "."
                        GoTo end1
                    End If
                Next

                'check Password
                str1 = Me.txtPswd.Text
                str2 = Me.txtConfirm.Text

                If StrComp(str1, str2, CompareMethod.Text) = 0 Then
                Else
                    strM = "Passwords do not match."
                    GoTo end1
                End If

                If Len(str1) = 0 Or Len(str2) = 0 Then
                    strM = "Password cannot be blank."
                    GoTo end1
                End If

                Dim ctI As Short
                Dim tblC As DataTable
                tblC = tblConfiguration
                Dim int1 As Int32

                varP2 = str1
                strP = str1

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

                        GoTo end1
                    End If
                End If

            Case "UserName"
                'check txtFirstName
                str1 = Me.txtFirstName.Text
                If Len(str1) = 0 Then
                    strM = "First Name cannot be blank."
                    GoTo end1
                End If

                str2 = Me.txtMiddleName.Text

                'check txtLastName
                str3 = Me.txtLastName.Text
                If Len(str3) = 0 Then
                    strM = "Last Name cannot be blank."
                    GoTo end1
                End If

                'check for unique entry
                strUN = str1 & str2 & str3
                For Count1 = 0 To tblPersonnel.Rows.Count - 1
                    str1 = NZ(tblPersonnel.Rows(Count1).Item("CHARFIRSTNAME"), "")
                    str2 = NZ(tblPersonnel.Rows(Count1).Item("CHARMIDDLENAME"), "")
                    str3 = NZ(tblPersonnel.Rows(Count1).Item("CHARLASTNAME"), "")

                    str4 = str1 & str2 & str3

                    If StrComp(strUN, str4, CompareMethod.Text) = 0 Then
                        strM = "The User Name:" & ChrW(10) & ChrW(10) & strUN & ChrW(10) & ChrW(10) & "already exists."
                        GoTo end1
                    End If

                Next

        End Select

        ValidateEntries = True

end1:
        If ValidateEntries Then
        Else
            MsgBox(strM, vbInformation, "Invalid entry...")
        End If

    End Function

End Class