Option Compare Text

Imports System.DirectoryServices
Imports System.DirectoryServices.AccountManagement
Imports System.Web.Configuration

Module modLDAP

    '*****
    'https://msdn.microsoft.com/en-us/library/windows/desktop/aa378184(v=vs.85).aspx
    Private Declare Auto Function LogonUser Lib "advapi32.dll" (ByVal lpszUsername As String, _
                                                            ByVal lpszDomain As String, _
                                                            ByVal lpszPassword As String, _
                                                            ByVal dwLogonType As Integer, _
                                                            ByVal dwLogonProvider As Integer, _
                                                            ByRef phToken As IntPtr) As Integer ' Boolean
    Const LOGON32_LOGON_INTERACTIVE As Long = 2
    Const LOGON32_LOGON_NETWORK As Long = 3
    Const LOGON32_PROVIDER_DEFAULT As Long = 1


    '******

    Function AuthenticateUser(user As String, pass As String) As Boolean

        'https://www.codeproject.com/Questions/471016/How-to-ensure-I-could-connect-to-LDAP-successfully

        AuthenticateUser = False

        Dim strM As String

        boolCountLogin = True

        Cursor.Current = Cursors.WaitCursor

        Dim strDomain As String ' = "gubbsinc.local"


        If INTWINAUTH = 2 Then 'Non-LDAP

            Dim DC As PrincipalContext

            Try

                DC = New PrincipalContext(ContextType.Domain)

            Catch ex As Exception

                boolCountLogin = False

                strM = "Problem authenticating user:"
                strM = strM & ChrW(10) & "PrincipalContext: " & ex.Message
                strM = strM & ChrW(10) & "Please attempt to login in again."
                strM = strM & ChrW(10) & "(Note that this login attempt does not count against the configured number of login attempts.)"
                MsgBox(strM, vbInformation, "Problem...")
                GoTo end1

            End Try

            Try

                'If DC.ValidateCredentials(user, pass, DC.Options.ServerBind) Then
                If DC.ValidateCredentials(user, pass) Then
                    'MsgBox("Yes")
                    AuthenticateUser = True
                Else
                    'MsgBox("No")
                    AuthenticateUser = False
                    strM = "Either UserName or Password is incorrect."
                    MsgBox(strM, vbInformation, "Invalid action...")
                End If

            Catch ex As Exception

                boolCountLogin = False
                strM = "Problem authenticating user:"
                strM = strM & ChrW(10) & "ValidateCredentials: " & ex.Message
                strM = strM & ChrW(10) & "Please attempt to login in again."
                strM = strM & ChrW(10) & "(Note that this login attempt does not count against the configured number of login attempts.)"
                MsgBox(strM, vbInformation, "Problem...")

            End Try

end1:

            Try
                DC.Dispose()
            Catch ex As Exception

            End Try

        ElseIf INTWINAUTH = 3 Then 'ADVAPI32 LogonUser

            'http://stackoverflow.com/questions/8933684/issue-on-verifying-user-login-name-and-password
            'get domain from user logged on to the computer

            My.User.InitializeWithWindowsUser()
            Dim arr1() As String = My.User.Name.Split("\")
            Dim strAAA As String = My.User.Name.ToString
            If arr1.Length = 0 Then
                strM = "Problem obtaining domain from user"
                strM = strM & ChrW(10) & "Please attempt to login in again."
                strM = strM & ChrW(10) & "(Note that this login attempt does not count against the configured number of login attempts.)"
                MsgBox(strM, vbInformation, "Problem...")
                GoTo end1
            End If

            strDomain = arr1(0).ToString


            Dim token As IntPtr
            Dim var1, var2

            Try

                'If the function succeeds, the function returns nonzero.
                'If the function fails, it returns zero. To get extended error information, call GetLastError.

                var2 = LogonUser(user, strDomain, pass, LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, token)

                If var2 = 0 Then
                    strM = "Either UserName or Password is incorrect."
                    MsgBox(strM, vbInformation, "Invalid action...")
                    GoTo end2
                Else
                    AuthenticateUser = True
                End If

            Catch ex As Exception
                boolCountLogin = False
                strM = "Problem authenticating user (ADVAPI32):"
                strM = strM & ChrW(10) & ex.Message
                strM = strM & ChrW(10) & "Please attempt to login in again."
                strM = strM & ChrW(10) & "(Note that this login attempt does not count against the configured number of login attempts.)"
                MsgBox(strM, vbInformation, "Problem...")

            End Try

        Else

            GoTo end1

        End If


end2:

        Cursor.Current = Cursors.Default


    End Function

    Function AuthenticateUserLDAP(path As String, user As String, pass As String) As Boolean

        'path = path of LDAP server
        'user = username
        'pass = password

        Cursor.Current = Cursors.WaitCursor


        Dim var1
        Dim strM As String
        Dim strPath As String = "LDAP://" & path

        strM = "User is configured to use network account." & ChrW(10)
        strM = strM & "LDAP: " & path
        strM = strM & ChrW(10) & "UserName: " & user

        AuthenticateUserLDAP = False

        Try

            Dim de As New DirectoryEntry(strPath, user, pass, AuthenticationTypes.Secure)

            Try
                'run a search using those credentials.  
                'If it returns anything, then you're authenticated
                Dim ds As DirectorySearcher = New DirectorySearcher(de)
                ds.FindOne()

                strM = "Success."

            Catch ex1 As Exception

                strM = strM & ChrW(10) & ChrW(10) & "There was a problem with your login credentials:" & ChrW(10) & ChrW(10)
                strM = strM & ex1.Message
                Cursor.Current = Cursors.Default
                MsgBox(strM, vbInformation, "Invalid entry...")
                var1 = ex1.Message
                var1 = var1

                GoTo end1

                'otherwise, it will crash out so return false
            End Try

        Catch ex As Exception

            strM = strM & ChrW(10) & ChrW(10) & "There was a problem with your login credentials:" & ChrW(10) & ChrW(10)
            strM = strM & ex.Message
            Cursor.Current = Cursors.Default
            MsgBox(strM, vbInformation, "Invalid entry...")
            var1 = ex.Message
            var1 = var1

            GoTo end1

        End Try

        AuthenticateUserLDAP = True

end1:

        Cursor.Current = Cursors.Default

    End Function

    Function GetADUsers(ldapServerName As String, strUID As String, strPwd As String)

        Dim intLB As Short = 6
        Dim v(intLB, 1000)

        'http://stackoverflow.com/questions/28214732/active-directory-propertiestoload-get-all-properties

        ' this sample code reads all users from the Active Directory

        Dim entry As DirectoryEntry = Nothing
        Dim searcher As DirectorySearcher = Nothing
        Dim strM As String
        Dim var1, var2, var3, var4, var5


        ' create a directory entry object with current application context user
        ' we pass username and password as nothing to make it takes the current user credentials
        Try
            'entry = New DirectoryEntry("LDAP:\\" & ldapServerName, Nothing, Nothing, AuthenticationTypes.Secure)
            If Len(strUID) = 0 Or Len(strPwd) = 0 Then
                entry = New DirectoryEntry("LDAP://" & ldapServerName)
            Else
                entry = New DirectoryEntry("LDAP://" & ldapServerName, strUID, strPwd, AuthenticationTypes.Secure)
            End If

        Catch ex As Exception
            strM = "There was a problem with the variable 'entry'."
            strM = strM & ChrW(10) & ChrW(10) & ex.Message
            'MsgBox(strM, vbInformation, "Invalid call...")
            v(intLB, 1) = strM
            ReDim Preserve v(intLB, 1)
            GetADUsers = v
            GoTo end1
        End Try


        ' create a searcher for this directory entry.
        searcher = New DirectorySearcher(entry)

        ' specify the filter
        searcher.Filter = "(&(objectCategory=person)(objectClass=user))"

        '' specify the properties to be loaded
        'searcher.PropertiesToLoad.Add("mail")
        'searcher.PropertiesToLoad.Add("name")
        'searcher.PropertiesToLoad.Add("userPrincipalName")

        searcher.PropertiesToLoad.Add("samaccountname") 'network logon account
        searcher.PropertiesToLoad.Add("givenname") 'first name
        searcher.PropertiesToLoad.Add("sn") 'last name
        searcher.PropertiesToLoad.Add("displayname") 'network logon account
        searcher.PropertiesToLoad.Add("userprincipalname") 'domain name


        ' load results into search result collection
        Dim result As SearchResultCollection
        Try
            result = searcher.FindAll()
        Catch ex As Exception
            strM = "There was a problem with the call 'searcher.FindAll'."
            strM = strM & ChrW(10) & ChrW(10) & ex.Message
            'MsgBox(strM, vbInformation, "Invalid call...")
            v(intLB, 1) = strM
            ReDim Preserve v(intLB, 1)
            GetADUsers = v
            GoTo end1
        End Try

        'Additional information: Unable to cast object of type 'System.Collections.DictionaryEntry' to type 'System.DirectoryServices.DirectoryEntry'.
        'Additional information: Unable to cast object of type 'System.Collections.DictionaryEntry' to type 'System.DirectoryServices.DirectoryEntry'.

        ' loop through the collection


        Dim int1 As Int64 = 0
        For Each res As SearchResult In result
            int1 = int1 + 1
            If int1 > UBound(v, 2) Then
                ReDim Preserve v(intLB, UBound(v, 2) + 1000)
            End If

            ''the code below can be used during testing to return all available properties
            ''Dim prop As System.DirectoryServices.DirectoryEntry
            'Dim prop As System.Collections.DictionaryEntry
            'Console.WriteLine("Start New Properties")
            'For Each prop In res.Properties

            '    var1 = prop.Key
            '    var3 = prop.Value.ToString

            '    'var3 = "System.DirectoryServices.ResultPropertyValueCollection" {String}

            '    'foreach (var val in (property.Value as ResultPropertyValueCollection))
            '    Dim prop1 As System.DirectoryServices.ResultPropertyValueCollection
            '    For Each var3 In prop.Value
            '        var2 = var3.ToString
            '        Console.WriteLine(var1.ToString & ": " & var2.ToString)
            '    Next

            'Next
            'Console.WriteLine("End New Properties")

            ' get the properties
            'res.Properties("givenname").Item(0).ToString()
            'res.Properties("samaccountname").Item(0).ToString()
            'res.Properties("displayname").Item(0).ToString()
            'res.Properties("cn").Item(0).ToString()
            'res.Properties("userprincipalname").Item(0).ToString()
            Try
                var1 = NZ(res.Properties("sn").Item(0).ToString, "NA")
            Catch ex As Exception
                var1 = "NA"
            End Try
            v(1, int1) = var1
            Try
                var1 = NZ(res.Properties("givenname").Item(0).ToString, "NA")
            Catch ex As Exception
                var1 = "NA"
            End Try
            v(2, int1) = var1
            Try
                var1 = NZ(res.Properties("displayname").Item(0).ToString, "NA")
            Catch ex As Exception
                var1 = "NA"
            End Try
            v(3, int1) = var1
            Try
                var1 = NZ(res.Properties("samaccountname").Item(0).ToString, "NA")
            Catch ex As Exception
                var1 = "NA"
            End Try
            v(4, int1) = var1
            Try
                var1 = NZ(res.Properties("userprincipalname").Item(0).ToString, "NA")
            Catch ex As Exception
                var1 = "NA"
            End Try
            v(5, int1) = var1

        Next

        ReDim Preserve v(intLB, int1)

        GetADUsers = v

end1:

        ' release the resources
        Try
            entry.Dispose()
        Catch ex As Exception

        End Try
        Try
            searcher.Dispose()
        Catch ex As Exception

        End Try
        Try
            result.Dispose()
        Catch ex As Exception

        End Try

        'See more at: http://www.visual-basic-tutorials.com/Tutorials/Controls/DirectorySearcher.html#sthash.jH3kKBu4.dpuf

    End Function

    Function UseLDAP(strUID As String) As Boolean

        UseLDAP = False

        'need to set INTWINAUTH befor running this function
        Call GetINTWINAUTH()

        Dim strF As String = "CHARNETWORKACCOUNT = '" & strUID & "'"
        Dim rows() As DataRow = tblUserAccounts.Select(strF)
        Dim strLDAP As String = ""

        If rows.Length = 0 Then
            UseLDAP = False
        Else
            Dim strNetID As String = NZ(rows(0).Item("CHARNETWORKACCOUNT"), "")

            If Len(strNetID) = 0 Then
                GoTo end1
            Else

                If INTWINAUTH = 1 Then 'also need ldap server name
                    strLDAP = NZ(rows(0).Item("CHARLDAP"), "")
                    If Len(strLDAP) = 0 Then
                    Else
                        UseLDAP = True
                    End If
                Else
                    UseLDAP = True
                End If

            End If
        End If

end1:

    End Function


End Module
