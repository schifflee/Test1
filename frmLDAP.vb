Public Class frmLDAP

    Public strType As String
    Public boolCancel As String = True
    Public gUserID As String
    Public gPswd As String


    Private Sub frmLDAP_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call DoubleBufferControl(Me, "dgv")

        Call ControlDefaults(Me)

    End Sub

    Sub FormLoad()



        Call SizeControls()

        Call ShowPanels()

        Call LoadUsers()

    End Sub

    Sub ShowPanels()

        Dim pan As Panel

        Select Case strType

            Case "Existing"
                Me.panExisting.Visible = True
                Me.panTest.Visible = False
                Me.AcceptButton = Me.cmdRetrieve
            Case "Test"
                Me.panExisting.Visible = False
                Me.panTest.Visible = True

        End Select


    End Sub

    Sub SizeControls()


        Dim ns As New Size

        'Me.panTest.Size = Me.panExisting.Size ' New Size(0, 1)

        Dim pan As Panel = Me.panTest

        pan.Top = 20
        pan.Left = 20

        Dim a, b, c, d

        a = pan.Left + pan.Width

        Dim bw As Int16 = (Me.Width - Me.ClientSize.Width) / 2 'form border width
        Dim tbh As Int16 = Me.Height - Me.ClientSize.Height - 2 * bw 'titlebar height

        Me.Width = a + 20 + bw

        b = pan.Top + pan.Height

        Me.Height = b + 20 + tbh

        Me.panExisting.Left = pan.Left
        Me.panExisting.Top = pan.Top

        pan.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom

        Me.panExisting.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom


    End Sub

    Sub LoadUsers()

        '"LDAP://gubbs11.gubbsinc.local"

        Dim strU As String = Me.txtUserID.Text
        Dim strPwd As String = Me.txtPswd.Text
        Dim strLDAP As String = Me.txtLDAP.Text
        Dim var1
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String

        Cursor.Current = Cursors.WaitCursor

        Dim v = GetADUsers(strLDAP, strU, strPwd)
        Dim intLB = UBound(v, 1)

        'check for error
        var1 = NZ(v(intLB, 1), "")
        If Len(var1) = 0 Then
            Me.txtStatus.Text = "Success."
        Else
            Me.txtStatus.Text = var1
            GoTo end1
        End If

        ''legend:
        'searcher.PropertiesToLoad.Add("samaccountname") 'network logon account
        'searcher.PropertiesToLoad.Add("givenname") 'first name
        'searcher.PropertiesToLoad.Add("sn") 'last name
        'searcher.PropertiesToLoad.Add("displayname") 'network logon account
        'searcher.PropertiesToLoad.Add("userprincipalname") 'domain name

        Dim Count1 As Short
        Dim Count2 As Int32
        Dim intUB As Int64 = UBound(v, 2)

        'record data in a datatable
        Dim dtbl As New System.Data.DataTable
        For Count1 = 1 To 5

            Dim col1 As New DataColumn

            Select Case Count1
                Case 1
                    str1 = "sn"
                    str2 = "Last Name"
                Case 2
                    str1 = "givenname"
                    str2 = "First Name"
                Case 3
                    str1 = "displayname"
                    str2 = "Full Name"
                Case 4
                    str1 = "samaccountname"
                    str2 = "Network Account ID"
                Case 5
                    str1 = "userprincipalname"
                    str2 = "Full Network Account ID"
            End Select

            col1.ColumnName = str1
            col1.Caption = str2
            col1.DataType = System.Type.GetType("System.String")
            dtbl.Columns.Add(col1)

        Next

        'now enter data
        For Count2 = 1 To intUB

            Dim nr As System.Data.DataRow = dtbl.NewRow
            nr.BeginEdit()
            For Count1 = 1 To 5

                str1 = v(Count1, Count2)
                nr.Item(Count1 - 1) = str1

            Next
            nr.EndEdit()
            dtbl.Rows.Add(nr)

        Next

        Dim dgv As DataGridView = Me.dgvUsers
        Dim dv As DataView = New DataView(dtbl, "", "sn ASC", DataViewRowState.CurrentRows)

        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False
        dgv.DataSource = dv

        For Count1 = 1 To 5

            Select Case Count1
                Case 1
                    str1 = "sn"
                    str2 = "Last Name"
                Case 2
                    str1 = "givenname"
                    str2 = "First Name"
                Case 3
                    str1 = "displayname"
                    str2 = "Full Name"
                Case 4
                    str1 = "samaccountname"
                    str2 = "Network Account ID"
                Case 5
                    str1 = "userprincipalname"
                    str2 = "Full Network Account ID"
            End Select

            dgv.Columns(Count1 - 1).HeaderText = str2
            dgv.Columns(Count1 - 1).SortMode = DataGridViewColumnSortMode.NotSortable

        Next

        dgv.ColumnHeadersDefaultCellStyle.Font = New Font(dgv.Font, FontStyle.Bold)

        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells


end1:

        Cursor.Current = Cursors.Default

    End Sub

    Private Sub cmdTest_Click(sender As Object, e As EventArgs) Handles cmdRetrieve.Click

        Call LoadUsers()


    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click

        boolCancel = True
        Me.Visible = False

    End Sub

    Private Sub chkShowPswd_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowPswd.CheckedChanged

        If Me.chkShowPswd.Checked Then
            Me.txtPswd.PasswordChar = ""
        Else
            Me.txtPswd.PasswordChar = "*"
        End If

    End Sub

    Function ValidateNetAcct() As Boolean

        ValidateNetAcct = False

        Dim dgv As DataGridView = Me.dgvUsers

        Try
            Dim intRow As Int32 = dgv.CurrentRow.Index
            Dim strNet As String = dgv("samaccountname", intRow).Value

            'network account can be assigned to only one StudyDoc account
            Dim strF As String = "CHARNETWORKACCOUNT = '" & strNet & "'"
            Dim rows() As DataRow = tblUserAccounts.Select(strF)
            If rows.Length = 0 Then
            Else

                Dim strUID As String = rows(0).Item("CHARUSERID")
                Dim intP As Int64 = rows(0).Item("ID_TBLPERSONNEL")
                Dim strF1 As String = "ID_TBLPERSONNEL = " & intP
                Dim rowsP() As DataRow = tblPersonnel.Select(strF1)
                Dim strFN As String = rowsP(0).Item("CHARFIRSTNAME")
                Dim strMN As String = rowsP(0).Item("CHARMIDDLENAME")
                Dim strLN As String = rowsP(0).Item("CHARLASTNAME")
                Dim str1 As String

                If Len(strMN) = 0 Then
                    str1 = "The network account '" & strNet & "' has already been assigned to:" & ChrW(10) & ChrW(10) & "        " & strFN & " " & strLN & "."
                Else
                    str1 = "The network account '" & strNet & "' has already been assigned to:" & ChrW(10) & ChrW(10) & "        " & strFN & " " & strMN & " " & strLN & "."
                End If

                str1 = str1 & ChrW(10) & ChrW(10) & "A network account can be assigned to only one StudyDoc account."

                MsgBox(str1, vbInformation, "Invalid choice...")
                GoTo end1

            End If

            ValidateNetAcct = True

        Catch ex As Exception

        End Try


end1:

    End Function

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click

        If ValidateNetAcct() Then

            boolCancel = False
            Me.Visible = False

        End If

    End Sub

    Private Sub cmdOK1_Click(sender As Object, e As EventArgs) Handles cmdOK1.Click

        boolCancel = False
        Me.Visible = False

    End Sub

    Private Sub cmdExit1_Click(sender As Object, e As EventArgs) Handles cmdExit1.Click

        boolCancel = True
        Me.Visible = False

    End Sub

 
    Private Sub txtFilter_TextChanged(sender As Object, e As EventArgs) Handles txtFilter.TextChanged

        Dim dgv As DataGridView = Me.dgvUsers
        Dim dv As DataView = dgv.DataSource
        Dim strF As String
        Dim str1 As String = Me.txtFilter.Text

        If Len(str1) = 0 Then
            dv.RowFilter = Nothing
        Else
            strF = "sn LIKE '" & str1 & "*'"
            'dv.RowFilter = strF
            Try
                dv.RowFilter = strF
            Catch ex As Exception

            End Try

        End If



    End Sub

End Class