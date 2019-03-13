Option Compare Text

Public Class frmHeaderFooter

    Public strActiveC As String

    Private Sub frmHeaderFooter_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim int1 As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim strF As String
        Dim bool As Boolean
        Dim boolA As Short

        strF = "ID_TBLPERMISSIONS = " & id_tblPermissions
        Dim rows() As DataRow
        rows = tblPermissions.Select(strF)

        If rows.Length = 0 And boolRefresh = False Then
            'MsgBox("Guest does not have Edit privileges.", MsgBoxStyle.Information, "No no...")
            Me.cmdEdit.Enabled = False

        Else
            Me.cmdEdit.Enabled = True
        End If
        str1 = "NOTE: If any information is entered in the Header/Footer window, any header/footer information configured in the Report Template will be deleted." & ChrW(10) & ChrW(10)
        str1 = str1 & "The header and footer are divided into quadrants." & ChrW(10)
        str1 = str1 & "The user may configure each quandrant."
        Me.lblTitle.Text = str1


    End Sub

    Private Sub dgvFooter_CellMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs)

    End Sub

    Private Sub dgvHeader_CellMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs)

    End Sub

    Sub DoThis(ByVal cmd As String)
        Cursor.Current = Cursors.WaitCursor
        Dim int1 As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim strF As String
        Dim bool As Boolean
        Dim boolA As Short

        strF = "ID_TBLPERMISSIONS = " & id_tblPermissions
        Dim rows() As DataRow
        rows = tblPermissions.Select(strF)

        If StrComp(cmd, "Logoff", CompareMethod.Text) = 0 Then
        Else
            If rows.Length = 0 And boolRefresh = False Then
                MsgBox("Guest does not have Edit privileges.", MsgBoxStyle.Information, "No no...")
                Exit Sub
            End If
        End If


        Cursor.Current = Cursors.WaitCursor

        Select Case cmd
            Case "Edit"

                Call LockWindow(Not (BOOLHOME))

                Me.cmdEdit.Enabled = False
                Me.cmdSave.Enabled = True
                Me.cmdCancel.Enabled = True
                Me.cmdExit.Enabled = False

            Case "Save"

                Call SaveData()
                Call LockWindow(True)

                Me.cmdEdit.Enabled = True
                Me.cmdSave.Enabled = False
                Me.cmdCancel.Enabled = False
                Me.cmdExit.Enabled = True


            Case "Cancel"

                Call DoCancel(False)
                Call LockWindow(True)


                Me.cmdEdit.Enabled = True
                Me.cmdSave.Enabled = False
                Me.cmdCancel.Enabled = False
                Me.cmdExit.Enabled = True


        End Select

        Cursor.Current = Cursors.Default

    End Sub

    Sub SaveData()

        Dim var1
        Dim str1 As String
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim Count1 As Short
        Dim dt As Date
        Dim strF As String
        Dim strS As String

        dt = Now

        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTS = " & id_tblReports
        strS = "ID_TBLREPORTS ASC"
        dtbl = tblReportHeaders

        rows = dtbl.Select(strF, strS)

        rows(0).BeginEdit()

        For Count1 = 1 To 8
            Select Case Count1
                Case 1
                    str1 = "CHARHLT"
                Case 2
                    str1 = "CHARHRT"
                Case 3
                    str1 = "CHARHLB"
                Case 4
                    str1 = "CHARHRB"
                Case 5
                    str1 = "CHARFLT"
                Case 6
                    str1 = "CHARFRT"
                Case 7
                    str1 = "CHARFLB"
                Case 8
                    str1 = "CHARFRB"
            End Select

            var1 = Me.Controls(str1).Text
            rows(0).Item(str1) = var1

        Next

        'BOOLDIFFFIRSTPAGE
        'BOOLINCLUDELOGO
        str1 = "BOOLDIFFFIRSTPAGE"
        var1 = Me.chkDiffFirstPage.Checked
        If var1 = -1 Then
            rows(0).Item(str1) = -1
        Else
            rows(0).Item(str1) = 0
        End If

        str1 = "BOOLINCLUDELOGO"
        var1 = Me.chkIncludeLogo.Checked
        If var1 = -1 Then
            rows(0).Item(str1) = -1
        Else
            rows(0).Item(str1) = 0
        End If

        rows(0).Item("UPSIZE_TS") = dt

        rows(0).EndEdit()



        If boolGuWuOracle Then
            Try
                ta_tblReportHeaders.Update(tblReportHeaders)
            Catch ex As DBConcurrencyException
                'ds2005.TBLREPORTHEADERS.Merge('ds2005.TBLREPORTHEADERS, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblReportHeadersAcc.Update(tblReportHeaders)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLREPORTHEADERS.Merge('ds2005Acc.TBLREPORTHEADERS, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblReportHeadersSQLServer.Update(tblReportHeaders)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLREPORTHEADERS.Merge('ds2005Acc.TBLREPORTHEADERS, True)
            End Try
        End If




    End Sub

    Sub LockWindow(ByVal bool As Boolean)
        Dim Count1 As Short
        Dim str1 As String
        Me.CHARHLT.ReadOnly = bool
        Me.CHARHRT.ReadOnly = bool
        Me.CHARHLB.ReadOnly = bool
        Me.CHARHRB.ReadOnly = bool
        Me.CHARFLT.ReadOnly = bool
        Me.CHARFRT.ReadOnly = bool
        Me.CHARFLB.ReadOnly = bool
        Me.CHARFRB.ReadOnly = bool

        Me.chkDiffFirstPage.Enabled = Not (bool)
        Me.chkIncludeLogo.Enabled = Not (bool)

    End Sub

    Sub DoCancel(ByVal bool As Boolean)

        Call LoadData()


    End Sub

    Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Call DoThis("Cancel")

    End Sub

    Private Sub cmdEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Call DoThis("Edit")

    End Sub

    Private Sub cmdSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call DoThis("Save")

    End Sub

    Private Sub cmdExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub


    Sub LoadData()
        Dim h

        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim strS As String
        Dim strFM As String
        Dim strSM As String
        Dim str1 As String
        Dim var1, var2
        Dim Count1 As Short

        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTS = " & id_tblReports
        strS = "ID_TBLREPORTS ASC"
        dtbl = tblReportHeaders

        rows = dtbl.Select(strF, strS)

        If rows.Length = 0 Then 'add a new record
            Call FillHeaderFooterTable()
        End If

        'load DATA

        For Count1 = 1 To 8
            Select Case Count1
                Case 1
                    str1 = "CHARHLT"
                Case 2
                    str1 = "CHARHRT"
                Case 3
                    str1 = "CHARHLB"
                Case 4
                    str1 = "CHARHRB"
                Case 5
                    str1 = "CHARFLT"
                Case 6
                    str1 = "CHARFRT"
                Case 7
                    str1 = "CHARFLB"
                Case 8
                    str1 = "CHARFRB"
            End Select

            'var1 = rows(0).Item(str1)
            If rows.Length = 0 Then
                var1 = ""
            Else
                var1 = NZ(rows(0).Item(str1), "")
            End If
            Me.Controls(str1).Text = var1

        Next

        'BOOLDIFFFIRSTPAGE
        'BOOLINCLUDELOGO


        If rows.Length = 0 Then
            Me.chkDiffFirstPage.Checked = False
        Else
            str1 = "BOOLDIFFFIRSTPAGE"
            var1 = rows(0).Item(str1)
            If var1 = -1 Then
                Me.chkDiffFirstPage.Checked = True
            Else
                Me.chkDiffFirstPage.Checked = False
            End If

        End If

        If rows.Length = 0 Then
            Me.chkIncludeLogo.Checked = False
        Else
            str1 = "BOOLINCLUDELOGO"
            var1 = rows(0).Item(str1)
            If var1 = -1 Then
                Me.chkIncludeLogo.Checked = True
            Else
                Me.chkIncludeLogo.Checked = False
            End If
        End If

        'ID_TBLREPORTHEADERS
        'ID_TBLREPORTS
        'ID_TBLSTUDIES
        'BOOLDIFFFIRSTPAGE
        'BOOLINCLUDELOGO
        'CHARHLT 5
        'CHARHRT 6
        'CHARHLB 7
        'CHARHRB 8
        'CHARFLT 9
        'CHARFRT 10
        'CHARFLB 11
        'CHARFRB
        'UPSIZE_TS



        Call LockWindow(True)

    End Sub

    Private Sub cmiFieldCodes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmiFieldCodes.Click
        If Me.cmdEdit.Enabled Then
            MsgBox("Please go into Edit mode", MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        Call GetFieldCodes()
    End Sub

    Sub GetFieldCodes()
        Dim pos As Short
        Dim strT As String
        Dim str1 As String
        Dim strL As String
        Dim strR As String
        Dim strC As String
        Dim intL As Short
        Dim var1, var2

        Dim tbx As TextBox

        tbx = Me.Controls(strActiveC)

        If tbx Is Nothing Then
            Exit Sub
        End If


        'record position of cursor in text box
        'pos = tbx.SelectionStart + tbx.SelectionLength

        Try
            pos = tbx.SelectionStart + tbx.SelectionLength
        Catch ex As Exception
            pos = 0
        End Try

        tbx.SelectionStart = pos
        tbx.SelectionLength = 0

        Dim frm As New frmFieldCodes
        'Dim t, l, w, h, t1, l1

        't = tbx.Top
        'l = tbx.Left
        'w = tbx.Width
        'h = tbx.Height

        't1 = t + h + Me.Top + 10
        'l1 = l + (w / 2)

        Me.Cursor = New Cursor(Cursor.Current.Handle)

        'frm.Location = new system.drawing.point(l1, t1)

        frm.Location = New System.Drawing.Point(Cursor.Position.X, Cursor.Position.Y + 10)

        frm.ShowDialog()

        If frm.boolCancel Then

            'tbx.SelectionLength = 0
            'tbx.SelectionStart = pos

        Else

            strT = tbx.Text
            intL = Len(strT)
            If pos = 0 Then
                strL = "" 'Mid(strT, 1, pos)
                strR = strT 'Mid(strT, pos + 1, Len(strT) - pos)
                str1 = frm.strFC & " " & strR
            ElseIf pos = intL Then
                strL = strT 'Mid(strT, 1, pos)
                strR = "" 'Mid(strT, pos + 1, Len(strT) - pos)
                var1 = Mid(strT, intL, 1)
                If Asc(var1) = 32 Then 'space
                    str1 = strL & frm.strFC
                Else
                    str1 = strL & " " & frm.strFC
                End If
            Else
                strL = Mid(strT, 1, pos - 1)
                strR = Mid(strT, pos + 1, Len(strT) - pos)
                str1 = strL & " " & frm.strFC & " " & strR
            End If
            tbx.Text = str1
        End If

        tbx.SelectionStart = pos
        tbx.SelectionLength = 0

        frm.Dispose()
    End Sub

    Private Sub CHARFLB_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CHARFLB.Click
        strActiveC = "CHARFLB"
    End Sub

    Private Sub CHARFLT_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CHARFLT.Click
        strActiveC = "CHARFLT"

    End Sub

    Private Sub CHARFRB_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CHARFRB.Click
        strActiveC = "CHARFRB"

    End Sub

    Private Sub CHARFRT_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CHARFRT.Click
        strActiveC = "CHARFRT"

    End Sub

    Private Sub CHARHLB_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CHARHLB.Click
        strActiveC = "CHARHLB"

    End Sub

    Private Sub CHARHLT_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CHARHLT.Click
        strActiveC = "CHARHLT"

    End Sub

    Private Sub CHARHRB_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CHARHRB.Click
        strActiveC = "CHARHRB"

    End Sub

    Private Sub CHARHRT_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CHARHRT.Click
        strActiveC = "CHARHRT"

    End Sub
End Class