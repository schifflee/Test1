Option Compare Text

Public Class frmApplyDataToGroups

    Public boolCancel As Boolean = True
    Public boolFormLoad As Boolean


    Private Sub frmApplyDataToGroups_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim str1 As String

        str1 = "Select Route to" & ChrW(10) & "copy data FROM -->"
        Me.lblFrom.Text = str1

        str1 = "This action will copy Group/Route Details and Time Points to the selected Routes"
        Me.lblTitle.Text = str1


    End Sub

    Sub FormLoad()

        boolFormLoad = True

        Dim strF As String
        Dim strS As String
        Dim Count1 As Short
        Dim Count2 As Short
        Dim dgv As DataGridView
        Dim tbl As System.Data.DataTable
        Dim dv As System.Data.DataView

        Dim dv1 As New DataView ' = New DataView(frmSD.dgvGroupSummary.DataSource)
        Dim dv2 As New DataView ' = New DataView(frmSD.dgvGroupSummary.DataSource)
        Dim dv3 As New DataView ' = New DataView(frmSD.dgvGroupSummary.DataSource)

        tbl = frmSD.tblGroupSummary.Copy

        dv1 = New DataView(tbl)
        dv2 = New DataView(tbl)
        dv3 = New DataView(tbl)

        dv1.AllowDelete = False
        dv1.AllowNew = False

        Me.dgvGroupSummary.DataSource = dv1

        strF = "ID_TBLGUWUPKGROUPS = -1"
        strS = "ID_TBLGUWUPKGROUPS ASC, ColumnValue ASC"
        dv2.RowFilter = strF
        dv2.Sort = strS
        dv2.AllowDelete = False
        dv2.AllowNew = False
        Me.dgvFrom.DataSource = dv2

        dv3.RowFilter = strF
        dv3.Sort = strS
        dv3.AllowDelete = False
        dv3.AllowNew = False
        Me.dgvTo.DataSource = dv3

        dgv = Me.dgvFrom
        For Count1 = 1 To 3

            Select Case Count1
                Case 1
                    dgv = Me.dgvGroupSummary
                Case 2
                    dgv = Me.dgvTo
                Case 3
                    dgv = Me.dgvFrom
            End Select

            For Count2 = 0 To dgv.Columns.Count - 1
                dgv.Columns(Count2).Visible = False
            Next

            dgv.Columns("ColumnValue").Visible = True

            dgv.Columns("ColumnValue").HeaderText = "Item"

            'dgv.RowHeadersVisible = False
            dgv.RowHeadersWidth = 15
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            dgv.AllowUserToResizeColumns = True
            dgv.AllowUserToResizeRows = True
            dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

            Dim mw As Short
            Dim intDo As Short

            mw = 150

            intDo = 1

            If intDo = 1 Then
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

                'dgv.Columns("ColumnValue").MinimumWidth = mw
                'dgv.Columns.Item("ColumnValue").MinimumWidth = dgv.Width * 0.8
                Try
                    dgv.Columns.Item("ColummValue").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                Catch ex As Exception

                End Try
            Else
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                dgv.Columns("ColumnValue").MinimumWidth = mw
                Try
                    dgv.Columns.Item("ColummValue").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                Catch ex As Exception

                End Try
            End If

        Next

        boolFormLoad = False

        Call UpdateSummarySelection()


    End Sub


    Private Sub cmdOK1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK1.Click

        Dim dgv1 As DataGridView
        Dim strM As String


        dgv1 = Me.dgvFrom

        If dgv1.Rows.Count = 0 Then
            strM = "A Route must be chosen in the 'FROM' table."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            GoTo end1
        End If

        dgv1 = Me.dgvTo

        If dgv1.Rows.Count = 0 Then
            strM = "No items have been assigned as 'TO'."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            GoTo end1
        End If

        boolCancel = False
        Me.Visible = False

end1:

    End Sub

    Private Sub cmdCancel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel1.Click

        boolCancel = True
        Me.Visible = False

    End Sub

    Private Sub dgvGroupSummary_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvGroupSummary.SelectionChanged

        Call UpdateSummarySelection()


    End Sub

    Sub UpdateSummarySelection()
        If boolFormLoad Then
            Exit Sub
        End If

        Dim dgv3 As DataGridView
        Dim intRow3 As Short
        Dim id1 As Int64
        Dim id2 As Int64
        Dim boolF As Boolean

        dgv3 = Me.dgvGroupSummary

        If dgv3.Rows.Count = 0 Then
            GoTo end1
        ElseIf dgv3.CurrentRow Is Nothing Then
            intRow3 = 0
        Else
            intRow3 = dgv3.CurrentRow.Index
        End If

        'intRow3 = dgv3.CurrentRow.Index
        id1 = dgv3("ID_TBLGUWUPKGROUPS", intRow3).Value
        id2 = dgv3("ID_TBLGUWUPKROUTES", intRow3).Value

        If id1 = -1 Or id2 = -1 Then
            boolF = True
            'select either one above or one below
            Do Until id2 <> -1
                If intRow3 = 0 Then
                    intRow3 = intRow3 + 1
                    boolF = True
                ElseIf intRow3 = dgv3.Rows.Count - 1 Then
                    intRow3 = intRow3 - 1
                    boolF = False
                Else
                    If boolF Then
                        intRow3 = intRow3 + 1
                    Else
                        intRow3 = intRow3 - 1
                    End If
                End If
                'id2 = dgv3("ID_TBLGUWUPKROUTES", intRow3).Value
                Try
                    id2 = dgv3("ID_TBLGUWUPKROUTES", intRow3).Value
                Catch ex As Exception
                    Exit Do
                End Try
            Loop


            Try
                dgv3.CurrentCell = dgv3.Rows(intRow3).Cells("ColumnValue")
                dgv3.CurrentRow.Selected = True
            Catch ex As Exception

            End Try

            GoTo end1
        End If

end1:

    End Sub


    Private Sub cmdRemoveTo_Click(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub AddFrom()

        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim intRow1 As Short
        Dim strM As String
        Dim idG As Int64
        Dim idR As Int64
        Dim strf As String
        Dim dv As System.Data.DataView

        dgv1 = Me.dgvGroupSummary
        dgv2 = Me.dgvFrom

        If dgv1.Rows.Count = 0 Then
            Exit Sub
        End If

        If dgv2.Rows.Count > 0 Then
            strM = "Only one Route may be included in the FROM table"
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        If dgv1.CurrentRow Is Nothing Then
            strM = "Please select a route."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        intRow1 = dgv1.CurrentRow.Index
        idG = dgv1("ID_TBLGUWUPKGROUPS", intRow1).Value
        idR = dgv1("ID_TBLGUWUPKROUTES", intRow1).Value

        dv = dgv2.DataSource

        strf = "ID_TBLGUWUPKGROUPS = " & idG & " AND (ID_TBLGUWUPKROUTES = " & idR & " OR ID_TBLGUWUPKROUTES = -1)"
        dv.RowFilter = strf

        Dim strS As String
        strS = "ID_TBLGUWUPKGROUPS ASC" ', ColumnValue ASC"
        dv.Sort = strS

    End Sub


    Sub AddTo(ByVal boolAll As Boolean)

        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim dgv3 As DataGridView
        Dim intRows As Short
        Dim intRow1 As Short
        Dim strM As String
        Dim idG As Int64
        Dim idGN As Int64
        Dim idGY(10) As Int64
        Dim idR As Int64
        Dim idRN As Int64
        Dim idRY(10) As Int64
        Dim arrG(10) As Int64
        Dim Count1 As Short

        Dim strf As String
        Dim dv1 As System.Data.DataView
        Dim dv2 As System.Data.DataView
        Dim dv3 As System.Data.DataView

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String

        Dim int1 As Short
        Dim id As Int64
        Dim id1 As Int64
        Dim id2 As Int64

        Dim intGRows As Short

        dgv1 = Me.dgvGroupSummary
        dgv2 = Me.dgvTo
        dgv3 = Me.dgvFrom

        intRows = dgv1.Rows.Count

        ReDim idGY(intRows)
        ReDim idRY(intRows)
        ReDim arrG(intRows)

        If dgv1.Rows.Count = 0 Then
            Exit Sub
        End If

        If dgv1.CurrentRow Is Nothing Then
            strM = "Please select one or more routes."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        If dgv3.Rows.Count = 0 Then
            strM = "A Route must first be added to the FROM table"
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        'get exception from dgv3
        For Count1 = 0 To dgv3.Rows.Count - 1
            idRN = dgv3("ID_TBLGUWUPKROUTES", Count1).Value
            idGN = dgv3("ID_TBLGUWUPKGROUPS", Count1).Value
            If idRN = -1 Then
            Else
                Exit For
            End If
        Next

        int1 = 0

        If boolAll Then
            dgv1.SelectAll()
        End If

        For Each SelectedRow As DataGridViewRow In dgv1.SelectedRows
            int1 = int1 + 1
            idGY(int1) = SelectedRow.Cells("ID_TBLGUWUPKGROUPS").Value
            idRY(int1) = SelectedRow.Cells("ID_TBLGUWUPKROUTES").Value

            If idGY(int1) = idGN And idRY(int1) = idRN Then
                'build str1
                str3 = ""
                For Count1 = 0 To dgv3.Rows.Count - 1
                    str2 = dgv3("ColumnValue", Count1).Value
                    If Count1 = 0 Then
                        str3 = str2
                    Else
                        str3 = str3 & " - " & Trim(str2)
                    End If
                Next
                strM = "Please note that the selected Route '" & str3 & "' cannot be added because it is the FROM Route."
                MsgBox(strM, MsgBoxStyle.Information, "Please note...")
            End If
        Next

        intRows = int1

        'find number of intgrows
        int1 = 0
        intGRows = 0
        For Count1 = 1 To intRows
            id = idRY(Count1)
            If id = -1 Then
                intGRows = intGRows + 1
                arrG(intGRows) = idGY(intRows)
            End If
        Next

        'build strF
        int1 = 0
        'strf = "(a = 1 and (b = 1 or b = 2)) or (a = 2 and (b = 1 or b = 2)) not (a = 2 and b = 2)"
        For Count1 = 1 To intRows
            id1 = idGY(Count1)
            id2 = idRY(Count1)
            If Count1 = 1 Then
                strf = "(ID_TBLGUWUPKGROUPS = " & id1 & " AND (ID_TBLGUWUPKROUTES = " & id2 & " OR ID_TBLGUWUPKROUTES = -1))"
            Else
                strf = strf & " OR (ID_TBLGUWUPKGROUPS = " & id1 & " AND (ID_TBLGUWUPKROUTES = " & id2 & " OR ID_TBLGUWUPKROUTES = -1))"
            End If
        Next
        strf = strf & "AND NOT (ID_TBLGUWUPKGROUPS = " & idGN & " AND ID_TBLGUWUPKROUTES = " & idRN & ")"

        dv2 = dgv2.DataSource

        dv2.RowFilter = strf

        Dim strS As String
        strS = "ID_TBLGUWUPKGROUPS ASC" ', ColumnValue ASC"
        dv2.Sort = strS

    End Sub

    Sub RemoveTo()

        Dim dgv2 As DataGridView
        Dim strf As String
        Dim dv As System.Data.DataView

        dgv2 = Me.dgvTo

        If dgv2.Rows.Count = 0 Then
            Exit Sub
        End If

        strf = "ID_TBLGUWUPKGROUPS = -1"

        dv = dgv2.DataSource
        dv.RowFilter = strf

    End Sub

    Private Sub RemoveFrom()

        Dim dgv2 As DataGridView
        Dim strf As String
        Dim dv As System.Data.DataView

        dgv2 = Me.dgvFrom

        If dgv2.Rows.Count = 0 Then
            Exit Sub
        End If

        strf = "ID_TBLGUWUPKGROUPS = -1"

        dv = dgv2.DataSource
        dv.RowFilter = strf

        dgv2 = Me.dgvTo

        If dgv2.Rows.Count = 0 Then
            Exit Sub
        End If

        strf = "ID_TBLGUWUPKGROUPS = -1"

        dv = dgv2.DataSource
        dv.RowFilter = strf

    End Sub

    Private Sub cmdAddFrom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddFrom.Click

        Call AddFrom()

    End Sub

    Private Sub cmdRemoveFrom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemoveFrom.Click

        Call RemoveFrom()

    End Sub



    Private Sub cmdRemoveTo_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemoveTo.Click

        Call RemoveTo()

    End Sub

    Private Sub cmdAddAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddAll.Click

        Call AddTo(True)

    End Sub

    Private Sub cmdAddTo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddTo.Click

        Call AddTo(False)

    End Sub
End Class