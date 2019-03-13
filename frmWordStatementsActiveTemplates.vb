Public Class frmWordStatementsActiveTemplates

    Private Sub frmWordStatementsActiveTemplates_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call DoubleBufferControl(Me, "dgv")

        Call ControlDefaults(Me)

        Call PlaceForm()

        Call dgvConfig()
        Call dgvFilter()

        If gboolAuditTrail Then
            Me.lblNote.Visible = True
        Else
            Me.lblNote.Visible = False
        End If

    End Sub

    Sub PlaceForm()

        Dim a, b, c, d, e


    End Sub

    Sub dgvConfig()

        Dim dgv As DataGridView = Me.dgvReportStatements
        Dim tbl1 As DataTable = tblWordStatements
        Dim dv As DataView = New DataView(tbl1, "ID_TBLWORDSTATEMENTS > 0", "CHARTITLE ASC", DataViewRowState.CurrentRows)

        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        dgv.DataSource = dv

        Dim Count1 As Short
        Dim str1 As String

        For Count1 = 0 To dgv.ColumnCount - 1
            dgv.Columns(Count1).Visible = False
        Next

        str1 = "CHARTITLE"
        dgv.Columns(str1).HeaderText = "Report Template"
        dgv.Columns(str1).Visible = True

        str1 = "CHARWORDSTATEMENT"
        dgv.Columns(str1).HeaderText = "Status"
        dgv.Columns(str1).Visible = True

        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect


    End Sub

    Sub dgvFilter()


        Dim strF As String
        Dim dv As DataView = Me.dgvReportStatements.DataSource

        Try
            If Me.rbAll.Checked Then
                strF = "ID_TBLWORDSTATEMENTS > 0"
            ElseIf Me.rbActive.Checked Then
                strF = "CHARWORDSTATEMENT = 'Active'"
            ElseIf Me.rbInactive.Checked Then
                strF = "CHARWORDSTATEMENT = 'Inactive'"
            End If

            dv.RowFilter = strF
        Catch ex As Exception

        End Try


    End Sub

    Private Sub dgvReportStatements_DoubleClick(sender As Object, e As EventArgs) Handles dgvReportStatements.DoubleClick

        Call dgvDoubleClick()

    End Sub

    Sub dgvDoubleClick()

        Dim dgv As DataGridView = Me.dgvReportStatements
        Dim dv As DataView = dgv.DataSource
        Dim dtbl As DataTable = dv.ToTable

        Dim intRow As Int16
        Dim intRows As Int16
        Dim strM As String
        Dim str1 As String
        Dim str2 As String
        Dim intSel As Int16
        Dim strF As String
        Dim var1
        Dim id As Int64

        intRows = dgv.Rows.Count

        Try
            If dgv.RowCount = 0 Then
                GoTo end1
            End If
            If dgv.CurrentRow Is Nothing Then
                strM = "Please select a row"
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If
            intRow = dgv.CurrentRow.Index

            str1 = dgv("CHARWORDSTATEMENT", intRow).Value
            id = dgv("ID_TBLWORDSTATEMENTS", intRow).Value

            If StrComp(str1, "Active", CompareMethod.Text) = 0 Then
                str2 = "Inactive"
            Else
                str2 = "Active"
            End If

            strF = "ID_TBLWORDSTATEMENTS = " & id

            Dim rows() As DataRow = dtbl.Select(strF)
            rows(0).BeginEdit()
            rows(0).Item("CHARWORDSTATEMENT") = str2
            rows(0).EndEdit()

            dgv("CHARWORDSTATEMENT", intRow).Value = str2

            'choose another row in order to actuate
            If intRow = 0 Then
                intSel = intRow + 1
            ElseIf intRow = intRows - 1 Then
                intSel = intRow - 1
            Else
                intSel = intRow + 1
            End If
            If Me.rbAll.Checked Then
            Else
                Try
                    dgv.CurrentCell = dgv.Rows(intSel).Cells("CHARTITLE")
                Catch ex As Exception
                    var1 = ex.Message
                End Try
            End If


            If intRows = 1 Then
                'need to do something else to trigger change
                Try
                    'dgv.ClearSelection()
                    'dgv.Select()
                    Me.gbActive.Focus()

                Catch ex As Exception
                    var1 = ex.Message
                End Try

            End If

        Catch ex As Exception
            var1 = ex.Message
        End Try

end1:

    End Sub

    Private Sub rbAll_CheckedChanged(sender As Object, e As EventArgs) Handles rbAll.CheckedChanged

        Call dgvFilter()

    End Sub

    Private Sub rbActive_CheckedChanged(sender As Object, e As EventArgs) Handles rbActive.CheckedChanged

        Call dgvFilter()

    End Sub

    Private Sub rbInactive_CheckedChanged(sender As Object, e As EventArgs) Handles rbInactive.CheckedChanged

        Call dgvFilter()

    End Sub

    Private Sub cmdDone_Click(sender As Object, e As EventArgs) Handles cmdDone.Click

        'save everything
        Call DoSave()

        Me.Dispose()

    End Sub

    Sub DoSave()


        If boolGuWuOracle Then
            Try
                ta_tblWordStatements.Update(tblWordStatements)
            Catch ex As DBConcurrencyException
                'ds2005.TBLWORDSTATEMENTS.Merge('ds2005.TBLWORDSTATEMENTS, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblWordStatementsAcc.Update(tblWordStatements)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLWORDSTATEMENTS.Merge('ds2005Acc.TBLWORDSTATEMENTS, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblWordStatementsSQLServer.Update(tblWordStatements)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLWORDSTATEMENTS.Merge('ds2005Acc.TBLWORDSTATEMENTS, True)
            End Try
        End If


    End Sub

    Private Sub cmdCancelEdit_Click(sender As Object, e As EventArgs) Handles cmdCancelEdit.Click

        'Dim dgv As DataGridView = Me.dgvReportStatements
        'Dim dv As DataView = dgv.DataSource
        Dim dtbl As DataTable = tblWordStatements ' dv.ToTable
        dtbl.RejectChanges()

        Me.Dispose()

    End Sub

    Private Sub dgvReportStatements_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvReportStatements.CellContentClick

    End Sub
End Class