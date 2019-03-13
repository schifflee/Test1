Option Compare Text

Public Class frmConfigAssayPers

    Public boolCancel As Boolean = True
    Public idA As Int64
    Public idS As Int64
    Public idS1 As Int64

    'Private Sub cmdOK1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

    '    'boolCancel = False
    '    'Me.Visible = False


    'End Sub

    'Private Sub cmdCancel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel1.Click

    '    boolCancel = True
    '    Me.Visible = False

    'End Sub

    Sub FormLoad()

        Dim dgv As DataGridView
        Dim strF As String
        Dim strS As String
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim boolVis As Boolean
        Dim cw As Single

        'load dgvSource
        strF = "BOOLACTIVE = -1"
        strS = "CHARLASTNAME ASC"
        Dim dv As System.Data.DataView = New DataView(tblPersonnel, strF, strS, DataViewRowState.CurrentRows)

        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        'If dv.Count = 0 Then
        '    Me.lblNone.Visible = True
        'Else
        '    Me.lblNone.Visible = False
        'End If

        dgv = Me.dgvSource

        dgv.DataSource = dv

        For Count1 = 0 To dgv.Columns.Count - 1
            str3 = dgv.Columns(Count1).Name
            str1 = "Col" & Count1
            str2 = "Col" & Count1
            boolVis = False
            Select Case str3
                Case "CHARLASTNAME"
                    str1 = "CHARLASTNAME"
                    str2 = "Last Name"
                    boolVis = True
                Case "CHARFIRSTNAME"
                    str1 = "CHARFIRSTNAME"
                    str2 = "First Name"
                    boolVis = True
                Case "CHARMIDDLENAME"
                    str1 = "CHARMIDDLENAME"
                    str2 = "Middle Name"
                    boolVis = True
            End Select

            dgv.Columns(Count1).Visible = boolVis
            dgv.Columns(Count1).HeaderText = str2

        Next

        cw = 25

        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgv.AllowUserToResizeColumns = True
        dgv.AllowUserToResizeRows = True
        dgv.RowHeadersWidth = cw '25
        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        dgv.AutoResizeColumns()


        'dO dgvpi
        dgv = Me.dgvPI

        cw = 25

        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgv.AllowUserToResizeColumns = True
        dgv.AllowUserToResizeRows = True
        dgv.RowHeadersWidth = cw '25
        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        dgv.AutoResizeColumns()


        'do dgvanal
        dgv = Me.dgvAnal
        cw = 25

        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgv.AllowUserToResizeColumns = True
        dgv.AllowUserToResizeRows = True
        dgv.RowHeadersWidth = cw '25
        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        dgv.AutoResizeColumns()


    End Sub

    Private Sub cmdAddPI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddPI.Click

        Call AddPers(Me.dgvPI, "PI")

    End Sub

    Sub RemovePers(ByVal dgvD As DataGridView, ByVal strPer As String)

        Dim dgvS As DataGridView
        Dim intRow As Short
        Dim strM As String
        Dim id As Int64
        Dim id1 As Int64
        Dim strL As String
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim maxID As Int64
        Dim boolMaxID As Boolean
        Dim Count1 As Short
        Dim dv As System.Data.DataView
        Dim boolGo As Boolean
        Dim dt As Date
        Dim intM As Short
        Dim strF As String

        dt = Now

        If dgvD.RowCount = 0 Then
            Exit Sub
        End If

        tbl = tblGuWuAssayPERS

        If dgvD.CurrentRow Is Nothing Then
            strM = "Please select a person from the " & strPer & " list."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            GoTo end1
        End If

        For Each sr As DataGridViewRow In dgvD.SelectedRows
            id = sr.Cells("ID_TBLPERSONNEL").Value
            strF = "ID_TBLPERSONNEL = " & id & " AND ID_TBLGUWUASSAY = " & idA & " AND CHARROLE = '" & strPer & "'"
            Erase rows
            rows = tbl.Select(strF)
            rows(0).BeginEdit()
            rows(0).Item("DTREMOVED") = dt
            rows(0).EndEdit()

        Next

        dgvD.AutoResizeColumns()

end1:

    End Sub

    Sub AddPers(ByVal dgvD As DataGridView, ByVal strPer As String)

        Dim dgvS As DataGridView
        Dim intRow As Short
        Dim strM As String
        Dim id As Int64
        Dim id1 As Int64
        Dim strL As String
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim maxID As Int64
        Dim boolMaxID As Boolean
        Dim Count1 As Short
        Dim dv As System.Data.DataView
        Dim boolGo As Boolean
        Dim boolGo1 As Boolean
        Dim dt As Date
        Dim intM As Short
        Dim strF As String
        Dim strLN As String
        Dim strMN As String
        Dim strFN As String
        Dim var1

        dt = Now

        dgvS = Me.dgvSource

        If dgvS.RowCount = 0 Then
            Exit Sub
        End If

        If dgvS.CurrentRow Is Nothing Then
            strM = "Please select a person from the Source list."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            GoTo end1
        End If

        intRow = dgvS.CurrentRow.Index

        intM = 0
        tbl = tblGuWuAssayPERS
        For Each sr As DataGridViewRow In dgvS.Rows
            If sr.Selected Then
                id = sr.Cells("ID_TBLPERSONNEL").Value
                intRow = sr.Index

                'ensure id is unique
                dv = dgvD.DataSource
                boolGo = True
                For Count1 = 0 To dv.Count - 1
                    id1 = dv(Count1).Item("ID_TBLPERSONNEL")
                    If id1 = id Then
                        boolGo = False
                        Exit For
                    End If
                Next

                If boolGo Then 'continue

                    'check for existing, only previously removed
                    boolGo1 = True
                    strF = "ID_TBLPERSONNEL = " & id & " AND ID_TBLGUWUASSAY = " & idA & " AND CHARROLE = '" & strPer & "'"
                    Erase rows
                    rows = tbl.Select(strF)
                    If rows.Length = 0 Then
                        boolGo1 = True
                    Else
                        boolGo1 = False
                    End If

                    If boolGo1 Then
                        intM = intM + 1
                        If intM = 1 Then
                            maxID = GetMaxID("TBLGUWUASSAYPERS", 1, False)
                        Else
                            maxID = maxID + 1
                        End If
                        Dim nRow As DataRow = tbl.NewRow
                        nRow.BeginEdit()
                        nRow("CHARROLE") = strPer
                        nRow("ID_TBLGUWUASSAY") = idA
                        nRow("ID_TBLGUWUSTUDIES") = idS
                        nRow("ID_TBLSTUDIES") = idS1
                        nRow("ID_TBLPERSONNEL") = id
                        nRow("DTASSIGNED") = dt
                        nRow("ID_TBLGUWUASSAYPERS") = maxID
                        var1 = FullName(intRow) 'debugging
                        nRow("CHARPERSONNEL") = CStr(var1)
                        boolMaxID = True
                        nRow.EndEdit()

                        tbl.Rows.Add(nRow)

                    Else 'modify existing row
                        rows(0).BeginEdit()
                        rows(0).Item("DTASSIGNED") = dt
                        rows(0).Item("DTREMOVED") = DBNull.Value
                        var1 = FullName(intRow) 'debugging
                        rows(0).Item("CHARPERSONNEL") = CStr(var1)
                        rows(0).EndEdit()
                    End If

                End If
            End If
        Next
        For Each sr As DataGridViewRow In dgvS.SelectedRows

        Next


        If intM > 1 Then
            boolMaxID = PutMaxID("TBLGUWUASSAYPERS", maxID)
        End If

        dgvD.AutoResizeColumns()


end1:

    End Sub

    Function FullName(ByVal intRow As Int16)

        Dim strLN As String
        Dim strMN As String
        Dim strFN As String
        Dim dgv As DataGridView

        dgv = Me.dgvSource

        strLN = Trim(NZ(dgv("CHARLASTNAME", intRow).Value, ""))
        strMN = Trim(NZ(dgv("CHARMIDDLENAME", intRow).Value, ""))
        strFN = Trim(NZ(dgv("CHARFIRSTNAME", intRow).Value, ""))

        If Len(strMN) = 0 Then
            FullName = strLN & ", " & strFN
        Else
            FullName = strLN & ", " & strFN & " " & strMN
        End If


    End Function

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        boolCancel = False
        Me.Visible = False

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        tblGuWuAssayPERS.RejectChanges()

        boolCancel = True
        Me.Visible = False

    End Sub

    Private Sub cmdRemovePI_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdRemovePI.Click

        Call RemovePers(Me.dgvPI, "PI")

    End Sub

    Private Sub cmdAddAnal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAddAnal.Click

        Call AddPers(Me.dgvAnal, "Analyst")

    End Sub

    Private Sub cmdRemoveAnal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdRemoveAnal.Click

        Call RemovePers(Me.dgvAnal, "Analyst")

    End Sub
End Class