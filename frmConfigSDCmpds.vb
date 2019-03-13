Option Compare Text

Public Class frmConfigSDCmpds

    Public boolCancel As Boolean = True
    Public idA As Int64
    Public idS As Int64
    Public idS1 As Int64


    Private Sub frmConfigSDCmpds_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub cmdOK1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK1.Click

        boolCancel = False
        Me.Visible = False

    End Sub

    Private Sub cmdCancel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel1.Click

        tblGuWuAssignedCmpd.RejectChanges()

        boolCancel = True
        Me.Visible = False

    End Sub

    Sub FormLoad()

        Call Configure_dgvCmpd()

        Call ConfigSource()


    End Sub

    Sub Configure_dgvCmpd()
        Dim dgv As DataGridView
        Dim Count1 As Short
        Dim cw As Single

        'If boolSourceSD Then
        '    dgv = Me.dgvCmpd
        'Else
        '    GoTo end1
        'End If

        cw = 15

        dgv = Me.dgvCmpd

        For Count1 = 1 To 1
            Select Case Count1
                Case 1
                    dgv = Me.dgvCmpd
                    cw = 10


            End Select
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

            Select Case Count1
                Case 1
                    Call ConfigCmpdTableSD(True)

            End Select
        Next

end1:

    End Sub

    Sub ConfigSource()

        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim dgv As DataGridView
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim boolVis As Boolean
        Dim cw As Single

        tbl = tblGuWuCompounds
        dgv = Me.dgvSource

        strF = "BOOLINCLUDE = -1"

        Dim dv As System.Data.DataView = New DataView(tbl, strF, Nothing, DataViewRowState.CurrentRows)

        dv.AllowNew = False
        dv.AllowDelete = False
        dv.AllowEdit = False

        dgv.DataSource = dv

        dgv.ReadOnly = True

        For Count1 = 0 To dgv.Columns.Count - 1
            str3 = dgv.Columns(Count1).Name
            boolVis = False
            str2 = ""
            Select Case str3
                Case "ID_TBLGUWUCOMPOUNDS"
                    str1 = "ID_TBLGUWUCOMPOUNDS"
                    str2 = "ID_TBLGUWUCOMPOUNDS"
                    boolVis = False
                Case "CHARANALYTENAME"
                    str1 = "CHARANALYTENAME"
                    str2 = "Compound"
                    boolVis = True
                Case "CHARCOMPANYID"
                    str1 = "CHARCOMPANYID"
                    str2 = "Company ID"
                    boolVis = True
                Case "CHARCOMMENTS"
                    str1 = "CHARCOMMENTS"
                    str2 = "Comments"
                    boolVis = True
                Case "CHARALIAS"
                    str1 = "CHARALIAS"
                    str2 = "Alias"
                    boolVis = True
                Case "CHARIUPAC"
                    str1 = "CHARIUPAC"
                    str2 = "IUPAC"
                    boolVis = True
                Case "BOOLINCLUDE"
                    str1 = "BOOLINCLUDE"
                    str2 = "BOOLINCLUDE"
                    boolVis = False
                Case "NUMMW"
                    str1 = "NUMMW"
                    str2 = "Mol. Wt."
                    boolVis = True
                Case "CHARCHEMFORMULA"
                    str1 = "CHARCHEMFORMULA"
                    str2 = "Chem. Formula"
                    boolVis = True
                Case "CHARALTNAME1"
                    str1 = "CHARALTNAME1"
                    str2 = "User Text 1"
                    boolVis = True
                Case "CHARALTNAME2"
                    str1 = "CHARALTNAME2"
                    str2 = "User Text 2"
                    boolVis = True
                Case "CHARALTNAME3"
                    str1 = "CHARALTNAME3"
                    str2 = "User Text 3"
                    boolVis = True
                Case "CHARALTNAME4"
                    str1 = "CHARALTNAME4"
                    str2 = "User Text 4"
                    boolVis = True

            End Select

            dgv.Columns(Count1).HeaderText = str2
            dgv.Columns(Count1).Visible = boolVis
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


    End Sub


    Sub ConfigCmpdTableSD(ByVal boolW As Boolean)
        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim boolRO As Boolean
        Dim int1 As Short
        Dim intCol As Short
        Dim boolVis As Boolean

        dgv = Me.dgvCmpd
        'hide all columns
        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns.Item(Count1).Visible = False
            dgv.Columns.Item(Count1).ReadOnly = True
            'dgv.Columns.Item(Count1).DisplayIndex = dgv.Columns.Count - 1
            'dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        'enter column headertexts
        For Count1 = 0 To dgv.ColumnCount - 1
            str4 = dgv.Columns(Count1).Name
            str1 = ""
            boolRO = False
            Select Case str4
                Case "CHARCOMPOUND"
                    str1 = "CHARCOMPOUND"
                    str2 = "Compound"
                    boolRO = True
                    boolVis = True
                    intCol = Count1
                Case "CHARCOMPANYID"
                    str1 = "CHARCOMPANYID"
                    str2 = "Company ID"
                    boolRO = True
                    boolVis = True
                Case "CHARCOMPOUNDTYPE"
                    str1 = "CHARCOMPOUNDTYPE"
                    str2 = "Type"
                    boolRO = True
                    boolVis = True
            End Select
            If Len(str1) = 0 Then
            Else
                dgv.Columns.Item(str1).Visible = boolVis
                dgv.Columns.Item(str1).HeaderText = str2
                dgv.Columns.Item(str1).ReadOnly = boolRO
            End If
        Next

        'set first row as current row
        If dgv.Rows.Count = 0 Then

        Else
            dgv.CurrentCell = dgv.Rows.Item(0).Cells(intCol)
        End If
        dgv.AutoResizeColumns()

end1:

    End Sub




    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click

        Dim dgvS As DataGridView
        Dim dgvD As DataGridView
        Dim Count1 As Short
        Dim intRows As Short
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim row As DataGridViewRow
        Dim strM As String
        Dim maxID As Int64
        Dim var1, var2, var3
        Dim boolGo As Boolean
        Dim id As Int64
        Dim idC As Int64
        Dim boolMaxID As Boolean
        Dim intType As Short
        Dim strType As String

        dgvS = Me.dgvSource
        dgvD = Me.dgvCmpd

        If dgvS.SelectedRows.Count = 0 Then
            strM = "One or more rows must be selected in the Source table."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        strM = "Click 'Yes' if compound is Analyte."
        strM = strM & ChrW(10) & ChrW(10)
        strM = strM & "Click 'No' if compound is Internal Standard."
        intType = MsgBox(strM, MsgBoxStyle.YesNoCancel, "Choose compound type...")
        strType = "Analyte"
        If intType = 2 Then 'Cancel
            Exit Sub
        ElseIf intType = 6 Then 'Yes
            strType = "Analyte"
        Else 'no
            strType = "Int Std"
        End If

        tbl = tblGuWuAssignedCmpd

        maxID = GetMaxID("TBLGUWUASSIGNEDCMPD", 1, False)


        boolMaxID = False
        For Each row In dgvS.SelectedRows

            idC = row.Cells("ID_TBLGUWUCOMPOUNDS").Value
            'do not allow replicates
            boolGo = True
            For Count1 = 0 To dgvD.RowCount - 1
                id = dgvD("ID_TBLGUWUCOMPOUNDS", Count1).Value
                If id = idC Then
                    boolGo = False
                    Exit For
                End If
            Next

            If boolGo Then

                Dim nRow As DataRow = tbl.NewRow
                nRow.BeginEdit()

                maxID = maxID + 1
                boolMaxID = True
                nRow("ID_TBLGUWUASSIGNEDCMPD") = maxID
                nRow("ID_TBLGUWUSTUDIES") = idS
                nRow("ID_TBLSTUDIES") = idS1
                nRow("ID_TBLGUWUASSAY") = idA
                nRow("ID_TBLGUWUCOMPOUNDS") = idC

                var1 = row.Cells("CHARANALYTENAME").Value
                var2 = row.Cells("CHARCOMPANYID").Value

                nRow("CHARCOMPOUND") = var1
                nRow("CHARCOMPANYID") = var2
                nRow("CHARCOMPOUNDTYPE") = strType

                nRow.EndEdit()

                tbl.Rows.Add(nRow)

            End If

        Next

        If boolMaxID Then
            Dim bool As Boolean

            bool = PutMaxID("TBLGUWUASSIGNEDCMPD", maxID)

        End If

        Me.dgvCmpd.AutoResizeColumns()


    End Sub

    Private Sub cmdRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemove.Click

        Dim dgvS As DataGridView
        Dim dgvD As DataGridView
        Dim Count1 As Short
        Dim intRows As Short
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim row As DataGridViewRow
        Dim strM As String
        Dim maxID As Int64
        Dim var1, var2, var3
        Dim boolGo As Boolean
        Dim id As Int64
        Dim idC As Int64
        Dim strF As String

        dgvS = Me.dgvSource
        dgvD = Me.dgvCmpd

        If dgvD.SelectedRows.Count = 0 Then
            strM = "One or more rows must be selected in the Configured Compounds table."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        tbl = tblGuWuAssignedCmpd

        For Each row In dgvD.SelectedRows

            id = row.Cells("ID_TBLGUWUASSIGNEDCMPD").Value
            strF = "ID_TBLGUWUASSIGNEDCMPD = " & id
            rows = tbl.Select(strF)
            rows(0).Delete()

        Next

        Me.dgvCmpd.AutoResizeColumns()


    End Sub
End Class