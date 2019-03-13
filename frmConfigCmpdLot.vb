Option Compare Text

Public Class frmConfigCmpdLot
    Public boolCancel As Boolean = True
    Public idA As Int64
    Public idS As Int64
    Public idS1 As Int64
    Public idC As Int64
    Public idC1 As Int64


    Private Sub cmdOK1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK1.Click

        boolCancel = False
        Me.Visible = False

    End Sub

    Private Sub cmdCancel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel1.Click

        boolCancel = True
        Me.Visible = False

    End Sub

    Private Sub frmConfigCmpdLot_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Sub FormLoad()

        Dim dgv As DataGridView
        Dim strF As String
        Dim strS As String
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim boolVis As Boolean

        'load dgvSource
        strF = "ID_TBLGUWUCOMPOUNDS = " & idC1
        strS = "CHARLOTNUMBER ASC"
        Dim dv As System.Data.DataView = New DataView(tblGuWuCompoundsInd, strF, strS, DataViewRowState.CurrentRows)

        dv.RowFilter = strF
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        If dv.Count = 0 Then
            Me.lblNone.Visible = True
        Else
            Me.lblNone.Visible = False
        End If

        dgv = Me.dgvSource

        dgv.DataSource = dv

        For Count1 = 0 To dgv.Columns.Count - 1
            str3 = dgv.Columns(Count1).Name
            str1 = "Col" & Count1
            str2 = "Col" & Count1
            boolVis = False
            Select Case str3
                Case "ID_TBLGUWUCOMPOUNDSIND"
                    str1 = "ID_TBLGUWUCOMPOUNDSIND"
                    str2 = "ID_TBLGUWUCOMPOUNDSIND"
                    boolVis = False
                Case "ID_TBLGUWUCOMPOUNDS"
                    str1 = "ID_TBLGUWUCOMPOUNDS"
                    str2 = "ID_TBLGUWUCOMPOUNDS"
                    boolVis = False
                Case "CHARLOTNUMBER"
                    str1 = "CHARLOTNUMBER"
                    str2 = "Lot Number"
                    boolVis = True
                Case "CHARPHYSICALDESCRIPTION"
                    str1 = "CHARPHYSICALDESCRIPTION"
                    str2 = "Physical Descr."
                    boolVis = True
                Case "CHARSTORAGECONDITIONS"
                    str1 = "CHARSTORAGECONDITIONS"
                    str2 = "Storage Conditions"
                    boolVis = True
                Case "CHARDATERECEIVED"
                    str1 = "CHARDATERECEIVED"
                    str2 = "Date Received"
                    boolVis = True
                Case "CHAREXPIRATIONRETESTDATE"
                    str1 = "CHAREXPIRATIONRETESTDATE"
                    str2 = "Expiration Retest Date"
                    boolVis = True
                Case "CHARAMOUNTRECEIVED"
                    str1 = "CHARAMOUNTRECEIVED"
                    str2 = "Amount Received"
                    boolVis = True
                Case "CHARSUPPLIER"
                    str1 = "CHARSUPPLIER"
                    str2 = "Supplier"
                    boolVis = True
                Case "CHARPURITY"
                    str1 = "CHARPURITY"
                    str2 = "Purity"
                    boolVis = True
                Case "CHARPERCENTWATER"
                    str1 = "CHARPERCENTWATER"
                    str2 = "% Water"
                    boolVis = True
                Case "CHARCOMMENTS"
                    str1 = "CHARCOMMENTS"
                    str2 = "Comments"
                    boolVis = True
                Case "BOOLINCLUDE"
                    str1 = "BOOLINCLUDE"
                    str2 = "BOOLINCLUDE"
                    boolVis = False
                Case "CHARCERTOFANALYSIS"
                    str1 = "CHARCERTOFANALYSIS"
                    str2 = "Cert. of Analysis?"
                    boolVis = True

            End Select

            dgv.Columns(Count1).Visible = boolVis
            dgv.Columns(Count1).HeaderText = str2

        Next

        Dim cw As Single

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


        dgv = Me.dgvLot
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

        Call DoAddLabel()


    End Sub

    Sub DoAddLabel()

        Dim str1 As String

        str1 = "<- Add"

        If Me.dgvLot.RowCount = 0 Then
            str1 = "<- Add"
        Else
            str1 = "<- Replace"
        End If

        Me.cmdAdd.Text = str1

    End Sub

    Sub AddClick()

        Dim dgvS As DataGridView
        Dim dgvD As DataGridView
        Dim intRow As Short
        Dim strM As String
        Dim id As Int64
        Dim strL As String
        Dim tbl As System.Data.DataTable
        Dim maxID As Int64

        dgvS = Me.dgvSource
        dgvD = Me.dgvLot

        If dgvS.RowCount = 0 Then
            Exit Sub
        End If

        If dgvS.CurrentRow Is Nothing Then
            strM = "Please select a Lot Number."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            GoTo end1
        End If

        intRow = dgvS.CurrentRow.Index

        If dgvD.RowCount = 0 Then 'add a row
            id = dgvS("ID_TBLGUWUCOMPOUNDSIND", intRow).Value
            strL = NZ(dgvS("CHARLOTNUMBER", intRow).Value, "NA")

            'Public idA As Int64
            'Public idS As Int64
            'Public idS1 As Int64
            'Public idC As Int64
            'Public idC1 As Int64

            tbl = tblGuWuAssignedCmpdLot
            Dim nRow As DataRow = tbl.NewRow
            nRow.BeginEdit()
            nRow("CHARLOTNUMBER") = strL
            nRow("ID_TBLGUWUASSIGNEDCMPDLOT") = GetMaxID("TBLGUWUCOMPOUNDSIND", 1, True)
            nRow("ID_TBLGUWUASSIGNEDCMPD") = idC
            nRow("ID_TBLGUWUASSAY") = idA
            nRow("ID_TBLGUWUSTUDIES") = idS
            nRow("ID_TBLSTUDIES") = idS1
            nRow("ID_TBLGUWUCOMPOUNDS") = idC1
            nRow("ID_TBLGUWUCOMPOUNDSIND") = id

            nRow.EndEdit()

            tbl.Rows.Add(nRow)

        Else 'replace info
            Dim dv As System.Data.DataView

            id = dgvS("ID_TBLGUWUCOMPOUNDSIND", intRow).Value
            strL = NZ(dgvS("CHARLOTNUMBER", intRow).Value, "NA")

            dv = dgvD.DataSource
            dv.AllowEdit = True
            dv(0).BeginEdit()
            dv(0).Item("ID_TBLGUWUCOMPOUNDSIND") = id
            dv(0).Item("CHARLOTNUMBER") = strL
            dv(0).EndEdit()
            dv.AllowEdit = False
        End If

        Me.dgvLot.AutoResizeColumns()

        Call DoAddLabel()

end1:
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click

        Call AddClick()

    End Sub

    Private Sub cmdRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemove.Click



        Dim dgvS As DataGridView
        Dim dgvD As DataGridView
        Dim intRow As Short
        Dim strM As String
        Dim id As Int64
        Dim strL As String
        Dim tbl As System.Data.DataTable
        Dim maxID As Int64

        dgvS = Me.dgvSource
        dgvD = Me.dgvLot

        If dgvD.RowCount = 0 Then 'ignore

        Else 'replace info
            Dim dv As System.Data.DataView

            dv = dgvD.DataSource
            dv.AllowDelete = True
            dv(0).Delete()
            dv.AllowDelete = False
        End If

        Me.dgvLot.AutoResizeColumns()

        Call DoAddLabel()

end1:

    End Sub

    Private Sub dgvSource_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSource.DoubleClick

        Call AddClick()

    End Sub
End Class