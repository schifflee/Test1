Option Compare Text

Public Class frmReplacePersonnel

    Public boolCancel As Boolean = True
    Public boolFormLoad As Boolean

    Private Sub frmReplacePersonnel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call ControlDefaults(Me)

        boolFormLoad = True

        Dim dgv1 As DataGridView

        Dim str1 As String

        str1 = "Select the Study from which Contributing Personnel are to be imported and replaced"
        str1 = str1 & ChrW(10) & "Then click OK or Cancel"
        Me.lblTitle.Text = str1

        Call ConfigStudy()
        Call ConfigPersonnel()

        'select appropriate row in dgvStudy
        Dim intID As Int64
        Dim dgv As DataGridView
        Dim Count1 As Int16

        dgv = Me.dgvStudy
        dgv.CurrentCell = dgv.Rows.Item(0).Cells("CHARWATSONSTUDYNAME")
        For Count1 = 0 To dgv.RowCount - 1

            intID = dgv("ID_TBLSTUDIES", Count1).Value
            If intID = id_tblStudies Then
                dgv.CurrentCell = dgv.Rows.Item(Count1).Cells("CHARWATSONSTUDYNAME")
                Exit For
            End If

        Next

        boolFormLoad = False
        Call StudySelect()

        Call SizeForm()

    End Sub

    Sub SizeForm()

        'Dim w, r

        'w = frmH.Width
        'r = frmH.lbxSymbol.Left + frmH.lbxSymbol.Width
        'Me.Left = 30
        'Me.Width = r - 30

    End Sub

    Sub StudySelect()

        If boolFormLoad Then
            Exit Sub
        End If

        Dim intID As Int64
        Dim intRow As Int16
        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim strF As String
        Dim strS As String

        dgv1 = Me.dgvStudy
        dgv2 = Me.dgvPersonnel

        intRow = dgv1.CurrentRow.Index

        intID = dgv1("ID_TBLSTUDIES", intRow).Value
        strF = "ID_TBLSTUDIES = " & intID

        Dim dv As System.Data.DataView
        dv = dgv2.DataSource

        dv.RowFilter = strF
        strS = "INTORDER ASC"
        dv.Sort = strS


    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        Dim dgv As DataGridView
        Dim intMaxID As Int64
        Dim strF As String
        Dim Count1 As Int16
        Dim Count2 As Short
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim var1
        Dim strName As String



        dgv = Me.dgvPersonnel
        tbl = tblContributingPersonnel

        intMaxID = GetMaxID("TBLCONTRIBUTINGPERSONNEL", dgv.RowCount, True)

        'first delete all rows from tbl
        strF = "ID_TBLSTUDIES = " & id_tblStudies
        rows = tbl.Select(strF)

        For Count1 = 0 To rows.Length - 1
            rows(Count1).BeginEdit()
            rows(Count1).Delete()
            rows(Count1).EndEdit()
        Next

        'add rows to tbl
        For Count1 = 0 To dgv.RowCount - 1

            Dim nr As DataRow = tbl.NewRow
            nr.BeginEdit()
            For Count2 = 1 To dgv.ColumnCount - 1
                strName = dgv.Columns(Count2).Name
                var1 = dgv(strName, Count1).Value
                nr.Item(strName) = var1
            Next
            intMaxID = intMaxID + 1
            nr.Item("ID_TBLCONTRIBUTINGPERSONNEL") = intMaxID
            nr.Item("ID_TBLSTUDIES") = id_tblStudies
            nr.EndEdit()
            tbl.Rows.Add(nr)
        Next

        'Call PutMaxID("TBLCONTRIBUTINGPERSONNEL", intMaxID)

        boolCancel = False
        Me.Visible = False

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        boolCancel = True
        Me.Visible = False

    End Sub

    Sub ConfigStudy()

        Dim tbl As System.Data.DataTable
        Dim strS As String
        Dim intRows As Int16

        tbl = tblStudies
        intRows = tbl.Rows.Count

        Dim dv As System.Data.DataView = New DataView(tbl)
        strS = "CHARWATSONSTUDYNAME ASC"
        dv.Sort = strS
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        Dim dgv As DataGridView
        dgv = Me.dgvStudy

        dgv.DataSource = dv

        Dim Count1 As Short

        For Count1 = 0 To dgv.ColumnCount - 1
            dgv.Columns(Count1).Visible = False
        Next

        dgv.Columns("CHARWATSONSTUDYNAME").Visible = True
        dgv.Columns("CHARWATSONSTUDYNAME").HeaderText = "Study"

        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv.Columns("CHARWATSONSTUDYNAME").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        dgv.Columns(0).Width = 100


    End Sub

    Sub ConfigPersonnel()

        Dim tbl As System.Data.DataTable

        tbl = tblContributingPersonnel

        Dim dv As System.Data.DataView = New DataView(tbl)
        Dim strS As String

        strS = "INTORDER ASC"

        dv.Sort = strS
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False


        Dim dgvS As DataGridView
        Dim dgvD As DataGridView

        dgvS = frmH.dgvContributingPersonnel
        dgvD = Me.dgvPersonnel

        dgvD.DataSource = dv

        Dim Count1 As Short

        For Count1 = 0 To dgvD.ColumnCount - 1
            dgvD.Columns(Count1).Visible = dgvS.Columns(Count1).Visible
            dgvD.Columns(Count1).HeaderText = dgvS.Columns(Count1).HeaderText
            dgvD.Columns(Count1).AutoSizeMode = DataGridViewAutoSizeColumnsMode.Fill
        Next

        dgvD.Columns("CHARCPNAME").AutoSizeMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
        dgvD.Columns("CHARCPTITLE").AutoSizeMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
        dgvD.Columns("CHARCPROLE").AutoSizeMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
        dgvD.Columns(0).Width = 100

    End Sub

    Private Sub dgvStudy_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvStudy.SelectionChanged

        'Call StudySelect()
        If boolFormLoad Then
            Exit Sub
        End If

        Call StudySelect()

    End Sub

End Class