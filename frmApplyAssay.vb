Option Compare Text

Public Class frmApplyAssay
    Public boolCancel As Boolean = True
    Public boolFromApply As Boolean = False
    Public boolFromNew As Boolean = False
    Public idS As Int64
    Public idT As Int64
    Public idA As Int64


    Private Sub frmApplyAssay_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Sub FormLoad()

        If boolFromNew Then
            Me.panTemplate.Visible = False
            Me.Height = Me.panNew.Top + Me.panNew.Height + 30
        Else
            Me.panNew.Visible = False
            Me.panTemplate.Visible = True
            Me.panTemplate.Top = Me.panNew.Top
            Me.Height = Me.panTemplate.Top + Me.panTemplate.Height + 30
        End If

        'reposition form
        Dim w, h, t1, h1, w1

        w = My.Computer.Screen.WorkingArea.Width
        h = My.Computer.Screen.WorkingArea.Height

        t1 = Me.Top
        h1 = Me.Height
        w1 = Me.Width

        Me.Top = (h - h1) / 2
        Me.Left = (w - w1) / 2

        Dim tbl As System.Data.DataTable
        Dim tbl1 As System.Data.DataTable
        Dim dv As System.Data.DataView
        Dim dv1 As System.Data.DataView
        Dim dgv As DataGridView
        Dim dgv1 As DataGridView
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim boolVis As Boolean
        Dim strS As String

        Dim strF As String

        strF = "BOOLINCLUDE = -1"

        tbl = tblGuWuStudies
        tbl1 = tblGuWuAssay

        dv = New DataView(tbl, strF, Nothing, DataViewRowState.CurrentRows)
        dv1 = New DataView(tbl1, strF, Nothing, DataViewRowState.CurrentRows)

        dgv = Me.dgvStudy
        dgv1 = Me.dgvAssay

        'fill dgvStudies
        'do dgvAssay first
        dgv1.DataSource = dv1
        dv1.AllowDelete = False
        dv1.AllowEdit = False
        dv1.AllowNew = False

        For Count1 = 0 To dgv1.Columns.Count - 1
            str1 = dgv1.Columns(Count1).Name
            str2 = str1
            boolVis = False
            Select Case str1
                Case "CHARASSAYNAME"
                    str2 = "Assay"
                    boolVis = True
                    dgv1.Columns.Item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            End Select
            dgv1.Columns(Count1).ReadOnly = True
            dgv1.Columns(Count1).Visible = boolVis
            dgv1.Columns(Count1).HeaderText = str2
        Next

        dgv1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        dgv1.AllowUserToResizeColumns = True
        dgv1.AllowUserToResizeRows = True
        dgv1.RowHeadersWidth = 10
        dgv1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        dgv1.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgv1.AutoResizeColumns()
        dgv1.CurrentCell = dgv1.Item("CHARASSAYNAME", 0)


        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False
        If Me.rbStudyName.Checked Then
            strS = "CHARSTUDYNAME ASC"
        Else
            strS = "CHARSTUDYNUMBER ASC"
        End If
        dv.Sort = strS
        'now dow dgvStudy
        dgv.DataSource = dv

        For Count1 = 0 To dgv.Columns.Count - 1
            str1 = dgv.Columns(Count1).Name
            str2 = str1
            boolVis = False
            Select Case str1
                Case "CHARSTUDYNAME"
                    str2 = "Study Name"
                    boolVis = True
                    dgv.Columns.Item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                Case "CHARSTUDYNUMBER"
                    str2 = "Study Number"
                    boolVis = True
                    dgv.Columns.Item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            End Select
            dgv.Columns(Count1).ReadOnly = True
            dgv.Columns(Count1).Visible = boolVis
            dgv.Columns(Count1).HeaderText = str2
        Next

        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        dgv.AllowUserToResizeColumns = True
        dgv.AllowUserToResizeRows = True
        dgv.RowHeadersWidth = 10
        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgv.AutoResizeColumns()
        dgv.CurrentCell = dgv.Item("CHARSTUDYNAME", 0)
        dgv.CurrentRow.Selected = True

    End Sub

    Sub FilterStudy()

        Dim dgv As DataGridView
        Dim intRow As Short
        Dim strF As String
        Dim id As Int64
        Dim dv As System.Data.DataView
        Dim strS As String

        dgv = Me.dgvStudy
        'intRow = dgv.CurrentRow.Index

        If dgv.Rows.Count = 0 Then
            id = -1
        Else
            If dgv.CurrentRow Is Nothing Then
                intRow = 0
            Else
                intRow = dgv.CurrentRow.Index
            End If
            id = dgv("ID_TBLGUWUSTUDIES", intRow).Value
        End If

        strF = "ID_TBLGUWUSTUDIES = " & id & " AND ID_TBLGUWUASSAY <> " & idA
        strS = "CHARASSAYNAME ASC"
        dv = Me.dgvAssay.DataSource
        dv.RowFilter = strF
        dv.Sort = strS

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        boolCancel = True
        Me.Visible = False

    End Sub

    Private Sub cmdApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        boolCancel = False
        Me.Visible = False

    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        If FinalNameValidate() Then

            boolCancel = False
            Me.Visible = False

        End If

    End Sub

    Private Sub chkApplyTemplate_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkApplyTemplate.CheckedChanged

        If Me.chkApplyTemplate.Checked Then
            Me.panTemplate.Visible = True
            Me.Height = Me.panTemplate.Top + Me.panTemplate.Height + 60
        Else
            Me.panTemplate.Visible = False
            Me.Height = Me.panNew.Top + Me.panNew.Height + 60
        End If

        'reposition form
        Dim w, h, t1, h1, w1

        w = My.Computer.Screen.WorkingArea.Width
        h = My.Computer.Screen.WorkingArea.Height

        t1 = Me.Top
        h1 = Me.Height
        w1 = Me.Width

        Me.Top = (h - h1) / 2
        Me.Left = (w - w1) / 2

    End Sub

    Private Sub dgvStudy_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvStudy.SelectionChanged

        Call FilterStudy()

    End Sub

    Private Sub txtStudyFilter_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtStudyFilter.TextChanged

        Dim str1 As String
        Dim boolName As Boolean

        If Me.rbStudyName.Checked Then
            boolName = True
        Else
            boolName = False
        End If

        str1 = Me.txtStudyFilter.Text

        Dim dv As System.Data.DataView
        Dim strF As String
        Dim strS As String

        dv = Me.dgvStudy.DataSource

        'If boolName Then
        '    strF = "CHARSTUDYNAME IS LIKE %" & str1 & "%"
        '    strS = "CHARSTUDYNAME ASC"
        'Else
        '    strF = "CHARSTUDYNUMBER IS LIKE %" & str1 & "%"
        '    strS = "CHARSTUDYNUMBER ASC"
        'End If

        'dv.RowFilter = strF
        'dv.Sort = strS

        'If boolName Then
        '    strF = "CHARSTUDYNAME IS LIKE *" & str1 & "*"
        '    strS = "CHARSTUDYNAME ASC"
        'Else
        '    strF = "CHARSTUDYNUMBER IS LIKE *" & str1 & "*"
        '    strS = "CHARSTUDYNUMBER ASC"
        'End If

        'dv.RowFilter = strF
        'dv.Sort = strS

        'If boolName Then
        '    strF = "CHARSTUDYNAME  LIKE *" & str1 & "*"
        '    strS = "CHARSTUDYNAME ASC"
        'Else
        '    strF = "CHARSTUDYNUMBER  LIKE *" & str1 & "*"
        '    strS = "CHARSTUDYNUMBER ASC"
        'End If

        'dv.RowFilter = strF
        'dv.Sort = strS

        If boolName Then
            strF = "CHARSTUDYNAME  LIKE '*" & str1 & "*'"
            strS = "CHARSTUDYNAME ASC"
        Else
            strF = "CHARSTUDYNUMBER  LIKE '*" & str1 & "*'"
            strS = "CHARSTUDYNUMBER ASC"
        End If

        dv.RowFilter = strF
        dv.Sort = strS

    End Sub

    Private Sub txtAssayName_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtAssayName.Validating

        Dim str1 As String
        Dim boolhit As Boolean
        Dim strM As String
        Dim Count1 As Short
        Dim strName As String

        'ensure name is unique

        str1 = Me.txtAssayName.Text

        If Len(str1) = 0 Then
            Exit Sub
        Else
            strName = str1
        End If

        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String

        strF = "ID_TBLGUWUSTUDIES = " & idS
        tbl = tblGuWuAssay
        rows = tbl.Select(strF)

        If rows.Length = 0 Then
            boolhit = False
        Else

            For Count1 = 0 To rows.Length - 1
                str1 = rows(Count1).Item("CHARASSAYNAME")
                If StrComp(str1, strName, CompareMethod.Text) = 0 Then
                    strM = "Proposed Assay Name '" & strName & "' already exists in this study."
                    strM = strM & ChrW(10) & ChrW(10) & "Please enter a unique name."
                    MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
                    e.Cancel = True
                End If
            Next
        End If

    End Sub

    Function FinalNameValidate() As Boolean

        Dim str1 As String
        Dim strM As String

        'ensure name is unique

        str1 = Me.txtAssayName.Text

        If Len(str1) = 0 Then
            strM = "Proposed Assay Name cannot be blank."
            strM = strM & ChrW(10) & ChrW(10) & "Please enter a unique name."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
            FinalNameValidate = False
            GoTo end1
        Else
            FinalNameValidate = True
        End If

        ''make sure selected assay isn't the same assay
        'If boolFromApply Then
        '    If idA = idT Then
        '        strM = "The chosen Assay Template is the same template."
        '        strM = strM & ChrW(10) & ChrW(10) & "Please enter a unique name."
        '        MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
        '        FinalNameValidate = False
        '    Else
        '        FinalNameValidate = True
        '    End If
        'Else
        '    FinalNameValidate = True
        'End If

end1:

    End Function

    Private Sub rbStudyName_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbStudyName.CheckedChanged

        Call FilterChange()
    End Sub

    Private Sub rbStudyNumber_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbStudyNumber.CheckedChanged

        Call FilterChange()

    End Sub

    Sub FilterChange()

        Dim strS As String
        Dim dv As System.Data.DataView

        dv = Me.dgvStudy.DataSource

        If Me.rbStudyName.Checked Then
            strS = "CHARSTUDYNAME ASC"
        Else
            strS = "CHARSTUDYNUMBER ASC"
        End If
        Try
            dv.Sort = strS
        Catch ex As Exception

        End Try

    End Sub

    Private Sub dgvAssay_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAssay.CellClick

        Call RecordidT()

    End Sub


    Private Sub dgvAssay_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAssay.SelectionChanged

        Call RecordidT()

    End Sub

    Sub RecordidT()

        Dim dgv As DataGridView
        Dim intRow As Short

        dgv = Me.dgvAssay
        If dgv.Rows.Count = 0 Then
            idT = -1
        Else
            If dgv.CurrentRow Is Nothing Then
                intRow = 0
            Else
                intRow = dgv.CurrentRow.Index
            End If
            idT = dgv("ID_TBLGUWUASSAY", intRow).Value
        End If

    End Sub
End Class