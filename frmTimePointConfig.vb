Option Compare Text

Public Class frmTimePointConfig

    Public boolCancel As Boolean = True
    Public tblTP As New System.Data.DataTable
    Public dvTimePoints As System.Data.DataView
    Public idS As Int64
    Public idS1 As Int64
    Public idA As Int64
    Public idG As Int64
    Public idR As Int64
    Public boolTP As Boolean = False
    Public boolTPName As Boolean = False

    Private Sub frmTimePointConfig_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Sub FormLoad()

        Dim dgv As DataGridView
        Dim tbl As System.Data.DataTable
        Dim dv As System.Data.DataView
        Dim dv1 As System.Data.DataView
        Dim strF As String
        Dim strS As String
        Dim intRow As Short
        Dim id As Int64

        Dim str1 As String

        str1 = "NOTE: All actions performed in the Time Point Saved Sets pane are final and cannot be reversed."
        Me.lblTP.Text = str1

        tbl = TBLGUWUTPNAMESCONFIG
        dgv = Me.dgvTimepointSets
        strS = "CHARTPSETNAME ASC"

        If Me.rbAll.Checked Then
            strF = "BOOLINCLUDE > -2"
        ElseIf Me.rbActive.Checked Then
            strF = "BOOLINCLUDE = -1"
        ElseIf Me.rbInactive.Checked Then
            strF = "BOOLINCLUDE = 0"
        End If

        dv = New DataView(tbl, strF, strS, DataViewRowState.CurrentRows)

        dv.AllowDelete = False
        dv.AllowNew = False
        dv.Sort = strS

        dgv.DataSource = dv

        Call ConfigTP(dgv, "CHARTPSETNAME", "Set Name")

        If dgv.Rows.Count = 0 Then
            GoTo end1
        End If

        If dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If

        dgv.CurrentCell = dgv.Rows(intRow).Cells("CHARTPSETNAME")
        dgv.CurrentRow.Selected = True
        id = dgv("ID_TBLGUWUTPNAMESCONFIG", intRow).Value


        tbl = TBLGUWUTPCONFIG
        dgv = Me.dgvTP

        strF = "ID_TBLGUWUTPNAMESCONFIG = " & id
        strS = "NUMTIMEPOINT ASC"

        dv1 = New DataView(tbl, strF, strS, DataViewRowState.CurrentRows)

        dv1.AllowDelete = False
        dv1.AllowNew = False
        dv1.Sort = strS

        dgv.DataSource = dv1

        Call ConfigTP(dgv, "NUMTIMEPOINT", "Time" & ChrW(10) & "Points" & ChrW(10) & "(hrs)")

        If dgv.Rows.Count = 0 Then
            GoTo end1
        End If

        If dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If

        dgv.CurrentCell = dgv.Rows(intRow).Cells("NUMTIMEPOINT")
        dgv.CurrentRow.Selected = True

end1:

        'now load existing time points
        dgv = Me.dgvTimePoints

        Dim var1


        dgv.DataSource = Me.dvTimePoints

        var1 = Me.dvTimePoints.RowFilter

        Call ConfigTP(dgv, "NUMTIMEPOINT", "Time" & ChrW(10) & "Points" & ChrW(10) & "(hrs)")

        If dgv.Rows.Count = 0 Then
            GoTo end2
        End If

        If dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If

        dgv.CurrentCell = dgv.Rows(intRow).Cells("NUMTIMEPOINT")
        dgv.CurrentRow.Selected = True

end2:

    End Sub

    Sub ChangeTPSet()

        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim intRow As Short
        Dim dv As System.Data.DataView
        Dim strF As String
        Dim strS As String
        Dim id As Int64

        dgv1 = Me.dgvTimepointSets
        If dgv1.Rows.Count = 0 Then
            Exit Sub
        End If

        If dgv1.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv1.CurrentRow.Index
        End If

        id = dgv1("ID_TBLGUWUTPNAMESCONFIG", intRow).Value
        strF = "ID_TBLGUWUTPNAMESCONFIG = " & id
        strS = "NUMTIMEPOINT ASC"

        dgv2 = Me.dgvTP

        dv = dgv2.DataSource
        'dv.RowFilter = strF
        Try
            dv.RowFilter = strF
            If dgv2.Rows.Count = 0 Then
                GoTo end1
            Else
                intRow = 0
            End If
            dgv2.CurrentCell = dgv2.Rows(intRow).Cells("NUMTIMEPOINT")
            dgv2.CurrentRow.Selected = True
        Catch ex As Exception

        End Try



end1:

    End Sub

    Sub ConfigTP(ByVal dgv As DataGridView, ByVal strColumn As String, ByVal strName As String)

        Dim cw As Int16
        Dim Count1 As Short

        cw = 15

        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgv.AllowUserToResizeColumns = True
        dgv.AllowUserToResizeRows = True
        dgv.RowHeadersWidth = cw '25
        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        For Count1 = 0 To dgv.ColumnCount - 1
            dgv.Columns(Count1).Visible = False
        Next

        dgv.Columns(strColumn).Visible = True
        dgv.Columns(strColumn).ReadOnly = True
        dgv.Columns(strColumn).HeaderText = strName
        If StrComp(strName, "Set Name", CompareMethod.Text) = 0 Then
        Else
            dgv.Columns(strColumn).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        End If

    End Sub

    Private Sub cmdOK1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK1.Click

        boolCancel = False
        Me.Visible = False

    End Sub

    Private Sub cmdCancel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel1.Click

        boolCancel = True
        Me.Visible = False

    End Sub

    Private Sub dgvTimepointSets_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvTimepointSets.SelectionChanged

        Call ChangeTPSet()

    End Sub

    Private Sub rbAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbAll.CheckedChanged

        Call ChangeActiveFilter()

    End Sub

    Private Sub rbActive_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbActive.CheckedChanged

        Call ChangeActiveFilter()

    End Sub

    Private Sub rbInactive_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbInactive.CheckedChanged

        Call ChangeActiveFilter()

    End Sub

    Sub ChangeActiveFilter()

        Dim dgv As DataGridView
        Dim tbl As System.Data.DataTable
        Dim dv As System.Data.DataView
        Dim dv1 As System.Data.DataView
        Dim strF As String
        Dim strS As String
        Dim intRow As Short
        Dim id As Int64

        tbl = TBLGUWUTPNAMESCONFIG
        dgv = Me.dgvTimepointSets

        dv = dgv.DataSource

        If Me.rbAll.Checked Then
            strF = "BOOLINCLUDE > -2"
        ElseIf Me.rbActive.Checked Then
            strF = "BOOLINCLUDE = -1"
        ElseIf Me.rbInactive.Checked Then
            strF = "BOOLINCLUDE = 0"
        End If

        'dv.RowFilter = strF
        Try
            dv.RowFilter = strF
        Catch ex As Exception
            strS = ""
        End Try

    End Sub

    Private Sub cmdApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdApply.Click

        Dim dvS As System.Data.DataView
        Dim Count1 As Short
        Dim maxID As Int64
        Dim dgv As DataGridView

        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim strM As String

        dgv = Me.dgvTP

        Dim var1

        If dgv.RowCount = 0 Then
            strM = "A Time Point Set with actual values must be chosen."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        tbl = tblGuWuRTTimePoints


        strF = "ID_TBLGUWUPKROUTES = " & idR
        rows = tbl.Select(strF)

        'first delete all rows in dvd
        For Count1 = 0 To rows.Length - 1
            rows(Count1).Delete()
        Next

        maxID = GetMaxID("TBLGUWURTTIMEPOINTS", dgv.RowCount, True)

        Dim num1 As Single

        'now add dvs to dvd
        For Count1 = 0 To dgv.RowCount - 1
            Dim rowV As DataRow = tbl.NewRow
            rowV.BeginEdit()
            maxID = maxID + 1
            rowV("ID_TBLGUWURTTIMEPOINTS") = maxID
            rowV("ID_TBLGUWUSTUDIES") = idS
            rowV("ID_TBLSTUDIES") = idS1
            rowV("ID_TBLGUWUASSAY") = idA
            rowV("ID_TBLGUWUPKGROUPS") = idG
            rowV("ID_TBLGUWUPKROUTES") = idR
            var1 = dgv("NUMTIMEPOINT", Count1).Value
            rowV("NUMTIMEPOINT") = dgv("NUMTIMEPOINT", Count1).Value
            rowV.EndEdit()

            tbl.Rows.Add(rowV)

        Next

        Dim bool As Boolean

        '20190219 LEE: Don't need anymore. Used GetMaxID
        'bool = PutMaxID("TBLGUWURTTIMEPOINTS", maxID)

    End Sub

    Private Sub cmdRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemove.Click

        Dim dgv As DataGridView
        Dim id As Int64
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String

        dgv = Me.dgvTimePoints

        tbl = tblGuWuRTTimePoints

        For Each sr As DataGridViewRow In dgv.SelectedRows

            id = sr.Cells("ID_TBLGUWURTTIMEPOINTS").Value

            strF = "ID_TBLGUWURTTIMEPOINTS = " & id
            Erase rows
            rows = tbl.Select(strF)

            If rows.Length = 0 Then

            Else
                rows(0).Delete()
            End If

        Next



    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click

        Call AddCmpd()

    End Sub

    Sub AddCmpd()

        Dim strMin As String
        Dim strHour As String
        Dim strDay As String
        Dim boolDo As Boolean
        Dim boolMin As Boolean
        Dim boolHour As Boolean
        Dim boolDay As Boolean
        Dim strM As String
        Dim varE
        Dim boolE As Boolean
        Dim numL As Single

        strMin = Me.txtMin.Text
        strHour = Me.txtHour.Text
        strDay = Me.txtDay.Text

        boolDo = False
        boolMin = False
        boolHour = False
        boolDay = False
        boolE = False
        strM = "Entry must be numeric >= 0"

        If Len(strMin) = 0 Then
            If Len(strHour) = 0 Then
                If Len(strDay) = 0 Then
                Else
                    boolDay = True
                    boolDo = True
                End If
            Else
                boolHour = True
                boolDo = True
            End If
        Else
            boolDo = True
            boolMin = True
        End If

        If boolDo Then 'continue
            If boolMin Then
                varE = strMin
            Else
                If boolHour Then
                    varE = strHour
                Else
                    If boolDay Then
                        varE = strDay
                    End If
                End If
            End If
        End If

        If IsNumeric(varE) Then
            If CDec(varE) < 0 Then
                boolE = True
            Else
                numL = 0
                If boolMin Then
                    numL = RoundToDecimalRAFZ(CDec(varE) / 60, 3)
                ElseIf boolHour Then
                    numL = CDec(varE)
                ElseIf boolDay Then
                    numL = RoundToDecimalRAFZ(CDec(varE) * 24, 3)
                End If

                Call AddTimepoint(numL)


            End If
        Else
            boolE = True
        End If

        If boolE Then
            MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
        Else
            'now clear items
            Me.txtMin.Text = ""
            Me.txtHour.Text = ""
            Me.txtDay.Text = ""

        End If



    End Sub

    Sub AddTimepoint(ByVal numL As Single)

        Dim tbl As System.Data.DataTable
        Dim var1
        Dim Count1 As Short
        Dim rows() As DataRow
        Dim maxID As Int64
        Dim dgv As DataGridView
        Dim boolGo As Boolean
        Dim num1 As Single

        tbl = tblGuWuRTTimePoints

        Dim strM As String

        'first ensure time point is unique
        boolGo = True
        dgv = Me.dgvTimePoints
        For Count1 = 0 To dgv.Rows.Count - 1
            num1 = dgv("NUMTIMEPOINT", Count1).Value
            If num1 = numL Then
                strM = "Time point '" & numL & "' already exists."
                MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
                boolGo = False
                Exit For
            End If
        Next

        If boolGo Then
            Dim nRow As DataRow = tbl.NewRow

            nRow.BeginEdit()
            maxID = GetMaxID("TBLGUWURTTIMEPOINTS", 1, True)
            nRow("ID_TBLGUWURTTIMEPOINTS") = maxID
            nRow("ID_TBLGUWUSTUDIES") = idS
            nRow("ID_TBLSTUDIES") = idS1
            nRow("ID_TBLGUWUASSAY") = idA
            nRow("ID_TBLGUWUPKGROUPS") = idG
            nRow("ID_TBLGUWUPKROUTES") = idR
            nRow("NUMTIMEPOINT") = numL
            nRow.EndEdit()

            tbl.Rows.Add(nRow)

        End If




    End Sub


    Private Sub txtMin_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtMin.Validating

        Dim var1
        Dim strM As String

        var1 = Me.txtMin.Text

        If Len(var1) = 0 Then
            Exit Sub
        End If

        If IsNumeric(var1) Then
        Else
            strM = "Entry must be numeric."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
            Me.txtMin.Focus()
            e.Cancel = True
        End If

    End Sub

    Private Sub txtHour_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtHour.Validating

        Dim var1
        Dim strM As String

        var1 = Me.txtHour.Text

        If Len(var1) = 0 Then
            Exit Sub
        End If

        If IsNumeric(var1) Then
        Else
            strM = "Entry must be numeric."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
            Me.txtHour.Focus()
            e.Cancel = True
        End If

    End Sub

    Private Sub txtDay_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDay.Validating

        Dim var1
        Dim strM As String

        var1 = Me.txtDay.Text

        If Len(var1) = 0 Then
            Exit Sub
        End If

        If IsNumeric(var1) Then
        Else
            strM = "Entry must be numeric."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
            Me.txtDay.Focus()
            e.Cancel = True
        End If
    End Sub

    Private Sub cmdSaveSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveSet.Click

        Dim strTitle As String
        Dim strM As String
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim intGo As Short
        Dim strF As String
        Dim boolGo As Boolean
        Dim boolO As Boolean
        Dim id As Int64
        Dim id1 As Int64

        strM = "Enter a name for this Time Point Set"
        strTitle = InputBox(strM, "Enter a Time Point Set name...")
        If Len(strTitle) = 0 Then
            Exit Sub
        End If

        intGo = 2
        'determine if unique
        tbl = TBLGUWUTPNAMESCONFIG
        strF = "CHARTPSETNAME = '" & strTitle & "'"
        rows = tbl.Select(strF)
        boolGo = True
        boolO = False
        If rows.Length = 0 Then 'continue
            boolGo = True
        Else
            strM = "The Time Point Set Name '" & strTitle & "' already exists."
            strM = strM & ChrW(10) & ChrW(10)
            strM = strM & "Do you wish to overwrite (OK) the existing set?"
            strM = strM & ChrW(10) & ChrW(10)
            strM = strM & "Or"
            strM = strM & ChrW(10) & ChrW(10)
            strM = strM & "Do you wish to use a different name (Cancel)?"
            intGo = MsgBox(strM, MsgBoxStyle.OkCancel, "Name must be unique...")
            If intGo = 1 Then 'overwrite
                boolGo = True
                boolO = True
            Else
                boolGo = False
            End If

        End If

        If boolGo Then
            Dim tbl1 As System.Data.DataTable
            Dim rows1() As DataRow
            Dim strF1 As String
            Dim Count1 As Short
            Dim dgvS As DataGridView
            Dim dgvD As DataGridView
            Dim maxID As Int64
            Dim boolM As Boolean
            Dim str1 As String

            tbl1 = TBLGUWUTPCONFIG

            If boolO Then 'overwrite
                'first delete existing rows
                id = rows(0).Item("ID_TBLGUWUTPNAMESCONFIG")
                strF1 = "ID_TBLGUWUTPNAMESCONFIG = " & id
                rows1 = tbl1.Select(strF1)
                For Count1 = 0 To rows1.Length - 1
                    rows1(Count1).Delete()
                Next

            Else 'add new

                'add a row to names table
                Dim nRow1 As DataRow = tbl.NewRow
                maxID = GetMaxID("TBLGUWUTPNAMESCONFIG", 1, True)
                nRow1.BeginEdit()
                nRow1("ID_TBLGUWUTPNAMESCONFIG") = maxID
                nRow1("CHARTPSETNAME") = strTitle
                nRow1("BOOLINCLUDE") = -1
                nRow1.EndEdit()
                tbl.Rows.Add(nRow1)

                id = maxID

            End If

            'now add information from dgvS
            dgvS = Me.dgvTimePoints
            maxID = GetMaxID("TBLGUWUTPCONFIG", dgvS.Rows.Count, True)

            For Count1 = 0 To dgvS.Rows.Count - 1
                Dim nRow As DataRow = tbl1.NewRow
                maxID = maxID + 1
                nRow.BeginEdit()
                nRow("ID_TBLGUWUTPCONFIG") = maxID
                nRow("ID_TBLGUWUTPNAMESCONFIG") = id
                nRow("NUMTIMEPOINT") = dgvS("NUMTIMEPOINT", Count1).Value
                nRow.EndEdit()
                tbl1.Rows.Add(nRow)
            Next
            ''20190219 LEE: Don't need anymore. Used GetMaxID
            'boolM = PutMaxID("TBLGUWUTPCONFIG", maxID)

            'now select row
            dgvD = Me.dgvTP
            For Count1 = 0 To dgvD.Rows.Count - 1
                str1 = dgvD("CHARTPSETNAME", Count1).Value
                If StrComp(str1, strTitle, CompareMethod.Text) = 0 Then
                    dgvD.CurrentCell = dgvD.Rows(Count1).Cells("CHARTPSETNAME")
                    dgvD.CurrentRow.Selected = True
                    Exit For
                End If
            Next

            boolTP = True
            boolTPName = True

        End If

    End Sub

    Private Sub cmdDeactivate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeactivate.Click

        Call Activate(False)

    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click

        Dim dgv As DataGridView
        Dim id As Int64
        Dim intRow As Short
        Dim strM As String
        Dim strF As String
        Dim strF1 As String
        Dim tbl As System.Data.DataTable
        Dim tbl1 As System.Data.DataTable
        Dim rows() As DataRow
        Dim rows1() As DataRow
        Dim Count1 As Short
        Dim intGo As Short
        Dim strN As String

        dgv = Me.dgvTimepointSets

        If dgv.Rows.Count = 0 Then
            Exit Sub
        End If

        If dgv.CurrentRow Is Nothing Then
            strM = "Please select a Set Name."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            GoTo end1
        End If

        intRow = dgv.CurrentRow.Index
        strN = dgv("CHARTPSETNAME", intRow).Value

        strM = "Are you sure you wish to delete Time Point Set Name '" & strN & "'?"
        intGo = MsgBox(strM, MsgBoxStyle.OkCancel, "Delete Time Point Set...")
        If intGo = 1 Then 'continue
        Else
            GoTo end1
        End If

        id = dgv("ID_TBLGUWUTPNAMESCONFIG", intRow).Value

        'first delete tp
        tbl1 = TBLGUWUTPCONFIG
        strF = "ID_TBLGUWUTPNAMESCONFIG = " & id
        rows1 = tbl1.Select(strF)
        For Count1 = 0 To rows1.Length - 1
            rows1(Count1).Delete()
        Next

        'now delete tp name
        tbl = TBLGUWUTPNAMESCONFIG
        strF = "ID_TBLGUWUTPNAMESCONFIG = " & id
        rows = tbl.Select(strF)
        If rows.Length = 0 Then
        Else
            rows(0).Delete()
        End If

        boolTPName = True
        boolTP = True


end1:

    End Sub

    Private Sub cmdActivate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdActivate.Click

        Call Activate(True)

    End Sub

    Sub Activate(ByVal boolI As Boolean)

        Dim dgv As DataGridView
        Dim id As Int64
        Dim intRow As Short
        Dim strM As String
        Dim strF As String
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim intB As Short

        If boolI Then
            intB = -1
        Else
            intB = 0
        End If

        dgv = Me.dgvTimepointSets

        If dgv.Rows.Count = 0 Then
            Exit Sub
        End If

        If dgv.CurrentRow Is Nothing Then
            strM = "Please select a Set Name."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            GoTo end1
        End If

        intRow = dgv.CurrentRow.Index
        id = dgv("ID_TBLGUWUTPNAMESCONFIG", intRow).Value

        tbl = TBLGUWUTPNAMESCONFIG
        strF = "ID_TBLGUWUTPNAMESCONFIG = " & id
        rows = tbl.Select(strF)
        rows(0).BeginEdit()
        rows(0).Item("BOOLINCLUDE") = intB
        rows(0).EndEdit()

        boolTPName = True

end1:

    End Sub
End Class