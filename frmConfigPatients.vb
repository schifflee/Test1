Option Compare Text

Public Class frmConfigPatients

    Public boolCancel As Boolean = True
    Public boolFromGroupSummary As Boolean = False
    Public idS As Int64
    Public idS1 As Int64
    Public idA As Int64
    Public boolFormLoad As Boolean = False


    Private Sub cmdOK1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK1.Click

        boolCancel = False
        Me.Visible = False

    End Sub

    Private Sub cmdCancel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel1.Click

        boolCancel = True
        tblGuWuPKSubjects.RejectChanges()
        Me.Visible = False

    End Sub

    Private Sub rbSerial_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSerial.CheckedChanged

        Call SerialSelect()

    End Sub

    Sub SetSerialSelect()
        Dim strF As String
        Dim dgv As DataGridView
        Dim Count1 As Short
        Dim intRows As Short
        Dim int1 As Short

        dgv = Me.dgvPatients

        If dgv.Rows.Count = 0 Then
        Else
            int1 = dgv("BOOLSERIALBLEED", 0).Value

            If int1 = -1 Then
                Me.rbSerial.Checked = True
            Else
                Me.rbSerialNon.Checked = True
            End If

        End If

    End Sub

    Sub SerialSelect()

        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim idG As Int64
        Dim idR As Int64
        Dim dgv As DataGridView
        Dim Count1 As Short
        Dim intRows As Short
        Dim int1 As Short

        dgv = Me.dgvPatients
        tbl = tblGuWuPKSubjects

        If dgv.Rows.Count = 0 Then
        Else
            idG = dgv("ID_TBLGUWUPKGROUPS", 0).Value
            idR = dgv("ID_TBLGUWUPKROUTES", 0).Value
            strF = "ID_TBLGUWUPKGROUPS = " & idG & " AND ID_TBLGUWUPKROUTES = " & idR
            rows = tbl.Select(strF)
            If Me.rbSerial.Checked Then
                int1 = -1
            Else
                int1 = 0
            End If
            For Count1 = 0 To rows.Length - 1
                rows(Count1).BeginEdit()
                rows(Count1).Item("BOOLSERIALBLEED") = int1
                rows(Count1).EndEdit()
            Next
        End If

        Try

            If Me.rbSerial.Checked Then
                Me.panTimePoints.Visible = False
                Me.dgvPatients.Columns("NUMTIMEPOINT").Visible = False
                Me.dgvPatients.Columns("NUMPATIENTGROUP").Visible = False
            Else
                Me.panTimePoints.Visible = True
                Me.dgvPatients.Columns("NUMTIMEPOINT").Visible = True
                Me.dgvPatients.Columns("NUMPATIENTGROUP").Visible = False
            End If

            Me.dgvPatients.AutoResizeColumns()
        Catch ex As Exception

        End Try

        Call FillPatientsCheck(Me.dgvPatients)


    End Sub

    Private Sub rbSerialNon_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbSerialNon.CheckedChanged

        Call SerialSelect()

    End Sub

    Sub ChangePatients()

        Dim id As Int64
        Dim intRow As Short
        Dim strF As String
        Dim dv As System.Data.DataView
        Dim dgv As DataGridView

        Try
            dv = Me.dgvPatients.DataSource
            Try
                dv.AllowEdit = True
            Catch ex As Exception

            End Try

            dgv = Me.dgvGroupSummary
            If dgv.Rows.Count = 0 Then
                id = -1
            ElseIf dgv.CurrentRow Is Nothing Then
                id = dgv.Item("ID_TBLGUWUPKROUTES", 0).Value
            Else
                intRow = dgv.CurrentRow.Index
                id = dgv.Item("ID_TBLGUWUPKROUTES", intRow).Value
            End If
            strF = "ID_TBLGUWUPKROUTES = " & id
            dv.RowFilter = strF

            Try
                'now fill cmpds
                Call FillPatient(dv)
            Catch ex As Exception

            End Try

            Call DoLabel(Me.dgvPatients, Me.lblPatients, "Patients")


        Catch ex As Exception

        End Try



    End Sub

    Sub ChangeTimePoints()

        Dim id As Int64
        Dim intRow As Short
        Dim strF As String
        Dim dv As System.Data.DataView
        Dim dgv As DataGridView

        Try
            dv = Me.dgvGroupTimePoints.DataSource
            dgv = Me.dgvGroupSummary
            If dgv.Rows.Count = 0 Then
                id = -1
            ElseIf dgv.CurrentRow Is Nothing Then
                id = dgv.Item("ID_TBLGUWUPKROUTES", 0).Value
            Else
                intRow = dgv.CurrentRow.Index
                id = dgv.Item("ID_TBLGUWUPKROUTES", intRow).Value
            End If
            strF = "ID_TBLGUWUPKROUTES = " & id
            dv.RowFilter = strF

            Call DoLabel(Me.dgvGroupTimePoints, Me.lblTimePoints, "Time Points")


        Catch ex As Exception

        End Try


    End Sub

    Sub DoLabel(ByVal dgv As DataGridView, ByVal lbl As Label, ByVal strT As String)

        Dim intC As Int16
        intC = dgv.RowCount
        strT = strT & " - " & intC
        lbl.Text = strT

    End Sub

    Sub FillPatient(ByVal dv As System.Data.DataView)

        Dim tbl As System.Data.DataTable
        Dim tbl1 As System.Data.DataTable
        Dim rows() As DataRow
        Dim rows1() As DataRow
        Dim strF As String
        Dim strF1 As String
        Dim strS As String
        Dim intRows As Int16
        Dim Count1 As Int16
        Dim id As Int64
        Dim id1 As Int64
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strName As String
        Dim arr1(3, 1) As Int64
        Dim num1 As Single

        intRows = dv.Count
        ReDim arr1(3, intRows)

        'tbl = tblGuWuCompounds
        'tbl1 = tblGuWuAssignedCmpd

        tbl = tblGuWuPKSubjects
        tbl1 = tblGuWuRTTimePoints

        Dim var1

        For Count1 = 0 To intRows - 1
            'who knows why I have to do this, but I do
            'arr1(0, Count1) = dv(Count1).Item("ID_TBLGUWUASSIGNEDCMPD")
            'arr1(1, Count1) = dv(Count1).Item("ID_TBLGUWUCOMPOUNDS")

            arr1(0, Count1) = NZ(dv(Count1).Item("ID_tblGuWuRTTimePoints"), -1)
            arr1(1, Count1) = dv(Count1).Item("ID_tblGuWuPKSubjects")
        Next

        For Count1 = 0 To intRows - 1
            id = arr1(0, Count1) 'dv(Count1).Item("ID_tblGuWuRTTimePoints")
            id1 = arr1(1, Count1) 'dv(Count1).Item("ID_tblGuWuPKSubjects")
            strF = "ID_tblGuWuRTTimePoints = " & id
            Erase rows1
            rows1 = tbl1.Select(strF)

            If rows1.Length = 0 Then
                num1 = -1
            Else
                num1 = rows1(0).Item("NUMTIMEPOINT")
            End If


            strF1 = "ID_tblGuWuPKSubjects = " & id1
            Erase rows
            rows = tbl.Select(strF1)

            rows(0).BeginEdit()
            rows(0).Item("NUMTIMEPOINT") = num1
            rows(0).EndEdit()

        Next

        Me.dgvPatients.AutoResizeColumns()


    End Sub

    Private Sub dgvGroupSummary_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvGroupSummary.SelectionChanged

        Dim dgv As DataGridView
        Dim intRow As Short
        Dim id1 As Int64
        Dim id2 As Int64
        Dim intRows As Short
        Dim intT As Short
        Dim int1 As Short
        Dim boolCrap As Boolean

        If boolFromGroupSummary Then
            Exit Sub
        End If

        boolFromGroupSummary = True

        dgv = Me.dgvGroupSummary

        If dgv.Rows.Count = 0 Then
            GoTo end1
        ElseIf dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If

        id1 = dgv("ID_TBLGUWUPKROUTES", intRow).Value

        intRows = dgv.Rows.Count


        If id1 = -1 Then
            id2 = -1
            intT = intRow
            int1 = 0
            boolCrap = False
            Do Until id2 <> -1
                int1 = int1 + 1
                If intT = intRows - 1 And intRows <> 1 Then 'go back
                    intT = intT - 1
                ElseIf intT = 0 And intRows <> 1 Then 'go forward
                    intT = intT + 1
                ElseIf intRows = 1 Then
                    boolCrap = True
                    Exit Do
                Else 'go forward
                    intT = intT + 1
                End If

                id2 = dgv("ID_TBLGUWUPKROUTES", intT).Value

                If int1 > 25 Then
                    boolCrap = True
                    Exit Do
                End If
            Loop

            If boolCrap Then
            Else
                intRow = intT
                dgv.CurrentCell = dgv.Rows(intRow).Cells("ColumnValue")
                dgv.CurrentRow.Selected = True
            End If
        End If

end1:

        Call ChangeTimePoints()

        Call ChangePatients()

        Call SetSerialSelect()

        Call SetSerial()

        boolFromGroupSummary = False

    End Sub

    Sub FormLoad()

        boolFormLoad = True

        Call Configure_dgvTimePoint()

        Call Configure_dgvPatient()

        boolFormLoad = False
        Call SetSerial()
        boolFormLoad = True
        Call SerialSelect()

        boolFormLoad = False

        'Call SetSerial()

        Call BaseName()

        Call Locks(True)

        boolFormLoad = True
        Me.dgvGroupSummary.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        boolFormLoad = False

    End Sub

    Sub SetSerial()

        Dim dgv As DataGridView
        Dim boolS As Boolean
        Dim intRow As Short
        Dim intS As Short

        dgv = Me.dgvPatients
        boolS = True

        If boolFormLoad Then
            Exit Sub
        End If

        If dgv.RowCount = 0 Then
            boolS = True
        Else
            If dgv.CurrentRow Is Nothing Then
                intRow = 0
            Else
                intRow = dgv.CurrentRow.Index
            End If
            intS = -1
            intS = dgv("BOOLSERIALBLEED", intRow).Value
            If intS = -1 Then
                boolS = True
            Else
                boolS = False
            End If
        End If


        Try
            If boolS Then
                Me.rbSerial.Checked = True
                dgv.Columns("NUMTIMEPOINT").Visible = False
            Else
                Me.rbSerialNon.Checked = True
                dgv.Columns("NUMTIMEPOINT").Visible = True
            End If
        Catch ex As Exception

        End Try

    End Sub

    Sub Configure_dgvPatient()
        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim Count1 As Short
        Dim cw

        dgv = Me.dgvPatients


        For Count1 = 1 To 1
            Select Case Count1
                Case 1
                    dgv = Me.dgvPatients
                    cw = 10


            End Select
            dgv.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv.AllowUserToResizeColumns = True
            dgv.AllowUserToResizeRows = True
            dgv.RowHeadersWidth = cw '25
            dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            'dgv.Columns("NUMTIMEPOINT").SortMode = DataGridViewColumnSortMode.NotSortable
            Try
                dgv.Columns("CHARSUBJECTNAME").SortMode = DataGridViewColumnSortMode.NotSortable
            Catch ex As Exception
                Dim str1 As String
                str1 = ""
            End Try



            Select Case Count1
                Case 1
                    'Call ConfigSubjectTableSD(True)

            End Select
        Next

        'dv = dgv.DataSource
        'dv.AllowEdit = True

        ''make some columns checkboxes
        'Dim column As New DataGridViewCheckBoxColumn()
        'With column
        '    .HeaderText = "Serial" & ChrW(10) & "Bleed"
        '    .Name = "BOOLSERIAL"
        '    .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        '    .FlatStyle = FlatStyle.Standard
        '    .CellTemplate = New DataGridViewCheckBoxCell()
        '    '.CellTemplate.Style.BackColor = Color.Beige
        'End With
        'dgv.Columns.Insert(0, column)


        'Dim column1 As New DataGridViewCheckBoxColumn()
        'With column1
        '    .HeaderText = "Terminal" & ChrW(10) & "Bleed"
        '    .Name = "BOOLTERMINAL"
        '    .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        '    .FlatStyle = FlatStyle.Standard
        '    .CellTemplate = New DataGridViewCheckBoxCell()
        '    '.CellTemplate.Style.BackColor = Color.Beige
        'End With
        'dgv.Columns.Insert(0, column1)

end1:

    End Sub


    Sub Configure_dgvTimePoint()
        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim Count1 As Short
        Dim cw

        For Count1 = 1 To 1
            Select Case Count1
                Case 1
                    dgv = Me.dgvGroupTimePoints
                    cw = 10


            End Select
            dgv.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv.AllowUserToResizeColumns = True
            dgv.AllowUserToResizeRows = True
            dgv.RowHeadersWidth = cw '25
            dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            'dgv.Columns("NUMTIMEPOINT").SortMode = DataGridViewColumnSortMode.NotSortable
            Try
                dgv.Columns("NUMTIMEPOINT").SortMode = DataGridViewColumnSortMode.NotSortable
            Catch ex As Exception
                Dim str1 As String
                str1 = ""
            End Try

            Select Case Count1
                Case 1
                    'Call ConfigTimePointTableSD(True)

            End Select
        Next



end1:

    End Sub

    Private Sub rbIncrement_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbIncrement.CheckedChanged

        Call BaseName()

    End Sub

    Sub BaseName()
        Dim str1 As String

        Try
            If Me.rbIncrement.Checked Then
                Me.txtBaseName.Enabled = True
                Me.lbltxtUniqueID.Enabled = False
                Me.dgvPatients.Columns("CHARSUBJECTNAME").Visible = True
                Me.dgvPatients.Columns("CHARUNIQUEID").Visible = False
            ElseIf Me.rbUnique.Checked Then
                Me.txtBaseName.Enabled = False
                Me.lbltxtUniqueID.Enabled = True
                Me.dgvPatients.Columns("CHARSUBJECTNAME").Visible = False
                Me.dgvPatients.Columns("CHARUNIQUEID").Visible = True
            ElseIf Me.rbBothNames.Checked Then
                Me.txtBaseName.Enabled = True
                Me.lbltxtUniqueID.Enabled = True
                Me.dgvPatients.Columns("CHARSUBJECTNAME").Visible = True
                Me.dgvPatients.Columns("CHARUNIQUEID").Visible = True
            End If

            'Me.dgvPatients.ReadOnly = False
            Me.dgvPatients.Columns("CHARUNIQUEID").ReadOnly = False
            Me.dgvPatients.Columns("CHARSUBJECTNAME").ReadOnly = True

        Catch ex As Exception

        End Try

        Me.txtBaseName.Text = ""
        'Me.txtUniqueName.Text = ""

    End Sub

    Private Sub rbBothNames_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbBothNames.CheckedChanged

        Call BaseName()

    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click

        Call AddBaseName()

    End Sub

    Sub AddBaseName()

        Dim dgv As DataGridView
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim strS As String
        Dim strBaseName As String
        Dim strM As String
        Dim maxID As Int64
        Dim idG As Int64
        Dim idR As Int64
        Dim intRow As Short

        strBaseName = Me.txtBaseName.Text

        If Me.rbIncrement.Checked Or Me.rbBothNames.Checked Then 'base name must have a value

            If Len(strBaseName) = 0 Then
                strM = "Base Name cannot be blank"
                MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
                GoTo end1
            End If

        ElseIf Me.rbUnique.Checked Then 'add base name row


        End If

        dgv = Me.dgvGroupSummary
        If dgv.Rows.Count = 0 Then
            GoTo end1
        End If

        If dgv.CurrentRow Is Nothing Then
            strM = "Please select a Route."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            GoTo end1
        End If

        intRow = dgv.CurrentRow.Index

        idG = dgv("ID_TBLGUWUPKGROUPS", intRow).Value
        idR = dgv("ID_TBLGUWUPKROUTES", intRow).Value

        If idR = -1 Then
            strM = "Please select a Route."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            GoTo end1
        End If

        tbl = tblGuWuPKSubjects
        Dim nrow As DataRow = tbl.NewRow
        nrow.BeginEdit()
        nrow("ID_TBLGUWUPKSUBJECTS") = GetMaxID("TBLGUWUPKSUBJECTS", 1, True)
        nrow("ID_TBLGUWUSTUDIES") = idS
        nrow("ID_TBLSTUDIES") = idS1
        nrow("ID_TBLGUWUASSAY") = idA
        nrow("ID_TBLGUWUPKGROUPS") = idG
        nrow("ID_TBLGUWUPKROUTES") = idR

        nrow("CHARUNIQUEID") = DBNull.Value
        nrow("BOOLTERMINALBLEED") = 0

        If Me.rbSerial.Checked Then
            nrow("BOOLSERIALBLEED") = -1
            nrow("ID_TBLGUWURTTIMEPOINTS") = -1
            nrow("NUMPATIENTGROUP") = -1
        Else
            nrow("BOOLSERIALBLEED") = 0
            nrow("ID_TBLGUWURTTIMEPOINTS") = DBNull.Value
            nrow("NUMPATIENTGROUP") = 1
        End If

        'determine subject name
        If Me.rbIncrement.Checked Or Me.rbBothNames.Checked Then 'base name must have a value
            nrow("CHARSUBJECTNAME") = Me.txtBaseName.Text & "-" & GetIncr(Me.txtBaseName.Text)
        ElseIf Me.rbUnique.Checked Then 'add base name row
            nrow("CHARSUBJECTNAME") = DBNull.Value
        End If

        nrow.EndEdit()
        tbl.Rows.Add(nrow)

end1:


    End Sub

    Sub Locks(ByVal bool As Boolean)

        Me.gbxBleeds.Enabled = bool
        Me.gbxIncrement.Enabled = bool
        Me.gbxNameType.Enabled = bool

        Me.dgvPatients.ReadOnly = bool

        Me.cmdAdd.Enabled = Not (bool)
        Me.cmdRemove.Enabled = Not (bool)
        Me.cmdAddTP.Enabled = Not (bool)


    End Sub

    Function GetIncr(ByVal txtBN As String) As Short

        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim strS As String
        Dim arr1(1)
        Dim intRows As Short
        Dim Count1 As Short
        Dim var1, var2, var3
        Dim maxI As Short

        tbl = tblGuWuPKSubjects
        strF = "ID_TBLGUWUASSAY = " & idA & " AND CHARSUBJECTNAME LIKE '" & txtBN & "-*'"
        strS = "CHARSUBJECTNAME ASC"
        rows = tbl.Select(strF, strS)
        intRows = rows.Length
        If intRows = 0 Then
            GetIncr = 1
            GoTo end1
        End If

        ReDim arr1(intRows)
        'find incr
        maxI = 0
        For Count1 = 0 To intRows - 1
            var1 = rows(Count1).Item("CHARSUBJECTNAME")
            var2 = Mid(var1, Len(txtBN) + 2)
            If Len(var2) = 0 Then
                GetIncr = 1
                GoTo end1
            Else
                If IsNumeric(var2) Then
                    If var2 > maxI Then
                        maxI = var2
                    End If
                End If
            End If
        Next

        GetIncr = maxI + 1


end1:

    End Function


    Private Sub dgvPatients_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvPatients.CellValidating


        Dim dgv As DataGridView
        Dim intCol As Short
        Dim strCol As String
        Dim var1
        Dim intRow As Short
        Dim boolE As Boolean
        Dim strM As String

        dgv = Me.dgvPatients
        intCol = e.ColumnIndex
        intRow = e.RowIndex

        strCol = dgv.Columns(intCol).Name
        If StrComp(strCol, "NUMPATIENTGROUP", CompareMethod.Text) = 0 Then 'continue
        Else
            GoTo end1
        End If

        'entry must be integer
        boolE = False
        var1 = dgv(intCol, intRow).Value
        If Len(NZ(var1, "")) = 0 Then
        Else
            If IsNumeric(var1) Then
                If var1 < 1 Then
                    boolE = True
                Else
                    If IsInt(var1) Then
                    Else
                        boolE = True
                    End If
                End If
            End If
        End If

        If boolE Then
            strM = "Entry must be integer > 0"
            MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
            e.Cancel = True
            GoTo end1
        End If

end1:

    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click

        Call DoThis("cmdEdit")


    End Sub

    Sub DoThis(ByVal cmd As String)
        Dim strM As String
        Dim intM As Short

        Select Case cmd
            Case "cmdEdit"
                Call Locks(False)

                Me.cmdEdit.Enabled = False
                Me.cmdSave.Enabled = True
                Me.cmdCancel.Enabled = True
                Me.cmdExit.Enabled = False

            Case "cmdSave"
                Call Locks(True)

                Me.cmdEdit.Enabled = True
                Me.cmdSave.Enabled = False
                Me.cmdCancel.Enabled = False
                Me.cmdExit.Enabled = True

            Case "cmdCancel"
                strM = "NOTE: This will cancel ALL edits made during this Configure Patients session."
                strM = strM & ChrW(10) & ChrW(10)
                strM = strM & "Do you wish to continue?"
                'intM = MsgBox(strM, MsgBoxStyle.YesNo, "Do you wish to continue?")
                intM = 6
                If intM = 6 Then 'yes
                Else
                    GoTo end1
                End If
                Call Locks(True)
                tblGuWuPKSubjects.RejectChanges()

                Me.cmdEdit.Enabled = True
                Me.cmdSave.Enabled = False
                Me.cmdCancel.Enabled = False
                Me.cmdExit.Enabled = True

                Call SetSerialSelect()

        End Select

end1:

    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        Call DoThis("cmdSave")


    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Call DoThis("cmdCancel")

    End Sub

    Private Sub cmdRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemove.Click

        Dim dgv As DataGridView
        Dim id As Int64
        Dim intRow As Short
        Dim strM As String
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String

        dgv = Me.dgvPatients

        If dgv.Rows.Count = 0 Then
            GoTo end1
        End If

        If dgv.CurrentRow Is Nothing Then
            strM = "Please select a Patient row."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            GoTo end1
        End If

        intRow = dgv.CurrentRow.Index
        id = dgv("ID_TBLGUWUPKSUBJECTS", intRow).Value
        strF = "ID_TBLGUWUPKSUBJECTS = " & id
        tbl = tblGuWuPKSubjects
        rows = tbl.Select(strF)
        rows(0).Delete()

end1:

    End Sub

    Sub AddTP()

        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim intRow1 As Short
        Dim intRow2 As Short
        Dim id1 As Int64
        Dim numTP As Single
        Dim strM As String

        dgv1 = Me.dgvPatients
        dgv2 = Me.dgvGroupTimePoints

        If dgv1.Rows.Count = 0 Then
            GoTo end1
        End If

        If dgv2.Rows.Count = 0 Then
            GoTo end1
        End If

        If dgv1.CurrentRow Is Nothing Then
            strM = "Please select a Patient row."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            GoTo end1
        End If

        If dgv2.CurrentRow Is Nothing Then
            strM = "Please select a Time Point row."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            GoTo end1
        End If

        intRow1 = dgv1.CurrentRow.Index
        intRow2 = dgv2.CurrentRow.Index

        dgv1("NUMTIMEPOINT", intRow1).Value = dgv2("NUMTIMEPOINT", intRow2).Value
        dgv1("ID_TBLGUWURTTIMEPOINTS", intRow1).Value = dgv2("ID_TBLGUWURTTIMEPOINTS", intRow1).Value
        dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit)

end1:

    End Sub

    Private Sub cmdAddTP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddTP.Click

        Call AddTP()

    End Sub

    Private Sub dgvPatients_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPatients.CellValueChanged

        Dim dgv As DataGridView
        Dim intCol As Short
        Dim strCol As String
        Dim bool As Boolean
        Dim dv As System.Data.DataView
        Dim boolA As Boolean

        dgv = Me.dgvPatients
        intCol = e.ColumnIndex
        strCol = dgv.Columns(intCol).Name
        If StrComp(strCol, "BOOLTERMINAL", CompareMethod.Text) = 0 Then 'continue
        Else
            GoTo end1
        End If

        If e.RowIndex = -1 Then
            GoTo end1
        End If

        dv = dgv.DataSource
        boolA = dv.AllowEdit
        dv.AllowEdit = True
        bool = dgv(e.ColumnIndex, e.RowIndex).Value

        dv(e.RowIndex).BeginEdit()
        If bool Then
            dv(e.RowIndex).Item("BOOLTERMINALBLEED") = -1
        Else
            dv(e.RowIndex).Item("BOOLTERMINALBLEED") = 0
        End If
        dv(e.RowIndex).EndEdit()
        dv.AllowEdit = boolA

end1:


    End Sub

    Private Sub dgvPatients_CurrentCellDirtyStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvPatients.CurrentCellDirtyStateChanged

        If Me.dgvPatients.IsCurrentCellDirty Then
            dgvPatients.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If

    End Sub

    Private Sub dgvGroupTimePoints_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvGroupTimePoints.CellDoubleClick

        Call AddTP()

    End Sub

    Private Sub cmdExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdExit.Click

        boolCancel = False
        Me.Visible = False

    End Sub
End Class