Public Class frmApplyTemplateTables

    Public boolCancel As Boolean = True
    Public gTblStudies As Int64 = 0
    Public boolFormLoad As Boolean = False

    Private Sub frmApplyTemplateTables_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call DoubleBufferControl(Me, "dgv")

        Call ControlDefaults(Me)

        boolFormLoad = True

        Call StartupPosition()

        Call LoadExistingTables()

        Call LoadTemplates()

        boolFormLoad = False

    End Sub


    Sub Import()

        Dim intR As Short
        Dim strM As String

        Dim dv As DataView
        Dim Count1 As Int32
        Dim Count2 As Int32
        Dim Count3 As Int32
        Dim int1 As Int32
        Dim int2 As Int32
        Dim var1, var2, var3, var4, var5, var6
        Dim id1 As Int64
        Dim id2 As Int64
        Dim id3 As Int64
        Dim strF As String
        Dim dtbl1 As DataTable
        Dim dtbl2 As New DataTable
        Dim maxID As Int64
        Dim omaxID As Int64
        Dim rows() As DataRow
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim boolAdd As Boolean
        Dim strF1 As String
        Dim strF2 As String


        Dim intOrder As Int16 = 999

        Dim idTRT1 As Int64
        Dim idTRT2 As Int64

        dtbl1 = tblReportTables
        int1 = dtbl1.Columns.Count 'debug

        'debug
        dv = frmH.dgvReportTableConfiguration.DataSource
        Dim tblAA As DataTable = dv.ToTable
        For Count1 = 0 To tblAA.Columns.Count - 1
            str1 = tblAA.Columns(Count1).ColumnName
            str1 = str1
        Next

        'first make all boolinclude = 0
        For Count1 = 0 To dtbl1.Rows.Count - 1
            dtbl1.Rows(Count1).BeginEdit()
            dtbl1.Rows(Count1).Item("BOOLINCLUDE") = 0
            dtbl1.Rows(Count1).EndEdit()
        Next

        dv = Me.dgvD.DataSource
        Dim tblD As DataTable = dv.ToTable
        rows = tblD.Select("", "", DataViewRowState.CurrentRows)

        For Count1 = 0 To rows.Length - 1

            intOrder = intOrder + 1
            boolAdd = False

            'check to see if table already exists and simply needs to be made visible
            var1 = rows(Count1).Item("ID_TBLCONFIGREPORTTABLES")
            var2 = rows(Count1).Item("CHARHEADINGTEXT")
            var3 = NZ(rows(Count1).Item("CHARFCID"), "")
            idTRT1 = rows(Count1).Item("ID_TBLREPORTTABLE")

            'first look for id and included
            strF = "ID_TBLCONFIGREPORTTABLES = " & var1 & " AND CHARFCID = '" & var3 & "'"
            Dim rowsA() As DataRow = dtbl1.Select(strF, "", DataViewRowState.CurrentRows)

            If rowsA.Length = 0 Then
                'add to table
                Dim nr1 As DataRow = dtbl1.NewRow
                nr1.BeginEdit()
                For Count2 = 0 To Me.dgvD.Columns.Count - 1
                    str1 = dgvD.Columns(Count2).Name
                    If dtbl1.Columns.Contains(str1) Then
                        nr1(str1) = rows(Count1).Item(str1)
                    End If
                Next

                'next column is CHARTABLENAME
                int1 = dgvD.Columns.Count
                var4 = nr1("ID_TBLCONFIGREPORTTABLES")
                Dim rowsC() As DataRow = tblConfigReportTables.Select("ID_TBLCONFIGREPORTTABLES = " & var4, "", DataViewRowState.CurrentRows)
                str1 = NZ(rowsC(0).Item("CHARTABLENAME"), "NA")
                nr1(int1) = str1

                'last columns are cmpds, make included
                int1 = int1 + 1
                For Count2 = int1 To dtbl1.Columns.Count - 1
                    nr1(Count2) = -1
                Next

                nr1("INTORDER") = intOrder
                nr1("ID_TBLSTUDIES") = id_tblStudies

                'make boolinclude true
                Try
                    nr1("BOOLINCLUDE") = -1
                Catch ex As Exception
                    var1 = ex.Message

                End Try

                nr1.EndEdit()
                dtbl1.Rows.Add(nr1)

                boolAdd = True

            Else
                'simply modify table entry
                rowsA(0).BeginEdit()

                For Count2 = 0 To Me.dgvD.Columns.Count - 1
                    str1 = dgvD.Columns(Count2).Name
                    If dtbl1.Columns.Contains(str1) Then
                        Select Case str1
                            Case "ID_TBLTABLEPROPERTIES"
                            Case "ID_TBLREPORTTABLE"
                            Case "ID_TBLCONFIGREPORTTABLES"
                            Case "ID_TBLSTUDIES"
                            Case Else
                                var1 = rows(Count1).Item(str1)
                                rowsA(0).Item(str1) = rows(Count1).Item(str1)
                        End Select
                    Else
                        var1 = var1
                    End If
                Next

                idTRT2 = rowsA(0).Item("ID_TBLREPORTTABLE")

                rowsA(0).Item("BOOLINCLUDE") = -1
                rowsA(0)("INTORDER") = intOrder

                rowsA(0).EndEdit()

                boolAdd = False

            End If

        Next




        'update TBLTABLEPROPERTIES
        'legend:
        'dtbl2 = tblTableProperties
        'dtbl2 = ta_tblTablePropertiesAcc.GetDataBy_ID_TBLSTUDIES(id_tblStudies)

        maxID = GetMaxID("TBLTABLEPROPERTIES", 1, False)
        omaxID = maxID

        id1 = rows(0).Item("ID_TBLSTUDIES")

        dtbl2.Clear()
        dtbl2.AcceptChanges()

        If boolGuWuOracle Then
            'dtbl2 = ta_tblTableProperties.GetDataBy_ID_TBLSTUDIES(id1)
        ElseIf boolGuWuAccess Then
            dtbl2 = ta_tblTablePropertiesAcc.GetDataBy_ID_TBLSTUDIES(id1)
        ElseIf boolGuWuSQLServer Then
            dtbl2 = ta_tblTablePropertiesSQLServer.GetDataBy_ID_TBLSTUDIES(id1)
        End If


        Dim intAdded As Short
        For Count1 = 0 To rows.Length - 1

            intAdded = 0

            'id1 = tblProp.Rows(Count1).Item("oID_tblStudy")
            'id2 = tblProp.Rows(Count1).Item("oID_tblReportTable")
            'id3 = tblProp.Rows(Count1).Item("nID_tblReportTable")
            'intAdded = tblProp.Rows(Count1).Item("BOOLADDED")

            var1 = rows(Count1).Item("ID_TBLCONFIGREPORTTABLES")
            var2 = rows(Count1).Item("CHARHEADINGTEXT")
            var3 = NZ(rows(Count1).Item("CHARFCID"), "")
            idTRT1 = rows(Count1).Item("ID_TBLREPORTTABLE")
            idTRT2 = idTRT1

            'check to see if table already exists is original data
            'legend
            'dtbl1 = tblReportTables

            'must also look at fcid
            strF = "ID_TBLCONFIGREPORTTABLES = " & var1 & " AND CHARFCID = '" & var3 & "'"
            Dim rowsA() As DataRow = dtbl1.Select(strF, "", DataViewRowState.CurrentRows)
            int1 = rowsA.Length 'debug

            id2 = idTRT1
            id3 = id2

            'get data from dtbl2
            strF1 = "ID_TBLREPORTTABLE = " & idTRT1
            Dim rowsTP() As DataRow = dtbl2.Select(strF1)

            If rowsA.Length = 0 Then

                maxID = maxID + 1
                'add to table
                Dim nr1 As DataRow = tblTableProperties.NewRow
                nr1.BeginEdit()
                For Count2 = 0 To dtbl2.Columns.Count - 1
                    str1 = dtbl2.Columns(Count2).ColumnName
                    nr1(str1) = rowsTP(0).Item(str1)
                Next

                nr1("ID_TBLTABLEPROPERTIES") = maxID
                nr1("ID_TBLSTUDIES") = id_tblStudies

                nr1.EndEdit()

                tblTableProperties.Rows.Add(nr1)

                boolAdd = True

                intAdded = -1

                id2 = nr1.Item("ID_TBLREPORTTABLE")

            Else

                strF2 = "ID_TBLREPORTTABLE = " & rowsA(0).Item("ID_TBLREPORTTABLE")
                Dim rowsTP1() As DataRow = tblTableProperties.Select(strF2, "", DataViewRowState.CurrentRows)
                int1 = rowsTP1.Length 'debug\

                If rowsTP1.Length = 0 Then

                    'make new row
                    maxID = maxID + 1
                    'add to table
                    Dim nr1 As DataRow = tblTableProperties.NewRow
                    nr1.BeginEdit()
                    For Count2 = 0 To dtbl2.Columns.Count - 1
                        str1 = dtbl2.Columns(Count2).ColumnName
                        nr1(str1) = rowsTP(0).Item(str1)
                    Next

                    nr1("ID_TBLTABLEPROPERTIES") = maxID
                    nr1("ID_TBLSTUDIES") = id_tblStudies

                    nr1.EndEdit()

                    tblTableProperties.Rows.Add(nr1)

                    boolAdd = True

                    intAdded = -1

                    id2 = nr1.Item("ID_TBLREPORTTABLE")

                Else
                    'simply modify table entry
                    'legend
                    'dtbl1 = tblReportTables
                    rowsTP1(0).BeginEdit()

                    For Count2 = 0 To dtbl2.Columns.Count - 1
                        str1 = dtbl2.Columns(Count2).ColumnName
                        Select Case str1
                            Case "ID_TBLTABLEPROPERTIES"
                            Case "ID_TBLREPORTTABLE"
                            Case "ID_TBLCONFIGREPORTTABLES"
                            Case "ID_TBLSTUDIES"
                            Case Else
                                var1 = rowsTP(0).Item(str1) 'debug
                                rowsTP1(0).Item(str1) = rowsTP(0).Item(str1)
                        End Select

                    Next

                    idTRT2 = rowsTP1(0).Item("ID_TBLREPORTTABLE")

                    rowsTP1(0).EndEdit()

                    boolAdd = False

                    intAdded = 0
                End If

             

            End If

            'CheckForAutoAssignSamplesTable(ByVal idO As Int64, ByVal idN As Int64, ByVal idTStudiesO As Int64, ByVal idTStudiesN As Int64, ByVal boolGetNew As Boolean, ByVal intAdded As Short)
            Call CheckForAutoAssignSamplesTable(id2, id3, gTblStudies, id_tblStudies, True, intAdded)

        Next

        If maxID = omaxID Then
        Else
            Call PutMaxID("TBLTABLEPROPERTIES", maxID)
        End If

        'Pause(1)

        'Call CheckForTblProperties(-1, 0)

end1:

        Cursor.Current = Cursors.Default

        Me.Visible = False

    End Sub

    Sub ConfigReportTable()

        Dim Count1 As Int32
        Dim dgv As DataGridView = Me.dgvD

        Dim strF As String
        Dim strS As String

        Dim intF As Short
        Dim int1 As Int16

        Dim dvTableRows As DataView

        Dim var1

        Dim tbl1 As DataTable
        If boolGuWuOracle Then
            'tbl1 = ta_tblReportTable.GetDataBy_ID_TBLSTUDIES(gTblStudies)
        ElseIf boolGuWuAccess Then
            tbl1 = ta_tblReportTableAcc.GetDataBy_ID_TBLSTUDIES(gTblStudies)
        ElseIf boolGuWuSQLServer Then
            tbl1 = ta_tblReportTableSQLServer.GetDataBy_ID_TBLSTUDIES(gTblStudies)
        End If
        int1 = tblReportTable.Rows.Count 'DEBUG

        strF = "ID_TBLSTUDIES = " & gTblStudies & " AND BOOLINCLUDE = -1"
        strS = "INTORDER ASC"

        'dvTableRows = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
        dvTableRows = New DataView(tbl1)

        dvTableRows.AllowDelete = False
        dvTableRows.AllowEdit = False
        dvTableRows.AllowNew = False

        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv.DataSource = dvTableRows

        Try
            dvTableRows.RowFilter = strF
        Catch ex As Exception
            var1 = ex.Message
        End Try
        dvTableRows.Sort = strS

        If boolFormLoad Then

            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            'hide all columns
            For Count1 = 0 To dgv.ColumnCount - 1
                dgv.Columns(Count1).Visible = False
            Next
            'make two columns visible
            'dgv.Columns("ID_TBLREPORTTABLE").Visible = True'DEBUG

            'dgv.Columns("ID_TBLCONFIGREPORTTABLES").Visible = True'DEBUG

            dgv.Columns("CHARHEADINGTEXT").Visible = True
            dgv.Columns("CHARFCID").Visible = True
            'dgv.Columns("BOOLINCLUDE").Visible = True

            dgv.Columns("CHARHEADINGTEXT").HeaderText = "Table Title"
            dgv.Columns("CHARFCID").HeaderText = "FC ID"
            dgv.Columns("BOOLINCLUDE").HeaderText = "Incl."

            dgv.Columns("BOOLINCLUDE").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            dgv.Columns("CHARFCID").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

            dgv.Columns("CHARHEADINGTEXT").ReadOnly = True
            dgv.Columns("CHARFCID").ReadOnly = True
            dgv.Columns("BOOLINCLUDE").ReadOnly = True

            'dgv.Columns("CHARHEADINGTEXT").MinimumWidth = dgv.Width * 0.75
            'dgv.Columns("CHARHEADINGTEXT").Width = dgv.Width * 0.75

            ''make boolInclude a checkbox
            'Try
            '    Dim chk As New DataGridViewCheckBoxColumn()
            '    dgv.Columns.Add(chk)
            '    chk.HeaderText = "Incl."
            '    chk.Name = "chk"
            '    chk.DisplayIndex = 0
            '    chk.TrueValue = -1
            '    chk.FalseValue = 0
            '    chk.ThreeState = False
            '    chk.ReadOnly = True
            'Catch ex As Exception
            '    var1 = ex.Message
            'End Try

            dgv.RowHeadersWidth = dgv.RowHeadersWidth * 0.5

            dgv.ColumnHeadersDefaultCellStyle.Font = New Font(dgv.Font, FontStyle.Bold)

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        End If


        dgv.AutoResizeRows()
        dgv.AutoResizeColumns()

        ''fill chk
        'Try
        '    For Count1 = 0 To dgv.Rows.Count - 1
        '        dgv("chk", Count1).Value = NZ(dgv("boolInclude", Count1).Value, 0)
        '        'int1 = NZ(dgv("boolInclude", Count1).Value, 0)
        '        'If int1 = 0 Then
        '        '    dgv("chk", Count1).Value = False
        '        'Else
        '        '    dgv("chk", Count1).Value = True
        '        'End If
        '        ''verify
        '        'var1 = dgv("chk", Count1).Value
        '        'var1 = var1

        '    Next
        'Catch ex As Exception
        '    var1 = ex.Message
        'End Try

    End Sub

    Sub LoadExistingTables()


        Dim Count1 As Int32
        Dim dgv As DataGridView = Me.dgvS
        Dim var1

        dgv.DataSource = frmH.dgvReportTableConfiguration.DataSource

        'hide all columns
        For Count1 = 0 To dgv.ColumnCount - 1
            dgv.Columns(Count1).Visible = False
        Next
        'make two columns visible
        'dgv.Columns("ID_TBLREPORTTABLE").Visible = True'DEBUG

        'dgv.Columns("ID_TBLCONFIGREPORTTABLES").Visible = True'DEBUG

        dgv.Columns("CHARHEADINGTEXT").Visible = True
        dgv.Columns("CHARFCID").Visible = True
        'dgv.Columns("BOOLINCLUDE").Visible = True

        dgv.Columns("CHARHEADINGTEXT").HeaderText = "Table Title"
        dgv.Columns("CHARFCID").HeaderText = "FC ID"
        dgv.Columns("BOOLINCLUDE").HeaderText = "Incl."

        dgv.Columns("BOOLINCLUDE").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        dgv.Columns("CHARFCID").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        dgv.Columns("CHARHEADINGTEXT").ReadOnly = True
        dgv.Columns("CHARFCID").ReadOnly = True
        dgv.Columns("BOOLINCLUDE").ReadOnly = True

        dgv.RowHeadersWidth = dgv.RowHeadersWidth * 0.5
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        dgv.ColumnHeadersDefaultCellStyle.Font = New Font(dgv.Font, FontStyle.Bold)

        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

        dgv.AutoResizeRows()
        dgv.AutoResizeColumns()

        'dgv.Columns("CHARHEADINGTEXT").MinimumWidth = dgv.Width * 0.75
        'dgv.Columns("CHARHEADINGTEXT").Width = dgv.Width * 0.75

        ''select first row
        'DataGridView1.CurrentCell = DataGridView1.Rows(1).Cells(0)
        'dgv.Rows(index).Selected = True


end1:

    End Sub

    Sub LoadTemplates()

        Dim dgv As DataGridView = Me.dgvT
        Dim strF As String
        Dim strS As String
        Dim str1 As String = "CHARTEMPLATENAME"

        strF = "BOOLACTIVE = -1"
        strS = str1 & " ASC"

        Dim dv As DataView = New DataView(tblTemplates, strF, strS, DataViewRowState.CurrentRows)

        dv.AllowNew = False
        dv.AllowEdit = False
        dv.AllowDelete = False

        dgv.DataSource = dv

        Dim Count1 As Int16

        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns(Count1).Visible = False
        Next

        dgv.Columns(str1).HeaderText = "Study Template"
        dgv.Columns(str1).Visible = True

        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font(dgv.DefaultCellStyle.Font, FontStyle.Bold)

        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        dgv.RowHeadersWidth = dgv.RowHeadersWidth * 0.5

    End Sub

    Sub StartupPosition()

        Dim w, h

        'w = My.Computer.Screen.WorkingArea.Width
        'h = My.Computer.Screen.WorkingArea.Height

        w = SystemInformation.WorkingArea.Width
        h = SystemInformation.WorkingArea.Height

        Dim l1, w1

        l1 = w * 0.05
        w1 = w * 0.9

        Me.Left = l1
        Me.Width = w1

        Me.Top = h * 0.05
        Me.Height = h * 0.9

    End Sub

    Private Sub cmdExitSave_Click(sender As Object, e As EventArgs) Handles cmdExitSave.Click

        Dim intR As Short
        Dim strM As String

        strM = "Do you wish to continue?"
        intR = MsgBox(strM, vbOKCancel, "Continue?...")
        If intR = 1 Then
        Else
            Exit Sub
        End If

        Cursor.Current = Cursors.WaitCursor

        Call Import()
        Cursor.Current = Cursors.WaitCursor

        Call UpdateReportStatements()

        'pesky
        Call OrderDGV(frmH.dgvReportTableConfiguration, "INTORDER", "ID_TBLREPORTTABLE")
        Cursor.Current = Cursors.WaitCursor
        Call OrderReportTableConfig()
        Cursor.Current = Cursors.WaitCursor
        Call SetComboCell(frmH.dgvReportTableConfiguration, "CHARPAGEORIENTATION")
        Cursor.Current = Cursors.WaitCursor
        Call AssessSampleAssignment()

        boolCancel = False

        Cursor.Current = Cursors.Default

        Me.Visible = False

    End Sub

    Sub UpdateReportStatements()

        Dim strF As String
        strF = "ID_TBLSTUDIES = " & gTblStudies & " AND ID_TBLCONFIGREPORTTYPE = 100"
        Dim rows() As DataRow = tblReportStatements.Select(strF)

        Dim var1, var2

        var1 = rows(0).Item("CHARSTATEMENT")
        var2 = rows(0).Item("ID_TBLWORDSTATEMENTS")

        Dim dgv1 As DataGridView = frmH.dgvReportStatements
        Dim dv As DataView = dgv1.DataSource
        dv(0).BeginEdit()
        dv(0).Item("CHARSTATEMENT") = var1
        dv(0).Item("ID_TBLWORDSTATEMENTS") = var2
        dv(0).EndEdit()


    End Sub

    Private Sub cmdExitNoSave_Click(sender As Object, e As EventArgs) Handles cmdExitNoSave.Click

        boolCancel = True

        Me.Visible = False

    End Sub


    Private Sub dgvT_SelectionChanged(sender As Object, e As EventArgs) Handles dgvT.SelectionChanged

        Call GetStudyID()

        Call ConfigReportTable()

    End Sub

    Sub GetStudyID()

        Dim dgv As DataGridView = dgvT
        Dim id1 As Int64

        Dim intRow As Int16 = dgv.CurrentRow.Index

        gTblStudies = dgv("ID_TBLSTUDIES", intRow).Value

    End Sub

End Class