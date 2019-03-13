Option Compare Text

Public Class frmImportTables

    Public gTblStudies As Int64 = 0
    Public ttTables As New DataTable
    Public gTableRows() As DataRow
    Public dvTableRows As DataView
    Public boolFormLoad As Boolean = False
    Public dvAdded As DataView
    Public boolSave As Boolean = False
    Public tblProp As New DataTable
    Public boolT As Boolean = True
    Public arrC(2, 1000)
    '1=ID_TBLREPORTTABLE, 2=color

    Sub DoLabels()

        Dim str1 As String

        If boolT Then
            str1 = "Choose a Report Template:"
        Else
            str1 = "Choose a Report Template:"
        End If

        Me.lblStudy.Text = str1

    End Sub

    Private Sub frmImportTables_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load


        Call DoubleBufferControl(Me, "dgv")

        Call ControlDefaults(Me)

        Call StartupPosition()

        boolFormLoad = True

        Call CreateTblProp()

        Try
            dvTableRows.AllowDelete = False
            dvTableRows.AllowEdit = False
            dvTableRows.AllowNew = False
        Catch ex As Exception

        End Try

        Call LoadStudies()

        Try
            Call LoadAdded()
        Catch ex As Exception

        End Try


        Call ConfigReportTable()

        boolFormLoad = False

        Me.cbxStudy.SelectedIndex = 0

        Call DoLabels()


    End Sub

    Sub CreateTblProp()

        Dim col1 As New DataColumn
        col1.ColumnName = "oID_tblStudy"
        col1.Caption = "Orig tblStudy"
        col1.DataType = System.Type.GetType("System.Int64")
        tblProp.Columns.Add(col1)

        Dim col2 As New DataColumn
        col2.ColumnName = "oID_tblReportTable"
        col2.Caption = "Orig ID_tblReportTable"
        col2.DataType = System.Type.GetType("System.Int64")
        tblProp.Columns.Add(col2)

        Dim col3 As New DataColumn
        col3.ColumnName = "nID_tblReportTable"
        col3.Caption = "New ID_tblReportTable"
        col3.DataType = System.Type.GetType("System.Int64")
        tblProp.Columns.Add(col3)

        Dim col4 As New DataColumn
        col4.ColumnName = "BOOLADDED"
        col4.Caption = "Added"
        col4.DataType = System.Type.GetType("System.Int16")
        tblProp.Columns.Add(col4)

        'Dim col2 As New DataColumn
        'col2.ColumnName = "LevelNumber"
        'col2.Caption = "Level"
        'col2.DataType = System.Type.GetType("System.Int16")
        'tblProp.Columns.Add(col2)

        'Dim col2 As New DataColumn
        'col2.ColumnName = "LevelNumber"
        'col2.Caption = "Level"
        'col2.DataType = System.Type.GetType("System.Int16")
        'tblProp.Columns.Add(col2)

        Dim dv As DataView = New DataView(tblProp, "oID_tblStudy > 0", "oID_tblReportTable ASC", DataViewRowState.CurrentRows)

        Me.dgvTblProps.DataSource = dv
        Me.dgvTblProps.AutoResizeColumns()


        Dim Count1 As Integer


        'establish this table
        ttTables = tblReportTables.Copy

        ttTables.AcceptChanges()

        'For Count1 = 0 To tblReportTables.Columns.Count - 1

        '    Dim col10 As New DataColumn
        '    col10.ColumnName = tblReportTables.Columns(Count1).ColumnName
        '    col10.Caption = tblReportTables.Columns(Count1).Caption
        '    col10.DataType = tblReportTables.Columns(Count1).DataType
        '    ttTables.Columns.Add(col10)


        'Next
    End Sub


    Sub LoadAdded()

        Dim strF As String
        Dim strS As String
        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim str1 As String
        Dim str2 As String

        dgv1 = Me.dgvAdded
        dgv2 = Me.dgvReportTableConfiguration

        If Me.rbAll.Checked Then
            strF = "ID_TBLSTUDIES = " & id_tblStudies
            strS = "INTORDER ASC"
        Else
            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDE <> 0"
            strS = "INTORDER ASC"
        End If

        'dvAdded = New DataView(tblReportTable, strF, strS, DataViewRowState.Added)
        'dvAdded = New DataView(ttTables, strF, strS, DataViewRowState.Added)
        dvAdded = New DataView(ttTables, strF, strS, DataViewRowState.CurrentRows)
        dvAdded.Sort = "ID_TBLREPORTTABLE ASC"

        dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv1.DataSource = dvAdded

        Dim int1 As Int16 = dvAdded.Count

        Dim Count1 As Short
        Dim Count2 As Short
        For Count1 = 0 To dgv1.Columns.Count - 1

            dgv1.Columns(Count1).Visible = False

        Next

        For Count1 = 1 To 3

            Select Case Count1
                Case 1
                    str1 = "CHARHEADINGTEXT"
                    str2 = "Table Title"
                Case 2
                    str1 = "CHARFCID"
                    str2 = "FC ID"
                Case 3
                    str1 = "BOOLINCLUDE"
                    str2 = "Incl."
            End Select

            dgv1.Columns(str1).Visible = True
            dgv1.Columns(str1).HeaderText = str2
            dgv1.Columns(str1).ReadOnly = True
        Next

        'For Count1 = 0 To dgv1.Columns.Count - 1

        '    str1 = dgv1.Columns(Count1).Name
        '    For Count2 = 0 To dgv2.Columns.Count - 1
        '        str2 = dgv1.Columns(Count2).Name
        '        If StrComp(str1, str2, CompareMethod.Text) = 0 Then
        '            dgv1.Columns(Count1).Visible = dgv2.Columns(Count2).Visible
        '            dgv1.Columns(Count1).HeaderText = dgv2.Columns(Count2).HeaderText
        '        End If
        '    Next

        '    'dgv1.Columns(Count1).Visible = dgv2.Columns(Count1).Visible
        '    'dgv1.Columns(Count1).HeaderText = dgv2.Columns(Count1).HeaderText
        '    'dgv1.Columns(Count1).ReadOnly = dgv2.Columns(Count1).ReadOnly

        'Next

        dvAdded.AllowDelete = False
        dvAdded.AllowEdit = False
        dvAdded.AllowNew = False

        dgv1.RowHeadersWidth = 25
        dgv1.AutoResizeRows()

        dgv1.AutoResizeColumns()

        dgv1.ColumnHeadersDefaultCellStyle.Font = New Font(dgv1.Font, FontStyle.Bold)

        Call ApplyColors()

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


    Sub ConfigReportTable()

        Dim Count1 As Int32
        Dim dgv As DataGridView = Me.dgvReportTableConfiguration

        Dim strF As String
        Dim strS As String

        Dim intF As Short
        Dim int1 As Int16

        Dim var1

        If Me.rbShowAllRTConfig.Checked Then
            intF = -1
        Else
            intF = 0
        End If

        Dim tbl1 As DataTable

        If boolGuWuOracle Then
            'tbl1 = ta_tblReportTable.GetDataBy_ID_TBLSTUDIES(gTblStudies)
        ElseIf boolGuWuAccess Then
            tbl1 = ta_tblReportTableAcc.GetDataBy_ID_TBLSTUDIES(gTblStudies)
        ElseIf boolGuWuSQLServer Then
            tbl1 = ta_tblReportTableSQLServer.GetDataBy_ID_TBLSTUDIES(gTblStudies)
        End If
        int1 = tblReportTable.Rows.Count 'DEBUG

        If Me.rbShowAllRTConfig.Checked Then
            strF = "ID_TBLSTUDIES = " & gTblStudies
        Else
            strF = "ID_TBLSTUDIES = " & gTblStudies & " AND BOOLINCLUDE = -1"
        End If

        strS = "INTORDER ASC"

        'dvTableRows = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
        dvTableRows = New DataView(tbl1)

        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv.DataSource = dvTableRows

        dvTableRows.RowFilter = strF
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

            'make boolInclude a checkbox
            Try
                Dim chk As New DataGridViewCheckBoxColumn()
                dgv.Columns.Add(chk)
                chk.HeaderText = "Incl."
                chk.Name = "chk"
                chk.DisplayIndex = 0
                chk.TrueValue = -1
                chk.FalseValue = 0
                chk.ThreeState = False
                chk.ReadOnly = True
            Catch ex As Exception
                var1 = ex.Message
            End Try

            dgv.RowHeadersWidth = 25

            dgv.ColumnHeadersDefaultCellStyle.Font = New Font(dgv.Font, FontStyle.Bold)

        End If


        dgv.AutoResizeRows()
        dgv.AutoResizeColumns()

        'fill chk
        Try
            For Count1 = 0 To dgv.Rows.Count - 1
                dgv("chk", Count1).Value = NZ(dgv("boolInclude", Count1).Value, 0)
                'int1 = NZ(dgv("boolInclude", Count1).Value, 0)
                'If int1 = 0 Then
                '    dgv("chk", Count1).Value = False
                'Else
                '    dgv("chk", Count1).Value = True
                'End If
                ''verify
                'var1 = dgv("chk", Count1).Value
                'var1 = var1

            Next
        Catch ex As Exception
            var1 = ex.Message
        End Try

    End Sub

    Sub LoadStudies()

        Dim dtbl As DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim strS As String

        If boolT Then
            strF = "BOOLACTIVE = -1"
            strS = "CHARTEMPLATENAME ASC"
            dtbl = tblTemplates
        Else
            strF = "ID_TBLSTUDIES <> " & id_tblStudies
            strS = "CHARWATSONSTUDYNAME ASC"
            dtbl = tblStudies
        End If

        rows = dtbl.Select(strF, strS)

        Me.cbxStudy.DataSource = rows

        If boolT Then
            Me.cbxStudy.DisplayMember = "CHARTEMPLATENAME"
            Me.cbxStudy.ValueMember = "ID_TBLSTUDIES"
        Else
            Me.cbxStudy.DisplayMember = "CHARWATSONSTUDYNAME"
            Me.cbxStudy.ValueMember = "ID_TBLSTUDIES"
        End If

        'select first row at end of formload
        Me.cbxStudy.SelectedIndex = -1

        strS = strS

        'Call GetStudyID()

    End Sub


    Private Sub cbxStudy_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbxStudy.SelectedIndexChanged

        If boolFormLoad Then
            'Exit Sub
        End If

        Dim checkInt As Short

        'If (IsNothing(cbxStudy.SelectedValue)) Then 'Don't show tables until Study is chosen
        '    panTables.Visible = False
        'Else
        '    If Me.rbShowAllRTConfig.Checked = True Then 'Change to just included tables if a Study was chosen
        '        rbShowIncludedRTConfig.Checked = True
        '    Else
        '        panTables.Visible = True
        '        Call GetStudyID()
        '    End If
        'End If

        Try
            Call GetStudyID()
        Catch ex As Exception

        End Try

    End Sub

    Sub GetStudyID()

        Dim var1

        Try

            Me.txtValue.Text = NZ(Me.cbxStudy.SelectedValue.ToString, 0)

            'gTblStudies = CLng(Me.txtValue.Text)
            Try
                gTblStudies = CLng(Me.txtValue.Text)
            Catch ex As Exception
                gTblStudies = 0
            End Try

            Call ConfigReportTable()

            'Call FillReportTable()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub rbShowIncludedRTConfig_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbShowIncludedRTConfig.CheckedChanged

        Call ConfigReportTable()

        'Call FillReportTable()

    End Sub

    Private Sub rbShowAllRTConfig_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbShowAllRTConfig.CheckedChanged

        'Call ConfigReportTable()

        'Call FillReportTable()
        Try
            'Call FillReportTable()
            Call ConfigReportTable()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub cmdAdd_Click(sender As System.Object, e As System.EventArgs) Handles cmdAdd.Click

        Call AddRows()

    End Sub

    Sub AddRows()

        Dim dgvS As DataGridView
        Dim dgvD As DataGridView

        Dim id1 As Int64

        Dim var1

        dgvS = Me.dgvReportTableConfiguration
        dgvD = Me.dgvAdded

        'first clear selected rows
        dgvD.ClearSelection()

        Dim dv As DataView
        'dv = dgvD.DataSource

        'dv.AllowNew = True

        Dim maxID As Int64
        Dim maxIDo As Int64

        maxID = GetMaxID("TBLREPORTTABLE", 1, False)
        maxIDo = maxID

        Dim Count1 As Integer
        Dim Count2 As Integer
        Dim Count3 As Short
        Dim strF As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim int1 As Short

        Dim rowSel As DataGridViewRow
        Dim strIS As String
        Dim arrS(dgvS.Rows.Count)

        int1 = 0
        For Each rowSel In dgvS.SelectedRows
            int1 = int1 + 1
            arrS(int1) = rowSel.Index
        Next

        For Count3 = int1 To 1 Step -1
            int1 = arrS(Count3)
            rowSel = dgvS.Rows(int1)

            Dim nR As DataRow = ttTables.NewRow
            nR.BeginEdit()
            For Count1 = 0 To dgvS.Columns.Count - 1
                var1 = rowSel.Cells(Count1).Value 'debug
                str1 = dgvS.Columns(Count1).Name
                If ttTables.Columns.Contains(str1) Then
                    nR.Item(str1) = rowSel.Cells(Count1).Value
                End If
            Next
            'get analytes
            For Count2 = 0 To tblAnalytesHome.Rows.Count - 1
                'if IntStd, then skip
                strIS = NZ(tblAnalytesHome.Rows(Count2).Item("IsIntStd"), "Yes")
                If StrComp(strIS, "Yes", CompareMethod.Text) = 0 Then 'skip
                Else
                    Try
                        str3 = tblAnalytesHome.Rows(Count2).Item("AnalyteDescription")
                        nR.Item(str3) = True
                    Catch ex As Exception

                    End Try
                End If
            Next
            'add charTableName
            id1 = nR.Item("ID_TBLCONFIGREPORTTABLES")
            Dim rowsID() As DataRow
            strF = "ID_TBLCONFIGREPORTTABLES = " & id1
            rowsID = tblConfigReportTables.Select(strF)
            nR.Item("CHARTABLENAME") = rowsID(0).Item("CHARTABLENAME")

            'enter new maxid
            maxID = maxID + 1
            nR.Item("ID_TBLREPORTTABLE") = maxID
            'enter correct id_tblstudies
            nR.Item("ID_TBLSTUDIES") = id_tblStudies

            nR.EndEdit()

            ttTables.Rows.Add(nR)

            'select last row and color it green
            dgvD.ClearSelection()
            dgvD.Rows(dgvD.Rows.Count - 1).Selected = True
            Dim drow As DataGridViewRow = dgvD.Rows(dgvD.Rows.Count - 1)
            drow.DefaultCellStyle.BackColor = Color.FromArgb(255, 207, 176) ' Color.FromArgb(229, 239, 249)


            id1 = rowSel.Cells("ID_TBLREPORTTABLE").Value

            'add items to tblProp
            Dim nRow As DataRow = tblProp.NewRow
            nRow.BeginEdit()
            nRow("oID_tblStudy") = CLng(Me.txtValue.Text)
            nRow("oID_tblReportTable") = id1
            nRow("nID_tblReportTable") = maxID
            nRow("BOOLADDED") = -1
            nRow.EndEdit()
            tblProp.Rows.Add(nRow)

        Next

        Call RecordColors()

        'For Each rowSel In dgvS.SelectedRows

        '    Dim nR As DataRow = ttTables.NewRow
        '    nR.BeginEdit()
        '    For Count1 = 0 To dgvS.Columns.Count - 1
        '        var1 = rowSel.Cells(Count1).Value 'debug
        '        str1 = dgvS.Columns(Count1).Name
        '        nR.Item(str1) = rowSel.Cells(Count1).Value
        '    Next
        '    'get analytes
        '    For Count2 = 0 To tblAnalytesHome.Rows.Count - 1
        '        'if IntStd, then skip
        '        strIS = NZ(tblAnalytesHome.Rows(Count2).Item("IsIntStd"), "Yes")
        '        If StrComp(strIS, "Yes", CompareMethod.Text) = 0 Then 'skip
        '        Else
        '            Try
        '                nR.Item(tblAnalytesHome.Rows(Count2).Item("AnalyteDescription")) = True
        '            Catch ex As Exception

        '            End Try
        '        End If
        '    Next
        '    'add charTableName
        '    id1 = nR.Item("ID_TBLCONFIGREPORTTABLES")
        '    Dim rowsID() As DataRow
        '    strF = "ID_TBLCONFIGREPORTTABLES = " & id1
        '    rowsID = tblConfigReportTables.Select(strF)
        '    nR.Item("CHARTABLENAME") = rowsID(0).Item("CHARTABLENAME")

        '    'enter new maxid
        '    maxID = maxID + 1
        '    nR.Item("ID_TBLREPORTTABLE") = maxID
        '    'enter correct id_tblstudies
        '    nR.Item("ID_TBLSTUDIES") = id_tblStudies

        '    nR.EndEdit()

        '    ttTables.Rows.Add(nR)

        '    id1 = rowSel.Cells("ID_TBLREPORTTABLE").Value

        '    'add items to tblProp
        '    Dim nRow As DataRow = tblProp.NewRow
        '    nRow.BeginEdit()
        '    nRow("oID_tblStudy") = CLng(Me.txtValue.Text)
        '    nRow("oID_tblReportTable") = id1
        '    nRow("nID_tblReportTable") = maxID
        '    nRow.EndEdit()
        '    tblProp.Rows.Add(nRow)

        'Next

        If maxIDo = maxID Then
        Else
            Call PutMaxID("TBLREPORTTABLE", maxID)
        End If

        dgvD.AutoResizeColumns()

        dgvD.ClearSelection()

    End Sub

    Private Sub cmdExitSave_Click(sender As System.Object, e As System.EventArgs) Handles cmdExitSave.Click

        Cursor.Current = Cursors.WaitCursor

        Call Import()

    End Sub

    Sub Import()


        Dim intR As Short
        Dim strM As String

        strM = "Do you wish to continue?"
        intR = MsgBox(strM, vbOKCancel, "Continue?...")
        If intR = 1 Then
        Else
            Exit Sub
        End If

        boolSave = True

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

        Dim intOrder As Int16 = 999

        Dim idTRT1 As Int64
        Dim idTRT2 As Int64


        Cursor.Current = Cursors.WaitCursor

        dv = Me.dgvTblProps.DataSource
        'debug
        int1 = dv.Count
        int1 = int1

        dtbl1 = tblReportTables

        'first do added
        rows = ttTables.Select("", "", DataViewRowState.Added)

        For Count1 = 0 To rows.Length - 1

            intOrder = intOrder + 1
            boolAdd = False

            'check to see if table already exists and simply needs to be made visible
            var1 = rows(Count1).Item("ID_TBLCONFIGREPORTTABLES")
            var2 = rows(Count1).Item("CHARHEADINGTEXT")
            var3 = NZ(rows(Count1).Item("CHARFCID"), "")
            idTRT1 = rows(Count1).Item("ID_TBLREPORTTABLE")

            'first look for id and included
            strF = "ID_TBLCONFIGREPORTTABLES = " & var1 & " AND BOOLINCLUDE = 0"
            Dim rowsA() As DataRow = dtbl1.Select(strF, "", DataViewRowState.CurrentRows)

            If rowsA.Length = 0 Then

                'add to table
                Dim nr1 As DataRow = dtbl1.NewRow
                nr1.BeginEdit()
                For Count2 = 0 To Me.dgvAdded.Columns.Count - 1
                    str1 = dgvAdded.Columns(Count2).Name
                    If dtbl1.Columns.Contains(str1) Then
                        nr1(str1) = rows(Count1).Item(str1)
                    End If
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

                For Count2 = 0 To Me.dgvAdded.Columns.Count - 1
                    str1 = dgvAdded.Columns(Count2).Name
                    If dtbl1.Columns.Contains(str1) Then
                        If StrComp(str1, "ID_TBLREPORTTABLE", CompareMethod.Text) = 0 Then 'skip
                        Else
                            var1 = rows(Count1).Item(str1)
                            rowsA(0).Item(str1) = rows(Count1).Item(str1)
                        End If
                    End If
                Next

                idTRT2 = rowsA(0).Item("ID_TBLREPORTTABLE")

                rowsA(0).Item("BOOLINCLUDE") = -1
                rowsA(0).Item("INTORDER") = intOrder
                rowsA(0).Item("ID_TBLSTUDIES") = id_tblStudies

                rowsA(0).EndEdit()

                boolAdd = False

            End If

            'now update tblProps

            If boolAdd Then
                'do nothing
            Else
                'update tblprobs
                strF1 = "nID_tblReportTable = " & idTRT1
                Dim rowsTP() As DataRow = tblProp.Select(strF1)
                If rowsTP.Length = 0 Then
                    var1 = var1
                Else
                    rowsTP(0).BeginEdit()
                    rowsTP(0).Item("nID_tblReportTable") = idTRT2
                    rowsTP(0).Item("BOOLADDED") = 0
                    rowsTP(0).EndEdit()
                End If
            End If

            'legend
            ''add items to tblProp
            'Dim nRow As DataRow = tblProp.NewRow
            'nRow.BeginEdit()
            'nRow("oID_tblStudy") = CLng(Me.txtValue.Text)
            'nRow("oID_tblReportTable") = id1
            'nRow("nID_tblReportTable") = maxID
            'nRow.EndEdit()
            'tblProp.Rows.Add(nRow)


 

        Next

        'col1.ColumnName = "oID_tblStudy"
        'col2.ColumnName = "oID_tblReportTable"
        'col3.ColumnName = "nID_tblReportTable"

        'update TBLTABLEPROPERTIES
        'legend:
        'dtbl2 = tblTableProperties
        'dtbl2 = ta_tblTablePropertiesAcc.GetDataBy_ID_TBLSTUDIES(id_tblStudies)
        maxID = GetMaxID("TBLTABLEPROPERTIES", 1, False)
        omaxID = maxID

        Dim intAdded As Short
        For Count1 = 0 To tblProp.Rows.Count - 1

            id1 = tblProp.Rows(Count1).Item("oID_tblStudy")
            id2 = tblProp.Rows(Count1).Item("oID_tblReportTable")
            id3 = tblProp.Rows(Count1).Item("nID_tblReportTable")
            intAdded = tblProp.Rows(Count1).Item("BOOLADDED")

            dtbl2.Clear()
            dtbl2.AcceptChanges()
            If boolGuWuOracle Then
                'dtbl2 = ta_tblTableProperties.GetDataBy_ID_TBLSTUDIES(id1)
            ElseIf boolGuWuAccess Then
                dtbl2 = ta_tblTablePropertiesAcc.GetDataBy_ID_TBLSTUDIES(id1)
            ElseIf boolGuWuSQLServer Then
                dtbl2 = ta_tblTablePropertiesSQLServer.GetDataBy_ID_TBLSTUDIES(id1)
            End If

            strF = "ID_TBLSTUDIES = " & id1 & " AND ID_TBLREPORTTABLE = " & id2
            Erase rows
            rows = dtbl2.Select(strF)
            If rows.Length = 0 Then
            Else

                If intAdded = -1 Then

                    'rows need to be added
                    For Count2 = 0 To rows.Length - 1

                        Dim nR1 As DataRow = tblTableProperties.NewRow
                        nR1.BeginEdit()
                        For Count3 = 0 To dtbl2.Columns.Count - 1
                            nR1.Item(Count3) = rows(Count2).Item(Count3)
                        Next
                        'replace maxid
                        maxID = maxID + 1
                        nR1.Item("ID_TBLTABLEPROPERTIES") = maxID
                        'replace studyid
                        nR1.Item("ID_TBLSTUDIES") = id_tblStudies
                        'replace ID_TBLREPORTTABLE
                        nR1.Item("ID_TBLREPORTTABLE") = id3
                        nR1.EndEdit()
                        tblTableProperties.Rows.Add(nR1)

                    Next

                Else

                    'rows need to be updated
                    For Count2 = 0 To rows.Length - 1

                        strF = "ID_TBLREPORTTABLE = " & id3 'don't need id_tblstudies
                        Dim rowsN() As DataRow = tblTableProperties.Select(strF)
                        If rowsN.Length = 0 Then
                        Else
                            rowsN(0).BeginEdit()
                            For Count3 = 0 To dtbl2.Columns.Count - 1
                                'don't enter all data
                                str1 = dtbl2.Columns(Count3).ColumnName
                                Select Case str1
                                    Case "ID_TBLTABLEPROPERTIES"
                                    Case "ID_TBLREPORTTABLE"
                                    Case "ID_TBLCONFIGREPORTTABLES"
                                    Case "ID_TBLSTUDIES"

                                    Case Else
                                        rowsN(0).Item(Count3) = rows(Count2).Item(Count3)
                                End Select

                            Next
                            rowsN(0).EndEdit()
                        End If

                    Next

                End If

                'now make entries for tblAutoAssignSamples
                '(ByVal idO As Int64, ByVal idN As Int64, ByVal idTStudiesO As Int64, ByVal idTStudiesN As Int64, ByVal boolGetNew As Boolean, ByVal intAdded As Short)
                Call CheckForAutoAssignSamplesTable(id2, id3, id1, id_tblStudies, True, intAdded)
               
            End If
        Next

        If maxID = omaxID Then
        Else
            Call PutMaxID("TBLTABLEPROPERTIES", maxID)
        End If

        'Pause(1)

        'Call CheckForTblProperties(-1, 0)

        Call OrderDGV(frmH.dgvReportTableConfiguration, "INTORDER", "ID_TBLREPORTTABLE")
        Call OrderReportTableConfig()
        Call SetComboCell(frmH.dgvReportTableConfiguration, "CHARPAGEORIENTATION")
        Call AssessSampleAssignment()

        Cursor.Current = Cursors.Default

        Me.Visible = False

    End Sub

    Public Sub FormLoad()

        Cursor.Current = Cursors.Default

    End Sub


    Sub RecordColors()

        Dim dgv As DataGridView = Me.dgvAdded
        Dim Count1 As Int16
        Dim var1, var2

        '1=ID_TBLREPORTTABLE, 2=color
        For Count1 = 0 To dgv.Rows.Count - 1
            var1 = dgv("ID_TBLREPORTTABLE", Count1).Value
            var2 = dgv.Rows(Count1).DefaultCellStyle.BackColor
            arrC(1, Count1 + 1) = var1
            arrC(2, Count1 + 1) = var2
        Next

    End Sub

    Sub ApplyColors()

        Dim dgv As DataGridView = Me.dgvAdded
        Dim Count1 As Int16
        Dim Count2 As Int16
        Dim var1, var2, var3

        'redo colors
        '1=ID_TBLREPORTTABLE, 2=color
        For Count1 = 1 To dgv.Rows.Count
            var1 = dgv("ID_TBLREPORTTABLE", Count1 - 1).Value
            For Count2 = 1 To UBound(arrC, 2)
                var3 = arrC(1, Count2)
                var2 = arrC(2, Count2)
                If var1 = var3 Then
                    dgv.Rows(Count1 - 1).DefaultCellStyle.BackColor = var2
                    Exit For
                End If

            Next
        Next

    End Sub

    Sub RemoveAddedRows(boolAll As Boolean)

        Dim int3 As Integer

        Dim dgvS As DataGridView
        Dim strF As String
        Dim rowsAdded() As DataRow
        Dim rowsProp() As DataRow
        Dim intRows As Int32

        Dim id1 As Int64

        Dim var1, var2, var3

        dgvS = Me.dgvAdded

        Dim dv As DataView
        dv = dgvS.DataSource

        Dim Count1 As Integer
        Dim Count2 As Integer

        Dim rowSel As DataGridViewRow

        Dim arrR()

        If boolAll Then
            intRows = dv.Count
        Else
            intRows = dgvS.SelectedRows.Count
        End If

        ReDim arrR(intRows)

        Dim int1 As Int32
        Dim int2 As Int32
        Dim strM As String

        If boolAll Then

            ttTables.RejectChanges()

        Else

            'VB stuff: when grid is refreshed, all colors revert to default

            int1 = 0

            For Each rowSel In dgvS.SelectedRows
                id1 = rowSel.Cells("ID_TBLREPORTTABLE").Value
                int1 = int1 + 1
                arrR(int1) = id1
            Next

            dv.AllowDelete = True

            For Count1 = 1 To int1

                id1 = arrR(Count1)
                int2 = dv.Find(id1)

                If int2 = -1 Then
                Else

                    'allow only added rows to be deleted
                    strF = "ID_TBLREPORTTABLE = " & id1
                    Dim rows() As DataRow = ttTables.Select(strF, "", DataViewRowState.Added)
                    If rows.Length = 0 Then
                        strM = "Only added rows may be removed."
                        strM = strM & ChrW(10) & ChrW(10) & "Please redo remove action and select only added (orange) rows."
                        MsgBox(strM, vbInformation, "Invalid action...")
                        GoTo end1
                    End If


                    dv.Delete(int2)

                    'remove from tblprop
                    strF = "nID_tblReportTable = " & id1
                    'rowsProp = tblProp.Select(strF)
                    Try
                        rowsProp = tblProp.Select(strF)
                        rowsProp(0).Delete()
                    Catch ex As Exception

                    End Try

                End If

            Next

end1:

            dv.AllowDelete = False

            dgvS.AutoResizeColumns()

            Call ApplyColors()


        End If

    End Sub



    Private Sub cmdExitNoSave_Click(sender As System.Object, e As System.EventArgs) Handles cmdExitNoSave.Click

        boolSave = False

        Me.Visible = False

    End Sub

    Private Sub cmdRemove_Click(sender As System.Object, e As System.EventArgs) Handles cmdRemove.Click

        Call RemoveAddedRows(False)

    End Sub

    Private Sub gbRTC_Enter(sender As Object, e As EventArgs) Handles gbRTC.Enter

    End Sub

    Private Sub lblSource_Click(sender As Object, e As EventArgs) Handles lblSource.Click

    End Sub

    Private Sub rbTemplate_CheckedChanged(sender As Object, e As EventArgs) Handles rbTemplate.CheckedChanged

        Call DoLabels()

        boolT = True
        Call LoadStudies()

        Try
            Me.cbxStudy.SelectedIndex = 0
        Catch ex As Exception

        End Try

    End Sub

    Private Sub rbStudy_CheckedChanged(sender As Object, e As EventArgs) Handles rbStudy.CheckedChanged

        Call DoLabels()

        boolT = False
        Call LoadStudies()
        Try
            Me.cbxStudy.SelectedIndex = 0
        Catch ex As Exception

        End Try


    End Sub

    Private Sub rbIncl_CheckedChanged(sender As Object, e As EventArgs) Handles rbIncl.CheckedChanged

        Try
            Call LoadAdded()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub rbAll_CheckedChanged(sender As Object, e As EventArgs) Handles rbAll.CheckedChanged

        Try
            Call LoadAdded()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub dgvReportTableConfiguration_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvReportTableConfiguration.CellContentClick

    End Sub

    Private Sub dgvReportTableConfiguration_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles dgvReportTableConfiguration.CellFormatting

        Dim dgv1 As DataGridView = Me.dgvReportTableConfiguration
        Dim var1

        Dim strN As String

        strN = dgv1.Columns(e.ColumnIndex).Name

        Try
            If StrComp(strN, "chk", CompareMethod.Text) = 0 Then
                'Debug.Print("dgv1_CellFormatting Row=" & e.RowIndex & ", Col=" & e.ColumnIndex & ", Value=" & e.Value & ", DesiredType=" & e.DesiredType.ToString & ", CellStyle=" & e.CellStyle.ToString)
                If dgv1.Rows(e.RowIndex).Cells("boolInclude").Value = 0 Then e.Value = False Else e.Value = True
                e.FormattingApplied = True
            End If
        Catch ex As Exception
            var1 = ex.Message
        End Try
    

    End Sub

    Private Sub dgvReportTableConfiguration_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles dgvReportTableConfiguration.CurrentCellDirtyStateChanged

        Dim dgv1 As DataGridView = Me.dgvReportTableConfiguration
        Dim var1

        Dim strN As String

        Try

            strN = dgv1.Columns(dgv1.CurrentCell.ColumnIndex).Name

            If StrComp(strN, "chk", CompareMethod.Text) = 0 Then
                If dgv1.IsCurrentCellDirty Then
                    dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit)
                End If
            End If
        Catch ex As Exception
            var1 = ex.Message
        End Try
  

    End Sub

  
    Private Sub dgvAdded_Sorted(sender As Object, e As EventArgs) Handles dgvAdded.Sorted

        Call ApplyColors()

    End Sub
End Class