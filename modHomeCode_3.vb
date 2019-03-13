Option Compare Text

Module modHomeCode_3


    Sub UpdateWord_dgv()

        If boolFormLoad Then
            Exit Sub
        End If


        'filter dgvReportStatementWord
        Dim intRow As Short
        Dim intCol As Short
        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim Count1 As Short
        Dim dv2 As System.Data.DataView ' DataView
        Dim strF2 As String
        Dim id1 As Int64
        Dim id2 As Int64
        Dim var1
        Dim boolGo As Boolean
        Dim strS As String

        dgv1 = frmH.dgvReportStatements
        dgv2 = frmH.dgvReportStatementWord

        If dgv1.Rows.Count = 0 Then
            GoTo end1
        End If
        If dgv1.CurrentRow Is Nothing Then
            'select first row
            intRow = 0
            For Count1 = 0 To dgv1.Columns.Count - 1
                If dgv1.Columns(Count1).Visible Then
                    dgv1.CurrentCell = dgv1.Rows(intRow).Cells(Count1)
                    dgv1.Rows(intRow).Selected = True
                    Exit For
                End If
            Next
        Else
            intRow = dgv1.CurrentRow.Index
        End If

        Try
            id1 = NZ(dgv1("ID_TBLCONFIGBODYSECTIONS", intRow).Value, 0)
        Catch ex As NullReferenceException
            id1 = 0
        End Try

        Try
            id2 = NZ(dgv1("ID_TBLWORDSTATEMENTS", intRow).Value, 0)
        Catch ex As NullReferenceException
            id2 = 0
        End Try

        'id1 = NZ(dgv1("ID_TBLCONFIGBODYSECTIONS", intRow).Value, 0)
        'id2 = NZ(dgv1("ID_TBLWORDSTATEMENTS", intRow).Value, 0)
        strF2 = "ID_TBLCONFIGBODYSECTIONS = " & id1 & " AND CHARWORDSTATEMENT = 'Active'"

        dv2 = dgv2.DataSource

        Try
            dv2.RowFilter = strF2
            strS = "CHARTITLE ASC"
            dv2.Sort = strS
        Catch ex As NullReferenceException
            'MsgBox(ex.Message)
        End Try


        dgv2.AutoResizeRows()

        'select appropriate row
        boolGo = False
        For Count1 = 0 To dgv2.Rows.Count - 1
            var1 = dgv2("ID_TBLWORDSTATEMENTS", Count1).Value
            If var1 = id2 Then
                intRow = Count1
                boolGo = True
                Exit For
            End If
        Next


        If boolGo Then
            For Count1 = 0 To dgv2.Columns.Count - 1
                If dgv2.Columns(Count1).Visible Then
                    dgv2.CurrentCell = dgv2.Rows(intRow).Cells(Count1)
                    dgv2.Rows(intRow).Selected = True
                    Exit For
                End If
            Next
        Else 'select first row
            If dgv2.Rows.Count = 0 Then
            Else
                intRow = 0
                For Count1 = 0 To dgv2.Columns.Count - 1
                    If dgv2.Columns(Count1).Visible Then
                        dgv2.CurrentCell = dgv2.Rows(intRow).Cells(Count1)
                        dgv2.Rows(intRow).Selected = True
                        Exit For
                    End If
                Next
            End If
        End If

end1:

    End Sub

    Sub ReportTableHeaderConfig()

        Dim tbl As System.Data.DataTable
        Dim tbl1 As System.Data.DataTable
        Dim dv As System.Data.DataView
        Dim ct1 As Short
        Dim Count1 As Short
        Dim var1, var2, var3
        Dim str1 As String
        Dim str2 As String
        Dim intRows As Short

        'first populate cbxReportTableTypes
        tbl = tblConfigReportType
        dv = tbl.DefaultView
        ct1 = dv.Count
        str1 = "charReportType ASC"
        dv.Sort = str1
        ct1 = dv.Count

        Call ConfigReportTablesColumns()

        'configure tblHeaderLabels
        tbl = tblHeaderLabels
        var1 = tbl.Rows.Count 'debug
        If tbl.Columns.Count = 0 Then

            Dim col1 As New DataColumn
            col1.DataType = System.Type.GetType("System.Int64")
            col1.ColumnName = "id_tblConfigHeaderLookup"
            col1.ReadOnly = False
            tbl.Columns.Add(col1)
            Dim col2 As New DataColumn
            col2.DataType = System.Type.GetType("System.Int16")
            col2.ColumnName = "id_tblConfigReportTables"
            col2.ReadOnly = False
            tbl.Columns.Add(col2)
            Dim col3 As New DataColumn
            col3.DataType = System.Type.GetType("System.String")
            col3.ColumnName = "charColumnLabel"
            col3.ReadOnly = False
            tbl.Columns.Add(col3)
            Dim col4 As New DataColumn
            col4.DataType = System.Type.GetType("System.Boolean")
            'col4.DataType = System.Type.GetType("System.Short")
            col4.ColumnName = "boolInclude"
            col4.ReadOnly = False
            tbl.Columns.Add(col4)
            Dim col5 As New DataColumn
            col5.DataType = System.Type.GetType("System.Int16")
            col5.ColumnName = "intOrder"
            col5.ReadOnly = False
            tbl.Columns.Add(col5)
            Dim col6 As New DataColumn
            col6.DataType = System.Type.GetType("System.String")
            col6.ColumnName = "charUserLabel"
            col6.ReadOnly = False
            tbl.Columns.Add(col6)
            Dim col7 As New DataColumn
            col7.DataType = System.Type.GetType("System.Int64")
            col7.ColumnName = "id_tblStudies"
            col7.ReadOnly = False
            tbl.Columns.Add(col7)
            Dim col8 As New DataColumn
            col8.DataType = System.Type.GetType("System.Int64")
            col8.ColumnName = "id_tblReportTableHeaderConfig"
            col8.ReadOnly = False
            tbl.Columns.Add(col8)
            Dim col9 As New DataColumn
            col9.DataType = System.Type.GetType("System.String")
            col9.ColumnName = "charWatsonTable"
            col9.ReadOnly = False
            tbl.Columns.Add(col9)
            Dim col10 As New DataColumn
            col10.DataType = System.Type.GetType("System.String")
            col10.ColumnName = "charWatsonField"
            col10.ReadOnly = False
            tbl.Columns.Add(col10)

        End If

        'populate tbl
        tbl1 = tblConfigHeaderLookup
        ct1 = tbl1.Rows.Count
        intRows = tbl.Rows.Count 'tbl = tblConfigReportType
        If intRows = 0 Then 'initialize
            Dim row As DataRow
            For Count1 = 0 To ct1 - 1

                str2 = tbl1.Rows.Item(Count1).Item("charColumnLabel")
                If InStr(1, str2, "Subject Number", CompareMethod.Text) > 0 Then
                ElseIf InStr(1, str2, "Group Number", CompareMethod.Text) > 0 Then
                Else
                    row = tbl.NewRow
                    row("id_tblConfigHeaderLookup") = tbl1.Rows.Item(Count1).Item("id_tblConfigHeaderLookup")
                    row("id_tblConfigReportTables") = tbl1.Rows.Item(Count1).Item("id_tblConfigReportTables")
                    row("charColumnLabel") = tbl1.Rows.Item(Count1).Item("charColumnLabel")
                    var1 = tbl1.Rows.Item(Count1).Item("charColumnLabel")
                    row("charUserLabel") = tbl1.Rows.Item(Count1).Item("charColumnLabel")
                    If StrComp(var1, "Matrix", CompareMethod.Text) = 0 Then 'debugging
                        var2 = tbl1.Rows.Item(Count1).Item("BOOLDEFAULT") 'True
                        var3 = var2
                    End If
                    row("boolInclude") = tbl1.Rows.Item(Count1).Item("BOOLDEFAULT") 'True
                    row("intOrder") = tbl1.Rows.Item(Count1).Item("intOrder")
                    row("id_tblStudies") = 0
                    row("id_tblReportTableHeaderConfig") = 0
                    row("charWatsonTable") = tbl1.Rows.Item(Count1).Item("charWatsonTable")
                    row("charWatsonField") = tbl1.Rows.Item(Count1).Item("charWatsonField")
                    tbl.Rows.Add(row)
                End If

            Next

            var1 = tbl.Rows.Count 'debug

            Call ConfigReportTableHeadersColumns()

            'now populate dgReportTables based on selection
            Call ReportTableHeaderPopulate()

            'now populate frmh.dgvreporttableheaderconfig based on dgReportTableHeader selection
            Call ReportTableHeaderConfigPopulate()

        Else 're-initialize

            'debug
            var1 = tbl.Rows.Count
            var2 = tbl1.Rows.Count

            Dim intRow As Int32

            intRow = -1
            For Count1 = 0 To ct1 - 1
                If Count1 = 169 Then
                    var1 = var1 'debug
                End If
                str2 = tbl1.Rows.Item(Count1).Item("charColumnLabel")
                str2 = tbl1.Rows(Count1).Item("charColumnLabel")
                If InStr(1, str2, "Subject Number", CompareMethod.Text) > 0 Then
                    var1 = var1
                ElseIf InStr(1, str2, "Group Number", CompareMethod.Text) > 0 Then
                    var1 = var1
                Else
                    intRow = intRow + 1
                    tbl.Rows(intRow).BeginEdit()
                    tbl.Rows(intRow).Item("id_tblConfigHeaderLookup") = tbl1.Rows.Item(Count1).Item("id_tblConfigHeaderLookup")
                    tbl.Rows(intRow).Item("id_tblConfigReportTables") = tbl1.Rows.Item(Count1).Item("id_tblConfigReportTables")
                    tbl.Rows(intRow).Item("charColumnLabel") = tbl1.Rows.Item(Count1).Item("charColumnLabel")
                    tbl.Rows(intRow).Item("charUserLabel") = tbl1.Rows.Item(Count1).Item("charColumnLabel")
                    tbl.Rows(intRow).Item("boolInclude") = tbl1.Rows.Item(Count1).Item("BOOLDEFAULT") 'True
                    tbl.Rows(intRow).Item("intOrder") = tbl1.Rows.Item(Count1).Item("intOrder")
                    tbl.Rows(intRow).Item("id_tblStudies") = 0
                    tbl.Rows(intRow).Item("id_tblReportTableHeaderConfig") = 0
                    tbl.Rows(intRow).Item("charWatsonTable") = tbl1.Rows.Item(Count1).Item("charWatsonTable")
                    tbl.Rows(intRow).Item("charWatsonField") = tbl1.Rows.Item(Count1).Item("charWatsonField")
                    tbl.Rows(intRow).EndEdit()
                End If
            Next

        End If



    End Sub

    Sub DoCancelRTHConfig()

        tblReportTableHeaderConfig.RejectChanges()

        Call ReportTableHeaderPopulateData()

    End Sub

    Sub ReportTableHeaderPopulateData()

        Dim tbl1 As System.Data.DataTable
        Dim tbl2 As System.Data.DataTable
        Dim tbl3 As System.Data.DataTable
        Dim tbl4 As System.Data.DataTable
        Dim drows1() As DataRow
        Dim drows2() As DataRow
        Dim drows3() As DataRow
        Dim drows4() As DataRow
        Dim drowsF() As DataRow
        Dim ct1 As Short
        Dim ct2 As Short
        Dim ct3 As Short
        Dim ct4 As Short
        Dim ctF As Short
        Dim dv As System.Data.DataView
        Dim Count1 As Short
        Dim Count2 As Short
        Dim var1, var2, var3
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strF As String
        Dim int1 As Short
        Dim int2 As Short
        Dim intS As Long
        Dim strS As String
        Dim boolExists As Boolean
        Dim row As DataRow
        Dim drowsE() As DataRow
        Dim intRowsE As Short
        Dim strF1 As String
        Dim maxID As Int64
        Dim maxIDo As Int64

        tbl1 = tblReportTableHeaderConfig
        tbl2 = tblHeaderLabels
        ct2 = tbl2.Rows.Count
        tbl3 = tblConfigHeaderLookup
        ct3 = tbl3.Rows.Count

        strF = "id_tblStudies = " & id_tblStudies
        drows1 = tbl1.Select(strF)
        ct1 = drows1.Length

        maxID = GetMaxID("tblReportTableHeaderConfig", 1, False)
        maxIDo = maxID

        If ct1 = 0 Then 'use default values
            For Count1 = 0 To ct3 - 1 'get info from tblConfigHeaderLookup

                str2 = tbl3.Rows.Item(Count1).Item("charColumnLabel")
                If InStr(1, str2, "Subject Number", CompareMethod.Text) > 0 Then
                ElseIf InStr(1, str2, "Group Number", CompareMethod.Text) > 0 Then
                Else
                    var1 = tbl3.Rows.Item(Count1).Item("id_tblConfigHeaderLookup")
                    strF = "id_tblConfigHeaderLookup = " & var1
                    drows2 = tbl2.Select(strF)
                    ct2 = drows2.Length
                    'row = tbl2.Rows.item(Count1)
                    drows2(0).Item("id_tblConfigHeaderLookup") = tbl3.Rows.Item(Count1).Item("id_tblConfigHeaderLookup")
                    drows2(0).Item("id_tblConfigReportTables") = tbl3.Rows.Item(Count1).Item("id_tblConfigReportTables")
                    drows2(0).Item("charColumnLabel") = tbl3.Rows.Item(Count1).Item("charColumnLabel")
                    var2 = tbl3.Rows.Item(Count1).Item("boolDefault") 'True
                    If var2 = 0 Then
                        drows2(0).Item("boolInclude") = False
                    Else
                        drows2(0).Item("boolInclude") = True
                    End If
                    drows2(0).Item("intOrder") = tbl3.Rows.Item(Count1).Item("intOrder")
                    drows2(0).Item("charUserLabel") = tbl3.Rows.Item(Count1).Item("charColumnLabel")
                    drows2(0).Item("id_tblStudies") = id_tblStudies
                    drows2(0).Item("id_tblReportTableHeaderConfig") = 0
                    drows2(0).Item("charWatsonTable") = tbl3.Rows.Item(Count1).Item("charWatsonTable")
                    drows2(0).Item("charWatsonField") = tbl3.Rows.Item(Count1).Item("charWatsonField")
                    drows2(0).EndEdit()
                End If

        
            Next
        Else 'enter existing data

            For Count1 = 0 To ct2 - 1 'coming from tblHeaderLabels
                var1 = tbl2.Rows.Item(Count1).Item("id_tblConfigHeaderLookup") 'coming from tblHeaderLabels
                boolExists = False
                strF1 = "id_tblStudies = " & id_tblStudies & " AND id_tblConfigHeaderLookup = " & var1
                drowsE = tbl1.Select(strF1)
                intRowsE = drowsE.Length
                If intRowsE = 0 Then
                    boolExists = False
                Else
                    boolExists = True
                End If
                'For Count2 = 0 To drows1.Length - 1
                '    var2 = drows1(Count2).Item("id_tblConfigHeaderLookup") 'coming from tblReportTableHeaderConfig
                '    If var1 = var2 Then 'exists
                '        boolExists = True
                '        Exit For
                '    End If
                'Next
                If boolExists Then 'update row
                    row = tbl2.Rows.Item(Count1)
                    row.BeginEdit()
                    row("id_tblReportTableHeaderConfig") = drowsE(0).Item("id_tblConfigReportTables")
                    row("id_tblStudies") = drowsE(0).Item("id_tblStudies")
                    row("id_tblConfigReportTables") = drowsE(0).Item("id_tblConfigReportTables")
                    row("id_tblConfigHeaderLookup") = drowsE(0).Item("id_tblConfigHeaderLookup")
                    row("charUserLabel") = drowsE(0).Item("charUserLabel")
                    row("intOrder") = drowsE(0).Item("intOrder")
                    var1 = drowsE(0).Item("boolInclude")
                    If var1 = -1 Then
                        row("boolInclude") = True
                    Else
                        row("boolInclude") = False
                    End If
                    row.EndEdit()

                    '***
                    'row("id_tblConfigHeaderLookup") = tbl1.Rows.item(Count1).Item("id_tblConfigHeaderLookup")
                    'row("id_tblConfigReportTables") = tbl1.Rows.item(Count1).Item("id_tblConfigReportTables")
                    'row("charColumnLabel") = tbl1.Rows.item(Count1).Item("charColumnLabel")
                    'row("boolInclude") = True
                    'row("intOrder") = tbl1.Rows.item(Count1).Item("intOrder")
                    'row("charUserLabel") = tbl1.Rows.item(Count1).Item("charColumnLabel")
                    'row("id_tblStudies") = 0
                    'row("id_tblReportTableHeaderConfig") = 0
                    'row("charWatsonTable") = tbl1.Rows.item(Count1).Item("charWatsonTable")
                    'row("charWatsonField") = tbl1.Rows.item(Count1).Item("charWatsonField")

                    '***

                Else 'add row to tbl1
                    Dim nRow As DataRow = tbl1.NewRow
                    maxID = maxID + 1
                    nRow.BeginEdit()
                    nRow("id_tblReportTableHeaderConfig") = maxID
                    nRow("id_tblStudies") = id_tblStudies
                    nRow("id_tblConfigReportTables") = tbl2.Rows.Item(Count1).Item("id_tblConfigReportTables")
                    nRow("id_tblConfigHeaderLookup") = tbl2.Rows.Item(Count1).Item("id_tblConfigHeaderLookup")
                    nRow("charUserLabel") = tbl2.Rows.Item(Count1).Item("charUserLabel")
                    nRow("intOrder") = tbl2.Rows.Item(Count1).Item("intOrder")
                    var1 = tbl2.Rows.Item(Count1).Item("boolInclude")
                    nRow("boolInclude") = tbl2.Rows.Item(Count1).Item("boolInclude")
                    nRow.EndEdit()

                    tbl1.Rows.Add(nRow)

                End If
            Next
        End If

        If maxID = maxIDo Then
        Else
            PutMaxID("tblReportTableHeaderConfig", maxID)
            'save tblReportTableHeaderConfig

            Try
                If boolGuWuOracle Then
                    Try
                        ta_tblReportTableHeaderConfig.Update(tblReportTableHeaderConfig)
                    Catch ex As DBConcurrencyException
                        ''msgbox("aaReport Table Header Config: " & ex.Message)
                        'ds2005.TBLREPORTTABLEHEADERCONFIG.Merge('ds2005.TBLREPORTTABLEHEADERCONFIG, True)
                    End Try

                ElseIf boolGuWuAccess Then
                    Try
                        ta_tblReportTableHeaderConfigAcc.Update(tblReportTableHeaderConfig)
                    Catch ex As DBConcurrencyException
                        ''msgbox("aaReport Table Header Config: " & ex.Message)
                        'ds2005Acc.TBLREPORTTABLEHEADERCONFIG.Merge('ds2005Acc.TBLREPORTTABLEHEADERCONFIG, True)
                    End Try

                ElseIf boolGuWuSQLServer Then
                    Try
                        ta_tblReportTableHeaderConfigSQLServer.Update(tblReportTableHeaderConfig)
                    Catch ex As DBConcurrencyException
                        ''msgbox("aaReport Table Header Config: " & ex.Message)
                        'ds2005Acc.TBLREPORTTABLEHEADERCONFIG.Merge('ds2005Acc.TBLREPORTTABLEHEADERCONFIG, True)
                    End Try

                End If
            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try

        End If
        'refresh dg
        dv = tbl2.DefaultView
        dv.AllowNew = False
        dv.AllowDelete = False
        dv.AllowEdit = True
        frmH.dgvReportTableHeaderConfig.DataSource = dv
        'frmH.dgvReportTableHeaderConfig.Refresh()

        ''autosize grid
        'Dim ts1 As DataGridTableStyle
        'Dim gs As DataGridColumnStyle
        'int1 = 0
        'ts1 = frmH.frmh.dgvreporttableheaderconfig.TableStyles(0)
        'For Each gs In ts1.GridColumnStyles
        '    int1 = int1 + 1
        'Next
        'Call AutoSizeGrid(100, dv, frmH.frmh.dgvreporttableheaderconfig, dv.Count, int1, 0, False)

    End Sub
    Sub SaveReportTableHeaderConfig()

        Dim tbl1 As System.Data.DataTable
        Dim tbl2 As System.Data.DataTable
        Dim tbl3 As System.Data.DataTable
        Dim tbl4 As System.Data.DataTable
        Dim drows1() As DataRow
        Dim drows2() As DataRow
        Dim drows3() As DataRow
        Dim drows4() As DataRow
        Dim drowsF() As DataRow
        Dim ct1 As Short
        Dim ct2 As Short
        Dim ct3 As Short
        Dim ct4 As Short
        Dim ctF As Short
        Dim dv As System.Data.DataView
        Dim Count1 As Short
        Dim Count2 As Short
        Dim var1, var2, var3
        Dim str1 As String
        Dim strF As String
        Dim int1 As Short
        Dim int2 As Short
        Dim intS As Long
        Dim strS As String
        Dim boolExists As Boolean
        Dim row As DataRow

        tbl1 = tblReportTableHeaderConfig
        tbl2 = tblHeaderLabels
        ct2 = tbl2.Rows.Count
        tbl3 = tblConfigHeaderLookup
        ct3 = tbl3.Rows.Count


        Dim tblMax As System.Data.DataTable
        Dim rowsMax() As DataRow
        Dim strFMax As String
        Dim maxID, maxID1

        'If boolGuWuOracle Then
        '    ta_tblMaxID.Fill(tblMaxID)
        'ElseIf boolGuWuAccess Then
        '    ta_tblMaxIDAcc.Fill(tblMaxID)
        'ElseIf boolGuWuSQLServer Then
        '    ta_tblMaxIDSQLServer.Fill(tblMaxID)
        'End If
        'strFMax = "charTable = 'tblReportTableHeaderConfig'"
        'tblMax = tblMaxID
        'rowsMax = tblMax.Select(strFMax)
        'maxID = rowsMax(0).Item("nummaxid")

        maxID = GetMaxID("tblReportTableHeaderConfig", 1, False) 'if maxid increment is 1, then getmaxid already does putmaxid
        maxID1 = maxID

        'first see if there are existing data
        strF = "id_tblStudies = " & id_tblStudies
        drows1 = tbl1.Select(strF)
        ct1 = drows1.Length
        If ct1 = 0 Then 'add brand new row
            For Count1 = 0 To ct2 - 1
                row = tbl1.NewRow
                maxID = maxID + 1
                row.BeginEdit()
                row("id_tblReportTableHeaderConfig") = maxID
                row("id_tblStudies") = id_tblStudies
                row("id_tblConfigReportTables") = tbl2.Rows.Item(Count1).Item("id_tblConfigReportTables")
                row("id_tblConfigHeaderLookup") = tbl2.Rows.Item(Count1).Item("id_tblConfigHeaderLookup")
                row("charUserLabel") = tbl2.Rows.Item(Count1).Item("charUserLabel")
                row("intOrder") = tbl2.Rows.Item(Count1).Item("intOrder")
                var1 = tbl2.Rows.Item(Count1).Item("boolInclude")
                If var1 Then
                    row("boolInclude") = -1
                Else
                    row("boolInclude") = 0
                End If
                row.EndEdit()
                tbl1.Rows.Add(row)
            Next
        ElseIf ct1 > 0 Then 'update and add new rows as needed
            For Count1 = 0 To ct2 - 1
                var1 = tbl2.Rows.Item(Count1).Item("id_tblConfigHeaderLookup") 'from tblHeaderLabels
                boolExists = False
                For Count2 = 0 To drows1.Length - 1
                    var2 = drows1(Count2).Item("id_tblConfigHeaderLookup")
                    If var1 = var2 Then 'exists
                        boolExists = True
                        Exit For
                    End If
                Next
                If boolExists Then 'update row
                    'row = tbl1.Rows.item(Count1)
                    row = drows1(Count2)
                    row.BeginEdit()
                Else 'add a new row to tbl1
                    row = tbl1.NewRow
                    row.BeginEdit()
                    maxID = maxID + 1
                    row.Item("id_tblReportTableHeaderConfig") = maxID
                End If
                row("id_tblStudies") = id_tblStudies
                row("id_tblConfigReportTables") = tbl2.Rows.Item(Count1).Item("id_tblConfigReportTables")
                row("id_tblConfigHeaderLookup") = tbl2.Rows.Item(Count1).Item("id_tblConfigHeaderLookup")
                row("charUserLabel") = tbl2.Rows.Item(Count1).Item("charUserLabel")
                row("intOrder") = tbl2.Rows.Item(Count1).Item("intOrder")
                var1 = tbl2.Rows.Item(Count1).Item("boolInclude")
                If var1 Then
                    row("boolInclude") = -1
                Else
                    row("boolInclude") = 0
                End If
                row.EndEdit()
                If boolExists Then 'update row
                    'row.EndEdit()
                Else 'add row
                    tbl1.Rows.Add(row)
                End If
            Next
        Else
        End If

        'frmH.dgvReportTableHeaderConfig.Update()

        Dim dvCheck As System.Data.DataView = New DataView(tblReportTableHeaderConfig)
        dvCheck.RowStateFilter = DataViewRowState.ModifiedCurrent
        Dim int10 As Short
        int10 = 1
        If int10 = 0 Then
        Else

            Call FillAuditTrailTemp(tblReportTableHeaderConfig)

            If boolGuWuOracle Then
                Try
                    ta_tblReportTableHeaderConfig.Update(tblReportTableHeaderConfig)
                Catch ex As DBConcurrencyException
                    ''msgbox("aaReport Table Header Config: " & ex.Message)
                    'ds2005.TBLREPORTTABLEHEADERCONFIG.Merge('ds2005.TBLREPORTTABLEHEADERCONFIG, True)
                End Try

            ElseIf boolGuWuAccess Then
                Try
                    ta_tblReportTableHeaderConfigAcc.Update(tblReportTableHeaderConfig)
                Catch ex As DBConcurrencyException
                    ''msgbox("aaReport Table Header Config: " & ex.Message)
                    'ds2005Acc.TBLREPORTTABLEHEADERCONFIG.Merge('ds2005Acc.TBLREPORTTABLEHEADERCONFIG, True)
                End Try

            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblReportTableHeaderConfigSQLServer.Update(tblReportTableHeaderConfig)
                Catch ex As DBConcurrencyException
                    ''msgbox("aaReport Table Header Config: " & ex.Message)
                    'ds2005Acc.TBLREPORTTABLEHEADERCONFIG.Merge('ds2005Acc.TBLREPORTTABLEHEADERCONFIG, True)
                End Try
            End If
        End If

        If maxID = maxID1 Then
        Else

            Call PutMaxID("tblReportTableHeaderConfig", maxID)

            'rowsMax(0).BeginEdit()
            'rowsMax(0).Item("nummaxid") = maxID
            'rowsMax(0).EndEdit()
            'If boolGuWuOracle Then
            '    Try
            '        ta_tblMaxID.Update(tblMaxID)
            '    Catch ex As DBConcurrencyException
            '        'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
            '    End Try
            'ElseIf boolGuWuAccess Then
            '    Try
            '        ta_tblMaxIDAcc.Update(tblMaxID)
            '    Catch ex As DBConcurrencyException
            '        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '    End Try
            'ElseIf boolGuWuSQLServer Then
            '    Try
            '        ta_tblMaxIDSQLServer.Update(tblMaxID)
            '    Catch ex As DBConcurrencyException
            '        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '    End Try
            'End If


        End If


    End Sub

    Sub ReportTableHeaderConfigPopulate()

        Dim tbl As System.Data.DataTable
        Dim dv As System.Data.DataView
        Dim ct1 As Short
        Dim Count1 As Short
        Dim var1
        Dim str1 As String
        Dim strF As String
        Dim int1 As Short
        Dim int2 As Short
        Dim intS As Long
        Dim drows() As DataRow
        Dim strS As String
        Dim boolE As Boolean
        Dim dgv As DataGridView
        Dim dgv1 As DataGridView

        dgv = frmH.dgvReportTables
        dgv1 = frmH.dgvReportTableHeaderConfig

        'determine index of dgReportTables
        If dgv.Rows.Count = 0 Then
            Exit Sub
        End If

        If dgv.CurrentRow Is Nothing Then
            int1 = 0
        Else
            int1 = frmH.dgvReportTables.CurrentRow.Index
        End If

        dv = dgv.DataSource
        intS = dv.Item(int1).Item("id_tblConfigReportTables")
        boolE = dv.Item(int1).Item("boolAllowHeaderExcl")

        '20180808 LEE:
        'Ignore the boolE thing.
        'If shown, allow config
        boolE = True

        'assign appropriate tables to dgReportTables
        '20190221 LEE:
        'Regr table must remove four items
        If intS = 2 Then
            strF = "id_tblConfigReportTables = " & intS & " AND (ID_TBLCONFIGHEADERLOOKUP = 97 OR ID_TBLCONFIGHEADERLOOKUP = 280)"
        Else
            strF = "id_tblConfigReportTables = " & intS
        End If

        strS = "intOrder ASC"
        tbl = tblHeaderLabels

        ''debug
        'Dim intI As Int64 = tbl.Rows.Count
        'intI = intI

        dv = tbl.DefaultView
        dv.RowFilter = strF
        dv.AllowNew = False
        dv.AllowEdit = True
        dv.AllowDelete = False
        dv.Sort = strS
        dgv1.DataSource = dv

        Call ConfigReportTableHeadersColumns()

        '20180808 LEE:
        'don't do this anymore
        'make notes in Notes box instead
        'Select Case intS
        '    Case 13, 14, 15 'recovery tables
        '        'change backcolor to denote that columns are read-only
        '        boolE = False
        '        dgv1.Columns("intOrder").DefaultCellStyle.BackColor = Color.LightGray
        '        dgv1.Columns("intOrder").DefaultCellStyle.ForeColor = Color.DarkGray
        '        dgv1.Columns("intOrder").DefaultCellStyle.SelectionBackColor = Color.LightGray
        '        dgv1.Columns("intOrder").DefaultCellStyle.SelectionForeColor = Color.DarkGray

        '        dgv1.Columns("boolInclude").DefaultCellStyle.BackColor = Color.LightGray
        '        dgv1.Columns("boolInclude").DefaultCellStyle.ForeColor = Color.DarkGray
        '        dgv1.Columns("boolInclude").DefaultCellStyle.SelectionBackColor = Color.LightGray
        '        dgv1.Columns("boolInclude").DefaultCellStyle.SelectionForeColor = Color.DarkGray

        '    Case Else
        '        dgv1.Columns("intOrder").DefaultCellStyle.BackColor = Color.White
        '        dgv1.Columns("intOrder").DefaultCellStyle.ForeColor = Color.Black
        '        dgv1.Columns("intOrder").DefaultCellStyle.SelectionBackColor = Color.DodgerBlue
        '        dgv1.Columns("intOrder").DefaultCellStyle.SelectionForeColor = Color.White

        '        dgv1.Columns("boolInclude").DefaultCellStyle.BackColor = Color.White
        '        dgv1.Columns("boolInclude").DefaultCellStyle.ForeColor = Color.Black
        '        dgv1.Columns("boolInclude").DefaultCellStyle.SelectionBackColor = Color.DodgerBlue
        '        dgv1.Columns("boolInclude").DefaultCellStyle.SelectionForeColor = Color.White

        'End Select

        'select first item
        If dv.Count = 0 Then
        Else
            'find first visible row
            For Count1 = 0 To dgv1.ColumnCount - 1
                If dgv1.Columns(Count1).Visible Then
                    Exit For
                End If
            Next
            dgv1.CurrentCell = dgv1.Rows(0).Cells(Count1)
            dgv1.Rows(0).Selected = True
        End If

        'determine if users are allowed to exclude table headings
        If boolE Then
            dgv1.Columns("intOrder").ReadOnly = False
            dgv1.Columns("boolInclude").ReadOnly = False
        Else
            dgv1.Columns("intOrder").ReadOnly = True
            dgv1.Columns("boolInclude").ReadOnly = True
        End If

        Select Case intS
            Case 36, 37, 38
                dgv1.Columns("intOrder").Visible = False
            Case Else
                dgv1.Columns("intOrder").Visible = True
        End Select


        'enter notes
        Select Case intS
            Case 32
                str1 = "Column 'Sample Name' not available if Report Table Configuration - 'Format table as a grid...' checkbox is checked or if number of QC Levels > 1."
                'str1 = str1 & ChrW(10) & "Columns 'Order' and 'Include' are ignored in this table. Order and Include are hard-coded."
                str1 = str1 & ChrW(10) & "Column 'Watson Run ID': 'Include' and 'Order' are ignored"
                str1 = str1 & ChrW(10) & "Column 'Sample Name': 'Order' is ignored"
                str1 = str1 & ChrW(10) & "Column 'Analysis Date': 'Include' and 'Order' are ignored"
            Case 1
                str1 = "None"
            Case 35
                str1 = "Column 'Watson Run ID': 'Include' and 'Order' are ignored"
                str1 = str1 & ChrW(10) & "Column 'Sample Name': 'Order' is ignored"
                str1 = str1 & ChrW(10) & "Column 'Analysis Date': 'Include' and 'Order' are ignored"
            Case Else
                str1 = "Columns 'Order' and 'Include' are ignored in this table. Order and Include are hard-coded."
                str1 = str1 & ChrW(10) & "Column 'Watson Run ID' is provided for user text modification only."
                str1 = str1 & ChrW(10) & "Column 'Analysis Date' is provided for user text modification only."

        End Select
        frmH.lblNotes2.Text = str1

    End Sub

    Sub ReportTableHeaderPopulate()

        Dim tbl As System.Data.DataTable
        Dim dv As System.Data.DataView
        Dim ct1 As Short
        Dim Count1 As Short
        Dim var1
        Dim str1 As String
        Dim strF As String
        Dim int1 As Short
        Dim intS As Short
        Dim drows() As DataRow
        Dim strS As String
        Dim dgv As DataGridView

        dgv = frmH.dgvReportTables

        If boolCont Then 'continue
        Else 'stop
            Exit Sub
        End If
        'determine index of selected cbxReportTableTypes
        tbl = tblConfigReportType

        strS = "intOrder ASC"
        'tbl = frmh.qryConfigReportTables
        tbl = tblConfigReportTables
        dv = tbl.DefaultView

        dv.AllowNew = False
        dv.AllowEdit = False
        dv.AllowDelete = False
        'strF = "boolInclude = " & True
        strF = "boolInclude = -1"
        dv.RowFilter = strF
        dv.Sort = strS
        dgv.DataSource = dv

        'select first item
        If dv.Count = 0 Then
        Else
            'find first visible row
            int1 = 0
            For Count1 = 0 To dgv.ColumnCount - 1
                If dgv.Columns(Count1).Visible Then
                    int1 = Count1
                    Exit For
                End If
            Next
            dgv.CurrentCell = dgv.Rows(0).Cells(int1)
            dgv.Rows(int1).Selected = True
        End If

    End Sub

    Sub ReportTableHeaderFilter()
        Dim dv As System.Data.DataView
        Dim Count1 As Short
        Dim strF As String
        Dim var1, var2
        Dim strS As String
        Dim tbl As System.Data.DataTable

        tbl = tblConfigReportTables
        dv = tbl.DefaultView

        'filter according to contents of dgvReportTableConfiguration
        Dim dvR As System.Data.DataView
        dvR = frmH.dgvReportTableConfiguration.DataSource
        If dvR.Count = 0 Then
        Else
            strF = "(id_tblConfigReportType = 1000 OR id_tblConfigReportType = 2000) "
            For Count1 = 0 To dvR.Count - 1
                var1 = dvR(Count1).Item("id_tblConfigReportTables")
                var2 = dvR(Count1).Item("BOOLINCLUDE")
                If var2 Then
                    strF = strF & " OR id_tblConfigReportTables = " & var1 ' & " OR "
                End If
            Next
        End If

        'dv.RowFilter = strF

        ''''''''''''''console.writeline(strF)
        dv.RowFilter = strF
        dv.AllowNew = False
        dv.AllowEdit = False
        dv.AllowDelete = False
        strS = "intOrder ASC"
        dv.Sort = strS
        frmH.dgvReportTables.DataSource = dv

        Call ConfigReportTablesColumns()

    End Sub

    Sub ConfigReportTablesColumns()

        Dim dgv As DataGridView
        Dim Count1 As Short
        Dim int1 As Short
        Dim str1 As String
        Dim var1, var2

        dgv = frmH.dgvReportTables
        For Count1 = 0 To dgv.ColumnCount - 1
            dgv.Columns(Count1).Visible = False
        Next



        Try
            dgv.Columns("CHARTABLENAME").Visible = True
            dgv.Columns("CHARTABLENAME").HeaderText = "Table"
            dgv.Columns("CHARTABLENAME").SortMode = DataGridViewColumnSortMode.NotSortable

        Catch ex As Exception

        End Try

        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        'dgv.AutoResizeColumns()
        dgv.RowHeadersWidth = 25
        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        'dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        dgv.AutoResizeRows()



    End Sub

    Sub ConfigReportTableHeadersColumns()

        Dim dgv As DataGridView
        Dim Count1 As Short
        Dim int1 As Short
        Dim str1 As String
        Dim var1, var2

        If boolFormLoad Then
        Else
            'Exit Sub
        End If

        dgv = frmH.dgvReportTableHeaderConfig
        For Count1 = 0 To dgv.ColumnCount - 1
            dgv.Columns(Count1).Visible = False
        Next

        Try
            str1 = "CHARCOLUMNLABEL"
            dgv.Columns(str1).Visible = True
            dgv.Columns(str1).HeaderText = "Suggested Label"
            dgv.Columns(str1).ReadOnly = True
            dgv.Columns(str1).DisplayIndex = 0
            dgv.Columns(str1).SortMode = DataGridViewColumnSortMode.NotSortable

            str1 = "charUserLabel"
            dgv.Columns(str1).Visible = True
            dgv.Columns(str1).DisplayIndex = 1
            dgv.Columns(str1).HeaderText = "User Label **"
            dgv.Columns(str1).SortMode = DataGridViewColumnSortMode.NotSortable

            str1 = "intOrder"
            dgv.Columns(str1).Visible = True
            dgv.Columns(str1).DisplayIndex = 2
            dgv.Columns(str1).HeaderText = "Order"
            dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgv.Columns(str1).SortMode = DataGridViewColumnSortMode.NotSortable

            str1 = "boolInclude"
            dgv.Columns(str1).Visible = True
            dgv.Columns(str1).DisplayIndex = 3
            dgv.Columns(str1).HeaderText = "Include"
            dgv.Columns(str1).SortMode = DataGridViewColumnSortMode.NotSortable

        Catch ex As Exception

        End Try


        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        Try
            dgv.Columns("CHARCOLUMNLABEL").Width = dgv.Width * 0.32
            dgv.Columns("charUserLabel").Width = dgv.Width * 0.32
        Catch ex As Exception

        End Try

        dgv.RowHeadersWidth = 25
        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        'dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        dgv.AutoResizeRows()


    End Sub

    Sub CalcSampleCount()

        Dim dgv As DataGridView
        Dim Count1 As Short
        Dim int1 As Short
        Dim var1
        Dim intTot As Int16
        Dim intReport As Short
        Dim bool As Boolean
        Dim intB As Short

        dgv = frmH.dgvSampleReceipt
        int1 = dgv.Rows.Count
        intTot = 0
        intReport = 0
        For Count1 = 0 To int1 - 1
            intB = NZ(dgv("boolUse", Count1).Value, 0)
            If intB = 0 Then
                bool = False
            Else
                bool = True
            End If
            'bool = NZ(dgv("boolUse", Count1).Value, False)
            var1 = NZ(dgv("numSampleNumber", Count1).Value, 0)
            intTot = intTot + var1
            If bool Then
                intReport = intReport + var1
            End If
        Next

        frmH.txtSRecTotal.Text = intTot
        frmH.txtSRecTotalReport.Text = intReport


    End Sub

    Sub ShowSummaryTable()
        If boolFormLoad Then
            Exit Sub
        End If
        If frmH.cmdEdit.Enabled Then
            'Exit Sub
        End If


        'Call ConfigureSummaryTable()
        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim strF As String
        Dim id As Long

        dgv = frmH.dgvSummaryData
        dv = dgv.DataSource

        id = id_tblStudies

        If frmH.rbShowAllSummaryTable.Checked Then
            strF = "id_tblStudies = " & id & " AND (boolInclude = 0 OR boolInclude = -1)"
        ElseIf frmH.rbShowIncludedSummaryTable.Checked Then
            strF = "id_tblStudies = " & id & " AND boolInclude = -1"
        End If

        'dv.RowFilter = strF
        Try
            dv.RowFilter = strF
        Catch ex As Exception

        End Try

    End Sub


    Sub dgvReportStatementsCellContentClick(ByVal orig)
        If boolFormLoad Then
            Exit Sub
        End If
        'MsgBox(orig)
        Dim intRow As Short
        Dim intCol As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim intA As Short
        Dim intB As Short
        Dim intS As Short
        Dim intI As Short
        Dim boolGo As Boolean
        Dim boolGoI As Boolean
        Dim dgv As DataGridView
        Dim var1, var2
        Dim boolPB As Boolean

        'intRow = e.RowIndex
        intRow = frmH.dgvReportStatements.CurrentRow.Index
        intCol = frmH.dgvReportStatements.CurrentCell.ColumnIndex
        If oldCurrentRowRS = intRow Then
        Else
            'Call UpdateWord_dgv()

            'oldCurrentRowRS = intRow
        End If

        If frmH.cmdEdit.Enabled Then
            Exit Sub
        End If

        Dim int1 As Short

        'intCol = e.ColumnIndex
        boolGo = False
        boolGoI = False
        intA = -1
        intB = -1
        intI = -1
        int1 = frmH.dgvReportStatements.Columns.Count
        boolPB = False

        str1 = frmH.dgvReportStatements.Columns.Item(intCol).Name
        If StrComp(str1, "boolUStatements", CompareMethod.Text) = 0 Then
            boolGo = True
            str2 = "boolGW"
            str3 = "boolUseStatements"
            str4 = "boolGuWu"
        ElseIf StrComp(str1, "boolGW", CompareMethod.Text) = 0 Then
            boolGo = True
            str2 = "boolUStatements"
            str3 = "boolGuWu"
            str4 = "boolUseStatements"
        ElseIf StrComp(str1, "boolI", CompareMethod.Text) = 0 Then
            boolGoI = True
            str3 = "boolInclude"
        ElseIf StrComp(str1, "boolPB", CompareMethod.Text) = 0 Then
            boolPB = True
            str3 = "BOOLPAGEBREAK"
        End If

        Dim bool As Boolean

        If boolGo Then
            Dim dtbl As System.Data.DataTable
            dtbl = tblReportstatements
            dtbl.Columns.Item("charStatement").ReadOnly = False

            bool = frmH.dgvReportStatements.Rows.Item(intRow).Cells(str1).Value
            frmH.dgvReportStatements.Rows.Item(intRow).Cells(str2).Value = Not (bool)
            frmH.dgvReportStatements.Rows.Item(intRow).Cells("charStatement").Value = ""

            If bool Then
                frmH.dgvReportStatements.Rows.Item(intRow).Cells(str3).Value = -1
                frmH.dgvReportStatements.Rows.Item(intRow).Cells(str4).Value = 0
            Else
                frmH.dgvReportStatements.Rows.Item(intRow).Cells(str3).Value = 0
                frmH.dgvReportStatements.Rows.Item(intRow).Cells(str4).Value = -1
            End If

            frmH.dgvReportStatements.CommitEdit(DataGridViewDataErrorContexts.Commit)

            'frmh.dgvReportStatements.CommitEdit(DataGridViewDataErrorContexts.Commit)
            bool = frmH.dgvReportStatements.Rows.Item(intRow).Cells(str1).Value

            'dtbl.Columns.item("charStatement").ReadOnly = True
            Call ReportStatementChangeCbxFill()

        End If

        If boolGoI Then 'modify bound column cell values
            bool = frmH.dgvReportStatements.Rows.Item(intRow).Cells(str1).Value
            If bool Then
                frmH.dgvReportStatements.Rows.Item(intRow).Cells(str3).Value = -1
            Else
                frmH.dgvReportStatements.Rows.Item(intRow).Cells(str3).Value = 0
            End If
            frmH.dgvReportStatements.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If

        If boolPB Then 'modify bound column cell values
            bool = frmH.dgvReportStatements.Rows.Item(intRow).Cells(str1).Value
            If bool Then
                frmH.dgvReportStatements.Rows.Item(intRow).Cells(str3).Value = -1
            Else
                frmH.dgvReportStatements.Rows.Item(intRow).Cells(str3).Value = 0
            End If
            frmH.dgvReportStatements.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If

    End Sub


    Sub ResizeDV(ByVal dv As DataGridView, ByVal boolOne As Boolean)

        Dim ct3 As Short
        Dim Count2 As Short
        Dim int1 As Short

        If boolOne Then
            int1 = 0
        Else
            int1 = 1
        End If
        ct3 = dv.Columns.Count
        For Count2 = int1 To ct3 - 1
            'frmH.dgvCompanyAnalRef.Columns.item(Count2).SortMode = DataGridViewColumnSortMode.NotSortable
            dv.AutoResizeColumn(Count2, DataGridViewAutoSizeColumnMode.AllCells)
        Next
        dv.AllowUserToResizeColumns = True
        dv.AllowUserToResizeRows = True

        dv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

    End Sub

    Sub ResizeRows(ByVal dgv As DataGridView)

        Dim var11
        Dim int1 As Short
        Dim Count2 As Short

        int1 = dgv.Rows.Count
        dgv.Refresh()
        For Count2 = 0 To int1 - 1
            dgv.Rows.Item(Count2).Height = 18 'var11 * 0.8
        Next
        dgv.Refresh()

    End Sub


    Sub GenericDGVRowInsert(ByVal dgv As DataGridView, ByVal dtbl As System.Data.DataTable, ByVal strTable As String, ByVal strID As String)
        Dim int1 As Short
        Dim int2 As Short
        Dim intNRNumber As Short
        Dim dv As System.Data.DataView
        Dim intOrder As Short
        Dim Count1 As Short
        Dim ct1 As Short
        Dim drows() As DataRow
        Dim strS As String

        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim maxID, maxID1

        maxID = 1
        If Len(strTable) = 0 Then
        Else

            maxID = GetMaxID(strTable, 1, True) 'if maxid increment is 1, then getmaxid already does putmaxid
            'Call PutMaxID(strTable, maxID)

            'If boolGuWuOracle Then
            '    ta_tblMaxID.Fill(tblMaxID)
            'ElseIf boolGuWuAccess Then
            '    ta_tblMaxIDAcc.Fill(tblMaxID)
            'ElseIf boolGuWuSQLServer Then
            '    ta_tblMaxIDSQLServer.Fill(tblMaxID)
            'End If
            'strF = "charTable = '" & strTable & "'"
            'tbl = tblMaxID
            'rows = tbl.Select(strF)
            'maxID = rows(0).Item("NUMMAXID")
            'maxID1 = maxID
            'maxID = maxID + 1
            'rows(0).BeginEdit()
            'rows(0).Item("NUMMAXID") = maxID
            'rows(0).EndEdit()
            'If boolGuWuOracle Then
            '    Try
            '        ta_tblMaxID.Update(tblMaxID)
            '    Catch ex As DBConcurrencyException
            '        'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
            '    End Try
            'ElseIf boolGuWuAccess Then
            '    Try
            '        ta_tblMaxIDAcc.Update(tblMaxID)
            '    Catch ex As DBConcurrencyException
            '        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '    End Try
            'ElseIf boolGuWuSQLServer Then
            '    Try
            '        ta_tblMaxIDSQLServer.Update(tblMaxID)
            '    Catch ex As DBConcurrencyException
            '        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '    End Try
            'End If

        End If

        dv = dgv.DataSource
        ct1 = dv.Count ' tbl.Rows.Count

        If dgv.SelectedRows.Count = 0 And ct1 = 0 Then
            int1 = 0
            intOrder = 1
        ElseIf dgv.SelectedRows.Count = 0 And ct1 > 0 Then
            int1 = ct1
            intOrder = ct1 + 1
        Else
            int1 = dgv.CurrentRow.Index + 1
            intOrder = dgv.CurrentRow.Index + 2
        End If

        strF = "id_tblStudies = " & id_tblStudies
        drows = dtbl.Select(strF, "intOrder ASC")
        'renumber rows
        For Count1 = int1 To ct1 - 1
            drows(Count1).BeginEdit()
            int2 = drows(Count1).Item("intOrder")
            drows(Count1).Item("intOrder") = int2 + 1
            drows(Count1).EndEdit()
        Next
        'now add row
        Dim row As DataRow = dtbl.NewRow()
        row.BeginEdit()
        If Len(strTable) = 0 Then
        Else
            row.Item(strID) = maxID
        End If
        row.Item("id_tblStudies") = id_tblStudies
        row.Item("intOrder") = intOrder
        row.EndEdit()
        dtbl.Rows.Add(row)

        dv = dtbl.DefaultView
        strF = "id_tblStudies = " & id_tblStudies
        strS = "intOrder ASC"
        dv.RowFilter = strF
        dv.Sort = strS
        dv.AllowNew = False
        dv.AllowEdit = True
        dv.AllowDelete = False
        dgv.DataSource = dv


    End Sub

    Sub GenericDGVRowDelete(ByVal dgv As DataGridView)
        Dim Count1 As Short
        Dim ct1 As Short
        Dim ct2 As Short
        Dim srow As DataGridViewSelectedRowCollection
        Dim dv As System.Data.DataView
        Dim int1 As Short

        ct2 = dgv.SelectedRows.Count

        If ct2 = 0 Then
            MsgBox("Action could not be completed. Remember to select one or more entire row in order to delete rows.", MsgBoxStyle.Information, "No rows selected to delete...")
            Exit Sub
        End If

        int1 = dgv.CurrentRow.Index
        dv = dgv.DataSource
        srow = dgv.SelectedRows
        If srow.Count = 0 Then 'don't do anything
            MsgBox("No rows have been selected.", MsgBoxStyle.Information, "Select some rows...")
        Else
            For Count1 = ct2 - 1 To 0 Step -1
                int1 = srow(Count1).Index
                dv(int1).Row.Delete()
            Next

            'now reorder items
            ct1 = dgv.Rows.Count
            For Count1 = 0 To ct1 - 1
                dv(Count1).BeginEdit()
                dv(Count1).Item("intOrder") = Count1 + 1
                dv(Count1).EndEdit()
            Next
        End If


    End Sub

    Sub DoCancelSampleReceipt()

        frmH.dgvSampleReceipt.CommitEdit(DataGridViewDataErrorContexts.Commit)

        tblSampleReceipt.RejectChanges()
        Call SampleReceiptChange()

        Call ReorderSRec() 'funny that this needs to be called

    End Sub

    Sub SetReportConfigType()

        Dim int1 As Short
        Dim int2 As Short
        Dim strF As String
        Dim rows() As DataRow
        Dim tbl As System.Data.DataTable
        Dim dgv As DataGridView

        dgv = frmH.dgvReports
        If dgv.RowCount = 0 Then
            id_tblConfigReportType = -1
        Else
            If dgv.CurrentRow Is Nothing Then
                int1 = 0
            Else
                int1 = dgv.CurrentRow.Index
            End If
        End If

        Try
            int2 = dgv.Item("ID_TBLCONFIGREPORTTYPE", int1).Value
            id_tblConfigReportType = int2
        Catch ex As Exception
            id_tblConfigReportType = -1
        End Try

        If id_tblConfigReportType > 1 And id_tblConfigReportType < 5 Then 'is validation
            frmH.gbMethValApplyGuWu.Visible = False
            frmH.cmdMethValUpdate.Visible = True
            frmH.dgvMethodValData.Height = (frmH.gbMethValApplyGuWu.Top + frmH.gbMethValApplyGuWu.Height) - frmH.dgvMethodValData.Top
        Else
            frmH.gbMethValApplyGuWu.Visible = True
            frmH.cmdMethValUpdate.Visible = True
            frmH.dgvMethodValData.Height = (frmH.gbMethValApplyGuWu.Top) - frmH.dgvMethodValData.Top - 5
        End If


    End Sub

    Sub FillTableStuffMethVal(ByVal boolStabilityOnly As Boolean)

        'this must be called AFTER tblAnalysisResultsHome gets prepared
        '20190208 LEE:
        'This updates dgvMethodValData
        'called by:
        '  cmdmethValUpdate click
        '  Report Table config Save - this updates dgvMethodValData with stability conditions statements
        '  FillMethValExistingGuWu

        Dim dtbl1 As System.Data.DataTable
        Dim dtbl2 As System.Data.DataTable
        Dim dtbl3 As System.Data.DataTable
        Dim dtblD As System.Data.DataTable
        Dim rowsD() As DataRow
        Dim dvD As System.Data.DataView
        Dim strF1 As String
        Dim strF2 As String
        Dim strF3 As String
        Dim strFD As String
        Dim intRows As Short
        Dim intCols As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim boolGo As Boolean
        Dim arrA(20, 3)
        Dim int1 As Int16
        Dim int2 As Int16
        Dim int3 As Int16
        Dim int4 As Int16
        Dim int5 As Int16

        Dim strV As String
        Dim var1, var2, var3, var4, var5, var6
        Dim idR As Short
        Dim idS2 As Int64
        Dim dgvD As DataGridView
        Dim strAnal As String
        Dim strIS As String
        Dim boolE As Boolean

        dgvD = frmH.dgvMethodValData
        intRows = dgvD.Rows.Count
        dvD = dgvD.DataSource

        dtbl1 = tblAssignedSamples
        dtbl2 = tblAssignedSamplesHelper
        dtbl3 = tblConfigReportTables
        dtblD = tblMethodValidationData

        'DON'T LOOK ANYMORE
        'Dim con As New ADODB.Connection
        'Dim rs As New ADODB.Recordset
        'rs.CursorLocation = CursorLocationEnum.adUseClient

        Dim strACon As String

        'strFD = "ID_TBLSTUDIES = " & id_tblStudies & " AND INTCOLUMNNUMBER = " & Count2
        'rowsD = dtblD.Select(strFD)
        'var1 = NZ(rowsD(0).Item("CHARARCHIVEPATH"), "")
        'If Len(var1) = 0 Then
        '    con.Open(constrCur)
        'Else
        '    strACon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & var1 & ";"
        '    con.Open(strACon)
        'End If

        '''''''''''console.writeline("Open con: " & Now)

        strFD = "ID_TBLSTUDIES = " & id_tblStudies & " AND INTCOLUMNNUMBER = 1" ' & Count2
        rowsD = dtblD.Select(strFD)
        var1 = NZ(rowsD(0).Item("CHARARCHIVEPATH"), "")

        boolE = True
        'If Len(var1) = 0 Then
        '    con.Open(constrCur)
        'Else
        '    strACon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & var1 & ";"
        '    'ensure var1 exists
        '    If System.IO.File.Exists(var1) Then
        '        con.Open(strACon)
        '        boolE = True
        '    Else
        '        boolE = False
        '    End If
        'End If

        boolE = False

        intCols = ctAnalytes
        '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
        '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
        '10=UseIntStd, 11=IntStd, 12=MasterAssayID, 13=Original AnalyteDescription
        ReDim arrA(20, intCols)
        '1=analytedescription,2=analyteid, 3=analyteindex
        int1 = 0
        For Count1 = 1 To dgvD.Columns.Count - 1
            str1 = dgvD.Columns(Count1).HeaderText
            For Count2 = 1 To ctAnalytes
                str2 = arrAnalytes(1, Count2)
                If StrComp(str1, str2, CompareMethod.Text) = 0 Then
                    int1 = int1 + 1
                    For Count3 = 1 To 13
                        arrA(Count3, int1) = arrAnalytes(Count3, Count2)
                    Next
                    Exit For
                End If
            Next
        Next

        ' 1  Summary of Analytical Runs
        ' 2  Summary of Regression Constants
        ' 3  Summary of Back-Calculated Calibration Std Conc
        ' 4  Summary of Interpolated QC Std Conc
        ' 5  Summary of Samples
        ' 6  Summary of Reassayed Samples
        ' 7  Summary of Repeat Samples
        ' 11  Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision
        ' 12  Summary of Interpolated Dilution QC Concentrations
        ' 13  Summary of Combined Recovery
        ' 14  Summary of True Recovery
        ' 15  Summary of Suppression/Enhancement
        ' 17  Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments
        ' 18  Summary of [Period Temp] Stability in Matrix
        ' 19  Summary of Freeze/Thaw [#Cycles] Stability in Matrix
        ' 21  [Period Temp] Final Extract Stability of Interpolated QC Std Concentrations
        ' 22  [Period Temp] Stock Solution Stability Assessment
        ' 23  [Period Temp] Spiking Solution Stability Assessment
        ' 29  [Period Temp] Long-Term QC Std Storage Stability
        ' 30  Incurred Samples
        ' 1000  QA Events Table Columns (Events/Management)
        ' 2000  QA Events Table Rows (Critical Phases)

        '''''''''''console.writeline("Start FillInterIntraQCStats: " & Now)

        Call FillInterIntraQCStats()

        Dim dtblT As System.Data.DataTable
        dtblT = tblReportTable

        '''''''''''console.writeline("Start block: " & Now)
        Dim boolDo As Boolean
        Dim boolProp As Boolean = False '20190215 LEE:
        Dim strCol As String
        Dim strField As String

        For Count1 = 0 To intRows - 1
            'str1 = dgvD(0, Count1).Value
            strCol = dgvD(0, Count1).Value
            dvD(Count1).BeginEdit()

            For Count2 = 1 To intCols
                strAnal = arrA(1, Count2)
                strIS = NZ(arrA(11, Count2), "NA")
                strFD = "ID_TBLSTUDIES = " & id_tblStudies & " AND INTCOLUMNNUMBER = " & Count2
                'rowsD = dtblD.Select(strFD)
                'idS2 = NZ(rowsD(0).Item("ID_TBLSTUDIES"), -1)

                strV = ""

                boolDo = True
                boolProp = False
                strField = ""

                If boolStabilityOnly Then
                    Select Case strCol
                        Case "Freeze/Thaw Stability"
                            'strV = GetStabPeriod(dtblT, id_tblStudies, 19, strCol)
                            
                        Case "Maximum # of Freeze/thaw Cycles"
                            boolProp = True
                            strField = "CHARMAXNUMBERFREEZETHAW" '20190215 LEE: this is 
                        Case "Stability under Storage Conditions" 'deprecated, is now benchtop stability 20181110
                        Case "Bench-top Stability"
                            'strV = GetStabPeriod(dtblT, id_tblStudies, 21, strCol)
                            'Case "Is Stability >= Maximum Storage Duration"
                        Case "Process Stability"
                            'strV = GetStabPeriod(dtblT, id_tblStudies, 21, strCol)
                            'Case "Refrigerated Stability in Matrix" 'deprecated now Reinjection Stability 20181110
                        Case "Reinjection Stability"
                            'strV = GetStabPeriod(dtblT, id_tblStudies, 18, strCol)
                        Case "Long-term Storage Stability"
                            'strV = GetStabPeriod(dtblT, id_tblStudies, 29, strCol)

                            '20190109 LEE:
                        Case "Whole Blood Stability"
                            'strV = GetStabPeriod(dtblT, id_tblStudies, 29, strCol)
                        Case "Stock Solution Stability"
                            'strV = GetStabPeriod(dtblT, id_tblStudies, 29, strCol)
                        Case "Spiking Solution Stability"
                            'strV = GetStabPeriod(dtblT, id_tblStudies, 29, strCol)
                        Case "Autosampler Stability"
                            'strV = GetStabPeriod(dtblT, id_tblStudies, 29, strCol)
                        Case "Batch Reinjection Stability"


                        Case Else
                            boolDo = False
                    End Select

                    If boolDo Then
                        strV = GetStabPeriod(dtblT, id_tblStudies, 0, strCol)
                    End If

                Else

                    '20190206 LEE
                    Dim dvAA As DataView = frmH.dgvWatsonAnalRef.DataSource

                    Select Case strCol
                        Case "Validation Corporate Study/Project Number"
                        Case "Validation Protocol Number"
                        Case "Validation Report Title"
                        Case "Validation Report Number"
                        Case "Lab Method Title"
                        Case "Lab Method Number"
                            'Case "Analytical Method Type" '20190212 LEE: deprecated
                        Case "Assay Technique" '20190212 LEE:
                        Case "Sponsor Method Validation Study Number"
                        Case "Sponsor Method Validation Title"
                        Case "Extraction Procedure Description"
                        Case "Sample Size Units"
                        Case "Sample Size"
                        Case "Species"
                        Case "Anticoagulant/Preservative"
                        Case "Matrix"
                        Case "Maximum Run Size"
                        Case "QC Concentrations"
                            idR = 11 '
                            If boolE Then
                                'strV = GetQCs(con, arrA, Count2, idR, idS2)
                            End If

                            'Dim dvAA As DataView = frmH.dgvWatsonAnalRef.DataSource

                            If Count2 > frmH.dgvWatsonAnalRef.ColumnCount - 1 Then
                                strV = ""
                            Else
                                int2 = FindRowDV("ULOQ Units", dvAA)
                                int3 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
                                int5 = FindRowDV("Is Internal Standard?", dvAA)
                                strV = "None"

                                '20190206 LEE
                                'evalutate count2
                                str3 = NZ(dvAA(int5).Item(Count2), "")
                                If StrComp(str3, "Yes", CompareMethod.Text) = 0 Then
                                    Exit For
                                End If

                                var4 = dvAA.Item(int2).Item(Count2)
                                str1 = NZ(frmH.dgvStudyConfig(1, int3).Value, "")

                                If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
                                Else
                                    var4 = str1
                                End If

                                If IsDBNull(var4) Then
                                    var4 = "[NA]"
                                ElseIf Len(var4) = 0 Or StrComp(NZ(var4, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                                    var4 = "[NA]"
                                Else
                                    var4 = var4
                                End If

                                var1 = strAnal
                                var2 = ReturnQCStds(CStr(var1), False)
                                str1 = var2 & " " & var4 ' & " for " & var1
                                strV = str1
                            End If
                           

                        Case "Standard Curve Concentrations"
                            'If boolE Then
                            '    'strV = GetCalibrStds(con, arrA, Count2, idR, idS2)
                            'End If

                            'Dim dvAA As DataView = frmH.dgvWatsonAnalRef.DataSource
                            If Count2 > frmH.dgvWatsonAnalRef.ColumnCount - 1 Then
                                strV = ""
                            Else
                                int2 = FindRowDV("ULOQ Units", dvAA)
                                int3 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
                                int5 = FindRowDV("Is Internal Standard?", dvAA)
                                strV = "None"

                                str3 = NZ(dvAA(int5).Item(Count2), "")
                                If StrComp(str3, "Yes", CompareMethod.Text) = 0 Then
                                    Exit For
                                End If

                                var4 = dvAA.Item(int2).Item(Count2)
                                str1 = NZ(frmH.dgvStudyConfig(1, int3).Value, "")

                                If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
                                Else
                                    var4 = str1
                                End If

                                If IsDBNull(var4) Then
                                    var4 = "[NA]"
                                ElseIf Len(var4) = 0 Or StrComp(NZ(var4, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                                    var4 = "[NA]"
                                Else
                                    var4 = var4
                                End If

                                var1 = strAnal
                                var2 = ReturnCalibrStds(CStr(var1), Count2, False)
                                str1 = var2 & " " & var4 ' & " for " & var1
                                strV = str1

                            End If

                        Case "Lower Limit of Quantification"
                            strV = GetLLOQ(Count2, strAnal)
                        Case "Upper Limit of Quantification"
                            strV = GetULOQ(Count2, strAnal)
                        Case "Average Recovery of Analyte (%)"
                            strV = GetRecovery(dtblT, id_tblStudies, 14, strAnal, False, strIS)
                        Case "Average Recovery of IS (%)"
                            If StrComp(strIS, "NA", CompareMethod.Text) = 0 Then
                                strV = "NA"
                            Else
                                strV = GetRecovery(dtblT, id_tblStudies, 14, strIS, True, strIS)
                            End If
                        Case "Inter-day QC Accuracy Range (%Bias)"
                            strV = GetInterAccMin(strAnal) & " to " & GetInterAccMax(strAnal)
                        Case "Inter-day QC Precision Range (%CV)"
                            strV = GetInterPrecMin(strAnal) & " to " & GetInterPrecMax(strAnal)
                        Case "Intra-day QC Accuracy Range (%Bias)"
                            strV = GetIntraAccMin(strAnal) & " to " & GetIntraAccMax(strAnal)
                        Case "Intra-day QC Precision Range (%CV)"
                            strV = GetIntraPrecMin(strAnal) & " to " & GetIntraPrecMax(strAnal)
                        Case "Freeze/Thaw Stability"
                            strV = NZ(rowsD(0).Item("CHARDEMONSTRATEDFREEZETHAW"), "")
                            If Len(strV) = 0 Then
                                'try getstability
                                strV = GetStabPeriod(dtblT, id_tblStudies, 19, strCol) '3rd parameter no longer used
                            End If
                        Case "Maximum # of Freeze/thaw Cycles"
                            strV = rowsD(0).Item("CHARMAXNUMBERFREEZETHAW")
                            If Len(strV) = 0 Then
                                'try getstability
                                strV = GetStabPeriod(dtblT, id_tblStudies, 19, strCol) '3rd parameter no longer used
                            End If
                        Case "Stability under Storage Conditions" 'deprecated now Benchtop stability 20181110
                        Case "Bench-top Stability"
                            'strV = GetStabPeriod(dtblT, id_tblStudies, 21, str1)
                            strV = NZ(rowsD(0).Item("CHARSTABILITYUNDERSTORAGECOND"), "")
                            If Len(strV) = 0 Then
                                'try getstability
                                strV = GetStabPeriod(dtblT, id_tblStudies, 19, strCol) '3rd parameter no longer used
                            End If
                        Case "Is Stability >= Maximum Storage Duration"
                            strV = rowsD(0).Item("CHARSTABILITYMAXSTORAGEDUR")
                        Case "Process Stability"
                            'strV = GetStabPeriod(dtblT, id_tblStudies, 21, str1)
                            strV = NZ(rowsD(0).Item("CHARPROCSTABILITY"), "")
                            If Len(strV) = 0 Then
                                'try getstability
                                strV = GetStabPeriod(dtblT, id_tblStudies, 19, strCol) '3rd parameter no longer used
                            End If
                        Case "Refrigerated Stability in Matrix" 'deprecated now Reinjection Stability 20181110
                        Case "Reinjection Stability"
                            'strV = GetStabPeriod(dtblT, id_tblStudies, 18, str1)
                            strV = NZ(rowsD(0).Item("CHARREFRSTAB"), "")
                            If Len(strV) = 0 Then
                                'try getstability
                                strV = GetStabPeriod(dtblT, id_tblStudies, 19, strCol) '3rd parameter no longer used
                            End If
                        Case "Long-term Storage Stability"
                            strV = NZ(rowsD(0).Item("CHARLTSTORSTAB"), "")
                            If Len(strV) = 0 Then
                                'try getstability
                                strV = GetStabPeriod(dtblT, id_tblStudies, 19, strCol) '3rd parameter no longer used
                            End If

                            '20190109 LEE:
                        Case "Whole Blood Stability"
                            strV = NZ(rowsD(0).Item("CHARBLOOD"), "")
                            If Len(strV) = 0 Then
                                'try getstability
                                strV = GetStabPeriod(dtblT, id_tblStudies, 19, strCol) '3rd parameter no longer used
                            End If
                        Case "Stock Solution Stability"
                            strV = NZ(rowsD(0).Item("CHARSTOCKSOLUTION"), "")
                            If Len(strV) = 0 Then
                                'try getstability
                                strV = GetStabPeriod(dtblT, id_tblStudies, 19, strCol) '3rd parameter no longer used
                            End If
                        Case "Spiking Solution Stability"
                            strV = NZ(rowsD(0).Item("CHARSPIKING"), "")
                            If Len(strV) = 0 Then
                                'try getstability
                                strV = GetStabPeriod(dtblT, id_tblStudies, 19, strCol) '3rd parameter no longer used
                            End If
                        Case "Autosampler Stability"
                            strV = NZ(rowsD(0).Item("CHARAUTOSAMPLER"), "")
                            If Len(strV) = 0 Then
                                'try getstability
                                strV = GetStabPeriod(dtblT, id_tblStudies, 19, strCol) '3rd parameter no longer used
                            End If
                        Case "Batch Reinjection Stability"
                            strV = NZ(rowsD(0).Item("CHARBATCHREINJECTION"), "")
                            If Len(strV) = 0 Then
                                'try getstability
                                strV = GetStabPeriod(dtblT, id_tblStudies, 19, strCol) '3rd parameter no longer used
                            End If


                        Case "Dilution Integrity"
                            '
                            strV = rowsD(0).Item("CHARDILINTEGR")
                        Case "Analyte Selectivity"
                            strV = rowsD(0).Item("CHARANALSELECT")
                        Case "Internal Standard Selectivity"
                            strV = rowsD(0).Item("CHARISSELECT")
                    End Select
                End If

                If Len(strV) = 0 Or boolDo = False Then
                    If Len(strV) = 0 Then 'get from charstability values
                        strV = GetStability()
                    End If
                Else
                    dvD(Count1).Item(Count2) = strV
                End If

            Next

            dvD(Count1).EndEdit()
        Next

        '''''''''''console.writeline("End block: " & Now)

        'If con.State = ADODB.ObjectStateEnum.adStateOpen Then
        '    con.Close()
        'End If

        'con = Nothing

        '''''''''''console.writeline("Close con: " & Now)

    End Sub


    Function GetStability() As String

        Dim dtbl As DataTable = tblReportTable
        Dim strF As String



    End Function

    Function GetReportType()

        Dim int1 As Short
        Dim int2 As Short
        Dim strF As String
        Dim rows() As DataRow
        Dim tbl As System.Data.DataTable
        Dim dgv As DataGridView

        GetReportType = -1

        dgv = frmH.dgvReports
        If dgv.RowCount = 0 Then
            id_tblConfigReportType = -1
        Else
            If dgv.CurrentRow Is Nothing Then
                int1 = 0
            Else
                int1 = dgv.CurrentRow.Index
            End If
        End If

        Try
            int2 = dgv.Item("ID_TBLCONFIGREPORTTYPE", int1).Value
            id_tblConfigReportType = int2
        Catch ex As Exception
            id_tblConfigReportType = -1
        End Try

        GetReportType = id_tblConfigReportType

    End Function

    Sub SelectedRefresh()

        Dim str1 As String
        Dim str2 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim dgv As DataGridView

        Cursor.Current = Cursors.WaitCursor

        str1 = frmH.lbxTab1.SelectedItem
        frmH.dgvCompanyAnalRef.Columns.Item("Item").Frozen = False
        'If frmH.rbRBS_Col.Checked Then
        '    frmH.dgvReportStatements.Columns.Item("CHARHEADINGTEXT").Frozen = False
        'Else
        '    frmH.dgvReportStatements.Columns.Item("charSectionName").Frozen = False
        'End If

        Try
            If frmH.rbRBS_Col.Checked Then
                frmH.dgvReportStatements.Columns.Item("CHARHEADINGTEXT").Frozen = False
            Else
                frmH.dgvReportStatements.Columns.Item("charSectionName").Frozen = False
            End If

        Catch ex As Exception
            Exit Sub
        End Try

        Cursor.Current = Cursors.WaitCursor

        'address pesky items
        Call ReorderSRec()

        Cursor.Current = Cursors.WaitCursor

        If StrComp(str1, "Analytical Reference Standard", CompareMethod.Text) = 0 Then
            str2 = AnalRefHook()
            If Len(str2) > 0 Then
                Select Case str2
                    Case Is = "CRLWor_AnalRefStandard"
                        Call ComboBoxCRLAnalRefFill()
                End Select
            End If
            Call ResizeRows(frmH.dgvCompanyAnalRef)
            Call ResizeRows(frmH.dgvWatsonAnalRef)
            Call HideAnalRefRows()

            frmH.dgvCompanyAnalRef.Columns.Item("Item").Frozen = True

        End If

        Cursor.Current = Cursors.WaitCursor

        If StrComp(str1, "Method Validation Data", CompareMethod.Text) = 0 Then
            Call SetReportConfigType()
        End If

        If StrComp(str1, "Report Body Sections", CompareMethod.Text) = 0 Then

            boolRSCFill = True
            boolRSCFill = False
            Cursor.Current = Cursors.WaitCursor

            Call OrderReportStatementCol()

            Cursor.Current = Cursors.WaitCursor

            'If frmH.rbRBS_Col.Checked Then
            '    frmH.dgvReportStatements.Columns.Item("CHARHEADINGTEXT").Frozen = True
            'Else
            '    frmH.dgvReportStatements.Columns.Item("charSectionName").Frozen = True
            'End If

            Cursor.Current = Cursors.WaitCursor

            Call SetRBSCmds()

            Cursor.Current = Cursors.WaitCursor


        End If

        'If StrComp(str1, "Appendices", CompareMethod.Text) = 0 Then
        '    'Call AppendixUpdateCB()
        'End If

        If StrComp(str1, "QA Event Table", CompareMethod.Text) = 0 Then
            'configure qa table
            Call QATableInitialize()
        End If
        Cursor.Current = Cursors.WaitCursor

        'If StrComp(str1, "Report Table Header Configuration", CompareMethod.Text) = 0 Then
        If StrComp(str1, "Configure Column Headings", CompareMethod.Text) = 0 Then
            'filter ReportTableHeader appropriately
            Try
                Call ReportTableHeaderFilter()
                Call ReportTableHeaderConfigPopulate()
            Catch ex As Exception

            End Try
        End If
        Cursor.Current = Cursors.WaitCursor

        If StrComp(str1, "Report Table Configuration", CompareMethod.Text) = 0 Then
            'Call AssessSampleAssignment()'this gets called in rtfilter
            Call RTFilter()
            Call SizecmdOrder(frmH.dgvReportTableConfiguration, frmH.cmdOrderReportTableConfig, "INTORDER")
        End If
        Cursor.Current = Cursors.WaitCursor

        'If StrComp(str1, "Add/Edit Top Level Data", CompareMethod.Text) = 0 Then
        '    'configure a comboboxcell for dgvDataCompany
        '    Dim dvD as system.data.dataview
        '    'dvD = frmH.dgvDataCompany.DataSource
        '    dvD = frmH.dgvStudyConfig.DataSource
        '    Try
        '        int1 = FindRowDV("Table Date Format", dvD)
        '        Dim cbx As New DataGridViewComboBoxCell
        '        cbx = cbxDateFormat.Clone
        '        cbx.AutoComplete = True
        '        cbx.MaxDropDownItems = 20
        '        cbx.DisplayStyleForCurrentCellOnly = True
        '        cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
        '        frmH.dgvStudyConfig("Value", int1) = cbx
        '    Catch ex As Exception

        '    End Try

        '    Try
        '        int1 = FindRowDV("Text Date Format", dvD)
        '        Dim cbx1 As New DataGridViewComboBoxCell
        '        cbx1 = cbxDateFormat.Clone
        '        cbx1.AutoComplete = True
        '        cbx1.MaxDropDownItems = 20
        '        cbx1.DisplayStyleForCurrentCellOnly = True
        '        cbx1.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
        '        frmH.dgvStudyConfig("Value", int1) = cbx1
        '    Catch ex As Exception

        '    End Try

        '    frmH.dgvStudyConfig.Refresh()

        'End If
        Cursor.Current = Cursors.WaitCursor

        If StrComp(str1, "Choose Study & Report", CompareMethod.Text) = 0 Then
            Call OrderReportsHome()
        End If
        Cursor.Current = Cursors.WaitCursor

        'fill bool checkboxes in Contributing Personnel
        If StrComp(str1, "Contributing Personnel", CompareMethod.Text) = 0 Then
            Call CPDisplayOrder()
            Call UpdateCPBool()
        End If
        Cursor.Current = Cursors.WaitCursor

        If StrComp(str1, "Summary Table", CompareMethod.Text) = 0 Then
            Call OrderSummaryTable()
            Call UpdateBoolSummaryTable()
        End If

        If StrComp(str1, "Sample Receipt", CompareMethod.Text) = 0 Then
            frmH.dgvSampleReceipt.AutoResizeColumns()
        End If
        Cursor.Current = Cursors.WaitCursor

        'check for hook
        Call HookAnalysis()
        Cursor.Current = Cursors.WaitCursor

        'set focus back to lbx
        frmH.lbxTab1.Select()

        If boolFormLoad Then
        Else
            If StrComp(str1, "Appendices and Figures", CompareMethod.Text) = 0 Then
                'Call frmH.OpenAppFig()
            End If

            If StrComp(str1, "Sample/QC/Calibr Std Details", CompareMethod.Text) = 0 Then
                'Call frmH.OpenAnalDetails()

                'int1 = frmH.lbxTab1.SelectedIndex
                'int2 = frmH.tab1.TabIndex
                'If int1 = int2 Then
                'Else
                '    frmH.tab1.SelectedIndex = int1
                '    frmH.tab1.Refresh()
                'End If
            End If
        End If

        frmH.dgvMethodValData.AutoResizeRows()

        Cursor.Current = Cursors.Default

    End Sub

    Sub CorrectSampleReceipt(ByVal boolWatson, ByVal boolManual)

        Dim bool As Boolean
        Dim int1 As Short
        Dim Count1 As Short
        Dim strField As String

        Dim dv As System.Data.DataView
        dv = frmH.dgvSampleReceipt.DataSource
        int1 = dv.Count

        If boolWatson Then 'correct boolUseWatson field
            strField = "boolUseWatson"
            If frmH.chkUseWatsonSampleNumber.CheckState = CheckState.Checked Then
                bool = True
            Else
                bool = False
            End If

        ElseIf boolManual Then 'correct boolUseManual field
            strField = "boolUseManual"
            If frmH.chkManualSampleNumber.CheckState = CheckState.Checked Then
                bool = True
            Else
                bool = False
            End If

        End If

        Dim intS As Short
        If bool Then
            intS = -1
        Else
            intS = 0
        End If

        For Count1 = 0 To int1 - 1
            dv(Count1).BeginEdit()
            dv(Count1).Item(strField) = intS
            dv(Count1).EndEdit()
        Next

    End Sub

    Sub OrderDGV(ByVal dgv As DataGridView, ByVal strS As String, ByVal strID As String)

        Dim dv As System.Data.DataView
        Dim Count1 As Short
        Dim Count2 As Short
        Dim intRows As Short
        Dim tbl As New System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim var1, var2, var3, var4, var5, var6
        'Dim oStrSort

        Dim col1 As New DataColumn
        col1.ColumnName = "ID"
        tbl.Columns.Add(col1)
        Dim col2 As New DataColumn
        col2.ColumnName = "intOldOrder"
        tbl.Columns.Add(col2)
        Dim col3 As New DataColumn
        col3.ColumnName = "intNewOrder"
        tbl.Columns.Add(col3)

        dv = dgv.DataSource
        intRows = dv.Count
        dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)

        'oStrSort = dv.Sort
        'MsgBox(oStrSort)

        'dv.Sort = strS

        frmH.Cursor.Current = Cursors.WaitCursor

        'record id and order
        For Count1 = 0 To intRows - 1
            'var1 = dv(Count1).Item(strID)
            var1 = dgv.Item(strID, Count1).Value
            Dim row As DataRow = tbl.NewRow
            row("ID") = CStr(var1)
            row("intOldOrder") = dgv.Item("INTORDER", Count1).Value
            row("intNewOrder") = Count1 + 1
            tbl.Rows.Add(row)
        Next

        'turn off dv sort
        'dv.Sort = Nothing

        Dim intRow As Int16
        'now re-order columns
        For Count1 = 0 To intRows - 1
            var1 = tbl.Rows.Item(Count1).Item("ID")
            var2 = tbl.Rows.Item(Count1).Item("intOldOrder")
            var3 = tbl.Rows.Item(Count1).Item("intNewOrder")
            'dgv.Item(strS, Count1).Value = var2
            'dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            intRow = FindRowInDGV(strID, var1, dgv)

            dv(intRow).BeginEdit()
            dv(intRow).Item(strS) = var3
            dv(intRow).EndEdit()

            'For Count2 = 0 To dv.Count - 1
            '    var4 = dv(Count2).Item(strID)
            '    If var1 = var4 Then
            '        dv(Count2).BeginEdit()
            '        dv(Count2).Item(strS) = var3
            '        dv(Count2).EndEdit()
            '        Exit For
            '    End If
            'Next
        Next

        dgv.AutoResizeRows()

        'Exit Sub

        'turn sort back on
        'dv.Sort = "INTORDER ASC"

        'dgv.AutoResizeRows()

        frmH.Cursor.Current = Cursors.Default

    End Sub

    Sub RTFilter()

        If boolFormLoad Then
            Exit Sub
        End If

        Dim dv As System.Data.DataView
        Dim strF, strTxtFilter As String
        Dim dtbl As System.Data.DataTable

        dtbl = tblReportTables

        'debug
        Dim int1 As Short
        Dim Count1 As Short
        Dim str1 As String
        int1 = dtbl.Columns.Count
        For Count1 = 0 To int1 - 1
            str1 = dtbl.Columns(Count1).ColumnName
            str1 = str1
        Next

        strTxtFilter = frmH.txtFilterSamples.Text

        If frmH.cmdEdit.Enabled Then
        Else
            If StrComp(strTxtFilter, "", CompareMethod.Text) = 0 Then
                frmH.cmdClearFilters.Enabled = False
                frmH.cmdRTCUp.Enabled = True
                frmH.cmdRTCDown.Enabled = True
                frmH.lblRTCUpDown.ForeColor = Color.White
            Else
                frmH.cmdClearFilters.Enabled = True
                frmH.cmdRTCUp.Enabled = False
                frmH.cmdRTCDown.Enabled = False
                frmH.lblRTCUpDown.ForeColor = Color.LightGray
            End If
        End If

        dv = dtbl.DefaultView
        dv.Sort = "INTORDER ASC"
        If frmH.rbShowIncludedRTConfig.Checked Then
            strF = "BOOLINCLUDE =  " & True 'leave as TRUE. Underlying cache table has boolean field
            strF = strF & String.Format(" AND {0} LIKE '%{1}%'", "CHARHEADINGTEXT", strTxtFilter)
        Else
            strF = strF & String.Format("{0} LIKE '%{1}%'", "CHARHEADINGTEXT", strTxtFilter)
        End If
        dv.RowFilter = strF
        dv.AllowEdit = True
        dv.AllowNew = False
        dv.AllowDelete = False
        'frmh.dgReportTableConfiguration.DataSource = dv
        frmH.dgvReportTableConfiguration.DataSource = dv

        Call AssessSampleAssignment()

        frmH.dgvReportTableConfiguration.AutoResizeRows()



    End Sub

    Sub RBFilter()

        If boolFormLoad Then
            Exit Sub
        End If
        If boolCont Then
        Else
            Exit Sub
        End If

        'Dim dv as system.data.dataview
        Dim strF As String
        Dim tbl As System.Data.DataTable
        Dim dtbl As System.Data.DataTable
        Dim strS As String
        Dim dgv As DataGridView
        Dim num1
        Dim rows() As DataRow
        Dim strRBSFilter As String 'Company filter
        Dim strRBSTypeFilter As String 'Report Type
        Dim int1 As Short
        Dim boolCI As Boolean
        Dim str1 As String
        Dim numRBSTypeFilter As Short
        Dim numRBSFilter As Short

        Cursor.Current = Cursors.WaitCursor

        boolStopRBS = True

        strRBSFilter = NZ(frmH.cbxRBSFilter.Text, "[None]")
        strRBSTypeFilter = NZ(frmH.cbxRBSTypeFilter.Text, "[None]")

        'find ID_TBLCONFIGREPORTTYPE	
        If StrComp(strRBSTypeFilter, "[None]", CompareMethod.Text) = 0 Then
            numRBSTypeFilter = 0
        Else
            'find numRBSTypeFilter
            str1 = "CHARREPORTTYPE = '" & strRBSTypeFilter & "'"
            tbl = tblConfigReportType
            rows = tbl.Select(str1)
            numRBSTypeFilter = rows(0).Item("ID_TBLCONFIGREPORTTYPE")
        End If

        tbl = tblConfiguration
        strF = "CHARCONFIGGROUP = 'Global' AND CHARCONFIGCATEGORY = 'Global Settings' AND CHARCONFIGTITLE = 'Company ID'"
        Erase rows
        rows = tbl.Select(strF)
        num1 = rows(0).Item("CHARCONFIGVALUE")
        'check to see if company id has been configured
        If Len(NZ(num1, "")) = 0 Then
            If StrComp(strRBSFilter, "Company", CompareMethod.Text) = 0 Then
                MsgBox("A Company ID has not been configured in the the Administration - Global Parameters - Global Settings window. Please contact your StudyDoc Administrator.", MsgBoxStyle.Information, "Company ID has not been configured...")
                boolFormLoad = True
                boolFormLoad = True
                frmH.cbxRBSFilter.Text = "[None]"
                strRBSFilter = "[None]"
                boolFormLoad = False
                'frmh.rbBodySectionFilterNo.Checked = True
                boolFormLoad = False
            End If
        Else
            If StrComp(strRBSFilter, "Company", CompareMethod.Text) = 0 Then
                numRBSFilter = num1
            ElseIf StrComp(strRBSFilter, "[None]", CompareMethod.Text) = 0 Then
                numRBSFilter = num1
            Else 'use str
                numRBSFilter = CInt(strRBSFilter)
            End If
        End If


        'check to see if company has its own report body sections
        strF = "ID = '" & CStr(num1) & "'"
        Erase rows
        rows = tblReportCompanies.Select(strF)
        int1 = rows.Length
        boolCI = True 'company has its own report body sections
        If int1 = 0 Then
            boolCI = False
        End If
        If boolCI Then
        Else
            numRBSTypeFilter = 1
        End If

        dgv = frmH.dgvReportStatements
        dtbl = tblReportstatements

        If frmH.rbShowIncludedRBody.Checked Then
            strF = "id_tblStudies = " & id_tblStudies & " AND boolInclude = -1"
        Else
            If StrComp(strRBSFilter, "Company", CompareMethod.Text) = 0 Then
                If StrComp(strRBSTypeFilter, "[None]", CompareMethod.Text) = 0 Then
                    strF = "id_tblStudies = " & id_tblStudies & " AND (NUMCOMPANY = " & numRBSFilter & " OR NUMCOMPANY = 0)"
                Else 'a Report type has been chosen
                    strF = "id_tblStudies = " & id_tblStudies & " AND (NUMCOMPANY = " & numRBSFilter & " OR NUMCOMPANY = 0) AND (ID_TBLCONFIGREPORTTYPE = " & numRBSTypeFilter & " OR ID_TBLCONFIGREPORTTYPE = 0)"
                End If
            ElseIf StrComp(strRBSFilter, "[None]", CompareMethod.Text) = 0 Then
                If StrComp(strRBSTypeFilter, "[None]", CompareMethod.Text) = 0 Then
                    strF = "id_tblStudies = " & id_tblStudies
                Else 'a Report type has been chosen
                    strF = "id_tblStudies = " & id_tblStudies & " AND (ID_TBLCONFIGREPORTTYPE = " & numRBSTypeFilter & " OR ID_TBLCONFIGREPORTTYPE = 0)"
                End If
            Else 'a company id has been chosen
                If StrComp(strRBSTypeFilter, "[None]", CompareMethod.Text) = 0 Then
                    strF = "id_tblStudies = " & id_tblStudies & " AND (NUMCOMPANY = " & numRBSFilter & " OR NUMCOMPANY = 0)"
                Else 'a Report type has been chosen
                    strF = "id_tblStudies = " & id_tblStudies & " AND (NUMCOMPANY = " & numRBSFilter & " OR NUMCOMPANY = 0) AND (ID_TBLCONFIGREPORTTYPE = " & numRBSTypeFilter & " OR ID_TBLCONFIGREPORTTYPE = 0)"
                End If
            End If
        End If

        '''''''''''''''''''console.writeline(strF)

        Try
            Dim dv As System.Data.DataView
            dv = dgv.DataSource
            dv.RowFilter = strF
            strS = "INTORDER ASC"
            dv.Sort = strS
        Catch ex As Exception

        End Try


        Call OrderReportStatementCol() 'pesky

        boolStopRBS = False

        Cursor.Current = Cursors.WaitCursor


    End Sub

    Sub ApplyTemplate(ByVal TemplateName As String)

        'Note: Do not use GetMaxID in ApplyTemplate - too many trips to the database

        '20190218 LEE: Frontage reports crashing when 5 users attempt to load 5 new studies and apply template simultaneously.
        '"Cannot insert duplicate key in object 'dbo.TBLREPORTTABLE'. The duplicate key value is 5126
        'This looks like a maxID thing. 
        'We're going to have to use GetMaxID

        Dim numStudyD As Int32
        Dim numStudyS As Int32
        Dim tbl1 As System.Data.DataTable
        Dim tbl2 As System.Data.DataTable
        Dim rows() As DataRow
        Dim rowsD() As DataRow
        Dim rowsS() As DataRow
        Dim strF As String
        Dim strS As String
        Dim strD As String
        Dim var1, var2
        Dim col As DataColumn
        Dim str1 As String
        Dim str2 As String
        Dim int1 As Int32
        Dim intS As Int32
        Dim Count1 As Int32
        Dim Count2 As Int32
        Dim Count3 As Int32
        Dim boolGo As Int32
        Dim varRow
        Dim drowsmaxid() As DataRow
        Dim maxid As Int64
        Dim maxid1
        Dim ctP As Short
        Dim strSort As String

        Dim strFld As String
        Dim strFld1 As String
        Dim tbl As System.Data.DataTable
        Dim contr As ComboBox
        Dim drows1() As DataRow

        'Dim id_TBLREPORTS As Int64
        Dim dtbl1 As System.Data.DataTable

        'tblMaxID = tblMaxID

        numStudyD = id_tblStudies 'destination
        'find template id_tblStudies
        strF = "charTemplateName = '" & TemplateName & "'"
        tbl1 = tblTemplates
        rows = tbl1.Select(strF)
        numStudyS = rows(0).Item("id_tblStudies") 'source
        varRow = rows(0).Item("id_tblTemplates")

        If numStudyD = numStudyS Then
            str1 = "This study is the template study for '" & TemplateName & "' and cannot be applied to itself."
            MsgBox(str1, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        Cursor.Current = Cursors.WaitCursor
        Call PositionProgress()
        frmH.lblProgress.Text = "Applying Template " & TemplateName & "..."
        frmH.lblProgress.Visible = True
        frmH.lblProgress.Refresh()

        frmH.panProgress.Visible = True
        frmH.panProgress.Refresh()

        'now do Home page
        ctP = 1
        frmH.pb1.Value = ctP
        frmH.pb1.Maximum = 20
        frmH.pb1.Visible = True
        frmH.pb1.Refresh()

        'retrieve appropriate study template record
        Dim tblT As System.Data.DataTable
        Dim rowsZ() As DataRow
        Dim rec
        tblT = tblTemplateAttributes
        strF = "id_tblTab1 = 1 AND id_tblTemplates = " & varRow
        Erase rowsZ
        rowsZ = tblT.Select(strF)
        boolGo = rowsZ(0).Item("boolInclude")
        If boolGo = -1 Then
            'Word document template
            tbl1 = tblReports
            strS = "id_tblStudies = " & numStudyS
            strD = "id_tblStudies = " & numStudyD
            rowsS = tbl1.Select(strS)
            intS = rowsS.Length

            'Erase drowsmaxid
            'drowsmaxid = tblMaxID.Select("charTable='tblReports'")
            'maxid = drowsmaxid(0).Item("numMaxID")
            'maxid1 = maxid

            rowsD = tbl1.Select(strD)
            int1 = rowsD.Length

            If intS = 0 Then
            Else
                'rowsD = tbl1.Select(strD)
                'int1 = rowsD.Length

                '20190219 LEE:
                maxid = GetMaxID("tblReports", intS, True)
                maxid1 = maxid

                If int1 = 0 Then
                    Dim rowsA As DataRow = tbl1.NewRow 'tbl1 = tblReports
                    rowsA.BeginEdit()
                    maxid = maxid + 1
                    rowsA.Item("id_tblReports") = maxid
                    'record for later
                    id_tblReports = maxid
                    For Each col In tbl1.Columns

                        str1 = col.ColumnName
                        var1 = rowsS(0).Item(str1)
                        If StrComp(str1, "charReportTitle", CompareMethod.Text) = 0 Then
                        ElseIf StrComp(str1, "id_tblReports", CompareMethod.Text) = 0 Then
                        ElseIf StrComp(str1, "charReportNumber", CompareMethod.Text) = 0 Then
                        ElseIf StrComp(str1, "dtReportDraftIssueDate", CompareMethod.Text) = 0 Then
                        ElseIf StrComp(str1, "dtReportFinalIssueDate", CompareMethod.Text) = 0 Then
                        ElseIf StrComp(str1, "id_tblStudies", CompareMethod.Text) = 0 Then
                            rowsA.Item(str1) = id_tblStudies
                        Else
                            rowsA.Item(str1) = var1
                        End If

                        If StrComp(str1, "BOOLDISPLAYATTACHMENTS", CompareMethod.Text) = 0 Then
                            If NZ(var1, 0) = 0 Then
                                gboolDisplayAttachments = False
                            Else
                                gboolDisplayAttachments = True
                            End If
                        End If

                        If StrComp(str1, "BOOLREADONLYTABLES", CompareMethod.Text) = 0 Then
                            If NZ(var1, 0) = 0 Then
                                gboolReadOnlyTables = False
                            Else
                                gboolReadOnlyTables = True
                            End If
                        End If

                    Next
                    rowsA.EndEdit()
                    tbl1.Rows.Add(rowsA)
                Else
                    rowsD(0).BeginEdit()
                    For Each col In tbl1.Columns
                        If col.Ordinal = 0 Then
                        Else
                            str1 = col.ColumnName
                            var1 = rowsS(0).Item(str1)

                            If StrComp(str1, "charReportTitle", CompareMethod.Text) = 0 Then
                            ElseIf StrComp(str1, "id_tblReports", CompareMethod.Text) = 0 Then
                                id_tblReports = var1
                            ElseIf StrComp(str1, "charReportNumber", CompareMethod.Text) = 0 Then
                            ElseIf StrComp(str1, "dtReportDraftIssueDate", CompareMethod.Text) = 0 Then
                            ElseIf StrComp(str1, "dtReportFinalIssueDate", CompareMethod.Text) = 0 Then
                            ElseIf StrComp(str1, "id_tblStudies", CompareMethod.Text) = 0 Then
                                rowsD(0).Item(str1) = id_tblStudies
                            Else
                                rowsD(0).Item(str1) = var1
                            End If
                        End If

                        If StrComp(str1, "BOOLDISPLAYATTACHMENTS", CompareMethod.Text) = 0 Then
                            If NZ(var1, 0) = 0 Then
                                gboolDisplayAttachments = False
                            Else
                                gboolDisplayAttachments = True
                            End If
                        End If

                        If StrComp(str1, "BOOLREADONLYTABLES", CompareMethod.Text) = 0 Then
                            If NZ(var1, 0) = 0 Then
                                gboolReadOnlyTables = False
                            Else
                                gboolReadOnlyTables = True
                            End If
                        End If

                    Next
                    rowsD(0).EndEdit()
                End If
            End If
        End If

        ''20190219 LEE: Don't need anymore. Used GetMaxID
        'If maxid1 = maxid Then 'do no action
        'Else

        '    drowsmaxid(0).BeginEdit()
        '    drowsmaxid(0).Item("numMaxID") = maxid
        '    drowsmaxid(0).EndEdit()

        '    If boolGuWuOracle Then
        '        Try
        '            ta_tblMaxID.Update(tblMaxID)
        '        Catch ex As DBConcurrencyException
        '            'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
        '        End Try
        '    ElseIf boolGuWuAccess Then
        '        Try
        '            ta_tblMaxIDAcc.Update(tblMaxID)
        '        Catch ex As DBConcurrencyException
        '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
        '        End Try
        '    ElseIf boolGuWuSQLServer Then
        '        Try
        '            ta_tblMaxIDSQLServer.Update(tblMaxID)
        '        Catch ex As DBConcurrencyException
        '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
        '        End Try
        '    End If

        'End If



        'now do tblReportHeaders
        boolGo = 0 ' -1 'ReportHeaders disabled in 2.0.62

        If boolGo = -1 Then
            tbl1 = tblReportHeaders
            strS = "id_tblStudies = " & numStudyS
            strD = "id_tblStudies = " & numStudyD
            Erase rowsS
            Erase rowsD
            rowsS = tbl1.Select(strS)
            rowsD = tbl1.Select(strD)
            int1 = rowsD.Length
            intS = rowsS.Length

            If intS = 0 Then 'add default records 'won't happen
                'Call CheckForTblProperties(-1)
            Else

                'Erase drowsmaxid
                'drowsmaxid = tblMaxID.Select("charTable='tblReportHeaders'")
                'maxid = drowsmaxid(0).Item("numMaxID") 'tblTableProperties

                '20190219 LEE:
                maxid = GetMaxID("tblReportHeaders", intS, True)
                maxid1 = maxid

                'first delete rows
                For Count1 = int1 - 1 To 0 Step -1
                    rowsD(Count1).Delete()
                Next
                'now add rows
                For Count2 = 0 To intS - 1 'add row
                    Dim rowsA As DataRow = tbl1.NewRow
                    rowsA.BeginEdit()
                    For Count1 = 0 To tbl1.Columns.Count - 1
                        var1 = rowsS(Count2).Item(Count1)
                        rowsA.Item(Count1) = var1
                    Next
                    maxid = maxid + 1
                    rowsA.Item("ID_TBLREPORTHEADERS") = maxid
                    rowsA.Item("id_tblStudies") = numStudyD
                    rowsA.Item("ID_TBLREPORTS") = id_tblReports
                    rowsA.EndEdit()
                    tbl1.Rows.Add(rowsA)
                Next
            End If

            ' ''20190219 LEE: Don't need anymore. Used GetMaxID
            'If maxid1 = maxid Then 'do no action
            'Else

            '    drowsmaxid(0).BeginEdit() 'tblReportTable
            '    drowsmaxid(0).Item("numMaxID") = maxid
            '    drowsmaxid(0).EndEdit()

            '    If boolGuWuOracle Then
            '        Try
            '            ta_tblMaxID.Update(tblMaxID)
            '        Catch ex As DBConcurrencyException
            '            'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
            '        End Try
            '    ElseIf boolGuWuAccess Then
            '        Try
            '            ta_tblMaxIDAcc.Update(tblMaxID)
            '        Catch ex As DBConcurrencyException
            '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '        End Try
            '    ElseIf boolGuWuSQLServer Then
            '        Try
            '            ta_tblMaxIDSQLServer.Update(tblMaxID)
            '        Catch ex As DBConcurrencyException
            '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '        End Try
            '    End If

            'End If

        End If


        'now do tblCustumFieldCodes
        'strF = "id_tblTab1 = 36"
        'Erase rowsZ
        'rowsZ = tblT.Select(strF)
        'boolGo = rowsZ(0).Item("boolInclude")

        boolGo = -1

        If boolGo = -1 Then

            tbl1 = tblCustomFieldCodes
            strS = "id_tblStudies = " & numStudyS
            strD = "id_tblStudies = " & numStudyD
            Erase rowsS
            Erase rowsD
            rowsS = tbl1.Select(strS)
            rowsD = tbl1.Select(strD)
            int1 = rowsD.Length
            intS = rowsS.Length

            If intS = 0 Then 'add default records 'won't happen
                'Call CheckForTblProperties(-1)
            Else

                'Erase drowsmaxid
                'drowsmaxid = tblMaxID.Select("charTable='tblCustomFieldCodes'")
                'maxid = drowsmaxid(0).Item("numMaxID") 'tblTableProperties

                '20190219 LEE:
                maxid = GetMaxID("tblCustomFieldCodes", intS, True)
                maxid1 = maxid


                'first delete rows
                For Count1 = int1 - 1 To 0 Step -1
                    rowsD(Count1).Delete()
                Next
                'now add rows
                For Count2 = 0 To intS - 1 'add row
                    Dim rowsA As DataRow = tbl1.NewRow
                    rowsA.BeginEdit()
                    For Count1 = 0 To tbl1.Columns.Count - 1
                        var1 = rowsS(Count2).Item(Count1)
                        rowsA.Item(Count1) = var1
                    Next
                    maxid = maxid + 1
                    rowsA.Item("ID_TBLCUSTOMFIELDCODES") = maxid
                    rowsA.Item("id_tblStudies") = numStudyD
                    rowsA.EndEdit()
                    tbl1.Rows.Add(rowsA)
                Next
            End If

            'If maxid1 = maxid Then 'do no action
            'Else

            '    drowsmaxid(0).BeginEdit() 'tblCustomFieldCodes
            '    drowsmaxid(0).Item("numMaxID") = maxid
            '    drowsmaxid(0).EndEdit()

            '    If boolGuWuOracle Then
            '        Try
            '            ta_tblMaxID.Update(tblMaxID)
            '        Catch ex As DBConcurrencyException
            '            'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
            '        End Try
            '    ElseIf boolGuWuAccess Then
            '        Try
            '            ta_tblMaxIDAcc.Update(tblMaxID)
            '        Catch ex As DBConcurrencyException
            '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '        End Try
            '    ElseIf boolGuWuSQLServer Then
            '        Try
            '            ta_tblMaxIDSQLServer.Update(tblMaxID)
            '        Catch ex As DBConcurrencyException
            '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '        End Try
            '    End If

            'End If

            Call FillFCRW()

        End If

        'now do data
        ctP = ctP + 1
        frmH.pb1.Value = ctP
        frmH.pb1.Refresh()

        strF = "id_tblTab1 = 2 AND id_tblTemplates = " & varRow
        Erase rowsZ
        rowsZ = tblT.Select(strF)
        boolGo = rowsZ(0).Item("boolInclude")
        If boolGo = -1 Then
            tbl1 = tblData
            strS = "id_tblStudies = " & numStudyS
            strD = "id_tblStudies = " & numStudyD
            Erase rowsS
            Erase rowsD
            rowsS = tbl1.Select(strS)
            rowsD = tbl1.Select(strD)
            int1 = rowsD.Length
            intS = rowsS.Length
            'Erase drowsmaxid
            'drowsmaxid = tblMaxID.Select("charTable='tblData'")
            'maxid = drowsmaxid(0).Item("numMaxID")
            'maxid1 = maxid

            '20190219 LEE:
            maxid = GetMaxID("tblData", intS, True)
            maxid1 = maxid

            If int1 = 0 Then
                For Count2 = 0 To intS - 1 'add rows
                    Dim rowsA As DataRow = tbl1.NewRow
                    rowsA.BeginEdit()
                    maxid = maxid + 1
                    'DEBUG

                    For Count1 = 0 To tbl1.Columns.Count - 1
                        str1 = tbl1.Columns.Item(Count1).ColumnName

                        If StrComp(str1, "ID_TBLDATA", CompareMethod.Text) = 0 Then
                            var1 = maxid
                            rowsA.Item(Count1) = var1
                        ElseIf StrComp(str1, "ID_TBLSTUDIES", CompareMethod.Text) = 0 Then

                        ElseIf StrComp(str1, "CHARCORPORATESTUDYID", CompareMethod.Text) = 0 Then
                            var1 = "[NONE]"
                            rowsA.Item(Count1) = var1
                        ElseIf StrComp(str1, "CHARPROTOCOLNUMBER", CompareMethod.Text) = 0 Then
                            var1 = "[NONE]"
                            rowsA.Item(Count1) = var1
                        ElseIf StrComp(str1, "CHARSPONSORSTUDYNUMBER", CompareMethod.Text) = 0 Then
                            var1 = "[NONE]"
                            rowsA.Item(Count1) = var1
                        ElseIf StrComp(str1, "CHARSPONSORSTUDYTITLE", CompareMethod.Text) = 0 Then
                            var1 = "[NONE]"
                            rowsA.Item(Count1) = var1
                        Else
                            var1 = rowsS(Count2).Item(Count1)
                            rowsA.Item(Count1) = var1

                            'now check to see if there is a cbx to modify
                            Dim boolCBX As Boolean
                            Dim boolCorp As Boolean

                            boolCorp = False
                            boolCBX = False
                            strFld = "1"

                            Select Case str1
                                Case "ID_TBLASSAYTECHNIQUE"
                                    boolCBX = True
                                    boolCorp = False
                                    strFld = "3"
                                    contr = frmH.cbxAssayTechnique
                                Case "ID_TBLANTICOAGULANT"
                                    boolCBX = True
                                    boolCorp = False
                                    strFld = "1"
                                    contr = frmH.cbxAnticoagulant
                                Case "ID_SUBMITTEDBY"
                                    boolCBX = False
                                    boolCorp = True
                                    strFld = "1"
                                    contr = frmH.cbxSubmittedBy
                                Case "ID_SUBMITTEDTO"
                                    boolCBX = False
                                    boolCorp = True
                                    strFld = "1"
                                    contr = frmH.cbxSubmittedTo
                                Case "ID_INSUPPORTOF"
                                    boolCBX = False
                                    boolCorp = True
                                    strFld = "1"
                                    contr = frmH.cbxInSupportOf
                            End Select

                            If boolCBX Then
                                tbl = tblDropdownBoxContent 'for data page update code
                                var2 = "[None]"
                                If Len(var1) = 0 Then
                                    var2 = "[None]"
                                Else
                                    str2 = "id_tblDropdownBoxContent = " & var1
                                    drows1 = tbl.Select(str2)
                                    If drows1.Length = 0 Then
                                    Else
                                        var2 = drows1(0).Item("charValue")
                                    End If
                                End If
                                contr.Text = var2
                                If StrComp(strFld, "3", CompareMethod.Text) = 0 Then
                                    'do cbxAcronym also
                                    var2 = drows1(0).Item("charAcronym")
                                    frmH.cbxAssayTechniqueAcronym.Text = var2
                                End If
                            End If

                            If boolCorp Then
                                'tbl = tblCorporateAddresses
                                tbl = tblCorporateNickNames
                                strFld = "charNickname"
                                'strFld1 = "id_tblCorporateAddresses"
                                strFld1 = "id_tblCorporateNickNames"

                                var2 = "[None]"
                                If Len(var1) = 0 Or var1 = 0 Then
                                    var2 = "[None]"
                                Else
                                    str2 = strFld1 & " = '" & var1 & "'"
                                    drows1 = tbl.Select(str2, "id_tblCorporateNickNames ASC")
                                    If drows1.Length = 0 Then
                                    Else
                                        var2 = drows1(0).Item(strFld)
                                    End If
                                End If
                                contr.Text = var2

                            End If

                        End If
                        'var1 = rowsS(Count2).Item(Count1)
                        'rowsA.Item(Count1) = var1
                    Next
                    rowsA.Item("id_tblStudies") = id_tblStudies
                    rowsA.EndEdit()
                    tbl1.Rows.Add(rowsA)
                Next
            Else
                For Count2 = 0 To intS - 1 'edit rows
                    rowsD(Count2).BeginEdit()
                    For Count1 = 1 To tbl1.Columns.Count - 1
                        str1 = tbl1.Columns.Item(Count1).ColumnName
                        If StrComp(str1, "ID_TBLDATA", CompareMethod.Text) = 0 Then
                            'ElseIf StrComp(str1, "ID_TBLSTUDIES", CompareMethod.Text) = 0 Then
                            'ElseIf StrComp(str1, "CHARCORPORATESTUDYID	", CompareMethod.Text) = 0 Then
                            'ElseIf StrComp(str1, "CHARPROTOCOLNUMBER	", CompareMethod.Text) = 0 Then
                            'ElseIf StrComp(str1, "CHARSPONSORSTUDYNUMBER	", CompareMethod.Text) = 0 Then
                            'ElseIf StrComp(str1, "CHARSPONSORSTUDYTITLE	", CompareMethod.Text) = 0 Then

                        ElseIf StrComp(str1, "CHARCORPORATESTUDYID", CompareMethod.Text) = 0 Then
                        ElseIf StrComp(str1, "CHARPROTOCOLNUMBER", CompareMethod.Text) = 0 Then
                        ElseIf StrComp(str1, "CHARSPONSORSTUDYNUMBER", CompareMethod.Text) = 0 Then
                        ElseIf StrComp(str1, "CHARSPONSORSTUDYTITLE", CompareMethod.Text) = 0 Then
                        Else
                            var1 = rowsS(Count2).Item(Count1)
                            rowsD(Count2).Item(Count1) = var1

                            'now check to see if there is a cbx to modify
                            Dim boolCBX As Boolean
                            Dim boolCorp As Boolean

                            boolCorp = False
                            boolCBX = False
                            strFld = "1"

                            Select Case str1
                                Case "ID_TBLASSAYTECHNIQUE"
                                    boolCBX = True
                                    boolCorp = False
                                    strFld = "3"
                                    contr = frmH.cbxAssayTechnique
                                Case "ID_TBLANTICOAGULANT"
                                    boolCBX = True
                                    boolCorp = False
                                    strFld = "1"
                                    contr = frmH.cbxAnticoagulant
                                Case "ID_SUBMITTEDBY"
                                    boolCBX = False
                                    boolCorp = True
                                    strFld = "1"
                                    contr = frmH.cbxSubmittedBy
                                Case "ID_SUBMITTEDTO"
                                    boolCBX = False
                                    boolCorp = True
                                    strFld = "1"
                                    contr = frmH.cbxSubmittedTo
                                Case "ID_INSUPPORTOF"
                                    boolCBX = False
                                    boolCorp = True
                                    strFld = "1"
                                    contr = frmH.cbxInSupportOf
                            End Select

                            If boolCBX Then
                                tbl = tblDropdownBoxContent 'for data page update code
                                Erase drows1
                                var2 = "[None]"
                                If Len(var1) = 0 Then
                                    var2 = "[None]"
                                Else
                                    str2 = "id_tblDropdownBoxContent = " & var1
                                    drows1 = tbl.Select(str2)
                                    If drows1.Length = 0 Then
                                    Else
                                        var2 = drows1(0).Item("charValue")
                                    End If
                                End If
                                contr.Text = var2
                                If StrComp(strFld, "3", CompareMethod.Text) = 0 Then
                                    'do cbxAcronym also
                                    var2 = drows1(0).Item("charAcronym")
                                    frmH.cbxAssayTechniqueAcronym.Text = var2
                                End If
                            End If

                            If boolCorp Then
                                'tbl = tblCorporateAddresses
                                tbl = tblCorporateNickNames
                                strFld = "charNickname"
                                'strFld1 = "id_tblCorporateAddresses"
                                strFld1 = "id_tblCorporateNickNames"

                                var2 = "[None]"
                                If Len(var1) = 0 Or var1 = 0 Then
                                    var2 = "[None]"
                                Else
                                    str2 = strFld1 & " = '" & var1 & "'"
                                    drows1 = tbl.Select(str2, "id_tblCorporateNickNames ASC")
                                    If drows1.Length = 0 Then
                                    Else
                                        var2 = drows1(0).Item(strFld)
                                    End If
                                End If
                                contr.Text = var2

                            End If

                        End If
                    Next
                    rowsD(Count2).Item("id_tblStudies") = id_tblStudies
                    rowsD(Count2).EndEdit()
                Next
            End If

        End If

        ' ''20190219 LEE: Don't need anymore. Used GetMaxID
        'If maxid1 = maxid Then 'do no action
        'Else

        '    drowsmaxid(0).BeginEdit()
        '    drowsmaxid(0).Item("numMaxID") = maxid
        '    drowsmaxid(0).EndEdit()
        '    If boolGuWuOracle Then
        '        Try
        '            ta_tblMaxID.Update(tblMaxID)
        '        Catch ex As DBConcurrencyException
        '            'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
        '        End Try
        '    ElseIf boolGuWuAccess Then
        '        Try
        '            ta_tblMaxIDAcc.Update(tblMaxID)
        '        Catch ex As DBConcurrencyException
        '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
        '        End Try
        '    ElseIf boolGuWuSQLServer Then
        '        Try
        '            ta_tblMaxIDSQLServer.Update(tblMaxID)
        '        Catch ex As DBConcurrencyException
        '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
        '        End Try
        '    End If

        'End If

        'now fill data tab
        Call FillDataTabData(False)


        'now do Summary Table
        ctP = ctP + 1
        frmH.pb1.Value = ctP
        frmH.pb1.Refresh()

        strF = "id_tblTab1 = 4 AND id_tblTemplates = " & varRow
        Erase rowsZ
        rowsZ = tblT.Select(strF)
        boolGo = rowsZ(0).Item("boolInclude")
        If boolGo = -1 Then
            'do boolInclude and intOrder
            tbl1 = tblSummaryData
            strS = "id_tblStudies = " & numStudyS
            strD = "id_tblStudies = " & numStudyD
            Erase rowsS
            Erase rowsD
            rowsS = tbl1.Select(strS)
            rowsD = tbl1.Select(strD)
            int1 = rowsD.Length
            intS = rowsS.Length

            'first delete all rows
            For Count2 = 0 To int1 - 1 'edit rows
                rowsD(Count2).Delete()
            Next

            'now add rows
            For Count2 = 0 To intS - 1 'add rows
                Dim rowsA As DataRow = tbl1.NewRow
                rowsA.BeginEdit()
                For Count1 = 0 To tbl1.Columns.Count - 1
                    var1 = rowsS(Count2).Item(Count1)
                    rowsA.Item(Count1) = var1
                Next
                rowsA.Item("id_tblStudies") = id_tblStudies
                rowsA.EndEdit()
                tbl1.Rows.Add(rowsA)
            Next


            '    If int1 = 0 Then
            '        For Count2 = 0 To intS - 1 'add rows
            '            Dim rowsA As DataRow = tbl1.NewRow
            '            rowsA.BeginEdit()
            '            For Count1 = 0 To tbl1.Columns.Count - 1
            '                var1 = rowsS(Count2).Item(Count1)
            '                rowsA.Item(Count1) = var1
            '            Next
            '            rowsA.Item("id_tblStudies") = id_tblStudies
            '            rowsA.EndEdit()
            '            tbl1.Rows.Add(rowsA)
            '        Next
            '    ElseIf int1 < intS Then
            '        'delete all rowsD
            '        For Count2 = 0 To int1 - 1 'edit rows
            '            'rowsD(Count2).BeginEdit()
            '            rowsD(Count2).Delete()
            '            'rowsD(Count2).EndEdit()
            '        Next
            '        'now add rows
            '        For Count2 = 0 To intS - 1 'add rows
            '            Dim rowsA As DataRow = tbl1.NewRow
            '            rowsA.BeginEdit()
            '            For Count1 = 0 To tbl1.Columns.Count - 1
            '                var1 = rowsS(Count2).Item(Count1)
            '                rowsA.Item(Count1) = var1
            '            Next
            '            rowsA.Item("id_tblStudies") = id_tblStudies
            '            rowsA.EndEdit()
            '            tbl1.Rows.Add(rowsA)
            '        Next
            '    Else
            '        For Count2 = 0 To intS - 1 'edit rows
            '            rowsD(Count2).BeginEdit()
            '            For Count1 = 1 To tbl1.Columns.Count - 1
            '                var1 = rowsS(Count2).Item(Count1)
            '                rowsD(Count2).Item(Count1) = var1
            '            Next
            '            rowsD(Count2).Item("id_tblStudies") = id_tblStudies
            '            rowsD(Count2).EndEdit()
            '        Next
            '    End If
        End If

        'now do Appendices and Figures
        ctP = ctP + 1
        frmH.pb1.Value = ctP
        frmH.pb1.Refresh()

        strF = "id_tblTab1 = 13 AND id_tblTemplates = " & varRow
        Erase rowsZ
        rowsZ = tblT.Select(strF)
        If rowsZ.Length = 0 Then
            boolGo = -1
        Else
            boolGo = rowsZ(0).Item("boolInclude")
        End If
        If boolGo = -1 Then
            tbl1 = tblAppFigs
            strS = "id_tblStudies = " & numStudyS
            strD = "id_tblStudies = " & numStudyD
            Erase rowsS
            Erase rowsD

            var1 = tbl1.Rows.Count 'debug
            rowsS = tbl1.Select(strS)
            rowsD = tbl1.Select(strD)
            int1 = rowsD.Length
            intS = rowsS.Length

            'Erase drowsmaxid
            'drowsmaxid = tblMaxID.Select("charTable='tblAppFigs'")
            'maxid = drowsmaxid(0).Item("numMaxID")
            'maxid1 = maxid

            '20190219 LEE:
            maxid = GetMaxID("tblAppFigs", intS, True)
            maxid1 = maxid

            'first delete rows
            For Count1 = int1 - 1 To 0 Step -1
                rowsD(Count1).Delete()
            Next
            'now add rows
            For Count2 = 0 To intS - 1 'add rows
                Dim rowsA As DataRow = tbl1.NewRow
                rowsA.BeginEdit()
                For Count1 = 0 To tbl1.Columns.Count - 1
                    var1 = rowsS(Count2).Item(Count1)
                    rowsA.Item(Count1) = var1
                Next
                maxid = maxid + 1
                rowsA.Item("id_tblStudies") = id_tblStudies
                rowsA.Item("id_tblAppFigs") = maxid
                rowsA.EndEdit()
                tbl1.Rows.Add(rowsA)
            Next

        End If

        ' ''20190219 LEE: Don't need anymore. Used GetMaxID
        'If maxid1 = maxid Then 'do no action
        'Else

        '    drowsmaxid(0).BeginEdit()
        '    drowsmaxid(0).Item("numMaxID") = maxid
        '    drowsmaxid(0).EndEdit()
        '    If boolGuWuOracle Then
        '        Try
        '            ta_tblMaxID.Update(tblMaxID)
        '        Catch ex As DBConcurrencyException
        '            'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
        '        End Try
        '    ElseIf boolGuWuAccess Then
        '        Try
        '            ta_tblMaxIDAcc.Update(tblMaxID)
        '        Catch ex As DBConcurrencyException
        '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
        '        End Try
        '    ElseIf boolGuWuSQLServer Then
        '        Try
        '            ta_tblMaxIDSQLServer.Update(tblMaxID)
        '        Catch ex As DBConcurrencyException
        '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
        '        End Try
        '    End If

        'End If


        'now do tblReportTable 

        'NOTE: tblReportTable gets filled before this from old code, but is incomplete, so we will delete those rows and do it again
        ''20190218 LEE: Hmmm. Can't find anything with tblReportTable being done earlier
        '20171110 LEE:  Aack! following code was erasing tblAppFig
        'need to assign tblReportTable to tbl1
        'tbl1 = tblReportTable
        '20171206 LEE:
        'No, don't assign tbl1 early - it gets assigned right away
        'simple remove .clear and .acceptchanges commands

        'NOTE: tblReportTable is only filled for a given study
        'must go and get data from template study for source
        If boolGuWuAccess Then
            'tbl1.Clear()
            'tbl1.AcceptChanges()
            tbl1 = ta_tblReportTableAcc.GetDataBy_ID_TBLSTUDIES(numStudyS)
            int1 = tbl1.Rows.Count 'DEBUG
        ElseIf boolGuWuSQLServer Then
            'tbl1.Clear()
            'tbl1.AcceptChanges()
            tbl1 = ta_tblReportTableSQLServer.GetDataBy_ID_TBLSTUDIES(numStudyS)
            int1 = tbl1.Rows.Count 'DEBUG
        ElseIf boolGuWuOracle Then
            str1 = "Need Oracle code here:  ApplyTemplate"
            MsgBox(str1, vbInformation, "Oracle code...")
        End If

        'now set tblReportTable to tbl2 to keep code consistent
        tbl2 = tblReportTable

        Dim tblRID As New DataTable
        Dim intRID As Int64
        Dim col1 As New DataColumn
        col1.ColumnName = "IDS"
        col1.DataType = System.Type.GetType("System.Int64")
        tblRID.Columns.Add(col1)
        Dim col2 As New DataColumn
        col2.ColumnName = "IDD"
        col2.DataType = System.Type.GetType("System.Int64")
        tblRID.Columns.Add(col2)

        ctP = ctP + 1
        frmH.pb1.Value = ctP
        frmH.pb1.Refresh()

        'strF = "id_tblTab1 = 5 AND id_tblTemplates = " & varRow
        'Erase rowsZ
        'rowsZ = tblT.Select(strF)
        'boolGo = rowsZ(0).Item("boolInclude")

        Dim intD As Short
        boolGo = -1
        If boolGo = -1 Then

            'tbl1 = tblReportTable
            strS = "id_tblStudies = " & numStudyS
            strD = "id_tblStudies = " & numStudyD
            Erase rowsS
            Erase rowsD
            rowsS = tbl1.Select(strS)
            rowsD = tbl2.Select(strD)
            intD = rowsD.Length
            intS = rowsS.Length

            Dim tblRT As System.Data.DataTable
            Dim intRT As Short
            Dim rowsRT() As DataRow
            Dim rowsSS() As DataRow
            Dim boolG As Boolean
            tblRT = tblConfigReportTables
            strF = "id_tblConfigReportType < 1000"
            rowsRT = tblRT.Select(strF)
            intRT = rowsRT.Length

            boolG = False
            'Erase drowsmaxid
            'drowsmaxid = tblMaxID.Select("charTable='tblReportTable'")
            'maxid = drowsmaxid(0).Item("numMaxID")
            'maxid1 = maxid

            '20190219 LEE:
            maxid = GetMaxID("tblReportTable", intS, True)
            maxid1 = maxid

            'now delete rowsd
            For Count1 = 0 To intD - 1
                rowsD(Count1).Delete()
            Next

            int1 = tbl2.Rows.Count 'debug
            int1 = int1

            'update tblReportTable
            If boolGuWuOracle Then
                Try
                    ta_tblReportTable.Update(tbl2) 'tbl2 = tblReportTable
                Catch ex As DBConcurrencyException
                    'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblReportTableAcc.Update(tbl2)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblReportTableSQLServer.Update(tbl2)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
                End Try
            End If


            'now add rows
            '20190218 LEE: Frontage reports crashing when 5 users attempt to load 5 new studies and apply template simultaneously.
            '"Cannot insert duplicate key in object 'dbo.TBLREPORTTABLE'. The duplicate key value is 5126
            'This looks like a maxID thing

            For Count2 = 0 To intS - 1

                maxid = maxid + 1
                Dim nr As DataRow = tbl2.NewRow
                nr.BeginEdit()
                For Count1 = 0 To tbl1.Columns.Count - 1
                    var2 = tbl1.Columns(Count1).ColumnName 'debug
                    var1 = rowsS(Count2).Item(Count1)
                    nr.Item(Count1) = var1
                Next

                'These id's need to be added for later
                intRID = rowsS(Count2).Item("ID_TBLREPORTTABLE")
                Dim nrID As DataRow = tblRID.NewRow
                nrID.BeginEdit()
                nrID("IDS") = intRID
                nrID("IDD") = maxid
                nrID.EndEdit()
                tblRID.Rows.Add(nrID)

                nr.Item("id_tblStudies") = numStudyD
                nr.Item("id_tblReportTable") = maxid
                nr.EndEdit()
                tbl2.Rows.Add(nr)

            Next

            int1 = tbl2.Rows.Count 'debug

            'If maxid1 = maxid Then 'do no action
            'Else

            '    drowsmaxid(0).BeginEdit()
            '    drowsmaxid(0).Item("numMaxID") = maxid
            '    drowsmaxid(0).EndEdit()
            '    If boolGuWuOracle Then
            '        Try
            '            ta_tblMaxID.Update(tblMaxID)
            '        Catch ex As DBConcurrencyException
            '            'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
            '        End Try
            '    ElseIf boolGuWuAccess Then
            '        Try
            '            ta_tblMaxIDAcc.Update(tblMaxID)
            '        Catch ex As DBConcurrencyException
            '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '        End Try
            '    ElseIf boolGuWuSQLServer Then
            '        Try
            '            ta_tblMaxIDSQLServer.Update(tblMaxID)
            '        Catch ex As DBConcurrencyException
            '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '        End Try
            '    End If

            '    'save tblReportTable
            '    If boolGuWuOracle Then
            '        Try
            '            ta_tblReportTable.Update(tbl2)
            '        Catch ex As DBConcurrencyException
            '            'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
            '        End Try
            '    ElseIf boolGuWuAccess Then
            '        Try
            '            ta_tblReportTableAcc.Update(tbl2)
            '        Catch ex As DBConcurrencyException
            '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '        End Try
            '    ElseIf boolGuWuSQLServer Then
            '        Try
            '            ta_tblReportTableSQLServer.Update(tbl2)
            '        Catch ex As DBConcurrencyException
            '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '        End Try
            '    End If

            '    'now fill tblReportTable
            '    If boolGuWuAccess Then
            '        tblReportTable.Clear()
            '        tblReportTable.AcceptChanges()
            '        tblReportTable = ta_tblReportTableAcc.GetDataBy_ID_TBLSTUDIES(numStudyD)
            '        int1 = tblReportTable.Rows.Count 'DEBUG
            '    ElseIf boolGuWuSQLServer Then
            '        tblReportTable.Clear()
            '        tblReportTable.AcceptChanges()
            '        tblReportTable = ta_tblReportTableSQLServer.GetDataBy_ID_TBLSTUDIES(numStudyD)
            '        int1 = tblReportTable.Rows.Count 'DEBUG
            '    ElseIf boolGuWuOracle Then
            '        str1 = "Need Oracle code here:  ApplyTemplate"
            '        MsgBox(str1, vbInformation, "Oracle code...")
            '    End If

            '    int1 = int1 'debug

            'End If

            'save tblReportTable
            If boolGuWuOracle Then
                Try
                    ta_tblReportTable.Update(tbl2)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblReportTableAcc.Update(tbl2)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblReportTableSQLServer.Update(tbl2)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
                End Try
            End If

            'now fill tblReportTable
            If boolGuWuAccess Then
                tblReportTable.Clear()
                tblReportTable.AcceptChanges()
                tblReportTable = ta_tblReportTableAcc.GetDataBy_ID_TBLSTUDIES(numStudyD)
                int1 = tblReportTable.Rows.Count 'DEBUG
            ElseIf boolGuWuSQLServer Then
                tblReportTable.Clear()
                tblReportTable.AcceptChanges()
                tblReportTable = ta_tblReportTableSQLServer.GetDataBy_ID_TBLSTUDIES(numStudyD)
                int1 = tblReportTable.Rows.Count 'DEBUG
            ElseIf boolGuWuOracle Then
                str1 = "Need Oracle code here:  ApplyTemplate"
                MsgBox(str1, vbInformation, "Oracle code...")
            End If

        End If



        'TBLTABLEPROPERTIES

        'NOTE: TBLTABLEPROPERTIES gets filled before this from old code, but is incomplete, so we will delete those rows and do it again

        ctP = ctP + 1
        frmH.pb1.Value = ctP
        frmH.pb1.Refresh()

        'strF = "id_tblTab1 = 5 AND id_tblTemplates = " & varRow
        'Erase rowsZ
        'rowsZ = tblT.Select(strF)
        'boolGo = rowsZ(0).Item("boolInclude")

        If boolGuWuAccess Then
            'tbl1.Clear()
            'tbl1.AcceptChanges()
            tbl1 = ta_tblTablePropertiesAcc.GetDataBy_ID_TBLSTUDIES(numStudyS)
            int1 = tbl1.Rows.Count 'DEBUG
        ElseIf boolGuWuSQLServer Then
            'tbl1.Clear()
            'tbl1.AcceptChanges()
            tbl1 = ta_tblTablePropertiesSQLServer.GetDataBy_ID_TBLSTUDIES(numStudyS)
            int1 = tbl1.Rows.Count 'DEBUG
        ElseIf boolGuWuOracle Then
            str1 = "Need Oracle code here:  ApplyTemplate"
            MsgBox(str1, vbInformation, "Oracle code...")

        End If

        tbl2 = tblTableProperties

        boolGo = -1
        If boolGo = -1 Then

            'tbl1 = tblTableProperties
            'tbl2 = tblReportTable
            strS = "id_tblStudies = " & numStudyS
            strD = "id_tblStudies = " & numStudyD
            Erase rowsS
            Erase rowsD
            rowsS = tbl1.Select(strS)
            rowsD = tbl2.Select(strD)
            intD = rowsD.Length
            intS = rowsS.Length

            Dim rowsSS() As DataRow
            Dim rowsT() As DataRow

            'Erase drowsmaxid
            'drowsmaxid = tblMaxID.Select("charTable='tblTableProperties'")
            'maxid = drowsmaxid(0).Item("numMaxID")
            'maxid1 = maxid

            '20190219 LEE:
            maxid = GetMaxID("tblTableProperties", intS, True)
            maxid1 = maxid

            'delete rowsd, if they exist
            For Count1 = 0 To intD - 1
                rowsD(Count1).Delete()
            Next

            'update tblReportTable
            If boolGuWuOracle Then
                Try
                    ta_tblTableProperties.Update(tbl2)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblTablePropertiesAcc.Update(tbl2)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblTablePropertiesSQLServer.Update(tbl2)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
                End Try
            End If

            'now add rows
            For Count2 = 0 To intS - 1 'edit rows
                maxid = maxid + 1
                Dim nr As DataRow = tbl2.NewRow
                nr.BeginEdit()
                For Count1 = 0 To tbl1.Columns.Count - 1
                    var1 = rowsS(Count2).Item(Count1)
                    nr.Item(Count1) = var1
                Next
                nr.Item("id_tblStudies") = numStudyD

                'must find  ID_TBLREPORTTABLE
                var1 = rowsS(Count2).Item("ID_TBLREPORTTABLE")
                strF = "IDS = " & var1
                Erase rowsSS
                rowsSS = tblRID.Select(strF)
                If rowsSS.Length = 0 Then
                    var2 = var1
                Else
                    var2 = rowsSS(0).Item("IDD")
                End If

                nr.Item("ID_TBLREPORTTABLE") = var2
                nr.Item("ID_TBLTABLEPROPERTIES") = maxid
                nr.EndEdit()
                tbl2.Rows.Add(nr)

            Next

            ' ''20190219 LEE: Don't need anymore. Used GetMaxID
            'If maxid1 = maxid Then 'do no action
            'Else

            '    drowsmaxid(0).BeginEdit()
            '    drowsmaxid(0).Item("numMaxID") = maxid
            '    drowsmaxid(0).EndEdit()
            '    If boolGuWuOracle Then
            '        Try
            '            ta_tblMaxID.Update(tblMaxID)
            '        Catch ex As DBConcurrencyException
            '            'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
            '        End Try
            '    ElseIf boolGuWuAccess Then
            '        Try
            '            ta_tblMaxIDAcc.Update(tblMaxID)
            '        Catch ex As DBConcurrencyException
            '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '        End Try
            '    ElseIf boolGuWuSQLServer Then
            '        Try
            '            ta_tblMaxIDSQLServer.Update(tblMaxID)
            '        Catch ex As DBConcurrencyException
            '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '        End Try
            '    End If

            'End If

            'save tblTableProperties
            If boolGuWuOracle Then
                Try
                    ta_tblTableProperties.Update(tbl2)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblTablePropertiesAcc.Update(tbl2)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblTablePropertiesSQLServer.Update(tbl2)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
                End Try
            End If

            'now fill tblAutoAssignSamples
            If boolGuWuAccess Then
                tblTableProperties.Clear()
                tblTableProperties.AcceptChanges()
                tblTableProperties = ta_tblTablePropertiesAcc.GetDataBy_ID_TBLSTUDIES(numStudyD)
                int1 = tblReportTable.Rows.Count 'DEBUG
            ElseIf boolGuWuSQLServer Then
                tblTableProperties.Clear()
                tblTableProperties.AcceptChanges()
                tblTableProperties = ta_tblTablePropertiesSQLServer.GetDataBy_ID_TBLSTUDIES(numStudyD)
                int1 = tblTableProperties.Rows.Count 'DEBUG
            ElseIf boolGuWuOracle Then
                str1 = "Need Oracle code here:  ApplyTemplate"
                MsgBox(str1, vbInformation, "Oracle code...")

            End If

        End If


        '20181112 LEE:
        'Upon new study, tblMethodValidationData isn't updating
        'Call RealMethValExecute(True)
        '20190205 LEE:
        'Hmmm. RealMethValExecute should only get called if study is Sample Analysis
        'So may throw an error. Embed in Try-Catch
        'Alturas COR-2017-2 Throws an error here if comes from ApplyTemplate
        'Also, should do boolformload here for this sub - need to ignore any false negatives to boolHit in RealMethValExecute
        Dim boolFL As Boolean = boolFormLoad
        boolFormLoad = True
        Try
            Call RealMethValExecute(False)
            If boolGuWuOracle Then
                Try
                    ta_tblMethodValidationData.Update(tblMethodValidationData)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblMethodValidationDataAcc.Update(tblMethodValidationData)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblMethodValidationDataSQLServer.Update(tblMethodValidationData)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
                End Try
            End If
        Catch ex As Exception
            var1 = var1
        End Try
        boolFormLoad = boolFL

   

        '******

        'TBLAUTOASSIGNSAMPLES
        ctP = ctP + 1
        frmH.pb1.Value = ctP
        frmH.pb1.Refresh()

        '20171206 LEE:
        tbl1 = tblAutoAssignSamples

        'strF = "id_tblTab1 = 5 AND id_tblTemplates = " & varRow
        'Erase rowsZ
        'rowsZ = tblT.Select(strF)
        'boolGo = rowsZ(0).Item("boolInclude")

        boolGo = -1
        If boolGo = -1 Then

            'NOTE: TBLAUTOASSIGNESAMPLES is only filled for a given study
            'must go and get data from template study for source
            If boolGuWuAccess Then
                'tbl1.Clear()
                'tbl1.AcceptChanges()
                tbl1 = ta_TBLAUTOASSIGNSAMPLESAcc.GetDataBy_ID_TBLSTUDIES(numStudyS)
                int1 = tbl1.Rows.Count 'DEBUG
            ElseIf boolGuWuSQLServer Then
                'tbl1.Clear()
                'tbl1.AcceptChanges()
                tbl1 = ta_TBLAUTOASSIGNSAMPLESSQLServer.GetDataBy_ID_TBLSTUDIES(numStudyS)
                int1 = tbl1.Rows.Count 'DEBUG
            ElseIf boolGuWuOracle Then
                str1 = "Need Oracle code here:  ApplyTemplate"
                MsgBox(str1, vbInformation, "Oracle code...")

            End If

            tbl2 = tblAutoAssignSamples

            'tbl1 = tblAutoAssignSamples
            strS = "id_tblStudies = " & numStudyS
            strD = "id_tblStudies = " & numStudyD
            Erase rowsS
            Erase rowsD
            rowsS = tbl1.Select(strS)
            intS = rowsS.Length
            rowsD = tbl2.Select(strD)
            intD = rowsD.Length


            Dim rowsSS() As DataRow
            Dim rowsT() As DataRow

            'Erase drowsmaxid
            'drowsmaxid = tblMaxID.Select("charTable='TBLAUTOASSIGNSAMPLES'")
            'maxid = drowsmaxid(0).Item("numMaxID")
            'maxid1 = maxid

            '20190219 LEE:
            maxid = GetMaxID("TBLAUTOASSIGNSAMPLES", intS, True)
            maxid1 = maxid

            'delete rowsd if they exist
            For Count1 = 0 To intD - 1
                rowsD(Count1).Delete()
            Next

            'update tblReportTable
            If boolGuWuOracle Then
                Try
                    'ta_TBLAUTOASSIGNSAMPLES.Update(tbl2)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_TBLAUTOASSIGNSAMPLESAcc.Update(tbl2)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_TBLAUTOASSIGNSAMPLESSQLServer.Update(tbl2)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
                End Try
            End If

            'now add rows
            For Count2 = 0 To intS - 1

                maxid = maxid + 1
                Dim nr As DataRow = tbl2.NewRow
                nr.BeginEdit()

                Try
                    For Count1 = 0 To tbl1.Columns.Count - 1
                        var1 = rowsS(Count2).Item(Count1)
                        nr.Item(Count1) = var1
                    Next

                Catch ex As Exception
                    var1 = ex.Message
                End Try

                'must findID_TBLREPORTTABLE
                var1 = rowsS(Count2).Item("ID_TBLREPORTTABLE")
                strF = "IDS = " & var1
                Erase rowsSS
                rowsSS = tblRID.Select(strF)
                If rowsSS.Length = 0 Then
                    var2 = var1
                Else
                    var2 = rowsSS(0).Item("IDD")
                End If

                nr.Item("ID_TBLREPORTTABLE") = var2
                nr.Item("id_tblStudies") = numStudyD
                nr.Item("ID_TBLAUTOASSIGNSAMPLES") = maxid
                nr.EndEdit()
                tbl2.Rows.Add(nr)

            Next

            ''20190219 LEE: Don't need anymore. Used GetMaxID
            'If maxid1 = maxid Then 'do no action
            'Else

            '    drowsmaxid(0).BeginEdit()
            '    drowsmaxid(0).Item("numMaxID") = maxid
            '    drowsmaxid(0).EndEdit()
            '    If boolGuWuOracle Then
            '        Try
            '            ta_tblMaxID.Update(tblMaxID)
            '        Catch ex As DBConcurrencyException
            '            'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
            '        End Try
            '    ElseIf boolGuWuAccess Then
            '        Try
            '            ta_tblMaxIDAcc.Update(tblMaxID)
            '        Catch ex As DBConcurrencyException
            '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '        End Try
            '    ElseIf boolGuWuSQLServer Then
            '        Try
            '            ta_tblMaxIDSQLServer.Update(tblMaxID)
            '        Catch ex As DBConcurrencyException
            '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '        End Try
            '    End If

            'End If

            'save tblAutoAssignSamples
            If boolGuWuOracle Then
                'Try
                '    ta_tblAutoAssignSamples.Update(tbl2)
                'Catch ex As DBConcurrencyException
                '    'ds2005.TBLAUTOASSIGNSAMPLES.Merge('ds2005.TBLMAXID, True)
                'End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_TBLAUTOASSIGNSAMPLESAcc.Update(tbl2)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLAUTOASSIGNSAMPLES.Merge('ds2005Acc.TBLAUTOASSIGNSAMPLES, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_TBLAUTOASSIGNSAMPLESSQLServer.Update(tbl2)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLAUTOASSIGNSAMPLES.Merge('ds2005Acc.TBLAUTOASSIGNSAMPLES, True)
                End Try
            End If


            'now fill tblAutoAssignSamples
            If boolGuWuAccess Then
                tblAutoAssignSamples.Clear()
                tblAutoAssignSamples.AcceptChanges()
                tblAutoAssignSamples = ta_TBLAUTOASSIGNSAMPLESAcc.GetDataBy_ID_TBLSTUDIES(numStudyD)
                int1 = tblAutoAssignSamples.Rows.Count 'DEBUG
            ElseIf boolGuWuSQLServer Then
                tblAutoAssignSamples.Clear()
                tblAutoAssignSamples.AcceptChanges()
                tblAutoAssignSamples = ta_TBLAUTOASSIGNSAMPLESSQLServer.GetDataBy_ID_TBLSTUDIES(numStudyD)
                int1 = tbl2.Rows.Count 'DEBUG
            ElseIf boolGuWuOracle Then
                str1 = "Need Oracle code here:  ApplyTemplate"
                MsgBox(str1, vbInformation, "Oracle code...")

            End If

        End If

        '******

        ctP = ctP + 1
        frmH.pb1.Value = ctP
        frmH.pb1.Refresh()

        'now must toggle a configuration routine
        Call DoRTConfigCancel() 'this will repopulate dgv


        'now do Report Table Header Configuration
        ctP = ctP + 1
        frmH.pb1.Value = ctP
        frmH.pb1.Refresh()

        strF = "id_tblTab1 = 6 AND id_tblTemplates = " & varRow
        Erase rowsZ
        rowsZ = tblT.Select(strF)
        boolGo = rowsZ(0).Item("boolInclude")
        boolGo = -1
        If boolGo = -1 Then


            '****

            'tblReportTableHeaderConfig is a GetDataBy query
            'must set it to normal, then set back
            'ID_TBLCONFIGHEADERLOOKUP	ID_TBLCONFIGREPORTTABLES	CHARCOLUMNLABEL	INTORDER	CHARWATSONTABLE	CHARWATSONFIELD	BOOLDEFAULT	UPSIZE_TS

            tblReportTableHeaderConfig.Clear()
            tblReportTableHeaderConfig.AcceptChanges()
            tblReportTableHeaderConfig.BeginLoadData()
            If boolGuWuOracle Then
                ta_tblReportTableHeaderConfig.ClearBeforeFill = True
                ta_tblReportTableHeaderConfig.Fill(tblReportTableHeaderConfig)
            ElseIf boolGuWuAccess Then
                ta_tblReportTableHeaderConfigAcc.ClearBeforeFill = True
                ta_tblReportTableHeaderConfigAcc.Fill(tblReportTableHeaderConfig)
            ElseIf boolGuWuSQLServer Then
                ta_tblReportTableHeaderConfigSQLServer.ClearBeforeFill = True
                ta_tblReportTableHeaderConfigSQLServer.Fill(tblReportTableHeaderConfig)
            End If
            tblReportTableHeaderConfig.EndLoadData()
            int1 = tblReportTableHeaderConfig.Rows.Count 'DEBUG


            '****

            tbl1 = tblReportTableHeaderConfig

            strS = "id_tblStudies = " & numStudyS
            strD = "id_tblStudies = " & numStudyD
            strSort = "id_tblReportTableHeaderConfig ASC"

            Erase rowsS
            Erase rowsD
            rowsS = tbl1.Select(strS, strSort)
            rowsD = tbl1.Select(strD, strSort)
            int1 = rowsD.Length
            intS = rowsS.Length

            'Erase drowsmaxid
            'drowsmaxid = tblMaxID.Select("charTable='tblReportTableHeaderConfig'")
            'maxid = drowsmaxid(0).Item("numMaxID")
            'maxid1 = maxid

            '20190219 LEE:
            maxid = GetMaxID("tblReportTableHeaderConfig", intS, True)
            maxid1 = maxid

            If int1 = 0 Then
                For Count2 = 0 To intS - 1 'add rows
                    Dim rowsA As DataRow = tbl1.NewRow
                    rowsA.BeginEdit()
                    maxid = maxid + 1
                    rowsA.Item("id_tblReportTableHeaderConfig") = maxid
                    For Count1 = 0 To tbl1.Columns.Count - 1
                        If StrComp(tbl1.Columns.Item(Count1).ColumnName, "id_tblReportTableHeaderConfig", CompareMethod.Text) = 0 Then
                        Else
                            var1 = rowsS(Count2).Item(Count1)
                            rowsA.Item(Count1) = var1
                        End If
                    Next
                    rowsA.Item("id_tblStudies") = id_tblStudies
                    rowsA.EndEdit()
                    tbl1.Rows.Add(rowsA)
                Next
            Else
                For Count2 = 0 To intS - 1 'edit rows
                    If Count2 > int1 - 1 Then
                        maxid = maxid + 1
                        Dim rowsA As DataRow = tbl1.NewRow
                        rowsA.Item("id_tblReportTableHeaderConfig") = maxid
                        For Count1 = 0 To tbl1.Columns.Count - 1
                            If StrComp(tbl1.Columns.Item(Count1).ColumnName, "id_tblReportTableHeaderConfig", CompareMethod.Text) = 0 Then
                            Else
                                var1 = rowsS(Count2).Item(Count1)
                                rowsA.Item(Count1) = var1
                            End If
                        Next
                        rowsA.Item("id_tblStudies") = id_tblStudies
                        rowsA.EndEdit()
                        tbl1.Rows.Add(rowsA)

                    Else
                        rowsD(Count2).BeginEdit()
                        For Count1 = 0 To tbl1.Columns.Count - 1
                            str1 = tbl1.Columns.Item(Count1).ColumnName

                            Select Case UCase(str1)
                                Case UCase("id_tblReportTableHeaderConfig")
                                Case UCase("ID_TBLSTUDIES")
                                Case Else
                                    var1 = rowsS(Count2).Item(Count1)
                                    rowsD(Count2).Item(Count1) = var1
                            End Select

                            'If StrComp(tbl1.Columns.Item(Count1).ColumnName, "id_tblReportTableHeaderConfig", CompareMethod.Text) = 0 Then
                            'Else
                            '    var1 = rowsS(Count2).Item(Count1)
                            '    rowsD(Count2).Item(Count1) = var1
                            '    '''''''''''''''''''''''''console.writeline(str1 & ": " & var1)
                            'End If
                        Next
                        'rowsD(Count2).Item("id_tblStudies") = id_tblStudies

                        rowsD(Count2).EndEdit()
                    End If

                Next
            End If
        End If

        ''20190219 LEE: Don't need anymore. Used GetMaxID
        'If maxid1 = maxid Then 'do no action
        'Else

        '    drowsmaxid(0).BeginEdit()
        '    drowsmaxid(0).Item("numMaxID") = maxid
        '    drowsmaxid(0).EndEdit()
        '    If boolGuWuOracle Then
        '        Try
        '            ta_tblMaxID.Update(tblMaxID)
        '        Catch ex As DBConcurrencyException
        '            'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
        '        End Try
        '    ElseIf boolGuWuAccess Then
        '        Try
        '            ta_tblMaxIDAcc.Update(tblMaxID)
        '        Catch ex As DBConcurrencyException
        '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
        '        End Try
        '    ElseIf boolGuWuSQLServer Then
        '        Try
        '            ta_tblMaxIDSQLServer.Update(tblMaxID)
        '        Catch ex As DBConcurrencyException
        '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
        '        End Try
        '    End If
        'End If

        '*****

        'update database
        If boolGuWuOracle Then
            Try
                ta_tblReportTableHeaderConfig.Update(tblReportTableHeaderConfig)
            Catch ex As DBConcurrencyException
                ''msgbox("aaReport Table Header Config: " & ex.Message)
                'ds2005.TBLREPORTTABLEHEADERCONFIG.Merge('ds2005.TBLREPORTTABLEHEADERCONFIG, True)
            End Try

        ElseIf boolGuWuAccess Then
            Try
                ta_tblReportTableHeaderConfigAcc.Update(tblReportTableHeaderConfig)
            Catch ex As DBConcurrencyException
                ''msgbox("aaReport Table Header Config: " & ex.Message)
                'ds2005Acc.TBLREPORTTABLEHEADERCONFIG.Merge('ds2005Acc.TBLREPORTTABLEHEADERCONFIG, True)
            End Try

        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblReportTableHeaderConfigSQLServer.Update(tblReportTableHeaderConfig)
            Catch ex As DBConcurrencyException
                ''msgbox("aaReport Table Header Config: " & ex.Message)
                'ds2005Acc.TBLREPORTTABLEHEADERCONFIG.Merge('ds2005Acc.TBLREPORTTABLEHEADERCONFIG, True)
            End Try
        End If


        'put back
        tblReportTableHeaderConfig.Clear()
        tblReportTableHeaderConfig.AcceptChanges()
        If boolGuWuOracle Then
            'tblReportTableHeaderConfig = ta_tblReportTableHeaderConfig.GetDataBy_ID_TBLSTUDIES(id_tblStudies)
        ElseIf boolGuWuAccess Then
            tblReportTableHeaderConfig = ta_tblReportTableHeaderConfigAcc.GetDataBy_ID_TBLSTUDIES(id_tblStudies)
        ElseIf boolGuWuSQLServer Then
            tblReportTableHeaderConfig = ta_tblReportTableHeaderConfigSQLServer.GetDataBy_ID_TBLSTUDIES(id_tblStudies)
        End If
        int1 = tblReportTableHeaderConfig.Rows.Count


        '*****

        ctP = ctP + 1
        frmH.pb1.Value = ctP
        frmH.pb1.Refresh()

        'call to configure Report Table Header Configuration
        Call ReportTableHeaderPopulateData()


        'now do Report Body Sections
        ctP = ctP + 1
        frmH.pb1.Value = ctP
        frmH.pb1.Refresh()

        Dim strCol As String

        strF = "id_tblTab1 = 9 And id_tblTemplates = " & varRow
        Erase rowsZ
        rowsZ = tblT.Select(strF)
        boolGo = rowsZ(0).Item("boolInclude")
        '20190219 LEE: tblReportStatements has multiple primary keys: id_tblStudies, id_tblConfigReportType, ID_TBLCONFIGBODYSECTIONS
        'no maxid
        If boolGo = -1 Then
            tbl1 = tblReportStatements
            strS = "id_tblStudies = " & numStudyS
            strD = "id_tblStudies = " & numStudyD
            Erase rowsS
            Erase rowsD
            rowsS = tbl1.Select(strS)
            rowsD = tbl1.Select(strD)
            int1 = rowsD.Length
            intS = rowsS.Length
            If int1 = 0 Then
                For Count2 = 0 To intS - 1 'add rows
                    Dim rowsA As DataRow = tbl1.NewRow
                    rowsA.BeginEdit()
                    For Count1 = 0 To tbl1.Columns.Count - 1
                        var1 = rowsS(Count2).Item(Count1)
                        strCol = tbl1.Columns.Item(Count1).ColumnName 'for debugging
                        rowsA.Item(Count1) = var1
                    Next
                    rowsA.Item("id_tblStudies") = id_tblStudies
                    rowsA.EndEdit()
                    tbl1.Rows.Add(rowsA)
                Next
            Else
                'delete all rows in rowsd
                For Count1 = int1 - 1 To 0 Step -1
                    rowsD(Count1).Delete()
                Next
                'tbl1.AcceptChanges()
                rowsD = tbl1.Select(strD)
                int1 = rowsD.Length

                'add rows
                For Count2 = 0 To intS - 1 'add rows
                    Dim rowsA As DataRow = tbl1.NewRow
                    rowsA.BeginEdit()
                    For Count1 = 0 To tbl1.Columns.Count - 1
                        var1 = rowsS(Count2).Item(Count1)
                        strCol = tbl1.Columns.Item(Count1).ColumnName 'for debugging
                        rowsA.Item(Count1) = var1
                    Next
                    rowsA.Item("id_tblStudies") = id_tblStudies
                    rowsA.EndEdit()
                    tbl1.Rows.Add(rowsA)
                Next
            End If

        End If

        'update database
        If boolGuWuOracle Then
            Try
                ta_tblReportStatements.Update(tblReportStatements)
            Catch ex As DBConcurrencyException
                ''msgbox("aaReport Table Header Config: " & ex.Message)
                'ds2005.TBLREPORTTABLEHEADERCONFIG.Merge('ds2005.TBLREPORTTABLEHEADERCONFIG, True)
            End Try

        ElseIf boolGuWuAccess Then
            Try
                ta_tblReportStatementsAcc.Update(tblReportStatements)
            Catch ex As DBConcurrencyException
                ''msgbox("aaReport Table Header Config: " & ex.Message)
                'ds2005Acc.TBLREPORTTABLEHEADERCONFIG.Merge('ds2005Acc.TBLREPORTTABLEHEADERCONFIG, True)
            End Try

        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblReportStatementsSQLServer.Update(tblReportStatements)
            Catch ex As DBConcurrencyException
                ''msgbox("aaReport Table Header Config: " & ex.Message)
                'ds2005Acc.TBLREPORTTABLEHEADERCONFIG.Merge('ds2005Acc.TBLREPORTTABLEHEADERCONFIG, True)
            End Try
        End If


        'now do Contributing Personnel
        ctP = ctP + 1
        frmH.pb1.Value = ctP
        frmH.pb1.Refresh()

        strF = "id_tblTab1 = 8 AND id_tblTemplates = " & varRow
        Erase rowsZ
        rowsZ = tblT.Select(strF)
        boolGo = rowsZ(0).Item("boolInclude")
        If boolGo = -1 Then
            tbl1 = tblContributingPersonnel
            strS = "id_tblStudies = " & numStudyS
            strD = "id_tblStudies = " & numStudyD
            Erase rowsS
            Erase rowsD
            rowsS = tbl1.Select(strS)
            rowsD = tbl1.Select(strD)
            int1 = rowsD.Length
            intS = rowsS.Length

            'Erase drowsmaxid
            'drowsmaxid = tblMaxID.Select("charTable='tblContributingPersonnel'")
            'maxid = drowsmaxid(0).Item("numMaxID")
            'maxid1 = maxid

            '20190219 LEE:
            maxid = GetMaxID("tblContributingPersonnel", intS, True)
            maxid1 = maxid

            'first delete rows
            For Count1 = int1 - 1 To 0 Step -1
                rowsD(Count1).Delete()
            Next
            'now add rows
            For Count2 = 0 To intS - 1 'add rows
                Dim rowsA As DataRow = tbl1.NewRow
                rowsA.BeginEdit()
                For Count1 = 0 To tbl1.Columns.Count - 1
                    var1 = rowsS(Count2).Item(Count1)
                    rowsA.Item(Count1) = var1
                Next
                maxid = maxid + 1
                rowsA.Item("id_tblContributingPersonnel") = maxid
                rowsA.Item("id_tblStudies") = id_tblStudies
                rowsA.EndEdit()
                tbl1.Rows.Add(rowsA)
            Next

        End If

        ''20190219 LEE: Don't need anymore. Used GetMaxID
        'If maxid1 = maxid Then 'do no action
        'Else

        '    drowsmaxid(0).BeginEdit()
        '    drowsmaxid(0).Item("numMaxID") = maxid
        '    drowsmaxid(0).EndEdit()
        '    If boolGuWuOracle Then
        '        Try
        '            ta_tblMaxID.Update(tblMaxID)
        '        Catch ex As DBConcurrencyException
        '            'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
        '        End Try
        '    ElseIf boolGuWuAccess Then
        '        Try
        '            ta_tblMaxIDAcc.Update(tblMaxID)
        '        Catch ex As DBConcurrencyException
        '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
        '        End Try
        '    ElseIf boolGuWuSQLServer Then
        '        Try
        '            ta_tblMaxIDSQLServer.Update(tblMaxID)
        '        Catch ex As DBConcurrencyException
        '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
        '        End Try
        '    End If
        'End If

        '20190219 LEE:
        If boolGuWuOracle Then
            'Try
            '    ta_tblContributingPersonnel.Update(tblContributingPersonnel)
            'Catch ex As DBConcurrencyException
            '    'ds2005.TBLAUTOASSIGNSAMPLES.Merge('ds2005.TBLMAXID, True)
            'End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblContributingPersonnelAcc.Update(tblContributingPersonnel)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLAUTOASSIGNSAMPLES.Merge('ds2005Acc.TBLAUTOASSIGNSAMPLES, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblContributingPersonnelSQLServer.Update(tblContributingPersonnel)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLAUTOASSIGNSAMPLES.Merge('ds2005Acc.TBLAUTOASSIGNSAMPLES, True)
            End Try
        End If

        ctP = ctP + 1
        frmH.pb1.Value = ctP
        frmH.pb1.Refresh()

        boolRSCFill = True
        Call ReportStatementsFillCharSection()
        boolRSCFill = False


        ctP = ctP + 1
        frmH.pb1.Value = ctP
        frmH.pb1.Refresh()

        'now do tblTableLegends
        boolGo = -1

        If boolGo = -1 Then
            tbl1 = tblTableLegends
            strS = "id_tblStudies = " & numStudyS
            strD = "id_tblStudies = " & numStudyD
            Erase rowsS
            Erase rowsD
            rowsS = tbl1.Select(strS)
            rowsD = tbl1.Select(strD)
            int1 = rowsD.Length
            intS = rowsS.Length

            If intS = 0 Then 'add default records 'won't happen
                'Call CheckForTblProperties(-1)
            Else

                'Erase drowsmaxid
                'drowsmaxid = tblMaxID.Select("charTable='tblTableLegends'")
                'maxid = drowsmaxid(0).Item("numMaxID") 'tblTableProperties
                'maxid1 = maxid

                '20190219 LEE:
                maxid = GetMaxID("tblTableLegends", intS, True)
                maxid1 = maxid

                'first delete rows
                For Count1 = int1 - 1 To 0 Step -1
                    rowsD(Count1).Delete()
                Next
                'now add rows
                For Count2 = 0 To intS - 1 'add row
                    Dim rowsA As DataRow = tbl1.NewRow
                    rowsA.BeginEdit()
                    For Count1 = 0 To tbl1.Columns.Count - 1
                        var1 = rowsS(Count2).Item(Count1)
                        rowsA.Item(Count1) = var1
                    Next
                    maxid = maxid + 1
                    rowsA.Item("ID_tblTableLegends") = maxid
                    rowsA.Item("id_tblStudies") = numStudyD
                    rowsA.EndEdit()
                    tbl1.Rows.Add(rowsA)
                Next
            End If

            ''20190219 LEE: Don't need anymore. Used GetMaxID
            'If maxid1 = maxid Then 'do no action
            'Else
            '    drowsmaxid(0).BeginEdit() 'tblTableLegends
            '    drowsmaxid(0).Item("numMaxID") = maxid
            '    drowsmaxid(0).EndEdit()

            '    If boolGuWuOracle Then
            '        Try
            '            ta_tblMaxID.Update(tblMaxID)
            '        Catch ex As DBConcurrencyException
            '            'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
            '        End Try
            '    ElseIf boolGuWuAccess Then
            '        Try
            '            ta_tblMaxIDAcc.Update(tblMaxID)
            '        Catch ex As DBConcurrencyException
            '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '        End Try
            '    ElseIf boolGuWuSQLServer Then
            '        Try
            '            ta_tblMaxIDSQLServer.Update(tblMaxID)
            '        Catch ex As DBConcurrencyException
            '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '        End Try
            '    End If

            'End If

        End If

        ctP = ctP + 1
        frmH.pb1.Value = ctP
        frmH.pb1.Refresh()

        'perform a save action
        boolOR = True
        Call DoThisApplyTemplate()
        boolOR = False

        ctP = ctP + 1

        If ctP > frmH.pb1.Maximum Then
            frmH.pb1.Maximum = frmH.pb1.Maximum + 10
        End If
        frmH.pb1.Value = ctP
        frmH.pb1.Refresh()

        frmH.pb1.Value = frmH.pb1.Maximum
        frmH.pb1.Refresh()

        'Clean up Table Configuration tab

        'set Report Table Config - Show Included
        frmH.rbShowIncludedRTConfig.Checked = True

        'Do the filter
        Call RTFilter()

        'Order the tables
        Try
            Call OrderDGV(frmH.dgvReportTableConfiguration, "INTORDER", "ID_TBLREPORTTABLE")
        Catch ex As Exception

        End Try


        MsgBox("Template application complete.", MsgBoxStyle.Information, "Action complete...")

        'frmH.lblProgress.Visible = False
        'frmH.pb1.Visible = False

        frmH.panProgress.Visible = False
        frmH.panProgress.Refresh()

        frmH.Refresh()
        'SendKeys.Send("%")

        Cursor.Current = Cursors.Default


    End Sub

    Sub CheckForAutoAssignSamplesTable(ByVal idO As Int64, ByVal idN As Int64, ByVal idTStudiesO As Int64, ByVal idTStudiesN As Int64, ByVal boolGetNew As Boolean, ByVal intAdded As Short)

        'intRow is dgvReportTables selected row
        'idO is original id_tblReportTable
        'idN as new id_tblReportTable

        'idTStudiesO is original id_tblStudies
        'idTStudiesN is new id_tblStudies

        Dim maxID As Int64
        Dim maxID1 As Int64
        Dim strF As String
        Dim strF1 As String
        Dim Count1 As Short
        Dim int1 As Int16
        Dim var1, var2
        Dim str1 As String


        'need to get data for old tblStudies
        Dim tblAAS1 As DataTable
        Dim tblAAS2 As DataTable

        If boolGetNew Then
            If boolGuWuOracle Then
                'tblAAS1 = ta_TBLAUTOASSIGNSAMPLES.GetDataBy_ID_TBLSTUDIES(idTStudiesO)
            ElseIf boolGuWuAccess Then
                tblAAS1 = ta_TBLAUTOASSIGNSAMPLESAcc.GetDataBy_ID_TBLSTUDIES(idTStudiesO)
            ElseIf boolGuWuSQLServer Then
                tblAAS1 = ta_TBLAUTOASSIGNSAMPLESSQLServer.GetDataBy_ID_TBLSTUDIES(idTStudiesO)
            End If
            int1 = tblAAS1.Rows.Count 'DEBUG
        Else
            tblAAS1 = tblAutoAssignSamples
        End If

        tblAAS2 = tblAutoAssignSamples

        strF = "ID_TBLREPORTTABLE = " & idO
        Dim rowsS() As DataRow = tblAAS1.Select(strF)

        If intAdded = -1 Then 'added

            maxID = GetMaxID("TBLAUTOASSIGNSAMPLES", 1, True)
            maxID1 = maxID

            Dim nr As DataRow = tblAAS2.NewRow

            nr.BeginEdit()
            For Count1 = 0 To tblAAS2.Columns.Count - 1
                nr(Count1) = rowsS(0).Item(Count1)
            Next

            'set individual values
            maxID = maxID + 1
            nr("ID_TBLSTUDIES") = id_tblStudies ' idTStudiesN
            nr("ID_TBLAUTOASSIGNSAMPLES") = maxID
            nr("ID_TBLREPORTTABLE") = idN

            var1 = nr("ID_TBLSTUDIES")
            var1 = var1 'debug

            nr.EndEdit()
            tblAAS2.Rows.Add(nr)

            'If maxID = maxID1 Then
            'Else
            '    Call PutMaxID("TBLAUTOASSIGNSAMPLES", maxID)
            'End If

        Else

            'modify existing row
            strF1 = "ID_TBLREPORTTABLE = " & idN
            Dim rowsA() As DataRow = tblAAS2.Select(strF1)

            If rowsA.Length = 0 Then
                var1 = var1
            Else
                rowsA(0).BeginEdit()
                For Count1 = 0 To tblAAS2.Columns.Count - 1
                    'ignore some columns
                    str1 = tblAAS2.Columns(Count1).ColumnName
                    Select Case str1
                        Case "ID_TBLAUTOASSIGNSAMPLES"
                        Case "ID_TBLSTUDIES"
                        Case "ID_TBLCONFIGREPORTTABLES"
                        Case "ID_TBLREPORTTABLE"
                        Case Else
                            rowsA(0).Item(Count1) = rowsS(0).Item(Count1)
                    End Select

                Next
            End If



        End If




    End Sub

    Sub CheckForTblProperties(ByVal intRow As Short, ByVal idO As Int64)

        'intRow is dgvReportTables selected row
        'idO is original id_tblReportTable

        Dim dtbl As System.Data.DataTable
        Dim dtbl1 As System.Data.DataTable
        Dim rows() As DataRow
        Dim rows1() As DataRow
        Dim strF As String
        Dim strF1 As String
        Dim Count1 As Short
        Dim Count2 As Short
        Dim tbl As System.Data.DataTable
        Dim str1 As String
        Dim maxID As Int64
        Dim maxID1 As Int64
        Dim var1, var2, var3, var4, var5
        Dim rowsM() As DataRow
        Dim dt As Date
        Dim dgv As DataGridView
        Dim id As Int64
        Dim intS As Short
        Dim intE As Short
        Dim rowsS() As DataRow
        Dim id1 As Int64
        Dim idN As Int64

        dtbl = tblTableProperties
        id = idO
        If intRow = -1 Then
            strF = "ID_TBLSTUDIES = " & id_tblStudies
            strF1 = "ID_TBLSTUDIES = " & id_tblStudies
            intS = 0
        Else
            strF1 = "ID_TBLSTUDIES = " & id_tblStudies
            dgv = frmH.dgvReportTableConfiguration
            idN = dgv("ID_TBLREPORTTABLE", intRow).Value
            id1 = dgv("ID_TBLCONFIGREPORTTABLES", intRow).Value
            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idN
            intS = 0
            intE = 0
        End If
        rows = dtbl.Select(strF)
        dt = Now

        dtbl1 = tblReportTable
        rows1 = dtbl1.Select(strF)
        If intRow = -1 Then
            intE = rows1.Length - 1
        Else
            intE = 0
        End If

        If rows.Length = 0 Then 'add default rows to table

            maxID = GetMaxID("tblTableProperties", 1, False) 'if maxid increment is 1, then getmaxid already does putmaxid, 20190306 LEE: but incremented later
            maxID1 = maxID

            If intRow = -1 Then 'add default tableproperties

                For Count1 = intS To intE

                    'var1 = rows1(Count1).Item("ID_TBLREPORTTABLE")
                    'var2 = rows1(Count1).Item("ID_TBLCONFIGREPORTTABLES")
                    If intRow = -1 Then
                        var1 = rows1(Count1).Item("ID_TBLREPORTTABLE")
                        var2 = rows1(Count1).Item("ID_TBLCONFIGREPORTTABLES")
                    Else
                        var1 = dgv("ID_TBLREPORTTABLE", intRow).Value
                        var2 = dgv("ID_TBLCONFIGREPORTTABLES", intRow).Value
                    End If
                    maxID = maxID + 1

                    Dim row As DataRow = dtbl.NewRow
                    row.BeginEdit()

                    row.Item("ID_TBLTABLEPROPERTIES") = maxID
                    row.Item("ID_TBLREPORTTABLE") = var1
                    row.Item("ID_TBLCONFIGREPORTTABLES") = var2
                    row.Item("ID_TBLSTUDIES") = id_tblStudies

                    row.Item("boolBQLSHOWCONC") = -1
                    row.Item("boolCSSHOWREJVALUES") = -1
                    '20181220 LEE:
                    'Change default of boolCSREPORTACCVALUES
                    row.Item("boolCSREPORTACCVALUES") = 0 ' -1
                    row.Item("boolQCREPORTACCVALUES") = -1
                    row.Item("boolSTATSMEAN") = -1
                    row.Item("boolSTATSSD") = -1
                    row.Item("boolSTATSCV") = -1
                    row.Item("boolTHEORETICAL") = 0
                    row.Item("BOOLINCLANOVA") = -1
                    row.Item("BOOLINCLANOVASUMSTATS") = -1
                    If var2 = 13 Or var2 = 14 Or var2 = 15 Then 'recovery, stock soln tables
                        row.Item("boolSTATSBIAS") = 0
                    Else
                        If var2 = 22 And var2 = 23 Then 'Stock soln and Spiking soln stability
                            row.Item("boolSTATSBIAS") = 0
                        Else
                            row.Item("boolSTATSBIAS") = -1
                        End If

                    End If
                    row.Item("boolSTATSN") = -1
                    row.Item("boolStatsDiff") = 0
                    row.Item("boolStatsDiffCol") = 0
                    row.Item("boolStatsRegr") = 0
                    row.Item("boolStatsNR") = 0
                    row.Item("boolStatsLetter") = 0

                    row.Item("NUMISCRIT1") = 20
                    row.Item("NUMISCRIT1LEVEL") = System.DBNull.Value
                    row.Item("NUMISCRIT2") = System.DBNull.Value

                    row.Item("UPSIZE_TS") = dt

                    row.Item("BOOLBQLLEGEND") = 0
                    row.Item("NUMSAMPLEG1") = 103
                    row.Item("NUMSAMPLEG2") = 0
                    row.Item("NUMSAMPLEG3") = 0
                    row.Item("NUMSAMPLEG4") = 0
                    row.Item("NUMSAMPLES1") = 103
                    row.Item("NUMSAMPLES2") = 102
                    row.Item("NUMSAMPLES3") = 106
                    row.Item("NUMSAMPLES4") = 101
                    row.Item("CHARSAMPLEGAD1") = "ASC"
                    row.Item("CHARSAMPLEGAD2") = "ASC"
                    row.Item("CHARSAMPLEGAD3") = "ASC"
                    row.Item("CHARSAMPLEGAD4") = "ASC"
                    row.Item("CHARSAMPLESAD1") = "ASC"
                    row.Item("CHARSAMPLESAD2") = "ASC"
                    row.Item("CHARSAMPLESAD3") = "ASC"
                    row.Item("CHARSAMPLESAD4") = "ASC"
                    row.Item("BOOLINCLUDEPSAE") = 0

                    row.Item("BOOLRCCONC") = -1
                    row.Item("BOOLRCPA") = 0
                    row.Item("BOOLRCPARATIO") = 0
                    row.Item("BOOLINCLUDEISTBL") = 0
                    row.Item("BOOLMEANACCURACY") = 0
                    row.Item("BOOLNONELEG") = -1
                    row.Item("BOOLPOSLEG") = -1
                    row.Item("BOOLNEGLEG") = 0
                    row.Item("BOOLCUSTOMLEG") = 0

                    row.Item("CHARTITLELEG") = "%Difference = "
                    row.Item("CHARNUMLEG") = "(Mean Old - Mean New) x 100"
                    row.Item("CHARDENLEG") = "(Mean New)"

                    row.Item("BOOLINCLUDEDATE") = 0
                    row.Item("BOOLDIFFERENCE") = -1
                    row.Item("BOOLRECOVERY") = 0
                    row.Item("BOOLINCLUDEWATSONLABELS") = 0

                    row.Item("boolStatsRE") = 0

                    row.Item("BOOLCONVERTTIME") = 0
                    row.Item("BOOLCONVERTTEMP") = 0
                    row.Item("CHARISCONC") = ""

                    row.Item("BOOLINTRARUNSUMSTATS") = 0
                    row.Item("BOOLDOINDREC") = 0

                    row.Item("BOOLISCOMBINELEVELS") = -1
                    row.Item("BOOLREASSAYREASLETTERS") = 0


                    row.Item("CHARCARRYOVERLABEL") = "Blank"
                    row.Item("BOOLMFTABLE") = 0
                    row.Item("BOOLINCLMFCOLS") = 0
                    row.Item("BOOLINCLINTSTDNMF") = 0
                    row.Item("BOOLCALCINTSTDNMF") = 0
                    row.Item("NUMPRECCRITLOTS") = 0
                    row.Item("BOOLREGRULOQ") = 0
                    row.Item("INTQCLEVELGROUP") = 0

                    row.Item("BOOLCONCCOMMENTS") = 0

                    row.Item("BOOLADHOCSTABCOMPCOLUMNS") = 0

                    'BOOLISCOMBINELEVELS
                    'CHARCARRYOVERLABEL
                    'BOOLMFTABLE
                    'BOOLINCLMFCOLS
                    'BOOLINCLINTSTDNMF
                    'BOOLCALCINTSTDNMF
                    'NUMPRECCRITLOTS
                    'BOOLREGRULOQ
                    'INTQCLEVELGROUP


                    row.EndEdit()
                    dtbl.Rows.Add(row)
                Next

            Else 'add duplicate tableproperties

                strF = "ID_TBLREPORTTABLE = " & id & " AND ID_TBLCONFIGREPORTTABLES = " & id1 & " AND ID_TBLSTUDIES = " & id_tblStudies
                rowsS = dtbl.Select(strF)
                var1 = rowsS.Length
                If var1 = 0 Then
                    str1 = "Problem duplicating table."
                    MsgBox(str1, MsgBoxStyle.Information, "Problem duplicating table...")
                Else
                    For Count1 = 0 To 0

                        var1 = idN
                        var2 = id1

                        maxID = maxID + 1

                        Dim row As DataRow = dtbl.NewRow
                        row.BeginEdit()

                        '20190113 LEE:
                        'why doing hard entries?
                        'Use loop instead
                        For Count2 = 0 To dtbl.Columns.Count - 1

                            var3 = UCase(dtbl.Columns(Count2).ColumnName)

                            Select Case var3
                                Case "ID_TBLTABLEPROPERTIES"
                                    row.Item("ID_TBLTABLEPROPERTIES") = maxID
                                Case "ID_TBLREPORTTABLE"
                                    row.Item("ID_TBLREPORTTABLE") = var1
                                Case "ID_TBLCONFIGREPORTTABLES"
                                    row.Item("ID_TBLCONFIGREPORTTABLES") = var2
                                Case "ID_TBLSTUDIES"
                                    row.Item("ID_TBLSTUDIES") = id_tblStudies
                                Case "UPSIZE_TS"
                                    row.Item("UPSIZE_TS") = dt
                                Case Else
                                    row.Item(var3) = rowsS(0).Item(var3)
                            End Select

                        Next

                        row.EndEdit()
                        dtbl.Rows.Add(row)

                    Next
                End If
            End If

            If maxID1 = maxID Then 'ignore
            Else

                Call PutMaxID("tblTableProperties", maxID)

            End If

            If boolGuWuOracle Then
                Try
                    ta_tblTableProperties.Update(tblTableProperties)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLTABLEPROPERTIES.Merge('ds2005.TBLTABLEPROPERTIES, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblTablePropertiesAcc.Update(tblTableProperties)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLTABLEPROPERTIES.Merge('ds2005Acc.TBLTABLEPROPERTIES, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblTablePropertiesSQLServer.Update(tblTableProperties)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLTABLEPROPERTIES.Merge('ds2005Acc.TBLTABLEPROPERTIES, True)
                End Try
            End If


        Else

            'check for new tables

            Dim strS As String
            Dim int1 As Int64
            Dim int2 As Int64
            Dim intF As Short

            strS = "ID_TBLCONFIGREPORTTABLES ASC"
            intRow = -1
            strF = "ID_TBLSTUDIES = " & id_tblStudies
            strF1 = "ID_TBLSTUDIES = " & id_tblStudies
            intS = 0

            Dim dv As System.Data.DataView = New DataView(dtbl, strF, strS, DataViewRowState.CurrentRows)
            'dtbl = tblTableProperties
            rows1 = dtbl1.Select(strF) 'tblReportTable

            ''get max id
            maxID = GetMaxID("tblTableProperties", 1, False) 'if maxid increment is 1, then getmaxid already does putmaxid
            maxID1 = maxID

            Dim idCR As Int64

            For Count1 = 0 To rows1.Length - 1
                int1 = rows1(Count1).Item("ID_TBLCONFIGREPORTTABLES")
                intF = dv.Find(int1)
                If intF = -1 Then
                    'add to dv
                    Dim rowView As DataRowView = dv.AddNew

                    ' Change values in the DataViewRow.
                    var1 = rows1(Count1).Item("ID_TBLREPORTTABLE")
                    var2 = rows1(Count1).Item("ID_TBLCONFIGREPORTTABLES")
                    idCR = var2

                    maxID = maxID + 1

                    rowView("ID_TBLTABLEPROPERTIES") = maxID
                    rowView("ID_TBLREPORTTABLE") = var1
                    rowView("ID_TBLCONFIGREPORTTABLES") = var2
                    rowView("ID_TBLSTUDIES") = id_tblStudies

                    rowView("boolBQLSHOWCONC") = -1
                    rowView("boolCSSHOWREJVALUES") = -1
                    '20181220 LEE:
                    'Change default of boolCSREPORTACCVALUES
                    rowView("boolCSREPORTACCVALUES") = 0 '-1
                    rowView("boolQCREPORTACCVALUES") = -1
                    rowView("boolSTATSMEAN") = -1
                    rowView("boolSTATSSD") = -1
                    rowView("boolSTATSCV") = -1
                    rowView("boolTHEORETICAL") = 0
                    rowView("BOOLINCLANOVA") = -1
                    rowView("BOOLINCLANOVASUMSTATS") = -1
                    rowView("boolSTATSBIAS") = -1
                    If var2 = 13 Or var2 = 14 Or var2 = 15 Then 'recovery, stock soln tables
                        rowView("boolSTATSBIAS") = 0
                    Else
                        If var2 = 22 And var2 = 23 Then 'Stock soln and Spiking soln stability
                            rowView("boolSTATSBIAS") = 0
                        Else
                            rowView("boolSTATSBIAS") = -1
                        End If

                    End If

                    rowView("boolSTATSN") = -1
                    rowView("boolStatsDiff") = 0
                    rowView("boolStatsDiffCol") = 0
                    rowView("boolStatsRegr") = 0
                    rowView("boolStatsNR") = 0
                    rowView("boolStatsLetter") = 0

                    rowView("NUMISCRIT1") = 20
                    rowView("NUMISCRIT1LEVEL") = System.DBNull.Value
                    rowView("NUMISCRIT2") = System.DBNull.Value

                    rowView("UPSIZE_TS") = dt

                    rowView("BOOLBQLLEGEND") = 0
                    rowView("NUMSAMPLEG1") = 103
                    rowView("NUMSAMPLEG2") = 0
                    rowView("NUMSAMPLEG3") = 0
                    rowView("NUMSAMPLEG4") = 0
                    rowView("NUMSAMPLES1") = 103
                    rowView("NUMSAMPLES2") = 102
                    rowView("NUMSAMPLES3") = 106
                    rowView("NUMSAMPLES4") = 101
                    rowView("CHARSAMPLEGAD1") = "ASC"
                    rowView("CHARSAMPLEGAD2") = "ASC"
                    rowView("CHARSAMPLEGAD3") = "ASC"
                    rowView("CHARSAMPLEGAD4") = "ASC"
                    rowView("CHARSAMPLESAD1") = "ASC"
                    rowView("CHARSAMPLESAD2") = "ASC"
                    rowView("CHARSAMPLESAD3") = "ASC"
                    rowView("CHARSAMPLESAD4") = "ASC"
                    rowView("BOOLINCLUDEPSAE") = 0

                    rowView("BOOLRCCONC") = -1
                    rowView("BOOLRCPA") = 0
                    rowView("BOOLRCPARATIO") = 0
                    rowView("BOOLINCLUDEISTBL") = 0
                    rowView("BOOLMEANACCURACY") = 0
                    rowView("BOOLNONELEG") = -1
                    rowView("BOOLPOSLEG") = -1
                    rowView("BOOLNEGLEG") = 0
                    rowView("BOOLCUSTOMLEG") = 0

                    rowView("CHARTITLELEG") = "%Difference = "
                    rowView("CHARNUMLEG") = "(Mean Old - Mean New) x 100"
                    rowView("CHARDENLEG") = "(Mean New)"

                    rowView("BOOLINCLUDEDATE") = 0
                    rowView("BOOLDIFFERENCE") = 0
                    rowView("BOOLRECOVERY") = 0
                    rowView("BOOLINCLUDEWATSONLABELS") = 0

                    rowView("BOOLSTATSRE") = 0

                    rowView("BOOLCONVERTTIME") = 0
                    rowView("BOOLCONVERTTEMP") = 0
                    rowView("CHARISCONC") = ""

                    rowView("BOOLINTRARUNSUMSTATS") = 0
                    rowView("BOOLDOINDREC") = 0

                    'BOOLISCOMBINELEVELS
                    rowView("BOOLISCOMBINELEVELS") = -1
                    rowView("BOOLREASSAYREASLETTERS") = 0

                    If boolAllowAcc(idCR) Then
                    Else
                        rowView("boolSTATSBIAS") = 0
                        rowView("boolStatsDiff") = 0
                        rowView("boolStatsDiffCol") = 0
                        rowView("BOOLSTATSRE") = 0
                        rowView("boolTHEORETICAL") = 0
                    End If

                    rowView.Item("CHARCARRYOVERLABEL") = "Blank"
                    rowView.Item("BOOLMFTABLE") = 0
                    rowView.Item("BOOLINCLMFCOLS") = 0
                    rowView.Item("BOOLINCLINTSTDNMF") = 0
                    rowView.Item("BOOLCALCINTSTDNMF") = 0
                    rowView.Item("NUMPRECCRITLOTS") = 0
                    rowView.Item("BOOLREGRULOQ") = 0
                    rowView.Item("INTQCLEVELGROUP") = 0

                    rowView.Item("BOOLCONCCOMMENTS") = 0

                    rowView.Item("BOOLADHOCSTABCOMPCOLUMNS") = 0

                    rowView.EndEdit()

                End If
            Next

            If maxID1 = maxID Then 'ignore
            Else


                Call PutMaxID("tblTableProperties", maxID)

                'rowsM(0).BeginEdit()
                'rowsM(0).Item("NUMMAXID") = maxID
                'rowsM(0).EndEdit()
                'If boolGuWuOracle Then
                '    Try
                '        ta_tblMaxID.Update(tblMaxID)
                '    Catch ex As DBConcurrencyException
                '        'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
                '    End Try
                'ElseIf boolGuWuAccess Then
                '    Try
                '        ta_tblMaxIDAcc.Update(tblMaxID)
                '    Catch ex As DBConcurrencyException
                '        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
                '    End Try
                'ElseIf boolGuWuSQLServer Then
                '    Try
                '        ta_tblMaxIDSQLServer.Update(tblMaxID)
                '    Catch ex As DBConcurrencyException
                '        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
                '    End Try
                'End If

                If boolGuWuOracle Then
                    Try
                        ta_tblTableProperties.Update(tblTableProperties)
                    Catch ex As DBConcurrencyException
                        'ds2005.TBLTABLEPROPERTIES.Merge('ds2005.TBLTABLEPROPERTIES, True)
                    End Try
                ElseIf boolGuWuAccess Then
                    Try
                        ta_tblTablePropertiesAcc.Update(tblTableProperties)
                    Catch ex As DBConcurrencyException
                        'ds2005Acc.TBLTABLEPROPERTIES.Merge('ds2005Acc.TBLTABLEPROPERTIES, True)
                    End Try
                ElseIf boolGuWuSQLServer Then
                    Try
                        ta_tblTablePropertiesSQLServer.Update(tblTableProperties)
                    Catch ex As DBConcurrencyException
                        'ds2005Acc.TBLTABLEPROPERTIES.Merge('ds2005Acc.TBLTABLEPROPERTIES, True)
                    End Try
                End If

            End If

        End If

    End Sub

    Sub AssessSampleAssignment()

        Dim tblA As System.Data.DataTable
        Dim rowsA() As DataRow
        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim int1 As Short
        Dim int2 As Short
        'Dim dgv1 As DataGridView
        'Dim dv1 as system.data.dataview
        Dim int1A As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim strF As String
        Dim bool As Boolean
        Dim colrW, colrS, colrT, colrA
        Dim boolTimer As Boolean
        Dim boolI As Boolean
        Dim boolIS As Boolean
        Dim int3 As Short
        Dim boolA As Boolean
        Dim boolISA As Boolean
        Dim str1 As String
        Dim str2 As String

        Dim tblAnal As System.Data.DataTable
        Dim rowsAnal() As DataRow
        Dim intRowsAnal As Short
        Dim var1, var2, var3, var4, var10

        Dim boolCont As Boolean
        Dim boolInclude As Boolean
        Dim rows() As DataRow

        Dim strMatrix As String
        Dim idT As Long

        strF = "ID_TBLSTUDIES = " & id_tblStudies
        Dim strS As String
        strS = "ID_TBLSTUDIES ASC"
        Dim dvProps As System.Data.DataView = New DataView(tblTableProperties, strF, strS, DataViewRowState.CurrentRows)
        Dim intProps As Short

        tblAnal = tblAnalytesHome
        int1A = tblAnal.Rows.Count

        boolI = False 'for internal standard
        boolInclude = False
        boolTimer = False
        tblA = tblAssignedSamples
        'dgv = frmh.dgvTables
        dgv = frmH.dgvReportTableConfiguration
        dv = dgv.DataSource
        Try
            int1 = dv.Count
        Catch ex As Exception
            int1 = 0
        End Try

        'dgv1 = frmh.dgvAnalytes
        'dv1 = dgv1.DataSource
        'int1A = dv1.Count

        'pesky
        Dim nP1 As New Padding(0, 6, 0, 6)
        dgv.DefaultCellStyle.Padding = nP1


        Dim idCT As Long
        Dim intOIS As Short
        Dim boolOIS As Boolean

        For Count1 = 0 To int1 - 1

            var1 = dv(Count1).Item("ID_TBLCONFIGREPORTTABLES")
            var10 = dv(Count1).Item("BOOLINCLUDE") 'returns boolean
            idT = dv(Count1).Item("ID_TBLREPORTTABLE")
            idCT = var1 ' dv(Count1).Item("ID_CONFIGREPORTTABLES")

            strF = "ID_TBLCONFIGREPORTTABLES = " & var1
            Erase rows
            rows = tblConfigReportTables.Select(strF)
            boolIS = rows(0).Item("boolincludeis")
            'further evaluate boolIS because tblReportProperties can override
            strF = "ID_TBLREPORTTABLE = " & idT
            dvProps.RowFilter = strF
            If dvProps.Count = 0 Then
            Else
                intProps = NZ(dvProps(0).Item("BOOLINCLUDEISTBL"), 0)
                intOIS = NZ(dvProps(0).Item("BOOLCUSTOMLEG"), 0)
                If intProps = -1 Then
                    boolIS = True
                    boolOIS = False
                    Select Case idCT
                        Case 22
                            If intOIS = -1 Then
                                boolOIS = True
                            Else
                                boolOIS = False
                            End If
                    End Select
                Else
                    boolIS = False
                    boolOIS = False
                End If
            End If
            bool = dv(Count1).Item("boolRequiresSampleAssignment")
            colrT = Color.White
            boolI = -1
            boolCont = True
            If boolI = -1 Then
                If bool = -1 And var10 Then
                    'loop through Analytes
                    For Count2 = 0 To int1A - 1
                        colrA = Color.White
                        str2 = tblAnal.Rows.Item(Count2).Item("IsIntStd")
                        boolISA = False
                        If StrComp(str2, "Yes", CompareMethod.Text) = 0 Then
                            boolISA = True
                        Else
                            boolISA = False
                        End If
                        If boolIS And idCT <> 35 Then 'evaluate analytes for int std, ignore 35: Carryover'If boolIS Then 'evaluate analytes for int std
                            If boolISA Then 'continue
                                str1 = tblAnal.Rows.Item(Count2).Item("AnalyteDescription")
                                var3 = NZ(tblAnal.Rows.Item(Count2).Item("AnalyteIndex"), 0)
                                var4 = NZ(tblAnal.Rows.Item(Count2).Item("MasterAssayID"), 0)
                                strMatrix = NZ(tblAnal.Rows.Item(Count2).Item("MATRIX"), "AA")
                                'determine if analyte is checked
                                'var2 = dv(Count1).Item(str1)
                                'If var2 = -1 Then 'continue
                                'check for table entry for
                                var1 = dv(Count1).Item("id_tblConfigReportTables")
                                strF = "id_tblConfigReportTables = " & var1 & " AND CHARANALYTE = '" & CleanText(str1) & "' AND BOOLINTSTD = -1 AND ID_TBLREPORTTABLE = " & idT ' & " AND MATRIX '" & strMatrix & "'"
                                Erase rowsA
                                rowsA = tblA.Select(strF)

                                int2 = rowsA.Length
                                If int2 = 0 Then 'color the row
                                    colrT = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
                                    colrA = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
                                    boolTimer = True
                                    boolCont = False
                                End If
                                'End If
                            Else '

                                If boolOIS Then
                                Else
                                    str1 = tblAnal.Rows.Item(Count2).Item("AnalyteDescription")
                                    var3 = NZ(tblAnal.Rows.Item(Count2).Item("ANALYTEID"), 0)
                                    var4 = NZ(tblAnal.Rows.Item(Count2).Item("MasterAssayID"), 0)
                                    strMatrix = NZ(tblAnal.Rows.Item(Count2).Item("MATRIX"), "AA")
                                    'determine if analyte is checked
                                    Try
                                        var2 = dv(Count1).Item(str1)
                                        If var2 = -1 Then 'continue
                                            'check for table entry for
                                            var1 = dv(Count1).Item("id_tblConfigReportTables")
                                            'strF = "id_tblConfigReportTables = " & var1 & " AND AnalyteIndex = " & var3 & " AND MasterAssayID = " & var4 & " AND BOOLINTSTD = 0 AND ID_TBLREPORTTABLE = " & idT
                                            strF = "id_tblConfigReportTables = " & var1 & " AND ANALYTEID = " & var3 & " AND BOOLINTSTD = 0 AND ID_TBLREPORTTABLE = " & idT & " AND SAMPLETYPEID = '" & strMatrix & "'"
                                            Erase rowsA
                                            rowsA = tblA.Select(strF)
                                            int2 = rowsA.Length
                                            If int2 = 0 Then 'color the row
                                                colrT = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
                                                colrA = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
                                                boolTimer = True
                                                boolCont = False
                                            End If
                                        Else
                                        End If
                                    Catch ex As Exception
                                        Dim abc As String
                                        abc = str1 & " does not exist"
                                    End Try
                                End If
                             

                            End If
                        Else 'ignore analyte int std
                            If boolISA Then
                            Else
                                str1 = tblAnal.Rows.Item(Count2).Item("AnalyteDescription")
                                var3 = NZ(tblAnal.Rows.Item(Count2).Item("ANALYTEID"), 0)
                                var4 = NZ(tblAnal.Rows.Item(Count2).Item("MasterAssayID"), 0)
                                strMatrix = NZ(tblAnal.Rows.Item(Count2).Item("MATRIX"), "AA")
                                'determine if analyte is checked

                                Try
                                    var2 = dv(Count1).Item(str1)
                                    If var2 = -1 Then 'continue
                                        'check for table entry for
                                        var1 = dv(Count1).Item("id_tblConfigReportTables")
                                        'strF = "id_tblConfigReportTables = " & var1 & " AND AnalyteIndex = " & var3 & " AND MasterAssayID = " & var4 & " AND BOOLINTSTD = 0 AND ID_TBLREPORTTABLE = " & idT
                                        strF = "id_tblConfigReportTables = " & var1 & " AND ANALYTEID = " & var3 & " AND BOOLINTSTD = 0 AND ID_TBLREPORTTABLE = " & idT & " AND SAMPLETYPEID = '" & strMatrix & "'"
                                        Erase rowsA
                                        rowsA = tblA.Select(strF)
                                        int2 = rowsA.Length
                                        If int2 = 0 Then 'color the row
                                            colrT = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
                                            colrA = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
                                            boolTimer = True
                                            boolCont = False
                                        End If
                                    Else
                                    End If
                                Catch ex As Exception
                                    Dim abcd As String
                                    abcd = str1 & " doesn't exist"
                                End Try
                            End If
                        End If
                    Next
                End If
            End If
            If colrT = dgv.Rows.Item(Count1).DefaultCellStyle.BackColor Then
            Else
                dgv.Rows.Item(Count1).DefaultCellStyle.BackColor = colrT
            End If
        Next

        If boolTimer Then
            frmH.TimerRTC.Enabled = True
            frmH.llblAssignedSamples.Visible = True
        Else
            frmH.TimerRTC.Enabled = False
            frmH.cmdAssignSamples.BackColor = Color.Gainsboro
            frmH.llblAssignedSamples.Visible = False
        End If

    End Sub

    Sub ColorMethodValRows()

        Dim int1 As Short
        Dim int2 As Long
        Dim Count1 As Short

        Dim str1 As String
        Dim dgv As DataGridView
        Dim var1, var2
        Dim str2 As String
        Dim strM As String

        Dim boolVal As Boolean
        boolVal = False
        Dim dgvR As DataGridView
        Dim idR As Int64
        Dim intCol As Short
        Dim colrT, colrA

        dgv = frmH.dgvMethodValData
        dgvR = frmH.dgvReports

        Dim intRows As Short

        intRows = dgv.Rows.Count

        If dgvR.Rows.Count = 0 Then
        Else
            idR = dgvR("ID_TBLCONFIGREPORTTYPE", 0).Value
            If idR > 1 And idR < 1000 Then
                boolVal = True
            Else
                boolVal = False
            End If
        End If

        If boolVal = False Then
            frmH.lbl1.Visible = True
            frmH.lbl2.Visible = False
        Else
            frmH.lbl1.Visible = False
            frmH.lbl2.Visible = True
        End If

        For Count1 = 0 To intRows - 1

            Try

                'If dgvR.Rows.Count = 0 Then
                'Else
                '    idR = dgvR("ID_TBLCONFIGREPORTTYPE", 0).Value
                '    If idR > 1 And idR < 1000 Then
                '        boolVal = True
                '    Else
                '        boolVal = False
                '    End If
                'End If

                'If boolVal = False Then
                '    frmH.lbl1.Visible = True
                '    frmH.lbl2.Visible = False
                'Else
                '    frmH.lbl1.Visible = False
                '    frmH.lbl2.Visible = True
                'End If

                If boolVal Then 'check for readonly
                    var1 = dgv(0, Count1).Value

                    strM = ""
                    colrT = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
                    Select Case var1
                        Case "Validation Corporate Study/Project Number"
                            strM = "To change '" & var1 & "', change the 'Corporate Study/Project Number' text box on the 'Add/Edit Top Level Data' page."
                            dgv.Rows(Count1).DefaultCellStyle.BackColor = Color.FromArgb(255, 207, 176) 'orange
                        Case "Validation Protocol Number"
                            strM = "To change '" & var1 & "', change the 'Protocol Number' text box on the 'Add/Edit Top Level Data' page."
                            dgv.Rows(Count1).DefaultCellStyle.BackColor = Color.FromArgb(255, 207, 176) 'orange
                        Case "Validation Report Title"
                            strM = "To change '" & var1 & "', change the 'Report Title' text box of the 'Configured Reports' table on the 'Choose Study & Report' page."
                            dgv.Rows(Count1).DefaultCellStyle.BackColor = Color.FromArgb(255, 207, 176) 'orange
                        Case "Validation Report Number"
                            strM = "To change '" & var1 & "', change the 'Report Number' text box of the 'Configured Reports' table on the  'Choose Study & Report' page."
                            dgv.Rows(Count1).DefaultCellStyle.BackColor = Color.FromArgb(255, 207, 176) 'orange
                            'Case "Analytical Method Type" '20190212 LEE: deprecated
                            '    strM = "To change '" & var1 & "', change the 'Assay Technique' dropdown box on the 'Add/Edit Top Level Data' page."
                            '    dgv.Rows(Count1).DefaultCellStyle.BackColor = Color.FromArgb(255, 207, 176)'orange
                        Case "Assay Technique" '20190212 LEE: 
                            strM = "To change '" & var1 & "', change the 'Assay Technique' dropdown box on the 'Add/Edit Top Level Data' page."
                            dgv.Rows(Count1).DefaultCellStyle.BackColor = Color.FromArgb(255, 207, 176)

                            '20190220 LEE:
                        Case "Freeze/Thaw Stability", "Bench-top Stability", "Process Stability", "Reinjection Stability", "Batch Reinjection Stability", "Long-term Storage Stability", "Whole Blood Stability", "Stock Solution Stability", "Spiking Solution Stability", "Autosampler Stability" '20190212 LEE: 
                            strM = "To change '" & var1 & "', corresponding Stability Conditions Summary cell in the Advanced Table Configuration window - Stability Tab."
                            dgv.Rows(Count1).DefaultCellStyle.BackColor = Color.FromArgb(255, 207, 176)
                        Case "Maximum # of Freeze/thaw Cycles" '20190212 LEE: 
                            strM = "To change '" & var1 & "', corresponding [#Cylces] Information cell in the Advanced Table Configuration window - Stability Tab."
                            dgv.Rows(Count1).DefaultCellStyle.BackColor = Color.FromArgb(255, 207, 176)
                        Case "Anticoagulant/Preservative" '20190220 LEE: 
                            strM = "To change '" & var1 & "', change the 'Anticoagulant' dropdown box on the 'Add/Edit Top Level Data' page."
                            dgv.Rows(Count1).DefaultCellStyle.BackColor = Color.FromArgb(255, 207, 176)
                        Case Else
                            colrT = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer)) 'white
                    End Select

                Else
                    'colrT = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer)) 'white
                    dgv.Rows(Count1).DefaultCellStyle.BackColor = Color.FromArgb(255, 255, 255) 'white
                End If

                'dgv.Rows.Item(Count1).DefaultCellStyle.BackColor = colrT

                'dgv.Rows(Count1).DefaultCellStyle.BackColor = Color.FromArgb(231, 86, 56)

            Catch ex As Exception

                var1 = ex.Message

            End Try

        Next

        var1 = var1


        'If int2 = 0 Then 'color the row
        '    colrT = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
        '    colrA = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
        'End If

        'If colrT = dgv.Rows.Item(Count1).DefaultCellStyle.BackColor Then
        'Else
        '    dgv.Rows.Item(Count1).DefaultCellStyle.BackColor = colrT
        'End If


    End Sub

    Sub HookAnalysis()
        Dim tbl As System.Data.DataTable
        Dim tbl1 As System.Data.DataTable
        Dim intTab As Short
        Dim intHook As Short
        Dim strF As String
        Dim rows() As DataRow
        Dim rows1() As DataRow
        Dim str1 As String
        Dim int1 As Short


        tbl = tblHooks
        tbl1 = tblTab1
        str1 = frmH.lbxTab1.SelectedItem
        strF = "CHARITEM = '" & str1 & "' AND INTFORM = 1"
        rows1 = tbl1.Select(strF)
        int1 = rows1(0).Item("ID_TBLTAB1")
        strF = "ID_TBLTAB1 = " & int1
        rows = tbl.Select(strF)
        If rows.Length = 0 Then 'make button invisible
            frmH.cmdHook.Visible = False
            frmH.cmdHook.Enabled = False
        Else
            int1 = rows(0).Item("BOOLSHOW")
            If int1 = -1 Then 'make button visible
                frmH.cmdHook.Location = New System.Drawing.Point(939, 97 + frmH.cmdHook.Height)
                frmH.cmdHook.Visible = True
                frmH.cmdHook.Enabled = True
                frmH.cmdHook.BringToFront()
            Else
                frmH.cmdHook.Visible = False
                frmH.cmdHook.Enabled = False
            End If
        End If



    End Sub

    Sub HookFill_CRL_AnalRefStandard()
        Dim constr As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim strDB As String
        Dim strMDW As String
        Dim tbl As System.Data.DataTable
        Dim strF As String
        Dim int1 As Short
        Dim strUID As String
        Dim strPswd As String
        Dim strProv As String
        Dim cnstr As String
        Dim rows() As DataRow
        Dim strErr As String


        str1 = frmH.lbxTab1.SelectedItem
        strF = "CHARHOOK = 'CRLWor_AnalRefStandard'"
        tbl = tblHooks
        rows = tbl.Select(strF)

        If rows.Length = 0 Then 'ignore everything
            'str1 = "Hook named 'AnalRefStandard' has not been configured." & Chr(10)
            'str1 = str1 & "Please contact your StudyDoc Administrator."
            'MsgBox(str1, MsgBoxStyle.Information, "Proper Hook not configured...")
            GoTo end1
        End If

        'strUID = "User ID=larry_e;"
        'strPswd = "Password=gwoman22;"
        'constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\GubbsInc\BusinessApps\MaterialReceipt\MatReceiptTables.mdb;Jet OLEDB:System database=C:\GubbsInc\BusinessApps\SecurityDefinitions\BAC0002.MDW"
        strUID = "User ID=" & rows(0).Item("CHARUID") & ";"
        'strPswd = "Password=" & Coding(Decode(rows(0).Item("CHARPSWD"), True), False) & ";"
        strPswd = "Password=" & PasswordUnEncrypt(NZ(rows(0).Item("CHARPSWD"), "NA")) & ";"
        strProv = rows(0).Item("CHARCONNECTIONSTRING") & ";"
        cnstr = strProv & strUID & strPswd


        'fill tblHooks1 with pertinent data
        'If boolANSI Then
        'str1 = "SELECT tblCode.chemId, tblCode.ChemName, tblCode.AnotherName, tblBottleSource.BottleId, tblBottleSource.StoreTemp, tblBottleSource.PhyColor, tblBottleSource.PhyState, tblBottleSource.[Lot#], tblBottleSource.ReceiptDate, tblBottleSource.ExpDate, tblBottleSource.Amt, tblBottleSource.SourceName, tblBottleSource.[% Purity], tblBottleSource.corrWater, tblBottleSource.Comment, tblBottleSource.btlUnits "
        'str2 = "FROM tblCode INNER JOIN tblBottleSource ON tblCode.chemId = tblBottleSource.ChemId "
        'str3 = "WHERE (((tblBottleSource.Disposition)='In Use') AND ((tblBottleSource.StorageDept) Like 'BAC%')) "
        'str4 = "ORDER BY tblCode.chemId;"
        'strSQL = str1 & str2 & str3 & str4
        'Else

        str1 = "SELECT tblCode.chemId, tblCode.ChemName, tblCode.AnotherName, tblBottleSource.BottleId, tblBottleSource.StoreTemp, tblBottleSource.PhyColor, tblBottleSource.PhyState, tblBottleSource.[Lot#], tblBottleSource.ReceiptDate, tblBottleSource.ExpDate, tblBottleSource.Amt, tblBottleSource.SourceName, tblBottleSource.[% Purity], tblBottleSource.corrWater, tblBottleSource.Comment, tblBottleSource.btlUnits "
        str2 = "FROM tblCode, tblBottleSource "
        str2 = str2 & "WHERE tblCode.chemId = tblBottleSource.ChemId "
        str3 = "AND (((tblBottleSource.Disposition)='In Use') AND ((tblBottleSource.StorageDept) Like 'BAC%')) "
        str4 = "ORDER BY tblCode.chemId;"
        strSQL = str1 & str2 & str3 & str4

        'sample
        'str1 = "SELECT DISTINCT ASSAYANALYTES.STUDYID, ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.ANALYTEINDEX, GLOBALANALYTES.ANALYTEDESCRIPTION, GLOBALANALYTES.PROJECTID, ASSAY.MASTERASSAYID "
        'str2 = "FROM ASSAYANALYTES, ASSAY, GLOBALANALYTES, STUDY "
        'str2 = str2 & "WHERE((ASSAYANALYTES.ANALYTEID = GLOBALANALYTES.GLOBALANALYTEID) AND (ASSAYANALYTES.STUDYID = STUDY.STUDYID)) AND ((ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID)) "
        'str3 = "AND (((ASSAYANALYTES.STUDYID) = " & wStudyID & ") And ((GLOBALANALYTES.ACTIVE) = -1)) "
        'str4 = "ORDER BY ASSAYANALYTES.ANALYTEID;"
        'sample
        'End If
        ''''''''''''''''''''''''''console.writeline(strSQL)

        'str1 = "SELECT tblBottleSource.BottleId "
        'str2 = "FROM tblBottleSource;"
        ''str3 = "WHERE (((tblBottleSource.Disposition)='In Use') AND ((tblBottleSource.StorageDept) Like 'BAC*')) "
        ''str4 = "ORDER BY tblBottleSource.chemId;"
        'strSQL = str1 & str2 ' & str3 & str4


        'tblHook1.Clear()
        Err.Clear()
        strErr = ""
        On Error GoTo err1
        Dim con As New OleDb.OleDbConnection(cnstr)
        On Error GoTo err2
        Dim da As New OleDb.OleDbDataAdapter(strSQL, con)
        On Error GoTo err3
        tblHook1.Clear()
        tblHook1.AcceptChanges()
        tblHook1.BeginLoadData()
        da.Fill(tblHook1) 'Retrieve data into DataTable.
        tblHook1.EndLoadData()
        boolHook1 = True
        On Error GoTo 0

        ''''''''''''''''''''''''''console.writeline(cnstr)

        'Dim var1
        'var1 = tblHook1.Columns.Count
        'Dim Count1 As Short
        'For Count1 = 0 To var1 - 1
        '    '''''''''''''''''''''''''console.writeline(tblHook1.Columns.item(Count1).ColumnName)
        'Next

err1:
        If Err.Number <> 0 Then
            Err.Clear()
            On Error GoTo 0
            boolHook1 = False
            strErr = "Hmmm. There seems to be a problem with the connectionstring for the CRL-BAC Material Receipt Tables datatabase."
            strErr = strErr & Chr(10) & Chr(10) & "Please contact your StudyDoc system administrator."
            MsgBox(strErr)
            GoTo end1
        End If

err2:
        If Err.Number <> 0 Then
            Err.Clear()
            On Error GoTo 0
            boolHook1 = False
            strErr = "Hmmm. There seems to be a problem establishing a dataadapter for the CRL-BAC Material Receipt Tables datatabase."
            strErr = strErr & Chr(10) & Chr(10) & "Please contact your StudyDoc system administrator."
            MsgBox(strErr)
            GoTo end2
        End If

err3:
        If Err.Number <> 0 Then
            Err.Clear()
            On Error GoTo 0
            boolHook1 = False
            strErr = "Hmmm. There seems to be a problem loading tables from the CRL-BAC Material Receipt Tables datatabase."
            strErr = strErr & Chr(10) & Chr(10) & "Please contact your StudyDoc system administrator."
            MsgBox(strErr)
            GoTo end3
        End If

end3:
        da.Dispose()
end2:

        con.Close()
        con.Dispose()

end1:


        If boolHook1 Then
            int1 = 0
        Else 'update boolErr
            int1 = -1
        End If
        rows(0).BeginEdit()
        rows(0).Item("BOOLERROR") = int1
        rows(0).EndEdit()

        If boolGuWuOracle Then
            On Error Resume Next
            ta_tblHooks.Update(tblHooks)
            If Err.Number <> 0 Then
                ''msgbox("aaChromatography: " & ex.Message)
                'ds2005.TBLHOOKS.Merge('ds2005.TBLHOOKS, True)
            End If

        ElseIf boolGuWuAccess Then

            On Error Resume Next
            ta_tblHooksAcc.Update(tblHooks)
            If Err.Number <> 0 Then
                ''msgbox("aaChromatography: " & ex.Message)
                'ds2005Acc.TBLHOOKS.Merge('ds2005Acc.TBLHOOKS, True)
            End If
        ElseIf boolGuWuSQLServer Then

            On Error Resume Next
            ta_tblHooksSQLServer.Update(tblHooks)
            If Err.Number <> 0 Then
                ''msgbox("aaChromatography: " & ex.Message)
                'ds2005Acc.TBLHOOKS.Merge('ds2005Acc.TBLHOOKS, True)
            End If
        End If



        'On Error GoTo 0


    End Sub

    Sub Create_tblQCStds()

        Dim tbl As System.Data.DataTable

        tbl = tblQCStds

        Dim col1 As New DataColumn
        col1.ColumnName = "AnalyteDescription"
        col1.Caption = "Analyte"
        col1.DataType = System.Type.GetType("System.String")
        tbl.Columns.Add(col1)

        Dim col2 As New DataColumn
        col2.ColumnName = "LevelNumber"
        col2.Caption = "Level"
        col2.DataType = System.Type.GetType("System.Int16")
        tbl.Columns.Add(col2)

        Dim col3 As New DataColumn
        col3.ColumnName = "Concentration"
        col3.Caption = "Conc."
        col3.DataType = System.Type.GetType("System.Decimal")
        tbl.Columns.Add(col3)

        Dim col8 As New DataColumn
        col8.ColumnName = "QCNAME"
        col8.Caption = "QC Name"
        col8.DataType = System.Type.GetType("System.String")
        tbl.Columns.Add(col8)

        Dim col4 As New DataColumn
        col4.ColumnName = "NumReps"
        col4.Caption = "# of Reps"
        col4.DataType = System.Type.GetType("System.Int16")
        tbl.Columns.Add(col4)

        Dim col9 As New DataColumn
        col9.ColumnName = "AssayID"
        col9.Caption = "AssayID"
        col9.DataType = System.Type.GetType("System.Int64")
        tbl.Columns.Add(col9)

        Dim col5 As New DataColumn
        col5.ColumnName = "MasterAssayID"
        col5.Caption = "MasterAssayID"
        col5.DataType = System.Type.GetType("System.Int64")
        tbl.Columns.Add(col5)

        Dim col6 As New DataColumn
        col6.ColumnName = "ID"
        col6.Caption = "Analyte ID"
        col6.DataType = System.Type.GetType("System.Int64")
        tbl.Columns.Add(col6)

        Dim col7 As New DataColumn
        col7.ColumnName = "Index"
        col7.Caption = "Analyte Index"
        col7.DataType = System.Type.GetType("System.Int16")
        tbl.Columns.Add(col7)

        Dim col10 As New DataColumn
        col10.ColumnName = "FlagPercent"
        col10.Caption = "Flag Percent"
        tbl.Columns.Add(col10)

    End Sub
    Sub GetMethodInfo()

        Dim dv As System.Data.DataView
        Dim Count1 As Short
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim str1 As String
        Dim str2 As String
        Dim dgv As DataGridView
        Dim dgv2 As DataGridView
        Dim dvMVD As System.Data.DataView
        Dim strCol As String
        Dim strAC As String
        Dim numSSize As String
        Dim strSpecies As String
        Dim strMatrix As String
        Dim dvDC As System.Data.DataView
        Dim dvDW As System.Data.DataView
        Dim Count2 As Short
        Dim strRet As String
        Dim strFind As String
        Dim var1
        Dim boolMVal As Boolean
        Dim dvR As System.Data.DataView


        dvDC = frmH.dgvDataCompany.DataSource
        dvDW = frmH.dgvDataWatson.DataSource
        dgv = frmH.dgvMethodValData
        dgv2 = frmH.dgvMethValExistingGuWu
        dvMVD = frmH.dgvMethValExistingGuWu.DataSource
        int3 = dvMVD.Count

        dv = frmH.dgvMethodValData.DataSource

        'if the study is sample analysis, then don't replace
        dvR = frmH.dgvReports.DataSource
        boolMVal = False
        If dvR.Count = 0 Then
        Else
            str1 = NZ(dvR(0).Item("CHARREPORTTYPE"), "Sample Analysis")
            If Len(str1) = 0 Then
            Else
                If StrComp(str1, "Sample Analysis", CompareMethod.Text) = 0 Then
                    boolMVal = True
                End If
            End If
        End If

        If boolMVal Then
        Else
            For Count1 = 1 To 4
                Select Case Count1
                    Case 1 'Anticoagulant
                        strRet = NZ(frmH.cbxAnticoagulant.Text, "")
                        strFind = "Anticoagulant/Preservative"
                    Case 2 'Sample Size
                        str1 = "Sample Size"
                        int1 = FindRowDV(str1, dvDW)
                        strRet = dvDW(int1).Item("Value")
                        strFind = "Sample Size"
                    Case 3 'Species
                        str1 = "Species"
                        int1 = FindRowDV(str1, dvDW)
                        strRet = dvDW(int1).Item("Value")
                        strFind = "Species"
                    Case 4 'Matrix
                        str1 = "Matrix"
                        int1 = FindRowDV(str1, dvDW)
                        strRet = dvDW(int1).Item("Value")
                        strFind = "Matrix"
                    Case 5 'Corporate Study/Project Number
                        'str1 = "Validation Corporate Study/Project Number"
                        'int1 = FindRowDV(str1, dvDC)

                        'If int1 = -1 Then
                        '    strRet = "[None]"
                        'Else
                        '    strRet = dvDC(int1).Item("Value")
                        'End If

                        'strFind = "Validation Corporate Study/Project Number"
                    Case 6 'Validation Protocol Number
                        'str1 = "Validation Protocol Number"
                        'int1 = FindRowDV(str1, dvDC)


                        'If int1 = -1 Then
                        '    strRet = "[None]"
                        'Else
                        '    strRet = dvDC(int1).Item("Value")
                        'End If
                        ''strRet = dvDC(int1).Item("Value")
                        'strFind = "Validation Protocol Number"

                End Select

                int2 = FindRowDV(strFind, dv)
                For Count2 = 0 To int3 - 1
                    'strCol = dg2.Item(Count2, 0)
                    strCol = dgv2(0, Count2).Value
                    dv(int2).BeginEdit()
                    'dv(int2).Item(strCol) = strRet
                    Try
                        dv(int2).Item(strCol) = strRet
                    Catch ex As Exception
                        'MsgBox(ex.Message)
                    End Try
                    dv(int2).EndEdit()

                    'var1 = dv(int2).Item(strCol)
                    'If Len(NZ(var1, "")) = 0 Then
                    '    dv(int2).BeginEdit()
                    '    dv(int2).Item(strCol) = strRet
                    '    dv(int2).EndEdit()
                    'End If
                Next
            Next
        End If


    End Sub


    Sub cbxStudyCorrect()
        Dim str1 As String
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim int1 As Short
        Dim int2 As Short

        If boolFormLoad Then
            Exit Sub
        End If

        If frmH.dgvwStudy.Rows.Count = 0 Then
            GoTo end2
        ElseIf frmH.dgvwStudy.CurrentRow Is Nothing Then
            GoTo end2
        End If

        str1 = NZ(frmH.cbxStudy.Text, "")
        int1 = frmH.dgvwStudy.CurrentRow.Index
        If Len(str1) = 0 Then
            frmH.cbxStudy.SelectedIndex = int1
            GoTo end1
        End If

        strF = "STUDYNAME = '" & str1 & "'"
        tbl = tblwSTUDY
        rows = tbl.Select(strF)
        int2 = rows.Length
        If int2 = 0 Then
            frmH.cbxStudy.SelectedIndex = int1
            GoTo end1
        End If


end1:

        'MsgBox("cbxStudyCorrect")'for debugging

        'select entire text
        boolQuickFind = True
        boolQFDone = True
        'frmh.cbxStudy.Focus()

        If frmH.cbxStudy.Focused Then
            'SendKeys.Send("{HOME}")
            'SendKeys.Send("+{END}")
            frmH.cbxStudy.SelectAll()
        End If


end2:

    End Sub

    Sub QALIQUOT(ByVal strG As String)

        If tblAssignedSamples.Columns.Contains("ALIQUOTFACTOR") Then
            MsgBox("Yes: " & strG)
        Else
            MsgBox("No: " & strG)
        End If

    End Sub


    Sub AddCols_tblAss()

        Dim tbl1 As System.Data.DataTable
        Dim tbl2 As System.Data.DataTable
        Dim int1 As Short
        Dim int2 As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim boolFormat As Boolean
        Dim boolAdd As Boolean
        Dim boolFInt As Boolean
        Dim varF, VAR1

        tbl1 = tblAssignedSamples
        int1 = tbl1.Columns.Count
        tbl2 = tblAnalysisResultsHome
        int2 = tbl2.Columns.Count

        'for initialization
        'Dim dgv2 As DataGridView
        'int1 = dgv.Columns.Count

        If int2 = 0 Then
            GoTo end1
        End If

        'add unbound columns to tbl
        For Count1 = 0 To int2 - 1
            str1 = tbl2.Columns.Item(Count1).ColumnName
            'str2 = tbl2.Columns.item(Count1).HeaderText
            varF = tbl2.Columns(Count1).DataType.ToString

            If StrComp(str1, "ASSAYDATETIME", CompareMethod.Text) = 0 Then
                VAR1 = 0
            End If

            boolAdd = True
            boolFormat = False
            boolFInt = False
            Select Case str1
                Case "RUNID"
                    boolAdd = False
                Case "ANALYTEINDEX"
                    boolAdd = False
                Case "MASTERASSAYID"
                    boolAdd = False
                Case "RUNSAMPLESEQUENCENUMBER"
                    boolAdd = False
                Case "CONCENTRATION"
                    boolFormat = True
                Case "ALIQUOTFACTOR"
                    boolFInt = True
            End Select
            If boolAdd Then
                Dim nc As New DataColumn
                If tbl1.Columns.Contains(str1) Then
                Else
                    nc.ColumnName = str1
                    nc.Caption = str2
                    nc.DataType = System.Type.GetType(varF)
                    'If boolFormat Then
                    '    nc.DataType = System.Type.GetType("System.Double")
                    'End If
                    'If boolFInt Then
                    '    nc.DataType = System.Type.GetType("System.Single")
                    'End If
                    tbl1.Columns.Add(nc)
                End If
            End If
        Next

        'add stuff to tbl1
        If tbl1.Columns.Contains("BOOLEXCLSAMPLECHK") Then
        Else
            'now add a checkbox column
            Dim nchk As New DataColumn
            nchk.ColumnName = "BOOLEXCLSAMPLECHK"
            nchk.Caption = "Excl. Sample"
            nchk.DataType = System.Type.GetType("System.Boolean")
            nchk.AllowDBNull = True
            tbl1.Columns.Add(nchk)
        End If


        VAR1 = VAR1

end1:


    End Sub

    Sub RemoveBOOLEXCLSAMPLECHK()

        If tblAssignedSamples.Columns.Contains("BOOLEXCLSAMPLECHK") Then
            tblAssignedSamples.Columns.Remove("BOOLEXCLSAMPLECHK")
        End If

    End Sub

    Sub SetStatementTitle()

        'Exit Sub

        Dim dgv As DataGridView

        dgv = frmH.dgvReportStatementWord

        Try
            If frmH.rbSections.Checked Then
                dgv.Columns("CHARTITLE").HeaderText = "Statements"
            Else
                dgv.Columns("CHARTITLE").HeaderText = "Available Word Report Templates"
            End If

        Catch ex As Exception
            Dim var1

        End Try

    End Sub

    Sub UpdateTablePropBools()

        '??? This isn't needed
        Exit Sub

        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim str1 As String
        Dim intRows As Short
        Dim intT As Int64
        Dim idTR As Int64
        Dim idCRT As Int64
        Dim Count1 As Short
        Dim row() As DataRow
        Dim var1
        Dim tbl As System.Data.DataTable

        dgv = frmH.dgvReportTableConfiguration
        dv = dgv.DataSource
        str1 = "BOOLINCLUDE = TRUE"
        'dv.RowFilter = str1
        intRows = dv.Count
        ''wdd.visible = True

        'intRows = 4 'for testing
        For Count1 = 0 To intRows - 1

            var1 = dv(Count1).Item("BOOLINCLUDE")
            idTR = dv(Count1).Item("ID_TBLREPORTTABLE")
            idCRT = dv(Count1).Item("ID_TBLCONFIGREPORTTABLES")

            Try
                Call SetTablePropertiesBool(idTR, idCRT)

            Catch ex As Exception

            End Try

        Next

    End Sub

    Sub AssessQCs()

        '20160714 LEE:
        'This function can take a LONG time if study type is method validation and # of analytes is large (e.g. 4)
        'Actually, this function is not needed for method validation
        'For now, will enter a hard number cutoff:
        '   - if # QC levels > than say, 20, then ignore

        Dim dgv1 As DataGridView
        Dim dv1 As System.Data.DataView
        Dim tbl1 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim intCt1 As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim Count4 As Short
        Dim Count5 As Short
        Dim Count6 As Short
        Dim boolQC As Boolean
        Dim strF As String
        Dim strS As String
        Dim tbl2 As System.Data.DataTable
        Dim rows2() As DataRow
        Dim rows3() As DataRow
        Dim int30 As Short
        Dim int40 As Short
        Dim int50 As Short
        Dim int60 As Short
        Dim ctQCs As Short
        Dim var1, var2, var3, var4, var5, var6
        Dim int1 As Short
        Dim int2 As Short
        Dim numMean, numSD
        Dim numRepDilnQC
        Dim numRepQC
        Dim strAnal As String
        Dim dv3 As System.Data.DataView
        Dim boolDoAssigned As Boolean = False
        Dim tblBC As System.Data.DataTable
        Dim tblQC As System.Data.DataTable
        Dim strSQL As String
        Dim drows() As DataRow
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim nomConc As Single
        Dim intLevels As Short
        Dim intLevel As Short
        Dim boolPSAE As Boolean = False
        Dim tblP As System.Data.DataTable
        Dim strFP As String
        Dim rowsP() As DataRow
        Dim tblQCLevels As System.Data.DataTable
        Dim dv2 As System.Data.DataView
        Dim arrConc(100)
        Dim num1 As Single
        Dim num2 As Single
        Dim num3 As Single
        Dim num4 As Single

        Dim arrAcc(100) As Double
        Dim arrPrec(100) As Double

        Dim idCR As Int16
        Dim boolMethVal As Boolean
        strF = "ID_TBLSTUDIES = " & id_tblStudies ' & " AND CHARREPORTTYPE = 'Sample Analysis'"
        Dim dr() As DataRow = tblReports.Select(strF)
        If dr.Length = 0 Then
            str1 = "Sample Analysis"
        Else
            str1 = NZ(dr(0).Item("CHARREPORTTYPE"), "Sample Analysis")
        End If

        If InStr(1, str1, "Method", CompareMethod.Text) > 0 Then
            boolMethVal = True
            idCR = 11
        Else
            boolMethVal = False
            idCR = 4
        End If

        Dim intEnd As Short
        Dim strFFF As String

        Dim boolExPSAE As Boolean
        If frmH.chkPSAE.Checked Then
            boolExPSAE = True
        End If

        Dim dvT1 As System.Data.DataView
        If boolExPSAE Then
            dvT1 = New DataView(tblRegConAll)
        Else
            dvT1 = New DataView(tblRegCon)
        End If
        Dim tblRID As DataTable = dvT1.ToTable

        Try

            'determine if QCs are built from assigned samples
            tblP = tblTableProperties
            strFP = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLCONFIGREPORTTABLES = " & idCR
            rowsP = tblP.Select(strFP)
            int1 = rowsP(0).Item("BOOLINCLUDEPSAE")
            If int1 = 0 Then
                boolPSAE = False
            Else
                boolPSAE = True
            End If

            tbl1 = tblReportTable
            'strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLCONFIGREPORTTABLES = 4 AND BOOLREQUIRESSAMPLEASSIGNMENT = -1"
            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLCONFIGREPORTTABLES = " & idCR & " AND BOOLREQUIRESSAMPLEASSIGNMENT = -1 AND BOOLINCLUDE = -1"
            rows1 = tbl1.Select(strF)
            intCt1 = rows1.Length
            If intCt1 = 0 Then 'ignore
                boolDoAssigned = False
                tbl2 = tblBCQCConcs
            Else
                boolDoAssigned = True
                tbl2 = tblAssignedSamples
                'Exit Sub
            End If

            'tbl2 = tblAssignedSamples
            'tblQC = tblBCQCConcs

            'Dim arrBCStdActual(intCt1)

            For Count1 = 1 To ctAnalytes

                ReDim arrAcc(100)
                ReDim arrPrec(100)

                strFFF = GetARSRuns(tblRID, arrAnalytes(2, Count1), arrAnalytes(16, Count1), False)


                strAnal = tblAnalytesHome.Rows.Item(Count1 - 1).Item("AnalyteDescription")

                'find ctQCs = number of QC levels
                Dim blQCLevels As System.Data.DataTable

                If boolDoAssigned Then
                    strF = "ID_TBLCONFIGREPORTTABLES = " & idCR & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & CleanText(strAnal) & "' AND BOOLINTSTD = 0"
                    'strF = strF & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND (ELIMINATEDFLAG <> 'Y' AND BOOLEXCLSAMPLE <> -1)"
                    strF = strF & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND (ELIMINATEDFLAG <> 'Y' AND BOOLEXCLSAMPLE <> -1)"

                    strS = "ID_TBLSTUDIES ASC"
                    dv2 = New DataView(tbl2, strF, strS, DataViewRowState.CurrentRows)
                    tblQCLevels = dv2.ToTable("a", True, "CHARHELPER1", "NOMCONC")
                    ctQCs = tblQCLevels.Rows.Count
                    intLevels = tblQCLevels.Rows.Count
                    intEnd = ctQCs
                Else

                    'str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID,
                    ' " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG,
                    ' " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID,
                    ' ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION,
                    ' ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYANALYTES.ANALYTEID "

                    'If boolPSAE Then
                    '    strF = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ELIMINATEDFLAG <> 'Y' AND RUNTYPEID > 0"
                    'Else
                    '    strF = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ELIMINATEDFLAG <> 'Y' AND RUNTYPEID <> 3"
                    'End If
                    If boolPSAE Then
                        strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ELIMINATEDFLAG <> 'Y' AND RUNTYPEID > 0"
                    Else
                        strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ELIMINATEDFLAG <> 'Y' AND RUNTYPEID <> 3"
                    End If

                    '20170928 LEE: need to add matrix for multiple matrix studies

                    If Len(strFFF) = 0 Then
                        If boolPSAE Then
                            strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ELIMINATEDFLAG <> 'Y' AND RUNTYPEID > 0 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "'"
                        Else
                            strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ELIMINATEDFLAG <> 'Y' AND RUNTYPEID <> 3 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "'"
                        End If
                    Else
                        If boolPSAE Then
                            strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ELIMINATEDFLAG <> 'Y' AND RUNTYPEID > 0 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "' AND (" & strFFF & ")"
                        Else
                            strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ELIMINATEDFLAG <> 'Y' AND RUNTYPEID <> 3 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "' AND (" & strFFF & ")"
                        End If
                    End If

                    strS = "ASSAYLEVEL ASC"
                    Try
                        dv2 = New DataView(tbl2, strF, strS, DataViewRowState.CurrentRows)
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    tblQCLevels = dv2.ToTable("a", True, "ASSAYLEVEL")
                    intLevels = tblQCLevels.Rows.Count


                    'str1 = "SELECT DISTINCT ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, 
                    'ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, 
                    'ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS,  ASSAYREPS.FLAGPERCENT "
                    'PSAE Stuff not needed here
                    'If boolPSAE Then
                    '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID > 0"
                    'Else
                    '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID <> 3"
                    'End If
                    'str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1)
                    str1 = "ANALYTEID = " & arrAnalytes(2, Count1)
                    'must account for SAMPLETYPEID (matrix) for multiple matrix studies
                    str1 = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "'"
                    strS = "LEVELNUMBER"
                    drows = tblBCQCs.Select(str1, strS)
                    int1 = drows.Length
                    ctQCs = int1

                    'use smallest number
                    intEnd = ctQCs
                    If ctQCs <= intLevels Then
                        intEnd = ctQCs
                    Else
                        intEnd = intLevels
                    End If
                    ctQCs = intEnd

                End If

                For Count2 = 0 To intEnd - 1
                    If boolDoAssigned Then
                        nomConc = tblQCLevels.Rows.Item(Count2).Item("NOMCONC")
                        strF = "ID_TBLCONFIGREPORTTABLES = " & idCR & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & CleanText(strAnal) & "' AND BOOLINTSTD = 0 AND NOMCONC = " & nomConc & " AND CHARHELPER2 IS NULL AND (ELIMINATEDFLAG = 'N' AND BOOLEXCLSAMPLE = 0)"
                    Else
                        nomConc = NZ(drows(Count2).Item("CONCENTRATION"), 0)
                        intLevel = tblQCLevels.Rows(Count2).Item("ASSAYLEVEL")
                        'If boolPSAE Then
                        '    strF = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ASSAYLEVEL = " & intLevel & " AND RUNTYPEID > 0 AND ELIMINATEDFLAG = 'N' AND RUNANALYTEREGRESSIONSTATUS = 3"
                        'Else
                        '    strF = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ASSAYLEVEL = " & intLevel & " AND RUNTYPEID <> 3 AND ELIMINATEDFLAG = 'N' AND RUNANALYTEREGRESSIONSTATUS = 3"
                        'End If
                        If boolPSAE Then
                            strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ASSAYLEVEL = " & intLevel & " AND RUNTYPEID > 0 AND ELIMINATEDFLAG = 'N' AND RUNANALYTEREGRESSIONSTATUS = 3"
                        Else
                            strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ASSAYLEVEL = " & intLevel & " AND RUNTYPEID <> 3 AND ELIMINATEDFLAG = 'N' AND RUNANALYTEREGRESSIONSTATUS = 3"
                        End If

                        '20170928 LEE: need to add matrix for multiple matrix studies


                        '"' AND (" & strFFF & ")"
                        If Len(strFFF) = 0 Then
                            If boolPSAE Then
                                strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ASSAYLEVEL = " & intLevel & " AND RUNTYPEID > 0 AND ELIMINATEDFLAG = 'N' AND RUNANALYTEREGRESSIONSTATUS = 3 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "'"
                            Else
                                strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ASSAYLEVEL = " & intLevel & " AND RUNTYPEID <> 3 AND ELIMINATEDFLAG = 'N' AND RUNANALYTEREGRESSIONSTATUS = 3 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "'"
                            End If
                        Else
                            If boolPSAE Then
                                strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ASSAYLEVEL = " & intLevel & " AND RUNTYPEID > 0 AND ELIMINATEDFLAG = 'N' AND RUNANALYTEREGRESSIONSTATUS = 3 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "' AND (" & strFFF & ")"
                            Else
                                strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ASSAYLEVEL = " & intLevel & " AND RUNTYPEID <> 3 AND ELIMINATEDFLAG = 'N' AND RUNANALYTEREGRESSIONSTATUS = 3 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "' AND (" & strFFF & ")"
                            End If
                        End If

                    End If

                    Erase rows2
                    'rows2 = tbl2.Select(strF)
                    Try 'Error here
                        rows2 = tbl2.Select(strF)
                        int1 = rows2.Length
                    Catch ex As Exception
                        int1 = 0
                    End Try
                    'int1 = rows2.Length
                    If int1 = 0 Then
                        var3 = 0
                        var4 = 0
                    Else

                        numMean = SigFigOrDec(MeanDR(rows2, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False), LSigFig, False)
                        numSD = SigFigOrDec(StdDevDR(rows2, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False), LSigFig, False)

                        If CDec(numMean) = 0 Then
                            var3 = 0 'CDec(Format(RoundToDecimal(numSD / numMean * 100, 10), "#0.0")) 'for testing
                            var4 = 0 'CDec(Format(RoundToDecimal(((numMean / var1) - 1) * 100, 10), "#0.0")) 'for testing
                        Else
                            var3 = CDec(Format(RoundToDecimalRAFZ(RoundToDecimalRAFZ(numSD / numMean * 100, intQCDec + 1), intQCDec), strQCDec)) 'precision/CV
                            var4 = CDec(Format(RoundToDecimalRAFZ(RoundToDecimalRAFZ(((numMean / nomConc) - 1) * 100, intQCDec + 1), intQCDec), strQCDec)) 'accuracy/bias
                        End If

                    End If
                    arrPrec(Count2 + 1) = var3
                    arrAcc(Count2 + 1) = var4
                    var1 = var1
                Next

                'legend:
                int30 = FindRow("QC Mean Accuracy Min", tblWatsonAnalRefTable, "Item")
                int40 = FindRow("QC Mean Accuracy Max", tblWatsonAnalRefTable, "Item")
                int50 = FindRow("QC Precision Min", tblWatsonAnalRefTable, "Item")
                int60 = FindRow("QC Precision Max", tblWatsonAnalRefTable, "Item")
                If ctQCs = 0 Then
                    var1 = Format(0, strQCDec)
                    tblWatsonAnalRefTable.Rows.Item(int30).BeginEdit()
                    tblWatsonAnalRefTable.Rows.Item(int30).Item(Count1) = var1
                    tblWatsonAnalRefTable.Rows.Item(int30).EndEdit()
                    var1 = Format(0, strQCDec)
                    tblWatsonAnalRefTable.Rows.Item(int40).BeginEdit()
                    tblWatsonAnalRefTable.Rows.Item(int40).Item(Count1) = var1
                    tblWatsonAnalRefTable.Rows.Item(int40).EndEdit()
                    var1 = Format(0, strQCDec)
                    tblWatsonAnalRefTable.Rows.Item(int50).BeginEdit()
                    tblWatsonAnalRefTable.Rows.Item(int50).Item(Count1) = var1
                    tblWatsonAnalRefTable.Rows.Item(int50).EndEdit()
                    var1 = Format(0, strQCDec)
                    tblWatsonAnalRefTable.Rows.Item(int60).BeginEdit()
                    tblWatsonAnalRefTable.Rows.Item(int60).Item(Count1) = var1
                    tblWatsonAnalRefTable.Rows.Item(int60).EndEdit()
                Else
                    var1 = Format(CDec(GetMin(arrAcc, intEnd)), strQCDec)
                    tblWatsonAnalRefTable.Rows.Item(int30).BeginEdit()
                    tblWatsonAnalRefTable.Rows.Item(int30).Item(Count1) = var1
                    tblWatsonAnalRefTable.Rows.Item(int30).EndEdit()
                    var1 = Format(CDec(GetMax(arrAcc, intEnd)), strQCDec)
                    tblWatsonAnalRefTable.Rows.Item(int40).BeginEdit()
                    tblWatsonAnalRefTable.Rows.Item(int40).Item(Count1) = var1
                    tblWatsonAnalRefTable.Rows.Item(int40).EndEdit()
                    var1 = Format(CDec(GetMin(arrPrec, intEnd)), strQCDec)
                    tblWatsonAnalRefTable.Rows.Item(int50).BeginEdit()
                    tblWatsonAnalRefTable.Rows.Item(int50).Item(Count1) = var1
                    tblWatsonAnalRefTable.Rows.Item(int50).EndEdit()
                    var1 = Format(CDec(GetMax(arrPrec, intEnd)), strQCDec)
                    tblWatsonAnalRefTable.Rows.Item(int60).BeginEdit()
                    tblWatsonAnalRefTable.Rows.Item(int60).Item(Count1) = var1
                    tblWatsonAnalRefTable.Rows.Item(int60).EndEdit()
                End If

                ''Find QC and Diln QC Reps
                numRepDilnQC = 0
                numRepQC = 0
                If ctQCs = 0 Or intLevels = 0 Then
                Else
                    If boolDoAssigned Then
                        var1 = tblQCLevels.Rows.Item(0).Item("NOMCONC")
                        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & CleanText(strAnal) & "' AND BOOLINTSTD = 0 AND NOMCONC = " & var1
                        Erase rows2
                        rows2 = tbl2.Select(strF)
                        var2 = rows2(0).Item("RUNID")
                        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & CleanText(strAnal) & "' AND BOOLINTSTD = 0 AND NOMCONC = " & var1 & " AND RUNID = " & var2 & " AND ID_TBLCONFIGREPORTTABLES = " & idCR & ""
                        Erase rows3
                        rows3 = tbl2.Select(strF)
                        numRepQC = rows3.Length

                        'find diln QC Reps

                        'must get unique id_tblReportTable. Get from Dilution table
                        Try

                            numRepDilnQC = 0
                            For Count3 = 1 To 3
                                Select Case Count3
                                    Case 1 'Dilution table
                                        var5 = 12
                                    Case 2 'ANOVA table
                                        var5 = 11
                                    Case 3 'QC table
                                        var5 = 4
                                End Select

                                Select Case Count3
                                    Case 1 'dilution table
                                        strF = GetBOOLSTATSNRFilter(12) 'this will return a partial string
                                        'e.g.  strF = "ID_TBLREPORTTABLE = " & var1
                                        strF = strF & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & CleanText(strAnal) & "'"
                                    Case 2, 3 'ANOVA, QC tables
                                        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & CleanText(strAnal) & "' AND BOOLINTSTD = 0 AND ALIQUOTFACTOR <> 1  AND ID_TBLCONFIGREPORTTABLES = " & var5
                                End Select

                                Erase rows2
                                rows2 = tbl2.Select(strF)
                                If rows2.Length = 0 Then
                                Else

                                    Dim tblA As DataTable = rows2.CopyToDataTable
                                    'Dim tblB As DataTable = myDT.DefaultView.ToTable(True, "name")
                                    Dim tblB As DataTable = tblA.DefaultView.ToTable("a", True, "ID_TBLREPORTTABLE")

                                    For Count2 = 0 To tblB.Rows.Count - 1

                                        var6 = tblB.Rows(Count2).Item("ID_TBLREPORTTABLE")

                                        '20190228 LEE: Redo with GetBOOLSTATSNRFilter
                                        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND " & GetBOOLSTATSNRFilter(12) '12 = Dilution Table

                                        'check to see if report is used
                                        'strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & var6
                                        Dim rowsRT() As DataRow = tblReportTable.Select(strF)
                                        If rowsRT.Length = 0 Then
                                        Else
                                            var1 = NZ(rowsRT(0).Item("BOOLINCLUDE"), 0)
                                            If var1 = 0 Then
                                            Else

                                                'strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & strAnal & "' AND BOOLINTSTD = 0 AND CHARHELPER1 = 'QC Diln'"
                                                strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & CleanText(strAnal) & "' AND BOOLINTSTD = 0 AND ALIQUOTFACTOR <> 1  AND ID_TBLCONFIGREPORTTABLES = " & var5 & " AND ID_TBLREPORTTABLE = " & var6
                                                Erase rows2
                                                rows2 = tbl2.Select(strF)
                                                'var1 = rows2(0).Item("NOMCONC")
                                                If rows2.Length = 0 Then
                                                Else


                                                    For Count5 = 0 To rows2.Length - 1
                                                        var2 = rows2(Count5).Item("RUNID")
                                                        For Count4 = 0 To rows2.Length - 1
                                                            var1 = rows2(Count4).Item("NOMCONC")
                                                            'strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & strAnal & "' AND BOOLINTSTD = 0 AND CHARHELPER1 = 'QC Diln' AND NOMCONC = " & var1
                                                            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & CleanText(strAnal) & "' AND BOOLINTSTD = 0 AND ALIQUOTFACTOR <> 1 AND NOMCONC = " & var1 & " AND RUNID = " & var2 & " AND ID_TBLCONFIGREPORTTABLES = " & var5 & " AND ID_TBLREPORTTABLE = " & var6
                                                            Erase rows3
                                                            rows3 = tbl2.Select(strF)
                                                            var3 = rows3.Length
                                                            If var3 > numRepDilnQC Then
                                                                numRepDilnQC = var3
                                                            End If
                                                        Next Count4

                                                    Next Count5

                                                End If

                                            End If
                                        End If

                                    Next Count2

                                End If

                                var1 = var1 'debug

                            Next Count3

                        Catch ex As Exception
                            var1 = ex.Message
                        End Try

                    Else

                        Count5 = 0
                        Dim int201 As Short
                        Dim drowsF() As DataRow
                        Dim intF As Short
                        Dim maxRep As Short
                        Dim drowsR() As DataRow

                        'Dim arrBCQCs(5, 200) '1=LevelNumber, 2=Concentration, 3=ID, 4=#ofReplicates

                        numRepDilnQC = 0
                        numRepQC = 0
                        maxRep = 0
                        For Count2 = 1 To ctQCs
                            'var1 = arrBCQCs(1, Count2)
                            var1 = tblBCQCs.Rows(Count2 - 1).Item("LEVELNUMBER")
                            var2 = tblBCQCs.Rows(Count2 - 1).Item("CONCENTRATION")

                            'str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND LEVELNUMBER = " & var1 & " and MASTERASSAYID = " & arrAnalytes(12, Count1) & " and CONCENTRATION = " & var2
                            str1 = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND LEVELNUMBER = " & var1 & " AND CONCENTRATION = " & var2
                            ''''''debugwriteline(str1)
                            drowsR = tblQCReps.Select(str1, "REPLICATENUMBER ASC")
                            int1 = drowsR.Length
                            var4 = drowsR(int1 - 1).Item("REPLICATENUMBER")
                            If var4 > maxRep Then
                                maxRep = maxRep + 1
                            End If
                            ''
                        Next

                        'find number of regression parameters
                        Dim int10 As Short
                        'Dim dvT As system.data.dataview = New DataView(tblRegCon)
                        Dim dvT As System.Data.DataView = New DataView(tblRegConAll)
                        'If boolPSAE Then
                        '    str1 = "STUDYID = " & wStudyID & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID > 0 AND ANALYTEID = " & arrAnalytes(2, Count1)
                        'Else
                        '    str1 = "STUDYID = " & wStudyID & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID <> 3 AND ANALYTEID = " & arrAnalytes(2, Count1)
                        'End If
                        If boolPSAE Then
                            str1 = "STUDYID = " & wStudyID & " AND RUNTYPEID > 0 AND ANALYTEID = " & arrAnalytes(2, Count1)
                        Else
                            str1 = "STUDYID = " & wStudyID & " AND RUNTYPEID <> 3 AND ANALYTEID = " & arrAnalytes(2, Count1)
                        End If


                        '"' AND (" & strFFF & ")"
                        If Len(strFFF) = 0 Then
                            If boolPSAE Then
                                str1 = "STUDYID = " & wStudyID & " AND RUNTYPEID > 0 AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "'"
                            Else
                                str1 = "STUDYID = " & wStudyID & " AND RUNTYPEID <> 3 AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "'"
                            End If
                        Else

                            If boolPSAE Then
                                str1 = "STUDYID = " & wStudyID & " AND RUNTYPEID > 0 AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "' AND (" & strFFF & ")"
                            Else
                                str1 = "STUDYID = " & wStudyID & " AND RUNTYPEID <> 3 AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "' AND (" & strFFF & ")"
                            End If
                        End If

                        'str1 = "STUDYID = " & wStudyID & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1)
                        Try
                            dvT.RowFilter = str1
                        Catch ex As Exception
                            var1 = var1
                        End Try

                        int10 = dvT.Count
                        Dim tblT As System.Data.DataTable = dvT.ToTable("a", True, "REGRESSIONPARAMETERID")
                        Dim intRP As Short
                        Dim intTRows As Short
                        Dim var8
                        intRP = tblT.Rows.Count

                        'find number of table rows
                        Dim tblTT As System.Data.DataTable = dvT.ToTable("b", True, "RSQUARED")
                        intTRows = tblTT.Rows.Count

                        Try
                            For Count2 = 0 To int10 - 1 Step intRP 'step  because tblRegCon has multirow entries
                                'need maxRep rows for each accepted run
                                int201 = CInt(dvT(Count2).Item("RUNID"))
                                For Count3 = 0 To maxRep - 1
                                    'establish array going across table ctQC number of times
                                    For Count4 = 1 To ctQCs
                                        'var2 = arrBCQCs(1, Count4) '.Item("LevelNumber")
                                        'var3 = arrBCQCs(2, Count4) 'CONCENTRATION
                                        var2 = tblBCQCs.Rows(Count4 - 1).Item("LEVELNUMBER")
                                        var3 = tblBCQCs.Rows(Count4 - 1).Item("CONCENTRATION")

                                        Count5 = Count5 + 1
                                        'If boolPSAE Then
                                        '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNID = " & int201 & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID > 0"
                                        'Else
                                        '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNID = " & int201 & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3"
                                        'End If
                                        If boolPSAE Then
                                            str1 = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND RUNID = " & int201 & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID > 0"
                                        Else
                                            str1 = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND RUNID = " & int201 & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3"
                                        End If


                                        '"' AND (" & strFFF & ")"
                                        If Len(strFFF) = 0 Then

                                            If boolPSAE Then
                                                str1 = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND RUNID = " & int201 & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID > 0 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "'"
                                            Else
                                                str1 = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND RUNID = " & int201 & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "'"
                                            End If
                                        Else

                                            If boolPSAE Then
                                                str1 = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND RUNID = " & int201 & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID > 0 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "' AND (" & strFFF & ")"
                                            Else
                                                str1 = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND RUNID = " & int201 & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "' AND (" & strFFF & ")"
                                            End If
                                        End If

                                        ''''''''''''''''''''''''''''''''console.writeline(id_tblStudies & ": " & str1)
                                        Erase drowsF
                                        drowsF = tblBCQCConcs.Select(str1, "RUNSAMPLESEQUENCENUMBER ASC")
                                        intF = drowsF.Length
                                        If intF < Count3 + 1 Then 'enter base values
                                        Else 'enter in total set of values
                                            ''evaluate QC Rep #
                                            var8 = drowsF(Count3).Item("AliquotFactor")
                                            If var8 = 1 Then
                                                numRepQC = Count3 + 1
                                            Else
                                                numRepDilnQC = Count3 + 1
                                            End If

                                        End If
                                    Next
                                Next
                            Next
                        Catch ex As Exception
                            var1 = ex.Message
                            var1 = var1
                        End Try


                        'record #ofReplicates in tblWatsonAnalRef: QCConcentrationCount
                        int1 = FindRow("# of QC Replicates", tblWatsonAnalRefTable, "Item")
                        tblWatsonAnalRefTable.Rows.Item(int1).BeginEdit()
                        tblWatsonAnalRefTable.Rows.Item(int1).Item(Count1) = numRepQC
                        tblWatsonAnalRefTable.Rows.Item(int1).EndEdit()
                        int1 = FindRow("# of Dilution QC Replicates", tblWatsonAnalRefTable, "Item")
                        tblWatsonAnalRefTable.Rows.Item(int1).BeginEdit()
                        tblWatsonAnalRefTable.Rows.Item(int1).Item(Count1) = numRepDilnQC
                        tblWatsonAnalRefTable.Rows.Item(int1).EndEdit()

                    End If
                End If
                'record items in tblWatsonAnalRef
                int1 = FindRow("# of QC Levels", tblWatsonAnalRefTable, "Item")
                tblWatsonAnalRefTable.Rows.Item(int1).BeginEdit()
                tblWatsonAnalRefTable.Rows.Item(int1).Item(Count1) = ctQCs
                tblWatsonAnalRefTable.Rows.Item(int1).EndEdit()

                'record #ofReplicates in tblWatsonAnalRef: QCConcentrationCount
                int1 = FindRow("# of QC Replicates", tblWatsonAnalRefTable, "Item")
                tblWatsonAnalRefTable.Rows.Item(int1).BeginEdit()
                tblWatsonAnalRefTable.Rows.Item(int1).Item(Count1) = numRepQC
                tblWatsonAnalRefTable.Rows.Item(int1).EndEdit()
                int1 = FindRow("# of Dilution QC Replicates", tblWatsonAnalRefTable, "Item")
                tblWatsonAnalRefTable.Rows.Item(int1).BeginEdit()
                tblWatsonAnalRefTable.Rows.Item(int1).Item(Count1) = numRepDilnQC
                tblWatsonAnalRefTable.Rows.Item(int1).EndEdit()


            Next

            var1 = var1


        Catch ex As Exception

            ctQCs = 0
            'legend:
            int30 = FindRow("QC Mean Accuracy Min", tblWatsonAnalRefTable, "Item")
            int40 = FindRow("QC Mean Accuracy Max", tblWatsonAnalRefTable, "Item")
            int50 = FindRow("QC Precision Min", tblWatsonAnalRefTable, "Item")
            int60 = FindRow("QC Precision Max", tblWatsonAnalRefTable, "Item")
            If Count1 = 0 Then
                Count1 = 1
            End If
            If ctQCs = 0 Then
                var1 = Format(0, strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int30).BeginEdit()
                tblWatsonAnalRefTable.Rows.Item(int30).Item(Count1) = var1
                tblWatsonAnalRefTable.Rows.Item(int30).EndEdit()
                var1 = Format(0, strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int40).BeginEdit()
                tblWatsonAnalRefTable.Rows.Item(int40).Item(Count1) = var1
                tblWatsonAnalRefTable.Rows.Item(int40).EndEdit()
                var1 = Format(0, strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int50).BeginEdit()
                tblWatsonAnalRefTable.Rows.Item(int50).Item(Count1) = var1
                tblWatsonAnalRefTable.Rows.Item(int50).EndEdit()
                var1 = Format(0, strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int60).BeginEdit()
                tblWatsonAnalRefTable.Rows.Item(int60).Item(Count1) = var1
                tblWatsonAnalRefTable.Rows.Item(int60).EndEdit()
            Else
                var1 = Format(CDec(GetMin(arrAcc, ctQCs)), strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int30).BeginEdit()
                tblWatsonAnalRefTable.Rows.Item(int30).Item(Count1) = var1
                tblWatsonAnalRefTable.Rows.Item(int30).EndEdit()
                var1 = Format(CDec(GetMax(arrAcc, ctQCs)), strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int40).BeginEdit()
                tblWatsonAnalRefTable.Rows.Item(int40).Item(Count1) = var1
                tblWatsonAnalRefTable.Rows.Item(int40).EndEdit()
                var1 = Format(CDec(GetMin(arrPrec, ctQCs)), strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int50).BeginEdit()
                tblWatsonAnalRefTable.Rows.Item(int50).Item(Count1) = var1
                tblWatsonAnalRefTable.Rows.Item(int50).EndEdit()
                var1 = Format(CDec(GetMax(arrPrec, ctQCs)), strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int60).BeginEdit()
                tblWatsonAnalRefTable.Rows.Item(int60).Item(Count1) = var1
                tblWatsonAnalRefTable.Rows.Item(int60).EndEdit()
            End If

            'record items in tblWatsonAnalRef
            int1 = FindRow("# of QC Levels", tblWatsonAnalRefTable, "Item")
            tblWatsonAnalRefTable.Rows.Item(int1).Item(Count1) = ctQCs

            'record #ofReplicates in tblWatsonAnalRef: QCConcentrationCount
            int1 = FindRow("# of QC Replicates", tblWatsonAnalRefTable, "Item")
            tblWatsonAnalRefTable.Rows.Item(int1).BeginEdit()
            tblWatsonAnalRefTable.Rows.Item(int1).Item(Count1) = 0 'numRepQC
            tblWatsonAnalRefTable.Rows.Item(int1).EndEdit()
            int1 = FindRow("# of Dilution QC Replicates", tblWatsonAnalRefTable, "Item")
            tblWatsonAnalRefTable.Rows.Item(int1).BeginEdit()
            tblWatsonAnalRefTable.Rows.Item(int1).Item(Count1) = 0 'numRepDilnQC
            tblWatsonAnalRefTable.Rows.Item(int1).EndEdit()

        End Try


        'now do Calibr Stds

        'tblBCStds
        'str1 = "SELECT DISTINCT ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, 
        'ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID "

        'tblBCStdConcs
        'str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, 
        '" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG,
        ' ANARUNANALYTERESULTS.ANALYTEINDEX, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR,
        ' ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.CONCENTRATION, 
        'ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS("")

        Try

            'determine if Calibr Stds are built from assigned samples
            tblP = tblTableProperties
            strFP = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLCONFIGREPORTTABLES = 3"
            rowsP = tblP.Select(strFP)
            var1 = rowsP(0).Item("BOOLINCLUDEPSAE")
            If int1 = 0 Then
                boolPSAE = False
            Else
                boolPSAE = True
            End If

            'don't use boolIncludePSAE from tblTableProperties
            'use tblTableProperties
            boolPSAE = boolIncludePSAE


            tbl1 = tblReportTable
            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLCONFIGREPORTTABLES = 3 AND BOOLREQUIRESSAMPLEASSIGNMENT = -1 AND BOOLINCLUDE = -1"
            rows1 = tbl1.Select(strF)
            intCt1 = rows1.Length
            If intCt1 = 0 Then 'ignore
                boolDoAssigned = False
                tbl2 = tblBCStdConcs
            Else
                boolDoAssigned = True
                tbl2 = tblAssignedSamples
                'Exit Sub
            End If

            'tbl2 = tblAssignedSamples
            'tblQC = tblBCQCConcs

            'Dim arrBCStdActual(intCt1)

            For Count1 = 1 To ctAnalytes
                ReDim arrAcc(100)
                ReDim arrPrec(100)

                strAnal = tblAnalytesHome.Rows.Item(Count1 - 1).Item("AnalyteDescription")

                strFFF = GetARSRuns(tblRID, arrAnalytes(2, Count1), arrAnalytes(16, Count1), False)

                'find ctQCs = number of QC levels
                If boolDoAssigned Then
                    strF = "ID_TBLCONFIGREPORTTABLES = 3 AND ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & CleanText(strAnal) & "' AND BOOLINTSTD = 0"
                    'strF = strF & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND (ELIMINATEDFLAG <> 'Y' AND BOOLEXCLSAMPLE <> -1)"
                    strF = strF & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND (ELIMINATEDFLAG <> 'Y' AND BOOLEXCLSAMPLE <> -1)"
                    strS = "ID_TBLSTUDIES ASC"
                    dv2 = New DataView(tbl2, strF, strS, DataViewRowState.CurrentRows)
                    tblQCLevels = dv2.ToTable("a", True, "NOMCONC")
                    ctQCs = tblQCLevels.Rows.Count
                Else

                    'str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID,
                    ' " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG,
                    ' " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID,
                    ' ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION,
                    ' ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYANALYTES.ANALYTEID "

                    'If boolPSAE Then
                    '    strF = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND ELIMINATEDFLAG <> 'Y' AND RUNTYPEID > 0"
                    'Else
                    '    strF = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND ELIMINATEDFLAG <> 'Y' AND RUNTYPEID <> 3"
                    'End If
                    If boolPSAE Then
                        strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ELIMINATEDFLAG <> 'Y' AND RUNTYPEID > 0"
                    Else
                        strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ELIMINATEDFLAG <> 'Y' AND RUNTYPEID <> 3"
                    End If

                  

                    '"' AND (" & strFFF & ")"
                    If Len(strFFF) = 0 Then
                        If boolPSAE Then
                            strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ELIMINATEDFLAG <> 'Y' AND RUNTYPEID > 0 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "'"
                        Else
                            strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ELIMINATEDFLAG <> 'Y' AND RUNTYPEID <> 3 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "'"
                        End If
                    Else
                        If boolPSAE Then
                            strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ELIMINATEDFLAG <> 'Y' AND RUNTYPEID > 0 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "' AND (" & strFFF & ")"
                        Else
                            strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ELIMINATEDFLAG <> 'Y' AND RUNTYPEID <> 3 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "' AND (" & strFFF & ")"
                        End If
                    End If

                    strS = "ASSAYLEVEL ASC"
                    Try
                        dv2 = New DataView(tbl2, strF, strS, DataViewRowState.CurrentRows)
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    tblQCLevels = dv2.ToTable("a", True, "ASSAYLEVEL")
                    intLevels = tblQCLevels.Rows.Count


                    'str1 = "SELECT DISTINCT ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, 
                    'ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, 
                    'ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS,  ASSAYREPS.FLAGPERCENT "

                    'If boolPSAE Then
                    '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ANALYTEID = " & arrAnalytes(2, Count1) '  ' & " AND RUNTYPEID > 0"
                    'Else
                    '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ANALYTEID = " & arrAnalytes(2, Count1) ' & " AND RUNTYPEID <> 3"
                    'End If
                    If boolPSAE Then
                        str1 = "ANALYTEID = " & arrAnalytes(2, Count1) '  ' & " AND RUNTYPEID > 0"
                    Else
                        str1 = "ANALYTEID = " & arrAnalytes(2, Count1) ' & " AND RUNTYPEID <> 3"
                    End If
                    str1 = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "'"
                    'this doesn't need strFFF

                    strS = "LEVELNUMBER"
                    Try
                        drows = tblBCStds.Select(str1, strS)
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    int1 = drows.Length
                    ctQCs = int1
                End If


                Dim intCt As Short
                intCt = 0
                For Count2 = 0 To ctQCs - 1
                    If Count2 > intLevels - 1 Then
                        'Exit For
                    End If
                    '''''''''''console.writeline("Start " & Count2)
                    If boolDoAssigned Then
                        nomConc = tblQCLevels.Rows.Item(Count2).Item("NOMCONC")
                        strF = "ID_TBLCONFIGREPORTTABLES = 3 AND ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & CleanText(strAnal) & "' AND BOOLINTSTD = 0 AND NOMCONC = " & nomConc & " AND CHARHELPER2 IS NULL AND (ELIMINATEDFLAG = 'N' AND BOOLEXCLSAMPLE = 0)"
                    Else
                        nomConc = NZ(drows(Count2).Item("CONCENTRATION"), 0)
                        'intLevel = tblQCLevels.Rows(Count2).Item("ASSAYLEVEL")
                        intLevel = drows(Count2).Item("LEVELNUMBER")
                        'If boolPSAE Then
                        '    strF = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND ASSAYLEVEL = " & intLevel & " AND RUNTYPEID > 0 AND ELIMINATEDFLAG = 'N' AND RUNANALYTEREGRESSIONSTATUS = 3"
                        'Else
                        '    strF = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND ASSAYLEVEL = " & intLevel & " AND RUNTYPEID <> 3 AND ELIMINATEDFLAG = 'N' AND RUNANALYTEREGRESSIONSTATUS = 3"
                        'End If
                        If boolPSAE Then
                            strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ASSAYLEVEL = " & intLevel & " AND RUNTYPEID > 0 AND ELIMINATEDFLAG = 'N' AND RUNANALYTEREGRESSIONSTATUS = 3"
                        Else
                            strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ASSAYLEVEL = " & intLevel & " AND RUNTYPEID <> 3 AND ELIMINATEDFLAG = 'N' AND RUNANALYTEREGRESSIONSTATUS = 3"
                        End If

                     

                        '"' AND (" & strFFF & ")"
                        If Len(strFFF) = 0 Then
                            If boolPSAE Then
                                strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ASSAYLEVEL = " & intLevel & " AND RUNTYPEID > 0 AND ELIMINATEDFLAG = 'N' AND RUNANALYTEREGRESSIONSTATUS = 3 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "'"
                            Else
                                strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ASSAYLEVEL = " & intLevel & " AND RUNTYPEID <> 3 AND ELIMINATEDFLAG = 'N' AND RUNANALYTEREGRESSIONSTATUS = 3 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "'"
                            End If
                        Else
                            If boolPSAE Then
                                strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ASSAYLEVEL = " & intLevel & " AND RUNTYPEID > 0 AND ELIMINATEDFLAG = 'N' AND RUNANALYTEREGRESSIONSTATUS = 3 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "' AND (" & strFFF & ")"
                            Else
                                strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND ASSAYLEVEL = " & intLevel & " AND RUNTYPEID <> 3 AND ELIMINATEDFLAG = 'N' AND RUNANALYTEREGRESSIONSTATUS = 3 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "' AND (" & strFFF & ")"
                            End If
                        End If

                    End If

                    Erase rows2
                    'rows2 = tbl2.Select(strF)
                    Try 'Error here

                        rows2 = tbl2.Select(strF)
                        int1 = rows2.Length
                    Catch ex As Exception
                        int1 = 0
                    End Try
                    'int1 = rows2.Length
                    If int1 = 0 Then
                        var3 = "NA"
                        var4 = "NA"
                    Else

                        'For Count3 = 0 To int1 - 1
                        '    var3 = rows2(Count3).Item("RUNID")
                        '    var4 = var3
                        'Next
                        'numMean = SigFigOrDec(Mean(int1, arrBCStdActual), LSigFig, True)
                        'numSD = SigFigOrDec(StdDev(int1, arrBCStdActual), LSigFig, True)

                        'var3 = rows2(0).Item("RUNTYPEID")


                        'fill arrconc
                        ReDim arrConc(int1)
                        ''''''''''''console.writeline("StartCalibr")
                        'For Count3 = 0 To int1 - 1
                        '    num1 = rows2(Count3).Item("CONCENTRATION")
                        '    num2 = rows2(Count3).Item("ALIQUOTFACTOR")
                        '    num3 = num1 / num2
                        '    num4 = SigFigOrDec(num3, LSigFig, False)
                        '    '''''''''''console.writeline(num4)
                        '    arrConc(Count3 + 1) = num4
                        'Next
                        ''''''''''''console.writeline("EndCalibr")

                        'numMean = SigFigOrDec(Mean(int1, arrConc), LSigFig, False)
                        'numSD = SigFigOrDec(StdDev(int1, arrConc), LSigFig, False)

                        If Count2 = 6 Then 'debugging
                            var1 = 0
                        End If
                        numMean = SigFigOrDec(MeanDR(rows2, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False), LSigFig, False)
                        numSD = SigFigOrDec(StdDevDR(rows2, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False), LSigFig, False)

                        var1 = rows2.Length

                        'var3 = CDec(Format(RoundToDecimal(numSD / numMean * 100, 10), "#0.0")) 'for testing
                        'var4 = CDec(Format(RoundToDecimal(((numMean / var1) - 1) * 100, 10), "#0.0")) 'for testing
                        If CDec(numMean) = 0 Then
                            var3 = 0 'CDec(Format(RoundToDecimal(numSD / numMean * 100, 10), "#0.0")) 'for testing
                            var4 = 0 'CDec(Format(RoundToDecimal(((numMean / var1) - 1) * 100, 10), "#0.0")) 'for testing
                        Else
                            'var3 = CDec(Format(RoundToDecimal(numSD / numMean * 100, 10), "#0.0")) 'for testing
                            'var4 = CDec(Format(RoundToDecimal(((numMean / nomConc) - 1) * 100, 10), "#0.0")) 'for testing

                            var3 = CDec(Format(RoundToDecimalRAFZ(RoundToDecimalRAFZ(numSD / numMean * 100, intQCDec + 1), intQCDec), strQCDec)) 'precision/CV
                            var4 = CDec(Format(RoundToDecimalRAFZ(RoundToDecimalRAFZ(((numMean / nomConc) - 1) * 100, intQCDec + 1), intQCDec), strQCDec)) 'accuracy/bias
                            intCt = intCt + 1

                            arrPrec(intCt) = var3
                            arrAcc(intCt) = var4

                        End If



                    End If


                Next

                ReDim Preserve arrPrec(intCt)
                ReDim Preserve arrAcc(intCt)
                ctQCs = intCt

                ''debugging
                ' '''''''''console.writeline("Start arrPrec AssessQCs")
                'For Count2 = 1 To intCt
                '    If Count2 = 1 Then
                '        var1 = arrPrec(Count2)
                '    Else
                '        var1 = var1 & ChrW(9) & arrPrec(Count2)
                '    End If
                'Next
                ' '''''''''console.writeline(var1)
                ' '''''''''console.writeline("End arrPrec AssessQCs")

                ''debugging
                ' '''''''''console.writeline("Start arrAcc AssessQCs")
                'For Count2 = 1 To intCt
                '    If Count2 = 1 Then
                '        var1 = arrAcc(Count2)
                '    Else
                '        var1 = var1 & ChrW(9) & arrAcc(Count2)
                '    End If
                'Next
                ' '''''''''console.writeline(var1)
                ' '''''''''console.writeline("End arrAcc AssessQCs")

                'legend:
                int30 = FindRow("Analyte Mean Accuracy Min", tblWatsonAnalRefTable, "Item") 'bias
                int40 = FindRow("Analyte Mean Accuracy Max", tblWatsonAnalRefTable, "Item")
                int50 = FindRow("Analyte Precision Min", tblWatsonAnalRefTable, "Item") '%CV
                int60 = FindRow("Analyte Precision Max", tblWatsonAnalRefTable, "Item")
                If ctQCs = 0 Then
                    var1 = Format(0, strQCDec)
                    tblWatsonAnalRefTable.Rows.Item(int30).BeginEdit()
                    tblWatsonAnalRefTable.Rows.Item(int30).Item(Count1) = var1
                    tblWatsonAnalRefTable.Rows.Item(int30).EndEdit()
                    var1 = Format(0, strQCDec)
                    tblWatsonAnalRefTable.Rows.Item(int40).BeginEdit()
                    tblWatsonAnalRefTable.Rows.Item(int40).Item(Count1) = var1
                    tblWatsonAnalRefTable.Rows.Item(int40).EndEdit()
                    var1 = Format(0, strQCDec)
                    tblWatsonAnalRefTable.Rows.Item(int50).BeginEdit()
                    tblWatsonAnalRefTable.Rows.Item(int50).Item(Count1) = var1
                    tblWatsonAnalRefTable.Rows.Item(int50).EndEdit()
                    var1 = Format(0, strQCDec)
                    tblWatsonAnalRefTable.Rows.Item(int60).BeginEdit()
                    tblWatsonAnalRefTable.Rows.Item(int60).Item(Count1) = var1
                    tblWatsonAnalRefTable.Rows.Item(int60).EndEdit()
                Else
                    var1 = Format(CDec(GetMin(arrAcc, ctQCs)), strQCDec)
                    tblWatsonAnalRefTable.Rows.Item(int30).BeginEdit()
                    tblWatsonAnalRefTable.Rows.Item(int30).Item(Count1) = var1
                    tblWatsonAnalRefTable.Rows.Item(int30).EndEdit()
                    var1 = Format(CDec(GetMax(arrAcc, ctQCs)), strQCDec)
                    tblWatsonAnalRefTable.Rows.Item(int40).BeginEdit()
                    tblWatsonAnalRefTable.Rows.Item(int40).Item(Count1) = var1
                    tblWatsonAnalRefTable.Rows.Item(int40).EndEdit()
                    var1 = Format(CDec(GetMin(arrPrec, ctQCs)), strQCDec)
                    tblWatsonAnalRefTable.Rows.Item(int50).BeginEdit()
                    tblWatsonAnalRefTable.Rows.Item(int50).Item(Count1) = var1
                    tblWatsonAnalRefTable.Rows.Item(int50).EndEdit()
                    var1 = Format(CDec(GetMax(arrPrec, ctQCs)), strQCDec)
                    tblWatsonAnalRefTable.Rows.Item(int60).BeginEdit()
                    tblWatsonAnalRefTable.Rows.Item(int60).Item(Count1) = var1
                    tblWatsonAnalRefTable.Rows.Item(int60).EndEdit()
                End If




            Next

        Catch ex As Exception

            ctQCs = 0
            int1 = 1
            int30 = FindRow("Analyte Mean Accuracy Min", tblWatsonAnalRefTable, "Item") 'bias
            int40 = FindRow("Analyte Mean Accuracy Max", tblWatsonAnalRefTable, "Item")
            int50 = FindRow("Analyte Precision Min", tblWatsonAnalRefTable, "Item") '%CV
            int60 = FindRow("Analyte Precision Max", tblWatsonAnalRefTable, "Item")
            If ctQCs = 0 Then
                var1 = Format(0, strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int30).BeginEdit()
                tblWatsonAnalRefTable.Rows.Item(int30).Item(int1) = var1
                tblWatsonAnalRefTable.Rows.Item(int30).EndEdit()
                var1 = Format(0, strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int40).BeginEdit()
                tblWatsonAnalRefTable.Rows.Item(int40).Item(int1) = var1
                tblWatsonAnalRefTable.Rows.Item(int40).EndEdit()
                var1 = Format(0, strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int50).BeginEdit()
                tblWatsonAnalRefTable.Rows.Item(int50).Item(int1) = var1
                tblWatsonAnalRefTable.Rows.Item(int50).EndEdit()
                var1 = Format(0, strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int60).BeginEdit()
                tblWatsonAnalRefTable.Rows.Item(int60).Item(int1) = var1
                tblWatsonAnalRefTable.Rows.Item(int60).EndEdit()
            Else
                var1 = Format(CDec(GetMin(arrAcc, ctQCs)), strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int30).BeginEdit()
                tblWatsonAnalRefTable.Rows.Item(int30).Item(int1) = var1
                tblWatsonAnalRefTable.Rows.Item(int30).EndEdit()
                var1 = Format(CDec(GetMax(arrAcc, ctQCs)), strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int40).BeginEdit()
                tblWatsonAnalRefTable.Rows.Item(int40).Item(int1) = var1
                tblWatsonAnalRefTable.Rows.Item(int40).EndEdit()
                var1 = Format(CDec(GetMin(arrPrec, ctQCs)), strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int50).BeginEdit()
                tblWatsonAnalRefTable.Rows.Item(int50).Item(int1) = var1
                tblWatsonAnalRefTable.Rows.Item(int50).EndEdit()
                var1 = Format(CDec(GetMax(arrPrec, ctQCs)), strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int60).BeginEdit()
                tblWatsonAnalRefTable.Rows.Item(int60).Item(int1) = var1
                tblWatsonAnalRefTable.Rows.Item(int60).EndEdit()
            End If

        End Try

        Try

            'now do min regression
            'str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.STUDYID, ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX,
            ' ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, ANARUNREGPARAMETERS.PARAMETERVALUE, ANALYTICALRUNANALYTES.RSQUARED,
            ' ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAY.MASTERASSAYID,  ANALYTICALRUN.RUNTYPEID "

            'determine if Calibr Stds are built from assigned samples
            tblP = tblTableProperties
            strFP = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLCONFIGREPORTTABLES = 2"
            rowsP = tblP.Select(strFP)
            int1 = rowsP(0).Item("BOOLINCLUDEPSAE")
            If int1 = 0 Then
                boolPSAE = False
            Else
                boolPSAE = True
            End If

            'don't use boolIncludePSAE from tblTableProperties
            'use tblTableProperties
            boolPSAE = boolIncludePSAE


            'RegrCon table cannot have samples assigned
            'tbl1 = tblReportTable
            'strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLCONFIGREPORTTABLES = 2 AND BOOLREQUIRESSAMPLEASSIGNMENT = -1"
            'rows1 = tbl1.Select(strF)
            'intCt1 = rows1.Length
            'If intCt1 = 0 Then 'ignore
            '    boolDoAssigned = False
            '    tbl2 = tblBCStdConcs
            'Else
            '    boolDoAssigned = True
            '    tbl2 = tblAssignedSamples
            '    'Exit Sub
            'End If

            'tbl2 = tblAssignedSamples
            'tblQC = tblBCQCConcs

            'Dim arrBCStdActual(intCt1)

            For Count1 = 1 To ctAnalytes

                int30 = FindRow("Minimum r^2", tblWatsonAnalRefTable, "Item")
                'these filters are bad
                'masterassayid isn't applicable anymore
                'If boolIncludePSAE Then
                '    strF = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " and MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND RUNTYPEID > 0 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                'Else
                '    strF = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " and MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                'End If
                If boolIncludePSAE Then
                    strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND RUNTYPEID > 0 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                Else
                    strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                End If


                '"' AND (" & strFFF & ")"
                If Len(strFFF) = 0 Then
                    If boolIncludePSAE Then
                        strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND RUNTYPEID > 0 AND RUNANALYTEREGRESSIONSTATUS <> 4 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "'"
                    Else
                        strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "'"
                    End If
                Else
                    If boolIncludePSAE Then
                        strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND RUNTYPEID > 0 AND RUNANALYTEREGRESSIONSTATUS <> 4 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "' AND (" & strFFF & ")"
                    Else
                        strF = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4 AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "' AND (" & strFFF & ")"
                    End If
                End If

                Dim rowsR() As DataRow
                Dim intRows As Short
                Try
                    rowsR = tblRegConAll.Select(strF)
                Catch ex As Exception
                    var1 = var1
                End Try

                intRows = rowsR.Length
                ReDim arrAcc(intRows)
                For Count2 = 0 To intRows - 1
                    var1 = rowsR(Count2).Item("RSQUARED")
                    arrPrec(Count2 + 1) = var1
                Next
                str1 = "0."
                For Count2 = 1 To LRegrSigFigs
                    str1 = str1 & "0"
                Next
                var1 = Format(CDec(GetMin(arrPrec, intRows)), str1)
                tblWatsonAnalRefTable.Rows.Item(int30).BeginEdit()
                tblWatsonAnalRefTable.Rows.Item(int30).Item(Count1) = var1
                tblWatsonAnalRefTable.Rows.Item(int30).EndEdit()

            Next

        Catch ex As Exception
            'int30 = FindRow("Analyte Mean Accuracy Min", tblWatsonAnalRefTable, "Item") 'bias
            'tblWatsonAnalRefTable.Rows.Item(int30).BeginEdit()
            'var1 = 0 'Format(CDec(GetMin(arrPrec, intRows)), str1)
            'tblWatsonAnalRefTable.Rows.Item(int30).Item(Count1) = var1
            'tblWatsonAnalRefTable.Rows.Item(int30).EndEdit()
        End Try

end1:

    End Sub

End Module
