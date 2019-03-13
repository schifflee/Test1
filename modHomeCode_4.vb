Option Compare Text

Imports System
Imports System.IO
Imports System.Text

Module modHomeCode_4


    Function GetTabColumn(ByVal strN As String) As String

        GetTabColumn = ""

        Dim dtbl1 As System.Data.DataTable
        Dim Count1 As Short
        Dim Count2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim boolIgnore As Boolean = False

        dtbl1 = tblData
        str1 = ""

        Select Case strN
            Case Is = "ID_TBLDATA"
                boolIgnore = True
            Case Is = "ID_TBLSTUDIES"
                boolIgnore = True
            Case Is = "ID_TBLASSAYTECHNIQUE"
                boolIgnore = True
            Case Is = "ID_TBLANTICOAGULANT"
                boolIgnore = True
            Case Is = "CHARCORPORATESTUDYID"
                str1 = "Data"
            Case Is = "CHARPROTOCOLNUMBER"
                str1 = "Data"
            Case Is = "ID_TBLVOLUMEUNITS"
                boolIgnore = True
            Case Is = "ID_TBLTEMPERATURES"
                boolIgnore = True
            Case Is = "ID_SUBMITTEDBY"
                boolIgnore = True
            Case Is = "ID_SUBMITTEDTO"
                boolIgnore = True
            Case Is = "ID_INSUPPORTOF"
                boolIgnore = True
            Case Is = "CHARDATAARCHIVALLOCATION"
                str1 = "Data"
            Case Is = "CHARSPONSORSTUDYNUMBER"
                str1 = "Data"
            Case Is = "CHARSPONSORSTUDYTITLE"
                str1 = "Data"
            Case Is = "UPSIZE_TS"
                boolIgnore = True
            Case Is = "NUMSIGFIGS"
                str1 = "Config"
            Case Is = "CHARDATEFORMAT"
                str1 = "Config"
            Case Is = "CHARTEXTDATEFORMAT"
                str1 = "Config"
            Case Is = "NUMDECIMALS"
                str1 = "Config"
            Case Is = "BOOLUSESIGFIGS"
                str1 = "Config"
            Case Is = "CHARTIMEZONE"
                str1 = "Config"
            Case Is = "CHAROUTLIERMETHOD"
                str1 = "Data"
            Case Is = "BOOLENTIREREPORT"
                str1 = "Config"
            Case Is = "BOOLUSESPECRND"
                str1 = "Config"
            Case Is = "NUMREGRSIGFIGS"
                str1 = "Config"
            Case Is = "NUMR2SIGFIGS"
                str1 = "Config"
            Case Is = "CHARUNITS"
                str1 = "Config"
            Case Is = "INTQCPERCDECPLACES"
                str1 = "Config"
            Case Is = "BOOLQAEVENTBORDER"
                str1 = "Config"
            Case Is = "BOOLALLOWEXCLSAMPLES"
                str1 = "Config"
            Case Is = "BOOLALLOWGUWUACCCRIT"
                str1 = "Config"
            Case Is = "DTSTUDYSTARTDATE"
                str1 = "Data"
            Case Is = "DTSTUDYENDDATE"
                str1 = "Data"
            Case Is = "INTCOMMAFORMAT"
                str1 = "Config"
            Case Is = "BOOLBLUEHYPERLINK"
                str1 = "Config"
            Case Is = "BOOLREDBOLDFONT"
                str1 = "Config"


            Case Is = "NUMSIGFIGSAREA"
                str1 = "Config"
            Case Is = "NUMDECIMALSAREA"
                str1 = "Config"
            Case Is = "BOOLUSESIGFIGSAREA"
                str1 = "Config"
            Case Is = "BOOLUSESPECRNDAREA"
                str1 = "Config"


            Case Is = "NUMSIGFIGSAREARATIO"
                str1 = "Config"
            Case Is = "NUMDECIMALSAREARATIO"
                str1 = "Config"
            Case Is = "BOOLUSESIGFIGSAREARATIO"
                str1 = "Config"
            Case Is = "BOOLUSESPECRNDAREARATIO"
                str1 = "Config"



            Case Is = "BOOLUSESIGFIGSREGR"
                str1 = "Config"
            Case Is = "NUMREGRDEC"
                str1 = "Config"
            Case Is = "BOOLUSEREGRSCINOT"
                str1 = "Config"

            Case Is = "BOOLNOMCONCPAREN"
                str1 = "Config"
            Case Is = "CHARSTPAGE"
                str1 = "Config"

            Case Is = "BOOLTABLEDTTIMESTAMP"
                str1 = "Config"

            Case Is = "BOOLFOOTNOTEQCMEAN"
                str1 = "Config"
            Case Is = "BOOLFLIPHEADER"
                str1 = "Config"

            Case Is = "BOOLQCNA"
                str1 = "Config"

            Case Is = "BOOLBQL"
                str1 = "Config"

            Case Is = "CHARBQL"
                str1 = "Config"

            Case Is = "BOOLIGNOREFC"
                str1 = "Config"

            Case Is = "BOOLPSL"
                str1 = "Config"

            Case Is = "CHARCAPTIONTRAILER"
                str1 = "Config"

            Case Is = "BOOLRECSIGFIG"
                str1 = "Config"

            Case Is = "BOOLSD2"
                str1 = "Config"

            Case Is = "BOOLDIFFCOLSTATS"
                str1 = "Config"

            Case Is = "CHARCAPTIONFOLLOW"
                str1 = "Config"

            Case Is = "BOOLUSERSD"
                str1 = "Config"

            Case Is = "BOOLTABLELABELSECTION"
                str1 = "Config"

            Case Is = "NUMTABLEFONTSIZE"
                str1 = "Config"

                '20190108 LEE:
            Case Is = "BOOLCALIBRTABLETITLE"
                str1 = "Config"

                'don't think these are needed here

                'Case "BOOLROUNDFIVEEVEN"
                '    str1 = True

                'Case "BOOLROUNDFIVEAWAY"
                '    str1 = True

                'Case "BOOLCRITFULLPREC"
                '    str1 = True

                'Case "BOOLCRITROUNDED"
                '    str1 = True

                'Case "BOOLMEANFULLPREC"
                '    str1 = True

                'Case "BOOLMEANROUNDED"
                '    str1 = True


        End Select

        If boolIgnore Then
        Else
            GetTabColumn = str1
        End If

    End Function

    Sub FillAnalRunSum()

        'This is for frmHome Analytical Runs

        'Dim tbl as System.Data.DataTable
        Dim dtbl As System.Data.DataTable
        Dim strF As String
        Dim str1 As String
        Dim str2 As String
        Dim Count1 As Short
        Dim Count2 As Short
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim int4 As Short
        Dim var1, var2, var3, var4
        'Dim dv as system.data.dataview
        Dim drows() As DataRow
        Dim strS As String

        'tbl = tblAnalRunSum
        dtbl = tblAnalyticalRunSummary
        'dv = dtbl.DefaultView

        'filter tblAnalRunSum, if needed

        Dim bool As Boolean
        bool = boolStopRBS
        boolStopRBS = True

        Dim strF1 As String = ""
        Dim strF2 As String = ""
        Dim strF3 As String = ""
        Dim strF4 As String = ""
        Dim strF5 As String = ""
        Dim strF6 As String = ""
        Dim strFTot As String = ""

        '20171219 LEE: boolInThisRunsAssayID is for runs with calibration curves

        If frmH.chkAll.Checked Then
            'strF1 = "RUNTYPEID > 0 AND boolInThisRunsAssayID = 'Yes' AND RUNANALYTEREGRESSIONSTATUS > -2" 'do -2 here because blank row has -1
            '20171118 LEE: Hmmm. boolInThisRunsAssayID = 'Yes' filters out 'No Regression Performed', which is not intended
            strF1 = "RUNTYPEID > 0 AND RUNANALYTEREGRESSIONSTATUS > -2" 'do -2 here because blank row has -1
            strFTot = strF1
        Else
            If frmH.chkAccepted.Checked Then
                strF2 = "RUNANALYTEREGRESSIONSTATUS = 3"
                If Len(strFTot) = 0 Then
                    strFTot = strF2
                Else
                    strFTot = strFTot & " OR " & strF2
                End If
            End If
            If frmH.chkRejected.Checked Then
                strF3 = "RUNANALYTEREGRESSIONSTATUS = 4"
                If Len(strFTot) = 0 Then
                    strFTot = strF3
                Else
                    strFTot = strFTot & " OR " & strF3
                End If
            End If
            If frmH.chkRegrPerformed.Checked Then
                strF4 = "RUNANALYTEREGRESSIONSTATUS = 2"
                If Len(strFTot) = 0 Then
                    strFTot = strF4
                Else
                    strFTot = strFTot & " OR " & strF4
                End If
            End If
            If frmH.chkNoRegrPerformed.Checked Then
                strF5 = "RUNANALYTEREGRESSIONSTATUS = 1"
                If Len(strFTot) = 0 Then
                    strFTot = strF5
                Else
                    strFTot = strFTot & " OR " & strF5
                End If
            End If

            If frmH.chkPSAE.Checked Then
                'strF6 = "RUNTYPEID > 0"
                strF6 = "RUNTYPEID = 3"
                If Len(strFTot) = 0 Then
                    strFTot = strF6
                Else
                    'strFTot = "(" & strFTot & ") AND " & strF6
                    strFTot = "(" & strFTot & ") OR " & strF6
                End If
            Else
                strF6 = "RUNTYPEID <> 3"
                If Len(strFTot) = 0 Then
                    strFTot = "RUNANALYTEREGRESSIONSTATUS = -10" 'strF6
                Else
                    strFTot = "(" & strFTot & ") AND " & strF6
                    'strFTot = "(" & strFTot & ") OR " & strF6
                End If
            End If

        End If


        '20180716 LEE:
        If Len(strFTot) = 0 Then
            strFTot = "RUNANALYTEREGRESSIONSTATUS = -10"
        Else
            'strFTot = "(" & strFTot & ") AND boolInThisRunsAssayID = 'Yes' OR RUNANALYTEREGRESSIONSTATUS = -1" 'RUNANALYTEREGRESSIONSTATUS = -1 to get blank rows
            '20180717 LEE
            strFTot = "(" & strFTot & ") OR RUNANALYTEREGRESSIONSTATUS = -1" 'RUNANALYTEREGRESSIONSTATUS = -1 to get blank rows
        End If


        'If Len(strFTot) = 0 Then
        '    strFTot = "RUNANALYTEREGRESSIONSTATUS = -10"
        'Else

        '    '20180713 LEE:
        '    '??? Why al this????

        '    'If frmH.chkAll.Checked Then
        '    '    '20171118 LEE: Hmmm. boolInThisRunsAssayID = 'Yes' filters out 'No Regression Performed', which is not intended

        '    '    '20171219 LEE:
        '    '    'Not true. Problem was assigning boolInThisRunsAssayID in DoPrepare
        '    '    'strFTot = "(" & strFTot & ") AND boolInThisRunsAssayID = 'Yes'"
        '    '    '20180316 LEE:
        '    '    'NO! boolInThisRunsAssayID refers only to runs with calibr curves
        '    'ElseIf frmH.chkNoRegrPerformed.Checked Then
        '    '    strFTot = "(" & strFTot & ") OR RUNANALYTEREGRESSIONSTATUS = -1" 'RUNANALYTEREGRESSIONSTATUS = -1 to get blank rows
        '    'Else
        '    '    'strFTot = "(" & strFTot & ") AND boolInThisRunsAssayID = 'Yes' OR RUNANALYTEREGRESSIONSTATUS = -1" 'RUNANALYTEREGRESSIONSTATUS = -1 to get blank rows
        '    '    '20180316 LEE:
        '    '    'get rid of boolInThisRunsAssayID!!!
        '    '    strFTot = "(" & strFTot & ") OR RUNANALYTEREGRESSIONSTATUS = -1" 'RUNANALYTEREGRESSIONSTATUS = -1 to get blank rows
        '    'End If
        '    'strFTot = "(" & strFTot & ") AND boolInThisRunsAssayID = 'Yes' OR RUNANALYTEREGRESSIONSTATUS = -1" 'RUNANALYTEREGRESSIONSTATUS = -1 to get blank rows
        'End If

        ''''''''''''''''console.writeline(strF)
        Dim dv2 As System.Data.DataView = New DataView(tblAnalRunSum)
        dv2.RowFilter = strFTot

        int1 = dv2.Count 'debug
        int1 = int1
        'ReSort this correctly


        Dim tbl As System.Data.DataTable = dv2.ToTable

        'loop through tbl, find appropriate dv entries, then record comment value and boolInclude
        int1 = tbl.Rows.Count
        For Count1 = 0 To int1 - 1
            var1 = NZ(tbl.Rows.Item(Count1).Item("Watson Run ID"), "") 'Watson Run ID
            int2 = Len(var1)
            If Len(var1) = 0 Then
            Else
                'find appropriate entry in dv
                var2 = tbl.Rows.Item(Count1).Item("Analyte") 'Analyte

                '20190206 LEE:
                strF = "id_tblStudies = " & id_tblStudies & " and intWatsonRunID = " & var1 & " and charAnalyte = '" & CleanText(CStr(var2)) & "'"
                drows = dtbl.Select(strF)

                If drows.Length = 0 Then
                    'tbl.Rows.item(Count1).Item("User Comments") = ""
                    'tbl.Rows.item(Count1).Item("boolInclude") = True
                Else
                    'record entry
                    var3 = drows(0).Item("charUserComments")
                    tbl.Rows.Item(Count1).Item("User Comments") = var3
                    var3 = drows(0).Item("boolInclude")
                    tbl.Rows.Item(Count1).Item("boolInclude") = var3
                    var3 = drows(0).Item("boolIncludeRegr")
                    tbl.Rows.Item(Count1).Item("boolIncludeRegr") = var3

                End If
            End If
        Next

        Dim dv As System.Data.DataView = New DataView(tbl)
        dv.AllowNew = False
        dv.AllowDelete = False
        dv.AllowEdit = True
        'frmh.dganalyticalRunSummary.DataSource = dv

        Dim dgv As DataGridView = frmH.dgvAnalyticalRunSummary

        dgv.DataSource = dv
        'autosize columns
        dgv.AllowUserToResizeColumns = True
        dgv.AllowUserToResizeRows = True
        dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)


        dgv.Columns.Item("Watson Run ID").Width = 60
        dgv.Columns.Item("User Comments").Width = 500

        dgv.Columns("RUNANALYTEREGRESSIONSTATUS").Visible = False

        'dgv.SelectionMode = DataGridViewSelectionMode.ColumnHeaderSelect
        dgv.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect

        dgv.Refresh()

        'record intusercomments and boolexcludepsae
        If boolFromAnalSum Then
        Else
            Dim dv1 As System.Data.DataView
            dv1 = frmH.dgvReports.DataSource
            If dv1.Count = 0 Then
                frmH.rbUseWatsonComments.Checked = True
                'frmH.rbAnalRunsExclPSAE.Checked = True

                frmH.chkAll.Checked = True
                frmH.chkAccepted.Checked = False
                frmH.chkRejected.Checked = False
                frmH.chkRegrPerformed.Checked = False
                frmH.chkNoRegrPerformed.Checked = False
                frmH.chkPSAE.Checked = False

            Else
                int1 = frmH.dgvReports.CurrentRow.Index
                int2 = NZ(dv1.Item(int1).Item("intUserComments"), 1)
                Select Case int2
                    Case 1
                        frmH.rbUseWatsonComments.Checked = True
                    Case 2
                        frmH.rbUseUserComments.Checked = True
                End Select

                'BOOLALLAR chkAll
                'BOOLACCAR chkAccepted
                'BOOLREJAR chkRejected
                'BOOLREGRAR chkRegrPerformed
                'BOOLNOREGRAR chkNoRegrPerformed
                'BOOLINCLPSAE chkPSAE

                int2 = NZ(dv1.Item(int1).Item("BOOLALLAR"), -1) 'checkbox will take 1 or -1 for checked
                frmH.chkAll.Checked = int2

                int2 = NZ(dv1.Item(int1).Item("BOOLACCAR"), 0)
                frmH.chkAccepted.Checked = int2

                int2 = NZ(dv1.Item(int1).Item("BOOLREJAR"), 0)
                frmH.chkRejected.Checked = int2

                int2 = NZ(dv1.Item(int1).Item("BOOLREGRAR"), 0)
                frmH.chkRegrPerformed.Checked = int2

                int2 = NZ(dv1.Item(int1).Item("BOOLNOREGRAR"), 0)
                frmH.chkNoRegrPerformed.Checked = int2

                int2 = NZ(dv1.Item(int1).Item("BOOLINCLPSAE"), 0)
                frmH.chkPSAE.Checked = int2

            End If
        End If

        Call NumberAnalSumRows(dgv)

        Call SetAnalRunSummaryColumns()


        boolStopRBS = bool

        ''set tblTableProperties
        ''20160907 LEE: boolexcludepsae in tblTableProperties deprecated
        'Dim tblT As System.Data.DataTable
        'Dim rowsT() As DataRow
        'tblT = tblTableProperties
        'strF = "ID_TBLCONFIGREPORTTABLES = 1 AND ID_TBLSTUDIES = " & id_tblStudies
        'rowsT = tblT.Select(strF)

        'If rowsT.Length = 0 Then
        'Else
        '    rowsT(0).BeginEdit()
        '    If frmH.rbAnalRunsExclPSAE.Checked Then
        '        rowsT(0).Item("BOOLINCLUDEPSAE") = 0
        '    Else
        '        rowsT(0).Item("BOOLINCLUDEPSAE") = 1
        '    End If
        '    rowsT(0).EndEdit()
        'End If

    End Sub

    Sub NumberAnalSumRows(dgv As DataGridView)

        Dim Count1 As Short
        Dim intCt As Short = 0
        Dim intRS As Short = 0

        dgv.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedHeaders
        'myGrid.RowHeadersWidth = ClsConsts.vSettings.LengthMyMoviesRowHeader;

        Dim num1 As Single
        num1 = dgv.RowHeadersWidth
        'dgv.RowHeadersWidth = 75


        For Count1 = 0 To dgv.Rows.Count - 1

            intRS = NZ(dgv("RUNANALYTEREGRESSIONSTATUS", Count1).Value, 1)
            If intRS = -1 Then
                intCt = 0
            Else
                intCt = intCt + 1
                dgv.Rows(Count1).HeaderCell.Value = intCt.ToString
            End If

        Next

        dgv.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders)

        dgv.Refresh()

    End Sub

    Sub FillTableReports(ByVal boolInitial As Boolean)

        Dim intWid As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim strDTable As String
        Dim dtbls As System.Data.DataTable
        Dim dtbl As System.Data.DataTable
        Dim dg As DataGrid
        Dim boolRO As Boolean
        Dim var1, var2, var3
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim bool As Boolean
        Dim drow As DataRow
        Dim drows() As DataRow
        'Dim erows() As DataRow
        Dim int1 As Short
        Dim dv As System.Data.DataView
        Dim dgv As DataGridView
        Dim boolA As Short
        Dim strF As String
        Dim rows() As DataRow
        Dim strS As String


        strDTable = "tblReportTables"
        'tblReportTables = new System.Data.DataTable
        dtbls = tblReportTables
        dtbl = tblReportTable
        dgv = frmH.dgvReportTableConfiguration
        boolRO = True
        intWid = 250 'autosize min col width

        frmH.dgvReportTableConfiguration.AllowUserToResizeColumns = True
        frmH.dgvReportTableConfiguration.AllowUserToResizeRows = True
        frmH.dgvReportTableConfiguration.RowHeadersWidth = 25
        'frmH.dgvReportTableConfiguration.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        'frmH.dgvReportTableConfiguration.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
        frmH.dgvReportTableConfiguration.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        frmH.dgvReportTableConfiguration.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        frmH.dgvReportTableConfiguration.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        frmH.dgvReportTableConfiguration.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

        'frmH.dgvReportTableConfiguration.AllowUserToResizeRows = True
        'frmH.dgvReportTableConfiguration.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        'frmH.dgvReportTableConfiguration.AutoResizeRows()
        'frmH.dgvReportTableConfiguration.ColumnHeadersDefaultCellStyle.Font.Bold=True

        Dim drowF() As DataRow
        Dim boolF As Boolean

        strF = "ID_TBLSTUDIES = " & id_tblStudies
        strS = "INTORDER ASC"
        rows = dtbl.Select(strF, strS)

        'first check to ensure that dtbl has newly-added tables

        Dim dtbl2 As System.Data.DataTable
        Dim rows2() As DataRow
        Dim rows3() As DataRow
        Dim strF1 As String
        Dim strS1 As String
        Dim strF2 As String
        Dim boolHit As Boolean
        Dim maxID As Int64
        strF1 = "ID_TBLCONFIGREPORTTYPE < 1000"
        strS1 = "ID_TBLCONFIGREPORTTABLES ASC"
        dtbl2 = tblConfigReportTables
        rows2 = dtbl2.Select(strF1, strS1)
        boolHit = False

        maxID = GetMaxID("TBLREPORTTABLE", 1, False)

        For Count1 = 0 To rows2.Length - 1
            var1 = rows2(Count1).Item("ID_TBLCONFIGREPORTTABLES")
            strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLCONFIGREPORTTABLES = " & var1
            rows3 = dtbl.Select(strF2)
            If rows3.Length > 0 Or id_tblStudies = 0 Then 'OK
            Else 'add row
                boolHit = True
                'maxID = GetMaxID("TBLREPORTTABLE", 1, True)
                maxID = maxID + 1
                Dim nRow1 As DataRow = dtbl.NewRow

                nRow1.BeginEdit()
                nRow1("ID_TBLREPORTTABLE") = maxID
                nRow1("ID_TBLSTUDIES") = id_tblStudies
                nRow1("ID_TBLCONFIGREPORTTABLES") = var1
                nRow1("INTORDER") = rows2(Count1).Item("INTORDER") '100
                Select Case var1
                    Case 36, 37, 38
                        nRow1("CHARPAGEORIENTATION") = "L"
                    Case Else
                        nRow1("CHARPAGEORIENTATION") = "P"
                End Select

                nRow1("BOOLINCLUDE") = 0
                nRow1("BOOLREQUIRESSAMPLEASSIGNMENT") = rows2(Count1).Item("BOOLREQUIRESSAMPLEASSIGNMENT") ' 0
                Try
                    nRow1("BOOLPLACEHOLDER") = rows2(Count1).Item("BOOLPLACEHOLDER") ' 0
                Catch ex As Exception
                    var1 = var1
                End Try
                nRow1("UPSIZE_TS") = System.DBNull.Value
                nRow1("CHARSTABILITYPERIOD") = System.DBNull.Value
                nRow1("CHARHEADINGTEXT") = rows2(Count1).Item("CHARTABLENAME")
                nRow1("CHARSTYLE") = "Style 1"
                nRow1("INTEGNUM") = 1

                nRow1.EndEdit()

                dtbl.Rows.Add(nRow1)

            End If
        Next

        If boolHit Then 'save

            PutMaxID("TBLREPORTTABLE", maxID)

            If boolGuWuOracle Then
                Try
                    ta_tblReportTable.Update(tblReportTable)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLREPORTTABLE.Merge('ds2005.TBLREPORTTABLE, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblReportTableAcc.Update(tblReportTable)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLREPORTTABLE.Merge('ds2005Acc.TBLREPORTTABLE, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblReportTableSQLServer.Update(tblReportTable)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLREPORTTABLE.Merge('ds2005Acc.TBLREPORTTABLE, True)
                End Try
            End If

            strF = "ID_TBLSTUDIES = " & id_tblStudies
            strS = "INTORDER ASC"
            rows = dtbl.Select(strF, strS)

        End If

        tblReportTables = dtbl.Clone 'copies structure, but not data
        dtbls = tblReportTables

        'debug
        For Count1 = dtbls.Columns.Count - 1 To 0 Step -1
            str1 = dtbls.Columns(Count1).ColumnName
            str1 = str1
        Next

        Try
            Dim col1 As New DataColumn
            col1.DataType = System.Type.GetType("System.String")
            col1.ColumnName = "CHARTABLENAME"
            col1.Caption = "Table"
            dtbls.Columns.Add(col1)


        Catch ex As Exception

        End Try

        Try
            'alter datatype
            dtbls.Columns("BOOLINCLUDE").DataType = System.Type.GetType("System.Boolean")
            dtbls.Columns("boolRequiresSampleAssignment").DataType = System.Type.GetType("System.Boolean")
            Try
                dtbls.Columns("BOOLPLACEHOLDER").DataType = System.Type.GetType("System.Boolean")
            Catch ex As Exception
                var1 = var1
            End Try

        Catch ex As Exception

        End Try

        'delete all rows from table
        boolLoad = True 'to disable a dgvReportTablesConfig Validation event
        dtbls.Rows.Clear()
        boolLoad = False

        drows = dtbl.Select(strF)
        int1 = drows.Length

        'fill rows
        For Count2 = 0 To int1 - 1
            If Count2 = 21 Then
                var1 = 1
            End If
            'str1 = "id_tblConfigReportType < 1000" ' & int1
            var1 = drows(Count2).Item("ID_TBLCONFIGREPORTTABLES")
            str1 = "ID_TBLCONFIGREPORTTABLES = " & var1
            'Dim eRows() As DataRow = tblConfigReportTables.Select(str1, "intOrder ASC")
            Dim eRows() As DataRow = tblConfigReportTables.Select(str1)
            var2 = eRows(0).Item("CHARTABLENAME")
            drow = dtbls.NewRow
            drow.BeginEdit()
            drow("CHARTABLENAME") = var2
            For Count3 = 0 To dtbl.Columns.Count - 1
                str1 = dtbl.Columns(Count3).ColumnName
                var1 = drows(Count2).Item(str1) 'debugging
                drow(str1) = drows(Count2).Item(str1)
            Next
            drow.EndEdit()
            dtbls.Rows.Add(drow)
        Next


        'refresh datagrid
        dv = dtbls.DefaultView
        dv.AllowDelete = False
        dv.AllowNew = False
        dv.AllowEdit = True
        dv.Sort = strS

        'dg.DataSource = dv
        'dgv.AutoGenerateColumns = True
        dgv.DataSource = dv
        'dgv.AutoGenerateColumns = False
        'dg.Refresh()

        'first hide all columns
        For Count1 = 0 To dgv.ColumnCount - 1
            dgv.Columns(Count1).Visible = False
        Next

        dgv.Columns.Item("CHARSTYLE").DisplayIndex = 9
        dgv.Columns.Item("CHARSTYLE").Visible = False
        dgv.Columns.Item("CHARSTYLE").HeaderText = "Table Style"

        dgv.Columns.Item("CHARSTABILITYPERIOD").DisplayIndex = 8
        dgv.Columns.Item("CHARSTABILITYPERIOD").Visible = False 'True
        dgv.Columns.Item("CHARSTABILITYPERIOD").HeaderText = "Period Temp**"

        Try
            dgv.Columns.Item("BOOLPLACEHOLDER").DisplayIndex = 7
            dgv.Columns.Item("BOOLPLACEHOLDER").Visible = True
            dgv.Columns.Item("BOOLPLACEHOLDER").HeaderText = "B*"
        Catch ex As Exception
            var1 = var1
        End Try

        dgv.Columns.Item("CHARFCID").DisplayIndex = 6
        dgv.Columns.Item("CHARFCID").Visible = True
        dgv.Columns.Item("CHARFCID").HeaderText = "FC ID*"

        dgv.Columns.Item("BOOLINCLUDE").DisplayIndex = 5
        dgv.Columns.Item("BOOLINCLUDE").Visible = True
        dgv.Columns.Item("BOOLINCLUDE").HeaderText = "Include"

        dgv.Columns.Item("CHARPAGEORIENTATION").DisplayIndex = 4
        dgv.Columns.Item("CHARPAGEORIENTATION").Visible = True ' True 'this is now a dropdownbox column - see below
        dgv.Columns.Item("CHARPAGEORIENTATION").HeaderText = "P/L*"

        'dgv.Columns.Item("cOrientation").DisplayIndex = 4
        'dgv.Columns.Item("cOrientation").Visible = False 'this is now a dropdownbox column - see below
        'dgv.Columns.Item("cOrientation").HeaderText = "P/L*"


        dgv.Columns.Item("INTORDER").DisplayIndex = 3
        dgv.Columns.Item("INTORDER").Visible = True
        dgv.Columns.Item("INTORDER").HeaderText = "Order"

        dgv.Columns.Item("CHARTABLENAME").DisplayIndex = 2
        'dgv.Columns.Item("CHARTABLENAME").Visible = True
        dgv.Columns.Item("CHARTABLENAME").HeaderText = "StudyDoc Table Title"

        dgv.Columns.Item("CHARHEADINGTEXT").DisplayIndex = 1
        dgv.Columns.Item("CHARHEADINGTEXT").Visible = True
        dgv.Columns.Item("CHARHEADINGTEXT").HeaderText = "Table Title"

        dgv.Columns.Item("boolRequiresSampleAssignment").DisplayIndex = 0
        dgv.Columns.Item("boolRequiresSampleAssignment").Visible = True
        dgv.Columns.Item("boolRequiresSampleAssignment").HeaderText = "SA*"

        dgv.Columns.Item("ID_TBLREPORTTABLE").Visible = False



        'dgv.Columns.Item("Table").Frozen = True
        'dgv.Columns.Item("CHARHEADINGTEXT").Frozen = True
        int1 = dgv.Columns.Count


        For Count1 = 0 To int1 - 1
            dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
            'If Count1 = 1 Then
            '    dgv.Columns.Item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            'Else
            '    dgv.AutoResizeColumn(Count1)
            '    dgv.Columns.Item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            'End If
        Next
        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
        'dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        'dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)

        dgv.Columns.Item("INTORDER").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv.Columns.Item("CHARPAGEORIENTATION").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv.Columns.Item("CHARSTABILITYPERIOD").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        'dgv.Columns.Item("INTORDER").SortMode = DataGridViewColumnSortMode.Automatic

        'dgv.Columns.Item("boolFigure").Visible = False

        'dgv.Columns.Item("Table").Width = 250
        dgv.Columns.Item("CHARHEADINGTEXT").MinimumWidth = 350

        'now do comboboxcell
        For Count1 = 0 To dgv.Rows.Count - 1
            Dim cbx As New DataGridViewComboBoxCell
            cbx.Items.Add("P")
            cbx.Items.Add("L")
            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            Try
                dgv("CHARPAGEORIENTATION", Count1) = cbx
            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try
        Next

        Call SetComboCell(frmH.dgvReportTableConfiguration, "CHARPAGEORIENTATION")

        dgv.Refresh()

        'set cmdOrderReportTableConfig Position
        Call SizecmdOrder(dgv, frmH.cmdOrderReportTableConfig, "INTORDER")

        'move cmdresize
        'frmH.cmdResize.Left = frmH.cmdOrderReportTableConfig.Left + frmH.cmdOrderReportTableConfig.Width + 10

        'dgv.AllowUserToResizeRows = True



    End Sub


    Sub OrderReportTableConfig()

        Try

            Dim dgv As DataGridView

            dgv = frmH.dgvReportTableConfiguration

            'dgv.AutoGenerateColumns = False

            dgv.Columns.Item("BOOLINCLUDE").DisplayIndex = 0
            dgv.Columns.Item("boolRequiresSampleAssignment").DisplayIndex = 1
            dgv.Columns.Item("CHARHEADINGTEXT").DisplayIndex = 2
            dgv.Columns.Item("CHARTABLENAME").DisplayIndex = 3
            dgv.Columns.Item("INTORDER").DisplayIndex = 4
            'Try
            '    dgv.Columns.Item("cOrientation").DisplayIndex = 5
            'Catch ex As Exception

            'End Try
            dgv.Columns.Item("CHARPAGEORIENTATION").DisplayIndex = 5
            dgv.Columns.Item("CHARFCID").DisplayIndex = 6
            Try
                dgv.Columns.Item("BOOLPLACEHOLDER").DisplayIndex = 7
            Catch ex As Exception

            End Try
            dgv.Columns.Item("CHARSTABILITYPERIOD").DisplayIndex = 8
            dgv.Columns.Item("CHARSTYLE").DisplayIndex = 9

            'dgv.AutoGenerateColumns = True

            dgv.Columns.Item("CHARHEADINGTEXT").Width = 350
            dgv.Columns.Item("CHARTABLENAME").Width = 250

            dgv.AutoResizeRows()

        Catch ex As Exception

        End Try

    

    End Sub

    Sub SizecmdOrder(ByVal dgv As DataGridView, ByVal cmd As Control, ByVal strField As String)

        Dim Count1 As Short
        Dim str1 As String
        'set cmdOrderReportTableConfig Position
        Dim wd1, wd2
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        'Dim dgv As DataGridView

        wd1 = 0
        wd2 = 0

        'pesky
        Call OrderReportTableConfig()

        wd2 = dgv.RowHeadersWidth
        wd1 = wd1 + wd2
        int1 = dgv.Columns(strField).Index
        int2 = dgv.Columns(strField).DisplayIndex
        For Count1 = 0 To dgv.Columns.Count - 1
            str1 = dgv.Columns(Count1).Name
            If dgv.Columns(Count1).Visible Then
                int3 = dgv.Columns(Count1).DisplayIndex
                If int3 < int2 Then
                    wd2 = dgv.Columns.Item(Count1).Width
                    wd1 = wd1 + wd2
                End If
            End If
            'If InStr(1, str1, "Include", CompareMethod.Text) > 0 Then
            '    int1 = Count1
            '    Exit For
            'Else
            '    If dgv.Columns(Count1).Visible Then
            '        wd2 = dgv.Columns.Item(Count1).Width
            '        wd1 = wd1 + wd2
            '    End If
            'End If
        Next
        'wd2 = dgv.Columns.Item("boolRequiresSampleAssignment").Width
        'wd1 = wd1 + wd2
        'wd2 = dgv.Columns.Item("CHARHEADINGTEXT").Width
        'wd1 = wd1 + wd2
        'wd2 = dgv.Columns.Item("BOOLINCLUDE").Width
        'wd1 = wd1 + wd2
        cmd.Left = dgv.Left + wd1
        dgv.Columns.Item(int1).MinimumWidth = cmd.Width
        dgv.Columns.Item(int1).Width = cmd.Width

        frmH.cmdOrderReportTableConfig.BringToFront()

        frmH.cmdShowGroups.Left = frmH.cmdOrderReportTableConfig.Left + frmH.cmdOrderReportTableConfig.Width + 5
        frmH.cmdShowGroups.Top = frmH.cmdOrderReportTableConfig.Top

    End Sub

    Sub FillStuff()
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        'Dim rs As New ADODB.Recordset
        Dim var1, var2, var3
        Dim strPath As String
        'Dim fso As New Scripting.FileSystemObject
        'Dim fo, fi
        Dim intCt As Short
        Dim Count1 As Short
        Dim int1 As Short
        Dim tbl As System.Data.DataTable
        Dim drow As DataRow
        Dim boolExists As Boolean

        'clear all form control text boxes
        Call LoadFormClear()
        'enter path info in dgHome
        str1 = "Report Templates"
        str2 = "charConfigTitle = '" & str1 & "'"
        Dim drows1() As DataRow
        'drows1 = tblConfiguration.Select(str2)
        'strPath = drows1(0).Item("charConfigValue")

        strPath = ""

        With tbl_dgHome
            Dim dr As DataRow = .NewRow
            dr(0) = str1
            dr(1) = strPath
            .Rows.Add(dr)
        End With

        ''get .doc files from path
        ''ensure folder exists
        'boolExists = fso.FolderExists(strPath)
        'If boolExists Then 'continue
        'Else
        '    MsgBox("This folder:" & Chr(10) & Chr(10) & strPath & Chr(10) & Chr(10) & "does not exist." & Chr(10) & Chr(10) & "Please inspect the 'Report Template Path' value of the Configuration module.", MsgBoxStyle.Information, "Path does not exist...")
        '    GoTo end1
        'End If
        'fo = fso.GetFolder(strPath)
        'For Each fi In fo.files
        '    var1 = fi.Name
        '    If StrComp(Microsoft.VisualBasic.Right(var1, 4), ".doc", CompareMethod.Text) = 0 Then
        '        If InStr(var1, "StudyDoc", CompareMethod.Text) > 0 Then
        '        Else
        '            'frmh.lbxReportTemplates.Items.Add(var1)
        '            cbxxReportTemplates.Items.Add(var1)
        '        End If
        '    End If
        'Next
        'rs.Close()
        'rs = Nothing

        'bind tbl_dgHome to dgHome
        frmH.dgHome.DataSource = tbl_dgHome
        ''''debugWriteLine("7")


        'Call AutoSizeGrid(500, tbl_dgHome, frmh.dgHome, tbl_dgHome.Rows.Count, tbl_dgHome.Columns.Count, 0, False)

        'fill SubmittedTo,InSupportOf,SubmittedBy cbxboxes
        Call FillCorporateNames()
        'choose first items
        If frmH.cbxSubmittedTo.Items.Count = 0 Then
        Else
            frmH.cbxSubmittedTo.SelectedIndex = 0
        End If
        If frmH.cbxInSupportOf.Items.Count = 0 Then
        Else
            frmH.cbxInSupportOf.SelectedIndex = 0
        End If
        If frmH.cbxSubmittedBy.Items.Count = 0 Then
        Else
            frmH.cbxSubmittedBy.SelectedIndex = 0
        End If

        drows1 = Nothing

end1:

    End Sub

    Sub FillCorporateNames()
        Dim int1 As Short
        Dim var1, var2, var3
        Dim tbl As System.Data.DataTable
        Dim str1 As String
        Dim Count1 As Short
        Dim strST As String
        Dim strIS As String
        Dim strSB As String

        boolStopCBX = True

        'record the values in cbxs
        strST = NZ(frmH.cbxSubmittedTo.Text, "[None]")
        strIS = NZ(frmH.cbxInSupportOf.Text, "[None]")
        strSB = NZ(frmH.cbxSubmittedBy.Text, "[None]")

        'first clear cbxboxes
        frmH.cbxSubmittedTo.Items.Clear()
        frmH.cbxInSupportOf.Items.Clear()
        frmH.cbxSubmittedBy.Items.Clear()

        'tbl = tblCorporateAddresses
        tbl = tblCorporateNickNames
        'str1 = "boolInclude = TRUE"
        str1 = "boolinclude = -1"
        Dim dRows() As DataRow = tbl.Select(str1, "charNickname ASC")

        int1 = dRows.Length
        var3 = ""

        frmH.cbxSubmittedTo.Items.Add("[None]")
        frmH.cbxInSupportOf.Items.Add("[None]")
        frmH.cbxSubmittedBy.Items.Add("[None]")

        For Count1 = 0 To int1 - 1
            var1 = dRows(Count1).Item("charNickname")
            frmH.cbxSubmittedTo.Items.Add(var1)
            frmH.cbxInSupportOf.Items.Add(var1)
            frmH.cbxSubmittedBy.Items.Add(var1)
        Next

        'now select original values
        frmH.cbxSubmittedTo.Text = strST
        frmH.cbxInSupportOf.Text = strIS
        frmH.cbxSubmittedBy.Text = strSB

        boolStopCBX = False

        dRows = Nothing
        tbl.Dispose()


    End Sub

    Sub LoadFormClear()
        'this will clear text values upon Form Load
        Dim cont As Control
        Dim tp As TabPage
        Dim var1

        For Each cont In frmH.Controls
            var1 = cont.Name
            If InStr(var1, "char", CompareMethod.Text) > 0 Then
                cont.Text = ""
            End If
        Next

        For Each tp In frmH.tab1.TabPages
            For Each cont In tp.Controls
                var1 = cont.Name
                If InStr(var1, "char", CompareMethod.Text) > 0 Then
                    cont.Text = ""
                End If
            Next
        Next
    End Sub


    Sub UpdateOutstandingItems(ByVal dtS As DateTime, ByVal dtE As DateTime)

        Dim tblMax As System.Data.DataTable
        Dim rowsMax() As DataRow
        Dim strFMax As String
        Dim maxID
        Dim Count1 As Short
        Dim dv As System.Data.DataView
        Dim str1 As String
        Dim int1 As Short
        Dim var1

        'clear items out of tblOutstanding
        'clear contents of tblOutstandingItems
        intErrCount = 0
        str1 = "ID_TBLSTUDIES = " & id_tblStudies
        Dim drowsO() As DataRow
        drowsO = tblOutstandingItems.Select(str1)
        int1 = drowsO.Length
        For Count1 = int1 - 1 To 0 Step -1
            'tblOutstandingItems.Rows.Remove(drowsO(Count1))
            drowsO(Count1).Delete()
        Next

        maxID = GetMaxID("tblOutstandingItems", ctArrReportNA + 1, True) 'if maxid increment is 1, then getmaxid already does putmaxid
        '20190219 LEE: Don't need anymore. Used GetMaxID
        'Call PutMaxID("tblOutstandingItems", maxID + ctArrReportNA + 1)

        'tblMax = tblMaxID
        'strFMax = "charTable = 'tblOutstandingItems'"
        'Erase rowsMax
        'rowsMax = tblMax.Select(strFMax)
        'maxID = rowsMax(0).Item("nummaxid")
        'maxID = maxID
        'rowsMax(0).BeginEdit()
        'rowsMax(0).Item("nummaxid") = maxID + ctArrReportNA + 1
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

        Dim frm As New frmReportErrMsg
        For Count1 = 1 To ctArrReportNA
            var1 = NZ(arrReportNA(1, Count1), "")
            If Len(var1) = 0 Then
            Else
                maxID = maxID + 1
                Dim newrow As DataRow = tblOutstandingItems.NewRow
                newrow.BeginEdit()
                newrow("id_tblOutstandingItems") = maxID
                newrow("id_tblStudies") = id_tblStudies
                newrow("charSectionName") = arrReportNA(1, Count1) 'section name
                newrow("charTabName") = arrReportNA(3, Count1) 'tab name
                newrow("charLocation") = arrReportNA(2, Count1) 'report item
                newrow("charValue") = arrReportNA(4, Count1) 'tab item
                newrow("CHARFIELDCODE") = arrReportNA(5, Count1) 'field code
                newrow.EndEdit()
                tblOutstandingItems.Rows.Add(newrow)

            End If
        Next


        If boolGuWuOracle Then
            Try
                ta_tblOutstandingItems.Update(tblOutstandingItems)
            Catch ex As DBConcurrencyException
                'ds2005.TBLOUTSTANDINGITEMS.Merge('ds2005.TBLOUTSTANDINGITEMS, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblOutstandingItemsAcc.Update(tblOutstandingItems)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLOUTSTANDINGITEMS.Merge('ds2005Acc.TBLOUTSTANDINGITEMS, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblOutstandingItemsSQLServer.Update(tblOutstandingItems)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLOUTSTANDINGITEMS.Merge('ds2005Acc.TBLOUTSTANDINGITEMS, True)
            End Try
        End If

        frm.dgvReportErrMsg.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        frm.dgvReportErrMsg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        dv = New DataView(tblOutstandingItems)
        str1 = "id_tblStudies = " & id_tblStudies
        dv.RowFilter = str1
        frm.dgvReportErrMsg.DataSource = dv
        'configure dgv
        frm.dgvReportErrMsg.Columns.Item("ID_TBLOUTSTANDINGITEMS").Visible = False
        frm.dgvReportErrMsg.Columns.Item("id_tblStudies").Visible = False
        frm.dgvReportErrMsg.Columns.Item("UPSIZE_TS").Visible = False

        frm.dgvReportErrMsg.Columns.Item("charSectionName").Visible = False
        frm.dgvReportErrMsg.Columns.Item("charSectionName").HeaderText = "Report Section Name"
        frm.dgvReportErrMsg.Columns.Item("charTabName").HeaderText = "StudyDoc Tab Name"
        frm.dgvReportErrMsg.Columns.Item("charLocation").HeaderText = "Report Item"
        frm.dgvReportErrMsg.Columns.Item("charValue").HeaderText = "Item Within Tab"
        frm.dgvReportErrMsg.Columns.Item("CHARFIELDCODE").HeaderText = "Field Code"

        For Count1 = 0 To frm.dgvReportErrMsg.ColumnCount - 1
            frm.dgvReportErrMsg.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
            frm.dgvReportErrMsg.Columns.Item(Count1).DisplayIndex = frm.dgvReportErrMsg.ColumnCount - 1
        Next
        frm.dgvReportErrMsg.Columns.Item("charSectionName").DisplayIndex = 0 'HeaderText = "Report Section Name"
        frm.dgvReportErrMsg.Columns.Item("charLocation").DisplayIndex = 1 '.HeaderText = "Report Item"
        frm.dgvReportErrMsg.Columns.Item("charTabName").DisplayIndex = 2 '.HeaderText = "StudyDoc Tab Name"
        frm.dgvReportErrMsg.Columns.Item("charValue").DisplayIndex = 3 '.HeaderText = "Item Within Tab"


        frm.dgvReportErrMsg.AutoResizeColumns()
        frm.dgvReportErrMsg.AutoResizeRows()

        frm.dgvReportErrMsg.Columns.Item("charSectionName").DisplayIndex = 0 'HeaderText = "Report Section Name"
        frm.dgvReportErrMsg.Columns.Item("charLocation").DisplayIndex = 1 '.HeaderText = "Report Item"
        frm.dgvReportErrMsg.Columns.Item("charTabName").DisplayIndex = 2 '.HeaderText = "StudyDoc Tab Name"
        frm.dgvReportErrMsg.Columns.Item("charValue").DisplayIndex = 3 '.HeaderText = "Item Within Tab"

        frm.txtReportTitle.Text = frmH.lblReportTitle.Text

        str1 = "Report Started:  " & Format(dtS, "MMMM dd, yyyy   hh:mm:ss tt")
        frm.lblStart.Text = str1

        str1 = "Report Finished:  " & Format(dtE, "MMMM dd, yyyy   hh:mm:ss tt")
        frm.lblEnd.Text = str1

        frm.lblEnd.Left = frm.lblStart.Left + frm.lblStart.Width + 20

        Dim intMin As Long
        intMin = DateDiff(DateInterval.Minute, dtS, dtE)
        If intMin = 0 Then
            intMin = 1
        End If

        If intMin = 1 Then
            str1 = "Total Time:  ~ " & intMin & " minute"
        Else
            str1 = "Total Time:  ~ " & intMin & " minutes"
        End If

        frm.lblTotal.Text = str1

        frm.lblEnd.Visible = True
        frm.lblStart.Visible = True
        frm.lblTotal.Visible = True

        If ctArrReportNA = 0 Then

            str1 = "There were no field code anomolies found in this report." & ChrW(10) & ChrW(10)
            str1 = str1 & "This report has all field codes satisfied."
            frm.lblTitle.Text = str1

            frm.lbl2.Visible = False


        End If

        frm.ShowDialog()
        frm.Dispose()
        frmH.Refresh()
        'SendKeys.Send("%")


        'configure dgv
        'frm.dgvReportErrMsg.Columns.item(0).Visible = False
        'frm.dgvReportErrMsg.Columns.item("charSectionName").Visible = False
        'frm.dgvReportErrMsg.Columns.item("charSectionName").HeaderText = "Report Section Name"
        'frm.dgvReportErrMsg.Columns.item("charTabName").HeaderText = "StudyDoc Tab Name"
        'frm.dgvReportErrMsg.Columns.item("charLocation").HeaderText = "Location Within Tab"
        'frm.dgvReportErrMsg.Columns.item("charValue").HeaderText = "StudyDoc Value Item"



    End Sub


    Sub AddColumnsAnalRefTable()

        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim dtbl As System.Data.DataTable
        Dim boolRO As Boolean
        Dim int1 As Short
        Dim var1, var2
        Dim dv As System.Data.DataView
        Dim tblS As System.Data.DataTable
        Dim strF As String
        Dim strS As String
        Dim drows() As DataRow
        Dim ct1 As Short
        Dim ct2 As Short
        Dim boolC As Boolean
        Dim str1 As String
        Dim str2 As String
        Dim strAnal As String
        Dim ct3 As Short
        Dim intCols As Short
        Dim dgv As DataGridView
        Dim strIsIS As String

        Dim dtblAH As DataTable
        Dim dtblUA As DataTable
        Dim dtblUIS As DataTable

        'this table should only be unique compounds, not C1 or matrix

        Dim dvAHA As DataView
        Dim dvAHIS As DataView

        'Dim arrAnalytes(15, 51) '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
        '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
        '10=UseIntStd, 11=IntStd, 12=MasterAssayID, 13=IsCoadminCmpd,14=OriginalAnalyteDescription,15=intGroup

        dtblAH = tblAnalytesHome ' tblAnalyteGroups
        strF = "IsIntStd = 'No'"
        strS = ""
        dvAHA = New DataView(dtblAH, strF, strS, DataViewRowState.CurrentRows)
        dtblUA = dvAHA.ToTable("a", True, "OriginalAnalyteDescription")
        'dtblUA = dvAHA.ToTable("a", True, "ANALYTEID")
        strF = "IsIntStd = 'Yes'"
        strS = ""
        dvAHIS = New DataView(dtblAH, strF, strS, DataViewRowState.CurrentRows)
        dtblUIS = dvAHA.ToTable("b", True, "INTSTD")

        'If boolUseGroups Then
        '    dtblAH = tblAnalyteGroups
        '    dvAHA = New DataView(dtblAH, "", "", DataViewRowState.CurrentRows)
        '    dtblUA = dvAHA.ToTable("a", True, "ANALYTEDESCRIPTION")
        '    dvAHIS = New DataView(dtblAH, "", "", DataViewRowState.CurrentRows)
        '    dtblUIS = dvAHA.ToTable("b", True, "INTSTD")
        'End If

        Try

            boolFromCAR = True
            tblS = tblAnalRefStandards
            If boolUseGroups Then
                strF = "id_tblStudies = " & id_tblStudies & " AND INTGROUP > 0"
            Else
                strF = "id_tblStudies = " & id_tblStudies & " AND INTGROUP = 0"
            End If

            drows = tblS.Select(strF, "ID_TBLANALREFSTANDARDS ASC")
            ct1 = drows.Length

            If ct1 = 0 Then
                If boolUseGroups Then
                    ct2 = dtblUA.Rows.Count ' + dtblUIS.Rows.Count
                Else
                    ct2 = ctAnalytes + ctAnalytes_IS
                End If
            Else
                ct2 = ct1
            End If

            'dg = frmh.dgCompanyAnalRef
            dtbl = tblCompanyAnalRefTable
            boolRO = False

            dgv = frmH.dgvCompanyAnalRef
            Dim boolF As Boolean
            'boolF = dgv.Columns.item("Item").Frozen
            'dgv.Columns.item("Item").Frozen = False 'must do this first before fooling with table or VStudio has a fit

            int1 = dgv.Columns.Count 'for debugging
            boolF = dgv.Columns.Item("Item").Frozen
            dgv.Columns.Item("Item").Frozen = False 'must do this first before fooling with table or VStudio has a fit

            'delete all columns except col 0,1,2
            int1 = dtbl.Columns.Count
            For Count2 = int1 - 1 To 0 Step -1
                str1 = dtbl.Columns.Item(Count2).ColumnName
                If StrComp(str1, "ID_TBLDATATABLEROWTITLES", CompareMethod.Text) = 0 Then
                ElseIf StrComp(str1, "BOOLINCLUDE", CompareMethod.Text) = 0 Then
                ElseIf StrComp(str1, "Item", CompareMethod.Text) = 0 Then
                Else
                    dtbl.Columns.Remove(dtbl.Columns.Item(Count2))
                End If
            Next

            'first do Analytes

            Dim dtbl1 As DataTable
            For Count3 = 1 To 2

                Select Case Count3
                    Case 1
                        dtbl1 = dtblUA
                        strIsIS = "No"
                    Case 2
                        dtbl1 = dtblUIS
                        strIsIS = "Yes"
                End Select

                For Count2 = 1 To dtbl1.Rows.Count

                    Select Case Count3
                        Case 1
                            strAnal = dtbl1.Rows(Count2 - 1).Item("OriginalAnalyteDescription")
                        Case 2
                            strAnal = dtbl1.Rows(Count2 - 1).Item("INTSTD")
                    End Select

                    Dim dc As New DataColumn
                    dc.DataType = System.Type.GetType("System.String")
                    If ct1 = 0 Then
                        'If boolUseGroups Then
                        '    var1 = strAnal ' dtbl1.Rows(Count2 - 1).Item("OriginalAnalyteDescription")
                        '    var2 = var1
                        'Else
                        '    var1 = arrAnalytes(1, Count2)
                        '    var2 = var1
                        'End If
                        var1 = strAnal
                        var2 = var1

                    Else
                        var1 = drows(Count2 - 1).Item("CHARCOLUMNNAME")
                        var2 = NZ(drows(Count2 - 1).Item("charAnalyteParent"), var1)
                        Count3 = drows(Count2 - 1).Item("ID_TBLANALREFSTANDARDS")
                    End If

                    If dtbl.Columns.Contains(var1) Then
                    Else
                        dc.ColumnName = var1
                        dc.Caption = var2

                        dc.ReadOnly = boolRO
                        'On Error Resume Next
                        dtbl.Columns.Add(dc)
                        'On Error GoTo 0

                        'now check for default data
                        If ct1 = 0 Then 'enter default data
                            dv = dtbl.DefaultView
                            str1 = "Is Replicate?"
                            int1 = FindRowDVByCol(str1, dv, "Item")
                            dtbl.Rows.Item(int1).BeginEdit()
                            dtbl.Rows.Item(int1).Item(var1) = "No"
                            dtbl.Rows.Item(int1).EndEdit()

                            str1 = "Is Configured in Watson?"
                            int1 = FindRowDVByCol(str1, dv, "Item")
                            dtbl.Rows.Item(int1).BeginEdit()
                            dtbl.Rows.Item(int1).Item(var1) = "Yes"
                            dtbl.Rows.Item(int1).EndEdit()

                            str1 = "Analyte Name"
                            int1 = FindRowDVByCol(str1, dv, "Item")
                            dtbl.Rows.Item(int1).BeginEdit()
                            dtbl.Rows.Item(int1).Item(var1) = strAnal ' dtbl.Rows(Count2).Item("OriginalAnalyteDescription") ' arrAnalytes(14, Count2)
                            dtbl.Rows.Item(int1).EndEdit()

                            str1 = "Is Coadministered Cmpd?"
                            int1 = FindRowDVByCol(str1, dv, "Item")
                            dtbl.Rows.Item(int1).BeginEdit()
                            dtbl.Rows.Item(int1).Item(var1) = "No"
                            dtbl.Rows.Item(int1).EndEdit()

                            str1 = "Analyte Parent"
                            int1 = FindRowDVByCol(str1, dv, "Item")
                            dtbl.Rows.Item(int1).BeginEdit()
                            dtbl.Rows.Item(int1).Item(var1) = strAnal ' dtbl.Rows(Count2).Item("OriginalAnalyteDescription") 'arrAnalytes(14, Count2)
                            dtbl.Rows.Item(int1).EndEdit()

                            str1 = "Is Internal Standard?"
                            int1 = FindRowDVByCol(str1, dv, "Item")
                            dtbl.Rows.Item(int1).BeginEdit()
                            dtbl.Rows.Item(int1).Item(var1) = strIsIS ' arrAnalytes(9, Count2)
                            dtbl.Rows.Item(int1).EndEdit()

                        End If

                    End If

                Next

            Next



            Dim dv1 As System.Data.DataView = New DataView(dtbl)
            'dg.DataSource = dv
            frmH.dgvCompanyAnalRef.DataSource = dv1
            frmH.dgvCompanyAnalRef.Refresh()
            For Count1 = 0 To frmH.dgvCompanyAnalRef.Columns.Count - 1
                frmH.dgvCompanyAnalRef.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
            Next

            Call HideAnalRefRows()

            Dim bool As Boolean
            int1 = dgv.Rows.Count

            For Count1 = 0 To int1 - 1
                var1 = dgv("ID_TBLDATATABLEROWTITLES", Count1).Value
                bool = False
                Select Case var1
                    Case Is = 205 'Is Internal Standard?
                        bool = True
                End Select
                dgv.Rows.Item(Count1).ReadOnly = bool
            Next

            boolFromCAR = False

            str1 = AnalRefHook()
            If Len(str1) > 0 Then
                'now look if row is companyid and hook is there
                Select Case str1
                    Case "CRLWor_AnalRefStandard"
                        Call ComboBoxCRLAnalRefFill()
                        Call PopulateFromCRLAnalRefHook()
                End Select
            End If

            Call ResizeRows(frmH.dgvCompanyAnalRef)
            Call ResizeRows(frmH.dgvWatsonAnalRef)

            dgv.Columns.Item("Item").Frozen = boolF

        Catch ex As Exception
            var1 = ex.Message
        End Try

    End Sub

    Sub ComboBoxCRLAnalRefFill()
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim tbl3 As System.Data.DataTable
        Dim rows3() As DataRow
        Dim dgv As DataGridView
        Dim str1 As String
        Dim strF As String
        Dim Count1 As Short
        Dim ct2 As Short

        ct2 = frmH.dgvCompanyAnalRef.Columns.Count
        dgv = frmH.dgvCompanyAnalRef
        str1 = frmH.lbxTab1.SelectedItem
        strF = "CHARHOOK = 'CRLWor_AnalRefStandard'"
        tbl = tblHooks
        rows = tbl.Select(strF)
        If rows.Length = 0 Or boolHook1 = False Then 'ignore everything
        Else
            Try
                Dim tbl1 As System.Data.DataTable
                tbl1 = tblHook1
                Dim dv2 As System.Data.DataView = New DataView(tbl1)
                Dim tbl2 As System.Data.DataTable = dv2.ToTable("a", True, "BottleID")
                For Count1 = 0 To ct2 - 1
                    str1 = frmH.dgvCompanyAnalRef.Columns.Item(Count1).Name
                    If StrComp(str1, "ID_TBLDATATABLEROWTITLES", CompareMethod.Text) = 0 Then 'ignore
                    ElseIf StrComp(str1, "BOOLINCLUDE", CompareMethod.Text) = 0 Then 'ignore
                    ElseIf StrComp(str1, "Item", CompareMethod.Text) = 0 Then 'ignore
                    Else
                        Dim cbx As New DataGridViewComboBoxCell
                        cbx = cbxCompanyID.Clone
                        cbx.AutoComplete = True
                        cbx.MaxDropDownItems = 20
                        cbx.DisplayStyleForCurrentCellOnly = True
                        cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
                        dgv(Count1, 0) = cbx
                    End If
                Next
            Catch ex As Exception

            End Try

        End If

    End Sub


    Sub PopulateFromCRLAnalRefHook()
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim tbl1 As System.Data.DataTable
        Dim tbl2 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim rows2a() As DataRow
        Dim dgv As DataGridView
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim var1, var2
        Dim strF As String
        Dim Count1 As Short
        Dim Count2 As Short
        Dim ct2 As Short
        Dim strID As String
        Dim int1 As Short
        Dim int2 As Short
        Dim dv As System.Data.DataView
        Dim intRows As Short
        Dim intCols As Short
        Dim strT As String
        Dim intCol As Short
        Dim strCol As String
        Dim intID As Long


        dgv = frmH.dgvCompanyAnalRef
        dv = dgv.DataSource
        intRows = dv.Count
        intCols = dgv.Columns.Count
        tbl = tblHooks
        tbl1 = tblHook1
        tbl2 = tblDataTableRowTitles

        strF = "CHARHOOK = 'CRLWor_AnalRefStandard'"
        rows = tbl.Select(strF)
        If rows.Length = 0 Then 'ignore everything
        Else
            If rows(0).Item("BOOLINCLUDE") = -1 Then
                For Count1 = 0 To intCols - 1
                    strCol = dgv.Columns.Item(Count1).Name
                    If StrComp(strCol, "ID_TBLDATATABLEROWTITLES", CompareMethod.Text) = 0 Then 'IGNORE
                    ElseIf StrComp(strCol, "BOOLINCLUDE", CompareMethod.Text) = 0 Then 'IGNORE
                    ElseIf StrComp(strCol, "Item", CompareMethod.Text) = 0 Then 'IGNORE
                    Else
                        intCol = Count1
                        strID = NZ(dv(0).Item(Count1), "") 'retrieve BottleID
                        If Len(strID) = 0 Then
                        Else
                            strF = "BottleID = '" & strID & "'"
                            rows1 = tbl1.Select(strF)
                            intID = dv(0).Item("ID_TBLDATATABLEROWTITLES") 'id_tbldatatablerowtitles
                            'determine if there is a hook for this
                            strF = "ID_TBLDATATABLEROWTITLES = " & intID
                            Erase rows2a
                            rows2a = tbl2.Select(strF)
                            var1 = NZ(rows2a(0).Item("CHARHOOK"), "")
                            If Len(var1) = 0 Then 'ignore
                            Else
                                If rows1.Length = 0 Then 'ignore
                                Else
                                    For Count2 = 1 To intRows - 1
                                        str1 = dv(Count2).Item("Item") 'row title
                                        If StrComp(str1, "Analyte Name", CompareMethod.Text) = 0 Then
                                            var2 = var1
                                        End If
                                        If StrComp(str1, "Physical Description", CompareMethod.Text) = 0 Then
                                            strT = ""
                                            'retrieve PhyColor
                                            var1 = NZ(rows1(0).Item("PhyColor"), "")
                                            If Len(var1) = 0 Then
                                            Else
                                                strT = strT & var1 & " "
                                            End If
                                            'retrieve PhyState
                                            var1 = NZ(rows1(0).Item("PhyState"), "")
                                            If Len(var1) = 0 Then
                                            Else
                                                strT = strT & var1 & " "
                                            End If

                                            dv(Count2).BeginEdit()
                                            dv(Count2).Item(intCol) = strT
                                            dv(Count2).EndEdit()

                                        ElseIf StrComp(str1, "Amount Received", CompareMethod.Text) = 0 Then
                                            strT = ""
                                            'retrieve PhyColor
                                            var1 = NZ(rows1(0).Item("Amt"), "")
                                            If Len(var1) = 0 Then
                                            Else
                                                strT = strT & var1 & " "
                                            End If
                                            'retrieve PhyState
                                            var1 = NZ(rows1(0).Item("btlUnits"), "")
                                            If Len(var1) = 0 Then
                                            Else
                                                strT = strT & var1 & " "
                                            End If

                                            dv(Count2).BeginEdit()
                                            dv(Count2).Item(intCol) = strT
                                            dv(Count2).EndEdit()

                                        ElseIf StrComp(str1, "Purity", CompareMethod.Text) = 0 Then
                                            strT = ""
                                            'retrieve Purity
                                            var1 = NZ(rows1(0).Item("% Purity"), "")
                                            If Len(var1) = 0 Then
                                            Else
                                                strT = strT & var1 & "%"
                                            End If

                                            dv(Count2).BeginEdit()
                                            dv(Count2).Item(intCol) = strT
                                            dv(Count2).EndEdit()

                                        ElseIf StrComp(str1, "Date Received", CompareMethod.Text) = 0 Then
                                            strT = ""
                                            'retrieve Date
                                            var1 = NZ(rows1(0).Item("ReceiptDate"), "")
                                            If Len(var1) = 0 Then
                                            Else
                                                'format date
                                                strT = Format(var1, LDateFormat)
                                            End If

                                            dv(Count2).BeginEdit()
                                            dv(Count2).Item(intCol) = strT
                                            dv(Count2).EndEdit()

                                        ElseIf StrComp(str1, "Expiration/Retest Date", CompareMethod.Text) = 0 Then
                                            strT = ""
                                            'retrieve Date
                                            var1 = NZ(rows1(0).Item("ExpDate"), "")
                                            If Len(var1) = 0 Then
                                            Else
                                                'format date
                                                strT = Format(var1, LDateFormat)
                                            End If

                                            dv(Count2).BeginEdit()
                                            dv(Count2).Item(intCol) = strT
                                            dv(Count2).EndEdit()

                                        Else
                                            strF = "CHARDATATABLENAME = 'tblCompanyAnalRefTable' AND CHARROWNAME = '" & str1 & "'"
                                            Erase rows2
                                            rows2 = tbl2.Select(strF)
                                            If rows2.Length = 0 Then 'ignore
                                            Else
                                                var2 = NZ(rows2(0).Item("CHARHOOK"), "")
                                                If Len(var2) = 0 Then
                                                Else
                                                    var1 = rows1(0).Item(var2)
                                                    dv(Count2).BeginEdit()
                                                    dv(Count2).Item(intCol) = NZ(var1, "")
                                                    dv(Count2).EndEdit()
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        End If
                    End If
                Next
            End If

        End If


    End Sub

    Sub PopulateSingleFromCRLAnalRefHook(ByVal intCol)
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim tbl1 As System.Data.DataTable
        Dim tbl2 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim rows2a() As DataRow
        Dim dgv As DataGridView
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim var1, var2
        Dim strF As String
        Dim Count1 As Short
        Dim Count2 As Short
        Dim ct2 As Short
        Dim strID As String
        Dim int1 As Short
        Dim int2 As Short
        Dim dv As System.Data.DataView
        Dim intRows As Short
        Dim intCols As Short
        Dim strT As String
        Dim intID As Long



        dgv = frmH.dgvCompanyAnalRef
        dv = dgv.DataSource
        intRows = dv.Count
        intCols = dgv.Columns.Count
        tbl = tblHooks
        tbl1 = tblHook1
        tbl2 = tblDataTableRowTitles
        strF = "CHARHOOK = 'CRLWor_AnalRefStandard'"
        rows = tbl.Select(strF)
        If rows.Length = 0 Then 'ignore everything
        Else
            strID = NZ(dv(0).Item(intCol), "") 'retrieve BottleID
            If Len(strID) = 0 Then
            Else
                strF = "BottleID = '" & strID & "'"
                rows1 = tbl1.Select(strF)
                If rows1.Length = 0 Then 'ignore
                Else
                    For Count2 = 1 To intRows - 1
                        str1 = dv(Count2).Item("Item") 'row title
                        intID = dv(Count2).Item("ID_TBLDATATABLEROWTITLES") 'id_tbldatatablerowtitles
                        'determine if there is a hook for this

                        strF = "ID_TBLDATATABLEROWTITLES = " & intID
                        Erase rows2a
                        rows2a = tbl2.Select(strF)
                        var1 = NZ(rows2a(0).Item("CHARHOOK"), "")
                        If Len(var1) = 0 Then 'ignore
                        Else
                            If StrComp(str1, "Analyte Name", CompareMethod.Text) = 0 Then
                                var2 = var1
                            End If
                            If StrComp(str1, "Physical Description", CompareMethod.Text) = 0 Then
                                strT = ""
                                'retrieve PhyColor
                                var1 = NZ(rows1(0).Item("PhyColor"), "")
                                If Len(var1) = 0 Then
                                Else
                                    strT = strT & var1 & " "
                                End If
                                'retrieve PhyState
                                var1 = NZ(rows1(0).Item("PhyState"), "")
                                If Len(var1) = 0 Then
                                Else
                                    strT = strT & var1 & " "
                                End If

                                'dv(Count2).BeginEdit()
                                'dv(Count2).Item(intCol) = strT
                                'dv(Count2).EndEdit()

                            ElseIf StrComp(str1, "Amount Received", CompareMethod.Text) = 0 Then
                                strT = ""
                                'retrieve PhyColor
                                var1 = NZ(rows1(0).Item("Amt"), "")
                                If Len(var1) = 0 Then
                                Else
                                    strT = strT & var1 & " "
                                End If
                                'retrieve PhyState
                                var1 = NZ(rows1(0).Item("btlUnits"), "")
                                If Len(var1) = 0 Then
                                Else
                                    strT = strT & var1 & " "
                                End If

                                'dv(Count2).BeginEdit()
                                'dv(Count2).Item(intCol) = strT
                                'dv(Count2).EndEdit()

                            ElseIf StrComp(str1, "Purity", CompareMethod.Text) = 0 Then
                                strT = ""
                                'retrieve Purity
                                var1 = NZ(rows1(0).Item("% Purity"), "")
                                If Len(var1) = 0 Then
                                Else
                                    strT = strT & var1 & "%"
                                End If

                                'dv(Count2).BeginEdit()
                                'dv(Count2).Item(intCol) = strT
                                'dv(Count2).EndEdit()

                            ElseIf StrComp(str1, "Date Received", CompareMethod.Text) = 0 Then
                                strT = ""
                                'retrieve Date
                                var1 = NZ(rows1(0).Item("ReceiptDate"), "")
                                If Len(var1) = 0 Then
                                Else
                                    'format date
                                    strT = Format(var1, LDateFormat)
                                End If

                                'dv(Count2).BeginEdit()
                                'dv(Count2).Item(intCol) = strT
                                'dv(Count2).EndEdit()

                            ElseIf StrComp(str1, "Expiration/Retest Date", CompareMethod.Text) = 0 Then
                                strT = ""
                                'retrieve Date
                                var1 = NZ(rows1(0).Item("ExpDate"), "")
                                If Len(var1) = 0 Then
                                Else
                                    'format date
                                    strT = Format(var1, LDateFormat)
                                End If

                                'dv(Count2).BeginEdit()
                                'dv(Count2).Item(intCol) = strT
                                'dv(Count2).EndEdit()

                            Else
                                strT = ""
                                strF = "CHARDATATABLENAME = 'tblCompanyAnalRefTable' AND CHARROWNAME = '" & str1 & "'"
                                Erase rows2
                                rows2 = tbl2.Select(strF)
                                If rows2.Length = 0 Then 'ignore
                                Else
                                    var2 = NZ(rows2(0).Item("CHARHOOK"), "")
                                    If Len(var2) = 0 Then
                                    Else
                                        var1 = rows1(0).Item(var2)
                                        strT = CStr(NZ(var1, ""))
                                        'dv(Count2).BeginEdit()
                                        'dv(Count2).Item(intCol) = NZ(var1, "")
                                        'dv(Count2).EndEdit()
                                    End If
                                End If
                            End If

                            If Len(strT) > 0 Then
                                var1 = strT 'for debugging
                            End If

                            dv(Count2).BeginEdit()
                            dv(Count2).Item(intCol) = strT
                            dv(Count2).EndEdit()

                        End If
                    Next
                End If
            End If
        End If

        dgv.Refresh()

    End Sub

    Sub AddColumnsWatsonAnalRefTable()

        Dim Count1 As Short
        Dim Count2 As Short
        Dim dg As DataGrid
        Dim dtbl As System.Data.DataTable
        Dim boolRO As Boolean
        Dim int1 As Short
        Dim int2 As Short
        Dim var1, var2
        Dim dv As System.Data.DataView
        Dim incr As Short
        Dim str1 As String
        Dim str2 As String

        'dg = frmh.dgWatsonAnalRef
        dtbl = tblWatsonAnalRefTable
        boolRO = True

        'delete all columns except col(0)
        int1 = dtbl.Columns.Count
        For Count2 = int1 - 1 To 1 Step -1
            dtbl.Columns.Remove(dtbl.Columns.Item(Count2))
        Next
        'dg.TableStyles.Clear()
        int2 = 0
        incr = 0
        Try
            For Count2 = 1 To ctAnalytes + ctAnalytes_IS
                var1 = arrAnalytes(1, Count2)
                If IsDBNull(var1) Then
                    var2 = var2 'debug
                End If
                'If dtbl.Columns.Contains(var1) Then
                '    incr = incr + 1
                '    str1 = arrAnalytes(1, Count2)
                '    str2 = str1 & "_Cal" & incr
                '    arrAnalytes(1, Count2) = str2
                'Else
                'End If
                int2 = int2 + 1
                Dim dc As New DataColumn
                dc.DataType = System.Type.GetType("System.String")

                dc.ColumnName = arrAnalytes(1, Count2)
                dc.Caption = arrAnalytes(1, Count2)
                dc.ReadOnly = boolRO
                dtbl.Columns.Add(dc)
            Next
        Catch ex As Exception
            var1 = ex.Message
        End Try

        'dv = dtbl.DefaultView
        dv = New DataView(dtbl)
        frmH.dgvWatsonAnalRef.DataSource = dv


    End Sub

    Sub FillAnalysisResultsTable(ByVal cn)

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim rs As New ADODB.Recordset
        Dim fld As ADODB.Field
        Dim drow As DataRow
        Dim var1, var2, var3, var4
        Dim row1() As DataRow
        Dim row2() As DataRow
        Dim strF As String
        Dim intRows1 As Int64
        Dim intRows2 As Int64
        Dim Count1 As Int64
        Dim Count2 As Int64
        Dim Count3 As Int64
        Dim Count4 As Int64
        Dim tbl1 As System.Data.DataTable
        Dim tbl2 As System.Data.DataTable
        Dim tbl3 As System.Data.DataTable
        Dim tbl4 As System.Data.DataTable
        Dim int1 As Int64
        Dim int2 As Int64
        Dim int3 As Int64
        Dim int4 As Int64
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim rows3() As DataRow
        Dim rows4() As DataRow
        Dim strS As String
        Dim dcol As DataColumn
        Dim int10 As Int64

        Dim boolMDB As Boolean = True

        tbl4 = tblAssignedSamples
        int4 = tbl4.Rows.Count 'DEBUG

        If boolANSI Then
            str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, STUDY.STUDYNAME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME "
            str2 = "FROM (ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) INNER JOIN STUDY ON ANALYTICALRUNANALYTES.STUDYID = STUDY.STUDYID "
            'str3 = "WHERE(((ASSAY.STUDYID)=" & wStudyID & ") And ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) <> 4)) "
            str3 = "WHERE(((ASSAY.STUDYID)=" & wStudyID & ")) "
            str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"
        Else

        End If


        'add sampletypeid (matrix)
        If boolANSI Then
            str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, STUDY.STUDYNAME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, CONFIGSAMPLETYPES.SAMPLETYPEKEY, CONFIGSAMPLETYPES.SAMPLETYPEID, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME "
            str2 = "FROM (((ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID)) ON (ASSAY.STUDYID = ANALYTICALRUN.STUDYID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.RUNID = ANALYTICALRUN.RUNID)) ON (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID)) INNER JOIN STUDY ON ANALYTICALRUNANALYTES.STUDYID = STUDY.STUDYID) LEFT JOIN DESIGNSAMPLE ON ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) LEFT JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
            str3 = "WHERE(((ASSAY.STUDYID) = " & wStudyID & ")) "
            str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"
        Else

        End If


        'add analyteid and ANALYTICALRUNSAMPLE.ASSAYDATETIME
        If boolANSI Then
            str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ANALYTICALRUNSAMPLE.STUDYID, STUDY.STUDYNAME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, CONFIGSAMPLETYPES.SAMPLETYPEKEY, CONFIGSAMPLETYPES.SAMPLETYPEID, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME "
            str2 = "FROM ASSAYANALYTES INNER JOIN ((((ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) INNER JOIN STUDY ON ANALYTICALRUNANALYTES.STUDYID = STUDY.STUDYID) LEFT JOIN DESIGNSAMPLE ON ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) LEFT JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ASSAYANALYTES.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAYANALYTES.STUDYID = ANALYTICALRUN.STUDYID) "
            str3 = "WHERE(((ASSAY.STUDYID) = " & wStudyID & ")) "
            str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"


            'add RUNSAMPLEORDERNUMBER
            str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ANALYTICALRUNSAMPLE.STUDYID, STUDY.STUDYNAME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, CONFIGSAMPLETYPES.SAMPLETYPEKEY, CONFIGSAMPLETYPES.SAMPLETYPEID, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME "
            str2 = "FROM ASSAYANALYTES INNER JOIN ((((ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) INNER JOIN STUDY ON ANALYTICALRUNANALYTES.STUDYID = STUDY.STUDYID) LEFT JOIN DESIGNSAMPLE ON ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) LEFT JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ASSAYANALYTES.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAYANALYTES.STUDYID = ANALYTICALRUN.STUDYID) "
            str3 = "WHERE(((ASSAY.STUDYID) = " & wStudyID & ")) "
            str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"


            'add , ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS
            str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ANALYTICALRUNSAMPLE.STUDYID, STUDY.STUDYNAME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, CONFIGSAMPLETYPES.SAMPLETYPEKEY, CONFIGSAMPLETYPES.SAMPLETYPEID, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS "
            str2 = "FROM ASSAYANALYTES INNER JOIN ((((ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) INNER JOIN STUDY ON ANALYTICALRUNANALYTES.STUDYID = STUDY.STUDYID) LEFT JOIN DESIGNSAMPLE ON ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) LEFT JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ASSAYANALYTES.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAYANALYTES.STUDYID = ANALYTICALRUN.STUDYID) "
            str3 = "WHERE(((ASSAY.STUDYID) = " & wStudyID & ")) "
            str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"


            'add REPORTEDCONC

            If boolAccess Then
                str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNSAMPLE.STUDYID, STUDY.STUDYNAME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, CONFIGSAMPLETYPES.SAMPLETYPEKEY, CONFIGSAMPLETYPES.SAMPLETYPEID, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, (ANARUNANALYTERESULTS.CONCENTRATION/ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC "
                str2 = "FROM ASSAYANALYTES INNER JOIN ((((ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID)) ON (ASSAY.STUDYID = ANALYTICALRUN.STUDYID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.RUNID = ANALYTICALRUN.RUNID)) ON (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID)) INNER JOIN STUDY ON ANALYTICALRUNANALYTES.STUDYID = STUDY.STUDYID) LEFT JOIN DESIGNSAMPLE ON ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) LEFT JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ASSAYANALYTES.STUDYID = ANALYTICALRUN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ANALYTICALRUN.ASSAYID) "
                str3 = "WHERE(((ASSAY.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"
            Else
                str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".STUDY.STUDYNAME, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & "." & strAnaRunPeak & ".ANALYTEHEIGHT, " & strSchema & "." & strAnaRunPeak & ".ANALYTEAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTDNAME, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & "." & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, (" & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION/" & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC "
                str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN ((((" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & "." & strAnaRunPeak & " INNER JOIN (" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID)) ON (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) ON (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) INNER JOIN " & strSchema & ".STUDY ON " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".STUDY.STUDYID) LEFT JOIN " & strSchema & ".DESIGNSAMPLE ON " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) LEFT JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".DESIGNSAMPLE.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = ANALYTICALRUN.ASSAYID) "
                str3 = "WHERE(((" & strSchema & ".ASSAY.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"
            End If

            strSQL = str1 & str2 & str3 & str4


            'str1 = "SELECT SECUSERACCOUNTS.* "
            'str2 = "FROM SECUSERACCOUNTS;"
            'strSQL = str1 & str2
            'rs.CursorLocation = CursorLocationEnum.adUseClient
            'rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            'rs.ActiveConnection = Nothing

            If boolAccess Then
                str1 = "SELECT SECUSERACCOUNTS.* "
                str2 = "FROM SECUSERACCOUNTS;"
            Else
                str1 = "SELECT " & strSchema & ".SECUSERACCOUNTS.* "
                str2 = "FROM " & strSchema & ".SECUSERACCOUNTS;"
            End If
            strSQL = str1 & str2

            rs.CursorLocation = CursorLocationEnum.adUseClient
            rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            rs.ActiveConnection = Nothing

            If rs.RecordCount = 0 Then
                boolMDB = True
            Else
                boolMDB = False
            End If

            rs.Close()

            If boolAccess Then
                'add RECORDTIMESTAMP
                str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNSAMPLE.STUDYID, STUDY.STUDYNAME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, CONFIGSAMPLETYPES.SAMPLETYPEKEY, CONFIGSAMPLETYPES.SAMPLETYPEID, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, (ANARUNANALYTERESULTS.CONCENTRATION/ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC, ANARUNANALYTERESULTS.RECORDTIMESTAMP "
                str2 = "FROM ASSAYANALYTES INNER JOIN ((((ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID)) ON (ASSAY.STUDYID = ANALYTICALRUN.STUDYID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.RUNID = ANALYTICALRUN.RUNID)) ON (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID)) INNER JOIN STUDY ON ANALYTICALRUNANALYTES.STUDYID = STUDY.STUDYID) LEFT JOIN DESIGNSAMPLE ON ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) LEFT JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ASSAYANALYTES.STUDYID = ANALYTICALRUN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ANALYTICALRUN.ASSAYID) "
                str3 = "WHERE(((ASSAY.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"


                '20160224 LEE: re-arranged joins to get complete CONFIGSAMPLETYPES.SAMPLETYPEID (matrix)
                str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNSAMPLE.STUDYID, STUDY.STUDYNAME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, CONFIGSAMPLETYPES.SAMPLETYPEKEY, CONFIGSAMPLETYPES.SAMPLETYPEID, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, (ANARUNANALYTERESULTS.CONCENTRATION/ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC, ANARUNANALYTERESULTS.RECORDTIMESTAMP "
                str2 = "FROM ((((ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID)) ON (ASSAY.STUDYID = ANALYTICALRUN.STUDYID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.RUNID = ANALYTICALRUN.RUNID)) ON (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID)) INNER JOIN STUDY ON ANALYTICALRUNANALYTES.STUDYID = STUDY.STUDYID) LEFT JOIN DESIGNSAMPLE ON ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) INNER JOIN ASSAYANALYTES ON (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX)) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                str3 = "WHERE(((ANALYTICALRUNSAMPLE.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"

                '20160225 LEE: had to re-optimize: previous query would non open in design mode for Ricerca 032852, but can open in Frontage
                'str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNSAMPLE.STUDYID, STUDY.STUDYNAME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, CONFIGSAMPLETYPES.SAMPLETYPEKEY, CONFIGSAMPLETYPES.SAMPLETYPEID, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, (ANARUNANALYTERESULTS.CONCENTRATION/ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC, ANARUNANALYTERESULTS.RECORDTIMESTAMP "
                str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNSAMPLE.STUDYID, STUDY.STUDYNAME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, CONFIGSAMPLETYPES.SAMPLETYPEKEY, CONFIGSAMPLETYPES.SAMPLETYPEID, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, (ANARUNANALYTERESULTS.CONCENTRATION/ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC, ANARUNANALYTERESULTS.RECORDTIMESTAMP "
                str2 = "FROM STUDY INNER JOIN ((ASSAYANALYTES INNER JOIN ((ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) LEFT JOIN DESIGNSAMPLE ON ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) ON (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (ASSAYANALYTES.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAYANALYTES.STUDYID = ANALYTICALRUN.STUDYID)) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON STUDY.STUDYID = ASSAY.STUDYID "
                str3 = "WHERE(((ASSAY.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"

                'ANARUNRAWANALYTEPEAK
                'RUNANALYTEREGRESSIONSTATUS

                '20160816 LEE: Had to completely redo. Ricerca PLASMA study was not returning any Recovery samples because no regression was performed in those samples
                str1 = "SELECT " & strAnaRunPeak & ".ANALYTEINDEX, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strAnaRunPeak & ".RUNID, " & strAnaRunPeak & ".STUDYID, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, " & strAnaRunPeak & ".RECORDTIMESTAMP, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUN.ASSAYID, ANARUNANALYTERESULTS.CONCENTRATION, ([ANARUNANALYTERESULTS].[CONCENTRATION]/[ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR]) AS REPORTEDCONC, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, STUDY.STUDYNAME, ASSAY.MASTERASSAYID, ASSAYANALYTES.ANALYTEID "
                str2 = "FROM (STUDY INNER JOIN (CONFIGSAMPLETYPES INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN (" & strAnaRunPeak & " LEFT JOIN ANARUNANALYTERESULTS ON (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNSAMPLE.RUNID = " & strAnaRunPeak & ".RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = " & strAnaRunPeak & ".STUDYID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER)) ON (ANALYTICALRUN.RUNID = ANALYTICALRUNSAMPLE.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNSAMPLE.STUDYID)) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strAnaRunPeak & ".ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID)) ON (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON CONFIGSAMPLETYPES.SAMPLETYPEKEY = ASSAY.SAMPLETYPEKEY) ON STUDY.STUDYID = ASSAY.STUDYID) INNER JOIN ASSAYANALYTES ON (ANALYTICALRUN.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ANALYTICALRUNANALYTES.STUDYID = ASSAYANALYTES.STUDYID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX) "
                str3 = "WHERE (((" & strAnaRunPeak & ".STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"

                ' (Round([ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR],12)) AS ALIQUOTFACTOR, 
                '20171124 LEE: Must round aliquotfactor to 12 to account for floating point probs with such dilnF as 1/11 or 1/51
                str1 = "SELECT " & strAnaRunPeak & ".ANALYTEINDEX, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strAnaRunPeak & ".RUNID, " & strAnaRunPeak & ".STUDYID, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, " & strAnaRunPeak & ".RECORDTIMESTAMP, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL,  (Round([ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR],12)) AS ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUN.ASSAYID, ANARUNANALYTERESULTS.CONCENTRATION, ([ANARUNANALYTERESULTS].[CONCENTRATION]/[ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR]) AS REPORTEDCONC, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, STUDY.STUDYNAME, ASSAY.MASTERASSAYID, ASSAYANALYTES.ANALYTEID "
                str2 = "FROM (STUDY INNER JOIN (CONFIGSAMPLETYPES INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN (" & strAnaRunPeak & " LEFT JOIN ANARUNANALYTERESULTS ON (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNSAMPLE.RUNID = " & strAnaRunPeak & ".RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = " & strAnaRunPeak & ".STUDYID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER)) ON (ANALYTICALRUN.RUNID = ANALYTICALRUNSAMPLE.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNSAMPLE.STUDYID)) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strAnaRunPeak & ".ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID)) ON (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON CONFIGSAMPLETYPES.SAMPLETYPEKEY = ASSAY.SAMPLETYPEKEY) ON STUDY.STUDYID = ASSAY.STUDYID) INNER JOIN ASSAYANALYTES ON (ANALYTICALRUN.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ANALYTICALRUNANALYTES.STUDYID = ASSAYANALYTES.STUDYID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX) "
                str3 = "WHERE (((" & strAnaRunPeak & ".STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"

                'ANARUNRAWANALYTEPEAK
                'date

            Else

                'Add user id info. User id info not exported to .mdb
                str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".STUDY.STUDYNAME, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & "." & strAnaRunPeak & ".ANALYTEHEIGHT, " & strSchema & "." & strAnaRunPeak & ".ANALYTEAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTDNAME, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & "." & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, (" & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION/" & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".ANARUNANALYTERESULTS.RECORDTIMESTAMP, " & strSchema & ".SECUSERACCOUNTS.LOGINNAME, " & strSchema & ".SECUSERACCOUNTS.FIRSTNAME, " & strSchema & ".SECUSERACCOUNTS.MIDDLEINITIAL, " & strSchema & ".SECUSERACCOUNTS.LASTNAME "
                str2 = "FROM " & strSchema & ".SECUSERACCOUNTS INNER JOIN (" & strSchema & ".ASSAYANALYTES INNER JOIN ((((" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & "." & strAnaRunPeak & " INNER JOIN (" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID)) ON (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) ON (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) INNER JOIN " & strSchema & ".STUDY ON " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".STUDY.STUDYID) LEFT JOIN " & strSchema & ".DESIGNSAMPLE ON " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) LEFT JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".DESIGNSAMPLE.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID)) ON " & strSchema & ".SECUSERACCOUNTS.USERID = " & strSchema & "." & strAnaRunPeak & ".USERID "
                str3 = "WHERE(((" & strSchema & ".ASSAY.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"

                '20160224 LEE: re-arranged joins to get complete CONFIGSAMPLETYPES.SAMPLETYPEID (matrix)
                str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".STUDY.STUDYNAME, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & "." & strAnaRunPeak & ".ANALYTEHEIGHT, " & strSchema & "." & strAnaRunPeak & ".ANALYTEAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTDNAME, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & "." & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, (" & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION/" & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".ANARUNANALYTERESULTS.RECORDTIMESTAMP "
                str2 = "FROM ((((" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & "." & strAnaRunPeak & " INNER JOIN (" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID)) ON (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) ON (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) INNER JOIN " & strSchema & ".STUDY ON " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".STUDY.STUDYID) LEFT JOIN " & strSchema & ".DESIGNSAMPLE ON " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                str3 = "WHERE(((" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"

                '20160225 LEE: had to re-optimize: previous query would non open in design mode for Ricerca 032852, but can open in Frontage
                'str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".STUDY.STUDYNAME, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & "." & strAnaRunPeak & ".ANALYTEHEIGHT, " & strSchema & "." & strAnaRunPeak & ".ANALYTEAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTDNAME, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & "." & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, (" & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION/" & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".ANARUNANALYTERESULTS.RECORDTIMESTAMP "
                str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".STUDY.STUDYNAME, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & "." & strAnaRunPeak & ".ANALYTEHEIGHT, " & strSchema & "." & strAnaRunPeak & ".ANALYTEAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTDNAME, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & "." & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, (" & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION/" & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".ANARUNANALYTERESULTS.RECORDTIMESTAMP "
                str2 = "FROM " & strSchema & ".STUDY INNER JOIN ((" & strSchema & ".ASSAYANALYTES INNER JOIN ((" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & "." & strAnaRunPeak & " INNER JOIN (" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX =" & strSchema & ". ANARUNANALYTERESULTS.ANALYTEINDEX)) LEFT JOIN " & strSchema & ".DESIGNSAMPLE ON " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON " & strSchema & ".STUDY.STUDYID = " & strSchema & ".ASSAY.STUDYID "
                str3 = "WHERE(((" & strSchema & ".ASSAY.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"

                'RUNANALYTEREGRESSIONSTATUS

                '20160816 LEE: Had to completely redo. Ricerca PLASMA study was not returning any Recovery samples because no regression was performed in those samples
                str1 = "SELECT " & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & "." & strAnaRunPeak & ".RUNID, " & strSchema & "." & strAnaRunPeak & ".STUDYID, " & strSchema & "." & strAnaRunPeak & ".ANALYTEHEIGHT, " & strSchema & "." & strAnaRunPeak & ".ANALYTEAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTDNAME, " & strSchema & "." & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, " & strSchema & "." & strAnaRunPeak & ".RECORDTIMESTAMP, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, (" & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION/" & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".STUDY.STUDYNAME, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTES.ANALYTEID "
                str2 = "FROM (" & strSchema & ".STUDY INNER JOIN (" & strSchema & ".CONFIGSAMPLETYPES INNER JOIN (" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN (" & strSchema & "." & strAnaRunPeak & " LEFT JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & "." & strAnaRunPeak & ".RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & "." & strAnaRunPeak & ".STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY = " & strSchema & ".ASSAY.SAMPLETYPEKEY) ON " & strSchema & ".STUDY.STUDYID = " & strSchema & ".ASSAY.STUDYID) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ANALYTICALRUN.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX) "
                str3 = "WHERE (((" & strSchema & "." & strAnaRunPeak & ".STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"

            End If


        End If

        'ANARUNANALYTERESULTS.RECORDTIMESTAMP
        strSQL = str1 & str2 & str3 & str4

        'Console.WriteLine("tblAnalysisResultsHome: " & strSQL)
        rs.CursorLocation = CursorLocationEnum.adUseClient
        Try
            rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try


        'Try
        '    rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        'Catch ex As Exception
        '    If boolMDB Then
        '        'add RECORDTIMESTAMP
        '        str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNSAMPLE.STUDYID, STUDY.STUDYNAME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, ANARUNRAWANALYTEPEAK_INJECT.RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, ANARUNRAWANALYTEPEAK_INJECT.SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANARUNRAWANALYTEPEAK_INJECT.ANALYTEHEIGHT, ANARUNRAWANALYTEPEAK_INJECT.ANALYTEAREA, ANARUNRAWANALYTEPEAK_INJECT.INTERNALSTANDARDHEIGHT, ANARUNRAWANALYTEPEAK_INJECT.INTERNALSTANDARDAREA, ANARUNRAWANALYTEPEAK_INJECT.INTERNALSTDNAME, CONFIGSAMPLETYPES.SAMPLETYPEKEY, CONFIGSAMPLETYPES.SAMPLETYPEID, ANARUNRAWANALYTEPEAK_INJECT.ANALYTEPEAKRETENTIONTIME, ANARUNRAWANALYTEPEAK_INJECT.INTERNALSTANDARDRETENTIONTIME, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, (ANARUNANALYTERESULTS.CONCENTRATION/ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC, ANARUNANALYTERESULTS.RECORDTIMESTAMP "
        '        str2 = "FROM ASSAYANALYTES INNER JOIN ((((ANALYTICALRUNANALYTES INNER JOIN (ANARUNRAWANALYTEPEAK_INJECT INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID)) ON (ASSAY.STUDYID = ANALYTICALRUN.STUDYID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.RUNID = ANALYTICALRUN.RUNID)) ON (ANARUNRAWANALYTEPEAK_INJECT.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANARUNRAWANALYTEPEAK_INJECT.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANARUNRAWANALYTEPEAK_INJECT.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANARUNRAWANALYTEPEAK_INJECT.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID)) INNER JOIN STUDY ON ANALYTICALRUNANALYTES.STUDYID = STUDY.STUDYID) LEFT JOIN DESIGNSAMPLE ON ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) LEFT JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ASSAYANALYTES.STUDYID = ANALYTICALRUN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ANALYTICALRUN.ASSAYID) "
        '        str3 = "WHERE(((ASSAY.STUDYID) = " & wStudyID & ")) "
        '        str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"

        '    Else

        '        'Add user id info. User id info not exported to .mdb
        '        str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNSAMPLE.STUDYID, STUDY.STUDYNAME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, ANARUNRAWANALYTEPEAK_INJECT.RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, ANARUNRAWANALYTEPEAK_INJECT.SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANARUNRAWANALYTEPEAK_INJECT.ANALYTEHEIGHT, ANARUNRAWANALYTEPEAK_INJECT.ANALYTEAREA, ANARUNRAWANALYTEPEAK_INJECT.INTERNALSTANDARDHEIGHT, ANARUNRAWANALYTEPEAK_INJECT.INTERNALSTANDARDAREA, ANARUNRAWANALYTEPEAK_INJECT.INTERNALSTDNAME, CONFIGSAMPLETYPES.SAMPLETYPEKEY, CONFIGSAMPLETYPES.SAMPLETYPEID, ANARUNRAWANALYTEPEAK_INJECT.ANALYTEPEAKRETENTIONTIME, ANARUNRAWANALYTEPEAK_INJECT.INTERNALSTANDARDRETENTIONTIME, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, (ANARUNANALYTERESULTS.CONCENTRATION/ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC, ANARUNANALYTERESULTS.RECORDTIMESTAMP, SECUSERACCOUNTS.LOGINNAME, SECUSERACCOUNTS.FIRSTNAME, SECUSERACCOUNTS.MIDDLEINITIAL, SECUSERACCOUNTS.LASTNAME "
        '        str2 = "FROM SECUSERACCOUNTS INNER JOIN (ASSAYANALYTES INNER JOIN ((((ANALYTICALRUNANALYTES INNER JOIN (ANARUNRAWANALYTEPEAK_INJECT INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID)) ON (ASSAY.STUDYID = ANALYTICALRUN.STUDYID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.RUNID = ANALYTICALRUN.RUNID)) ON (ANARUNRAWANALYTEPEAK_INJECT.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANARUNRAWANALYTEPEAK_INJECT.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANARUNRAWANALYTEPEAK_INJECT.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANARUNRAWANALYTEPEAK_INJECT.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID)) INNER JOIN STUDY ON ANALYTICALRUNANALYTES.STUDYID = STUDY.STUDYID) LEFT JOIN DESIGNSAMPLE ON ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) LEFT JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ASSAYANALYTES.STUDYID = ANALYTICALRUN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ANALYTICALRUN.ASSAYID)) ON SECUSERACCOUNTS.USERID = ANARUNRAWANALYTEPEAK_INJECT.USERID "
        '        str3 = "WHERE(((ASSAY.STUDYID) = " & wStudyID & ")) "
        '        str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"

        '    End If

        '    strSQL = str1 & str2 & str3 & str4
        '    rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        'End Try

        rs.ActiveConnection = Nothing

        tblAnalysisResultsHome.Clear()
        tblAnalysisResultsHome.AcceptChanges()

        tblAnalysisResultsHome.BeginLoadData()
        daDoPr.Fill(tblAnalysisResultsHome, rs)
        tblAnalysisResultsHome.EndLoadData()

        var1 = tblAnalysisResultsHome.Rows.Count 'debug

        'int1 = tblAnalysisResultsHome.Rows.Count 'for debugging

        rs.Close()

        'now get extra anal runs
        '20160224 LEE: Don't need this anymore

        'If boolANSI Then
        '    str1 = "SELECT DISTINCT " & strAnaRunPeak & ".RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strAnaRunPeak & ".SAMPLENAME, ASSAY.MASTERASSAYID, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, " & strAnaRunPeak & ".STUDYID, " & strAnaRunPeak & ".ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANARUNANALYTERESULTS.ELIMINATEDFLAG, ASSAY.ASSAYID, STUDY.STUDYNAME, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME "
        '    'str2 = "FROM ((ASSAY INNER JOIN " & strAnaRunPeak & " ON (ASSAY.RUNID = " & strAnaRunPeak & ".RUNID) AND (ASSAY.STUDYID = " & strAnaRunPeak & ".STUDYID)) LEFT JOIN ANARUNANALYTERESULTS ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) INNER JOIN STUDY ON " & strAnaRunPeak & ".STUDYID = STUDY.STUDYID "
        '    str2 = "FROM STUDY INNER JOIN ((ASSAY INNER JOIN " & strAnaRunPeak & " ON (ASSAY.STUDYID = " & strAnaRunPeak & ".STUDYID) AND (ASSAY.RUNID = " & strAnaRunPeak & ".RUNID)) LEFT JOIN ANARUNANALYTERESULTS ON (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) ON STUDY.STUDYID = ASSAY.STUDYID "
        '    str3 = "WHERE(((" & strAnaRunPeak & ".STUDYID) = " & wStudyID & ")) "
        '    str4 = "ORDER BY " & strAnaRunPeak & ".RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"
        'Else
        '    str1 = "SELECT DISTINCT " & strAnaRunPeak & ".RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strAnaRunPeak & ".SAMPLENAME, ASSAY.MASTERASSAYID, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, " & strAnaRunPeak & ".STUDYID, " & strAnaRunPeak & ".ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANARUNANALYTERESULTS.ELIMINATEDFLAG, ASSAY.ASSAYID, STUDY.STUDYNAME, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME "
        '    str2 = "FROM STUDY, ASSAY, " & strAnaRunPeak & ", ANARUNANALYTERESULTS "
        '    str2 = str2 & "WHERE (((ASSAY.STUDYID = " & strAnaRunPeak & ".STUDYID) AND (ASSAY.RUNID = " & strAnaRunPeak & ".RUNID)) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID(+)) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID(+)) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER(+)) AND (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX(+))) AND STUDY.STUDYID = ASSAY.STUDYID "
        '    'str2 = str2 & "WHERE (((ASSAY.STUDYID = " & strAnaRunPeak & ".STUDYID) AND (ASSAY.RUNID = " & strAnaRunPeak & ".RUNID)) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID(+)) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) AND STUDY.STUDYID = ASSAY.STUDYID "

        '    str3 = "AND (((" & strAnaRunPeak & ".STUDYID) = " & wStudyID & ")) "
        '    str4 = "ORDER BY " & strAnaRunPeak & ".RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"

        'End If

        'If boolANSI Then
        '    str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, " & strAnaRunPeak & ".RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strAnaRunPeak & ".SAMPLENAME, ASSAY.MASTERASSAYID, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, " & strAnaRunPeak & ".STUDYID, " & strAnaRunPeak & ".ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANARUNANALYTERESULTS.ELIMINATEDFLAG, ASSAY.ASSAYID, STUDY.STUDYNAME, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME "
        '    str2 = "FROM ASSAYANALYTES INNER JOIN (STUDY INNER JOIN ((ASSAY INNER JOIN " & strAnaRunPeak & " ON (ASSAY.RUNID = " & strAnaRunPeak & ".RUNID) AND (ASSAY.STUDYID = " & strAnaRunPeak & ".STUDYID)) LEFT JOIN ANARUNANALYTERESULTS ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON STUDY.STUDYID = ASSAY.STUDYID) ON (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID) AND (ASSAYANALYTES.STUDYID = ASSAY.STUDYID) "
        '    str3 = "WHERE(((" & strAnaRunPeak & ".STUDYID) = " & wStudyID & ")) "
        '    str4 = "ORDER BY " & strAnaRunPeak & ".RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"

        '    'add stuff from AnalyticalRunSample
        '    str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, " & strAnaRunPeak & ".RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strAnaRunPeak & ".SAMPLENAME, ASSAY.MASTERASSAYID, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, " & strAnaRunPeak & ".STUDYID, " & strAnaRunPeak & ".ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANARUNANALYTERESULTS.ELIMINATEDFLAG, ASSAY.ASSAYID, STUDY.STUDYNAME, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME "
        '    str2 = "FROM ANALYTICALRUNSAMPLE INNER JOIN (ASSAYANALYTES INNER JOIN (STUDY INNER JOIN ((ASSAY INNER JOIN " & strAnaRunPeak & " ON (ASSAY.RUNID = " & strAnaRunPeak & ".RUNID) AND (ASSAY.STUDYID = " & strAnaRunPeak & ".STUDYID)) LEFT JOIN ANARUNANALYTERESULTS ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON STUDY.STUDYID = ASSAY.STUDYID) ON (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID) AND (ASSAYANALYTES.STUDYID = ASSAY.STUDYID)) ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = " & strAnaRunPeak & ".RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = " & strAnaRunPeak & ".STUDYID) "
        '    str3 = "WHERE(((" & strAnaRunPeak & ".STUDYID) = " & wStudyID & ")) "
        '    str4 = "ORDER BY " & strAnaRunPeak & ".RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"

        '    If boolAccess Then
        '        'add RUNSAMPLEORDERNUMBER
        '        str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, " & strAnaRunPeak & ".RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strAnaRunPeak & ".SAMPLENAME, ASSAY.MASTERASSAYID, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, " & strAnaRunPeak & ".STUDYID, " & strAnaRunPeak & ".ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANARUNANALYTERESULTS.ELIMINATEDFLAG, ASSAY.ASSAYID, STUDY.STUDYNAME, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, ANARUNANALYTERESULTS.RECORDTIMESTAMP "
        '        str2 = "FROM ANALYTICALRUNSAMPLE INNER JOIN (ASSAYANALYTES INNER JOIN (STUDY INNER JOIN ((ASSAY INNER JOIN " & strAnaRunPeak & " ON (ASSAY.RUNID = " & strAnaRunPeak & ".RUNID) AND (ASSAY.STUDYID = " & strAnaRunPeak & ".STUDYID)) LEFT JOIN ANARUNANALYTERESULTS ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON STUDY.STUDYID = ASSAY.STUDYID) ON (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID) AND (ASSAYANALYTES.STUDYID = ASSAY.STUDYID)) ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = " & strAnaRunPeak & ".RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = " & strAnaRunPeak & ".STUDYID) "
        '        str3 = "WHERE(((" & strAnaRunPeak & ".STUDYID) = " & wStudyID & ")) "
        '        str4 = "ORDER BY " & strAnaRunPeak & ".RUNID, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"
        '    Else

        '        'add user id info
        '        str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & "." & strAnaRunPeak & ".RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & "." & strAnaRunPeak & ".ANALYTEHEIGHT, " & strSchema & "." & strAnaRunPeak & ".ANALYTEAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTDNAME, " & strSchema & "." & strAnaRunPeak & ".STUDYID, " & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".STUDY.STUDYNAME, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & "." & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, " & strSchema & ".ANARUNANALYTERESULTS.RECORDTIMESTAMP, " & strSchema & ".SECUSERACCOUNTS.LOGINNAME, " & strSchema & ".SECUSERACCOUNTS.FIRSTNAME, " & strSchema & ".SECUSERACCOUNTS.MIDDLEINITIAL, " & strSchema & ".SECUSERACCOUNTS.LASTNAME "
        '        str2 = "FROM (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN (" & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".STUDY INNER JOIN ((" & strSchema & ".ASSAY INNER JOIN " & strSchema & "." & strAnaRunPeak & " ON (" & strSchema & ".ASSAY.STUDYID = " & strSchema & "." & strAnaRunPeak & ".STUDYID) AND (" & strSchema & ".ASSAY.RUNID = " & strSchema & "." & strAnaRunPeak & ".RUNID)) LEFT JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX)) ON " & strSchema & ".STUDY.STUDYID = " & strSchema & ".ASSAY.STUDYID) ON (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAY.ASSAYID)) ON (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & "." & strAnaRunPeak & ".STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & "." & strAnaRunPeak & ".RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER)) INNER JOIN " & strSchema & ".SECUSERACCOUNTS ON " & strSchema & "." & strAnaRunPeak & ".USERID = " & strSchema & ".SECUSERACCOUNTS.USERID "
        '        str3 = "WHERE(((" & strSchema & "." & strAnaRunPeak & ".STUDYID) = " & wStudyID & ")) "
        '        str4 = "ORDER BY " & strSchema & "." & strAnaRunPeak & ".RUNID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"


        '    End If

        'End If

        'strSQL = str1 & str2 & str3 & str4

        ''console.writeline("tblARHTemp: " & strSQL)


        'rs.CursorLocation = CursorLocationEnum.adUseClient
        'rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        ''Try
        ''    rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

        ''Catch ex As Exception

        ''    If boolMDB Then
        ''        'add RUNSAMPLEORDERNUMBER
        ''        str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ANARUNRAWANALYTEPEAK_INJECT.RUNID, ANARUNRAWANALYTEPEAK_INJECT.RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANARUNRAWANALYTEPEAK_INJECT.SAMPLENAME, ASSAY.MASTERASSAYID, ANARUNRAWANALYTEPEAK_INJECT.ANALYTEHEIGHT, ANARUNRAWANALYTEPEAK_INJECT.ANALYTEAREA, ANARUNRAWANALYTEPEAK_INJECT.INTERNALSTANDARDHEIGHT, ANARUNRAWANALYTEPEAK_INJECT.INTERNALSTANDARDAREA, ANARUNRAWANALYTEPEAK_INJECT.INTERNALSTDNAME, ANARUNRAWANALYTEPEAK_INJECT.STUDYID, ANARUNRAWANALYTEPEAK_INJECT.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANARUNANALYTERESULTS.ELIMINATEDFLAG, ASSAY.ASSAYID, STUDY.STUDYNAME, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANARUNRAWANALYTEPEAK_INJECT.ANALYTEPEAKRETENTIONTIME, ANARUNRAWANALYTEPEAK_INJECT.INTERNALSTANDARDRETENTIONTIME, ANARUNANALYTERESULTS.RECORDTIMESTAMP "
        ''        str2 = "FROM ANALYTICALRUNSAMPLE INNER JOIN (ASSAYANALYTES INNER JOIN (STUDY INNER JOIN ((ASSAY INNER JOIN ANARUNRAWANALYTEPEAK_INJECT ON (ASSAY.RUNID = ANARUNRAWANALYTEPEAK_INJECT.RUNID) AND (ASSAY.STUDYID = ANARUNRAWANALYTEPEAK_INJECT.STUDYID)) LEFT JOIN ANARUNANALYTERESULTS ON (ANARUNRAWANALYTEPEAK_INJECT.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (ANARUNRAWANALYTEPEAK_INJECT.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANARUNRAWANALYTEPEAK_INJECT.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANARUNRAWANALYTEPEAK_INJECT.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON STUDY.STUDYID = ASSAY.STUDYID) ON (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID) AND (ASSAYANALYTES.STUDYID = ASSAY.STUDYID)) ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNRAWANALYTEPEAK_INJECT.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNRAWANALYTEPEAK_INJECT.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNRAWANALYTEPEAK_INJECT.STUDYID) "
        ''        str3 = "WHERE(((ANARUNRAWANALYTEPEAK_INJECT.STUDYID) = " & wStudyID & ")) "
        ''        str4 = "ORDER BY ANARUNRAWANALYTEPEAK_INJECT.RUNID, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"
        ''    Else

        ''        'add user id info
        ''        str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ANARUNRAWANALYTEPEAK_INJECT.RUNID, ANARUNRAWANALYTEPEAK_INJECT.RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANARUNRAWANALYTEPEAK_INJECT.SAMPLENAME, ASSAY.MASTERASSAYID, ANARUNRAWANALYTEPEAK_INJECT.ANALYTEHEIGHT, ANARUNRAWANALYTEPEAK_INJECT.ANALYTEAREA, ANARUNRAWANALYTEPEAK_INJECT.INTERNALSTANDARDHEIGHT, ANARUNRAWANALYTEPEAK_INJECT.INTERNALSTANDARDAREA, ANARUNRAWANALYTEPEAK_INJECT.INTERNALSTDNAME, ANARUNRAWANALYTEPEAK_INJECT.STUDYID, ANARUNRAWANALYTEPEAK_INJECT.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANARUNANALYTERESULTS.ELIMINATEDFLAG, ASSAY.ASSAYID, STUDY.STUDYNAME, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANARUNRAWANALYTEPEAK_INJECT.ANALYTEPEAKRETENTIONTIME, ANARUNRAWANALYTEPEAK_INJECT.INTERNALSTANDARDRETENTIONTIME, ANARUNANALYTERESULTS.RECORDTIMESTAMP, SECUSERACCOUNTS.LOGINNAME, SECUSERACCOUNTS.FIRSTNAME, SECUSERACCOUNTS.MIDDLEINITIAL, SECUSERACCOUNTS.LASTNAME "
        ''        str2 = "FROM (ANALYTICALRUNSAMPLE INNER JOIN (ASSAYANALYTES INNER JOIN (STUDY INNER JOIN ((ASSAY INNER JOIN ANARUNRAWANALYTEPEAK_INJECT ON (ASSAY.STUDYID = ANARUNRAWANALYTEPEAK_INJECT.STUDYID) AND (ASSAY.RUNID = ANARUNRAWANALYTEPEAK_INJECT.RUNID)) LEFT JOIN ANARUNANALYTERESULTS ON (ANARUNRAWANALYTEPEAK_INJECT.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANARUNRAWANALYTEPEAK_INJECT.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANARUNRAWANALYTEPEAK_INJECT.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANARUNRAWANALYTEPEAK_INJECT.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) ON STUDY.STUDYID = ASSAY.STUDYID) ON (ASSAYANALYTES.STUDYID = ASSAY.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID)) ON (ANALYTICALRUNSAMPLE.STUDYID = ANARUNRAWANALYTEPEAK_INJECT.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNRAWANALYTEPEAK_INJECT.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNRAWANALYTEPEAK_INJECT.RUNSAMPLESEQUENCENUMBER)) INNER JOIN SECUSERACCOUNTS ON ANARUNRAWANALYTEPEAK_INJECT.USERID = SECUSERACCOUNTS.USERID "
        ''        str3 = "WHERE(((ANARUNRAWANALYTEPEAK_INJECT.STUDYID) = " & wStudyID & ")) "
        ''        str4 = "ORDER BY ANARUNRAWANALYTEPEAK_INJECT.RUNID, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"

        ''    End If

        ''    strSQL = str1 & str2 & str3 & str4
        ''    rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

        ''End Try
        'rs.ActiveConnection = Nothing


        ' '''''''''console.writeline("tblARHTemp: " & strSQL)

        'tblARHTemp.Clear()

        'tblARHTemp.BeginLoadData()
        'daDoPr.Fill(tblARHTemp, rs)
        'tblARHTemp.EndLoadData()


        ''int2 = tblARHTemp.Rows.Count 'FOR TESTING
        ' '''''''''''''''''console.writeline("rs rows: " & int1)

        'tbl1 = tblAnalysisResultsHome
        'tbl2 = tblAnalytesHome
        'tbl3 = tblARHTemp


        ''make distinct table in tblarhtemp for runid
        'Dim dv1 As System.Data.DataView = New DataView(tbl3)
        'Dim tbl3a As System.Data.DataTable = dv1.ToTable("a", True, "RUNID", "ANALYTEID")
        'int3 = tbl3a.Rows.Count
        'For Count3 = 0 To int3 - 1
        '    var1 = tbl3a.Rows.Item(Count3).Item("RUNID")
        '    var2 = tbl3a.Rows.Item(Count3).Item("ANALYTEID")
        '    strF = "RUNID = " & var1 & " AND ANALYTEID = " & var2
        '    rows1 = tbl1.Select(strF)
        '    int1 = rows1.Length
        '    If int1 = 0 Then
        '        rows3 = tbl3.Select(strF)
        '        int3 = rows3.Length
        '        For Count2 = 0 To int3 - 1
        '            Dim row As DataRow = tbl1.NewRow
        '            row.BeginEdit()
        '            For Each dcol In tbl3.Columns
        '                var1 = rows3(Count2).Item(dcol.ColumnName)
        '                row(dcol.ColumnName) = var1
        '            Next
        '            row.EndEdit()
        '            tbl1.Rows.Add(row)
        '        Next
        '    End If
        ''Next

        'rs.Close()


        tbl1 = tblAnalysisResultsHome
        tbl2 = tblAnalytesHome

        'Update CHARANALYTE
        Dim strF2 As String

        'sometimes analytes have identical index and m_assayid
        'must INCLUDE assayid

        'this looks like it would take a long time, but it actually doesn't
        If boolUseGroups Then

            'don't need to filter for IS because tblCalStdGroupAssayIDsAll has only Analytes
            For Count1 = 0 To tblCalStdGroupAssayIDsAll.Rows.Count - 1
                var1 = tblCalStdGroupAssayIDsAll.Rows(Count1).Item("ANALYTEID")
                var2 = tblCalStdGroupAssayIDsAll.Rows(Count1).Item("ASSAYID")
                var3 = tblCalStdGroupAssayIDsAll.Rows(Count1).Item("ANALYTEDESCRIPTION_C")
                var4 = tblCalStdGroupAssayIDsAll.Rows(Count1).Item("INTGROUP")
                strF = "ANALYTEID = " & var1 & " AND ASSAYID = " & var2

                Erase row1
                row1 = tbl1.Select(strF)
                If row1.Length = 0 Then
                Else
                    For Count2 = 0 To row1.Length - 1
                        row1(Count2).BeginEdit()
                        row1(Count2).Item("CHARANALYTE") = var3
                        Try
                            row1(Count2).Item("INTGROUP") = var4
                        Catch ex As Exception
                            var1 = var1
                        End Try
                        row1(Count2).EndEdit()
                    Next
                End If
                var1 = var1 'DEBUG
            Next
            'tblCalStdGroupAssayIDsAll
        Else

            strF = "IsIntStd = 'No'"
            Erase row2
            row2 = tbl2.Select(strF)

            intRows1 = tbl1.Rows.Count
            intRows2 = row2.Length

            For Count1 = 0 To intRows2 - 1
                'retrieve CHARANALYTE from tblAnalytes
                var1 = row2(Count1).Item("ANALYTEINDEX")
                var2 = row2(Count1).Item("MASTERASSAYID")
                var3 = row2(Count1).Item("ANALYTEDESCRIPTION")
                var4 = row2(Count1).Item("ANALYTEID")
                strF = "ANALYTEINDEX = " & var1 & " AND MASTERASSAYID = " & var2 & " AND ANALYTEID = " & var4

                Erase row1
                row1 = tbl1.Select(strF)
                If row1.Length = 0 Then
                Else
                    For Count2 = 0 To row1.Length - 1
                        row1(Count2).BeginEdit()
                        row1(Count2).Item("CHARANALYTE") = var3
                        row1(Count2).EndEdit()
                    Next
                End If
                var1 = var1 'DEBUG
            Next
        End If


        'Note: tblAssignedSamples may have reference to additional studies than the parent study
        'must add rows for these additional studies

        'make distinct table for id_tblStudies2 in tbl4
        strF = "ID_TBLSTUDIES = " & id_tblStudies
        strS = "ID_TBLSTUDIES ASC"

        '20160224 LEE: tbl4 (tblAssignedSamples) will have to be set again because tblAssignedSamples is filtered for current study
        int4 = tbl4.Rows.Count 'DEBUG
        Dim dv2 As System.Data.DataView = New DataView(tbl4, strF, strS, DataViewRowState.CurrentRows)
        int4 = dv2.Count 'debug
        Dim tbl4a As System.Data.DataTable = dv2.ToTable("b", True, "ID_TBLSTUDIES2")
        int4 = tbl4a.Rows.Count
        Dim id As Int64
        Dim lng1 As Int64
        For Count4 = 0 To int4 - 1

            lng1 = tbl4a.Rows(Count4).Item("ID_TBLSTUDIES2")
            id = GetWStudyID(lng1)
            If lng1 = id_tblStudies Then 'ignore
            Else
                If boolAccess Then

                    ''20160224 LEE: re-arranged joins to get complete CONFIGSAMPLETYPES.SAMPLETYPEID (matrix)
                    'str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNSAMPLE.STUDYID, STUDY.STUDYNAME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, CONFIGSAMPLETYPES.SAMPLETYPEKEY, CONFIGSAMPLETYPES.SAMPLETYPEID, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, (ANARUNANALYTERESULTS.CONCENTRATION/ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC, ANARUNANALYTERESULTS.RECORDTIMESTAMP "
                    'str2 = "FROM ((((ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID)) ON (ASSAY.STUDYID = ANALYTICALRUN.STUDYID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.RUNID = ANALYTICALRUN.RUNID)) ON (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID)) INNER JOIN STUDY ON ANALYTICALRUNANALYTES.STUDYID = STUDY.STUDYID) LEFT JOIN DESIGNSAMPLE ON ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) INNER JOIN ASSAYANALYTES ON (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX)) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                    'str3 = "WHERE(((ANALYTICALRUNSAMPLE.STUDYID) = " & id & ")) "
                    'str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"

                    ''20160225 LEE: had to re-optimize: previous query would non open in design mode for Ricerca 032852, but can open in Frontage
                    'str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, CONFIGSAMPLETYPES.SAMPLETYPEKEY, CONFIGSAMPLETYPES.SAMPLETYPEID, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, (ANARUNANALYTERESULTS.CONCENTRATION/ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC, ANARUNANALYTERESULTS.RECORDTIMESTAMP "
                    'str2 = "FROM (ASSAYANALYTES INNER JOIN ((ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) LEFT JOIN DESIGNSAMPLE ON ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) ON (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (ASSAYANALYTES.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAYANALYTES.STUDYID = ANALYTICALRUN.STUDYID)) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                    'str3 = "WHERE(((ASSAY.STUDYID) = " & id & ")) "
                    'str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"

                    '20160816 LEE: Had to completely redo. Ricerca PLASMA study was not returning any Recovery samples because no regression was performed in those samples
                    str1 = "SELECT " & strAnaRunPeak & ".ANALYTEINDEX, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strAnaRunPeak & ".RUNID, " & strAnaRunPeak & ".STUDYID, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, " & strAnaRunPeak & ".RECORDTIMESTAMP, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUN.ASSAYID, ANARUNANALYTERESULTS.CONCENTRATION, ([ANARUNANALYTERESULTS].[CONCENTRATION]/[ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR]) AS REPORTEDCONC, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, STUDY.STUDYNAME, ASSAY.MASTERASSAYID, ASSAYANALYTES.ANALYTEID "
                    str2 = "FROM (STUDY INNER JOIN (CONFIGSAMPLETYPES INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN (" & strAnaRunPeak & " LEFT JOIN ANARUNANALYTERESULTS ON (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNSAMPLE.RUNID = " & strAnaRunPeak & ".RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = " & strAnaRunPeak & ".STUDYID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER)) ON (ANALYTICALRUN.RUNID = ANALYTICALRUNSAMPLE.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNSAMPLE.STUDYID)) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strAnaRunPeak & ".ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID)) ON (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON CONFIGSAMPLETYPES.SAMPLETYPEKEY = ASSAY.SAMPLETYPEKEY) ON STUDY.STUDYID = ASSAY.STUDYID) INNER JOIN ASSAYANALYTES ON (ANALYTICALRUN.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ANALYTICALRUNANALYTES.STUDYID = ASSAYANALYTES.STUDYID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX) "
                    str3 = "WHERE (((" & strAnaRunPeak & ".STUDYID)=" & id & ")) "
                    str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"

                    ' (Round([ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR],12)) AS ALIQUOTFACTOR, 
                    '20171124 LEE: Must round aliquotfactor to 12 to account for floating point probs with such dilnF as 1/11 or 1/51
                    str1 = "SELECT " & strAnaRunPeak & ".ANALYTEINDEX, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strAnaRunPeak & ".RUNID, " & strAnaRunPeak & ".STUDYID, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, " & strAnaRunPeak & ".RECORDTIMESTAMP, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, (Round([ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR],12)) AS ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUN.ASSAYID, ANARUNANALYTERESULTS.CONCENTRATION, ([ANARUNANALYTERESULTS].[CONCENTRATION]/[ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR]) AS REPORTEDCONC, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, STUDY.STUDYNAME, ASSAY.MASTERASSAYID, ASSAYANALYTES.ANALYTEID "
                    str2 = "FROM (STUDY INNER JOIN (CONFIGSAMPLETYPES INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN (" & strAnaRunPeak & " LEFT JOIN ANARUNANALYTERESULTS ON (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNSAMPLE.RUNID = " & strAnaRunPeak & ".RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = " & strAnaRunPeak & ".STUDYID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER)) ON (ANALYTICALRUN.RUNID = ANALYTICALRUNSAMPLE.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNSAMPLE.STUDYID)) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strAnaRunPeak & ".ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID)) ON (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON CONFIGSAMPLETYPES.SAMPLETYPEKEY = ASSAY.SAMPLETYPEKEY) ON STUDY.STUDYID = ASSAY.STUDYID) INNER JOIN ASSAYANALYTES ON (ANALYTICALRUN.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ANALYTICALRUNANALYTES.STUDYID = ASSAYANALYTES.STUDYID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX) "
                    str3 = "WHERE (((" & strAnaRunPeak & ".STUDYID)=" & id & ")) "
                    str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"

                    'ANARUNRAWANALYTEPEAK
                Else

                    ''20160224 LEE: re-arranged joins to get complete CONFIGSAMPLETYPES.SAMPLETYPEID (matrix)
                    'str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".STUDY.STUDYNAME, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & "." & strAnaRunPeak & ".ANALYTEHEIGHT, " & strSchema & "." & strAnaRunPeak & ".ANALYTEAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTDNAME, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & "." & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, (" & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION/" & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".ANARUNANALYTERESULTS.RECORDTIMESTAMP "
                    'str2 = "FROM ((((" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & "." & strAnaRunPeak & " INNER JOIN (" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID)) ON (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) ON (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) INNER JOIN " & strSchema & ".STUDY ON " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".STUDY.STUDYID) LEFT JOIN " & strSchema & ".DESIGNSAMPLE ON " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                    'str3 = "WHERE(((" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID) = " & id & ")) "
                    'str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"


                    ''20160225 LEE: had to re-optimize: previous query would non open in design mode for Ricerca 032852, but can open in Frontage
                    'str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & "." & strAnaRunPeak & ".ANALYTEHEIGHT, " & strSchema & "." & strAnaRunPeak & ".ANALYTEAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTDNAME, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & "." & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, (" & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION/" & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".ANARUNANALYTERESULTS.RECORDTIMESTAMP "
                    'str2 = "FROM (" & strSchema & ".ASSAYANALYTES INNER JOIN ((" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & "." & strAnaRunPeak & " INNER JOIN (" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX)) LEFT JOIN " & strSchema & ".DESIGNSAMPLE ON " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                    'str3 = "WHERE(((" & strSchema & ".ASSAY.STUDYID) = " & id & ")) "
                    'str4 = "ORDER BY " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"

                    '20160816 LEE: Had to completely redo. Ricerca PLASMA study was not returning any Recovery samples because no regression was performed in those samples
                    str1 = "SELECT " & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & "." & strAnaRunPeak & ".RUNID, " & strSchema & "." & strAnaRunPeak & ".STUDYID, " & strSchema & "." & strAnaRunPeak & ".ANALYTEHEIGHT, " & strSchema & "." & strAnaRunPeak & ".ANALYTEAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTDNAME, " & strSchema & "." & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, " & strSchema & "." & strAnaRunPeak & ".RECORDTIMESTAMP, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, (" & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION/" & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".STUDY.STUDYNAME, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTES.ANALYTEID "
                    str2 = "FROM (" & strSchema & ".STUDY INNER JOIN (" & strSchema & ".CONFIGSAMPLETYPES INNER JOIN (" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN (" & strSchema & "." & strAnaRunPeak & " LEFT JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & "." & strAnaRunPeak & ".RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & "." & strAnaRunPeak & ".STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY = " & strSchema & ".ASSAY.SAMPLETYPEKEY) ON " & strSchema & ".STUDY.STUDYID = " & strSchema & ".ASSAY.STUDYID) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ANALYTICALRUN.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX) "
                    str3 = "WHERE (((" & strSchema & "." & strAnaRunPeak & ".STUDYID)=" & id & ")) "
                    str4 = "ORDER BY " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"

                End If

                '21061228 LEE: Remember this needs to be set in AssignSamples.SetAnalysisResultsTable

                strSQL = str1 & str2 & str3 & str4

                '''''''''console.writeline("tblAnalysisResultsHome: " & strSQL)
                '''''''''''''console.writeline(strSQL)
                rs.CursorLocation = CursorLocationEnum.adUseClient
                rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

                rs.ActiveConnection = Nothing

                'tblAnalysisResultsHome.Clear()

                var1 = rs.RecordCount 'debug

                'tblAnalysisResultsHome.Clear()
                'tblAnalysisResultsHome.AcceptChanges()
                'do not clear here
                tblAnalysisResultsHome.BeginLoadData()
                daDoPr.Fill(tblAnalysisResultsHome, rs)
                tblAnalysisResultsHome.EndLoadData()

                int1 = tblAnalysisResultsHome.Rows.Count 'for debugging

                rs.Close()

                ''Update CHARANALYTE
                'strF = "IsIntStd = 'No'"
                'Erase row2
                'row2 = tbl2.Select(strF)

                'intRows1 = tbl1.Rows.Count
                'intRows2 = row2.Length
                'For Count1 = 0 To intRows2 - 1
                '    'retrieve CHARANALYTE from tblAnalytes
                '    var1 = row2(Count1).Item("ANALYTEINDEX")
                '    var2 = row2(Count1).Item("MASTERASSAYID")
                '    var3 = row2(Count1).Item("ANALYTEDESCRIPTION")
                '    var4 = row2(Count1).Item("ANALYTEID")
                '    strF = "ANALYTEINDEX = " & var1 & " AND MASTERASSAYID = " & var2 & " AND ANALYTEID = " & var4
                '    Erase row1
                '    row1 = tbl1.Select(strF)
                '    If row1.Length = 0 Then
                '    Else
                '        For Count2 = 0 To row1.Length - 1
                '            row1(Count2).BeginEdit()
                '            row1(Count2).Item("CHARANALYTE") = var3
                '            'row1(Count2).Item("ANALYTEID") = var4
                '            row1(Count2).EndEdit()
                '        Next
                '    End If
                'Next

                '*****

                '20160224 LEE: There is a problem here. 
                If boolUseGroups Then
                    For Count1 = 0 To tblCalStdGroupAssayIDsAll.Rows.Count - 1
                        var1 = tblCalStdGroupAssayIDsAll.Rows(Count1).Item("ANALYTEID")
                        var2 = tblCalStdGroupAssayIDsAll.Rows(Count1).Item("ASSAYID")
                        var3 = tblCalStdGroupAssayIDsAll.Rows(Count1).Item("ANALYTEDESCRIPTION_C")
                        var4 = tblCalStdGroupAssayIDsAll.Rows(Count1).Item("INTGROUP")

                        'Note: 2nd study will have different ASSAYID
                        strF = "ANALYTEID = " & var1 ' & " AND ASSAYID = " & var2

                        Erase row1
                        row1 = tbl1.Select(strF)
                        If row1.Length = 0 Then
                        Else
                            For Count2 = 0 To row1.Length - 1
                                row1(Count2).BeginEdit()
                                row1(Count2).Item("CHARANALYTE") = var3
                                Try
                                    row1(Count2).Item("INTGROUP") = var4
                                Catch ex As Exception
                                    var1 = var1
                                End Try
                                row1(Count2).EndEdit()
                            Next
                        End If
                        var1 = var1 'DEBUG
                    Next
                    'tblCalStdGroupAssayIDsAll
                Else

                    strF = "IsIntStd = 'No'"
                    Erase row2
                    row2 = tbl2.Select(strF)

                    intRows1 = tbl1.Rows.Count
                    intRows2 = row2.Length

                    For Count1 = 0 To intRows2 - 1
                        'retrieve CHARANALYTE from tblAnalytes
                        var1 = row2(Count1).Item("ANALYTEINDEX")
                        var2 = row2(Count1).Item("MASTERASSAYID")
                        var3 = row2(Count1).Item("ANALYTEDESCRIPTION")
                        var4 = row2(Count1).Item("ANALYTEID")
                        strF = "ANALYTEINDEX = " & var1 & " AND MASTERASSAYID = " & var2 & " AND ANALYTEID = " & var4

                        Erase row1
                        row1 = tbl1.Select(strF)
                        If row1.Length = 0 Then
                        Else
                            For Count2 = 0 To row1.Length - 1
                                row1(Count2).BeginEdit()
                                row1(Count2).Item("CHARANALYTE") = var3
                                row1(Count2).EndEdit()
                            Next
                        End If
                        var1 = var1 'DEBUG
                    Next
                End If

                '*****




            End If
        Next

        'Call CheckAssSamples(1)

end1:

        If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
            rs.Close()
        End If

        rs = Nothing



    End Sub

    Sub FillDoPrepareTables()

        'tblReassayReport
        'add two unbound columns
        If tblReassayReport.Columns.Contains("numRR") Then
        Else
            Dim colA As New DataColumn
            colA.ColumnName = "numRR" 'Reason for Reassay number
            tblReassayReport.Columns.Add(colA)
        End If
        If tblReassayReport.Columns.Contains("numRCR") Then
        Else
            Dim colB As New DataColumn
            colB.ColumnName = "numRCR" 'Reason for Reported Concentration number
            tblReassayReport.Columns.Add(colB)
        End If
        dsDoPr.Tables.Add(tblReassayReport)

        dsDoPr.Tables.Add(tblReassayReasons)
        dsDoPr.Tables.Add(tblBCStds)
        dsDoPr.Tables.Add(tblBCStdsAssayID)
        dsDoPr.Tables.Add(tblBCStdConcs)
        dsDoPr.Tables.Add(tblBCQCs)
        dsDoPr.Tables.Add(tblQCAI)
        dsDoPr.Tables.Add(tblQCReps)
        dsDoPr.Tables.Add(tblAccAnalRuns)
        dsDoPr.Tables.Add(tblwSTUDY)
        dsDoPr.Tables.Add(tblwPROJECTS)


        If tblBCStdConcs.Columns.Contains("NomConc") Then
        Else
            Dim col1 As New DataColumn
            col1.ColumnName = "NomConc"
            col1.DataType = System.Type.GetType("System.Decimal")
            tblBCQCConcs.Columns.Add(col1)
        End If

        If tblBCStdConcs.Columns.Contains("QCLABEL") Then
        Else
            Dim col1 As New DataColumn
            col1.ColumnName = "QCLABEL"
            col1.DataType = System.Type.GetType("System.String")
            tblBCQCConcs.Columns.Add(col1)
        End If

        If tblBCQCConcs.Columns.Contains("AnalyteDescription") Then
        Else
            Dim col101 As New DataColumn
            col101.ColumnName = "AnalyteDescription"
            col101.DataType = System.Type.GetType("System.String")
            tblBCQCConcs.Columns.Add(col101)
        End If

        dsDoPr.Tables.Add(tblBCQCConcs)

        dsDoPr.Tables.Add(tblQCF)
        dsDoPr.Tables.Add(tblRegCon)
        dsDoPr.Tables.Add(tblReassay)

        If tblAnalysisResultsHome.Columns.Contains("CHARANALYTE") Then
        Else
            Dim col1 As New DataColumn
            col1.ColumnName = "CHARANALYTE"
            col1.Caption = "Analyte"
            col1.DataType = System.Type.GetType("System.String")
            tblAnalysisResultsHome.Columns.Add(col1)
        End If

        If tblAnalysisResultsHome.Columns.Contains("INTGROUP") Then
        Else
            Dim col1 As New DataColumn
            col1.ColumnName = "INTGROUP"
            col1.Caption = "Group"
            col1.DataType = System.Type.GetType("System.Int16")
            tblAnalysisResultsHome.Columns.Add(col1)
        End If

        dsDoPr.Tables.Add(tblAnalysisResultsHome)

        dsDoPr.Tables.Add(tblARHTemp)
        dsDoPr.Tables.Add(tblSAMPLERESULTSCONFLICT)


    End Sub

    Sub FillAssignedSamplesDGV()

        Dim dgv As DataGridView
        'Dim tbl1 as System.Data.DataTable
        Dim strF As String
        Dim strS As String
        Dim Count1 As Int64
        Dim Count2 As Int64
        Dim Count3 As Int64
        Dim Count4 As Int64
        Dim int1 As Int64
        Dim int2 As Int64
        Dim int3 As Int64
        Dim int4 As Int64
        Dim int5 As Int64
        Dim str1 As String
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim var1, var2, var3, var4, var5, var6, var7
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim dv As System.Data.DataView
        Dim boolAdd As Boolean
        Dim dgv2 As DataGridView
        Dim wRunID
        Dim dv2 As System.Data.DataView
        Dim dv3 As System.Data.DataView
        Dim num1 As Long
        Dim num2 As Long
        Dim num3 As Long
        Dim boolFormat As Boolean
        Dim boolP As Boolean
        Dim strP As String
        Dim intOrig As String

        Dim tbl1 As System.Data.DataTable
        Dim rows10() As DataRow
        Dim tbl2 As System.Data.DataTable
        Dim rows20() As DataRow
        Dim intU As Int64

        'intOrig = Me.cbxFilterRunID.SelectedIndex
        'Me.cbxFilterRunID.SelectedIndex = 0

        'dgv = Me.dgvAssignedSamples
        'dgv2 = Me.dgvAnalyticalRuns

        ''20181216 LEE:
        ''suddenly tblAssignedSamples.AddCols_tblAss actions aren't recognized
        ''do again
        'Call AddCols_tblAss()

        tbl1 = tblAnalysisResultsHome
        tbl2 = tblAssignedSamples
        strF = "ID_TBLSTUDIES = " & id_tblStudies

        ''''''''console.writeline("Start tbl1 columns")
        'For Count1 = 0 To tbl1.Columns.Count - 1 'debugging
        '    '''''''console.writeline(tbl1.Columns(Count1).ColumnName)
        'Next
        ''''''''console.writeline("End tbl1 columns")

        Dim strF1 As String

        'find unique id_tblstudies2 in tbl2
        Dim dvU As System.Data.DataView = New DataView(tbl2, strF, "ID_TBLSTUDIES ASC", DataViewRowState.CurrentRows)
        Dim tblU As System.Data.DataTable = dvU.ToTable("a", True, "ID_TBLSTUDIES2")

        'int4 = tbl1.Rows.Count 'for debugging
        'int5 = tbl1.Columns.Count 'for debugging
        int2 = tbl1.Columns.Count 'debug

        'Dim cn As New ADODB.Connection
        'Dim rs As New ADODB.Recordset

        For Count4 = 0 To tblU.Rows.Count - 1
            intU = tblU.Rows(Count4).Item("ID_TBLSTUDIES2")

            'If intU = id_tblStudies Then
            'Else 'make new rs

            '    If cn.State = ADODB.ObjectStateEnum.adStateOpen Then
            '    Else
            '        cn.Open(constrCur)
            '    End If

            '    'find runid
            '    'str1 = "id_tblStudies = " & rows20(Count1).Item("id_tblStudies2")
            '    str1 = "id_tblStudies = " & intU 'rows20(Count1).Item("id_tblStudies2")
            '    Erase rows2
            '    rows2 = tblStudies.Select(str1)
            '    wRunID = rows2(0).Item("int_WatsonStudyID")
            '    If boolANSI Then
            '        str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, STUDY.STUDYNAME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME "
            '        str2 = "FROM (ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) INNER JOIN STUDY ON ANALYTICALRUNANALYTES.STUDYID = STUDY.STUDYID "
            '        str3 = "WHERE(((ASSAY.STUDYID)=" & wRunID & ") AND ((ANALYTICALRUN.RUNTYPEID) <> 3) And ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) <> 4)) "
            '        str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"
            '    Else
            '        str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, STUDY.STUDYNAME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME "
            '        str2 = "FROM ANALYTICALRUNANALYTES ," & strAnaRunPeak & " ,ASSAY ,ANALYTICALRUN ,ANALYTICALRUNSAMPLE , ANARUNANALYTERESULTS, STUDY "
            '        str2 = str2 & "WHERE ((((((ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) AND (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID)) AND (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) AND (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) AND ANALYTICALRUNANALYTES.STUDYID = STUDY.STUDYID "
            '        str3 = "AND ((ASSAY.STUDYID)=" & wRunID & ") AND (((ANALYTICALRUN.RUNTYPEID)<>3) AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)<>4)) "
            '        str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"
            '    End If
            '    strSQL = str1 & str2 & str3 & str4

            '    ''''''''''''console.writeline("FillAssignedSamplesDGV")
            '    ''''''''''''console.writeline(strSQL)

            '    rs.CursorLocation = CursorLocationEnum.adUseClient
            '    rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            '    rs.ActiveConnection = Nothing
            'End If

            strF1 = "ID_TBLSTUDIES = " & intU 'id_tblStudies
            strF1 = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLSTUDIES2 = " & intU
            Erase rows20
            rows20 = tbl2.Select(strF1) 'tblAssignedSamples
            int1 = rows20.Length
            var7 = GetWStudyID(intU)

            For Count1 = 0 To int1 - 1
                Erase rows1
                'var1 = dv3(Count1).Item("RUNID") ' dgv.Item("RUNID", Count1).Value ' 

                'int4 = dv3.Count

                var1 = rows20(Count1).Item("RUNID") ' dgv.Item("RUNID", Count1).Value ' 
                var2 = rows20(Count1).Item("ANALYTEINDEX") 'dgv.Item("ANALYTEINDEX", Count1).Value '
                var3 = rows20(Count1).Item("MASTERASSAYID") 'dgv.Item("MASTERASSAYID", Count1).Value ' 
                var4 = rows20(Count1).Item("RUNSAMPLESEQUENCENUMBER") 'dgv.Item("RUNSAMPLESEQUENCENUMBER", Count1).Value ' 
                var5 = rows20(Count1).Item("BOOLINTSTD") 'dgv.Item("RUNSAMPLESEQUENCENUMBER", Count1).Value ' 
                var6 = rows20(Count1).Item("CHARANALYTE")
                'var7 = rows20(Count1).Item("ANALYTEID")

                'var6 = rows20(Count1).Item("ID_TBLSTUDIES")
                'var7 = GetWStudyID(var6)

                'var5 = dgv.Item("BOOLINTSTD", Count1).Value ' dv2(Count1).Item("BOOLINTSTD")'

                'strF = "RUNID = '" & var1 & "' AND ANALYTEINDEX = '" & var2 & "' AND MASTERASSAYID = '" & var3 & "' AND RUNSAMPLESEQUENCENUMBER = '" & var4 & "' AND STUDYID = " & var7 & " AND CHARANALYTE = '" & var6 & "'"

                If var5 = -1 Then
                    strF = "RUNID = " & var1 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & " AND STUDYID = " & var7 & " AND INTERNALSTDNAME = '" & CleanText(CStr(var6)) & "'"
                Else
                    strF = "RUNID = " & var1 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & " AND STUDYID = " & var7 & " AND CHARANALYTE = '" & CleanText(CStr(var6)) & "'"
                End If
                '20180126 LEE: the previous strF's are incorrect
                'If IntStd (var5 = -1), then strF should be as below
                'If var5 = -1 Then
                '    strF = "RUNID = " & var1 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & " AND STUDYID = " & var7 & " AND CHARANALYTE = '" & var6 & "'"
                'Else
                '    strF = "RUNID = " & var1 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & " AND STUDYID = " & var7 & " AND CHARANALYTE = '" & var6 & "'"
                'End If

                rows1 = tbl1.Select(strF) 'tblAnalysisResultsHome
                int3 = rows1.Length

                'Dim varTT
                If int3 = 0 Then
                    GoTo next1
                Else
                End If

                'varTT = rows1(0).Item("INTERNALSTDNAME")'debugging

                num1 = CLng(rows20(Count1).Item("id_tblStudies"))
                num2 = CLng(rows20(Count1).Item("id_tblStudies2"))
                num3 = id_tblStudies 'CLng(Me.txtStudyID.Text)

                If InStr(1, var6, "Ver", CompareMethod.Text) > 0 Then
                    var1 = var1
                End If

                rows20(Count1).BeginEdit() 'tblAssignedSamples

                '20190124 LEE: Hmm. Add columns to Assigned Samples isn't running
                'update boolexclsamplechk
                var1 = rows20(Count1).Item("BOOLEXCLSAMPLE")
                If IsDBNull(var1) Then
                    rows20(Count1).Item("BOOLEXCLSAMPLE") = 0
                    rows20(Count1).Item("BOOLEXCLSAMPLECHK") = False
                Else
                    If var1 = -1 Then
                        rows20(Count1).Item("BOOLEXCLSAMPLECHK") = True
                    Else
                        rows20(Count1).Item("BOOLEXCLSAMPLECHK") = False
                    End If
                End If
         

                num2 = num3
                If num2 = num3 Then 'data from current study can be used

                    For Count2 = 0 To int2 - 1
                        'str1 = dgv2.Columns.item(Count2).Name
                        str1 = tbl1.Columns.Item(CInt(Count2)).ColumnName

                        If InStr(1, str1, "GROUP", CompareMethod.Text) > 0 And var5 = -1 Then
                            str1 = str1
                            Dim vAA, vBB
                            vAA = NZ(rows1(0).Item(str1), "") '20190130 LEE: Added NZ
                            vBB = NZ(rows1(0).Item("CHARANALYTE"), "") '20190130 LEE: Added NZ
                            If StrComp(vBB, "d3 IS", CompareMethod.Text) = 0 Then
                                vAA = vAA
                            End If
                        End If

                        boolAdd = True
                        boolFormat = False
                        '20180126 LEE: Problem here:
                        'INTGROUP was previously added to TBLASSIGNSAMPLES
                        'this value should not be replaced by tblAnalysisResultsHome
                        'add CASE "INTGROUP" to select statement below
                        Select Case str1
                            Case "RUNID"
                                boolAdd = False
                            Case "ANALYTEINDEX"
                                boolAdd = False
                            Case "MASTERASSAYID"
                                boolAdd = False
                            Case "RUNSAMPLESEQUENCENUMBER"
                                boolAdd = False
                            Case "CHARANALYTE"
                                boolAdd = False
                            Case "CONCENTRATION"
                                boolFormat = True
                            Case "INTGROUP"
                                boolAdd = False
                        End Select
                        If boolAdd Then
                            var2 = rows1(0).Item(str1) 'tblAnalysisResultsHome

                            'dgv.Item(str1, Count1).Value = var2
                            rows20(Count1).Item(str1) = var2
                            'dgv.Rows.item(Count1).Cells(str1).Value = rows1(0).Item(str1)
                            'dgv.Update()
                        End If
                        If boolFormat Then
                            'var2 = CDec(NZ(rows1(0).Item(str1), 0))
                            '20160510 LEE: Aack! Do not set to 0!!!
                            var2 = rows1(0).Item(str1) 'tblAnalysisResultsHome
                            'dgv.Item(str1, Count1).Value = var2
                            Try
                                rows20(Count1).Item(str1) = var2
                            Catch ex As Exception
                                var2 = var2
                            End Try

                        End If
                        'var1 = dgv.Item(str1, Count1).Value
                    Next

                Else 'data must be retrieved from a different Watson study
                    'Dim cn As New ADODB.Connection
                    'If cn.State = ADODB.ObjectStateEnum.adStateOpen Then
                    'Else
                    '    cn.Open(constrCur)
                    'End If
                    'Dim rs As New ADODB.Recordset

                    ''find runid
                    'str1 = "id_tblStudies = " & rows20(Count1).Item("id_tblStudies2")
                    'Erase rows2
                    'rows2 = tblStudies.Select(str1)
                    'wRunID = rows2(0).Item("int_WatsonStudyID")
                    'If boolANSI Then
                    '    str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, STUDY.STUDYNAME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME "
                    '    str2 = "FROM (ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) INNER JOIN STUDY ON ANALYTICALRUNANALYTES.STUDYID = STUDY.STUDYID "
                    '    str3 = "WHERE(((ASSAY.STUDYID)=" & wRunID & ") AND ((ANALYTICALRUN.RUNTYPEID) <> 3) And ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) <> 4)) "
                    '    str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"
                    'Else
                    '    str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, STUDY.STUDYNAME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME "
                    '    str2 = "FROM ANALYTICALRUNANALYTES ," & strAnaRunPeak & " ,ASSAY ,ANALYTICALRUN ,ANALYTICALRUNSAMPLE , ANARUNANALYTERESULTS, STUDY "
                    '    str2 = str2 & "WHERE ((((((ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) AND (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID)) AND (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) AND (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) AND ANALYTICALRUNANALYTES.STUDYID = STUDY.STUDYID "
                    '    str3 = "AND ((ASSAY.STUDYID)=" & wRunID & ") AND (((ANALYTICALRUN.RUNTYPEID)<>3) AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)<>4)) "
                    '    str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"
                    'End If
                    'strSQL = str1 & str2 & str3 & str4

                    '''''''''''''console.writeline("FillAssignedSamplesDGV")
                    '''''''''''''console.writeline(strSQL)

                    'rs.CursorLocation = CursorLocationEnum.adUseClient
                    'rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    'rs.ActiveConnection = Nothing

                    'rs.Filter = ""
                    'rs.Filter = strF
                    'Dim intLL As Int16
                    'intLL = rs.RecordCount

                    'If rs.EOF And rs.BOF Then
                    'Else
                    '    For Count3 = 0 To rs.Fields.Count - 1
                    '        var2 = rs.Fields(Count3).Value
                    '        str1 = rs.Fields(Count3).Name
                    '        boolAdd = True
                    '        boolFormat = False
                    '        Select Case str1
                    '            Case "RUNID"
                    '                boolAdd = False
                    '            Case "ANALYTEINDEX"
                    '                boolAdd = False
                    '            Case "MASTERASSAYID"
                    '                boolAdd = False
                    '            Case "RUNSAMPLESEQUENCENUMBER"
                    '                boolAdd = False
                    '            Case "CONCENTRATION"
                    '                boolFormat = True

                    '        End Select
                    '        If boolAdd Then
                    '            'dgv.Item(str1, Count1).Value = rs.Fields(str1).Value
                    '            var2 = rs.Fields(str1).Value 'debugging
                    '            rows20(Count1).Item(str1) = var2 'rs.Fields(str1).Value
                    '        End If
                    '        If boolFormat Then
                    '            var2 = CDec(NZ(rs.Fields(str1).Value, 0))
                    '            rows20(Count1).Item(str1) = var2
                    '            'dgv.Item(str1, Count1).Value = var2
                    '        End If
                    '    Next
                    'End If

                End If
                rows20(Count1).EndEdit()
next1:
            Next

            'If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
            '    rs.Close()
            'End If

        Next

        'rs = Nothing
        'If cn.State = ADODB.ObjectStateEnum.adStateOpen Then
        '    cn.Close()
        'End If
        'cn = Nothing

        tblAssignedSamples.AcceptChanges()
        'dgv.ReadOnly = True

        ''debug:  lisst(tblAssignedSamples)
        ''''''''console.writeline("Start FillAssignedSamplesDVG")
        'For Count1 = 0 To rows20.Length - 1
        '    var1 = ""
        '    If Count1 = 0 Then
        '        For Count2 = 0 To tblAssignedSamples.Columns.Count - 1
        '            var2 = tblAssignedSamples.Columns(Count2).ColumnName
        '            var1 = var1 & ";" & var2
        '        Next
        '        '''''''console.writeline(var1)
        '        var1 = ""
        '    End If
        '    For Count2 = 0 To tblAssignedSamples.Columns.Count - 1
        '        var2 = rows20(Count1).Item(Count2)
        '        var1 = var1 & ";" & NZ(var2, "NULL")
        '    Next
        '    '''''''console.writeline(var1)
        'Next
        ''''''''console.writeline("End FillAssignedSamplesDVG")

    End Sub


    Sub OpenAssignedSamples(boolViewOnly As Boolean)

        Dim var1
        Dim str1 As String

        If BOOLASSIGNSAMPLES Or boolViewOnly Then
        Else
            MsgBox("User not allowed to Assign Samples.", MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        var1 = NZ(frmH.cbxStudy.Text, "")
        If Len(var1) = 0 Then
            MsgBox("First choose a study.", MsgBoxStyle.Information, "Choose a study...")
            Exit Sub
        End If

        If frmH.cmdEdit.Enabled Or boolViewOnly Then
        Else
            str1 = "Please save the study before assigning samples."
            MsgBox(str1, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        Cursor.Current = Cursors.WaitCursor

        'evaluate dgv
        Dim dv As System.Data.DataView
        Dim intRows As Short
        Dim boolGo As Boolean
        Dim Count1 As Short
        Dim bool1 As Short
        Dim bool2 As Short
        Dim int1 As Short
        Dim intRow As Short
        Dim dgv As DataGridView

        dgv = frmH.dgvReportTableConfiguration
        dv = dgv.DataSource
        intRows = dv.Count
        boolGo = False
        For Count1 = 0 To intRows - 1
            bool1 = dv(Count1).Item("BOOLINCLUDE")
            If bool1 = -1 Then 'evaluate further
                bool2 = dv(Count1).Item("boolRequiresSampleAssignment")
                If bool2 = -1 Then
                    boolGo = True
                    Exit For
                End If
            End If

        Next

        If boolViewOnly Then
            boolGo = True
        End If

        'boolGo = True

        If boolGo Then
            Dim frm As New frmAssignSamples
            frm.boolFormLoad = True
            frm.panLabels.Visible = True
            'frm.txtStudy.Text = cbxStudy.Text
            frm.Text = "  Assign samples for study " & var1 & "..."
            frm.txtStudyID.Text = id_tblStudies
            frm.boolViewOnly = boolViewOnly
            'Call frm.PlaceControls(boolViewOnly)

            'If boolViewOnly Then
            '    frm.panAccCrit.Visible = False
            '    frm.panLabels.Visible = False
            '    frm.panCbxStudy.Visible = False
            '    frm.panExtra.Visible = False
            '    frm.cmdEdit.Visible = False
            '    frm.cmdOK.Visible = False
            '    frm.cmdReset.Visible = False
            '    frm.dgvTables.Visible = False
            '    frm.lblTables.Visible = False
            '    frm.lblColored.Visible = False
            '    Dim a, b, c, d
            '    a = frm.panAssSamples.Top + frm.panAssSamples.Height
            '    b = frm.dgvAnalyticalRuns.Top
            '    frm.dgvAnalyticalRuns.Height = a - b
            'End If

            frm.ShowDialog()
            frmH.Refresh()
            'frm.Dispose()
            frm.Close()

            '20190130 LEE
            'why is this form not being disposed?
            frm.Dispose()

            Cursor.Current = Cursors.WaitCursor
            Call AssessSampleAssignment()
            Cursor.Current = Cursors.WaitCursor
            Call AssessQCs()
            Cursor.Current = Cursors.WaitCursor
        Else
            MsgBox("There are no tables that require sample assignment", MsgBoxStyle.Information, "No need...")
        End If

        frmH.Focus()
        frmH.dgvReportStatements.Focus()

        Cursor.Current = Cursors.Default

    End Sub


    Sub ViewAnalRuns()

        Dim var1
        var1 = NZ(frmH.cbxStudy.Text, "")
        If Len(var1) = 0 Then
            MsgBox("First choose a study.", MsgBoxStyle.Information, "Choose a study...")
            Exit Sub
        End If

        Cursor.Current = Cursors.WaitCursor

        'evaluate dgv
        Dim dv As System.Data.DataView
        Dim intRows As Short
        Dim boolGo As Boolean
        Dim Count1 As Short
        Dim bool1 As Short
        Dim bool2 As Short
        Dim int1 As Short

        'dv = frmH.dgvReportTableConfiguration.DataSource
        'intRows = dv.Count
        'boolGo = False
        'For Count1 = 0 To intRows - 1
        '    bool1 = dv(Count1).Item("Include")
        '    If bool1 = -1 Then 'evaluate further
        '        bool2 = dv(Count1).Item("boolRequiresSampleAssignment")
        '        If bool2 = -1 Then
        '            boolGo = True
        '            Exit For
        '        End If
        '    End If

        'Next

        Dim frm As New frmAssignSamples

        frm.boolViewOnly = True
        'frm.txtStudy.Text = frmH.cbxStudy.Text
        frm.Text = "  Review samples for study " & var1 & "..."
        frm.txtStudyID.Text = id_tblStudies
        frm.lblColored.Visible = False

        'hide a bunch of stuff
        'frm.lblWait.Top = 0
        'frm.lblWait.Left = 0
        'frm.lblWait.Width = frm.Width
        'frm.lblWait.Height = frm.lbl1.Top ' - frm.lbl1.Height
        'frm.lblWait.Text = ""
        'frm.lblWait.Visible = True
        'frm.lblWait.BringToFront()
        'frm.cmdExit.BringToFront()

        Dim t1, t2, h1


        frm.cmdReturn.Visible = False
        frm.lbl2.Visible = False
        frm.cbxStudy.Visible = False

        frm.lbldgvNomConc.Visible = False
        frm.dgvNomConc.VirtualMode = False

        't1 = frm.dgvAnalyticalRuns.Top
        't2 = frm.dgvAssignedSamples.Top + frm.dgvAssignedSamples.Height
        'h1 = t2 - t1

        'frm.dgvAnalyticalRuns.Height = h1 ' frm.Height - (frm.dgvAnalyticalRuns.Top + 100)
        'frm.dgvAnalyticalRuns.Width = frm.Width - frm.dgvAnalyticalRuns.Left - 40
        'frm.dgvAnalyticalRuns.BringToFront()

        frm.cmdEdit.Visible = False
        frm.cmdOK.Visible = False
        frm.cmdReset.Visible = False

        'frm.lbl3.BringToFront()
        'frm.dgvAnalytes.BringToFront()

        frm.panLabels.Visible = False

        frm.ShowDialog()
        'Me.Refresh()

        frm.Close()

        '20190130 LEE:
        'why isn't this form being disposed?
        frm.Dispose()

        Cursor.Current = Cursors.Default
    End Sub

    Sub CleanUpDirs()

        Dim strPath As String
        Dim var1
        Dim dt As Date
        Dim dtNow As Date
        Dim dtCheck As Date

        strPath = "C:\Labintegrity\StudyDoc\Temp\"

        dtNow = Now

        For Each foundFile As String In My.Computer.FileSystem.GetFiles(strPath)
            dt = FileDateTime(foundFile)
            dtCheck = DateAdd("d", 14, dt)
            If dtCheck > dtNow Then 'ignore
            Else 'delete
                Try
                    File.Delete(foundFile)
                Catch ex As Exception
                End Try
            End If
        Next

        'now do TempReports
        strPath = "C:\Labintegrity\StudyDoc\TempReport\"

        If Directory.Exists(strPath) Then
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(strPath)
                dt = FileDateTime(foundFile)
                dtCheck = DateAdd("d", 14, dt)
                If dtCheck > dtNow Then 'ignore
                Else 'delete
                    Try
                        File.Delete(foundFile)
                    Catch ex As Exception
                    End Try
                End If
            Next
        End If


    End Sub

    Sub FillHeaderFooterTable()
        Dim h

        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim strS As String
        Dim strFM As String
        Dim strSM As String
        Dim str1 As String
        Dim var1, var2
        Dim Count1 As Short

        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTS = " & id_tblReports
        strS = "ID_TBLREPORTS ASC"
        dtbl = tblReportHeaders

        rows = dtbl.Select(strF, strS)

        If rows.Length = 0 Then 'add a new record
            Dim rowsM() As DataRow
            Dim tbl As System.Data.DataTable
            Dim maxID As Int64

            maxID = GetMaxID("tblReportHeaders", 1, True) 'if maxid increment is 1, then getmaxid already does putmaxid
            'Call PutMaxID("tblReportHeaders", maxID)

            'If boolGuWuOracle Then
            '    ta_tblMaxID.Fill(tblMaxID)
            'ElseIf boolGuWuAccess Then
            '    ta_tblMaxIDAcc.Fill(tblMaxID)
            'ElseIf boolGuWuSQLServer Then
            '    ta_tblMaxIDSQLServer.Fill(tblMaxID)
            'End If

            'tbl = tblMaxID
            'strFM = "charTable = 'tblReportHeaders'"
            'rowsM = tbl.Select(strFM)
            'maxID = rowsM(0).Item("NUMMAXID")
            'maxID = maxID + 1
            'rowsM(0).BeginEdit()
            'rowsM(0).Item("NUMMAXID") = maxID
            'rowsM(0).EndEdit()

            ''tbl.AcceptChanges()
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

            'enter data
            Dim newRow As DataRow = dtbl.NewRow
            newRow.BeginEdit()
            For Count1 = 1 To 8
                Select Case Count1
                    Case 1
                        str1 = "CHARHLT"
                        var2 = "Project Number: [CORPORATESTUDY/PROJECTNUMBER]"
                    Case 2
                        str1 = "CHARHRT"
                        var2 = "Page [PAGENUMBER]"
                    Case 3
                        str1 = "CHARHLB"
                        var2 = "Final Report"
                    Case 4
                        str1 = "CHARHRB"
                        var2 = "For [SUBMITTEDTO]"
                    Case 5
                        str1 = "CHARFLT"
                        var2 = System.DBNull.Value
                    Case 6
                        str1 = "CHARFRT"
                        var2 = System.DBNull.Value
                    Case 7
                        str1 = "CHARFLB"
                        var2 = System.DBNull.Value
                    Case 8
                        str1 = "CHARFRB"
                        var2 = System.DBNull.Value
                End Select

                newRow.Item(str1) = var2

                ''reselect rows
                'Erase rows
                'rows = dtbl.Select(strF, strS)

            Next


            'BOOLDIFFFIRSTPAGE
            'BOOLINCLUDELOGO
            str1 = "BOOLDIFFFIRSTPAGE"
            newRow.Item(str1) = -1
            str1 = "BOOLINCLUDELOGO"
            newRow.Item(str1) = 0

            'ID_TBLREPORTHEADERS
            'ID_TBLREPORTS
            'ID_TBLSTUDIES
            newRow.Item("ID_TBLREPORTHEADERS") = maxID
            newRow.Item("ID_TBLREPORTS") = id_tblReports
            newRow.Item("ID_TBLSTUDIES") = id_tblStudies

            newRow.EndEdit()

            dtbl.Rows.Add(newRow)


            If boolGuWuOracle Then
                Try
                    ta_tblReportHeaders.Update(tblReportHeaders)
                Catch ex As DBConcurrencyException
                    ''msgbox("aaQA Tables: " & ex.Message)
                    'ds2005.TBLREPORTHEADERS.Merge('ds2005.TBLREPORTHEADERS, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblReportHeadersAcc.Update(tblReportHeaders)
                Catch ex As DBConcurrencyException
                    ''msgbox("aaQA Tables: " & ex.Message)
                    'ds2005Acc.TBLREPORTHEADERS.Merge('ds2005Acc.TBLREPORTHEADERS, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblReportHeadersSQLServer.Update(tblReportHeaders)
                Catch ex As DBConcurrencyException
                    ''msgbox("aaQA Tables: " & ex.Message)
                    'ds2005Acc.TBLREPORTHEADERS.Merge('ds2005Acc.TBLREPORTHEADERS, True)
                End Try
            End If

            'reselect rows
            Erase rows
            rows = dtbl.Select(strF, strS)

        End If
    End Sub

    Sub FillPatientsCheck(ByVal dgv As DataGridView)

        Dim Count1 As Short
        Dim dv As System.Data.DataView
        Dim intRows As Short
        Dim int1 As Short
        Dim bool As Boolean
        Dim boolE As Boolean

        Try
            dv = dgv.DataSource
        Catch ex As Exception
            Exit Sub
        End Try

        Try
            intRows = dv.Count
        Catch ex As Exception
            Exit Sub
        End Try

        If intRows = 0 Then
            Exit Sub
        End If

        boolE = dv.AllowEdit
        dv.AllowEdit = True
        For Count1 = 0 To intRows - 1
            int1 = dv(Count1).Item("BOOLTERMINALBLEED")
            bool = False
            If int1 = -1 Then
                bool = True
            Else
                bool = False
            End If
            dv(Count1).BeginEdit()
            dv(Count1).Item("BOOLTERMINAL") = bool
            dv(Count1).EndEdit()
        Next
        dv.AllowEdit = boolE

    End Sub

    Sub ModMethChoice()

        Dim str1 As String
        Dim boolArchive As Boolean

        boolArchive = False
        If frmH.rbArchive.Checked Then
            boolArchive = True
        Else
            boolArchive = False
        End If

        If boolArchive Then
            str1 = "2. Retrieve information from an Existing StudyDoc-configured Watson Validation Study archived .mdb file listed in the dropdown box below or Browse..."

        Else
            str1 = "2. Retrieve information from an Existing StudyDoc-configured Watson Validation Study by choosing a study from the dropdown box below"

        End If

        str1 = "2. Retrieve information from an Existing StudyDoc-configured Watson Validation Study by choosing a study from the dropdown box below"

        frmH.lblM2.Text = str1

        frmH.cbxArchivedMDB.Top = frmH.cbxMethValExistingGuWu.Top
        frmH.cmdBrowseMDB.Top = frmH.cbxArchivedMDB.Top

        If boolArchive Then
            frmH.cbxMethValExistingGuWu.Visible = False
            frmH.cbxArchivedMDB.Visible = True
            frmH.cmdBrowseMDB.Visible = True
        Else
            frmH.cbxMethValExistingGuWu.Visible = True
            frmH.cbxArchivedMDB.Visible = False
            frmH.cmdBrowseMDB.Visible = False
        End If

        frmH.cbxMethValExistingGuWu.Visible = True
        frmH.cbxArchivedMDB.Visible = False
        frmH.cmdBrowseMDB.Visible = False




    End Sub

    Sub FillArchivedMDB()

        Exit Sub

        Dim strPath As String
        Dim arr1(100)


        strPath = "C:\Labintegrity\StudyDoc\ArchivedMDBs\"
        strPath = "C:\LabIntegrity\StudyDoc\ArchivedMDBs\"

        frmH.cbxArchivedMDB.Items.Clear()

        If System.IO.Directory.Exists(strPath) Then
        Else
            Exit Sub
        End If

        For Each fi As String In My.Computer.FileSystem.GetFiles(strPath, FileIO.SearchOption.SearchTopLevelOnly, "*.mdb")

            frmH.cbxArchivedMDB.Items.Add(strPath & fi)

        Next

        frmH.cbxArchivedMDB.DropDownWidth = frmH.gbMethValApplyGuWu.Width - frmH.cbxArchivedMDB.Left - 10


    End Sub



End Module


