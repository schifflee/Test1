Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.ComponentModel.PropertyDescriptorCollection
Imports Word = Microsoft.Office.Interop.Word
Imports Microsoft.VisualBasic
Imports System.IO

Module modRepeatAnalysis

    Sub SRSummaryRepeatSamples_7(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal idTR As Int64)


        'This table uses a number of database tables:
        'tblRepeatAllRunSamples - A table with a set of entries with all runsamples (including duplicates) for each 
        '                         analyteid/sample that has been reassayed (All Analytes, All Matrices)
        'tblRepeatTableRows - A table with a single set of entries for each row of the repeat table.
        'tblDuplicates - A table with the original and all the duplicate entries for each row of the repeat table

        'NDL Note: If we want to include *every* entry in SAMPRESCONFLICTDEC, and then in another routine, filter out by the latest timepoint, we could use:
        'SELECT DISTINCT DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.TREATMENTEVENTID, SAMPLERESULTS.ACCEPTANCETIMESTAMP, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, SAMPLERESULTS.CONCENTRATION, ANARUNANALYTERESULTS.CONCENTRATION, ASSAY.SAMPLETYPEKEY, ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER, ANARUNANALYTERESULTS.RUNID, SAMPLERESULTS.DESIGNSAMPLEID, ANALYTICALRUN.STUDYID, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS
        'FROM SAMPLERESULTSCONFLICT AS SAMPLERESULTSCONFLICT_4, SAMPLERESULTSCONFLICT AS SAMPLERESULTSCONFLICT_5, ((ASSAYANALYTES INNER JOIN ANALYTICALRUNANALYTES ON (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX)) INNER JOIN (ANALYTICALRUN INNER JOIN ASSAY ON ANALYTICALRUN.ASSAYID = ASSAY.ASSAYID) ON (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID)) INNER JOIN ((SAMPRESCONFLICTDEC INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSAMPLE INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID) AND (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID)) ON (SAMPRESCONFLICTDEC.DESIGNSAMPLEID = SAMPLERESULTS.DESIGNSAMPLEID) AND (SAMPRESCONFLICTDEC.ANALYTEID = SAMPLERESULTS.ANALYTEID)) INNER JOIN (ANARUNANALYTERESULTS INNER JOIN ANALYTICALRUNSAMPLE ON (ANARUNANALYTERESULTS.RUNID = ANALYTICALRUNSAMPLE.RUNID) AND (ANARUNANALYTERESULTS.STUDYID = ANALYTICALRUNSAMPLE.STUDYID) AND (ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER = ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER)) ON SAMPRESCONFLICTDEC.DESIGNSAMPLEID = ANALYTICALRUNSAMPLE.DESIGNSAMPLEID) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (ASSAYANALYTES.ANALYTEID = SAMPRESCONFLICTDEC.ANALYTEID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNSAMPLE.STUDYID) AND (ANALYTICALRUN.RUNID = ANALYTICALRUNSAMPLE.RUNID)
        'WHERE (((SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3))
        'ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR;


        Dim str1, str2, str3, str4, strSQL, strTName, strF, strS As String
        Dim CountAnalyteID, CountMatrix As Short
        Dim Count2, Count3, Count4, Count5 As Short
        Dim Count1A As Int16
        Dim Count2A As Int16

        Dim boolPlaceHolderTemp As Boolean = False
        Dim boolDoPlaceHolderTemp As Boolean = False

        Dim intR As Integer

        Dim arrReassay(9, 100), arrReasons(100), arrReasonsC(100), arrRepeatRows(), arrRepeat(7, 100)
        '1=ANALYTEID, 2=DESIGNSUBJECTTAG, 3=SUBJECTGROUPNAME, 4=ENDDAY, 5=ENDHOUR, 6=RUNID, 7=DESIGNSAMPLEID
        Dim ctReassay, ctReassayRows, ctReasons, ctRepeatRows, ctLegend, ctRPool, ctRepeatTableRows As Int32
        Dim intDo As Short
        Dim dvDo As System.Data.DataView
        Dim int1, int2, int3, int4, intExp, intRow As Short
        Dim ctCols As Short
        Dim arrOrder(6, 100)
        Dim tbl, tbl1, tbl2 As System.Data.DataTable
        Dim dr(), dr1(), dr2(), dr3() As DataRow
        Dim var1, var2, var3, var4
        Dim numLLOQ, numULOQ As Decimal
        Dim strBQL, strAQL As String
        Dim lng1, lng2 As Int64
        Dim dv As System.Data.DataView
        Dim wrdselection As Microsoft.Office.Interop.Word.Selection
        Dim tblD As System.Data.DataTable
        Dim dvD As System.Data.DataView

        Dim rowSC() As DataRow
        Dim strFSC, strFld, strPaste, strPasteT, strTempInfo As String
        Dim dup1, dup2, numMeanDup, numOrig, numRV As Object
        Dim numSpaceAfter As Single
        Dim numSpaceAfterNew As Single

        '1=DESIGNSAMPLEID, 2=ANALYTEID, 3=DESIGNSUBJECTTAG, 4=TimePoint, 5=RUNID, 6=RUNSAMPLEORDERNUMBER, 7=DECISIONCODE
        Dim fld As ADODB.Field
        Dim fonts, fontsize
        Dim strConcReason, strAnalyteAndMatrixFilter, strAnalyteMatrixandDesignIDFilter As String

        Dim v1, v2, vU
        Dim strM, strM1

        Dim tblRepeatTableRows As New DataTable
        Dim dvRepeatAllRunSamples, dvRepeatTableRows As DataView
        Dim dvCalStdGroupAssayIDsAcc As New DataView(tblCalStdGroupAssayIDsAcc)
        Dim dvDuplicates As DataView
        Dim tblDuplicates As New DataTable
        Dim ctDuplicates As Short
        Dim boolFirstLine, boolFirstLineEntry As Boolean

        Dim tbl1A As DataTable
        Dim tbl2A As DataTable
        Dim strR As String = "_xyz_"
        Dim strR1 As String = "_abc_"

        Dim strSdvRepeatTableRowsSort As String = ""
        Dim strSdvRepeatAllRunSamplesSort As String = ""

        strSdvRepeatTableRowsSort = "DESIGNSUBJECTTAG ASC, WEEK ASC, ENDDAY ASC, ENDHOUR ASC, ENDMINUTE ASC, ENDSECOND ASC, SAMPLERESULTS_RUNID ASC"
        strSdvRepeatAllRunSamplesSort = strSdvRepeatTableRowsSort

        Dim boolBQL As Boolean = False
        Dim boolAQL As Boolean = False
        Dim boolNA As Boolean = False
        Dim boolNR As Boolean = False

        Dim strConcUnits As String
        Dim numDF As Decimal

        Dim strTNameO As String 'original Table Name

        Dim charFCID As String
        strF = "ID_TBLREPORTTABLE = " & idTR
        Dim rowsTR() As DataRow = tblReportTable.Select(strF)
        var1 = rowsTR(0).Item("CHARFCID")
        charFCID = NZ(var1, "NA")

        With wd

            fontsize = wd.ActiveDocument.Styles("Normal").Font.Size ' .Selection.Font.Size
            fonts = fontsize '.Selection.Font.Size

            Dim intTableID As Short
            intTableID = 7

            Dim strWRunId As String = GetWatsonColH(intTableID)

            dvDo = frmH.dgvReportTableConfiguration.DataSource
            intDo = FindRowDVNumByCol(intTableID, dvDo, "id_tblconfigreporttables")

            '***
            intDo = FindRowDVNumByCol(idTR, dvDo, "ID_TBLREPORTTABLE")

            'Get table name
            var1 = dvDo(intDo).Item("CHARHEADINGTEXT")
            strTNameO = NZ(var1, "[NONE]")

            'get Temperature info
            var1 = dvDo(intDo).Item("CHARSTABILITYPERIOD")
            strTempInfo = NZ(var1, "[NONE]")

            '***

            ctPB = ctPB + 1
            If ctPB > frmH.pb1.Maximum Then
                ctPB = 1
            End If
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()


            '20160313 LEE:
            'moved tblRepeatAllRunSamples to DoPrepare
            'this table is needed in Sample Conc table if Concentration is used

            '*** Start Duplicate Logic 1 ****

            '20160313 LEE:
            'moved tblRepeatAllRunSamples to DoPrepare
            'this table is needed in Sample Conc table if Concentration is used

            'Create Collapsed table - One entry per Decision [AnaRunAnalyteResults.Concentration, RunID, RunSequence, and OriginalValue 
            'not included, as they will differ between samples]
            'Note the SampleResults.Concentration is the concentration of the result only and not an individual sample per se 
            '(e.g. it may be the mean of the individual runsamples).  Same with SampleResults.CalibrationRangeFlag (shortened to CalibrationRangeFlag).

            dvRepeatAllRunSamples = New DataView(tblRepeatAllRunSamples)
            '20160219 LEE: 
            ' - there were duplicate instances of RUNID in orignal query,  I see you use both in different Unique tables, will leave for now
            ' - there were duplicate instances of ALIQUOTFACTOR in original query, I see you use both in different Unique tables, will leave for now

            'tblRepeatTableRows = dvRepeatAllRunSamples.ToTable("tblRepeatTableRows", True, "ANALYTEID", "DESIGNSUBJECTTAG", "DESIGNSAMPLEID", "SUBJECTGROUPNAME",
            '                                                   "USERSAMPLEID", "TREATMENTEVENTID", "ENDDAY", "ENDHOUR", "ENDMINUTE", "ENDSECOND", "STUDYID",
            '                                                   "REASSAYREASON", "REASSAYCONCREASON", "DECISIONCODE", "SAMPLETYPEID", "SAMPLERESULTS.CONCENTRATION",
            '                                                   "CALIBRATIONRANGEFLAG", "SAMPLERESULTS.ALIQUOTFACTOR", "SAMPLERESULTS.RUNID", "WEEK", "VISITTEXT")


            tblRepeatTableRows = dvRepeatAllRunSamples.ToTable("tblRepeatTableRows", True, "ANALYTEID", "DESIGNSUBJECTTAG", "DESIGNSAMPLEID", "SUBJECTGROUPNAME",
            "USERSAMPLEID", "TREATMENTEVENTID", "ENDDAY", "ENDHOUR", "ENDMINUTE", "ENDSECOND", "STUDYID",
            "REASSAYREASON", "REASSAYCONCREASON", "DECISIONCODE", "SAMPLETYPEID", "SAMPLERESULTS_CONCENTRATION",
            "CALIBRATIONRANGEFLAG", "SAMPLERESULTS_ALIQUOTFACTOR", "SAMPLERESULTS_RUNID", "WEEK", "VISITTEXT", "STARTDAY", "STARTHOUR", "STARTMINUTE", "STARTSECOND")

            ''20180219 LEE: added "ARS_ALIQUOTFACTOR"
            'tblRepeatTableRows = dvRepeatAllRunSamples.ToTable("tblRepeatTableRows", True, "ANALYTEID", "DESIGNSUBJECTTAG", "DESIGNSAMPLEID", "SUBJECTGROUPNAME",
            '                                       "USERSAMPLEID", "TREATMENTEVENTID", "ENDDAY", "ENDHOUR", "ENDMINUTE", "ENDSECOND", "STUDYID",
            '                                       "REASSAYREASON", "REASSAYCONCREASON", "DECISIONCODE", "SAMPLETYPEID", "SAMPLERESULTS_CONCENTRATION",
            '                                       "CALIBRATIONRANGEFLAG", "SAMPLERESULTS_ALIQUOTFACTOR", "ARS_ALIQUOTFACTOR", "SAMPLERESULTS_RUNID", "WEEK", "VISITTEXT", "STARTDAY", "STARTHOUR", "STARTMINUTE", "STARTSECOND")



            dvRepeatTableRows = New DataView(tblRepeatTableRows)
            'need to sort here since original recordset sort doesn't seem to do the trick consistently
            dvRepeatTableRows.Sort = strSdvRepeatTableRowsSort

            '*** End Duplicate Logic 1 ****


            '20160217 LEE: Keep order the same as Report Table Config page:
            'strS = "MATRIX ASC, ANALYTEDESCRIPTION ASC, ANALYTEDESCRIPTION_C ASC, INTGROUP ASC"
            'but this may change if we add functionality to let user modify sort order

            '20160218 LEE: Use this for testing order
            'Dim intR As Short
            'strM = "Enter OK for 'Matrix'"
            'intR = MsgBox(strM, vbOKCancel)
            'If intR = 1 Then
            '    gSortAnalytes = "Matrix"
            'Else
            '    gSortAnalytes = "Analyte"
            'End If

            Dim boolM As Boolean = False
            If StrComp(gSortAnalytes, "Matrix", CompareMethod.Text) = 0 Then
                tbl1A = tblMatrices
                tbl2A = tblAnalyteIDs
                boolM = True
            Else
                tbl1A = tblAnalyteIDs
                tbl2A = tblMatrices
                boolM = False
            End If

            Dim strMatrix As String
            Dim strAnalyteID, strAnalyteDescription As String

            '20180216 LEE: Don't need matrix/calibr/analyte grouping anymore
            'user is now able to sort compounds as desired
            'but still want to combine analtyes with differing calibration curves
            'make a new analyte table without groups
            Dim dv11 As New DataView(tblAnalytesHome, "IsIntStd = 'No'", "INTORDER ASC, OriginalAnalyteDescription ASC", DataViewRowState.CurrentRows)
            Dim tbl11 As DataTable = dv11.ToTable("a", True, "AnalyteID", "OriginalAnalyteDescription", "Matrix", "ConcUnits", "IsIntStd")

            strF = "IsIntStd = 'No'"
            strS = "INTORDER ASC" ', IsIntStd ASC, OriginalAnalyteDescription ASC"
            Dim rows11() As DataRow = tbl11.Select()
            Dim intRowsAnal As Short = rows11.Length

            For Count1A = 0 To 0 '20180219 LEE' tbl1A.Rows.Count - 1 'Iterate through each Matrix (but keep different calibration ranges together)

                'strTName = strTNameO 'reset strTName

                If boolM Then
                    strMatrix = tblMatrices.Rows(Count1A).Item("Matrix")
                Else
                    strAnalyteID = tblAnalyteIDs.Rows(Count1A).Item("AnalyteID")
                    strAnalyteDescription = tblAnalyteIDs.Rows(Count1A).Item("AnalyteDescription")
                End If


                For Count2A = 0 To intRowsAnal - 1 '20180219 LEE' tbl2A.Rows.Count - 1 'Iterate through each AnalyteID, and generate the information

                    '20171128 LEE:
                    strTName = strTNameO 'reset strTName

                    boolPlaceHolderTemp = False
                    boolDoPlaceHolderTemp = False

                    Dim arrLegend(4, 1000) 'Reason for Reassay

                    'If boolM Then
                    '    strAnalyteID = tblAnalyteIDs.Rows(Count2A).Item("AnalyteID")
                    '    strAnalyteDescription = tblAnalyteIDs.Rows(Count2A).Item("AnalyteDescription")
                    'Else
                    '    strMatrix = tblMatrices.Rows(Count2A).Item("Matrix")
                    'End If

                    ''20160311 LEE:
                    ''get units from tblanalyteshome
                    'Dim strF1 As String
                    'strF1 = "ANALYTEID = " & strAnalyteID & " AND MATRIX = '" & strMatrix & "'"
                    'Dim rowsUnits() As DataRow = tblAnalytesHome.Select(strF1)
                    'strConcUnits = rowsUnits(0).Item("ConcUnits")

                    '20180216 LEE: 
                    Try
                        strAnalyteID = rows11(Count2A).Item("AnalyteID")
                        strAnalyteDescription = rows11(Count2A).Item("OriginalAnalyteDescription")
                        strMatrix = rows11(Count2A).Item("Matrix")
                        strConcUnits = rows11(Count2A).Item("ConcUnits")
                        'intGroup = rows11(Count2A).Item("INTGROUP")
                    Catch ex As Exception
                        var1 = var1 'debug
                    End Try

                    If (Not (boolGenerateTableForThisAnalyteIDandMatrix(intDo, strAnalyteID, strMatrix))) Then
                        GoTo next1
                    End If

                    intTCur = intTCur + 1

                    '**This may no longer be needed - not used anywhere.
                    Dim strSString As String
                    Dim strGroupCheck As String
                    Call GetGroupSort(idTR) 'retrieve grouping and sorting information

                    strSString = GetSString()

                    If intGroups = 0 Then
                        strGroupCheck = "[None]"
                    Else
                        strGroupCheck = arrGroups(1, 1)
                    End If
                    '**

                    'Select for Matrix and AnalyteID
                    strAnalyteAndMatrixFilter = "SAMPLETYPEID = '" & strMatrix & "' AND ANALYTEID = " & strAnalyteID
                    dvRepeatTableRows.RowFilter = strAnalyteAndMatrixFilter
                    'need to sort here since original recordset sort doesn't seem to do the trick consistently
                    dvRepeatTableRows.Sort = strSdvRepeatTableRowsSort

                    'SKIP this AnalyteID if there are no reassays in this matrix to report.
                    If (dvRepeatTableRows.Count = 0) Then
                        GoTo Next1
                    End If

                    ctRepeatTableRows = dvRepeatTableRows.Count

                    strM = "Creating " & strTName & " For " & strAnalyteDescription & "..."
                    strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    strM1 = strM
                    frmH.lblProgress.Text = strM
                    frmH.Refresh()

                    'page setup according to configuration
                    str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")

                    'insert page break
                    Call InsertPageBreak(wd)
                    '.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                    Call PageSetup(wd, str1) 'L=Landscape, P=Portrait

                    If boolPlaceHolder Then
                        GoTo skipPlaceholder
                    End If

                    'determine number of columns and order of columns
                    tbl1 = tblReportTableHeaderConfig
                    strF = "id_tblStudies = " & id_tblStudies & " AND id_tblConfigReportTables = " & intTableID & " AND boolInclude = -1"
                    dr1 = tbl1.Select(strF, "intOrder ASC")
                    int1 = dr1.Length
                    ctCols = int1

                    '20160216 LEE: determine if there are more than one unique weeks
                    strF = "WEEK IS NOT NULL"
                    Dim dvWeeks As DataView = New DataView(tblRepeatAllRunSamples, strF, "", DataViewRowState.CurrentRows)
                    Dim tblWeeks As DataTable = dvWeeks.ToTable("a", True, "WEEK")
                    Dim boolDoWeeks As Boolean = False
                    If tblWeeks.Rows.Count > 1 Then
                        boolDoWeeks = True
                    End If


                    '1=ColumnHeader, 2=Include(X), 3=Order, 4=ReportColumnHeader
                    'arrOrder
                    ' 1=ColumnLabel, 2=Order, 3=id_tblConfigReportTables, 4=UserLabel, 5=id_tblConfigHeaderLookup, 6=WatsonField
                    tbl = tblConfigHeaderLookup
                    arrOrder.Clear(arrOrder, 0, arrOrder.Length)
                    int1 = 0
                    For Count2 = 1 To ctCols  'For each item to be reported in table, paste the column information (header title, etc) into arrOrder

                        int1 = int1 + 1

                        arrOrder(2, int1) = int1 'dr1(int1 - 1).Item("intOrder")
                        arrOrder(3, int1) = dr1(Count2 - 1).Item("id_tblConfigHeaderLookup")
                        arrOrder(4, int1) = dr1(Count2 - 1).Item("charUserLabel")
                        str2 = dr1(Count2 - 1).Item("charUserLabel") 'just checking
                        arrOrder(5, int1) = dr1(Count2 - 1).Item("id_tblConfigReportTables")
                        'find column label
                        str1 = "id_tblConfigHeaderLookup = " & arrOrder(3, int1) & " AND id_tblConfigReportTables = " & intTableID & ""
                        dr = tbl.Select(str1)
                        arrOrder(1, int1) = dr(0).Item("CHARCOLUMNLABEL")
                        arrOrder(6, int1) = dr(0).Item("CHARWATSONFIELD")

                        var1 = arrOrder(6, int1)

                        If StrComp(var1.ToString, "ENDDAY", CompareMethod.Text) = 0 And boolDoWeeks Then 'add another column

                            var2 = dr1(Count2 - 1).Item("id_tblConfigReportTables")
                            int1 = int1 + 1

                            'first increment day up
                            For Count3 = 1 To 6
                                arrOrder(Count3, int1) = arrOrder(Count3, int1 - 1)
                            Next
                            arrOrder(2, int1) = int1 'set the order + 1

                            arrOrder(2, int1 - 1) = int1 - 1 'dr1(int1 - 1).Item("intOrder")
                            arrOrder(3, int1 - 1) = -1 ' dr1(int1 - 1).Item("id_tblConfigHeaderLookup")
                            arrOrder(4, int1 - 1) = "Week" ' dr1(int1 - 1).Item("charUserLabel")
                            arrOrder(5, int1 - 1) = var2
                            arrOrder(1, int1 - 1) = "Week" ' dr(0).Item("CHARCOLUMNLABEL")
                            arrOrder(6, int1 - 1) = "WEEK" ' dr(0).Item("CHARWATSONFIELD")

                        End If
                    Next

                    ctCols = int1


                    'Now order tblRepeatTableRows
                    strS = "DESIGNSUBJECTTAG ASC"
                    dvRepeatTableRows = tblRepeatTableRows.DefaultView
                    dvRepeatTableRows.Sort = strS

skipPlaceholder:

                    wrdselection = wd.Selection()


                    Try

                        '20180913 LEE:
                        Call IncrNextTableNumber(wd)

                        If boolPlaceHolder Then
                            '.ActiveDocument.Tables.Add(Range:=wrdselection.Range, NumRows:=21, NumColumns:=1, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)
                            .ActiveDocument.Tables.Add(Range:=wrdselection.Range, NumRows:=1, NumColumns:=1, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                        Else
                            '.ActiveDocument.Tables.Add(Range:=wrdselection.Range, NumRows:=2 + ctRepeatTableRows, NumColumns:=ctCols, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)
                            .ActiveDocument.Tables.Add(Range:=wrdselection.Range, NumRows:=2 + ctRepeatTableRows, NumColumns:=ctCols, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                        End If

                        .Selection.Tables.Item(1).Columns.PreferredWidth = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPoints
                        .Selection.Tables.Item(1).Select()

                        Call SetCellPaddingZero(.Selection.Tables.Item(1))

                        Call removeBorderButLeaveTopAndBottom(wd)
                        .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

                        If boolPlaceHolder Then

                            .Selection.Tables.Item(1).Select()
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone

                            'strTName = strTName.Replace("[MATRIX]", strMatrix) 'Need to do this here for now
                            '20160221 LEE: Made function to update. 
                            strTName = UpdateAnalyteMatrix(strTName, strAnalyteDescription, strMatrix, False, 0, False)
                            Call EnterTableNumber(wd, strTName, 3, strAnalyteDescription, strTempInfo, intTableID, 1, idTR)
                            'var1 = dvDo(intDo).Item("CHARHEADINGTEXT") 'Then change it back
                            'strTName = NZ(var1, "[NONE]")
                            Call MoveOneCellDown(wd)

                            .Selection.TypeParagraph()
                            .Selection.TypeParagraph()

                            'enter a table record in tblTableN
                            'ctTableN = ctTableN + 1
                            Dim dtblr1 As DataRow = tblTableN.NewRow
                            dtblr1.BeginEdit()
                            dtblr1.Item("TableNumber") = ctTableN
                            dtblr1.Item("AnalyteName") = strAnalyteDescription
                            dtblr1.Item("TableName") = strTNameO
                            dtblr1.Item("TableID") = intTableID
                            dtblr1.Item("CHARFCID") = charFCID
                            dtblr1.Item("TableNameNew") = strTName
                            tblTableN.Rows.Add(dtblr1)

                            GoTo next1
                        End If

                        .Selection.Tables.Item(1).Select()
                        Call GlobalTableParaFormat(wd)

                        '20171220 LEE: Do not set table size, use the style default table
                        '.Selection.Font.Size = fontsize - 1
                        .Selection.Tables.Item(1).Cell(1, 1).Select()

                        'enter headings
                        For Count4 = 1 To ctCols
                            var1 = arrOrder(1, Count4)
                            var2 = arrOrder(4, Count4)
                            Select Case var1
                                Case "Dup 1 Result"
                                    var2 = var2 & ChrW(10) & "(" & strConcUnits & ")"
                                Case "Dup 2 Result"
                                    var2 = var2 & ChrW(10) & "(" & strConcUnits & ")"
                                Case "Mean of Dups"
                                    var2 = var2 & ChrW(10) & "(" & strConcUnits & ")"
                                Case "Orig Result"
                                    var2 = var2 & ChrW(10) & "(" & strConcUnits & ")"
                                Case "Reported Result"
                                    var2 = var2 & ChrW(10) & "(" & strConcUnits & ")"

                            End Select
                            .Selection.Tables.Item(1).Cell(1, Count4).Select()
                            .Selection.Text = var2 ' arrOrder(4, Count4)
                            'override default spacing
                            .Selection.ParagraphFormat.SpaceAfter = 0
                        Next

                        'border top and bottom of range
                        .Selection.Tables.Item(1).Cell(1, 1).Select()
                        '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=2, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                        .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom 'bottom align headers

                        'begin entering sample info
                        .Selection.Tables.Item(1).Cell(3, 1).Select()

                        '1=DESIGNSAMPLEID, 2=ANALYTEID, 3=DESIGNSUBJECTTAG, 4=TimePoint, 5=RUNID, 6=RUNSAMPLEORDERNUMBER, 7=DECISIONCODE


                        '''''''''''wdd.visible = True

                        '*** Start Duplicate Logic 2 ***

                        Dim strL As String
                        Dim ctItems As Short = 0

                        strL = frmH.lblProgress.Text

                        boolFirstLine = True

                        Dim ctMax As Short = 20
                        Dim ctIncr As Short = 0

                        For Count2 = 0 To ctRepeatTableRows - 1  'For each Repeat (set of Duplicates)

                            boolFirstLineEntry = True

                            ctIncr = ctIncr + 1
                            '20160220 LEE: For large tables, repeating every entry can slow down process because of screen redraw
                            'implement ctMax to screen redraw only every ctmax items
                            If Count2 = 0 Then '
                                str1 = strL & ChrW(10) & "Processing Repeat " & Count2 + 1 & " of " & ctRepeatTableRows & "..."
                                frmH.lblProgress.Text = str1
                                frmH.lblProgress.Refresh()
                            ElseIf ctIncr >= ctMax Then
                                str1 = strL & ChrW(10) & "Processing Repeat " & Count2 + 1 & " of " & ctRepeatTableRows & "..."
                                frmH.lblProgress.Text = str1
                                frmH.lblProgress.Refresh()
                                ctIncr = 0
                            ElseIf Count2 = ctRepeatTableRows - 1 Then
                                str1 = strL & ChrW(10) & "Processing Repeat " & ctRepeatTableRows & " of " & ctRepeatTableRows & "..."
                                frmH.lblProgress.Text = str1
                                frmH.lblProgress.Refresh()
                            End If


                            'filter dvRepeatTableRows
                            dvRepeatTableRows.RowFilter = strAnalyteAndMatrixFilter 'Reset RowFilter
                            'need to sort here since original recordset sort doesn't seem to do the trick consistently
                            dvRepeatTableRows.Sort = strSdvRepeatTableRowsSort

                            var1 = dvRepeatTableRows(Count2).Item("DESIGNSAMPLEID")
                            strAnalyteMatrixandDesignIDFilter = strAnalyteAndMatrixFilter & " AND DESIGNSAMPLEID = " & var1
                            dvRepeatTableRows.RowFilter = strAnalyteMatrixandDesignIDFilter
                            'need to sort here since original recordset sort doesn't seem to do the trick consistently
                            dvRepeatTableRows.Sort = strSdvRepeatTableRowsSort
                            If dvRepeatTableRows.Count = 0 Then
                                'skip
                                GoTo nextCount2
                            End If
                            If dvRepeatTableRows.Count <> 1 Then
                                'There should ALWAYS be only 1 entry for a single analyteID and DesignSampleID (i.e. there should be a single result)
                                MsgBox("Error: In Repeat Samples table, one of the samples (AnalyteID = " & strAnalyteID & _
                                       "DesignSampleID = " & dvRepeatTableRows(Count2).Item("DESIGNSAMPLEID") & ") has more than one " & _
                                       "Result associated with it.")
                            End If

                            dvRepeatAllRunSamples.RowFilter = strAnalyteMatrixandDesignIDFilter
                            'need to sort here since original recordset sort doesn't seem to do the trick consistently
                            dvRepeatAllRunSamples.Sort = strSdvRepeatAllRunSamplesSort

                            'NDL: COMMENT THIS OUT WHEN TESTING (TO GET ALL REPEATS, INCLUDING SINGLES)
                            'If dvRepeatAllRunSamples.Count < 3 Then 'skip
                            '    int1 = .Selection.Tables.Item(1).Rows.Count
                            '    If .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber) = int1 Then
                            '    Else
                            '        .Selection.Rows.Delete()

                            '    End If
                            '    GoTo nextCount2
                            'End If

                            ctItems = ctItems + 1

                            '****NDL 2-Feb-2016
                            'tblDuplicates (one column table):  Make Table of Duplicate concentration values for easy access
                            'tblDuplicates is just a list of the duplicate concentrations, along with
                            'their ULOQs and LLOQs (these may differ between duplicates).
                            tblDuplicates = dvRepeatAllRunSamples.ToTable("tblDuplicates", False, "AR_CONCENTRATION", "ARS_ALIQUOTFACTOR", "ANARUNANALYTERESULTS_RUNID", "RUNSAMPLESEQUENCENUMBER", "NM", "VEC", "CONCENTRATIONSTATUS", "ORIGINALVALUE")
                            tblDuplicates.Columns.Add("Duplicates")

                            ctDuplicates = tblDuplicates.Rows.Count

                            For Count3 = 0 To ctDuplicates - 1

                                Dim intRowID As Integer

                                '1. Find Quantitation Limits (ULOQ, LLOQ) for Duplicate
                                numLLOQ = NZ(tblDuplicates.Rows(Count3).Item("NM"), 1000000)
                                numULOQ = NZ(tblDuplicates.Rows(Count3).Item("VEC"), 0)
                                var2 = tblDuplicates.Rows(Count3).Item("AR_CONCENTRATION") 'Not to be confused with SampleResults Concentration

                                '20180219 LEE:
                                numDF = NZ(tblDuplicates.Rows(Count3).Item("ARS_ALIQUOTFACTOR"), 1)

                                '20160412 LEE:
                                'var2 = NZ(var2, 0) 'commented this line out
                                'Note Ricerca 034324:
                                'Sometimes ANARUNANALYTERESULTS_CONCENTRATION is NULL, so we should not call it 0
                                'If isnull, then Logic should evaluate CONCENTRATIONSTATUS
                                Dim boolDo As Boolean = False
                                If IsDBNull(var2) Then
                                    'evaluate CONCENTRATIONSTATUS
                                    var3 = tblDuplicates.Rows(Count3).Item("CONCENTRATIONSTATUS")



                                    str1 = NZ(var3, "")
                                    str3 = "NA"
                                    'if var2=null and len(str1)=0 then there is a problem
                                    If Len(str1) = 0 Then
                                        var2 = 0 'set to 0 because must do something. Can review this later
                                        boolDo = True
                                    Else
                                        If StrComp(str1, "VEC", CompareMethod.Text) = 0 Then

                                            'If boolBQLSHOWCONC Then
                                            '    If boolLUseSigFigs Then
                                            '        strAQL = AQL() & "(>" & DisplayNum(SigFigOrDec(numULOQ, LSigFig, False), LSigFig, False) & ")"
                                            '    Else
                                            '        strAQL = AQL() & "(>" & Format(RoundToDecimalRAFZ(numULOQ, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                            '    End If
                                            'Else
                                            '    strAQL = AQL()
                                            'End If

                                            '****

                                            If boolBQLSHOWCONC Then
                                                If boolBQLLEGEND Then
                                                    If boolLUseSigFigs Then
                                                        strAQL = str3 & " (" & AQL() & ")"
                                                    Else
                                                        strAQL = str3 & " (" & AQL() & ")"
                                                    End If
                                                Else
                                                    If boolLUseSigFigs Then
                                                        strAQL = str3 & strR1 & AQL() & "(>" & DisplayNum(SigFigOrDec(numULOQ / numDF, LSigFig, False), LSigFig, False) & ")"
                                                    Else
                                                        strAQL = str3 & strR1 & AQL() & "(>" & Format(RoundToDecimalRAFZ(numULOQ / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                    End If
                                                End If
                                                boolNA = True
                                            Else
                                                If boolBQLLEGEND Then
                                                    strAQL = AQL()
                                                Else
                                                    If boolLUseSigFigs Then
                                                        strAQL = AQL() & "(>" & DisplayNum(SigFigOrDec(numULOQ / numDF, LSigFig, False), LSigFig, False) & ")"
                                                    Else
                                                        strAQL = AQL() & "(>" & Format(RoundToDecimalRAFZ(numULOQ / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                    End If
                                                End If

                                            End If

                                            '****

                                            str2 = strAQL

                                        ElseIf StrComp(str1, "NM", CompareMethod.Text) = 0 Then

                                            'If boolBQLSHOWCONC Then
                                            '    If boolLUseSigFigs Then
                                            '        strBQL = BQL() & "(<" & DisplayNum(SigFigOrDec(numLLOQ, LSigFig, False), LSigFig, False) & ")"
                                            '    Else
                                            '        strBQL = BQL() & "(<" & Format(RoundToDecimalRAFZ(numLLOQ, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                            '    End If
                                            'Else
                                            '    strBQL = BQL()
                                            'End If

                                            '****

                                            If boolBQLSHOWCONC Then
                                                If boolBQLLEGEND Then
                                                    If boolLUseSigFigs Then
                                                        strBQL = str3 & " (" & BQL() & ")"
                                                    Else
                                                        strBQL = str3 & " (" & BQL() & ")"
                                                    End If
                                                Else
                                                    If boolLUseSigFigs Then
                                                        strBQL = str3 & strR1 & BQL() & "(<" & DisplayNum(SigFigOrDec(numLLOQ / numDF, LSigFig, False), LSigFig, False) & ")"
                                                    Else
                                                        strBQL = str3 & strR1 & BQL() & "(<" & Format(RoundToDecimalRAFZ(numLLOQ / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                    End If
                                                End If

                                            Else
                                                If boolBQLLEGEND Then
                                                    strBQL = BQL()
                                                Else
                                                    If boolLUseSigFigs Then
                                                        strBQL = BQL() & "(<" & DisplayNum(SigFigOrDec(numLLOQ / numDF, LSigFig, False), LSigFig, False) & ")"
                                                    Else
                                                        strBQL = BQL() & "(<" & Format(RoundToDecimalRAFZ(numLLOQ / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                    End If
                                                End If

                                            End If

                                            '****

                                            str2 = strBQL

                                        Else

                                            'assume it's BQL

                                            'If boolBQLSHOWCONC Then
                                            '    If boolLUseSigFigs Then
                                            '        strBQL = BQL() & "(<" & DisplayNum(SigFigOrDec(numLLOQ, LSigFig, False), LSigFig, False) & ")"
                                            '    Else
                                            '        strBQL = BQL() & "(<" & Format(RoundToDecimalRAFZ(numLLOQ, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                            '    End If
                                            'Else
                                            '    strBQL = BQL()
                                            'End If

                                            '***

                                            If boolBQLSHOWCONC Then
                                                If boolBQLLEGEND Then
                                                    If boolLUseSigFigs Then
                                                        strBQL = str3 & " (" & BQL() & ")"
                                                    Else
                                                        strBQL = str3 & " (" & BQL() & ")"
                                                    End If
                                                Else
                                                    If boolLUseSigFigs Then
                                                        strBQL = str3 & strR1 & BQL() & "(<" & DisplayNum(SigFigOrDec(numLLOQ / numDF, LSigFig, False), LSigFig, False) & ")"
                                                    Else
                                                        strBQL = str3 & strR1 & BQL() & "(<" & Format(RoundToDecimalRAFZ(numLLOQ / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                    End If
                                                End If

                                            Else
                                                If boolBQLLEGEND Then
                                                    strBQL = BQL()
                                                Else
                                                    If boolLUseSigFigs Then
                                                        strBQL = BQL() & "(<" & DisplayNum(SigFigOrDec(numLLOQ / numDF, LSigFig, False), LSigFig, False) & ")"
                                                    Else
                                                        strBQL = BQL() & "(<" & Format(RoundToDecimalRAFZ(numLLOQ / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                    End If
                                                End If

                                            End If

                                            '***

                                            str2 = strBQL
                                        End If

                                        tblDuplicates.Rows(Count3).Item("Duplicates") = str2

                                    End If
                                Else
                                    'do previous logic
                                    boolDo = True
                                End If

                                'Do a Manual check of LLOQ and ULOQ.  If I knew that Watson always had the correct Concentration Status
                                '(for all supported versions), then I would just rely on that.  But I can't, so I do it manually,
                                'and check against it.
                                If boolDo Then

                                    Try
                                        If boolLUseSigFigs Then
                                            str1 = DisplayNum(SigFigOrDec(var2, LSigFig, False), LSigFig, False)
                                        Else
                                            str1 = Format(RoundToDecimalRAFZ(var2, LSigFig), GetRegrDecStr(LSigFig))
                                        End If
                                    Catch ex As Exception
                                        str1 = "NA"
                                    End Try

                                    If (var2 < numLLOQ) Then

                                        'If boolBQLSHOWCONC Then
                                        '    If boolLUseSigFigs Then
                                        '        strBQL = BQL() & "(<" & DisplayNum(SigFigOrDec(numLLOQ, LSigFig, False), LSigFig, False) & ")"
                                        '    Else
                                        '        strBQL = BQL() & "(<" & Format(RoundToDecimalRAFZ(numLLOQ, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                        '    End If
                                        'Else
                                        '    strBQL = BQL()
                                        'End If

                                        '*****

                                        If boolBQLSHOWCONC Then
                                            If boolBQLLEGEND Then
                                                If boolLUseSigFigs Then
                                                    strBQL = str1 & " (" & BQL() & ")"
                                                Else
                                                    strBQL = str1 & " (" & BQL() & ")"
                                                End If
                                            Else
                                                If boolLUseSigFigs Then
                                                    strBQL = str1 & strR1 & BQL() & "(<" & DisplayNum(SigFigOrDec(numLLOQ / numDF, LSigFig, False), LSigFig, False) & ")"
                                                Else
                                                    strBQL = str1 & strR1 & BQL() & "(<" & Format(RoundToDecimalRAFZ(numLLOQ / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                End If
                                            End If

                                        Else

                                            If boolBQLLEGEND Then
                                                strBQL = BQL()
                                            Else
                                                If boolLUseSigFigs Then
                                                    strBQL = BQL() & "(<" & DisplayNum(SigFigOrDec(numLLOQ / numDF, LSigFig, False), LSigFig, False) & ")"
                                                Else
                                                    strBQL = BQL() & "(<" & Format(RoundToDecimalRAFZ(numLLOQ / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                End If
                                            End If

                                        End If

                                        '*****

                                        tblDuplicates.Rows(Count3).Item("Duplicates") = strBQL

                                    ElseIf (var2 > numULOQ) Then

                                        'If boolBQLSHOWCONC Then
                                        '    If boolLUseSigFigs Then
                                        '        strAQL = AQL() & "(>" & DisplayNum(SigFigOrDec(numULOQ, LSigFig, False), LSigFig, False) & ")"
                                        '    Else
                                        '        strAQL = AQL() & "(>" & Format(RoundToDecimalRAFZ(numULOQ, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                        '    End If
                                        'Else
                                        '    strAQL = AQL()
                                        'End If

                                        '****

                                        If boolBQLSHOWCONC Then
                                            If boolBQLLEGEND Then
                                                If boolLUseSigFigs Then
                                                    strAQL = str1 & " (" & AQL() & ")"
                                                Else
                                                    strAQL = str1 & " (" & AQL() & ")"
                                                End If
                                            Else
                                                If boolLUseSigFigs Then
                                                    strAQL = str1 & strR1 & AQL() & "(>" & DisplayNum(SigFigOrDec(numULOQ / numDF, LSigFig, False), LSigFig, False) & ")"
                                                Else
                                                    strAQL = str1 & strR1 & AQL() & "(>" & Format(RoundToDecimalRAFZ(numULOQ / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                End If
                                            End If

                                        Else
                                            If boolBQLLEGEND Then
                                                strAQL = AQL()
                                            Else
                                                If boolLUseSigFigs Then
                                                    strAQL = AQL() & "(>" & DisplayNum(SigFigOrDec(numULOQ / numDF, LSigFig, False), LSigFig, False) & ")"
                                                Else
                                                    strAQL = AQL() & "(>" & Format(RoundToDecimalRAFZ(numULOQ / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                End If
                                            End If

                                        End If

                                        '****

                                        tblDuplicates.Rows(Count3).Item("Duplicates") = strAQL

                                    Else

                                        If ((StrComp(NZ(tblDuplicates.Rows(Count3).Item("CONCENTRATIONSTATUS"), ""), "VEC") = 0) Or _
                                            (StrComp(NZ(tblDuplicates.Rows(Count3).Item("CONCENTRATIONSTATUS"), ""), "NM") = 0)) Then
                                            'Something's Wrong - the Watson concentration status is entered,
                                            'but it doesn't agree with our check  (this should never happen)
                                            MsgBox("Error: In Repeat Samples table, one of the samples (Run= " & intRowID & " AND AnalyteID = " & strAnalyteID & _
                                                   "Run Sequence # = " & tblDuplicates.Rows(Count3).Item("RUNSAMPLESEQUENCENUMBER") & ") is marked out of range in Watson; but it " & _
                                                   "is between LLOQ (" & numLLOQ & ") and ALOQ (" & numULOQ & "). Diluted Concentration = " & var2 & ".  It has not been" &
                                                   "marked out of range in the Repeat Samples table.")
                                        End If

                                        'Calculate Concentration Value (be sure to use the Aliquot Factor for the repeat pool,
                                        'as this could in theory be different for each Duplicate
                                        var3 = tblDuplicates.Rows(Count3).Item("ARS_ALIQUOTFACTOR") 'Take the Aliquot Factor Analytical Run.
                                        tblDuplicates.Rows(Count3).Item("Duplicates") = strCalcConcWithDilution(var2, var3)
                                    End If
                                End If

                            Next
                            dvDuplicates = New DataView(tblDuplicates)

                            Dim strStartDay As String = ""
                            Dim strStartHour As String = ""
                            Dim strStartMinute As String = ""

                            For Count3 = 1 To ctCols

                                strFld = arrOrder(6, Count3)
                                str1 = arrOrder(1, Count3)
                                Select Case str1
                                    Case "Subject"
                                        var1 = dvRepeatTableRows(0).Item("DESIGNSUBJECTTAG")
                                        If StrComp(var1, "303", CompareMethod.Text) = 0 Then
                                            var1 = var1
                                        End If
                                    Case "Custom ID"
                                        var1 = NZ(dvRepeatTableRows(0).Row("USERSAMPLEID"), "NA")
                                    Case "Treatment"
                                        var1 = dvRepeatTableRows(0).Item("TREATMENTID")
                                    Case "%Diff between Dups"
                                        If (ctDuplicates = 3) Then
                                            dup1 = tblDuplicates.Rows(1).Item("Duplicates")
                                            dup2 = tblDuplicates.Rows(2).Item("Duplicates")
                                            If IsNumeric(dup1) And IsNumeric(dup2) Then 'Can be calcuulated
                                                var1 = Format((CDec(dup1) - CDec(dup2)) / CDec(dup1) * 100, strQCDec) & "%"
                                            Else
                                                var1 = "NA"
                                                boolNA = True
                                            End If
                                        Else
                                            var1 = "NA"
                                            boolNA = True
                                        End If
                                    Case "%Diff between Orig and Mean Dup"
                                        Dim boolAllNumeric As Boolean = True
                                        Dim sumDuplicates As Decimal = 0
                                        Dim ctNumericDuplicates As Short = 0

                                        'Check if numeric, and calculates sum and count
                                        For Count4 = 0 To tblDuplicates.Rows.Count - 1 'Check duplicates AND Orig
                                            If (Not (IsNumeric(tblDuplicates.Rows(Count4).Item("Duplicates")))) Then
                                                boolAllNumeric = False
                                            ElseIf (Count4 = 0) Then 'set Original
                                                numOrig = CDec(tblDuplicates.Rows(Count4).Item("Duplicates"))
                                            Else 'Find sum/count of duplicates
                                                sumDuplicates = sumDuplicates + CDec(tblDuplicates.Rows(Count4).Item("Duplicates"))
                                                ctNumericDuplicates = ctNumericDuplicates + 1
                                            End If
                                        Next

                                        'Perform Calculation
                                        If boolAllNumeric Then
                                            If boolLUseSigFigs Then
                                                var2 = CDec(SigFigOrDec(sumDuplicates / ctNumericDuplicates, LSigFig, False))
                                            Else
                                                var2 = CDec(RoundToDecimalRAFZ(sumDuplicates / ctNumericDuplicates, LSigFig))
                                            End If
                                            var1 = Format((var2 - numOrig) / numOrig * 100, strQCDec) & "%"
                                        Else
                                            var1 = "NA"
                                            boolNA = True
                                        End If
                                    Case "Week"
                                        var1 = NZ(dvRepeatTableRows(0).Row("WEEK"), "")
                                        var1 = var1 'debug
                                    Case "Day"
                                        var3 = dvRepeatTableRows.Count 'DEBUG
                                        var1 = NZ(dvRepeatTableRows(0).Row("ENDDAY"), 1)
                                        var1 = var1 'debug
                                        'check for startday
                                        strStartDay = NZ(dvRepeatTableRows(0).Row("STARTDAY"), "")

                                        If Len(strStartDay) = 0 Then
                                        Else
                                            var1 = strStartDay & " to " & var1
                                        End If

                                    Case "Dup 1 Result"
                                        dvDuplicates.RowFilter = "ORIGINALVALUE = 'N'"

                                        If (dvDuplicates.Count = 0) Then
                                            var1 = "NA"
                                            boolNA = True
                                        Else
                                            var1 = CStr(dvDuplicates(0).Item("Duplicates"))
                                        End If
                                    Case "Dup 2 Result"
                                        dvDuplicates.RowFilter = "ORIGINALVALUE = 'N'"
                                        If (dvDuplicates.Count < 2) Then
                                            var1 = "NA"
                                            boolNA = True
                                        Else  'If more than 2 duplicates, just put them all here.
                                            var1 = ""
                                            For Count4 = 1 To dvDuplicates.Count - 1 'Start at 2nd duplicate
                                                var1 = var1 & CStr(dvDuplicates(Count4).Item("Duplicates"))
                                                If Count4 <> dvDuplicates.Count - 1 Then
                                                    var1 = var1 & strR
                                                End If
                                            Next
                                            var1 = var1 'debug
                                        End If

                                    Case "Dup Run(s)"  'LARRY - USE THIS WHEN IMPLEMENTING...
                                        dvDuplicates.RowFilter = "ORIGINALVALUE = 'N'"

                                        'create a unique dataset
                                        Dim tblURunID As DataTable = dvDuplicates.ToTable("a", True, "ANARUNANALYTERESULTS_RUNID")
                                        var1 = ""
                                        If tblURunID.Rows.Count = 0 Then
                                            var1 = "NA" 'should never happen
                                            boolNA = True
                                        Else
                                            For Count5 = 0 To tblURunID.Rows.Count - 1
                                                var1 = var1 & CStr(tblURunID.Rows(Count5).Item("ANARUNANALYTERESULTS_RUNID"))
                                                If Count5 <> tblURunID.Rows.Count - 1 Then
                                                    var1 = var1 & strR 'replace strR later with soft return
                                                End If
                                            Next
                                        End If
                                        var1 = var1 'debug

                                        'If (dvDuplicates.Count = 0) Then
                                        '    var1 = "NA" 'This should never happen
                                        '    boolNA = True
                                        'Else
                                        '    var1 = ""
                                        '    For Count4 = 0 To dvDuplicates.Count - 1  'Start at 1st duplicate
                                        '        Dim boolThereAlready = False
                                        '        'Only add if not yet there.
                                        '        For Count5 = 0 To Count4 - 1
                                        '            If (dvDuplicates(Count4).Item("ANARUNANALYTERESULTS.RUNID") = dvDuplicates(Count5).Item("ANARUNANALYTERESULTS.RUNID")) Then
                                        '                boolThereAlready = True
                                        '            End If
                                        '        Next
                                        '        If Not (boolThereAlready) Then 'Add Run
                                        '            var1 = var1 & CStr(dvDuplicates(Count4).Item("ANARUNANALYTERESULTS.RUNID"))
                                        '            If Count4 <> dvDuplicates.Count - 1 Then
                                        '                'var1 = var1 & ","
                                        '                var1 = var1 & strR 'replace later with soft return
                                        '            End If
                                        '        End If
                                        '    Next
                                        'End If

                                    Case "Mean of Dups"
                                        Dim boolAllNumeric As Boolean = True
                                        Dim sumDuplicates As Decimal = 0
                                        Dim ctNumericDuplicates As Short = 0
                                        For Count4 = 1 To tblDuplicates.Rows.Count - 1 'Check duplicates
                                            If (Not (IsNumeric(tblDuplicates.Rows(Count4).Item("Duplicates")))) Then
                                                boolAllNumeric = False
                                            Else
                                                sumDuplicates = sumDuplicates + CDec(tblDuplicates.Rows(Count4).Item("Duplicates"))
                                                ctNumericDuplicates = ctNumericDuplicates + 1
                                            End If
                                        Next

                                        If boolAllNumeric Then
                                            If boolLUseSigFigs Then
                                                numMeanDup = CStr(DisplayNum(SigFigOrDec(sumDuplicates / ctNumericDuplicates, LSigFig, False), LSigFig, False))
                                            Else
                                                numMeanDup = CStr(Format(RoundToDecimalRAFZ(sumDuplicates / ctNumericDuplicates, LSigFig), GetRegrDecStr(LSigFig)))
                                            End If
                                            var1 = numMeanDup
                                        Else
                                            var1 = "NA"
                                            boolNA = True
                                        End If

                                    Case "Orig Result"
                                        dvDuplicates.RowFilter = "ORIGINALVALUE = 'Y'"
                                        If (dvDuplicates.Count <> 1) Then
                                            var1 = "NA" 'This should never happen
                                            boolNA = True
                                        Else
                                            var1 = CStr(dvDuplicates(0).Item("Duplicates"))
                                        End If

                                        'Case "Orig Run"  'LARRY - USE THIS WHEN IMPLEMENTING...
                                    Case "Watson Run ID"  'LARRY - USE THIS WHEN IMPLEMENTING...
                                        dvDuplicates.RowFilter = "ORIGINALVALUE = 'Y'"
                                        If (dvDuplicates.Count <> 1) Then
                                            var1 = "NA" 'This should never happen
                                            boolNA = True
                                        Else
                                            var1 = CStr(dvDuplicates(0).Item("ANARUNANALYTERESULTS_RUNID"))
                                        End If

                                    Case "Reported Result"

                                        'Count2 = 0 To ctRepeatTableRows - 1
                                        If Count2 = ctRepeatTableRows - 1 Then
                                            var1 = var1 'debug
                                        End If

                                        'If it is VEC or NM, it must be because it refers to a run/runsequence that is ALQ or BLQ.
                                        'Otherwise, it would be a mean, median, or some other average, and could not be outside of range,
                                        'since the numbers have to be within range to be part of the average.  In the latter case, there is no run in 
                                        'the SampleResults table.
                                        'If I want to report the LLOQ or ULOQ in the results, I have to find the referenced run and report its
                                        'LLOQ or ULOQ.  If there is no referenced run, I cannot (in theory) report LLOQ or ULOQ, as different runs
                                        'might have different ranges.  In this case, I just report AQL() or BQL().

                                        '20160502 LEE:
                                        'when evaluating NM/VEC, must look at CALIBRATIONRANGEFLAG, which comes from SAMPLERESULTS
                                        'CONCENTRATIONSTATUS comes from ANARUNANALYTERESULTS, which may not have the appropriate value for Reported Value

                                        var2 = dvRepeatTableRows(0).Item("SAMPLERESULTS_CONCENTRATION")

                                        '20180219 LEE:
                                        'numDF = NZ(dvRepeatTableRows(0).Item("ARS_ALIQUOTFACTOR"), 1)
                                        numDF = NZ(dvRepeatTableRows(0).Item("SAMPLERESULTS_ALIQUOTFACTOR"), 1)

                                        'var3 = dvRepeatTableRows.Count 'debug
                                        'var4 = tblRepeatTableRows.Rows.Count 'debug

                                        '20160314 LEE: concentration may be null
                                        'should not report 0, should report NR
                                        var2 = NZ(var2, "NR")
                                        If StrComp(var2, "NR", CompareMethod.Text) = 0 Then
                                            boolNR = True
                                        End If

                                        '1. Find Quantitation Limits for Result Value
                                        '1a. First, find row that Sample Results refers to
                                        Dim intRowID As Integer
                                        'intRowID = NZ(tblRepeatTableRows.Rows(0).Item("SampleResults_RunID"), -1) 'Not to be confused with ANARUNANALYTERESULTS.RUNID
                                        intRowID = NZ(dvRepeatTableRows(0).Item("SampleResults_RunID"), -1)
                                        Dim boolInRange As Boolean = False

                                        If (intRowID <> -1) Then  'Note that some results don't have Runs associated with them (e.g. medians).
                                            '1b. Then, find ULOQ and LLOQ by filtering the rungroups 
                                            dvCalStdGroupAssayIDsAcc.RowFilter = "RunID = " & intRowID & " AND AnalyteID = " & strAnalyteID
                                            If (dvCalStdGroupAssayIDsAcc.Count <> 1) Then
                                                MsgBox("Error: SRSummaryRepeatSamples - More than one row associated with RunID = " & intRowID & " AND AnalyteID = " & strAnalyteID, vbInformation, "Problem in Repeat Samples table creation...")
                                            End If
                                            numULOQ = dvCalStdGroupAssayIDsAcc(0).Item("ULOQ")
                                            numLLOQ = dvCalStdGroupAssayIDsAcc(0).Item("LLOQ")

                                            'check AQL/BQL manually
                                            If IsNumeric(var2) Then

                                                Try
                                                    If boolLUseSigFigs Then
                                                        str1 = DisplayNum(SigFigOrDec(var2, LSigFig, False), LSigFig, False)
                                                    Else
                                                        str1 = Format(RoundToDecimalRAFZ(var2, LSigFig), GetRegrDecStr(LSigFig))
                                                    End If
                                                Catch ex As Exception
                                                    str1 = "NA"
                                                End Try

                                                If (var2 < numLLOQ) Then

                                                    'If boolBQLSHOWCONC Then
                                                    '    If boolLUseSigFigs Then
                                                    '        strBQL = BQL() & "<(" & DisplayNum(SigFigOrDec(numLLOQ, LSigFig, False), LSigFig, False) & ")"
                                                    '    Else
                                                    '        strBQL = BQL() & "<(" & Format(RoundToDecimalRAFZ(numLLOQ, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                    '    End If
                                                    'Else
                                                    '    strBQL = BQL()
                                                    'End If

                                                    '*****

                                                    If boolBQLSHOWCONC Then
                                                        If boolBQLLEGEND Then
                                                            If boolLUseSigFigs Then
                                                                strBQL = str1 & " (" & BQL() & ")"
                                                            Else
                                                                strBQL = str1 & " (" & BQL() & ")"
                                                            End If
                                                        Else
                                                            If boolLUseSigFigs Then
                                                                strBQL = str1 & strR1 & BQL() & "(<" & DisplayNum(SigFigOrDec(numLLOQ / numDF, LSigFig, False), LSigFig, False) & ")"
                                                            Else
                                                                strBQL = str1 & strR1 & BQL() & "(<" & Format(RoundToDecimalRAFZ(numLLOQ / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                            End If
                                                        End If

                                                    Else

                                                        If boolBQLLEGEND Then
                                                            strBQL = BQL()
                                                        Else
                                                            If boolLUseSigFigs Then
                                                                strBQL = BQL() & "(<" & DisplayNum(SigFigOrDec(numLLOQ / numDF, LSigFig, False), LSigFig, False) & ")"
                                                            Else
                                                                strBQL = BQL() & "(<" & Format(RoundToDecimalRAFZ(numLLOQ / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                            End If
                                                        End If

                                                    End If

                                                    '*****


                                                    numRV = strBQL

                                                ElseIf (var2 > numULOQ) Then

                                                    'If boolBQLSHOWCONC Then
                                                    '    If boolLUseSigFigs Then
                                                    '        strAQL = AQL() & ">(" & DisplayNum(SigFigOrDec(numULOQ, LSigFig, False), LSigFig, False) & ")"
                                                    '    Else
                                                    '        strAQL = AQL() & ">(" & Format(RoundToDecimalRAFZ(numULOQ, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                    '    End If
                                                    'Else
                                                    '    strAQL = AQL()
                                                    'End If


                                                    '*****

                                                    If boolBQLSHOWCONC Then
                                                        If boolBQLLEGEND Then
                                                            If boolLUseSigFigs Then
                                                                strAQL = str1 & " (" & AQL() & ")"
                                                            Else
                                                                strAQL = str1 & " (" & AQL() & ")"
                                                            End If
                                                        Else
                                                            If boolLUseSigFigs Then
                                                                strAQL = str1 & strR1 & AQL() & "(>" & DisplayNum(SigFigOrDec(numULOQ / numDF, LSigFig, False), LSigFig, False) & ")"
                                                            Else
                                                                strAQL = str1 & strR1 & AQL() & "(>" & Format(RoundToDecimalRAFZ(numULOQ / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                            End If
                                                        End If

                                                    Else
                                                        If boolBQLLEGEND Then
                                                            strAQL = AQL()
                                                        Else
                                                            If boolLUseSigFigs Then
                                                                strAQL = AQL() & "(>" & DisplayNum(SigFigOrDec(numULOQ / numDF, LSigFig, False), LSigFig, False) & ")"
                                                            Else
                                                                strAQL = AQL() & "(>" & Format(RoundToDecimalRAFZ(numULOQ / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                            End If
                                                        End If

                                                    End If

                                                    '*****

                                                    numRV = strAQL
                                                Else
                                                    boolInRange = True 'keep as true for now

                                                    'assume bql
                                                    If boolBQLSHOWCONC Then
                                                        If boolBQLLEGEND Then
                                                            If boolLUseSigFigs Then
                                                                strBQL = str1 & " (" & BQL() & ")"
                                                            Else
                                                                strBQL = str1 & " (" & BQL() & ")"
                                                            End If
                                                        Else
                                                            If boolLUseSigFigs Then
                                                                strBQL = str1 & strR1 & BQL() & "(<" & DisplayNum(SigFigOrDec(numLLOQ / numDF, LSigFig, False), LSigFig, False) & ")"
                                                            Else
                                                                strBQL = str1 & strR1 & BQL() & "(<" & Format(RoundToDecimalRAFZ(numLLOQ / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                            End If
                                                        End If

                                                    Else

                                                        If boolBQLLEGEND Then
                                                            strBQL = BQL()
                                                        Else
                                                            If boolLUseSigFigs Then
                                                                strBQL = BQL() & "(<" & DisplayNum(SigFigOrDec(numLLOQ / numDF, LSigFig, False), LSigFig, False) & ")"
                                                            Else
                                                                strBQL = BQL() & "(<" & Format(RoundToDecimalRAFZ(numLLOQ / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                            End If
                                                        End If

                                                    End If

                                                    numRV = strBQL

                                                End If
                                            Else
                                                boolInRange = True
                                            End If
                                        Else
                                            boolInRange = True 'Can't check (not Rows) - make it true for now
                                        End If

                                        'Even if it looks good with manual check - check the CalibrationRangeFlag in the SampleResults table
                                        '(this is equivalent in relation to Sample Results as the ConcentrationStatus is to AnaRunAnalyteResults)
                                        Dim boolHasCRF As Boolean = False
                                        var3 = NZ(dvRepeatTableRows(0).Item("CALIBRATIONRANGEFLAG"), "")
                                        If Len(var3) = 0 Then
                                        Else
                                            boolHasCRF = True
                                        End If

                                        If boolInRange Then
                                            If (StrComp(NZ(dvRepeatTableRows(0).Item("CALIBRATIONRANGEFLAG"), ""), "VEC") = 0) Then
                                                numRV = strAQL ' AQL()
                                                boolInRange = False
                                            ElseIf (StrComp(NZ(dvRepeatTableRows(0).Item("CALIBRATIONRANGEFLAG"), ""), "NM") = 0) Then
                                                numRV = strBQL ' BQL()
                                                boolInRange = False
                                            End If
                                        End If

                                        'If it still looks good, report it.
                                        If boolInRange And IsNumeric(var2) Then
                                            'Concentration Value
                                            'var3 = dvRepeatTableRows(0).Item("SAMPLERESULTS_ALIQUOTFACTOR")

                                            '20180219 LEE: use ARS_ALIQUOTFACTOR
                                            'var3 = dvRepeatTableRows(0).Item("ARS_ALIQUOTFACTOR")
                                            var3 = dvRepeatTableRows(0).Item("SAMPLERESULTS_ALIQUOTFACTOR")
                                            numRV = strCalcConcWithDilution(var2, var3)
                                            var1 = numRV
                                        Else
                                            '20160502 LEE:
                                            'Ricerca Sample Analysis, Cmpd 1, Urine
                                            'Design Sample 605 has var2=NULL, but has NM in CALIBRATIONRANGEFLAG
                                            'New logic: if var2=NULL, but CALIBRATIONRANGEFLAG <> NULL, then report CALIBRATIONRANGEFLAG
                                            If IsNumeric(var2) Or boolHasCRF Then
                                                var1 = numRV
                                            Else
                                                var1 = var2
                                            End If
                                        End If

                                    Case "Time"

                                        Dim varE
                                        Dim vT, vH, vM
                                        Dim vTS, vHS, vMS

                                        'these are endhour and endminutes
                                        var1 = NZ(dvRepeatTableRows(0).Item("ENDHOUR"), 0)
                                        vH = RoundToDecimal(var1, 3)
                                        var2 = NZ(dvRepeatTableRows(0).Item("ENDMINUTE"), 0)
                                        vM = RoundToDecimal(var2 / 60, 3)

                                        vT = vH + vM
                                        str1 = vT & "h"

                                        varE = str1
                                        varE = varE 'debug

                                        'look for StartHour and StartMinute
                                        strStartHour = CStr(NZ(dvRepeatTableRows(0).Item("STARTHOUR"), "")) '
                                        strStartMinute = CStr(NZ(dvRepeatTableRows(0).Item("STARTMINUTE"), "")) '

                                        If Len(strStartHour) <> 0 Or Len(strStartMinute) <> 0 Then

                                            var1 = NZ(dvRepeatTableRows(0).Item("STARTHOUR"), 0)
                                            vHS = RoundToDecimal(var1, 3)
                                            var2 = NZ(dvRepeatTableRows(0).Item("STARTMINUTE"), 0)
                                            vMS = RoundToDecimal(var2 / 60, 3)

                                            vTS = vHS + vMS
                                            str1 = vTS & "h"

                                            varE = str1 & " to " & varE

                                        End If

                                        var1 = varE

                                    Case "Reported"
                                        var1 = dvRepeatTableRows(0).Row("REASSAYCONCREASON")
                                End Select

                                If InStr(1, NZ(var1, ""), BQL, CompareMethod.Text) > 0 Then
                                    boolBQL = True
                                End If
                                If InStr(1, NZ(var1, ""), AQL, CompareMethod.Text) > 0 Then
                                    boolAQL = True
                                End If

                                If boolFirstLineEntry Then
                                    strPasteT = var1
                                    boolFirstLineEntry = False
                                Else
                                    strPasteT = strPasteT & ChrW(9) & var1
                                End If
                            Next Count3



                            If boolFirstLine Then
                                strPaste = strPasteT
                                boolFirstLine = False
                            Else
                                strPaste = strPaste & ChrW(10) & strPasteT
                            End If

                            ''tttt
                            ''insert a line
                            'If Count2 = ctRepeat Then
                            'Else
                            '    .Selection.InsertRowsBelow(1)
                            'End If

nextCount2:
                        Next Count2

                        '**** End Duplicate Logic 2 ****

                        var1 = ctItems 'debugging

                        'now check for extra rows
                        int1 = .Selection.Tables.Item(1).Rows.Count
                        int2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)

                        ctrsRepeat(1, Count2A) = int1 - 2

                        If ctItems = 0 Then

                            .Selection.Tables.Item(1).Cell(int1, 1).Select()
                            str1 = "No Repeat Samples defined"
                            .Selection.Text = str1
                            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                            Try
                                .Selection.Cells.Merge()
                                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                            Catch ex As Exception

                            End Try

                        Else

                            If intRow = int1 Then 'no problem
                            Else
                                'record table number
                                int3 = wd.ActiveDocument.Tables.Count
                                'select next row
                                .Selection.Tables.Item(1).Cell(intRow + 1, 1).Select()
                            End If

                        End If

                        'ttttt
                        'now paste

                        'wd.Visible = True

                        .Selection.Tables.Item(1).Cell(3, 1).Select()

                        If (IsNothing(strPaste)) Then
                        Else

                            Dim rng1 As Word.Range
                            Dim tblW As Word.Table

                            tblW = .Selection.Tables.Item(1)
                            Try
                                rng1 = wd.ActiveDocument.Range(Start:=tblW.Cell(3, 1).Range.Start, End:=tblW.Cell(tblW.Rows.Count, tblW.Columns.Count).Range.End)
                            Catch ex As Exception
                                var1 = ex.Message
                                var1 = var1
                            End Try


                            'send strpaste to clipboard
                            Try
                                Clipboard.Clear()
                            Catch ex As Exception

                            End Try
                            'give time to set
                            Pause(0.1)
                            Try
                                Clipboard.SetText(strPaste, TextDataFormat.Text)
                                'give time to set
                                Pause(0.1)
                            Catch ex As Exception
                                'MsgBox("SetText: " & ex.Message)
                            End Try

                            'debug
                            ''console.writeline(strPaste)

                            'select appropriate rows
                            rng1.Select()
                            'paste from clipboard
                            Try
                                .Selection.PasteAndFormat(Microsoft.Office.Interop.Word.WdRecoveryType.wdPasteDefault)
                            Catch ex As Exception
                                'MsgBox("Paste: " & ex.Message)
                            End Try

                            'the paste action removes the range object and any table formatting, must reset it
                            Call GlobalTableParaFormat(wd)
                            rng1 = wd.ActiveDocument.Range(Start:=tblW.Cell(3, 1).Range.Start, End:=tblW.Cell(tblW.Rows.Count, tblW.Columns.Count).Range.End)
                            rng1.Select()
                            'the paste action removes paragraph formatting, must format again
                            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalTop

                            '20171220 LEE: Do not set table size, use the style default table
                            '.Selection.Font.Size = fontsize - 1

                            'Enforce minimal 5-point spacing in Table
                            '.Selection.Tables.Item(1).Select()
                            '20160218 LEE: not whole table, just table body, not headers
                            'the table body is currently selected
                            Call EnforceMinimumTableVerticalSpacing(wd, 5)



                            '20160219 LEE
                            'replace '_xyz_' with chrw(11)
                            With rng1.Find
                                .ClearFormatting()
                                .Text = strR
                                .Replacement.ClearFormatting()
                                .Replacement.Text = "," & ChrW(11)
                                .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Forward:=True, Wrap:=Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue)
                            End With

                            'replace '_abc_' with chrw(11)
                            With rng1.Find
                                .ClearFormatting()
                                .Text = strR1
                                .Replacement.ClearFormatting()
                                '.Replacement.Text = "," & ChrW(11)
                                .Replacement.Text = ChrW(11)
                                .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Forward:=True, Wrap:=Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue)
                            End With

                            '*****

                            'end ttttt
                            If IsNothing(strPaste) Then
                            Else
                                'Align reported Reason to Left (If we have lines of data)
                                '.Selection.Columns.Last.Select()
                                .Selection.Tables(1).Columns(ctCols).Select()
                                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                                .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalTop 'wdCellAlignVerticalBottom

                                'but put header back as center and bottom
                                .Selection.Tables.Item(1).Cell(1, ctCols).Select()
                                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                                .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

                            End If
                        End If

                    Catch ex As Exception

                        str1 = "There was a problem preparing table:"
                        str1 = strM1 & ChrW(10) & ChrW(10) & str1
                        str1 = str1 & ChrW(10) & ChrW(10)
                        str1 = str1 & ex.Message
                        MsgBox(str1, vbInformation, "Problem...")

                    End Try


                    'enter legend array
                    ctLegend = 0
                    'If boolBQL Then
                    '    ctLegend = ctLegend + 1
                    '    arrLegend(1, ctLegend) = BQL()
                    '    If StrComp(BQL, "BQL", CompareMethod.Text) = 0 Then
                    '        arrLegend(2, ctLegend) = "Below Quantitation Limit"
                    '    Else
                    '        arrLegend(2, ctLegend) = "Below Limit of Quantitation"
                    '    End If
                    '    arrLegend(3, ctLegend) = False
                    'End If

                    'If boolAQL Then
                    '    ctLegend = ctLegend + 1
                    '    arrLegend(1, ctLegend) = AQL()
                    '    If StrComp(AQL, "AQL", CompareMethod.Text) = 0 Then
                    '        arrLegend(2, ctLegend) = "Above Quantitation Limit"
                    '    Else
                    '        arrLegend(2, ctLegend) = "Above Limit of Quantitation"
                    '    End If
                    '    arrLegend(3, ctLegend) = False
                    'End If

                    If boolNA Then
                        ctLegend = ctLegend + 1
                        arrLegend(1, ctLegend) = "NA"
                        arrLegend(2, ctLegend) = "Not Applicable"
                        arrLegend(3, ctLegend) = False
                    End If

                    If boolNR Then
                        ctLegend = ctLegend + 1
                        arrLegend(1, ctLegend) = "NR"
                        arrLegend(2, ctLegend) = "Not Reportable"
                        arrLegend(3, ctLegend) = False
                    End If

                    ctLegend = ctLegend + 1

                    arrLegend(1, ctLegend) = BQL()
                    If boolBQLLEGEND Then
                        If boolLUseSigFigs Then
                            arrLegend(2, ctLegend) = BQLVerbose() & " (" & DisplayNum(SigFigOrDec(numLLOQ, LSigFig, False), LSigFig, False) & " " & strConcUnits & ")"
                        Else
                            arrLegend(2, ctLegend) = BQLVerbose() & " (" & Format(SigFigOrDec(numLLOQ, LSigFig, False), GetRegrDecStr(LSigFig)) & " " & strConcUnits & ")"
                        End If
                    Else
                        arrLegend(2, ctLegend) = BQLVerbose()
                    End If

                    arrLegend(3, ctLegend) = False
                    arrLegend(4, ctLegend) = False

                    ctLegend = ctLegend + 1

                    arrLegend(1, ctLegend) = AQL()
                    'If boolBQLLEGEND Then

                    'Else
                    '    arrLegend(2, ctLegend) = "Above Quantitation Limit"
                    'End If
                    If boolBQLLEGEND Then
                        If boolLUseSigFigs Then
                            arrLegend(2, ctLegend) = AQLVerbose() & " (" & DisplayNum(SigFigOrDec(numULOQ, LSigFig, False), LSigFig, False) & " " & strConcUnits & ")"
                        Else
                            arrLegend(2, ctLegend) = AQLVerbose() & " (" & Format(SigFigOrDec(numULOQ, LSigFig, False), GetRegrDecStr(LSigFig)) & " " & strConcUnits & ")"
                        End If
                    Else
                        arrLegend(2, ctLegend) = AQLVerbose()
                    End If

                    arrLegend(3, ctLegend) = False
                    arrLegend(4, ctLegend) = False


                    ReDim Preserve arrLegend(4, ctLegend)
                    ''autofit to contents the table

                    'autofit table

                    'wd.Visible = True

                    Call AutoFitTable(wd, True)


                    '20160218 LEE: moved this earlier so that headers can be put back to default
                    ''Enforce minimal 5-point spacing in Table
                    '.Selection.Tables.Item(1).Select()
                    ''20160218 LEE: first record default setting
                    ''need to use this in header
                    'numSpaceAfter = wd.Selection.ParagraphFormat.SpaceAfter
                    'Call EnforceMinimumTableVerticalSpacing(wd, 5)

                    'Don't allow rows to break across pages
                    .Selection.Rows.AllowBreakAcrossPages = False

                    '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                    '.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)

                    'border bottom of this table
                    int1 = .Selection.Tables.Item(1).Rows.Count
                    .Selection.Tables.Item(1).Cell(int1, 1)
                    .Selection.Rows.Select()
                    .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                    'enter table number
                    dv = frmH.dgvDataWatson.DataSource
                    'dv1 = frmH.dgCompanyAnalRef.DataSource
                    var1 = strAnalyteDescription
                    str2 = "Summary of " & var1
                    var1 = frmH.cbxAnticoagulant.Text
                    str2 = str2 & " " & var1
                    int1 = FindRowDV("Species", dv)
                    var1 = dv.Item(int1).Item(1)
                    var1 = Capit(CStr(LowerCase(var1)))
                    str2 = str2 & " " & var1
                    'int1 = FindRowDV("Matrix", dv) '7-Feb-2016 Can't do that anymore (with 2 matrices in study)
                    'var1 = dv.Item(int1).Item(1)
                    var1 = strMatrix
                    var1 = Capit(CStr(LowerCase(var1)))
                    str2 = str2 & " " & var1
                    'str2 = str2 & " Study Samples For " & Sheets("Data").Range("SubmittedTo1").Offset(0, 1).Value
                    str2 = str2 & " Repeat Samples for " & NZ(strSponsor, "NA")

                    dv = frmH.dgvDataCompany.DataSource
                    int1 = FindRowDV("Sponsor Study Number", dv)
                    var1 = NZ(dv.Item(int1).Item(1), "NA")
                    If StrComp(var1, "NA", CompareMethod.Text) = 0 Then
                    Else
                        str2 = str2 & " Study " & var1
                    End If

                    Dim tbl3 As Word.Table
                    tbl3 = .Selection.Tables.Item(1)

                    Dim Oldrng As Microsoft.Office.Interop.Word.Range
                    .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine)
                    Oldrng = wd.Selection.Range

                    tbl3.Select()

                    'strTName = strTName.Replace("[MATRIX]", strMatrix) 'Need to do this here for now
                    '20160221 LEE: Made function to update. 
                    strTName = UpdateAnalyteMatrix(strTName, strAnalyteDescription, strMatrix, False, 0, False)
                    Call EnterTableNumber(wd, strTName, 3, strAnalyteDescription, strTempInfo, intTableID, 1, idTR)
                    var1 = dvDo(intDo).Item("CHARHEADINGTEXT") 'Then change it back
                    strTName = NZ(var1, "[NONE]")

                    'Call EnterTableNumber(wd, str2, 3)
                    'return to old spot

                    Oldrng.Select()

                    'enter a table record in tblTableN
                    'ctTableN = ctTableN + 1
                    Dim dtblr As DataRow = tblTableN.NewRow
                    dtblr.BeginEdit()
                    dtblr.Item("TableNumber") = ctTableN
                    dtblr.Item("AnalyteName") = strAnalyteDescription
                    dtblr.Item("TableName") = strTNameO
                    dtblr.Item("TableID") = intTableID
                    dtblr.Item("CHARFCID") = charFCID
                    dtblr.Item("TableNameNew") = strTName
                    tblTableN.Rows.Add(dtblr)


                    'split table, if needed
                    'ctLegend = ctReasons
                    'str1 = Application.StatusBar
                    str1 = frmH.lblProgress.Text

                    '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                    '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)

                    'autofit table
                    Call AutoFitTable(wd, True)

                    strM = "Finalizing " & strTName & "..."
                    strM1 = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    str1 = strM1

                    frmH.lblProgress.Text = strM1
                    frmH.Refresh()

                    'this is the one to use
                    Call SplitTable(wd, 3, ctLegend, arrLegend, str1, False, ctLegend, False, False, False, intTableID)

                    ''autofit table
                    'Call AutoFitTable(wd, True)

                    tbl3.Select()

                    'ttttt
                    'Sub SplitTable(ByVal wd As Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, 
                    'ByVal arr As Object, ByVal strT As String, ByVal DoLegend As Boolean, ByVal intSplitRows As Short, ByVal boolSmallFont As Boolean)

                    Call MoveOneCellDown(wd)
                    Call InsertLegend(wd, intTableID, idTR, False, 1)

skip7:
next1:
                Next Count2A
            Next Count1A
        End With


    End Sub

End Module
