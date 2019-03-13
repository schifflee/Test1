Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.ComponentModel.PropertyDescriptorCollection
Imports Word = Microsoft.Office.Interop.Word
Imports Microsoft.VisualBasic
Imports System.IO

Module modReassaySamplesBU


    Sub SRSummaryReassaySamplesNewBU(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal idTR As Int64)


        'SAMPLERESULTSCONFLICT	Contains accepted reassay samples, original results by run id and seq number.
        'SAMPRESCONFLICTCHOICES	Contains the reassay samples with mean or median choices for the Reassay. 
        'SAMPRESCONFLICTDEC	Contains the concentration decision choices for the Reassay Selection function. of Watson.

        Dim boolPlaceHolderTemp As Boolean = False
        Dim boolDoPlaceHolderTemp As Boolean = False

        Dim intR As Integer

        Dim numNomConc As Decimal
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim Count1A As Int16
        Dim Count2A As Int16
        Dim Count2 As Short
        Dim Count3 As Short
        Dim Count4 As Short
        Dim Count5 As Short
        Dim Count6 As Short
        Dim varE

        Dim strA As String
        Dim strB As String
        Dim intLeg As Int16
        Dim ctLegend1 As Int16

        Dim numDF As Decimal

        'The following from original Excel code

        ''****Start Re-assay Samples
        'Dim arrReassayRows(100)
        'Dim arrReassay(9, 100)
        'Dim arrReasons(100)
        'Dim arrReasonsC(100)
        '1=ANALYTEID, 2=DESIGNSUBJECTTAG, 3=SUBJECTGROUPNAME, 4=ENDDAY, 5=ENDHOUR, 6=RUNID, 7=DESIGNSAMPLEID
        Dim ctReassay, ctReassayRows, ctReasons, ctRepeatRows, ctLegend, ctRPool, ctRepeatTableRows As Int32
        Dim ctReasonsC As Short
        Dim boolStop As Boolean
        Dim bool As Boolean
        Dim intDo As Short
        Dim dvDo As System.Data.DataView
        Dim strTName As String
        Dim tbl1 As System.Data.DataTable
        Dim strF As String
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim int4 As Short
        Dim ctCols As Short
        Dim arrOrder(6, 100)
        Dim tbl As System.Data.DataTable
        Dim dr() As DataRow
        Dim dr1() As DataRow
        Dim dr2() As DataRow
        Dim dr3() As DataRow
        Dim var1, var2, var3
        Dim lng1 As Int64
        Dim lng2 As Int64
        Dim dv As System.Data.DataView

        Dim strConcUnits As String

        Dim wrdselection As Microsoft.Office.Interop.Word.Selection
        Dim tblD As System.Data.DataTable
        Dim dvD As System.Data.DataView

        Dim rowSC() As DataRow
        Dim strFSC As String
        Dim strFld As String
        Dim posrow1, posrow2 As Short
        Dim boolGo As Boolean
        Dim strTempInfo
        Dim strS As String


        Dim intExp As Short

        Dim fonts
        Dim fontsize

        Dim boolShowConcReason As Boolean = False
        Dim boolShowReasReason As Boolean = False
        Dim int8 As Short

        Dim v1, v2, vU

        Dim tbl1A As DataTable
        Dim tbl2A As DataTable
        Dim strR As String = "_xyz_"
        Dim strR1 As String = "_abc_"

        Dim strSubj As String
        Dim boolHit As Boolean
        Dim intDSId As Int64
        Dim intL As Short

        Dim fld As ADODB.Field
        Dim strConcReason, strAnalyteAndMatrixFilter, strAnalyteMatrixandDesignIDFilter As String
        Dim strM, strM1


        Dim tblRepeatTableRows As New DataTable
        Dim dvRepeatAllRunSamples, dvRepeatTableRows As DataView
        Dim dvCalStdGroupAssayIDsAcc As New DataView(tblCalStdGroupAssayIDsAcc)
        Dim dvDuplicates As DataView
        Dim tblDuplicates As New DataTable
        Dim ctDuplicates As Short
        Dim boolFirstLine, boolFirstLineEntry As Boolean

        Dim strSdvRepeatTableRowsSort As String = "" ' "DESIGNSUBJECTTAG ASC, WEEK ASC, ENDDAY ASC, ENDHOUR ASC, SAMPLERESULTS.RUNID ASC"
        Dim strSdvRepeatAllRunSamplesSort As String = "" '"DESIGNSUBJECTTAG ASC, WEEK ASC, ENDDAY ASC, ENDHOUR ASC, SAMPLERESULTS.RUNID ASC, RUNSAMPLESEQUENCENUMBER ASC"

        Dim dup1, dup2, numMeanDup, numOrig, numRV As Object
        Dim numSpaceAfter As Single
        Dim numSpaceAfterNew As Single

        Dim strStartDay As String = ""
        Dim strStartHour As String = ""
        Dim strStartMinute As String = ""

        '1=DESIGNSAMPLEID, 2=ANALYTEID, 3=DESIGNSUBJECTTAG, 4=TimePoint, 5=RUNID, 6=RUNSAMPLEORDERNUMBER, 7=DECISIONCODE

        Dim numLLOQ As Decimal
        Dim numULOQ As Decimal
        Dim strBQL, strAQL As String
        Dim numBQL As Decimal

        Dim tblAG As DataTable = tblAnalyteGroups
        Dim intGroup As Int16

        Dim strTNameO As String 'original Table Name

        Dim charFCID As String
        strF = "ID_TBLREPORTTABLE = " & idTR
        Dim rowsTR() As DataRow = tblReportTable.Select(strF)
        var1 = rowsTR(0).Item("CHARFCID")
        charFCID = NZ(var1, "NA")


        ' wd.Visible = True

        With wd

            ''''''''wdd.visible = True

            fontsize = wd.ActiveDocument.Styles("Normal").Font.Size ' .Selection.Font.Size
            fonts = fontsize ' .Selection.Font.Size

            'dvDo = frmH.dgvReportTableConfiguration.DataSource
            'strTName = "Summary of Reassayed Samples"
            'intDo = FindRowDVByCol(strTName, dvDo, "Table")

            Dim intTableID As Short
            intTableID = 6

            Dim strWRunId As String = GetWatsonColH(intTableID)

            dvDo = frmH.dgvReportTableConfiguration.DataSource
            intDo = FindRowDVNumByCol(intTableID, dvDo, "id_tblconfigreporttables")

            ''Get table name
            'var1 = dvDo(intDo).Item("Table")
            'strTName = NZ(var1, "[NONE]")

            '***
            intDo = FindRowDVNumByCol(idTR, dvDo, "ID_TBLREPORTTABLE")
            intLeg = 0

            'intLegStart = 96
            'boolPro = False

            'Get table name
            'var1 = dvDo(intDo).Item("Table")
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

            Dim strSortARS As String


            '*** Start Duplicate Logic 1 ****

            '20160313 LEE:
            'moved tblRepeatAllRunSamples to DoPrepare
            'this table is needed in Sample Conc table if Concentration is used

            'Create Collapsed table - One entry per Decision [AnaRunAnalyteResults.Concentration, RunID, RunSequence, and OriginalValue 
            'not included, as they will differ between samples]
            'Note the SampleResults.Concentration is the concentration of the result only and not an individual sample per se 
            '(e.g. it may be the mean of the individual runsamples).  Same with SampleResults.CalibrationRangeFlag (shortened to CalibrationRangeFlag).

            strSdvRepeatTableRowsSort = "DESIGNSUBJECTTAG ASC, WEEK ASC, ENDDAY ASC, ENDHOUR ASC, ENDMINUTE ASC, ENDSECOND ASC, SAMPLERESULTS_RUNID ASC"

            strSortARS = "DESIGNSUBJECTTAG ASC, WEEK ASC, ENDDAY ASC, ENDHOUR ASC, ENDMINUTE ASC, ENDSECOND ASC, ANARUNANALYTERESULTS_RUNID ASC"

            strSdvRepeatAllRunSamplesSort = strSdvRepeatTableRowsSort

            dvRepeatAllRunSamples = New DataView(tblRepeatAllRunSamples, "", strSdvRepeatTableRowsSort, DataViewRowState.CurrentRows)
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



            'Dim strSdvRepeatTableRowsSort As String = "" ' "DESIGNSUBJECTTAG ASC, WEEK ASC, ENDDAY ASC, ENDHOUR ASC, SAMPLERESULTS.RUNID ASC"
            'Dim strSdvRepeatAllRunSamplesSort As String = "" '"DESIGNSUBJECTTAG ASC, WEEK ASC, ENDDAY ASC, ENDHOUR ASC, SAMPLERESULTS.RUNID ASC, RUNSAMPLESEQUENCENUMBER ASC"


            dvRepeatTableRows = New DataView(tblRepeatTableRows)
            'need to sort here since original recordset sort doesn't seem to do the trick consistently
            dvRepeatTableRows.Sort = strSdvRepeatTableRowsSort

            '*** End Duplicate Logic 1 ****

            ''''''''''''wdd.visible = True

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

            Dim strAnalyteID, strAnalyteDescription As String
            Dim strMatrix As String

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

            ''debug
            'var1 = ""
            'For Count1A = 0 To tbl11.Columns.Count - 1
            '    var1 = var1 & ";" & tbl11.Columns(Count1A).ColumnName
            'Next
            'Console.WriteLine(var1)

            'For Count2A = 0 To rows11.Length - 1
            '    var1 = ""
            '    For Count1A = 0 To tbl11.Columns.Count - 1
            '        var1 = var1 & ";" & rows11(Count2A).Item(Count1A)
            '    Next
            '    Console.WriteLine(var1)
            'Next



            ' For Count1A = 0 To tbl1A.Rows.Count - 1 'Iterate through each Matrix (but keep different calibration ranges together)
            For Count1A = 0 To 0 '20180216 LEE: ignore this loop

                'strTName = strTNameO 'reset strTName

                'If boolM Then
                '    strMatrix = tblMatrices.Rows(Count1A).Item("Matrix")
                'Else
                '    strAnalyteID = tblAnalyteIDs.Rows(Count1A).Item("AnalyteID")
                '    strAnalyteDescription = tblAnalyteIDs.Rows(Count1A).Item("AnalyteDescription")
                'End If

                ' For Count2A = 0 To tbl2A.Rows.Count - 1 'Iterate through each AnalyteID, and generate the information
                For Count2A = 0 To intRowsAnal - 1 'Iterate through each AnalyteID, and generate the information

                    '20180201 LEE:
                    intLeg = 0
                    ctLegend1 = 0

                    '20171128 LEE:
                    strTName = strTNameO 'reset strTName

                    boolPlaceHolderTemp = False
                    boolDoPlaceHolderTemp = False

                    Dim boolBQL As Boolean = False
                    Dim boolAQL As Boolean = False
                    Dim boolNA As Boolean = False
                    Dim boolNR As Boolean = False

                    Dim arrLegend1(4, 100)
                    '1= Actual string to search in table
                    '2= Definition of string
                    '3= Not used
                    '4= True: Do not look for item in table, but add buffer row to row count.  False: Look for item in table; if found, add buffer row to row count


                    Dim arrReassayRows(100)
                    Dim arrReassay(9, 100)
                    Dim arrReasons(100)
                    Dim arrReasonsC(100)

                    Dim arrLegend(4, 1000) 'Reason for Reassay
                    Dim arrLegendC(3, 1000) 'Reason for Reported Concentration

                    Dim strPaste As String = ""
                    Dim strPasteT As String = ""

                    Dim strF2 As String
                    Dim strF3 As String


                    'If boolM Then
                    '    strAnalyteID = tblAnalyteIDs.Rows(Count2A).Item("AnalyteID")
                    '    strAnalyteDescription = tblAnalyteIDs.Rows(Count2A).Item("AnalyteDescription")
                    'Else
                    '    strMatrix = tblMatrices.Rows(Count2A).Item("Matrix")
                    'End If

                    ''find intGroup
                    'strF = "ANALYTEID = " & strAnalyteID & " AND MATRIX = '" & strMatrix & "'"
                    'Dim rowsAG() As DataRow = tblAG.Select(strF)
                    'If rowsAG.Length = 0 Then
                    'Else
                    '    intGroup = rowsAG(0).Item("INTGROUP")
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

                    strAnalyteAndMatrixFilter = "SAMPLETYPEID = '" & strMatrix & "' AND ANALYTEID = " & strAnalyteID

                    If (Not (boolGenerateTableForThisAnalyteIDandMatrix(intDo, strAnalyteID, strMatrix))) Then
                        GoTo next1
                    End If

                    ' ''strM = "Entering " & strTName & " for " & arrAnalytes(1, Count2A) & "..."
                    ' ''strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    ' ''strM1 = strM
                    ' ''frmH.lblProgress.Text = strM ' "Entering Summary of Reassayed Samples for " & arrAnalytes(1, Count2A) & " Concentrations Table..."
                    ' ''frmH.lblProgress.Refresh()

                    '' ''check if table is to be generated
                    ' ''strDo = arrAnalytes(1, Count2A) 'record column name

                    ' ''If UseAnalyte(CStr(strDo)) Then
                    ' ''Else
                    ' ''    GoTo end1
                    ' ''End If

                    ' ''bool = dvDo.Item(intDo).Item(strDo) 'find boolean value of dvDo column

                    intTCur = intTCur + 1

                    strM = "Creating " & strTName & " For " & strAnalyteDescription & "..."
                    strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    strM1 = strM
                    frmH.lblProgress.Text = strM
                    frmH.Refresh()

                    If boolPlaceHolder Then

                        'insert page break
                        Call InsertPageBreak(wd)
                        '.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                        Call PageSetup(wd, str1) 'L=Landscape, P=Portrait

                        GoTo skipPlaceholder

                    End If


                    'first get unique items for table
                    'NDL 6-Feb-2016: Filter for Matrix, but don't filter for calibration sets.  Even if a calibration 
                    ' set sub-analyte is not selected, a repeat may include repeat samples with the different LOQ's.  
                    ' So we report all calibration sets regardless of selection.
                    strF = makeRunMatrixFilter(intDo, strAnalyteID, strMatrix) & " AND ORIGINALVALUE = 'Y'"
                    'Console.WriteLine(strF)
                    'strSdvRepeatTableRowsSort = "DESIGNSUBJECTTAG ASC, WEEK ASC, ENDDAY ASC, ENDHOUR ASC, ENDMINUTE ASC, ENDSECOND ASC, SAMPLERESULTS_RUNID ASC"
                    'strS = "DESIGNSUBJECTTAG ASC, RUNID ASC, ENDDAY ASC, ENDHOUR ASC, ENDMINUTE ASC"
                    strS = "DESIGNSUBJECTTAG ASC, WEEK ASC, ENDDAY ASC, ENDHOUR ASC, ENDMINUTE ASC, ENDSECOND, RUNID ASC"
                    Dim dv1 As System.Data.DataView = New DataView(tblReassayReport, strF, strS, DataViewRowState.CurrentRows)
                    'Dim tblUGrps As System.Data.DataTable = dv1.ToTable("a", True, "DESIGNSUBJECTTAG", "ENDDAY", "ENDHOUR", "ENDMINUTE", "RUNID", "DESIGNSAMPLEID")
                    ' Dim tblUGrps As System.Data.DataTable = dv1.ToTable("a", True, "DESIGNSUBJECTTAG", "ENDDAY", "ENDHOUR", "ENDMINUTE", "RUNID", "DESIGNSAMPLEID", "STARTDAY", "STARTHOUR", "STARTMINUTE")
                    Dim tblUGrps As System.Data.DataTable = dv1.ToTable("a", True, "DESIGNSUBJECTTAG", "WEEK", "ENDDAY", "ENDHOUR", "ENDMINUTE", "ENDSECOND", "RUNID", "DESIGNSAMPLEID", "STARTDAY", "STARTHOUR", "STARTMINUTE", "STARTSECOND", "USERSAMPLEID")
                    Dim intUGrps As Short
                    intUGrps = tblUGrps.Rows.Count

                    'determine number of blank rows to skip
                    Dim tblUSubj As System.Data.DataTable = dv1.ToTable("a", True, "DESIGNSUBJECTTAG")
                    'Nope. Need to keep Design Subject separate
                    'Dim tblUSubj As System.Data.DataTable = dv1.ToTable("a", True, "DESIGNSUBJECTTAG", "DESIGNSAMPLEID")
                    Dim intUSubj As Short
                    intUSubj = tblUSubj.Rows.Count

                    If intUSubj = 0 Or intUGrps = 0 Then 'no reassayed samples
                        strM = "The analyte '" & strAnalyteDescription & "' has no reassayed samples."
                        strM = strM & ChrW(10) & ChrW(10)
                        strM = strM & "If you wish to skip this table, click 'Yes'."
                        strM = strM & ChrW(10) & ChrW(10)
                        strM = strM & "If you wish to include a placeholder table for this table, click 'No'."

                        intR = MsgBox(strM, MsgBoxStyle.YesNo, "No samples...")
                        If intR = 6 Then 'Yes
                            GoTo next1
                        Else
                            boolDoPlaceHolderTemp = True
                            boolPlaceHolderTemp = boolPlaceHolder
                            boolPlaceHolder = True
                            'insert page break
                            Call InsertPageBreak(wd)
                            Call PageSetup(wd, str1) 'L=Landscape, P=Portrait
                            GoTo skipPlaceholder
                        End If

                    End If

                    'page setup according to configuration
                    str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")

                    'insert page break
                    Call InsertPageBreak(wd)
                    Call PageSetup(wd, str1) 'L=Landscape, P=Portrait

                    'determine number of columns and order of columns
                    tbl1 = tblReportTableHeaderConfig
                    strF = "id_tblStudies = " & id_tblStudies & " AND id_tblConfigReportTables = " & intTableID & " AND boolInclude = -1"
                    dr1 = tbl1.Select(strF, "intOrder ASC")
                    int1 = dr1.Length
                    ctCols = int1
                    '1=ColumnHeader, 2=Include(X), 3=Order, 4=ReportColumnHeader
                    'arrOrder
                    ' 1=ColumnLabel, 2=Order, 3=id_tblConfigReportTables, 4=UserLabel, 5=id_tblConfigHeaderLookup
                    tbl = tblConfigHeaderLookup
                    arrOrder.Clear(arrOrder, 0, arrOrder.Length)
                    For Count2 = 1 To ctCols
                        arrOrder(2, Count2) = Count2 'dr1(Count2 - 1).Item("intOrder")
                        arrOrder(3, Count2) = dr1(Count2 - 1).Item("id_tblConfigHeaderLookup")
                        arrOrder(4, Count2) = dr1(Count2 - 1).Item("charUserLabel")
                        str2 = dr1(Count2 - 1).Item("charUserLabel") 'just checking
                        arrOrder(5, Count2) = dr1(Count2 - 1).Item("id_tblConfigReportTables")
                        'find column label
                        str1 = "id_tblConfigHeaderLookup = " & arrOrder(3, Count2) & " AND id_tblConfigReportTables = " & intTableID
                        dr = tbl.Select(str1)
                        arrOrder(1, Count2) = dr(0).Item("CHARCOLUMNLABEL")
                        arrOrder(6, Count2) = dr(0).Item("CHARWATSONFIELD")
                    Next

                    'generate reason numbers and record reasons legend
                    int3 = 0
                    'Include even unchecked Analytes, as long as they have the right matrix
                    'because sometimes a sample is re-assayed with a different calibration set
                    '(i.e. different sub-analyte).
                    strF = "ANALYTEID = " & strAnalyteID
                    dr1 = tblReassayReasons.Select(strF) 'this is rsReassay1
                    dvD = New DataView(tblReassayReport)
                    dvD.RowFilter = strF
                    int1 = dr1.Length

                    'first do Reassay Reasons
                    'generate distinct table
                    tblD = dvD.ToTable("a", True, "REASSAYREASON")
                    int2 = tblD.Rows.Count

                    'enter reasons into array
                    'str1 = ""
                    'For Count5 = 0 To tblReassayReport.Columns.Count - 1
                    '    str1 = str1 & tblReassayReport.Columns(Count5).ColumnName & ";"
                    'Next
                    '''''''''''''console.writeline("First try")
                    ''''''''''''''Console.WriteLine(str1)
                    Dim intCt As Short
                    intCt = 0

                    '20160220 LEE: need to filter this for unique subject and week/day/hr/min
                    Dim rowsUG() As DataRow
                    For Count4 = 1 To intUSubj

                        strSubj = tblUSubj.Rows(Count4 - 1).Item("DESIGNSUBJECTTAG")

                        strF = "DESIGNSUBJECTTAG = '" & strSubj & "'"
                        rowsUG = tblUGrps.Select(strF, strS)
                        intL = rowsUG.Length
                        intL = intL 'debug

                        For Count6 = 1 To intL
                            intDSId = rowsUG(Count6 - 1).Item("DESIGNSAMPLEID")

                            '20180226 LEE: Note that Reassay Reason and Reported Reason are correct
                            'sometimes ORIGINALVALUE is switched. Addressed in later code

                            'populate tblreassayreport with Reason Number
                            'str1 = "ANALYTEID = " & strAnalyteID & " AND DESIGNSUBJECTTAG = '" & strSubj & "' AND DESIGNSAMPLEID = " & intDSId
                            '20180124 LEE: Must also make sure there is an original value = Y
                            str1 = "ANALYTEID = " & strAnalyteID & " AND DESIGNSUBJECTTAG = '" & strSubj & "' AND DESIGNSAMPLEID = " & intDSId & " AND ORIGINALVALUE = 'Y'"
                            '20180124 LEE: and MATRIX!!! SAMPLETYPEID
                            str1 = "ANALYTEID = " & strAnalyteID & " AND DESIGNSUBJECTTAG = '" & strSubj & "' AND DESIGNSAMPLEID = " & intDSId & " AND ORIGINALVALUE = 'Y' AND SAMPLETYPEID = '" & strMatrix & "'"
                            'Console.WriteLine(str1)
                            dr2 = tblReassayReport.Select(str1)

                            If dr2.Length = 0 Then
                            Else


                                ''check
                                'var1 = dr2(0).Item("SAMPLETYPEID")
                                'var1 = var1

                                'ensure reason is unique to set
                                var1 = dr2(0).Item("REASSAYREASON")
                                boolHit = False
                                For Count5 = 1 To intCt
                                    var2 = arrLegend(2, Count5)
                                    If StrComp(var1, var2, CompareMethod.Text) = 0 Then
                                        boolHit = True
                                        int1 = Count5
                                        Exit For
                                    End If
                                Next
                                If boolHit Then
                                Else
                                    intCt = intCt + 1
                                    arrLegend(1, intCt) = intCt ' Count2 + 1 'intCt
                                    arrLegend(2, intCt) = dr2(0).Item("REASSAYREASON")
                                    arrLegend(3, intCt) = False
                                    arrReasons(intCt) = dr2(0).Item("REASSAYREASON")
                                    int1 = intCt
                                End If

                                For Count3 = 0 To dr2.Length - 1
                                    dr2(Count3).BeginEdit()
                                    dr2(Count3).Item("numRR") = int1 ' intCt 'Count2 + 1
                                    dr2(Count3).EndEdit()
                                Next
                            End If

                        Next

                    Next

                    ctReasons = intCt 'int2
                    ReDim Preserve arrReasons(ctReasons)

                    'next do Report Concentration Reasons
                    'generate distinct table
                    tblD.Clear()
                    tblD = dvD.ToTable("a", True, "REASSAYCONCREASON")
                    int2 = tblD.Rows.Count
                    If int2 > UBound(arrLegendC, 2) Then
                        ReDim Preserve arrLegendC(3, UBound(arrLegendC, 2) + 100)
                    End If
                    If int2 > UBound(arrReasonsC, 1) Then
                        ReDim Preserve arrReasonsC(UBound(arrReasonsC, 1) + 100)
                    End If

                    'enter reasons into array
                    intCt = 0

                    '20180226 LEE: Note that Reassay Reason and Reported Reason are correct
                    'sometimes ORIGINALVALUE is switched. Addressed in later code

                    For Count4 = 1 To intUSubj

                        strSubj = tblUSubj.Rows(Count4 - 1).Item("DESIGNSUBJECTTAG")

                        strF = "DESIGNSUBJECTTAG = '" & strSubj & "'"
                        rowsUG = tblUGrps.Select(strF, strS)
                        intL = rowsUG.Length
                        intL = intL 'debug

                        For Count6 = 1 To intL

                            intDSId = rowsUG(Count6 - 1).Item("DESIGNSAMPLEID")


                            'populate tblreassayreport with Conc Reason Number
                            'str1 = "ANALYTEID = " & strAnalyteID & " AND DESIGNSUBJECTTAG = '" & strSubj & "' AND DESIGNSAMPLEID = " & intDSId
                            '20180124 LEE: Must also make sure there is an original value = Y
                            str1 = "ANALYTEID = " & strAnalyteID & " AND DESIGNSUBJECTTAG = '" & strSubj & "' AND DESIGNSAMPLEID = " & intDSId & " AND ORIGINALVALUE = 'Y'"
                            '20180124 LEE: and MATRIX!!! SAMPLETYPEID
                            str1 = "ANALYTEID = " & strAnalyteID & " AND DESIGNSUBJECTTAG = '" & strSubj & "' AND DESIGNSAMPLEID = " & intDSId & " AND ORIGINALVALUE = 'Y' AND SAMPLETYPEID = '" & strMatrix & "'"

                            dr2 = tblReassayReport.Select(str1)
                            Dim ar2() As DataRow = tblReassayReport.Select(str1)
                            If dr2.Length = 0 Then
                            Else

                                '****

                                'ensure reason is unique to set
                                var1 = dr2(0).Item("REASSAYCONCREASON")
                                boolHit = False
                                For Count5 = 1 To intCt
                                    var2 = arrLegendC(2, Count5)
                                    If StrComp(var1, var2, CompareMethod.Text) = 0 Then
                                        boolHit = True
                                        int1 = Count5
                                        Exit For
                                    End If
                                Next
                                If boolHit Then
                                Else
                                    intCt = intCt + 1
                                    '1=index, 2=Reason, 3=??
                                    If BOOLREASSAYREASLETTERS Then
                                        str1 = ChrW(64 + intCt)
                                        arrLegendC(1, intCt) = str1 ' Count2 + 1 'intCt
                                    Else
                                        arrLegendC(1, intCt) = intCt ' Count2 + 1 'intCt
                                    End If
                                    'arrLegendC(1, intCt) = intCt ' Count2 + 1 'intCt
                                    arrLegendC(2, intCt) = dr2(0).Item("REASSAYCONCREASON")
                                    arrLegendC(3, intCt) = False
                                    arrReasonsC(intCt) = dr2(0).Item("REASSAYCONCREASON")
                                    int1 = intCt
                                End If

                                For Count3 = 0 To dr2.Length - 1
                                    dr2(Count3).BeginEdit()
                                    If BOOLREASSAYREASLETTERS Then
                                        str1 = ChrW(64 + intCt)
                                        dr2(Count3).Item("numRCR") = str1
                                    Else
                                        dr2(Count3).Item("numRCR") = int1
                                    End If
                                    'dr2(Count3).Item("numRCR") = int1 ' intCt 'Count2 + 1
                                    dr2(Count3).EndEdit()
                                Next

                                '****
                            End If

                        Next Count6


                    Next Count4

                    ctReasonsC = intCt 'int2
                    ReDim Preserve arrReasonsC(ctReasonsC)

                    '***Redo this entire mess

                    ''first get unique items for table
                    'strF = "ANALYTEID = " & arrAnalytes(2, Count2A) & " AND ORIGINALVALUE = 'Y'"
                    'strS = "DESIGNSUBJECTTAG ASC, RUNID ASC, ENDDAY ASC, ENDHOUR ASC"
                    'Dim dv1 as system.data.dataview = New DataView(tblReassayReport, strF, strS, DataViewRowState.CurrentRows)
                    'Dim tblUGrps As System.Data.DataTable = dv1.ToTable("a", True, "DESIGNSUBJECTTAG", "ENDDAY", "ENDHOUR", "RUNID", "DESIGNSAMPLEID")
                    'Dim intUGrps As Short
                    'intUGrps = tblUGrps.Rows.Count

                    ''determine number of blank rows to skip
                    'Dim tblUSubj As System.Data.DataTable = dv1.ToTable("a", True, "DESIGNSUBJECTTAG")
                    'Dim intUSubj As Short
                    'intUSubj = tblUSubj.Rows.Count

                    'determine number of table rows
                    Dim intTRows As Short
                    Dim rowsU() As DataRow

                    intTRows = 0
                    intTRows = intTRows + 1 'for table header
                    intTRows = intTRows + 1 'for blank row
                    intTRows = intTRows + intUGrps 'for unique groups
                    intTRows = intTRows + intUSubj - 1 'for blank row between groups

                    ctrsReassayed(1, Count2A) = intUGrps

skipPlaceholder:

                    wrdselection = wd.Selection()

                    'If intUSubj = 0 Or intUGrps = 0 Then 'no reassayed samples
                    '    'MsgBox("This study has no reassayed samples", MsgBoxStyle.Information, "No data..")
                    '    ''wdd.visible = True
                    '    var1 = var1
                    '    var1 = var1
                    '    GoTo end1
                    'End If


                    Try

                        '20180913 LEE:
                        Call IncrNextTableNumber(wd)

                        If boolPlaceHolder Then
                            .ActiveDocument.Tables.Add(Range:=wrdselection.Range, NumRows:=1, NumColumns:=1, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                        Else
                            .ActiveDocument.Tables.Add(Range:=wrdselection.Range, NumRows:=intTRows, NumColumns:=ctCols, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                        End If

                        .Selection.Tables.Item(1).Columns.PreferredWidth = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPoints


                        '****Start
                        .Selection.Tables.Item(1).Select()

                        Call SetCellPaddingZero(.Selection.Tables.Item(1))

                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                        .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
                        removeBorderButLeaveTopAndBottom(wd)
                        '.Selection.Font.Size = 11

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

                            If boolDoPlaceHolderTemp Then
                                boolPlaceHolder = boolPlaceHolderTemp
                            End If

                            GoTo next1

                        End If

                        .Selection.Tables.Item(1).Select()
                        Call GlobalTableParaFormat(wd)
                        '20171220 LEE: Do not set table size, use the style default table
                        '.Selection.Font.Size = fontsize - 1
                        .Selection.Tables.Item(1).Cell(1, 1).Select()

                        'enter headings
                        For Count4 = 1 To ctCols
                            .Selection.Tables.Item(1).Cell(1, Count4).Select()
                            var1 = arrOrder(1, Count4)
                            var2 = arrOrder(4, Count4)
                            Select Case var1
                                Case "OriginalConc."
                                    var2 = var2 & ChrW(10) & "(" & strConcUnits & ")"
                                Case "ReassayConc."
                                    var2 = var2 & ChrW(10) & "(" & strConcUnits & ")"
                                Case "ReportedConc."
                                    var2 = var2 & ChrW(10) & "(" & strConcUnits & ")"
                            End Select
                            .Selection.Text = var2
                        Next

                        'border top and bottom of range
                        .Selection.Tables.Item(1).Cell(1, 1).Select()
                        '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                        .Selection.Tables.Item(1).Cell(2, 1).Select()


                        int8 = 2

                        '*&*&

                        Dim rows1() As DataRow
                        Dim rows2() As DataRow
                        Dim rows3() As DataRow
                        Dim rows4() As DataRow
                        Dim Count7 As Short
                        Dim rowsReps() As DataRow
                        Dim intRowsReps As Short
                        Dim intRowsO As Short
                        Dim strNo As String
                        Dim intRowsRepsN As Short

                        Dim intCtLbl As Int32 = 0
                        Dim maxCtLbl As Short = 25

                        Dim intN As Int16 = 0

                        Dim ctRows As Short = 0
                        Dim strL As String

                        strL = frmH.lblProgress.Text

                        Try

                            '20171221 LEE: Logic must be redone
                            'Subject can be analyzed twice
                            'Need to look at Design Samples

                            For Count2 = 0 To intUSubj - 1

                                strSubj = tblUSubj.Rows(Count2).Item("DESIGNSUBJECTTAG")
                                '20171221 LEE: Also evaluate DesignSampleID
                                'intDSId = tblUSubj.Rows(Count2).Item("DESIGNSAMPLEID")

                                strF = "DESIGNSUBJECTTAG = '" & strSubj & "'"
                                'strF = "DESIGNSUBJECTTAG = '" & strSubj & "' AND DESIGNSAMPLEID = " & intDSId
                                rowsU = tblUGrps.Select(strF, strS)
                                intL = rowsU.Length
                                intL = intL 'debug
                                For Count4 = 0 To intL - 1

                                    intCtLbl = intCtLbl + 1
                                    If intCtLbl > maxCtLbl Then
                                        str1 = strL & ChrW(10) & "Processing Subject " & Count2 + 1 & " of " & intUSubj & "..."
                                        str1 = str1 & ChrW(10) & "Sample " & Count4 + 1 & " of " & intL & "..."
                                        frmH.lblProgress.Text = str1
                                        frmH.lblProgress.Refresh()
                                        intCtLbl = 0
                                    End If

                                    intDSId = rowsU(Count4).Item("DESIGNSAMPLEID")

                                    ''20180226 LEE:
                                    ''dvDuplicates
                                    ''LI-00016 Study POP01 Furosemide Human Plasma has several original and repeat samples marked incorrectly in SAMPLERESULTSCONFLICT.ORIGINALVALUE (Y and N are switched).
                                    ''This is obvious because the Reassay RunID (e.g. 3) is before the Original RunID (e.g. 7)
                                    ''There must be some problem with the query because actaul Reassay Watson table provided by LI-00016 shows the correct assignment.
                                    ''need to check here to ensure correct assignment
                                    'strF2 = "DESIGNSAMPLEID = " & intDSId & " AND ANALYTEID = " & strAnalyteID & " AND SAMPLETYPEID = '" & strMatrix & "'"
                                    'Dim rowsDups() As DataRow = tblReassayReport.Select(strF2)
                                    ''Dim rowsDups() As DataRow = tblReassayReport.Select(strF2, "RUNID ASC")
                                    ''first entry ORIGINALVALUE should be "Y"
                                    'var1 = rowsDups(0).Item("ORIGINALVALUE")
                                    'If StrComp(var1, "Y", CompareMethod.Text) = 0 Then
                                    'Else
                                    '    'need to reassign ORIGINALVALUE
                                    '    For Count3 = 0 To rowsDups.Length - 1
                                    '        rowsDups(Count3).BeginEdit()
                                    '        If Count3 = 0 Then
                                    '            rowsDups(Count3).Item("ORIGINALVALUE") = "Y"
                                    '        Else
                                    '            rowsDups(Count3).Item("ORIGINALVALUE") = "N"
                                    '        End If
                                    '        rowsDups(Count3).EndEdit()
                                    '    Next
                                    'End If

                                    'strF2 = "DESIGNSAMPLEID = " & intDSId & " AND ORIGINALVALUE = 'Y' AND ANALYTEID = " & strAnalyteID
                                    '20180124 LEE: include matrix
                                    strF2 = "DESIGNSAMPLEID = " & intDSId & " AND ORIGINALVALUE = 'Y' AND ANALYTEID = " & strAnalyteID & " AND SAMPLETYPEID = '" & strMatrix & "'"
                                    Erase rows1
                                    rows1 = tblReassayReport.Select(strF2)

                                    ''20180606 LEE:
                                    ''debug
                                    'var1 = ""
                                    'For Count3 = 0 To tblReassayReport.Columns.Count - 1
                                    '    var1 = var1 & tblReassayReport.Columns(Count3).ColumnName & "\"
                                    'Next
                                    'Console.WriteLine(var1)
                                    'For Count5 = 0 To tblReassayReport.Rows.Count - 1
                                    '    var1 = ""
                                    '    For Count3 = 0 To tblReassayReport.Columns.Count - 1
                                    '        var1 = var1 & tblReassayReport.Rows(Count5).Item(Count3) & "\"
                                    '    Next
                                    '    Console.WriteLine(var1)
                                    'Next

                                    intRowsO = rows1.Length '

                                    strF3 = "DESIGNSAMPLEID = " & intDSId & " AND ORIGINALVALUE = 'N' AND ANALYTEID = " & strAnalyteID
                                    '20180124 LEE: include matrix
                                    strF3 = "DESIGNSAMPLEID = " & intDSId & " AND ORIGINALVALUE = 'N' AND ANALYTEID = " & strAnalyteID & " AND SAMPLETYPEID = '" & strMatrix & "'"
                                    Erase rowsReps
                                    Try
                                        rowsReps = tblReassayReport.Select(strF3)
                                    Catch ex As Exception
                                        var1 = var1
                                    End Try

                                    intRowsRepsN = rowsReps.Length 'not used


                                    'get unique "No" rows
                                    Dim dvNo As System.Data.DataView = New DataView(tblReassayReport, strF3, "", DataViewRowState.CurrentRows)
                                    Dim tblUNo As DataTable = dvNo.ToTable("aN", True, "DESIGNSAMPLEID", "ORIGINALVALUE", "ENDDAY", "ENDHOUR", "ENDMINUTE", "RUNID")
                                    intRowsReps = tblUNo.Rows.Count

                                    'strF = "DESIGNSAMPLEID = " & intDSId & " AND ANALYTEID = " & strAnalyteID
                                    '20180124 LEE: include matrix
                                    strF = "DESIGNSAMPLEID = " & intDSId & " AND ANALYTEID = " & strAnalyteID & " AND SAMPLETYPEID = '" & strMatrix & "'"
                                    Erase rows4
                                    rows4 = tblSampleDesign.Select(strF)
                                    Dim intRRT As Int16
                                    intRRT = rows4.Length '
                                    intRRT = intRRT

                                    '20180226 LEE:
                                    'strNo isn't used anywhere in this code
                                    ''first find number of reassay Nos
                                    'For Count6 = 0 To intRowsReps - 1
                                    '    If Count6 = 0 Then
                                    '        strNo = tblUNo.Rows(Count6).Item("RUNID").ToString
                                    '    Else
                                    '        strNo = strNo & ", " & tblUNo.Rows(Count6).Item("RUNID").ToString
                                    '    End If
                                    'Next


                                    '*** begin use in new

                                    'var1 = dvRepeatTableRows(Count2).Item("DESIGNSAMPLEID")
                                    strAnalyteMatrixandDesignIDFilter = strAnalyteAndMatrixFilter & " AND DESIGNSAMPLEID = " & intDSId
                                    dvRepeatTableRows.RowFilter = strAnalyteMatrixandDesignIDFilter
                                    dvRepeatTableRows.Sort = strSdvRepeatTableRowsSort
                                    'If dvRepeatTableRows.Count = 0 Then
                                    '    'skip
                                    '    GoTo nextCount2
                                    'End If
                                    If dvRepeatTableRows.Count > 1 Or dvRepeatTableRows.Count = 0 Then
                                        'There should ALWAYS be only 1 entry for a single analyteID and DesignSampleID (i.e. there should be a single result)
                                        Try
                                            str1 = "Error: In Repeat Samples table, one of the samples (AnalyteID = " & strAnalyteID & _
                                               "DesignSampleID = " & dvRepeatTableRows(Count2).Item("DESIGNSAMPLEID") & ") has more than one " & _
                                               "Result associated with it."
                                        Catch ex As Exception
                                            str1 = str1
                                        End Try
                                        'MsgBox(str1, vbInformation, "Problem")

                                        '20171221 LEE: Cannot go any further, skip to next count4
                                        GoTo skipNextCount4

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

                                    '****NDL 2-Feb-2016
                                    'tblDuplicates (one column table):  Make Table of Duplicate concentration values for easy access
                                    'tblDuplicates is just a list of the duplicate concentrations, along with
                                    'their ULOQs and LLOQs (these may differ between duplicates).
                                    'tblDuplicates = dvRepeatAllRunSamples.ToTable("tblDuplicates", False, "ANARUNANALYTERESULTS.CONCENTRATION", "ANALYTICALRUNSAMPLE.ALIQUOTFACTOR", "ANARUNANALYTERESULTS.RUNID", "RUNSAMPLESEQUENCENUMBER", "NM", "VEC", "CONCENTRATIONSTATUS", "ORIGINALVALUE")
                                    Dim tblDD As DataTable
                                    tblDD = dvRepeatAllRunSamples.ToTable("tblDD", False, "AR_CONCENTRATION", "ARS_ALIQUOTFACTOR", "ANARUNANALYTERESULTS_RUNID", "RUNSAMPLESEQUENCENUMBER", "NM", "VEC", "CONCENTRATIONSTATUS", "ORIGINALVALUE")
                                    Dim rowsDD() As DataRow = tblDD.Select("", "ANARUNANALYTERESULTS_RUNID ASC")
                                    tblDuplicates = rowsDD.CopyToDataTable
                                    tblDuplicates.Columns.Add("Duplicates")

                                    ctDuplicates = tblDuplicates.Rows.Count

                                    For Count3 = 0 To ctDuplicates - 1

                                        Dim intRowID As Integer

                                        '1. Find Quantitation Limits (ULOQ, LLOQ) for Duplicate
                                        numLLOQ = NZ(tblDuplicates.Rows(Count3).Item("NM"), 1000000)
                                        numULOQ = NZ(tblDuplicates.Rows(Count3).Item("VEC"), 0)
                                        var2 = tblDuplicates.Rows(Count3).Item("AR_CONCENTRATION") 'Not to be confused with SampleResults Concentration

                                        '20180218 LEE:  must adjust numLLOQ/numULOQ for dilution factor
                                        numDF = NZ(tblDuplicates.Rows(Count3).Item("ARS_ALIQUOTFACTOR"), 1)


                                        '20160412 LEE:
                                        'var2 = NZ(var2, 0) 'commented this line out
                                        'Note Ricerca 034324:
                                        'Sometimes ANARUNANALYTERESULTS.CONCENTRATION is NULL, so we should not call it 0
                                        'If isnull, then Logic should evaluate CONCENTRATIONSTATUS
                                        Dim boolDo As Boolean = False
                                        If IsDBNull(var2) Then
                                            'evaluate CONCENTRATIONSTATUS
                                            var3 = tblDuplicates.Rows(Count3).Item("CONCENTRATIONSTATUS")
                                            str1 = NZ(var3, "")
                                            'if var2=null and len(str1)=0 then there is a problem
                                            str3 = "NA"
                                            If Len(str1) = 0 Then
                                                var2 = 0 'set to 0 because must do something. Can review this later
                                                boolDo = True
                                            Else
                                                If StrComp(str1, "VEC", CompareMethod.Text) = 0 Then

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

                                                    '20171221 LEE: No, do not treat as BQL
                                                    'just report value
                                                    str2 = str1

                                                    'If boolBQLSHOWCONC Then
                                                    '    If boolBQLLEGEND Then
                                                    '        If boolLUseSigFigs Then
                                                    '            strBQL = str3 & " (" & BQL() & ")"
                                                    '        Else
                                                    '            strBQL = str3 & " (" & BQL() & ")"
                                                    '        End If
                                                    '    Else
                                                    '        If boolLUseSigFigs Then
                                                    '            strBQL = str3 & strR1 & BQL() & "(<" & DisplayNum(SigFigOrDec(numLLOQ, LSigFig, False), LSigFig, False) & ")"
                                                    '        Else
                                                    '            strBQL = str3 & strR1 & BQL() & "(<" & Format(RoundToDecimalRAFZ(numLLOQ, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                    '        End If
                                                    '    End If

                                                    'Else
                                                    '    If boolBQLLEGEND Then
                                                    '        strBQL = BQL()
                                                    '    Else
                                                    '        If boolLUseSigFigs Then
                                                    '            strBQL = BQL() & "(<" & DisplayNum(SigFigOrDec(numLLOQ, LSigFig, False), LSigFig, False) & ")"
                                                    '        Else
                                                    '            strBQL = BQL() & "(<" & Format(RoundToDecimalRAFZ(numLLOQ, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                    '        End If
                                                    '    End If

                                                    'End If

                                                    ''***

                                                    'str2 = strBQL

                                                    '20180201 LEE: Add to arrLegend
                                                    'Add to Legend Array
                                                    strA = var3
                                                    If StrComp(str1, "Follows >ULOQ sample", CompareMethod.Text) = 0 Then 'skip

                                                    Else
                                                        If StrComp(strA, "N.R.", CompareMethod.Text) = 0 Then
                                                            strB = "Not Reported"

                                                        Else
                                                            strB = "[Manual Entry]"
                                                        End If

                                                        ctLegend1 = ctLegend1 + SetLegendArray(arrLegend1, intLeg, strB, strA, False)
                                                        intLeg = ctLegend1
                                                    End If



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

                                                tblDuplicates.Rows(Count3).Item("Duplicates") = strBQL

                                            ElseIf (var2 > numULOQ) Then

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

                                                tblDuplicates.Rows(Count3).Item("Duplicates") = strAQL

                                            Else
                                                If ((StrComp(NZ(tblDuplicates.Rows(Count3).Item("CONCENTRATIONSTATUS"), ""), "VEC") = 0) Or (StrComp(NZ(tblDuplicates.Rows(Count3).Item("CONCENTRATIONSTATUS"), ""), "NM") = 0)) Then
                                                    'Something's Wrong - the Watson concentration status is entered,
                                                    'but it doesn't agree with our check  (this should never happen)
                                                    'MsgBox("Error: In Repeat Samples table, one of the samples (Run= " & intRowID & " AND AnalyteID = " & strAnalyteID & _
                                                    '       "Run Sequence # = " & tblDuplicates.Rows(Count3).Item("RUNSAMPLESEQUENCENUMBER") & ") is marked out of range in Watson; but it " & _
                                                    '       "is between LLOQ (" & numLLOQ & ") and ALOQ (" & numULOQ & "). Diluted Concentration = " & var2 & ".  It has not been" &
                                                    '       "marked out of range in the Repeat Samples table.")

                                                End If

                                                'Calculate Concentration Value (be sure to use the Aliquot Factor for the repeat pool,
                                                'as this could in theory be different for each Duplicate
                                                var3 = tblDuplicates.Rows(Count3).Item("ARS_ALIQUOTFACTOR") 'Take the Aliquot Factor Analytical Run.
                                                tblDuplicates.Rows(Count3).Item("Duplicates") = strCalcConcWithDilution(var2, var3)

                                            End If
                                        End If

                                    Next Count3
                                    dvDuplicates = New DataView(tblDuplicates)

                                    '*** end use in new


                                    For Count7 = 0 To intRowsO - 1

                                        intN = intN + 1

                                        int8 = int8 + 1
                                        If int8 > .Selection.Tables.Item(1).Rows.Count Then
                                            .Selection.Tables.Item(1).Cell(int8 - 1, 1).Select()
                                            .Selection.InsertRowsBelow(1)

                                        End If

                                        strStartDay = ""
                                        strStartHour = ""
                                        strStartMinute = ""
                                        Dim vT, vH, vM, vS
                                        Dim vTS, vHS, vMS, vSS

                                        Try


                                            For Count3 = 1 To ctCols
                                                varE = ""
                                                Select Case arrOrder(1, Count3)
                                                    Case "Day"
                                                        'str1 = CStr(NZ(rows(Count4).Item("ENDDAY"), "NA")) ' 
                                                        str1 = CStr(NZ(rowsU(Count4).Item("ENDDAY"), 0)) '
                                                        strStartDay = CStr(NZ(rowsU(Count4).Item("STARTDAY"), "")) '

                                                        If Len(strStartDay) = 0 Then
                                                            varE = str1
                                                        Else
                                                            varE = str1 & " to " & strStartDay
                                                        End If

                                                    Case "Custom ID"
                                                        Try
                                                            varE = NZ(rowsU(Count4).Item("USERSAMPLEID"), "NA")
                                                        Catch ex As Exception
                                                            varE = varE 'debug
                                                        End Try


                                                    Case "Orig Watson ID"
                                                        str1 = CStr(rowsU(Count4).Item("RUNID"))
                                                        '20180226 LEE:
                                                        'No, should be rows1
                                                        If rows1.Length = 0 Then
                                                            str1 = "NA"
                                                            boolNA = True
                                                        Else
                                                            str1 = CStr(rows1(0).Item("RUNID"))
                                                        End If

                                                        varE = str1
                                                    Case "Reason for Reassay"
                                                        'intDSId = tblUGrps.Rows(Count4).Item("DESIGNSAMPLEID")
                                                        'strF2 = "DESIGNSAMPLEID = " & intDSId & " AND ORIGINALVALUE = 'Y' AND ANALYTEID = " & arrAnalytes(2, Count2A)
                                                        'Erase rows1
                                                        'rows1 = tblReassayReport.Select(strF2)
                                                        'int1 = rows1.Length 'debugging
                                                        'varE = rows1(0).Item("numRR")
                                                        'varE = rowsReps(Count7).Item("numRR")
                                                        'rows1
                                                        boolShowReasReason = True
                                                        dvDuplicates.RowFilter = "ORIGINALVALUE = 'N'"
                                                        If (dvDuplicates.Count = 0) Then
                                                            var1 = "NA" 'This should never happen
                                                            boolNA = True
                                                        ElseIf dvDuplicates.Count = 1 Then
                                                            var1 = rows1(Count7).Item("numRR")
                                                            'var1 = rowsReps(0).Item("numRR")
                                                        Else
                                                            var1 = rows1(Count7).Item("numRR")
                                                            'var1 = rowsReps(0).Item("numRR")
                                                        End If

                                                        varE = var1
                                                        'varE = varE 'debug

                                                    Case "Reassd Watson ID"

                                                        'varE = rowsReps(Count7).Item("RUNID")

                                                        'dvDuplicates.RowFilter = "ORIGINALVALUE = 'N'"
                                                        'If (dvDuplicates.Count = 0) Then
                                                        '    var1 = "NA" 'This should never happen
                                                        '    boolNA = True
                                                        'Else

                                                        '    dvDuplicates.RowFilter = "ORIGINALVALUE = 'N'"
                                                        '    var1 = ""
                                                        '    For Count5 = 0 To dvDuplicates.Count - 1 '
                                                        '        var1 = var1 & CStr(dvDuplicates(Count5).Item("ANARUNANALYTERESULTS_RUNID"))
                                                        '        If Count5 <> dvDuplicates.Count - 1 Then
                                                        '            var1 = var1 & strR 'replace strR later with soft return
                                                        '        End If
                                                        '    Next
                                                        '    varE = var1 'debug

                                                        'End If

                                                        '20180226 LEE:
                                                        'get From rowsreps
                                                        If rowsReps.Length = 0 Then
                                                            var1 = "NA" 'This should never happen
                                                            boolNA = True
                                                        Else
                                                            For Count5 = 0 To rowsReps.Length - 1
                                                                var2 = rowsReps(Count5).Item("RUNID")
                                                                If Count5 = 0 Then
                                                                    var1 = var2
                                                                Else
                                                                    var1 = var1 & strR & var2
                                                                End If
                                                            Next
                                                        End If

                                                        varE = var1

                                                    Case "Subject"
                                                        varE = strSubj
                                                    Case "Time"
                                                        'these are endhour and endminutes
                                                        var1 = NZ(rowsU(Count4).Item("ENDHOUR"), 0)
                                                        vH = RoundToDecimal(var1, 3)
                                                        var2 = NZ(rowsU(Count4).Item("ENDMINUTE"), 0)
                                                        vM = RoundToDecimal(var2 / 60, 3)

                                                        vT = vH + vM
                                                        str1 = vT & "h"

                                                        varE = str1
                                                        varE = varE 'debug

                                                        'look for StartHour and StartMinute
                                                        strStartHour = CStr(NZ(rowsU(Count4).Item("STARTHOUR"), "")) '
                                                        strStartMinute = CStr(NZ(rowsU(Count4).Item("STARTMINUTE"), "")) '

                                                        If Len(strStartHour) <> 0 Or Len(strStartMinute) <> 0 Then
                                                            var1 = NZ(rowsU(Count4).Item("STARTHOUR"), 0)
                                                            vHS = RoundToDecimal(var1, 3)
                                                            var2 = NZ(rowsU(Count4).Item("STARTMINUTE"), 0)
                                                            vMS = RoundToDecimal(var2 / 60, 3)

                                                            vTS = vHS + vMS
                                                            str1 = vTS & "h"

                                                            varE = str1 & " to " & varE

                                                        End If


                                                    Case "Treatment"
                                                        'TREATMENTDESC
                                                        int1 = rows4.Length
                                                        If int1 = 0 Then
                                                            varE = "NA"
                                                            boolNA = True
                                                        Else
                                                            'varE = rows4(0).Item("TREATMENTDESC")
                                                            varE = rows4(0).Item("TREATMENTDESC")
                                                        End If

                                                    Case "Matrix"
                                                        'Matrix
                                                        int1 = rows4.Length
                                                        If int1 = 0 Then
                                                            varE = "NA"
                                                            boolNA = True
                                                        Else
                                                            'varE = rows4(0).Item("TREATMENTDESC")
                                                            '20180124 LEE: now get from tblReassayReport
                                                            'varE = rows4(0).Item("SAMPLETYPEID")
                                                            varE = rows1(Count7).Item("SAMPLETYPEID")
                                                        End If


                                                    Case "Reason for Reported Conc."
                                                        'var1 = CStr(dr2(Count2 - 1).Item("numRCR"))
                                                        boolShowConcReason = True
                                                        varE = rows1(0).Item("numRCR")
                                                        varE = varE 'debug
                                                    Case "Reported Watson Run ID"
                                                        int1 = rows4.Length
                                                        If int1 = 0 Then
                                                            varE = ""
                                                        Else
                                                            varE = rows4(0).Item("RUNID")
                                                        End If

                                                    Case "OriginalConc."
                                                        dvDuplicates.RowFilter = "ORIGINALVALUE = 'Y'"
                                                        If (dvDuplicates.Count <> 1) Then
                                                            varE = "NA" 'This should never happen
                                                            boolNA = True
                                                        Else
                                                            varE = CStr(dvDuplicates(0).Item("Duplicates"))
                                                        End If
                                                        varE = varE 'debug

                                                    Case "ReassayConc."
                                                        dvDuplicates.RowFilter = "ORIGINALVALUE = 'N'"
                                                        var1 = ""
                                                        For Count5 = 0 To dvDuplicates.Count - 1 '
                                                            var2 = CStr(dvDuplicates(Count5).Item("Duplicates"))
                                                            If Count5 = 0 Then
                                                                var1 = var2
                                                            Else
                                                                var1 = var1 & strR & var2
                                                            End If
                                                            'If Count5 <> dvDuplicates.Count - 1 Then
                                                            '    var1 = var1 & strR 'replace strR later with soft return
                                                            'End If
                                                        Next
                                                        varE = var1 'debug
                                                    Case "ReportedConc."

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

                                                        '20160314 LEE: concentration may be null
                                                        'should not report 0, should report NR
                                                        var2 = NZ(var2, "NR")
                                                        If StrComp(var2, "NR", CompareMethod.Text) = 0 Then
                                                            boolNR = True
                                                        End If

                                                        '1. Find Quantitation Limits for Result Value
                                                        '1a. First, find row that Sample Results refers to
                                                        Dim intRowID As Integer
                                                        'intRowID = NZ(tblRepeatTableRows.Rows(0).Item("SampleResults.RunID"), -1) 'Not to be confused with ANARUNANALYTERESULTS.RUNID
                                                        intRowID = NZ(dvRepeatTableRows(0).Item("SampleResults_RunID"), -1)
                                                        Dim boolInRange As Boolean = False

                                                        If (intRowID <> -1) Then  'Note that some results don't have Runs associated with them (e.g. medians).
                                                            '1b. Then, find ULOQ and LLOQ by filtering the rungroups 
                                                            '20171221 LEE: Aack! Need to account for matrix
                                                            'dvCalStdGroupAssayIDsAcc.RowFilter = "RunID = " & intRowID & " AND AnalyteID = " & strAnalyteID
                                                            dvCalStdGroupAssayIDsAcc.RowFilter = "RunID = " & intRowID & " AND AnalyteID = " & strAnalyteID & " AND MATRIX = '" & strMatrix & "'"

                                                            If (dvCalStdGroupAssayIDsAcc.Count <> 1) Then
                                                                str1 = "Error: SRSummaryRepeatSamples - More than one row associated with RunID = " & intRowID & " AND AnalyteID = " & strAnalyteID & " AND MATRIX = '" & strMatrix & "'"
                                                                MsgBox(str1, vbInformation, "Problem in Repeat Samples table creation...")
                                                            End If
                                                            numULOQ = dvCalStdGroupAssayIDsAcc(0).Item("ULOQ")
                                                            numLLOQ = dvCalStdGroupAssayIDsAcc(0).Item("LLOQ")

                                                            'check AQL/BQL manually
                                                            If IsNumeric(var2) Then

                                                                If boolLUseSigFigs Then
                                                                    str1 = DisplayNum(SigFigOrDec(var2, LSigFig, False), LSigFig, False)
                                                                Else
                                                                    str1 = Format(RoundToDecimalRAFZ(var2, LSigFig), GetRegrDecStr(LSigFig))
                                                                End If

                                                                If (var2 < numLLOQ) Then

                                                                    If boolBQLSHOWCONC Then
                                                                        If boolBQLLEGEND Then
                                                                            If boolLUseSigFigs Then
                                                                                strBQL = str1 & " (" & BQL() & ")"
                                                                            Else
                                                                                strBQL = str1 & " (" & BQL() & ")"
                                                                            End If
                                                                        Else
                                                                            If boolLUseSigFigs Then
                                                                                strBQL = str1 & strR1 & BQL() & "<(" & DisplayNum(SigFigOrDec(numLLOQ / numDF, LSigFig, False), LSigFig, False) & ")"
                                                                            Else
                                                                                strBQL = str1 & strR1 & BQL() & "<(" & Format(RoundToDecimalRAFZ(numLLOQ / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                                            End If
                                                                        End If

                                                                    Else
                                                                        If boolBQLLEGEND Then
                                                                            strBQL = BQL()
                                                                        Else
                                                                            If boolLUseSigFigs Then
                                                                                strBQL = BQL() & "<(" & DisplayNum(SigFigOrDec(numLLOQ / numDF, LSigFig, False), LSigFig, False) & ")"
                                                                            Else
                                                                                strBQL = BQL() & "<(" & Format(RoundToDecimalRAFZ(numLLOQ / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                                            End If
                                                                        End If

                                                                    End If

                                                                    numRV = strBQL

                                                                ElseIf (var2 > numULOQ) Then

                                                                    If boolBQLSHOWCONC Then
                                                                        If boolBQLLEGEND Then
                                                                            If boolLUseSigFigs Then
                                                                                strAQL = str1 & " (" & AQL() & ")"
                                                                            Else
                                                                                strAQL = str1 & " (" & AQL() & ")"
                                                                            End If
                                                                        Else
                                                                            If boolLUseSigFigs Then
                                                                                strAQL = str1 & strR1 & AQL() & ">(" & DisplayNum(SigFigOrDec(numULOQ / numDF, LSigFig, False), LSigFig, False) & ")"
                                                                            Else
                                                                                strAQL = str1 & strR1 & AQL() & ">(" & Format(RoundToDecimalRAFZ(numULOQ / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                                            End If
                                                                        End If

                                                                    Else
                                                                        If boolBQLLEGEND Then
                                                                            strAQL = AQL()
                                                                        Else
                                                                            If boolLUseSigFigs Then
                                                                                strAQL = AQL() & ">(" & DisplayNum(SigFigOrDec(numULOQ / numDF, LSigFig, False), LSigFig, False) & ")"
                                                                            Else
                                                                                strAQL = AQL() & ">(" & Format(RoundToDecimalRAFZ(numULOQ / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                                            End If
                                                                        End If

                                                                    End If

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
                                                                                strBQL = str1 & strR1 & BQL() & "<(" & DisplayNum(SigFigOrDec(numLLOQ / numDF, LSigFig, False), LSigFig, False) & ")"
                                                                            Else
                                                                                strBQL = str1 & strR1 & BQL() & "<(" & Format(RoundToDecimalRAFZ(numLLOQ / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                                            End If
                                                                        End If

                                                                    Else
                                                                        If boolBQLLEGEND Then
                                                                            strBQL = BQL()
                                                                        Else
                                                                            If boolLUseSigFigs Then
                                                                                strBQL = BQL() & "<(" & DisplayNum(SigFigOrDec(numLLOQ / numDF, LSigFig, False), LSigFig, False) & ")"
                                                                            Else
                                                                                strBQL = BQL() & "<(" & Format(RoundToDecimalRAFZ(numLLOQ / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
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
                                                            var3 = dvRepeatTableRows(0).Item("SAMPLERESULTS_ALIQUOTFACTOR")
                                                            numRV = strCalcConcWithDilution(var2, var3)
                                                            varE = numRV
                                                        Else
                                                            '20160502 LEE:
                                                            'Ricerca Sample Analysis, Cmpd 1, Urine
                                                            'Design Sample 605 has var2=NULL, but has NM in CALIBRATIONRANGEFLAG
                                                            'New logic: if var2=NULL, but CALIBRATIONRANGEFLAG <> NULL, then report CALIBRATIONRANGEFLAG
                                                            If IsNumeric(var2) Or boolHasCRF Then
                                                                varE = numRV
                                                            Else
                                                                varE = var2
                                                            End If
                                                        End If
                                                End Select

                                                If Count3 = 1 Then
                                                    strPasteT = varE
                                                Else
                                                    strPasteT = strPasteT & ChrW(9) & varE
                                                End If

                                                'tttt
                                                '.Selection.Tables.Item(1).Cell(int8, Count3).Select()
                                                '.Selection.TypeText(Text:=CStr(NZ(varE, "")))

                                            Next Count3

                                        Catch ex As Exception
                                            var1 = var1
                                        End Try

                                        ''console.writeline(strPasteT)

                                        'If Count2 = 0 And Count4 = 0 And Count7 = 0 Then
                                        '    strPaste = strPasteT
                                        'Else
                                        '    strPaste = strPaste & ChrW(10) & strPasteT
                                        'End If

                                        If Len(strPaste) = 0 Then
                                            strPaste = strPasteT
                                        Else
                                            strPaste = strPaste & ChrW(10) & strPasteT
                                        End If

                                        ctRows = ctRows + 1

                                    Next Count7

skipNextCount4:

                                Next Count4




                                'no, don't add an extra line
                                int8 = int8 + 1
                                If int8 > .Selection.Tables.Item(1).Rows.Count Then
                                    '.Selection.Tables.Item(1).Cell(int8 - 1, 1).Select()
                                    '.Selection.InsertRowsBelow(1)

                                    ''tttt
                                    'varE = ""
                                    'For Count3 = 1 To ctCols
                                    '    If Count3 = 1 Then
                                    '        strPasteT = varE
                                    '    Else
                                    '        strPasteT = strPasteT & ChrW(9) & varE
                                    '    End If
                                    'Next

                                    'If Count2 = 1 Then
                                    '    strPaste = strPasteT
                                    'Else
                                    '    strPaste = strPaste & ChrW(10) & strPasteT
                                    'End If

                                End If

                            Next Count2

                            str1 = strL & ChrW(10) & "Processing Subject " & intUSubj & " of " & intUSubj & "..."
                            str1 = str1 & ChrW(10) & "Sample " & intL & " of " & intL & "..."
                            frmH.lblProgress.Text = str1
                            frmH.lblProgress.Refresh()

                        Catch ex As Exception

                            var1 = 1

                        End Try

                        '*****

                        .Selection.Tables.Item(1).Cell(3, 1).Select()

                        If IsNothing(strPaste) Then
                        Else

                            Dim rng1 As Word.Range
                            Dim tblW As Word.Table

                            tblW = .Selection.Tables.Item(1)
                            'Try
                            '    rng1 = wd.ActiveDocument.Range(Start:=tblW.Cell(3, 1).Range.Start, End:=tblW.Cell(tblW.Rows.Count, tblW.Columns.Count).Range.End)
                            '    rng1 = wd.ActiveDocument.Range(Start:=tblW.Cell(3, 1).Range.Start, End:=tblW.Cell(ctRows + 3 - 1, ctCols).Range.End)
                            'Catch ex As Exception
                            '    var1 = ex.Message
                            '    var1 = var1
                            'End Try

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
                            'select appropriate rows
                            '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                            '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdColumn, Extend:=True)


                            Dim boolV As Boolean
                            boolV = wd.Visible
                            wd.WindowState = Word.WdWindowState.wdWindowStateMinimize
                            wd.Visible = True

                            Try
                                rng1 = wd.ActiveDocument.Range(Start:=tblW.Cell(3, 1).Range.Start, End:=tblW.Cell(ctRows + 3 - 1, ctCols).Range.End)
                                rng1.Select()
                            Catch ex As Exception
                                .Selection.SelectRow()
                                .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=ctRows - 1, Extend:=True)
                                var1 = ex.Message
                                var1 = var1
                            End Try

                            Pause(0.1)

                            '.Selection.SelectRow()
                            '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=ctRows - 1, Extend:=True)
                            ' paste from clipboard
                            ''''console.writeline("Start Paste")

                            Try
                                .Selection.PasteAndFormat(Microsoft.Office.Interop.Word.WdRecoveryType.wdPasteDefault)
                            Catch ex As Exception
                                'MsgBox("Paste: " & ex.Message)
                            End Try
                            ''console.writeline("Skip2")
                            ''console.writeline(strPaste)
                            ''''console.writeline("End Paste")

                            'Pause(0.5)

                            wd.Visible = boolV
                            ''' 
                            'rng1 = wd.ActiveDocument.Range(Start:=tblW.Cell(3, 1).Range.Start, End:=tblW.Cell(tblW.Rows.Count, tblW.Columns.Count).Range.End)
                            'rng1.Select()

                            'the paste action removes the range object and any table formatting, must reset it
                            Call GlobalTableParaFormat(wd)
                            Try
                                rng1 = wd.ActiveDocument.Range(Start:=tblW.Cell(3, 1).Range.Start, End:=tblW.Cell(ctRows + 3 - 1, ctCols).Range.End)
                                rng1.Select()
                            Catch ex As Exception
                                .Selection.SelectRow()
                                .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=ctRows - 1, Extend:=True)
                                var1 = ex.Message
                                var1 = var1
                            End Try

                            'the paste action removes paragraph formatting, must replace it
                            'make cell alignment wdCellAlignVerticalTop
                            '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                            '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdColumn, Extend:=True)
                            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalTop 'wdCellAlignVerticalBottom

                            '20171220 LEE: Do not set table size, use the style default table
                            '.Selection.Font.Size = fontsize - 1

                            'replace '_xyz_' with chrw(11)
                            With rng1.Find
                                .ClearFormatting()
                                .Text = strR
                                .Replacement.ClearFormatting()
                                '.Replacement.Text = "," & ChrW(11)
                                .Replacement.Text = ChrW(11)
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
                        End If


                        '*****

                        'delete extra rows
                        'remove unused rows
                        Call RemoveRows(wd, 1)

                        ctrsReassayed(1, Count2A) = intN

                    Catch ex As Exception

                        str1 = "There was a problem preparing table:"
                        str1 = strM1 & ChrW(10) & ChrW(10) & str1
                        str1 = str1 & ChrW(10) & ChrW(10)
                        str1 = str1 & ex.Message
                        MsgBox(str1, vbInformation, "Problem...")

                    End Try

                    Dim wtbl1 As Word.Table
                    wtbl1 = .Selection.Tables.Item(1)


                    'record position
                    'strTName = strTName.Replace("[MATRIX]", strMatrix) 'Need to do this here for now
                    '20160221 LEE: Made function to update. 
                    strTName = UpdateAnalyteMatrix(strTName, strAnalyteDescription, strMatrix, False, 0, False)
                    Call EnterTableNumber(wd, strTName, 3, strAnalyteDescription, strTempInfo, intTableID, 1, idTR)
                    var1 = dvDo(intDo).Item("CHARHEADINGTEXT") 'Then change it back
                    strTName = NZ(var1, "[NONE]")

                    'enter a table record in tblTableN
                    'ctTableN = ctTableN + 1
                    Dim dtblr As DataRow = tblTableN.NewRow
                    dtblr.BeginEdit()
                    dtblr.Item("TableNumber") = ctTableN
                    dtblr.Item("AnalyteName") = strAnalyteDescription 'arrAnalytes(1, Count2A)
                    dtblr.Item("TableName") = strTNameO
                    dtblr.Item("TableID") = intTableID
                    dtblr.Item("CHARFCID") = charFCID
                    dtblr.Item("TableNameNew") = strTName
                    tblTableN.Rows.Add(dtblr)


                    posrow1 = .Selection.Tables.Item(1).Rows.Count + 2

                    'enter arrlegenD for reasons
                    int1 = 1
                    If int1 > UBound(arrLegend, 2) Then
                        ReDim Preserve arrLegend(UBound(arrLegend, 1), UBound(arrLegend, 2) + 10)
                    End If
                    arrLegend(1, int1) = ""
                    arrLegend(2, int1) = ""
                    arrLegend(3, int1) = False

                    '20180328 LEE: Need to account for boolShowReasReason
                    If boolShowReasReason Then

                        int1 = int1 + 1
                        If int1 > UBound(arrLegend, 2) Then
                            ReDim Preserve arrLegend(UBound(arrLegend, 1), UBound(arrLegend, 2) + 10)
                        End If
                        arrLegend(1, int1) = ""
                        arrLegend(2, int1) = "REASON FOR REASSAY:"
                        arrLegend(3, int1) = False
                        For Count4 = 1 To ctReasons
                            int1 = int1 + 1
                            If int1 > UBound(arrLegend, 2) Then
                                ReDim Preserve arrLegend(UBound(arrLegend, 1), UBound(arrLegend, 2) + 10)
                            End If
                            arrLegend(1, int1) = Count4
                            var1 = arrReasons(Count4)
                            If StrComp(var1, "AQL", CompareMethod.Text) = 0 Then
                                arrReasons(Count4) = "Above Quantification Limit (AQL)"
                            ElseIf StrComp(var1, "ALQ", CompareMethod.Text) = 0 Then
                                arrReasons(Count4) = "Above Limit of Quantification (ALQ)"
                            ElseIf StrComp(var1, "BQL", CompareMethod.Text) = 0 Then
                                arrReasons(Count4) = "Below Quantification Limit (BQL)"
                            ElseIf StrComp(var1, "BLQ", CompareMethod.Text) = 0 Then
                                arrReasons(Count4) = "Below Limit of Quantification Limit (BLQ)"
                            End If
                            arrLegend(2, int1) = arrReasons(Count4)
                            arrLegend(3, int1) = False
                        Next

                    End If

                    If boolShowConcReason Then

                        If boolShowReasReason Then
                            int1 = int1 + 1
                            If int1 > UBound(arrLegend, 2) Then
                                ReDim Preserve arrLegend(UBound(arrLegend, 1), UBound(arrLegend, 2) + 10)
                            End If
                            arrLegend(1, int1) = ""
                            arrLegend(2, int1) = ""
                            arrLegend(3, int1) = False
                        End If

                        int1 = int1 + 1
                        If int1 > UBound(arrLegend, 2) Then
                            ReDim Preserve arrLegend(UBound(arrLegend, 1), UBound(arrLegend, 2) + 10)
                        End If
                        arrLegend(1, int1) = ""
                        arrLegend(2, int1) = "REASON FOR REPORTED CONCENTRATION:"
                        arrLegend(3, int1) = False
                        For Count4 = 1 To ctReasonsC
                            int1 = int1 + 1
                            If int1 > UBound(arrLegend, 2) Then
                                ReDim Preserve arrLegend(UBound(arrLegend, 1), UBound(arrLegend, 2) + 10)
                            End If
                            var1 = arrLegendC(1, Count4) 'debug
                            arrLegend(1, int1) = arrLegendC(1, Count4)
                            arrLegend(2, int1) = arrLegendC(2, Count4)
                            arrLegend(3, int1) = False
                        Next

                        ReDim Preserve arrLegend(4, int1)

                    End If

                    str1 = frmH.lblProgress.Text

                    wtbl1.Select()

                    Dim boolOld As Boolean = True
                    'strM = "Do old (Yes) or new (No)"
                    'intR = MsgBox(strM, vbYesNo)
                    'If intR = 6 Then
                    '    boolOld = True
                    'Else
                    '    boolOld = False
                    'End If
                    boolOld = False

                    Dim ctLegends As Short = int1

                    '*****



                    '*****

                    'autofit table
                    '20180201 LEE: Make boolVis TRUE
                    Call AutoFitTable(wd, True)

                    strM = "Finalizing " & strTName & "..."
                    strM1 = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    str1 = strM1

                    frmH.lblProgress.Text = strM1
                    frmH.Refresh()

                    'tttt
                    '20160107 Larry: Call SplitTable, but don't add legend. Wait till end since legend can be very large.

                    If boolNA Then
                        ctLegend1 = ctLegend1 + 1
                        arrLegend1(1, ctLegend1) = "NA"
                        arrLegend1(2, ctLegend1) = "Not Applicable"
                        arrLegend1(3, ctLegend1) = False
                        arrLegend1(4, ctLegend1) = False
                    End If

                    If boolNR Then
                        ctLegend1 = ctLegend1 + 1
                        arrLegend1(1, ctLegend1) = "NR"
                        arrLegend1(2, ctLegend1) = "Not Reportable"
                        arrLegend1(3, ctLegend1) = False
                        arrLegend1(4, ctLegend1) = False
                    End If

                    ctLegend1 = ctLegend1 + 1

                    arrLegend1(1, ctLegend1) = BQL()
                    If boolBQLLEGEND Then
                        If boolLUseSigFigs Then
                            arrLegend1(2, ctLegend1) = BQLVerbose() & " (" & DisplayNum(SigFigOrDec(numLLOQ / numDF, LSigFig, False), LSigFig, False) & " " & strConcUnits & ")"
                        Else
                            arrLegend1(2, ctLegend1) = BQLVerbose() & " (" & Format(SigFigOrDec(numLLOQ, LSigFig / numDF, False), GetRegrDecStr(LSigFig)) & " " & strConcUnits & ")"
                        End If
                    Else
                        arrLegend1(2, ctLegend1) = BQLVerbose()
                    End If

                    arrLegend1(3, ctLegend1) = False
                    arrLegend1(4, ctLegend1) = False

                    ctLegend1 = ctLegend1 + 1

                    arrLegend1(1, ctLegend1) = AQL()
                    'If boolBQLLEGEND Then

                    'Else
                    '    arrLegend(2, ctLegend) = "Above Quantitation Limit"
                    'End If
                    If boolBQLLEGEND Then
                        If boolLUseSigFigs Then
                            arrLegend1(2, ctLegend1) = AQLVerbose() & " (" & DisplayNum(SigFigOrDec(numULOQ / numDF, LSigFig, False), LSigFig, False) & " " & strConcUnits & ")"
                        Else
                            arrLegend1(2, ctLegend1) = AQLVerbose() & " (" & Format(SigFigOrDec(numULOQ / numDF, LSigFig, False), GetRegrDecStr(LSigFig)) & " " & strConcUnits & ")"
                        End If
                    Else
                        arrLegend1(2, ctLegend1) = AQLVerbose()
                    End If

                    arrLegend1(3, ctLegend1) = False
                    arrLegend1(4, ctLegend1) = False

                    ReDim Preserve arrLegend1(4, ctLegend1)

                    'Call SplitTable(wd, 2, ctLegend, arrLegend, str1, False, ctLegend + 1, False, False, False, intTableID)

                    Call SplitTable(wd, 2, ctLegend1, arrLegend1, str1, False, int1 + 1, False, False, False, intTableID)

                    'If boolBQLLEGEND Then

                    'Else
                    '    If boolOld Then '
                    '        Call SplitTable(wd, 3, ctLegends, arrLegend, str1, True, int1 + 1, False, False, False, intTableID)
                    '    Else
                    '        Call SplitTable(wd, 3, 0, arrLegend, str1, False, int1 + 1, False, False, False, intTableID)
                    '    End If
                    'End If


                    'autofit table
                    Call AutoFitTable(wd, False)

                    ''Sub SplitTable(ByVal wd As Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, 
                    ''ByVal arr As Object, ByVal strT As String, ByVal DoLegend As Boolean, ByVal intSplitRows As Short, ByVal boolSmallFont As Boolean)

                    'record posrow2
                    posrow2 = .Selection.Tables.Item(1).Rows.Count
                    int1 = .Selection.Tables.Item(1).Columns.Count


                    'Here!!!
                    ''''''''wdd.visible = True

                    If (boolShowConcReason Or boolShowReasReason) And boolOld = False Then

                        Dim myTable As Microsoft.Office.Interop.Word.Table
                        Dim myRange As Microsoft.Office.Interop.Word.Range
                        Dim intRows As Short

                        myTable = .Selection.Tables.Item(1)

                        intRows = myTable.Rows.Count

                        myTable.Cell(intRows, 1).Select()
                        int8 = intRows

                        'insert below ctReasonsC + 2
                        '.Selection.InsertRowsBelow(ctReasonsC + 2)
                        .Selection.InsertRowsBelow(ctLegends)

                        'at this point, all inserted rows are selected
                        'if line space after is > 3, then reduce it to three
                        Dim rngI As Word.Range
                        rngI = .Selection.Range
                        var1 = rngI.ParagraphFormat.SpaceAfter
                        If var1 > 3 Then
                            rngI.ParagraphFormat.SpaceAfter = 3
                        End If

                        'clear borders
                        .Selection.Borders.Enable = False

                        'ensure the last line has a border
                        myTable.Cell(int8, 1).Select()
                        .Selection.SelectRow()

                        Dim intC As Short
                        intC = .Selection.Cells.Count
                        If intC = 1 Then
                        Else
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        End If


                        ''start in 2nd row down
                        'int8 = int8 + 1

                        For Count2 = 1 To ctLegends
                            myTable.Cell(int8 + Count2, 1).Select()
                            .Selection.SelectRow()
                            Try
                                .Selection.Cells.Merge()
                            Catch ex As Exception

                            End Try

                            'set tabs
                            Try
                                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                                .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
                                With .Selection.ParagraphFormat
                                    .LeftIndent = 28 'InchesToPoints(0.38)
                                    .SpaceBeforeAuto = False
                                    .SpaceAfterAuto = False
                                End With
                                With .Selection.ParagraphFormat
                                    .SpaceBeforeAuto = False
                                    .SpaceAfterAuto = False
                                    .FirstLineIndent = -28 'InchesToPoints(-0.38)
                                End With
                                .Selection.ParagraphFormat.TabStops.Add(Position:=28, _
                                    Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                            Catch ex As Exception

                            End Try
                        Next
                        'int8 = int8 + 1
                        'int8 = int8 + 1
                        myTable.Cell(int8, 1).Select()

                        For Count2 = 1 To ctLegends

                            int8 = int8 + 1
                            myTable.Cell(int8, 1).Select()
                            'Herehere

                            var1 = arrLegend(1, Count2)
                            '                var1 = Replace(CStr(var1), "-", NBH, 1, -1, CompareMethod.Text)
                            var1 = Replace(CStr(var1), "-", NBH, 1, -1, vbTextCompare)
                            If arrLegend(3, Count2) Then
                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
                                .Selection.Font.Bold = False
                                typeInSuperscript(wd, CStr(var1))
                            Else
                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
                                .Selection.Font.Bold = False
                                .Selection.TypeText(Text:=CStr(var1))
                            End If
                            If Len(var1) = 0 Then
                            Else
                                .Selection.TypeText(Text:=vbTab)
                                .Selection.TypeText(Text:="=")
                                .Selection.TypeText(Text:=vbTab)
                            End If
                            var2 = arrLegend(2, Count2)
                            var2 = Replace(CStr(var2), "-", NBH, 1, -1, vbTextCompare)
                            '                var2 = Replace(CStr(var2), "-", NBH, 1, -1, CompareMethod.Text)
                            .Selection.TypeText(Text:=CStr(var2))

                        Next

                    End If

                    Call MoveOneCellDown(wd)

                    Call InsertLegend(wd, intTableID, idTR, False, 1)

                    var1 = var1 'debug

next1:
                Next Count2A

            Next Count1A

        End With

    End Sub


End Module
