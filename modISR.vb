
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.ComponentModel.PropertyDescriptorCollection
Imports Word = Microsoft.Office.Interop.Word
Imports Microsoft.VisualBasic
Imports System.IO



Module modISR

    Public numISRPercent As Single = 0
    Public ctISRTot As Int16 = 0
    Public ctISRPass As Int16 = 0

    Sub ISR_01_02_30(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal idTR As Int64)

        'Agilux has given us 7.4 data (D:\GoogleDrives\LabIntegrity\GoogleDrive\LabIntegrity\Clients\Agilux\Projects\StudyDoc\20140205_ExampleData\SampleAnalysis\IY0009RBBS\), 
        'so I inspected the tables. What I found was not a new sample type, but a new column called “ANALYSISTYPE” in table ANALYTICALRUNSAMPLE. 
        'In the Agilux data, all records were NULL for this column, except for the ISR samples, whose entry is ‘ISR’.

        'SAMPLERESULTSCONFLICT	Contains accepted reassay samples, original results by run id and seq number.
        'SAMPRESCONFLICTCHOICES	Contains the reassay samples with mean or median choices for the Reassay. 
        'SAMPRESCONFLICTDEC	Contains the concentration decision choices for the Reassay Selection function. of Watson.

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

        Dim numDF As Decimal

        Dim boolConcStatus As Boolean = False

        'The following from original Excel code

        ''****Start Re-assay Samples
        Dim arrReassayRows(100)
        Dim arrReassay(9, 100)
        Dim arrReasons(100)
        Dim arrReasonsC(100)
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
        Dim var1, var2, var3, var4, var5, var6, var7
        Dim lng1 As Int64
        Dim lng2 As Int64
        Dim dv As System.Data.DataView

        Dim strConcUnits As String
        Dim boolJustTable As Boolean = False

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
        Dim strPaste As String
        Dim strPasteT As String

        Dim intExp As Short

        Dim fonts
        Dim fontsize

        Dim boolShowConcReason As Boolean = False
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

        Dim vT, vH, vM, vS
        Dim vTS, vHS, vMS, vSS
        Dim strColLabel As String
        Dim arrMean(100)
        Dim intMean As Short
        Dim numMean As Decimal = 0
        Dim numMeanTot As Decimal = 0
        Dim numISR As Decimal


        Dim tblRepeatTableRows As New DataTable
        Dim dvRepeatAllRunSamples, dvRepeatTableRows As DataView
        Dim dvCalStdGroupAssayIDsAcc As New DataView(tblCalStdGroupAssayIDsAcc)
        Dim dvDuplicates As DataView
        Dim tblDuplicates As New DataTable
        Dim ctDuplicates As Short
        Dim boolFirstLine, boolFirstLineEntry As Boolean

        Dim strSdvRepeatTableRowsSort As String = "" ' "DESIGNSUBJECTTAG ASC, WEEK ASC, ENDDAY ASC, ENDHOUR ASC, SAMPLERESULTS.RUNID ASC"
        Dim strSdvRepeatAllRunSamplesSort As String = "" ' strSdvRepeatTableRowsSort ' "DESIGNSUBJECTTAG ASC, WEEK ASC, ENDDAY ASC, ENDHOUR ASC, SAMPLERESULTS.RUNID ASC, RUNSAMPLESEQUENCENUMBER ASC"

        Dim dup1, dup2, numMeanDup, numOrig, numRV As Object
        Dim numSpaceAfter As Single
        Dim numSpaceAfterNew As Single

        Dim strStartDay As String = ""
        Dim strStartHour As String = ""
        Dim strStartMinute As String = ""

        '1=DESIGNSAMPLEID, 2=ANALYTEID, 3=DESIGNSUBJECTTAG, 4=TimePoint, 5=RUNID, 6=RUNSAMPLEORDERNUMBER, 7=DECISIONCODE

        Dim boolBQL As Boolean = False
        Dim boolAQL As Boolean = False
        Dim boolNA As Boolean = False
        Dim boolNR As Boolean = False

        Dim numLLOQ, numULOQ As Decimal
        Dim strBQL, strAQL As String
        Dim numBQL As Decimal

        Dim strTNameO As String 'original Table Name

        Dim strA As String
        Dim strB As String

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
            intTableID = 30
            dvDo = frmH.dgvReportTableConfiguration.DataSource
            intDo = FindRowDVNumByCol(intTableID, dvDo, "id_tblconfigreporttables")

            Dim strSString As String
            Dim strGroupCheck As String
            'find idt of this study
            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLCONFIGREPORTTABLES = 5"
            Dim rowsTT() As DataRow = tblReportTable.Select(strF)
            Dim idTTT As Int64
            idTTT = rowsTT(0).Item("ID_TBLREPORTTABLE")

            'get sort from SampleConcentrations table
            Try
                Call GetGroupSort(idTTT) 'retrieve grouping and sorting information
            Catch ex As Exception
                var1 = var1 'debug
            End Try
            strSString = GetSString()


            ''Get table name
            'var1 = dvDo(intDo).Item("Table")
            'strTName = NZ(var1, "[NONE]")

            '***
            intDo = FindRowDVNumByCol(idTR, dvDo, "ID_TBLREPORTTABLE")
            'intLeg = 0
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


            '*** Start Duplicate Logic 1 ****

            '20161109 LEE: It looks like Logic 1 isn't needed in ISR table, will leave here a bit longer

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

            strSdvRepeatTableRowsSort = "DESIGNSUBJECTTAG ASC, WEEK ASC, ENDDAY ASC, ENDHOUR ASC, ENDMINUTE ASC, ENDSECOND ASC, SAMPLERESULTS_RUNID ASC"
            strSdvRepeatAllRunSamplesSort = strSdvRepeatTableRowsSort
            'strSString = "DESIGNSUBJECTTAG ASC, WEEK ASC, ENDDAY ASC, ENDHOUR ASC, ENDMINUTE ASC, ENDSECOND ASC, SRUNID ASC"

            tblRepeatTableRows = dvRepeatAllRunSamples.ToTable("tblRepeatTableRows", True, "ANALYTEID", "DESIGNSUBJECTTAG", "DESIGNSAMPLEID", "SUBJECTGROUPNAME",
                                                               "USERSAMPLEID", "TREATMENTEVENTID", "ENDDAY", "ENDHOUR", "ENDMINUTE", "ENDSECOND", "STUDYID",
                                                               "REASSAYREASON", "REASSAYCONCREASON", "DECISIONCODE", "SAMPLETYPEID", "SAMPLERESULTS_CONCENTRATION",
                                                               "CALIBRATIONRANGEFLAG", "SAMPLERESULTS_ALIQUOTFACTOR", "SAMPLERESULTS_RUNID", "WEEK", "VISITTEXT", "STARTDAY", "STARTHOUR", "STARTMINUTE", "STARTSECOND")


            dvRepeatTableRows = New DataView(tblRepeatTableRows)
            'need to sort here since original recordset sort doesn't seem to do the trick consistently
            Try
                dvRepeatTableRows.Sort = strSdvRepeatTableRowsSort
            Catch ex As Exception
                var1 = var1
            End Try


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


            Dim strAnalyteID, strAnalyteDescription As String
            Dim strMatrix As String
            Dim intGroup As Short
            Dim strAnal As String
            Dim strAnalC As String
            Dim AnalyteID As Int64

            If boolCanDoISR Or boolPlaceHolder Then
            Else

                strM = "This data is stored in Watson " & strWatsonVersion & " which does not support ISR samples."
                strM = strM & ChrW(10) & "Table placeholder(s) will be made for this table."
                MsgBox(strM, vbInformation, "Invalid action...")

            End If

            For Count1A = 0 To 0 '20180219 LEE:' tbl1A.Rows.Count - 1 'Iterate through each Matrix (but keep different calibration ranges together)

                ctISRTot = 0
                ctISRPass = 0

                If boolM Then
                    strMatrix = tblMatrices.Rows(Count1A).Item("Matrix")
                Else
                    strAnalyteID = tblAnalyteIDs.Rows(Count1A).Item("AnalyteID")
                    strAnalyteDescription = tblAnalyteIDs.Rows(Count1A).Item("AnalyteDescription")
                    AnalyteID = CInt(strAnalyteID)
                End If

                For Count2A = 0 To intRowsAnal - 1 '20180219 LEE' tbl2A.Rows.Count - 1 'Iterate through each AnalyteID, and generate the information

                    strTName = strTNameO 'reset strTName

                    Dim arrLegend(4, 1000) 'Reason for Reassay
                    Dim arrLegendC(3, 1000) 'Reason for Reported Concentration

                    strA = ""
                    strB = ""

                    'If boolM Then
                    '    strAnalyteID = tblAnalyteIDs.Rows(Count2A).Item("AnalyteID")
                    '    strAnalyteDescription = tblAnalyteIDs.Rows(Count2A).Item("AnalyteDescription")
                    '    AnalyteID = CInt(strAnalyteID)
                    'Else
                    '    strMatrix = tblMatrices.Rows(Count2A).Item("Matrix")
                    'End If

                    '20180216 LEE: 
                    Try
                        strAnalyteID = rows11(Count2A).Item("AnalyteID")
                        AnalyteID = CInt(strAnalyteID)
                        strAnalyteDescription = rows11(Count2A).Item("OriginalAnalyteDescription")
                        strMatrix = rows11(Count2A).Item("Matrix")
                        strConcUnits = rows11(Count2A).Item("ConcUnits")
                        'intGroup = rows11(Count2A).Item("INTGROUP")
                    Catch ex As Exception
                        var1 = var1 'debug
                    End Try

                    If (Not (boolGenerateTableForThisAnalyteIDandMatrix(intDo, strAnalyteID, strMatrix))) Then
                        GoTo next2
                    End If

                    If boolCanDoISR Then
                    Else
                        str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                        Call InsertPageBreak(wd)
                        Call PageSetup(wd, str1) 'L=Landscape, P=Portrait
                        boolJustTable = True
                        GoTo next1
                    End If

                    Dim numCrit As Decimal
                    numCrit = ReturnISRCrit1(strAnalyteID, idTR)

                    ''20160311 LEE:
                    ''get units from tblanalyteshome
                    'Dim strF1 As String
                    'strF1 = "ANALYTEID = " & strAnalyteID & " AND MATRIX = '" & strMatrix & "'"
                    'Dim rowsUnits() As DataRow = tblAnalytesHome.Select(strF1)
                    'strConcUnits = rowsUnits(0).Item("ConcUnits")

                    'gstrAnal = arrAnalytes(1, Count2A)  'Only used for Regression Equation; not correct with Analyte Groups
                    'gnumAnal = Count2A

                    strAnalyteAndMatrixFilter = "SAMPLETYPEID = '" & strMatrix & "' AND ANALYTEID = " & strAnalyteID

                    'If (Not (boolGenerateTableForThisAnalyteIDandMatrix(intDo, strAnalyteID, strMatrix))) Then
                    '    GoTo next1
                    'End If

                    intTCur = intTCur + 1

                    strM = "Creating " & strTName & " For " & strAnalyteDescription & "..."
                    strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    strM1 = strM
                    frmH.lblProgress.Text = strM
                    frmH.Refresh()


                    'first get unique items for table
                    'NDL 6-Feb-2016: Filter for Matrix, but don't filter for calibration sets.  Even if a calibration 
                    ' set sub-analyte is not selected, a repeat may include repeat samples with the different LOQ's.  
                    ' So we report all calibration sets regardless of selection.
                    strF = makeRunMatrixFilter(intDo, strAnalyteID, strMatrix) ' & " AND ORIGINALVALUE = 'Y'"
                    strS = "DESIGNSUBJECTTAG ASC, RUNID ASC, ENDDAY ASC, ENDHOUR ASC"
                    Dim dv1 As System.Data.DataView = New DataView(tblISR, strF, strS, DataViewRowState.CurrentRows)
                    'Dim tblUGrps As System.Data.DataTable = dv1.ToTable("a", True, "DESIGNSUBJECTTAG", "ENDDAY", "ENDHOUR", "ENDMINUTE", "RUNID", "DESIGNSAMPLEID")
                    Dim tblUGrps As System.Data.DataTable = dv1.ToTable("a", True, "DESIGNSUBJECTTAG", "ENDDAY", "ENDHOUR", "ENDMINUTE", "RUNID", "DESIGNSAMPLEID", "STARTDAY", "STARTHOUR", "STARTMINUTE")
                    Dim intUGrps As Short
                    intUGrps = tblUGrps.Rows.Count

                    'determine number of blank rows to skip
                    Dim tblUSubj As System.Data.DataTable = dv1.ToTable("a", True, "DESIGNSUBJECTTAG")
                    'Nope. Need to keep Design Subject separate
                    'Dim tblUSubj As System.Data.DataTable = dv1.ToTable("a", True, "DESIGNSUBJECTTAG", "DESIGNSAMPLEID")
                    Dim intUSubj As Short
                    intUSubj = tblUSubj.Rows.Count

                    If intUSubj = 0 Or intUGrps = 0 Then 'no reassayed samples
                        GoTo next1
                    End If

                    'page setup according to configuration
                    str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                    Dim boolPortrait As Boolean = True
                    If StrComp(str1, "P", CompareMethod.Text) = 0 Then
                        boolPortrait = True
                    Else
                        boolPortrait = False
                    End If

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
                        arrOrder(4, Count2) = NZ(dr1(Count2 - 1).Item("charUserLabel"), "")
                        str2 = dr1(Count2 - 1).Item("charUserLabel") 'just checking
                        arrOrder(5, Count2) = dr1(Count2 - 1).Item("id_tblConfigReportTables")
                        'find column label
                        str1 = "id_tblConfigHeaderLookup = " & arrOrder(3, Count2) & " AND id_tblConfigReportTables = " & intTableID
                        dr = tbl.Select(str1)
                        var1 = dr(0).Item("CHARCOLUMNLABEL") 'debug
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

                            'populate tblreassayreport with Reason Number
                            str1 = "ANALYTEID = " & strAnalyteID & " AND DESIGNSUBJECTTAG = '" & strSubj & "' AND DESIGNSAMPLEID = " & intDSId
                            dr2 = tblReassayReport.Select(str1)

                            If dr2.Length = 0 Then
                            Else
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

                    For Count4 = 1 To intUSubj

                        strSubj = tblUSubj.Rows(Count4 - 1).Item("DESIGNSUBJECTTAG")

                        strF = "DESIGNSUBJECTTAG = '" & strSubj & "'"
                        rowsUG = tblUGrps.Select(strF, strS)
                        intL = rowsUG.Length
                        intL = intL 'debug

                        For Count6 = 1 To intL

                            intDSId = rowsUG(Count6 - 1).Item("DESIGNSAMPLEID")

                            'populate tblreassayreport with Conc Reason Number
                            str1 = "ANALYTEID = " & strAnalyteID & " AND DESIGNSUBJECTTAG = '" & strSubj & "' AND DESIGNSAMPLEID = " & intDSId
                            dr2 = tblReassayReport.Select(str1)
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
                                    arrLegendC(1, intCt) = intCt ' Count2 + 1 'intCt
                                    arrLegendC(2, intCt) = dr2(0).Item("REASSAYCONCREASON")
                                    arrLegendC(3, intCt) = False
                                    arrReasonsC(intCt) = dr2(0).Item("REASSAYCONCREASON")
                                    int1 = intCt
                                End If

                                For Count3 = 0 To dr2.Length - 1
                                    dr2(Count3).BeginEdit()
                                    dr2(Count3).Item("numRCR") = int1 ' intCt 'Count2 + 1
                                    dr2(Count3).EndEdit()
                                Next

                                '****
                            End If

                        Next


                    Next

                    ctReasonsC = intCt 'int2

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
                    Dim rows() As DataRow
                    intTRows = 0
                    intTRows = intTRows + 1 'for table header
                    intTRows = intTRows + 1 'for blank row
                    intTRows = intTRows + intUGrps 'for unique groups
                    intTRows = intTRows + intUSubj - 1 'for blank row between groups

                    ctrsReassayed(1, Count2A) = intUGrps

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

                            GoTo next1
                        End If

                        .Selection.Tables.Item(1).Select()
                        Call GlobalTableParaFormat(wd)

                        '20171220 LEE:
                        'Do not set table size, use the style default table
                        '.Selection.Font.Size = fontsize - 1
                        .Selection.Tables.Item(1).Cell(1, 1).Select()

                        'enter headings
                        For Count4 = 1 To ctCols
                            .Selection.Tables.Item(1).Cell(1, Count4).Select()
                            var1 = arrOrder(1, Count4)
                            var2 = arrOrder(4, Count4)
                            Select Case var1
                                Case "Original Conc."
                                    var2 = var2 & ChrW(10) & "(" & strConcUnits & ")"
                                Case "ISR Conc."
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


                        int8 = 1


                        Dim strF2 As String
                        Dim strF3 As String
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
                        Dim numPD As Single

                        Dim intCtLbl As Int32 = 0
                        Dim maxCtLbl As Short = 25

                        Dim intN As Int16 = 0

                        Dim ctRows As Short = 0
                        Dim strL As String

                        strL = frmH.lblProgress.Text

                        Dim rowsISR() As DataRow
                        Dim intRowsISR As Int16
                        Dim rowsDS() As DataRow

                        Dim boolDoMean As Boolean = False

                        Try

                            'filter tblISR for analyteid and matrix
                            strF2 = "ANALYTEID = " & strAnalyteID & " AND SAMPLETYPEID = '" & strMatrix & "'"
                            'sort according to Sample Conc table
                            rowsISR = tblISR.Select(strF2, strSString)
                            intRowsISR = rowsISR.Length


                            For Count4 = 0 To intRowsISR - 1

                                intDSId = rowsISR(Count4).Item("DESIGNSAMPLEID")

                                'filter tbldesignsample for intdsid
                                strF = strF2 & " AND DESIGNSAMPLEID = " & intDSId
                                rowsDS = tblSampleDesign.Select(strF)

                                'Now record data

                                intMean = 0
                                boolDoMean = False
                                boolConcStatus = False

                                For Count3 = 1 To ctCols
                                    varE = ""
                                    strColLabel = arrOrder(1, Count3)
                                    Select Case strColLabel
                                        Case "Sample Count"
                                            var1 = Count4 + 1
                                            varE = var1.ToString
                                        Case "Subject"
                                            varE = rowsISR(Count4).Item("DESIGNSUBJECTTAG")
                                        Case "Custom ID"
                                            varE = rowsISR(Count4).Item("USERSAMPLEID")
                                        Case "Treatment"
                                            varE = rowsISR(Count4).Item("TREATMENTDESC")
                                        Case "Day"
                                            str1 = CStr(NZ(rowsISR(Count4).Item("ENDDAY"), 0)) '
                                            strStartDay = CStr(NZ(rowsISR(Count4).Item("STARTDAY"), "")) '
                                            If Len(strStartDay) = 0 Then
                                                varE = str1
                                            Else
                                                varE = str1 & " to " & strStartDay
                                            End If
                                        Case "Time"
                                            'these are endhour and endminutes
                                            var1 = NZ(rowsISR(Count4).Item("ENDHOUR"), 0)
                                            vH = RoundToDecimal(var1, 3)
                                            var2 = NZ(rowsISR(Count4).Item("ENDMINUTE"), 0)
                                            vM = RoundToDecimal(var2 / 60, 3)

                                            vT = vH + vM
                                            str1 = vT & "h"

                                            varE = str1
                                            varE = varE 'debug

                                            'look for StartHour and StartMinute
                                            strStartHour = CStr(NZ(rowsISR(Count4).Item("STARTHOUR"), "")) '
                                            strStartMinute = CStr(NZ(rowsISR(Count4).Item("STARTMINUTE"), "")) '

                                            If Len(strStartHour) <> 0 Or Len(strStartMinute) <> 0 Then

                                                var1 = NZ(rowsISR(Count4).Item("STARTHOUR"), 0)
                                                vHS = RoundToDecimal(var1, 3)
                                                var2 = NZ(rowsISR(Count4).Item("STARTMINUTE"), 0)
                                                vMS = RoundToDecimal(var2 / 60, 3)

                                                vTS = vHS + vMS
                                                str1 = vTS & "h"

                                                varE = str1 & " to " & varE

                                            End If

                                        Case "Original Conc."

                                            numOrig = -1

                                            If rowsDS.Length = 0 Then
                                                varE = "NA"
                                                boolNA = True
                                            Else

                                                '20180219 LEE: (0) probably isn't correct
                                                'getting wrong RunID
                                                var1 = rowsDS(0).Item("CONCENTRATION")
                                                var4 = NZ(rowsDS(0).Item("ALIQUOTFACTOR"), 1)
                                                numDF = var4
                                                var2 = rowsDS(0).Item("CALIBRATIONRANGEFLAG")
                                                var3 = rowsDS(0).Item("CALIBRATIONRANGE")
                                                numOrig = -1

                                                If IsDBNull(var2) Then

                                                    'get concentration
                                                    If IsDBNull(var1) Then
                                                        varE = "NS"
                                                    Else
                                                        If var4 = 0 Then
                                                            var5 = var1
                                                        Else
                                                            var5 = var1 / var4
                                                        End If
                                                        numOrig = SigFigOrDec(var5, LSigFig, False)
                                                        varE = SigFigOrDecString(numOrig, LSigFig, False)
                                                    End If

                                                Else

                                                    'If StrComp(var2, "NM", CompareMethod.Text) = 0 Then
                                                    '    If boolLUseSigFigs Then
                                                    '        strBQL = BQL() & "(<" & DisplayNum(SigFigOrDec(var3, LSigFig, False), LSigFig, False) & ")"
                                                    '    Else
                                                    '        strBQL = BQL() & "(<" & Format(RoundToDecimalRAFZ(var3, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                    '    End If
                                                    '    varE = strBQL
                                                    'Else
                                                    '    If boolLUseSigFigs Then
                                                    '        strAQL = AQL() & "(>" & DisplayNum(SigFigOrDec(var3, LSigFig, False), LSigFig, False) & ")"
                                                    '    Else
                                                    '        strAQL = AQL() & "(>" & Format(RoundToDecimalRAFZ(var3, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                    '    End If
                                                    '    varE = strBQL
                                                    'End If


                                                    '*****

                                                    If IsNumeric(var2) Then
                                                        If boolLUseSigFigs Then
                                                            str1 = DisplayNum(SigFigOrDec(var2, LSigFig, False), LSigFig, False)
                                                        Else
                                                            str1 = Format(RoundToDecimalRAFZ(var2, LSigFig), GetRegrDecStr(LSigFig))
                                                        End If
                                                    Else
                                                        If boolLUseSigFigs Then
                                                            str1 = DisplayNum(SigFigOrDec(var1 / numDF, LSigFig, False), LSigFig, False)
                                                        Else
                                                            str1 = Format(RoundToDecimalRAFZ(var1 / numDF, LSigFig), GetRegrDecStr(LSigFig))
                                                        End If
                                                        str1 = str1
                                                    End If


                                                    If StrComp(var2, "NM", CompareMethod.Text) = 0 Then 'bql

                                                        If boolBQLSHOWCONC Then
                                                            If boolBQLLEGEND Then
                                                                If boolLUseSigFigs Then
                                                                    strBQL = str1 & " (" & BQL() & ")"
                                                                Else
                                                                    strBQL = str1 & " (" & BQL() & ")"
                                                                End If
                                                            Else
                                                                If boolLUseSigFigs Then
                                                                    strBQL = str1 & strR1 & BQL() & "(<" & DisplayNum(SigFigOrDec(var3 / numDF, LSigFig, False), LSigFig, False) & ")"
                                                                Else
                                                                    strBQL = str1 & strR1 & BQL() & "(<" & Format(RoundToDecimalRAFZ(var3 / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                                End If
                                                            End If

                                                        Else

                                                            If boolBQLLEGEND Then
                                                                strBQL = BQL()
                                                            Else
                                                                If boolLUseSigFigs Then
                                                                    strBQL = BQL() & "(<" & DisplayNum(SigFigOrDec(var3 / numDF, LSigFig, False), LSigFig, False) & ")"
                                                                Else
                                                                    strBQL = BQL() & "(<" & Format(RoundToDecimalRAFZ(var3 / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                                End If
                                                            End If

                                                        End If

                                                        varE = strBQL

                                                    ElseIf StrComp(var2, "VEC", CompareMethod.Text) = 0 Then 'aql

                                                        If boolBQLSHOWCONC Then
                                                            If boolBQLLEGEND Then
                                                                If boolLUseSigFigs Then
                                                                    strAQL = str1 & " (" & AQL() & ")"
                                                                Else
                                                                    strAQL = str1 & " (" & AQL() & ")"
                                                                End If
                                                            Else
                                                                If boolLUseSigFigs Then
                                                                    strAQL = str1 & strR1 & AQL() & "(>" & DisplayNum(SigFigOrDec(var3 / numDF, LSigFig, False), LSigFig, False) & ")"
                                                                Else
                                                                    strAQL = str1 & strR1 & AQL() & "(>" & Format(RoundToDecimalRAFZ(var3 / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                                End If
                                                            End If

                                                        Else
                                                            If boolBQLLEGEND Then
                                                                strAQL = AQL()
                                                            Else
                                                                If boolLUseSigFigs Then
                                                                    strAQL = AQL() & "(>" & DisplayNum(SigFigOrDec(var3 / numDF, LSigFig, False), LSigFig, False) & ")"
                                                                Else
                                                                    strAQL = AQL() & "(>" & Format(RoundToDecimalRAFZ(var3 / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                                End If
                                                            End If

                                                        End If

                                                        var3 = strAQL

                                                    End If

                                                    '*****

                                                End If
                                            End If


                                        Case "Original Run ID"

                                            If rowsDS.Length = 0 Then
                                                var1 = "NA"
                                                boolNA = True
                                            Else
                                                var1 = NZ(rowsDS(0).Item("RUNID"), "NA")
                                            End If

                                            varE = var1
                                            If StrComp(varE.ToString, "NA", CompareMethod.Text) = 0 Then
                                                boolNA = True
                                            End If

                                        Case "ISR Conc."

                                            numISR = -1

                                            'ISR may be replicate

                                            var1 = NZ(rowsISR(Count4).Item("CONCENTRATION"), "NA")
                                            var2 = NZ(rowsISR(Count4).Item("ALIQUOTFACTOR"), 1)
                                            numDF = var2
                                            var4 = NZ(rowsISR(Count4).Item("CONCENTRATIONSTATUS"), "") 'NM or VEC
                                            var3 = NZ(rowsISR(Count4).Item("NM"), "") 'BQL number
                                            var5 = NZ(rowsISR(Count4).Item("VEC"), "") 'AQL number
                                            '20170926 LEE: This assumes that consistent LLOQ and ULOQ are used in all retrieved runid's
                                            If IsNumeric(var3) Then
                                                numLLOQ = var3
                                            End If
                                            If IsNumeric(var5) Then
                                                numULOQ = var5
                                            End If
                                            If IsNumeric(var1) Then
                                                If var2 = 0 Then
                                                    var6 = var1
                                                Else
                                                    var6 = var1 / var2
                                                End If
                                                numISR = SigFigOrDec(var6, LSigFig, False)
                                                varE = SigFigOrDecString(numISR, LSigFig, False)

                                                '******

                                                str1 = varE

                                                If StrComp(var4, "NM", CompareMethod.Text) = 0 Then 'bql

                                                    If boolBQLSHOWCONC Then
                                                        If boolBQLLEGEND Then
                                                            If boolLUseSigFigs Then
                                                                strBQL = str1 & " (" & BQL() & ")"
                                                            Else
                                                                strBQL = str1 & " (" & BQL() & ")"
                                                            End If
                                                        Else
                                                            If IsNumeric(var3) Then
                                                                If boolLUseSigFigs Then
                                                                    strBQL = str1 & strR1 & BQL() & "(<" & DisplayNum(SigFigOrDec(var3 / numDF, LSigFig, False), LSigFig, False) & ")"
                                                                Else
                                                                    strBQL = str1 & strR1 & BQL() & "(<" & Format(RoundToDecimalRAFZ(var3 / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                                End If
                                                            Else
                                                                strBQL = str1 & strR1 & BQL() & "(<NA)"
                                                                boolNA = True
                                                            End If

                                                        End If

                                                    Else

                                                        If boolBQLLEGEND Then
                                                            strBQL = BQL()
                                                        Else

                                                            If IsNumeric(var3) Then
                                                                If boolLUseSigFigs Then
                                                                    strBQL = BQL() & "(<" & DisplayNum(SigFigOrDec(var3 / numDF, LSigFig, False), LSigFig, False) & ")"
                                                                Else
                                                                    strBQL = BQL() & "(<" & Format(RoundToDecimalRAFZ(var3 / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                                End If
                                                            Else
                                                                strBQL = str1 & strR1 & BQL() & "(<NA)"
                                                                boolNA = True
                                                            End If

                                                        End If

                                                    End If

                                                    varE = strBQL

                                                ElseIf StrComp(var4, "VEC", CompareMethod.Text) = 0 Then 'aql

                                                    If boolBQLSHOWCONC Then
                                                        If boolBQLLEGEND Then
                                                            If boolLUseSigFigs Then
                                                                strAQL = str1 & " (" & AQL() & ")"
                                                            Else
                                                                strAQL = str1 & " (" & AQL() & ")"
                                                            End If
                                                        Else

                                                            If IsNumeric(var5) Then
                                                                If boolLUseSigFigs Then
                                                                    strAQL = str1 & strR1 & AQL() & "(>" & DisplayNum(SigFigOrDec(var5 / numDF, LSigFig, False), LSigFig, False) & ")"
                                                                Else
                                                                    strAQL = str1 & strR1 & AQL() & "(>" & Format(RoundToDecimalRAFZ(var5 / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                                End If
                                                            Else
                                                                strAQL = str1 & strR1 & AQL() & "(>NA)"
                                                                boolNA = True
                                                            End If

                                                        End If

                                                    Else
                                                        If boolBQLLEGEND Then
                                                            strAQL = AQL()
                                                        Else
                                                            If IsNumeric(var5) Then
                                                                If boolLUseSigFigs Then
                                                                    strAQL = AQL() & "(>" & DisplayNum(SigFigOrDec(var5 / numDF, LSigFig, False), LSigFig, False) & ")"
                                                                Else
                                                                    strAQL = AQL() & "(>" & Format(RoundToDecimalRAFZ(var5 / numDF, LSigFig), GetRegrDecStr(LSigFig)) & ")"
                                                                End If
                                                            Else
                                                                strAQL = str1 & strR1 & AQL() & "(>NA)"
                                                                boolNA = True
                                                            End If

                                                        End If

                                                    End If

                                                    var3 = strAQL

                                                End If

                                                '******


                                            Else
                                                If Len(var4) = 0 Then
                                                    varE = "NA"
                                                    boolNA = True
                                                Else
                                                    varE = var4
                                                    boolConcStatus = True
                                                End If
                                                If StrComp(varE.ToString, "NA", CompareMethod.Text) = 0 Then
                                                    boolNA = True
                                                End If
                                            End If





                                        Case "ISR Run ID", "Watson Run ID"
                                            varE = NZ(rowsISR(Count4).Item("RUNID"), "NA")
                                            If StrComp(varE.ToString, "NA", CompareMethod.Text) = 0 Then
                                                boolNA = True
                                            End If
                                        Case "Mean Original+ISR"

                                            numMean = -1
                                            If numOrig = -1 Or numISR = -1 Then
                                                varE = "NA"
                                                boolNA = True
                                            Else
                                                var1 = (numOrig + numISR) / 2
                                                numMean = SigFigOrDec(var1, LSigFig, False) 'this function evaluates appropriate rounding
                                                varE = SigFigOrDecString(numMean, LSigFig, False) 'this function evaluates appropriate rounding
                                                boolDoMean = True
                                            End If

                                        Case "%Difference"

                                            '(ISR Result – Original Result)*100/(Mean)
                                            numPD = -1
                                            'If numISR = -1 Or numOrig = -1 Or numMean <= 0 Then
                                            If numISR = -1 Or numOrig = -1 Then
                                                varE = "NA"
                                                boolNA = True
                                            Else
                                                'var1 = ((numISR - numOrig) * 100) / numMean
                                                var1 = (numISR - numOrig)
                                                If boolDoMean Then
                                                    varE = Format(CalcCVPercent(var1, numMean, intQCDec), strQCDec)
                                                Else
                                                    'calculate mean
                                                    'Since Mean column is not shown, this value should be full precision
                                                    var2 = (numOrig + numISR) / 2
                                                    numMean = var2 ' SigFigOrDec(var2, LSigFig, False) 'this function evaluates appropriate rounding
                                                    varE = Format(CalcCVPercent(var1, numMean, intQCDec), strQCDec)
                                                End If

                                                numPD = CDec(varE)

                                            End If

                                        Case "Pass/Fail"
                                            If numPD = -1 Then
                                                If boolConcStatus Then
                                                    varE = "NA"
                                                    boolNA = True
                                                Else
                                                    varE = "NA"
                                                    boolNA = True
                                                End If
                                                If StrComp(varE.ToString, "NA", CompareMethod.Text) = 0 Then
                                                    boolNA = True
                                                End If
                                            Else
                                                If Math.Abs(numPD) > numCrit Then
                                                    varE = "Fail"
                                                Else
                                                    varE = "Pass"
                                                End If
                                            End If
                                        Case "Group"
                                            varE = rowsISR(Count4).Item("SUBJECTGROUPNAME")
                                        Case "Gender"
                                            varE = NZ(rowsISR(Count4).Item("GENDER"), "NA")
                                            If StrComp(varE, "NA", CompareMethod.Text) = 0 Then
                                                boolNA = True
                                            End If
                                        Case "Visit Text"
                                            varE = NZ(rowsISR(Count4).Item("VISITTEXT"), "NA")
                                            If StrComp(varE, "NA", CompareMethod.Text) = 0 Then
                                                boolNA = True
                                            End If
                                        Case "Time Text"
                                            varE = NZ(rowsISR(Count4).Item("TIMETEXT"), "NA")
                                            If StrComp(varE, "NA", CompareMethod.Text) = 0 Then
                                                boolNA = True
                                            End If
                                    End Select

                                    If Count3 = 1 Then
                                        strPasteT = varE
                                    Else
                                        strPasteT = strPasteT & ChrW(9) & varE
                                    End If

                                Next Count3

                                'count pass/fail for stats, pass/fail column may not be shown
                                Try
                                    If numPD = -1 Then
                                        If boolConcStatus Then

                                        Else

                                        End If
                                    Else
                                        If Math.Abs(numPD) > numCrit Then
                                            varE = "Fail"
                                        Else
                                            varE = "Pass"
                                            ctISRPass = ctISRPass + 1
                                        End If
                                        ctISRTot = ctISRTot + 1
                                    End If
                                Catch ex As Exception

                                End Try

                                ''console.writeline(strPasteT)

                                'If Count2 = 0 And Count4 = 0 And Count7 = 0 Then
                                If Count4 = 0 Then
                                    strPaste = strPasteT
                                Else
                                    strPaste = strPaste & ChrW(10) & strPasteT
                                End If
                                ctRows = ctRows + 1

                                str1 = strL & ChrW(10) & "Processing ISR " & Count4 & " of " & intRowsISR & "..."
                                frmH.lblProgress.Text = str1
                                frmH.lblProgress.Refresh()

                            Next Count4

                            str1 = strL & ChrW(10) & "Completed processing ISR " & intRowsISR & " of " & intRowsISR & "..." & ChrW(10) & "Entering data..."
                            frmH.lblProgress.Text = str1
                            frmH.lblProgress.Refresh()

                        Catch ex As Exception
                            var1 = ex.Message
                        End Try

                        '*****

                        Try
                            numISRPercent = RoundToDecimalRAFZ(ctISRPass / ctISRTot * 100, intQCDec)
                        Catch ex As Exception
                            numISRPercent = 0
                        End Try

                        .Selection.Tables.Item(1).Cell(3, 1).Select()

                        If IsNothing(strPaste) Then
                        Else

                            Dim rng1 As Word.Range
                            Dim tblW As Word.Table

                            tblW = .Selection.Tables.Item(1)
                            Try
                                rng1 = wd.ActiveDocument.Range(Start:=tblW.Cell(3, 1).Range.Start, End:=tblW.Cell(intRowsISR + 3 - 1, ctCols).Range.End)
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
                            'select appropriate rows
                            '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                            '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdColumn, Extend:=True)

                            Try
                                rng1 = wd.ActiveDocument.Range(Start:=tblW.Cell(3, 1).Range.Start, End:=tblW.Cell(intRowsISR + 3 - 1, ctCols).Range.End)
                                rng1.Select()
                            Catch ex As Exception
                                .Selection.SelectRow()
                                .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=ctRows - 1, Extend:=True)
                                var1 = ex.Message
                                var1 = var1
                            End Try

                            Pause(0.1)

                            'paste from clipboard
                            ''''console.writeline("Start Paste")
                            Try
                                .Selection.PasteAndFormat(Microsoft.Office.Interop.Word.WdRecoveryType.wdPasteDefault)
                            Catch ex As Exception
                                'MsgBox("Paste: " & ex.Message)
                            End Try
                            ''console.writeline("Skip2")
                            ''console.writeline(strPaste)
                            ''''console.writeline("End Paste")
                            ''' 


                            'the paste action removes the range object and any table formatting, must reset it
                            Call GlobalTableParaFormat(wd)

                            rng1 = wd.ActiveDocument.Range(Start:=tblW.Cell(3, 1).Range.Start, End:=tblW.Cell(tblW.Rows.Count, tblW.Columns.Count).Range.End)
                            rng1.Select()
                            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdColumn, Extend:=True)
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
                    'var1 = dvDo(intDo).Item("CHARHEADINGTEXT") 'Then change it back
                    'strTName = NZ(var1, "[NONE]")

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


                    str1 = frmH.lblProgress.Text

                    wtbl1.Select()

                    Dim intR As Integer
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

                    Dim arrLegend1(4, 100)
                    '1= Actual string to search in table
                    '2= Definition of string
                    '3= Not used
                    '4= True: Do not look for item in table, but add buffer row to row count.  False: Look for item in table; if found, add buffer row to row count
                    Dim ctLegend1 As Short = 0

                    If boolNA Then
                        ctLegend1 = ctLegend1 + 1
                        arrLegend1(1, ctLegend1) = "NA"
                        arrLegend1(2, ctLegend1) = "Not Applicable"
                        arrLegend1(3, ctLegend1) = False
                        arrLegend1(4, ctLegend1) = False
                    End If

                    ctLegend1 = ctLegend1 + 1

                    arrLegend1(1, ctLegend1) = BQL()
                    If boolBQLLEGEND Then
                        If boolLUseSigFigs Then
                            arrLegend1(2, ctLegend1) = BQLVerbose() & " (" & DisplayNum(SigFigOrDec(numLLOQ, LSigFig, False), LSigFig, False) & " " & strConcUnits & ")"
                        Else
                            arrLegend1(2, ctLegend1) = BQLVerbose() & " (" & Format(SigFigOrDec(numLLOQ, LSigFig, False), GetRegrDecStr(LSigFig)) & " " & strConcUnits & ")"
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
                            arrLegend1(2, ctLegend1) = AQLVerbose() & " (" & DisplayNum(SigFigOrDec(numULOQ, LSigFig, False), LSigFig, False) & " " & strConcUnits & ")"
                        Else
                            arrLegend1(2, ctLegend1) = AQLVerbose() & " (" & Format(SigFigOrDec(numULOQ, LSigFig, False), GetRegrDecStr(LSigFig)) & " " & strConcUnits & ")"
                        End If
                    Else
                        arrLegend1(2, ctLegend1) = AQLVerbose()
                    End If

                    arrLegend1(3, ctLegend1) = False
                    arrLegend1(4, ctLegend1) = False

                    ReDim Preserve arrLegend1(4, ctLegend1)


                    'autofit table
                    'Call AutoFitTable(wd, False)
                    '20180219 LEE: must show doc to autofit
                    Call AutoFitTable(wd, True)

                    strM = "Finalizing " & strTName & "..."
                    strM1 = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    str1 = strM1

                    frmH.lblProgress.Text = strM1
                    frmH.Refresh()

                    '20160107 Larry: Call SplitTable, but don't add legend. Wait till end since legend can be very large.
                    '20160107 Larry: Not true. Re-assay reasons should be small
                    'I think that is a holdover from Re-assay table
                    'boolOld = True
                    'If boolOld Then '
                    '    Call SplitTable(wd, 3, ctLegends, arrLegend, str1, True, int1 + 1, False, False, False, intTableID)
                    'Else
                    '    Call SplitTable(wd, 3, 0, arrLegend, str1, False, int1 + 1, False, False, False, intTableID)
                    'End If
                    Call SplitTable(wd, 3, ctLegend1, arrLegend1, str1, False, int1 + 1, False, False, False, intTableID)

                    'autofit table
                    'Call AutoFitTable(wd, False)
                    '20180219 LEE: must show doc to autofit
                    Call AutoFitTable(wd, True)

                    ''Sub SplitTable(ByVal wd As Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, 
                    ''ByVal arr As Object, ByVal strT As String, ByVal DoLegend As Boolean, ByVal intSplitRows As Short, ByVal boolSmallFont As Boolean)

                    'record posrow2
                    posrow2 = .Selection.Tables.Item(1).Rows.Count
                    int1 = .Selection.Tables.Item(1).Columns.Count


                    'Here!!!
                    ''''''''wdd.visible = True

                    Call MoveOneCellDown(wd)

                    '20190108 LEE:
                    'InsertLegend evaluates boolNoneLeg
                    'If boolNONELEG Then
                    'Else
                    '    'analyteid
                    '    Call InsertLegend(wd, intTableID, idTR, False, AnalyteID)
                    'End If
                    'analyteid
                    Call InsertLegend(wd, intTableID, idTR, False, AnalyteID)


next1:

                    If boolJustTable Then

                        If gNumMatrix = 1 Then
                            strA = strAnalC
                        Else
                            strA = strAnal 'strAnalC has '..Matrix', don't want to pass that here
                        End If

                        'No, just strAnal
                        ' strA = strAnal
                        strA = strAnalyteDescription
                        str1 = strA ' NZ(rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION"), "")
                        'Call JustTable(wd, str1, str2, strDo, strTName, intTableID)
                        If Len(str1) = 0 Then
                        Else
                            strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, False)
                            Call JustTable(wd, str1, strTName, strA, strTName, intTableID, strTempInfo, "", strTNameO, intGroup, idTR)
                        End If

                    End If

next2:

                Next Count2A
            Next Count1A
        End With

    End Sub


End Module
