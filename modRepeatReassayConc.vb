
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.ComponentModel.PropertyDescriptorCollection
Imports Word = Microsoft.Office.Interop.Word
Imports Microsoft.VisualBasic
Imports System.IO


Module modRepeatReassayConc


    Sub SRSummaryOfSC_5(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal idTR As Int64)

        'Summary of Sample Concentrations table

        Dim numNomConc As Decimal
        Dim BACStudy As String
        Dim rs As New ADODB.Recordset
        Dim constr As String
        Dim dbPath As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        'From modGroups: tblAnalyteGroups: "ANALYTEDESCRIPTION", "ANALYTEID", "INTSTD", "INTGROUP", "ANALYTEDESCRIPTION_C", "MATRIX", "INTCALSET", "CALIBRSET"
        Dim Count2A As Int32
        Dim Count2 As Int32
        Dim Count3 As Int32
        Dim Count4 As Int32
        Dim Count5 As Int32
        Dim Count6 As Int32
        Dim Count7 As Int32
        Dim Count1A As Int32
        Dim var1, var2, var3, var4, var5, var6, var7, var8, var9
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim int4 As Short
        Dim int5 As Short
        Dim arrTemp(2, 50)
        Dim num1 As Object
        Dim num2 As Object
        Dim num3 As Object
        Dim arrBCStdActual()
        Dim strA As String
        Dim strB As String
        Dim ctLegend As Short
        Dim intLeg As Short
        Dim lng1 As Long
        Dim lng2 As Long
        Dim boolPortrait As Boolean
        Dim intLastAnal As Short
        Dim arrOrder(7, 10)
        Dim ctCols 'number of columns in a table
        Dim strSub1 As String
        Dim strSub2 As String
        Dim pos1 As Short
        Dim pos2 As Short
        Dim numSum As Object
        Dim numMean As Object
        Dim numSD As Object
        Dim arrSampleConcs()
        '1=ANALYTEINDEX, 2=RUND, 3=CONCENTRATION, 4=DESIGNSAMPLEID, 5=ALIQUOTFACTOR, 6=CONCENTRATIONSTATUS, 7=REPLICATE
        Dim arrSampleDesign(1, 1)
        '1=DESIGNSUBJECTTAG, 2=SUBJECTGROUPNAME, 3=ENDDAY, 4=ENDHOUR, 5=DESIGNSAMPLEID
        Dim ctSampleConcs As Long
        Dim ctSampleDesign As Long
        Dim rsReassay As New ADODB.Recordset
        Dim rsDesign As New ADODB.Recordset
        Dim numBQL As Decimal ' Object
        Dim strBQL As String
        Dim numAQL As Decimal ' Object
        Dim strAQL As String

        Dim strP1 As String
        Dim strP2 As String
        Dim strP3 As String
        Dim intCTSDMax As Short = 50
        Dim intCTSD As Short = 0

        Dim boolNA As Boolean
        Dim boolNR As Boolean

        Dim boolNum As Boolean = True

        Dim dvDo As System.Data.DataView
        Dim intDo As Short
        Dim bool As Boolean
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim dv As System.Data.DataView
        Dim dv1 As System.Data.DataView
        Dim drows() As DataRow
        Dim drowsF() As DataRow
        Dim arrBCQCs(1, 1) '1=LevelNumber, 2=Concentration
        Dim arrBCQCConcs(1, 1)
        Dim int10 As Short
        Dim intF As Short
        Dim Count2A0 As Short
        Dim ctP As Short
        Dim boolGo As Boolean
        Dim ctQCs As Short
        Dim ctAnalyticalRuns As Short
        Dim Count20 As Short
        Dim tbl1 As System.Data.DataTable
        Dim dr1() As DataRow
        Dim strF As String
        Dim strS As String
        Dim strTName As String
        Dim strTempInfo As String
        Dim strConcUnits As String
        Dim strDoseUnits As String

        Dim intAnalyteID As Int64
        Dim intDesignSampleID As Int64

        Dim arrSS(2, 1) 'superscripted row
        '1=Row, 2=Col
        Dim intSS As Int32 = 0
        Dim lbSS As Short = 3


        Dim fonts
        Dim fontsize

        Dim strM As String
        Dim strM1 As String

        Dim tbl1A As DataTable
        Dim tbl2A As DataTable
        Dim strR As String = "_xyz_"
        Dim strTNameO As String 'original Table Name

        Dim charFCID As String
        strF = "ID_TBLREPORTTABLE = " & idTR
        Dim rowsTR() As DataRow = tblReportTable.Select(strF)
        var1 = rowsTR(0).Item("CHARFCID")
        charFCID = NZ(var1, "NA")

        Dim tblAG As DataTable = tblAnalyteGroups

        Dim doc As Word.Document

        With wd

            doc = wd.ActiveDocument
            Call SpellingOff(doc, False)

            fontsize = wd.ActiveDocument.Styles("Normal").Font.Size ' .Selection.Font.Size
            fonts = fontsize ' .Selection.Font.Size

            'dvDo = frmH.dgvReportTableConfiguration.DataSource
            'strTName = "Summary of Samples"
            'intDo = FindRowDVByCol(strTName, dvDo, "Table")

            Dim intTableID As Short
            intTableID = 5

            Dim strWRunId As String = GetWatsonColH(intTableID)

            dvDo = frmH.dgvReportTableConfiguration.DataSource
            'intDo = FindRowDVNumByCol(intTableID, dvDo, "id_tblconfigreporttables")

            ''Get table name
            'var1 = dvDo(intDo).Item("Table")
            'strTName = NZ(var1, "[NONE]")

            '***
            Dim intTC As Int64
            Try
                intDo = FindRowDVNumByCol(idTR, dvDo, "ID_TBLREPORTTABLE")
            Catch ex As Exception
                MsgBox("intDo Prob")
            End Try

            Dim strSString As String
            Dim strGroupCheck As String
            Call GetGroupSort(idTR) 'retrieve grouping and sorting information

            strSString = GetSString()

            If intGroups = 0 Then
                strGroupCheck = "[None]"
            Else
                strGroupCheck = arrGroups(1, 1)
            End If

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


            Dim intGroup As Int16

            Dim strErr As String
            Dim vDay, vTime, vSubject

            Dim strAD As String = ""

            For Count1A = 0 To 0 '20180216 LEE: ignore this loop 'tbl1A.Rows.Count - 1 'Iterate through each Matrix (but keep different calibration ranges together)

                'strTName = strTNameO 'reset strTName

                intGroup = -2

                'If boolM Then
                '    strMatrix = tblMatrices.Rows(Count1A).Item("Matrix")
                'Else
                '    strAnalyteID = tblAnalyteIDs.Rows(Count1A).Item("AnalyteID")
                '    intAnalyteID = tblAnalyteIDs.Rows(Count1A).Item("AnalyteID")
                '    strAnalyteDescription = tblAnalyteIDs.Rows(Count1A).Item("AnalyteDescription")
                'End If

                For Count2A = 0 To intRowsAnal - 1 '20180216 LEE: tbl2A.Rows.Count - 1 'Iterate through each AnalyteID, and generate the information

                    '20171128 LEE:
                    strTName = strTNameO 'reset strTName

                    Dim arrLegend(4, 1000)

                    Dim boolDate As Boolean = False

                    'If boolM Then
                    '    strAnalyteID = tblAnalyteIDs.Rows(Count2A).Item("AnalyteID")
                    '    intAnalyteID = tblAnalyteIDs.Rows(Count2A).Item("AnalyteID")
                    '    strAnalyteDescription = tblAnalyteIDs.Rows(Count2A).Item("AnalyteDescription")
                    'Else
                    '    strMatrix = tblMatrices.Rows(Count2A).Item("Matrix")
                    'End If

                    ''find intGroup
                    'strF = "ANALYTEID = " & intAnalyteID & " AND MATRIX = '" & strMatrix & "'"
                    'Dim rowsAG() As DataRow = tblAG.Select(strF)
                    'If rowsAG.Length = 0 Then
                    'Else
                    '    intGroup = rowsAG(0).Item("INTGROUP")
                    'End If



                    '20180216 LEE: 
                    Try
                        strAnalyteID = rows11(Count2A).Item("AnalyteID")
                        intAnalyteID = rows11(Count2A).Item("AnalyteID")
                        strAnalyteDescription = rows11(Count2A).Item("OriginalAnalyteDescription")
                        strMatrix = rows11(Count2A).Item("Matrix")
                        strConcUnits = rows11(Count2A).Item("ConcUnits")
                        '20180629 LEE:
                        'BIG note: since we're reporting all concentration levels in this report
                        'there may be more than one INTGROUP
                        'need to evaluate group for the analytical run for each sample
                        'intGroup = GetGroup(intAnalyteID, strMatrix) ' rows11(Count2A).Item("INTGROUP") 'need intGroup for overall BQL/AQL, rows11 does not have intGroup

                    Catch ex As Exception
                        var1 = var1 'debug
                    End Try

                    If (Not (boolGenerateTableForThisAnalyteIDandMatrix(intDo, strAnalyteID, strMatrix))) Then
                        GoTo next1
                    End If

                    frmH.lblProgress.Text = "Entering Summary of " & strAnalyteDescription & " in " & strMatrix & " Concentrations Table..."
                    frmH.lblProgress.Refresh()

                    ctQCs = arrctQCs(3, Count2A)

                    'Start Table
                    intTCur = intTCur + 1

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

                    'determine number of columns and order of columns
                    'find column number
                    '1=ColumnHeader, 2=Include(X), 3=Order, 4=ReportColumnHeader

                    tbl1 = tblReportTableHeaderConfig
                    strF = "id_tblStudies = " & id_tblStudies & " AND id_tblConfigReportTables = 5 AND boolInclude = -1"
                    dr1 = tbl1.Select(strF, "intOrder ASC")
                    int1 = dr1.Length
                    ctCols = int1

                    Dim boolHasDate As Boolean
                    Dim idTCHL As Int16

                    '20160220 LEE: determine if there are more than one unique weeks
                    '20180627 LEE: Don't need to do anymore. Added Weeks as a user column choice
                    'strF = "WEEK IS NOT NULL"
                    'Dim dvWeeks As DataView = New DataView(tblSampleDesign, strF, "", DataViewRowState.CurrentRows)
                    'Dim tblWeeks As DataTable = dvWeeks.ToTable("a", True, "WEEK")
                    'Dim boolDoWeeks As Boolean = False
                    'If tblWeeks.Rows.Count > 1 Then
                    '    boolDoWeeks = True
                    'End If

                    ctLegend = 0
                    intLeg = 0

                    Count2 = 0
                    '1=ColumnHeader, 2=Include(X), 3=Order, 4=ReportColumnHeader
                    'arrOrder
                    ' 1=ColumnLabel, 2=Order, 3=id_tblConfigReportTables, 4=UserLabel, 5=id_tblConfigHeaderLookup, '6=CHARWATSONFIELD
                    Dim tbl As System.Data.DataTable
                    Dim dr() As DataRow
                    tbl = tblConfigHeaderLookup

                    int1 = 0
                    For Count2 = 1 To ctCols

                        int1 = int1 + 1
                        If int1 > UBound(arrOrder, 2) Then
                            ReDim Preserve arrOrder(7, UBound(arrOrder, 2) + 10)
                        End If

                        arrOrder(4, int1) = NZ(dr1(Count2 - 1).Item("charUserLabel"), "")
                        arrOrder(2, int1) = Count2 'dr1(Count2 - 1).Item("intOrder")
                        arrOrder(3, int1) = dr1(Count2 - 1).Item("id_tblConfigHeaderLookup")
                        idTCHL = arrOrder(3, int1)
                        If idTCHL = 214 Or idTCHL = 219 Then
                            boolHasDate = True
                        End If
                        'find column label
                        str1 = "id_tblConfigHeaderLookup = " & arrOrder(3, int1)
                        dr = tbl.Select(str1)
                        arrOrder(1, int1) = dr(0).Item("charColumnLabel")
                        arrOrder(5, int1) = dr1(Count2 - 1).Item("id_tblConfigReportTables")
                        'find CHARWATSONFIELD
                        Erase dr
                        str1 = "id_tblConfigReportTables = " & dr1(Count2 - 1).Item("id_tblConfigReportTables")
                        str1 = str1 & " AND id_tblConfigHeaderLookup = " & dr1(Count2 - 1).Item("id_tblConfigHeaderLookup")
                        dr = tbl.Select(str1)
                        arrOrder(6, int1) = dr(0).Item("CHARWATSONFIELD")

                        var1 = arrOrder(6, int1)
                        var2 = arrOrder(1, int1) 'DEBUG

                        '20180627 LEE: Don't need to do anymore. Added Weeks as a user column choice
                        'If StrComp(var1.ToString, "ENDDAY", CompareMethod.Text) = 0 And boolDoWeeks Then 'add another column

                        '    var2 = dr1(Count2 - 1).Item("id_tblConfigReportTables")
                        '    int1 = int1 + 1
                        '    If int1 > UBound(arrOrder, 2) Then
                        '        ReDim Preserve arrOrder(7, UBound(arrOrder, 2) + 10)
                        '    End If

                        '    'first increment day up
                        '    For Count3 = 1 To 6
                        '        arrOrder(Count3, int1) = arrOrder(Count3, int1 - 1)
                        '    Next
                        '    arrOrder(2, int1) = int1 'set the order + 1

                        '    'arrOrder(2, int1) = int1  'dr1(int1 - 1).Item("intOrder")
                        '    'arrOrder(3, int1) = -1 ' dr1(int1 - 1).Item("id_tblConfigHeaderLookup")
                        '    'arrOrder(4, int1) = "Week" ' dr1(int1 - 1).Item("charUserLabel")
                        '    'arrOrder(5, int1) = var2
                        '    'arrOrder(1, int1) = "Week" ' dr(0).Item("CHARCOLUMNLABEL")
                        '    'arrOrder(6, int1) = "WEEK" ' dr(0).Item("CHARWATSONFIELD")

                        '    arrOrder(2, int1 - 1) = int1 - 1 'dr1(int1 - 1).Item("intOrder")
                        '    arrOrder(3, int1 - 1) = -1 ' dr1(int1 - 1).Item("id_tblConfigHeaderLookup")
                        '    arrOrder(4, int1 - 1) = "Week" ' dr1(int1 - 1).Item("charUserLabel")
                        '    arrOrder(5, int1 - 1) = var2
                        '    arrOrder(1, int1 - 1) = "Week" ' dr(0).Item("CHARCOLUMNLABEL")
                        '    arrOrder(6, int1 - 1) = "WEEK" ' dr(0).Item("CHARWATSONFIELD")

                        '    ' 1=ColumnLabel, 2=Order, 3=id_tblConfigReportTables, 4=UserLabel, 5=id_tblConfigHeaderLookup, '6=CHARWATSONFIELD

                        'End If
                    Next

                    ctCols = int1
                    'Dim arrOrder(7, 100)
                    ' 1=ColumnLabel, 2=Order, 3=id_tblConfigReportTables, 4=UserLabel, 5=id_tblConfigHeaderLookup, 6=WatsonTitle
                    '6=CHARWATSONFIELD

                    ReDim Preserve arrOrder(7, ctCols)

                    '20160311 LEE:
                    'get units from tblanalyteshome
                    Dim strF1 As String
                    'strF1 = "ANALYTEID = " & intAnalyteID & " AND MATRIX = '" & strMatrix & "'"
                    'Dim rowsUnits() As DataRow = tblAnalytesHome.Select(strF1)
                    'strConcUnits = rowsUnits(0).Item("ConcUnits")

                    'dv = frmH.dgvWatsonAnalRef.DataSource
                    'int2 = FindRowDV("LLOQ Units", dv)
                    'int4 = FindRowDV("ULOQ Units", dv)

                    'var1 = dv.Item(int2).Item(Count2A) 'Sheets("AnalRefTables").Range("ULOQUnits").Offset(0, Count2A).Value
                    'strConcUnits = var1

                    'Dim int111
                    'int111 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
                    'str1 = NZ(frmH.dgvStudyConfig(1, int111).Value, "")
                    'If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
                    'Else
                    '    strConcUnits = str1
                    'End If

                    strF = makeRunMatrixAndSelectedAnalyteFilter(intDo, strAnalyteID, strMatrix)

                    '********END******

                    'str4 = "DESIGNSUBJECTTAG ASC"
                    'prepare sort string
                    'str4 = " DESIGNSUBJECTTAG ASC, SUBJECTGROUPNAME ASC, TREATMENTID ASC, GENDERID ASC, ENDDAY ASC, ENDHOUR ASC, ENDMINUTE ASC, ENDSECOND ASC"

                    If Len(strSString) = 0 Then
                        'str4 = " DESIGNSUBJECTTAG ASC, ENDDAY ASC, ENDHOUR ASC, ENDMINUTE ASC, ENDSECOND ASC"
                        'If InStr(1, strSString, "Start", CompareMethod.Text) > 0 Then
                        '    str4 = "DESIGNSUBJECTTAG ASC, SERIALSTARTTIME ASC"
                        'Else
                        '    str4 = "DESIGNSUBJECTTAG ASC, SERIALENDTIME ASC"
                        'End If
                        str4 = "DESIGNSUBJECTTAG ASC, SERIALENDTIME ASC, SERIALSTARTTIME ASC"
                    Else
                        'str4 = " " & strSString & ", ENDDAY ASC, ENDHOUR ASC, ENDMINUTE ASC, ENDSECOND ASC"
                        If InStr(1, strSString, "Start", CompareMethod.Text) > 0 Then
                            str4 = " " & strSString & ", SERIALENDTIME ASC"
                        Else
                            str4 = " " & strSString & ", SERIALSTARTTIME ASC"
                        End If
                    End If

                    Erase drows
                    ''console.writeline(strF)
                    '20160304 LEE: using above strF, null values (Median conc, Mean conc) get excluded
                    'actaully, we don't need the filter to include accepted analytical runs
                    'the underlying Watson table (SAMPLERESULTS) values by definition come from accepted analytical runs
                    'we should only need to filter by analyteid and matrix type

                    str1 = "ANALYTEID = " & intAnalyteID & " AND SAMPLETYPEID = '" & strMatrix & "'"

                    ''debug
                    'Console.WriteLine("str1: " & str1)
                    'Console.WriteLine("str4: " & str4)
                    'Console.WriteLine("Start")
                    'For Count2 = 0 To tblSampleDesign.Columns.Count - 1
                    '    Console.WriteLine(tblSampleDesign.Columns(Count2).ColumnName)
                    'Next
                    'Console.WriteLine("End")

                    drows = tblSampleDesign.Select(str1, str4)
                    int1 = drows.Length

                    ''debug
                    'Try
                    '    Dim CountAA As Int16
                    '    Dim CountB As Int16
                    '    Console.WriteLine("START")
                    '    var1 = ""
                    '    For CountAA = 0 To tblSampleDesign.Columns.Count - 1
                    '        var2 = tblSampleDesign.Columns(CountAA).ColumnName
                    '        var1 = var1 & ";" & var2
                    '    Next
                    '    Console.WriteLine(var1)
                    '    For CountB = 0 To drows.Length - 1
                    '        var1 = ""
                    '        For CountAA = 0 To tblSampleDesign.Columns.Count - 1
                    '            var2 = NZ(drows(CountB).Item(CountAA), "GUBBS")
                    '            var1 = var1 & ";" & var2
                    '        Next
                    '        Console.WriteLine(var1)
                    '    Next
                    '    Console.WriteLine("END")
                    'Catch ex As Exception
                    '    var3 = ex.Message
                    'End Try


                    ctSampleDesign = int1
                    Count2 = 0

                    'find strDoseUnits
                    strDoseUnits = "NA"
                    Try
                        For Count2 = 0 To drows.Length - 1
                            var1 = NZ(drows(Count2).Item("DOSEUNITSDESCRIPTION"), "")
                            If Len(var1) = 0 Then
                            Else
                                strDoseUnits = var1
                                Exit For
                            End If
                        Next
                    Catch ex As Exception
                        var1 = ex.Message
                    End Try


                    'get runid's

                    Dim dvRunIds As System.Data.DataView = New DataView(tblSampleDesign, strF, str4, DataViewRowState.CurrentRows)
                    Dim tblRunIds As System.Data.DataTable = dvRunIds.ToTable("A", True, "RUNID")
                    Dim intRID As Int16 = tblRunIds.Rows.Count
                    Dim intE1 As Short
                    Dim CountA As Short
                    Dim intColRunDate As Short = 0

                    If intRID = 0 Then
                        varR = "NA"
                    Else
                        varR = "NA"
                        If intRID < 3 Then
                            For CountA = 0 To intRID - 1
                                If CountA = 0 Then
                                    varR = tblRunIds.Rows(CountA).Item("RUNID")
                                Else
                                    var1 = tblRunIds.Rows(CountA).Item("RUNID")
                                    varR = varR & " and " & var1
                                End If
                            Next
                        Else
                            For CountA = 0 To intRID - 1
                                If CountA = 0 Then
                                    varR = tblRunIds.Rows(CountA).Item("RUNID")
                                ElseIf CountA = tblRunIds.Rows.Count - 1 Then
                                    var1 = tblRunIds.Rows(CountA).Item("RUNID")
                                    varR = varR & ", and " & var1
                                Else
                                    var1 = tblRunIds.Rows(CountA).Item("RUNID")
                                    varR = varR & ", " & var1
                                End If
                            Next
                        End If
                    End If
                    ctrsSamples(2, Count2A) = varR

                    tblRunIds.Dispose()
                    dvRunIds.Dispose()

                    'populate arrOrder for column order and header info

                    For Count3 = 1 To ctCols
                        'find the first entry in arrorder
                        int1 = Count3
                        Count4 = 0

                        'arrOrder
                        ' 1=ColumnLabel, 2=Order, 3=id_tblConfigReportTables, 4=UserLabel, 5=id_tblConfigHeaderLookup
                        '6=CHARWATSONFIELD
                        str1 = arrOrder(1, int1)



                        Select Case str1

                            Case "Concentration"
                                str2 = arrOrder(4, Count3) 'user label 
                                arrOrder(4, Count3) = str2 & ChrW(10) & "(" & strConcUnits & ")"

                            Case "Dose Amount"
                                str2 = arrOrder(4, Count3) 'user label
                                arrOrder(4, Count3) = str2 & ChrW(10) & "(" & strDoseUnits & ")"
                                'ctSampleDesign
                        End Select

                    Next

                    Dim ubASD As Short = 6
                    Erase arrSampleDesign
                    ReDim arrSampleDesign(ctCols + ubASD, ctSampleDesign)
                    '1=DESIGNSUBJECTTAG, 2=SUBJECTGROUPNAME, 3=ENDDAY, 4=ENDHOUR, 5=CONCENTRATION, 6=ALIQUOTFACTOR, 7=ID, 8=DESIGNSAMPLEID, 9=GENDER
                    'Order Columns + 6
                    'ctcols+1=ALIQUTOFACTOR, +2=RUNID, +3=DESIGNSAMPLEID, +4=COMMENTMEMO

                    strP1 = "Entering Summary of " & strAnalyteDescription & " in " & strMatrix & " Concentrations table..."
                    strP1 = strP1 & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    strP2 = "Preparing Sample 1 of " & ctSampleDesign & " samples..."
                    strP3 = strP1 & ChrW(10) & strP2
                    frmH.lblProgress.Text = strP3
                    frmH.lblProgress.Refresh()


                    intCTSD = 0

                    For Count2 = 0 To ctSampleDesign - 1


                        intCTSD = intCTSD + 1
                        If intCTSD > intCTSDMax Then
                            strP2 = "Preparing Sample " & Count2 + 1 & " of " & ctSampleDesign & " samples..."
                            strP3 = strP1 & ChrW(10) & strP2
                            frmH.lblProgress.Text = strP3
                            frmH.lblProgress.Refresh()
                            intCTSD = 0
                        End If


                        Try
                            intColRunDate = 0
                            For Count3 = 1 To ctCols
                                'find the first entry in arrorder
                                int1 = Count3
                                Count4 = 0

                                'arrOrder
                                ' 1=ColumnLabel, 2=Order, 3=id_tblConfigReportTables, 4=UserLabel, 5=id_tblConfigHeaderLookup, 6=WatsonTitle
                                str1 = arrOrder(1, int1)
                                str2 = arrOrder(6, int1)

                                Select Case str1
                                    'Case "Group"
                                    '    str2 = "SUBJECTGROUPNAME"
                                    'Case "Subject"
                                    '    str2 = "DESIGNSUBJECTTAG"
                                    'Case "Gender"
                                    '    str2 = "GENDERID"
                                    'Case "Treatment"
                                    '    str2 = "TREATMENTID"
                                    'Case "Day"
                                    '    str2 = "ENDDAY"
                                    'Case "Time"
                                    '    str2 = "ENDHOUR"
                                    'Case "Concentration"
                                    '    str2 = "CONCENTRATION"
                                    'Case "Watson Run ID"
                                    '    str2 = "RUNID"
                                    Case "Analyte"
                                        str2 = "Analyte"

                                    Case "Sample Count"
                                        str2 = "Sample Count"
                                End Select

                                'evaluate
                                Dim varH, varM, varS
                                Dim varH1, varM1, varS1, varT

                                Dim varHS, varM1S, varS1S, varTS

                                varHS = ""
                                varM1S = ""
                                varS1S = ""
                                varTS = ""

                                Dim strStartDay As String = ""
                                Dim strStartHour As String = ""
                                Dim strStartMinute As String = ""
                                Dim strStartSecond As String = ""

                                '   Case "Sample Count"
                                'str1 = Count2.ToString

                                If StrComp(str2, "Analyte", CompareMethod.Text) = 0 Then
                                    var1 = strAnalyteDescription

                                ElseIf StrComp(str2, "Custom ID", CompareMethod.Text) = 0 Then
                                    var1 = var2.ToString

                                ElseIf StrComp(str2, "Sample Count", CompareMethod.Text) = 0 Then
                                    var2 = Count2 + 1
                                    var1 = var2.ToString

                                ElseIf StrComp(str2, "CONCENTRATION", CompareMethod.Text) = 0 Then
                                    'No! If concentration is null, then don't report
                                    var1 = NZ(drows(Count2).Item(str2), "")
                                    If Len(var1) = 0 Then 'report CONCENTRATIONSTATUS instead

                                        'from Watson Data Dictionary 7.3
                                        'CONCENTRATIONSTATUS  If = “NM” or “VEC”, deal with calibration ranges, if not empty, display it
                                        'CALIBRATIONRANGEFLAG  Possible values are NM, VEC, or null  (NM=bql, VEC=aql)
                                        'CALBRATIONRANGE  If out of calibration range, this value is the NM or VEC value

                                        var2 = NZ(drows(Count2).Item("CONCENTRATIONSTATUS"), "NR")
                                        var3 = NZ(drows(Count2).Item("CALIBRATIONRANGEFLAG"), "")
                                        var4 = NZ(drows(Count2).Item("CALIBRATIONRANGE"), 0)
                                        var5 = NZ(drows(Count2).Item("ALIQUOTFACTOR"), 1)

                                        If StrComp(var2, "NM", CompareMethod.Text) = 0 Then 'bql/aql
                                            If boolLUseSigFigs Then
                                                strBQL = BQL() & "(<" & DisplayNum(SigFigOrDec(CDec(var4 / var5), LSigFig, False), LSigFig, False) & ")"
                                            Else
                                                strBQL = BQL() & "(<" & Format(CDec(var4 / var5), GetRegrDecStr(LSigFig)) & ")"
                                            End If
                                            var1 = strBQL
                                        ElseIf StrComp(var2, "VEC", CompareMethod.Text) = 0 Then 'bql/aql
                                            If boolLUseSigFigs Then
                                                strAQL = AQL() & "(>" & DisplayNum(SigFigOrDec(CDec(var4 / var5), LSigFig, False), LSigFig, False) & ")"
                                            Else
                                                strAQL = AQL() & "(>" & Format(CDec(var4 / var5), GetRegrDecStr(LSigFig)) & ")"
                                            End If
                                            var1 = strAQL
                                        Else
                                            var1 = var2
                                        End If

                                        If StrComp(var2, "NR", CompareMethod.Text) = 0 Then
                                            boolNR = True

                                            'Add to Legend Array
                                            strA = "NR"
                                            strB = "Not Reported"
                                            ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, strB, strA, False)
                                            intLeg = ctLegend

                                        End If

                                        'evaluate for bql and aql later

                                        'var2 = NZ(drows(Count2).Item("CALIBRATIONRANGEFLAG"), "")
                                        'var3 = NZ(drows(Count2).Item("CALIBRATIONRANGE"), "")


                                    End If

                                ElseIf StrComp(str2, "WEEK", CompareMethod.Text) = 0 Then
                                    var1 = NZ(drows(Count2).Item("WEEK"), "")

                                ElseIf StrComp(str2, "ENDHOUR,ENDMINUTE,ENDSECOND", CompareMethod.Text) = 0 Then
                                    'var1 = NZ(drows(Count2).Item("ENDHOUR"), 0)
                                    var1 = NZ(drows(Count2).Item("ENDHOUR"), 0)
                                    varH = RoundToDecimal(NZ(drows(Count2).Item("ENDHOUR"), 0), 3)
                                    varM = NZ(drows(Count2).Item("ENDMINUTE"), 0)
                                    varM1 = RoundToDecimal(varM / 60, 3)
                                    varS = NZ(drows(Count2).Item("ENDSECOND"), 0)
                                    varS1 = RoundToDecimal(varS / 3600, 3)

                                    'check for both end and start
                                    strStartHour = NZ(drows(Count2).Item("STARTHOUR"), "")
                                    strStartMinute = NZ(drows(Count2).Item("STARTMINUTE"), "")
                                    strStartSecond = NZ(drows(Count2).Item("STARTMINUTE"), "")
                                    '
                                    If Len(strStartHour) <> 0 Or Len(strStartMinute) <> 0 Or Len(strStartSecond) <> 0 Then
                                        var1 = NZ(drows(Count2).Item("STARTHOUR"), 0)
                                        varHS = RoundToDecimal(NZ(drows(Count2).Item("STARTHOUR"), 0), 3)
                                        varM = NZ(drows(Count2).Item("STARTMINUTE"), 0)
                                        varM1S = RoundToDecimal(varM / 60, 3)
                                        varS = NZ(drows(Count2).Item("STARTSECOND"), 0)
                                        varS1S = RoundToDecimal(varS / 3600, 3)
                                    End If

                                ElseIf StrComp(str2, "STARTHOUR,STARTMINUTE,STARTSECOND", CompareMethod.Text) = 0 Then
                                    var1 = NZ(drows(Count2).Item("STARTHOUR"), 0)
                                    varH = RoundToDecimal(NZ(drows(Count2).Item("STARTHOUR"), 0), 3)
                                    varM = NZ(drows(Count2).Item("STARTMINUTE"), 0)
                                    varM1 = RoundToDecimal(varM / 60, 3)
                                    varS = NZ(drows(Count2).Item("STARTSECOND"), 0)
                                    varS1 = RoundToDecimal(varS / 3600, 3)

                                ElseIf StrComp(str2, "RUNDATE", CompareMethod.Text) = 0 Then

                                Else
                                    var1 = drows(Count2).Item(str2)
                                End If

                                'evaluate
                                '20160301 LEE: added
                                If StrComp(str2, "ASSAYDATETIME", CompareMethod.Text) = 0 Then
                                    ''sometimes ASSAYDATETIME is null
                                    If Len(NZ(var1, "")) = 0 Then
                                        var1 = "NA"
                                        'get from tblCalStdGroupAssayIDsAll
                                        var1 = drows(Count2).Item("RUNID")
                                        Dim rrr() As DataRow = tblCalStdGroupAssayIDsAll.Select("RUNID = " & var1)
                                        var1 = rrr(0).Item("RUNDATE")
                                        var1 = Format(var1, LDateFormat)
                                    Else
                                        var1 = Format(var1, LDateFormat)
                                    End If

                                    'even this sometimes has null (null RunID). See later code: search for intColRunDate
                                    intColRunDate = Count3

                                    boolDate = True

                                End If

                                If StrComp(str2, "RUNDATE", CompareMethod.Text) = 0 Then
                                    Try
                                        var1 = NZ(drows(Count2).Item("RUNID"), 0)
                                        Dim rrr() As DataRow = tblCalStdGroupAssayIDsAll.Select("RUNID = " & var1)
                                        var1 = rrr(0).Item("RUNDATE")
                                        var1 = Format(var1, LDateFormat)
                                        var1 = var1 'debug
                                    Catch ex As Exception
                                        var2 = ex.Message
                                        var1 = "NA"
                                    End Try

                                    'even this sometimes has null (null RunID). See later code: search for intColRunDate
                                    intColRunDate = Count3

                                    boolDate = True

                                End If

                                If StrComp(str1, "Time", vbTextCompare) = 0 Then

                                    varT = varH + varM1 + varS1
                                    If Len(strStartHour) <> 0 Or Len(strStartMinute) <> 0 Or Len(strStartSecond) <> 0 Then
                                        varTS = varHS + varM1S + varS1S
                                        varT = varTS & " to " & varT
                                        var1 = varT
                                    Else
                                        If varH = 0 And varM = 0 And varS = 0 Then
                                            var1 = 0
                                        Else
                                            var1 = varT
                                        End If
                                    End If

                                ElseIf StrComp(str1, "Start Time", vbTextCompare) = 0 Then

                                    If varH = 0 And varM = 0 And varS = 0 Then
                                        var1 = 0
                                    Else
                                        varT = varH + varM1 + varS1
                                        var1 = varT

                                    End If

                                ElseIf StrComp(str1, "Gender", CompareMethod.Text) = 0 Then
                                    Select Case NZ(var1, 0)
                                        Case 1
                                            var1 = "Male"
                                        Case 2
                                            var1 = "Female"
                                        Case 0
                                            var1 = "[None]"
                                    End Select

                                ElseIf StrComp(str1, "Dil Factor", CompareMethod.Text) = 0 Then

                                    var2 = 1 / NZ(var1, 1)
                                    If IsInt(var2) Then
                                        var1 = var2
                                    Else
                                        var3 = RoundToDecimalRAFZ(var2, 1)
                                        var1 = var3
                                    End If

                                ElseIf StrComp(str1, "Visit Text", CompareMethod.Text) = 0 Then
                                    var1 = drows(Count2).Item("VISITTEXT")

                                ElseIf StrComp(str1, "Time Text", CompareMethod.Text) = 0 Then
                                    var1 = drows(Count2).Item("TIMETEXT")

                                End If
                                'arrSampleDesign(Count3, Count2 + 1) = NZ(var1, 0)
                                arrSampleDesign(Count3, Count2 + 1) = var1

                            Next Count3

                        Catch ex As Exception
                            var4 = ex.Message
                            var4 = var4
                        End Try

                        'Order Columns + 6
                        'ctcols+1=ALIQUTOFACTOR, +2=RUNID, +3=DESIGNSAMPLEID, +4=COMMENTMEMO, +5=CalibrationRangeFlag, +6=CalibrationRange
                        arrSampleDesign(ctCols + 1, Count2 + 1) = drows(Count2).Item("ALIQUOTFACTOR")
                        var4 = NZ(drows(Count2).Item("COMMENTMEMO"), "")
                        arrSampleDesign(ctCols + 4, Count2 + 1) = var4
                        var4 = NZ(drows(Count2).Item("CALIBRATIONRANGEFLAG"), "")
                        arrSampleDesign(ctCols + 5, Count2 + 1) = var4
                        var4 = NZ(drows(Count2).Item("CALIBRATIONRANGE"), "")
                        arrSampleDesign(ctCols + 6, Count2 + 1) = var4

                        'here
                        'check for null RUNID
                        Dim intRunIDUse As Short = 0
                        var1 = NZ(drows(Count2).Item("RUNID"), "")
                        If Len(var1) = 0 Then

                            'tblGetDecRunID
                            intDesignSampleID = drows(Count2).Item("DESIGNSAMPLEID")
                            strF = "ANALYTEID = " & intAnalyteID & " AND DESIGNSAMPLEID = " & intDesignSampleID
                            Dim intDC As Short

                            'first find decision code from tblSAMPRESCONFLICTDEC
                            Dim rowsSDEC() As DataRow
                            Try
                                rowsSDEC = tblSAMPRESCONFLICTDEC.Select(strF)
                            Catch ex As Exception
                                var4 = ex.Message
                                var4 = var4
                            End Try

                            'tblSAMPRESCONFLICTDEC is sorted by RECORDTIMESTAMP DESC
                            'if multiple records are returned, the latest decision is used
                            If rowsSDEC.Length = 0 Then
                                var3 = "NA"
                            Else
                                intDC = rowsSDEC(0).Item("DECISIONCODE")
                                '1=Choose one, 2=Mean, 3=Median
                                '1: shouldn't occur - a RUNID should exist
                                '2: for now, lets list all runids
                                '3: find the actual runid
                                var3 = ""
                                If intDC = 0 Then
                                    'RUNID is null in tblSampleDesign
                                    'first look in rowsSDEC
                                    For Count3 = 0 To rowsSDEC.Length - 1
                                        var2 = NZ(rowsSDEC(0).Item("RUNID"), "")
                                        If Len(var2) = 0 Then
                                        Else
                                            var3 = var2
                                            Exit For
                                        End If
                                    Next
                                    If Len(var3) = 0 Then
                                        'take last runid in tblReassayReport
                                        Dim strFRR As String = "ANALYTEID = " & intAnalyteID & " AND DESIGNSAMPLEID = " & intDesignSampleID
                                        Dim rowsRR() As System.Data.DataRow = tblReassayReport.Select(strFRR, "RUNID DESC", DataViewRowState.CurrentRows)
                                        If rowsRR.Length = 0 Then
                                            var3 = "NA"
                                        Else
                                            var3 = rowsRR(0).Item("RUNID")
                                            intRunIDUse = var3
                                        End If
                                    End If

                                ElseIf intDC = 2 Then

                                    '20160313 LEE: Note: may get this from tblSAMPRESCONFLICTDEC in the future
                                    'if Concentration is null

                                    'get all runids from tblReassayReport
                                    Dim dvMean As DataView = New DataView(tblReassayReport, strF, "RUNID", DataViewRowState.CurrentRows)
                                    Dim tblMean As DataTable = dvMean.ToTable("a", True, "RUNID")
                                    For Count3 = 0 To tblMean.Rows.Count - 1
                                        var4 = tblMean.Rows(Count3).Item("RUNID")
                                        If Count3 = 0 Then
                                            var3 = var4
                                        Else
                                            var3 = var3 & "," & var4
                                        End If
                                    Next
                                    intRunIDUse = var4
                                ElseIf intDC = 3 Then
                                    'get runid from tblGetDecRunID which should return a single record if DECISIONCODE = 3
                                    Dim rowsMedian() As DataRow = tblGetDecRunID.Select(strF)
                                    If rowsMedian.Length = 0 Then
                                        var3 = "NA"
                                    Else
                                        var3 = NZ(rowsMedian(0).Item("RUNID"), "NA")
                                    End If
                                    If IsNumeric(var3) Then
                                        intRunIDUse = var3
                                    End If
                                End If
                            End If

                            'need to reset ASSAYDATETIME or RUNDATE
                            If IsNumeric(var3) Then
                                Try
                                    var1 = intRunIDUse
                                    Dim rrr() As DataRow = tblCalStdGroupAssayIDsAll.Select("RUNID = " & var1)
                                    var1 = rrr(0).Item("RUNDATE")
                                    var1 = Format(var1, LDateFormat)
                                    var1 = var1 'debug
                                    arrSampleDesign(intColRunDate, Count2 + 1) = var1
                                    'intColRunDate
                                Catch ex As Exception
                                    var2 = ex.Message
                                    arrSampleDesign(intColRunDate, Count2 + 1) = "NA"
                                End Try
                            End If

                        Else
                            var3 = var1
                        End If

                        arrSampleDesign(ctCols + 2, Count2 + 1) = var3 ' drows(Count2).Item("RUNID")

                        arrSampleDesign(ctCols + 3, Count2 + 1) = drows(Count2).Item("DESIGNSAMPLEID")


                    Next Count2

                    strP2 = "Finished preparing Sample " & ctSampleDesign & " of " & ctSampleDesign & " samples..."
                    strP3 = strP1 & ChrW(10) & strP2
                    frmH.lblProgress.Text = strP3
                    frmH.lblProgress.Refresh()

                    'int1 = ctSampleDesign 'number of table rows

                    'open reassay recordset
                    str3 = "ANALYTEID=" & strAnalyteID
                    Erase drows
                    drows = tblReassay.Select(str3)
                    int1 = drows.Length
                    Count2 = 0

                    wrdSelection = wd.Selection()

                    ''''''''wdd.visible = True

                    ctrsSamples(1, Count2A) = ctSampleDesign

                    strP2 = "Creating Word table..." & ChrW(10) & "If the table is large, this may take a few moments..."
                    strP3 = strP1 & ChrW(10) & strP2
                    frmH.lblProgress.Text = strP3
                    frmH.lblProgress.Refresh()

                    '20160312 LEE:
                    'amazingly, all these formatting actions below take a long time if table is large
                    'e.g. AbbVie M13099DAA. Table is ~65 pages


                    Try

                        '20180913 LEE:
                        Call IncrNextTableNumber(wd)

                        If boolPlaceHolder Then
                            .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=1, NumColumns:=1, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                        Else
                            .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=2 + ctSampleDesign, NumColumns:=ctCols, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                        End If

                        .Selection.Tables.Item(1).Rows.AllowBreakAcrossPages = False
                        .Selection.Tables.Item(1).Columns.PreferredWidth = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPoints

                        .Selection.Tables.Item(1).Select()

                        Call SetCellPaddingZero(.Selection.Tables.Item(1))

                        .Selection.Rows.AllowBreakAcrossPages = False


                        '20160908 LEE:
                        'check to see if any columns need wordwrap = false
                        'do this before adding the table number row
                        For Count2 = 1 To ctCols
                            idTCHL = arrOrder(3, Count2)
                            If idTCHL = 214 Or idTCHL = 219 Then '20160908 LEE: these are date columns. Don't let them wrap
                                boolHasDate = True
                                'select column
                                .Selection.Tables.Item(1).Cell(1, Count2).Select()
                                .Selection.SelectColumn()
                                Call DoCells(.Selection.Cells)
                                var1 = var1
                            End If
                        Next Count2


                        ''''wdd.visible = True

                        'With .Selection.Tables.Item(1)
                        '    .TopPadding = 0 'InchesToPoints(0)
                        '    .BottomPadding = 0 ' InchesToPoints(0)
                        '    .LeftPadding = 2 'InchesToPoints(0.03)
                        '    .RightPadding = 2 'InchesToPoints(0.03)
                        '    '.WordWrap = True
                        '    '.FitText = False
                        'End With

                        .Selection.Tables.Item(1).Select()

                        Call SetCellPadding(.Selection.Tables.Item(1))

                        Call removeBorderButLeaveTopAndBottom(wd)
                        '.Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
                        .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalTop
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

                        '20171220 LEE: Do not set table size, use the style default table
                        '.Selection.Font.Size = fontsize - 1
                        .Selection.Tables.Item(1).Cell(1, 1).Select()

                        'enter headings
                        For Count4 = 1 To ctCols
                            .Selection.Tables.Item(1).Cell(1, Count4).Select()
                            .Selection.Text = Replace(arrOrder(4, Count4), " ", ChrW(10), 1, -1, CompareMethod.Text)
                        Next


                        'border top and bottom of range
                        .Selection.Tables.Item(1).Cell(1, 1).Select()

                        '.Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        'format alignment and table fit
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                        .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom 'bottom align header row

                        'autofit table
                        Call AutoFitTable(wd, False)

                        'begin entering sample info
                        Count4 = 5 'Counter for row selection
                        Count5 = 0 'Counter for arr
                        .Selection.Tables.Item(1).Cell(3, 1).Select()
                        'arrSampleConcs(1 To 7, 1 To ctSampleDesign)
                        '1=ANALYTEINDEX, 2=RUNID, 3=CONCENTRATION, 4=DESIGNSAMPLEID, 5=ALIQUOTFACTOR, 6=CONCENTRATIONSTATUS, 7=REPLICATES
                        'arrSampleDesign()
                        '1=DESIGNSUBJECTTAG, 2=SUBJECTGROUPNAME, 3=ENDDAY, 4=ENDHOUR, 5=CONCENTRATION, 6=ALIQUOTFACTOR, 7=ID

                        '*****
                        'generat array
                        Dim arrPaste()
                        Dim strPaste As String = ""
                        Dim strPasteT As String = ""
                        ReDim arrPaste(ctSampleDesign)
                        Dim numBQLA As Decimal ' Single
                        Dim numAQLA As Decimal ' Single
                        Dim numDF As Decimal ' Single

                        Dim intColConc As Short = 0

                        strM = "Entering Summary of " & strAnalyteDescription & " in " & strMatrix & " Concentrations table..."
                        strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."

                        strP1 = "Entering Summary of " & strAnalyteDescription & " in " & strMatrix & " Concentrations table..."
                        strP1 = strP1 & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                        strP2 = "Sample 1 of " & ctSampleDesign & " samples..."
                        strP3 = strP1 & ChrW(10) & strP2
                        frmH.lblProgress.Text = strP3
                        frmH.lblProgress.Refresh()

                        strPaste = ""
                        Erase arrSS
                        ReDim arrSS(lbSS, ctSampleDesign)
                        '1=Row, 2=Col

                        intCTSD = 0
                        For Count2 = 1 To ctSampleDesign

                            intCTSD = intCTSD + 1
                            If intCTSD > intCTSDMax Then
                                strP2 = "Sample " & Count2 & " of " & ctSampleDesign & " samples..."
                                strP3 = strP1 & ChrW(10) & strP2
                                frmH.lblProgress.Text = strP3
                                frmH.lblProgress.Refresh()
                                intCTSD = 0
                            End If

                            strErr = ""

                            If Count2 = 2644 Then
                                var1 = var1 'debug
                            End If
                            Count4 = Count4 + 1
                            Count5 = Count5 + 1

                            'Find RunID
                            Dim strRunID As String

                            var1 = NZ(arrSampleDesign(ctCols + 2, Count2), 0) 'runid
                            'Note: RunID may be "NA"
                            If IsNumeric(var1) Then
                                If var1 = 0 Then 'search for RunID
                                    Try
                                        var2 = arrSampleDesign(ctCols + 3, Count2) 'designsampleid
                                        str1 = "DESIGNSAMPLEID = " & var2
                                        Erase drowsF
                                        drowsF = tblReassay.Select(str1)
                                        int1 = drowsF.Length
                                    Catch ex As Exception
                                        var4 = ex.Message
                                        var4 = var4
                                    End Try

                                    Try
                                        If int1 = 0 Then
                                            strRunID = "Problem"
                                        Else
                                            strRunID = drowsF(0).Item("RUNID")
                                        End If
                                    Catch ex As Exception
                                        var4 = ex.Message
                                        var4 = var4
                                    End Try

                                    '.Selection.TypeText Text:=CStr(NZ(var3, "Problem"))
                                Else
                                    Try
                                        strRunID = CStr(arrSampleDesign(ctCols + 2, Count2))
                                    Catch ex As Exception
                                        var4 = ex.Message
                                        var4 = var4
                                    End Try

                                End If
                            Else
                                strRunID = var1
                            End If

                            str1 = " "

                            vDay = "NA"
                            vTime = "NA"
                            vSubject = "NA"
                            Try
                                For Count3 = 1 To ctCols

                                    Try
                                        var1 = arrSampleDesign(Count3, Count2)
                                        str1 = CStr(NZ(arrSampleDesign(Count3, Count2), " ")) 'value
                                        var2 = arrOrder(1, Count3) 'Header title
                                    Catch ex As Exception
                                        var4 = ex.Message
                                        var4 = var4
                                    End Try

                                    Select Case var2


                                        Case "Dose Amount"
                                            str1 = NZ(var1, "NA")

                                        Case "Day"
                                            vDay = NZ(arrSampleDesign(Count3, Count2), "NA")
                                            vDay = vDay 'debug
                                        Case "Time"
                                            vTime = NZ(arrSampleDesign(Count3, Count2), "NA")
                                            vTime = vTime 'debug
                                        Case "Subject"
                                            vSubject = NZ(arrSampleDesign(Count3, Count2), "NA")
                                        Case "Concentration"
                                            'Look up numBQL for this run
                                            intColConc = Count3

                                            '20180219 LEE:
                                            'if

                                            If StrComp(strRunID, "Problem") = 0 Or StrComp(strRunID, "NA", CompareMethod.Text) = 0 Then
                                                strErr = "ERROR - SRSummary: The RunID for one of the Samples cannot be determined."
                                                strErr = strErr & ChrW(10) & "Table row: " & Count2 & ", Subject: " & vSubject & ", Day: " & vDay & ", Time: " & vTime
                                                'console.writeline(strErr)
                                                'MsgBox(strErr, vbInformation, "Problem...")
                                                str1 = "NA" 'Can't report, because we don't knowif it's within range.
                                            Else

                                                'arrSampleDesign
                                                '1=DESIGNSUBJECTTAG, 2=SUBJECTGROUPNAME, 3=ENDDAY, 4=ENDHOUR, 5=CONCENTRATION, 6=ALIQUOTFACTOR, 7=ID, 8=DESIGNSAMPLEID, 9=GENDER
                                                'Order Columns + 4
                                                'ctcols+1=ALIQUTOFACTOR, +2=RUNID, +3=DESIGNSAMPLEID, +4=COMMENTMEMO

                                                numBQL = getRunAnalyteLLOQ(strRunID, strAnalyteID, intGroup)
                                                numAQL = getRunAnalyteULOQ(strRunID, strAnalyteID, intGroup)
                                                '20160518 LEE: This may be numeric, NM, VEC, or some other text like 'Out of Range', or NR
                                                var5 = arrSampleDesign(Count3, Count2) 'this value may be text, e.g. NR, BQL, or AQL
                                                If IsNumeric(var5) Then
                                                    boolNum = True
                                                Else
                                                    boolNum = False
                                                    If StrComp(var5, "NR", CompareMethod.Text) = 0 Then
                                                        boolNR = True

                                                        'Add to Legend Array
                                                        strA = "NR"
                                                        strB = "Not Reported"
                                                        ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, strB, strA, True)
                                                        intLeg = ctLegend

                                                    End If
                                                End If

                                                Try
                                                    numDF = NZ(arrSampleDesign(ctCols + 1, Count2), 1) 'ALIQUOTFACTOR
                                                    numBQLA = numBQL / numDF
                                                    numAQLA = numAQL / numDF
                                                    If boolBQLLEGEND Then
                                                        strBQL = BQL()
                                                        strAQL = AQL()
                                                    Else
                                                        'if samples have been diluted, then LLOQ must be adjusted accordingly
                                                        If boolLUseSigFigs Then
                                                            strBQL = BQL() & "(<" & DisplayNum(SigFigOrDec(numBQLA, LSigFig, False), LSigFig, False) & ")"
                                                            strAQL = AQL() & "(>" & DisplayNum(SigFigOrDec(numAQLA, LSigFig, False), LSigFig, False) & ")"
                                                        Else
                                                            strBQL = BQL() & "(<" & Format(numBQLA, GetRegrDecStr(LSigFig)) & ")"
                                                            strAQL = AQL() & "(>" & Format(numAQLA, GetRegrDecStr(LSigFig)) & ")"
                                                        End If

                                                    End If

                                                Catch ex As Exception
                                                    var1 = ex.Message
                                                End Try

                                                If boolNum Then

                                                    num1 = NZ(arrSampleDesign(Count3, Count2), 0) 'CONCENTRATION
                                                    'var1 = NZ(arrSampleDesign(ctCols + 1, Count2), 1) 'ALIQUOTFACTOR
                                                    num1 = num1 / numDF 'ALIQUOTFACTOR
                                                    If boolLUseSigFigs Then
                                                        num1 = SigFigOrDec(num1, LSigFig, False)
                                                    Else
                                                        num1 = RoundToDecimalRAFZ(num1, LSigFig)
                                                    End If
                                                    'numBQLA = numBQL / numDF
                                                    'numAQLA = numAQL / numDF

                                                    'If boolBQLLEGEND Then
                                                    '    strBQL = BQL()
                                                    '    strAQL = AQL()
                                                    'Else
                                                    '    'if samples have been diluted, then LLOQ must be adjusted accordingly
                                                    '    If boolLUseSigFigs Then
                                                    '        strBQL = BQL() & "(<" & DisplayNum(SigFigOrDec(numBQLA, LSigFig, False), LSigFig, False) & ")"
                                                    '        strAQL = AQL() & "(>" & DisplayNum(SigFigOrDec(numAQLA, LSigFig, False), LSigFig, False) & ")"
                                                    '    Else
                                                    '        strBQL = BQL() & "(<" & Format(numBQLA, GetRegrDecStr(LSigFig)) & ")"
                                                    '        strAQL = AQL() & "(>" & Format(numAQLA, GetRegrDecStr(LSigFig)) & ")"
                                                    '    End If

                                                    'End If

                                                Else
                                                    'Concentration was null, so 'NR' reported
                                                    'look for DECISIONREASON, but don't report it yet
                                                    'need to implement a superscript function for this table first
                                                    ''don't do any of this for now
                                                    'just report CONCENTRATIONSTATUS from earlier
                                                    'Dim drNR() As DataRow
                                                    'Dim strFNR As String
                                                    '' '1=DESIGNSUBJECTTAG, 2=SUBJECTGROUPNAME, 3=ENDDAY, 4=ENDHOUR, 5=CONCENTRATION, 6=ALIQUOTFACTOR, 7=ID, 8=DESIGNSAMPLEID, 9=GENDER

                                                    'strFNR = "ANALYTEID = " & intAnalyteID & " AND SAMPLETYPEID = '" & strMatrix & "' AND DESIGNSAMPLEID = " & NZ(arrSampleDesign(8, Count2), 0)
                                                    'drNR = tblSAMPRESCONFLICTDEC.Select(strFNR)
                                                    'var4 = GetDECISIONREASONValue(False, intAnalyteID, drNR(0))
                                                    'str1 = var5

                                                    ''Add to Legend Array
                                                    'intLeg = intLeg + 1
                                                    'strA = ChrW(96 + intLeg)
                                                    'ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA,TRUE)

                                                    ''20160518 LEE:
                                                    'If StrComp(var5, "NM", CompareMethod.Text) = 0 Then 'bql
                                                    '    boolNum = True
                                                    '    num1 = 0
                                                    'ElseIf StrComp(var5, "VEC", CompareMethod.Text) = 0 Then 'AQL
                                                    '    boolNum = True
                                                    '    num1 = numAQLA + 1
                                                    'ElseIf StrComp(var5, "NR", CompareMethod.Text) = 0 Then 'report NR
                                                    '    str1 = var5
                                                    'ElseIf StrComp(var5, "Out of Range", CompareMethod.Text) = 0 Then 'repor NR
                                                    '    str1 = "NR"
                                                    'Else
                                                    '    str1 = var5
                                                    'End If

                                                    '20170710 LEE:
                                                    If StrComp(var5, "NR", CompareMethod.Text) = 0 Or StrComp(var5, "Out of Range", CompareMethod.Text) = 0 Then 'report NR
                                                        str1 = "NR"

                                                        'Add to Legend Array
                                                        strA = "NR"
                                                        str2 = "Not Reported"
                                                        ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str2, strA, False)
                                                        intLeg = ctLegend
                                                        boolNR = True

                                                    Else
                                                        str1 = var5
                                                    End If

                                                End If

                                                If boolNum Then
                                                    If num1 < numBQLA Then
                                                        If boolBQLSHOWCONC Then
                                                            If boolBQLLEGEND Then
                                                                If boolLUseSigFigs Then
                                                                    str1 = DisplayNum(num1, LSigFig, False) & " "
                                                                Else
                                                                    str1 = Format(num1, GetRegrDecStr(LSigFig)) & " "
                                                                End If

                                                                str1 = str1 & "(" & strBQL & ")"
                                                            Else
                                                                If boolLUseSigFigs Then
                                                                    str1 = DisplayNum(num1, LSigFig, False) & " " 'soft return
                                                                Else
                                                                    str1 = Format(num1, GetRegrDecStr(LSigFig)) & " " 'soft return
                                                                End If

                                                                str1 = str1 & strBQL
                                                            End If
                                                        Else
                                                            str1 = strBQL
                                                        End If

                                                    ElseIf num1 > numAQLA Then

                                                        Try
                                                            If boolBQLSHOWCONC Then
                                                                If boolBQLLEGEND Then
                                                                    If boolLUseSigFigs Then
                                                                        str1 = DisplayNum(num1, LSigFig, False) & " "
                                                                    Else
                                                                        str1 = Format(num1, GetRegrDecStr(LSigFig)) & " "
                                                                    End If

                                                                    str1 = str1 & "(" & strAQL & ")"
                                                                Else
                                                                    If boolLUseSigFigs Then
                                                                        str1 = DisplayNum(num1, LSigFig, False) & " " 'soft return
                                                                    Else
                                                                        str1 = Format(num1, GetRegrDecStr(LSigFig)) & " " 'soft return
                                                                    End If

                                                                    str1 = str1 & strAQL
                                                                End If
                                                            Else
                                                                str1 = strAQL
                                                            End If
                                                        Catch ex As Exception
                                                            var4 = ex.Message
                                                            var4 = var4
                                                        End Try

                                                    Else

                                                        Try
                                                            If boolLUseSigFigs Then
                                                                str1 = DisplayNum(num1, LSigFig, False)
                                                            Else
                                                                str1 = Format(num1, GetRegrDecStr(LSigFig))
                                                            End If
                                                        Catch ex As Exception
                                                            var4 = ex.Message
                                                            var4 = var4
                                                        End Try

                                                    End If
                                                Else

                                                End If


                                                'check to see if there is a COMMENTMEMO
                                                Try
                                                    If BOOLCONCCOMMENTS Then
                                                        var2 = NZ(arrSampleDesign(ctCols + 4, Count2), "")
                                                        If Len(var2) = 0 Then
                                                        Else
                                                            'superscript this row and column
                                                            intSS = intSS + 1
                                                            arrSS(1, intSS) = Count2 + 2 'account for header rows
                                                            arrSS(2, intSS) = intColConc

                                                            'Add to Legend Array
                                                            intLeg = intLeg + 1
                                                            strA = ChrW(96 + intLeg)
                                                            ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, var2.ToString, strA, True)
                                                            'strA = ChrW(96 + ctLegend)
                                                            arrSS(3, intSS) = strA
                                                            'superscript later
                                                            intLeg = ctLegend

                                                        End If
                                                    End If

                                                Catch ex As Exception
                                                    var3 = ex.Message
                                                End Try

                                            End If
                                        Case "Watson Run ID"
                                            str1 = strRunID

                                        Case "Visit Text"
                                            str1 = NZ(var1, "")

                                        Case "Time Text"
                                            str1 = NZ(var1, "")

                                        Case strGroupCheck
                                            If Count2 = 1 Then
                                                strSub1 = str1
                                                strSub2 = str1
                                            Else
                                                strSub2 = str1
                                            End If

                                        Case "Time", "Start Time"
                                            str1 = "a" 'debug

                                        Case Else

                                    End Select

                                    If Count3 = 1 Then
                                        strPasteT = str1
                                    Else
                                        strPasteT = strPasteT & ChrW(9) & str1
                                    End If

                                    If StrComp(str1, "NA", CompareMethod.Text) = 0 Then
                                        boolNA = True
                                    End If

                                Next
                            Catch ex As Exception
                                var4 = ex.Message
                                var4 = var4
                            End Try

                            Try
                                If Count2 = 1 Then
                                    strPaste = strPasteT
                                Else
                                    strPaste = strPaste & ChrW(10) & strPasteT
                                End If
                            Catch ex As Exception
                                var4 = ex.Message
                                var4 = var4
                            End Try


                        Next Count2

                        ReDim Preserve arrSS(lbSS, intSS)

                        strP2 = "Finished Sample " & ctSampleDesign & " of " & ctSampleDesign & " samples..."
                        strP3 = strP1 & ChrW(10) & strP2
                        frmH.lblProgress.Text = strP3
                        frmH.lblProgress.Refresh()

                        strM = "Entering Summary of " & strAnalyteDescription & " in " & strMatrix & " Concentrations table..."
                        strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                        frmH.lblProgress.Text = strM
                        frmH.lblProgress.Refresh()

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
                        ''select appropriate rows
                        '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdColumn, Extend:=True)

                        Dim rng1 As Word.Range
                        Dim tblW As Word.Table

                        tblW = .Selection.Tables.Item(1)

                        Try
                            rng1 = wd.ActiveDocument.Range(Start:=tblW.Cell(3, 1).Range.Start, End:=tblW.Cell(tblW.Rows.Count, ctCols).Range.End)
                            rng1.Select()
                        Catch ex As Exception
                            'select appropriate rows
                            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdColumn, Extend:=True)
                            var1 = ex.Message
                            var1 = var1
                        End Try

                        Pause(0.1)

                        'paste from clipboard
                        Try
                            .Selection.PasteAndFormat(Microsoft.Office.Interop.Word.WdRecoveryType.wdPasteDefault)
                        Catch ex As Exception
                            'MsgBox("Paste: " & ex.Message)
                        End Try


                        'the paste action removes the range object and any table formatting, must reset it
                        Call GlobalTableParaFormat(wd)

                        Try
                            rng1 = wd.ActiveDocument.Range(Start:=tblW.Cell(3, 1).Range.Start, End:=tblW.Cell(tblW.Rows.Count, ctCols).Range.End)
                            rng1.Select()
                        Catch ex As Exception
                            'select appropriate rows
                            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdColumn, Extend:=True)
                            var1 = ex.Message
                            var1 = var1
                        End Try
                        ''the paste action removes paragraph formatting, must replace it
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                        .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

                        '20171220 LEE: Do not set table size, use the style default table
                        '.Selection.Font.Size = fontsize - 1
                        '*****

                    Catch ex As Exception

                        str1 = "There was a problem preparing table:"
                        str1 = strM1 & ChrW(10) & ChrW(10) & str1
                        str1 = str1 & ChrW(10) & ChrW(10)
                        str1 = str1 & ex.Message
                        MsgBox(str1, vbInformation, "Problem...")

                    End Try

                    'superscript cells
                    Try
                        For Count2 = 1 To intSS

                            int1 = arrSS(1, Count2) 'row
                            int2 = arrSS(2, Count2) 'col
                            str1 = arrSS(3, Count2) 'superscript
                            .Selection.Tables(1).Cell(int1, int2).Select()
                            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine)
                            Call typeInSuperscriptFontSize12WithSpace(wd, str1)
                        Next
                    Catch ex As Exception
                        var1 = ex.Message
                    End Try

                    str1 = "Entering Summary of " & strAnalyteDescription & " Concentrations table: Grouping data..."
                    str1 = str1 & ChrW(10) & "Table " & intTCur & " of " & intTTot & " Tables..."
                    str1 = str1 & ChrW(10) & "If the table is large, this step may be lengthy..."
                    frmH.lblProgress.Text = str1
                    frmH.lblProgress.Refresh()

                    'find columns to evaluate
                    Dim arrCols(2, ctCols)
                    '1=header, 2=column#
                    Dim intCtCols As Short
                    intCtCols = 0
                    For Count2 = 1 To intGroups
                        var1 = arrGroups(1, Count2)
                        For Count3 = 1 To ctCols
                            var2 = arrOrder(1, Count3)
                            If StrComp(var1, var2, CompareMethod.Text) = 0 Then
                                intCtCols = intCtCols + 1
                                arrCols(1, intCtCols) = var1
                                arrCols(2, intCtCols) = Count3
                                Exit For
                            End If
                        Next
                    Next

                    'now record rows that need to be insert
                    'arrSampleDesign: 1=Cols , 2=Rows
                    Dim intTblRows As Short
                    Dim intTblRow As Short
                    Dim arrInsertRows(ctSampleDesign)
                    Dim intInsertRows As Short
                    Dim intOffS As Short

                    int1 = .Selection.Tables.Item(1).Rows.Count

                    intInsertRows = 0

                    intOffS = int1 - ctSampleDesign
                    '''''wdd.visible = True
                    For Count2 = 1 To intCtCols
                        int1 = arrCols(2, Count2)
                        intTblRows = .Selection.Tables.Item(1).Rows.Count
                        intTblRow = intTblRows
                        For Count3 = 1 To ctSampleDesign
                            If Count3 = 1 Then
                                var1 = CStr(NZ(arrSampleDesign(int1, Count3), 0))
                                var2 = var1
                            Else
                                var2 = var1
                                var1 = CStr(NZ(arrSampleDesign(int1, Count3), 0))
                            End If
                            If StrComp(var1, var2, CompareMethod.Text) = 0 Then
                            Else

                                intInsertRows = intInsertRows + 1
                                If intInsertRows > UBound(arrInsertRows) Then
                                    ReDim Preserve arrInsertRows(intInsertRows)
                                End If
                                arrInsertRows(intInsertRows) = Count3 + intOffS

                            End If
                        Next
                    Next

                    'sort these rows
                    var1 = var1 'debug
                    For Count2 = 1 To intInsertRows
                        int1 = arrInsertRows(Count2)
                        For Count3 = Count2 To intInsertRows
                            int2 = arrInsertRows(Count3)
                            If int2 < int1 Then
                                int3 = int1
                                int1 = int2
                                int2 = int3
                                arrInsertRows(Count2) = int1
                                arrInsertRows(Count3) = int2
                            End If
                        Next
                    Next
                    var1 = var1 'debug
                    ''debug
                    '''''''''console.writeline("Begin")
                    'For Count2 = 1 To intInsertRows
                    '    ''''''''console.writeline(arrInsertRows(Count2))
                    'Next
                    '''''''''console.writeline("End")

                    strP1 = "Entering Summary of " & strAnalyteDescription & " in " & strMatrix & " Concentrations table..."
                    strP1 = strP1 & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    strP2 = "Inserting Row 1 of " & intInsertRows & " Group Rows..."
                    strP3 = strP1 & ChrW(10) & strP2
                    frmH.lblProgress.Text = strP3
                    frmH.lblProgress.Refresh()

                    intCTSD = 0

                    For Count2 = intInsertRows To 1 Step -1

                        intCTSD = intCTSD + 1
                        If intCTSD > intCTSDMax Then
                            strP2 = "Inserting Row " & intInsertRows - Count2 & " of " & intInsertRows & " Group Rows..."
                            strP3 = strP1 & ChrW(10) & strP2
                            frmH.lblProgress.Text = strP3
                            frmH.lblProgress.Refresh()
                            intCTSD = 0
                        End If

                        If Count2 = intInsertRows Then
                            int1 = arrInsertRows(Count2)
                            int2 = int1
                        Else
                            int1 = int2
                            int2 = arrInsertRows(Count2)
                        End If
                        If int1 = int2 And Count2 <> intInsertRows Then 'skip
                            int1 = int2
                        Else
                            '''''wdd.visible = True
                            .Selection.Tables.Item(1).Cell(int2, 1).Select()
                            .Selection.InsertRowsAbove(1)
                        End If
                    Next

                    strP2 = "Inserting Row " & intInsertRows & " of " & intInsertRows & " Group Rows..."
                    strP3 = strP1 & ChrW(10) & strP2
                    frmH.lblProgress.Text = strP3
                    frmH.lblProgress.Refresh()

                    int1 = .Selection.Tables.Item(1).Rows.Count
                    .Selection.Tables.Item(1).Cell(int1, 1).Select()
                    'border bottom of thellos table
                    .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                    .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                    strM = "Final Formatting of Summary of " & strAnalyteDescription & " in " & strMatrix & " Concentrations Table..."
                    strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " Tables..."
                    strM = strM & ChrW(10) & "If the table is large, this step may be lengthy..."
                    frmH.lblProgress.Text = strM
                    frmH.lblProgress.Refresh()

                    'enter table number
                    dv = frmH.dgvDataWatson.DataSource
                    'dv1 = frmH.dgCompanyAnalRef.DataSource
                    var1 = strAnalyteDescription
                    str2 = "Summary of " & var1 & " Concentrations in"
                    var1 = frmH.cbxAnticoagulant.Text
                    str2 = str2 & " " & var1 & " Buffered"
                    int1 = FindRowDV("Species", dv)
                    var1 = dv.Item(int1).Item(1)
                    var1 = Capit(CStr(LowerCase(var1)))
                    str2 = str2 & " " & var1
                    int1 = FindRowDV("Matrix", dv)
                    var1 = dv.Item(int1).Item(1)
                    var1 = Capit(CStr(LowerCase(var1)))
                    str2 = str2 & " " & var1
                    'str2 = str2 & " Study Samples For " & Sheets("Data").Range("SubmittedTo1").Offset(0, 1).Value
                    str2 = str2 & " Study Samples For " & NZ(strSponsor, "[NA]")

                    'determine if Study Number is to be added
                    dv = frmH.dgvDataCompany.DataSource
                    int1 = FindRowDV("Sponsor Study Number", dv)
                    var1 = NZ(dv.Item(int1).Item(1), "NA")
                    If StrComp(var1, "NA", CompareMethod.Text) = 0 Then
                    Else
                        str2 = str2 & " Study " & var1
                    End If

                    'strTName = strTName.Replace("[MATRIX]", strMatrix) 'Need to do this here for now
                    '20160221 LEE: Made function to update. 
                    strTName = UpdateAnalyteMatrix(strTName, strAnalyteDescription, strMatrix, False, 0, False)
                    Call EnterTableNumber(wd, strTName, 3, strAnalyteDescription, strTempInfo, intTableID, 1, idTR)
                    'Note: strTName is byRef and will return Table, number, caption, label


                    ''''''''wdd.visible = True


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


                    If boolNA Then
                        ctLegend = ctLegend + 1
                        arrLegend(1, ctLegend) = "NA"
                        arrLegend(2, ctLegend) = "Not Applicable"
                        arrLegend(3, ctLegend) = False
                        arrLegend(4, ctLegend) = False
                    End If

                    'NR done earlier

                    If StrComp(BQL, "BQL", CompareMethod.Text) = 0 Then

                        ctLegend = ctLegend + 1

                        arrLegend(1, ctLegend) = BQL()
                        If boolBQLLEGEND Then
                            If boolLUseSigFigs Then
                                arrLegend(2, ctLegend) = BQLVerbose() & " (" & DisplayNum(SigFigOrDec(numBQL, LSigFig, False), LSigFig, False) & " " & strConcUnits & ")"
                            Else
                                arrLegend(2, ctLegend) = BQLVerbose() & " (" & Format(SigFigOrDec(numBQL, LSigFig, False), GetRegrDecStr(LSigFig)) & " " & strConcUnits & ")"
                            End If

                        Else
                            arrLegend(2, ctLegend) = BQLVerbose()
                        End If
                        arrLegend(3, ctLegend) = False

                        ctLegend = ctLegend + 1

                        arrLegend(1, ctLegend) = AQL()
                        If boolBQLLEGEND Then
                            If boolLUseSigFigs Then
                                arrLegend(2, ctLegend) = AQLVerbose() & " (" & DisplayNum(SigFigOrDec(numAQL, LSigFig, False), LSigFig, False) & " " & strConcUnits & ")"
                            Else
                                arrLegend(2, ctLegend) = AQLVerbose() & " (" & Format(SigFigOrDec(numAQL, LSigFig, False), GetRegrDecStr(LSigFig)) & " " & strConcUnits & ")"
                            End If

                        Else
                            arrLegend(2, ctLegend) = AQLVerbose()
                        End If

                    Else

                        ctLegend = ctLegend + 1

                        arrLegend(1, ctLegend) = BQL()
                        If boolBQLLEGEND Then
                            If boolLUseSigFigs Then
                                arrLegend(2, ctLegend) = BQLVerbose() & " (" & DisplayNum(SigFigOrDec(numBQL, LSigFig, False), LSigFig, False) & " " & strConcUnits & ")"
                            Else
                                arrLegend(2, ctLegend) = BQLVerbose() & " (" & Format(SigFigOrDec(numBQL, LSigFig, False), GetRegrDecStr(LSigFig)) & " " & strConcUnits & ")"
                            End If

                        Else
                            arrLegend(2, ctLegend) = BQLVerbose() & ""
                        End If
                        arrLegend(3, ctLegend) = False

                        ctLegend = ctLegend + 1

                        arrLegend(1, ctLegend) = AQL()
                        If boolBQLLEGEND Then
                            If boolLUseSigFigs Then
                                arrLegend(2, ctLegend) = AQLVerbose() & " (" & DisplayNum(SigFigOrDec(numAQL, LSigFig, False), LSigFig, False) & " " & strConcUnits & ")"
                            Else
                                arrLegend(2, ctLegend) = AQLVerbose() & " (" & Format(SigFigOrDec(numAQL, LSigFig, False), GetRegrDecStr(LSigFig)) & " " & strConcUnits & ")"
                            End If

                        Else
                            arrLegend(2, ctLegend) = AQLVerbose()
                        End If
                    End If

                    arrLegend(3, ctLegend) = False
                    arrLegend(4, ctLegend) = False

                    ReDim Preserve arrLegend(4, ctLegend)

                    str1 = frmH.lblProgress.Text

                    'autofit table
                    '20160518 LEE: Too many examples of table not fitting correctly. 
                    'Call AutoFitTable(wd, boolDate)

                    '20180201 LEE: Getting big problems with tables not page breaking correctly
                    'make boolVis true no matter what
                    Call AutoFitTable(wd, True)

                    '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)
                    '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)

                    ReDim Preserve arrLegend(4, ctLegend)

                    strM = "Finalizing " & strTName & "..."
                    strM1 = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    str1 = strM1

                    frmH.lblProgress.Text = strM1
                    frmH.Refresh()

                    Call SplitTable(wd, 2, ctLegend, arrLegend, str1, False, ctLegend + 1, False, False, False, intTableID)
                    'Sub SplitTable(ByVal wd As Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, 
                    'ByVal arr As Object, ByVal strT As String, ByVal DoLegend As Boolean, ByVal intSplitRows As Short, ByVal boolSmallFont As Boolean)

                    ''''''''wdd.visible = True

                    'autofit table
                    Call AutoFitTable(wd, False)

                    Call MoveOneCellDown(wd)

                    Call InsertLegend(wd, intTableID, idTR, False, 1)


                    ''''''''wdd.visible = True

                    var1 = "" 'debugging

next1:

                Next Count2A
            Next Count1A

            Call SpellingOff(doc, True)

        End With

    End Sub


End Module
