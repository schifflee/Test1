Option Compare Text

Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.ComponentModel.PropertyDescriptorCollection
Imports Word = Microsoft.Office.Interop.Word
Imports Microsoft.VisualBasic
Imports System.IO


Module modStyle4


    'from Style2


    Sub MVCarryover_v1_35(ByVal wd As Microsoft.Office.Interop.Word.application, ByVal idTR As Int64)

        Dim numNomConc As Decimal
        Dim var1, var2, var3, var4, var5, var10
        Dim dvDo As system.data.dataview
        Dim strTName As String
        Dim intDo As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim Count4 As Short
        Dim Count5 As Short
        Dim Count10 As Short
        Dim strDo As String
        Dim bool As Boolean
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strTerm1HighConc As String
        Dim tbl1 As System.Data.DataTable
        Dim dv1 As system.data.dataview
        Dim rows1() As DataRow
        Dim intRows1 As Short
        Dim strF1 As String
        Dim tbl2 As System.Data.DataTable
        Dim dv2 As system.data.dataview
        Dim rows2() As DataRow
        Dim intRows2 As Short
        Dim strF2 As String
        Dim tbl3 As System.Data.DataTable
        Dim dv3 As system.data.dataview
        Dim rows3() As DataRow
        Dim intRows3 As Short
        Dim strF3 As String
        Dim intTableID As Short
        Dim tbl4 As System.Data.DataTable
        Dim dv4 As system.data.dataview
        Dim rows4() As DataRow
        Dim intRows4 As Short
        Dim strF4 As String
        Dim strS As String
        Dim intNumRuns As Short
        Dim dv As system.data.dataview
        Dim tblNumRuns As System.Data.DataTable
        Dim tblLevels As System.Data.DataTable
        Dim intNumLevels As Short
        Dim intTblRows As Short
        Dim wrdSelection As Microsoft.Office.Interop.Word.selection
        Dim strF As String
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim int4 As Short
        Dim int10 As Short
        Dim intRowsX As Short
        Dim tblX As System.Data.DataTable
        Dim varNom
        Dim strConcUnits As String
        Dim intLeg As Short
        Dim ctQCLegend As Short
        Dim ctDilLeg As Short
        Dim strA As String
        Dim strB As String

        Dim boolV As Boolean = wd.Visible

        Dim intRowsSection As Int16

        Dim int8Max As Int32

        Dim strCOBL As String 'carryover blank label

        Dim intLLOQ As Short = 0

        Dim intSS As Short
        Dim intEE As Short

        Dim ctLegend As Short
        Dim fontsize
        Dim boolPro As Boolean

        Dim boolHasSampleName As Boolean = False
        '20180809 LEE:

        Dim boolIULOQ As Boolean = Not (boolIncludePSAE) 'BOOLINCLUDEPSAE defines the exclusion of ULOQ column and is NOT
        Dim boolInjCol As Boolean = Not (BOOLCONCCOMMENTS) 'BOOLCONCCOMMENTS defines the exclusion of Inj column and is NOT

        Dim hi, lo
        Dim rows10() As DataRow
        Dim rows11() As DataRow
        Dim intRowsAnal As Short
        Dim arrFP(2, 20) 'FlagPercent array: 1=hi, 2=lo
        Dim strFP As String
        Dim numMean As Decimal
        Dim numBias As Decimal
        Dim numSD As Decimal
        Dim numSDIS As Decimal
        Dim tblZ As System.Data.DataTable
        Dim dvAn As system.data.dataview
        Dim tblAnGo As New System.Data.DataTable
        Dim p1, p2, p3, p4, p5, p6, p7, p8, p9, p10
        Dim strM As String
        Dim fonts
        Dim numDF As Decimal
        Dim DilFactor
        Dim strF2a As String
        Dim strTempInfo As String
        Dim rowsX() As DataRow
        Dim intLegStart As Short
        Dim boolJustTable As Boolean

        Dim intExp As Short
        Dim ctExp As Short
        Dim int8 As Short
        Dim intStart As Short

        Dim rows2E() As DataRow
        Dim nE As Short
        Dim nI As Short
        Dim boolOutHeadE As Boolean = False
        Dim boolOutHeadI As Boolean = False
        Dim boolDeleteRows As Boolean = False

        Dim tblRID As System.Data.DataTable
        Dim numRID As Short
        Dim ctTbl As Short
        Dim varAnal, varIS
        Dim v1, v2, vU


        Dim vAnalyteIndex
        Dim vMasterAssayID
        Dim vAnalyteID
        Dim tblAG As DataTable = tblAnalyteGroups 'tblAnalyteGroups has all analytes, not just accepted

        Dim intGroup As Short
        Dim strAnal As String
        Dim strAnalC As String
        Dim strMatrix As String
        Dim strTNameO As String
        Dim intRunID As Int16
        Dim strDECISIONREASON As String
        Dim boolExFromAS As Boolean

        Dim numPrec As Single
        Dim numPrecIS As Single
        Dim numTheor As Single
        Dim rowsData() As DataRow

        Dim boolBSNCol As Boolean = False

        Dim arrP(1, 1)
        Dim strPaste As String
        Dim strPasteT As String
        Dim strR As String = "_xyz_"

      

        Dim charFCID As String
        strF = "ID_TBLREPORTTABLE = " & idTR
        Dim rowsTR() As DataRow = tblReportTable.Select(strF)
        var1 = rowsTR(0).Item("CHARFCID")
        charFCID = NZ(var1, "NA")

        boolJustTable = False

        Cursor.Current = Cursors.WaitCursor

        ''''wdd.visible = True

        fontsize = wd.ActiveDocument.Styles("Normal").Font.Size ' wd.Selection.Font.Size
        fonts = fontsize ' wd.Selection.Font.Size

        Dim intBlankReps As Short

        Dim intIc As Short
        Dim intIcMax As Short = 5


        With wd

            intTableID = 35

            Dim strWRunId As String = GetWatsonColH(intTableID)

            dvDo = frmH.dgvReportTableConfiguration.DataSource
            strF = "id_tblconfigreporttables = " & intTableID
            intDo = FindRowDVNumByCol(intTableID, dvDo, "id_tblconfigreporttables")

            '***
            intDo = FindRowDVNumByCol(idTR, dvDo, "ID_TBLREPORTTABLE")

            var1 = dvDo(intDo).Item("CHARHEADINGTEXT")
            strTName = NZ(var1, "[NONE]")
            strTNameO = strTName

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

            tbl1 = tblAnalysisResultsHome
            tbl2 = tblAssignedSamples
            tbl3 = tblAssignedSamplesHelper
            tbl4 = tblAnalytesHome

            Dim tbl5 As System.Data.DataTable
            tbl5 = tblTableProperties
            strF = "ID_TBLREPORTTABLE = " & idTR
            Dim rowsTP() As DataRow = tbl5.Select(strF)

            Dim strBlankLabel As String = NZ(rowsTP(0).Item("CHARCARRYOVERLABEL"), "Blank")

            'ensure data has been entered
            strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idTR
            rowsX = tbl2.Select(strF)


            strF = "IsIntStd = 'No'"
            strS = "INTORDER ASC, IsIntStd ASC, AnalyteDescription ASC"
            rows11 = tblAnalytesHome.Select(strF, strS)
            intRowsAnal = rows11.Length

            ctTbl = 0
            For Count1 = 1 To intRowsAnal

                boolJustTable = False

                Dim arrLegend(4, 20)

                strTName = strTNameO

                ctLegend = 0

                Dim int11 As Short
                If boolSTATSDIFFCOL Then
                    int11 = 2
                Else
                    int11 = 1
                End If

                'for legend stuff

                intExp = 0
                ctExp = 0

                intLeg = 0
                ctQCLegend = 0
                ctDilLeg = 0
                ctLegend = 0
                strA = ""
                strB = ""
                arrLegend.Clear(arrLegend, 0, arrLegend.Length)
                arrFP.Clear(arrFP, 0, arrFP.Length)
                intLegStart = 96

                Dim strSuper As String
                Dim intSuper As Short = 96

                '20180711 LEE
                boolBSNCol = False

                'check if table is to be generated
                'strDo = arrAnalytes(1, Count1) 'record column name
                strDo = rows11(Count1 - 1).Item("ANALYTEDESCRIPTION")

                If UseAnalyte(CStr(strDo)) Then
                Else
                    GoTo next1
                End If

                bool = dvDo.Item(intDo).Item(strDo) 'find boolean value of dvDo column
                strAnal = strDo

                Dim strM1 As String
                If bool Then 'continue

                    ctTbl = ctTbl + 1

                    intTCur = intTCur + 1

                    'setup tables
                    If boolUseGroups Then
                        intGroup = tblAG.Rows(Count1 - 1).Item("INTGROUP")
                        strAnal = tblAG.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                        strAnalC = tblAG.Rows(Count1 - 1).Item("ANALYTEDESCRIPTION_C")
                        vAnalyteID = tblAG.Rows.Item(Count1 - 1).Item("ANALYTEID")
                        strMatrix = tblAG.Rows(Count1 - 1).Item("MATRIX")
                    Else
                        var1 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteIndex")
                        var2 = tbl4.Rows.Item(Count1 - 1).Item("MasterAssayID")
                        var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                        var4 = tbl4.Rows.Item(Count1 - 1).Item("ANALYTEID")
                        intGroup = 0
                        vAnalyteIndex = var1
                        vMasterAssayID = var2
                        vAnalyteID = var4
                        strMatrix = ""
                    End If

                    'ensure data has been entered
                    strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & CleanText(strDo) & "' AND ID_TBLREPORTTABLE = " & idTR
                    rowsX = tbl2.Select(strF)
                    If rowsX.Length = 0 Then

                        strM = "Creating " & strTName & " For " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & "..."
                        strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                        frmH.lblProgress.Text = strM
                        frmH.Refresh()
                        'MsgBox("Samples have not been assigned to this table.", MsgBoxStyle.Information, "Samples have not been assigned...")
                        boolJustTable = True
                        'page setup according to configuration
                        str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                        'insert page break
                        'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                        Call InsertPageBreak(wd)
                        Call PageSetup(wd, str1) 'L=Landscape, P=Portrait
                        GoTo end1
                    Else
                        boolJustTable = False
                    End If

                    'page setup according to configuration
                    str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                    'insert page break
                    'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                    Call InsertPageBreak(wd)
                    Call PageSetup(wd, str1) 'L=Landscape, P=Portrait

                    'ReDim arrBCQCs(8, 50) '1=LevelNumber, 2=NomConcentration, 3=ID, 4=FlagPercent, 5=Hello, 6=Lo, 7=#ofReplicates, 8=ASSAYID

                    strM = "Creating " & strTName & " For " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & "..."
                    strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    strM1 = strM
                    frmH.lblProgress.Text = strM
                    frmH.Refresh()

                    'get strConcUnits
                    int1 = FindRowDV("ULOQ Units", frmH.dgvWatsonAnalRef.DataSource)
                    strConcUnits = NZ(frmH.dgvWatsonAnalRef(Count1, int1).Value, "ng/mL")

                    int1 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
                    str1 = NZ(frmH.dgvStudyConfig(1, int1).Value, "")

                    If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
                    Else
                        strConcUnits = str1
                    End If


                    ''setup tables
                    'If boolUseGroups Then
                    '    intGroup = tblAG.Rows(Count1 - 1).Item("INTGROUP")
                    '    strAnal = tblAG.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                    '    strAnalC = tblAG.Rows(Count1 - 1).Item("ANALYTEDESCRIPTION_C")
                    '    vAnalyteID = tblAG.Rows.Item(Count1 - 1).Item("ANALYTEID")
                    '    strMatrix = tblAG.Rows(Count1 - 1).Item("MATRIX")
                    'Else
                    '    var1 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteIndex")
                    '    var2 = tbl4.Rows.Item(Count1 - 1).Item("MasterAssayID")
                    '    var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                    '    var4 = tbl4.Rows.Item(Count1 - 1).Item("ANALYTEID")
                    '    intGroup = 0
                    '    vAnalyteIndex = var1
                    '    vMasterAssayID = var2
                    '    vAnalyteID = var4
                    '    strMatrix = ""
                    'End If

                    ''setup tables
                    'var1 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteIndex")
                    'var2 = tbl4.Rows.Item(Count1 - 1).Item("MasterAssayID")
                    'var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                    'var4 = tbl4.Rows.Item(Count1 - 1).Item("ANALYTEID")
                    'strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                    'strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                    'strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                    'strF2 = strF2 & "ANALYTEINDEX = " & var1 & " AND "
                    'strF2 = strF2 & "ANALYTEID = " & var4 & " AND "
                    'strF2 = strF2 & "MASTERASSAYID = " & var2 ' & " AND "


                    If boolUseGroups Then
                        strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                        strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                        strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                        strF2 = strF2 & "INTGROUP = " & intGroup
                    Else
                        strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                        strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                        strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                        strF2 = strF2 & "ANALYTEINDEX = " & var1 & " AND "
                        strF2 = strF2 & "MASTERASSAYID = " & var2 ' & " AND "
                        'strF2 = strF2 & "ANALYTEID = " & var4 ' & "' AND "
                        'strF2 = strF2 & "BOOLINTSTD = 0"
                    End If

                    strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                    rows2 = tbl2.Select(strF2, strS)

                    Dim strFData As String
                    strFData = strF2

                    int1 = rows2.Length 'debug
                    dv2 = New DataView(tbl2, strF2, strS, DataViewRowState.CurrentRows)
                    Dim intProcRows As Short
                    intProcRows = dv2.Count
                    Dim intPR As Short
                    strF3 = strF2 & " AND CHARHELPER1 = 'BLANK'"
                    Dim dvPR As DataView = New DataView(tbl2, strF3, strS, DataViewRowState.CurrentRows)
                    intPR = dvPR.Count

                    'find # of LLOQ rows
                    strF3 = strF2 & " AND CHARHELPER1 = 'LLOQ'"
                    Dim dvLLOQ As DataView = New DataView(tbl2, strF3, strS, DataViewRowState.CurrentRows)
                    Dim intLLOQRows As Int16 = dvLLOQ.Count

                    'find # of ULOQ rows
                    strF3 = strF2 & " AND CHARHELPER1 = 'ULOQ'"
                    Dim dvULOQ As DataView = New DataView(tbl2, strF3, strS, DataViewRowState.CurrentRows)
                    Dim intULOQRows As Int16 = dvULOQ.Count

                    intRowsSection = 0
                    If intPR > intRowsSection Then
                        intRowsSection = intPR
                    End If
                    If intLLOQRows > intRowsSection Then
                        intRowsSection = intLLOQRows
                    End If
                    If intULOQRows > intRowsSection Then
                        intRowsSection = intULOQRows
                    End If

                    Dim tblRunID As System.Data.DataTable = dv2.ToTable("tt", True, "RUNID")
                    intNumRuns = tblRunID.Rows.Count

                    'get strConcUnits
                    intRunID = 0
                    int1 = 0
                    Do Until intRunID > 0
                        var1 = tblRunID.Rows(int1).Item("RUNID")
                        If IsDBNull(var1) Then
                        Else
                            intRunID = var1
                        End If
                        int1 = int1 + 1
                    Loop
                    strConcUnits = GetConcUnits(intRunID)

                    'LLOQ and Blanks are Required.  Find out what third type is (usually ULOQ)
                    Dim tblTerm1s As System.Data.DataTable = dv2.ToTable("Term1's", True, "CHARHELPER1")
                    If boolIULOQ Then

                        If tblTerm1s.Rows.Count = 3 Then
                        Else 'Something is wrong
                            strM = "The Carryover Table has the wrong number of Term1 assignments." & ChrW(10)
                            strM = strM & "It must have exactly 3 types, LLOQ, ULOQ (or QC-High for example)., and Blank." & ChrW(10)
                            strM = strM & "The table will error out. The user should go back to Assign Samples and ensure every analytical run has LLOQ, ULOQ (or QC-High for example), and Blank samples assigned."
                            MsgBox(strM, vbInformation, "Problem...")
                            '& "(though ULOQ can also be QC-High or other)")
                        End If
                    Else


                        If tblTerm1s.Rows.Count = 2 Then
                        ElseIf tblTerm1s.Rows.Count = 3 Then '20180810 LEE: user may assign ULOQ, but choose not to show column
                        Else 'Something is wrong

                            strM = "The Carryover Table has the wrong number of Term1 assignments." & ChrW(10)
                            strM = strM & "It must have exactly 2 types, LLOQ and Blank." & ChrW(10)
                            strM = strM & "The table will error out. The user should go back to Assign Samples and ensure every analytical run has LLOQ and Blank samples assigned."

                            MsgBox(strM, vbInformation, "Problem...")
                            '& "(though ULOQ can also be QC-High or other)")
                        End If
                    End If

                    '

                    'Find which one is not LLOQ or Blank
                    If boolIULOQ Then
                        For Count10 = 0 To tblTerm1s.Rows.Count - 1
                            str1 = tblTerm1s.Rows(Count10).Item("CHARHELPER1").ToString
                            If (Not (StrComp(str1, "LLOQ") = 0) And Not (StrComp(str1, "Blank") = 0)) Then
                                strTerm1HighConc = str1
                            End If
                        Next
                    Else
                        strTerm1HighConc = ""
                    End If


                    'find number of table rows to generate
                    intRowsX = 0

                    boolOutHeadE = False
                    boolOutHeadI = False
                    boolDeleteRows = False

                    'generate table
                    intTblRows = 0
                    intTblRows = intTblRows + 2 'for header
                    intTblRows = intTblRows + intRowsSection 'section rows
                    'each section has:
                    '  1 blank row
                    '  n replicates (may be different for each run) => done already with intpr
                    '  if stats section
                    '     1 blank row
                    '     n stats rows

                    Dim intCSN As Short
                    intCSN = countNumStatsRows()
                    int1 = 1  '1 = blank row

                    If intCSN = 0 Then
                        int2 = int1
                    Else
                        int2 = int1 + 1 + intCSN
                    End If
                    intTblRows = intTblRows + (int2 * intNumRuns)
                    intTblRows = intTblRows 'debug

                    wrdSelection = wd.Selection()


                    'determine number of columns

                    Dim intColWRID As Short = 0
                    Dim intColBSN As Short = 0
                    Dim intColDate As Short = 0
                    Dim intColLLOQInjNum As Short = 0
                    Dim intColLLOQPA As Short = 0
                    Dim intColLLOQISPA As Short = 0
                    Dim intColULOQInjNum As Short = 0
                    Dim intColULOQPA As Short = 0
                    Dim intColULOQISPA As Short = 0
                    Dim intColBlankInjNum As Short = 0
                    Dim intColBlankPA As Short = 0
                    Dim intColBlankISPA As Short = 0
                    Dim intColCO As Short = 0
                    Dim intColISCO As Short = 0

                    Dim tblCHL As DataTable = tblConfigHeaderLookup
                    Dim tblRHC As DataTable = tblReportTableHeaderConfig

                    Dim strFCHL As String
                    Dim strFRHC As String

                    strFRHC = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLCONFIGREPORTTABLES = " & intTableID

                    '20160825 LEE: Hmm. Seems to be a replicate problem in tblReportTableHeaderConfig
                    'make a unique table for now
                    Dim dv10 As System.Data.DataView = New DataView(tblRHC, strFRHC, "INTORDER ASC", DataViewRowState.CurrentRows)
                    Dim tblRHCa As DataTable = dv10.ToTable("a", True, "CHARUSERLABEL", "ID_TBLCONFIGHEADERLOOKUP", "BOOLINCLUDE")

                    Dim rowsRHC() As DataRow = tblRHC.Select(strFRHC, "INTORDER ASC")

                    rowsRHC = tblRHC.Select(strFRHC, "INTORDER ASC")

                    Dim strDateCol As String = "Analysis Date"
                    Dim rowsDC() As DataRow = tblRHCa.Select("ID_TBLCONFIGHEADERLOOKUP = 296")
                    If rowsDC.Length = 0 Then
                    Else
                        '20180809 LEE:
                        'Remember, Analysis Date here is only for Column Labeling
                        Try
                            strDateCol = NZ(rowsDC(0).Item("CHARUSERLABEL"), "Analysis Date")
                        Catch ex As Exception
                            var1 = var1
                        End Try
                    End If

                    rowsDC = tblRHCa.Select("ID_TBLCONFIGHEADERLOOKUP = 248") 'Sample Name
                    If rowsDC.Length = 0 Then
                        boolBSNCol = False
                    Else
                        'var3 = tblRHCa.Rows(0).Item("BOOLINCLUDE")
                        '20180808 LEE:
                        var3 = rowsDC(0).Item("BOOLINCLUDE")
                        If var3 <> 0 Then
                            boolBSNCol = True
                        Else
                            boolBSNCol = False
                        End If
                    End If

                    'int std column
                    var1 = NZ(rowsTP(0).Item("BOOLINCLUDEISTBL"), "0")
                    Dim boolIntStd As Boolean
                    If var1 = 0 Then
                        boolIntStd = False
                    Else
                        boolIntStd = True
                    End If

                    Dim intCols As Short

                    intCols = 1 'runid
                    intColWRID = intCols

                    If boolBSNCol Then ' sample name
                        intCols = intCols + 1
                        intColBSN = intCols
                    End If

                    'do LLOQ Section
                    If boolInjCol Then
                        intCols = intCols + 1
                        intColLLOQInjNum = intCols
                    End If
                    intCols = intCols + 1 'Peak area
                    intColLLOQPA = intCols
                    If boolIntStd Then
                        intCols = intCols + 1
                        intColLLOQISPA = intCols
                    End If
                    intCols = intCols + 1 'spacer

                    If boolIULOQ Then
                        'do ULOQ Section
                        If boolInjCol Then
                            intCols = intCols + 1
                            intColULOQInjNum = intCols
                        End If
                        intCols = intCols + 1 'Peak area
                        intColULOQPA = intCols
                        If boolIntStd Then
                            intCols = intCols + 1
                            intColULOQISPA = intCols
                        End If
                        intCols = intCols + 1 'spacer
                    End If

                    'do blank section
                    If boolInjCol Then
                        intCols = intCols + 1
                        intColBlankInjNum = intCols
                    End If
                    intCols = intCols + 1 'Peak area
                    intColBlankPA = intCols
                    If boolIntStd Then
                        intCols = intCols + 1
                        intColBlankISPA = intCols
                    End If
                    intCols = intCols + 1 'spacer

                    intCols = intCols + 1 'Analyte%Carryover
                    intColCO = intCols

                    If boolIntStd Then
                        'IS%Carryover column
                        intCols = intCols + 1
                        intColISCO = intCols
                    End If

                    ReDim arrP(intCols, intTblRows + 100)

                    '*****

                    Try

                        '20180913 LEE:
                        Call IncrNextTableNumber(wd)

                        If boolPlaceHolder Then
                            .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=1, NumColumns:=1, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                        Else
                            .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=intTblRows, NumColumns:=intCols, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                        End If

                        .Selection.Tables.Item(1).Select()

                        '20180711 LEE:
                        'make this table wordwrap = false


                        'Call SetCellPaddingZero(.Selection.Tables.Item(1))
                        '20180711 LEE:
                        'SetCellPaddingZero makes SplitTable complicated

                        .Selection.Tables.Item(1).Rows.AllowBreakAcrossPages = False
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                        .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
                        .Selection.Tables.Item(1).Columns.PreferredWidth = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPoints
                        '.Selection.Tables.Item(1).Columns.Item(1).Width = 86
                        For Count2 = 1 To intNumLevels
                            '.Selection.Tables.item(1).Columns.item(Count2 + 1).Width = 50
                        Next
                        .Selection.Tables.Item(1).Select()

                        'remove border, but leave top and bottom
                        Call removeBorderButLeaveTopAndBottom(wd)

                        .Selection.Tables.Item(1).Cell(1, 1).Select()

                        Dim wdTbl As Word.Table = .Selection.Tables.Item(1)

                        Dim intCol As Short = 0

                        If boolPlaceHolder Then

                            .Selection.Tables.Item(1).Select()
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone

                            strA = arrAnalytes(14, Count1)
                            strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, False)
                            Call EnterTableNumber(wd, strTName, 3, strA, strTempInfo, intTableID, intGroup, idTR)
                            Call MoveOneCellDown(wd)

                            .Selection.TypeParagraph()
                            .Selection.TypeParagraph()

                            'enter a table record in tblTableN
                            'ctTableN = ctTableN + 1
                            Dim dtblr1 As DataRow = tblTableN.NewRow
                            dtblr1.BeginEdit()
                            dtblr1.Item("TableNumber") = ctTableN
                            dtblr1.Item("AnalyteName") = arrAnalytes(1, Count1)
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

                        '*****

                        'determine column headers

                        Dim intColCt As Short = 0

                        '20180809 LEE:
                        'Redo this code
                        For Count2 = 0 To tblRHCa.Rows.Count - 1

                            var1 = tblRHCa.Rows(Count2).Item("CHARUSERLABEL")
                            var2 = tblRHCa.Rows(Count2).Item("ID_TBLCONFIGHEADERLOOKUP")
                            var3 = NZ(tblRHCa.Rows(Count2).Item("BOOLINCLUDE"), 0)

                            Select Case var2

                                Case 247 'Watson Run ID
                                    intColCt = intColCt + 1
                                    .Selection.Tables.Item(1).Cell(2, intColCt).Select()
                                    intColWRID = intColCt
                                    If BOOLINCLUDEDATE Then
                                        '.Selection.TypeText(Text:=var1 & ChrW(10) & "(Analysis Date)")
                                        '20180420 LEE:
                                        '.Selection.TypeText(var1 & ChrW(10) & "(" & GetAnalysisDateLabel(intTableID) & ")")
                                        '20180809 LEE:
                                        .Selection.TypeText(var1 & ChrW(10) & "(" & strDateCol & ")")
                                    Else
                                        .Selection.TypeText(Text:=var1)
                                    End If
                                Case 248 'Sample Name

                                    If var3 = 0 Then
                                    Else
                                        intColCt = intColCt + 1
                                        intColBSN = intColCt
                                        .Selection.Tables.Item(1).Cell(2, intColCt).Select()
                                        .Selection.TypeText(Text:=var1)
                                        boolHasSampleName = True
                                    End If
                                Case 296 'Analysis Date - ignore


                            End Select

                        Next Count2

                        '*****

                        'do LLOQ section

                        intColCt = intColCt + 1
                        intSuper = intSuper + 1
                        strSuper = ChrW(intSuper)
                        .Selection.Tables.Item(1).Cell(1, intColCt).Select()
                        str1 = "LLOQ"
                        .Selection.TypeText(Text:=str1)

                        '20190108 LEE
                        If boolNONELEG Then
                        Else
                            .Selection.TypeText(Text:=ChrW(160))
                            Call typeInSuperscript(wd, strSuper)
                        End If

                        If boolInjCol Then
                            .Selection.Tables.Item(1).Cell(2, intColCt).Select()
                            str1 = "Injection" & ChrW(10) & "#"
                            .Selection.TypeText(Text:=str1)
                            intColCt = intColCt + 1
                        End If

                        intColLLOQPA = intColCt
                        .Selection.Tables.Item(1).Cell(2, intColCt).Select()
                        str1 = "Peak" & ChrW(10) & "Area"
                        .Selection.TypeText(Text:=str1)

                        If boolIntStd Then
                            intColCt = intColCt + 1
                            .Selection.Tables.Item(1).Cell(2, intColCt).Select()
                            str1 = "IS Peak" & ChrW(10) & "Area"
                            .Selection.TypeText(Text:=str1)
                        End If

                        'do ULOQ section
                        If boolIULOQ Then

                            intColCt = intColCt + 2
                            .Selection.Tables.Item(1).Cell(1, intColCt).Select()
                            str1 = strTerm1HighConc  'Previously "ULOQ"
                            .Selection.TypeText(Text:=str1)
                            'If (StrComp(strTerm1HighConc, "ULOQ", CompareMethod.Text) = 0) Then
                            '20190108 LEE
                            If (InStr(1, strTerm1HighConc, "ULOQ", CompareMethod.Text) > 0) And boolNONELEG = False Then
                                intSuper = intSuper + 1
                                strSuper = ChrW(intSuper)
                                .Selection.TypeText(Text:=ChrW(160))
                                Call typeInSuperscript(wd, strSuper)
                            End If

                            If boolInjCol Then
                                .Selection.Tables.Item(1).Cell(2, intColCt).Select()
                                str1 = "Injection" & ChrW(10) & "#"
                                .Selection.TypeText(Text:=str1)
                                intColCt = intColCt + 1
                            End If

                            .Selection.Tables.Item(1).Cell(2, intColCt).Select()
                            str1 = "Peak" & ChrW(10) & "Area"
                            .Selection.TypeText(Text:=str1)

                            If boolIntStd Then
                                intColCt = intColCt + 1
                                .Selection.Tables.Item(1).Cell(2, intColCt).Select()
                                str1 = "IS Peak" & ChrW(10) & "Area"
                                .Selection.TypeText(Text:=str1)
                            End If


                        End If

                        'do Blanks section

                        intColCt = intColCt + 2
                        .Selection.Tables.Item(1).Cell(1, intColCt).Select()
                        str1 = strBlankLabel
                        .Selection.TypeText(Text:=str1)

                        If boolInjCol Then
                            .Selection.Tables.Item(1).Cell(2, intColCt).Select()
                            str1 = "Injection" & ChrW(10) & "#"
                            .Selection.TypeText(Text:=str1)
                            intColCt = intColCt + 1
                        End If


                        intColBlankPA = intColCt
                        .Selection.Tables.Item(1).Cell(2, intColCt).Select()
                        str1 = "Peak" & ChrW(10) & "Area"
                        .Selection.TypeText(Text:=str1)

                        If boolIntStd Then
                            intColCt = intColCt + 1
                            .Selection.Tables.Item(1).Cell(2, intColCt).Select()
                            str1 = "IS Peak" & ChrW(10) & "Area"
                            .Selection.TypeText(Text:=str1)
                        End If

                        intColCt = intColCt + 2
                        .Selection.Tables.Item(1).Cell(2, intColCt).Select()
                        If boolIntStd Then
                            str1 = "Analyte %" & ChrW(10) & "Carryover" 'NBS
                        Else
                            str1 = "%" & ChrW(10) & "Carryover"
                        End If

                        intSuper = intSuper + 1
                        strSuper = ChrW(intSuper)

                        .Selection.TypeText(Text:=str1)

                        '20190108
                        If boolNONELEG Then
                        Else
                            .Selection.TypeText(Text:=ChrW(160))
                            .Selection.Font.Superscript = True
                            .Selection.TypeText(Text:=strSuper)
                            .Selection.Font.Superscript = False
                        End If

                        If boolIntStd Then
                            intColCt = intColCt + 1
                            'superscript letter is same as previous
                            strSuper = ChrW(intSuper)
                            .Selection.Tables.Item(1).Cell(2, intColCt).Select()
                            str1 = "IS %" & ChrW(10) & "Carryover"
                            .Selection.TypeText(Text:=str1)

                            '20190108
                            If boolNONELEG Then
                            Else
                                .Selection.TypeText(Text:=ChrW(160))
                                .Selection.Font.Superscript = True
                                .Selection.TypeText(Text:=strSuper)
                                .Selection.Font.Superscript = False
                            End If

                        End If

                        ''''wdd.visible = True

                        .Selection.Tables.Item(1).Cell(2, 1).Select()
                        'bottom border this row
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle


                        'merge and border blanks
                        If boolInjCol Then 'find beginning merge range
                            intSS = intColBlankInjNum
                        Else
                            intSS = intColBlankPA
                        End If
                        .Selection.Tables.Item(1).Cell(1, intSS).Select()
                        int1 = 0
                        If boolInjCol Then
                            int1 = int1 + 1
                        End If
                        If boolIntStd Then
                            int1 = int1 + 1
                        End If
                        If int1 = 0 Then
                        Else
                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=int1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                            .Selection.Cells.Merge()
                            With .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                            End With
                        End If

                        If boolIULOQ Then

                            'merge and border ULOQ
                            If boolInjCol Then 'find beginning merge range
                                intSS = intColULOQInjNum
                            Else
                                intSS = intColULOQPA
                            End If
                            .Selection.Tables.Item(1).Cell(1, intSS).Select()
                            int1 = 0
                            If boolInjCol Then
                                int1 = int1 + 1
                            End If
                            If boolIntStd Then
                                int1 = int1 + 1
                            End If
                            If int1 = 0 Then
                            Else
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=int1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                                .Selection.Cells.Merge()
                                With .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                    .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                End With
                            End If

                        End If

                        'merge and border LLOQ
                        If boolInjCol Then 'find beginning merge range
                            intSS = intColLLOQInjNum
                        Else
                            intSS = intColLLOQPA
                        End If

                        .Selection.Tables.Item(1).Cell(1, intSS).Select()
                        int1 = 0
                        If boolInjCol Then
                            int1 = int1 + 1
                        End If
                        If boolIntStd Then
                            int1 = int1 + 1
                        End If
                        If int1 = 0 Then
                        Else
                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=int1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                            .Selection.Cells.Merge()
                            With .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                            End With
                        End If


                        'begin entering data'
                        intStart = 4
                        Dim boolExit As Boolean = False

                        Dim numAnal As Single
                        Dim numIS As Single

                        Dim tblAR As New System.Data.DataTable
                        tblAR.Columns.Add("AR", Type.GetType("System.Decimal"))
                        tblAR.Columns.Add("ARIS", Type.GetType("System.Decimal"))

                        Dim rowsAR() As DataRow

                        frmH.lblProgress.Text = strM1 ' & ChrW(10) & "Processing Run ID " & var10
                        frmH.Refresh()

                        intStart = 4
                        Dim intColStart As Short
                        intColStart = 2
                        Dim intNRow As Single
                        Dim numMeanLLOQ As Single
                        Dim numMeanLLOQIS As Single

                        intNRow = 4
                        intIc = 0
                        int8Max = 0


                        strM = "Creating " & strTName & " For " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & "..."
                        strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                        strM1 = strM
                        frmH.lblProgress.Text = strM & ChrW(10) & ChrW(10) & "1 of " & intNumRuns & " analytical runs..."
                        frmH.Refresh()

                        For Count2 = 0 To intNumRuns - 1

                            intIc = intIc + 1
                            If intIc > intIcMax Then

                                strM = "Creating " & strTName & " For " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & "..."
                                strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                                strM1 = strM
                                frmH.lblProgress.Text = strM & ChrW(10) & ChrW(10) & Count2 & " of " & intNumRuns & " analytical runs..."
                                frmH.Refresh()

                                intIc = 0
                            End If


                            var10 = tblRunID.Rows(Count2).Item("RUNID")

                            intStart = intNRow

                            'Do Quick Runthrough for LLOQ and ULOQ to see how many rows each has
                            Dim intNumRowsInSection As Short = 0
                            For Count4 = 1 To 3
                                Select Case Count4
                                    Case 1
                                        strF = strFData & " AND RUNID = " & var10 & " AND CHARHELPER1 = 'LLOQ'"
                                    Case 2
                                        strF = strFData & " AND RUNID = " & var10 & " AND CHARHELPER1 = '" & strTerm1HighConc & "'"
                                    Case 3
                                        strF = strFData & " AND RUNID = " & var10 & " AND CHARHELPER1 = 'Blank'"
                                End Select

                                Dim tbl2SB As System.Data.DataTable = dv2.ToTable
                                strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                                rows2 = tbl2SB.Select(strF, strS)
                                int3 = rows2.Length
                                If (int3 > intNumRowsInSection) Then
                                    intNumRowsInSection = int3
                                End If
                            Next

                            'Need to find number of Blank reps
                            Dim intBReps As Short
                            strF = strFData & " AND RUNID = " & var10 & " AND CHARHELPER1 = 'Blank'"
                            Dim tbl2BReps As System.Data.DataTable = dv2.ToTable
                            strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                            rows2 = tbl2BReps.Select(strF, strS)
                            intBReps = rows2.Length

                            intLLOQ = 0

                            'Now go through and put in actual values

                            Dim intColInj As Short '
                            Dim intColAPA As Short
                            Dim intColISPA As Short

                            For Count4 = 1 To 3

                                Select Case Count4
                                    Case 1
                                        strF = strFData & " AND RUNID = " & var10 & " AND CHARHELPER1 = 'LLOQ'"
                                        intColInj = intColLLOQInjNum
                                        intColAPA = intColLLOQPA
                                        intColISPA = intColLLOQISPA

                                    Case 2

                                        If boolIULOQ Then
                                        Else
                                            GoTo skipNext4
                                        End If
                                        intColInj = intColULOQInjNum
                                        intColAPA = intColULOQPA
                                        intColISPA = intColULOQISPA
                                        strF = strFData & " AND RUNID = " & var10 & " AND CHARHELPER1 = '" & strTerm1HighConc & "'"

                                    Case 3
                                        intColInj = intColBlankInjNum
                                        intColAPA = intColBlankPA
                                        intColISPA = intColBlankISPA
                                        strF = strFData & " AND RUNID = " & var10 & " AND CHARHELPER1 = 'Blank'"

                                End Select

                                Dim tbl2SB As System.Data.DataTable = dv2.ToTable
                                strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                                rows2 = tbl2SB.Select(strF, strS)
                                int3 = rows2.Length

                                If Count4 = 1 Then
                                    tblAR.Clear()
                                    intLLOQ = int3
                                End If

                                For Count3 = 0 To int3 - 1

                                    If Count3 + intStart > UBound(arrP, 2) Then
                                        ReDim Preserve arrP(intCols, UBound(arrP, 2) + 10)
                                    End If

                                    'enter runid
                                    If Count4 = 1 And Count3 = 0 And intColWRID <> 0 Then

                                        If BOOLINCLUDEDATE Then
                                            str2 = CStr(var10)
                                            str1 = GetDateFromRunID(NZ(var10, 0), LDateFormat, intGroup, idTR)
                                            'str3 = str2 & ChrW(10) & "(" & str1 & ")"
                                            str3 = str2 & strR & "(" & str1 & ")" 'use soft vertical return (11) instead of line feed (10)

                                            '20181129 LEE:
                                            'not consistent with other reports
                                            'try this instead
                                            arrP(intColWRID, Count3 + intStart + 1) = "(" & str1 & ")"

                                        Else
                                            str3 = CStr(var10)
                                        End If
                                        'arrP(intColWRID, Count3 + intStart) = str3
                                        arrP(intColWRID, Count3 + intStart) = CStr(var10)

                                    End If

                                    'enter sample name
                                    If Count4 = 3 And intColBSN <> 0 Then
                                        var1 = NZ(rows2(Count3).Item("SAMPLENAME"), "NA")
                                        If StrComp(var1, "NA", CompareMethod.Text) = 0 Then
                                        Else
                                            'strip off RunID and Seq#
                                            var1 = ShortSampleName(var1.ToString)
                                        End If
                                        arrP(intColBSN, Count3 + intStart) = var1

                                    End If

                                    If boolInjCol Then
                                        'enter injection number
                                        var1 = NZ(rows2(Count3).Item("RUNSAMPLEORDERNUMBER"), "NA")
                                        arrP(intColInj, Count3 + intStart) = var1
                                    End If

                                    'enter Analyte Peak Area
                                    var1 = NZ(rows2(Count3).Item("ANALYTEAREA"), "NA")
                                    If IsNumeric(var1) Then
                                        If boolLUseSigFigsArea Then
                                            var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                        Else
                                            var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                        End If
                                        numAnal = var1
                                    Else
                                        numAnal = 0
                                    End If
                                    arrP(intColAPA, Count3 + intStart) = var1

                                    If Count4 = 3 Then
                                        Dim nrow As DataRow = tblAR.NewRow
                                        nrow.BeginEdit()
                                        nrow.Item("AR") = var1
                                        nrow.EndEdit()
                                        tblAR.Rows.Add(nrow)
                                    End If

                                    If boolIntStd Then

                                        'enter Analyte Peak Area
                                        var1 = NZ(rows2(Count3).Item("INTERNALSTANDARDAREA"), "NA")
                                        If IsNumeric(var1) Then
                                            If boolLUseSigFigsArea Then
                                                var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                            Else
                                                var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                            End If
                                            numIS = var1
                                        Else
                                            numIS = 0
                                        End If
                                        '.Selection.Tables.Item(1).Cell(Count3 + intStart, intColStart + 2).Select()
                                        '.Selection.TypeText(CStr(var1))
                                        arrP(intColISPA, Count3 + intStart) = var1

                                        If Count4 = 3 Then
                                            tblAR.Rows(Count3).BeginEdit()
                                            tblAR.Rows(Count3).Item("ARIS") = var1
                                            tblAR.Rows(Count3).EndEdit()
                                        End If

                                    End If

                                    If Count4 = 3 Then
                                        var1 = tblAR.Rows(Count3).Item("AR")
                                        var2 = numMeanLLOQ
                                        If var2 = 0 Then
                                            'var4 = "NA"
                                            var4 = Format(var2, strQCDec)
                                        Else
                                            var3 = var1 / var2 * 100
                                            var4 = Format(var3, strQCDec)
                                        End If
                                        '.Selection.Tables.Item(1).Cell(Count3 + intStart, intColCO).Select()
                                        '.Selection.TypeText(CStr(var4))
                                        arrP(intColCO, Count3 + intStart) = var4

                                    End If

                                    If boolIntStd Then
                                        If Count4 = 3 Then
                                            var1 = tblAR.Rows(Count3).Item("ARIS")
                                            var2 = numMeanLLOQIS
                                            If var2 = 0 Then
                                                'var4 = "NA"
                                                var4 = Format(var2, strQCDec)
                                            Else
                                                var3 = var1 / var2 * 100
                                                var4 = Format(var3, strQCDec)
                                            End If
                                            '.Selection.Tables.Item(1).Cell(Count3 + intStart, intColISCO).Select()
                                            '.Selection.TypeText(CStr(var4))
                                            arrP(intColISCO, Count3 + intStart) = var4
                                        End If
                                    End If

                                Next Count3


                                Dim intRow As Short
                                'intRow is the position to start next section
                                'the next section is either a new analytical run or a Stats section
                                intRow = intStart + intNumRowsInSection + 1

                                'the following needed for mean calculation
                                Dim tbl2SA As System.Data.DataTable = dv2.ToTable
                                strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                                rows2E = tbl2SA.Select(strF, strS)
                                int3 = rows2E.Length

                                If intCSN = 0 Then

                                    int8 = intRow - 1

                                    intNRow = int8 + 1

                                    '****

                                    If Count4 = 1 Then

                                        numMeanLLOQ = 0
                                        numMeanLLOQIS = 0

                                        Try
                                            'record Mean of Peak Area

                                            'var1 = MeanDRArea(rows2E, "ANALYTEAREA", False, "ALIQUOTFACTOR", False, False)
                                            '20180720 LEE: MeanDR can accept Area
                                            var1 = MeanDR(rows2E, "ANALYTEAREA", False, "ALIQUOTFACTOR", False, False)
                                            If boolLUseSigFigsArea Then  'NDL 14-Jan-2015 Changes this to SigFigsArea, as we are dealing with Peak areas comparisons here.
                                                numMean = SigFigOrDecString(var1, LSigFigArea, True)
                                            Else
                                                numMean = RoundToDecimalRAFZ(var1, LSigFigArea)
                                            End If

                                            Call InsertQCTables(intTableID, idTR, charFCID, NZ(varNom, 0), 1, "Mean", numMean, CSng(var10), Count1, strDo, 0, 0, False)

                                            numMeanLLOQ = numMean

                                            If boolIntStd Then

                                                'var1 = MeanDRArea(rows2E, "INTERNALSTANDARDAREA", False, "ALIQUOTFACTOR", False, False)
                                                '20180720 LEE: MeanDR can accept Area
                                                var1 = MeanDR(rows2E, "INTERNALSTANDARDAREA", False, "ALIQUOTFACTOR", False, False)
                                                If boolLUseSigFigsArea Then  'NDL 14-Jan-2015 Changes this to SigFigsArea, as we are dealing with Peak areas comparisons here.
                                                    numMean = SigFigOrDecString(var1, LSigFigArea, True)
                                                Else
                                                    numMean = RoundToDecimalRAFZ(var1, LSigFigArea)
                                                End If

                                                Call InsertQCTables(intTableID, idTR, charFCID, NZ(varNom, 0), 1, "MeanIS", numMean, CSng(var10), Count1, strDo, 0, 0, False)

                                                numMeanLLOQIS = numMean

                                            End If
                                        Catch ex As Exception

                                        End Try

                                    End If

                                    '****

                                Else

                                    'now enter Mean/Bias/n labels
                                    'int8 = intRow - 1
                                    'int8 = intRow + intBReps - 2

                                    If Count4 = 1 Then

                                        'subtract a row to allow int8 counter ability to count
                                        int8 = intRow - 1

                                        'If intNumRuns = 1 Then
                                        '    If intCSN = 0 Then
                                        '        int8 = intRow + intBReps - 2
                                        '    Else
                                        '        int8 = intRow + intBReps - 3
                                        '    End If

                                        'Else
                                        '    If intCSN = 0 Then
                                        '        int8 = intRow + intBReps - 2
                                        '    Else
                                        '        int8 = intRow + intBReps - 2
                                        '    End If
                                        'End If


                                        If int8 > int8Max Then
                                            int8Max = int8
                                        End If

                                        '.add
                                        int1 = 0
                                        If boolSTATSMEAN Then
                                            int8 = int8 + 1
                                            int1 = int1 + 1
                                            '.Selection.Tables.Item(1).Cell(int8, 1).Select()
                                            '.Selection.TypeText("Mean")
                                            arrP(1, int8) = "Mean"
                                            '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, 1).Select()
                                        End If
                                        If boolSTATSSD Then
                                            int8 = int8 + 1
                                            int1 = int1 + 1
                                            '.Selection.Tables.Item(1).Cell(int8, 1).Select()
                                            '.Selection.TypeText("S.D.") '((Mean/NomConc)-1)*100)
                                            arrP(1, int8) = "S.D."
                                        End If
                                        If boolSTATSCV Then
                                            int8 = int8 + 1
                                            int1 = int1 + 1
                                            '.Selection.Tables.Item(1).Cell(int8, 1).Select()
                                            '.Selection.TypeText(ReturnPrecLabel()) '(sd/mean)*100)
                                            arrP(1, int8) = ReturnPrecLabel()
                                        End If

                                        If boolSTATSN Then
                                            int8 = int8 + 1
                                            int1 = int1 + 1
                                            '.Selection.Tables.Item(1).Cell(int8, 1).Select()
                                            '.Selection.TypeText("n")
                                            arrP(1, int8) = "n"
                                        End If
                                    End If

                                    ''now start enter stats for analyte
                                    'Dim tbl2SA As System.Data.DataTable = dv2.ToTable
                                    'strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                                    'rows2E = tbl2SA.Select(strF, strS)
                                    'int3 = rows2E.Length

                                    'int8 = intRow - 1
                                    'messes up here
                                    If boolSTATSMEAN Or boolSTATSSD Or boolSTATSCV Or boolSTATSN Then
                                        'int8 = intRow + intBReps - int1
                                        'int8 = intRow - int1
                                        int8 = int8 - int1
                                    Else
                                        int8 = intRow + intBReps - 1
                                    End If

                                    '*****

                                    If Count4 = 1 Then

                                        numMeanLLOQ = 0
                                        numMeanLLOQIS = 0

                                        Try
                                            'var1 = MeanDRArea(rows2E, "ANALYTEAREA", False, "ALIQUOTFACTOR", False, False)
                                            '20180720 LEE: MeanDR can accept Area
                                            var1 = MeanDR(rows2E, "ANALYTEAREA", False, "ALIQUOTFACTOR", False, False)
                                            If boolLUseSigFigsArea Then  'NDL 14-Jan-2015 Changes this to SigFigsArea, as we are dealing with Peak areas comparisons here.
                                                numMean = SigFigOrDecString(var1, LSigFigArea, True)
                                            Else
                                                numMean = RoundToDecimalRAFZ(var1, LSigFigArea)
                                            End If

                                            Call InsertQCTables(intTableID, idTR, charFCID, NZ(varNom, 0), 1, "Mean", numMean, CSng(var10), Count1, strDo, 0, 0, False)
                                            If boolLUseSigFigsArea Then
                                                var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                            Else
                                                var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea + 2), strAreaDec)
                                            End If

                                            numMeanLLOQ = numMean

                                            If boolIntStd Then

                                                'var1 = MeanDRArea(rows2E, "INTERNALSTANDARDAREA", False, "ALIQUOTFACTOR", False, False)
                                                '20180720 LEE: MeanDR can accept Area
                                                var1 = MeanDR(rows2E, "INTERNALSTANDARDAREA", False, "ALIQUOTFACTOR", False, False)
                                                If boolLUseSigFigsArea Then  'NDL 14-Jan-2015 Changes this to SigFigsArea, as we are dealing with Peak areas comparisons here.
                                                    numMean = SigFigOrDecString(var1, LSigFigArea, True)
                                                Else
                                                    numMean = RoundToDecimalRAFZ(var1, LSigFigArea)
                                                End If

                                                Call InsertQCTables(intTableID, idTR, charFCID, NZ(varNom, 0), 1, "MeanIS", numMean, CSng(var10), Count1, strDo, 0, 0, False)
                                                If boolLUseSigFigsArea Then
                                                    var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                                Else
                                                    var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea + 2), strAreaDec)
                                                End If

                                                numMeanLLOQIS = numMean

                                            End If
                                        Catch ex As Exception

                                        End Try

                                        If boolSTATSMEAN Then
                                            Try
                                                'enter Mean of Peak Area
                                                int8 = int8 + 1

                                                'var1 = MeanDRArea(rows2E, "ANALYTEAREA", False, "ALIQUOTFACTOR", False, False)
                                                '20180720 LEE: MeanDR can accept Area
                                                var1 = MeanDR(rows2E, "ANALYTEAREA", False, "ALIQUOTFACTOR", False, False)
                                                If boolLUseSigFigsArea Then  'NDL 14-Jan-2015 Changes this to SigFigsArea, as we are dealing with Peak areas comparisons here.
                                                    numMean = SigFigOrDecString(var1, LSigFigArea, True)
                                                Else
                                                    numMean = RoundToDecimalRAFZ(var1, LSigFigArea)
                                                End If

                                                Call InsertQCTables(intTableID, idTR, charFCID, NZ(varNom, 0), 1, "Mean", numMean, CSng(var10), Count1, strDo, 0, 0, False)
                                                If boolLUseSigFigsArea Then
                                                    var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                                Else
                                                    var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea + 2), strAreaDec)
                                                End If
                                                '.Selection.Tables.Item(1).Cell(int8, intColStart + 1).Select()
                                                '.Selection.TypeText(CStr(var1))
                                                arrP(intColLLOQPA, int8) = CStr(var1)
                                                numMeanLLOQ = numMean

                                                If boolIntStd Then

                                                    'var1 = MeanDRArea(rows2E, "INTERNALSTANDARDAREA", False, "ALIQUOTFACTOR", False, False)
                                                    '20180720 LEE: MeanDR can accept Area
                                                    var1 = MeanDR(rows2E, "INTERNALSTANDARDAREA", False, "ALIQUOTFACTOR", False, False)
                                                    If boolLUseSigFigsArea Then  'NDL 14-Jan-2015 Changes this to SigFigsArea, as we are dealing with Peak areas comparisons here.
                                                        numMean = SigFigOrDecString(var1, LSigFigArea, True)
                                                    Else
                                                        numMean = RoundToDecimalRAFZ(var1, LSigFigArea)
                                                    End If

                                                    Call InsertQCTables(intTableID, idTR, charFCID, NZ(varNom, 0), 1, "MeanIS", numMean, CSng(var10), Count1, strDo, 0, 0, False)
                                                    If boolLUseSigFigsArea Then
                                                        var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                                    Else
                                                        var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea + 2), strAreaDec)
                                                    End If
                                                    '.Selection.Tables.Item(1).Cell(int8, intColStart + 2).Select()
                                                    '.Selection.TypeText(CStr(var1))
                                                    'arrP(intColStart + 2, int8) = CStr(var1)
                                                    arrP(intColLLOQISPA, int8) = CStr(var1)
                                                    numMeanLLOQIS = numMean
                                                End If
                                            Catch ex As Exception

                                            End Try

                                        End If


                                        Try
                                            If boolLUseSigFigs Then
                                                numSD = SigFigOrDecString(var1, LSigFig, False)
                                            Else
                                                numSD = RoundToDecimalRAFZ(var1, LSigFig)
                                            End If
                                            Call InsertQCTables(intTableID, idTR, charFCID, NZ(varNom, 0), 1, "SD", numSD, CSng(var10), Count1, strDo, 0, 0, False)
                                            var1 = numSD

                                            If boolIntStd Then

                                                var1 = StdDevDRArea(rows2E, "INTERNALSTANDARDAREA", False, "ALIQUOTFACTOR", False, False)
                                                If boolLUseSigFigs Then
                                                    numSDIS = SigFigOrDecString(var1, LSigFig, True)
                                                Else
                                                    numSDIS = RoundToDecimalRAFZ(var1, LSigFig)
                                                End If
                                                Call InsertQCTables(intTableID, idTR, charFCID, NZ(varNom, 0), 1, "SDIS", numSDIS, CSng(var10), Count1, strDo, 0, 0, False)

                                            End If

                                        Catch ex As Exception

                                        End Try
                                        If boolSTATSSD Then
                                            int8 = int8 + 1
                                            If int3 < gSDMax Then
                                                '.Selection.Tables.Item(1).Cell(int8, intColStart + 1).Select()
                                                '.Selection.TypeText("NA")
                                                var1 = "NA"
                                                arrP(intColLLOQPA, int8) = CStr(var1)
                                                If boolIntStd Then
                                                    '.Selection.Tables.Item(1).Cell(int8, intColStart + 2).Select()
                                                    '.Selection.TypeText("NA")
                                                    var1 = "NA"
                                                    arrP(intColLLOQISPA, int8) = CStr(var1)
                                                End If
                                            Else
                                                Try
                                                    'enter SD of Peak Area

                                                    var1 = StdDevDRArea(rows2E, "ANALYTEAREA", False, "ALIQUOTFACTOR", False, False)
                                                    If boolLUseSigFigs Then
                                                        numSD = SigFigOrDecString(var1, LSigFig, False)
                                                    Else
                                                        numSD = RoundToDecimalRAFZ(var1, LSigFig)
                                                    End If
                                                    var1 = numSD
                                                    If boolLUseSigFigsArea Then
                                                        var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                                    Else
                                                        var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea + 2), strAreaDec)
                                                    End If
                                                    '.Selection.Tables.Item(1).Cell(int8, intColStart + 1).Select()
                                                    '.Selection.TypeText(CStr(var1))
                                                    arrP(intColLLOQPA, int8) = CStr(var1)

                                                    If boolIntStd Then

                                                        var1 = StdDevDRArea(rows2E, "INTERNALSTANDARDAREA", False, "ALIQUOTFACTOR", False, False)
                                                        If boolLUseSigFigs Then
                                                            numSDIS = SigFigOrDecString(var1, LSigFig, True)
                                                        Else
                                                            numSDIS = RoundToDecimalRAFZ(var1, LSigFig)
                                                        End If
                                                        var1 = numSD
                                                        If boolLUseSigFigsArea Then
                                                            var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                                        Else
                                                            var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea + 2), strAreaDec)
                                                        End If
                                                        '.Selection.Tables.Item(1).Cell(int8, intColStart + 2).Select()
                                                        '.Selection.TypeText(CStr(var1))
                                                        arrP(intColLLOQISPA, int8) = CStr(var1)

                                                    End If

                                                Catch ex As Exception

                                                End Try

                                            End If

                                        End If


                                        Try
                                            numPrec = CalcCVPercent(numSD, numMeanLLOQ, intQCDec)
                                            Call InsertQCTables(intTableID, idTR, charFCID, varNom, 1, "Precision", numPrec, CSng(var10), Count1, strDo, 0, 0, False)
                                            If boolIntStd Then
                                                numPrecIS = CalcCVPercent(numSDIS, numMeanLLOQIS, intQCDec)
                                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, 1, "PrecisionIS", numPrecIS, CSng(var10), Count1, strDo, 0, 0, False)
                                            End If
                                        Catch ex As Exception

                                        End Try
                                        If boolSTATSCV Then
                                            Try
                                                'enter %CV
                                                int8 = int8 + 1
                                                If int3 < gSDMax Then
                                                    '.Selection.Tables.Item(1).Cell(int8, intColStart + 1).Select()
                                                    '.Selection.TypeText("NA")
                                                    var1 = "NA"
                                                    arrP(intColLLOQPA, int8) = CStr(var1)
                                                    If boolIntStd Then
                                                        '.Selection.Tables.Item(1).Cell(int8, intColStart + 2).Select()
                                                        '.Selection.TypeText("NA")
                                                        var1 = "NA"
                                                        arrP(intColLLOQISPA, int8) = CStr(var1)
                                                    End If
                                                Else

                                                    Try
                                                        '.Selection.Tables.Item(1).Cell(int8, intColStart + 1).Select()
                                                        '.Selection.TypeText(Format(numPrec, strQCDec))
                                                        var1 = Format(numPrec, strQCDec)
                                                        arrP(intColLLOQPA, int8) = CStr(var1)

                                                        If boolIntStd Then
                                                            '.Selection.Tables.Item(1).Cell(int8, intColStart + 2).Select()
                                                            '.Selection.TypeText(Format(numPrecIS, strQCDec))
                                                            var1 = Format(numPrecIS, strQCDec)
                                                            arrP(intColLLOQISPA, int8) = CStr(var1)
                                                        End If

                                                    Catch ex As Exception

                                                    End Try

                                                End If

                                            Catch ex As Exception

                                            End Try
                                        End If


                                        Try
                                            Call InsertQCTables(intTableID, idTR, charFCID, varNom, 1, "n", numPrec, CSng(var10), Count1, strDo, 0, 0, False)
                                            If boolIntStd Then
                                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, 1, "nIS", numPrec, CSng(var10), Count1, strDo, 0, 0, False)
                                            End If
                                        Catch ex As Exception

                                        End Try
                                        If boolSTATSN Then
                                            Try
                                                'enter n
                                                int8 = int8 + 1
                                                '.Selection.Tables.Item(1).Cell(int8, intColStart + 1).Select()
                                                '.Selection.TypeText(CStr(int3))
                                                arrP(intColLLOQPA, int8) = CStr(int3)

                                                If boolIntStd Then

                                                    '.Selection.Tables.Item(1).Cell(int8, intColStart + 2).Select()
                                                    '.Selection.TypeText(CStr(int3))
                                                    arrP(intColLLOQISPA, int8) = CStr(int3)

                                                End If

                                            Catch ex As Exception

                                            End Try
                                        End If

                                        'intNRow = int8 + 3
                                        If boolSTATSMEAN Then
                                            intNRow = int8 + 2
                                        Else
                                            intNRow = int8 ' - 1
                                        End If

                                    End If 'Here

                                    '*****

                                    If int8 > int8Max Then
                                        int8Max = int8
                                    End If

                                    var1 = var1 'debug

                                End If 'from IF intCSN > 0



                                If int8 > int8Max Then
                                    int8Max = int8
                                End If

skipNext4:

                            Next Count4 '1=LLOQ, 2=ULOQ, 3=BLANK

                        Next Count2 'intNumRuns

                        strPaste = ""
                        strPasteT = ""
                        Dim intSt As Short = 4 'the first row that has data

                        If int8Max > UBound(arrP, 2) Then
                            ReDim Preserve arrP(intCols, int8Max)
                        End If
                        For Count4 = intSt To int8Max  'rows
                            For Count2 = 1 To intCols  'columns
                                var1 = arrP(Count2, Count4)
                                var2 = NZ(var1, "")
                                If Count2 = 1 Then
                                    strPasteT = var2
                                Else
                                    strPasteT = strPasteT & ChrW(9) & var2
                                End If
                            Next
                            If Count4 = intSt Then
                                strPaste = strPasteT
                            Else
                                strPaste = strPaste & ChrW(10) & strPasteT
                            End If
                        Next

                        ''debug
                        'Console.WriteLine("Start")
                        'Console.WriteLine(strPaste)
                        'Console.WriteLine("End")

                        'paste contents

                        'send strpaste to clipboard
                        If IsNothing(strPaste) Then
                        Else
                            Dim rng1 As Word.Range
                            Dim tblW As Word.Table

                            tblW = .Selection.Tables.Item(1)
                            Try
                                rng1 = wd.ActiveDocument.Range(Start:=tblW.Cell(4, 1).Range.Start, End:=tblW.Cell(tblW.Rows.Count, tblW.Columns.Count).Range.End)
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
                                var1 = ex.Message
                            End Try

                            'select appropriate rows
                            rng1.Select()
                            'paste from clipboard
                            Try
                                .Selection.PasteAndFormat(Microsoft.Office.Interop.Word.WdRecoveryType.wdPasteDefault)
                            Catch ex As Exception
                                var1 = ex.Message
                            End Try

                            'the paste action removes the range object and any table formatting, must reset it
                            Call GlobalTableParaFormat(wd)
                            rng1 = wd.ActiveDocument.Range(Start:=tblW.Cell(4, 1).Range.Start, End:=tblW.Cell(tblW.Rows.Count, tblW.Columns.Count).Range.End)
                            rng1.Select()
                            'the paste action removes paragraph formatting, must format again
                            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

                            Call GlobalTableParaFormat(wd)

                            ''Enforce minimal 5-point spacing in Table
                            ''.Selection.Tables.Item(1).Select()
                            ''20160218 LEE: not whole table, just table body, not headers
                            ''the table body is currently selected
                            'Call EnforceMinimumTableVerticalSpacing(wd, 5)

                            '20171220 LEE: Do not set table size, use the style default table
                            '.Selection.Font.Size = fontsize - 1

                            '20160219 LEE
                            'replace '_xyz_' with chrw(11) vertical tab

                            With rng1.Find
                                .ClearFormatting()
                                .Text = strR
                                .Replacement.ClearFormatting()
                                .Replacement.Text = ChrW(11)
                                .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Forward:=True, Wrap:=Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue)
                            End With

                            If BOOLINCLUDEDATE Then
                                '20180711 LEE:
                                'must format this column as wraptext=false
                                'select column
                                .Selection.Tables.Item(1).Cell(1, 1).Select()
                                .Selection.SelectColumn()
                                Call DoCells(.Selection.Cells)
                                var1 = var1

                            End If

                            If boolBSNCol Then
                                'must format this column as wraptext=false
                                'select column
                                .Selection.Tables.Item(1).Cell(1, intColBSN).Select()
                                .Selection.SelectColumn()
                                Call DoCells(.Selection.Cells)
                                var1 = var1
                            End If

                            '20180712 LEE:
                            'need to do autofit here to establish wordwraps
                            If BOOLINCLUDEDATE Or boolBSNCol Then
                                Call AutoFitTable(wd, True)
                            End If

                            '*****

                        End If
                        'end paste contents


                        strM = "Creating " & strTName & " For " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & "..."
                        strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                        strM1 = strM
                        frmH.lblProgress.Text = strM & ChrW(10) & ChrW(10) & intNumRuns & " of " & intNumRuns & " analytical runs..."
                        frmH.Refresh()

                        strM = "Creating " & strTName & " For " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & "..."
                        strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                        strM1 = strM
                        frmH.lblProgress.Text = strM & ChrW(10) & ChrW(10) & "Final formatting..."
                        frmH.Refresh()

                        'END NEW STUFF
                        .Selection.Tables.Item(1).Cell(.Selection.Tables.Item(1).Rows.Count, 1).Select()

                        'bottom border this row
                        .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.MoveRight Unit:=Microsoft.Office.Interop.Word.wdunits.word.wdunits.wdCharacter, Count:=ctQCs, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                        'autofit window
                        'dim bool
                        .Selection.Tables.Item(1).Select()
                        ''select body rows
                        'Dim rngS As Word.Range
                        'rngS = .Selection.Tables.Item(1).Rows(3).Range
                        'rngS.End = .Selection.Tables.Item(1).Rows(.Selection.Tables.Item(1).Rows.Count).Range.End
                        'rngS.Select()


                        ''''wdd.visible = True

                        ''now force fit column 5
                        '.Selection.Tables.Item(1).Cell(3, 5).Select()
                        '.Selection.SelectColumn()
                        '.Selection.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.wdpreferredwidthtype.wdPreferredWidthPercent ' wdPreferredWidthPercent
                        '.Selection.Columns.PreferredWidth = 1
                        ''pesky. do it again
                        '.Selection.Tables.Item(1).Cell(3, 5).Select()
                        '.Selection.SelectColumn()
                        '.Selection.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.wdpreferredwidthtype.wdPreferredWidthPercent ' wdPreferredWidthPercent
                        '.Selection.Columns.PreferredWidth = 1


                        ''''wdd.visible = True

                        .Selection.Tables.Item(1).Cell(1, 1).Select()

                        'remove unused rows
                        Call RemoveRows(wd, ctTbl)

                    Catch ex As Exception

                        str1 = "There was a problem preparing table:"
                        str1 = strM1 & ChrW(10) & ChrW(10) & str1
                        str1 = str1 & ChrW(10) & ChrW(10)
                        str1 = str1 & ex.Message
                        MsgBox(str1, vbInformation, "Problem...")

                    End Try

                    .Selection.Tables.Item(1).Cell(1, 1).Select()

                    'str1 = str2 & " Final Extract Stability: Summary of " & rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION") & " Interpolated QC Standard Concentrations."

                    '***
                    strA = rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")
                    strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, False)
                    Call EnterTableNumber(wd, strTName, 3, strA, strTempInfo, intTableID, intGroup, idTR)
                    '***
                    'Call EnterTableNumber(wd, str1, 3)

                    ''now force fit column 4
                    '.Selection.Tables.Item(1).Cell(intStart, 4).Select()
                    '.Selection.SelectColumn()
                    '.Selection.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent ' wdPreferredWidthPercent
                    '.Selection.Columns.PreferredWidth = 1

                    ''now force fit column 7
                    '.Selection.Tables.Item(1).Cell(intStart, 7).Select()
                    '.Selection.SelectColumn()
                    '.Selection.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent ' wdPreferredWidthPercent
                    '.Selection.Columns.PreferredWidth = 1


                    If boolIntStd And boolHasSampleName Then
                        With .Selection.Tables.Item(1)
                            '.TopPadding = .TopPadding
                            '.BottomPadding = .BottomPadding
                            .LeftPadding = 0
                            .RightPadding = 1.44
                            '.WordWrap = True
                            '.FitText = False
                            '.Spacing = 1
                            'MsgBox(.Spacing())
                        End With
                    End If

                    'enter a table record in tblTableN
                    'ctTableN = ctTableN + 1
                    Dim dtblr As DataRow = tblTableN.NewRow
                    dtblr.BeginEdit()
                    dtblr.Item("TableNumber") = ctTableN
                    dtblr.Item("AnalyteName") = strDo 'arrAnalytes(1, Count1)
                    dtblr.Item("TableName") = strTNameO
                    dtblr.Item("TableID") = intTableID
                    dtblr.Item("CHARFCID") = charFCID
                    dtblr.Item("TableNameNew") = strTName
                    tblTableN.Rows.Add(dtblr)

                    'split table, if needed
                    str1 = frmH.lblProgress.Text

                    intSuper = 96

                    ctLegend = ctLegend + 1
                    intLeg = intLeg + 1
                    arrLegend(1, intLeg) = "NA"
                    arrLegend(2, intLeg) = "Not Applicable"
                    arrLegend(3, intLeg) = False

                    '20190108 LEE:
                    If boolNONELEG Then
                    Else
                        ctLegend = ctLegend + 1
                        intLeg = intLeg + 1
                        intSuper = intSuper + 1
                        strSuper = ChrW(intSuper)
                        arrLegend(1, intLeg) = strSuper
                        arrLegend(2, intLeg) = "Lower Limit of Quantitation"
                        arrLegend(3, intLeg) = True
                        arrLegend(4, intLeg) = True

                        If boolIULOQ Then
                            If (InStr(1, strTerm1HighConc, "ULOQ", CompareMethod.Text) > 0) Then
                                ctLegend = ctLegend + 1
                                intLeg = intLeg + 1
                                intSuper = intSuper + 1
                                strSuper = ChrW(intSuper)
                                arrLegend(1, intLeg) = strSuper
                                arrLegend(2, intLeg) = "Upper Limit of Quantitation"
                                arrLegend(3, intLeg) = True
                                arrLegend(4, intLeg) = True
                            End If
                        End If

                        ctLegend = ctLegend + 1
                        intLeg = intLeg + 1
                        intSuper = intSuper + 1
                        strSuper = ChrW(intSuper)
                        arrLegend(1, intLeg) = strSuper
                        If boolSTATSMEAN Or intLLOQ > 1 Then
                            str1 = "The ratio of Peak Area in " & strBlankLabel & " to the mean LLOQ Peak Area x 100"
                            str1 = "(" & strBlankLabel & " Peak Area)/(Mean LLOQ Peak Area) x 100"
                        Else
                            If intLLOQ = 1 Then
                                str1 = "The ratio of Peak Area in " & strBlankLabel & " to the LLOQ Peak Area x 100"
                                str1 = "(" & strBlankLabel & " Peak Area)/(LLOQ Peak Area) x 100"
                            Else
                                str1 = "The ratio of Peak Area in " & strBlankLabel & " to the mean LLOQ Peak Area x 100"
                                str1 = "(" & strBlankLabel & " Peak Area)/(Mean LLOQ Peak Area) x 100"
                            End If
                        End If
                        arrLegend(2, intLeg) = str1
                        arrLegend(3, intLeg) = True
                        arrLegend(4, intLeg) = True

                        ctLegend = ctLegend + 1
                        intLeg = intLeg + 1
                        intSuper = intSuper + 1
                        strSuper = ChrW(intSuper)
                        arrLegend(1, intLeg) = strSuper
                        str1 = "CF is expressed as (Peak Area of " & strBlankLabel & " Sample)/(Peak Area of the Preceding ULOQ)"
                        arrLegend(2, intLeg) = str1 ' "CF is expressed as the ratio of Peak Area of Carryover assessment Blank sample to Peak Area of the preceding ULOQ"
                        arrLegend(3, intLeg) = True
                        arrLegend(4, intLeg) = False

                    End If



                    'If boolIntStd Or boolHasSampleName Or BOOLINCLUDEDATE Then
                    '    Call AutoFitTable(wd, True)
                    'Else
                    '    Call AutoFitTable(wd, False)
                    'End If

                    '20180712 LEE:
                    'Do TRUE all the time
                    Call AutoFitTable(wd, True)

                    strM = "Finalizing " & strTName & "..."
                    strM1 = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    str1 = strM1

                    frmH.lblProgress.Text = strM1
                    frmH.Refresh()

                    '*****
                    '20180712 LEE:
                    'This table has problems because of the .wordwrap = false stuff for some of the columns in combination with autofit
                    'SplitTable shows two different non-reproduceable errors:
                    '   - Table splits two early resulting in too many pages because column entries are still wrapping
                    '   - Table row ranges somehow get screwed up resulting in legends being totally screwed up
                    'None of these things happen in verbose mode
                    'The resolution is to make wd = visible for the entire SplitTable action

                    wd.Visible = True

                    'autofit again
                    Call AutoFitTable(wd, True)

                    '20180719 LEE:
                    'Autofit is not working for 2nd table if document is large and samplename is long
                    'At this point document is visible
                    'try waiting 5 seconds
                    Pause(5)
                    'try saving the document
                    wd.ActiveDocument.Save()
                    'autfit again
                    Call AutoFitTable(wd, True)
                    'wait another 5 seconds
                    Pause(5)

                    '****

                    str1 = frmH.lblProgress.Text

                    Call SplitTable(wd, 4, intLeg, arrLegend, str1, False, ctLegend + 2, False, False, False, intTableID)
                    'Sub SplitTable(ByVal wd As Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, 
                    'ByVal arr As Object, ByVal strT As String, ByVal DoLegend As Boolean, ByVal intSplitRows As Short, ByVal boolSmallFont As Boolean)

                    'autofit window again
                    .Selection.Tables.Item(1).Select()
                    'autofit table
                    If (boolIntStd And boolHasSampleName) Or BOOLINCLUDEDATE Then
                        Call AutoFitTable(wd, True)
                    Else
                        Call AutoFitTable(wd, False)
                    End If

                    wd.Visible = boolV

                    'pesky
                    'bottom-border row1
                    .Selection.Tables.Item(1).Cell(1, 1).Select()
                    .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                    '.Selection.MoveRight Unit:=Microsoft.Office.Interop.Word.wdunits.word.wdunits.wdCharacter, Count:=ctQCs, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
                    .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                    'move to line below table
                    Call MoveOneCellDown(wd)

                    Call InsertLegend(wd, intTableID, idTR, False, 1)


                Else
                    boolJustTable = False

                End If

end1:
                If boolJustTable Then
                    var1 = strTempInfo
                    'replace numeric with verbose in strTempInfo
                    'look for a sequence of characters that is numeric
                    var3 = ""
                    var4 = ""
                    Dim bool1 As Boolean
                    Dim bool2 As Boolean
                    bool1 = False 'Start
                    bool2 = False 'End
                    str2 = ""
                    For Count2 = 1 To Len(var1)
                        var2 = Mid(var1, Count2, 1)
                        If StrComp(var2, " ", CompareMethod.Text) = 0 Then
                            var2 = "a"
                        End If
                        If IsNumeric(var2) Then
                            var3 = var3 & var2
                            If IsNumeric(var3) Then
                                var4 = var3
                                bool1 = True
                            Else
                            End If
                        Else
                            If bool1 Then
                                bool2 = True
                            End If
                        End If
                        If bool1 And bool2 Then
                            Exit For
                        End If
                    Next

                    If bool1 = False Then
                        var2 = "[NA]"
                    Else
                        var2 = VerboseNumber(var4, True)
                        str2 = Replace(var1, CStr(var4), var2, 1, 1, CompareMethod.Text)
                    End If

                    str1 = str2 & " " & strTName ' " Final Extract Stability: Summary of " & rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION") & " Interpolated QC Standard Concentrations."
                    str2 = str1
                    'str1 = rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")
                    ''Call JustTable(wd, str1, str2, strDo, strTName, intTableID)
                    'Call JustTable(wd, str1, strTName, strDo, strTName, intTableID, strTempInfo, "")

                    str1 = NZ(rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION"), "")
                    'Call JustTable(wd, str1, str2, strDo, strTName, intTableID)
                    If Len(str1) = 0 Then
                    Else
                        strA = str1
                        strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, False)
                        Call JustTable(wd, str1, strTName, strDo, strTName, intTableID, strTempInfo, "", strTNameO, intGroup, idTR)
                    End If

                End If

next1:

            Next
end2:
        End With

        wd.Visible = boolV

    End Sub

    Sub MVSelectivity_v1_34(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal idTR As Int64)

        Dim boolOC As Boolean = False 'bool if eliminated
        Dim numNomConc As Decimal
        Dim var1, var2, var3, var4, var5, var10
        Dim dvDo As system.data.dataview
        Dim strTName As String
        Dim intDo As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim Count4 As Short
        Dim Count5 As Short
        Dim Count10 As Short
        Dim strDo As String
        Dim bool As Boolean
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim tbl1 As System.Data.DataTable
        Dim dv1 As system.data.dataview
        Dim rows1() As DataRow
        Dim intRows1 As Short
        Dim strF1 As String
        Dim tbl2 As System.Data.DataTable
        Dim dv2 As system.data.dataview
        Dim rows2() As DataRow
        Dim intRows2 As Short
        Dim strF2 As String
        Dim tbl3 As System.Data.DataTable
        Dim dv3 As system.data.dataview
        Dim rows3() As DataRow
        Dim intRows3 As Short
        Dim strF3 As String
        Dim intTableID As Short
        Dim tbl4 As System.Data.DataTable
        Dim dv4 As system.data.dataview
        Dim rows4() As DataRow
        Dim intRows4 As Short
        Dim strF4 As String
        Dim strS As String
        Dim intNumRuns As Short
        Dim dv As system.data.dataview
        Dim tblNumRuns As System.Data.DataTable
        Dim tblLevels As System.Data.DataTable
        Dim intNumLevels As Short
        Dim intTblRows As Short
        Dim wrdSelection As Microsoft.Office.Interop.Word.selection
        Dim strF As String
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim int4 As Short
        Dim int10 As Short
        Dim intRowsX As Short
        Dim tblX As System.Data.DataTable
        Dim varNom
        Dim strConcUnits As String
        Dim intLeg As Short
        Dim ctQCLegend As Short
        Dim ctDilLeg As Short
        Dim strA As String
        Dim strB As String

        Dim ctLegend As Short
        Dim fontsize
        Dim boolPro As Boolean

        Dim hi, lo
        Dim rows10() As DataRow
        Dim rows11() As DataRow
        Dim intRowsAnal As Short
        Dim arrFP(2, 20) 'FlagPercent array: 1=hi, 2=lo
        Dim strFP As String
        Dim numMean As Decimal
        Dim numBias As Decimal
        Dim numSD As Decimal
        Dim tblZ As System.Data.DataTable
        Dim dvAn As system.data.dataview
        Dim tblAnGo As New System.Data.DataTable
        Dim p1, p2, p3, p4, p5, p6, p7, p8, p9, p10
        Dim strM As String
        Dim fonts
        Dim numDF As Decimal
        Dim DilFactor
        Dim strF2a As String
        Dim strTempInfo As String
        Dim rowsX() As DataRow
        Dim intLegStart As Short
        Dim boolJustTable As Boolean

        Dim intExp As Short
        Dim ctExp As Short
        Dim int8 As Short
        Dim intStart As Short

        Dim rows2E() As DataRow
        Dim nE As Short
        Dim nI As Short
        Dim boolOutHeadE As Boolean = False
        Dim boolOutHeadI As Boolean = False
        Dim boolDeleteRows As Boolean = False

        Dim tblRID As System.Data.DataTable
        Dim numRID As Short
        Dim ctTbl As Short
        Dim varAnal, varIS
        Dim v1, v2, vU

        Dim vAnalyteIndex
        Dim vMasterAssayID
        Dim vAnalyteID
        Dim tblAG As DataTable = tblAnalyteGroups 'tblAnalyteGroups has all analytes, not just accepted

        Dim intGroup As Short
        Dim strAnal As String
        Dim strAnalC As String
        Dim strMatrix As String
        Dim strTNameO As String
        Dim intRunID As Int16
        Dim strDECISIONREASON As String
        Dim boolExFromAS As Boolean

        Dim numPrec As Single
        Dim numTheor As Single
        Dim rowsData() As DataRow

        Dim charFCID As String
        strF = "ID_TBLREPORTTABLE = " & idTR
        Dim rowsTR() As DataRow = tblReportTable.Select(strF)
        var1 = rowsTR(0).Item("CHARFCID")
        charFCID = NZ(var1, "NA")

        boolJustTable = False

        Cursor.Current = Cursors.WaitCursor

        ''''wdd.visible = True

        fontsize = wd.ActiveDocument.Styles("Normal").Font.Size ' wd.Selection.Font.Size
        fonts = fontsize ' wd.Selection.Font.Size

        With wd

            intTableID = 34

            Dim strWRunId As String = GetWatsonColH(intTableID)

            dvDo = frmH.dgvReportTableConfiguration.DataSource
            strF = "id_tblconfigreporttables = " & intTableID
            intDo = FindRowDVNumByCol(intTableID, dvDo, "id_tblconfigreporttables")

            '***
            intDo = FindRowDVNumByCol(idTR, dvDo, "ID_TBLREPORTTABLE")

            var1 = dvDo(intDo).Item("CHARHEADINGTEXT")
            strTName = NZ(var1, "[NONE]")
            strTNameO = strTName

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

            tbl1 = tblAnalysisResultsHome
            tbl2 = tblAssignedSamples
            tbl3 = tblAssignedSamplesHelper
            tbl4 = tblAnalytesHome

            Dim tbl5 As System.Data.DataTable
            tbl5 = tblTableProperties
            strF = "ID_TBLREPORTTABLE = " & idTR

            'ensure data has been entered
            strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idTR
            rowsX = tbl2.Select(strF)


            strF = "IsIntStd = 'No'"
            strS = "INTORDER ASC, IsIntStd ASC, AnalyteDescription ASC"
            rows11 = tblAnalytesHome.Select(strF, strS)
            intRowsAnal = rows11.Length

            ctTbl = 0
            For Count1 = 1 To intRowsAnal

                boolJustTable = False

                Dim arrLegend(4, 20)

                strTName = strTNameO

                ctLegend = 0

                Dim int11 As Short
                If boolSTATSDIFFCOL Then
                    int11 = 2
                Else
                    int11 = 1
                End If

                'for legend stuff

                intExp = 0
                ctExp = 0

                intLeg = 0
                ctQCLegend = 0
                ctDilLeg = 0
                ctLegend = 0
                strA = ""
                strB = ""
                arrLegend.Clear(arrLegend, 0, arrLegend.Length)
                arrFP.Clear(arrFP, 0, arrFP.Length)
                intLegStart = 96


                'check if table is to be generated
                'strDo = arrAnalytes(1, Count1) 'record column name
                strDo = rows11(Count1 - 1).Item("ANALYTEDESCRIPTION")

                If UseAnalyte(CStr(strDo)) Then
                Else
                    GoTo next1
                End If

                bool = dvDo.Item(intDo).Item(strDo) 'find boolean value of dvDo column
                strAnal = strDo

                Dim strM1 As String
                If bool Then 'continue

                    intTCur = intTCur + 1

                    ctTbl = ctTbl + 1

                    'setup tables
                    If boolUseGroups Then
                        intGroup = tblAG.Rows(Count1 - 1).Item("INTGROUP")
                        strAnal = tblAG.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                        strAnalC = tblAG.Rows(Count1 - 1).Item("ANALYTEDESCRIPTION_C")
                        vAnalyteID = tblAG.Rows.Item(Count1 - 1).Item("ANALYTEID")
                        strMatrix = tblAG.Rows(Count1 - 1).Item("MATRIX")
                    Else
                        var1 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteIndex")
                        var2 = tbl4.Rows.Item(Count1 - 1).Item("MasterAssayID")
                        var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                        var4 = tbl4.Rows.Item(Count1 - 1).Item("ANALYTEID")
                        intGroup = 0
                        vAnalyteIndex = var1
                        vMasterAssayID = var2
                        vAnalyteID = var4
                        strMatrix = ""
                    End If

                    'ensure data has been entered
                    strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & CleanText(strDo) & "' AND ID_TBLREPORTTABLE = " & idTR
                    rowsX = tbl2.Select(strF)
                    If rowsX.Length = 0 Then
                        'strM = "Creating Summary of " & strTempInfo & " Final Extract Stability Table ...."
                        strM = "Creating " & strTName & " For " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & "..."
                        strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                        frmH.lblProgress.Text = strM
                        frmH.Refresh()
                        'MsgBox("Samples have not been assigned to this table.", MsgBoxStyle.Information, "Samples have not been assigned...")
                        boolJustTable = True
                        'page setup according to configuration
                        str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                        'insert page break
                        'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                        Call InsertPageBreak(wd)
                        Call PageSetup(wd, str1) 'L=Landscape, P=Portrait
                        GoTo end1
                    Else
                        boolJustTable = False
                    End If

                    'page setup according to configuration
                    str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                    'insert page break
                    'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                    Call InsertPageBreak(wd)
                    Call PageSetup(wd, str1) 'L=Landscape, P=Portrait

                    'ReDim arrBCQCs(8, 50) '1=LevelNumber, 2=NomConcentration, 3=ID, 4=FlagPercent, 5=Hello, 6=Lo, 7=#ofReplicates, 8=ASSAYID
                    strM = "Creating Summary of " & strTempInfo & " Final Extract Stability Table For " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & "..."
                    strM = "Creating " & strTName & " For " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & "..."
                    strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    strM1 = strM
                    frmH.lblProgress.Text = strM
                    frmH.Refresh()

                    ''setup tables
                    'If boolUseGroups Then
                    '    intGroup = tblAG.Rows(Count1 - 1).Item("INTGROUP")
                    '    strAnal = tblAG.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                    '    strAnalC = tblAG.Rows(Count1 - 1).Item("ANALYTEDESCRIPTION_C")
                    '    vAnalyteID = tblAG.Rows.Item(Count1 - 1).Item("ANALYTEID")
                    '    strMatrix = tblAG.Rows(Count1 - 1).Item("MATRIX")
                    'Else
                    '    var1 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteIndex")
                    '    var2 = tbl4.Rows.Item(Count1 - 1).Item("MasterAssayID")
                    '    var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                    '    var4 = tbl4.Rows.Item(Count1 - 1).Item("ANALYTEID")
                    '    intGroup = 0
                    '    vAnalyteIndex = var1
                    '    vMasterAssayID = var2
                    '    vAnalyteID = var4
                    '    strMatrix = ""
                    'End If

                    ''setup tables
                    'var1 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteIndex")
                    'var2 = tbl4.Rows.Item(Count1 - 1).Item("MasterAssayID")
                    'var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                    'var4 = tbl4.Rows.Item(Count1 - 1).Item("ANALYTEID")
                    'strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                    'strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                    'strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                    'strF2 = strF2 & "ANALYTEINDEX = " & var1 & " AND "
                    'strF2 = strF2 & "ANALYTEID = " & var4 & " AND "
                    'strF2 = strF2 & "MASTERASSAYID = " & var2 ' & " AND "

                    If boolUseGroups Then
                        strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                        strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                        strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                        strF2 = strF2 & "INTGROUP = " & intGroup
                    Else
                        strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                        strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                        strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                        strF2 = strF2 & "ANALYTEINDEX = " & var1 & " AND "
                        strF2 = strF2 & "MASTERASSAYID = " & var2 ' & " AND "
                        'strF2 = strF2 & "ANALYTEID = " & var4 ' & "' AND "
                        'strF2 = strF2 & "BOOLINTSTD = 0"
                    End If
                    strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                    rows2 = tbl2.Select(strF2, strS)

                    Dim strFData As String
                    strFData = strF2

                    int1 = rows2.Length 'debug
                    dv2 = New DataView(tbl2, strF2, strS, DataViewRowState.CurrentRows)
                    Dim intProcRows As Short
                    intProcRows = dv2.Count

                    'get strConcUnits
                    int1 = FindRowDV("ULOQ Units", frmH.dgvWatsonAnalRef.DataSource)

                    strConcUnits = NZ(frmH.dgvWatsonAnalRef(Count1, int1).Value, "ng/mL")

                    int1 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
                    str1 = NZ(frmH.dgvStudyConfig(1, int1).Value, "")

                    If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
                    Else
                        strConcUnits = str1
                    End If

                    'find number of table rows to generate
                    intRowsX = 0

                    boolOutHeadE = False
                    boolOutHeadI = False
                    boolDeleteRows = False

                    'generate table
                    intTblRows = 0
                    intTblRows = intTblRows + 3 'for Blank with IS header
                    intTblRows = intTblRows + 1 'for blank row
                    intTblRows = intTblRows + intProcRows 'for number of data rows
                    intTblRows = intTblRows + 2 'for two blank rows
                    intTblRows = intTblRows + 3 'for Blank with IS header
                    intTblRows = intTblRows + 1 'for blank row

                    'Increment for Statistics Sections
                    Dim intCSN As Short
                    intCSN = countNumStatsRows()
                    intTblRows = intTblRows + intCSN

                    wrdSelection = wd.Selection()

                    Dim intCols As Short
                    intCols = 11



                    Try

                        '20180913 LEE:
                        Call IncrNextTableNumber(wd)

                        If boolPlaceHolder Then
                            .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=1, NumColumns:=1, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                        Else
                            .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=intTblRows, NumColumns:=intCols, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                        End If

                        .Selection.Tables.Item(1).Select()

                        Call SetCellPaddingZero(.Selection.Tables.Item(1))

                        .Selection.Tables.Item(1).Rows.AllowBreakAcrossPages = False
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                        .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
                        .Selection.Tables.Item(1).Columns.PreferredWidth = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPoints
                        '.Selection.Tables.Item(1).Columns.Item(1).Width = 86
                        For Count2 = 1 To intNumLevels
                            '.Selection.Tables.item(1).Columns.item(Count2 + 1).Width = 50
                        Next
                        .Selection.Tables.Item(1).Select()


                        ''''wdd.visible = True

                        'remove border, but leave top and bottom
                        removeBorderButLeaveTopAndBottom(wd)

                        'border top and bottom of range
                        .Selection.Tables.Item(1).Cell(1, 1).Select()

                        If boolPlaceHolder Then

                            .Selection.Tables.Item(1).Select()
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone

                            strA = arrAnalytes(14, Count1)
                            strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, False)
                            Call EnterTableNumber(wd, strTName, 3, strA, strTempInfo, intTableID, intGroup, idTR)
                            Call MoveOneCellDown(wd)

                            .Selection.TypeParagraph()
                            .Selection.TypeParagraph()

                            'enter a table record in tblTableN
                            'ctTableN = ctTableN + 1
                            Dim dtblr1 As DataRow = tblTableN.NewRow
                            dtblr1.BeginEdit()
                            dtblr1.Item("TableNumber") = ctTableN
                            dtblr1.Item("AnalyteName") = arrAnalytes(1, Count1)
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

                        'enter 1st Row Headers for Blank with IS

                        Dim rowsW() As DataRow
                        Dim intRowsW As Short
                        Dim rowsWO() As DataRow
                        Dim intRowsWO As Short
                        Dim rowsLLOQ() As DataRow
                        Dim intRowsLLOQ As Short

                        Dim tblW As System.Data.DataTable = dv2.ToTable
                        Dim tblWO As System.Data.DataTable = dv2.ToTable
                        Dim tblLLOQ As System.Data.DataTable = dv2.ToTable

                        strF = strFData & " AND CHARHELPER2 = 'Blank With IS'"
                        strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                        rowsW = tblW.Select(strF, strS)
                        intRowsW = rowsW.Length

                        strF = strFData & " AND CHARHELPER2 = 'Blank WithOut IS'"
                        strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                        rowsWO = tblWO.Select(strF, strS)
                        intRowsWO = rowsWO.Length

                        strF = strFData & " AND CHARHELPER2 = 'LLOQ'"
                        strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                        rowsLLOQ = tblLLOQ.Select(strF, strS)
                        intRowsLLOQ = rowsLLOQ.Length

                        Dim varNomConc
                        Try
                            varNomConc = rowsLLOQ(0).Item("NOMCONC")
                        Catch ex As Exception
                            varNomConc = 0
                        End Try

                        Dim strUnits As String
                        Dim dvWA = frmH.dgvWatsonAnalRef.DataSource
                        int2 = FindRowDV("LLOQ Units", dvWA)
                        strUnits = NZ(dvWA(int2).Item(1), "NA")

                        int1 = InStr(strWRunId, " ", CompareMethod.Text)
                        If int1 = 0 Then
                            str2 = strWRunId
                        Else
                            str1 = Mid(strWRunId, 1, int1 - 1)
                            str2 = Mid(strWRunId, int1 + 1, Len(strWRunId))
                        End If


                        'Blank With IS
                        If int1 = 0 Then
                        Else
                            .Selection.Tables.Item(1).Cell(2, 1).Select()
                            'str1 = "Watson"
                            .Selection.TypeText(Text:=str1)
                        End If

                        .Selection.Tables.Item(1).Cell(3, 1).Select()
                        str1 = str2 '"Run ID"
                        .Selection.TypeText(Text:=str1)

                        .Selection.Tables.Item(1).Cell(3, 2).Select()
                        str1 = "Lot #"
                        .Selection.TypeText(Text:=str1)

                        .Selection.Tables.Item(1).Cell(1, 1).Select()
                        str1 = "Blank With Int. Std. Samples"
                        .Selection.TypeText(Text:=str1)



                        .Selection.Tables.Item(1).Cell(2, 3).Select()
                        str1 = "Peak" & ChrW(160) & "Area"
                        .Selection.TypeText(Text:=str1)

                        .Selection.Tables.Item(1).Cell(3, 3).Select()
                        '20181014 LEE
                        'change to 'Analyte'
                        str1 = "Analyte" 'strAnal
                        .Selection.TypeText(Text:=str1)


                        .Selection.Tables.Item(1).Cell(3, 4).Select()
                        str1 = "Int." & ChrW(160) & "Std."
                        .Selection.TypeText(Text:=str1)

                        '.Selection.Tables.Item(1).Cell(3, 5).Select()
                        'str1 = "Peak Area Ratio"
                        '.Selection.TypeText(Text:=str1)

                        .Selection.Tables.Item(1).Cell(2, 5).Select()
                        str1 = "Analyte/IS"
                        .Selection.TypeText(Text:=str1)

                        .Selection.Tables.Item(1).Cell(3, 5).Select()
                        str1 = "Area" & ChrW(160) & "Ratio"
                        .Selection.TypeText(Text:=str1)

                        'LLOQ Samples
                        .Selection.Tables.Item(1).Cell(1, 7).Select()
                        str1 = "LLOQ Samples (" & varNomConc & " " & strUnits & ")"
                        .Selection.TypeText(Text:=str1)

                        int1 = InStr(strWRunId, " ", CompareMethod.Text)
                        If int1 = 0 Then
                            str2 = strWRunId
                        Else
                            str1 = Mid(strWRunId, 1, int1 - 1)
                            str2 = Mid(strWRunId, int1 + 1, Len(strWRunId))
                        End If

                        If int1 = 0 Then
                        Else
                            .Selection.Tables.Item(1).Cell(2, 7).Select()
                            str1 = str1 ' "Watson"
                            .Selection.TypeText(Text:=str1)
                        End If

                        .Selection.Tables.Item(1).Cell(3, 7).Select()

                        If BOOLINCLUDEDATE Then
                            'str1 = str2 & ChrW(10) & "(Analysis Date)"
                            '20180420 LEE:
                            str1 = str2 & ChrW(10) & "(" & GetAnalysisDateLabel(intTableID) & ")"

                        Else
                            str1 = str2
                        End If

                        .Selection.TypeText(Text:=str1)

                        .Selection.Tables.Item(1).Cell(2, 8).Select()
                        str1 = "Peak" & ChrW(160) & "Area"
                        .Selection.TypeText(Text:=str1)

                        .Selection.Tables.Item(1).Cell(3, 8).Select()
                        '20181014 LEE
                        'change to 'Analyte'
                        str1 = "Analyte" 'strAnal
                        .Selection.TypeText(Text:=str1)

                        .Selection.Tables.Item(1).Cell(3, 9).Select()
                        str1 = "Int." & ChrW(160) & "Std."
                        .Selection.TypeText(Text:=str1)

                        '.Selection.Tables.Item(1).Cell(3, 10).Select()
                        'str1 = "Peak Area Ratio"
                        '.Selection.TypeText(Text:=str1)

                        .Selection.Tables.Item(1).Cell(2, 10).Select()
                        str1 = "Analyte/IS"
                        .Selection.TypeText(Text:=str1)

                        .Selection.Tables.Item(1).Cell(3, 10).Select()
                        str1 = "Area" & ChrW(160) & "Ratio"
                        .Selection.TypeText(Text:=str1)

                        .Selection.Tables.Item(1).Cell(2, 11).Select()
                        str1 = "% of Run"
                        .Selection.TypeText(Text:=str1)

                        .Selection.Tables.Item(1).Cell(3, 11).Select()
                        str1 = "LLOQ"
                        .Selection.TypeText(Text:=str1)

                        '20190108 LEE:
                        If boolNONELEG Then
                        Else
                            .Selection.TypeText(Text:=" ")
                            .Selection.Font.Superscript = True
                            .Selection.TypeText(Text:="a")
                            .Selection.Font.Superscript = False
                        End If


                        '****



                        ''''wdd.visible = True

                        .Selection.Tables.Item(1).Cell(3, 1).Select()
                        'bottom border this row
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                        '20181014 LEE: Don't need this border
                        '20181219 LEE: ?? Yes it is needed
                        'but for some reason, the entire row is getting underlined
                        'try on cell at a time
                        'that didn't work
                        'try moving it up in order
                        .Selection.Tables.Item(1).Cell(2, 8).Select()
                        'With .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                        '    .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        'End With
                        '.Selection.Tables.Item(1).Cell(2, 9).Select()
                        'With .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                        '    .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        'End With
                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        .Selection.Cells.Merge()
                        With .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                            .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        End With

                        'merge and border
                        .Selection.Tables.Item(1).Cell(1, 7).Select()
                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=3, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        .Selection.Cells.Merge()
                        With .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                            .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        End With

                        .Selection.Tables.Item(1).Cell(1, 1).Select()
                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=4, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        .Selection.Cells.Merge()
                        With .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                            .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        End With


                        .Selection.Tables.Item(1).Cell(2, 3).Select()
                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        .Selection.Cells.Merge()
                        With .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                            .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        End With
                        '20181014 LEE:
                        'Hmmm. For some reason cols 1 and 2 are also getting underlined
                        'upon investigation, it seems to be the top border of the next row
                        .Selection.Tables.Item(1).Cell(3, 1).Select()
                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        With .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop)
                            .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                        End With

                        'begin entering data'
                        Try

                            intStart = 5
                            Dim boolExit As Boolean = False

                            Dim numAnal As Single
                            Dim numIS As Single

                            Dim tblAR As New System.Data.DataTable
                            Dim col2 As New DataColumn
                            tblAR.Columns.Add("AR", Type.GetType("System.Decimal"))

                            Dim rowsAR() As DataRow

                            'strM = "Creating Summary of " & strTempInfo & " Final Extract Stability Table For " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & "..."
                            frmH.lblProgress.Text = strM1 ' & ChrW(10) & "Processing Run ID " & var10
                            frmH.Refresh()


                            'enter Blank with IS
                            Dim numAA As Single
                            Dim numISA As Single
                            Dim arrAreaRatioAnal(2, intRowsW)
                            '1=BlankwIS AR, 2=Blankw/oIS AR
                            Dim arrAreaRatioLLOQ(2, intRowsLLOQ)
                            '1=wIS, 2=w/oIS

                            For Count3 = 0 To intRowsW - 1

                                var10 = rowsW(Count3).Item("RUNID")
                                'enter runid
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 1).Select()
                                If BOOLINCLUDEDATE Then
                                    str2 = CStr(var10)
                                    str1 = GetDateFromRunID(NZ(var10, 0), LDateFormat, intGroup, idTR)
                                    .Selection.TypeText(str2 & ChrW(10) & "(" & str1 & ")")
                                Else
                                    .Selection.TypeText(CStr(var10))
                                End If

                                'enter lot#
                                var1 = NZ(rowsW(Count3).Item("CHARHELPER1"), "NA")
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 2).Select()
                                .Selection.TypeText(CStr(var1))

                                'enter analytearea
                                var1 = NZ(rowsW(Count3).Item("ANALYTEAREA"), "NA")
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 3).Select()
                                If IsNumeric(var1) Then
                                    If boolLUseSigFigsArea Then
                                        var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                    Else
                                        var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                    End If
                                End If

                                .Selection.TypeText(CStr(var1))
                                If IsNumeric(var1) Then
                                    numAA = var1
                                Else
                                    numAA = -1
                                End If

                                'enter ISarea
                                var1 = NZ(rowsW(Count3).Item("INTERNALSTANDARDAREA"), "NA")
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 4).Select()
                                If IsNumeric(var1) Then
                                    If boolLUseSigFigsArea Then
                                        var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                    Else
                                        var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                    End If
                                End If

                                .Selection.TypeText(CStr(var1))
                                If IsNumeric(var1) Then
                                    numISA = var1
                                Else
                                    numISA = -1
                                End If

                                'enter arearatio
                                If numISA <> -1 And numAA <> -1 Then
                                    If numISA = 0 Then
                                        var2 = 0
                                    Else
                                        '20181014 LEE:
                                        'Hmmm. Why is this a percent?
                                        'var2 = numAA / numISA * 100
                                        var2 = numAA / numISA
                                    End If
                                    'var1 = SigFigOrDecString(var2, LSigFig, False)
                                    If IsNumeric(var2) Then
                                        If boolLUseSigFigsAreaRatio Then
                                            var1 = SigFigAreaRatio(CSng(var2), LSigFigAreaRatio, False, True) 'special rounding incorporated
                                        Else
                                            var1 = Format(RoundToDecimalRAFZ((CSng(var2)), LSigFigAreaRatio), strAreaDecAreaRatio)
                                        End If
                                    End If
                                Else
                                    var1 = "NA"
                                End If
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 5).Select()
                                .Selection.TypeText(CStr(var1))
                                If IsNumeric(var1) Then
                                    arrAreaRatioAnal(1, Count3) = CDec(var1)
                                Else
                                    arrAreaRatioAnal(1, Count3) = CDec(0)
                                End If

                            Next Count3

                            'now do lloq
                            For Count3 = 0 To intRowsLLOQ - 1

                                var10 = rowsLLOQ(Count3).Item("RUNID")
                                'enter runid
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 7).Select()
                                If BOOLINCLUDEDATE Then
                                    str2 = CStr(var10)
                                    str1 = GetDateFromRunID(NZ(var10, 0), LDateFormat, intGroup, idTR)
                                    .Selection.TypeText(str2 & ChrW(10) & "(" & str1 & ")")
                                Else
                                    .Selection.TypeText(CStr(var10))
                                End If


                                'enter analytearea
                                var1 = NZ(rowsLLOQ(Count3).Item("ANALYTEAREA"), "NA")
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 8).Select()
                                If IsNumeric(var1) Then
                                    If boolLUseSigFigsArea Then
                                        var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                    Else
                                        var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea + 2), strAreaDec)
                                    End If
                                End If

                                .Selection.TypeText(CStr(var1))
                                If IsNumeric(var1) Then
                                    numAA = var1
                                Else
                                    numAA = -1
                                End If

                                'enter ISarea
                                var1 = NZ(rowsLLOQ(Count3).Item("INTERNALSTANDARDAREA"), "NA")
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 9).Select()
                                If IsNumeric(var1) Then
                                    If boolLUseSigFigsArea Then
                                        var1 = SigFigArea(CSng(var1), LSigFigArea, False, True) 'special rounding incorporated
                                    Else
                                        var1 = Format(RoundToDecimalRAFZ(CSng(var1), LSigFigArea + 2), strAreaDec)
                                    End If
                                End If

                                .Selection.TypeText(CStr(var1))
                                If IsNumeric(var1) Then
                                    numISA = var1
                                Else
                                    numISA = -1
                                End If

                                'enter arearatio
                                '20181014 LEE:
                                'formulae were referencing Area rather than area ratio
                                If numISA <> -1 And numAA <> -1 Then
                                    If numISA = 0 Then
                                        var2 = 0
                                    Else
                                        '20181014 LEE:
                                        'Hmmm. Why is this a percent?
                                        'var2 = numAA / numISA * 100
                                        var2 = numAA / numISA
                                    End If
                                    If IsNumeric(var2) Then
                                        If boolLUseSigFigsAreaRatio Then
                                            var1 = SigFigAreaRatio(CSng(var2), LSigFigAreaRatio, False, True) 'special rounding incorporated
                                        Else
                                            var1 = Format(RoundToDecimalRAFZ(CSng(var2), LSigFigAreaRatio), strAreaDecAreaRatio)
                                        End If
                                    Else
                                        var1 = "NA"
                                    End If

                                    'var1 = CStr(SigFigOrDecString(var2, LSigFig, False))
                                Else
                                    var1 = "NA"
                                End If
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 10).Select()
                                .Selection.TypeText(CStr(var1))
                                arrAreaRatioLLOQ(1, Count3) = CDec(var1)

                            Next Count3

                            'now do:
                            'enter %RunLLOQ
                            'need arearatioaverage
                            Dim numAveAR As Decimal = 0
                            Dim intAveAR As Short = 0
                            numAveAR = 0
                            For Count3 = 0 To intRowsLLOQ - 1
                                var2 = arrAreaRatioLLOQ(1, Count3)
                                If IsNumeric(var2) Then
                                    numAveAR = numAveAR + var2
                                    intAveAR = intAveAR + 1
                                End If
                            Next Count3
                            If intAveAR = 0 Then
                                var2 = 0
                            Else
                                var2 = CDec(numAveAR / intAveAR)
                            End If
                            For Count3 = 0 To intRowsW - 1
                                var1 = arrAreaRatioAnal(1, Count3)
                                If IsNumeric(var1) And IsNumeric(var2) Then
                                    If var2 = 0 Then
                                        var4 = Format(0, strQCDec)
                                    Else
                                        var3 = var1 / var2 * 100
                                        var4 = Format(var3, strQCDec)
                                    End If
                                Else
                                    var4 = "NA"
                                End If
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 11).Select()
                                .Selection.TypeText(CStr(var4))
                            Next Count3


                            'now do Blank Without IS

                            If intRowsW > intRowsLLOQ Then
                                intStart = intStart + intRowsW + 2
                            Else
                                intStart = intStart + intRowsLLOQ + 2
                            End If

                            Dim intHS2 As Short
                            intHS2 = intStart

                            'enter headings

                            int1 = InStr(strWRunId, " ", CompareMethod.Text)
                            If int1 = 0 Then
                                str2 = strWRunId
                            Else
                                str1 = Mid(strWRunId, 1, int1 - 1)
                                str2 = Mid(strWRunId, int1 + 1, Len(strWRunId))
                            End If

                            'Blank With IS
                            If int1 = 0 Then
                            Else
                                .Selection.Tables.Item(1).Cell(intStart + 1, 1).Select()
                                str1 = str1 ' "Watson"
                                .Selection.TypeText(Text:=str1)
                            End If

                            .Selection.Tables.Item(1).Cell(intStart + 2, 1).Select()
                            str1 = str2 '"Run ID"
                            .Selection.TypeText(Text:=str1)

                            .Selection.Tables.Item(1).Cell(intStart + 2, 2).Select()
                            str1 = "Lot #"
                            .Selection.TypeText(Text:=str1)

                            .Selection.Tables.Item(1).Cell(intStart, 1).Select()
                            str1 = "Blank Without Int. Std. Samples"
                            .Selection.TypeText(Text:=str1)

                            .Selection.Tables.Item(1).Cell(intStart + 1, 3).Select()
                            str1 = "Peak Area"
                            .Selection.TypeText(Text:=str1)

                            .Selection.Tables.Item(1).Cell(intStart + 2, 3).Select()
                            '20181014 LEE
                            'change to 'Analyte'
                            str1 = "Analyte" 'strAnal
                            .Selection.TypeText(Text:=str1)

                            .Selection.Tables.Item(1).Cell(intStart + 2, 4).Select()
                            str1 = "Int. Std."
                            .Selection.TypeText(Text:=str1)

                            'LLOQ Samples
                            .Selection.Tables.Item(1).Cell(intStart, 7).Select()
                            str1 = "LLOQ Samples (" & varNomConc & " " & strUnits & ")"
                            .Selection.TypeText(Text:=str1)

                            int1 = InStr(strWRunId, " ", CompareMethod.Text)
                            If int1 = 0 Then
                                str2 = strWRunId
                            Else
                                str1 = Mid(strWRunId, 1, int1 - 1)
                                str2 = Mid(strWRunId, int1 + 1, Len(strWRunId))
                            End If

                            If int1 = 0 Then
                            Else
                                .Selection.Tables.Item(1).Cell(intStart + 1, 7).Select()
                                str1 = str1 ' "Watson"
                                .Selection.TypeText(Text:=str1)
                            End If

                            .Selection.Tables.Item(1).Cell(intStart + 2, 7).Select()

                            If BOOLINCLUDEDATE Then
                                'str1 = str2 & ChrW(10) & "(Analysis Date)"
                                '20180420 LEE:
                                str1 = str2 & ChrW(10) & "(" & GetAnalysisDateLabel(intTableID) & ")"

                            Else
                                str1 = str2
                            End If

                            .Selection.TypeText(Text:=str1)

                            .Selection.Tables.Item(1).Cell(intStart + 1, 8).Select()
                            str1 = "Peak Area"
                            .Selection.TypeText(Text:=str1)

                            .Selection.Tables.Item(1).Cell(intStart + 2, 8).Select()
                            '20181014 LEE
                            'change to 'Analyte'
                            str1 = "Analyte" 'strAnal
                            .Selection.TypeText(Text:=str1)

                            .Selection.Tables.Item(1).Cell(intStart + 2, 9).Select()
                            str1 = "Int. Std."
                            .Selection.TypeText(Text:=str1)

                            .Selection.Tables.Item(1).Cell(intStart, 11).Select()
                            str1 = "% of Lowest"
                            .Selection.TypeText(Text:=str1)

                            .Selection.Tables.Item(1).Cell(intStart + 1, 11).Select()
                            str1 = "Acceptable" & ChrW(160) & "IS"
                            .Selection.TypeText(Text:=str1)

                            .Selection.Tables.Item(1).Cell(intStart + 2, 11).Select()
                            str1 = "Response"
                            .Selection.TypeText(Text:=str1)

                            '20190108 LEE:
                            If boolNONELEG Then
                            Else
                                .Selection.TypeText(Text:=" ")
                                .Selection.Font.Superscript = True
                                .Selection.TypeText(Text:="b")
                                .Selection.Font.Superscript = False
                            End If


                            .Selection.Tables.Item(1).Cell(intStart + 2, 1).Select()
                            '20181014 LEE:
                            '.Selection.Tables.Item(1).Cell(intStart + 2, 2).Select()
                            'bottom border this row
                            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle


                            'merge and border

                            '20181014 LEE: Don't need this border
                            '20181219 LEE: ?? Yes we need this border!
                            .Selection.Tables.Item(1).Cell(intStart + 1, 8).Select()
                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                            .Selection.Cells.Merge()
                            With .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                            End With


                            .Selection.Tables.Item(1).Cell(intStart, 7).Select()
                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=3, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                            .Selection.Cells.Merge()
                            With .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                            End With

                            .Selection.Tables.Item(1).Cell(intStart, 1).Select()
                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=4, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                            .Selection.Cells.Merge()
                            With .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                            End With

                            ''20181014 LEE: Don't need this border
                            ''20181219 LEE: ?? Yes we need this border!
                            '.Selection.Tables.Item(1).Cell(intStart + 1, 8).Select()
                            '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                            '.Selection.Cells.Merge()
                            'With .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                            '    .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                            'End With

                            .Selection.Tables.Item(1).Cell(intStart + 1, 3).Select()
                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                            .Selection.Cells.Merge()
                            With .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                            End With
                            '20181014 LEE:
                            'Hmmm. For some reason cols 1 and 2 are also getting underlined
                            'upon investigation, it seems to be the top border of the next row
                            .Selection.Tables.Item(1).Cell(intStart + 2, 1).Select()
                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                            With .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop)
                                .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                            End With

                            intStart = intStart + 4

                            For Count3 = 0 To intRowsWO - 1

                                var10 = rowsWO(Count3).Item("RUNID")
                                'enter runid
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 1).Select()
                                If BOOLINCLUDEDATE Then
                                    str2 = CStr(var10)
                                    str1 = GetDateFromRunID(NZ(var10, 0), LDateFormat, intGroup, idTR)
                                    .Selection.TypeText(str2 & ChrW(10) & "(" & str1 & ")")
                                Else
                                    .Selection.TypeText(CStr(var10))
                                End If


                                'enter lot#
                                var1 = NZ(rowsWO(Count3).Item("CHARHELPER1"), "NA")
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 2).Select()
                                .Selection.TypeText(CStr(var1))

                                'enter analytearea
                                var1 = NZ(rowsWO(Count3).Item("ANALYTEAREA"), "NA")
                                If IsNumeric(var1) Then
                                    If boolLUseSigFigsArea Then
                                        var1 = SigFigArea(CSng(var1), LSigFigArea, False, True) 'special rounding incorporated
                                    Else
                                        var1 = Format(RoundToDecimalRAFZ(CSng(var1), LSigFigArea), strAreaDec)
                                    End If
                                End If
                                If IsNumeric(var1) Then
                                    numAA = var1
                                Else
                                    numAA = -1
                                End If
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 3).Select()

                                '20181219 LEE: Frontage wants to see this information
                                'var1 = "NA"  'Because we are comparing Blank with IS to Blank, the Analyte Area is not applicable
                                .Selection.TypeText(CStr(var1))


                                'enter ISarea
                                var1 = NZ(rowsWO(Count3).Item("INTERNALSTANDARDAREA"), "NA")
                                If IsNumeric(var1) Then
                                    If boolLUseSigFigsArea Then
                                        var1 = SigFigArea(CSng(var1), LSigFigArea, False, True) 'special rounding incorporated
                                    Else
                                        var1 = Format(RoundToDecimalRAFZ(CSng(var1), LSigFigArea), strAreaDec)
                                    End If
                                End If
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 4).Select()
                                .Selection.TypeText(CStr(var1))
                                If IsNumeric(var1) Then
                                    numISA = var1
                                Else
                                    numISA = -1
                                End If
                                arrAreaRatioAnal(2, Count3) = NZ(var1, 0)

                            Next

                            'now do lloq
                            For Count3 = 0 To intRowsLLOQ - 1

                                var10 = rowsLLOQ(Count3).Item("RUNID")
                                'enter runid
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 7).Select()
                                If BOOLINCLUDEDATE Then
                                    str2 = CStr(var10)
                                    str1 = GetDateFromRunID(NZ(var10, 0), LDateFormat, intGroup, idTR)
                                    .Selection.TypeText(str2 & ChrW(10) & "(" & str1 & ")")
                                Else
                                    .Selection.TypeText(CStr(var10))
                                End If

                                'enter analytearea
                                var1 = NZ(rowsLLOQ(Count3).Item("ANALYTEAREA"), "NA")
                                If IsNumeric(var1) Then
                                    If boolLUseSigFigsArea Then
                                        var1 = SigFigArea(CSng(var1), LSigFigArea, False, True) 'special rounding incorporated
                                    Else
                                        var1 = Format(RoundToDecimalRAFZ(CSng(var1), LSigFigArea), strAreaDec)
                                    End If
                                    numAA = var1
                                Else
                                    numAA = -1
                                End If
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 8).Select()
                                '20181219 LEE: Frontage wants to have all data shown
                                'var1 = "NA" '
                                .Selection.TypeText(CStr(var1))

                                'enter ISarea
                                var1 = NZ(rowsLLOQ(Count3).Item("INTERNALSTANDARDAREA"), "NA")
                                If IsNumeric(var1) Then
                                    If boolLUseSigFigsArea Then
                                        var1 = SigFigArea(CSng(var1), LSigFigArea, False, True) 'special rounding incorporated
                                    Else
                                        var1 = Format(RoundToDecimalRAFZ(CSng(var1), LSigFigArea), strAreaDec)
                                    End If
                                End If
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 9).Select()
                                .Selection.TypeText(CStr(var1))
                                If IsNumeric(var1) Then
                                    numISA = var1
                                Else
                                    numISA = -1
                                End If
                                arrAreaRatioLLOQ(2, Count3) = NZ(var1, 0)

                                ''enter %RunLLOQ
                                ' ''''wdd.visible = True
                                'If Count3 > intRowsWO - 1 Then
                                'Else
                                '    var1 = arrAreaRatioAnal(1, Count3)
                                '    var2 = arrAreaRatioAnal(2, Count3)
                                '    If IsNumeric(var1) And IsNumeric(var2) Then
                                '        If var2 = 0 Then
                                '            'var4 = CStr(SigFigOrDecString(0, LSigFig, False))
                                '            var4 = Format(0, strQCDec)
                                '        Else
                                '            var3 = var1 / var2 * 100
                                '            'var4 = CStr(SigFigOrDecString(var3, LSigFig, False))
                                '            var4 = Format(var3, strQCDec)
                                '        End If
                                '    Else
                                '        var4 = "NA"
                                '    End If
                                '    .Selection.Tables.Item(1).Cell(Count3 + intStart, 11).Select()
                                '    .Selection.TypeText(CStr(var4))
                                'End If

                            Next

                            '20181014 LEE:
                            'now do:
                            'enter %RunLLOQ
                            'need arearatioaverage
                            'need arearatioaverage
                            numAveAR = 0
                            intAveAR = 0
                            For Count3 = 0 To intRowsLLOQ - 1
                                var2 = arrAreaRatioLLOQ(1, Count3)
                                If IsNumeric(var2) Then
                                    numAveAR = numAveAR + var2
                                    intAveAR = intAveAR + 1
                                End If
                            Next Count3
                            If intAveAR = 0 Then
                                var2 = 0
                            Else
                                var2 = CDec(numAveAR / intAveAR)
                            End If

                            For Count3 = 0 To intRowsW - 1
                                var1 = arrAreaRatioAnal(2, Count3)
                                If IsNumeric(var1) And IsNumeric(var2) Then
                                    If var2 = 0 Then
                                        var4 = Format(0, strQCDec)
                                    Else
                                        var3 = var1 / var2 * 100
                                        var4 = Format(var3, strQCDec)
                                    End If
                                Else
                                    var4 = "NA"
                                End If
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 11).Select()
                                .Selection.TypeText(CStr(var4))
                            Next Count3

                            'END NEW

                            Dim intRow As Short
                            intRow = intStart + int3 + 1

                            GoTo endStats

                            'now enter Mean/Bias/n labels
                            int8 = intRow - 1
                            If boolSTATSMEAN Then
                                int8 = int8 + 1
                                .Selection.Tables.Item(1).Cell(int8, 2).Select()
                                .Selection.TypeText("Mean")
                                '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, 1).Select()
                            End If
                            If boolSTATSSD Then
                                int8 = int8 + 1
                                .Selection.Tables.Item(1).Cell(int8, 2).Select()
                                .Selection.TypeText("S.D.") '((Mean/NomConc)-1)*100)
                                '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, 1).Select()
                                .Selection.Tables.Item(1).Cell(int1 + int8, 1).Select()
                            End If
                            If boolSTATSCV Then
                                int8 = int8 + 1
                                .Selection.Tables.Item(1).Cell(int8, 2).Select()
                                .Selection.TypeText(ReturnPrecLabel()) '(sd/mean)*100)
                                '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, 1).Select()
                                .Selection.Tables.Item(1).Cell(int1 + int8, 1).Select()
                            End If

                            If boolSTATSN Then
                                int8 = int8 + 1
                                .Selection.Tables.Item(1).Cell(int8, 2).Select()
                                '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, 1).Select()
                                .Selection.TypeText("n")
                            End If

                            'now start enter stats for analyte
                            Dim tbl2SA As System.Data.DataTable = dv2.ToTable
                            strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                            rows2E = tbl2SA.Select("RUNSAMPLEORDERNUMBER > -1", strS)
                            int3 = rows2E.Length

                            int8 = intRow - 1
                            .Selection.Tables.Item(1).Cell(int8, 2).Select()

                            Dim numMeanRT As Single
                            Dim numMeanIS As Single
                            Dim numMeanRTIS As Single
                            Dim numMeanAR As Single
                            Dim numSDIS As Single
                            Dim numSDRTIS As Single
                            Dim numSDAR As Single
                            Dim numPrecRT As Single
                            Dim numPrecIS As Single
                            Dim numPrecRTIS As Single
                            Dim numPrecAR As Single


                            Try
                                numMean = MeanDR(rows2E, "ANALYTEAREA", False, "ALIQUOTFACTOR", False, False)
                                var1 = numMean
                                If boolLUseSigFigsArea Then
                                    var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                Else
                                    var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                End If
                                numMean = CDec(var1)
                                Call InsertQCTables(intTableID, idTR, charFCID, NZ(varNom, 0), 1, "Mean", numMean, CSng(var10), Count1, strDo, 0, 0, False)
                            Catch ex As Exception

                            End Try
                            If boolSTATSMEAN Then
                                Try
                                    'enter Mean of Peak Area
                                    int8 = int8 + 1
                                    .Selection.Tables.Item(1).Cell(int8, 3).Select()
                                    numMean = MeanDR(rows2E, "ANALYTEAREA", False, "ALIQUOTFACTOR", False, False)
                                    Call InsertQCTables(intTableID, idTR, charFCID, NZ(varNom, 0), 1, "Mean", numMean, CSng(var10), Count1, strDo, 0, 0, False)
                                    var1 = numMean
                                    If boolLUseSigFigsArea Then
                                        var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                    Else
                                        var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                    End If
                                    numMean = CDec(var1)
                                    .Selection.TypeText(CStr(var1))
                                    '.Selection.TypeText(CStr(Format(numMean, "0")))
                                Catch ex As Exception

                                End Try

                                Try
                                    'enter Mean of RT
                                    .Selection.Tables.Item(1).Cell(int8, 4).Select()
                                    numMeanRT = MeanDR(rows2E, "ANALYTEPEAKRETENTIONTIME", False, "ALIQUOTFACTOR", False, False)
                                    .Selection.TypeText(CStr(Format(numMeanRT, "0.00")))
                                Catch ex As Exception

                                End Try

                                Try
                                    'enter Mean of Peak Area
                                    .Selection.Tables.Item(1).Cell(int8, 6).Select()
                                    numMeanIS = MeanDR(rows2E, "INTERNALSTANDARDAREA", False, "ALIQUOTFACTOR", False, False)
                                    var1 = numMeanIS
                                    If boolLUseSigFigsArea Then
                                        var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                    Else
                                        var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                    End If
                                    .Selection.TypeText(CStr(var1))
                                    '.Selection.TypeText(CStr(Format(numMeanIS, "0")))
                                Catch ex As Exception

                                End Try

                                Try
                                    'enter Mean of RT
                                    .Selection.Tables.Item(1).Cell(int8, 7).Select()
                                    numMeanRTIS = MeanDR(rows2E, "INTERNALSTANDARDRETENTIONTIME", False, "ALIQUOTFACTOR", False, False)
                                    .Selection.TypeText(CStr(Format(numMeanRTIS, "0.00")))
                                Catch ex As Exception

                                End Try

                                Try
                                    'enter Mean of Area Ratio
                                    .Selection.Tables.Item(1).Cell(int8, 8).Select()
                                    numMeanAR = MeanDR(rowsAR, "AR", False, "ALIQUOTFACTOR", False, False)
                                    var1 = numMeanAR
                                    If boolLUseSigFigsArea Then
                                        var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                    Else
                                        var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea + 2), strAreaDec)
                                    End If
                                    .Selection.TypeText(CStr(var1))
                                Catch ex As Exception

                                End Try

                            End If

                            Dim numSDRT As Single


                            Try
                                var1 = StdDevDR(rows2E, "ANALYTEAREA", False, "ALIQUOTFACTOR", False, False)
                                If boolLUseSigFigsArea Then
                                    numSD = SigFigOrDecString(var1, LSigFigArea, True) 'special rounding incorporated
                                Else
                                    numSD = RoundToDecimalRAFZ(var1, LSigFigArea)
                                End If

                                Call InsertQCTables(intTableID, idTR, charFCID, NZ(varNom, 0), 1, "SD", numSD, CSng(var10), Count1, strDo, 0, 0, False)
                            Catch ex As Exception

                            End Try
                            If boolSTATSSD Then
                                int8 = int8 + 1
                                If int3 < gSDMax Then
                                    .Selection.Tables.Item(1).Cell(int8, 3).Select()
                                    .Selection.TypeText("NA")
                                    .Selection.Tables.Item(1).Cell(int8, 4).Select()
                                    .Selection.TypeText("NA")
                                    .Selection.Tables.Item(1).Cell(int8, 6).Select()
                                    .Selection.TypeText("NA")
                                    .Selection.Tables.Item(1).Cell(int8, 7).Select()
                                    .Selection.TypeText("NA")
                                    .Selection.Tables.Item(1).Cell(int8, 8).Select()
                                    .Selection.TypeText("NA")
                                Else
                                    Try
                                        'enter SD of Peak Area
                                        .Selection.Tables.Item(1).Cell(int8, 3).Select()
                                        var1 = StdDevDR(rows2E, "ANALYTEAREA", False, "ALIQUOTFACTOR", False, False)
                                        If boolLUseSigFigsArea Then
                                            numSD = SigFigOrDecString(var1, LSigFigArea, True) 'special rounding incorporated
                                        Else
                                            numSD = RoundToDecimalRAFZ(var1, LSigFigArea)
                                        End If

                                        var1 = numSD
                                        If boolLUseSigFigsArea Then
                                            var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                        Else
                                            var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                        End If
                                        .Selection.TypeText(CStr(var1))

                                    Catch ex As Exception

                                    End Try

                                    Try
                                        'enter SD of RT
                                        .Selection.Tables.Item(1).Cell(int8, 4).Select()
                                        numSDRT = StdDevDR(rows2E, "ANALYTEPEAKRETENTIONTIME", False, "ALIQUOTFACTOR", False, False)
                                        If IsNumeric(numSDRT) Then
                                            numSDRT = RoundToDecimalRAFZ(numSDRT, 2)
                                        End If
                                        .Selection.TypeText(CStr(Format(numSDRT, "0.00")))
                                        '.Selection.TypeText(CStr(SigFigOrDecString(numSDRT, LSigFig, False)))
                                    Catch ex As Exception

                                    End Try

                                    Try
                                        'enter SD of Peak Area
                                        .Selection.Tables.Item(1).Cell(int8, 6).Select()
                                        var1 = StdDevDR(rows2E, "INTERNALSTANDARDAREA", False, "ALIQUOTFACTOR", False, False)
                                        If boolLUseSigFigs Then
                                            numSDIS = SigFigOrDec(var1, LSigFigArea, False)
                                        Else
                                            numSDIS = RoundToDecimalRAFZ(var1, LSigFigArea)
                                        End If

                                        var1 = numSDIS
                                        If boolLUseSigFigsArea Then
                                            var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                        Else
                                            var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea + 2), strAreaDec)
                                        End If
                                        .Selection.TypeText(CStr(var1))


                                    Catch ex As Exception

                                    End Try

                                    Try
                                        'enter SD of RT
                                        .Selection.Tables.Item(1).Cell(int8, 7).Select()
                                        numSDRTIS = StdDevDR(rows2E, "INTERNALSTANDARDRETENTIONTIME", False, "ALIQUOTFACTOR", False, False)
                                        If IsNumeric(numSDRTIS) Then
                                            numSDRTIS = RoundToDecimalRAFZ(numSDRTIS, 2)
                                        End If
                                        .Selection.TypeText(CStr(Format(numSDRTIS, "0.00")))
                                        '.Selection.TypeText(CStr(SigFigOrDecString(numSDRTIS, LSigFig, False)))
                                    Catch ex As Exception

                                    End Try

                                    Try
                                        'enter SD of Area Ratio
                                        .Selection.Tables.Item(1).Cell(int8, 8).Select()
                                        numSDAR = StdDevDR(rowsAR, "AR", False, "ALIQUOTFACTOR", False, False)
                                        var1 = numSDAR
                                        If boolLUseSigFigsArea Then
                                            var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                        Else
                                            var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea + 2), strAreaDec)
                                        End If
                                        .Selection.TypeText(CStr(var1))
                                    Catch ex As Exception

                                    End Try

                                End If

                            End If


                            Try
                                numPrec = CalcCVPercent(numSD, numMean, intQCDec)
                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, 1, "Precision", numPrec, CSng(var10), Count1, strDo, 0, 0, False)
                            Catch ex As Exception

                            End Try
                            If boolSTATSCV Then
                                Try
                                    'enter %CV
                                    int8 = int8 + 1

                                    If int3 < gSDMax Then
                                        .Selection.Tables.Item(1).Cell(int8, 3).Select()
                                        .Selection.TypeText("NA")
                                        .Selection.Tables.Item(1).Cell(int8, 4).Select()
                                        .Selection.TypeText("NA")
                                        .Selection.Tables.Item(1).Cell(int8, 6).Select()
                                        .Selection.TypeText("NA")
                                        .Selection.Tables.Item(1).Cell(int8, 7).Select()
                                        .Selection.TypeText("NA")
                                        .Selection.Tables.Item(1).Cell(int8, 8).Select()
                                        .Selection.TypeText("NA")
                                    Else

                                        Try
                                            .Selection.Tables.Item(1).Cell(int8, 3).Select()

                                            .Selection.TypeText(Format(numPrec, strQCDec))
                                        Catch ex As Exception

                                        End Try

                                        Try
                                            .Selection.Tables.Item(1).Cell(int8, 4).Select()
                                            numPrecRT = RoundToDecimalA(RoundToDecimalRAFZ((numSDRT / numMeanRT * 100), intQCDec + 4), intQCDec)
                                            .Selection.TypeText(Format(numPrecRT, strQCDec))
                                        Catch ex As Exception

                                        End Try

                                        Try
                                            .Selection.Tables.Item(1).Cell(int8, 6).Select()
                                            numPrecIS = RoundToDecimalA(RoundToDecimalRAFZ((numSDIS / numMeanIS * 100), intQCDec + 4), intQCDec)
                                            .Selection.TypeText(Format(numPrecIS, strQCDec))
                                        Catch ex As Exception

                                        End Try

                                        Try
                                            .Selection.Tables.Item(1).Cell(int8, 7).Select()
                                            numPrecRTIS = RoundToDecimalA(RoundToDecimalRAFZ((numSDRTIS / numMeanRTIS * 100), intQCDec + 4), intQCDec)
                                            .Selection.TypeText(Format(numPrecRTIS, strQCDec))
                                        Catch ex As Exception

                                        End Try

                                        Try
                                            .Selection.Tables.Item(1).Cell(int8, 8).Select()
                                            numPrecAR = RoundToDecimalA(RoundToDecimalRAFZ((numSDAR / numMeanAR * 100), intQCDec + 4), intQCDec)
                                            .Selection.TypeText(Format(numPrecAR, strQCDec))
                                        Catch ex As Exception

                                        End Try


                                    End If

                                Catch ex As Exception

                                End Try
                            End If


                            Try
                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, 1, "n", int3, CSng(var10), Count1, strDo, 0, 0, False)
                            Catch ex As Exception

                            End Try
                            If boolSTATSN Then
                                Try
                                    'enter n
                                    int8 = int8 + 1
                                    .Selection.Tables.Item(1).Cell(int8, 3).Select()
                                    .Selection.TypeText(CStr(int3))
                                    .Selection.Tables.Item(1).Cell(int8, 4).Select()
                                    .Selection.TypeText(CStr(int3))
                                    .Selection.Tables.Item(1).Cell(int8, 6).Select()
                                    .Selection.TypeText(CStr(int3))
                                    .Selection.Tables.Item(1).Cell(int8, 7).Select()
                                    .Selection.TypeText(CStr(int3))
                                    .Selection.Tables.Item(1).Cell(int8, 8).Select()
                                    .Selection.TypeText(CStr(int3))
                                    '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                    '.Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 2).Select()


                                Catch ex As Exception

                                End Try
                            End If

endStats:

                        Catch ex As Exception

                            str1 = "There was a problem preparing table:"
                            str1 = strM1 & ChrW(10) & ChrW(10) & str1
                            str1 = str1 & ChrW(10) & ChrW(10)
                            str1 = str1 & ex.Message
                            MsgBox(str1, vbInformation, "Problem...")

                        End Try


                        'END NEW STUFF
                        .Selection.Tables.Item(1).Cell(.Selection.Tables.Item(1).Rows.Count, 1).Select()

                        'bottom border this row
                        .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.MoveRight Unit:=Microsoft.Office.Interop.Word.wdunits.word.wdunits.wdCharacter, Count:=ctQCs, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                        'autofit window
                        .Selection.Tables.Item(1).Select()
                        'autofit table
                        Call AutoFitTable(wd, False)

                        ''''wdd.visible = True

                        ''now force fit column 5
                        '.Selection.Tables.Item(1).Cell(3, 5).Select()
                        '.Selection.SelectColumn()
                        '.Selection.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.wdpreferredwidthtype.wdPreferredWidthPercent ' wdPreferredWidthPercent
                        '.Selection.Columns.PreferredWidth = 1
                        ''pesky. do it again
                        '.Selection.Tables.Item(1).Cell(3, 5).Select()
                        '.Selection.SelectColumn()
                        '.Selection.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.wdpreferredwidthtype.wdPreferredWidthPercent ' wdPreferredWidthPercent
                        '.Selection.Columns.PreferredWidth = 1


                        ''''wdd.visible = True

                        .Selection.Tables.Item(1).Cell(1, 1).Select()

                        'remove unused rows
                        Call RemoveRows(wd, ctTbl)

                        'str1 = str2 & " Final Extract Stability: Summary of " & rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION") & " Interpolated QC Standard Concentrations."

                        '***
                        strA = rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")
                        strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, False)
                        Call EnterTableNumber(wd, strTName, 3, strA, strTempInfo, intTableID, intGroup, idTR)
                        '***
                        'Call EnterTableNumber(wd, str1, 3)

                        'now force fit column 5
                        .Selection.Tables.Item(1).Cell(intStart, 6).Select()
                        .Selection.SelectColumn()
                        .Selection.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent ' wdPreferredWidthPercent
                        .Selection.Columns.PreferredWidth = 1


                        'enter a table record in tblTableN
                        'ctTableN = ctTableN + 1
                        Dim dtblr As DataRow = tblTableN.NewRow
                        dtblr.BeginEdit()
                        dtblr.Item("TableNumber") = ctTableN
                        dtblr.Item("AnalyteName") = strDo 'arrAnalytes(1, Count1)
                        dtblr.Item("TableName") = strTNameO
                        dtblr.Item("TableID") = intTableID
                        dtblr.Item("CHARFCID") = charFCID
                        dtblr.Item("TableNameNew") = strTName
                        tblTableN.Rows.Add(dtblr)

                        'split table, if needed
                        str1 = frmH.lblProgress.Text

                        ctLegend = ctLegend + 1
                        intLeg = intLeg + 1
                        arrLegend(1, intLeg) = "NA"
                        arrLegend(2, intLeg) = "Not Applicable"
                        arrLegend(3, intLeg) = False
                        arrLegend(4, intLeg) = False

                        '20190108 LEE:
                        If boolNONELEG Then
                        Else
                            intLeg = intLeg + 1
                            arrLegend(1, intLeg) = "a"
                            arrLegend(2, intLeg) = "(Area Ratio of Blank With IS) / (Area Ratio of LLOQ Samples) where 'LLOQ' = Lower Limit of Quantitation"
                            arrLegend(3, intLeg) = True
                            arrLegend(4, intLeg) = True

                            intLeg = intLeg + 1
                            arrLegend(1, intLeg) = "b"
                            arrLegend(2, intLeg) = "(Int. Std. Peak Area of Blank Without Int. Std.) / (Int. Std. Peak Area of LLOQ Samples)"
                            arrLegend(3, intLeg) = True
                            arrLegend(4, intLeg) = True
                        End If

                        Call AutoFitTable(wd, BOOLINCLUDEDATE)

                        strM = "Finalizing " & strTName & "..."
                        strM1 = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                        str1 = strM1

                        frmH.lblProgress.Text = strM1
                        frmH.Refresh()

                        Call SplitTable(wd, 4, intLeg, arrLegend, str1, False, ctLegend + 2, False, False, False, intTableID)
                        'Sub SplitTable(ByVal wd As Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, 
                        'ByVal arr As Object, ByVal strT As String, ByVal DoLegend As Boolean, ByVal intSplitRows As Short, ByVal boolSmallFont As Boolean)

                        'autofit window again
                        .Selection.Tables.Item(1).Select()
                        'autofit table
                        Call AutoFitTable(wd, False)

                        'pesky
                        'bottom-border row1
                        .Selection.Tables.Item(1).Cell(1, 1).Select()
                        .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.MoveRight Unit:=Microsoft.Office.Interop.Word.wdunits.word.wdunits.wdCharacter, Count:=ctQCs, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        'move to line below table
                        Call MoveOneCellDown(wd)

                        Call InsertLegend(wd, intTableID, idTR, False, 1)

                    Catch ex As Exception

                        'come out of table
                        'wd.visible = True 'debug
                        Call MoveOneCellDown(wd)
                        wd.Selection.TypeParagraph()
                        wd.Selection.TypeParagraph()
                        .Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter) ', Count:=4, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

                        strM = "There was a problem creating '" & strTName & "'" & ChrW(10) & ChrW(10) & ex.Message
                        MsgBox(strM, MsgBoxStyle.Information, "Problem...")
                        GoTo end2
                    End Try


                Else
                    boolJustTable = False

                End If

end1:
                If boolJustTable Then
                    var1 = strTempInfo
                    'replace numeric with verbose in strTempInfo
                    'look for a sequence of characters that is numeric
                    var3 = ""
                    var4 = ""
                    Dim bool1 As Boolean
                    Dim bool2 As Boolean
                    bool1 = False 'Start
                    bool2 = False 'End
                    str2 = ""
                    For Count2 = 1 To Len(var1)
                        var2 = Mid(var1, Count2, 1)
                        If StrComp(var2, " ", CompareMethod.Text) = 0 Then
                            var2 = "a"
                        End If
                        If IsNumeric(var2) Then
                            var3 = var3 & var2
                            If IsNumeric(var3) Then
                                var4 = var3
                                bool1 = True
                            Else
                            End If
                        Else
                            If bool1 Then
                                bool2 = True
                            End If
                        End If
                        If bool1 And bool2 Then
                            Exit For
                        End If
                    Next

                    If bool1 = False Then
                        var2 = "[NA]"
                    Else
                        var2 = VerboseNumber(var4, True)
                        str2 = Replace(var1, CStr(var4), var2, 1, 1, CompareMethod.Text)
                    End If

                    str1 = str2 & " " & strTName ' " Final Extract Stability: Summary of " & rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION") & " Interpolated QC Standard Concentrations."
                    str2 = str1
                    'str1 = rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")
                    ''Call JustTable(wd, str1, str2, strDo, strTName, intTableID)
                    'Call JustTable(wd, str1, strTName, strDo, strTName, intTableID, strTempInfo, "")

                    str1 = NZ(rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION"), "")
                    strA = str1
                    'Call JustTable(wd, str1, str2, strDo, strTName, intTableID)
                    If Len(str1) = 0 Then
                    Else
                        strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, False)
                        Call JustTable(wd, str1, strTName, strDo, strTName, intTableID, strTempInfo, "", strTNameO, intGroup, idTR)
                    End If

                End If

next1:

            Next
end2:
        End With


    End Sub



    Sub MVSummaryOfIQCBetweenRun_11(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal idTR As Int64)

        'boolQCREPORTACCVALUES

        Dim intLastRow As Int32 = 1
        Dim intCRow As Int32 = 1

        Dim numNomConc As Decimal
        Dim numSD As Decimal
        Dim var1, var2, var3, var4, var5, var10
        Dim dvDo As System.Data.DataView
        Dim strTName As String
        Dim intDo As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim Count4 As Short
        Dim Count5 As Short
        Dim strDo As String
        Dim bool As Boolean
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim tbl1 As System.Data.DataTable
        Dim dv1 As System.Data.DataView
        Dim rows1() As DataRow
        Dim intRows1 As Short
        Dim strF1 As String
        Dim tbl2 As System.Data.DataTable
        Dim dv2 As System.Data.DataView
        Dim rows2() As DataRow
        Dim intRows2 As Short
        Dim strF2 As String
        Dim tbl3 As System.Data.DataTable
        Dim dv3 As System.Data.DataView
        Dim rows3() As DataRow
        Dim intRows3 As Short
        Dim strF3 As String
        Dim intTableID As Short
        Dim tbl4 As System.Data.DataTable
        Dim dv4 As System.Data.DataView
        Dim rows4() As DataRow
        Dim intRows4 As Short
        Dim strF4 As String
        Dim strS As String
        Dim intNumRuns As Short
        Dim dv As System.Data.DataView
        Dim tblNumRuns As System.Data.DataTable
        Dim tblLevels As System.Data.DataTable
        Dim intNumLevels As Short
        Dim intTblRows As Short
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim strF As String
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim int4 As Short
        Dim int10 As Short
        Dim intRowsX As Short
        Dim tblX As System.Data.DataTable
        Dim varNom
        Dim strConcUnits As String
        Dim intLeg As Short
        Dim ctQCLegend As Short
        Dim ctDilLeg As Short
        Dim strA As String
        Dim strB As String

        Dim ctLegend As Short
        Dim fontsize
        Dim hi, lo
        Dim rows10() As DataRow
        Dim rows11() As DataRow
        Dim intRowsAnal As Short
        Dim arrFP(2, 20) 'FlagPercent array
        Dim strFP As String
        Dim numMean As Decimal
        Dim numBias As Decimal
        Dim tblZ As System.Data.DataTable
        Dim tblAnova As New System.Data.DataTable
        Dim ReturnAnova(1)
        Dim dvAn As System.Data.DataView
        Dim tblAnGo As New System.Data.DataTable
        Dim p1, p2, p3, p4, p5, p6, p7, p8, p9, p10
        Dim strM As String
        Dim numDF As Decimal
        Dim rowsX() As DataRow
        Dim intLegStart As Short
        Dim boolPro As Boolean
        Dim intRow As Short
        Dim boolJustTable As Boolean
        Dim strTempInfo As String
        Dim intExp As Short
        Dim ctExp As Short
        Dim int8 As Short
        Dim intN As Short
        Dim rowsActual() As DataRow
        Dim strFActual As String
        Dim v1, v2, vU
        Dim charFCID As String
        Dim intFirstAnova As Int16 = 0
        Dim boolFirstAnova As Boolean = False

        Dim varConc
        Dim boolOC As Boolean ' if eliminated

        Dim num1 As Single
        Dim num2 As Single
        Dim num3 As Single

        Dim vAnalyteIndex
        Dim vMasterAssayID
        Dim vAnalyteID
        Dim tblAG As DataTable = tblAnalyteGroups 'tblAnalyteGroups has all analytes, not just accepted

        Dim intGroup As Short
        Dim strAnal As String
        Dim strAnalC As String
        Dim strMatrix As String
        Dim strTNameO As String
        Dim intRunID As Int16
        Dim strDECISIONREASON As String
        Dim boolExFromAS As Boolean

        Dim strFnr1 As String
        Dim strFnr2 As String
        Dim rowsEE() As DataRow

        Dim rows2All() As DataRow
        Dim nE As Short
        Dim nI As Short

        Dim int12 As Short

        Dim numPrec As Single
        Dim numTheor As Single

        Dim boolHasOutlier As Boolean = False

        'make a table for different stats
        Dim dtblStats As New DataTable
        For Count1 = 1 To 8
            Dim col As New DataColumn
            Select Case Count1
                Case 1
                    str1 = "numMean"
                Case 2
                    str1 = "numSD"
                Case 3
                    str1 = "numPrec"
                Case 4
                    str1 = "numAcc"
                Case 5
                    str1 = "intN"
                Case 6
                    str1 = "boolAll"
                Case 7
                    str1 = "QCLevel"
                Case 8
                    str1 = "RunID"
            End Select
            col.ColumnName = str1
            dtblStats.Columns.Add(str1)
        Next

        'make a table for Accuracy for boolDiffCol
        Dim dtblAccDiffCol As New DataTable
        For Count1 = 1 To 5
            Dim col As New DataColumn
            Select Case Count1
                Case 1
                    str1 = "numAcc"
                Case 2
                    str1 = "boolOut"
                Case 3
                    str1 = "QCLevel"
                Case 4
                    str1 = "RunID"
                Case 5
                    str1 = "BOOLOUTLIER"
            End Select
            col.ColumnName = str1
            dtblAccDiffCol.Columns.Add(str1)
        Next

        boolJustTable = False

        Cursor.Current = Cursors.WaitCursor

        ''wdd.visible = True

        Dim fonts
        fontsize = wd.ActiveDocument.Styles("Normal").Font.Size ' wd.Selection.Font.Size
        fonts = fontsize ' wd.Selection.Font.Size

        Dim rowsFC() As DataRow = tblReportTable.Select("ID_TBLREPORTTABLE = " & idTR)
        charFCID = NZ(rowsFC(0).Item("CHARFCID"), "NA")

        With wd

            'dvDo = frmH.dgvReportTableConfiguration.DataSource
            'strTName = "Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision"
            'intDo = FindRowDVByCol(strTName, dvDo, "Table")

            intTableID = 11

            Dim strWRunId As String = GetWatsonColH(intTableID)

            dvDo = frmH.dgvReportTableConfiguration.DataSource
            strF = "id_tblconfigreporttables = " & intTableID
            intDo = FindRowDVNumByCol(intTableID, dvDo, "id_tblconfigreporttables")

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
            strTName = NZ(var1, "[NONE]")

            strTNameO = strTName


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

            tbl1 = tblAnalysisResultsHome
            tbl2 = tblAssignedSamples
            tbl3 = tblAssignedSamplesHelper
            tbl4 = tblAnalytesHome

            'ensure data has been entered
            strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idTR
            rowsX = tbl2.Select(strF)
            'If rowsX.Length = 0 Then
            '    strM = "Creating Summary of Interpolated QC Standard Concentrations Table ...."
            '    frmH.lblProgress.Text = strM
            '    frmH.Refresh()
            '    MsgBox("Samples have not been assigned to this table.", MsgBoxStyle.Information, "Samples have not been assigned...")
            '    GoTo end2
            'End If


            strF = "IsIntStd = 'No'"
            strS = "INTORDER ASC, IsIntStd ASC, AnalyteDescription ASC"
            rows11 = tblAnalytesHome.Select(strF, strS)
            intRowsAnal = rows11.Length

            'build tblAnova
            tblAnova.Columns.Add("GROUP", Type.GetType("System.Int16"))
            tblAnova.Columns.Add("Conc", Type.GetType("System.Decimal"))
            tblAnova.Columns.Add("NomConc", Type.GetType("System.Decimal"))
            tblAnova.Columns.Add("ELIMINATEDFLAG", Type.GetType("System.String"))
            tblAnova.Columns.Add("BOOLEXCLSAMPLE", Type.GetType("System.Int16"))

            'var1 = NZ(rows1(Count4).Item("ELIMINATEDFLAG"), "N")
            'var3 = NZ(rows1(Count4).Item("BOOLEXCLSAMPLE"), "N")

            Dim tblIntraSum As New System.Data.DataTable
            Dim rowsTIS() As DataRow

            'build tblz to record stats info
            tblIntraSum.Columns.Add("GROUP", Type.GetType("System.Int16"))
            tblIntraSum.Columns.Add("Conc", Type.GetType("System.Decimal"))
            tblIntraSum.Columns.Add("NomConc", Type.GetType("System.Decimal"))
            tblIntraSum.Columns.Add("intSetNum", Type.GetType("System.Int16"))
            tblIntraSum.Columns.Add("boolAll", Type.GetType("System.Boolean"))

            For Count1 = 1 To intRowsAnal

                boolHasOutlier = False

                boolJustTable = False

                Dim arrLegend(4, 20)

                strTName = strTNameO

                ctLegend = 0

                Dim int11 As Short
                If boolSTATSDIFFCOL Then
                    int11 = 2
                Else
                    int11 = 1
                End If

                tblAnova.Clear()

                'check if table is to be generated
                'strDo = arrAnalytes(1, Count1) 'record column name
                strDo = rows11(Count1 - 1).Item("ANALYTEDESCRIPTION")

                If UseAnalyte(CStr(strDo)) Then
                Else
                    GoTo next1
                End If

                bool = dvDo.Item(intDo).Item(strDo) 'find boolean value of dvDo column

                Dim strM1 As String
                If bool Then 'continue

                    intTCur = intTCur + 1

                    'ensure data has been entered
                    strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & CleanText(strDo) & "' AND ID_TBLREPORTTABLE = " & idTR
                    rowsX = tbl2.Select(strF)

                    'setup tables
                    If boolUseGroups Then
                        intGroup = tblAG.Rows(Count1 - 1).Item("INTGROUP")
                        strAnal = tblAG.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                        strAnalC = tblAG.Rows(Count1 - 1).Item("ANALYTEDESCRIPTION_C")
                        vAnalyteID = tblAG.Rows.Item(Count1 - 1).Item("ANALYTEID")
                        strMatrix = tblAG.Rows(Count1 - 1).Item("MATRIX")
                    Else
                        var1 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteIndex")
                        var2 = tbl4.Rows.Item(Count1 - 1).Item("MasterAssayID")
                        var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                        var4 = tbl4.Rows.Item(Count1 - 1).Item("ANALYTEID")
                        intGroup = 0
                        vAnalyteIndex = var1
                        vMasterAssayID = var2
                        vAnalyteID = var4
                        strMatrix = ""
                    End If

                    If rowsX.Length = 0 Then
                        strM = "Creating " & strTName & "...."
                        strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                        frmH.lblProgress.Text = strM
                        frmH.Refresh()
                        'MsgBox("Samples have not been assigned to this table.", MsgBoxStyle.Information, "Samples have not been assigned...")
                        boolJustTable = True
                        'page setup according to configuration
                        str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                        'insert page break
                        'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)


                        Call InsertPageBreak(wd)
                        Call PageSetup(wd, str1) 'L=Landscape, P=Portrait

                        GoTo end1
                    Else
                        boolJustTable = False
                    End If


                    'for legend stuff
                    intExp = 0
                    ctExp = 0

                    intLeg = 0
                    ctQCLegend = 0
                    ctDilLeg = 0
                    ctLegend = 0
                    strA = ""
                    strB = ""
                    arrLegend.Clear(arrLegend, 0, arrLegend.Length)
                    arrFP.Clear(arrFP, 0, arrFP.Length)
                    intLegStart = 96

                    'page setup according to configuration
                    Dim strOrientation As String
                    strOrientation = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                    'insert page break
                    'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)

                    Call InsertPageBreak(wd)
                    Call PageSetup(wd, strOrientation) 'L=Landscape, P=Portrait

                    'ReDim arrBCQCs(8, 50) '1=LevelNumber, 2=NomConcentration, 3=ID, 4=FlagPercent, 5=Hello, 6=Lo, 7=#ofReplicates, 8=ASSAYID
                    strM = "Creating " & strTName & " For " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & "..."
                    strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    strM1 = strM
                    frmH.lblProgress.Text = strM
                    frmH.Refresh()

                    ''get strConcUnits
                    'int1 = FindRowDV("ULOQ Units", frmH.dgvWatsonAnalRef.DataSource)
                    'strConcUnits = NZ(frmH.dgvWatsonAnalRef(Count1, int1).Value, "ng/mL")

                    'int1 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
                    'str1 = NZ(frmH.dgvStudyConfig(1, int1).Value, "")

                    'If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
                    'Else
                    '    strConcUnits = str1
                    'End If

                    'setup tables
                    'Legend: Dim arrAnalytes(14, 51) '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
                    '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
                    ''10=UseIntStd, 11=IntStd, 12=MasterAssayID,13=IsCoadministeredCmpd,14=Original Analyte Description

                    var1 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteID")
                    vAnalyteID = var1
                    var1 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteIndex")
                    vAnalyteIndex = var1
                    var2 = tbl4.Rows.Item(Count1 - 1).Item("MasterAssayID")
                    vMasterAssayID = var2
                    var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                    strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                    strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                    strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                    strF2 = strF2 & "ANALYTEINDEX = " & var1 & " AND "
                    strF2 = strF2 & "MASTERASSAYID = " & var2 & " AND "
                    strF2 = strF2 & "CHARANALYTE = '" & CleanText(CStr(var3)) & "'"
                    'strF2 = strF2 & "BOOLINTSTD = 0"

                    If boolUseGroups Then
                        strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                        strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                        strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                        strF2 = strF2 & "INTGROUP = " & intGroup
                    Else
                        strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                        strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                        strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                        strF2 = strF2 & "ANALYTEINDEX = " & var1 & " AND "
                        strF2 = strF2 & "MASTERASSAYID = " & var2 & " AND "
                        strF2 = strF2 & "CHARANALYTE = '" & CleanText(CStr(var3)) & "'"
                        'strF2 = strF2 & "ANALYTEID = " & var4 ' & "' AND "
                        'strF2 = strF2 & "BOOLINTSTD = 0"
                    End If

                    strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                    rows2 = tbl2.Select(strF2, strS)
                    int1 = rows2.Length 'debug
                    dv2 = New DataView(tbl2, strF2, strS, DataViewRowState.CurrentRows)
                    int1 = dv2.Count 'debug

                    'find number of runs used
                    tblNumRuns = dv2.ToTable("a", True, "RUNID")
                    intNumRuns = tblNumRuns.Rows.Count

                    'get strConcUnits
                    intRunID = 0
                    int1 = 0
                    Do Until intRunID > 0
                        var1 = tblNumRuns.Rows(int1).Item("RUNID")
                        If IsDBNull(var1) Then
                        Else
                            intRunID = var1
                        End If
                        int1 = int1 + 1
                    Loop
                    strConcUnits = GetConcUnits(intRunID)

                    'establish table of level numbers
                    'must be sorted by nomconc!
                    'make new dv
                    Dim dvNL As New DataView(tbl2, strF2, "NOMCONC ASC", DataViewRowState.CurrentRows)
                    tblLevels = dvNL.ToTable("b", True, "NOMCONC", "CHARHELPER1")
                    intNumLevels = tblLevels.Rows.Count

                    Dim intRowsXTot As Short = 0

                    'find number of table rows to generate
                    intRowsX = 0
                    For Count2 = 0 To intNumRuns - 1
                        '.Selection.Tables.item(1).Cell(int1, 1).Select()
                        'enter runid
                        var10 = tblNumRuns.Rows.Item(Count2).Item("RUNID")
                        '.Selection.TypeText(CStr(var10))
                        '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)

                        intRowsX = 0
                        For Count3 = 0 To intNumLevels - 1
                            varNom = tblLevels.Rows.Item(Count3).Item("NOMCONC")
                            dv2.RowFilter = ""
                            'don't know why, but must make a long filter here or
                            'both analytes get returned in dv2.rowfilter
                            strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom
                            dv2.RowFilter = strF
                            int2 = dv2.Count
                            If int2 > intRowsX Then
                                intRowsX = int2
                            End If
                        Next

                        intRowsXTot = intRowsXTot + intRowsX

                    Next

                    'generate table
                    intTblRows = 0
                    intTblRows = intTblRows + 3 'for header
                    intTblRows = intTblRows + 1 'for blank row
                    intTblRows = intTblRows + (intRowsXTot) 'for number of data rows
                    intTblRows = intTblRows + (1 * intNumRuns) 'for a blank row after each run set
                    'intTblRows = intTblRows + (3 * intNumRuns) 'for Mean/Bias/n section for each run set

                    'Increment for Statistics Sections
                    Dim intCSN As Short
                    intCSN = countNumStatsRows()
                    intTblRows = intTblRows + (intCSN * intNumRuns)

                    If intCSN > 0 Then
                        intTblRows = intTblRows + (1 * intNumRuns) - 1 'for a blank row after each Mean/Bias/n set, except last set
                    Else
                        intTblRows = intTblRows - 1 'subtract an unneeded blank row
                    End If

                    'BOOLINTRARUNSUMSTATS
                    'boolINCLANOVASUMSTATS: this is inter-run stats

                    If boolINCLANOVA Or boolINCLANOVASUMSTATS Or BOOLINTRARUNSUMSTATS Then
                        intTblRows = intTblRows + 2 'for Summary Stats heading
                    End If

                    If boolINCLANOVA Then
                        intTblRows = intTblRows + 6 ' 14 '9 '8 'for Anonva section
                    End If

                    If boolINCLANOVASUMSTATS Or BOOLINTRARUNSUMSTATS Then
                        intTblRows = intTblRows + 6 ' 8 '9 '8 'for between run section
                    End If

                    If boolINCLANOVASUMSTATS Or BOOLINTRARUNSUMSTATS Then
                        intTblRows = intTblRows + 6 ' 8 '9 '8 'for between run section
                    End If

                    If boolINCLANOVA = False And boolINCLANOVASUMSTATS = False And BOOLINTRARUNSUMSTATS = False Then
                        intTblRows = intTblRows + 1
                    End If

                    wrdSelection = wd.Selection()

                    Dim intCols As Short
                    If boolSTATSDIFFCOL Then
                        intCols = (intNumLevels * 2) + 1
                    Else
                        intCols = intNumLevels + 1
                    End If

                    Dim tblLevelCrit As New DataTable
                    Dim col1 As New DataColumn
                    col1.ColumnName = "NomConc"
                    col1.DataType = System.Type.GetType("System.Decimal")
                    tblLevelCrit.Columns.Add(col1)
                    Dim col2 As New DataColumn
                    col2.ColumnName = "Crit"
                    col2.DataType = System.Type.GetType("System.Decimal")
                    tblLevelCrit.Columns.Add(col2)


                    Try


                        '20180913 LEE:
                        Call IncrNextTableNumber(wd)

                        If boolPlaceHolder Then
                            .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=1, NumColumns:=1, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                        Else
                            .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=intTblRows, NumColumns:=intCols, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                        End If

                        .Selection.Tables.Item(1).Select()

                        Call SetCellPaddingZero(.Selection.Tables.Item(1))

                        .Selection.Font.Size = 10
                        .Selection.Tables.Item(1).Rows.AllowBreakAcrossPages = False
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                        .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
                        .Selection.Tables.Item(1).Columns.PreferredWidth = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPoints
                        '.Selection.Tables.Item(1).Columns.Item(1).Width = 86
                        For Count2 = 1 To intNumLevels
                            '.Selection.Tables.item(1).Columns.item(Count2 + 1).Width = 50
                        Next
                        .Selection.Tables.Item(1).Select()

                        'remove border, but leave top and bottom
                        removeBorderButLeaveTopAndBottom(wd)

                        'border top and bottom of range
                        .Selection.Tables.Item(1).Cell(1, 1).Select()


                        If boolPlaceHolder Then

                            .Selection.Tables.Item(1).Select()
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone


                            strA = arrAnalytes(14, Count1)
                            If gNumMatrix = 1 Then
                                strA = strAnalC
                            Else
                                strA = strAnal 'strAnalC has '..Matrix', don't want to pass that here
                            End If
                            'No, just strAnal
                            strA = strAnal
                            strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, False)
                            Call EnterTableNumber(wd, strTName, 3, strA, strTempInfo, intTableID, intGroup, idTR)
                            Call MoveOneCellDown(wd)

                            .Selection.TypeParagraph()
                            .Selection.TypeParagraph()

                            'enter a table record in tblTableN
                            'ctTableN = ctTableN + 1
                            Dim dtblr1 As DataRow = tblTableN.NewRow
                            dtblr1.BeginEdit()
                            dtblr1.Item("TableNumber") = ctTableN
                            dtblr1.Item("AnalyteName") = arrAnalytes(1, Count1)
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


                        .Selection.SelectRow()
                        .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=2, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.MoveRight Unit:=Microsoft.Office.Interop.Word.wdunits.word.wdunits.wdCharacter, Count:=ctQCs, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                        '.Selection.MoveLeft(Microsoft.Office.Interop.Word.WdUnits.wdCharacter, 1)

                        'Enter row titles
                        .Selection.Tables.Item(1).Cell(2, 2).Select()
                        For Count2 = 0 To intNumLevels - 1
                            'var1 = arrBCQCs(3, Count2)
                            var1 = NZ(tblLevels.Rows.Item(Count2).Item("NOMCONC"), "")
                            var2 = NZ(tblLevels.Rows.Item(Count2).Item("CHARHELPER1"), "")
                            '.Selection.TypeText(Text:=var1)

                            '******determine if the level is a diln level
                            Dim strE As String
                            var3 = ReturnStdQC(var2.ToString) ' var2 ' & ChrW(10) & var1 & " " & strConcUnits
                            'strE = ChrW(10) & var1 & " " & strConcUnits
                            .Selection.TypeText(Text:=var3)

                            dv2.RowFilter = ""
                            strF = strF2 & " AND NOMCONC = " & CDbl(var1)
                            dv2.RowFilter = strF
                            'check for aliquot factor
                            Dim numDS As Single
                            If dv2.Count = 0 Then

                            Else
                                numDS = dv2(0).Item("ALIQUOTFACTOR")
                                If numDS <> 1 Then
                                    'record legend
                                    'var1 = NZ(DilQCFactor(Count2), 1)
                                    intLeg = intLeg + 1
                                    ctDilLeg = ctDilLeg + 1
                                    ctLegend = ctLegend + 1
                                    'configure first legend item
                                    'var4 = numDSNomConc 'tblLevels.Rows(Count2 - 1).Item("NOMCONC")
                                    'var3 = Format(1 / CDec(numDS), "0")
                                    var3 = GetDilnFactor(CDec(numDS)) '20190220 LEE
                                    strA = Chr(96 + intLeg) 'debugging
                                    arrLegend(1, intLeg) = Chr(96 + intLeg) 'a,b,c,etc
                                    'var: units
                                    Dim strAN As String = GetAN(var3)

                                    If boolLUseSigFigs Then
                                        'arrLegend(2, intLeg) = "Dilution QCs undiluted concentration " & DisplayNum(SigFigOrDec(Val(var1), LSigFig, False), LSigFig, False) & " " & strConcUnits & "; a 1:" & var3 & " dilution with blank matrix was performed prior to extraction and analysis."
                                        arrLegend(2, intLeg) = "Dilution QCs undiluted concentration " & DisplayNum(SigFigOrDec(Val(var1), LSigFig, False), LSigFig, False) & " " & strConcUnits & "; " & strAN & " " & var3 & "-fold dilution with blank matrix was performed prior to extraction and analysis."
                                    Else
                                        'arrLegend(2, intLeg) = "Dilution QCs undiluted concentration " & Format(CDbl(Val(var1)), LSigFig) & " " & strConcUnits & "; a 1:" & var3 & " dilution with blank matrix was performed prior to extraction and analysis."
                                        arrLegend(2, intLeg) = "Dilution QCs undiluted concentration " & Format(CDbl(Val(var1)), LSigFig) & " " & strConcUnits & "; " & strAN & " " & var3 & "-fold dilution with blank matrix was performed prior to extraction and analysis."
                                    End If

                                    arrLegend(3, intLeg) = True
                                    arrLegend(4, intLeg) = True

                                    'enter superscript

                                    If boolRedBoldFont Then
                                        .Selection.Font.Bold = True
                                        .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                    End If

                                    Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                    .Selection.Font.Bold = False
                                    .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic

                                End If

                            End If
                            '.Selection.TypeText(strE)

                            '******

                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=int11)
                        Next

                        'Enter nom. conc. row titles
                        .Selection.Tables.Item(1).Cell(3, 2).Select()
                        For Count2 = 0 To intNumLevels - 1
                            'var1 = arrBCQCs(3, Count2)
                            If boolLUseSigFigs Then
                                var1 = CStr(DisplayNum(SigFigOrDec(tblLevels.Rows.Item(Count2).Item("NOMCONC"), LSigFig, False), LSigFig, False))
                            Else
                                var1 = CStr(Format(RoundToDecimalRAFZ(tblLevels.Rows.Item(Count2).Item("NOMCONC"), LSigFig), GetRegrDecStr(LSigFig)))
                            End If

                            If LboolNomConcParen Then
                                If StrComp(strOrientation, "P", CompareMethod.Text) = 0 Then
                                    If boolSTATSDIFFCOL Then
                                        var1 = "(" & var1 & ChrW(160) & strConcUnits & ")"
                                    Else
                                        var1 = "(" & var1 & ChrW(160) & strConcUnits & ")"
                                    End If
                                Else
                                    var1 = "(" & var1 & ChrW(160) & strConcUnits & ")"
                                End If
                            Else
                                var1 = var1 & ChrW(160) & strConcUnits
                            End If

                            .Selection.TypeText(Text:=var1)
                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                            If boolSTATSDIFFCOL Then
                                .Selection.TypeText(Text:=ReturnDiffLabel)
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                            End If
                        Next

                        .Selection.Tables.Item(1).Cell(3, 1).Select()
                        'bottom border this row
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .Selection.Tables.Item(1).Cell(3, 1).Select()

                        'begin entering data'

                        If BOOLINCLUDEDATE Then
                            .Selection.Tables.Item(1).Cell(2, 1).Select()
                            .Selection.TypeText(strWRunId)
                            .Selection.Tables.Item(1).Cell(3, 1).Select()
                            '.Selection.TypeText("(Analysis Date)")
                            '20180420 LEE:
                            .Selection.TypeText("(" & GetAnalysisDateLabel(intTableID) & ")")
                        Else
                            .Selection.TypeText(strWRunId)
                        End If

                        int1 = 5 'row position counter
                        'intLeg = 0
                        'ctQCLegend = 0
                        'ctDilLeg = 0
                        'ctLegend = 0
                        'strA = ""
                        'strB = ""

                        ''''''wdd.visible = True

                        For Count2 = 0 To intNumRuns - 1

                            .Selection.Tables.Item(1).Cell(int1, 1).Select()
                            'enter runid
                            var10 = tblNumRuns.Rows.Item(Count2).Item("RUNID")

                            'strM = "Creating Summary of Interpolated QC Standard Concentrations Table For " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & "..."
                            frmH.lblProgress.Text = strM1 & ChrW(10) & "Processing Run ID " & var10
                            frmH.Refresh()

                            .Selection.TypeText(CStr(var10))
                            If BOOLINCLUDEDATE Then
                                .Selection.Tables.Item(1).Cell(int1 + 1, 1).Select()
                                str1 = GetDateFromRunID(NZ(var10, 0), LDateFormat, intGroup, idTR)
                                .Selection.TypeText("(" & str1 & ")")
                                .Selection.Tables.Item(1).Cell(int1, 1).Select()
                            End If
                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)

                            'determine intRowsX for this runid

                            intRowsX = 0
                            For Count3 = 0 To intNumLevels - 1
                                varNom = tblLevels.Rows.Item(Count3).Item("NOMCONC")
                                dv2.RowFilter = ""
                                'don't know why, but must make a long filter here or
                                'both analytes get returned in dv2.rowfilter
                                strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom
                                dv2.RowFilter = strF
                                int2 = dv2.Count
                                If int2 > intRowsX Then
                                    intRowsX = int2
                                End If
                            Next



                            'start filling in data by columns
                            'intRowsX = 0
                            int12 = -1


                            For Count3 = 0 To intNumLevels - 1

                                Dim boolDoExtraStats As Boolean = False

                                int12 = int12 + 1
                                intN = 0

                                varNom = tblLevels.Rows.Item(Count3).Item("NOMCONC")

                                'determine hi and lo (nom*flagpercent)
                                'get AssayID from tblBCStdConcs
                                'strF = "CONCENTRATION = " & varNom & " AND ANALYTEID = " & vAnalyteID & " AND MASTERASSAYID = " & vMasterAssayID & " AND ANALYTEINDEX = " & vAnalyteIndex & " AND RUNID = " & var10 & " AND FLAGPERCENT IS NOT NULL"
                                'strF = "CONCENTRATION = #" & varNom & "# AND ANALYTEID = " & vAnalyteID & " AND MASTERASSAYID = " & vMasterAssayID & " AND ANALYTEINDEX = " & vAnalyteIndex & " AND RUNID = " & var10
                                'this strF is super goofy
                                'if Conc < 1, then the query return 0 records
                                'must do something different
                                strF = "ANALYTEID = " & vAnalyteID & " AND MASTERASSAYID = " & vMasterAssayID & " AND ANALYTEINDEX = " & vAnalyteIndex & " AND RUNID = " & var10 & " AND ANALYTEFLAGPERCENT IS NOT NULL"
                                'do not do 'ANALYTEFLAGPERCENT IS NOT NULL', sometimes user forgets to assign flag
                                strF = "ANALYTEID = " & vAnalyteID & " AND MASTERASSAYID = " & vMasterAssayID & " AND ANALYTEINDEX = " & vAnalyteIndex & " AND RUNID = " & var10 ' & " AND ANALYTEFLAGPERCENT IS NOT NULL"

                                If boolUseGroups Then
                                    strF = "ANALYTEID = " & vAnalyteID & " AND RUNID = " & var10
                                Else
                                    strF = "ANALYTEID = " & vAnalyteID & " AND MASTERASSAYID = " & vMasterAssayID & " AND ANALYTEINDEX = " & vAnalyteIndex & " AND RUNID = " & var10
                                End If

                                'if Conc < 1, then the query return 0 records
                                'must do something different
                                var1 = GetANALYTEFLAGPERCENTAnova(varNom, var10, vAnalyteID, tblLevelCrit)

                                var1 = CDec(var1)
                                arrFP(1, int12) = var1
                                arrFP(2, int12) = var1
                                v1 = var1
                                v2 = var1
                                vU = 0

                                Call SetHighAndLowCriteria(varNom, var1, var1, hi, lo)

                                'start entering data
                                dv2.RowFilter = ""
                                'don't know why, but must make a long filter here or
                                'both analytes get returned in dv2.rowfilter
                                strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom
                                dv2.RowFilter = strF
                                int2 = dv2.Count

                                If dv2.Count = 0 Then

                                    'fill in NA's
                                    If boolQCNA Then
                                        For Count4 = 0 To intRowsX - 1
                                            .Selection.Tables.Item(1).Cell(int1 + Count4, (Count3 * int11) + 2).Select()
                                            str1 = "NA"
                                            .Selection.TypeText(str1)
                                            If boolSTATSDIFFCOL Then
                                                .Selection.Tables.Item(1).Cell(int1 + Count4, (Count3 * int11) + 3).Select()
                                                str1 = "NA"
                                                .Selection.TypeText(str1)
                                            End If
                                        Next
                                    End If


                                    GoTo dvNo1
                                End If
                                'create rows1 from tbl1 which will contain data
                                strF = ""
                                strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                                Dim tbl2S As System.Data.DataTable = dv2.ToTable
                                rows1 = tbl2S.Select("RUNID > 0", strS)
                                ''20180716 LEE:
                                ''must account for null concentrations
                                'Try
                                '    rows1 = tbl2S.Select("RUNID > 0 AND CONCENTRATION IS NOT NULL", strS)
                                'Catch ex As Exception
                                '    var1 = var1
                                'End Try
                                rows2All = rows1
                                int3 = rows1.Length
                                nE = int3

                                'REDO hi/lo
                                If nE = 0 Then
                                    vU = -1
                                    v1 = -1 ' NZ(rows1(0).Item("NUMMAXACCCRIT"), 0)
                                    v2 = -1 ' NZ(rows1(0).Item("NUMMINACCCRIT"), 0)
                                    arrFP(1, int12) = v1
                                    arrFP(2, int12) = v2
                                    Call SetHighAndLowCriteria(varNom, v1, v2, hi, lo)
                                    var1 = var1
                                Else
                                    vU = rows1(0).Item("BOOLUSEGUWUACCCRIT")
                                    If vU = -1 Then
                                        v1 = NZ(rows1(0).Item("NUMMAXACCCRIT"), 0)
                                        v2 = NZ(rows1(0).Item("NUMMINACCCRIT"), 0)
                                        arrFP(1, int12) = v1
                                        arrFP(2, int12) = v2
                                        Call SetHighAndLowCriteria(varNom, v1, v2, hi, lo)
                                    End If
                                End If



                                ''''''wdd.visible = True


                                'now do rows actual
                                'strFActual = "(" & strF & ") AND (ELIMINATEDFLAG = 'N' OR BOOLEXCLSAMPLE = 0)"
                                If gAllowExclSamples And LAllowExclSamples Then
                                    'strFActual = "ELIMINATEDFLAG = 'N' AND BOOLEXCLSAMPLE = 0"
                                    strFActual = "(ELIMINATEDFLAG = 'N' OR ELIMINATEDFLAG IS NULL) AND BOOLEXCLSAMPLE = 0"
                                Else
                                    'strFActual = "ELIMINATEDFLAG = 'N'"
                                    strFActual = "(ELIMINATEDFLAG = 'N' OR ELIMINATEDFLAG IS NULL)"
                                End If
                                '20180716 LEE:
                                'must account for null concentrations
                                If gAllowExclSamples And LAllowExclSamples Then
                                    'strFActual = "ELIMINATEDFLAG = 'N' AND BOOLEXCLSAMPLE = 0"
                                    strFActual = "(ELIMINATEDFLAG = 'N' OR ELIMINATEDFLAG IS NULL) AND BOOLEXCLSAMPLE = 0 AND CONCENTRATION IS NOT NULL"
                                Else
                                    'strFActual = "ELIMINATEDFLAG = 'N'"
                                    strFActual = "(ELIMINATEDFLAG = 'N' OR ELIMINATEDFLAG IS NULL) AND CONCENTRATION IS NOT NULL"
                                End If
                                Try
                                    rowsActual = tbl2S.Select(strFActual, strS) 'this rows2E in Ad Hoc Stability
                                Catch ex As Exception
                                    var1 = var1
                                End Try
                                intN = rowsActual.Length


                                'just because gAllowExclSamples And LAllowExclSamples doesn't mean StudyDoc was actually used
                                'check for BOOLUSEGUWUACCCRIT
                                ''This has already been accounted for above
                                'If gAllowExclSamples And LAllowExclSamples Then
                                '    var1 = rowsActual(0).Item("BOOLUSEGUWUACCCRIT")
                                '    If var1 = 0 Then 'must sort again
                                '        Erase rowsActual
                                '        strFActual = "ELIMINATEDFLAG = 'N'"
                                '        rowsActual = tbl2S.Select(strFActual, strS)
                                '        intN = rowsActual.Length
                                '    End If
                                'End If

                                'enter data
                                Dim boolEnterDiff As Boolean

                                For Count4 = 0 To intRowsX - 1 'int3 - 1

                                    boolEnterDiff = True
                                    boolOC = False

                                    .Selection.Tables.Item(1).Cell(int1 + Count4, (Count3 * int11) + 2).Select()
                                    If Count4 > int3 - 1 Then

                                        If boolQCNA Then
                                            str1 = "NA"
                                        Else
                                            str1 = ""
                                        End If

                                        .Selection.TypeText(str1)
                                        boolEnterDiff = False
                                        boolOC = True
                                    Else
                                        var1 = rows1(Count4).Item("CONCENTRATION")
                                        varConc = var1
                                        var1 = NZ(var1, 0)
                                        numDF = rows1(Count4).Item("ALIQUOTFACTOR")
                                        var1 = var1 / numDF
                                        If boolLUseSigFigs Then
                                            var2 = SigFigOrDec(var1, LSigFig, False)
                                        Else
                                            var2 = RoundToDecimalRAFZ(var1, LSigFig)
                                        End If

                                        var1 = NZ(rows1(Count4).Item("ELIMINATEDFLAG"), "N")
                                        var3 = NZ(rows1(Count4).Item("BOOLEXCLSAMPLE"), 0)
                                        'add rows to tblAnova
                                        Dim rowsAn As DataRow = tblAnova.NewRow
                                        rowsAn.Item("GROUP") = var10
                                        rowsAn.Item("Conc") = var2
                                        rowsAn.Item("NomConc") = varNom
                                        If IsDBNull(varConc) Then
                                            var1 = "Y"
                                        End If
                                        rowsAn.Item("ELIMINATEDFLAG") = var1
                                        rowsAn.Item("BOOLEXCLSAMPLE") = var3

                                        tblAnova.Rows.Add(rowsAn)

                                        If IsDBNull(varConc) Then

                                            boolExFromAS = False
                                            boolOC = True

                                            If gAllowExclSamples And LAllowExclSamples Then
                                                If var3 = -1 Then
                                                    var1 = "Y"
                                                    boolExFromAS = True
                                                Else
                                                    'var1 = "N"
                                                    'don't assign "N", Watson may override
                                                End If
                                            End If

                                            var1 = "Y"

                                            intLeg = intLeg + 1
                                            strA = ChrW(intLeg + intLegStart)

                                            vU = NZ(rows1(Count4).Item("BOOLUSEGUWUACCCRIT"), 0)

                                            boolDoExtraStats = True

                                            'Remember, tblAssignedSamples does not have DECISIONREASON
                                            Dim var6
                                            var6 = "No Value: " & GetDECISIONREASONValue(boolExFromAS, vAnalyteID, rows1(Count4))

                                            'Set Legend String
                                            str1 = GetLegendStringExcluded(arrFP(1, int12), arrFP(2, int12), vU, var6, intTableID, True, "")
                                            boolHasOutlier = HasOutlier(str1, boolHasOutlier)

                                            'Add to Legend Array
                                            ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                            If boolRedBoldFont Then
                                                .Selection.Font.Bold = True
                                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                            End If

                                            .Selection.TypeText(Text:="NV")

                                            Call typeInSuperscriptFontSize12WithSpace(wd, strA)

                                            boolEnterDiff = True 'FALSE

                                        ElseIf StrComp(var1, "Y", vbTextCompare) = 0 Or (gAllowExclSamples And LAllowExclSamples And var3 = -1) Then

                                            boolExFromAS = False
                                            boolOC = True

                                            If gAllowExclSamples And LAllowExclSamples Then
                                                If var3 = -1 Then
                                                    var1 = "Y"
                                                    boolExFromAS = True
                                                Else
                                                    'var1 = "N"
                                                    'don't assign "N", Watson may override
                                                End If
                                            End If

                                            intLeg = intLeg + 1
                                            strA = ChrW(intLeg + intLegStart)

                                            var1 = "Y"

                                            vU = NZ(rows1(Count4).Item("BOOLUSEGUWUACCCRIT"), 0)

                                            boolDoExtraStats = True

                                            'Remember, tblAssignedSamples does not have DECISIONREASON
                                            Dim var6
                                            var6 = GetDECISIONREASONValue(boolExFromAS, vAnalyteID, rows1(Count4))

                                            'Set Legend String
                                            str1 = GetLegendStringExcluded(arrFP(1, int12), arrFP(2, int12), vU, var6, intTableID, True, "")
                                            boolHasOutlier = HasOutlier(str1, boolHasOutlier)

                                            'Add to Legend Array
                                            ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                            If boolRedBoldFont Then
                                                .Selection.Font.Bold = True
                                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                            End If


                                            If boolLUseSigFigs Then
                                                .Selection.TypeText(Text:=CStr(DisplayNum(var2, LSigFig, False)))
                                            Else
                                                .Selection.TypeText(Text:=CStr(Format(var2, GetRegrDecStr(LSigFig))))
                                            End If


                                            Call typeInSuperscriptFontSize12WithSpace(wd, strA)

                                            boolEnterDiff = True 'FALSE

                                        Else

                                            'determine if value is outside acceptance criteria
                                            vU = NZ(rows1(Count4).Item("BOOLUSEGUWUACCCRIT"), 0)
                                            If LAllowGuWuAccCrit And gAllowGuWuAccCrit And vU = -1 Then
                                                v1 = NZ(rows1(Count4).Item("NUMMAXACCCRIT"), 0)
                                                v2 = NZ(rows1(Count4).Item("NUMMINACCCRIT"), 0)
                                                Call SetHighAndLowCriteria(varNom, v1, v2, hi, lo)

                                                arrFP(1, int12) = v1
                                                arrFP(2, int12) = v2
                                            End If

                                            'If var2 > hi Or var2 < lo Then 'flag
                                            If OutsideAccCrit(var2, varNom, v1, v2, NZ(vU, 0)) Then

                                                intLeg = intLeg + 1
                                                strA = ChrW(intLeg + intLegStart)

                                                If LAllowGuWuAccCrit And gAllowGuWuAccCrit And vU = -1 Then
                                                    v1 = NZ(rows1(Count4).Item("NUMMAXACCCRIT"), 0)
                                                    v2 = NZ(rows1(Count4).Item("NUMMINACCCRIT"), 0)
                                                    If v1 = v2 Then
                                                        str1 = "Value outside of acceptance criteria (" & ChrW(177) & " " & v1 & " % theoretical) but included in summary statistics."
                                                    Else
                                                        str1 = "Value outside of acceptance criteria (+" & v1 & "/-" & v2 & " % theoretical) but included in summary statistics."
                                                    End If

                                                    'str1 = "Value outside of acceptance criteria (+" & v1 & "/-" & v2 & " % theoretical) but included in summary statistics."
                                                Else
                                                    str1 = "Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimal(arrFP(1, int12), 0) & "% theoretical) but included in summary statistics."
                                                End If

                                                'str1 = "Value outside of acceptance criteria (" & RoundToDecimal(arrFP(Count3), 0) & "% theoretical) but included in summary statistics."

                                                'Add to Legend Array
                                                ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)
                                                boolHasOutlier = HasOutlier(str1, boolHasOutlier)

                                                If boolRedBoldFont Then
                                                    .Selection.Font.Bold = True
                                                    .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                End If

                                                If boolLUseSigFigs Then
                                                    .Selection.TypeText(Text:=CStr(DisplayNum(var2, LSigFig, False)))
                                                Else
                                                    .Selection.TypeText(Text:=CStr(Format(var2, GetRegrDecStr(LSigFig))))
                                                End If


                                                Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                                boolEnterDiff = True
                                            Else
                                                If boolLUseSigFigs Then
                                                    .Selection.TypeText(Text:=CStr(DisplayNum(var2, LSigFig, False)))
                                                Else
                                                    .Selection.TypeText(Text:=CStr(Format(var2, GetRegrDecStr(LSigFig))))
                                                End If
                                                boolEnterDiff = True
                                            End If

                                        End If

                                    End If

                                    If boolSTATSDIFFCOL Then
                                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                        If boolEnterDiff Then
                                            var1 = SigFigOrDec(var2, LSigFig, False)
                                            'var3 = Format(((var1 / varNom) - 1) * 100, strQCDec)
                                            'var3 = Format(RoundToDecimal(((var1 / varNom) - 1) * 100, intQCDec), strQCDec)

                                            If boolTHEORETICAL Then
                                                var3 = CalcREPercent(var2, varNom, intQCDec)
                                                numTheor = 100 + CDec(var3)

                                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, Count3 + 1, "Accuracy", numTheor, CSng(var10), Count1, strDo, v1, v2, boolOC)
                                            Else
                                                var3 = Format(RoundToDecimal(CalcREPercent(var2, varNom, intQCDec), intQCDec), strQCDec)

                                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, Count3 + 1, "Accuracy", var3, CSng(var10), Count1, strDo, v1, v2, boolOC)
                                            End If

                                            'add to dtblAccDiffCol 
                                            'legend
                                            'Select Case Count1
                                            '    Case 1
                                            '        str1 = "numAcc"
                                            '    Case 2
                                            '        str1 = "boolAll"
                                            '    Case 3
                                            '        str1 = "QCLevel"
                                            '    Case 4
                                            '        str1 = "RunID"
                                            'End Select
                                            Dim nrD As DataRow = dtblAccDiffCol.NewRow
                                            nrD.BeginEdit()
                                            nrD.Item("numAcc") = var3
                                            nrD.Item("boolOut") = boolOC
                                            nrD.Item("QCLevel") = Count3 + 1
                                            nrD.Item("RunID") = CInt(var10)
                                            nrD.Item("BOOLOUTLIER") = boolOC.ToString
                                            nrD.EndEdit()
                                            dtblAccDiffCol.Rows.Add(nrD)

                                        Else

                                            If boolQCNA Then
                                                var3 = "NA"
                                            Else
                                                var3 = ""
                                            End If

                                        End If
                                        .Selection.TypeText(Text:=CStr(var3))
                                        .Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                    End If

                                    var1 = var1 'debug

                                Next

dvNo1:
                                'now enter Mean/Bias/n
                                If Count3 = 0 Then
                                    int8 = 0
                                    If boolSTATSMEAN Then
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, 1).Select()
                                        .Selection.TypeText("Intra-Run Mean:")
                                    End If
                                    If boolSTATSSD Then
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, 1).Select()
                                        .Selection.TypeText("Intra-Run S.D.:") '((Mean/NomConc)-1)*100)
                                    End If
                                    If boolSTATSCV Then
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, 1).Select()
                                        .Selection.TypeText("Intra-Run " & ReturnPrecLabel() & ":")
                                    End If
                                    If boolSTATSBIAS And boolSTATSMEAN Then
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, 1).Select()
                                        .Selection.TypeText("Intra-Run %Bias:")
                                    End If
                                    If boolTHEORETICAL And boolSTATSMEAN Then
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, 1).Select()
                                        .Selection.TypeText("Intra-Run %Theoretical:")
                                    End If
                                    If boolSTATSDIFF And boolSTATSMEAN Then
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, 1).Select()
                                        .Selection.TypeText("Intra-Run %Diff:")
                                    End If
                                    If BOOLSTATSRE And boolSTATSMEAN Then
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, 1).Select()
                                        .Selection.TypeText("Intra-Run %RE:")
                                    End If
                                    If boolSTATSN Then
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, 1).Select()
                                        .Selection.TypeText("n:")
                                    End If

                                End If


                                If dv2.Count = 0 Then
                                    GoTo dvNo2
                                End If

                                int8 = 0

                                Dim nr1 As DataRow = dtblStats.NewRow
                                nr1.BeginEdit()
                                nr1.Item("boolAll") = False
                                nr1.Item("RunID") = var10
                                nr1.Item("QCLevel") = Count3 '0-based
                                nr1.EndEdit()
                                dtblStats.Rows.Add(nr1)

                                Dim nr2 As DataRow = dtblStats.NewRow
                                nr2.BeginEdit()
                                nr2.Item("boolAll") = True
                                nr2.Item("RunID") = var10
                                nr2.Item("QCLevel") = Count3 '0-based
                                nr2.EndEdit()
                                dtblStats.Rows.Add(nr2)

                                strFnr1 = "boolAll = FALSE and RunID = " & var10 & " AND QCLevel = " & Count3
                                strFnr2 = "boolAll = TRUE and RunID = " & var10 & " AND QCLevel = " & Count3

                                v1 = arrFP(1, int12)
                                v2 = arrFP(2, int12)

                                Dim boolMean As Boolean = True
                                If boolSTATSMEAN Then
                                    Try
                                        'enter Mean
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, (Count3 * int11) + 2).Select()
                                        'var1 = MeanDR(rows1, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                        var1 = MeanDR(rowsActual, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                        If boolLUseSigFigs Then
                                            numMean = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                        Else
                                            numMean = RoundToDecimalRAFZ(var1, LSigFig)
                                        End If

                                        If nE = 0 Then
                                            boolMean = False
                                        Else

                                        End If

                                        rowsEE = dtblStats.Select(strFnr1)
                                        rowsEE(0).BeginEdit()
                                        rowsEE(0).Item("numMean") = numMean
                                        rowsEE(0).EndEdit()

                                        'now do all
                                        var1 = MeanDR(rows2All, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                        If boolLUseSigFigs Then
                                            var2 = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                        Else
                                            var2 = RoundToDecimalRAFZ(var1, LSigFig)
                                        End If
                                        rowsEE = dtblStats.Select(strFnr2)
                                        rowsEE(0).BeginEdit()
                                        rowsEE(0).Item("numMean") = var2
                                        rowsEE(0).EndEdit()

                                        '.Selection.TypeText(CStr(numMean))

                                        If intN = 0 Then
                                        Else
                                            Call InsertAnovaTables(intTableID, idTR, charFCID, False, True, varNom, Count3 + 1, "Mean", numMean, False, False, CSng(var10), Count1, strDo, 0, 0, False)
                                            'enter numMean in tblIntraSum
                                            Dim nTIS As DataRow = tblIntraSum.NewRow
                                            nTIS.BeginEdit()
                                            nTIS.Item("GROUP") = var10
                                            nTIS.Item("Conc") = numMean
                                            nTIS.Item("NomConc") = varNom
                                            nTIS.Item("intSetNum") = int12 + 1
                                            nTIS.Item("boolAll") = False
                                            nTIS.EndEdit()
                                            tblIntraSum.Rows.Add(nTIS)
                                        End If

                                        rowsEE = dtblStats.Select(strFnr2)
                                        var1 = rowsEE(0).Item("numMean") ' CDec(NZ(rowsEE(0).Item("numMean"), 0))
                                        If boolLUseSigFigs Then
                                            '.Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                            str2 = CStr(DisplayNum(var1, LSigFig, False))
                                        Else
                                            '.Selection.TypeText(Text:=CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                            str3 = GetRegrDecStr(LSigFig)
                                            'str2 = CStr(Format(var1, GetRegrDecStr(LSigFig)))
                                            str2 = Format(CDec(var1), GetRegrDecStr(LSigFig))
                                        End If

                                        Dim nTIS1 As DataRow = tblIntraSum.NewRow
                                        nTIS1.BeginEdit()
                                        nTIS1.Item("GROUP") = var10
                                        nTIS1.Item("Conc") = CDec(str2)
                                        nTIS1.Item("NomConc") = varNom
                                        nTIS1.Item("intSetNum") = int12 + 1
                                        nTIS1.Item("boolAll") = True
                                        nTIS1.EndEdit()
                                        tblIntraSum.Rows.Add(nTIS1)

                                        'vU = NZ(dv2(Count4).Item("BOOLUSEGUWUACCCRIT"), 0)
                                        If LAllowGuWuAccCrit And gAllowGuWuAccCrit And vU = -1 Then
                                            v1 = arrFP(1, int12) ' NZ(dv2(Count4).Item("NUMMAXACCCRIT"), 0)
                                            v2 = arrFP(2, int12) ' NZ(dv2(Count4).Item("NUMMINACCCRIT"), 0)
                                            Call SetHighAndLowCriteria(varNom, v1, v2, hi, lo)
                                        End If

                                        'determine if value is outside acceptance criteria
                                        'If (numMean > hi Or numMean < lo) And boolFootNoteQCMean Then 'flag
                                        If (OutsideAccCrit(numMean, varNom, v1, v2, NZ(vU, 0))) And boolFootNoteQCMean Then 'flag
                                            intLeg = intLeg + 1
                                            strA = ChrW(intLeg + intLegStart)

                                            'Set Legend String
                                            str1 = GetLegendStringIncluded(v1, v2, vU)
                                            boolHasOutlier = HasOutlier(str1, boolHasOutlier)
                                            'Add to Legend Array
                                            ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                            If boolRedBoldFont Then
                                                .Selection.Font.Bold = True
                                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                            End If

                                            If intN = 0 Then
                                                .Selection.TypeText(Text:="NA")
                                            Else
                                                If boolLUseSigFigs Then
                                                    .Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                                Else
                                                    .Selection.TypeText(Text:=CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                                End If

                                                Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                            End If

                                            boolEnterDiff = True

                                        Else

                                            If intN = 0 Then
                                                str1 = "NA"
                                            Else
                                                If boolLUseSigFigs Then
                                                    '.Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                                    str1 = CStr(DisplayNum(numMean, LSigFig, False))
                                                Else
                                                    '.Selection.TypeText(Text:=CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                                    str1 = CStr(Format(numMean, GetRegrDecStr(LSigFig)))
                                                End If
                                            End If

                                            If boolQCREPORTACCVALUES Or boolDoExtraStats = False Then
                                                .Selection.TypeText(Text:=str1)
                                            Else
                                                .Selection.TypeText(Text:=str1 & ChrW(160) & "(" & str2 & ")")
                                            End If

                                            boolEnterDiff = True
                                        End If

                                    Catch ex As Exception

                                    End Try

                                End If

                                If boolSTATSSD Then

                                    Try
                                        'enter Mean
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, (Count3 * int11) + 2).Select()

                                        '20170716 LEE: Depricate intN<gSDMax
                                        var1 = StdDevDR(rowsActual, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)

                                        If boolLUseSigFigs Then
                                            numSD = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                            '.Selection.TypeText(CStr(DisplayNum(numSD, LSigFig, False)))
                                            str1 = CStr(DisplayNum(numSD, LSigFig, False))
                                        Else
                                            numSD = RoundToDecimalRAFZ(var1, LSigFig)
                                            '.Selection.TypeText(CStr(Format(numSD, GetRegrDecStr(LSigFig))))
                                            str1 = CStr(Format(numSD, GetRegrDecStr(LSigFig)))
                                        End If

                                        rowsEE = dtblStats.Select(strFnr1)
                                        rowsEE(0).BeginEdit()
                                        rowsEE(0).Item("numSD") = numSD
                                        rowsEE(0).EndEdit()

                                        'now do all
                                        var1 = StdDevDR(rows2All, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                        If boolLUseSigFigs Then
                                            var2 = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                        Else
                                            var2 = RoundToDecimalRAFZ(var1, LSigFig)
                                        End If
                                        rowsEE = dtblStats.Select(strFnr2)
                                        rowsEE(0).BeginEdit()
                                        rowsEE(0).Item("numSD") = var2
                                        rowsEE(0).EndEdit()

                                        If intN < gSDMax Then
                                            Call InsertAnovaTables(intTableID, idTR, charFCID, False, True, varNom, Count3 + 1, "SD", 0, True, False, CSng(var10), Count1, strDo, 0, 0, False)
                                        Else
                                            Call InsertAnovaTables(intTableID, idTR, charFCID, False, True, varNom, Count3 + 1, "SD", numSD, False, False, CSng(var10), Count1, strDo, 0, 0, False)
                                        End If

                                        If intN < gSDMax Then
                                            str1 = "NA"
                                        End If

                                        If boolQCREPORTACCVALUES Or boolDoExtraStats = False Then
                                            .Selection.TypeText(Text:=str1)
                                        Else

                                            If nE < gSDMax Then
                                                str2 = "NA"
                                            Else
                                                rowsEE = dtblStats.Select(strFnr2)
                                                var1 = rowsEE(0).Item("numSD")
                                                If boolLUseSigFigs Then
                                                    '.Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                                    str2 = CStr(DisplayNum(var1, LSigFig, False))
                                                Else
                                                    '.Selection.TypeText(Text:=CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                                    str2 = CStr(Format(CDec(var1), GetRegrDecStr(LSigFig)))
                                                End If
                                            End If

                                            .Selection.TypeText(Text:=str1 & ChrW(160) & "(" & str2 & ")")
                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If

                                If boolSTATSCV Then
                                    Try
                                        'enter %CV
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, (Count3 * int11) + 2).Select()

                                        '20170716 LEE: Depricate intN<gSDMax

                                        numPrec = CalcCVPercent(numSD, numMean, intQCDec)
                                        '.Selection.TypeText(Format(numPrec, strQCDec))
                                        str1 = Format(numPrec, strQCDec)

                                        rowsEE = dtblStats.Select(strFnr1)
                                        rowsEE(0).BeginEdit()
                                        rowsEE(0).Item("numPrec") = numPrec
                                        rowsEE(0).EndEdit()

                                        'now do all
                                        rowsEE = dtblStats.Select(strFnr2)
                                        var1 = rowsEE(0).Item("numMean")
                                        var2 = rowsEE(0).Item("numSD")
                                        var3 = CalcCVPercent(var2, var1, intQCDec)
                                        rowsEE(0).BeginEdit()
                                        rowsEE(0).Item("numPrec") = CDec(var3)
                                        rowsEE(0).EndEdit()

                                        If intN = 0 Then
                                            Call InsertAnovaTables(intTableID, idTR, charFCID, False, True, varNom, Count3 + 1, "Precision", 0, True, False, CSng(var10), Count1, strDo, 0, 0, False)
                                        Else
                                            Call InsertAnovaTables(intTableID, idTR, charFCID, False, True, varNom, Count3 + 1, "Precision", numPrec, False, False, CSng(var10), Count1, strDo, 0, 0, False)
                                        End If

                                        If intN < gSDMax Then
                                            str1 = "NA"
                                        End If

                                        If boolQCREPORTACCVALUES Or boolDoExtraStats = False Then
                                            .Selection.TypeText(Text:=str1)
                                        Else
                                            If nE < gSDMax Then
                                                str2 = "NA"
                                            Else
                                                rowsEE = dtblStats.Select(strFnr2)
                                                var1 = rowsEE(0).Item("numPrec")
                                                str2 = Format(CDec(var1), strQCDec)
                                            End If

                                            .Selection.TypeText(Text:=str1 & ChrW(160) & "(" & str2 & ")")
                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If


                                'bias
                                '****
                                If boolSTATSBIAS Then
                                    If boolSTATSDIFFCOL And BOOLDIFFCOLSTATS Then
                                        'get average of diffcol
                                        'Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", var3, CSng(var10), Count1, strDo, v1, v2, FALSE)
                                        numBias = GetBiasFromDiffCol(idTR, varNom, Count3 + 1, var10, False)
                                    Else
                                        numBias = CalcREPercent(numMean, varNom, intQCDec)
                                        '.Selection.TypeText(Format(numBias, strQCDec))
                                        If intN = 0 Then
                                        Else
                                            Call InsertAnovaTables(intTableID, idTR, charFCID, False, True, varNom, Count3 + 1, "Accuracy", numBias, False, False, CSng(var10), Count1, strDo, 0, 0, False)
                                        End If
                                    End If
                                    str1 = Format(numBias, strQCDec)

                                    If intN = 0 Then
                                        str1 = "NA"
                                    End If
                                End If
                                '****

                                If boolSTATSBIAS And boolSTATSMEAN Then
                                    Try
                                        'enter %Bias
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, (Count3 * int11) + 2).Select()

                                        rowsEE = dtblStats.Select(strFnr1)
                                        rowsEE(0).BeginEdit()
                                        rowsEE(0).Item("numAcc") = numBias
                                        rowsEE(0).EndEdit()

                                        'now do all
                                        rowsEE = dtblStats.Select(strFnr2)
                                        var1 = rowsEE(0).Item("numMean")
                                        var2 = varNom
                                        var3 = CalcREPercent(var1, var2, intQCDec)
                                        rowsEE(0).BeginEdit()
                                        rowsEE(0).Item("numAcc") = CDec(var3)
                                        rowsEE(0).EndEdit()

                                        If boolQCREPORTACCVALUES Or boolDoExtraStats = False Then
                                            .Selection.TypeText(Text:=str1)
                                        Else

                                            rowsEE = dtblStats.Select(strFnr2)
                                            var1 = CDec(rowsEE(0).Item("numAcc"))
                                            str2 = Format(CDec(var1), strQCDec)
                                            If nE = 0 Then
                                                str2 = "NA"
                                            End If
                                            .Selection.TypeText(Text:=str1 & ChrW(160) & "(" & str2 & ")")

                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If

                                'If boolTHEORETICAL And boolSTATSDIFFCOL = False Then
                                '    If intN = 0 Then
                                '        str1 = "NA"
                                '    Else
                                '        numTheor = CalcREPercent(numMean, varNom, intQCDec)
                                '        numTheor = 100 + CDec(numTheor)
                                '        '.Selection.TypeText(Format(numTheor, strQCDec))
                                '        str1 = Format(numTheor, strQCDec)
                                '        Call InsertAnovaTables(intTableID, idTR, charFCID, False, True, varNom, Count3 + 1, "Accuracy", numTheor, False, False, CSng(var10), Count1, strDo, 0, 0, FALSE)
                                '    End If

                                'End If

                                '*****
                                If boolTHEORETICAL Then
                                    If boolSTATSDIFFCOL And BOOLDIFFCOLSTATS Then
                                        'get average of diffcol
                                        'Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", var3, CSng(var10), Count1, strDo, v1, v2, FALSE)
                                        numTheor = GetBiasFromDiffCol(idTR, varNom, Count3 + 1, var10, False)
                                        numTheor = 100 + CDec(numTheor)
                                    Else
                                        numTheor = CalcREPercent(numMean, varNom, intQCDec)
                                        numTheor = 100 + CDec(numTheor)
                                        '.Selection.TypeText(Format(numBias, strQCDec))
                                        If intN = 0 Then
                                        Else
                                            Call InsertAnovaTables(intTableID, idTR, charFCID, False, True, varNom, Count3 + 1, "Accuracy", numTheor, False, False, CSng(var10), Count1, strDo, 0, 0, False)
                                        End If
                                    End If
                                    If intN = 0 Then
                                        str1 = "NA"
                                    Else
                                        str1 = Format(numTheor, strQCDec)
                                    End If
                                End If
                                '*****

                                If boolTHEORETICAL And boolSTATSMEAN Then
                                    Try
                                        'enter %Theor
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, (Count3 * int11) + 2).Select()

                                        rowsEE = dtblStats.Select(strFnr1)
                                        rowsEE(0).BeginEdit()
                                        rowsEE(0).Item("numAcc") = numTheor
                                        rowsEE(0).EndEdit()

                                        'now do all
                                        rowsEE = dtblStats.Select(strFnr2)
                                        var1 = rowsEE(0).Item("numMean")
                                        var2 = varNom
                                        var3 = CalcREPercent(var1, var2, intQCDec)
                                        rowsEE(0).BeginEdit()
                                        rowsEE(0).Item("numAcc") = CDec(var3)
                                        rowsEE(0).EndEdit()

                                        If boolQCREPORTACCVALUES Or boolDoExtraStats = False Then
                                            .Selection.TypeText(Text:=str1)
                                        Else

                                            rowsEE = dtblStats.Select(strFnr2)
                                            var1 = CDec(rowsEE(0).Item("numAcc"))
                                            str2 = Format(CDec(var1), strQCDec)
                                            If nE = 0 Then
                                                str2 = "NA"
                                            End If
                                            .Selection.TypeText(Text:=str1 & ChrW(160) & "(" & str2 & ")")

                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If

                                'If boolSTATSDIFF And boolSTATSDIFFCOL = False Then
                                '    If intN = 0 Then
                                '        str1 = "NA"
                                '    Else
                                '        numBias = CalcREPercent(numMean, varNom, intQCDec)
                                '        str1 = Format(numBias, strQCDec)
                                '        Call InsertAnovaTables(intTableID, idTR, charFCID, False, True, varNom, Count3 + 1, "Accuracy", numBias, False, False, CSng(var10), Count1, strDo, 0, 0, FALSE)
                                '    End If

                                'End If

                                '****
                                If boolSTATSDIFF Then
                                    If boolSTATSDIFFCOL And BOOLDIFFCOLSTATS Then
                                        'get average of diffcol
                                        'Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", var3, CSng(var10), Count1, strDo, v1, v2, FALSE)
                                        numBias = GetBiasFromDiffCol(idTR, varNom, Count3 + 1, var10, False)
                                    Else
                                        numBias = CalcREPercent(numMean, varNom, intQCDec)
                                        '.Selection.TypeText(Format(numBias, strQCDec))
                                        If intN = 0 Then
                                        Else
                                            Call InsertAnovaTables(intTableID, idTR, charFCID, False, True, varNom, Count3 + 1, "Accuracy", numBias, False, False, CSng(var10), Count1, strDo, 0, 0, False)
                                        End If
                                    End If
                                    If intN = 0 Then
                                        str1 = "NA"
                                    Else
                                        str1 = Format(numBias, strQCDec)
                                    End If
                                End If
                                '****

                                If boolSTATSDIFF And boolSTATSMEAN Then
                                    Try
                                        'enter %Diff
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, (Count3 * int11) + 2).Select()

                                        rowsEE = dtblStats.Select(strFnr1)
                                        rowsEE(0).BeginEdit()
                                        rowsEE(0).Item("numAcc") = numBias
                                        rowsEE(0).EndEdit()

                                        'now do all
                                        rowsEE = dtblStats.Select(strFnr2)
                                        var1 = rowsEE(0).Item("numMean")
                                        var2 = varNom
                                        var3 = CalcREPercent(var1, var2, intQCDec)
                                        rowsEE(0).BeginEdit()
                                        rowsEE(0).Item("numAcc") = CDec(var3)
                                        rowsEE(0).EndEdit()

                                        If boolQCREPORTACCVALUES Or boolDoExtraStats = False Then
                                            .Selection.TypeText(Text:=str1)
                                        Else

                                            rowsEE = dtblStats.Select(strFnr2)
                                            var1 = CDec(rowsEE(0).Item("numAcc"))
                                            str2 = Format(var1, strQCDec)
                                            If nE = 0 Then
                                                str2 = "NA"
                                            End If
                                            .Selection.TypeText(Text:=str1 & ChrW(160) & "(" & str2 & ")")

                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If

                                'If BOOLSTATSRE And boolSTATSDIFFCOL = False Then
                                '    If intN = 0 Then
                                '        str1 = "NA"
                                '    Else
                                '        numBias = CalcREPercent(numMean, varNom, intQCDec)
                                '        str1 = Format(numBias, strQCDec)
                                '        Call InsertAnovaTables(intTableID, idTR, charFCID, False, True, varNom, Count3 + 1, "Accuracy", numBias, False, False, CSng(var10), Count1, strDo, 0, 0, FALSE)
                                '    End If

                                'End If


                                If BOOLSTATSRE Then
                                    If boolSTATSDIFFCOL And BOOLDIFFCOLSTATS Then
                                        'get average of diffcol
                                        'Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", var3, CSng(var10), Count1, strDo, v1, v2, FALSE)
                                        numBias = GetBiasFromDiffCol(idTR, varNom, Count3 + 1, var10, False)
                                    Else
                                        numBias = CalcREPercent(numMean, varNom, intQCDec)
                                        '.Selection.TypeText(Format(numBias, strQCDec))
                                        If intN = 0 Then
                                        Else
                                            Call InsertAnovaTables(intTableID, idTR, charFCID, False, True, varNom, Count3 + 1, "Accuracy", numBias, False, False, CSng(var10), Count1, strDo, 0, 0, False)
                                        End If
                                    End If
                                    If intN = 0 Then
                                        str1 = "NA"
                                    Else
                                        str1 = Format(numBias, strQCDec)
                                    End If
                                End If
                                '****

                                If BOOLSTATSRE And boolSTATSMEAN Then
                                    Try
                                        'enter %RE
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, (Count3 * int11) + 2).Select()

                                        rowsEE = dtblStats.Select(strFnr1)
                                        rowsEE(0).BeginEdit()
                                        rowsEE(0).Item("numAcc") = numBias
                                        rowsEE(0).EndEdit()

                                        'now do all
                                        rowsEE = dtblStats.Select(strFnr2)
                                        var1 = rowsEE(0).Item("numMean")
                                        var2 = varNom
                                        var3 = CalcREPercent(var1, var2, intQCDec)
                                        rowsEE(0).BeginEdit()
                                        rowsEE(0).Item("numAcc") = CDec(var3)
                                        rowsEE(0).EndEdit()

                                        If boolQCREPORTACCVALUES Or boolDoExtraStats = False Then
                                            .Selection.TypeText(Text:=str1)
                                        Else

                                            rowsEE = dtblStats.Select(strFnr2)
                                            var1 = CDec(rowsEE(0).Item("numAcc"))
                                            str2 = Format(var1, strQCDec)
                                            If nE = 0 Then
                                                str2 = "NA"
                                            End If
                                            .Selection.TypeText(Text:=str1 & ChrW(160) & "(" & str2 & ")")

                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If


                                Call InsertAnovaTables(intTableID, idTR, charFCID, False, True, varNom, Count3 + 1, "n", intN, False, False, CSng(var10), Count1, strDo, 0, 0, False)

                                If boolSTATSN Then
                                    Try
                                        'enter n
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, (Count3 * int11) + 2).Select()
                                        '.Selection.TypeText(CStr(int2))
                                        '.Selection.TypeText(CStr(intN))
                                        str1 = CStr(intN)


                                        rowsEE = dtblStats.Select(strFnr1)
                                        rowsEE(0).BeginEdit()
                                        rowsEE(0).Item("intN") = intN
                                        rowsEE(0).EndEdit()

                                        'now do all
                                        rowsEE = dtblStats.Select(strFnr2)
                                        rowsEE(0).BeginEdit()
                                        rowsEE(0).Item("intN") = nE
                                        rowsEE(0).EndEdit()

                                        If boolQCREPORTACCVALUES Or boolDoExtraStats = False Then
                                            .Selection.TypeText(Text:=str1)
                                        Else

                                            str2 = CStr(nE)

                                            .Selection.TypeText(Text:=str1 & ChrW(160) & "(" & str2 & ")")

                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If

                                intCRow = int1 + intRowsX + int8
                                If intCRow > intLastRow Then
                                    intLastRow = intCRow
                                End If

dvNo2:

nextCount3:
                            Next

                            'increase row position counter
                            If Count2 = intNumRuns - 1 Then
                                int1 = int1 + intRowsX + int8 + 1 '4
                            Else
                                int1 = int1 + intRowsX + int8 + 2 '5
                            End If

                        Next

                        .Selection.Tables.Item(1).Cell(int1 - 1, 1).Select()

                        'bottom border this row
                        .Selection.SelectRow()


                        '.Selection.MoveRight Unit:=Microsoft.Office.Interop.Word.wdunits.word.wdunits.wdCharacter, Count:=ctQCs, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                        intRow = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)
                        .Selection.Tables.Item(1).Cell(intRow + 1, 1).Select()

                        '.Selection.ParagraphFormat.PageBreakBefore = True

                        intRow = intRow + 2
                        .Selection.Tables.Item(1).Cell(intRow + 1, 1).Select()

                        ''autofit window
                        '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitContent)
                        '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitContent)
                        '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                        '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                        '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitContent)
                        '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitContent)

                        'go back and merge line 1
                        .Selection.Tables.Item(1).Cell(1, 2).Select()
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        Try
                            .Selection.Cells.Merge()
                        Catch ex As Exception

                        End Try
                        .Selection.Font.Bold = False
                        .Selection.TypeText(Text:="Nominal Concentrations")
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                    Catch ex As Exception

                        str1 = "There was a problem preparing table:"
                        str1 = strM1 & ChrW(10) & ChrW(10) & str1
                        str1 = str1 & ChrW(10) & ChrW(10)
                        str1 = str1 & ex.Message
                        MsgBox(str1, vbInformation, "Problem...")

                    End Try


                    .Selection.Tables.Item(1).Cell(1, 1).Select()
                    'go to end of table
                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdColumn)


                    ''enter table number

                    int1 = 4
                    .Selection.Tables.Item(1).Cell(intRow, 1).Select()
                    If boolINCLANOVA Or boolINCLANOVASUMSTATS Or BOOLINTRARUNSUMSTATS Then
                        With .Selection.ParagraphFormat
                            '.PageBreakBefore = True
                        End With
                    End If
                    'int2 = intRow + 1
                    Dim boolGo As Boolean
                    Dim intRows As Short
                    Dim strAnova As String
                    Dim strSumStats As String

                    strAnova = "ANOVA Section"
                    strSumStats = "Summary Statistics Section"
                    'strSumStats = ""

                    ''wdd.visible = True

                    int2 = intRow - 1 ' + 1
                    If boolINCLANOVA Or boolINCLANOVASUMSTATS Or BOOLINTRARUNSUMSTATS Then
                        intFirstAnova = int2 + 1
                    End If

                    If boolINCLANOVASUMSTATS Then 'this is inter-run stats

                        'If boolINCLANOVA Then
                        int1 = 0
                        'int2 = intRow + 1
                        intRows = intRow + 1 + 8

                        For Count3 = 1 To 9 '16 '14 '8
                            str1 = ""
                            int1 = int1 + 1
                            boolGo = True
                            Select Case Count3
                                Case 1
                                    str1 = strSumStats ' strAnova
                                Case 2
                                    If boolSTATSMEAN Then
                                        'str1 = "Mean Observed Conc."
                                        str1 = "Inter-Run Mean:"
                                    Else
                                        boolGo = False
                                    End If
                                Case 3
                                    If boolSTATSSD Then
                                        str1 = "Inter-Run S.D.:"
                                    Else
                                        boolGo = False
                                    End If
                                Case 4
                                    If boolSTATSCV Then
                                        str1 = "Inter-Run " & ReturnPrecLabel() & ":"
                                    Else
                                        boolGo = False
                                    End If
                                Case 5
                                    If boolSTATSBIAS And boolSTATSMEAN Then
                                        str1 = "Inter-Run %Bias:"
                                    Else
                                        boolGo = False
                                    End If
                                Case 6
                                    If boolTHEORETICAL And boolSTATSMEAN Then
                                        str1 = "Inter-Run %Theoretical:"
                                    Else
                                        boolGo = False
                                    End If
                                Case 7
                                    If boolSTATSDIFF And boolSTATSMEAN Then
                                        str1 = "Inter-Run %Diff:"
                                    Else
                                        boolGo = False
                                    End If
                                Case 8
                                    If BOOLSTATSRE And boolSTATSMEAN Then
                                        str1 = "Inter-Run %RE:"
                                    Else
                                        boolGo = False
                                    End If
                                Case 9
                                    If boolSTATSN Then
                                        str1 = "n:"
                                    Else
                                        boolGo = False
                                    End If

                            End Select
                            If boolGo Then
                                int2 = int2 + 1
                            End If
                            If Len(str1) = 0 Then
                            Else
                                '.Selection.Tables.Item(1).Cell(Count3, 1).Select()
                                .Selection.Tables.Item(1).Cell(int2, 1).Select()
                                .Selection.TypeText(str1)

                            End If
                        Next

                        int2 = int2 + 1

                    End If

                    int1 = 0
                    'int2 = intRow + 1
                    If BOOLINTRARUNSUMSTATS Then

                        intRows = intRow + 1 + 8
                        For Count3 = 1 To 9 '16 '14 '8

                            If boolINCLANOVASUMSTATS And Count3 = 1 Then
                                'don't do strSumStats
                                GoTo nextBOOLINTRARUNSUMSTATS
                            End If
                            str1 = ""
                            int1 = int1 + 1
                            boolGo = True
                            Select Case Count3
                                Case 1
                                    str1 = strSumStats ' strAnova
                                Case 2
                                    If boolSTATSMEAN Then
                                        'str1 = "Mean Observed Conc."
                                        str1 = "Intra-Run Mean:"
                                    Else
                                        boolGo = False
                                    End If
                                Case 3
                                    If boolSTATSSD Then
                                        str1 = "Intra-Run S.D.:"
                                    Else
                                        boolGo = False
                                    End If
                                Case 4
                                    If boolSTATSCV Then
                                        str1 = "Intra-Run " & ReturnPrecLabel() & ":"
                                    Else
                                        boolGo = False
                                    End If
                                Case 5
                                    If boolSTATSBIAS And boolSTATSMEAN Then
                                        str1 = "Intra-Run %Bias:"
                                    Else
                                        boolGo = False
                                    End If
                                Case 6
                                    If boolTHEORETICAL And boolSTATSMEAN Then
                                        str1 = "Intra-Run %Theoretical:"
                                    Else
                                        boolGo = False
                                    End If
                                Case 7
                                    If boolSTATSDIFF And boolSTATSMEAN Then
                                        str1 = "Intra-Run %Diff:"
                                    Else
                                        boolGo = False
                                    End If
                                Case 8
                                    If BOOLSTATSRE And boolSTATSMEAN Then
                                        str1 = "Intra-Run %RE:"
                                    Else
                                        boolGo = False
                                    End If
                                Case 9
                                    If boolSTATSN Then
                                        str1 = "n:"
                                    Else
                                        boolGo = False
                                    End If

                            End Select
                            If boolGo Then
                                int2 = int2 + 1
                            End If
                            If Len(str1) = 0 Then
                            Else
                                '.Selection.Tables.Item(1).Cell(Count3, 1).Select()
                                .Selection.Tables.Item(1).Cell(int2, 1).Select()
                                .Selection.TypeText(str1)

                            End If

nextBOOLINTRARUNSUMSTATS:

                        Next

                        int2 = int2 + 1

                    End If



                    If boolINCLANOVA Then
                        'intRow = int2
                        'int2 = intRow + 1
                        For Count3 = 1 To 4 '16 '14 '8
                            str1 = ""
                            boolGo = True
                            Select Case Count3
                                Case 1
                                    str1 = strAnova ' strSumStats
                                Case 2
                                    str1 = "Between Run Precision (" & ReturnPrecLabel() & "):"
                                Case 3
                                    str1 = "Within Run Precision (" & ReturnPrecLabel() & "):"
                                    'Case 4
                                    '    If boolSTATSDIFF Then
                                    '        str1 = "Within Run %Difference:"
                                    '    Else
                                    '        boolGo = False
                                    '    End If
                                    'Case 5
                                    '    If BOOLSTATSRE Then
                                    '        str1 = "Within Run %RE:"
                                    '    Else
                                    '        boolGo = False
                                    '    End If
                                Case 4 '13
                                    str1 = "Number of Between Run Runs:"
                            End Select
                            If boolGo Then
                                int2 = int2 + 1
                            End If
                            If Len(str1) = 0 Then
                            Else
                                '.Selection.Tables.Item(1).Cell(Count3, 1).Select()
                                .Selection.Tables.Item(1).Cell(int2, 1).Select()
                                .Selection.TypeText(str1)

                            End If
                        Next
                    End If

                    var1 = "a"

                    int12 = -1
                    For Count3 = 0 To intNumLevels - 1

                        Dim boolDoExtraStats As Boolean = False

                        int12 = int12 + 1
                        varNom = tblLevels.Rows.Item(Count3).Item("NOMCONC")
                        strF = strF2 & " AND NOMCONC = " & varNom
                        dv2.RowFilter = ""
                        dv2.RowFilter = strF
                        int2 = dv2.Count


                        'just because gAllowExclSamples And LAllowExclSamples doesn't mean StudyDoc was actually used
                        'check for BOOLUSEGUWUACCCRIT
                        ''this has already been taken into account before
                        'If gAllowExclSamples And LAllowExclSamples Then
                        '    var1 = dv2(0).Item("BOOLUSEGUWUACCCRIT")
                        '    If var1 = 0 Then 'must sort again
                        '        Erase rowsActual
                        '        strF = strF2 & "AND NOMCONC = " & varNom & " AND ELIMINATEDFLAG = 'N'"
                        '        dv2.RowFilter = ""
                        '        dv2.RowFilter = strF
                        '        int2 = dv2.Count
                        '    End If
                        'End If

                        ''create rows1 from tbl1 which will contain data
                        ''set strF to '' because converting right away to datarow and don't need to filter again
                        'strF = ""

                        Erase rows1

                        Dim tbl2SS As System.Data.DataTable = dv2.ToTable
                        'now do rows actual
                        'strFActual = "(" & strF & ") AND (ELIMINATEDFLAG = 'N' OR BOOLEXCLSAMPLE = 0)"
                        If gAllowExclSamples And LAllowExclSamples Then
                            'strFActual = "ELIMINATEDFLAG = 'N' AND BOOLEXCLSAMPLE = 0"
                            strFActual = "(ELIMINATEDFLAG = 'N' OR ELIMINATEDFLAG IS NULL) AND BOOLEXCLSAMPLE = 0"
                        Else
                            'strFActual = "ELIMINATEDFLAG = 'N'"
                            strFActual = "(ELIMINATEDFLAG = 'N' OR ELIMINATEDFLAG IS NULL)"
                        End If
                        strS = "RUNID ASC,RUNSAMPLEORDERNUMBER ASC"
                        rowsActual = tbl2SS.Select(strFActual, strS)
                        int3 = rowsActual.Length
                        intN = int3

                        strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                        rows1 = tbl2SS.Select("RUNID > 0", strS)
                        ''20180716 LEE:
                        ''must account for null concentrations
                        'Try
                        '    rows1 = tbl2SS.Select("RUNID > 0 AND CONCENTRATION IS NOT NULL", strS)
                        'Catch ex As Exception
                        '    var1 = var1
                        'End Try
                        rows2All = rows1
                        int3 = rows1.Length
                        nE = int3

                        If intN <> nE Then
                            boolDoExtraStats = True
                        End If

                        ''debug
                        ''console.writeline("Start")
                        'For Count4 = 0 To int3 - 1
                        '    var1 = "NomConc: " & rows1(Count4).Item("NOMCONC") & ", ELIMINATEDFLAG: " & rows1(Count4).Item("ELIMINATEDFLAG") & ", BOOLEXCLSAMPLE: " & rows1(Count4).Item("BOOLEXCLSAMPLE")
                        '    'console.writeline(var1)
                        'Next
                        ''console.writeline("End")

                        If nE = 0 Then
                            vU = -1
                            v1 = -1
                            v2 = -1
                            Call SetHighAndLowCriteria(varNom, v1, v2, hi, lo)
                        Else
                            vU = rows1(0).Item("BOOLUSEGUWUACCCRIT")
                            v1 = arrFP(1, int12)
                            v2 = arrFP(2, int12)
                            Call SetHighAndLowCriteria(varNom, v1, v2, hi, lo)
                        End If


                        'start ANOVA section
                        dvAn = tblAnova.DefaultView

                        'retrieve anova
                        dvAn.RowFilter = ""
                        dvAn.RowFilter = "NomConc = " & varNom
                        tblAnGo.Clear()
                        tblAnGo = dvAn.ToTable
                        int2 = tblAnGo.Rows.Count
                        ReturnAnova = ANOVA_OneWay(tblAnGo)

                        ''''wdd.visible = True

                        Dim intCC As Short
                        'intRows = intRow + 1 + 17
                        intCC = intFirstAnova ' intRow + 1
                        int2 = 0

                        'boolINCLANOVASUMSTATS
                        'boolINCLANOVA
                        '20160927 LEE: Do these stats, but just don't print if boolINCLANOVASUMSTATS
                        Dim boolAAAB As Boolean = True

                        If boolAAAB Then

                            Dim nr2 As DataRow = dtblStats.NewRow
                            nr2.BeginEdit()
                            nr2.Item("boolAll") = True
                            nr2.Item("RunID") = 0
                            nr2.Item("QCLevel") = Count3 '0-based
                            nr2.EndEdit()
                            dtblStats.Rows.Add(nr2)

                            strFnr2 = "boolAll = TRUE and RunID = 0 AND QCLevel = " & Count3

                            For Count4 = 1 To 9 ' intRow + 1 To intRows
                                int2 = int2 + 1
                                str1 = ""
                                boolGo = True
                                str2 = ""

                                Select Case int2
                                    Case 1 'enter Mean Obs Conc
                                        str1 = "Inter-Run Mean"
                                        var1 = MeanDR(rowsActual, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                        If boolLUseSigFigs Then
                                            numMean = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                        Else
                                            numMean = RoundToDecimalRAFZ(var1, LSigFig)
                                        End If

                                        'now do all
                                        var1 = MeanDR(rows2All, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                        If boolLUseSigFigs Then
                                            var2 = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                        Else
                                            var2 = RoundToDecimalRAFZ(var1, LSigFig)
                                        End If
                                        str2 = CStr(var2)
                                        rowsEE = dtblStats.Select(strFnr2)
                                        rowsEE(0).BeginEdit()
                                        rowsEE(0).Item("numMean") = var2
                                        rowsEE(0).EndEdit()


                                        If boolSTATSMEAN Then

                                            If boolLUseSigFigs Then
                                                str1 = CStr(DisplayNum(numMean, LSigFig, False))
                                            Else
                                                str1 = CStr(Format(numMean, GetRegrDecStr(LSigFig)))
                                            End If

                                            Call InsertAnovaTables(intTableID, idTR, charFCID, True, False, varNom, Count3 + 1, "Mean", numMean, False, False, CSng(var10), Count1, strDo, 0, 0, False)
                                        Else
                                            boolGo = False
                                        End If

                                    Case 2 'Inter-Run S.D.

                                        var1 = StdDevDR(rowsActual, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                        If boolLUseSigFigs Then
                                            numSD = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                        Else
                                            numSD = RoundToDecimalRAFZ(var1, LSigFig)
                                        End If

                                        'now do all
                                        var1 = StdDevDR(rows2All, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                        If boolLUseSigFigs Then
                                            var2 = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                        Else
                                            var2 = RoundToDecimalRAFZ(var1, LSigFig)
                                        End If
                                        rowsEE = dtblStats.Select(strFnr2)
                                        rowsEE(0).BeginEdit()
                                        rowsEE(0).Item("numSD") = var2
                                        rowsEE(0).EndEdit()

                                        If nE < gSDMax Then
                                            str2 = "NA"
                                        Else
                                            str2 = CStr(var2)
                                        End If

                                        If boolSTATSSD Then

                                            If intN < gSDMax Then
                                                str1 = "NA"
                                            Else
                                                If boolLUseSigFigs Then
                                                    str1 = CStr(DisplayNum(numSD, LSigFig, False))
                                                Else
                                                    str1 = CStr(Format(numSD, GetRegrDecStr(LSigFig)))
                                                End If
                                            End If

                                            Call InsertAnovaTables(intTableID, idTR, charFCID, True, False, varNom, Count3 + 1, "SD", numSD, False, False, CSng(var10), Count1, strDo, 0, 0, False)
                                        Else
                                            Call InsertAnovaTables(intTableID, idTR, charFCID, True, False, varNom, Count3 + 1, "SD", numSD, True, False, CSng(var10), Count1, strDo, 0, 0, False)
                                            boolGo = False
                                        End If

                                        var1 = var1

                                    Case 3 'Inter-Run %CV
                                        numPrec = CalcCVPercent(numSD, numMean, intQCDec)
                                        If boolSTATSCV Then

                                            If intN < gSDMax Then
                                                str1 = "NA"
                                            Else
                                                str1 = Format(numPrec, strQCDec) 'CStr(var1)
                                            End If

                                            Call InsertAnovaTables(intTableID, idTR, charFCID, True, False, varNom, Count3 + 1, "Precision", numPrec, False, False, CSng(var10), Count1, strDo, 0, 0, False)

                                            'now do all
                                            rowsEE = dtblStats.Select(strFnr2)
                                            var1 = rowsEE(0).Item("numMean")
                                            var2 = rowsEE(0).Item("numSD")
                                            var3 = CalcCVPercent(var2, var1, intQCDec)
                                            rowsEE(0).BeginEdit()
                                            rowsEE(0).Item("numPrec") = CDec(var3)
                                            rowsEE(0).EndEdit()

                                            If nE < gSDMax Then
                                                str2 = "NA"
                                            Else
                                                str2 = Format(CDec(var3), strQCDec)
                                            End If

                                        Else
                                            boolGo = False
                                            Call InsertAnovaTables(intTableID, idTR, charFCID, True, False, varNom, Count3 + 1, "Precision", numPrec, True, False, CSng(var10), Count1, strDo, 0, 0, False)
                                        End If

                                    Case 4 'Inter-Run Bias
                                        'var1 = (((numMean / varNom) - 1) * 100)
                                        'numBias = CDec(Format(var1, strQCDec))

                                        If boolSTATSDIFFCOL Then
                                            numBias = GetBiasFromDiffCol(idTR, varNom, Count3 + 1, 0, False)
                                        Else
                                            numBias = CalcREPercent(numMean, varNom, intQCDec)
                                        End If

                                        If boolSTATSBIAS And boolSTATSMEAN Then


                                            str1 = Format(numBias, strQCDec)
                                            Call InsertAnovaTables(intTableID, idTR, charFCID, True, False, varNom, Count3 + 1, "Accuracy", numBias, False, False, CSng(var10), Count1, strDo, 0, 0, False)

                                            'now do all


                                            If boolSTATSDIFFCOL Then

                                                var3 = GetAveDiffColAcc(Count3 + 1, dtblAccDiffCol, 0, True)

                                            Else

                                                'legend
                                                'strFnr1 = "boolAll = FALSE and RunID = " & var10 & " AND QCLevel = " & Count3
                                                'strFnr2 = "boolAll = TRUE and RunID = " & var10 & " AND QCLevel = " & Count3

                                                rowsEE = dtblStats.Select(strFnr2)
                                                var1 = rowsEE(0).Item("numMean")
                                                var2 = varNom
                                                var3 = CalcREPercent(var1, var2, intQCDec)

                                            End If

                                            rowsEE(0).BeginEdit()
                                            rowsEE(0).Item("numAcc") = CDec(var3)
                                            rowsEE(0).EndEdit()

                                            str2 = Format(CDec(var3), strQCDec)
                                            If nE = 0 Then
                                                str2 = "NA"
                                            End If

                                            If intN = 0 Then
                                                str1 = "NA"
                                            End If

                                        Else
                                            boolGo = False
                                        End If

                                    Case 5 'Inter-Run Theor

                                        'var1 = CDec(Format((((numMean / varNom) - 1) * 100), strQCDec))
                                        'numTheor = 100 + var1

                                        If boolSTATSDIFFCOL Then
                                            numTheor = GetBiasFromDiffCol(idTR, varNom, Count3 + 1, 0, False)
                                        Else
                                            numTheor = CalcREPercent(numMean, varNom, intQCDec)
                                        End If

                                        numTheor = 100 + CDec(numTheor)

                                        If boolTHEORETICAL And boolSTATSMEAN Then
                                            str1 = Format(numTheor, strQCDec)
                                            Call InsertAnovaTables(intTableID, idTR, charFCID, True, False, varNom, Count3 + 1, "Accuracy", numTheor, False, False, CSng(var10), Count1, strDo, 0, 0, False)

                                            'now do all
                                            If boolSTATSDIFFCOL Then
                                                var3 = GetAveDiffColAcc(Count3 + 1, dtblAccDiffCol, 0, True)
                                            Else

                                                rowsEE = dtblStats.Select(strFnr2)
                                                var1 = rowsEE(0).Item("numMean")
                                                var2 = varNom
                                                var3 = CalcREPercent(var1, var2, intQCDec)

                                            End If

                                            var3 = 100 + CDec(var3)

                                            rowsEE(0).BeginEdit()
                                            rowsEE(0).Item("numAcc") = CDec(var3)
                                            rowsEE(0).EndEdit()

                                            str2 = Format(CDec(var3), strQCDec)
                                            If nE = 0 Then
                                                str2 = "NA"
                                            End If

                                            If intN = 0 Then
                                                str1 = "NA"
                                            End If

                                        Else
                                            boolGo = False
                                        End If

                                    Case 6 'Inter-Run %Diff
                                        'var1 = (((numMean / varNom) - 1) * 100)
                                        'numBias = CDec(Format(var1, strQCDec))

                                        'numBias = CalcREPercent(numMean, varNom, intQCDec)
                                        If boolSTATSDIFFCOL Then
                                            numBias = GetBiasFromDiffCol(idTR, varNom, Count3 + 1, 0, False)
                                        Else
                                            numBias = CalcREPercent(numMean, varNom, intQCDec)
                                        End If

                                        If boolSTATSDIFF And boolSTATSMEAN Then
                                            str1 = Format(numBias, strQCDec)
                                            Call InsertAnovaTables(intTableID, idTR, charFCID, True, False, varNom, Count3 + 1, "Accuracy", numBias, False, False, CSng(var10), Count1, strDo, 0, 0, False)

                                            'now do all
                                            If boolSTATSDIFFCOL Then

                                                var3 = GetAveDiffColAcc(Count3 + 1, dtblAccDiffCol, 0, True)

                                            Else
                                                rowsEE = dtblStats.Select(strFnr2)
                                                var1 = rowsEE(0).Item("numMean")
                                                var2 = varNom
                                                var3 = CalcREPercent(var1, var2, intQCDec)

                                            End If

                                            rowsEE(0).BeginEdit()
                                            rowsEE(0).Item("numAcc") = CDec(var3)
                                            rowsEE(0).EndEdit()

                                            str2 = Format(CDec(var3), strQCDec)
                                            If nE = 0 Then
                                                str2 = "NA"
                                            End If

                                            If intN = 0 Then
                                                str1 = "NA"
                                            End If

                                        Else
                                            boolGo = False
                                        End If

                                    Case 7 'Inter-Run %RE
                                        'var1 = (((numMean / varNom) - 1) * 100)
                                        'numBias = CDec(Format(var1, strQCDec))

                                        'numBias = CalcREPercent(numMean, varNom, intQCDec)
                                        If boolSTATSDIFFCOL Then
                                            numBias = GetBiasFromDiffCol(idTR, varNom, Count3 + 1, 0, False)
                                        Else
                                            numBias = CalcREPercent(numMean, varNom, intQCDec)
                                        End If

                                        If BOOLSTATSRE And boolSTATSMEAN Then
                                            str1 = Format(numBias, strQCDec)
                                            Call InsertAnovaTables(intTableID, idTR, charFCID, True, False, varNom, Count3 + 1, "Accuracy", numBias, False, False, CSng(var10), Count1, strDo, 0, 0, False)

                                            'now do all
                                            If boolSTATSDIFFCOL Then

                                                var3 = GetAveDiffColAcc(Count3 + 1, dtblAccDiffCol, 0, True)

                                            Else

                                                rowsEE = dtblStats.Select(strFnr2)
                                                var1 = rowsEE(0).Item("numMean")
                                                var2 = varNom
                                                var3 = CalcREPercent(var1, var2, intQCDec)

                                            End If

                                            rowsEE(0).BeginEdit()
                                            rowsEE(0).Item("numAcc") = CDec(var3)
                                            rowsEE(0).EndEdit()

                                            str2 = Format(CDec(var3), strQCDec)
                                            If nE = 0 Then
                                                str2 = "NA"
                                            End If

                                            If intN = 0 Then
                                                str1 = "NA"
                                            End If

                                        Else
                                            boolGo = False
                                        End If



                                    Case 8 'n
                                        If boolSTATSN Then
                                            str1 = CStr(intN)
                                            str2 = CStr(nE)
                                            Call InsertAnovaTables(intTableID, idTR, charFCID, True, False, varNom, Count3 + 1, "n", intN, False, False, CSng(var10), Count1, strDo, 0, 0, False)
                                        Else
                                            boolGo = False
                                        End If

                                End Select

                                If boolINCLANOVASUMSTATS Then

                                    If boolGo Then
                                        intCC = intCC + 1
                                    End If
                                    If Len(str1) = 0 Then
                                    Else
                                        '.Selection.Tables.Item(1).Cell(Count4, Count3 + 2).Select()
                                        .Selection.Tables.Item(1).Cell(intCC, (Count3 * int11) + 2).Select()
                                        If int2 = 1 Then
                                            'v1 =
                                            'determine if value is outside acceptance criteria
                                            'HERE
                                            'If numMean > hi Or numMean < lo Then 'flag
                                            If OutsideAccCrit(numMean, varNom, v1, v2, NZ(vU, 0)) Then
                                                intLeg = intLeg + 1
                                                strA = ChrW(intLeg + intLegStart)

                                                'Set Legend String
                                                str1 = GetLegendStringIncluded(v1, v2, vU)
                                                boolHasOutlier = HasOutlier(str1, boolHasOutlier)
                                                'Add to Legend Array
                                                ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                                If boolRedBoldFont Then
                                                    .Selection.Font.Bold = True
                                                    .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                End If

                                                If boolLUseSigFigs Then
                                                    .Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                                Else
                                                    .Selection.TypeText(Text:=CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                                End If


                                                Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                                If boolQCREPORTACCVALUES = False And boolDoExtraStats Then
                                                    .Selection.TypeText(ChrW(160) & "(" & str2 & ")")
                                                End If
                                                'boolEnterDiff = True
                                            Else
                                                If boolQCREPORTACCVALUES = False And boolDoExtraStats Then
                                                    .Selection.TypeText(str1 & ChrW(160) & "(" & str2 & ")")
                                                Else
                                                    .Selection.TypeText(str1)
                                                End If
                                            End If

                                        Else

                                            If boolQCREPORTACCVALUES = False And boolDoExtraStats Then
                                                .Selection.TypeText(str1 & ChrW(160) & "(" & str2 & ")")
                                            Else
                                                .Selection.TypeText(str1)
                                            End If

                                        End If

                                    End If

                                End If


                            Next

                            If boolINCLANOVASUMSTATS Then
                                intCC = intCC + 1
                            End If

                        End If


                        '****

                        int2 = 0
                        '20160927 LEE: Do these stats, but just don't print if BOOLINTRARUNSUMSTATS
                        Dim boolAAAA As Boolean = True
                        If boolAAAA Then

                            'get n
                            Dim nTIS As Short
                            strF = "NomConc = " & varNom & " AND boolAll = FALSE"
                            Erase rowsTIS
                            rowsTIS = tblIntraSum.Select(strF)
                            nTIS = rowsTIS.Length

                            Dim nTISX As Short
                            Dim rowsTISX() As DataRow
                            strF = "NomConc = " & varNom & " AND boolAll = TRUE"
                            rowsTISX = tblIntraSum.Select(strF)
                            nTISX = rowsTISX.Length

                            Dim numMean1 As Decimal
                            Dim numSD1 As Decimal
                            Dim numPrec1 As Decimal
                            Dim numBias1 As Decimal
                            Dim numTheor1 As Decimal

                            boolDoExtraStats = False

                            Dim CountRun As Short

                            For Count4 = 1 To 9 ' intRow + 1 To intRows
                                int2 = int2 + 1
                                str1 = ""
                                str2 = ""
                                boolGo = True


                                Select Case int2
                                    Case 1 'enter Mean Obs Conc
                                        str1 = "Intra-Run Mean"
                                        ' Function StdDevDR(ByVal r() As DataRow, ByVal strCol As String, ByVal boolAliq As Boolean, ByVal strAliq As String, ByVal boolSF As Boolean, ByVal boolUseIS As Boolean)
                                        var1 = MeanDR(rowsTIS, "Conc", False, "", True, False)
                                        If boolLUseSigFigs Then
                                            numMean = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                        Else
                                            numMean = RoundToDecimalRAFZ(var1, LSigFig)
                                        End If

                                        var1 = MeanDR(rowsTISX, "Conc", False, "", True, False)
                                        If boolLUseSigFigs Then
                                            numMean1 = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                        Else
                                            numMean1 = RoundToDecimalRAFZ(var1, LSigFig)
                                        End If

                                        If numMean = numMean1 Or nTISX = 0 Then
                                            boolDoExtraStats = False
                                        Else
                                            boolDoExtraStats = True
                                        End If

                                        If boolSTATSMEAN Then
                                            If boolLUseSigFigs Then
                                                str1 = CStr(DisplayNum(numMean, LSigFig, False))
                                            Else
                                                str1 = CStr(Format(numMean, GetRegrDecStr(LSigFig)))
                                            End If

                                            If boolLUseSigFigs Then
                                                str2 = CStr(DisplayNum(numMean1, LSigFig, False))
                                            Else
                                                str2 = CStr(Format(numMean1, GetRegrDecStr(LSigFig)))
                                            End If

                                        Else
                                            boolGo = False
                                        End If

                                    Case 2 'Intra-Run S.D.

                                        var1 = StdDevDR(rowsTIS, "Conc", False, "", True, False)
                                        If boolLUseSigFigs Then
                                            numSD = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                        Else
                                            numSD = RoundToDecimalRAFZ(var1, LSigFig)
                                        End If

                                        var1 = StdDevDR(rowsTISX, "Conc", False, "", True, False)
                                        If boolLUseSigFigs Then
                                            numSD1 = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                        Else
                                            numSD1 = RoundToDecimalRAFZ(var1, LSigFig)
                                        End If

                                        If boolSTATSSD Then
                                            If nTIS < gSDMax Then
                                                str1 = "NA"
                                            Else
                                                If boolLUseSigFigs Then
                                                    str1 = CStr(DisplayNum(numSD, LSigFig, False))
                                                Else
                                                    str1 = CStr(Format(numSD, GetRegrDecStr(LSigFig)))
                                                End If
                                            End If

                                            If nTISX < gSDMax Then
                                                str2 = "NA"
                                            Else
                                                If boolLUseSigFigs Then
                                                    str2 = CStr(DisplayNum(numSD1, LSigFig, False))
                                                Else
                                                    str2 = CStr(Format(numSD1, GetRegrDecStr(LSigFig)))
                                                End If
                                            End If


                                        Else
                                            boolGo = False
                                        End If

                                    Case 3 'Intra-Run %CV
                                        numPrec = CalcCVPercent(numSD, numMean, intQCDec)
                                        numPrec1 = CalcCVPercent(numSD1, numMean1, intQCDec)
                                        If boolSTATSCV Then
                                            str1 = Format(numPrec, strQCDec) 'CStr(var1)
                                            str2 = Format(numPrec1, strQCDec) 'CStr(var1)
                                            Call InsertAnovaTables(intTableID, idTR, charFCID, True, False, varNom, Count3 + 1, "IntraRunPrecision", numPrec, False, False, CSng(var10), Count1, strDo, 0, 0, False)

                                        Else
                                            boolGo = False
                                        End If

                                        If nTIS < gSDMax Then
                                            str1 = "NA"
                                        End If
                                        If nTISX < gSDMax Then
                                            str2 = "NA"
                                        End If

                                    Case 4 'Intra-Run Bias
                                        'var1 = (((numMean / varNom) - 1) * 100)
                                        'numBias = CDec(Format(var1, strQCDec))

                                        If boolSTATSDIFFCOL Then
                                            'this needs to be average of diffcol for each analytical run
                                            Dim v10
                                            Dim num10 As Decimal
                                            Dim numTot As Decimal = 0
                                            Dim intTot As Short = 0
                                            For CountRun = 0 To intNumRuns - 1
                                                v10 = tblNumRuns.Rows.Item(CountRun).Item("RUNID")
                                                num10 = RoundToDecimal(RoundToDecimalRAFZ(GetBiasFromDiffCol(idTR, varNom, Count3 + 1, CInt(v10), False), intQCDec + 4), intQCDec)
                                                If num10 = -1 Then
                                                Else
                                                    numTot = numTot + num10
                                                    intTot = intTot + 1
                                                End If
                                            Next
                                            If intTot = 0 Then
                                                numBias = -1
                                            Else
                                                numBias = RoundToDecimal(RoundToDecimalRAFZ(numTot / intTot, intQCDec + 4), intQCDec)
                                            End If

                                            'do all
                                            numTot = 0
                                            intTot = 0
                                            For CountRun = 0 To intNumRuns - 1
                                                v10 = tblNumRuns.Rows.Item(CountRun).Item("RUNID")
                                                num10 = RoundToDecimal(RoundToDecimalRAFZ(GetAveDiffColAcc(Count3 + 1, dtblAccDiffCol, CInt(v10), True), intQCDec + 4), intQCDec)
                                                If num10 = -1 Then
                                                Else
                                                    numTot = numTot + num10
                                                    intTot = intTot + 1
                                                End If
                                            Next
                                            If intTot = 0 Then
                                                numBias1 = -1
                                            Else
                                                numBias1 = RoundToDecimal(RoundToDecimalRAFZ(numTot / intTot, intQCDec + 4), intQCDec)
                                            End If

                                        Else
                                            numBias = CalcREPercent(numMean, varNom, intQCDec)
                                            numBias1 = CalcREPercent(numMean1, varNom, intQCDec)
                                        End If


                                        If boolSTATSBIAS And boolSTATSMEAN Then
                                            If numBias = -1 Then
                                                str1 = "NA"
                                            Else
                                                str1 = Format(numBias, strQCDec)
                                                Call InsertAnovaTables(intTableID, idTR, charFCID, True, False, varNom, Count3 + 1, "IntraRunAccuracy", numBias, False, False, CSng(var10), Count1, strDo, 0, 0, False)
                                            End If
                                            If numBias = -1 Then
                                                str2 = "NA"
                                            Else
                                                str2 = Format(numBias1, strQCDec)
                                            End If
                                            If nTISX = 0 Then
                                                str2 = "NA"
                                            End If

                                            If nTIS = 0 Then
                                                str1 = "NA"
                                            End If

                                        Else
                                            boolGo = False
                                        End If


                                    Case 5 'Intra-Run Theor

                                        'var1 = CDec(Format((((numMean / varNom) - 1) * 100), strQCDec))
                                        'numTheor = 100 + var1
                                        If boolSTATSDIFFCOL Then

                                            'this needs to be average of diffcol for each analytical run
                                            Dim v10
                                            Dim num10 As Decimal
                                            Dim numTot As Decimal = 0
                                            Dim intTot As Short = 0
                                            For CountRun = 0 To intNumRuns - 1
                                                v10 = tblNumRuns.Rows.Item(CountRun).Item("RUNID")
                                                num10 = RoundToDecimal(RoundToDecimalRAFZ(GetBiasFromDiffCol(idTR, varNom, Count3 + 1, CInt(v10), False), intQCDec + 4), intQCDec)
                                                If num10 = -1 Then
                                                Else
                                                    numTot = numTot + num10
                                                    intTot = intTot + 1
                                                End If
                                            Next
                                            If intTot = 0 Then
                                                numTheor = -1
                                            Else
                                                numTheor = RoundToDecimal(RoundToDecimalRAFZ(numTot / intTot, intQCDec + 4), intQCDec)
                                            End If
                                            numTheor = 100 + CDec(numTheor)

                                            'do all
                                            numTot = 0
                                            intTot = 0
                                            For CountRun = 0 To intNumRuns - 1
                                                v10 = tblNumRuns.Rows.Item(CountRun).Item("RUNID")
                                                num10 = RoundToDecimal(RoundToDecimalRAFZ(GetAveDiffColAcc(Count3 + 1, dtblAccDiffCol, CInt(v10), True), intQCDec + 4), intQCDec)
                                                If num10 = -1 Then
                                                Else
                                                    numTot = numTot + num10
                                                    intTot = intTot + 1
                                                End If
                                            Next
                                            If intTot = 0 Then
                                                numTheor1 = -1
                                            Else
                                                numTheor1 = RoundToDecimal(RoundToDecimalRAFZ(numTot / intTot, intQCDec + 4), intQCDec)
                                            End If
                                            numTheor1 = 100 + CDec(numTheor1)

                                        Else
                                            numTheor = CalcREPercent(numMean, varNom, intQCDec)
                                            numTheor = 100 + CDec(numTheor)

                                            numTheor1 = CalcREPercent(numMean1, varNom, intQCDec)
                                            numTheor1 = 100 + CDec(numTheor)
                                        End If

                                        If boolTHEORETICAL And boolSTATSMEAN Then
                                            If numTheor = -1 Then
                                                str1 = "NA"
                                            Else
                                                str1 = Format(numTheor, strQCDec)
                                                Call InsertAnovaTables(intTableID, idTR, charFCID, True, False, varNom, Count3 + 1, "IntraRunAccuracy", numTheor, False, False, CSng(var10), Count1, strDo, 0, 0, False)
                                            End If
                                            If numTheor1 = -1 Then
                                                str2 = "NA"
                                            Else
                                                str2 = Format(numTheor1, strQCDec)
                                            End If
                                            If nTISX = 0 Then
                                                str2 = "NA"
                                            End If

                                            If nTIS = 0 Then
                                                str1 = "NA"
                                            End If

                                        Else
                                            boolGo = False
                                        End If


                                    Case 6 'Intra-Run %Diff
                                        'var1 = (((numMean / varNom) - 1) * 100)
                                        'numBias = CDec(Format(var1, strQCDec))

                                        If boolSTATSDIFFCOL Then
                                            'this needs to be average of diffcol for each analytical run
                                            Dim v10
                                            Dim num10 As Decimal
                                            Dim numTot As Decimal = 0
                                            Dim intTot As Short = 0
                                            For CountRun = 0 To intNumRuns - 1
                                                v10 = tblNumRuns.Rows.Item(CountRun).Item("RUNID")
                                                num10 = RoundToDecimal(RoundToDecimalRAFZ(GetBiasFromDiffCol(idTR, varNom, Count3 + 1, CInt(v10), False), intQCDec + 4), intQCDec)
                                                If num10 = -1 Then
                                                Else
                                                    numTot = numTot + num10
                                                    intTot = intTot + 1
                                                End If
                                            Next
                                            If intTot = 0 Then
                                                numBias = -1
                                            Else
                                                numBias = RoundToDecimal(RoundToDecimalRAFZ(numTot / intTot, intQCDec + 4), intQCDec)
                                            End If

                                            'do all
                                            numTot = 0
                                            intTot = 0
                                            For CountRun = 0 To intNumRuns - 1
                                                v10 = tblNumRuns.Rows.Item(CountRun).Item("RUNID")
                                                num10 = RoundToDecimal(RoundToDecimalRAFZ(GetAveDiffColAcc(Count3 + 1, dtblAccDiffCol, CInt(v10), True), intQCDec + 4), intQCDec)
                                                If num10 = -1 Then
                                                Else
                                                    numTot = numTot + num10
                                                    intTot = intTot + 1
                                                End If
                                            Next
                                            If intTot = 0 Then
                                                numBias1 = -1
                                            Else
                                                numBias1 = RoundToDecimal(RoundToDecimalRAFZ(numTot / intTot, intQCDec + 4), intQCDec)
                                            End If

                                        Else
                                            numBias = CalcREPercent(numMean, varNom, intQCDec)
                                            numBias1 = CalcREPercent(numMean1, varNom, intQCDec)
                                        End If

                                        If boolSTATSDIFF And boolSTATSMEAN Then

                                            If numBias = -1 Then
                                                str1 = "NA"
                                            Else
                                                str1 = Format(numBias, strQCDec)
                                                Call InsertAnovaTables(intTableID, idTR, charFCID, True, False, varNom, Count3 + 1, "IntraRunAccuracy", numBias, False, False, CSng(var10), Count1, strDo, 0, 0, False)
                                            End If
                                            If numBias = -1 Then
                                                str2 = "NA"
                                            Else
                                                str2 = Format(numBias1, strQCDec)
                                            End If
                                            If nTISX = 0 Then
                                                str2 = "NA"
                                            End If

                                            If nTIS = 0 Then
                                                str1 = "NA"
                                            End If

                                        Else
                                            boolGo = False
                                        End If


                                    Case 7 'Intra-Run %RE
                                        'var1 = (((numMean / varNom) - 1) * 100)
                                        'numBias = CDec(Format(var1, strQCDec))

                                        If boolSTATSDIFFCOL Then
                                            'this needs to be average of diffcol for each analytical run
                                            Dim v10
                                            Dim num10 As Decimal
                                            Dim numTot As Decimal = 0
                                            Dim intTot As Short = 0
                                            For CountRun = 0 To intNumRuns - 1
                                                v10 = tblNumRuns.Rows.Item(CountRun).Item("RUNID")
                                                num10 = RoundToDecimal(RoundToDecimalRAFZ(GetBiasFromDiffCol(idTR, varNom, Count3 + 1, CInt(v10), False), intQCDec + 4), intQCDec)
                                                If num10 = -1 Then
                                                Else
                                                    numTot = numTot + num10
                                                    intTot = intTot + 1
                                                End If
                                            Next
                                            If intTot = 0 Then
                                                numBias = -1
                                            Else
                                                numBias = RoundToDecimal(RoundToDecimalRAFZ(numTot / intTot, intQCDec + 4), intQCDec)
                                            End If


                                            'do all
                                            numTot = 0
                                            intTot = 0
                                            For CountRun = 0 To intNumRuns - 1
                                                v10 = tblNumRuns.Rows.Item(CountRun).Item("RUNID")
                                                num10 = RoundToDecimal(RoundToDecimalRAFZ(GetAveDiffColAcc(Count3 + 1, dtblAccDiffCol, CInt(v10), True), intQCDec + 4), intQCDec)
                                                If num10 = -1 Then
                                                Else
                                                    numTot = numTot + num10
                                                    intTot = intTot + 1
                                                End If
                                            Next
                                            If intTot = 0 Then
                                                numBias1 = -1
                                            Else
                                                numBias1 = RoundToDecimal(RoundToDecimalRAFZ(numTot / intTot, intQCDec + 4), intQCDec)
                                            End If

                                        Else
                                            numBias = CalcREPercent(numMean, varNom, intQCDec)
                                            numBias1 = CalcREPercent(numMean1, varNom, intQCDec)
                                        End If

                                        If BOOLSTATSRE And boolSTATSMEAN Then
                                            If numBias = -1 Then
                                                str1 = "NA"
                                            Else
                                                str1 = Format(numBias, strQCDec)
                                                Call InsertAnovaTables(intTableID, idTR, charFCID, True, False, varNom, Count3 + 1, "IntraRunAccuracy", numBias, False, False, CSng(var10), Count1, strDo, 0, 0, False)
                                            End If
                                            If numBias = -1 Then
                                                str2 = "NA"
                                            Else
                                                str2 = Format(numBias1, strQCDec)
                                            End If
                                            If nTISX = 0 Then
                                                str2 = "NA"
                                            End If

                                            If nTIS = 0 Then
                                                str1 = "NA"
                                            End If

                                        Else
                                            boolGo = False
                                        End If


                                    Case 8 'n
                                        If boolSTATSN Then
                                            str1 = CStr(nTIS)
                                            str2 = CStr(nTISX)
                                            Call InsertAnovaTables(intTableID, idTR, charFCID, False, True, varNom, Count3 + 1, "IntraRunN", nTIS, False, False, CSng(var10), Count1, strDo, 0, 0, False)
                                        Else
                                            boolGo = False
                                        End If

                                End Select


                                If BOOLINTRARUNSUMSTATS Then

                                    If boolINCLANOVASUMSTATS And Count4 = 1 Then
                                        'move intcc back one
                                        intCC = intCC - 1
                                    End If

                                    If boolGo Then
                                        intCC = intCC + 1
                                    End If

                                    If Len(str1) = 0 Then
                                    Else
                                        '.Selection.Tables.Item(1).Cell(Count4, Count3 + 2).Select()
                                        .Selection.Tables.Item(1).Cell(intCC, (Count3 * int11) + 2).Select()
                                        If int2 = 1 Then

                                            'determine if value is outside acceptance criteria
                                            'HERE
                                            'If numMean > hi Or numMean < lo Then 'flag
                                            If OutsideAccCrit(numMean, varNom, v1, v2, NZ(vU, 0)) Then 'flag

                                                intLeg = intLeg + 1
                                                strA = ChrW(intLeg + intLegStart)

                                                'Set Legend String
                                                str1 = GetLegendStringIncluded(v1, v2, vU)
                                                boolHasOutlier = HasOutlier(str1, boolHasOutlier)
                                                'Add to Legend Array
                                                ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                                If boolRedBoldFont Then
                                                    .Selection.Font.Bold = True
                                                    .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                End If

                                                .Selection.TypeText(Text:=CStr(numMean))

                                                Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                                If boolQCREPORTACCVALUES = False And boolDoExtraStats Then
                                                    .Selection.TypeText(ChrW(160) & "(" & str2 & ")")
                                                End If
                                                'boolEnterDiff = True
                                            Else
                                                If boolQCREPORTACCVALUES = False And boolDoExtraStats Then
                                                    .Selection.TypeText(str1 & ChrW(160) & "(" & str2 & ")")
                                                Else
                                                    .Selection.TypeText(str1)
                                                End If

                                                'boolEnterDiff = True
                                            End If

                                        Else
                                            If boolQCREPORTACCVALUES = False And boolDoExtraStats Then
                                                .Selection.TypeText(str1 & ChrW(160) & "(" & str2 & ")")
                                            Else
                                                .Selection.TypeText(str1)
                                            End If
                                        End If

                                    End If

                                End If

                            Next

                            If BOOLINTRARUNSUMSTATS Then
                                intCC = intCC + 1
                            End If

                        End If

                        '****
                        int2 = 0

                        ''wdd.visible = True
                        'boolINCLANOVASUMSTATS
                        If boolINCLANOVA Then
                            For Count4 = 1 To 3 'intRow + 1 To intRows
                                int2 = int2 + 1
                                str1 = ""
                                boolGo = True

                                Select Case int2

                                    Case 1 'Between Run Precision (%CV)
                                        var1 = ReturnAnova(0)
                                        If IsNumeric(var1) Then
                                            str1 = Format(var1, strQCDec)
                                        Else
                                            str1 = var1
                                        End If

                                    Case 2 'Within Run Precision (%CV)
                                        var1 = ReturnAnova(1)
                                        If IsNumeric(var1) Then
                                            str1 = Format(var1, strQCDec)
                                        Else
                                            str1 = var1
                                        End If
                                    Case 3 'Number of runs
                                        'find number of runs by doing a distinct
                                        tblZ = dv2.ToTable("a", True, "RUNID")
                                        int1 = tblZ.Rows.Count
                                        str1 = CStr(int1)
                                End Select

                                If boolGo Then
                                    intCC = intCC + 1
                                End If

                                If Len(str1) = 0 Then
                                Else
                                    .Selection.Tables.Item(1).Cell(intCC, (Count3 * int11) + 2).Select()
                                    .Selection.TypeText(str1)
                                End If
                            Next

                        End If

                    Next

                    Dim intRowE As Short
                    intRow = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdEndOfRangeRowNumber) + 1
                    intRowE = .Selection.Tables.Item(1).Rows.Count - 2

                    ''wdd.visible = True

                    'autofit table
                    Call AutoFitTable(wd, False)

                    ''''''wdd.visible = True

                    'delete extra rows
                    For Count2 = intRowE To intRow Step -1
                        .Selection.Tables.Item(1).Cell(Count2, 1).Select()
                        .Selection.Rows.Delete()
                    Next

                    'remove unused rows
                    Call RemoveRows(wd, 1)

                    ''wdd.visible = True

                    'check to see if the addition of legends is going to require the insertion of a page break
                    If boolINCLANOVA Or boolINCLANOVASUMSTATS Or BOOLINTRARUNSUMSTATS Then

                        ''wdd.visible = True

                        'autofit table
                        Call AutoFitTable(wd, False)

                        ''wdd.visible = True

                        Dim pT As Int64
                        p2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                        pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                        int1 = .Selection.Tables.Item(1).Rows.Count
                        int2 = int1 + intLeg + 2
                        'insert intLeg + 2 rows below selection
                        .Selection.InsertRowsBelow(intLeg + 2)
                        .Selection.Tables.Item(1).Cell(int1, 1).Select()

                        For Count2 = int1 To int2
                            .Selection.Tables.Item(1).Cell(Count2, 1).Select()
                            pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                            If pT <> p2 Then
                                boolFirstAnova = True
                                Exit For
                            End If
                        Next

                        'wd.Visible = True

                        If boolFirstAnova Then

                            ''wdd.visible = True

                            '*****20180714 LEE:
                            'DO NOT FORCE PAGE BREAK!
                            'Messes up SplitTable

                            '.Selection.Tables.Item(1).Cell(intFirstAnova, 1).Select()
                            'With .Selection.ParagraphFormat
                            '    .PageBreakBefore = True
                            'End With

                            ''PageBreakBefore for some reason  top-borders the selection
                            ''must remove underline
                            '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                            '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                            '.Selection.Tables.Item(1).Cell(intFirstAnova, 1).Select()

                            '******

                            'remove unused rows
                            Call RemoveRows(wd, 1)
                        Else
                            Call RemoveRows(wd, 1)
                        End If

                        ''wdd.visible = True

                        'autofit table
                        Call AutoFitTable(wd, False)

                        'If boolHasOutlier Then
                        '    'go to end of table and add a row
                        '    .Selection.Tables.Item(1).Rows(.Selection.Tables.Item(1).Rows.Count).Select()
                        '    .Selection.InsertRowsBelow(1)
                        '    .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        '    .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                        '    .Selection.Cells.Merge()
                        '    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                        '    .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
                        '    '20180419 LEE: Added 'statistical'
                        '    .Selection.TypeText("The statistical results within parentheses were calculated including the outlier value.")
                        'End If

                    End If

                    'enter table number
                    str1 = "Summary of " & rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION") & " Interpolated QC Standard Concentrations with Between Run and Within Run Summary Statistics"

                    '***
                    strA = rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")
                    If gNumMatrix = 1 Then
                        strA = strAnalC
                    Else
                        strA = strAnal 'strAnalC has '..Matrix', don't want to pass that here
                    End If
                    'No, just strAnal
                    strA = strAnal
                    strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, False)
                    Call EnterTableNumber(wd, strTName, 5, strA, strTempInfo, intTableID, intGroup, idTR)
                    '***

                    'Call EnterTableNumber(wd, str1, 4)

                    'enter a table record in tblTableN
                    'ctTableN = ctTableN + 1
                    Dim dtblr As DataRow = tblTableN.NewRow
                    dtblr.BeginEdit()
                    dtblr.Item("TableNumber") = ctTableN
                    dtblr.Item("AnalyteName") = strDo 'arrAnalytes(1, Count1)
                    dtblr.Item("TableName") = strTNameO
                    dtblr.Item("TableID") = intTableID
                    dtblr.Item("CHARFCID") = charFCID
                    dtblr.Item("TableNameNew") = strTName
                    tblTableN.Rows.Add(dtblr)

                    .Selection.Tables.Item(1).Cell(intRow + 1, 1).Select()

                    'split table, if needed
                    str1 = frmH.lblProgress.Text

                    ctLegend = ctLegend + 1
                    intLeg = intLeg + 1
                    arrLegend(1, intLeg) = "NA"
                    arrLegend(2, intLeg) = "Not Applicable"
                    arrLegend(3, intLeg) = False

                    '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                    '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)


                    'strM = "Creating Summary of Interpolated QC Standard Concentrations Table For " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & "..."

                    'autofit table
                    Call AutoFitTable(wd, BOOLINCLUDEDATE)

                    strM = "Finalizing " & strTName & "..."
                    strM1 = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    str1 = strM1

                    frmH.lblProgress.Text = strM1
                    frmH.Refresh()

                    Call SplitTable(wd, 4, intLeg, arrLegend, str1, False, intLeg + 2, False, True, boolFirstAnova, intTableID)

                    Call RemoveRows(wd, 1)

                    'Call NoSplitTable(wd, 4, intLeg, arrLegend, str1, False, True, 5, ctLegend + 5, False, True)
                    'move to line below table

                    'autofit table
                    Call AutoFitTable(wd, False)

                    Call MoveOneCellDown(wd)

                    Call InsertLegend(wd, intTableID, idTR, False, 1)

                End If

end1:

                If boolJustTable Then

                    If gNumMatrix = 1 Then
                        strA = strAnalC
                    Else
                        strA = strAnal 'strAnalC has '..Matrix', don't want to pass that here
                    End If
                    'No, just strAnal
                    strA = strAnal
                    str1 = strA ' NZ(rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION"), "")
                    'Call JustTable(wd, str1, str2, strDo, strTName, intTableID)
                    If Len(str1) = 0 Then
                    Else
                        strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, False)
                        Call JustTable(wd, str1, strTName, strDo, strTName, intTableID, strTempInfo, "", strTNameO, intGroup, idTR)
                    End If

                End If

                'If boolJustTable Then
                '    'str1 = "Summary of " & rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION") & " Interpolated QC Standard Concentrations with Between Run and Within Run Summary Statistics"
                '    str2 = str1
                '    'str1 = rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")
                '    ''Call JustTable(wd, str1, str2, strDo, strTName, intTableID)
                '    'Call JustTable(wd, str1, strTName, strDo, strTName, intTableID, strTempInfo, "")

                '    str1 = NZ(rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION"), "")
                '    'Call JustTable(wd, str1, str2, strDo, strTName, intTableID)
                '    If Len(str1) = 0 Then
                '    Else
                '        Call JustTable(wd, str1, strTName, strDo, strTName, intTableID, strTempInfo, "")
                '    End If

                'End If

next1:

            Next


end2:
            'clear formatting again
            .Selection.Find.ClearFormatting()


            If .ActiveWindow.View.SplitSpecial = Word.WdSpecialPane.wdPaneNone Then
                .ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdPrintView
            Else
                .ActiveWindow.View.Type = Word.WdViewType.wdPrintView
            End If

        End With


    End Sub


    Sub SRSummaryOfAnalRuns_1(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal idTR As Int64)

        Dim BACStudy As String
        Dim rs As New ADODB.Recordset
        Dim constr As String
        Dim dbPath As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim Count4 As Short
        Dim Count5 As Short
        Dim Count6 As Short
        Dim Count7 As Short
        Dim var1, var2, var3, var4, var5, var6, var7, var8, var9
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        'Dim frmh.arrregcon()
        Dim arrTemp(2, 50)
        Dim num1 As Object
        Dim num2 As Object
        Dim num3 As Object
        Dim arrBCStdActual()

        Dim ctLegend As Short
        Dim lng1 As Integer
        Dim lng2 As Integer
        Dim boolPortrait As Boolean
        Dim intLastAnal As Short
        Dim arrOrder()
        Dim ctCols 'number of columns in a table
        Dim strSub1 As String
        Dim strSub2 As String
        Dim pos1 As Short
        Dim pos2 As Short
        Dim tblARS As System.Data.DataTable
        Dim ct1 As Short

        Dim ctStart As Short
        Dim ctEnd As Short
        Dim Count10, Count9 As Short
        Dim dt As Date
        Dim dvDo As System.Data.DataView
        Dim bool As Boolean
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim intDo As Short
        Dim strDo As String
        Dim ctanalyticalruns As Short
        Dim strTName As String
        Dim tblS As System.Data.DataTable 'datatable for tblConfigHeader that has report table column names
        Dim strF As String
        Dim rowsS() As DataRow
        Dim intCols As Short
        Dim strS As String
        Dim tblHL As System.Data.DataTable 'datatable for tblConfigHeaderLookup that has dv column names
        Dim rowsHL() As DataRow
        Dim intColsHL As Short
        Dim strTempInfo As String
        Dim strA As String

        Dim strM As String
        Dim strM1 As String

        'tblS and tblHL have in common:
        '  ID_TBLCONFIGREPORTTABLES and ID_TBLCONFIGHEADERLOOKUP

        Dim strMatrix As String
        Dim intGroup As Short

        Dim fontsize

        Dim strTNameO As String
        Dim idTCHL As Int16

        Dim charFCID As String
        strF = "ID_TBLREPORTTABLE = " & idTR
        Dim rowsTR() As DataRow = tblReportTable.Select(strF)
        var1 = rowsTR(0).Item("CHARFCID")
        charFCID = NZ(var1, "NA")

        fontsize = wd.ActiveDocument.Styles("Normal").Font.Size 'wd.Selection.Font.Size

        ''''''wdd.visible = True

        Dim intTableID As Short
        intTableID = 1

        Dim strWRunId As String = GetWatsonColH(intTableID)

        'retrieve report table column header information
        tblS = tblReportTableHeaderConfig
        strF = "id_tblStudies = " & id_tblStudies & " AND boolInclude = -1 AND ID_TBLCONFIGREPORTTABLES = " & intTableID
        strS = "intOrder"
        rowsS = tblS.Select(strF, strS)
        'determine number of rows
        intCols = rowsS.Length
        Dim arrAnalyticalRuns(intCols, 500)

        'retrieve default config table column header information
        tblHL = tblConfigHeaderLookup
        strF = "id_tblStudies = '" & id_tblStudies & "AND ID_TBLCONFIGREPORTTABLES = 1"

        'find number of Analytical Runs from tables
        ''

        dvDo = frmH.dgvReportTableConfiguration.DataSource
        intDo = FindRowDVNumByCol(intTableID, dvDo, "id_tblconfigreporttables")

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
        strTName = NZ(var1, "[NONE]")
        strTNameO = strTName

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

        Dim dvARS1 As System.Data.DataView
        'tblARS = tblAnalyticalRunSummary
        int1 = tblAnalyticalRunSummary.Rows.Count 'debug
        Dim dgvARS As DataGridView = frmH.dgvAnalyticalRunSummary
        dvARS1 = frmH.dgvAnalyticalRunSummary.DataSource
        ct1 = dvARS1.Count 'ct = Number of Analytical Runs in the AnalyticalRunSummary Table.

        '20160908 LEE: Going forward, need to filter dvARS, which directly affects dgvAnalyticalRunSummary
        'so instead, convert to new table, then back to dvARS so code doesn't change much
        tblARS = dvARS1.ToTable
        Dim dvARS As System.Data.DataView = New DataView(tblARS)


        'First, Record unique AnalyteID's in table (this should be global)
        Dim dtAnalytes, dtUniqueAnalyteIDs As New DataTable
        Dim dvAnalytes As New DataView(dtAnalytes)
        Dim ctUniqueAnalyteIDs As Short
        Dim strAnalyteID As String
        Dim boolPSAE As Boolean = False

        dtAnalytes.Columns.Add("AnalyteID")
        For x As Integer = 1 To (arrAnalytes.GetUpperBound(1))
            If (StrComp(arrAnalytes(2, x), "", CompareMethod.Text) = 0) Then
            Else
                Dim row As DataRow = dtAnalytes.NewRow
                row.Item("AnalyteID") = arrAnalytes(2, x)
                dtAnalytes.Rows.Add(row)

            End If
        Next

        dtUniqueAnalyteIDs = dvAnalytes.ToTable("dtUniqueAnalyteIDsTable", True, "AnalyteID")
        ctUniqueAnalyteIDs = dtUniqueAnalyteIDs.Rows.Count

        strM = "Creating " & strTName & "..."
        strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
        frmH.lblProgress.Text = strM
        frmH.Refresh()

        'Write a Table for each Unique AnalyteID
        For Count9 = 0 To ctUniqueAnalyteIDs - 1

            strTName = strTNameO

            Dim arrLegend(4, 20)

            Dim boolAnalytesExistForTable As Boolean = False
            Dim thisAnalyteID As Int64 = dtUniqueAnalyteIDs.Rows(Count9).Item("AnalyteID")
            Dim thisAnalyteName As String
            Dim thisAnalyteColumnName As String
            Dim firstAnalyte As Short 'Number of Analyte in arrAnalytes
            bool = False

            'check if table is to be generated
            'Cycle through Analytes and see if any with AnalyteID are checked
            For Count10 = 1 To ctAnalytes
                var1 = arrAnalytes(2, Count10)
                var2 = CLng(NZ(dtUniqueAnalyteIDs.Rows(Count9).Item("AnalyteID"), 0))
                If (StrComp(var1, var2, CompareMethod.Text) = 0) Then 'If the Analyte has the right AnalyteID...
                    strDo = arrAnalytes(1, Count10) 'record Analyte name
                    If (UseAnalyte(CStr(strDo))) Then 'If Analyte is in Checked in the Table Configuration...
                        boolAnalytesExistForTable = True
                        If (Not (bool)) Then 'If we haven't already found an Analyte with the correct AnalyteID that is checked
                            bool = dvDo.Item(intDo).Item(strDo)  'Then check this Analyte; maybe *it* is checked.
                        End If
                        '20160206 LEE: Don't need to keep looping
                        If bool Then
                            Exit For
                        End If
                    End If
                End If
            Next

            If Not (boolAnalytesExistForTable) Then
                GoTo NEXT1
            End If

            Dim boolHasDate As Boolean = False

            If bool Then 'continue    (otherwise, if no Analytes checked, go to next AnalyteID)

                intTCur = intTCur + 1

                strM = "Creating " & strTName & " ..."
                strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                strM1 = strM
                frmH.lblProgress.Text = strM
                frmH.Refresh()

                'page setup according to configuration
                str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                'insert page break
                Call InsertPageBreak(wd)
                Call PageSetup(wd, str1) 'L=Landscape, P=Portrait

                Count1 = 0
                firstAnalyte = 0

                '20160908 LEE: This table should reproduce exactly what is from the Analytical Run Summary window, filtered for analyteid
                'Note that ANALYTEID is text in dvARS
                strF = "ANALYTEID = '" & thisAnalyteID & "'"
                'strFTot = "(" & strFTot & ") OR RUNANALYTEREGRESSIONSTATUS = -1" to get blank rows
                dvARS.RowFilter = strF

                ''debug
                'Console.WriteLine(strF & ": " & Count9)
                'var1 = ""
                'For Count2 = 0 To tblARS.Columns.Count - 1
                '    var1 = var1 & ";" & tblARS.Columns(Count2).ColumnName
                'Next
                'Console.WriteLine(var1)
                'For Count3 = 0 To dvARS.Count - 1
                '    var1 = ""
                '    For Count2 = 0 To tblARS.Columns.Count - 1
                '        var1 = var1 & ";" & dvARS(Count3).Item(Count2)
                '    Next
                '    Console.WriteLine(var1)
                'Next


                'Go through all Analytical Runs, write out the ones which are relevant to this AnalyteID
                For Count2 = 0 To dvARS.Count - 1

                    'Check if it's the right AnalyteID, the Analyte is present in the run, 
                    'and the run is selected in Analytical Run Summary, and the Compound is selected in the Table Configuration

                    '20171219 LEE: boolInThisRunsAssayID only defines if run has a calibration run
                    'shouldn't be used here to evaluate

                    var1 = dvARS(Count2).Item("AnalyteID")
                    var2 = dvARS(Count2).Item("boolInThisRunsAssayID")
                    var3 = dvARS(Count2).Item("boolInclude")

                    'Be careful here about the blank row
                    thisAnalyteName = dvARS(Count2).Item("Analyte")
                    thisAnalyteColumnName = dvARS(Count2).Item("Analyte_C")

                    '20160908 LEE: var4 analysis excludes rows that should be included
                    'shouldn't have to make this evaluation since it's been handled at the Analytical Run Summary level and we have filtered out empty thisAnalyteColumnName
                    'If String.IsNullOrEmpty(thisAnalyteColumnName) Then
                    '    var4 = False
                    'Else
                    '    var4 = dvDo.Item(intDo).Item(thisAnalyteColumnName)
                    'End If
                    var4 = True

                    '20171219 LEE: This logic isn't working for complex multi-analyte multi-matrix multi-calcurve studies
                    'e.g. Alturas POP01, Ricerca/Concord 034116

                    '20171219 LEE: boolInThisRunsAssayID only defines if run has a calibration run
                    'shouldn't be used here to evaluate

                    '20180717 LEE:
                    'Do not need to make this evaluation anymore
                    'Should show whatever is in dvARS
                    'If ((StrComp(var1, CStr(thisAnalyteID), CompareMethod.Text) = 0) And (StrComp(var2, "Yes", CompareMethod.Text) = 0) And var3 And var4) Then
                    ''If ((StrComp(var1, CStr(thisAnalyteID), CompareMethod.Text) = 0) And var3 And var4) Then
                    '20180717 LEE:
                    'Do not need to make this evaluation anymore
                    'Should show whatever is in dvARS
                    'If (firstAnalyte = 0) Then
                    '    'Find First Analyte that is in this Analyte Index
                    '    For Count10 = 1 To ctAnalytes
                    '        var1 = arrAnalytes(1, Count10)
                    '        var2 = dvARS(Count2).Item("Analyte_C")
                    '        If (StrComp(var1, var2, CompareMethod.Text) = 0) Then
                    '            firstAnalyte = Count10
                    '            Exit For
                    '        End If
                    '    Next
                    'End If

                    '20181129 LEE:
                    'Aack! var3 (boolInclude) evaluation was removed when the previous logic was removed
                    'put back in
                    var3 = CInt(NZ(dvARS(Count2).Item("boolInclude"), 0))
                    If var3 = 0 Then
                        GoTo nextCount2
                    End If

                    'Go ahead and put it into the array
                    Count1 = Count1 + 1
                    If Count2 > UBound(arrAnalyticalRuns, 2) Then
                        ReDim Preserve arrAnalyticalRuns(intCols, UBound(arrAnalyticalRuns, 2) + 500)
                    End If
                    For Count3 = 1 To intCols
                        int1 = rowsS(Count3 - 1).Item("ID_TBLCONFIGHEADERLOOKUP")
                        strF = "ID_TBLCONFIGHEADERLOOKUP = " & int1 & " AND ID_TBLCONFIGREPORTTABLES = 1"
                        Erase rowsHL
                        rowsHL = tblHL.Select(strF)
                        str2 = NZ(rowsHL(0).Item("CHARCOLUMNLABEL"), "NA")

                        If StrComp(str2, "Comments", CompareMethod.Text) = 0 Then
                            If frmH.rbUseWatsonComments.Checked Then
                                arrAnalyticalRuns(Count3, Count1) = dvARS.Item(Count2).Item("Watson Comments") 'Watson comments
                            Else
                                arrAnalyticalRuns(Count3, Count1) = dvARS.Item(Count2).Item("User Comments") 'User Comments
                            End If
                        ElseIf StrComp(str2, "Pass/Fail", CompareMethod.Text) = 0 Then
                            var1 = dvARS.Item(Count2).Item(str2)
                            arrAnalyticalRuns(Count3, Count1) = var1
                            ''20160908 LEE: Don't know why this is being done
                            'If StrComp(var1, "Accepted", CompareMethod.Text) = 0 Then
                            '    arrAnalyticalRuns(Count3, Count1) = "Passed"
                            'ElseIf StrComp(var1, "Rejected", CompareMethod.Text) = 0 Then
                            '    arrAnalyticalRuns(Count3, Count1) = "Failed"
                            'End If
                        Else
                            var1 = dvARS.Item(Count2).Item(str2) 'debug
                            arrAnalyticalRuns(Count3, Count1) = dvARS.Item(Count2).Item(str2) 'rs.Fields("RUNID").Value
                        End If
                    Next

                    'Else
                    '    var1 = var1
                    'End If

nextCount2:

                Next Count2
                'var1 = Sheets("AnalRuns").Range("AnalRunsHome").Offset(Count3, 0).Value

                ctanalyticalruns = Count1

                'enter Summary of Analytical Runs
                '20180717 LEE:
                'Do not need to make this evaluation anymore
                'Should show whatever is in dvARS
                firstAnalyte = 1
                If (firstAnalyte <> 0) Then
                    'Note that we are referring to the table by its first StudyDoc analyte, despite
                    'the fact that it represents multiple sub-Analytes (different matrices, etc)
                    'strM = "Creating " & strTName & ": " & arrAnalytes(1, firstAnalyte) & "..."
                    strM = "Creating " & strTName & ": " & thisAnalyteName & "..."
                    strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    frmH.lblProgress.Text = strM
                    frmH.Refresh()
                    'wd.Selection.TypeParagraph()

                    wrdSelection = wd.Selection()
                    Dim intRows As Short
                    Dim intRow As Short

                    'int1 = (ctanalyticalruns * 2) - 1
                    int1 = (ctanalyticalruns * 2)

                    If boolPlaceHolder Then
                        intRows = 1
                        intCols = 1
                    Else
                        If boolExcludeEntireTableTitle Then
                            intRows = int1 + 1
                        Else
                            intRows = int1 + 2
                        End If

                    End If


                    With wd
                        '.Selection.Goto What:=wdGoToBookmark, Name:="tblSummaryofAnalyticalRuns"


                        Try

                            '20180913 LEE:
                            Call IncrNextTableNumber(wd)

                            .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=intRows, NumColumns:=intCols, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)

                            .Selection.Tables.Item(1).Rows.AllowBreakAcrossPages = False

                            wrdSelection = wd.Selection()

                            .Selection.Tables.Item(1).Select()

                            Call SetCellPaddingZero(.Selection.Tables.Item(1))

                            With .Selection 'remove initial borders
                                '.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                                .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                                '.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                                .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                                .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                                .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                                .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                                .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                            End With

                            If boolPlaceHolder Then

                                .Selection.Tables.Item(1).Select()
                                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone

                                strA = arrAnalytes(14, Count10)
                                'strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup)
                                Call EnterTableNumber(wd, strTName, 3, strA, strTempInfo, intTableID, intGroup, idTR)
                                Call MoveOneCellDown(wd)

                                .Selection.TypeParagraph()
                                .Selection.TypeParagraph()

                                'enter a table record in tblTableN
                                'ctTableN = ctTableN + 1
                                Dim dtblr1 As DataRow = tblTableN.NewRow
                                dtblr1.BeginEdit()
                                dtblr1.Item("TableNumber") = ctTableN
                                'dtblr1.Item("AnalyteName") = arrAnalytes(1, Count1)
                                dtblr1.Item("AnalyteName") = thisAnalyteName ' arrAnalytes(1, firstAnalyte)
                                dtblr1.Item("TableName") = strTNameO
                                dtblr1.Item("TableID") = intTableID
                                dtblr1.Item("CHARFCID") = charFCID
                                dtblr1.Item("TableNameNew") = strTName
                                tblTableN.Rows.Add(dtblr1)

                                GoTo next1
                            End If

                            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalTop '.wdCellAlignVerticalBottom

                            .Selection.Tables.Item(1).Select()
                            Call GlobalTableParaFormat(wd)

                            '20171220 LEE:
                            'Do not set table size, use the style default table
                            '.Selection.Font.Size = fontsize - 1
                            .Selection.Tables.Item(1).Cell(1, 1).Select()


                            '.Selection.HomeKey(Unit:=wdLine)
                            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine)

                            ''20160908 LEE: Don't do this anymore
                            'For Count1 = 1 To intCols

                            '    Select Case rowsS(Count1 - 1).Item("ID_TBLCONFIGHEADERLOOKUP")
                            '        Case Is = 87 'Analysis Date
                            '            num1 = 80
                            '        Case Is = 89 'Comments
                            '            num1 = 200
                            '        Case Is = 90 'Extraction Date
                            '            num1 = 80
                            '        Case Is = 91 'Notebook ID
                            '            num1 = 115
                            '        Case Is = 92 'Samples
                            '            num1 = 100
                            '        Case Is = 93 'Watson Run ID
                            '            num1 = 50
                            '        Case Is = 159 'Pass/Fail
                            '            num1 = 65
                            '        Case Is = 215 'run type
                            '            num1 = 80
                            '        Case Is = 216 'matrix
                            '            num1 = 60
                            '        Case Is = 217 'lloq
                            '            num1 = 50
                            '        Case Is = 218 'uloq
                            '            num1 = 50
                            '    End Select

                            '    .Selection.Tables.Item(1).Columns.Item(Count1).PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPoints
                            '    .Selection.Tables.Item(1).Columns.Item(Count1).PreferredWidth = num1
                            '    ''
                            'Next

                            'enter table heading
                            str1 = "Summary of Analytical Runs: " & thisAnalyteName ' arrAnalytes(14, firstAnalyte)

                            '20160908 LEE:
                            'check to see if any columns need wordwrap = false
                            'do this before adding the table number row
                            For Count2 = 1 To intCols
                                var1 = arrAnalyticalRuns(Count2, Count1)
                                var2 = rowsS(Count2 - 1).Item("CHARUSERLABEL")
                                idTCHL = rowsS(Count2 - 1).Item("ID_TBLCONFIGHEADERLOOKUP")
                                If idTCHL = 87 Or idTCHL = 90 Then '20160908 LEE: these are date columns. Don't let them wrap
                                    boolHasDate = True
                                    'select column
                                    .Selection.Tables.Item(1).Cell(1, Count2).Select()
                                    .Selection.SelectColumn()
                                    Call DoCells(.Selection.Cells)
                                End If
                            Next Count2



                            '***
                            'We also use the firstAnalyte when entering the Table Number
                            strA = thisAnalyteName ' NZ(arrAnalytes(14, firstAnalyte), "NA")
                            'strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup)
                            Call EnterTableNumber(wd, strTName, 3, strA, strTempInfo, intTableID, 1, idTR)
                            '***
                            'enter a table record in tblTableN
                            'ctTableN = ctTableN + 1
                            Dim dtblr As DataRow = tblTableN.NewRow
                            dtblr.BeginEdit()
                            dtblr.Item("TableNumber") = ctTableN
                            dtblr.Item("AnalyteName") = thisAnalyteName ' arrAnalytes(1, firstAnalyte)
                            dtblr.Item("TableName") = strTNameO
                            dtblr.Item("TableID") = intTableID
                            dtblr.Item("CHARFCID") = charFCID
                            dtblr.Item("TableNameNew") = strTName
                            tblTableN.Rows.Add(dtblr)

                            ''ensure the table is selected
                            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToTable, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToFirst, Count:=ctTableN, Name:="")

                            If boolExcludeEntireTableTitle Then
                                intRow = 1
                            Else
                                intRow = 2
                            End If
                            .Selection.Tables.Item(1).Cell(intRow, 1).Select()


                            'enter headings
                            For Count1 = 1 To intCols
                                str1 = NZ(rowsS(Count1 - 1).Item("CHARUSERLABEL"), "NA")
                                .Selection.Tables.Item(1).Cell(intRow, Count1).Range.Text = str1
                            Next

                            'fit columns
                            .Selection.Tables.Item(1).Cell(intRow, 1).Select()
                            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                            '.Selection.Font.Size = 10
                            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom 'align bottom the header row
                            .Selection.Font.Bold = False
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                            .Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

                            intRow = intRow + 1

                            .Selection.Tables.Item(1).Cell(intRow, 1).Select()

                            '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)


                            Count5 = 0
                            Dim ctP As Short
                            'For Count1 = 1 To ctanalyticalruns * ctAnalytes Step ctAnalytes
                            ctP = 0
                            boolPSAE = False
                            Dim intARow As Short
                            intARow = -1
                            For Count1 = 1 To ctanalyticalruns

                                intARow = intARow + 2

                                If Count1 > ctP Then
                                    'strM = "Entering Analytical Run # " & arrAnalyticalRuns(1, Count1) & " of " & ctanalyticalruns & " for " & arrAnalytes(1, Count10) & "..."
                                    strM = "Entering " & Count1 & " of " & ctanalyticalruns & " for " & arrAnalytes(1, Count10) & "..."
                                    strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                                    frmH.lblProgress.Text = strM
                                    frmH.Refresh()
                                    ctP = ctP + 5
                                End If
                                Count3 = 0
                                Count4 = 0
                                For Count2 = 1 To intCols
                                    var1 = NZ(arrAnalyticalRuns(Count2, Count1), "")
                                    var2 = rowsS(Count2 - 1).Item("CHARUSERLABEL")
                                    'idTCHL = rowsS(Count2 - 1).Item("ID_TBLCONFIGHEADERLOOKUP")
                                    'If idTCHL = 87 Or idTCHL = 90 Then '20160908 LEE: these are date columns. Don't let them wrap
                                    '    Call DoCells(.Selection.Tables.Item(1).Cell(intRow + intARow, Count2))
                                    '    'With .Selection.Tables.Item(1).Cell(intRow + intARow, Count2)
                                    '    '    .WordWrap = False
                                    '    '    .FitText = True
                                    '    'End With
                                    'End If
                                    If InStr(1, var1, "PSAE", CompareMethod.Text) > 0 Then
                                        boolPSAE = True
                                    End If
                                    If InStr(1, var2, "Date", CompareMethod.Text) > 0 Then
                                        If Len(var2) = 0 Then
                                            '.Selection.TypeText(CStr(NZ(var1, "NA")))
                                            .Selection.Tables.Item(1).Cell(intRow + intARow, Count2).Range.Text = CStr(NZ(var1, "NA"))
                                        Else
                                            If IsDate(var1) Then
                                                '.Selection.TypeText(Format(CDate(var1), LDateFormat))
                                                .Selection.Tables.Item(1).Cell(intRow + intARow, Count2).Range.Text = Format(CDate(var1), LDateFormat)
                                            Else
                                                '.Selection.TypeText(CStr(NZ(var1, "NA")))
                                                .Selection.Tables.Item(1).Cell(intRow + intARow, Count2).Range.Text = CStr(NZ(var1, "NA"))
                                            End If
                                        End If
                                    Else
                                        '.Selection.TypeText(CStr(NZ(var1, "NA")))
                                        .Selection.Tables.Item(1).Cell(intRow + intARow, Count2).Range.Text = CStr(NZ(var1, "NA"))
                                    End If
                                    If Count2 = intCols Then
                                        Count5 = Count5 + 1
                                    Else

                                    End If
                                Next

                            Next

                        Catch ex As Exception

                            str1 = "There was a problem preparing table:"
                            str1 = strM1 & ChrW(10) & ChrW(10) & str1
                            str1 = str1 & ChrW(10) & ChrW(10)
                            str1 = str1 & ex.Message
                            MsgBox(str1, vbInformation, "Problem...")

                        End Try



                        'now split table if needed
                        arrLegend(1, 1) = "NR"
                        '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitContent)

                        arrLegend(2, 1) = "Not Reported"
                        arrLegend(3, 1) = False

                        arrLegend(1, 2) = "NA"
                        arrLegend(2, 2) = "Not Applicable"
                        arrLegend(3, 2) = False
                        ctLegend = 2

                        If boolPSAE Then
                            arrLegend(1, 3) = "PSAE"
                            arrLegend(2, 3) = "Pre-Study Assay Evaluation"
                            arrLegend(3, 3) = False
                            ctLegend = 3

                        End If

                        str1 = frmH.lblProgress.Text

                        'autofit table
                        Call AutoFitTable(wd, boolHasDate)

                        Pause(0.1)

                        strM = "Finalizing " & strTName & "..."
                        strM1 = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                        str1 = strM1

                        frmH.lblProgress.Text = strM1
                        frmH.Refresh()

                        Call SplitTable(wd, 2, ctLegend, arrLegend, str1, False, 2, False, False, False, intTableID)
                        'autofit table
                        Call AutoFitTable(wd, boolHasDate)

                        Call MoveOneCellDown(wd)
                        Call InsertLegend(wd, intTableID, idTR, False, 1)


                    End With
                End If
            End If

NEXT1:

        Next

end1:


    End Sub


End Module
