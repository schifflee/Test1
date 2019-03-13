Option Compare Text

Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.ComponentModel.PropertyDescriptorCollection
Imports Word = Microsoft.Office.Interop.Word
Imports Microsoft.VisualBasic
Imports System.IO

Module Module1
    'Public strSponsor As String
    'Public strCompany As String
    'Public strInSupportOf As String
    'Public ctAnalytes As Short
    Public arrR2(20)
    Public arrBC(20)
    Public arrQC(20)
    Public arrSSC(20)
    Public boolGender
    Friend frmH As frmHome_01
    Friend frmC As frmConsole
    Friend frmSD As frmSDHome
    Friend frmAbort As frmAbort
    'Friend frmUpdate As frmUpdateCheck
    Public boolSplitTable As Boolean = False
    Public ctRealLegend As Short = 0

    Sub AssignedQCsTable_4(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal idTR As Int64)

        Dim numNomConc As Decimal
        Dim var1, var2, var3, var4, var5, var6, var7, var10
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
        Dim varNom, varConc
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
        '1=Max, 2=Min
        Dim strFP As String
        Dim numMean As Decimal
        Dim numBias As Decimal
        Dim numSD As Decimal
        Dim numPrec As Decimal
        Dim boolPro As Boolean

        Dim tblZ As New System.Data.DataTable
        Dim dvAn As System.Data.DataView
        Dim tblAnGo As New System.Data.DataTable
        Dim p1, p2, p3, p4, p5, p6, p7, p8, p9, p10
        Dim strM As String
        Dim numDF As Decimal
        Dim rowsX() As DataRow
        Dim intLegStart As Short
        Dim ctQCs As Short
        Dim intD As Short

        Dim DilQCFactor()

        Erase DilQCFactor
        ReDim DilQCFactor(ctQCs)

        Dim varAliquot
        Dim boolJustTable As Boolean
        Dim strTempInfo As String

        Dim intExp As Short
        Dim ctExp As Short
        Dim int8 As Short

        Dim numTheor As Single

        Dim v1, v2, vU

        Dim intRC As Short

        Dim boolOC As Boolean = False 'bool if eliminated

        Dim intGroup As Short
        Dim strAnal As String
        Dim strAnalC As String
        Dim strMatrix As String
        Dim intRunID As Int16
        Dim strDECISIONREASON As String
        Dim boolExFromAS As Boolean

        Dim charFCID As String
        strF = "ID_TBLREPORTTABLE = " & idTR
        Dim rowsTR() As DataRow = tblReportTable.Select(strF)
        var1 = rowsTR(0).Item("CHARFCID")
        charFCID = NZ(var1, "NA")

        boolJustTable = False

        Cursor.Current = Cursors.WaitCursor

        '''wdd.visible = True

        Dim fonts
        fontsize = wd.ActiveDocument.Styles("Normal").Font.Size ' wd.Selection.Font.Size
        fonts = fontsize 'wd.Selection.Font.Size

        Dim strTNameO As String
        With wd

            'dvDo = frmH.dgvReportTableConfiguration.DataSource
            'strTName = "Summary of Interpolated QC Std Conc"
            'intDo = FindRowDVByCol(strTName, dvDo, "Table")

            intTableID = 4

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

            ''legend
            'tbl1 = tblAnalysisResultsHome
            'tbl2 = tblAssignedSamples
            'tbl3 = tblAssignedSamplesHelper
            'tbl4 = tblAnalytesHome

            'ensure data has been entered
            strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idTR
            rowsX = tbl2.Select(strF)
            'If rowsX.Length = 0 Then
            '    strM = "Creating Summary of Interpolated QC Standard Concentrations Table ...."
            '    frmH.lblProgress.Text = strM
            '    MsgBox("Samples have not been assigned to this table.", MsgBoxStyle.Information, "Samples have not been assigned...")
            '    GoTo end2
            'End If

            strF = "IsIntStd = 'No'"
            'strS = ReturnSort(False)
            strS = "INTORDER ASC, IsIntStd ASC, AnalyteDescription ASC"
            rows11 = tblAnalytesHome.Select(strF, strS)
            intRowsAnal = rows11.Length

            ''build tblAnova
            'tblZ.Columns.Add("Group", Type.GetType("System.Int16"))
            'tblZ.Columns.Add("Conc", Type.GetType("System.Decimal"))
            'tblZ.Columns.Add("NomConc", Type.GetType("System.Decimal"))

            'build tblz to record stats info
            tblZ.Columns.Add("NomConc", Type.GetType("System.Decimal"))
            tblZ.Columns.Add("Conc", Type.GetType("System.Decimal"))
            tblZ.Columns.Add("ALIQUOTFACTOR", Type.GetType("System.Decimal"))
            tblZ.Columns.Add("ELIMINATEDFLAG", Type.GetType("System.String"))
            tblZ.Columns.Add("BOOLOUTLIER", Type.GetType("System.Boolean"))
            tblZ.Columns.Add("HI", Type.GetType("System.Decimal"))
            tblZ.Columns.Add("LO", Type.GetType("System.Decimal"))
            tblZ.Columns.Add("v1", Type.GetType("System.Decimal"))
            tblZ.Columns.Add("v2", Type.GetType("System.Decimal"))
            tblZ.Columns.Add("CHARHELPER1", Type.GetType("System.String"))
            tblZ.Columns.Add("ASSAYLEVEL", Type.GetType("System.Decimal"))

            Dim arrDup(100, 3)
            '1=strDo, 2=Replicate
            Dim intDup As Short
            Dim intDupCt As Short
            Dim boolDupGo As Boolean

            Dim vAnalyteIndex
            Dim vMasterAssayID
            Dim vAnalyteID
            Dim tblAG As DataTable = tblAnalyteGroups 'tblAnalyteGroups has all analytes, not just accepted

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

                'clear tblz
                tblZ.Rows.Clear()

                'check if table is to be generated
                'strDo = arrAnalytes(1, Count1) 'record column name
                var1 = rows11(Count1 - 1).Item("ANALYTEDESCRIPTION")

                If UseAnalyte(CStr(var1)) Then
                Else
                    GoTo end1
                End If

                strDo = arrAnalytes(1, Count1)
                var2 = var1
                bool = dvDo.Item(intDo).Item(strDo) 'find boolean value of dvDo column

                Dim strM1 As String
                If bool Then 'continue

                    intTCur = intTCur + 1

                    'ensure data has been entered
                    strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & CleanText(strDo) & "' AND ID_TBLREPORTTABLE = " & idTR

                    rowsX = tbl2.Select(strF)

                    'need to have strmatrix before doing booljusttable

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
                        'page setup according to configuration
                        str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                        'insert page break
                        'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                        Call InsertPageBreak(wd)
                        Call PageSetup(wd, str1) 'L=Landscape, P=Portrait
                        boolJustTable = True
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

                    ''legend:
                    'tbl1 = tblAnalysisResultsHome
                    'tbl2 = tblAssignedSamples
                    'tbl3 = tblAssignedSamplesHelper
                    'tbl4 = tblAnalytesHome

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

                    If BOOLINCLUDEDATE Then
                        strS = "ASSAYDATETIME ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                    Else
                        strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                    End If
                    rows2 = tbl2.Select(strF2, strS)
                    int1 = rows2.Length 'debug
                    dv2 = New DataView(tbl2, strF2, strS, DataViewRowState.CurrentRows)
                    int2 = dv2.Count 'debug

                    'find number of runs used
                    tblNumRuns = dv2.ToTable("a", True, "RUNID")
                    intNumRuns = tblNumRuns.Rows.Count

                    'establish table of level numbers
                    'must be sorted by nomconc!
                    'make new dv
                    'Dim dvNL As New DataView(tbl2, strF2, "NOMCONC ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC", DataViewRowState.CurrentRows)
                    '20171124 LEE:
                    Dim dvNL As New DataView(tbl2, strF2, "NOMCONC ASC, ALIQUOTFACTOR DESC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC", DataViewRowState.CurrentRows)
                    'tblLevels = dvNL.ToTable("b", True, "ALIQUOTFACTOR", "NOMCONC", "CHARHELPER1", "ASSAYLEVEL")
                    'NOTE: Must add ASSAYLEVEL after the fact because sometimes study may have different assay levels for the same nomconc
                    'tblLevels = dvNL.ToTable("b", True, "ALIQUOTFACTOR", "NOMCONC", "CHARHELPER1") ', "ASSAYLEVEL")
                    '20171124 LEE:
                    tblLevels = dvNL.ToTable("b", True, "ALIQUOTFACTOR", "NOMCONC", "CHARHELPER1") ', "ASSAYLEVEL")
                    intNumLevels = tblLevels.Rows.Count
                    ctQCs = intNumLevels
                    Erase DilQCFactor
                    ReDim DilQCFactor(intNumLevels)

                    tblLevels.Columns.Add("ASSAYLEVEL", Type.GetType("System.Decimal"))


                    For Count2 = 0 To tblLevels.Rows.Count - 1
                        tblLevels.Rows(Count2).BeginEdit()
                        tblLevels.Rows(Count2).Item("ASSAYLEVEL") = Count2 + 1
                        tblLevels.Rows(Count2).EndEdit()
                    Next

                    'NOTE: This AssayLevel thing can screw up the underlying data. Fix it
                    For Count2 = 0 To tblLevels.Rows.Count - 1
                        var1 = tblLevels.Rows(Count2).Item("NOMCONC")
                        var2 = tblLevels.Rows(Count2).Item("CHARHELPER1")
                        strF = "NOMCONC = " & var1 & " AND CHARHELPER1 = '" & var2 & "'"
                        dv2.RowFilter = strF
                        For Count3 = 0 To dv2.Count - 1
                            dv2(Count3).BeginEdit()
                            dv2(Count3).Item("ASSAYLEVEL") = Count2 + 1
                            dv2(Count3).EndEdit()
                        Next
                    Next
                    dv2.RowFilter = Nothing

                    '20171118 LEE: need to sort appropriately when Diln and Hi have same assay levels and/or NOMCONC
                    'Dim rowsLevel() As DataRow = tblLevels.Select("", "ASSAYLEVEL ASC, NOMCONC ASC, ALIQUOTFACTOR DESC")
                    'Dim rowsLevel() As DataRow = tblLevels.Select("", "NOMCONC ASC, ALIQUOTFACTOR DESC")
                    '20171124 LEE:
                    Dim rowsLevel() As DataRow = tblLevels.Select("", "NOMCONC ASC, ALIQUOTFACTOR DESC")

                    'establish diln qcs
                    For Count2 = 0 To intNumLevels - 1
                        'var2 = NZ(tblLevels.Rows(Count2).Item("ALIQUOTFACTOR"), 1)
                        var2 = NZ(rowsLevel(Count2).Item("ALIQUOTFACTOR"), 1)
                        '20171118 LEE:
                        'var1 = CDec(NZ(var2, 1))
                        var1 = CDec(Math.Round(NZ(var2, 1), intDFDec))
                        DilQCFactor(Count2 + 1) = var1
                        'If var1 <> 1 Then
                        '    DilQCFactor(Count2 + 1) = var1
                        'Else
                        '    DilQCFactor(Count2 + 1) = 1
                        'End If
                    Next

                    intLeg = 0
                    ctQCLegend = 0
                    ctDilLeg = 0
                    ctLegend = 0

                    'dv = frmH.dgvWatsonAnalRef.DataSource
                    'int1 = FindRowDV("LLOQ Units", dv)
                    'var2 = dv.Item(int1).Item(1)

                    'int1 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
                    'str1 = NZ(frmH.dgvStudyConfig(1, int1).Value, "")

                    'If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
                    'Else
                    '    var2 = str1
                    'End If

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

                    For Count2 = 1 To ctQCs
                        'var1 = DilQCFactor(Count2)
                        'var1 = tblLevels.Rows(Count2 - 1).Item("ALIQUOTFACTOR")
                        var1 = NZ(DilQCFactor(Count2), 1)
                        If var1 = 1 Then
                        Else
                            intLeg = intLeg + 1
                            ctDilLeg = ctDilLeg + 1
                            ctLegend = ctLegend + 1
                            'configure first legend item
                            'var4 = arrBCQCs(2, Count2)
                            var4 = tblLevels.Rows(Count2 - 1).Item("NOMCONC")
                            'var4 = Format(arrBCQCs(2, Count2), "0")
                            'var2 = Sheets("AnalRefTables").Range("LLOQUnits").Offset(0, Count1).Value
                            'var3 = Format(1 / CDec(var1), "0")
                            var3 = GetDilnFactor(CDec(var1)) '20190220 LEE
                            var1 = Chr(96 + intLeg) 'debugging
                            arrLegend(1, intLeg) = Chr(96 + intLeg) 'a,b,c,etc
                            ' arrLegend(2, intLeg) = "Dilution QCs undiluted concentration " & var4 & " " & strConcUnits & "; a 1:" & var3 & " dilution with blank matrix was performed prior to extraction and analysis."
                            Dim strAN As String = GetAN(var3)

                            arrLegend(2, intLeg) = "Dilution QCs undiluted concentration " & var4 & " " & strConcUnits & "; " & strAN & " " & var3 & "-fold dilution with blank matrix was performed prior to extraction and analysis."
                            arrLegend(3, intLeg) = True
                            arrLegend(4, intLeg) = True
                            'arrQCLegend(1, intLeg) = Chr(96 + intLeg) 'a,b,c,etc
                            'arrQCLegend(2, intLeg) = "Dilution QCs undiluted concentration " & var4 & " " & var2 & "; a 1:" & var3 & " dilution with blank matrix was performed prior to extraction and analysis."
                            'arrQCLegend(3, intLeg) = True
                            'ctQCLegend = ctQCLegend + 1
                        End If
                    Next


                    'find number of table rows to generate
                    Dim intRowsXTot As Short = 0

                    For Count2 = 0 To intNumRuns - 1
                        'enter runid
                        var10 = tblNumRuns.Rows.Item(Count2).Item("RUNID")
                        '.Selection.TypeText(CStr(var10))
                        '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)

                        intRowsX = 0
                        For Count3 = 0 To intNumLevels - 1
                            'varNom = tblLevels.Rows.Item(Count3).Item("NOMCONC")
                            'var1 = tblLevels.Rows.Item(Count3).Item("ASSAYLEVEL")
                            '20171118 LEE:
                            varNom = rowsLevel(Count3).Item("NOMCONC")
                            var1 = rowsLevel(Count3).Item("ASSAYLEVEL")
                            dv2.RowFilter = ""
                            'don't know why, but must make a long filter here or
                            'both analytes get returned in dv2.rowfilter
                            strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom & " AND ASSAYLEVEL = " & var1
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

                    'Increment for Statistics Sections
                    'intTblRows = intTblRows + countNumStatsRows()

                    Dim intCSN As Short
                    intCSN = countNumStatsRows()
                    intTblRows = intTblRows + intCSN

                    If intCSN > 0 Then
                    Else
                        intTblRows = intTblRows - 1 'subtract an unneeded blank row
                    End If

                    If boolQCREPORTACCVALUES Then
                    Else
                        ctExp = ctExp + 2 'for stats headings
                        intTblRows = intTblRows + 2 'for two titles

                        'Increment for Statistics Sections
                        intTblRows = intTblRows + intCSN
                        ctExp = ctExp + intCSN

                        If intCSN > 0 Then
                            intTblRows = intTblRows + 1 '(1 * intNumRuns) - 1 'for a blank row after each Mean/Bias/n set, except last set
                            ctExp = ctExp + 1 '(1 * intNumRuns) - 1
                        End If

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

                            strA = strAnal
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

                        .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=2, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.MoveRight Unit:=Microsoft.Office.Interop.Word.wdunits.word.wdunits.wdCharacter, Count:=ctQCs, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                        '.Selection.MoveLeft(Microsoft.Office.Interop.Word.WdUnits.wdCharacter, 1)

                        'Enter row titles
                        .Selection.Tables.Item(1).Cell(2, 2).Select()

                        '''wdd.visible = True
                        int1 = -1
                        For Count2 = 0 To (intNumLevels * int11) - 1 Step int11
                            int1 = int1 + 1
                            'var1 = arrBCQCs(3, Count2)
                            'var1 = NZ(tblLevels.Rows.Item(int1).Item("CHARHELPER1"), "")
                            '20171118 LEE:
                            var1 = NZ(rowsLevel(int1).Item("CHARHELPER1"), "")
                            var3 = ReturnStdQC(var1.ToString)
                            .Selection.TypeText(Text:=var3)
                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=int11)
                        Next

                        'enter dilution qc superscripts
                        Count3 = 0

                        Dim boolEnterDiff As Boolean

                        'For Count2 = ctDilLeg - 1 To 0 Step -1
                        int1 = 0
                        For Count2 = 1 To ctQCs
                            var1 = DilQCFactor(Count2)

                            If var1 = 1 Then
                            Else
                                Count3 = Count3 + 1
                                If int11 = 1 Then
                                    int1 = Count2 + 1
                                Else
                                    int1 = Count2 * int11
                                End If
                                .Selection.Tables.Item(1).Cell(2, int1).Select()
                                .Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCharacter)
                                .Selection.MoveLeft(Microsoft.Office.Interop.Word.WdUnits.wdCharacter)
                                '.Selection.TypeText(Text:=CStr(Chr(Count3 + 96))) 'units
                                .Selection.TypeText(" " & CStr(Chr(Count3 + 96))) 'units
                                'superscript the footnote
                                .Selection.MoveLeft(Microsoft.Office.Interop.Word.WdUnits.wdCharacter, 1, Word.WdMovementType.wdExtend)
                                .Selection.Font.Superscript = True
                                .Selection.Font.Size = 12
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)

                            End If
                        Next

                        '''''wdd.visible = True

                        'Enter nom. conc. row titles
                        .Selection.Tables.Item(1).Cell(3, 2).Select()
                        'For Count2 = 0 To intNumLevels - 1
                        int1 = -1
                        For Count2 = 0 To (intNumLevels * int11) - 1 Step int11
                            int1 = int1 + 1
                            'var1 = arrBCQCs(3, Count2)
                            'If boolLUseSigFigs Then
                            '    var1 = CStr(SigFigOrDecString(tblLevels.Rows.Item(int1).Item("NOMCONC"), LSigFig, False))
                            'Else
                            '    var1 = CStr(RoundToDecimalRAFZ(tblLevels.Rows.Item(int1).Item("NOMCONC"), LSigFig))
                            'End If
                            '20171118 LEE
                            If boolLUseSigFigs Then
                                var1 = CStr(SigFigOrDecString(rowsLevel(int1).Item("NOMCONC"), LSigFig, False))
                            Else
                                var1 = CStr(RoundToDecimalRAFZ(rowsLevel(int1).Item("NOMCONC"), LSigFig))
                            End If
                            var1 = var1 & ChrW(160) & strConcUnits
                            If intNumLevels > 5 Then
                                .Selection.Font.Size = .Selection.Font.Size - 1
                            End If
                            .Selection.TypeText(Text:=var1)
                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=1)
                            If int11 = 2 Then
                                .Selection.TypeText(Text:=ReturnDiffLabel)
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=1)
                            End If
                        Next

                        .Selection.Tables.Item(1).Cell(3, 1).Select()
                        'bottom border this row
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .Selection.Tables.Item(1).Cell(3, 1).Select()

                        'begin entering data'
                        If BOOLINCLUDEDATE Then
                            '.Selection.Tables.Item(1).Cell(2, 1).Select()
                            '.Selection.TypeText("Watson Run ID")
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
                        strA = ""
                        strB = ""

                        Dim boolSameNumRows As Boolean = True
                        Dim int2Same As Short
                        Dim vCH1

                        For Count2 = 0 To intNumRuns - 1

                            ''''''''wdd.visible = True

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

                            'find accurate introwsX
                            int2 = 0
                            intRowsX = 0
                            For Count3 = 0 To intNumLevels - 1

                                'varNom = tblLevels.Rows.Item(Count3).Item("NOMCONC")
                                'var1 = tblLevels.Rows.Item(Count3).Item("ASSAYLEVEL")
                                'vCH1 = tblLevels.Rows.Item(Count3).Item("CHARHELPER1")
                                '20171118 LEE:
                                varNom = rowsLevel(Count3).Item("NOMCONC")
                                var1 = rowsLevel(Count3).Item("ASSAYLEVEL")
                                vCH1 = rowsLevel(Count3).Item("CHARHELPER1")
                                var2 = CDec(DilQCFactor(Count3 + 1))


                                dv2.RowFilter = ""
                                'don't know why, but must make a long filter here or
                                'both analytes get returned in dv2.rowfilter

                                'If INTQCLEVELGROUP = 0 Then 'use assaylevel
                                '    strF = strF2 & " AND RUNID = " & var10 & " AND ASSAYLEVEL = " & var1
                                '    'strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND RUNID = " & int20 & " AND ASSAYLEVEL = " & var2 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                                'ElseIf INTQCLEVELGROUP = 1 Then 'use NomConc
                                '    '20171118 LEE: need aliquot factor too because sometimes NomConc is same for Diln and Hi samples
                                '    strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom & " AND ALIQUOTFACTOR = " & var2
                                'ElseIf INTQCLEVELGROUP = 2 Then 'use Level Label
                                '    strF = strF2 & " AND RUNID = " & var10 & " AND CHARHELPER1 = '" & vCH1 & "'"
                                'Else
                                '    strF = strF2 & " AND RUNID = " & var10 & " AND ASSAYLEVEL = " & var1
                                'End If


                                'Function ReturnQCGroup(boolA, boolE, boolR, a, b, c, d, e, f) As String
                                '20180823 LEE:

                                'boolA: TRUE if include aliquot factor
                                'boolE: TRUE if include ELIMINATEDFLAG
                                'boolR: TRUE if include RUNID
                                'a=rundid
                                'b=assaylevel
                                'c=aliquotfactor
                                'd=nomconc
                                'e=CHARHELPER1 or QCLabel
                                'f=ELIMINATEDFLAG
                                str1 = ReturnQCGroup(False, False, True, var10, var1, var2, varNom, vCH1, "")
                                strF = strF2 & str1

                                '******

                                'strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom & " AND ASSAYLEVEL = " & var1
                                dv2.RowFilter = strF
                                int2 = dv2.Count

                                If int2 > intRowsX Then
                                    intRowsX = int2
                                End If

                            Next

                            'start filling in data by rows

                            intD = -1

                            ''wdd.visible = True

                            For Count3 = 0 To (intNumLevels * int11) - 1 Step int11

                                intD = intD + 1
                                'varNom = tblLevels.Rows.Item(intD).Item("NOMCONC")
                                'var4 = tblLevels.Rows.Item(intD).Item("ASSAYLEVEL")
                                'vCH1 = tblLevels.Rows.Item(intD).Item("CHARHELPER1")
                                '20171118 LEE
                                varNom = rowsLevel(intD).Item("NOMCONC")
                                var4 = rowsLevel(intD).Item("ASSAYLEVEL")
                                vCH1 = rowsLevel(intD).Item("CHARHELPER1")

                                'determine hi and lo (nom*flagpercent)
                                strF = "CONCENTRATION = '" & varNom & "'"

                                'if Conc < 1, then the query return 0 records
                                'must do something different
                                'var1 = GetANALYTEFLAGPERCENT(varNom, var10, vAnalyteID)
                                var1 = GetANALYTEFLAGPERCENTAnova(varNom, var10, vAnalyteID, tblLevelCrit)
                                v1 = var1
                                v2 = var1
                                'var1 = CDec(NZ(rows10(0).Item("FLAGPERCENT"), 15))
                                arrFP(1, intD) = var1
                                arrFP(2, intD) = var1
                                Call SetHighAndLowCriteria(varNom, var1, var1, hi, lo)

                                'start entering data
                                dv2.RowFilter = ""

                                'don't know why, but must make a long filter here or
                                'both analytes get returned in dv2.rowfilter
                                var1 = CDec(DilQCFactor(intD + 1))

                                'If INTQCLEVELGROUP = 0 Then 'use assaylevel
                                '    strF = strF2 & " AND RUNID = " & var10 & " AND ASSAYLEVEL = " & var4 & " AND ALIQUOTFACTOR = " & var1
                                '    'strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND RUNID = " & int20 & " AND ASSAYLEVEL = " & var2 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                                'ElseIf INTQCLEVELGROUP = 1 Then 'use NomConc
                                '    strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom & " AND ALIQUOTFACTOR = " & var1
                                'ElseIf INTQCLEVELGROUP = 2 Then 'use Level Label
                                '    strF = strF2 & " AND RUNID = " & var10 & " AND CHARHELPER1 = '" & vCH1 & "' AND ALIQUOTFACTOR = " & var1
                                'Else
                                '    strF = strF2 & " AND RUNID = " & var10 & " AND ASSAYLEVEL = " & var4 & " AND ALIQUOTFACTOR = " & var1
                                'End If

                                'Function ReturnQCGroup(boolA, boolE, boolR, a, b, c, d, e, f) As String
                                '20180823 LEE:

                                'boolA: TRUE if include aliquot factor
                                'boolE: TRUE if include ELIMINATEDFLAG
                                'boolR: TRUE if include RUNID
                                'a=rundid
                                'b=assaylevel
                                'c=aliquotfactor
                                'd=nomconc
                                'e=CHARHELPER1 or QCLabel
                                'f=ELIMINATEDFLAG

                                str1 = ReturnQCGroup(True, False, True, var10, var4, var1, varNom, vCH1, "")
                                strF = strF2 & str1


                                'strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom & " AND ALIQUOTFACTOR = " & var1
                                '''''''''''console.writeline(strF)
                                Try
                                    dv2.RowFilter = strF
                                    int2 = dv2.Count
                                Catch ex As Exception
                                    'MsgBox(ex.Message & ChrW(10) & strF)
                                    int2 = 0
                                    '''''''''''console.writeline("Previous strF is bad")
                                    'Exit Sub
                                End Try
                                'dv2.RowFilter = strF
                                'int2 = dv2.Count

                                'create rows1 from tbl1 which will contain data
                                strF = ""
                                If int2 = 0 Then
                                Else
                                    strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idTR & " AND ANALYTEID = " & vAnalyteID
                                End If
                                For Count4 = 0 To int2 - 1
                                    If Count4 = 0 Then
                                        str1 = " AND ((RUNID = " & var10 & " AND RUNSAMPLEORDERNUMBER = " & var4 & ") OR "
                                    ElseIf Count4 = int2 - 1 Then
                                        str1 = " (RUNID = " & var10 & " AND RUNSAMPLEORDERNUMBER = " & var4 & "))"
                                    Else
                                        str1 = " (RUNID = " & var10 & " AND RUNSAMPLEORDERNUMBER = " & var4 & ") OR "
                                    End If
                                Next

                                Erase rows1
                                If Len(strF) = 0 Then
                                    strF = "(RUNID = 0 AND ANALYTEINDEX = 0 AND MASTERASSAYID = 0 AND ANALYTEID = 0 AND RUNSAMPLEORDERNUMBER = 0)"
                                End If

                                'rows1 = tbl1.Select(strF)
                                Dim tbl2R As System.Data.DataTable = dv2.ToTable
                                strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                                rows1 = tbl2R.Select(strF, strS)
                                int3 = rows1.Length

                                'DO hi / lo again
                                'varNom = tblLevels.Rows.Item(intD).Item("NOMCONC")
                                '20171118 LEE
                                varNom = rowsLevel(intD).Item("NOMCONC")
                                'determine hi and lo (nom*flagpercent)
                                strF = "CONCENTRATION = '" & varNom & "'"
                                'here
                                'rows10 = tblBCQCs.Select(strF)

                                'determine hi and lo (nom*flagpercent)
                                'if Conc < 1, then the query return 0 records
                                'must do something different
                                'if Conc < 1, then the query return 0 records
                                'must do something different



                                If int3 = 0 Then
                                    var2 = 15 'CDec(NZ(rows1(0).Item("NUMMINACCCRIT"), 0))
                                    var3 = 15 'CDec(NZ(rows1(0).Item("NUMMAXACCCRIT"), 0))
                                    vU = 0 'CDec(NZ(rows1(0).Item("BOOLUSEGUWUACCCRIT"), 0))
                                    If gAllowGuWuAccCrit And LAllowGuWuAccCrit And vU = -1 Then
                                        arrFP(1, intD) = var3
                                        arrFP(2, intD) = var2
                                        Call SetHighAndLowCriteria(varNom, var3, var2, hi, lo)
                                        v1 = var2
                                        v2 = var3
                                    Else
                                        arrFP(1, intD) = v1
                                        arrFP(2, intD) = v1
                                        Call SetHighAndLowCriteria(varNom, v1, v1, hi, lo)
                                        v1 = v1
                                        v2 = v1
                                    End If
                                Else
                                    'var1 = CDec(NZ(rows1(0).Item("ANALYTEFLAGPERCENT"), 15))
                                    'var1 = CDec(NZ(rows1(0).Item("FLAGPERCENT"), 15))
                                    'var1 = GetANALYTEFLAGPERCENT(varNom, var10, vAnalyteID)
                                    var1 = GetANALYTEFLAGPERCENTAnova(varNom, var10, vAnalyteID, tblLevelCrit)
                                    var2 = CDec(NZ(rows1(0).Item("NUMMINACCCRIT"), 0))
                                    var3 = CDec(NZ(rows1(0).Item("NUMMAXACCCRIT"), 0))

                                    vU = CDec(NZ(rows1(0).Item("BOOLUSEGUWUACCCRIT"), 0))
                                    If gAllowGuWuAccCrit And LAllowGuWuAccCrit And vU = -1 Then
                                        arrFP(1, intD) = var3
                                        arrFP(2, intD) = var2
                                        Call SetHighAndLowCriteria(varNom, var3, var2, hi, lo)
                                        v1 = var2
                                        v2 = var3
                                    Else
                                        arrFP(1, intD) = var1
                                        arrFP(2, intD) = var1
                                        Call SetHighAndLowCriteria(varNom, var1, var1, hi, lo)
                                        v1 = var1
                                        v2 = var1
                                    End If

                                End If

                                var1 = var1 'debug


                                ''''wdd.visible = True

                                'HereD
                                'enter data
                                For Count4 = 0 To intRowsX - 1 'int3 - 1

                                    boolEnterDiff = True
                                    boolOC = False

                                    .Selection.Tables.Item(1).Cell(int1 + Count4, (intD * int11) + 2).Select()

                                    ''''wdd.visible = True

                                    If Count4 > int3 - 1 Then
                                        If boolQCNA Then
                                            str1 = "NA"
                                        Else
                                            str1 = ""
                                        End If

                                        .Selection.TypeText(str1)
                                        boolEnterDiff = False
                                    Else
                                        'var1 = NZ(rows1(Count4).Item("CONCENTRATION"), "GAA")
                                        'If StrComp(var1, "GAA", CompareMethod.Text) = 0 Then
                                        '    var1 = "GAA"
                                        'End If
                                        '20160510 LEE: NO! Do not set to 0!!
                                        'var1 = NZ(rows1(Count4).Item("CONCENTRATION"), 0)
                                        var1 = rows1(Count4).Item("CONCENTRATION")
                                        'numDF = rows1(Count4).Item("ALIQUOTFACTOR")
                                        numDF = rows1(Count4).Item("ALIQUOTFACTOR")
                                        If IsDBNull(var1) Then
                                            var2 = var1
                                        Else
                                            var1 = var1 / numDF
                                            If boolLUseSigFigs Then
                                                var2 = SigFigOrDec(var1, LSigFig, False)
                                            Else
                                                var2 = RoundToDecimalRAFZ(var1, LSigFig)
                                            End If
                                        End If
                                        varConc = var2

                                        var1 = NZ(rows1(Count4).Item("ELIMINATEDFLAG"), "N")
                                        var3 = NZ(rows1(Count4).Item("BOOLEXCLSAMPLE"), 0)
                                        If StrComp(var1, "Y", CompareMethod.Text) = 0 Then
                                            boolExFromAS = False
                                        Else
                                            If gAllowExclSamples And LAllowExclSamples Then
                                                If var3 = -1 Then
                                                    var1 = "Y"
                                                    boolExFromAS = True
                                                Else
                                                    'var1 = "N"
                                                    'don't assign "N", Watson may override
                                                End If
                                            End If
                                        End If

                                        If IsDBNull(var2) Then
                                            var1 = "Y"
                                        End If

                                        'add rows to tblAnova
                                        'Dim rowsAn As DataRow = tblAnova.NewRow
                                        'rowsAn.Item("Group") = var10
                                        'rowsAn.Item("Conc") = var2
                                        'rowsAn.Item("NomConc") = varNom
                                        'tblAnova.Rows.Add(rowsAn)

                                        'add rows to tblz
                                        'NomConc
                                        'Conc
                                        'ALIQUOTFACTOR
                                        'ELIMINATEDFLAG
                                        'BOOLOUTLIER
                                        'HI
                                        'LO

                                        Dim rowsz As DataRow = tblZ.NewRow
                                        rowsz.BeginEdit()
                                        rowsz.Item("NomConc") = varNom
                                        rowsz.Item("Conc") = var2
                                        rowsz.Item("ALIQUOTFACTOR") = numDF
                                        rowsz.Item("ELIMINATEDFLAG") = var1
                                        If StrComp(var1, "Y", vbTextCompare) = 0 And IsDBNull(var2) = False Then
                                            rowsz.Item("BOOLOUTLIER") = True
                                        Else
                                            rowsz.Item("BOOLOUTLIER") = False
                                        End If
                                        rowsz.Item("HI") = hi
                                        rowsz.Item("LO") = lo
                                        rowsz.Item("v1") = v1
                                        rowsz.Item("v2") = v2
                                        rowsz.Item("CHARHELPER1") = vCH1
                                        rowsz.Item("ASSAYLEVEL") = var4
                                        rowsz.EndEdit()
                                        tblZ.Rows.Add(rowsz)

                                        'Note: earlier var1 has been reassigned according to LAllowGuWuAccCrit
                                        If StrComp(var1, "Y", vbTextCompare) = 0 And IsDBNull(var2) = False Then

                                            intExp = intExp + 1
                                            intLeg = intLeg + 1
                                            strA = ChrW(intLeg + intLegStart)
                                            boolOC = True

                                            'Remember, tblAssignedSamples does not have DECISIONREASON
                                            var6 = GetDECISIONREASONValue(boolExFromAS, vAnalyteID, rows1(Count4))

                                            ''Set Legend String
                                            str1 = GetLegendStringExcluded(arrFP(1, intD), arrFP(2, intD), vU, var6, intTableID, True, "")
                                            ''Add to Legend Array
                                            ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                            If boolRedBoldFont Then
                                                .Selection.Font.Bold = True
                                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                            End If

                                            If boolLUseSigFigs Then
                                                .Selection.TypeText(Text:=DisplayNum(var2, LSigFig, False))
                                            Else
                                                .Selection.TypeText(Text:=Format(var2, GetRegrDecStr(LSigFig)))
                                            End If

                                            Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                            '.Selection.TypeText Text:="NR"

                                            boolEnterDiff = True 'False

                                        Else

                                            'determine if value is outside acceptance criteria
                                            If IsDBNull(varConc) Then

                                                boolOC = True
                                                intExp = intExp + 1
                                                intLeg = intLeg + 1
                                                strA = ChrW(intLeg + intLegStart)

                                                'Remember, tblAssignedSamples does not have DECISIONREASON
                                                var6 = GetDECISIONREASONValue(boolExFromAS, vAnalyteID, rows1(Count4))

                                                var6 = "No Value: " & NZ(var6, "No reason recorded.")
                                                var1 = "Y"

                                                ''Set Legend String
                                                str1 = GetLegendStringExcluded(arrFP(1, intD), arrFP(2, intD), vU, var6, intTableID, True, "")
                                                ''Add to Legend Array
                                                ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                                If boolRedBoldFont Then
                                                    .Selection.Font.Bold = True
                                                    .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                End If

                                                .Selection.TypeText(Text:="NV")
                                                Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                                '.Selection.TypeText Text:="NR"

                                                boolEnterDiff = True 'False

                                                'If var2 > hi Or var2 < lo Then 'flag
                                            ElseIf OutsideAccCrit(varConc, varNom, v1, v2, NZ(vU, 0)) Then

                                                intLeg = intLeg + 1
                                                strA = ChrW(intLeg + intLegStart)
                                                'var1 = arrFP(intD)

                                                'Set Legend String
                                                str1 = GetLegendStringIncluded(arrFP(1, intD), arrFP(2, intD), vU)
                                                'Add to Legend Array
                                                ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                                If boolRedBoldFont Then
                                                    .Selection.Font.Bold = True
                                                    .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                End If

                                                If boolLUseSigFigs Then
                                                    .Selection.TypeText(Text:=CStr(DisplayNum(var2, LSigFig, False)))
                                                Else
                                                    .Selection.TypeText(Text:=Format(var2, GetRegrDecStr(LSigFig)))
                                                End If


                                                Call typeInSuperscriptFontSize12WithSpace(wd, strA)

                                                boolEnterDiff = True

                                            Else

                                                If boolLUseSigFigs Then
                                                    .Selection.TypeText(Text:=CStr(DisplayNum(var2, LSigFig, False)))
                                                Else
                                                    .Selection.TypeText(Text:=Format(var2, GetRegrDecStr(LSigFig)))
                                                End If

                                                boolEnterDiff = True
                                            End If
                                        End If

                                    End If

                                    If boolSTATSDIFFCOL Then
                                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                        If boolEnterDiff Then
                                            If boolLUseSigFigs Then
                                                var1 = SigFigOrDec(var2, LSigFig, False)
                                            Else
                                                var1 = RoundToDecimal(var2, LSigFig)
                                            End If
                                            'var3 = Format(((var1 / varNom) - 1) * 100, strQCDec)
                                            'var3 = Format(RoundToDecimal(((var1 / varNom) - 1) * 100, intQCDec), strQCDec)
                                            If boolTHEORETICAL Then
                                                var3 = CalcREPercent(var1, varNom, intQCDec)
                                                numTheor = 100 + CDec(var3)
                                                var3 = numTheor

                                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, intD + 1, "Accuracy", numTheor, CSng(var10), Count1, strDo, v1, v2, boolOC)
                                            Else
                                                var3 = Format(RoundToDecimal(CalcREPercent(var1, varNom, intQCDec), intQCDec), strQCDec)

                                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, intD + 1, "Accuracy", var3, CSng(var10), Count1, strDo, v1, v2, boolOC)

                                            End If

                                            '20180430 LEE:
                                            'must account for Endogenous Cmpds, NomConc = 0
                                            var3 = NomConcZero(varNom, var3)

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

                                Next

                                var1 = var1 'debugging

                            Next

                            'increase row position counter
                            If Count2 = intNumRuns - 1 Then
                                int1 = int1 + intRowsX + 1 '4
                            Else
                                int1 = int1 + intRowsX + 1 '5
                            End If

                        Next

                        'now enter Mean/SD/Precision/Bias/n
                        '**

                        intD = -1
                        Dim vAL
                        For Count3 = 0 To (intNumLevels * int11) - 1 Step int11
                            intD = intD + 1
                            'varNom = tblLevels.Rows.Item(intD).Item("NOMCONC")
                            'varAliquot = tblLevels.Rows.Item(intD).Item("ALIQUOTFACTOR")
                            'vCH1 = tblLevels.Rows.Item(intD).Item("CHARHELPER1")
                            'vAL = tblLevels.Rows.Item(intD).Item("ASSAYLEVEL")
                            '20171118 LEE:
                            varNom = rowsLevel(intD).Item("NOMCONC")
                            varAliquot = rowsLevel(intD).Item("ALIQUOTFACTOR")
                            vCH1 = rowsLevel(intD).Item("CHARHELPER1")
                            vAL = rowsLevel(intD).Item("ASSAYLEVEL")

                            If Count3 = 0 Then

                                int8 = 0
                                intRowsX = -1
                                If boolSTATSMEAN Then
                                    If boolQCREPORTACCVALUES Then
                                    Else
                                        If intExp = 0 Then
                                        Else
                                            int8 = int8 + 1
                                        End If
                                    End If
                                End If

                                typeStatsLabels(wd, int8, int1 + intRowsX, 1, False)

                                If boolQCREPORTACCVALUES Then

                                Else

                                    If intExp = 0 Then
                                    Else


                                        intRowsX = int8 + 1
                                        int8 = 0

                                        typeStatsLabels(wd, int8, int1 + intRowsX, 1, False)
                                    End If

                                End If

                            End If

                            If boolQCREPORTACCVALUES Then
                                intRowsX = -1
                            Else
                                If intExp = 0 Then
                                Else
                                    intRowsX = 0
                                    If Count3 = 0 Then
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX, 2).Select()
                                        .Selection.TypeText(Text:="Summary Statistics Excluding Outlier Values")
                                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                        Try
                                            .Selection.Cells.Merge()
                                        Catch ex As Exception

                                        End Try
                                        With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                            .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                        End With
                                    End If
                                End If
                            End If

                            'strF = "NomConc = " & CDbl(varNom) & " AND ELIMINATEDFLAG = 'N'"
                            ''need assay level
                            'If INTQCLEVELGROUP = 0 Then 'use assaylevel
                            '    strF = "ASSAYLEVEL = " & vAL & " AND ALIQUOTFACTOR = " & NZ(varAliquot, 1) & " AND ELIMINATEDFLAG = 'N'"
                            '    'strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND RUNID = " & int20 & " AND ASSAYLEVEL = " & var2 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                            'ElseIf INTQCLEVELGROUP = 1 Then 'use NomConc
                            '    strF = "NOMCONC = " & varNom & " AND ALIQUOTFACTOR = " & NZ(varAliquot, 1) & " AND ELIMINATEDFLAG = 'N'"
                            'ElseIf INTQCLEVELGROUP = 2 Then 'use Level Label
                            '    strF = "CHARHELPER1 = '" & vCH1 & "' AND ALIQUOTFACTOR = " & NZ(varAliquot, 1) & " AND ELIMINATEDFLAG = 'N'"
                            'Else
                            '    strF = "ASSAYLEVEL = " & vAL & " AND ALIQUOTFACTOR = " & NZ(varAliquot, 1) & " AND ELIMINATEDFLAG = 'N'"
                            'End If

                            'strF = "NomConc = " & CDbl(varNom) & " AND ELIMINATEDFLAG = 'N' AND ALIQUOTFACTOR = " & NZ(varAliquot, 1)

                            'Function ReturnQCGroup(boolA, boolE, boolR, a, b, c, d, e, f) As String
                            '20180823 LEE:

                            'boolA: TRUE if include aliquot factor
                            'boolE: TRUE if include ELIMINATEDFLAG
                            'boolR: TRUE if include RUNID
                            'a=rundid
                            'b=assaylevel
                            'c=aliquotfactor
                            'd=nomconc
                            'e=CHARHELPER1 or QCLabel
                            'f=ELIMINATEDFLAG

                            str1 = ReturnQCGroup(True, True, False, 0, vAL, varAliquot, varNom, vCH1, "N")
                            strF = str1

                            Erase rows1
                            'here
                            'strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                            rows1 = tblZ.Select(strF)
                            int2 = rows1.Length
                            'Dim numMean As Decimal
                            'Dim numBias As Decimal
                            'Dim numSD As Decimal
                            'Dim numPrec As Decimal

                            int8 = 0
                            Try
                                If rows1.Length = 0 Then
                                    var1 = "NA"
                                    numMean = 0
                                Else
                                    'var1 = MeanDR(rows1, "Conc", False, "ALIQUOTFACTOR", True, False) 'boolAliqu = FALSE because tblZ has corrected values
                                    var1 = MeanDR(rows1, "Conc", False, "ALIQUOTFACTOR", True, False) 'boolAliqu = FALSE because tblZ has corrected values
                                    If boolLUseSigFigs Then
                                        numMean = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                    Else
                                        numMean = RoundToDecimalRAFZ(var1, LSigFig)
                                    End If

                                    Call InsertQCTables(intTableID, idTR, charFCID, varNom, intD + 1, "Mean", numMean, CSng(var10), Count1, strDo, 0, 0, False)
                                End If
                            Catch ex As Exception

                            End Try
                            If boolSTATSMEAN Then
                                Try
                                    'enter Mean
                                    int8 = int8 + 1
                                    .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()

                                    'determine if value is outside acceptance criteria
                                    'if rows1.length > 1 then look for highest value
                                    If rows1.Length > 1 Then
                                        var1 = 0
                                        For Count2 = 0 To rows1.Length - 1
                                            var2 = NZ(rows1(Count2).Item("v1"), 0)
                                            If IsNumeric(var2) Then
                                                If var2 > var1 Then
                                                    hi = rows1(Count2).Item("HI")
                                                    lo = rows1(Count2).Item("LO")
                                                    v1 = rows1(Count2).Item("v1")
                                                    v2 = rows1(Count2).Item("v2")
                                                    var1 = var2
                                                End If
                                            End If
                                        Next
                                    Else
                                        hi = rows1(0).Item("HI")
                                        lo = rows1(0).Item("LO")
                                        v1 = rows1(0).Item("v1")
                                        v2 = rows1(0).Item("v2")
                                    End If

                                    'If (numMean > hi Or numMean < lo) And boolFootNoteQCMean Then 'flag
                                    If (OutsideAccCrit(numMean, varNom, v1, v2, NZ(vU, 0))) And boolFootNoteQCMean Then 'flag
                                        intLeg = intLeg + 1
                                        strA = ChrW(intLeg + intLegStart)

                                        'Set Legend String
                                        str1 = GetLegendStringIncluded(arrFP(1, intD), arrFP(2, intD), vU)
                                        'Add to Legend Array
                                        ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                        If boolRedBoldFont Then
                                            .Selection.Font.Bold = True
                                            .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                        End If

                                        If boolLUseSigFigs Then
                                            .Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                        Else
                                            .Selection.TypeText(Text:=Format(numMean, GetRegrDecStr(LSigFig)))
                                        End If

                                        Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                        boolEnterDiff = True
                                    Else
                                        If boolLUseSigFigs Then
                                            .Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                        Else
                                            .Selection.TypeText(Text:=Format(numMean, GetRegrDecStr(LSigFig)))
                                        End If

                                        boolEnterDiff = True
                                    End If



                                Catch ex As Exception

                                End Try
                            End If

                            Try
                                If rows1.Length = 0 Then
                                    var1 = "NA"
                                    numSD = 0
                                Else
                                    'var1 = StdDevDR(rows1, "Conc", False, "ALIQUOTFACTOR", True, False) 'boolAliqu = FALSE because tblZ has corrected values
                                    var1 = StdDevDR(rows1, "Conc", False, "ALIQUOTFACTOR", True, False) 'boolAliqu = FALSE because tblZ has corrected values
                                    If boolLUseSigFigs Then
                                        numSD = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                    Else
                                        numSD = RoundToDecimalRAFZ(var1, LSigFig)
                                    End If

                                End If
                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, intD + 1, "SD", numSD, CSng(var10), Count1, strDo, 0, 0, False)
                            Catch ex As Exception

                            End Try
                            If boolSTATSSD Then
                                Try
                                    'enter SD
                                    int8 = int8 + 1
                                    .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                    If int2 < gSDMax Then
                                        .Selection.TypeText("NA")
                                    Else
                                        .Selection.TypeText(CStr(numSD))
                                    End If

                                Catch ex As Exception

                                End Try
                            End If

                            Try
                                If numMean = 0 Then
                                    numPrec = 0
                                Else
                                    If int2 < gSDMax Then
                                    Else
                                        numPrec = CalcCVPercent(numSD, numMean, intQCDec)
                                        Call InsertQCTables(intTableID, idTR, charFCID, varNom, intD + 1, "Precision", numPrec, CSng(var10), Count1, strDo, 0, 0, False)
                                    End If

                                End If

                            Catch ex As Exception

                            End Try

                            If boolSTATSCV Then
                                Try
                                    'enter %CV
                                    int8 = int8 + 1
                                    .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                    If int2 < gSDMax Then
                                        .Selection.TypeText("NA")
                                    Else
                                        .Selection.TypeText(Text:=Format(numPrec, strQCDec))
                                    End If

                                Catch ex As Exception

                                End Try
                            End If

                            If varNom <= 0 Then

                            Else

                            End If

                            If boolSTATSBIAS And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then

                                Try
                                    numBias = CalcREPercent(numMean, varNom, intQCDec)

                                    If int2 = 0 Then
                                    Else
                                        Call InsertQCTables(intTableID, idTR, charFCID, varNom, intD + 1, "Accuracy", numBias, CSng(var10), Count1, strDo, 0, 0, False)
                                    End If

                                Catch ex As Exception

                                End Try

                            Else
                                'get numbias from average of %Bias columns
                                numBias = GetBiasFromDiffCol(idTR, varNom, intD + 1, 0, False)
                            End If


                            If boolSTATSBIAS And boolSTATSMEAN Then
                                Try
                                    'record %theor
                                    int8 = int8 + 1
                                    .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                    'If varNom <= 0 Then
                                    '    .Selection.TypeText("NA")
                                    'Else
                                    '    .Selection.TypeText(Text:=Format(numBias, strQCDec))
                                    'End If

                                    '.Selection.TypeText(Text:=Format(numBias, strQCDec))
                                    '20180430 LEE:
                                    'must account for Endogenous Cmpds, NomConc = 0
                                    .Selection.TypeText(Text:=NomConcZero(varNom, numBias))

                                Catch ex As Exception

                                End Try
                            End If

                            If boolTHEORETICAL And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                Try
                                    numTheor = CalcREPercent(numMean, varNom, intQCDec)
                                    numTheor = 100 + CDec(numTheor)
                                    If int2 = 0 Then
                                    Else
                                        Call InsertQCTables(intTableID, idTR, charFCID, varNom, intD + 1, "Accuracy", numTheor, CSng(var10), Count1, strDo, 0, 0, False)
                                    End If

                                Catch ex As Exception

                                End Try
                            Else
                                'get numbias from average of %Bias columns
                                numTheor = GetBiasFromDiffCol(idTR, varNom, intD + 1, 0, False)
                                numTheor = 100 + CDec(numTheor)
                            End If
                            If boolTHEORETICAL And boolSTATSMEAN Then
                                Try
                                    'record %theor
                                    int8 = int8 + 1
                                    .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                    '.Selection.TypeText(Text:=Format(NomConcZero(varNom, numTheor), strQCDec))
                                    '20180430 LEE:
                                    'must account for Endogenous Cmpds, NomConc = 0
                                    .Selection.TypeText(Text:=NomConcZero(varNom, numTheor))
                                Catch ex As Exception

                                End Try
                            End If

                            If boolSTATSDIFF And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                Try
                                    numBias = CalcREPercent(numMean, varNom, intQCDec)
                                    If int2 = 0 Then
                                    Else
                                        Call InsertQCTables(intTableID, idTR, charFCID, varNom, intD + 1, "Accuracy", numBias, CSng(var10), Count1, strDo, 0, 0, False)
                                    End If

                                Catch ex As Exception

                                End Try
                            Else
                                'get numbias from average of %Bias columns
                                numBias = GetBiasFromDiffCol(idTR, varNom, intD + 1, 0, False)
                            End If
                            If boolSTATSDIFF And boolSTATSMEAN Then
                                Try
                                    'record %Diff
                                    int8 = int8 + 1
                                    .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                    '.Selection.TypeText(Text:=Format(numBias, strQCDec))
                                    '20180430 LEE:
                                    'must account for Endogenous Cmpds, NomConc = 0
                                    .Selection.TypeText(Text:=NomConcZero(varNom, numBias))

                                Catch ex As Exception

                                End Try
                            End If

                            If BOOLSTATSRE And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                Try
                                    numBias = CalcREPercent(numMean, varNom, intQCDec)
                                    If int2 = 0 Then
                                    Else
                                        Call InsertQCTables(intTableID, idTR, charFCID, varNom, intD + 1, "Accuracy", numBias, CSng(var10), Count1, strDo, 0, 0, False)
                                    End If

                                Catch ex As Exception

                                End Try
                            Else
                                'get numbias from average of %Bias columns
                                numBias = GetBiasFromDiffCol(idTR, varNom, intD + 1, 0, False)
                            End If

                            If BOOLSTATSRE And boolSTATSMEAN Then
                                Try
                                    'record %RE
                                    int8 = int8 + 1
                                    .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                    '.Selection.TypeText(Text:=Format(numBias, strQCDec))
                                    '20180430 LEE:
                                    'must account for Endogenous Cmpds, NomConc = 0
                                    .Selection.TypeText(Text:=NomConcZero(varNom, numBias))
                                Catch ex As Exception

                                End Try
                            End If

                            Try
                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, intD + 1, "n", int2, CSng(var10), Count1, strDo, 0, 0, False)
                            Catch ex As Exception

                            End Try
                            If boolSTATSN Then
                                Try
                                    'enter n
                                    int8 = int8 + 1
                                    .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                    .Selection.TypeText(CStr(int2))

                                Catch ex As Exception

                                End Try
                            End If

                            If boolQCREPORTACCVALUES Then


                            Else
                                If intExp = 0 Then
                                Else
                                    intRowsX = int8 + 2
                                    If Count3 = 0 Then
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX, 2).Select()
                                        .Selection.TypeText(Text:="Summary Statistics Including Outlier Values")
                                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                        Try
                                            .Selection.Cells.Merge()
                                        Catch ex As Exception

                                        End Try
                                        With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                            .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                        End With
                                    End If

                                    'If INTQCLEVELGROUP = 0 Then 'use assaylevel
                                    '    strF = "ASSAYLEVEL = " & vAL & " AND ALIQUOTFACTOR = " & NZ(varAliquot, 1)
                                    '    'strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND RUNID = " & int20 & " AND ASSAYLEVEL = " & var2 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                                    'ElseIf INTQCLEVELGROUP = 1 Then 'use NomConc
                                    '    strF = "NOMCONC = " & varNom & " AND ALIQUOTFACTOR = " & NZ(varAliquot, 1)
                                    'ElseIf INTQCLEVELGROUP = 2 Then 'use Level Label
                                    '    strF = "CHARHELPER1 = '" & vCH1 & "' AND ALIQUOTFACTOR = " & NZ(varAliquot, 1)
                                    'Else
                                    '    strF = "ASSAYLEVEL = " & vAL & " AND ALIQUOTFACTOR = " & NZ(varAliquot, 1)
                                    'End If

                                    'Function ReturnQCGroup(boolA, boolE, boolR, a, b, c, d, e, f) As String
                                    '20180823 LEE:

                                    'boolA: TRUE if include aliquot factor
                                    'boolE: TRUE if include ELIMINATEDFLAG
                                    'boolR: TRUE if include RUNID
                                    'a=rundid
                                    'b=assaylevel
                                    'c=aliquotfactor
                                    'd=nomconc
                                    'e=CHARHELPER1 or QCLabel
                                    'f=ELIMINATEDFLAG

                                    str1 = ReturnQCGroup(True, False, False, 0, vAL, varAliquot, varNom, vCH1, "")
                                    strF = str1

                                    'strF = "NomConc = " & CDec(varNom) ' & " AND ELIMINATEDFLAG = 'N'"
                                    Erase rows1
                                    rows1 = tblZ.Select(strF)
                                    int2 = rows1.Length

                                    intRowsX = int8 + 2 '7

                                    int8 = 0

                                    If boolSTATSMEAN Then
                                        Try
                                            'enter Mean
                                            int8 = int8 + 1
                                            .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                            If rows1.Length = 0 Then
                                                numMean = 0
                                            Else
                                                'var1 = MeanDR(rows1, "Conc", True, "ALIQUOTFACTOR", True, False)
                                                'pass FALSE, means already corrected
                                                var1 = MeanDR(rows1, "Conc", False, "ALIQUOTFACTOR", True, False)
                                                If boolLUseSigFigs Then
                                                    numMean = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                                Else
                                                    numMean = RoundToDecimalRAFZ(var1, LSigFig)
                                                End If

                                            End If

                                            '.Selection.TypeText(CStr(numMean))

                                            'determine if value is outside acceptance criteria
                                            hi = rows1(0).Item("HI")
                                            lo = rows1(0).Item("LO")
                                            v1 = rows1(0).Item("v1")
                                            v2 = rows1(0).Item("v2")
                                            'If (numMean > hi Or numMean < lo) And boolFootNoteQCMean Then 'flag
                                            If (OutsideAccCrit(numMean, varNom, v1, v2, NZ(vU, 0))) And boolFootNoteQCMean Then 'flag
                                                intLeg = intLeg + 1
                                                strA = ChrW(intLeg + intLegStart)

                                                'Set Legend String
                                                str1 = GetLegendStringIncluded(v1, v2, vU)
                                                'Add to Legend Array
                                                ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                                If boolRedBoldFont Then
                                                    .Selection.Font.Bold = True
                                                    .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                End If

                                                If boolLUseSigFigs Then
                                                    .Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                                Else
                                                    .Selection.TypeText(Text:=Format(numMean, GetRegrDecStr(LSigFig)))
                                                End If

                                                Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                                boolEnterDiff = True
                                            Else
                                                If boolLUseSigFigs Then
                                                    .Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                                Else
                                                    .Selection.TypeText(Text:=Format(numMean, GetRegrDecStr(LSigFig)))
                                                End If
                                                boolEnterDiff = True
                                            End If

                                        Catch ex As Exception

                                        End Try
                                    End If
                                    If boolSTATSSD Then
                                        Try
                                            'enter SD
                                            int8 = int8 + 1
                                            .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                            If int2 < gSDMax Then
                                                .Selection.TypeText("NA")
                                            Else
                                                If rows1.Length = 0 Then
                                                    numSD = 0
                                                Else
                                                    'pass FALSE, means already corrected
                                                    'var1 = StdDevDR(rows1, "Conc", True, "ALIQUOTFACTOR", True, False)
                                                    var1 = StdDevDR(rows1, "Conc", False, "ALIQUOTFACTOR", True, False)
                                                    If boolLUseSigFigs Then
                                                        numSD = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                                    Else
                                                        numSD = RoundToDecimalRAFZ(var1, LSigFig)
                                                    End If

                                                End If
                                                .Selection.TypeText(CStr(numSD))
                                            End If

                                        Catch ex As Exception

                                        End Try
                                    End If
                                    If boolSTATSCV Then
                                        Try
                                            'enter %CV
                                            int8 = int8 + 1
                                            .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                            If int2 < gSDMax Then
                                                .Selection.TypeText("NA")
                                            Else
                                                If numMean = 0 Then
                                                    numPrec = 0
                                                Else
                                                    numPrec = CalcCVPercent(numSD, numMean, intQCDec)
                                                End If
                                                .Selection.TypeText(Text:=Format(numPrec, strQCDec))
                                            End If

                                        Catch ex As Exception

                                        End Try
                                    End If

                                    If boolSTATSBIAS And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                        Try
                                            numBias = CalcREPercent(numMean, varNom, intQCDec)
                                        Catch ex As Exception

                                        End Try
                                    Else
                                        'get numbias from average of %Bias columns
                                        numBias = GetBiasFromDiffCol(idTR, varNom, intD + 1, 0, True)
                                    End If

                                    If boolSTATSBIAS And boolSTATSMEAN Then
                                        Try
                                            'record %Bias
                                            int8 = int8 + 1
                                            .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                            '.Selection.TypeText(Text:=Format(numBias, strQCDec))
                                            '20180430 LEE:
                                            'must account for Endogenous Cmpds, NomConc = 0
                                            .Selection.TypeText(Text:=NomConcZero(varNom, numBias))

                                        Catch ex As Exception

                                        End Try
                                    End If

                                    If boolTHEORETICAL And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                        Try
                                            numTheor = CalcREPercent(numMean, varNom, intQCDec)
                                            numTheor = 100 + CDec(numTheor)
                                        Catch ex As Exception

                                        End Try
                                    Else
                                        'get numbias from average of %Bias columns
                                        numTheor = GetBiasFromDiffCol(idTR, varNom, intD + 1, 0, True)
                                        numTheor = 100 + CDec(numTheor)
                                    End If

                                    If boolTHEORETICAL And boolSTATSMEAN Then
                                        Try
                                            'record %Bias
                                            int8 = int8 + 1
                                            .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                            '.Selection.TypeText(Text:=Format(numTheor, strQCDec))
                                            '20180430 LEE:
                                            'must account for Endogenous Cmpds, NomConc = 0
                                            .Selection.TypeText(Text:=NomConcZero(varNom, numTheor))
                                        Catch ex As Exception

                                        End Try
                                    End If


                                    If boolSTATSDIFF And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                        Try
                                            numBias = CalcREPercent(numMean, varNom, intQCDec)
                                        Catch ex As Exception

                                        End Try
                                    Else
                                        'get numbias from average of %Bias columns
                                        numBias = GetBiasFromDiffCol(idTR, varNom, intD + 1, 0, True)
                                    End If

                                    If boolSTATSDIFF And boolSTATSMEAN Then
                                        Try
                                            'record %Bias
                                            int8 = int8 + 1
                                            .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                            '.Selection.TypeText(Text:=Format(numBias, strQCDec))
                                            '20180430 LEE:
                                            'must account for Endogenous Cmpds, NomConc = 0
                                            .Selection.TypeText(Text:=NomConcZero(varNom, numBias))
                                        Catch ex As Exception

                                        End Try
                                    End If


                                    If BOOLSTATSRE And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                        Try
                                            numBias = CalcREPercent(numMean, varNom, intQCDec)
                                        Catch ex As Exception

                                        End Try
                                    Else
                                        'get numbias from average of %Bias columns
                                        numBias = GetBiasFromDiffCol(idTR, varNom, intD + 1, 0, True)
                                    End If

                                    If BOOLSTATSRE And boolSTATSMEAN Then
                                        Try
                                            'record %RE
                                            int8 = int8 + 1
                                            .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                            '.Selection.TypeText(Text:=Format(numBias, strQCDec))
                                            '20180430 LEE:
                                            'must account for Endogenous Cmpds, NomConc = 0
                                            .Selection.TypeText(Text:=NomConcZero(varNom, numBias))

                                        Catch ex As Exception

                                        End Try
                                    End If

                                    If boolSTATSN Then
                                        Try
                                            'enter n
                                            int8 = int8 + 1
                                            .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                            .Selection.TypeText(CStr(int2))

                                        Catch ex As Exception

                                        End Try
                                    End If
                                End If

                            End If

                            'End If
                        Next

                        'bottom border this row
                        .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.MoveRight Unit:=Microsoft.Office.Interop.Word.wdunits.word.wdunits.wdCharacter, Count:=ctQCs, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                        'If boolQCREPORTACCVALUES Then
                        'Else
                        '    If intExp = 0 Then
                        '        Call DeleteRows(ctExp, wd)
                        '    End If
                        'End If

                        'Call DeleteTableRows(wd)

                        'remove unused rows
                        Call RemoveRows(wd, 1)

                        'autofit table
                        Call AutoFitTable(wd, False)

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


                    'go to end of table
                    .Selection.Tables.Item(1).Cell(1, 1).Select()
                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdColumn)

                    'enter table number
                    '***
                    'strA = rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")
                    If gNumMatrix = 1 Then
                        strA = strAnalC
                    Else
                        strA = strAnal 'strAnalC has '..Matrix', don't want to pass that here
                    End If
                    'No. Now just send strAnal
                    strA = strAnal
                    strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, False)
                    Call EnterTableNumber(wd, strTName, 5, strA, strTempInfo, intTableID, intGroup, idTR)
                    '***

                    'Call EnterTableNumber(wd, str1, 3)

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

                    ''''wdd.visible = True

                    '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                    '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)

                    'autofit window AGAIN
                    'autofit table
                    Call AutoFitTable(wd, BOOLINCLUDEDATE)

                    strM = "Finalizing " & strTName & "..."
                    strM1 = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    str1 = strM1

                    frmH.lblProgress.Text = strM1
                    frmH.Refresh()

                    Call SplitTable(wd, 4, intLeg, arrLegend, str1, False, ctLegend + 2, False, True, False, intTableID)
                    'Sub SplitTable(ByVal wd As Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, 
                    'ByVal arr As Object, ByVal strT As String, ByVal DoLegend As Boolean, ByVal intSplitRows As Short, ByVal boolSmallFont As Boolean)

                    'autofit window AGAIN
                    'autofit table
                    'autofit table
                    Call AutoFitTable(wd, BOOLINCLUDEDATE)

                    'move to line below table
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

next1:

            Next


end2:
        End With

    End Sub



    Sub MVSummaryFinalExtractQC_21(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal idTR As Int64)

        '20191017 LEE: Divert to MVAdHocQCStability_31

        Dim numNomConc As Decimal
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

        Dim int12 As Short
        Dim boolEnterDiff As Boolean

        Dim ctLegend As Short
        Dim fontsize
        Dim boolPro As Boolean

        Dim hi, lo
        Dim rows10() As DataRow
        Dim rows11() As DataRow
        Dim intRowsAnal As Short
        Dim arrFP(2, 20) 'FlagPercent array
        '1=max, 2=min
        Dim strFP As String
        Dim numMean As Decimal
        Dim numBias As Decimal
        Dim numSD As Decimal
        Dim tblZ As System.Data.DataTable
        Dim dvAn As System.Data.DataView
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

        Dim strDECISIONREASON As String
        Dim boolExFromAS As Boolean

        Dim numPrec As Single
        Dim numTheor As Single

        Dim v1, v2, vU
        Dim intRC As Short

        Dim varConc

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

        Dim boolOC As Boolean = False 'bool if eliminated

        Dim charFCID As String
        strF = "ID_TBLREPORTTABLE = " & idTR
        Dim rowsTR() As DataRow = tblReportTable.Select(strF)
        var1 = rowsTR(0).Item("CHARFCID")
        charFCID = NZ(var1, "NA")


        boolJustTable = False

        Cursor.Current = Cursors.WaitCursor

        '''''''''''wdd.visible = True

        fontsize = wd.ActiveDocument.Styles("Normal").Font.Size ' wd.Selection.Font.Size
        fonts = fontsize ' wd.Selection.Font.Size

        With wd

            intTableID = 21

            Dim strWRunId As String = GetWatsonColH(intTableID)

            dvDo = frmH.dgvReportTableConfiguration.DataSource
            strF = "id_tblconfigreporttables = " & intTableID
            intDo = FindRowDVNumByCol(intTableID, dvDo, "id_tblconfigreporttables")

            ''Get table name
            'var1 = dvDo(intDo).Item("Table")
            'strTName = NZ(var1, "[NONE]")

            ''get Temperature info
            'var1 = dvDo(intDo).Item("PERIODTEMP")
            'strTempInfo = NZ(var1, "[NONE]")

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
            '    strM = "Creating Summary of " & strTempInfo & " Final Extract Stability Table ...."
            '    frmH.lblProgress.Text = strM
            '    frmH.Refresh()
            '    MsgBox("Samples have not been assigned to this table.", MsgBoxStyle.Information, "Samples have not been assigned...")
            '    GoTo end2
            'End If

            strF = "IsIntStd = 'No'"
            strS = "INTORDER ASC, IsIntStd ASC, AnalyteDescription ASC"
            rows11 = tblAnalytesHome.Select(strF, strS)

            intRowsAnal = rows11.Length

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

                Dim strM1 As String
                If bool Then 'continue
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

                    intTCur = intTCur + 1

                    If rowsX.Length = 0 Then
                        strM = "Creating " & strTName & "...."
                        strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                        frmH.lblProgress.Text = strM

                        frmH.Refresh()
                        'MsgBox("Samples have not been assigned to this table.", MsgBoxStyle.Information, "Samples have not been assigned...")
                        'page setup according to configuration
                        str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                        'insert page break
                        ' wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                        Call InsertPageBreak(wd)
                        Call PageSetup(wd, str1) 'L=Landscape, P=Portrait
                        boolJustTable = True
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

                    ''setup tables
                    'var1 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteIndex")
                    'var2 = tbl4.Rows.Item(Count1 - 1).Item("MasterAssayID")
                    'var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")

                    'vAnalyteIndex = var1
                    'vMasterAssayID = var2
                    'vAnalyteID = tbl4.Rows.Item(Count1 - 1).Item("AnalyteID")

                    'strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                    'strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                    'strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                    'strF2 = strF2 & "ANALYTEINDEX = " & var1 & " AND "
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
                    For Count2 = 0 To intNumLevels - 1 'check for any null values
                        var3 = tblLevels.Rows.Item(Count2).Item("NOMCONC")
                        If IsDBNull(var3) Then
                            str1 = "The Nominal Concentration for some assigned samples for " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & " have not been configured."
                            str1 = str1 & ChrW(10) & "When this action is finished, please navigate to the Assigned Samples window and correct this problem."
                            If boolDisableWarnings Then
                            Else
                                MsgBox(str1, MsgBoxStyle.Information, "Nom Conc problem...")
                            End If
                            GoTo end1
                        End If
                        var3 = tblLevels.Rows.Item(Count2).Item("CHARHELPER1")
                        If IsDBNull(var3) Then
                            str1 = "The Term 1 designation for some assigned samples for " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & " have not been configured."
                            str1 = str1 & ChrW(10) & "When this action is finished, please navigate to the Assigned Samples window and correct this problem."
                            If boolDisableWarnings Then
                            Else
                                MsgBox(str1, MsgBoxStyle.Information, "Nom Conc problem...")
                            End If
                            GoTo end1
                        End If

                    Next

                    'find number of table rows to generate
                    intRowsX = 0

                    boolOutHeadE = False
                    boolOutHeadI = False
                    boolDeleteRows = False

                    Dim intRowsXTot As Short = 0



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
                    intTblRows = intTblRows + (1 * intNumRuns) 'for one blank rows after each run set
                    'intTblRows = intTblRows + (5 * intNumRuns) 'for Mean/Bias/n section for each run set

                    'Increment for Statistics Sections

                    Dim intCSN As Short
                    intCSN = countNumStatsRows()
                    intTblRows = intTblRows + intCSN

                    If intCSN > 0 Then
                    Else
                        intTblRows = intTblRows - 1 'subtract an unneeded blank row
                    End If

                    'intTblRows = intTblRows + (2 * intNumRuns) - 2 'for two blank rows after each Mean/Bias/n set, except last set
                    'intTblRows = intTblRows + (2 * intNumRuns) - 3 'for two blank rows after each Mean/Bias/n set, except last set
                    'NDL commented this out: it was preventing rows from displaying in this display (when presenting a single set).
                    If boolQCREPORTACCVALUES Then

                    Else

                        'For Count2 = 1 To intNumRuns 'must loop because this table may have several sections
                        'intTblRows = intTblRows + (3 * intNumRuns) 'for stats headings

                        intTblRows = intTblRows + 2 'for stats headings
                        ctExp = ctExp + 2 'for stats headings

                        'Increment for Statistics Sections
                        'intTblRows = intTblRows + (countNumStatsRows() * intNumRuns)
                        intTblRows = intTblRows + (intCSN)
                        'ctExp = ctExp + (countNumStatsRows() * intNumRuns)
                        ctExp = ctExp + (intCSN)

                        If intCSN > 0 Then
                            intTblRows = intTblRows + 1 'one more blank row
                            ctExp = ctExp + 1 'one more blank row
                        End If

                        'intTblRows = intTblRows + (1 * intNumRuns) - 1 'for a blank row after each Mean/Bias/n set, except last set
                        'ctExp = ctExp + (1 * intNumRuns) - 1
                        'Next
                    End If
                    wrdSelection = wd.Selection()

                    Dim intCols As Short
                    If boolSTATSDIFFCOL Then
                        intCols = (intNumLevels * 2) + 1
                    Else
                        intCols = intNumLevels + 1
                    End If


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

                        .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=2, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.MoveRight Unit:=Microsoft.Office.Interop.Word.wdunits.word.wdunits.wdCharacter, Count:=ctQCs, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                        'Enter QC ID row titles
                        .Selection.Tables.Item(1).Cell(2, 2).Select()
                        For Count2 = 0 To intNumLevels - 1
                            'var1 = arrBCQCs(3, Count2)
                            var2 = tblLevels.Rows.Item(Count2).Item("CHARHELPER1")
                            var3 = ReturnStdQC(var2.ToString)
                            'var3 = var2
                            .Selection.TypeText(Text:=var3)
                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                            If boolSTATSDIFFCOL Then
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                            End If

                        Next


                        'Enter nom. conc. row titles
                        .Selection.Tables.Item(1).Cell(3, 2).Select()
                        For Count2 = 0 To intNumLevels - 1
                            'var1 = arrBCQCs(3, Count2)
                            If boolLUseSigFigs Then
                                var1 = CStr(SigFigOrDecString(tblLevels.Rows.Item(Count2).Item("NOMCONC"), LSigFig, False))
                            Else
                                var1 = CStr(Format(tblLevels.Rows.Item(Count2).Item("NOMCONC"), GetRegrDecStr(LSigFig)))
                            End If

                            'var3 = var1 & ChrW(160) & strConcUnits

                            If LboolNomConcParen Then
                                var3 = "(" & var1 & ChrW(160) & strConcUnits & ")"
                            Else
                                var3 = var1 & ChrW(160) & strConcUnits
                            End If
                            .Selection.TypeText(Text:=var3)

                            '******determine if the level is a diln level
                            Dim strE As String
                            'var3 = var1 ' & ChrW(10) & var1 & " " & strConcUnits
                            'strE = ChrW(10) & var1 & " " & strConcUnits
                            '.Selection.TypeText(Text:=var3)

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
                                        'arrLegend(2, intLeg) = "Dilution QCs undiluted concentration " & Format(RoundToDecimalRAFZ(CDbl(Val(var1)), LSigFig), GetRegrDecStr(LSigFig)) & " " & strConcUnits & "; a 1:" & var3 & " dilution with blank matrix was performed prior to extraction and analysis."
                                        arrLegend(2, intLeg) = "Dilution QCs undiluted concentration " & Format(RoundToDecimalRAFZ(CDbl(Val(var1)), LSigFig), GetRegrDecStr(LSigFig)) & " " & strConcUnits & "; " & strAN & " " & var3 & "-fold dilution with blank matrix was performed prior to extraction and analysis."
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
                            .Selection.TypeText(strE)

                            '******


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

                        'enter Watson Run ID column header
                        .Selection.Tables.Item(1).Cell(3, 1).Select()

                        If BOOLINCLUDEDATE Then
                            .Selection.Tables.Item(1).Cell(2, 1).Select()
                            .Selection.TypeText(strWRunId)
                            .Selection.Tables.Item(1).Cell(3, 1).Select()
                            '.Selection.TypeText("(Analysis Date)")
                            '20180420 LEE:
                            .Selection.TypeText("(" & GetAnalysisDateLabel(intTableID) & ")")
                        Else
                            If boolSTATSDIFFCOL Then
                                int1 = InStr(strWRunId, " ", CompareMethod.Text)
                                If int1 = 0 Then
                                    .Selection.Tables.Item(1).Cell(3, 1).Select()
                                    .Selection.TypeText(strWRunId)
                                Else
                                    str1 = Mid(strWRunId, 1, int1 - 1)
                                    str2 = Mid(strWRunId, int1 + 1, Len(strWRunId))
                                    .Selection.Tables.Item(1).Cell(2, 1).Select()
                                    .Selection.TypeText(str1)
                                    .Selection.Tables.Item(1).Cell(3, 1).Select()
                                    .Selection.TypeText(str2)
                                End If
                            Else
                                str1 = Replace(strWRunId, " ", ChrW(160), 1, -1, CompareMethod.Text)
                                '.Selection.TypeText("Watson" & ChrW(160) & "Run" & ChrW(160) & "ID")
                                .Selection.TypeText(str1)
                            End If

                        End If

                        'begin entering data'
                        intStart = 5
                        int1 = 5 'row position counter
                        Dim boolExit As Boolean = False

                        'now do rows2E


                        For Count2 = 0 To intNumRuns - 1

                            '.Selection.Tables.Item(1).Cell(intStart, 1).Select()
                            .Selection.Tables.Item(1).Cell(int1, 1).Select()
                            'enter runid
                            var10 = tblNumRuns.Rows.Item(Count2).Item("RUNID")


                            frmH.lblProgress.Text = strM1 & ChrW(10) & "Processing Run ID " & var10
                            frmH.Refresh()

                            .Selection.TypeText(CStr(var10))
                            If BOOLINCLUDEDATE Then
                                'If Count2 = 0 Then
                                '    .Selection.Tables.Item(1).Cell(intStart + 1, 1).Select()
                                'Else
                                '    .Selection.Tables.Item(1).Cell(int1 + 1, 1).Select()
                                'End If
                                .Selection.Tables.Item(1).Cell(int1 + 1, 1).Select()
                                str1 = GetDateFromRunID(NZ(var10, 0), LDateFormat, intGroup, idTR)
                                .Selection.TypeText("(" & str1 & ")")
                                'If Count2 = 0 Then
                                '    .Selection.Tables.Item(1).Cell(intStart, 1).Select()
                                'Else
                                '    .Selection.Tables.Item(1).Cell(int1, 1).Select()
                                'End If
                                .Selection.Tables.Item(1).Cell(int1, 1).Select()

                            End If

                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)

                            'must first determine if there are any outliers
                            intExp = 0


                            '20160816 LEE: intRowsX needs to be determined for each dataset
                            '****
                            dv2.RowFilter = ""
                            'don't know why, but must make a long filter here or
                            'both analytes get returned in dv2.rowfilter

                            intRowsX = 0

                            int12 = -1
                            For Count3 = 0 To tblLevels.Rows.Count - 1
                                int12 = int12 + 1
                                varNom = tblLevels.Rows.Item(int12).Item("NOMCONC")
                                strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom
                                dv2.RowFilter = strF
                                int2 = dv2.Count
                                If int2 > intRowsX Then
                                    intRowsX = int2
                                End If
                            Next

                            '****

                            For Count3 = 0 To intNumLevels - 1

                                varNom = tblLevels.Rows.Item(Count3).Item("NOMCONC")
                                'start evaluating data
                                dv2.RowFilter = ""
                                'don't know why, but must make a long filter here or
                                'both analytes get returned in dv2.rowfilter
                                strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom
                                dv2.RowFilter = strF
                                int2 = dv2.Count

                                'create rows1 from tbl1 which will contain data
                                strF = ""

                                Erase rows1
                                Dim tbl2B As System.Data.DataTable = dv2.ToTable
                                rows2 = tbl2B.Select("RUNID > 0")
                                int3 = rows2.Length

                                'evaluate data
                                For Count4 = 0 To intRowsX - 1 'int3 - 1
                                    var1 = NZ(rows2(Count4).Item("ELIMINATEDFLAG"), "N")
                                    var2 = NZ(rows2(Count4).Item("BOOLEXCLSAMPLE"), 0)
                                    If StrComp(var1, "Y", CompareMethod.Text) = 0 Or (gAllowExclSamples And LAllowExclSamples And var2 = -1) Then
                                        boolExFromAS = False
                                    Else
                                        If gAllowExclSamples And LAllowExclSamples Then
                                            If var2 = -1 Then
                                                var1 = "Y"
                                                boolExFromAS = True
                                            Else
                                                'var1 = "N"
                                                'don't assign "N", Watson may override
                                            End If
                                        End If
                                        intExp = intExp + 1
                                        boolExit = True
                                        Exit For
                                    End If
                                Next

                                If boolExit Then
                                    Exit For
                                End If

                            Next


                            int12 = -1
                            'start filling in data by columns
                            'intRowsX = 0

                            'int1 = 0

                            For Count3 = 0 To (intNumLevels * int11) - 1 Step int11
                                int12 = int12 + 1
                                varNom = tblLevels.Rows.Item(int12).Item("NOMCONC")


                                'determine hi and lo (nom*flagpercent)
                                strF = "CONCENTRATION = '" & varNom & "'"
                                'rows10 = tblBCQCs.Select(strF)

                                'determine hi and lo (nom*flagpercent)
                                'strF = "CONCENTRATION = " & varNom & " AND ANALYTEID = " & vAnalyteID & " AND MASTERASSAYID = " & vMasterAssayID & " AND ANALYTEINDEX = " & vAnalyteIndex & " AND CONCENTRATION = " & varNom & " AND RUNID = " & var10
                                'if Conc < 1, then the query return 0 records
                                'must do something different
                                var1 = GetANALYTEFLAGPERCENT(varNom, var10, vAnalyteID)

                                'var1 = CDec(NZ(rows10(0).Item("FLAGPERCENT"), 15))
                                arrFP(1, int12) = var1
                                arrFP(2, int12) = var1
                                Call SetHighAndLowCriteria(varNom, var1, var1, hi, lo)

                                'start entering data
                                dv2.RowFilter = ""
                                'don't know why, but must make a long filter here or
                                'both analytes get returned in dv2.rowfilter
                                strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom
                                dv2.RowFilter = strF
                                int2 = dv2.Count
                                'nI = int2

                                'create rows1 from tbl1 which will contain data
                                strF = ""
                                Dim tbl2s As System.Data.DataTable = dv2.ToTable
                                rows2 = tbl2s.Select("RUNID > -1")
                                int3 = rows2.Length

                                If int3 > 0 Then
                                    'do hi/lo again
                                    strF = "CONCENTRATION = '" & varNom & "'"
                                    'rows10 = tblBCQCs.Select(strF)

                                    ''determine hi and lo (nom*flagpercent)
                                    'strF = "CONCENTRATION = " & varNom & " AND ANALYTEID = " & vAnalyteID & " AND MASTERASSAYID = " & vMasterAssayID & " AND ANALYTEINDEX = " & vAnalyteIndex & " AND CONCENTRATION = " & varNom & " AND RUNID = " & var10
                                    'rows10 = tblQCRunIDs.Select(strF, "LEVELNUMBER ASC")

                                    vU = rows2(0).Item("BOOLUSEGUWUACCCRIT")
                                    If gAllowGuWuAccCrit And LAllowGuWuAccCrit And vU = -1 Then
                                        var1 = CDec(NZ(rows2(0).Item("NUMMAXACCCRIT"), 0))
                                        var2 = CDec(NZ(rows2(0).Item("NUMMINACCCRIT"), 0))
                                        arrFP(1, int12) = var1
                                        arrFP(2, int12) = var2

                                        Call SetHighAndLowCriteria(varNom, var1, var2, hi, lo)
                                        v1 = var1
                                        v2 = var2

                                    Else
                                        'if Conc < 1, then the query return 0 records

                                        'must do something different
                                        var1 = GetANALYTEFLAGPERCENT(varNom, var10, vAnalyteID)

                                        'var1 = CDec(NZ(rows10(0).Item("FLAGPERCENT"), 15))
                                        arrFP(1, int12) = var1
                                        arrFP(2, int12) = var1

                                        Call SetHighAndLowCriteria(varNom, var1, var1, hi, lo)
                                        v1 = var1
                                        v2 = var1

                                    End If
                                Else
                                    var1 = var1 'debug
                                End If

                                'enter data
                                For Count4 = 0 To intRowsX - 1 'int3 - 1

                                    boolOC = False

                                    If Count4 > int3 - 1 Then
                                    Else
                                        boolEnterDiff = False
                                        If Count4 = 0 Then
                                            'enter charhelper2
                                            str1 = NZ(dv2(0).Item("CHARHELPER2"), "")
                                            '.Selection.Tables.Item(1).Cell(intStart + Count4 + 1, 1).Select()
                                            .Selection.Tables.Item(1).Cell(int1 + Count4 + 1, 1).Select()
                                            .Selection.TypeText(str1)
                                        End If

                                    End If

                                    '.Selection.Tables.Item(1).Cell(intStart + Count4, Count3 + 2).Select()
                                    .Selection.Tables.Item(1).Cell(int1 + Count4, Count3 + 2).Select()
                                    If Count4 > int3 - 1 Then
                                        If boolQCNA Then
                                            str1 = "NA"
                                        Else
                                            str1 = ""
                                        End If

                                        .Selection.TypeText(str1)
                                        boolEnterDiff = False
                                    Else
                                        var1 = rows2(Count4).Item("CONCENTRATION")
                                        varConc = var1
                                        var1 = NZ(var1, 0)
                                        numDF = rows2(Count4).Item("ALIQUOTFACTOR")
                                        var1 = var1 / numDF
                                        If boolLUseSigFigs Then
                                            var2 = SigFigOrDec(var1, LSigFig, False)
                                        Else
                                            var2 = RoundToDecimalRAFZ(var1, LSigFig)
                                        End If

                                        var1 = NZ(rows2(Count4).Item("ELIMINATEDFLAG"), "N")
                                        var3 = NZ(rows2(Count4).Item("BOOLEXCLSAMPLE"), 0)
                                        vU = NZ(rows2(Count4).Item("BOOLUSEGUWUACCCRIT"), 0)

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

                                            boolEnterDiff = True 'False
                                            intExp = intExp + 1
                                            intLeg = intLeg + 1
                                            strA = ChrW(intLeg + intLegStart)

                                            '20160305 LEE:
                                            'Added DECISIONREASON code
                                            Dim var6
                                            'Remember, tblAssignedSamples does not have DECISIONREASON
                                            var6 = "No Value: " & GetDECISIONREASONValue(boolExFromAS, vAnalyteID, rows2(Count4))
                                            ''Set Legend String
                                            str1 = GetLegendStringExcluded(v1, v2, vU, var6, intTableID, True, "")
                                            'Add to Legend Array
                                            ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                            If boolRedBoldFont Then
                                                .Selection.Font.Bold = True
                                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                            End If
                                            .Selection.TypeText("NV")
                                            Call typeInSuperscriptFontSize12WithSpace(wd, strA)

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

                                            boolEnterDiff = True 'False
                                            intExp = intExp + 1
                                            intLeg = intLeg + 1
                                            strA = ChrW(intLeg + intLegStart)

                                            '20160305 LEE:
                                            'Added DECISIONREASON code
                                            Dim var6
                                            'Remember, tblAssignedSamples does not have DECISIONREASON
                                            var6 = GetDECISIONREASONValue(boolExFromAS, vAnalyteID, rows2(Count4))
                                            ''Set Legend String
                                            str1 = GetLegendStringExcluded(arrFP(1, int12), arrFP(2, int12), vU, var6, intTableID, True, "")
                                            'Add to Legend Array
                                            ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                            If boolRedBoldFont Then
                                                .Selection.Font.Bold = True
                                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                            End If

                                            If boolLUseSigFigs Then
                                                .Selection.TypeText(CStr(DisplayNum(var2, LSigFig, False)))
                                            Else
                                                .Selection.TypeText(Format(var2, GetRegrDecStr(LSigFig)))
                                            End If
                                            Call typeInSuperscriptFontSize12WithSpace(wd, strA)

                                        Else

                                            boolEnterDiff = True
                                            'determine if value is outside acceptance criteria
                                            If OutsideAccCrit(var2, varNom, v1, v2, NZ(vU, 0)) Then
                                                'If var2 > hi Or var2 < lo Then 'flag
                                                intLeg = intLeg + 1
                                                strA = ChrW(intLeg + intLegStart)

                                                'Set Legend String
                                                str1 = GetLegendStringIncluded(v1, v2, vU)
                                                'Add to Legend Array
                                                ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                                If boolRedBoldFont Then
                                                    .Selection.Font.Bold = True
                                                    .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                End If

                                                If boolLUseSigFigs Then
                                                    .Selection.TypeText(Text:=CStr(DisplayNum(var2, LSigFig, False)))
                                                Else
                                                    .Selection.TypeText(Text:=Format(var2, GetRegrDecStr(LSigFig)))
                                                End If

                                                Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                            Else
                                                If boolLUseSigFigs Then
                                                    .Selection.TypeText(Text:=CStr(DisplayNum(var2, LSigFig, False)))
                                                Else
                                                    .Selection.TypeText(Text:=Format(var2, GetRegrDecStr(LSigFig)))
                                                End If

                                                boolEnterDiff = True
                                            End If
                                        End If

                                    End If

                                    If boolSTATSDIFFCOL Then
                                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                        If boolEnterDiff Then
                                            'var3 = Format(((var2 / varNom) - 1) * 100, strQCDec)
                                            'var3 = Format(RoundToDecimal(((var2 / varNom) - 1) * 100, intQCDec), strQCDec)


                                            If boolTHEORETICAL Then
                                                var3 = CalcREPercent(var2, varNom, intQCDec)
                                                numTheor = 100 + CDec(var3)

                                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", numTheor, CSng(var10), Count1, strDo, v1, v2, boolOC)

                                            Else
                                                var3 = Format(RoundToDecimal(CalcREPercent(var2, varNom, intQCDec), intQCDec), strQCDec)

                                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", var3, CSng(var10), Count1, strDo, v1, v2, boolOC)

                                            End If

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

                                Next Count4 'intRowsX Data

                                'remove Stats from here and put them outside Count3



                            Next Count3 'numLevels

                            'MsgBox("int12 = " & int12)

                            'increase row position counter
                            If Count2 = intNumRuns - 1 Then
                                int1 = int1 + intRowsX + 1
                                'intStart = int1 + int8 + 1
                            Else
                                int1 = int1 + intRowsX + 1
                                'intStart = int1 + int8 + 1
                            End If

                            '''''''wdd.visible = True

                        Next Count2 'numRuns

                        '*****start stats

                        '.Selection.Tables.Item(1).Cell(intStart + Count4 + 1, 1).Select()
                        '.Selection.Tables.Item(1).Cell(int1 + Count4 + 1, 1).Select()
                        .Selection.Tables.Item(1).Cell(int1, 1).Select()
                        int1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)
                        '.Selection.Tables.Item(1).Cell(int1 + 1, 2).Select()
                        'begin doing stats
                        If boolQCREPORTACCVALUES Then
                        Else
                            If intExp = 0 Then
                            Else
                                '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                '.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                                'enter some blank spaces to fool PageBreak function
                                '.selection.typetext(Text:="  ")
                                If boolOutHeadE Then
                                Else
                                    .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                    .Selection.TypeText(Text:="Summary Statistics Excluding Outlier Values")
                                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                    Try
                                        .Selection.Cells.Merge()
                                    Catch ex As Exception

                                    End Try
                                    '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                    With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                        .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                    End With
                                    boolOutHeadE = True
                                End If


                                .Selection.Tables.Item(1).Cell(int1 + 1, 1).Select()
                            End If
                        End If

                        int1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)

                        'now enter Mean/Bias/n
                        'If Count3 = 0 Then
                        int8 = 0
                        Call typeStatsLabels(wd, int8, int1 - 1, 1, False)

                        If boolQCREPORTACCVALUES Then
                        Else
                            If intExp = 0 Then
                            Else
                                int8 = int8 + 1
                                .Selection.Tables.Item(1).Cell(int1 + int8, 1).Select()
                                '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                '.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                                'enter some blank spaces to fool PageBreak function
                                '.selection.typetext(Text:="  ")
                                If boolOutHeadI Then
                                Else
                                    .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                    .Selection.TypeText(Text:="Summary Statistics Including Outlier Values")
                                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                    Try
                                        .Selection.Cells.Merge()
                                        With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                            .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                        End With
                                    Catch ex As Exception

                                    End Try
                                    '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                    boolOutHeadI = True
                                End If

                                Call typeStatsLabels(wd, int8, int1, 1, False)

                            End If
                        End If
                        'End If

                        int12 = -1
                        For Count3 = 0 To (intNumLevels * int11) - 1 Step int11

                            int12 = int12 + 1
                            varNom = tblLevels.Rows.Item(int12).Item("NOMCONC")

                            dv2.RowFilter = ""
                            'don't know why, but must make a long filter here or
                            'both analytes get returned in dv2.rowfilter
                            strF = strF2 & " AND NOMCONC = " & varNom
                            dv2.RowFilter = strF
                            int2 = dv2.Count
                            nI = int2

                            'create rows1 from tbl1 which will contain data
                            strF = ""
                            Dim tbl2t As System.Data.DataTable = dv2.ToTable
                            rows2 = tbl2t.Select("RUNID > -1")


                            'now do rows2E
                            strF = ""
                            Erase rows2E
                            'rows2E = tbl1.Select(strF)
                            If gAllowExclSamples And LAllowExclSamples Then
                                rows2E = tbl2t.Select("ELIMINATEDFLAG = 'N' AND BOOLEXCLSAMPLE = 0")
                            Else
                                rows2E = tbl2t.Select("ELIMINATEDFLAG = 'N'")
                            End If
                            nE = rows2E.Length


                            int8 = 0
                            .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 2).Select()
                            v1 = arrFP(1, int12)
                            v2 = arrFP(2, int12)

                            Dim boolMean As Boolean = True
                            Try
                                var1 = MeanDR(rows2E, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                If boolLUseSigFigs Then
                                    numMean = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                Else
                                    numMean = RoundToDecimalRAFZ(var1, LSigFig)
                                End If

                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Mean", numMean, CSng(var10), Count1, strDo, 0, 0, False)
                            Catch ex As Exception

                            End Try
                            If boolSTATSMEAN Then
                                Try
                                    'enter Mean
                                    int8 = int8 + 1

                                    If nE = 0 Then
                                        boolMean = False
                                        .Selection.TypeText("NA")
                                    Else
                                        If (OutsideAccCrit(numMean, varNom, v1, v2, NZ(vU, 0))) And boolFootNoteQCMean Then 'flag
                                            'OutsideAccCrit(varConc, varNom, v1, v2, NZ(vU, 0))
                                            intLeg = intLeg + 1
                                            strA = ChrW(intLeg + intLegStart)
                                            'Set Legend String
                                            str1 = GetLegendStringIncluded(v1, v2, vU)
                                            'Add to Legend Array
                                            ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                            If boolRedBoldFont Then
                                                .Selection.Font.Bold = True
                                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                            End If

                                            If boolLUseSigFigs Then
                                                .Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                            Else
                                                .Selection.TypeText(Text:=Format(numMean, GetRegrDecStr(LSigFig)))
                                            End If

                                            Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                            boolEnterDiff = True
                                        Else
                                            If boolLUseSigFigs Then
                                                .Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                            Else
                                                .Selection.TypeText(Text:=Format(numMean, GetRegrDecStr(LSigFig)))
                                            End If

                                            boolEnterDiff = True
                                        End If
                                    End If

                                    '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                    .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 2).Select()

                                Catch ex As Exception

                                End Try
                            End If

                            Try
                                var1 = StdDevDR(rows2E, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                If boolLUseSigFigs Then
                                    numSD = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                Else
                                    numSD = RoundToDecimalRAFZ(var1, LSigFig)
                                End If


                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "SD", numSD, CSng(var10), Count1, strDo, 0, 0, False)
                            Catch ex As Exception

                            End Try
                            If boolSTATSSD Then
                                Try
                                    'enter SD
                                    int8 = int8 + 1

                                    If nE < gSDMax Then
                                        .Selection.TypeText("NA")
                                    Else
                                        var1 = StdDevDR(rows2E, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                        If boolLUseSigFigs Then
                                            .Selection.TypeText(CStr(numSD))
                                        Else
                                            .Selection.TypeText(Format(numSD, GetRegrDecStr(LSigFig)))
                                        End If

                                    End If
                                    '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                    .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 2).Select()
                                Catch ex As Exception

                                End Try
                            End If


                            Try
                                If nE < gSDMax Then
                                Else
                                    numPrec = CalcCVPercent(numSD, numMean, intQCDec)
                                    Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Precision", numPrec, CSng(var10), Count1, strDo, 0, 0, False)
                                End If

                            Catch ex As Exception

                            End Try
                            If boolSTATSCV Then
                                Try
                                    'enter %CV
                                    int8 = int8 + 1
                                    If nE < gSDMax Then
                                        .Selection.TypeText("NA")
                                    Else
                                        .Selection.TypeText(Format(numPrec, strQCDec))
                                    End If

                                    '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                    .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 2).Select()
                                Catch ex As Exception

                                End Try
                            End If

                            If boolSTATSBIAS And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                Try
                                    numBias = CalcREPercent(numMean, varNom, intQCDec)
                                    If nE = 0 Then
                                    Else
                                        Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", numBias, CSng(var10), Count1, strDo, 0, 0, False)
                                    End If

                                Catch ex As Exception

                                End Try
                            End If
                            If boolSTATSBIAS And boolSTATSMEAN Then
                                Try
                                    'enter %Bias
                                    int8 = int8 + 1
                                    If boolMean Then
                                        .Selection.TypeText(Format(numBias, strQCDec))
                                    Else
                                        .Selection.TypeText("NA")
                                    End If

                                    .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 2).Select()
                                Catch ex As Exception

                                End Try
                            End If

                            If boolTHEORETICAL And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                Try
                                    numTheor = CalcREPercent(numMean, varNom, intQCDec)
                                    numTheor = 100 + CDec(numTheor)
                                    If nE = 0 Then
                                    Else
                                        Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", numTheor, CSng(var10), Count1, strDo, 0, 0, False)
                                    End If

                                Catch ex As Exception

                                End Try
                            End If

                            If boolTHEORETICAL And boolSTATSMEAN Then
                                Try
                                    'enter %theoretical
                                    int8 = int8 + 1

                                    If boolMean Then
                                        .Selection.TypeText(Format(numTheor, strQCDec))
                                    Else
                                        .Selection.TypeText("NA")
                                    End If
                                    .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 2).Select()
                                Catch ex As Exception

                                End Try

                            End If

                            If boolSTATSDIFF And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                Try
                                    numBias = CalcREPercent(numMean, varNom, intQCDec)
                                    If nE = 0 Then
                                    Else
                                        Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", numBias, CSng(var10), Count1, strDo, 0, 0, False)
                                    End If

                                Catch ex As Exception

                                End Try
                            End If

                            If boolSTATSDIFF And boolSTATSMEAN Then
                                Try
                                    'enter %Bias
                                    int8 = int8 + 1
                                    If boolMean Then
                                        .Selection.TypeText(Format(numBias, strQCDec))
                                    Else
                                        .Selection.TypeText("NA")
                                    End If

                                    '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                    .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 2).Select()
                                Catch ex As Exception

                                End Try
                            End If

                            If BOOLSTATSRE And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                Try
                                    numBias = CalcREPercent(numMean, varNom, intQCDec)
                                    If nE = 0 Then
                                    Else
                                        Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", numBias, CSng(var10), Count1, strDo, 0, 0, False)
                                    End If

                                Catch ex As Exception

                                End Try
                            End If

                            If BOOLSTATSRE And boolSTATSMEAN Then
                                Try
                                    'enter %RE
                                    int8 = int8 + 1

                                    If boolMean Then
                                        .Selection.TypeText(Format(numBias, strQCDec))
                                    Else
                                        .Selection.TypeText("NA")
                                    End If

                                    '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                    .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 2).Select()
                                Catch ex As Exception

                                End Try

                            End If


                            Try
                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "n", nE, CSng(var10), Count1, strDo, 0, 0, False)
                            Catch ex As Exception

                            End Try
                            If boolSTATSN Then
                                Try
                                    'enter n
                                    int8 = int8 + 1
                                    '.Selection.TypeText(CStr(int2))
                                    .Selection.TypeText(CStr(nE))
                                    '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                    '.Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 2).Select()


                                Catch ex As Exception

                                End Try
                            End If

                            If boolQCREPORTACCVALUES Then
                            Else
                                If intExp = 0 Then
                                    If boolDeleteRows Then
                                    Else
                                        Call DeleteRows(ctExp / intNumRuns, wd)
                                        boolDeleteRows = True
                                    End If

                                Else
                                    int8 = int8 + 2
                                    .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 2).Select()

                                    '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                    '.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                                    ''enter some blank spaces to fool PageBreak function
                                    ''.selection.typetext(Text:="  ")
                                    '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                    '.Selection.TypeText(Text:="Summary Statistics Including Outlier Values")
                                    '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                    ' Try
                                    ''.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                    'With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                    '    .LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleSingle
                                    'End With

                                    boolMean = True


                                    If boolSTATSMEAN Then
                                        Try
                                            'enter Mean
                                            int8 = int8 + 1
                                            var1 = MeanDR(rows2, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                            If boolLUseSigFigs Then
                                                numMean = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                            Else
                                                numMean = RoundToDecimalRAFZ(var1, LSigFig)
                                            End If

                                            If nI = 0 Then
                                                boolMean = False
                                                .Selection.TypeText("NA")
                                            Else
                                                'determine if value is outside acceptance criteria
                                                'If (numMean > hi Or numMean < lo) And boolFootNoteQCMean Then 'flag
                                                If (OutsideAccCrit(numMean, varNom, v1, v2, NZ(vU, 0))) And boolFootNoteQCMean Then 'flag
                                                    intLeg = intLeg + 1
                                                    strA = ChrW(intLeg + intLegStart)

                                                    'Set Legend String
                                                    str1 = GetLegendStringIncluded(v1, v2, vU)
                                                    'Add to Legend Array
                                                    ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                                    If boolRedBoldFont Then
                                                        .Selection.Font.Bold = True
                                                        .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                    End If

                                                    If boolLUseSigFigs Then
                                                        .Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                                    Else
                                                        .Selection.TypeText(Text:=Format(numMean, GetRegrDecStr(LSigFig)))
                                                    End If

                                                    Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                                    boolEnterDiff = True
                                                Else
                                                    If boolLUseSigFigs Then
                                                        .Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                                    Else
                                                        .Selection.TypeText(Text:=Format(numMean, GetRegrDecStr(LSigFig)))
                                                    End If
                                                    boolEnterDiff = True
                                                End If
                                            End If

                                            '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                            .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 2).Select()

                                        Catch ex As Exception

                                        End Try
                                    End If
                                    If boolSTATSSD Then
                                        Try
                                            'enter SD
                                            int8 = int8 + 1

                                            If nI < gSDMax Then
                                                .Selection.TypeText("NA")
                                            Else
                                                var1 = StdDevDR(rows2, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                                If boolLUseSigFigs Then
                                                    numSD = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                                Else
                                                    numSD = RoundToDecimalRAFZ(var1, LSigFig)
                                                End If

                                                .Selection.TypeText(CStr(numSD))

                                            End If

                                            '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                            .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 2).Select()
                                        Catch ex As Exception

                                        End Try
                                    End If

                                    If boolSTATSCV Then
                                        Try
                                            'enter %CV
                                            int8 = int8 + 1
                                            If nE < gSDMax Then
                                                .Selection.TypeText("NA")
                                            Else
                                                var1 = Format(numSD / numMean * 100, strQCDec)
                                                numPrec = CalcCVPercent(numSD, numMean, intQCDec)
                                                .Selection.TypeText(Format(numPrec, strQCDec))
                                            End If

                                            '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                            .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 2).Select()
                                        Catch ex As Exception

                                        End Try
                                    End If
                                    If boolSTATSBIAS And boolSTATSMEAN Then
                                        Try
                                            'enter %Bias
                                            int8 = int8 + 1
                                            'var1 = (((numMean / varNom) - 1) * 100)
                                            'var1 = Format(var1, strQCDec)
                                            '.Selection.TypeText(CStr(var1))
                                            numBias = CalcREPercent(numMean, varNom, intQCDec)
                                            If boolMean Then
                                                .Selection.TypeText(Format(numBias, strQCDec))
                                            Else
                                                .Selection.TypeText("NA")
                                            End If

                                            '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                            .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 2).Select()
                                        Catch ex As Exception

                                        End Try
                                    End If

                                    If boolTHEORETICAL And boolSTATSMEAN Then
                                        Try
                                            'enter %theoretical
                                            int8 = int8 + 1
                                            'var1 = (((numMean / varNom) - 1) * 100)
                                            'var1 = Format(var1, strQCDec)
                                            'var1 = Format(100 + var1, strQCDec)
                                            '.Selection.TypeText(CStr(var1))

                                            numTheor = CalcREPercent(numMean, varNom, intQCDec)
                                            numTheor = 100 + CDec(numTheor)
                                            If boolMean Then
                                                .Selection.TypeText(Format(numTheor, strQCDec))
                                            Else
                                                .Selection.TypeText("NA")
                                            End If

                                            .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 2).Select()
                                        Catch ex As Exception

                                        End Try

                                    End If

                                    If boolSTATSDIFF And boolSTATSMEAN Then
                                        Try
                                            'enter %Bias
                                            int8 = int8 + 1
                                            'var1 = (((numMean / varNom) - 1) * 100)
                                            'var1 = Format(var1, strQCDec)
                                            '.Selection.TypeText(CStr(var1))

                                            numBias = CalcREPercent(numMean, varNom, intQCDec)
                                            If boolMean Then
                                                .Selection.TypeText(Format(numBias, strQCDec))
                                            Else
                                                .Selection.TypeText("NA")
                                            End If

                                            '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                            .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 2).Select()
                                        Catch ex As Exception

                                        End Try
                                    End If

                                    If BOOLSTATSRE And boolSTATSMEAN Then
                                        Try
                                            'enter %RE
                                            int8 = int8 + 1
                                            'var1 = (((numMean / varNom) - 1) * 100)
                                            'var1 = Format(var1, strQCDec)
                                            '.Selection.TypeText(CStr(var1))

                                            numBias = CalcREPercent(numMean, varNom, intQCDec)
                                            If boolMean Then
                                                .Selection.TypeText(Format(numBias, strQCDec))
                                            Else
                                                .Selection.TypeText("NA")
                                            End If

                                            '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                            .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 2).Select()
                                        Catch ex As Exception

                                        End Try
                                    End If

                                    If boolSTATSN Then
                                        Try
                                            'enter n
                                            int8 = int8 + 1
                                            .Selection.TypeText(CStr(nI))
                                            '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                            '.Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 2).Select()

                                        Catch ex As Exception

                                        End Try
                                    End If
                                End If

                            End If

                        Next



                        '*****end stats


                        'bottom border this row
                        .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.MoveRight Unit:=Microsoft.Office.Interop.Word.wdunits.word.wdunits.wdCharacter, Count:=ctQCs, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                        'autofit window
                        .Selection.Tables.Item(1).Select()
                        'autofit table
                        Call AutoFitTable(wd, False)

                        'go back and merge line 1
                        .Selection.Tables.Item(1).Cell(1, 2).Select()
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.Cells.Merge()
                        If intNumLevels < 2 Then
                        Else
                            Try
                                .Selection.Cells.Merge()
                            Catch ex As Exception

                            End Try
                        End If
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


                    'go to end of table
                    .Selection.Tables.Item(1).Cell(1, 1).Select()
                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdColumn)

                    'enter table number
                    var1 = strTempInfo
                    'replace numeric with verbose in strTempInfo
                    'look for a sequence of characters that is numeric
                    var3 = ""
                    var4 = ""
                    Dim bool1 As Boolean
                    Dim bool2 As Boolean
                    bool1 = False 'Start
                    bool2 = False 'End
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

                    str1 = str2 & " Final Extract Stability: Summary of " & rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION") & " Interpolated QC Standard Concentrations."

                    '***
                    strA = rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")
                    strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, False)
                    Call EnterTableNumber(wd, strTName, 3, strA, strTempInfo, intTableID, intGroup, idTR)
                    '***
                    'Call EnterTableNumber(wd, str1, 3)

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

                    '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                    '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)

                    Call AutoFitTable(wd, BOOLINCLUDEDATE)

                    strM = "Finalizing " & strTName & "..."
                    strM1 = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    str1 = strM1

                    frmH.lblProgress.Text = strM1
                    frmH.Refresh()


                    Call SplitTable(wd, 4, intLeg, arrLegend, str1, False, ctLegend + 2, False, False, False, intTableID)
                    'Sub SplitTable(ByVal wd As Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, 
                    'ByVal arr As Object, ByVal strT As String, ByVal DoLegend As Boolean, ByVal intSplitRows As Short, ByVal boolSmallFont As Boolean)

                    .Selection.Tables.Item(1).Select()
                    'autofit table
                    Call AutoFitTable(wd, False)
                    'move to line below table
                    Call MoveOneCellDown(wd)

                    Call InsertLegend(wd, intTableID, idTR, False, 1)

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

                    'If bool1 = False Then
                    '    var2 = "[NA]"
                    'Else
                    '    var2 = VerboseNumber(var4, True)
                    '    str2 = Replace(var1, CStr(var4), var2, 1, 1, CompareMethod.Text)
                    'End If

                    'str1 = str2 & " Final Extract Stability: Summary of " & rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION") & " Interpolated QC Standard Concentrations."
                    'str2 = str1
                    'str1 = rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")
                    ''Call JustTable(wd, str1, str2, strDo, strTName, intTableID)
                    'Call JustTable(wd, str1, strTName, strDo, strTName, intTableID, strTempInfo, "")

                    str1 = NZ(rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION"), "")
                    'If gNumMatrix = 1 Then
                    '    strA = strAnalC
                    'Else
                    '    strA = strAnal 'strAnalC has '..Matrix', don't want to pass that here
                    'End If
                    str1 = strA
                    'Call JustTable(wd, str1, str2, strDo, strTName, intTableID)

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

                    'If Len(str1) = 0 Then
                    'Else
                    '    Call JustTable(wd, str1, strTName, strDo, strTName, intTableID, strTempInfo, "")
                    'End If

                End If

next1:

            Next
end2:
        End With

    End Sub

    Sub MVSummaryFTStabilityQC_19(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal idTR As Int64)

        '20191017 LEE: Divert to MVAdHocQCStability_31

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
        Dim boolPro As Boolean

        Dim varConc

        Dim hi, lo
        Dim rows10() As DataRow
        Dim rows11() As DataRow
        Dim intRowsAnal As Short
        Dim arrFP(2, 20) 'FlagPercent array
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
        Dim boolSTable As Boolean
        Dim boolJustTable As Boolean

        Dim intExp As Short
        Dim ctExp As Short
        Dim int8 As Short

        Dim rows2E() As DataRow
        Dim nE As Short
        Dim nI As Short
        Dim boolOutHeadE As Boolean = False
        Dim boolOutHeadI As Boolean = False
        Dim boolDeleteRows As Boolean = False
        Dim boolOutlier As Boolean = False
        Dim intStart As Short

        Dim numPrec As Single
        Dim numTheor As Single

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

        Dim v1, v2, vU
        Dim intRC As Short

        boolJustTable = False

        boolSTable = False

        Cursor.Current = Cursors.WaitCursor

        fontsize = wd.ActiveDocument.Styles("Normal").Font.Size 'wd.Selection.Font.Size
        fonts = fontsize ' wd.Selection.Font.Size

        Dim charFCID As String
        strF = "ID_TBLREPORTTABLE = " & idTR
        Dim rowsTR() As DataRow = tblReportTable.Select(strF)
        var1 = rowsTR(0).Item("CHARFCID")
        charFCID = NZ(var1, "NA")

        With wd

            intTableID = 19

            Dim strWRunId As String = GetWatsonColH(intTableID)

            dvDo = frmH.dgvReportTableConfiguration.DataSource

            strF = "id_tblconfigreporttables = " & intTableID
            intDo = FindRowDVNumByCol(intTableID, dvDo, "id_tblconfigreporttables")

            ''Get table name
            'var1 = dvDo(intDo).Item("Table")
            'strTName = NZ(var1, "[NONE]")

            ''get Temperature info
            'var1 = dvDo(intDo).Item("PERIODTEMP")
            'strTempInfo = NZ(var1, "[NONE]")

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
            '    strM = "Creating Summary of " & strTempInfo & " Freeze/Thaw Stability in Matrix Table ...."
            '    frmH.lblProgress.Text = strM
            '    frmH.Refresh()
            '    MsgBox("Samples have not been assigned to this table.", MsgBoxStyle.Information, "Samples have not been assigned...")
            '    boolSTable = True
            '    GoTo end2
            'End If

            strF = "IsIntStd = 'No'"
            strS = "INTORDER ASC, IsIntStd ASC, AnalyteDescription ASC"
            rows11 = tblAnalytesHome.Select(strF, strS)
            intRowsAnal = rows11.Length

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

                Dim strM1 As String
                If bool Then 'continue
                    'ensure data has been entered
                    strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & CleanText(strDo) & "' AND ID_TBLREPORTTABLE = " & idTR
                    rowsX = tbl2.Select(strF)

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

                    If rowsX.Length = 0 Then
                        strM = "Creating " & strTName & "...."
                        strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                        strM1 = strM
                        frmH.lblProgress.Text = strM
                        frmH.Refresh()
                        'MsgBox("Samples have not been assigned to this table.", MsgBoxStyle.Information, "Samples have not been assigned...")
                        'page setup according to configuration
                        str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                        'insert page break
                        'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                        Call InsertPageBreak(wd)
                        Call PageSetup(wd, str1) 'L=Landscape, P=Portrait
                        boolJustTable = True
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

                    'determine if there are any outliers
                    'var1 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteIndex")
                    'var2 = tbl4.Rows.Item(Count1 - 1).Item("MasterAssayID")
                    'var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")

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

                    'vAnalyteIndex = var1
                    'vMasterAssayID = var2
                    'vAnalyteID = tbl4.Rows.Item(Count1 - 1).Item("AnalyteID")

                    'strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                    'strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                    'strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                    'strF2 = strF2 & "ANALYTEINDEX = " & var1 & " AND "
                    'strF2 = strF2 & "MASTERASSAYID = " & var2 & " AND (ELIMINATEDFLAG = 'Y' OR BOOLEXCLSAMPLE = -1)"

                    If boolUseGroups Then
                        strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                        strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                        strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                        strF2 = strF2 & "INTGROUP = " & intGroup & " AND (ELIMINATEDFLAG = 'Y' OR BOOLEXCLSAMPLE = -1)"
                    Else
                        strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                        strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                        strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                        strF2 = strF2 & "ANALYTEINDEX = " & var1 & " AND "
                        strF2 = strF2 & "MASTERASSAYID = " & var2 & " AND (ELIMINATEDFLAG = 'Y' OR BOOLEXCLSAMPLE = -1)"
                        'strF2 = strF2 & "ANALYTEID = " & var4 ' & "' AND "
                        'strF2 = strF2 & "BOOLINTSTD = 0"
                    End If

                    strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                    rows2 = tbl2.Select(strF2, strS)
                    int1 = rows2.Length 'debug
                    If int1 > 0 Then
                        boolOutlier = True
                    End If

                    'setup tables
                    'var1 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteIndex")
                    'var2 = tbl4.Rows.Item(Count1 - 1).Item("MasterAssayID")
                    'var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                    'strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                    'strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                    'strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                    'strF2 = strF2 & "ANALYTEINDEX = " & var1 & " AND "
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
                    int1 = rows2.Length 'debug
                    dv2 = New DataView(tbl2, strF2, strS, DataViewRowState.CurrentRows)
                    int1 = dv2.Count 'debug

                    'find number of runs used
                    tblNumRuns = dv2.ToTable("a", True, "RUNID")
                    intNumRuns = tblNumRuns.Rows.Count

                    ''search for dilution samples
                    'Dim boolHasDil As Boolean = False
                    'Dim dvW = frmH.dgvWatsonAnalRef.DataSource
                    'Dim strUnits As String
                    'int1 = FindRowDV("LLOQ Units", dv)
                    'strUnits = dvW.Item(int1).Item(1)

                    'int1 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvstudyconfig.DataSource)
                    'str1 = NZ(frmH.dgvStudyConfig(1, int1).Value, "")

                    'If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
                    'Else
                    '    strUnits = str1
                    'End If

                    'Dim tblDS as System.Data.Datatable = dv2.ToTable("ds", True, "NOMCONC", "ALIQUOTFACTOR")
                    'Dim strFDS As String
                    'Dim rowsDS() As DataRow
                    'Dim numDS As Single
                    'Dim numDSNomConc As Single
                    'strFDS = "NOMCONC = " & nomconc
                    'rowsDS = tblDS.Select(strFDS)
                    'numDS = rowsDS(0).Item("ALIQUOTFACTOR")
                    'numDSNomConc = rowsDS(0).Item("NOMCONC")
                    'If var1 <> 1 Then
                    '    boolHasDil = True
                    '    'record legend
                    '    'var1 = NZ(DilQCFactor(Count2), 1)
                    '    intLeg = intLeg + 1
                    '    ctDilLeg = ctDilLeg + 1
                    '    ctLegend = ctLegend + 1
                    '    'configure first legend item
                    '    var4 = numDSNomConc 'tblLevels.Rows(Count2 - 1).Item("NOMCONC")
                    '    var3 = Format(1 / CDec(var1), "0")
                    '    var1 = Chr(96 + intLeg) 'debugging
                    '    arrLegend(1, intLeg) = Chr(96 + intLeg) 'a,b,c,etc
                    '    'var: units
                    '    arrLegend(2, intLeg) = "Dilution QCs undiluted concentration " & SigFigOrDec(numDSNomConc, LSigFig, False, True) & " " & strUnits & "; a 1:" & var3 & " dilution with blank matrix was performed prior to extraction and analysis."
                    '    arrLegend(3, intLeg) = True
                    '    arrLegend(4, intLeg) = True
                    'End If

                    'establish table of level numbers
                    'must be sorted by nomconc!
                    'make new dv
                    Dim int12 As Short

                    Dim dvNL As New DataView(tbl2, strF2, "NOMCONC ASC", DataViewRowState.CurrentRows)
                    tblLevels = dvNL.ToTable("b", True, "NOMCONC", "CHARHELPER1")
                    intNumLevels = tblLevels.Rows.Count
                    For Count2 = 0 To intNumLevels - 1 'check for any null values
                        var3 = tblLevels.Rows.Item(Count2).Item("NOMCONC")
                        If IsDBNull(var3) Then
                            str1 = "The Nominal Concentration for some assigned samples for " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & " have not been configured."
                            str1 = str1 & ChrW(10) & "When this action is finished, please navigate to the Assigned Samples window and correct this problem."
                            If boolDisableWarnings Then
                            Else
                                MsgBox(str1, MsgBoxStyle.Information, "Nom Conc problem...")
                            End If
                            GoTo end1
                        End If
                        var3 = tblLevels.Rows.Item(Count2).Item("CHARHELPER1")
                        If IsDBNull(var3) Then
                            str1 = "The Term 1 designation for some assigned samples for " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & " have not been configured."
                            str1 = str1 & ChrW(10) & "When this action is finished, please navigate to the Assigned Samples window and correct this problem."
                            If boolDisableWarnings Then
                            Else
                                MsgBox(str1, MsgBoxStyle.Information, "Nom Conc problem...")
                            End If
                            GoTo end1
                        End If

                    Next

                    'find number of table rows to generate
                    intRowsX = 0
                    Dim intRowsXTot As Short = 0

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
                    intTblRows = intTblRows + 2 'for header
                    intTblRows = intTblRows + 1 'for blank row
                    intTblRows = intTblRows + (intRowsXTot) 'for number of data rows
                    'intTblRows = intTblRows + (intNumRuns * intRowsX) 'for number of data rows
                    intTblRows = intTblRows + (1 * intNumRuns) 'for a blank row after each run set

                    'intTblRows = intTblRows + (5 * intNumRuns) 'for Mean/Bias/n section for each run set

                    'Increment for Statistics Sections
                    Dim intCSN As Short
                    intCSN = countNumStatsRows()
                    intTblRows = intTblRows + (intCSN * intNumRuns)

                    If intCSN > 0 Then
                        intTblRows = intTblRows + (1 * intNumRuns) - 1 'for a blank row after each Mean/Bias/n set, except last set
                    Else
                        intTblRows = intTblRows - 1 'subtract an unneeded blank row
                    End If


                    If boolQCREPORTACCVALUES Then

                    Else
                        If boolOutlier Then
                            intTblRows = intTblRows + (3 * intNumRuns) 'for stats headings
                            ctExp = ctExp + 3 'for stats headings

                            'Increment for Statistics Sections
                            intTblRows = intTblRows + (intCSN * intNumRuns)
                            ctExp = ctExp + (intCSN * intNumRuns)

                            If intCSN > 0 Then
                                intTblRows = intTblRows + (1 * intNumRuns) - 1 'for a blank row after each Mean/Bias/n set, except last set
                                ctExp = ctExp + (1 * intNumRuns) - 1
                            End If
                        
                        End If

                    End If

                    wrdSelection = wd.Selection()

                    Dim intCols As Short
                    If boolSTATSDIFFCOL Then
                        intCols = (intNumLevels * 2) + 1
                    Else
                        intCols = intNumLevels + 1
                    End If

                    Dim boolEnterDiff As Boolean


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

                        .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.MoveRight Unit:=Microsoft.Office.Interop.Word.wdunits.word.wdunits.wdCharacter, Count:=ctQCs, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                        '.Selection.MoveLeft(Microsoft.Office.Interop.Word.WdUnits.wdCharacter, 1)

                        'Enter nom. conc. row titles
                        .Selection.Tables.Item(1).Cell(2, 2).Select()
                        Dim strE As String
                        For Count2 = 0 To intNumLevels - 1
                            'var1 = arrBCQCs(3, Count2)
                            If boolLUseSigFigs Then
                                var1 = CStr(SigFigOrDecString(tblLevels.Rows.Item(Count2).Item("NOMCONC"), LSigFig, False))
                            Else
                                var1 = CStr(Format(RoundToDecimalRAFZ(tblLevels.Rows.Item(Count2).Item("NOMCONC"), LSigFig), GetRegrDecStr(LSigFig)))
                            End If

                            var2 = tblLevels.Rows.Item(Count2).Item("CHARHELPER1")
                            var3 = ReturnStdQC(var2.ToString)
                            'var2 = var3
                            'var3 = var2 & " " & var1 & " " & strConcUnits
                            'var3 = var2 ' & ChrW(10) & var1 & " " & strConcUnits
                            If LboolNomConcParen Then
                                strE = ChrW(10) & "(" & var1 & ChrW(160) & strConcUnits & ")"
                            Else
                                strE = ChrW(10) & var1 & ChrW(160) & strConcUnits
                            End If

                            .Selection.TypeText(Text:=var3)

                            '****determine if this level is a dilution level
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
                                        'arrLegend(2, intLeg) = "Dilution QCs undiluted concentration " & Format(RoundToDecimalRAFZ(CDbl(Val(var1)), LSigFig), GetRegrDecStr(LSigFig)) & " " & strConcUnits & "; a 1:" & var3 & " dilution with blank matrix was performed prior to extraction and analysis."
                                        arrLegend(2, intLeg) = "Dilution QCs undiluted concentration " & Format(RoundToDecimalRAFZ(CDbl(Val(var1)), LSigFig), GetRegrDecStr(LSigFig)) & " " & strConcUnits & "; " & strAN & " " & var3 & "-fold dilution with blank matrix was performed prior to extraction and analysis."
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

                            .Selection.TypeText(strE)


                            '******

                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                            If boolSTATSDIFFCOL Then
                                .Selection.TypeText(Text:=ReturnDiffLabel)
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                            End If
                        Next

                        .Selection.Tables.Item(1).Cell(2, 1).Select()
                        'bottom border this row
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .Selection.Tables.Item(1).Cell(2, 1).Select()

                        'begin entering data'
                        If BOOLINCLUDEDATE Then
                            '.Selection.TypeText(strWRunId & ChrW(10) & "(Analysis Date)")
                            '20180420 LEE:
                            .Selection.TypeText(strWRunId & ChrW(10) & "(" & GetAnalysisDateLabel(intTableID) & ")")
                        Else
                            .Selection.TypeText(strWRunId)
                        End If

                        int1 = 4 'row position counter
                        For Count2 = 0 To intNumRuns - 1

                            .Selection.Tables.Item(1).Cell(int1, 1).Select()
                            'enter runid
                            var10 = tblNumRuns.Rows.Item(Count2).Item("RUNID")

                            'strM = "Creating Summary of " & strTempInfo & " Freeze/Thaw Stability in Matrix Table For " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & "..."
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

                            'start filling in data by columns
                            'intRowsX = 0

                            '20160816 LEE: intRowsX needs to be determined for each dataset
                            '****
                            dv2.RowFilter = ""
                            'don't know why, but must make a long filter here or
                            'both analytes get returned in dv2.rowfilter

                            intRowsX = 0
                            int12 = -1
                            For Count3 = 0 To tblLevels.Rows.Count - 1
                                int12 = int12 + 1
                                varNom = tblLevels.Rows.Item(int12).Item("NOMCONC")
                                strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom
                                dv2.RowFilter = strF
                                int2 = dv2.Count
                                If int2 > intRowsX Then
                                    intRowsX = int2
                                End If
                            Next

                            '****


                            int12 = -1
                            For Count3 = 0 To (intNumLevels * int11) - 1 Step int11
                                int12 = int12 + 1
                                varNom = tblLevels.Rows.Item(int12).Item("NOMCONC")

                                'determine hi and lo (nom*flagpercent)
                                strF = "CONCENTRATION = '" & varNom & "'"
                                'rows10 = tblBCQCs.Select(strF)

                                'determine hi and lo (nom*flagpercent)
                                'strF = "CONCENTRATION = " & varNom & " AND ANALYTEID = " & vAnalyteID & " AND MASTERASSAYID = " & vMasterAssayID & " AND ANALYTEINDEX = " & vAnalyteIndex & " AND CONCENTRATION = " & varNom & " AND RUNID = " & var10

                                'if Conc < 1, then the query return 0 records
                                'must do something different
                                var1 = GetANALYTEFLAGPERCENT(varNom, var10, vAnalyteID)
                                'var1 = CDec(NZ(rows10(0).Item("FLAGPERCENT"), 15))
                                arrFP(1, int12) = var1
                                arrFP(2, int12) = var1
                                Call SetHighAndLowCriteria(varNom, var1, var1, hi, lo)

                                'start entering data
                                dv2.RowFilter = ""
                                'don't know why, but must make a long filter here or
                                'both analytes get returned in dv2.rowfilter
                                strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom
                                dv2.RowFilter = strF
                                int2 = dv2.Count

                                'get inexcluded dataset
                                strF = ""
                                'For Count4 = 0 To dv2.Count - 1
                                '    var2 = dv2(Count4).Item("ANALYTEINDEX")
                                '    var3 = dv2(Count4).Item("MASTERASSAYID")
                                '    var4 = dv2(Count4).Item("RUNSAMPLEORDERNUMBER")
                                '    var5 = dv2(Count4).Item("ANALYTEID")

                                '    If Count4 <> dv2.Count - 1 Then
                                '        strF = strF & "(RUNID = " & var10 & " AND ANALYTEID = " & var5 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLEORDERNUMBER = " & var4 & ") OR "
                                '    Else
                                '        strF = strF & "(RUNID = " & var10 & " AND ANALYTEID = " & var5 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLEORDERNUMBER = " & var4 & ")"
                                '    End If
                                'Next
                                Erase rows2
                                'rows2 = tbl1.Select(strF)
                                Dim tbl2s As System.Data.DataTable = dv2.ToTable
                                rows2 = tbl2s.Select("RUNID > -1")
                                int3 = rows2.Length
                                'int3 = rows2.Length
                                nI = int3

                                'redo hi/lo
                                If rows2.Length = 0 Then
                                    vU = 0
                                    v1 = 0
                                    v2 = 0
                                Else
                                    vU = rows2(0).Item("BOOLUSEGUWUACCCRIT")
                                    If gAllowGuWuAccCrit And LAllowGuWuAccCrit And vU = -1 Then
                                        v1 = CDec(NZ(rows2(0).Item("NUMMAXACCCRIT"), 0))
                                        v2 = CDec(NZ(rows2(0).Item("NUMMINACCCRIT"), 0))
                                        arrFP(1, int12) = v1
                                        arrFP(2, int12) = v2

                                        Call SetHighAndLowCriteria(varNom, v1, v2, hi, lo)

                                    Else
                                        strF = "CONCENTRATION = '" & varNom & "'"
                                        'rows10 = tblBCQCs.Select(strF)

                                        'determine hi and lo (nom*flagpercent)
                                        'strF = "CONCENTRATION = " & varNom & " AND ANALYTEID = " & vAnalyteID & " AND MASTERASSAYID = " & vMasterAssayID & " AND ANALYTEINDEX = " & vAnalyteIndex & " AND CONCENTRATION = " & varNom & " AND RUNID = " & var10

                                        'if Conc < 1, then the query return 0 records
                                        'must do something different
                                        var1 = GetANALYTEFLAGPERCENT(varNom, var10, vAnalyteID)

                                        'var1 = CDec(NZ(rows10(0).Item("FLAGPERCENT"), 15))
                                        arrFP(1, int12) = var1
                                        arrFP(2, int12) = var1

                                        Call SetHighAndLowCriteria(varNom, var1, var1, hi, lo)
                                        v1 = var1
                                        v2 = var1

                                    End If
                                End If


                                'create rows1 from tbl1 which will contain data
                                strF = ""

                                Erase rows2E
                                'rows2E = tbl1.Select(strF)
                                If gAllowExclSamples And LAllowExclSamples Then
                                    rows2E = tbl2s.Select("ELIMINATEDFLAG = 'N' AND BOOLEXCLSAMPLE = 0")
                                Else
                                    rows2E = tbl2s.Select("ELIMINATEDFLAG = 'N'")
                                End If
                                nE = rows2E.Length

                                'enter data

                                For Count4 = 0 To intRowsX - 1 'int3 - 1

                                    boolOC = False
                                    boolEnterDiff = False
                                    .Selection.Tables.Item(1).Cell(int1 + Count4, Count3 + 2).Select()
                                    If Count4 > int3 - 1 Then
                                        If boolQCNA Then
                                            str1 = "NA"
                                        Else
                                            str1 = ""
                                        End If
                                        .Selection.TypeText(str1)
                                        boolEnterDiff = False
                                    Else
                                        var1 = rows2(Count4).Item("CONCENTRATION")
                                        varConc = var1
                                        var1 = NZ(var1, 0)
                                        numDF = rows2(Count4).Item("ALIQUOTFACTOR")
                                        var1 = var1 / numDF
                                        If boolLUseSigFigs Then
                                            var2 = SigFigOrDec(var1, LSigFig, False)
                                        Else
                                            var2 = RoundToDecimalRAFZ(var1, LSigFig)
                                        End If

                                        var1 = NZ(rows2(Count4).Item("ELIMINATEDFLAG"), "N")
                                        var3 = NZ(rows2(Count4).Item("BOOLEXCLSAMPLE"), 0)
                                        vU = NZ(rows2(Count4).Item("BOOLUSEGUWUACCCRIT"), 0)
                                        boolEnterDiff = True 'FALSE

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

                                            intExp = intExp + 1
                                            intLeg = intLeg + 1
                                            strA = ChrW(intLeg + intLegStart)

                                            '20160305 LEE:
                                            'Added DECISIONREASON code
                                            Dim var6
                                            'Remember, tblAssignedSamples does not have DECISIONREASON
                                            var6 = "No Value: " & GetDECISIONREASONValue(boolExFromAS, vAnalyteID, rows2(Count4))
                                            'Set Legend String
                                            str1 = GetLegendStringExcluded(v1, v2, vU, var6, intTableID, True, "")
                                            'Add to Legend Array
                                            ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)


                                            If boolRedBoldFont Then
                                                .Selection.Font.Bold = True
                                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                            End If

                                            .Selection.TypeText(Text:="NV")

                                            Call typeInSuperscriptFontSize12WithSpace(wd, strA)

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

                                            intExp = intExp + 1
                                            intLeg = intLeg + 1
                                            strA = ChrW(intLeg + intLegStart)

                                            '20160305 LEE:
                                            'Added DECISIONREASON code
                                            Dim var6
                                            'Remember, tblAssignedSamples does not have DECISIONREASON
                                            var6 = GetDECISIONREASONValue(boolExFromAS, vAnalyteID, rows2(Count4))
                                            'Set Legend String
                                            str1 = GetLegendStringExcluded(arrFP(1, int12), arrFP(2, int12), vU, var6, intTableID, True, "")
                                            'Add to Legend Array
                                            ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)


                                            If boolRedBoldFont Then
                                                .Selection.Font.Bold = True
                                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                            End If

                                            If boolLUseSigFigs Then
                                                .Selection.TypeText(Text:=DisplayNum(var2, LSigFig, False))
                                            Else
                                                .Selection.TypeText(Text:=Format(var2, GetRegrDecStr(LSigFig)))
                                            End If

                                            Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                            '.Selection.TypeText Text:="NR"
                                        Else

                                            boolEnterDiff = True
                                            'determine if value is outside acceptance criteria

                                            'If var2 > hi Or var2 < lo Then 'flag
                                            If OutsideAccCrit(var2, varNom, v1, v2, NZ(vU, 0)) Then
                                                intLeg = intLeg + 1
                                                strA = ChrW(intLeg + intLegStart)

                                                'Set Legend String
                                                str1 = GetLegendStringIncluded(v1, v2, vU)
                                                'Add to Legend Array
                                                ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                                If boolRedBoldFont Then
                                                    .Selection.Font.Bold = True
                                                    .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                End If

                                                If boolLUseSigFigs Then
                                                    .Selection.TypeText(Text:=CStr(DisplayNum(var2, LSigFig, False)))
                                                Else
                                                    .Selection.TypeText(Text:=Format(var2, GetRegrDecStr(LSigFig)))
                                                End If

                                                Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                            Else
                                                If boolLUseSigFigs Then
                                                    .Selection.TypeText(Text:=CStr(DisplayNum(var2, LSigFig, False)))
                                                Else
                                                    .Selection.TypeText(Text:=Format(var2, GetRegrDecStr(LSigFig)))
                                                End If
                                                boolEnterDiff = True
                                            End If
                                        End If

                                    End If

                                    If boolSTATSDIFFCOL Then
                                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                        If boolEnterDiff Then
                                            'var3 = Format(((var2 / varNom) - 1) * 100, strQCDec)
                                            'var3 = Format(RoundToDecimal(((var2 / varNom) - 1) * 100, intQCDec), strQCDec)


                                            If boolTHEORETICAL Then
                                                var3 = CalcREPercent(var2, varNom, intQCDec)
                                                numTheor = 100 + CDec(var3)

                                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", numTheor, CSng(var10), Count1, strDo, v1, v2, boolOC)

                                            Else
                                                var3 = Format(RoundToDecimal(CalcREPercent(var2, varNom, intQCDec), intQCDec), strQCDec)

                                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", var3, CSng(var10), Count1, strDo, v1, v2, boolOC)

                                            End If

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

                                Next Count4


                                'now enter Mean/Bias/n for each analytical run

                                int8 = 0

                                If boolQCREPORTACCVALUES Then
                                Else
                                    If boolOutlier Then
                                        int8 = int8 + 1
                                    End If
                                    If boolOutlier And Count3 = 0 Then
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, 1).Select()

                                        .Selection.TypeText(Text:="Summary Statistics Excluding Outlier Values")
                                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                        Try
                                            .Selection.Cells.Merge()
                                        Catch ex As Exception
                                        End Try
                                        With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                            .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                        End With

                                    End If
                                End If


                                If Count3 = 0 Then

                                    Call typeStatsLabels(wd, int8, int1 + intRowsX, 1, False)

                                    If boolQCREPORTACCVALUES Then
                                    Else
                                        If boolOutlier Then
                                            int8 = int8 + 2
                                        End If
                                        If boolOutlier And Count3 = 0 Then
                                            .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, 1).Select()

                                            .Selection.TypeText(Text:="Summary Statistics Including Outlier Values")
                                            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                            Try
                                                .Selection.Cells.Merge()
                                            Catch ex As Exception
                                            End Try
                                            With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                                .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                            End With

                                            typeStatsLabels(wd, int8, int1 + intRowsX, 1, False)
                                        End If
                                    End If
                                End If

                                int8 = 0
                                If boolQCREPORTACCVALUES Then
                                Else
                                    If boolOutlier Then
                                        int8 = int8 + 1
                                    End If
                                End If

                                v1 = arrFP(1, int12)
                                v2 = arrFP(2, int12)

                                Dim boolMean As Boolean = True

                                Try
                                    var1 = MeanDR(rows2E, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                    If boolLUseSigFigs Then
                                        numMean = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                    Else
                                        numMean = RoundToDecimalRAFZ(var1, LSigFig)
                                    End If

                                    Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Mean", numMean, CSng(var10), Count1, strDo, 0, 0, False)
                                Catch ex As Exception

                                End Try
                                If boolSTATSMEAN Then
                                    Try
                                        'enter Mean
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                        If nE = 0 Then
                                            .Selection.TypeText(Text:="NA")
                                        Else

                                            If nE = 0 Then
                                                boolMean = False
                                                .Selection.TypeText("NA")
                                            Else

                                                'determine if value is outside acceptance criteria
                                                'If (numMean > hi Or numMean < lo) And boolFootNoteQCMean Then 'flag
                                                If (OutsideAccCrit(numMean, varNom, v1, v2, NZ(vU, 0))) And boolFootNoteQCMean Then 'flag
                                                    intLeg = intLeg + 1
                                                    strA = ChrW(intLeg + intLegStart)
                                                    'Set Legend String
                                                    str1 = GetLegendStringIncluded(v1, v2, vU)
                                                    'Add to Legend Array
                                                    ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                                    If boolRedBoldFont Then
                                                        .Selection.Font.Bold = True
                                                        .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                    End If

                                                    If boolLUseSigFigs Then
                                                        .Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                                    Else
                                                        .Selection.TypeText(Text:=Format(numMean, GetRegrDecStr(LSigFig)))
                                                    End If

                                                    Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                                    boolEnterDiff = True
                                                Else
                                                    If boolLUseSigFigs Then
                                                        .Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                                    Else
                                                        .Selection.TypeText(Text:=Format(numMean, GetRegrDecStr(LSigFig)))
                                                    End If
                                                    boolEnterDiff = True
                                                End If
                                            End If

                                        End If
                                        '.Selection.TypeText(CStr(numMean))




                                    Catch ex As Exception

                                    End Try
                                End If

                                Try
                                    var1 = StdDevDR(rows2E, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                    If boolLUseSigFigs Then
                                        numSD = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                    Else
                                        numSD = RoundToDecimalRAFZ(var1, LSigFig)
                                    End If
                                    Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "SD", numSD, CSng(var10), Count1, strDo, 0, 0, False)
                                Catch ex As Exception

                                End Try
                                If boolSTATSSD Then
                                    Try
                                        'enter SD
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()

                                        If nE < gSDMax Then
                                            .Selection.TypeText("NA")
                                        Else
                                            var1 = StdDevDR(rows2E, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                            If boolLUseSigFigs Then
                                                .Selection.TypeText(DisplayNum(numSD, LSigFig, False))
                                            Else
                                                .Selection.TypeText(Format(numSD, GetRegrDecStr(LSigFig)))
                                            End If

                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If


                                Try
                                    If nE < gSDMax Then
                                    Else
                                        numPrec = CalcCVPercent(numSD, numMean, intQCDec)
                                        Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Precision", numPrec, CSng(var10), Count1, strDo, 0, 0, False)
                                    End If

                                Catch ex As Exception

                                End Try
                                If boolSTATSCV Then
                                    Try
                                        'enter %CV
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                        If nE < gSDMax Then
                                            .Selection.TypeText("NA")
                                        Else
                                            .Selection.TypeText(Format(numPrec, strQCDec))
                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If


                                If boolSTATSBIAS And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                    Try
                                        numBias = CalcREPercent(numMean, varNom, intQCDec)
                                        If nE = 0 Then
                                        Else
                                            Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", numBias, CSng(var10), Count1, strDo, 0, 0, False)
                                        End If

                                    Catch ex As Exception

                                    End Try
                                Else
                                    'get numbias from average of %Bias columns
                                    numBias = GetBiasFromDiffCol(idTR, varNom, int12 + 1, 0, False)
                                End If
                                If boolSTATSBIAS And boolSTATSMEAN Then
                                    Try
                                        'enter %Bias
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                        If nE = 0 Then
                                            .Selection.TypeText(Text:="NA")
                                        Else
                                            .Selection.TypeText(Format(numBias, strQCDec))
                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If


                                If boolTHEORETICAL And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                    Try
                                        numTheor = CalcREPercent(numMean, varNom, intQCDec)
                                        numTheor = 100 + CDec(numTheor)
                                        If nE = 0 Then
                                        Else
                                            Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", numTheor, CSng(var10), Count1, strDo, 0, 0, False)
                                        End If

                                    Catch ex As Exception

                                    End Try
                                Else
                                    'get numbias from average of %Bias columns
                                    numTheor = GetBiasFromDiffCol(idTR, varNom, int12 + 1, 0, False)
                                    numTheor = 100 + CDec(numTheor)
                                End If

                                If boolTHEORETICAL And boolSTATSMEAN Then
                                    Try
                                        'enter %theoretical
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()

                                        If nE = 0 Then
                                            .Selection.TypeText(Text:="NA")
                                        Else
                                            .Selection.TypeText(Format(numTheor, strQCDec))
                                        End If

                                    Catch ex As Exception

                                    End Try

                                End If


                                If boolSTATSDIFF And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                    Try
                                        numBias = CalcREPercent(numMean, varNom, intQCDec)
                                        If nE = 0 Then
                                        Else
                                            Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", numBias, CSng(var10), Count1, strDo, 0, 0, False)
                                        End If

                                    Catch ex As Exception

                                    End Try
                                Else
                                    'get numbias from average of %Bias columns
                                    numBias = GetBiasFromDiffCol(idTR, varNom, int12 + 1, 0, False)
                                End If

                                If boolSTATSDIFF And boolSTATSMEAN Then
                                    Try
                                        'enter %Bias
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                        If nE = 0 Then
                                            .Selection.TypeText(Text:="NA")
                                        Else

                                            .Selection.TypeText(Format(numBias, strQCDec))
                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If


                                If BOOLSTATSRE And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then

                                    Try
                                        numBias = CalcREPercent(numMean, varNom, intQCDec)
                                        If nE = 0 Then
                                        Else
                                            Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", numBias, CSng(var10), Count1, strDo, 0, 0, False)
                                        End If

                                    Catch ex As Exception

                                    End Try
                                Else
                                    'get numbias from average of %Bias columns
                                    numBias = GetBiasFromDiffCol(idTR, varNom, int12 + 1, 0, False)
                                End If
                                If BOOLSTATSRE And boolSTATSMEAN Then
                                    Try
                                        'enter %RE
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                        If nE = 0 Then
                                            .Selection.TypeText(Text:="NA")
                                        Else

                                            .Selection.TypeText(Format(numBias, strQCDec))
                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If


                                Try
                                    Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "n", nE, CSng(var10), Count1, strDo, 0, 0, False)
                                Catch ex As Exception

                                End Try
                                If boolSTATSN Then
                                    Try
                                        'enter n
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                        .Selection.TypeText(CStr(nE))
                                    Catch ex As Exception

                                    End Try
                                End If

                                If boolQCREPORTACCVALUES Then
                                Else
                                    If boolOutlier Then
                                        int8 = int8 + 2

                                        boolMean = True
                                        If boolSTATSMEAN Then
                                            Try
                                                'enter Mean
                                                int8 = int8 + 1
                                                .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()

                                                If nE = 0 Then
                                                    .Selection.TypeText(Text:="NA")
                                                Else
                                                    var1 = MeanDR(rows2, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                                    If boolLUseSigFigs Then
                                                        numMean = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                                    Else
                                                        numMean = RoundToDecimalRAFZ(var1, LSigFig)
                                                    End If

                                                    If nI = 0 Then
                                                        boolMean = False
                                                        .Selection.TypeText("NA")
                                                    Else
                                                        '.Selection.TypeText(CStr(numMean))

                                                        'determine if value is outside acceptance criteria
                                                        'If (numMean > hi Or numMean < lo) And boolFootNoteQCMean Then 'flag
                                                        If (OutsideAccCrit(numMean, varNom, v1, v2, NZ(vU, 0))) And boolFootNoteQCMean Then 'flag
                                                            intLeg = intLeg + 1
                                                            strA = ChrW(intLeg + intLegStart)

                                                            'Set Legend String
                                                            str1 = GetLegendStringIncluded(v1, v2, vU)
                                                            'Add to Legend Array
                                                            ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                                            If boolRedBoldFont Then
                                                                .Selection.Font.Bold = True
                                                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                            End If

                                                            If boolLUseSigFigs Then
                                                                .Selection.TypeText(Text:=DisplayNum(numMean, LSigFig, False))
                                                            Else
                                                                .Selection.TypeText(Text:=Format(numMean, GetRegrDecStr(LSigFig)))
                                                            End If

                                                            Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                                            boolEnterDiff = True
                                                        Else
                                                            If boolLUseSigFigs Then
                                                                .Selection.TypeText(Text:=DisplayNum(numMean, LSigFig, False))
                                                            Else
                                                                .Selection.TypeText(Text:=Format(numMean, GetRegrDecStr(LSigFig)))
                                                            End If
                                                            boolEnterDiff = True
                                                        End If
                                                    End If
                                                End If

                                            Catch ex As Exception

                                            End Try
                                        End If
                                        If boolSTATSSD Then
                                            Try
                                                'enter SD
                                                int8 = int8 + 1
                                                .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                                If nI < gSDMax Then
                                                    .Selection.TypeText("NA")
                                                Else
                                                    var1 = StdDevDR(rows2, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                                    If boolLUseSigFigs Then
                                                        numSD = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                                    Else
                                                        numSD = RoundToDecimalRAFZ(var1, LSigFig)
                                                    End If

                                                    If boolLUseSigFigs Then
                                                        .Selection.TypeText(Text:=DisplayNum(numSD, LSigFig, False))
                                                    Else
                                                        .Selection.TypeText(Text:=Format(numSD, GetRegrDecStr(LSigFig)))
                                                    End If
                                                End If

                                            Catch ex As Exception

                                            End Try
                                        End If

                                        If boolSTATSCV Then
                                            Try
                                                'enter %CV
                                                int8 = int8 + 1
                                                .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                                If nI < gSDMax Then
                                                    .Selection.TypeText("NA")
                                                Else
                                                    numPrec = CalcCVPercent(numSD, numMean, intQCDec)
                                                    .Selection.TypeText(Format(numPrec, strQCDec))
                                                End If

                                            Catch ex As Exception

                                            End Try
                                        End If
                                        If boolSTATSBIAS And boolSTATSMEAN Then
                                            Try
                                                'enter %Bias
                                                int8 = int8 + 1
                                                .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                                'var1 = (((numMean / varNom) - 1) * 100)
                                                'var1 = Format(var1, strQCDec)
                                                '.Selection.TypeText(CStr(var1))

                                                numBias = CalcREPercent(numMean, varNom, intQCDec)

                                                If boolMean Then
                                                    .Selection.TypeText(Format(numBias, strQCDec))
                                                Else
                                                    .Selection.TypeText("NA")
                                                End If

                                            Catch ex As Exception

                                            End Try
                                        End If

                                        If boolTHEORETICAL And boolSTATSMEAN Then
                                            Try
                                                'enter %theoretical
                                                int8 = int8 + 1
                                                .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                                'var1 = (((numMean / varNom) - 1) * 100)
                                                'var1 = Format(var1, strQCDec)
                                                'var1 = Format(100 + var1, strQCDec)
                                                '.Selection.TypeText(CStr(var1))
                                                If nI = 0 Then
                                                    .Selection.TypeText(Text:="NA")
                                                Else
                                                    numTheor = CalcREPercent(numMean, varNom, intQCDec)
                                                    numTheor = 100 + CDec(numTheor)
                                                    .Selection.TypeText(Format(numTheor, strQCDec))
                                                End If

                                            Catch ex As Exception

                                            End Try

                                        End If

                                        If boolSTATSDIFF And boolSTATSMEAN Then
                                            Try
                                                'enter %diff
                                                int8 = int8 + 1
                                                .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                                'var1 = (((numMean / varNom) - 1) * 100)
                                                'var1 = Format(var1, strQCDec)
                                                '.Selection.TypeText(CStr(var1))
                                                If nI = 0 Then
                                                    .Selection.TypeText(Text:="NA")
                                                Else
                                                    numBias = CalcREPercent(numMean, varNom, intQCDec)
                                                    .Selection.TypeText(Format(numBias, strQCDec))
                                                End If

                                            Catch ex As Exception

                                            End Try
                                        End If

                                        If BOOLSTATSRE And boolSTATSMEAN Then
                                            Try
                                                'enter %RE
                                                int8 = int8 + 1
                                                .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                                'var1 = (((numMean / varNom) - 1) * 100)
                                                'var1 = Format(var1, strQCDec)
                                                '.Selection.TypeText(CStr(var1))
                                                If nI = 0 Then
                                                    .Selection.TypeText(Text:="NA")
                                                Else
                                                    numBias = CalcREPercent(numMean, varNom, intQCDec)
                                                    .Selection.TypeText(Format(numBias, strQCDec))
                                                End If

                                            Catch ex As Exception

                                            End Try
                                        End If

                                        If boolSTATSN Then
                                            Try
                                                'enter n
                                                int8 = int8 + 1
                                                .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                                .Selection.TypeText(CStr(nI))

                                            Catch ex As Exception

                                            End Try
                                        End If
                                    End If
                                End If

                                'end Mean/Bias/n

                            Next Count3 'number of QC levels

                            'increase row position counter
                            If Count2 = intNumRuns - 1 Then
                                int1 = int1 + intRowsX + int8 + 1
                            Else
                                If intCSN = 0 Then 'this is number of stats rows
                                    int1 = int1 + intRowsX + int8 + 1
                                Else
                                    int1 = int1 + intRowsX + int8 + 2
                                End If

                            End If

                        Next Count2 'number of run ids

                        'bottom border this row
                        .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.MoveRight Unit:=Microsoft.Office.Interop.Word.wdunits.word.wdunits.wdCharacter, Count:=ctQCs, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                        'If boolQCREPORTACCVALUES Then
                        'Else
                        '    If intExp = 0 Then
                        '        Call DeleteRows(ctExp, wd)
                        '    End If
                        'End If

                        'autofit table
                        Call AutoFitTable(wd, False)

                        'go back and merge line 1
                        .Selection.Tables.Item(1).Cell(1, 2).Select()
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.Cells.Merge()
                        If intNumLevels < 2 Then
                        Else
                            Try
                                .Selection.Cells.Merge()
                            Catch ex As Exception

                            End Try
                        End If
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


                    'go to end of table
                    .Selection.Tables.Item(1).Cell(1, 1).Select()
                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdColumn)

                    'enter table number
                    str1 = "Summary of " & rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION") & " Freeze/Thaw " & strTempInfo & " Stability in Matrix"

                    '***
                    strA = rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")
                    strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, False)
                    Call EnterTableNumber(wd, strTName, 3, strA, strTempInfo, intTableID, intGroup, idTR)
                    '***

                    'Call EnterTableNumber(wd, str1, 3)

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

                    '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                    '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)

                    Call AutoFitTable(wd, BOOLINCLUDEDATE)

                    strM = "Finalizing " & strTName & "..."
                    strM1 = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    str1 = strM1

                    frmH.lblProgress.Text = strM1
                    frmH.Refresh()

                    Call SplitTable(wd, 4, intLeg, arrLegend, str1, False, ctLegend + 5, False, False, False, intTableID)
                    'Sub SplitTable(ByVal wd As Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, 
                    'ByVal arr As Object, ByVal strT As String, ByVal DoLegend As Boolean, ByVal intSplitRows As Short, ByVal boolSmallFont As Boolean)

                    'move to line below table
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

                '    'str1 = rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")
                '    'str2 = "Summary of " & rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION") & " Freeze/Thaw " & strTempInfo & " Stability in Matrix"
                '    ''Call JustTable(wd, str1, str2, strDo, strTName, intTableID, strTempInfo,"")
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
            If boolSTable Then
                ''enter table number
                'str1 = "Summary of " & rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION") & " Freeze/Thaw " & strTempInfo & " Stability in Matrix"
                'Call EnterTableNumber(wd, str1, 3)

            End If
        End With

    End Sub


    Sub MVSummarySpikingSolnStability_23(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal idTR As Int64)

        '20191017 LEE: Divert to MVAdHocQCStabilityComparison_32

        'BIG NOTE: Stock Soln Stability has no calibr standards
        '   So there are no data in ANARUNANALYTERESULTS
        '   So can't ELIMINATEFLAG any data, because ELIMINATEFLAG is in ANARUNANALYTERESULTS

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
        Dim tbl2a As New System.Data.DataTable
        Dim tbl2b As New System.Data.DataTable
        Dim dv2 As system.data.dataview
        Dim rows2() As DataRow
        Dim rows2a() As DataRow
        Dim rows2b() As DataRow
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
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim strF As String
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim int4 As Short
        Dim int10 As Short
        Dim intRowsX As Short
        Dim tblX As New System.Data.DataTable
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
        Dim arrFP(20) 'FlagPercent array
        Dim strFP As String
        Dim numMean As Decimal
        Dim numBias As Decimal
        Dim numSD As Decimal
        Dim tblZ As New System.Data.DataTable
        Dim dvAn As system.data.dataview
        Dim p1, p2, p3, p4, p5, p6, p7, p8, p9, p10
        Dim strM As String
        Dim fonts
        Dim numDF As Decimal
        Dim DilFactor
        Dim strF2a As String
        Dim strTempInfo As String
        Dim rowsQC() As DataRow
        Dim rowsRS() As DataRow
        Dim introwsQC As Short
        Dim introwsRS As Short
        Dim arr1(1)
        Dim numA, numB
        Dim numA1 As Decimal
        Dim numB1 As Decimal
        Dim boolIS As Boolean
        Dim strX As String
        Dim rowsX() As DataRow
        Dim boolX As Boolean
        Dim intCols As Short
        Dim col1, col2, col3, col4, col5, col6 As Short
        Dim intLegStart As Short
        Dim strUnits As String
        Dim varNomConc
        Dim numStd As Short
        Dim rowStd() As DataRow
        Dim intStdRows As Short
        Dim boolJustTable As Boolean

        Dim strMatrix As String
        Dim intGroup As Short

        Dim intExp As Short
        Dim ctExp As Short
        Dim int8 As Short

        Dim rows2E() As DataRow
        Dim nE As Short
        Dim nI As Short
        Dim boolOutHeadE As Boolean = False
        Dim boolOutHeadI As Boolean = False
        Dim boolDeleteRows As Boolean = False
        Dim boolOutlier As Boolean = False
        Dim intStart As Short


        Dim vAnalyteIndex
        Dim vMasterAssayID
        Dim vAnalyteID
        Dim tblAG As DataTable = tblAnalyteGroups 'tblAnalyteGroups has all analytes, not just accepted

        Dim strAnal As String
        Dim strAnalC As String
        Dim strTNameO As String
        Dim intRunID As Int16
        Dim strDECISIONREASON As String
        Dim boolExFromAS As Boolean

        Dim v1, v2, vU


        Dim numPrec As Single
        Dim numTheor As Single

        Dim numMeanFP As Decimal 'full precision
        Dim numAFP As Decimal
        Dim numBFP As Decimal


        '''''wdd.visible = True

        boolJustTable = False

        Cursor.Current = Cursors.WaitCursor

        fontsize = wd.ActiveDocument.Styles("Normal").Font.Size ' wd.Selection.Font.Size
        fonts = fontsize ' wd.Selection.Font.Size

        Dim charFCID As String
        strF = "ID_TBLREPORTTABLE = " & idTR
        Dim rowsTR() As DataRow = tblReportTable.Select(strF)
        var1 = rowsTR(0).Item("CHARFCID")
        charFCID = NZ(var1, "NA")

        With wd

            intTableID = 23

            Dim strWRunId As String = GetWatsonColH(intTableID)

            dvDo = frmH.dgvReportTableConfiguration.DataSource
            strF = "id_tblconfigreporttables = " & intTableID
            intDo = FindRowDVNumByCol(intTableID, dvDo, "id_tblconfigreporttables")

            ''Get table name
            'var1 = dvDo(intDo).Item("Table")
            'strTName = NZ(var1, "[NONE]")

            ''get Temperature info
            'var1 = dvDo(intDo).Item("PERIODTEMP")
            'strTempInfo = NZ(var1, "[NONE]")

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

            'find lloq units
            dv = frmH.dgvWatsonAnalRef.DataSource
            int2 = FindRowDV("LLOQ Units", dv)
            strUnits = dv(int2).Item(1)

            int1 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
            str1 = NZ(frmH.dgvStudyConfig(1, int1).Value, "")

            If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
            Else
                strUnits = str1
            End If

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
            '    strM = "Creating Spiking Solution Stability Assessment ...."
            '    frmH.lblProgress.Text = strM
            '    frmH.Refresh()
            '    MsgBox("Samples have not been assigned to this table.", MsgBoxStyle.Information, "Samples have not been assigned...")
            '    GoTo end2
            'End If

            strF = "IsIntStd = 'No'"
            'strS = "AnalyteDescription ASC"
            strS = "INTORDER ASC, IsIntStd ASC, AnalyteDescription ASC"
            rows11 = tblAnalytesHome.Select(strF, strS)
            intRowsAnal = rows11.Length

            If tblX.Columns.Contains("Ratio") Then
            Else
                tblX.Columns.Add("Ratio", Type.GetType("System.Double"))
                tblX.Columns.Add("ANALYTEAREA", Type.GetType("System.Double"))
                tblX.Columns.Add("INTERNALSTANDARDAREA", Type.GetType("System.Double"))
            End If

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

                strX = rows11(Count1 - 1).Item("IsIntStd")
                boolX = False
                boolJustTable = False
                If StrComp(strX, "Yes", CompareMethod.Text) = 0 Then
                    'check for boolIntStd in tbl2
                    strF = "IsIntStd = 'Yes'"
                    var1 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteIndex")
                    var2 = tbl4.Rows.Item(Count1 - 1).Item("MasterAssayID")
                    var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                    strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                    strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                    strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                    strF2 = strF2 & "BOOLINTSTD = -1"
                    Erase rowsX
                    rowsX = tbl2.Select(strF2)
                    int1 = rowsX.Length
                    If int1 > 0 Then
                        bool = True
                        boolX = True
                    Else
                        bool = False
                    End If
                Else
                    bool = dvDo.Item(intDo).Item(strDo) 'find boolean value of dvDo column
                End If

                Dim strM1 As String
                If bool Then 'continue
                    'ensure data has been entered
                    strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & CleanText(strDo) & "' AND ID_TBLREPORTTABLE = " & idTR
                    rowsX = tbl2.Select(strF)

                    intTCur = intTCur + 1

                    'setup tables
                    If boolUseGroups Then
                        If boolX Then
                            intGroup = -1 'tblAG.Rows(Count1 - 1).Item("INTGROUP")
                            strAnal = tbl4.Rows.Item(Count1 - 1).Item("OriginalAnalyteDescription") ' tblAG.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                            strAnalC = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription") ' tblAG.Rows(Count1 - 1).Item("ANALYTEDESCRIPTION_C")
                            vAnalyteID = 0 ' tblAG.Rows.Item(Count1 - 1).Item("ANALYTEID")

                            'find matrix
                            Dim strFF As String
                            strFF = "INTSTD = '" & CleanText(strAnal) & "'"
                            Dim rowsFF() As DataRow = tblAG.Select(strFF, "", DataViewRowState.CurrentRows)
                            If rowsFF.Length = 0 Then
                                strMatrix = "NA"
                            Else
                                strMatrix = NZ(rowsFF(0).Item("MATRIX"), "NA")
                            End If

                        Else
                            intGroup = tblAG.Rows(Count1 - 1).Item("INTGROUP")
                            strAnal = tblAG.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                            strAnalC = tblAG.Rows(Count1 - 1).Item("ANALYTEDESCRIPTION_C")
                            vAnalyteID = tblAG.Rows.Item(Count1 - 1).Item("ANALYTEID")
                            strMatrix = tblAG.Rows(Count1 - 1).Item("MATRIX")
                        End If

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
                        'page setup according to configuration
                        str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                        'insert page break
                        'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                        Call InsertPageBreak(wd)
                        Call PageSetup(wd, str1) 'L=Landscape, P=Portrait
                        boolJustTable = True
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

                    ''setup tables
                    'If boolUseGroups Then
                    '    If boolX Then
                    '        intGroup = -1 'tblAG.Rows(Count1 - 1).Item("INTGROUP")
                    '        strAnal = tbl4.Rows.Item(Count1 - 1).Item("OriginalAnalyteDescription") ' tblAG.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                    '        strAnalC = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription") ' tblAG.Rows(Count1 - 1).Item("ANALYTEDESCRIPTION_C")
                    '        vAnalyteID = 0 ' tblAG.Rows.Item(Count1 - 1).Item("ANALYTEID")
                    '        strMatrix = "" 'tblAG.Rows(Count1 - 1).Item("MATRIX")
                    '    Else
                    '        intGroup = tblAG.Rows(Count1 - 1).Item("INTGROUP")
                    '        strAnal = tblAG.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                    '        strAnalC = tblAG.Rows(Count1 - 1).Item("ANALYTEDESCRIPTION_C")
                    '        vAnalyteID = tblAG.Rows.Item(Count1 - 1).Item("ANALYTEID")
                    '        strMatrix = tblAG.Rows(Count1 - 1).Item("MATRIX")
                    '    End If

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

                    'find outliers
                    strF2 = ""
                    If boolUseGroups Then
                        strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                        strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                        strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                        strF2 = strF2 & "INTGROUP = " & intGroup & " AND (ELIMINATEDFLAG = 'Y' OR BOOLEXCLSAMPLE = -1)"
                    Else
                        var1 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteIndex")
                        var2 = tbl4.Rows.Item(Count1 - 1).Item("MasterAssayID")
                        var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                        strF2 = strF2 & "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                        strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                        strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                        strF2 = strF2 & "ANALYTEINDEX = " & var1 & " AND "
                        strF2 = strF2 & "MASTERASSAYID = " & var2 & " AND (ELIMINATEDFLAG = 'Y' OR BOOLEXCLSAMPLE = -1)"
                    End If

                    strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                    rows2 = tbl2.Select(strF2, strS)
                    int1 = rows2.Length 'debug
                    If int1 > 0 Then
                        boolOutlier = True
                    End If

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
                    Erase rows2
                    rows2 = tbl2.Select(strF2, strS)
                    int1 = rows2.Length 'debug
                    dv2 = New DataView(tbl2, strF2, strS, DataViewRowState.CurrentRows)
                    int1 = dv2.Count 'debug
                    strS = "CONCENTRATION ASC"
                    strF = "ANALYTEID = " & vAnalyteID ' "ANALYTEINDEX = " & var1 & " AND MASTERASSAYID = " & var2
                    rowStd = tblBCStds.Select(strF, strS)
                    intStdRows = rowStd.Length

                    'find number of runs used
                    tblNumRuns = dv2.ToTable("a", True, "RUNID")
                    intNumRuns = tblNumRuns.Rows.Count
                    'intNumRuns = 1


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


                    'establish number of QCs evaluated
                    'this will actually give number of columns
                    'must be sorted by nomconc!
                    'make new dv
                    Dim dvNL As New DataView(tbl2, strF2, "NOMCONC ASC", DataViewRowState.CurrentRows)
                    tblLevels = dvNL.ToTable("b", True, "NOMCONC")
                    intNumLevels = tblLevels.Rows.Count
                    For Count2 = 0 To intNumLevels - 1 'check for any null values
                        var3 = tblLevels.Rows.Item(Count2).Item("NOMCONC")
                        If IsDBNull(var3) Then
                            str1 = "The Nominal Concentration for some assigned samples for " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & " have not been configured."
                            str1 = str1 & ChrW(10) & "When this action is finished, please navigate to the Assigned Samples window and correct this problem."
                            If boolDisableWarnings Then
                            Else
                                MsgBox(str1, MsgBoxStyle.Information, "Nom Conc problem...")
                            End If
                            GoTo end1
                        End If
                    Next

                    ReDim arr1(intNumLevels)

                    'find number of table rows to generate
                    intRowsX = intNumLevels

                    'find introwsQC and dataviews CHARHELPER1
                    If boolX Then
                        strF = " AND CHARHELPER1 = 'Old Spiking Solution' AND BOOLINTSTD = -1"
                    Else
                        strF = " AND CHARHELPER1 = 'Old Spiking Solution' AND BOOLINTSTD = 0"
                    End If
                    strF = strF2 & strF
                    strS = "NOMCONC ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                    dv1 = New DataView(tbl2, strF, strS, DataViewRowState.CurrentRows)
                    tbl2a = dv1.ToTable
                    introwsQC = tbl2a.Rows.Count
                    If introwsQC = 0 Then
                        str1 = "Appropriate assigned samples for " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & " have not been configured as 'Old Spiking Solution'."
                        str1 = str1 & ChrW(10) & "When this action is finished, please navigate to the Assigned Samples window and correct this problem."
                        MsgBox(str1, MsgBoxStyle.Information, "Nom Conc problem...")
                        GoTo end1
                    End If

                    'find introwsRS and dataviews CHARHELPER1
                    If boolX Then
                        strF = " AND CHARHELPER1 = 'New Spiking Solution' AND BOOLINTSTD = -1"
                    Else
                        strF = " AND CHARHELPER1 = 'New Spiking Solution' AND BOOLINTSTD = 0"
                    End If
                    strF = strF2 & strF
                    strS = "NOMCONC ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                    dv3 = New DataView(tbl2, strF, strS, DataViewRowState.CurrentRows)
                    tbl2b = dv3.ToTable
                    introwsRS = tbl2b.Rows.Count
                    If introwsRS = 0 Then
                        str1 = "Appropriate assigned samples for " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & " have not been configured as 'New Spiking Solution'."
                        str1 = str1 & ChrW(10) & "When this action is finished, please navigate to the Assigned Samples window and correct this problem."
                        MsgBox(str1, MsgBoxStyle.Information, "Nom Conc problem...")
                        GoTo end1
                    End If

                    Dim intRowsP As Short
                    If introwsQC > introwsRS Then
                        intRowsP = introwsQC / intNumLevels
                    Else
                        intRowsP = introwsRS / intNumLevels
                    End If

                    'generate table
                    intTblRows = 0
                    intTblRows = intTblRows + 2 'for header
                    intTblRows = intTblRows + 1 'for blank row
                    intTblRows = intTblRows + introwsRS + introwsQC
                    'intTblRows = intTblRows + 10 'for Old/New set

                    'Increment for Statistics Sections
                    Dim intCSN As Short
                    intCSN = countNumStatsRows()
                    intTblRows = intTblRows + (2 * intNumRuns * intCSN)

                    If intCSN > 0 Then
                    Else
                        intTblRows = intTblRows - 1 'subtract an unneeded blank row
                    End If

                    intTblRows = intTblRows + (2 * intNumRuns) 'for blank row
                    'intTblRows = intTblRows + 1 'for %Difference
                    intTblRows = intTblRows + 3 'for %Difference


                    If boolQCREPORTACCVALUES Then

                    Else
                        If boolOutlier Then
                            intTblRows = intTblRows + (3 * intNumRuns) 'for stats headings
                            ctExp = ctExp + 3 'for stats headings

                            'Increment for Statistics Sections
                            intTblRows = intTblRows + (intCSN * 2 * intNumRuns)
                            ctExp = ctExp + (intCSN * 2 * intNumRuns)

                            If intCSN > 0 Then
                                intTblRows = intTblRows + (2 * intNumRuns) - 1 'for a blank row after each Mean/Bias/n set, except last set
                                ctExp = ctExp + (2 * intNumRuns) - 1
                            End If
                       
                        End If


                    End If

                    'intTblRows = intTblRows + 2 'for footer

                    wrdSelection = wd.Selection()

                    'intCols = 2 + intNumLevels
                    intCols = 1 + (3 * intNumLevels)


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


                        'remove border, but leave top and bottom
                        removeBorderButLeaveTopAndBottom(wd)

                        'border top and bottom of range
                        .Selection.Tables.Item(1).Cell(1, 1).Select()

                        If boolPlaceHolder Then

                            .Selection.Tables.Item(1).Select()
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone

                            strA = arrAnalytes(14, Count1)

                            If boolX Then
                                strA = "Internal Standard " & strA
                            End If

                            strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, boolX)
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

                        .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.MoveRight Unit:=Microsoft.Office.Interop.Word.wdunits.word.wdunits.wdCharacter, Count:=ctQCs, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        'format selection
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                        .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

                        'Enter  row titles
                        'column 1
                        .Selection.Tables.Item(1).Cell(2, 1).Select()
                        str1 = strWRunId & tblNumRuns.Rows.Item(0).Item("RUNID")
                        .Selection.TypeText(Text:=str1)

                        int1 = 4 + intRowsP + 1 'row position counter

                        int8 = -1
                        If boolQCREPORTACCVALUES Then
                        Else
                            If boolOutlier Then
                                int8 = int8 + 1
                                .Selection.Tables.Item(1).Cell(int1 + int8, 2).Select()
                                .Selection.TypeText(Text:="Summary Statistics Excluding Outlier Values")
                                .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                Try
                                    .Selection.Cells.Merge()
                                Catch ex As Exception
                                End Try
                                With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                    .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                End With
                                .Selection.Tables.Item(1).Cell(int1 + 2, 1).Select()
                            End If
                        End If


                        Dim intHome1 As Short
                        Dim intHome2 As Short
                        Dim intT8 As Short
                        intT8 = int8
                        Dim intCCol As Short

                        'column 2
                        For Count2 = 1 To intNumLevels
                            intCCol = ((Count2 - 1) * 3) + 2
                            intHome1 = int1
                            int8 = intT8

                            'enter Peak Area headings
                            str1 = "Analyte"
                            .Selection.Tables.Item(1).Cell(1, intCCol).Select()
                            .Selection.TypeText(Text:=str1)
                            str1 = "Peak" & ChrW(160) & "Area"
                            .Selection.Tables.Item(1).Cell(2, intCCol).Select()
                            .Selection.TypeText(Text:=str1)

                            str1 = "Internal" & ChrW(160) & "Standard"
                            .Selection.Tables.Item(1).Cell(1, intCCol + 1).Select()
                            .Selection.TypeText(Text:=str1)
                            str1 = "Peak" & ChrW(160) & "Area"
                            .Selection.Tables.Item(1).Cell(2, intCCol + 1).Select()
                            .Selection.TypeText(Text:=str1)

                            typeStatsLabels(wd, int8, int1, intCCol, False)

                            If boolQCREPORTACCVALUES Then
                            Else
                                If boolOutlier Then
                                    int8 = int8 + 2
                                    .Selection.Tables.Item(1).Cell(int1 + int8, 2).Select()
                                    .Selection.TypeText(Text:="Summary Statistics Including Outlier Values")
                                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                    Try
                                        .Selection.Cells.Merge()
                                    Catch ex As Exception
                                    End Try
                                    With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                        .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                    End With
                                    .Selection.Tables.Item(1).Cell(int1 + 2, 1).Select()
                                End If
                            End If

                            int8 = int8 + 1 'for blank line
                            int8 = int8 + intRowsP
                            int8 = int8 + 1 'for blank line

                            intHome2 = int8 + 1

                            typeStatsLabels(wd, int8, int1, intCCol, False)
                        Next


                        '.Selection.Tables.item(1).Cell(int1 + 10, 2).Select()
                        'str1 = "%Difference"
                        '.Selection.TypeText(Text:=str1)

                        'start doing columns of stdconcs (nomconc)

                        'begin entering data'

                        For Count2 = 0 To intNumLevels - 1

                            intCCol = (Count2 * 3) + 2
                            'for legend stuff
                            intLeg = 0
                            ctQCLegend = 0
                            ctDilLeg = 0
                            ctLegend = 0
                            strA = ""
                            strB = ""
                            arrLegend.Clear(arrLegend, 0, arrLegend.Length)
                            arrFP.Clear(arrFP, 0, arrFP.Length)
                            intLegStart = 96

                            'reset starting position
                            int1 = 4

                            'retrieve nomconc
                            varNomConc = tblLevels.Rows.Item(Count2).Item("NOMCONC")

                            'find Std number
                            numStd = 0
                            For Count3 = 0 To intStdRows - 1
                                var1 = rowStd(Count3).Item("CONCENTRATION")
                                var1 = NZ(var1, 0)
                                If CDec(var1) = CDec(varNomConc) Then
                                    numStd = Count3 + 1
                                    Exit For
                                End If
                            Next

                            'retrive runid
                            var10 = tblNumRuns.Rows.Item(0).Item("RUNID")  'NDL: This assumes all values are from same run!!!  If they aren't, report will crash.

                            'enter column headers
                            .Selection.Tables.Item(1).Cell(1, intCCol + 2).Select()
                            If boolRCPA Then
                                str1 = "Peak Area" & ChrW(10) & "Recovery Std " & numStd
                            Else
                                str1 = "Peak Area Ratio" & ChrW(10) & "Recovery Std " & numStd
                            End If

                            str1 = Replace(str1, " ", ChrW(160), 1, -1, CompareMethod.Text) 'replace all spaces with nbs
                            .Selection.TypeText(Text:=str1)

                            .Selection.Tables.Item(1).Cell(2, intCCol + 2).Select()

                            If LboolNomConcParen Then
                                If boolLUseSigFigs Then
                                    str1 = "(" & DisplayNum(SigFigOrDec(varNomConc, LSigFig, False), LSigFig, False) & ChrW(160) & strConcUnits & ")"
                                Else
                                    str1 = "(" & Format(RoundToDecimalRAFZ(varNomConc, LSigFig), GetRegrDecStr(LSigFig)) & ChrW(160) & strConcUnits & ")"
                                End If

                            Else
                                If boolLUseSigFigs Then
                                    str1 = DisplayNum(SigFigOrDec(varNomConc, LSigFig, False), LSigFig, False) & ChrW(160) & strConcUnits
                                Else
                                    str1 = Format(RoundToDecimalRAFZ(varNomConc, LSigFig), GetRegrDecStr(LSigFig)) & ChrW(160) & strConcUnits
                                End If

                            End If

                            .Selection.TypeText(Text:=str1)

                            'strM = "Creating Spiking Solution Assessment For " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & "..."
                            frmH.lblProgress.Text = strM1 & ChrW(10) & "Processing Run ID " & var10
                            frmH.Refresh()

                            'start filling in data by rows
                            'intRowsX = 0
                            For Count3 = 0 To 1

                                'column 3
                                strF = "NOMCONC = " & varNomConc & " AND RUNID = " & var10
                                If Count3 = 0 Then 'get Old Spiking Solution
                                    Erase rows2a
                                    rows2a = tbl2a.Select(strF)
                                Else 'get new Spiking solution
                                    Erase rows2a
                                    rows2a = tbl2b.Select(strF)
                                End If
                                int2 = rows2a.Length
                                strF = ""

                                Erase rows1
                                'rows1 = tbl1.Select(strF)
                                rows1 = rows2a
                                int3 = rows1.Length

                                int8 = -1

                                'get area ratios
                                tblX.Clear()
                                For Count4 = 0 To int3 - 1
                                    var1 = RoundToDecimalRAFZ(rows1(Count4).Item("ANALYTEAREA"), 0)
                                    var2 = RoundToDecimalRAFZ(NZ(rows1(Count4).Item("INTERNALSTANDARDAREA"), 0), 0)
                                    If var2 = 0 Then
                                        var3 = var1
                                    Else
                                        var3 = var1 / var2
                                        var3 = RoundToDecimalRAFZ(var3, 5)
                                    End If
                                    Dim nr1 As DataRow = tblX.NewRow
                                    nr1.BeginEdit()
                                    nr1.Item("Ratio") = var3
                                    nr1.Item("ANALYTEAREA") = var1
                                    nr1.Item("INTERNALSTANDARDAREA") = var2
                                    nr1.EndEdit()
                                    tblX.Rows.Add(nr1)

                                    int8 = int8 + 1
                                    'enter peak areas
                                    .Selection.Tables.Item(1).Cell(int1 + int8, intCCol).Select()
                                    .Selection.TypeText(var1)
                                    .Selection.Tables.Item(1).Cell(int1 + int8, intCCol + 1).Select()
                                    .Selection.TypeText(var2)

                                    .Selection.Tables.Item(1).Cell(int1 + int8, intCCol + 2).Select()
                                    If boolRCPA Then
                                        '.Selection.TypeText(Format(var1, "0"))
                                        .Selection.TypeText(Format(var1, strAreaDecAreaRatio))
                                    Else
                                        .Selection.TypeText(Format(RoundToDecimalRAFZ(var3, 5), "0.00000"))
                                    End If

                                Next
                                strF = "Ratio > 0"
                                Erase rows2b
                                rows2b = tblX.Select(strF)

                                int8 = int8 + 1 'blank line
                                'int8 = -1

                                If boolSTATSMEAN Then
                                    Try
                                        'Row1 Mean
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + int8, intCCol + 2).Select()
                                        If boolRCPA Then
                                            var1 = MeanDR(rows2b, "ANALYTEAREA", False, "gaga", False, False)
                                            If boolLUseSigFigsArea Then
                                                .Selection.TypeText(CStr(DisplayNum(SigFigArea(var1, LSigFigArea, True, False), LSigFigArea, False)))
                                                var2 = SigFigOrDec(var1, LSigFigArea, False)
                                            Else
                                                .Selection.TypeText(Format(RoundToDecimalRAFZ(var1, LSigFigArea), GetRegrDecStr(LSigFigArea)))
                                                var2 = RoundToDecimalRAFZ(var1, LSigFigArea)
                                            End If
                                        Else
                                            var1 = MeanDR(rows2b, "Ratio", False, "gaga", False, False)

                                            If boolLUseSigFigsAreaRatio Then
                                                .Selection.TypeText(CStr(DisplayNum(SigFigAreaRatio(var1, LSigFigAreaRatio, True, False), LSigFigAreaRatio, False)))
                                                var2 = SigFigOrDec(var1, LSigFigAreaRatio, False)
                                            Else
                                                .Selection.TypeText(Format(RoundToDecimalRAFZ(var1, LSigFigAreaRatio), GetRegrDecStr(LSigFigAreaRatio)))
                                                var2 = RoundToDecimalRAFZ(var1, LSigFigAreaRatio)
                                            End If

                                        End If

                                        numMean = var2
                                        numMeanFP = CDec(var1)

                                        If Count3 = 0 Then
                                            numA = numMean
                                            numAFP = numMeanFP
                                        Else
                                            numB = numMean
                                            numBFP = numMeanFP
                                        End If
                                    Catch ex As Exception

                                    End Try
                                End If
                                If boolSTATSSD Then
                                    Try
                                        'Row2 SD
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + int8, intCCol + 2).Select()
                                        If rows2b.Length < gSDMax Then
                                            .Selection.TypeText("NA")
                                            var3 = 0
                                        Else
                                            If boolRCPA Then
                                                var1 = StdDevDR(rows2b, "ANALYTEAREA", False, "gaga", False, False)
                                                If boolLUseSigFigsArea Then
                                                    .Selection.TypeText(DisplayNum(SigFigArea(var1, LSigFigArea, True, False), LSigFigArea, False))
                                                Else
                                                    .Selection.TypeText(Format(RoundToDecimalRAFZ(var1, LSigFigArea), GetRegrDecStr(LSigFigArea)))
                                                End If
                                                'var2 = RoundToDecimal(var1, 0)
                                                'var3 = var2 'SigFigOrDec(var2, LSigFig, False)
                                                '.Selection.TypeText(CStr(var3))
                                            Else
                                                var1 = StdDevDR(rows2b, "Ratio", False, "gaga", False, False)
                                                If boolLUseSigFigsAreaRatio Then
                                                    .Selection.TypeText(DisplayNum(SigFigAreaRatio(var1, LSigFigAreaRatio, True, False), LSigFigAreaRatio, False))
                                                Else
                                                    .Selection.TypeText(Format(RoundToDecimalRAFZ(var1, LSigFigAreaRatio), GetRegrDecStr(LSigFigAreaRatio)))
                                                End If
                                                'var2 = RoundToDecimal(var1, 8)
                                                'var3 = SigFigOrDec(var2, LSigFig, False)
                                                '.Selection.TypeText(CStr(DisplayNum(var3, LSigFig, False)))
                                            End If

                                        End If

                                        numSD = var3
                                    Catch ex As Exception

                                    End Try
                                End If


                                Try
                                    If rows2b.Length < gSDMax Then
                                    Else
                                        numPrec = CalcCVPercent(numSD, numMean, intQCDec)
                                        Call InsertQCTables(intTableID, idTR, charFCID, varNom, Count3, "Precision", numPrec, CSng(var10), Count1, strDo, 0, 0, False)
                                    End If

                                Catch ex As Exception

                                End Try
                                If boolSTATSCV Then
                                    Try
                                        'Row3 %CV
                                        int8 = int8 + 1
                                        If rows2b.Length < gSDMax Then
                                            var3 = "NA"
                                            .Selection.Tables.Item(1).Cell(int1 + int8, intCCol + 2).Select()
                                            .Selection.TypeText(CStr(var3))
                                        Else
                                            .Selection.Tables.Item(1).Cell(int1 + int8, intCCol + 2).Select()
                                            .Selection.TypeText(Format(numPrec, strQCDec))
                                        End If


                                    Catch ex As Exception

                                    End Try
                                End If


                                If boolSTATSBIAS And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                    Try
                                        numBias = CalcREPercent(numMean, varNom, intQCDec)
                                        If rows2b.Length = 0 Then
                                        Else
                                            Call InsertQCTables(intTableID, idTR, charFCID, varNom, Count3, "Accuracy", numBias, CSng(var10), Count1, strDo, 0, 0, False)
                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If
                                If boolSTATSBIAS And boolSTATSMEAN Then
                                    Try
                                        'Row3 %CV
                                        int8 = int8 + 1

                                        .Selection.Tables.Item(1).Cell(int1 + int8, intCCol + 2).Select()
                                        .Selection.TypeText(Format(numBias, strQCDec))

                                    Catch ex As Exception

                                    End Try
                                End If


                                If boolTHEORETICAL And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then

                                    Try
                                        numTheor = CalcREPercent(numMean, varNom, intQCDec)
                                        numTheor = 100 + CDec(numTheor)
                                        If rows2b.Length = 0 Then
                                        Else
                                            Call InsertQCTables(intTableID, idTR, charFCID, varNom, Count3, "Accuracy", numTheor, CSng(var10), Count1, strDo, 0, 0, False)
                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If
                                If boolTHEORETICAL And boolSTATSMEAN Then
                                    Try
                                        'enter %theoretical
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + int8, intCCol + 2).Select()
                                        .Selection.TypeText(Format(numTheor, strQCDec))

                                    Catch ex As Exception

                                    End Try

                                End If



                                If boolSTATSDIFF And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                    Try
                                        numBias = CalcREPercent(numMean, varNom, intQCDec)
                                        If rows2b.Length = 0 Then
                                        Else
                                            Call InsertQCTables(intTableID, idTR, charFCID, varNom, Count3, "Accuracy", numBias, CSng(var10), Count1, strDo, 0, 0, False)
                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If
                                If boolSTATSDIFF And boolSTATSMEAN Then
                                    Try
                                        'Row3 %RE
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + int8, intCCol + 2).Select()
                                        .Selection.TypeText(Format(numBias, strQCDec))
                                    Catch ex As Exception

                                    End Try
                                End If



                                If BOOLSTATSRE And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                    Try
                                        numBias = CalcREPercent(numMean, varNom, intQCDec)
                                        If rows2b.Length = 0 Then
                                        Else
                                            Call InsertQCTables(intTableID, idTR, charFCID, varNom, Count3, "Accuracy", numBias, CSng(var10), Count1, strDo, 0, 0, False)
                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If
                                If BOOLSTATSRE And boolSTATSMEAN Then
                                    Try
                                        'Row3 %RE
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + int8, intCCol + 2).Select()
                                        .Selection.TypeText(Format(numBias, strQCDec))

                                    Catch ex As Exception

                                    End Try
                                End If


                                Try
                                    Call InsertQCTables(intTableID, idTR, charFCID, varNom, Count3, "n", CInt(var3), CSng(var10), Count1, strDo, 0, 0, False)
                                Catch ex As Exception

                                End Try
                                If boolSTATSN Then
                                    Try
                                        'Row4 n
                                        int8 = int8 + 1
                                        var3 = rows2b.Length
                                        .Selection.Tables.Item(1).Cell(int1 + int8, intCCol + 2).Select()
                                        .Selection.TypeText(CStr(var3))
                                    Catch ex As Exception

                                    End Try
                                End If

                                'Row1 Column 1
                                .Selection.Tables.Item(1).Cell(int1, 1).Select()
                                If Count3 = 0 Then
                                    var1 = tbl2a.Rows.Item(0).Item("CHARHELPER1")
                                Else
                                    var1 = tbl2b.Rows.Item(0).Item("CHARHELPER1")
                                End If
                                str1 = Replace(NZ(var1, "NA"), " ", ChrW(160), 1, -1, CompareMethod.Text)
                                .Selection.TypeText(CStr(str1))


                                If Count3 = 0 Then
                                    int1 = int1 + int8 + 2 '5
                                Else
                                    int1 = int1 + int8 + 2 '5
                                End If

                            Next

                            'do %Difference/Mean Acc
                            Try

                                'Note: The equations below are backwards
                                'at this point, numA and numB are full precision
                                'numA = Old, numB = New
                                '20161206 LEE: gboolMeanFullPrec has been deprecated. Only gboolMeanRounded = true (e.g. gboolMeanFullPrec = false) is allowed
                                If gboolMeanFullPrec Then

                                    If boolMEANACCURACY Then
                                        If boolPOSLEG Then
                                            var1 = (numAFP - numBFP) / ((numAFP + numBFP) / 2) * 100
                                        Else
                                            var1 = (numBFP - numAFP) / ((numAFP + numBFP) / 2) * 100
                                        End If
                                    ElseIf BOOLDIFFERENCE Then
                                        If boolPOSLEG Then
                                            var1 = RoundToDecimalRAFZ((numAFP - numBFP) / numAFP * 100, 5) '(T0-Tn)/((Tn+T0)/2)
                                        Else
                                            var1 = RoundToDecimalRAFZ((numBFP - numAFP) / numAFP * 100, 5) '(Tn-T0)/((Tn+T0)/2)
                                        End If
                                    ElseIf boolRECOVERY Then
                                        If boolPOSLEG Then
                                            var1 = RoundToDecimalRAFZ(numAFP / numBFP * 100, 5) 'T0/Tn
                                        Else
                                            var1 = RoundToDecimalRAFZ(numBFP / numAFP * 100, 5) 'Tn/T0
                                        End If
                                    Else
                                        If boolPOSLEG Then
                                            var1 = ((numAFP / numBFP) - 1) * 100
                                        Else
                                            var1 = (1 - (numAFP / numBFP)) * 100
                                        End If
                                    End If

                                    var2 = var1 ' RoundToDecimal(var1, 5)

                                Else

                                    ''legend
                                    'If boolLUseSigFigsArea Then
                                    '    .Selection.TypeText(CStr(DisplayNum(SigFigOrDec(var1, LSigFig, False), LSigFigArea, False)))
                                    'Else
                                    '    .Selection.TypeText(Format(RoundToDecimalRAFZ(var1, LSigFigArea), GetRegrDecStr(LSigFigArea)))
                                    'End If

                                    If boolRCPA Then
                                        If boolLUseSigFigsArea Then
                                            numA1 = SigFigArea(RoundToDecimalA(numA, LSigFigArea), LSigFigArea, True, False)
                                        Else
                                            numA1 = RoundToDecimalRAFZ(numA, LSigFigArea)
                                        End If
                                    Else 'peak area ratio
                                        'numA1 = numA
                                        If boolLUseSigFigsAreaRatio Then
                                            numA1 = SigFigAreaRatio(RoundToDecimalA(numA, LSigFigAreaRatio), LSigFigAreaRatio, True, False)
                                        Else
                                            numA1 = RoundToDecimalRAFZ(numA, LSigFigAreaRatio)
                                        End If

                                        'numB1 = numB
                                        If boolLUseSigFigsAreaRatio Then
                                            numB1 = SigFigAreaRatio(RoundToDecimalA(numB, LSigFigAreaRatio), LSigFigAreaRatio, True, False)
                                        Else
                                            numB1 = RoundToDecimalRAFZ(numB, LSigFigAreaRatio)
                                        End If
                                    End If

                                    var1 = ReturnDiff(numA1, numB1)

                                    'If boolMEANACCURACY Then
                                    '    If boolPOSLEG Then
                                    '        var1 = (numA1 - numB1) / ((numA1 + numB1) / 2) * 100
                                    '    Else
                                    '        var1 = (numB1 - numA1) / ((numA1 + numB1) / 2) * 100
                                    '    End If
                                    'ElseIf BOOLDIFFERENCE Then
                                    '    If boolPOSLEG Then
                                    '        var1 = RoundToDecimalRAFZ((numA1 - numB1) / numA1 * 100, 5) '(Tn-TO)/((Tn+T0)/2)
                                    '    Else
                                    '        var1 = RoundToDecimalRAFZ((numB1 - numA1) / numA1 * 100, 5) '(TO-Tn)/((Tn+T0)/2)
                                    '    End If
                                    'ElseIf boolRECOVERY Then
                                    '    If boolPOSLEG Then
                                    '        var1 = RoundToDecimalRAFZ(numA1 / numB1 * 100, 5) 'Tn/T0
                                    '    Else
                                    '        var1 = RoundToDecimalRAFZ(numB1 / numA1 * 100, 5) 'T0/Tn
                                    '    End If
                                    'Else
                                    '    If boolPOSLEG Then
                                    '        var1 = ((numA1 / numB1) - 1) * 100
                                    '    Else
                                    '        var1 = (1 - (numA1 / numB1)) * 100
                                    '    End If
                                    'End If

                                    var2 = var1 ' RoundToDecimal(var1, 5)

                                End If

                                var3 = ""
                            Catch ex As Exception
                                var3 = "NA"
                            End Try

                            If Len(var3) = 0 Then
                                var3 = Format(RoundToDecimalRAFZ(var2, intQCDec), strQCDec)
                            End If

                            .Selection.Tables.Item(1).Cell(int1, intCCol + 2).Select()
                            .Selection.TypeText(CStr(var3))
                            Call InsertQCTables(intTableID, idTR, charFCID, varNom, Count2 + 1, "Diff", var3, CSng(var10), Count1, strDo, 0, 0, False)

                        Next

                        'bottom border this row
                        .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.MoveRight Unit:=Microsoft.Office.Interop.Word.wdunits.word.wdunits.wdCharacter, Count:=ctQCs, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                        'If boolQCREPORTACCVALUES Then
                        'Else
                        '    If intExp = 0 Then
                        '        Call DeleteRows(ctExp, wd)
                        '    End If
                        'End If


                    Catch ex As Exception

                        str1 = "There was a problem preparing table:"
                        str1 = strM1 & ChrW(10) & ChrW(10) & str1
                        str1 = str1 & ChrW(10) & ChrW(10)
                        str1 = str1 & ex.Message
                        MsgBox(str1, vbInformation, "Problem...")

                    End Try


                    'autofit window
                    .Selection.Tables.Item(1).Select()
                    'autofit table
                    Call AutoFitTable(wd, False)

                    ''enter Difference information ' NO! Use InsertLegend instead
                    'add a line space
                    .Selection.Tables.Item(1).Cell(int1, 2).Select()

                    str1 = GetLegendTitle(intTableID, idTR) ' "%Difference"

                    '20190225 LEE:
                    Dim strDiff As String = ""
                    strDiff = str1

                    'if there is an '=' sign, then remove it
                    str1 = Replace(str1, " = ", "", 1, -1, CompareMethod.Text)
                    str1 = Replace(str1, "= ", "", 1, -1, CompareMethod.Text)
                    str1 = Replace(str1, " =", "", 1, -1, CompareMethod.Text)

                    'str1 = "%Difference"
                    .Selection.TypeText(Text:=str1)
                    'ctLegend = ctLegend + 1
                    'intLeg = intLeg + 1
                    'strA = ChrW(intLeg + intLegStart)
                    '.selection.font.superscript = True
                    '.Selection.TypeText(Text:=strA)
                    'str1 = "%Difference = ((Mean Old - Mean New)/Mean New) x 100"
                    'arrLegend(1, intLeg) = strA
                    'arrLegend(2, intLeg) = str1
                    'arrLegend(3, intLeg) = True

                    'enter table number
                    var1 = strTempInfo
                    'replace numeric with verbose in strTempInfo
                    'look for a sequence of characters that is numeric
                    var3 = ""
                    var4 = ""
                    Dim bool1 As Boolean
                    Dim bool2 As Boolean
                    bool1 = False 'Start
                    bool2 = False 'End
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

                    'remove unused rows
                    Call RemoveRows(wd, 0)

                    str1 = str2 & " Spiking Solution Stability Assessment: Summary of " & rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")

                    '***
                    strA = rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")
                    If boolX Then
                        strA = "Internal Standard " & strA
                    End If
                    strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, boolX)
                    Call EnterTableNumber(wd, strTName, 3, strA, strTempInfo, intTableID, intGroup, idTR)
                    '***

                    'Call EnterTableNumber(wd, str1, 3)

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

                    .Selection.Tables.Item(1).Select()
                    'autofit table
                    Call AutoFitTable(wd, BOOLINCLUDEDATE)

                    strM = "Finalizing " & strTName & "..."
                    strM1 = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    str1 = strM1

                    ''split table, if needed
                    frmH.lblProgress.Text = strM1
                    frmH.Refresh()


                    Call SplitTable(wd, 4, intLeg, arrLegend, str1, True, ctLegend + 2, False, False, False, intTableID)
                    'Sub SplitTable(ByVal wd As Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, 
                    'ByVal arr As Object, ByVal strT As String, ByVal DoLegend As Boolean, ByVal intSplitRows As Short, ByVal boolSmallFont As Boolean)

                    Call MoveOneCellDown(wd)

                    Call InsertLegend(wd, intTableID, idTR, boolX, 1)

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
                    'If bool1 = False Then
                    '    var2 = "[NA]"
                    'Else
                    '    var2 = VerboseNumber(var4, True)
                    '    str2 = Replace(var1, CStr(var4), var2, 1, 1, CompareMethod.Text)
                    'End If

                    'str1 = str2 & " Spiking Solution Stability Assessment: Summary of " & rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")
                    'str2 = str1
                    'str1 = rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")
                    ''Call JustTable(wd, str1, str2, strDo, strTName, intTableID)
                    'Call JustTable(wd, str1, strTName, strDo, strTName, intTableID, strTempInfo, "")

                    str1 = NZ(rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION"), "")
                    'Call JustTable(wd, str1, str2, strDo, strTName, intTableID)
                    If Len(str1) = 0 Then
                    Else
                        strA = strAnal
                        strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, boolX)
                        Call JustTable(wd, str1, strTName, strDo, strTName, intTableID, strTempInfo, "", strTNameO, intGroup, idTR)
                    End If

                End If

next1:

            Next
end2:
        End With

    End Sub


    Sub MVSummaryLongTermQCStability_29(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal idTR As Int64)

        '20180807 LEE:
        'deprecate this table
        'Use Ad Hoc Stability Comparison instead

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
        Dim tbl2a As New System.Data.DataTable
        Dim tbl2b As New System.Data.DataTable
        Dim dv2 As system.data.dataview
        Dim rows2() As DataRow
        Dim rows2a() As DataRow
        Dim rows2b() As DataRow
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
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim strF As String
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim int3a As Short
        Dim int4 As Short
        Dim int10 As Short
        Dim intRowsX As Short
        Dim tblX As New System.Data.DataTable
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
        Dim strMsg As String = ""

        Dim varConc

        Dim hi, lo
        Dim rows10() As DataRow
        Dim rows11() As DataRow
        Dim intRowsAnal As Short
        Dim arrFP(2, 20) 'FlagPercent array
        '1=max, 2=min
        Dim strFP As String
        Dim numMean As Decimal
        Dim numBias As Decimal
        Dim numSD As Decimal
        Dim tblZ As New System.Data.DataTable
        Dim dvAn As system.data.dataview
        Dim p1, p2, p3, p4, p5, p6, p7, p8, p9, p10
        Dim strM As String
        Dim fonts
        Dim numDF As Decimal
        Dim DilFactor
        Dim strF2a As String
        Dim strTempInfo As String
        Dim rowsQC() As DataRow
        Dim rowsRS() As DataRow
        Dim introwsQC As Short
        Dim introwsRS As Short
        Dim arr1(1)
        Dim numA, numB
        Dim numA1 As Decimal
        Dim numB1 As Decimal
        Dim boolIS As Boolean
        Dim strX As String
        Dim rowsX() As DataRow
        'Dim boolX As Boolean
        Dim intCols As Short
        Dim col1, col2, col3, col4, col5, col6 As Short
        Dim intLegStart As Short
        Dim strUnits As String
        Dim intIncr As Short
        Dim intQCRows As Short
        Dim boolJustTable As Boolean

        Dim intExp As Short
        Dim ctExp As Short
        Dim int8 As Short
        Dim intX As Short

        Dim rows2E() As DataRow
        Dim nE As Short
        Dim nI As Short
        Dim boolOutHeadE As Boolean = False
        Dim boolOutHeadI As Boolean = False
        Dim boolDeleteRows As Boolean = False
        Dim boolOutlier As Boolean = False
        Dim intStart As Short
        Dim numAO, numBO
        Dim numAO1 As Decimal
        Dim numBO1 As Decimal


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

        Dim intSID As Int64

        Dim v1, v2, vU


        Dim numPrec As Single
        Dim numTheor As Single

        Dim numMeanFP As Decimal 'full precision
        Dim numAFP As Decimal
        Dim numBFP As Decimal

        boolJustTable = False

        Cursor.Current = Cursors.WaitCursor

        fontsize = wd.ActiveDocument.Styles("Normal").Font.Size ' wd.Selection.Font.Size
        fonts = fontsize ' wd.Selection.Font.Size

        Dim charFCID As String
        strF = "ID_TBLREPORTTABLE = " & idTR
        Dim rowsTR() As DataRow = tblReportTable.Select(strF)
        var1 = rowsTR(0).Item("CHARFCID")
        charFCID = NZ(var1, "NA")

        With wd

            intTableID = 29

            Dim strWRunId As String = GetWatsonColH(intTableID)

            dvDo = frmH.dgvReportTableConfiguration.DataSource
            strF = "id_tblconfigreporttables = " & intTableID
            intDo = FindRowDVNumByCol(intTableID, dvDo, "id_tblconfigreporttables")

            ''Get table name
            'var1 = dvDo(intDo).Item("Table")
            'strTName = NZ(var1, "[NONE]")

            ''get Temperature info
            'var1 = dvDo(intDo).Item("PERIODTEMP")
            'strTempInfo = NZ(var1, "[NONE]")

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

            'find lloq units
            dv = frmH.dgvWatsonAnalRef.DataSource
            int2 = FindRowDV("LLOQ Units", dv)
            strUnits = dv(int2).Item(1)

            int1 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
            str1 = NZ(frmH.dgvStudyConfig(1, int1).Value, "")

            If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
            Else
                strUnits = str1
            End If

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
            '    strM = "Creating " & strTempInfo & " Long-Term QC Standard Storage Stability Assessment ...."
            '    frmH.lblProgress.Text = strM
            '    frmH.Refresh()
            '    MsgBox("Samples have not been assigned to this table.", MsgBoxStyle.Information, "Samples have not been assigned...")
            '    GoTo end2
            'End If

            strF = "IsIntStd = 'No'"
            'strS = "AnalyteDescription ASC"
            strS = "INTORDER ASC, IsIntStd ASC, AnalyteDescription ASC"
            rows11 = tblAnalytesHome.Select(strF, strS)
            intRowsAnal = rows11.Length

            If tblX.Columns.Contains("Ratio") Then
            Else
                tblX.Columns.Add("Ratio", Type.GetType("System.Double"))
                tblX.Columns.Add("AF", Type.GetType("System.Double"))
                tblX.Columns.Add("ELIMINATEDFLAG", Type.GetType("System.String"))
            End If

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

                strX = rows11(Count1 - 1).Item("IsIntStd")
                bool = dvDo.Item(intDo).Item(strDo) 'find boolean value of dvDo column

                Dim strM1 As String
                If bool Then 'continue

                    'ensure data has been entered
                    strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & CleanText(strDo) & "' AND ID_TBLREPORTTABLE = " & idTR
                    rowsX = tbl2.Select(strF)

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

                    If rowsX.Length = 0 Then
                        strM = "Creating " & strTName & "...."
                        strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                        strM1 = strM
                        frmH.lblProgress.Text = strM
                        frmH.Refresh()
                        'MsgBox("Samples have not been assigned to this table.", MsgBoxStyle.Information, "Samples have not been assigned...")
                        'page setup according to configuration
                        str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                        'insert page break
                        'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                        Call InsertPageBreak(wd)
                        Call PageSetup(wd, str1) 'L=Landscape, P=Portrait
                        boolJustTable = True
                        GoTo end1
                    Else
                        boolJustTable = False
                    End If


                    'ReDim arrBCQCs(8, 50) '1=LevelNumber, 2=NomConcentration, 3=ID, 4=FlagPercent, 5=Hello, 6=Lo, 7=#ofReplicates, 8=ASSAYID
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

                    If boolPlaceHolder Then
                        'go directly to table
                        GoTo ph
                    End If

                    'determine if excluded data exists
                    strF2 = ""
                    var1 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteIndex")
                    var2 = tbl4.Rows.Item(Count1 - 1).Item("MasterAssayID")
                    var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                    strF2 = strF2 & "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                    strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                    strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                    strF2 = strF2 & "CHARANALYTE = '" & CleanText(CStr(var3)) & "' AND (ELIMINATEDFLAG = 'Y' OR BOOLEXCLSAMPLE = -1)"

                    'strF2 = strF2 & "ANALYTEINDEX = " & var1 & " AND "
                    'strF2 = strF2 & "MASTERASSAYID = " & var2 & " AND ELIMINATEDFLAG = 'Y'"

                    strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                    rows2 = tbl2.Select(strF2, strS)
                    int1 = rows2.Length 'debug
                    If int1 > 0 Then
                        boolOutlier = True
                    End If

                    ''setup tables
                    'strF2 = ""
                    'var1 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteIndex")
                    'var2 = tbl4.Rows.Item(Count1 - 1).Item("MasterAssayID")
                    'var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")


                    'strF2 = strF2 & "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                    'strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                    'strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                    'strF2 = strF2 & "CHARANALYTE = '" & CleanText(cstr(var3)) & "'" ' & " AND "

                    ''strF2 = strF2 & "ANALYTEINDEX = " & var1 & " AND "
                    ''strF2 = strF2 & "MASTERASSAYID = " & var2 ' & " AND "

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
                    int1 = rows2.Length 'debug
                    dv2 = New DataView(tbl2, strF2, strS, DataViewRowState.CurrentRows)
                    int1 = dv2.Count 'debug

                    'find number of runs used
                    tblNumRuns = dv2.ToTable("a", True, "RUNID")
                    intNumRuns = tblNumRuns.Rows.Count
                    If intNumRuns < 2 Then
                        str1 = "There should be at least 2 Run ID's configured for Long-Term QC Standard Storage Stability Assessment for " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & "."
                        str1 = str1 & ChrW(10) & "When this action is finished, please navigate to the Assigned Samples window and correct this problem."
                        MsgBox(str1, MsgBoxStyle.Information, "Nom Conc problem...")
                        boolJustTable = True

                        'page setup according to configuration
                        str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                        'insert page break
                        'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                        Call InsertPageBreak(wd)
                        Call PageSetup(wd, str1) 'L=Landscape, P=Portrait

                        strMsg = "Run ID's need to be assigned in Assign Samples window."
                        GoTo jt ' end1
                    End If
                    'intNumRuns = 1

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

                    'establish number of QCs evaluated
                    'this will actually give number of columns
                    'must be sorted by nomconc!
                    'make new dv
                    Dim dvNL As New DataView(tbl2, strF2, "NOMCONC ASC", DataViewRowState.CurrentRows)
                    tblLevels = dvNL.ToTable("b", True, "NOMCONC")
                    intNumLevels = tblLevels.Rows.Count
                    For Count2 = 0 To intNumLevels - 1 'check for any null values
                        var3 = tblLevels.Rows.Item(Count2).Item("NOMCONC")
                        If IsDBNull(var3) Then
                            str1 = "The Nominal Concentration for some assigned samples for " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & " have not been configured."
                            str1 = str1 & ChrW(10) & "When this Report Generation action is finished, please navigate to the Assigned Samples window and correct this problem."
                            If boolDisableWarnings Then
                            Else
                                MsgBox(str1, MsgBoxStyle.Information, "Nom Conc problem...")
                            End If
                            strMsg = "Nominal Concentrations need to be assigned in Assign Samples window."

                            'page setup according to configuration
                            str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                            'insert page break
                            'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                            Call InsertPageBreak(wd)
                            Call PageSetup(wd, str1) 'L=Landscape, P=Portrait

                            GoTo jt ' end1
                        End If
                    Next


                    'page setup according to configuration
                    str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                    'insert page break
                    'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                    Call InsertPageBreak(wd)
                    Call PageSetup(wd, str1) 'L=Landscape, P=Portrait

                    ReDim arr1(intNumLevels)

                    'find number of table rows to generate
                    intRowsX = intNumLevels

                    'for reference
                    'tbl1 = tblAnalysisResultsHome
                    'tbl2 = tblAssignedSamples
                    'tbl3 = tblAssignedSamplesHelper
                    'tbl4 = tblAnalytesHome

                    'now determine if data is retrieved from more than one study
                    Dim tblU As System.Data.DataTable
                    Dim rowsU() As DataRow
                    Dim dvU As System.Data.DataView
                    Dim boolDup As Boolean = False

                    strF = "ID_TBLSTUDIES = " & id_tblStudies
                    dvU = New DataView(tbl2, strF, "ID_TBLSTUDIES ASC", DataViewRowState.CurrentRows)
                    tblU = dvU.ToTable("a", True, "ID_TBLSTUDIES2")
                    If tblU.Rows.Count > 1 Then
                        boolDup = True
                    End If

                    'find introwsQC and dataviews CHARHELPER1
                    strF = " AND CHARHELPER2 = 'Long Term Analysis'"
                    strF = strF2 & strF
                    strS = "NOMCONC ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                    dv1 = New DataView(tbl2, strF, strS, DataViewRowState.CurrentRows)
                    tbl2a = dv1.ToTable
                    introwsQC = tbl2a.Rows.Count
                    If introwsQC = 0 Then
                        str1 = "Appropriate assigned samples for " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & " have not been configured as 'Long Term Analysis'."
                        str1 = str1 & ChrW(10) & "When this action is finished, please navigate to the Assigned Samples window and correct this problem."
                        MsgBox(str1, MsgBoxStyle.Information, "Nom Conc problem...")
                        GoTo end1
                    End If

                    intIncr = CInt(introwsQC / intNumLevels)

                    'find introwsRS and dataviews CHARHELPER2
                    strF = " AND CHARHELPER2 = 'T(0) Analysis'"
                    strF = strF2 & strF
                    strS = "NOMCONC ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                    dv3 = New DataView(tbl2, strF, strS, DataViewRowState.CurrentRows)
                    tbl2b = dv3.ToTable
                    introwsRS = tbl2b.Rows.Count
                    If introwsRS = 0 Then
                        str1 = "Appropriate assigned samples for " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & " have not been configured as 'T(0) Analysis'."
                        str1 = str1 & ChrW(10) & "When this action is finished, please navigate to the Assigned Samples window and correct this problem."
                        MsgBox(str1, MsgBoxStyle.Information, "Nom Conc problem...")
                        GoTo end1
                    End If

                    intQCRows = CInt(introwsRS / intNumLevels)


                    'generate table
                    intTblRows = 0
                    intTblRows = intTblRows + 2 'for header
                    intTblRows = intTblRows + 1 'for blank row
                    intTblRows = intTblRows + intIncr 'for Stored Set
                    intTblRows = intTblRows + 1 'for blank row

                    'Increment for Statistics Sections
                    Dim intCSN As Short
                    intCSN = countNumStatsRows()
                    intTblRows = intTblRows + intCSN

                    If intCSN > 0 Then
                    Else
                        intTblRows = intTblRows - 1 'subtract an unneeded blank row
                    End If

                    intTblRows = intTblRows + 1 'for blank row
                    intTblRows = intTblRows + intQCRows 'for QC set
                    intTblRows = intTblRows + 1 'for blank row

                    'Increment for Statistics Sections
                    intTblRows = intTblRows + intCSN

                    If boolQCREPORTACCVALUES Then

                    Else
                        If boolOutlier Then
                            For Count2 = 1 To 2
                                intTblRows = intTblRows + (3 * intNumRuns) 'for stats headings
                                ctExp = ctExp + 3 'for stats headings

                                'Increment for Statistics Sections
                                intTblRows = intTblRows + intCSN
                                ctExp = ctExp + intCSN

                                If intCSN > 0 Then
                                    intTblRows = intTblRows + 1 '(1 * intNumRuns) - 1 'for a blank row after each Mean/Bias/n set, except last set
                                    ctExp = ctExp + 1 '(1 * intNumRuns) - 1
                                End If
                            
                            Next

                            intTblRows = intTblRows - 1 'get one too many rows
                            ctExp = ctExp - 1

                        End If

                    End If

                    intTblRows = intTblRows + 1 'for blank row
                    intTblRows = intTblRows + 1 'for %Difference


                    'intTblRows = intTblRows + 3 'for T0 Mean/SD/etc
                    'intTblRows = intTblRows + 1 'for blank row
                    'intTblRows = intTblRows + 1 'for %Difference

ph:
                    'boolplaceholder needs wrdSelection

                    wrdSelection = wd.Selection()


                    If boolSTATSDIFFCOL Then
                        intCols = (intNumLevels * 2) + 1
                    Else
                        intCols = 1 + intNumLevels
                    End If


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

                        .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.MoveRight Unit:=Microsoft.Office.Interop.Word.wdunits.word.wdunits.wdCharacter, Count:=ctQCs, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        'format selection
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                        .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

                        're-set parameters
                        intLeg = 0
                        ctQCLegend = 0
                        ctDilLeg = 0
                        ctLegend = 0
                        strA = ""
                        strB = ""

                        'Enter  row titles

                        int1 = 4 'row position counter

                        'column 1
                        .Selection.Tables.Item(1).Cell(2, 1).Select()
                        If BOOLINCLUDEDATE Then
                            'str1 = strWRunId & ChrW(10) & "(Analysis Date)"
                            '20180420 LEE:
                            str1 = strWRunId & ChrW(10) & "(" & GetAnalysisDateLabel(intTableID) & ")"
                        Else
                            str1 = strWRunId
                        End If
                        .Selection.TypeText(Text:=str1)

                        .Selection.Tables.Item(1).Cell(int1, 1).Select()
                        'var1 = tblNumRuns.Rows.item(0).Item("RUNID")
                        var1 = tbl2a.Rows.Item(0).Item("RUNID")
                        str1 = var1 & " (Stored)"
                        .Selection.TypeText(Text:=str1)
                        If BOOLINCLUDEDATE Then
                            .Selection.Tables.Item(1).Cell(int1 + 1, 1).Select()
                            str1 = GetDateFromRunID(NZ(var1, 0), LDateFormat, intGroup, idTR)
                            .Selection.TypeText("(" & str1 & ")")
                            .Selection.Tables.Item(1).Cell(int1, 1).Select()
                        End If
                        'column 2
                        int8 = -1

                        If boolQCREPORTACCVALUES Then
                        Else
                            If boolOutlier Then
                                int8 = int8 + 1
                                .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + int8, 2).Select()
                                .Selection.TypeText(Text:="Summary Statistics Excluding Outlier Values")
                                .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                Try
                                    .Selection.Cells.Merge()
                                Catch ex As Exception
                                End Try
                                With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                    .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                End With
                            End If
                        End If

                        intStart = int8
                        Call typeStatsLabels(wd, int8, int1 + intIncr + 1, 1, False)

                        If boolQCREPORTACCVALUES Then
                        Else
                            If boolOutlier Then
                                If boolOutlier Then
                                    int8 = int8 + 2
                                    .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + int8, 2).Select()
                                    .Selection.TypeText(Text:="Summary Statistics Including Outlier Values")
                                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                    Try
                                        .Selection.Cells.Merge()
                                    Catch ex As Exception
                                    End Try
                                    With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                        .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                    End With

                                    Call typeStatsLabels(wd, int8, int1 + intIncr + 1, 1, False)

                                End If
                            End If
                        End If

                        int8 = int8 + 2

                        '.Selection.Tables.Item(1).Cell(int1 + intIncr + 7, 1).Select()
                        .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + int8, 1).Select()
                        'var1 = tblNumRuns.Rows.item(1).Item("RUNID")
                        var1 = tbl2b.Rows.Item(0).Item("RUNID")
                        str1 = var1 & " (T"
                        .Selection.TypeText(Text:=str1)
                        .Selection.Font.Subscript = True
                        str1 = "0"
                        .Selection.TypeText(Text:=str1)
                        .Selection.Font.Subscript = False
                        str1 = ")"
                        .Selection.TypeText(Text:=str1)

                        If BOOLINCLUDEDATE Then
                            .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + int8 + 1, 1).Select()
                            str1 = GetDateFromRunID(NZ(var1, 0), LDateFormat, intGroup, idTR)
                            .Selection.TypeText("(" & str1 & ")")
                            .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + int8, 1).Select()
                        End If

                        If boolQCREPORTACCVALUES Then
                        Else
                            If boolOutlier Then
                                int8 = int8 + 1
                                .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + intQCRows + int8, 2).Select()
                                .Selection.TypeText(Text:="Summary Statistics Excluding Outlier Values")
                                .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                Try
                                    .Selection.Cells.Merge()
                                Catch ex As Exception
                                End Try
                                With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                    .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                End With
                            End If
                        End If

                        Call typeStatsLabels(wd, int8, int1 + intIncr + 1 + intQCRows, 1, False)

                        If boolQCREPORTACCVALUES Then
                        Else
                            If boolOutlier Then
                                If boolOutlier Then
                                    int8 = int8 + 2
                                    .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + intQCRows + int8, 2).Select()
                                    .Selection.TypeText(Text:="Summary Statistics Including Outlier Values")
                                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                    Try
                                        .Selection.Cells.Merge()
                                    Catch ex As Exception
                                    End Try
                                    With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                        .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                    End With

                                    typeStatsLabels(wd, int8, int1 + intIncr + 1 + intQCRows, 1, False)
                                End If
                            End If
                        End If

                        '''''''''''wdd.visible = True
                        Dim boolEnterDiff As Boolean
                        Dim int12 As Short
                        int12 = -1

                        Dim intRC As Short

                        'begin entering data by columns
                        For Count2 = 0 To (intNumLevels * int11) - 1 Step int11

                            int12 = int12 + 1
                            '''''''''''wdd.visible = True

                            'reset starting position
                            int1 = 4

                            'retrieve nomconc
                            varNom = tblLevels.Rows.Item(int12).Item("NOMCONC")

                            'determine hi and lo (nom*flagpercent)
                            strF = "CONCENTRATION = '" & varNom & "'"
                            'rows10 = tblBCQCs.Select(strF)
                            var10 = tbl2a.Rows.Item(0).Item("RUNID")

                            'determine hi and lo (nom*flagpercent)
                            'strF = "CONCENTRATION = " & varNom & " AND ANALYTEID = " & vAnalyteID & " AND MASTERASSAYID = " & vMasterAssayID & " AND ANALYTEINDEX = " & vAnalyteIndex & " AND CONCENTRATION = " & varNom & " AND RUNID = " & var10
                            'if Conc < 1, then the query return 0 records
                            'must do something different
                            var1 = GetANALYTEFLAGPERCENT(varNom, var10, vAnalyteID)
                            v1 = var1
                            v2 = v1

                            'var1 = CDec(NZ(rows10(0).Item("FLAGPERCENT"), 15))
                            arrFP(1, int12) = var1
                            arrFP(2, int12) = var1
                            Call SetHighAndLowCriteria(varNom, var1, var1, hi, lo)

                            'strM = "Creating " & strTempInfo & " Long-Term QC Standard Storage Stability Assessment For " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & "..."
                            frmH.lblProgress.Text = strM1 & ChrW(10) & "Processing Run ID " & var10
                            frmH.Refresh()

                            'for reference
                            'tbl1 = tblAnalysisResultsHome
                            'tbl2 = tblAssignedSamples
                            'tbl3 = tblAssignedSamplesHelper
                            'tbl4 = tblAnalytesHome

                            'start filling in data by rows
                            'intRowsX = 0

                            For Count3 = 0 To 1
                                'column 3
                                If Count3 = 0 Then
                                    'retrive runid for long term
                                    var10 = tbl2a.Rows.Item(0).Item("RUNID")
                                    var1 = tbl2a.Rows.Item(0).Item("ID_TBLSTUDIES2")
                                    intSID = GetWStudyID(var1)
                                Else
                                    'retrive runid for T(0)
                                    var10 = tbl2b.Rows.Item(0).Item("RUNID")
                                    var1 = tbl2b.Rows.Item(0).Item("ID_TBLSTUDIES2")
                                    intSID = GetWStudyID(var1)
                                End If
                                strF = "NOMCONC = " & varNom & " AND RUNID = " & var10 & " AND STUDYID = " & intSID
                                If Count3 = 0 Then 'get Long Term Analysis
                                    Erase rows2a
                                    rows2a = tbl2a.Select(strF)
                                Else 'get T(0) Analysis
                                    Erase rows2a
                                    rows2a = tbl2b.Select(strF)
                                End If
                                int2 = rows2a.Length

                                If Count3 = 0 Then
                                    'enter column headers
                                    .Selection.Tables.Item(1).Cell(2, 2 + Count2).Select()
                                    str1 = rows2a(0).Item("CHARHELPER1")
                                    .Selection.Tables.Item(1).Cell(2, 2 + Count2).Select()
                                    If boolLUseSigFigs Then
                                        var1 = SigFigOrDecString(varNom, LSigFig, False)
                                    Else
                                        var1 = RoundToDecimalRAFZ(varNom, LSigFig)
                                    End If


                                    '******determine if the level is a diln level
                                    Dim strE As String
                                    ' var3 = str1 ' & ChrW(10) & var1 & " " & strConcUnits
                                    var3 = ReturnStdQC(str1)
                                    If boolLUseSigFigs Then
                                        If LboolNomConcParen Then
                                            strE = ChrW(10) & "(" & DisplayNum(var1, LSigFig, False) & ChrW(160) & strConcUnits & ")"
                                        Else
                                            strE = ChrW(10) & DisplayNum(var1, LSigFig, False) & ChrW(160) & strConcUnits
                                        End If
                                    Else
                                        If LboolNomConcParen Then
                                            strE = ChrW(10) & "(" & Format(RoundToDecimalRAFZ(var1, LSigFig), GetRegrDecStr(LSigFig)) & ChrW(160) & strConcUnits & ")"
                                        Else
                                            strE = ChrW(10) & Format(RoundToDecimalRAFZ(var1, LSigFig), GetRegrDecStr(LSigFig)) & ChrW(160) & strConcUnits
                                        End If
                                    End If

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
                                            Dim strAN As String = GetAN(var3)

                                            strA = Chr(96 + intLeg) 'debugging
                                            arrLegend(1, intLeg) = Chr(96 + intLeg) 'a,b,c,etc
                                            'var: units
                                            If boolLUseSigFigs Then
                                                'arrLegend(2, intLeg) = "Dilution QCs undiluted concentration " & DisplayNum(SigFigOrDec(Val(var1), LSigFig, False), LSigFig, False) & " " & strConcUnits & "; a 1:" & var3 & " dilution with blank matrix was performed prior to extraction and analysis."
                                                arrLegend(2, intLeg) = "Dilution QCs undiluted concentration " & DisplayNum(SigFigOrDec(Val(var1), LSigFig, False), LSigFig, False) & " " & strConcUnits & "; " & strAN & " " & var3 & "-fold dilution with blank matrix was performed prior to extraction and analysis."
                                            Else
                                                'arrLegend(2, intLeg) = "Dilution QCs undiluted concentration " & Format(RoundToDecimalRAFZ(CDbl(Val(var1)), LSigFig), GetRegrDecStr(LSigFig)) & " " & strConcUnits & "; a 1:" & var3 & " dilution with blank matrix was performed prior to extraction and analysis."
                                                arrLegend(2, intLeg) = "Dilution QCs undiluted concentration " & Format(RoundToDecimalRAFZ(CDbl(Val(var1)), LSigFig), GetRegrDecStr(LSigFig)) & " " & strConcUnits & "; " & strAN & " " & var3 & "-fold dilution with blank matrix was performed prior to extraction and analysis."
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
                                    .Selection.TypeText(strE)

                                    '******

                                    '.Selection.TypeText(Text:=str2)

                                    If boolSTATSDIFFCOL Then
                                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                        .Selection.TypeText(Text:=ReturnDiffLabel)
                                        .Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                    End If

                                End If

                                strF = ""

                                Erase rows1
                                'rows1 = tbl1.Select(strF)

                                rows1 = rows2a
                                int3a = rows1.Length

                                'REDO HI/LO
                                If int3a = 0 Then
                                    vU = 0
                                Else
                                    vU = rows1(0).Item("BOOLUSEGUWUACCCRIT")
                                    If gAllowGuWuAccCrit And LAllowGuWuAccCrit And vU = -1 Then
                                        v1 = CDec(NZ(rows1(0).Item("NUMMAXACCCRIT"), 0))
                                        v2 = CDec(NZ(rows1(0).Item("NUMMINACCCRIT"), 0))
                                        arrFP(1, int12) = v1
                                        arrFP(2, int12) = v2
                                        Call SetHighAndLowCriteria(varNom, v1, v2, hi, lo)
                                    End If
                                End If

                                'here man
                                'for reference
                                'tbl1 = tblAnalysisResultsHome
                                'tbl2 = tblAssignedSamples
                                'tbl3 = tblAssignedSamplesHelper
                                'tbl4 = tblAnalytesHome

                                'get area ratios
                                tblX.Clear()
                                For Count4 = 0 To int3a - 1

                                    boolOC = False

                                    'if data is from two different studies, then enter study number
                                    If boolDup Then
                                        If Count4 = 0 Then
                                            '.Selection.Tables.Item(1).Cell(int1 + intIncr + 2 + int8, 1).Select()
                                            .Selection.Tables.Item(1).Cell(int1 + Count4 + 1, 1).Select()
                                            var1 = rows2a(0).Item("CHARSTUDYNAME2")
                                            str1 = "Study " & NZ(var1, "[NA]")
                                            .Selection.TypeText(Text:=str1)
                                        End If
                                    End If

                                    var1 = rows1(Count4).Item("CONCENTRATION")
                                    varConc = var1
                                    var1 = NZ(var1, 0)
                                    If boolLUseSigFigs Then
                                        var2 = SigFigOrDec(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                    Else
                                        var2 = RoundToDecimalRAFZ(var1, LSigFig)
                                    End If

                                    var3 = rows1(Count4).Item("ALIQUOTFACTOR")
                                    var4 = NZ(rows1(Count4).Item("ELIMINATEDFLAG"), "N")
                                    var5 = NZ(rows1(Count4).Item("BOOLEXCLSAMPLE"), 0)

                                    Dim nr1 As DataRow = tblX.NewRow
                                    nr1.BeginEdit()
                                    nr1.Item("Ratio") = var2
                                    nr1.Item("AF") = var3
                                    If gAllowExclSamples And LAllowExclSamples And var5 = -1 Then
                                        nr1.Item("ELIMINATEDFLAG") = "Y"
                                    Else
                                        nr1.Item("ELIMINATEDFLAG") = var4
                                    End If

                                    nr1.EndEdit()
                                    tblX.Rows.Add(nr1)

                                    boolEnterDiff = False
                                    If Count3 > -1 Then
                                        'print value
                                        .Selection.Tables.Item(1).Cell(int1 + Count4, 2 + Count2).Select()


                                        '.Selection.TypeText(var1)
                                        If boolLUseSigFigs Then
                                            var4 = DisplayNum(var2, LSigFig, False)
                                        Else
                                            var4 = Format(var2, GetRegrDecStr(LSigFig))
                                        End If
                                        var1 = NZ(rows1(Count4).Item("ELIMINATEDFLAG"), "N")
                                        var3 = NZ(rows1(Count4).Item("BOOLEXCLSAMPLE"), 0)
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

                                            boolEnterDiff = True 'FALSE
                                            intExp = intExp + 1
                                            intLeg = intLeg + 1
                                            strA = ChrW(intLeg + intLegStart)

                                            '20160305 LEE:
                                            'Added DECISIONREASON code
                                            Dim var6
                                            Dim strDecReason As String
                                            'var6 = rows1(Count4).Item("DECISIONREASON")
                                            'Remember, tblAssignedSamples does not have DECISIONREASON
                                            var6 = "No Value: " & GetDECISIONREASONValue(boolExFromAS, vAnalyteID, rows1(Count4))
                                            'Set Legend String
                                            str1 = GetLegendStringExcluded(v1, v2, vU, var6, intTableID, True, "")
                                            'Add to Legend Array
                                            ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                            If boolRedBoldFont Then
                                                .Selection.Font.Bold = True
                                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                            End If

                                            .Selection.TypeText(Text:="NV")

                                            Call typeInSuperscriptFontSize12WithSpace(wd, strA)

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

                                            boolEnterDiff = True 'FALSE
                                            intExp = intExp + 1
                                            intLeg = intLeg + 1
                                            strA = ChrW(intLeg + intLegStart)

                                            '20160305 LEE:
                                            'Added DECISIONREASON code
                                            Dim var6
                                            Dim strDecReason As String
                                            'var6 = rows1(Count4).Item("DECISIONREASON")
                                            'Remember, tblAssignedSamples does not have DECISIONREASON
                                            var6 = GetDECISIONREASONValue(boolExFromAS, vAnalyteID, rows1(Count4))
                                            'Set Legend String
                                            str1 = GetLegendStringExcluded(v1, v2, vU, var6, intTableID, True, "")
                                            'Add to Legend Array
                                            ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                            If boolRedBoldFont Then
                                                .Selection.Font.Bold = True
                                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                            End If

                                            If boolLUseSigFigs Then
                                                .Selection.TypeText(Text:=DisplayNum(var2, LSigFig, False))
                                            Else
                                                .Selection.TypeText(Text:=Format(var2, GetRegrDecStr(LSigFig)))
                                            End If

                                            Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                            '.Selection.TypeText Text:="NR"
                                        Else

                                            'determine if value is outside acceptance criteria
                                            'If var2 > hi Or var2 < lo Then 'flag
                                            If OutsideAccCrit(var2, varNom, v1, v2, NZ(vU, 0)) Then
                                                boolEnterDiff = True
                                                intLeg = intLeg + 1
                                                strA = ChrW(intLeg + intLegStart)

                                                'Set Legend String
                                                str1 = GetLegendStringIncluded(v1, v2, vU)
                                                'Add to Legend Array
                                                ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                                If boolRedBoldFont Then
                                                    .Selection.Font.Bold = True
                                                    .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                End If

                                                If boolLUseSigFigs Then
                                                    .Selection.TypeText(Text:=DisplayNum(var2, LSigFig, False))
                                                Else
                                                    .Selection.TypeText(Text:=Format(var2, GetRegrDecStr(LSigFig)))
                                                End If

                                                Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                            Else
                                                .Selection.TypeText(Text:=CStr(DisplayNum(var2, LSigFig, False)))
                                                boolEnterDiff = True
                                            End If
                                        End If

                                    End If

                                    If boolSTATSDIFFCOL Then
                                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                        If boolEnterDiff Then
                                            'var3 = Format(((var2 / varNom) - 1) * 100, strQCDec)
                                            'var3 = Format(RoundToDecimal(((var2 / varNom) - 1) * 100, intQCDec), strQCDec)
                                            If boolTHEORETICAL Then
                                                var3 = CalcREPercent(var2, varNom, intQCDec)
                                                numTheor = 100 + CDec(var3)

                                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", numTheor, CSng(var10), Count1, strDo, v1, v2, boolOC)

                                            Else
                                                var3 = Format(RoundToDecimal(CalcREPercent(var2, varNom, intQCDec), intQCDec), strQCDec)

                                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", var3, CSng(var10), Count1, strDo, v1, v2, boolOC)

                                            End If
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

                                Next

                                'this is INCLUDED
                                strF = "Ratio >= 0"
                                Erase rows2b
                                rows2b = tblX.Select(strF)
                                var1 = rows2b.Length 'debugging
                                nI = rows2b.Length

                                'do EXCLUDED
                                strF = "Ratio >= 0 AND ELIMINATEDFLAG = 'N'"
                                Erase rows2E
                                rows2E = tblX.Select(strF)
                                nE = rows2E.Length

                                int8 = -1


                                If boolQCREPORTACCVALUES Then
                                Else
                                    If boolOutlier Then
                                        int8 = int8 + 1
                                    End If
                                End If

                                intIncr = int3a

                                v1 = arrFP(1, int12)
                                v2 = arrFP(2, int12)

                                If boolSTATSMEAN Then
                                    Try
                                        'Row1 Mean
                                        int8 = int8 + 1
                                        var1 = MeanDR(rows2E, "Ratio", True, "AF", False, False)
                                        If boolLUseSigFigs Then
                                            var2 = RoundToDecimalA(var1, LSigFig)
                                            var3 = SigFigOrDec(var2, LSigFig, False)
                                        Else
                                            var3 = RoundToDecimalRAFZ(var1, LSigFig)
                                        End If

                                        numMean = var3
                                        numMeanFP = CDec(var1)
                                        If Count3 = 0 Then
                                            numA = numMean   'numA = mumMean = var3
                                            numA1 = var1
                                        Else
                                            numB = numMean   'numB = mumMean = var3
                                            numB1 = var1
                                        End If

                                        If Count3 > -1 Then
                                            .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + int8, 2 + Count2).Select()
                                        Else
                                            .Selection.Tables.Item(1).Cell(int1, 2 + Count2).Select()
                                        End If

                                        'determine if value is outside acceptance criteria
                                        'If (numMean > hi Or numMean < lo) And boolFootNoteQCMean Then 'flag
                                        If (OutsideAccCrit(numMean, varNom, v1, v2, NZ(vU, 0))) And boolFootNoteQCMean Then 'flag
                                            intLeg = intLeg + 1
                                            strA = ChrW(intLeg + intLegStart)

                                            'Set Legend String
                                            str1 = GetLegendStringIncluded(v1, v2, vU)
                                            'Add to Legend Array
                                            ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                            If boolRedBoldFont Then
                                                .Selection.Font.Bold = True
                                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                            End If

                                            '.Selection.TypeText(Text:=CStr(numMean))
                                            If boolLUseSigFigs Then
                                                .Selection.TypeText(CStr(DisplayNum(numMean, LSigFig, False)))
                                            Else
                                                .Selection.TypeText(Format(numMean, GetRegrDecStr(LSigFig)))
                                            End If

                                            Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                            boolEnterDiff = True
                                        Else
                                            '.Selection.TypeText(Text:=CStr(numMean))
                                            If boolLUseSigFigs Then
                                                .Selection.TypeText(CStr(DisplayNum(numMean, LSigFig, False)))
                                            Else
                                                .Selection.TypeText(Format(numMean, GetRegrDecStr(LSigFig)))
                                            End If

                                            boolEnterDiff = True
                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If
                                If boolSTATSSD Then
                                    Try
                                        'Row2 SD
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + int8, 2 + Count2).Select()
                                        If rows2b.Length < gSDMax Then
                                            .Selection.TypeText("NA")
                                            var3 = 0
                                        Else
                                            var1 = StdDevDR(rows2E, "Ratio", True, "AF", False, False)
                                            If boolLUseSigFigs Then
                                                var2 = RoundToDecimalA(var1, LSigFig)
                                                var3 = SigFigOrDec(var2, LSigFig, False)
                                            Else
                                                var3 = RoundToDecimalRAFZ(var1, LSigFig)
                                            End If

                                            If Count3 > -1 Then
                                                If boolLUseSigFigs Then
                                                    .Selection.TypeText(CStr(DisplayNum(var3, LSigFig, False)))
                                                Else
                                                    .Selection.TypeText(Format(RoundToDecimalRAFZ(var3, LSigFig), GetRegrDecStr(LSigFig)))
                                                End If
                                            Else
                                                .Selection.TypeText("NA")
                                            End If
                                            numSD = var3
                                        End If


                                    Catch ex As Exception

                                    End Try
                                End If


                                Try
                                    If rows2b.Length < gSDMax Then
                                    Else
                                        numPrec = CalcCVPercent(numSD, numMean, intQCDec)
                                        Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Precision", numPrec, CSng(var10), Count1, strDo, 0, 0, False)
                                    End If

                                Catch ex As Exception

                                End Try
                                If boolSTATSCV Then
                                    Try
                                        'Row3 %CV
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + int8, 2 + Count2).Select()
                                        If rows2b.Length < gSDMax Then
                                            .Selection.TypeText("NA")
                                        Else
                                            If Count3 > -1 Then
                                                .Selection.TypeText(Format(numPrec, strQCDec))
                                            Else
                                                .Selection.TypeText("NA")
                                            End If
                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If



                                If boolSTATSBIAS And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                    Try
                                        numBias = CalcREPercent(numMean, varNom, intQCDec)
                                        If rows2b.Length = 0 Then
                                        Else
                                            Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", numBias, CSng(var10), Count1, strDo, 0, 0, False)
                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If
                                If boolSTATSBIAS And boolSTATSMEAN Then
                                    Try
                                        'Row4 %Bias '((Mean/NomConc)-1)*100)
                                        int8 = int8 + 1

                                        If Count3 > -1 Then
                                            .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + int8, 2 + Count2).Select()
                                            .Selection.TypeText(Format(numBias, strQCDec))
                                        End If
                                    Catch ex As Exception

                                    End Try
                                End If



                                If boolTHEORETICAL And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                    Try
                                        numTheor = CalcREPercent(numMean, varNom, intQCDec)
                                        numTheor = 100 + CDec(numTheor)
                                        If rows2b.Length = 0 Then
                                        Else
                                            Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", numTheor, CSng(var10), Count1, strDo, 0, 0, False)
                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If
                                If boolTHEORETICAL And boolSTATSMEAN Then
                                    Try
                                        'enter %theoretical
                                        int8 = int8 + 1
                                        If Count3 > -1 Then
                                            .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + int8, 2 + Count2).Select()
                                            .Selection.TypeText(Format(numTheor, strQCDec))
                                        End If

                                    Catch ex As Exception

                                    End Try

                                End If


                                If boolSTATSDIFF And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                    Try
                                        numBias = CalcREPercent(numMean, varNom, intQCDec)
                                        If rows2b.Length = 0 Then
                                        Else
                                            Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", numBias, CSng(var10), Count1, strDo, 0, 0, False)
                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If
                                If boolSTATSDIFF And boolSTATSMEAN Then
                                    Try
                                        'Row4 %Bias '((Mean/NomConc)-1)*100)
                                        int8 = int8 + 1
                                        If Count3 > -1 Then
                                            .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + int8, 2 + Count2).Select()
                                            .Selection.TypeText(Format(numBias, strQCDec))
                                        End If
                                    Catch ex As Exception

                                    End Try
                                End If


                                If BOOLSTATSRE And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                    Try
                                        numBias = CalcREPercent(numMean, varNom, intQCDec)
                                        If rows2b.Length = 0 Then
                                        Else
                                            Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", numBias, CSng(var10), Count1, strDo, 0, 0, False)
                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If
                                If BOOLSTATSRE And boolSTATSMEAN Then
                                    Try
                                        'Row4 %RE '((Mean/NomConc)-1)*100)
                                        int8 = int8 + 1
                                        If Count3 > -1 Then
                                            .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + int8, 2 + Count2).Select()
                                            .Selection.TypeText(Format(numBias, strQCDec))
                                        End If
                                    Catch ex As Exception

                                    End Try
                                End If


                                Try
                                    Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "n", nE, CSng(var10), Count1, strDo, 0, 0, False)
                                Catch ex As Exception

                                End Try
                                If boolSTATSN Then
                                    Try
                                        'Row5 n
                                        int8 = int8 + 1
                                        var3 = rows2b.Length
                                        If Count3 > -1 Then
                                            .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + int8, 2 + Count2).Select()
                                            .Selection.TypeText(CStr(nE))
                                        End If
                                    Catch ex As Exception

                                    End Try
                                End If

                                If boolQCREPORTACCVALUES Then
                                Else
                                    If boolOutlier Then
                                        int8 = int8 + 2

                                        If boolSTATSMEAN Then
                                            Try
                                                'Row1 Mean
                                                int8 = int8 + 1
                                                var1 = MeanDR(rows2b, "Ratio", True, "AF", False, False)
                                                If boolLUseSigFigs Then
                                                    var2 = RoundToDecimalA(var1, LSigFig)
                                                    var3 = SigFigOrDec(var2, LSigFig, False)
                                                Else
                                                    var3 = RoundToDecimalRAFZ(var1, LSigFig)
                                                End If

                                                numMean = var3
                                                If Count3 = 0 Then
                                                    numAO = numMean
                                                    numAO1 = var1
                                                Else
                                                    numBO = numMean
                                                    numBO1 = var1
                                                End If

                                                If Count3 > -1 Then
                                                    .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + int8, 2 + Count2).Select()
                                                Else
                                                    .Selection.Tables.Item(1).Cell(int1, 2 + Count2).Select()
                                                End If

                                                'determine if value is outside acceptance criteria
                                                'If (numMean > hi Or numMean < lo) And boolFootNoteQCMean Then 'flag
                                                If (OutsideAccCrit(numMean, varNom, v1, v2, NZ(vU, 0))) And boolFootNoteQCMean Then 'flag
                                                    intLeg = intLeg + 1
                                                    strA = ChrW(intLeg + intLegStart)

                                                    'Set Legend String
                                                    str1 = GetLegendStringIncluded(v1, v2, vU)
                                                    'Add to Legend Array
                                                    ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                                    If boolRedBoldFont Then
                                                        .Selection.Font.Bold = True
                                                        .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                    End If

                                                    '.Selection.TypeText(Text:=CStr(numMean))
                                                    If boolLUseSigFigs Then
                                                        .Selection.TypeText(CStr(DisplayNum(numMean, LSigFig, False)))
                                                    Else
                                                        .Selection.TypeText(Format(numMean, GetRegrDecStr(LSigFig)))
                                                    End If

                                                    Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                                    boolEnterDiff = True
                                                Else
                                                    '.Selection.TypeText(Text:=CStr(numMean))
                                                    If boolLUseSigFigs Then
                                                        .Selection.TypeText(CStr(DisplayNum(numMean, LSigFig, False)))
                                                    Else
                                                        .Selection.TypeText(Format(numMean, GetRegrDecStr(LSigFig)))
                                                    End If

                                                    boolEnterDiff = True
                                                End If


                                            Catch ex As Exception

                                            End Try
                                        End If
                                        If boolSTATSSD Then
                                            Try
                                                'Row2 SD
                                                int8 = int8 + 1
                                                .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + int8, 2 + Count2).Select()
                                                If nI < gSDMax Then
                                                    .Selection.TypeText("NA")
                                                    var3 = 0
                                                Else
                                                    var1 = StdDevDR(rows2b, "Ratio", True, "AF", False, False)
                                                    If boolLUseSigFigs Then
                                                        var2 = RoundToDecimalA(var1, LSigFig)
                                                        var3 = SigFigOrDec(var2, LSigFig, False)
                                                    Else
                                                        var3 = RoundToDecimalRAFZ(var1, LSigFig)
                                                    End If

                                                    If Count3 > -1 Then
                                                        If boolLUseSigFigs Then
                                                            .Selection.TypeText(CStr(DisplayNum(var3, LSigFig, False)))
                                                        Else
                                                            .Selection.TypeText(Format(RoundToDecimalRAFZ(var3, LSigFig), GetRegrDecStr(LSigFig)))
                                                        End If

                                                    Else
                                                        .Selection.TypeText("NA")
                                                    End If
                                                End If

                                                numSD = var3
                                            Catch ex As Exception

                                            End Try
                                        End If


                                        If boolSTATSCV Then
                                            Try
                                                'Row3 %CV
                                                int8 = int8 + 1
                                                .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + int8, 2 + Count2).Select()
                                                If nI < gSDMax Then
                                                    .Selection.TypeText("NA")
                                                Else
                                                    'var3 = Format(numSD / numMean * 100, strQCDec)
                                                    If Count3 > -1 Then
                                                        numPrec = CalcCVPercent(numSD, numMean, intQCDec)
                                                        .Selection.TypeText(Format(numPrec, strQCDec))
                                                    Else
                                                        .Selection.TypeText("NA")
                                                    End If
                                                End If

                                            Catch ex As Exception

                                            End Try
                                        End If
                                        If boolSTATSBIAS And boolSTATSMEAN Then
                                            Try
                                                'Row4 %Bias '((Mean/NomConc)-1)*100)
                                                int8 = int8 + 1
                                                'var3 = Format(((numMean / varNom) - 1) * 100, strQCDec)
                                                If Count3 > -1 Then
                                                    .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + int8, 2 + Count2).Select()
                                                    '.Selection.TypeText(CStr(var3))

                                                    numBias = CalcREPercent(numMean, varNom, intQCDec)
                                                    .Selection.TypeText(Format(numBias, strQCDec))
                                                End If
                                            Catch ex As Exception

                                            End Try
                                        End If
                                        If boolTHEORETICAL And boolSTATSMEAN Then
                                            Try
                                                'enter %theoretical
                                                int8 = int8 + 1
                                                'var1 = (((numMean / varNom) - 1) * 100)
                                                'var1 = Format(var1, strQCDec)
                                                'var1 = Format(100 + var1, strQCDec)
                                                If Count3 > -1 Then
                                                    .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + int8, 2 + Count2).Select()
                                                    '.Selection.TypeText(CStr(var1))

                                                    numTheor = CalcREPercent(numMean, varNom, intQCDec)
                                                    numTheor = 100 + CDec(numTheor)
                                                    .Selection.TypeText(Format(numTheor, strQCDec))
                                                End If
                                            Catch ex As Exception

                                            End Try

                                        End If

                                        If boolSTATSDIFF And boolSTATSMEAN Then
                                            Try
                                                'Row4 %Bias '((Mean/NomConc)-1)*100)
                                                int8 = int8 + 1
                                                'var3 = Format(((numMean / varNom) - 1) * 100, strQCDec)
                                                If Count3 > -1 Then
                                                    .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + int8, 2 + Count2).Select()
                                                    '.Selection.TypeText(CStr(var3))

                                                    numBias = CalcREPercent(numMean, varNom, intQCDec)
                                                    .Selection.TypeText(Format(numBias, strQCDec))
                                                End If
                                            Catch ex As Exception

                                            End Try
                                        End If

                                        If BOOLSTATSRE And boolSTATSMEAN Then
                                            Try
                                                'Row4 %RE '((Mean/NomConc)-1)*100)
                                                int8 = int8 + 1
                                                'var3 = Format(((numMean / varNom) - 1) * 100, strQCDec)
                                                If Count3 > -1 Then
                                                    .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + int8, 2 + Count2).Select()
                                                    '.Selection.TypeText(CStr(var3))

                                                    numBias = CalcREPercent(numMean, varNom, intQCDec)
                                                    .Selection.TypeText(Format(numBias, strQCDec))
                                                End If
                                            Catch ex As Exception

                                            End Try
                                        End If

                                        If boolSTATSN Then
                                            Try
                                                'Row5 n
                                                int8 = int8 + 1
                                                var3 = rows2b.Length
                                                If Count3 > -1 Then
                                                    .Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + int8, 2 + Count2).Select()
                                                    '.Selection.TypeText(CStr(var3))
                                                    .Selection.TypeText(CStr(nI))

                                                End If
                                            Catch ex As Exception

                                            End Try
                                        End If

                                    End If
                                End If

                                If Count3 = 0 Then
                                    int1 = int1 + intIncr + int8 + 3 '7
                                Else
                                    'int1 = int1 + 3
                                    int1 = int1 + intQCRows + int8 + 1 + 2 '5
                                End If

                            Next

                            If Count2 = 0 Then 'record intX for %Difference
                                intX = int1 'to place %Difference info
                            End If

                            'If Count2 = 0 Then
                            'do %Difference
                            If gboolMeanFullPrec Then
                                '20161206 LEE: gboolMeanFullPrec has been deprecated. Only gboolMeanRounded = true (e.g. gboolMeanFullPrec = false) is allowed
                                var1 = ((numA1 / numB1) - 1) * 100
                            Else
                                var1 = ((numA / numB) - 1) * 100
                            End If

                            var2 = RoundToDecimalA(var1, intQCDec + 3)
                            var3 = Format(RoundToDecimalRAFZ(var2, intQCDec), strQCDec)
                            Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Diff", var3, CSng(var10), Count1, strDo, 0, 0, False)

                            '.Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + 6, 2 + Count2).Select()
                            .Selection.Tables.Item(1).Cell(intX, 2 + Count2).Select()
                            .Selection.TypeText(CStr(var3))

                            If boolQCREPORTACCVALUES Then
                            Else
                                If boolOutlier Then
                                    If gboolMeanFullPrec Then
                                        '20161206 LEE: gboolMeanFullPrec has been deprecated. Only gboolMeanRounded = true (e.g. gboolMeanFullPrec = false) is allowed
                                        var1 = ((numAO1 / numBO1) - 1) * 100
                                    Else
                                        var1 = ((numAO / numBO) - 1) * 100
                                    End If

                                    var2 = RoundToDecimalA(var1, intQCDec + 3)
                                    var3 = Format(RoundToDecimalRAFZ(var2, intQCDec), strQCDec)
                                    '.Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + 6, 2 + Count2).Select()
                                    .Selection.Tables.Item(1).Cell(intX + 1, 2 + Count2).Select()
                                    .Selection.TypeText(CStr(var3))

                                End If
                            End If

                            ''add superscript for legend
                            'ctLegend = ctLegend + 1
                            'intLeg = intLeg + 1
                            'strA = ChrW(intLeg + intLegStart)

                            'End If

                        Next


                        'bottom border this row
                        .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.MoveRight Unit:=Microsoft.Office.Interop.Word.wdunits.word.wdunits.wdCharacter, Count:=ctQCs, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                        'If boolQCREPORTACCVALUES Then
                        'Else
                        '    If intExp = 0 Then
                        '        Call DeleteRows(ctExp, wd)
                        '    End If
                        'End If

                        'autofit window
                        .Selection.Tables.Item(1).Select()
                        'autofit table
                        Call AutoFitTable(wd, False)

                        'go back and merge line 1
                        .Selection.Tables.Item(1).Cell(1, 2).Select()
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.Cells.Merge()
                        If intNumLevels < 2 Then
                        Else
                            Try
                                .Selection.Cells.Merge()
                            Catch ex As Exception

                            End Try
                        End If
                        .Selection.Font.Bold = False
                        .Selection.TypeText(Text:="Nominal Concentrations")
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                        'enter Difference information
                        '.Selection.Tables.Item(1).Cell(int1 + intIncr + 1 + 11, 1).Select()
                        If boolQCREPORTACCVALUES Then
                            .Selection.Tables.Item(1).Cell(intX, 1).Select()
                            str1 = "%Difference"
                            .Selection.TypeText(Text:=str1)
                            ctLegend = ctLegend + 1
                            intLeg = intLeg + 1
                            strA = ChrW(intLeg + intLegStart)
                            .Selection.Font.Superscript = True
                            .Selection.TypeText(Text:=" " & strA)
                            'str1 = "%Difference" ' = ((Mean Stored - Mean T" & ""
                            arrLegend(1, intLeg) = strA
                            arrLegend(2, intLeg) = str1
                            arrLegend(3, intLeg) = True
                        Else
                            If boolOutlier Then
                                .Selection.Tables.Item(1).Cell(intX, 1).Select()
                                str1 = "%Difference (Excluding Outliers)"
                                str1 = "%Difference"
                                .Selection.TypeText(Text:=str1)
                                ctLegend = ctLegend + 1
                                intLeg = intLeg + 1
                                strA = ChrW(intLeg + intLegStart)
                                .Selection.Font.Superscript = True
                                .Selection.TypeText(Text:=" " & strA)
                                .Selection.Font.Superscript = False
                                .Selection.TypeText(Text:=" (Excluding Outliers)")


                                str1 = "%Difference" ' = ((Mean Stored - Mean T" & ""
                                arrLegend(1, intLeg) = strA
                                arrLegend(2, intLeg) = str1
                                arrLegend(3, intLeg) = True

                                .Selection.Tables.Item(1).Cell(intX + 1, 1).Select()
                                str1 = "%Difference (Including Outliers)"
                                str1 = "%Difference"
                                .Selection.TypeText(Text:=str1)
                                'ctLegend = ctLegend + 1
                                'intLeg = intLeg + 1
                                'strA = ChrW(intLeg + intLegStart)
                                .Selection.Font.Superscript = True
                                .Selection.TypeText(Text:=" " & strA)
                                .Selection.Font.Superscript = False
                                .Selection.TypeText(Text:=" (Including Outliers)")

                                'str1 = "%Difference" ' = ((Mean Stored - Mean T" & ""
                                'arrLegend(1, intLeg) = strA
                                'arrLegend(2, intLeg) = str1
                                'arrLegend(3, intLeg) = True

                            Else
                                .Selection.Tables.Item(1).Cell(intX, 1).Select()
                                str1 = "%Difference"
                                .Selection.TypeText(Text:=str1)
                                ctLegend = ctLegend + 1
                                intLeg = intLeg + 1
                                strA = ChrW(intLeg + intLegStart)
                                .Selection.Font.Superscript = True
                                .Selection.TypeText(Text:=" " & strA)
                                'str1 = "%Difference" ' = ((Mean Stored - Mean T" & ""
                                arrLegend(1, intLeg) = strA
                                arrLegend(2, intLeg) = str1
                                arrLegend(3, intLeg) = True
                            End If
                        End If


                    Catch ex As Exception

                        str1 = "There was a problem preparing table:"
                        str1 = strM1 & ChrW(10) & ChrW(10) & str1
                        str1 = str1 & ChrW(10) & ChrW(10)
                        str1 = str1 & ex.Message
                        MsgBox(str1, vbInformation, "Problem...")

                    End Try



                    'enter table number
                    var1 = strTempInfo
                    'replace numeric with verbose in strTempInfo
                    'look for a sequence of characters that is numeric
                    var3 = ""
                    var4 = ""
                    Dim bool1 As Boolean
                    Dim bool2 As Boolean
                    bool1 = False 'Start
                    bool2 = False 'End
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

                    str1 = str2 & " Long-Term QC Standard Storage Stability Assessment: Summary of Interpolated " & rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION") & " QC Standard Concentrations."
                    '***
                    strA = rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")
                    strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, False)
                    Call EnterTableNumber(wd, strTName, 3, strA, strTempInfo, intTableID, intGroup, idTR)
                    '***
                    'Call EnterTableNumber(wd, str1, 3)

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

                    Call AutoFitTable(wd, BOOLINCLUDEDATE)
                    '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                    '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)

                    strM = "Finalizing " & strTName & "..."
                    strM1 = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    str1 = strM1

                    frmH.lblProgress.Text = strM1
                    frmH.Refresh()

                    Call SplitTable(wd, 4, intLeg, arrLegend, str1, False, ctLegend + 2, False, False, False, intTableID)
                    'Sub SplitTable(ByVal wd As Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, 
                    'ByVal arr As Object, ByVal strT As String, ByVal DoLegend As Boolean, ByVal intSplitRows As Short, ByVal boolSmallFont As Boolean)
                    'go to last line and finish Difference
                    .Selection.Tables.Item(1).Select()
                    int1 = .Selection.Tables.Item(1).Rows.Count
                    .Selection.Tables.Item(1).Cell(int1, 1).Select()
                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine)
                    str1 = arrLegend(1, ctLegend)
                    '.Selection.Font.Superscript = True
                    '.Selection.TypeText(Text:=str1)
                    '.Selection.Font.Superscript = False
                    str1 = " = ((Mean Stored - Mean T" & ""
                    .Selection.TypeText(Text:=str1)
                    .Selection.Font.Subscript = True
                    str1 = "0"
                    .Selection.TypeText(Text:=str1)
                    .Selection.Font.Subscript = False
                    str1 = ")/(Mean T"
                    .Selection.TypeText(Text:=str1)
                    .Selection.Font.Subscript = True
                    str1 = "0"
                    .Selection.TypeText(Text:=str1)
                    .Selection.Font.Subscript = False
                    str1 = ")) x 100"
                    .Selection.TypeText(Text:=str1)

                    'move to line below table
                    Call MoveOneCellDown(wd)

                    Call InsertLegend(wd, intTableID, idTR, False, 1)

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
                    'int1 = Len(var1)
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
                    'If bool1 = False Then
                    '    var2 = "[NA]"
                    'Else
                    '    var2 = VerboseNumber(var4, True)
                    '    str2 = Replace(var1, CStr(var4), var2, 1, 1, CompareMethod.Text)
                    'End If

                    'str1 = str2 & " Long-Term QC Standard Storage Stability Assessment: Summary of Interpolated " & rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION") & " QC Standard Concentrations."
                    'str2 = str1
                    'str1 = rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")
                    ''Call JustTable(wd, str1, str2, strDo, strTName, intTableID)
                    'Call JustTable(wd, str1, strTName, strDo, strTName, intTableID, strTempInfo, "")

jt:

                    'str1 = NZ(rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION"), "")
                    ''Call JustTable(wd, str1, str2, strDo, strTName, intTableID)
                    'If Len(str1) = 0 Then
                    'Else
                    '    Call JustTable(wd, str1, strTName, strDo, strTName, intTableID, strTempInfo, strMsg)
                    'End If

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


                End If

next1:

            Next
end2:
        End With

    End Sub

    Sub MVSystemSuit_v1_33(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal idTR As Int64)

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

            intTableID = 33

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

                    ctTbl = ctTbl + 1

                    'ensure data has been entered
                    strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & CleanText(strDo) & "' AND ID_TBLREPORTTABLE = " & idTR
                    rowsX = tbl2.Select(strF)

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

                    'find number of unique Run ID's
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

                    'find number of table rows to generate
                    intRowsX = 0

                    boolOutHeadE = False
                    boolOutHeadI = False
                    boolDeleteRows = False

                    'generate table
                    intTblRows = 0
                    intTblRows = intTblRows + 2 'for header
                    intTblRows = intTblRows + 1 'for blank row
                    intTblRows = intTblRows + intProcRows 'for number of data rows
                    intTblRows = intTblRows + 1 'for one blank row before stats

                    'Increment for Statistics Sections
                    Dim intCSN As Short
                    intCSN = countNumStatsRows()
                    intTblRows = intTblRows + intCSN

                    If intCSN > 0 Then
                    Else
                        intTblRows = intTblRows - 1 'subtract an unneeded blank row
                    End If

                    'add rows for different number of runs
                    int1 = ((intNumRuns - 1) * 8) + (intNumRuns - 1)
                    intTblRows = intTblRows + int1

                    wrdSelection = wd.Selection()

                    Dim intCols As Short
                    intCols = 8


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

                        '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=2, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        ''.Selection.MoveRight Unit:=Microsoft.Office.Interop.Word.wdunits.word.wdunits.wdCharacter, Count:=ctQCs, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
                        '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleSingle
                        '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleSingle

                        'enter 1st Row Headers
                        int1 = InStr(strWRunId, " ", CompareMethod.Text)
                        If int1 = 0 Then
                            .Selection.Tables.Item(1).Cell(3, 1).Select()
                            str2 = strWRunId
                        Else
                            str1 = Mid(strWRunId, 1, int1 - 1)
                            str2 = Mid(strWRunId, int1 + 1, Len(strWRunId))
                        End If

                        If int1 = 0 Then
                        Else
                            .Selection.Tables.Item(1).Cell(1, 1).Select()
                            str1 = str1 ' "Watson"
                            .Selection.TypeText(Text:=str1)
                        End If

                        .Selection.Tables.Item(1).Cell(2, 1).Select()
                        If BOOLINCLUDEDATE Then
                            'str1 = str2 & ChrW(10) & "(Analysis Date)"
                            '20180420 LEE:
                            str1 = str2 & ChrW(10) & "(" & GetAnalysisDateLabel(intTableID) & ")"
                        Else
                            str1 = str2
                        End If

                        .Selection.TypeText(Text:=str1)

                        .Selection.Tables.Item(1).Cell(2, 2).Select()
                        str1 = "Replicate"
                        .Selection.TypeText(Text:=str1)

                        .Selection.Tables.Item(1).Cell(1, 3).Select()
                        str1 = strAnal
                        .Selection.TypeText(Text:=str1)

                        .Selection.Tables.Item(1).Cell(1, 6).Select()
                        str1 = "Internal Standard"
                        .Selection.TypeText(Text:=str1)

                        .Selection.Tables.Item(1).Cell(1, 8).Select()
                        str1 = strAnal & "/IS"
                        .Selection.TypeText(Text:=str1)

                        'enter 2nd row headers
                        .Selection.Tables.Item(1).Cell(2, 3).Select()
                        str1 = "Peak Area"
                        .Selection.TypeText(Text:=str1)

                        .Selection.Tables.Item(1).Cell(2, 4).Select()
                        str1 = "Ret. Time"
                        .Selection.TypeText(Text:=str1)

                        .Selection.Tables.Item(1).Cell(2, 6).Select()
                        str1 = "Peak Area"
                        .Selection.TypeText(Text:=str1)

                        .Selection.Tables.Item(1).Cell(2, 7).Select()
                        str1 = "Ret. Time"
                        .Selection.TypeText(Text:=str1)

                        .Selection.Tables.Item(1).Cell(2, 8).Select()
                        str1 = "Area Ratio"
                        .Selection.TypeText(Text:=str1)

                        ''''wdd.visible = True

                        .Selection.Tables.Item(1).Cell(2, 1).Select()
                        'bottom border this row
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                        'merge and border
                        .Selection.Tables.Item(1).Cell(1, 3).Select()
                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        .Selection.Cells.Merge()
                        With .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                            .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        End With
                        .Selection.Tables.Item(1).Cell(1, 5).Select()
                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        .Selection.Cells.Merge()
                        With .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                            .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        End With

                        'begin entering data'
                        intStart = 4
                        Dim boolExit As Boolean = False

                        Dim numAnal As Single
                        Dim numIS As Single

                        Dim tblAR As New System.Data.DataTable
                        tblAR.Columns.Add("AR", Type.GetType("System.Decimal"))
                        tblAR.Columns.Add("RunID", Type.GetType("System.Decimal"))

                        Dim rowsAR() As DataRow

                        'For Count2 = 0 To intProcRows - 1

                        ''legend

                        ''find number of runs used
                        'tblNumRuns = dv2.ToTable("a", True, "RUNID")
                        'intNumRuns = tblNumRuns.Rows.Count

                        ''find number of different Run Identifiers
                        'tblRID = dv2.ToTable("b", True, "RUNID")
                        'numRID = tblRID.Rows.Count

                        ''re-establish tblLevels to get charhelper2's
                        'tblRID = dv2.ToTable("b", True, "RUNID", "CHARHELPER2")

                        ''end legend

                        Dim strFData1 As String
                        strFData1 = strFData ' & " AND RUNID = " & var10
                        int1 = dv2.Count 'debug
                        Dim tbl2SB As System.Data.DataTable = dv2.ToTable

                        strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                        rows2 = tbl2SB.Select("RUNSAMPLEORDERNUMBER > -1", strS)
                        int3 = rows2.Length

                        'strM = "Creating Summary of " & strTempInfo & " Final Extract Stability Table For " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & "..."
                        frmH.lblProgress.Text = strM1 ' & ChrW(10) & "Processing Run ID " & var10
                        frmH.Refresh()

                        Dim numRT As Single

                        Dim CountR As Short

                        For CountR = 0 To intNumRuns - 1


                            'var10 = rows2(Count3).Item("RUNID")

                            var10 = tblNumRuns.Rows(CountR).Item("RUNID")

                            strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                            rows2 = tbl2SB.Select("RUNID = " & var10, strS)
                            int3 = rows2.Length

                            For Count3 = 0 To int3 - 1

                                'var10 = rows2(Count3).Item("RUNID")
                                'enter runid
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 1).Select()

                                If Count3 = 0 Then
                                    If BOOLINCLUDEDATE Then
                                        str1 = GetDateFromRunID(NZ(var10, 0), LDateFormat, intGroup, idTR)
                                        .Selection.Tables.Item(1).Cell(Count3 + intStart, 1).Select()
                                        .Selection.TypeText(CStr(var10))
                                        .Selection.Tables.Item(1).Cell(Count3 + intStart + 1, 1).Select()
                                        .Selection.TypeText("(" & str1 & ")")
                                    Else
                                        .Selection.TypeText(CStr(var10))
                                    End If
                                End If

                                'enter replicate
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 2).Select()
                                .Selection.TypeText(CStr(Count3 + 1))

                                'enter Analyte Area
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
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 3).Select()
                                .Selection.TypeText(CStr(var1))

                                'enter Analyte RT
                                var1 = NZ(rows2(Count3).Item("ANALYTEPEAKRETENTIONTIME"), "NA")
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 4).Select()
                                If IsNumeric(var1) Then
                                    numRT = var1
                                    .Selection.TypeText(Format(numRT, "0.00"))
                                Else
                                    .Selection.TypeText(Format(var1, "0.00"))
                                End If


                                'enter IS Area
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
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 6).Select()
                                .Selection.TypeText(CStr(var1))

                                'enter IS RT
                                var1 = NZ(rows2(Count3).Item("INTERNALSTANDARDRETENTIONTIME"), "NA")
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 7).Select()
                                If IsNumeric(var1) Then
                                    numRT = var1
                                    .Selection.TypeText(Format(numRT, "0.00"))
                                Else
                                    .Selection.TypeText(Format(var1, "0.00"))
                                End If

                                'enter ratio
                                If numIS = 0 Then
                                    var1 = 1
                                Else
                                    var1 = numAnal / numIS
                                    If boolLUseSigFigsAreaRatio Then
                                        var1 = SigFigAreaRatio(var1, LSigFigAreaRatio, False, True) 'special rounding incorporated
                                    Else
                                        var1 = Format(RoundToDecimalRAFZ(var1, LSigFigAreaRatio), strAreaDecAreaRatio)
                                    End If
                                End If
                                .Selection.Tables.Item(1).Cell(Count3 + intStart, 8).Select()
                                .Selection.TypeText(CStr(var1))

                                var2 = CSng(var1)
                                Dim nrow As DataRow = tblAR.NewRow
                                nrow.BeginEdit()
                                nrow.Item("AR") = var2
                                nrow.Item("RunID") = var10
                                nrow.EndEdit()
                                tblAR.Rows.Add(nrow)

                            Next

                            'Next Count2

                            Erase rowsAR
                            rowsAR = tblAR.Select("RunId = " & var10)

                            Dim intRow As Short
                            intRow = intStart + int3 + 1

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
                            'rows2E = tbl2SA.Select("RUNSAMPLEORDERNUMBER > -1", strS)
                            Erase rows2E
                            rows2E = tbl2SA.Select("RUNID = " & var10, strS)
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

                            Catch ex As Exception

                            End Try
                            If boolSTATSMEAN Then
                                Try
                                    'enter Mean of Peak Area
                                    int8 = int8 + 1

                                    'numMean = MeanDRArea(rows2E, "ANALYTEAREA", False, "ALIQUOTFACTOR", False, False)
                                    '20180720 LEE: MeanDR can accept Area
                                    numMean = MeanDR(rows2E, "ANALYTEAREA", False, "ALIQUOTFACTOR", False, False)
                                    var1 = numMean
                                    If boolLUseSigFigsArea Then
                                        var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                    Else
                                        var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                    End If
                                    numMean = CDec(var1)
                                    Call InsertQCTables(intTableID, idTR, charFCID, NZ(varNom, 0), 1, "MeanSuitAnalyte", numMean, CSng(var10), Count1, strDo, 0, 0, False)

                                    .Selection.Tables.Item(1).Cell(int8, 3).Select()
                                    .Selection.TypeText(CStr(var1))
                                    var1 = var1 'debug
                                Catch ex As Exception

                                End Try

                                Try
                                    'enter Mean of RT
                                    .Selection.Tables.Item(1).Cell(int8, 4).Select()
                                    numMeanRT = MeanDR(rows2E, "ANALYTEPEAKRETENTIONTIME", False, "ALIQUOTFACTOR", True, False)
                                    .Selection.TypeText(CStr(Format(numMeanRT, "0.00")))
                                    numMeanRT = CDec(Format(numMeanRT, "0.00"))
                                    Call InsertQCTables(intTableID, idTR, charFCID, NZ(varNom, 0), 1, "MeanSuitAnalyteRT", numMeanRT, CSng(var10), Count1, strDo, 0, 0, False)
                                Catch ex As Exception

                                End Try

                                Try
                                    'enter Mean of Int Std Peak Area
                                    .Selection.Tables.Item(1).Cell(int8, 6).Select()
                                    'numMeanIS = MeanDRArea(rows2E, "INTERNALSTANDARDAREA", False, "ALIQUOTFACTOR", False, False)
                                    '20180720 LEE: MeanDR can accept Area
                                    numMeanIS = MeanDR(rows2E, "INTERNALSTANDARDAREA", False, "ALIQUOTFACTOR", False, False)
                                    var1 = numMeanIS
                                    If boolLUseSigFigsArea Then
                                        var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                    Else
                                        var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                    End If
                                    numMeanIS = CDec(var1)
                                    .Selection.TypeText(CStr(var1))
                                    Call InsertQCTables(intTableID, idTR, charFCID, NZ(varNom, 0), 1, "MeanSuitIS", numMeanIS, CSng(var10), Count1, strDo, 0, 0, False)
                                Catch ex As Exception

                                End Try

                                Try
                                    'enter Mean Int Std of RT
                                    .Selection.Tables.Item(1).Cell(int8, 7).Select()
                                    numMeanRTIS = MeanDR(rows2E, "INTERNALSTANDARDRETENTIONTIME", False, "ALIQUOTFACTOR", True, False)
                                    .Selection.TypeText(CStr(Format(numMeanRTIS, "0.00")))
                                    numMeanRTIS = CDec(Format(numMeanRTIS, "0.00"))
                                    Call InsertQCTables(intTableID, idTR, charFCID, NZ(varNom, 0), 1, "MeanSuitISRT", numMeanRTIS, CSng(var10), Count1, strDo, 0, 0, False)
                                Catch ex As Exception

                                End Try

                                Try
                                    'enter Mean of Area Ratio
                                    .Selection.Tables.Item(1).Cell(int8, 8).Select()
                                    'numMeanAR = MeanDRArea(rowsAR, "AR", False, "ALIQUOTFACTOR", False, False)
                                    '20180720 LEE: MeanDR can accept Area
                                    numMeanAR = MeanDR(rows2E, "AR", False, "ALIQUOTFACTOR", False, False)
                                    var1 = numMeanAR
                                    'If boolLUseSigFigsArea Then
                                    '    var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                    'Else
                                    '    var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), GetRegrDecStr(LSigFigArea))
                                    'End If

                                    If boolLUseSigFigsAreaRatio Then
                                        var1 = SigFigAreaRatio(var1, LSigFigAreaRatio, False, True) 'special rounding incorporated
                                    Else
                                        var1 = Format(RoundToDecimalRAFZ(var1, LSigFigAreaRatio), strAreaDecAreaRatio)
                                    End If


                                    .Selection.TypeText(CStr(var1))
                                    numMeanAR = CDec(var1)
                                    Call InsertQCTables(intTableID, idTR, charFCID, NZ(varNom, 0), 1, "MeanSuitAreaRatio", numMeanAR, CSng(var10), Count1, strDo, 0, 0, False)
                                    '.Selection.TypeText(CStr(SigFigOrDecString(numMeanAR, LSigFig, False)))
                                Catch ex As Exception

                                End Try

                            End If

                            Dim numSDRT As Single

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
                                        numSD = StdDevDRArea(rows2E, "ANALYTEAREA", False, "ALIQUOTFACTOR", False, False)
                                        Call InsertQCTables(intTableID, idTR, charFCID, NZ(varNom, 0), 1, "Mean", numSD, CSng(var10), Count1, strDo, 0, 0, False)
                                        var1 = numSD
                                        If boolLUseSigFigsArea Then
                                            var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                        Else
                                            var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                        End If
                                        numSD = CDec(var1)
                                        .Selection.TypeText(CStr(var1))
                                        Call InsertQCTables(intTableID, idTR, charFCID, NZ(varNom, 0), 1, "SDSuitAnalyte", numSD, CSng(var10), Count1, strDo, 0, 0, False)
                                        'If numSD < 100 Then
                                        '    .Selection.TypeText(CStr(SigFigOrDecString(numSD, LSigFig, False)))
                                        'Else
                                        '    .Selection.TypeText(CStr(Format(numSD, "0")))
                                        'End If
                                    Catch ex As Exception

                                    End Try

                                    Try
                                        'enter SD of RT
                                        .Selection.Tables.Item(1).Cell(int8, 4).Select()
                                        numSDRT = StdDevDR(rows2E, "ANALYTEPEAKRETENTIONTIME", False, "ALIQUOTFACTOR", True, False)
                                        .Selection.TypeText(Format(numSDRT, "0.00"))
                                        numSDRT = CDec(Format(numSDRT, "0.00"))
                                        Call InsertQCTables(intTableID, idTR, charFCID, NZ(varNom, 0), 1, "SDSuitAnalyteRT", numSDRT, CSng(var10), Count1, strDo, 0, 0, False)
                                    Catch ex As Exception

                                    End Try

                                    Try
                                        'enter SD of IntStd Peak Area
                                        .Selection.Tables.Item(1).Cell(int8, 6).Select()
                                        numSDIS = StdDevDRArea(rows2E, "INTERNALSTANDARDAREA", False, "ALIQUOTFACTOR", False, False)
                                        '.Selection.TypeText(CStr(Format(numSDIS, "0")))
                                        var1 = numSDIS
                                        If boolLUseSigFigsArea Then
                                            var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                        Else
                                            var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                        End If
                                        numSDIS = CDec(var1)
                                        .Selection.TypeText(CStr(var1))
                                        Call InsertQCTables(intTableID, idTR, charFCID, NZ(varNom, 0), 1, "SDSuitIS", numSDIS, CSng(var10), Count1, strDo, 0, 0, False)


                                    Catch ex As Exception

                                    End Try

                                    Try
                                        'enter SD of RT
                                        .Selection.Tables.Item(1).Cell(int8, 7).Select()
                                        numSDRTIS = StdDevDR(rows2E, "INTERNALSTANDARDRETENTIONTIME", False, "ALIQUOTFACTOR", True, False)
                                        .Selection.TypeText(Format(numSDRTIS, "0.00"))
                                        numSDRTIS = CDec(Format(numSDRTIS, "0.00"))
                                        Call InsertQCTables(intTableID, idTR, charFCID, NZ(varNom, 0), 1, "SDSuitISRT", numSDRTIS, CSng(var10), Count1, strDo, 0, 0, False)
                                    Catch ex As Exception

                                    End Try

                                    Try
                                        'enter SD of Area Ratio
                                        .Selection.Tables.Item(1).Cell(int8, 8).Select()
                                        numSDAR = StdDevDRArea(rowsAR, "AR", False, "ALIQUOTFACTOR", False, False)
                                        var1 = numSDAR
                                        'If boolLUseSigFigsArea Then
                                        '    var1 = SigFigArea(var1, LSigFigArea, False, True) 'special rounding incorporated
                                        'Else
                                        '    var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                        'End If

                                        If boolLUseSigFigsAreaRatio Then
                                            var1 = SigFigAreaRatio(var1, LSigFigAreaRatio, False, True) 'special rounding incorporated
                                        Else
                                            var1 = Format(RoundToDecimalRAFZ(var1, LSigFigAreaRatio), strAreaDecAreaRatio)
                                        End If

                                        numSDAR = CDec(var1)
                                        .Selection.TypeText(CStr(var1))
                                        Call InsertQCTables(intTableID, idTR, charFCID, NZ(varNom, 0), 1, "SDSuitAreaRatio", numSDAR, CSng(var10), Count1, strDo, 0, 0, False)
                                        '.Selection.TypeText(CStr(SigFigOrDecString(numSDAR, LSigFig, False)))
                                    Catch ex As Exception

                                    End Try

                                End If

                            End If

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
                                            numPrec = CalcCVPercent(numSD, numMean, intQCDec)
                                            Call InsertQCTables(intTableID, idTR, charFCID, varNom, 1, "PrecisionSuitAnalyte", numPrec, CSng(var10), Count1, strDo, 0, 0, False)
                                            .Selection.TypeText(Format(numPrec, strQCDec))
                                        Catch ex As Exception

                                        End Try

                                        Try
                                            .Selection.Tables.Item(1).Cell(int8, 4).Select()
                                            numPrecRT = RoundToDecimalA(RoundToDecimalRAFZ((numSDRT / numMeanRT * 100), intQCDec + 4), intQCDec)
                                            .Selection.TypeText(Format(numPrecRT, strQCDec))
                                            Call InsertQCTables(intTableID, idTR, charFCID, varNom, 1, "PrecisionSuitAnalyteRT", numPrecRT, CSng(var10), Count1, strDo, 0, 0, False)
                                        Catch ex As Exception

                                        End Try

                                        Try
                                            .Selection.Tables.Item(1).Cell(int8, 6).Select()
                                            numPrecIS = RoundToDecimalA(RoundToDecimalRAFZ((numSDIS / numMeanIS * 100), intQCDec + 4), intQCDec)
                                            var1 = (numSDIS / numMeanIS) 'debug
                                            .Selection.TypeText(Format(numPrecIS, strQCDec))
                                            Call InsertQCTables(intTableID, idTR, charFCID, varNom, 1, "PrecisionSuitIS", numPrecIS, CSng(var10), Count1, strDo, 0, 0, False)
                                        Catch ex As Exception

                                        End Try

                                        Try
                                            .Selection.Tables.Item(1).Cell(int8, 7).Select()
                                            numPrecRTIS = RoundToDecimalA(RoundToDecimalRAFZ((numSDRTIS / numMeanRTIS * 100), intQCDec + 4), intQCDec)
                                            .Selection.TypeText(Format(numPrecRTIS, strQCDec))
                                            Call InsertQCTables(intTableID, idTR, charFCID, varNom, 1, "PrecisionSuitISRT", numPrecRTIS, CSng(var10), Count1, strDo, 0, 0, False)
                                        Catch ex As Exception

                                        End Try

                                        Try
                                            .Selection.Tables.Item(1).Cell(int8, 8).Select()
                                            numPrecAR = RoundToDecimalA(RoundToDecimalRAFZ((numSDAR / numMeanAR * 100), intQCDec + 4), intQCDec)
                                            .Selection.TypeText(Format(numPrecAR, strQCDec))
                                            Call InsertQCTables(intTableID, idTR, charFCID, varNom, 1, "PrecisionSuitAreaRatio", numPrecAR, CSng(var10), Count1, strDo, 0, 0, False)
                                        Catch ex As Exception

                                        End Try


                                    End If

                                Catch ex As Exception

                                End Try
                            End If

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
                                    Call InsertQCTables(intTableID, idTR, charFCID, varNom, 1, "n", int3, CSng(var10), Count1, strDo, 0, 0, False)

                                Catch ex As Exception

                                End Try
                            End If

                            intStart = int8 + 2

                        Next

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
                    .Selection.Tables.Item(1).Cell(5, 5).Select()
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

                    Call AutoFitTable(wd, BOOLINCLUDEDATE)

                    strM = "Finalizing " & strTName & "..."
                    strM1 = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    str1 = strM1

                    frmH.lblProgress.Text = strM1
                    frmH.Refresh()

                    Call SplitTable(wd, 4, intLeg, arrLegend, str1, False, ctLegend + 2, False, False, False, intTableID)
                    'Sub SplitTable(ByVal wd As Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, 
                    'ByVal arr As Object, ByVal strT As String, ByVal DoLegend As Boolean, ByVal intSplitRows As Short, ByVal boolSmallFont As Boolean)

                    'autofit window
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

                    'If bool1 = False Then
                    '    var2 = "[NA]"
                    'Else
                    '    var2 = VerboseNumber(var4, True)
                    '    str2 = Replace(var1, CStr(var4), var2, 1, 1, CompareMethod.Text)
                    'End If

                    'str1 = str2 & " " & strTName ' " Final Extract Stability: Summary of " & rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION") & " Interpolated QC Standard Concentrations."
                    'str2 = str1
                    'str1 = rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")
                    ''Call JustTable(wd, str1, str2, strDo, strTName, intTableID)
                    'Call JustTable(wd, str1, strTName, strDo, strTName, intTableID, strTempInfo, "")

                    str1 = NZ(rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION"), "")
                    'Call JustTable(wd, str1, str2, strDo, strTName, intTableID)
                    If Len(str1) = 0 Then
                    Else
                        strA = strAnal
                        str1 = strA
                        strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, False)
                        Call JustTable(wd, str1, strTName, strDo, strTName, intTableID, strTempInfo, "", strTNameO, intGroup, idTR)
                    End If

                End If

next1:

            Next
end2:
        End With



    End Sub


End Module
